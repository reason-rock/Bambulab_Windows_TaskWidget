import json
import os
import ssl
import time
import inspect
import threading
import tempfile
import atexit
import msvcrt
from typing import Any

import paho.mqtt.client as mqtt
from dotenv import load_dotenv
from winotify import Notification

try:
    import PyTaskbar
    HAS_TASKBAR = True
except Exception:
    PyTaskbar = None
    HAS_TASKBAR = False
try:
    import pystray
    from PIL import Image, ImageDraw, ImageFont
    HAS_TRAY = True
except Exception:
    pystray = None
    Image = None
    ImageDraw = None
    ImageFont = None
    HAS_TRAY = False


# =========================
# 환경변수 로드
# =========================
load_dotenv()

USER_ID = os.getenv("BBL_USER_ID")
ACCESS_TOKEN = os.getenv("BBL_ACCESS_TOKEN")
DEVICE_ID = os.getenv("BBL_DEVICE_ID")

BROKER = "us.mqtt.bambulab.com"
PORT = 8883

REPORT_TOPIC = f"device/{DEVICE_ID}/report"
REQUEST_TOPIC = f"device/{DEVICE_ID}/request"

APP_ID = "Bambu Monitor"


# =========================
# 설정 검증
# =========================
def validate_env() -> None:
    missing = []
    if not USER_ID:
        missing.append("BBL_USER_ID")
    if not ACCESS_TOKEN:
        missing.append("BBL_ACCESS_TOKEN")
    if not DEVICE_ID:
        missing.append("BBL_DEVICE_ID")

    if missing:
        raise RuntimeError(
            "다음 환경변수가 비어 있습니다: "
            + ", ".join(missing)
            + "\n.env 파일을 확인하세요."
        )


# =========================
# 전역 상태
# =========================
latest_print: dict[str, Any] = {}
sequence_id = 1
prev_state: str | None = None
prev_percent: int | None = None
prev_task: str | None = None

taskbar = None
taskbar_ready = False
taskbar_logged_api = False

tray_icon = None
tray_ready = False
tray_lock = threading.Lock()
stop_event = threading.Event()

# Windows Taskbar progress flag fallback values (TBPFLAG)
TBPF_NOPROGRESS = 0
TBPF_INDETERMINATE = 1
TBPF_NORMAL = 2
TBPF_ERROR = 4
TBPF_PAUSED = 8

instance_lock_file = None


def acquire_single_instance_lock() -> bool:
    global instance_lock_file

    lock_name = f"bbmonitor_{DEVICE_ID or 'default'}.lock"
    lock_path = os.path.join(tempfile.gettempdir(), lock_name)

    try:
        instance_lock_file = open(lock_path, "w")
        msvcrt.locking(instance_lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        instance_lock_file.write(str(os.getpid()))
        instance_lock_file.flush()
        atexit.register(release_single_instance_lock)
        return True
    except OSError:
        print("[APP] 이미 실행 중입니다. 중복 실행을 종료합니다.")
        return False


def release_single_instance_lock() -> None:
    global instance_lock_file

    if instance_lock_file is None:
        return

    try:
        instance_lock_file.seek(0)
        msvcrt.locking(instance_lock_file.fileno(), msvcrt.LK_UNLCK, 1)
    except OSError:
        pass
    finally:
        try:
            instance_lock_file.close()
        except Exception:
            pass
        instance_lock_file = None
# =========================
# 유틸
# =========================
def next_seq() -> str:
    global sequence_id
    s = str(sequence_id)
    sequence_id += 1
    return s


def deep_merge(dst: dict[str, Any], src: dict[str, Any]) -> dict[str, Any]:
    for k, v in src.items():
        if isinstance(v, dict) and isinstance(dst.get(k), dict):
            deep_merge(dst[k], v)
        else:
            dst[k] = v
    return dst


def request_pushall(client: mqtt.Client) -> None:
    payload = {
        "pushing": {
            "sequence_id": next_seq(),
            "command": "pushall",
            "version": 1,
            "push_target": 1,
        }
    }
    client.publish(REQUEST_TOPIC, json.dumps(payload), qos=1)


def human_state(raw: str | None) -> str:
    mapping = {
        "RUNNING": "printing",
        "PAUSE": "paused",
        "FINISH": "finished",
        "FAILED": "failed",
        "IDLE": "idle",
        "PREPARE": "preparing",
        "SLICING": "slicing",
    }
    if not raw:
        return "unknown"
    return mapping.get(raw, raw.lower())


def summarize_state(print_obj: dict[str, Any]) -> str:
    state = human_state(print_obj.get("gcode_state"))
    percent = int(print_obj.get("mc_percent", 0) or 0)
    remaining = print_obj.get("mc_remaining_time")
    task = print_obj.get("subtask_name") or print_obj.get("task_name") or "-"
    return (
        f"state={state} | percent={percent}% | "
        f"remaining_min={remaining} | task={task}"
    )


def set_console_title(title: str) -> None:
    # pythonw 실행 시 cmd 창 팝업 방지
    return

def _tray_color_for_state(state: str) -> tuple[int, int, int, int]:
    if state == "printing":
        return (220, 53, 69, 255)  # red
    if state in ("idle", "finished"):
        return (46, 160, 67, 255)  # green
    if state == "paused":
        return (245, 158, 11, 255)  # amber
    if state == "failed":
        return (185, 28, 28, 255)  # dark red
    return (100, 116, 139, 255)  # gray



def _get_tray_font(size: int):
    if ImageFont is None:
        return None
    try:
        return ImageFont.truetype("segoeuib.ttf", size)
    except Exception:
        try:
            return ImageFont.truetype("segoeui.ttf", size)
        except Exception:
            return ImageFont.load_default()


def _build_tray_battery_icon(percent: int, state: str):
    if not HAS_TRAY or Image is None or ImageDraw is None:
        return None

    clamped = max(0, min(int(percent), 100))
    display_percent = 100 if state == "idle" else clamped

    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 시스템 트레이(실제 16x16 축소)에서 보이도록 세로 영역을 거의 꽉 채움
    body = (2, 4, 57, 60)
    cap = (57, 24, 63, 40)
    inner = (5, 7, 54, 57)

    draw.rounded_rectangle(body, radius=6, outline=(30, 41, 59, 255), width=3, fill=(255, 255, 255, 235))
    draw.rounded_rectangle(cap, radius=2, fill=(30, 41, 59, 255))

    fill_color = _tray_color_for_state(state)
    inner_w = inner[2] - inner[0]
    filled = int(inner_w * (display_percent / 100.0))

    if filled > 0:
        draw.rectangle((inner[0], inner[1], inner[0] + filled, inner[3]), fill=fill_color)

    # 16x16 축소시 '%'는 가독성을 크게 떨어뜨려 아이콘에는 숫자만 표시
    label = f"{display_percent}"
    font_size = 34
    font = _get_tray_font(font_size)
    max_w = (inner[2] - inner[0]) - 4
    max_h = (inner[3] - inner[1]) - 4

    while font_size > 10:
        if hasattr(draw, "textbbox"):
            text_bbox = draw.textbbox((0, 0), label, font=font)
            text_w = text_bbox[2] - text_bbox[0]
            text_h = text_bbox[3] - text_bbox[1]
        else:
            text_w, text_h = draw.textsize(label, font=font)

        if text_w <= max_w and text_h <= max_h:
            break

        font_size -= 2
        font = _get_tray_font(font_size)

    cx = (inner[0] + inner[2]) / 2.0
    cy = (inner[1] + inner[3]) / 2.0

    if hasattr(draw, "textbbox"):
        # bbox 원점 오프셋까지 반영해서 시각적으로 정확한 중앙 정렬
        x0, y0, x1, y1 = draw.textbbox((0, 0), label, font=font)
        tx = cx - ((x0 + x1) / 2.0)
        ty = cy - ((y0 + y1) / 2.0)
        draw.text((tx + 2, ty + 2), label, font=font, fill=(0, 0, 0, 180))
        draw.text((tx, ty), label, font=font, fill=(255, 255, 255, 255))
    else:
        text_w, text_h = draw.textsize(label, font=font)
        tx = cx - (text_w / 2.0)
        ty = cy - (text_h / 2.0)
        draw.text((tx + 2, ty + 2), label, font=font, fill=(0, 0, 0, 180))
        draw.text((tx, ty), label, font=font, fill=(255, 255, 255, 255))

    return img


def on_tray_exit(icon, item) -> None:
    print("[TRAY] 종료 메뉴 클릭")
    stop_event.set()
    try:
        icon.stop()
    except Exception:
        pass

def init_tray_icon() -> None:
    global tray_icon, tray_ready

    if not HAS_TRAY:
        print("[TRAY] pystray/Pillow 미설치. 트레이 아이콘 비활성화.")
        print("[TRAY] 설치: pip install pystray pillow")
        return

    try:
        icon_image = _build_tray_battery_icon(0, "unknown")
        menu = pystray.Menu(pystray.MenuItem("종료", on_tray_exit))
        tray_icon = pystray.Icon("bambu_monitor", icon_image, "Bambu Monitor", menu)

        thread = threading.Thread(target=tray_icon.run, daemon=True)
        thread.start()

        tray_ready = True
        print("[TRAY] initialized (system tray icon)")
    except Exception as e:
        tray_icon = None
        tray_ready = False
        print("[TRAY] init failed:", e)


def update_tray_icon(percent: int, state: str, task: str) -> None:
    if not tray_ready or tray_icon is None:
        return

    icon_image = _build_tray_battery_icon(percent, state)
    if icon_image is None:
        return

    clamped = max(0, min(int(percent), 100))
    display_percent = 100 if state == "idle" else clamped
    short_task = (task or "-")[:42]

    try:
        with tray_lock:
            tray_icon.icon = icon_image
            tray_icon.title = f"Bambu Monitor | {display_percent}% | {state} | {short_task}"
    except Exception as e:
        print("[TRAY] update failed:", e)


def shutdown_tray_icon() -> None:
    global tray_icon, tray_ready

    if tray_icon is None:
        return

    try:
        tray_icon.stop()
    except Exception as e:
        print("[TRAY] stop failed:", e)
    finally:
        tray_icon = None
        tray_ready = False
# =========================
# 작업표시줄 진행률
# =========================
def _resolve_progress_type_for_state(state: str):
    """
    PyTaskbar ProgressType 매핑
    문서 예시:
    - NOPROGRESS
    - INDETERMINATE
    - NORMAL
    - PAUSED
    - ERROR
    """
    if not hasattr(PyTaskbar, "ProgressType"):
        # PyTaskbar enum이 없을 때는 Win32 TBPFLAG 값으로 fallback
        if state == "printing":
            return TBPF_ERROR
        if state == "paused":
            return TBPF_PAUSED
        if state == "failed":
            return TBPF_ERROR
        if state == "finished":
            return TBPF_NORMAL
        if state == "idle":
            return TBPF_NORMAL
        return TBPF_NOPROGRESS

    pt = PyTaskbar.ProgressType

    if state == "printing":
        return getattr(pt, "ERROR", TBPF_ERROR)
    if state == "paused":
        return getattr(pt, "PAUSED", TBPF_PAUSED)
    if state == "failed":
        return getattr(pt, "ERROR", TBPF_ERROR)
    if state == "finished":
        return getattr(pt, "NORMAL", TBPF_NORMAL)
    if state == "idle":
        return getattr(pt, "NORMAL", TBPF_NORMAL)

    return getattr(pt, "NOPROGRESS", TBPF_NOPROGRESS)


def _call_taskbar_method(names: list[str], *args) -> bool:
    if taskbar is None:
        return False
    for name in names:
        fn = getattr(taskbar, name, None)
        if callable(fn):
            fn(*args)
            return True
    return False


def init_taskbar() -> None:
    global taskbar, taskbar_ready

    if not HAS_TASKBAR:
        print("[TASKBAR] PyTaskbar import 실패. 작업표시줄 기능 비활성화.")
        return

    try:
        if not hasattr(PyTaskbar, "TaskbarProgress"):
            print("[TASKBAR] TaskbarProgress 클래스가 없습니다.")
            return

        tb_cls = PyTaskbar.TaskbarProgress
        sig = inspect.signature(tb_cls)
        if len(sig.parameters) == 0:
            taskbar = tb_cls()
        else:
            hwnd = None
            try:
                import ctypes
                hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            except Exception:
                hwnd = None

            if hwnd:
                try:
                    taskbar = tb_cls(hwnd)
                except Exception:
                    taskbar = tb_cls()
            else:
                taskbar = tb_cls()

        taskbar_ready = True
        print("[TASKBAR] initialized with TaskbarProgress")
    except Exception as e:
        taskbar = None
        taskbar_ready = False
        print("[TASKBAR] init failed:", e)


def update_taskbar_progress(percent: int, state: str) -> None:
    global taskbar_logged_api

    if not taskbar_ready or taskbar is None:
        return

    clamped = max(0, min(int(percent), 100))

    try:
        if not taskbar_logged_api:
            print("[TASKBAR] methods:", [m for m in dir(taskbar) if not m.startswith("_")])
            taskbar_logged_api = True

        progress_type = _resolve_progress_type_for_state(state)
        _call_taskbar_method(
            ["set_progress_type", "setProgressType", "SetProgressType", "set_state", "SetState"],
            progress_type,
        )

        if state in ("printing", "paused", "failed", "finished", "idle"):
            if state == "finished":
                _call_taskbar_method(["set_progress", "setProgress", "SetProgress"], 100)
                # 완료 강조
                if hasattr(taskbar, "flash_done"):
                    taskbar.flash_done()
            elif state == "idle":
                # idle일 때도 초록색 바를 보여 요구사항을 만족
                _call_taskbar_method(["set_progress", "setProgress", "SetProgress"], 100)
            else:
                _call_taskbar_method(["set_progress", "setProgress", "SetProgress"], clamped)
        else:
            # unknown 등
            _call_taskbar_method(["reset", "Reset"])

    except Exception as e:
        print("[TASKBAR] update failed:", e)

def notify(title: str, message: str, important: bool = False) -> None:
    try:
        duration = "long" if important else "short"
        toast = Notification(
            app_id=APP_ID,
            title=title,
            msg=message,
            duration=duration,
        )
        toast.show()

        # 토스트가 환경에 따라 보이지 않을 때를 대비해 트레이 풍선 알림도 같이 표시
        if important and tray_ready and tray_icon is not None and hasattr(tray_icon, "notify"):
            try:
                tray_icon.notify(message, title)
            except Exception:
                pass

    except Exception as e:
        print("[TOAST] notify failed:", e)
def handle_notifications(print_obj: dict[str, Any]) -> None:
    global prev_state, prev_percent, prev_task

    state = human_state(print_obj.get("gcode_state"))
    percent = int(print_obj.get("mc_percent", 0) or 0)
    task = print_obj.get("subtask_name") or print_obj.get("task_name") or "Unknown job"

    if prev_state != "printing" and state == "printing":
        notify("출력 시작", f"{task} 출력을 시작했습니다. ({percent}%)", important=True)

    if prev_state != "finished" and state == "finished":
        notify("출력 완료", f"{task} 출력이 끝났습니다.", important=True)

    if prev_state != "failed" and state == "failed":
        notify("3D Print Failed", task)

    if state == "printing":
        if prev_percent is None:
            prev_percent = percent
        else:
            prev_bucket = prev_percent // 25
            curr_bucket = percent // 25
            if curr_bucket > prev_bucket and percent < 100:
                notify("3D Print Progress", f"{task}: {percent}%")
        prev_percent = percent
    else:
        prev_percent = percent

    prev_state = state
    prev_task = task


# =========================
# MQTT 콜백
# =========================
def on_connect(client: mqtt.Client, userdata, flags, reason_code, properties=None):
    print(f"[MQTT] connected: {reason_code}")
    client.subscribe(REPORT_TOPIC, qos=1)
    request_pushall(client)


def on_message(client: mqtt.Client, userdata, msg: mqtt.MQTTMessage):
    try:
        payload = json.loads(msg.payload.decode("utf-8", errors="ignore"))
    except Exception as e:
        print("[MQTT] invalid json:", e)
        return

    if "print" not in payload or not isinstance(payload["print"], dict):
        return

    deep_merge(latest_print, payload["print"])

    state = human_state(latest_print.get("gcode_state"))
    percent = int(latest_print.get("mc_percent", 0) or 0)
    task = latest_print.get("subtask_name") or latest_print.get("task_name") or "-"

    if state == "printing":
        set_console_title(f"H2S {percent}% - {task}")
    elif state == "finished":
        set_console_title(f"H2S Finished - {task}")
    elif state == "idle":
        set_console_title("H2S Idle")
    elif state == "failed":
        set_console_title(f"H2S Failed - {task}")
    elif state == "paused":
        set_console_title(f"H2S Paused {percent}% - {task}")
    else:
        set_console_title(f"H2S {state} - {task}")

    update_taskbar_progress(percent, state)
    update_tray_icon(percent, state, task)
    handle_notifications(latest_print)

    print("[PRINT]", summarize_state(latest_print))


def on_disconnect(client: mqtt.Client, userdata, disconnect_flags, reason_code, properties=None):
    print(f"[MQTT] disconnected: {reason_code}")


# =========================
# 메인
# =========================
def main():
    if not acquire_single_instance_lock():
        return

    stop_event.clear()
    validate_env()
    init_taskbar()
    init_tray_icon()

    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
    client.username_pw_set(username=f"u_{USER_ID}", password=ACCESS_TOKEN)

    context = ssl.create_default_context()
    client.tls_set_context(context)

    client.on_connect = on_connect
    client.on_message = on_message
    client.on_disconnect = on_disconnect

    client.connect(BROKER, PORT, keepalive=60)
    client.loop_start()

    try:
        while not stop_event.is_set():
            for _ in range(60):
                if stop_event.is_set():
                    break
                time.sleep(1)
            if stop_event.is_set():
                break
            request_pushall(client)
    except KeyboardInterrupt:
        print("stopping...")
        stop_event.set()
    finally:
        client.loop_stop()
        client.disconnect()
        shutdown_tray_icon()


if __name__ == "__main__":
    main()