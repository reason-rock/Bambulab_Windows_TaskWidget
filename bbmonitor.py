import atexit
import ctypes
import inspect
import json
import logging
import msvcrt
import os
import shutil
import ssl
import subprocess
import tempfile
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
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


load_dotenv()

BROKER = "us.mqtt.bambulab.com"
PORT = 8883
APP_ID = "Bambu Monitor"
APP_USER_MODEL_ID = "reason_rock.BambuMonitor"
CONFIG_FILENAME = "bambu_monitor_config.json"

TBPF_NOPROGRESS = 0
TBPF_INDETERMINATE = 1
TBPF_NORMAL = 2
TBPF_ERROR = 4
TBPF_PAUSED = 8


config_lock = threading.Lock()
settings_request_event = threading.Event()
stop_event = threading.Event()

taskbar = None
taskbar_ready = False
taskbar_logged_api = False

tray_icon = None
tray_ready = False
tray_lock = threading.Lock()

instance_lock_file = None

sequence_id = 1
config_data: dict[str, Any] = {}
printer_states: dict[str, dict[str, Any]] = {}
printer_prev: dict[str, dict[str, Any]] = {}
tray_rotation_index = 0
printer_filament_info: dict[str, dict[str, Any]] = {}
printer_runout_notified: dict[str, bool] = {}

monitor_client: mqtt.Client | None = None
monitor_client_lock = threading.Lock()

mqtt_logger = logging.getLogger("mqtt")
if not mqtt_logger.handlers:
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("[MQTT] %(message)s"))
    mqtt_logger.addHandler(handler)
    mqtt_logger.setLevel(logging.DEBUG)

RECONNECT_BASE_DELAY = 2
RECONNECT_MAX_DELAY = 120
RECONNECT_MAX_ATTEMPTS = 20
reconnect_attempts = 0
mqtt_connected = False
mqtt_connected_time: float | None = None


def config_path() -> Path:
    if getattr(__import__("sys"), "frozen", False):
        return Path(__import__("sys").executable).resolve().with_name(CONFIG_FILENAME)
    return Path(__file__).resolve().with_name(CONFIG_FILENAME)


def default_config() -> dict[str, Any]:
    env_device = os.getenv("BBL_DEVICE_ID", "").strip()
    env_alias = os.getenv("BBL_PRINTER_NAME", "").strip()
    printers = []
    if env_device:
        printers.append({"alias": env_alias or env_device, "device_id": env_device})
    return {
        "user_id": os.getenv("BBL_USER_ID", "").strip(),
        "access_token": os.getenv("BBL_ACCESS_TOKEN", "").strip(),
        "email": os.getenv("BBL_EMAIL", "").strip(),
        "printers": printers,
    }


def load_config() -> dict[str, Any]:
    path = config_path()
    cfg = default_config()
    if path.exists():
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                cfg.update({k: v for k, v in raw.items() if k in cfg})
        except Exception as e:
            print("[CONFIG] load failed:", e)

    printers: list[dict[str, str]] = []
    for item in cfg.get("printers", []):
        if not isinstance(item, dict):
            continue
        device_id = str(item.get("device_id", "")).strip()
        alias = str(item.get("alias", item.get("name", ""))).strip() or device_id
        if device_id:
            printers.append({"alias": alias, "device_id": device_id})

    cfg["user_id"] = str(cfg.get("user_id", "")).strip()
    cfg["access_token"] = str(cfg.get("access_token", "")).strip()
    cfg["email"] = str(cfg.get("email", "")).strip()
    cfg["printers"] = printers
    return cfg


def save_config(cfg: dict[str, Any]) -> None:
    cleaned = {
        "user_id": str(cfg.get("user_id", "")).strip(),
        "access_token": str(cfg.get("access_token", "")).strip(),
        "email": str(cfg.get("email", "")).strip(),
        "printers": [],
    }
    for item in cfg.get("printers", []):
        if not isinstance(item, dict):
            continue
        device_id = str(item.get("device_id", "")).strip()
        alias = str(item.get("alias", item.get("name", ""))).strip() or device_id
        if device_id:
            cleaned["printers"].append({"alias": alias, "device_id": device_id})

    path = config_path()
    path.write_text(json.dumps(cleaned, ensure_ascii=False, indent=2), encoding="utf-8")


def validate_config(cfg: dict[str, Any]) -> None:
    missing = []
    if not cfg.get("user_id"):
        missing.append("user_id")
    if not cfg.get("access_token"):
        missing.append("access_token")
    if not cfg.get("printers"):
        missing.append("printers")
    if missing:
        raise RuntimeError("다음 설정값이 비어 있습니다: " + ", ".join(missing))


def acquire_single_instance_lock() -> bool:
    global instance_lock_file

    lock_path = os.path.join(tempfile.gettempdir(), "bbmonitor.lock")

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


def next_seq() -> str:
    global sequence_id
    value = str(sequence_id)
    sequence_id += 1
    return value


def deep_merge(dst: dict[str, Any], src: dict[str, Any]) -> dict[str, Any]:
    for key, value in src.items():
        if isinstance(value, dict) and isinstance(dst.get(key), dict):
            deep_merge(dst[key], value)
        else:
            dst[key] = value
    return dst


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
        f"state={state} | percent={percent}% | remaining_min={remaining} | task={task}"
    )


def printer_label(device_id: str) -> str:
    for printer in config_data.get("printers", []):
        if printer.get("device_id") == device_id:
            return printer.get("alias") or printer.get("name") or device_id
    return device_id


def extract_filament_info(print_obj: dict[str, Any]) -> dict[str, Any]:
    """AMS/필라멘트 스풀 정보를 추출합니다."""
    info: dict[str, Any] = {
        "ams_slots": [],
        "active_tray": None,
        "target_tray": None,
        "active_filament_type": None,
        "active_filament_color": None,
        "has_ams": False,
    }

    # AMS 정보 추출
    ams_data = print_obj.get("ams")
    if ams_data and isinstance(ams_data, dict):
        info["has_ams"] = True
        ams_list = ams_data.get("ams") or []
        if isinstance(ams_list, list):
            for ams_unit in ams_list:
                if not isinstance(ams_unit, dict):
                    continue
                ams_id = ams_unit.get("id", "?")
                trays = ams_unit.get("tray") or []
                for tray in trays:
                    if not isinstance(tray, dict):
                        continue
                    slot_info = {
                        "ams_id": ams_id,
                        "tray_id": tray.get("id", "?"),
                        "type": tray.get("type", ""),
                        "color": tray.get("color", ""),
                        "weight": tray.get("weight"),
                        "remain": tray.get("remain"),
                        "tag_uid": tray.get("tag_uid", ""),
                        "name": tray.get("name", ""),
                    }
                    info["ams_slots"].append(slot_info)

    # 현재 활성 트레이 정보
    tray_now = print_obj.get("tray_now")
    tray_tar = print_obj.get("tray_tar")
    info["active_tray"] = tray_now
    info["target_tray"] = tray_tar

    # 현재 필라멘트 타입/컬러
    info["active_filament_type"] = print_obj.get("filament_type")
    info["active_filament_color"] = print_obj.get("filament_color")

    # AMS에서 활성 트레이 정보 찾기
    if info["has_ams"] and tray_now is not None:
        for slot in info["ams_slots"]:
            if str(slot.get("tray_id")) == str(tray_now):
                if not info["active_filament_type"]:
                    info["active_filament_type"] = slot.get("type", "")
                if not info["active_filament_color"]:
                    info["active_filament_color"] = slot.get("color", "")
                info["active_slot_detail"] = slot
                break

    return info


def format_filament_summary(filament_info: dict[str, Any]) -> str:
    """필라멘트 정보를 요약 문자열로 반환합니다."""
    if not filament_info:
        return ""

    parts = []

    active_type = filament_info.get("active_filament_type")
    if active_type:
        parts.append(f"{active_type}")

    active_tray = filament_info.get("active_tray")
    if active_tray is not None:
        parts.append(f"트레이:{active_tray}")

    ams_slots = filament_info.get("ams_slots", [])
    if ams_slots:
        slot_summaries = []
        for slot in ams_slots:
            t = slot.get("type", "?")
            remain = slot.get("remain")
            if remain is not None:
                slot_summaries.append(f"{t}({remain}%)")
            else:
                slot_summaries.append(t)
        parts.append(f"AMS: {', '.join(slot_summaries)}")

    return " | ".join(parts) if parts else ""


def check_runout_error(print_obj: dict[str, Any]) -> bool:
    """필라멘트 run-out 에러 여부를 확인합니다."""
    state = print_obj.get("gcode_state", "")
    if state != "PAUSE":
        return False

    # print_error 코드로 run-out 확인 (문자열 또는 정수)
    print_error = print_obj.get("print_error")
    if print_error is not None:
        print_error_str = str(print_error).lower().strip()
        # Bambu Lab run-out 관련 에러 코드
        runout_error_codes = [
            "07008102",  # filament runout
            "07008033",  # AMS filament runout
            "07008101",  # filament runout related
            "128102",  # 0x1f526 runout decimal variant
        ]
        for code in runout_error_codes:
            if code in print_error_str:
                return True

    # hw_switch_action 필드로 확인
    hw_switch = str(print_obj.get("hw_switch_action", "") or "").lower()
    if "filament_runout" in hw_switch or "runout" in hw_switch:
        return True

    # ams_status 필드 확인 (정수형 AMS 상태 코드)
    ams_status = print_obj.get("ams_status")
    if ams_status is not None:
        try:
            ams_status_int = int(ams_status)
            # AMS 상태 코드 중 run-out 관련 값
            if ams_status_int in (0x08000102, 0x08008102, 0x07008102):
                return True
        except (ValueError, TypeError):
            pass

    # filament_runout 필드 직접 확인
    if print_obj.get("filament_runout"):
        return True

    return False


def request_pushall(client: mqtt.Client, device_id: str) -> None:
    payload = {
        "pushing": {
            "sequence_id": next_seq(),
            "command": "pushall",
            "version": 1,
            "push_target": 1,
        }
    }
    client.publish(f"device/{device_id}/request", json.dumps(payload), qos=1)


def set_console_title(title: str) -> None:
    return


def set_app_user_model_id() -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception as e:
        print("[APP] AppUserModelID set failed:", e)


def _tray_color_for_state(state: str) -> tuple[int, int, int, int]:
    if state == "printing":
        return (220, 53, 69, 255)
    if state in ("idle", "finished"):
        return (46, 160, 67, 255)
    if state == "paused":
        return (245, 158, 11, 255)
    if state == "failed":
        return (185, 28, 28, 255)
    if state == "disconnected":
        return (108, 117, 125, 255)
    return (100, 116, 139, 255)


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
    body = (2, 4, 57, 60)
    cap = (57, 24, 63, 40)
    inner = (5, 7, 54, 57)

    draw.rounded_rectangle(
        body, radius=6, outline=(30, 41, 59, 255), width=3, fill=(255, 255, 255, 235)
    )
    draw.rounded_rectangle(cap, radius=2, fill=(30, 41, 59, 255))

    fill_color = _tray_color_for_state(state)
    inner_w = inner[2] - inner[0]
    filled = int(inner_w * (display_percent / 100.0))
    if filled > 0:
        draw.rectangle(
            (inner[0], inner[1], inner[0] + filled, inner[3]), fill=fill_color
        )

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


def on_tray_exit(icon=None, item=None) -> None:
    stop_event.set()
    try:
        if icon is not None:
            icon.stop()
    except Exception:
        pass


def open_settings_from_tray(icon=None, item=None) -> None:
    settings_request_event.set()


def _find_bambu_studio_target() -> str | None:
    candidates = [
        shutil.which("bambu-studio.exe"),
        shutil.which("BambuStudio.exe"),
        shutil.which("Bambu Studio.exe"),
        os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Programs",
            "Bambu Studio",
            "bambu-studio.exe",
        ),
        os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Programs",
            "Bambu Studio",
            "BambuStudio.exe",
        ),
        os.path.join(
            os.environ.get("PROGRAMFILES", ""), "Bambu Studio", "bambu-studio.exe"
        ),
        os.path.join(
            os.environ.get("PROGRAMFILES", ""), "Bambu Studio", "BambuStudio.exe"
        ),
        os.path.join(
            os.environ.get("PROGRAMFILES(X86)", ""), "Bambu Studio", "bambu-studio.exe"
        ),
        os.path.join(
            os.environ.get("PROGRAMFILES(X86)", ""), "Bambu Studio", "BambuStudio.exe"
        ),
        os.path.join(
            os.environ.get("APPDATA", ""),
            "Microsoft",
            "Windows",
            "Start Menu",
            "Programs",
            "Bambu Studio.lnk",
        ),
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None


def open_bambu_studio(icon=None, item=None) -> None:
    target = _find_bambu_studio_target()
    if not target:
        notify(
            "Bambu Studio", "Bambu Studio 실행 파일을 찾지 못했습니다.", important=True
        )
        return
    try:
        if target.lower().endswith(".lnk"):
            os.startfile(target)
        else:
            subprocess.Popen([target], close_fds=True)
    except Exception as e:
        notify("Bambu Studio", "Bambu Studio를 실행하지 못했습니다.", important=True)
        print("[TRAY] failed to launch Bambu Studio:", e)


def init_tray_icon() -> None:
    global tray_icon, tray_ready
    if not HAS_TRAY:
        print("[TRAY] pystray/Pillow 미설치. 트레이 아이콘 비활성화.")
        return
    try:
        icon_image = _build_tray_battery_icon(0, "unknown")
        menu = pystray.Menu(
            pystray.MenuItem("Bambu Studio 열기", open_bambu_studio, default=True),
            pystray.MenuItem("설정 열기", open_settings_from_tray),
            pystray.MenuItem("종료", on_tray_exit),
        )
        tray_icon = pystray.Icon("bambu_monitor", icon_image, "Bambu Monitor", menu)
        threading.Thread(target=tray_icon.run, daemon=True).start()
        tray_ready = True
    except Exception as e:
        tray_icon = None
        tray_ready = False
        print("[TRAY] init failed:", e)


def update_tray_icon(
    percent: int, state: str, task: str, label: str, filament_summary: str = ""
) -> None:
    if not tray_ready or tray_icon is None:
        return
    icon_image = _build_tray_battery_icon(percent, state)
    if icon_image is None:
        return
    clamped = max(0, min(int(percent), 100))
    display_percent = 100 if state == "idle" else clamped
    short_task = (task or "-")[:30]

    status_prefix = ""
    if not mqtt_connected:
        if reconnect_attempts >= RECONNECT_MAX_ATTEMPTS:
            status_prefix = "[연결 끊김] "
        else:
            status_prefix = f"[재연결 {reconnect_attempts}/{RECONNECT_MAX_ATTEMPTS}] "

    filament_part = f"\n{filament_summary}" if filament_summary else ""
    try:
        with tray_lock:
            tray_icon.icon = icon_image
            tray_icon.title = f"{status_prefix}{label} | {display_percent}% | {state} | {short_task}{filament_part}"
    except Exception as e:
        mqtt_logger.debug("tray update failed: %s", e)


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


def _resolve_progress_type_for_state(state: str):
    if not hasattr(PyTaskbar, "ProgressType"):
        if state == "printing":
            return TBPF_ERROR
        if state == "paused":
            return TBPF_PAUSED
        if state == "failed":
            return TBPF_ERROR
        if state in ("finished", "idle"):
            return TBPF_NORMAL
        return TBPF_NOPROGRESS

    pt = PyTaskbar.ProgressType
    if state == "printing":
        return getattr(pt, "ERROR", TBPF_ERROR)
    if state == "paused":
        return getattr(pt, "PAUSED", TBPF_PAUSED)
    if state == "failed":
        return getattr(pt, "ERROR", TBPF_ERROR)
    if state in ("finished", "idle"):
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
        tb_cls = PyTaskbar.TaskbarProgress
        sig = inspect.signature(tb_cls)
        if len(sig.parameters) == 0:
            taskbar = tb_cls()
        else:
            hwnd = None
            try:
                hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            except Exception:
                hwnd = None
            taskbar = tb_cls(hwnd) if hwnd else tb_cls()
        taskbar_ready = True
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
            print(
                "[TASKBAR] methods:", [m for m in dir(taskbar) if not m.startswith("_")]
            )
            taskbar_logged_api = True
        progress_type = _resolve_progress_type_for_state(state)
        _call_taskbar_method(
            [
                "set_progress_type",
                "setProgressType",
                "SetProgressType",
                "set_state",
                "SetState",
            ],
            progress_type,
        )
        if state == "finished":
            _call_taskbar_method(["set_progress", "setProgress", "SetProgress"], 100)
            if hasattr(taskbar, "flash_done"):
                taskbar.flash_done()
        elif state == "idle":
            _call_taskbar_method(["set_progress", "setProgress", "SetProgress"], 100)
        elif state in ("printing", "paused", "failed"):
            _call_taskbar_method(
                ["set_progress", "setProgress", "SetProgress"], clamped
            )
        else:
            _call_taskbar_method(["reset", "Reset"])
    except Exception as e:
        print("[TASKBAR] update failed:", e)


def notify(title: str, message: str, important: bool = False) -> None:
    try:
        toast = Notification(
            app_id=APP_ID,
            title=title,
            msg=message,
            duration="long" if important else "short",
        )
        toast.show()
        if (
            important
            and tray_ready
            and tray_icon is not None
            and hasattr(tray_icon, "notify")
        ):
            try:
                tray_icon.notify(message, title)
            except Exception:
                pass
    except Exception as e:
        print("[TOAST] notify failed:", e)


def choose_primary_printer() -> tuple[str, dict[str, Any]] | None:
    if not printer_states:
        return None

    def rank(item: tuple[str, dict[str, Any]]) -> tuple[int, int]:
        device_id, state_obj = item
        state = human_state(state_obj.get("gcode_state"))
        percent = int(state_obj.get("mc_percent", 0) or 0)
        priority = {
            "failed": 5,
            "paused": 4,
            "printing": 3,
            "preparing": 2,
            "finished": 1,
            "idle": 0,
        }.get(state, -1)
        return (priority, percent)

    return max(printer_states.items(), key=rank)


def get_rotating_printer() -> tuple[str, dict[str, Any]] | None:
    global tray_rotation_index

    if not printer_states:
        return None

    ordered_ids = [
        printer.get("device_id")
        for printer in config_data.get("printers", [])
        if printer.get("device_id") in printer_states
    ]
    if not ordered_ids:
        ordered_ids = list(printer_states.keys())

    tray_rotation_index %= len(ordered_ids)
    device_id = ordered_ids[tray_rotation_index]
    return device_id, printer_states[device_id]


def rotate_tray_target() -> None:
    global tray_rotation_index
    if not printer_states:
        return
    tray_rotation_index = (tray_rotation_index + 1) % len(printer_states)
    update_app_status()


def update_app_status() -> None:
    primary = choose_primary_printer()
    rotating = get_rotating_printer()

    if primary is None:
        display_state = "disconnected" if not mqtt_connected else "unknown"
        update_taskbar_progress(0, display_state)
    else:
        device_id, print_obj = primary
        state = human_state(print_obj.get("gcode_state"))
        percent = int(print_obj.get("mc_percent", 0) or 0)
        task = print_obj.get("subtask_name") or print_obj.get("task_name") or "-"
        label = printer_label(device_id)

        if state == "printing":
            set_console_title(f"{label} {percent}% - {task}")
        elif state == "finished":
            set_console_title(f"{label} Finished - {task}")
        elif state == "idle":
            set_console_title(f"{label} Idle")
        elif state == "failed":
            set_console_title(f"{label} Failed - {task}")
        elif state == "paused":
            set_console_title(f"{label} Paused {percent}% - {task}")
        else:
            set_console_title(f"{label} {state} - {task}")

        update_taskbar_progress(percent, state)

    if rotating is None:
        display_state = "disconnected" if not mqtt_connected else "unknown"
        task_text = (
            "재연결 중..."
            if not mqtt_connected and reconnect_attempts < RECONNECT_MAX_ATTEMPTS
            else ("연결 끊김 - 앱 재시작 필요" if not mqtt_connected else "설정 필요")
        )
        update_tray_icon(0, display_state, task_text, "Bambu Monitor")
        return

    device_id, print_obj = rotating
    state = human_state(print_obj.get("gcode_state"))
    percent = int(print_obj.get("mc_percent", 0) or 0)
    task = print_obj.get("subtask_name") or print_obj.get("task_name") or "-"
    label = printer_label(device_id)
    filament_info = printer_filament_info.get(device_id, {})
    filament_summary = format_filament_summary(filament_info) if filament_info else ""
    update_tray_icon(percent, state, task, label, filament_summary)


def handle_notifications(device_id: str, print_obj: dict[str, Any]) -> None:
    prev = printer_prev.setdefault(device_id, {"state": None, "percent": None})
    state = human_state(print_obj.get("gcode_state"))
    percent = int(print_obj.get("mc_percent", 0) or 0)
    task = print_obj.get("subtask_name") or print_obj.get("task_name") or "Unknown job"
    label = printer_label(device_id)

    if prev["state"] != "printing" and state == "printing":
        notify(
            "출력 시작",
            f"{label}: {task} 출력을 시작했습니다. ({percent}%)",
            important=True,
        )
    if prev["state"] != "finished" and state == "finished":
        notify("출력 완료", f"{label}: {task} 출력이 끝났습니다.", important=True)
    if prev["state"] != "failed" and state == "failed":
        notify("3D Print Failed", f"{label}: {task}", important=True)

    prev_percent = prev["percent"]
    if state == "printing":
        if prev_percent is not None:
            prev_bucket = int(prev_percent) // 25
            curr_bucket = percent // 25
            if curr_bucket > prev_bucket and percent < 100:
                notify("3D Print Progress", f"{label}: {task} {percent}%")

    # 필라멘트 run-out 감지
    if check_runout_error(print_obj):
        if not printer_runout_notified.get(device_id, False):
            filament_type = print_obj.get("filament_type", "")
            type_info = f" ({filament_type})" if filament_type else ""
            notify(
                "필라멘트 부족",
                f"{label}: 필라멘트가 소진되었습니다{type_info}. 교체해주세요.",
                important=True,
            )
            printer_runout_notified[device_id] = True
    else:
        printer_runout_notified[device_id] = False

    prev["state"] = state
    prev["percent"] = percent


def topic_to_device_id(topic: str) -> str | None:
    parts = topic.split("/")
    if len(parts) >= 3 and parts[0] == "device":
        return parts[1]
    return None


def on_connect(client: mqtt.Client, userdata, flags, reason_code, properties=None):
    global reconnect_attempts, mqtt_connected, mqtt_connected_time
    rc_val = reason_code.value if hasattr(reason_code, "value") else reason_code
    if rc_val != 0:
        mqtt_logger.warning("connect failed: rc=%s", rc_val)
        return
    mqtt_logger.info("connected to broker")
    reconnect_attempts = 0
    mqtt_connected = True
    mqtt_connected_time = time.time()
    for printer in config_data.get("printers", []):
        device_id = printer.get("device_id")
        if not device_id:
            continue
        client.subscribe(f"device/{device_id}/report", qos=1)
        request_pushall(client, device_id)
    update_app_status()


def on_message(client: mqtt.Client, userdata, msg: mqtt.MQTTMessage):
    try:
        payload = json.loads(msg.payload.decode("utf-8", errors="ignore"))
    except Exception as e:
        mqtt_logger.warning("invalid json: %s", e)
        return

    print_obj = payload.get("print")
    if not isinstance(print_obj, dict):
        return

    device_id = topic_to_device_id(msg.topic)
    if not device_id:
        return

    state_obj = printer_states.setdefault(device_id, {})
    deep_merge(state_obj, print_obj)

    # 필라멘트/AMS 정보 추출
    filament_info = extract_filament_info(state_obj)
    if filament_info.get("has_ams") or filament_info.get("active_filament_type"):
        printer_filament_info[device_id] = filament_info

    handle_notifications(device_id, state_obj)
    update_app_status()
    mqtt_logger.debug("[%s] %s", printer_label(device_id), summarize_state(state_obj))
    if filament_info.get("has_ams"):
        mqtt_logger.debug(
            "[%s] 필라멘트: %s",
            printer_label(device_id),
            format_filament_summary(filament_info),
        )


def _reconnect_delay(attempt: int) -> float:
    delay = RECONNECT_BASE_DELAY * (2**attempt)
    jitter = delay * 0.1 * (time.time() % 1)
    return min(delay + jitter, RECONNECT_MAX_DELAY)


def _schedule_reconnect(client: mqtt.Client) -> None:
    global reconnect_attempts, mqtt_connected
    if stop_event.is_set():
        return
    reconnect_attempts += 1
    if reconnect_attempts > RECONNECT_MAX_ATTEMPTS:
        mqtt_logger.error(
            "재연결 %d회 시도 후 포기합니다. 설정을 확인하거나 앱을 재시작하세요.",
            RECONNECT_MAX_ATTEMPTS,
        )
        notify(
            "MQTT 연결 끊김",
            f"{RECONNECT_MAX_ATTEMPTS}회 재연결 시도 후 실패했습니다. 설정을 확인하거나 앱을 재시작하세요.",
            important=True,
        )
        mqtt_connected = False
        update_app_status()
        return
    delay = _reconnect_delay(reconnect_attempts)
    mqtt_logger.info(
        "%d/%d회 재연결 시도 (%.1f초 후)...",
        reconnect_attempts,
        RECONNECT_MAX_ATTEMPTS,
        delay,
    )
    update_app_status()
    time.sleep(delay)
    if stop_event.is_set():
        return
    try:
        client.reconnect()
    except Exception as e:
        mqtt_logger.warning("reconnect failed: %s", e)


def on_disconnect(
    client: mqtt.Client, userdata, disconnect_flags, reason_code, properties=None
):
    global mqtt_connected, mqtt_connected_time
    rc_val = reason_code.value if hasattr(reason_code, "value") else reason_code
    mqtt_connected = False
    mqtt_connected_time = None
    if stop_event.is_set():
        mqtt_logger.info("disconnected (shutdown): rc=%s", rc_val)
        return
    mqtt_logger.warning("disconnected: rc=%s", rc_val)
    threading.Thread(target=_schedule_reconnect, args=(client,), daemon=True).start()


def stop_monitor_client() -> None:
    global monitor_client
    with monitor_client_lock:
        client = monitor_client
        monitor_client = None
    if client is None:
        return
    try:
        client.loop_stop()
    except Exception:
        pass
    try:
        client.disconnect()
    except Exception:
        pass


def start_monitor_client() -> None:
    global monitor_client
    validate_config(config_data)

    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
    client.username_pw_set(
        username=f"u_{config_data['user_id']}", password=config_data["access_token"]
    )
    client.tls_set_context(ssl.create_default_context())
    client.on_connect = on_connect
    client.on_message = on_message
    client.on_disconnect = on_disconnect
    client.on_log = lambda client, userdata, level, buf: mqtt_logger.debug(buf)
    client.connect(BROKER, PORT, keepalive=60)
    client.loop_start()

    with monitor_client_lock:
        monitor_client = client


def apply_new_config(cfg: dict[str, Any]) -> None:
    global config_data, printer_states, printer_prev
    with config_lock:
        config_data = cfg
        save_config(config_data)
    printer_states = {}
    printer_prev = {}
    stop_monitor_client()
    start_monitor_client()
    update_app_status()


def _parse_printers_text(text: str) -> list[dict[str, str]]:
    printers = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if "|" in line:
            alias, device_id = [part.strip() for part in line.split("|", 1)]
        else:
            alias, device_id = line, line
        if device_id:
            printers.append({"alias": alias or device_id, "device_id": device_id})
    return printers


def open_settings_dialog(force: bool = False) -> bool:
    current = load_config()
    if not force:
        try:
            validate_config(current)
            return False
        except RuntimeError:
            pass

    result: dict[str, Any] = {"saved": False}
    root = tk.Tk()
    root.title("Bambu Monitor 설정")
    root.resizable(True, True)
    root.attributes("-topmost", True)
    root.geometry("540x420")

    frame = ttk.Frame(root, padding=16)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    frame.rowconfigure(4, weight=1)

    ttk.Label(frame, text="사용자 ID").grid(row=0, column=0, sticky="w", pady=(0, 8))
    user_id_var = tk.StringVar(value=current.get("user_id", ""))
    ttk.Entry(frame, textvariable=user_id_var).grid(
        row=0, column=1, sticky="ew", pady=(0, 8)
    )

    ttk.Label(frame, text="Access Token").grid(row=1, column=0, sticky="w", pady=(0, 8))
    access_token_var = tk.StringVar(value=current.get("access_token", ""))
    ttk.Entry(frame, textvariable=access_token_var, show="*").grid(
        row=1, column=1, sticky="ew", pady=(0, 8)
    )

    ttk.Label(frame, text="이메일(선택)").grid(row=2, column=0, sticky="w", pady=(0, 8))
    email_var = tk.StringVar(value=current.get("email", ""))
    ttk.Entry(frame, textvariable=email_var).grid(
        row=2, column=1, sticky="ew", pady=(0, 8)
    )

    ttk.Label(frame, text="프린터 목록").grid(row=3, column=0, sticky="nw", pady=(0, 8))
    help_text = "한 줄에 하나씩 입력하세요. 형식: 별명 | DEVICE_ID"
    ttk.Label(frame, text=help_text, foreground="#666666").grid(
        row=3, column=1, sticky="w", pady=(0, 8)
    )

    printers_box = tk.Text(frame, height=12)
    printers_box.grid(row=4, column=0, columnspan=2, sticky="nsew")
    printers_box.insert(
        "1.0",
        "\n".join(
            f"{item.get('alias', item.get('name', item['device_id']))} | {item['device_id']}"
            for item in current.get("printers", [])
        ),
    )

    buttons = ttk.Frame(frame)
    buttons.grid(row=5, column=0, columnspan=2, sticky="e", pady=(12, 0))

    def on_save() -> None:
        cfg = {
            "user_id": user_id_var.get().strip(),
            "access_token": access_token_var.get().strip(),
            "email": email_var.get().strip(),
            "printers": _parse_printers_text(printers_box.get("1.0", "end")),
        }
        try:
            validate_config(cfg)
            apply_new_config(cfg)
        except Exception as e:
            messagebox.showerror("설정 오류", str(e), parent=root)
            return

        result["saved"] = True
        root.destroy()

    def on_cancel() -> None:
        root.destroy()

    ttk.Button(buttons, text="취소", command=on_cancel).pack(side="right", padx=(8, 0))
    ttk.Button(buttons, text="저장", command=on_save).pack(side="right")

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return bool(result["saved"])


def settings_watcher() -> None:
    while not stop_event.is_set():
        if settings_request_event.wait(timeout=1):
            settings_request_event.clear()
            try:
                open_settings_dialog(force=True)
            except Exception as e:
                print("[CONFIG] dialog failed:", e)


def tray_rotation_watcher() -> None:
    while not stop_event.is_set():
        for _ in range(5):
            if stop_event.is_set():
                return
            time.sleep(1)
        if stop_event.is_set():
            return
        if len(printer_states) > 1:
            rotate_tray_target()


def main() -> None:
    global config_data

    if not acquire_single_instance_lock():
        return

    stop_event.clear()
    set_app_user_model_id()
    config_data = load_config()
    init_taskbar()
    init_tray_icon()

    threading.Thread(target=settings_watcher, daemon=True).start()
    threading.Thread(target=tray_rotation_watcher, daemon=True).start()

    if not open_settings_dialog(force=False):
        validate_config(config_data)

    start_monitor_client()

    try:
        while not stop_event.is_set():
            for _ in range(60):
                if stop_event.is_set():
                    break
                time.sleep(1)
            if stop_event.is_set():
                break
            if not mqtt_connected:
                mqtt_logger.debug(
                    "skipping pushall: not connected (attempt %d)", reconnect_attempts
                )
                continue
            with monitor_client_lock:
                client = monitor_client
            if client is not None:
                for printer in config_data.get("printers", []):
                    device_id = printer.get("device_id")
                    if device_id:
                        request_pushall(client, device_id)
    except KeyboardInterrupt:
        stop_event.set()
    finally:
        stop_monitor_client()
        shutdown_tray_icon()


if __name__ == "__main__":
    main()
