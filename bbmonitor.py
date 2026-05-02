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
from datetime import datetime
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
_APP_DIR = (
    Path(__import__("sys").executable).resolve().parent
    if getattr(__import__("sys"), "frozen", False)
    else Path(__file__).resolve().parent
)
APP_ICON_PATH = str(_APP_DIR / "app_icon.ico")

TBPF_NOPROGRESS = 0
TBPF_INDETERMINATE = 1
TBPF_NORMAL = 2
TBPF_ERROR = 4
TBPF_PAUSED = 8


config_lock = threading.Lock()
settings_request_event = threading.Event()
status_dashboard_event = threading.Event()
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

    # 사람이 읽기 편한 형태로 포맷
    state_kr = _state_korean(state)
    parts = [f"🖨️ {state_kr}"]

    if state not in ("idle", "disconnected"):
        parts.append(f"{percent}%")

    if remaining and state == "printing":
        time_str = _format_remaining_time(remaining)
        parts.append(f"(남은 시간: {time_str})")

    if task and task != "-":
        parts.append(f"| {task}")

    return " ".join(parts)


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
    # Bambu Lab brand colors
    if state == "printing":
        return (255, 107, 0, 255)  # #FF6B00 Bambu Orange
    if state == "idle":
        return (134, 142, 150, 255)  # #868E96 Gray
    if state == "finished":
        return (81, 207, 102, 255)  # #51CF66 Green
    if state == "paused":
        return (255, 212, 59, 255)  # #FFD43B Yellow
    if state == "failed":
        return (255, 107, 107, 255)  # #FF6B6B Red
    if state == "disconnected":
        return (73, 80, 87, 255)  # #495057 Dark Gray
    return (100, 116, 139, 255)


def _get_tray_font(size: int, bold: bool = True):
    if ImageFont is None:
        return None
    font_names = (
        ["segoeuib.ttf", "seguihis.ttf", "malgunbd.ttf"]
        if bold
        else ["segoeui.ttf", "malgun.ttf"]
    )
    for name in font_names:
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _build_tray_battery_icon(percent: int, state: str):
    if not HAS_TRAY or Image is None or ImageDraw is None:
        return None

    clamped = max(0, min(int(percent), 100))
    display_percent = 100 if state == "idle" else clamped

    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Improved body dimensions with better proportions
    body = (3, 6, 54, 58)
    cap = (54, 22, 60, 42)
    inner = (7, 10, 50, 54)

    # Draw shadow
    draw.rounded_rectangle(
        (body[0] + 1, body[1] + 1, body[2] + 1, body[3] + 1),
        radius=7,
        fill=(0, 0, 0, 40),
    )

    # Draw body with subtle gradient effect (outer)
    draw.rounded_rectangle(
        body, radius=7, outline=(45, 52, 70, 200), width=2, fill=(250, 251, 252, 245)
    )

    # Draw cap with matching style
    draw.rounded_rectangle(cap, radius=3, fill=(45, 52, 70, 200))

    # Determine what to display
    fill_color = _tray_color_for_state(state)
    if state == "disconnected":
        label = "?"
        display_percent = -1
    elif state == "idle":
        label = "✓"
        display_percent = -1
    else:
        label = f"{display_percent}"

    # Draw fill with gradient effect
    if display_percent > 0:
        inner_w = inner[2] - inner[0]
        filled = int(inner_w * (display_percent / 100.0))
        if filled > 0:
            # Main fill
            draw.rounded_rectangle(
                (inner[0], inner[1], inner[0] + filled, inner[3]),
                radius=4,
                fill=fill_color,
            )
            # Highlight gradient on top portion
            highlight_y = inner[1] + ((inner[3] - inner[1]) * 0.3)
            highlight_color = (
                min(fill_color[0] + 40, 255),
                min(fill_color[1] + 40, 255),
                min(fill_color[2] + 40, 255),
                80,
            )
            draw.rectangle(
                (inner[0] + 1, inner[1] + 1, inner[0] + filled - 1, int(highlight_y)),
                fill=highlight_color,
            )
    elif state == "idle":
        # Show full green fill for idle
        draw.rounded_rectangle(inner, radius=4, fill=fill_color)

    # Text rendering with improved stroke effect
    if display_percent == -1:
        # Special symbols for disconnected/idle
        font_size = 28
        font = _get_tray_font(font_size, bold=True)
        cx = (inner[0] + inner[2]) / 2.0
        cy = (inner[1] + inner[3]) / 2.0

        if hasattr(draw, "textbbox"):
            x0, y0, x1, y1 = draw.textbbox((0, 0), label, font=font)
            tx = cx - ((x0 + x1) / 2.0)
            ty = cy - ((y0 + y1) / 2.0)
            # Multi-layer stroke for cleaner look
            for dx, dy in [
                (-1, -1),
                (-1, 1),
                (1, -1),
                (1, 1),
                (0, -1),
                (0, 1),
                (-1, 0),
                (1, 0),
            ]:
                draw.text((tx + dx, ty + dy), label, font=font, fill=(0, 0, 0, 160))
            draw.text((tx, ty), label, font=font, fill=(255, 255, 255, 255))
        else:
            text_w, text_h = draw.textsize(label, font=font)
            tx = cx - (text_w / 2.0)
            ty = cy - (text_h / 2.0)
            for dx, dy in [
                (-1, -1),
                (-1, 1),
                (1, -1),
                (1, 1),
                (0, -1),
                (0, 1),
                (-1, 0),
                (1, 0),
            ]:
                draw.text((tx + dx, ty + dy), label, font=font, fill=(0, 0, 0, 160))
            draw.text((tx, ty), label, font=font, fill=(255, 255, 255, 255))
    else:
        # Number rendering
        font_size = 32
        font = _get_tray_font(font_size, bold=True)
        max_w = (inner[2] - inner[0]) - 6
        max_h = (inner[3] - inner[1]) - 6

        while font_size > 12:
            if hasattr(draw, "textbbox"):
                text_bbox = draw.textbbox((0, 0), label, font=font)
                text_w = text_bbox[2] - text_bbox[0]
                text_h = text_bbox[3] - text_bbox[1]
            else:
                text_w, text_h = draw.textsize(label, font=font)
            if text_w <= max_w and text_h <= max_h:
                break
            font_size -= 2
            font = _get_tray_font(font_size, bold=True)

        cx = (inner[0] + inner[2]) / 2.0
        cy = (inner[1] + inner[3]) / 2.0

        if hasattr(draw, "textbbox"):
            x0, y0, x1, y1 = draw.textbbox((0, 0), label, font=font)
            tx = cx - ((x0 + x1) / 2.0)
            ty = cy - ((y0 + y1) / 2.0)
            # Multi-layer stroke
            for dx, dy in [
                (-1, -1),
                (-1, 1),
                (1, -1),
                (1, 1),
                (0, -1),
                (0, 1),
                (-1, 0),
                (1, 0),
            ]:
                draw.text((tx + dx, ty + dy), label, font=font, fill=(0, 0, 0, 140))
            draw.text((tx, ty), label, font=font, fill=(255, 255, 255, 255))
        else:
            text_w, text_h = draw.textsize(label, font=font)
            tx = cx - (text_w / 2.0)
            ty = cy - (text_h / 2.0)
            for dx, dy in [
                (-1, -1),
                (-1, 1),
                (1, -1),
                (1, 1),
                (0, -1),
                (0, 1),
                (-1, 0),
                (1, 0),
            ]:
                draw.text((tx + dx, ty + dy), label, font=font, fill=(0, 0, 0, 140))
            draw.text((tx, ty), label, font=font, fill=(255, 255, 255, 255))

    return img


def _state_emoji(state: str) -> str:
    """Return emoji for printer state."""
    emojis = {
        "printing": "🖨️",
        "paused": "⏸️",
        "finished": "✅",
        "failed": "❌",
        "idle": "💤",
        "preparing": "⚙️",
        "disconnected": "🔴",
    }
    return emojis.get(state, "❓")


def _state_korean(state: str) -> str:
    """Return Korean label for printer state."""
    labels = {
        "printing": "프린팅 중",
        "paused": "일시정지",
        "finished": "완료",
        "failed": "실패",
        "idle": "유휴",
        "preparing": "준비 중",
        "disconnected": "연결 끊김",
    }
    return labels.get(state, state)


def _format_remaining_time(minutes: int | None) -> str:
    """Format remaining minutes to human-readable string."""
    if minutes is None or minutes <= 0:
        return ""
    if minutes >= 60:
        hours = minutes // 60
        mins = minutes % 60
        return f"{hours}시간 {mins}분"
    return f"{minutes}분"


def on_tray_exit(icon=None, item=None) -> None:
    stop_event.set()
    try:
        if icon is not None:
            icon.stop()
    except Exception:
        pass


def open_settings_from_tray(icon=None, item=None) -> None:
    settings_request_event.set()


_dashboard_window = None
_dashboard_lock = threading.Lock()


def open_status_dashboard(icon=None, item=None) -> None:
    """Open status dashboard popup showing all printer statuses."""
    global _dashboard_window

    with _dashboard_lock:
        if _dashboard_window is not None:
            try:
                _dashboard_window.lift()
                _dashboard_window.focus_force()
                return
            except Exception:
                _dashboard_window = None

    # Run dashboard in main thread via event
    status_dashboard_event.set()


def _create_status_dashboard() -> None:
    """Create and show the status dashboard window."""
    global _dashboard_window

    # Brand colors
    ORANGE = "#FF6B00"
    DARK = "#1A1A2E"
    GRAY_BG = "#F8F9FA"
    GRAY_LIGHT = "#E9ECEF"
    GRAY_MID = "#6C757D"
    GRAY_DARK = "#343A40"
    WHITE = "#FFFFFF"
    GREEN = "#51CF66"
    RED = "#FF6B6B"
    YELLOW = "#FFD43B"
    GRAY = "#868E96"
    DARK_GRAY = "#495057"

    FONT_FAMILY = "Malgun Gothic"

    dash = tk.Toplevel()
    dash.title("Bambu Monitor - 상태 대시보드")
    dash.configure(bg=GRAY_BG)
    dash.geometry("520x600")
    dash.minsize(480, 400)

    # Center window
    dash.update_idletasks()
    screen_w = dash.winfo_screenwidth()
    screen_h = dash.winfo_screenheight()
    win_w, win_h = 520, 600
    x = (screen_w - win_w) // 2
    y = (screen_h - win_h) // 2
    dash.geometry(f"{win_w}x{win_h}+{x}+{y}")

    # Set icon
    try:
        if os.path.exists(APP_ICON_PATH):
            dash.iconbitmap(APP_ICON_PATH)
    except Exception:
        pass

    # Style
    style = ttk.Style(dash)
    style.theme_use("clam")
    style.configure("Dashboard.TFrame", background=GRAY_BG)
    style.configure("DashHeader.TFrame", background=DARK)
    style.configure(
        "DashTitle.TLabel",
        background=DARK,
        foreground=WHITE,
        font=(FONT_FAMILY, 13, "bold"),
    )
    style.configure(
        "DashSubtitle.TLabel",
        background=DARK,
        foreground="#ADB5BD",
        font=(FONT_FAMILY, 9),
    )
    style.configure("Card.TFrame", background=WHITE)
    style.configure(
        "CardTitle.TLabel",
        background=WHITE,
        foreground=DARK,
        font=(FONT_FAMILY, 11, "bold"),
    )
    style.configure(
        "CardInfo.TLabel", background=WHITE, foreground=GRAY_DARK, font=(FONT_FAMILY, 9)
    )
    style.configure(
        "CardValue.TLabel",
        background=WHITE,
        foreground=DARK,
        font=(FONT_FAMILY, 10, "bold"),
    )
    style.configure(
        "StatusGreen.Horizontal.TProgressbar",
        background=GREEN,
        troughcolor=GRAY_LIGHT,
        thickness=8,
        borderwidth=0,
    )
    style.configure(
        "StatusOrange.Horizontal.TProgressbar",
        background=ORANGE,
        troughcolor=GRAY_LIGHT,
        thickness=8,
        borderwidth=0,
    )
    style.configure(
        "StatusRed.Horizontal.TProgressbar",
        background=RED,
        troughcolor=GRAY_LIGHT,
        thickness=8,
        borderwidth=0,
    )
    style.configure(
        "StatusYellow.Horizontal.TProgressbar",
        background=YELLOW,
        troughcolor=GRAY_LIGHT,
        thickness=8,
        borderwidth=0,
    )
    style.configure(
        "StatusGray.Horizontal.TProgressbar",
        background=GRAY,
        troughcolor=GRAY_LIGHT,
        thickness=8,
        borderwidth=0,
    )

    # Header
    header = ttk.Frame(dash, style="DashHeader.TFrame")
    header.pack(fill="x", padx=0, pady=0)
    header.configure(height=60)

    header_inner = ttk.Frame(header, style="DashHeader.TFrame")
    header_inner.pack(fill="both", expand=True, padx=16, pady=(12, 12))

    # Logo
    logo_canvas = tk.Canvas(
        header_inner, width=32, height=32, bg=DARK, highlightthickness=0
    )
    logo_canvas.pack(side="left", padx=(0, 10))
    logo_canvas.create_oval(3, 3, 29, 29, fill=ORANGE, outline="")
    logo_canvas.create_text(16, 16, text="B", fill=WHITE, font=("Segoe UI", 12, "bold"))

    title_frame = ttk.Frame(header_inner, style="DashHeader.TFrame")
    title_frame.pack(side="left", fill="x", expand=True)
    ttk.Label(title_frame, text="상태 대시보드", style="DashTitle.TLabel").pack(
        anchor="w"
    )
    ttk.Label(
        title_frame,
        text="모든 프린터의 현재 상태를 확인합니다",
        style="DashSubtitle.TLabel",
    ).pack(anchor="w", pady=(2, 0))

    # Update time label
    update_time_var = tk.StringVar(value="마지막 업데이트: -")
    update_time_label = tk.Label(
        header_inner,
        textvariable=update_time_var,
        bg=DARK,
        fg="#ADB5BD",
        font=(FONT_FAMILY, 8),
        anchor="e",
    )
    update_time_label.pack(side="right", padx=(0, 0))

    # Scrollable content
    content_canvas = tk.Canvas(dash, bg=GRAY_BG, highlightthickness=0)
    content_scrollbar = ttk.Scrollbar(
        dash, orient="vertical", command=content_canvas.yview
    )
    content_frame = ttk.Frame(content_canvas, style="Dashboard.TFrame")

    content_frame.bind(
        "<Configure>",
        lambda e: content_canvas.configure(scrollregion=content_canvas.bbox("all")),
    )
    content_canvas.create_window(
        (0, 0), window=content_frame, anchor="nw", tags="content_window"
    )
    content_canvas.configure(yscrollcommand=content_scrollbar.set)

    content_canvas.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=12)
    content_scrollbar.pack(side="right", fill="y", padx=(0, 12), pady=12)

    def on_canvas_configure(event):
        content_canvas.itemconfig("content_window", width=event.width)

    content_canvas.bind("<Configure>", on_canvas_configure)

    def _on_mousewheel(event):
        content_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    content_canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # Container for printer cards
    cards_container = ttk.Frame(content_frame, style="Dashboard.TFrame")
    cards_container.pack(fill="both", expand=True)

    state_colors = {
        "printing": (ORANGE, "StatusOrange.Horizontal.TProgressbar"),
        "idle": (GRAY, "StatusGray.Horizontal.TProgressbar"),
        "finished": (GREEN, "StatusGreen.Horizontal.TProgressbar"),
        "failed": (RED, "StatusRed.Horizontal.TProgressbar"),
        "paused": (YELLOW, "StatusYellow.Horizontal.TProgressbar"),
        "disconnected": (DARK_GRAY, "StatusGray.Horizontal.TProgressbar"),
    }

    def refresh_dashboard():
        nonlocal cards_container
        # Clear existing cards
        for widget in cards_container.winfo_children():
            widget.destroy()

        # Update time
        if mqtt_connected_time:
            update_str = datetime.fromtimestamp(mqtt_connected_time).strftime(
                "%H:%M:%S"
            )
            update_time_var.set(f"마지막 업데이트: {update_str}")
        else:
            update_time_var.set("마지막 업데이트: -")

        printers = config_data.get("printers", [])
        if not printers:
            no_printer_frame = tk.Frame(
                cards_container,
                bg=WHITE,
                highlightbackground=GRAY_LIGHT,
                highlightthickness=1,
            )
            no_printer_frame.pack(fill="x", pady=(0, 8), padx=4)
            tk.Label(
                no_printer_frame,
                text="설정된 프린터가 없습니다",
                bg=WHITE,
                fg=GRAY_MID,
                font=(FONT_FAMILY, 10),
            ).pack(pady=20)
            return

        for printer in printers:
            device_id = printer.get("device_id", "")
            if not device_id:
                continue

            alias = printer.get("alias", "") or device_id
            state_obj = printer_states.get(device_id, {})
            state = (
                human_state(state_obj.get("gcode_state"))
                if state_obj
                else "disconnected"
            )
            percent = int(state_obj.get("mc_percent", 0) or 0) if state_obj else 0
            remaining = state_obj.get("mc_remaining_time") if state_obj else None
            task = (
                (state_obj.get("subtask_name") or state_obj.get("task_name") or "")
                if state_obj
                else ""
            )
            nozzle_temp = state_obj.get("nozzle_temp") if state_obj else None
            bed_temp = state_obj.get("bed_temp") if state_obj else None
            fil_info = printer_filament_info.get(device_id, {})

            emoji = _state_emoji(state)
            state_text = _state_korean(state)
            color, progress_style = state_colors.get(
                state, (GRAY, "StatusGray.Horizontal.TProgressbar")
            )
            display_percent = 100 if state == "idle" else percent

            # Card
            card = tk.Frame(
                cards_container,
                bg=WHITE,
                highlightbackground=GRAY_LIGHT,
                highlightthickness=1,
            )
            card.pack(fill="x", pady=(0, 8), padx=4)

            card_inner = tk.Frame(card, bg=WHITE)
            card_inner.pack(fill="x", padx=12, pady=10)

            # Title row
            title_row = tk.Frame(card_inner, bg=WHITE)
            title_row.pack(fill="x", pady=(0, 8))

            tk.Label(
                title_row,
                text=f"{emoji} {alias}",
                bg=WHITE,
                fg=DARK,
                font=(FONT_FAMILY, 11, "bold"),
            ).pack(side="left")

            state_badge = tk.Label(
                title_row,
                text=f" {state_text} ",
                bg=color,
                fg=WHITE,
                font=(FONT_FAMILY, 8, "bold"),
            )
            state_badge.pack(side="right")

            # Progress bar
            if state not in ("idle", "disconnected"):
                progress_frame = tk.Frame(card_inner, bg=WHITE)
                progress_frame.pack(fill="x", pady=(0, 6))

                progress_bar = ttk.Progressbar(
                    progress_frame,
                    style=progress_style,
                    length=200,
                    maximum=100,
                    value=display_percent,
                )
                progress_bar.pack(side="left", fill="x", expand=True)

                percent_label = tk.Label(
                    progress_frame,
                    text=f"{display_percent}%",
                    bg=WHITE,
                    fg=DARK,
                    font=(FONT_FAMILY, 9, "bold"),
                    width=5,
                    anchor="e",
                )
                percent_label.pack(side="right", padx=(8, 0))

            # Info grid
            info_frame = tk.Frame(card_inner, bg=WHITE)
            info_frame.pack(fill="x", pady=(0, 4))

            row = 0

            # Task name
            if task:
                tk.Label(
                    info_frame,
                    text="작업:",
                    bg=WHITE,
                    fg=GRAY_MID,
                    font=(FONT_FAMILY, 8),
                    anchor="w",
                ).grid(row=row, column=0, sticky="w", pady=1)
                tk.Label(
                    info_frame,
                    text=task[:40],
                    bg=WHITE,
                    fg=DARK,
                    font=(FONT_FAMILY, 9),
                    anchor="w",
                ).grid(row=row, column=1, sticky="w", padx=(8, 0), pady=1)
                row += 1

            # Remaining time
            if remaining and state == "printing":
                time_str = _format_remaining_time(remaining)
                tk.Label(
                    info_frame,
                    text="남은 시간:",
                    bg=WHITE,
                    fg=GRAY_MID,
                    font=(FONT_FAMILY, 8),
                    anchor="w",
                ).grid(row=row, column=0, sticky="w", pady=1)
                tk.Label(
                    info_frame,
                    text=time_str,
                    bg=WHITE,
                    fg=ORANGE,
                    font=(FONT_FAMILY, 9, "bold"),
                    anchor="w",
                ).grid(row=row, column=1, sticky="w", padx=(8, 0), pady=1)
                row += 1

            # Temperatures
            if nozzle_temp and bed_temp:
                tk.Label(
                    info_frame,
                    text="온도:",
                    bg=WHITE,
                    fg=GRAY_MID,
                    font=(FONT_FAMILY, 8),
                    anchor="w",
                ).grid(row=row, column=0, sticky="w", pady=1)
                temp_text = f"노즐 {nozzle_temp}°C / 베드 {bed_temp}°C"
                tk.Label(
                    info_frame,
                    text=temp_text,
                    bg=WHITE,
                    fg=DARK,
                    font=(FONT_FAMILY, 9),
                    anchor="w",
                ).grid(row=row, column=1, sticky="w", padx=(8, 0), pady=1)
                row += 1

            # Filament info
            if fil_info.get("active_filament_type"):
                fil_type = fil_info["active_filament_type"]
                fil_color = fil_info.get("active_filament_color", "")
                ams_info = ""
                if fil_info.get("has_ams"):
                    active_tray = fil_info.get("active_tray")
                    if active_tray is not None:
                        ams_info = f" (AMS 트레이 {active_tray})"
                color_info = f" [{fil_color}]" if fil_color else ""
                tk.Label(
                    info_frame,
                    text="필라멘트:",
                    bg=WHITE,
                    fg=GRAY_MID,
                    font=(FONT_FAMILY, 8),
                    anchor="w",
                ).grid(row=row, column=0, sticky="w", pady=1)
                tk.Label(
                    info_frame,
                    text=f"{fil_type}{color_info}{ams_info}",
                    bg=WHITE,
                    fg=DARK,
                    font=(FONT_FAMILY, 9),
                    anchor="w",
                ).grid(row=row, column=1, sticky="w", padx=(8, 0), pady=1)

        # Schedule next refresh
        if dash.winfo_exists():
            dash.after(3000, refresh_dashboard)

    def on_close():
        global _dashboard_window
        try:
            content_canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        _dashboard_window = None
        dash.destroy()

    dash.protocol("WM_DELETE_WINDOW", on_close)
    _dashboard_window = dash

    # Initial refresh
    refresh_dashboard()


def _build_tray_menu() -> pystray.Menu:
    """Build dynamic context menu based on current printer states."""
    menu_items = []

    # Header
    menu_items.append(pystray.MenuItem("Bambu Monitor v1.0", None, enabled=False))
    menu_items.append(pystray.Menu.SEPARATOR)

    # Current primary printer status
    primary = choose_primary_printer()
    if primary:
        device_id, state_obj = primary
        state = human_state(state_obj.get("gcode_state"))
        percent = int(state_obj.get("mc_percent", 0) or 0)
        label = printer_label(device_id)
        remaining = state_obj.get("mc_remaining_time")
        task = state_obj.get("subtask_name") or state_obj.get("task_name") or ""
        emoji = _state_emoji(state)
        state_text = _state_korean(state)

        if state == "printing":
            time_str = _format_remaining_time(remaining)
            time_info = f" (남은 시간: {time_str})" if time_str else ""
            status_line = f"{emoji} {label}: {state_text} ({percent}%){time_info}"
        else:
            status_line = f"{emoji} {label}: {state_text}"
        if task:
            status_line += f"\n   작업: {task[:25]}"
        menu_items.append(pystray.MenuItem(status_line, None, enabled=False))
    else:
        if not mqtt_connected and reconnect_attempts >= RECONNECT_MAX_ATTEMPTS:
            menu_items.append(
                pystray.MenuItem("🔴 연결 끊김 - 앱 재시작 필요", None, enabled=False)
            )
        elif not mqtt_connected:
            menu_items.append(
                pystray.MenuItem(
                    f"🟡 재연결 중 ({reconnect_attempts}/{RECONNECT_MAX_ATTEMPTS})",
                    None,
                    enabled=False,
                )
            )
        else:
            menu_items.append(
                pystray.MenuItem("⚪ 프린터 대기 중...", None, enabled=False)
            )
    menu_items.append(pystray.Menu.SEPARATOR)

    # Individual printer submenu
    printer_menu_items = []
    for printer in config_data.get("printers", []):
        device_id = printer.get("device_id", "")
        if not device_id:
            continue

        alias = printer.get("alias", "") or device_id
        state_obj = printer_states.get(device_id, {})
        state = human_state(state_obj.get("gcode_state"))
        percent = int(state_obj.get("mc_percent", 0) or 0)
        emoji = _state_emoji(state) if state_obj else "⚪"
        state_text = _state_korean(state) if state_obj else "미연결"

        if state == "printing":
            item_text = f"{emoji} {alias} | {percent}% {state_text}"
        else:
            item_text = f"{emoji} {alias} | {state_text}"

        def make_printer_action(did: str):
            def action(icon=None, item=None):
                pass

            return action

        printer_menu_items.append(
            pystray.MenuItem(item_text, make_printer_action(device_id))
        )

    if printer_menu_items:
        menu_items.append(
            pystray.MenuItem("프린터 목록", pystray.Menu(*printer_menu_items))
        )
        menu_items.append(pystray.Menu.SEPARATOR)

    # Tools section
    menu_items.append(pystray.MenuItem("📊 상태 보기", open_status_dashboard))
    menu_items.append(
        pystray.MenuItem("🚀 Bambu Studio 열기", open_bambu_studio, default=True)
    )
    menu_items.append(pystray.MenuItem("⚙️ 설정 열기", open_settings_from_tray))
    menu_items.append(pystray.Menu.SEPARATOR)
    menu_items.append(pystray.MenuItem("종료", on_tray_exit))

    return pystray.Menu(*menu_items)


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
        menu = _build_tray_menu()
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
    # Build enhanced tooltip
    primary = choose_primary_printer()
    tooltip_parts = ["Bambu Monitor"]
    tooltip_parts.append("━━━━━━━━━━━━━━━━")

    if primary:
        device_id, state_obj = primary
        p_label = printer_label(device_id)
        p_state = human_state(state_obj.get("gcode_state"))
        p_percent = (
            int(state_obj.get("mc_percent", 0) or 0)
            if p_state not in ("idle", "disconnected")
            else (100 if p_state == "idle" else 0)
        )
        p_remaining = state_obj.get("mc_remaining_time")
        p_task = state_obj.get("subtask_name") or state_obj.get("task_name") or ""
        p_nozzle = state_obj.get("nozzle_temp")
        p_bed = state_obj.get("bed_temp")
        fil_info = printer_filament_info.get(device_id, {})

        tooltip_parts.append(f"프린터: {p_label}")
        tooltip_parts.append(f"상태: {_state_korean(p_state)} ({p_percent}%)")

        if p_remaining and p_state == "printing":
            tooltip_parts.append(f"남은 시간: {_format_remaining_time(p_remaining)}")

        if p_task:
            tooltip_parts.append(f"작업: {p_task[:30]}")

        tooltip_parts.append("━━━━━━━━━━━━━━━━")

        if fil_info.get("active_filament_type"):
            fil_type = fil_info["active_filament_type"]
            fil_color = fil_info.get("active_filament_color", "")
            color_info = f" ({fil_color})" if fil_color else ""
            tooltip_parts.append(f"필라멘트: {fil_type}{color_info}")

        if p_nozzle and p_bed:
            tooltip_parts.append(f"온도: 노즐 {p_nozzle}°C / 베드 {p_bed}°C")
    else:
        if not mqtt_connected:
            tooltip_parts.append("프린터 연결 대기 중...")
        else:
            tooltip_parts.append("설정된 프린터 없음")

    if filament_summary:
        tooltip_parts.append(filament_summary)

    tooltip_text = "\n".join(tooltip_parts)

    try:
        with tray_lock:
            tray_icon.icon = icon_image
            tray_icon.title = tooltip_text
            # Update dynamic menu
            try:
                tray_icon.menu = _build_tray_menu()
            except Exception as menu_err:
                mqtt_logger.debug("menu update failed: %s", menu_err)
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


def notify(
    title: str,
    message: str,
    important: bool = False,
    icon_path: str | None = None,
) -> None:
    try:
        # Use provided icon, app icon, or None
        notify_icon = icon_path
        if not notify_icon:
            if os.path.exists(APP_ICON_PATH):
                notify_icon = APP_ICON_PATH

        toast = Notification(
            app_id=APP_ID,
            title=title,
            msg=message,
            duration="long" if important else "short",
            icon=notify_icon,
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
    task = (
        print_obj.get("subtask_name") or print_obj.get("task_name") or "알 수 없는 작업"
    )
    label = printer_label(device_id)
    remaining = print_obj.get("mc_remaining_time")
    filament_info = printer_filament_info.get(device_id, {})

    if prev["state"] != "printing" and state == "printing":
        # 예상 소요 시간 포함
        time_info = ""
        if remaining and remaining > 0:
            time_info = f"\n예상 소요 시간: {_format_remaining_time(remaining)}"
        notify(
            "🖨️ 출력 시작",
            f"{label}: {task} 출력을 시작했습니다. ({percent}%){time_info}",
            important=True,
        )
    if prev["state"] != "finished" and state == "finished":
        notify(
            "✅ 출력 완료",
            f"{label}: {task} 출력이 성공적으로 완료되었습니다.",
            important=True,
        )
    if prev["state"] != "failed" and state == "failed":
        print_error = print_obj.get("print_error", "")
        error_info = f"\n오류 코드: {print_error}" if print_error else ""
        notify(
            "❌ 출력 실패",
            f"{label}: {task} 출력에 실패했습니다.{error_info}",
            important=True,
        )

    prev_percent = prev["percent"]
    if state == "printing":
        if prev_percent is not None:
            prev_bucket = int(prev_percent) // 25
            curr_bucket = percent // 25
            if curr_bucket > prev_bucket and percent < 100:
                # 진행률 마일스톤 알림에 남은 시간 포함
                time_info = ""
                if remaining and remaining > 0:
                    time_info = f" (남은 시간: {_format_remaining_time(remaining)})"
                notify(
                    f"📊 진행률 {percent}%",
                    f"{label}: {task}{time_info}",
                )

    # 필라멘트 run-out 감지
    if check_runout_error(print_obj):
        if not printer_runout_notified.get(device_id, False):
            filament_type = print_obj.get("filament_type", "")
            # AMS 트레이 정보 포함
            ams_info = ""
            if filament_info.get("has_ams"):
                active_tray = filament_info.get("active_tray")
                if active_tray is not None:
                    ams_info = f" (AMS 트레이 {active_tray})"
            type_info = f" ({filament_type})" if filament_type else ""
            notify(
                "⚠️ 필라멘트 부족",
                f"{label}: 필라멘트가 소진되었습니다{type_info}{ams_info}. 교체해주세요.",
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

    result: dict[str, Any] = {"saved": False, "connect": False}

    # ── Color Constants ──
    ORANGE = "#FF6B00"
    DARK = "#1A1A2E"
    GRAY_BG = "#F8F9FA"
    GRAY_LIGHT = "#E9ECEF"
    GRAY_MID = "#6C757D"
    GRAY_DARK = "#343A40"
    WHITE = "#FFFFFF"
    GREEN = "#28A745"
    RED = "#DC3545"
    ORANGE_LIGHT = "#FFF3E0"

    # ── Font Setup ──
    FONT_FAMILY = "Malgun Gothic"
    FONT_TITLE = (FONT_FAMILY, 14, "bold")
    FONT_SUBTITLE = (FONT_FAMILY, 9)
    FONT_LABEL = (FONT_FAMILY, 10)
    FONT_HINT = (FONT_FAMILY, 8)
    FONT_STEP = (FONT_FAMILY, 9, "bold")
    FONT_BUTTON = (FONT_FAMILY, 10)
    FONT_TREE = (FONT_FAMILY, 9)

    # ── Window Setup ──
    root = tk.Tk()
    root.title("Bambu Monitor 설정")
    root.configure(bg=GRAY_BG)
    root.resizable(True, True)
    root.attributes("-topmost", True)
    root.minsize(600, 580)

    # Center window
    root.update_idletasks()
    win_w, win_h = 620, 620
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = (screen_w - win_w) // 2
    y = (screen_h - win_h) // 2
    root.geometry(f"{win_w}x{win_h}+{x}+{y}")

    # ── Style Configuration ──
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure("Main.TFrame", background=GRAY_BG)
    style.configure("Header.TFrame", background=DARK)
    style.configure("Step.TFrame", background=WHITE)
    style.configure("Field.TFrame", background=WHITE)
    style.configure("Button.TFrame", background=GRAY_BG)

    style.configure("Title.TLabel", background=DARK, foreground=WHITE, font=FONT_TITLE)
    style.configure(
        "Subtitle.TLabel", background=DARK, foreground="#ADB5BD", font=FONT_SUBTITLE
    )
    style.configure(
        "StepActive.TLabel", background=ORANGE, foreground=WHITE, font=FONT_STEP
    )
    style.configure(
        "StepInactive.TLabel",
        background=GRAY_LIGHT,
        foreground=GRAY_MID,
        font=FONT_STEP,
    )
    style.configure(
        "Field.TLabel", background=WHITE, foreground=GRAY_DARK, font=FONT_LABEL
    )
    style.configure(
        "Hint.TLabel", background=WHITE, foreground=GRAY_MID, font=FONT_HINT
    )
    style.configure(
        "Section.TLabel",
        background=WHITE,
        foreground=DARK,
        font=(FONT_FAMILY, 11, "bold"),
    )

    style.configure(
        "Field.TEntry",
        fieldbackground=WHITE,
        background=WHITE,
        foreground=GRAY_DARK,
        bordercolor=GRAY_LIGHT,
        lightcolor=GRAY_LIGHT,
        darkcolor=GRAY_LIGHT,
        focusthickness=2,
        focuscolor=ORANGE,
    )
    style.map(
        "Field.TEntry",
        fieldbackground=[("focus", WHITE)],
        bordercolor=[("focus", ORANGE)],
    )

    style.configure(
        "Orange.TButton",
        background=ORANGE,
        foreground=WHITE,
        font=FONT_BUTTON,
        borderwidth=0,
        focuscolor=ORANGE,
        padding=(20, 8),
    )
    style.map(
        "Orange.TButton",
        background=[("active", "#E55D00"), ("pressed", "#CC5200")],
        foreground=[("active", WHITE)],
    )

    style.configure(
        "Secondary.TButton",
        background=GRAY_LIGHT,
        foreground=GRAY_DARK,
        font=FONT_BUTTON,
        borderwidth=0,
        focuscolor=GRAY_LIGHT,
        padding=(16, 8),
    )
    style.map(
        "Secondary.TButton",
        background=[("active", "#DEE2E6"), ("pressed", "#CED4DA")],
        foreground=[("active", GRAY_DARK)],
    )

    style.configure(
        "Danger.TButton",
        background=RED,
        foreground=WHITE,
        font=(FONT_FAMILY, 9),
        borderwidth=0,
        focuscolor=RED,
        padding=(8, 4),
    )
    style.map(
        "Danger.TButton", background=[("active", "#C82333"), ("pressed", "#BD2130")]
    )

    style.configure(
        "Small.TButton",
        background=ORANGE,
        foreground=WHITE,
        font=(FONT_FAMILY, 9),
        borderwidth=0,
        focuscolor=ORANGE,
        padding=(8, 4),
    )
    style.map(
        "Small.TButton", background=[("active", "#E55D00"), ("pressed", "#CC5200")]
    )

    style.configure(
        "Printer.Treeview",
        background=WHITE,
        foreground=GRAY_DARK,
        fieldbackground=WHITE,
        font=FONT_TREE,
        rowheight=28,
        borderwidth=0,
    )
    style.configure(
        "Printer.Treeview.Heading",
        background=GRAY_LIGHT,
        foreground=GRAY_DARK,
        font=(FONT_FAMILY, 9, "bold"),
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "Printer.Treeview",
        background=[("selected", ORANGE_LIGHT)],
        foreground=[("selected", DARK)],
    )
    style.map("Printer.Treeview.Heading", background=[("active", "#DEE2E6")])

    # ── Main Container ──
    main_frame = ttk.Frame(root, style="Main.TFrame")
    main_frame.pack(fill="both", expand=True)

    # ── Header Section ──
    header = ttk.Frame(main_frame, style="Header.TFrame")
    header.pack(fill="x", padx=0, pady=0)
    header.configure(height=70)

    header_inner = ttk.Frame(header, style="Header.TFrame")
    header_inner.pack(fill="both", expand=True, padx=20, pady=(16, 16))

    # Logo icon (drawn with canvas)
    logo_canvas = tk.Canvas(
        header_inner, width=36, height=36, bg=DARK, highlightthickness=0
    )
    logo_canvas.pack(side="left", padx=(0, 12))
    logo_canvas.create_oval(4, 4, 32, 32, fill=ORANGE, outline="")
    logo_canvas.create_text(18, 18, text="B", fill=WHITE, font=("Segoe UI", 14, "bold"))

    title_frame = ttk.Frame(header_inner, style="Header.TFrame")
    title_frame.pack(side="left", fill="x", expand=True)
    ttk.Label(title_frame, text="Bambu Monitor 설정", style="Title.TLabel").pack(
        anchor="w"
    )
    ttk.Label(
        title_frame,
        text="프린터 연결을 위한 인증 정보와 프린터를 설정합니다",
        style="Subtitle.TLabel",
    ).pack(anchor="w", pady=(2, 0))

    # ── Step Indicators ──
    step_frame = ttk.Frame(main_frame, style="Main.TFrame")
    step_frame.pack(fill="x", padx=20, pady=(16, 8))

    step1_frame = tk.Frame(step_frame, bg=ORANGE, padx=12, pady=4)
    step1_frame.pack(side="left")
    tk.Label(
        step1_frame, text="1. 인증 정보", bg=ORANGE, fg=WHITE, font=FONT_STEP
    ).pack()

    step_connector = tk.Frame(step_frame, bg=GRAY_LIGHT, height=2, width=30)
    step_connector.pack(side="left", fill="y", padx=(8, 8), pady=8)

    step2_frame = tk.Frame(step_frame, bg=GRAY_LIGHT, padx=12, pady=4)
    step2_frame.pack(side="left")
    tk.Label(
        step2_frame, text="2. 프린터 목록", bg=GRAY_LIGHT, fg=GRAY_MID, font=FONT_STEP
    ).pack()

    # ── Content Area with Scroll ──
    content_canvas = tk.Canvas(main_frame, bg=GRAY_BG, highlightthickness=0)
    content_scrollbar = ttk.Scrollbar(
        main_frame, orient="vertical", command=content_canvas.yview
    )
    content_frame = ttk.Frame(content_canvas, style="Main.TFrame")

    content_frame.bind(
        "<Configure>",
        lambda e: content_canvas.configure(scrollregion=content_canvas.bbox("all")),
    )
    content_canvas.create_window(
        (0, 0), window=content_frame, anchor="nw", tags="content_window"
    )
    content_canvas.configure(yscrollcommand=content_scrollbar.set)

    content_canvas.pack(side="left", fill="both", expand=True, padx=(20, 0), pady=8)
    content_scrollbar.pack(side="right", fill="y", padx=(0, 20), pady=8)

    def on_canvas_configure(event):
        content_canvas.itemconfig("content_window", width=event.width)

    content_canvas.bind("<Configure>", on_canvas_configure)

    def _on_mousewheel(event):
        content_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    content_canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # ── Auth Section ──
    auth_card = tk.Frame(
        content_frame, bg=WHITE, highlightbackground=GRAY_LIGHT, highlightthickness=1
    )
    auth_card.pack(fill="x", pady=(0, 12))

    auth_inner = tk.Frame(auth_card, bg=WHITE)
    auth_inner.pack(fill="x", padx=16, pady=12)

    tk.Label(
        auth_inner, text="인증 정보", bg=WHITE, fg=DARK, font=(FONT_FAMILY, 11, "bold")
    ).pack(anchor="w", pady=(0, 12))

    # User ID Field
    user_id_container = tk.Frame(auth_inner, bg=WHITE)
    user_id_container.pack(fill="x", pady=(0, 12))
    tk.Label(
        user_id_container, text="사용자 ID *", bg=WHITE, fg=GRAY_DARK, font=FONT_LABEL
    ).pack(anchor="w")
    tk.Label(
        user_id_container,
        text="Bambu Lab 웹 또는 앱에서 확인할 수 있는 고유 사용자 식별자입니다",
        bg=WHITE,
        fg=GRAY_MID,
        font=FONT_HINT,
    ).pack(anchor="w", pady=(2, 4))
    user_id_var = tk.StringVar(value=current.get("user_id", ""))
    user_id_entry = ttk.Entry(
        user_id_container, textvariable=user_id_var, style="Field.TEntry"
    )
    user_id_entry.pack(fill="x")

    # Access Token Field
    token_container = tk.Frame(auth_inner, bg=WHITE)
    token_container.pack(fill="x", pady=(0, 12))
    tk.Label(
        token_container, text="Access Token *", bg=WHITE, fg=GRAY_DARK, font=FONT_LABEL
    ).pack(anchor="w")
    tk.Label(
        token_container,
        text="Bambu Lab API 인증에 사용되는 토큰입니다 (입력 시 ••••로 표시됨)",
        bg=WHITE,
        fg=GRAY_MID,
        font=FONT_HINT,
    ).pack(anchor="w", pady=(2, 4))
    access_token_var = tk.StringVar(value=current.get("access_token", ""))
    token_entry = ttk.Entry(
        token_container, textvariable=access_token_var, style="Field.TEntry", show="•"
    )
    token_entry.pack(fill="x")

    # Toggle password visibility
    show_password_var = tk.BooleanVar(value=False)

    def toggle_password_visibility():
        if show_password_var.get():
            token_entry.configure(show="")
        else:
            token_entry.configure(show="•")

    pw_check = tk.Checkbutton(
        token_container,
        text="토큰 표시",
        variable=show_password_var,
        command=toggle_password_visibility,
        bg=WHITE,
        fg=GRAY_MID,
        selectcolor=WHITE,
        activebackground=WHITE,
        activeforeground=GRAY_MID,
        font=FONT_HINT,
    )
    pw_check.pack(anchor="w", pady=(4, 0))

    # Email Field
    email_container = tk.Frame(auth_inner, bg=WHITE)
    email_container.pack(fill="x")
    tk.Label(
        email_container, text="이메일 (선택)", bg=WHITE, fg=GRAY_DARK, font=FONT_LABEL
    ).pack(anchor="w")
    tk.Label(
        email_container,
        text="알림 수신 등에 사용될 수 있는 이메일 주소입니다",
        bg=WHITE,
        fg=GRAY_MID,
        font=FONT_HINT,
    ).pack(anchor="w", pady=(2, 4))
    email_var = tk.StringVar(value=current.get("email", ""))
    email_entry = ttk.Entry(
        email_container, textvariable=email_var, style="Field.TEntry"
    )
    email_entry.pack(fill="x")

    # ── Printer List Section ──
    printer_card = tk.Frame(
        content_frame, bg=WHITE, highlightbackground=GRAY_LIGHT, highlightthickness=1
    )
    printer_card.pack(fill="x", pady=(0, 12))

    printer_inner = tk.Frame(printer_card, bg=WHITE)
    printer_inner.pack(fill="x", padx=16, pady=12)

    printer_header = tk.Frame(printer_inner, bg=WHITE)
    printer_header.pack(fill="x", pady=(0, 8))

    tk.Label(
        printer_header,
        text="프린터 목록 *",
        bg=WHITE,
        fg=DARK,
        font=(FONT_FAMILY, 11, "bold"),
    ).pack(side="left")
    tk.Label(
        printer_header,
        text="모니터링할 프린터를 추가하세요 (최소 1대)",
        bg=WHITE,
        fg=GRAY_MID,
        font=FONT_HINT,
    ).pack(side="left", padx=(8, 0))

    # Printer Treeview
    tree_frame = tk.Frame(printer_inner, bg=GRAY_LIGHT)
    tree_frame.pack(fill="both", expand=True, pady=(0, 8))

    columns = ("alias", "device_id", "status")
    printer_tree = ttk.Treeview(
        tree_frame, columns=columns, show="headings", height=4, style="Printer.Treeview"
    )
    printer_tree.heading("alias", text="별칭", anchor="w")
    printer_tree.heading("device_id", text="Device ID", anchor="w")
    printer_tree.heading("status", text="상태", anchor="center")

    printer_tree.column("alias", width=150, minwidth=100, stretch=True)
    printer_tree.column("device_id", width=200, minwidth=150, stretch=True)
    printer_tree.column("status", width=80, minwidth=70, stretch=False)

    tree_scroll = ttk.Scrollbar(
        tree_frame, orient="vertical", command=printer_tree.yview
    )
    printer_tree.configure(yscrollcommand=tree_scroll.set)

    printer_tree.pack(side="left", fill="both", expand=True)
    tree_scroll.pack(side="right", fill="y")

    # Populate tree
    def get_printer_status(device_id: str) -> tuple[str, str]:
        if device_id in printer_states:
            state_raw = printer_states[device_id].get("gcode_state", "")
            if state_raw:
                return ("연결됨", GREEN)
        return ("미연결", GRAY_MID)

    def refresh_tree():
        for item in printer_tree.get_children():
            printer_tree.delete(item)
        for printer in current.get("printers", []):
            device_id = printer.get("device_id", "")
            alias = printer.get("alias", "") or device_id
            status_text, status_color = get_printer_status(device_id)
            printer_tree.insert(
                "",
                "end",
                iid=device_id,
                values=(alias, device_id, status_text),
                tags=(status_color,),
            )
        printer_tree.tag_configure(GREEN, foreground=GREEN)
        printer_tree.tag_configure(GRAY_MID, foreground=GRAY_MID)

    refresh_tree()

    # Printer Edit Controls
    printer_edit_frame = tk.Frame(printer_inner, bg=WHITE)
    printer_edit_frame.pack(fill="x", pady=(0, 4))

    edit_inner = tk.Frame(printer_edit_frame, bg=WHITE)
    edit_inner.pack(fill="x")

    tk.Label(edit_inner, text="별칭:", bg=WHITE, fg=GRAY_DARK, font=FONT_HINT).grid(
        row=0, column=0, padx=(0, 4), sticky="w"
    )
    alias_entry_var = tk.StringVar()
    alias_entry = ttk.Entry(
        edit_inner, textvariable=alias_entry_var, style="Field.TEntry", width=18
    )
    alias_entry.grid(row=0, column=1, padx=(0, 8))

    tk.Label(
        edit_inner, text="Device ID:", bg=WHITE, fg=GRAY_DARK, font=FONT_HINT
    ).grid(row=0, column=2, padx=(0, 4), sticky="w")
    device_entry_var = tk.StringVar()
    device_entry = ttk.Entry(
        edit_inner, textvariable=device_entry_var, style="Field.TEntry", width=24
    )
    device_entry.grid(row=0, column=3, padx=(0, 8))

    btn_frame = tk.Frame(edit_inner, bg=WHITE)
    btn_frame.grid(row=0, column=4)

    def add_printer():
        device_id = device_entry_var.get().strip()
        alias = alias_entry_var.get().strip() or device_id
        if not device_id:
            messagebox.showwarning("입력 오류", "Device ID를 입력해주세요", parent=root)
            return
        if device_id in [
            printer_tree.item(i)["values"][1] for i in printer_tree.get_children()
        ]:
            messagebox.showwarning(
                "중복 오류", "이미 추가된 Device ID입니다", parent=root
            )
            return
        printer_tree.insert(
            "",
            "end",
            iid=device_id,
            values=(alias, device_id, "미연결"),
            tags=(GRAY_MID,),
        )
        printer_tree.tag_configure(GRAY_MID, foreground=GRAY_MID)
        alias_entry_var.set("")
        device_entry_var.set("")
        alias_entry.focus_set()

    def delete_printer():
        selected = printer_tree.selection()
        if not selected:
            messagebox.showinfo(
                "선택 필요", "삭제할 프린터를 선택해주세요", parent=root
            )
            return
        for item in selected:
            printer_tree.delete(item)

    def edit_printer():
        selected = printer_tree.selection()
        if not selected:
            messagebox.showinfo(
                "선택 필요", "수정할 프린터를 선택해주세요", parent=root
            )
            return
        item = selected[0]
        values = printer_tree.item(item)["values"]
        alias_entry_var.set(values[0])
        device_entry_var.set(values[1])
        printer_tree.delete(item)

    ttk.Button(btn_frame, text="추가", style="Small.TButton", command=add_printer).pack(
        side="left", padx=(0, 4)
    )
    ttk.Button(
        btn_frame, text="수정", style="Secondary.TButton", command=edit_printer
    ).pack(side="left", padx=(0, 4))
    ttk.Button(
        btn_frame, text="삭제", style="Danger.TButton", command=delete_printer
    ).pack(side="left")

    # ── Buttons Section ──
    button_frame = tk.Frame(content_frame, bg=GRAY_BG)
    button_frame.pack(fill="x", pady=(4, 8))

    def get_printers_from_tree() -> list[dict[str, str]]:
        printers = []
        for item in printer_tree.get_children():
            values = printer_tree.item(item)["values"]
            alias, device_id = values[0], values[1]
            if device_id:
                printers.append({"alias": alias or device_id, "device_id": device_id})
        return printers

    def on_save_and_connect() -> None:
        cfg = {
            "user_id": user_id_var.get().strip(),
            "access_token": access_token_var.get().strip(),
            "email": email_var.get().strip(),
            "printers": get_printers_from_tree(),
        }
        try:
            validate_config(cfg)
            apply_new_config(cfg)
        except Exception as e:
            messagebox.showerror("설정 오류", str(e), parent=root)
            return
        result["saved"] = True
        result["connect"] = True
        root.destroy()

    def on_save() -> None:
        cfg = {
            "user_id": user_id_var.get().strip(),
            "access_token": access_token_var.get().strip(),
            "email": email_var.get().strip(),
            "printers": get_printers_from_tree(),
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

    # Right-aligned buttons
    tk.Button(
        button_frame,
        text="취소",
        bg=GRAY_LIGHT,
        fg=GRAY_DARK,
        font=FONT_BUTTON,
        relief="flat",
        padx=20,
        pady=8,
        cursor="hand2",
        activebackground="#DEE2E6",
        activeforeground=GRAY_DARK,
        command=on_cancel,
    ).pack(side="right", padx=(8, 0))

    tk.Button(
        button_frame,
        text="저장 후 연결",
        bg=ORANGE,
        fg=WHITE,
        font=FONT_BUTTON,
        relief="flat",
        padx=20,
        pady=8,
        cursor="hand2",
        activebackground="#E55D00",
        activeforeground=WHITE,
        command=on_save_and_connect,
    ).pack(side="right", padx=(8, 0))

    tk.Button(
        button_frame,
        text="저장",
        bg=GRAY_DARK,
        fg=WHITE,
        font=FONT_BUTTON,
        relief="flat",
        padx=20,
        pady=8,
        cursor="hand2",
        activebackground="#23272B",
        activeforeground=WHITE,
        command=on_save,
    ).pack(side="right")

    # ── Cleanup and Run ──
    def on_close():
        try:
            content_canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
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


def status_dashboard_watcher() -> None:
    while not stop_event.is_set():
        if status_dashboard_event.wait(timeout=1):
            status_dashboard_event.clear()
            try:
                _create_status_dashboard()
            except Exception as e:
                print("[DASHBOARD] failed:", e)


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
    threading.Thread(target=status_dashboard_watcher, daemon=True).start()

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
