import json
import os
import ssl
import time
from typing import Any

import paho.mqtt.client as mqtt
from dotenv import load_dotenv

load_dotenv()

USER_ID = os.getenv("BBL_USER_ID")
ACCESS_TOKEN = os.getenv("BBL_ACCESS_TOKEN")
DEVICE_ID = os.getenv("BBL_DEVICE_ID")

BROKER = "us.mqtt.bambulab.com"
PORT = 8883

REPORT_TOPIC = f"device/{DEVICE_ID}/report"
REQUEST_TOPIC = f"device/{DEVICE_ID}/request"

latest_print: dict[str, Any] = {}
sequence_id = 1


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


def summarize_state(print_obj: dict[str, Any]) -> str:
    state = print_obj.get("gcode_state", "UNKNOWN")
    percent = print_obj.get("mc_percent", 0)
    remaining = print_obj.get("mc_remaining_time")
    stage = print_obj.get("mc_print_stage")
    sub_stage = print_obj.get("mc_print_sub_stage")
    task = print_obj.get("subtask_name") or print_obj.get("task_name") or "-"

    human = {
        "RUNNING": "printing",
        "PAUSE": "paused",
        "FINISH": "finished",
        "FAILED": "failed",
        "IDLE": "idle",
        "PREPARE": "preparing",
    }.get(state, state.lower())

    return (
        f"state={human} | percent={percent}% | "
        f"remaining_min={remaining} | stage={stage}/{sub_stage} | task={task}"
    )


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

    # print_error 코드로 run-out 확인
    print_error = print_obj.get("print_error")
    if print_error is not None:
        print_error_str = str(print_error).lower().strip()
        runout_error_codes = [
            "07008102",  # filament runout
            "07008033",  # AMS filament runout
            "07008101",  # filament runout related
            "128102",  # decimal variant
        ]
        for code in runout_error_codes:
            if code in print_error_str:
                return True

    # hw_switch_action 필드로 확인
    hw_switch = str(print_obj.get("hw_switch_action", "") or "").lower()
    if "filament_runout" in hw_switch or "runout" in hw_switch:
        return True

    # ams_status 필드 확인
    ams_status = print_obj.get("ams_status")
    if ams_status is not None:
        try:
            ams_status_int = int(ams_status)
            if ams_status_int in (0x08000102, 0x08008102, 0x07008102):
                return True
        except (ValueError, TypeError):
            pass

    # filament_runout 필드 직접 확인
    if print_obj.get("filament_runout"):
        return True

    return False


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

    if "print" in payload and isinstance(payload["print"], dict):
        deep_merge(latest_print, payload["print"])
        print("[PRINT]", summarize_state(latest_print))

        # 필라멘트/AMS 정보 출력
        filament_info = extract_filament_info(latest_print)
        if filament_info.get("has_ams") or filament_info.get("active_filament_type"):
            summary = format_filament_summary(filament_info)
            if summary:
                print("[FILAMENT]", summary)

        # Run-out 에러 감지
        if check_runout_error(latest_print):
            filament_type = latest_print.get("filament_type", "")
            type_info = f" ({filament_type})" if filament_type else ""
            print(f"[RUNOUT] ⚠ 필라멘트가 소진되었습니다{type_info}. 교체해주세요.")


def on_disconnect(
    client: mqtt.Client, userdata, disconnect_flags, reason_code, properties=None
):
    print(f"[MQTT] disconnected: {reason_code}")


def main():
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
        while True:
            time.sleep(60)
            # 드물게 상태 재동기화
            request_pushall(client)
    except KeyboardInterrupt:
        print("stopping...")
    finally:
        client.loop_stop()
        client.disconnect()


if __name__ == "__main__":
    main()
