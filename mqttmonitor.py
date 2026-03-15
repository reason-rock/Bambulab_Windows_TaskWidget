import json
import os
import ssl
import time
from typing import Any

import paho.mqtt.client as mqtt

from dotenv import load_dotenv
import os

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


def on_disconnect(client: mqtt.Client, userdata, disconnect_flags, reason_code, properties=None):
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