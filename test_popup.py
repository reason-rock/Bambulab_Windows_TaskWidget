import time
from datetime import datetime

from winotify import Notification

APP_ID = "Bambu Monitor"


def send_toast(title: str, message: str, important: bool = False) -> None:
    duration = "long" if important else "short"
    toast = Notification(
        app_id=APP_ID,
        title=title,
        msg=message,
        duration=duration,
    )
    toast.show()


def main() -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    send_toast("팝업 테스트(시작)", f"테스트 시작 알림입니다. {now}", important=True)
    time.sleep(2)
    send_toast("팝업 테스트(완료)", f"테스트 완료 알림입니다. {now}", important=True)
    print("[OK] 알림 2건 전송 완료")
    print("[TIP] 알림이 안 보이면 Windows 알림 설정/집중 모드를 확인하세요.")


if __name__ == "__main__":
    main()
