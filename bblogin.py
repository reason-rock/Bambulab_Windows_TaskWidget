import os
import requests
import getpass
from dotenv import load_dotenv

load_dotenv()

EMAIL = os.getenv("BBL_EMAIL")
PASSWORD = os.getenv("BBL_PASSWORD")

if not EMAIL or not PASSWORD:
    raise RuntimeError("BBL_EMAIL/BBL_PASSWORD 가 .env에 설정되어야 합니다.")

BASE = "https://api.bambulab.com"

session = requests.Session()
session.headers.update({
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0",
})

# 1) 비밀번호 로그인
resp = session.post(
    f"{BASE}/v1/user-service/user/login",
    json={
        "account": EMAIL,
        "password": PASSWORD,
    },
    timeout=30,
)
resp.raise_for_status()
data = resp.json()
print("LOGIN =", data)

access_token = data.get("accessToken")

# 2) verifyCode 필요 시 코드 로그인
if not access_token:
    login_type = data.get("loginType")
    if login_type != "verifyCode":
        raise RuntimeError(f"예상 못한 로그인 상태: {data}")

    print("이메일로 온 인증코드를 입력하세요.")
    code = getpass.getpass("Verification code: ").strip()

    resp2 = session.post(
        f"{BASE}/v1/user-service/user/login",
        json={
            "account": EMAIL,
            "code": code,
        },
        timeout=30,
    )
    resp2.raise_for_status()
    data2 = resp2.json()
    print("VERIFY LOGIN =", data2)

    access_token = data2.get("accessToken")
    if not access_token:
        raise RuntimeError(f"accessToken 발급 실패: {data2}")

print("ACCESS_TOKEN 확보 완료")
print(access_token[:20] + "...")

# 3) 인증 헤더 설정
session.headers["Authorization"] = f"Bearer {access_token}"

# 4) USER_ID(uid) 조회
pref_resp = session.get(
    f"{BASE}/v1/design-user-service/my/preference",
    timeout=30,
)
print("PREFERENCE STATUS =", pref_resp.status_code)
print("PREFERENCE BODY =", pref_resp.text)
pref_resp.raise_for_status()

pref_data = pref_resp.json()
user_id = pref_data.get("uid") or pref_data.get("data", {}).get("uid")
if not user_id:
    raise RuntimeError(f"uid를 찾지 못했습니다: {pref_data}")

print("USER_ID =", user_id)
print("MQTT_USERNAME =", f"u_{user_id}")

# 5) 바인드된 프린터 목록 조회
bind_resp = session.get(
    f"{BASE}/v1/iot-service/api/user/bind",
    timeout=30,
)
print("BIND STATUS =", bind_resp.status_code)
print("BIND BODY =", bind_resp.text)
bind_resp.raise_for_status()

bind_data = bind_resp.json()

# 응답 구조 대응
devices = []
if isinstance(bind_data, dict):
    if isinstance(bind_data.get("devices"), list):
        devices = bind_data["devices"]
    elif isinstance(bind_data.get("data"), dict) and isinstance(bind_data["data"].get("devices"), list):
        devices = bind_data["data"]["devices"]
    elif isinstance(bind_data.get("data"), list):
        devices = bind_data["data"]

if not devices:
    print("기기 목록을 찾지 못했습니다. bind 응답 구조를 직접 확인하세요.")
else:
    print("\n=== DEVICES ===")
    for i, d in enumerate(devices, start=1):
        dev_id = d.get("dev_id") or d.get("deviceId") or d.get("device_id")
        name = d.get("name") or d.get("dev_name") or d.get("devName")
        model = d.get("dev_model_name") or d.get("model") or d.get("productName")
        product = d.get("dev_product_name") or d.get("product") or d.get("type")
        print(f"[{i}] name={name} | model={model} | product={product} | dev_id={dev_id}")
