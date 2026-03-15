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

# 1) 먼저 비밀번호로 시도
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
if access_token:
    print("바로 로그인 성공")
else:
    login_type = data.get("loginType")
    if login_type != "verifyCode":
        raise RuntimeError(f"예상 못한 로그인 상태: {data}")

    print("이메일로 온 인증코드를 입력하세요.")
    code = getpass.getpass("Verification code: ").strip()

    # 2) 인증코드로 다시 로그인
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
