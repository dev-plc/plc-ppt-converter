"""Canva 업로드 모듈 — PL Church PPT Converter 연동

PPTX 파일을 Canva Connect API로 가져오기합니다.

설정:
    1. https://www.canva.com/developers/apps 에서 앱 생성
    2. OAuth 2.0 으로 access_token 발급 후 환경변수 설정:
       set CANVA_ACCESS_TOKEN=your_token_here   (Windows)
       export CANVA_ACCESS_TOKEN=your_token_here  (Mac/Linux)

단독 실행:
    python canva_upload.py output.pptx "에베소서 제4강"
    python canva_upload.py output.pptx "에베소서 제4강" --token YOUR_TOKEN
"""

import sys
import os
import argparse
import base64
import time

try:
    import requests
except ImportError:
    print("오류: requests 패키지가 필요합니다 → pip install requests")
    sys.exit(1)

CANVA_API_BASE = "https://api.canva.com/rest/v1"
POLL_INTERVAL  = 2    # 초
POLL_TIMEOUT   = 60   # 초 (최대 대기)


def upload_to_canva(pptx_path: str, design_name: str, access_token: str) -> dict:
    """PPTX 파일을 Canva로 가져오기.

    Returns:
        {"design_id": str, "edit_url": str}
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    name_b64 = base64.b64encode(design_name.encode("utf-8")).decode("ascii")

    with open(pptx_path, "rb") as f:
        file_bytes = f.read()

    resp = requests.post(
        f"{CANVA_API_BASE}/imports",
        headers=headers,
        files={
            "file": (
                os.path.basename(pptx_path),
                file_bytes,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        },
        data={"name_base64": name_b64},
        timeout=60,
    )

    if not resp.ok:
        raise RuntimeError(f"업로드 요청 실패 ({resp.status_code}): {resp.text}")

    job_id = resp.json().get("job", {}).get("id")
    if not job_id:
        raise RuntimeError(f"잡 ID를 받지 못했습니다: {resp.text}")

    print(f"  Canva 처리 중 (job: {job_id})")

    deadline = time.time() + POLL_TIMEOUT
    attempt  = 0
    while time.time() < deadline:
        time.sleep(POLL_INTERVAL)
        attempt += 1

        poll = requests.get(
            f"{CANVA_API_BASE}/imports/{job_id}",
            headers=headers,
            timeout=30,
        )
        poll.raise_for_status()

        job = poll.json().get("job", {})
        status = job.get("status", "")

        if status == "success":
            designs = job.get("result", {}).get("designs", [])
            if not designs:
                raise RuntimeError("가져오기 완료됐지만 디자인 ID가 없습니다.")
            design_id = designs[0]["id"]
            edit_url  = f"https://www.canva.com/design/{design_id}/edit"
            return {"design_id": design_id, "edit_url": edit_url}

        if status == "failed":
            err = job.get("error", {})
            raise RuntimeError(f"Canva 가져오기 실패: {err}")

        if attempt % 5 == 0:
            print(f"  대기 중... ({int(time.time() % deadline)}초)")

    raise TimeoutError(f"Canva 가져오기 시간 초과 ({POLL_TIMEOUT}초)")


def main():
    parser = argparse.ArgumentParser(description="PPTX → Canva 업로더")
    parser.add_argument("pptx",         help="업로드할 PPTX 경로")
    parser.add_argument("name", nargs="?", help="Canva 디자인 이름 (생략 시 파일명 사용)")
    parser.add_argument("--token",      help="Canva Access Token (없으면 CANVA_ACCESS_TOKEN 환경변수)")
    args = parser.parse_args()

    token = args.token or os.environ.get("CANVA_ACCESS_TOKEN", "").strip()
    if not token:
        print("오류: Canva Access Token이 필요합니다.")
        print()
        print("발급 방법:")
        print("  1. https://www.canva.com/developers/apps → 새 앱 생성")
        print("  2. Scopes: design:content:write, design:content:read")
        print("  3. OAuth 2.0 Authorization Code Flow 로 token 발급")
        print("  4. set CANVA_ACCESS_TOKEN=<token>  또는  --token 옵션 사용")
        sys.exit(1)

    if not os.path.exists(args.pptx):
        print(f"오류: 파일 없음 → {args.pptx}")
        sys.exit(1)

    design_name = args.name or os.path.splitext(os.path.basename(args.pptx))[0]

    print(f"[1/2] 업로드: {os.path.basename(args.pptx)}  →  Canva")
    result = upload_to_canva(args.pptx, design_name, token)

    print(f"[2/2] 완료!")
    print(f"  디자인 ID : {result['design_id']}")
    print(f"  편집 링크 : {result['edit_url']}")


if __name__ == "__main__":
    main()
