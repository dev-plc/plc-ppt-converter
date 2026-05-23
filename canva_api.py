"""Canva Connect API 클라이언트  (OAuth 2.0 + PKCE)

사용법:
    python canva_api.py auth              # 브라우저로 Canva 로그인
    python canva_api.py me                # 로그인된 사용자 정보 확인
    python canva_api.py upload <pptx>     # PPTX를 Canva에 임포트
    python canva_api.py logout            # 저장된 토큰 삭제

convert_pptx.py에서 직접 사용:
    python convert_pptx.py input.pptx --canva
"""

import os
import sys
import json
import time
import base64
import hashlib
import secrets
import argparse
import threading
import webbrowser
import urllib.parse
import urllib.request
import urllib.error
import http.server
from pathlib import Path

# ── 설정 ──────────────────────────────────────────────────────────────────────
CLIENT_ID     = "OC-AZ5O0t4gEaqt"
CLIENT_SECRET = os.environ.get("CANVA_CLIENT_SECRET", "")
REDIRECT_URI = "http://127.0.0.1:8080/callback"
SCOPES       = "design:content:write design:content:read asset:write profile:read"

AUTH_URL     = "https://www.canva.com/api/oauth/authorize"
TOKEN_URL    = "https://api.canva.com/rest/v1/oauth/token"
API_BASE     = "https://api.canva.com/rest/v1"

TOKEN_FILE   = Path(os.environ.get("DATA_DIR", str(Path(__file__).parent))) / ".canva_token.json"


# ── PKCE ──────────────────────────────────────────────────────────────────────

def _pkce_pair():
    verifier  = base64.urlsafe_b64encode(secrets.token_bytes(32)).rstrip(b"=").decode()
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode()).digest()
    ).rstrip(b"=").decode()
    return verifier, challenge


# ── 토큰 관리 ─────────────────────────────────────────────────────────────────

def _save_token(token: dict):
    token["saved_at"] = time.time()
    TOKEN_FILE.write_text(json.dumps(token, indent=2), encoding="utf-8")
    os.chmod(TOKEN_FILE, 0o600)

def _load_token() -> dict | None:
    if not TOKEN_FILE.exists():
        return None
    try:
        return json.loads(TOKEN_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None

def _is_expired(token: dict) -> bool:
    saved_at   = token.get("saved_at", 0)
    expires_in = token.get("expires_in", 3600)
    return time.time() > saved_at + expires_in - 60  # 1분 여유

def _refresh_token(token: dict) -> dict:
    refresh_tok = token.get("refresh_token")
    if not refresh_tok:
        raise RuntimeError("refresh_token 없음. 재인증 필요:\n  python canva_api.py auth")

    data = urllib.parse.urlencode({
        "grant_type":    "refresh_token",
        "refresh_token": refresh_tok,
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
    }).encode()

    req = urllib.request.Request(
        TOKEN_URL, data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urllib.request.urlopen(req) as r:
            new_token = json.loads(r.read())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"토큰 갱신 실패 ({e.code}). 재인증 필요:\n  python canva_api.py auth") from e

    new_token.setdefault("refresh_token", refresh_tok)
    _save_token(new_token)
    return new_token

def get_access_token() -> str:
    token = _load_token()
    if token is None:
        raise RuntimeError("인증 필요:\n  python canva_api.py auth")
    if _is_expired(token):
        token = _refresh_token(token)
    return token["access_token"]


# ── OAuth 인증 흐름 ────────────────────────────────────────────────────────────

def cmd_auth():
    if not CLIENT_SECRET:
        print("오류: CANVA_CLIENT_SECRET 환경변수가 설정되지 않았습니다.")
        print("  PowerShell: $env:CANVA_CLIENT_SECRET='시크릿값'")
        sys.exit(1)
    verifier, challenge = _pkce_pair()
    state = secrets.token_urlsafe(16)

    auth_params = {
        "response_type":         "code",
        "client_id":             CLIENT_ID,
        "redirect_uri":          REDIRECT_URI,
        "scope":                 SCOPES,
        "state":                 state,
        "code_challenge":        challenge,
        "code_challenge_method": "S256",
    }
    auth_url = AUTH_URL + "?" + urllib.parse.urlencode(auth_params)

    code_holder: dict = {}

    class _Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urllib.parse.urlparse(self.path)
            params = dict(urllib.parse.parse_qsl(parsed.query))

            if params.get("state") != state:
                self.send_error(400, "state mismatch")
                return

            code_holder["code"]  = params.get("code")
            code_holder["error"] = params.get("error")

            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            if code_holder.get("code"):
                self.wfile.write(
                    "<html><body><h2>✅ Canva 인증 완료!</h2>"
                    "<p>이 탭을 닫으세요.</p></body></html>".encode("utf-8")
                )
            else:
                self.wfile.write(
                    f"<html><body><h2>❌ 인증 실패</h2>"
                    f"<p>{params.get('error_description', '')}</p></body></html>".encode("utf-8")
                )

        def log_message(self, *_):
            pass  # 콘솔 로그 억제

    # 사용 가능한 포트 자동 탐색 (8080 → 8081 → ... → 8090)
    server = None
    port   = 8080
    for p in range(8080, 8091):
        try:
            server = http.server.HTTPServer(("127.0.0.1", p), _Handler)
            port   = p
            break
        except OSError:
            continue
    if server is None:
        print("오류: 포트 8080~8090 모두 사용 중입니다.")
        sys.exit(1)

    # 사용 포트가 바뀌면 redirect URI도 갱신
    actual_redirect = f"http://127.0.0.1:{port}/callback"
    auth_params["redirect_uri"] = actual_redirect
    auth_url = AUTH_URL + "?" + urllib.parse.urlencode(auth_params)

    def _serve():
        # 코드를 받을 때까지 계속 요청 처리 (favicon 등 복수 요청 대응)
        while not code_holder.get("code") and not code_holder.get("error"):
            server.handle_request()

    thread = threading.Thread(target=_serve, daemon=True)
    thread.start()

    print("브라우저에서 Canva 로그인 페이지가 열립니다...")
    print(f"자동으로 열리지 않으면 아래 URL을 직접 복사하세요:\n{auth_url}\n")
    webbrowser.open(auth_url)
    thread.join(timeout=120)
    server.server_close()

    code = code_holder.get("code")
    if not code:
        err = code_holder.get("error", "알 수 없는 오류")
        print(f"인증 실패: {err}")
        sys.exit(1)

    # Authorization code → Access token 교환
    token_data = urllib.parse.urlencode({
        "grant_type":    "authorization_code",
        "code":          code,
        "redirect_uri":  actual_redirect,
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code_verifier": verifier,
    }).encode()

    req = urllib.request.Request(
        TOKEN_URL, data=token_data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urllib.request.urlopen(req) as r:
            token = json.loads(r.read())
    except urllib.error.HTTPError as e:
        print(f"토큰 발급 실패 ({e.code}): {e.read().decode()}")
        sys.exit(1)

    _save_token(token)
    print("인증 완료! 토큰이 저장되었습니다.")


# ── 공통 API 요청 ─────────────────────────────────────────────────────────────

def _api_request(method: str, url: str, *, json_body=None,
                 raw_body: bytes = None, extra_headers: dict = None) -> dict:
    access_token = get_access_token()
    hdrs = {"Authorization": f"Bearer {access_token}"}

    body = None
    if json_body is not None:
        body = json.dumps(json_body).encode()
        hdrs["Content-Type"] = "application/json"
    elif raw_body is not None:
        body = raw_body

    if extra_headers:
        hdrs.update(extra_headers)

    req = urllib.request.Request(url, data=body, headers=hdrs, method=method)
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"API 오류 {e.code}: {e.read().decode()}") from e


# ── PPTX 업로드 (Import) ──────────────────────────────────────────────────────

def _build_multipart(boundary: str, filename: str,
                     file_bytes: bytes, mime: str) -> bytes:
    b  = boundary.encode()
    fn = filename.encode()
    return b"".join([
        b"--" + b + b"\r\n",
        b'Content-Disposition: form-data; name="asset_upload"; filename="' + fn + b'"\r\n',
        b"Content-Type: " + mime.encode() + b"\r\n\r\n",
        file_bytes,
        b"\r\n--" + b + b"--\r\n",
    ])


def upload_pptx(pptx_path: str, verbose: bool = True) -> str:
    """PPTX를 Canva에 임포트하고 디자인 URL을 반환한다."""
    path = Path(pptx_path)
    if not path.exists():
        raise FileNotFoundError(f"파일 없음: {pptx_path}")

    if verbose:
        print(f"[Canva] 업로드 중: {path.name}")

    access_token = get_access_token()
    boundary  = secrets.token_hex(16)
    mime_type = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    body = _build_multipart(boundary, path.name, path.read_bytes(), mime_type)

    hdrs = {
        "Authorization":   f"Bearer {access_token}",
        "Content-Type":    f"multipart/form-data; boundary={boundary}",
        "Import-Metadata": json.dumps({
            "title":     path.stem,
            "mime_type": mime_type,
        }),
    }

    req = urllib.request.Request(
        f"{API_BASE}/imports", data=body, headers=hdrs, method="POST"
    )
    try:
        with urllib.request.urlopen(req) as r:
            result = json.loads(r.read())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"업로드 실패 {e.code}: {e.read().decode()}") from e

    job_id = result.get("job", {}).get("id", "")
    if verbose:
        print(f"[Canva] 처리 중... (job: {job_id})")

    # 완료 대기 (최대 60초 폴링)
    for _ in range(30):
        time.sleep(2)
        try:
            req2 = urllib.request.Request(
                f"{API_BASE}/imports/{job_id}",
                headers={"Authorization": f"Bearer {access_token}"},
            )
            with urllib.request.urlopen(req2) as r:
                st = json.loads(r.read())
        except Exception:
            continue

        status    = st.get("job", {}).get("status", "")
        design_id = st.get("job", {}).get("design", {}).get("id", "")

        if status == "success" and design_id:
            url = f"https://www.canva.com/design/{design_id}/edit"
            if verbose:
                print(f"[Canva] 완료! 디자인 ID: {design_id}")
                print(f"[Canva] 열기: {url}")
            return url
        elif status == "failed":
            raise RuntimeError(f"Canva 임포트 실패: {st}")

        if verbose:
            print(f"[Canva] 상태: {status}...")

    raise TimeoutError("Canva 임포트 시간 초과. 나중에 Canva에서 직접 확인하세요.")


def cmd_upload(pptx_path: str):
    try:
        upload_pptx(pptx_path, verbose=True)
    except (RuntimeError, FileNotFoundError, TimeoutError) as e:
        print(f"오류: {e}")
        sys.exit(1)


# ── 사용자 정보 ───────────────────────────────────────────────────────────────

def cmd_me():
    try:
        result = _api_request("GET", f"{API_BASE}/users/me")
    except RuntimeError as e:
        print(f"오류: {e}")
        sys.exit(1)
    user = result.get("user", {})
    print(f"이름  : {user.get('display_name', '-')}")
    print(f"이메일: {user.get('email', '-')}")
    print(f"ID    : {user.get('id', '-')}")


# ── 로그아웃 ──────────────────────────────────────────────────────────────────

def cmd_logout():
    if TOKEN_FILE.exists():
        TOKEN_FILE.unlink()
        print("토큰이 삭제되었습니다.")
    else:
        print("저장된 토큰이 없습니다.")


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Canva Connect API 클라이언트")
    sub    = parser.add_subparsers(dest="cmd")

    sub.add_parser("auth",   help="OAuth 인증 (브라우저 로그인)")
    sub.add_parser("me",     help="로그인된 사용자 정보")
    sub.add_parser("logout", help="저장된 토큰 삭제")

    up = sub.add_parser("upload", help="PPTX 파일을 Canva에 업로드")
    up.add_argument("pptx", help="업로드할 PPTX 경로")

    args = parser.parse_args()

    if   args.cmd == "auth":   cmd_auth()
    elif args.cmd == "me":     cmd_me()
    elif args.cmd == "logout": cmd_logout()
    elif args.cmd == "upload": cmd_upload(args.pptx)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
