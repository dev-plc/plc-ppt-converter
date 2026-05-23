# NAS 배포 스크립트 (PowerShell)
# 실행: .\deploy_nas.ps1
#
# 사전 조건:
#   1. SSH 키가 %USERPROFILE%\.ssh\id_ed25519 에 있어야 합니다
#   2. python canva_api.py auth 로 인증이 완료되어야 합니다
#   3. 이 스크립트는 프로젝트 폴더에서 실행하세요

$NAS_USER = "plc_admin"
$NAS_HOST = "211.104.170.129"
$NAS_DIR  = "/volume1/docker/plc-ppt-bot"
$SSH_KEY  = "$env:USERPROFILE\.ssh\id_ed25519"

function nas-ssh($cmd) {
    & ssh -i $SSH_KEY "${NAS_USER}@${NAS_HOST}" $cmd
}
function nas-scp($src, $dst) {
    & scp -O -i $SSH_KEY $src "${NAS_USER}@${NAS_HOST}:$dst"
}

# ── 환경변수 확인 ──────────────────────────────────────────────────────────────
if (-not $env:TELEGRAM_BOT_TOKEN) {
    Write-Error "TELEGRAM_BOT_TOKEN 환경변수가 없습니다."
    exit 1
}
if (-not $env:CANVA_CLIENT_SECRET) {
    Write-Error "CANVA_CLIENT_SECRET 환경변수가 없습니다."
    exit 1
}

# ── Canva 토큰 파일 확인 ───────────────────────────────────────────────────────
if (-not (Test-Path ".canva_token.json")) {
    Write-Error ".canva_token.json 없음. 먼저 'python canva_api.py auth' 실행하세요."
    exit 1
}

Write-Host "[1/5] NAS 디렉토리 생성..." -ForegroundColor Cyan
nas-ssh "mkdir -p ${NAS_DIR}/docker ${NAS_DIR}/data && chown -R ${NAS_USER}:users ${NAS_DIR}"

Write-Host "[2/5] 소스 파일 전송..." -ForegroundColor Cyan
nas-scp @("convert_pptx.py","canva_api.py","telegram_bot.py","requirements.txt") "${NAS_DIR}/"
nas-scp @("docker/Dockerfile","docker/docker-compose.yml") "${NAS_DIR}/docker/"

# 표지.png → cover.png (한글 파일명 인코딩 문제 우회)
Write-Host "       표지.png → cover.png 전송..." -ForegroundColor Cyan
$TMP_COVER = [System.IO.Path]::GetTempPath() + "cover.png"
Copy-Item "표지.png" $TMP_COVER -Force
nas-scp $TMP_COVER "${NAS_DIR}/cover.png"
Remove-Item $TMP_COVER

Write-Host "[3/5] Canva 토큰 전송..." -ForegroundColor Cyan
nas-scp ".canva_token.json" "${NAS_DIR}/data/.canva_token.json"

Write-Host "[4/5] .env 생성..." -ForegroundColor Cyan
$TMP_ENV = [System.IO.Path]::GetTempFileName()
@"
TELEGRAM_BOT_TOKEN=$env:TELEGRAM_BOT_TOKEN
CANVA_CLIENT_SECRET=$env:CANVA_CLIENT_SECRET
DATA_DIR=/app/data
"@ | Set-Content -Path $TMP_ENV -Encoding UTF8
nas-scp $TMP_ENV "${NAS_DIR}/.env"
Remove-Item $TMP_ENV

Write-Host "[5/5] Docker 빌드 및 실행..." -ForegroundColor Cyan
nas-ssh "cd ${NAS_DIR}/docker && sudo /usr/local/bin/docker compose up -d --build"

Write-Host ""
Write-Host "배포 완료!" -ForegroundColor Green
Write-Host "로그 확인: ssh -i `"$SSH_KEY`" ${NAS_USER}@${NAS_HOST} `"sudo /usr/local/bin/docker logs -f plc-ppt-bot`""
