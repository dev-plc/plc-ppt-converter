# NAS 최초 배포 스크립트 (처음 한 번만 실행)
# 이후 업데이트는 update_nas.ps1 사용
#
# 실행 전:
#   $env:TELEGRAM_BOT_TOKEN = "봇토큰"
#   $env:CANVA_CLIENT_SECRET = "Canva시크릿"

$NAS_USER = "plc_admin"
$NAS_HOST = "211.104.170.129"
$NAS_DIR  = "/volume1/docker/plc-ppt-bot"
$SSH_KEY  = "$env:USERPROFILE\.ssh\id_ed25519"

function nas-ssh($cmd) { & ssh -i $SSH_KEY "${NAS_USER}@${NAS_HOST}" $cmd }
function nas-scp($src, $dst) { & scp -O -i $SSH_KEY $src "${NAS_USER}@${NAS_HOST}:$dst" }

if (-not $env:TELEGRAM_BOT_TOKEN) { Write-Error "TELEGRAM_BOT_TOKEN 환경변수가 없습니다."; exit 1 }
if (-not $env:CANVA_CLIENT_SECRET) { Write-Error "CANVA_CLIENT_SECRET 환경변수가 없습니다."; exit 1 }

Write-Host "[1/4] NAS 디렉토리 생성..." -ForegroundColor Cyan
nas-ssh "mkdir -p ${NAS_DIR}/data && chown -R ${NAS_USER}:users ${NAS_DIR}"

Write-Host "[2/4] Docker 파일 전송..." -ForegroundColor Cyan
nas-scp "docker/Dockerfile"         "${NAS_DIR}/Dockerfile"
nas-scp "docker/docker-compose.yml" "${NAS_DIR}/docker-compose.yml"

Write-Host "[3/4] .env 생성..." -ForegroundColor Cyan
$TMP_ENV = [System.IO.Path]::GetTempFileName()
@"
TELEGRAM_BOT_TOKEN=$env:TELEGRAM_BOT_TOKEN
CANVA_CLIENT_SECRET=$env:CANVA_CLIENT_SECRET
DATA_DIR=/app/data
"@ | Set-Content -Path $TMP_ENV -Encoding UTF8
nas-scp $TMP_ENV "${NAS_DIR}/.env"
Remove-Item $TMP_ENV

Write-Host "[4/4] Docker 빌드 및 시작 (GitHub 클론 포함, 2~3분 소요)..." -ForegroundColor Cyan
nas-ssh "cd ${NAS_DIR} && sudo /usr/local/bin/docker compose down; sudo /usr/local/bin/docker compose build --no-cache && sudo /usr/local/bin/docker compose up -d"

Write-Host ""
Write-Host "배포 완료!" -ForegroundColor Green
Write-Host "로그 확인: ssh -i `"$SSH_KEY`" ${NAS_USER}@${NAS_HOST} `"sudo /usr/local/bin/docker logs -f plc-ppt-bot`""
Write-Host ""
Write-Host "이후 코드 업데이트는 update_nas.ps1 실행하세요."
