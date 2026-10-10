# NAS 배포 스크립트 — Dockerfile(설치 패키지)이 바뀌었을 때 실행
# 코드만 바뀌었을 때는 update_nas.ps1 로 충분
#
# 봇 토큰: PowerShell에 $env:TELEGRAM_BOT_TOKEN 이 있으면 NAS .env를 새로 만들고,
#          없으면 NAS에 이미 있는 .env를 그대로 사용합니다.

$NAS_USER = "plc_admin"
$NAS_HOST = "211.104.170.129"
$NAS_DIR  = "/volume1/docker/plc-ppt-bot"
$SSH_KEY  = "$env:USERPROFILE\.ssh\id_ed25519"

function nas-ssh($cmd) { & ssh -i $SSH_KEY "${NAS_USER}@${NAS_HOST}" $cmd }
function nas-scp($src, $dst) { & scp -O -i $SSH_KEY $src "${NAS_USER}@${NAS_HOST}:$dst" }

Write-Host "[1/4] NAS 디렉토리 확인..." -ForegroundColor Cyan
nas-ssh "mkdir -p ${NAS_DIR}/data && chown -R ${NAS_USER}:users ${NAS_DIR}"

Write-Host "[2/4] Docker 파일 전송..." -ForegroundColor Cyan
nas-scp "docker/Dockerfile"         "${NAS_DIR}/Dockerfile"
nas-scp "docker/docker-compose.yml" "${NAS_DIR}/docker-compose.yml"

Write-Host "[3/4] .env 확인..." -ForegroundColor Cyan
if ($env:TELEGRAM_BOT_TOKEN) {
    $TMP_ENV = [System.IO.Path]::GetTempFileName()
    # ASCII로 저장 (UTF8은 BOM이 붙어 변수 이름이 깨질 수 있음)
    @"
TELEGRAM_BOT_TOKEN=$env:TELEGRAM_BOT_TOKEN
DATA_DIR=/app/data
"@ | Set-Content -Path $TMP_ENV -Encoding ASCII
    nas-scp $TMP_ENV "${NAS_DIR}/.env"
    Remove-Item $TMP_ENV
    Write-Host "       새 .env 전송 완료"
} else {
    $exists = nas-ssh "test -f ${NAS_DIR}/.env && echo yes"
    if ($exists -ne "yes") {
        Write-Error "NAS에 .env가 없습니다. 먼저 `$env:TELEGRAM_BOT_TOKEN = '봇토큰' 을 설정하고 다시 실행하세요."
        exit 1
    }
    Write-Host "       NAS의 기존 .env 사용 (봇 토큰 그대로)"
}

Write-Host "[4/4] Docker 빌드 및 시작 (GitHub 클론 + LibreOffice 설치, 5~10분 소요)..." -ForegroundColor Cyan
nas-ssh "cd ${NAS_DIR} && sudo /usr/local/bin/docker compose down; sudo /usr/local/bin/docker compose build --no-cache && sudo /usr/local/bin/docker compose up -d"

Write-Host ""
Write-Host "배포 완료!" -ForegroundColor Green
Write-Host "로그 확인: ssh -i `"$SSH_KEY`" ${NAS_USER}@${NAS_HOST} `"sudo /usr/local/bin/docker logs -f plc-ppt-bot`""
Write-Host ""
Write-Host "이후 코드만 바뀌면 update_nas.ps1 실행하세요."
