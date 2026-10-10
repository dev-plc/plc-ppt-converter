# NAS 코드 업데이트 스크립트
# GitHub master에 push 후 이 스크립트만 실행하면 끝 (PowerShell: .\update_nas.ps1)
#
# 컨테이너 재시작 시 entrypoint가 자동으로 git pull 실행
# $env:TELEGRAM_BOT_TOKEN 을 설정하고 실행하면 봇 토큰도 함께 교체

$NAS_USER = "plc_admin"
# NAS 주소: 시놀로지 DDNS (공인 IP가 바뀌어도 따라감). 다른 주소를 쓰려면 $env:NAS_HOST 지정
$NAS_HOST = if ($env:NAS_HOST) { $env:NAS_HOST } else { "plcnas01.synology.me" }
$NAS_DIR  = "/volume1/docker/plc-ppt-bot"
$SSH_KEY  = "$env:USERPROFILE\.ssh\id_ed25519"

function Fail($msg) { Write-Host "실패: $msg" -ForegroundColor Red; exit 1 }
function nas-ssh($cmd) {
    & ssh -i $SSH_KEY -o ConnectTimeout=15 "${NAS_USER}@${NAS_HOST}" $cmd
    if ($LASTEXITCODE -ne 0) { Fail "NAS 명령 실패 (종료 코드 $LASTEXITCODE): $cmd" }
}
function nas-scp($src, $dst) {
    & scp -O -i $SSH_KEY -o ConnectTimeout=15 $src "${NAS_USER}@${NAS_HOST}:$dst"
    if ($LASTEXITCODE -ne 0) { Fail "파일 전송 실패: $src" }
}

# ── 접속 확인 (실패하면 여기서 멈춤) ───────────────────────────────────────────
if (-not (Test-Path $SSH_KEY)) {
    Fail "SSH 키가 없습니다: $SSH_KEY`n      이 PC/계정에서 처음이면 NAS_SETUP.md 3장대로 키를 만들고 NAS에 등록하세요."
}
& ssh -i $SSH_KEY -o ConnectTimeout=15 -o BatchMode=yes "${NAS_USER}@${NAS_HOST}" "echo ok" | Out-Null
if ($LASTEXITCODE -ne 0) {
    Fail "NAS(${NAS_HOST}:22)에 SSH로 접속할 수 없습니다. 네트워크·방화벽·키 등록을 확인하세요."
}

if ($env:TELEGRAM_BOT_TOKEN) {
    # 봇 토큰 교체: .env를 새로 올리고 컨테이너를 다시 만든다 (재빌드 없음)
    Write-Host "새 봇 토큰 반영 중..." -ForegroundColor Cyan
    $TMP_ENV = [System.IO.Path]::GetTempFileName()
    @"
TELEGRAM_BOT_TOKEN=$env:TELEGRAM_BOT_TOKEN
DATA_DIR=/app/data
"@ | Set-Content -Path $TMP_ENV -Encoding ASCII
    nas-scp $TMP_ENV "${NAS_DIR}/.env"
    Remove-Item $TMP_ENV
    nas-ssh "cd ${NAS_DIR} && sudo /usr/local/bin/docker compose up -d --force-recreate"
} else {
    Write-Host "NAS 컨테이너 재시작 중..." -ForegroundColor Cyan
    nas-ssh "sudo /usr/local/bin/docker restart plc-ppt-bot"
}
Write-Host "완료! 컨테이너가 GitHub master에서 최신 코드를 자동으로 받아옵니다." -ForegroundColor Green
