# NAS 코드 업데이트 스크립트
# GitHub master에 push 후 이 스크립트만 실행하면 끝
#
# 컨테이너 재시작 시 entrypoint가 자동으로 git pull 실행

$NAS_USER = "plc_admin"
$NAS_HOST = "211.104.170.129"
$SSH_KEY  = "$env:USERPROFILE\.ssh\id_ed25519"

Write-Host "NAS 컨테이너 재시작 중..." -ForegroundColor Cyan
& ssh -i $SSH_KEY "${NAS_USER}@${NAS_HOST}" "sudo /usr/local/bin/docker restart plc-ppt-bot"
Write-Host "완료! 컨테이너가 GitHub master에서 최신 코드를 자동으로 받아옵니다." -ForegroundColor Green
