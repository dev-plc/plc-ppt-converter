#!/bin/sh
set -e
echo "[PLC Bot] GitHub에서 최신 코드 받는 중..."
git pull origin master 2>&1 || echo "[PLC Bot] git pull 실패, 기존 버전으로 실행"
pip install -q --no-cache-dir -r requirements.txt 2>&1 || echo "[PLC Bot] 패키지 설치 확인 실패"
cp "표지.png" cover.png 2>/dev/null || true
if [ -f docker/fonts.conf ]; then
    cp docker/fonts.conf /etc/fonts/local.conf 2>/dev/null && fc-cache -f >/dev/null 2>&1 || true
fi
command -v soffice >/dev/null 2>&1 || echo "[PLC Bot] LibreOffice 없음 → 비교 이미지 없이 동작 (deploy_nas.ps1 재실행 필요)"
echo "[PLC Bot] 봇 시작..."
exec python telegram_bot.py
