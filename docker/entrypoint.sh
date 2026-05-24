#!/bin/sh
set -e
echo "[PLC Bot] GitHub에서 최신 코드 받는 중..."
git pull origin master 2>&1 || echo "[PLC Bot] git pull 실패, 기존 버전으로 실행"
cp "표지.png" cover.png 2>/dev/null || true
echo "[PLC Bot] 봇 시작..."
exec python telegram_bot.py
