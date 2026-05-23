"""PLC PPT Telegram Bot

PPTX 파일을 받아 변환 후 Canva에 업로드하고 결과를 돌려줍니다.

실행 전 환경변수 설정 (PowerShell):
    $env:TELEGRAM_BOT_TOKEN  = "봇토큰"
    $env:CANVA_CLIENT_SECRET = "Canva시크릿"

실행:
    python telegram_bot.py
"""

import os
import sys
import logging
import tempfile
from pathlib import Path

from telegram import Update
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    ContextTypes, filters,
)

# ── 로컬 모듈 ─────────────────────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from convert_pptx import parse_before, build_after
from canva_api import upload_pptx

# ── 설정 ──────────────────────────────────────────────────────────────────────
BOT_TOKEN   = os.environ.get("TELEGRAM_BOT_TOKEN", "")
# 로컬: 표지.png / Docker: cover.png (한글 파일명 인코딩 우회)
_BASE = Path(__file__).parent
COVER_IMAGE = next(
    (str(_BASE / n) for n in ("cover.png", "표지.png") if (_BASE / n).exists()),
    None,
)

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)


# ── 핸들러 ────────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "안녕하세요! PLC PPT 변환 봇입니다.\n\n"
        "📎 Before PPTX 파일을 전송하면\n"
        "   ① 변환된 After PPTX\n"
        "   ② Canva 편집 링크\n"
        "를 돌려드립니다."
    )


async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    # PPTX 파일인지 확인
    if not (doc.file_name or "").lower().endswith(".pptx"):
        await update.message.reply_text("⚠️ PPTX 파일만 처리할 수 있습니다.")
        return

    await update.message.reply_text(f"📥 받았습니다: {doc.file_name}\n변환 중...")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path    = Path(tmp)
        input_path  = tmp_path / doc.file_name
        output_path = tmp_path / (Path(doc.file_name).stem + "_converted.pptx")

        # 파일 다운로드
        tg_file = await ctx.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(str(input_path))

        # ── 변환 ──────────────────────────────────────────────────────────────
        try:
            data = parse_before(str(input_path))
        except Exception as e:
            await update.message.reply_text(f"❌ 파싱 오류:\n{e}")
            return

        cover = COVER_IMAGE
        try:
            n = build_after(data, str(output_path), cover_image=cover)
        except Exception as e:
            await update.message.reply_text(f"❌ 변환 오류:\n{e}")
            return

        t = data["title"]
        await update.message.reply_text(
            f"✅ 변환 완료 ({n}슬라이드)\n"
            f"📖 {t['series']} — {t['lecture_title']}"
        )

        # ── PPTX 전송 ─────────────────────────────────────────────────────────
        with open(output_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=output_path.name,
                caption="변환된 PPTX입니다.",
            )

        # ── Canva 업로드 ──────────────────────────────────────────────────────
        await update.message.reply_text("☁️ Canva에 업로드 중...")
        try:
            canva_url = upload_pptx(str(output_path), verbose=False)
            await update.message.reply_text(
                f"🎨 Canva 편집 링크:\n{canva_url}"
            )
        except RuntimeError as e:
            await update.message.reply_text(
                f"⚠️ Canva 업로드 실패:\n{e}\n\n"
                "PPTX는 위에서 받으실 수 있습니다."
            )


async def handle_other(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📎 PPTX 파일을 첨부해서 보내주세요.")


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    if not BOT_TOKEN:
        print("오류: TELEGRAM_BOT_TOKEN 환경변수가 설정되지 않았습니다.")
        print("  PowerShell: $env:TELEGRAM_BOT_TOKEN='봇토큰'")
        sys.exit(1)

    if not os.environ.get("CANVA_CLIENT_SECRET"):
        print("경고: CANVA_CLIENT_SECRET 미설정 → Canva 업로드 불가")

    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.ALL, handle_other))

    print("봇 시작됨. Ctrl+C로 종료.")
    app.run_polling()


if __name__ == "__main__":
    main()
