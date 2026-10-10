"""PLC PPT Telegram Bot

Before PPTX 파일을 받아 변환하고,
원본 슬라이드 ↔ 변환 슬라이드 이미지를 짝지어 보낸 뒤 변환된 PPTX를 돌려줍니다.
원본 1장이 여러 장으로 나뉘면 변환 이미지를 앨범으로 묶고 "1 → 2페이지"를 붙입니다.

실행 전 환경변수 설정 (PowerShell):
    $env:TELEGRAM_BOT_TOKEN  = "봇토큰"

실행:
    python telegram_bot.py
"""

import os
import sys
import logging
import asyncio
import tempfile
from contextlib import ExitStack
from pathlib import Path

from telegram import Update, InputMediaPhoto
from telegram.error import RetryAfter
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    ContextTypes, filters,
)

# ── 로컬 모듈 ─────────────────────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from convert_pptx import parse_before, build_after
import preview

# ── 설정 ──────────────────────────────────────────────────────────────────────
BOT_TOKEN   = os.environ.get("TELEGRAM_BOT_TOKEN", "")
# 로컬: 표지.png / Docker: cover.png (한글 파일명 인코딩 우회)
_BASE = Path(__file__).parent
COVER_IMAGE = next(
    (str(_BASE / n) for n in ("cover.png", "표지.png") if (_BASE / n).exists()),
    None,
)
SEND_INTERVAL = 0.4   # 메시지 사이 간격(초) — 텔레그램 전송 제한 대비

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
)
# httpx는 요청 주소(봇 토큰 포함)를 INFO로 남기므로 경고 이상만 기록
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger(__name__)


# ── 전송 도우미 ───────────────────────────────────────────────────────────────

async def _send(make_call):
    """전송 제한(RetryAfter)에 걸리면 기다렸다가 다시 보낸다."""
    for _ in range(5):
        try:
            result = await make_call()
            await asyncio.sleep(SEND_INTERVAL)
            return result
        except RetryAfter as e:
            delay = e.retry_after
            delay = delay.total_seconds() if hasattr(delay, "total_seconds") else float(delay)
            await asyncio.sleep(delay + 1)
    return None

async def _send_photo(message, path, caption=None):
    with open(path, "rb") as f:
        data = f.read()
    return await _send(lambda: message.reply_photo(
        photo=data, caption=caption, read_timeout=60, write_timeout=60))

async def _send_album(message, paths, caption=None):
    """사진 여러 장을 앨범(최대 10장)으로. 첫 장에 캡션."""
    for start in range(0, len(paths), 10):
        part = paths[start:start + 10]
        if len(part) == 1:
            await _send_photo(message, part[0], caption if start == 0 else None)
            continue
        with ExitStack() as stack:
            media = [
                InputMediaPhoto(stack.enter_context(open(p, "rb")).read(),
                                caption=(caption if start == 0 and k == 0 else None))
                for k, p in enumerate(part)
            ]
            await _send(lambda: message.reply_media_group(
                media=media, read_timeout=90, write_timeout=90))

async def send_comparisons(message, pairs):
    """원본 이미지 → 변환 이미지(들) 순서로 전송."""
    for orig, convs in pairs:
        if orig:
            await _send_photo(message, orig)
        if not convs:
            await _send(lambda: message.reply_text("(빈 슬라이드 — 변환에서 제외)"))
        elif len(convs) == 1:
            await _send_photo(message, convs[0])
        else:
            await _send_album(message, convs, caption=preview.split_caption(len(convs)))


# ── 핸들러 ────────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "안녕하세요! PLC PPT 변환 봇입니다.\n\n"
        "📎 Before PPTX 파일을 전송하면\n"
        "   ① 원본 ↔ 변환 슬라이드 비교 이미지\n"
        "   ② 변환된 After PPTX\n"
        "를 돌려드립니다.\n"
        "원본 1장이 2장으로 나뉘면 \"1 → 2페이지\"로 표시됩니다."
    )


async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    doc = msg.document

    # PPTX 파일인지 확인
    if not (doc.file_name or "").lower().endswith(".pptx"):
        await msg.reply_text("⚠️ PPTX 파일만 처리할 수 있습니다.")
        return

    await msg.reply_text(f"📥 받았습니다: {doc.file_name}\n변환 중...")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path    = Path(tmp)
        input_path  = tmp_path / doc.file_name
        output_path = tmp_path / (Path(doc.file_name).stem + "_converted.pptx")

        # 파일 다운로드
        tg_file = await ctx.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(str(input_path))

        # ── 변환 (blocking → 별도 스레드에서 실행) ────────────────────────────
        try:
            data = await asyncio.to_thread(parse_before, str(input_path))
        except Exception as e:
            await msg.reply_text(f"❌ 파싱 오류:\n{e}")
            return

        try:
            n, slide_map = await asyncio.to_thread(
                build_after, data, str(output_path), COVER_IMAGE
            )
        except Exception as e:
            log.exception("변환 오류")
            await msg.reply_text(f"❌ 변환 오류:\n{e}")
            return

        t = data["title"]
        splits = sum(1 for outs in slide_map if len(outs) > 1)
        summary = (
            f"✅ 변환 완료: 원본 {len(slide_map)}장 → 변환 {n}장\n"
            f"📖 {t['series']} — {t['lecture_title']}"
        )
        if splits:
            summary += f"\n✂️ 나뉜 슬라이드 {splits}장"

        try:
            # ── 원본 ↔ 변환 비교 이미지 ──────────────────────────────────────
            if preview.available():
                await msg.reply_text(summary + "\n🖼 원본 ↔ 변환 비교 이미지를 보냅니다...")
                orig_imgs, conv_imgs = await asyncio.gather(
                    asyncio.to_thread(preview.render_slides, str(input_path),
                                      str(tmp_path / "pv_orig")),
                    asyncio.to_thread(preview.render_slides, str(output_path),
                                      str(tmp_path / "pv_conv")),
                )
                if orig_imgs and conv_imgs:
                    pairs = preview.pair_slides(orig_imgs, conv_imgs, slide_map)
                    await send_comparisons(msg, pairs)
                else:
                    await msg.reply_text("⚠️ 비교 이미지 생성에 실패했습니다. PPTX만 보냅니다.")
            else:
                await msg.reply_text(summary)
        except Exception:
            log.exception("비교 이미지 전송 오류")
            await msg.reply_text("⚠️ 비교 이미지 전송 중 오류가 있었습니다. PPTX를 보냅니다.")
        finally:
            # ── PPTX 전송 (비교 이미지가 실패해도 항상) ────────────────────────
            with open(output_path, "rb") as f:
                await _send(lambda: msg.reply_document(
                    document=f,
                    filename=output_path.name,
                    caption=f"변환된 PPTX입니다. (원본 {len(slide_map)}장 → {n}장)",
                    read_timeout=120, write_timeout=120,
                ))


async def handle_other(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📎 PPTX 파일을 첨부해서 보내주세요.")


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    if not BOT_TOKEN:
        print("오류: TELEGRAM_BOT_TOKEN 환경변수가 설정되지 않았습니다.")
        print("  PowerShell: $env:TELEGRAM_BOT_TOKEN='봇토큰'")
        sys.exit(1)

    if not preview.available():
        print("경고: LibreOffice/pdftoppm 없음 → 비교 이미지 없이 PPTX만 전송")

    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.ALL, handle_other))

    print("봇 시작됨. Ctrl+C로 종료.")
    app.run_polling()


if __name__ == "__main__":
    main()
