"""PL Church Sermon PPT Converter  (Before → After)

Usage:
    python convert_pptx.py 에베소서 제4강(before).pptx
    python convert_pptx.py input.pptx output.pptx --png
"""

import sys
import io
import re
import os
import argparse

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

# ── 색상 ──────────────────────────────────────────────────────────────────────
BG_COLOR   = RGBColor(0x1B, 0x20, 0x45)
HEADER_BG  = RGBColor(0x23, 0x2A, 0x56)
GOLD       = RGBColor(0xF5, 0xC5, 0x18)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

# ── 슬라이드 크기 (표준 16:9, 1920×1080 기준) ────────────────────────────────
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.50)

# ── 폰트 ─────────────────────────────────────────────────────────────────────
FONT_CLASSIC_BOLD = "국립박물관문화재단클래식 Bold"
FONT_CLASSIC_MED  = "국립박물관문화재단클래식 Medium"
FONT_NOTO         = "Noto Sans KR"

MAX_BULLETS = 3   # 슬라이드당 최대 불렛 수


# ── Before PPTX 파서 ──────────────────────────────────────────────────────────

def _para_text(para) -> str:
    return "".join(run.text for run in para.runs).strip()

def _tf_by_name(slide, name):
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == name:
            return shape.text_frame
    return None

def _split_series_title(full_text: str):
    m = re.search(r"(.*?제\s*\d+\s*강)", full_text)
    if m:
        return m.group(1).strip(), full_text[m.end():].strip()
    return "", full_text.strip()

def parse_title_slide(slide):
    title_tf = _tf_by_name(slide, "제목 1")
    sub_tf   = _tf_by_name(slide, "부제목 2")

    full_text = ""
    if title_tf:
        for para in title_tf.paragraphs:
            t = _para_text(para)
            if t:
                full_text += t

    series, lecture_title = _split_series_title(full_text)

    sub_paras = []
    if sub_tf:
        for para in sub_tf.paragraphs:
            t = _para_text(para)
            if t:
                sub_paras.append(t)

    pastor = "Rev. Dr. Paul Junghoon Lee"

    if not lecture_title and sub_paras:
        non_pastor   = [t for t in sub_paras if "Rev" not in t and "Dr." not in t]
        pastor_lines = [t for t in sub_paras if "Rev" in t or "Dr." in t]
        lecture_title = non_pastor[0] if non_pastor else ""
        if pastor_lines:
            pastor = pastor_lines[0]
    else:
        for t in sub_paras:
            if t:
                pastor = t
                break

    return series, lecture_title, pastor

def parse_content_slide(slide):
    title_tf   = _tf_by_name(slide, "제목 1")
    content_tf = _tf_by_name(slide, "내용 개체 틀 2")

    heading_parts = []
    if title_tf:
        for para in title_tf.paragraphs:
            t = _para_text(para)
            if t:
                heading_parts.append(t)
    heading = " ".join(heading_parts)

    bullets = []
    if content_tf:
        for para in content_tf.paragraphs:
            t = _para_text(para)
            if t:
                bullets.append(t)

    return heading, bullets

def parse_before(path: str) -> dict:
    prs = Presentation(path)
    all_slides = list(prs.slides)
    series, lecture_title, pastor = parse_title_slide(all_slides[0])

    content_slides = []
    for slide in all_slides[1:]:
        heading, bullets = parse_content_slide(slide)
        if heading or bullets:
            content_slides.append({"heading": heading, "bullets": bullets})

    return {
        "title": {"series": series, "lecture_title": lecture_title, "pastor": pastor},
        "content_slides": content_slides,
    }


# ── 핵심어 판별 (골드 색상 적용 대상) ─────────────────────────────────────────

# 성경 약자 목록
_BIBLE_ABBRS = (
    "창|출|레|민|신|수|삿|룻|삼|왕|대|스|느|에|욥|시|잠|전|아|사|렘|애|겔|단"
    "|호|욜|암|옵|욘|미|나|합|습|학|슥|말"
    "|마|막|눅|요|행|롬|고|갈|엡|빌|골|살|딤|딛|몬|히|약|벧|유|계"
)
_BIBLE_RE = re.compile(rf"(?:{_BIBLE_ABBRS})\s*\d+\s*[：:]\s*\d+")

def _is_key_term(text: str) -> bool:
    """짧은 단독 핵심어이면 True → 골드 색상 적용"""
    t = text.strip()
    if len(t) > 28:
        return False
    if t and t[0] in ('"', '"', '"', "'", "'", "'", "「", "『", "＂"):
        return False
    if _BIBLE_RE.search(t):
        return False
    if t.count("(") + t.count("（") >= 2:
        return False
    return True


# ── After PPTX 빌더 ───────────────────────────────────────────────────────────

def _solid_bg(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def _add_rect(slide, left, top, width, height, fill_color: RGBColor):
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        left, top, width, height,
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape

def _add_textbox(slide, left, top, width, height,
                 text, font_name, size_pt,
                 bold=None, italic=None, color=WHITE,
                 align=PP_ALIGN.LEFT, word_wrap=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf  = box.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name  = font_name
    run.font.size  = Pt(size_pt)
    if bold    is not None: run.font.bold   = bold
    if italic  is not None: run.font.italic = italic
    run.font.color.rgb = color
    return box

def build_title_slide(prs, series: str, lecture_title: str, pastor: str,
                      cover_image: str = None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    if cover_image and os.path.exists(cover_image):
        # 표지.png 를 슬라이드 전체 배경으로 (성경책·별 장식 모두 포함됨)
        slide.shapes.add_picture(cover_image, 0, 0, SLIDE_W, SLIDE_H)
    else:
        # 표지 이미지 없으면 단색 배경 + 별 텍스트 대체
        _solid_bg(slide, BG_COLOR)
        _add_textbox(
            slide,
            SLIDE_W - Inches(1.5), Inches(0.3),
            Inches(1.2), Inches(0.9),
            "✦", FONT_CLASSIC_BOLD, 48,
            bold=False, color=GOLD, align=PP_ALIGN.CENTER,
        )

    # ── 시리즈명 (예: 에베소서 제4강)
    _add_textbox(
        slide,
        Inches(1.5), Inches(1.9),
        SLIDE_W - Inches(3), Inches(0.6),
        series, FONT_CLASSIC_MED, 22,
        color=WHITE, align=PP_ALIGN.CENTER,
    )

    # ── 강의 제목 (중앙, 크게)
    _add_textbox(
        slide,
        Inches(1.5), Inches(2.55),
        SLIDE_W - Inches(3), Inches(2.2),
        lecture_title, FONT_CLASSIC_BOLD, 64,
        bold=True, color=WHITE, align=PP_ALIGN.CENTER,
    )

    # ── 작은 별 + 목사님 이름
    pastor_box = slide.shapes.add_textbox(
        Inches(1.5), Inches(4.85),
        SLIDE_W - Inches(3), Inches(0.9),
    )
    tf = pastor_box.text_frame
    tf.word_wrap = False
    p  = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER

    star_run = p.add_run()
    star_run.text = "✦  "
    star_run.font.name  = FONT_CLASSIC_BOLD
    star_run.font.size  = Pt(14)
    star_run.font.color.rgb = GOLD

    name_run = p.add_run()
    name_run.text = pastor
    name_run.font.name   = FONT_CLASSIC_MED
    name_run.font.size   = Pt(24)
    name_run.font.italic = True
    name_run.font.color.rgb = WHITE

    return slide

def build_content_slide(prs, heading: str, bullets: list, is_continued=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _solid_bg(slide, BG_COLOR)

    HEADER_H = Inches(1.20)

    # ── 헤더 배경 박스
    _add_rect(slide, Inches(0), Inches(0), SLIDE_W, HEADER_H, HEADER_BG)

    # ── 헤더 텍스트
    head_label = heading + (" (계속)" if is_continued else "")
    _add_textbox(
        slide,
        Inches(0.4), Inches(0.12),
        SLIDE_W - Inches(0.8), HEADER_H - Inches(0.1),
        head_label, FONT_CLASSIC_BOLD, 30,
        bold=True, color=WHITE, align=PP_ALIGN.LEFT,
    )

    # ── 본문 불렛 영역
    body_top = HEADER_H + Inches(0.20)
    body_h   = SLIDE_H  - body_top - Inches(0.2)
    body_box = slide.shapes.add_textbox(
        Inches(0.45), body_top,
        SLIDE_W - Inches(0.8), body_h,
    )
    tf = body_box.text_frame
    tf.word_wrap = True

    # hanging indent: 두 번째 줄이 "-" 뒤 텍스트와 정렬되도록
    # marL = 들여쓰기 총량, indent = 첫 줄은 그만큼 왼쪽으로 당김
    _MARL   =  int(Inches(0.52))   # 두 번째 줄 시작 위치 (EMU)
    _INDENT = -int(Inches(0.52))   # 첫 줄 = marL + indent = 0 위치

    for i, bullet_text in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment    = PP_ALIGN.LEFT
        p.space_before = Pt(22)
        p.space_after  = Pt(4)

        # hanging indent XML 적용
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL",   str(_MARL))
        pPr.set("indent", str(_INDENT))

        run = p.add_run()
        run.text = f"-  {bullet_text}"

        is_key = _is_key_term(bullet_text)
        run.font.name  = FONT_NOTO
        run.font.size  = Pt(28)
        run.font.bold  = True
        run.font.color.rgb = GOLD if is_key else WHITE

    return slide


def _split_bullets(bullets: list, max_per_slide: int) -> list:
    n = len(bullets)
    if n <= max_per_slide:
        return [bullets]
    num_chunks = (n + max_per_slide - 1) // max_per_slide
    base_size  = n // num_chunks
    remainder  = n % num_chunks
    chunks, idx = [], 0
    for i in range(num_chunks):
        size = base_size + (1 if i < remainder else 0)
        chunks.append(bullets[idx : idx + size])
        idx += size
    return chunks


def build_after(data: dict, output_path: str, cover_image: str = None):
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    t = data["title"]
    build_title_slide(prs, t["series"], t["lecture_title"], t["pastor"],
                      cover_image=cover_image)

    for cs in data["content_slides"]:
        heading = cs["heading"]
        bullets = cs["bullets"]

        if not bullets:
            build_content_slide(prs, heading, [])
            continue

        for j, chunk in enumerate(_split_bullets(bullets, MAX_BULLETS)):
            build_content_slide(prs, heading, chunk, is_continued=(j > 0))

    prs.save(output_path)
    return len(prs.slides)


# ── PNG 내보내기 (PowerPoint COM, Windows 전용) ────────────────────────────────

def export_png(pptx_path: str, out_dir: str, width_px=1920, height_px=1080):
    import comtypes.client
    pptx_path = os.path.abspath(pptx_path)
    out_dir   = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    app = comtypes.client.CreateObject("PowerPoint.Application")
    app.Visible = 1
    try:
        prs = app.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)
        try:
            for i, slide in enumerate(prs.Slides, start=1):
                png_path = os.path.join(out_dir, f"slide_{i:03d}.png")
                slide.Export(png_path, "PNG", width_px, height_px)
                print(f"  PNG: {os.path.basename(png_path)}")
        finally:
            prs.Close()
    finally:
        app.Quit()

    print(f"  완료 → {out_dir}")


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="PL Church PPT 변환기")
    parser.add_argument("input",           help="Before PPTX 경로")
    parser.add_argument("output", nargs="?", help="출력 경로 (생략 시 자동)")
    parser.add_argument("--png", action="store_true",
                        help="변환 후 PNG 내보내기 (PowerPoint 필요)")
    parser.add_argument("--png-dir",       help="PNG 저장 폴더")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"오류: 파일 없음 → {args.input}")
        sys.exit(1)

    out_path = args.output or (os.path.splitext(args.input)[0] + "_converted.pptx")

    print(f"[1/3] 파싱: {os.path.basename(args.input)}")
    data = parse_before(args.input)
    t = data["title"]
    print(f"      시리즈: {t['series']}")
    print(f"      제목  : {t['lecture_title']}")
    print(f"      목사  : {t['pastor']}")
    print(f"      내용 슬라이드: {len(data['content_slides'])}개 (분리 전)")

    # 표지 이미지: 스크립트와 같은 폴더의 표지.png 자동 탐색
    script_dir  = os.path.dirname(os.path.abspath(__file__))
    cover_image = os.path.join(script_dir, "표지.png")
    if os.path.exists(cover_image):
        print(f"[2/3] 변환 중... (표지 이미지 적용)")
    else:
        cover_image = None
        print(f"[2/3] 변환 중... (표지.png 없음, 기본 배경 사용)")

    n = build_after(data, out_path, cover_image=cover_image)
    print(f"[3/3] 완료 → {os.path.basename(out_path)}  ({n} 슬라이드)")

    if args.png:
        png_dir = args.png_dir or os.path.splitext(out_path)[0] + "_slides"
        print(f"[+]  PNG 내보내기 → {png_dir}")
        export_png(out_path, png_dir)


if __name__ == "__main__":
    main()
