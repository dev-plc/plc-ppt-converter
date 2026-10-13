"""PL Church Sermon PPT Converter  (Before → After)

Usage:
    python convert_pptx.py 에베소서 제4강(before).pptx
    python convert_pptx.py input.pptx output.pptx --png

디자인 기준: 완성샘플(마태복음 62강·신명기 2강) 1920×1080 PNG에서 측정한 값.
"""

import sys
import io
import re
import os
import argparse

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

# ── 색상 (완성샘플 측정값) ────────────────────────────────────────────────────
BG_TOP     = RGBColor(0x09, 0x14, 0x42)   # 배경 그라데이션 위
BG_BOTTOM  = RGBColor(0x19, 0x16, 0x29)   # 배경 그라데이션 아래
BG_COLOR   = RGBColor(0x12, 0x15, 0x36)   # 그라데이션 중간값 (단색 대체용)
HEADER_BG  = RGBColor(0x3C, 0x40, 0x59)   # 제목 글자에 맞춘 박스
GOLD       = RGBColor(0xFF, 0xCD, 0x4A)   # 원본 형광펜 강조
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)   # 제목
BODY_COLOR = RGBColor(0xEE, 0xEE, 0xEE)   # 본문
PASTOR_COLOR = RGBColor(0xFE, 0xEF, 0xC5) # 표지 목사님 이름

# ── 슬라이드 크기 (표준 16:9, 1920×1080 기준) ────────────────────────────────
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ── 폰트 ─────────────────────────────────────────────────────────────────────
FONT_CLASSIC_MED  = "국립박물관문화재단클래식 Medium"   # 제목·표지
FONT_NOTO         = "Noto Sans KR"

# ── 본문/제목 레이아웃 (인치·pt) ──────────────────────────────────────────────
BODY_PT_MAX   = 32      # 완성샘플 본문 크기
BODY_PT_MIN   = 28      # 이 크기로도 안 들어갈 때만 슬라이드 분리
LINE_RATIO    = 1.28    # 줄 간격 = 글자 크기 × 1.28 (샘플: 32pt → 41pt)
MIN_GAP_RATIO = 0.25    # 불렛 사이 최소 간격 (줄 간격 대비)
MAX_GAP_RATIO = 2.2     # 불렛 사이 최대 간격

HEAD_PT       = 38      # 제목 크기
HEAD_LINE_PT  = 40      # 제목 줄 간격
HEAD_X        = 0.38    # 제목 박스 위치
HEAD_Y        = 0.33
HEAD_PAD_X    = 0.13
HEAD_PAD_Y    = 0.12

BODY_X        = 0.50    # 본문 불렛("-") 시작 위치
BODY_RIGHT    = 13.08   # 본문 오른쪽 끝
DASH_PREFIX   = "-  "    # 불렛: PowerPoint 불렛 기능 대신 글자로 넣는다
                         # (Canva 등은 사용자 불렛을 동그라미(•)로 바꿔 버림)
BODY_GAP_TOP  = 0.42    # 제목 박스 아래 ~ 본문 시작
CONTENT_BOTTOM = 7.5 * 0.75   # 화면 아래 1/4은 비워 둔다 (멀리서 앞사람 머리에 가려 안 보임)
BODY_BOTTOM   = CONTENT_BOTTOM   # 본문 아래 한계
BODY_TOP_NOHEAD = 0.45  # 제목 없는 슬라이드의 본문 시작

MIN_PICTURE_AREA = 0.02  # 슬라이드 면적 대비 이 비율 미만 이미지는 장식으로 간주

_NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_A = f'{{{_NS_A}}}'


# ── 글자 폭 추정 (Noto Sans KR Bold 실측, em 단위) ─────────────────────────────

_NOTO_W = {
    ' ': .227, '!': .370, '"': .574, '#': .590, '$': .590, '%': .963, '&': .740,
    "'": .325, '(': .378, ')': .378, '*': .507, '+': .590, ',': .325, '-': .370,
    '.': .325, '/': .387, ':': .325, ';': .325, '<': .590, '=': .590, '>': .590,
    '?': .514, '@': 1.007, '[': .378, '\\': .387, ']': .378, '^': .590, '_': .567,
    '`': .626, '{': .378, '|': .296, '}': .378, '~': .590,
    'A': .641, 'B': .681, 'C': .656, 'D': .714, 'E': .615, 'F': .585, 'G': .717,
    'H': .757, 'I': .330, 'J': .568, 'K': .686, 'L': .578, 'M': .853, 'N': .749,
    'O': .770, 'P': .667, 'Q': .770, 'R': .682, 'S': .624, 'T': .625, 'U': .748,
    'V': .619, 'W': .915, 'X': .627, 'Y': .580, 'Z': .613,
    'a': .591, 'b': .644, 'c': .527, 'd': .644, 'e': .581, 'f': .372, 'g': .597,
    'h': .640, 'i': .304, 'j': .306, 'k': .604, 'l': .315, 'm': .964, 'n': .641,
    'o': .626, 'p': .644, 'q': .644, 'r': .437, 's': .495, 't': .421, 'u': .637,
    'v': .576, 'w': .863, 'x': .562, 'y': .574, 'z': .511,
    '‘': .325, '’': .325, '“': .574, '”': .574, '·': .592, '…': 1.0, '–': .544,
    '—': .908,
}

def _char_em(c: str, classic: bool = False) -> float:
    o = ord(c)
    if 0xAC00 <= o <= 0xD7A3 or 0x3130 <= o <= 0x318F:
        return 1.0 if classic else 0.92
    if c.isdigit():
        return 0.62 if classic else 0.59
    if c in _NOTO_W:
        w = _NOTO_W[c]
        return w * 1.05 if classic else w   # 클래식체 폭은 실측값이 없어 조금 넉넉히
    if o >= 0x2E80:
        return 1.0
    return 0.6

def _text_em(s: str, classic: bool = False) -> float:
    return sum(_char_em(c, classic) for c in s)

def _count_lines(text: str, width_em: float, classic: bool = False) -> int:
    """어절(공백) 단위 줄바꿈으로 필요한 줄 수 추정. 한 어절이 줄보다 길면 글자 단위로 끊는다."""
    space = _char_em(' ', classic)
    lines, cur = 1, 0.0
    for word in text.split():
        w = _text_em(word, classic)
        if cur > 0 and cur + space + w <= width_em:
            cur += space + w
            continue
        if cur > 0:
            lines += 1
            cur = 0.0
        while w > width_em:
            lines += 1
            w -= width_em
        cur = w
    return lines


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

def _parse_title_loose(slide):
    """'제목 1' shape이 없는 표지 (예: 이전 변환본을 다시 넣은 경우):
    위에서부터 텍스트를 읽어 '제N강' → 시리즈, 'Rev/Dr.' → 목사, 나머지 → 제목."""
    shapes = sorted((s for s in slide.shapes if s.has_text_frame),
                    key=lambda s: (s.top or 0, s.left or 0))
    texts = [t for s in shapes for p in s.text_frame.paragraphs if (t := _para_text(p))]
    series, pastor, rest = "", "Rev. Dr. Paul Junghoon Lee", []
    for t in texts:
        if "Rev" in t or "Dr." in t:
            pastor = re.sub(r'^[✦★*\s]+', '', t)
        elif not series and re.search(r"제\s*\d+\s*강", t):
            series, remainder = _split_series_title(t)
            if remainder:
                rest.append(remainder)
        else:
            rest.append(t)
    lecture_title = max(rest, key=len) if rest else ""
    return series, lecture_title, pastor

def parse_title_slide(slide):
    title_tf = _tf_by_name(slide, "제목 1")
    sub_tf   = _tf_by_name(slide, "부제목 2")
    if title_tf is None and sub_tf is None:
        return _parse_title_loose(slide)

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


_BULLET_PREFIX_RE = re.compile(r'^[\-–—•·▪◦●○■□➢►✓※]\s*')
# 이 글자로 끝나는 문단은 문장이 이어지는 것으로 보고 다음 문단과 한 불렛으로 합친다
# (완성샘플: "…상징하며," + "가나안 정복은…" → 한 불렛)
_CONTINUATION_ENDINGS = (',', '，', '는', '며', '를')

def _join_runs(runs: list) -> list:
    """인접한 같은 색 run 병합, 빈 run 제거."""
    out = []
    for r in runs:
        if not r['text']:
            continue
        if out and out[-1]['gold'] == r['gold']:
            out[-1] = {'text': out[-1]['text'] + r['text'], 'gold': r['gold']}
        else:
            out.append(dict(r))
    return out

def _para_runs(para) -> list:
    """단락을 [{text, gold}] run 목록으로 반환.
    원본에 <a:highlight>가 있는 run만 gold=True. 줄바꿈(<a:br>)은 공백으로.
    """
    runs = []
    for child in para._p:
        tag = child.tag.split('}')[-1]
        if tag in ('r', 'fld'):
            t_el = child.find(f'{_A}t')
            text = t_el.text if t_el is not None and t_el.text else ''
            rPr = child.find(f'{_A}rPr')
            gold = rPr is not None and rPr.find(f'{_A}highlight') is not None
        elif tag == 'br':
            text, gold = ' ', False
        else:
            continue
        runs.append({'text': re.sub(r'[\s ​]+', ' ', text), 'gold': gold})

    runs = _join_runs(runs)
    if not runs:
        return []
    runs[0]['text'] = _BULLET_PREFIX_RE.sub('', runs[0]['text'].lstrip())
    runs[-1]['text'] = runs[-1]['text'].rstrip()
    runs = _join_runs(runs)
    # 형광펜이 공백에만 걸린 경우는 강조로 보지 않는다
    for r in runs:
        if r['gold'] and not r['text'].strip():
            r['gold'] = False
    runs = _join_runs(runs)
    if not ''.join(r['text'] for r in runs).strip():
        return []
    return runs

def _bullet_full_text(bullet: list) -> str:
    return ''.join(r['text'] for r in bullet)

def _merge_continuations(bullets: list) -> list:
    out = []
    for b in bullets:
        if out and _bullet_full_text(out[-1]).rstrip().endswith(_CONTINUATION_ENDINGS):
            out[-1] = _join_runs(out[-1] + [{'text': ' ', 'gold': False}] + b)
        else:
            out.append(b)
    return out


def _placeholder_type(shape):
    if not shape.is_placeholder:
        return None
    try:
        return shape.placeholder_format.type
    except Exception:
        return None

_TITLE_PH = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.VERTICAL_TITLE}
_SKIP_PH  = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE,
             PP_PLACEHOLDER.HEADER}

def _is_title_shape(shape) -> bool:
    return (_placeholder_type(shape) in _TITLE_PH
            or shape.name.startswith(("제목", "Title")))

def _picture_crop(el):
    """<a:srcRect> 자르기 비율 (l, t, r, b) — 0~1."""
    src = el.find(f'.//{_A}srcRect')
    if src is None:
        return (0.0, 0.0, 0.0, 0.0)
    vals = []
    for k in ('l', 't', 'r', 'b'):
        v = int(src.get(k, '0') or 0) / 100000.0
        vals.append(min(max(v, 0.0), 0.95))
    return tuple(vals)

def _iter_shapes(shapes, transform):
    """그룹 내부까지 (shape, (left, top, width, height)) 를 슬라이드 좌표로 반환."""
    for shape in shapes:
        try:
            box = transform(shape.left or 0, shape.top or 0,
                            shape.width or 0, shape.height or 0)
        except Exception:
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            xfrm = shape._element.grpSpPr.find(f'{_A}xfrm')
            if xfrm is None:
                continue
            off, ext = xfrm.find(f'{_A}off'), xfrm.find(f'{_A}ext')
            choff, chext = xfrm.find(f'{_A}chOff'), xfrm.find(f'{_A}chExt')
            if None in (off, ext, choff, chext):
                continue
            ox, oy = int(off.get('x')), int(off.get('y'))
            ex, ey = int(ext.get('cx')), int(ext.get('cy'))
            cx, cy = int(choff.get('x')), int(choff.get('y'))
            cw, ch = int(chext.get('cx')) or 1, int(chext.get('cy')) or 1

            def child_tf(l, t, w, h, ox=ox, oy=oy, ex=ex, ey=ey, cx=cx, cy=cy, cw=cw, ch=ch):
                return transform(ox + (l - cx) * ex / cw, oy + (t - cy) * ey / ch,
                                 w * ex / cw, h * ey / ch)
            yield from _iter_shapes(shape.shapes, child_tf)
        else:
            yield shape, box

def _extract_pictures(slide, scale_x, scale_y) -> list:
    """의미 있는 크기의 이미지를 After 좌표(EMU)·원본 자르기 정보와 함께 추출."""
    def to_target(l, t, w, h):
        return (int(l * scale_x), int(t * scale_y), int(w * scale_x), int(h * scale_y))

    slide_area = int(SLIDE_W) * int(SLIDE_H)
    pics = []
    for shape, (l, t, w, h) in _iter_shapes(slide.shapes, to_target):
        el = shape._element
        tag = el.tag.split('}')[-1]
        if tag not in ('pic', 'sp'):
            continue
        blip = el.find(f'.//{_A}blip')
        if blip is None:
            continue
        rId = blip.get(f'{{{_NS_R}}}embed')
        if not rId or rId not in slide.part.rels:
            continue
        if w * h < slide_area * MIN_PICTURE_AREA:
            continue
        try:
            blob = slide.part.rels[rId].target_part.blob
        except Exception:
            continue
        try:
            px_size = Image.open(io.BytesIO(blob)).size
        except Exception:
            px_size = None
        pics.append({'blob': blob, 'crop': _picture_crop(el), 'px_size': px_size,
                     'left': l, 'top': t, 'width': w, 'height': h})
    return pics

# 하위 호환용 (이전 버전 API)
def _extract_images(slide, scale_x=1.0, scale_y=1.0) -> list:
    return [{'bytes': p['blob'], 'left': p['left'], 'top': p['top'],
             'width': p['width'], 'height': p['height']}
            for p in _extract_pictures(slide, scale_x, scale_y)]

def parse_content_slide(slide, scale_x=1.0, scale_y=1.0):
    """슬라이드 → dict(heading, has_title_shape, bullets, pictures)."""
    title_shapes, text_shapes = [], []
    for shape in slide.shapes:
        if not shape.has_text_frame or _placeholder_type(shape) in _SKIP_PH:
            continue
        if _is_title_shape(shape):
            title_shapes.append(shape)
        elif shape.text_frame.text.strip():
            text_shapes.append(shape)
    text_shapes.sort(key=lambda s: ((s.top or 0), (s.left or 0)))

    # 제목 안의 줄바꿈(<a:br>)은 공백으로 이어 붙인다
    heading = " ".join(
        _bullet_full_text(_para_runs(p)) for s in title_shapes for p in s.text_frame.paragraphs
    )
    heading = re.sub(r'\s+', ' ', heading).strip()

    paras = [p for s in text_shapes for p in s.text_frame.paragraphs if _para_text(p)]

    # 제목 shape 자체가 없으면 첫 단락이 제목 역할 (마태복음 42강 6번 슬라이드)
    if not title_shapes and len(paras) >= 2 and len(_para_text(paras[0])) <= 30:
        heading = _para_text(paras[0])
        paras = paras[1:]

    bullets = [b for p in paras if (b := _para_runs(p))]
    bullets = _merge_continuations(bullets)

    return {
        'heading': heading,
        'has_title_shape': bool(title_shapes),
        'bullets': bullets,
        'pictures': _extract_pictures(slide, scale_x, scale_y),
    }

def _is_image_only_pptx(slides) -> bool:
    """모든 슬라이드가 이미지 전용(텍스트 shape 없음)이면 True."""
    for slide in slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                return False
    return True

def parse_before(path: str) -> dict:
    prs = Presentation(path)
    all_slides = list(prs.slides)

    if _is_image_only_pptx(all_slides):
        raise ValueError(
            "이 파일은 이미지 기반 슬라이드로만 구성되어 있습니다.\n"
            "이미 변환된 파일(After/Complete)이나 1차 작업물로 보입니다.\n"
            "텍스트 shape이 있는 원본 Before 파일을 사용해 주세요."
        )

    # 소스 슬라이드 크기가 After 기준(SLIDE_W×SLIDE_H)과 다를 경우 좌표 스케일링
    scale_x = int(SLIDE_W) / int(prs.slide_width)
    scale_y = int(SLIDE_H) / int(prs.slide_height)

    series, lecture_title, pastor = parse_title_slide(all_slides[0])

    content_slides = []
    prev_heading = ""
    for src, slide in enumerate(all_slides[1:], start=2):
        cs = parse_content_slide(slide, scale_x, scale_y)
        heading, bullets, pictures = cs['heading'], cs['bullets'], cs['pictures']

        if bullets or heading:
            # 글자가 있는 슬라이드에서 화면 전체를 덮는 그림은 배경으로 보고 제외
            full = int(SLIDE_W) * int(SLIDE_H) * 0.85
            pictures = [p for p in pictures if p['width'] * p['height'] < full]

        if not bullets and not heading:
            if not pictures:
                continue                      # 빈 슬라이드
            kind = 'image'                    # 지도·도표 등 이미지 전용 → 그대로 유지
        elif pictures:
            kind = 'text_image'
        else:
            kind = 'text'
            # 제목이 비어 있으면 앞 슬라이드 제목을 이어서 사용 (완성샘플 방식)
            if not heading and cs['has_title_shape']:
                heading = prev_heading

        if kind == 'text' and heading:
            prev_heading = heading

        content_slides.append({
            "src": src, "kind": kind, "heading": heading,
            "bullets": bullets, "pictures": pictures,
            "images": pictures,  # 하위 호환
        })

    return {
        "title": {"series": series, "lecture_title": lecture_title, "pastor": pastor},
        "content_slides": content_slides,
        "n_source_slides": len(all_slides),
    }


# ── 텍스트 배치 계산 ──────────────────────────────────────────────────────────

def _heading_layout(heading: str, max_w_in: float):
    """제목 → (글자크기, 줄 수, 박스 폭, 박스 높이) 인치.
    한 줄에 아깝게 안 들어가면 34pt까지 줄여 한 줄로, 그래도 길면 두 줄(넘치면 더 작게)."""
    inner_w = max_w_in - 2 * HEAD_PAD_X

    def n_lines(sz):
        return _count_lines(heading, inner_w * 72 / sz, classic=True)

    size = next((s for s in (HEAD_PT, 36, 34) if n_lines(s) == 1), None)
    if size is None:
        size = next((s for s in (HEAD_PT, 34, 30) if n_lines(s) <= 2), 30)
    lines = n_lines(size)
    text_w = _text_em(heading, classic=True) * size / 72
    if lines == 1:
        box_w = min(max_w_in, text_w * 1.10 + 2 * HEAD_PAD_X + 0.05)
    else:
        box_w = max_w_in
    line_pt = HEAD_LINE_PT * size / HEAD_PT
    box_h = 2 * HEAD_PAD_Y + lines * line_pt / 72
    return size, lines, box_w, box_h

_WIDTH_SAFETY = 0.95   # 글꼴·렌더러 차이를 감안해 폭을 조금 좁게 잡고 줄 수를 추정

def _hang_in(size) -> float:
    """"-  " 글자 폭(인치) = 내어쓰기 폭."""
    return _text_em(DASH_PREFIX) * size / 72

def _bullet_lines(bullets, size, width_in):
    width_em = (width_in - _hang_in(size)) * _WIDTH_SAFETY * 72 / size
    return [_count_lines(_bullet_full_text(b), width_em) for b in bullets]

def _block_height_pt(lines, size):
    pitch = size * LINE_RATIO
    return sum(lines) * pitch + max(len(lines) - 1, 0) * MIN_GAP_RATIO * pitch

def _fit_size(bullets, width_in, avail_in, sizes=None):
    """bullets가 들어가는 가장 큰 글자 크기. 없으면 None."""
    for size in sizes or range(BODY_PT_MAX, BODY_PT_MIN - 1, -2):
        if _block_height_pt(_bullet_lines(bullets, size, width_in), size) <= avail_in * 72:
            return size
    return None

def _partition(heights, k):
    """연속 구간 k개로 나눠 가장 높은 구간을 최소화 → 구간 경계 인덱스 목록."""
    n = len(heights)
    pre = [0]
    for h in heights:
        pre.append(pre[-1] + h)
    INF = float('inf')
    best = [[INF] * (n + 1) for _ in range(k + 1)]
    cut = [[0] * (n + 1) for _ in range(k + 1)]
    best[0][0] = 0
    for j in range(1, k + 1):
        for i in range(1, n + 1):
            for m in range(j - 1, i):
                v = max(best[j - 1][m], pre[i] - pre[m])
                if v < best[j][i]:
                    best[j][i], cut[j][i] = v, m
    bounds, i = [], n
    for j in range(k, 0, -1):
        bounds.append((cut[j][i], i))
        i = cut[j][i]
    return list(reversed(bounds))

def plan_chunks(bullets, width_in, avail_in):
    """한 슬라이드에 넣을 수 있으면 [(bullets, size)].
    32→28pt로도 안 들어갈 때만 높이가 고르게 나뉘도록 분리한다."""
    if not bullets:
        return [([], BODY_PT_MAX)]
    size = _fit_size(bullets, width_in, avail_in)
    if size:
        return [(bullets, size)]

    pitch = BODY_PT_MAX * LINE_RATIO
    heights = [l * pitch + MIN_GAP_RATIO * pitch
               for l in _bullet_lines(bullets, BODY_PT_MAX, width_in)]
    for k in range(2, len(bullets) + 1):
        groups = [bullets[a:b] for a, b in _partition(heights, k)]
        sizes = [_fit_size(g, width_in, avail_in) for g in groups]
        if all(sizes):
            common = min(sizes)   # 나뉜 슬라이드끼리 글자 크기 통일
            return [(g, common) for g in groups]

    # 불렛 하나가 너무 길면 크기를 더 줄여서라도 한 장에
    out = []
    for b in bullets:
        s = _fit_size([b], width_in, avail_in, sizes=range(BODY_PT_MIN, 19, -2)) or 20
        out.append(([b], s))
    return out


# ── After PPTX 빌더 ───────────────────────────────────────────────────────────

_BG_PNG = None

def _bg_png() -> bytes:
    """세로 그라데이션 배경 이미지 (모든 뷰어에서 동일하게 보이도록 그림으로 깐다)."""
    global _BG_PNG
    if _BG_PNG is None:
        h = 1080
        img = Image.new('RGB', (16, h))
        top, bot = tuple(BG_TOP), tuple(BG_BOTTOM)
        for y in range(h):
            t = y / (h - 1)
            c = tuple(round(a + (b - a) * t) for a, b in zip(top, bot))
            for x in range(16):
                img.putpixel((x, y), c)
        buf = io.BytesIO()
        img.save(buf, 'PNG')
        _BG_PNG = buf.getvalue()
    return _BG_PNG

def _solid_bg(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def _add_background(slide):
    _solid_bg(slide, BG_COLOR)
    slide.shapes.add_picture(io.BytesIO(_bg_png()), 0, 0, SLIDE_W, SLIDE_H)

def _set_font(run, name, size_pt, color, bold=None, italic=None):
    """latin/ea/cs 글꼴을 모두 지정 (한글은 ea 글꼴로 그려지므로 필수)."""
    f = run.font
    f.size = Pt(size_pt)
    f.color.rgb = color
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    rPr = run._r.get_or_add_rPr()
    rPr.set('lang', 'ko-KR')
    for tag in ('latin', 'ea', 'cs'):
        el = rPr.find(f'{_A}{tag}')
        if el is None:
            el = rPr.makeelement(f'{_A}{tag}', {})
            rPr.append(el)
        el.set('typeface', name)

def _plain_frame(tf, anchor=MSO_ANCHOR.TOP, wrap=True):
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor

def _add_textbox(slide, left, top, width, height,
                 text, font_name, size_pt,
                 bold=None, italic=None, color=WHITE,
                 align=PP_ALIGN.LEFT, word_wrap=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf  = box.text_frame
    _plain_frame(tf, wrap=word_wrap)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    _set_font(run, font_name, size_pt, color, bold=bold, italic=italic)
    return box

def _add_heading(slide, heading: str, x_in: float, max_w_in: float) -> float:
    """글자 폭에 맞춘 박스 제목. (박스 아래쪽 y, 박스 오른쪽 x) 인치를 반환."""
    size, lines, box_w, box_h = _heading_layout(heading, max_w_in)
    box = slide.shapes.add_textbox(Inches(x_in), Inches(HEAD_Y), Inches(box_w), Inches(box_h))
    box.fill.solid()
    box.fill.fore_color.rgb = HEADER_BG
    box.line.fill.background()
    tf = box.text_frame
    # 한 줄 제목은 줄바꿈 금지 (글꼴 폭이 추정보다 넓어도 박스 밖으로 두 줄이 되지 않게)
    _plain_frame(tf, anchor=MSO_ANCHOR.MIDDLE, wrap=(lines > 1))
    tf.margin_left = tf.margin_right = Inches(HEAD_PAD_X)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = Pt(HEAD_LINE_PT * size / HEAD_PT)
    run = p.add_run()
    run.text = heading
    _set_font(run, FONT_CLASSIC_MED, size, WHITE, bold=False)
    return HEAD_Y + box_h, x_in + box_w

def _bullet_ppr(p, size):
    """불렛 기능은 끄고(buNone) "-  " 폭만큼 내어쓰기 → 줄바꿈된 줄이 글자 시작에 맞춰짐."""
    hang = int(Inches(_hang_in(size)))
    pPr = p._p.get_or_add_pPr()
    pPr.set('marL', str(hang))
    pPr.set('indent', str(-hang))
    for tag in ('buClr', 'buSzPct', 'buFontTx', 'buFont', 'buChar', 'buNone', 'buAutoNum'):
        for el in pPr.findall(f'{_A}{tag}'):
            pPr.remove(el)
    pPr.append(pPr.makeelement(f'{_A}buNone', {}))

def _add_body(slide, bullets, size, left_in, right_in, top_in, bottom_in):
    """불렛을 영역 안에 배치. 남는 공간은 불렛 사이 간격 → 위쪽 여백 순으로 배분."""
    if not bullets:
        return
    width_in = right_in - left_in
    avail_pt = (bottom_in - top_in) * 72
    pitch = size * LINE_RATIO
    lines = _bullet_lines(bullets, size, width_in)
    text_pt = sum(lines) * pitch
    n = len(bullets)
    free = avail_pt - text_pt
    gap = 0.0
    if n > 1:
        gap = min(MAX_GAP_RATIO * pitch, max(MIN_GAP_RATIO * pitch, free * 0.75 / (n - 1)))
    used = text_pt + (n - 1) * gap
    top_off = max(0.0, (avail_pt - used) * 0.25)

    box_top = top_in + top_off / 72
    box_h = max(used / 72 + 0.3, bottom_in - box_top)
    box = slide.shapes.add_textbox(Inches(left_in), Inches(box_top),
                                   Inches(width_in), Inches(box_h))
    tf = box.text_frame
    _plain_frame(tf)
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment    = PP_ALIGN.LEFT
        p.line_spacing = Pt(pitch)
        p.space_before = Pt(0 if i == 0 else gap)
        p.space_after  = Pt(0)
        _bullet_ppr(p, size)
        dash = p.add_run()
        dash.text = DASH_PREFIX
        _set_font(dash, FONT_NOTO, size, BODY_COLOR, bold=True)
        for rd in bullet:
            run = p.add_run()
            run.text = rd['text']
            _set_font(run, FONT_NOTO, size, GOLD if rd['gold'] else BODY_COLOR, bold=True)

def _place_picture(slide, pic, left, top, width, height, cover=False):
    """이미지 배치. cover=True면 영역을 꽉 채우도록 가운데 기준으로 잘라낸다."""
    l, t, r, b = pic['crop']
    shape = slide.shapes.add_picture(io.BytesIO(pic['blob']), int(left), int(top),
                                     int(width), int(height))
    cl = ct = cr = cb = 0.0
    if cover:
        if pic['px_size']:
            pw, ph = pic['px_size']
            vis_aspect = (pw * (1 - l - r)) / max(ph * (1 - t - b), 1)
        else:
            vis_aspect = pic['width'] / max(pic['height'], 1)
        box_aspect = width / max(height, 1)
        if vis_aspect > box_aspect:            # 이미지가 더 넓음 → 좌우 자르기
            extra = 1 - box_aspect / vis_aspect
            cl = cr = extra / 2
        else:                                  # 이미지가 더 높음 → 위 30% / 아래 70%
            extra = 1 - vis_aspect / box_aspect
            ct, cb = extra * 0.3, extra * 0.7
    shape.crop_left   = l + (1 - l - r) * cl
    shape.crop_right  = r + (1 - l - r) * cr
    shape.crop_top    = t + (1 - t - b) * ct
    shape.crop_bottom = b + (1 - t - b) * cb
    return shape

def build_title_slide(prs, series: str, lecture_title: str, pastor: str,
                      cover_image: str = None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    if cover_image and os.path.exists(cover_image):
        # 표지.png 를 슬라이드 전체 배경으로 (성경책·별 장식 모두 포함됨)
        slide.shapes.add_picture(cover_image, 0, 0, SLIDE_W, SLIDE_H)
    else:
        _add_background(slide)

    full_w = SLIDE_W - Inches(1.0)

    # ── 시리즈명 (예: 마태복음 제62강)
    _add_textbox(slide, Inches(0.5), Inches(1.60), full_w, Inches(0.80),
                 series, FONT_CLASSIC_MED, 40, bold=False, color=BODY_COLOR, align=PP_ALIGN.CENTER)

    # ── 강의 제목 (가운데, 크게) — 길면 한 줄에 맞도록 줄인다
    title_pt = 64
    while title_pt > 40 and _text_em(lecture_title, classic=True) * title_pt / 72 > 12.0:
        title_pt -= 4
    _add_textbox(slide, Inches(0.5), Inches(2.85), full_w, Inches(1.30),
                 lecture_title, FONT_CLASSIC_MED, title_pt,
                 bold=False, color=BODY_COLOR, align=PP_ALIGN.CENTER)

    # ── 목사님 이름 (크림색)
    _add_textbox(slide, Inches(0.5), Inches(4.50), full_w, Inches(0.80),
                 pastor, FONT_CLASSIC_MED, 36, bold=False, color=PASTOR_COLOR, align=PP_ALIGN.CENTER)
    return slide

def build_image_slide(prs, pictures):
    """이미지 전용 슬라이드: 원본 배치를 그대로 옮기되,
    아래 1/4을 침범하면 전체를 비율 그대로 줄여 위쪽 3/4 안에 가운데 정렬."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_background(slide)
    limit = int(Inches(CONTENT_BOTTOM))
    bottom = max(p['top'] + p['height'] for p in pictures)
    scale = min(1.0, limit / bottom) if bottom > 0 else 1.0
    if scale < 1.0:
        cx = (min(p['left'] for p in pictures)
              + max(p['left'] + p['width'] for p in pictures)) / 2
        for pic in pictures:
            left = int(SLIDE_W) / 2 + (pic['left'] - cx) * scale
            _place_picture(slide, pic, left, pic['top'] * scale,
                           pic['width'] * scale, pic['height'] * scale)
    else:
        for pic in pictures:
            _place_picture(slide, pic, pic['left'], pic['top'], pic['width'], pic['height'])
    return slide

SW_IN = 13.333   # 슬라이드 폭 (인치)

def build_content_slide(prs, heading: str, bullets: list, size: int = BODY_PT_MAX):
    """제목 + 불렛 슬라이드."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_background(slide)
    if heading:
        head_bottom, _ = _add_heading(slide, heading, HEAD_X, SW_IN - 2 * HEAD_X)
        top = head_bottom + BODY_GAP_TOP
    else:
        top = BODY_TOP_NOHEAD
    _add_body(slide, bullets, size, BODY_X, BODY_RIGHT, top, BODY_BOTTOM)
    return slide

def _text_geometry(heading):
    """전체 폭 텍스트 슬라이드의 (본문 폭, 본문 높이) 인치."""
    top = BODY_TOP_NOHEAD
    if heading:
        _, _, _, box_h = _heading_layout(heading, SW_IN - 2 * HEAD_X)
        top = HEAD_Y + box_h + BODY_GAP_TOP
    return BODY_RIGHT - BODY_X, BODY_BOTTOM - top

def _image_column(pictures):
    """이미지 열 (side, 기본 폭 인치): 원본 이미지 중심이 오른쪽이면 오른쪽 열.
    폭은 원본 이미지 폭을 따르되 슬라이드의 35~45% (완성샘플 41~49%)."""
    W = int(SLIDE_W)
    xmin = min(p['left'] for p in pictures)
    xmax = max(p['left'] + p['width'] for p in pictures)
    side = 'right' if (xmin + xmax) / 2 >= W / 2 else 'left'
    width_in = Emu(xmax - xmin).inches
    width_in = min(max(width_in, SW_IN * 0.35), SW_IN * 0.45)
    return side, width_in

def _text_image_layout(heading, side, col_w):
    """텍스트+이미지 슬라이드 배치 (인치)."""
    col_x = SW_IN - col_w if side == 'right' else 0.0
    lay = {
        'side': side, 'col_w': col_w, 'col_x': col_x,
        'text_left':  BODY_X if side == 'right' else col_w + BODY_X,
        'text_right': col_x - 0.35 if side == 'right' else SW_IN - 0.30,
        'head_x': HEAD_X if side == 'right' else col_w + HEAD_X,
        'img_top': HEAD_Y, 'body_top': BODY_TOP_NOHEAD,
    }
    lay['head_max'] = SW_IN - lay['head_x'] - HEAD_X
    if heading:
        _, _, box_w, box_h = _heading_layout(heading, lay['head_max'])
        head_bottom = HEAD_Y + box_h
        lay['body_top'] = head_bottom + BODY_GAP_TOP
        # 제목 박스가 이미지 열을 침범하면 이미지는 제목 아래부터 (팀 켈러 슬라이드)
        if side == 'left' or lay['head_x'] + box_w > col_x - 0.15:
            lay['img_top'] = head_bottom
    return lay

def build_text_image_slide(prs, heading, bullets, pictures, size, lay):
    """텍스트 + 이미지: 원본의 좌우 배치를 유지해 텍스트 열 / 이미지 열로 나눈다."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_background(slide)
    if heading:
        _add_heading(slide, heading, lay['head_x'], lay['head_max'])

    # 이미지 열: 제목 높이부터 화면 3/4 지점까지·옆 끝까지 채움, 여러 장이면 세로로 쌓기
    ordered = sorted(pictures, key=lambda p: (p['top'], p['left']))
    gap = 0.08
    total_h = CONTENT_BOTTOM - lay['img_top']   # 이미지도 아래 1/4은 비움
    each_h = (total_h - gap * (len(ordered) - 1)) / len(ordered)
    for i, pic in enumerate(ordered):
        y = lay['img_top'] + i * (each_h + gap)
        _place_picture(slide, pic, Inches(lay['col_x']), Inches(y),
                       Inches(lay['col_w']), Inches(each_h), cover=True)

    _add_body(slide, bullets, size, lay['text_left'], lay['text_right'],
              lay['body_top'], BODY_BOTTOM)
    return slide

def build_after(data: dict, output_path: str, cover_image: str = None):
    """After PPTX 저장. (전체 슬라이드 수, slide_map) 반환.
    slide_map[i] = 원본 i+1번 슬라이드가 변환된 슬라이드 번호 목록 (1부터)."""
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    n_src = data.get("n_source_slides") or (len(data["content_slides"]) + 1)
    slide_map = [[] for _ in range(n_src)]

    t = data["title"]
    build_title_slide(prs, t["series"], t["lecture_title"], t["pastor"],
                      cover_image=cover_image)
    slide_map[0].append(1)

    for idx, cs in enumerate(data["content_slides"]):
        src = cs.get("src", idx + 2)
        kind = cs.get("kind", "text")
        heading = cs.get("heading", "")
        bullets = cs.get("bullets", [])
        pictures = cs.get("pictures", cs.get("images", []))
        out = slide_map[src - 1] if src - 1 < len(slide_map) else []

        if kind == 'image':
            build_image_slide(prs, pictures)
            out.append(len(prs.slides))
            continue

        if kind == 'text_image':
            # 글이 넘치면 분리하기 전에 이미지 열을 좁혀 본다
            side, col_w = _image_column(pictures)
            for cw in (col_w, SW_IN * 0.35):
                lay = _text_image_layout(heading, side, min(col_w, cw))
                width_in = lay['text_right'] - lay['text_left']
                avail_in = BODY_BOTTOM - lay['body_top']
                if _fit_size(bullets, width_in, avail_in):
                    break
        else:
            width_in, avail_in = _text_geometry(heading)

        chunks = plan_chunks(bullets, width_in, avail_in)
        for j, (chunk, size) in enumerate(chunks):
            if kind == 'text_image' and j == 0:
                build_text_image_slide(prs, heading, chunk, pictures, size, lay)
            else:
                if kind == 'text_image':
                    # 이미지 슬라이드에서 넘친 불렛은 전체 폭 텍스트 슬라이드로
                    size = _fit_size(chunk, *_text_geometry(heading)) or size
                build_content_slide(prs, heading, chunk, size)
            out.append(len(prs.slides))

    prs.save(output_path)
    return len(prs.slides), slide_map


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


def describe_map(slide_map) -> list:
    """['원본 1 → 1', '원본 5 → 5, 6 (1 → 2페이지)', ...]"""
    lines = []
    for i, outs in enumerate(slide_map, start=1):
        if not outs:
            lines.append(f"원본 {i} → (빈 슬라이드, 제외)")
        elif len(outs) == 1:
            lines.append(f"원본 {i} → {outs[0]}")
        else:
            lines.append(f"원본 {i} → {', '.join(map(str, outs))}  (1 → {len(outs)}페이지)")
    return lines


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

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
    try:
        data = parse_before(args.input)
    except ValueError as e:
        print(f"오류: {e}")
        sys.exit(1)
    t = data["title"]
    print(f"      시리즈: {t['series']}")
    print(f"      제목  : {t['lecture_title']}")
    print(f"      목사  : {t['pastor']}")
    print(f"      내용 슬라이드: {len(data['content_slides'])}개 (분리 전)")

    # 표지 이미지: cover.png → 표지.png 순으로 탐색
    script_dir  = os.path.dirname(os.path.abspath(__file__))
    cover_image = next(
        (os.path.join(script_dir, n) for n in ("cover.png", "표지.png")
         if os.path.exists(os.path.join(script_dir, n))),
        None,
    )
    if cover_image:
        print(f"[2/3] 변환 중... (표지 이미지 적용)")
    else:
        print(f"[2/3] 변환 중... (표지 이미지 없음, 기본 배경 사용)")

    n, slide_map = build_after(data, out_path, cover_image=cover_image)
    print(f"[3/3] 완료 → {os.path.basename(out_path)}  "
          f"(원본 {len(slide_map)}장 → {n}장)")
    for line in describe_map(slide_map):
        if "→" in line and ("페이지" in line or "제외" in line):
            print(f"      {line}")

    if args.png:
        png_dir = args.png_dir or os.path.splitext(out_path)[0] + "_slides"
        print(f"[+]  PNG 내보내기 → {png_dir}")
        export_png(out_path, png_dir)


if __name__ == "__main__":
    main()
