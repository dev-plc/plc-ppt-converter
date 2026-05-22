"""PPTX 구조 분석 스크립트 - before/after 파일 비교용"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

def rgb_to_hex(rgb):
    if rgb is None:
        return None
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"

def analyze_pptx(path):
    prs = Presentation(path)
    w = prs.slide_width.inches
    h = prs.slide_height.inches
    print(f"\n=== {path} ===")
    print(f"슬라이드 크기: {w:.2f}\" x {h:.2f}\" ({prs.slide_width.emu} x {prs.slide_height.emu} EMU)")
    print(f"슬라이드 수: {len(prs.slides)}")

    for i, slide in enumerate(prs.slides):
        print(f"\n--- 슬라이드 {i+1} ---")
        bg = slide.background
        fill = bg.fill
        print(f"  배경 fill type: {fill.type}")
        try:
            if fill.fore_color and fill.fore_color.rgb:
                print(f"  배경색: {rgb_to_hex(fill.fore_color.rgb)}")
        except:
            pass

        for j, shape in enumerate(slide.shapes):
            print(f"  [Shape {j}] name={shape.name!r} type={shape.shape_type} "
                  f"left={shape.left} top={shape.top} w={shape.width} h={shape.height}")

            try:
                sf = shape.fill
                if sf.type is not None:
                    try:
                        print(f"    fill_type={sf.type} color={rgb_to_hex(sf.fore_color.rgb)}")
                    except:
                        print(f"    fill_type={sf.type}")
            except:
                pass

            if shape.has_text_frame:
                tf = shape.text_frame
                for k, para in enumerate(tf.paragraphs):
                    for run in para.runs:
                        font = run.font
                        try:
                            color = rgb_to_hex(font.color.rgb) if font.color and font.color.type else None
                        except:
                            color = None
                        size = round(font.size.pt, 1) if font.size else None
                        bold = font.bold
                        italic = font.italic
                        name = font.name
                        text = run.text[:60] if run.text else ""
                        if text.strip():
                            print(f"    para[{k}] run: font={name!r} size={size}pt bold={bold} italic={italic} "
                                  f"color={color} | {text!r}")

if __name__ == "__main__":
    base = r"C:\Users\myc43\OneDrive - ETERNAL LIBRTY POLICY INSTITUTE\바탕 화면\plc ppt"
    files = [
        f"{base}\\에베소서 제4강(before).pptx",
        f"{base}\\에베소서 제4강(after).pptx",
    ]
    for f in files:
        try:
            analyze_pptx(f)
        except Exception as e:
            print(f"오류: {f}: {e}")
