"""슬라이드 미리보기 이미지 생성 (LibreOffice → PDF → JPG)

텔레그램 봇에서 원본/변환 슬라이드를 짝지어 보내는 데 사용.
LibreOffice(soffice)와 poppler(pdftoppm)가 없으면 None을 반환하므로
호출하는 쪽은 미리보기 없이 계속 진행하면 된다.

Usage (로컬 확인용):
    python preview.py input.pptx out_dir
"""

import glob
import os
import shutil
import subprocess
import sys
import tempfile


def _soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def available() -> bool:
    return bool(_soffice() and shutil.which("pdftoppm"))


def render_slides(pptx_path: str, out_dir: str, width: int = 1280,
                  timeout: int = 300):
    """PPTX의 각 슬라이드를 JPG로 저장하고 경로 목록(슬라이드 순서)을 반환.
    도구가 없거나 실패하면 None."""
    soffice, pdftoppm = _soffice(), shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        return None
    os.makedirs(out_dir, exist_ok=True)

    # 한글·공백 파일명 문제를 피하려고 ASCII 이름으로 복사해서 변환
    work = tempfile.mkdtemp(prefix="pv_", dir=out_dir)
    src = os.path.join(work, "deck.pptx")
    shutil.copyfile(pptx_path, src)
    # 동시에 여러 변환이 돌아도 충돌하지 않도록 매번 별도 프로필 사용
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        subprocess.run(
            [soffice, f"-env:UserInstallation=file://{profile}", "--headless",
             "--convert-to", "pdf", "--outdir", work, src],
            check=True, timeout=timeout,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        pdf = os.path.join(work, "deck.pdf")
        if not os.path.exists(pdf):
            return None
        prefix = os.path.join(work, "slide")
        subprocess.run(
            [pdftoppm, "-jpeg", "-jpegopt", "quality=85",
             "-scale-to-x", str(width), "-scale-to-y", "-1", pdf, prefix],
            check=True, timeout=timeout,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except (subprocess.SubprocessError, OSError):
        return None
    finally:
        shutil.rmtree(profile, ignore_errors=True)

    files = glob.glob(prefix + "-*.jpg")
    files.sort(key=lambda p: int(p.rsplit("-", 1)[1].split(".")[0]))
    return files or None


def pair_slides(orig_imgs: list, conv_imgs: list, slide_map: list) -> list:
    """원본 슬라이드마다 (원본 이미지, [변환 이미지...]) 목록.
    slide_map[i] = 원본 i+1번이 변환된 슬라이드 번호(1부터) 목록."""
    pairs = []
    for i, outs in enumerate(slide_map):
        orig = orig_imgs[i] if i < len(orig_imgs) else None
        convs = [conv_imgs[n - 1] for n in outs if 0 < n <= len(conv_imgs)]
        pairs.append((orig, convs))
    return pairs


def split_caption(n_pages: int) -> str:
    """원본 1장이 여러 장으로 나뉘었을 때 붙이는 설명."""
    return f"1 → {n_pages}페이지"


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    imgs = render_slides(sys.argv[1], sys.argv[2])
    if imgs is None:
        print("미리보기 생성 실패 (soffice / pdftoppm 확인)")
        sys.exit(1)
    for p in imgs:
        print(p)
