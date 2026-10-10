FROM python:3.12-slim

# git: 코드 자동 업데이트 / libreoffice-impress·poppler-utils: 원본↔변환 비교 이미지
# fonts-noto-cjk: 미리보기 한글 글꼴
RUN apt-get update && apt-get install -y --no-install-recommends \
        git libreoffice-impress poppler-utils fonts-noto-cjk fontconfig \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# GitHub에서 최초 클론
RUN git clone --depth=1 https://github.com/dev-plc/plc-ppt-converter.git .

# 패키지 설치 (entrypoint에서도 매번 확인)
RUN pip install --no-cache-dir -r requirements.txt

# 표지.png → cover.png (한글 파일명 인코딩 우회)
RUN cp 표지.png cover.png

# 미리보기용 글꼴 대체 설정
RUN cp docker/fonts.conf /etc/fonts/local.conf && fc-cache -f

RUN chmod +x docker/entrypoint.sh

CMD ["/bin/sh", "docker/entrypoint.sh"]
