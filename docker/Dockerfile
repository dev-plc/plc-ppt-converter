FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends git \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# GitHub에서 최초 클론
RUN git clone --depth=1 https://github.com/dev-plc/plc-ppt-converter.git .

# 패키지 설치 (requirements.txt 변경 시에만 --build 필요)
RUN pip install --no-cache-dir -r requirements.txt

# 표지.png → cover.png (한글 파일명 인코딩 우회)
RUN cp 표지.png cover.png

RUN chmod +x docker/entrypoint.sh

CMD ["/bin/sh", "docker/entrypoint.sh"]
