# PL Church PPT Converter — 프로젝트 컨텍스트

## 작업 폴더
`C:\Users\myc43\OneDrive - ETERNAL LIBRTY POLICY INSTITUTE\바탕 화면\plc ppt`

## GitHub
- 레포: https://github.com/dev-plc/plc-ppt-converter
- 계정: dev-plc (dev@plch.or.kr)
- 브랜치: master
- 동기화: `git pull` / `git push`

## 핵심 파일
- `convert_pptx.py` — 메인 변환 스크립트 (Before PPTX → After PPTX + PNG)
- `표지.png` — 타이틀 슬라이드 배경 (성경책 아이콘 + 별 장식 포함)
- `samples/` — Before 샘플 파일 3개 (에베소서, 레위기, 마태복음)

## 환경
- Python 3.12+
- 패키지: python-pptx, pillow, comtypes, fonttools
- 폰트: 국립박물관문화재단클래식 B/M (사용자 폰트 디렉토리 설치됨)
- OS: Windows 11, PowerShell

## 실행 방법
```bash
python convert_pptx.py samples/에베소서\ 제4강\(before\).pptx
python convert_pptx.py input.pptx output.pptx --png
```

## 디자인 상수 (convert_pptx.py 상단)
- BG_COLOR   = #1B2045  (배경)
- HEADER_BG  = #232A56  (헤더 박스)
- GOLD       = #F5C518  (핵심어 강조)
- MAX_BULLETS = 3       (슬라이드당 최대 불렛)
- 본문 폰트 크기: 28pt Bold

## Before PPTX 파싱 주의
- 에베소서/레위기: Shape '제목 1'에 시리즈명+강의제목 합쳐짐
- 마태복음: Shape '제목 1'=시리즈명, Shape '부제목 2'=강의제목+목사명 (분리)
- 핵심어 감지: 28자 이하, 따옴표 없음, 성경구절 패턴 없음 → 골드 색상
