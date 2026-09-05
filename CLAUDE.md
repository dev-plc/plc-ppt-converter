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

## 개발 기록 (옵시디언 연동)

이 저장소의 과거 개발 기록, 트러블슈팅 노트, 설계 결정 사항은 아래 경로에
정리되어 있습니다:

G:\내 드라이브\EleaZar\dev-notes\plc-ppt-converter\

과거에 겪었던 문제, 설계 이유, 트러블슈팅 이력이 필요한 작업을 할 때는
먼저 위 경로를 확인하세요.

그 폴더를 채우는 것은 `dev-plc/knowledge-base` 저장소입니다. 기본 브랜치에 .md 가
푸시되면 `.github/workflows/sync-obsidian.yml` 이 그리로 미러링하고, 데스크톱이
그 저장소를 위 경로로 `fetch` + `reset --hard` 해서 받아갑니다. **볼트에서 직접
고친 것은 다음 동기화 때 덮어쓰입니다** — 손으로 쓰는 메모는 `dev-notes` 바깥에
두세요. (예전에는 Actions 가 Google Drive API 로 직접 올렸는데, auth/drive 가
제한된 스코프라 OAuth 앱을 프로덕션으로 게시할 수 없고 테스트 상태의 리프레시
토큰이 7일마다 만료돼서 걷어냈습니다.)