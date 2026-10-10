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

## 디자인 상수 (convert_pptx.py 상단, 완성샘플 측정값)
- BG_TOP/BG_BOTTOM = #091442 → #191629  (세로 그라데이션 배경, 그림으로 깖)
- HEADER_BG  = #3C4059  (제목 글자 폭에 맞춘 박스)
- GOLD       = #FFCD4A  (원본 <a:highlight> 글자만)
- BODY_PT_MAX/MIN = 32/28  (본문 크기 — 28pt로도 안 들어갈 때만 슬라이드 분리)
- 제목: 국립박물관문화재단클래식 Bold 38pt

## Before PPTX 파싱 주의
- 에베소서/레위기: Shape '제목 1'에 시리즈명+강의제목 합쳐짐
- 마태복음: Shape '제목 1'=시리즈명, Shape '부제목 2'=강의제목+목사명 (분리)
- 제목이 비어 있으면 앞 슬라이드 제목 반복, 제목 shape 자체가 없으면 첫 단락이 제목
- 쉼표 등으로 끝나는 문단은 다음 문단과 한 불렛으로 병합 (_CONTINUATION_ENDINGS)
- 슬라이드 종류: text / text_image(이미지 열 + 글) / image(이미지만, 원본 배치 유지)
- 이미지 전용 파일(이미 변환된 After/1차 작업물)은 ValueError로 거부

## 텔레그램 봇 / NAS
- build_after()는 (슬라이드 수, slide_map)을 반환 — slide_map[i] = 원본 i+1번의 변환 슬라이드 번호들
- preview.py가 LibreOffice로 원본·변환 슬라이드를 JPG로 렌더 → 봇이 짝지어 전송
- 코드만 바뀌면 update_nas.ps1, Dockerfile(apt 패키지)이 바뀌면 deploy_nas.ps1 재실행

## 개발 기록 (옵시디언 연동)

이 저장소의 과거 개발 기록, 트러블슈팅 노트, 설계 결정 사항은 아래 경로에
정리되어 있습니다:

G:\내 드라이브\EleaZar\dev-notes\plc-ppt-converter\

과거에 겪었던 문제, 설계 이유, 트러블슈팅 이력이 필요한 작업을 할 때는
먼저 위 경로를 확인하세요.