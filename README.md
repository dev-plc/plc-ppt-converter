# PL Church Sermon PPT Converter

PL교회 설교 자료 PPT를 자동으로 PL 디자인 규칙에 맞게 변환합니다.

## 사용법

```bash
python convert_pptx.py "에베소서 제4강(before).pptx"
python convert_pptx.py input.pptx output.pptx --png
```

- 출력 PPTX는 입력 파일과 같은 폴더에 `_converted.pptx`로 저장됩니다.
- `--png` 옵션: 슬라이드별 PNG(1920×1080)도 함께 내보냅니다. (PowerPoint 설치 필요)

## 의존성 설치

```bash
pip install python-pptx pillow comtypes
```

## 폰트 설치 (필수)

`C:\Users\[사용자]\AppData\Local\Microsoft\Windows\Fonts\` 에 아래 파일 설치:

- `국립박물관문화재단클래식B.otf`
- `국립박물관문화재단클래식M.otf`
- `국립박물관문화재단클래식L.otf`

Noto Sans KR: [Google Fonts](https://fonts.google.com/noto/specimen/Noto+Sans+KR) 에서 설치

## 디자인 규칙

| 역할 | 값 |
|------|-----|
| 배경 | `#1B2045` |
| 헤더 박스 | `#232A56` |
| 강조 (골드) | `#F5C518` |
| 기본 텍스트 | `#FFFFFF` |
| 제목 폰트 | 국립박물관문화재단클래식 Bold |
| 본문 폰트 | Noto Sans KR Bold |

## 파일 구조

```
plc ppt/
├── convert_pptx.py     # 메인 변환 스크립트
├── 표지.png             # 타이틀 슬라이드 배경 이미지
├── samples/            # Before 샘플 파일
│   ├── 에베소서 제4강(before).pptx
│   ├── 레위기 제12강(before).pptx
│   └── 마태복음 제41강(before).pptx
└── analyze_pptx.py     # PPTX 구조 분석 도구 (개발용)
```
