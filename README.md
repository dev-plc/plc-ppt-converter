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

## 디자인 규칙 (완성샘플 측정값)

| 역할 | 값 |
|------|-----|
| 배경 | 세로 그라데이션 `#091442` → `#191629` |
| 제목 박스 | 글자 폭에 맞춘 박스 `#3C4059`, 국립박물관문화재단클래식 Bold 38pt |
| 본문 | Noto Sans KR Bold 32pt (넘치면 28pt까지 축소), `#EEEEEE`, "-" 불렛 |
| 강조 (골드) | `#FFCD4A` — 원본에 형광펜이 칠해진 글자만 |
| 표지 목사님 이름 | `#FEEFC5` |

## 변환 규칙

- **슬라이드 분리 최소화**: 원본 1장 → 변환 1장이 기본. 32pt → 28pt로 줄여도 안 들어갈 때만 나눈다.
- **문단 이어 붙이기**: 쉼표(,) 등으로 끝나 문장이 이어지는 문단은 다음 문단과 한 불렛으로 합친다.
- **빈 제목**: 원본 제목이 비어 있으면 앞 슬라이드 제목을 이어서 쓴다.
- **이미지만 있는 슬라이드** (지도·도표 등): 원본 배치 그대로 옮긴다.
- **글 + 이미지 슬라이드**: 원본에서 이미지가 있던 쪽(좌/우)에 이미지 열, 반대쪽에 글.

## 텔레그램 봇

Before PPTX를 보내면 원본 슬라이드 이미지 → 변환 슬라이드 이미지 순서로 짝지어 보낸 뒤
변환된 PPTX를 보냅니다. 원본 1장이 2장으로 나뉘면 변환 이미지를 앨범으로 묶고
"1 → 2페이지"를 붙입니다. 비교 이미지는 LibreOffice로 만들며(`preview.py`),
NAS Docker 이미지에 포함되어 있습니다.

## 파일 구조

```
plc ppt/
├── convert_pptx.py     # 메인 변환 스크립트
├── telegram_bot.py     # 텔레그램 봇 (NAS Docker에서 실행)
├── preview.py          # 원본/변환 슬라이드 미리보기 이미지 (LibreOffice)
├── 표지.png             # 타이틀 슬라이드 배경 이미지
├── samples/            # Before 샘플 파일
│   ├── 에베소서 제4강(before).pptx
│   ├── 레위기 제12강(before).pptx
│   └── 마태복음 제41강(before).pptx
└── analyze_pptx.py     # PPTX 구조 분석 도구 (개발용)
```
