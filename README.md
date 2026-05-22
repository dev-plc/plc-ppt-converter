# PL Church Sermon PPT Converter

PL교회 설교 자료 PPT를 자동으로 PL 디자인 규칙에 맞게 변환하고, Canva로 업로드합니다.

## 사용법

```bash
# 기본 변환
python convert_pptx.py "에베소서 제4강(before).pptx"

# 변환 + PNG 내보내기 (PowerPoint 설치 필요)
python convert_pptx.py input.pptx output.pptx --png

# 변환 + Canva 자동 업로드
python convert_pptx.py input.pptx --canva

# Canva 토큰 직접 지정
python convert_pptx.py input.pptx --canva --canva-token YOUR_TOKEN
```

- 출력 PPTX는 입력 파일과 같은 폴더에 `_converted.pptx`로 저장됩니다.
- `--canva` 실행 후 출력되는 편집 링크로 Canva에서 바로 열 수 있습니다.

## Canva 연동 설정

### 1단계 — Canva 앱 생성

1. [https://www.canva.com/developers/apps](https://www.canva.com/developers/apps) 접속
2. **Create an app** 클릭 → 앱 이름 입력 (예: `PL PPT Importer`)
3. **Scopes** 에서 다음 항목 체크:
   - `design:content:write`
   - `design:content:read`
4. **Redirect URL** 에 `http://localhost:8080/callback` 추가
5. **Client ID** 와 **Client Secret** 복사

### 2단계 — Access Token 발급

OAuth 2.0 Authorization Code Flow를 사용합니다.

브라우저에서 아래 URL 접속 (`CLIENT_ID` 교체):

```
https://www.canva.com/api/oauth/authorize?
  response_type=code&
  client_id=CLIENT_ID&
  scope=design:content:write%20design:content:read&
  redirect_uri=http://localhost:8080/callback
```

리디렉션된 URL에서 `code=` 값 복사 후 토큰 교환:

```bash
curl -X POST https://api.canva.com/rest/v1/oauth/token \
  -u "CLIENT_ID:CLIENT_SECRET" \
  -d "grant_type=authorization_code" \
  -d "code=AUTHORIZATION_CODE" \
  -d "redirect_uri=http://localhost:8080/callback"
```

응답의 `access_token` 값을 환경변수에 설정:

```bat
# Windows PowerShell
$env:CANVA_ACCESS_TOKEN = "your_access_token_here"

# Windows 영구 설정
setx CANVA_ACCESS_TOKEN "your_access_token_here"
```

### 3단계 — 단독 업로드 (선택)

이미 변환된 PPTX를 Canva에 올리려면:

```bash
python canva_upload.py output.pptx "에베소서 제4강"
```

## 의존성 설치

```bash
pip install python-pptx pillow comtypes requests
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
├── convert_pptx.py     # 메인 변환 스크립트 (--canva 플래그 포함)
├── canva_upload.py     # Canva 업로드 모듈 (단독 실행 가능)
├── 표지.png             # 타이틀 슬라이드 배경 이미지
├── samples/            # Before 샘플 파일
│   ├── 에베소서 제4강(before).pptx
│   ├── 레위기 제12강(before).pptx
│   └── 마태복음 제41강(before).pptx
└── analyze_pptx.py     # PPTX 구조 분석 도구 (개발용)
```
