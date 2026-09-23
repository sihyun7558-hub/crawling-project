# MUSINSA Ranking Crawler

MUSINSA 랭킹 페이지의 상품 정보를 자동으로 수집하고, 이미지와 데이터를 저장하며 상품명 키워드를 분석할 수 있도록 개발한 Python 데스크톱 애플리케이션입니다.

처음에는 Playwright를 이용해 랭킹 정보를 수집하는 간단한 프로그램으로 시작했습니다. 이후 실제 사용 과정에서 필요한 기능을 단계적으로 추가하여 이미지 저장, GUI 갤러리, CSV/Excel 출력, 수집 범위 설정, 중단 기능, 키워드 분석까지 확장했습니다.

> 이 프로젝트는 개인 학습 및 포트폴리오 목적으로 제작했으며 MUSINSA의 공식 프로젝트가 아닙니다. 사이트 구조 변경에 따라 선택자(selector)가 동작하지 않을 수 있습니다.

## 주요 기능

- Playwright 기반 동적 페이지 크롤링
- 상품 순위 / 브랜드 / 상품명 / 가격 수집
- TOP 10 / 20 / 50 / 100 빠른 범위 설정 및 직접 범위 입력
- 상품 이미지 자동 다운로드
- 수집 이미지 GUI 갤러리 표시
- CSV 및 Excel(.xlsx) 저장
- 상품명 기반 키워드 빈도 TOP 10 분석
- 진행률 표시 및 크롤링 중단
- 결과 저장 폴더 선택
- 별도 Thread에서 크롤링을 실행하여 GUI 응답성 유지

## Tech Stack

- **Language:** Python
- **Browser Automation:** Playwright
- **GUI:** CustomTkinter / Tkinter
- **HTTP:** Requests
- **Image:** Pillow
- **Excel:** OpenPyXL
- **Text Analysis:** `re`, `collections.Counter`
- **Concurrency:** `threading`

## 처리 흐름

```text
MUSINSA Ranking URL
        ↓
Playwright로 페이지 접근 및 스크롤
        ↓
순위 / 브랜드 / 상품명 / 가격 / 이미지 URL 추출
        ↓
상품 이미지 다운로드 및 GUI Gallery 표시
        ↓
CSV / Excel 저장
        ↓
상품명 토큰 추출 및 불용어 제거
        ↓
키워드 빈도 TOP 10 표시
```

## 프로젝트 구조

```text
crawling-project/
├── crawler.py
├── requirements.txt
├── README.md
├── .gitignore
├── screenshots/
└── archive/
```

`screenshots/`에는 실행 화면을 추가할 수 있으며, `archive/`에는 실제로 보관 중인 이전 버전(v1.3, v3.0 등)이 있을 경우 원본 스냅샷을 보관하는 용도로 사용할 수 있습니다.

## 설치

Python 가상환경 사용을 권장합니다.

```bash
python -m venv .venv
```

macOS / Linux:

```bash
source .venv/bin/activate
```

Windows:

```bash
.venv\\Scripts\\activate
```

필요한 패키지를 설치합니다.

```bash
pip install -r requirements.txt
playwright install chromium
```

## 실행

```bash
python crawler.py
```

프로그램 실행 후 랭킹 페이지 URL과 원하는 순위 범위를 입력하고 **추출 시작** 버튼을 누릅니다.

## 저장 데이터

수집된 데이터는 다음 컬럼으로 CSV와 Excel 파일에 저장됩니다.

| 컬럼 | 내용 |
| --- | --- |
| 순위 | 랭킹 순위 |
| 브랜드 | 상품 브랜드 |
| 상품명 | 상품명 |
| 가격 | 상품 가격 |
| 파일명 | 다운로드한 이미지 파일명 |

이미지와 CSV/Excel 결과는 사용자가 지정한 저장 폴더에 생성됩니다.

## 키워드 분석

수집된 상품명에서 한글과 영문 단어를 추출하고 불용어를 제거한 뒤 `Counter`를 사용해 빈도를 계산합니다. 가장 많이 등장한 키워드 10개를 GUI에서 간단한 막대 형태로 확인할 수 있습니다.

## 개발 과정

```text
v1.x  기본 Playwright 크롤링 + GUI + 진행률
  ↓
v3.x  이미지 다운로드 + Gallery + CSV / Excel 저장
  ↓
v5.0  순위 범위 설정 + 중단 기능 + 키워드 분석 + UI 개선
```

이 버전 흐름은 보관 중인 실제 소스 스냅샷을 기준으로 관리하는 것을 원칙으로 합니다. Git 이력이 없는 과거 버전은 과거 날짜의 커밋으로 재작성하지 않고 archive 또는 release 형태로 보관합니다.

## 구현하면서 배운 점

단순히 웹페이지의 데이터를 가져오는 것에서 끝나지 않고 **수집 → 저장 → 확인 → 분석**까지 이어지는 흐름을 하나의 데스크톱 프로그램으로 구성했습니다. 동적 페이지에서 필요한 데이터를 추출하기 위해 Playwright를 사용했고, 크롤링 작업을 별도 Thread에서 실행하여 GUI가 멈추지 않도록 구성했습니다.

또한 수집한 결과를 CSV/Excel과 이미지로 저장하고 상품명 데이터를 다시 분석하면서, 웹 데이터 수집 이후의 데이터 처리 과정까지 직접 구현했습니다.

## Known Limitations

- MUSINSA의 HTML 구조나 CSS selector가 변경되면 크롤링 로직 수정이 필요할 수 있습니다.
- 현재 상품 페이지 열기 기능은 상품 식별자 대신 상품명을 사용하고 있어 개선이 필요합니다.
- 일부 예외 처리가 포괄적으로 작성되어 있어 향후 logging과 세분화된 오류 처리가 필요합니다.
- GUI와 크롤링/저장 로직이 하나의 파일에 있어 모듈 분리가 필요합니다.

## Future Improvements

- GUI / crawler / exporter / analyzer 모듈 분리
- 상품 ID 및 상품 URL 직접 수집
- selector 관리 구조 개선
- `logging` 기반 오류 기록
- 테스트 코드 추가
- 데이터 분석 및 시각화 기능 확장

## Disclaimer

웹사이트의 이용약관, robots 정책 및 요청 빈도를 확인하고 준수해야 합니다. 수집한 데이터와 이미지를 재배포하거나 상업적으로 사용하는 경우에는 해당 콘텐츠의 권리와 이용 조건을 별도로 확인해야 합니다.
