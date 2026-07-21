# HLB글로벌(주) 특수관계자 명세서 생성기 (v2)

회계 담당자가 엑셀 5종을 올리면 서버가 **당기 특관자 명세서 · 지배구조 · 오류목록** 3개 파일을 생성하여 **다운로드만** 제공하는 도구입니다.
입력/출력의 **내용(셀 값·표·미리보기)은 화면·로그에 절대 표시하지 않습니다.**

## 핵심 동작

- 입력: 모든 엑셀 형식 지원 (`.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`)
- 업로드 즉시 선행 검증: 확장자·매직바이트·워크북 오픈·시트/헤더·손상/암호화/0바이트 차단
- 사업자명(거래처명) 정규화: "특관자 상호 정리"의 들여쓰기 시드 → ChatGPT API → 휴리스틱 폴백
- 38.1~38.5 도메인 집계 후 명세서 생성, 검증용 잔액명세서로 자체 검증
- 출력 안전성: openpyxl 재오픈 + LibreOffice 재계산 → 수식오류(`#REF!` 등) **0건** 확인 후에만 다운로드 제공

## 입력 5종 (슬롯)

1. 전기 이월 소스 — 전기말 계정별 잔액 명세서
2. 당기 거래 내역 소스 — 당기 계정별 원장(분개 라인)
3. 당기 채권채무 잔액 검증용 소스 — 당기말 계정별 잔액 명세서 **(필수)**
4. 특관자 상호 정리 — 정규화 기준/시드
5. 전기 특관자 명세서 — 양식/수식 템플릿

> 파일명이 제각각이어도 내용 기반으로 자동 분류하며, 화면에서 슬롯을 수동 교정할 수 있습니다.

## ⚠️ 분기 선택은 연초 누적(YTD)입니다

이 앱의 분기는 독립된 3개월 구간이 아닙니다. 선택 연도의 1월부터 분기말까지를
항상 누적합니다.

- 1분기: 1월~3월
- 2분기: 1월~6월
- 3분기: 1월~9월
- 4분기: 1월~12월

예를 들어 2026년 2분기 명세서의 집계와 `39.1` 상세 거래는 모두
2026년 1월~6월이어야 합니다. 4월~6월만 추출하면 잘못된 결과입니다.

## 설정

```bash
cp secret.env.example secret.env
# secret.env 를 열어 OPENAI_API_KEY 를 입력 (키는 화면/로그에 노출되지 않음)
```

`secret.env.example`:

```
OPENAI_API_KEY=
OPENAI_MODEL=gpt-4.1-mini
```

## 명세서 고정 템플릿 (필수 자산)

당기 특관자 명세서는 사업보고서 게재용 **고정 양식**(`특관자`/`39.1`/`39.2` 시트)을
템플릿으로 사용합니다. 이 파일은 실제 거래 데이터가 담긴 '정답' 파일이라 **저장소에 커밋하지 않습니다**(`.gitignore`).

```bash
# 직전 분기 정답(또는 검증된 빈 양식) xlsx 를 아래 경로에 배치
cp <정답_특관자명세서.xlsx> app/assets/statement_template.xlsx
```

- 동작 원리: `특관자` 시트는 전부 `=SUMIF('39.2'!...)` 수식이므로, 우리는 **`39.2` 시트의
  거래/채권채무 값 컬럼만** 거래처명(정규화 키) 매칭으로 주입합니다. 보고서 수치는 자동 계산됩니다.
- 전기말·38.3~38.5·사채명세 등 **수기/전기 영역은 템플릿 값을 그대로 보존**합니다(분기별 수동 갱신 전제).
- 템플릿이 없으면 단순표 폴백으로 산출됩니다(서식 없음).
- 데이터가 있으나 양식에 행이 없는 법인은 **오류목록('템플릿 미반영')** 에 기록됩니다.

## 실행

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# 일반 실행 (권장, 리로드 없음 → 안정적인 단일 프로세스)
./run.sh                 # 또는: uvicorn app.main:app --host 0.0.0.0 --port 8000
# http://localhost:8000
```

### 개발 중 핫리로드

```bash
./dev.sh                 # app/ 만 감시, .venv/.sessions/__pycache__ 제외
```

> **무한 리로드 주의**: `.venv` 가 프로젝트 폴더 안에 있으면 IDE 인덱싱/파일시스템
> 이벤트로 인해 `--reload` 가 의존성 파일 변화를 감지해 서버가 계속 재시작됩니다.
> `--reload-dir app` 만으로는 macOS 에서 완전히 막히지 않습니다.
> **확실한 해결책은 venv 를 프로젝트 밖에 두는 것**입니다(자발적 리로드 0 확인):
>
> ```bash
> python -m venv ~/.venvs/hg-affiliate
> source ~/.venvs/hg-affiliate/bin/activate
> pip install -r requirements.txt
> uvicorn app.main:app --reload --reload-dir app --port 8000
> ```

> LibreOffice(headless)가 설치되어 있으면 출력 재계산 검증을 수행합니다.
> 미설치 시 openpyxl 재오픈 + 수식 정적 점검으로 대체합니다.
> macOS: `brew install --cask libreoffice`, Ubuntu: `apt-get install libreoffice-calc`

## Docker

```bash
docker build -t hg-affiliate .
docker run -p 8000:8000 -e PORT=8000 --env-file secret.env hg-affiliate
```

## Cloud Run 배포

`secret.env` 와 `app/assets/statement_template.xlsx` 가 로컬에 있어야 합니다(둘 다 git 제외).

```bash
chmod +x deploy.sh
./deploy.sh
```

기본값: 프로젝트 `hg-affiliate`, 리전 `asia-southeast1`, 서비스 `hg-affiliate-runtime`.
환경변수로 `GCP_PROJECT`, `GCP_REGION`, `GCP_SERVICE` 를 덮어쓸 수 있습니다.

배포 후 접속 주소:

- **https://hg-affiliate.web.app** (Firebase Hosting → Cloud Run rewrite)
- Cloud Run 직접 URL (`deploy.sh` 마지막에 출력)

`firebase.json`(저장소 루트)이 `hg-affiliate-runtime` 으로 모든 요청을 프록시합니다.
Hosting만 다시 배포하려면 저장소 루트에서 `firebase deploy --only hosting` 을 실행하세요.

Cloud Run 은 세션 상태를 인스턴스 메모리에 두므로 **`--max-instances 1`** 로 배포합니다.
(인스턴스가 여러 개면 업로드마다 세션을 잃어 `undefined`·검증 실패가 날 수 있음)

## 구조

```
app/
  main.py            FastAPI 라우트
  config.py          설정/슬롯 정의 (secret.env 로드)
  session.py         세션 격리 + 만료 삭제
  excel_io.py        다중 포맷 읽기
  validation.py      입력 선행 검증
  classifier.py      내용 기반 자동 분류
  normalize.py       사업자명 정규화 (시드 + OpenAI + 폴백)
  output_check.py    출력 안전성 검증
  pipeline.py        처리 오케스트레이션
  domain/            38.x 집계 + 출력 3종 생성
  templates/, static/  단일 페이지 UI
```

## 비고

- 38.5(경영진 보상)은 인사·급여 자료가 없으면 표만 유지하고 숫자는 공란입니다.
- 전환사채(투자/발행)는 회차별 보유자 확인이 필요하여 오류목록에 검토 항목으로 남깁니다.
- 도메인 집계는 계정과목/거래처/차변·대변 등 컬럼명을 키워드로 인식합니다. 인식 실패 시 오류목록에 기록됩니다.
