# 당기 기간 선택 및 데이터 추출 설계

- 작성일: 2026-06-12
- 대상 앱: `hg-affiliate-app` (HLB글로벌 특관자 명세서 생성기)
- 상태: 설계 합의 완료 → 구현 계획 단계로 이행 예정

## 1. 배경 및 용어

특관자 명세서는 **분기 누적(YTD)** 기준으로 작성된다.

- **당기** 기간 정의(누적):
  - 1Q = 해당년도 1월~3월
  - 2Q = 해당년도 1월~6월
  - 3Q = 해당년도 1월~9월
  - 4Q = 해당년도 1월~12월
- **전기** = 직전 결산년도(전년도 연간 결산 = 전년도 4분기). 당기가 어느 분기든
  전기는 항상 직전년도 4분기로 동일. (기존 spec/가드와 일치)

샘플(`_sample_input/`)이 정의를 그대로 증명한다:

| 파일 | 슬롯 | 기간 |
|---|---|---|
| `1)_전기이월소스_25.4Q…` | `prev_balance` | 25.4Q |
| `2)_당기거래내역소스_26.1Q…` | `current_ledger` | 26.1Q |
| `3)_…채권채무잔액검증용_26.1Q…` | `current_balance` | 26.1Q |
| `4)_특관자상호정리…` | `name_pivot` | — |
| `5)_전기_특관자명세서…` | `prev_statement` | — |

## 2. 목표 / 비목표

**목표**
1. 당기 기간(년도 + 분기)을 사용자가 선택하는 옵션 제공.
2. 선택된 당기 기간(해당년도 1월~분기말월)에 해당하는 데이터를 입력 소스에서
   추출하여 '당기 소스'로 사용.
3. 날짜 파싱 신뢰도를 높이기 위해 결정론 파서 + AI(ChatGPT) 보조 파서 사용.
4. 선택 년도로 전기 가드(`period_check`)를 정밀화.

**비목표**
- 잔액명세서(당기말/전기말)에 대한 기간 필터링 (아래 3.3 근거 참조).
- 출력물 제목에 기간 라벨 자동 삽입(추후 별도 검토).
- 전기 기간을 분기별로 회전(정의상 항상 전년도 4Q이므로 불필요).

## 3. 설계

### 3.1 기간 모델
신규 모듈 `app/domain/period_extract.py`.

- `Period(year: int, quarter: int)`; `end_month = quarter * 3`.
- 채택 조건: `parsed.year == year AND parsed.month <= end_month`.
  - 시작은 항상 1월이라 일(day) 정밀도 불필요 → 월 경계 비교.
- `parse_year_month(value) -> tuple[int, int] | None` (결정론):
  - `datetime`/`date` 객체 → (year, month)
  - 숫자(엑셀 serial, 합리적 범위) → `datetime(1899,12,30)+days`
  - 문자열: `2026-03-15`, `2026.3.15`, `2026/03/15`, `20260315`,
    `2026년 3월`, `2026-03` 등에서 년·월 추출(정규식)
  - 그 외 → `None`

### 3.2 당기 선택 UI
- `index.html` 상단(업로드 위)에 **년도 입력**(숫자, 2000~2099) + **분기 라디오**(1Q/2Q/3Q/4Q). 둘 다 필수.
- `app.js`: 실행 시 `year`, `quarter`를 `/api/run` 바디에 포함.
- `main.py /api/run`: `year`(2000~2099), `quarter`(1~4) 검증. 누락/범위 밖이면 400.

### 3.3 당기 데이터 추출 — **당기 원장에만 적용**
**근거**: 잔액명세서는 54/51개 '시트=계정' 스냅샷이며, 일부 시트의 `예금일자/만기일자`는
거래 발생일이 아니라 계정 속성 날짜다. 이를 기간 필터에 쓰면 채권채무 잔액이 손상된다.
잔액명세서는 "해당 분기말 시점 스냅샷"이므로 분기별로 올바른 파일을 업로드해 그대로 쓴다.

추출 흐름(원장 프레임에 대해):
1. 날짜 컬럼 해석: `columns.resolve_column(df, "date")`.
   - **버그 수정**: `COLUMN_ALIASES["date"]`에 **"날짜"** 추가(샘플 원장 헤더가 "날짜").
     기타 흔한 별칭(`회계일자/발생일자/거래일자/전기일자`)도 보강.
   - 날짜 컬럼 미식별 → **다운로드 차단(오류목록 기록)**.
2. 행 단위 파싱:
   - 날짜 셀이 빈값/합계·소계 라벨 → 비거래행으로 제외(실패 아님).
   - 값이 있는 셀: (a) `parse_year_month` 결정론 파싱.
   - (a) 실패한 '비어있지 않은' 셀들 → (b) **ChatGPT 배치 파서**로 `YYYY-MM` 재시도.
   - (a)·(b) 모두 실패한 셀이 **1건이라도** 남으면 → **다운로드 차단(엄격)**.
   - OPENAI 키 없음/AI 호출 실패로 (b) 불가한데 미파싱 셀이 남아도 → 차단.
3. 채택: 파싱된 (year, month)가 `Period` 조건을 만족하는 행만 당기로 남김.
4. 필터링된 원장 프레임을 기존 집계(`aggregate_ledger`, `_verify_reconstruction`)에 전달.
   다운스트림은 변경 없음(이미 좁혀진 당기 부분집합을 받음).

### 3.4 AI 보조 날짜 파서
`normalize.py`의 OpenAI 패턴을 따른다(클라이언트 생성/재시도/JSON 파싱 동일 구조).
- `period_extract.py` 내 `_openai_parse_dates(values, settings)`:
  - 배치(예: 60개), `response_format=json_object`, `temperature=0`, 재시도/백오프.
  - 프롬프트: "각 문자열에서 거래 발생 연·월을 추출, `YYYY-MM` 또는 파싱불가면 `null`로
    JSON 반환." 회계 날짜 표기(한글/구분자 혼용) 처리 지시.
  - 반환 JSON을 `{원본문자열: (year, month)}`로 변환. `null`/형식오류는 미파싱으로 둠.
- API 키 값은 화면·로그 비노출(기존 원칙 유지). 추출 수치(행 내용)도 로그 금지.

### 3.5 전기 가드 연동
- `period_check.evaluate_prior_period(prev_text, current_text, selected_year=None)`로 확장:
  - `selected_year` 주어지면 기대 전기연도 = `selected_year - 1`, 기대 분기 = 4.
  - 전기 소스 표식 연도가 `selected_year-1`이 아니거나, 1~3분기로 보이면 경고.
  - `selected_year` 없으면 기존 휴리스틱(min-year 비교) 유지(fallback).
- `pipeline.run_pipeline`이 `Period`를 받아 가드에 `selected_year` 전달.

### 3.6 데이터 흐름 (run_pipeline)
```
/api/run(year, quarter) → run_pipeline(slot_paths, settings, output_dir,
                                       slot_filenames, period)
  0) 전기 가드(period.year 연동)
  1) 프레임 로드
  1.5) 당기 원장 = extract_current_period(ledger, period)   ← 신규
       · 날짜컬럼 미식별/미파싱 잔존 → 차단 반환(PipelineOutcome.ok=False)
  2~8) 기존과 동일 (필터링된 원장으로 집계)
```

## 4. 오류 처리 / 엣지 케이스
- 원장 미업로드: 기존대로 38.1/38.4 건너뜀(경고). 기간 선택은 여전히 필수(전기 가드·라벨용).
- 날짜 컬럼 없음 / 유효 날짜 0건 / 미파싱 잔존: 다운로드 차단 + 오류목록 사유 기록.
- AI 파싱으로 구제된 행 수는 정보성 메모로 남김(수치 내용은 비노출).
- 합계/footer/빈 날짜 행: 정상 제외(차단 아님).

## 5. 테스트 (TDD, 샘플 사용)
- 위치: `hg-affiliate-app/tests/`. 샘플 파일은 민감데이터 → 없으면 `skip`.
- **버그 가드**: 실제 원장의 "날짜" 컬럼이 `resolve_column(df,"date")`로 해석되는지(현재 실패).
- 실제 원장(26.1Q) → `Period(2026,1)` 필터 시 전 행 채택, 행수 일치.
- 실제 원장 → `Period(2026,4)`도 전 행 채택(month≤12).
- 합성 경계: `2025-12`·`2026-04` 행 주입 후 `Period(2026,1)`에서 제외 확인.
- `parse_year_month`: ISO/점/슬래시/`YYYYMMDD`/`YYYY년 M월`/엑셀serial/`datetime`.
- 미파싱 셀 잔존 → 추출이 '차단' 신호 반환.
- AI 파서: 호출은 모킹(실네트워크 없이), JSON 반환 → (year,month) 변환 검증.
- 전기 가드: `selected_year=2026` → 전기 소스 "25.4Q"는 무경고, "26.1Q"는 경고.

## 6. 보안 / 운영
- `_sample_input/`을 `.gitignore`에 추가(실거래 데이터, 커밋 금지).
- 입력/출력 내용·산출 수치 화면·로그 비노출 원칙 유지.

## 7. 변경 파일 요약
- `app/domain/columns.py` — `date` 별칭에 "날짜" 등 추가.
- `app/domain/period_extract.py` — 신규(Period, 파서, 추출, AI 보조).
- `app/domain/period_check.py` — `selected_year` 연동.
- `app/pipeline.py` — `period` 인자, 원장 추출 단계, 차단 처리.
- `app/main.py` — `/api/run`의 `year/quarter` 검증·전달.
- `app/templates/index.html`, `app/static/app.js` — 당기 선택 UI.
- `.gitignore`(root) — `_sample_input/` 추가.
- `tests/test_period_extract.py`(+가드 테스트) — 신규.
