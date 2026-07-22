# 39.1 업체별 잔액 그룹핑 + 다운로드 안정화 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 39.1 상세의 잔액 소계를 (업체=회사명 본지점합산, 계정) 단위로 재편해 업체별 연속 블록으로 출력하고, Cloud Run 인스턴스 재활용으로 죽는 다운로드 버튼을 클라이언트 blob 프리페치로 고치고, testrun 5개 파일로 2026 2Q 명세서를 생성한다.

**Architecture:** `detail_group_key`가 잔액 그룹의 유일한 정의다. 키를 (정규회사명, 계정)으로 줄이고 출력 순서를 회사 최초 등장 → 회사 내 계정 최초 등장으로 정렬한다. 다운로드는 실행 완료 직후 프런트가 결과 파일(총 ~200KB)을 fetch→blob으로 확보해 서버 세션 소멸과 무관하게 저장 가능하게 한다.

**Tech Stack:** Python 3.9, pandas, openpyxl, FastAPI, pytest, vanilla JS.

## Global Constraints

- **YTD 불변식 (AGENTS.md):** 1Q=1~3월, 2Q=1~6월, 3Q=1~9월, 4Q=1~12월. 2Q 상세는 반드시 1~6월 포함.
- `39.1`/`39.2`/기간/집계 코드를 만지면 해당 테스트 모듈 + 전체 스위트를 실행해야 함.
- `testrun/`, `sample_input*`, 생성된 명세서는 절대 커밋 금지.
- 거래처명·적요·금액 등 회계 데이터 값을 로그/터미널/테스트 출력에 찍지 않는다 (구조·건수만).
- 사용자 확정 요구: **업체별 연속 블록 + 업체마다 계정별 잔액 소계. 업체 단위 채권·채무 합산(순액/총액) 행은 만들지 않는다.**
- 작업 브랜치: `fix/partner-grouped-balance-and-download` (in-place 브랜치. `.venv`/대용량 데이터가 저장소 경로에 상대적이므로 worktree 사용 금지).
- 테스트 실행 인터프리터: `/Users/jaekyu/Documents/code_p/hg_affiliate/hg-affiliate-app/.venv/bin/python` (venv의 console-script는 깨져 있으니 반드시 `python -m pytest` 형태 사용).
- 모든 pytest 실행의 cwd는 `hg-affiliate-app/`.

---

## File Structure

- Modify `hg-affiliate-app/app/domain/statement_detail.py`: 그룹 키 축소 + 회사 블록 정렬.
- Modify `hg-affiliate-app/app/domain/statement_template.py`: 그룹 키 타입 힌트만 갱신.
- Modify `hg-affiliate-app/tests/test_statement_detail.py`: 새 그룹핑/정렬 회귀 + 기존 기대치 갱신.
- Modify `hg-affiliate-app/app/main.py`: CORS `expose_headers` 추가.
- Modify `hg-affiliate-app/app/static/app.js`: 결과 파일 blob 프리페치 다운로드.
- Modify `hg-affiliate-app/tests/test_api.py`: CORS 노출 헤더 회귀.
- Modify `hg-affiliate-app/run.sh`, `hg-affiliate-app/dev.sh`: venv 인터프리터 직접 지정(활성화 스크립트 의존 제거).

### Task 0: 브랜치 생성과 기존 미커밋 작업 베이스라인 커밋

현재 워킹트리에는 이전 세션의 전기이월(39.1 opening rows) 작업이 미커밋 상태로 있다 (전체 스위트 green 확인됨: 144 passed, 11 skipped).

- [ ] **Step 1: 브랜치 생성**

```bash
cd /Users/jaekyu/Documents/code_p/hg_affiliate
git checkout -b fix/partner-grouped-balance-and-download
```

- [ ] **Step 2: 베이스라인 커밋 (소스/테스트만, 데이터 제외)**

```bash
git add hg-affiliate-app/app/domain/statement.py hg-affiliate-app/app/domain/statement_detail.py hg-affiliate-app/app/domain/statement_template.py hg-affiliate-app/app/pipeline.py hg-affiliate-app/tests/test_statement_detail.py hg-affiliate-app/tests/test_statement_template_period.py
git commit -m "feat: carry prior-year opening balances into 39.1 detail"
git status --short   # testrun/, sample_input* 이 staged 되지 않았는지 확인
```

### Task 1: 39.1 잔액 그룹을 (업체, 계정)으로 재편하고 업체 블록으로 정렬

**Files:**
- Modify: `hg-affiliate-app/app/domain/statement_detail.py:185-192` (`detail_group_key`), `:477-493` (`build_statement_detail` 마지막 그룹 방출부)
- Modify: `hg-affiliate-app/app/domain/statement_template.py:237` (타입 힌트)
- Test: `hg-affiliate-app/tests/test_statement_detail.py`

**Interfaces:**
- Produces: `detail_group_key(row: StatementDetailRow) -> tuple[str, str]` — (정규회사명, 공백제거·casefold 계정 키). 거래처코드는 더 이상 그룹 식별에 쓰지 않는다.
- `build_statement_detail(...)` 반환 순서 계약: 같은 회사(canonical)의 모든 행이 하나의 연속 블록. 회사 순서 = 최초 등장(전기이월 행 포함) 순. 회사 안에서 계정 그룹 순서 = 최초 등장 순. 각 (회사, 계정) 그룹의 마지막 행에만 `balance`(분류된 bucket이 있을 때) 부여 — 기존과 동일한 방향 규칙(`uses_credit_balance`).
- Consumes(불변): `_write_statement_detail`은 `detail_group_key` 변화 지점마다 그룹 시작행을 갱신하므로 연속 블록 계약만 지키면 수식 범위는 자동으로 (회사, 계정) 단위가 된다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_statement_detail.py`에 추가:

```python
def test_detail_orders_rows_into_contiguous_company_blocks():
    ledger = _ledger(
        [
            ("2026-01-03", "상품매출", "A sale"),
            ("2026-02-01", "상품매출", "B sale"),
            ("2026-03-01", "미지급금", "A payable"),
            ("2026-04-01", "상품매출", "A sale 2"),
        ]
    )
    ledger.loc[1, "거래처명"] = "특관자B"
    ledger.loc[1, "거래처코드"] = "V002"

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A", "특관자B": "특관자B"},
        canonical={"특관자A", "특관자B"},
        period=Period(2026, 2),
    )

    assert [str(row.canonical_name) for row in rows] == [
        "특관자A",
        "특관자A",
        "특관자A",
        "특관자B",
    ]
    assert [row.description for row in rows] == [
        "A sale",
        "A sale 2",
        "A payable",
        "B sale",
    ]
    assert [row.balance for row in rows] == [None, 200.0, 100.0, 100.0]


def test_detail_merges_balance_across_partner_codes_of_one_company():
    ledger = _ledger(
        [
            ("2026-01-03", "외상매출금", "head office"),
            ("2026-02-01", "외상매출금", "branch"),
        ]
    )
    ledger.loc[0, ["계정코드", "차변", "대변"]] = ["1080000", "100", "0"]
    ledger.loc[1, ["계정코드", "거래처코드", "거래처명", "차변", "대변"]] = [
        "1080000",
        "V002",
        "특관자A지점",
        "50",
        "0",
    ]

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A", "특관자A지점": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert len(rows) == 2
    assert [row.balance for row in rows] == [None, 150.0]
```

또한 기존 `test_detail_merges_opening_aliases_only_within_same_partner_code_and_account`(line 479)를 새 의미로 개명·확장한다 — 그룹 키에서 거래처코드가 빠졌으므로 전기이월 별칭 병합은 (회사, 계정) 안에서 코드가 달라도 일어난다:

```python
def test_detail_merges_opening_aliases_within_company_and_account():
    ledger = _ledger([("2026-06-30", "외상매출금", "collection")])
    ledger.loc[0, ["계정코드", "거래처코드", "거래처명", "차변", "대변"]] = [
        "1080000",
        "V001",
        "특관자A",
        "0",
        "150",
    ]
    opening = pd.DataFrame(
        [
            {
                "계정과목": "외상매출금",
                "거래처": "특관자A",
                "거래처코드": "V001",
                "금액": "100",
            },
            {
                "계정과목": "외상매출금",
                "거래처": "특관자㈜",
                "거래처코드": "V009",
                "금액": "50",
            },
        ]
    )

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A", "특관자㈜": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
        prev_balance=opening,
    )

    assert len(rows) == 2
    assert rows[0].date == "[전기이월]"
    assert rows[0].debit == 150.0
    assert rows[1].balance == 0.0
```

- [ ] **Step 2: RED 확인**

```bash
cd /Users/jaekyu/Documents/code_p/hg_affiliate/hg-affiliate-app
.venv/bin/python -m pytest tests/test_statement_detail.py -q
```

Expected: 새 테스트 2개 FAIL (`특관자A` 블록이 흩어지고 코드별 잔액이 2개), 개명 테스트는 병합 결과가 달라 FAIL 가능. 기존 테스트 결과를 기록해 둘 것.

- [ ] **Step 3: 최소 구현**

`app/domain/statement_detail.py`의 `detail_group_key`를 다음으로 교체:

```python
def detail_group_key(row: StatementDetailRow) -> tuple[str, str]:
    """Return the contiguous 39.1 balance group identity: (company, account)."""
    return (
        _clean_text(row.canonical_name),
        _account_key(row.account_name),
    )
```

`build_statement_detail`의 마지막 방출부(`rows: list[StatementDetailRow] = []` 이후 전체)를 다음으로 교체:

```python
    group_order = {key: index for index, key in enumerate(grouped)}
    company_order: dict[str, int] = {}
    for key in grouped:
        company_order.setdefault(key[0], group_order[key])

    rows: list[StatementDetailRow] = []
    for key in sorted(
        grouped, key=lambda item: (company_order[item[0]], group_order[item])
    ):
        group = grouped[key]
        last = group[-1]
        # The template's summary and 39.2 detail blocks intentionally reflect only
        # classified buckets.  Leaving P blank for an unclassified account preserves
        # that account-reflection boundary and keeps the 39.1 self-check auditable.
        if last.bucket is not None:
            total_debit = sum(row.debit for row in group)
            total_credit = sum(row.credit for row in group)
            balance = (
                total_credit - total_debit
                if uses_credit_balance(last)
                else total_debit - total_credit
            )
            group[-1] = replace(last, balance=balance)
        rows.extend(group)
    return rows
```

그룹 dict 타입 힌트 2곳(`grouped`, `_opening_rows`의 `by_group`)을 `dict[tuple[str, str], list[StatementDetailRow]]`로 갱신.

`app/domain/statement_template.py:237`의 `previous_group_key: tuple[str, str, str] | None`을 `tuple[str, str] | None`로 갱신.

- [ ] **Step 4: GREEN + 인접 스위트**

```bash
.venv/bin/python -m pytest tests/test_statement_detail.py tests/test_statement_template_period.py tests/test_pipeline_period.py -q
```

Expected: PASS. 기존 테스트 중 (회사, 계정, 코드) 3-tuple 그룹핑이나 흩어진 순서를 전제한 기대치가 깨지면, **새 계약(업체 연속 블록 + 계정별 잔액)에 맞게 기대치만 갱신**한다. 잔액 방향 규칙·bucket 분류 기대치는 바꾸지 않는다.

- [ ] **Step 5: 전체 스위트**

```bash
.venv/bin/python -m pytest -q
```

Expected: all pass (sample 스킵 제외), 신규 실패 0.

- [ ] **Step 6: 커밋**

```bash
git add hg-affiliate-app/app/domain/statement_detail.py hg-affiliate-app/app/domain/statement_template.py hg-affiliate-app/tests/test_statement_detail.py hg-affiliate-app/tests/test_statement_template_period.py
git commit -m "fix: group 39.1 balances per company and account"
```

(템플릿 테스트 파일은 실제 수정된 경우에만 add.)

### Task 2: 다운로드 버튼을 인스턴스 수명과 무관하게 만들기 (blob 프리페치)

**근본 원인(조사 완료):** Cloud Run(max-instances=1, scale-to-zero)에서 실행 완료 후 유휴 ~15분이 지나면 인스턴스가 종료되어 인스턴스 로컬 디스크의 세션·출력 파일이 사라진다. 이후 다운로드 클릭은 새 인스턴스에서 404 (2026-07-22 07:26 UTC, Edge/Windows에서 5회 404 로그로 확인). 로컬 curl/배포 curl 재현에서는 warm 상태라 정상.

**수정:** 실행 완료 직후 프런트가 결과 파일 3개(~200KB)를 fetch해 blob URL로 버튼에 연결한다. 서버 인스턴스가 죽어도 브라우저 메모리의 blob으로 저장 가능. 파일명 보존을 위해 CORS로 `Content-Disposition`을 노출한다.

**Files:**
- Modify: `hg-affiliate-app/app/main.py:30-41` (CORSMiddleware)
- Modify: `hg-affiliate-app/app/static/app.js` (`renderResult`, `run`)
- Test: `hg-affiliate-app/tests/test_api.py`

**Interfaces:**
- Produces: `/download/*` 응답의 `Content-Disposition`이 브라우저 JS에서 읽힘 (`expose_headers`).
- app.js: `prefetchDownloads(downloads) -> Promise<[{name, href, filename, prefetched}]>`; `renderResult`는 async.

- [ ] **Step 1: 실패하는 API 테스트 작성**

`tests/test_api.py`에 추가 (기존 파일의 TestClient 사용 패턴을 그대로 따를 것):

```python
def test_cors_exposes_content_disposition_for_downloads():
    from starlette.testclient import TestClient

    from app.main import app

    client = TestClient(app)
    response = client.get(
        "/api/config", headers={"Origin": "https://hg-affiliate.web.app"}
    )
    exposed = response.headers.get("access-control-expose-headers", "")
    assert "content-disposition" in exposed.lower()
```

- [ ] **Step 2: RED 확인**

```bash
.venv/bin/python -m pytest tests/test_api.py -q
```

Expected: 새 테스트 FAIL (`access-control-expose-headers` 없음).

- [ ] **Step 3: CORS 노출 헤더 추가**

`app/main.py`의 `add_middleware` 호출에 한 줄 추가:

```python
app.add_middleware(
    CORSMiddleware,
    allow_origins=[...기존 그대로...],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)
```

- [ ] **Step 4: GREEN 확인 후 app.js 프리페치 구현**

`app/static/app.js`에서 `renderResult`를 다음으로 교체하고 헬퍼를 추가:

```js
let objectUrls = [];

function parseDownloadFilename(res, fallback) {
  const cd = res.headers.get("Content-Disposition") || "";
  const star = cd.match(/filename\*=utf-8''([^;]+)/i);
  if (star) {
    try { return decodeURIComponent(star[1]); } catch (e) { /* fallthrough */ }
  }
  const plain = cd.match(/filename="?([^";]+)"?/i);
  return plain ? plain[1] : fallback;
}

// 실행 직후 결과 파일을 blob 으로 미리 받아 둔다.
// Cloud Run 인스턴스가 내려가 세션이 사라져도 이미 받은 blob 은 저장 가능하다.
async function prefetchDownloads(downloads) {
  objectUrls.forEach(u => URL.revokeObjectURL(u));
  objectUrls = [];
  return Promise.all((downloads || []).map(async d => {
    try {
      const res = await fetch(apiUrl(d.url));
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const href = URL.createObjectURL(blob);
      objectUrls.push(href);
      return {
        name: d.name,
        href,
        filename: parseDownloadFilename(res, `${d.name}.xlsx`),
        prefetched: true,
      };
    } catch (e) {
      return { name: d.name, href: apiUrl(d.url), filename: `${d.name}.xlsx`, prefetched: false };
    }
  }));
}

async function renderResult(data) {
  const result = document.getElementById("result");
  const box = document.getElementById("result-box");
  result.classList.remove("hidden");
  if (!data.ok && (!data.downloads || !data.downloads.length)) {
    box.className = "rounded-lg border border-red-200 bg-red-50 p-4";
    box.innerHTML = `<p class="text-sm font-medium text-red-700">${escapeHtml(data.message || "실패")}</p>`;
    return;
  }
  box.className = "rounded-lg border border-emerald-200 bg-emerald-50 p-4";
  const links = await prefetchDownloads(data.downloads);
  const buttons = links.map(l =>
    `<a href="${l.href}" download="${escapeHtml(l.filename)}" class="inline-flex items-center gap-2 rounded-md bg-white border px-3 py-2 text-sm font-medium text-indigo-700 hover:bg-indigo-50 mr-2 mb-2">
       ⬇ ${escapeHtml(l.name)}
     </a>`).join("");
  const stale = links.some(l => !l.prefetched)
    ? `<p class="text-xs text-amber-600 mt-2">일부 파일을 미리 받아두지 못했습니다. 버튼이 동작하지 않으면 다시 실행하세요.</p>`
    : "";
  box.innerHTML = `
    <p class="text-sm font-medium ${data.ok ? "text-emerald-700" : "text-amber-700"}">${escapeHtml(data.message)}</p>
    <p class="text-xs text-slate-500 mt-1">검토/오류 건수: <b>${data.error_count ?? 0}</b> (상세는 오류목록 파일 참고)</p>
    <div class="mt-3">${buttons}</div>${stale}`;
}
```

`run()`의 두 `renderResult(...)` 호출을 `await renderResult(...)`로 갱신.

- [ ] **Step 5: 전체 스위트**

```bash
.venv/bin/python -m pytest -q
```

Expected: all pass.

- [ ] **Step 6: 커밋**

```bash
git add hg-affiliate-app/app/main.py hg-affiliate-app/app/static/app.js hg-affiliate-app/tests/test_api.py
git commit -m "fix: prefetch result downloads so buttons survive instance recycling"
```

### Task 3: 로컬 실행 환경 복구 (깨진 venv + 실행 스크립트 견고화)

**배경(조사 완료):** 프로젝트 폴더가 `projectPark` → `code_p/hg_affiliate`로 이동되면서 `.venv`의 console-script shebang과 `activate`의 `VIRTUAL_ENV`가 옛 경로를 가리킨다. 그 결과 `run.sh`/`dev.sh`가 시스템 uvicorn(jinja2 없음)으로 폴백해 로컬 서버가 시작조차 안 된다. 또한 로컬 포트 8000은 무관한 앱(telegram-auto-read)이 점유 중이다.

- [ ] **Step 1: venv 재생성**

```bash
cd /Users/jaekyu/Documents/code_p/hg_affiliate/hg-affiliate-app
python3 -m venv --clear .venv
.venv/bin/python -m pip install --upgrade pip -q
.venv/bin/python -m pip install -r requirements.txt -q
.venv/bin/python -m pytest -q   # 재생성 후 전체 스위트 재확인
head -1 .venv/bin/uvicorn       # shebang 이 현재 경로인지 확인
```

Expected: 스위트 all pass, shebang이 `/Users/jaekyu/Documents/code_p/hg_affiliate/...`.

- [ ] **Step 2: run.sh / dev.sh 가 activate 에 의존하지 않게 수정**

두 스크립트에서 `source .venv/bin/activate` + `exec uvicorn ...` 패턴을 venv 인터프리터 직접 실행으로 교체. `run.sh`:

```bash
#!/usr/bin/env bash
# 운영/일반 실행 (리로드 없음 → 안정적인 단일 프로세스).
# 회계 담당자가 도구를 사용할 때는 이 스크립트를 쓰세요.
set -euo pipefail
cd "$(dirname "$0")"

PORT="${PORT:-8000}"

PY="python3"
if [ -x ".venv/bin/python" ]; then
  PY=".venv/bin/python"
fi

exec "$PY" -m uvicorn app.main:app --host 0.0.0.0 --port "$PORT"
```

`dev.sh`도 동일하게 `exec "$PY" -m uvicorn app.main:app --port "$PORT" --reload ...` (기존 reload 옵션 유지).

- [ ] **Step 3: 실제 기동 확인 후 커밋**

```bash
PORT=8901 ./run.sh &  # 몇 초 뒤
curl -s http://127.0.0.1:8901/api/config   # {"has_key":...} 확인 후 kill %1
git add hg-affiliate-app/run.sh hg-affiliate-app/dev.sh
git commit -m "fix: run scripts use venv interpreter directly"
```

### Task 4: testrun 5개 파일로 2026 2Q 명세서 생성 + 실사용 검증

- [ ] **Step 1: 로컬 서버로 전체 플로우 실행 (다운로드 수정 검증 겸용)**

로컬 서버 기동(`PORT=8901 ./run.sh`) 후 API로: 세션 생성 → testrun 1)~5) 업로드(자동 슬롯 배정 확인) → `/api/run {year:2026, quarter:2}` → ok=true 확인 → `/download/{sid}/당기_특관자_명세서`를 받아 `testrun/당기_특관자_명세서_2026_2Q.xlsx`로 저장. 회계 값은 출력하지 않는다.

- [ ] **Step 2: 구조 검증 스크립트 (수치 미출력)**

scratchpad에 검증 스크립트를 작성해 실행:

1. `39.1`의 각 정규회사(M열)가 **정확히 하나의 연속 블록**인지 (블록 수 == 회사 수).
2. P열 수식이 각 (회사, 계정) 그룹 마지막 행에만 존재하고, SUM 범위가 그룹 시작~끝과 일치하는지.
3. 날짜 있는 행의 연·월이 모두 2026년 1~6월이고 min month == 1 인지.
4. `verify_output()` ok.
5. 오류목록/지배구조 파일 생성 확인.

Expected: 모두 통과. 개수(회사 수, 그룹 수, 행 수)만 출력.

- [ ] **Step 3: 브라우저 실사용 검증 (다운로드 버튼)**

browse 스킬(헤드리스 브라우저)로 `http://127.0.0.1:8901` 접속 → 5개 파일 업로드 → 2026/2분기 실행 → 결과 버튼 렌더 확인 → **서버 프로세스 kill** → 다운로드 버튼 클릭 → blob 덕분에 파일이 정상 저장되는지 확인. (이것이 Cloud Run 인스턴스 소멸 시나리오의 로컬 재현이다.)

- [ ] **Step 4: LibreOffice 렌더 확인**

```bash
soffice --headless --convert-to pdf --outdir <scratchpad> "testrun/당기_특관자_명세서_2026_2Q.xlsx"
```

39.1 페이지에서 기간 라벨·행 클리핑·수식 오류가 없는지 눈으로 확인.

- [ ] **Step 5: git 오염 확인**

```bash
git status --short   # testrun/ 파일이 추적되지 않는지 확인
```

### Task 5: 머지 + 배포 + 최종 검증

- [ ] **Step 1: 전체 스위트 최종 실행 후 main 머지**

```bash
.venv/bin/python -m pytest -q
git checkout main && git merge fix/partner-grouped-balance-and-download
```

- [ ] **Step 2: Cloud Run + Hosting 배포**

Docker 데몬 확인 후 `hg-affiliate-app/deploy.sh` 실행 (기존 배포 스크립트 그대로). 실패 시 원인만 보고하고 로컬 검증 결과로 마무리.

- [ ] **Step 3: 배포 검증**

배포된 `/api/config` 응답 + `/api/config`에 Origin 헤더로 `access-control-expose-headers` 확인. (실데이터 재실행은 사용자 몫으로 남긴다.)

- [ ] **Step 4: superpowers:verification-before-completion 후 결과 보고**

## Self-Review 결과

- 요구 1(업체별 잔액) → Task 1. 사용자 확정: 업체 합산 행 없음, 계정별 소계 유지 — 반영됨.
- 요구 2(다운로드 디버깅) → Task 2 (근본 원인: 인스턴스 재활용, 증거: Cloud Run 404 로그) + Task 3(로컬 환경).
- 요구 3(2Q 생성) → Task 4.
- 타입 일관성: `detail_group_key` 2-tuple로 통일 (statement_template 힌트 포함).
