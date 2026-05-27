# Vibe Coding Prompt: 특관자명세서 자동 생성 Web App

> 이 문서는 Cursor, Claude Code, Replit, Bolt.new, Lovable 등 vibe coding 도구에
> 입력하기 위한 **단일 프롬프트**이다. 아래 내용 전체를 한 번에 붙여넣고 시작할 것.

---

## 0. 프로젝트 한 줄 요약

> "분기별 회계 데이터(전기이월 잔액명세서 + 당기 계정별원장)와 특관자 매핑 테이블을 업로드하면
> 한국채택국제회계기준(K-IFRS) 주석용 **특수관계자 거래내역 명세서** Excel 파일을
> 자동으로 생성·다운로드하는 SaaS-style 사내 웹앱을 만든다.
> 프론트/백엔드 모두 Firebase에 배포하고, 소스코드는 GitHub에 둔다."

---

## 1. 비즈니스 컨텍스트

### 1-1. 무엇을 만드는가
회사(예: 에이치엘비글로벌(주))의 재무팀이 분기별 결산 시 작성해야 하는
**특수관계자(특관자) 주석용 명세서**(K-IFRS 1024호) 작성 자동화 도구.

### 1-2. 현재 수기 작업
재무팀은 매분기마다 다음 4개 Excel 파일을 가공해
하나의 명세서 파일을 만든다 (현재 약 1~2일 소요):

1. `전기이월소스_계정별잔액명세서` — 전 분기말 BS 계정별 거래처별 잔액 (50+ 시트)
2. `당기거래내역소스_계정별원장` — 당분기 모든 거래 (단일 시트, 18,000+ 행)
3. `채권채무잔액검증용` — 검증 보조
4. `특관자상호정리` — 거래처명 → 회사명(본지점합산) 매핑 (피벗 형태)

### 1-3. 산출물
한 개의 xlsx 파일에 다음 3개 시트:

- `39.1 (2)` — 거래처별·계정과목별 상세 거래내역 (전기이월 + 당분기 거래)
- `특관자` — 회사명별 요약(38.1 거래내역, 38.2 채권채무, 38.3 대손충당금, 38.4 자금거래)
- `39.2` — 회사명 × 거래유형 cross-tab 요약

---

## 2. 핵심 비즈니스 로직 (반드시 구현)

### 2-1. 입력 → 출력 파이프라인

```
[Step 1] 4개 파일 업로드
   ├── 전기이월소스 (.xlsx, 다시트)
   ├── 당기거래내역소스 (.xlsx, Sheet1)
   ├── 검증용소스 (.xlsx, optional)
   └── 특관자상호정리 (.xlsx, Sheet1)

[Step 2] 거래처 → 회사명 매핑 빌드
   - 특관자상호정리.xlsx 1열을 "부모(회사명) → 자식(거래처명)" 트리로 파싱
   - 부모 행 다음에 들여쓰기 없이 나오는 행이 자식 (alias)
   - "[제외]" 표기된 자식은 매핑 제외
   - 결과: {거래처명: 회사명} dict + {회사명: 분류} (종속/관계/기타특수관계자/기타)

[Step 3] 전기이월 entries 생성
   - 전기이월소스의 각 BS 계정 시트에서 거래처별 잔액 row 추출
   - 거래처명을 회사명으로 매핑 (매핑 안되면 SKIP — 특관자 아님)
   - GROSS 금액 사용 (대손충당금 차감 전)
   - 18200 장기외상매출금 시트는 EXCLUDE (2021년 구 (주)넥스트사이언스 데이터)

[Step 4] 당기 거래내역 추가
   - 당기거래내역소스의 모든 row 중 거래처가 특관자 매핑된 것만 포함
   - 8190000 + 거래처코드 88570 + 차변<0 조건은 EXCLUDE
     (전대료 reclassification 분개로, GROSS 임차료만 표시해야 함)

[Step 5] 동일 거래처코드 내 중복 거래처명 통합
   - 예: 89687 "에이치엘비(주)"와 "에이치엘비㈜" (㈜는 단일 특수기호)
   - [전기이월] level에서 (회사명, 계정코드, [전기이월]) 기준 합산

[Step 6] 미지급금/선수금 분리, 미수금/미수수익 분리
   - 동일 거래처/회사라도 계정과목은 절대 합산 금지
   - 미지급금(2530000), 선수금(2590000), 미수금(1200000), 미수수익(1160000) 각각 별도 행
   - 유사 분리: 매입채무/매입선급금, 매출채권/선수수익 등

[Step 7] 39.1 (2) 시트 작성
   - 상단(R1~R10): 요약 합계 (매출등/매입등/채권채무/총잔액)
   - R15 헤더: 매입매출 | 자금거래 | 채권채무 | 수익.비용 | 구분계정과목 |
              계정코드 | 계정과목 | 날짜 | 적요 | 거래처코드 | 거래처명 |
              회사명(본지점합산) | 차변 | 대변 | 잔액
   - R16~: detail rows

[Step 8] 특관자 시트 작성
   - 38.1 (당기 + 전기): 회사명별 매출/기타수익/매입/기타비용 (분류 4개 그룹별 합계)
   - 38.2 (당기말 + 전기말): 매출채권/대여금/기타채권/기타채무/투자전환사채/전환사채
   - 38.3: 대손충당금 (대손상각비 = 당기 변동분)
   - 38.4: 자금거래 (자금대여/상환/차입/상환/지분취득/처분/자본불입)
   - R107 (에이치엘비(주) 지분취득)은 외부 워크북 link 또는 수동 입력 필드로 제공

[Step 9] 39.2 시트 작성
   - "no, 구분(대분류), 구분(소분류=회사명), 거래처명, 매출, 기타수익, 매입, 기타비용"
   - 상세 컬럼: 매출, 기타수익, 이자수익, 매입, 기타비용, 자산매각, 자산취득
   - 두 번째 테이블: 회사명별 매출채권/대여금/기타채권/차입금/매입채무/기타채무/지분취득

[Step 10] 동일 워크북에 모두 저장 후 사용자에게 다운로드 제공
```

### 2-2. v1 → v2 매핑 변경 처리
v2 변경 사례를 반드시 처리:
- `에이치밸류에셋(주)`가 별도 회사로 분리 (v1에서는 에이치엘비제넥스 하위 [제외] 항목)
- 즉, 매핑 파일을 새로 업로드하면 모든 로직이 그 매핑을 따라야 함 (하드코딩 금지)

### 2-3. 검증 규칙
파일 업로드 직후 다음을 검사하고 경고 표시:
- 매핑 파일에 등장하지 않은 거래처가 당기/전기 데이터에 존재 → 회사명=`(미매핑)` 으로 표시하고 사용자에게 listing
- 차변·대변 양쪽 모두 비어있는 row → skip
- 음수 차변 + 거래처코드 88570 + 8190000 조합 → 자동 제외 (Step 4 규칙)
- 18200 장기외상매출금 → 자동 제외 (사용자 토글 가능)

---

## 3. 기술 스택 (확정)

| Layer | 선택 | 이유 |
|---|---|---|
| Frontend | **React 18 + TypeScript + Vite + Tailwind CSS + shadcn/ui** | Firebase Hosting 호환 최강, vibe coding 도구 친화 |
| State | **TanStack Query + Zustand** | 서버 상태 + 클라이언트 상태 분리 |
| File Upload UI | **react-dropzone** | drag & drop UX |
| Backend | **Cloud Functions for Firebase, 2nd gen, Python 3.11** | `openpyxl`/`pandas` 그대로 사용 가능 (Excel 처리에 압도적 유리) |
| Excel 처리 | **openpyxl** (formula 보존) + **pandas** (집계) | |
| Storage | **Firebase Storage** | 업로드된 입력 파일 + 생성된 출력 파일 |
| DB | **Firestore** | 사용자별 작업 이력, 매핑 테이블, 분기별 메타데이터 |
| Auth | **Firebase Auth (Google + Email)** | 사내 사용자 한정 (allow-list domain) |
| Hosting | **Firebase Hosting** (frontend) + **Cloud Run** (functions auto-deploy) | |
| CI/CD | **GitHub Actions → Firebase deploy** | main 브랜치 push 시 자동 배포 |
| 패키지 관리 | npm (frontend), pip (functions) | |

### 3-1. 폴더 구조

```
hg-affiliate-app/
├── .github/workflows/
│   └── deploy.yml                  # GitHub Actions: build + firebase deploy
├── firebase.json
├── .firebaserc
├── firestore.rules
├── storage.rules
├── README.md
├── apps/
│   ├── web/                        # Frontend (React + Vite)
│   │   ├── src/
│   │   │   ├── components/
│   │   │   │   ├── FileUpload.tsx
│   │   │   │   ├── MappingEditor.tsx
│   │   │   │   ├── JobProgress.tsx
│   │   │   │   ├── DownloadButton.tsx
│   │   │   │   └── ui/             # shadcn components
│   │   │   ├── pages/
│   │   │   │   ├── HomePage.tsx
│   │   │   │   ├── NewJobPage.tsx
│   │   │   │   ├── JobDetailPage.tsx
│   │   │   │   ├── MappingPage.tsx
│   │   │   │   └── LoginPage.tsx
│   │   │   ├── lib/
│   │   │   │   ├── firebase.ts
│   │   │   │   ├── api.ts
│   │   │   │   └── types.ts
│   │   │   ├── App.tsx
│   │   │   └── main.tsx
│   │   ├── index.html
│   │   ├── vite.config.ts
│   │   ├── tailwind.config.js
│   │   ├── tsconfig.json
│   │   └── package.json
│   └── functions/                  # Backend (Python Cloud Functions)
│       ├── main.py                 # Function entrypoints
│       ├── core/
│       │   ├── __init__.py
│       │   ├── mapping.py          # 특관자상호정리 parser
│       │   ├── prev_balance.py     # 전기이월 추출
│       │   ├── current_ledger.py   # 당기거래 처리
│       │   ├── merge.py            # 통합/필터링/sum
│       │   ├── sheet_391.py        # 39.1 (2) writer
│       │   ├── sheet_tukgwanja.py  # 특관자 writer
│       │   ├── sheet_392.py        # 39.2 writer
│       │   └── rules.py            # exclude / business rules
│       ├── tests/
│       │   ├── fixtures/           # 작은 샘플 xlsx
│       │   ├── test_mapping.py
│       │   ├── test_rules.py
│       │   └── test_end_to_end.py
│       ├── requirements.txt        # openpyxl, pandas, firebase-admin, functions-framework
│       └── pyproject.toml
└── docs/
    ├── BUSINESS_LOGIC.md           # Section 2 내용 옮김
    ├── DEPLOYMENT.md
    └── CHANGELOG.md
```

---

## 4. 사용자 흐름 (UX)

### 4-1. 첫 화면 (Home)
- 로그인 안되어 있으면 → Google 로그인 버튼
- 로그인 후 → 최근 작업(분기) 카드 리스트 + "새 분기 생성" CTA
- 사이드바: Jobs / Mapping / Settings

### 4-2. 새 작업 생성 (`/new`)
1. **분기 선택**: 연도/분기 dropdown (예: 2026 1Q)
2. **파일 업로드 (4개 zone)**:
   - 전기이월소스 (필수)
   - 당기거래내역소스 (필수)
   - 검증용소스 (선택)
   - 특관자상호정리 (필수 — Mapping 페이지에 저장된 버전이 있으면 기본값 자동 선택 가능)
3. **검증 결과 미리보기**:
   - 매핑 안된 거래처 listing (펼쳐서 보기, 일괄 매핑 버튼)
   - 자동 제외될 row 수 (장기외상매출금, 음수 8190000 등)
   - 토글 옵션: "장기외상매출금(18200) 제외", "8190000 음수 제외", "동일 거래처코드 통합"
4. **수동 입력 필드**:
   - 특관자 R107 외부 link 값 (예: 지분취득 13,392,708,800)
   - 기타 수동 보정 필드 (필요시)
5. **[생성] 버튼** 클릭 → 백엔드 호출 → 진행률 표시 (parsing → 매핑 → detail → summary → write)
6. **완료** → 다운로드 + 미리보기(3 시트 탭) + 검증 통계 카드

### 4-3. 매핑 관리 (`/mapping`)
- 현재 활성 매핑(v2 등) 표시: 회사명(부모) → 거래처명(자식) 트리 view
- 추가/수정/삭제 (드래그앤드롭으로 alias 이동)
- 변경 사항 저장 시 새 버전 생성 (v1, v2, v3...) — 과거 분기 결과 재현 가능
- xlsx 업로드로 일괄 import / xlsx 내보내기

### 4-4. 작업 상세 (`/jobs/:id`)
- 입력 파일 정보, 사용한 매핑 버전, 생성 시각
- 생성된 결과 xlsx 다운로드
- 시트별 미리보기 (페이지네이션)
- 같은 분기 이전 작업과의 diff (셀 단위 비교)
- (선택) 정답 파일 업로드 → 자동 비교 → 일치율 % 표시

---

## 5. UI/UX 가이드

- 한국어 인터페이스. 회계용어는 원어 유지.
- 색상: 차분한 비즈니스 톤 (slate-50 배경, primary는 blue-600).
- 폰트: `Pretendard` 또는 시스템 기본.
- 모든 숫자는 천단위 구분(,)과 단위 표기(원).
- 음수는 빨강 + 괄호 표기: `(33,810,000)`.
- 테이블은 가상 스크롤 (TanStack Table + virtualizer).
- 모든 destructive action은 confirm dialog.

---

## 6. 백엔드 API 설계 (Cloud Functions)

각 함수는 HTTPS callable, Firebase Auth ID token 검증.

| Function | Method | 입력 | 출력 |
|---|---|---|---|
| `mapping_parse` | POST | xlsx URL | parsed mapping JSON |
| `mapping_save` | POST | mapping JSON, version label | mapping doc id |
| `mapping_list` | GET | - | versions list |
| `job_validate` | POST | 4 file URLs + mapping id | 검증 결과 (unmapped 거래처, 제외 예정 row 수 등) |
| `job_run` | POST | 4 file URLs + mapping id + options + manual overrides | job id (async) |
| `job_status` | GET | job id | progress, status, result URL |
| `job_diff` | POST | job id + answer xlsx URL | cell-by-cell diff |

Long-running `job_run`은 Cloud Tasks 또는 Pub/Sub로 비동기 처리. Firestore에 job document 만들고 functions가 업데이트, 클라이언트는 Firestore listener로 실시간 진행률 표시.

---

## 7. Firestore 스키마

```
users/{uid}
  email, displayName, role

mappings/{mappingId}
  version (e.g., "v2")
  label, createdAt, createdBy
  entries: [{ parent: "에이치엘비(주)", aliases: ["에이치엘비(주)", "에이치엘비㈜"], category: "기타특수관계자" }]
  excluded: ["에이치밸류에셋 주식회사"]  # [제외] 표기 처리

jobs/{jobId}
  uid, createdAt, quarter (e.g., "2026-1Q"), status
  inputs: { prevBalance, currentLedger, validation, mappingId }
  options: { excludeLongAR: true, excludeNeg8190000: true, mergeDuplicateVendors: true }
  overrides: { tukgwanjaR107: 13392708800, ... }
  progress: { step: "writing_392", pct: 0.85 }
  result: { outputUrl, stats: { rowsWritten, unmappedVendors, ... } }
  error: null | "..."
```

---

## 8. 보안 규칙

- **Firebase Auth**: 회사 도메인만 허용 (Cloud Function에서 email 도메인 검사 — 예: `@hlbglobal.com`).
- **Storage 규칙**: `/users/{uid}/uploads/**` 본인만 R/W, `/users/{uid}/outputs/**` 본인만 R, write는 functions(service account)만.
- **Firestore 규칙**: jobs는 자기 것만, mappings는 인증된 사용자 read all + write는 role=admin 만.
- 비밀키는 Firebase Secret Manager 사용 (env file commit 금지).

---

## 9. 단계별 구현 순서 (vibe coding이 이 순서를 그대로 따를 것)

### Phase 0 — 부트스트랩 (30분)
1. `npm create vite@latest apps/web -- --template react-ts`
2. `cd apps/web && npm install` + Tailwind/shadcn 초기화
3. `firebase init` (hosting, functions[python], storage, firestore, emulators)
4. Firebase project 생성, GitHub repo 생성, push
5. GitHub Actions workflow 작성: main push → emulator test → deploy

### Phase 1 — 핵심 백엔드 (Python core) (2일)
1. `apps/functions/core/mapping.py`: 특관자상호정리.xlsx 파싱 (트리 구조 인식)
2. `core/prev_balance.py`: 전기이월 추출 (모든 시트 순회, 거래처/금액 컬럼 자동 탐지)
3. `core/current_ledger.py`: 당기 거래 추출
4. `core/rules.py`: 모든 비즈니스 규칙 (Section 2 그대로 — exclude, merge, split)
5. `core/sheet_391.py` / `sheet_tukgwanja.py` / `sheet_392.py`: 시트 작성기
6. 본 프로젝트의 4Q 작업완료 파일과 1Q 정답 파일을 fixture로 테스트:
   - 1Q 입력 → 결과가 1Q 정답과 100% 일치해야 PASS

### Phase 2 — Functions endpoints + Storage 연동 (1일)
1. HTTPS callable wrapper 추가
2. Storage 업로드 URL signed 발급, 파일 read/write
3. Firestore job document 생성/업데이트
4. Pub/Sub trigger로 long-running 분리

### Phase 3 — Frontend MVP (2일)
1. Auth flow (Google + 도메인 allow-list)
2. 새 작업 페이지 (4 file dropzone + validation preview)
3. 진행률 + 다운로드
4. 작업 이력 페이지

### Phase 4 — 매핑 편집 UI (1일)
1. Tree view (회사명 부모 + alias 자식)
2. 드래그앤드롭으로 alias 이동
3. 버전 관리 (저장 시 새 버전 생성)
4. xlsx import/export

### Phase 5 — 폴리싱 (1일)
1. 결과 미리보기 (3 sheet tab)
2. 검증/diff 기능
3. 에러 핸들링, loading state
4. README + DEPLOYMENT.md + screenshots

---

## 10. 테스트 케이스 (반드시 통과)

본 프로젝트(`[재무]_특관자_주석_생성/`)의 실 데이터를 사용:

1. **재현 테스트**: `2025.4Q_작업완료/` 입력 4개 파일 + 구버전 매핑(v1) → 출력이 기존 `4)특관자명세서.xlsx`와 일치
2. **v2 적용 테스트**: `2026.1Q_작업중/` 입력 + v2 매핑 → 출력이 `5)2026_1Q_특관자명세서_정답.xlsx`와 **100% 셀-by-셀 일치**
3. **장기외상매출금 제외 토글**: 토글 OFF 시 4)2026_1Q_특관자명세서.xlsx의 결과 재현 (장기외상매출금 포함)
4. **에이치밸류에셋 분리**: v2 매핑 적용 시 39.2에 별도 행으로 등장
5. **음수 8190000 처리**: 88570 + 8190000 + 차변<0 entries 12건 자동 제외 확인
6. **거래처명 통합**: 89687 "에이치엘비(주)"와 "에이치엘비㈜" 전기이월 합산 확인

`apps/functions/tests/test_end_to_end.py`에 위 케이스 모두 작성. CI에서 emulator로 실행.

---

## 11. 배포

### 11-1. Firebase 설정
- `firebase.json`:
  ```json
  {
    "hosting": {
      "public": "apps/web/dist",
      "rewrites": [
        { "source": "/api/**", "function": "api" },
        { "source": "**", "destination": "/index.html" }
      ]
    },
    "functions": [
      { "source": "apps/functions", "runtime": "python311", "region": "asia-northeast3" }
    ],
    "firestore": { "rules": "firestore.rules" },
    "storage": { "rules": "storage.rules" },
    "emulators": {
      "auth": { "port": 9099 },
      "functions": { "port": 5001 },
      "firestore": { "port": 8080 },
      "storage": { "port": 9199 },
      "hosting": { "port": 5000 }
    }
  }
  ```
- 리전: `asia-northeast3` (서울).

### 11-2. GitHub Actions (`.github/workflows/deploy.yml`)
- PR: emulator로 lint + test
- main push: Firebase 토큰으로 hosting + functions deploy

### 11-3. 비용 고려
- Functions 2nd gen + Storage + Firestore 무료 한도 내에서 운영 가능 (사내 사용량 가정)
- 큰 xlsx 처리 시 메모리 1GiB 이상 필요할 수 있음 → function memory 옵션 조정

---

## 12. 참고 자료 (vibe coding 도구에 함께 제공)

- 본 프로젝트 폴더 `[재무]_특관자_주석_생성/` 전체 (실 데이터)
  - `2025.4Q_작업완료/4)특관자명세서.xlsx` ← 정상 산출물 참고
  - `2026.1Q_작업중/5)2026_1Q_특관자명세서_정답.xlsx` ← 신규 정답
  - `2026.1Q_작업중/특관자상호정리_v2.xlsx` ← 신규 매핑
- `오류분석및수정보고서.md` ← 9개 버그와 비즈니스 룰 상세
- K-IFRS 1024 (특수관계자 공시) 표준 — 사용자 별도 제공

---

## 13. 첫 출력 요구사항 (vibe coding에 명시)

이 프롬프트를 받은 즉시 다음을 순서대로 출력할 것:

1. 최종 폴더 구조 트리 (Section 3-1 확정안)
2. `package.json` (frontend) + `requirements.txt` (functions) 초안
3. `firebase.json`, `.firebaserc` (project id placeholder), `firestore.rules`, `storage.rules`
4. `apps/functions/main.py` + `core/` 모듈 6개 (mapping/prev_balance/current_ledger/rules/sheet_391/sheet_tukgwanja/sheet_392) — Section 2 비즈니스 로직 그대로 구현
5. `apps/web/src/App.tsx`, `pages/NewJobPage.tsx`, `components/FileUpload.tsx`, `MappingEditor.tsx` — Section 4 UX 그대로
6. `apps/functions/tests/test_end_to_end.py` — Section 10 테스트 케이스 1번 + 2번
7. `.github/workflows/deploy.yml`
8. `README.md` — 빠른 시작(로컬 emulator 실행 ~3분), 배포(~10분), 첫 분기 생성(~5분)

각 파일은 **그대로 복사해 붙여넣어 실행 가능한 완전한 코드**여야 한다.
부분 코드, `...`, `TODO` 금지. 작성 후 본 프로젝트의 1Q 정답 파일로 검증하는 셀-by-셀 일치 테스트가 PASS 됨을 확인할 때까지 반복 수정한다.

---

## 14. Done Definition

- [ ] `firebase emulators:start`로 로컬 실행 가능
- [ ] 1Q 정답 파일과 100.000% 셀-by-셀 일치 (테스트 PASS)
- [ ] Google 로그인으로 접속 → 4 파일 업로드 → 결과 다운로드 (E2E)
- [ ] 매핑 편집 → 새 버전 저장 → 그 버전으로 작업 생성 (E2E)
- [ ] GitHub main push → 자동 배포 → 운영 URL에서 동일 시나리오 동작
- [ ] README.md만 보고 새 개발자가 30분 내 로컬 실행 성공
