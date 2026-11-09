# 대단원·#태그 AI 자동판정 (GAS) — 구현계획서 v2

작성: 2026-11-07 · 대상: gas-project-problem-ledger (D조 문항장부 바운드 스크립트)
근거: v1 계획서 + `AiTagging.gs` v1 + 레포 `prefix-ds@b2c78f9` 전 파일 + 두 스프레드시트 재실측(2026-11-07) + Claude API 사양(Opus 5.5, API 버전 2026-10-21)
상태: **N1~N9 권장안 확정(2026-11-07)** · 발견 1~7 재검증 완료(§8) · `AiTagging.gs` v2 + `MainMenu.gs` 작성 완료(문법 검사만 통과, 실제 실행 전 — §6 Step 0-b 스모크부터)

---

## 0. 요약 — v1에서 무엇이 바뀌나

v1의 큰 틀(장부 A열 ↔ Data_DS L열 매칭 → 문제·풀이를 AI에 읽혀 L·M 기록, 과목 결정 순서, strict 검증 대응, 원본 시트 17열 고정, `_태그사전`, `_AI판정로그`, 파일럿 먼저)은 **그대로 유지**합니다. 아래 항목을 고칩니다.

### 0-1. 실행 불가 (고치지 않으면 한 건도 처리 못 함)

| # | v1 | 문제 | v2 |
|---|---|---|---|
| C1 | `TEMPERATURE: 0` 전송 | Opus 5.5는 `temperature`·`top_p`·`top_k`를 받지 않음 → **모든 요청이 400** → 전 행 `api_error` | 필드 삭제 (기본 샘플링만 가능) |
| C2 | `max_tokens: 600` | Opus 5.5는 **thinking이 항상 켜져 있고 끌 수 없음**(끄면 400). thinking 토큰도 `max_tokens`에 포함 → 답을 쓰기 전에 **자주** 잘려 `stop_reason: "max_tokens"`, 텍스트 없음 → v1은 `parse_error`로 뭉뚱그림 | `max_tokens: 16000`, `output_config.effort` 명시(N2), `stop_reason` 검사 → `max_tokens` 상태로 구분 기록 |
| C3 | `max_tokens: 0` 예열 호출 | ① `output_config.format`(구조화 출력)과 함께 보내면 400 ② effort·thinking 설정이 본 요청과 다르면 캐시가 어긋남 ③ 배치 작업은 예열 이득이 없음 | 예열 삭제. 동기 경로에서는 과목별 첫 건만 단독 호출하고 나머지는 동시 호출 |
| C4 | 4.5분 tick을 트리거로 반복 | 개인(gmail) 계정은 **트리거 총 실행시간이 하루 90분**. Opus thinking 응답은 건당 수십 초라 하루에 끝나지 않고, CroP 삽입·문서 생성과 같은 90분을 나눠 씀 | **Message Batches API**로 제출하고 결과만 짧게 받아 기록(N1). 단가도 50% 낮음 |
| C5 | `LockService.getScriptLock()`을 tick 내내 보유 | 스크립트 잠금은 **프로젝트 전체가 공유**. `crop-image-insert.gs`의 `cropRunBatch_`는 `tryLock(10초)` 실패 시 재예약 없이 끝남 → AI tick이 돌면 **CroP 삽입이 조용히 멈춤**. 반대 방향은 더 나쁨: v1 tick은 **자기 트리거를 지운 뒤** `tryLock(15초)`하므로 잠금 실패 = AI 실행 자체가 소멸 | 자체 임대(lease, ScriptProperties 만료시각)로 바꾸고 공용 잠금은 임대 획득 순간(수 ms)에만. 임대 중이면 끝내지 않고 **2분 뒤 재예약** |
| C6 | 대상 ~2,000행 가정, `no_source`를 재실행 건너뛰기 목록에 넣음 | 실측(§1): Data_DS에 jangbu_id가 있는 행은 4,749행 이후 약 1,550행뿐이고 **S는 S0001~S0368만** 있음. [써킷] S5001~은 0건. 그런데 `no_source`를 건너뛰기 목록에 넣으면 **변환 후에도 영원히 건너뜀**(v1 §6 "재실행하면 자연히 처리"와 코드가 모순) | `no_source`는 건너뛰기 목록에서 제외(API 호출이 없으니 매번 다시 확인해도 비용 0). 행마다 로그를 남기지 않고 실행 요약에 개수만 |

### 0-2. 코드 버그 (돌기는 하지만 멈추거나 데이터를 망침)

| # | 위치 | 문제 | v2 |
|---|---|---|---|
| B1 | `ai_judgeBatch_` | `fetchAll`이 예외를 던지면 그 청크의 결과가 `null`로 남고 `ai_finalize_`에서 `res.error` 접근 → TypeError → tick 중단 | 결과 기본값을 `{error:'fetch_exception'}`으로, 청크 실패 시 건별 재시도 |
| B2 | `ai_tagTick_` 전체 | tick 안 어디서든 예외가 나면 다음 트리거가 예약되지 않음 → **진행 상태만 남고 조용히 멈춤** | 본문 `try/catch` → 오류를 state·로그에 남기고 재예약(연속 3회 실패 시 중지 + toast) |
| B3 | `ai_tagStop` | 진행 중 tick이 끝나며 state를 다시 저장하고 트리거를 새로 만듦 → **중지가 안 됨** | 저장·재예약 직전에 state가 아직 있는지(같은 runId인지) 다시 읽고, 없으면 그냥 끝냄 |
| B4 | `ai_flushWrites_` | `write_error` 로그의 ID 열이 빈칸 → 건너뛰기 목록이 그 행을 못 잡음. 이미 `ok/review`로 센 행을 `fail`로 또 셈 | ID를 기록하고, 집계는 기록 결과가 확정된 뒤 한 번만 |
| B5 | `ai_userText_`, `ai_loadDataDS_` | `truncated`, `dup`을 계산만 하고 로그에 안 남김 | 로그 "그림·잘림·중복" 열에 기록 |
| B6 | 기록 경로 | 수집 시점과 기록 시점 사이(동기 최대 4.5분, 배치 최대 수 시간)에 사람이 L·M을 적으면 **AI가 덮어씀**. 행 번호만 믿고 씀 | 기록 직전 그 행의 A(ID)·L·M을 다시 읽어 ① ID가 다르면 `id_moved` ② L이 채워졌으면 `write_skipped`로 남기고 쓰지 않음 |
| B7 | `ai_appendLog_` | 행마다 `appendRow` → 느리고 쓰기 할당량을 많이 씀 | 모아서 `setValues` 한 번 |
| B8 | `ai_loadDataDS_` | tick마다 Data_DS 6,300행 × A~L(문제·풀이 본문 포함, 수십 MB)을 읽음. 앞 4,700여 행은 jangbu_id도 없음 | A·L 두 열만 읽어 `jangbu_id → 행` 색인 → 필요한 행의 B·C만 연속 구간으로 읽음 |
| B9 | 재시도 | 429/529의 `retry-after` 헤더 무시 | 헤더 값 우선, 없으면 지수 백오프 |
| B10 | `ai_parseResponse_` | `stop_reason`을 안 봄(`refusal`, `max_tokens`가 `parse_error`로 뭉뚱그려짐). 정규식으로 JSON 추출 | `stop_reason` 분기 + 구조화 출력으로 `text` 블록을 그대로 `JSON.parse` |

### 0-3. 계획서 문구 오류

| # | v1 | 정정 |
|---|---|---|
| E1 | `CONFIG.TARGET_SHEETS`, `CONFIG.USE_PROMPT_CACHE` 등 | 레포에 전역 `CONFIG`가 이미 있음(`구글문서버전복원.gs:13`). Apps Script는 모든 파일이 전역을 공유하므로 같은 이름을 쓰면 충돌. 코드대로 **`AI_CFG`** 로 표기 |
| E2 | §3-6 P열 예시 `(haiku-5-5, c=0.86)` | D5(Opus 단독)·D6(P열 기록 안 함)과 모순 → 삭제 |
| E3 | 출력 예시 `unit_candidates` / 동시 5건 / 10행마다 flush | 코드는 `units_used` / 4건 / 8행. v2는 구조화 출력 스키마(§2)로 통일 |
| E4 | 메뉴 블록을 `// ---- 메뉴 완성 ----` 앞에 | `MainMenu.gs`에 그런 주석이 없음 → **`menu.addToUi();` 바로 앞** |
| E5 | 2,000행 $25, 파일럿 $1~2 | 출력 150토큰만 계산 — thinking 토큰(출력 단가로 과금)을 빠뜨림. §5에서 재추정 |
| E6 | "6300행 텍스트라도 수 초" | 본문 포함 수십 MB라 근거 없음 → B8 |
| E7 | "S#### 표본은 못 봄" | 실측으로 S0001~S0368 확인(§1) |
| E8 | §3-1 `normalizeId_`·`parseRowInput_` 등 기존 헬퍼 재사용 | 코드는 `ai_` 자체 함수만 씀 — 다른 파일 의존이 없어 이쪽이 맞음. 문구만 정정 |
| E9 | §6 "`no_source`는 변환 후 재실행하면 자연히 처리" | 코드와 모순(C6) → v2에서 실제로 그렇게 동작하도록 수정 |

---

## 1. 실측 (2026-11-07 재확인·추가분)

v1 §1의 실측은 아래를 제외하고 그대로 유효합니다.

**Latex 변환 · Data_DS**
- L `jangbu_id`는 **값**(수식 아님). 2~4,748행(`연구실01회조판1_…` 등)은 비어 있고, **4,749행부터 약 1,550행**에만 ID가 있음.
- S 행: **S0001~S0368** (4,749행 `24모의X01_1공통01` = S0001 … S0044~S0046 = `24모의X01_5기하28~30`). S0369 이후와 S5001~은 없음.
- A `id`는 구 명명(`24모의X01_5기하28`)과 새 명명(`26X모의14_1공통22`)이 섞일 수 있음 → 과목코드 정규식 `_(1공통|3확통|4미적|5기하)\d{1,2}$`은 둘 다 잡음(유지).

**D조 문항장부**
- `Data!A2` = `={'[풀세트]'!2:2000;'[써킷]'!2:751;'[문항]'!2:199;'[풀세트]'!2001:3500;'[써킷]'!752:1251}` (v1과 같음), `Data!R2` = `=countif($Q$2:$Q,Q2)`.
- `[풀세트]` 2001행 = S0001: **B(연도) 빈칸**, E `X.01`, F `1 공통`(값), G 1, K `=iferror(VLOOKUP(L…),"")`. S0369 = E `25.X.01` → Data_DS에 아직 없음.
- `[써킷]` 752행 = S5001: E `써킷.X.00`, F `1 공통`(값), G 9, K `=vlookup(L…)`(iferror 없음 → L 빈칸이면 `#N/A`, 정상).
- 선택항목 E2:G16 대단원 15개·(사)과목 4종 — v1과 같음.
- `문제이동(일괄).gs`는 행을 옮기지 않고 같은 행의 **E·G를 바꿈** → 행 번호는 안정적이지만, 결과를 받을 때 A(ID)를 다시 확인하는 것이 안전(B6).

**레포**
- `ai_`/`AI_` 접두어 이름은 레포 어디에도 없음 → 충돌 없음.
- 트리거·잠금을 쓰는 스크립트: `crop-image-insert.gs`(ScriptLock, 트리거 `cropContinueInsert`/`cropContinueVerify`), `문항구글문서 자동생성.gs`(트리거 `processNextBatch_`). 모두 같은 90분/일 트리거 예산을 씀.

**지금 처리 가능한 범위**: `[풀세트]` S0001~S0368 — `L2001:M2368`을 읽어 보니 **전부 빈칸**이므로 368행 모두 대상. 나머지 S 행은 Latex 변환 후 같은 메뉴로 다시 실행하면 자동 포함.

---

## 2. API 호출 사양 (Opus 5.5)

```
POST https://api.anthropic.com/v1/messages            (배치: /v1/messages/batches 의 params 로 같은 본문)
x-api-key: <키>
anthropic-version: 2026-10-21
content-type: application/json

{
  "model": "claude-opus-5-5",
  "max_tokens": 16000,
  "output_config": {
    "effort": "medium",                                  // N2. 기본값도 medium이지만 명시
    "format": { "type": "json_schema", "schema": <과목별 스키마> }
  },
  "system": [
    { "type": "text", "text": "<규칙 + 허용 대단원>" },
    { "type": "text", "text": "<과목별 태그 사전>",
      "cache_control": { "type": "ephemeral", "ttl": "1h" } }
  ],
  "messages": [ { "role": "user", "content": "[문항] S0001 (과목: 1 공통, 번호: 1)\n\n[문제]\n…\n\n[풀이]\n…" } ]
}
```

- **보내지 않는 것**: `temperature`·`top_p`·`top_k`(400), `thinking`(생략 = adaptive, 끌 수 없음), 어시스턴트 prefill(400).
- **과목별 스키마** — 대단원을 `enum`으로 묶어 **허용 목록 밖 값이 원천적으로 나오지 않게** 함(strict 검증 대응이 코드 검사에서 API 보장으로 바뀜). 키워드는 사전이 수백 개라 enum으로 넣지 않고 코드에서 검증(v1 방식 유지).
  ```json
  {"type":"object","additionalProperties":false,
   "required":["unit","units_used","keywords","confidence","reason"],
   "properties":{
     "unit":{"type":"string","enum":["11 지수함수와 로그함수","12 삼각함수","13 수열","21 함수의 극한과 연속","22 미분","23 적분"]},
     "units_used":{"type":"array","items":{"type":"string","enum":[…같은 목록…]}},
     "keywords":{"type":"array","items":{"type":"string"}},
     "confidence":{"type":"number"},
     "reason":{"type":"string"}}}
  ```
  (`minimum`/`maxItems` 같은 수치 제약은 구조화 출력이 지원하지 않음 → 0~1 범위·3개 제한은 코드에서)
- **응답 읽기**: `content`에서 `type === "text"` 블록만 골라 `JSON.parse`(앞에 빈 `thinking` 블록이 올 수 있음). `stop_reason`이 `end_turn`이 아니면 → `max_tokens`/`refused` 상태로 기록하고 장부에 안 씀.
- **캐시**: 캐시 접두부(규칙+사전)는 과목별로 바이트가 같아야 함 → 사전은 결정적 정렬(`localeCompare(…,'ko')`), 날짜·실행ID 등은 user 메시지에만. 최소 캐시 길이 512토큰은 사전만으로 넘음. 배치는 처리 시간이 길어 `ttl: "1h"`.
- **이미지**: 풀이의 `![…](drive 링크)`는 모델이 볼 수 없으므로 보내기 전에 `[그림]`으로 치환(토큰 절약)하고 `has_figure` 표시(N5).
- **거절(refusal)**: HTTP 200 + `stop_reason: "refusal"`로 옴. 수학 문항에서 드물지만 `refused`로 기록(N4).

---

## 3. 설계 (v1 대비 바뀐 부분)

### 3-1. 실행 방식 — 배치 제출 + 결과 수거 (N1 권장안)

```
[메뉴] 제출  : 대상 수집 → 시작 전 집계 확인창 → 요청 생성 → POST /v1/messages/batches (300건 단위로 나눔)
               → state = {runId, mode, batches:[{id, n}], counts} 저장 → 10분 뒤 수거 트리거
[트리거] 수거: 각 batch GET → processing_status === 'ended' 인 것만 results_url(JSONL) 받기
               → custom_id 로 행 찾기 → 검증 → 재확인(B6) → L·M 묶음 기록 → 로그 묶음 기록
               → 남은 batch 있으면 10분 뒤 재예약, 없으면 완료 요약 toast
[메뉴] 즉시 판정: 현재 시트에서 선택한 행(최대 8행)을 동기 호출(fetchAll)로 바로 판정·기록 — 스모크 테스트·소량 재시도용
```
- 수거 tick은 수 초~수십 초 → 90분/일 예산 문제와 잠금 경합(C4·C5)이 사실상 사라짐.
- 배치는 보통 1시간 안에, 늦어도 24시간 안에 끝남. 결과는 29일 보관.
- `custom_id`는 영문·숫자·`_`·`-` 64자 이내 → `F2001_S0001`(F=[풀세트], C=[써킷]) 형식. 행 메타데이터는 state에 저장하지 않고(ScriptProperties 값당 9KB 한계) **수거 시점에 그 행을 다시 읽어 재구성**(과목 결정도 그때 다시 — 사이에 F·G가 바뀌었으면 바뀐 값 기준).
- 실행은 한 번에 하나만: 제출 중인 배치가 있으면 새 제출을 막고 "상태 확인/중지"를 안내.
- 중지 = 남은 배치에 `POST …/cancel` + 트리거 삭제 + state 삭제. 이미 끝난 결과는 버리지 않고 "결과만 수거" 선택지 제공.
- UrlFetch POST 한 번에 50MB 한도 → 요청 1건이 system(~20KB)+user(~10KB) 수준이라 300건 ≈ 9MB로 여유.

### 3-2. 대상 수집

```
시트: AI_CFG.TARGET_SHEETS = ['[풀세트]','[써킷]']          (D1)
접두어: 시작 시 S / D / 둘 다 선택 (기본 S)                  (N3, 레포의 D/S 선택 관례)
행 조건: A=^[DS]\d{4}$ 이고 선택한 접두어, E·G 채워짐, L 빈칸 (D2), 행 범위 입력 시 그 범위
제외 : 로그의 보류·실패 상태(hold_subject·unit_invalid·write_error·refused) — 강제 모드면 포함
       no_source(Data_DS 없음) — 제외가 아니라 그냥 "이번엔 못 함"으로 집계
```
- 시작 전 확인창에 시트별 **후보 수 / Data_DS 있음 / 없음(no_source) / 보류로 건너뜀**을 보여 주고 진행 여부를 물음(`문항구글문서 자동생성.gs` v3의 "시작 전 요약 확인" 관례).
- Data_DS 적재는 B8대로 A·L만 먼저 읽고, 요청을 만들 때 필요한 행의 B·C만 읽음.

### 3-3. 과목 결정 — v1 유지

① F가 값이고 4종 중 하나 → F ② Data_DS id 과목코드 ③ G≤22 → `1 공통` ④ 셋 다 없으면 15개 전체 허용 + review. 경로끼리 다르면 `hold_subject`(D8). 파일럿은 사람 L의 과목 사용.
- 참고(실측): S 행은 F가 값으로 채워져 있고 Data_DS 코드도 있어 대부분 ①②가 함께 성립.

### 3-4. 태그 사전 — v1 유지 + 보완

- 생성·병합 방식(기존 D·E 보존, 정규화 키 인접 정렬)은 v1 그대로.
- **빈도 1 키워드의 기본 "사용" 값**(N8): 오타·일회성 표현(`다항함수 결저`, `중복조합(?)`)이 대부분이라 기본 FALSE를 권장.
- 잡음 키워드(`문항 폐기`, `사실상 미분` 등)는 덕수님이 E열 체크 해제.
- 사전은 **제출 시점에 한 번 읽어 그 실행의 모든 요청에 고정**(배치는 자연히 그렇게 됨). 실행 중 사전을 고쳐도 그 실행에는 반영되지 않음 — 다음 실행부터.

### 3-5. 프롬프트 — v1 문구 유지, 세 가지 수정

1. 출력 형식 지시(“JSON 객체 하나만, 코드펜스 금지”)는 구조화 출력이 보장하므로 삭제.
2. "unit에는 번호가 가장 큰 것" 규칙은 그대로 두되 최종 결정은 코드(D4) — v1과 같음.
3. `[그림]` 자리표시 안내 한 줄 추가: "풀이에 [그림]이 있으면 그림 내용은 볼 수 없다. 본문 서술만으로 판단하고, 그림 없이 판단이 어려우면 confidence를 낮춘다."
- "생각을 출력하라" 류의 지시는 넣지 않음(Opus 5.5는 내부 추론을 응답에 옮기라는 요청을 `reasoning_extraction`으로 거절할 수 있음). `reason` 한 문장은 허용 범위.

### 3-6. 검증·기록

- 검증 순서: `stop_reason` → JSON → 대단원(enum이라 사실상 항상 통과, 방어용 검사 유지) → 키워드 사전 정규화(버린 키워드 기록) → 과목 불일치(D8) → confidence(D7).
- **기록 직전 재확인(B6)**: 행 묶음의 A·L·M을 한 번에 다시 읽어, A가 그 ID인지, L이 아직 빈칸인지 확인. 통과한 행만 L·M을 연속 구간 단위로 `setValues`. 예외 시 행 단위로 다시 시도해 문제 행만 `write_error`.
- D2에 따라 M에 사람 값이 있어도 덮어쓰되, **이전 M 값을 로그에 남김**(되돌리기용, §3-8).
- 쓰지 않는 열: A, F, I, J, K, Q, P(D6), 그리고 Data 시트 전체. 원본 시트 열 추가 금지(v1과 같음).

### 3-7. 로그 시트 `_AI판정로그` (v2 열 정의, 25열)

시각 · 실행ID · 모드 · 시트 · 행 · ID · 상태 · 과목 · 과목경로 · 대단원(AI) · 키워드(AI) · 버린 키워드 · 복수 대단원 · confidence · 사유(AI) · 모델/effort · 입력자수 · 그림/잘림/중복 · 토큰 in/cacheR/cacheW/out · 이전 M · 비고/오류 · 대단원(사람) · 키워드(사람) · 대단원 일치 · 키워드 겹침

- 상태 값: `ok` · `review` · `hold_subject` · `unit_invalid` · `api_error` · `max_tokens` · `refused` · `write_skipped`(사람 값 발견) · `id_moved` · `write_error` · `undone` · `pilot` · `summary`
- `no_source`는 행마다 남기지 않고 `summary` 행에 개수만.
- 기존 v1 로그 시트가 있으면 열이 달라지므로 v2는 시트를 새로 만들고 v1 시트는 `_AI판정로그_v1`로 이름만 바꿔 보존.

### 3-8. 되돌리기 (신규)

메뉴 `AI 판정 · 실행 되돌리기`: 실행ID 입력 → 그 실행의 `ok`/`review` 로그 행마다 현재 L·M이 **AI가 쓴 값과 같을 때만** L을 비우고 M을 "이전 M"으로 복원 → 로그 `undone`. 사람이 그 뒤에 고친 행은 건드리지 않음. (P열 이력을 남기지 않기로 한 D6의 안전장치)

### 3-9. 잠금·트리거·오류 복구

- 잠금: ScriptLock은 **임대 획득 순간(수 ms)에만** 쓰고, 실제 작업 중에는 ScriptProperties `AI_TAG_LEASE`(만료시각 = 지금+6분)로 표시. 다른 AI 실행이 임대 중이면 2분 뒤 재예약. CroP·문서생성 스크립트는 건드리지 않음.
- 수거 핸들러 시작 시 자기 트리거 정리(`ai_deleteTriggers_`), 끝에 state 재확인 후 재예약(B3).
- 예외: `try/catch`로 잡아 로그 `api_error`/비고에 남기고 재예약, 연속 3회면 중지 + toast(B2).
- 트리거 실행 중 `SpreadsheetApp.getUi()` 금지 — v1과 같음.

### 3-10. API 키 보관 (N6)

v1은 ScriptProperties. 이 스프레드시트를 편집자로 공유받은 사람은 스크립트 편집기에서 스크립트 속성을 볼 수 있음 → **UserProperties**(키를 저장한 본인만 읽음, 그 사람이 만든 트리거에서도 읽힘)로 바꾸고, Anthropic Console에서 이 키의 워크스페이스에 **지출 한도**를 걸어 두기를 권장.

### 3-11. 파일·메뉴

- `AiTagging.gs` v2 (레포에 작성 완료 — 전체 교체본, `ai_`/`AI_` 접두어, 다른 파일 의존 없음, 전역 이름 충돌 검사 통과)
- `MainMenu.gs` — `menu.addToUi();` 바로 앞에 추가(작성 완료):
  ```js
  // ---- 그룹 : AI 대단원·태그 판정 (AiTagging.gs) ----
  menu.addSeparator();
  menu.addItem('AI 판정 · API 키 저장 (내 계정)', 'ai_setApiKey');
  menu.addItem('AI 태그 사전 만들기/갱신 (_태그사전)', 'ai_buildTagDict');
  menu.addItem('AI 판정 ① 파일럿 제출 (장부에 쓰지 않음)', 'ai_pilotSubmit');
  menu.addItem('AI 판정 ② 본 실행 제출', 'ai_tagSubmit');
  menu.addItem('AI 판정 · 강제 제출 (보류/실패 행 포함)', 'ai_tagSubmitForce');
  menu.addItem('AI 판정 · 선택 행 즉시 판정 (최대 8행)', 'ai_tagSelectionNow');
  menu.addItem('AI 판정 · 상태 / 지금 수거', 'ai_tagStatus');
  menu.addItem('AI 판정 · 중지 (배치 취소)', 'ai_tagStop');
  menu.addItem('AI 판정 · 실행 되돌리기', 'ai_undoRun');
  ```
- 트리거 핸들러: `ai_batchPoll_` (밑줄 함수도 트리거 대상 가능 — 레포의 `processNextBatch_`와 같은 방식)
- 로그 구간 방식: 한 실행의 로그 행은 **연속 구간**(제출 시 `submitted`로 한 번에 기록, state에 `logStart`·`logCount`). 수거 때 그 구간을 통째로 읽어 메모리에서 갱신한 뒤 한 번에 되씀 — 제출 시 알아낸 정보(입력 자수·그림/잘림/중복·이전 M·사람 L/M)가 ScriptProperties(9KB) 없이 보존됨. 실행이 진행 중일 때는 즉시 판정·되돌리기가 로그에 끼어들지 못하게 막음(구간 연속성 보장)

---

## 4. 결정

### 4-1. 확정 사항 (2026-11-07) — 유지/보완

| # | 결정 | v2 처리 |
|---|---|---|
| D1 | [풀세트]+[써킷] | 유지. 현재 [써킷] 대상은 Data_DS에 0건이라 전부 no_source로 집계됨 |
| D2 | L 빈칸이면 대상, M은 덮어씀 | 유지 + 이전 M을 로그에 보존(되돌리기용) |
| D3 | 중복 jangbu_id는 마지막 행 | 유지 + 중복 수를 로그에 기록 |
| D4 | 공통 내 복수 개념은 번호 큰 쪽 | 유지 |
| D5 | Opus 5.5 단독 | 유지. effort는 N2에서 결정 |
| D6 | P열 이력 없음 | 유지 + 되돌리기 메뉴(§3-8)로 보완 |
| D7 | 낮은 confidence는 기록 + review | 유지. 모델이 스스로 매긴 confidence는 보정이 안 돼 있으므로 기준값 0.6은 파일럿 결과로 다시 정함 |
| D8 | 과목 경로 불일치 → 보류 | 유지 |
| D9 | 사전은 1회 생성 후 덕수님 정리 | 유지 + N8 |
| D10 | 파일럿 과목별 25행 × 4 | 유지. 25행이면 정확도 오차가 ±15%p 안팎이라 "큰 문제 발견"용으로 보고, 세밀한 비교가 필요하면 N2의 두 번째 실행을 같은 표본으로 |

### 4-2. 추가 결정 — **2026-11-07 N1~N9 모두 권장안으로 확정** (코드에 반영됨)

| # | 질문 | 확정(권장안) | 다른 선택지와 이유 |
|---|---|---|---|
| **N1** | 실행 방식 | **배치 API(본 실행·파일럿) + 소량 즉시 판정 메뉴** | 동기 fetchAll(v1): 결과가 바로 보이지만 하루 90분 트리거 예산·잠금 경합·UrlFetch 응답 대기(시간 설정 불가) 위험이 있고 단가가 2배 |
| **N2** | effort | **파일럿을 같은 100행으로 medium·low 두 번 돌려 비교 후 결정**(대단원 정확도 차이 ≤2%p이고 키워드 겹침이 비슷하면 low) | medium 고정: 비교 비용(배치 약 $2) 없이 시작. high 이상: 분류 작업이라 이득이 적을 가능성이 큼 |
| **N3** | 대상 접두어 | **S만(기본값), 시작 시 D/S/둘 다 선택 가능** | 둘 다: L이 빈 D행도 함께 채워짐 — 사람이 일부러 비워 둔 D행이 있다면 원치 않는 기록 |
| **N4** | 거절(refusal) 대비 다른 모델 자동 재시도(`fallbacks`) | **쓰지 않음, `refused`로 기록 후 사람 검토** | 사용: 거절 시 다른 모델이 답함 — D5(Opus 단독)와 어긋나고, 배치 API에서는 이 옵션 자체를 받지 않음. 수학 문항이라 거절은 드물 것 |
| **N5** | 그림이 있는 문항 | **v2는 그림 미전달 + `has_figure` 표시, 파일럿에서 그림 유무별 정확도를 따로 봄** | 그림 전달(Drive 이미지 → base64): 기하·그래프 문항 품질이 오를 수 있으나 구현·비용이 늘어남 → 파일럿에서 차이가 크면 v3에서 |
| **N6** | API 키 보관 위치 | **UserProperties + Console 지출 한도** | ScriptProperties(v1): 스크립트 편집 권한이 있는 사람이면 누구나 키를 볼 수 있음 |
| **N7** | 로그 시트 v1 → v2 | **v1 시트 이름 변경 보존, v2 새로 생성** | 같은 시트 계속 사용: 열 위치가 달라 건너뛰기 목록·요약이 잘못 읽힘 |
| **N8** | 사전의 빈도 1 키워드 | **기본 "사용 안 함"(E=FALSE), 덕수님이 필요한 것만 켬** | 기본 사용: 오타·일회성 표현까지 후보가 되어 키워드 품질이 떨어짐 |
| **N9** | `review` 행 표시 | **장부에는 표시 없이 로그만**(원본 시트 서식·열 변경 없음) | L·M 셀 배경색 표시: 한눈에 보이지만 원본 시트 서식을 바꿈 — 원하시면 되돌리기 때 색도 지우도록 구현 |

---

## 5. 비용·시간 재추정

행당 가정(파일럿 토큰 열로 확정): 비캐시 입력(문제+풀이) ~2.5k · 캐시 읽기(규칙+사전) ~4k · **출력(thinking 포함) 1~2.5k**.
단가(Opus 5.5): 입력 $4 · 출력 $20 · 캐시 읽기 $0.20 · 5분 캐시 쓰기 $5 · 1시간 캐시 쓰기 $8 (MTok당), **배치는 입력·출력 50%**.

| 범위 | 동기 | 배치 |
|---|---|---|
| 행당 | 약 $0.03~0.06 | 약 $0.015~0.03 |
| 즉시 판정 스모크 5행 | ~$0.3 | — |
| 파일럿 100행 (effort 1회) | ~$3~6 | **~$1.5~3** (N2로 2회면 두 배) |
| 지금 가능한 S0001~S0368 | ~$11~22 | **~$6~11** |
| S 전체 ~2,000행(변환 완료 후) | ~$60~120 | **~$30~60** |

v1의 "2,000행 $25"는 thinking 출력을 빼고 계산한 값이라 실제보다 낮습니다. 시간은 배치 1회당 보통 1시간 이내(최대 24시간), 수거 tick은 회당 수십 초.

---

## 6. 단계 계획

| 단계 | 내용 | 완료 기준 |
|---|---|---|
| 0-a 준비 | `AiTagging.gs`·`MainMenu.gs`를 Apps Script 프로젝트에 반영(`clasp push`) → 시트 새로고침 → 메뉴 'API 키 저장'(N6, 내 계정에만 저장) → Console 지출 한도 → `ai_buildTagDict` → `_태그사전` 정리(N8) | 사전 정리 끝 |
| 0-b 스모크 | **즉시 판정**으로 이미 L·M이 있는 D행 3~5개(장부 쓰기 끔) → 요청 형식·`stop_reason`·토큰·소요 시간 확인 | 400·파싱 오류 0건, 캐시 읽기 토큰 > 0 |
| 0-c 파일럿 | 배치 100행 × effort(medium[, low]) → 과목별 대단원 정확도·키워드 겹침률·1개 이상 겹침·그림 유무별 정확도·행당 비용 | N2·N5·D7 기준값 결정, 프롬프트·사전 확정 |
| 1 본 실행 | S0001~S0368에서 행 범위 50행 먼저 → 눈으로 확인 → 나머지 | L·M 기록 + 로그 |
| 2 검토 | `review`·`hold_subject`·`refused` 행 사람 검토, 필요하면 되돌리기·강제 재제출 | — |
| 3 후속 | Latex 변환이 늘면 같은 메뉴 재실행(no_source가 자동으로 대상이 됨). Pipeline 연동·그림 전달(N5)은 그 뒤 검토 | — |

구현 산출물: `AiTagging.gs` v2 전체 교체본, `MainMenu.gs` 수정(§3-11). 레포가 clasp 루트이므로 레포에서 고친 뒤 `clasp push` 또는 편집기에 붙여넣기.

---

## 7. 리스크·주의

- 원본 시트 열 추가 · Data 시트 쓰기 · A열 쓰기 금지(2026-10-24 Data #REF! 사고 재발 방지) — v1과 같음.
- 파일럿 편향: 사전은 사람 태그(Data!M)로 만들었으므로 표본 행의 사람 태그 문자열이 사전에 들어 있음. 실제 S 행보다 키워드 일치율이 조금 높게 나올 수 있음.
- 선택항목!I열(드롭다운)은 Data!M에서 자동 생성 → AI가 쓴 조합도 드롭다운에 나타남(사전 밖 단어는 쓰지 않으므로 어휘는 늘지 않음).
- 배치 결과가 늦게 오는 동안 E·G가 바뀌거나(문제이동) 사람이 L을 채우면 → 수거 시 재구성·재확인으로 걸러짐(B6).
- Opus 5.5 rate limit은 계정 등급별. 배치는 별도 한도라 본 실행 영향이 적음.
- 같은 사전·프롬프트로 다시 돌려도 결과가 조금씩 다를 수 있음(temperature를 낮출 수 없음). 재실행은 L 빈칸 행만 대상이므로 기존 기록은 바뀌지 않음.

---

## 8. 발견 1~7 재검증 (2026-11-07, 코드 작성 전)

| 발견 | 재검증 방법 | 결과 |
|---|---|---|
| 1 `temperature` → 400 | Claude API 사양표의 Opus 5.5 행 "Sampling(temperature/top_p/top_k): Removed – 400" 재확인. v1이 쓰는 `anthropic-version: 2023-06-01`은 무관(모델이 거부). 2026-10-21 버전은 모든 모델에서 거부 | **유지.** v2는 샘플링 파라미터를 보내지 않음 |
| 2 `max_tokens: 600` | "Size `max_tokens` for the thinking as well as the reply — thinking counts toward `max_tokens` … a limit sized for a no-thinking route cuts replies off" 재확인 | **유지, 표현 수정**: "항상"이 아니라 "자주" 잘림. v2 16000 + `stop_reason` 분기 |
| 3 368행 | Data_DS L2:L6304 전체 재스캔: S는 S0001~S0368만, [써킷] S5001~ 없음. 추가로 `[풀세트]!L2001:M2368` 읽음 → 전부 빈칸 | **유지·보강**: 368행 전부가 대상. `no_source` 영구 건너뜀(v1 `SKIP_STATUSES`) 코드 줄 재확인 |
| 4 잠금 경합 | `crop-image-insert.gs:362-363` `tryLock(10s)` 실패 시 `return`만 하고 재예약 없음, `cropContinueInsert`도 그대로 종료. v1 `ai_tagTick_`은 `ai_deleteTriggers_()` **뒤에** `tryLock(15s)` | **유지·보강**: 잠금 실패 시 AI 실행이 통째로 사라지는 문제까지 확인. v2는 임대 방식 + 실패 시 2분 뒤 재예약 |
| 5 트리거 90분/일 → 배치 | Apps Script 할당량(소비자 계정 트리거 총 실행시간 90분/일) · 배치 문서 "All Messages API features supported (caching 등)", "`fallbacks`는 Batches에서 거부"(N4와 일치) · 구조화 출력은 Opus 5.5 지원 목록에 포함 | **유지.** 불확실 사항 1건: 동기 `UrlFetchApp.fetch`의 응답 대기 한도는 공식 문서에 명시가 없음(설정 불가). thinking 응답이 길면 즉시 판정(동기)이 시간 초과될 수 있어 즉시 판정은 8행 상한·재시도·`fetch_exception` 기록으로 방어. 본 작업은 배치라 영향 없음 |
| 6 사람 값 덮어쓰기 | v1 `ai_flushWrites_`: `sh.getRange(start, 12, height, 2).setValues(values)` — 기록 전 A·L·M 재확인 없음 | **유지.** v2는 기록 직전 A·L·M 재확인(`id_moved` / `write_skipped`) |
| 7 B1~B10 | 코드 줄 단위 재확인. B1: `fetchAll` 예외 → `continue` → `results[i]`가 `null`인 채 `ai_finalize_`의 `res.error` 접근 → TypeError(맨 윗줄 `res && res.tokens`만 보호됨). B2: `try/finally`만 있고 `catch` 없음. B3: `ai_tagStop` 후 tick 끝에서 `ai_saveState_`+트리거 생성. B4: `write_error` 로그 ID 열 `''`, `state.counts.fail++` 이중 집계. B5·B7·B8·B9·B10 동일 | **전부 유지.** v2 반영 내역은 부록 A |

재검증 중 추가로 확인한 점: `ai_`/`AI_` 접두어는 레포 전 파일에 없음(충돌 없음), `MainMenu.gs`에 `// ---- 메뉴 완성 ----` 주석 없음(E4 맞음), `CONFIG` 전역은 `구글문서버전복원.gs:13`(E1 맞음).

## 부록 A. `AiTagging.gs` v1 → v2 함수별 정정 목록

(v1 함수 → v2 함수. 작성된 `AiTagging.gs` v2 기준)

| v1 | v2 | 정정 |
|---|---|---|
| `AI_CFG` | `AI_CFG` + `AI_LOG` | `TEMPERATURE` 삭제, `MAX_TOKENS` 600→16000, `EFFORT` 추가, `API_VERSION` `2026-10-21`, `CACHE_TTL '1h'`, `SKIP_STATUSES`에서 `no_source` 제거·`refused` 추가, 배치(`BATCH_MAX=300`, `BATCH_MAX_BYTES`, `POLL_MS=10분`), 임대(`LEASE_MS`, `LEASE_RETRY_MS`), `MAX_POLL_ERRORS`, `NOW_MAX_ROWS`, `DICT_DEFAULT_ON_MIN_FREQ`. 키는 UserProperties. 로그 열 상수 `AI_LOG` |
| `ai_request_` | `ai_params_` | `temperature` 제거, `output_config:{effort, format:{json_schema}}` 추가. 동기·배치 공용 본문 |
| — | `ai_schema_` | 과목별 대단원 enum 스키마 |
| `ai_prewarm_` | (삭제) | C3 |
| `ai_systemBlocks_` | 동일 | 출력 형식 지시 삭제, `[그림]` 안내 추가, `cache_control`에 `ttl` |
| `ai_userText_` | 동일 | 이미지 마크다운 → `[그림]`, `flags`(그림/잘림/중복)·`inChars` 반환 |
| `ai_judgeBatch_` | `ai_judgeNow_` | 즉시 판정 전용. 결과 기본값 `fetch_exception`, `fetchAll` 예외 시 건별 재시도, `retry-after` 존중 |
| `ai_parseResponse_` | `ai_parseMessage_` | `stop_reason` 분기(`refusal`→`refused`, `max_tokens`), `text` 블록 `JSON.parse`(정규식 추출 삭제) |
| `ai_loadDataDS_` | `ai_loadDsIndex_` + `ai_loadDsText_` | A·L만 읽어 색인, 필요한 행의 B·C만 연속 구간으로 읽기, 중복 수 보존 |
| `ai_loadSkipSet_` | 동일 | 새 로그 열 기준, ID 빈 행 무시, 이후 `ok/review/undone`이 있으면 해제 |
| `ai_collectItems_` | `ai_buildItem_` + `ai_collectCandidates_` | 접두어 필터(N3), `no_source`는 집계만, 과목 불일치는 `hold` 목록 |
| `ai_resolveSubject_` | 동일 | item 객체 기반으로 정리 |
| `ai_pickPilotRows_` | 동일 | item 배열 반환 |
| `ai_finalize_` | 동일 | `max_tokens`/`refused`/`parse_error` 구분, 복수 대단원 열, 사람 값 비교를 파일럿·즉시 판정 공용으로 |
| `ai_flushWrites_` | 동일 + `ai_writeRows_` | 기록 직전 A·L·M 재확인(B6: `id_moved`/`write_skipped`), 이전 M 보존, `write_error`에 ID·사유 기록, 집계는 `ai_recount_`로 한 번만 |
| `ai_appendLog_` | `ai_logRow_` + `ai_readLogBlock_`/`ai_writeLogBlock_` | 실행 단위 연속 구간을 통째로 읽고 한 번에 되씀 |
| `ai_tagTick_` | `ai_batchPoll_` + `ai_processResults_` | 수거 전용. 임대 잠금, `try/catch` 재예약(연속 3회 실패 시 중지), 저장·재예약 전 `ai_stateAlive_` 확인 |
| `ai_start_` | `ai_submit_` | 접두어·행 범위 → 수집 → 시작 전 집계 확인 → 배치 제출(실패 시 만든 배치 취소) → 로그 구간 기록 → state → 10분 뒤 수거 예약 |
| `ai_tagStart`/`Force`/`ai_tagPilot` | `ai_tagSubmit`/`ai_tagSubmitForce`/`ai_pilotSubmit` | 배치 제출 |
| `ai_tagStop` | 동일 | 배치 취소 + "끝난 결과는 수거 후 종료 / 즉시 중지(`aborted`)" 선택 |
| `ai_tagStatus` | 동일 | 배치별 처리 현황 조회 + "지금 수거" |
| 신규 | `ai_setApiKey`·`ai_tagSelectionNow`·`ai_undoRun`·`ai_acquireLease_`·`ai_releaseLease_`·`ai_apiFetch_`·`ai_clusters_`·`ai_recount_`·`ai_applyFinal_` | — |
| `ai_buildTagDict` | 동일 | 빈도 1 키워드 기본 "사용 안 함"(N8), 기존 D·E 보존 |
| `ai_getLogSheet_` | 동일 | v1 로그 시트(열 22개) 발견 시 `_AI판정로그_v1`로 이름 변경 보존(N7) |
