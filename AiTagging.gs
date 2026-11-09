/**************************************************************
 * AiTagging.gs  v4  (2026-11-07)
 * 대단원(L)·#태그(M) AI 자동판정 — D조 문항장부 바운드 스크립트
 *
 * v4: v3 검토 반영 — ID 재탐색 시 중복 ID 보류, 본 배치 제출 때도 ID 재탐색, 예열~본 배치 사이 사전 변경 감지(캐시 미적중 경고),
 *     제출 직후 로그/state 저장 실패 시 배치 취소(고아 배치 방지), 중지 메뉴의 로그 구간 예외 내성
 *
 * 흐름 (Message Batches API, 2단계)
 *   제출  [풀세트]/[써킷]에서 A=S/D####, E·G 채워짐, L 빈칸인 행을 찾아 로그에 'queued'로 적고,
 *         과목별 첫 행 1건씩(≤4건)만 먼저 배치로 보낸다(예열 배치: 1시간 프롬프트 캐시를 만들어 둠, N10).
 *   수거  트리거(ai_batchPoll_)가 예열 배치가 끝나면 그 결과를 기록하고, 'queued' 행을 그 시점의 시트 값으로
 *         다시 만들어 본 배치(300건 단위)로 제출한다. 이후 10분마다 끝난 배치의 결과를 받아
 *         검증 → 기록 직전 그 행의 A·L·M을 다시 읽어(사람이 적었으면 건너뜀, ID가 옮겨졌으면 찾아서) L·M 기록 → 로그 갱신.
 *
 * 메뉴 진입점 (MainMenu.gs — v2와 같음)
 *   ai_setApiKey          API 키 저장 (내 계정 UserProperties)
 *   ai_buildTagDict       태그 사전 만들기/갱신 (_태그사전)
 *   ai_pilotSubmit        파일럿 제출: 이미 L·M이 있는 D행 표본으로 일치율 측정 (장부에 쓰지 않음)
 *   ai_tagSubmit          본 실행 제출
 *   ai_tagSubmitForce     본 실행 제출 — 로그의 보류/실패 행도 다시
 *   ai_tagSelectionNow    선택 행(최대 8행) 즉시 판정 (동기 호출, 스모크 테스트·소량 재시도)
 *   ai_tagStatus          상태 / 지금 수거
 *   ai_tagStop            중지 (배치 취소)
 *   ai_undoRun            실행 되돌리기 (AI가 쓴 값이 그대로일 때만 L 비우고 M 복원)
 *
 * 사전 준비
 *   1) 메뉴 'API 키 저장'으로 키 저장 (Anthropic Console 에서 이 키의 워크스페이스에 지출 한도 설정 권장)
 *   2) 메뉴 '태그 사전 만들기/갱신' → _태그사전 시트의 D(정규형)·E(사용) 정리
 *
 * 절대 하지 않는 것
 *   - 원본 시트에 열 추가 / Data 시트 쓰기 / A열 쓰기 / F·I·J·K·Q 수식열 쓰기 / P열 쓰기(D6)
 *   - 트리거 실행 중 SpreadsheetApp.getUi() 호출 (toast·Logger 사용)
 *   - 프로젝트 공용 ScriptLock 을 오래 잡기 (임대 획득 순간에만 — CroP 삽입과 충돌 방지)
 *
 * API 주의 (Opus 5.5, anthropic-version 2026-10-21) — 공식 문서로 확인(2026-11-07)
 *   - temperature/top_p/top_k 보내면 400. thinking 은 끌 수 없고 max_tokens 에 포함됨 → 16000. 샘플링은 effort 로만.
 *   - 대단원은 구조화 출력(json_schema enum)으로 허용 목록 밖 값이 나오지 않게 한다. 수치 제약(min/max)은 스키마에 못 씀.
 *   - 배치는 입력·출력 단가 50%. 결과는 29일 보관. 보통 1시간 안, 최대 24시간. max_tokens:0 예열·fallbacks 는 배치 불가.
 *   - 배치 안에서는 캐시 적중이 best-effort 라 "공유 접두부 1건을 먼저 1시간 캐시로 써 두고 나머지를 제출"하는 것이 권장 방식.
 **************************************************************/

const AI_CFG = {
  // --- 참조(Latex 변환) ---
  LATEX_SS_ID: '1wu8GdvmpyxCqSKS5RiSlRhWGH4B2OMoPN0TGP94ti54',
  DS_SHEET: 'Data_DS',
  DS_COL: { ID: 1, PROBLEM: 2, SOLUTION: 3, JANGBU: 12 },      // A, B, C, L

  // --- 주 작업(D조 문항장부) ---
  TARGET_SHEETS: ['[풀세트]', '[써킷]'],                         // D1
  SHEET_CODE: { '[풀세트]': 'F', '[써킷]': 'C' },                 // custom_id 에 쓰는 시트 약호
  SELECT_SHEET: '선택항목',
  DATA_SHEET: 'Data',
  LOG_SHEET: '_AI판정로그',
  LOG_SHEET_V1: '_AI판정로그_v1',                                // 열 수가 다른 옛 로그가 있으면 이 이름으로 보존 (N7)
  DICT_SHEET: '_태그사전',
  COL: { ID: 1, YEAR: 2, SET: 5, SUBJ: 6, NO: 7, UNIT: 12, TAG: 13 },   // A B E F G L M
  YEAR_MAX: 2026,
  ID_RE: /^[DS]\d{4}$/,
  DEFAULT_PREFIXES: 'S',                                         // N3: 기본 S, 시작 시 S/D/SD 선택

  // --- 모델/API ---
  MODEL: 'claude-opus-5-5',                                      // D5
  EFFORT: 'medium',                                              // 기본값. 제출 때 바꿀 수 있음 (N2·N12)
  EFFORT_LEVELS: ['low', 'medium', 'high', 'xhigh', 'max'],
  API_BASE: 'https://api.anthropic.com/v1',
  API_VERSION: '2026-10-21',
  API_KEY_PROP: 'ANTHROPIC_API_KEY',                             // N6: UserProperties
  MAX_TOKENS: 16000,                                             // thinking 포함
  CACHE_TTL: '1h',                                               // 배치는 처리 시간이 길어 1시간 캐시
  MAX_PROBLEM_CHARS: 4000,
  MAX_SOLUTION_CHARS: 12000,
  MAX_KEYWORDS: 3,

  // --- 배치/수거 ---
  WARM_FIRST: true,                                              // N10: 과목별 1건 예열 배치 → 본 배치
  BATCH_MAX: 300,                                                // 본 배치 1개당 요청 수
  BATCH_MAX_BYTES: 40 * 1024 * 1024,                             // UrlFetch POST 50MB 한도 아래
  POLL_FIRST_MS: 3 * 60 * 1000,                                  // 예열 배치 첫 확인
  POLL_MS: 10 * 60 * 1000,                                       // 수거 간격
  LEASE_MS: 6 * 60 * 1000,                                       // 임대(실행 중 표시) 만료
  LEASE_RETRY_MS: 2 * 60 * 1000,                                 // 임대 중이면 이 뒤에 재시도
  MAX_POLL_ERRORS: 3,                                            // 연속 오류 n회면 자동 수거 중지

  // --- 즉시 판정(동기) ---
  NOW_MAX_ROWS: 8,
  CONCURRENCY: 4,
  RETRY_WAITS_MS: [2000, 5000, 12000],

  // --- 판정 정책 ---
  REVIEW_BELOW: 0.6,                                             // D7
  PILOT_PER_SUBJECT: 25,                                         // D10
  SKIP_STATUSES: ['hold_subject', 'unit_invalid', 'write_error', 'refused'],   // 재실행 시 건너뜀 (강제 모드 제외). no_source 는 제외(C6)
  DICT_DEFAULT_ON_MIN_FREQ: 2,                                   // N8: 빈도 1 키워드는 기본 '사용 안 함'

  // --- 상태 ---
  PROP_STATE: 'AI_TAG_STATE',
  PROP_LEASE: 'AI_TAG_LEASE',
  PROP_LAST: 'AI_TAG_LAST_SUMMARY',
  POLL_FN: 'ai_batchPoll_',

  // Data_DS id 과목코드 → (사)과목
  SUBJ_CODE: { '1공통': '1 공통', '3확통': '3 확률과 통계', '4미적': '4 미적분', '5기하': '5 기하' },
};

/** _AI판정로그 열 (1-based). 시트는 전체 텍스트 서식('@') — "1/2" 같은 값이 날짜로 바뀌지 않게 */
const AI_LOG = {
  HEADER: ['시각', '실행ID', '모드', '시트', '행', 'ID', '상태', '과목', '과목경로', '대단원(AI)', '키워드(AI)',
    '버린 키워드', '복수 대단원', 'confidence', '사유(AI)', '모델/effort', '입력자수', '그림/잘림/중복',
    '토큰 in/cacheR/cacheW/out', '이전 M', '비고/오류', '대단원(사람)', '키워드(사람)', '대단원 일치', '키워드 겹침'],
  TIME: 1, RUN: 2, MODE: 3, SHEET: 4, ROW: 5, ID: 6, STATUS: 7, SUBJ: 8, PATH: 9, UNIT: 10, TAGS: 11,
  DROPPED: 12, MULTI: 13, CONF: 14, REASON: 15, MODEL: 16, INCHARS: 17, FLAGS: 18, TOKENS: 19, PREV_M: 20,
  NOTE: 21, H_UNIT: 22, H_TAGS: 23, M_UNIT: 24, M_OVL: 25,
};
/** 아직 결과가 없는(진행 중) 상태 */
const AI_PENDING = ['queued', 'submitted'];

/* =====================================================================
 * 메뉴 진입점
 * ===================================================================== */

function ai_setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Anthropic API 키 저장 (내 계정에만 저장)',
    '키를 붙여넣으세요 (sk-ant-…). 이 키로 만든 트리거는 내 계정으로 실행됩니다.\n비워 두고 확인하면 저장된 키를 지웁니다.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = String(res.getResponseText() || '').trim();
  const up = PropertiesService.getUserProperties();
  if (!key) { up.deleteProperty(AI_CFG.API_KEY_PROP); ui.alert('저장된 키를 지웠습니다.'); return; }
  if (!/^sk-ant-/.test(key)) { ui.alert('키 형식이 아닙니다 (sk-ant-… 로 시작해야 합니다).'); return; }
  up.setProperty(AI_CFG.API_KEY_PROP, key);
  ui.alert('키를 저장했습니다. Anthropic Console 에서 이 키의 워크스페이스에 지출 한도를 걸어 두세요.');
}

function ai_tagSubmit() { ai_submit_('tag', false); }
function ai_tagSubmitForce() { ai_submit_('tag', true); }
function ai_pilotSubmit() { ai_submit_('pilot', false); }

function ai_tagStatus() {
  const ui = SpreadsheetApp.getUi();
  const st = ai_loadState_();
  if (!st) {
    const last = PropertiesService.getScriptProperties().getProperty(AI_CFG.PROP_LAST) || '';
    ui.alert('진행 중인 AI 판정 작업이 없습니다.' + (last ? '\n\n마지막 요약:\n' + last : ''));
    return;
  }
  if (!ai_apiKey_()) { ui.alert(`진행 중: 실행ID ${st.runId} (${st.mode}, ${st.phase})\n\n배치 상태 조회와 수거는 API 키를 저장한 계정에서만 됩니다. (메뉴 "API 키 저장")`); return; }
  const lines = st.batches.map(b => {
    if (b.done) return `${b.id} [${b.kind}]: 수거 완료 (${b.n}건)`;
    const r = ai_apiFetch_(`/messages/batches/${b.id}`, 'get', null);
    if (!r.json) return `${b.id} [${b.kind}]: 상태 조회 실패 (${r.code})`;
    const c = r.json.request_counts || {};
    return `${b.id} [${b.kind}]: ${r.json.processing_status} (처리중 ${c.processing || 0} / 성공 ${c.succeeded || 0} / 오류 ${c.errored || 0} / 취소 ${c.canceled || 0} / 만료 ${c.expired || 0})`;
  });
  const msg =
    `실행ID ${st.runId} · 모드 ${st.mode} · 단계 ${st.phase} · ${st.model}/${st.effort}\n시작 ${st.startedAt}\n` +
    `대상 ${st.total}건 (예열 ${st.warmN}건 · 대기 ${st.queuedN}건) · Data_DS 없음 ${st.noSource} · 건너뜀 ${st.skipped} · 과목보류 ${st.hold}\n` +
    `수거 결과: 기록 ${st.counts.ok} / review ${st.counts.review} / 보류·실패 ${st.counts.fail}\n` +
    `수거 시도 ${st.polls}회 · 연속 오류 ${st.errors}회 · 예약 트리거 ${ai_hasTrigger_() ? '있음' : '없음'}${st.stopping ? ' · 중지 요청됨' : ''}` +
    `${st.lastError ? '\n마지막 오류: ' + st.lastError : ''}${st.cacheWarn ? '\n⚠️ ' + st.cacheWarn : ''}\n\n` +
    lines.join('\n') + '\n\n지금 수거할까요? (끝난 배치만 받고, 예열이 끝났으면 본 배치를 제출합니다)';
  if (ui.alert('AI 판정 상태', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  st.errors = 0; ai_saveState_(st);
  const outcome = ai_batchPoll_();
  const after = ai_loadState_();
  if (outcome === 'busy') ui.alert('다른 수거 실행이 진행 중입니다. 2분 뒤 자동으로 다시 시도합니다.');
  else if (outcome === 'error') ui.alert('수거 중 오류가 났습니다. 로그(실행 기록)를 확인하세요. 트리거는 다시 예약했습니다.');
  else if (after) ui.alert(`수거했습니다. 단계 ${after.phase}, 남은 배치 ${after.batches.filter(b => !b.done).length}개.`);
  else ui.alert('모든 배치를 수거하고 작업을 마쳤습니다.\n\n' + (PropertiesService.getScriptProperties().getProperty(AI_CFG.PROP_LAST) || ''));
}

function ai_tagStop() {
  const ui = SpreadsheetApp.getUi();
  const st = ai_loadState_();
  if (!st) { ai_deleteTriggers_(); ui.alert('진행 중인 작업이 없습니다. (남은 트리거는 정리했습니다)'); return; }
  const r = ui.alert('AI 판정 중지',
    `실행ID ${st.runId} (${st.mode}, ${st.phase})\n\n[예] 남은 배치를 취소하고, 이미 끝난 결과는 수거한 뒤 종료 (대기 행은 제출하지 않음)\n[아니오] 지금 즉시 중지 (끝난 결과도 받지 않음 — 로그는 'aborted')\n[취소] 그대로 둠`,
    ui.ButtonSet.YES_NO_CANCEL);
  if (r === ui.Button.CANCEL) return;
  const failed = [];
  st.batches.filter(b => !b.done).forEach(b => {
    const res = ai_apiFetch_(`/messages/batches/${b.id}/cancel`, 'post', {});
    if (!res.json) failed.push(`${b.id} (${res.code})`);
  });
  ai_deleteTriggers_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 로그 구간 표시 변경 — 구간이 깨져 있어도(사람이 로그에 행을 넣은 경우) 중지는 되어야 하므로 예외는 안내만
  const markAborted = statuses => {
    try {
      const block = ai_readLogBlock_(ai_getLogSheet_(ss), st);
      block.rows.forEach(row => { if (statuses.indexOf(row[AI_LOG.STATUS - 1]) >= 0) { row[AI_LOG.STATUS - 1] = 'aborted'; row[AI_LOG.TIME - 1] = ai_now_(); } });
      ai_writeLogBlock_(block);
      return '';
    } catch (e) { return `\n⚠️ 로그 상태는 바꾸지 못했습니다: ${String(e.message || e).slice(0, 160)}`; }
  };
  if (r === ui.Button.YES) {
    const warn = markAborted(['queued']);                       // 대기 행은 제출하지 않으므로 지금 aborted 로
    st.stopping = true; st.errors = 0; ai_saveState_(st);
    ScriptApp.newTrigger(AI_CFG.POLL_FN).timeBased().after(60 * 1000).create();
    ui.alert(`취소를 요청했습니다. 1분 뒤부터 끝난 결과를 수거하고 종료합니다.${failed.length ? '\n취소 실패: ' + failed.join(', ') : ''}${warn}`);
    return;
  }
  const warn = markAborted(AI_PENDING);
  PropertiesService.getScriptProperties().deleteProperty(AI_CFG.PROP_STATE);
  ui.alert(`즉시 중지했습니다. (진행 중이던 행은 로그에 'aborted' — 다음 제출에서 다시 대상이 됩니다)${failed.length ? '\n취소 실패: ' + failed.join(', ') : ''}${warn}`);
}

/* =====================================================================
 * 제출 (tag / pilot) — 대상 수집 → 로그 'queued' → 예열 배치(과목별 1건) 제출
 * ===================================================================== */

function ai_submit_(mode, force) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ai_precheck_(ui, ss)) return;
  if (ai_loadState_()) { ui.alert('진행 중인 AI 판정 작업이 있습니다. "상태 / 지금 수거" 또는 "중지" 메뉴를 먼저 쓰세요.'); return; }

  // 접두어·행 범위 (본 실행만)
  let prefixes = 'D', range = null;
  if (mode === 'tag') {
    const p = ui.prompt('대상 접두어', `S / D / SD 중 하나 (비우면 ${AI_CFG.DEFAULT_PREFIXES})`, ui.ButtonSet.OK_CANCEL);
    if (p.getSelectedButton() !== ui.Button.OK) return;
    prefixes = (String(p.getResponseText() || '').trim().toUpperCase() || AI_CFG.DEFAULT_PREFIXES);
    if (!/^(S|D|SD|DS)$/.test(prefixes)) { ui.alert('접두어는 S, D, SD 중 하나여야 합니다.'); return; }
    const r = ui.prompt('행 범위 (선택)',
      `비워 두면 ${AI_CFG.TARGET_SHEETS.join(', ')} 전체. 현재 시트의 일부만 하려면 예: 2001-2050`, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    const txt = String(r.getResponseText() || '').trim();
    if (txt) {
      const m = txt.match(/^(\d+)(?:\s*-\s*(\d+))?$/);
      if (!m) { ui.alert('범위 형식이 잘못되었습니다. 예: 2001-2050'); return; }
      const active = ss.getActiveSheet().getName();
      if (AI_CFG.TARGET_SHEETS.indexOf(active) < 0) { ui.alert(`현재 시트(${active})는 대상 시트가 아닙니다.`); return; }
      range = { name: active, from: Math.max(2, Number(m[1])), to: Number(m[2] || m[1]) };
    }
  }
  // effort (N12: 파일럿 비교용. 비우면 기본값)
  const ep = ui.prompt('effort', `${AI_CFG.EFFORT_LEVELS.join(' / ')} 중 하나 (비우면 ${AI_CFG.EFFORT})`, ui.ButtonSet.OK_CANCEL);
  if (ep.getSelectedButton() !== ui.Button.OK) return;
  const effort = String(ep.getResponseText() || '').trim().toLowerCase() || AI_CFG.EFFORT;
  if (AI_CFG.EFFORT_LEVELS.indexOf(effort) < 0) { ui.alert(`effort 는 ${AI_CFG.EFFORT_LEVELS.join(', ')} 중 하나여야 합니다.`); return; }

  const ctx = ai_loadContext_(ss, mode, force);
  ctx.effort = effort;
  const run = { items: [], hold: [], noSource: 0, skipped: 0 };
  if (mode === 'pilot') {
    run.items = ai_pickPilotRows_(ss, ctx);
  } else {
    const sheets = range ? [range] : AI_CFG.TARGET_SHEETS.map(n => ({ name: n, from: 2, to: 0 }));
    sheets.forEach(s => {
      const sh = ss.getSheetByName(s.name); if (!sh) return;
      ai_collectCandidates_(sh, s.from, s.to, prefixes, ctx, run);
    });
  }
  const n = run.items.length;

  // 예열 대상: 과목별 첫 행 1건 (N10). 대상이 과목 수 이하면 예열 없이 전부 본 배치
  let warm = [], queued = run.items;
  if (AI_CFG.WARM_FIRST) {
    const seen = {};
    run.items.forEach(it => { const k = it.subject || '*'; if (!seen[k]) { seen[k] = true; warm.push(it); } });
    if (warm.length < n) queued = run.items.filter(it => warm.indexOf(it) < 0); else { warm = []; queued = run.items; }
  }
  const head = mode === 'pilot'
    ? `파일럿: 사람이 L·M을 적어 둔 D행 표본 ${n}행 (과목별 ≤${AI_CFG.PILOT_PER_SUBJECT}). 장부에는 쓰지 않습니다.`
    : `본 실행: 접두어 ${prefixes}, 대상 ${range ? `${range.name} ${range.from}~${range.to}` : AI_CFG.TARGET_SHEETS.join(', ')}${force ? ' (강제: 보류/실패 행 포함)' : ''}`;
  const msg = `${head}\n\n제출 ${n}행 · Data_DS 없음 ${run.noSource} · 보류로 건너뜀 ${run.skipped} · 과목 경로 불일치(보류) ${run.hold.length}\n` +
    `모델 ${AI_CFG.MODEL} / effort ${effort}\n예상 비용 약 $${(n * 0.015).toFixed(1)}~${(n * 0.03).toFixed(1)} (배치 단가, thinking 포함 추정)\n` +
    (warm.length ? `\n예열 배치 ${warm.length}건을 먼저 보내고(약 3분 뒤 확인), 끝나면 나머지 ${queued.length}건을 본 배치로 제출합니다.` : '\n전부 한 번에 본 배치로 제출합니다.') +
    `\n\n제출할까요? (결과는 10분마다 자동 수거)`;
  if (!n) { ui.alert(msg.replace(/\n\n제출할까요.*$/, '\n\n제출할 행이 없습니다.')); if (run.hold.length) ai_appendHoldLog_(ss, mode, run.hold, effort); return; }
  if (ui.alert('AI 판정 제출', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const runId = ai_runId_(mode);
  // 첫 배치 요청 생성 (예열이면 예열분만, 아니면 전부)
  const first = warm.length ? warm : queued;
  const requests = ai_buildRequests_(ctx, first);
  let batches;
  try {
    batches = ai_submitBatches_(requests, warm.length ? 'warm' : 'main');
  } catch (e) {
    ui.alert(`제출 중 오류가 나서 이미 만든 배치는 취소했습니다.\n${e.message}\n\n(Anthropic Console 의 Batches 화면에서 남은 배치가 없는지 확인하세요)`);
    return;
  }

  // 로그: 보류 → 예열(submitted) → 대기(queued). 한 실행의 행은 연속 구간.
  // 배치는 이미 서버에 있으므로, 여기서 실패하면(할당량·예외) 배치를 취소해 고아 배치(과금만 되고 수거 못 함)를 막는다
  try {
    const logSh = ai_getLogSheet_(ss);
    const logStart = logSh.getLastRow() + 1;
    const rows = [];
    run.hold.forEach(it => rows.push(ai_logRow_(runId, mode, it, 'hold_subject', `과목 경로 불일치 (${it.subjPath}: F="${it.fVal}", DS="${it.dsCode}", G=${it.gNo})`, effort)));
    first.forEach(it => rows.push(ai_logRow_(runId, mode, it, 'submitted', '', effort)));
    if (warm.length) queued.forEach(it => rows.push(ai_logRow_(runId, mode, it, 'queued', '', effort)));
    logSh.getRange(logStart, 1, rows.length, AI_LOG.HEADER.length).setValues(rows);

    const state = {
      runId, mode, force: !!force, prefixes, model: AI_CFG.MODEL, effort,
      phase: warm.length ? 'warm' : 'main',
      startedAt: ai_now_(), batches,
      total: n, warmN: first.length, queuedN: warm.length ? queued.length : 0,
      noSource: run.noSource, skipped: run.skipped, hold: run.hold.length,
      logStart, logCount: rows.length, counts: { ok: 0, review: 0, fail: 0 }, polls: 0, errors: 0, stopping: false,
      promptFp: ai_promptFp_(ctx),                                  // 예열~본 배치 사이 사전/대단원 표 변경 감지용
    };
    ai_saveState_(state);
    ai_deleteTriggers_();
    ScriptApp.newTrigger(AI_CFG.POLL_FN).timeBased().after(warm.length ? AI_CFG.POLL_FIRST_MS : AI_CFG.POLL_MS).create();
  } catch (e) {
    batches.forEach(b => { try { ai_apiFetch_(`/messages/batches/${b.id}/cancel`, 'post', {}); } catch (e2) {} });
    PropertiesService.getScriptProperties().deleteProperty(AI_CFG.PROP_STATE);
    ui.alert(`배치는 만들었지만 로그/상태 저장에 실패해 배치를 취소했습니다. 다시 시도하세요.\n${e.message}`);
    return;
  }
  ui.alert(`제출했습니다. 실행ID ${runId}\n${warm.length ? '예열' : '본'} 배치: ${batches.map(b => `${b.id} (${b.n}건)`).join('\n')}\n\n` +
    (warm.length ? `약 3분 뒤 예열 결과를 확인하고 나머지 ${queued.length}건을 자동 제출합니다. ` : '') +
    `이후 10분마다 자동 수거합니다. "상태 / 지금 수거" 메뉴로 바로 진행할 수도 있습니다.`);
}

function ai_precheck_(ui, ss) {
  if (!ai_apiKey_()) { ui.alert('API 키가 없습니다. 메뉴 "AI 판정 · API 키 저장"을 먼저 실행하세요.'); return false; }
  if (!ss.getSheetByName(AI_CFG.DICT_SHEET)) { ui.alert(`'${AI_CFG.DICT_SHEET}' 시트가 없습니다. 먼저 "태그 사전 만들기/갱신"을 실행하세요.`); return false; }
  return true;
}

function ai_appendHoldLog_(ss, mode, hold, effort) {
  const logSh = ai_getLogSheet_(ss);
  const runId = ai_runId_('hold');
  const rows = hold.map(it => ai_logRow_(runId, mode, it, 'hold_subject', `과목 경로 불일치 (${it.subjPath}: F="${it.fVal}", DS="${it.dsCode}", G=${it.gNo})`, effort));
  logSh.getRange(logSh.getLastRow() + 1, 1, rows.length, AI_LOG.HEADER.length).setValues(rows);
}

function ai_runId_(mode) {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMdd-HHmmss') + ({ pilot: 'P', tag: 'T', now: 'N', hold: 'H' }[mode] || 'X');
}

/** items → Messages Batches 요청 목록. it.flags/inChars 도 채움 (로그용) */
function ai_buildRequests_(ctx, items) {
  const dsText = ai_loadDsText_(items.map(it => it.dsRow));
  return items.map(it => {
    const u = ai_userText_(it, dsText[it.dsRow] || { problem: '', solution: '' });
    it.flags = u.flags; it.inChars = u.inChars;
    return { custom_id: ai_customId_(it), params: ai_params_(ctx, it.subject, u.text) };
  });
}

/**
 * 요청을 개수·크기 기준으로 나눠 배치 생성. 실패하면 이미 만든 배치는 취소하고 throw.
 * 배치 생성 POST 는 "요청이 서버에 닿기 전" 실패(네트워크 예외·429·529)만 재시도 — 5xx 재시도는 중복 배치를 만들 수 있음 (N13)
 */
function ai_submitBatches_(requests, kind) {
  const batches = [];
  let chunk = [], bytes = 0;
  const flush = () => {
    if (!chunk.length) return;
    const r = ai_apiFetch_('/messages/batches', 'post', { requests: chunk }, { createOnly: true });
    if (!r.json || !r.json.id) throw new Error(`배치 제출 실패 (${r.code}): ${String(r.body || '').slice(0, 300)}`);
    batches.push({ id: r.json.id, n: chunk.length, kind, done: false });
    chunk = []; bytes = 0;
  };
  try {
    requests.forEach(q => {
      const len = JSON.stringify(q).length;
      if (chunk.length >= AI_CFG.BATCH_MAX || bytes + len > AI_CFG.BATCH_MAX_BYTES) flush();
      chunk.push(q); bytes += len;
    });
    flush();
  } catch (e) {
    batches.forEach(b => { try { ai_apiFetch_(`/messages/batches/${b.id}/cancel`, 'post', {}); } catch (e2) {} });
    throw e;
  }
  return batches;
}

/* =====================================================================
 * 수거 (트리거 핸들러) — 끝난 배치의 결과를 받아 검증·기록. 예열 단계가 끝나면 본 배치 제출
 * 반환: 'done' | 'ongoing' | 'busy' | 'error' | 'none'
 * ===================================================================== */

function ai_batchPoll_() {
  ai_deleteTriggers_();
  const props = PropertiesService.getScriptProperties();
  const state = ai_loadState_();
  if (!state) return 'none';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ai_acquireLease_()) {
    ScriptApp.newTrigger(AI_CFG.POLL_FN).timeBased().after(AI_CFG.LEASE_RETRY_MS).create();
    Logger.log('다른 AI 판정 실행이 진행 중 — 2분 뒤 재시도');
    return 'busy';
  }
  try {
    state.polls++;
    const logSh = ai_getLogSheet_(ss);
    let ctxLite = null, block = null;

    // 1) 끝난 배치 수거
    for (const b of state.batches) {
      if (b.done) continue;
      const r = ai_apiFetch_(`/messages/batches/${b.id}`, 'get', null);
      if (!r.json) throw new Error(`배치 상태 조회 실패 ${b.id} (${r.code})`);
      if (r.json.processing_status !== 'ended') continue;
      const jr = ai_apiFetch_(r.json.results_url, 'get', null, { raw: true });
      if (jr.code !== 200) throw new Error(`결과 다운로드 실패 ${b.id} (${jr.code})`);
      const lines = String(jr.body || '').split('\n').map(s => s.trim()).filter(Boolean).map(s => JSON.parse(s));
      if (!ctxLite) { ctxLite = ai_loadContextLite_(ss); block = ai_readLogBlock_(logSh, state); }
      ai_processResults_(ss, state, ctxLite, block, lines);
      b.done = true;
      ai_writeLogBlock_(block);
      ai_recount_(state, block);
      if (!ai_stateAlive_(state.runId)) { Logger.log('중지됨 — 저장하지 않고 종료'); return 'done'; }
      ai_saveState_(state);
    }

    // 2) 예열이 끝났으면 본 배치 제출 (중지 요청 중이면 제출하지 않음)
    if (state.phase === 'warm' && !state.stopping && state.batches.every(b => b.done)) {
      if (!block) block = ai_readLogBlock_(logSh, state);
      const added = ai_submitMain_(ss, state, block);            // 'queued' 행 → 'submitted', 배치 추가
      ai_writeLogBlock_(block);
      state.phase = 'main';
      if (!ai_stateAlive_(state.runId)) return 'done';
      ai_saveState_(state);
      ss.toast(`예열 완료 → 본 배치 ${added.length}개(${state.queuedN}건) 제출${state.cacheWarn ? ' ⚠️ ' + state.cacheWarn : ''}`, '⏳ AI 판정', state.cacheWarn ? 20 : 8);
    }

    state.errors = 0;
    const remaining = state.batches.filter(b => !b.done).length;
    if (remaining) {
      if (!ai_stateAlive_(state.runId)) return 'done';
      ai_saveState_(state);
      ScriptApp.newTrigger(AI_CFG.POLL_FN).timeBased().after(AI_CFG.POLL_MS).create();
      ss.toast(`배치 ${remaining}개 처리 대기 중 (기록 ${state.counts.ok}, review ${state.counts.review}) — 10분 뒤 다시 확인`, '⏳ AI 판정', 8);
      return 'ongoing';
    }
    if (!block) block = ai_readLogBlock_(logSh, state);
    const summary = ai_finishSummary_(ss, state, block);
    props.deleteProperty(AI_CFG.PROP_STATE);
    ss.toast(summary.replace(/\n/g, ' / ').slice(0, 300), '✅ AI 판정 완료', 15);
    Logger.log(summary);
    return 'done';
  } catch (e) {
    Logger.log('수거 오류: ' + (e && e.stack || e));
    if (!ai_stateAlive_(state.runId)) return 'done';
    state.errors = (state.errors || 0) + 1;
    state.lastError = String(e.message || e).slice(0, 200);
    ai_saveState_(state);
    if (state.errors >= AI_CFG.MAX_POLL_ERRORS) {
      ss.toast(`수거가 ${state.errors}회 연속 실패해 자동 수거를 멈췄습니다. "상태 / 지금 수거"로 재개하세요. ${state.lastError}`, '⚠️ AI 판정', 20);
    } else {
      ScriptApp.newTrigger(AI_CFG.POLL_FN).timeBased().after(AI_CFG.POLL_MS).create();
    }
    return 'error';
  } finally {
    ai_releaseLease_();
  }
}

/**
 * 예열 뒤 본 배치 제출: 로그 블록의 'queued' 행을 그 시점의 시트 값으로 다시 만들어 요청 생성 → 배치 생성 → 'submitted'
 * (사이에 사람이 L을 적었으면 write_skipped, ID가 바뀌었으면 id_moved 로 남기고 보내지 않음)
 */
function ai_submitMain_(ss, state, block) {
  const idx = [];
  block.rows.forEach((r, i) => { if (r[AI_LOG.STATUS - 1] === 'queued') idx.push(i); });
  if (!idx.length) return [];
  const ctx = ai_loadContext_(ss, state.mode, true);              // 건너뛰기 집합은 제출 때 이미 적용됨
  ctx.effort = state.effort;

  // 시트별로 모아 행을 다시 읽음
  const bySheet = {};
  idx.forEach(i => { const r = block.rows[i]; (bySheet[String(r[AI_LOG.SHEET - 1])] = bySheet[String(r[AI_LOG.SHEET - 1])] || []).push(i); });
  const items = [];
  Object.keys(bySheet).forEach(name => {
    const sh = ss.getSheetByName(name);
    const list = bySheet[name];
    if (!sh) { list.forEach(i => { block.rows[i][AI_LOG.STATUS - 1] = 'write_error'; block.rows[i][AI_LOG.NOTE - 1] = '시트 없음'; }); return; }
    const rowOf = {}; list.forEach(i => { rowOf[Number(block.rows[i][AI_LOG.ROW - 1])] = i; });
    const idCache = {};
    ai_clusters_(Object.keys(rowOf).map(Number), 1).forEach(cl => {
      const vals = sh.getRange(cl.start, 1, cl.len, AI_CFG.COL.TAG).getDisplayValues();
      for (let r = cl.start; r < cl.start + cl.len; r++) {
        const i = rowOf[r]; if (i === undefined) continue;
        const lr = block.rows[i];
        const wantId = String(lr[AI_LOG.ID - 1]);
        let it = ai_buildItem_(name, r, vals[r - cl.start], ctx, state.mode);
        let moved = '';
        lr[AI_LOG.TIME - 1] = ai_now_();
        if (!it || it.id !== wantId) {                                 // 그 사이 행이 옮겨짐 → A열에서 ID 재탐색 (N11, 기록 때와 같은 규칙)
          const f = ai_findRowById_(sh, wantId, idCache);
          if (!f.row || f.dup) { lr[AI_LOG.STATUS - 1] = 'id_moved'; lr[AI_LOG.NOTE - 1] = `제출 시점에 그 행의 ID가 "${vals[r - cl.start][0]}"${f.dup ? `, ${wantId} 가 시트에 ${f.dup}번 있음` : `, 시트에서 ${wantId} 를 찾지 못함`}`; continue; }
          it = ai_buildItem_(name, f.row, sh.getRange(f.row, 1, 1, AI_CFG.COL.TAG).getDisplayValues()[0], ctx, state.mode);
          if (!it) { lr[AI_LOG.STATUS - 1] = 'id_moved'; lr[AI_LOG.NOTE - 1] = `${f.row}행의 ID를 읽지 못함`; continue; }
          moved = `행 이동 ${r}→${f.row}`; lr[AI_LOG.ROW - 1] = f.row;
        }
        if (state.mode === 'tag' && it.unitHuman) { lr[AI_LOG.STATUS - 1] = 'write_skipped'; lr[AI_LOG.NOTE - 1] = ai_joinNote_(moved, `사람이 먼저 적음: L="${it.unitHuman}"`); continue; }
        if (!it.dsRow) { lr[AI_LOG.STATUS - 1] = 'no_source'; lr[AI_LOG.NOTE - 1] = ai_joinNote_(moved, '제출 시점에 Data_DS 에서 사라짐'); continue; }
        if (state.mode === 'tag' && it.subjMismatch) { lr[AI_LOG.STATUS - 1] = 'hold_subject'; lr[AI_LOG.NOTE - 1] = ai_joinNote_(moved, `과목 경로 불일치 (${it.subjPath})`); continue; }
        it.logIdx = i; it.moved = moved;
        items.push(it);
      }
    });
  });
  if (!items.length) return [];
  const requests = ai_buildRequests_(ctx, items);
  // 예열 때와 system(규칙+사전)이 달라졌으면 1시간 캐시를 못 쓴다 — 막지는 않고 로그·toast 로 알림 (N15)
  const fpNow = ai_promptFp_(ctx), fpOld = state.promptFp || {};
  const changed = Object.keys(fpNow).filter(k => fpOld[k] !== undefined && fpOld[k] !== fpNow[k]);
  if (changed.length) {
    state.cacheWarn = `예열 뒤 사전/대단원 표가 바뀜 (${changed.join(', ')}) → 본 배치는 캐시 미적중, 입력 비용 증가`;
    Logger.log(state.cacheWarn);
  }
  const batches = ai_submitBatches_(requests, 'main');            // 실패 시 throw → 수거 오류 경로(재시도)
  items.forEach(it => {
    const lr = block.rows[it.logIdx];
    lr[AI_LOG.STATUS - 1] = 'submitted'; lr[AI_LOG.SUBJ - 1] = it.subject || ''; lr[AI_LOG.PATH - 1] = it.subjPath || '';
    lr[AI_LOG.INCHARS - 1] = it.inChars; lr[AI_LOG.FLAGS - 1] = it.flags || ''; lr[AI_LOG.PREV_M - 1] = it.prevM || '';
    if (it.moved) lr[AI_LOG.NOTE - 1] = ai_joinNote_(lr[AI_LOG.NOTE - 1], it.moved);
    if (state.mode !== 'tag') { lr[AI_LOG.H_UNIT - 1] = it.unitHuman || ''; lr[AI_LOG.H_TAGS - 1] = it.tagHuman || ''; }
  });
  batches.forEach(b => state.batches.push(b));
  return batches;
}

/** 결과 JSONL 한 묶음을 로그 블록에 반영하고, tag 모드면 L·M 기록 */
function ai_processResults_(ss, state, ctx, block, lines) {
  const pending = {};   // sheetName → { row → {id, values, logRow} }
  lines.forEach(line => {
    const cid = String(line.custom_id || '');
    const m = cid.match(/^([FC])(\d+)_([DS]\d{4})$/);
    if (!m) { Logger.log('custom_id 해석 불가: ' + cid); return; }
    const sheetName = Object.keys(AI_CFG.SHEET_CODE).find(k => AI_CFG.SHEET_CODE[k] === m[1]);
    const row = Number(m[2]), id = m[3];
    const i = block.index[`${sheetName}|${row}|${id}`];
    if (i === undefined) { Logger.log('로그 블록에 없는 결과: ' + cid); return; }
    const lr = block.rows[i];
    if (lr[AI_LOG.STATUS - 1] !== 'submitted') return;              // 이미 반영됨
    const subject = String(lr[AI_LOG.SUBJ - 1] || '');
    const res = line.result || {};
    let fin;
    if (res.type === 'succeeded') {
      fin = ai_finalize_(ctx, state.mode, subject, ai_parseMessage_(res.message), { unit: lr[AI_LOG.H_UNIT - 1], tags: lr[AI_LOG.H_TAGS - 1] });
    } else if (res.type === 'errored') {
      const err = res.error && res.error.error ? res.error.error : res.error || {};
      fin = { status: 'api_error', note: `${err.type || ''} ${String(err.message || '').slice(0, 200)}`.trim() };
    } else {
      fin = { status: res.type === 'canceled' ? 'canceled' : res.type === 'expired' ? 'expired' : 'api_error', note: res.type || '' };
    }
    ai_applyFinal_(lr, fin);
    if (state.mode === 'tag' && fin.write) {
      (pending[sheetName] = pending[sheetName] || {})[row] = { id, values: fin.write, logRow: lr };
    }
  });
  Object.keys(pending).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { Object.values(pending[name]).forEach(p => { p.logRow[AI_LOG.STATUS - 1] = 'write_error'; p.logRow[AI_LOG.NOTE - 1] = '시트 없음'; }); return; }
    ai_flushWrites_(sh, pending[name]);
  });
}

/** finalize 결과를 로그 행 배열에 반영 */
function ai_applyFinal_(lr, fin) {
  lr[AI_LOG.TIME - 1] = ai_now_();
  lr[AI_LOG.STATUS - 1] = fin.status;
  lr[AI_LOG.UNIT - 1] = fin.unit || '';
  lr[AI_LOG.TAGS - 1] = fin.tags || '';
  lr[AI_LOG.DROPPED - 1] = fin.dropped || '';
  lr[AI_LOG.MULTI - 1] = fin.multi || '';
  lr[AI_LOG.CONF - 1] = fin.conf === undefined ? '' : fin.conf;
  lr[AI_LOG.REASON - 1] = fin.reason || '';
  lr[AI_LOG.TOKENS - 1] = fin.tokens || '';
  lr[AI_LOG.NOTE - 1] = fin.note || '';
  if (fin.match) { lr[AI_LOG.M_UNIT - 1] = fin.match.unit; lr[AI_LOG.M_OVL - 1] = `${fin.match.overlap}/${fin.match.humanN}`; }
}

/** 로그 블록에서 상태를 다시 세어 counts 갱신 (이중 집계 방지). 진행 중 상태는 세지 않음 */
function ai_recount_(state, block) {
  const c = { ok: 0, review: 0, fail: 0 };
  block.rows.forEach(r => {
    const s = r[AI_LOG.STATUS - 1];
    if (s === 'ok' || s === 'pilot') c.ok++;
    else if (s === 'review') c.review++;
    else if (AI_PENDING.indexOf(s) < 0) c.fail++;
  });
  state.counts = c;
}

/* =====================================================================
 * 즉시 판정 (동기) — 선택 행 최대 8행
 * ===================================================================== */

function ai_tagSelectionNow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ai_precheck_(ui, ss)) return;
  if (ai_loadState_()) { ui.alert('배치 작업이 진행 중입니다. 끝난 뒤 사용하세요.'); return; }
  const sh = ss.getActiveSheet();
  if (AI_CFG.TARGET_SHEETS.indexOf(sh.getName()) < 0) { ui.alert(`현재 시트(${sh.getName()})는 대상 시트가 아닙니다.`); return; }
  const rowsSet = {};
  const rl = ss.getActiveRangeList();
  const ranges = rl ? rl.getRanges() : (ss.getActiveRange() ? [ss.getActiveRange()] : []);
  ranges.forEach(rg => { for (let r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) if (r >= 2) rowsSet[r] = 1; });
  const rows = Object.keys(rowsSet).map(Number).sort((a, b) => a - b);
  if (!rows.length) { ui.alert('행을 선택한 뒤 실행하세요.'); return; }
  if (rows.length > AI_CFG.NOW_MAX_ROWS) { ui.alert(`즉시 판정은 최대 ${AI_CFG.NOW_MAX_ROWS}행입니다. (선택 ${rows.length}행)`); return; }
  const w = ui.alert('즉시 판정', `${rows.length}행을 동기 호출로 판정합니다 (${AI_CFG.MODEL}/${AI_CFG.EFFORT}, 건당 수십 초).\n\n장부 L·M에 기록할까요? (L이 비어 있는 행만 기록되고, 이미 L이 있는 행은 '판정만'과 같습니다)\n[예] 기록  [아니오] 판정만 (로그에만)`, ui.ButtonSet.YES_NO_CANCEL);
  if (w === ui.Button.CANCEL) return;
  const writeMode = w === ui.Button.YES;

  const ctx = ai_loadContext_(ss, 'now', true);
  ctx.effort = AI_CFG.EFFORT;
  const items = [], problems = [];
  rows.forEach(row => {
    const r = sh.getRange(row, 1, 1, AI_CFG.COL.TAG).getDisplayValues()[0];
    const it = ai_buildItem_(sh.getName(), row, r, ctx, 'now');
    if (!it) { problems.push(`${row}행: ID 아님`); return; }
    if (!it.dsRow) { problems.push(`${row}행 ${it.id}: Data_DS 없음`); return; }
    if (it.subjMismatch) { problems.push(`${row}행 ${it.id}: 과목 경로 불일치 (${it.subjPath}) → 판정은 F 기준으로 진행, 기록 안 함`); it.noWrite = true; }
    items.push(it);
  });
  if (!items.length) { ui.alert('판정할 행이 없습니다.\n' + problems.join('\n')); return; }

  const reqs = ai_buildRequests_(ctx, items).map(q => q.params);
  const results = ai_judgeNow_(reqs);
  const runId = ai_runId_('now');
  const logRows = [], pending = {};
  items.forEach((it, k) => {
    const lr = ai_logRow_(runId, 'now', it, 'submitted', '', ctx.effort);
    const fin = ai_finalize_(ctx, 'now', it.subject, results[k], { unit: it.unitHuman, tags: it.tagHuman });
    ai_applyFinal_(lr, fin);
    if (fin.write && writeMode && !it.noWrite) pending[it.row] = { id: it.id, values: fin.write, logRow: lr };
    else if (fin.write) lr[AI_LOG.NOTE - 1] = ai_joinNote_(lr[AI_LOG.NOTE - 1], '기록 안 함(판정만)');
    logRows.push(lr);
  });
  if (Object.keys(pending).length) ai_flushWrites_(sh, pending);
  const logSh = ai_getLogSheet_(ss);
  logSh.getRange(logSh.getLastRow() + 1, 1, logRows.length, AI_LOG.HEADER.length).setValues(logRows);
  const lines = logRows.map(lr => `${lr[AI_LOG.ROW - 1]}행 ${lr[AI_LOG.ID - 1]}: ${lr[AI_LOG.STATUS - 1]} · ${lr[AI_LOG.UNIT - 1]} · ${lr[AI_LOG.TAGS - 1]} (conf ${lr[AI_LOG.CONF - 1]}, 토큰 ${lr[AI_LOG.TOKENS - 1]})${lr[AI_LOG.M_UNIT - 1] !== '' ? ` · 사람 대단원 일치 ${lr[AI_LOG.M_UNIT - 1]}, 키워드 ${lr[AI_LOG.M_OVL - 1]}` : ''}${lr[AI_LOG.NOTE - 1] ? ' · ' + lr[AI_LOG.NOTE - 1] : ''}`);
  ui.alert(`즉시 판정 결과 (실행ID ${runId})\n\n${lines.join('\n')}${problems.length ? '\n\n' + problems.join('\n') : ''}`);
}

/** 동기 호출: fetchAll + 재시도. params 와 같은 길이의 parsed 결과 배열 */
function ai_judgeNow_(params) {
  const results = params.map(() => ({ error: 'fetch_exception', raw: '' }));
  let pending = params.map((p, i) => i);
  for (let attempt = 0; attempt <= AI_CFG.RETRY_WAITS_MS.length && pending.length; attempt++) {
    if (attempt > 0) Utilities.sleep(AI_CFG.RETRY_WAITS_MS[attempt - 1]);
    const next = [];
    for (let c = 0; c < pending.length; c += AI_CFG.CONCURRENCY) {
      const idx = pending.slice(c, c + AI_CFG.CONCURRENCY);
      let resps;
      try { resps = UrlFetchApp.fetchAll(idx.map(i => ai_fetchAllReq_('/messages', 'post', params[i]))); }
      catch (e) {
        // 묶음 자체가 실패하면 건별로 시도해 문제 건만 가려냄
        resps = idx.map(i => { try { const f = ai_fetchOptions_('/messages', 'post', params[i]); return UrlFetchApp.fetch(f.url, f.options); } catch (e2) { return null; } });
      }
      resps.forEach((resp, k) => {
        const i = idx[k];
        if (!resp) { results[i] = { error: 'fetch_exception', raw: '응답 없음(시간 초과 가능)' }; next.push(i); return; }
        const code = resp.getResponseCode(), body = resp.getContentText();
        if (code === 200) { try { results[i] = ai_parseMessage_(JSON.parse(body)); } catch (e) { results[i] = { error: 'parse_error', raw: body.slice(0, 200) }; } }
        else if (code === 429 || code === 529 || code >= 500) {
          results[i] = { error: `api_${code}`, raw: body.slice(0, 200) }; next.push(i);
          const ra = ai_retryAfter_(resp); if (ra > 0) Utilities.sleep(Math.min(ra, 30) * 1000);
        } else results[i] = { error: `api_${code}`, raw: body.slice(0, 200) };
      });
    }
    pending = next;
  }
  return results;
}

/* =====================================================================
 * 되돌리기
 * ===================================================================== */

function ai_undoRun() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ai_loadState_()) { ui.alert('배치 작업이 진행 중입니다. 끝난 뒤 사용하세요.'); return; }
  const p = ui.prompt('실행 되돌리기', '되돌릴 실행ID를 입력하세요 (예: 261107-153000T). 로그 B열에 있습니다.', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const runId = String(p.getResponseText() || '').trim();
  const logSh = ss.getSheetByName(AI_CFG.LOG_SHEET);
  if (!runId || !logSh || logSh.getLastRow() < 2) { ui.alert('실행ID 또는 로그 시트가 없습니다.'); return; }
  const vals = logSh.getRange(2, 1, logSh.getLastRow() - 1, AI_LOG.HEADER.length).getValues();
  const targets = [];
  vals.forEach((r, k) => {
    if (String(r[AI_LOG.RUN - 1]) !== runId) return;
    const st = String(r[AI_LOG.STATUS - 1]);
    if ((st === 'ok' || st === 'review') && (r[AI_LOG.MODE - 1] === 'tag' || r[AI_LOG.MODE - 1] === 'now'))
      targets.push({ logRow: k + 2, sheetName: String(r[AI_LOG.SHEET - 1]), row: Number(r[AI_LOG.ROW - 1]), id: String(r[AI_LOG.ID - 1]),
        unit: String(r[AI_LOG.UNIT - 1]), tags: String(r[AI_LOG.TAGS - 1]), prevM: String(r[AI_LOG.PREV_M - 1] || ''),
        note: String(r[AI_LOG.NOTE - 1] || '') });
  });
  if (!targets.length) { ui.alert(`실행 ${runId} 에 되돌릴(ok/review) 행이 없습니다.`); return; }
  if (ui.alert('실행 되돌리기', `실행 ${runId}: ${targets.length}행.\nAI가 쓴 값이 지금도 그대로인 행만 L을 비우고 M을 이전 값으로 되돌립니다.\n(기록 때 ID를 찾아 다른 행에 썼으면 그 행을 봅니다)\n진행할까요?`, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  let undone = 0, kept = 0;
  const statusUpdates = {};
  const bySheet = {};
  targets.forEach(t => {
    const mv = t.note.match(/행 이동 \d+→(\d+)/);                 // 기록 때 ID를 찾아 옮겨 쓴 경우
    if (mv) t.row = Number(mv[1]);
    (bySheet[t.sheetName] = bySheet[t.sheetName] || []).push(t);
  });
  Object.keys(bySheet).forEach(name => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    const list = bySheet[name].sort((a, b) => a.row - b.row);
    const writes = {};
    ai_clusters_(list.map(t => t.row), 1).forEach(cl => {
      const cur = sh.getRange(cl.start, 1, cl.len, AI_CFG.COL.TAG).getDisplayValues();
      list.filter(t => t.row >= cl.start && t.row < cl.start + cl.len).forEach(t => {
        const c = cur[t.row - cl.start];
        const same = ai_normId_(c[AI_CFG.COL.ID - 1]) === t.id && String(c[AI_CFG.COL.UNIT - 1]).trim() === t.unit && String(c[AI_CFG.COL.TAG - 1]).trim() === t.tags;
        if (same) { writes[t.row] = ['', t.prevM]; statusUpdates[t.logRow] = 'undone'; undone++; } else kept++;
      });
    });
    ai_writeRows_(sh, AI_CFG.COL.UNIT, 2, writes);
  });
  const su = {}; Object.keys(statusUpdates).forEach(r => { su[r] = [statusUpdates[r]]; });
  ai_writeRows_(logSh, AI_LOG.STATUS, 1, su);
  ui.alert(`되돌리기 완료: ${undone}행 복원, ${kept}행은 사람이 고쳐서 그대로 둠.`);
}

/* =====================================================================
 * 컨텍스트: 대단원 표, 태그 사전, Data_DS 색인, 건너뛰기 집합
 * ===================================================================== */

function ai_loadContext_(ss, mode, force) {
  const units = ai_loadUnits_(ss);
  const dict = ai_loadDict_(ss, units);
  const ds = ai_loadDsIndex_();
  const skip = (mode === 'tag' && !force) ? ai_loadSkipSet_(ss) : {};
  return { units, dict, ds, skip, effort: AI_CFG.EFFORT, systemCache: {}, schemaCache: {} };
}

/** 수거 시 필요한 것만 (Data_DS 안 읽음) */
function ai_loadContextLite_(ss) {
  const units = ai_loadUnits_(ss);
  return { units, dict: ai_loadDict_(ss, units) };
}

/** 선택항목!E2:G → { list, set, bySubject, subjectOfUnit } */
function ai_loadUnits_(ss) {
  const sh = ss.getSheetByName(AI_CFG.SELECT_SHEET);
  const vals = sh.getRange(2, 5, 30, 3).getValues();           // E,F,G
  const out = { list: [], set: {}, bySubject: {}, subjectOfUnit: {} };
  vals.forEach(r => {
    const unit = String(r[0] || '').trim(), subj = String(r[2] || '').trim();
    if (!unit || !subj) return;
    out.list.push(unit); out.set[unit] = true;
    (out.bySubject[subj] = out.bySubject[subj] || []).push(unit);
    out.subjectOfUnit[unit] = subj;
  });
  if (out.list.length !== 15) Logger.log(`⚠️ 대단원 표가 15개가 아님: ${out.list.length}개`);
  return out;
}

/** _태그사전 → { canonOf:{normKey:canonical}, unitsOf:{canonical:{unit:true}}, listBySubject:{subj:[...]} } */
function ai_loadDict_(ss, units) {
  const sh = ss.getSheetByName(AI_CFG.DICT_SHEET);
  const out = { canonOf: {}, unitsOf: {}, listBySubject: {} };
  if (!sh || sh.getLastRow() < 2) return out;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();   // A 키워드, B 빈도, C 대단원들, D 정규형, E 사용
  vals.forEach(r => {
    const kw = ai_cleanKw_(r[0]); if (!kw) return;
    const use = (r[4] === '' || r[4] === null) ? true : (r[4] === true || String(r[4]).toUpperCase() === 'TRUE');
    if (!use) return;
    const canon = ai_cleanKw_(r[3]) || kw;
    out.canonOf[ai_norm_(kw)] = canon;
    out.canonOf[ai_norm_(canon)] = canon;
    const us = out.unitsOf[canon] = out.unitsOf[canon] || {};
    String(r[2] || '').split(',').map(x => x.trim()).filter(Boolean).forEach(u => { us[u] = true; });
  });
  const sortKo = (a, b) => a.localeCompare(b, 'ko');          // 결정적 정렬 → 캐시 접두부 고정
  Object.keys(units.bySubject).forEach(subj => {
    const allowed = units.bySubject[subj];
    out.listBySubject[subj] = Object.keys(out.unitsOf).filter(c => {
      const us = Object.keys(out.unitsOf[c]);
      return us.length === 0 || us.some(u => allowed.indexOf(u) >= 0);
    }).sort(sortKo);
  });
  out.listBySubject['*'] = Object.keys(out.unitsOf).sort(sortKo);
  return out;
}

/** Data_DS 의 A(id)·L(jangbu_id)만 읽어 색인: jangbuId → {dsRow, dsId, subjCode, dup}  (D3: 중복은 마지막 행) */
function ai_loadDsIndex_() {
  const sh = SpreadsheetApp.openById(AI_CFG.LATEX_SS_ID).getSheetByName(AI_CFG.DS_SHEET);
  const last = sh.getLastRow();
  const map = {};
  if (last < 2) return map;
  const ids = sh.getRange(2, AI_CFG.DS_COL.ID, last - 1, 1).getValues();
  const jbs = sh.getRange(2, AI_CFG.DS_COL.JANGBU, last - 1, 1).getValues();
  for (let i = 0; i < jbs.length; i++) {
    const jid = ai_normId_(jbs[i][0]); if (!jid) continue;
    const dsId = String(ids[i][0] || '').trim();
    const m = dsId.match(/_(1공통|3확통|4미적|5기하)\d{1,2}$/);
    const prev = map[jid];
    map[jid] = { dsRow: i + 2, dsId, subjCode: m ? AI_CFG.SUBJ_CODE[m[1]] : '', dup: prev ? prev.dup + 1 : 1 };
  }
  return map;
}

/** 필요한 Data_DS 행의 B(문제)·C(풀이)만 연속 구간으로 읽기: dsRow → {problem, solution} */
function ai_loadDsText_(dsRows) {
  const sh = SpreadsheetApp.openById(AI_CFG.LATEX_SS_ID).getSheetByName(AI_CFG.DS_SHEET);
  const out = {};
  ai_clusters_(dsRows, 100).forEach(cl => {
    const vals = sh.getRange(cl.start, AI_CFG.DS_COL.PROBLEM, cl.len, 2).getValues();
    vals.forEach((r, k) => { out[cl.start + k] = { problem: String(r[0] || ''), solution: String(r[1] || '') }; });
  });
  return out;
}

/** 로그에서 영구 보류 상태의 ID 집합 (tag 모드, 강제 아님일 때) */
function ai_loadSkipSet_(ss) {
  const sh = ss.getSheetByName(AI_CFG.LOG_SHEET);
  const set = {};
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < AI_LOG.HEADER.length) return set;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, AI_LOG.STATUS).getValues();
  vals.forEach(r => {
    if (String(r[AI_LOG.MODE - 1]) !== 'tag') return;
    const id = ai_normId_(r[AI_LOG.ID - 1]); if (!id) return;
    const st = String(r[AI_LOG.STATUS - 1]);
    if (AI_CFG.SKIP_STATUSES.indexOf(st) >= 0) set[id] = st;
    else if (st === 'ok' || st === 'review' || st === 'undone') delete set[id];   // 그 뒤에 성공/되돌림이 있으면 해제
  });
  return set;
}

/* =====================================================================
 * 대상 행 수집
 * ===================================================================== */

/** 한 행(표시값 A~M)으로 item 만들기. ID가 아니면 null */
function ai_buildItem_(sheetName, row, r, ctx, mode) {
  const id = ai_normId_(r[AI_CFG.COL.ID - 1]);
  if (!AI_CFG.ID_RE.test(id)) return null;
  const fRaw = String(r[AI_CFG.COL.SUBJ - 1] || '').trim();
  const it = {
    sheetName, row, id,
    year: parseInt(String(r[AI_CFG.COL.YEAR - 1] || '').trim(), 10),
    setVal: String(r[AI_CFG.COL.SET - 1] || '').trim(),
    fVal: fRaw.startsWith('#') ? '' : fRaw,                       // #N/A 등은 빈칸 취급
    gNo: parseInt(String(r[AI_CFG.COL.NO - 1] || '').trim(), 10),
    unitHuman: String(r[AI_CFG.COL.UNIT - 1] || '').trim(),
    tagHuman: String(r[AI_CFG.COL.TAG - 1] || '').trim(),
  };
  it.prevM = it.tagHuman;
  const src = ctx.ds[id] || null;
  it.dsRow = src ? src.dsRow : 0; it.dsCode = src ? src.subjCode : ''; it.dup = src ? src.dup : 0;
  const s = ai_resolveSubject_(ctx, it, mode);
  it.subject = s.subject; it.subjPath = s.path; it.subjMismatch = s.mismatch;
  return it;
}

/** 시트의 from~to 행에서 후보 수집 → run.items / run.hold / run.noSource / run.skipped */
function ai_collectCandidates_(sh, from, to, prefixes, ctx, run) {
  const last = sh.getLastRow();
  const endRow = to ? Math.min(to, last) : last;
  if (from > endRow) return;
  const vals = sh.getRange(from, 1, endRow - from + 1, AI_CFG.COL.TAG).getDisplayValues();
  vals.forEach((r, k) => {
    const it = ai_buildItem_(sh.getName(), from + k, r, ctx, 'tag');
    if (!it || prefixes.indexOf(it.id[0]) < 0) return;
    if (it.unitHuman) return;                                     // D2: L 빈칸만
    if (!it.setVal || !Number.isFinite(it.gNo)) return;
    if (Number.isFinite(it.year) && it.year > AI_CFG.YEAR_MAX) return;
    if (ctx.skip[it.id]) { run.skipped++; return; }
    if (!it.dsRow) { run.noSource++; return; }                    // C6: 로그 없이 집계만 — 변환되면 다음 실행에 자동 포함
    if (it.subjMismatch) { run.hold.push(it); return; }           // D8
    run.items.push(it);
  });
}

/**
 * 과목 결정: ① F가 값이면 F ② Data_DS id 과목코드 ③ G≤22 → '1 공통'. 경로끼리 다르면 mismatch.
 * 파일럿은 사람 L의 과목. 즉시 판정은 사람 L이 있어도 실제 경로로(스모크 테스트 목적).
 */
function ai_resolveSubject_(ctx, it, mode) {
  if (mode === 'pilot' && it.unitHuman && ctx.units.subjectOfUnit[it.unitHuman])
    return { subject: ctx.units.subjectOfUnit[it.unitHuman], path: 'L', mismatch: false };
  const subjects = Object.keys(ctx.units.bySubject);
  const votes = [];
  if (it.fVal && subjects.indexOf(it.fVal) >= 0) votes.push({ s: it.fVal, p: 'F' });
  if (it.dsCode) votes.push({ s: it.dsCode, p: 'DS' });
  if (Number.isFinite(it.gNo) && it.gNo >= 1 && it.gNo <= 22) votes.push({ s: '1 공통', p: 'G' });
  if (!votes.length) return { subject: '', path: 'none', mismatch: false };
  const first = votes[0];
  return { subject: first.s, path: votes.map(v => v.p).join('+'), mismatch: votes.some(v => v.s !== first.s) };
}

/** 파일럿 표본: 과목별 N행, 시트 전체에 고르게 (사람 L·M 있고 Data_DS 에 있는 D행) */
function ai_pickPilotRows_(ss, ctx) {
  const bySubj = {};
  AI_CFG.TARGET_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    const last = sh.getLastRow(); if (last < 2) return;
    const vals = sh.getRange(2, 1, last - 1, AI_CFG.COL.TAG).getDisplayValues();
    vals.forEach((r, k) => {
      const it = ai_buildItem_(name, k + 2, r, ctx, 'pilot');
      if (!it || it.id[0] !== 'D') return;
      if (!it.unitHuman || !it.tagHuman || !ctx.units.set[it.unitHuman] || !it.dsRow) return;
      (bySubj[it.subject] = bySubj[it.subject] || []).push(it);
    });
  });
  const picked = [];
  Object.keys(bySubj).sort().forEach(subj => {
    const arr = bySubj[subj], N = Math.min(AI_CFG.PILOT_PER_SUBJECT, arr.length);
    for (let j = 0; j < N; j++) picked.push(arr[Math.floor(j * arr.length / N)]);
  });
  return picked.sort((a, b) => a.sheetName === b.sheetName ? a.row - b.row : a.sheetName.localeCompare(b.sheetName));
}

/* =====================================================================
 * 프롬프트·요청 본문
 * ===================================================================== */

/** 과목별 system 블록 (규칙+허용 대단원, 사전). 사전 블록에 1시간 캐시 */
function ai_systemBlocks_(ctx, subject) {
  const key = subject || '*';
  if (ctx.systemCache[key]) return ctx.systemCache[key];
  const allowed = subject ? ctx.units.bySubject[subject] : ctx.units.list;
  const kws = ctx.dict.listBySubject[key] || ctx.dict.listBySubject['*'] || [];
  const rules =
`당신은 한국 수능·모의고사 수학 문항을 교육과정 대단원과 핵심 개념 태그로 분류하는 전문가다.
입력으로 문제와 그 풀이(LaTeX 섞인 한국어)가 주어진다.

과제
(1) 이 문항이 속한 대단원 1개
(2) 풀이에 실제로 사용된 핵심 개념 키워드, 중요한 순서로 최대 ${AI_CFG.MAX_KEYWORDS}개

규칙
1. 대단원은 [허용 대단원] 목록의 문자열 중 하나를 숫자·띄어쓰기까지 글자 그대로 쓴다.
2. 판단 기준은 풀이에서 실제로 쓰인 개념이다. 문제의 소재나 겉모습보다 풀이 과정을 우선한다.
3. 둘 이상의 대단원 개념이 함께 쓰였으면 units_used 에 모두 적고, unit 에는 앞 두 자리 번호가 가장 큰 것을 적는다. (예: 13 수열과 22 미분이 함께 쓰이면 unit 은 "22 미분")
4. 키워드는 [태그 사전]에 있는 문자열만 글자 그대로 쓴다. 사전에 없는 말을 새로 만들지 않는다. 적절한 것이 없으면 빈 배열로 둔다. 풀이의 핵심 발상을 담은 키워드를 앞에 둔다.
5. 풀이에 [그림] 표시가 있으면 그 자리에 있던 그림은 볼 수 없다. 본문 서술만으로 판단하고, 그림 없이 판단이 어려우면 confidence 를 낮춘다.
6. confidence 는 0~1 사이 실수. 대단원이 애매하거나 키워드가 사전과 잘 맞지 않으면 낮게 준다. reason 은 판정 근거 한 문장.
7. 경계 주의: 확률변수·확률분포(이산/연속, 이항분포, 정규분포)·통계적 추정은 "33 통계"다. "32 확률"은 사건의 확률(덧셈·곱셈정리, 조건부확률, 독립시행)까지다.

[허용 대단원]
${allowed.map(u => '- ' + u).join('\n')}`;
  const dict = `[태그 사전] (한 줄에 하나, 이 중에서만 고른다)\n${kws.join('\n')}`;
  const blocks = [{ type: 'text', text: rules }, { type: 'text', text: dict, cache_control: { type: 'ephemeral', ttl: AI_CFG.CACHE_TTL } }];
  ctx.systemCache[key] = blocks;
  return blocks;
}

/** 과목별 출력 스키마 (대단원 enum). 수치 제약은 구조화 출력이 지원하지 않아 코드에서 검사 */
function ai_schema_(ctx, subject) {
  const key = subject || '*';
  if (ctx.schemaCache[key]) return ctx.schemaCache[key];
  const allowed = subject ? ctx.units.bySubject[subject] : ctx.units.list;
  const schema = {
    type: 'object', additionalProperties: false,
    required: ['unit', 'units_used', 'keywords', 'confidence', 'reason'],
    properties: {
      unit: { type: 'string', enum: allowed },
      units_used: { type: 'array', items: { type: 'string', enum: allowed } },
      keywords: { type: 'array', items: { type: 'string' } },
      confidence: { type: 'number' },
      reason: { type: 'string' },
    },
  };
  ctx.schemaCache[key] = schema;
  return schema;
}

/** user 메시지: 그림 마크다운 → [그림], 길이 제한. {text, flags, inChars} */
function ai_userText_(it, src) {
  const figRe = /!\[[^\]]*\]\([^)]*\)/g, figTest = /!\[[^\]]*\]\([^)]*\)/;
  const p0 = String(src.problem || ''), s0 = String(src.solution || '');
  const hasFigure = figTest.test(p0) || figTest.test(s0);
  const p1 = p0.replace(figRe, '[그림]'), s1 = s0.replace(figRe, '[그림]');
  const p = ai_truncate_(p1, AI_CFG.MAX_PROBLEM_CHARS), s = ai_truncate_(s1, AI_CFG.MAX_SOLUTION_CHARS);
  const flags = [hasFigure ? '그림' : '', (p.truncated || s.truncated) ? '잘림' : '', it.dup > 1 ? `중복${it.dup}` : ''].filter(Boolean).join('/');
  const text = `[문항] ${it.id}  (과목: ${it.subject || '미정'}, 번호: ${Number.isFinite(it.gNo) ? it.gNo : '?'})\n\n[문제]\n${p.text}\n\n[풀이]\n${s.text}`;
  return { text, flags, inChars: p0.length + s0.length };
}

/** Messages 요청 본문 (동기·배치 공용). temperature 등 샘플링 파라미터는 보내지 않는다. effort 는 ctx.effort */
function ai_params_(ctx, subject, userText) {
  return {
    model: AI_CFG.MODEL,
    max_tokens: AI_CFG.MAX_TOKENS,
    output_config: { effort: ctx.effort || AI_CFG.EFFORT, format: { type: 'json_schema', schema: ai_schema_(ctx, subject) } },
    system: ai_systemBlocks_(ctx, subject),
    messages: [{ role: 'user', content: userText }],
  };
}

function ai_customId_(it) { return `${AI_CFG.SHEET_CODE[it.sheetName]}${it.row}_${it.id}`; }

/* =====================================================================
 * 응답 해석·검증
 * ===================================================================== */

/** Messages 응답 객체 → {unit, unitsUsed, keywords, confidence, reason, tokens} 또는 {error, raw, tokens} */
function ai_parseMessage_(msg) {
  const u = (msg && msg.usage) || {};
  const tokens = `${u.input_tokens || 0}/${u.cache_read_input_tokens || 0}/${u.cache_creation_input_tokens || 0}/${u.output_tokens || 0}`;
  const stop = msg && msg.stop_reason;
  if (stop === 'refusal') return { error: 'refused', raw: (msg.stop_details && (msg.stop_details.category || msg.stop_details.reason)) || '', tokens };
  if (stop === 'max_tokens') return { error: 'max_tokens', raw: 'thinking+출력이 max_tokens 초과', tokens };
  const text = ((msg && msg.content) || []).filter(c => c.type === 'text').map(c => c.text).join('\n').trim();
  if (!text) return { error: 'parse_error', raw: `stop=${stop}, 텍스트 없음`, tokens };
  try {
    const o = JSON.parse(text);
    return {
      unit: String(o.unit || '').trim(),
      unitsUsed: Array.isArray(o.units_used) ? o.units_used.map(String) : [],
      keywords: Array.isArray(o.keywords) ? o.keywords.map(String) : [],
      confidence: Number(o.confidence), reason: String(o.reason || '').slice(0, 300), tokens,
    };
  } catch (e) { return { error: 'parse_error', raw: text.slice(0, 200), tokens }; }
}

/**
 * 검증·확정. human = {unit, tags} (파일럿·즉시 판정에서 비교용)
 * 반환 {status, unit, tags, dropped, multi, conf, reason, tokens, note, write:[unit,tags]|null, match}
 */
function ai_finalize_(ctx, mode, subject, res, human) {
  const out = { status: 'api_error', unit: '', tags: '', dropped: '', multi: '', conf: '', reason: '', tokens: (res && res.tokens) || '', note: '', write: null, match: null };
  if (!res || res.error) {
    const e = res ? res.error : 'no_result';
    out.status = (e === 'refused' || e === 'max_tokens' || e === 'parse_error') ? e : 'api_error';
    out.note = `${e}${res && res.raw ? ': ' + String(res.raw).slice(0, 200) : ''}`;
    return out;
  }
  out.conf = Number.isFinite(res.confidence) ? Math.round(Math.min(1, Math.max(0, res.confidence)) * 100) / 100 : '';
  out.reason = res.reason;

  // 대단원: 허용 목록 안에서, 복수면 번호 큰 것 (D4). enum 이라 보통 통과하지만 방어
  const allowed = subject ? ctx.units.bySubject[subject] : ctx.units.list;
  const fix = u => { const n = ai_norm_(u); return ctx.units.list.find(x => ai_norm_(x) === n) || ''; };
  const cands = [];
  res.unitsUsed.map(fix).forEach(u => { if (u && allowed.indexOf(u) >= 0 && cands.indexOf(u) < 0) cands.push(u); });
  const single = fix(res.unit);
  if (single && allowed.indexOf(single) >= 0 && cands.indexOf(single) < 0) cands.push(single);
  if (!cands.length) {
    out.status = 'unit_invalid';
    out.note = `AI unit="${res.unit}" units_used=${JSON.stringify(res.unitsUsed)} 허용=${subject || '전체'}`;
    return out;
  }
  const unit = cands.reduce((a, b) => (parseInt(b, 10) > parseInt(a, 10) ? b : a));
  out.unit = unit;
  out.multi = cands.length > 1 ? cands.join(' | ') : '';
  const notes = [];
  if (single && single !== unit) notes.push(`AI unit ${single}→규칙 ${unit}`);

  // 키워드: 사전 정규화 (D9)
  const kept = [], dropped = [];
  res.keywords.forEach(k => {
    const c = ctx.dict.canonOf[ai_norm_(k)];
    if (c) { if (kept.indexOf(c) < 0) kept.push(c); } else { const ck = ai_cleanKw_(k); if (ck) dropped.push(ck); }
  });
  const tags = kept.slice(0, AI_CFG.MAX_KEYWORDS);
  out.tags = tags.join(', ');
  out.dropped = dropped.join(', ');

  // 사람 값과 비교 (파일럿·즉시 판정)
  const hUnit = String((human && human.unit) || '').trim(), hTags = String((human && human.tags) || '').trim();
  if (hUnit || hTags) {
    const hk = hTags.split(',').map(ai_cleanKw_).filter(Boolean).map(ai_norm_);
    out.match = { unit: hUnit === unit ? 1 : 0, overlap: tags.filter(t => hk.indexOf(ai_norm_(t)) >= 0).length, humanN: hk.length };
  }
  if (mode === 'pilot') { out.status = 'pilot'; out.note = notes.join('; '); return out; }

  // 본 실행·즉시 판정: 기록 (D7)
  const low = out.conf !== '' && out.conf < AI_CFG.REVIEW_BELOW;
  const review = low || cands.length > 1 || !subject || !tags.length || dropped.length > 0;
  out.write = [unit, out.tags];
  out.status = review ? 'review' : 'ok';
  if (low) notes.push(`conf<${AI_CFG.REVIEW_BELOW}`);
  if (cands.length > 1) notes.push('복수 대단원');
  if (!subject) notes.push('과목 미정');
  if (!tags.length) notes.push('키워드 없음');
  if (dropped.length) notes.push('버린 키워드 있음');
  out.note = notes.join('; ');
  return out;
}

/**
 * L·M 기록. pending = { row → {id, values:[unit,tags], logRow} }
 * 기록 직전 그 행의 A·L·M 을 다시 읽어 (B6)
 *   - ID가 다르면 A열에서 그 ID를 찾아 그 행에 기록(N11) — 못 찾으면 id_moved
 *   - L이 채워졌으면 write_skipped 로 두고 쓰지 않는다
 * strict 검증 예외 등 실패는 write_error.
 */
function ai_flushWrites_(sh, pending) {
  const rows = Object.keys(pending).map(Number).sort((a, b) => a - b);
  if (!rows.length) return;
  const writes = {};
  const byRow = {};                    // 실제 기록 행 → pending 항목
  const idCache = {};                  // ID 찾기용 A열 캐시 (필요할 때 1회)
  const relocated = [];                // {p, from, to}
  ai_clusters_(rows, 1).forEach(cl => {
    const cur = sh.getRange(cl.start, 1, cl.len, AI_CFG.COL.TAG).getDisplayValues();
    for (let r = cl.start; r < cl.start + cl.len; r++) {
      const p = pending[r]; if (!p) continue;
      const c = cur[r - cl.start];
      const lr = p.logRow;
      if (ai_normId_(c[AI_CFG.COL.ID - 1]) !== p.id) {
        const f = ai_findRowById_(sh, p.id, idCache);
        if (!f.row || f.dup) { lr[AI_LOG.STATUS - 1] = 'id_moved'; lr[AI_LOG.NOTE - 1] = ai_joinNote_(lr[AI_LOG.NOTE - 1], `그 행의 ID가 "${c[AI_CFG.COL.ID - 1]}"${f.dup ? `, ${p.id} 가 시트에 ${f.dup}번 있음` : `, 시트에서 ${p.id} 를 찾지 못함`}`); continue; }
        relocated.push({ p, from: r, to: f.row }); continue;
      }
      if (String(c[AI_CFG.COL.UNIT - 1] || '').trim()) { lr[AI_LOG.STATUS - 1] = 'write_skipped'; lr[AI_LOG.NOTE - 1] = ai_joinNote_(lr[AI_LOG.NOTE - 1], `사람이 먼저 적음: L="${c[AI_CFG.COL.UNIT - 1]}"`); continue; }
      lr[AI_LOG.PREV_M - 1] = String(c[AI_CFG.COL.TAG - 1] || '');
      writes[r] = p.values; byRow[r] = p;
    }
  });
  // ID를 찾아 옮겨 쓰는 행: 그 행의 L·M 다시 확인
  relocated.forEach(x => {
    const c = sh.getRange(x.to, 1, 1, AI_CFG.COL.TAG).getDisplayValues()[0];
    const lr = x.p.logRow;
    if (String(c[AI_CFG.COL.UNIT - 1] || '').trim()) { lr[AI_LOG.STATUS - 1] = 'write_skipped'; lr[AI_LOG.NOTE - 1] = ai_joinNote_(lr[AI_LOG.NOTE - 1], `행 이동 ${x.from}→${x.to}; 사람이 먼저 적음: L="${c[AI_CFG.COL.UNIT - 1]}"`); return; }
    lr[AI_LOG.PREV_M - 1] = String(c[AI_CFG.COL.TAG - 1] || '');
    lr[AI_LOG.NOTE - 1] = ai_joinNote_(lr[AI_LOG.NOTE - 1], `행 이동 ${x.from}→${x.to}`);
    writes[x.to] = x.p.values; byRow[x.to] = x.p;
  });
  ai_writeRows_(sh, AI_CFG.COL.UNIT, 2, writes, (r, err) => {
    const lr = byRow[r].logRow;
    lr[AI_LOG.STATUS - 1] = 'write_error'; lr[AI_LOG.NOTE - 1] = String(err).slice(0, 200);
  });
}

/** rowValues = { row → [v1, v2, …] } 를 연속 구간으로 묶어 col 부터 width 열에 기록. 실패한 구간은 행 단위로 재시도 */
function ai_writeRows_(sh, col, width, rowValues, onError) {
  const rows = Object.keys(rowValues).map(Number).sort((a, b) => a - b);
  ai_clusters_(rows, 1).forEach(cl => {
    const values = []; for (let r = cl.start; r < cl.start + cl.len; r++) values.push(rowValues[r]);
    try { sh.getRange(cl.start, col, cl.len, width).setValues(values); }
    catch (e) {
      for (let r = cl.start; r < cl.start + cl.len; r++) {
        try { sh.getRange(r, col, 1, width).setValues([rowValues[r]]); }
        catch (e2) { if (onError) onError(r, e2); else Logger.log(`기록 실패 ${r}행: ${e2}`); }
      }
    }
  });
}

/* =====================================================================
 * 로그 시트
 * ===================================================================== */

function ai_getLogSheet_(ss) {
  let sh = ss.getSheetByName(AI_CFG.LOG_SHEET);
  if (sh && sh.getLastColumn() < AI_LOG.HEADER.length) {              // 열 수가 다른 옛 로그 → 이름 바꿔 보존 (N7)
    let name = AI_CFG.LOG_SHEET_V1, k = 2;
    while (ss.getSheetByName(name)) name = `${AI_CFG.LOG_SHEET_V1}_${k++}`;
    sh.setName(name); sh = null;
  }
  if (!sh) {
    sh = ss.insertSheet(AI_CFG.LOG_SHEET);
    sh.getRange(1, 1, 1, AI_LOG.HEADER.length).setValues([AI_LOG.HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // 전체 텍스트 서식: "1/2"(키워드 겹침)·"10/500/0/40"(토큰) 같은 값이 날짜·수식으로 바뀌지 않게.
  // 기록이 시트 끝을 넘어 자동 추가되는 행은 서식이 없으므로, 미리 여유 행을 두고 서식을 건다
  const need = sh.getLastRow() + 5000;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(1, 1, sh.getMaxRows(), AI_LOG.HEADER.length).setNumberFormat('@');
  return sh;
}

function ai_logRow_(runId, mode, it, status, note, effort) {
  const r = new Array(AI_LOG.HEADER.length).fill('');
  r[AI_LOG.TIME - 1] = ai_now_(); r[AI_LOG.RUN - 1] = runId; r[AI_LOG.MODE - 1] = mode;
  r[AI_LOG.SHEET - 1] = it.sheetName; r[AI_LOG.ROW - 1] = it.row; r[AI_LOG.ID - 1] = it.id; r[AI_LOG.STATUS - 1] = status;
  r[AI_LOG.SUBJ - 1] = it.subject || ''; r[AI_LOG.PATH - 1] = it.subjPath || '';
  r[AI_LOG.MODEL - 1] = `${AI_CFG.MODEL}/${effort || AI_CFG.EFFORT}`;
  r[AI_LOG.INCHARS - 1] = it.inChars === undefined ? '' : it.inChars; r[AI_LOG.FLAGS - 1] = it.flags || '';
  r[AI_LOG.PREV_M - 1] = it.prevM || ''; r[AI_LOG.NOTE - 1] = note || '';
  if (mode !== 'tag') { r[AI_LOG.H_UNIT - 1] = it.unitHuman || ''; r[AI_LOG.H_TAGS - 1] = it.tagHuman || ''; }
  return r;
}

/** 실행의 로그 구간(연속 행)을 읽어 {sh, start, rows, index}. 구간이 깨졌으면(사이에 다른 행) throw */
function ai_readLogBlock_(logSh, state) {
  const rows = logSh.getRange(state.logStart, 1, state.logCount, AI_LOG.HEADER.length).getValues();
  const bad = rows.findIndex(r => String(r[AI_LOG.RUN - 1]) !== state.runId);
  if (bad >= 0) throw new Error(`로그 구간이 어긋남: ${state.logStart + bad}행의 실행ID가 "${rows[bad][AI_LOG.RUN - 1]}" (기대 ${state.runId}). 로그 시트에 행을 넣거나 지우지 마세요.`);
  const index = {};
  rows.forEach((r, i) => { index[`${r[AI_LOG.SHEET - 1]}|${Number(r[AI_LOG.ROW - 1])}|${r[AI_LOG.ID - 1]}`] = i; });
  return { sh: logSh, start: state.logStart, rows, index };
}
function ai_writeLogBlock_(block) { block.sh.getRange(block.start, 1, block.rows.length, AI_LOG.HEADER.length).setValues(block.rows); }

/** 완료 요약 (파일럿은 일치율 집계). 로그 끝에 summary 행 */
function ai_finishSummary_(ss, state, block) {
  const cnt = {};
  block.rows.forEach(r => { const s = r[AI_LOG.STATUS - 1]; cnt[s] = (cnt[s] || 0) + 1; });
  let s = `[${state.runId}] 모드 ${state.mode}, ${state.model}/${state.effort}\n대상 ${state.total} · Data_DS 없음 ${state.noSource} · 건너뜀 ${state.skipped}\n` +
    `결과: ${Object.keys(cnt).sort().map(k => `${k} ${cnt[k]}`).join(', ')}`;
  // 토큰 합계 → 대략 비용 (배치 50%)
  let tin = 0, tcr = 0, tcw = 0, tout = 0;
  block.rows.forEach(r => { const m = String(r[AI_LOG.TOKENS - 1]).match(/^(\d+)\/(\d+)\/(\d+)\/(\d+)$/); if (m) { tin += +m[1]; tcr += +m[2]; tcw += +m[3]; tout += +m[4]; } });
  const half = state.mode === 'now' ? 1 : 0.5;
  const usd = (tin * 4 + tcr * 0.2 + tcw * 8 + tout * 20) / 1e6 * half;
  s += `\n토큰 in ${tin} / cacheR ${tcr} / cacheW ${tcw} / out ${tout} ≈ $${usd.toFixed(2)}`;
  if (state.cacheWarn) s += `\n⚠️ ${state.cacheWarn}`;
  if (state.mode === 'pilot') {
    const by = {};
    block.rows.filter(r => r[AI_LOG.STATUS - 1] === 'pilot').forEach(r => {
      const subj = String(r[AI_LOG.SUBJ - 1] || '?');
      const b = by[subj] = by[subj] || { n: 0, unitOk: 0, ovl: 0, any: 0, fig: 0, figOk: 0 };
      b.n++; b.unitOk += Number(r[AI_LOG.M_UNIT - 1]) || 0;
      const mm = String(r[AI_LOG.M_OVL - 1]).match(/^(\d+)\/(\d+)$/);
      if (mm) { const o = +mm[1], h = +mm[2]; b.ovl += h ? o / h : 0; b.any += o > 0 ? 1 : 0; }
      if (/그림/.test(String(r[AI_LOG.FLAGS - 1]))) { b.fig++; b.figOk += Number(r[AI_LOG.M_UNIT - 1]) || 0; }
    });
    s += '\n파일럿 일치율 (과목: 표본 · 대단원 정확 · 키워드 겹침률 평균 · 1개 이상 겹침 · 그림 문항 대단원 정확)';
    const tot = { n: 0, unitOk: 0, ovl: 0, any: 0, fig: 0, figOk: 0 };
    Object.keys(by).sort().forEach(k => {
      const b = by[k]; Object.keys(tot).forEach(f => { tot[f] += b[f]; });
      s += `\n${k}: ${b.n} · ${(100 * b.unitOk / b.n).toFixed(0)}% · ${(100 * b.ovl / b.n).toFixed(0)}% · ${(100 * b.any / b.n).toFixed(0)}% · ${b.fig ? `${(100 * b.figOk / b.fig).toFixed(0)}% (${b.fig})` : '-'}`;
    });
    if (tot.n) s += `\n전체: ${tot.n} · ${(100 * tot.unitOk / tot.n).toFixed(0)}% · ${(100 * tot.ovl / tot.n).toFixed(0)}% · ${(100 * tot.any / tot.n).toFixed(0)}% · ${tot.fig ? `${(100 * tot.figOk / tot.fig).toFixed(0)}% (${tot.fig})` : '-'}`;
  }
  const row = new Array(AI_LOG.HEADER.length).fill('');
  row[AI_LOG.TIME - 1] = ai_now_(); row[AI_LOG.RUN - 1] = state.runId; row[AI_LOG.MODE - 1] = state.mode; row[AI_LOG.STATUS - 1] = 'summary';
  row[AI_LOG.NOTE - 1] = s.replace(/\n/g, ' | ');
  ai_getLogSheet_(ss).appendRow(row);
  PropertiesService.getScriptProperties().setProperty(AI_CFG.PROP_LAST, s);
  return s;
}

/* =====================================================================
 * 태그 사전 만들기/갱신  (_태그사전: A 키워드, B 빈도, C 함께 쓰인 대단원, D 정규형, E 사용)
 * ===================================================================== */

function ai_buildTagDict() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const data = ss.getSheetByName(AI_CFG.DATA_SHEET);
  if (!data) { ui.alert(`'${AI_CFG.DATA_SHEET}' 시트를 찾을 수 없습니다.`); return; }

  const last = data.getLastRow();
  const vals = data.getRange(2, AI_CFG.COL.UNIT, last - 1, 2).getValues();   // L, M (읽기만)
  const agg = {};
  vals.forEach(r => {
    const unit = String(r[0] || '').trim();
    String(r[1] || '').split(',').map(ai_cleanKw_).filter(Boolean).forEach(kw => {
      const a = agg[kw] = agg[kw] || { n: 0, units: {} };
      a.n++; if (unit) a.units[unit] = true;
    });
  });

  // 기존 사전의 D(정규형)·E(사용) 보존
  let sh = ss.getSheetByName(AI_CFG.DICT_SHEET);
  const keep = {};
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
      const kw = ai_cleanKw_(r[0]); if (kw) keep[kw] = { canon: ai_cleanKw_(r[3]), use: r[4] };
    });
  } else if (!sh) {
    sh = ss.insertSheet(AI_CFG.DICT_SHEET);
  }

  // N8: 새 키워드는 빈도 ≥ 2 일 때만 기본 사용
  const rows = Object.keys(agg).map(kw => [kw, agg[kw].n, Object.keys(agg[kw].units).sort().join(', '),
    keep[kw] ? keep[kw].canon : '',
    keep[kw] && keep[kw].use !== '' && keep[kw].use !== null ? (keep[kw].use === true || String(keep[kw].use).toUpperCase() === 'TRUE') : agg[kw].n >= AI_CFG.DICT_DEFAULT_ON_MIN_FREQ]);
  rows.sort((a, b) => {
    const ka = ai_norm_(a[0]), kb = ai_norm_(b[0]);
    return ka < kb ? -1 : ka > kb ? 1 : a[0].localeCompare(b[0], 'ko');
  });

  sh.clear();
  sh.getRange(1, 1, 1, 5).setValues([['키워드', '빈도', '함께 쓰인 대단원', '정규형(비우면 그대로)', '사용']]).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 4).setNumberFormat('@');              // 키워드·정규형이 수식/날짜로 바뀌지 않게
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
    sh.getRange(2, 5, rows.length, 1).insertCheckboxes();
  }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);

  const groups = {};
  rows.forEach(r => { const k = ai_norm_(r[0]); (groups[k] = groups[k] || []).push(r[0]); });
  const dupGroups = Object.values(groups).filter(g => g.length > 1).length;
  const off = rows.filter(r => r[4] === false).length;
  ui.alert(`태그 사전 갱신 완료: 원자 키워드 ${rows.length}개 (사용 ${rows.length - off} / 사용 안 함 ${off})\n` +
    `표기만 다른 묶음 ${dupGroups}개가 인접해 있습니다.\nD열(정규형)에 대표 표기를 적고, E열 체크로 사용 여부를 정하세요.\n` +
    `빈도 1 키워드는 기본 '사용 안 함'입니다. (기존 D·E 입력은 유지했습니다)`);
}

/* =====================================================================
 * API 호출
 * ===================================================================== */

function ai_apiKey_() { return PropertiesService.getUserProperties().getProperty(AI_CFG.API_KEY_PROP) || ''; }

/** {url, options} — fetch(url, options) 또는 fetchAll 용 Object.assign({url}, options) */
function ai_fetchOptions_(path, method, payload) {
  const options = {
    method, muteHttpExceptions: true,
    headers: { 'x-api-key': ai_apiKey_(), 'anthropic-version': AI_CFG.API_VERSION },
  };
  if (payload !== null && payload !== undefined) { options.contentType = 'application/json'; options.payload = JSON.stringify(payload); }
  return { url: /^https?:/.test(path) ? path : AI_CFG.API_BASE + path, options };
}
function ai_fetchAllReq_(path, method, payload) { const f = ai_fetchOptions_(path, method, payload); return Object.assign({ url: f.url }, f.options); }

/** retry-after 헤더(초). 헤더 이름 대소문자 무시 */
function ai_retryAfter_(resp) {
  const h = resp.getHeaders() || {};
  const k = Object.keys(h).find(x => x.toLowerCase() === 'retry-after');
  const v = k ? Number(h[k]) : NaN;
  return Number.isFinite(v) && v > 0 ? v : 0;
}

/**
 * 단건 호출 + 재시도(429/529/5xx·네트워크, retry-after 우선).
 * 반환 {code, body, json}. opts.raw 면 json 해석 안 함 (JSONL 결과 다운로드).
 * opts.createOnly(배치 생성): 요청이 서버에 닿기 전 실패(네트워크 예외·429·529)만 재시도, 5xx 는 재시도 안 함 (중복 생성 방지, N13)
 */
function ai_apiFetch_(path, method, payload, opts) {
  const f = ai_fetchOptions_(path, method, payload);
  const url = f.url, o = f.options;
  let last = { code: 0, body: '', json: null };
  for (let attempt = 0; attempt <= AI_CFG.RETRY_WAITS_MS.length; attempt++) {
    if (attempt > 0) Utilities.sleep(AI_CFG.RETRY_WAITS_MS[attempt - 1]);
    let resp;
    try { resp = UrlFetchApp.fetch(url, o); } catch (e) { last = { code: 0, body: String(e), json: null }; continue; }
    const code = resp.getResponseCode(), body = resp.getContentText();
    last = { code, body, json: null };
    if (code >= 200 && code < 300) {
      if (!(opts && opts.raw)) { try { last.json = JSON.parse(body); } catch (e) { last.json = null; } }
      return last;
    }
    const retryable = (code === 429 || code === 529) || (code >= 500 && !(opts && opts.createOnly));
    if (retryable) {
      const ra = ai_retryAfter_(resp);
      if (ra > 0) Utilities.sleep(Math.min(ra, 60) * 1000);
      continue;
    }
    return last;   // 4xx(및 createOnly 의 5xx): 재시도 안 함
  }
  return last;
}

/* =====================================================================
 * 상태·임대·트리거·유틸
 * ===================================================================== */

function ai_loadState_() {
  const s = PropertiesService.getScriptProperties().getProperty(AI_CFG.PROP_STATE);
  return s ? JSON.parse(s) : null;
}
function ai_saveState_(st) { PropertiesService.getScriptProperties().setProperty(AI_CFG.PROP_STATE, JSON.stringify(st)); }
/** 중지 메뉴가 state 를 지웠는지 확인 (B3) */
function ai_stateAlive_(runId) { const s = ai_loadState_(); return !!(s && s.runId === runId); }

/** 임대: 프로젝트 공용 ScriptLock 은 몇 ms 만 잡고, 실제 실행 중 표시는 ScriptProperties 만료시각으로 (C5) */
function ai_acquireLease_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return false;
  try {
    const p = PropertiesService.getScriptProperties();
    const until = Number(p.getProperty(AI_CFG.PROP_LEASE) || 0);
    if (until > Date.now()) return false;
    p.setProperty(AI_CFG.PROP_LEASE, String(Date.now() + AI_CFG.LEASE_MS));
    return true;
  } finally { lock.releaseLock(); }
}
function ai_releaseLease_() { PropertiesService.getScriptProperties().deleteProperty(AI_CFG.PROP_LEASE); }

function ai_deleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === AI_CFG.POLL_FN) ScriptApp.deleteTrigger(t); });
}
function ai_hasTrigger_() { return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === AI_CFG.POLL_FN); }

/** 정렬된/비정렬 행 번호 → 연속 구간 [{start, len}] (gap 이하 간격은 한 구간으로) */
function ai_clusters_(rows, gap) {
  const r = Array.from(new Set(rows.map(Number).filter(n => Number.isFinite(n) && n > 0))).sort((a, b) => a - b);
  const out = [];
  r.forEach(n => {
    const c = out[out.length - 1];
    if (c && n - (c.start + c.len - 1) <= gap) c.len = n - c.start + 1; else out.push({ start: n, len: 1 });
  });
  return out;
}
function ai_joinNote_(a, b) { return [String(a || ''), String(b || '')].filter(Boolean).join('; '); }

/**
 * 시트 A열에서 ID 찾기 → {row, dup}. 같은 ID가 2개 이상이면 dup=개수(어느 행인지 정할 수 없으므로 호출자가 보류).
 * cache 는 시트별로 A열을 1회만 읽기 위한 객체(호출자가 빈 객체를 넘김)
 */
function ai_findRowById_(sh, id, cache) {
  if (!cache.colA) { const n = sh.getLastRow(); cache.colA = n >= 2 ? sh.getRange(2, 1, n - 1, 1).getDisplayValues().map(v => ai_normId_(v[0])) : []; }
  let row = 0, dup = 0;
  cache.colA.forEach((v, k) => { if (v === id) { dup++; if (!row) row = k + 2; } });
  return { row, dup: dup > 1 ? dup : 0 };
}

/** system(규칙+사전) 지문: 과목 키 → 문자열 길이·해시. 예열~본 배치 사이에 사전이 바뀌었는지 비교용 */
function ai_promptFp_(ctx) {
  const fp = {};
  Object.keys(ctx.systemCache).forEach(k => {
    const s = ctx.systemCache[k].map(b => b.text).join('\u0001');
    let h = 5381; for (let i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0;
    fp[k] = `${s.length}:${h}`;
  });
  return fp;
}
function ai_normId_(v) {
  let s = String(v === null || v === undefined ? '' : v).trim();
  if (/^\d+(\.0+)?$/.test(s)) s = String(parseInt(s, 10));
  return s;
}
function ai_cleanKw_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/[\u0000-\u001f\u007f\u200b\ufeff]/g, '').replace(/\s+/g, ' ').trim();
}
/** 비교용 정규화: 공백 제거·소문자 */
function ai_norm_(v) { return ai_cleanKw_(v).replace(/\s+/g, '').toLowerCase(); }
function ai_truncate_(s, max) {
  s = String(s || '');
  return s.length > max ? { text: s.slice(0, max) + '\n…(이하 생략)', truncated: true } : { text: s, truncated: false };
}
function ai_now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
