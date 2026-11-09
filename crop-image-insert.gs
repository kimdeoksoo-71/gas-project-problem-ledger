/**
 * CroP 분할 이미지 → 문항 구글문서 일괄 삽입  (crop-image-insert.gs)
 *
 * CroP 앱이 만든 문항별 PNG(<시험지명>_문제_1공통07.png, <시험지명>_해설_1공통07[_cN].png)를
 * 구글드라이브 폴더에 올려 두면, 문항장부(Data 시트)에서 해당 문항을 찾아
 * 그 구글문서의 맨 위에  [문제 이미지] → [해설 이미지 c1 → c2 → …]  순서로 삽입한다.
 *
 * ── 매칭 규칙 (2026-10-24 개편) ───────────────────────────────────────────
 *   시트 E열 (사)세트 값과 파일명의 시험지명을 둘 다 "표준 세트명"(YYX모의NN / YYX서킷NN)으로 바꿔 비교한다.
 *
 *     E열 (사)세트          표준 세트명      파일명 (crops)
 *     X.02            →    24X모의02   ←   24X모의02_문제_1공통01.png
 *     X.0             →    24X모의00   ←   24X모의00_…
 *     써킷.X.05 / 써킷.X05 → 24X서킷05  ←   24X서킷05_…
 *     25.X.0          →    25X모의00   ←   25X모의00_…
 *     25.X.03         →    25X모의03   ←   25X모의03_…
 *     25.써킷X.03      →    25X서킷03   ←   25X서킷03_…
 *     26.X.14         →    26X모의14   ←   26X모의14_…  (구 이름 26모의X14_… 도 인식)
 *
 *   - 연도가 없는 값(X.02, 써킷.X.05)은 CROP_CFG.NO_YEAR_DEFAULT(기본 '24')년으로 본다.
 *   - 회차 번호는 정수로 비교한다 (X.0 = X.00, X05 = X.05).
 *   - 써킷/서킷, 점(.)·공백 유무 차이는 무시한다.
 *   - 과목은 앞 숫자 코드로만 비교 (시트 "3 확률과 통계" ↔ 파일명 3확통), 번호는 정수.
 *   - 규칙으로 안 되는 시험지명은 CROP_CFG.EXAM_MAP 에 '시험지명': 'E열 값' 으로 직접 적는다.
 *   - 매칭이 의심되면 cropCheckSetNames() 를 먼저 돌려 E열 값 ↔ 표준 세트명 대응표를 로그로 확인.
 *
 * ── 파일키 호환 ──────────────────────────────────────────────────────────
 *   색인의 파일키는 표준 세트명으로 쓴다 (예: 26X모의14_1공통22).
 *   2026-10-19 에 넣은 427문항은 구 파일키(26모의X14_1공통22)로 색인·P열·문서 로그가 남아 있으므로,
 *   색인 보존 / P열 중복검사 / 검증 / 중복정리 모두 구 키도 같은 문항으로 인정한다 → 다시 넣지 않는다.
 *
 * ── 실행 순서 ────────────────────────────────────────────────────────────
 *   0) cropCheckSetNames() — (선택) E열 세트명 ↔ 표준 세트명, 폴더 시험지명 대응을 로그로 확인
 *   1) cropBuildIndex()    — 폴더를 훑어 '_CroP색인' 시트를 만든다 (문서는 건드리지 않음)
 *   2) cropStartInsert()   — 색인에서 상태가 빈 행만 처리. 4분마다 끊고 트리거로 이어서 돌린다.
 *      cropStopInsert()    — 중지 (예약된 트리거 삭제)
 *   3) cropVerify()        — ✅ 행의 문서를 열어 이미지 수·로그 줄 수 점검 (색인 M~O열)
 *      cropRetryErrors() / cropFixDuplicates() / cropVerifyReset() — 후처리
 *
 * 처리된 행은 색인 시트에 ✅ 가 남고, 문서 끝과 P열(이력)에 "[yy-MM-dd] CroP 이미지 삽입: …" 한 줄이 남는다.
 *
 * ── 2026-11-09 추가 ──────────────────────────────────────────────────────
 *   - 이미지 폴더를 IMAGE_DS(고정 ID)로 지정. CroP PNG 가 '20 문항 관리/CROPS' 에서 PBMAI/IMAGE_DS 로 옮겨져 CROPS 는 비어 있다.
 *   - CROP_CFG.ONLY_IDS: 삽입 대상을 문항ID 범위로 제한 (색인에 대기 행이 더 있어도 범위 밖은 건드리지 않음).
 *   - 삽입이 끝난 문항은 원본 시트([풀세트]/[써킷]/[문항]) O열(등록)에 1 을 적는다 (CROP_CFG.WRITE_REG).
 */

const CROP_CFG = {
  // ── 이미지 폴더 ─────────────────────────────────────────────
  FOLDER_ID: '1fPY6942RrCcsmoHL-3HbtSk8F210ngXi',   // PBMAI/IMAGE_DS (2026-11-09). 비우면 '20 문항 관리' 아래 FOLDER_NAME 으로 찾음
  FOLDER_NAME: 'CROPS',

  // ── 문항장부 ────────────────────────────────────────────────
  TARGET_SHEETS: ['Data'],
  INDEX_SHEET: '_CroP색인',
  COL_ID: 1,     // A  ID
  COL_SET: 5,    // E  (사)세트
  COL_SUBJ: 6,   // F  (사)과목   예) "1 공통"
  COL_NO: 7,     // G  (사)번호
  COL_DOC: 14,   // N  구글문서 링크 (셀 링크 또는 텍스트)
  COL_REG: 15,   // O  등록 (0/1 목록 검증). 삽입 완료 시 1
  COL_PLOG: 16,  // P  이력

  // ── 삽입 대상 제한 ──────────────────────────────────────────
  //    null 이면 색인의 대기 행 전부. {prefix, from, to} 면 그 문항ID 범위만 삽입한다 (예: S0001~S1014).
  //    범위 밖 대기 행은 색인에 대기(빈 상태)로 그대로 남는다.
  ONLY_IDS: { prefix: 'S', from: 1015, to: 1090 },

  // ── 세트명 해석 ─────────────────────────────────────────────
  NO_YEAR_DEFAULT: '24',         // E열에 연도가 없을 때(X.02, 써킷.X.05) 붙일 연도

  // ── 시험지명 → (사)세트 예외 매핑 (규칙으로 안 되는 것만) ──────
  //    키: 파일명의 시험지명,  값: E열에 실제로 적힌 세트명 그대로
  //    예) '26X서킷TB01': '26.써킷TB.01'
  EXAM_MAP: {
  },

  // ── 삽입 모양 ───────────────────────────────────────────────
  IMG_MAX_WIDTH_PX: 600,   // 이 폭(px)보다 넓은 이미지는 비율 유지해 줄임. 0 이면 원본 크기 그대로
  ADD_CAPTION: false,      // true 면 이미지 위에 "▣ 26X모의14 1공통22" 한 줄 추가
  WRITE_LOG: true,         // 문서 끝 + P열에 처리 로그 한 줄
  WRITE_REG: true,         // 삽입 완료(이미삽입됨 포함) 시 원본 시트 O열(등록)에 1
  LOG_TAG: 'CroP 이미지 삽입',

  // ── 배치 실행 ───────────────────────────────────────────────
  TIME_BUDGET_MS: 4 * 60 * 1000,   // 큰 이미지 5~6장짜리 문항은 1건에 1분 가까이 걸릴 수 있어 6분 한도에 여유를 둔다
  NEXT_RUN_DELAY_MS: 60 * 1000,
  SLEEP_MS: 200,
  DRY_RUN: false,          // true 면 문서를 열어 확인만 하고 넣지 않음 (색인에 '🧪' 표시)

  TRIGGER_HANDLER: 'cropContinueInsert',
  PROP_RUNNING: 'CROP_INSERT_RUNNING',
  VERIFY_TRIGGER_HANDLER: 'cropContinueVerify',
  PROP_VERIFYING: 'CROP_VERIFY_RUNNING',
};

// 색인 시트 열 (1-based).  M~O 는 cropVerify 가 채우는 검증 열
const CROP_IDX = {
  KEY: 1, ITEM_ID: 2, SHEET: 3, ROW: 4, MATCH_KEY: 5, DOC_URL: 6,
  PROB_FILE: 7, SOL_FILES: 8, SOL_COUNT: 9, STATUS: 10, MSG: 11, TIME: 12,
  V_IMGS: 13, V_LOGS: 14, V_RESULT: 15,
  HEADER: ['파일키', '문항ID', '시트', '행', '사용세트식별', '문서URL',
           '문제파일ID', '해설파일ID(순서대로)', '해설조각수', '상태', '메시지', '처리시각',
           '문서이미지수', '로그줄수', '검증'],
};

const CROP_FILE_RE = /^(.+?)_(문제|해설)_(\d)(공통|확통|미적|기하)(\d{1,2})(?:_c(\d+))?\.png$/i;

/* ═══════════════════════════════════════════════════════════════
 * 세트명 정규화
 * ═══════════════════════════════════════════════════════════════ */

/**
 * 파일명의 시험지명 → 표준 세트명 정보.
 *   새 이름: 24X모의02, 25X서킷03, 26X모의14
 *   구 이름: 26모의X14, 26서킷X04 (2026-10-19 이전 업로드분)
 * @return {{canon:string, yy:string, kind:string, nn:string}|null}
 */
function cropParseExam_(exam) {
  const s = cropNfc_(exam).replace(/\s+/g, '');
  let m = s.match(/^(\d{2})X(모의|서킷|써킷)(\d{1,3})$/i);           // 새 이름
  if (!m) {
    const o = s.match(/^(\d{2})(모의|서킷|써킷)X(\d{1,3})$/i);         // 구 이름
    if (o) m = [o[0], o[1], o[2], o[3]];
  }
  if (!m) return null;
  return cropMakeCanon_(m[1], m[2], m[3]);
}

/**
 * 시트 E열 (사)세트 값 → 표준 세트명. X 계열이 아니면 '' (예: '24년 D조 창고', 'CS.써킷.01').
 *   X.02 / X.0 / 써킷.X.05 / 써킷.X05 / 25.X.0 / 25.X.03 / 25.써킷X.03 / 26.X.14 / 26.써킷X.04
 */
function cropSetToCanon_(setName) {
  const s = cropNfc_(setName).replace(/\s+/g, '');
  const m = s.match(/^(?:(\d{2})\.)?(써킷|서킷|모의)?\.?X\.?(\d{1,3})$/i);
  if (!m) return '';
  const yy = m[1] || CROP_CFG.NO_YEAR_DEFAULT;
  return cropMakeCanon_(yy, m[2] || '모의', m[3]).canon;
}

function cropMakeCanon_(yy, kind, nn) {
  const k = /^(써킷|서킷)$/.test(kind) ? '서킷' : '모의';
  const n2 = String(parseInt(nn, 10)).padStart(2, '0');
  return { canon: `${yy}X${k}${n2}`, yy: String(yy), kind: k, nn: n2 };
}

/** 시트 쪽 세트 식별자: 표준 세트명이 되면 그것, 아니면 공백 뺀 원문 */
function cropSetIdForSheet_(setName) {
  return cropSetToCanon_(setName) || cropNfc_(setName).replace(/\s+/g, '');
}

/** 파일 쪽 세트 식별자: EXAM_MAP 우선(값은 E열 원문), 없으면 시험지명을 표준 세트명으로 */
function cropSetIdForExam_(exam) {
  const e = cropNfc_(exam);
  for (const k in CROP_CFG.EXAM_MAP) {
    if (cropNfc_(k) === e) return cropSetIdForSheet_(CROP_CFG.EXAM_MAP[k]);
  }
  const p = cropParseExam_(e);
  return p ? p.canon : '';
}

/** 시험지명을 파일키용 표준 이름으로 (해석 안 되면 원문) */
function cropCanonExamName_(exam) {
  const p = cropParseExam_(exam);
  return p ? p.canon : cropNfc_(exam);
}

/**
 * 파일키의 별칭들: [표준 키, 구 키]. 예) '26X모의14_1공통22' → ['26X모의14_1공통22', '26모의X14_1공통22']
 * 구 키는 2026-10-19 삽입분의 색인·P열·문서 로그에 남아 있는 형식.
 */
function cropKeyAliases_(fileKey) {
  const key = cropNfc_(fileKey);
  const i = key.indexOf('_');
  if (i < 0) return [key];
  const p = cropParseExam_(key.slice(0, i));
  if (!p) return [key];
  const rest = key.slice(i);
  const canon = p.canon + rest;
  const legacy = `${p.yy}${p.kind}X${p.nn}` + rest;
  return canon === legacy ? [canon] : [canon, legacy];
}

/** text 안에 fileKey(별칭 포함)의 로그가 있는가 */
function cropHasLog_(text, fileKey) {
  const t = cropNfc_(text);
  return cropKeyAliases_(fileKey).some(k => t.indexOf(`${CROP_CFG.LOG_TAG}: ${k}`) !== -1);
}

/* ═══════════════════════════════════════════════════════════════
 * 0단계 (선택): 세트명 대응 점검
 * ═══════════════════════════════════════════════════════════════ */
function cropCheckSetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSets = {};    // 원문 → {canon, n}
  CROP_CFG.TARGET_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, CROP_CFG.COL_SET, sh.getLastRow() - 1, 1).getValues().forEach(([v]) => {
      const raw = cropNfc_(v).trim();
      if (!raw) return;
      const o = sheetSets[raw] || (sheetSets[raw] = { canon: cropSetToCanon_(raw), n: 0 });
      o.n++;
    });
  });

  const lines = ['== 시트 E열 X 세트 → 표준 세트명 =='];
  const canonFrom = {};
  Object.keys(sheetSets).sort().forEach(raw => {
    const o = sheetSets[raw];
    if (!o.canon) return;
    lines.push(`  ${raw}  →  ${o.canon}  (${o.n}행)`);
    (canonFrom[o.canon] || (canonFrom[o.canon] = [])).push(raw);
  });
  const clash = Object.keys(canonFrom).filter(c => canonFrom[c].length > 1);
  if (clash.length) {
    lines.push('', '⚠ 서로 다른 E열 값이 같은 표준 세트명으로 합쳐짐:');
    clash.forEach(c => lines.push(`  ${c} ← ${canonFrom[c].join(' , ')}`));
  }
  const xLike = Object.keys(sheetSets).filter(r => !sheetSets[r].canon && /X/i.test(r));
  if (xLike.length) lines.push('', 'ℹ X 가 들어 있지만 규칙에 안 맞아 제외된 E열 값: ' + xLike.join(', '));

  const folder = cropGetFolder_();
  if (folder) {
    const exams = {};
    const it = folder.getFilesByType(MimeType.PNG);
    while (it.hasNext()) {
      const m = cropNfc_(it.next().getName()).match(CROP_FILE_RE);
      if (m) exams[m[1]] = (exams[m[1]] || 0) + 1;
    }
    lines.push('', '== 폴더 시험지명 → 세트 식별자 (시트에 있음?) ==');
    Object.keys(exams).sort().forEach(e => {
      const id = cropSetIdForExam_(e);
      const inSheet = id && Object.keys(sheetSets).some(r => cropSetIdForSheet_(r) === id);
      lines.push(`  ${e} → ${id || '(해석불가: EXAM_MAP 필요)'}  ${inSheet ? '✅' : '❌ 시트에 없음'}  PNG ${exams[e]}`);
    });
  } else {
    lines.push('', `⚠ 이미지 폴더를 찾지 못해 파일 쪽 점검은 건너뜀 ('20 문항 관리/${CROP_CFG.FOLDER_NAME}')`);
  }
  const out = lines.join('\n');
  Logger.log(out);
  cropAlert_(out.length > 1800 ? out.slice(0, 1800) + '\n… (전체는 실행 로그 참고)' : out);
}

/* ═══════════════════════════════════════════════════════════════
 * 1단계: 색인 만들기
 * ═══════════════════════════════════════════════════════════════ */
function cropBuildIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = cropGetFolder_();
  if (!folder) { cropAlert_(`⚠️ 이미지 폴더를 찾을 수 없습니다. (FOLDER_ID 또는 '20 문항 관리/${CROP_CFG.FOLDER_NAME}')`); return; }

  // 1) 폴더의 PNG 를 파일키별로 묶기 (파일키는 표준 세트명 기준이라 구/새 이름이 섞여도 한 문항으로 모인다)
  const groups = {};      // fileKey -> { exam, code, subj, no, prob:{id,name}, sols:[{c,id,name}] }
  let pngCount = 0;
  const badNames = [];
  const it = folder.getFilesByType(MimeType.PNG);
  while (it.hasNext()) {
    const f = it.next();
    const name = cropNfc_(f.getName());     // 맥에서 올린 파일은 한글이 NFD(자소 분리)로 올 수 있어 NFC 로 맞춤
    const m = name.match(CROP_FILE_RE);
    if (!m) { badNames.push(name); continue; }
    pngCount++;
    const [, exam, kind, code, subj, noStr, cStr] = m;
    const no = parseInt(noStr, 10);
    const fileKey = `${cropCanonExamName_(exam)}_${code}${subj}${String(no).padStart(2, '0')}`;
    const g = groups[fileKey] || (groups[fileKey] = { exam, code, subj, no, prob: null, sols: [] });
    if (kind === '문제') {
      if (g.prob) g.dupProb = true;
      g.prob = { id: f.getId(), name };
    } else {
      g.sols.push({ c: cStr ? parseInt(cStr, 10) : 0, id: f.getId(), name });
    }
  }
  Logger.log(`PNG ${pngCount}개, 문항키 ${Object.keys(groups).length}개, 이름규칙 불일치 ${badNames.length}개`);
  if (badNames.length) Logger.log('규칙 불일치 파일(최대 20): ' + badNames.slice(0, 20).join(', '));

  // 2) 시트 조회표: 매칭키 -> [{sheetName,row,id,docUrl}]
  const lookup = cropBuildSheetLookup_(ss);

  // 3) 기존 색인에서 ✅ 상태는 보존 (구 파일키도 표준 키로 바꿔 읽음 → 재실행해도 중복 삽입 방지)
  const keep = cropReadExistingIndex_(ss);

  // 4) 색인 행 만들기
  const keys = Object.keys(groups).sort();
  const rows = keys.map(fileKey => {
    const g = groups[fileKey];
    const setId = cropSetIdForExam_(g.exam);
    const matchKey = setId ? cropNormKey_(setId, g.code, g.no) : '';
    const hits = matchKey ? (lookup[matchKey] || []) : [];

    // 해설 조각 정렬·연속성 검사
    g.sols.sort((a, b) => a.c - b.c);
    let solMsg = '';
    if (g.sols.length > 1) {
      const cs = g.sols.map(s => s.c);
      if (!cs.every((c, i) => c === i + 1)) solMsg = `해설 조각번호 이상(${cs.join(',')})`;
    }

    let status = '';
    const msg = [];
    if (!setId)                { status = '❌ 세트규칙없음'; msg.push(`시험지명 "${g.exam}" → EXAM_MAP 에 추가 필요`); }
    else if (hits.length === 0){ status = '❌ 매칭없음';    msg.push(`시트에 "${matchKey}" 없음`); }
    else if (hits.length > 1)  { status = '❌ 중복';        msg.push(hits.map(h => `${h.sheetName}!${h.row}(${h.id}, E="${h.setRaw}")`).join(' / ')); }
    if (!g.prob)               { status = status || '❌ 문제파일없음'; msg.push('문제 PNG 없음'); }
    if (g.dupProb)             { status = status || '❌ 문제파일중복'; msg.push('문제 PNG 2개 이상 (구/새 이름이 함께 올라갔는지 확인)'); }
    if (solMsg)                { status = status || '❌ 조각누락'; msg.push(solMsg); }
    if (g.sols.length === 0)   { msg.push('해설 PNG 없음 → 문제만 삽입'); }

    const hit = hits.length === 1 ? hits[0] : null;
    if (hit && !hit.docUrl)    { status = status || '❌ 문서링크없음'; msg.push('N열 비어있음'); }

    const base = [fileKey, hit ? hit.id : '', hit ? hit.sheetName : '', hit ? hit.row : '', matchKey,
                  hit ? hit.docUrl : '', g.prob ? g.prob.id : '', g.sols.map(s => s.id).join(','), g.sols.length];

    const prev = keep[fileKey];
    if (prev && /^✅/.test(prev.status)) {       // 실제로 넣은 행만 보존 (🧪 DRY_RUN 은 다시 대기로)
      return base.concat([prev.status, prev.msg, prev.time, prev.vImgs, prev.vLogs, prev.vResult]);
    }
    return base.concat([status, msg.join(' | '), '', '', '', '']);
  });

  // 5) 색인 시트 쓰기
  let sh = ss.getSheetByName(CROP_CFG.INDEX_SHEET);
  if (!sh) sh = ss.insertSheet(CROP_CFG.INDEX_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, 1, CROP_IDX.HEADER.length).setValues([CROP_IDX.HEADER]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, CROP_IDX.HEADER.length).setValues(rows);
  sh.setFrozenRows(1);

  // 6) 요약
  const cnt = {};
  rows.forEach(r => { const s = r[CROP_IDX.STATUS - 1] || '(대기)'; cnt[s] = (cnt[s] || 0) + 1; });
  const summary = [
    `색인 완료: 문항키 ${rows.length}개 (PNG ${pngCount}개)`,
    ...Object.keys(cnt).sort().map(k => `  ${k}: ${cnt[k]}`),
    badNames.length ? `  이름규칙 불일치 PNG: ${badNames.length}개 (로그 참고)` : '',
    '',
    `'${CROP_CFG.INDEX_SHEET}' 시트에서 ❌ 행을 확인한 뒤 cropStartInsert 를 실행하세요.`,
  ].filter(Boolean).join('\n');
  Logger.log(summary);
  cropAlert_(summary);
}

/* ═══════════════════════════════════════════════════════════════
 * 2단계: 삽입 (시작 / 이어서 / 중지)
 * ═══════════════════════════════════════════════════════════════ */
function cropStartInsert() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(CROP_CFG.INDEX_SHEET)) {
    cropAlert_(`⚠️ '${CROP_CFG.INDEX_SHEET}' 시트가 없습니다. cropBuildIndex 를 먼저 실행하세요.`);
    return;
  }
  cropDeleteTriggers_();
  PropertiesService.getScriptProperties().setProperty(CROP_CFG.PROP_RUNNING, '1');
  cropAlert_(cropRunBatch_().summary);
}

function cropContinueInsert() {   // 트리거 핸들러
  if (PropertiesService.getScriptProperties().getProperty(CROP_CFG.PROP_RUNNING) !== '1') { cropDeleteTriggers_(); return; }
  Logger.log(cropRunBatch_().summary);
}

function cropStopInsert() {
  PropertiesService.getScriptProperties().deleteProperty(CROP_CFG.PROP_RUNNING);
  cropDeleteTriggers_();
  cropAlert_('⏹ 중지했습니다. 다시 시작하려면 cropStartInsert 를 실행하세요 (처리된 행은 건너뜁니다).');
}

/** 시간 예산 안에서 색인의 '대기' 행을 처리하고, 남았으면 다음 트리거를 예약한다. */
function cropRunBatch_() {
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) return { summary: '다른 실행이 진행 중입니다. 잠시 후 다시 시도하세요.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const idx = ss.getSheetByName(CROP_CFG.INDEX_SHEET);
    const lastRow = idx.getLastRow();
    if (lastRow < 2) { cropFinish_(); return { summary: '색인이 비어 있습니다.' }; }

    const data = idx.getRange(2, 1, lastRow - 1, CROP_IDX.HEADER.length).getValues();
    const targetSheets = {};
    CROP_CFG.TARGET_SHEETS.forEach(n => { const s = ss.getSheetByName(n); if (s) targetSheets[n] = s; });

    const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
    let done = 0, failed = 0, remaining = 0, outOfScope = 0, stoppedForTime = false;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const st = String(row[CROP_IDX.STATUS - 1] || '').trim();
      const pending = st === '' || (st.startsWith('🧪') && !CROP_CFG.DRY_RUN);   // 대기 행 (+ DRY_RUN 흔적은 실제 실행 때 다시 처리)
      if (!pending) continue;
      if (!cropInScope_(row[CROP_IDX.ITEM_ID - 1])) { outOfScope++; continue; }   // 대상 범위 밖 → 대기로 둔다
      if (Date.now() - t0 > CROP_CFG.TIME_BUDGET_MS) { stoppedForTime = true; remaining++; continue; }

      const sheetRow = i + 2;
      const fileKey = row[CROP_IDX.KEY - 1];
      const stamp = Utilities.formatDate(new Date(), tz, 'yy-MM-dd HH:mm');
      try {
        const res = cropProcessItem_(row, targetSheets, tz);
        idx.getRange(sheetRow, CROP_IDX.STATUS, 1, 3).setValues([[res.status, res.msg, stamp]]);
        done++;
      } catch (e) {
        idx.getRange(sheetRow, CROP_IDX.STATUS, 1, 3).setValues([['❌ 오류', String(e && e.message || e).slice(0, 500), stamp]]);
        Logger.log(`❌ ${fileKey}: ${e && e.stack || e}`);
        failed++;
      }
      Utilities.sleep(CROP_CFG.SLEEP_MS);
    }

    const elapsed = Math.round((Date.now() - t0) / 1000);
    if (stoppedForTime) {
      cropDeleteTriggers_();
      ScriptApp.newTrigger(CROP_CFG.TRIGGER_HANDLER).timeBased().after(CROP_CFG.NEXT_RUN_DELAY_MS).create();
      return { summary: `이번 실행: 성공 ${done}, 실패 ${failed} (${elapsed}초). 남은 ${remaining}개는 약 ${Math.round(CROP_CFG.NEXT_RUN_DELAY_MS / 1000)}초 뒤 자동으로 이어서 처리합니다.` };
    }
    cropFinish_();
    return { summary: `✅ 전체 완료: 이번 실행 성공 ${done}, 실패 ${failed} (${elapsed}초). 결과는 '${CROP_CFG.INDEX_SHEET}' 시트 참고.` +
      (outOfScope ? `\n대상 범위(${cropScopeLabel_()}) 밖이라 건너뛴 대기 행 ${outOfScope}개` : '') };
  } finally {
    lock.releaseLock();
  }
}

/** 문항ID 가 CROP_CFG.ONLY_IDS 범위 안인가 (ONLY_IDS 가 null 이면 항상 true) */
function cropInScope_(itemId) {
  const o = CROP_CFG.ONLY_IDS;
  if (!o) return true;
  const m = String(itemId || '').trim().match(/^([A-Za-z]+)(\d+)$/);
  if (!m || m[1] !== o.prefix) return false;
  const n = parseInt(m[2], 10);
  return n >= o.from && n <= o.to;
}

function cropScopeLabel_() {
  const o = CROP_CFG.ONLY_IDS;
  if (!o) return '전체';
  const pad = n => String(n).padStart(4, '0');
  return `${o.prefix}${pad(o.from)}~${o.prefix}${pad(o.to)}`;
}

function cropFinish_() {
  PropertiesService.getScriptProperties().deleteProperty(CROP_CFG.PROP_RUNNING);
  cropDeleteTriggers_();
}



/* ═══════════════════════════════════════════════════════════════
 * 3단계: 검증 — ✅ 행의 문서를 실제로 열어 이미지 수와 로그 줄 수를 센다
 *   로그줄수 = 1 이 정상. 2 이상이면 같은 문항에 두 번 들어간 것(시간 초과로 죽은 뒤 재처리된 경우).
 *   문서이미지수 < 1+해설조각수 이면 이미지가 빠진 것. 초과는 원래 있던 그림일 수 있어 참고만.
 *   결과는 색인 M~O 열. 시간 예산 안에 못 끝내면 트리거로 이어서 돈다. 재실행 시 검증 열이 빈 행만 본다.
 * ═══════════════════════════════════════════════════════════════ */
function cropVerify() {
  cropDeleteTriggers_(CROP_CFG.VERIFY_TRIGGER_HANDLER);
  PropertiesService.getScriptProperties().setProperty(CROP_CFG.PROP_VERIFYING, '1');
  cropAlert_(cropVerifyBatch_().summary);
}

function cropContinueVerify() {   // 트리거 핸들러
  if (PropertiesService.getScriptProperties().getProperty(CROP_CFG.PROP_VERIFYING) !== '1') { cropDeleteTriggers_(CROP_CFG.VERIFY_TRIGGER_HANDLER); return; }
  Logger.log(cropVerifyBatch_().summary);
}

/** 검증 열(M~O)을 비워 처음부터 다시 검증할 때 */
function cropVerifyReset() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CROP_CFG.INDEX_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  sh.getRange(2, CROP_IDX.V_IMGS, sh.getLastRow() - 1, 3).clearContent();
  cropAlert_('검증 열을 비웠습니다. cropVerify 를 실행하세요.');
}

function cropVerifyBatch_() {
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) return { summary: '다른 실행이 진행 중입니다.' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const idx = ss.getSheetByName(CROP_CFG.INDEX_SHEET);
    const lastRow = idx.getLastRow();
    if (lastRow < 2) return { summary: '색인이 비어 있습니다.' };
    idx.getRange(1, 1, 1, CROP_IDX.HEADER.length).setValues([CROP_IDX.HEADER]);
    const data = idx.getRange(2, 1, lastRow - 1, CROP_IDX.HEADER.length).getValues();
    let done = 0, flagged = 0, remaining = 0, stopped = false;

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!/^✅/.test(String(r[CROP_IDX.STATUS - 1] || ''))) continue;
      if (String(r[CROP_IDX.V_RESULT - 1] || '') !== '') continue;        // 이미 검증한 행
      if (Date.now() - t0 > CROP_CFG.TIME_BUDGET_MS) { stopped = true; remaining++; continue; }

      const key = String(r[CROP_IDX.KEY - 1]);
      const expected = 1 + (parseInt(r[CROP_IDX.SOL_COUNT - 1], 10) || 0);
      let imgs = '', logs = '', result = '';
      try {
        const doc = DocumentApp.openById(cropExtractDocId_(r[CROP_IDX.DOC_URL - 1]));
        const body = doc.getBody();
        imgs = body.getImages().length;
        logs = 0;
        const n = body.getNumChildren();
        for (let c = 0; c < n; c++) {
          const el = body.getChild(c);
          if (el.getType() === DocumentApp.ElementType.PARAGRAPH && cropHasLog_(el.asParagraph().getText(), key)) logs++;
        }
        const notes = [];
        if (logs !== 1) notes.push(logs === 0 ? '⚠ 로그없음' : `⚠ 중복(로그 ${logs}회)`);
        if (imgs < expected) notes.push(`⚠ 이미지부족(${imgs}/${expected})`);
        else if (imgs > expected && logs <= 1) notes.push(`ℹ 이미지초과(${imgs}/${expected}, 기존 그림 가능)`);
        result = notes.length ? notes.join(' ') : '✅ 정상';
        if (notes.some(s => s.startsWith('⚠'))) flagged++;
      } catch (e) {
        result = '❌ 열기실패: ' + String(e && e.message || e).slice(0, 200);
        flagged++;
      }
      idx.getRange(i + 2, CROP_IDX.V_IMGS, 1, 3).setValues([[imgs, logs, result]]);
      done++;
      Utilities.sleep(100);
    }

    if (stopped) {
      cropDeleteTriggers_(CROP_CFG.VERIFY_TRIGGER_HANDLER);
      ScriptApp.newTrigger(CROP_CFG.VERIFY_TRIGGER_HANDLER).timeBased().after(CROP_CFG.NEXT_RUN_DELAY_MS).create();
      return { summary: `검증 ${done}건 (⚠/❌ ${flagged}건). 남은 ${remaining}건은 1분 뒤 자동으로 이어서 검증합니다.` };
    }
    PropertiesService.getScriptProperties().deleteProperty(CROP_CFG.PROP_VERIFYING);
    cropDeleteTriggers_(CROP_CFG.VERIFY_TRIGGER_HANDLER);
    return { summary: `✅ 검증 완료: 이번 실행 ${done}건, ⚠/❌ ${flagged}건. '${CROP_CFG.INDEX_SHEET}' O열(검증)에서 ⚠ 행을 확인하세요.` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 중복 삽입된 문항 정리: 검증에서 '⚠ 중복' 으로 나온 행에 대해
 *   문서 맨 위의 이미지 (1+해설조각수)장과 로그 줄 1개를 지워 한 벌만 남긴다.
 *   실행 후 cropVerifyReset → cropVerify 로 다시 확인할 것.
 */
function cropFixDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idx = ss.getSheetByName(CROP_CFG.INDEX_SHEET);
  const data = idx.getRange(2, 1, idx.getLastRow() - 1, CROP_IDX.HEADER.length).getValues();
  let fixed = 0;
  const errs = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (!/⚠ 중복/.test(String(r[CROP_IDX.V_RESULT - 1] || ''))) continue;
    const key = String(r[CROP_IDX.KEY - 1]);
    const expected = 1 + (parseInt(r[CROP_IDX.SOL_COUNT - 1], 10) || 0);
    try {
      const doc = DocumentApp.openById(cropExtractDocId_(r[CROP_IDX.DOC_URL - 1]));
      const body = doc.getBody();
      // 맨 위부터 이미지 문단 expected 개 제거
      let removed = 0, c = 0;
      while (removed < expected && c < body.getNumChildren()) {
        const el = body.getChild(c);
        if (el.getType() === DocumentApp.ElementType.PARAGRAPH && el.asParagraph().getNumChildren() > 0 &&
            el.asParagraph().getChild(0).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          body.removeChild(el); removed++;
        } else c++;
      }
      // 로그 줄 하나 제거 (뒤에서부터, 구 파일키 로그 포함)
      for (let k = body.getNumChildren() - 1; k >= 0; k--) {
        const el = body.getChild(k);
        if (el.getType() === DocumentApp.ElementType.PARAGRAPH && cropHasLog_(el.asParagraph().getText(), key)) {
          if (k === body.getNumChildren() - 1) el.asParagraph().clear(); else body.removeChild(el);   // 마지막 문단은 지울 수 없어 비움(setText('')은 '빈 텍스트 요소' 오류, 2026-11-10)
          break;
        }
      }
      doc.saveAndClose();
      idx.getRange(i + 2, CROP_IDX.V_RESULT).setValue(`🔧 중복 1벌 제거(이미지 ${removed}장) → 재검증 필요`);
      fixed++;
    } catch (e) { errs.push(`${key}: ${e && e.message || e}`); }
  }
  cropAlert_(`중복 정리 ${fixed}건` + (errs.length ? `\n실패:\n${errs.join('\n')}` : '') + `\n\ncropVerifyReset → cropVerify 로 다시 확인하세요.`);
}

/** '❌ 오류' 행을 대기로 되돌려 cropStartInsert 에서 다시 처리하게 한다 */
function cropRetryErrors() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CROP_CFG.INDEX_SHEET);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const st = sh.getRange(2, CROP_IDX.STATUS, n, 3).getValues();
  let cnt = 0;
  st.forEach((r, i) => { if (/^❌ 오류/.test(String(r[0]))) { sh.getRange(i + 2, CROP_IDX.STATUS, 1, 3).setValues([['', '', '']]); cnt++; } });
  cropAlert_(`${cnt}개 행을 대기로 되돌렸습니다. cropStartInsert 를 실행하세요.`);
}

/* ═══════════════════════════════════════════════════════════════
 * 보조 함수
 * ═══════════════════════════════════════════════════════════════ */

/**
 * 매칭키: 세트 식별자(표준 세트명 또는 원문) + 과목 앞 숫자 코드 + 정수 번호.
 *   예) "26X모의14:1:22"
 */
function cropNormKey_(setId, subj, no) {
  const s = cropNfc_(setId).replace(/\s+/g, '');
  const fm = String(subj == null ? '' : subj).match(/\d/);
  const n = parseInt(String(no).replace(/\D/g, ''), 10);
  if (!s || !fm || !Number.isFinite(n)) return '';
  return `${s}:${fm[0]}:${n}`;
}

/** 대상 시트들을 한 번에 읽어 매칭키 → [{sheetName,row,id,docUrl,setRaw}] */
function cropBuildSheetLookup_(ss) {
  const lookup = {};
  CROP_CFG.TARGET_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { Logger.log(`⚠️ 시트 없음: ${name}`); return; }
    const last = sh.getLastRow();
    if (last < 2) return;
    const n = last - 1;
    const vals = sh.getRange(2, 1, n, CROP_CFG.COL_PLOG).getValues();
    const rich = sh.getRange(2, CROP_CFG.COL_DOC, n, 1).getRichTextValues();
    for (let i = 0; i < n; i++) {
      const r = vals[i];
      const id = String(r[CROP_CFG.COL_ID - 1] || '').trim();
      const setRaw = cropNfc_(r[CROP_CFG.COL_SET - 1]).trim();
      const subj = r[CROP_CFG.COL_SUBJ - 1], no = r[CROP_CFG.COL_NO - 1];
      if (!id || setRaw === '' || subj === '' || no === '') continue;
      const key = cropNormKey_(cropSetIdForSheet_(setRaw), subj, no);
      if (!key) continue;
      let docUrl = '';
      try { docUrl = (rich[i][0] && rich[i][0].getLinkUrl()) || ''; } catch (e) {}
      if (!docUrl) docUrl = String(r[CROP_CFG.COL_DOC - 1] || '').trim();
      (lookup[key] || (lookup[key] = [])).push({ sheetName: name, row: i + 2, id, docUrl, setRaw });
    }
  });
  return lookup;
}

/** 기존 색인 시트의 상태 보존용 읽기: 표준 fileKey → {status,msg,time,vImgs,vLogs,vResult} */
function cropReadExistingIndex_(ss) {
  const out = {};
  const sh = ss.getSheetByName(CROP_CFG.INDEX_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;
  const width = Math.min(sh.getLastColumn(), CROP_IDX.HEADER.length);
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  vals.forEach(r => {
    const raw = String(r[CROP_IDX.KEY - 1] || '');
    if (!raw) return;
    const k = cropKeyAliases_(raw)[0];          // 구 파일키(26모의X14_…) → 표준 키(26X모의14_…)
    const item = {
      status: String(r[CROP_IDX.STATUS - 1] || ''), msg: r[CROP_IDX.MSG - 1], time: r[CROP_IDX.TIME - 1],
      vImgs: width >= CROP_IDX.V_IMGS ? r[CROP_IDX.V_IMGS - 1] : '',
      vLogs: width >= CROP_IDX.V_LOGS ? r[CROP_IDX.V_LOGS - 1] : '',
      vResult: width >= CROP_IDX.V_RESULT ? r[CROP_IDX.V_RESULT - 1] : '',
    };
    if (!out[k] || /^✅/.test(item.status)) out[k] = item;   // 같은 문항이 구/새 키로 둘 다 있으면 ✅ 쪽 우선
  });
  return out;
}

function cropGetFolder_() {
  if (CROP_CFG.FOLDER_ID) {
    try { return DriveApp.getFolderById(CROP_CFG.FOLDER_ID); } catch (e) { return null; }
  }
  const roots = DriveApp.getFoldersByName('20 문항 관리');
  if (!roots.hasNext()) return null;
  const sub = roots.next().getFoldersByName(CROP_CFG.FOLDER_NAME);
  return sub.hasNext() ? sub.next() : null;
}

function cropFitWidth_(img) {
  const max = CROP_CFG.IMG_MAX_WIDTH_PX;
  if (!max) return;
  try {
    const w = img.getWidth(), h = img.getHeight();
    if (w > max) { img.setWidth(max); img.setHeight(Math.round(h * max / w)); }
  } catch (e) { Logger.log('이미지 크기 조정 실패(무시): ' + e); }
}

function cropExtractDocId_(url) {
  const s = String(url || '').trim();
  let m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9-_]+)/);     if (m) return m[1];
  m = s.match(/^[-\w]{25,}$/);                if (m) return m[0];
  return null;
}

function cropDeleteTriggers_(handler) {
  const h = handler || CROP_CFG.TRIGGER_HANDLER;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === h) ScriptApp.deleteTrigger(t);
  });
}

/** 한글 유니코드 정규화(NFC). 맥 업로드 파일명(NFD) ↔ 시트 값(NFC) 불일치 방지 */
function cropNfc_(s) {
  s = (s == null) ? '' : String(s);
  try { return s.normalize('NFC'); } catch (e) { return s; }
}

/** 진단용: 폴더의 PNG 이름 몇 개가 NFC 인지, 규칙에 맞는지, 어떤 세트로 해석되는지 로그로 보여준다 */
function cropDebugNames() {
  const folder = cropGetFolder_();
  if (!folder) { Logger.log('폴더 없음'); return; }
  const it = folder.getFilesByType(MimeType.PNG);
  let i = 0;
  while (it.hasNext() && i < 5) {
    const raw = it.next().getName();
    const nfc = cropNfc_(raw);
    const m = nfc.match(CROP_FILE_RE);
    Logger.log(`${raw}  | 원본이 NFC? ${raw === nfc}, 규칙일치? ${!!m}, 세트 → ${m ? cropSetIdForExam_(m[1]) || '(해석불가)' : '-'}`);
    i++;
  }
}

/** 에디터/트리거 실행에서는 getUi 가 실패하므로 로그로 대체 */
function cropAlert_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}