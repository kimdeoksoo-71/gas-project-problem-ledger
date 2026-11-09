/**
 * crop-data-plog-fix.gs  (2026-10-24)
 *
 * 문제: 'Data' 시트 A2 는 ={'[풀세트]'!2:2000;'[써킷]'!2:751;'[문항]'!2:199} 배열 수식이다.
 *       CroP 이미지 삽입 스크립트가 TARGET_SHEETS=['Data'] 기준으로 P열(이력) 로그를
 *       Data 시트 셀에 직접 써서, 배열이 펼쳐질 자리가 막혀 A2 가 #REF! 가 됐다.
 *       → Data 가 통째로 비고, Data 를 참조하는 '문제검색' 이 "결과 없음" 이 된다.
 *
 * 이 파일이 하는 일
 *   1) cropRepairDataPLog()  — Data 시트 P열에 잘못 써진 CroP 로그를 원본 시트
 *                              ([풀세트]/[써킷]/[문항]) 의 같은 문항ID 행 P열로 옮기고,
 *                              Data P열을 비워 배열 수식을 되살린다. (한 번만 실행)
 *   2) cropSourcePCell_()    — 앞으로 P열 로그는 문항ID로 원본 시트 행을 찾아 거기에 쓴다.
 *                              cropProcessItem_ 가 이 함수를 쓰도록 교체본을 아래에 둔다.
 *
 * 적용 순서
 *   ① cropStopInsert 실행 (진행 중인 삽입 중지)
 *   ② 이 파일을 Apps Script 프로젝트에 새 파일로 추가
 *   ③ 기존 crop-image-insert.gs 에서 function cropProcessItem_ 전체를 지운다
 *      (이 파일에 교체본이 있음. 같은 이름이 두 개 있으면 어느 쪽이 쓰일지 보장되지 않음)
 *   ④ cropRepairDataPLog 실행 → 알림에서 옮긴 개수 확인 → '문제검색' 확인
 *   ⑤ cropStartInsert 로 남은 대기 행 이어서 처리
 */

const CROP_SOURCE_SHEETS = ['[풀세트]', '[써킷]', '[문항]'];
const CROP_DATA_SHEET = 'Data';

/* ═══════════════════════════════════════════════════════════════
 * 문항ID → 원본 시트 P열 셀
 * ═══════════════════════════════════════════════════════════════ */
let CROP_SRC_ID_MAP_ = null;   // 실행마다 새로 만든다 (트리거 실행 간에는 유지되지 않음)

function cropBuildSourceIdMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = {};
  CROP_SOURCE_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const last = sh.getLastRow();
    if (last < 2) return;
    sh.getRange(2, 1, last - 1, 1).getValues().forEach((r, i) => {
      const id = String(r[0] || '').trim();
      if (!id) return;
      if (map[id]) { map[id].dup = true; return; }
      map[id] = { sheet: sh, row: i + 2, dup: false };
    });
  });
  return map;
}

/** 문항ID 로 원본 시트의 P열 셀을 돌려준다. 없거나 ID가 중복이면 null */
function cropSourcePCell_(itemId) {
  if (!CROP_SRC_ID_MAP_) CROP_SRC_ID_MAP_ = cropBuildSourceIdMap_();
  const hit = CROP_SRC_ID_MAP_[String(itemId || '').trim()];
  if (!hit || hit.dup) return null;
  return hit.sheet.getRange(hit.row, CROP_CFG.COL_PLOG);
}

function cropAppendLine_(cell, line) {
  const prev = String(cell.getValue() || '').replace(/\s+$/g, '');
  cell.setValue(prev ? prev + '\n' + line : line);
}

/* ═══════════════════════════════════════════════════════════════
 * cropProcessItem_ 교체본 — P열 로그를 원본 시트에 쓴다
 *   (기존 파일의 같은 이름 함수는 지울 것)
 * ═══════════════════════════════════════════════════════════════ */
function cropProcessItem_(row, targetSheets, tz) {
  const fileKey = String(row[CROP_IDX.KEY - 1]);
  const itemId  = String(row[CROP_IDX.ITEM_ID - 1] || '').trim();
  const docUrl  = String(row[CROP_IDX.DOC_URL - 1] || '').trim();
  const probId  = String(row[CROP_IDX.PROB_FILE - 1] || '').trim();
  const solIds  = String(row[CROP_IDX.SOL_FILES - 1] || '').split(',').map(s => s.trim()).filter(Boolean);

  if (!docUrl) throw new Error('문서URL 없음');
  if (!probId) throw new Error('문제파일ID 없음');
  const docId = cropExtractDocId_(docUrl);
  if (!docId) throw new Error('문서 ID 추출 실패: ' + docUrl);

  // Data 시트(배열 수식)에는 절대 쓰지 않는다. 문항ID 로 원본 시트 행을 찾는다.
  let pCell = null;
  if (CROP_CFG.WRITE_LOG || CROP_CFG.WRITE_REG) {
    pCell = cropSourcePCell_(itemId);
    if (!pCell) throw new Error(`원본 시트(${CROP_SOURCE_SHEETS.join(', ')})에서 문항ID '${itemId}' 를 찾지 못했거나 중복입니다`);
  }
  // 삽입 완료 표시: 같은 행 O열(등록) = 1  (2026-11-09)
  const markReg = () => { if (CROP_CFG.WRITE_REG && pCell) pCell.offset(0, CROP_CFG.COL_REG - CROP_CFG.COL_PLOG).setValue(1); };

  const logLine = `[${Utilities.formatDate(new Date(), tz, 'yy-MM-dd')}] ${CROP_CFG.LOG_TAG}: ${fileKey} (문제 1 + 해설 ${solIds.length})`;

  // 이미 넣은 문항이면 건너뜀 (P열 로그로 2차 확인, 구 파일키 로그도 인정)
  if (CROP_CFG.WRITE_LOG && pCell && cropHasLog_(String(pCell.getValue() || ''), fileKey)) {
    if (!CROP_CFG.DRY_RUN) markReg();
    return { status: '✅ 이미삽입됨', msg: 'P열에 동일 로그 있어 건너뜀' };
  }

  const blobs = [DriveApp.getFileById(probId).getBlob()]
    .concat(solIds.map(id => DriveApp.getFileById(id).getBlob()));

  if (CROP_CFG.DRY_RUN) {
    DocumentApp.openById(docId);   // 열 수 있는지만 확인
    return { status: '🧪 DRY_RUN', msg: `삽입 예정 이미지 ${blobs.length}장 (${blobs.map(b => b.getName()).join(', ')})` };
  }

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  let at = 0;
  if (CROP_CFG.ADD_CAPTION) {
    body.insertParagraph(at++, `▣ ${fileKey.replace('_', ' ')}`).setBold(true);
  }
  blobs.forEach(b => {
    const img = body.insertImage(at++, b);        // 새 문단으로 삽입됨 (문제 → 해설 c1 → c2 …)
    cropFitWidth_(img);
  });
  if (CROP_CFG.WRITE_LOG) body.appendParagraph(logLine);
  doc.saveAndClose();

  if (CROP_CFG.WRITE_LOG && pCell) cropAppendLine_(pCell, logLine);
  markReg();

  return { status: '✅', msg: `이미지 ${blobs.length}장 삽입` };
}

/* ═══════════════════════════════════════════════════════════════
 * 1회용 복구: Data P열 → 원본 시트 P열로 이동, Data P열 비우기
 * ═══════════════════════════════════════════════════════════════ */
function cropRepairDataPLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(CROP_DATA_SHEET);
  if (!data) { cropAlert_(`'${CROP_DATA_SHEET}' 시트가 없습니다.`); return; }

  // 삽입이 돌고 있으면 먼저 멈추게 한다 (돌면서 Data 에 계속 쓰기 때문)
  if (PropertiesService.getScriptProperties().getProperty(CROP_CFG.PROP_RUNNING) ||
      ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === CROP_CFG.TRIGGER_HANDLER)) {
    cropAlert_('CroP 삽입이 아직 진행(예약) 중입니다. cropStopInsert 를 먼저 실행한 뒤 다시 실행하세요.');
    return;
  }

  // 배열이 막혀 있을 때만 동작 (배열이 정상이면 P열 값은 원본에서 펼쳐진 값이라 건드리면 안 됨)
  const a2 = data.getRange('A2');
  const formula = a2.getFormula();
  if (!formula) { cropAlert_('Data!A2 에 수식이 없습니다. 중단합니다.'); return; }
  if (a2.getDisplayValue() !== '#REF!') {
    cropAlert_('Data!A2 배열 수식이 정상입니다(#REF! 아님). 복구할 것이 없습니다.');
    return;
  }

  // 배열 수식에서 Data 행 → 원본 시트 행 대응을 읽는다 (ID로 못 찾을 때 대비)
  const segs = [];
  let dataRow = 2;
  const re = /'([^']+)'!(\d+):(\d+)/g;
  let m;
  while ((m = re.exec(formula)) !== null) {
    const from = +m[2], to = +m[3];
    segs.push({ name: m[1], dataFrom: dataRow, srcFrom: from, len: to - from + 1 });
    dataRow += to - from + 1;
  }
  const rowToSource = r => {
    const s = segs.find(g => r >= g.dataFrom && r < g.dataFrom + g.len);
    if (!s) return null;
    const sh = ss.getSheetByName(s.name);
    return sh ? sh.getRange(s.srcFrom + (r - s.dataFrom), CROP_CFG.COL_PLOG) : null;
  };

  // 색인: 파일키 → 문항ID
  const keyToId = {};
  const idx = ss.getSheetByName(CROP_CFG.INDEX_SHEET);
  if (idx && idx.getLastRow() > 1) {
    idx.getRange(2, 1, idx.getLastRow() - 1, CROP_IDX.ITEM_ID).getValues().forEach(r => {
      const k = String(r[CROP_IDX.KEY - 1] || '').trim();
      const id = String(r[CROP_IDX.ITEM_ID - 1] || '').trim();
      if (k && id) keyToId[k] = id;
    });
  }

  CROP_SRC_ID_MAP_ = cropBuildSourceIdMap_();
  const last = data.getLastRow();
  if (last < 2) { cropAlert_('Data 시트에 행이 없습니다.'); return; }
  const pRange = data.getRange(2, CROP_CFG.COL_PLOG, last - 1, 1);
  const pVals = pRange.getValues();

  const tagRe = new RegExp(CROP_CFG.LOG_TAG.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ':\\s*(\\S+)');
  let cells = 0, moved = 0, skipped = 0, byRow = 0;
  const unresolved = [];
  const backup = [['Data행', '원래 값']];

  pVals.forEach((r, i) => {
    const v = String(r[0] || '');
    if (!v.trim()) return;
    cells++;
    const rowNo = i + 2;
    backup.push([rowNo, v]);
    v.split('\n').map(s => s.trim()).filter(Boolean).forEach(line => {
      const mm = line.match(tagRe);
      if (!mm) return;                       // CroP 로그가 아닌 줄(원본에서 펼쳐졌던 이력 사본)은 무시
      const key = mm[1];
      let cell = cropSourcePCell_(keyToId[key]);
      if (!cell) { cell = rowToSource(rowNo); if (cell) byRow++; }
      if (!cell) { unresolved.push(`Data ${rowNo}행: ${line}`); return; }
      if (cropHasLog_(String(cell.getValue() || ''), key)) { skipped++; return; }
      cropAppendLine_(cell, line);
      moved++;
    });
  });

  if (unresolved.length) {
    cropAlert_(`옮길 곳을 못 찾은 로그 ${unresolved.length}개가 있어 Data P열을 비우지 않았습니다.\n` +
      unresolved.slice(0, 15).join('\n') + (unresolved.length > 15 ? '\n…' : ''));
    return;
  }

  // 백업 시트에 원래 값을 남기고 Data P열 비우기
  const bName = '_DataP백업_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMdd_HHmm');
  const b = ss.insertSheet(bName);
  b.getRange(1, 1, backup.length, 2).setValues(backup);
  pRange.clearContent();
  SpreadsheetApp.flush();

  const ok = data.getRange('A2').getDisplayValue() !== '#REF!';
  cropAlert_(`Data P열 셀 ${cells}개 정리\n원본 시트로 옮긴 로그 ${moved}줄 (이미 있던 것 ${skipped}줄 건너뜀, 행번호로 찾은 것 ${byRow}줄)\n` +
    `원래 값 백업: '${bName}' 시트\n\nData!A2 배열 수식: ${ok ? '✅ 정상 복구' : '⚠ 아직 #REF! — 다른 열에 남은 값이 있는지 확인하세요'}`);
}