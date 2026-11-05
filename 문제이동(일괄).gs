/**
* 문항이동 배치 처리 (수정본 v2)
* - '문항이동' 시트의 T(문항ID), U(세트명), V(번호)를 읽어
* - [풀세트], [써킷], [문항] 중 A열=ID인 유일한 행을 찾아
*   E열=U, G열=V로 갱신
* - 해당 행 N열의 구글문서 링크로 들어가 문서 맨 끝에 로그 1줄 추가
*   [YY-MM-DD] (이전E) (이전G)번 ⇒ (이후U) (이후V)번
* - 그리고 "문서 맨 끝에 추가된 1줄(마지막 문단)"을
*   같은 행 P열에 기존 내용 아래로 첨부
* - 완료된(성공) 행 수를 알림
*
* [v2 수정 사항]
* 1) DocumentApp.openByUrl() → 문서ID 추출 후 openById() 로 변경
*    - N열에 기록된 링크가 ".../document/d/ID" 형태(끝에 /edit 없음)라서
*      openByUrl()이 "Invalid argument: url" 오류를 던지며 실패했음.
*    - hist_copyDocText_toP.gs 와 동일하게 ID를 정규식으로 추출해 openById() 사용.
* 2) 처리 순서 변경: 문서 로그 기록에 "성공한 뒤에만" E/G열을 갱신
*    - 기존에는 E/G를 먼저 고치고 문서 기록에서 실패 → 시트만 바뀌고
*      이력은 안 남는 "반쪽 처리" 상태가 됨. 이제는 실패 시 아무것도 안 바꿈.
* 3) 실패 사유를 행별로 수집해 완료 알림창에 표시 (Logger에도 기록)
* 4) N열 링크 추출 강화: 리치텍스트 링크(부분 링크 포함), HYPERLINK 수식,
*    일반 텍스트 모두 지원
*/

const MOVE_CFG = {
  MOVE_SHEET_NAME: '문항이동',
  TARGET_SHEETS: ['[풀세트]', '[써킷]', '[문항]'],

  // 문항이동 시트 컬럼
  COL_ID: 20,      // T
  COL_USE_SET: 21, // U
  COL_USE_NO: 22,  // V

  // 대상 시트 컬럼
  TGT_COL_KEY: 1,   // A
  TGT_COL_SET: 5,   // E
  TGT_COL_NO: 7,    // G
  TGT_COL_DOC: 14,  // N
  TGT_COL_PLOG: 16, // P  생성/이동이력(시트에 붙여넣기)
};

function moveItemsAndLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const moveSheet = ss.getSheetByName(MOVE_CFG.MOVE_SHEET_NAME);
  if (!moveSheet) {
    ui.alert(`시트 "${MOVE_CFG.MOVE_SHEET_NAME}" 를 찾을 수 없습니다.`);
    return;
  }

  const startRow = promptRow_(ui, '시작 행 번호를 입력하세요');
  if (!startRow) return;

  const endRow = promptRow_(ui, '끝 행 번호를 입력하세요');
  if (!endRow) return;

  if (startRow > endRow) {
    ui.alert('시작 행이 끝 행보다 클 수 없습니다.');
    return;
  }

  const targetSheets = MOVE_CFG.TARGET_SHEETS
    .map(name => ss.getSheetByName(name))
    .filter(sh => sh);

  if (targetSheets.length === 0) {
    ui.alert('대상 시트([풀세트], [써킷], [문항])를 하나도 찾지 못했습니다.');
    return;
  }

  let successCount = 0;
  let skipCount = 0;
  const errors = []; // ✅ 행별 실패 사유 수집
  const todayTag = formatYYMMDD_(); // [YY-MM-DD]

  for (let row = startRow; row <= endRow; row++) {
    try {
      const id = String(moveSheet.getRange(row, MOVE_CFG.COL_ID).getDisplayValue()).trim();
      const newSet = String(moveSheet.getRange(row, MOVE_CFG.COL_USE_SET).getDisplayValue()).trim(); // U
      const newNo  = String(moveSheet.getRange(row, MOVE_CFG.COL_USE_NO).getDisplayValue()).trim();  // V

      if (!id || !newSet || !newNo) {
        skipCount++;
        errors.push(`${row}행: T/U/V 값 비어있음 (건너뜀)`);
        continue;
      }

      // ID 형식: D+4자리 (예: D0123)
      if (!/^D\d{4}$/.test(id)) {
        throw new Error(`ID 형식이 예상과 다름(예: D0123): ${id}`);
      }

      // 세 시트에서 A열=ID 인 유일한 행 찾기
      const hit = findUniqueRowById_(targetSheets, id);
      if (!hit) {
        skipCount++;
        errors.push(`${row}행(${id}): 대상 시트에서 ID를 찾지 못함`);
        continue;
      }

      const { sheet: targetSheet, rowIndex } = hit;

      // 업데이트 전에 "기존 값" 확보
      const oldSet = String(targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_SET).getDisplayValue()).trim(); // 기존 E
      const oldNo  = String(targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_NO).getDisplayValue()).trim();  // 기존 G

      // ✅ [v2] 시트를 고치기 "전에" 문서부터 확인/기록한다.
      //    (문서 기록 실패 시 시트도 건드리지 않아 반쪽 처리 방지)
      const docId = getDocIdFromCell_(targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_DOC));
      if (!docId) {
        throw new Error(`구글문서 링크(N열)에서 문서ID 추출 실패`);
      }

      // 로그 1줄 생성
      const logLine = `[${todayTag}] ${oldSet} ${oldNo}번 ⇒ ${newSet} ${newNo}번`;

      // 1) 문서 끝에 로그 추가
      // 2) 문서 마지막 문단(=방금 추가된 1줄)을 다시 읽어서 반환
      const lastLine = appendLogToDocAndGetLastLine_(docId, logLine);

      // ✅ 문서 기록이 성공했을 때만 E, G 업데이트
      targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_SET).setValue(newSet); // E
      targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_NO).setValue(newNo);   // G

      // 그 1줄을 P열에 "기존 내용 아래"로 첨부
      appendLineToCell_(targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_PLOG), lastLine);

      successCount++;
    } catch (err) {
      const msg = err && err.message ? err.message : String(err);
      Logger.log(`Row ${row} 실패: ${msg}`);
      errors.push(`${row}행: ${msg}`);
      skipCount++;
    }
  }

  // ✅ 실패 사유를 알림창에 함께 표시 (최대 10건)
  let report = `완료!\n성공: ${successCount}행\n스킵/실패: ${skipCount}행`;
  if (errors.length > 0) {
    const shown = errors.slice(0, 10);
    report += `\n\n[실패 사유]\n` + shown.join('\n');
    if (errors.length > 10) {
      report += `\n... 외 ${errors.length - 10}건 (실행기록 로그 참조)`;
    }
  }
  ui.alert(report);
}

/** ---------- helpers ---------- */

function promptRow_(ui, message) {
  const res = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  const n = parseInt(String(res.getResponseText()).trim(), 10);
  if (!Number.isFinite(n) || n < 1) {
    ui.alert('유효한 행 번호를 입력하세요.');
    return null;
  }
  return n;
}

function findUniqueRowById_(targetSheets, id) {
  let found = null;

  for (const sh of targetSheets) {
    const finder = sh.getRange(1, MOVE_CFG.TGT_COL_KEY, sh.getMaxRows(), 1)
      .createTextFinder(id)
      .matchEntireCell(true);

    const cell = finder.findNext();
    if (cell) {
      if (found) throw new Error(`ID 중복 발견: "${id}" (${found.sheet.getName()} / ${sh.getName()})`);
      found = { sheet: sh, rowIndex: cell.getRow() };
    }
  }

  return found;
}

/**
 * ✅ [v2] 셀에서 구글문서 "ID"를 추출한다.
 * 지원 형태:
 *  - 셀 전체 리치텍스트 링크
 *  - 부분 텍스트에만 걸린 리치텍스트 링크 (runs 순회)
 *  - =HYPERLINK("url","라벨") 수식
 *  - 일반 텍스트 URL (".../document/d/ID" 또는 ".../open?id=ID")
 */
function getDocIdFromCell_(range) {
  const candidates = [];

  // 1) 리치텍스트 링크
  try {
    const rt = range.getRichTextValue();
    if (rt) {
      const link = rt.getLinkUrl();
      if (link) candidates.push(link);
      for (const run of rt.getRuns()) {
        const runLink = run.getLinkUrl();
        if (runLink) candidates.push(runLink);
      }
    }
  } catch (e) {}

  // 2) HYPERLINK 수식
  try {
    const f = String(range.getFormula() || '');
    const fm = f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
    if (fm) candidates.push(fm[1]);
  } catch (e) {}

  // 3) 일반 텍스트
  candidates.push(String(range.getDisplayValue() || ''));

  for (const c of candidates) {
    const id = extractDocIdFromUrl_(c);
    if (id) return id;
  }
  return null;
}

// URL 문자열에서 문서 ID 추출
function extractDocIdFromUrl_(url) {
  const s = String(url || '').trim();
  if (!s) return null;
  let m = s.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return null;
}

/**
 * 문서 맨 끝에 line을 추가하고,
 * 문서의 "마지막 문단 텍스트"를 반환한다.
 * ✅ [v2] openByUrl(url) → openById(docId) 로 변경
 *    (N열의 ".../document/d/ID" 형태 링크는 /edit 이 없어
 *     openByUrl 이 Invalid argument 오류를 던졌음)
 */
function appendLogToDocAndGetLastLine_(docId, line) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  body.appendParagraph(line);

  // 저장 전에 마지막 문단 읽기 (저장/닫기는 뒤에서)
  const n = body.getNumChildren();
  let lastText = line;

  if (n > 0) {
    const last = body.getChild(n - 1);
    // Paragraph/Text 둘 다 대응
    try {
      lastText = String(last.asParagraph().getText() || '').trim() || line;
    } catch (e) {
      try {
        lastText = String(last.getText() || '').trim() || line;
      } catch (e2) {
        lastText = line;
      }
    }
  }

  doc.saveAndClose();
  return lastText;
}

/**
 * 셀에 기존 내용 아래로 줄바꿈 후 line을 추가
 */
function appendLineToCell_(range, line) {
  const cur = String(range.getDisplayValue() || '').trim();
  const add = String(line || '').trim();
  if (!add) return;

  if (!cur) {
    range.setValue(add);
  } else {
    range.setValue(cur + '\n' + add);
  }
}

// YY-MM-DD (예: 25-01-26)
function formatYYMMDD_() {
  const d = new Date();
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yy}-${mm}-${dd}`;
}