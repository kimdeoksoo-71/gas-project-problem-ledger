/**************************************************
 * 이동이력 보수(補修) 기록기 (v2)
 *
 * [용도]
 * '문제 일괄이동' 실행 중 E/G열(시트)만 바뀌고
 * 구글문서 이력 기록이 누락된 행들에 대한 간이 보수.
 *
 * 전제: 해당 문서들의 "마지막 행"에는 이동 전의
 *       세트명·번호가 이미 기록되어 있음.
 *       따라서 그 바로 아래에
 *       "⇒ (현재 E열값) (현재 G열값)번"
 *       한 줄만 덧붙이면 이력이 완성됨.
 *
 * [사용법]
 * 1. '문항이동' 시트 T열에 문항 ID만 기록 (U, V열은 비워도 됨)
 * 2. relogMoveHistory 실행
 * 3. 입력창 하나에 "시작행-끝행" 입력 (예: 2-42)
 * 4. 각 행마다:
 *    - [풀세트]/[써킷]/[문항]에서 A열=ID인 행을 찾고
 *    - N열 링크로 구글문서를 열어 맨 끝에 새 행을 만든 뒤
 *    - "[yy-mm-dd] ⇒ (현재 E열값) (현재 G열값)번" 을 입력
 *      ※ E/G열 = "대상 시트([풀세트]/[써킷]/[문항])"의 E/G열
 *        ('문항이동' 시트의 E/G열이 아님)
 *    - 같은 줄을 대상 시트 P열에도 기존 내용 아래로 첨부
 *
 * [안전장치]
 * - 문서 마지막 줄이 이미 같은 내용이면 건너뜀 (중복 실행 방지, P열도 안 건드림)
 * - E/G열은 절대 건드리지 않음 (문서와 P열에만 기록)
 *
 * ※ 문제이동(일괄).gs (v2) 의 MOVE_CFG, findUniqueRowById_,
 *    getDocIdFromCell_, appendLineToCell_, formatYYMMDD_ 를 그대로 사용한다.
 **************************************************/

function relogMoveHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const moveSheet = ss.getSheetByName(MOVE_CFG.MOVE_SHEET_NAME);
  if (!moveSheet) {
    ui.alert(`시트 "${MOVE_CFG.MOVE_SHEET_NAME}" 를 찾을 수 없습니다.`);
    return;
  }

  // ✅ 시작행/끝행을 입력창 하나로 받기 (예: 2-42)
  const range = promptRowRange_(ui);
  if (!range) return;
  const { startRow, endRow } = range;

  const targetSheets = MOVE_CFG.TARGET_SHEETS
    .map(name => ss.getSheetByName(name))
    .filter(sh => sh);

  if (targetSheets.length === 0) {
    ui.alert('대상 시트([풀세트], [써킷], [문항])를 하나도 찾지 못했습니다.');
    return;
  }

  let successCount = 0;
  let skipCount = 0;
  const errors = [];
  const todayTag = formatYYMMDD_(); // [yy-mm-dd] 날짜 태그 (문제이동(일괄).gs의 헬퍼)

  for (let row = startRow; row <= endRow; row++) {
    try {
      const id = String(moveSheet.getRange(row, MOVE_CFG.COL_ID).getDisplayValue()).trim(); // T

      if (!id) {
        skipCount++;
        errors.push(`${row}행: T열 ID 비어있음 (건너뜀)`);
        continue;
      }
      if (!/^[DS]\d{4}$/.test(id)) {
        throw new Error(`ID 형식이 예상과 다름(예: D0123, S0001): ${id}`);
      }

      // 세 시트에서 A열=ID 인 유일한 행 찾기
      const hit = findUniqueRowById_(targetSheets, id);
      if (!hit) {
        skipCount++;
        errors.push(`${row}행(${id}): 대상 시트에서 ID를 찾지 못함`);
        continue;
      }

      const { sheet: targetSheet, rowIndex } = hit;

      // 현재 E열(세트), G열(번호) 값 읽기
      const curSet = String(targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_SET).getDisplayValue()).trim();
      const curNo  = String(targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_NO).getDisplayValue()).trim();

      if (!curSet || !curNo) {
        skipCount++;
        errors.push(`${row}행(${id}): 대상 행의 E열 또는 G열이 비어있음 (건너뜀)`);
        continue;
      }

      // N열 문서 링크 → 문서 ID
      const docId = getDocIdFromCell_(targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_DOC));
      if (!docId) {
        throw new Error(`구글문서 링크(N열)에서 문서ID 추출 실패`);
      }

      // ✅ 기록할 한 줄: [yy-mm-dd] ⇒ (현재 E열값) (현재 G열값)번
      const newLine = `[${todayTag}] ⇒ ${curSet} ${curNo}번`;

      // 문서 열어서 마지막 줄 확인 후, 바로 아래에 새 행 추가
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();

      // 중복 실행 방지: 마지막 비어있지 않은 줄이 이미 같으면 건너뜀
      if (lastNonEmptyLine_(body) === newLine) {
        doc.saveAndClose();
        skipCount++;
        errors.push(`${row}행(${id}): 이미 같은 줄이 문서 끝에 있음 (건너뜀)`);
        continue;
      }

      body.appendParagraph(newLine);
      doc.saveAndClose();

      // ✅ 같은 줄을 대상 시트 P열에도 기존 내용 아래로 첨부
      appendLineToCell_(targetSheet.getRange(rowIndex, MOVE_CFG.TGT_COL_PLOG), newLine);

      successCount++;
    } catch (err) {
      const msg = err && err.message ? err.message : String(err);
      Logger.log(`Row ${row} 실패: ${msg}`);
      errors.push(`${row}행: ${msg}`);
      skipCount++;
    }
  }

  let report = `보수기록 완료!\n성공: ${successCount}행\n스킵/실패: ${skipCount}행`;
  if (errors.length > 0) {
    const shown = errors.slice(0, 10);
    report += `\n\n[상세]\n` + shown.join('\n');
    if (errors.length > 10) {
      report += `\n... 외 ${errors.length - 10}건 (실행기록 로그 참조)`;
    }
  }
  ui.alert(report);
}

/** ---------- helpers ---------- */

// "2-42" / "2~42" / "2,42" / "2 42" 형식을 모두 허용
function promptRowRange_(ui) {
  const res = ui.prompt(
    '보수기록 실행 범위',
    '시작행-끝행 을 입력하세요 (예: 2-42)',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return null;

  const raw = String(res.getResponseText()).trim();
  const m = raw.match(/^(\d+)\s*[-~,\s]\s*(\d+)$/);
  if (!m) {
    ui.alert(`입력 형식이 올바르지 않습니다: "${raw}"\n예: 2-42`);
    return null;
  }

  const startRow = parseInt(m[1], 10);
  const endRow = parseInt(m[2], 10);

  if (startRow < 1 || endRow < startRow) {
    ui.alert('시작 행이 1 이상이고, 끝 행은 시작 행보다 크거나 같아야 합니다.');
    return null;
  }

  return { startRow, endRow };
}

// 문서 본문의 마지막 비어있지 않은 줄 텍스트 (없으면 '')
function lastNonEmptyLine_(body) {
  for (let i = body.getNumChildren() - 1; i >= 0; i--) {
    const el = body.getChild(i);
    let t = '';
    try { t = String(el.asParagraph().getText() || '').trim(); }
    catch (e) { try { t = String(el.getText() || '').trim(); } catch (e2) {} }
    if (t) return t;
  }
  return '';
}