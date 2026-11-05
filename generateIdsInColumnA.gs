/**
 * A열 ID(D 또는 S + 4자리) 자동 생성
 * - 첫 번째 ID(예: 1001 · D1001 · S0001)와 새로 만들 행의 개수를 입력받아
 * - 활성 시트의 마지막 행 아래부터 A열에 ID를 순서대로 채움
 *   숫자만 넣으면 D (기존 방식 그대로), S로 시작하면 S
 * - 활성 시트와 [풀세트]·[써킷]·[문항] A열에 이미 있는 ID와 중복되면 실행 전에 경고 후 중단
 *   (문제이동 등은 세 시트에서 ID가 유일하다고 가정한다)
 */
const GEN_ID_SHEETS = ['[풀세트]', '[써킷]', '[문항]'];

function generateIdsInColumnA() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // === 1) 첫 번째 ID 입력 ===
  const resStart = ui.prompt(
    'ID 자동 생성',
    '첫 번째 ID를 입력하세요 (예: 1001 또는 D1001 → D, S0001 → S)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resStart.getSelectedButton() !== ui.Button.OK) return;

  const startText = String(resStart.getResponseText()).trim().toUpperCase();
  const sm = startText.match(/^([DS])?(\d{1,4})$/);
  if (!sm) {
    ui.alert('⚠️ 0~9999 사이의 숫자(앞에 D 또는 S 가능)를 입력하세요. (예: 1001, D1001, S0001)');
    return;
  }
  const prefix = sm[1] || 'D';
  const startNum = parseInt(sm[2], 10);
  const fmt = n => `${prefix}${String(n).padStart(4, '0')}`;

  // === 2) 새로 만들 행의 개수 입력 ===
  const resCount = ui.prompt(
    'ID 자동 생성',
    '새로 만들 행의 개수를 입력하세요 (예: 50)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resCount.getSelectedButton() !== ui.Button.OK) return;

  const count = parseInt(String(resCount.getResponseText()).trim(), 10);
  if (!Number.isFinite(count) || count < 1) {
    ui.alert('⚠️ 1 이상의 개수를 입력하세요.');
    return;
  }

  // 4자리 범위 초과 확인 (예: 9990에서 20개 → D10009 불가)
  const endNum = startNum + count - 1;
  if (endNum > 9999) {
    ui.alert(`⚠️ 마지막 번호가 9999를 넘습니다.\n(시작: ${startNum}, 개수: ${count} → 마지막: ${endNum})`);
    return;
  }

  // === 3) 기존 A열 ID와 중복 확인 (활성 시트 + 세 원본 시트) ===
  const existingSet = new Set();
  const sheetsToCheck = [sheet].concat(
    GEN_ID_SHEETS.map(n => ss.getSheetByName(n)).filter(sh => sh && sh.getSheetId() !== sheet.getSheetId())
  );
  sheetsToCheck.forEach(sh => {
    const lr = sh.getLastRow();
    if (lr < 1) return;
    sh.getRange(1, 1, lr, 1).getDisplayValues()
      .map(r => String(r[0]).trim().toUpperCase())
      .filter(v => /^[DS]\d{4}$/.test(v))
      .forEach(v => existingSet.add(v));
  });

  const dupes = [];
  for (let n = startNum; n <= endNum; n++) {
    const id = fmt(n);
    if (existingSet.has(id)) dupes.push(id);
  }
  if (dupes.length > 0) {
    ui.alert(`⚠️ 이미 A열에 존재하는 ID가 있어 중단합니다.\n중복: ${dupes.slice(0, 10).join(', ')}${dupes.length > 10 ? ` 외 ${dupes.length - 10}개` : ''}`);
    return;
  }

  // === 4) 마지막 행 아래부터 ID 일괄 입력 ===
  const lastRow = sheet.getLastRow();
  const values = [];
  for (let n = startNum; n <= endNum; n++) {
    values.push([fmt(n)]);
  }

  const writeStartRow = lastRow + 1;
  sheet.getRange(writeStartRow, 1, count, 1).setValues(values);

  ui.alert(
    `완료!\n` +
    `시트: "${sheet.getName()}"\n` +
    `행 ${writeStartRow} ~ ${writeStartRow + count - 1}\n` +
    `ID: ${fmt(startNum)} ~ ${fmt(endNum)} (총 ${count}개)`
  );
}
