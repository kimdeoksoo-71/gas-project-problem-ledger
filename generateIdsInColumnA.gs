/**
 * A열 ID(D+4자리) 자동 생성
 * - 첫 번째 숫자(예: 1001 또는 D1001)와 새로 만들 행의 개수를 입력받아
 * - 활성 시트의 마지막 행 아래부터 A열에 D0001 형식의 ID를 순서대로 채움
 * - 기존 A열에 이미 있는 ID와 중복되면 실행 전에 경고 후 중단
 */
function generateIdsInColumnA() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // === 1) 첫 번째 숫자 입력 ===
  const resStart = ui.prompt(
    'ID 자동 생성',
    '첫 번째 숫자를 입력하세요 (예: 1001 또는 D1001)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resStart.getSelectedButton() !== ui.Button.OK) return;

  const startText = String(resStart.getResponseText()).trim().toUpperCase();
  const startNum = parseInt(startText.replace(/^D/, ''), 10);

  if (!Number.isFinite(startNum) || startNum < 0 || startNum > 9999) {
    ui.alert('⚠️ 0~9999 사이의 숫자를 입력하세요. (예: 1001 또는 D1001)');
    return;
  }

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

  // === 3) 기존 A열 ID와 중복 확인 ===
  const lastRow = sheet.getLastRow();
  if (lastRow >= 1) {
    const existing = sheet.getRange(1, 1, lastRow, 1).getDisplayValues()
      .map(r => String(r[0]).trim().toUpperCase())
      .filter(v => /^D\d{4}$/.test(v));
    const existingSet = new Set(existing);

    const dupes = [];
    for (let n = startNum; n <= endNum; n++) {
      const id = `D${String(n).padStart(4, '0')}`;
      if (existingSet.has(id)) dupes.push(id);
    }
    if (dupes.length > 0) {
      ui.alert(`⚠️ 이미 A열에 존재하는 ID가 있어 중단합니다.\n중복: ${dupes.slice(0, 10).join(', ')}${dupes.length > 10 ? ` 외 ${dupes.length - 10}개` : ''}`);
      return;
    }
  }

  // === 4) 마지막 행 아래부터 ID 일괄 입력 ===
  const values = [];
  for (let n = startNum; n <= endNum; n++) {
    values.push([`D${String(n).padStart(4, '0')}`]);
  }

  const writeStartRow = lastRow + 1;
  sheet.getRange(writeStartRow, 1, count, 1).setValues(values);

  ui.alert(
    `완료!\n` +
    `시트: "${sheet.getName()}"\n` +
    `행 ${writeStartRow} ~ ${writeStartRow + count - 1}\n` +
    `ID: D${String(startNum).padStart(4, '0')} ~ D${String(endNum).padStart(4, '0')} (총 ${count}개)`
  );
}