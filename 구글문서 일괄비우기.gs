/******************************************************
 * 구글문서 내용 전체 삭제 자동화
 ******************************************************/
function clearDocsByRange() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  // === 1. 입력받기 ===
  const res = ui.prompt('첫 문항번호와 끝 문항번호를 입력하세요', '예: 0001,0010', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const [startNum, endNum] = res.getResponseText().split(',').map(s => s.trim());

  if (!startNum || !endNum) {
    ui.alert('⚠️ 입력 형식이 잘못되었습니다. 예: 0001,0010');
    return;
  }

  // === 2. 데이터 범위 ===
  const data = sheet.getDataRange().getValues();
  let successCount = 0;

  // === 3. 각 행 반복 ===
  for (let i = 1; i < data.length; i++) {
    const idStr = data[i][0]; // A열 (예: D0001)
    const link = data[i][13]; // N열 (0부터 시작하므로 13번째가 N열)
    if (!idStr || !link) continue;

    const num = idStr.replace(/^D/, '');
    if (num < startNum || num > endNum) continue;

    try {
      const docIdMatch = link.match(/[-\w]{25,}/);
      if (!docIdMatch) throw new Error('문서 ID를 찾을 수 없음');

      const doc = DocumentApp.openById(docIdMatch[0]);
      const body = doc.getBody();

      // === 3-1. 문서 내용 전체 삭제 ===
      body.clear(); // 문서 전체 내용 제거

      doc.saveAndClose();
      successCount++;

    } catch (err) {
      Logger.log(`오류 (${idStr}): ${err}`);
    }
  }

  // === 4. 완료 메시지 ===
  ui.alert(`✅ 완료: ${successCount}개의 문서 내용을 모두 삭제했습니다.`);
}
