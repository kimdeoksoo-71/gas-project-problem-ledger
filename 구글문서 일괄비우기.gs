/******************************************************
 * 구글문서 내용 전체 삭제 자동화
 ******************************************************/
function clearDocsByRange() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  // === 1. 입력받기 ===
  // 숫자만 넣으면 D (기존 입력 방식 그대로). S 문서는 S0001,S0010 처럼 접두어를 붙인다.
  const res = ui.prompt('첫 문항ID와 끝 문항ID를 입력하세요', 'D 문서: 0001,0010  (또는 D0001,D0010)\nS 문서: S0001,S0010', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const parts = res.getResponseText().split(',').map(s => s.trim().toUpperCase().match(/^([DS])?(\d{1,4})$/));
  const prefix = parts[0] ? (parts[0][1] || 'D') : '';
  const prefix2 = parts[1] ? (parts[1][1] || 'D') : '';
  const startNum = parts[0] ? parseInt(parts[0][2], 10) : NaN;
  const endNum = parts[1] ? parseInt(parts[1][2], 10) : NaN;

  if (parts.length !== 2 || !Number.isFinite(startNum) || !Number.isFinite(endNum) || startNum > endNum || prefix !== prefix2) {
    ui.alert('⚠️ 입력 형식이 잘못되었습니다. 예: 0001,0010 또는 S0001,S0010');
    return;
  }

  // 지우기 전에 한 번 더 확인 (되돌리려면 문서 버전 기록에서 복원해야 한다)
  const pad = n => String(n).padStart(4, '0');
  const ok = ui.alert('구글문서 일괄비우기',
    `시트 "${sheet.getName()}"에서 ${prefix}${pad(startNum)} ~ ${prefix}${pad(endNum)} 행의 N열 문서 내용을 모두 지웁니다.\n계속할까요?`,
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  // === 2. 데이터 범위 ===
  const data = sheet.getDataRange().getValues();
  let successCount = 0;

  // === 3. 각 행 반복 ===
  for (let i = 1; i < data.length; i++) {
    const idStr = String(data[i][0] || '').trim(); // A열 (예: D0001 · S0001)
    const link = String(data[i][13] || '').trim(); // N열 (0부터 시작하므로 13번째가 N열)
    if (!idStr || !link) continue;

    // 접두어까지 정확히 맞는 ID만 (예전에는 D만 떼고 문자열로 비교했다)
    const m = idStr.match(/^([DS])(\d{4})$/);
    if (!m || m[1] !== prefix) continue;
    const num = parseInt(m[2], 10);
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
