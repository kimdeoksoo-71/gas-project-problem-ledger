function clearTUVColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return; // 2행부터니까 2행보다 작으면 아무것도 안함

  // T열(20), U열(21), V열(22) - 2행부터 마지막행까지
  sheet.getRange(2, 20, lastRow - 1, 3).clearContent();
}
