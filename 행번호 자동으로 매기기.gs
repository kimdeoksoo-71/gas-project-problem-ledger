function generateSerialNumbers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 600; // A열의 일련번호를 붙이기 시작하는 행의 행번호
  const startNumber = 4000; // D1001부터 시작할 숫자
  const lastRow = sheet.getLastRow(); // 마지막 행

  for (let i = startRow; i <= lastRow; i++) {
    const serialNumber = `D${(startNumber + i - 1).toString().padStart(4, '0')}`;
    sheet.getRange(i, 1).setValue(serialNumber); // A열에 일련번호 입력
  }
}
