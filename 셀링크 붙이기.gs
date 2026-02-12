function addCellLinksInRange() {
  // 스프레드시트 및 시트 선택
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // 시작 행과 끝 행을 지정 (예: 2행부터 10행까지)
  var startRow = 2;
  var endRow = 47;

  // A열의 지정한 범위 가져오기
  var range = sheet.getRange(startRow, 1, endRow - startRow + 1, 1);

  // 각 셀에 대해 링크 추가
  range.getValues().forEach((row, index) => {
    var rowIndex = startRow + index;
    var cell = sheet.getRange(rowIndex, 1);
    var cellLink = spreadsheet.getUrl() + "#gid=" + sheet.getSheetId() + "&range=" + cell.getA1Notation();
    
    // 셀에 링크 추가
    cell.setFormula('=HYPERLINK("' + cellLink + '", "' + cell.getValue() + '")');
  });
}