function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setActiveSelection("A1"); // A1 셀을 선택하도록 설정
}
