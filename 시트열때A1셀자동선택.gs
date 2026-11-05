/**
 * ⚠️ 이 함수는 원래 이름이 onOpen 이었다. MainMenu.gs 의 onOpen(메뉴 만들기)과 이름이 겹쳐
 * Apps Script가 파일 순서상 뒤의 것 하나만 실행했고, 지금까지는 메뉴 쪽이 실행되고 있었다
 * (메뉴가 보였으므로). clasp push로 파일 순서가 바뀌면 메뉴가 사라질 수 있어 이름을 바꾼다.
 * 지금 동작(시트를 열 때 A1 선택 안 함)을 그대로 두려고 어디서도 부르지 않는다.
 * A1 자동 선택이 필요하면 MainMenu.gs onOpen 끝에서 selectA1OnOpen_() 을 부르면 된다.
 */
function selectA1OnOpen_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setActiveSelection("A1"); // A1 셀을 선택하도록 설정
}
