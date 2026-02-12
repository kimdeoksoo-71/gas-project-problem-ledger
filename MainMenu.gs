function onInstall() { onOpen(); }

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🥑 문항관리 메뉴');


  // ---- 그룹 1 : 구글문서 만들기 --
  menu.addItem('문항구글문서 일괄생성', 'startCreateDocs');
  menu.addItem('구글문서 링크 일괄추출 (빈 시트에서)','insertTextIntoGoogleDocs')
  menu.addItem('셀 링크 붙이기', 'addCellLinksInRange')

  // ---- 그룹 2 : 구글문서에 입력하기 ---
  menu.addSeparator();
  menu.addItem('구글문서 일괄입력 : 그림 & 문구', 'insertImageAndLogToDocs');
  menu.addItem('구글문서 일괄입력 : 문구만', 'insertTextIntoGoogleDocs');
  menu.addItem('구글문서 일괄비우기','clearDocsByRange');
  

  
 
  // ---- 그룹 2: 문제 이동 ----
  menu.addSeparator();
  menu.addItem('문제 일괄이동', 'moveItemsAndLog');
  menu.addItem('문제이동 초기화', 'clearTUVColumns');

  

  // ---- 메뉴 완성 ----
  menu.addToUi();
}



