function onInstall() { onOpen(); }

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🥑 문항관리 메뉴');


  // ---- 그룹  : 구글문서 만들기 --
  menu.addItem('ID 자동생성','generateIdsInColumnA');
  menu.addItem('문항구글문서 일괄생성', 'startCreateDocs');
  menu.addItem('셀 링크 붙이기', 'addCellLinksInRange');
  menu.addSeparator();
  menu.addItem('INSERTIMG 파일명 일괄변경 (D+4자리)','renameFilesInFolder');

  // ---- 그룹  : 구글문서에 입력하기 ---
  menu.addSeparator();
  menu.addItem('구글문서/시트 일괄입력 : 그림 & 문구', 'insertImageAndLogToDocs');
  menu.addItem('구글문서/시트 일괄입력 : 문구만', 'insertTextIntoGoogleDocs');
  menu.addItem('구글문서 일괄비우기','clearDocsByRange');
  


  // ---- 그룹 : 문제 이동 ----
  menu.addSeparator();
  menu.addItem('문제 일괄이동', 'moveItemsAndLog');
  menu.addItem('문제이동 초기화', 'clearTUVColumns');
  menu.addItem("구글문서 '페이지 없음' 일괄설정", 'setPagelessWithPrompt');
  menu.addItem("이동이력 보수기록", 'relogMoveHistory');
  

// ---- 그룹 : CroP 이미지 삽입 (MainMenu.gs onOpen 안, menu.addToUi() 앞에 붙여넣기) ----
  menu.addSeparator();
  menu.addItem('CroP 이미지 ⓪ 세트명 대응 점검', 'cropCheckSetNames');
  menu.addItem('CroP 이미지 ① 색인 만들기', 'cropBuildIndex');
  menu.addItem('CroP 이미지 ② 삽입 시작/재개', 'cropStartInsert');
  menu.addItem('CroP 이미지 ⏹ 중지', 'cropStopInsert');
  menu.addItem('CroP 이미지 ③ 검증 (문서 열어 확인)', 'cropVerify');
  menu.addItem('CroP 이미지 · 오류행 다시 대기로', 'cropRetryErrors');
  menu.addItem('CroP 이미지 · 중복 1벌 제거', 'cropFixDuplicates');
  menu.addToUi();
}



