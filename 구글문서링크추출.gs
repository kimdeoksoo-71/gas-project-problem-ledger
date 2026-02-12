function saveDocumentLinks() {
  // 첫 문서와 마지막 문서의 일련번호를 사용자로부터 입력받음
  const startNumber = parseInt(Browser.inputBox("첫 문서의 일련번호를 입력하세요 (예: 1001)"));
  const endNumber = parseInt(Browser.inputBox("마지막 문서의 일련번호를 입력하세요 (예: 1010)"));

  // Google Drive에서 "문항 구글문서"라는 이름의 폴더를 찾음
  const folder = DriveApp.getFoldersByName("문항 구글문서").next();
  const files = folder.getFiles();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); // 시트를 초기화하여 이전 데이터를 삭제

  const fileList = [];

  // 파일을 하나씩 확인하면서 조건에 맞는 파일의 이름과 링크를 리스트에 추가
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    // 파일명이 "D"로 시작하고 4자리 숫자로 되어 있는지 확인
    const match = fileName.match(/^D(\d{4})$/);
    if (match) {
      const fileNumber = parseInt(match[1]);

      // 입력된 범위 내에 있는 파일만 리스트에 추가
      if (fileNumber >= startNumber && fileNumber <= endNumber) {
        fileList.push({ name: fileName, url: file.getUrl() });
      }
    }
  }

  // 파일 리스트를 파일명 순으로 정렬
  fileList.sort((a, b) => a.name.localeCompare(b.name));

  // 정렬된 파일 리스트를 시트에 추가
  fileList.forEach((file, index) => {
    sheet.getRange(index + 1, 1).setValue(file.name); // A열에 파일명 추가
    sheet.getRange(index + 1, 2).setValue(file.url); // B열에 파일 링크 추가
  });

  SpreadsheetApp.flush();
  Browser.msgBox("작업이 완료되었습니다!");
}
