/**
 * 구글드라이브 폴더 내 파일명 일괄 변환 스크립트
 * 파일명 형식: D + 4자리 숫자 (예: D1345, D1346, ...)
 * 정렬: 파일명 기준 내림차순 → 시작번호부터 1씩 증가
 */

const TARGET_FOLDER_ID = "1I1h83bHIxk9pzvSRV3YgdUee4E9k9MHb";

function renameFilesInFolder() {
  const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const files = folder.getFiles();

  // 1) 파일 목록 수집
  const fileList = [];
  while (files.hasNext()) {
    const file = files.next();
    fileList.push(file);
  }

  const totalCount = fileList.length;

  if (totalCount === 0) {
    SpreadsheetApp.getUi().alert("폴더에 파일이 없습니다.");
    return;
  }

  // 2) Alert로 총 개수 안내 + 시작번호 입력
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "파일명 일괄 변환",
    `폴더 내 파일 총 ${totalCount}개가 있습니다.\n\n시작번호 4자리를 입력하세요 (예: 1345):`,
    ui.ButtonSet.OK_CANCEL
  );

  // 취소 처리
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("작업이 취소되었습니다.");
    return;
  }

  const inputText = response.getResponseText().trim();

  // 입력값 검증: 4자리 숫자인지 확인
  if (!/^\d{4}$/.test(inputText)) {
    ui.alert("오류: 4자리 숫자를 정확히 입력해주세요. (예: 1345)");
    return;
  }

  let startNumber = parseInt(inputText, 10);

  // 3) 파일명 기준 내림차순 정렬
  fileList.sort((a, b) => {
    const nameA = a.getName().toLowerCase();
    const nameB = b.getName().toLowerCase();
    if (nameA > nameB) return -1;
    if (nameA < nameB) return 1;
    return 0;
  });

  // 4) 파일명 변환 실행
  const results = [];
  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < fileList.length; i++) {
    const file = fileList[i];
    const oldName = file.getName();
    const extension = getExtension(oldName);
    const currentNumber = startNumber + i;
    const newName = "D" + String(currentNumber).padStart(4, "0") + extension;

    try {
      file.setName(newName);
      results.push(`✓ ${oldName} → ${newName}`);
      successCount++;
    } catch (e) {
      results.push(`✗ ${oldName} → 변환 실패: ${e.message}`);
      errorCount++;
    }
  }

  // 5) 결과 알림
  const summary = [
    "═══ 파일명 변환 완료 ═══",
    "",
    `총 파일 수: ${totalCount}`,
    `성공: ${successCount}개`,
    `실패: ${errorCount}개`,
    `번호 범위: D${String(startNumber).padStart(4, "0")} ~ D${String(startNumber + totalCount - 1).padStart(4, "0")}`,
    "",
    "── 변환 내역 ──",
    ...results,
  ].join("\n");

  ui.alert("파일명 변환 결과", summary, ui.ButtonSet.OK);
}

/**
 * 파일명에서 확장자 추출 (점 포함)
 * 확장자가 없으면 빈 문자열 반환
 */
function getExtension(filename) {
  const dotIndex = filename.lastIndexOf(".");
  if (dotIndex === -1 || dotIndex === 0) return "";
  return filename.substring(dotIndex);
}