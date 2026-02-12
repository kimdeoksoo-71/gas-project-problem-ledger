/**************************************************
 * D0001 ~ D9999 형식의 구글문서 일괄 생성기 (시트용)
 * - 폴더: "내 드라이브/20 문항관리/문항 구글문서"
 * - 한 번에 20개씩 생성
 * - 이름 중복 시 경고 후 중단
 * - 생성 시 스프레드시트 A열에서 같은 이름을 찾아
 *   N열(14번째 열)에 구글문서 링크 기록
 * - 타임 트리거로 자동 이어서 실행
 **************************************************/

const BATCH_SIZE = 20;
const FOLDER_PATH = ['20 문항 관리', '문항 구글문서'];
const PROP_KEY = 'DOC_GENERATOR_STATE';

/**************************************************
 * 메인 실행 함수
 **************************************************/
function startCreateDocs() {
  const ui = SpreadsheetApp.getUi();
  const start = Number(ui.prompt('시작 번호를 입력하세요', '예: 1', ui.ButtonSet.OK_CANCEL).getResponseText());
  const end = Number(ui.prompt('끝 번호를 입력하세요', '예: 120', ui.ButtonSet.OK_CANCEL).getResponseText());
  if (!start || !end || start > end) {
    ui.alert('⚠️ 잘못된 범위입니다.');
    return;
  }

  const folder = getTargetFolder();
  if (!folder) {
    ui.alert('⚠️ 지정된 폴더를 찾을 수 없습니다.');
    return;
  }

  // 중복 파일 검사
  const existing = checkExistingFiles(folder, start, end);
  if (existing.length > 0) {
    ui.alert(`⚠️ 다음 파일명이 이미 존재합니다:\n${existing.join('\n')}\n\n작업을 중단합니다.`);
    return;
  }

  // 상태 저장
  const state = { start, end, current: start, created: 0, folderId: folder.getId(), sheetName: SpreadsheetApp.getActiveSheet().getName() };
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify(state));

  // 즉시 1차 실행
  processNextBatch_();
}

/**************************************************
 * 배치 실행 함수 (내부용)
 **************************************************/
function processNextBatch_() {
  const stateJson = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (!stateJson) return;
  const state = JSON.parse(stateJson);
  const folder = DriveApp.getFolderById(state.folderId);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(state.sheetName);

  let count = 0;
  for (let i = 0; i < BATCH_SIZE && state.current <= state.end; i++, state.current++) {
    const name = `D${String(state.current).padStart(4, '0')}`;

    // 구글문서 생성
    const doc = DocumentApp.create(name);
    const file = DriveApp.getFileById(doc.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    // 링크를 시트 N열에 기록
    const data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
    const rowIndex = data.findIndex(r => r[0] === name);
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 2, 14).setValue(`https://docs.google.com/document/d/${doc.getId()}`);
    }

    count++;
    state.created++;
  }

  // 상태 업데이트
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify(state));

  if (state.current <= state.end) {
    // 다음 배치 예약 (30초 후)
    ScriptApp.newTrigger('processNextBatch_')
      .timeBased()
      .after(30 * 1000)
      .create();
  } else {
    // 완료 처리
    SpreadsheetApp.getUi().alert(`✅ 총 ${state.created}개의 문서를 생성했습니다.`);
    PropertiesService.getScriptProperties().deleteProperty(PROP_KEY);
  }
}

/**************************************************
 * 지정된 경로의 폴더 찾기
 **************************************************/
function getTargetFolder() {
  let folder = DriveApp.getRootFolder();
  for (const name of FOLDER_PATH) {
    const folders = folder.getFoldersByName(name);
    if (folders.hasNext()) folder = folders.next();
    else return null;
  }
  return folder;
}

/**************************************************
 * 파일 이름 중복 확인
 **************************************************/
function checkExistingFiles(folder, start, end) {
  const existing = [];
  const files = folder.getFiles();
  const existingNames = [];
  while (files.hasNext()) existingNames.push(files.next().getName());
  for (let i = start; i <= end; i++) {
    const name = `D${String(i).padStart(4, '0')}`;
    if (existingNames.includes(name)) existing.push(name);
  }
  return existing;
}
