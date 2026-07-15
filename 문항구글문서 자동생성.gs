/**************************************************
 * D0001 ~ D9999 형식의 구글문서 일괄 생성기 (시트용) - 개선판
 * - 폴더: "내 드라이브/20 문항 관리/문항 구글문서"
 * - 배치 크기 고정 대신 "시간 기준"으로 최대한 생성
 *   (1회 실행당 4분 30초까지 작업 후 다음 배치 자동 예약)
 * - 이름 중복 시 경고 후 중단
 * - 생성 시 스프레드시트 A열에서 같은 이름을 찾아
 *   N열(14번째 열)에 구글문서 링크 기록
 * - 타임 트리거로 자동 이어서 실행
 * - 개선사항:
 *   1) 시간 기반 배치 → 실행마다 가능한 최대 개수 자동 처리
 *   2) 일회성 트리거 잔해 자동 정리 (트리거 20개 제한 대응)
 *   3) 트리거 실행 시 ui.alert 에러 방지 → toast로 완료 알림
 *   4) A열 데이터를 루프 밖에서 1회만 읽어 속도 개선
 *   5) 10개마다 상태 저장 → 중간에 죽어도 이어서 실행 가능
 **************************************************/

const FOLDER_PATH = ['20 문항 관리', '문항 구글문서'];
const PROP_KEY = 'DOC_GENERATOR_STATE';
const TIME_LIMIT_MS = 4.5 * 60 * 1000; // 1회 실행당 최대 작업 시간 (4분 30초)
const TRIGGER_DELAY_MS = 30 * 1000;    // 다음 배치까지 대기 시간 (30초)
const SAVE_EVERY = 10;                 // 상태 저장 주기 (문서 N개마다)

/**************************************************
 * 메인 실행 함수 (메뉴에서 실행)
 **************************************************/
function startCreateDocs() {
  const ui = SpreadsheetApp.getUi();

  // 이미 진행 중인 작업이 있는지 확인
  const existingState = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (existingState) {
    const prev = JSON.parse(existingState);
    const resume = ui.alert(
      '진행 중인 작업 발견',
      `이전 작업이 남아 있습니다.\n(범위: ${prev.start}~${prev.end}, 현재: ${prev.current}, 생성: ${prev.created}개)\n\n이어서 진행할까요?\n"아니오"를 누르면 이전 작업을 삭제하고 새로 시작합니다.`,
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (resume === ui.Button.CANCEL) return;
    if (resume === ui.Button.YES) {
      processNextBatch_();
      return;
    }
    // NO → 이전 상태 삭제 후 새로 시작
    PropertiesService.getScriptProperties().deleteProperty(PROP_KEY);
    deleteMyTriggers_('processNextBatch_');
  }

  const startRes = ui.prompt('시작 번호를 입력하세요', '예: 1', ui.ButtonSet.OK_CANCEL);
  if (startRes.getSelectedButton() !== ui.Button.OK) return;
  const start = Number(startRes.getResponseText());

  const endRes = ui.prompt('끝 번호를 입력하세요', '예: 120', ui.ButtonSet.OK_CANCEL);
  if (endRes.getSelectedButton() !== ui.Button.OK) return;
  const end = Number(endRes.getResponseText());

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
  const state = {
    start,
    end,
    current: start,
    created: 0,
    folderId: folder.getId(),
    sheetName: SpreadsheetApp.getActiveSheet().getName()
  };
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify(state));

  // 즉시 1차 실행
  processNextBatch_();
}

/**************************************************
 * 배치 실행 함수 (내부용 / 트리거 실행)
 * - 시간 제한(TIME_LIMIT_MS)에 도달할 때까지 계속 생성
 **************************************************/
function processNextBatch_() {
  const startTime = Date.now();

  // ✅ 이 함수를 가리키는 기존 일회성 트리거 정리 (잔해 누적 방지)
  deleteMyTriggers_('processNextBatch_');

  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty(PROP_KEY);
  if (!stateJson) return;
  const state = JSON.parse(stateJson);

  const folder = DriveApp.getFolderById(state.folderId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(state.sheetName);
  if (!sheet) {
    // 시트를 못 찾으면 상태를 지우고 종료 (무한 재시도 방지)
    props.deleteProperty(PROP_KEY);
    return;
  }

  // ✅ A열 데이터를 루프 밖에서 한 번만 읽기 (속도 개선)
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    : [];

  let batchCount = 0;

  while (state.current <= state.end) {
    // ✅ 시간이 다 됐으면 중단하고 다음 배치 예약
    if (Date.now() - startTime > TIME_LIMIT_MS) break;

    const name = `D${String(state.current).padStart(4, '0')}`;

    // 구글문서 생성 후 대상 폴더로 이동
    const doc = DocumentApp.create(name);
    const file = DriveApp.getFileById(doc.getId());
    file.moveTo(folder);

    // 링크를 시트 N열에 기록
    const rowIndex = data.findIndex(r => r[0] === name);
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 2, 14).setValue(`https://docs.google.com/document/d/${doc.getId()}`);
    }

    state.current++;
    state.created++;
    batchCount++;

    // ✅ 중간에 죽어도 이어갈 수 있게 주기적으로 상태 저장
    if (batchCount % SAVE_EVERY === 0) {
      props.setProperty(PROP_KEY, JSON.stringify(state));
    }
  }

  // 최종 상태 저장
  props.setProperty(PROP_KEY, JSON.stringify(state));

  if (state.current <= state.end) {
    // 남은 작업이 있으면 다음 배치 예약
    ScriptApp.newTrigger('processNextBatch_')
      .timeBased()
      .after(TRIGGER_DELAY_MS)
      .create();
    ss.toast(
      `이번 배치 ${batchCount}개 생성 (누적 ${state.created}개 / 남은 개수 ${state.end - state.current + 1}개)\n30초 후 자동으로 계속됩니다.`,
      '⏳ 진행 중',
      10
    );
  } else {
    // ✅ 완료 처리 (트리거 실행 시 ui.alert는 에러가 나므로 toast 사용)
    props.deleteProperty(PROP_KEY);
    ss.toast(`✅ 총 ${state.created}개의 문서를 생성했습니다.`, '완료', 10);
  }
}

/**************************************************
 * 이 프로젝트에서 특정 함수를 가리키는 트리거 모두 삭제
 **************************************************/
function deleteMyTriggers_(fnName) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(t);
    }
  });
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