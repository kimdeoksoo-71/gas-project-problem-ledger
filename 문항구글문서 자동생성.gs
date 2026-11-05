/**************************************************
 * 문항 구글문서 일괄 생성기 (시트용) - v3
 * - 문서 이름: 접두어 + 4자리 번호 (D0001~D9999 · S0001~S9999)
 *   ✅ v3: 시작할 때 접두어 D / S 를 고른다 (기존에는 D 고정)
 * - 폴더: "내 드라이브/20 문항 관리/문항 구글문서"
 * - 배치 크기 고정 대신 "시간 기준"으로 최대한 생성
 *   (1회 실행당 4분 30초까지 작업 후 다음 배치 자동 예약)
 * - 이름 중복 시 경고 후 중단
 * - 생성 시 활성 시트 A열에서 같은 이름을 찾아
 *   N열(14번째 열)에 구글문서 링크 기록
 * - 타임 트리거로 자동 이어서 실행
 * - 생성된 문서를 '페이지 없음(Pageless)' 형식으로 자동 설정
 *
 * ✅ v3 하루 한도 대응
 *   개인 계정(gmail.com)은 Apps Script로 만들 수 있는 문서가 하루 250개다
 *   (Google Apps Script 할당량 표 'Documents created').
 *   - 이 스크립트가 하루(한국 시간 날짜 기준)에 만든 개수를 세어 250개에서 멈추고
 *     다음 날 00:10에 자동으로 이어서 시작한다.
 *   - 다른 스크립트가 같은 날 문서를 만들어 한도에 먼저 닿으면
 *     "하루 호출 한도" 오류를 잡아 1시간 뒤 다시 시도한다.
 *   - 진행 상태를 문서 1개마다 저장한다(v2는 10개마다라 중간에 멈추면 최대 9개가 두 번 생겼다).
 *   - 이어서 할 때 같은 이름의 문서가 이미 폴더에 있으면 새로 만들지 않고 그 문서를 쓴다.
 *
 * ⚠️ 사전 설정 (최초 1회 필수):
 *   Apps Script 편집기 왼쪽 [서비스(Services)] 옆 ➕ 클릭
 *   → "Google Docs API" 선택 → 식별자 "Docs" 그대로 → 추가
 *   (이 설정이 없으면 "Docs is not defined" 오류가 발생합니다)
 **************************************************/

const FOLDER_PATH = ['20 문항 관리', '문항 구글문서'];
const PROP_KEY = 'DOC_GENERATOR_STATE';
const CD_DAILY_KEY = 'DOC_GENERATOR_DAILY';   // { date: 'yyyy-MM-dd'(KST), count }
const TIME_LIMIT_MS = 4.5 * 60 * 1000;        // 1회 실행당 최대 작업 시간 (4분 30초)
const TRIGGER_DELAY_MS = 30 * 1000;           // 다음 배치까지 대기 시간 (30초)
const CD_DAILY_LIMIT = 250;                   // 하루 생성 한도 (개인 계정 Apps Script 할당량)
const CD_QUOTA_RETRY_MS = 60 * 60 * 1000;     // 한도 오류 시 다시 시도까지 (1시간)
const CD_PREFIXES = ['D', 'S'];               // 고를 수 있는 접두어
const CD_TZ = 'Asia/Seoul';

/**************************************************
 * 메인 실행 함수 (메뉴에서 실행)
 **************************************************/
function startCreateDocs() {
  const ui = SpreadsheetApp.getUi();

  // ✅ Docs API 고급 서비스 활성화 여부 확인 (페이지 없음 설정에 필요)
  if (typeof Docs === 'undefined') {
    ui.alert(
      '⚠️ Google Docs API 서비스가 활성화되지 않았습니다.\n\n' +
      'Apps Script 편집기 왼쪽 [서비스] 옆 ➕ 를 눌러\n' +
      '"Google Docs API"를 추가한 뒤 다시 실행하세요.\n\n' +
      '(문서의 "페이지 없음" 설정에 필요합니다)'
    );
    return;
  }

  // 이미 진행 중인 작업이 있는지 확인
  const props = PropertiesService.getScriptProperties();
  const existingState = props.getProperty(PROP_KEY);
  if (existingState) {
    const prev = JSON.parse(existingState);
    const p = prev.prefix || 'D';
    const resume = ui.alert(
      '진행 중인 작업 발견',
      `이전 작업이 남아 있습니다.\n` +
      `(시트: ${prev.sheetName}, 범위: ${cdDocName_(p, prev.start)}~${cdDocName_(p, prev.end)}, ` +
      `다음: ${cdDocName_(p, prev.current)}, 생성: ${prev.created}개)\n\n` +
      `이어서 진행할까요?\n"아니오"를 누르면 이전 작업을 삭제하고 새로 시작합니다.`,
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (resume === ui.Button.CANCEL) return;
    if (resume === ui.Button.YES) {
      processNextBatch_();
      return;
    }
    // NO → 이전 상태 삭제 후 새로 시작
    props.deleteProperty(PROP_KEY);
    deleteMyTriggers_('processNextBatch_');
  }

  // ✅ v3: 접두어 선택
  const prefixRes = ui.prompt('문서 ID 접두어', `접두어를 입력하세요: ${CD_PREFIXES.join(' 또는 ')}`, ui.ButtonSet.OK_CANCEL);
  if (prefixRes.getSelectedButton() !== ui.Button.OK) return;
  const prefix = String(prefixRes.getResponseText()).trim().toUpperCase();
  if (CD_PREFIXES.indexOf(prefix) === -1) {
    ui.alert(`⚠️ 접두어는 ${CD_PREFIXES.join(' 또는 ')} 중 하나여야 합니다. (입력: "${prefix}")`);
    return;
  }

  const startRes = ui.prompt('시작 번호를 입력하세요', `숫자만 입력 (예: 1 → ${cdDocName_(prefix, 1)})`, ui.ButtonSet.OK_CANCEL);
  if (startRes.getSelectedButton() !== ui.Button.OK) return;
  const start = cdParseNum_(startRes.getResponseText(), prefix);

  const endRes = ui.prompt('끝 번호를 입력하세요', `숫자만 입력 (예: 120 → ${cdDocName_(prefix, 120)})`, ui.ButtonSet.OK_CANCEL);
  if (endRes.getSelectedButton() !== ui.Button.OK) return;
  const end = cdParseNum_(endRes.getResponseText(), prefix);

  if (!start || !end || start > end || end > 9999) {
    ui.alert('⚠️ 잘못된 범위입니다. (1~9999, 시작 ≤ 끝)');
    return;
  }

  const folder = getTargetFolder();
  if (!folder) {
    ui.alert('⚠️ 지정된 폴더를 찾을 수 없습니다.');
    return;
  }

  // ✅ v3: 활성 시트 A열에 이 ID들이 있는지 확인 (다른 시트에서 실행하는 실수 방지)
  const sheet = SpreadsheetApp.getActiveSheet();
  const rowOf = cdBuildRowMap_(sheet);
  const missing = [];
  for (let i = start; i <= end; i++) {
    const name = cdDocName_(prefix, i);
    if (!rowOf[name]) missing.push(name);
  }
  const total = end - start + 1;
  if (missing.length === total) {
    ui.alert(
      `⚠️ 현재 시트 "${sheet.getName()}"의 A열에 ${cdDocName_(prefix, start)}~${cdDocName_(prefix, end)}가 하나도 없습니다.\n` +
      `ID가 있는 시트를 연 뒤 다시 실행하세요.`
    );
    return;
  }

  // 중복 파일 검사
  const existing = checkExistingFiles(folder, prefix, start, end);
  if (existing.length > 0) {
    ui.alert(`⚠️ 다음 파일명이 이미 존재합니다:\n${existing.slice(0, 30).join('\n')}` +
      `${existing.length > 30 ? `\n… 외 ${existing.length - 30}개` : ''}\n\n작업을 중단합니다.`);
    return;
  }

  // ✅ v3: 시작 전 요약 확인
  const days = Math.ceil(total / CD_DAILY_LIMIT);
  let summary =
    `시트: ${sheet.getName()}\n` +
    `만들 문서: ${cdDocName_(prefix, start)} ~ ${cdDocName_(prefix, end)} (${total}개)\n` +
    `하루 ${CD_DAILY_LIMIT}개 한도라 약 ${days}일에 걸쳐 자동으로 이어서 만듭니다.\n` +
    `(오늘 이미 만든 개수: ${cdLoadDaily_(props).count}개)`;
  if (missing.length > 0) {
    summary += `\n\n⚠️ 이 시트 A열에 없는 ID ${missing.length}개는 문서만 만들고 N열 링크는 쓰지 못합니다:\n` +
      `${missing.slice(0, 10).join(', ')}${missing.length > 10 ? ` 외 ${missing.length - 10}개` : ''}`;
  }
  summary += '\n\n시작할까요?';
  if (ui.alert('문항구글문서 일괄생성', summary, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // 상태 저장
  const state = {
    prefix,
    start,
    end,
    current: start,
    created: 0,
    folderId: folder.getId(),
    sheetName: sheet.getName()
  };
  props.setProperty(PROP_KEY, JSON.stringify(state));

  // 즉시 1차 실행
  processNextBatch_();
}

/**************************************************
 * 배치 실행 함수 (내부용 / 트리거 실행)
 * - 시간 제한(TIME_LIMIT_MS) 또는 하루 한도에 도달할 때까지 계속 생성
 **************************************************/
function processNextBatch_() {
  const startTime = Date.now();

  // ✅ 이 함수를 가리키는 기존 일회성 트리거 정리 (잔해 누적 방지)
  deleteMyTriggers_('processNextBatch_');

  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty(PROP_KEY);
  if (!stateJson) return;
  const state = JSON.parse(stateJson);
  const prefix = state.prefix || 'D';   // v2에서 남은 상태는 접두어가 없다 → D

  const folder = DriveApp.getFolderById(state.folderId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(state.sheetName);
  if (!sheet) {
    // 시트를 못 찾으면 상태를 지우고 종료 (무한 재시도 방지)
    props.deleteProperty(PROP_KEY);
    return;
  }

  // ✅ A열을 루프 밖에서 한 번만 읽기 (ID → 행)
  const rowOf = cdBuildRowMap_(sheet);
  const daily = cdLoadDaily_(props);

  let batchCount = 0;
  let stopReason = '';   // '' | 'time' | 'daily' | 'quota'

  while (state.current <= state.end) {
    // ✅ 시간이 다 됐으면 중단하고 다음 배치 예약
    if (Date.now() - startTime > TIME_LIMIT_MS) { stopReason = 'time'; break; }
    // ✅ 하루 한도에 닿았으면 내일로
    if (daily.count >= CD_DAILY_LIMIT) { stopReason = 'daily'; break; }

    const name = cdDocName_(prefix, state.current);

    // ✅ 이어서 할 때 안전장치: 같은 이름 문서가 이미 폴더에 있으면 그 문서를 쓴다
    let docId = null;
    const same = folder.getFilesByName(name);
    while (same.hasNext()) {
      const f = same.next();
      if (!f.isTrashed()) { docId = f.getId(); break; }
    }
    if (!docId) {
      let doc;
      try {
        doc = DocumentApp.create(name);
      } catch (e) {
        if (cdIsQuotaError_(e)) { stopReason = 'quota'; break; }
        throw e;
      }
      docId = doc.getId();
      DriveApp.getFileById(docId).moveTo(folder);

      // ✅ '페이지 없음(Pageless)' 형식으로 설정
      cdSetPageless_(docId, name);

      daily.count++;
      cdSaveDaily_(props, daily);
      state.created++;
    }

    // 링크를 시트 N열에 기록
    const row = rowOf[name];
    if (row) {
      sheet.getRange(row, 14).setValue(`https://docs.google.com/document/d/${docId}`);
    }

    state.current++;
    batchCount++;

    // ✅ v3: 문서 1개마다 상태 저장 (중간에 멈춰도 같은 문서를 두 번 만들지 않는다)
    props.setProperty(PROP_KEY, JSON.stringify(state));
  }

  // 최종 상태 저장
  props.setProperty(PROP_KEY, JSON.stringify(state));

  if (state.current > state.end) {
    // ✅ 완료 처리 (트리거 실행 시 ui.alert는 에러가 나므로 toast 사용)
    props.deleteProperty(PROP_KEY);
    ss.toast(`✅ 총 ${state.created}개의 문서를 생성했습니다.`, '완료', 10);
    return;
  }

  const left = state.end - state.current + 1;
  if (stopReason === 'daily') {
    const at = cdNextDayAt_(0, 10);   // 다음 날 00:10 (KST)
    ScriptApp.newTrigger('processNextBatch_').timeBased().at(at).create();
    ss.toast(
      `오늘 ${daily.count}개를 만들어 하루 한도(${CD_DAILY_LIMIT}개)에 닿았습니다.\n` +
      `남은 ${left}개는 ${Utilities.formatDate(at, CD_TZ, 'M/d HH:mm')}에 자동으로 이어서 만듭니다.`,
      '⏸ 하루 한도', 15
    );
  } else if (stopReason === 'quota') {
    ScriptApp.newTrigger('processNextBatch_').timeBased().after(CD_QUOTA_RETRY_MS).create();
    ss.toast(
      `Google의 하루 문서 생성 한도에 닿았습니다(다른 작업이 만든 문서 포함).\n` +
      `다음: ${cdDocName_(prefix, state.current)} · 남은 ${left}개 · 1시간 뒤 다시 시도합니다.`,
      '⏸ 하루 한도', 15
    );
  } else {
    // 시간 제한 → 다음 배치 예약
    ScriptApp.newTrigger('processNextBatch_').timeBased().after(TRIGGER_DELAY_MS).create();
    ss.toast(
      `이번 배치 ${batchCount}개 처리 (누적 생성 ${state.created}개 / 남은 개수 ${left}개)\n30초 후 자동으로 계속됩니다.`,
      '⏳ 진행 중',
      10
    );
  }
}

/** 문서 이름: 접두어 + 4자리 번호 */
function cdDocName_(prefix, n) {
  return `${prefix}${String(n).padStart(4, '0')}`;
}

/** 번호 입력: "120" 또는 "S0120"(접두어가 맞을 때만) → 120, 잘못되면 0 */
function cdParseNum_(text, prefix) {
  const t = String(text || '').trim().toUpperCase();
  const m = t.match(/^([A-Z])?(\d{1,4})$/);
  if (!m) return 0;
  if (m[1] && m[1] !== prefix) return 0;
  return parseInt(m[2], 10);
}

/** 활성 시트 A열(표시 값) → { ID: 행번호 } (처음 나온 행) */
function cdBuildRowMap_(sheet) {
  const map = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  const vals = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || '').trim();
    if (v && !map[v]) map[v] = i + 2;
  }
  return map;
}

/** 오늘(KST) 이 스크립트가 만든 문서 수 */
function cdLoadDaily_(props) {
  const today = Utilities.formatDate(new Date(), CD_TZ, 'yyyy-MM-dd');
  try {
    const d = JSON.parse(props.getProperty(CD_DAILY_KEY) || '{}');
    if (d.date === today) return { date: today, count: Number(d.count) || 0 };
  } catch (e) {}
  return { date: today, count: 0 };
}

function cdSaveDaily_(props, daily) {
  props.setProperty(CD_DAILY_KEY, JSON.stringify(daily));
}

/** "하루에 너무 많이 호출" 류의 할당량 오류인가 */
function cdIsQuotaError_(e) {
  const msg = String(e && e.message ? e.message : e);
  return /too many times|quota|하루|할당량|한도/i.test(msg);
}

/** 다음 날 hh:mm (KST) */
function cdNextDayAt_(hh, mm) {
  const tomorrow = new Date(Date.now() + 24 * 60 * 60 * 1000);
  const ymd = Utilities.formatDate(tomorrow, CD_TZ, 'yyyy-MM-dd');
  return new Date(`${ymd}T${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}:00+09:00`);
}

/**************************************************
 * ✅ 구글문서를 '페이지 없음(Pageless)' 형식으로 설정
 * - Docs API 고급 서비스 필요 ([서비스]에서 "Google Docs API" 추가)
 * - 실패해도 문서 생성 자체는 유지 (로그만 남김)
 * - v3: 이름을 setPageless_ → cdSetPageless_ 로 바꿨다.
 *   구글문서버전복원.gs 에 인자가 다른 setPageless_(docId, token)가 있어
 *   (Apps Script는 모든 파일이 전역을 공유한다) 파일 순서에 따라 한쪽이 덮였다.
 **************************************************/
function cdSetPageless_(docId, name) {
  try {
    Docs.Documents.batchUpdate(
      {
        requests: [
          {
            updateDocumentStyle: {
              documentStyle: {
                documentFormat: { documentMode: 'PAGELESS' }
              },
              fields: 'documentFormat'
            }
          }
        ]
      },
      docId
    );
  } catch (e) {
    Logger.log(`⚠️ 페이지 없음 설정 실패 (${name}): ${e.message}`);
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
 * 파일 이름 중복 확인 (접두어별)
 **************************************************/
function checkExistingFiles(folder, prefix, start, end) {
  const existing = [];
  const files = folder.getFiles();
  const existingNames = new Set();
  while (files.hasNext()) existingNames.add(files.next().getName());
  for (let i = start; i <= end; i++) {
    const name = cdDocName_(prefix, i);
    if (existingNames.has(name)) existing.push(name);
  }
  return existing;
}
