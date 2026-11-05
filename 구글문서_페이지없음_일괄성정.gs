/**************************************************
 * 구글문서 '페이지 없음' 일괄 설정 (단독 실행용)
 *
 * - "문항 구글문서" 폴더에서 D####(범위) 문서를 찾아
 * - Docs API로 각 문서를 '페이지 없음(PAGELESS)'으로 전환
 * - 문서 내용은 건드리지 않음 (레이아웃 설정만 변경)
 *
 * 실행 방법 2가지 (아무거나 사용):
 *
 * ① setPagelessByConfig
 *    - 편집기에서 바로 실행 (독립형 프로젝트 OK, 시트 불필요)
 *    - 아래 PAGELESS_CFG의 번호를 수정한 뒤 실행
 *    - 결과는 실행 로그에 출력
 *
 * ② setPagelessWithPrompt
 *    - 스프레드시트에 바인딩된 프로젝트 전용
 *    - 시트 화면(메뉴)에서 실행하면 입력창이 뜸
 *    - MainMenu.gs에 추가 예:
 *      menu.addItem("구글문서 '페이지 없음' 일괄설정", 'setPagelessWithPrompt');
 **************************************************/

const PAGELESS_CFG = {
  START_NUM: 1452,              // 첫 문서 번호 (4자리)
  END_NUM: 1489,                // 끝 문서 번호 (4자리)
  FOLDER_NAME: '문항 구글문서',  // 문서가 들어있는 폴더명
};

/** ① 편집기 실행용: PAGELESS_CFG의 범위 사용 */
function setPagelessByConfig() {
  const summary = runSetPageless_(PAGELESS_CFG.START_NUM, PAGELESS_CFG.END_NUM);
  Logger.log(summary);
}

/** ② 시트 메뉴 실행용: 입력창으로 범위 받기 */
function setPagelessWithPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "'페이지 없음' 일괄 설정",
    '첫 번호와 끝 번호를 숫자 4자리로 입력하세요.\n예: 1452,1489',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const parts = res.getResponseText().split(',').map(v => parseInt(v.trim(), 10));
  const s = parts[0], e = parts[1];
  if (!Number.isFinite(s) || !Number.isFinite(e) || s > e) {
    ui.alert('⚠️ 입력 형식이 잘못되었습니다. 예: 1452,1489');
    return;
  }

  const summary = runSetPageless_(s, e);
  ui.alert(summary);
}

/** 공통 실행부: 범위 내 문서를 모두 PAGELESS로 전환 */
function runSetPageless_(startNum, endNum) {
  // === 1. 폴더에서 대상 문서 수집 ===
  const folderIter = DriveApp.getFoldersByName(PAGELESS_CFG.FOLDER_NAME);
  if (!folderIter.hasNext()) {
    return `⚠️ "${PAGELESS_CFG.FOLDER_NAME}" 폴더를 찾을 수 없습니다.`;
  }
  const folder = folderIter.next();

  const targets = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const m = f.getName().match(/^D(\d{4})$/);
    if (!m) continue;
    const n = parseInt(m[1], 10);
    if (n >= startNum && n <= endNum) targets.push({ name: f.getName(), id: f.getId() });
  }
  targets.sort((a, b) => a.name.localeCompare(b.name));

  if (targets.length === 0) {
    return '⚠️ 범위에 해당하는 문서를 폴더에서 찾지 못했습니다.';
  }

  // === 2. 문서별로 PAGELESS 설정 ===
  const token = ScriptApp.getOAuthToken();
  let ok = 0;
  const failed = [];

  for (const t of targets) {
    const res = UrlFetchApp.fetch(
      `https://docs.googleapis.com/v1/documents/${t.id}:batchUpdate`,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          requests: [{
            updateDocumentStyle: {
              documentStyle: { documentFormat: { documentMode: 'PAGELESS' } },
              fields: 'documentFormat'
            }
          }]
        }),
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    if (res.getResponseCode() === 200) {
      ok++;
      Logger.log(`✅ ${t.name}: 페이지 없음 설정 완료`);
    } else {
      failed.push(t.name);
      Logger.log(`❌ ${t.name}: HTTP ${res.getResponseCode()} ${res.getContentText().slice(0, 200)}`);
    }

    Utilities.sleep(200); // API 과호출 방지
  }

  // === 3. 결과 요약 ===
  let summary = `✅ '페이지 없음' 설정 완료\n대상: ${targets.length}개 / 성공: ${ok}개 / 실패: ${failed.length}개`;
  if (failed.length) summary += `\n\n실패 문서:\n${failed.join(', ')}`;
  return summary;
}