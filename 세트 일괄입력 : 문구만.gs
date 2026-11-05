function insertTextIntoGoogleDocs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const rowInput = Browser.inputBox('행 번호를 입력하세요 (예: 1, 7, 15, 26-73)');
  const rows = parseRowInput_(rowInput);

  if (!rows.length) {
    Browser.msgBox('유효한 행 번호가 없습니다.');
    return;
  }

  const textToInsert = Browser.inputBox('삽입할 문구를 입력하세요');

  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  const dateStr = Utilities.formatDate(new Date(), tz, 'yy-MM-dd');
  const entry = `[${dateStr}] ${String(textToInsert || '')}`;

  const targetSheetNames = ['[풀세트]', '[써킷]', '[문항]'];

  const cache = buildIdCache_(ss, targetSheetNames);

  const notFound = [];

  for (const i of rows) {
    const idRaw = sheet.getRange(i, 1).getValue();
    const id = normalize_(idRaw);
    const docUrl = String(sheet.getRange(i, 14).getValue() || '').trim();

    // 1) 구글문서 업데이트 - 문서 맨 끝에 새 단락으로 삽입
    const docId = docUrl ? extractDocId(docUrl) : null;
    if (docId) {
      try {
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();
        body.appendParagraph(entry);
        doc.saveAndClose();
      } catch (e) {}
    }

    // 2) [풀세트]/[써킷]/[문항] 원본 시트 P열 기록
    if (id) {
      const matched = appendToOriginalP_(ss, cache, id, entry);
      if (!matched) notFound.push(`행${i} id="${idRaw}"`);
    }
  }

  if (notFound.length) {
    Browser.msgBox('완료!\n\n⚠ 원본 미발견:\n' + notFound.join('\n'));
  } else {
    Browser.msgBox('완료!');
  }
}

/* =========================
 * id 정규화: 숫자/문자열 불일치 방지
 * 123 → "123", "123.0" → "123", " abc " → "abc"
 * ========================= */
function normalize_(v) {
  let s = String(v ?? '').trim();
  s = s.replace(/\.0+$/, '');
  return s;
}

/* =========================
 * 원본 시트 A열 캐싱
 * { normalizedId: { sheet, row } }
 * ========================= */
function buildIdCache_(ss, sheetNames) {
  const map = {};
  for (const name of sheetNames) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    const lastRow = sh.getLastRow();
    if (lastRow < 1) continue;

    const ids = sh.getRange(1, 1, lastRow, 1).getValues();
    for (let r = 0; r < ids.length; r++) {
      const nid = normalize_(ids[r][0]);
      if (nid && !map[nid]) {
        map[nid] = { sheetName: name, row: r + 1 };
      }
    }
  }
  return map;
}

/* =========================
 * 캐시에서 id 매칭 → P열 append
 * ========================= */
function appendToOriginalP_(ss, cache, id, entry) {
  const hit = cache[id];
  if (!hit) return false;

  const sh = ss.getSheetByName(hit.sheetName);
  const cell = sh.getRange(hit.row, 16); // P열
  const prev = String(cell.getValue() || '').replace(/\s+$/g, '');
  cell.setValue(prev ? prev + '\n' + entry : entry);
  return true;
}

/* =========================
 * parseRowInput_ · extractDocId 는 '세트 일괄입력 : 그림과 내용.gs' 의 것을 함께 쓴다.
 * (두 파일에 같은 이름으로 조금씩 다른 정의가 있어, 파일 순서상 뒤의 것 하나가
 *  두 메뉴 모두에 쓰이고 있었다. 그쪽 정의가 이쪽 기능을 모두 포함한다:
 *  빈 입력 처리 · 25자 이상 ID 추출 폴백)
 * ========================= */
