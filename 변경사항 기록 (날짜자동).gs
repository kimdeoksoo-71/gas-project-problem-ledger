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

  const targetSheetNames = ['풀세트', '써킷', '문항'];

  for (const i of rows) {
    const id = String(sheet.getRange(i, 1).getValue() || '').trim();       // A열 id
    const docUrl = String(sheet.getRange(i, 14).getValue() || '').trim();  // N열
    if (!docUrl) continue;

    const docId = extractDocId(docUrl);
    if (!docId) continue;

    // 1) 구글문서: 마지막 문단 끝에 "Shift+Enter(줄바꿈)"로 entry 추가 (문단 유지)
    try {
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      const paras = body.getParagraphs();

      if (paras.length === 0) {
        body.appendParagraph(entry);
      } else {
        const last = paras[paras.length - 1];
        const text = last.editAsText();
        const raw = text.getText() || '';

        if (raw.trim() === '') {
          // 마지막 문단이 비어있으면 그냥 채움 (불필요한 줄바꿈 X)
          text.setText(entry);
        } else {
          // 이미 끝이 줄바꿈이면 추가 줄바꿈 안 넣고 바로 붙임
          const rtrim = raw.replace(/[ \t]+$/g, '');
          const glue = rtrim.endsWith('\n') ? '' : '\n'; // <- Shift+Enter
          text.appendText(glue + entry);
        }
      }

      doc.saveAndClose();
    } catch (e) {
      // 실패해도 조용히 패스
    }

    // 2) [풀세트]/[써킷]/[문항] 시트의 P열: 줄바꿈(행을 달리하여) 추가
    if (id) {
      try {
        appendToMatchedSheetP_(ss, targetSheetNames, id, entry);
      } catch (e) {}
    }
  }

  Browser.msgBox('완료!');
}

/* =========================
 * 행 입력 파서
 * 예: 1, 7, 15, 26-73
 * ========================= */
function parseRowInput_(input) {
  const set = new Set();

  String(input)
    .split(',')
    .map(s => s.trim())
    .forEach(part => {
      if (/^\d+$/.test(part)) {
        set.add(Number(part));
      } else if (/^\d+\s*-\s*\d+$/.test(part)) {
        let [a, b] = part.split('-').map(Number);
        if (a > b) [a, b] = [b, a];
        for (let i = a; i <= b; i++) set.add(i);
      }
    });

  return [...set].filter(n => n > 0).sort((a, b) => a - b);
}

/* =========================
 * id 매칭 → P열 append
 * ========================= */
function appendToMatchedSheetP_(ss, sheetNames, id, entry) {
  for (const name of sheetNames) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    if (lastRow < 1) continue;

    const ids = sh.getRange(1, 1, lastRow, 1).getValues();
    for (let r = 0; r < ids.length; r++) {
      if (String(ids[r][0] || '').trim() === id) {
        const cell = sh.getRange(r + 1, 16); // P열
        const prev = String(cell.getValue() || '');
        cell.setValue(prev ? prev + '\n' + entry : entry);
        return;
      }
    }
  }
}

/* =========================
 * Google Docs ID 추출
 * ========================= */
function extractDocId(url) {
  const s = String(url || '').trim();

  let m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (m) return m[1];

  return null;
}
