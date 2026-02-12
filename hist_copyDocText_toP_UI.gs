/*************************************************
 * [활성 시트 기준]
 * N열(구글문서 URL) → 문서의 "이미지 제외 텍스트"를 읽어
 * P열(이력)에 붙여넣기
 *************************************************/

const HIST = {
  COL_URL: 14,   // N
  COL_OUT: 16,   // P
};

function hist_copyDocText_toP_UI() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet(); // ✅ 활성 시트
  if (!sh) {
    ui.alert("활성 시트를 찾을 수 없음");
    return;
  }

  const a = ui.prompt("작업 시작 행 번호", ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const startRow = parseInt(a.getResponseText(), 10);

  const b = ui.prompt("작업 마지막 행 번호", ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const endRow = parseInt(b.getResponseText(), 10);

  if (!Number.isFinite(startRow) || !Number.isFinite(endRow) || startRow < 1 || endRow < startRow) {
    ui.alert("행 번호가 올바르지 않음");
    return;
  }

  hist_copyDocText_toP_(sh, startRow, endRow);
  ui.alert(`완료: ${startRow} ~ ${endRow}`);
}

function hist_copyDocText_toP_(sh, startRow, endRow) {
  for (let r = startRow; r <= endRow; r++) {
    try {
      const url = getUrlFromCell_(sh, r, HIST.COL_URL);
      if (!url) {
        sh.getRange(r, HIST.COL_OUT).setValue("SKIP: URL 없음");
        continue;
      }

      const docId = extractDocId_(url);
      if (!docId) {
        sh.getRange(r, HIST.COL_OUT).setValue("SKIP: 문서 ID 추출 실패");
        continue;
      }

      const text = readDocTextWithoutImages_(docId).trim();
      sh.getRange(r, HIST.COL_OUT).setValue(text);
    } catch (e) {
      sh.getRange(r, HIST.COL_OUT).setValue(`ERROR: ${e.message}`);
    }
  }
}

function getUrlFromCell_(sh, row, col) {
  const cell = sh.getRange(row, col);
  const rt = cell.getRichTextValue();

  if (rt) {
    const link = rt.getLinkUrl();
    if (link) return link;
    for (const run of rt.getRuns()) {
      if (run.getLinkUrl()) return run.getLinkUrl();
    }
  }

  const v = String(cell.getValue() || "").trim();
  const m = v.match(/https?:\/\/docs\.google\.com\/document\/[^\s]+/);
  return m ? m[0] : v;
}

function extractDocId_(url) {
  const m = url.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : "";
}

function readDocTextWithoutImages_(docId) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const out = [];

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    const type = el.getType();

    if (type === DocumentApp.ElementType.INLINE_IMAGE) continue;

    if (
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM
    ) {
      const t = el.asText().getText().trim();
      if (t) out.push(t);
      continue;
    }

    if (type === DocumentApp.ElementType.TABLE) {
      const table = el.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        const cells = [];
        for (let c = 0; c < row.getNumCells(); c++) {
          cells.push(row.getCell(c).getText().trim());
        }
        const line = cells.join("\t").trim();
        if (line) out.push(line);
      }
    }
  }

  return out.join("\n");
}
