function insertImageAndLogToDocs() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const activeSh = ss.getActiveSheet();

  // === 1) 작업할 "행번호" 입력 (자유형식) ===
  const resRows = ui.prompt(
    '작업할 행번호를 입력하세요',
    '예: 1, 7, 15, 26-73',
    ui.ButtonSet.OK_CANCEL
  );
  if (resRows.getSelectedButton() !== ui.Button.OK) return;

  const rows = parseRowInput_(resRows.getResponseText());
  if (!rows.length) {
    ui.alert('⚠️ 유효한 행번호가 없습니다.');
    return;
  }

  // 기록사항
  const resNote = ui.prompt('기록사항을 입력하세요', '예: 검토 완료', ui.ButtonSet.OK_CANCEL);
  if (resNote.getSelectedButton() !== ui.Button.OK) return;
  const noteStr = (resNote.getResponseText() || '').trim();

  // === 2) 오늘 날짜로 [YY-MM-DD] 자동 생성 ===
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  const dateStr = Utilities.formatDate(new Date(), tz, 'yy-MM-dd');
  const logText = `[${dateStr}] ${noteStr}`;

  // === 3) 폴더 찾기 ===
  const root = DriveApp.getFoldersByName('20 문항 관리');
  if (!root.hasNext()) {
    ui.alert('⚠️ "20 문항 관리" 폴더를 찾을 수 없습니다.');
    return;
  }
  const folder20 = root.next();

  const insertFolderIter = folder20.getFoldersByName('INSERTIMG');
  if (!insertFolderIter.hasNext()) {
    ui.alert('⚠️ "INSERTIMG" 폴더를 찾을 수 없습니다.');
    return;
  }
  const insertFolder = insertFolderIter.next();

  // === 4) 원본 시트들 캐시 (id->row, id->docUrl, P기존값) ===
  const targetSheetNames = ['[풀세트]', '[써킷]', '[문항]']; // ✅ 대괄호 포함 주의
  const sheetInfoList = targetSheetNames
    .map(name => ss.getSheetByName(name))
    .filter(Boolean)
    .map(sh => prepareSheetCache_(sh));

  let successCount = 0;

  // === 5) 입력된 행번호만 처리 ===
  for (const r of rows) {
    // 활성시트의 A열에서 id만 가져옴 (요청사항: 각 행 id는 A열)
    const idStr = String(activeSh.getRange(r, 1).getValue() || '').trim();
    if (!idStr) continue;

    // 원본은 [풀세트]/[써킷]/[문항]에서 id로 찾음
    const src = findSourceById_(sheetInfoList, idStr);
    if (!src) continue;

    const link = String(src.docUrl || '').trim(); // 원본 시트의 N열
    if (!link) continue;

    try {
      const docId = extractDocId(link);
      if (!docId) throw new Error('문서 ID를 찾을 수 없음');

      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();

      // === 이미지 삽입 (문서 맨 위) ===
      let file = null;
      const iter1 = insertFolder.getFilesByName(idStr + '.png');
      if (iter1.hasNext()) file = iter1.next();
      else {
        const iter2 = insertFolder.getFilesByName(idStr + '.jpg');
        if (iter2.hasNext()) file = iter2.next();
      }

      if (file) {
        body.insertImage(0, file.getBlob());
        body.insertParagraph(1, ''); // 기존과 동일 위치/방식
      } else {
        Logger.log(`이미지 없음: ${idStr}`);
      }

      // === 날짜 및 기록 (문서) : 기존과 동일(새 문단으로 append) ===
      body.appendParagraph(logText);
      doc.saveAndClose();

      // === 날짜 및 기록 (시트 P열) : id 매칭된 원본 행에 누적(줄바꿈) ===
      queueAppendLogSingle_(src.cache, idStr, logText);

      successCount++;
    } catch (err) {
      Logger.log(`오류 (${idStr}): ${err}`);
    }
  }

  // === 6) P열 배치 쓰기 ===
  flushPUpdates_(sheetInfoList);

  ui.alert(`✅ 완료: ${successCount}개의 문서 + P열 누적 기록이 반영되었습니다.`);
}

/* =========================
 * 행 입력 파서
 * 예: 1, 7, 15, 26-73
 * ========================= */
function parseRowInput_(input) {
  const set = new Set();

  String(input || '')
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

  return [...set].filter(n => Number.isFinite(n) && n > 0).sort((a, b) => a - b);
}

/**
 * 시트 캐시 준비:
 * - A열 id -> row
 * - N열 docUrl 캐시 (id->url)
 * - P열 기존값 읽기
 */
function prepareSheetCache_(sh) {
  const lastRow = sh.getLastRow();
  const cache = {
    sh,
    lastRow,
    idRowMap: {},   // id -> rowNumber
    nUrlMap: {},    // id -> docUrl (N열)
    pOldMap: {},    // rowNumber -> oldText
    pWriteMap: {}   // rowNumber -> newText
  };

  if (lastRow < 2) return cache;

  const numRows = lastRow - 1;

  const aVals = sh.getRange(2, 1, numRows, 1).getValues();    // A2:A
  const nVals = sh.getRange(2, 14, numRows, 1).getValues();   // N2:N
  const pVals = sh.getRange(2, 16, numRows, 1).getValues();   // P2:P

  for (let i = 0; i < numRows; i++) {
    const row = i + 2;
    const id = String(aVals[i][0] || '').trim();
    if (!id) continue;

    cache.idRowMap[id] = row;
    cache.nUrlMap[id] = String(nVals[i][0] || '').trim();

    const p = pVals[i][0];
    cache.pOldMap[row] = (p !== '' && p != null) ? String(p) : '';
  }

  return cache;
}

/**
 * id로 원본 시트/행/문서링크 찾기 (첫 매칭 반환)
 */
function findSourceById_(sheetInfoList, idStr) {
  for (const cache of sheetInfoList) {
    const row = cache.idRowMap[idStr];
    if (!row) continue;
    return { cache, row, docUrl: cache.nUrlMap[idStr] };
  }
  return null;
}

/**
 * 특정 cache(=특정 원본 시트)에서만 P열에 누적(줄바꿈)
 */
function queueAppendLogSingle_(cache, idStr, logText) {
  const row = cache.idRowMap[idStr];
  if (!row) return;

  const base = (cache.pWriteMap[row] != null) ? cache.pWriteMap[row] : cache.pOldMap[row];
  const trimmed = String(base || '').trim();
  cache.pWriteMap[row] = trimmed ? (trimmed + '\n' + logText) : logText;
}

/**
 * pWriteMap에 쌓인 변경사항을 시트에 배치 반영
 */
function flushPUpdates_(sheetInfoList) {
  sheetInfoList.forEach(cache => {
    const rows = Object.keys(cache.pWriteMap).map(r => parseInt(r, 10)).sort((a, b) => a - b);
    if (rows.length === 0) return;

    let start = rows[0];
    let prev = rows[0];

    for (let i = 1; i <= rows.length; i++) {
      const cur = rows[i];
      const isBreak = (i === rows.length) || (cur !== prev + 1);

      if (isBreak) {
        const end = prev;
        const height = end - start + 1;

        const values = [];
        for (let r = start; r <= end; r++) {
          const v = (cache.pWriteMap[r] != null) ? cache.pWriteMap[r] : cache.pOldMap[r];
          values.push([v]);
        }
        cache.sh.getRange(start, 16, height, 1).setValues(values);

        if (i < rows.length) start = cur;
      }
      prev = cur;
    }
  });
}

/**
 * Google Docs ID 추출
 */
function extractDocId(url) {
  const s = String(url || '').trim();

  let m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (m) return m[1];

  // fallback: 문서ID처럼 보이는 토큰
  m = s.match(/[-\w]{25,}/);
  if (m) return m[0];

  return null;
}
