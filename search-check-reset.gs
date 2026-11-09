/**
 * Phase 72 (Mathory) 부록 D — 문제검색 시트: 검색 조건(1행 B·C·E·I·L·M)이 바뀌면 U열 '선택' 체크를 전부 해제한다.
 *
 * U열 체크는 행 위치에 묶여 있어(A2 FILTER 결과) 조건을 바꾸면 같은 자리에 다른 문항이 온다.
 * W1(선택 ID 목록)·X1(개수 경고)은 수식이 소유하고, 이 트리거는 체크만 지운다.
 *
 * ⚠ 단순 트리거 onEdit은 프로젝트에 하나뿐이어야 한다 — 다른 파일에 onEdit을 만들지 말 것.
 * ⚠ 사람 편집에만 돈다(다른 스크립트·API가 1행을 바꾸면 안 돈다). 지금 1행 조건은 드롭다운으로 사람만 바꾼다.
 * ⚠ `clearTUVColumns`(TUV열 지우기.gs — 문항이동 시트용 수동 메뉴)와 무관. 합치지 않는다.
 */
const SCR = {
  SHEET: '문제검색',
  COND_COLS: [2, 3, 5, 9, 12, 13],   // B C E I L M — 각각 '선택항목' 드롭다운(A2 FILTER 조건)
  COL_CHECK: 21,                      // U '선택' 체크박스
};

function onEdit(e) {
  try {
    const r = e && e.range; if (!r) return;
    const sh = r.getSheet(); if (sh.getName() !== SCR.SHEET) return;
    if (r.getRow() !== 1 || r.getNumRows() !== 1) return;                 // 1행 한 셀/구간만
    const c0 = r.getColumn(), c1 = c0 + r.getNumColumns() - 1;
    if (!SCR.COND_COLS.some((c) => c >= c0 && c <= c1)) return;           // 조건 셀이 아니면 무시(라벨 편집 등)
    const last = sh.getLastRow(); if (last < 2) return;
    const rng = sh.getRange(2, SCR.COL_CHECK, last - 1, 1);
    const vals = rng.getValues();
    if (!vals.some((v) => v[0] === true)) return;                         // 켜진 체크가 없으면 쓰기 0
    rng.setValue(false);                                                  // clearContent 금지 — 체크박스 모양이 남아야 한다
  } catch (_) { /* 단순 트리거는 조용히 */ }
}
