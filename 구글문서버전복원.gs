/**************************************************
 * 구글문서 버전 복원 + '페이지 없음' 설정 — 편집기 실행용
 *
 * 기존 복원 코드에 아래 기능이 추가된 버전:
 *   복원 성공 직후, 해당 문서를 '페이지 없음(PAGELESS)'으로 전환
 *
 * 사용법 (기존과 동일)
 * 1) CONFIG 값 수정
 * 2) previewRestoreDocs 실행 → 로그 확인
 * 3) restoreDocs 실행 → 복원 + 페이지 없음 설정
 **************************************************/

const CONFIG = {
  PREFIX: 'D',                      // 문서 ID 접두어 (D 또는 S)
  START_NUM: 1452,                  // 첫 문서 번호
  END_NUM: 1489,                    // 끝 문서 번호
  CUTOFF: '2026-07-15 09:00',       // 기준 시각: 이 시각 "이전"의 마지막 버전으로 복원
  FOLDER_NAME: '문항 구글문서',      // 문서가 들어있는 폴더명
  SET_PAGELESS: true,               // 복원 후 '페이지 없음' 설정 여부
};

/** 1단계: 미리보기 (아무것도 바꾸지 않음) */
function previewRestoreDocs() { runRestore_(true); }

/** 2단계: 실제 복원 + 페이지 없음 설정 */
function restoreDocs() { runRestore_(false); }

function runRestore_(dryRun) {
  const cutoff = parseKstDateTime_(CONFIG.CUTOFF);
  if (!cutoff) throw new Error('CONFIG.CUTOFF 형식 오류. 예: 2026-07-15 09:00');

  const folderIter = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  if (!folderIter.hasNext()) throw new Error(`"${CONFIG.FOLDER_NAME}" 폴더를 찾을 수 없습니다.`);
  const folder = folderIter.next();

  const targets = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const m = f.getName().match(new RegExp(`^${CONFIG.PREFIX}(\\d{4})$`));
    if (!m) continue;
    const n = parseInt(m[1], 10);
    if (n >= CONFIG.START_NUM && n <= CONFIG.END_NUM) {
      targets.push({ name: f.getName(), id: f.getId() });
    }
  }
  targets.sort((a, b) => a.name.localeCompare(b.name));

  if (targets.length === 0) throw new Error('범위에 해당하는 문서를 폴더에서 찾지 못했습니다.');
  Logger.log(`대상 문서 ${targets.length}개 (${targets[0].name} ~ ${targets[targets.length - 1].name})`);
  Logger.log(`기준 시각: ${CONFIG.CUTOFF} 이전 마지막 버전 / 모드: ${dryRun ? '미리보기' : '★실제 복원★'} / 페이지없음: ${CONFIG.SET_PAGELESS ? 'ON' : 'OFF'}`);
  Logger.log('──────────────────────────────');

  const token = ScriptApp.getOAuthToken();
  let ok = 0, fail = 0;

  for (const t of targets) {
    try {
      const revisions = listRevisions_(t.id, token);

      let chosen = null;
      for (const rev of revisions) {
        const rt = new Date(rev.modifiedTime);
        if (rt < cutoff && (!chosen || rt > new Date(chosen.modifiedTime))) chosen = rev;
      }

      if (!chosen) {
        Logger.log(`❌ ${t.name}: 기준 시각 이전의 버전이 없음`);
        fail++;
        continue;
      }

      const revTimeStr = Utilities.formatDate(new Date(chosen.modifiedTime), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

      if (dryRun) {
        Logger.log(`📋 ${t.name}: ${revTimeStr} 버전으로 복원 예정` + (CONFIG.SET_PAGELESS ? ' + 페이지 없음 설정 예정' : ''));
        ok++;
        continue;
      }

      const exportUrl = chosen.exportLinks &&
        chosen.exportLinks['application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
      if (!exportUrl) {
        Logger.log(`❌ ${t.name}: 해당 버전의 내보내기 링크 없음`);
        fail++;
        continue;
      }

      const docxRes = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (docxRes.getResponseCode() !== 200) {
        Logger.log(`❌ ${t.name}: 버전 다운로드 실패 (HTTP ${docxRes.getResponseCode()})`);
        fail++;
        continue;
      }

      const updateRes = UrlFetchApp.fetch(
        `https://www.googleapis.com/upload/drive/v3/files/${t.id}?uploadType=media&supportsAllDrives=true`,
        {
          method: 'patch',
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          payload: docxRes.getContent(),
          headers: { Authorization: 'Bearer ' + token },
          muteHttpExceptions: true
        }
      );
      if (updateRes.getResponseCode() !== 200) {
        Logger.log(`❌ ${t.name}: 복원 업로드 실패 (HTTP ${updateRes.getResponseCode()})`);
        fail++;
        continue;
      }

      // ✅ 추가된 부분: 복원 직후 '페이지 없음' 설정
      let pagelessMsg = '';
      if (CONFIG.SET_PAGELESS) {
        const pr = restoreSetPageless_(t.id, token);
        pagelessMsg = pr.ok ? ' + 페이지 없음 설정 완료' : ` (⚠ 페이지 없음 설정 실패: ${pr.msg})`;
      }

      Logger.log(`✅ ${t.name}: ${revTimeStr} 버전으로 복원 완료${pagelessMsg}`);
      ok++;
    } catch (err) {
      Logger.log(`❌ ${t.name}: ${err && err.message ? err.message : err}`);
      fail++;
    }
  }

  Logger.log('──────────────────────────────');
  Logger.log(`${dryRun ? '미리보기' : '복원'} 종료 — 정상: ${ok}개 / 실패: ${fail}개`);
}

/** Docs API로 문서를 '페이지 없음(PAGELESS)'으로 전환
 *  (이름을 setPageless_ → restoreSetPageless_ 로 바꿨다 — 문항구글문서 자동생성.gs 의
 *   같은 이름 함수와 전역에서 겹쳐 파일 순서에 따라 한쪽이 덮였다) */
function restoreSetPageless_(docId, token) {
  const res = UrlFetchApp.fetch(
    `https://docs.googleapis.com/v1/documents/${docId}:batchUpdate`,
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
  if (res.getResponseCode() === 200) return { ok: true };
  return { ok: false, msg: `HTTP ${res.getResponseCode()} ${res.getContentText().slice(0, 200)}` };
}

/** Drive API로 버전 목록 전체 가져오기 (페이지네이션 포함) */
function listRevisions_(fileId, token) {
  const out = [];
  let pageToken = '';
  do {
    const url = `https://www.googleapis.com/drive/v3/files/${fileId}/revisions` +
      `?fields=nextPageToken,revisions(id,modifiedTime,exportLinks)&pageSize=1000` +
      (pageToken ? `&pageToken=${pageToken}` : '');
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      throw new Error(`버전 목록 조회 실패 (HTTP ${res.getResponseCode()})`);
    }
    const data = JSON.parse(res.getContentText());
    if (data.revisions) out.push(...data.revisions);
    pageToken = data.nextPageToken || '';
  } while (pageToken);
  return out;
}

/** "2026-07-15 09:00" 형식 → 한국시간 Date */
function parseKstDateTime_(s) {
  const m = String(s || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return new Date(`${m[1]}-${m[2]}-${m[3]}T${String(m[4]).padStart(2, '0')}:${m[5]}:00+09:00`);
}