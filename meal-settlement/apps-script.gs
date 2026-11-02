/**
 * 스마트 식대정산 — 구글 시트 자동 동기화 서버 (Apps Script)
 * 만든이: 구병순
 *
 * 쓰는 방법
 * 1) script.google.com 에서 새 프로젝트를 만들고 이 코드를 Code.gs 에 통째로 붙여 넣는다.
 *    (기존 식대정산 스크립트가 있으면 그 내용을 지우고 이 코드로 바꿔도 된다)
 * 2) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → 배포
 *    (기존 배포를 고칠 때는 배포 관리 → 연필 → 버전: 새 버전 → 배포. 주소가 그대로 유지된다)
 * 3) 나온 웹 앱 URL(…/exec)을 앱의 [연동 설정]에 붙여 넣는다. 휴대폰·PC 각각 한 번씩.
 *
 * 저장 위치
 * - 시트: 스크립트 속성 SHEET_ID 가 있으면 그 시트, 없으면 처음 실행 때 "식대정산_동기화" 시트를 만들어 씀
 * - 영수증 사진: 스크립트 속성 FOLDER_ID 가 있으면 그 폴더, 없으면 드라이브의 "식대정산_영수증" 폴더
 */

const REC_COLS = ['id', 'userId', 'author', 'corporation', 'siteName', 'sheetTitle', 'date', 'storeName',
  'category', 'amount', 'approvalNo', 'memo', 'updatedAt', 'receiptIds'];

function doGet() {
  return out_({ ok: true, app: '스마트 식대정산 동기화', time: Date.now() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'uploadImage') return out_(uploadImage_(req));
    if (req.action === 'getImage') return out_(getImage_(req));
    if (req.action === 'sync') {
      lock.waitLock(20000);
      return out_(sync_(req));
    }
    return out_({ ok: false, error: '알 수 없는 요청: ' + req.action });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function sync_(req) {
  const ss = book_();
  const recSh = sheet_(ss, 'records', REC_COLS);
  const delSh = sheet_(ss, 'deleted', ['id', 'deletedAt']);
  const metaSh = sheet_(ss, 'meta', ['key', 'value', 'updatedAt']);

  // 1) 삭제 목록
  const delRows = rows_(delSh);
  const deleted = {};
  delRows.forEach(r => { deleted[String(r[0])] = Number(r[1]) || 0; });
  (req.deletedIds || []).forEach(id => { deleted[String(id)] = Date.now(); });

  // 2) 내역: 더 최신(updatedAt)인 쪽이 이김, 삭제된 id 는 빼냄
  const map = {};
  rows_(recSh).forEach(r => { const o = toObj_(r); map[String(o.id)] = o; });
  (req.records || []).forEach(o => {
    const id = String(o.id);
    if (deleted[id]) return;
    const cur = map[id];
    if (!cur || (Number(o.updatedAt) || 0) >= (Number(cur.updatedAt) || 0)) map[id] = o;
  });
  Object.keys(deleted).forEach(id => { delete map[id]; });

  // 3) 예산·일별 계획(meta): 키마다 더 최신인 쪽이 이김
  const meta = {};
  rows_(metaSh).forEach(r => { meta[String(r[0])] = { key: String(r[0]), value: String(r[1]), updatedAt: Number(r[2]) || 0 }; });
  (req.meta || []).forEach(m => {
    if (!m || !m.key) return;
    const cur = meta[m.key];
    if (!cur || (Number(m.updatedAt) || 0) > cur.updatedAt) meta[m.key] = { key: m.key, value: String(m.value), updatedAt: Number(m.updatedAt) || 0 };
  });

  // 4) 시트에 다시 쓰기
  const recs = Object.keys(map).map(k => map[k]).sort((a, b) => String(a.date).localeCompare(String(b.date)) || Number(a.id) - Number(b.id));
  write_(recSh, REC_COLS, recs.map(o => REC_COLS.map(c => o[c] === undefined || o[c] === null ? '' : String(o[c]))));
  write_(delSh, ['id', 'deletedAt'], Object.keys(deleted).map(id => [id, String(deleted[id])]));
  const metaList = Object.keys(meta).map(k => meta[k]);
  write_(metaSh, ['key', 'value', 'updatedAt'], metaList.map(m => [m.key, m.value, String(m.updatedAt)]));

  return {
    ok: true,
    records: recs,
    deletedIds: Object.keys(deleted),
    meta: metaList,
    time: Date.now()
  };
}

function uploadImage_(req) {
  const m = String(req.dataUrl || '').match(/^data:([^;]+);base64,(.+)$/);
  if (!m) throw new Error('사진 형식이 올바르지 않습니다.');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], req.name || ('receipt_' + Date.now() + '.jpg'));
  const file = folder_().createFile(blob);
  return { ok: true, fileId: file.getId() };
}

function getImage_(req) {
  const blob = DriveApp.getFileById(req.fileId).getBlob();
  return { ok: true, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

// ── 도우미 ──
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function book_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet && SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const ss = SpreadsheetApp.create('식대정산_동기화');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);
  const it = DriveApp.getFoldersByName('식대정산_영수증');
  const f = it.hasNext() ? it.next() : DriveApp.createFolder('식대정산_영수증');
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}
function sheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, sh.getLastColumn()).getDisplayValues();
}
function toObj_(r) {
  const o = {};
  REC_COLS.forEach((c, i) => { o[c] = r[i]; });
  return o;
}
function write_(sh, cols, data) {
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(cols.length, sh.getLastColumn())).clearContent();
  if (!data.length) return;
  const rng = sh.getRange(2, 1, data.length, cols.length);
  rng.setNumberFormat('@'); // 날짜·승인번호가 자동 변환되지 않게 글자로 저장
  rng.setValues(data);
}
