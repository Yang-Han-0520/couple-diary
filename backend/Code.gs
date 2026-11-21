/**
 * 우리 기록 — 저장 서버 (Google Apps Script, 구글 시트에 붙여서 사용)
 *
 * 시트에 탭 2개가 자동으로 생긴다.
 *   기록 : 일정·기념일·지도·맛집·영화·버킷리스트 등 모든 글 기록 (종류, id, 내용(JSON), 수정시각, 수정한사람, 삭제됨)
 *   사진 : 사진 목록 (id, 작은사진 파일id, 큰사진 파일id, 찍은날, 올린사람, 올린시각, 삭제됨)
 * 사진 파일은 구글 드라이브 "우리기록_사진" 폴더에 저장한다 (공유 설정 없음 = 본인 계정만 열람).
 *
 * 비밀키(SECRET)를 아는 요청만 읽고 쓸 수 있다. 공개 저장소에는 비밀키를 넣지 않는다.
 */
const SECRET = '__SECRET__';
const TZ = 'Asia/Seoul';
const PHOTO_FOLDER = '우리기록_사진';
const MAX_CELL = 49000; // 구글 시트 한 칸 최대 5만 자
const MAX_THUMBS_PER_CALL = 30;

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let p = {};
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'bad_json' });
  }
  return handle_(p);
}

function handle_(p) {
  const t0 = Date.now();
  if (!p.k || p.k !== SECRET) return out_({ ok: false, error: 'unauthorized' });
  try {
    let r;
    switch (p.a) {
      case 'ping': r = { pong: true }; break;
      case 'load': r = load_(); break;
      case 'put': r = locked_(function () { return put_(p.c, p.id, p.data, p.by); }); break;
      case 'del': r = locked_(function () { return del_(p.c, p.id, p.by); }); break;
      case 'upPhoto': r = upPhoto_(p); break;
      case 'thumbs': r = { thumbs: thumbs_(p.ids || []) }; break;
      case 'full': r = { full: full_(String(p.id || '')) }; break;
      case 'delPhoto': r = locked_(function () { return delPhoto_(String(p.id || '')); }); break;
      default: return out_({ ok: false, error: 'unknown_action' });
    }
    return out_(Object.assign({ ok: true, ms: Date.now() - t0 }, r));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- 시트 ----------
function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // 글이 수식·날짜·숫자로 바뀌지 않게 글자 형식으로 고정
    sh.getRange(1, 1, sh.getMaxRows(), header.length).setNumberFormat('@');
  }
  return sh;
}
function shRec_() { return sheet_('기록', ['종류', 'id', '내용', '수정시각', '수정한사람', '삭제됨']); }
function shPhoto_() { return sheet_('사진', ['id', '작은사진', '큰사진', '찍은날', '올린사람', '올린시각', '삭제됨']); }

function rows_(sh, ncol) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, ncol).getValues();
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function cell_(v) { return "'" + String(v == null ? '' : v); }
function txt_(v) { v = String(v); return v.charAt(0) === "'" ? v.slice(1) : v; }
function appendRow_(sh, row) { sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]); }
function findRow_(rows, col0, val, col1, val1) {
  for (let i = 0; i < rows.length; i++) {
    if (txt_(rows[i][col0]) === val && (col1 == null || txt_(rows[i][col1]) === val1)) return i + 2;
  }
  return -1;
}

// ---------- 글 기록 ----------
function load_() {
  const recs = rows_(shRec_(), 6)
    .filter(function (r) { return txt_(r[1]) !== '' && txt_(r[5]) !== 'TRUE'; })
    .map(function (r) {
      let data = null;
      try { data = JSON.parse(txt_(r[2])); } catch (e) {}
      return { c: txt_(r[0]), id: txt_(r[1]), data: data, updated: txt_(r[3]), by: txt_(r[4]) };
    });
  const photos = rows_(shPhoto_(), 7)
    .filter(function (r) { return txt_(r[0]) !== '' && txt_(r[6]) !== 'TRUE'; })
    .map(function (r) { return { id: txt_(r[0]), taken: txt_(r[3]), by: txt_(r[4]), created: txt_(r[5]) }; });
  return { records: recs, photos: photos };
}

function put_(c, id, data, by) {
  c = String(c || ''); id = String(id || '');
  if (!c || !id) throw new Error('no_id');
  const json = JSON.stringify(data == null ? {} : data);
  if (json.length > MAX_CELL) throw new Error('too_long');
  const sh = shRec_();
  const t = now_();
  const row = [cell_(c), cell_(id), cell_(json), cell_(t), cell_(by || ''), cell_('FALSE')];
  const idx = findRow_(rows_(sh, 2), 0, c, 1, id);
  if (idx > 0) sh.getRange(idx, 1, 1, 6).setValues([row]);
  else appendRow_(sh, row);
  return { updated: t };
}

function del_(c, id, by) {
  const sh = shRec_();
  const idx = findRow_(rows_(sh, 2), 0, String(c || ''), 1, String(id || ''));
  if (idx > 0) sh.getRange(idx, 4, 1, 3).setValues([[cell_(now_()), cell_(by || ''), cell_('TRUE')]]);
  return {};
}

// ---------- 사진 ----------
function folder_() {
  const props = PropertiesService.getScriptProperties();
  const fid = props.getProperty('PHOTO_FOLDER_ID');
  if (fid) {
    try { return DriveApp.getFolderById(fid); } catch (e) {}
  }
  const f = DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}

function blob_(b64, name) {
  return Utilities.newBlob(Utilities.base64Decode(String(b64)), 'image/jpeg', name);
}

function upPhoto_(p) {
  const id = String(p.id || '');
  if (!id || !p.thumb || !p.full) throw new Error('no_photo');
  const sh = shPhoto_();
  if (findRow_(rows_(sh, 1), 0, id) > 0) return { dup: true }; // 같은 사진을 두 번 올리지 않음
  const f = folder_();
  const tf = f.createFile(blob_(p.thumb, id + '_s.jpg'));
  const ff = f.createFile(blob_(p.full, id + '_l.jpg'));
  return locked_(function () {
    if (findRow_(rows_(sh, 1), 0, id) > 0) { tf.setTrashed(true); ff.setTrashed(true); return { dup: true }; }
    const t = now_();
    appendRow_(sh, [cell_(id), cell_(tf.getId()), cell_(ff.getId()), cell_(p.taken || ''), cell_(p.by || ''), cell_(t), cell_('FALSE')]);
    return { created: t };
  });
}

function photoMap_() {
  const m = {};
  rows_(shPhoto_(), 7).forEach(function (r) {
    if (txt_(r[6]) !== 'TRUE') m[txt_(r[0])] = { s: txt_(r[1]), l: txt_(r[2]) };
  });
  return m;
}

function thumbs_(ids) {
  const m = photoMap_();
  const out = {};
  ids.slice(0, MAX_THUMBS_PER_CALL).forEach(function (id) {
    const x = m[String(id)];
    if (!x) return;
    try { out[id] = Utilities.base64Encode(DriveApp.getFileById(x.s).getBlob().getBytes()); } catch (e) {}
  });
  return out;
}

function full_(id) {
  const x = photoMap_()[id];
  if (!x) return null;
  return Utilities.base64Encode(DriveApp.getFileById(x.l).getBlob().getBytes());
}

function delPhoto_(id) {
  const sh = shPhoto_();
  const rows = rows_(sh, 3);
  const idx = findRow_(rows, 0, id);
  if (idx < 0) return {};
  sh.getRange(idx, 7).setValue(cell_('TRUE'));
  // 드라이브 파일은 휴지통으로 (30일 안에 복구 가능)
  [rows[idx - 2][1], rows[idx - 2][2]].forEach(function (fid) {
    try { DriveApp.getFileById(txt_(fid)).setTrashed(true); } catch (e) {}
  });
  return {};
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
