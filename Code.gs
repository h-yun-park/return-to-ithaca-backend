/**
 * 이타카로의 귀환 · 도장 기록 백엔드 (Google Apps Script)
 *
 * 아래 SHEET_ID의 구글 시트에 기록하는 스크립트예요. 시트 탭 3개를 씁니다.
 *  - 선원    : 학번 1개당 1줄. 이름, 도장 개수, 부스별 도장 시각
 *  - 부스코드 : 부스별 4자리 코드 (setup() 실행 시 자동 생성, 직접 바꿔도 됨)
 *  - 기록    : 모든 도장 시도 로그 (성공/실패)
 *  - 사진    : 출석 인증 탭에서 올린 지류 도장판 사진 (원본은 드라이브 폴더)
 *
 * 처음 한 번 setup()을 실행한 뒤 웹 앱으로 배포하세요. README.md 참고.
 */

const BOOTHS = [
  ['lotus', '로토파고스 · 망각의 연꽃'],
  ['cyclops', '키클롭스의 섬'],
  ['laestrygonians', '라이스트뤼고네스'],
  ['aiaia', '아이아이에 섬'],
  ['sirens', '세이렌의 바다'],
  ['scylla', '스퀼라와 카륍디스'],
  ['thrinacia', '트리나키아'],
  ['ogygia', '오기기아'],
  ['scheria', '스케리아'],
];
// 기록용 구글 시트 ID (시트 주소의 /d/ 와 /edit 사이)
const SHEET_ID = '1zlrrzxQMxnNVYc7AYMyChkpvEFrGdjGK1yLy78F7FeQ';
const NEED = 7;
const SAILOR_HEAD = ['학번', '이름', '도장 수', '이타카 자격', '등록 시각', '마지막 도장']
  .concat(BOOTHS.map(b => b[1]));
const PHOTO_FOLDER = '이타카로의 귀환 도장판 사진';
const PHOTO_HEAD = ['제출 시각', '학번', '이름', '웹 도장 수', '사진', '원본', '확인'];
const MAX_PHOTO_BYTES = 8 * 1024 * 1024;
const MAX_FAILS = 5;          // 같은 선원이 코드 틀리는 허용 횟수
const FAIL_WINDOW_SEC = 600;  // 위 횟수를 세는 시간(10분)

/* ---------- 최초 설정: 편집기에서 한 번 실행 ---------- */
function setup() {
  const ss = ss_();
  const sailors = ss.getSheetByName('선원') || ss.insertSheet('선원');
  sailors.getRange(1, 1, 1, SAILOR_HEAD.length).setValues([SAILOR_HEAD]).setFontWeight('bold');
  sailors.setFrozenRows(1);
  sailors.setFrozenColumns(2);
  sailors.getRange('A:A').setNumberFormat('@');

  const codes = ss.getSheetByName('부스코드') || ss.insertSheet('부스코드');
  if (codes.getLastRow() < 2) {
    codes.getRange(1, 1, 1, 3).setValues([['부스 ID', '부스', '코드']]).setFontWeight('bold');
    const used = {};
    const rows = BOOTHS.map(b => {
      let c;
      do { c = String(Math.floor(1000 + Math.random() * 9000)); } while (used[c]);
      used[c] = true;
      return [b[0], b[1], c];
    });
    codes.getRange(2, 1, rows.length, 3).setValues(rows);
    codes.getRange('C:C').setNumberFormat('@');
  }

  const log = ss.getSheetByName('기록') || ss.insertSheet('기록');
  if (log.getLastRow() < 1) {
    log.getRange(1, 1, 1, 5).setValues([['시각', '학번', '이름', '부스', '결과']]).setFontWeight('bold');
    log.setFrozenRows(1);
  }
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(s1);
  setupPhotos();
}

/* ---------- 사진 탭 + 드라이브 폴더 준비 (한 번 실행) ---------- */
function setupPhotos() {
  const ss = ss_();
  const sh = ss.getSheetByName('사진') || ss.insertSheet('사진');
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, PHOTO_HEAD.length).setValues([PHOTO_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.setColumnWidth(5, 140);
  }
  folder_();
}

/* ---------- 웹 앱 입구 ---------- */
function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = e.parameter || {}; }
  return handle_(p);
}

function handle_(p) {
  let out;
  try {
    const sid = clean_(p.sno, 12);
    if (!/^\d{8,10}$/.test(sid)) throw new Error('학번을 숫자로 정확히 적어 주세요.');
    if (p.action === 'register') out = register_(sid, clean_(p.name, 30));
    else if (p.action === 'stamp') out = stamp_(sid, clean_(p.name, 30), clean_(p.booth, 30), clean_(p.code, 10));
    else if (p.action === 'status') out = status_(sid);
    else if (p.action === 'photo') out = photo_(sid, clean_(p.name, 30), String(p.image || ''), String(p.thumb || ''));
    else throw new Error('알 수 없는 요청이에요.');
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 동작 ---------- */
function register_(sid, name) {
  if (!name) throw new Error('이름을 적어 주세요.');
  return withLock_(() => {
    const sh = sheet_('선원');
    const r = findRow_(sh, sid);
    if (r) {
      sh.getRange(r, 2).setValue(name);
      return { name: name, stamps: stampsOf_(sh, r) };
    }
    sh.appendRow([sid, name, 0, '', new Date(), ''].concat(BOOTHS.map(() => '')));
    return { name: name, stamps: [] };
  });
}

function stamp_(sid, name, booth, code) {
  const idx = BOOTHS.findIndex(b => b[0] === booth);
  if (idx < 0) throw new Error('없는 부스예요.');
  const cache = CacheService.getScriptCache();
  const fk = 'fail_' + sid;  // sid = 학번
  const fails = Number(cache.get(fk) || 0);
  if (fails >= MAX_FAILS) throw new Error('코드를 여러 번 틀렸어요. 10분 뒤에 다시 시도하거나 RA에게 문의하세요.');

  return withLock_(() => {
    const sh = sheet_('선원');
    let r = findRow_(sh, sid);
    if (!r) {
      if (!name) throw new Error('먼저 이름을 등록해 주세요.');
      sh.appendRow([sid, name, 0, '', new Date(), ''].concat(BOOTHS.map(() => '')));
      r = sh.getLastRow();
    }
    const who = sh.getRange(r, 2).getValue();
    if (String(code) !== codeOf_(booth)) {
      cache.put(fk, String(fails + 1), FAIL_WINDOW_SEC);
      log_(sid, who, booth, '코드 틀림');
      throw new Error('코드가 맞지 않아요. 부스 RA에게 다시 확인해 주세요.');
    }
    const col = SAILOR_HEAD.length - BOOTHS.length + idx + 1;
    const cell = sh.getRange(r, col);
    let already = true;
    if (!cell.getValue()) {
      already = false;
      cell.setValue(new Date());
      sh.getRange(r, 6).setValue(new Date());
    }
    const stamps = stampsOf_(sh, r);
    sh.getRange(r, 3, 1, 2).setValues([[stamps.length, stamps.length >= NEED ? 'O' : '']]);
    log_(sid, who, booth, already ? '이미 받음' : '도장');
    return { name: who, stamps: stamps, already: already };
  });
}

function status_(sid) {
  const sh = sheet_('선원');
  const r = findRow_(sh, sid);
  if (!r) return { registered: false, stamps: [] };
  return { registered: true, name: sh.getRange(r, 2).getValue(), stamps: stampsOf_(sh, r) };
}

function photo_(sid, name, image, thumb) {
  const m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/.exec(image);
  if (!m) throw new Error('사진 파일을 읽지 못했어요. 다시 골라 주세요.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > MAX_PHOTO_BYTES) throw new Error('사진이 너무 커요.');
  const sailors = sheet_('선원');
  let r = findRow_(sailors, sid);
  if (!r) {
    if (!name) throw new Error('먼저 이름과 학번을 등록해 주세요.');
    withLock_(() => sailors.appendRow([sid, name, 0, '', new Date(), ''].concat(BOOTHS.map(() => ''))));
    r = findRow_(sailors, sid);
  }
  const who = sailors.getRange(r, 2).getValue();
  const count = stampsOf_(sailors, r).length;
  const when = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd_HHmmss');
  const ext = m[1] === 'image/png' ? 'png' : m[1] === 'image/webp' ? 'webp' : 'jpg';
  const file = folder_().createFile(Utilities.newBlob(bytes, m[1], sid + '_' + who + '_' + when + '.' + ext));

  return withLock_(() => {
    const sh = sheet_('사진');
    sh.appendRow([new Date(), sid, who, count, '', file.getUrl(), false]);
    const row = sh.getLastRow();
    sh.getRange(row, 7).insertCheckboxes();
    sh.setRowHeight(row, 120);
    if (/^data:image\/(?:jpeg|png);base64,/.test(thumb) && thumb.length < 1500000) {
      try {
        const img = SpreadsheetApp.newCellImage().setSourceUrl(thumb).setAltTextTitle(sid + ' 도장판').build();
        sh.getRange(row, 5).setValue(img);
      } catch (err) {
        sh.getRange(row, 5).setValue('미리보기 실패 · 원본 링크 확인');
      }
    }
    return { saved: true, stamps: stampsOf_(sailors, r) };
  });
}

/* ---------- 도우미 ---------- */
function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  const f = DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}
function sheet_(n) {
  const sh = ss_().getSheetByName(n);
  if (!sh) throw new Error('시트 설정이 안 됐어요. setup()을 먼저 실행해 주세요.');
  return sh;
}
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function findRow_(sh, sid) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const hit = sh.getRange(2, 1, last - 1, 1).createTextFinder(sid).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}
function stampsOf_(sh, r) {
  const start = SAILOR_HEAD.length - BOOTHS.length + 1;
  const vals = sh.getRange(r, start, 1, BOOTHS.length).getValues()[0];
  return BOOTHS.filter((b, i) => vals[i] !== '' && vals[i] !== null).map(b => b[0]);
}
function codeOf_(booth) {
  const rows = sheet_('부스코드').getDataRange().getValues();
  const row = rows.find(x => x[0] === booth);
  return row ? String(row[2]).trim() : '';
}
function log_(sid, name, booth, result) {
  sheet_('기록').appendRow([new Date(), sid, name, booth, result]);
}
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, '').trim().slice(0, max)
    .replace(/^[=+\-@]/, "'$&"); // 시트 수식 주입 방지
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
