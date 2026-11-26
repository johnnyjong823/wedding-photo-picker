/**
 * 婚紗照選擇器：照片清單 + 兩人挑選紀錄同步（Google Sheet 當資料庫）
 *
 * 使用方式：新的 Google 試算表 →「擴充功能 → Apps Script」，把整份貼上，
 * 把 PHOTO_FOLDER 換成婚紗照的雲端資料夾網址，執行 setup()，再部署成網頁應用程式。
 * 詳細步驟見 ../SETUP.md
 *
 * 工作表「照片」：分組後發布的照片清單（網頁「⚙ → 照片分組」發布）
 * 工作表「紀錄」：每列一個欄位，例如
 *   v.h.<卡片id> 老公的選擇、v.w.<卡片id> 老婆的選擇、f.<卡片id> 兩人一起決定的結果、cfg.tags 自訂用途
 *   值是 JSON：{ f:[要精修的照片id], t:{ 照片id:['小卡','相本'] } }
 */

// ★ 換成婚紗照所在的 Google 雲端資料夾網址（或資料夾 id）。子資料夾也會一起掃。
const PHOTO_FOLDER = '';

const SHEET = '紀錄';
const HEADERS = ['欄位', '值', '版本', '更新時間'];
const PHOTO_SHEET = '照片';
const PHOTO_HEADERS = ['照片id', '檔名', '拍攝時間', '組', '資料夾'];
const KEY_RE = /^((v\.[hw]|f)\.[\w-]{6,80}|cfg\.tags)$/;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📷 婚紗照選擇器')
    .addItem('① 初始化（setup）', 'setup')
    .addItem('🔑 顯示同步密碼', 'showKey')
    .addItem('✏️ 自訂同步密碼', 'setMyKey')
    .addItem('🔄 換成隨機密碼', 'resetKey')
    .addToUi();
}

/** 建立工作表與同步密碼（重跑也不會壞） */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [[SHEET, HEADERS], [PHOTO_SHEET, PHOTO_HEADERS]].forEach(([name, head]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#f6e4e7');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
  });
  const blank = ss.getSheets().find(s => s.getName() !== SHEET && s.getName() !== PHOTO_SHEET && s.getLastRow() === 0);
  if (blank) ss.deleteSheet(blank);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SYNC_KEY')) props.setProperty('SYNC_KEY', newKey_());
  if (!props.getProperty('REV')) props.setProperty('REV', '0');
  if (!props.getProperty('MREV')) props.setProperty('MREV', '0');

  // 順便確認資料夾讀得到（第一次會要求雲端硬碟權限）
  let folderMsg = '⚠ 還沒填 PHOTO_FOLDER';
  const fid = folderId_();
  if (fid) {
    const folder = DriveApp.getFolderById(fid);
    folderMsg = '照片資料夾：' + folder.getName();
    const acc = folder.getSharingAccess();
    if (acc !== DriveApp.Access.ANYONE_WITH_LINK && acc !== DriveApp.Access.ANYONE) {
      folderMsg += '\n⚠ 資料夾還沒設成「知道連結的任何人可檢視」，手機會看不到照片';
    }
  }

  const key = props.getProperty('SYNC_KEY');
  Logger.log('同步密碼：' + key);
  Logger.log(folderMsg);
  ss.toast('同步密碼：' + key + '\n' + folderMsg, '✅ 初始化完成', 30);
}

function showKey() {
  const key = PropertiesService.getScriptProperties().getProperty('SYNC_KEY');
  const msg = key ? `同步密碼：${key}\n\n打開選擇器輸入這組密碼，或用「⚙ → 複製連結」把網址傳給另一半。`
    : '請先執行「① 初始化」';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function resetKey() {
  PropertiesService.getScriptProperties().setProperty('SYNC_KEY', newKey_());
  showKey();
}

/** 自己設同步密碼（要從試算表上方選單執行，編輯器裡跑不出輸入框） */
function setMyKey() {
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {
    Logger.log('請回到試算表，從上方選單「📷 婚紗照選擇器 → ✏️ 自訂同步密碼」執行。');
    return;
  }
  const r = ui.prompt('✏️ 自訂同步密碼',
    '至少 6 個字，不能有空白（中文、英文、數字都可以）。\n改了之後兩支手機都要重新輸入，或重新傳一次連結。',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  if (key.length < 6 || /\s/.test(key)) {
    ui.alert('沒有更改：密碼要至少 6 個字，而且不能有空白。');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('SYNC_KEY', key);
  showKey();
}

/**
 * ?key=…&since=版本        → 之後的挑選紀錄變更（另附 mrev：照片清單版本）
 * ?key=…&manifest=1        → 發布過的照片清單
 * ?key=…&scan=1            → 掃描雲端資料夾（給「照片分組」用，含小縮圖）
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.key) return json({ ok: true, msg: 'photo picker endpoint is running', folder: !!folderId_() });
  const P = props_(); // 一次讀完，Properties 每日讀取有配額
  if (!P.SYNC_KEY || p.key !== P.SYNC_KEY) return json({ ok: false, error: 'bad key' });

  try {
    if (!sheet_(SHEET)) throw new Error('請先在 Apps Script 執行 setup()（① 初始化）');
    if (p.scan) return json({ ok: true, photos: scan_() });
    const rev = parseInt(P.REV, 10) || 0;
    const mrev = parseInt(P.MREV, 10) || 0;
    if (p.manifest) return json({ ok: true, mrev, photos: readManifest_() });

    let since = parseInt(p.since, 10) || 0;
    let reset = false;
    if (since > rev) { since = 0; reset = true; } // 工作表被清掉重來時，讓網頁整份重抓
    if (since === rev) return json({ ok: true, rev, mrev, changes: [] });
    // 先讀 REV 再讀資料列；寫入端是「寫列 → flush → 設 REV」，所以這裡讀到的列一定涵蓋 rev
    const changes = readRows_()
      .filter(r => Number(r[2]) > since && Number(r[2]) <= rev)
      .map(r => ({ k: String(r[0]), v: parse_(r[1]) }));
    return json({ ok: true, rev, mrev, changes, reset });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/**
 * { key, changes: { 'v.h.xxx': {...}, 'f.xxx': null(=清除) } }
 * { key, manifest: [ {id, name, t, g, fd}, … ] }   → 取代整份照片清單
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!sheet_(SHEET)) throw new Error('請先在 Apps Script 執行 setup()（① 初始化）');
    const d = JSON.parse(e.postData.contents);
    if (!keyOk_(d.key)) return json({ ok: false, error: 'bad key' });
    if (Array.isArray(d.manifest)) return json(writeManifest_(d.manifest));

    const changes = d.changes || {};
    const keys = Object.keys(changes).filter(k => KEY_RE.test(k)).slice(0, 3000);
    if (!keys.length) return json({ ok: true, rev: getProp_('REV'), written: [] });

    const rev = getProp_('REV') + 1;
    const now = new Date();
    const rows = readRows_();
    const idx = {};
    rows.forEach((r, i) => { idx[r[0]] = i; });
    keys.forEach(k => {
      const row = [k, toCell_(changes[k]), rev, now];
      if (k in idx) rows[idx[k]] = row;
      else { idx[k] = rows.length; rows.push(row); }
    });
    const sh = sheet_(SHEET);
    sh.getRange(2, 1, rows.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
    SpreadsheetApp.flush(); // 先真的寫進去再公布新版本，放鎖後別人才讀得到
    PropertiesService.getScriptProperties().setProperty('REV', String(rev));
    return json({ ok: true, rev, written: keys });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 照片清單 ---------- */

function writeManifest_(list) {
  const rows = list.slice(0, 5000).map(p => [
    String(p.id || '').replace(/[^\w-]/g, ''),
    String(p.name || '').slice(0, 200),
    Number(p.t) || '',
    Number(p.g) || 0,
    String(p.fd || '').slice(0, 200),
  ]).filter(r => r[0]);
  const sh = sheet_(PHOTO_SHEET);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, PHOTO_HEADERS.length).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, PHOTO_HEADERS.length).setValues(rows);
  }
  SpreadsheetApp.flush();
  const mrev = getProp_('MREV') + 1;
  PropertiesService.getScriptProperties().setProperty('MREV', String(mrev));
  return { ok: true, mrev, count: rows.length };
}

function readManifest_() {
  const sh = sheet_(PHOTO_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, PHOTO_HEADERS.length).getValues()
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), name: String(r[1]), t: Number(r[2]) || 0, g: Number(r[3]) || 0, fd: String(r[4] || '') }));
}

/** 列出資料夾（含子資料夾）所有照片：檔名、拍攝時間、120px 小縮圖（網頁拿來算相似度） */
function scan_() {
  const root = folderId_();
  if (!root) throw new Error('Code.gs 的 PHOTO_FOLDER 還沒填雲端資料夾網址');
  DriveApp.getFolderById(root); // 確認有權限（也讓 Apps Script 要到雲端硬碟授權）
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const fields = 'nextPageToken,files(id,name,mimeType,thumbnailLink,createdTime,imageMediaMetadata(time,width,height,rotation))';

  const photos = [];
  const queue = [{ id: root, path: '' }];
  while (queue.length) {
    const fd = queue.shift();
    let pageToken = '';
    do {
      const url = 'https://www.googleapis.com/drive/v3/files'
        + '?q=' + encodeURIComponent(`'${fd.id}' in parents and trashed=false`)
        + '&pageSize=1000&orderBy=name&supportsAllDrives=true&includeItemsFromAllDrives=true'
        + '&fields=' + encodeURIComponent(fields)
        + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
      const r = JSON.parse(UrlFetchApp.fetch(url, { headers: auth }).getContentText());
      (r.files || []).forEach(f => {
        if (f.mimeType === 'application/vnd.google-apps.folder') {
          queue.push({ id: f.id, path: fd.path ? fd.path + '/' + f.name : f.name });
        } else if (/^image\//.test(f.mimeType)) {
          const m = f.imageMediaMetadata || {};
          photos.push({
            id: f.id, name: f.name, fd: fd.path,
            t: exifTime_(m.time),
            w: m.width || 0, h: m.height || 0, rot: m.rotation || 0,
            thumb: f.thumbnailLink || '',
          });
        }
      });
      pageToken = r.nextPageToken || '';
    } while (pageToken);
  }

  // 縮圖一次抓 40 張（平行），抓不到就留空，網頁會只用拍攝時間分組
  const withThumb = photos.filter(p => /=s\d+/.test(p.thumb || ''));
  for (let i = 0; i < withThumb.length; i += 40) {
    const chunk = withThumb.slice(i, i + 40);
    const reqs = chunk.map(p => ({
      url: p.thumb.replace(/=s\d+(-[\w-]+)?$/, '') + '=s120',
      headers: auth, muteHttpExceptions: true,
    }));
    UrlFetchApp.fetchAll(reqs).forEach((res, j) => {
      if (res.getResponseCode() !== 200) return;
      const type = String(res.getHeaders()['Content-Type'] || 'image/jpeg').split(';')[0];
      chunk[j].th = 'data:' + type + ';base64,' + Utilities.base64Encode(res.getContent());
    });
  }
  photos.forEach(p => { delete p.thumb; if (!p.th) p.th = ''; });
  return photos;
}

/** EXIF "2025:10:12 14:03:22" → 毫秒（當地時間，只拿來比先後與間隔） */
function exifTime_(s) {
  const m = /^(\d{4}):(\d\d):(\d\d) (\d\d):(\d\d):(\d\d)/.exec(s || '');
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]) : 0;
}

function folderId_() {
  const s = String(PHOTO_FOLDER || '').trim();
  const m = /folders\/([\w-]+)/.exec(s) || /[?&]id=([\w-]+)/.exec(s);
  return m ? m[1] : (/^[\w-]{10,}$/.test(s) ? s : '');
}

/* ---------- 共用 ---------- */

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function readRows_() {
  const sh = sheet_(SHEET);
  const last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
}

function getProp_(k) {
  return parseInt(PropertiesService.getScriptProperties().getProperty(k), 10) || 0;
}

function keyOk_(key) {
  const k = PropertiesService.getScriptProperties().getProperty('SYNC_KEY');
  return !!k && key === k;
}

function props_() {
  return PropertiesService.getScriptProperties().getProperties();
}

function newKey_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

/** 存成 JSON 文字（開頭一定是 { [ " 或 null，不會被當成公式） */
function toCell_(v) {
  const s = JSON.stringify(v == null ? null : v);
  if (s.length > 40000) throw new Error('value too long');
  return s;
}

function parse_(c) {
  try { return JSON.parse(String(c)); } catch (e) { return null; }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
