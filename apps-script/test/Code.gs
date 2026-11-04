/**
 * 百鮮年度計劃暨會議管理工作台 V3.1.3 — Google Apps Script 資料後端（個人帳號版）
 * V3.0.0：年度計劃格式改版（策略展開 12 欄、專案各月行動計劃 dstep）
 * V2.9.3：NotebookLM 匯出改存到指定資料夾（NLM_FOLDER_ID），舊檔自動搬入
 * V2.9.0：三組彙整同步到 Google 試算表（action=nlmsheet）；AI 回覆長度上限提高
 * V2.8.0：操作記錄顯示修改前後的值（舊值 → 新值），競爭者分析分數標示是哪一家、哪個構面
 * V2.7.0：匯出到 NotebookLM（action=nlmexport）：在附件資料夾下的「NotebookLM 資料」建立／更新 Google 文件
 * V2.6.5：共事特質可由主持人鎖定（鎖定後同仁只能提名）
 * V2.6.4：上傳的 PPT 自動轉成 Google 簡報，報告時以投影片模式播放（action=toslides）
 * V2.5.2：錄音轉逐字稿依副檔名送正確格式（iPhone m4a）、顯示詳細錯誤；錄音可刪除（移到垃圾桶）
 * V2.5.1：已報到者重複報到不覆蓋第一次報到時間
 * V2.5.0：新增「log」工作表記錄所有操作（登入、密碼、人員管理、資料寫入），管理員可在網頁「操作記錄」檢視
 * 版本：V1.0 人員寫死於網頁、生日認證碼；V2.0 人員匯入、工號登入、管理月會、例行會議；V2.1 列席者簽到、例行會議簽到與追蹤事項；V2.2 部門輪值、會議錄音與 AI 會議記錄
 * 功能：表格資料同步、附件上傳（存到雲端硬碟）、競爭者基本資料自動帶入與登記截圖辨識（需 Gemini API 金鑰）
 *
 * 第一次使用或更新後：
 *   1. 上方函式選單選 setup → 執行 → 完成授權
 *   2. 部署 → 管理部署作業 → 鉛筆 → 版本選「新版本」→ 部署
 * 圖片辨識（選用）：專案設定（齒輪）→ 指令碼屬性 → 新增屬性 GEMINI_API_KEY
 * 使用者：執行 setup 後，試算表會建立「users」工作表（工號、姓名、部門、課別、職稱、角色、分組、在職、生日…）
 *   登入帳號＝工號，預設密碼＝生日月日 4 碼，第一次登入需改密碼；角色填「管理員」者可進入主持人畫面
 *   也可以在網頁「人員管理」用 Excel 匯入
 */
const SHEET_NAME = 'docs';
const FOLDER_NAME = '百鮮年度計劃工作台附件';
const GEMINI_MODEL = 'gemini-3.8-flash';   // 若 Google 更換模型名稱，改這裡（或在指令碼屬性新增 GEMINI_MODEL 覆蓋）


/* ================= 使用者與登入（工號＋密碼，記住裝置 180 天） ================= */
const USERS_SHEET = 'users';
const USER_COLS = ['工號', '姓名', '部門', '課別', '職稱', '角色', '分組', '在職', '生日', '密碼雜湊', '憑證版本', 'LINE ID'];
const LINE_CHANNEL_ID = '2010620345';   // LINE Login 頻道 ID（LIFF ID 前半段）
const DEPT_CODE = { '業務部': 'SALES', '廠務部': 'FAC', '研發技術部': 'RD', '品質部': 'QA', '管理部': 'ADM', '總經理室': 'GM' };
const DEPT_CODES = ['SALES', 'FAC', 'RD', 'QA', 'ADM'];
const STRAT_COL_DEPT = ['SALES', 'SALES', 'SALES', 'SALES', 'FAC', 'FAC', 'FAC', 'RD', 'QA', 'ADM', 'ADM', 'ADM'];   // V3.0.0：12 欄（新增行銷設計組）
const TOKEN_DAYS = 180;
const SEED_USERS = [
  ['', '廖鴻文', '業務部', 'A'], ['', '張菱玲', '業務部', 'A'], ['', '黃世婷', '廠務部', 'A'], ['', '藍培達', '廠務部', 'A'], ['', '胡雅欣', '研發技術部', 'A'], ['', '洪瑋慈', '品質部', 'A'], ['', '曾珮瑜', '管理部', 'A'],
  ['', '廖韋智', '業務部', 'B'], ['', '何翊菁', '業務部', 'B'], ['', '李德盟', '廠務部', 'B'], ['', '藍姸雅', '研發技術部', 'B'], ['', '洪褕庭', '品質部', 'B'], ['', '汪靜榆', '管理部', 'B'], ['', '羅心伶', '管理部', 'B'],
  ['', '陳岳峰', '業務部', 'C'], ['', '許詠鈞', '業務部', 'C'], ['', '賴意雱', '業務部', 'C'], ['', '王俊毅', '廠務部', 'C'], ['', '林思廷', '研發技術部', 'C'], ['', '林皇龍', '品質部', 'C'], ['', '陳育仙', '管理部', 'C'], ['', '李怡徵', '管理部', 'C'],
  ['A021', '林奕君', '總經理室', 'A'], ['', '林裕閔', '總經理室', 'C'], ['', '林奕瀧', '總經理室', 'B']
];
function users_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.getRange('A:A').setNumberFormat('@'); sh.getRange('I:I').setNumberFormat('@');
    sh.appendRow(USER_COLS);
    // 由舊版 people 工作表帶入生日認證碼
    const old = {}; const ps = ss.getSheetByName('people');
    if (ps) ps.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) old[String(r[0]).trim()] = normCode_(r[2]); });
    SEED_USERS.forEach(r => sh.appendRow([r[0], r[1], r[2], '', '', r[1] === '林奕君' ? '管理員' : '一般', r[3], '是', old[r[1]] || '', '', 1]));
  }
  return sh;
}
function usersBust_() { try { CacheService.getScriptCache().remove('users_rows'); } catch (e) {} }
function readUsers_() {
  let v = null; const cache = CacheService.getScriptCache();
  try { const c = cache.get('users_rows'); if (c) v = JSON.parse(c); } catch (e) {}
  if (!v) {
    const sh0 = users_(); v = sh0.getDataRange().getValues();
    if (v[0].length < 12) { sh0.getRange(1, 12).setValue('LINE ID'); }
    try { const s = JSON.stringify(v); if (s.length < 95000) cache.put('users_rows', s, 30); } catch (e) {}
  }
  const byId = {}, byName = {};
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; if (!r[1]) continue;
    const u = { row: i + 1, id: String(r[0] || '').trim().toUpperCase(), name: String(r[1]).trim(), dept: String(r[2] || '').trim(), unit: String(r[3] || '').trim(), title: String(r[4] || '').trim(),
      role: String(r[5] || '一般').trim(), group: String(r[6] || '').trim(), active: String(r[7] || '是').trim() !== '否', bday: normCode_(r[8]), hash: String(r[9] || ''), ver: Number(r[10] || 1), line: String(r[11] || '') };
    if (u.id) byId[u.id] = u; byName[u.name] = u;
  }
  return { byId, byName };
}
function normCode_(x) { let s = String(x == null ? '' : x).replace(/\D/g, ''); if (s && s.length < 4) s = ('0000' + s).slice(-4); return s; }
function authSecret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('AUTH_SECRET');
  if (!s) { s = Utilities.getUuid(); p.setProperty('AUTH_SECRET', s); }
  return s;
}
const b64_ = s => Utilities.base64EncodeWebSafe(s).replace(/=+$/, '');
function sig_(payload) { return b64_(Utilities.computeHmacSha256Signature(payload, authSecret_())); }
function pwHash_(id, pw) { return b64_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, authSecret_() + '|' + id + '|' + pw)); }
function makeToken_(u) { const payload = u.id + '|' + (Date.now() + TOKEN_DAYS * 864e5) + '|' + u.ver; return b64_(payload) + '.' + sig_(payload); }
function verifyToken_(tok, U) {
  if (!tok) return null;
  const parts = String(tok).split('.'); if (parts.length !== 2) return null;
  let payload; try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0] + '==='.slice((parts[0].length + 3) % 4))).getDataAsString(); } catch (e) { return null; }
  if (sig_(payload) !== parts[1]) return null;
  const [id, exp, ver] = payload.split('|'); const u = U.byId[id];
  if (!u || !u.active || Number(exp) < Date.now() || Number(ver) !== u.ver) return null;
  return u;
}
/** 回傳身分：管理員視同主持人（'__host__'），其他人回傳姓名 */
function whoOf_(u) { return u ? (u.role === '管理員' ? '__host__' : u.name) : null; }
function publicUser_(u) { return { id: u.id, name: u.name, dept: u.dept, unit: u.unit, title: u.title, role: u.role, group: u.group, active: u.active, lineBound: !!u.line }; }
function login_(req) {
  const id = String(req.id || '').trim().toUpperCase(), cache = CacheService.getScriptCache(), key = 'fail_' + id;
  const fails = Number(cache.get(key) || 0);
  if (fails >= 5) return { ok: false, error: '錯誤次數過多，請 10 分鐘後再試，或請管理員協助' };
  let U = readUsers_(), u = U.byId[id];
  if (!u && id === VIEWER_ID) { ensureViewer_(); U = readUsers_(); u = U.byId[id]; }                // V3.1.3：唯讀測試帳號自動建立
  if (!u || !u.active) return { ok: false, error: '查無此工號，請確認或聯絡管理員' };
  const pw = String(req.pw || '').trim();
  const ok = u.hash ? pwHash_(u.id, pw) === u.hash : (!!u.bday && normCode_(pw) === u.bday);
  if (!ok) { cache.put(key, String(fails + 1), 600); return { ok: false, error: (u.hash ? '密碼不正確' : (u.bday ? '密碼不正確（尚未改過密碼者，預設為生日月日 4 碼）' : '尚未設定預設密碼，請聯絡管理員')) + '（剩 ' + (4 - fails) + ' 次）' }; }
  cache.remove(key);
  return { ok: true, token: makeToken_(u), user: publicUser_(u), mustChange: !u.hash };
}
/** V3.1.3：唯讀測試帳號（角色「檢視」，只能看不能改，不需改密碼） */
const VIEWER_ID = 'TEST', VIEWER_PW = '000000';
function ensureViewer_() {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { if (readUsers_().byId[VIEWER_ID]) return; const sh = users_();
    sh.appendRow([VIEWER_ID, '測試帳號', '', '', '', '檢視', '', '是', '', pwHash_(VIEWER_ID, VIEWER_PW), 1, '']); usersBust_(); } finally { lock.releaseLock(); }
}
function changePw_(req) {
  const U = readUsers_(), u = verifyToken_(req.auth && req.auth.token, U);
  if (!u) return { ok: false, auth: true, error: '登入已失效，請重新登入' };
  if (u.role === '檢視') return { ok: false, error: '測試帳號不能修改密碼' };
  const pw = String(req.newpw || '');
  if (pw.length < 6) return { ok: false, error: '新密碼至少 6 個字元' };
  if (normCode_(pw) === u.bday && /^\d+$/.test(pw)) return { ok: false, error: '新密碼不能與生日相同' };
  users_().getRange(u.row, 10).setValue(pwHash_(u.id, pw));
  usersBust_(); return { ok: true };
}
function adminUser_(req, U) { const u = verifyToken_(req.auth && req.auth.token, U); return u && u.role === '管理員' ? u : null; }
/** 管理員：匯入／更新人員（以工號為準；名單中沒有的人不會被刪除） */
function importUsers_(req) {
  const U = readUsers_(); if (!adminUser_(req, U)) return { ok: false, error: '需要管理員權限' };
  const sh = users_(); let add = 0, upd = 0;
  (req.rows || []).forEach(r => {
    const id = String(r.id || '').trim().toUpperCase(), name = String(r.name || '').trim(); if (!name) return;
    const ex = (id && U.byId[id]) || U.byName[name];
    const vals = [id, name, r.dept || '', r.unit || '', r.title || '', r.role || '一般', r.group || '', r.active === '否' ? '否' : '是'];
    if (ex) { sh.getRange(ex.row, 1, 1, 8).setValues([vals]); if (r.bday) sh.getRange(ex.row, 9).setValue(normCode_(r.bday)); upd++; }
    else { sh.appendRow(vals.concat([normCode_(r.bday), '', 1])); add++; }
  });
  usersBust_(); return { ok: true, added: add, updated: upd };
}
function resetPw_(req) {
  const U = readUsers_(); if (!adminUser_(req, U)) return { ok: false, error: '需要管理員權限' };
  const u = U.byId[String(req.id || '').toUpperCase()]; if (!u) return { ok: false, error: '查無此工號' };
  const sh = users_(); if (req.mode !== 'revoke') sh.getRange(u.row, 10).setValue(''); sh.getRange(u.row, 11).setValue(u.ver + 1);
  usersBust_(); return { ok: true };
}
function groupOf_(name, U) { const u = U.byName[name]; return u ? u.group : ''; }
/** 誰可以寫哪一份資料 */
function can_(who, w, U, docs) {
  if (who === '__host__') return true;
  if ((U.byName[who] || {}).role === '檢視') return false;                          // V3.1.3：唯讀帳號不能寫入
  const parts = String(w.id).split('__'), prefix = parts[0], suffix = parts[1];
  const cells = Object.keys((w.patch && w.patch.cells) || {}), items = Object.keys((w.patch && w.patch.items) || {});
  const cur = (docs[w.id] || {}).items || {}, curCells = (docs[w.id] || {}).cells || {};
  const me = U.byName[who] || {}, myDept = DEPT_CODE[me.dept] || '';
  const own = k => { const v = w.patch.items[k], old = cur[k]; return v === null ? !!(old && old.by === who) : !!(v && v.by === who && (!old || old.by === who)); };
  if (prefix === 'attn') return !cells.length && items.every(k => k === who);       // 只能寫自己的報到／請假
  if (prefix === 'fb') return !cells.length && items.every(k => k === who);         // 工作坊回饋：只能寫自己的
  if (prefix === 'rta') return !cells.length && items.every(k => k === who);        // 例行會議簽到／請假：只能寫自己的
  if (prefix === 'traits' && suffix === 'X') return !cells.length && items.every(own);  // 共事特質：自己送出的
  if (prefix === 'rt' && suffix === 'X') return !cells.length && items.every(own);   // 例行會議：只有上傳者本人（管理員另可）
  const cfg = (docs['config__X'] || {}).cells || {};
  if (/^mm\d{6}$/.test(prefix) && suffix === 'X') {                                 // 月會：主持人管流程、記錄人寫記錄
    if (curCells.closed === '1') return false;                                       // 會議已關閉
    const META = ['host', 'recorder', 'cohosts', 'date', 'time', 'place', 'mins', 'order', 'flowOk', 'closed', 'hostDept', 'recDept', 'to', 'subtitle'];
    const cohosts = String(curCells.cohosts || '').split(',').filter(String);
    const myDeptName = (U.byName[who] || {}).dept || '';
    const inHostDept = !!curCells.hostDept && myDeptName === curCells.hostDept, inRecDept = !!curCells.recDept && myDeptName === curCells.recDept;
    const memberOf = (n, dn) => !n || ((U.byName[n] || {}).dept === dn);
    const isHost = who === curCells.host || cohosts.indexOf(who) >= 0 || inHostDept, isRec = who === curCells.recorder || inRecDept;
    if (!isHost && !isRec) return false;
    return cells.every(k => {
      if (k === 'host') return inHostDept && memberOf(w.patch.cells[k], curCells.hostDept);        // 輪值部門自行決定主持人
      if (k === 'recorder') return inRecDept && memberOf(w.patch.cells[k], curCells.recDept);      // 輪值部門自行決定記錄人
      if (['cohosts', 'hostDept', 'recDept'].indexOf(k) >= 0) return false;                       // 只有管理員可以改
      if (META.indexOf(k) >= 0) return isHost;                                                     // 會議設定：主持人
      return isHost || isRec;                                                                      // 會議記錄：主持人或記錄人
    });
  }
  if (/^mr\d{6}$/.test(prefix)) {                                                  // 月會報告：會議關閉後鎖定
    if ((((docs['mm' + prefix.slice(2) + '__X'] || {}).cells) || {}).closed === '1') return false;
  }
  const ANNUAL_OPEN = ['attn', 'rt', 'rta', 'cal', 'kact', 'dproj', 'dstep'];                       // 年度計劃關閉後仍可編輯的資料
  if (cfg['plan.closed'] === '1' && !/^m[mr]\d{6}$/.test(prefix) && !/^rq\d{4}q\d$/.test(prefix) && ANNUAL_OPEN.indexOf(prefix) < 0) return false;   // V3.0.9：季度檢討（rqYYYYqN__部門）年度計劃關閉後仍可填
  if (prefix === 'snap') return false;                                               // 第一天版本：只有主持人
  if (DEPT_CODES.indexOf(suffix) >= 0) return myDept === 'GM' || myDept === suffix;    // 部門資料（含月會報告）：自己部門
  if (prefix === 'cal') return !cells.length && items.every(k => { const v = w.patch.items[k], old = cur[k];
      if (myDept === 'GM') return true; if (old && old.dept !== myDept) return false; return v === null || (v && v.dept === myDept); });
  if (prefix === 'ops') {
    if (suffix === 'X') return !items.length && (myDept === 'FAC' || myDept === 'GM');
    if (['A', 'B', 'C'].indexOf(suffix) < 0) return false;
    if (myDept !== 'GM' && groupOf_(who, U) !== suffix) return false;
    return !cells.length && items.every(k => { const v = w.patch.items[k], old = cur[k], o = old || v;
      if (o && o.kind === 'sug') return v === null ? !!(old && (old.by === who || myDept === 'GM')) : v.by === who;
      return true; });
  }
  if (prefix === 'strategy' && suffix === 'X') {
    if (myDept === 'GM') return true;
    return !items.length && cells.every(k => { const m = /^r\d+c(\d+)$/.exec(k); return !!m && STRAT_COL_DEPT[+m[1]] === myDept; });
  }
  if (prefix === 'traits8' && suffix === 'X') {                                   // V2.6.5 特質鎖定：只能提名
    if (cells.indexOf('locked') >= 0) return false;                                 // 鎖定只有主持人可以設定
    if (curCells.locked === '1') return !cells.length && items.every(k => { const v = w.patch.items[k], o = cur[k] || v; return !!(o && o.kind === 'nom'); });
  }
  if (prefix === 'config' || prefix === 'draw') return false;
  if (['A', 'B', 'C'].indexOf(suffix) >= 0) return groupOf_(who, U) === suffix;
  if (suffix === 'X') return ['traits8', 'tags', 'strategy'].indexOf(prefix) >= 0;
  return false;
}



/* ================= 操作記錄（V2.5.0） ================= */
const LOG_SHEET = 'log';
function logSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss.insertSheet(LOG_SHEET); sh.appendRow(['時間', '工號', '姓名', '動作', '資料', '內容']); sh.setFrozenRows(1); }
  return sh;
}
function nowTpe_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'); }
function cut_(v, n) { const t = typeof v === 'string' ? v : JSON.stringify(v); return t == null ? '' : (t.length > n ? t.slice(0, n) + '…' : t); }
function addLog_(rows) {
  if (!rows || !rows.length) return;
  try { const sh = logSheet_(); rows.forEach(r => sh.appendRow([r[0] || nowTpe_(), r[1] || '', r[2] || '', r[3] || '', r[4] || '', cut_(r[5] || '', 1500)])); } catch (e) {}
}
/** 寫入內容摘要：欄位 → 新值；項目 → 新增／修改／刪除 */
function patchSummary_(w, old) {
  const out = [], p = w.patch || {}, oi = (old && old.items) || {}, oc = (old && old.cells) || {};
  const SLOT = { trad: '傳統盤', proj: '專案代工', future: '未來目標' }, DIM = ['損益', '總體策略', '市場/品牌', '產品/研發', '供應鏈管理', '生產', '通路/銷售', '系統/流程', '組織/人力', '核心競爭力'];
  Object.keys(p.cells || {}).forEach(k => {
    const m = /^(trad|proj|future)\.(sr|su)(\d+)$/.exec(k), lab = m ? '（' + SLOT[m[1]] + '・' + (DIM[+m[3]] || '') + '・' + (m[2] === 'sr' ? '對方' : '百鮮') + '分數）' : '';
    out.push(k + lab + '：' + (oc[k] != null && oc[k] !== '' ? cut_(oc[k], 60) + ' → ' : '') + cut_(p.cells[k], 120));
  });
  Object.keys(p.items || {}).forEach(k => {
    const v = p.items[k];
    if (v === null) { const o = oi[k] || {}; out.push('刪除 ' + k + (o.name || o.resp || o.title || o.text ? '（' + cut_(o.name || o.resp || o.title || o.text, 60) + '）' : '')); return; }
    const label = v && (v.name || v.resp || v.title || v.text || v.goal) ? '（' + cut_(v.name || v.resp || v.title || v.text || v.goal, 60) + '）' : '';
    out.push((oi[k] ? '修改 ' : '新增 ') + k + label + '：' + cut_(v, 200));
  });
  return out.join('\n');
}
function logs_(req) {
  const U = readUsers_(); if (!adminUser_(req, U)) return { ok: false, error: '需要管理員權限' };
  const sh = logSheet_(), last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [], more: false };
  const N = Math.min(last - 1, 20000), vals = sh.getRange(last - N + 1, 1, N, 6).getDisplayValues();
  const q = String(req.q || '').trim().toLowerCase(), off = Number(req.offset || 0), lim = Math.min(500, Number(req.limit || 200));
  const out = [];
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    if (q && r.join(' ').toLowerCase().indexOf(q) < 0) continue;
    out.push(r);
  }
  return { ok: true, rows: out.slice(off, off + lim).map(r => ({ time: r[0], id: r[1], name: r[2], action: r[3], doc: r[4], detail: r[5] })), more: out.length > off + lim };
}

/* ================= 列席者／例行會議簽到（不需帳號：點名字＋簽到碼） ================= */
function fnv_(str) { let h = 2166136261; for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); } return h >>> 0; }
function todayTpe_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd'); }
function codeFor_(seed, slot) { return String(fnv_(seed + '|' + slot) % 10000).padStart(4, '0'); }
/** scope：mm202701（管理月會）或 rt:<key>（例行會議） */
function meetingInfo_(scope, docs) {
  if (/^mm\d{6}$/.test(scope)) {
    const c = ((docs[scope + '__X'] || {}).cells) || {}; if (!c.date) return null;
    return { kind: 'mm', title: (Number(scope.slice(2, 6)) - 1911) + '年' + Number(scope.slice(6)) + '月份管理月會' + (c.subtitle ? '－' + c.subtitle : ''), date: c.date, time: c.time || '14:00', seed: 'linco-local|' + c.date, names: [] };
  }
  if (/^rt:/.test(scope)) {
    const v = (((docs['rt__X'] || {}).items) || {})[scope.slice(3)]; if (!v || !v.date) return null;
    return { kind: 'rt', title: v.series || v.title || v.type, date: v.date, time: v.time || '', seed: 'rt|' + scope.slice(3), names: v.participants || [] };
  }
  return null;
}
function guestMeta_(req) {
  const { docs } = readAll_(), mi = meetingInfo_(String(req.scope || ''), docs);
  if (!mi) return { ok: false, error: '找不到這場會議' };
  if (mi.date !== todayTpe_()) return { ok: false, error: '目前不是會議期間（' + mi.date + '）' };
  return { ok: true, title: mi.title, date: mi.date, time: mi.time, kind: mi.kind, names: mi.names };
}
function guestCheckin_(req) {
  const scope = String(req.scope || ''), name = String(req.name || '').trim().slice(0, 20);
  if (!name) return { ok: false, error: '請輸入姓名' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const { docs, rows } = readAll_(), mi = meetingInfo_(scope, docs);
    if (!mi) return { ok: false, error: '找不到這場會議' };
    if (mi.date !== todayTpe_()) return { ok: false, error: '目前不是會議期間' };
    const slot = Math.floor(Date.now() / 1000 / 30), code = String(req.code || '').trim();
    if (code !== codeFor_(mi.seed, slot) && code !== codeFor_(mi.seed, slot - 1)) return { ok: false, error: '簽到碼錯誤或已過期，請輸入螢幕上最新的簽到碼' };
    const id = mi.kind === 'mm' ? 'attn__' + mi.date.replace(/-/g, '') : 'rta__' + scope.slice(3);
    const key = mi.kind === 'mm' && !readUsers_().byName[name] ? '列席:' + name : name;
    const d = docs[id] || { cells: {}, items: {} }; d.items = d.items || {}; d.cells = d.cells || {};
    if (!(d.items[key] && d.items[key].checkin)) d.items[key] = Object.assign({}, d.items[key] || {}, { checkin: new Date().toISOString(), guest: true });
    d.updatedAt = new Date().toISOString(); const text = JSON.stringify(d), sh = sheet_();
    if (rows[id]) sh.getRange(rows[id], 2, 1, 2).setValues([[text, d.updatedAt]]); else sh.appendRow([id, text, d.updatedAt]);
    docsBust_();
    const exp = Date.parse(mi.date + 'T23:59:59+08:00'), payload = 'G|' + scope + '|' + exp;
    return { ok: true, token: 'G' + b64_(payload) + '.' + sig_(payload), scope: scope, name: name, title: mi.title };
  } finally { lock.releaseLock(); }
}
function verifyGuest_(tok) {
  if (!tok || tok[0] !== 'G') return null; const parts = tok.slice(1).split('.'); if (parts.length !== 2) return null;
  let payload; try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0] + '==='.slice((parts[0].length + 3) % 4))).getDataAsString(); } catch (e) { return null; }
  if (sig_(payload) !== parts[1]) return null; const [g, scope, exp] = payload.split('|');
  return Number(exp) >= Date.now() ? scope : null;
}
/** 列席者只能看到當次會議的資料 */
function guestDocs_(scope, docs) {
  const out = {}, mi = meetingInfo_(scope, docs); if (!mi) return out;
  Object.keys(docs).forEach(id => {
    if (mi.kind === 'mm') { const ym = scope.slice(2); if (id.indexOf('mr' + ym + '__') === 0 || id === 'mm' + ym + '__X' || /^(dproj|dkpi|kact)__/.test(id) || id === 'attn__' + mi.date.replace(/-/g, '') || id === 'config__X') out[id] = docs[id]; }
    else { if (id === 'rta__' + scope.slice(3)) out[id] = docs[id]; }
  });
  if (mi.kind === 'rt') { const k = scope.slice(3); out['rt__X'] = { cells: {}, items: { [k]: ((docs['rt__X'] || {}).items || {})[k] } }; }
  return out;
}

/** 第一次請手動執行：完成授權、建立 docs 工作表與附件資料夾 */
function setup() {
  sheet_();
  users_();
  logSheet_();
  authSecret_();
  folder_();
  UrlFetchApp.getRequest('https://www.google.com');
  Logger.log('設定完成：docs 工作表與附件資料夾已建立');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['id', 'json', 'updatedAt']); }
  return sh;
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function readAll_() {
  const values = sheet_().getDataRange().getValues();
  const docs = {}, rows = {};
  for (let i = 1; i < values.length; i++) {
    const id = values[i][0];
    if (!id) continue;
    try { docs[id] = JSON.parse(values[i][1] || '{}'); } catch (e) { docs[id] = {}; }
    rows[id] = i + 1;
  }
  return { docs, rows };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const tk = (e && e.parameter && e.parameter.tk) || '';
  if (tk[0] === 'G') { const sc = verifyGuest_(tk); if (!sc) return json_({ ok: false, auth: true, guest: true, error: '會議已結束，無法再檢視' }); return json_({ ok: true, guest: sc, docs: guestDocs_(sc, cachedDocs_()) }); }
  const U = readUsers_(), u = verifyToken_(tk, U);
  if (!u) return json_({ ok: false, auth: true, error: '請先登入' });
  const all = cachedDocs_(), since = Number((e.parameter && e.parameter.since) || 0), now = Date.now();
  let docs = all, partial = false;
  if (u.role === '檢視') { docs = {}; Object.keys(all).forEach(id => { if (!/^(rv\d|review__|rq\d)/.test(id)) docs[id] = all[id]; }); }   // V3.1.3：測試帳號看不到年度檢討數據
  if (u.role !== '管理員' && ['總經理室', '業務部'].indexOf(u.dept) < 0) {                    // 個人業績（業務別分析）只給管理員、總經理室、業務部
    const rs = Object.keys(docs).filter(id => /^rv\d/.test(id) && ((docs[id] || {}).cells || {}).restrict === '1').map(id => id.replace(/__X$/, ''));
    if (rs.length) { const d2 = {}; Object.keys(docs).forEach(id => { if (!rs.some(p => id === p + '__X' || id.indexOf(p + 'p') === 0)) d2[id] = docs[id]; }); docs = d2; }
  }
  if (since) { const base = docs; docs = {}; partial = true; Object.keys(base).forEach(id => { const t = Date.parse((base[id] || {}).updatedAt || '') || 0; if (t > since - 15000) docs[id] = base[id]; }); }
  return json_({ ok: true, docs: docs, partial: partial, now: now, roster: Object.values(U.byName).map(publicUser_), me: publicUser_(u) });
}

function upload_(req, me) {
  const bytes = Utilities.base64Decode(req.data);
  if (bytes.length > 10 * 1024 * 1024) return { ok: false, error: '檔案超過 10MB' };
  const blob = Utilities.newBlob(bytes, req.mime || 'application/octet-stream', req.name || 'file');
  const file = folder_().createFile(blob);
  if (me) try { file.setDescription('by:' + me.id); } catch (e) {}
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: file.getId(), url: file.getUrl(), name: file.getName() };
}

function parse_(req) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, error: '尚未設定圖片辨識金鑰（GEMINI_API_KEY），請手動填寫' };
  const prompt = '這是台灣經濟部「公司登記」或「商業登記」的查詢截圖。請擷取資料，只輸出 JSON，欄位：' +
    'ubn(統一編號), name(公司名稱), capital(資本總額，只要數字與逗號), rep(代表人姓名), address(公司所在地), ' +
    'setup(核准設立日期，照圖上原樣), business(所營事業，每項一行，含代碼)。圖上沒有的欄位給空字串。';
  const body = {
    contents: [{ parts: [{ inline_data: { mime_type: req.mime || 'image/jpeg', data: req.data } }, { text: prompt }] }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0 }
  };
  const props = PropertiesService.getScriptProperties();
  const models = [props.getProperty('GEMINI_MODEL') || GEMINI_MODEL]
    .concat((props.getProperty('GEMINI_FALLBACK') || '').split(',').map(x => x.trim()).filter(Boolean));
  let res = null;
  for (const model of models) {
    for (let attempt = 0; attempt < 3; attempt++) {
      res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + key, {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code === 200) break;
      if (code !== 503 && code !== 429 && code !== 500) break;   // 只有「忙碌／暫時錯誤」才重試
      Utilities.sleep(2000 * (attempt + 1));
    }
    if (res.getResponseCode() === 200) break;
  }
  const code = res.getResponseCode();
  if (code === 503 || code === 429) return { ok: false, error: 'AI 辨識服務目前較忙碌，請過 1–2 分鐘再試，或先手動填寫' };
  if (code !== 200) return { ok: false, error: '圖片辨識失敗（' + code + '）：' + res.getContentText().slice(0, 200) };
  const out = JSON.parse(res.getContentText());
  const text = (((out.candidates || [])[0] || {}).content || {}).parts ? out.candidates[0].content.parts.map(p => p.text || '').join('') : '';
  let fields = {};
  try { fields = JSON.parse(text.replace(/```json|```/g, '').trim()); } catch (e) { return { ok: false, error: '辨識結果無法解讀，請手動填寫' }; }
  return { ok: true, fields: fields };
}


/** 呼叫 Gemini（含忙碌重試與備用模型）；search=true 時使用 Google 搜尋 */
function gemini_(body) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) return { error: '尚未設定 AI 金鑰（GEMINI_API_KEY）' };
  const models = [props.getProperty('GEMINI_MODEL') || GEMINI_MODEL]
    .concat((props.getProperty('GEMINI_FALLBACK') || '').split(',').map(x => x.trim()).filter(Boolean));
  let res = null;
  for (const model of models) {
    for (let attempt = 0; attempt < 3; attempt++) {
      res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + key, {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true
      });
      const c = res.getResponseCode();
      if (c === 200 || (c !== 503 && c !== 429 && c !== 500)) break;
      Utilities.sleep(2000 * (attempt + 1));
    }
    if (res.getResponseCode() === 200) break;
  }
  const code = res.getResponseCode();
  if (code === 503 || code === 429) return { error: 'AI 服務目前較忙碌，請過 1–2 分鐘再試，或先手動填寫' };
  if (code !== 200) return { error: 'AI 查詢失敗（' + code + '）：' + res.getContentText().slice(0, 200) };
  const out = JSON.parse(res.getContentText());
  const parts = (((out.candidates || [])[0] || {}).content || {}).parts || [];
  const cand = (out.candidates || [])[0] || {};
  return { text: parts.map(p => p.text || '').join(''), finish: cand.finishReason || (out.promptFeedback && out.promptFeedback.blockReason) || '' };
}

function jsonFrom_(text) {
  const m = String(text || '').replace(/```json|```/g, '').match(/\{[\s\S]*\}/);
  if (!m) return null;
  try { return JSON.parse(m[0]); } catch (e) { return null; }
}

function commas_(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function rocDate_(s) { s = String(s || ''); return s.length === 7 ? s.slice(0, 3) + '年' + s.slice(3, 5) + '月' + s.slice(5, 7) + '日' : s; }

/** 經濟部商工開放資料（以統一編號查詢）；失敗時回傳 null */
function gcis_(ubn) {
  try {
    const base = 'https://data.gcis.nat.gov.tw/od/data/api/';
    const q = '?$format=json&$filter=Business_Accounting_NO%20eq%20' + ubn + '&$skip=0&$top=1';
    const r1 = UrlFetchApp.fetch(base + '5F64D864-61CB-4D0D-8AD9-492047CC1EA6' + q, { muteHttpExceptions: true });
    if (r1.getResponseCode() !== 200) return null;
    const a = JSON.parse(r1.getContentText() || '[]')[0];
    if (!a || !a.Company_Name) return null;
    const f = {
      ubn: ubn, name: a.Company_Name,
      capital: a.Capital_Stock_Amount ? commas_(a.Capital_Stock_Amount) : '',
      rep: a.Responsible_Name || '', address: a.Company_Location || '',
      setup: rocDate_(a.Company_Setup_Date)
    };
    try {
      const r2 = UrlFetchApp.fetch(base + '236EE382-4942-41A9-BD03-CA0709025E7C' + q, { muteHttpExceptions: true });
      const b = JSON.parse(r2.getContentText() || '[]')[0];
      if (b && b.Cmp_Business) f.business = b.Cmp_Business.map(x => (x.Business_Item || '') + ' ' + (x.Business_Item_Desc || '')).join('\n');
    } catch (e) {}
    return f;
  } catch (e) { return null; }
}

/** 依公司名稱／統編／網址自動查詢競爭者基本資料 */
function lookup_(req) {
  let fields = {}, sources = [];
  let ubn = String(req.ubn || '').replace(/\D/g, '');
  const ask = '請用網路搜尋查詢台灣公司「' + (req.name || '') + '」' + (ubn ? '（統一編號 ' + ubn + '）' : '') + (req.url ? '（網站 ' + req.url + '）' : '') +
    ' 的公司登記資料與官方網站（可參考經濟部商工登記公示資料、台灣公司網等公開來源）。只輸出一個 JSON 物件，不要其他文字：' +
    '{"ubn":"統一編號8碼","name":"公司登記全名","capital":"資本總額（元，數字加逗號）","rep":"代表人","address":"公司所在地",' +
    '"setup":"核准設立日期（民國年，例如 077年03月21日）","business":"所營事業（每項一行，含代碼）","url":"官方網站網址","note":"一句話說明這家公司的主要業務"}。' +
    '查不到的欄位給空字串，不要猜測。';
  const g = gemini_({ contents: [{ parts: [{ text: ask }] }], tools: [{ google_search: {} }] });
  if (g.text) { const j = jsonFrom_(g.text); if (j) { fields = j; sources.push('AI 網路搜尋'); } }
  if (!ubn && fields.ubn) ubn = String(fields.ubn).replace(/\D/g, '');
  if (/^\d{8}$/.test(ubn)) {
    const official = gcis_(ubn);
    if (official) { Object.keys(official).forEach(k => { if (official[k]) fields[k] = official[k]; }); sources.unshift('經濟部商工開放資料'); }
  }
  if (!sources.length) return { ok: false, error: g.error || '查不到這家公司的資料，請確認公司名稱或改填統一編號後再試' };
  return { ok: true, fields: fields, sources: sources };
}

/** 一般 AI 文字處理（例如彙整共事特質） */
function ai_(req) {
  const prompt = String(req.prompt || '').slice(0, 60000);
  if (!prompt) return { ok: false, error: '沒有內容' };
  const gc = { temperature: 0.3, maxOutputTokens: 8192 }; if (req.json) gc.responseMimeType = 'application/json';
  const g = gemini_({ contents: [{ parts: [{ text: prompt }] }], generationConfig: gc });
  return g.text ? { ok: true, text: g.text } : { ok: false, error: g.error || 'AI 沒有回應' };
}

/** 讀取附件資料夾內的圖片（產生 PPT 用） */
function file64_(req) {
  const f = DriveApp.getFileById(req.id);
  const folderId = folder_().getId();
  let inFolder = false; const ps = f.getParents();
  while (ps.hasNext()) { if (ps.next().getId() === folderId) { inFolder = true; break; } }
  if (!inFolder) return { ok: false, error: '只能讀取附件資料夾內的檔案' };
  if (f.getSize() > 6 * 1024 * 1024) return { ok: false, error: '圖片太大' };
  const b = f.getBlob();
  return { ok: true, mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}

/** 會議錄音轉逐字稿（Gemini；單一檔案上限約 18MB） */
/** 錄音 MIME：依副檔名判斷（iPhone 錄的是 MP4/AAC，要用 audio/m4a；Chrome 錄的是 WebM/Opus） */
function audioMime_(f) {
  const n = String(f.getName() || '').toLowerCase(), t = String(f.getBlob().getContentType() || '').split(';')[0];
  if (/\.(m4a|mp4|aac)$/.test(n) || /mp4|m4a|aac/.test(t)) return 'audio/m4a';
  if (/\.webm$/.test(n) || /webm/.test(t)) return 'audio/webm';
  if (/\.mp3$/.test(n) || /mpeg|mp3/.test(t)) return 'audio/mp3';
  if (/\.wav$/.test(n) || /wav/.test(t)) return 'audio/wav';
  if (/\.ogg$/.test(n) || /ogg/.test(t)) return 'audio/ogg';
  if (/\.flac$/.test(n) || /flac/.test(t)) return 'audio/flac';
  return t || 'audio/webm';
}
function transcribe_(req) {
  let f;
  try { f = DriveApp.getFileById(req.id); } catch (e) { return { ok: false, error: '找不到錄音檔（可能已刪除）' }; }
  const folderId = folder_().getId();
  let inFolder = false; const ps = f.getParents(); while (ps.hasNext()) { if (ps.next().getId() === folderId) { inFolder = true; break; } }
  if (!inFolder) return { ok: false, error: '只能轉錄附件資料夾內的錄音' };
  const size = f.getSize();
  if (size < 2000) return { ok: false, error: '錄音檔幾乎是空的（' + size + ' bytes），可能沒有錄到聲音，請檢查麥克風權限' };
  if (size > 15 * 1024 * 1024) return { ok: false, error: '錄音檔太大，請分段錄音' };
  const mime = audioMime_(f), data = Utilities.base64Encode(f.getBlob().getBytes());
  const prompt = '請把這段會議錄音逐字轉錄成繁體中文，盡量區分不同發言者（發言者1、發言者2…），保留數字與專有名詞，不要摘要。如果錄音中沒有可辨識的說話內容，只回覆「（無可辨識的語音）」。';
  const g = gemini_({ contents: [{ parts: [{ inline_data: { mime_type: mime, data: data } }, { text: prompt }] }], generationConfig: { temperature: 0, maxOutputTokens: 32768 } });
  if (g.text && g.text.trim()) return { ok: true, text: g.text.trim() };
  return { ok: false, error: (g.error || ('AI 沒有回傳文字' + (g.finish ? '（原因：' + g.finish + '）' : ''))) + '｜檔案：' + f.getName() + '，' + Math.round(size / 1024) + ' KB，格式 ' + mime };
}

/** V2.6.4 上傳的 PPT 轉成 Google 簡報（投影片模式播放） */
function toSlides_(req) {
  let f; try { f = DriveApp.getFileById(req.id); } catch (e) { return { ok: false, error: '找不到檔案' }; }
  const folder = folder_(), folderId = folder.getId(); let inFolder = false; const ps = f.getParents(); while (ps.hasNext()) { if (ps.next().getId() === folderId) { inFolder = true; break; } }
  if (!inFolder) return { ok: false, error: '只能轉換附件資料夾內的檔案' };
  if (f.getMimeType() === 'application/vnd.google-apps.presentation') return { ok: true, sid: f.getId() };
  const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + f.getId() + '/copy?fields=id', {
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ name: f.getName().replace(/\.(pptx?|key)$/i, '') + '（投影片）', mimeType: 'application/vnd.google-apps.presentation', parents: [folderId] }), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return { ok: false, error: '轉換失敗（' + r.getResponseCode() + '）：' + r.getContentText().slice(0, 200) };
  const sid = JSON.parse(r.getContentText()).id;
  try { DriveApp.getFileById(sid).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { ok: true, sid: sid };
}
/** V2.7.0：匯出到 NotebookLM——建立或更新 Google 文件（同一份文件 id 不變，NotebookLM 會自動同步） */
const NLM_FOLDER_ID = '1QWcpP0UZzaemOPMcpZm_ft6ZoqWxrsot';   // V2.9.3：NotebookLM 資料夾（同仁要求存取、管理員核准）
function nlmFolder_() {
  try { const f = DriveApp.getFolderById(NLM_FOLDER_ID); if (!f.isTrashed()) return f; } catch (e) {}
  const p = folder_(), it = p.getFoldersByName('NotebookLM 資料');
  return it.hasNext() ? it.next() : p.createFolder('NotebookLM 資料');
}
/** 舊版建在其他資料夾的檔案，搬到 NotebookLM 資料夾（連結不變） */
function nlmMoveIn_(id, folder) { try { const f = DriveApp.getFileById(id); const ps = f.getParents(); let inF = false; while (ps.hasNext()) if (ps.next().getId() === folder.getId()) inF = true; if (!inF) f.moveTo(folder); } catch (e) {} }
function nlmExport_(req, me) {
  if (!me || me.role !== '管理員') return { ok: false, error: '只有管理員可以匯出' };
  const key = String(req.key || '').replace(/[^\w-]/g, ''), title = String(req.title || '').slice(0, 200), html = String(req.html || '');
  if (!key || !title || !html) return { ok: false, error: '資料不完整' };
  const props = PropertiesService.getScriptProperties(), pk = 'NLM_' + key, folder = nlmFolder_();
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  let id = props.getProperty(pk), created = false;
  if (id) { try { const f = DriveApp.getFileById(id); if (f.isTrashed()) id = null; else if (f.getName() !== title) f.setName(title); } catch (e) { id = null; } }
  if (id) nlmMoveIn_(id, folder);
  if (!id) {
    const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files?fields=id', { method: 'post', contentType: 'application/json', headers: auth,
      payload: JSON.stringify({ name: title, mimeType: 'application/vnd.google-apps.document', parents: [folder.getId()] }), muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return { ok: false, error: '建立文件失敗（' + r.getResponseCode() + '）：' + r.getContentText().slice(0, 200) };
    id = JSON.parse(r.getContentText()).id; props.setProperty(pk, id); created = true;
  }
  const u = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + id + '?uploadType=media', { method: 'patch', contentType: 'text/html; charset=UTF-8', headers: auth,
    payload: Utilities.newBlob(html, 'text/html').getBytes(), muteHttpExceptions: true });
  if (u.getResponseCode() !== 200) return { ok: false, error: '寫入文件失敗（' + u.getResponseCode() + '）：' + u.getContentText().slice(0, 200) };
  return { ok: true, id: id, created: created, url: 'https://docs.google.com/document/d/' + id + '/edit', folderUrl: folder.getUrl() };
}
/** V2.9.0：三組彙整寫入 Google 試算表（每張原始表格一個分頁；同一份試算表連結不變，NotebookLM 會自動同步） */
function nlmSheet_(req, me) {
  if (!me || me.role !== '管理員') return { ok: false, error: '只有管理員可以同步' };
  const key = String(req.key || '').replace(/[^\w-]/g, ''), title = String(req.title || '').slice(0, 200), sheets = req.sheets || [];
  if (!key || !title || !sheets.length) return { ok: false, error: '資料不完整' };
  const props = PropertiesService.getScriptProperties(), pk = 'NLMS_' + key, folder = nlmFolder_();
  let id = props.getProperty(pk), created = false;
  if (id) { try { const f = DriveApp.getFileById(id); if (f.isTrashed()) id = null; else if (f.getName() !== title) f.setName(title); } catch (e) { id = null; } }
  if (id) nlmMoveIn_(id, folder);
  if (!id) {
    const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files?fields=id', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify({ name: title, mimeType: 'application/vnd.google-apps.spreadsheet', parents: [folder.getId()] }), muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return { ok: false, error: '建立試算表失敗（' + r.getResponseCode() + '）：' + r.getContentText().slice(0, 200) };
    id = JSON.parse(r.getContentText()).id; props.setProperty(pk, id); created = true;
  }
  const ss = SpreadsheetApp.openById(id), names = [];
  sheets.forEach(function (s, si) {
    const name = String(s.name || ('表' + (si + 1))).slice(0, 90); names.push(name);
    let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name, si); else ss.setActiveSheet(sh), ss.moveActiveSheet(si + 1);
    sh.clear();
    const rows = (s.rows || []).map(function (r) { return (r || []).map(function (v) { return v == null ? '' : String(v).slice(0, 49000); }); });
    const w = Math.max.apply(null, rows.map(function (r) { return r.length; }).concat([1]));
    const vals = rows.map(function (r) { while (r.length < w) r.push(''); return r; });
    if (!vals.length) return;
    const rg = sh.getRange(1, 1, vals.length, w); rg.setNumberFormat('@'); rg.setValues(vals); rg.setWrap(true).setVerticalAlignment('top');
    sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#F1EBDD'); sh.setFrozenRows(1);
    for (let c = 1; c <= w; c++) { const len = Math.max.apply(null, vals.map(function (r) { return Math.max.apply(null, String(r[c - 1]).split('\n').map(function (l) { return l.length; })); })); sh.setColumnWidth(c, Math.min(420, Math.max(70, len * 14))); }
  });
  ss.getSheets().forEach(function (sh) { if (names.indexOf(sh.getName()) < 0 && ss.getSheets().length > 1) ss.deleteSheet(sh); });
  return { ok: true, id: id, created: created, url: 'https://docs.google.com/spreadsheets/d/' + id + '/edit', folderUrl: folder.getUrl() };
}
/** 刪除錄音／附件：管理員或上傳者本人，移到雲端硬碟垃圾桶 */
function trash_(req, me) {
  let f; try { f = DriveApp.getFileById(req.id); } catch (e) { return { ok: true }; }
  const folderId = folder_().getId(); let inFolder = false; const ps = f.getParents(); while (ps.hasNext()) { if (ps.next().getId() === folderId) { inFolder = true; break; } }
  if (!inFolder) return { ok: false, error: '只能刪除附件資料夾內的檔案' };
  if (me.role !== '管理員' && String(f.getDescription() || '') !== 'by:' + me.id) return { ok: false, error: '只有上傳者或管理員可以刪除檔案（已從列表移除）' };
  f.setTrashed(true); return { ok: true };
}

/* ================= LINE 綁定（通知小卡回覆用） ================= */
function lineSub_(idToken) {
  if (!idToken) return null;
  const r = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', { method: 'post', payload: { id_token: idToken, client_id: LINE_CHANNEL_ID }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return null;
  const j = JSON.parse(r.getContentText()); return j && j.sub ? j.sub : null;
}
function lineAuth_(req) {
  const sub = lineSub_(req.idToken); if (!sub) return { ok: false, error: 'LINE 驗證失敗，請重新開啟' };
  const U = readUsers_(); const u = Object.values(U.byName).find(x => x.line === sub && x.active);
  if (!u) return { ok: true, bound: false };
  return { ok: true, bound: true, token: makeToken_(u), user: publicUser_(u) };
}
function lineBind_(req) {
  const sub = lineSub_(req.idToken); if (!sub) return { ok: false, error: 'LINE 驗證失敗，請重新開啟' };
  const r = login_(req); if (!r.ok) return r;
  const U = readUsers_(), u = U.byId[String(req.id || '').trim().toUpperCase()];
  Object.values(U.byName).forEach(x => { if (x.line === sub && x.row !== u.row) users_().getRange(x.row, 12).setValue(''); });
  users_().getRange(u.row, 12).setValue(sub);
  usersBust_(); return r;
}

/** 讀取用快取：多人同時開會時，避免每次輪詢都重讀整張試算表（寫入後立即失效） */
function cachedDocs_() {
  const cache = CacheService.getScriptCache();
  try { const n = Number(cache.get('docs_n') || 0); if (n) { const keys = []; for (let i = 0; i < n; i++) keys.push('docs_' + i); const got = cache.getAll(keys);
    if (keys.every(k => got[k] != null)) return JSON.parse(keys.map(k => got[k]).join('')); } } catch (e) {}
  const docs = readAll_().docs;
  try { const s = JSON.stringify(docs), parts = {}; let n = 0; for (let i = 0; i < s.length; i += 90000) parts['docs_' + (n++)] = s.slice(i, i + 90000);
    if (n <= 80) { parts.docs_n = String(n); cache.putAll(parts, 5); } } catch (e) {}
  return docs;
}
function docsBust_() { try { CacheService.getScriptCache().remove('docs_n'); } catch (e) {} }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '資料格式錯誤' }); }
  try {
    if (req.action === 'login') { const r = login_(req); addLog_([[null, String(req.id || '').toUpperCase(), r.user ? r.user.name : '', r.ok ? '登入' : '登入失敗', '', r.ok ? '' : r.error]]); return json_(r); }
    if (req.action === 'changepw') { const r = changePw_(req); const u = verifyToken_(req.auth && req.auth.token, readUsers_()); addLog_([[null, u ? u.id : '', u ? u.name : '', r.ok ? '修改密碼' : '修改密碼失敗', '', r.ok ? '' : r.error]]); return json_(r); }
    if (req.action === 'import') { const r = importUsers_(req); const u = verifyToken_(req.auth && req.auth.token, readUsers_()); addLog_([[null, u ? u.id : '', u ? u.name : '', r.ok ? '人員管理' : '人員管理（拒絕）', 'users', r.ok ? (req.rows || []).map(x => (x.id || '') + ' ' + (x.name || '') + ' ' + (x.dept || '') + ' ' + (x.role || '') + (x.active === '否' ? ' 離職' : '')).join('\n') : r.error]]); return json_(r); }
    if (req.action === 'resetpw') { const r = resetPw_(req); const u = verifyToken_(req.auth && req.auth.token, readUsers_()); addLog_([[null, u ? u.id : '', u ? u.name : '', r.ok ? (req.mode === 'revoke' ? '強制登出' : '重設密碼') : '重設密碼（拒絕）', 'users', String(req.id || '')]]); return json_(r); }
    if (req.action === 'gmeta') return json_(guestMeta_(req));
    if (req.action === 'gcheckin') { const r = guestCheckin_(req); addLog_([[null, '', String(req.name || ''), r.ok ? '列席簽到' : '列席簽到失敗', String(req.scope || ''), r.ok ? '' : (r.error || '')]]); return json_(r); }
    if (req.action === 'lineauth') { const r = lineAuth_(req); if (r.bound) addLog_([[null, r.user.id, r.user.name, 'LINE 登入', '', '']]); return json_(r); }
    if (req.action === 'linebind') { const r = lineBind_(req); addLog_([[null, String(req.id || '').toUpperCase(), r.user ? r.user.name : '', r.ok ? 'LINE 綁定' : 'LINE 綁定失敗', '', r.ok ? '' : r.error]]); return json_(r); }
  } catch (err) { return json_({ ok: false, error: String(err) }); }
  const U = readUsers_();
  const me_ = verifyToken_(req.auth && req.auth.token, U);
  const who = whoOf_(me_);
  if (!who) return json_({ ok: false, auth: true, error: '登入已失效，請重新登入' });
  if (me_.role === '檢視' && ['file64', 'sync', 'write', 'writes'].indexOf(req.action) < 0 && !req.writes) return json_({ ok: false, error: '測試帳號只能檢視' });
  try {
    if (req.action === 'logs') return json_(logs_(req));
    if (req.action === 'upload') { const r = upload_(req, me_); addLog_([[null, me_.id, me_.name, r.ok ? '上傳附件' : '上傳失敗', '', String(req.name || '') + (r.ok ? '' : '：' + r.error)]]); return json_(r); }
    if (req.action === 'parse') return json_(parse_(req));
    if (req.action === 'lookup') return json_(lookup_(req));
    if (req.action === 'ai') return json_(ai_(req));
    if (req.action === 'file64') return json_(file64_(req));
    if (req.action === 'transcribe') { const r = transcribe_(req); addLog_([[null, me_.id, me_.name, r.ok ? '錄音轉逐字稿' : '錄音轉逐字稿失敗', String(req.id || ''), r.ok ? String(r.text.length) + ' 字' : r.error]]); return json_(r); }
    if (req.action === 'toslides') { const r = toSlides_(req); addLog_([[null, me_.id, me_.name, r.ok ? '簡報轉投影片' : '簡報轉投影片失敗', String(req.id || ''), r.ok ? '' : r.error]]); return json_(r); }
    if (req.action === 'nlmsheet') { const r = nlmSheet_(req, me_); addLog_([[null, me_.id, me_.name, r.ok ? '同步試算表到 NotebookLM' : '同步試算表失敗', String(req.key || ''), r.ok ? String(req.title || '') : r.error]]); return json_(r); }
    if (req.action === 'nlmexport') { const r = nlmExport_(req, me_); addLog_([[null, me_.id, me_.name, r.ok ? '匯出到 NotebookLM' : '匯出到 NotebookLM 失敗', String(req.key || ''), r.ok ? String(req.title || '') : r.error]]); return json_(r); }
    if (req.action === 'trash') { const r = trash_(req, me_); addLog_([[null, me_.id, me_.name, r.ok ? '刪除檔案' : '刪除檔案（拒絕）', String(req.id || ''), r.ok ? '' : r.error]]); return json_(r); }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const { docs, rows } = readAll_();
    const sh = sheet_();
    const denied = [], logRows = [];
    (req.writes || []).forEach(w => {
      if (!can_(who, w, U, docs)) { denied.push(w.id); logRows.push([null, me_.id, me_.name, '寫入（拒絕）', w.id, patchSummary_(w, docs[w.id])]); return; }
      logRows.push([null, me_.id, me_.name, '修改', w.id, patchSummary_(w, docs[w.id])]);
      const d = docs[w.id] || { cells: {}, items: {} };
      d.cells = d.cells || {}; d.items = d.items || {};
      Object.keys((w.patch && w.patch.cells) || {}).forEach(k => { d.cells[k] = w.patch.cells[k]; });
      const pfx = String(w.id).split('__')[0];
      Object.keys((w.patch && w.patch.items) || {}).forEach(k => {
        let v = w.patch.items[k];
        const old = d.items[k];
        // V2.5.1：已報到者以第一次報到時間為準，重複報到不覆蓋（避免變成遲到）
        if (v && (pfx === 'attn' || pfx === 'rta') && old && old.checkin && v.checkin && v.checkin !== old.checkin) v = Object.assign({}, v, { checkin: old.checkin, dev: old.dev });
        d.items[k] = v === null ? null : Object.assign({}, old || {}, v);
      });
      d.updatedAt = new Date().toISOString();
      docs[w.id] = d;
      const text = JSON.stringify(d);
      if (rows[w.id]) sh.getRange(rows[w.id], 2, 1, 2).setValues([[text, d.updatedAt]]);
      else { sh.appendRow([w.id, text, d.updatedAt]); rows[w.id] = sh.getLastRow(); }
    });
    docsBust_();
    addLog_(logRows);
    if (denied.length) return json_({ ok: false, denied: denied, error: '沒有權限修改其他組別或主持人設定的內容' });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
