/*************************************************************
 * 夜自習管理系統　後端（Google Apps Script）
 *
 * 資料放在這份指令碼綁定的 Google 試算表，前端為 index.html。
 * 部署方式與操作步驟見 docs/DEPLOY_GUIDE.md。
 *************************************************************/

/** 通關密語：存在試算表「設定」工作表的一列（A 欄 password、B 欄你的密語），
 *  或在 Apps Script「專案設定 → 指令碼屬性」新增 PASSWORD。程式裡不寫死任何密語。 */
const PWD_KEY = 'password';

/** 「設定」工作表裡這幾列是系統用的，其餘每一列都是可調參數（前端「設定與檢查」頁也會寫回這裡）。
 *  參數以「項目／值」兩欄呈現，直接在試算表改完，前端按「從雲端重新載入」就生效。 */
const CFG_SYS_KEYS = { password: 1, meta: 1, config: 1, conflicts: 1, updatedAt: 1 };

/** 範本檔名（上傳到雲端硬碟後，用檔名尋找） */
const TPL_NAMES = ['座位表-範本.docx', '違規單-範本.docx', '名單範本.xlsx'];

/** 可選：指定範本所在資料夾的 ID，留空則用檔名全雲端搜尋 */
const TPL_FOLDER_ID = '';

const SHEET_DEF = {
  students: {
    name: '名單',
    cols: ['key', '班級', '座號', '姓名', '電話', '便當', '家長輪值', '備註', '一', '二', '三', '四', '五', '啟用', '學號'],
  },
  enrollments: { name: '異動', cols: ['id', 'key', '生效日', '一', '二', '三', '四', '五', '原因', '建立時間'] },
  calendar: { name: '日曆', cols: ['日期', '狀態', '說明'] },
  leaves: { name: '請假', cols: ['id', 'key', '日期', '事由', '已繳紙本'] },
  violations: { name: '違規', cols: ['id', 'key', '日期', '事由', '點數', '已繳紙本'] },
  checkins: { name: '點名', cols: ['日期', 'key', '狀態'] },
  seats: { name: '座位', cols: ['格號', 'key', '下位'] },
  duty: { name: '輪值', cols: ['日期', '輪值教師', '記事'] },
  dutyNotes: { name: '輪值記事', cols: ['id', '日期', '時間', '內容', '學生'] },
  config: { name: '設定', cols: ['項目', '值'] },
};
/* ------------------------------------------------------------------ */
/* 效能：回應快取與資料世代                                            */
/* ------------------------------------------------------------------ */

/* loadAll 的整包回應存在 Apps Script 快取（分片，單鍵上限約 100 KB）。
   每次寫入會把「資料世代」往前推，舊世代的快取不再被引用，因此不會讀到舊資料。 */
const CACHE_PREFIX = 'ns_data';
const CACHE_SCHEMA = 'v2';
/* 快取存活 90 秒：老師可能直接在試算表改資料，所以不能放太久；
   程式自己存檔會把「世代」往前推，快取立刻失效，不必等逾時 */
const CACHE_TTL_SEC = 90;
const CACHE_CHUNK_BYTES = 90000;
const CACHE_MAX_BYTES = 900000;
const GEN_PROP = 'DATA_GENERATION';

/* 前端項目 → 工作表：只寫真的有變動的那幾張 */
const PARTS_TO_SHEET = {
  students: 'students', enrollments: 'enrollments', calendar: 'calendar', leaves: 'leaves',
  violations: 'violations', checkins: 'checkins', seats: 'seats', seatPair: 'seats',
  duty: 'duty', dutyNotes: 'dutyNotes',
  meta: 'config', opts: 'config', config: 'config', conflicts: 'config',
};

function generation_() {
  try {
    return PropertiesService.getScriptProperties().getProperty(GEN_PROP) || '0';
  } catch (err) { return '0'; }
}

function bumpGeneration_() {
  const next = String(Date.now());
  try { PropertiesService.getScriptProperties().setProperty(GEN_PROP, next); } catch (err) { /* 寫不進去時仍回新值 */ }
  return next;
}

function cacheKeyFor_(gen) { return CACHE_PREFIX + '_' + CACHE_SCHEMA + '_' + String(gen || '0'); }

/* 大包資料切成 90 KB 一片，避免超過 CacheService 單鍵上限 */
function putCacheChunked_(key, text, ttl) {
  const cache = CacheService.getScriptCache();
  const parts = Math.ceil(text.length / CACHE_CHUNK_BYTES) || 1;
  const batch = {};
  batch[key + '_n'] = String(parts);
  for (let i = 0; i < parts; i++) {
    batch[key + '_p' + i] = text.substring(i * CACHE_CHUNK_BYTES, (i + 1) * CACHE_CHUNK_BYTES);
  }
  cache.putAll(batch, ttl || CACHE_TTL_SEC);
}

function getCacheChunked_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const total = parseInt(cache.get(key + '_n') || '0', 10) || 0;
    if (total < 1) return '';
    const keys = [];
    for (let i = 0; i < total; i++) keys.push(key + '_p' + i);
    const got = cache.getAll(keys);
    let text = '';
    for (let i = 0; i < total; i++) {
      const piece = got[key + '_p' + i];
      if (!piece) return '';   /* 有一片掉了就當作沒有快取 */
      text += piece;
    }
    return text;
  } catch (err) { return ''; }
}

function removeCacheChunked_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const total = parseInt(cache.get(key + '_n') || '0', 10) || 0;
    const keys = [key + '_n'];
    for (let i = 0; i < total; i++) keys.push(key + '_p' + i);
    cache.removeAll(keys);
  } catch (err) { /* 清不掉也不影響正確性，舊世代本來就不會被讀到 */ }
}

/* ------------------------------------------------------------------ */
/* 入口                                                                */
/* ------------------------------------------------------------------ */

function doGet(e) {
  /* 只有這段會去搜雲端硬碟，忙線時不要讓整支 API 掛掉 */
  let templates = [];
  try {
    templates = TPL_NAMES.map(function (n) { return { name: n, found: findTemplate_(n) !== null }; });
  } catch (err) {
    templates = TPL_NAMES.map(function (n) { return { name: n, found: null }; });
  }
  return json_({
    ok: true,
    service: '夜自習管理系統後端',
    version: '1.2.0',
    sheets: Object.keys(SHEET_DEF).map(function (k) { return SHEET_DEF[k].name; }),
    templates: templates,
    message: '請用 POST 呼叫，或直接開啟 index.html 使用。',
  });
}

function doPost(e) {
  resetMemo_();
  let req = {};
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '請求格式錯誤（不是合法 JSON）' });
  }
  const expectPwd = password_();
  if (!expectPwd) {
    return json_({
      ok: false,
      error: '後端還沒設定通關密語：請在試算表「設定」工作表新增一列（A 欄 password、B 欄你的密語），'
        + '或在 Apps Script「專案設定 → 指令碼屬性」新增 PASSWORD',
    });
  }
  if (String(req.password || '') !== expectPwd) {
    return json_({ ok: false, error: '通關密語不正確' });
  }
  try {
    switch (req.action) {
      case 'ping':
        return json_({ ok: true, pong: new Date().toISOString(), generation: generation_() });
      case 'loadAll': {
        const loadT0 = Date.now();
        const gen = generation_();
        /* 前端手上就是這一版：不讀表、也不傳資料，只回一句「沒變」 */
        if (!req.fresh && req.generation && String(req.generation) === gen) {
          return json_({ ok: true, notModified: true, generation: gen, serverMs: Date.now() - loadT0 });
        }
        const cached = req.fresh === true ? '' : getCacheChunked_(cacheKeyFor_(gen));
        if (cached) {
          /* 快取命中：把先前存好的 JSON 原樣送出，省掉十張表的讀取與重新組裝 */
          return jsonRaw_('{"ok":true,"cached":true,"generation":"' + gen + '","serverMs":'
            + (Date.now() - loadT0) + ',"state":' + cached + '}');
        }
        const state = loadState_();
        const packed = JSON.stringify(state);
        if (packed.length <= CACHE_MAX_BYTES) {
          try { putCacheChunked_(cacheKeyFor_(gen), packed, CACHE_TTL_SEC); } catch (errPut) { /* 放不進快取就照常回資料 */ }
        }
        return json_({ ok: true, state: state, generation: gen, cached: false, serverMs: Date.now() - loadT0 });
      }
      case 'saveAll': {
        const saveT0 = Date.now();
        const before = generation_();
        const result = saveState_(req);
        const gen = bumpGeneration_();
        removeCacheChunked_(cacheKeyFor_(before));
        return json_({
          ok: true, savedAt: new Date().toISOString(), generation: gen,
          wrote: result.wrote, serverMs: Date.now() - saveT0,
        });
      }
      case 'getTemplate':
        return json_({ ok: true, name: req.name, base64: getTemplateBase64_(req.name) });
      default:
        return json_({ ok: false, error: '未知的動作：' + req.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* 已經是 JSON 字串時直接送出，省掉一次 parse 再 stringify */
function jsonRaw_(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* 工作表                                                              */
/* ------------------------------------------------------------------ */

/* 同一次執行內只讀一次試算表與工作表：Apps Script 每次呼叫都有成本，重複讀取是慢的主因 */
let MEMO = { ss: null, sheets: {}, rows: {}, pwd: undefined };
function resetMemo_() { MEMO = { ss: null, sheets: {}, rows: {}, pwd: undefined }; }

function book_() {
  if (MEMO.ss) return MEMO.ss;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('這份指令碼沒有綁定試算表，請用「擴充功能 > Apps Script」從試算表開啟');
  MEMO.ss = ss;
  return ss;
}

function writeHeader_(sh, def) {
  sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function sheet_(key) {
  const def = SHEET_DEF[key];
  if (MEMO.sheets[key]) return MEMO.sheets[key];
  const ss = book_();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    writeHeader_(sh, def);
    MEMO.rows[key] = [];
  }
  MEMO.sheets[key] = sh;
  return sh;
}

/* 只補缺的工作表。欄位標題的校正留給「初次設定」，平常讀寫不走這條，省掉十幾次往返 */
function ensureSheets_() {
  const ss = book_();
  const have = {};
  ss.getSheets().forEach(function (sh) { have[sh.getName()] = true; });
  let created = 0;
  Object.keys(SHEET_DEF).forEach(function (k) {
    const def = SHEET_DEF[k];
    if (have[def.name]) return;
    const sh = ss.insertSheet(def.name);
    writeHeader_(sh, def);
    MEMO.sheets[k] = sh;
    MEMO.rows[k] = [];
    created++;
  });
  return created;
}

/* 欄位標題校正：手動跑「初次設定」時才檢查，平時不多花往返 */
function verifyHeaders_() {
  Object.keys(SHEET_DEF).forEach(function (k) {
    const def = SHEET_DEF[k];
    const sh = sheet_(k);
    const width = Math.max(sh.getLastColumn(), def.cols.length);
    const head = sh.getRange(1, 1, 1, width).getValues()[0].map(function (v) { return String(v || '').trim(); });
    let need = false;
    for (let i = 0; i < def.cols.length; i++) {
      if (head[i] !== String(def.cols[i])) { need = true; break; }
    }
    if (need) sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]);
  });
}

function readRows_(key) {
  if (MEMO.rows[key]) return MEMO.rows[key];
  const def = SHEET_DEF[key];
  const sh = sheet_(key);
  const last = sh.getLastRow();
  let rows = [];
  if (last >= 2) {
    rows = sh.getRange(2, 1, last - 1, def.cols.length).getValues().filter(function (r) {
      return r.some(function (c) { return c !== '' && c !== null; });
    });
  }
  MEMO.rows[key] = rows;
  return rows;
}

/* 兩份資料是不是一樣：一樣就整張不寫，寫入是整個流程最貴的一段（可省下大半時間） */
function sameRows_(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const x = a[i] || [], y = b[i] || [];
    for (let j = 0; j < y.length; j++) {
      const p = x[j] === null || x[j] === undefined ? '' : String(x[j]);
      const q = y[j] === null || y[j] === undefined ? '' : String(y[j]);
      if (p !== q) return false;
    }
  }
  return true;
}

/* 每張表用哪一欄當主鍵；沒有主鍵的表（設定）就整表比對後重寫 */
const ROW_KEY_COLS = {
  students: 0, enrollments: 0, calendar: 0, leaves: 0, violations: 0,
  seats: 0, duty: 0, dutyNotes: 0,
};

function rowKeyOf_(key, row) {
  /* 點名表沒有單一主鍵，用「日期＋學生」當鍵 */
  if (key === 'checkins') {
    return String(row[0] == null ? '' : row[0]) + '|' + String(row[1] == null ? '' : row[1]);
  }
  const col = ROW_KEY_COLS[key];
  if (col === undefined) return '';
  return String(row[col] == null ? '' : row[col]);
}

function sameRow_(a, b) {
  const n = Math.max(a.length, b.length);
  for (let j = 0; j < n; j++) {
    const p = a[j] === null || a[j] === undefined ? '' : String(a[j]);
    const q = b[j] === null || b[j] === undefined ? '' : String(b[j]);
    if (p !== q) return false;
  }
  return true;
}

function rewriteRows_(key, sh, old, rows) {
  const def = SHEET_DEF[key];
  const last = sh.getLastRow();
  if (rows.length) sh.getRange(2, 1, rows.length, def.cols.length).setValues(rows);
  /* 資料變少時才需要清掉後面多出來的列 */
  if (last - 1 > rows.length) {
    sh.getRange(2 + rows.length, 1, last - 1 - rows.length, def.cols.length).clearContent();
  }
  return true;
}

/* 有主鍵的表只寫真正變動的那幾列。
   例：點名表累積到 6000 列時，點一位學生只寫 1 列，不必整張重寫。 */
function writeRowsByKey_(key, sh, old, rows) {
  const def = SHEET_DEF[key];
  const width = def.cols.length;
  /* 表上夾雜空白列時位置會對不上，改走整表重寫比較安全 */
  const last = sh.getLastRow();
  if (last - 1 !== old.length) return false;
  const pos = {};
  old.forEach(function (r, i) {
    const k = rowKeyOf_(key, r);
    if (k) pos[k] = i;
  });
  const keep = {};
  const updates = [];
  const appends = [];
  rows.forEach(function (r) {
    const k = rowKeyOf_(key, r);
    const at = pos[k];
    if (at === undefined) { appends.push(r); return; }
    keep[at] = 1;
    if (!sameRow_(old[at], r)) updates.push({ at: at, values: r });
  });
  const dels = [];
  old.forEach(function (r, i) {
    if (keep[i]) return;
    if (!rowKeyOf_(key, r)) return;
    dels.push(i);
  });
  const changed = updates.length + appends.length + dels.length;
  /* 變動範圍太大時整表重寫反而快，也避免幾百次刪列 */
  if (dels.length > 30 || changed > Math.max(60, Math.round(rows.length * 0.5))) return false;

  /* 更新列：連續的合併成一次 setValues */
  updates.sort(function (a, b) { return a.at - b.at; });
  let i = 0;
  while (i < updates.length) {
    let j = i;
    while (j + 1 < updates.length && updates[j + 1].at === updates[j].at + 1) j++;
    const block = [];
    for (let x = i; x <= j; x++) block.push(updates[x].values);
    sh.getRange(updates[i].at + 2, 1, block.length, width).setValues(block);
    i = j + 1;
  }
  /* 刪列：由下往上刪，上面的列號才不會位移 */
  dels.sort(function (a, b) { return b - a; });
  for (let d = 0; d < dels.length; d++) sh.deleteRow(dels[d] + 2);
  /* 新增列：接在最後一筆之後 */
  if (appends.length) {
    sh.getRange(2 + (rows.length - appends.length), 1, appends.length, width).setValues(appends);
  }
  return true;
}

function writeRows_(key, rows) {
  const sh = sheet_(key);
  const old = readRows_(key);
  if (sameRows_(old, rows)) return false;
  const useKey = key === 'checkins' || ROW_KEY_COLS[key] !== undefined;
  /* 小表整表重寫最省事；大表才走列級差異寫入 */
  if (useKey && rows.length && old.length && (rows.length + old.length) > 40) {
    if (!writeRowsByKey_(key, sh, old, rows)) rewriteRows_(key, sh, old, rows);
  } else {
    rewriteRows_(key, sh, old, rows);
  }
  MEMO.rows[key] = rows;
  return true;
}

function loadState_() {
  ensureSheets_();
  const cfg = {};
  readRows_('config').forEach(function (r) { cfg[r[0]] = r[1]; });
  const opts = {};
  Object.keys(cfg).forEach(function (k) {
    const key = String(k || '').trim();
    if (!key || CFG_SYS_KEYS[key]) return;
    opts[key] = cfg[k] === null || cfg[k] === undefined ? '' : String(cfg[k]);
  });

  const students = readRows_('students').map(function (r) {
    return {
      key: String(r[0]), cls: Number(r[1]), seat: pad2_(r[2]), name: String(r[3] || ''),
      phone: String(r[4] || ''), lunch: String(r[5] || ''), parentDuty: r[6] === true || String(r[6]).toUpperCase() === 'V',
      note: String(r[7] || ''),
      days: [num_(r[8]), num_(r[9]), num_(r[10]), num_(r[11]), num_(r[12])],
      active: !(r[13] === false || String(r[13]) === 'FALSE'),
      sid: String(r[14] || ''),
    };
  });

  const enrollments = readRows_('enrollments').map(function (r) {
    return {
      id: String(r[0]), key: String(r[1]), date: dateStr_(r[2]),
      days: [num_(r[3]), num_(r[4]), num_(r[5]), num_(r[6]), num_(r[7])],
      reason: String(r[8] || ''), createdAt: String(r[9] || ''),
    };
  });

  const calendar = readRows_('calendar').map(function (r) {
    return { date: dateStr_(r[0]), status: String(r[1] || '上課'), label: String(r[2] || '') };
  });

  const leaves = readRows_('leaves').map(function (r) {
    return {
      id: String(r[0]), key: String(r[1]), date: dateStr_(r[2]), reason: String(r[3] || ''),
      paper: r[4] === true || String(r[4]).toUpperCase() === 'V',
    };
  });

  const violations = readRows_('violations').map(function (r) {
    return {
      id: String(r[0]), key: String(r[1]), date: dateStr_(r[2]), reason: String(r[3] || ''),
      points: Number(r[4] || 1),
      paper: r[5] === true || String(r[5]).toUpperCase() === 'V',
    };
  });

  const checkins = readRows_('checkins').map(function (r) {
    return { date: dateStr_(r[0]), key: String(r[1]), status: String(r[2] || '') };
  });

  const seats = {};
  const seatPair = {};
  readRows_('seats').forEach(function (r) {
    if (r[0] === '' || r[0] === null) return;
    if (r[1]) seats[String(r[0])] = String(r[1]);
    if (r[2]) seatPair[String(r[0])] = String(r[2]);
  });

  const duty = readRows_('duty').map(function (r) {
    return { date: dateStr_(r[0]), teacher: String(r[1] || ''), note: String(r[2] || '') };
  });

  const dutyNotes = readRows_('dutyNotes').map(function (r) {
    return {
      id: String(r[0]), date: dateStr_(r[1]), time: timeStr_(r[2]),
      text: String(r[3] || ''), key: String(r[4] || ''),
    };
  });

  return {
    meta: safeJson_(cfg.meta, { school: '', term: '' }),
    opts: opts,
    config: safeJson_(cfg.config, { refDate: '', totals: {} }),
    students: students,
    enrollments: enrollments,
    calendar: calendar,
    leaves: leaves,
    violations: violations,
    checkins: checkins,
    seats: seats,
    seatPair: seatPair,
    duty: duty,
    dutyNotes: dutyNotes,
    conflicts: safeJson_(cfg.conflicts, []),
  };
}

/* 寫入：state 是完整狀態；parts 是這次真正有變動的項目，只寫對應的工作表。
   沒帶 parts 就是全寫（第一次存檔或整包匯入用）。 */
function saveState_(req) {
  ensureSheets_();
  const st = req.state || {};
  const parts = Array.isArray(req.parts) && req.parts.length ? req.parts : null;
  const wantSheet = function (sheetKey) {
    if (!parts) return true;
    for (let i = 0; i < parts.length; i++) {
      if (PARTS_TO_SHEET[parts[i]] === sheetKey) return true;
    }
    return false;
  };
  const wrote = [];
  /* 只寫這次真的有送來資料的表，避免半套 payload 把整張表清空 */
  const ready = function (sheetKey, names) {
    if (!wantSheet(sheetKey)) return false;
    if (!parts) return true;
    for (let i = 0; i < names.length; i++) {
      if (st[names[i]] === undefined) return false;
    }
    return true;
  };

  if (ready('students', ['students'])) {
    const students = (st.students || []).map(function (s) {
      const d = s.days || [0, 0, 0, 0, 0];
      return [s.key, s.cls, s.seat, s.name, s.phone || '', s.lunch || '', s.parentDuty ? 'V' : '',
        s.note || '', d[0], d[1], d[2], d[3], d[4], s.active === false ? 'FALSE' : '', s.sid || ''];
    });
    if (writeRows_('students', students)) wrote.push('students');
  }
  if (ready('enrollments', ['enrollments'])) {
    const rows = (st.enrollments || []).map(function (e) {
      const d = e.days || [0, 0, 0, 0, 0];
      return [e.id, e.key, e.date, d[0], d[1], d[2], d[3], d[4], e.reason || '', e.createdAt || ''];
    });
    if (writeRows_('enrollments', rows)) wrote.push('enrollments');
  }
  if (ready('calendar', ['calendar'])) {
    const rows = (st.calendar || []).map(function (c) { return [c.date, c.status || '上課', c.label || '']; });
    if (writeRows_('calendar', rows)) wrote.push('calendar');
  }
  if (ready('leaves', ['leaves'])) {
    const rows = (st.leaves || []).map(function (l) { return [l.id, l.key, l.date, l.reason || '', l.paper ? 'V' : '']; });
    if (writeRows_('leaves', rows)) wrote.push('leaves');
  }
  if (ready('violations', ['violations'])) {
    const rows = (st.violations || []).map(function (v) {
      return [v.id, v.key, v.date, v.reason || '', v.points || 1, v.paper ? 'V' : ''];
    });
    if (writeRows_('violations', rows)) wrote.push('violations');
  }
  if (ready('checkins', ['checkins'])) {
    const rows = (st.checkins || []).map(function (c) { return [c.date, c.key, c.status || '']; });
    if (writeRows_('checkins', rows)) wrote.push('checkins');
  }
  if (ready('seats', ['seats', 'seatPair'])) {
    const seatPair = st.seatPair || {};
    const seatKeys = Object.keys(st.seats || {}).concat(Object.keys(seatPair))
      .filter(function (k, i, a) { return a.indexOf(k) === i && (st.seats[k] || seatPair[k]); })
      .sort(function (a, b) { return Number(a) - Number(b); });
    const rows = seatKeys.map(function (k) { return [Number(k), (st.seats || {})[k] || '', seatPair[k] || ''] });
    if (writeRows_('seats', rows)) wrote.push('seats');
  }
  if (ready('duty', ['duty'])) {
    const rows = (st.duty || []).map(function (d) { return [d.date, d.teacher || '', d.note || '']; });
    if (writeRows_('duty', rows)) wrote.push('duty');
  }
  if (ready('dutyNotes', ['dutyNotes'])) {
    const rows = (st.dutyNotes || []).map(function (n) {
      return [n.id, n.date, n.time || '', n.text || '', n.key || ''];
    });
    if (writeRows_('dutyNotes', rows)) wrote.push('dutyNotes');
  }
  /* 設定工作表：保留不認識的列（例如 password），只改這次有送來的項目 */
  if (wantSheet('config')) {
    const cfgMap = {};
    readRows_('config').forEach(function (r) { if (r[0]) cfgMap[String(r[0])] = r[1]; });
    if (st.meta !== undefined) cfgMap.meta = JSON.stringify(st.meta || {});
    if (st.config !== undefined) cfgMap.config = JSON.stringify(st.config || {});
    if (st.conflicts !== undefined) cfgMap.conflicts = JSON.stringify(st.conflicts || []);
    if (st.opts !== undefined) {
      Object.keys(st.opts || {}).forEach(function (k) {
        const key = String(k || '').trim();
        if (!key || CFG_SYS_KEYS[key]) return;
        const v = st.opts[k];
        cfgMap[key] = v === null || v === undefined ? '' : String(v);
      });
    }
    cfgMap.updatedAt = new Date().toISOString();
    const configRows = Object.keys(cfgMap).map(function (k) { return [k, cfgMap[k]]; });
    if (writeRows_('config', configRows)) wrote.push('config');
  }

  return { wrote: wrote };
}

/* ------------------------------------------------------------------ */
/* 範本檔                                                              */
/* ------------------------------------------------------------------ */

function findTemplate_(name) {
  if (TPL_FOLDER_ID) {
    try {
      const it = DriveApp.getFolderById(TPL_FOLDER_ID).getFilesByName(name);
      if (it.hasNext()) return it.next();
    } catch (err) { /* 資料夾失效時退回全雲端搜尋 */ }
  }
  const it = DriveApp.getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function getTemplateBase64_(name) {
  const cache = CacheService.getScriptCache();
  const ck = 'tpl64_' + name;
  try {
    const hit = cache.get(ck);
    if (hit) return hit;
  } catch (err) { /* 快取讀不到就照常回雲端硬碟拿 */ }
  const file = findTemplate_(name);
  if (!file) {
    throw new Error('找不到範本檔「' + name + '」，請先把 templates 資料夾內的三個檔案上傳到雲端硬碟');
  }
  const b64 = Utilities.base64Encode(file.getBlob().getBytes());
  try {
    /* 快取單鍵上限約 100 KB，太大的檔案就每次回硬碟拿 */
    if (b64.length < 95000) cache.put(ck, b64, 21600);
  } catch (err) { /* 放不進快取就算了，功能不受影響 */ }
  return b64;
}

/** 第一次使用時執行：建立工作表並檢查範本是否都在 */
function 初次設定() {
  ensureSheets_();
  verifyHeaders_();
  resetMemo_();
  const report = TPL_NAMES.map(function (n) {
    return n + '：' + (findTemplate_(n) ? '找到' : '尚未上傳');
  });
  SpreadsheetApp.getUi().alert('工作表已建立\n\n' + report.join('\n')
    + '\n\n通關密語：' + (password_() ? '已設定' : '尚未設定，請在「設定」工作表新增一列（A 欄 password、B 欄你的密語）')
    + '\n\n可調參數：也在「設定」工作表，一項一列（A 欄項目、B 欄值）。'
    + '\n學校名稱、學年度與學期、退班點數門檻、違規事由清單、請假事由清單、便當選項、匯出檔名前綴，'
    + '\n改完存檔，回到網頁按「從雲端重新載入」就生效。');
}

/* ------------------------------------------------------------------ */
/* 小工具                                                              */
/* ------------------------------------------------------------------ */

/** 讀通關密語：先找試算表「設定」工作表的 password 列，其次 Apps Script 指令碼屬性
 *  同一次執行只讀一次，存檔流程不會為了密語再讀一次工作表 */
function password_() {
  if (MEMO.pwd !== undefined) return MEMO.pwd;
  let pwd = '';
  try {
    const row = readRows_('config').filter(function (r) {
      return String(r[0]).trim() === PWD_KEY;
    })[0];
    if (row && String(row[1] || '').trim()) pwd = String(row[1]).trim();
  } catch (err) { /* 工作表還沒建立時往下找指令碼屬性 */ }
  if (!pwd) pwd = PropertiesService.getScriptProperties().getProperty('PASSWORD') || '';
  MEMO.pwd = pwd;
  return pwd;
}

function pad2_(v) {
  const s = String(v === null || v === undefined ? '' : v).trim();
  if (/^\d+$/.test(s)) return s.length < 2 ? '0' + s : s;
  return s;
}

function num_(v) {
  return Number(v) === 1 ? 1 : 0;
}

function dateStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  const m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) {
    return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  }
  return s;
}

function timeStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  if (typeof v === 'number') {
    const mins = Math.round(v * 24 * 60);
    return ('0' + Math.floor(mins / 60)).slice(-2) + ':' + ('0' + (mins % 60)).slice(-2);
  }
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
  return s;
}

function safeJson_(text, fallback) {
  if (!text) return fallback;
  try {
    return JSON.parse(text);
  } catch (err) {
    return fallback;
  }
}
