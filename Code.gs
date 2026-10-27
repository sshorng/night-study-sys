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
/* 入口                                                                */
/* ------------------------------------------------------------------ */

function doGet(e) {
  return json_({
    ok: true,
    service: '夜自習管理系統後端',
    version: '1.0.0',
    sheets: Object.keys(SHEET_DEF).map(function (k) { return SHEET_DEF[k].name; }),
    templates: TPL_NAMES.map(function (n) { return { name: n, found: findTemplate_(n) !== null }; }),
    message: '請用 POST 呼叫，或直接開啟 index.html 使用。',
  });
}

function doPost(e) {
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
        return json_({ ok: true, pong: new Date().toISOString() });
      case 'loadAll':
        return json_({ ok: true, state: loadState_() });
      case 'saveAll':
        saveState_(req.state || {});
        return json_({ ok: true, savedAt: new Date().toISOString() });
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

/* ------------------------------------------------------------------ */
/* 工作表                                                              */
/* ------------------------------------------------------------------ */

function book_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('這份指令碼沒有綁定試算表，請用「擴充功能 > Apps Script」從試算表開啟');
  return ss;
}

function sheet_(key) {
  const def = SHEET_DEF[key];
  const ss = book_();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureSheets_() {
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
  const def = SHEET_DEF[key];
  const sh = sheet_(key);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, def.cols.length).getValues();
  return values.filter(function (r) {
    return r.some(function (c) { return c !== '' && c !== null; });
  });
}

function writeRows_(key, rows) {
  const def = SHEET_DEF[key];
  const sh = sheet_(key);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, def.cols.length).clearContent();
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, def.cols.length).setValues(rows);
}

/* ------------------------------------------------------------------ */
/* 讀取與寫入狀態                                                       */
/* ------------------------------------------------------------------ */

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

function saveState_(st) {
  ensureSheets_();
  const students = (st.students || []).map(function (s) {
    return [s.key, s.cls, s.seat, s.name, s.phone || '', s.lunch || '', s.parentDuty ? 'V' : '',
      s.note || '', s.days[0], s.days[1], s.days[2], s.days[3], s.days[4], s.active === false ? 'FALSE' : '',
      s.sid || ''];
  });
  const enrollments = (st.enrollments || []).map(function (e) {
    return [e.id, e.key, e.date, e.days[0], e.days[1], e.days[2], e.days[3], e.days[4], e.reason || '', e.createdAt || ''];
  });
  const calendar = (st.calendar || []).map(function (c) { return [c.date, c.status || '上課', c.label || '']; });
  const leaves = (st.leaves || []).map(function (l) { return [l.id, l.key, l.date, l.reason || '', l.paper ? 'V' : '']; });
  const violations = (st.violations || []).map(function (v) {
    return [v.id, v.key, v.date, v.reason || '', v.points || 1, v.paper ? 'V' : ''];
  });
  const checkins = (st.checkins || []).map(function (c) { return [c.date, c.key, c.status || '']; });
  const seatPair = st.seatPair || {};
  const seatKeys = Object.keys(st.seats || {}).concat(Object.keys(seatPair))
    .filter(function (k, i, a) { return a.indexOf(k) === i && (st.seats[k] || seatPair[k]); })
    .sort(function (a, b) { return Number(a) - Number(b); });
  const seatRows = seatKeys.map(function (k) { return [Number(k), (st.seats || {})[k] || '', seatPair[k] || '']; });
  const duty = (st.duty || []).map(function (d) { return [d.date, d.teacher || '', d.note || '']; });
  const dutyNotes = (st.dutyNotes || []).map(function (n) {
    return [n.id, n.date, n.time || '', n.text || '', n.key || ''];
  });
  /* 設定工作表要保留不認識的列（例如 password），只更新以下幾項 */
  const cfgMap = {};
  readRows_('config').forEach(function (r) { if (r[0]) cfgMap[String(r[0])] = r[1]; });
  cfgMap.meta = JSON.stringify(st.meta || {});
  cfgMap.config = JSON.stringify(st.config || {});
  cfgMap.conflicts = JSON.stringify(st.conflicts || []);
  cfgMap.updatedAt = new Date().toISOString();
  /* 可調參數逐列寫回「設定」工作表，前端與試算表兩邊都能改 */
  const optRows = st.opts || {};
  Object.keys(optRows).forEach(function (k) {
    const key = String(k || '').trim();
    if (!key || CFG_SYS_KEYS[key]) return;
    cfgMap[key] = optRows[k] === null || optRows[k] === undefined ? '' : String(optRows[k]);
  });
  const configRows = Object.keys(cfgMap).map(function (k) { return [k, cfgMap[k]]; });

  writeRows_('students', students);
  writeRows_('enrollments', enrollments);
  writeRows_('calendar', calendar);
  writeRows_('leaves', leaves);
  writeRows_('violations', violations);
  writeRows_('checkins', checkins);
  writeRows_('seats', seatRows);
  writeRows_('duty', duty);
  writeRows_('dutyNotes', dutyNotes);
  writeRows_('config', configRows);
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
  const file = findTemplate_(name);
  if (!file) {
    throw new Error('找不到範本檔「' + name + '」，請先把 templates 資料夾內的三個檔案上傳到雲端硬碟');
  }
  return Utilities.base64Encode(file.getBlob().getBytes());
}

/** 第一次使用時執行：建立工作表並檢查範本是否都在 */
function 初次設定() {
  ensureSheets_();
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

/** 讀通關密語：先找試算表「設定」工作表的 password 列，其次 Apps Script 指令碼屬性 */
function password_() {
  try {
    const row = readRows_('config').filter(function (r) {
      return String(r[0]).trim() === PWD_KEY;
    })[0];
    if (row && String(row[1] || '').trim()) return String(row[1]).trim();
  } catch (err) { /* 工作表還沒建立時往下找指令碼屬性 */ }
  return PropertiesService.getScriptProperties().getProperty('PASSWORD') || '';
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
