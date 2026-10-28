/**
 * ระบบติดตามสถานะการสอบเทียบเครื่องมือวัด (Tool Calibration Tracking System)
 * กลุ่มวิศวกรรม กองบริหารการบินเกษตร
 *
 * Code.gs — ฝั่งเซิร์ฟเวอร์ (Google Apps Script ผูกกับ Google Sheet ฐานข้อมูล)
 *  - เมนู "เครื่องมือวัด" ในชีต: ตั้งค่าเริ่มต้นระบบ / ซ่อมสูตร / ส่งรายงานแจ้งเตือน / ตั้งเวลาแจ้งเตือน
 *  - API: doGet()/doPost() ให้หน้าเว็บ index.html (วางบน GitHub Pages) อ่าน/เขียนข้อมูล
 */

var TZ = 'Asia/Bangkok';

var SHEET = {
  TOOLS: 'ทะเบียนเครื่องมือ',
  HIST: 'ประวัติสอบเทียบ',
  USAGE: 'บันทึกการใช้งาน',
  SETTINGS: 'ตั้งค่า',
  RETIRED: 'เครื่องมือที่จำหน่ายแล้ว'
};

// คอลัมน์ของแท็บทะเบียนเครื่องมือ (ลำดับนี้คือลำดับในชีต — ห้ามสลับ ถ้าจะเพิ่มให้ต่อท้าย)
var TOOL_COLS = [
  ['tool_id',       'Tool ID'],
  ['group',         'กลุ่มงาน'],
  ['description',   'ชื่อรายการ'],
  ['model',         'รุ่น (Model)'],
  ['part_no',       'P/N'],
  ['serial_no',     'S/N'],
  ['asset_no',      'เลขครุภัณฑ์'],
  ['location',      'สถานที่ / หน่วยถือครอง'],
  ['cal_interval',  'รอบสอบเทียบ (เดือน)'],
  ['last_cal',      'วันที่สอบเทียบล่าสุด'],
  ['next_due',      'ครบกำหนดครั้งถัดไป'],
  ['cal_result',    'ผลสอบเทียบล่าสุด'],
  ['cert_no',       'เลขที่ใบรับรอง'],
  ['cert_link',     'ลิงก์ใบรับรอง'],
  ['cal_lab',       'ห้องปฏิบัติการ'],
  ['state',         'สถานะพิเศษ'],
  ['status',        'สถานะ (คำนวณอัตโนมัติ)'],
  ['cal_cost',      'ราคาสอบเทียบ (บาท)'],
  ['legacy_plan',   'แผนเดิม ปีงบ 65–70'],
  ['legacy_status', 'สถานะเดิม (ข้อมูล ก.ค. 66)'],
  ['remarks',       'หมายเหตุ'],
  ['photo',         'รูปภาพ (ลิงก์)']
];
var HIST_COLS = ['บันทึกเมื่อ', 'Tool ID', 'วันที่สอบเทียบ', 'ผลสอบเทียบ', 'เลขที่ใบรับรอง', 'ลิงก์ใบรับรอง', 'ห้องปฏิบัติการ', 'รอบ (เดือน)', 'ครบกำหนดครั้งถัดไป', 'ผู้ตรวจรับ', 'หมายเหตุ'];
var USAGE_COLS = ['วันเวลา', 'Tool ID', 'สถานะขณะใช้งาน', 'ผู้ใช้งาน', 'ทะเบียนอากาศยาน', 'เลขที่ใบสั่งงาน (Job Card)', 'หมายเหตุ'];

var RESULTS = ['ผ่าน', 'ไม่ผ่าน', 'เกินเกณฑ์ (OOT)'];
var STATES = ['ส่งสอบเทียบ', 'ส่งซ่อม', 'ชำรุด', 'ไม่ต้องสอบเทียบ (NCR)'];

var SETTING_ROWS = [
  ['ชื่อหน่วยงาน', 'กองบริหารการบินเกษตร กรมฝนหลวงและการบินเกษตร', ''],
  ['แจ้งเตือนล่วงหน้า (วัน)', 30, 'เครื่องมือที่เหลือไม่เกินจำนวนวันนี้จะแสดง "ใกล้ครบกำหนด"'],
  ['รหัสผู้ดูแล (PIN)', '', 'ใช้ยืนยันตอนบันทึกผลสอบเทียบ / เปลี่ยนสถานะผ่านเว็บแอป — ตั้งเองอย่างน้อย 4 หลัก'],
  ['อีเมลรับแจ้งเตือน', '', 'คั่นหลายอีเมลด้วยเครื่องหมายจุลภาค ว่าง = ส่งถึงเจ้าของสคริปต์'],
  ['ลิงก์หน้าเว็บ (สำหรับ QR Code)', '', 'ลิงก์ GitHub Pages ของระบบ เช่น https://ชื่อผู้ใช้.github.io/tool-cal/ ว่าง = ใช้ลิงก์ของหน้าที่เปิดพิมพ์ป้าย'],
  ['ขนาด QR บนป้าย (มม.)', 14, 'แนะนำ 14–16 มม.']
];

// ------------------------------------------------------------------ เมนูในชีต

function onOpen() {
  SpreadsheetApp.getUi().createMenu('เครื่องมือวัด')
    .addItem('1) ตั้งค่าเริ่มต้นระบบ (ครั้งแรก)', 'setup')
    .addItem('ซ่อมสูตรคำนวณ / จัดรูปแบบใหม่', 'refreshFormulas')
    .addSeparator()
    .addItem('ส่งรายงานแจ้งเตือนทางอีเมลตอนนี้', 'sendReportNow')
    .addItem('ตั้งเวลาส่งรายงานทุกวันจันทร์', 'installWeeklyTrigger')
    .addToUi();
}

/** สร้างแท็บทั้งหมด หัวตาราง สูตร และนำเข้าข้อมูลเครื่องมือเริ่มต้น (รันซ้ำได้ ไม่ลบข้อมูลเดิม) */
function setup() {
  requireSheetUi_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  var tools = ensureSheet_(ss, SHEET.TOOLS, TOOL_COLS.map(function (c) { return c[1]; }));
  ensureSheet_(ss, SHEET.HIST, HIST_COLS);
  ensureSheet_(ss, SHEET.USAGE, USAGE_COLS);
  var st = ss.getSheetByName(SHEET.SETTINGS);
  if (!st) {
    st = ss.insertSheet(SHEET.SETTINGS);
    st.getRange(1, 1, 1, 3).setValues([['รายการ', 'ค่า', 'คำอธิบาย']]).setFontWeight('bold').setBackground('#E8EDF2');
    st.getRange(2, 1, SETTING_ROWS.length, 3).setValues(SETTING_ROWS);
    st.setColumnWidth(1, 200); st.setColumnWidth(2, 320); st.setColumnWidth(3, 520);
    st.setFrozenRows(1);
  }

  // นำเข้าข้อมูลเริ่มต้นเฉพาะเมื่อทะเบียนยังว่าง
  var imported = 0;
  if (tools.getLastRow() < 2 && typeof SEED_TOOLS !== 'undefined') {
    var rows = SEED_TOOLS.map(function (s) {
      var o = {};
      SEED_KEYS.forEach(function (k, i) { o[k] = s[i]; });
      o.cal_interval = 12;
      return TOOL_COLS.map(function (c) { return o[c[0]] === undefined ? '' : o[c[0]]; });
    });
    // ป้องกัน S/N ที่ขึ้นต้นด้วย 0 หายไป
    tools.getRange(2, colOf_('serial_no'), rows.length, 1).setNumberFormat('@');
    tools.getRange(2, colOf_('tool_id'), rows.length, 1).setNumberFormat('@');
    tools.getRange(2, 1, rows.length, TOOL_COLS.length).setValues(rows);
    imported = rows.length;
  }

  // ลบแท็บเปล่าเริ่มต้น (ชีต1 / Sheet1) ถ้าไม่มีข้อมูล
  ['ชีต1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  refreshFormulas();
  ss.setActiveSheet(tools);

  var msg = 'ตั้งค่าเรียบร้อย' + (imported ? ' นำเข้าเครื่องมือ ' + imported + ' รายการ' : '') +
    '\n\nขั้นต่อไป:\n1) ไปที่แท็บ "ตั้งค่า" แล้วตั้งรหัสผู้ดูแล (PIN)\n2) Deploy เป็นเว็บแอป (ผู้มีสิทธิ์เข้าถึง: ทุกคน) แล้วนำลิงก์ไปใส่ใน index.html\n3) กรอกวันที่สอบเทียบล่าสุดจากใบรับรองของเครื่องมือแต่ละชิ้น';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/** ใส่สูตรวันครบกำหนด / สถานะ ให้ทุกแถว พร้อมตัวเลือก dropdown และสีสถานะ */
function refreshFormulas() {
  requireSheetUi_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET.TOOLS);
  if (!sh) throw new Error('ยังไม่มีแท็บ ' + SHEET.TOOLS + ' — กรุณารัน "ตั้งค่าเริ่มต้นระบบ" ก่อน');
  var last = Math.max(sh.getLastRow(), 2);
  var n = last - 1;

  var nextF = [], statF = [];
  for (var r = 2; r <= last; r++) {
    var f = rowFormulas_(r);
    nextF.push([f.next]);
    statF.push([f.status]);
  }
  sh.getRange(2, colOf_('next_due'), n, 1).setFormulas(nextF);
  sh.getRange(2, colOf_('status'), n, 1).setFormulas(statF);

  // รูปแบบ
  var hdr = sh.getRange(1, 1, 1, TOOL_COLS.length);
  hdr.setValues([TOOL_COLS.map(function (c) { return c[1]; })]);
  hdr.setFontWeight('bold').setBackground('#13202E').setFontColor('#FFFFFF').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  sh.getRange(2, colOf_('last_cal'), n, 2).setNumberFormat('d mmm yyyy');
  sh.getRange(2, colOf_('cal_cost'), n, 1).setNumberFormat('#,##0');
  sh.getRange(2, colOf_('next_due'), n, 1).setBackground('#F3F5F7');
  sh.getRange(2, colOf_('status'), n, 1).setFontWeight('bold').setHorizontalAlignment('center');
  var widths = { tool_id: 80, group: 160, description: 340, model: 110, serial_no: 110, asset_no: 190, location: 120,
    cal_interval: 90, last_cal: 120, next_due: 120, cal_result: 120, cert_link: 160, state: 150, status: 130,
    legacy_plan: 260, legacy_status: 120, remarks: 260, photo: 220 };
  Object.keys(widths).forEach(function (k) { sh.setColumnWidth(colOf_(k), widths[k]); });

  // Dropdown
  var maxR = Math.max(last + 200, 300);
  sh.getRange(2, colOf_('cal_result'), maxR, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(RESULTS, true).setAllowInvalid(false).build());
  sh.getRange(2, colOf_('state'), maxR, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATES, true).setAllowInvalid(false).build());
  sh.getRange(2, colOf_('last_cal'), maxR, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('กรอกเป็นวันที่ เช่น 15/8/2026 (ปี ค.ศ.)').build());

  // สีสถานะ
  var stRange = sh.getRange(2, colOf_('status'), maxR, 1);
  var rules = sh.getConditionalFormatRules().filter(function (rule) {
    return !rule.getRanges().some(function (rg) { return rg.getColumn() === colOf_('status'); });
  });
  [['พร้อมใช้งาน', '#E3F2EA', '#1E6B45'], ['ใกล้ครบกำหนด', '#FDF0D9', '#8A5300'],
   ['หมดอายุ', '#A8261B', '#FFFFFF'], ['ห้ามใช้งาน', '#A8261B', '#FFFFFF'], ['ส่งซ่อม', '#A8261B', '#FFFFFF'],
   ['ชำรุด', '#A8261B', '#FFFFFF'], ['ส่งสอบเทียบ', '#E3ECF6', '#1C4E80'], ['รอข้อมูล', '#ECEFF2', '#3A4652'],
   ['NCR', '#ECEFF2', '#3A4652']].forEach(function (c) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(c[0])
      .setBackground(c[1]).setFontColor(c[2]).setRanges([stRange]).build());
  });
  sh.setConditionalFormatRules(rules);

  [SHEET.HIST, SHEET.USAGE].forEach(function (name) {
    var s = ss.getSheetByName(name);
    if (!s) return;
    s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold').setBackground('#13202E').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  });
}

/** สูตรวันครบกำหนด / สถานะ ของแถว r */
function rowFormulas_(r) {
  var cI = colLetter_('cal_interval'), cL = colLetter_('last_cal'), cN = colLetter_('next_due'),
      cR = colLetter_('cal_result'), cS = colLetter_('state');
  var warnRef = "'" + SHEET.SETTINGS + "'!$B$3";
  return {
    next: '=IF(OR(' + cL + r + '="",' + cI + r + '=""),"",EDATE(' + cL + r + ',' + cI + r + '))',
    status: '=IF($A' + r + '="","",' +
      'IF(' + cS + r + '="ไม่ต้องสอบเทียบ (NCR)","NCR",' +
      'IF(' + cS + r + '<>"",' + cS + r + ',' +
      'IF(OR(' + cR + r + '="ไม่ผ่าน",' + cR + r + '="เกินเกณฑ์ (OOT)"),"ห้ามใช้งาน",' +
      'IF(' + cN + r + '="","รอข้อมูล",' +
      'IF(' + cN + r + '<TODAY(),"หมดอายุ",' +
      'IF(' + cN + r + '-TODAY()<=' + warnRef + ',"ใกล้ครบกำหนด","พร้อมใช้งาน")))))))'
  };
}

// ------------------------------------------------------------------ เว็บแอป

/**
 * API สำหรับหน้าเว็บ (index.html บน GitHub Pages)
 *  GET  ?fn=apiBoot | apiList | apiGetTool&id=M-001        → อ่านข้อมูล
 *  POST {"fn":"apiLogUsage","args":[...]} (Content-Type: text/plain) → เขียนข้อมูล
 * ตอบกลับเป็น JSON: {"ok":true,"data":...} หรือ {"ok":false,"error":"..."}
 */
var API_READ = { apiBoot: 0, apiList: 0, apiGetTool: 1 };
var API_WRITE = { apiLogUsage: 1, apiRecordCalibration: 2, apiSetState: 4, apiCheckPin: 1, apiUploadPhoto: 3, apiSaveTool: 2, apiRetireTool: 3 };
var PHOTO_FOLDER = 'Tool Calibration – รูปเครื่องมือ';

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.fn) return json_({ ok: true, data: 'Tool Calibration API — พร้อมใช้งาน' });
  if (!API_READ.hasOwnProperty(p.fn)) return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + p.fn });
  return run_(p.fn, p.fn === 'apiGetTool' ? [p.id] : []);
}

function doPost(e) {
  var body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'ข้อมูลไม่ถูกต้อง' }); }
  var fn = String(body.fn || '');
  if (!API_WRITE.hasOwnProperty(fn) && !API_READ.hasOwnProperty(fn)) return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + fn });
  var args = Array.isArray(body.args) ? body.args.slice(0, 4) : [];
  return run_(fn, args);
}

function run_(fn, args) {
  try {
    var fns = { apiBoot: apiBoot, apiList: apiList, apiGetTool: apiGetTool, apiLogUsage: apiLogUsage,
      apiRecordCalibration: apiRecordCalibration, apiSetState: apiSetState, apiCheckPin: apiCheckPin, apiUploadPhoto: apiUploadPhoto,
      apiSaveTool: apiSaveTool, apiRetireTool: apiRetireTool };
    return json_({ ok: true, data: fns[fn].apply(null, args) });
  } catch (err) {
    return json_({ ok: false, error: (err && err.message) || String(err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** ข้อมูลตั้งต้นของหน้าเว็บ */
function apiBoot() {
  var cfg = settings_();
  return {
    org: cfg.org,
    warnDays: cfg.warnDays,
    qrBase: cfg.qrBase,
    qrSize: cfg.qrSize,
    hasPin: !!cfg.pin,
    today: fmtDate_(new Date())
  };
}

/** รายการเครื่องมือทั้งหมด + สถานะ (ใช้ในหน้าภาพรวมและพิมพ์ป้าย) */
function apiList() {
  var cfg = settings_();
  return readTools_().map(function (t) { return publicTool_(t, cfg); });
}

/** ข้อมูลเครื่องมือ 1 ชิ้น (หน้าผลการสแกน) */
function apiGetTool(id) {
  var cfg = settings_();
  var t = findTool_(id);
  if (!t) return null;
  var out = publicTool_(t, cfg);
  out.history = readHistory_(t.tool_id).slice(0, 8);
  var since = t.last_cal instanceof Date ? t.last_cal : null;
  var usage = readUsage_(t.tool_id);
  out.usageSinceCal = usage.filter(function (u) { return !since || (u.ts && u.ts >= since); }).length;
  out.recentUsage = usage.slice(0, 5).map(function (u) {
    return { ts: u.ts ? Utilities.formatDate(u.ts, TZ, "yyyy-MM-dd'T'HH:mm") : '', user: u.user, aircraft: u.aircraft, job: u.job };
  });
  return out;
}

/** ช่างบันทึกการใช้เครื่องมือกับงาน (ไม่ต้องใช้ PIN) */
function apiLogUsage(p) {
  p = p || {};
  var t = findTool_(p.id);
  if (!t) throw new Error('ไม่พบเครื่องมือ ' + p.id);
  if (!str_(p.user)) throw new Error('กรุณากรอกชื่อผู้ใช้งาน');
  if (!str_(p.job) && !str_(p.aircraft)) throw new Error('กรุณากรอกทะเบียนอากาศยานหรือเลขที่ใบสั่งงาน');
  var st = computeStatus_(t, settings_());
  withLock_(function () {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.USAGE).appendRow([
      new Date(), t.tool_id, st.label, clean_(p.user, 80), clean_(str_(p.aircraft).toUpperCase(), 20), clean_(p.job, 40), clean_(p.note, 300)
    ]);
  });
  return { ok: true, status: st.label };
}

/** ผู้ดูแลบันทึกผลการสอบเทียบ (ต้องใช้ PIN) */
function apiRecordCalibration(pin, p) {
  checkPin_(pin);
  p = p || {};
  var t = findTool_(p.id);
  if (!t) throw new Error('ไม่พบเครื่องมือ ' + p.id);
  var d = parseIsoDate_(p.date);
  if (!d) throw new Error('กรุณาระบุวันที่สอบเทียบ');
  if (d > new Date()) throw new Error('วันที่สอบเทียบต้องไม่เป็นวันในอนาคต');
  if (RESULTS.indexOf(p.result) < 0) throw new Error('กรุณาเลือกผลการสอบเทียบ');
  var interval = Number(p.interval || t.cal_interval || 12);
  if (!(interval > 0 && interval <= 120)) throw new Error('รอบสอบเทียบต้องอยู่ระหว่าง 1–120 เดือน');
  var next = addMonths_(d, interval);
  if (str_(p.certLink) && !safeUrl_(p.certLink)) throw new Error('ลิงก์ใบรับรองต้องขึ้นต้นด้วย https://');

  withLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheetByName(SHEET.HIST).appendRow([
      new Date(), t.tool_id, d, p.result, clean_(p.certNo, 60), safeUrl_(p.certLink), clean_(p.lab, 120), interval, next, clean_(p.by, 80), clean_(p.note, 300)
    ]);
    var sh = ss.getSheetByName(SHEET.TOOLS);
    var r = t._row;
    setCell_(sh, r, 'cal_interval', interval);
    setCell_(sh, r, 'last_cal', d);
    setCell_(sh, r, 'cal_result', p.result);
    setCell_(sh, r, 'cert_no', clean_(p.certNo, 60));
    setCell_(sh, r, 'cert_link', safeUrl_(p.certLink));
    setCell_(sh, r, 'cal_lab', clean_(p.lab, 120));
    // ผ่าน → กลับมาใช้งานได้ (ล้างสถานะส่งสอบเทียบ) / ไม่ผ่าน → คงห้ามใช้ตามผล
    if (p.result === 'ผ่าน' && (t.state === 'ส่งสอบเทียบ')) setCell_(sh, r, 'state', '');
    // คืนสูตรวันครบกำหนด เผื่อเคยถูกพิมพ์ทับ
    sh.getRange(r, colOf_('next_due')).setFormula(rowFormulas_(r).next);
  });
  SpreadsheetApp.flush();
  return apiGetTool(t.tool_id);
}

/** ผู้ดูแลเปลี่ยนสถานะพิเศษ (ส่งสอบเทียบ / ส่งซ่อม / ชำรุด / NCR / ปกติ) */
function apiSetState(pin, id, state, note) {
  checkPin_(pin);
  var t = findTool_(id);
  if (!t) throw new Error('ไม่พบเครื่องมือ ' + id);
  state = str_(state);
  if (state && STATES.indexOf(state) < 0) throw new Error('สถานะไม่ถูกต้อง');
  withLock_(function () {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.TOOLS);
    setCell_(sh, t._row, 'state', state);
    if (str_(note)) {
      var stamp = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
      var prev = str_(t.remarks);
      setCell_(sh, t._row, 'remarks', (prev ? prev + '\n' : '') + stamp + ' ' + (state || 'กลับสู่สถานะปกติ') + ': ' + clean_(note, 300).replace(/^'/, ''));
    }
  });
  SpreadsheetApp.flush();
  return apiGetTool(t.tool_id);
}

/** ผู้ดูแลอัปโหลดรูปเครื่องมือจากมือถือ (ต้องใช้ PIN) → เก็บใน Google Drive และใส่ลิงก์ในชีต */
function apiUploadPhoto(pin, id, dataUrl) {
  checkPin_(pin);
  var t = findTool_(id);
  if (!t) throw new Error('ไม่พบเครื่องมือ ' + id);
  var m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=]+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('ไฟล์รูปไม่ถูกต้อง');
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 3 * 1024 * 1024) throw new Error('รูปใหญ่เกิน 3 MB');
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER);
  var ext = m[1] === 'jpeg' ? 'jpg' : m[1];
  var name = t.tool_id + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '.' + ext;
  var file = folder.createFile(Utilities.newBlob(bytes, 'image/' + m[1], name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  withLock_(function () {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.TOOLS);
    var hdr = sh.getRange(1, colOf_('photo'));
    if (!str_(hdr.getValue())) hdr.setValue('รูปภาพ (ลิงก์)').setFontWeight('bold').setBackground('#13202E').setFontColor('#FFFFFF');
    setCell_(sh, t._row, 'photo', 'https://drive.google.com/file/d/' + file.getId() + '/view');
  });
  SpreadsheetApp.flush();
  return apiGetTool(t.tool_id);
}

/** ฟิลด์ที่แก้ไขได้จากหน้าเว็บ (ไม่รวมข้อมูลสอบเทียบ ซึ่งบันทึกผ่าน "บันทึกผลการสอบเทียบ") */
var EDITABLE = ['group', 'description', 'model', 'part_no', 'serial_no', 'asset_no', 'location', 'cal_interval', 'cal_cost', 'remarks', 'photo'];
var ID_PREFIX = { 'กลุ่มซ่อมบำรุงอากาศยาน': 'M', 'กลุ่มอิเล็กทรอนิกส์การบิน': 'A' };

/** ผู้ดูแลเพิ่มเครื่องมือใหม่ (ไม่มี p.tool_id เดิม) หรือแก้ไขข้อมูลเครื่องมือ (p.mode = 'edit') */
function apiSaveTool(pin, p) {
  checkPin_(pin);
  p = p || {};
  var desc = str_(p.description), group = str_(p.group);
  if (!desc) throw new Error('กรุณากรอกชื่อรายการ');
  if (!group) throw new Error('กรุณาเลือกกลุ่มงาน');
  var interval = str_(p.cal_interval) === '' ? 12 : Number(p.cal_interval);
  if (!(interval > 0 && interval <= 120)) throw new Error('รอบสอบเทียบต้องอยู่ระหว่าง 1–120 เดือน');
  var cost = str_(p.cal_cost) === '' ? '' : Number(String(p.cal_cost).replace(/,/g, ''));
  if (cost !== '' && !(cost >= 0)) throw new Error('ราคาสอบเทียบไม่ถูกต้อง');
  if (str_(p.photo) && !safeUrl_(p.photo)) throw new Error('ลิงก์รูปต้องขึ้นต้นด้วย https://');
  var val = {
    group: clean_(group, 80), description: clean_(desc, 200), model: clean_(p.model, 60), part_no: clean_(p.part_no, 60),
    serial_no: clean_(p.serial_no, 60), asset_no: clean_(p.asset_no, 80), location: clean_(p.location, 80),
    cal_interval: interval, cal_cost: cost, remarks: clean_(p.remarks, 500), photo: safeUrl_(p.photo)
  };
  var all = readTools_();
  var dupSn = val.serial_no && all.filter(function (t) {
    return str_(t.serial_no).toUpperCase() === val.serial_no.toUpperCase() && t.tool_id !== str_(p.tool_id).toUpperCase();
  });
  var id;
  withLock_(function () {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.TOOLS);
    if (p.mode === 'edit') {
      var t = findTool_(p.tool_id);
      if (!t || t.tool_id !== str_(p.tool_id).toUpperCase()) throw new Error('ไม่พบเครื่องมือ ' + p.tool_id);
      id = t.tool_id;
      EDITABLE.forEach(function (k) { setCell_(sh, t._row, k, val[k]); });
      sh.getRange(t._row, colOf_('serial_no')).setNumberFormat('@').setValue(val.serial_no);
    } else {
      id = str_(p.tool_id).toUpperCase();
      if (id) {
        if (!/^[A-Z0-9][A-Z0-9\-]{1,19}$/.test(id)) throw new Error('Tool ID ใช้ได้เฉพาะ A–Z, 0–9 และ - (2–20 ตัว)');
        if (all.some(function (t) { return t.tool_id === id; })) throw new Error('Tool ID ' + id + ' มีอยู่แล้ว');
      } else {
        id = nextToolId_(all, ID_PREFIX[group] || 'T');
      }
      var r = sh.getLastRow() + 1;
      var row = TOOL_COLS.map(function (c) { return c[0] === 'tool_id' ? id : (val.hasOwnProperty(c[0]) ? val[c[0]] : ''); });
      sh.getRange(r, colOf_('tool_id')).setNumberFormat('@');
      sh.getRange(r, colOf_('serial_no')).setNumberFormat('@');
      sh.getRange(r, 1, 1, TOOL_COLS.length).setValues([row]);
      var f = rowFormulas_(r);
      sh.getRange(r, colOf_('next_due')).setFormula(f.next);
      sh.getRange(r, colOf_('status')).setFormula(f.status);
      sh.getRange(r, colOf_('last_cal'), 1, 2).setNumberFormat('d mmm yyyy');
    }
  });
  SpreadsheetApp.flush();
  var out = apiGetTool(id);
  if (dupSn && dupSn.length) out.warning = 'S/N นี้ซ้ำกับ ' + dupSn.map(function (t) { return t.tool_id; }).join(', ');
  return out;
}

/** ผู้ดูแลนำเครื่องมือออกจากทะเบียน (จำหน่าย / ไม่ใช้แล้ว) — ย้ายแถวไปแท็บ "เครื่องมือที่จำหน่ายแล้ว" ไม่ลบประวัติ */
function apiRetireTool(pin, id, reason) {
  checkPin_(pin);
  reason = clean_(reason, 300).replace(/^'/, '');
  if (!reason) throw new Error('กรุณาระบุเหตุผลที่นำออกจากทะเบียน');
  var t = findTool_(id);
  if (!t || t.tool_id !== str_(id).toUpperCase()) throw new Error('ไม่พบเครื่องมือ ' + id);
  withLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET.TOOLS);
    var arc = ensureSheet_(ss, SHEET.RETIRED, ['วันที่นำออก', 'เหตุผล'].concat(TOOL_COLS.map(function (c) { return c[1]; })));
    var values = sh.getRange(t._row, 1, 1, TOOL_COLS.length).getValues()[0];
    arc.appendRow([new Date(), reason].concat(values));
    sh.deleteRow(t._row);
  });
  return { ok: true, tool_id: t.tool_id };
}

function nextToolId_(all, prefix) {
  var max = 0, re = new RegExp('^' + prefix + '-(\\d+)$');
  all.forEach(function (t) { var m = re.exec(t.tool_id); if (m) max = Math.max(max, Number(m[1])); });
  return prefix + '-' + ('00' + (max + 1)).slice(-3);
}

function apiCheckPin(pin) {
  checkPin_(pin);
  return true;
}

// ------------------------------------------------------------------ แจ้งเตือนทางอีเมล

/** เมนู: ส่งรายงานทันที */
function sendReportNow() {
  requireSheetUi_();
  SpreadsheetApp.getUi().alert(buildAndSendReport_());
}

/** ตัวจัดการ trigger รายสัปดาห์ — รับเฉพาะการเรียกจาก trigger ของโปรเจกต์นี้ */
function weeklyReport(e) {
  var uid = e && e.triggerUid;
  var ok = uid && ScriptApp.getProjectTriggers().some(function (tr) { return tr.getUniqueId() === String(uid); });
  if (!ok) throw new Error('not allowed');
  buildAndSendReport_();
}

function buildAndSendReport_() {
  var cfg = settings_();
  var all = readTools_().map(function (t) { return publicTool_(t, cfg); });
  var pick = function (keys) { return all.filter(function (t) { return keys.indexOf(t.statusKey) >= 0; }); };
  var bad = pick(['expired', 'bad']), warn = pick(['warn']), cal = pick(['cal']), unknown = pick(['unknown']);
  var url = cfg.qrBase || '';

  var row = function (t) {
    return '<tr><td style="padding:4px 8px;font-family:monospace">' + esc_(t.tool_id) + '</td><td style="padding:4px 8px">' + esc_(t.description) +
      '</td><td style="padding:4px 8px">' + esc_(t.location) + '</td><td style="padding:4px 8px">' + (t.next_due || '-') +
      '</td><td style="padding:4px 8px">' + esc_(t.statusLabel) + (t.daysLeft !== null ? ' (' + t.daysLeft + ' วัน)' : '') + '</td></tr>';
  };
  var block = function (title, list, color) {
    if (!list.length) return '';
    return '<h3 style="color:' + color + ';margin:18px 0 6px">' + title + ' (' + list.length + ')</h3>' +
      '<table style="border-collapse:collapse;font-size:13px" border="1" bordercolor="#D9DEE3"><tr style="background:#E8EDF2">' +
      '<th>Tool ID</th><th>รายการ</th><th>สถานที่</th><th>ครบกำหนด</th><th>สถานะ</th></tr>' + list.map(row).join('') + '</table>';
  };
  var html = '<div style="font-family:sans-serif"><h2>รายงานสถานะการสอบเทียบเครื่องมือวัด</h2>' +
    '<p>ข้อมูล ณ ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + ' · เครื่องมือทั้งหมด ' + all.length + ' รายการ</p>' +
    block('ห้ามใช้งาน / หมดอายุ', bad, '#A8261B') +
    block('ใกล้ครบกำหนด (ภายใน ' + cfg.warnDays + ' วัน)', warn, '#8A5300') +
    block('อยู่ระหว่างส่งสอบเทียบ', cal, '#1C4E80') +
    (unknown.length ? '<p style="margin-top:16px">ยังไม่มีข้อมูลวันที่สอบเทียบ ' + unknown.length + ' รายการ</p>' : '') +
    (url ? '<p><a href="' + url + '">เปิดระบบติดตามสถานะ</a></p>' : '') + '</div>';

  if (!bad.length && !warn.length && !cal.length && !unknown.length) return 'ไม่มีรายการต้องแจ้งเตือน';
  var to = cfg.email || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: to,
    subject: '[Tool Calibration] ห้ามใช้ ' + bad.length + ' · ใกล้ครบกำหนด ' + warn.length + ' รายการ',
    htmlBody: html
  });
  return 'ส่งรายงานถึง ' + to;
}

function installWeeklyTrigger() {
  requireSheetUi_();
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'weeklyReport' || tr.getHandlerFunction() === 'sendDueReport') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('weeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).inTimezone(TZ).create();
  try { SpreadsheetApp.getUi().alert('ตั้งเวลาส่งรายงานทุกวันจันทร์ ประมาณ 08:00 น. เรียบร้อย'); } catch (e) {}
}

// ------------------------------------------------------------------ ภายใน

function computeStatus_(t, cfg) {
  var state = str_(t.state);
  var next = t.next_due instanceof Date ? t.next_due : null;
  if (!next && t.last_cal instanceof Date && Number(t.cal_interval) > 0) next = addMonths_(t.last_cal, Number(t.cal_interval));
  var days = next ? dayDiff_(today_(), next) : null;
  var res = { key: 'ok', label: 'พร้อมใช้งาน', usable: true, next: next, days: days };
  if (state === 'ไม่ต้องสอบเทียบ (NCR)') return set_(res, 'ncr', 'ไม่ต้องสอบเทียบ (NCR)', true);
  if (state === 'ส่งสอบเทียบ') return set_(res, 'cal', 'อยู่ระหว่างส่งสอบเทียบ', false);
  if (state === 'ส่งซ่อม') return set_(res, 'bad', 'ห้ามใช้งาน — ส่งซ่อม', false);
  if (state === 'ชำรุด') return set_(res, 'bad', 'ห้ามใช้งาน — ชำรุด', false);
  if (t.cal_result === 'ไม่ผ่าน' || t.cal_result === 'เกินเกณฑ์ (OOT)') return set_(res, 'bad', 'ห้ามใช้งาน — ผลสอบเทียบ' + t.cal_result, false);
  if (!next) return set_(res, 'unknown', 'ยังไม่มีข้อมูลการสอบเทียบ', false);
  if (days < 0) return set_(res, 'expired', 'ห้ามใช้งาน — หมดอายุการสอบเทียบ', false);
  if (days <= cfg.warnDays) return set_(res, 'warn', 'ใกล้ครบกำหนดสอบเทียบ', true);
  return res;
}

function set_(o, k, l, u) { o.key = k; o.label = l; o.usable = u; return o; }

function publicTool_(t, cfg) {
  var st = computeStatus_(t, cfg);
  return {
    tool_id: str_(t.tool_id), group: str_(t.group), description: str_(t.description), model: str_(t.model),
    part_no: str_(t.part_no), serial_no: str_(t.serial_no), asset_no: str_(t.asset_no), location: str_(t.location),
    cal_interval: t.cal_interval === '' ? '' : Number(t.cal_interval),
    last_cal: t.last_cal instanceof Date ? fmtDate_(t.last_cal) : '',
    next_due: st.next ? fmtDate_(st.next) : '',
    cal_result: str_(t.cal_result), cert_no: str_(t.cert_no), cert_link: safeUrl_(t.cert_link), cal_lab: str_(t.cal_lab),
    state: str_(t.state), legacy_plan: str_(t.legacy_plan), legacy_status: str_(t.legacy_status), remarks: str_(t.remarks),
    photo: photoUrl_(t.photo), photo_raw: safeUrl_(t.photo),
    statusKey: st.key, statusLabel: st.label, usable: st.usable, daysLeft: st.days
  };
}

function readTools_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.TOOLS);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, TOOL_COLS.length).getValues();
  var out = [];
  vals.forEach(function (row, i) {
    if (!str_(row[0])) return;
    var o = { _row: i + 2 };
    TOOL_COLS.forEach(function (c, j) { o[c[0]] = row[j]; });
    o.tool_id = str_(o.tool_id).toUpperCase();
    o.last_cal = fixBE_(o.last_cal);
    o.next_due = fixBE_(o.next_due);
    out.push(o);
  });
  return out;
}

function findTool_(id) {
  id = str_(id).toUpperCase();
  if (!id) return null;
  var all = readTools_();
  for (var i = 0; i < all.length; i++) if (all[i].tool_id === id) return all[i];
  // เผื่อสแกน/พิมพ์ด้วย S/N
  var bySn = all.filter(function (t) { return str_(t.serial_no).toUpperCase() === id; });
  return bySn.length === 1 ? bySn[0] : null;
}

function readHistory_(id) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.HIST);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HIST_COLS.length).getValues()
    .filter(function (r) { return str_(r[1]).toUpperCase() === id; })
    .map(function (r) {
      return {
        cal_date: r[2] instanceof Date ? fmtDate_(fixBE_(r[2])) : str_(r[2]), result: str_(r[3]), cert_no: str_(r[4]),
        cert_link: safeUrl_(r[5]), lab: str_(r[6]), by: str_(r[9])
      };
    })
    .sort(function (a, b) { return a.cal_date < b.cal_date ? 1 : -1; });
}

function readUsage_(id) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.USAGE);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, USAGE_COLS.length).getValues()
    .filter(function (r) { return str_(r[1]).toUpperCase() === id; })
    .map(function (r) { return { ts: r[0] instanceof Date ? r[0] : null, user: str_(r[3]), aircraft: str_(r[4]), job: str_(r[5]) }; })
    .reverse();
}

function settings_() {
  var cfg = { org: '', warnDays: 30, pin: '', email: '', qrBase: '', qrSize: 14 };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return cfg;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var k = str_(r[0]), v = r[1];
    if (k === 'ชื่อหน่วยงาน') cfg.org = str_(v);
    else if (k === 'แจ้งเตือนล่วงหน้า (วัน)') cfg.warnDays = Number(v) > 0 ? Number(v) : 30;
    else if (k === 'รหัสผู้ดูแล (PIN)') cfg.pin = str_(v);
    else if (k === 'อีเมลรับแจ้งเตือน') cfg.email = str_(v);
    else if (k === 'ลิงก์หน้าเว็บ (สำหรับ QR Code)' || k === 'ลิงก์สำหรับ QR Code') cfg.qrBase = str_(v);
    else if (k === 'ขนาด QR บนป้าย (มม.)') cfg.qrSize = Number(v) >= 8 ? Number(v) : 14;
  });
  return cfg;
}

function checkPin_(pin) {
  var cfg = settings_();
  if (!cfg.pin || cfg.pin.length < 4) throw new Error('ยังไม่ได้ตั้งรหัสผู้ดูแล (PIN) ในแท็บ "ตั้งค่า" ของฐานข้อมูล');
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pin_fail') || 0);
  if (fails >= 5) throw new Error('ใส่รหัสผิดหลายครั้ง ระบบล็อกการยืนยันไว้ 15 นาที');
  if (str_(pin) !== cfg.pin) {
    cache.put('pin_fail', String(fails + 1), 900);
    Utilities.sleep(1200);
    throw new Error('รหัสผู้ดูแลไม่ถูกต้อง');
  }
}

/** ฟังก์ชันจัดการระบบใช้ได้จากเมนูในชีตเท่านั้น (กันการเรียกผ่านเว็บแอป) */
function requireSheetUi_() {
  try { SpreadsheetApp.getUi(); } catch (e) { throw new Error('ใช้ได้จากเมนูในชีตเท่านั้น'); }
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function colOf_(key) {
  for (var i = 0; i < TOOL_COLS.length; i++) if (TOOL_COLS[i][0] === key) return i + 1;
  throw new Error('unknown column ' + key);
}
function colLetter_(key) {
  var n = colOf_(key), s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function setCell_(sh, row, key, val) { sh.getRange(row, colOf_(key)).setValue(val); }

function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
/** ข้อความจากผู้ใช้ที่จะเขียนลงชีต: ตัดความยาว และกันไม่ให้กลายเป็นสูตร */
function clean_(v, max) {
  var s = str_(v).slice(0, max || 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function esc_(s) { return str_(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function safeUrl_(u) { u = str_(u); return /^https?:\/\//i.test(u) ? u : ''; }
/** ลิงก์รูป: ไฟล์ใน Google Drive → ลิงก์ภาพย่อที่แสดงในหน้าเว็บได้ / ลิงก์รูปทั่วไปใช้ตามเดิม */
function photoUrl_(u) {
  u = safeUrl_(u);
  var m = /drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=\w+&)?id=)([\w-]{20,})/.exec(u);
  return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w800' : u;
}

function today_() {
  var s = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd').split('-');
  return new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]));
}
function fmtDate_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
/** ปี-เดือน-วัน ตามเวลาประเทศไทย (ไม่ขึ้นกับโซนเวลาของโปรเจกต์) */
function ymd_(d) { return fmtDate_(d).split('-').map(Number); }
function dayDiff_(a, b) {
  var x = ymd_(a), y = ymd_(b);
  return Math.round((Date.UTC(y[0], y[1] - 1, y[2]) - Date.UTC(x[0], x[1] - 1, x[2])) / 86400000);
}
function addMonths_(d, m) {
  var p = ymd_(d), y = p[0], mo = p[1] - 1 + m, day = p[2];
  var last = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
  return new Date(y, mo, Math.min(day, last));
}
function parseIsoDate_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str_(s));
  if (!m) return null;
  var y = Number(m[1]); if (y > 2400) y -= 543;
  return new Date(y, Number(m[2]) - 1, Number(m[3]));
}
/** กันกรณีกรอกปี พ.ศ. ลงช่องวันที่ (เช่น 2569) → แปลงเป็น ค.ศ. */
function fixBE_(d) {
  if (!(d instanceof Date)) return d;
  if (d.getFullYear() > 2400) return new Date(d.getFullYear() - 543, d.getMonth(), d.getDate());
  return d;
}

