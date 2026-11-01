/**
 * ระบบรายงาน Pet Shop R9 — Google Apps Script
 * ------------------------------------------------------------
 * ชีตที่ใช้ (สร้างอัตโนมัติจากเมนู "ตั้งค่าเริ่มต้น"):
 *   ตั้งค่า        : ชื่อรายงาน, ฟอนต์, รหัสผ่านผู้ดูแล ฯลฯ
 *   พนักงาน        : ชื่อ-สกุล | ชื่อเล่น | รหัส(#) | ใช้งาน
 *   Target         : ปี | รหัส | ชื่อเล่น | Target/วัน | ม.ค.–ธ.ค. | รวมปี
 *   DATA_นัดหมาย   : ข้อมูลจากไฟล์ TLP_01_01
 *   POS_yyyy-MM    : ข้อมูลจากไฟล์ TLPH-DATA (POS) แยกเดือน
 *   IMED_yyyy-MM   : ข้อมูลจากไฟล์ TLP_05_15 (ImedX) เฉพาะสินค้า 2-, A01–A17
 *   MANUAL         : ยอดโทร/LINE ที่พนักงานกรอกเอง (ต่อวัน)
 *   LOG_นำเข้า     : ประวัติการนำเข้าไฟล์
 */

const TZ = 'Asia/Bangkok';
const SH = {
  SET: 'ตั้งค่า', STAFF: 'พนักงาน', TARGET: 'Target', APPT: 'DATA_นัดหมาย',
  MANUAL: 'MANUAL', LOG: 'LOG_นำเข้า', STOCK: 'STOCK_ล่าสุด', CHECK: 'CHECK_Sameday', CLAIM: 'เคลมสินค้า', GRABLIST: 'GRAB_รายการ'
};
const SET_KEYS = {
  title1: 'ชื่อรายงานนัดหมาย',
  title2: 'ชื่อรายงานประจำวัน',
  title3: 'ชื่อรายงานรายสัปดาห์',
  fontIds: 'ฟอนต์ (Drive File ID)',
  fontName: 'ชื่อฟอนต์',
  prepOnly: 'เตรียมสินค้า: แสดงเฉพาะนัดที่มี #รหัส',
  pinto: 'เคสปิ่นโต: นับเฉพาะนัดที่ Note มีคำว่า "ปิ่นโต"',
  vets: 'คำนำหน้าผู้สั่งที่เป็นหมอ (คั่นด้วย ,)',
  pin: 'รหัสผ่านผู้ดูแล',
  grabBuffer: 'Grab: หักสต็อกกันไว้ (ชิ้น)',
  shifts: 'ตารางเวร: ตัวเลือกกะ (คั่นด้วย ,)',
  qrUrl: 'ลิงก์หน้าเว็บ (QR)'
};
const SPECIAL_TARGETS = [
  { code: 'TEAM',  label: 'นัดหมายรวมทีม (ว่าง = รวมจากรายคน)' },
  { code: 'VISIT', label: 'VISIT (ลูกค้าไม่ซ้ำชื่อ POS+ImedX)' },
  { code: 'BILL',  label: 'BILL (ยอดเฉลี่ยต่อบิล บาท)' }
];
const APPT_HEADERS = ['วัน-เวลาที่สร้างนัด','วันที่นัด','เวลาที่นัด','HN','ชื่อ-สกุล','ชื่อสัตว์เลี้ยง','Note','Follow','ทำนัดจาก','Member','Doctor'];
const POS_HEADERS  = ['วันที่','เวลา','เลขที่เอกสาร','ลูกค้า','พนักงานขาย','รหัสสินค้า','ชื่อสินค้า','จำนวน','สุทธิ','สถานะ','ผู้ให้บริการ','เลขที่การจัดส่ง'];
const IMED_HEADERS = ['วันที่','เวลา','VN/AN','HN','ชื่อคนไข้','ชื่อเจ้าของ','ผู้สั่ง','รหัส','รายการ','จำนวน','มูลค่ารวม'];
const MONTHS_TH = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
const MONTHS_SHORT = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
const DEFAULT_SHIFTS = '07.30-16.30,08.00-17.00,09.00-18.00,10.00-19.00,11.00-20.00,13.00-22.00,15.00-24.00,OFF,ลา,ชั้น 2,ลงเวร R8,ประชุม,MOD/13.00,-';
const PRODUCT_RE = /^(2-|A(0[1-9]|1[0-7])-)/i;   // สินค้ากลุ่ม 2-, A01 ถึง A17

/* =========================================================
 * เว็บแอป + เมนู
 * ========================================================= */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Page');
  t.startTab = (e && e.parameter && e.parameter.page) || 'upload';
  t.cabId = (e && e.parameter && e.parameter.cab) || '';
  return t.evaluate()
    .setTitle('รายงาน Pet Shop')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📅 ระบบรายงาน')
    .addItem('1) ตั้งค่าเริ่มต้น (สร้างชีต)', 'setupSheets')
    .addItem('2) เปิดหน้าโปรแกรม', 'openAppDialog')
    .addItem('ลิงก์หน้าเว็บสำหรับพนักงาน', 'showWebLink')
    .addItem('ล้างข้อมูลรหัสฝั่งโรงพยาบาล', 'cleanupNonProduct')
    .addItem('ดูพื้นที่ที่ใช้ไป', 'showUsage')
    .addToUi();
}

function openAppDialog() {
  const t = HtmlService.createTemplateFromFile('Page');
  t.startTab = 'upload';
  t.cabId = '';
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(1000).setHeight(720), 'ระบบรายงาน Pet Shop');
}

function showWebLink() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('ยังไม่ได้ Deploy เป็นเว็บแอป\n\nไปที่ ส่วนขยาย > Apps Script > ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เลือกประเภท "เว็บแอป"');
    return;
  }
  ui.showModalDialog(HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:14px">ลิงก์สำหรับพนักงาน:<br><br>' +
    '<a href="' + url + '" target="_blank">' + url + '</a><br><br>' +
    'หน้าเฉพาะ: ใส่ <b>?page=appt</b> / <b>daily</b> / <b>weekly</b> / <b>upload</b> / <b>staff</b> / <b>target</b> ต่อท้ายลิงก์</div>'
  ).setWidth(560).setHeight(200), 'ลิงก์หน้าเว็บ');
}

/* =========================================================
 * สร้างชีต
 * ========================================================= */
function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  let s = ss.getSheetByName(SH.SET);
  if (!s) {
    s = ss.insertSheet(SH.SET);
    const rows = [
      ['หัวข้อ', 'ค่า', 'คำอธิบาย'],
      [SET_KEYS.title1, 'รายงานจำนวนนัดหมายราย คน แผนก Pet Shop R9', ''],
      [SET_KEYS.title2, 'สรุปยอดประจำวัน แผนก Pet Shop R9', ''],
      [SET_KEYS.title3, 'สรุปรายสัปดาห์ แผนก Pet Shop R9', ''],
      [SET_KEYS.fontIds, '', 'อัปโหลดไฟล์ฟอนต์ Thonglor (.ttf/.otf) ลง Google Drive แล้วใส่ File ID (ตัวปกติ, ตัวหนา คั่นด้วย ,)'],
      [SET_KEYS.fontName, 'Thonglor', ''],
      [SET_KEYS.prepOnly, false, 'ติ๊ก = แสดงเฉพาะนัดที่มี #รหัส (ค่าเริ่มต้น: แสดงทุกนัด นัดที่ไม่มี # จะแสดงว่า "ไม่ระบุ")'],
      [SET_KEYS.pinto, false, 'ติ๊กออก = นับทุกนัดในวันที่เลือก'],
      [SET_KEYS.vets, 'สพ.ญ.,น.สพ.,นสพ.,สพ.ช.', 'ใช้หาเคสที่หมอสั่งจ่ายสินค้า'],
      [SET_KEYS.pin, '1234', 'รหัสผ่านของ Username admin (หัวหน้างาน) — ว่าง = 1234'],
      [SET_KEYS.grabBuffer, 0, 'จำนวนที่หักออกจากสต็อกจริงก่อนส่งให้ Grab (กันขายเกิน) ใส่ 0 = ส่งตามจริง']
    ];
    s.getRange(1, 1, rows.length, 3).setValues(rows);
    s.getRange('B7:B8').insertCheckboxes();
    s.getRange('B7').setValue(false); s.getRange('B8').setValue(false);
    styleHeader_(s, 3);
    s.setColumnWidth(1, 330); s.setColumnWidth(2, 360); s.setColumnWidth(3, 520);
  }

  // เพิ่มหัวข้อตั้งค่าที่ยังไม่มี (กรณีอัปเดตโค้ดเวอร์ชันใหม่)
  const have = s.getRange(1, 1, Math.max(1, s.getLastRow()), 1).getValues().map(r => String(r[0]).trim());
  if (have.indexOf(SET_KEYS.qrUrl) < 0) s.appendRow([SET_KEYS.qrUrl, '', 'ว่าง = ใช้ลิงก์ Apps Script · ใส่ลิงก์ GitHub Pages (เช่น https://ชื่อ.github.io/petshop/) เพื่อให้ QR ตู้สั้นลงและไม่เปลี่ยน']);
  if (have.indexOf(SET_KEYS.shifts) < 0) s.appendRow([SET_KEYS.shifts, DEFAULT_SHIFTS, 'ตัวเลือกในช่องตารางเวร (พิมพ์ค่าอื่นเองได้ เช่น OFF แทน 18)']);
  if (have.indexOf(SET_KEYS.grabBuffer) < 0) s.appendRow([SET_KEYS.grabBuffer, 0, 'จำนวนที่หักออกจากสต็อกจริงก่อนส่งให้ Grab (กันขายเกิน) ใส่ 0 = ส่งตามจริง']);

  s = ss.getSheetByName(SH.STAFF);
  if (!s) {
    s = ss.insertSheet(SH.STAFF);
    s.getRange(1, 1, 2, 4).setValues([
      ['ชื่อ-สกุล', 'ชื่อเล่น', 'รหัส (#)', 'ใช้งาน'],
      ['นายอภิวัฒน์ ไสลสุข', 'อาร์ม', 'AM', true]
    ]);
    s.getRange('D2:D200').insertCheckboxes();
    s.getRange('C2:C200').setNumberFormat('@');
    styleHeader_(s, 4);
    s.setColumnWidth(1, 260); s.setColumnWidth(2, 120);
  }

  s = ss.getSheetByName(SH.TARGET);
  if (!s) {
    s = ss.insertSheet(SH.TARGET);
    s.getRange(1, 1, 1, 17).setValues([['ปี (ค.ศ.)', 'รหัส', 'ชื่อ', 'Target/วัน'].concat(MONTHS_SHORT).concat(['รวมทั้งปี'])]);
    styleHeader_(s, 17);
    s.setColumnWidth(3, 220);
  }

  ensureSheet_(SH.APPT, APPT_HEADERS);
  ensureSheet_(SH.MANUAL, ['วันที่', 'โทร/LINE จำนวนเคส', 'โทร/LINE ยอดเงิน', 'บันทึกเมื่อ']);
  ensureSheet_(SH.LOG, ['เวลา', 'ไฟล์', 'ประเภท', 'จำนวนรายการ', 'ช่วงวันที่', 'หมายเหตุ']);

  ['Sheet1', 'แผ่นงาน1', 'ชีต1'].forEach(n => {
    const x = ss.getSheetByName(n);
    if (x && ss.getSheets().length > 1 && x.getLastRow() === 0) ss.deleteSheet(x);
  });
  try { SpreadsheetApp.getActive().toast('สร้างชีตเรียบร้อย ✅', 'ระบบรายงาน', 5); } catch (e) {}
  Logger.log('สร้างชีตเรียบร้อย');
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(s, headers.length);
    // ตัดคอลัมน์/แถวว่างเริ่มต้น (ชีตใหม่มี 26 คอลัมน์ × 1000 แถว) เพื่อประหยัดโควตา 10 ล้านเซลล์
    try {
      if (s.getMaxColumns() > headers.length) s.deleteColumns(headers.length + 1, s.getMaxColumns() - headers.length);
      if (s.getMaxRows() > 2) s.deleteRows(3, s.getMaxRows() - 2);
    } catch (e) {}
  }
  return s;
}

function styleHeader_(s, n) {
  s.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff');
  s.setFrozenRows(1);
}

function sheet_(name) {
  let s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) { setupSheets(); s = SpreadsheetApp.getActive().getSheetByName(name); }
  return s;
}

/* =========================================================
 * ตั้งค่า / พนักงาน / Target
 * ========================================================= */
function getSettings_() {
  const v = sheet_(SH.SET).getRange('A2:B40').getValues();
  const m = {};
  v.forEach(r => { if (r[0]) m[String(r[0]).trim()] = r[1]; });
  return {
    title1: m[SET_KEYS.title1] || 'รายงานจำนวนนัดหมาย',
    title2: m[SET_KEYS.title2] || 'สรุปยอดประจำวัน',
    title3: m[SET_KEYS.title3] || 'สรุปรายสัปดาห์',
    fontIds: String(m[SET_KEYS.fontIds] || '').split(',').map(x => x.trim()).filter(String),
    fontName: String(m[SET_KEYS.fontName] || 'Thonglor').trim(),
    prepOnlyCoded: m[SET_KEYS.prepOnly] === true,
    pintoOnly: m[SET_KEYS.pinto] === true,
    vets: String(m[SET_KEYS.vets] || 'สพ.ญ.,น.สพ.').split(',').map(x => x.trim()).filter(String),
    pin: String(m[SET_KEYS.pin] || '').trim(),
    grabBuffer: Math.max(0, Number(m[SET_KEYS.grabBuffer]) || 0),
    shifts: String(m[SET_KEYS.shifts] || DEFAULT_SHIFTS).split(',').map(x => x.trim()).filter(String)
  };
}

function cleanCode_(c) { return String(c || '').replace(/#/g, '').trim().toUpperCase(); }

function getStaff_(includeInactive) {
  const s = sheet_(SH.STAFF);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, 4).getValues()
    .map(r => ({ name: String(r[0]).trim(), nick: String(r[1]).trim(), code: cleanCode_(r[2]), active: r[3] !== false }))
    .filter(x => x.name && x.code && (includeInactive || x.active));
}

/** คืนค่า { CODE: {day, months:[12]} } ของปีที่ระบุ */
function getTargets_(year) {
  const s = sheet_(SH.TARGET);
  const n = s.getLastRow() - 1;
  const out = {};
  if (n < 1) return out;
  s.getRange(2, 1, n, 16).getValues().forEach(r => {
    if (Number(r[0]) !== Number(year) || !r[1]) return;
    out[cleanCode_(r[1])] = { name: String(r[2] || ''), day: Number(r[3]) || 0, months: r.slice(4, 16).map(x => Number(x) || 0) };
  });
  return out;
}

function checkPin_() {
  if (!CURRENT_USER) throw new Error('กรุณาเข้าสู่ระบบ');
  if (CURRENT_USER.role !== 'admin') throw new Error('เฉพาะหัวหน้างาน');
}

const STAFF_HEADERS = ['ชื่อ-สกุล', 'ชื่อเล่น', 'รหัส (#)', 'ใช้งาน (นัดหมาย)', 'ชื่อเล่น EN', 'รหัสพนักงาน', 'เบอร์โทร', 'อยู่ในตารางเวร', 'ส่งรายงาน HR'];

/** พนักงานทั้งหมด (ทุกคอลัมน์) */
function getAllStaffRows_() {
  const s = sheet_(SH.STAFF);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  const w = Math.max(4, Math.min(9, s.getLastColumn()));
  return s.getRange(2, 1, n, 9).getValues().map(r => ({
    name: String(r[0]).trim(), nick: String(r[1]).trim(), code: cleanCode_(r[2]), active: r[3] !== false,
    en: String(r[4] || '').trim(), emp: String(r[5] || '').trim(), phone: String(r[6] || '').trim(),
    roster: w >= 8 ? r[7] === true : false, hr: w >= 9 ? r[8] === true : false
  })).filter(x => x.name);
}


/** ข้อมูลสำหรับหน้า พนักงาน / Target (เฉพาะหัวหน้างาน) */
function getAdminData(year, pin) {
  checkPin_(pin);
  year = Number(year) || Number(Utilities.formatDate(new Date(), TZ, 'yyyy'));
  const all = getAllStaffRows_();
  const t = getTargets_(year);
  const rows = [];
  all.filter(s => s.code).forEach(s => rows.push({ code: s.code, name: s.nick ? s.name + ' (' + s.nick + ')' : s.name, kind: 'staff', day: (t[s.code] || {}).day || '', months: (t[s.code] || {}).months || [] }));
  SPECIAL_TARGETS.forEach(x => rows.push({ code: x.code, name: x.label, kind: 'special', day: (t[x.code] || {}).day || '', months: (t[x.code] || {}).months || [] }));
  return { year: year, staff: all, targets: rows };
}

function saveStaff(list, pin) {
  bumpVersion_();
  checkPin_(pin);
  const codes = {}, names = {};
  const rows = (list || []).map(x => [String(x.name || '').trim(), String(x.nick || '').trim(), cleanCode_(x.code), x.active !== false,
    String(x.en || '').trim(), String(x.emp || '').trim(), String(x.phone || '').trim(), x.roster === true, x.hr === true])
    .filter(r => r[0] || r[2]);
  rows.forEach(r => {
    if (!r[0]) throw new Error('กรุณากรอกชื่อ-สกุลให้ครบ');
    if (names[r[0]]) throw new Error('ชื่อ ' + r[0] + ' ซ้ำกัน');
    names[r[0]] = 1;
    if (r[2]) { if (codes[r[2]]) throw new Error('รหัส #' + r[2] + ' ซ้ำกัน'); codes[r[2]] = 1; }
  });
  const s = sheet_(SH.STAFF);
  s.getRange(1, 1, 1, 9).setValues([STAFF_HEADERS]);
  styleHeader_(s, 9);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 9).clearContent().removeCheckboxes();
  if (rows.length) {
    s.getRange(2, 3, rows.length, 1).setNumberFormat('@');
    s.getRange(2, 6, rows.length, 2).setNumberFormat('@');
    s.getRange(2, 1, rows.length, 9).setValues(rows);
    s.getRange(2, 4, rows.length, 1).insertCheckboxes();
    s.getRange(2, 8, rows.length, 2).insertCheckboxes();
  }
  return 'บันทึกพนักงาน ' + rows.length + ' คนแล้ว';
}

function saveTargets(year, rows, pin) {
  bumpVersion_();
  checkPin_(pin);
  year = Number(year);
  const s = sheet_(SH.TARGET);
  const n = s.getLastRow() - 1;
  const keep = n > 0 ? s.getRange(2, 1, n, 16).getValues().filter(r => r[0] !== '' && Number(r[0]) !== year) : [];
  const add = (rows || []).map(x => {
    const m = [];
    for (let i = 0; i < 12; i++) m.push(x.months && x.months[i] !== '' && x.months[i] != null ? Number(x.months[i]) || 0 : '');
    return [year, cleanCode_(x.code), x.name || '', x.day === '' || x.day == null ? '' : Number(x.day) || 0].concat(m);
  }).filter(r => r[3] !== '' || r.slice(4).some(v => v !== ''));
  const all = keep.concat(add).sort((a, b) => a[0] - b[0]);
  if (n > 0) s.getRange(2, 1, n, 17).clearContent();
  if (all.length) {
    s.getRange(2, 1, all.length, 16).setValues(all);
    s.getRange(2, 17, all.length, 1).setFormulas(all.map((_, i) => ['=SUM(E' + (i + 2) + ':P' + (i + 2) + ')']));
  }
  return 'บันทึก Target ปี ' + year + ' แล้ว';
}

/* =========================================================
 * ฟอนต์ (ฝังจาก Google Drive)
 * ========================================================= */
function getFontCss() {
  const st = getSettings_();
  let css = '';
  st.fontIds.slice(0, 2).forEach((id, i) => {
    try {
      const f = DriveApp.getFileById(id);
      const name = f.getName().toLowerCase();
      const fmt = /\.otf$/.test(name) ? ['font/otf', 'opentype'] : /\.woff2$/.test(name) ? ['font/woff2', 'woff2'] : /\.woff$/.test(name) ? ['font/woff', 'woff'] : ['font/ttf', 'truetype'];
      css += "@font-face{font-family:'" + st.fontName + "';font-weight:" + (i ? 700 : 400) +
        ";src:url(data:" + fmt[0] + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes()) + ") format('" + fmt[1] + "');}";
    } catch (e) { /* ไม่มีสิทธิ์อ่านไฟล์ฟอนต์ */ }
  });
  return { css: css, name: st.fontName };
}

/* =========================================================
 * ตัวช่วยวันที่ / ตัวเลข
 * ========================================================= */
function pad_(x) { x = String(x); return x.length < 2 ? '0' + x : x; }
function dayKey_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function keyToDate_(k) { return Utilities.parseDate(k, TZ, 'yyyy-MM-dd'); }
function addDays_(k, n) { return dayKey_(new Date(keyToDate_(k).getTime() + n * 86400000)); }
function thaiShort_(k) { const p = k.split('-'); return Number(p[2]) + '/' + Number(p[1]) + '/' + p[0]; }
function num_(v) { const n = Number(String(v == null ? '' : v).replace(/,/g, '').trim()); return isNaN(n) ? 0 : n; }
function fixYear_(y) { y = Number(y); if (y < 100) y = y >= 60 ? 2500 + y - 543 : 2000 + y; if (y > 2400) y -= 543; return y; }
function todayKey_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function normName_(s) { return String(s || '').replace(/^\s*คุณ\s*/, '').replace(/\s+/g, ' ').trim().toLowerCase(); }

/** "08/08/2026 18:55" หรือ Date → Date */
function parseDT_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2})[:.](\d{2}))?/);
  if (!m) return null;
  return Utilities.parseDate(pad_(m[1]) + '/' + pad_(m[2]) + '/' + fixYear_(m[3]) + ' ' + pad_(m[4] || '0') + ':' + (m[5] || '00'), TZ, 'dd/MM/yyyy HH:mm');
}

/** "1 พ.ย. 2026 17:13:11" → {key, time} */
function parseThaiShortDT_(v) {
  if (v instanceof Date) return { key: dayKey_(v), time: Utilities.formatDate(v, TZ, 'HH:mm') };
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\s*(\S+?)\s*(\d{2,4})\s+(\d{1,2}):(\d{2})/);
  if (m) {
    const mi = MONTHS_SHORT.indexOf(m[2]) >= 0 ? MONTHS_SHORT.indexOf(m[2]) : MONTHS_TH.indexOf(m[2]);
    if (mi >= 0) return { key: fixYear_(m[3]) + '-' + pad_(mi + 1) + '-' + pad_(m[1]), time: pad_(m[4]) + ':' + m[5] };
  }
  const d = parseDT_(s);
  return d ? { key: dayKey_(d), time: Utilities.formatDate(d, TZ, 'HH:mm') } : null;
}

/** "ตั้งแต่วันที่ 1 มกราคม 2569 ถึงวันที่ 31 ธันวาคม 2569" → {start, end} (yyyy-MM-dd) */
function parseThaiRange_(text) {
  const re = new RegExp('(\\d{1,2})\\s*(' + MONTHS_TH.concat(MONTHS_SHORT).map(x => x.replace(/\./g, '\\.')).join('|') + ')\\s*(\\d{2,4})', 'g');
  const out = [];
  let m;
  while ((m = re.exec(text))) {
    let mi = MONTHS_TH.indexOf(m[2]);
    if (mi < 0) mi = MONTHS_SHORT.indexOf(m[2]);
    out.push(fixYear_(m[3]) + '-' + pad_(mi + 1) + '-' + pad_(m[1]));
  }
  return out.length >= 2 ? { start: out[0], end: out[1] } : null;
}

/* =========================================================
 * นำเข้าไฟล์
 * ========================================================= */
function detectType_(rows) {
  const head = rows.slice(0, 12).map(r => r.join('|')).join('\n');
  if (head.indexOf('วัน-เวลาที่สร้างนัด') >= 0) return 'APPT';
  if (head.indexOf('วันเวลาที่ขาย') >= 0 && head.indexOf('เลขที่เอกสาร') >= 0) return 'POS';
  if (head.indexOf('รายงานอัตราการจ่ายยา') >= 0 || head.indexOf('VN/AN') >= 0) return 'IMED';
  if (head.indexOf('รายงานเตรียมตรวจนับ') >= 0) return 'STOCK';
  if (head.indexOf('รายงานราคาขายสินค้า') >= 0) return 'MASTER';
  if (/^STOREID\|ITEMID/i.test(head) || head.indexOf('STOREID|ITEMID|ITEMNAME') >= 0) return 'GRAB';
  return '';
}

/** เรียกจากหน้าเว็บ: ส่งข้อความ CSV ทั้งไฟล์มา */
function importCsvText(text, fileName) {
  auth_();
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    text = String(text || '').replace(/^﻿/, '');
    const rows = Utilities.parseCsv(text);
    const type = detectType_(rows);
    let res;
    if (type === 'APPT') res = importAppt_(rows);
    else if (type === 'POS') res = importPos_(rows);
    else if (type === 'IMED') res = importImed_(rows);
    else if (type === 'STOCK') res = importStock_(rows);
    else if (type === 'GRAB') {
      return { ok: false, type: type, message: 'ไฟล์ Grab ให้วางที่แท็บ "Grab Stock" เพื่อเติมสต็อกและดาวน์โหลดไฟล์กลับ' };
    } else if (type === 'MASTER') res = importMaster_(rows);
    else throw new Error('ไม่รู้จักรูปแบบไฟล์นี้ (รองรับ TLP_01_01, TLPH-DATA, TLP_05_15, TLP_14_09, TLP_18_05, ไฟล์ Grab)');
    bumpVersion_();
    sheet_(SH.LOG).appendRow([new Date(), fileName || '', type, res.count, res.range || '', (res.note || '') + (CURRENT_USER ? ' · โดย ' + CURRENT_USER.name : '')]);
    res.ok = true; res.type = type;
    return res;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ไฟล์นัดหมาย TLP_01_01 ---------- */
function parseApptRows_(rows) {
  const recs = [], seen = {};
  let last = null;
  rows.forEach(r => {
    const a = String(r[0] || '').trim();
    if (/^\d+$/.test(a)) {
      const c = parseDT_(r[1]), ap = parseDT_(r[2]);
      if (c && ap) {
        last = {
          created: c, appt: ap, time: String(r[3] || '').trim(), hn: String(r[4] || '').trim(),
          name: String(r[5] || '').replace(/\s+/g, ' ').trim(), pet: String(r[6] || '').trim(),
          note: String(r[11] || ''), follow: String(r[9] || '').trim(), from: String(r[12] || '').trim(),
          member: String(r[13] || '').trim(), doctor: String(r[14] || '').trim()
        };
        const k = [c.getTime(), last.hn, ap.getTime(), last.time, last.pet].join('|');
        if (!seen[k]) { seen[k] = 1; recs.push(last); } else last = null;
        return;
      }
    }
    // บรรทัดต่อของ Note ที่ถูกตัดข้ามหน้า (ช่อง A, B, E ว่าง แต่ช่อง L มีข้อความ)
    if (!a && last && r[11] && !String(r[1] || '').trim() && !String(r[4] || '').trim()) last.note += String(r[11]);
  });
  recs.forEach(x => x.note = x.note.trim());
  return recs;
}

/** รวมข้อมูลนัดเก่า + ใหม่ แล้วตัดรายการซ้ำ (pure)
 *  ชั้น 1: นัดเดียวกันที่ถูก export ซ้ำ (วันเวลาสร้าง + HN + วันที่นัด + เวลานัด + สัตว์) → ใช้ข้อมูลจากไฟล์ใหม่
 *  ชั้น 2: นัดที่ถูกสร้างซ้ำ (HN + วันที่นัด + เวลานัด + สัตว์ + Note เหมือนกัน) → เก็บรายการที่สร้างก่อน */
function mergeAppt_(oldRecs, newRecs) {
  const norm = x => String(x || '').replace(/\s+/g, ' ').trim().toLowerCase();
  const minute = d => Math.floor(d.getTime() / 60000);
  const k1 = r => [minute(r.created), norm(r.hn), dayKey_(r.appt), norm(r.time), norm(r.pet)].join('|');
  const k2 = r => [norm(r.hn), dayKey_(r.appt), norm(r.time), norm(r.pet), norm(r.note)].join('|');
  const m1 = {};
  oldRecs.forEach(r => m1[k1(r)] = r);
  let updated = 0, added = 0;
  newRecs.forEach(r => { if (m1[k1(r)]) updated++; else added++; m1[k1(r)] = r; });
  const m2 = {};
  Object.keys(m1).map(k => m1[k]).sort((a, b) => a.created - b.created).forEach(r => { if (!m2[k2(r)]) m2[k2(r)] = r; });
  const all = Object.keys(m2).map(k => m2[k])
    .sort((a, b) => a.appt - b.appt || String(a.time).localeCompare(String(b.time)) || a.created - b.created);
  return { all: all, added: Math.max(0, all.length - oldRecs.length), updated: updated, removed: Object.keys(m1).length - all.length };
}

function importAppt_(rows) {
  const recs = parseApptRows_(rows);
  if (!recs.length) throw new Error('ไม่พบรายการนัดหมายในไฟล์');
  const ks = recs.map(r => dayKey_(r.appt)).sort();
  const m = mergeAppt_(readAppt_(), recs);
  const all = m.all;
  const s = sheet_(SH.APPT);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, APPT_HEADERS.length).clearContent();
  if (all.length) {
    s.getRange(2, 3, all.length, 9).setNumberFormat('@');
    s.getRange(2, 1, all.length, APPT_HEADERS.length).setValues(all.map(r => [r.created, r.appt, r.time, r.hn, r.name, r.pet, r.note, r.follow, r.from, r.member, r.doctor]));
    s.getRange(2, 1, all.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    s.getRange(2, 2, all.length, 1).setNumberFormat('dd/MM/yyyy');
  }
  const rg = thaiShort_(ks[0]) + ' - ' + thaiShort_(ks[ks.length - 1]);
  return {
    count: recs.length, range: rg, note: 'ใหม่ ' + m.added + ' / อัปเดต ' + m.updated + ' / ตัดซ้ำ ' + m.removed,
    message: 'นำเข้านัดหมาย ' + recs.length + ' รายการ (วันที่นัด ' + rg + ') → เพิ่มใหม่ ' + m.added + ', ซ้ำกับข้อมูลเดิม (อัปเดตแทน) ' + m.updated +
      ', ตัดรายการซ้ำ ' + m.removed + ' · สะสมในระบบทั้งหมด ' + all.length + ' รายการ'
  };
}

/* =========================================================
 * รายงาน 4: นัดหมายประจำเดือน (ตามวันที่นัด)
 * ========================================================= */
function getApptMonth(monthKey, toKey) {
  auth_();
  monthKey = monthKey || todayKey_().slice(0, 7);
  toKey = toKey || monthKey;
  if (toKey < monthKey) { const t = toKey; toKey = monthKey; monthKey = t; }
  const months = [];
  for (let k = monthKey; k <= toKey && months.length < 36;) {
    months.push(k);
    const y = Number(k.slice(0, 4)), m = Number(k.slice(5, 7));
    k = m === 12 ? (y + 1) + '-01' : y + '-' + pad_(m + 1);
  }
  const inRange = {};
  months.forEach(m => inRange[m] = 1);
  const staff = getStaff_();
  const allStaff = getStaff_(true);   // รวมคนที่ย้ายแผนก/ปิดใช้งาน เพื่อแสดงชื่อในรายการเก่า
  const tByYear = {};
  const targetsOf = code => months.map(mk => {
    const y = Number(mk.slice(0, 4));
    tByYear[y] = tByYear[y] || getTargets_(y);
    return ((tByYear[y][code] || { months: [] }).months[Number(mk.slice(5, 7)) - 1]) || 0;
  });
  const targetOf = code => targetsOf(code).reduce((a, b) => a + b, 0);
  const rows = readAppt_().filter(r => inRange[dayKey_(r.appt).slice(0, 7)]).map(r => {
    const j = matchStaff_(r.note, allStaff);
    const act = j >= 0 && allStaff[j].active;
    const tag = anyTag_(r.note);
    return {
      d: dayKey_(r.appt), time: r.time, created: Utilities.formatDate(r.created, TZ, 'd/M/yyyy HH:mm'),
      hn: r.hn, name: r.name, pet: r.pet, note: r.note, follow: r.follow, from: r.from, member: r.member,
      code: act ? allStaff[j].code : '', who: j >= 0 ? (allStaff[j].nick || allStaff[j].name) : (tag ? '#' + tag : 'ไม่ระบุ')
    };
  });
  const mName = mk => MONTHS_TH[Number(mk.slice(5, 7)) - 1] + ' ' + mk.slice(0, 4);
  return {
    title: 'รายงานนัดหมาย ' + (months.length > 1 ? mName(monthKey) + ' - ' + mName(toKey) : 'ประจำเดือน ' + mName(monthKey)),
    monthKey: monthKey, toKey: toKey, months: months, rows: rows, imports: lastImports_(),
    staff: staff.map(s => ({ code: s.code, name: s.nick ? s.name + ' (' + s.nick + ')' : s.name, nick: s.nick || s.name, target: targetOf(s.code), tm: targetsOf(s.code) })),
    monthLabels: months.map(mk => MONTHS_SHORT[Number(mk.slice(5, 7)) - 1] + ' ' + String(Number(mk.slice(0, 4)) + 543).slice(2))
  };
}

function readAppt_() {
  const s = sheet_(SH.APPT);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, APPT_HEADERS.length).getValues()
    .filter(r => r[0] instanceof Date && r[1] instanceof Date)
    .map(r => ({ created: r[0], appt: r[1], time: String(r[2]), hn: String(r[3]), name: String(r[4]), pet: String(r[5]), note: String(r[6]), follow: String(r[7]), from: String(r[8]), member: String(r[9]), doctor: String(r[10]) }));
}

/* ---------- ข้อมูลรายเดือน (POS / IMED) ---------- */
function monthSheetName_(prefix, key) { return prefix + '_' + key.slice(0, 7); }

/** เขียนแถวลงชีตรายเดือน โดยแทนที่ข้อมูลเดิมของวันที่ที่อยู่ในไฟล์ */
function writeMonthly_(prefix, headers, rows) {
  const byMonth = {};
  rows.forEach(r => (byMonth[r[0].slice(0, 7)] = byMonth[r[0].slice(0, 7)] || []).push(r));
  Object.keys(byMonth).forEach(mk => {
    const add = byMonth[mk];
    const days = {};
    add.forEach(r => days[r[0]] = 1);
    const s = ensureSheet_(prefix + '_' + mk, headers);
    const n = s.getLastRow() - 1;
    const old = n > 0 ? s.getRange(2, 1, n, headers.length).getDisplayValues().filter(r => r[0] && !days[r[0]]) : [];
    const all = old.concat(add).sort((a, b) => (a[0] + a[1]).localeCompare(b[0] + b[1]));
    if (n > 0) s.getRange(2, 1, n, headers.length).clearContent();
    s.getRange(2, 1, all.length, headers.length).setNumberFormat('@');
    s.getRange(2, 1, all.length, headers.length).setValues(all.map(r => r.map(String)));
  });
}

function readMonthly_(prefix, monthKey, width) {
  const s = SpreadsheetApp.getActive().getSheetByName(prefix + '_' + monthKey);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, width).getDisplayValues();
}

/** อ่านข้อมูลช่วงวันที่ [start, end] จากชีตรายเดือน */
function readRange_(prefix, width, start, end) {
  const months = {};
  for (let k = start; k <= end; k = addDays_(k, 1)) months[k.slice(0, 7)] = 1;
  let out = [];
  Object.keys(months).forEach(mk => {
    out = out.concat(readMonthly_(prefix, mk, width).filter(r => r[0] >= start && r[0] <= end));
  });
  return out;
}

/* ---------- POS TLPH-DATA ---------- */
function importPos_(rows) {
  const h = rows[0].map(x => String(x).trim());
  const ix = name => h.indexOf(name);
  const c = { dt: ix('วันเวลาที่ขาย'), doc: ix('เลขที่เอกสาร'), cus: ix('ลูกค้า'), sale: ix('พนักงานขาย'), code: ix('รหัสสินค้า'), pname: ix('ชื่อสินค้า'), qty: ix('จำนวน'), net: ix('สุทธิ'), st: ix('สถานะ'), prov: ix('ผู้ให้บริการ'), del: ix('เลขที่การจัดส่ง') };
  if (c.dt < 0 || c.doc < 0 || c.net < 0) throw new Error('ไฟล์ POS ไม่มีคอลัมน์ วันเวลาที่ขาย / เลขที่เอกสาร / สุทธิ');
  const out = [];
  rows.slice(1).forEach(r => {
    const d = parseThaiShortDT_(r[c.dt]);
    if (!d || !String(r[c.doc] || '').trim()) return;
    const g = i => i >= 0 ? String(r[i] == null ? '' : r[i]).trim() : '';
    out.push([d.key, d.time, g(c.doc), g(c.cus), g(c.sale), g(c.code), g(c.pname), String(num_(r[c.qty])), String(num_(r[c.net])), g(c.st), g(c.prov), g(c.del)]);
  });
  if (!out.length) throw new Error('ไม่พบรายการขายในไฟล์ POS');
  writeMonthly_('POS', POS_HEADERS, out);
  const ks = out.map(r => r[0]).sort();
  return { count: out.length, range: thaiShort_(ks[0]) + ' - ' + thaiShort_(ks[ks.length - 1]), message: 'นำเข้า POS ' + out.length + ' รายการ (' + thaiShort_(ks[0]) + ' - ' + thaiShort_(ks[ks.length - 1]) + ')' };
}

/* ---------- ImedX TLP_05_15 ---------- */
function importImed_(rows) {
  let hi = rows.findIndex(r => r.some(x => String(x).trim() === 'VN/AN'));
  if (hi < 0) throw new Error('ไม่พบหัวตาราง VN/AN ในไฟล์ ImedX');
  const h = rows[hi].map(x => String(x).replace(/\s+/g, '').trim());
  const ix = name => h.indexOf(name);
  const c = { vn: ix('VN/AN'), hn: ix('HN'), pet: ix('ชื่อคนไข้'), own: ix('ชื่อเจ้าของ'), date: ix('วันที่จ่าย'), time: ix('เวลา'), ord: ix('ผู้สั่ง'), code: ix('รหัส'), item: ix('รายการ'), qty: ix('จำนวน'), val: ix('มูลค่ารวม') };
  if (c.vn < 0 || c.date < 0 || c.code < 0 || c.val < 0) throw new Error('ไฟล์ ImedX ไม่มีคอลัมน์ที่ต้องใช้ (VN/AN, วันที่จ่าย, รหัส, มูลค่ารวม)');
  const out = [];
  let total = 0;
  rows.slice(hi + 1).forEach(r => {
    if (!/^\s*[\d,]+\.?\s*$/.test(String(r[0] || ''))) return;
    total++;
    const code = String(r[c.code] || '').trim();
    if (!PRODUCT_RE.test(code)) return;          // เก็บเฉพาะสินค้า 2-, A01–A17
    const d = parseDT_(r[c.date]);
    if (!d) return;
    const g = i => i >= 0 ? String(r[i] == null ? '' : r[i]).trim() : '';
    out.push([dayKey_(d), g(c.time).replace('.', ':'), g(c.vn), g(c.hn), g(c.pet), g(c.own).replace(/\s+/g, ' '), g(c.ord), code, g(c.item), String(num_(r[c.qty])), String(num_(r[c.val]))]);
  });
  if (!total) throw new Error('ไม่พบรายการในไฟล์ ImedX');
  if (out.length) writeMonthly_('IMED', IMED_HEADERS, out);
  const ks = out.map(r => r[0]).sort();
  const rg = ks.length ? thaiShort_(ks[0]) + ' - ' + thaiShort_(ks[ks.length - 1]) : '';
  return { count: out.length, range: rg, note: 'ทั้งไฟล์ ' + total + ' รายการ', message: 'นำเข้า ImedX เฉพาะสินค้า 2-, A01–A17 จำนวน ' + out.length + ' จาก ' + total + ' รายการ (' + rg + ')' };
}

function getImports() { auth_(); return lastImports_(); }

function lastImports_() {
  const s = sheet_(SH.LOG);
  const n = s.getLastRow() - 1;
  if (n < 1) return {};
  const out = {};
  s.getRange(2, 1, n, 5).getValues().forEach(r => {
    if (r[0] instanceof Date) out[r[2]] = Utilities.formatDate(r[0], TZ, 'd/M/yyyy HH:mm') + (r[4] ? ' (ข้อมูล ' + r[4] + ')' : '');
  });
  return out;
}

/* =========================================================
 * รายงาน 1: นัดหมายรายคน
 * ========================================================= */
function matchStaff_(note, staff) {
  let best = -1, pos = Infinity;
  staff.forEach((s, i) => {
    const p = String(note).search(new RegExp('#\\s?' + s.code.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'));
    if (p >= 0 && (p < pos || (p === pos && s.code.length > staff[best].code.length))) { best = i; pos = p; }
  });
  return best;
}

function anyTag_(note) { const m = String(note).match(/#\s?([A-Za-z]{2,4})/); return m ? m[1].toUpperCase() : ''; }

function getApptReport(dateKey) {
  auth_();
  dateKey = dateKey || todayKey_();
  const st = getSettings_();
  const rep = buildApptReport_(readAppt_(), getStaff_(), getTargets_(Number(dateKey.slice(0, 4))), dateKey, st.prepOnlyCoded, getStaff_(true));
  rep.title = st.title1;
  rep.imports = lastImports_();
  return rep;
}

function buildApptReport_(recs, staff, targets, dateKey, prepOnlyCoded, allStaff) {
  allStaff = allStaff || staff;
  const mIdx = Number(dateKey.slice(5, 7)) - 1;
  const monthKey = dateKey.slice(0, 7);
  const tomorrow = addDays_(dateKey, 1);
  const rows = staff.map(s => {
    const t = targets[s.code] || { day: 0, months: [] };
    return { name: s.nick ? s.name + ' (' + s.nick + ')' : s.name, code: s.code, dayT: t.day || 0, monthT: t.months[mIdx] || 0, today: 0, month: 0 };
  });
  const created = [], prep = [];
  // แสดงชื่อเล่นผู้ทำนัด (รวมคนที่ย้ายแผนกแล้ว) — ถ้ารหัสไม่อยู่ในรายชื่อ แสดง #รหัส, ไม่มี # แสดง "ไม่ระบุ"
  const item = r => {
    const j = matchStaff_(r.note, allStaff), tag = anyTag_(r.note);
    const who = j >= 0 ? (allStaff[j].nick || allStaff[j].name) : (tag ? '#' + tag : 'ไม่ระบุ');
    return {
      created: Utilities.formatDate(r.created, TZ, 'HH:mm'), appt: thaiShort_(dayKey_(r.appt)), time: r.time,
      hn: r.hn, name: r.name, pet: r.pet, note: r.note, who: who, tag: who
    };
  };
  recs.forEach(r => {
    const i = matchStaff_(r.note, staff);
    const ck = dayKey_(r.created), ak = dayKey_(r.appt);
    if (ck === dateKey) { if (i >= 0) rows[i].today++; created.push(Object.assign(item(r), { _t: r.created.getTime() })); }
    if (ak.slice(0, 7) === monthKey && i >= 0) rows[i].month++;
    if (ak === tomorrow && (!prepOnlyCoded || anyTag_(r.note))) prep.push(item(r));
  });
  created.sort((a, b) => a._t - b._t).forEach(x => delete x._t);
  prep.sort((a, b) => String(a.time).localeCompare(String(b.time)));
  const pct = (a, b) => b ? a / b : 0;
  rows.forEach(r => {
    r.todayPct = pct(r.today, r.dayT); r.monthPct = pct(r.month, r.monthT);
    r.remDay = Math.max(0, r.dayT - r.today); r.remMonth = Math.max(0, r.monthT - r.month);
  });
  const sum = f => rows.reduce((a, r) => a + r[f], 0);
  const team = targets.TEAM || { day: 0, months: [] };
  const dayT = team.day || sum('dayT');
  const monthT = team.months[mIdx] || sum('monthT');
  const total = { today: sum('today'), month: sum('month'), dayT: dayT, monthT: monthT, remDay: sum('remDay'), remMonth: sum('remMonth') };
  total.todayPct = pct(total.today, dayT); total.monthPct = pct(total.month, monthT);
  return { dateKey: dateKey, dateText: thaiShort_(dateKey), tomorrowText: thaiShort_(tomorrow), monthText: MONTHS_TH[mIdx] + ' ' + dateKey.slice(0, 4), rows: rows, total: total, created: created, prep: prep };
}

/* =========================================================
 * รายงาน 2: สรุปยอดประจำวัน
 * ========================================================= */
function getDailyReport(saleKey, apptKey) {
  auth_();
  saleKey = saleKey || todayKey_();
  apptKey = apptKey || saleKey;
  const st = getSettings_();
  const pos = readRange_('POS', POS_HEADERS.length, saleKey, saleKey);
  const imed = readRange_('IMED', IMED_HEADERS.length, saleKey, saleKey);
  const appts = readAppt_().filter(r => dayKey_(r.appt) === apptKey && (!st.pintoOnly || /ปิ่นโต/.test(r.note)));
  const rep = buildDailyReport_(pos, imed, appts, st.vets);
  rep.manual = getManual_(saleKey);
  rep.title = st.title2;
  rep.saleKey = saleKey; rep.apptKey = apptKey;
  rep.saleText = thaiShort_(saleKey); rep.apptText = thaiShort_(apptKey);
  rep.hasPos = pos.length > 0; rep.hasImed = imed.length > 0;
  rep.imports = lastImports_();
  return rep;
}

/** สรุปบิล POS: ตัดบิลที่ถูก VOID และบิลที่ยอดสุทธิ <= 0 */
function posBills_(pos) {
  const voided = {};
  pos.forEach(r => { if (/^VOID-/i.test(r[2])) voided[r[2].replace(/^VOID-/i, '')] = 1; });
  const bills = {};
  pos.forEach(r => {
    const doc = r[2];
    if (/^VOID-/i.test(doc) || voided[doc] || /ยกเลิก/.test(r[9])) return;
    const b = bills[doc] = bills[doc] || { doc: doc, cus: r[3], net: 0, del: '', prov: '', time: r[1] };
    b.net += num_(r[8]);
    if (r[11]) { b.del = r[11]; b.prov = r[10]; }
  });
  return Object.keys(bills).map(k => bills[k]).filter(b => b.net > 0);
}

/** สรุป VN ImedX (เฉพาะสินค้า) ที่ยอดรวม > 0 */
function imedVisits_(imed, filterFn) {
  const v = {};
  imed.forEach(r => {
    if (filterFn && !filterFn(r)) return;
    const x = v[r[2]] = v[r[2]] || { vn: r[2], hn: r[3], pet: r[4], own: r[5], ord: r[6], val: 0 };
    x.val += num_(r[10]);
  });
  return Object.keys(v).map(k => v[k]).filter(x => x.val > 0);
}

/** VISIT / BILL ของ 1 วัน จาก POS + ImedX รวมกัน
 *  BILL  = บิล POS (สุทธิ > 0) + VN ImedX (สินค้า > 0)
 *  VISIT = ลูกค้าไม่ซ้ำ จับคู่ชื่อหลังตัด "คุณ" (ลูกค้าไม่มีชื่อ/ลูกค้าทั่วไป นับ 1 ต่อ 1 บิล) */
function visitBill_(pos, imed) {
  const bills = posBills_(pos), vns = imedVisits_(imed);
  const names = {};
  let anon = 0;
  bills.map(b => b.cus).concat(vns.map(v => v.own)).forEach(n => {
    const k = normName_(n);
    if (!k || k === 'ลูกค้าทั่วไป' || k === '-') anon++; else names[k] = 1;
  });
  return { bill: bills.length + vns.length, visit: Object.keys(names).length + anon,
    posAmt: bills.reduce((a, b) => a + b.net, 0), imedAmt: vns.reduce((a, v) => a + v.val, 0) };
}

function buildDailyReport_(pos, imed, appts, vets) {
  const bills = posBills_(pos);
  const visits = imedVisits_(imed);
  const sumF = (a, f) => a.reduce((s, x) => s + x[f], 0);

  // 1) Traffic
  const traffic = { bill: bills.length, billAmt: sumF(bills, 'net'), visit: visits.length, visitAmt: sumF(visits, 'val') };
  traffic.count = traffic.bill + traffic.visit; traffic.amount = traffic.billAmt + traffic.visitAmt;
  traffic.customers = visitBill_(pos, imed).visit;

  // 2) ปิ่นโตมาตามนัด: จับคู่ชื่อเจ้าของ (ตัด "คุณ")
  const buy = {};
  bills.forEach(b => { const k = normName_(b.cus); if (k) (buy[k] = buy[k] || { pos: 0, imed: 0 }).pos += b.net; });
  visits.forEach(v => { const k = normName_(v.own); if (k) (buy[k] = buy[k] || { pos: 0, imed: 0 }).imed += v.val; });
  const seen = {}, pinto = [];
  appts.forEach(a => {
    const k = normName_(a.name);
    if (!k || seen[k]) return;
    seen[k] = 1;
    if (buy[k]) pinto.push({ name: a.name.replace(/^\s*คุณ\s*/, ''), pet: a.pet, note: a.note, pos: buy[k].pos, imed: buy[k].imed, total: buy[k].pos + buy[k].imed });
  });
  const pintoSum = { appts: Object.keys(seen).length, count: pinto.length, amount: sumF(pinto, 'total') };

  // 4) Same day: เลขที่การจัดส่งไม่ซ้ำ
  const del = {};
  bills.forEach(b => { if (b.del) { const x = del[b.del] = del[b.del] || { del: b.del, prov: b.prov, net: 0 }; x.net += b.net; } });
  const sameday = Object.keys(del).map(k => del[k]).filter(x => x.net > 0);
  const prov = {};
  sameday.forEach(x => { const p = x.prov || 'อื่น ๆ'; prov[p] = prov[p] || { count: 0, amount: 0 }; prov[p].count++; prov[p].amount += x.net; });

  // 5) หมอสั่งจ่าย
  const isVet = r => vets.some(p => String(r[6]).trim().indexOf(p) === 0);
  const vetV = imedVisits_(imed, isVet);

  return {
    traffic: traffic,
    pinto: pinto.sort((a, b) => b.total - a.total), pintoSum: pintoSum,
    sameday: { count: sameday.length, amount: sumF(sameday, 'net'), byProvider: prov },
    vet: { count: vetV.length, amount: sumF(vetV, 'val'), list: vetV.map(v => ({ vn: v.vn, own: v.own.replace(/^\s*คุณ\s*/, ''), pet: v.pet, ord: v.ord, val: v.val })) }
  };
}

function getManual_(key) {
  const s = sheet_(SH.MANUAL);
  const n = s.getLastRow() - 1;
  if (n < 1) return { count: '', amount: '' };
  const r = s.getRange(2, 1, n, 3).getDisplayValues().filter(x => x[0] === key).pop();
  return r ? { count: r[1], amount: r[2] } : { count: '', amount: '' };
}

function saveManual(key, count, amount) {
  bumpVersion_();
  auth_();
  const s = sheet_(SH.MANUAL);
  const n = s.getLastRow() - 1;
  const keys = n > 0 ? s.getRange(2, 1, n, 1).getDisplayValues().map(r => r[0]) : [];
  const i = keys.indexOf(key);
  const row = [key, count === '' ? '' : num_(count), amount === '' ? '' : num_(amount), new Date()];
  s.getRange(i >= 0 ? i + 2 : s.getLastRow() + 1, 1, 1, 1).setNumberFormat('@');
  if (i >= 0) s.getRange(i + 2, 1, 1, 4).setValues([row]); else s.appendRow(row);
  return 'บันทึกแล้ว';
}

function manualRange_(start, end) {
  const s = sheet_(SH.MANUAL);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, 3).getDisplayValues().filter(r => r[0] >= start && r[0] <= end);
}

/* =========================================================
 * รายงาน 3: รายสัปดาห์ (สัปดาห์ จันทร์–อาทิตย์ ภายในเดือน)
 * ========================================================= */
function getWeeklyReport(monthKey) {
  auth_();
  monthKey = monthKey || todayKey_().slice(0, 7);
  const st = getSettings_();
  const y = Number(monthKey.slice(0, 4)), mIdx = Number(monthKey.slice(5, 7)) - 1;
  const start = monthKey + '-01';
  const daysIn = new Date(y, mIdx + 1, 0).getDate();
  const end = monthKey + '-' + pad_(daysIn);
  const staff = getStaff_();
  const targets = getTargets_(y);
  const team = targets.TEAM || { months: [] };
  const apptT = team.months[mIdx] || staff.reduce((a, s) => a + ((targets[s.code] || { months: [] }).months[mIdx] || 0), 0);
  const visitT = (targets.VISIT || { months: [] }).months[mIdx] || 0;
  const billT = (targets.BILL || { months: [] }).months[mIdx] || 0;

  // แบ่งสัปดาห์
  const weeks = [];
  let cur = null;
  // W1 = 1–7, W2 = 8–14, W3 = 15–21, W4 = 22–สิ้นเดือน
  const today = todayKey_();
  for (let d = 1; d <= daysIn; d++) {
    const k = monthKey + '-' + pad_(d);
    if (!cur || (d - 1) % 7 === 0 && d <= 22) { cur = { start: k, end: k, days: 0 }; weeks.push(cur); }
    cur.end = k; cur.days++;
  }
  weeks.forEach(w => { w.future = w.start > today; w.endEff = w.end < today ? w.end : today; });
  const wOf = k => weeks.findIndex(w => k >= w.start && k <= w.end);
  weeks.forEach((w, i) => Object.assign(w, {
    label: 'W' + (i + 1), range: Number(w.start.slice(8)) + '-' + Number(w.end.slice(8)) + ' ' + MONTHS_SHORT[mIdx],
    apptT: apptT * w.days / daysIn, visitT: visitT * w.days / daysIn, billT: billT * w.days / daysIn,
    appt: 0, visit: 0, bill: 0, posAmt: 0, imedAmt: 0, manual: 0, manualAmt: 0
  }));

  readAppt_().forEach(r => {
    const k = dayKey_(r.appt);
    if (k < start || k > end || matchStaff_(r.note, staff) < 0) return;
    weeks[wOf(k)].appt++;
  });
  const pos = readRange_('POS', POS_HEADERS.length, start, end);
  const byDay = {};
  pos.forEach(r => (byDay[r[0]] = byDay[r[0]] || []).push(r));

  const imed = readRange_('IMED', IMED_HEADERS.length, start, end);
  const imDay = {};
  imed.forEach(r => (imDay[r[0]] = imDay[r[0]] || []).push(r));
  const allDays = {};
  Object.keys(byDay).concat(Object.keys(imDay)).forEach(k => allDays[k] = 1);
  Object.keys(allDays).forEach(k => {
    const w = weeks[wOf(k)];
    if (!w) return;
    const x = visitBill_(byDay[k] || [], imDay[k] || []);
    w.visit += x.visit; w.bill += x.bill; w.posAmt += x.posAmt; w.imedAmt += x.imedAmt;
  });
  manualRange_(start, end).forEach(r => { const w = weeks[wOf(r[0])]; if (w) { w.manual += num_(r[1]); w.manualAmt += num_(r[2]); } });

  const tot = { label: 'รวมทั้งเดือน', range: MONTHS_TH[mIdx] + ' ' + y, apptT: apptT, visitT: visitT, billT: billT };
  ['appt', 'visit', 'bill', 'posAmt', 'imedAmt', 'manual', 'manualAmt'].forEach(f => tot[f] = weeks.reduce((a, w) => a + w[f], 0));
  return { title: st.title3, monthKey: monthKey, monthText: MONTHS_TH[mIdx] + ' ' + y, weeks: weeks, total: tot, imports: lastImports_() };
}

/* =========================================================
 * STOCK (TLP_14_09) + Grab Stock
 * ========================================================= */
const STOCK_HEADERS = ['รหัส', 'รายการ', 'หน่วย', 'MIN', 'MAX', 'คงเหลือ', 'Exp_date'];

function importStock_(rows) {
  const head = rows.slice(0, 5).map(r => r.join(' ')).join(' ');
  const m = head.match(/ณ\s*วันที่\s*(\d{1,2}\s*\S+\s*\d{4})\s*(\d{1,2}:\d{2})?/);
  let asOf = '';
  let asOfKey = '';
  if (m) { const k = parseThaiRange_(m[1] + ' ' + m[1]); asOf = k ? thaiShort_(k.start) + (m[2] ? ' ' + m[2] : '') : m[1]; if (k) asOfKey = k.start + ' ' + (m[2] ? pad_(m[2].split(':')[0]) + ':' + m[2].split(':')[1] : '00:00'); }
  const map = {}, order = [];
  rows.forEach(r => {
    if (r.length < 9 || !/^\s*[\d,]+\.?\s*$/.test(String(r[0] || ''))) return;
    const code = String(r[1] || '').trim().toUpperCase();
    if (!code || !PRODUCT_RE.test(code)) return;          // เก็บเฉพาะสินค้า Pet shop (2-, A01–A17)
    if (!map[code]) { map[code] = [code, String(r[2] || '').trim(), String(r[3] || '').trim(), num_(r[5]), num_(r[6]), 0, String(r[9] || '').trim()]; order.push(code); }
    map[code][5] += num_(r[8]);               // รวมทุก Batch
  });
  if (!order.length) throw new Error('ไม่พบรายการสินค้าในไฟล์ STOCK');
  const s = ensureSheet_(SH.STOCK, STOCK_HEADERS);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, STOCK_HEADERS.length).clearContent();
  s.getRange(2, 1, order.length, 3).setNumberFormat('@');
  s.getRange(2, 1, order.length, STOCK_HEADERS.length).setValues(order.map(c => map[c]));
  PropertiesService.getScriptProperties().setProperties({ stockAsOf: asOf, stockAsOfKey: asOfKey });
  bumpBase_(); bumpVersion_();
  return { count: order.length, range: asOf, message: 'นำเข้า STOCK เฉพาะสินค้า Pet shop (2-, A01–A17) ' + order.length + ' รายการ (ข้อมูล ณ ' + asOf + ')' };
}

function readStockMap_() {
  const s = SpreadsheetApp.getActive().getSheetByName(SH.STOCK);
  const out = {};
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 6).getValues().forEach(r => { if (r[0]) out[String(r[0]).trim()] = { name: r[1], qty: Number(r[5]) || 0 }; });
  return out;
}

function csvCell_(v) {
  v = v == null ? '' : String(v);
  return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}

/** pure: เติมสต็อกลงไฟล์ Grab — กฎเดียวกับที่ทำอยู่: Current = Max = คงเหลือ, >0 = in stock, 0 = out-of-stock indefinitely */
function fillGrab_(rows, stock, buffer, claims) {
  claims = claims || {};
  const hi = rows.findIndex(r => r.some(x => String(x).trim() === 'ITEMCODE') && r.some(x => String(x).trim() === 'CURRENTSTOCK'));
  if (hi < 0) throw new Error('ไม่พบแถวหัวตาราง STOREID ในไฟล์ Grab');
  const h = rows[hi].map(x => String(x).trim().toUpperCase());
  const c = { code: h.indexOf('ITEMCODE'), name: h.indexOf('ITEMNAME'), cur: h.indexOf('CURRENTSTOCK'), max: h.indexOf('MAXSTOCK'), st: h.indexOf('ITEMSTATUS') };
  if (c.code < 0 || c.cur < 0 || c.st < 0) throw new Error('ไฟล์ Grab ไม่มีคอลัมน์ ITEMCODE / CURRENTSTOCK / ITEMSTATUS');
  const sum = { total: 0, matched: 0, inStock: 0, out: 0, notFound: [], changed: [], claimed: 0, codes: [] };
  rows.forEach((r, i) => {
    if (i <= hi || !String(r[0] || '').trim()) return;
    sum.total++;
    const code = String(r[c.code] || '').trim();
    sum.codes.push(code);
    const st = stock[code];
    if (!st) { sum.notFound.push({ code: code, name: r[c.name], status: r[c.st] }); return; }
    sum.matched++;
    const cl = claims[code.toUpperCase()] || 0;          // หักสินค้าเคลม
    if (cl) sum.claimed++;
    const q = Math.max(0, Math.floor(st.qty) - cl - buffer);   // ติดลบ → 0
    const old = String(r[c.cur] || '');
    const val = q > 0 ? String(q) : '';
    r[c.cur] = val;
    if (c.max >= 0) r[c.max] = val;
    r[c.st] = q > 0 ? 'in stock' : 'out-of-stock indefinitely';
    if (q > 0) sum.inStock++; else sum.out++;
    if (old !== val) sum.changed.push({ code: code, name: r[c.name], old: old || '0', now: val || '0' });
  });
  const csv = rows.map(r => r.map(csvCell_).join(',')).join('\n') + '\n';
  return { csv: csv, summary: sum };
}

/** เรียกจากหน้าเว็บ: ส่งไฟล์ Grab มา → คืนไฟล์ CSV ที่เติมสต็อกแล้ว */
function buildGrabStock(text, fileName) {
  auth_();
  text = String(text || '').replace(/^﻿/, '');
  const rows = Utilities.parseCsv(text);
  if (detectType_(rows) !== 'GRAB') throw new Error('ไฟล์นี้ไม่ใช่ไฟล์สต็อกของ Grab');
  const stock = readStockMap_();
  if (!Object.keys(stock).length) throw new Error('ยังไม่มีข้อมูล STOCK — นำเข้าไฟล์ TLP_14_09 ก่อน');
  const st = getSettings_();
  const claimMap = {};
  readClaims_().forEach(c => claimMap[c.code] = (claimMap[c.code] || 0) + c.qty);
  const res = fillGrab_(rows, stock, st.grabBuffer, claimMap);
  saveGrabCodes_(res.summary.codes);
  delete res.summary.codes;
  res.fileName = fileName || 'grab_stock.csv';
  res.stockAsOf = PropertiesService.getScriptProperties().getProperty('stockAsOf') || '';
  res.today = thaiShort_(todayKey_());
  res.buffer = st.grabBuffer;
  sheet_(SH.LOG).appendRow([new Date(), res.fileName, 'GRAB', res.summary.matched, res.stockAsOf, 'in stock ' + res.summary.inStock + ' / หมด ' + res.summary.out + ' / ไม่พบ ' + res.summary.notFound.length]);
  return res;
}

/** สำรอง: บันทึกไฟล์ลง Google Drive แล้วคืนลิงก์ดาวน์โหลด */
function saveCsvToDrive(csv, fileName) {
  auth_();
  const it = DriveApp.getFoldersByName('Grab Stock Export');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Grab Stock Export');
  const f = folder.createFile(fileName, csv, MimeType.CSV);
  return 'https://drive.google.com/uc?export=download&id=' + f.getId();
}

/* =========================================================
 * รายงาน 5: ลูกค้าซื้อสินค้า แต่ยังไม่มีนัดในระบบ
 *   - ซื้อสินค้า (POS + ImedX) ในช่วงวันที่เลือก เฉพาะกลุ่มสินค้าที่เลือก
 *   - "มีนัดแล้ว" = มีนัดในระบบที่วันที่นัด >= วันที่ซื้อครั้งล่าสุด (จับคู่ HN หรือชื่อหลังตัด "คุณ")
 * ========================================================= */
function getNoApptReport(start, end, groups, minBill) {
  auth_();
  end = end || todayKey_();
  start = start || end.slice(0, 8) + '01';
  if (start > end) { const t = start; start = end; end = t; }
  const pos = readRange_('POS', POS_HEADERS.length, start, end);
  const imed = readRange_('IMED', IMED_HEADERS.length, start, end);
  const res = buildNoAppt_(pos, imed, readAppt_(), groups && groups.length ? groups : ['2', 'A01', 'A09'], Number(minBill) || 0, start, end);
  res.start = start; res.end = end; res.startText = thaiShort_(start); res.endText = thaiShort_(end);
  const fl = latestFollows_();
  res.rows.forEach(x => { x.follow = fl[x.key] || null; });
  res.today = todayKey_();
  res.imports = lastImports_();
  return res;
}

function groupOf_(code) {
  code = String(code || '').trim().toUpperCase();
  if (/^2-/.test(code)) return '2';
  const m = code.match(/^(A\d{2})-/);
  return m ? m[1] : '';
}

function buildNoAppt_(pos, imed, appts, groups, minBill, start, end) {
  const want = {};
  groups.forEach(g => want[String(g).toUpperCase().replace(/-$/, '')] = 1);
  // นัดหมาย: วันที่นัดล่าสุด ต่อ HN และต่อชื่อ
  const apHN = {}, apName = {};
  appts.forEach(a => {
    const k = dayKey_(a.appt), hn = String(a.hn).trim().toUpperCase(), nm = normName_(a.name);
    if (hn && (!apHN[hn] || apHN[hn] < k)) apHN[hn] = k;
    if (nm && (!apName[nm] || apName[nm] < k)) apName[nm] = k;
  });
  // บิล: POS (ตัด VOID) + ImedX ต่อ VN
  const voided = {};
  pos.forEach(r => { if (/^VOID-/i.test(r[2])) voided[r[2].replace(/^VOID-/i, '')] = 1; });
  const bills = {};
  const addItem = (id, date, cus, hn, who, code, name, qty, amt) => {
    if (!want[groupOf_(code)] || amt <= 0) return;
    const b = bills[id] = bills[id] || { date: date, cus: cus, hn: hn, who: who, amt: 0, items: [] };
    b.amt += amt; b.items.push({ code: String(code).trim().toUpperCase(), name: name, qty: qty, amt: amt });
  };
  pos.forEach(r => {
    if (/^VOID-/i.test(r[2]) || voided[r[2]] || /ยกเลิก/.test(r[9])) return;
    addItem('P' + r[2], r[0], r[3], '', r[4], r[5], r[6], num_(r[7]), num_(r[8]));
  });
  imed.forEach(r => addItem('I' + r[0] + r[2], r[0], r[5], r[3], r[6], r[7], r[8], num_(r[9]), num_(r[10])));

  // เดือนในช่วงที่เลือก
  const months = [];
  if (start && end) for (let k = start.slice(0, 7); k <= end.slice(0, 7);) {
    months.push(k);
    const y = Number(k.slice(0, 4)), m = Number(k.slice(5, 7));
    k = m === 12 ? (y + 1) + '-01' : y + '-' + pad_(m + 1);
  }
  const lastM = months[months.length - 1];

  const cus = {};
  Object.keys(bills).forEach(id => {
    const b = bills[id];
    const k = normName_(b.cus);
    if (!k || k === 'ลูกค้าทั่วไป' || k === '-') return;
    const c = cus[k] = cus[k] || { name: String(b.cus).replace(/^\s*คุณ\s*/, '').replace(/\s+/g, ' ').trim(), hn: {}, bills: 0, amt: 0, maxBill: 0, last: '', prod: {}, who: {}, mon: {} };
    c.bills++; c.amt += b.amt; c.maxBill = Math.max(c.maxBill, b.amt);
    if (b.date > c.last) c.last = b.date;
    if (b.hn) c.hn[String(b.hn).trim().toUpperCase()] = 1;
    const mk = b.date.slice(0, 7);
    c.mon[mk] = (c.mon[mk] || 0) + b.amt;
    b.items.forEach(it => {
      const p = c.prod[it.code] = c.prod[it.code] || { code: it.code, name: it.name, last: '', amt: 0, qty: 0, m: {}, q: {} };
      if (b.date >= p.last) { p.last = b.date; p.name = it.name; }
      p.amt += it.amt; p.qty += it.qty;
      p.m[mk] = (p.m[mk] || 0) + it.amt; p.q[mk] = (p.q[mk] || 0) + it.qty;
    });
    if (b.who) c.who[String(b.who).replace(/\(.*\)/, '').trim()] = 1;
  });
  const r2 = x => Math.round(x * 100) / 100;
  const rows = [];
  let withAppt = 0;
  Object.keys(cus).forEach(k => {
    const c = cus[k];
    if (c.maxBill < minBill) return;
    const hns = Object.keys(c.hn);
    let ap = apName[k] || '';
    hns.forEach(h => { if (apHN[h] && apHN[h] > ap) ap = apHN[h]; });
    const has = !!ap && ap >= c.last;
    if (has) withAppt++;
    const prods = Object.keys(c.prod).map(pk => {
      const p = c.prod[pk];
      const boughtM = Object.keys(p.m).length;
      return {
        code: p.code, name: p.name, last: p.last, lastText: thaiShort_(p.last), amt: r2(p.amt), qty: p.qty,
        months: months.map(m => p.m[m] ? r2(p.m[m]) : 0), qtys: months.map(m => p.q[m] || 0),
        // ซื้อประจำ (≥2 เดือนก่อนหน้า) แต่ยังไม่ซื้อเดือนล่าสุด
        missing: months.length > 1 && !p.m[lastM] && boughtM >= 2
      };
    }).sort((a, b) => b.last.localeCompare(a.last) || b.amt - a.amt);
    rows.push({
      key: k, name: c.name, hn: hns.join(', '), bills: c.bills, amt: r2(c.amt), maxBill: r2(c.maxBill),
      last: c.last, lastText: thaiShort_(c.last), appt: ap, apptText: ap ? thaiShort_(ap) : '', hasAppt: has,
      monthAmt: months.map(m => r2(c.mon[m] || 0)), prods: prods, missing: prods.some(p => p.missing),
      items: prods.slice(0, 4).map(p => p.name), who: Object.keys(c.who).slice(0, 3).join(', ')
    });
  });
  rows.sort((a, b) => (a.hasAppt - b.hasAppt) || (b.amt - a.amt));
  return { rows: rows, total: rows.length, withAppt: withAppt, without: rows.length - withAppt, groups: Object.keys(want),
    months: months, monthText: months.map(m => MONTHS_SHORT[Number(m.slice(5, 7)) - 1] + ' ' + String(Number(m.slice(0, 4)) + 543).slice(2)) };
}

/* =========================================================
 * รายงาน 6: ตรวจสอบยอด Sameday (Grab / LINE MAN) จาก POS
 * ========================================================= */
const DAYS_TH = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

function providerOf_(prov, del) {
  const p = String(prov || '').trim();
  if (p) return /grab/i.test(p) ? 'Grab' : /line/i.test(p) ? 'LINE MAN' : p;
  return /^GM-/i.test(del) ? 'Grab' : /^LME-/i.test(del) ? 'LINE MAN' : 'อื่น ๆ';
}

/** pure: รวมรายการตามเลขที่การจัดส่ง */
function buildSameday_(pos, seller) {
  const voided = {};
  pos.forEach(r => { if (/^VOID-/i.test(r[2])) voided[r[2].replace(/^VOID-/i, '')] = 1; });
  const sellers = {}, map = {};
  pos.forEach(r => {
    const del = String(r[11] || '').trim();
    if (!del) return;
    if (r[4]) sellers[r[4]] = 1;
    const x = map[del] = map[del] || { del: del, prov: providerOf_(r[10], del), time: '', net: 0, gross: 0, cancelled: false, docs: {}, voids: {}, sellers: {} };
    const isVoid = /^VOID-/i.test(r[2]) || voided[r[2]] || /ยกเลิก/.test(r[9]);
    if (!/^VOID-/i.test(r[2])) { x.docs[r[2]] = 1; if (r[1] > x.time) x.time = r[1]; x.gross += num_(r[8]); }
    if (!isVoid) x.net += num_(r[8]); else if (!/^VOID-/i.test(r[2])) x.voids[r[2]] = 1;
    if (r[4]) x.sellers[r[4]] = 1;
  });
  const list = Object.keys(map).map(k => {
    const x = map[k];
    x.cancelled = x.net <= 0;
    x.voidDocs = Object.keys(x.voids).join(', '); delete x.voids;
    x.docs = Object.keys(x.docs).join(', ');
    x.sellerList = Object.keys(x.sellers);
    delete x.sellers;
    x.net = Math.round(x.net * 100) / 100; x.gross = Math.round(x.gross * 100) / 100;
    return x;
  }).filter(x => !seller || x.sellerList.indexOf(seller) >= 0)
    .sort((a, b) => b.time.localeCompare(a.time));
  const sum = {};
  list.forEach(x => { if (x.cancelled) return; const s = sum[x.prov] = sum[x.prov] || { count: 0, amount: 0 }; s.count++; s.amount += x.net; });
  return { list: list, sum: sum, sellers: Object.keys(sellers).sort() };
}

function getSamedayReport(dateKey, seller) {
  auth_();
  dateKey = dateKey || todayKey_();
  const pos = readRange_('POS', POS_HEADERS.length, dateKey, dateKey);
  const r = buildSameday_(pos, seller || '');
  const chk = readChecks_(dateKey);
  r.list.forEach(x => { const c = chk[x.del]; x.checked = !!(c && c.on); x.checkBy = c ? c.by : ''; });
  const d = keyToDate_(dateKey);
  r.dateKey = dateKey;
  r.dateText = 'วัน' + DAYS_TH[d.getDay()] + 'ที่ ' + Number(dateKey.slice(8)) + ' ' + MONTHS_TH[Number(dateKey.slice(5, 7)) - 1] + ' ' + dateKey.slice(0, 4);
  r.hasPos = pos.length > 0;
  r.seller = seller || '';
  r.imports = lastImports_();
  return r;
}

function readChecks_(dateKey) {
  const s = ensureSheet_(SH.CHECK, ['วันที่', 'เลขที่การจัดส่ง', 'ตรวจแล้ว', 'ผู้ตรวจ', 'เวลา']);
  const out = {};
  if (s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 4).getDisplayValues().forEach(r => {
    if (r[0] === dateKey) out[r[1]] = { on: r[2] === 'TRUE' || r[2] === 'ใช่', by: r[3] };
  });
  return out;
}

/** ติ๊ก / เอาติ๊กออก รายการ (บันทึกถาวร) */
function saveSamedayCheck(dateKey, dels, on, by) {
  auth_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = ensureSheet_(SH.CHECK, ['วันที่', 'เลขที่การจัดส่ง', 'ตรวจแล้ว', 'ผู้ตรวจ', 'เวลา']);
    const n = s.getLastRow() - 1;
    const keys = n > 0 ? s.getRange(2, 1, n, 2).getDisplayValues().map(r => r[0] + '|' + r[1]) : [];
    const add = [];
    [].concat(dels).forEach(del => {
      const row = [dateKey, del, on ? 'TRUE' : 'FALSE', (CURRENT_USER && CURRENT_USER.name) || by || '', new Date()];
      const i = keys.indexOf(dateKey + '|' + del);
      if (i >= 0) s.getRange(i + 2, 1, 1, 5).setValues([row]); else add.push(row);
    });
    if (add.length) {
      const r0 = s.getLastRow() + 1;
      s.getRange(r0, 1, add.length, 3).setNumberFormat('@');
      s.getRange(r0, 1, add.length, 5).setValues(add);
    }
    return 'ok';
  } finally { lock.releaseLock(); }
}

/* ---------- สินค้าเคลม (หักออกจาก STOCK ก่อนส่ง Grab) ---------- */
function readClaims_() {
  const s = ensureSheet_(SH.CLAIM, ['รหัสสินค้า', 'จำนวนเคลม', 'หมายเหตุ', 'บันทึกเมื่อ']);
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 3).getDisplayValues()
    .map(r => ({ code: String(r[0]).trim().toUpperCase(), qty: num_(r[1]), note: r[2] }))
    .filter(x => x.code && x.qty > 0);
}

function saveGrabCodes_(codes) {
  const s = ensureSheet_(SH.GRABLIST, ['ITEMCODE ในไฟล์ Grab ล่าสุด']);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 1).clearContent();
  if (codes.length) { s.getRange(2, 1, codes.length, 1).setNumberFormat('@'); s.getRange(2, 1, codes.length, 1).setValues(codes.map(c => [c])); }
}

/** รายการเคลม + ชื่อสินค้า + STOCK ปัจจุบัน / หลังปรับ + มีใน Grab หรือไม่ */
function getClaims() {
  auth_();
  const stock = readStockMap_();
  const g = SpreadsheetApp.getActive().getSheetByName(SH.GRABLIST);
  const inGrab = {};
  if (g && g.getLastRow() > 1) g.getRange(2, 1, g.getLastRow() - 1, 1).getDisplayValues().forEach(r => inGrab[String(r[0]).trim().toUpperCase()] = 1);
  const total = {};
  const list = readClaims_();
  list.forEach(c => total[c.code] = (total[c.code] || 0) + c.qty);
  return {
    stockAsOf: PropertiesService.getScriptProperties().getProperty('stockAsOf') || '',
    hasGrab: Object.keys(inGrab).length > 0,
    list: list.map(c => {
      const st = stock[c.code];
      const now = st ? Math.floor(st.qty) : null;
      return { code: c.code, qty: c.qty, note: c.note, name: st ? String(st.name) : '', found: !!st,
        stock: now, after: now == null ? null : Math.max(0, now - total[c.code]), grab: !!inGrab[c.code] };
    })
  };
}

function saveClaims(list) {
  bumpVersion_();
  auth_();
  const rows = (list || []).map(x => [String(x.code || '').trim().toUpperCase(), num_(x.qty), String(x.note || '').trim(), new Date()])
    .filter(r => r[0] && r[1] > 0);
  const s = ensureSheet_(SH.CLAIM, ['รหัสสินค้า', 'จำนวนเคลม', 'หมายเหตุ', 'บันทึกเมื่อ']);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 4).clearContent();
  if (rows.length) { s.getRange(2, 1, rows.length, 1).setNumberFormat('@'); s.getRange(2, 1, rows.length, 4).setValues(rows); }
  return getClaims();
}

/* =========================================================
 * ตารางเวร (เฉพาะหัวหน้างาน)
 *   MAIN_yyyy-MM : { start, end, cells: { 'yyyy-MM-dd|ชื่อ': 'กะ' } }
 *   HELP_yyyy-MM : { rows: ['yyyy-MM-dd'], cols: [{name,en,emp,phone}], cells: { 'วันที่|ชื่อ': 'กะ' } }
 * ========================================================= */
const ROSTER_SHEET = 'ตารางเวร_DATA';

function rosterGet_(key) {
  const s = ensureSheet_(ROSTER_SHEET, ['KEY', 'DATA (JSON)', 'บันทึกเมื่อ']);
  if (s.getLastRow() < 2) return null;
  const v = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === key) { try { return JSON.parse(v[i][1]); } catch (e) { return null; } }
  return null;
}

function rosterPut_(key, obj) {
  const s = ensureSheet_(ROSTER_SHEET, ['KEY', 'DATA (JSON)', 'บันทึกเมื่อ']);
  const n = s.getLastRow() - 1;
  const keys = n > 0 ? s.getRange(2, 1, n, 1).getValues().map(r => r[0]) : [];
  const i = keys.indexOf(key);
  const row = [key, JSON.stringify(obj), new Date()];
  if (i >= 0) s.getRange(i + 2, 1, 1, 3).setValues([row]); else s.appendRow(row);
}

function rosterSettings_() {
  const def = { title: 'ตารางการทำงาน PET SHOP R9', org: 'PC', dept: 'Pet shop', preparer: '', preparerTitle: '(หัวหน้าหน่วย เพ็ทช็อป)', approver: '', approverTitle: '(ผู้จัดการเพ็ทช็อปพระราม 9)' };
  try { return Object.assign(def, JSON.parse(PropertiesService.getScriptProperties().getProperty('rosterSettings') || '{}')); } catch (e) { return def; }
}

function getRoster(pin, monthKey) {
  checkPin_(pin);
  monthKey = monthKey || todayKey_().slice(0, 7);
  const y = Number(monthKey.slice(0, 4)), m = Number(monthKey.slice(5, 7));
  const last = monthKey + '-' + pad_(new Date(y, m, 0).getDate());
  const main = rosterGet_('MAIN_' + monthKey) || { start: monthKey + '-01', end: last, cells: {} };
  const help = rosterGet_('HELP_' + monthKey) || { rows: [], cols: [], cells: {} };
  return {
    monthKey: monthKey, main: main, help: help, settings: rosterSettings_(), options: getSettings_().shifts,
    staff: getAllStaffRows_().filter(s => s.roster).map(s => ({ name: s.name, nick: s.nick, en: s.en, emp: s.emp, phone: s.phone, hr: s.hr })),
    allStaff: getAllStaffRows_().map(s => ({ name: s.name, nick: s.nick, en: s.en, emp: s.emp, phone: s.phone }))
  };
}

function saveRoster(pin, monthKey, main, help, settings) {
  bumpVersion_();
  checkPin_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const clean = c => { const o = {}; Object.keys(c || {}).forEach(k => { const v = String(c[k] || '').trim(); if (v) o[k] = v; }); return o; };
    if (main) rosterPut_('MAIN_' + monthKey, { start: main.start, end: main.end, cells: clean(main.cells) });
    if (help) rosterPut_('HELP_' + monthKey, { rows: help.rows || [], cols: help.cols || [], cells: clean(help.cells) });
    if (settings) PropertiesService.getScriptProperties().setProperty('rosterSettings', JSON.stringify(settings));
    return 'บันทึกตารางเวรแล้ว';
  } finally { lock.releaseLock(); }
}

/* =========================================================
 * ผู้ใช้งาน / ล็อกอิน
 *   ทุกคำขอจากหน้าเว็บผ่าน api(token, ชื่อฟังก์ชัน, args)
 *   USERS: Username(รหัสพนักงาน) | ชื่อ | สิทธิ์ | ใช้งาน | salt | hash | ล็อกอินล่าสุด
 * ========================================================= */
var CURRENT_USER = null;
const USERS_SHEET = 'USERS';
const USER_HEADERS = ['Username (รหัสพนักงาน)', 'ชื่อ', 'สิทธิ์ (admin/staff)', 'ใช้งาน', 'salt', 'hash', 'ล็อกอินล่าสุด'];
const SESSION_SEC = 6 * 60 * 60;   // ล็อกอินค้างได้ 6 ชั่วโมง (นานสุดที่ระบบรองรับ)

function auth_() { if (!CURRENT_USER) throw new Error('กรุณาเข้าสู่ระบบ'); }

function api(token, fn, args) {
  const FN = {
    getApptReport: [getApptReport, 0], getApptMonth: [getApptMonth, 0], getDailyReport: [getDailyReport, 0], saveManual: [saveManual, 0],
    getWeeklyReport: [getWeeklyReport, 0], getNoApptReport: [getNoApptReport, 0], getSamedayReport: [getSamedayReport, 0],
    saveSamedayCheck: [saveSamedayCheck, 0], importCsvText: [importCsvText, 0], buildGrabStock: [buildGrabStock, 0],
    saveCsvToDrive: [saveCsvToDrive, 0], getClaims: [getClaims, 0], saveClaims: [saveClaims, 0], getImports: [getImports, 0],
    changePassword: [changePassword_, 0],
    getAdminData: [getAdminData, 1], saveStaff: [saveStaff, 1], saveTargets: [saveTargets, 1], getRoster: [getRoster, 1], saveRoster: [saveRoster, 1],
    getUsers: [getUsers_, 1], saveUsers: [saveUsers_, 1], resetPassword: [resetPassword_, 1],
    getVersion: [getVersion_, 0],
    saveFollow: [saveFollow_, 0], getFollowHistory: [getFollowHistory_, 0],
    getPurchaseReport: [getPurchaseReport_, 0], markOrdered: [markOrdered_, 0], getPurchaseConfig: [getPurchaseConfig_, 0], savePurchaseConfig: [savePurchaseConfig_, 1],
    getCabinets: [getCabinets_, 0], getCabinet: [getCabinet_, 0], saveCabinet: [saveCabinet_, 1], deleteCabinet: [deleteCabinet_, 1],
    uploadPhoto: [uploadPhoto_, 0], lookupProduct: [lookupProduct_, 0], searchProducts: [searchProducts_, 0],
    cabMove: [cabMove_, 0], saveShelf: [saveShelf_, 0], getShelfList: [getShelfList_, 0],
    getRefillPlan: [getRefillPlan_, 0], applyRefill: [applyRefill_, 0], syncCabinets: [syncCabinets_, 0]
  };
  const u = session_(token);
  if (!u) throw new Error('SESSION_EXPIRED');
  const f = FN[fn];
  if (!f) throw new Error('ไม่รู้จักคำสั่ง ' + fn);
  if (f[1] && u.role !== 'admin') throw new Error('เฉพาะหัวหน้างาน');
  CURRENT_USER = u;
  try {
    if (!CACHE_FN[fn]) return f[0].apply(null, args || []);
    // แคชผลรายงาน: ใช้ซ้ำได้จนกว่าจะมีการนำเข้า/แก้ข้อมูล (เวอร์ชันเปลี่ยน) หรือเปลี่ยนวัน
    const key = 'r_' + fn + '_' + getVersion_() + '_' + todayKey_() + '_' + md5_(JSON.stringify(args || []));
    const hit = cacheGetJson_(key);
    if (hit !== null) return hit;
    const res = f[0].apply(null, args || []);
    cachePutJson_(key, res, 1800);
    return res;
  } finally { CURRENT_USER = null; }
}

/* ---------- แคช (เร่งความเร็ว) ---------- */
const CACHE_FN = { getApptReport: 1, getApptMonth: 1, getDailyReport: 1, getWeeklyReport: 1, getNoApptReport: 1, getPurchaseReport: 1,
  getImports: 1, getCabinets: 1, getCabinet: 1, getRefillPlan: 1, getShelfList: 1, getClaims: 1, lookupProduct: 1, searchProducts: 1, getPurchaseConfig: 1 };
function md5_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function cacheGetJson_(key) {
  try {
    const c = CacheService.getScriptCache(), n = Number(c.get(key + '#n') || 0);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(key + '#' + i);
    const m = c.getAll(keys);
    let s = '';
    for (let i = 0; i < n; i++) { const v = m[key + '#' + i]; if (v == null) return null; s += v; }
    return JSON.parse(s);
  } catch (e) { return null; }
}
function cachePutJson_(key, obj, ttl) {
  try {
    const s = JSON.stringify(obj);
    if (!s || s.length > 2500000) return;
    const parts = {};
    let n = 0;
    for (let i = 0; i < s.length; i += 30000) parts[key + '#' + (n++)] = s.slice(i, i + 30000);
    parts[key + '#n'] = String(n);
    CacheService.getScriptCache().putAll(parts, ttl || 1800);
  } catch (e) {}
}
function baseVersion_() { return PropertiesService.getScriptProperties().getProperty('baseVersion') || '0'; }
function bumpBase_() { try { PropertiesService.getScriptProperties().setProperty('baseVersion', String(Date.now())); } catch (e) {} }

function hash_(salt, pass) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function newSalt_() { return Utilities.getUuid().replace(/-/g, ''); }
function usersSheet_() { return ensureSheet_(USERS_SHEET, USER_HEADERS); }
function readUsers_() {
  const s = usersSheet_();
  if (s.getLastRow() < 2) return [];
  const disp = s.getRange(2, 1, s.getLastRow() - 1, 6).getDisplayValues();
  return s.getRange(2, 1, s.getLastRow() - 1, 7).getValues().map((r, i) => ({
    row: i + 2, user: String(disp[i][0]).trim(), name: String(r[1]).trim(), role: String(r[2]).trim() === 'admin' ? 'admin' : 'staff',
    active: r[3] !== false, salt: String(disp[i][4]), hash: String(disp[i][5]), last: r[6]
  })).filter(x => x.user);
}

function session_(token) {
  if (!token) return null;
  const v = CacheService.getScriptCache().get('s_' + token);
  if (!v) return null;
  CacheService.getScriptCache().put('s_' + token, v, SESSION_SEC);   // ต่ออายุทุกครั้งที่ใช้งาน
  return JSON.parse(v);
}

/** เข้าสู่ระบบ — ถ้ายังไม่มีหัวหน้างานในระบบ ใช้ username "admin" + รหัสผ่านผู้ดูแลในชีตตั้งค่า (ว่าง = admin) */
function login(username, password) {
  username = String(username || '').trim(); password = String(password || '');
  if (!username || !password) throw new Error('กรุณากรอก Username และรหัสผ่าน');
  const cache = CacheService.getScriptCache(), fk = 'f_' + username.toLowerCase();
  const fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('ใส่รหัสผิดหลายครั้ง กรุณารอ 5 นาที');
  const users = readUsers_();
  let u = users.filter(x => x.user.toLowerCase() === username.toLowerCase())[0];
  let ok = false, me = null;
  if (u && u.active && hash_(u.salt, password) === u.hash) { ok = true; me = { user: u.user, name: u.name || u.user, role: u.role }; }
  // บัญชีหัวหน้างาน: Username "admin" + รหัสผ่านผู้ดูแลในชีตตั้งค่า (ว่าง = 1234)
  if (!ok && !u && username.toLowerCase() === 'admin') {
    if (checkAdminPass_(password)) { ok = true; me = { user: 'admin', name: 'หัวหน้างาน', role: 'admin', builtin: true }; }
  }
  if (!ok) { cache.put(fk, String(fails + 1), 300); throw new Error('Username หรือรหัสผ่านไม่ถูกต้อง'); }
  cache.remove(fk);
  if (u) usersSheet_().getRange(u.row, 7).setValue(new Date());
  const token = Utilities.getUuid();
  cache.put('s_' + token, JSON.stringify(me), SESSION_SEC);
  me.token = token;
  me.defaultPass = !!u && password === u.user;
  return me;
}

function whoami(token) { return session_(token); }
function logout(token) { if (token) CacheService.getScriptCache().remove('s_' + token); return true; }

function changePassword_(oldPass, newPass) {
  newPass = String(newPass || '');
  if (newPass.length < 4) throw new Error('รหัสผ่านใหม่ต้องยาวอย่างน้อย 4 ตัว');
  if (CURRENT_USER.builtin) {
    if (!checkAdminPass_(String(oldPass || ''))) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    const sa = newSalt_();
    PropertiesService.getScriptProperties().setProperties({ adminSalt: sa, adminHash: hash_(sa, newPass) });
    return 'เปลี่ยนรหัสผ่าน admin แล้ว';
  }
  const u = readUsers_().filter(x => x.user === CURRENT_USER.user)[0];
  if (!u) throw new Error('ไม่พบผู้ใช้');
  if (hash_(u.salt, String(oldPass || '')) !== u.hash) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  const salt = newSalt_();
  usersSheet_().getRange(u.row, 5, 1, 2).setNumberFormat('@').setValues([[salt, hash_(salt, newPass)]]);
  return 'เปลี่ยนรหัสผ่านแล้ว';
}

function getUsers_() {
  return {
    users: readUsers_().map(u => ({ user: u.user, name: u.name, role: u.role, active: u.active,
      last: u.last instanceof Date ? Utilities.formatDate(u.last, TZ, 'd/M/yyyy HH:mm') : '' })),
    staff: getAllStaffRows_().filter(s => s.emp && /[0-9A-Za-z]/.test(s.emp)).map(s => ({ emp: s.emp, name: s.name + (s.nick ? ' (' + s.nick + ')' : '') }))
  };
}

/** บันทึกรายชื่อผู้ใช้ — ผู้ใช้ใหม่ได้รหัสผ่านเริ่มต้น = Username (รหัสพนักงาน) */
function saveUsers_(list) {
  const old = {};
  readUsers_().forEach(u => old[u.user.toLowerCase()] = u);
  const seen = {};
  const rows = (list || []).map(x => ({ user: String(x.user || '').trim(), name: String(x.name || '').trim(), role: x.role === 'admin' ? 'admin' : 'staff', active: x.active !== false }))
    .filter(x => x.user).map(x => {
      const k = x.user.toLowerCase();
      if (seen[k]) throw new Error('Username ' + x.user + ' ซ้ำกัน');
      seen[k] = 1;
      const o = old[k];
      const salt = o ? o.salt : newSalt_();
      return [x.user, x.name, x.role, x.active, salt, o ? o.hash : hash_(salt, x.user), o ? o.last : ''];
    });
  if (seen.admin) throw new Error('Username admin สงวนไว้สำหรับหัวหน้างาน');
  if (!rows.length) { const s0 = usersSheet_(); if (s0.getLastRow() > 1) s0.getRange(2, 1, s0.getLastRow() - 1, 7).clearContent(); return getUsers_(); }
  const s = usersSheet_();
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 7).clearContent();
  s.getRange(2, 1, rows.length, 6).setNumberFormat('@');
  s.getRange(2, 1, rows.length, 7).setValues(rows.map(r => [String(r[0]), r[1], r[2], r[3], String(r[4]), String(r[5]), r[6]]));
  s.hideColumns(5, 2);
  return getUsers_();
}

function resetPassword_(username) {
  const u = readUsers_().filter(x => x.user === username)[0];
  if (!u) throw new Error('ไม่พบผู้ใช้ ' + username + ' (กดบันทึกก่อน)');
  const salt = newSalt_();
  usersSheet_().getRange(u.row, 5, 1, 2).setNumberFormat('@').setValues([[salt, hash_(salt, u.user)]]);
  return 'รีเซ็ตรหัสผ่านของ ' + u.user + ' เป็นรหัสพนักงานแล้ว';
}

/** รหัสผ่าน admin: ถ้าเคยเปลี่ยนในหน้าเว็บ ใช้รหัสที่เปลี่ยน, ไม่เคย → รหัสในชีตตั้งค่า (ว่าง = 1234) */
function checkAdminPass_(password) {
  const pr = PropertiesService.getScriptProperties();
  const h = pr.getProperty('adminHash');
  if (h) return hash_(pr.getProperty('adminSalt'), password) === h;
  return password === (getSettings_().pin || '1234');
}

/** ลืมรหัส admin: เปิด Apps Script แล้วรันฟังก์ชันนี้ → กลับไปใช้รหัสในชีตตั้งค่า (ว่าง = 1234) */
function resetAdminPassword() {
  PropertiesService.getScriptProperties().deleteProperty('adminHash');
  PropertiesService.getScriptProperties().deleteProperty('adminSalt');
  Logger.log('รีเซ็ตรหัส admin แล้ว — ใช้รหัสในชีตตั้งค่า (ว่าง = 1234)');
}

/* =========================================================
 * เวอร์ชันข้อมูล — หน้าเว็บใช้ตรวจว่าข้อมูลเปลี่ยนแล้วหรือยัง (ไม่ให้ค้างข้อมูลเก่า)
 * ========================================================= */
function bumpVersion_() { try { PropertiesService.getScriptProperties().setProperty('dataVersion', String(Date.now())); } catch (e) {} }
function getVersion_() { return PropertiesService.getScriptProperties().getProperty('dataVersion') || '0'; }

/* =========================================================
 * MASTER สินค้า (TLP_18_05): รหัส | ชื่อ | Barcode | หน่วย | ราคา
 * ========================================================= */
const SH_MASTER = 'MASTER_สินค้า';
function cleanBarcode_(b) { return String(b || '').replace(/[()\s]/g, '').trim(); }

function importMaster_(rows) {
  const hi = rows.findIndex(r => String(r[0]).trim() === 'รหัสยา');
  if (hi < 0) throw new Error('ไม่พบหัวตาราง รหัสยา ในไฟล์ TLP_18_05');
  const h = rows[hi].map(x => String(x).trim());
  const c = { code: 0, name: h.indexOf('ชื่อยา'), bc: h.indexOf('Barcode'), unit: h.indexOf('หน่วย'), price: h.indexOf('ราคา') };
  const seen = {}, out = [];
  rows.slice(hi + 1).forEach(r => {
    const code = String(r[0] || '').trim().toUpperCase();
    if (!code || seen[code] || !PRODUCT_RE.test(code)) return;     // เก็บเฉพาะสินค้า 2-, A01–A17
    seen[code] = 1;
    out.push([code, String(r[c.name] || '').trim(), cleanBarcode_(r[c.bc]), String(r[c.unit] || '').trim(), String(num_(r[c.price]))]);
  });
  if (!out.length) throw new Error('ไม่พบรายการสินค้าในไฟล์ Master');
  const s = ensureSheet_(SH_MASTER, ['รหัส', 'ชื่อสินค้า', 'Barcode', 'หน่วย', 'ราคา']);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 5).clearContent();
  s.getRange(2, 1, out.length, 5).setNumberFormat('@');
  s.getRange(2, 1, out.length, 5).setValues(out);
  bumpBase_();
  return { count: out.length, message: 'นำเข้า Master สินค้า ' + out.length + ' รายการ (มี Barcode ' + out.filter(r => r[2]).length + ')' };
}

function readMaster_() {
  const ck = 'b_master_' + baseVersion_(), hit = cacheGetJson_(ck);
  if (hit) return hit;
  const out = readMasterRaw_();
  cachePutJson_(ck, out, 21600);
  return out;
}
function readMasterRaw_() {
  const s = SpreadsheetApp.getActive().getSheetByName(SH_MASTER);
  const out = {};
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 5).getDisplayValues().forEach(r => { out[r[0]] = { code: r[0], name: r[1], barcode: r[2], unit: r[3], price: num_(r[4]) }; });
  return out;
}

/** STOCK เต็ม (รหัส → {name, unit, min, max, qty}) */
function readStockFull_() {
  const ck = 'b_stock_' + baseVersion_(), hit = cacheGetJson_(ck);
  if (hit) return hit;
  const out = readStockFullRaw_();
  cachePutJson_(ck, out, 21600);
  return out;
}
function readStockFullRaw_() {
  const s = SpreadsheetApp.getActive().getSheetByName(SH.STOCK);
  const out = {};
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 6).getValues().forEach(r => {
    const code = String(r[0]).trim().toUpperCase();
    if (code) out[code] = { name: String(r[1]), unit: String(r[2]), min: Number(r[3]) || 0, max: Number(r[4]) || 0, qty: Number(r[5]) || 0 };
  });
  return out;
}

/* =========================================================
 * รายงานสั่งซื้อ — รหัสกลาง = รหัสบริษัท (A01-223-02 → 223), หมวด = ส่วนแรก (A01)
 * ========================================================= */
const SH_CATOWN = 'หมวด_ผู้ดูแล', SH_SUP = 'บริษัท', SH_ORDER = 'สั่งซื้อ_LOG';
function catOf_(code) { const p = String(code).split('-'); return p[0] === '2' ? '2' : p[0]; }
function supOf_(code) { const p = String(code).split('-'); return p.length >= 3 ? p[1] : ''; }
function allCats_() { const a = ['2']; for (let i = 1; i <= 17; i++) a.push('A' + pad_(i)); return a; }

function getPurchaseConfig_() {
  const co = ensureSheet_(SH_CATOWN, ['หมวด', 'ผู้ดูแล']);
  const owners = {};
  if (co.getLastRow() > 1) co.getRange(2, 1, co.getLastRow() - 1, 2).getDisplayValues().forEach(r => { if (r[0]) owners[r[0]] = r[1]; });
  const sp = ensureSheet_(SH_SUP, ['รหัสบริษัท', 'ชื่อบริษัท', 'ผู้ดูแล (แทนหมวด)', 'หมายเหตุ']);
  const sups = {};
  if (sp.getLastRow() > 1) sp.getRange(2, 1, sp.getLastRow() - 1, 4).getDisplayValues().forEach(r => { if (r[0]) sups[r[0]] = { name: r[1], owner: r[2], note: r[3] }; });
  const stock = readStockFull_(), found = {};
  Object.keys(stock).forEach(c => { if (PRODUCT_RE.test(c)) { const s = supOf_(c); if (s) found[s] = (found[s] || 0) + 1; } });
  return { cats: allCats_(), owners: owners, sups: sups, supCodes: Object.keys(found).sort(), supCount: found,
    staff: getAllStaffRows_().map(s => s.nick ? s.nick + ' (' + s.name + ')' : s.name) };
}

function savePurchaseConfig_(owners, sups) {
  bumpVersion_();
  const co = ensureSheet_(SH_CATOWN, ['หมวด', 'ผู้ดูแล']);
  if (co.getLastRow() > 1) co.getRange(2, 1, co.getLastRow() - 1, 2).clearContent();
  const orows = Object.keys(owners || {}).filter(k => owners[k]).map(k => [k, owners[k]]);
  if (orows.length) { co.getRange(2, 1, orows.length, 1).setNumberFormat('@'); co.getRange(2, 1, orows.length, 2).setValues(orows); }
  const sp = ensureSheet_(SH_SUP, ['รหัสบริษัท', 'ชื่อบริษัท', 'ผู้ดูแล (แทนหมวด)', 'หมายเหตุ']);
  if (sp.getLastRow() > 1) sp.getRange(2, 1, sp.getLastRow() - 1, 4).clearContent();
  const srows = Object.keys(sups || {}).filter(k => sups[k] && (sups[k].name || sups[k].owner || sups[k].note)).map(k => [k, sups[k].name || '', sups[k].owner || '', sups[k].note || '']);
  if (srows.length) { sp.getRange(2, 1, srows.length, 1).setNumberFormat('@'); sp.getRange(2, 1, srows.length, 4).setValues(srows); }
  return getPurchaseConfig_();
}

/** ยอดขาย (จำนวน) ต่อรหัส จาก POS + ImedX ช่วงวันที่ */
function salesQty_(start, end) {
  const out = {}, dset = {};
  const pos = readRange_('POS', POS_HEADERS.length, start, end);
  pos.forEach(r => dset[r[0]] = 1);
  const voided = {};
  pos.forEach(r => { if (/^VOID-/i.test(r[2])) voided[r[2].replace(/^VOID-/i, '')] = 1; });
  pos.forEach(r => {
    if (/^VOID-/i.test(r[2]) || voided[r[2]] || /ยกเลิก/.test(r[9])) return;
    const c = String(r[5]).trim().toUpperCase(); out[c] = (out[c] || 0) + num_(r[7]);
  });
  readRange_('IMED', IMED_HEADERS.length, start, end).forEach(r => { dset[r[0]] = 1; const c = String(r[7]).trim().toUpperCase(); out[c] = (out[c] || 0) + num_(r[9]); });
  Object.defineProperty(out, '_days', { value: Object.keys(dset).length, enumerable: false });
  return out;
}

function lastOrders_() {
  const s = ensureSheet_(SH_ORDER, ['วันที่', 'รหัสบริษัท', 'สถานะ', 'ผู้บันทึก', 'หมายเหตุ', 'เวลา']);
  const out = {};
  if (s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 5).getDisplayValues().forEach(r => { out[r[1]] = { date: r[0], status: r[2], by: r[3], note: r[4] }; });
  return out;
}

function markOrdered_(sup, on, note) {
  bumpVersion_();
  const s = ensureSheet_(SH_ORDER, ['วันที่', 'รหัสบริษัท', 'สถานะ', 'ผู้บันทึก', 'หมายเหตุ', 'เวลา']);
  const r0 = s.getLastRow() + 1;
  s.getRange(r0, 1, 1, 2).setNumberFormat('@');
  s.getRange(r0, 1, 1, 6).setValues([[todayKey_(), String(sup), on ? 'สั่งแล้ว' : 'ยกเลิกสถานะ', CURRENT_USER ? CURRENT_USER.name : '', note || '', new Date()]]);
  return lastOrders_()[sup] || null;
}

/** days = ยอดขายย้อนหลังกี่วัน, cover = ต้องมีของพอขายกี่วัน */
function getPurchaseReport_(days, cover) {
  days = Math.max(7, Number(days) || 30); cover = Math.max(1, Number(cover) || 14);
  const end = todayKey_(), start = addDays_(end, -(days - 1));
  const stock = liveStock_();
  if (!Object.keys(stock).length) throw new Error('ยังไม่มีข้อมูล STOCK — นำเข้าไฟล์ TLP_14_09 ก่อน');
  const cfg = getPurchaseConfig_(), sold = salesQty_(start, end), master = readMaster_(), orders = lastOrders_();
  const items = [];
  const effDays = Math.max(1, Math.min(days, sold._days || days));   // ใช้จำนวนวันที่มีข้อมูลจริง
  let dead = 0;
  Object.keys(stock).forEach(code => {
    if (!PRODUCT_RE.test(code)) return;
    const st = stock[code], sq = sold[code] || 0;
    if (sq <= 0 && st.qty <= 0) { dead++; return; }          // ไม่มีของและไม่มียอดขาย = ไม่ต้องแสดง
    const avg = sq / effDays, cov = avg > 0 ? st.qty / avg : null;
    let status = 'ปกติ';
    if (st.qty <= 0) status = 'หมด';
    else if ((st.min > 0 && st.qty <= st.min) || (cov != null && cov < cover)) status = 'ใกล้หมด';
    const target = st.max > 0 ? st.max : Math.ceil(avg * cover * 2);
    const sup = supOf_(code), cat = catOf_(code), si = cfg.sups[sup] || {};
    items.push({ code: code, name: st.name || (master[code] || {}).name || '', unit: st.unit, cat: cat, sup: sup, supName: si.name || '',
      owner: si.owner || cfg.owners[cat] || '', qty: st.qty, sold: sq, avg: Math.round(avg * 100) / 100,
      cover: cov == null ? null : Math.round(cov), status: status, suggest: status === 'ปกติ' ? 0 : Math.max(0, target - st.qty),
      barcode: (master[code] || {}).barcode || '' });
  });
  const ordMap = {};
  Object.keys(orders).forEach(k => { ordMap[k] = orders[k]; });
  return { items: items, dead: dead, days: days, effDays: effDays, cover: cover, start: thaiShort_(start), end: thaiShort_(end),
    stockAsOf: PropertiesService.getScriptProperties().getProperty('stockAsOf') || '', orders: ordMap, owners: cfg.owners, sups: cfg.sups,
    hasSales: Object.keys(sold).length > 0 };
}

/* =========================================================
 * ตู้เก็บสินค้า / หน้าเชลฟ์ / แผนเติมของ
 *   ตู้        : ID | ชื่อตู้ | ที่ตั้ง | รูป | หมายเหตุ
 *   ตู้_สินค้า : ID ตู้ | รหัส | จำนวน | อัปเดต | โดย
 *   หน้าเชลฟ์  : รหัส | วางหน้าเชลฟ์ได้ (ชิ้น) | รูปสินค้า | หมายเหตุ
 *   หลักคิด    : STOCK ในระบบ (TLP_14_09) = หน้าเชลฟ์ + ในตู้  →  หน้าเชลฟ์ = STOCK − ในตู้
 *               ขายตัดหน้าเชลฟ์ก่อน ถ้าขายเกินหน้าเชลฟ์ แปลว่าหยิบจากตู้ไปขายแล้ว (ปรับยอดตู้)
 * ========================================================= */
const SH_CAB = 'ตู้', SH_CABI = 'ตู้_สินค้า', SH_SHELF = 'หน้าเชลฟ์';
function cabSheet_() { return ensureSheet_(SH_CAB, ['ID', 'ชื่อตู้', 'ที่ตั้ง', 'รูปตู้', 'หมายเหตุ']); }
function cabItemSheet_() { return ensureSheet_(SH_CABI, ['ID ตู้', 'รหัสสินค้า', 'จำนวน', 'อัปเดต', 'โดย']); }
function shelfSheet_() { return ensureSheet_(SH_SHELF, ['รหัสสินค้า', 'วางหน้าเชลฟ์ได้ (ชิ้น)', 'รูปสินค้า', 'หมายเหตุ']); }

function readCabs_() {
  const s = cabSheet_();
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 5).getDisplayValues().map((r, i) => ({ row: i + 2, id: r[0], name: r[1], loc: r[2], photo: r[3], note: r[4] })).filter(x => x.id);
}
function readCabItems_() {
  const s = cabItemSheet_();
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 5).getDisplayValues().map((r, i) => ({ row: i + 2, cab: r[0], code: r[1].toUpperCase(), qty: num_(r[2]), upd: r[3], by: r[4] })).filter(x => x.cab && x.code);
}
function writeCabItems_(items) {
  const s = cabItemSheet_();
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 5).clearContent();
  const rows = items.filter(x => x.qty > 0).map(x => [x.cab, x.code, x.qty, x.upd || '', x.by || '']);
  if (rows.length) { s.getRange(2, 1, rows.length, 2).setNumberFormat('@'); s.getRange(2, 1, rows.length, 5).setValues(rows); }
}
function readShelf_() {
  const s = shelfSheet_(), out = {};
  if (s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 4).getDisplayValues().forEach(r => { if (r[0]) out[r[0].toUpperCase()] = { cap: num_(r[1]), img: r[2], note: r[3] }; });
  return out;
}
function webUrl_() {
  // ถ้าตั้ง "ลิงก์หน้าเว็บ (QR)" ในชีตตั้งค่า เช่น https://ชื่อ.github.io/petshop/ จะใช้ลิงก์นั้นใน QR ตู้
  try { const v = sheet_(SH.SET).getRange('A2:B60').getValues().filter(r => String(r[0]).trim() === SET_KEYS.qrUrl)[0]; if (v && String(v[1]).trim()) return String(v[1]).trim().replace(/\/?$/, '/'); } catch (e) {}
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}
function nowText_() { return Utilities.formatDate(new Date(), TZ, 'd/M/yyyy HH:mm'); }

function getCabinets_() {
  const items = readCabItems_(), cnt = {};
  items.forEach(x => { const c = cnt[x.cab] = cnt[x.cab] || { items: 0, qty: 0 }; c.items++; c.qty += x.qty; });
  return { url: webUrl_(), cabs: readCabs_().map(c => Object.assign(c, cnt[c.id] || { items: 0, qty: 0 })) };
}

function saveCabinet_(cab) {
  bumpVersion_();
  const s = cabSheet_(), cabs = readCabs_();
  let id = String(cab.id || '').trim().toUpperCase();
  if (!id) { let n = cabs.length + 1; do { id = 'C' + pad_(n++); } while (cabs.some(c => c.id === id)); }
  const row = [id, String(cab.name || id), String(cab.loc || ''), String(cab.photo || ''), String(cab.note || '')];
  const ex = cabs.filter(c => c.id === id)[0];
  if (ex) s.getRange(ex.row, 1, 1, 5).setValues([row]);
  else { const r0 = s.getLastRow() + 1; s.getRange(r0, 1).setNumberFormat('@'); s.getRange(r0, 1, 1, 5).setValues([row]); }
  return getCabinet_(id);
}

function deleteCabinet_(id) {
  bumpVersion_();
  const ex = readCabs_().filter(c => c.id === id)[0];
  if (!ex) throw new Error('ไม่พบตู้ ' + id);
  if (readCabItems_().some(x => x.cab === id)) throw new Error('ตู้นี้ยังมีสินค้าอยู่ ย้ายหรือนำออกให้หมดก่อน');
  cabSheet_().deleteRow(ex.row);
  return getCabinets_();
}

function productInfo_(code, master, stock, shelf) {
  const m = master[code] || {}, st = stock[code] || {}, sh = shelf[code] || {};
  return { code: code, name: m.name || st.name || '', barcode: m.barcode || '', unit: m.unit || st.unit || '', price: m.price || 0,
    stock: st.qty == null ? null : st.qty, snap: st.snap == null ? null : st.snap, soldSince: st.sold || 0, cap: sh.cap || 0, img: sh.img || '' };
}

function getCabinet_(id) {
  const cab = readCabs_().filter(c => c.id === id)[0];
  if (!cab) throw new Error('ไม่พบตู้ ' + id);
  const master = readMaster_(), stock = liveStock_(), shelf = readShelf_();
  const all = readCabItems_(), inCab = {};
  all.forEach(x => inCab[x.code] = (inCab[x.code] || 0) + x.qty);
  const items = all.filter(x => x.cab === id).map(x => {
    const p = productInfo_(x.code, master, stock, shelf), now = p.stock == null ? null : Math.max(0, Math.floor(p.stock));
    const totalCab = inCab[x.code], effAll = now == null ? totalCab : Math.min(totalCab, now);
    return Object.assign(p, { qty: x.qty, upd: x.upd, by: x.by, inCabAll: totalCab, shelfNow: now == null ? null : now - effAll, over: now != null && totalCab > now ? totalCab - now : 0 });
  });
  items.sort((a, b) => a.name.localeCompare(b.name));
  return { cab: cab, items: items, url: webUrl_(), stockInfo: stockInfo_() };
}

/** ค้นหาสินค้าด้วย Barcode หรือรหัสสินค้า */
function lookupProduct_(q) {
  q = String(q || '').trim();
  if (!q) throw new Error('กรุณาสแกนหรือพิมพ์ Barcode / รหัสสินค้า');
  const master = readMaster_(), stock = liveStock_(), shelf = readShelf_();
  const bc = cleanBarcode_(q), up = q.toUpperCase();
  let code = master[up] || stock[up] ? up : '';
  if (!code) { const k = Object.keys(master).filter(c => master[c].barcode && (master[c].barcode === bc || master[c].barcode.replace(/^0+/, '') === bc.replace(/^0+/, '')))[0]; if (k) code = k; }
  // พิมพ์ตัวเลขท้าย Barcode / ท้ายรหัสสินค้า (เช่น 5 ตัวท้าย)
  if (!code && /^\d{3,7}$/.test(bc)) {
    const cand = {};
    Object.keys(master).forEach(c => { if ((master[c].barcode && master[c].barcode.slice(-bc.length) === bc) || c.slice(-bc.length) === bc) cand[c] = 1; });
    Object.keys(stock).forEach(c => { if (c.slice(-bc.length) === bc) cand[c] = 1; });
    const list = Object.keys(cand);
    if (list.length === 1) code = list[0];
    else if (list.length > 1) return { multi: list.slice(0, 20).map(c => productInfo_(c, master, stock, shelf)), q: q };
  }
  if (!code) throw new Error('ไม่พบสินค้า "' + q + '" — ลองพิมพ์ Barcode เต็ม, 5 ตัวท้าย หรือรหัสสินค้า');
  const p = productInfo_(code, master, stock, shelf);
  p.where = readCabItems_().filter(x => x.code === code).map(x => ({ cab: x.cab, qty: x.qty }));
  const inCab = p.where.reduce((a, w) => a + w.qty, 0);
  p.inCab = inCab;
  if (p.stock != null) { const now = Math.max(0, Math.floor(p.stock)); p.shelfNow = now - Math.min(inCab, now); p.over = Math.max(0, inCab - now); }
  p.stockInfo = stockInfo_();
  return p;
}

function searchProducts_(q) {
  q = String(q || '').trim().toLowerCase();
  if (q.length < 2) return [];
  const master = readMaster_(), stock = liveStock_(), shelf = readShelf_();
  const codes = {};
  Object.keys(master).forEach(c => codes[c] = 1);
  Object.keys(stock).forEach(c => { if (PRODUCT_RE.test(c)) codes[c] = 1; });
  const tail = c => /^\d+$/.test(q) && (((master[c] || {}).barcode || '').slice(-q.length) === q || c.slice(-q.length) === q.toUpperCase());
  return Object.keys(codes).filter(c => (c + ' ' + ((master[c] || {}).name || (stock[c] || {}).name || '') + ' ' + ((master[c] || {}).barcode || '')).toLowerCase().indexOf(q) >= 0)
    .sort((a, b) => tail(b) - tail(a))
    .slice(0, 30).map(c => productInfo_(c, master, stock, shelf));
}

/** สแกนเข้า/ออก/ย้ายตู้: delta บวก = เข้าตู้, ลบ = ออกจากตู้; toCab = ย้ายไปตู้อื่น */
function cabMove_(cabId, code, delta, toCab, setQty) {
  bumpVersion_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    code = String(code).toUpperCase(); delta = Number(delta) || 0;
    if (!readCabs_().some(c => c.id === cabId)) throw new Error('ไม่พบตู้ ' + cabId);
    if (toCab && !readCabs_().some(c => c.id === toCab)) throw new Error('ไม่พบตู้ ' + toCab);
    const items = readCabItems_(), by = CURRENT_USER ? CURRENT_USER.name : '', t = nowText_();
    const get = cab => { let x = items.filter(i => i.cab === cab && i.code === code)[0]; if (!x) { x = { cab: cab, code: code, qty: 0 }; items.push(x); } return x; };
    const x = get(cabId);
    if (setQty != null && setQty !== '') { x.qty = Math.max(0, Number(setQty) || 0); }
    else {
      if (x.qty + delta < 0) throw new Error('ในตู้มีแค่ ' + x.qty + ' ชิ้น');
      x.qty += delta;
      if (toCab) { const y = get(toCab); y.qty += -delta; y.upd = t; y.by = by; }
    }
    x.upd = t; x.by = by;
    writeCabItems_(items);
    return getCabinet_(cabId);
  } finally { lock.releaseLock(); }
}

function getShelfList_() {
  const master = readMaster_(), stock = liveStock_(), shelf = readShelf_();
  return Object.keys(shelf).map(c => productInfo_(c, master, stock, shelf)).sort((a, b) => a.name.localeCompare(b.name));
}

function saveShelf_(code, cap, img) {
  bumpVersion_();
  code = String(code).toUpperCase();
  const s = shelfSheet_(), n = s.getLastRow() - 1;
  const codes = n > 0 ? s.getRange(2, 1, n, 1).getDisplayValues().map(r => r[0].toUpperCase()) : [];
  const i = codes.indexOf(code);
  const row = [code, Number(cap) || 0, String(img || ''), ''];
  if (i >= 0) s.getRange(i + 2, 1, 1, 3).setValues([row.slice(0, 3)]);
  else { const r0 = s.getLastRow() + 1; s.getRange(r0, 1).setNumberFormat('@'); s.getRange(r0, 1, 1, 4).setValues([row]); }
  return lookupProduct_(code);
}

/** pure: แผนเติมของ */
function buildRefill_(stock, cabItems, shelf, cabs) {
  const byCode = {};
  cabItems.forEach(x => (byCode[x.code] = byCode[x.code] || []).push(x));
  const order = {};
  cabs.forEach((c, i) => order[c.id] = i);
  const codes = {};
  Object.keys(shelf).forEach(c => { if (shelf[c].cap > 0) codes[c] = 1; });
  Object.keys(byCode).forEach(c => codes[c] = 1);
  const picks = [], over = [], noCab = [], shelfRows = [];
  Object.keys(codes).forEach(code => {
    const T = Math.max(0, Math.floor((stock[code] || {}).qty || 0));
    const list = (byCode[code] || []).slice().sort((a, b) => (order[a.cab] || 0) - (order[b.cab] || 0));
    const C = list.reduce((a, x) => a + x.qty, 0), cap = (shelf[code] || {}).cap || 0;
    const S = Math.max(0, T - C);
    if (C > T) over.push({ code: code, stock: T, inCab: C, diff: C - T, cabs: list.map(x => ({ cab: x.cab, qty: x.qty })) });
    shelfRows.push({ code: code, stock: T, inCab: C, shelf: S, cap: cap });
    if (cap <= 0) return;
    let need = Math.max(0, cap - S);
    if (!need) return;
    if (C <= 0) { noCab.push({ code: code, stock: T, shelf: S, cap: cap, need: need }); return; }
    const avail = Math.min(C, T);    // หยิบได้ไม่เกินของที่มีจริงในระบบ
    let left = Math.min(need, avail);
    list.forEach(x => { if (left <= 0 || x.qty <= 0) return; const take = Math.min(left, x.qty); left -= take; picks.push({ cab: x.cab, code: code, take: take, cabQty: x.qty, shelf: S, cap: cap }); });
  });
  return { picks: picks, over: over, noCab: noCab, shelf: shelfRows };
}

function getRefillPlan_() {
  const master = readMaster_(), stock = liveStock_(), shelf = readShelf_(), cabs = readCabs_();
  const r = buildRefill_(stock, readCabItems_(), shelf, cabs);
  const info = c => productInfo_(c, master, stock, shelf);
  const groups = cabs.map(c => ({ cab: c, items: r.picks.filter(p => p.cab === c.id).map(p => Object.assign(info(p.code), p)) })).filter(g => g.items.length);
  return { groups: groups, over: r.over.map(o => Object.assign(info(o.code), o)), noCab: r.noCab.map(o => Object.assign(info(o.code), o)),
    stockAsOf: PropertiesService.getScriptProperties().getProperty('stockAsOf') || '', stockInfo: stockInfo_() };
}

/** ยืนยันเติมของ: หักจำนวนที่หยิบออกจากตู้ (ของไปอยู่หน้าเชลฟ์) */
function applyRefill_(picks) {
  bumpVersion_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const items = readCabItems_(), by = CURRENT_USER ? CURRENT_USER.name : '', t = nowText_();
    let n = 0;
    (picks || []).forEach(p => {
      const x = items.filter(i => i.cab === p.cab && i.code === String(p.code).toUpperCase())[0];
      if (!x) return;
      x.qty = Math.max(0, x.qty - (Number(p.take) || 0)); x.upd = t; x.by = by; n++;
    });
    writeCabItems_(items);
    return 'บันทึกการเติมของ ' + n + ' รายการแล้ว';
  } finally { lock.releaseLock(); }
}

/** ปรับยอดในตู้ให้ไม่เกิน STOCK ในระบบ (กรณีขายเกินหน้าเชลฟ์ = หยิบจากตู้ไปขาย) — ตัดจากตู้ท้าย ๆ ก่อน */
function syncCabinets_(codes) {
  bumpVersion_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const stock = liveStock_(), items = readCabItems_(), cabs = readCabs_(), order = {}, t = nowText_(), by = CURRENT_USER ? CURRENT_USER.name : '';
    cabs.forEach((c, i) => order[c.id] = i);
    const only = codes && codes.length ? codes.map(c => String(c).toUpperCase()) : null;
    let n = 0;
    const byCode = {};
    items.forEach(x => (byCode[x.code] = byCode[x.code] || []).push(x));
    Object.keys(byCode).forEach(code => {
      if (only && only.indexOf(code) < 0) return;
      const T = Math.max(0, Math.floor((stock[code] || {}).qty || 0));
      const list = byCode[code].sort((a, b) => (order[b.cab] || 0) - (order[a.cab] || 0));
      let over = list.reduce((a, x) => a + x.qty, 0) - T;
      list.forEach(x => { if (over <= 0) return; const cut = Math.min(over, x.qty); x.qty -= cut; over -= cut; x.upd = t; x.by = by + ' (ปรับตามระบบ)'; n++; });
    });
    writeCabItems_(items);
    return 'ปรับยอดในตู้ให้ตรงระบบแล้ว ' + n + ' รายการ';
  } finally { lock.releaseLock(); }
}

/** อัปโหลดรูป (ตู้/สินค้า) ไป Google Drive แล้วคืนลิงก์แสดงรูป */
function uploadPhoto_(base64, mime, name) {
  const it = DriveApp.getFoldersByName('PetShop รูปตู้และสินค้า');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('PetShop รูปตู้และสินค้า');
  const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), mime || 'image/jpeg', name || ('photo_' + Date.now() + '.jpg')));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1000';
}

/* =========================================================
 * ติดตามลูกค้า (โทร / เลื่อน / มารับ / ทำนัดใน ImedX แล้ว)
 *   ติดตาม : เวลา | key | ลูกค้า | HN | สถานะ | วันที่ (นัด/มารับ/โทรซ้ำ) | หมายเหตุ | ผู้ติดตาม
 * ========================================================= */
const SH_FOLLOW = 'ติดตาม';
const FOLLOW_HEADERS = ['เวลา', 'key', 'ลูกค้า', 'HN', 'สถานะ', 'วันที่', 'หมายเหตุ', 'ผู้ติดตาม'];
function followSheet_() { return ensureSheet_(SH_FOLLOW, FOLLOW_HEADERS); }

function readFollows_() {
  const s = followSheet_();
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 8).getDisplayValues().map(r => ({
    at: r[0], key: r[1], name: r[2], hn: r[3], status: r[4], date: r[5], note: r[6], by: r[7]
  })).filter(x => x.key);
}

/** สถานะล่าสุดต่อลูกค้า */
function latestFollows_() {
  const out = {};
  readFollows_().forEach(x => {
    const o = out[x.key] || (out[x.key] = { count: 0 });
    o.count++;
    Object.assign(o, { at: x.at, status: x.status, date: x.date, note: x.note, by: x.by });
  });
  return out;
}

function saveFollow_(f) {
  f = f || {};
  const key = normName_(f.key || f.name);
  if (!key) throw new Error('ไม่พบชื่อลูกค้า');
  if (!f.status) throw new Error('กรุณาเลือกสถานะ');
  bumpVersion_();
  const s = followSheet_(), r0 = s.getLastRow() + 1;
  s.getRange(r0, 1, 1, 8).setNumberFormat('@');
  s.getRange(r0, 1, 1, 8).setValues([[
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), key, String(f.name || ''), String(f.hn || ''),
    String(f.status), String(f.date || ''), String(f.note || ''), CURRENT_USER ? CURRENT_USER.name : ''
  ]]);
  return getFollowHistory_(key);
}

function getFollowHistory_(key) {
  key = normName_(key);
  const list = readFollows_().filter(x => x.key === key).reverse();
  return { key: key, list: list, latest: list[0] || null, count: list.length };
}

/* =========================================================
 * STOCK ปัจจุบัน = STOCK ณ เวลาอัปเดต (TLP_14_09) − ยอดขายหลังเวลานั้น (POS + ImedX)
 * ========================================================= */
function stockInfo_() {
  const pr = PropertiesService.getScriptProperties();
  const key = pr.getProperty('stockAsOfKey') || '';
  return { asOf: pr.getProperty('stockAsOf') || '', asOfKey: key, until: lastSaleTime_() };
}

/** เวลาขายล่าสุดที่มีข้อมูลในระบบ */
function lastSaleTime_() {
  const mk = todayKey_().slice(0, 7);
  let last = '';
  ['POS', 'IMED'].forEach(p => readMonthly_(p, mk, 2).forEach(r => { const t = r[0] + ' ' + r[1]; if (t > last) last = t; }));
  return last;
}

/** ยอดขายต่อรหัส หลังเวลา sinceKey ('yyyy-MM-dd HH:mm') */
function soldSince_(sinceKey) {
  const out = {};
  if (!sinceKey) return out;
  const start = sinceKey.slice(0, 10), end = todayKey_();
  if (start > end) return out;
  const after = r => (r[0] + ' ' + String(r[1]).replace('.', ':')) > sinceKey;
  const pos = readRange_('POS', POS_HEADERS.length, start, end);
  const voided = {};
  pos.forEach(r => { if (/^VOID-/i.test(r[2])) voided[r[2].replace(/^VOID-/i, '')] = 1; });
  pos.forEach(r => {
    if (!after(r) || /^VOID-/i.test(r[2]) || voided[r[2]] || /ยกเลิก/.test(r[9])) return;
    const c = String(r[5]).trim().toUpperCase(); out[c] = (out[c] || 0) + num_(r[7]);
  });
  readRange_('IMED', IMED_HEADERS.length, start, end).forEach(r => {
    if (!after(r)) return;
    const c = String(r[7]).trim().toUpperCase(); out[c] = (out[c] || 0) + num_(r[9]);
  });
  return out;
}

/** เหมือน readStockFull_ แต่ qty = คงเหลือตอนนี้ (snap = ณ เวลาอัปเดต, sold = ขายหลังจากนั้น) */
function liveStock_() {
  const ck = 'b_live_' + baseVersion_() + '_' + getVersion_(), hit = cacheGetJson_(ck);
  if (hit) return hit;
  const base = readStockFull_(), sold = soldSince_(stockInfo_().asOfKey), out = {};
  Object.keys(base).forEach(c => {
    const b = base[c], sd = sold[c] || 0;
    out[c] = { name: b.name, unit: b.unit, min: b.min, max: b.max, snap: b.qty, sold: sd, qty: b.qty - sd };
  });
  cachePutJson_(ck, out, 1800);
  return out;
}

/* =========================================================
 * ล้างข้อมูลรหัสฝั่งโรงพยาบาล (เก็บเฉพาะ 2-, A01–A17) — รันจากเมนู "ระบบรายงาน" หรือหน้า Apps Script
 * ========================================================= */
function cleanupNonProduct() {
  const ss = SpreadsheetApp.getActive();
  const report = [];
  const clean = (name, col) => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const n = sh.getLastRow() - 1, w = sh.getLastColumn();
    const vals = sh.getRange(2, 1, n, w).getDisplayValues();
    const keep = vals.filter(r => PRODUCT_RE.test(String(r[col]).trim()));
    if (keep.length === n) return;
    sh.getRange(2, 1, n, w).clearContent();
    if (keep.length) { sh.getRange(2, 1, keep.length, w).setNumberFormat('@'); sh.getRange(2, 1, keep.length, w).setValues(keep); }
    if (sh.getMaxRows() > keep.length + 50) sh.deleteRows(keep.length + 52, sh.getMaxRows() - keep.length - 51);
    report.push(name + ': ลบ ' + (n - keep.length) + ' แถว เหลือ ' + keep.length);
  };
  clean(SH.STOCK, 0);
  clean('MASTER_สินค้า', 0);
  ss.getSheets().forEach(sh => { if (/^IMED_\d{4}-\d{2}$/.test(sh.getName())) clean(sh.getName(), 7); });
  // ลดแถวว่างของทุกชีต (ประหยัดพื้นที่ 10 ล้านเซลล์)
  ss.getSheets().forEach(sh => {
    const last = Math.max(sh.getLastRow(), 1), max = sh.getMaxRows();
    if (max > last + 200) sh.deleteRows(last + 101, max - last - 100);
    const lc = Math.max(sh.getLastColumn(), 1), mc = sh.getMaxColumns();
    if (mc > lc + 5) sh.deleteColumns(lc + 3, mc - lc - 2);
  });
  bumpBase_(); bumpVersion_();
  const msg = report.length ? report.join('\n') : 'ไม่มีรหัสฝั่งโรงพยาบาลค้างอยู่แล้ว';
  try { ss.toast(msg, 'ล้างข้อมูลเสร็จ', 10); } catch (e) {}
  Logger.log(msg);
  return msg;
}

/** ดูว่าใช้เซลล์ไปเท่าไหร่ (จาก 10 ล้าน) — รันจากเมนู */
function showUsage() {
  const ss = SpreadsheetApp.getActive();
  let total = 0;
  const rows = ss.getSheets().map(sh => { const c = sh.getMaxRows() * sh.getMaxColumns(); total += c; return [sh.getName(), c]; })
    .sort((a, b) => b[1] - a[1]).slice(0, 8).map(r => r[0] + ': ' + r[1].toLocaleString()).join('\n');
  const msg = 'ใช้ไป ' + total.toLocaleString() + ' / 10,000,000 เซลล์ (' + (total / 1e5).toFixed(1) + '%)\n\nชีตที่ใหญ่ที่สุด:\n' + rows;
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
