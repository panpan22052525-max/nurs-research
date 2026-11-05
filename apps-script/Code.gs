/**
 * ระบบเสนอโครงร่างงานวิจัย คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏศรีสะเกษ
 * Backend: Google Apps Script + Google Sheets
 *
 * วิธีติดตั้ง
 * 1) เปิด script.google.com > โปรเจกต์ใหม่ (หรือเปิดจากชีต: ส่วนขยาย > Apps Script)
 * 2) วางโค้ดนี้ในไฟล์ Code.gs และสร้างไฟล์ HTML ชื่อ Index แล้ววางเนื้อหา Index.html
 * 3) รันฟังก์ชัน setup() หนึ่งครั้ง (อนุญาตสิทธิ์) เพื่อสร้างชีตและหัวตารางใน Google Sheets ตาม SHEET_ID
 * 4) ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป
 *    - เรียกใช้ในฐานะ: ฉัน
 *    - ผู้มีสิทธิ์เข้าถึง: ทุกคน  (จำเป็น เพื่อให้หน้าเว็บบน GitHub Pages ส่งข้อมูลเข้ามาได้)
 * 5) คัดลอก URL ที่ลงท้ายด้วย /exec ไปวางในไฟล์ config.js บน GitHub
 *
 * ไฟล์นี้ทำหน้าที่เป็น API หลังบ้านเท่านั้น ผู้ใช้จะเห็นแต่หน้าเว็บบน GitHub Pages
 */

/** Google Sheets ที่ใช้เก็บข้อมูล (ของคณะพยาบาลศาสตร์) */
const SHEET_ID = '1BZTbYXSFAPv2sZ6Tv41FwMN496FNinTpTv8CR2FRLp0';

const SHEETS = {
  proposals: {
    name: 'โครงร่างวิจัย',
    headers: ['เลขที่','วันเวลาที่ส่ง','สถานะ','ปีงบประมาณ','ชื่อโครงการ (ไทย)','ชื่อโครงการ (อังกฤษ)',
      'แหล่งทุน','ประเภทโครงการ','ประเภทการวิจัย','สาขา','กลุ่มเป้าหมาย','คำสำคัญ',
      'คำนำหน้า','หัวหน้าโครงการ','สัดส่วนหัวหน้า (%)','สาขาวิชา','โทรศัพท์','อีเมล','ผู้ร่วมวิจัย',
      'ความเป็นมา','วัตถุประสงค์','กรอบแนวคิด','รูปแบบการวิจัย','พื้นที่','กลุ่มตัวอย่าง','ขนาดตัวอย่าง',
      'เครื่องมือ/การวิเคราะห์','เดือนเริ่ม','ระยะเวลา (เดือน)','แผนการดำเนินงาน',
      'ยุทธศาสตร์','SDGs','การลดความเหลื่อมล้ำทางสุขภาพ','ภาวะสุขภาพที่มุ่งแก้ไข','กลุ่มเปราะบางที่ได้รับประโยชน์','แนวทางลดความเหลื่อมล้ำ','คำอธิบายการลดความเหลื่อมล้ำ','เป้าหมายการเผยแพร่','ประโยชน์ที่คาดว่าจะได้รับ','จริยธรรมการวิจัย',
      'ค่าตอบแทน','ค่าใช้สอย','ค่าวัสดุ','งบประมาณรวม','ลิงก์โครงร่างฉบับสมบูรณ์','เอกสารแนบทั้งหมด','ผู้ส่ง (บัญชี Google)']
  },
  members: { name: 'คณะผู้วิจัย', headers: ['เลขที่','บทบาท','ชื่อ–นามสกุล','สังกัด','สัดส่วน (%)'] },
  budget:  { name: 'งบประมาณ', headers: ['เลขที่','หมวด','รายการ','จำนวนเงิน (บาท)'] },
  faculty: { name: 'รายชื่ออาจารย์', headers: ['สาขาวิชา','คำนำหน้า / ตำแหน่ง','ชื่อ–นามสกุล','อีเมล','โทรศัพท์'] },
  drafts:  { name: 'ร่างที่บันทึกไว้', headers: ['อีเมลหัวหน้าโครงการ','หัวหน้าโครงการ','ชื่อโครงการ','บันทึกล่าสุด','จำนวนลิงก์เอกสาร','ลิงก์เอกสารแนบ','ข้อมูลร่าง (JSON)'] },
  files:   { name: 'เอกสารแนบ', headers: ['เลขที่','ประเภทเอกสาร','ชื่อไฟล์ใน Drive','ลิงก์','คณะเปิดอ่านได้หรือไม่','ตรวจเมื่อ'] }
};

/** API สำหรับหน้าเว็บบน GitHub Pages */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const d = req.data || {};
    let result;
    switch (req.action) {
      case 'submit':    result = submitProposal(d); break;
      case 'saveDraft': result = saveDraft(d); break;
      case 'faculty':   result = getFaculty(); break;
      case 'checkLink': result = checkDriveLink(String(d.url || '')); break;
      default: throw new Error('ไม่รู้จักคำสั่ง ' + req.action);
    }
    out = { ok: true, result: result };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** เปิด URL /exec ตรง ๆ เพื่อทดสอบว่า API ทำงาน */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'NURS SSKRU research proposal API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGetHtml_() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบเสนอโครงร่างงานวิจัย | คณะพยาบาลศาสตร์ มรภ.ศรีสะเกษ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setup() {
  Object.values(SHEETS).forEach(cfg => getSheet_(cfg));
}

function getSheet_(cfg) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(cfg.name);
  if (!sh) sh = ss.insertSheet(cfg.name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(cfg.headers);
    sh.getRange(1, 1, 1, cfg.headers.length)
      .setFontWeight('bold').setBackground('#1B6B3A').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else {
    const have = sh.getRange(1, 1, 1, cfg.headers.length).getDisplayValues()[0];
    if (have.join('|') !== cfg.headers.join('|')) {
      if (sh.getLastRow() === 1) {          // มีแต่หัวตาราง ยังไม่มีข้อมูล → อัปเดตหัวตารางให้อัตโนมัติ
        sh.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers])
          .setFontWeight('bold').setBackground('#1B6B3A').setFontColor('#ffffff');
      } else {
        throw new Error(`หัวตารางในแท็บ "${cfg.name}" ไม่ตรงกับรุ่นปัจจุบัน กรุณาเปลี่ยนชื่อแท็บเดิม แล้วรัน setup อีกครั้ง`);
      }
    }
  }
  return sh;
}

function submitProposal(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEETS.proposals);
    const ref = nextRef_(sh, d.fy);
    const sum = k => (d.budget[k] || []).reduce((a, r) => a + (Number(r.a) || 0), 0);
    const join = v => [].concat(v || []).join(', ');
    let user = '';
    try { user = Session.getActiveUser().getEmail(); } catch (e) {}

    sh.appendRow([
      ref, new Date(), 'รอพิจารณา', d.fy, d.titleTh, d.titleEn,
      d.fund, d.projType, d.resType, d.field, d.target, join(d.keywords),
      d.leadPrefix, d.leadName, Number(d.leadShare) || 0, d.leadDept, "'" + (d.leadTel || ''), d.leadEmail,
      (d.members || []).map(m => `${m.name} (${m.share || 0}%)`).join('; '),
      d.background, (d.objectives || []).map((o, i) => `${i + 1}) ${o}`).join('\n'), d.framework,
      d.design, d.area, d.sample, d.n, d.tools, d.start, d.months,
      (d.activities || []).map(a => `${a.name}: เดือน ${a.months.sort((x, y) => x - y).join(',')}`).join('\n'),
      d.strategy, join(d.sdgs), d.equityLevel, join(d.healthIssues), join(d.equityGroups), join(d.equityDims), d.equityDesc, join(d.outputs), d.benefit, d.ethics,
      sum('comp'), sum('use'), sum('mat'), Number(d.total) || 0,
      ((d.links || [])[0] || {}).url || '',
      (d.links || []).map(l => `${l.type}: ${l.url}`).join('\n'),
      user
    ]);

    const ms = getSheet_(SHEETS.members);
    ms.appendRow([ref, 'หัวหน้าโครงการ', `${d.leadPrefix} ${d.leadName}`, 'คณะพยาบาลศาสตร์', Number(d.leadShare) || 0]);
    (d.members || []).forEach(m => ms.appendRow([ref, 'ผู้ร่วมวิจัย', m.name, m.aff, Number(m.share) || 0]));

    const bs = getSheet_(SHEETS.budget);
    const CAT = { comp: 'ค่าตอบแทน', use: 'ค่าใช้สอย', mat: 'ค่าวัสดุ' };
    Object.keys(CAT).forEach(k => (d.budget[k] || [])
      .filter(r => r.n || Number(r.a))
      .forEach(r => bs.appendRow([ref, CAT[k], r.n, Number(r.a) || 0])));

    const fs = getSheet_(SHEETS.files);
    (d.links || []).forEach(l => {
      const c = checkDriveLink(l.url);
      fs.appendRow([ref, l.type, c.name || l.name || '', l.url,
        c.ok ? 'เปิดได้' : (c.drive ? 'ยังไม่ได้แชร์สิทธิ์ให้คณะ' : 'ลิงก์ภายนอก (ตรวจเอง)'), new Date()]);
    });

    notify_(ref, d);
    return ref;
  } finally {
    lock.releaseLock();
  }
}

function nextRef_(sh, fy) {
  const prefix = `NUR-${fy}-`;
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
  const max = ids.filter(v => String(v).startsWith(prefix))
    .map(v => Number(String(v).slice(prefix.length)) || 0)
    .reduce((a, b) => Math.max(a, b), 0);
  return prefix + String(max + 1).padStart(3, '0');
}

/** ส่งอีเมลยืนยันให้ผู้เสนอ (ปิดได้โดยลบการเรียกใช้ใน submitProposal) */
function notify_(ref, d) {
  if (!d.leadEmail) return;
  try {
    MailApp.sendEmail({
      to: d.leadEmail,
      subject: `[คณะพยาบาลศาสตร์ มรภ.ศรีสะเกษ] ได้รับโครงร่างวิจัย ${ref}`,
      htmlBody: `<p>เรียน ${d.leadPrefix} ${d.leadName}</p>
        <p>งานวิจัย คณะพยาบาลศาสตร์ ได้รับโครงร่างวิจัยของท่านแล้ว</p>
        <p><b>เลขที่:</b> ${ref}<br><b>ชื่อโครงการ:</b> ${d.titleTh}<br>
        <b>งบประมาณที่เสนอ:</b> ${Number(d.total).toLocaleString('th-TH')} บาท</p>
        <p>ระบบจะแจ้งผลการพิจารณาให้ทราบต่อไป</p>`
    });
  } catch (e) { console.warn(e); }
}

/** ดึงรหัสไฟล์/โฟลเดอร์จากลิงก์ Google Drive / Docs / Sheets / Slides */
function driveId_(url) {
  const m = String(url).match(/\/(?:d|folders)\/([\w-]{15,})/) || String(url).match(/[?&]id=([\w-]{15,})/);
  return m ? m[1] : '';
}

/**
 * ตรวจว่าบัญชีที่ deploy (บัญชีคณะ) เปิดไฟล์จากลิงก์ได้หรือไม่ — เรียกจากหน้าเว็บตอนผู้วิจัยวางลิงก์
 * คืนค่า { ok, drive, name }
 */
function checkDriveLink(url) {
  const id = driveId_(url);
  if (!id || !/(drive|docs)\.google\.com/.test(url)) return { ok: false, drive: false, name: '' };
  try { return { ok: true, drive: true, name: DriveApp.getFileById(id).getName() }; } catch (e) {}
  try { return { ok: true, drive: true, name: '📁 ' + DriveApp.getFolderById(id).getName() }; } catch (e) {}
  return { ok: false, drive: true, name: '' };
}

/** ตรวจลิงก์ในชีต "เอกสารแนบ" ซ้ำทั้งหมด (ใช้เมื่อผู้วิจัยแจ้งว่าแก้การแชร์แล้ว) */
function recheckAllLinks() {
  const sh = getSheet_(SHEETS.files);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const rows = sh.getRange(2, 1, n, 6).getValues();
  rows.forEach((r, i) => {
    if (!r[3]) return;
    const c = checkDriveLink(r[3]);
    r[2] = c.name || r[2];
    r[4] = c.ok ? 'เปิดได้' : (c.drive ? 'ยังไม่ได้แชร์สิทธิ์ให้คณะ' : 'ลิงก์ภายนอก (ตรวจเอง)');
    r[5] = new Date();
  });
  sh.getRange(2, 1, n, 6).setValues(rows);
}

/**
 * รายชื่ออาจารย์ประจำสาขา — กรอกเองในแท็บ "รายชื่ออาจารย์" (1 แถวต่อ 1 คน)
 * คำนำหน้าให้ใช้คำเดียวกับในเว็บ: อาจารย์ / อาจารย์ ดร. / ผู้ช่วยศาสตราจารย์ / ผู้ช่วยศาสตราจารย์ ดร. / รองศาสตราจารย์ / รองศาสตราจารย์ ดร. / ศาสตราจารย์
 */
function getFaculty() {
  const sh = getSheet_(SHEETS.faculty);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 5).getDisplayValues()
    .filter(r => r[2])
    .map(r => ({ dept: r[0].trim(), prefix: r[1].trim(), name: r[2].trim(), email: r[3].trim(), tel: r[4].trim() }));
}

/** บันทึกร่าง (เขียนอย่างเดียว ไม่ส่งข้อมูลร่างกลับไปที่หน้าเว็บ) — 1 แถวต่ออีเมลหัวหน้าโครงการ */
function saveDraft(d) {
  const email = String(d.leadEmail || '').trim().toLowerCase();
  if (!/^\S+@\S+\.\S+$/.test(email)) throw new Error('ต้องกรอกอีเมลหัวหน้าโครงการก่อนบันทึกเข้าระบบ');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEETS.drafts);
    const links = d.links || [];
    const row = [email, `${d.leadPrefix || ''} ${d.leadName || ''}`.trim(), d.titleTh || '', new Date(),
      links.length, links.map(l => `${l.type}: ${l.url}`).join('\n'), JSON.stringify(d).slice(0, 49000)];
    const n = sh.getLastRow() - 1;
    const emails = n > 0 ? sh.getRange(2, 1, n, 1).getDisplayValues().flat() : [];
    const at = emails.indexOf(email);
    if (at > -1) sh.getRange(at + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return { savedAt: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}
