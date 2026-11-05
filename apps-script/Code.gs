/**
 * ระบบเสนอโครงร่างงานวิจัย คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏศรีสะเกษ
 * Backend API: Google Apps Script + Google Sheets + Google Drive
 * หน้าเว็บอยู่บน GitHub Pages (index.html / track.html / admin.html) และเรียก API นี้ผ่าน config.js
 *
 * การติดตั้ง
 * 1) script.google.com > โปรเจกต์ใหม่ > วางโค้ดนี้
 * 2) ใส่รหัส Google Sheets ที่ SHEET_ID ด้านล่าง (บัญชีที่รันสคริปต์ต้องมีสิทธิ์แก้ไขชีต)
 * 3) รัน setup() หนึ่งครั้งและอนุญาตสิทธิ์ แล้วรัน createAdminKey() เพื่อสร้างรหัสผู้ดูแล (ดูในบันทึกการดำเนินการ)
 * 4) ทำให้ใช้งานได้ > เว็บแอป — เรียกใช้ในฐานะ: ฉัน, ผู้มีสิทธิ์เข้าถึง: ทุกคน
 * 5) นำ URL ที่ลงท้ายด้วย /exec ไปใส่ใน config.js
 * เมื่อแก้โค้ด: จัดการการทำให้ใช้งานได้ > แก้ไข > เวอร์ชันใหม่ (URL เดิมใช้ต่อได้)
 */

/** Google Sheets ที่ใช้เก็บข้อมูล */
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID') || 'ใส่รหัส Google Sheets ที่นี่';

const STATUSES = ['รอพิจารณา', 'กำลังพิจารณา', 'ให้แก้ไข', 'ผ่านการพิจารณา', 'ไม่ผ่านการพิจารณา'];
const MAX_UPLOAD_MB = 10;
const UPLOAD_TYPES = ['application/pdf', 'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'image/jpeg', 'image/png'];

const SHEETS = {
  proposals: {
    name: 'โครงร่างวิจัย',
    headers: ['เลขที่','วันเวลาที่ส่ง','สถานะ','ปีงบประมาณ','ชื่อโครงการ (ไทย)','ชื่อโครงการ (อังกฤษ)',
      'แหล่งทุน','ประเภทโครงการ','ประเภทการวิจัย','สาขา','กลุ่มเป้าหมาย','คำสำคัญ',
      'คำนำหน้า','หัวหน้าโครงการ','สัดส่วนหัวหน้า (%)','สาขาวิชา','โทรศัพท์','อีเมล','ผู้ร่วมวิจัย',
      'ความเป็นมา','วัตถุประสงค์','กรอบแนวคิด','รูปแบบการวิจัย','พื้นที่','กลุ่มตัวอย่าง','ขนาดตัวอย่าง',
      'เครื่องมือ/การวิเคราะห์','เดือนเริ่ม','ระยะเวลา (เดือน)','แผนการดำเนินงาน',
      'ยุทธศาสตร์','SDGs','การลดความเหลื่อมล้ำทางสุขภาพ','ภาวะสุขภาพที่มุ่งแก้ไข','กลุ่มเปราะบางที่ได้รับประโยชน์','แนวทางลดความเหลื่อมล้ำ','คำอธิบายการลดความเหลื่อมล้ำ','เป้าหมายการเผยแพร่','ประโยชน์ที่คาดว่าจะได้รับ','จริยธรรมการวิจัย',
      'ค่าตอบแทน','ค่าใช้สอย','ค่าวัสดุ','งบประมาณรวม','ลิงก์โครงร่างฉบับสมบูรณ์','เอกสารแนบทั้งหมด','ผู้ส่ง (บัญชี Google)',
      'หมายเหตุผู้พิจารณา','อัปเดตสถานะเมื่อ']
  },
  members: { name: 'คณะผู้วิจัย', headers: ['เลขที่','บทบาท','ชื่อ–นามสกุล','สังกัด','สัดส่วน (%)'] },
  budget:  { name: 'งบประมาณ', headers: ['เลขที่','หมวด','รายการ','จำนวนเงิน (บาท)'] },
  faculty: { name: 'รายชื่ออาจารย์', headers: ['สาขาวิชา','คำนำหน้า / ตำแหน่ง','ชื่อ–นามสกุล','อีเมล','โทรศัพท์'] },
  drafts:  { name: 'ร่างที่บันทึกไว้', headers: ['อีเมลหัวหน้าโครงการ','หัวหน้าโครงการ','ชื่อโครงการ','บันทึกล่าสุด','จำนวนลิงก์เอกสาร','ลิงก์เอกสารแนบ','ข้อมูลร่าง (JSON)'] },
  files:   { name: 'เอกสารแนบ', headers: ['เลขที่','ประเภทเอกสาร','ชื่อไฟล์ใน Drive','ลิงก์','คณะเปิดอ่านได้หรือไม่','ตรวจเมื่อ'] },
  history: { name: 'ประวัติสถานะ', headers: ['เลขที่','สถานะ','หมายเหตุ','วันเวลา','แจ้งอีเมล'] },
  uploads: { name: 'ไฟล์ที่อัปโหลด', headers: ['วันเวลา','อีเมลผู้อัปโหลด','ประเภทเอกสาร','ชื่อไฟล์','ขนาด (KB)','ลิงก์'] }
};

/* ============================== API ============================== */

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const d = req.data || {};
    let result;
    switch (req.action) {
      // ผู้วิจัย
      case 'submit':    result = submitProposal(d); break;
      case 'saveDraft': result = saveDraft(d); break;
      case 'faculty':   result = getFaculty(); break;
      case 'checkLink': result = checkDriveLink(String(d.url || '')); break;
      case 'upload':    result = uploadFile(d); break;
      case 'status':    result = trackStatus(d); break;
      // ผู้ดูแล (ต้องมีรหัสผู้ดูแล)
      case 'adminLogin':  requireAdmin_(req.key); result = { ok: true, statuses: STATUSES }; break;
      case 'adminList':   requireAdmin_(req.key); result = adminList(); break;
      case 'adminStatus': requireAdmin_(req.key); result = adminUpdateStatus(d); break;
      default: throw new Error('ไม่รู้จักคำสั่ง ' + req.action);
    }
    out = { ok: true, result: result };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'NURS SSKRU research proposal API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================== ตั้งค่า ============================== */

function setup() {
  Object.values(SHEETS).forEach(cfg => getSheet_(cfg));
}

/** สร้างรหัสผู้ดูแลใหม่ (รหัสเดิมใช้ไม่ได้ทันที) — ดูรหัสในบันทึกการดำเนินการ */
function createAdminKey() {
  const key = 'NURS-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY_HASH', hash_(key));
  console.log('รหัสผู้ดูแลใหม่: ' + key);
  return key;
}

function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8));
}

function requireAdmin_(key) {
  const want = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY_HASH');
  if (!want) throw new Error('ยังไม่ได้สร้างรหัสผู้ดูแล (รัน createAdminKey ใน Apps Script)');
  if (!key || hash_(String(key).trim()) !== want) {
    Utilities.sleep(800);
    throw new Error('รหัสผู้ดูแลไม่ถูกต้อง');
  }
}

function getSheet_(cfg) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(cfg.name);
  if (!sh) sh = ss.insertSheet(cfg.name);
  const style = r => r.setFontWeight('bold').setBackground('#1B6B3A').setFontColor('#ffffff');
  if (sh.getLastRow() === 0) {
    sh.appendRow(cfg.headers);
    style(sh.getRange(1, 1, 1, cfg.headers.length));
    sh.setFrozenRows(1);
    return sh;
  }
  const width = Math.max(sh.getLastColumn(), 1);
  const have = sh.getRange(1, 1, 1, width).getDisplayValues()[0];
  // เพิ่มคอลัมน์ใหม่ต่อท้ายได้เอง ถ้าคอลัมน์เดิมตรงกัน
  const sameStart = have.slice(0, Math.min(have.length, cfg.headers.length)).every((h, i) => h === cfg.headers[i]);
  if (sameStart && have.length < cfg.headers.length) {
    const extra = cfg.headers.slice(have.length);
    style(sh.getRange(1, have.length + 1, 1, extra.length).setValues([extra]));
  } else if (!sameStart) {
    if (sh.getLastRow() === 1) style(sh.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]));
    else throw new Error(`หัวตารางในแท็บ "${cfg.name}" ไม่ตรงกับรุ่นปัจจุบัน กรุณาเปลี่ยนชื่อแท็บเดิม แล้วรัน setup อีกครั้ง`);
  }
  return sh;
}

function rows_(cfg) {
  const sh = getSheet_(cfg);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cfg.headers.length).getValues();
}

/* ============================== ผู้วิจัย ============================== */

function submitProposal(d) {
  const email = String(d.leadEmail || '').trim();
  if (!d.titleTh || !/^\S+@\S+\.\S+$/.test(email)) throw new Error('ข้อมูลไม่ครบ: ต้องมีชื่อโครงการและอีเมลหัวหน้าโครงการ');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEETS.proposals);
    const ref = nextRef_(sh, d.fy);
    const budget = d.budget || {};
    const sum = k => (budget[k] || []).reduce((a, r) => a + (Number(r.a) || 0), 0);
    const join = v => [].concat(v || []).join(', ');
    const now = new Date();

    sh.appendRow([
      ref, now, STATUSES[0], d.fy, d.titleTh, d.titleEn,
      d.fund, d.projType, d.resType, d.field, d.target, join(d.keywords),
      d.leadPrefix, d.leadName, Number(d.leadShare) || 0, d.leadDept, "'" + (d.leadTel || ''), email,
      (d.members || []).map(m => `${m.name} (${m.share || 0}%)`).join('; '),
      d.background, (d.objectives || []).map((o, i) => `${i + 1}) ${o}`).join('\n'), d.framework,
      d.design, d.area, d.sample, d.n, d.tools, d.start, d.months,
      (d.activities || []).map(a => `${a.name}: เดือน ${[].concat(a.months || []).sort((x, y) => x - y).join(',')}`).join('\n'),
      d.strategy, join(d.sdgs), d.equityLevel, join(d.healthIssues), join(d.equityGroups), join(d.equityDims), d.equityDesc, join(d.outputs), d.benefit, d.ethics,
      sum('comp'), sum('use'), sum('mat'), Number(d.total) || 0,
      ((d.links || [])[0] || {}).url || '',
      (d.links || []).map(l => `${l.type}: ${l.url}`).join('\n'),
      '', '', now
    ]);

    const ms = getSheet_(SHEETS.members);
    ms.appendRow([ref, 'หัวหน้าโครงการ', `${d.leadPrefix || ''} ${d.leadName || ''}`.trim(), 'คณะพยาบาลศาสตร์', Number(d.leadShare) || 0]);
    (d.members || []).forEach(m => ms.appendRow([ref, 'ผู้ร่วมวิจัย', m.name, m.aff, Number(m.share) || 0]));

    const bs = getSheet_(SHEETS.budget);
    const CAT = { comp: 'ค่าตอบแทน', use: 'ค่าใช้สอย', mat: 'ค่าวัสดุ' };
    Object.keys(CAT).forEach(k => (budget[k] || [])
      .filter(r => r.n || Number(r.a))
      .forEach(r => bs.appendRow([ref, CAT[k], r.n, Number(r.a) || 0])));

    const fs = getSheet_(SHEETS.files);
    (d.links || []).forEach(l => {
      const c = checkDriveLink(l.url);
      fs.appendRow([ref, l.type, c.name || l.name || '', l.url,
        c.ok ? 'เปิดได้' : (c.drive ? 'ยังไม่ได้แชร์สิทธิ์ให้คณะ' : 'ลิงก์ภายนอก (ตรวจเอง)'), now]);
    });

    getSheet_(SHEETS.history).appendRow([ref, STATUSES[0], 'ได้รับโครงร่าง', now, 'ส่งอีเมลยืนยัน']);
    notifyReceived_(ref, d);
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

function mailFrame_(inner) {
  return `<div style="font-family:Tahoma,sans-serif;font-size:14px;line-height:1.7;color:#16241B">
    ${inner}
    <p style="color:#5D6E62;font-size:12px;margin-top:24px">งานวิจัย คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏศรีสะเกษ<br>
    ติดตามสถานะได้ที่หน้า “ติดตามสถานะ” ของระบบเสนอโครงร่างงานวิจัย โดยใช้เลขที่โครงร่างและอีเมลของท่าน</p></div>`;
}

function notifyReceived_(ref, d) {
  try {
    MailApp.sendEmail({
      to: d.leadEmail,
      subject: `[คณะพยาบาลศาสตร์ มรภ.ศรีสะเกษ] ได้รับโครงร่างวิจัย ${ref}`,
      htmlBody: mailFrame_(`<p>เรียน ${d.leadPrefix || ''} ${d.leadName || ''}</p>
        <p>งานวิจัย คณะพยาบาลศาสตร์ ได้รับโครงร่างวิจัยของท่านแล้ว</p>
        <p><b>เลขที่:</b> ${ref}<br><b>ชื่อโครงการ:</b> ${d.titleTh}<br>
        <b>งบประมาณที่เสนอ:</b> ${Number(d.total || 0).toLocaleString('th-TH')} บาท<br>
        <b>สถานะ:</b> ${STATUSES[0]}</p>`)
    });
  } catch (e) { console.warn(e); }
}

/** ติดตามสถานะ: ต้องระบุเลขที่และอีเมลหัวหน้าโครงการให้ตรงกัน */
function trackStatus(d) {
  const ref = String(d.ref || '').trim().toUpperCase();
  const email = String(d.email || '').trim().toLowerCase();
  if (!ref || !email) throw new Error('กรุณากรอกเลขที่โครงร่างและอีเมล');
  const H = SHEETS.proposals.headers, col = n => H.indexOf(n);
  const row = rows_(SHEETS.proposals).find(r => String(r[0]).toUpperCase() === ref &&
    String(r[col('อีเมล')]).trim().toLowerCase() === email);
  if (!row) { Utilities.sleep(600); throw new Error('ไม่พบโครงร่างที่ตรงกับเลขที่และอีเมลนี้'); }
  const history = rows_(SHEETS.history).filter(r => String(r[0]).toUpperCase() === ref)
    .map(r => ({ status: r[1], note: r[2], at: iso_(r[3]) }));
  return {
    ref: row[0], submitted: iso_(row[1]), status: row[2], fy: row[3], title: row[4],
    lead: `${row[col('คำนำหน้า')]} ${row[col('หัวหน้าโครงการ')]}`.trim(),
    total: Number(row[col('งบประมาณรวม')]) || 0,
    note: row[col('หมายเหตุผู้พิจารณา')], updated: iso_(row[col('อัปเดตสถานะเมื่อ')]),
    history: history
  };
}

function iso_(v) { return v instanceof Date ? v.toISOString() : (v ? String(v) : ''); }

/** อัปโหลดไฟล์เข้าโฟลเดอร์ Google Drive ของระบบ */
function uploadFile(d) {
  const email = String(d.email || '').trim().toLowerCase();
  if (!/^\S+@\S+\.\S+$/.test(email)) throw new Error('กรุณากรอกอีเมลหัวหน้าโครงการ (ขั้นตอนที่ 2) ก่อนอัปโหลดไฟล์');
  if (UPLOAD_TYPES.indexOf(d.mimeType) < 0) throw new Error('รองรับเฉพาะไฟล์ PDF, Word, Excel, JPG และ PNG');
  const bytes = Utilities.base64Decode(String(d.base64 || ''));
  if (!bytes.length) throw new Error('ไฟล์ว่างเปล่า');
  if (bytes.length > MAX_UPLOAD_MB * 1024 * 1024) throw new Error(`ไฟล์ใหญ่เกิน ${MAX_UPLOAD_MB} MB`);
  const safeName = String(d.name || 'file').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
  const folder = emailFolder_(email);
  const file = folder.createFile(Utilities.newBlob(bytes, d.mimeType, safeName));
  file.setDescription(`${d.type || 'เอกสารแนบ'} — อัปโหลดโดย ${email}`);
  getSheet_(SHEETS.uploads).appendRow([new Date(), email, d.type || '', safeName, Math.round(bytes.length / 1024), file.getUrl()]);
  return { id: file.getId(), name: safeName, url: file.getUrl() };
}

function uploadRoot_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('UPLOAD_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const folder = DriveApp.createFolder('ระบบเสนอโครงร่างวิจัย คณะพยาบาลศาสตร์ — ไฟล์แนบ');
  props.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}

function emailFolder_(email) {
  const root = uploadRoot_();
  const it = root.getFoldersByName(email);
  return it.hasNext() ? it.next() : root.createFolder(email);
}

/** ตรวจว่าบัญชีที่รันสคริปต์เปิดไฟล์จากลิงก์ได้หรือไม่ — คืนค่า { ok, drive, name } */
function driveId_(url) {
  const m = String(url).match(/\/(?:d|folders)\/([\w-]{15,})/) || String(url).match(/[?&]id=([\w-]{15,})/);
  return m ? m[1] : '';
}

function checkDriveLink(url) {
  const id = driveId_(url);
  if (!id || !/(drive|docs)\.google\.com/.test(url)) return { ok: false, drive: false, name: '' };
  try { return { ok: true, drive: true, name: DriveApp.getFileById(id).getName() }; } catch (e) {}
  try { return { ok: true, drive: true, name: '📁 ' + DriveApp.getFolderById(id).getName() }; } catch (e) {}
  return { ok: false, drive: true, name: '' };
}

/** ตรวจลิงก์ในแท็บ "เอกสารแนบ" ซ้ำทั้งหมด */
function recheckAllLinks() {
  const sh = getSheet_(SHEETS.files);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const rows = sh.getRange(2, 1, n, 6).getValues();
  rows.forEach(r => {
    if (!r[3]) return;
    const c = checkDriveLink(r[3]);
    r[2] = c.name || r[2];
    r[4] = c.ok ? 'เปิดได้' : (c.drive ? 'ยังไม่ได้แชร์สิทธิ์ให้คณะ' : 'ลิงก์ภายนอก (ตรวจเอง)');
    r[5] = new Date();
  });
  sh.getRange(2, 1, n, 6).setValues(rows);
}

/** รายชื่ออาจารย์ประจำสาขา — กรอกในแท็บ "รายชื่ออาจารย์" */
function getFaculty() {
  return rows_(SHEETS.faculty).map(r => r.map(v => String(v).trim()))
    .filter(r => r[2])
    .map(r => ({ dept: r[0], prefix: r[1], name: r[2], email: r[3], tel: r[4] }));
}

/** บันทึกร่าง (เขียนอย่างเดียว) — 1 แถวต่ออีเมลหัวหน้าโครงการ */
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

/* ============================== ผู้ดูแล ============================== */

function adminList() {
  const H = SHEETS.proposals.headers;
  const files = {}, members = {}, budget = {}, history = {};
  rows_(SHEETS.files).forEach(r => (files[r[0]] = files[r[0]] || []).push({ type: r[1], name: r[2], url: r[3], access: r[4] }));
  rows_(SHEETS.members).forEach(r => (members[r[0]] = members[r[0]] || []).push({ role: r[1], name: r[2], aff: r[3], share: r[4] }));
  rows_(SHEETS.budget).forEach(r => (budget[r[0]] = budget[r[0]] || []).push({ cat: r[1], item: r[2], amount: Number(r[3]) || 0 }));
  rows_(SHEETS.history).forEach(r => (history[r[0]] = history[r[0]] || []).push({ status: r[1], note: r[2], at: iso_(r[3]), mailed: r[4] }));
  const list = rows_(SHEETS.proposals).filter(r => r[0]).map(r => {
    const o = {};
    H.forEach((h, i) => o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]);
    o.files = files[r[0]] || []; o.members = members[r[0]] || [];
    o.budgetItems = budget[r[0]] || []; o.history = history[r[0]] || [];
    return o;
  });
  return { statuses: STATUSES, proposals: list, sheetUrl: 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit',
    uploadFolder: PropertiesService.getScriptProperties().getProperty('UPLOAD_FOLDER_ID') || '' };
}

function adminUpdateStatus(d) {
  const ref = String(d.ref || '').trim();
  const status = String(d.status || '');
  if (STATUSES.indexOf(status) < 0) throw new Error('สถานะไม่ถูกต้อง');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(SHEETS.proposals);
    const H = SHEETS.proposals.headers, col = n => H.indexOf(n) + 1;
    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().flat();
    const at = ids.indexOf(ref);
    if (at < 0) throw new Error('ไม่พบเลขที่ ' + ref);
    const rowNo = at + 2, now = new Date();
    sh.getRange(rowNo, col('สถานะ')).setValue(status);
    sh.getRange(rowNo, col('หมายเหตุผู้พิจารณา')).setValue(d.note || '');
    sh.getRange(rowNo, col('อัปเดตสถานะเมื่อ')).setValue(now);
    let mailed = 'ไม่ส่ง';
    if (d.notify) {
      const v = sh.getRange(rowNo, 1, 1, H.length).getValues()[0];
      const email = v[col('อีเมล') - 1];
      try {
        MailApp.sendEmail({
          to: email,
          subject: `[คณะพยาบาลศาสตร์ มรภ.ศรีสะเกษ] ผลการพิจารณาโครงร่างวิจัย ${ref}: ${status}`,
          htmlBody: mailFrame_(`<p>เรียน ${v[col('คำนำหน้า') - 1]} ${v[col('หัวหน้าโครงการ') - 1]}</p>
            <p>โครงร่างวิจัยเรื่อง <b>${v[col('ชื่อโครงการ (ไทย)') - 1]}</b> (เลขที่ ${ref})</p>
            <p>สถานะล่าสุด: <b>${status}</b></p>
            ${d.note ? `<p><b>ความเห็นผู้พิจารณา:</b><br>${String(d.note).replace(/</g, '&lt;').replace(/\n/g, '<br>')}</p>` : ''}`)
        });
        mailed = 'ส่งแล้ว: ' + email;
      } catch (e) { mailed = 'ส่งไม่สำเร็จ: ' + e.message; }
    }
    getSheet_(SHEETS.history).appendRow([ref, status, d.note || '', now, mailed]);
    return { ref: ref, status: status, updated: now.toISOString(), mailed: mailed };
  } finally {
    lock.releaseLock();
  }
}
