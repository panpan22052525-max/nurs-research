/* ส่งออกโครงร่างวิจัยเป็น Word (.doc) และ PDF (ผ่านหน้าต่างพิมพ์ของเบราว์เซอร์)
   ใช้ร่วมกันได้ทุกหน้า: ProposalExport.word(data, ref) / ProposalExport.pdf(data, ref) */
(function () {
  const esc = v => String(v ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const txt = v => esc(v).replace(/\r?\n/g, '<br>');
  const money = n => (+n || 0).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const list = v => [].concat(v || []).filter(Boolean);
  const thMonth = ym => { const m = /^(\d{4})-(\d{2})/.exec(ym || ''); if (!m) return ym || ''; return new Date(+m[1], +m[2] - 1, 1).toLocaleDateString('th-TH', { month: 'long', year: 'numeric' }); };
  const today = () => new Date().toLocaleDateString('th-TH', { dateStyle: 'long' });

  /* แปลงจำนวนเงินเป็นตัวอักษรไทย เช่น 19,500 → หนึ่งหมื่นเก้าพันห้าร้อยบาทถ้วน */
  function bahtText(num) {
    num = Math.round((+num || 0) * 100) / 100;
    const D = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'], P = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
    const read = s => {
      s = String(+s); if (s === '0') return '';
      if (s.length > 6) return read(s.slice(0, -6)) + 'ล้าน' + read(s.slice(-6));
      let out = '';
      for (let i = 0; i < s.length; i++) {
        const d = +s[i], pos = s.length - i - 1; if (!d) continue;
        if (pos === 1 && d === 1) out += 'สิบ';
        else if (pos === 1 && d === 2) out += 'ยี่สิบ';
        else if (pos === 0 && d === 1 && s.length > 1) out += 'เอ็ด';
        else out += D[d] + P[pos];
      }
      return out;
    };
    const [b, s] = num.toFixed(2).split('.');
    const baht = read(b), sat = read(s);
    if (!baht && !sat) return 'ศูนย์บาทถ้วน';
    return (baht ? baht + 'บาท' : '') + (sat ? sat + 'สตางค์' : 'ถ้วน');
  }

  const CATS = [['comp', 'หมวดค่าตอบแทน'], ['use', 'หมวดค่าใช้สอย'], ['mat', 'หมวดค่าวัสดุ']];
  const PARTS = ['ส่วนที่ 1 ข้อมูลทั่วไปของโครงการ', 'ส่วนที่ 2 คณะผู้วิจัย', 'ส่วนที่ 3 สาระสำคัญของโครงการ',
    'ส่วนที่ 4 ความสอดคล้องและผลผลิตที่คาดว่าจะได้รับ', 'ส่วนที่ 5 งบประมาณ', 'ส่วนที่ 6 เอกสารแนบ'];

  /* สร้างเนื้อหาเอกสาร — pb คือรูปแบบตัวขึ้นหน้าใหม่ของแต่ละปลายทาง */
  function build(d, ref, opt) {
    const PB = opt.pageBreak, logo = opt.logo;
    const row = (k, v) => `<tr><td class="k">${esc(k)}</td><td>${v ? v : '<span class="mute">—</span>'}</td></tr>`;
    const kv = rows => `<table class="kv">${rows.join('')}</table>`;
    const h1 = t => `<h1>${esc(t)}</h1>`, h2 = t => `<h2>${esc(t)}</h2>`;
    const para = v => v ? `<p class="body">${txt(v)}</p>` : '<p class="mute">—</p>';
    const lead = `${d.leadPrefix || ''} ${d.leadName || ''}`.trim();
    const fy = d.fy ? `ปีงบประมาณ พ.ศ. ${esc(d.fy)}` : '';

    // ปก
    let h = `<div class="cover">
      ${logo ? `<p class="c"><img src="${esc(logo)}" width="120" height="120" alt="ตราคณะ" style="width:2.8cm;height:auto"></p>` : ''}
      <p class="c cover-kicker">แบบเสนอโครงร่างงานวิจัย</p>
      <p class="c">${fy}</p>
      <p class="c cover-title">${txt(d.titleTh || '')}</p>
      <p class="c cover-en">${txt(d.titleEn || '')}</p>
      <p class="c" style="margin-top:1.1cm">เสนอโดย</p>
      <p class="c"><b>${esc(lead)}</b></p>
      <p class="c">${d.leadDept ? 'สาขาวิชา' + esc(d.leadDept) : ''}</p>
      <p class="c">คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏศรีสะเกษ</p>
      <p class="c" style="margin-top:.8cm">แหล่งทุน: ${esc(d.fund || '—')}</p>
      <p class="c">${ref ? `เลขที่โครงร่าง <b>${esc(ref)}</b> · ` : ''}วันที่ ${today()}</p>
    </div>`;

    // สารบัญ
    h += PB + h1('สารบัญ') + `<table class="toc">${PARTS.map(p => `<tr><td>${esc(p)}</td></tr>`).join('')}<tr><td>ลงนามรับรองของหัวหน้าโครงการ</td></tr></table>`;

    // 1 ข้อมูลทั่วไป
    h += PB + h1(PARTS[0]) + kv([
      row('ชื่อโครงการ (ภาษาไทย)', txt(d.titleTh)), row('ชื่อโครงการ (ภาษาอังกฤษ)', txt(d.titleEn)),
      row('ปีงบประมาณ', esc(d.fy)), row('แหล่งทุน', esc(d.fund)), row('ประเภทโครงการ', esc(d.projType)),
      row('ประเภทการวิจัย', esc(d.resType)), row('สาขาการพยาบาล', esc(d.field)), row('กลุ่มเป้าหมาย', esc(d.target)),
      row('คำสำคัญ', esc(list(d.keywords).join(', ')))]);

    // 2 คณะผู้วิจัย
    const team = [{ role: 'หัวหน้าโครงการ', name: lead, aff: 'คณะพยาบาลศาสตร์' + (d.leadDept ? ' สาขาวิชา' + d.leadDept : ''), share: d.leadShare }]
      .concat(list(d.members).map(m => ({ role: 'ผู้ร่วมวิจัย', name: m.name, aff: m.aff, share: m.share })));
    h += PB + h1(PARTS[1]) +
      `<table class="grid"><thead><tr><th style="width:8%">ลำดับ</th><th>ชื่อ–สกุล</th><th>บทบาท</th><th>สังกัด</th><th style="width:14%">สัดส่วน (%)</th></tr></thead><tbody>` +
      team.map((t, i) => `<tr><td class="c">${i + 1}</td><td>${esc(t.name)}</td><td>${esc(t.role)}</td><td>${esc(t.aff)}</td><td class="c">${esc(t.share || 0)}</td></tr>`).join('') +
      `</tbody></table>` + h2('ข้อมูลติดต่อหัวหน้าโครงการ') + kv([row('โทรศัพท์', esc(d.leadTel)), row('อีเมล', esc(d.leadEmail))]);

    // 3 สาระสำคัญ
    const months = Math.max(+d.months || 0, ...list(d.activities).flatMap(a => a.months || []), 0);
    const gantt = list(d.activities).length && months ? `<table class="grid gantt"><thead><tr><th>กิจกรรม</th>${Array.from({ length: months }, (_, i) => `<th>${i + 1}</th>`).join('')}</tr></thead><tbody>` +
      list(d.activities).map(a => `<tr><td>${esc(a.name)}</td>${Array.from({ length: months }, (_, i) => (a.months || []).includes(i + 1) ? '<td class="on">&#9632;</td>' : '<td></td>').join('')}</tr>`).join('') +
      `</tbody></table><p class="mute small">หมายเหตุ: ตัวเลขคือเดือนที่ของโครงการ เริ่มเดือน${esc(thMonth(d.start))}</p>` : '<p class="mute">—</p>';
    h += PB + h1(PARTS[2]) +
      h2('3.1 ความเป็นมาและความสำคัญของปัญหา') + para(d.background) +
      h2('3.2 วัตถุประสงค์ของการวิจัย') + (list(d.objectives).length ? list(d.objectives).map((o, i) => `<p class="body">${i + 1}) ${txt(o)}</p>`).join('') : '<p class="mute">—</p>') +
      h2('3.3 กรอบแนวคิดการวิจัย') + para(d.framework) +
      h2('3.4 ระเบียบวิธีวิจัย') + kv([row('รูปแบบการวิจัย', esc(d.design)), row('พื้นที่ดำเนินการวิจัย', esc(d.area)),
        row('ประชากรและกลุ่มตัวอย่าง', txt(d.sample)), row('ขนาดตัวอย่าง', d.n ? esc(d.n) + ' คน' : ''),
        row('เครื่องมือและการวิเคราะห์ข้อมูล', txt(d.tools)),
        row('ระยะเวลาดำเนินการ', d.months ? `${esc(d.months)} เดือน (เริ่ม${esc(thMonth(d.start))})` : '')]) +
      h2('3.5 แผนการดำเนินงาน') + gantt;

    // 4 ความสอดคล้อง
    h += PB + h1(PARTS[3]) + kv([
      row('ความสอดคล้องกับยุทธศาสตร์', esc(d.strategy)), row('เป้าหมายการพัฒนาที่ยั่งยืน (SDGs)', esc(list(d.sdgs).join(', '))),
      row('การลดความเหลื่อมล้ำทางสุขภาพ', esc(d.equityLevel)), row('ภาวะสุขภาพที่มุ่งแก้ไข', esc(list(d.healthIssues).join(', '))),
      row('กลุ่มเปราะบางที่ได้รับประโยชน์', esc(list(d.equityGroups).join(', '))), row('แนวทางลดความเหลื่อมล้ำ', esc(list(d.equityDims).join(', '))),
      row('คำอธิบายการลดความเหลื่อมล้ำ', txt(d.equityDesc)), row('เป้าหมายการเผยแพร่ผลงาน', esc(list(d.outputs).join(', '))),
      row('ประโยชน์ที่คาดว่าจะได้รับ', txt(d.benefit)), row('จริยธรรมการวิจัยในมนุษย์', esc(d.ethics))]);

    // 5 งบประมาณ
    const B = d.budget || {}; let no = 0, grand = 0, rows = '';
    CATS.forEach(([k, t]) => {
      const items = list(B[k]).filter(r => (r.n || '').trim() || +r.a); if (!items.length) return;
      const sub = items.reduce((a, r) => a + (+r.a || 0), 0); grand += sub;
      rows += `<tr class="cat"><td colspan="2"><b>${esc(t)}</b></td><td class="r"><b>${money(sub)}</b></td></tr>` +
        items.map(r => `<tr><td class="c">${++no}</td><td>${esc(r.n)}</td><td class="r">${money(r.a)}</td></tr>`).join('');
    });
    if (!grand && +d.total) grand = +d.total;
    h += PB + h1(PARTS[4]) +
      `<table class="grid"><thead><tr><th style="width:10%">ลำดับ</th><th>รายการ</th><th style="width:24%">จำนวนเงิน (บาท)</th></tr></thead><tbody>${rows || '<tr><td colspan="3" class="c mute">—</td></tr>'}` +
      `<tr class="total"><td colspan="2" class="r"><b>รวมงบประมาณทั้งสิ้น</b></td><td class="r"><b>${money(grand)}</b></td></tr></tbody></table>` +
      `<p class="c">(${bahtText(grand)})</p>`;

    // 6 เอกสารแนบ
    const links = list(d.links);
    h += PB + h1(PARTS[5]) + (links.length ? `<table class="grid"><thead><tr><th style="width:8%">ลำดับ</th><th style="width:32%">เอกสาร</th><th>ลิงก์</th></tr></thead><tbody>` +
      links.map((l, i) => `<tr><td class="c">${i + 1}</td><td>${esc(l.type)}${l.name ? `<br><span class="mute small">${esc(l.name)}</span>` : ''}</td><td class="url"><a href="${esc(l.url)}">${esc(l.url)}</a></td></tr>`).join('') +
      `</tbody></table>` : '<p class="mute">ไม่มีเอกสารแนบ</p>');

    // ลงนาม
    h += `<div class="sign">${h2('ลงนามรับรองของหัวหน้าโครงการ')}
      <p class="body">ข้าพเจ้าขอรับรองว่าข้อมูลในแบบเสนอโครงร่างงานวิจัยนี้เป็นความจริงทุกประการ และโครงการนี้ไม่ได้รับทุนซ้ำซ้อนจากแหล่งทุนอื่น</p>
      <p class="sig">ลงชื่อ ..................................................... หัวหน้าโครงการ</p>
      <p class="sig">( ${esc(lead) || '.....................................................'} )</p>
      <p class="sig">วันที่ ........ เดือน ............................ พ.ศ. ............</p></div>`;
    return h;
  }

  const fileName = (d, ref, ext) => (ref || 'โครงร่างวิจัย') + '_' + String(d.titleTh || 'proposal').replace(/[\\/:*?"<>|\n\r]+/g, ' ').trim().slice(0, 60) + '.' + ext;
  const logoUrl = () => { try { return new URL('logo.jpg', location.href).href } catch (e) { return '' } };

  /* ---------- Word ---------- */
  function word(d, ref) {
    const font = "'TH Sarabun New','TH SarabunPSK','Sarabun',sans-serif";
    const PB = `<br clear="all" style="mso-special-character:line-break;page-break-before:always">`;
    // เขียน style ลงในแต่ละ element โดยตรง เพราะ Word/LibreOffice/Google Docs อ่าน CSS class ได้ไม่ครบ
    const B = 'solid windowtext .5pt', G = '#EAF1E7';
    const RULES = [
      ['h1', 'font-size:20pt;font-weight:bold;color:#0F4526;margin:0 0 10pt 0'],
      ['h2', 'font-size:18pt;font-weight:bold;margin:12pt 0 4pt 0'],
      ['table', 'border-collapse:collapse;width:100%'],
      ['p.body', 'text-indent:1.25cm;text-align:justify'],
      ['.cover-kicker', 'font-size:24pt;font-weight:bold;color:#0F4526;margin-top:12pt'],
      ['.cover-title', 'font-size:20pt;font-weight:bold;margin-top:18pt'],
      ['.cover-en', 'font-size:18pt;font-style:italic'],
      ['table.kv td', 'vertical-align:top;padding:3pt 6pt'],
      ['table.kv td.k', 'width:34%;font-weight:bold'],
      ['table.grid th', `border:${B};padding:2pt 5pt;background:${G};font-weight:bold;text-align:center;vertical-align:middle`],
      ['table.grid td', `border:${B};padding:2pt 5pt;vertical-align:top`],
      ['table.gantt th, table.gantt td', 'font-size:14pt;text-align:center;padding:1pt 2pt'],
      ['table.gantt td:first-child', 'text-align:left'],
      ['td.on', 'background:#1B6B3A;color:#1B6B3A'],
      ['tr.cat td', 'background:#F6F8F4'], ['tr.total td', `background:${G}`],
      ['td.url', 'font-size:13pt;word-break:break-all'],
      ['table.toc td', 'padding:4pt 0;border-bottom:dotted #999999 .5pt'],
      ['p.sig', 'text-align:center;margin-top:14pt'],
      ['.c', 'text-align:center'], ['.r', 'text-align:right'], ['.mute', 'color:#666666'], ['.small', 'font-size:14pt']
    ];
    const box = document.createElement('div');
    box.innerHTML = build(d, ref, { pageBreak: '<hr data-pb>', logo: logoUrl() });
    RULES.forEach(([sel, st]) => box.querySelectorAll(sel).forEach(el => el.setAttribute('style', (el.getAttribute('style') || '') + ';' + st)));
    box.querySelectorAll('h1,h2').forEach(el => el.setAttribute('style', el.getAttribute('style') + `;font-family:${font}`));
    box.querySelectorAll('td.on').forEach(el => { el.setAttribute('bgcolor', '#1B6B3A'); el.innerHTML = '&nbsp;'; });
    box.querySelectorAll('table.grid th, tr.total td').forEach(el => el.setAttribute('bgcolor', G));
    box.querySelectorAll('table.grid').forEach(t => { t.setAttribute('border', '1'); t.setAttribute('cellspacing', '0'); t.setAttribute('cellpadding', '4'); });
    const body = box.innerHTML.replace(/<hr data-pb="">/g, PB);
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40" lang="th">
<head><meta charset="utf-8"><title>${esc(d.titleTh || 'โครงร่างวิจัย')}</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
@page Section1{size:21.0cm 29.7cm;margin:2.54cm 2.54cm 2.54cm 3.17cm;mso-header-margin:1.25cm;mso-footer-margin:1.25cm;mso-footer:f1;mso-title-page:yes}
div.Section1{page:Section1}
body,p,td,th,li{font-family:${font};font-size:16pt;line-height:150%;mso-ansi-font-size:16pt;mso-bidi-font-size:16pt;mso-bidi-font-family:'TH Sarabun New';mso-fareast-language:TH;mso-bidi-language:TH}
p{margin:0 0 4pt 0}
h1{font-family:${font};font-size:20pt;mso-bidi-font-size:20pt;font-weight:bold;color:#0F4526;margin:0 0 10pt 0;line-height:150%}
h2{font-family:${font};font-size:18pt;mso-bidi-font-size:18pt;font-weight:bold;margin:12pt 0 4pt 0;line-height:150%}
p.body{text-indent:1.25cm;text-align:justify;text-justify:inter-cluster}
.c{text-align:center} .r{text-align:right} .mute{color:#666} .small{font-size:14pt;mso-bidi-font-size:14pt}
.cover-kicker{font-size:24pt;mso-bidi-font-size:24pt;font-weight:bold;color:#0F4526;margin-top:12pt}
.cover-title{font-size:20pt;mso-bidi-font-size:20pt;font-weight:bold;margin-top:18pt}
.cover-en{font-size:18pt;mso-bidi-font-size:18pt;font-style:italic}
table{border-collapse:collapse;width:100%}
table.kv td{vertical-align:top;padding:3pt 6pt;border:none}
table.kv td.k{width:34%;font-weight:bold}
table.grid th,table.grid td{border:solid windowtext .5pt;padding:2pt 5pt;vertical-align:top}
table.grid th{background:#EAF1E7;font-weight:bold;text-align:center}
table.gantt td,table.gantt th{font-size:14pt;mso-bidi-font-size:14pt;text-align:center;padding:1pt 2pt}
table.gantt td:first-child{text-align:left}
td.on{background:#1B6B3A;color:#1B6B3A}
tr.cat td{background:#F6F8F4} tr.total td{background:#EAF1E7}
td.url{font-size:13pt;mso-bidi-font-size:13pt;word-break:break-all}
table.toc td{padding:4pt 0;border-bottom:dotted #999 .5pt}
p.sig{text-align:center;margin-top:14pt}
div.sign{margin-top:24pt}
p.MsoFooter{text-align:center;font-family:${font};font-size:14pt}
</style></head>
<body lang="TH"><div class="Section1">${body}
<table id="hrdftrtbl" style="margin-left:50in" border="0" cellspacing="0" cellpadding="0"><tr><td>
<div style="mso-element:footer" id="f1"><p class="MsoFooter"><span style="mso-field-code:' PAGE '"></span></p></div>
</td></tr></table></div></body></html>`;
    const blob = new Blob(['﻿', html], { type: 'application/msword' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = fileName(d, ref, 'doc');
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  /* ---------- PDF (พิมพ์ → บันทึกเป็น PDF) ---------- */
  function pdf(d, ref) {
    const w = window.open('', '_blank');
    if (!w) { alert('เบราว์เซอร์บล็อกหน้าต่างใหม่ กรุณาอนุญาตป๊อปอัปสำหรับเว็บไซต์นี้แล้วลองอีกครั้ง'); return false; }
    const body = build(d, ref, { pageBreak: '<div class="pb"></div>', logo: logoUrl() });
    w.document.open();
    w.document.write(`<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(fileName(d, ref, 'pdf').replace(/\.pdf$/, ''))}</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,400;0,700;1,400&display=swap">
<style>
@page{size:A4;margin:2.54cm 2.54cm 2.54cm 3.17cm;@bottom-center{content:counter(page);font-family:'Sarabun';font-size:11pt}}
@page:first{@bottom-center{content:none}}
:root{color-scheme:light}
body{font-family:'TH Sarabun New','Sarabun',sans-serif;font-size:14pt;line-height:1.5;color:#111;background:#fff;margin:0}
.sheet{max-width:17.3cm;margin:0 auto;padding:16px}
.bar{position:sticky;top:0;background:#0F4526;color:#fff;padding:10px 16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;font-size:12pt}
.bar button{font:inherit;font-weight:700;background:#F1DFA4;color:#0F4526;border:0;border-radius:8px;padding:6px 16px;cursor:pointer}
@media print{.bar{display:none}.sheet{max-width:none;padding:0}}
p{margin:0 0 3pt}
h1{font-size:17pt;color:#0F4526;margin:0 0 8pt;line-height:1.4}
h2{font-size:15pt;margin:12pt 0 4pt;line-height:1.4;break-after:avoid}
p.body{text-indent:1.25cm;text-align:justify}
.c{text-align:center}.r{text-align:right}.mute{color:#666}.small{font-size:12pt}
.pb{break-before:page;height:0}
@media screen{.pb{border-top:2px dashed #ccc;margin:28px 0}}
.cover{padding-top:1.5cm}
.cover-kicker{font-size:22pt;font-weight:700;color:#0F4526;margin-top:12pt}
.cover-title{font-size:18pt;font-weight:700;margin-top:16pt}
.cover-en{font-size:15pt;font-style:italic}
table{border-collapse:collapse;width:100%;margin-bottom:6pt}
tr{break-inside:avoid}
table.kv td{vertical-align:top;padding:3pt 6pt}
table.kv td.k{width:34%;font-weight:700}
table.grid th,table.grid td{border:.6pt solid #333;padding:2pt 5pt;vertical-align:top}
table.grid th{background:#EAF1E7;text-align:center}
table.gantt th,table.gantt td{font-size:12pt;text-align:center;padding:1pt 2pt}
table.gantt td:first-child{text-align:left}
td.on{background:#1B6B3A;color:#1B6B3A;-webkit-print-color-adjust:exact;print-color-adjust:exact}
tr.cat td{background:#F6F8F4}tr.total td{background:#EAF1E7}
th,tr.cat td,tr.total td{-webkit-print-color-adjust:exact;print-color-adjust:exact}
td.url{font-size:11pt;word-break:break-all}td.url a{color:#1B6B3A}
table.toc td{padding:4pt 0;border-bottom:1px dotted #999}
p.sig{text-align:center;margin-top:14pt}.sign{margin-top:24pt;break-inside:avoid}
</style></head><body>
<div class="bar"><span>ตรวจดูเอกสารแล้วกด “บันทึกเป็น PDF” แล้วเลือกปลายทาง “Save as PDF / บันทึกเป็น PDF”</span><button type="button" onclick="window.print()">บันทึกเป็น PDF</button></div>
<div class="sheet">${body}</div>
<script>
(document.fonts&&document.fonts.ready?document.fonts.ready:Promise.resolve()).then(function(){
  var imgs=[].slice.call(document.images).map(function(i){return i.complete?0:new Promise(function(r){i.onload=i.onerror=r})});
  Promise.all(imgs).then(function(){setTimeout(function(){window.print()},300)});
});
<\/script></body></html>`);
    w.document.close();
    return true;
  }

  window.ProposalExport = { word, pdf, bahtText };
})();
