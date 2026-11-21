/**
 * ระบบตรวจสินค้าใหม่ (RA) — หลังบ้านบน Google Apps Script
 *
 * เก็บข้อมูลใน Google Drive โฟลเดอร์ "RA ตรวจสินค้าใหม่" แยกโฟลเดอร์ตามสินค้า:
 *   RA ตรวจสินค้าใหม่/
 *   ├── สรุปรายการสินค้า        (Google Sheet: 1 แถวต่อ 1 สินค้า)
 *   └── <ชื่อสินค้า>/
 *       ├── data.json            (ข้อมูลทั้งหมดของสินค้า)
 *       ├── Ingredient - ...     (ไฟล์ที่จัดซื้อแนบ)
 *       ├── Nutrition - ...
 *       ├── Process Flow - ...
 *       └── <ชื่อสินค้า> v1 สรุปข้อมูล   (Google Doc สร้างตอนส่งให้ RA)
 *
 * ก่อนใช้งาน: ตั้ง TEAM_KEY (รหัสจัดซื้อ) และ RA_KEY (รหัส RA) แล้ว Deploy เป็น Web app
 * (Execute as: Me, Who has access: Anyone) — ดูขั้นตอนใน README.md
 */

const ROOT_ID = '1BEopI-OxSE1of60Zzutl9jUX_uBPF9g9'; // โฟลเดอร์ "RA ตรวจสินค้าใหม่"
const TEAM_KEY = 'CHANGE-ME';    // รหัสฝ่ายจัดซื้อ: สร้าง/แก้/ส่งงานได้ แต่ Approve/Reject ไม่ได้
const RA_KEY = 'CHANGE-ME-RA';   // รหัส RA: ตรวจ และ Approve/Reject ได้ (ให้เฉพาะ RA) — อย่าใส่รหัสจริงลง GitHub
const TZ = 'Asia/Bangkok';
const SEC = { ing: 'Ingredient', nut: 'Nutrition', flow: 'Process Flow' };
const HEAD = ['รหัส', 'ชื่อสินค้า', 'สถานะ', 'ผู้ส่ง', 'อัปเดตล่าสุด', 'Ingredient รวม %',
  'ไฟล์ Ingredient', 'ไฟล์ Nutrition', 'ไฟล์ Process Flow', 'ผู้ตรวจ (RA)', 'เหตุผล Reject',
  'โฟลเดอร์', 'ไฟล์สรุป', 'ข้อมูล (JSON)'];
const STATUS = { draft: 'ร่าง', submitted: 'รอ RA ตรวจ', rejected: 'ถูก Reject', approved: 'อนุมัติแล้ว' };

function doGet() {
  return out({ ok: true, msg: 'RA API ทำงานอยู่' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (TEAM_KEY === 'CHANGE-ME' || RA_KEY === 'CHANGE-ME-RA' || TEAM_KEY === RA_KEY)
      return out({ ok: false, code: 'no_key', error: 'ยังไม่ได้ตั้งรหัสใน Apps Script (TEAM_KEY และ RA_KEY ต้องตั้งและต้องไม่ซ้ำกัน)' });
    const role = req.key === RA_KEY ? 'ra' : req.key === TEAM_KEY ? 'buy' : null;
    if (!role) return out({ ok: false, code: 'bad_key', error: 'รหัสไม่ถูกต้อง' });

    // อัปโหลดไฟล์ไม่ต้องล็อก (ไม่แตะข้อมูลที่ใช้ร่วมกัน)
    if (req.action === 'upload') return out({ ok: true, file: upload(req) });

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (req.action === 'list') return out({ ok: true, role: role, items: listItems() });
      if (req.action === 'save') return out({ ok: true, item: saveItem(req.item, role) });
      if (req.action === 'delete') { deleteItem(req.id); return out({ ok: true }); }
      return out({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + req.action });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out({ ok: false, error: String((err && err.message) || err) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Sheet สรุปรายการ ---------- */
function sheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('สรุปรายการสินค้า');
    DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(ROOT_ID));
    props.setProperty('SHEET_ID', ss.getId());
    const sh = ss.getSheets()[0];
    sh.setName('รายการ');
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#e3f4e8');
  }
  return ss.getSheets()[0];
}

function listItems() {
  const sh = sheet();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, HEAD.length, n - 1, 1).getValues()
    .map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } })
    .filter(Boolean);
}

function findRow(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

/* ---------- บันทึกสินค้า ---------- */
function saveItem(item, role) {
  if (!item || !item.id) throw new Error('ไม่มีรหัสสินค้า');
  const shx = sheet();
  const rx = findRow(shx, item.id);
  let prev = null;
  if (rx > 0) { try { prev = JSON.parse(shx.getRange(rx, HEAD.length).getValue()); } catch (e) { prev = null; } }
  const was = prev ? prev.status : 'draft';
  if (role === 'buy') {
    // จัดซื้อ: แก้ได้เฉพาะร่าง/งานที่ถูก Reject และส่งได้แค่ "ร่าง" หรือ "ส่งให้ RA"
    if (was !== 'draft' && was !== 'rejected') throw new Error('งานนี้ส่งให้ RA แล้ว แก้ไขไม่ได้');
    if (item.status !== 'draft' && item.status !== 'submitted' && !(item.status === 'rejected' && was === 'rejected'))
      throw new Error('เฉพาะ RA เท่านั้นที่ Approve หรือ Reject ได้');
  } else if (prev && (item.status === 'approved' || item.status === 'rejected')) {
    // RA ตัดสิน: เปลี่ยนได้แค่สถานะและผลตรวจ เนื้อหาที่จัดซื้อส่งมาคงเดิม
    if (was !== 'submitted') throw new Error('งานนี้ไม่ได้อยู่ในสถานะรอ RA ตรวจ');
    const keep = { status: item.status, rejectReason: item.rejectReason || '', rejectSections: item.rejectSections || [],
      reviewedBy: item.reviewedBy || '', reviewedAt: item.reviewedAt || Date.now(), history: item.history || prev.history };
    item = Object.assign({}, prev, keep);
  }
  const name = String(item.name || '').trim();
  item.updatedAt = Date.now();

  // โฟลเดอร์ของสินค้า
  let folder = null;
  if (item.folder && item.folder.id) { try { folder = DriveApp.getFolderById(item.folder.id); } catch (e) { folder = null; } }
  if (!folder && name) {
    folder = DriveApp.getFolderById(ROOT_ID).createFolder(name);
    item.folder = { id: folder.getId(), url: folder.getUrl(), name: name };
  }
  if (folder && name && folder.getName() !== name) { folder.setName(name); item.folder.name = name; }

  // เอกสารสรุปตอนส่งให้ RA (1 ไฟล์ต่อการส่ง 1 ครั้ง)
  const ver = (item.history || []).filter(h => h.action === 'submit').length;
  if (folder && item.status === 'submitted' && ver && (!item.drive || item.drive.ver !== ver)) {
    item.drive = makeDoc(item, folder, ver);
  }
  // ติดสถานะที่ชื่อไฟล์สรุปเมื่อ RA ตัดสิน
  if (item.drive && (item.status === 'approved' || item.status === 'rejected') && item.drive.mark !== item.status) {
    try {
      const f = DriveApp.getFileById(item.drive.id);
      f.setName((item.status === 'approved' ? '[อนุมัติ] ' : '[REJECT] ') + item.drive.title);
      f.setDescription(item.status === 'rejected' ? 'RA Reject: ' + item.rejectReason : 'RA อนุมัติ ' + fmt(Date.now()));
      item.drive.mark = item.status;
    } catch (e) { /* ไฟล์สรุปถูกลบไปแล้ว */ }
  }

  const json = JSON.stringify(item);
  if (folder) {
    const fs = folder.getFilesByName('data.json');
    if (fs.hasNext()) fs.next().setContent(json);
    else folder.createFile('data.json', json, 'application/json');
  }

  const sh = sheet();
  const files = k => ((item.files && item.files[k]) || []).map(f => f.name).join(', ');
  const sum = (item.ingredients || []).reduce((a, r) => a + (parseFloat(r.pct) || 0), 0);
  const row = [item.id, name, STATUS[item.status] || item.status, item.createdBy || '', fmt(item.updatedAt),
    Math.round(sum * 100) / 100, files('ing'), files('nut'), files('flow'), item.reviewedBy || '',
    item.status === 'rejected' ? (item.rejectReason || '') : '',
    item.folder ? item.folder.url : '', item.drive ? item.drive.url : '', json];
  const r = findRow(sh, item.id);
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return item;
}

function deleteItem(id) {
  const sh = sheet();
  const r = findRow(sh, id);
  if (r < 0) return;
  try {
    const item = JSON.parse(sh.getRange(r, HEAD.length).getValue());
    if (item.status !== 'draft') throw new Error('ลบได้เฉพาะร่าง');
    if (item.folder && item.folder.id) DriveApp.getFolderById(item.folder.id).setTrashed(true); // ย้ายไปถังขยะ กู้คืนได้
  } catch (e) {
    if (String(e.message) === 'ลบได้เฉพาะร่าง') throw e;
  }
  sh.deleteRow(r);
}

/* ---------- อัปโหลดไฟล์ ---------- */
function upload(req) {
  if (!SEC[req.sec]) throw new Error('ส่วนของไฟล์ไม่ถูกต้อง');
  const ok = /^(application\/pdf|image\/(png|jpeg|gif|webp|heic)|text\/plain)$/.test(req.mimeType);
  if (!ok) throw new Error('รับเฉพาะ PDF รูปภาพ หรือ .txt');
  const folder = DriveApp.getFolderById(req.folderId);
  const bytes = Utilities.base64Decode(req.data);
  const blob = Utilities.newBlob(bytes, req.mimeType, SEC[req.sec] + ' - ' + req.fileName);
  const f = folder.createFile(blob);
  return { id: f.getId(), url: f.getUrl(), name: req.fileName, type: req.mimeType, size: bytes.length, at: Date.now(), by: req.by || '' };
}

/* ---------- เอกสารสรุป ---------- */
function makeDoc(item, folder, ver) {
  const title = item.name + ' v' + ver + ' สรุปข้อมูล (ส่ง RA ' + fmt(item.submittedAt || Date.now(), 'd/M/yyyy') + ')';
  const doc = DocumentApp.create(title);
  const b = doc.getBody();
  const H = DocumentApp.ParagraphHeading;
  b.appendParagraph(item.name).setHeading(H.TITLE);
  b.appendParagraph('ส่งโดย ' + (item.createdBy || '-') + ' · ' + fmt(item.submittedAt || Date.now()));

  b.appendParagraph('Ingredient').setHeading(H.HEADING2);
  fileList(b, item, 'ing');
  const ings = (item.ingredients || []).slice().sort((a, c) => (parseFloat(c.pct) || 0) - (parseFloat(a.pct) || 0));
  const sum = ings.reduce((a, r) => a + (parseFloat(r.pct) || 0), 0);
  const rows = [['#', 'วัตถุดิบ', '%']]
    .concat(ings.map((r, i) => [String(i + 1), String(r.name), (parseFloat(r.pct) || 0).toFixed(2)]))
    .concat([['', 'รวม', sum.toFixed(2)]]);
  const t = b.appendTable(rows);
  t.getRow(0).editAsText().setBold(true);
  t.getRow(rows.length - 1).editAsText().setBold(true);

  b.appendParagraph('Nutrition').setHeading(H.HEADING2);
  fileList(b, item, 'nut');
  b.appendParagraph('Process Flow').setHeading(H.HEADING2);
  fileList(b, item, 'flow');
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);
  return { id: file.getId(), url: doc.getUrl(), title: title, ver: ver, mark: '' };
}

function fileList(b, item, k) {
  const fs = (item.files && item.files[k]) || [];
  if (!fs.length) { b.appendParagraph('(ไม่มีไฟล์)'); return; }
  fs.forEach(f => b.appendListItem(f.name).setLinkUrl(f.url));
}

function fmt(t, pattern) {
  return Utilities.formatDate(new Date(t), TZ, pattern || 'd/M/yyyy HH:mm');
}

/** กดรันฟังก์ชันนี้ 1 ครั้งใน Apps Script เพื่ออนุญาตสิทธิ์ Drive / Sheets / Docs และสร้าง Sheet สรุป */
function setup() {
  sheet();
  DriveApp.getFolderById(ROOT_ID).getName();
  Logger.log('พร้อมใช้งาน');
}
