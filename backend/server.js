// PM-GOV1 Progress Dashboard API
// รับ "ไฟล์ Excel ทั้งไฟล์" จาก Make.com (multipart/form-data, field name = "file")
// แล้วแกะ (parse) เองฝั่ง server จาก sheet "DATA" - เหมือน pattern เดียวกับ CNX/HDY/CEI
// Deploy URL: https://pm-gov1-progress-api.onrender.com
// Make.com scenario webhook (trigger): https://hook.eu1.make.com/m58fg8qhrhydcl5nru3md2dumc5c9xqu
//
// ตั้งค่า module HTTP (Make a request) ใน Make.com:
//   URL: https://pm-gov1-progress-api.onrender.com/api/webhook/excel
//   Method: POST
//   Body type: Multipart/form-data
//   Field: Key = "file", Value = {{ไฟล์ Excel จาก module ก่อนหน้า}}  (เช่น {{1.data}})

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// --- ผู้ใช้งาน (login) ---
// ตั้งค่าจริงผ่าน environment variable AUTH_USERS บน Render (JSON array บรรทัดเดียว), เช่น:
//   [
//     {"username":"oat","name":"OAT","role":"admin","passwordHash":"$2b$10$..."},
//     {"username":"antika","name":"Antika Prasadsil","pmName":"Antika Prasadsil","passwordHash":"$2b$10$..."}
//   ]
//
// การมองเห็นข้อมูล:
//   - role:"admin"  หรือ ไม่ได้ใส่ pmName  -> เห็นข้อมูลทุกโครงการของทุก PM
//   - ใส่ pmName    -> เห็นเฉพาะโครงการที่ตัวเองเป็น PM เท่านั้น (กรองที่ server ก่อนส่งออก)
//   ค่า pmName ต้องตรงกับคอลัมน์ "PM Name" ในไฟล์ Excel (ระบบเทียบแบบไม่สนตัวพิมพ์เล็ก/ใหญ่
//   และไม่สนช่องว่างซ้ำ เช่น "Prapasiri   Rakkanpat" กับ "Prapasiri Rakkanpat" ถือว่าคนเดียวกัน)
//
// สร้าง hash รหัสผ่านได้จาก: node generate-hash.js "รหัสผ่านจริง"
// ⚠️ ถ้าไม่ตั้ง AUTH_USERS จะใช้ค่าเริ่มต้นด้านล่าง (username: admin / password: pmgov1) - เปลี่ยนก่อนใช้งานจริงเสมอ
let USERS;
try { USERS = JSON.parse(process.env.AUTH_USERS || 'null'); } catch { USERS = null; }
if (!Array.isArray(USERS) || !USERS.length) {
  USERS = [{ username: 'admin', name: 'PM-GOV1 Admin', role: 'admin', passwordHash: '$2b$10$bqSxasX72zv/WjwFMOYzden5FnVDWsATMQVD29KFMC3nvmWIhHLOi' }];
  console.log('⚠ AUTH_USERS ไม่ได้ตั้งค่า - ใช้บัญชีเริ่มต้น admin/pmgov1 (ควรเปลี่ยนก่อนใช้งานจริง)');
}

// เทียบชื่อ PM แบบยืดหยุ่น: ตัดช่องว่างซ้ำ/หัวท้าย และไม่สนตัวพิมพ์เล็กใหญ่
const normName = (s) => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();

// ทีม PM-GOV1 (5 คน) - sheet PaymentW/PaymentM/Stock Movement/SumStockM เป็นข้อมูลทั้งบริษัท (ทุกทีม)
// ต้อง "ล็อก" ให้เหลือแค่ 5 คนนี้เสมอ ไม่ว่าใครจะ login เข้ามา (รวมถึง admin) เพราะ Tab การเงิน/สต็อก
// ต้องไม่โชว์ข้อมูลของทีมอื่นเด็ดขาดตามที่ตกลงกันไว้
const TEAM_PM_NAMES = ['Antika Prasadsil', 'Lomdetch Puangsombut', 'Virojt Changyencham', 'Prapasiri Rakkanpat', 'Watanyu Anantakunakorn'];
const TEAM_SET = new Set(TEAM_PM_NAMES.map(normName));
const inTeam = (name) => TEAM_SET.has(normName(name));

const tokens = new Map(); // token -> { username, name, expires }
const TOKEN_TTL = 12 * 60 * 60 * 1000; // 12 ชั่วโมง

function requireAuth(req, res, next) {
  const h = req.headers['authorization'] || '';
  const token = h.startsWith('Bearer ') ? h.slice(7) : null;
  const entry = token ? tokens.get(token) : null;
  if (!entry || entry.expires < Date.now()) {
    if (token) tokens.delete(token);
    return res.status(401).json({ ok: false, error: 'Unauthorized - please sign in again' });
  }
  entry.expires = Date.now() + TOKEN_TTL; // ใช้งานต่อเนื่อง = ต่ออายุ session ให้อัตโนมัติ
  req.user = { username: entry.username, name: entry.name, pmName: entry.pmName || null, role: entry.role || 'pm' };
  next();
}

// --- เข้าสู่ระบบ ---
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ ok: false, error: 'Username and password are required' });

  const user = USERS.find(u => u.username.toLowerCase() === String(username).toLowerCase());
  if (!user) return res.status(401).json({ ok: false, error: 'Incorrect username or password' });

  const match = await bcrypt.compare(String(password), user.passwordHash);
  if (!match) return res.status(401).json({ ok: false, error: 'Incorrect username or password' });

  const token = crypto.randomBytes(24).toString('hex');
  tokens.set(token, { username: user.username, name: user.name || user.username, pmName: user.pmName || null, role: user.role || 'pm', expires: Date.now() + TOKEN_TTL });
  console.log(`✓ Login: ${user.username}`);
  res.json({ ok: true, token, name: user.name || user.username, scope: (user.role === 'admin' || !user.pmName) ? 'all' : user.pmName, expiresIn: TOKEN_TTL });
});

app.post('/api/logout', requireAuth, (req, res) => {
  const token = (req.headers['authorization'] || '').slice(7);
  tokens.delete(token);
  res.json({ ok: true });
});

// ให้ frontend เช็คว่า token ที่เก็บไว้ยังใช้ได้ไหม (ตอนเปิดหน้าเว็บใหม่)
app.get('/api/me', requireAuth, (req, res) => {
  res.json({ ok: true, user: req.user });
});

// เก็บ snapshot ล่าสุดไว้ใน memory
let latestData = {
  rows: [],
  revenue: [],
  paymentW: [],
  po: [],
  stock: [],
  stockSummary: [],
  projectInfo: [],
  updatedAt: null,
};

// เก็บข้อมูล debug ล่าสุดไว้เสมอ ดูผ่าน GET /api/debug/last-payload
let lastRaw = {
  receivedAt: null,
  contentType: null,
  mode: null,       // 'file' | 'json'
  sheetName: null,
  fileSize: null,
  preview: null,
};

// --- แกะไฟล์ Excel: sheet ความคืบหน้าโครงการ ---
// อ่าน sheet ชื่อ "DATA"/"PROGRESS1" (หรือ sheet แรกถ้าหาไม่เจอ) โดยแถวที่ 2 = header, แถวที่ 3 เป็นต้นไป = ข้อมูล
function parseProgressSheet(wb) {
  const KNOWN_NAMES = ['DATA', 'PROGRESS1'];
  const sheetName = wb.SheetNames.find(n => KNOWN_NAMES.includes(n.trim().toUpperCase())) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null, dateNF: 'yyyy-mm-dd' });

  const headerRow = aoa[0] || [];        // แถวแรกของ used-range (= แถว 2 ใน Excel เพราะแถว 1 ว่าง)
  const dataRows = aoa.slice(1);         // แถวถัดไปทั้งหมด (= แถว 3 เป็นต้นไปใน Excel)

  const rows = dataRows
    .filter(r => Array.isArray(r) && r.some(v => v !== null && v !== ''))
    .map(r => {
      const obj = {};
      headerRow.forEach((h, i) => {
        if (h === null || h === undefined || h === '') return;
        obj[h] = r[i] !== undefined ? r[i] : null;
      });
      return obj;
    });

  return { rows, sheetName };
}

// --- แกะไฟล์ Excel: sheet "Project Info" (ประวัติโครงการทั้งหมดย้อนหลังหลายปี โครงสร้างคอลัมน์เหมือน Progress1) ---
function parseProjectInfoSheet(wb) {
  const sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === 'PROJECT INFO');
  if (!sheetName) return { rows: [], sheetName: null };
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null, dateNF: 'yyyy-mm-dd' });

  const headerRow = aoa[0] || [];
  const dataRows = aoa.slice(1);

  const rows = dataRows
    .filter(r => Array.isArray(r) && r.some(v => v !== null && v !== ''))
    .map(r => {
      const obj = {};
      headerRow.forEach((h, i) => {
        if (h === null || h === undefined || h === '') return;
        obj[h] = r[i] !== undefined ? r[i] : null;
      });
      return obj;
    });

  return { rows, sheetName };
}

// --- แกะไฟล์ Excel: sheet "PaymentW" (เก็บเงินจากลูกค้า อัปเดตรายสัปดาห์) ---
// โครงสร้างพิเศษ: แถวโครงการ (มีเลขที่ No. + Project Code) ตามด้วยแถวย่อยรายงวด/ใบแจ้งหนี้
// (ไม่มี No./Project Code ซ้ำ ใช้ค่าคอลัมน์ D เป็นคำอธิบายงวดแทนชื่อโครงการ) สลับกับแถวหัวข้อ section
// (เช่น "TEAM 1", "Performed by PM") ที่ต้องข้ามทิ้ง
function parsePaymentW(wb) {
  const sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === 'PAYMENTW');
  if (!sheetName) return { rows: [], sheetName: null };
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null, dateNF: 'yyyy-mm-dd' });

  const hIdx = aoa.findIndex(r => Array.isArray(r) && r.includes('Project Code'));
  if (hIdx < 0) return { rows: [], sheetName };
  const dataRows = aoa.slice(hIdx + 1);

  const projects = [];
  let cur = null;
  for (const r of dataRows) {
    if (!Array.isArray(r) || r.every(v => v === null || v === '')) continue;
    const no = r[0], code = r[1];
    const isProjectRow = (typeof no === 'number' || (typeof no === 'string' && no.trim() !== '' && !isNaN(no))) && code;
    if (isProjectRow) {
      cur = {
        projectCode: String(code).trim(), customer: r[2], projectName: r[3], accType: r[4],
        contractNo: r[5], startDate: r[6], endedDate: r[7], creditTerm: r[8],
        sellingPrice: r[9], remainder: r[10], invoiced: r[12], collecting: r[13],
        collected: r[15], sale: r[16], pm: r[17], remark: r[18], installments: [],
      };
      projects.push(cur);
    } else if (cur && (r[3] || r[12] != null || r[15] != null)) {
      // แถวย่อย: งวด/ใบแจ้งหนี้ของโครงการปัจจุบัน - ข้ามถ้าเป็นแถวหัวข้อ section (col A มีตัวหนังสือ, col B ว่าง)
      if (typeof no === 'string' && !code) continue;
      cur.installments.push({
        desc: r[3], accType: r[4], invoiceDate: r[11], invoiced: r[12],
        collecting: r[13], collectedDate: r[14], collected: r[15],
      });
    }
  }
  return { rows: projects, sheetName };
}

// --- แกะไฟล์ Excel: sheet "PaymentM" เฉพาะตาราง PO (ฝั่งขวา คอลัมน์ T เป็นต้นไป) ---
// ตารางนี้เป็นรายการ PO ที่สั่งซื้อ/ว่าจ้างผู้รับเหมา-vendor ต่อโครงการ ใช้ track ว่าจ่าย/รับของครบหรือยัง
function parsePaymentMPO(wb) {
  const sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === 'PAYMENTM');
  if (!sheetName) return { rows: [], sheetName: null };
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null, dateNF: 'yyyy-mm-dd' });

  const hIdx = aoa.findIndex(r => Array.isArray(r) && r.some(v => String(v || '').replace(/\s+/g, ' ').trim() === 'PROJECT CODE'));
  if (hIdx < 0) return { rows: [], sheetName };
  const header = aoa[hIdx];
  const startCol = header.findIndex(v => String(v || '').replace(/\s+/g, ' ').trim() === 'PROJECT CODE');
  const dataRows = aoa.slice(hIdx + 1);

  const KEYS = ['projectCode','poNumber','approved','status','quotation','vendor','bup','item','price','poDate','paymentTerm','approvedDate','deliveryDate','product','category','customer','arrivalDate','receivedPrice','balanced','cur','remark'];
  const rows = dataRows
    .map(r => (Array.isArray(r) ? r.slice(startCol, startCol + KEYS.length) : []))
    .filter(r => r[0] !== null && r[0] !== undefined && r[0] !== '')
    .map(r => { const o = {}; KEYS.forEach((k,i)=>{ o[k] = r[i] !== undefined ? r[i] : null; }); o.projectCode = String(o.projectCode).trim(); return o; });

  return { rows, sheetName };
}

// --- แกะไฟล์ Excel: sheet "Stock Movement" (ของค้างสต็อก ระดับรายชิ้น) ---
function parseStockMovement(wb) {
  const sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === 'STOCK MOVEMENT');
  if (!sheetName) return { rows: [], sheetName: null };
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null, dateNF: 'yyyy-mm-dd' });

  const hIdx = aoa.findIndex(r => Array.isArray(r) && r.includes('PCODE'));
  if (hIdx < 0) return { rows: [], sheetName };
  const dataRows = aoa.slice(hIdx + 1);

  const KEYS = ['no','pcode','partNumber','description','pur','hasSN','zone','shelf','poNo','receiveNo','invNo','rest','unitPrice','rate','cur','unitPriceBaht','totalPriceBaht','stockValueBaht','inDate','daysIn','pmName'];
  const rows = dataRows
    .filter(r => Array.isArray(r) && r[1] !== null && r[1] !== undefined && r[1] !== '')
    .map(r => { const o = {}; KEYS.forEach((k,i)=>{ o[k] = r[i] !== undefined ? r[i] : null; }); o.pcode = String(o.pcode).trim(); return o; });

  return { rows, sheetName };
}

// --- แกะไฟล์ Excel: sheet "SumStockM" (สรุปของค้างสต็อก ระดับโครงการ ต่อ PM) ---
function parseSumStockM(wb) {
  const sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === 'SUMSTOCKM');
  if (!sheetName) return { rows: [], sheetName: null };
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null, dateNF: 'yyyy-mm-dd' });

  const hIdx = aoa.findIndex(r => Array.isArray(r) && r.includes('Project Code'));
  if (hIdx < 0) return { rows: [], sheetName };
  const dataRows = aoa.slice(hIdx + 2); // แถวถัดไปเป็น sub-header ของคอลัมน์ P จึงข้าม 2 แถว

  const KEYS = ['no','projectCode','projectName','customer','pm','progressPct','contractValue','stockValuePrev','stockValueCurr','diffStockValue','startDate','endDate','contractDays','stockAgeDays','firstInDate','poStatus','stockRemarkCurr','pmRemarkCurr','pmRemarkPrev'];
  const rows = dataRows
    .filter(r => Array.isArray(r) && r[1] !== null && r[1] !== undefined && r[1] !== '')
    .map(r => { const o = {}; KEYS.forEach((k,i)=>{ o[k] = r[i] !== undefined ? r[i] : null; }); o.projectCode = String(o.projectCode).trim(); return o; });

  return { rows, sheetName };
}

// --- แกะไฟล์ Excel: sheet "Revenue" ---
// โครงสร้างต่างจาก DATA: header 2 แถวซ้อนกัน (แถว 3 = หมวดหลัก/ไตรมาส, แถว 4 = เดือนย่อยใต้แต่ละไตรมาส)
// ข้อมูลเริ่มแถว 5 เป็นต้นไป - แปลงเป็น key เดียวโดยรวมบริบทไตรมาสเข้ากับ "Total" ย่อยกันชนกัน
// (Q1_Total, Q2_Total, ...) ส่วนเดือน (Jan..Dec) ใช้ชื่อเดือนตรง ๆ เพราะไม่ซ้ำกันอยู่แล้ว
function parseRevenueSheet(wb) {
  const sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === 'REVENUE');
  if (!sheetName) return { rows: [], sheetName: null };

  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null, dateNF: 'yyyy-mm-dd' });

  // หาแถว header คู่ (แถวที่มี "Project Code") แล้วแถวถัดไปคือ sub-header เดือน
  let h3Idx = aoa.findIndex(r => Array.isArray(r) && r.includes('Project Code'));
  if (h3Idx < 0) return { rows: [], sheetName };
  const h3 = aoa[h3Idx] || [];
  const h4 = aoa[h3Idx + 1] || [];
  const dataRows = aoa.slice(h3Idx + 2);

  const keys = [];
  let currentQuarter = null;
  const width = Math.max(h3.length, h4.length);
  for (let i = 0; i < width; i++) {
    const v3 = h3[i], v4 = h4[i];
    if (v3 && /^Q[1-4]$/.test(String(v3).trim())) currentQuarter = String(v3).trim();
    if (v4) {
      keys[i] = String(v4).trim() === 'Total' ? `${currentQuarter}_Total` : String(v4).trim();
    } else if (v3) {
      keys[i] = String(v3).trim();
    } else {
      keys[i] = null;
    }
  }

  const rows = dataRows
    .filter(r => Array.isArray(r) && r.some(v => v !== null && v !== ''))
    .map(r => {
      const obj = {};
      keys.forEach((k, i) => {
        if (!k) return;
        obj[k] = r[i] !== undefined ? r[i] : null;
      });
      return obj;
    })
    .filter(o => o['Project Code']); // ตัดแถวว่าง/แถวสรุปที่ไม่มีรหัสโครงการทิ้ง

  return { rows, sheetName };
}

// --- รับข้อมูลจาก Make.com ---
// รองรับ 2 ทาง: (1) ไฟล์ Excel จริงแบบ multipart (ทางหลัก) (2) JSON array ตรง ๆ (ทางสำรอง เผื่อ mapping ฝั่ง Make.com เอง)
app.post('/api/webhook/excel', upload.single('file'), (req, res) => {
  try {
    if (req.file) {
      const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
      const { rows, sheetName } = parseProgressSheet(wb);
      const { rows: revenue, sheetName: revenueSheetName } = parseRevenueSheet(wb);
      const { rows: paymentW, sheetName: paymentWSheetName } = parsePaymentW(wb);
      const { rows: po, sheetName: poSheetName } = parsePaymentMPO(wb);
      const { rows: stock, sheetName: stockSheetName } = parseStockMovement(wb);
      const { rows: stockSummary, sheetName: stockSummarySheetName } = parseSumStockM(wb);
      const { rows: projectInfo, sheetName: projectInfoSheetName } = parseProjectInfoSheet(wb);

      // PO ไม่มีคอลัมน์ PM ตรง ๆ - เดา PM เจ้าของจาก Project Code โดยอ้างอิงจาก Progress1 ก่อน (ครอบคลุมสุด)
      // แล้ว fallback ไปที่ PaymentW ถ้า Progress1 ไม่มีโครงการนั้น (เช่นโครงการเก่าที่ปิดไปแล้ว)
      const pmByCode = {};
      rows.forEach(r => { if (r['Project Code'] && r['PM Name']) pmByCode[String(r['Project Code']).trim()] = r['PM Name']; });
      paymentW.forEach(p => { if (p.projectCode && p.pm && !pmByCode[p.projectCode]) pmByCode[p.projectCode] = p.pm; });
      po.forEach(p => { p.pm = pmByCode[p.projectCode] || null; });

      // ล็อกทั้ง 5 ชุดข้อมูลนี้ให้เหลือแค่ทีม PM-GOV1 (5 คน) เสมอ - sheet ต้นทางเป็นข้อมูลทั้งบริษัท/ทั้งบริษัทย้อนหลัง
      const paymentWTeam = paymentW.filter(p => inTeam(p.pm));
      const poTeam = po.filter(p => inTeam(p.pm));
      const stockTeam = stock.filter(s => inTeam(s.pmName));
      const stockSummaryTeam = stockSummary.filter(s => inTeam(s.pm));
      const projectInfoTeam = projectInfo.filter(p => inTeam(p['PM Name']));

      latestData = { rows, revenue, paymentW: paymentWTeam, po: poTeam, stock: stockTeam, stockSummary: stockSummaryTeam, projectInfo: projectInfoTeam, updatedAt: new Date().toISOString() };
      lastRaw = {
        receivedAt: latestData.updatedAt,
        contentType: req.headers['content-type'] || null,
        mode: 'file',
        sheetName,
        revenueSheetName,
        fileSize: req.file.buffer.length,
        preview: rows[0] ? JSON.stringify(rows[0]).slice(0, 500) : null,
        revenuePreview: revenue[0] ? JSON.stringify(revenue[0]).slice(0, 500) : null,
      };
      console.log(`✓ PM-GOV1 webhook: parsed ${rows.length} rows from "${sheetName}", ${revenue.length} revenue rows from "${revenueSheetName || '(not found)'}"`);
      console.log(`  + PaymentW ${paymentW.length}→${paymentWTeam.length} (team) from "${paymentWSheetName || '(not found)'}", PO ${po.length}→${poTeam.length} (team) from "${poSheetName || '(not found)'}", Stock ${stock.length}→${stockTeam.length} (team) from "${stockSheetName || '(not found)'}", StockSummary ${stockSummary.length}→${stockSummaryTeam.length} (team) from "${stockSummarySheetName || '(not found)'}" (${req.file.buffer.length} bytes)`);
      console.log(`  + Project Info ${projectInfo.length}→${projectInfoTeam.length} (team) from "${projectInfoSheetName || '(not found)'}"`);

      // เตือนถ้ามีชื่อ PM ในไฟล์ที่ยังไม่มีบัญชีผูกไว้ (พิมพ์ชื่อผิด หรือมี PM ใหม่เข้าทีม)
      const linked = new Set(USERS.filter(u => u.pmName).map(u => normName(u.pmName)));
      if (linked.size) {
        const unlinked = [...new Set(rows.map(r => r['PM Name']).filter(Boolean))]
          .filter(n => !linked.has(normName(n)));
        if (unlinked.length) {
          console.log(`⚠ พบชื่อ PM ในไฟล์ที่ยังไม่มีบัญชีผูกไว้: ${JSON.stringify(unlinked)}`);
          console.log('  (คนเหล่านี้จะยัง login เข้ามาดูข้อมูลตัวเองไม่ได้ จนกว่าจะเพิ่มใน AUTH_USERS)');
        }
      }

      return res.json({ ok: true, mode: 'file', sheet: sheetName, count: rows.length, revenueSheet: revenueSheetName, revenueCount: revenue.length, paymentWCount: paymentWTeam.length, poCount: poTeam.length, stockCount: stockTeam.length, stockSummaryCount: stockSummaryTeam.length, projectInfoCount: projectInfoTeam.length, updatedAt: latestData.updatedAt });
    }

    // ทางสำรอง: body เป็น JSON array ของแถวข้อมูลตรง ๆ
    const payload = req.body;
    const rows = Array.isArray(payload) ? payload : payload?.rows || payload?.data || null;
    if (rows) {
      latestData = { rows, revenue: payload?.revenue || [], updatedAt: new Date().toISOString() };
      lastRaw = {
        receivedAt: latestData.updatedAt,
        contentType: req.headers['content-type'] || null,
        mode: 'json',
        sheetName: null,
        fileSize: null,
        preview: JSON.stringify(payload).slice(0, 500),
      };
      console.log(`✓ PM-GOV1 webhook: parsed ${rows.length} rows from JSON body`);
      return res.json({ ok: true, mode: 'json', count: rows.length, updatedAt: latestData.updatedAt });
    }

    lastRaw = {
      receivedAt: new Date().toISOString(),
      contentType: req.headers['content-type'] || null,
      mode: 'unrecognized',
      sheetName: null,
      fileSize: null,
      preview: JSON.stringify(payload || {}).slice(0, 300),
    };
    return res.status(400).json({ ok: false, error: 'No file (field "file") or JSON rows found in request', received: lastRaw });
  } catch (e) {
    console.error('Webhook error:', e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

// --- ให้ frontend ดึงข้อมูลไปแสดง (ต้อง login ก่อน) ---
// ถ้าผู้ใช้ผูกกับ pmName ไว้ จะกรองให้เหลือเฉพาะโครงการของตัวเองตั้งแต่ฝั่ง server
// (ปลอดภัยกว่ากรองฝั่งเบราว์เซอร์ เพราะข้อมูลคนอื่นไม่เคยถูกส่งออกไปเลย)
app.get('/api/projects', requireAuth, (req, res) => {
  const { pmName, role } = req.user;
  if (role === 'admin' || !pmName) return res.json(latestData);

  const target = normName(pmName);
  const rows = latestData.rows.filter(r => normName(r['PM Name']) === target);
  const revenue = (latestData.revenue || []).filter(r => normName(r['PM Name']) === target);
  const paymentW = (latestData.paymentW || []).filter(r => normName(r.pm) === target);
  const po = (latestData.po || []).filter(r => normName(r.pm) === target);
  const stock = (latestData.stock || []).filter(r => normName(r.pmName) === target);
  const stockSummary = (latestData.stockSummary || []).filter(r => normName(r.pm) === target);
  // projectInfo ไม่กรองรายบุคคล - เป็น tab ภาพรวมทีม (leaderboard/ประวัติ site) ที่ทุกคนในทีมควรเห็นเหมือนกันหมด
  // (ยังล็อกไว้แค่ทีม PM-GOV1 ตั้งแต่ตอนอ่านไฟล์แล้ว ไม่มีข้อมูลทีมอื่นหลุดออกมาแน่นอน)
  const projectInfo = latestData.projectInfo || [];

  // ถ้าผูก pmName ไว้แล้วแต่ไม่เจอโครงการเลย มักเกิดจากชื่อใน Excel สะกดไม่ตรงกับที่ตั้งไว้
  if (latestData.rows.length && !rows.length) {
    const available = [...new Set(latestData.rows.map(r => r['PM Name']).filter(Boolean))];
    console.log(`⚠ ผู้ใช้ "${req.user.username}" ผูกกับ pmName="${pmName}" แต่ไม่พบโครงการใดเลย`);
    console.log(`  ชื่อ PM ที่มีอยู่จริงในไฟล์: ${JSON.stringify(available)}`);
  }

  res.json({ ...latestData, rows, revenue, paymentW, po, stock, stockSummary, projectInfo, scope: pmName });
});

app.get('/api/health', (req, res) => {
  res.json({ ok: true, rows: latestData.rows.length, updatedAt: latestData.updatedAt });
});

// --- Debug: ดูว่า Make.com ส่งอะไรมาล่าสุดจริง ๆ (ต้อง login ก่อน) ---
app.get('/api/debug/last-payload', requireAuth, (req, res) => {
  res.json(lastRaw);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`PM-GOV1 API listening on :${PORT}`);

  // --- Keep-alive: กันไม่ให้ Render free tier sleep (spin down หลังไม่มี request ~15 นาที) ---
  // ยิง request เข้าตัวเองทุก 10 นาที ผ่าน public URL ของ Render (RENDER_EXTERNAL_URL มีให้อัตโนมัติ
  // บน Render เท่านั้น - รันในเครื่อง/ที่อื่นจะไม่ทำงาน ไม่กระทบอะไร)
  const selfUrl = process.env.RENDER_EXTERNAL_URL;
  if (selfUrl) {
    setInterval(() => {
      fetch(`${selfUrl}/api/health`)
        .then(() => console.log(`♥ keep-alive ping ok @ ${new Date().toISOString()}`))
        .catch(err => console.log('keep-alive ping failed:', err.message));
    }, 10 * 60 * 1000); // ทุก 10 นาที (สั้นกว่า timeout ของ Render ที่ ~15 นาที)
    console.log(`♥ keep-alive enabled → pinging ${selfUrl}/api/health every 10 min`);
  } else {
    console.log('ℹ keep-alive skipped (RENDER_EXTERNAL_URL not set - not running on Render, or running locally)');
  }
});
