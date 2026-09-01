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
const fs = require('fs');
const path = require('path');

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
  res.json({ ok: true, token, username: user.username, name: user.name || user.username, scope: (user.role === 'admin' || !user.pmName) ? 'all' : user.pmName, expiresIn: TOKEN_TTL });
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
  projectInfoAll: [], // ไม่ล็อกทีม - ใช้เฉพาะ card "ค้นหา site/โครงการทั้งบริษัท" ที่ตั้งใจให้เห็นข้ามทีมได้
  evm: {}, // sheet EVM ละเอียดต่อโครงการ (key = Project Code) - มีเฉพาะโครงการที่ทีมทำ sheet ไว้ ไม่ครบทุกโครงการ
  updatedAt: null,
};

// การตั้งค่าว่าการ์ดไหนให้ใครเห็นได้บ้าง (OAT จัดการเองผ่าน Tab "Admin") - key = cardId, value = 'all' หรือ array
// ของ username ที่อนุญาต การ์ดที่ไม่มี entry ในนี้ = ค่า default 'all' (ทุกคนเห็น)
//
// ⚠️ ความถาวรของการตั้งค่านี้ขึ้นกับที่เก็บ hosting: ระบบนี้เขียนค่าลงไฟล์ CARD_VISIBILITY_FILE บน disk เสมอ
// (คนละกลไกกับ latestData/Excel ที่อยู่ใน memory ล้วน ๆ) ซึ่งจะอยู่ถาวรข้ามการ redeploy ได้ "ก็ต่อเมื่อ" service
// บน Render ผูก Persistent Disk ไว้ที่ path ของไฟล์นี้เท่านั้น - ถ้าไม่ได้ผูกไว้ (ปกติ default ของ Render คือ
// disk แบบ ephemeral หายทุกครั้งที่ redeploy) ค่าจะยังคงอยู่ระหว่างที่ server รันต่อเนื่อง/restart ปกติ
// แต่จะรีเซ็ตกลับเป็นค่าเริ่มต้นเมื่อ redeploy โค้ดใหม่เหมือนเดิม - ถ้าอยากถาวรจริงข้ามทุกการ redeploy
// ต้องไปเปิด Persistent Disk บน Render แล้วตั้ง mount path ให้ตรงกับโฟลเดอร์ของไฟล์นี้
const CARD_VISIBILITY_FILE = path.join(__dirname, 'card-visibility-store.json');
const DEFAULT_CARD_VISIBILITY = {
  teamScorecard: ['oat'], // ค่าเริ่มต้น: การ์ดนี้เคย hardcode ไว้ให้ oat เห็นคนเดียว - seed ไว้แบบเดิมก่อน OAT จะมาปรับเองทีหลังได้
};

function loadCardVisibility() {
  try {
    const raw = fs.readFileSync(CARD_VISIBILITY_FILE, 'utf8');
    const parsed = JSON.parse(raw);
    console.log(`✓ โหลดค่าสิทธิ์การ์ดจากไฟล์ ${CARD_VISIBILITY_FILE} สำเร็จ`);
    return parsed;
  } catch (e) {
    console.log(`ℹ ยังไม่มีไฟล์ค่าสิทธิ์การ์ด (${e.code === 'ENOENT' ? 'ไฟล์ยังไม่เคยถูกสร้าง' : e.message}) - ใช้ค่าเริ่มต้นไปก่อน`);
    return { ...DEFAULT_CARD_VISIBILITY };
  }
}

function saveCardVisibility(data) {
  try {
    fs.writeFileSync(CARD_VISIBILITY_FILE, JSON.stringify(data, null, 2), 'utf8');
    return true;
  } catch (e) {
    console.log(`⚠ เขียนไฟล์ค่าสิทธิ์การ์ดไม่สำเร็จ: ${e.message} (การตั้งค่าจะยังใช้ได้ในรอบรันนี้ แต่จะหายถ้า server restart)`);
    return false;
  }
}

let CARD_VISIBILITY = loadCardVisibility();
if (!fs.existsSync(CARD_VISIBILITY_FILE)) saveCardVisibility(CARD_VISIBILITY); // สร้างไฟล์ตั้งแต่ startup รอบแรกเลย ไม่ต้องรอจนกว่าจะมีคนแก้ค่าครั้งแรก

// การตั้งค่าที่สอง แยกจาก CARD_VISIBILITY ข้างบน (คนละคำถามกัน): "การ์ดที่เห็นอยู่แล้ว ให้เห็นข้อมูลแค่ไหน"
// key = cardId, value = 'team-wide' (ทุกคนที่เห็นการ์ดนี้ได้ เห็นข้อมูลทีมเต็ม ไม่ผ่านการกรองรายบุคคล) หรือไม่มี entry
// = ค่า default (แต่ละคนเห็นแค่โครงการของตัวเอง ตามปกติ) เก็บไฟล์แยกต่างหาก ถาวรแบบเดียวกับ CARD_VISIBILITY
const CARD_DATA_SCOPE_FILE = path.join(__dirname, 'card-data-scope-store.json');
const DEFAULT_CARD_DATA_SCOPE = {
  dataFreshness: 'team-wide', // ค่าเริ่มต้น: Update PMS เคย hardcode ไว้เป็นทีมเต็มเสมอ - seed ไว้แบบเดิมก่อน OAT จะมาปรับเองทีหลังได้
};
function loadCardDataScope() {
  try {
    const raw = fs.readFileSync(CARD_DATA_SCOPE_FILE, 'utf8');
    console.log(`✓ โหลดค่าขอบเขตข้อมูลการ์ดจากไฟล์ ${CARD_DATA_SCOPE_FILE} สำเร็จ`);
    return JSON.parse(raw);
  } catch (e) {
    console.log(`ℹ ยังไม่มีไฟล์ขอบเขตข้อมูลการ์ด (${e.code === 'ENOENT' ? 'ไฟล์ยังไม่เคยถูกสร้าง' : e.message}) - ใช้ค่าเริ่มต้นไปก่อน`);
    return { ...DEFAULT_CARD_DATA_SCOPE };
  }
}
function saveCardDataScope(data) {
  try { fs.writeFileSync(CARD_DATA_SCOPE_FILE, JSON.stringify(data, null, 2), 'utf8'); return true; }
  catch (e) { console.log(`⚠ เขียนไฟล์ขอบเขตข้อมูลการ์ดไม่สำเร็จ: ${e.message}`); return false; }
}
let CARD_DATA_SCOPE = loadCardDataScope();
if (!fs.existsSync(CARD_DATA_SCOPE_FILE)) saveCardDataScope(CARD_DATA_SCOPE);

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
        const key = String(h).trim(); // หัวคอลัมน์บางอันมีช่องว่างนำหน้า (เช่น " Project Code") ต้องตัดออกก่อนเสมอ
        obj[key] = r[i] !== undefined ? r[i] : null;
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
        const key = String(h).trim(); // หัวคอลัมน์บางอันมีช่องว่างนำหน้า (เช่น " Project Code") ต้องตัดออกก่อนเสมอ
        obj[key] = r[i] !== undefined ? r[i] : null;
      });
      return obj;
    });

  return { rows, sheetName };
}

// --- แกะไฟล์ Excel: sheet ที่ชื่อเป็นรหัสโครงการล้วน ๆ (เช่น "335261233") ---
// เป็น sheet EVM (Earned Value Management) ละเอียดที่ทีมทำเองสำหรับโครงการที่ต้องเจาะลึกเป็นพิเศษ
// ไม่ตายตัวว่าต้องมีกี่ sheet - สแกนหาทุก sheet ที่ชื่อเป็นตัวเลขล้วน แล้ว parse ให้หมด เผื่ออนาคตมีเพิ่ม
//
// สำคัญ: จำนวน phase และจำนวนสัปดาห์ "ไม่เท่ากัน" ในแต่ละโครงการ (พบจริง: โครงการที่มีของต้องส่งมี 7 phase/35 สัปดาห์
// ส่วนโครงการที่ปรึกษา/MA ไม่มีของส่งมีแค่ 5 phase/40 สัปดาห์) - parser นี้จึงต้องหาโครงสร้างแบบไดนามิกทั้งหมด
// ไม่ hardcode ตำแหน่งแถว/คอลัมน์เด็ดขาด: หาความยาวสัปดาห์จากแถวหัวตาราง, หา phase จาก pattern คู่ P/A ที่ต่อกัน,
// และหาแถวสรุป (PV/EV/SPI/Health) จากการค้นชื่อ label ในคอลัมน์ A แทนตำแหน่งตายตัว
function parseEvmSheets(wb, knownProjectCodes) {
  const evmByCode = {};
  wb.SheetNames.forEach(sheetName => {
    const trimmed = sheetName.trim();
    if (!/^\d+$/.test(trimmed)) return;           // ต้องเป็นตัวเลขล้วนเท่านั้น
    if (!knownProjectCodes.has(trimmed)) return;   // ต้องตรงกับโครงการที่มีจริงใน Progress1

    try {
      const ws = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      if (aoa.length < 15) return; // สั้นเกินกว่าจะเป็นโครงสร้าง EVM จริง ข้ามอย่างปลอดภัย

      const title = String(aoa[0]?.[0] || '');
      const projectName = title.includes('-') ? title.split('-').slice(1).join('-').trim() : title;

      // หาความยาวช่วงสัปดาห์แบบไดนามิก - เริ่มคอลัมน์ D (index 3) ไล่จนกว่าจะเจอ cell ว่าง
      const headerRow = aoa[1] || [];
      let numWeeks = 0;
      while (headerRow[3 + numWeeks] !== null && headerRow[3 + numWeeks] !== undefined) numWeeks++;
      if (numWeeks === 0) return;
      const weekLabels = headerRow.slice(3, 3 + numWeeks).map(v => v === null ? null : String(v));
      const latestActualCol = 3 + numWeeks; // คอลัมน์ถัดจากสัปดาห์สุดท้าย = ค่า actual ล่าสุดของแต่ละ phase เสมอ

      // หา phase แบบไดนามิก - ไล่จากแถวที่ 3 (index 2) เป็นคู่ P/A ต่อเนื่องกันไปเรื่อย ๆ จนกว่า pattern จะขาด
      const phases = [];
      let i = 2;
      while (i + 1 < aoa.length) {
        const pRow = aoa[i] || [], aRow = aoa[i + 1] || [];
        const pName = String(pRow[0] || '').trim();
        const pMarker = String(pRow[2] || '').trim();
        const aMarker = String(aRow[2] || '').trim();
        if (pMarker !== 'P' || aMarker !== 'A' || !pName) break;
        const planWeekly = pRow.slice(3, 3 + numWeeks).map(v => typeof v === 'number' ? v : null);
        let lastPlanWeekIdx = -1;
        planWeekly.forEach((v, idx) => { if (v !== null) lastPlanWeekIdx = idx; });
        const actualLatestRaw = aRow[latestActualCol];
        phases.push({
          name: pName,
          weight: typeof pRow[1] === 'number' ? pRow[1] : 0,
          lastPlanWeekIdx,
          actualLatest: typeof actualLatestRaw === 'number' ? actualLatestRaw : 0,
        });
        i += 2;
      }
      if (!phases.length) return;

      // หาแถวสรุปด้วยการค้นชื่อ label ในคอลัมน์ A แทนตำแหน่งตายตัว (ตำแหน่งเลื่อนไปตามจำนวน phase ที่ไม่เท่ากัน)
      const findRow = (label) => aoa.findIndex(r => r && String(r[0] || '').trim() === label);
      const pvCumIdx = findRow('% Accumulate Planned Value [PV]');
      const evCumIdx = findRow('% Accumulate Earn Value [EV]');
      const spiStatusIdx = findRow('Schedule Performance Index [SPI]');
      const healthIdx = findRow('PROJECT HEALTH');
      if (pvCumIdx < 0 || evCumIdx < 0 || spiStatusIdx < 0 || healthIdx < 0) return; // โครงสร้างไม่ตรงตามที่คาด ข้ามอย่างปลอดภัย

      const pvCumRow = aoa[pvCumIdx] || [];
      const evCumRow = aoa[evCumIdx] || [];
      const spiRow = aoa[spiStatusIdx + 1] || []; // แถวถัดจากแถวสถานะ SPI คือแถวตัวเลข SPI เสมอ (ไม่มี label ของตัวเอง)
      const healthRow = aoa[healthIdx] || [];

      const pvCum = pvCumRow.slice(3, 3 + numWeeks).map(v => typeof v === 'number' ? v : null);
      const evCumFull = evCumRow.slice(3, 3 + numWeeks).map(v => typeof v === 'number' ? v : null);
      let asOfWeekIdx = -1;
      evCumFull.forEach((v, idx) => { if (v !== null) asOfWeekIdx = idx; }); // คอลัมน์สุดท้ายที่มี Actual จริง = "ข้อมูล ณ วันนี้" ของชีตนี้

      const spiWeekly = spiRow.slice(3, 3 + numWeeks).map(v => typeof v === 'number' ? v : null);
      const healthWeekly = healthRow.slice(3, 3 + numWeeks).map(v => v === null ? null : String(v));

      evmByCode[trimmed] = {
        projectCode: trimmed,
        projectName,
        weekLabels,
        phases,
        pvCum,
        evCum: evCumFull,
        asOfWeekIdx,
        spiCurrent: asOfWeekIdx >= 0 ? spiWeekly[asOfWeekIdx] : null,
        healthCurrent: asOfWeekIdx >= 0 ? healthWeekly[asOfWeekIdx] : null,
      };
    } catch (e) {
      console.log(`⚠ อ่าน EVM sheet "${sheetName}" ไม่สำเร็จ ข้ามไป: ${e.message}`);
    }
  });
  return evmByCode;
}


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
        collecting: r[13], collectedDate: r[14], collected: r[15], remainder: r[10],
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

  // หาหัวตาราง PO - ยืดหยุ่นขึ้น: ตัด non-breaking space (\u00A0) ที่ Excel ชอบแทรก, ไม่สนตัวพิมพ์เล็กใหญ่
  // และยอมรับทั้ง "PROJECT CODE" / "PROJECTCODE" (บางไฟล์ไม่มีเว้นวรรค)
  const normHdr = v => String(v || '').replace(/[\s\u00A0]+/g, ' ').trim().toUpperCase();
  const isCodeHdr = v => { const t = normHdr(v); return t === 'PROJECT CODE' || t === 'PROJECTCODE'; };
  // ⚠ sheet นี้มี "Project Code" 2 จุด: ตารางซ้าย (งวดจ่าย) กับตาราง PO ฝั่งขวา
  // ต้องเลือกอันที่เป็นตาราง PO เท่านั้น ดูจากการมีหัว "PO NUMBER" อยู่ถัดไปในแถวเดียวกัน
  const isPoNumHdr = v => normHdr(v).replace(/\s/g,'') === 'PONUMBER';
  const hIdx = aoa.findIndex(r => Array.isArray(r) && r.some(isCodeHdr) && r.some(isPoNumHdr));
  if (hIdx < 0) {
    // ช่วย debug: บอกว่าหัวตารางที่เจอจริงมีอะไรบ้าง จะได้รู้ว่าไฟล์ต้นทางเปลี่ยนรูปแบบไปหรือไม่
    const preview = aoa.slice(0, 5).map((r, i) =>
      `    แถว ${i}: ${JSON.stringify((r || []).filter(v => v !== null && v !== '').slice(0, 12))}`
    ).join('\n');
    console.log(`⚠ อ่าน PO จาก sheet "${sheetName}" ไม่ได้ - หาหัวตาราง "PROJECT CODE" ไม่เจอ`);
    console.log(`   จำนวนแถวทั้งหมดใน sheet: ${aoa.length}`);
    console.log(`   ตัวอย่าง 5 แถวแรกที่อ่านได้:\n${preview}`);
    return { rows: [], sheetName };
  }
  const header = aoa[hIdx];
  // เลือกคอลัมน์ "Project Code" ที่อยู่ติดกับ "PO NUMBER" (ตาราง PO) ไม่ใช่ตัวแรกของแถว
  const poNumCol = header.findIndex(isPoNumHdr);
  let startCol = -1;
  for (let i = 0; i < header.length; i++) {
    if (isCodeHdr(header[i]) && i < poNumCol) startCol = i;   // ตัวสุดท้ายที่อยู่ก่อน PO NUMBER
  }
  if (startCol < 0) startCol = header.findIndex(isCodeHdr);
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

      // แยก try/catch ต่อ sheet - ถ้า sheet ไหนพังจะเสียแค่ sheet นั้น ไม่ล้มทั้งการ sync
      // (เดิมถ้า parser ตัวไหนโยน error จะตกไป catch ใหญ่ท้ายสุด ทำให้ latestData ไม่ถูกอัปเดตเลย
      //  = ข้อมูลทุก sheet หายหมดพร้อมกัน ซึ่งเป็นอาการเดียวกับที่เคยเจอตอนแก้ไฟล์ PaymentW)
      const failed = [];
      const safeParse = (label, fn, fallback) => {
        try { return fn(); }
        catch (e) {
          console.error(`⚠ parse "${label}" ล้มเหลว: ${e.message} - ข้ามไปใช้ค่าว่างแทน ส่วนอื่นยังทำงานต่อ`);
          failed.push(label);
          return fallback;
        }
      };

      const { rows, sheetName } = safeParse('Progress1', () => parseProgressSheet(wb), { rows: [], sheetName: null });
      const { rows: revenue, sheetName: revenueSheetName } = safeParse('Revenue', () => parseRevenueSheet(wb), { rows: [], sheetName: null });
      const { rows: paymentW, sheetName: paymentWSheetName } = safeParse('PaymentW', () => parsePaymentW(wb), { rows: [], sheetName: null });
      const { rows: po, sheetName: poSheetName } = safeParse('PaymentM/PO', () => parsePaymentMPO(wb), { rows: [], sheetName: null });
      const { rows: stock, sheetName: stockSheetName } = safeParse('Stock Movement', () => parseStockMovement(wb), { rows: [], sheetName: null });
      const { rows: stockSummary, sheetName: stockSummarySheetName } = safeParse('SumStockM', () => parseSumStockM(wb), { rows: [], sheetName: null });
      const { rows: projectInfo, sheetName: projectInfoSheetName } = safeParse('Project Info', () => parseProjectInfoSheet(wb), { rows: [], sheetName: null });

      // Progress1 คือแกนหลัก ถ้าอ่านไม่ได้เลยแปลว่าไฟล์ผิดรูปแบบจริง ๆ - ไม่ควรเขียนทับข้อมูลเดิมที่ยังดีอยู่
      if (!rows.length) {
        console.error('⚠ อ่าน Progress1 ไม่ได้เลย - ยกเลิกการอัปเดต เก็บข้อมูลเดิมไว้ (กันข้อมูลหายทั้งระบบ)');
        return res.status(400).json({ ok: false, error: 'อ่านข้อมูลความคืบหน้าโครงการไม่ได้ หรือไม่มีข้อมูล - ไม่อัปเดตข้อมูลเดิม', failedSheets: failed });
      }

      // PO ไม่มีคอลัมน์ PM ตรง ๆ - เดา PM เจ้าของจาก Project Code โดยอ้างอิงจาก Progress1 ก่อน (ครอบคลุมสุด)
      // แล้ว fallback ไปที่ PaymentW ถ้า Progress1 ไม่มีโครงการนั้น (เช่นโครงการเก่าที่ปิดไปแล้ว)
      const pmByCode = {};
      rows.forEach(r => { if (r['Project Code'] && r['PM Name']) pmByCode[String(r['Project Code']).trim()] = r['PM Name']; });
      paymentW.forEach(p => { if (p.projectCode && p.pm && !pmByCode[p.projectCode]) pmByCode[p.projectCode] = p.pm; });
      // trim ทั้ง 2 ฝั่งให้ตรงกัน - ฝั่งสร้าง key ใช้ .trim() อยู่แล้ว ถ้าฝั่ง lookup ไม่ trim ด้วย
      // แล้วไฟล์มีช่องว่างแฝงเมื่อไหร่ จะหา pm ไม่เจอทุกแถว -> inTeam(null) กรองทิ้งหมด -> PO หายทั้งตาราง
      po.forEach(p => { p.pm = pmByCode[String(p.projectCode || '').trim()] || null; });

      // ล็อกทั้ง 5 ชุดข้อมูลนี้ให้เหลือแค่ทีม PM-GOV1 (5 คน) เสมอ - sheet ต้นทางเป็นข้อมูลทั้งบริษัท/ทั้งบริษัทย้อนหลัง
      const paymentWTeam = paymentW.filter(p => inTeam(p.pm));
      const poTeam = po.filter(p => inTeam(p.pm));
      if (po.length && !poTeam.length) {
        const noPm = po.filter(p => !p.pm).length;
        const sample = [...new Set(po.map(p => p.pm).filter(Boolean))].slice(0, 5);
        console.log(`⚠ PO ${po.length} รายการถูกกรองทิ้งหมด - ไม่มีรายการไหนเป็นของทีม PM-GOV1`);
        console.log(`   หา PM ไม่เจอ ${noPm} รายการ | PM ที่พบในไฟล์: ${JSON.stringify(sample)}`);
        console.log(`   (PO ไม่มีคอลัมน์ PM ต้องเดาจาก Project Code เทียบกับ Progress1/PaymentW)`);
      }
      const stockTeam = stock.filter(s => inTeam(s.pmName));
      const stockSummaryTeam = stockSummary.filter(s => inTeam(s.pm));
      const projectInfoTeam = projectInfo.filter(p => inTeam(p['PM Name']));

      // Progress1 (rows) เป็นข้อมูลทีมเราอยู่แล้วโดยธรรมชาติ - ใช้รหัสโครงการจากตรงนี้เป็น whitelist
      // ให้ parseEvmSheets เจอเฉพาะ sheet ที่ตรงกับโครงการทีมเราเท่านั้น ไม่มีทางหลุดข้ามทีมได้
      const knownProjectCodes = new Set(rows.map(r => String(r['Project Code'] || '').trim()).filter(Boolean));
      const evm = parseEvmSheets(wb, knownProjectCodes);

      latestData = { rows, revenue, paymentW: paymentWTeam, po: poTeam, stock: stockTeam, stockSummary: stockSummaryTeam, projectInfo: projectInfoTeam, projectInfoAll: projectInfo, evm, updatedAt: new Date().toISOString() };
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
      console.log(`  + EVM sheets found: ${Object.keys(evm).length} (${Object.keys(evm).join(', ') || 'none'})`);

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

      if (failed.length) console.error(`⚠ sync สำเร็จบางส่วน - sheet ที่อ่านไม่ได้: ${failed.join(', ')}`);
      return res.json({ ok: true, mode: 'file', sheet: sheetName, count: rows.length, revenueSheet: revenueSheetName, revenueCount: revenue.length, paymentWCount: paymentWTeam.length, poCount: poTeam.length, stockCount: stockTeam.length, stockSummaryCount: stockSummaryTeam.length, projectInfoCount: projectInfoTeam.length, failedSheets: failed, updatedAt: latestData.updatedAt });
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
// ==== Admin: จัดการว่าการ์ดไหนให้ใครเห็นได้บ้าง (Tab "Admin" เรียกใช้ 2 ตัวนี้) ====
// เฉพาะ role admin เท่านั้นที่เรียกได้ (username เป็นเงื่อนไขเพิ่มเติมฝั่ง frontend สำหรับโชว์ tab แต่ backend
// เช็คแค่ role พอ เพราะในระบบนี้มีแค่ oat คนเดียวที่เป็น admin - ถ้าอนาคตมี admin คนอื่นเพิ่ม จะจัดการได้ด้วยเหมือนกัน)
app.get('/api/admin/card-visibility', requireAuth, (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ ok: false, error: 'ต้องเป็น admin เท่านั้น' });
  res.json({ ok: true, visibility: CARD_VISIBILITY });
});

app.post('/api/admin/card-visibility', requireAuth, (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ ok: false, error: 'ต้องเป็น admin เท่านั้น' });
  const { cardId, mode, users } = req.body || {};
  if (!cardId || typeof cardId !== 'string') return res.status(400).json({ ok: false, error: 'ไม่มี cardId' });

  if (mode === 'all') {
    delete CARD_VISIBILITY[cardId]; // ค่า default = ทุกคนเห็น ไม่ต้องเก็บ entry ไว้ก็ได้ ชัดเจนกว่าว่าเป็นค่าปกติ
  } else if (mode === 'custom') {
    CARD_VISIBILITY[cardId] = Array.isArray(users) ? users.filter(u => typeof u === 'string') : [];
  } else {
    return res.status(400).json({ ok: false, error: 'mode ต้องเป็น "all" หรือ "custom"' });
  }

  const persisted = saveCardVisibility(CARD_VISIBILITY); // เขียนลงไฟล์ทันทีทุกครั้งที่แก้ - ดู CARD_VISIBILITY_FILE ด้านบนเรื่องความถาวรจริง
  console.log(`✓ Admin (${req.user.username}) แก้สิทธิ์การ์ด "${cardId}" -> ${mode==='all' ? 'ทุกคนเห็น' : `เฉพาะ [${(CARD_VISIBILITY[cardId]||[]).join(', ')}]`} (เขียนไฟล์${persisted?'สำเร็จ':'ไม่สำเร็จ'})`);
  res.json({ ok: true, visibility: CARD_VISIBILITY, persisted });
});

// ==== Admin: ขอบเขตข้อมูลภายในการ์ดที่เห็นอยู่แล้ว (ทีมเต็ม vs เฉพาะของตัวเอง) - คนละเรื่องกับ card-visibility ด้านบน ====
app.get('/api/admin/card-data-scope', requireAuth, (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ ok: false, error: 'ต้องเป็น admin เท่านั้น' });
  res.json({ ok: true, dataScope: CARD_DATA_SCOPE });
});

app.post('/api/admin/card-data-scope', requireAuth, (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ ok: false, error: 'ต้องเป็น admin เท่านั้น' });
  const { cardId, teamWide } = req.body || {};
  if (!cardId || typeof cardId !== 'string') return res.status(400).json({ ok: false, error: 'ไม่มี cardId' });

  if (teamWide) CARD_DATA_SCOPE[cardId] = 'team-wide';
  else delete CARD_DATA_SCOPE[cardId]; // ค่า default = เห็นแค่ของตัวเอง ไม่ต้องเก็บ entry ไว้ก็ได้

  const persisted = saveCardDataScope(CARD_DATA_SCOPE);
  console.log(`✓ Admin (${req.user.username}) แก้ขอบเขตข้อมูลการ์ด "${cardId}" -> ${teamWide ? 'ทีมเต็ม' : 'เฉพาะของตัวเอง'} (เขียนไฟล์${persisted?'สำเร็จ':'ไม่สำเร็จ'})`);
  res.json({ ok: true, dataScope: CARD_DATA_SCOPE, persisted });
});


// ===================== ผู้ช่วย AI สำหรับการ์ด TOR Breakdown =====================
// ใช้ Groq (โมเดล open-weight) ช่วยเฉพาะจุดที่กฎทำไม่ได้ ไม่ได้ส่งเอกสารทั้งฉบับ
// เหตุผล: free tier จำกัด 6,000 tokens/นาที ถ้าส่งทั้งฉบับจะชนเพดานและต้องหั่นจนเสียภาพรวม
// จึงส่งเฉพาะ "บรรทัดสั้น ๆ ที่กฎจัดประเภทไม่ได้" ครั้งละไม่เกิน 60 รายการ
// API key เก็บที่ฝั่ง server เท่านั้น (env GROQ_API_KEY) ไม่ให้หลุดไปหน้าเว็บ
const GROQ_KEY = process.env.GROQ_API_KEY || '';
// Groq ปลดโมเดลเป็นระยะ (llama-3.3-70b-versatile ย้ายไป Enterprise-only เมื่อ 16 ส.ค. 2569)
// จึงไม่ล็อกชื่อโมเดลตายตัว แต่ถามรายการที่ใช้ได้จริงจาก Groq แล้วเลือกตามลำดับที่ต้องการ
// แกะ JSON จากคำตอบของโมเดล - เผื่อกรณีมีข้อความนำ, ครอบ ```, หรือมีส่วนคิด (reasoning) ปนมา
function parseLooseJson(txt){
  if(!txt) return null;
  let t = String(txt);
  // ตัดส่วนคิดของ reasoning model ออก
  t = t.replace(/<think>[\s\S]*?<\/think>/gi,'')
       .replace(/<reasoning>[\s\S]*?<\/reasoning>/gi,'')
       .replace(/```(?:json)?/gi,'')
       .trim();
  const tryParse = v => { try { return JSON.parse(v); } catch { return null; } };
  // ถ้าเป็นหลาย object เรียงกันโดยไม่มีวงเล็บครอบ (เช่น {..}\n{..}) ให้รวมเป็น array ก่อน
  const standalone=[...t.matchAll(/^\s*(\{[^{}]*\})\s*$/gm)].map(m=>m[1]);
  if(standalone.length>1){
    const arr=standalone.map(tryParse).filter(Boolean);
    if(arr.length===standalone.length) return arr;
  }
  let r = tryParse(t);
  if(r) return r;
  // หา array หรือ object ที่สมดุลวงเล็บ (เอาก้อนที่ยาวที่สุด)
  for(const [open,close] of [['[',']'],['{','}']]){
    let best=null;
    for(let i=0;i<t.length;i++){
      if(t[i]!==open) continue;
      let depth=0, inStr=false, esc=false;
      for(let k=i;k<t.length;k++){
        const c=t[k];
        if(esc){ esc=false; continue; }
        if(c==='\\'){ esc=true; continue; }
        if(c==='"'){ inStr=!inStr; continue; }
        if(inStr) continue;
        if(c===open) depth++;
        else if(c===close){
          depth--;
          if(depth===0){
            const cand=t.slice(i,k+1);
            if(!best || cand.length>best.length) best=cand;
            break;
          }
        }
      }
    }
    if(best){ r=tryParse(best); if(r) return r; }
  }
  // แถวสุดท้าย: เก็บ object ทีละตัวจากข้อความ
  const objs=[...t.matchAll(/\{[^{}]*\}/g)].map(m=>tryParse(m[0])).filter(Boolean);
  return objs.length ? objs : null;
}

const GROQ_BASE = process.env.GROQ_BASE || 'https://api.groq.com/openai/v1';
const GROQ_MODEL_PREF = (process.env.GROQ_MODEL || '').trim();
const GROQ_FALLBACKS = [
  'openai/gpt-oss-120b',      // แทน llama-3.3-70b - คุณภาพดีสุดในกลุ่ม self-serve
  'qwen/qwen3.6-27b',
  'openai/gpt-oss-20b',       // เร็วและถูกกว่า ใช้กับงานจัดประเภทได้ดี
  'llama-3.3-70b-versatile',
  'llama-3.1-8b-instant',
];
let GROQ_MODEL_CACHE = { id:null, at:0 };

async function groqPickModel(ignorePref){
  // ignorePref = true เมื่อโมเดลที่ระบุไว้ใช้ไม่ได้แล้ว (ถูกปลดระวาง) ต้องหาตัวแทนจากรายการจริง
  if(GROQ_MODEL_PREF && !ignorePref) return GROQ_MODEL_PREF;
  if(GROQ_MODEL_CACHE.id && Date.now()-GROQ_MODEL_CACHE.at < 3600e3) return GROQ_MODEL_CACHE.id;
  try{
    const r = await fetch(GROQ_BASE+'/models', {
      headers:{ 'Authorization':'Bearer '+GROQ_KEY }
    });
    if(r.ok){
      const j = await r.json();
      const ids = (j.data||[]).map(m=>m.id);
      const pick = GROQ_FALLBACKS.find(m=>ids.includes(m))
                || ids.find(m=>/gpt-oss|qwen|llama/i.test(m) && !/guard|whisper|tts|orpheus|prompt/i.test(m));
      if(pick){ GROQ_MODEL_CACHE={id:pick, at:Date.now()}; console.log('[TOR AI] เลือกโมเดล:', pick); return pick; }
    }
  }catch(e){ console.warn('[TOR AI] ถามรายการโมเดลไม่ได้:', e.message); }
  return GROQ_FALLBACKS[0];
}
const TOR_TYPES_LIST = 'Document, Development, Planning, Requirement, Integration, Reporting, Test, Installation, Migration, Removal, Procurement, Design, Survey, Training, Meeting, Warranty, Configuration, Compliance, Penalty, SLA, Spec, Other';

// จำกัดการใช้งานต่อคนต่อวัน กัน quota หมดโดยไม่ตั้งใจ
const torAiUsage = {};   // { 'user|YYYY-MM-DD': count }
const TOR_AI_DAILY_LIMIT = Number(process.env.TOR_AI_DAILY_LIMIT || 40);

app.get('/api/tor/ai-status', requireAuth, async (req, res) => {
  const day = new Date().toISOString().slice(0,10);
  const key = (req.user && req.user.username) + '|' + day;
  let model = null, available = [];
  if (GROQ_KEY) {
    try {
      const r = await fetch(GROQ_BASE+'/models', { headers:{ 'Authorization':'Bearer '+GROQ_KEY } });
      if (r.ok) { const j = await r.json(); available = (j.data||[]).map(m=>m.id); }
    } catch {}
    model = await groqPickModel();
  }
  res.json({ ok:true, enabled: !!GROQ_KEY, model, available,
             used: torAiUsage[key] || 0, limit: TOR_AI_DAILY_LIMIT });
});

app.post('/api/tor/classify', requireAuth, async (req, res) => {
  if (!GROQ_KEY) return res.status(503).json({ ok:false, error:'ยังไม่ได้ตั้งค่า GROQ_API_KEY ที่เซิร์ฟเวอร์' });
  const day = new Date().toISOString().slice(0,10);
  const key = (req.user && req.user.username) + '|' + day;
  if ((torAiUsage[key]||0) >= TOR_AI_DAILY_LIMIT)
    return res.status(429).json({ ok:false, error:`ใช้ครบโควตาวันนี้แล้ว (${TOR_AI_DAILY_LIMIT} ครั้ง)` });

  const items = Array.isArray(req.body && req.body.items) ? req.body.items.slice(0,20) : [];
  const task = (req.body && req.body.task) || 'classify';
  if (!items.length) return res.json({ ok:true, result: [] });

  // ตัดข้อความให้สั้นก่อนส่ง ประหยัด token และอยู่ในเพดาน 6,000 TPM
  const lines = items.map((x,i)=>`${i+1}. ${String(x).replace(/\s+/g,' ').slice(0,110)}`).join('\n');

  const prompts = {
    classify: `จัดประเภทงานในเอกสาร TOR ภาษาไทยต่อไปนี้ ให้เลือกประเภทที่เหมาะที่สุดจากรายการนี้เท่านั้น: ${TOR_TYPES_LIST}
ตอบเป็น JSON object อย่างเดียว ไม่ต้องอธิบาย รูปแบบ: {"items":[{"n":1,"type":"Document"},{"n":2,"type":"Test"}]}

${lines}`,
    hidden: `ต่อไปนี้เป็นข้อความจากหมวดคุณลักษณะเฉพาะ (สเปค) ในเอกสาร TOR ภาษาไทย
หาข้อที่ไม่ใช่แค่สเปคอุปกรณ์ แต่เป็น "ข้อผูกพันที่ผู้ขายต้องทำ" เช่น ต้องจัดหา ต้องอบรม ต้องรับประกัน ต้องส่งมอบ ต้องเป็นยี่ห้อเดียวกัน
ตอบเป็น JSON object อย่างเดียว: {"items":[{"n":1,"why":"เหตุผลสั้นๆ"}]} ถ้าไม่มีให้ตอบ {"items":[]}

${lines}`,
    focus: `ต่อไปนี้เป็นรายการงานของงวดงานหนึ่งในโครงการ (เอกสาร TOR ภาษาไทย)
สรุปสั้นๆ ว่า Project Manager ควรโฟกัสอะไรในงวดนี้ และมีความเสี่ยงอะไรที่มักตกหล่น
ตอบเป็น JSON อย่างเดียว: {"focus":"...","risk":"..."} ความยาวไม่เกินอย่างละ 2 บรรทัด

${lines}`
  };
  const prompt = prompts[task] || prompts.classify;

  try {
    let model = await groqPickModel();
    let r = await fetch(GROQ_BASE+'/chat/completions', {
      method:'POST',
      headers:{ 'Content-Type':'application/json', 'Authorization':'Bearer '+GROQ_KEY },
      body: JSON.stringify({
        model,
        temperature: 0.1,
        // reasoning model (gpt-oss) กิน token ไปกับส่วนคิดก่อนตอบ ถ้าตั้งต่ำเกินจะตอบไม่ทันแล้วได้ค่าว่าง
        max_tokens: 2000,
        reasoning_effort: 'low',        // ให้คิดน้อย ๆ เพราะงานนี้เป็นการจัดหมวด ไม่ต้องใช้เหตุผลลึก
        response_format: { type:'json_object' },
        messages: [
          { role:'system', content:'คุณเป็นผู้ช่วยวิเคราะห์เอกสาร TOR ของหน่วยงานราชการไทย ตอบเป็น JSON เท่านั้น ห้ามมีข้อความอื่น' },
          { role:'user', content: prompt }
        ]
      })
    });
    // โมเดลถูกปลดระวาง -> ล้าง cache แล้วลองตัวถัดไปที่ใช้ได้จริง
    if (r.status === 404) {
      GROQ_MODEL_CACHE = { id:null, at:0 };
      const alt = await groqPickModel(true);   // ข้ามค่าที่ผู้ใช้ระบุ เพราะพิสูจน์แล้วว่าใช้ไม่ได้
      if (alt && alt !== model) {
        console.log(`[TOR AI] ${model} ใช้ไม่ได้ เปลี่ยนเป็น ${alt}`);
        model = alt;
        r = await fetch(GROQ_BASE+'/chat/completions', {
          method:'POST',
          headers:{ 'Content-Type':'application/json', 'Authorization':'Bearer '+GROQ_KEY },
          body: JSON.stringify({ model, temperature:0.1, max_tokens:1500, messages:[
            { role:'system', content:'คุณเป็นผู้ช่วยวิเคราะห์เอกสาร TOR ของหน่วยงานราชการไทย ตอบเป็น JSON เท่านั้น ห้ามมีข้อความอื่น' },
            { role:'user', content: prompt } ] })
        });
      }
    }
    // บางโมเดลไม่รองรับ response_format -> ลองใหม่โดยไม่ใส่
    if (r.status === 400) {
      const t400 = await r.clone().text();
      // json_validate_failed = โมเดลตอบไม่ทัน/ตอบไม่ตรงรูปแบบ · response_format = โมเดลไม่รองรับโหมดนี้
      // ทั้งสองกรณีแก้ด้วยการถอด json mode ออกแล้วให้ตัวแกะ JSON ฝั่งเราจัดการแทน
      if (/response_format|json_object|json_validate_failed/i.test(t400)) {
        console.log('[TOR AI] json mode ใช้ไม่ได้ ลองใหม่แบบไม่บังคับรูปแบบ');
        r = await fetch(GROQ_BASE+'/chat/completions', {
          method:'POST',
          headers:{ 'Content-Type':'application/json', 'Authorization':'Bearer '+GROQ_KEY },
          body: JSON.stringify({ model, temperature:0.1, max_tokens:4000, reasoning_effort:'low', messages:[
            { role:'system', content:'คุณเป็นผู้ช่วยวิเคราะห์เอกสาร TOR ของหน่วยงานราชการไทย ตอบเป็น JSON เท่านั้น ห้ามมีข้อความอื่น ห้ามอธิบาย' },
            { role:'user', content: prompt } ] })
        });
        // ยังไม่ผ่านอีก -> ลองโมเดลที่เบากว่า (คิดน้อย ตอบตรงกว่า)
        if (!r.ok) {
          const light = 'openai/gpt-oss-20b';
          if (light !== model) {
            console.log('[TOR AI] ลองโมเดลเบากว่า:', light);
            const r2 = await fetch(GROQ_BASE+'/chat/completions', {
              method:'POST',
              headers:{ 'Content-Type':'application/json', 'Authorization':'Bearer '+GROQ_KEY },
              body: JSON.stringify({ model: light, temperature:0, max_tokens:4000, messages:[
                { role:'system', content:'ตอบเป็น JSON เท่านั้น ห้ามมีข้อความอื่น' },
                { role:'user', content: prompt } ] })
            });
            if (r2.ok) { r = r2; model = light; }
          }
        }
      }
    }
    // 429 = ชนเพดาน tokens/นาที (free tier 6,000-8,000 TPM) - รอแล้วลองใหม่อัตโนมัติ
    // Groq บอกเวลาที่ต้องรอมาในหัวข้อ retry-after หรือในข้อความ
    let retried429 = 0;
    while (r.status === 429 && retried429 < 2) {
      const body429 = await r.clone().text();
      const ra = Number(r.headers.get('retry-after')) ||
                 Number((body429.match(/try again in ([\d.]+)s/i)||[])[1]) || 20;
      const waitMs = Math.min(45000, Math.ceil(ra * 1000) + 1500);
      console.log(`[TOR AI] ชนเพดาน TPM รอ ${Math.round(waitMs/1000)} วินาทีแล้วลองใหม่`);
      await new Promise(z => setTimeout(z, waitMs));
      retried429++;
      r = await fetch(GROQ_BASE+'/chat/completions', {
        method:'POST',
        headers:{ 'Content-Type':'application/json', 'Authorization':'Bearer '+GROQ_KEY },
        body: JSON.stringify({ model, temperature:0.1, max_tokens:4000, reasoning_effort:'low', messages:[
          { role:'system', content:'คุณเป็นผู้ช่วยวิเคราะห์เอกสาร TOR ของหน่วยงานราชการไทย ตอบเป็น JSON เท่านั้น ห้ามมีข้อความอื่น' },
          { role:'user', content: prompt } ] })
      });
    }
    if (r.status === 429) {
      const t = await r.text();
      const secs = Number((t.match(/try again in ([\d.]+)s/i)||[])[1]);
      return res.status(429).json({ ok:false,
        error: `ใช้โควตาต่อนาทีของ Groq เต็ม (free tier จำกัด ~8,000 tokens/นาที)` +
               (secs ? ` - ลองใหม่อีกครั้งใน ${Math.ceil(secs)} วินาที` : ' - รอสักครู่แล้วลองใหม่'),
        retryAfter: secs || 30 });
    }
    if (!r.ok) {
      const t = await r.text();
      return res.status(502).json({ ok:false, error:`Groq ตอบกลับ ${r.status}: ${t.slice(0,220)}` });
    }
    const j = await r.json();
    const msg = ((j.choices||[])[0]||{}).message || {};
    // โมเดลตระกูล gpt-oss เป็น reasoning model - บางครั้งใส่คำตอบไว้ใน reasoning แทน content
    let txt = msg.content || msg.reasoning_content || msg.reasoning || '';
    if (Array.isArray(txt)) txt = txt.map(x => (x && x.text) || '').join('');
    const result = parseLooseJson(txt);
    if (!result) {
      console.warn('[TOR AI] แกะ JSON ไม่ได้ model=' + model + ' ข้อความที่ได้:', String(txt).slice(0,400));
      return res.status(502).json({ ok:false,
        error:'AI ตอบกลับในรูปแบบที่อ่านไม่ได้ (โมเดล ' + model + ')',
        raw: String(txt).slice(0,300) });
    }
    // โมเดลอาจห่อผลไว้ใน items/result/data - คลี่ออกให้ frontend ได้รูปแบบเดิมเสมอ
    const unwrapped = (result && !Array.isArray(result) && (result.items||result.result||result.data))
                      || result;
    torAiUsage[key] = (torAiUsage[key]||0) + 1;
    const usage = j.usage || {};
    console.log(`[TOR AI] ${req.user.username} task=${task} model=${model} items=${items.length} tokens=${usage.total_tokens||'?'}`);
    res.json({ ok:true, result: unwrapped, model, tokens: usage.total_tokens || null,
               used: torAiUsage[key], limit: TOR_AI_DAILY_LIMIT });
  } catch (e) {
    res.status(500).json({ ok:false, error: String(e.message||e) });
  }
});

app.get('/api/projects', requireAuth, (req, res) => {
  const { pmName, role } = req.user;
  // rowsTeamWide: ข้อมูลทีมเต็มเสมอ ไม่ว่าใคร login (ไม่ถูกกรองรายบุคคล) - ใช้กับการ์ดที่ตั้งค่าให้เห็นทีมเต็มเสมอ
  // เช่น "Update PMS" ที่ OAT อยากให้ทุกคนในทีมเห็นสถานะทั้งทีม ไม่ใช่แค่โครงการของตัวเอง
  // paymentWTeamWide/poTeamWide/stockTeamWide/stockSummaryTeamWide เป็นคู่เดียวกัน สำหรับ card ฝั่ง Tab 2 (Finance & Stock)
  if (role === 'admin' || !pmName) return res.json({
    ...latestData,
    rowsTeamWide: latestData.rows,
    paymentWTeamWide: latestData.paymentW,
    poTeamWide: latestData.po,
    stockTeamWide: latestData.stock,
    stockSummaryTeamWide: latestData.stockSummary,
    cardVisibility: CARD_VISIBILITY,
    cardDataScope: CARD_DATA_SCOPE,
  });

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
  // projectInfoAll ไม่ต้อง filter ตรงนี้เลย - ไม่อยู่ใน object ที่ spread ทับด้านล่าง จึงหลุดผ่าน ...latestData
  // มาแบบเต็ม ไม่ล็อกทีม ตามที่ตั้งใจไว้สำหรับ card "ค้นหา site/โครงการทั้งบริษัท" โดยเฉพาะ

  // ถ้าผูก pmName ไว้แล้วแต่ไม่เจอโครงการเลย มักเกิดจากชื่อใน Excel สะกดไม่ตรงกับที่ตั้งไว้
  if (latestData.rows.length && !rows.length) {
    const available = [...new Set(latestData.rows.map(r => r['PM Name']).filter(Boolean))];
    console.log(`⚠ ผู้ใช้ "${req.user.username}" ผูกกับ pmName="${pmName}" แต่ไม่พบโครงการใดเลย`);
    console.log(`  ชื่อ PM ที่มีอยู่จริงในไฟล์: ${JSON.stringify(available)}`);
  }

  res.json({
    ...latestData,
    rows, revenue, paymentW, po, stock, stockSummary, projectInfo,
    rowsTeamWide: latestData.rows,
    paymentWTeamWide: latestData.paymentW,
    poTeamWide: latestData.po,
    stockTeamWide: latestData.stock,
    stockSummaryTeamWide: latestData.stockSummary,
    cardVisibility: CARD_VISIBILITY,
    cardDataScope: CARD_DATA_SCOPE,
    scope: pmName,
  });
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
