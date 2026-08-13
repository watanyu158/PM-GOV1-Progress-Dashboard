// PM-GOV1 Progress Dashboard API
// รับ "ไฟล์ Excel ทั้งไฟล์" จาก Make.com (multipart/form-data, field name = "file")
// แล้วแกะ (parse) เองฝั่ง server จาก sheet "DATA" — เหมือน pattern เดียวกับ CNX/HDY/CEI
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

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// เก็บ snapshot ล่าสุดไว้ใน memory
let latestData = {
  rows: [],
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

// --- แกะไฟล์ Excel ---
// อ่าน sheet ชื่อ "DATA" (หรือ sheet แรกถ้าหาไม่เจอ) โดยแถวที่ 2 = header, แถวที่ 3 เป็นต้นไป = ข้อมูล
// ตรงกับโครงสร้างไฟล์ All_Project_List / PM_GOV1_Progress_Update ทุกประการ
function parseExcelBuffer(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === 'DATA') || wb.SheetNames[0];
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

// --- รับข้อมูลจาก Make.com ---
// รองรับ 2 ทาง: (1) ไฟล์ Excel จริงแบบ multipart (ทางหลัก) (2) JSON array ตรง ๆ (ทางสำรอง เผื่อ mapping ฝั่ง Make.com เอง)
app.post('/api/webhook/excel', upload.single('file'), (req, res) => {
  try {
    if (req.file) {
      const { rows, sheetName } = parseExcelBuffer(req.file.buffer);
      latestData = { rows, updatedAt: new Date().toISOString() };
      lastRaw = {
        receivedAt: latestData.updatedAt,
        contentType: req.headers['content-type'] || null,
        mode: 'file',
        sheetName,
        fileSize: req.file.buffer.length,
        preview: rows[0] ? JSON.stringify(rows[0]).slice(0, 500) : null,
      };
      console.log(`✓ PM-GOV1 webhook: parsed ${rows.length} rows from sheet "${sheetName}" (${req.file.buffer.length} bytes)`);
      return res.json({ ok: true, mode: 'file', sheet: sheetName, count: rows.length, updatedAt: latestData.updatedAt });
    }

    // ทางสำรอง: body เป็น JSON array ของแถวข้อมูลตรง ๆ
    const payload = req.body;
    const rows = Array.isArray(payload) ? payload : payload?.rows || payload?.data || null;
    if (rows) {
      latestData = { rows, updatedAt: new Date().toISOString() };
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

// --- ให้ frontend ดึงข้อมูลไปแสดง ---
app.get('/api/projects', (req, res) => {
  res.json(latestData);
});

app.get('/api/health', (req, res) => {
  res.json({ ok: true, rows: latestData.rows.length, updatedAt: latestData.updatedAt });
});

// --- Debug: ดูว่า Make.com ส่งอะไรมาล่าสุดจริง ๆ ---
app.get('/api/debug/last-payload', (req, res) => {
  res.json(lastRaw);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`PM-GOV1 API listening on :${PORT}`);

  // --- Keep-alive: กันไม่ให้ Render free tier sleep (spin down หลังไม่มี request ~15 นาที) ---
  // ยิง request เข้าตัวเองทุก 10 นาที ผ่าน public URL ของ Render (RENDER_EXTERNAL_URL มีให้อัตโนมัติ
  // บน Render เท่านั้น — รันในเครื่อง/ที่อื่นจะไม่ทำงาน ไม่กระทบอะไร)
  const selfUrl = process.env.RENDER_EXTERNAL_URL;
  if (selfUrl) {
    setInterval(() => {
      fetch(`${selfUrl}/api/health`)
        .then(() => console.log(`♥ keep-alive ping ok @ ${new Date().toISOString()}`))
        .catch(err => console.log('keep-alive ping failed:', err.message));
    }, 10 * 60 * 1000); // ทุก 10 นาที (สั้นกว่า timeout ของ Render ที่ ~15 นาที)
    console.log(`♥ keep-alive enabled → pinging ${selfUrl}/api/health every 10 min`);
  } else {
    console.log('ℹ keep-alive skipped (RENDER_EXTERNAL_URL not set — not running on Render, or running locally)');
  }
});
