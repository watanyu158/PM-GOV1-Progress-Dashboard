// PM-GOV1 Progress Dashboard API
// รับข้อมูลจาก Make.com webhook (POST) แล้วเก็บไว้ให้ frontend ดึงไปแสดง (GET)
// รูปแบบเดียวกับ CNX / HDY / CEI: Excel -> Make.com -> POST /api/webhook/excel -> เก็บ in-memory -> GET /api/projects
// Deploy URL: https://pm-gov1-progress-api.onrender.com
// Make.com scenario webhook (trigger): https://hook.eu1.make.com/m58fg8qhrhydcl5nru3md2dumc5c9xqu

const express = require('express');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

// เก็บ snapshot ล่าสุดไว้ใน memory (เหมือน pattern เดิม)
let latestData = {
  rows: [],
  updatedAt: null,
};

// --- รับข้อมูลจาก Make.com ---
// Make.com จะอ่านชีต "DATA" ทั้งชีตแล้ว POST array ของ object (key = header คอลัมน์) มาที่นี่
app.post('/api/webhook/excel', (req, res) => {
  const payload = req.body;
  const rows = Array.isArray(payload) ? payload : payload.rows || payload.data || [];

  if (!Array.isArray(rows)) {
    return res.status(400).json({ ok: false, error: 'Expected an array of row objects' });
  }

  latestData = {
    rows,
    updatedAt: new Date().toISOString(),
  };

  console.log(`✓ PM-GOV1 webhook: received ${rows.length} rows @ ${latestData.updatedAt}`);
  res.json({ ok: true, count: rows.length, updatedAt: latestData.updatedAt });
});

// --- ให้ frontend ดึงข้อมูลไปแสดง ---
app.get('/api/projects', (req, res) => {
  res.json(latestData);
});

app.get('/api/health', (req, res) => {
  res.json({ ok: true, rows: latestData.rows.length, updatedAt: latestData.updatedAt });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PM-GOV1 API listening on :${PORT}`));
