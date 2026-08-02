// PJL — Project List Dashboard API
// รับข้อมูลจาก Make.com webhook (POST) แล้วเก็บไว้ให้ frontend ดึงไปแสดง (GET)
// รูปแบบเดียวกับ CNX / HDY / CEI: Excel -> Make.com -> POST /api/webhook/excel -> เก็บ in-memory -> GET /api/projects

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

  console.log(`✓ PJL webhook: received ${rows.length} rows @ ${latestData.updatedAt}`);
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
app.listen(PORT, () => console.log(`PJL API listening on :${PORT}`));
