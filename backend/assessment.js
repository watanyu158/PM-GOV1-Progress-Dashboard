// โมดูลแบบประเมิน 5 Key Indexes (One-on-One) - PM-GOV1 Dashboard
// หลักการ:
//  - เฉลยและคะแนนอยู่ฝั่ง server เท่านั้น (ส่งไปหน้าเว็บเฉพาะข้อความ + รหัสตัวเลือกแบบสุ่ม)
//  - น้อง ๆ ทำได้ "ครั้งเดียวต่อรอบ (ปี)" ส่งแล้วล็อก แก้ไม่ได้ และไม่เห็นคะแนนของตัวเอง
//  - ที่เก็บถาวรคือ Excel บน OneDrive ของ OAT ผ่าน Make.com (ASSESSMENT_MAKE_URL) เพราะ Render free ไม่มีดิสก์ถาวร
//  - บัญชี OAT (admin) และบัญชี pmName ของ OAT ทำได้แบบ "ทดลอง": ไม่บันทึกลง Excel ไม่นับในกราฟ ทำซ้ำได้
//
// ENV: ASSESSMENT_MAKE_URL, ASSESSMENT_MAKE_SECRET, ASSESSMENT_CYCLE (ค่าเริ่มต้น = ปี ค.ศ. ปัจจุบัน เวลาไทย),
//      ASSESSMENT_OPEN ("false" = ปิดรับ), ASSESSMENT_OWNER_PM (ค่าเริ่มต้น "Watanyu Anantakunakorn" = ไม่ถูกประเมิน)
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

module.exports = function registerAssessment(app, { requireAuth, normName, teamPmNames, dataDir }) {
  const BANK = JSON.parse(fs.readFileSync(path.join(__dirname, 'assessment-bank.json'), 'utf8'));
  const INDEXES = BANK.indexes;
  const CODE = { 'Attitude': 'att', 'Awareness': 'awr', 'Engagement': 'eng', 'Work-Life Balance': 'wlb', 'Process': 'prc' };
  const MAKE_URL = (process.env.ASSESSMENT_MAKE_URL || '').trim();
  const MAKE_SECRET = (process.env.ASSESSMENT_MAKE_SECRET || '').trim();
  const OWNER_PM = process.env.ASSESSMENT_OWNER_PM || 'Watanyu Anantakunakorn';
  const OPEN = String(process.env.ASSESSMENT_OPEN || 'true').toLowerCase() !== 'false';
  const cycle = () => (process.env.ASSESSMENT_CYCLE || String(new Date(Date.now() + 7 * 3600e3).getUTCFullYear())).trim();
  const STORE_FILE = path.join(dataDir, 'assessment-store.json');
  const CACHE_TTL = 2 * 60 * 1000;

  // ---- ดัชนีคำถาม/ตัวเลือก (ใช้ตรวจคำตอบ + คิดคะแนน) ----
  const OPT = new Map();   // optionId -> { qid, score }
  const QMAP = new Map();  // qid -> { index, optionIds:Set }
  for (const q of BANK.scenario) {
    QMAP.set(q.id, { index: q.index, optionIds: new Set(q.options.map(o => o.id)) });
    for (const o of q.options) {
      if (OPT.has(o.id)) throw new Error('assessment-bank.json: รหัสตัวเลือกซ้ำกัน ' + o.id);
      OPT.set(o.id, { qid: q.id, score: o.score });
    }
  }
  const SELF_IDS = BANK.self.map(s => s.id);
  const teamSet = new Set(teamPmNames.map(normName));
  const ownerNorm = normName(OWNER_PM);
  const assessedNames = teamPmNames.filter(n => normName(n) !== ownerNorm);
  const canonName = (n) => teamPmNames.find(t => normName(t) === normName(n)) || null;

  // ---- ที่เก็บ ----
  let rows = [];            // แถวทุกประเภทของรอบปัจจุบัน
  let loadedAt = 0;
  let lastError = null;
  let lastDiag = null;      // ผลตรวจการอ่านจาก Make ครั้งล่าสุด (แสดงให้ admin ดู)
  const trialResults = new Map(); // username -> ผลทดลองล่าสุด (อยู่ใน memory เท่านั้น)
  const busy = new Set();         // กันกดส่งซ้ำพร้อมกัน

  function loadLocal() {
    try { const j = JSON.parse(fs.readFileSync(STORE_FILE, 'utf8')); if (Array.isArray(j.rows)) rows = j.rows; } catch { /* ยังไม่มีไฟล์ */ }
  }
  function saveLocal() {
    try { fs.writeFileSync(STORE_FILE, JSON.stringify({ rows }, null, 1), 'utf8'); } catch (e) { console.log('⚠ assessment: เขียนไฟล์สำรองไม่สำเร็จ', e.message); }
  }
  loadLocal();
  console.log(MAKE_URL ? '✓ assessment: เก็บผลผ่าน Make.com (Excel)' : '⚠ assessment: ไม่ได้ตั้ง ASSESSMENT_MAKE_URL - เก็บเฉพาะไฟล์ในเครื่อง server (หายเมื่อ redeploy บน Render free)');

  async function makeCall(payload) {
    const r = await fetch(MAKE_URL, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ secret: MAKE_SECRET, ...payload }), signal: AbortSignal.timeout(25000),
    });
    const text = await r.text();
    let j = null; try { j = JSON.parse(text); } catch { /* ไม่ใช่ JSON */ }
    if (!r.ok) throw new Error(`Make ตอบ HTTP ${r.status}`);
    if (!j || j.ok === false) { console.log('assessment: Make ตอบไม่ถูกรูปแบบ ->', String(text).slice(0, 300)); throw new Error('Make ตอบกลับไม่ถูกรูปแบบ (ต้องเป็น JSON {"ok":true,...}) เริ่มต้นข้อความที่ได้: ' + String(text).slice(0, 80).replace(/\s+/g, ' ')); }
    return j;
  }
  const num = (v) => { const n = Number(v); return Number.isFinite(n) && v !== '' && v !== null ? n : null; };
  function normRow(r) {
    const o = (r && r.values && typeof r.values === 'object') ? r.values : (r || {});
    const out = { cycle: String(o.cycle ?? '').trim(), type: String(o.type ?? '').trim(), pmName: String(o.pmName ?? '').trim(),
      username: String(o.username ?? '').trim().toLowerCase(), submittedAt: String(o.submittedAt ?? '').trim(), answers: o.answers ?? '' };
    for (const c of Object.values(CODE)) for (const k of ['scn', 'self', 'oat']) out[`${c}_${k}`] = num(o[`${c}_${k}`]);
    return out;
  }
  // refresh: ดึงแถวทั้งหมดของรอบนี้จาก Make (force=true ข้ามแคช) - ถ้าล้มเหลวจะ throw เพื่อให้ผู้เรียกตัดสินใจ "ปิดไว้ก่อน"
  async function refresh(force) {
    if (!MAKE_URL) return;
    if (!force && Date.now() - loadedAt < CACHE_TTL) return;
    try {
      const j = await makeCall({ action: 'list', cycle: cycle() });
      const arr = Array.isArray(j.rows) ? j.rows : (Array.isArray(j.data) ? j.data : []);
      const all = arr.map(normRow);
      rows = all.filter(r => r.cycle === cycle());
      const byType = {}; rows.forEach(r => { byType[r.type || '?'] = (byType[r.type || '?'] || 0) + 1; });
      lastDiag = { at: Date.now(), rawCount: arr.length, kept: rows.length, byType,
        cycles: [...new Set(all.map(r => r.cycle))].slice(0, 6), keys: arr[0] ? Object.keys((arr[0] && arr[0].values) || arr[0]).slice(0, 8) : [] };
      console.log(`assessment: อ่านจาก Make ${arr.length} แถว ใช้ได้ ${rows.length} แถว (รอบ ${cycle()}) ประเภท=${JSON.stringify(byType)} cycles=${JSON.stringify(lastDiag.cycles)}`);
      loadedAt = Date.now(); lastError = null; saveLocal();
    } catch (e) { lastError = e.message; throw e; }
  }
  async function appendRow(row) {
    if (MAKE_URL) await makeCall({ action: 'append', row });
    rows.push(normRow(row)); saveLocal();
  }

  const cur = () => rows.filter(r => r.cycle === cycle());
  const latest = (type, pm) => cur().filter(r => r.type === type && normName(r.pmName) === normName(pm))
    .sort((a, b) => String(a.submittedAt).localeCompare(String(b.submittedAt))).pop() || null;
  const selfRowOf = (pm) => latest('self_test', pm);

  function who(user) {
    const pm = normName(user.pmName);
    const isAdmin = user.role === 'admin';
    const inTeam = teamSet.has(pm);
    const assessed = inTeam && pm !== ownerNorm && !isAdmin;
    const trial = isAdmin || (inTeam && pm === ownerNorm);
    return { isAdmin, assessed, trial, mode: assessed ? 'real' : (trial ? 'trial' : 'none'), pmName: canonName(user.pmName) };
  }
  const rnd = (n) => crypto.randomInt(n);
  function shuffle(a) { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = rnd(i + 1); [a[i], a[j]] = [a[j], a[i]]; } return a; }
  const r2 = (x) => Math.round(x * 100) / 100;

  function score(body) {
    const sc = (body && body.scenario) || {}, sf = (body && body.self) || {};
    const per = {}; INDEXES.forEach(i => per[i] = []);
    const picked = {};
    for (const [qid, meta] of QMAP) {
      const oid = sc[qid];
      if (typeof oid !== 'string' || !meta.optionIds.has(oid)) return { error: `คำตอบข้อทดสอบไม่ครบหรือไม่ถูกต้อง (${qid})` };
      const s = OPT.get(oid).score; per[meta.index].push(s); picked[qid] = s;
    }
    const scenario = {}, self = {};
    for (const i of INDEXES) scenario[i] = r2(per[i].reduce((a, b) => a + b, 0) / per[i].length);
    for (const s of BANK.self) {
      const v = Number(sf[s.id]);
      if (!Number.isInteger(v) || v < 1 || v > 5) return { error: `ยังไม่ได้ประเมินตนเองให้ครบ (${s.index})` };
      self[s.index] = v;
    }
    return { scenario, self, picked };
  }
  function toRow(type, user, pmName, extra) {
    const row = { cycle: cycle(), type, pmName, username: user.username, submittedAt: new Date().toISOString(), answers: '' };
    for (const c of Object.values(CODE)) for (const k of ['scn', 'self', 'oat']) row[`${c}_${k}`] = '';
    return Object.assign(row, extra);
  }

  // ---- สถานะของฉัน ----
  app.get('/api/assessment/status', requireAuth, async (req, res) => {
    const w = who(req.user);
    const base = { ok: true, cycle: cycle(), open: OPEN, mode: w.mode, isAdmin: w.isAdmin, storage: MAKE_URL ? 'make' : 'local' };
    if (w.mode === 'none') return res.json({ ...base, submitted: false });
    if (w.mode === 'trial') return res.json({ ...base, submitted: false, hasTrialResult: trialResults.has(req.user.username) });
    try { await refresh(false); } catch (e) { return res.status(503).json({ ok: false, error: 'ระบบจัดเก็บผลยังไม่พร้อม กรุณาลองใหม่อีกครั้งในอีกสักครู่' }); }
    const mine = selfRowOf(w.pmName);
    res.json({ ...base, submitted: !!mine, submittedAt: mine ? mine.submittedAt : null });
  });

  // ---- ขอแบบทดสอบ (ไม่มีเฉลย, สลับลำดับ, ไม่บอกว่าแต่ละข้อเป็น index ไหน) ----
  app.get('/api/assessment/form', requireAuth, async (req, res) => {
    const w = who(req.user);
    if (w.mode === 'none') return res.status(403).json({ ok: false, error: 'บัญชีนี้ไม่อยู่ในรายชื่อผู้ทำแบบประเมิน' });
    if (!OPEN) return res.status(403).json({ ok: false, error: 'ปิดรับแบบประเมินแล้ว' });
    if (w.mode === 'real') {
      try { await refresh(true); } catch { return res.status(503).json({ ok: false, error: 'ระบบจัดเก็บผลยังไม่พร้อม กรุณาลองใหม่อีกครั้ง' }); }
      if (selfRowOf(w.pmName)) return res.status(409).json({ ok: false, error: 'คุณทำแบบประเมินรอบนี้ไปแล้ว' });
    }
    res.set('Cache-Control', 'no-store');
    res.json({
      ok: true, cycle: cycle(),
      scenario: shuffle(BANK.scenario).map(q => ({ id: q.id, text: q.text, options: shuffle(q.options).map(o => ({ id: o.id, text: o.text })) })),
      self: BANK.self.map(s => ({ id: s.id, index: s.index, levels: s.levels })),
    });
  });

  // ---- ส่งคำตอบ: ครั้งเดียว ล็อก ----
  app.post('/api/assessment/submit', requireAuth, async (req, res) => {
    const w = who(req.user);
    if (w.mode === 'none') return res.status(403).json({ ok: false, error: 'บัญชีนี้ไม่อยู่ในรายชื่อผู้ทำแบบประเมิน' });
    if (!OPEN) return res.status(403).json({ ok: false, error: 'ปิดรับแบบประเมินแล้ว' });
    const s = score(req.body);
    if (s.error) return res.status(400).json({ ok: false, error: s.error });

    if (w.mode === 'trial') {
      const result = { at: new Date().toISOString(), scenario: s.scenario, self: s.self };
      trialResults.set(req.user.username, result);
      return res.json({ ok: true, trial: true, result });
    }
    if (busy.has(req.user.username)) return res.status(429).json({ ok: false, error: 'กำลังบันทึกอยู่ กรุณารอสักครู่' });
    busy.add(req.user.username);
    try {
      try { await refresh(true); } catch { return res.status(503).json({ ok: false, error: 'ระบบจัดเก็บผลยังไม่พร้อม ยังไม่ได้บันทึก กรุณากดส่งอีกครั้ง' }); }
      if (selfRowOf(w.pmName)) return res.status(409).json({ ok: false, error: 'คุณทำแบบประเมินรอบนี้ไปแล้ว แก้ไขไม่ได้' });
      const extra = { answers: JSON.stringify(s.picked) };
      for (const i of INDEXES) { extra[`${CODE[i]}_scn`] = s.scenario[i]; extra[`${CODE[i]}_self`] = s.self[i]; }
      const row = toRow('self_test', req.user, w.pmName, extra);
      try { await appendRow(row); } catch (e) { console.log('assessment append ล้มเหลว:', e.message); return res.status(503).json({ ok: false, error: 'บันทึกไม่สำเร็จ ยังไม่ได้ส่ง กรุณากดส่งอีกครั้ง' }); }
      console.log(`✓ assessment: ${req.user.username} ส่งแบบประเมินรอบ ${cycle()}`);
      res.json({ ok: true, submittedAt: row.submittedAt });
    } finally { busy.delete(req.user.username); }
  });

  // ---- โหมดทดลอง (OAT) ----
  app.get('/api/assessment/trial/result', requireAuth, (req, res) => {
    if (!who(req.user).trial) return res.status(403).json({ ok: false, error: 'เฉพาะบัญชีทดลอง' });
    res.json({ ok: true, result: trialResults.get(req.user.username) || null });
  });
  app.post('/api/assessment/trial/reset', requireAuth, (req, res) => {
    if (!who(req.user).trial) return res.status(403).json({ ok: false, error: 'เฉพาะบัญชีทดลอง' });
    trialResults.delete(req.user.username); res.json({ ok: true });
  });

  // ---- Admin (OAT): ดูผล + ให้คะแนนน้อง ----
  const adminOnly = (req, res, next) => req.user.role === 'admin' ? next() : res.status(403).json({ ok: false, error: 'ต้องเป็น admin เท่านั้น' });
  app.get('/api/assessment/admin/overview', requireAuth, adminOnly, async (req, res) => {
    let stale = null;
    try { await refresh(req.query.refresh === '1'); } catch (e) { stale = e.message; }
    const people = assessedNames.map(pm => {
      const sr = selfRowOf(pm), or = latest('oat_rating', pm);
      const pick = (r, k) => r ? Object.fromEntries(INDEXES.map(i => [i, r[`${CODE[i]}_${k}`]])) : null;
      return { pmName: pm, submitted: !!sr, submittedAt: sr ? sr.submittedAt : null, scenario: pick(sr, 'scn'), self: pick(sr, 'self'),
        oat: pick(or, 'oat'), oatAt: or ? or.submittedAt : null };
    });
    res.json({ ok: true, cycle: cycle(), storage: MAKE_URL ? 'make' : 'local', loadedAt, stale, diag: lastDiag, indexes: INDEXES,
      rubric: Object.fromEntries(BANK.self.map(s => [s.index, s.levels])), people });
  });
  app.post('/api/assessment/admin/oat-rating', requireAuth, adminOnly, async (req, res) => {
    const pm = canonName((req.body || {}).pmName);
    if (!pm || normName(pm) === ownerNorm) return res.status(400).json({ ok: false, error: 'ชื่อ PM ไม่อยู่ในรายชื่อที่ต้องประเมิน' });
    const rt = (req.body || {}).ratings || {};
    const extra = {};
    for (const i of INDEXES) {
      const v = Number(rt[i]);
      if (!Number.isInteger(v) || v < 1 || v > 5) return res.status(400).json({ ok: false, error: `ต้องให้คะแนน 1-5 ครบทุก index (${i})` });
      extra[`${CODE[i]}_oat`] = v;
    }
    try { await appendRow(toRow('oat_rating', req.user, pm, extra)); }
    catch (e) { return res.status(503).json({ ok: false, error: 'บันทึกไม่สำเร็จ กรุณาลองใหม่' }); }
    res.json({ ok: true });
  });
};
