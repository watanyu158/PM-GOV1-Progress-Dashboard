/*! asm-report.js - สร้างเอกสาร Word สรุปผลประเมิน 5 Key Indexes (ต่อคน) ฝั่งเบราว์เซอร์
 * ต้องโหลด lib/docx.min.js ก่อน (global: docx)
 * ใช้: const blob = await AsmReport.build({person, indexes, rubric, cycle, sources, generatedAt})
 *   person  = { pmName, submitted, submittedAt, oat, oatAt, scenario, self, items }   (oat/scenario/self = {index: 1-5})
 *   sources = { oat, scn, self }  แหล่งที่ใช้คิด Avg. (ตามช่องติ๊กในหน้าจอ)
 * ไม่มีเรื่อง IDP โดยตั้งใจ: เอกสารนี้สรุปผล + ประเด็นชวนคุย + พื้นที่บันทึก 1:1 ให้ VP แก้ไขต่อใน Word
 */
(function (root) {
  'use strict';

  // ---------- ดีไซน์ ----------
  const W = 9638;                                   // ความกว้างเนื้อหา A4 ขอบซ้ายขวา 2 ซม. (DXA)
  const FONT = 'TH SarabunPSK';
  const FONTS = { ascii: FONT, hAnsi: FONT, cs: FONT, eastAsia: FONT };
  const C = { navy: '1F3A5F', ink: '1F2937', muted: '6B7280', line: 'D9DEE5', soft: 'F3F5F8', soft2: 'EAF0F6',
    warn: 'FFF4DC', warnLine: 'E0B34F', note: 'FFFBEB', oat: '1F3A5F', scn: '2A9D8F', self: 'E07A1F', up: 'B54708', down: '175CD3' };
  const HEX = k => '#' + C[k];
  const SRC = { oat: 'VP', scn: 'ข้อทดสอบ', self: 'ตนเอง' };
  const LEGEND = { oat: 'VP ประเมิน', scn: 'ผลข้อทดสอบ', self: 'ประเมินตนเอง' };
  const SCALE = ['F1F4F8', 'DCE6F1', 'C3D5E8', 'A3BFDC', '86A9D1'];   // ระดับ 1-5 (น้ำเงินไล่เฉดอ่อน->เข้ม ไม่ใช้สีแดง/เขียว เพื่อไม่ตัดสินเกินจริง)
  const shade = v => v == null ? 'FFFFFF' : SCALE[Math.min(5, Math.max(1, Math.round(v))) - 1];
  const CANVAS_FONT = '"TH Sarabun New","Sarabun","Leelawadee UI","Noto Sans Thai","Tahoma",sans-serif';

  // ---------- ตัวช่วยตัวเลข ----------
  const num = v => (v == null || v === '' || isNaN(Number(v))) ? null : Number(v);
  const r1 = v => Math.round(v * 10) / 10;
  const fmt = v => v == null ? '–' : String(r1(v));
  const mean = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : null;
  const sg = d => d == null ? '–' : (d > 0 ? '+' : '') + String(r1(d));
  const thDate = s => {
    const d = s instanceof Date ? s : new Date(s);
    if (isNaN(d)) return '–';
    return d.toLocaleDateString('th-TH', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'Asia/Bangkok' });
  };
  const names = rs => rs.map(r => r.index).join(', ');

  // ---------- ประเด็นชวนคุย (ไม่ใช่แผนพัฒนา) ----------
  const BASE_Q = {
    'Attitude': ['ครั้งล่าสุดที่ได้รับ feedback ที่ไม่ตรงกับที่คิด ตอนนั้นตอบสนองอย่างไร และได้ปรับอะไรบ้าง',
      'ช่วงที่โครงการมีปัญหาหนัก ทีมเห็นพฤติกรรมอะไรจากคุณบ้าง และคุณอยากให้ทีมเห็นแบบไหน'],
    'Awareness': ['ยกตัวอย่างความเสี่ยงหรือ issue ที่เห็นก่อนคนอื่น เห็นได้อย่างไร และทำอะไรต่อ',
      'มีเรื่องไหนที่มารู้ช้าเกินไป ถ้าย้อนกลับไปได้ สัญญาณอะไรที่ควรจับได้ก่อนหน้านั้น'],
    'Engagement': ['stakeholder คนไหนที่สื่อสารยากหรือเริ่มห่างออกไป ตอนนี้ดูแลอย่างไร',
      'ครั้งล่าสุดที่เกิดความเข้าใจผิดระหว่างฝ่าย จัดการอย่างไร และได้ผลเป็นอย่างไร'],
    'Work-Life Balance': ['ช่วงที่ผ่านมางานไหนที่ถือไว้เองมากที่สุด และมอบหมายต่อให้ทีมได้หรือไม่ ติดอยู่ตรงไหน',
      'ช่วง peak ของโครงการ มีอะไรที่ช่วยให้ทั้งตัวเองและทีมไม่หมดแรง และมีอะไรที่ยังขาด'],
    'Process': ['อะไรทำให้การอัปเดต PMS หรือเอกสารตามรอบทำได้ยากที่สุด',
      'มี process ข้อไหนที่เห็นว่าไม่เหมาะกับโครงการ และอยากเสนอปรับอย่างไร']
  };

  // ---------- วิเคราะห์ข้อมูล ----------
  function analyze(inp) {
    const p = inp.person, IDX = inp.indexes;
    const tick = Object.assign({ oat: true, scn: true, self: true }, inp.sources || {});
    const tickN = ['oat', 'scn', 'self'].filter(k => tick[k]).length;
    const g = (o, i) => o ? num(o[i]) : null;
    const rows = IDX.map(i => {
      const V = g(p.oat, i), T = g(p.scenario, i), S = g(p.self, i);
      const used = [];
      if (tick.oat && V != null) used.push(V);
      if (tick.scn && T != null) used.push(T);
      if (tick.self && S != null) used.push(S);
      const Araw = mean(used);
      return { index: i, V, T, S, Araw, A: Araw == null ? null : r1(Araw), partial: used.length > 0 && used.length < tickN,
        dT: V != null && T != null ? r1(T - V) : null, dS: V != null && S != null ? r1(S - V) : null, dST: T != null && S != null ? r1(S - T) : null };
    });
    const col = k => rows.map(r => r[k]).filter(v => v != null);
    const m = k => { const a = col(k); return a.length ? mean(a) : null; };
    const overall = { V: m('V'), T: m('T'), S: m('S'), A: m('Araw') };
    overall.dT = overall.V != null && overall.T != null ? r1(overall.T - overall.V) : null;
    overall.dS = overall.V != null && overall.S != null ? r1(overall.S - overall.V) : null;
    const has = { oat: col('V').length > 0, scn: col('T').length > 0, self: col('S').length > 0 };
    const missing = ['oat', 'scn', 'self'].filter(k => tick[k] && !has[k]);
    const ranked = rows.filter(r => r.Araw != null).sort((a, b) => b.Araw - a.Araw);
    const spread = ranked.length > 1 ? ranked[0].Araw - ranked[ranked.length - 1].Araw : 0;
    const flat = ranked.length > 1 && spread < 0.3;
    const top = ranked.length ? ranked.filter(r => r.Araw >= ranked[0].Araw - 0.05).slice(0, 2) : [];
    const low = ranked.length > 1 ? ranked.slice().reverse().filter(r => r.Araw <= ranked[ranked.length - 1].Araw + 0.05).slice(0, 2) : [];
    const flag = {
      sUp: rows.filter(r => r.dS != null && r.dS >= 1), sDown: rows.filter(r => r.dS != null && r.dS <= -1),
      tUp: rows.filter(r => r.dT != null && r.dT >= 1), tDown: rows.filter(r => r.dT != null && r.dT <= -1),
      stUp: rows.filter(r => r.dST != null && r.dST >= 1), stDown: rows.filter(r => r.dST != null && r.dST <= -1)
    };
    const vpLow = rows.filter(r => r.V != null && r.V <= 2);
    return { rows, overall, has, missing, tick, tickN, ranked, top, low, flat, flag, vpLow,
      any: has.oat || has.scn || has.self };
  }

  // ข้อสังเกตรายด้านจากตัวเลข (ใช้คำว่า "อาจ" เพราะเป็นข้อสันนิษฐานเพื่อชวนคุย ไม่ใช่ข้อสรุป)
  function insightsOf(r) {
    const out = [];
    if (r.V != null) {
      if (r.dS == null) out.push('ยังไม่มีผลประเมินตนเองให้เทียบกับ VP');
      else if (r.dS >= 1) out.push(`ประเมินตนเองสูงกว่า VP ${fmt(r.dS)} ระดับ ควรให้ยกตัวอย่างงานจริงที่ทำให้เห็นต่างกัน (อาจเป็นเรื่องมุมมองหรือการตีความเกณฑ์)`);
      else if (r.dS <= -1) out.push(`ประเมินตนเองต่ำกว่า VP ${fmt(-r.dS)} ระดับ ผู้รับการประเมินอาจถ่อมตัวหรือยังไม่เห็นคุณค่าของงานที่ทำ ควรชวนเล่าผลงานด้านนี้`);
      else out.push('ประเมินตนเองใกล้เคียงกับ VP (ต่างกันไม่ถึง 1 ระดับ) มุมมองตรงกัน');
      if (r.dT == null) out.push('ยังไม่มีผลข้อทดสอบให้เทียบกับ VP');
      else if (r.dT >= 1) out.push(`ผลข้อทดสอบสูงกว่าที่ VP เห็นในงานจริง ${fmt(r.dT)} ระดับ น่าจะรู้แนวทางที่ควรทำ แต่ในงานจริงอาจยังทำได้ไม่สม่ำเสมอ`);
      else if (r.dT <= -1) out.push(`ผลข้อทดสอบต่ำกว่าที่ VP เห็นในงานจริง ${fmt(-r.dT)} ระดับ ในงานจริงอาจทำได้ดีกว่าการตอบโจทย์สมมติ หรือตีความโจทย์ต่างออกไป ลองดูรายข้อในภาคผนวก`);
      else out.push('ผลข้อทดสอบใกล้เคียงกับที่ VP เห็นในงานจริง');
    } else {
      out.push('VP ยังไม่ได้ให้คะแนนด้านนี้ จึงยังเทียบกับมุมมองของ VP ไม่ได้');
      if (r.dST != null) {
        if (r.dST >= 1) out.push(`ประเมินตนเองสูงกว่าผลข้อทดสอบ ${fmt(r.dST)} ระดับ`);
        else if (r.dST <= -1) out.push(`ประเมินตนเองต่ำกว่าผลข้อทดสอบ ${fmt(-r.dST)} ระดับ`);
        else out.push('ประเมินตนเองใกล้เคียงกับผลข้อทดสอบ');
      }
    }
    return out;
  }
  function promptsOf(r) {
    const base = BASE_Q[r.index] || [];
    const gap = [];
    if (r.dS != null && r.dS >= 1) gap.push(`ขอตัวอย่างงานที่ทำให้ประเมินตนเองระดับ ${fmt(r.S)} ขณะที่ VP เห็นที่ระดับ ${fmt(r.V)}`);
    else if (r.dS != null && r.dS <= -1) gap.push(`ชวนเล่างานที่ภูมิใจในด้านนี้ VP เห็นที่ระดับ ${fmt(r.V)} ซึ่งสูงกว่าที่ประเมินตนเอง (${fmt(r.S)})`);
    else if (r.dT != null && r.dT <= -1) gap.push('ชวนดูข้อทดสอบที่ตอบได้ระดับต่ำ ว่าตอนตอบคิดอย่างไร และในงานจริงทำอย่างไร');
    else if (r.dT != null && r.dT >= 1) gap.push('ชวนคุยว่าตัวเลือกที่เลือกในข้อทดสอบ ทำได้จริงในโครงการตอนนี้หรือไม่ ติดอะไร');
    return gap.length ? [base[0], gap[0]].filter(Boolean) : base.slice(0, 2);
  }

  // ---------- กราฟ (canvas -> PNG) ----------
  function mkCanvas(w, h) { const c = document.createElement('canvas'); c.width = w; c.height = h; return c; }
  async function toPng(c) {
    const b = await new Promise(res => c.toBlob(res, 'image/png'));
    return new Uint8Array(await b.arrayBuffer());
  }
  function drawRadar(series, labels) {
    const w = 1500, h = 1400, c = mkCanvas(w, h), g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, w, h);
    const cx = w / 2, cy = 640, R = 455, n = labels.length;
    const ang = k => -Math.PI / 2 + k * 2 * Math.PI / n;
    const pt = (k, v) => [cx + Math.cos(ang(k)) * R * v / 5, cy + Math.sin(ang(k)) * R * v / 5];
    for (let v = 1; v <= 5; v++) {
      g.beginPath();
      for (let k = 0; k < n; k++) { const [x, y] = pt(k, v); k ? g.lineTo(x, y) : g.moveTo(x, y); }
      g.closePath(); g.lineWidth = v === 3 ? 3.5 : 2; g.strokeStyle = v === 3 ? '#8FA0B5' : '#D8DEE6';
      g.setLineDash(v === 3 ? [14, 10] : []); g.stroke();
    }
    g.setLineDash([]); g.strokeStyle = '#E3E8EE'; g.lineWidth = 2;
    for (let k = 0; k < n; k++) { const [x, y] = pt(k, 5); g.beginPath(); g.moveTo(cx, cy); g.lineTo(x, y); g.stroke(); }
    g.font = `500 26px ${CANVAS_FONT}`; g.fillStyle = '#7B8794'; g.textAlign = 'left'; g.textBaseline = 'middle';
    for (let v = 1; v <= 5; v++) { const [x, y] = pt(0, v); g.fillText(String(v), x + 12, y + 2); }
    g.font = `700 40px ${CANVAS_FONT}`; g.fillStyle = '#1F2937';
    for (let k = 0; k < n; k++) {
      const a = ang(k), cs = Math.cos(a), sn = Math.sin(a), lx = cx + cs * (R + 62), ly = cy + sn * (R + 62);
      g.textAlign = cs > 0.3 ? 'left' : cs < -0.3 ? 'right' : 'center';
      g.textBaseline = sn < -0.5 ? 'bottom' : sn > 0.5 ? 'top' : 'middle';
      g.fillText(labels[k], lx, ly);
    }
    const style = {
      self: { dash: [4, 12], lw: 6, fill: 0, m: 'tri' }, scn: { dash: [26, 14], lw: 6, fill: 0.06, m: 'sq' }, oat: { dash: [], lw: 9, fill: 0.15, m: 'circ' }
    };
    const marker = (x, y, kind, color) => {
      g.fillStyle = '#fff'; g.strokeStyle = color; g.lineWidth = 5; g.setLineDash([]); g.beginPath();
      if (kind === 'circ') g.arc(x, y, 13, 0, Math.PI * 2);
      else if (kind === 'sq') g.rect(x - 12, y - 12, 24, 24);
      else { g.moveTo(x, y - 15); g.lineTo(x + 14, y + 11); g.lineTo(x - 14, y + 11); g.closePath(); }
      g.fill(); g.stroke();
    };
    ['self', 'scn', 'oat'].forEach(key => {
      const s = series.find(x => x.key === key); if (!s) return;
      const st = style[key], col = HEX(key), pts = [];
      s.values.forEach((v, k) => { if (v != null) pts.push(pt(k, Math.max(0, Math.min(5, v)))); });
      if (pts.length < 3) return;
      g.beginPath(); pts.forEach(([x, y], k) => k ? g.lineTo(x, y) : g.moveTo(x, y)); g.closePath();
      if (st.fill) { g.globalAlpha = st.fill; g.fillStyle = col; g.fill(); g.globalAlpha = 1; }
      g.lineWidth = st.lw; g.strokeStyle = col; g.setLineDash(st.dash); g.lineJoin = 'round'; g.lineCap = 'round'; g.stroke(); g.setLineDash([]);
      pts.forEach(([x, y]) => marker(x, y, st.m, col));
    });
    g.font = `600 32px ${CANVAS_FONT}`; g.textBaseline = 'middle'; g.textAlign = 'left';
    const items = series.map(s => ({ s, w: 90 + g.measureText(s.name).width }));
    const gap = 46, total = items.reduce((a, b) => a + b.w, 0) + gap * Math.max(0, items.length - 1);
    let x = (w - total) / 2; const y = h - 70;
    items.forEach(({ s, w: iw }) => {
      const st = style[s.key], col = HEX(s.key);
      g.strokeStyle = col; g.lineWidth = st.lw; g.setLineDash(st.dash); g.beginPath(); g.moveTo(x, y); g.lineTo(x + 64, y); g.stroke(); g.setLineDash([]);
      marker(x + 32, y, st.m, col);
      g.fillStyle = '#1F2937'; g.fillText(s.name, x + 82, y + 2); x += iw + gap;
    });
    return c;
  }
  function drawBars(series, labels) {
    const w = 1500, h = 660, c = mkCanvas(w, h), g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, w, h);
    const L = 70, Rm = 20, T = 80, B = 100, pw = w - L - Rm, ph = h - T - B, n = labels.length, gw = pw / n;
    const Y = v => T + ph - ph * v / 5;
    g.font = `500 26px ${CANVAS_FONT}`; g.textAlign = 'right'; g.textBaseline = 'middle';
    for (let v = 0; v <= 5; v++) {
      g.strokeStyle = v === 3 ? '#8FA0B5' : '#E3E8EE'; g.lineWidth = v === 3 ? 3 : 2; g.setLineDash(v === 3 ? [14, 10] : []);
      g.beginPath(); g.moveTo(L, Y(v)); g.lineTo(w - Rm, Y(v)); g.stroke();
      g.fillStyle = '#7B8794'; g.fillText(String(v), L - 12, Y(v));
    }
    g.setLineDash([]);
    const m = series.length, bw = Math.min(74, gw * 0.74 / m), gap = 10, gtot = bw * m + gap * (m - 1);
    labels.forEach((lab, k) => {
      const x0 = L + gw * k + (gw - gtot) / 2;
      series.forEach((s, j) => {
        const v = s.values[k]; if (v == null) return;
        const x = x0 + j * (bw + gap), y = Y(v);
        g.fillStyle = HEX(s.key); g.fillRect(x, y, bw, Y(0) - y);
        g.fillStyle = '#1F2937'; g.font = `700 28px ${CANVAS_FONT}`; g.textAlign = 'center'; g.textBaseline = 'bottom';
        g.fillText(fmt(v), x + bw / 2, y - 6);
      });
      g.font = `700 30px ${CANVAS_FONT}`; g.fillStyle = '#1F2937'; g.textAlign = 'center'; g.textBaseline = 'top';
      let lines = [lab];
      if (g.measureText(lab).width > gw - 14) { const cut = Math.max(lab.lastIndexOf(' '), lab.indexOf('-')); if (cut > 0) lines = [lab.slice(0, cut + (lab[cut] === '-' ? 1 : 0)).trim(), lab.slice(cut + 1).trim()]; }
      lines.forEach((t, i) => g.fillText(t, L + gw * k + gw / 2, Y(0) + 14 + i * 34));
    });
    g.font = `600 30px ${CANVAS_FONT}`; g.textBaseline = 'middle'; g.textAlign = 'left';
    let x = L; const y = 36;
    series.forEach(s => { g.fillStyle = HEX(s.key); g.fillRect(x, y - 13, 28, 26); g.fillStyle = '#1F2937'; g.fillText(s.name, x + 40, y + 1); x += 40 + g.measureText(s.name).width + 50; });
    return c;
  }

  // ---------- สร้างเอกสาร ----------
  async function build(inp) {
    const D = root.docx;
    if (!D) throw new Error('ยังไม่ได้โหลดไลบรารี docx');
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, AlignmentType,
      VerticalAlign, TabStopType, Tab, PageNumber, Header, Footer, ImageRun, LevelFormat, HeightRule, TableLayoutType, HeadingLevel } = D;
    const a = analyze(inp);
    if (!a.any) throw new Error('ยังไม่มีข้อมูลให้สรุป (ยังไม่มีทั้งผลข้อทดสอบ ประเมินตนเอง และคะแนนจาก VP)');
    const p = inp.person, IDX = inp.indexes, rubric = inp.rubric || {};
    const when = inp.generatedAt ? new Date(inp.generatedAt) : new Date();

    const series = ['oat', 'scn', 'self'].filter(k => a.has[k]).map(k => ({
      key: k, name: LEGEND[k], values: a.rows.map(r => k === 'oat' ? r.V : k === 'scn' ? r.T : r.S) }));
    const radarPng = await toPng(drawRadar(series, IDX));
    const barPng = await toPng(drawBars(series, IDX));

    // --- helpers ---
    const R = (text, o = {}) => new TextRun(Object.assign({ text: String(text), font: FONTS }, o));
    const pt = n => n * 2;
    const P = (kids, o = {}) => new Paragraph(Object.assign({ children: Array.isArray(kids) ? kids : [kids] }, o));
    const T = (text, size, o = {}) => P(R(text, { size: pt(size), bold: o.bold, italics: o.italics, color: o.color }),
      { alignment: o.align, spacing: Object.assign({ before: 0, after: 0 }, o.spacing || {}), keepNext: o.keepNext, keepLines: true });
    const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const noB = { top: none, bottom: none, left: none, right: none };
    const hair = { style: BorderStyle.SINGLE, size: 4, color: C.line };
    const fill = f => ({ type: ShadingType.CLEAR, fill: f, color: 'auto' });
    const cell = (kids, w, o = {}) => new TableCell(Object.assign({
      width: { size: w, type: WidthType.DXA }, children: Array.isArray(kids) ? kids : [kids],
      margins: { top: 70, bottom: 70, left: 110, right: 110 }, borders: noB }, o));
    const table = (colW, rows, o = {}) => new Table(Object.assign({
      width: { size: colW.reduce((x, y) => x + y, 0), type: WidthType.DXA }, columnWidths: colW, layout: TableLayoutType.FIXED, rows }, o));
    const gapP = h => P([], { spacing: { before: 0, after: 0, line: h || 120, lineRule: 'exact' } });
    const bullet = (text, o = {}) => P(R(text, { size: pt(o.size || 14.5), color: o.color }), { numbering: { reference: 'bul', level: 0 }, spacing: { before: 0, after: 40 }, keepLines: true, keepNext: o.keepNext });
    const label = text => T(text, 12.5, { bold: true, color: C.muted, spacing: { before: 100, after: 20 }, keepNext: true });
    const h1 = (text, brk) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: !!brk, children: [R(text, { size: pt(20), bold: true, color: C.navy })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.navy, space: 4 } }, spacing: { before: brk ? 0 : 280, after: 140 }, keepNext: true });
    const h2 = text => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [R(text, { size: pt(16), bold: true, color: C.navy })], spacing: { before: 220, after: 90 }, keepNext: true });
    const img = (data, wpx, hpx, alt) => new ImageRun({ type: 'png', data, transformation: { width: wpx, height: hpx }, altText: { title: alt, description: alt, name: alt } });

    // --- ส่วนหัวเอกสาร ---
    const title = table([W], [new TableRow({ children: [cell([
      T('ONE-ON-ONE REVIEW  ·  5 KEY INDEXES', 11.5, { color: 'B9C6D8', bold: true, spacing: { after: 40 } }),
      T('สรุปผลการประเมิน 5 Key Indexes', 27, { color: 'FFFFFF', bold: true, spacing: { after: 80 } }),
      T(p.pmName, 20, { color: 'FFFFFF', bold: true }),
      T('Project Manager  ·  Government 1', 14, { color: 'B9C6D8', spacing: { before: 20 } })
    ], W, { shading: fill(C.navy), margins: { top: 300, bottom: 300, left: 380, right: 380 } })] })]);

    const mw = [2409, 2409, 2409, 2411];
    const metaCell = (lab, val, w) => cell([T(lab, 11.5, { color: C.muted, bold: true }), T(val, 14, { bold: true })], w, { shading: fill(C.soft) });
    const meta = table(mw, [new TableRow({ children: [
      metaCell('รอบประเมิน', String(inp.cycle || '–'), mw[0]),
      metaCell('ส่งแบบประเมิน', p.submitted ? thDate(p.submittedAt) : 'ยังไม่ได้ส่ง', mw[1]),
      metaCell('VP ให้คะแนน', a.has.oat ? thDate(p.oatAt) : 'ยังไม่ได้ประเมิน', mw[2]),
      metaCell('จัดทำเอกสาร', thDate(when), mw[3])] })]);

    const miss = a.missing.map(k => SRC[k]);
    const banner = miss.length ? table([W], [new TableRow({ children: [cell([
      T(`ข้อมูลยังไม่ครบ: ยังไม่มี${miss.join(' และ ')}`, 14, { bold: true, color: '8A5A00' }),
      T('ค่า Avg. ในเอกสารนี้จึงเป็นค่าชั่วคราวที่เฉลี่ยจากข้อมูลที่มีอยู่ (ทำเครื่องหมาย *) เมื่อข้อมูลครบ ให้ดาวน์โหลดเอกสารใหม่จากระบบ', 13.5, { color: '8A5A00' })
    ], W, { shading: fill(C.warn), margins: { top: 90, bottom: 90, left: 180, right: 140 },
      borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color: C.warnLine } } })] })]) : null;

    // --- 1. ภาพรวม ---
    const o = a.overall;
    const tw = [2409, 2409, 2409, 2411];
    const gapB = { style: BorderStyle.SINGLE, size: 36, color: 'FFFFFF' };
    const tile = (lab, sub, val, w, kind) => {
      const dark = kind === 'avg', top = dark ? C.navy : C[kind];
      const fg = dark ? 'FFFFFF' : C.ink, fgm = dark ? 'C9D5E6' : C.muted;
      return cell([T(lab, 13, { bold: true, color: fgm }),
        P([R(fmt(val), { size: pt(30), bold: true, color: fg }), R(val == null ? '' : ' / 5', { size: pt(13), color: fgm })], { spacing: { before: 20, after: 0 } }),
        T(sub, 11.5, { color: fgm })], w, { shading: fill(dark ? C.navy : C.soft), margins: { top: 110, bottom: 110, left: 150, right: 100 },
        borders: { top: { style: BorderStyle.SINGLE, size: 24, color: top }, bottom: none, left: gapB, right: gapB } });
    };
    const avgBasis = ['oat', 'scn', 'self'].filter(k => a.tick[k]).map(k => SRC[k]).join(' + ') || '–';
    const tiles = table(tw, [new TableRow({ children: [
      tile('Avg. รวม', avgBasis, o.A, tw[0], 'avg'), tile('VP', 'เห็นจากงานจริง', o.V, tw[1], 'oat'),
      tile('ข้อทดสอบ', 'โจทย์สถานการณ์', o.T, tw[2], 'scn'), tile('ตนเอง', 'ประเมินเอง', o.S, tw[3], 'self')] })]);

    const sumKids = [];
    const sumBlock = (lab, lines) => { sumKids.push(label(lab)); lines.forEach(l => sumKids.push(T(l, 13.5, { spacing: { after: 30 } }))); };
    if (a.flat) sumBlock('ภาพรวมรายด้าน', [`ทั้ง 5 ด้านอยู่ในระดับใกล้เคียงกัน (ห่างกันไม่ถึง 0.3) ไม่มีด้านใดเด่นหรือด้อยกว่าด้านอื่นชัดเจน`]);
    else if (a.ranked.length) {
      sumBlock('จุดเด่น', [a.top.map(r => `${r.index} (${fmt(r.A)})`).join(', ')]);
      const lw = a.low.length ? a.low[0].Araw : null;
      const q = lw == null ? '' : lw < 3 ? 'ต่ำกว่าระดับมาตรฐาน (ระดับ 3) ควรให้ความสำคัญเป็นลำดับแรก'
        : lw < 4 ? 'ยังมีช่องว่างให้พัฒนา' : 'อยู่ในเกณฑ์ดี แต่ยังเป็นด้านที่ยกระดับต่อได้';
      const lines = [a.low.map(r => `${r.index} (${fmt(r.A)})`).join(', '), q].filter(Boolean);
      if (a.vpLow.length) lines.push(`VP ให้ระดับ 2 หรือต่ำกว่า: ${names(a.vpLow)}`);
      sumBlock('ควรให้ความสำคัญ', lines);
    }
    const f = a.flag, gl = [];
    if (a.has.oat) {
      if (f.sUp.length) gl.push(`ตนเองสูงกว่า VP: ${names(f.sUp)}`);
      if (f.sDown.length) gl.push(`ตนเองต่ำกว่า VP: ${names(f.sDown)}`);
      if (f.tUp.length) gl.push(`ข้อทดสอบสูงกว่า VP: ${names(f.tUp)}`);
      if (f.tDown.length) gl.push(`ข้อทดสอบต่ำกว่า VP: ${names(f.tDown)}`);
      const nOther = (a.has.scn ? 1 : 0) + (a.has.self ? 1 : 0);
      sumBlock('ความต่างของมุมมองที่ควรคุย', gl.length ? gl : nOther === 0 ? ['ยังไม่มีข้อมูลจากข้อทดสอบและตนเอง จึงยังเทียบมุมมองไม่ได้']
        : [(nOther === 2 ? 'สามมุมมอง' : 'มุมมองที่มีอยู่') + 'สอดคล้องกัน ต่างกันไม่ถึง 1 ระดับในทุกด้าน']);
    } else {
      if (f.stUp.length) gl.push(`ตนเองสูงกว่าข้อทดสอบ: ${names(f.stUp)}`);
      if (f.stDown.length) gl.push(`ตนเองต่ำกว่าข้อทดสอบ: ${names(f.stDown)}`);
      sumBlock('ความต่างของมุมมองที่ควรคุย', ['VP ยังไม่ได้ให้คะแนน จึงยังเทียบกับ VP ไม่ได้'].concat(gl));
    }
    const rw = [4900, 4738];
    const radarRow = table(rw, [new TableRow({ cantSplit: true, children: [
      cell([P(img(radarPng, 318, 297, 'กราฟเรดาร์ 5 Key Indexes'), { alignment: AlignmentType.CENTER })], rw[0], { margins: { top: 0, bottom: 0, left: 0, right: 0 } }),
      cell(sumKids, rw[1], { margins: { top: 0, bottom: 0, left: 200, right: 0 }, verticalAlign: VerticalAlign.TOP })] })]);

    // --- 2. ตารางคะแนน ---
    const cw = [2750, 800, 1250, 950, 900, 1494, 1494];
    const hcell = (txt, w, left) => cell([T(txt, 13, { bold: true, color: 'FFFFFF', align: left ? AlignmentType.LEFT : AlignmentType.CENTER })], w,
      { shading: fill(C.navy), verticalAlign: VerticalAlign.CENTER, margins: { top: 90, bottom: 90, left: 70, right: 70 } });
    const vcell = (v, w, o2 = {}) => cell([T(fmt(v) + (o2.star ? '*' : ''), 15, { align: AlignmentType.CENTER, bold: o2.bold })], w,
      { shading: fill(o2.shade === undefined ? shade(v) : o2.shade), verticalAlign: VerticalAlign.CENTER, borders: Object.assign({}, noB, { bottom: { style: BorderStyle.SINGLE, size: 8, color: 'FFFFFF' } }) });
    const dcell = (d, w) => cell([T(sg(d), 14.5, { align: AlignmentType.CENTER, bold: d != null && Math.abs(d) >= 1,
      color: d == null ? C.muted : Math.abs(d) < 1 ? C.ink : d > 0 ? C.up : C.down })], w, { verticalAlign: VerticalAlign.CENTER, borders: { top: none, left: none, right: none, bottom: hair } });
    const scoreRows = [new TableRow({ tableHeader: true, cantSplit: true, children: [hcell('Index', cw[0], true), hcell('VP', cw[1]), hcell('ข้อทดสอบ', cw[2]), hcell('ตนเอง', cw[3]),
      hcell('Avg.', cw[4]), hcell('Δ ข้อทดสอบ-VP', cw[5]), hcell('Δ ตนเอง-VP', cw[6])] })];
    a.rows.forEach(r => scoreRows.push(new TableRow({ cantSplit: true, children: [
      cell([T(r.index, 15, { bold: true })], cw[0], { verticalAlign: VerticalAlign.CENTER, borders: { top: none, left: none, right: none, bottom: hair } }),
      vcell(r.V, cw[1]), vcell(r.T, cw[2]), vcell(r.S, cw[3]), vcell(r.A, cw[4], { bold: true, star: r.partial, shade: 'FFFFFF' }), dcell(r.dT, cw[5]), dcell(r.dS, cw[6])] })));
    scoreRows.push(new TableRow({ cantSplit: true, children: [
      cell([T('เฉลี่ยทั้ง 5 ด้าน', 15, { bold: true })], cw[0], { shading: fill(C.soft2), verticalAlign: VerticalAlign.CENTER }),
      vcell(o.V, cw[1], { shade: C.soft2, bold: true }), vcell(o.T, cw[2], { shade: C.soft2, bold: true }), vcell(o.S, cw[3], { shade: C.soft2, bold: true }),
      vcell(o.A, cw[4], { shade: C.soft2, bold: true, star: a.rows.some(r => r.partial) }),
      cell([T(sg(o.dT), 14.5, { align: AlignmentType.CENTER, bold: true })], cw[5], { shading: fill(C.soft2), verticalAlign: VerticalAlign.CENTER }),
      cell([T(sg(o.dS), 14.5, { align: AlignmentType.CENTER, bold: true })], cw[6], { shading: fill(C.soft2), verticalAlign: VerticalAlign.CENTER })] }));
    const scoreTbl = table(cw, scoreRows);
    const avgNote = `Avg. = ค่าเฉลี่ยของ ${avgBasis} (ตามที่เลือกไว้ในหน้าจอ)` + (a.rows.some(r => r.partial) ? ` · * เฉลี่ยจากค่าที่มีอยู่ ยังไม่มี${miss.join(' และ ')}` : '');

    const guide = [
      ['VP', 'คะแนน 1-5 ที่ VP ให้ตาม rubric จากงานที่เห็นจริง'],
      ['ข้อทดสอบ', 'ค่าเฉลี่ยของโจทย์สถานการณ์ด้านละ 3 ข้อ (ระดับ 1-5) วัดว่าเลือกวิธีรับมืออย่างไรในสถานการณ์สมมติ'],
      ['ตนเอง', 'ระดับที่ผู้รับการประเมินเลือกเอง ตาม rubric เดียวกับ VP'],
      ['Δ', 'ผลต่างเทียบกับ VP (+ สูงกว่า VP, - ต่ำกว่า VP) ตัวหนาและมีสีเมื่อต่างตั้งแต่ 1 ระดับ']
    ];
    const guideBox = table([W], [new TableRow({ cantSplit: true, children: [cell([
      T('วิธีอ่านผล', 13.5, { bold: true, color: C.navy, spacing: { after: 40 } }),
      ...guide.map(([k, v]) => P([R(k + '  ', { size: pt(13.5), bold: true }), R(v, { size: pt(13.5) })], { spacing: { after: 20 }, keepLines: true })),
      T('สามค่านี้วัดคนละแบบ จึงใช้ประกอบกันเพื่อชวนคุย ไม่ควรใช้ค่าใดค่าหนึ่งตัดสินเดี่ยว ๆ', 13.5, { italics: true, color: C.muted, spacing: { before: 40 } })
    ], W, { shading: fill(C.soft), margins: { top: 110, bottom: 110, left: 180, right: 160 } })] })]);

    // --- 3. การ์ดรายด้าน ---
    const items = Array.isArray(p.items) ? p.items : null;
    const cardW = W;
    const cards = [];
    a.rows.forEach((r, ix) => {
      const rub = (lv) => ((rubric[r.index] || [])[Math.round(lv) - 1]) || '';
      const cl = 5100, cr = cardW - cl;
      const navyL = { style: BorderStyle.SINGLE, size: 36, color: C.navy };
      const head = cell([P([R(`${ix + 1}  `, { size: pt(17), bold: true, color: C.muted }), R(r.index, { size: pt(17), bold: true, color: C.navy }),
        new TextRun({ children: [new Tab(), `VP ${fmt(r.V)}  ·  ข้อทดสอบ ${fmt(r.T)}  ·  ตนเอง ${fmt(r.S)}  ·  Avg. ${fmt(r.A)}${r.partial ? '*' : ''}`], font: FONTS, size: pt(13), color: C.muted })],
        { tabStops: [{ type: TabStopType.RIGHT, position: cardW - 420 }], spacing: { after: 0 }, keepNext: true })], cardW, {
        columnSpan: 2, shading: fill(C.soft), margins: { top: 110, bottom: 110, left: 220, right: 160 },
        borders: { top: none, right: none, bottom: none, left: navyL } });

      // ซ้าย: ข้อสังเกต + เกณฑ์ rubric
      const L = [label('ข้อสังเกตจากข้อมูล')];
      insightsOf(r).forEach(t => L.push(bullet(t, { size: 13, keepNext: true })));
      const rl = [];
      if (r.V != null && r.S != null && Math.round(r.V) === Math.round(r.S)) rl.push(['VP และตนเอง', r.V, rub(r.V)]);
      else { if (r.V != null) rl.push(['VP', r.V, rub(r.V)]); if (r.S != null) rl.push(['ตนเอง', r.S, rub(r.S)]); }
      if (rl.length) {
        L.push(label('เกณฑ์ตาม rubric ของระดับที่ให้/เลือก'));
        rl.forEach(([who, lv, tx]) => L.push(P([R(`${who} · ระดับ ${fmt(lv)}  `, { size: pt(12.5), bold: true, color: C.navy }), R(tx, { size: pt(12.5) })], { spacing: { after: 30 }, keepLines: true, keepNext: true })));
      }
      // ขวา: ข้อทดสอบรายข้อ + ประเด็นชวนคุย
      const Rt = [];
      const its = (items || []).filter(x => x.index === r.index);
      if (its.length) {
        Rt.push(label('ข้อทดสอบรายข้อ'));
        const lows = its.filter(x => x.score <= 2).map(x => x.id);
        Rt.push(T(its.map(x => `${x.id} = ${x.score}`).join('   ·   ') + (lows.length ? `  (ระดับ 1-2: ${lows.join(', ')} ดูภาคผนวก)` : ''), 13.5, { spacing: { after: 20 }, keepNext: true }));
      }
      Rt.push(label('ประเด็นชวนคุย'));
      promptsOf(r).forEach(t => Rt.push(bullet(t, { size: 13, keepNext: true })));

      const body = [
        cell(L, cl, { margins: { top: 40, bottom: 80, left: 220, right: 140 }, borders: { top: none, right: none, bottom: none, left: navyL } }),
        cell(Rt, cr, { margins: { top: 40, bottom: 80, left: 140, right: 160 }, borders: { top: none, right: none, bottom: none, left: none } })];
      const foot = cell([P(R('ความเห็นเพิ่มเติมของ VP:  (พิมพ์ที่นี่)', { size: pt(13), italics: true, color: '9CA3AF' }), { spacing: { before: 0, after: 0 } })], cardW, {
        columnSpan: 2, shading: fill(C.note), margins: { top: 80, bottom: 80, left: 220, right: 160 },
        borders: { top: none, right: none, bottom: hair, left: navyL } });
      cards.push(table([cl, cr], [new TableRow({ cantSplit: true, children: [head] }), new TableRow({ cantSplit: true, children: body }), new TableRow({ cantSplit: true, children: [foot] })]));
      cards.push(gapP(160));
    });

    // --- 4. บันทึกการพูดคุย ---
    const nw = [2700, W - 2700];
    const noteRow = (lab, i, arr) => new TableRow({ cantSplit: true, height: { value: 1050, rule: HeightRule.ATLEAST }, children: [
      cell([T(lab, 14.5, { bold: true, color: C.navy, keepNext: i < arr.length - 1 })], nw[0], { shading: fill(C.soft), verticalAlign: VerticalAlign.TOP, margins: { top: 100, bottom: 100, left: 140, right: 110 }, borders: { top: none, left: none, right: none, bottom: { style: BorderStyle.SINGLE, size: 8, color: 'FFFFFF' } } }),
      cell([T('(พิมพ์บันทึกที่นี่)', 13.5, { italics: true, color: '9CA3AF', keepNext: i < arr.length - 1 })], nw[1], { shading: fill(C.note), verticalAlign: VerticalAlign.TOP, margins: { top: 100, bottom: 100, left: 160, right: 140 }, borders: { top: none, left: none, right: none, bottom: { style: BorderStyle.SINGLE, size: 8, color: 'FFFFFF' } } })] });
    const noteTbl = table(nw, ['ความเห็นโดยรวมของ VP', 'งานที่ภูมิใจที่สุด', 'เรื่องที่ท้าทายที่สุด', 'ข้อจำกัด / สิ่งที่ต้องการให้สนับสนุน', 'ข้อสรุปร่วมกัน'].map(noteRow));

    // --- ภาคผนวก ---
    let appendix = [];
    if (items && items.length) {
      const aw = [650, 1650, 800, 3250, 3288];
      const rowsA = [new TableRow({ tableHeader: true, cantSplit: true, children: [hcell('ข้อ', aw[0]), hcell('Index', aw[1], true), hcell('ระดับ', aw[2]), hcell('โจทย์สถานการณ์', aw[3], true), hcell('ตัวเลือกที่เลือก', aw[4], true)] })];
      const order = IDX.reduce((m, k, i) => (m[k] = i, m), {});
      items.slice().sort((x, y) => (order[x.index] - order[y.index]) || String(x.id).localeCompare(String(y.id))).forEach(x => rowsA.push(new TableRow({ cantSplit: true, children: [
        cell([T(x.id, 13, { align: AlignmentType.CENTER })], aw[0], { borders: { top: none, left: none, right: none, bottom: hair } }),
        cell([T(x.index, 13, { bold: true })], aw[1], { borders: { top: none, left: none, right: none, bottom: hair } }),
        cell([T(String(x.score), 15, { align: AlignmentType.CENTER, bold: true })], aw[2], { shading: fill(shade(x.score)), verticalAlign: VerticalAlign.CENTER, borders: { top: none, left: none, right: none, bottom: { style: BorderStyle.SINGLE, size: 8, color: 'FFFFFF' } } }),
        cell([T(x.q, 13)], aw[3], { borders: { top: none, left: none, right: none, bottom: hair } }),
        cell([T(x.a, 13)], aw[4], { borders: { top: none, left: none, right: none, bottom: hair } })] })));
      appendix = [h1('ภาคผนวก  คำตอบข้อทดสอบรายข้อ', true),
        T('ระดับของแต่ละข้อมาจากตัวเลือกที่เลือก (1-5) คะแนนข้อทดสอบรายด้านคือค่าเฉลี่ยของข้อในด้านนั้น ใช้ดูว่าข้อไหนที่ควรหยิบมาคุยต่อ', 13.5, { color: C.muted, spacing: { after: 120 } }),
        table(aw, rowsA)];
    }

    // --- ประกอบเอกสาร ---
    const hdr = new Header({ children: [P([R(`PM-GOV1  ·  One-on-One 5 Key Indexes  ·  ${p.pmName}`, { size: pt(11), color: '9CA3AF' })], { alignment: AlignmentType.RIGHT })] });
    const ftr = new Footer({ children: [P([R('เอกสารสำหรับการพูดคุย 1:1  ·  เป็นความลับ', { size: pt(11), color: '9CA3AF' }),
      new TextRun({ children: [new Tab(), 'หน้า ', PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: FONTS, size: pt(11), color: '9CA3AF' })],
      { tabStops: [{ type: TabStopType.RIGHT, position: W }], border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.line, space: 6 } } })] });
    const blank = new Header({ children: [P([])] });

    const children = [title, gapP(100), meta, gapP(120)];
    if (banner) children.push(banner, gapP(100));
    children.push(h1('1.  ภาพรวม'), tiles, gapP(160), radarRow,
      h1('2.  คะแนนรายด้านและความต่างของมุมมอง', true), scoreTbl,
      P(R(avgNote, { size: pt(12.5), color: C.muted }), { spacing: { before: 60, after: 100 } }),
      h2('เทียบ 3 มุมมองรายด้าน'), P(img(barPng, 600, 264, 'กราฟแท่งเทียบ VP ข้อทดสอบ ตนเอง รายด้าน'), { alignment: AlignmentType.CENTER, spacing: { after: 100 } }),
      guideBox,
      h1('3.  วิเคราะห์รายด้าน'),
      T('ข้อสังเกตด้านล่างสร้างจากตัวเลขโดยอัตโนมัติ เพื่อใช้เป็นจุดเริ่มต้นของการพูดคุย ไม่ใช่ข้อสรุป แก้ไขหรือเพิ่มความเห็นได้ตามต้องการ', 13.5, { color: C.muted, spacing: { after: 140 } }),
      ...cards,
      h1('4.  บันทึกการพูดคุย 1:1'), noteTbl, ...appendix);

    const doc = new Document({
      creator: 'PM-GOV1 Dashboard', title: `สรุปผลการประเมิน 5 Key Indexes - ${p.pmName}`, description: 'One-on-One 5 Key Indexes',
      styles: { default: { document: { run: { font: FONTS, size: pt(15), color: C.ink, language: { value: 'th-TH', bidirectional: 'th-TH' } }, paragraph: { spacing: { line: 276, lineRule: 'auto' } } } } },
      numbering: { config: [{ reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 340, hanging: 220 } }, run: { color: C.navy } } }] }] },
      sections: [{
        properties: { titlePage: true, page: { size: { width: 11906, height: 16838 }, margin: { top: 1247, bottom: 1134, left: 1134, right: 1134, header: 567, footer: 510 } } },
        headers: { default: hdr, first: blank }, footers: { default: ftr, first: ftr }, children }]
    });
    return Packer.toBlob(doc);
  }

  root.AsmReport = { build, analyze, insightsOf, promptsOf };
})(typeof window !== 'undefined' ? window : globalThis);
