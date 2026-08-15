// สร้าง AUTH_USERS ครบทั้งทีมในคำสั่งเดียว แล้วนำผลลัพธ์ไปวางใน Render → Environment
//
// วิธีใช้ (รันในเครื่องตัวเอง รหัสผ่านจริงไม่ถูกส่งออกไปไหน):
//   cd backend
//   npm install
//   node setup-users.js
//
// สคริปต์จะถามรหัสผ่านทีละคน (พิมพ์แล้วกด Enter) แล้วพ่น JSON บรรทัดเดียวออกมาให้ copy ไปวางได้เลย
//
// หมายเหตุเรื่องการมองเห็นข้อมูล:
//   - admin        -> เห็นทุกโครงการของทุก PM
//   - PM แต่ละคน   -> เห็นเฉพาะโครงการที่ตัวเองเป็นเจ้าของ (ผูกด้วยฟิลด์ pmName)
//   ค่า pmName ด้านล่างถูกตั้งให้ตรงกับคอลัมน์ "PM Name" ในไฟล์ Excel แล้ว ไม่ต้องแก้

const bcrypt = require('bcryptjs');
const readline = require('readline');

// รายชื่อทีม PM-GOV1 — แก้ username ได้ตามสะดวก แต่ pmName ต้องตรงกับใน Excel
const TEAM = [
  { username: 'oat',       name: 'OAT',                    role: 'admin' },
  { username: 'antika',    name: 'Antika Prasadsil',       pmName: 'Antika Prasadsil' },
  { username: 'lomdetch',  name: 'Lomdetch Puangsombut',   pmName: 'Lomdetch Puangsombut' },
  { username: 'virojt',    name: 'Virojt Changyencham',    pmName: 'Virojt Changyencham' },
  { username: 'prapasiri', name: 'Prapasiri Rakkanpat',    pmName: 'Prapasiri Rakkanpat' },
  { username: 'watanyu',   name: 'Watanyu Anantakunakorn', pmName: 'Watanyu Anantakunakorn' },
];

// อ่านบรรทัดจาก stdin ทีละบรรทัด รองรับทั้งพิมพ์เองในเทอร์มินัล และแบบ pipe เข้ามา
function makeReader() {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: process.stdin.isTTY });
  const queue = [];
  const waiting = [];
  let ended = false;
  rl.on('line', (line) => {
    if (waiting.length) waiting.shift()(line);
    else queue.push(line);
  });
  rl.on('close', () => { ended = true; while (waiting.length) waiting.shift()(''); });
  return {
    ask(prompt) {
      process.stdout.write(prompt);
      if (queue.length) { const v = queue.shift(); if (!process.stdin.isTTY) process.stdout.write('\n'); return Promise.resolve(v); }
      if (ended) return Promise.resolve('');
      return new Promise(resolve => waiting.push(v => { if (!process.stdin.isTTY) process.stdout.write('\n'); resolve(v); }));
    },
    close(){ rl.close(); },
  };
}

(async () => {
  console.log('\n=== สร้างบัญชีผู้ใช้ PM-GOV1 Dashboard ===');
  console.log('กรอกรหัสผ่านของแต่ละคน (เว้นว่างแล้วกด Enter = ข้ามคนนั้น ไม่สร้างบัญชี)\n');

  const reader = makeReader();
  const users = [];
  for (const member of TEAM) {
    const scope = member.role === 'admin' ? 'เห็นทุกโครงการ' : 'เห็นเฉพาะโครงการของตัวเอง';
    const pw = (await reader.ask(`รหัสผ่านของ ${member.name} [${member.username}] (${scope}): `)).trim();
    if (!pw) { console.log(`  ↳ ข้าม ${member.username}\n`); continue; }
    if (pw.length < 6) console.log('  ⚠ รหัสผ่านสั้นกว่า 6 ตัวอักษร แนะนำให้ตั้งยาวกว่านี้');
    const entry = { username: member.username, name: member.name, passwordHash: bcrypt.hashSync(pw, 10) };
    if (member.role) entry.role = member.role;
    if (member.pmName) entry.pmName = member.pmName;
    users.push(entry);
    console.log('  ✓ สร้างแล้ว\n');
  }
  reader.close();

  if (!users.length) { console.log('ไม่ได้สร้างบัญชีใดเลย'); return; }

  console.log('\n──────────────────────────────────────────────');
  console.log('คัดลอกทั้งบรรทัดด้านล่างนี้ไปวางใน Render:');
  console.log('  Render Dashboard → pm-gov1-progress-api → Environment');
  console.log('  → Add Environment Variable → Key: AUTH_USERS → Value: (วางด้านล่าง)');
  console.log('──────────────────────────────────────────────\n');
  console.log(JSON.stringify(users));
  console.log(`\n──────────────────────────────────────────────`);
  console.log(`สร้างทั้งหมด ${users.length} บัญชี: ${users.map(u => u.username).join(', ')}`);
  console.log('อย่าลืมกด Save Changes แล้วรอ Render restart ให้เสร็จก่อนทดสอบ login\n');
})();
