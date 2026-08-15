// เครื่องมือช่วยสร้าง bcrypt hash สำหรับตั้งรหัสผ่านผู้ใช้งานจริง
// รันในเครื่องตัวเอง (ไม่ต้องส่งรหัสผ่านจริงไปให้ใครดู) แล้วเอา hash ที่ได้ไปใส่ใน
// environment variable AUTH_USERS บน Render
//
// วิธีใช้:
//   cd backend
//   npm install
//   node generate-hash.js "รหัสผ่านที่ต้องการ"
//
// ตัวอย่างผลลัพธ์ที่ได้ เอาไปประกอบเป็น AUTH_USERS แบบนี้ (ใส่ใน Render → Environment):
//   [
//     {"username":"antika",  "name":"Antika Prasadsil",     "passwordHash":"<hash ที่ได้>"},
//     {"username":"lomdetch","name":"Lomdetch Puangsombut", "passwordHash":"<hash ที่ได้>"}
//   ]
// (ต้องเป็น JSON string บรรทัดเดียวตอนวางใน environment variable จริง)

const bcrypt = require('bcryptjs');

const password = process.argv[2];
if (!password) {
  console.log('วิธีใช้: node generate-hash.js "รหัสผ่านที่ต้องการ"');
  process.exit(1);
}

const hash = bcrypt.hashSync(password, 10);
console.log('\nPassword hash (bcrypt):');
console.log(hash);
console.log('\nเอา hash ด้านบนไปใส่ในฟิลด์ "passwordHash" ของ AUTH_USERS บน Render ครับ\n');
