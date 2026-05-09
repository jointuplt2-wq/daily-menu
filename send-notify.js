const fs = require('fs');
const https = require('https');

const TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const MEAL = process.env.MEAL; // 아침 | 점심 | 저녁

if (!TOKEN || !CHAT_ID || !MEAL) {
  console.error('❌ 환경변수 누락: TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, MEAL 필요');
  process.exit(1);
}

function pad(n) { return String(n).padStart(2, '0'); }
function todayKST() {
  const d = new Date(Date.now() + 9 * 60 * 60 * 1000);
  return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
}

const menuData = JSON.parse(fs.readFileSync('menu.json', 'utf-8'));
const today = todayKST();
const dayMenu = menuData[today];

const icon = MEAL === '아침' ? '🌅' : MEAL === '점심' ? '☀️' : '🌙';
const DAYS = ['일', '월', '화', '수', '목', '금', '토'];
const d = new Date(today);
const dayStr = DAYS[d.getDay()] + '요일';

let text;
if (!dayMenu || !dayMenu[MEAL] || !dayMenu[MEAL].length) {
  text = `${icon} *${MEAL} 메뉴* (${today.slice(5)} ${dayStr})\n\n등록된 메뉴가 없습니다.`;
} else {
  const items = dayMenu[MEAL].map(item => `• ${item}`).join('\n');
  text = `${icon} *경기케어 ${MEAL} 메뉴* (${today.slice(5)} ${dayStr})\n\n${items}\n\n🔗 [메뉴 보기](https://jointuplt2-wq.github.io/daily-menu/)`;
}

const body = JSON.stringify({
  chat_id: CHAT_ID,
  text,
  parse_mode: 'Markdown',
  disable_web_page_preview: false
});

const req = https.request({
  hostname: 'api.telegram.org',
  path: `/bot${TOKEN}/sendMessage`,
  method: 'POST',
  headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
}, res => {
  let data = '';
  res.on('data', chunk => data += chunk);
  res.on('end', () => {
    const result = JSON.parse(data);
    if (result.ok) console.log(`✅ ${MEAL} 메뉴 텔레그램 발송 완료`);
    else { console.error('❌ 발송 실패:', result.description); process.exit(1); }
  });
});
req.on('error', e => { console.error('❌ 요청 오류:', e.message); process.exit(1); });
req.write(body);
req.end();
