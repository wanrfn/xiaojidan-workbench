/* =========================================================
   小煎蛋的工作台  —  个人工作台 (本地优先 / PWA)
   数据全部存于浏览器 localStorage；可选接入同步服务实现多端实时同步
   ========================================================= */
'use strict';

/* ---------- 基础工具 ---------- */
const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const KEY = 'xiaojidan_workbench_v1';
const APP_VERSION = '20260920c'; // 缓存破版本号：每次改 JS 必须递增，并同步 index.html 的 ?v=

const todayStr = (d = new Date()) => {
  const z = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
};
const startOfDay = d => { d = new Date(d); d.setHours(0, 0, 0, 0); return d; };
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const WEEK = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];

function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove('show'), 1800);
}
function download(filename, content, type) {
  const blob = content instanceof Blob ? content : new Blob([content], { type: type || 'text/markdown' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = filename; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
/* ---------- 生成 .docx（纯 JS 打包，无需任何依赖） ---------- */
function crc32(buf) {
  let c, table = crc32._t;
  if (!table) { table = crc32._t = []; for (let n = 0; n < 256; n++) { c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; table[n] = c >>> 0; } }
  let crc = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ table[(crc ^ buf[i]) & 0xFF];
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
function buildZip(files) {
  const enc = new TextEncoder(); const chunks = []; const central = []; let offset = 0;
  for (const f of files) {
    const nameBytes = enc.encode(f.name); const data = f.data; const crc = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length + data.length); const dv = new DataView(local.buffer);
    dv.setUint32(0, 0x04034b50, true); dv.setUint16(4, 20, true); dv.setUint16(6, 0x0800, true); dv.setUint16(8, 0, true);
    dv.setUint16(10, 0, true); dv.setUint16(12, 0, true); dv.setUint32(14, crc, true);
    dv.setUint32(18, data.length, true); dv.setUint32(22, data.length, true); dv.setUint16(26, nameBytes.length, true); dv.setUint16(28, 0, true);
    local.set(nameBytes, 30); local.set(data, 30 + nameBytes.length); chunks.push(local);
    const cd = new Uint8Array(46 + nameBytes.length); const cdv = new DataView(cd.buffer);
    cdv.setUint32(0, 0x02014b50, true); cdv.setUint16(4, 20, true); cdv.setUint16(6, 20, true); cdv.setUint16(8, 0x0800, true);
    cdv.setUint16(10, 0, true); cdv.setUint16(12, 0, true); cdv.setUint32(16, crc, true); cdv.setUint32(20, data.length, true); cdv.setUint32(24, data.length, true);
    cdv.setUint16(28, nameBytes.length, true); cdv.setUint16(30, 0, true); cdv.setUint16(32, 0, true); cdv.setUint16(34, 0, true); cdv.setUint16(36, 0, true); cdv.setUint32(38, 0, true); cdv.setUint32(42, offset, true);
    cd.set(nameBytes, 46); central.push(cd); offset += local.length;
  }
  const centralSize = central.reduce((a, c) => a + c.length, 0); const centralOffset = offset;
  const end = new Uint8Array(22); const edv = new DataView(end.buffer);
  edv.setUint32(0, 0x06054b50, true); edv.setUint16(4, 0, true); edv.setUint16(6, 0, true); edv.setUint16(8, files.length, true); edv.setUint16(10, files.length, true);
  edv.setUint32(12, centralSize, true); edv.setUint32(16, centralOffset, true); edv.setUint16(20, 0, true);
  const total = chunks.reduce((a, c) => a + c.length, 0) + centralSize + 22; const all = new Uint8Array(total); let pos = 0;
  for (const c of chunks) { all.set(c, pos); pos += c.length; }
  for (const c of central) { all.set(c, pos); pos += c.length; }
  all.set(end, pos);
  return new Blob([all], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}
function mdToDocx(md) {
  const escXml = s => s.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const body = [];
  for (const raw of md.split('\n')) {
    const line = raw.replace(/\s+$/, '');
    if (!line.trim()) { body.push('<w:p/>'); continue; }
    let text = line, bold = false, size = null;
    const h = line.match(/^(#{1,3})\s+(.*)$/);
    if (h) { bold = true; size = h[1].length === 1 ? 18 : h[1].length === 2 ? 15 : 13; text = h[2]; }
    else if (/^工作汇报/.test(line) || /^【/.test(line)) { bold = true; size = 15; }
    const cm = line.match(/^[-*]\s+\[([ x])\]\s+(.*)$/);
    if (cm) text = '• ' + (cm[1] === 'x' ? '✅ ' : '⬜ ') + cm[2];
    else { const m = line.match(/^[-*]\s+(.*)$/); if (m) text = '• ' + m[1]; }
    text = text.replace(/^\[x\]\s*/, '✅ ').replace(/^\[ \]\s*/, '⬜ ').replace(/^>\s?/, '');
    const rpr = bold ? `<w:rPr><w:b/>${size ? `<w:sz w:val="${size * 2}"/>` : ''}</w:rPr>` : '';
    body.push(`<w:p><w:r>${rpr}<w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join('')}<w:sectPr/></w:body></w:document>`;
}
function buildDocx(md) {
  const enc = new TextEncoder();
  const files = [
    { name: '[Content_Types].xml', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>') },
    { name: '_rels/.rels', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>') },
    { name: 'word/document.xml', data: enc.encode(mdToDocx(md)) },
  ];
  return buildZip(files);
}
let lastReport = { md: '', name: '', docxName: '' };

/* ---------- 示例数据 ---------- */
const SAMPLE_RECIPES = [
  { id: 'r1', name: '西兰花炒虾仁', meal: 'lunch', emoji: '🥦', time: 15, tags: ['带饭友好', '高蛋白', '低卡'],
    ing: ['虾仁 200g', '西兰花 1 颗', '蒜末', '盐/黑胡椒'], steps: ['虾仁用盐料酒腌 10 分钟', '西兰花焯水 1 分钟', '热油爆香蒜末，下虾仁炒变色', '倒入西兰花翻炒，调味出锅'] },
  { id: 'r2', name: '番茄鸡蛋面', meal: 'dinner', emoji: '🍜', time: 12, tags: ['快手', '暖胃'],
    ing: ['挂面 1 把', '番茄 2 个', '鸡蛋 2 个', '葱花'], steps: ['番茄切块炒出汁', '加水煮开下面', '淋蛋液成蛋花', '盐调味撒葱花'] },
  { id: 'r3', name: '燕麦莓果碗', meal: 'breakfast', emoji: '🥣', time: 5, tags: ['免煮', '膳食纤维'],
    ing: ['燕麦 50g', '牛奶/酸奶', '蓝莓/草莓', '坚果'], steps: ['燕麦加牛奶浸泡', '铺上莓果与坚果', '可提前一晚冷藏隔夜燕麦'] },
  { id: 'r4', name: '鸡胸藜麦沙拉', meal: 'lunch', emoji: '🥗', time: 18, tags: ['带饭友好', '减脂', '高蛋白'],
    ing: ['鸡胸肉 1 块', '藜麦 60g', '黄瓜/小番茄', '油醋汁'], steps: ['藜麦煮熟沥干', '鸡胸煎熟切块', '蔬菜切丁', '拌入油醋汁'] },
  { id: 'r5', name: '蒸蛋羹', meal: 'breakfast', emoji: '🥚', time: 12, tags: ['低脂', '快手'],
    ing: ['鸡蛋 2 个', '温水 1.5 倍', '生抽/香油'], steps: ['蛋液加温水打匀过筛', '盖保鲜膜扎孔', '中火蒸 10 分钟', '淋生抽香油'] },
  { id: 'r6', name: '蒜香西兰花鸡丁', meal: 'dinner', emoji: '🍗', time: 20, tags: ['下饭', '家常'],
    ing: ['鸡腿肉 1 块', '西兰花', '蒜', '蚝油'], steps: ['鸡肉切丁腌制', '西兰花焯水', '炒香蒜末下鸡丁', '加蚝油与西兰花翻炒'] },
  { id: 'r7', name: '牛油果吐司', meal: 'breakfast', emoji: '🥑', time: 6, tags: ['免煮', '优质脂肪'],
    ing: ['全麦吐司', '牛油果', '鸡蛋(可水煮)', '黑胡椒'], steps: ['吐司烤脆', '牛油果压泥抹上', '铺水煮蛋切片', '撒黑胡椒'] },
  { id: 'r8', name: '韩式辣白菜豆腐汤', meal: 'dinner', emoji: '🍲', time: 18, tags: ['暖身', '开胃'],
    ing: ['嫩豆腐 1 盒', '辣白菜', '五花肉/午餐肉', '大葱'], steps: ['五花肉煸出油', '加辣白菜炒香', '加水与豆腐块煮开', '撒葱花'] },
  { id: 'r9', name: '全麦三明治', meal: 'breakfast', emoji: '🥪', time: 8, tags: ['免煮', '便携'],
    ing: ['全麦面包', '鸡蛋', '生菜', '番茄', '芝士'], steps: ['鸡蛋煮熟切片', '面包铺生菜番茄蛋芝士', '对半切好带走'] },
  { id: 'r10', name: '凉拌鸡丝荞麦面', meal: 'lunch', emoji: '🍝', time: 15, tags: ['带饭友好', '低卡', '清爽'],
    ing: ['荞麦面', '鸡胸丝', '黄瓜丝', '芝麻酱/醋'], steps: ['荞麦面煮熟过凉', '拌入鸡丝黄瓜丝', '调芝麻醋汁拌匀'] },
  { id: 'r11', name: '番茄龙利鱼', meal: 'dinner', emoji: '🐟', time: 20, tags: ['高蛋白', '清淡'],
    ing: ['龙利鱼', '番茄', '豆腐', '金针菇'], steps: ['鱼柳切块腌制', '番茄炒出汁加水', '下鱼块豆腐菇煮熟'] },
  { id: 'r12', name: '香蕉燕麦杯', meal: 'breakfast', emoji: '🍌', time: 4, tags: ['免煮', '膳食纤维'],
    ing: ['即食燕麦', '香蕉', '酸奶', '奇亚籽'], steps: ['杯底铺燕麦', '叠香蕉片与酸奶', '撒奇亚籽即食'] },
];

const SAMPLE_KNOWLEDGE = [
  { id: 'k1', tag: '入门', title: '复利是什么', body: '利息再投资产生“利滚利”。长期定投的核心动力，越早开始优势越大。' },
  { id: 'k2', tag: '工具', title: '基金 vs 股票', body: '基金由经理分散投资一篮子资产，波动小于单只股票，适合没时间盯盘的小白。' },
  { id: 'k3', tag: '风险', title: '资产配置', body: '把资金按风险分散到货币/债券/权益类，目标是在能承受的波动下争取收益。' },
  { id: 'k4', tag: '指标', title: '通胀与购买力', body: '钱放活期会被通胀悄悄稀释，理解 CPI 有助于判断“钱该放哪”。' },
  { id: 'k5', tag: '工具', title: '指数基金(ETF)', body: '跟踪沪深300、标普500等指数，费用低、透明，常被推荐为小白起点。' },
  { id: 'k6', tag: '纪律', title: '定投策略', body: '固定时间固定金额买入，平摊成本、克服追涨杀跌，贵在坚持。' },
  { id: 'k7', tag: '风险', title: '风险承受力', body: '投资前先问：这笔钱多久不用？亏 20% 会不会睡不着？答案决定仓位。' },
  { id: 'k8', tag: '趋势', title: '看懂财经日历', body: '关注央行利率、CPI、非农等数据发布日，市场常在这些节点波动。' },
];

const SAMPLE_FRESH = [
  { id: 'f1', cat: '穿搭', title: '美拉德风穿搭', desc: '棕咖色系叠穿，温暖高级，秋冬通勤也好看。', hot: '🔥 1.2w 讨论', link: 'https://www.douyin.com/search/%E7%BE%8E%E6%8B%89%E5%BE%B7%E9%A3%8E%E7%A9%BF%E6%90%AD' },
  { id: 'f2', cat: '餐厅', title: '城市 B istro 小酒馆', desc: '轻松氛围 + 人均百元的好拍照西餐，适合周末约朋友。', hot: '⭐ 新开', link: 'https://www.bilibili.com/search/all?keyword=%E5%B0%8F%E9%85%92%E9%A6%86' },
  { id: 'f3', cat: '游玩', title: '城市骑行路线', desc: '沿河绿道 + 咖啡店打卡，半天就能充好电。', hot: '', link: 'https://www.bilibili.com/search/all?keyword=%E5%9F%8E%E5%B8%82%E9%AA%91%E8%A1%8C' },
  { id: 'f4', cat: '潮流', title: 'City Walk 城市漫游', desc: '不赶景点，慢慢逛老街与独立书店，最近很火。', hot: '🔥 热门', link: 'https://www.xiaohongshu.com/search_result?keyword=citywalk' },
  { id: 'f5', cat: '理财投资', title: '闲钱自动攒计划', desc: '工资到账自动转一笔到货基，先储蓄后消费。', hot: '', link: 'https://www.bilibili.com/search/all?keyword=%E8%87%AA%E5%8A%A8%E6%94%AF%E5%87%BA' },
  { id: 'f6', cat: '副业', title: '知识付费小课', desc: '把你的编辑/外语专长做成 9.9 小课，边际成本低。', hot: '💡 可尝试', link: 'https://www.bilibili.com/search/all?keyword=%E7%9F%A5%E8%AF%86%E4%BB%98%E8%B4%B9' },
  { id: 'f7', cat: '娱乐', title: '沉浸式解压视频', desc: 'asmr / 整理收纳类视频，睡前放松很解压。', hot: '', link: 'https://www.bilibili.com/search/all?keyword=asmr%E8%A7%A3%E5%8E%8B' },
  { id: 'f8', cat: '潮流', title: '多巴胺穿搭', desc: '高饱和撞色，元气满满，适合拍照出片。', hot: '🔥 热门', link: 'https://www.douyin.com/search/%E5%A4%9A%E5%B7%B4%E8%83%BA%E7%A9%BF%E6%90%AD' },
  { id: 'f9', cat: '副业', title: 'AI 提效接单', desc: '用 AI 做PPT/排版/翻译接小单，外企背景很吃香。', hot: '💡 新思路', link: 'https://www.bilibili.com/search/all?keyword=AI%E6%8E%A5%E5%8D%95' },
  { id: 'f10', cat: '游玩', title: '近郊露营一日', desc: '租装备当天往返，成本可控，周末微度假。', hot: '', link: 'https://www.bilibili.com/search/all?keyword=%E8%BF%91%E9%83%8A%E9%9C%B2%E8%90%A5' },
];

const EN_PLAN = [
  { type: 'listen', icon: '👂', cls: 'listen', title: '听力 · 15min',
    desc: '会议录音精听：先盲听抓大意，再对照脚本跟读，最后听写关键句。',
    steps: ['盲听 1 遍，抓大意', '看脚本跟读 2 遍', '遮住脚本听写关键句', '对照原文修正'],
    material: 'A: Shall we get started? Thanks everyone for joining.\nB: Sure. First, the Q3 editorial calendar came in slightly behind target.\nA: Right — the main delay was the peer-review turnaround. Let’s tighten that.\nB: Agreed. I’ll circle back with the authors by Thursday.' },
  { type: 'speak', icon: '🗣️', cls: 'speak', title: '口语 · 影子跟读',
    desc: '跟读原生句子 3 遍，再自己录/说 1 分钟今日总结。',
    steps: ['逐句跟读 3 遍', '合上稿自己说 1 遍', '挑 1 句用在今日总结'],
    material: '1) Thanks for joining — let’s quickly run through today’s three points.\n2) Sorry, could you say that again a bit slower?\n3) That’s a good point — building on that, I’d suggest we…\n4) Great, I’ll follow up with a short summary by end of day.' },
  { type: 'vocab', icon: '📝', cls: 'vocab', title: '词汇 · 10 词',
    desc: '今天的 10 个外企地道表达，过一遍 + 自测 + 标记已掌握。',
    steps: ['过一遍 10 个表达', '遮盖中文自测', '标记已掌握的'],
    material: '' },
  { type: 'meet', icon: '💼', cls: 'meet', title: '会议 · 模拟开场',
    desc: '读熟开场/总结模板，用你自己的项目练一句，设想一个提问。',
    steps: ['读熟开场模板', '用自己项目练一句', '设想 1 个提问并准备回答'],
    material: '【开场】Hi everyone, thanks for joining. Let’s quickly run through today’s three points.\n【接话】That’s a good point — building on that, I’d suggest we…\n【没听清】Sorry, could you say that again a bit slower?\n【收尾】Great, I’ll circle back with a summary by EOD. Anything before we wrap up?' },
];
// 外企地道表达（按日期轮换取 10 个）
const EN_VOCAB = [
  { id: 'v1', en: 'peer review', cn: '同行评审', ex: 'The paper is under peer review.' },
  { id: 'v2', en: 'turnaround', cn: '处理时长 / 周转', ex: 'We need a faster turnaround.' },
  { id: 'v3', en: 'deadline', cn: '截止日期', ex: 'The deadline is this Friday.' },
  { id: 'v4', en: 'follow up', cn: '跟进', ex: 'I’ll follow up by EOD.' },
  { id: 'v5', en: 'circle back', cn: '稍后回来谈', ex: 'Let’s circle back to this later.' },
  { id: 'v6', en: 'align', cn: '对齐 / 达成一致', ex: 'Let’s align on the scope.' },
  { id: 'v7', en: 'bandwidth', cn: '精力 / 时间余量', ex: 'Do you have bandwidth this week?' },
  { id: 'v8', en: 'touch base', cn: '沟通一下', ex: 'Let’s touch base tomorrow.' },
  { id: 'v9', en: 'action item', cn: '待办事项', ex: 'That’s an action item for me.' },
  { id: 'v10', en: 'stakeholder', cn: '相关方 / 干系人', ex: 'Keep stakeholders informed.' },
  { id: 'v11', en: 'deliverable', cn: '交付物', ex: 'The deliverable is the report.' },
  { id: 'v12', en: 'ramp up', cn: '加快 / 启动', ex: 'We’ll ramp up next month.' },
];
function todayVocab() {
  const base = parseInt(todayStr().replace(/-/g, ''), 10) || 0;
  const start = base % EN_VOCAB.length;
  const out = [];
  for (let i = 0; i < 10; i++) out.push(EN_VOCAB[(start + i) % EN_VOCAB.length]);
  return out;
}
function speak(text) {
  try {
    if (!('speechSynthesis' in window)) { toast('当前浏览器不支持朗读'); return; }
    window.speechSynthesis.cancel();
    const u = new SpeechSynthesisUtterance(text);
    u.lang = 'en-US'; u.rate = 0.95;
    window.speechSynthesis.speak(u);
  } catch (e) { toast('朗读不可用'); }
}

/* ---------- 状态 ---------- */
function seed() {
  return {
    sidebar: [
      { id: 'work', name: '工作', icon: '💼', core: true },
      { id: 'english', name: '英语', icon: '📚', core: true },
      { id: 'finance', name: '理财', icon: '💰', core: true },
      { id: 'recipes', name: '食谱', icon: '🍳', core: true },
      { id: 'fresh', name: '新鲜玩意', icon: '✨', core: true },
    ],
    work: { todos: {} },
    english: { checkins: {}, subs: {}, vocabMastered: [] },
    finance: { knowledge: SAMPLE_KNOWLEDGE.slice(), reads: [] },
    recipes: { favs: [], dailySeed: {} },
    fresh: { items: SAMPLE_FRESH.slice(), favs: [] },
    custom: {},
    team: {
      activeTab: 'goals',
      groups: [
        { id: 'g1', name: '>2年组', target: '', subGroups: [
          { id: 'sg1', name: 'A组(2人)', target: '', memberIds: [] },
          { id: 'sg2', name: 'B组(3人)', target: '', memberIds: [] }
        ]},
        { id: 'g2', name: '1-2年组', target: '', subGroups: [
          { id: 'sg3', name: 'C组(3人)', target: '', memberIds: [] },
          { id: 'sg4', name: 'D组(3人)', target: '', memberIds: [] }
        ]}
      ],
      members: [],
      months: {},          // { '2026-09': { targets: {mid: '125'}, collapsed: false } }
      currentMonth: '',    // 当前查看的月份 'YYYY-MM'，空 = 自动取当月
      weeks: {},           // { '2026-09-14': { start:'2026-09-14', end:'2026-09-20', data:{mid:{completionRate,seriousErrors}} } }
      monthExtras: {},     // { '2026-09': { mid: { 月度加分/扣分项 } } }
      _selectedPeriod: ''  // 当前查看的日期区间 key
    },
    settings: { sync: { mode: 'local', url: '', enabled: false, cloudId: '', cloudUrl: '', cloudKey: '' } },
  };
}
function load() {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return seed();
    const s = JSON.parse(raw);
    // 简单补字段
    const base = seed();
    return Object.assign(base, s, {
      work: Object.assign(base.work, s.work),
      english: Object.assign(base.english, s.english),
      finance: Object.assign(base.finance, s.finance),
      recipes: Object.assign(base.recipes, s.recipes),
      fresh: Object.assign(base.fresh, s.fresh),
      team: Object.assign(base.team, s.team || {}),
      settings: Object.assign(base.settings, s.settings || {}, {
        sync: Object.assign(base.settings.sync, (s.settings && s.settings.sync) || {}),
      }),
    });
  } catch (e) { return seed(); }
}
let state = load();
// 老数据迁移：把成员身上的 personalTarget 归入 2026-09，并初始化当前月份
migrateTeamMonths();
normalizeWeeks();   // 旧的 'YYYY-Www' 周记录补上 start / end
if (!state.team.currentMonth) state.team.currentMonth = monthKey();
if (!state.team.monthExtras) state.team.monthExtras = {};
if (!state.team._selectedPeriod) state.team._selectedPeriod = periodList()[0] || '';
let _pushTimer = null, _lastPush = 0, _syncing = false;
function save(silent) {
  localStorage.setItem(KEY, JSON.stringify(state));
  if (!silent && syncActive()) schedulePush();
}

/* ---------- 顶部日期 / 问候 ---------- */
function greet() {
  const h = new Date().getHours();
  if (h < 6) return '夜深了'; if (h < 11) return '早上好'; if (h < 14) return '中午好';
  if (h < 18) return '下午好'; return '晚上好';
}
/* ---------- 财务月（每期 15 日 至 次月 14 日，按公司实际排期微调） ---------- */
const FIN_PERIODS_2026 = [
  ['2026-08-14', '2026-09-14'],
  ['2026-09-15', '2026-10-14'],
  ['2026-10-15', '2026-11-13'],
  ['2026-11-14', '2026-12-15'],
];
const fmtFin = d => `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}（${WEEK[d.getDay()]}）`;
function finPeriodStr(d) {
  const t = todayStr(d);
  for (const [s, e] of FIN_PERIODS_2026) {
    if (t >= s && t <= e) { const [a, b] = [new Date(s + 'T00:00:00'), new Date(e + 'T00:00:00')]; return `${fmtFin(a)} – ${fmtFin(b)}`; }
  }
  // 列表之外的日期，按通用规则推算：15日 至 次月14日
  const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  let start, end;
  if (dt.getDate() >= 15) { start = new Date(dt.getFullYear(), dt.getMonth(), 15); end = new Date(dt.getFullYear(), dt.getMonth() + 1, 14); }
  else { start = new Date(dt.getFullYear(), dt.getMonth() - 1, 15); end = new Date(dt.getFullYear(), dt.getMonth(), 14); }
  return `${fmtFin(start)} – ${fmtFin(end)}`;
}
function renderTopbar(title) {
  $('#pageTitle').textContent = title;
  const d = new Date();
  $('#pageDate').innerHTML = `${greet()} · ${todayStr(d)} ${WEEK[d.getDay()]}` +
    `<br><span style="font-size:11.5px;color:var(--ink-soft)">📅 当前财务月：${finPeriodStr(d)}</span>`;
}

/* ---------- 导航栏 ---------- */
let current = 'work';
function renderSidebar() {
  const nav = $('#nav'); nav.innerHTML = '';
  state.sidebar.forEach(it => {
    const a = document.createElement('div');
    a.className = 'nav-item' + (it.id === current ? ' active' : '');
    a.dataset.id = it.id;
    a.innerHTML = `<span class="ni-icon">${it.icon}</span><span class="ni-name">${esc(it.name)}</span>` +
      (state.sidebar.length > 1 ? `<span class="ni-remove" title="移除该板块" data-remove="${it.id}">✕</span>` : '');
    nav.appendChild(a);
  });
}
function go(id) {
  const it = state.sidebar.find(s => s.id === id);
  if (!it) return;
  current = id;
  renderSidebar();
  renderView();
  $('#sidebar').classList.remove('open'); $('#scrim').classList.remove('show');
  if (location.hash !== '#' + id) history.replaceState(null, '', '#' + id);
}

/* ---------- 视图分发 ---------- */
function renderView() {
  const it = state.sidebar.find(s => s.id === current) || state.sidebar[0];
  renderTopbar(it.name);
  const v = $('#view');
  if (it.core) {
    ({ work: viewWork, english: viewEnglish, finance: viewFinance, recipes: viewRecipes, fresh: viewFresh }[it.id] || viewWork)(v);
  } else {
    viewCustom(v, it);
  }
  updateStreakPill();
}

/* =================== 工作 =================== */
const CATS = ['稿件处理', '项目对接', '会议沟通', '人事招聘', '日常管理', '学习培训', '临时交办', '其他'];
const SCOPES = ['组内', '科室', '部门', '外部'];
let todoDraftDate = ''; // 添加框上次选的日期，便于连续把任务录到同一天
const OLD_CAT_MAP = {
  '组内': { cat: '其他', scope: '组内' },
  '科室': { cat: '其他', scope: '科室' },
  '部门': { cat: '其他', scope: '部门' },
  '面试': { cat: '人事招聘', scope: '组内' },
  '项目对接': { cat: '项目对接', scope: '组内' },
  '日常管理': { cat: '日常管理', scope: '组内' },
  '特殊任务': { cat: '临时交办', scope: '组内' },
  '其他': { cat: '其他', scope: '组内' },
};
function migrateTodos() {
  if (state.work.catSchema === 2) return;
  Object.keys(state.work.todos || {}).forEach(d => {
    (state.work.todos[d] || []).forEach(t => {
      if (t.scope && CATS.indexOf(t.cat) >= 0) return;
      const m = OLD_CAT_MAP[t.cat];
      if (m) { t.cat = m.cat; t.scope = m.scope; }
      else { t.scope = t.scope || SCOPES[0]; }
    });
  });
  state.work.catSchema = 2;
  save();
}
function ensureToday() {
  migrateTodos();
  const t = todayStr();
  if (!state.work.todos[t]) { state.work.todos[t] = []; save(); }
}
function workStats() {
  const now = new Date();
  const t = todayStr();
  const today = state.work.todos[t] || [];
  const dow = (now.getDay() + 6) % 7;
  const weekStart = addDays(startOfDay(now), -dow);
  let wTotal = 0, wDone = 0, mTotal = 0, mDone = 0;
  Object.keys(state.work.todos).forEach(d => {
    if (d > t) return; // 提前排的未来任务不计入已完成统计
    const dd = new Date(d);
    if (dd >= weekStart) { const arr = state.work.todos[d]; wTotal += arr.length; wDone += arr.filter(x => x.done).length; }
    if (dd.getFullYear() === now.getFullYear() && dd.getMonth() === now.getMonth()) {
      const arr = state.work.todos[d]; mTotal += arr.length; mDone += arr.filter(x => x.done).length;
    }
  });
  return { todayTotal: today.length, todayDone: today.filter(x => x.done).length, wTotal, wDone, mTotal, mDone };
}
function viewWork(v) {
  const workTab = state.work._tab || 'todo';
  v.innerHTML = `
  <div class="work-tabs" style="margin-bottom:16px">
    <button class="work-tab ${workTab==='todo'?'on':''}" data-act="work-tab" data-tab="todo">📝 To Do</button>
    <button class="work-tab ${workTab==='team'?'on':''}" data-act="work-tab" data-tab="team">👥 小组目标管理</button>
  </div>
  <div id="workPanel"></div>`;
  const panel = $('#workPanel');
  if (workTab === 'todo') renderWorkTodo(panel);
  else viewTeamGoals(panel);
}
function renderWorkTodo(v) {
  ensureToday();
  const t = todayStr();
  const today = state.work.todos[t] || [];
  const other = Object.keys(state.work.todos).filter(d => d !== t);
  const future = other.filter(d => d > t).sort();
  const past = other.filter(d => d < t).sort().reverse();
  const defDate = todoDraftDate || t;
  const st = workStats();
  v.innerHTML = `
  <div class="card card-soft">
    <div class="card-head"><span class="ch-emoji">📊</span><h2>效率概览</h2><span class="ch-sub">坚持记录，季度 / 年度总结更轻松</span></div>
    <div class="stats-row">
      <div class="mini-stat"><div class="n">${st.todayDone}/${st.todayTotal}</div><div class="l">今日完成</div></div>
      <div class="mini-stat"><div class="n">${st.wDone}/${st.wTotal}</div><div class="l">本周完成</div></div>
      <div class="mini-stat"><div class="n">${st.mDone}/${st.mTotal}</div><div class="l">本月完成</div></div>
      <div class="mini-stat"><div class="n">${st.wTotal ? Math.round(st.wDone / st.wTotal * 100) : 0}%</div><div class="l">本周完成率</div></div>
    </div>
  </div>
  <div class="card">
    <div class="card-head"><span class="ch-emoji">📝</span><h2>今日 To Do</h2>
      <span class="ch-sub">${t} ${WEEK[new Date().getDay()]}</span></div>
    <div class="todo-add">
      <input class="field" id="todoInput" placeholder="写点要做的事…（可提前排到未来日期）" />
      <span class="ch-sub">日期</span>
      <input type="date" class="field" id="todoDate" value="${defDate}" style="width:150px" />
      <span class="ch-sub">性质</span>
      <div class="seg" id="catSeg">
        ${CATS.map((c, i) => `<button data-cat="${c}" class="cat-${c} ${i === 0 ? 'on' : ''}">${c}</button>`).join('')}
      </div>
      <span class="ch-sub">范围</span>
      <div class="seg" id="scopeSeg">
        ${SCOPES.map((c, i) => `<button data-scope="${c}" class="scope-${c} ${i === 0 ? 'on' : ''}">${c}</button>`).join('')}
      </div>
      <button class="btn primary" data-act="todo-add">＋ 添加</button>
    </div>
    <div class="todo-group" id="todayGroup" style="margin-top:14px">
      ${today.length ? today.map(todoRow).join('') : '<div class="empty">还没有安排，先写一条吧～</div>'}
    </div>
    <div style="margin-top:14px"><button class="btn yellow sm" data-act="report-open">📊 生成半月/月/季/年报</button></div>
  </div>

  ${future.length ? `<div class="card card-soft"><div class="card-head"><span class="ch-emoji">🗓️</span><h2>未来安排</h2><span class="ch-sub">提前计划的待办，到期当天自动进入「今日 To Do」</span></div>
    ${future.map(d => historyDay(d)).join('')}
  </div>` : ''}

  ${past.length ? `<div class="card card-soft"><div class="card-head"><span class="ch-emoji">🗂️</span><h2>历史记录</h2><span class="ch-sub">点击日期可折叠</span></div>
    ${past.map(d => historyDay(d)).join('')}
  </div>` : ''}
  `;
}
function todoRow(tk) {
  const catPill = `<span class="pill cat-${esc(tk.cat || '其他')}">${esc(tk.cat || '其他')}</span>`;
  const scopePill = `<span class="pill scope-${esc(tk.scope || '组内')}">${esc(tk.scope || '组内')}</span>`;
  return `<div class="todo-item ${tk.done ? 'done' : ''}" data-id="${tk.id}">
    <div class="todo-check" data-act="todo-toggle" data-id="${tk.id}" data-date="${tk.date}">${tk.done ? '✓' : ''}</div>
    <div class="todo-text">${esc(tk.text)}</div>
    <div class="todo-meta">${catPill}${scopePill}</div>
    <div class="todo-edit" data-act="todo-edit" data-id="${tk.id}" data-date="${tk.date}" title="修改">✏️</div>
    <div class="todo-del" data-act="todo-del" data-id="${tk.id}" data-date="${tk.date}">🗑</div>
  </div>`;
}

/* =================== 小组目标管理 =================== */
// 小小组 → 马卡龙配色的全局映射（成员管理 / 周数据录入共用同一套颜色）
const SG_COLOR_KEYS = ['a', 'b', 'c', 'd'];
function sgColorMap() {
  const map = {};
  let i = 0;
  (state.team.groups || []).forEach(g => (g.subGroups || []).forEach(sg => {
    map[sg.id] = 'tm-sg-' + SG_COLOR_KEYS[i % SG_COLOR_KEYS.length];
    i++;
  }));
  return map;
}
// 扁平化的小小组列表，保持「大组 → 小小组」顺序
function allSubGroupList() {
  return (state.team.groups || []).reduce((acc, g) => acc.concat((g.subGroups || []).map(sg => ({ g, sg }))), []);
}

/* ---- 加减分规则（唯一数据源：规则页展示 + 自动算分都读它） ---- */
const TEAM_RULES = {
  plus: [
    { period: '每周', cat: '业务数据', items: [
      { key: 'bonusPersonalGoal', name: '个人完成率目标', pts: 1, auto: '完成率 ≥ 个人目标' },
      { key: 'bonusSubGroupGoal', name: '小小组完成率达标（小组）', pts: 1, auto: '小小组全员达标' },
      { key: 'bonusNoErrorWeek', name: '个人本周无严错产生', pts: 1, auto: '严错数 = 0' }
    ]},
    { period: '每月', cat: '业务数据', items: [
      { key: 'monthAttendTop1', name: '应出勤组内 Top 1', pts: 3 },
      { key: 'monthAttendTop2', name: '应出勤组内 Top 2', pts: 2 },
      { key: 'monthAttendTop3', name: '应出勤组内 Top 3', pts: 1 },
      { key: 'bonusNoErrorMonth', name: '无严错', pts: 2 }
    ]},
    { period: '每月', cat: '小组活动', items: [
      { key: 'bonusActivity', name: '组织技能分享 / 座谈会等组内线下活动', pts: 3 },
      { key: 'bonusInitiative', name: '主动接收临时紧急任务 / 主动补位 / 主动提建设性建议', pts: 2 },
      { key: 'bonusOnlineShare', name: '线上 case 分享 / tips 分享 / 疑难问题攻坚…', pts: 1 },
      { key: 'bonusEventOwner', name: '“周内大事件”负责人', pts: 2 },
      { key: 'bonusTea', name: '组织下午茶', pts: 1 }
    ]},
    { period: '每月', cat: '组会', items: [
      { key: 'bonusQuiz', name: '小组答题（小组）', pts: 1 }
    ]},
    { period: '每月', cat: '科室', items: [
      { key: 'bonusTrainer', name: '科室培训讲师', pts: 3 },
      { key: 'bonusDeptOther', name: '科室其他活动', pts: '1~3' }
    ]}
  ],
  minus: [
    { period: '每周', cat: '个人', items: [
      { key: 'deductSeriousErr', name: '每产生 1 个严错', pts: 0.5, auto: '严错数 × 0.5' }
    ]},
    { period: '每月', cat: '个人', items: [
      { key: 'deductErrorsGte3', name: '严错 ≥ 3 个', pts: 2, auto: '严错数 ≥ 3' },
      { key: 'deductQualityRule', name: '触及当月质量目标具体条例（如 compare、图表解释问题等）', pts: 1 },
      { key: 'deductDragGroup', name: '完成率低于小组目标且导致小组不达标（部分成员达标）', pts: 2 },
      { key: 'deductPending', name: '违规 pending / 返稿 / 跳 QC…', pts: 0.5 },
      { key: 'deductLow', name: 'Low 等级违规（未送美修 / 稿件 delay 48h / 宣传文件 / 多次违规修改稿件状态导致严重后果等）', pts: 1 },
      { key: 'deductMed', name: 'Medium 等级违规', pts: 2 },
      { key: 'deductHigh', name: 'High 等级违规', pts: 3 }
    ]},
    { period: '每月', cat: '小小组', items: [
      { key: 'deductGroupAllFail', name: '小小组成员完成率均不达标', pts: 2, auto: '小小组全员不达标' }
    ]}
  ]
};

// 统一重渲染入口：始终渲染到最外层 #workPanel，避免把 team 视图套进 #teamPanel 造成层级错乱
function renderTeam() {
  const panel = $('#workPanel');
  if (panel) viewTeamGoals(panel);
  else viewWork($('#view'));
}
function viewTeamGoals(v) {
  const tab = state.team.activeTab || 'goals';
  v.innerHTML = `
  <div class="team-tabs" style="margin-bottom:16px">
    <button class="team-tab ${tab==='goals'?'on':''}" data-act="team-tab" data-ttab="goals">🎯 目标设置</button>
    <button class="team-tab ${tab==='members'?'on':''}" data-act="team-tab" data-ttab="members">👥 成员管理</button>
    <button class="team-tab ${tab==='rules'?'on':''}" data-act="team-tab" data-ttab="rules">📐 加减分规则</button>
    <button class="team-tab ${tab==='weekly'?'on':''}" data-act="team-tab" data-ttab="weekly">📝 周数据录入</button>
    <button class="team-tab ${tab==='scoreboard'?'on':''}" data-act="team-tab" data-ttab="scoreboard">🏆 积分看板</button>
  </div>
  <div id="teamPanel"></div>`;
  const p = $('#teamPanel');
  if (tab === 'goals') renderTeamGoals(p);
  else if (tab === 'members') renderTeamMembers(p);
  else if (tab === 'rules') renderTeamRules(p);
  else if (tab === 'weekly') renderTeamWeekly(p);
  else renderTeamScoreboard(p);
}

/* ---- 加减分规则分栏 ---- */
function renderTeamRules(v) {
  const ruleTable = (side) => {
    const isPlus = side === 'plus';
    const groups = TEAM_RULES[side];
    // 按「每周 / 每月」分段，保留各段的分类小标题
    const periods = [];
    groups.forEach(gr => {
      let p = periods.find(x => x.period === gr.period);
      if (!p) { p = { period: gr.period, cats: [] }; periods.push(p); }
      p.cats.push(gr);
    });
    const body = periods.map(pd => `
      <tr class="tr-period"><td colspan="2">${pd.period}</td></tr>
      ${pd.cats.map(c => `
        <tr class="tr-cat"><td colspan="2">${esc(c.cat)}</td></tr>
        ${c.items.map(it => `<tr>
          <td class="tr-name">${esc(it.name)}${it.auto ? `<span class="tr-auto">自动：${esc(it.auto)}</span>` : ''}</td>
          <td class="tr-pts ${isPlus ? 'pts-plus' : 'pts-minus'}">${isPlus ? '+' : '-'}${it.pts}</td>
        </tr>`).join('')}
      `).join('')}
    `).join('');
    return `<div class="rule-col ${isPlus ? 'rule-col-plus' : 'rule-col-minus'}">
      <div class="rule-head ${isPlus ? 'rh-plus' : 'rh-minus'}">${isPlus ? '加分' : '扣分'}</div>
      <table class="rule-table"><tbody>${body}</tbody></table>
    </div>`;
  };

  v.innerHTML = `<div class="card">
    <div class="card-head"><h2>📐 加减分规则</h2><span class="ch-sub">周数据录入的得分按此规则自动计算</span></div>
    <div class="rule-grid">${ruleTable('plus')}${ruleTable('minus')}</div>
    <div class="rule-note">
      <b>自动计分说明：</b>「周数据录入」里只需填写 <b>完成率</b> 与 <b>严错数</b>，
      系统会据此自动判定 <b>个人完成率目标</b>（完成率 ≥ 个人目标 +1）、
      <b>小小组完成率达标</b>（小小组全员达标 +1）、
      <b>无严错</b>（严错数 = 0 时 +1，否则每个严错 -0.5），并即时给出本周得分。<br>
      标注「每月」的项按月统计，在成员行内展开「更多」后填写；每月积分 Top 3 有额外惊喜～
    </div>
  </div>`;
}

/* ---- 目标设置 ---- */
function renderTeamGoals(v) {
  const T = state.team;
  v.innerHTML = `<div class="card"><div class="card-head"><h2>🎯 小组目标设置</h2><span class="ch-sub">填写各层级目标完成率（%）</span></div>
  ${T.groups.map(g => `
    <div class="team-group-card" style="margin-bottom:20px">
      <div class="team-g-name">${esc(g.name)} <span class="team-g-target-label">大组目标</span>
        <input type="text" class="field team-target-input" data-gid="${g.id}" data-level="group" value="${esc(g.target||'')}" placeholder="%" style="width:80px"> %
      </div>
      <div class="team-subgroups">
        ${g.subGroups.map(sg => `
          <div class="team-sg-row">
            <span class="team-sg-name">${esc(sg.name)}</span>
            <span class="team-sg-target-label">小小组目标</span>
            <input type="text" class="field team-target-input" data-gid="${g.id}" data-sgid="${sg.id}" data-level="subgroup" value="${esc(sg.target||'')}" placeholder="%" style="width:80px"> %
            <span class="team-sg-members">
              ${(T.members.filter(m => sg.memberIds.includes(m.id))).map(m => `<span class="pill cat-其他">${esc(m.name)}<button data-act="team-rm-member-sg" data-mid="${m.id}" data-sgid="${sg.id}" title="移出此小组" style="border:none;background:none;font-size:11px;margin-left:2px;cursor:pointer">✕</button>`).join('')}
              ${sg.memberIds.length === 0 ? '<span style="color:var(--ink-faint);font-size:12px">暂无成员，请在「成员管理」中分配</span>' : ''}
            </span>
          </div>
        `).join('')}
      </div>
    </div>
  `).join('')}
  <div style="margin-top:12px"><button class="btn primary" data-act="team-save-goals">💾 保存目标</button></div>
  </div>`;
}

/* ---- 月份工具（成员目标按月存档） ---- */
// 切换查看的月份：离开的月份自动收起
function switchMonth(mk) {
  const T = state.team;
  const prev = curMonth();
  if (prev && prev !== mk) { const po = ensureMonth(prev); po.collapsed = true; }
  ensureMonth(mk).collapsed = false;
  T.currentMonth = mk;
  save(); renderTeam();
}
function monthKey(d) { d = d || new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); }
function monthLabel(k) { if (!k) return ''; const [y, m] = k.split('-'); return `${y}年${parseInt(m, 10)}月`; }
function curMonth() { return state.team.currentMonth || monthKey(); }
function ensureMonth(k) {
  const T = state.team;
  if (!T.months) T.months = {};
  if (!T.months[k]) T.months[k] = { targets: {}, collapsed: false };
  if (!T.months[k].targets) T.months[k].targets = {};
  return T.months[k];
}
// 该月某成员的目标值
function mTarget(k, mid) {
  const mo = (state.team.months || {})[k];
  if (mo && mo.targets && mo.targets[mid] != null && mo.targets[mid] !== '') return mo.targets[mid];
  // 兼容旧数据：仅当月回退到成员身上的 personalTarget
  if (k === monthKey()) { const m = state.team.members.find(x => x.id === mid); if (m && m.personalTarget) return m.personalTarget; }
  return '';
}
// 把老结构（成员身上的 personalTarget）一次性迁移到 2026-09
function migrateTeamMonths() {  const T = state.team;
  if (!T.months) T.months = {};
  const legacy = T.members.filter(m => m.personalTarget);
  if (!legacy.length) return false;
  const k = '2026-09';
  const mo = ensureMonth(k);
  let n = 0;
  legacy.forEach(m => { if (mo.targets[m.id] == null || mo.targets[m.id] === '') { mo.targets[m.id] = m.personalTarget; n++; } });
  if (!T.currentMonth) T.currentMonth = monthKey();
  return n > 0;
}

function renderTeamMembers(v) {
  const T = state.team;
  const showAdd = T._showAddForm;
  const activeM = curMonth();
  const thisM = monthKey();
  const addForm = showAdd ? `<div class="team-add-form" style="background:var(--green-50);border-radius:10px;padding:12px;margin-bottom:12px">
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <label class="tm-lb">姓名</label>
        <input class="field" id="newMName" placeholder="成员姓名" style="flex:1;min-width:120px">
        <label class="tm-lb">个人目标%</label>
        <input type="text" class="field" id="newMTarget" placeholder="如 125" style="width:80px">
        <label class="tm-lb">大组</label>
        <select class="field" id="newMGroup" style="width:130px">
          ${T.groups.map(g => `<option value="${g.id}">${esc(g.name)}</option>`).join('')}
        </select>
        <label class="tm-lb">小小组</label>
        <select class="field" id="newMSGroup" style="width:130px">
          ${((T.groups[0] || {}).subGroups || []).map(sg => `<option value="${sg.id}">${esc(sg.name)}</option>`).join('')}
        </select>
        <button class="btn primary sm" data-act="team-add-member-do-inline">✅ 确认添加</button>
      </div>
      <div style="margin-top:8px;font-size:12px;color:var(--ink-soft)">目标将计入当前查看月份：<b>${monthLabel(activeM)}</b></div>
    </div>` : '';

  const SG_COLORS = ['a', 'b', 'c', 'd']; // 马卡龙配色，A/B/C/D 依次
  const allSubGroups = T.groups.reduce((acc, g) => acc.concat((g.subGroups || []).map(sg => ({ g, sg }))), []);

  // 通用：渲染某个月份的「大组 → 小小组」表格
  const renderMonthTables = (mk) => {
    let si = 0;
    return T.groups.map(g => {
      const rows = [];
      (g.subGroups || []).forEach(sg => {
        const colorKey = 'tm-sg-' + SG_COLORS[si % SG_COLORS.length];
        si++;
        const mems = T.members.filter(m => m.subGroupId === sg.id || (m.groupId === g.id && m.subGroupId === sg.id));
        mems.forEach((m, i) => {
          const val = mTarget(mk, m.id);
          rows.push(`<tr>
            <td class="tm-td-grp ${colorKey}">${i === 0 ? esc(sg.name) : ''}</td>
            <td class="tm-td-name ${colorKey}">${esc(m.name)}</td>
            <td class="tm-td-target ${colorKey}"><input type="text" class="field tm-target-cell" data-mid="${m.id}" data-mk="${mk}" value="${esc(val)}" placeholder="—" style="width:78px;text-align:center"></td>
            <td class="tm-td-sgtarget ${colorKey}">${i === 0 ? (sg.target ? esc(sg.target) + '%' : '<span style="color:var(--ink-faint)">—</span>') : ''}</td>
            <td class="tm-td-act ${colorKey}">${mk === activeM ? `<button class="btn sm ghost" data-act="team-del-member" data-mid="${m.id}" style="color:var(--danger);padding:2px 6px">✕</button>` : ''}</td>
          </tr>`);
        });
        if (!mems.length) {
          rows.push(`<tr><td class="tm-td-grp ${colorKey}">${esc(sg.name)}</td><td colspan="3" class="${colorKey}" style="color:var(--ink-faint);font-size:12.5px">暂无成员</td><td class="tm-td-act ${colorKey}"></td></tr>`);
        }
      });
      const gT = T.members.filter(m => m.groupId === g.id).map(m => parseFloat(mTarget(mk, m.id))).filter(n => !isNaN(n));
      const avg = gT.length ? (gT.reduce((a, b) => a + b, 0) / gT.length).toFixed(2) : '';
      return `<div class="tm-group-block">
        <div class="tm-group-title">${esc(g.name)}</div>
        <table class="tm-table">
          <thead><tr><th style="width:90px">组别</th><th>Name</th><th style="width:90px">个人目标</th><th style="width:110px">小小组完成率目标</th><th style="width:40px"></th></tr></thead>
          <tbody>${rows.join('')}</tbody>
          <tfoot><tr class="tm-tfoot"><td colspan="3">入职年限总目标</td><td>${avg ? avg + '%' : '—'}</td><td></td></tr></tfoot>
        </table>
      </div>`;
    }).join('');
  };

  const legend = `<div class="tm-legend">
    ${allSubGroups.map((o, i) => `<span class="tm-lg"><i class="tm-dot tm-dot-${SG_COLORS[i % SG_COLORS.length]}"></i>${esc(o.sg.name)}</span>`).join('')}
  </div>`;

  // 月份下拉的全部候选（本月 + 当前查看月 + 已存档月份 + 未来若干月）
  const allMonthKeys = (act) => {
    const s = new Set([thisM, act, ...Object.keys(T.months || {})]);
    for (let i = 1; i <= 3; i++) s.add(monthKey(new Date(new Date().getFullYear(), new Date().getMonth() + i, 1)));
    return [...s].sort().reverse();
  };

  // 其他月份（除当前查看月）→ 折叠卡片；默认收起（用户手动展开的状态在 state 里保留）
  const histKeys = Object.keys(T.months || {}).filter(k => k !== activeM).sort().reverse();
  const histBlocks = histKeys.map(k => {
    const mo = T.months[k] || { targets: {}, collapsed: true };
    if (mo.collapsed === undefined) mo.collapsed = true;
    const open = !mo.collapsed;
    const filled = Object.keys(mo.targets || {}).filter(id => mo.targets[id] !== '' && mo.targets[id] != null).length;
    const isFuture = k > thisM;
    return `<div class="tm-month-fold ${open ? 'open' : ''}">
      <div class="tm-month-fold-head">
        <span class="tm-fold-hit" data-act="team-month-toggle" data-mk="${k}">
          <span class="tm-fold-arrow">${open ? '▾' : '▸'}</span>
          <span class="tm-fold-name">${monthLabel(k)}${isFuture ? ' <span class="tm-fold-tag">未来</span>' : ''}</span>
        </span>
        <span class="tm-fold-meta">${filled} 人已填</span>
        <button class="btn xs ghost" data-act="team-month-del" data-mk="${k}" title="删除该月记录">🗑</button>
      </div>
      ${open ? `<div class="tm-month-fold-body">${renderMonthTables(k)}
        <div style="margin-top:10px"><button class="btn primary sm" data-act="team-save-member-targets" data-mk="${k}">💾 保存 ${monthLabel(k)} 目标</button></div>
      </div>` : ''}
    </div>`;
  }).join('');

  v.innerHTML = `<div class="card"><div class="card-head"><h2>👥 成员管理（共${T.members.length}人）</h2><span class="ch-sub">目标按月存档，可直接在表格里填写</span></div>
  <div style="margin-bottom:12px;display:flex;gap:8px;flex-wrap:wrap;align-items:center">
    <button class="btn primary sm" data-act="team-add-member">${showAdd ? '－ 取消添加' : '＋ 添加成员'}</button>
    <label class="tm-lb" style="margin-left:4px">月份</label>
    <select class="field" id="tmMonthSel" style="width:150px">
      ${allMonthKeys(activeM).map(k => `<option value="${k}" ${k === activeM ? 'selected' : ''}>${monthLabel(k)}${k === thisM ? '（本月）' : ''}</option>`).join('')}
    </select>
    <button class="btn sm" data-act="team-month-new">＋ 新建月份</button>
    ${activeM !== thisM ? `<button class="btn sm ghost" data-act="team-month-goto" data-mk="${thisM}">回到本月</button>` : ''}
  </div>
  ${addForm}
  ${legend}
  ${renderMonthTables(activeM) || '<div class="empty">暂无分组，请先在「目标设置」中配置</div>'}
  <div style="margin-top:12px"><button class="btn primary" data-act="team-save-member-targets" data-mk="${activeM}">💾 保存 ${monthLabel(activeM)} 目标</button></div>

  ${histKeys.length ? `<div class="tm-hist-title">📁 其他月份（点击标题展开 / 收起，可查看也可修改）</div>
  ${histBlocks}` : ''}
  </div>`;

  // 月份下拉切换
  const mSel = $('#tmMonthSel');
  if (mSel) mSel.addEventListener('change', () => {
    switchMonth(mSel.value);
  });

  // 大组联动小小组（添加表单）
  if (showAdd) {
    const gSel = $('#newMGroup'), sgSel = $('#newMSGroup');
    if (gSel && sgSel) {
      gSel.addEventListener('change', () => {
        const g = T.groups.find(x => x.id === gSel.value);
        sgSel.innerHTML = ((g ? g.subGroups : []) || []).map(sg => `<option value="${sg.id}">${esc(sg.name)}</option>`).join('');
      });
    }
  }
}

/* ---- 新建月份弹窗 ---- */
function openMonthPicker() {
  const now = new Date();
  const opts = [];
  for (let i = -6; i <= 6; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const k = monthKey(d);
    opts.push(`<option value="${k}">${monthLabel(k)}${k === monthKey() ? '（本月）' : ''}</option>`);
  }
  openModal(`<h3>新建 / 切换月份</h3>
    <div class="row"><label>选择月份</label>
      <select class="field" id="mpMonth" style="width:100%">${opts.join('')}</select>
    </div>
    <div style="font-size:12px;color:var(--ink-soft);margin-top:6px">新建后该月目标为空表，可重新填写；已有月份会直接切换过去，不会清空数据。</div>
    <div class="row" style="margin-top:14px"><button class="btn primary" data-act="team-month-new-do">确定</button></div>`);
}

/* ---- 周数据录入（日期区间 + 极简录入 + 按规则自动算分） ---- */
// 周记录 key = 起始日期 'YYYY-MM-DD'
function getWeekKey(dateStr) { const d = new Date(dateStr); const jan1 = new Date(d.getFullYear(), 0, 1); const days = Math.floor((d - jan1) / 86400000); const w = Math.ceil((days + jan1.getDay() + 1) / 7); return `${d.getFullYear()}-W${String(w).padStart(2, '0')}`; }
function periodList() { return Object.keys(state.team.weeks || {}).sort().reverse(); }
function weekLabel(w) { return w && w.start ? `${w.start} ~ ${w.end || w.start}` : ''; }
function monthOfPeriod(w) { return w && w.start ? w.start.slice(0, 7) : ''; }
function periodsInMonth(mk) { return periodList().filter(k => monthOfPeriod(state.team.weeks[k]) === mk).sort(); }
function mondayStr(d) { const x = d ? new Date(d) : new Date(); const dow = (x.getDay() + 6) % 7; x.setDate(x.getDate() - dow); return dateStrOf(x); }
function sundayStr(d) { const x = new Date(mondayStr(d)); x.setDate(x.getDate() + 6); return dateStrOf(x); }
function dateStrOf(x) { return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`; }
function shiftDate(s, n) { const d = new Date(s); d.setDate(d.getDate() + n); return dateStrOf(d); }
// 兼容旧数据：把 '2026-W38' 形式的记录补上 start / end
function normalizeWeeks() {
  const T = state.team; if (!T.weeks) T.weeks = {};
  let changed = false;
  Object.keys(T.weeks).forEach(k => {
    const w = T.weeks[k]; if (!w) return;
    if (!w.start) { w.start = w.weekDate || todayStr(); changed = true; }
    if (!w.end) { w.end = shiftDate(w.start, 6); changed = true; }
    if (!w.data) { w.data = {}; changed = true; }
  });
  return changed;
}

// 某月是否已设置过任何成员目标
function monthHasTargets(mk) {
  const t = ((state.team.months || {})[mk] || {}).targets || {};
  return Object.keys(t).some(id => t[id] !== '' && t[id] != null);
}
// 个人目标「就近沿用」：优先取该月；该月没填则沿用最近一次设置过的目标
// 返回 { val, from, carried }：val=取到的值，from=实际来源月份，carried=是否为沿用值
function mTargetResolved(mk, mid) {
  const T = state.team;
  const own = mTarget(mk, mid);
  if (own !== '') return { val: own, from: mk, carried: false };
  const keys = Object.keys(T.months || {}).filter(x => x !== mk && monthHasTargets(x)).sort();
  const back = keys.filter(x => x < mk).reverse();   // 先往前（上月 / 更早）
  for (const x of back) { const v = mTarget(x, mid); if (v !== '') return { val: v, from: x, carried: true }; }
  const fwd = keys.filter(x => x > mk);              // 再往后（预先填好的下月目标）
  for (const x of fwd) { const v = mTarget(x, mid); if (v !== '') return { val: v, from: x, carried: true }; }
  const m = T.members.find(x => x.id === mid);
  if (m && m.personalTarget) return { val: m.personalTarget, from: '', carried: true };
  return { val: '', from: '', carried: false };
}
// 该成员在指定月份的个人目标（数值，未设置为 null）
function memberTarget(mid, mk) {
  const r = mTargetResolved(mk || curMonth(), mid);
  const n = parseFloat(r.val);
  return isNaN(n) ? null : n;
}
// 小小组达标判定：全员达标 meet / 全员不达标 allFail
function subGroupMeet(sgId, mk, rateMap) {
  const ms = state.team.members.filter(m => m.subGroupId === sgId);
  if (!ms.length) return { meet: false, allFail: false, judged: false };
  let meet = true, fail = true, judged = false;
  ms.forEach(m => {
    const t = memberTarget(m.id, mk);
    const r = rateMap[m.id];
    if (t == null || r == null || isNaN(r)) return;
    judged = true;
    if (r >= t) fail = false; else meet = false;
  });
  return { meet: judged && meet, allFail: judged && fail, judged };
}
// 本周得分（只含「每周」规则项）
function calcWeekScore(personalHit, subHit, errs) {
  let s = 0;
  if (personalHit) s += 1;   // 个人完成率目标 +1
  if (subHit) s += 1;        // 小小组完成率达标 +1
  if (!errs) s += 1;         // 个人本周无严错 +1
  s -= errs * 0.5;           // 每产生 1 个严错 -0.5
  return Math.round(s * 10) / 10;
}
// 月度加分/扣分项存储
function monthExtras(mk, mid) {
  const T = state.team;
  if (!T.monthExtras) T.monthExtras = {};
  if (!T.monthExtras[mk]) T.monthExtras[mk] = {};
  if (!T.monthExtras[mk][mid]) T.monthExtras[mk][mid] = {};
  return T.monthExtras[mk][mid];
}
// 该月某成员的严错合计
function monthErrSum(mk, mid) {
  return periodsInMonth(mk).reduce((a, k) => a + (parseInt((((state.team.weeks[k] || {}).data || {})[mid] || {}).seriousErrors) || 0), 0);
}
// 该月是否已有任何一周数据
function hasWeekData(mk, mid) {
  return periodsInMonth(mk).some(k => { const d = (((state.team.weeks[k] || {}).data || {})[mid] || {}); return d.completionRate !== undefined && d.completionRate !== ''; });
}
// 该月最新一周的完成率映射
function latestRateMap(mk) {
  const ks = periodsInMonth(mk); const last = ks[ks.length - 1]; const map = {};
  if (!last) return map;
  const data = (state.team.weeks[last] || {}).data || {};
  state.team.members.forEach(m => { const r = parseFloat((data[m.id] || {}).completionRate); map[m.id] = isNaN(r) ? null : r; });
  return map;
}
// 该成员该月的「每月」规则得分
function calcMonthExtra(member, mk) { return calcMonthExtraFrom(monthExtras(mk, member.id), member, mk); }
// 纯函数：给定月度项数据算分（便于输入时实时预览）
function calcMonthExtraFrom(x, member, mk) {
  x = x || {};
  const errSum = monthErrSum(mk, member.id);
  let s = 0;
  // ---- 加分 ----
  const rank = parseInt(x.attendanceRank);
  if (rank === 1) s += 3; else if (rank === 2) s += 2; else if (rank === 3) s += 1;
  if (errSum === 0 && hasWeekData(mk, member.id)) s += 2;      // 无严错
  s += (parseInt(x.bonusActivity) || 0) * 3;                   // 组织活动 ×3
  s += (parseInt(x.bonusInitiative) || 0) * 2;                 // 主动补位 ×2
  s += (parseInt(x.bonusOnlineShare) || 0) * 1;                // 线上分享 ×1
  s += (parseInt(x.bonusEventOwner) || 0) * 2;                 // 周内大事件 ×2
  s += (parseInt(x.bonusTea) || 0) * 1;                        // 下午茶 ×1
  s += (parseInt(x.bonusQuiz) || 0) * 1;                       // 小组答题 ×1
  s += (parseInt(x.bonusTrainer) || 0) * 3;                    // 科室讲师 ×3
  s += (parseFloat(x.bonusDeptOther) || 0);                    // 科室其他活动 1~3，直接填分值
  // ---- 扣分 ----
  if (errSum >= 3) s -= 2;                                     // 严错≥3
  if (x.deductQualityRule) s -= 1;                             // 触及质量条例
  if (x.deductDragGroup) s -= 2;                               // 完成率低致小组不达标
  s -= (parseFloat(x.deductPending) || 0) * 0.5;               // pending/返稿/跳QC
  s -= (parseFloat(x.deductLow) || 0) * 1;                     // Low 违规
  s -= (parseFloat(x.deductMed) || 0) * 2;                     // Medium 违规
  s -= (parseFloat(x.deductHigh) || 0) * 3;                    // High 违规
  const sj = subGroupMeet(member.subGroupId, mk, latestRateMap(mk));
  if (sj.allFail) s -= 2;                                      // 小小组成员完成率均不达标
  return Math.round(s * 10) / 10;
}
// 某成员某周的得分（周项）
function weekScoreOf(pk, mid) {
  const T = state.team; const wk = T.weeks[pk]; if (!wk) return 0;
  const mk = monthOfPeriod(wk) || curMonth();
  const m = T.members.find(x => x.id === mid); if (!m) return 0;
  const d = (wk.data || {})[mid] || {};
  const rate = parseFloat(d.completionRate); const tgt = memberTarget(mid, mk);
  const personalHit = (!isNaN(rate) && tgt != null) ? rate >= tgt : false;
  const rateMap = {};
  T.members.forEach(x => { const r = parseFloat(((wk.data || {})[x.id] || {}).completionRate); rateMap[x.id] = isNaN(r) ? null : r; });
  const sj = subGroupMeet(m.subGroupId, mk, rateMap);
  return calcWeekScore(personalHit, sj.meet, parseInt(d.seriousErrors) || 0);
}
function flagBadge(v) { return v === null ? '<span class="wk-badge wk-badge-na">—</span>' : v ? '<span class="wk-badge wk-badge-yes">✓</span>' : '<span class="wk-badge wk-badge-no">✕</span>'; }

function renderTeamWeekly(v) {
  const T = state.team;
  const list = periodList();
  const sel = (T._selectedPeriod && T.weeks[T._selectedPeriod]) ? T._selectedPeriod : (list[0] || '');
  T._selectedPeriod = sel;
  v.innerHTML = `<div class="card">
    <div class="card-head"><h2>📝 周数据录入</h2><span class="ch-sub">只需填写完成率与严错数，得分自动按规则计算</span></div>
    <div class="wk-toolbar">
      <span class="wk-tb-label">日期区间</span>
      <select class="field" id="wkPeriodSel" style="min-width:200px">
        ${list.length ? list.map(k => `<option value="${k}" ${k === sel ? 'selected' : ''}>${esc(weekLabel(T.weeks[k]))}</option>`).join('') : '<option value="">（暂无区间）</option>'}
      </select>
      <span class="wk-tb-sep"></span>
      <input type="date" class="field" id="wkNewStart" value="${mondayStr()}" title="起始日期">
      <span class="wk-tb-tilde">~</span>
      <input type="date" class="field" id="wkNewEnd" value="${sundayStr()}" title="结束日期">
      <button class="btn primary sm" data-act="team-create-period">＋ 创建区间</button>
      ${sel ? `<button class="btn sm ghost" data-act="team-del-period" data-k="${sel}" style="color:var(--danger);margin-left:auto">🗑 删除该区间</button>` : ''}
    </div>
    <div class="tm-legend">${allSubGroupList().map((o, i) => `<span class="tm-lg"><i class="tm-dot tm-dot-${SG_COLOR_KEYS[i % SG_COLOR_KEYS.length]}"></i>${esc(o.sg.name)}</span>`).join('')}</div>
    <div id="weeklyForm"></div>
  </div>`;
  renderWeeklyForm(sel);
}

function renderWeeklyForm(pk) {
  const T = state.team;
  const box = $('#weeklyForm'); if (!box) return;
  if (!pk || !T.weeks[pk]) { box.innerHTML = `<div class="empty">还没有任何日期区间。填好上方起止日期后点「＋ 创建区间」开始录入。</div>`; return; }
  if (!T.members.length) { box.innerHTML = '<div class="empty">请先在「成员管理」中添加成员</div>'; return; }

  const wk = T.weeks[pk];
  const mk = monthOfPeriod(wk) || curMonth();
  const cmap = sgColorMap();
  const rateMap = {};
  T.members.forEach(m => { const r = parseFloat((((wk.data || {})[m.id]) || {}).completionRate); rateMap[m.id] = isNaN(r) ? null : r; });

  const rows = [];
  allSubGroupList().forEach(({ sg }) => {
    const color = cmap[sg.id] || 'tm-sg-a';
    const mems = T.members.filter(m => m.subGroupId === sg.id);
    const j = subGroupMeet(sg.id, mk, rateMap);
    rows.push(`<tr class="wk-sg-head ${color}"><td colspan="9">${esc(sg.name)}<span class="wk-sg-cnt">${mems.length} 人</span>
      <span class="wk-sg-judge">小小组达标：${flagBadge(j.judged ? j.meet : null)}</span></td></tr>`);
    if (!mems.length) { rows.push(`<tr class="wk-row ${color}"><td class="wk-empty" colspan="9">暂无成员</td></tr>`); return; }
    mems.forEach(m => {
      const d = (wk.data || {})[m.id] || {};
      const rate = rateMap[m.id];
      const ti = mTargetResolved(mk, m.id);
      const tgt = isNaN(parseFloat(ti.val)) ? null : parseFloat(ti.val);
      const errs = parseInt(d.seriousErrors) || 0;
      const personalHit = (rate != null && tgt != null) ? rate >= tgt : false;
      const score = calcWeekScore(personalHit, j.meet, errs);
      rows.push(`<tr class="wk-row ${color}">
        <td class="wk-mem">${esc(m.name)}</td>
        <td class="wk-tgt"><span class="wk-tgt-wrap"><input type="text" inputmode="decimal" class="wk-tgt-in" data-mid="${m.id}" data-mk="${mk}" value="${esc(ti.val)}" placeholder="未设" title="与「成员管理」的 ${esc(monthLabel(mk))} 目标联动，改完自动同步"> %<i class="wk-carry" style="display:${ti.carried && ti.val !== '' ? '' : 'none'}" title="该月未单独设置，沿用 ${esc(ti.from ? monthLabel(ti.from) : '历史数据')} 的目标">↩</i></span></td>
        <td><input type="text" inputmode="decimal" class="field wk-rate" data-mid="${m.id}" value="${d.completionRate !== undefined && d.completionRate !== '' ? esc(d.completionRate) : ''}" placeholder="—" style="width:84px;text-align:center"> %</td>
        <td><input type="text" inputmode="decimal" class="field wk-err" data-mid="${m.id}" value="${d.seriousErrors !== undefined && d.seriousErrors !== '' ? esc(d.seriousErrors) : ''}" placeholder="0" style="width:60px;text-align:center"></td>
        <td class="wk-flag" data-flag="p-${m.id}">${flagBadge(rate == null || tgt == null ? null : personalHit)}</td>
        <td class="wk-flag" data-flag="s-${m.id}">${flagBadge(j.judged ? j.meet : null)}</td>
        <td class="wk-flag" data-flag="e-${m.id}">${flagBadge(errs === 0)}</td>
        <td class="wk-score ${score >= 0 ? 'score-pos' : 'score-neg'}" data-score="${m.id}"><b>${score}</b></td>
        <td><button class="btn sm ghost wk-more ${(T._openDetails || {})[m.id] ? 'on' : ''}" data-act="wk-toggle-more" data-mid="${m.id}" title="展开本月加分/扣分项">⋯</button></td>
      </tr>`);
      rows.push(`<tr class="wk-detail" data-detail="${m.id}" style="display:${(T._openDetails || {})[m.id] ? '' : 'none'}"><td colspan="9">${monthlyDetailHtml(m, mk)}</td></tr>`);
    });
  });

  const monthEmpty = !monthHasTargets(mk);
  const warn = monthEmpty ? `<div class="wk-warn">⚠️ <b>${esc(monthLabel(mk))}</b> 还没设置成员目标，下表「个人目标」为其他月份的沿用值（带 <i class="wk-carry">↩</i>）。<button class="btn xs" data-act="team-goto-members-month" data-mk="${mk}">去「成员管理」设置 ${esc(monthLabel(mk))} 目标 →</button></div>` : '';

  box.innerHTML = `${warn}<div class="wk-scroll"><table class="team-table wk-table">
    <thead><tr>
      <th style="min-width:120px">成员</th><th style="width:104px">个人目标</th><th style="width:142px">完成率</th><th style="width:90px">严错数</th>
      <th style="width:74px">个人达标</th><th style="width:82px">小小组达标</th><th style="width:64px">无严错</th><th style="width:74px">本周得分</th><th style="width:36px"></th>
    </tr></thead>
    <tbody>${rows.join('')}</tbody>
    <tfoot><tr class="wk-tfoot">
      <td colspan="3">本周合计（${esc(weekLabel(wk))}）</td>
      <td colspan="4" style="text-align:right">全员本周得分合计</td>
      <td class="wk-score score-pos" id="wkTotalCell"><b>${T.members.reduce((a, m) => a + weekScoreOf(pk, m.id), 0).toFixed(1)}</b></td>
      <td></td>
    </tr></tfoot>
  </table></div>
  <div class="wk-tipbox">💡 「个人目标」直接取自 <b>成员管理</b> 里 ${esc(monthLabel(mk))} 的目标，也可以在这里直接改（改完自动同步回成员管理）。「个人达标 / 小小组达标 / 无严错」由系统按 <b>加减分规则</b> 自动判定，本周得分 = 达标项加分 − 严错扣分；每月项目点行末 <b>⋯</b> 展开填写。</div>
  <div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap">
    <button class="btn primary" data-act="team-save-period" data-k="${pk}">💾 保存本周数据</button>
    <button class="btn yellow sm" data-act="team-show-rules">📋 查看加减分规则</button>
  </div>`;
}

// 月度加分/扣分项展开面板
function monthlyDetailHtml(m, mk) {
  const x = monthExtras(mk, m.id);
  const errSum = monthErrSum(mk, m.id);
  const sj = subGroupMeet(m.subGroupId, mk, latestRateMap(mk));
  const num = (k, label, mult) => `<label class="wk-mi"><span>${label}</span>
    <input type="text" inputmode="decimal" class="field wk-mx" data-mid="${m.id}" data-mk="${mk}" data-fld="${k}" value="${x[k] !== undefined && x[k] !== '' ? esc(x[k]) : ''}" placeholder="0" style="width:48px"><em>×${mult}</em></label>`;
  const chk = (k, label, pts) => `<label class="wk-mi wk-mi-chk"><input type="checkbox" class="wk-mxc" data-mid="${m.id}" data-mk="${mk}" data-fld="${k}" ${x[k] ? 'checked' : ''}><span>${label}</span><em>−${pts}</em></label>`;
  const extra = calcMonthExtra(m, mk);
  return `<div class="wk-detail-box">
    <div class="wk-detail-head">📆 ${esc(monthLabel(mk))} 月度项（每月统计一次）</div>
    <div class="wk-mgrid">
      <label class="wk-mi"><span>应出勤排名</span>
        <select class="field wk-mxs" data-mid="${m.id}" data-mk="${mk}" data-fld="attendanceRank" style="width:92px">
          <option value="" ${!x.attendanceRank ? 'selected' : ''}>无</option>
          <option value="1" ${x.attendanceRank == 1 ? 'selected' : ''}>Top1 +3</option>
          <option value="2" ${x.attendanceRank == 2 ? 'selected' : ''}>Top2 +2</option>
          <option value="3" ${x.attendanceRank == 3 ? 'selected' : ''}>Top3 +1</option>
        </select></label>
      ${num('bonusActivity', '组织活动', 3)}
      ${num('bonusInitiative', '主动补位', 2)}
      ${num('bonusOnlineShare', '线上分享', 1)}
      ${num('bonusEventOwner', '周内大事件', 2)}
      ${num('bonusTea', '组织下午茶', 1)}
      ${num('bonusQuiz', '小组答题', 1)}
      ${num('bonusTrainer', '科室讲师', 3)}
      ${num('bonusDeptOther', '科室其他活动', '1~3')}
      ${num('deductPending', 'pending/返稿/跳QC', 0.5)}
      ${num('deductLow', 'Low 违规', 1)}
      ${num('deductMed', 'Med 违规', 2)}
      ${num('deductHigh', 'High 违规', 3)}
      ${chk('deductQualityRule', '触及质量条例', 1)}
      ${chk('deductDragGroup', '完成率低致小组不达标', 2)}
    </div>
    <div class="wk-auto-line"><span class="wk-auto-tag">自动判定</span>
      本月严错合计 <b>${errSum}</b> 个 ·
      无严错 ${errSum === 0 && hasWeekData(mk, m.id) ? '<b class="wk-ok">+2</b>' : '<span class="wk-na">不满足</span>'} ·
      严错≥3 ${errSum >= 3 ? '<b class="wk-bad">−2</b>' : '<span class="wk-na">不满足</span>'} ·
      小小组均不达标 ${sj.allFail ? '<b class="wk-bad">−2</b>' : '<span class="wk-na">不满足</span>'}
    </div>
    <div class="wk-detail-total">当前月度项小计 <b class="${extra >= 0 ? 'wk-ok' : 'wk-bad'}" data-mxtotal="${m.id}">${extra}</b> 分</div>
  </div>`;
}

// 输入时实时重算（不落库）
function recalcWeeklyUI() {
  const T = state.team;
  const pk = T._selectedPeriod; const wk = T.weeks[pk]; if (!wk) return;
  const mk = monthOfPeriod(wk) || curMonth();
  const rateMap = {}, errMap = {};
  $$('.wk-rate').forEach(i => { const n = parseFloat(i.value); rateMap[i.dataset.mid] = isNaN(n) ? null : n; });
  $$('.wk-err').forEach(i => { const n = parseInt(i.value); errMap[i.dataset.mid] = isNaN(n) ? 0 : n; });
  allSubGroupList().forEach(({ sg }) => {
    const j = subGroupMeet(sg.id, mk, rateMap);
    const headFlag = $(`.wk-sg-head .wk-sg-judge`); // 组头单独刷新见下
    T.members.filter(m => m.subGroupId === sg.id).forEach(m => {
      const rate = rateMap[m.id], tgt = memberTarget(m.id, mk), errs = errMap[m.id] || 0;
      const personalHit = (rate != null && tgt != null) ? rate >= tgt : false;
      const setF = (pre, v) => { const el = $(`[data-flag="${pre}${m.id}"]`); if (el) el.innerHTML = flagBadge(v); };
      setF('p-', (rate == null || tgt == null) ? null : personalHit);
      setF('s-', j.judged ? j.meet : null);
      setF('e-', errs === 0);
      const sc = calcWeekScore(personalHit, j.meet, errs);
      const cell = $(`[data-score="${m.id}"]`);
      if (cell) { cell.innerHTML = `<b>${sc}</b>`; cell.className = 'wk-score ' + (sc >= 0 ? 'score-pos' : 'score-neg'); }
    });
  });
  // 组头达标徽标
  $$('.wk-sg-head').forEach((tr, idx) => {
    const o = allSubGroupList()[idx]; if (!o) return;
    const j = subGroupMeet(o.sg.id, mk, rateMap);
    const el = tr.querySelector('.wk-sg-judge');
    if (el) el.innerHTML = '小小组达标：' + flagBadge(j.judged ? j.meet : null);
  });
  const tc = $('#wkTotalCell');
  if (tc) {
    let t = 0;
    T.members.forEach(m => {
      const rate = rateMap[m.id], tgt = memberTarget(m.id, mk);
      const personalHit = (rate != null && tgt != null) ? rate >= tgt : false;
      const sj = subGroupMeet(m.subGroupId, mk, rateMap);
      t += calcWeekScore(personalHit, sj.meet, errMap[m.id] || 0);
    });
    tc.innerHTML = `<b>${t.toFixed(1)}</b>`;
  }
}

// 月度项输入时实时刷新该成员的小计
function recalcMonthExtraUI(mid, mk) {
  const m = state.team.members.find(x => x.id === mid);
  if (!m || !mk) return;
  const x = {};
  $$(`.wk-mx[data-mid="${mid}"], .wk-mxc[data-mid="${mid}"], .wk-mxs[data-mid="${mid}"]`).forEach(i => {
    x[i.dataset.fld] = i.type === 'checkbox' ? i.checked : (i.value || '').trim();
  });
  const extra = calcMonthExtraFrom(x, m, mk);
  const el = $(`[data-mxtotal="${mid}"]`);
  if (el) { el.textContent = extra; el.className = extra >= 0 ? 'wk-ok' : 'wk-bad'; }
}
// 失焦即落库：单个「完成率 / 严错数」输入
function commitWeekInput(input) {
  const T = state.team;
  const wk = T.weeks[T._selectedPeriod];
  const mid = input.dataset.mid;
  if (!wk || !mid) return;
  wk.data = wk.data || {}; wk.data[mid] = wk.data[mid] || {};
  const fld = input.classList.contains('wk-rate') ? 'completionRate' : 'seriousErrors';
  const clean = (input.value || '').trim().replace(/[^\d.]/g, '');
  wk.data[mid][fld] = clean;
  if (input.value !== clean) input.value = clean;
  save();
}
// 周面板里直接改「个人目标」→ 同步写回该月的 targets（与「成员管理」同一份数据）
function commitWeekTarget(input) {
  const mk = input.dataset.mk || curMonth();
  const mid = input.dataset.mid;
  if (!mid) return;
  const clean = (input.value || '').trim().replace(/[^\d.]/g, '');
  if (input.value !== clean) input.value = clean;
  const mo = ensureMonth(mk);
  if (clean === '') delete mo.targets[mid]; else mo.targets[mid] = clean;
  save();
  renderTeam();          // 整块刷新：同步「成员管理」、更新达标判定与得分、去掉沿用标记
  toast(`✅ 已同步为「成员管理」${monthLabel(mk)} 的目标`);
}
// 失焦即落库：单个月度项
function commitMonthExtra(mid, mk, fld) {
  if (!mid || !mk || !fld) return;
  const el = $(`.wk-mx[data-mid="${mid}"][data-fld="${fld}"], .wk-mxc[data-mid="${mid}"][data-fld="${fld}"], .wk-mxs[data-mid="${mid}"][data-fld="${fld}"]`);
  if (!el) return;
  monthExtras(mk, mid)[fld] = el.type === 'checkbox' ? el.checked : (el.value || '').trim();
  save();
}

/* ---- 积分核算引擎（基于 TEAM_RULES 自动核算） ---- */
// 周项得分：见 weekScoreOf()；月度项得分：见 calcMonthExtra()
// 某成员在某月的完整得分 = 该月各周得分之和 + 月度项得分
function monthTotalOf(mid, mk) {
  const m = state.team.members.find(x => x.id === mid); if (!m) return 0;
  const weeksSum = periodsInMonth(mk).reduce((a, k) => a + weekScoreOf(k, mid), 0);
  return Math.round((weeksSum + calcMonthExtra(m, mk)) * 10) / 10;
}
// 兼容旧签名：按周累计（截至 upToPeriod）
function calcMonthTotal(memberId, upToPeriod) {
  const mk = upToPeriod ? (monthOfPeriod(state.team.weeks[upToPeriod]) || curMonth()) : curMonth();
  let total = periodsInMonth(mk).filter(k => !upToPeriod || k <= upToPeriod).reduce((a, k) => a + weekScoreOf(k, memberId), 0);
  const m = state.team.members.find(x => x.id === memberId);
  if (m) total += calcMonthExtra(m, mk);
  return Math.round(total * 10) / 10;
}

/* ---- 积分看板 ---- */
function renderTeamScoreboard(v) {
  const T = state.team;
  const months = [...new Set(periodList().map(k => monthOfPeriod(T.weeks[k])).filter(Boolean))];
  if (!months.includes(curMonth())) months.push(curMonth());
  months.sort().reverse();
  const selMonth = (T._sbMonth && months.includes(T._sbMonth)) ? T._sbMonth : months[0];
  T._sbMonth = selMonth;

  v.innerHTML = `<div class="card"><div class="card-head"><h2>🏆 积分看板</h2>
    <span class="ch-sub">月份: <select class="field" id="sbMonthSelect" style="width:130px">
        ${months.map(m => `<option value="${m}" ${m === selMonth ? 'selected' : ''}>${esc(monthLabel(m))}</option>`).join('')}
      </select></span>
  </div>
  <div id="sbContent"></div></div>`;
  renderSBContent(selMonth);
}
function renderSBContent(mk) {
  const T = state.team;
  const c = $('#sbContent');
  if (!c) return;
  const monthWeeks = periodsInMonth(mk);

  if (monthWeeks.length === 0) { c.innerHTML = `<div class="empty">${esc(monthLabel(mk))} 暂无周数据，请先在「周数据录入」中创建日期区间</div>`; return; }

  // 月度累计排行（各周得分 + 月度项）
  const monthlyTotals = T.members.map(m => ({ id: m.id, name: m.name, total: monthTotalOf(m.id, mk) })).sort((a, b) => b.total - a.total);

  let html = `<div class="card card-soft" style="margin-bottom:16px">
    <h3 style="margin:0 0 10px;font-size:15px">🥇 ${esc(monthLabel(mk))} 累计积分 Top 排行 <span style="font-weight:400;color:var(--ink-faint);font-size:12px">（各周得分 + 月度项）</span></h3>
    <div class="podium">`;
  monthlyTotals.forEach((m, i) => {
    const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `${i + 1}`;
    html += `<div class="podium-item ${i < 3 ? 'podium-top3' : ''}">
      <span class="podium-rank">${medal}</span>
      <span class="podium-name">${esc(m.name)}</span>
      <span class="podium-score ${m.total >= 0 ? 'score-pos' : 'score-neg'}">${m.total}</span>
    </div>`;
  });
  html += `</div>
    ${monthlyTotals.length > 0 && monthlyTotals[0].total > 0 ? '<p style="margin:8px 0 0;font-size:12px;color:var(--ink-faint)">🎁 每月积分 Top 3 有价值不等的额外惊喜哦～</p>' : ''}
  </div>`;

  // 各周明细（按小小组上色）
  html += `<h3 style="margin:16px 0 10px;font-size:15px">📊 各周积分明细</h3>
    <div class="card card-soft"><div style="overflow-x:auto">
    <table class="team-table">
      <thead><tr><th>成员</th>${monthWeeks.map(k => `<th>${esc(weekLabel(T.weeks[k]))}</th>`).join('')}<th>月度项</th><th>月合计</th></tr></thead>
      <tbody>`;
  const cmap = sgColorMap();
  allSubGroupList().forEach(({ sg }) => {
    const color = cmap[sg.id] || 'tm-sg-a';
    const mems = T.members.filter(m => m.subGroupId === sg.id);
    if (!mems.length) return;
    html += `<tr class="wk-sg-head ${color}"><td colspan="${monthWeeks.length + 3}">${esc(sg.name)}<span class="wk-sg-cnt">${mems.length} 人</span></td></tr>`;
    mems.forEach(m => {
      html += `<tr class="wk-row ${color}"><td class="wk-mem">${esc(m.name)}</td>`;
      let wsum = 0;
      monthWeeks.forEach(k => {
        const d = (((T.weeks[k] || {}).data || {})[m.id] || {});
        const has = d.completionRate !== undefined && d.completionRate !== '';
        const s = has ? weekScoreOf(k, m.id) : '-';
        if (typeof s === 'number') wsum += s;
        const cls = typeof s === 'number' ? (s >= 0 ? 'score-pos' : 'score-neg') : '';
        html += `<td class="${cls}" style="text-align:center">${s}</td>`;
      });
      const extra = calcMonthExtra(m, mk);
      const total = Math.round((wsum + extra) * 10) / 10;
      html += `<td class="${extra >= 0 ? 'score-pos' : 'score-neg'}" style="text-align:center">${extra}</td>
        <td class="${total >= 0 ? 'score-pos' : 'score-neg'}" style="text-align:center;font-weight:800">${total}</td></tr>`;
    });
  });
  html += `</tbody></table></div></div>`;

  c.innerHTML = html;
}


function historyDay(d) {
  const list = state.work.todos[d] || [];
  const done = list.filter(x => x.done).length;
  return `<div class="todo-day collapsed" data-date="${d}">
    <div class="todo-day-head" data-act="day-toggle" data-date="${d}">
      <span class="caret">▾</span><span>📅 ${d} ${WEEK[new Date(d).getDay()]}</span>
      <span class="tw">${done}/${list.length}</span>
    </div>
    <div class="todo-day-body">${list.length ? list.map(todoRow).join('') : '<div class="empty">无记录</div>'}</div>
  </div>`;
}

/* ---------- 报告 ---------- */
function dateRange(type, base = new Date()) {
  const y = base.getFullYear(), m = base.getMonth(), d = base.getDate();
  if (type === 'week') {
    const dow = base.getDay(); // 0=Sun..6=Sat
    const mon = new Date(base); mon.setDate(d - ((dow + 6) % 7));
    const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
    const f = x => `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`;
    return { s: f(mon), e: f(sun) };
  }
  if (type === 'half') {
    if (d <= 15) return { s: `${y}-${String(m + 1).padStart(2, '0')}-01`, e: `${y}-${String(m + 1).padStart(2, '0')}-15` };
    const ld = daysInMonth(y, m);
    return { s: `${y}-${String(m + 1).padStart(2, '0')}-16`, e: `${y}-${String(m + 1).padStart(2, '0')}-${ld}` };
  }
  if (type === 'month') {
    const ld = daysInMonth(y, m);
    return { s: `${y}-${String(m + 1).padStart(2, '0')}-01`, e: `${y}-${String(m + 1).padStart(2, '0')}-${ld}` };
  }
  if (type === 'quarter') {
    const q = Math.floor(m / 3); const sm = q * 3; const em = sm + 2;
    const ld = daysInMonth(y, em);
    return { s: `${y}-${String(sm + 1).padStart(2, '0')}-01`, e: `${y}-${String(em + 1).padStart(2, '0')}-${ld}` };
  }
  if (type === 'year') return { s: `${y}-01-01`, e: `${y}-12-31` };
  return null;
}
function openReport() {
  const types = [['week', '周报'], ['half', '半月报'], ['month', '月报'], ['quarter', '季报'], ['year', '年报']];
  const now = new Date(); const y = now.getFullYear(), m = now.getMonth() + 1;
  const ms = String(m).padStart(2, '0');
  const monthStart = `${y}-${ms}-01`;
  const monthEnd = `${y}-${ms}-${String(daysInMonth(y, m - 1)).padStart(2, '0')}`;
  openModal(`<h2>📊 工作汇报</h2>
    <p style="color:var(--ink-soft);font-size:13px;margin:0 0 6px">选择区间，按你的模板汇总：<strong>4.本周/本期工作（组内 · 科室 · 部门）</strong> → <strong>5.下周/下阶段工作计划</strong>（未完成事项自动转为计划）。可导出 Markdown 或 Word 文件。</p>
    <div class="row" id="repTypes">
      ${types.map(([v, l]) => `<button class="btn sm" data-act="report-gen" data-range="${v}">${l}</button>`).join('')}
    </div>
    <div class="row" style="margin-top:10px;align-items:center;flex-wrap:wrap">
      <span class="ch-sub">自定义区间：</span>
      <label>起 <input type="date" class="field" id="repS" value="${monthStart}" style="width:150px"></label>
      <label>止 <input type="date" class="field" id="repE" value="${monthEnd}" style="width:150px"></label>
      <button class="btn primary sm" data-act="report-gen-custom">生成</button>
    </div>
    <div id="repOut" style="margin-top:10px"></div>`);
  // 防止点日期控件时冒泡误触关闭弹窗
  ['repS', 'repE'].forEach(id => { const el = $('#' + id); if (el) el.addEventListener('click', e => e.stopPropagation()); });
}
function genReport(range, custom) {
  let r, label, hd;
  if (range === 'custom') {
    if (!custom || !custom.s || !custom.e) { toast('请选择起止日期'); return; }
    r = { s: custom.s, e: custom.e }; label = '工作汇报';
    hd = { work: '本期工作', plan: '下阶段工作计划与目标' };
  } else {
    r = dateRange(range);
    const M = {
      week:    { title: '工作周报', work: '本周工作', plan: '下周工作计划与目标' },
      half:    { title: '工作半月报', work: '本期工作', plan: '下阶段工作计划与目标' },
      month:   { title: '工作月报', work: '本月工作', plan: '下月工作计划与目标' },
      quarter: { title: '工作季报', work: '本季工作', plan: '下季工作计划与目标' },
      year:    { title: '工作年报', work: '本年工作', plan: '明年工作计划与目标' },
    };
    label = M[range].title; hd = { work: M[range].work, plan: M[range].plan };
  }
  if (!r) { toast('区间无效'); return; }
  const all = [];
  Object.keys(state.work.todos).forEach(d => {
    if (d >= r.s && d <= r.e) state.work.todos[d].forEach(t => all.push(Object.assign({ date: d }, t)));
  });
  all.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
  const done = all.filter(t => t.done).length;
  const rate = all.length ? Math.round(done / all.length * 100) : 0;
  const open = all.filter(t => !t.done);
  const days = new Set(all.map(t => t.date)).size;
  const catOrder = CATS;
  const byCat = {}; catOrder.forEach(c => byCat[c] = []);
  all.forEach(t => { const c = t.cat || '其他'; (byCat[c] = byCat[c] || []).push(t); });
  const extraCats = Object.keys(byCat).filter(c => !catOrder.includes(c));
  const scopeStats = {}; SCOPES.forEach(s => scopeStats[s] = 0);
  all.forEach(t => { const s = t.scope || SCOPES[0]; scopeStats[s] = (scopeStats[s] || 0) + 1; });
  const scopeLine = SCOPES.filter(s => scopeStats[s] > 0).map(s => `${s} ${scopeStats[s]}`).join(' · ');

  const L = [];
  L.push(`# ${label}`);
  L.push('');
  L.push(`> 统计区间：**${r.s} 至 ${r.e}**　|　共 ${days} 天有记录`);
  L.push('> 自动汇总自「小煎蛋的工作台 · 工作」分栏');
  L.push('');
  if (all.length) {
    L.push(`本期共记录 ${all.length} 项任务，完成 ${done} 项（完成率 ${rate}%）。`);
    if (scopeLine) L.push(`归属分布：${scopeLine}。`);
  }
  else L.push('该区间暂无任何工作记录，先去「工作」分栏记几笔吧。');
  L.push('');
  // 4. 本周/本期工作（按「工作性质」分组，每条标注归属范围）
  L.push(`4.${hd.work}：`);
  catOrder.concat(extraCats).forEach(c => {
    const items = byCat[c]; if (!items || !items.length) return;
    L.push(`○${c}：`);
    items.forEach(t => L.push(`- ${t.text} · ${t.scope || SCOPES[0]} · ${t.date}${t.done ? ' ✓' : '（未完成）'}`));
  });
  L.push('');
  // 5. 下周/下阶段工作计划与目标（未完成事项自动转化为计划）
  L.push(`5.${hd.plan}：`);
  if (open.length) open.forEach(t => L.push(`- ${t.text}（${t.cat || '其他'} · ${t.scope || SCOPES[0]} · 原定 ${t.date}）`));
  else L.push('- （本期任务已全部完成，可在此补充新目标）');
  const md = L.join('\n');

  lastReport = { md, name: `工作汇报_${label}_${r.s}_${r.e}.md`, docxName: `工作汇报_${label}_${r.s}_${r.e}.docx` };
  const outHtml = `
    <div class="report-summary">
      <div class="stat-row">
        <div class="stat"><div class="num">${all.length}</div><div class="lab">任务总数</div></div>
        <div class="stat"><div class="num">${rate}%</div><div class="lab">完成率</div></div>
        <div class="stat"><div class="num">${done}</div><div class="lab">已完成</div></div>
        <div class="stat"><div class="num">${open.length}</div><div class="lab">待跟进</div></div>
      </div>
      <p style="color:var(--ink-soft);font-size:13px;margin:10px 0">区间 ${r.s} ~ ${r.e}，共 ${days} 天有记录。按「工作性质」分小标题，每条标注「归属范围」→ 5.下周/下阶段工作计划（未完成事项自动转为计划）。可导出 Markdown 或 Word(.docx)。</p>
      <pre class="report-md" style="white-space:pre-wrap;background:#fff;border:1px solid #eee;border-radius:10px;padding:12px;font-size:12px;line-height:1.6;max-height:340px;overflow:auto">${esc(md)}</pre>
    </div>
    <div style="margin-top:12px;display:flex;gap:8px;flex-wrap:wrap">
      <button class="btn sm yellow" data-act="report-export">⬇ 导出 Markdown</button>
      <button class="btn sm" data-act="report-export-docx">⬇ 导出 Word (.docx)</button>
      <button class="btn sm ghost" data-act="report-copy">📋 复制全文</button>
    </div>`;
  $('#repOut').innerHTML = outHtml;
  $('#repOut').dataset.payload = md;
}

/* 报告动作统一处理（#view 与 #modalRoot 共用：弹窗按钮渲染在 #modalRoot，不在 #view 内） */
function handleReport(act, el) {
  if (act === 'report-gen') { genReport(el.dataset.range); }
  else if (act === 'report-gen-custom') {
    const s = $('#repS').value, e = $('#repE').value;
    if (!s || !e) { toast('请选择起止日期'); return; }
    genReport('custom', { s, e });
  }
  else if (act === 'report-copy') {
    const payload = $('#repOut').dataset.payload || '';
    navigator.clipboard?.writeText(payload).then(() => toast('已复制全文')).catch(() => toast('复制失败'));
  }
  else if (act === 'report-export') {
    if (!lastReport.md) { toast('请先生成报告'); return; }
    download(lastReport.name, lastReport.md, 'text/markdown'); toast('已导出 Markdown');
  }
  else if (act === 'report-export-docx') {
    if (!lastReport.md) { toast('请先生成报告'); return; }
    download(lastReport.docxName, buildDocx(lastReport.md), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    toast('已导出 Word 文档');
  }
}

/* =================== 英语 =================== */
function enStreak() {
  let streak = 0; const d = new Date();
  while (true) {
    const k = todayStr(d);
    const c = state.english.checkins[k];
    if (c && (c.listen || c.speak || c.vocab || c.meet)) { streak++; d.setDate(d.getDate() - 1); }
    else break;
  }
  return streak;
}
function enHeatmap() {
  const cells = [];
  for (let i = 29; i >= 0; i--) {
    const d = addDays(new Date(), -i); const k = todayStr(d);
    const c = state.english.checkins[k];
    const n = c ? ['listen', 'speak', 'vocab', 'meet'].filter(x => c[x]).length : 0;
    cells.push({ k, n });
  }
  return cells;
}
function enSkillCount(type) { let n = 0; Object.values(state.english.checkins).forEach(c => { if (c && c[type]) n++; }); return n; }
function enTotalDays() { return Object.values(state.english.checkins).filter(c => c && (c.listen || c.speak || c.vocab || c.meet)).length; }
function viewEnglish(v) {
  const t = todayStr();
  const c = state.english.checkins[t] || {};
  const subs = state.english.subs[t] || {};
  const doneCount = ['listen', 'speak', 'vocab', 'meet'].filter(k => c[k]).length;
  const vocab = todayVocab();
  const mastered = state.english.vocabMastered;

  const planCards = EN_PLAN.map(p => {
    const done = c[p.type];
    const sub = subs[p.type] || {};
    let body;
    if (p.type === 'vocab') {
      body = `<div class="vocab-list">${vocab.map(w => {
        const ok = mastered.includes(w.id);
        return `<div class="vocab-item ${ok ? 'on' : ''}">
          <button class="vocab-spk" data-act="en-play" data-text="${encodeURIComponent(w.en)}" title="听发音">🔊</button>
          <div class="vocab-main"><b>${esc(w.en)}</b> <span class="vocab-cn">${esc(w.cn)}</span>
            <div class="vocab-ex">${esc(w.ex)}</div></div>
          <button class="btn xs ${ok ? 'yellow' : 'ghost'}" data-act="en-vocab" data-id="${w.id}">${ok ? '★ 已掌握' : '标记掌握'}</button>
        </div>`;
      }).join('')}</div>`;
    } else {
      body = `<div class="en-material">${esc(p.material).split('\n').map(l => `<div>${esc(l)}</div>`).join('')}</div>
        <button class="btn xs ghost" data-act="en-play" data-text="${encodeURIComponent(p.material)}">🔊 听发音</button>`;
    }
    const steps = (p.steps || []).map((s, i) => {
      const ck = sub[i];
      return `<label class="en-step ${ck ? 'on' : ''}"><input type="checkbox" ${ck ? 'checked' : ''} data-act="en-sub" data-type="${p.type}" data-idx="${i}"> <span>${esc(s)}</span></label>`;
    }).join('');
    return `<div class="card card-soft plan-card ${done ? 'done' : ''}">
      <div class="plan-top">
        <div class="plan-ic ${p.cls}">${p.icon}</div>
        <div><h3>${p.title}</h3><p class="plan-desc">${esc(p.desc)}</p></div>
      </div>
      ${body}
      <div class="en-steps">${steps}</div>
      <div class="checkin-row">
        <button class="btn sm ${done ? 'primary' : ''}" data-act="en-check" data-type="${p.type}">${done ? '✓ 已完成打卡' : '打卡完成'}</button>
      </div>
    </div>`;
  }).join('');

  v.innerHTML = `
  <div class="card">
    <div class="card-head"><span class="ch-emoji">📚</span><h2>每日英语打卡</h2>
      <span class="ch-sub">流利说 LV4 · 今日 ${doneCount}/4</span></div>
    <p style="color:var(--ink-soft);font-size:13.5px;margin:0 0 10px">
      你的发音不错，重点补<b>听力</b>与<b>输出</b>。下面四项都能<b>直接在页面完成</b>（点 🔊 可听发音，不用跳转外部网站），每天约 30 分钟，坚持一个月开会/闲聊更顺。</p>
    <div class="grid grid-2">${planCards}</div>
    <div style="margin-top:12px" class="pill yellow">🔥 连续打卡 ${enStreak()} 天</div>
  </div>
  <div class="card card-soft">
    <div class="card-head"><span class="ch-emoji">📈</span><h2>打卡轨迹（近 30 天）</h2>
      <span class="ch-sub">共打卡 ${enTotalDays()} 天</span></div>
    <div class="heatmap">${enHeatmap().map(c => `<div class="hm-cell hm-${c.n}" title="${c.k}">${c.n ? '●' : ''}</div>`).join('')}</div>
    <div class="skill-bars">
      ${EN_PLAN.map(p => `<div class="skill-bar"><span class="sb-name">${p.title.split(' ')[0]}</span><div class="sb-track"><div class="sb-fill" style="width:${Math.round(enSkillCount(p.type) / 30 * 100)}%"></div></div><span class="sb-num">${enSkillCount(p.type)} 天</span></div>`).join('')}
    </div>
  </div>`;
}

/* =================== 理财 =================== */
function viewFinance(v) {
  const tipIdx = new Date().getDate() % state.finance.knowledge.length;
  const tip = state.finance.knowledge[tipIdx];
  v.innerHTML = `
  <div class="card" style="background:linear-gradient(135deg,var(--green-100),var(--yellow-50))">
    <div class="card-head"><span class="ch-emoji">💡</span><h2>今日理财小知识</h2></div>
    <div style="display:flex;gap:12px;align-items:center">
      <div style="font-size:30px">${['🌱','📈','💡','🪙','📊'][tipIdx % 5]}</div>
      <div><b style="font-size:15px">${esc(tip.title)}</b><p style="margin:2px 0 0;color:var(--ink-soft);font-size:13px">${esc(tip.body)}</p></div>
    </div>
  </div>
  <div class="card">
    <div class="card-head"><span class="ch-emoji">📚</span><h2>理财知识库</h2>
      <span class="ch-sub">
        <button class="btn sm yellow" data-act="quiz-open">🧭 风险测评</button>
        <button class="btn sm ghost" data-act="know-add">＋ 记一条</button>
      </span></div>
    <div class="know-filterbar" id="knowFilter">
      ${['全部', ...Array.from(new Set(state.finance.knowledge.map(k => k.tag)))].map(tg => `<button class="chip ${tg === knowFilter ? 'on' : ''}" data-tag="${tg}">${tg}</button>`).join('')}
    </div>
    <div class="grid grid-3">
      ${state.finance.knowledge.filter(k => knowFilter === '全部' || k.tag === knowFilter).map(k => `
        <div class="card card-soft know-card">
          <span class="pill green">${esc(k.tag)}</span>
          <h3>${esc(k.title)}</h3><p>${esc(k.body)}</p>
        </div>`).join('')}
    </div>
  </div>
  <div class="card card-soft">
    <div class="card-head"><span class="ch-emoji">📡</span><h2>财经动向 / 趋势</h2>
      <span class="ch-sub"><a class="jump-link" href="https://www.bilibili.com/search/all?keyword=%E8%B4%A2%E7%BB%8F%E6%96%B0%E9%97%BB" target="_blank" rel="noopener">看原平台 ▶</a></span></div>
    <p style="color:var(--ink-soft);font-size:13.5px;margin:0">
      想了解实时行情与产品，可连接 <b>腾讯自选股 / 盈米 / Wind</b> 等连接器获取推送；
      当前为本地知识库，建议每周挑 1–2 个概念（如 index fund、资产配置）深入了解后再试小额产品。</p>
  </div>`;
}

/* =================== 食谱 =================== */
function daySeed(dStr) {
  // 用日期生成稳定“今日三餐”选择
  if (!state.recipes.dailySeed[dStr]) {
    const pool = SAMPLE_RECIPES.slice();
    const pick = (meal) => {
      const cands = pool.filter(r => r.meal === meal);
      const base = parseInt(dStr.replace(/-/g, ''), 10);
      return cands[base % cands.length];
    };
    state.recipes.dailySeed[dStr] = { breakfast: pick('breakfast').id, lunch: pick('lunch').id, dinner: pick('dinner').id };
    save();
  }
  return state.recipes.dailySeed[dStr];
}
function viewRecipes(v) {
  const t = todayStr();
  const seed = daySeed(t);
  const meals = [['breakfast', '🌞 早餐', '🍳'], ['lunch', '🍱 午餐（带饭友好）', '🍱'], ['dinner', '🌙 晚餐', '🍲']];
  const libFilter = recipeFilter === 'fav'
    ? SAMPLE_RECIPES.filter(r => state.recipes.favs.includes(r.id))
    : recipeFilter === 'all' ? SAMPLE_RECIPES : SAMPLE_RECIPES.filter(r => r.meal === recipeFilter);
  v.innerHTML = `
  <div class="card">
    <div class="card-head"><span class="ch-emoji">🍳</span><h2>今日三餐</h2>
      <span class="ch-sub"><button class="btn sm ghost" data-act="recipe-shuffle">🔄 换一批</button></span></div>
    <div class="grid grid-3">
      ${meals.map(([m, label, ic]) => {
        const r = SAMPLE_RECIPES.find(x => x.id === seed[m]);
        const fav = state.recipes.favs.includes(r.id);
        return `<div class="card card-soft recipe-card">
          <div class="rc-head"><div class="rc-thumb">${r.emoji}</div>
            <div><h3 style="margin:0;font-size:15px">${label}</h3>
            <span style="font-size:12px;color:var(--ink-soft)">${r.name} · 约 ${r.time} 分钟</span></div></div>
          <ul class="recipe-steps">${r.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ul>
          <div class="recipe-tags">${r.tags.map(tg => `<span class="pill green">${esc(tg)}</span>`).join('')}</div>
          <div class="checkin-row">
            <button class="btn sm ${fav ? 'yellow' : 'ghost'}" data-act="recipe-fav" data-id="${r.id}">${fav ? '★ 已收藏' : '☆ 收藏'}</button>
            <a class="jump-link" href="https://www.bilibili.com/search/all?keyword=${encodeURIComponent(r.name)}" target="_blank" rel="noopener">▶ 看做法</a>
          </div>
        </div>`;
      }).join('')}
    </div>
  </div>
  <div class="card card-soft">
    <div class="card-head"><span class="ch-emoji">🥘</span><h2>菜谱库</h2>
      <span class="ch-sub">${recipeFilter === 'fav' ? '我的收藏' : '按餐次挑选'}</span></div>
    <div class="filterbar" id="recipeFilter">
      ${[['all', '全部'], ['breakfast', '早餐'], ['lunch', '午餐'], ['dinner', '晚餐'], ['fav', '★ 收藏']].map(([v, l]) => `<button class="chip ${recipeFilter === v ? 'on' : ''}" data-act="recipe-filter" data-meal="${v}">${l}</button>`).join('')}
    </div>
    <div class="grid grid-3">
      ${libFilter.length ? libFilter.map(r => {
        const fav = state.recipes.favs.includes(r.id);
        const mealLabel = { breakfast: '早餐', lunch: '午餐', dinner: '晚餐' }[r.meal];
        return `<div class="card recipe-card">
          <div class="rc-head"><div class="rc-thumb">${r.emoji}</div>
            <div><h3 style="margin:0;font-size:15px">${r.name}</h3>
            <span style="font-size:12px;color:var(--ink-soft)">${mealLabel} · 约 ${r.time} 分钟</span></div></div>
          <ul class="recipe-steps">${r.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ul>
          <div class="recipe-tags">${r.tags.map(tg => `<span class="pill green">${esc(tg)}</span>`).join('')}</div>
          <div class="checkin-row">
            <button class="btn sm ${fav ? 'yellow' : 'ghost'}" data-act="recipe-fav" data-id="${r.id}">${fav ? '★ 已收藏' : '☆ 收藏'}</button>
            <a class="jump-link" href="https://www.bilibili.com/search/all?keyword=${encodeURIComponent(r.name)}" target="_blank" rel="noopener">▶ 看做法</a>
          </div>
        </div>`;
      }).join('') : '<div class="empty">这里还没有收藏，去上面点个 ☆ 吧～</div>'}
    </div>
  </div>`;
}

/* =================== 新鲜玩意 =================== */
const FRESH_CATS = ['全部', '穿搭', '餐厅', '游玩', '潮流', '理财投资', '副业', '娱乐', '收藏'];
let freshFilter = '全部';
let knowFilter = '全部';
let recipeFilter = 'all';
let quizSel = {};
function openQuiz() {
  quizSel = {};
  const qs = [
    { q: '这笔钱大概多久不会用到？', o: ['随时要用', '半年内', '1–3 年', '3 年以上'] },
    { q: '如果账户短期跌 20%，你会？', o: ['立刻卖出', '很焦虑', '还能忍', '无所谓当定投'] },
    { q: '你的主要目标是？', o: ['保本就好', '跑赢通胀', '稳健增值', '高收益敢冒险'] },
  ];
  openModal(`<h2>🧭 理财风险小测评</h2>
   <p style="color:var(--ink-soft);font-size:13px;margin:0 0 4px">3 题帮你判断适合的类型（仅供参考，非投资建议）</p>
   <div id="quizBox">${qs.map((it, i) => `<div class="quiz-q"><div class="qq">${i + 1}. ${it.q}</div>${it.o.map((o, j) => `<label class="quiz-opt" data-q="${i}" data-v="${j}">${o}</label>`).join('')}</div>`).join('')}</div>
   <div class="row"><button class="btn primary" data-act="quiz-calc">看结果</button></div>
   <div id="quizOut"></div>`);
}
function viewFresh(v) {
  const items = freshFilter === '收藏'
    ? state.fresh.items.filter(i => state.fresh.favs.includes(i.id))
    : state.fresh.items.filter(i => freshFilter === '全部' || i.cat === freshFilter);
  v.innerHTML = `
  <div class="card">
    <div class="card-head"><span class="ch-emoji">✨</span><h2>新鲜玩意</h2>
      <span class="ch-sub"><button class="btn sm ghost" data-act="fresh-add">＋ 添加灵感</button></span></div>
    <div class="filterbar" id="freshFilter">
      ${FRESH_CATS.map(c => `<button class="chip ${c === freshFilter ? 'on' : ''}" data-cat="${c}">${c}</button>`).join('')}
    </div>
    <div class="grid grid-3">
      ${items.length ? items.map(i => {
        const col = { '穿搭': 'var(--yellow-100)', '餐厅': '#FFE0E0', '游玩': '#E0EEFF', '潮流': 'var(--green-100)', '理财投资': '#E6F0FF', '副业': '#F3E8FF', '娱乐': '#FFF0D6' }[i.cat] || 'var(--green-100)';
        const fav = state.fresh.favs.includes(i.id);
        return `<div class="card card-soft fresh-card">
          <span class="fresh-cat" style="background:${col}">${esc(i.cat)}</span>
          <h3>${esc(i.title)}</h3><p>${esc(i.desc)}</p>
          ${i.hot ? `<span class="hot">${esc(i.hot)}</span>` : ''}
          <div class="checkin-row">
            <a class="jump-link" href="${i.link}" target="_blank" rel="noopener">▶ 看原平台</a>
            <button class="btn sm ${fav ? 'yellow' : 'ghost'}" data-act="fresh-fav" data-id="${i.id}">${fav ? '★' : '☆'}</button>
          </div>
        </div>`;
      }).join('') : '<div class="empty">这个分类还没有内容，点“添加灵感”记一条吧～</div>'}
    </div>
  </div>`;
}

/* =================== 自定义板块 =================== */
function viewCustom(v, it) {
  if (!state.custom[it.id]) state.custom[it.id] = { notes: '', icon: it.icon };
  const c = state.custom[it.id];
  v.innerHTML = `
  <div class="card">
    <div class="card-head"><span class="ch-emoji">${it.icon}</span><h2>${esc(it.name)}</h2>
      <span class="ch-sub">自定义板块</span></div>
    <p style="color:var(--ink-soft);font-size:13px;margin:0 0 8px">随手记点东西，数据同样本地保存：</p>
    <textarea class="field" id="customNotes" rows="8" placeholder="今天的灵感 / 待办 / 链接…">${esc(c.notes)}</textarea>
    <div style="margin-top:10px"><button class="btn primary sm" data-act="custom-save" data-id="${it.id}">保存</button></div>
  </div>`;
}

/* =================== 设置 / 数据 =================== */
function openSettings() {
  const s = state.settings.sync;
  openModal(`<h2>⚙ 设置 / 数据</h2>
    <div class="set-row"><span>导出全部数据（JSON）</span><button class="btn sm primary" data-act="export">下载</button></div>
    <div class="set-row"><span>导入数据（覆盖）</span><button class="btn sm" data-act="import">选择文件</button><input type="file" id="importFile" accept="application/json" style="display:none"></div>
    <div class="set-row"><span>恢复默认导航栏</span><button class="btn sm ghost" data-act="reset-nav">重置</button></div>
    <hr style="border:none;border-top:1px dashed var(--line);margin:14px 0">
    <div class="set-row" style="flex-direction:column;align-items:stretch;gap:8px">
      <div style="font-weight:800">🔄 多端实时同步</div>
      <p style="font-size:12.5px;color:var(--ink-soft);margin:0">想「不在家、电脑关机也能同步」就选 <b>云端 Supabase</b>（免费、邮箱注册、数据存在云端，两端随时一致）。在家且电脑常开也可用本地服务。</p>
      <label>同步方式
        <select id="syncMode" class="field">
          <option value="local" ${s.mode !== 'cloud' ? 'selected' : ''}>本地服务（仅在家）</option>
          <option value="cloud" ${s.mode === 'cloud' ? 'selected' : ''}>云端 Supabase（随时可用）</option>
        </select>
      </label>
      <div id="syncLocal">
        <label>服务地址 <input class="field" id="syncUrl" value="${esc(s.url)}" placeholder="http://电脑局域网IP:8787/workbench"></label>
        <label style="display:flex;gap:8px;align-items:center"><input type="checkbox" id="syncOn" ${s.enabled ? 'checked' : ''}> 启用自动同步</label>
      </div>
      <div id="syncCloud">
        <label>Supabase 项目 URL
          <input class="field" id="cloudUrl" value="${esc(s.cloudUrl || '')}" placeholder="https://xxxx.supabase.co（带 /rest/v1/ 也行）"></label>
        <label>anon key（公开密钥）
          <input class="field" id="cloudKey" type="password" value="${esc(s.cloudKey || '')}" placeholder="在 Supabase 项目 API 设置里复制"></label>
        <p style="font-size:12px;color:var(--ink-soft);margin:0">获取方法：① 打开 <b>supabase.com</b> 用邮箱注册 → 新建一个 Project；② 左侧 <b>SQL Editor</b> 粘贴运行我给你的建表语句；③ 左侧 <b>Project Settings → API</b> 复制 <b>Project URL</b> 与 <b>anon public key</b> 填上面。首次点「立即同步」会自动写入云端，不用手动建。</p>
      </div>
      <button class="btn sm yellow" data-act="sync-now">立即同步一次</button>
    </div>
    <hr style="border:none;border-top:1px dashed var(--line);margin:14px 0">
    <div class="set-row"><span>清空所有数据</span><button class="btn sm ghost" data-act="wipe">清空</button></div>
    <div class="set-row" style="font-size:11px;color:var(--ink-soft)"><span>当前版本</span><span style="font-family:monospace">v${APP_VERSION}</span></div>`);
  const modeSel = $('#syncMode');
  const applyMode = () => {
    const cloud = modeSel.value === 'cloud';
    $('#syncLocal').style.display = cloud ? 'none' : '';
    $('#syncCloud').style.display = cloud ? '' : 'none';
  };
  modeSel.addEventListener('change', applyMode);
  applyMode();
}

/* =================== 弹窗 =================== */
function openModal(html) {
  const root = $('#modalRoot');
  root.innerHTML = `<div class="modal-scrim" data-act="modal-close"></div><div class="modal">${html}<button class="close" data-act="modal-close">✕</button></div>`;
  root.classList.add('show');
}
function closeModal() { $('#modalRoot').classList.remove('show'); $('#modalRoot').innerHTML = ''; }

/* =================== 顶部连续打卡 =================== */
function updateStreakPill() {
  const p = $('#globalStreak');
  const s = enStreak();
  p.textContent = s > 0 ? `🔥 英语连打卡 ${s} 天` : '🌱 今日还没打卡';
}

/* =================== 事件：导航 / 顶栏 =================== */
$('#nav').addEventListener('click', e => {
  const rm = e.target.closest('[data-remove]');
  if (rm) {
    const id = rm.dataset.remove;
    if (state.sidebar.length <= 1) { toast('至少保留一个板块'); return; }
    state.sidebar = state.sidebar.filter(s => s.id !== id);
    delete state.custom[id];
    save(); renderSidebar();
    if (current === id) { current = state.sidebar[0].id; renderView(); }
    toast('已移除板块'); return;
  }
  const item = e.target.closest('.nav-item');
  if (item) go(item.dataset.id);
});
$('#hambBtn').addEventListener('click', () => { $('#sidebar').classList.add('open'); $('#scrim').classList.add('show'); });
$('#scrim').addEventListener('click', () => { $('#sidebar').classList.remove('open'); $('#scrim').classList.remove('show'); });
$('#settingsBtn').addEventListener('click', openSettings);
$('#addSectionBtn').addEventListener('click', () => {
  const emojis = ['📌', '🎯', '🌟', '📖', '🎨', '🏃', '🧘', '🎵', '🐱', '🌈'];
  openModal(`<h2>＋ 添加板块</h2>
    <div class="row"><label>名称</label><input class="field" id="newName" placeholder="如：读书 / 健身 / 旅行" style="flex:1"></div>
    <div class="row"><label>图标</label><div id="emojiPick" style="display:flex;gap:8px;flex-wrap:wrap">
      ${emojis.map((e, i) => `<button class="chip ${i === 0 ? 'on' : ''}" data-emoji="${e}">${e}</button>`).join('')}</div></div>
    <div class="row"><button class="btn primary" data-act="add-section-do">添加</button></div>`);
});

/* =================== 事件：主视图委托 =================== */
$('#view').addEventListener('click', e => {
  const el = e.target.closest('[data-act]'); if (!el) return;
  const act = el.dataset.act;
  if (act === 'work-tab') { state.work._tab = el.dataset.tab; save(); viewWork($('#view')); }
  else if (act === 'todo-add') {
    const input = $('#todoInput'); const text = input.value.trim();
    if (!text) { toast('先写点内容'); return; }
    const cat = ($('#catSeg .on') || {}).dataset?.cat || CATS[0];
    const sc = ($('#scopeSeg .on') || {}).dataset?.scope || SCOPES[0];
    const t = ($('#todoDate') && $('#todoDate').value) ? $('#todoDate').value.trim() : todayStr();
    todoDraftDate = t;
    (state.work.todos[t] = state.work.todos[t] || []).push({ id: uid(), text, cat, scope: sc, done: false, date: t });
    save(); input.value = ''; viewWork($('#view'));
    if (t !== todayStr()) toast('已记到 ' + t);
  }
  else if (act === 'todo-toggle') {
    const { id, date } = el.dataset;
    const item = (state.work.todos[date] || []).find(x => x.id === id);
    if (item) { item.done = !item.done; save(); viewWork($('#view')); }
  }
  else if (act === 'todo-del') {
    const { id, date } = el.dataset;
    state.work.todos[date] = (state.work.todos[date] || []).filter(x => x.id !== id);
    save(); viewWork($('#view'));
  }
  else if (act === 'todo-edit') {
    const { id, date } = el.dataset;
    const item = (state.work.todos[date] || []).find(x => x.id === id);
    if (!item) { toast('任务不存在'); return; }
    openModal(`<h2>✏️ 修改任务</h2>
      <div class="row"><label>内容</label><textarea class="field" id="editText" rows="2" style="flex:1">${esc(item.text)}</textarea></div>
      <div class="row"><label>工作性质</label><div class="seg" id="editCat">${CATS.map(c => `<button data-cat="${c}" class="cat-${c} ${c === item.cat ? 'on' : ''}">${c}</button>`).join('')}</div></div>
      <div class="row"><label>归属范围</label><div class="seg" id="editScope">${SCOPES.map(c => `<button data-scope="${c}" class="scope-${c} ${c === item.scope ? 'on' : ''}">${c}</button>`).join('')}</div></div>
      <div class="row"><label>日期</label><input type="date" class="field" id="editDate" value="${item.date}"></div>
      <div class="row"><button class="btn primary" data-act="todo-edit-do" data-id="${id}" data-date="${date}">保存</button></div>`);
  }
  else if (act === 'day-toggle') {
    const day = el.closest('.todo-day'); day.classList.toggle('collapsed');
  }
  /* ---- 小组目标管理事件 ---- */
  else if (act === 'team-tab') { state.team.activeTab = el.dataset.ttab; save(); renderTeam(); }
  else if (act === 'team-save-goals') {
    $$('.team-target-input').forEach(input => {
      const gid = input.dataset.gid, sgid = input.dataset.sgid, level = input.dataset.level;
      const val = input.value.trim().replace(/[^\d.]/g, '');
      if (level === 'group') { const g = state.team.groups.find(x => x.id === gid); if (g) g.target = val; }
      else if (level === 'subgroup') {
        const g = state.team.groups.find(x => x.id === gid);
        if (g) { const sg = g.subGroups.find(x => x.id === sgid); if (sg) sg.target = val; }
      }
    });
    save(); toast('目标已保存 ✅');
  }
  else if (act === 'team-add-member') {
    state.team._showAddForm = !state.team._showAddForm; save(); renderTeam();
  }
  else if (act === 'team-add-member-do-inline') {
    const name = ($('#newMName').value || '').trim();
    if (!name) { toast('请输入姓名'); return; }
    const mid = uid();
    const ntEl = $('#newMTarget');
    const nt = (ntEl ? (ntEl.value || '') : '').replace(/[^\d.]/g, '');
    state.team.members.push({ id: mid, name, personalTarget: nt, groupId: $('#newMGroup').value, subGroupId: $('#newMSGroup').value });
    // 目标同时写入当前查看的月份
    ensureMonth(curMonth()).targets[mid] = nt;
    const g = state.team.groups.find(x => x.id === $('#newMGroup').value);
    if (g) { const sg = g.subGroups.find(x => x.id === $('#newMSGroup').value); if (sg) sg.memberIds.push(mid); }
    state.team._showAddForm = false;
    save(); renderTeam(); toast('成员已添加 ✅');
  }
  else if (act === 'team-del-member') {
    const mid = el.dataset.mid;
    if (!confirm('确定删除该成员？')) return;
    state.team.members = state.team.members.filter(m => m.id !== mid);
    state.team.groups.forEach(g => g.subGroups.forEach(sg => sg.memberIds = sg.memberIds.filter(id => id !== mid)));
    save(); renderTeam(); toast('已删除');
  }
  else if (act === 'team-rm-member-sg') {
    const mid = el.dataset.mid, sgid = el.dataset.sgid;
    state.team.groups.forEach(g => g.subGroups.forEach(sg => { if (sg.id === sgid) sg.memberIds = sg.memberIds.filter(id => id !== mid); }));
    save(); renderTeam();
  }
  else if (act === 'team-save-member-targets') {
    const mk = el.dataset.mk || curMonth();
    const mo = ensureMonth(mk);
    let n = 0;
    $$('.tm-target-cell').forEach(input => {
      if (input.dataset.mk && input.dataset.mk !== mk) return; // 只存该月的格子
      const mid = input.dataset.mid;
      if (!state.team.members.find(x => x.id === mid)) return;
      mo.targets[mid] = (input.value || '').trim().replace(/[^\d.]/g, '');
      n++;
    });
    save(); renderTeam();
    toast(`✅ 已保存 ${monthLabel(mk)} 的 ${n} 位成员目标`);
  }
  else if (act === 'team-month-toggle') {
    const mk = el.dataset.mk;
    const mo = ensureMonth(mk);
    mo.collapsed = !mo.collapsed;
    save(); renderTeam();
  }
  else if (act === 'team-month-del') {
    const mk = el.dataset.mk;
    if (!confirm(`确定删除「${monthLabel(mk)}」的目标记录？\n该月已填写的目标将被清空（成员本身不会删除）。`)) return;
    delete state.team.months[mk];
    if (state.team.currentMonth === mk) state.team.currentMonth = monthKey();
    save(); renderTeam(); toast(`已删除 ${monthLabel(mk)} 的记录`);
  }
  else if (act === 'team-month-goto') {
    switchMonth(el.dataset.mk);
  }
  else if (act === 'team-month-new') {
    openMonthPicker();
  }
  else if (act === 'team-month-new-do') {
    const sel = $('#mpMonth');
    if (!sel) return;
    closeModal();
    switchMonth(sel.value);
    toast(`已切换到 ${monthLabel(sel.value)}`);
  }
  else if (act === 'team-edit-member') {
    state.team._editMemberId = el.dataset.mid; save();
    renderTeam();
  }
  else if (act === 'team-edit-member-cancel') {
    state.team._editMemberId = ''; save();
    renderTeam();
  }
  else if (act === 'team-edit-member-do') {
    const mid = el.dataset.mid;
    const m = state.team.members.find(x => x.id === mid);
    if (!m) { toast('❌ 成员不存在'); return; }
    const nameEl = $('#editMName');
    const name = nameEl ? (nameEl.value || '').trim() : m.name;
    if (!name) { toast('姓名不能为空'); return; }
    const targetInput = $('#editMTarget');
    let newTarget = targetInput ? (targetInput.value || '').trim() : '';
    // 容错：用户可能输入 "95%" "95 %"，去掉非数字字符，只保留数字和小数点
    newTarget = newTarget.replace(/[^\d.]/g, '');
    // 先落盘关键字段（姓名/个人目标），确保后面任何异常都不会把目标丢掉
    m.name = name;
    m.personalTarget = newTarget;
    // 大组/小小组：全部做空值保护，取不到就保持原值
    try {
      const gEl = $('#editMGroup'), sgEl = $('#editMSGroup');
      const newGid = gEl ? gEl.value : m.groupId;
      const newSgid = sgEl ? sgEl.value : m.subGroupId;
      if (newGid !== m.groupId || newSgid !== m.subGroupId) {
        state.team.groups.forEach(g => g.subGroups.forEach(sg => { sg.memberIds = (sg.memberIds || []).filter(id => id !== mid); }));
        const g = state.team.groups.find(gr => gr.id === newGid);
        if (g) { const sg = g.subGroups.find(s => s.id === newSgid); if (sg) { sg.memberIds = sg.memberIds || []; if (!sg.memberIds.includes(mid)) sg.memberIds.push(mid); } }
      }
      m.groupId = newGid;
      m.subGroupId = newSgid;
    } catch (err) {
      console.warn('team-edit-member-do 分组同步失败（目标已保存）:', err);
    }
    state.team._editMemberId = ''; save();
    renderTeam();
    toast('✅ 已保存 ' + esc(m.name) + ' 目标: ' + (newTarget || '(空)') + '%');
  }
  else if (act === 'team-goto-members-month') {
    const mk = el.dataset.mk || curMonth();
    ensureMonth(mk);
    state.team.currentMonth = mk;
    state.team.activeTab = 'members';
    save(); renderTeam();
    toast(`已切到「成员管理」的 ${monthLabel(mk)}，填完点「保存 ${monthLabel(mk)} 目标」`);
  }
  else if (act === 'team-create-period') {
    const sEl = $('#wkNewStart'), eEl = $('#wkNewEnd');
    const s = (sEl ? sEl.value : '') || mondayStr();
    const e = (eEl ? eEl.value : '') || sundayStr(s);
    if (!s) { toast('请先选择起始日期'); return; }
    if (e < s) { toast('❌ 结束日期不能早于起始日期'); return; }
    if (state.team.weeks[s]) { toast('该起始日期的区间已存在'); state.team._selectedPeriod = s; save(); renderTeam(); return; }
    state.team.weeks[s] = { start: s, end: e, data: {} };
    state.team._selectedPeriod = s; save();
    renderTeam(); toast(`✅ 已创建区间 ${s} ~ ${e}`);
  }
  else if (act === 'team-del-period') {
    const k = el.dataset.k;
    if (!state.team.weeks[k]) return;
    if (!confirm(`确定删除区间「${weekLabel(state.team.weeks[k])}」？\n该区间已录入的数据将一并删除。`)) return;
    delete state.team.weeks[k];
    if (state.team._selectedPeriod === k) state.team._selectedPeriod = periodList()[0] || '';
    save(); renderTeam(); toast('🗑 已删除该区间');
  }
  else if (act === 'wk-toggle-more') {
    const mid = el.dataset.mid;
    const row = $(`[data-detail="${mid}"]`);
    if (row) {
      const open = row.style.display !== 'none';
      row.style.display = open ? 'none' : '';
      el.classList.toggle('on', !open);
      if (!state.team._openDetails) state.team._openDetails = {};
      state.team._openDetails[mid] = !open;
      save();
    }
  }
  else if (act === 'team-save-period') {
    const k = el.dataset.k;
    const wk = state.team.weeks[k];
    if (!wk) { toast('区间不存在'); return; }
    wk.data = wk.data || {};
    const mk = monthOfPeriod(wk) || curMonth();
    // 每周项：完成率 / 严错数（自动计分的两个输入）
    $$('.wk-rate').forEach(i => {
      const mid = i.dataset.mid; if (!mid) return;
      wk.data[mid] = wk.data[mid] || {};
      wk.data[mid].completionRate = (i.value || '').trim().replace(/[^\d.]/g, '');
    });
    $$('.wk-err').forEach(i => {
      const mid = i.dataset.mid; if (!mid) return;
      wk.data[mid] = wk.data[mid] || {};
      wk.data[mid].seriousErrors = (i.value || '').trim().replace(/[^\d.]/g, '');
    });
    // 月度项
    $$('.wk-mx, .wk-mxc, .wk-mxs').forEach(i => {
      const mid = i.dataset.mid, fld = i.dataset.fld;
      if (!mid || !fld) return;
      const x = monthExtras(mk, mid);
      x[fld] = i.type === 'checkbox' ? i.checked : (i.value || '').trim();
    });
    save(); renderTeam(); toast(`✅ ${weekLabel(wk)} 数据已保存`);
  }
  else if (act === 'team-show-rules') { state.team.activeTab = 'rules'; save(); renderTeam(); }
  else if (act === 'report-open') openReport();
  else if (['report-gen', 'report-gen-custom', 'report-copy', 'report-export', 'report-export-docx'].includes(act)) handleReport(act, el);
  else if (act === 'quiz-open') { openQuiz(); }
  else if (act === 'know-filter') { knowFilter = el.dataset.tag; viewFinance($('#view')); }
  else if (act === 'recipe-filter') { recipeFilter = el.dataset.meal; viewRecipes($('#view')); }
  else if (act === 'en-check') {
    const t = todayStr(); const c = state.english.checkins[t] || {};
    c[el.dataset.type] = !c[el.dataset.type];
    state.english.checkins[t] = c; save(); viewEnglish($('#view'));
  }
  else if (act === 'en-play') {
    speak(decodeURIComponent(el.dataset.text || ''));
  }
  else if (act === 'en-sub') {
    const t = todayStr(); state.english.subs[t] = state.english.subs[t] || {};
    const sub = state.english.subs[t];
    const ty = el.dataset.type, idx = el.dataset.idx;
    sub[ty] = sub[ty] || {};
    sub[ty][idx] = !sub[ty][idx];
    save(); viewEnglish($('#view'));
  }
  else if (act === 'en-vocab') {
    const id = el.dataset.id; const m = state.english.vocabMastered;
    m.includes(id) ? m.splice(m.indexOf(id), 1) : m.push(id);
    save(); viewEnglish($('#view'));
  }
  else if (act === 'recipe-fav') {
    const id = el.dataset.id; const f = state.recipes.favs;
    f.includes(id) ? f.splice(f.indexOf(id), 1) : f.push(id);
    save(); viewRecipes($('#view'));
  }
  else if (act === 'recipe-shuffle') {
    const t = todayStr(); delete state.recipes.dailySeed[t]; viewRecipes($('#view'));
  }
  else if (act === 'fresh-fav') {
    const id = el.dataset.id; const f = state.fresh.favs;
    f.includes(id) ? f.splice(f.indexOf(id), 1) : f.push(id);
    save(); viewFresh($('#view'));
  }
  else if (act === 'know-add') {
    openModal(`<h2>＋ 记一条理财知识</h2>
      <div class="row"><label>标签</label><input class="field" id="kTag" placeholder="如：入门 / 工具 / 风险" style="flex:1"></div>
      <div class="row"><label>标题</label><input class="field" id="kTitle" placeholder="如：什么是定投" style="flex:1"></div>
      <div class="row"><label>内容</label><textarea class="field" id="kBody" rows="3" placeholder="一句话讲清楚"></textarea></div>
      <div class="row"><button class="btn primary" data-act="know-save">保存</button></div>`);
  }
  else if (act === 'know-save') {
    state.finance.knowledge.unshift({ id: uid(), tag: $('#kTag').value.trim() || '笔记', title: $('#kTitle').value.trim() || '未命名', body: $('#kBody').value.trim() });
    save(); closeModal(); viewFinance($('#view')); toast('已保存');
  }
  else if (act === 'fresh-add') {
    openModal(`<h2>＋ 添加灵感</h2>
      <div class="row"><label>分类</label><select class="field" id="fCat">${FRESH_CATS.filter(c => c !== '全部').map(c => `<option>${c}</option>`).join('')}</select></div>
      <div class="row"><label>标题</label><input class="field" id="fTitle" style="flex:1"></div>
      <div class="row"><label>描述</label><textarea class="field" id="fDesc" rows="2"></textarea></div>
      <div class="row"><label>原平台链接</label><input class="field" id="fLink" placeholder="https://..." style="flex:1"></div>
      <div class="row"><button class="btn primary" data-act="fresh-save">保存</button></div>`);
  }
  else if (act === 'fresh-save') {
    state.fresh.items.unshift({ id: uid(), cat: $('#fCat').value, title: $('#fTitle').value.trim() || '未命名', desc: $('#fDesc').value.trim(), hot: '', link: $('#fLink').value.trim() || '#' });
    save(); closeModal(); viewFresh($('#view')); toast('已添加');
  }
  else if (act === 'custom-save') {
    state.custom[current].notes = $('#customNotes').value; save(); toast('已保存');
  }
});

/* 导航栏分段选择（性质 / 范围） */
$('#view').addEventListener('click', e => {
  const seg = e.target.closest('#catSeg button, #scopeSeg button');
  if (seg) { $$('.seg').forEach(g => { if (g.contains(seg)) $$('button', g).forEach(b => b.classList.remove('on')); }); seg.classList.add('on'); }
});
/* 新鲜玩意筛选 */
$('#view').addEventListener('click', e => {
  const chip = e.target.closest('#freshFilter .chip');
  if (chip) { freshFilter = chip.dataset.cat; viewFresh($('#view')); }
});
/* 小组管理：日期区间选择 / 月选择 change */
$('#view').addEventListener('change', e => {
  if (e.target.id === 'wkPeriodSel') {
    state.team._selectedPeriod = e.target.value || ''; save();
    renderWeeklyForm(e.target.value);
    return;
  }
  // 月度项变动 → 刷新该行小计 + 顺手落库
  if (e.target.classList && (e.target.classList.contains('wk-mx') || e.target.classList.contains('wk-mxc') || e.target.classList.contains('wk-mxs'))) {
    recalcMonthExtraUI(e.target.dataset.mid, e.target.dataset.mk);
    commitMonthExtra(e.target.dataset.mid, e.target.dataset.mk, e.target.dataset.fld);
    return;
  }
  // 周面板改「个人目标」→ 同步回成员管理
  if (e.target.classList && e.target.classList.contains('wk-tgt-in')) {
    commitWeekTarget(e.target);
    return;
  }
  // 完成率 / 严错数 失焦即自动落库（避免忘记点保存导致丢数据）
  if (e.target.classList && (e.target.classList.contains('wk-rate') || e.target.classList.contains('wk-err'))) {
    commitWeekInput(e.target);
    return;
  }
  if (e.target.id === 'sbMonthSelect') {
    state.team._sbMonth = e.target.value; save();
    renderSBContent(e.target.value);
  }
  // 成员管理：切换大组时联动小小组
  if (e.target.classList.contains('team-sg-select')) {
    const mid = e.target.dataset.mid;
    const newGid = e.target.closest('.team-m-group').querySelector('[data-fld="groupId"]').value;
    const g = state.team.groups.find(gr => gr.id === newGid);
    if (g) { e.target.innerHTML = g.subGroups.map(sg => `<option value="${sg.id}" ${sg.id===e.target.value?'selected':''}>${esc(sg.name)}</option>`).join(''); }
  }
});
/* 周数据录入：完成率 / 严错数 输入即时重算得分 */
$('#view').addEventListener('input', e => {
  const t = e.target;
  if (!t || !t.classList) return;
  if (t.classList.contains('wk-rate') || t.classList.contains('wk-err')) recalcWeeklyUI();
  else if (t.classList.contains('wk-mx') || t.classList.contains('wk-mxs')) recalcMonthExtraUI(t.dataset.mid, t.dataset.mk);
});

/* =================== 弹窗事件委托 =================== */
$('#modalRoot').addEventListener('click', e => {
  const opt = e.target.closest('.quiz-opt');
  if (opt) {
    const q = opt.dataset.q;
    $$('#quizBox .quiz-opt').forEach(o => { if (o.dataset.q === q) o.classList.remove('on'); });
    opt.classList.add('on'); quizSel[q] = Number(opt.dataset.v); return;
  }
  const segBtn = e.target.closest('.seg button');
  if (segBtn) { const g = segBtn.closest('.seg'); $$('button', g).forEach(b => b.classList.remove('on')); segBtn.classList.add('on'); return; }
  const el = e.target.closest('[data-act]'); if (!el) return;
  const act = el.dataset.act;
  if (act === 'quiz-calc') {
    const vals = Object.values(quizSel);
    if (vals.length < 3) { toast('请完成全部 3 题'); return; }
    const score = vals.reduce((a, b) => a + b, 0);
    const map = [
      [2, '保守型', '优先货币基金、国债逆回购、定期存款，先把“安全垫”铺好。'],
      [4, '稳健型', '以债券基金 / “固收+”为主，少量指数基金定投。'],
      [6, '平衡型', '股债均衡配置，坚持指数基金定投，控制单只仓位。'],
      [8, '进取型', '可配置股票 / 行业 ETF，但务必控制仓位、设止损。'],
    ];
    const pick = map.find(m => score <= m[0]) || map[map.length - 1];
    $('#quizOut').innerHTML = `<div class="report-summary" style="margin-top:12px">
      <div class="stat"><div class="num">${pick[1]}</div><div class="lab">你的类型（得分 ${score}）</div></div>
      <p style="margin:10px 0 0;font-size:13.5px">${pick[2]}</p>
      <p style="margin:8px 0 0;font-size:12px;color:var(--ink-soft)">提示：投资有风险，先用小额试水，别一次 all in。</p>
    </div>`;
    return;
  }
  if (act === 'modal-close') { closeModal(); return; }
  if (act === 'add-section-do') {
    const name = ($('#newName').value || '').trim();
    if (!name) { toast('请填写名称'); return; }
    const icon = ($('#emojiPick .on') || {}).dataset?.emoji || '📌';
    const id = 'custom_' + uid();
    state.sidebar.push({ id, name, icon, core: false });
    state.custom[id] = { notes: '', icon };
    save(); closeModal(); renderSidebar(); toast('已添加板块');
  }
  if (act === 'export') {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = `workbench-backup-${todayStr()}.json`; a.click(); toast('已导出');
  }
  if (act === 'import') { $('#importFile').click(); }
  if (act === 'reset-nav') {
    state.sidebar = seed().sidebar; save(); renderSidebar(); renderView(); closeModal(); toast('导航栏已重置');
  }
  if (act === 'wipe') {
    if (confirm('确定清空所有数据？此操作不可恢复。')) { localStorage.removeItem(KEY); state = seed(); save(); closeModal(); renderSidebar(); renderView(); toast('已清空'); }
  }
  if (act === 'sync-now') {
    const s = state.settings.sync;
    s.mode = $('#syncMode').value;
    if (s.mode === 'cloud') {
      s.cloudUrl = ($('#cloudUrl').value || '').trim().replace(/\/+$/, '');
      s.cloudKey = ($('#cloudKey').value || '').trim();
      if (!s.cloudUrl || !s.cloudKey) { toast('请先填写 Supabase 项目 URL 和 anon key'); return; }
    } else {
      s.url = ($('#syncUrl').value || '').trim();
      s.enabled = $('#syncOn').checked;
      if (!s.url) { toast('请先填写同步服务地址'); return; }
    }
    save(true);
    syncNow();
  }
  if (act === 'todo-edit-do') {
    const id = el.dataset.id, odate = el.dataset.date;
    const text = ($('#editText').value || '').trim();
    if (!text) { toast('内容不能为空'); return; }
    const cat = ($('#editCat .on') || {}).dataset?.cat || CATS[0];
    const sc = ($('#editScope .on') || {}).dataset?.scope || SCOPES[0];
    const ndate = ($('#editDate').value || odate).trim();
    const arr = state.work.todos[odate] || [];
    const idx = arr.findIndex(x => x.id === id);
    if (idx < 0) { toast('任务不存在'); return; }
    const item = arr[idx];
    item.text = text; item.cat = cat; item.scope = sc;
    if (ndate !== odate) { item.date = ndate; arr.splice(idx, 1); (state.work.todos[ndate] = state.work.todos[ndate] || []).push(item); }
    save(); closeModal(); viewWork($('#view')); toast('已保存');
  }
  if (act === 'team-month-new-do') {
    const sel = $('#mpMonth');
    if (!sel) return;
    closeModal();
    switchMonth(sel.value);
    toast(`已切换到 ${monthLabel(sel.value)}`);
    return;
  }
  if (act === 'team-add-member-do') {
    const name = ($('#newMName').value || '').trim();
    if (!name) { toast('请输入姓名'); return; }
    const mid = uid();
    const ntEl = $('#newMTarget');
    const nt = (ntEl ? (ntEl.value || '') : '').replace(/[^\d.]/g, '');
    state.team.members.push({ id: mid, name, personalTarget: nt, groupId: $('#newMGroup').value, subGroupId: $('#newMSGroup').value });
    const g = state.team.groups.find(x => x.id === $('#newMGroup').value);
    if (g) { const sg = g.subGroups.find(x => x.id === $('#newMSGroup').value); if (sg) sg.memberIds.push(mid); }
    save(); closeModal(); state.team.activeTab = 'members'; renderTeam(); toast('成员已添加 ✅');
    return;
  }
  handleReport(act, el);
});
$('#modalRoot').addEventListener('click', e => {
  const em = e.target.closest('#emojiPick .chip');
  if (em) { $$('#emojiPick .chip').forEach(c => c.classList.remove('on')); em.classList.add('on'); }
});
$('#importFile') && $('#importFile').addEventListener('change', e => {
  const f = e.target.files[0]; if (!f) return;
  const r = new FileReader();
  r.onload = () => { try { state = JSON.parse(r.result); save(); closeModal(); renderSidebar(); renderView(); toast('导入成功'); } catch (err) { toast('文件格式错误'); } };
  r.readAsText(f);
});

/* =================== 同步（自动 + 合并） =================== */
// 深合并：数组按 id 取并集，对象递归合并，原始值以远端为准
function deepMerge(base, inc) {
  if (Array.isArray(base) && Array.isArray(inc)) {
    if (base.length && base[0] && typeof base[0] === 'object' && 'id' in base[0]) {
      const m = new Map(); [...base, ...inc].forEach(x => m.set(x.id, x)); return [...m.values()];
    }
    return Array.from(new Set([...base, ...inc].map(x => JSON.stringify(x)))).map(x => JSON.parse(x));
  }
  if (base && typeof base === 'object' && inc && typeof inc === 'object' && !Array.isArray(base)) {
    const out = { ...base };
    for (const k of Object.keys(inc)) out[k] = (k in out) ? deepMerge(out[k], inc[k]) : inc[k];
    return out;
  }
  return inc === undefined ? base : inc;
}

/* ============ 同步：本地服务 / 云端 Supabase 通用 ============ */
function syncActive() {
  const s = state.settings.sync;
  return s.mode === 'cloud' ? !!(s.cloudUrl && s.cloudKey) : !!(s.enabled && s.url);
}
function cloudBase() {
  // Supabase API 页面给的 URL 可能带 /rest/v1/，代码里还会再拼 /rest/v1/，
  // 所以这里统一去掉尾部 /rest/v1* 和斜杠，只保留基础域名
  return (state.settings.sync.cloudUrl || '').replace(/\/rest\/v\d*\/?$/, '').replace(/\/+$/, '');
}
async function pullRemote() {
  const s = state.settings.sync;
  if (s.mode === 'cloud') {
    const r = await fetch(`${cloudBase()}/rest/v1/sync?id=eq.workbench&select=data`, {
      headers: { 'apikey': s.cloudKey, 'Authorization': `Bearer ${s.cloudKey}` },
    });
    if (!r.ok) { const t = await r.text().catch(() => ''); throw new Error(`拉取失败 HTTP ${r.status}: ${t.slice(0, 200)}`); }
    const j = await r.json();
    return (j && j[0] && j[0].data) ? j[0].data : null;
  }
  const r = await fetch(s.url, { method: 'GET' });
  if (!r.ok) throw new Error('拉取失败 HTTP ' + r.status);
  const j = await r.json();
  return (j && j.state) ? j.state : null;
}
async function pushRemote(st) {
  const s = state.settings.sync;
  if (s.mode === 'cloud') {
    // 首次/后续都走 upsert：用固定 id 'workbench'，两端读写同一行即共享数据
    const r = await fetch(`${cloudBase()}/rest/v1/sync`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': s.cloudKey,
        'Authorization': `Bearer ${s.cloudKey}`,
        'Prefer': 'return=representation, resolution=merge-duplicates',
      },
      body: JSON.stringify({ id: 'workbench', data: st, updated_at: new Date().toISOString() }),
    });
    if (r.status === 401 || r.status === 403) { const t = await r.text().catch(() => ''); throw new Error(`认证失败 HTTP ${r.status}（可能用了 service_role key？请用 anon key）。详情: ${t.slice(0, 150)}`); }
    if (!r.ok) { const t = await r.text().catch(() => ''); throw new Error(`推送失败 HTTP ${r.status}: ${t.slice(0, 200)}`); }
  } else {
    await fetch(s.url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ updatedAt: Date.now(), state: st }) });
  }
}
function schedulePush() {
  if (!syncActive()) return;
  clearTimeout(_pushTimer);
  _pushTimer = setTimeout(doPush, 2500);
}
async function doPush() {
  if (_syncing || !syncActive()) return;
  const now = Date.now();
  if (now - _lastPush < 15000) { setTimeout(doPush, 15000 - (now - _lastPush)); return; } // 限流，避免触发免费额度
  _syncing = true;
  try { await pushRemote(state); _lastPush = Date.now(); } catch (e) {} finally { _syncing = false; }
}
async function doPull() {
  if (_syncing || !syncActive()) return;
  _syncing = true;
  try {
    const remote = await pullRemote();
    if (remote) { state = deepMerge(state, remote); save(true); renderView(); }
  } catch (e) {} finally { _syncing = false; }
}
async function syncNow() {
  if (!syncActive()) { toast('请先在同步设置里填好地址或云端密钥'); return; }
  if (_syncing) return;
  _syncing = true;
  try {
    const remote = await pullRemote();
    if (remote) state = deepMerge(state, remote);
    await pushRemote(state);
    _lastPush = Date.now();
    save(true);
    toast('同步成功 ✅');
    renderView();
  } catch (e) {
    // 显示完整错误详情，方便排查
    const msg = e.message || String(e);
    console.error('[同步错误]', e);
    toast('同步失败：' + msg);
  }
  _syncing = false;
}

/* =================== 每日 7 点自动建空白 To Do =================== */
function autoDaily() {
  const t = todayStr();
  if (!state.work.todos[t]) {
    ensureToday();
    toast('🌅 已为你准备好今天的空白 To Do 清单');
    if (current === 'work') viewWork($('#view'));
  }
}
// 打开时建；之后每分钟检查是否跨日（真实 7 点定时需在 PWA 后台周期同步，见 README）
ensureToday();
setInterval(autoDaily, 60 * 1000);

/* =================== 启动 =================== */
function boot() {
  const hash = location.hash.replace('#', '');
  if (hash && state.sidebar.find(s => s.id === hash)) current = hash;
  renderSidebar(); renderView();
  // 同步：打开即拉取一次；切回页面/打开时再拉；每 5 分钟自动拉取；改动自动推送（限流）
  if (syncActive()) { doPull(); }
  document.addEventListener('visibilitychange', () => { if (!document.hidden && syncActive()) doPull(); });
  setInterval(() => { if (syncActive()) doPull(); }, 5 * 60 * 1000);
  // 注销 Service Worker：避免 SW 缓存 index/app.js 导致「改了不生效」
  // 个人工具不需要离线，换来每次刷新都拉最新代码
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.getRegistrations().then(regs => regs.forEach(r => r.unregister())).catch(() => {});
  }
}
boot();
