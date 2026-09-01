/* =========================================================
   小煎蛋的工作台  —  个人工作台 (本地优先 / PWA)
   数据全部存于浏览器 localStorage；可选接入同步服务实现多端实时同步
   ========================================================= */
'use strict';

/* ---------- 基础工具 ---------- */
const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const KEY = 'xiaojidan_workbench_v1';

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
      weeks: {}
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
function renderTopbar(title) {
  $('#pageTitle').textContent = title;
  const d = new Date();
  $('#pageDate').textContent = `${greet()} · ${todayStr()} ${WEEK[d.getDay()]}`;
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
function viewTeamGoals(v) {
  const tab = state.team.activeTab || 'goals';
  v.innerHTML = `
  <div class="team-tabs" style="margin-bottom:16px">
    <button class="team-tab ${tab==='goals'?'on':''}" data-act="team-tab" data-ttab="goals">🎯 目标设置</button>
    <button class="team-tab ${tab==='members'?'on':''}" data-act="team-tab" data-ttab="members">👥 成员管理</button>
    <button class="team-tab ${tab==='weekly'?'on':''}" data-act="team-tab" data-ttab="weekly">📝 周数据录入</button>
    <button class="team-tab ${tab==='scoreboard'?'on':''}" data-act="team-tab" data-ttab="scoreboard">🏆 积分看板</button>
  </div>
  <div id="teamPanel"></div>`;
  const p = $('#teamPanel');
  if (tab === 'goals') renderTeamGoals(p);
  else if (tab === 'members') renderTeamMembers(p);
  else if (tab === 'weekly') renderTeamWeekly(p);
  else renderTeamScoreboard(p);
}

/* ---- 目标设置 ---- */
function renderTeamGoals(v) {
  const T = state.team;
  v.innerHTML = `<div class="card"><div class="card-head"><h2>🎯 小组目标设置</h2><span class="ch-sub">填写各层级目标完成率（%）</span></div>
  ${T.groups.map(g => `
    <div class="team-group-card" style="margin-bottom:20px">
      <div class="team-g-name">${esc(g.name)} <span class="team-g-target-label">大组目标</span>
        <input type="number" class="field team-target-input" data-gid="${g.id}" data-level="group" value="${g.target||''}" placeholder="%" style="width:80px"> %
      </div>
      <div class="team-subgroups">
        ${g.subGroups.map(sg => `
          <div class="team-sg-row">
            <span class="team-sg-name">${esc(sg.name)}</span>
            <span class="team-sg-target-label">小小组目标</span>
            <input type="number" class="field team-target-input" data-gid="${g.id}" data-sgid="${sg.id}" data-level="subgroup" value="${sg.target||''}" placeholder="%" style="width:80px"> %
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

/* ---- 成员管理 ---- */
function renderTeamMembers(v) {
  const T = state.team;
  v.innerHTML = `<div class="card"><div class="card-head"><h2>👥 成员管理（共${T.members.length}人）</h2><span class="ch-sub">共11人：>2年组5人 / 1-2年组6人</span></div>
  <div style="margin-bottom:12px"><button class="btn primary sm" data-act="team-add-member">＋ 添加成员</button></div>
  <div class="team-member-list">
    ${T.members.length === 0 ? '<div class="empty">暂无成员，点击上方按钮添加</div>' :
      T.members.map(m => {
        const g = T.groups.find(gr => gr.id === m.groupId);
        const sg = g ? g.subGroups.find(s => s.id === m.subGroupId) : null;
        return `<div class="team-member-row">
          <span class="team-m-name">${esc(m.name)}</span>
          <span class="team-m-target">个人目标: <input type="number" class="field" data-mid="${m.id}" data-fld="personalTarget" value="${m.personalTarget||''}" placeholder="%" style="width:70px"> %</span>
          <span class="team-m-group">
            所属: <select class="field" data-mid="${m.id}" data-fld="groupId" style="width:110px">
              ${T.groups.map(gg => `<option value="${gg.id}" ${gg.id===m.groupId?'selected':''}>${esc(gg.name)}</option>`).join('')}
            </select>
            <select class="field team-sg-select" data-mid="${m.id}" data-fld="subGroupId" style="width:110px">
              ${((g||T.groups[0]).subGroups||[]).map(ss => `<option value="${ss.id}" ${ss.id===m.subGroupId?'selected':''}>${esc(ss.name)}</option>`).join('')}
            </select>
          </span>
          <button class="btn sm ghost" data-act="team-del-member" data-mid="${m.id}" style="color:var(--danger)">删除</button>
        </div>`;
      }).join('')
    }
  </div>
  <div style="margin-top:12px"><button class="btn primary" data-act="team-save-members">💾 保存成员</button></div>
  </div>`;
}

/* ---- 周数据录入 ---- */
function getWeekKey(dateStr) { const d = new Date(dateStr); const jan1 = new Date(d.getFullYear(),0,1); const days = Math.floor((d - jan1)/(86400000)); const w = Math.ceil((days + jan1.getDay()+1)/7); return `${d.getFullYear()}-W${String(w).padStart(2,'0')}`; }
function getWeeksInMonth(year, month) { const weeks = []; Object.keys(state.team.weeks || {}).forEach(k => { const m = k.match(/^(\d{4})-W(\d{2})$/); if (!m) return; const d = new Date(year, month, 1); const jan1 = new Date(year,0,1); const targetWeek = Math.ceil((1 + jan1.getDay())/7); const wNum = parseInt(m[2]); /* simple: include if year matches and week is in range */ if (parseInt(m[1]) === year) weeks.push(k); }); return weeks.sort().reverse(); }

function renderTeamWeekly(v) {
  const T = state.team;
  const now = new Date();
  const curWeek = getWeekKey(todayStr());
  const weekKeys = Object.keys(T.weeks || {}).sort().reverse();
  const selectedWeek = T._selectedWeek || curWeek;

  v.innerHTML = `<div class="card"><div class="card-head"><h2>📝 周数据录入</h2>
    <span class="ch-sub">
      选择周: <select class="field" id="weekSelect" style="width:160px">
        <option value="">+ 新建一周</option>
        ${weekKeys.map(wk => `<option value="${wk}" ${wk===selectedWeek?'selected':''}>${wk}${wk===curWeek?' (本周)':''}</option>`).join('')}
      </select>
      ${selectedWeek && !T.weeks[selectedWeek] ? '' : selectedWeek ? `<span style="margin-left:8px;font-size:12px;color:var(--ink-soft)">日期: ${(T.weeks[selectedWeek]||{}).weekDate||''}</span>` : ''}
    </span>
  </div>
  <div id="weeklyForm"></div>
  </div>`;

  renderWeeklyForm(selectedWeek);
}
function renderWeeklyForm(weekKey) {
  const T = state.team;
  const form = $('#weeklyForm');
  if (!weekKey || !T.weeks[weekKey]) {
    form.innerHTML = `<div style="padding:20px;text-align:center">
      <p style="color:var(--ink-soft);margin-bottom:12px">选择已有周或新建一周来录入数据</p>
      <div class="row"><label>该周起始日期</label><input type="date" class="field" id="newWeekDate" value="${todayStr()}"></div>
      <div class="row"><button class="btn primary" data-act="team-create-week">创建新周</button></div>
    </div>`;
    return;
  }
  const wk = T.weeks[weekKey];
  const members = T.members;

  if (members.length === 0) { form.innerHTML = '<div class="empty">请先在「成员管理」中添加成员</div>'; return; }

  form.innerHTML = `
  <div style="overflow-x:auto">
  <table class="team-table">
    <thead><tr>
      <th>成员</th><th>所属小组</th><th>完成率%</th><th>严错数</th><th>无差错周</th>
      <th>个人达标</th><th>小小组达标</th><th>应出勤排名</th>
      <th>pending/误稿</th><th>Low违规</th><th>Med违规</th><th>High违规</th>
      <th>活动(×3)</th><th>主动补位(×2)</th><th>线上分享(×1)</th>
      <th>大事件(×2)</th><th>茶水(×1)</th><th>答题(×1)</th><th>讲师(×3)</th><th>其他推动</th>
      <th>触及质量条例</th><th>拖累小组</th><th>严错≥3</th><th>全员不达标</th>
      <th>周积分</th>
    </tr></thead>
    <tbody>
      ${members.map(m => {
        const d = (wk.data || {})[m.id] || {};
        const score = calcMemberWeekScore(m, d, wk, weekKey);
        const g = T.groups.find(gr => gr.id === m.groupId);
        const sg = g ? g.subGroups.find(s => s.id === m.subGroupId) : null;
        return `<tr>
          <td><b>${esc(m.name)}</b></td>
          <td><span class="pill scope-${sg?((['组内','科室','部门','外部'])[T.groups.indexOf(g)]||'其他'):'其他'}" style="font-size:11px">${sg?sg.name:'-'}</span></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="completionRate" value="${d.completionRate!==undefined?d.completionRate:''}" style="width:60px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="seriousErrors" value="${d.seriousErrors||0}" style="width:50px"></td>
          <td><input type="checkbox" class="team-cell" data-mid="${m.id}" data-fld="bonusNoErrorWeek" ${d.bonusNoErrorWeek?'checked':''}></td>
          <td><input type="checkbox" class="team-cell" data-mid="${m.id}" data-fld="bonusPersonalGoal" ${d.bonusPersonalGoal?'checked':''}></td>
          <td><input type="checkbox" class="team-cell" data-mid="${m.id}" data-fld="bonusSubGroupGoal" ${d.bonusSubGroupGoal?'checked':''}></td>
          <td><select class="field team-cell" data-mid="${m.id}" data-fld="attendanceRank" style="width:65px">
            <option value="" ${!d.attendanceRank?'selected':''}>-</option>
            ${[1,2,3].map(n=>`<option value="${n}" ${d.attendanceRank==n?'selected':''}>Top${n}</option>`).join('')}
            <option value="0" ${d.attendanceRank===0?'selected':''}>其他</option>
          </select></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="pendingCount" value="${d.pendingCount||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="lowViolations" value="${d.lowViolations||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="medViolations" value="${d.medViolations||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="highViolations" value="${d.highViolations||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="bonusActivity" value="${d.bonusActivity||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="bonusInitiative" value="${d.bonusInitiative||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="bonusOnlineShare" value="${d.bonusOnlineShare||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="bonusEventOwner" value="${d.bonusEventOwner||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="bonusTea" value="${d.bonusTea||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="bonusQuiz" value="${d.bonusQuiz||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="bonusTrainer" value="${d.bonusTrainer||0}" style="width:50px"></td>
          <td><input type="number" class="field team-cell" data-mid="${m.id}" data-fld="bonusOtherPush" value="${d.bonusOtherPush||0}" style="width:55px" placeholder="1~3每项"></td>
          <td><input type="checkbox" class="team-cell" data-mid="${m.id}" data-fld="deductQualityRule" ${d.deductQualityRule?'checked':''}></td>
          <td><input type="checkbox" class="team-cell" data-mid="${m.id}" data-fld="deductDragGroup" ${d.deductDragGroup?'checked':''}></td>
          <td><input type="checkbox" class="team-cell" data-mid="${m.id}" data-fld="deductErrorsGte3" ${d.deductErrorsGte3?'checked':''}></td>
          <td><input type="checkbox" class="team-cell" data-mid="${m.id}" data-fld="deductGroupAllFail" ${d.deductGroupAllFail?'checked':''}></td>
          <td class="team-score-cell ${score>=0?'score-pos':'score-neg'}"><b>${score}</b></td>
        </tr>`;
      }).join('')}
    </tbody>
  </table>
  </div>
  <div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap">
    <button class="btn primary" data-act="team-save-week" data-week="${weekKey}">💾 保存本周数据</button>
    <button class="btn yellow sm" data-act="team-show-rules" data-week="${weekKey}>📋 查看积分规则</button>
  </div>`;
}

/* ---- 积分核算引擎 ---- */
function calcMemberWeekScore(member, d, wk, weekKey) {
  let score = 0;
  // 每周加分
  if (d.bonusNoErrorWeek) score += 1;           // 个人本周无差错
  if (d.bonusPersonalGoal) score += 1;          // 个人完成率目标
  if (d.bonusSubGroupGoal) score += 1;         // 小小组完成率达标
  // 每月加分（按周录入，月度汇总时统计）
  // 应出勤排名
  if (d.attendanceRank === 1) score += 3;
  else if (d.attendanceRank === 2) score += 2;
  else if (d.attendanceRank === 3) score += 1;
  // 无严错
  const errCount = parseInt(d.seriousErrors) || 0;
  if (errCount === 0) score += 2;
  // 其他加分
  score += (parseInt(d.bonusActivity)||0) * 3;       // 组织活动 ×3
  score += (parseInt(d.bonusInitiative)||0) * 2;     // 主动补位 ×2
  score += (parseInt(d.bonusOnlineShare)||0) * 1;     // 线上分享 ×1
  score += (parseInt(d.bonusEventOwner)||0) * 2;      // 大事件负责人 ×2
  score += (parseInt(d.bonusTea)||0) * 1;             // 茶水 ×1
  score += (parseInt(d.bonusQuiz)||0) * 1;            // 答题 ×1
  score += (parseInt(d.bonusTrainer)||0) * 3;         // 讲师 ×3
  score += (parseInt(d.bonusOtherPush)||0);           // 其他推动

  // 扣分：每人每周
  score -= errCount * 0.5;                            // 每个严错 -0.5
  score -= (parseInt(d.pendingCount)||0) * 0.5;       // pending/误稿 -0.5 each
  // 扣分：每月
  if (d.deductErrorsGte3) score -= 2;                // 严错≥3个
  if (d.deductQualityRule) score -= 1;               // 触及质量条例
  if (d.deductDragGroup) score -= 2;                 // 拖累小组
  score -= (parseInt(d.lowViolations)||0) * 1;       // Low违规
  score -= (parseInt(d.medViolations)||0) * 1;       // Medium违规
  score -= (parseInt(d.highViolations)||0) * 3;      // High违规
  if (d.deductGroupAllFail) score -= 2;              // 全员不达标

  return Math.round(score * 10) / 10;
}

function calcMonthTotal(memberId, upToWeek) {
  let total = 0;
  const weeks = Object.keys(state.team.weeks || {}).sort();
  for (const wk of weeks) {
    if (upToWeek && wk > upToWeek) break;
    const wd = state.team.weeks[wk];
    if (!wd || !wd.data || !wd.data[memberId]) continue;
    total += calcMemberWeekScore(
      state.team.members.find(m => m.id === memberId) || {},
      wd.data[memberId], wd, wk
    );
  }
  return Math.round(total * 10) / 10;
}

/* ---- 积分看板 ---- */
function renderTeamScoreboard(v) {
  const T = state.team;
  const weekKeys = Object.keys(T.weeks || {}).sort().reverse();
  const selMonth = T._sbMonth || (new Date().getMonth() + 1);

  v.innerHTML = `<div class="card"><div class="card-head"><h2>🏆 积分看板</h2>
    <span class="ch-sub">
      月份: <select class="field" id="sbMonthSelect" style="width:100px">
        ${[1,2,3,4,5,6,7,8,9,10,11,12].map(m => `<option value="${m}" ${m==selMonth?'selected':''}>${m}月</option>`).join('')}
      </select>
    </span>
  </div>
  <div id="sbContent"></div></div>`;
  renderSBContent(selMonth);
}
function renderSBContent(month) {
  const T = state.team;
  const c = $('#sbContent');
  const year = new Date().getFullYear();
  const monthWeeks = Object.keys(T.weeks || {}).filter(k => {
    const m = k.match(/^(\d{4})-W(\d{2})$/);
    if (!m) return false;
    // 简单判断：根据 weekDate 判断月份
    const wd = T.weeks[k];
    if (!wd || !wd.weekDate) return false;
    const d = new Date(wd.weekDate);
    return d.getFullYear() === year && (d.getMonth() + 1) === month;
  }).sort();

  if (monthWeeks.length === 0) { c.innerHTML = '<div class="empty">该月暂无周数据，请先在「周数据录入」中添加</div>'; return; }

  // 月度累计排行
  const monthlyTotals = T.members.map(m => ({
    id: m.id, name: m.name,
    total: calcMonthTotal(m.id, monthWeeks[monthWeeks.length - 1])
  })).sort((a, b) => b.total - a.total);

  let html = `<div class="card card-soft" style="margin-bottom:16px">
    <h3 style="margin:0 0 10px;font-size:15px">🥇 ${month}月累计积分 Top 排行 <span style="font-weight:400;color:var(--ink-faint);font-size:12px">（截至最新一周）</span></h3>
    <div class="podium">`;
  monthlyTotals.forEach((m, i) => {
    const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `${i + 1}`;
    const isTop3 = i < 3;
    html += `<div class="podium-item ${isTop3 ? 'podium-top3' : ''}">
      <span class="podium-rank">${medal}</span>
      <span class="podium-name">${esc(m.name)}</span>
      <span class="podium-score ${m.total >= 0 ? 'score-pos' : 'score-neg'}">${m.total}</span>
    </div>`;
  });
  html += `</div>
    ${monthlyTotals.length > 0 && monthlyTotals[0].total > 0 ? '<p style="margin:8px 0 0;font-size:12px;color:var(--ink-faint)">🎁 每月积分 Top 3 有价值不等的额外惊喜哦～</p>' : ''}
  </div>`;

  // 各周明细
  html += `<h3 style="margin:16px 0 10px;font-size:15px">📊 各周积分明细</h3>
    <div class="card card-soft"><div style="overflow-x:auto">
    <table class="team-table">
      <thead><tr><th>成员</th>${monthWeeks.map(wk => `<th>${wk}<br><small>${(T.weeks[wk]||{}).weekDate||''}</small></th>`).join('')}<th>月合计</th></tr></thead>
      <tbody>`;
  T.members.forEach(m => {
    html += `<tr><td><b>${esc(m.name)}</b></td>`;
    let mTotal = 0;
    monthWeeks.forEach(wk => {
      const wd = T.weeks[wk];
      const d = (wd && wd.data) ? wd.data[m.id] : null;
      const s = d ? calcMemberWeekScore(m, d, wd, wk) : '-';
      if (typeof s === 'number') mTotal += s;
      const cls = typeof s === 'number' ? (s >= 0 ? 'score-pos' : 'score-neg') : '';
      html += `<td class="${cls}" style="text-align:center">${s}</td>`;
    });
    html += `<td class="${mTotal >= 0 ? 'score-pos' : 'score-neg'}" style="text-align:center;font-weight:800">${Math.round(mTotal*10)/10}</td></tr>`;
  });
  html += `</tbody></table></div></div>`;

  c.innerHTML = html;
}

/* ---- 积分规则弹窗 ---- */
function showTeamRules() {
  openModal(`<h2>📋 积分规则一览</h2>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;font-size:13px;line-height:1.8">
      <div><h4 style="color:var(--green-500);margin:0 0 6px">✅ 加分</h4>
        <table style="width:100%;border-collapse:collapse">
          <tr><td colspan="3" style="background:var(--green-50);font-weight:700;padding:4px 8px">每周</td></tr>
          <tr><td>个人完成率目标</td><td>+1</td></tr>
          <tr><td>小小组完成率达标</td><td>+1</td></tr>
          <tr><td>个人本周无差错</td><td>+1</td></tr>
          <tr><td colspan="3" style="background:var(--green-50);font-weight:700;padding:4px 8px;margin-top:6px">每月</td></tr>
          <tr><td>应出勤组内 Top 1</td><td>+3</td></tr>
          <tr><td>应出勤组内 Top 2</td><td>+2</td></tr>
          <tr><td>应出勤组内 Top 3</td><td>+1</td></tr>
          <tr><td>无严错</td><td>+2</td></tr>
          <tr><td>组织技能分享/座谈会等线下活动</td><td>+3/次</td></tr>
          <tr><td>主动接收临时紧急任务/补位/提建议</td><td>+2/次</td></tr>
          <tr><td>线上 case 分享/tips/疑难攻坚</td><td>+1/次</td></tr>
          <tr><td>周内大事件负责人</td><td>+2/次</td></tr>
          <tr><td>组织下茶水</td><td>+1/次</td></tr>
          <tr><td>小组答题</td><td>+1/次</td></tr>
          <tr><td>科室培训讲师</td><td>+3/次</td></tr>
          <tr><td>科室其他推动</td><td>+1~3</td></tr>
        </table>
      </div>
      <div><h4 style="color:#C62828;margin:0 0 6px">❌ 扣分</h4>
        <table style="width:100%;border-collapse:collapse">
          <tr><td colspan="2" style="background:#FFEBEE;font-weight:700;padding:4px 8px">每人/每周</td></tr>
          <tr><td>每产生 1 个严错</td><td>-0.5</td></tr>
          <tr><td>pending / 误稿 / 漏 QC</td><td>-0.5/次</td></tr>
          <tr><td colspan="2" style="background:#FFEBEE;font-weight:700;padding:4px 8px;margin-top:6px">每月</td></tr>
          <tr><td>严错 ≥ 3 个</td><td>-2</td></tr>
          <tr><td>触及当月质量目标具体条例</td><td>-1</td></tr>
          <tr><td>完成率低导致小组不达标</td><td>-2</td></tr>
          <tr><td>Low 等级违规（未送美修/delay）</td><td>-1/次</td></tr>
          <tr><td>Medium 等级违规</td><td>-1/次</td></tr>
          <tr><td>High 等级违规</td><td>-3/次</td></tr>
          <tr><td>小小组成员完成率均不达标</td><td>-2</td></tr>
        </table>
        <p style="margin-top:10px;color:var(--ink-faint);font-size:12px">⚠️ 积分规则：每周一同步上周积分情况，每月统计一次总积分<br>每月积分 Top 3 有价值不等的额外惊喜哦～</p>
      </div>
    </div>`);
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
    <div class="set-row"><span>清空所有数据</span><button class="btn sm ghost" data-act="wipe">清空</button></div>`);
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
  else if (act === 'team-tab') { state.team.activeTab = el.dataset.ttab; save(); viewTeamGoals($('#teamPanel')); }
  else if (act === 'team-save-goals') {
    $$('.team-target-input').forEach(input => {
      const gid = input.dataset.gid, sgid = input.dataset.sgid, level = input.level;
      const val = input.value.trim();
      if (level === 'group') { const g = state.team.groups.find(x => x.id === gid); if (g) g.target = val; }
      else if (level === 'subgroup') {
        const g = state.team.groups.find(x => x.id === gid);
        if (g) { const sg = g.subGroups.find(x => x.id === sgid); if (sg) sg.target = val; }
      }
    });
    save(); toast('目标已保存 ✅');
  }
  else if (act === 'team-add-member') {
    openModal(`<h2>＋ 添加成员</h2>
      <div class="row"><label>姓名</label><input class="field" id="newMName" placeholder="成员姓名" style="flex:1"></div>
      <div class="row"><label>个人目标%</label><input type="number" class="field" id="newMTarget" placeholder="如 95" style="flex:1"></div>
      <div class="row"><label>所属大组</label><select class="field" id="newMGroup" style="flex:1">
        ${state.team.groups.map(g => `<option value="${g.id}">${esc(g.name)}</option>`).join('')}
      </select></div>
      <div class="row"><label>所属小小组</label><select class="field" id="newMSGroup" style="flex:1">
        ${(state.team.groups[0].subGroups||[]).map(sg => `<option value="${sg.id}">${esc(sg.name)}</option>`).join('')}
      </select></div>
      <div class="row"><button class="btn primary" data-act="team-add-member-do">添加</button></div>`);
  }
  else if (act === 'team-add-member-do') {
    const name = $('#newMName').value.trim();
    if (!name) { toast('请输入姓名'); return; }
    const mid = uid();
    state.team.members.push({ id: mid, name, personalTarget: $('#newMTarget').value.trim(), groupId: $('#newMGroup').value, subGroupId: $('#newMSGroup').value });
    // 加入小小组
    const g = state.team.groups.find(x => x.id === $('#newMGroup').value);
    if (g) { const sg = g.subGroups.find(x => x.id === $('#newMSGroup').value); if (sg) sg.memberIds.push(mid); }
    save(); closeModal(); viewTeamGoals($('#teamPanel')); toast('成员已添加 ✅');
  }
  else if (act === 'team-del-member') {
    const mid = el.dataset.mid;
    if (!confirm('确定删除该成员？')) return;
    state.team.members = state.team.members.filter(m => m.id !== mid);
    state.team.groups.forEach(g => g.subGroups.forEach(sg => sg.memberIds = sg.memberIds.filter(id => id !== mid)));
    save(); viewTeamGoals($('#teamPanel')); toast('已删除');
  }
  else if (act === 'team-rm-member-sg') {
    const mid = el.dataset.mid, sgid = el.dataset.sgid;
    state.team.groups.forEach(g => g.subGroups.forEach(sg => { if (sg.id === sgid) sg.memberIds = sg.memberIds.filter(id => id !== mid); }));
    save(); viewTeamGoals($('#teamPanel'));
  }
  else if (act === 'team-save-members') {
    $$('[data-mid][data-fld]').forEach(input => {
      const mid = input.dataset.mid, fld = input.fld;
      const m = state.team.members.find(x => x.id === mid);
      if (!m) return;
      if (fld === 'groupId') {
        // 从旧小组移除，加入新小组
        const oldGid = m.groupId;
        const newGid = input.value;
        if (oldGid !== newGid) {
          const oldG = state.team.groups.find(g => g.id === oldGid);
          if (oldG) oldG.subGroups.forEach(sg => sg.memberIds = sg.memberIds.filter(id => id !== mid));
          m.groupId = newGid;
          m.subGroupId = '';
        }
      } else if (fld === 'subGroupId') {
        const oldSgid = m.subGroupId, newSgid = input.value;
        if (oldSgid !== newSgid) {
          state.team.groups.forEach(g => g.subGroups.forEach(sg => { if (sg.id === oldSgid) sg.memberIds = sg.memberIds.filter(id => id !== mid); }));
          m.subGroupId = newSgid;
          const g = state.team.groups.find(gr => gr.id === m.groupId);
          if (g) { const sg = g.subGroups.find(s => s.id === newSgid); if (sg && !sg.memberIds.includes(mid)) sg.memberIds.push(mid); }
        }
      } else { m[fld] = input.value.trim(); }
    });
    save(); toast('成员信息已保存 ✅');
  }
  else if (act === 'team-create-week') {
    const dateStr = $('#newWeekDate').value || todayStr();
    const wk = getWeekKey(dateStr);
    if (state.team.weeks[wk]) { toast('该周已存在'); return; }
    state.team.weeks[wk] = { weekDate: dateStr, data: {} };
    state.team._selectedWeek = wk; save();
    renderTeamWeekly($('#teamPanel')); toast(`已创建 ${wk}`);
  }
  else if (act === 'team-save-week') {
    const wk = el.dataset.week;
    if (!state.team.weeks[wk]) { toast('周数据不存在'); return; }
    state.team.weeks[wk].data = state.team.weeks[wk].data || {};
    $$('.team-cell').forEach(input => {
      const mid = input.dataset.mid, fld = input.fld;
      if (!mid || !fld) return;
      if (!state.team.weeks[wk].data[mid]) state.team.weeks[wk].data[mid] = {};
      let val;
      if (input.type === 'checkbox') val = input.checked;
      else if (input.type === 'number' || input.tagName === 'SELECT') val = input.value === '' ? '' : parseFloat(input.value);
      else val = input.value;
      state.team.weeks[wk].data[mid][fld] = val;
    });
    save(); toast(`${wk} 数据已保存 ✅`);
    renderWeeklyForm(wk); // 刷新显示积分
  }
  else if (act === 'team-show-rules') { showTeamRules(); }
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
/* 小组管理：周选择 / 月选择 change */
$('#view').addEventListener('change', e => {
  if (e.target.id === 'weekSelect') {
    state.team._selectedWeek = e.target.value || ''; save();
    renderWeeklyForm(e.target.value);
  }
  if (e.target.id === 'sbMonthSelect') {
    state.team._sbMonth = parseInt(e.target.value); save();
    renderSBContent(state.team._sbMonth);
  }
  // 成员管理：切换大组时联动小小组
  if (e.target.classList.contains('team-sg-select')) {
    const mid = e.target.dataset.mid;
    const newGid = e.target.closest('.team-m-group').querySelector('[data-fld="groupId"]').value;
    const g = state.team.groups.find(gr => gr.id === newGid);
    if (g) { e.target.innerHTML = g.subGroups.map(sg => `<option value="${sg.id}" ${sg.id===e.target.value?'selected':''}>${esc(sg.name)}</option>`).join(''); }
  }
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
  // 注册 PWA
  if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
}
boot();
