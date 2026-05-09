const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.author = 'Clayton';
pres.title = '中东局势分析 2026';

// ============================================================
// SLIDE DIMENSIONS
// ============================================================
pres.layout = 'LAYOUT_16x9';
const SLIDE_W = 10;
const SLIDE_H = 5.625;
const MARGIN = 0.5;
const CONTENT_X = MARGIN;
const CONTENT_Y = MARGIN;
const CONTENT_W = SLIDE_W - (2 * MARGIN);
const CONTENT_H = SLIDE_H - (2 * MARGIN);
const CENTER_X = SLIDE_W / 2;
const CENTER_Y = SLIDE_H / 2;

// ============================================================
// COLOR SYSTEM - Stamen Design 数据诗学
// ============================================================
const C = {
  ochre: 'C65D3B', ochreLight: 'D97B5D',
  sage: '7D9B76', sageLight: '9AB894',
  navy: '2E3A59', navyLight: '4A5570', navyDark: '1E2638',
  alert: 'C73E3E',
  bgDark: '1A1F2E', bgLight: 'F5F3EF',
  textDark: '2C2C2C', textLight: 'F5F3EF', textMuted: '6B6B6B',
  border: 'D4D0C8',
  gold: 'D4A843',
};

// ============================================================
// CONTAINER SYSTEM
// ============================================================
function createVirtualNode(type, data, parentX = 0, parentY = 0) {
  const opts = data.opts || {};
  const node = {
    type, data,
    absX: parentX + (opts.x || 0),
    absY: parentY + (opts.y || 0),
    w: opts.w || 0, h: opts.h || 0,
    children: []
  };
  node.addShape = function(shapeType, opts = {}) {
    const child = createVirtualNode('shape', { shapeType, opts }, node.absX, node.absY);
    node.children.push(child);
    return child;
  };
  node.addText = function(text, opts = {}) {
    const safeOpts = { fit: "shrink", ...opts };
    const child = createVirtualNode('text', { text, opts: safeOpts }, node.absX, node.absY);
    node.children.push(child);
    return child;
  };
  node.addImage = function(opts = {}) {
    const child = createVirtualNode('image', { opts }, node.absX, node.absY);
    node.children.push(child);
    return child;
  };
  node.addTable = function(tableData, opts = {}) {
    const child = createVirtualNode('table', { tableData, opts }, node.absX, node.absY);
    node.children.push(child);
    return child;
  };
  return node;
}

function flattenNode(node, realSlide, pres) {
  const absOpts = { ...node.data.opts, x: node.absX, y: node.absY };
  if (node.type === 'shape') realSlide.addShape(node.data.shapeType, absOpts);
  else if (node.type === 'text') realSlide.addText(node.data.text, absOpts);
  else if (node.type === 'image') realSlide.addImage(absOpts);
  else if (node.type === 'table') realSlide.addTable(node.data.tableData, absOpts);
  node.children.forEach(child => flattenNode(child, realSlide, pres));
}

const originalAddSlide = pres.addSlide.bind(pres);
pres.addSlide = function(options) {
  const realSlide = originalAddSlide(options);
  const virtualSlide = {
    children: [],
    _realSlide: realSlide,
    set background(val) { realSlide.background = val; },
    get background() { return realSlide.background; },
    addShape: function(shapeType, opts = {}) {
      const node = createVirtualNode('shape', { shapeType, opts }, 0, 0);
      this.children.push(node);
      return node;
    },
    addText: function(text, opts = {}) {
      const safeOpts = { fit: "shrink", ...opts };
      const node = createVirtualNode('text', { text, opts: safeOpts }, 0, 0);
      this.children.push(node);
      return node;
    },
    addImage: function(opts = {}) {
      const node = createVirtualNode('image', { opts }, 0, 0);
      this.children.push(node);
      return node;
    },
    addTable: function(tableData, opts = {}) {
      const node = createVirtualNode('table', { tableData, opts }, 0, 0);
      this.children.push(node);
      return node;
    },
    addChart: function(chartType, data, opts = {}) {
      realSlide.addChart(chartType, data, opts);
    },
    render: function() {
      this.children.forEach(child => flattenNode(child, realSlide, pres));
    }
  };
  return virtualSlide;
};

// ============================================================
// SLIDE 1: 封面
// ============================================================
let s1 = pres.addSlide();
s1.background = { color: C.bgDark };

// 顶部渐变条
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: 0.06, fill: { color: C.ochre } });

// 标题
s1.addText('中东局势', {
  x: 0.8, y: 1.2, w: 8, h: 1.2,
  fontSize: 44, fontFace: 'Georgia', color: C.textLight, bold: true, charSpacing: 2
});
s1.addText('冲突全景', {
  x: 0.8, y: 2.2, w: 8, h: 1.2,
  fontSize: 44, fontFace: 'Georgia', color: C.ochre, bold: true, charSpacing: 2
});

// 副标题
s1.addText('美伊霍尔木兹海峡对抗 · 以色列多线作战 · 全球能源冲击', {
  x: 0.8, y: 3.4, w: 7, h: 0.5,
  fontSize: 16, color: 'B0B0B0'
});

// 日期
s1.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 4.2, w: 2.8, h: 0.5,
  fill: { color: C.bgDark }, line: { color: C.ochre, width: 1.5 }, rectRadius: 0.05
});
s1.addText('2026.04.28 — 05.08', {
  x: 0.8, y: 4.2, w: 2.8, h: 0.5,
  fontSize: 14, fontFace: 'Consolas', color: C.ochre, align: 'center', valign: 'middle'
});

// 底部
s1.addText('Stamen Design 数据诗学', { x: 0.8, y: 5.1, w: 3, h: 0.3, fontSize: 10, color: '666666' });
s1.addText('7页演示文稿', { x: 7, y: 5.1, w: 2, h: 0.3, fontSize: 10, color: '666666', align: 'right' });

s1.render();

// ============================================================
// SLIDE 2: 局势总览
// ============================================================
let s2 = pres.addSlide();
s2.background = { color: C.bgLight };

// 顶部条
s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: 0.05, fill: { color: C.ochre } });

// Kicker
s2.addText('CHAPTER 01 · OVERVIEW', {
  x: CONTENT_X, y: 0.3, w: 4, h: 0.25,
  fontSize: 9, color: C.ochre, bold: true, charSpacing: 2
});

// 标题
s2.addText('局势总览：四大战线并行', {
  x: CONTENT_X, y: 0.55, w: 8, h: 0.6,
  fontSize: 28, fontFace: 'Georgia', color: C.textDark, bold: true
});

// 四大战线卡片
const warzones = [
  { title: '美伊冲突', color: C.ochre, desc: '霍尔木兹海峡爆发停火后最激烈海战', stats: ['击沉快艇 7艘', '航运暴跌 90%', '油价跳涨 +8%'] },
  { title: '以黎冲突', color: C.gold, desc: '距停火仅9天再度开战，贝鲁特遭袭', stats: ['停火维持 9天', '紧急状态至5.19', '4条战线'] },
  { title: '加沙地带', color: C.navy, desc: '名义停火实质未止，人道危机持续恶化', stats: ['儿童失学 1亿+', '危机等级 灾难', '援助受阻'] },
  { title: '外交谈判', color: C.alert, desc: '间接会谈无突破，和平协议已作废', stats: ['会谈 无突破', '协议 已作废', '停火概率 <10%'] },
];

const cardW = (CONTENT_W - 0.45) / 4;
warzones.forEach((wz, i) => {
  const cx = CONTENT_X + i * (cardW + 0.15);
  
  // 卡片背景
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx, y: 1.4, w: cardW, h: 2.8,
    fill: { color: 'FFFFFF' }, shadow: { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.08 },
    rectRadius: 0.08
  });
  
  // 顶部色条
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx, y: 1.4, w: cardW, h: 0.05,
    fill: { color: wz.color }, rectRadius: 0.02
  });
  
  // 标题
  s2.addText(wz.title, {
    x: cx + 0.15, y: 1.55, w: cardW - 0.3, h: 0.35,
    fontSize: 14, bold: true, color: C.textDark
  });
  
  // 描述
  s2.addText(wz.desc, {
    x: cx + 0.15, y: 1.9, w: cardW - 0.3, h: 0.8,
    fontSize: 10, color: '555555', lineSpacingMultiple: 1.3
  });
  
  // 数据行
  wz.stats.forEach((stat, j) => {
    const sy = 2.85 + j * 0.4;
    s2.addShape(pres.shapes.RECTANGLE, {
      x: cx + 0.15, y: sy, w: cardW - 0.3, h: 0.35,
      fill: { color: 'F8F6F2' }, rectRadius: 0.04
    });
    s2.addText(stat, {
      x: cx + 0.15, y: sy, w: cardW - 0.3, h: 0.35,
      fontSize: 10, color: wz.color, bold: true, align: 'center', valign: 'middle'
    });
  });
});

// 底部数据条
const metrics = [
  { value: '90%', label: '霍尔木兹海峡航运量暴跌', color: C.ochre },
  { value: '108-112', label: '布伦特原油 美元/桶', color: C.alert },
  { value: '60亿$', label: '伊朗石油出口损失', color: C.sage },
  { value: '250亿$', label: '美国战争成本', color: C.navyLight },
];

const mW = (CONTENT_W - 0.45) / 4;
metrics.forEach((m, i) => {
  const mx = CONTENT_X + i * (mW + 0.15);
  
  s2.addShape(pres.shapes.RECTANGLE, {
    x: mx, y: 4.5, w: mW, h: 0.8,
    fill: { color: C.bgDark }, rectRadius: 0.06
  });
  
  s2.addShape(pres.shapes.RECTANGLE, {
    x: mx, y: 4.5, w: 0.04, h: 0.8,
    fill: { color: m.color }
  });
  
  s2.addText(m.value, {
    x: mx + 0.15, y: 4.5, w: mW - 0.3, h: 0.45,
    fontSize: 20, fontFace: 'Georgia', color: m.color, bold: true
  });
  
  s2.addText(m.label, {
    x: mx + 0.15, y: 4.9, w: mW - 0.3, h: 0.35,
    fontSize: 8, color: '999999'
  });
});

// 页码
s2.addText('02 / 07', { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 9, color: C.textMuted, align: 'right' });

s2.render();

// ============================================================
// SLIDE 3: 美伊冲突时间线
// ============================================================
let s3 = pres.addSlide();
s3.background = { color: C.bgDark };

s3.addText('CHAPTER 02 · US-IRAN CONFLICT', {
  x: CONTENT_X, y: 0.3, w: 5, h: 0.25,
  fontSize: 9, color: C.ochre, bold: true, charSpacing: 2
});

s3.addText('美伊冲突时间线', {
  x: CONTENT_X, y: 0.55, w: 5, h: 0.5,
  fontSize: 28, fontFace: 'Georgia', color: C.textLight, bold: true
});

s3.addText('霍尔木兹海峡军事对抗升级全过程', {
  x: CONTENT_X, y: 1.0, w: 5, h: 0.3,
  fontSize: 12, color: '888888'
});

// 时间线事件
const events = [
  { date: '5月4日', text: '美国启动"自由计划"护航行动，美军击沉7艘伊朗快艇，伊朗向美军舰及商船发射导弹/无人机' },
  { date: '5月5日', text: '鲁比奥宣布"史诗怒火行动"结束，但封锁持续。伊朗不承认美方停火' },
  { date: '5月6日', text: '特朗普暂停"自由计划"。美军炸坏伊朗哈斯娜号油轮船舵，伊朗击落2架美以无人机' },
  { date: '5月7-8日', text: '最激烈冲突：美军空袭伊朗沿海，伊朗发射反舰导弹宣称重创3艘宙斯盾驱逐舰' },
];

events.forEach((ev, i) => {
  const ey = 1.5 + i * 0.75;
  
  // 时间线圆点
  s3.addShape(pres.shapes.OVAL, {
    x: CONTENT_X + 0.05, y: ey + 0.05, w: 0.12, h: 0.12,
    fill: { color: C.ochre }
  });
  
  // 日期
  s3.addText(ev.date, {
    x: CONTENT_X + 0.3, y: ey, w: 1.2, h: 0.25,
    fontSize: 10, fontFace: 'Consolas', color: C.ochre, bold: true
  });
  
  // 内容
  s3.addText(ev.text, {
    x: CONTENT_X + 0.3, y: ey + 0.25, w: 4.5, h: 0.45,
    fontSize: 10, color: 'CCCCCC', lineSpacingMultiple: 1.2
  });
});

// 右侧数据卡片
s3.addShape(pres.shapes.RECTANGLE, {
  x: 5.8, y: 1.3, w: 3.7, h: 1.2,
  fill: { color: '2A1F1F' }, line: { color: C.ochre, width: 1 }, rectRadius: 0.08
});
s3.addText('300+ 艘', {
  x: 5.8, y: 1.4, w: 3.7, h: 0.6,
  fontSize: 32, fontFace: 'Georgia', color: C.ochre, bold: true, align: 'center'
});
s3.addText('船只滞留霍尔木兹海峡', {
  x: 5.8, y: 2.0, w: 3.7, h: 0.3,
  fontSize: 10, color: 'AAAAAA', align: 'center'
});

// 小数据网格
const miniStats = [
  { value: '7', label: '快艇击沉' },
  { value: '2', label: '无人机击落' },
  { value: '3', label: '宙斯盾重创' },
  { value: '8%', label: '油价跳涨' },
];

miniStats.forEach((ms, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const mx = 5.8 + col * 1.85;
  const my = 2.8 + row * 0.9;
  
  s3.addShape(pres.shapes.RECTANGLE, {
    x: mx, y: my, w: 1.7, h: 0.75,
    fill: { color: '252A3A' }, rectRadius: 0.06
  });
  s3.addText(ms.value, {
    x: mx, y: my, w: 1.7, h: 0.45,
    fontSize: 22, fontFace: 'Georgia', color: C.textLight, bold: true, align: 'center'
  });
  s3.addText(ms.label, {
    x: mx, y: my + 0.4, w: 1.7, h: 0.3,
    fontSize: 9, color: '888888', align: 'center'
  });
});

// 经济损失
s3.addShape(pres.shapes.RECTANGLE, {
  x: 5.8, y: 4.6, w: 3.7, h: 0.8,
  fill: { color: '252A3A' }, rectRadius: 0.06
});
s3.addText('60亿美元', {
  x: 5.8, y: 4.6, w: 3.7, h: 0.5,
  fontSize: 24, fontFace: 'Georgia', color: C.sage, bold: true, align: 'center'
});
s3.addText('伊朗石油出口损失 + 美国250亿美元战争成本', {
  x: 5.8, y: 5.05, w: 3.7, h: 0.3,
  fontSize: 9, color: '888888', align: 'center'
});

s3.addText('03 / 07', { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 9, color: '666666', align: 'right' });
s3.render();

// ============================================================
// SLIDE 4: 以黎与加沙
// ============================================================
let s4 = pres.addSlide();
s4.background = { color: C.bgLight };

s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: 0.05, fill: { color: C.gold } });

s4.addText('CHAPTER 03 · ISRAEL-LEBANON & GAZA', {
  x: CONTENT_X, y: 0.3, w: 6, h: 0.25,
  fontSize: 9, color: C.ochre, bold: true, charSpacing: 2
});

// 左列 - 以黎
s4.addText('以色列-黎巴嫩冲突', {
  x: CONTENT_X, y: 0.6, w: 4, h: 0.4,
  fontSize: 20, fontFace: 'Georgia', color: C.textDark, bold: true
});

const lebanonEvents = [
  { date: '5月2日', title: '新一轮空袭启动', desc: '距停火仅9天，以色列发动新一轮空袭' },
  { date: '5月3日', title: '边境冲突持续', desc: '与真主党、胡塞武装交火持续' },
  { date: '5月6日', title: '贝鲁特遭空袭', desc: '打死真主党拉德万部队指挥官' },
  { date: '5月7日', title: '紧急状态延长', desc: '全国紧急状态延长至5月19日' },
];

lebanonEvents.forEach((ev, i) => {
  const ey = 1.15 + i * 0.7;
  s4.addShape(pres.shapes.RECTANGLE, {
    x: CONTENT_X, y: ey, w: 4.2, h: 0.6,
    fill: { color: 'FFFFFF' }, shadow: { type: 'outer', blur: 4, offset: 1, color: '000000', opacity: 0.06 },
    rectRadius: 0.06
  });
  s4.addShape(pres.shapes.RECTANGLE, {
    x: CONTENT_X, y: ey, w: 0.04, h: 0.6,
    fill: { color: C.gold }
  });
  s4.addText(ev.date + '  ' + ev.title, {
    x: CONTENT_X + 0.15, y: ey + 0.05, w: 3.9, h: 0.25,
    fontSize: 10, bold: true, color: C.textDark
  });
  s4.addText(ev.desc, {
    x: CONTENT_X + 0.15, y: ey + 0.3, w: 3.9, h: 0.25,
    fontSize: 9, color: '666666'
  });
});

// 以黎数据
['9天 停火维持', '4条 作战方向', '70% 军工被毁'].forEach((t, i) => {
  const tx = CONTENT_X + i * 1.45;
  s4.addShape(pres.shapes.RECTANGLE, {
    x: tx, y: 4.1, w: 1.35, h: 0.55,
    fill: { color: C.bgDark }, rectRadius: 0.05
  });
  s4.addText(t, {
    x: tx, y: 4.1, w: 1.35, h: 0.55,
    fontSize: 11, color: C.textLight, bold: true, align: 'center', valign: 'middle'
  });
});

// 右列 - 加沙
s4.addText('加沙地带人道危机', {
  x: 5.3, y: 0.6, w: 4, h: 0.4,
  fontSize: 20, fontFace: 'Georgia', color: C.textDark, bold: true
});

const gazaEvents = [
  { date: '持续状态', title: '名义停火，实质冲突', desc: '2025年10月名义停火，敌对行动未止' },
  { date: '人道援助', title: '援助船队被拦截', desc: '以色列拦截前往加沙的人道援助船队' },
  { date: '教育危机', title: '1亿+儿童教育中断', desc: '阿拉伯国家超1亿儿童教育中断' },
  { date: '国际评估', title: '灾难性人道危机', desc: '被国际组织称为"灾难性"危机' },
];

gazaEvents.forEach((ev, i) => {
  const ey = 1.15 + i * 0.7;
  s4.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: ey, w: 4.2, h: 0.6,
    fill: { color: 'FFFFFF' }, shadow: { type: 'outer', blur: 4, offset: 1, color: '000000', opacity: 0.06 },
    rectRadius: 0.06
  });
  s4.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: ey, w: 0.04, h: 0.6,
    fill: { color: C.navy }
  });
  s4.addText(ev.date + '  ' + ev.title, {
    x: 5.45, y: ey + 0.05, w: 3.9, h: 0.25,
    fontSize: 10, bold: true, color: C.textDark
  });
  s4.addText(ev.desc, {
    x: 5.45, y: ey + 0.3, w: 3.9, h: 0.25,
    fontSize: 9, color: '666666'
  });
});

['1亿+ 儿童失学', '灾难 危机等级', '持续 援助受阻'].forEach((t, i) => {
  const tx = 5.3 + i * 1.45;
  s4.addShape(pres.shapes.RECTANGLE, {
    x: tx, y: 4.1, w: 1.35, h: 0.55,
    fill: { color: C.bgDark }, rectRadius: 0.05
  });
  s4.addText(t, {
    x: tx, y: 4.1, w: 1.35, h: 0.55,
    fontSize: 11, color: C.textLight, bold: true, align: 'center', valign: 'middle'
  });
});

s4.addText('04 / 07', { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 9, color: C.textMuted, align: 'right' });
s4.render();

// ============================================================
// SLIDE 5: 能源冲击
// ============================================================
let s5 = pres.addSlide();
s5.background = { color: C.bgDark };

s5.addText('CHAPTER 04 · ENERGY IMPACT', {
  x: CONTENT_X, y: 0.3, w: 5, h: 0.25,
  fontSize: 9, color: C.ochre, bold: true, charSpacing: 2
});

s5.addText('全球能源冲击', {
  x: CONTENT_X, y: 0.55, w: 5, h: 0.5,
  fontSize: 28, fontFace: 'Georgia', color: C.textLight, bold: true
});

s5.addText('霍尔木兹海峡承载全球约30%海运原油', {
  x: CONTENT_X, y: 1.0, w: 5, h: 0.3,
  fontSize: 12, color: '888888'
});

// 左侧影响列表
const impacts = [
  { title: '航运危机', desc: '海峡紧急禁航，油轮绕道好望角，运费暴涨3倍' },
  { title: '油价飙升', desc: '布伦特突破92美元，若完全封锁或飙至150美元+' },
  { title: '炼油停摆', desc: '阿布扎比炼厂关闭，日均190万桶炼油能力停摆' },
  { title: '金融去美元化', desc: '伊朗推动能源去美元化，与中俄本币结算' },
];

impacts.forEach((imp, i) => {
  const iy = 1.5 + i * 0.75;
  s5.addShape(pres.shapes.RECTANGLE, {
    x: CONTENT_X, y: iy, w: 4.8, h: 0.65,
    fill: { color: '252A3A' }, line: { color: '3A3F50', width: 0.5 }, rectRadius: 0.06
  });
  s5.addText(imp.title, {
    x: CONTENT_X + 0.15, y: iy + 0.05, w: 4.5, h: 0.25,
    fontSize: 12, bold: true, color: C.textLight
  });
  s5.addText(imp.desc, {
    x: CONTENT_X + 0.15, y: iy + 0.3, w: 4.5, h: 0.3,
    fontSize: 9, color: 'AAAAAA'
  });
});

// 右侧大数字
s5.addShape(pres.shapes.RECTANGLE, {
  x: 5.5, y: 1.3, w: 4, h: 1.3,
  fill: { color: '2A1F1F' }, line: { color: C.ochre, width: 1 }, rectRadius: 0.1
});
s5.addText('30%', {
  x: 5.5, y: 1.35, w: 4, h: 0.8,
  fontSize: 48, fontFace: 'Georgia', color: C.ochre, bold: true, align: 'center'
});
s5.addText('全球海运原油经霍尔木兹海峡运输', {
  x: 5.5, y: 2.1, w: 4, h: 0.3,
  fontSize: 11, color: 'AAAAAA', align: 'center'
});

// 油价对比
const prices = [
  { label: '冲突前', value: '$92', color: C.navyLight },
  { label: '当前', value: '$108-112', color: C.ochre },
  { label: '若海峡封锁', value: '$150+', color: C.alert },
];

prices.forEach((p, i) => {
  const px = 5.5 + i * 1.35;
  const barH = [1.2, 1.7, 2.2][i];
  
  s5.addShape(pres.shapes.RECTANGLE, {
    x: px, y: 3.0 + (2.2 - barH), w: 1.15, h: barH,
    fill: { color: p.color }, rectRadius: 0.06
  });
  s5.addText(p.value, {
    x: px, y: 2.7, w: 1.15, h: 0.3,
    fontSize: 12, bold: true, color: p.color, align: 'center'
  });
  s5.addText(p.label, {
    x: px, y: 5.25, w: 1.15, h: 0.25,
    fontSize: 9, color: '888888', align: 'center'
  });
});

s5.addText('05 / 07', { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 9, color: '666666', align: 'right' });
s5.render();

// ============================================================
// SLIDE 6: 未来情景
// ============================================================
let s6 = pres.addSlide();
s6.background = { color: C.bgLight };

s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: 0.05, fill: { color: C.ochre } });

s6.addText('CHAPTER 05 · FUTURE SCENARIOS', {
  x: CONTENT_X, y: 0.3, w: 6, h: 0.25,
  fontSize: 9, color: C.ochre, bold: true, charSpacing: 2
});

s6.addText('未来情景预测', {
  x: CONTENT_X, y: 0.55, w: 6, h: 0.5,
  fontSize: 28, fontFace: 'Georgia', color: C.textDark, bold: true
});

s6.addText('未来三个月是关键窗口期，全面停火可能性不足10%', {
  x: CONTENT_X, y: 1.0, w: 6, h: 0.3,
  fontSize: 12, color: C.textMuted
});

const scenarios = [
  { name: '长期消耗战', prob: '50%', color: C.ochre, desc: '美以无法推翻伊朗政权，但持续施压，中东陷入"冷和平、热冲突"泥潭' },
  { name: '能源脱钩加速', prob: '30%', color: C.sage, desc: '全球能源格局重塑，伊朗与中俄本币结算扩大，石油美元体系受冲击' },
  { name: '全面战争', prob: '15%', color: C.alert, desc: '若误判导致核设施受损或海峡完全封锁，冲突可能升级为全面战争' },
  { name: '大国协调停火', prob: '5%', color: C.navy, desc: '需要中美俄等大国协调一致，当前外交僵局下概率极低' },
];

scenarios.forEach((sc, i) => {
  const sy = 1.5 + i * 0.7;
  
  s6.addShape(pres.shapes.RECTANGLE, {
    x: CONTENT_X, y: sy, w: CONTENT_W, h: 0.6,
    fill: { color: 'FFFFFF' }, shadow: { type: 'outer', blur: 4, offset: 1, color: '000000', opacity: 0.06 },
    rectRadius: 0.06
  });
  
  // 概率徽章
  s6.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: CONTENT_X + 0.1, y: sy + 0.1, w: 0.8, h: 0.4,
    fill: { color: sc.color + '20' }, rectRadius: 0.15
  });
  s6.addText(sc.prob, {
    x: CONTENT_X + 0.1, y: sy + 0.1, w: 0.8, h: 0.4,
    fontSize: 16, fontFace: 'Georgia', color: sc.color, bold: true, align: 'center', valign: 'middle'
  });
  
  // 名称
  s6.addText(sc.name, {
    x: CONTENT_X + 1.1, y: sy + 0.05, w: 1.8, h: 0.5,
    fontSize: 14, bold: true, color: C.textDark, valign: 'middle'
  });
  
  // 描述
  s6.addText(sc.desc, {
    x: CONTENT_X + 3, y: sy + 0.05, w: 5.8, h: 0.5,
    fontSize: 10, color: '555555', valign: 'middle', lineSpacingMultiple: 1.2
  });
});

// 关键判断
s6.addShape(pres.shapes.RECTANGLE, {
  x: CONTENT_X, y: 4.4, w: CONTENT_W, h: 0.8,
  fill: { color: C.bgDark }, rectRadius: 0.06
});
s6.addText('关键判断', {
  x: CONTENT_X + 0.2, y: 4.45, w: 2, h: 0.25,
  fontSize: 9, color: '888888', bold: true, charSpacing: 2
});
s6.addText('短期内局势难以缓和。长期消耗战是最可能路径，能源市场将持续承压，地缘政治风险溢价维持高位。', {
  x: CONTENT_X + 0.2, y: 4.7, w: CONTENT_W - 0.4, h: 0.4,
  fontSize: 11, color: C.textLight
});

s6.addText('06 / 07', { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 9, color: C.textMuted, align: 'right' });
s6.render();

// ============================================================
// SLIDE 7: 总结
// ============================================================
let s7 = pres.addSlide();
s7.background = { color: C.bgDark };

s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: 0.06, fill: { color: C.ochre } });

s7.addText('CONCLUSION', {
  x: 0, y: 0.8, w: SLIDE_W, h: 0.3,
  fontSize: 10, color: '888888', charSpacing: 4, align: 'center'
});

s7.addText([
  { text: '中东局势：', options: { fontSize: 32, fontFace: 'Georgia', color: C.textLight, bold: true } },
  { text: '高度紧张', options: { fontSize: 32, fontFace: 'Georgia', color: C.ochre, bold: true } },
  { text: '，和平前景', options: { fontSize: 32, fontFace: 'Georgia', color: C.textLight, bold: true } },
  { text: '严峻', options: { fontSize: 32, fontFace: 'Georgia', color: C.ochre, bold: true } },
], {
  x: 1, y: 1.2, w: 8, h: 0.8, align: 'center'
});

// 三个关键数字
const conclusions = [
  { number: '4', label: '条战线同时升级\n美伊、以黎、加沙、外交' },
  { number: '90%', label: '霍尔木兹海峡\n航运量暴跌' },
  { number: '<10%', label: '全面停火\n可能性' },
];

conclusions.forEach((c, i) => {
  const cx = 1.5 + i * 2.8;
  
  s7.addShape(pres.shapes.RECTANGLE, {
    x: cx, y: 2.4, w: 2.3, h: 1.5,
    fill: { color: '252A3A' }, line: { color: '3A3F50', width: 0.5 }, rectRadius: 0.08
  });
  
  s7.addText(c.number, {
    x: cx, y: 2.5, w: 2.3, h: 0.7,
    fontSize: 36, fontFace: 'Georgia', color: C.ochre, bold: true, align: 'center'
  });
  
  s7.addText(c.label, {
    x: cx, y: 3.2, w: 2.3, h: 0.6,
    fontSize: 10, color: 'AAAAAA', align: 'center', lineSpacingMultiple: 1.3
  });
});

// 来源
s7.addText('信息来源：今日头条、央视新闻、CGTN、以色列国防部、伊朗武装部队官方声明、美国中央司令部公告', {
  x: 1, y: 4.3, w: 8, h: 0.3,
  fontSize: 9, color: '666666', align: 'center'
});
s7.addText('文档生成时间：2026年5月8日', {
  x: 1, y: 4.6, w: 8, h: 0.3,
  fontSize: 9, color: '666666', align: 'center'
});

s7.addText('07 / 07', { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 9, color: '666666', align: 'right' });
s7.render();

// ============================================================
// SAVE
// ============================================================
pres.writeFile({ fileName: 'd:/TARE SPACE/middle-east-conflict-deck/deck.pptx' })
  .then(() => console.log('PPTX generated successfully!'))
  .catch(err => console.error('Error:', err));
