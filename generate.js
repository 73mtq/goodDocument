/**
 * generate.js —— 规范文档生成器 CLI 入口
 * =================================================================
 * 读取 config.json（可用 --config 指定别的文件），按配置生成 .docx。
 * 支持 --set a.b=c 临时覆盖任意参数，--out 指定输出文件名。
 *
 * 用法示例：
 *   node generate.js                                    # 用默认 config.json 生成
 *   node generate.js --set page.margins.left=2.8        # 临时改左边距
 *   node generate.js --set body.size=14 --set body.lineSpacing=2.0
 *   node generate.js --config 公文配置.json --out 公文.docx
 *   node generate.js --list-sizes                       # 查看中文字号对照表
 *   node generate.js --show-config                      # 查看生效的配置
 */

const fs = require("fs");
const path = require("path");
const { Packer, Document, Footer, Paragraph, TextRun, PageNumber, AlignmentType } = require("docx");
const {
  createFormatter, SIZE_NAME_TO_PT, resolveSize, makeFont,
  lineSpacingConfig, pageSizeDxa,
} = require("./formatter");

const CM_TO_DXA = 566.929;

/* ------------------------------------------------------------------
 * CLI 参数解析
 * ------------------------------------------------------------------ */

function parseArgs(argv) {
  const args = { set: [], _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--config" || a === "-c") { args.config = argv[++i]; continue; }
    if (a === "--out" || a === "-o") { args.out = argv[++i]; continue; }
    if (a === "--set" || a === "-s") { args.set.push(argv[++i]); continue; }
    if (a === "--list-sizes") { args.listSizes = true; continue; }
    if (a === "--show-config") { args.showConfig = true; continue; }
    if (a === "--dry-run") { args.dryRun = true; continue; }
    if (a === "--help" || a === "-h") { args.help = true; continue; }
    if (a.startsWith("--") && a.indexOf("=") > 2) {
      args.set.push(a.slice(2));
      continue;
    }
    args._.push(a);
  }
  return args;
}

/** 值类型转换：true/false → 布尔；纯数字 → 数值；其余字符串 */
function parseValue(s) {
  if (s === "true") return true;
  if (s === "false") return false;
  if (s !== "" && !isNaN(Number(s))) return Number(s);
  return s;
}

/** 按 "a.b.c" 点号路径设置对象属性 */
function setByPath(obj, keyPath, value) {
  const keys = keyPath.split(".");
  let cur = obj;
  for (let i = 0; i < keys.length - 1; i++) {
    const k = keys[i];
    if (cur[k] == null || typeof cur[k] !== "object") cur[k] = {};
    cur = cur[k];
  }
  cur[keys[keys.length - 1]] = value;
}

/** 将 --set 的 "a.b=value" 列表应用到 config */
function applyOverrides(config, setList) {
  for (const item of setList) {
    const eq = item.indexOf("=");
    if (eq < 0) { console.error("忽略无效 --set（缺 =）： " + item); continue; }
    setByPath(config, item.slice(0, eq).trim(), parseValue(item.slice(eq + 1).trim()));
  }
  return config;
}

/* ------------------------------------------------------------------
 * 示例文档内容（替换此处即可生成你自己的文档）
 * ------------------------------------------------------------------ */

function buildSampleContent(F, assetsDir) {
  const children = [];

  children.push(F.h1("一、研究背景与意义"));
  children.push(F.bodyPara(
    "随着信息技术的快速发展，规范化学术文档的排版质量日益受到重视。本文档用于演示一套可复用、可配置的排版规范，覆盖正文、四级标题、图、表与参考文献等核心要素，确保生成结果在 Word 中保持一致的视觉表现。"
  ));
  children.push(F.bodyPara(
    "其中，图表标题的居中问题是历史生成方案的主要痛点之一。本方案通过将图题/表题改用单倍行距并显式置零缩进，从根源上消除了因继承正文 1.5 倍行距而导致的标题上下漂移与视觉不居中现象。"
  ));

  children.push(F.h2("（一）图的排版规范"));
  children.push(F.bodyPara(
    "图片统一宽度等于正文宽度，高度按原始宽高比等比缩放，超高时以高度上限反向缩放；图片水平居中、不加边框，图题置于图下方，五号加粗居中，来源或注释以小五号左对齐排列。"
  ));
  children.push(...F.figure({
    imagePath: path.join(assetsDir, "figure1.png"),
    title: "各季度销售额对比",
    source: "数据来源：示例数据，仅用于排版演示。",
  }));

  children.push(F.h2("（二）表的排版规范"));
  children.push(F.bodyPara(
    "表题置于表上方，五号加粗居中；表格整体居中并占满正文宽度，单元格内容垂直且水平居中；注释或来源置于表下方，小五号左对齐。"
  ));
  children.push(...F.table({
    title: "主要指标年度对比",
    headers: ["指标", "2023 年", "2024 年", "同比增长"],
    rows: [
      ["营业收入（万元）", "1,200", "1,500", "25.0%"],
      ["净利润（万元）", "180", "225", "25.0%"],
      ["研发投入（万元）", "96", "135", "40.6%"],
    ],
    note: "注：以上数据为示例，不构成任何投资建议。",
  }));

  // 四级标题演示
  children.push(F.h2("（三）四级标题演示"));
  children.push(F.h3("1. 三级标题示例"));
  children.push(F.bodyPara("三级标题采用黑体四号，用于更细的章节划分。下方演示四级标题。"));
  children.push(F.h4("(1) 四级标题示例"));
  children.push(F.bodyPara("四级标题采用黑体小四。autoHeading() 还可根据“一、/（一）/1./(1)”前缀自动识别层级。"));

  children.push(F.h1("二、参考文献著录规范"));
  children.push(F.bodyPara(
    "参考文献遵循 GB/T 7714 顺序编码制：正文引用以 [n] 编号，逾 3 位作者用“等”或“et al”；强制标注文献类型标识，电子文献使用复合标识（如 [J/OL]），网络资源须含引用日期。"
  ));
  children.push(...F.references([
    { authors: "王建国, 李明, 张华, 等", title: "现代排版工程", type: "M", place: "北京", publisher: "科学出版社", year: "2022" },
    { authors: "LIU Y, CHEN X, ZHAO M, et al", title: "A study on document formatting consistency", type: "J", source: "Journal of Publishing Science", year: "2023", volume: "15", issue: "3", pages: "45-52" },
    { authors: "陈晓", title: "学术写作中的图表规范研究", type: "D", place: "上海", publisher: "复旦大学", year: "2021" },
    { authors: "全国信息与文献标准化技术委员会", title: "信息与文献 参考文献著录规则: GB/T 7714—2015", type: "S", place: "北京", publisher: "中国标准出版社", year: "2015" },
    { authors: "张伟, 刘洋", title: "电子文献引用规范的新进展", type: "J/OL", source: "中国科技期刊研究", year: "2024", volume: "35", issue: "2", pages: "88-95", url: "https://doi.org/10.1234/example.2024.02.012", accessDate: "2026-06-28" },
    { authors: "中华人民共和国国家统计局", title: "2024 年国民经济和社会发展统计公报", type: "EB/OL", url: "https://www.stats.gov.cn/example.htm", accessDate: "2026-06-28" },
  ]));

  return children;
}

/* ------------------------------------------------------------------
 * 文档装配（样式 / 页码 / 页边距）
 * ------------------------------------------------------------------ */

function buildDocument(config, children, meta = {}) {
  const { w, h } = pageSizeDxa(config.page.size);
  const m = config.page.margins;
  const unit = config.page.marginUnit === "cm" ? CM_TO_DXA : 1;
  const margins = {
    top: Math.round(m.top * unit), bottom: Math.round(m.bottom * unit),
    left: Math.round(m.left * unit), right: Math.round(m.right * unit),
  };

  // 页脚页码
  const pn = config.pageNumber || {};
  const pnFont = makeFont(pn.font || "宋体", pn.asciiFont || "Times New Roman");
  const footers = (pn.enabled !== false) ? {
    default: new Footer({
      children: [new Paragraph({
        alignment: AlignmentType[pn.align ? pn.align.toUpperCase() : "CENTER"] || AlignmentType.CENTER,
        spacing: lineSpacingConfig(1),
        children: [new TextRun({ children: [PageNumber.CURRENT], font: pnFont, size: resolveSize(pn.size) })],
      })],
    }),
  } : undefined;

  // 标题样式（用 default.heading1~4 覆盖内置，避免蓝色默认）
  const headingStyles = {};
  for (let lv = 1; lv <= 4; lv++) {
    const hc = config.headings["h" + lv];
    if (!hc) continue;
    headingStyles["heading" + lv] = {
      run: {
        font: makeFont(hc.font, hc.asciiFont),
        size: resolveSize(hc.size), bold: hc.bold !== false, color: hc.color || "000000",
      },
      paragraph: {
        spacing: {
          ...lineSpacingConfig(hc.lineSpacing),
          before: Math.round((hc.spaceBeforePt || 0) * 20),
          after: Math.round((hc.spaceAfterPt || 0) * 20),
        },
        outlineLevel: lv - 1,
      },
    };
  }

  const bodyFont = makeFont(config.body.font, config.body.asciiFont);
  return new Document({
    creator: meta.creator || "规范文档生成器",
    title: meta.title || "规范文档",
    styles: {
      default: {
        document: {
          run: { font: bodyFont, size: resolveSize(config.body.size), color: config.body.color || "000000" },
          paragraph: { spacing: lineSpacingConfig(config.body.lineSpacing) },
        },
        ...headingStyles,
      },
    },
    sections: [{
      properties: { page: { size: { width: w, height: h }, margin: margins } },
      ...(footers ? { footers } : {}),
      children,
    }],
  });
}

/* ------------------------------------------------------------------
 * 主流程
 * ------------------------------------------------------------------ */

function printUsage() {
  console.log(`规范文档生成器
用法: node generate.js [选项]

选项:
  -c, --config <路径>     指定配置文件（默认 config.json）
  -s, --set a.b=c         临时覆盖配置项，可重复（如 --set page.margins.left=2.8）
      --out <路径>        输出文件名（默认取 config.output.filename）
      --list-sizes        打印中文字号对照表
      --show-config       打印生效后的配置
      --dry-run           预览会生成什么，不写文件
  -h, --help              显示帮助

示例:
  node generate.js
  node generate.js --set body.size=14 --set body.lineSpacing=2.0
  node generate.js --config 公文配置.json --out 公文.docx
  node generate.js --set table.captionPosition=below
  node generate.js --dry-run
`);
}

function printSizes() {
  console.log("中文字号对照表（可用作 config 中所有 size 字段的值）：");
  for (const [name, pt] of Object.entries(SIZE_NAME_TO_PT)) {
    console.log("  " + name + " = " + pt + " pt");
  }
  console.log("（也可直接写磅值数字，如 12 表示 12pt）");
}

function main() {
  const args = parseArgs(process.argv.slice(2));

  if (args.help) { printUsage(); return; }
  if (args.listSizes) { printSizes(); return; }

  // 读取配置
  const cfgPath = path.resolve(args.config || "config.json");
  if (!fs.existsSync(cfgPath)) {
    console.error("找不到配置文件: " + cfgPath);
    console.error("请先创建 config.json，或用 --config 指定路径。");
    process.exit(1);
  }
  let config;
  try { config = JSON.parse(fs.readFileSync(cfgPath, "utf8")); }
  catch (e) { console.error("配置文件 JSON 解析失败: " + e.message); process.exit(1); }

  // 应用 CLI 覆盖
  if (args.set.length) applyOverrides(config, args.set);

  if (args.showConfig) {
    console.log("生效配置（已应用 --set 覆盖）：");
    console.log(JSON.stringify(config, null, 2));
    return;
  }

  // 构造格式化器与内容
  const F = createFormatter(config);
  const assetsDir = path.join(__dirname, "assets");
  const children = buildSampleContent(F, assetsDir);

  // dry-run 模式：预览会生成什么，不写文件
  if (args.dryRun) {
    console.log("[DRY-RUN] 不会修改任何文件\n");
    let paraCount = 0, tableCount = 0, otherCount = 0;
    for (const el of children) {
      const ctor = el.constructor && el.constructor.name;
      if (ctor === "Paragraph") paraCount++;
      else if (ctor === "Table") tableCount++;
      else otherCount++;
    }
    console.log("汇总：");
    console.log("  · 段落: " + paraCount);
    console.log("  · 表格: " + tableCount);
    if (otherCount > 0) console.log("  · 其它: " + otherCount);
    console.log("\n完成。");
    return;
  }

  // 装配并输出
  const doc = buildDocument(config, children, { title: "规范文档示例" });
  const outFile = path.resolve(args.out || (config.output && config.output.filename) || "规范文档示例.docx");

  Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(outFile, buf);
    console.log("已生成: " + outFile);
    console.log("正文宽度: " + F.contentWidth + " DXA ≈ " + (F.contentWidth / CM_TO_DXA).toFixed(2) + " cm");
  }).catch(e => { console.error("生成失败: " + e.message); process.exit(1); });
}

main();
