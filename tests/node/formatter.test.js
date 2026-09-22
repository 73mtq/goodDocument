const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { Packer, Document } = require("docx");
const {
  resolveSize, detectHeadingLevel, createFormatter, pageSizeDxa, contentWidthDxa,
} = require("../../formatter.js");

const ROOT = path.join(__dirname, "..", "..");
const config = JSON.parse(fs.readFileSync(path.join(ROOT, "config.json"), "utf8"));

async function documentXml(doc) {
  const buf = await Packer.toBuffer(doc);
  const zip = await JSZip.loadAsync(buf);
  return zip.file("word/document.xml").async("string");
}

test("resolveSize 支持中文字号与磅值", () => {
  assert.equal(resolveSize("小四"), 24);
  assert.equal(resolveSize("三号"), 32);
  assert.equal(resolveSize(12), 24);
  assert.equal(resolveSize(null), 24);
});

test("detectHeadingLevel 识别内置规则并忽略段首小数", () => {
  assert.equal(detectHeadingLevel("一、引言", config), 1);
  assert.equal(detectHeadingLevel("（一）研究背景", config), 2);
  assert.equal(detectHeadingLevel("1. 三级标题", config), 3);
  assert.equal(detectHeadingLevel("（1）四级标题", config), 4);
  assert.equal(detectHeadingLevel("1.5 倍行距说明", config), 0);
});

test("detectHeadingLevel 支持自定义规则", () => {
  const cfg = { headingDetect: { patterns: { h1: "^第.+章" } } };
  assert.equal(detectHeadingLevel("第一章 引言", cfg), 1);
});

test("bodyPara 首行缩进按 twips 输出", async () => {
  const F = createFormatter(config);
  const doc = new Document({ sections: [{ children: [F.bodyPara("测试段落")] }] });
  const xml = await documentXml(doc);
  assert.match(xml, /<w:ind[^>]*w:firstLine="480"/);
});

test("table 生成表头与数据行", async () => {
  const F = createFormatter(config);
  const out = F.table({
    title: "测试表",
    headers: ["A", "B"],
    rows: [["1", "2"], ["3", "4"]],
  });
  assert.equal(out[0].constructor.name, "Paragraph");
  assert.equal(out[1].constructor.name, "Table");
  const doc = new Document({ sections: [{ children: out }] });
  const xml = await documentXml(doc);
  assert.equal((xml.match(/<w:tr[ >]/g) || []).length, 3);
});

test("references 生成悬挂缩进", async () => {
  const F = createFormatter(config);
  const out = F.references([
    { authors: "张三", title: "排版测试", type: "J", source: "示例期刊", year: "2024" },
  ]);
  const doc = new Document({ sections: [{ children: out }] });
  const xml = await documentXml(doc);
  assert.match(xml, /<w:ind[^>]*w:hanging="480"/);
});

test("pageSizeDxa 与 contentWidthDxa 按页面设置计算", () => {
  assert.deepEqual(pageSizeDxa("Letter"), { w: 12240, h: 15840 });
  assert.deepEqual(pageSizeDxa("A4"), { w: 11906, h: 16838 });
  const expected = Math.round(
    11906 - config.page.margins.left * 566.929 - config.page.margins.right * 566.929
  );
  assert.equal(contentWidthDxa(config), expected);
});

/* ---- 图片尺寸策略（figure.sizeMode），与 Python 端语义保持一致 ---- */

const EMU_PER_CM = 360000;
const FIGURE_PNG = path.join(ROOT, "assets", "figure1.png");

async function figureSizeCm(overrides) {
  const cfg = JSON.parse(JSON.stringify(config));
  Object.assign(cfg.figure, overrides);
  const F = createFormatter(cfg);
  const out = F.figure({ imagePath: FIGURE_PNG, title: "测试图" });
  const xml = await documentXml(new Document({ sections: [{ children: out }] }));
  const extent = xml.match(/<wp:extent cx="(\d+)" cy="(\d+)"/);
  const xfrm = xml.match(/<a:ext cx="(\d+)"/);
  return {
    widthCm: Number(extent[1]) / EMU_PER_CM,
    heightCm: Number(extent[2]) / EMU_PER_CM,
    xfrmWidthCm: Number(xfrm[1]) / EMU_PER_CM,
  };
}

const closeTo = (actual, expected) => Math.abs(actual - expected) < 0.05;

test("fit 模式：宽度受 maxWidthCm 限制且等比", async () => {
  const size = await figureSizeCm({ sizeMode: "fit" });
  assert.ok(closeTo(size.widthCm, 15.5), `实际宽度 ${size.widthCm}`);
  assert.ok(closeTo(size.heightCm, 15.5 * 2 / 3), `实际高度 ${size.heightCm}`);
});

test("fixed 模式：不保持比例时强制拉伸到指定尺寸", async () => {
  const size = await figureSizeCm({
    sizeMode: "fixed", fixedWidthCm: 12, fixedHeightCm: 9, keepAspectRatio: false,
  });
  assert.ok(closeTo(size.widthCm, 12), `实际宽度 ${size.widthCm}`);
  assert.ok(closeTo(size.heightCm, 9), `实际高度 ${size.heightCm}`);
});

test("fixed 模式：只设宽度且保持比例即统一宽度", async () => {
  const size = await figureSizeCm({
    sizeMode: "fixed", fixedWidthCm: 10, fixedHeightCm: null, keepAspectRatio: true,
  });
  assert.ok(closeTo(size.widthCm, 10), `实际宽度 ${size.widthCm}`);
  assert.ok(closeTo(size.heightCm, 10 * 2 / 3), `实际高度 ${size.heightCm}`);
});

test("fixed 模式：目标宽度受正文可用宽度钳制", async () => {
  const size = await figureSizeCm({
    sizeMode: "fixed", fixedWidthCm: 20, fixedHeightCm: null, keepAspectRatio: true,
  });
  assert.ok(closeTo(size.widthCm, 15.5), `实际宽度 ${size.widthCm}`);
});

test("图片尺寸同时写入 wp:extent 与 a:xfrm/a:ext", async () => {
  const size = await figureSizeCm({
    sizeMode: "fixed", fixedWidthCm: 12, fixedHeightCm: 9, keepAspectRatio: false,
  });
  assert.ok(closeTo(size.xfrmWidthCm, size.widthCm), "a:ext 与 wp:extent 不一致");
});
