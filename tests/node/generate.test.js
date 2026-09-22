const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { Packer } = require("docx");
const { parseArgs, parseValue, setByPath, buildDocument, main } = require("../../generate.js");
const { createFormatter } = require("../../formatter.js");

const ROOT = path.join(__dirname, "..", "..");
const config = JSON.parse(fs.readFileSync(path.join(ROOT, "config.json"), "utf8"));

test("parseArgs 解析常用参数", () => {
  const args = parseArgs(["--set", "body.size=14", "--out", "a.docx", "--dry-run"]);
  assert.deepEqual(args.set, ["body.size=14"]);
  assert.equal(args.out, "a.docx");
  assert.equal(args.dryRun, true);
  assert.equal(typeof main, "function");
});

test("parseValue 转换布尔与数字", () => {
  assert.equal(parseValue("true"), true);
  assert.equal(parseValue("false"), false);
  assert.equal(parseValue("12.5"), 12.5);
  assert.equal(parseValue("宋体"), "宋体");
});

test("setByPath 支持点号路径", () => {
  const cfg = {};
  setByPath(cfg, "page.margins.left", 3.5);
  assert.equal(cfg.page.margins.left, 3.5);
});

test("buildDocument 组装的文档含页面边距与页脚页码", async () => {
  const F = createFormatter(config);
  const children = [F.h1("一、标题"), F.bodyPara("正文内容")];
  const doc = buildDocument(config, children, { title: "测试" });
  const buf = await Packer.toBuffer(doc);
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file("word/document.xml").async("string");
  assert.match(xml, /<w:pgMar[^>]*w:left="1701"/);
  const footerName = Object.keys(zip.files).find((name) => /footer\d*\.xml$/.test(name));
  assert.ok(footerName, "应包含页脚部件");
  const footer = await zip.file(footerName).async("string");
  assert.match(footer, /PAGE/);
});
