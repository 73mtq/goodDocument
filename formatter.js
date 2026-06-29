/**
 * formatter.js —— 规范文档排版核心库（配置驱动）
 * =================================================================
 * 依据传入的 config 对象，构造符合规范的 docx 段落/表格/图片/参考文献。
 * 所有字号、字体、行距、边距、对齐等均由 config 控制，便于外部调整。
 *
 * 导出：
 *   resolveSize(val)                 字号名/磅值 → docx 半磅值
 *   makeFont(zh, en)                 构造 {ascii,eastAsia,hAnsi,cs}
 *   lineSpacingConfig(ratio)         行距倍数 → docx spacing 行距对象
 *   detectHeadingLevel(text, cfg)    自动识别“一、/（一）/1./(1)”标题层级
 *   createFormatter(config)          返回 h1~h4 / bodyPara / figure / table / references
 */

const fs = require("fs");
const sizeOf = require("image-size");
const {
  Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Footer, AlignmentType, BorderStyle, WidthType, VerticalAlign,
  PageNumber, HeadingLevel,
} = require("docx");

/* ------------------------------------------------------------------
 * 常量
 * ------------------------------------------------------------------ */

const CM_TO_DXA = 566.929;            // 1 cm = 1440/2.54 DXA
const PT_TO_DXA = 20;                 // 1 pt = 20 DXA
const PX_PER_CM = 96 / 2.54;          // 1 cm ≈ 37.795 px（96 DPI）

// 中文字号名 → 磅值
const SIZE_NAME_TO_PT = {
  "初号": 42, "小初": 36,
  "一号": 26, "小一": 24,
  "二号": 22, "小二": 18,
  "三号": 16, "小三": 15,
  "四号": 14, "小四": 12,
  "五号": 10.5, "小五": 9,
  "六号": 7.5, "小六": 6.5,
  "七号": 5.5, "八号": 5,
};

const ALIGN_MAP = {
  left: AlignmentType.LEFT,
  right: AlignmentType.RIGHT,
  center: AlignmentType.CENTER,
  justify: AlignmentType.JUSTIFIED,
  both: AlignmentType.JUSTIFIED,
};

/* ------------------------------------------------------------------
 * 工具函数
 * ------------------------------------------------------------------ */

/** 字号解析：中文字号名(小四) 或 磅值数字(12) → docx 半磅值(24) */
function resolveSize(val) {
  if (val == null) return 24;
  if (typeof val === "number") return Math.round(val * 2);
  const s = String(val).trim();
  if (SIZE_NAME_TO_PT[s] != null) return Math.round(SIZE_NAME_TO_PT[s] * 2);
  const n = parseFloat(s);
  if (!isNaN(n)) return Math.round(n * 2);
  return 24; // 默认小四
}

/** 构造字体对象：中文 eastAsia + 英数 ascii/hAnsi */
function makeFont(zh, en) {
  return { ascii: en || "Times New Roman", eastAsia: zh || "宋体", hAnsi: en || "Times New Roman", cs: en || "Times New Roman" };
}

/** 行距倍数 → docx spacing 的 line/lineRule（1.5 倍=360，单倍=240） */
function lineSpacingConfig(ratio) {
  return { line: Math.round(240 * (ratio || 1)), lineRule: "auto" };
}

/** 磅 → DXA（用于段前段后、缩进） */
const ptToDxa = (pt) => Math.round(pt * PT_TO_DXA);

/** 页面尺寸：A4 / Letter，返回 {w,h} DXA */
function pageSizeDxa(size) {
  if (/letter/i.test(String(size))) return { w: 12240, h: 15840 };
  return { w: 11906, h: 16838 }; // A4 默认
}

/** 计算正文可用宽度（页宽 − 左右边距）DXA */
function contentWidthDxa(cfg) {
  const { w } = pageSizeDxa(cfg.page.size);
  const m = cfg.page.margins;
  const unit = cfg.page.marginUnit === "cm" ? CM_TO_DXA : 1;
  return Math.round(w - m.left * unit - m.right * unit);
}

/* ------------------------------------------------------------------
 * 标题层级自动识别
 * ------------------------------------------------------------------ */

const HEADING_PATTERNS = {
  h1: /^[一二三四五六七八九十百]+、/,
  h2: /^（[一二三四五六七八九十百]+）|^\([一二三四五六七八九十百]+\)/,
  h3: /^\d+[.、]/,
  h4: /^（\d+）|^\(\d+\)/,
};

/**
 * 根据文本前缀自动判断标题层级（1~4），非标题返回 0。
 */
function detectHeadingLevel(text, cfg) {
  if (!text) return 0;
  const t = String(text).trim();
  // 用户自定义 patterns 优先
  const pats = (cfg && cfg.headingDetect && cfg.headingDetect.patterns) || HEADING_PATTERNS;
  for (const lvl of ["h1", "h2", "h3", "h4"]) {
    const p = pats[lvl] || HEADING_PATTERNS[lvl];
    try { if (new RegExp(p).test(t)) return Number(lv.slice(1)); } catch (_) {}
  }
  // 兜底内置
  for (const [lvl, re] of Object.entries(HEADING_PATTERNS)) {
    if (re.test(t)) return Number(lvl.slice(1));
  }
  return 0;
}

/* ------------------------------------------------------------------
 * createFormatter(config) —— 返回所有构造函数
 * ------------------------------------------------------------------ */

function createFormatter(config) {
  const cfg = config;
  const CW = contentWidthDxa(cfg);

  /* ---- 字体/字号快捷取值 ---- */
  const bodyFont = makeFont(cfg.body.font, cfg.body.asciiFont);
  const bodySize = resolveSize(cfg.body.size);

  /* ---- 段落 run 工具 ---- */
  function bodyRun(text, extra = {}) {
    return new TextRun({ text, font: bodyFont, size: bodySize, color: cfg.body.color || "000000", ...extra });
  }
  function headRun(text, hcfg) {
    return new TextRun({
      text,
      font: makeFont(hcfg.font, hcfg.asciiFont),
      size: resolveSize(hcfg.size),
      bold: hcfg.bold !== false,
      color: hcfg.color || "000000",
    });
  }

  /* ---- 正文段落 ---- */
  function bodyPara(text, extra = {}) {
    const indentChars = cfg.body.firstLineIndentChars || 0;
    return new Paragraph({
      alignment: ALIGN_MAP[cfg.body.align] || AlignmentType.JUSTIFIED,
      spacing: { ...lineSpacingConfig(cfg.body.lineSpacing), after: ptToDxa(cfg.body.spaceAfterPt || 0) },
      indent: indentChars > 0 ? { firstLine: indentChars * bodySize } : undefined,
      children: [bodyRun(text)],
      ...extra,
    });
  }

  /* ---- 标题 h1~h4 ---- */
  function heading(level, text) {
    const key = "h" + level;
    const hcfg = cfg.headings[key];
    if (!hcfg) throw new Error("配置缺少 headings." + key);
    const hl = level === 1 ? HeadingLevel.HEADING_1
      : level === 2 ? HeadingLevel.HEADING_2
      : level === 3 ? HeadingLevel.HEADING_3 : HeadingLevel.HEADING_4;
    return new Paragraph({
      heading: hl,
      alignment: ALIGN_MAP[hcfg.align] || AlignmentType.LEFT,
      spacing: {
        ...lineSpacingConfig(hcfg.lineSpacing),
        before: ptToDxa(hcfg.spaceBeforePt || 0),
        after: ptToDxa(hcfg.spaceAfterPt || 0),
      },
      children: [headRun(text, hcfg)],
    });
  }
  const h1 = (t) => heading(1, t);
  const h2 = (t) => heading(2, t);
  const h3 = (t) => heading(3, t);
  const h4 = (t) => heading(4, t);

  /** 自动识别层级并生成标题段；非标题则作为正文段落 */
  function autoHeading(text) {
    const lvl = detectHeadingLevel(text, cfg);
    return lvl > 0 ? heading(lvl, text) : bodyPara(text);
  }

  /* ---- 图 ---- */
  function computeImageTransform(imgPath) {
    const dim = sizeOf(imgPath);
    const ratio = dim.width / dim.height;
    let w = cfg.figure.maxWidthCm * PX_PER_CM;
    let h = w / ratio;
    const maxH = cfg.figure.maxHeightCm * PX_PER_CM;
    if (h > maxH) { h = maxH; w = h * ratio; }
    const type = (dim.type === "jpg" || dim.type === "jpeg") ? "jpg" : "png";
    return { width: Math.round(w), height: Math.round(h), type };
  }

  let _figNo = 0, _tabNo = 0;
  const nextFigNo = () => ++_figNo;
  const nextTabNo = () => ++_tabNo;
  function resetCounters() { _figNo = 0; _tabNo = 0; }

  function figure({ imagePath, title, source }) {
    const f = cfg.figure;
    const no = nextFigNo();
    const tf = computeImageTransform(imagePath);
    const capFont = makeFont(f.captionText, f.captionAsciiFont || cfg.body.asciiFont);
    const out = [];

    // 图片：居中、无边框
    out.push(new Paragraph({
      alignment: ALIGN_MAP[f.align] || AlignmentType.CENTER,
      spacing: { ...lineSpacingConfig(1), before: 120, after: 60 },
      children: [new ImageRun({
        type: tf.type,
        data: fs.readFileSync(imagePath),
        transformation: { width: tf.width, height: tf.height },
        altText: { title: `${f.prefix}${no}`, description: title, name: `${f.prefix}${no}` },
      })],
    }));

    const capPara = new Paragraph({
      alignment: ALIGN_MAP[f.captionAlign] || AlignmentType.CENTER,
      spacing: { ...lineSpacingConfig(f.captionLineSpacing), before: 0, after: 40 },
      indent: { left: 0, right: 0, firstLine: 0 },
      children: [new TextRun({
        text: `${f.prefix}${no}  ${title}`, font: capFont,
        size: resolveSize(f.captionSize), bold: f.captionBold !== false,
      })],
    });

    const notePara = (src) => new Paragraph({
      alignment: ALIGN_MAP[f.noteAlign] || AlignmentType.LEFT,
      spacing: { ...lineSpacingConfig(f.noteLineSpacing), before: 0, after: 180 },
      indent: { left: 0, right: 0, firstLine: 0 },
      children: [new TextRun({ text: src, font: capFont, size: resolveSize(f.noteSize) })],
    });

    // 图题位置：below（图下）/ above（图上）
    // below → 图片→图题→注释；above → 图题→图片→注释
    if (f.captionPosition === "above") {
      out.unshift(capPara);              // [图题, 图片]
      if (source) out.push(notePara(source)); // [图题, 图片, 注释]
    } else {
      out.push(capPara);                 // [图片, 图题]
      if (source) out.push(notePara(source)); // [图片, 图题, 注释]
    }
    return out;
  }

  /* ---- 表 ---- */
  function table({ title, headers, rows, note, columnWidths }) {
    const t = cfg.table;
    const no = nextTabNo();
    const capFont = makeFont(t.captionText, t.captionAsciiFont || cfg.body.asciiFont);
    const cellFont = makeFont(t.cellFont || cfg.body.font, t.cellAsciiFont || cfg.body.asciiFont);
    const ncol = headers.length;

    // 列宽
    const widths = (columnWidths && columnWidths.length === ncol)
      ? [...columnWidths] : Array(ncol).fill(Math.floor(CW / ncol));
    const diff = CW - widths.reduce((a, b) => a + b, 0);
    widths[widths.length - 1] += diff;

    const borderDef = { style: BorderStyle.SINGLE, size: t.borderSize || 4, color: t.borderColor || "000000" };
    const cellBorders = { top: borderDef, bottom: borderDef, left: borderDef, right: borderDef };

    function cellPara(text, bold) {
      return new Paragraph({
        alignment: ALIGN_MAP[t.cellAlign] || AlignmentType.CENTER,
        spacing: { ...lineSpacingConfig(t.cellLineSpacing), before: 30, after: 30 },
        children: [new TextRun({
          text: String(text), font: cellFont, size: resolveSize(t.cellSize), bold: !!bold,
        })],
      });
    }
    function makeCell(text, width, bold) {
      return new TableCell({
        borders: cellBorders,
        width: { size: width, type: WidthType.DXA },
        verticalAlign: VerticalAlign[t.cellVAlign ? t.cellVAlign.toUpperCase() : "CENTER"] || VerticalAlign.CENTER,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [cellPara(text, bold)],
      });
    }

    const capPara = new Paragraph({
      alignment: ALIGN_MAP[t.captionAlign] || AlignmentType.CENTER,
      // 【关键修复】表题单倍行距 + 零缩进，杜绝继承正文 1.5 倍行距导致的居中漂移
      spacing: { ...lineSpacingConfig(t.captionLineSpacing), before: 180, after: 80 },
      indent: { left: 0, right: 0, firstLine: 0 },
      children: [new TextRun({
        text: `${t.prefix}${no}  ${title}`, font: capFont,
        size: resolveSize(t.captionSize), bold: t.captionBold !== false,
      })],
    });

    const notePara = (src) => new Paragraph({
      alignment: ALIGN_MAP[t.noteAlign] || AlignmentType.LEFT,
      spacing: { ...lineSpacingConfig(t.noteLineSpacing), before: 60, after: 180 },
      indent: { left: 0, right: 0, firstLine: 0 },
      children: [new TextRun({ text: src, font: capFont, size: resolveSize(t.noteSize) })],
    });

    const trs = [];
    trs.push(new TableRow({
      tableHeader: true,
      children: headers.map((h, i) => makeCell(h, widths[i], t.headerBold !== false)),
    }));
    (rows || []).forEach(r => {
      trs.push(new TableRow({ children: r.map((c, i) => makeCell(c, widths[i], false)) }));
    });

    const tbl = new Table({
      alignment: ALIGN_MAP[t.align] || AlignmentType.CENTER,
      width: { size: CW, type: WidthType.DXA },
      columnWidths: widths,
      borders: { ...cellBorders, insideHorizontal: borderDef, insideVertical: borderDef },
      rows: trs,
    });

    const out = [];
    if (t.captionPosition === "above") {
      out.push(capPara, tbl);
      if (note) out.push(notePara(note));
    } else {
      out.push(tbl, capPara);
      if (note) out.push(notePara(note));
    }
    return out;
  }

  /* ---- 参考文献 GB/T 7714 ---- */
  function references(items) {
    const r = cfg.references;
    const font = makeFont(r.font, r.asciiFont);
    const sz = resolveSize(r.size);
    const hang = ptToDxa(r.hangingIndentPt || 24);
    return items.map((it, i) => {
      const n = i + 1;
      let s = `[${n}] ${it.authors}. ${it.title}[${it.type}]. `;
      if (it.place && it.publisher) {
        s += `${it.place}: ${it.publisher}, ${it.year}.`;
      } else if (it.source) {
        s += `${it.source}, ${it.year}`;
        if (it.volume) s += `, ${it.volume}`;
        if (it.issue) s += `(${it.issue})`;
        s += `: ${it.pages || ""}.`;
      } else {
        s += `${it.year || ""}.`;
      }
      if (it.url) { s += ` ${it.url}`; if (it.accessDate) s += ` [${it.accessDate}].`; }
      return new Paragraph({
        alignment: ALIGN_MAP[r.align] || AlignmentType.JUSTIFIED,
        spacing: { ...lineSpacingConfig(r.lineSpacing), before: 0, after: 60 },
        indent: { left: hang, hanging: hang },
        children: [new TextRun({ text: s, font, size: sz })],
      });
    });
  }

  return {
    h1, h2, h3, h4, heading, autoHeading, bodyPara, bodyRun,
    figure, table, references,
    nextFigNo, nextTabNo, resetCounters,
    // 暴露内部常量供外部使用
    contentWidth: CW, bodyFont, bodySize,
  };
}

module.exports = {
  resolveSize, makeFont, lineSpacingConfig, detectHeadingLevel,
  pageSizeDxa, contentWidthDxa, createFormatter,
  CM_TO_DXA, PT_TO_DXA, SIZE_NAME_TO_PT,
};
