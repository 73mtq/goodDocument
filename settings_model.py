"""高级设置的声明式目录与配置值转换。"""

from font_sizes import SIZE_NAME_TO_PT
from ui_theme import LABEL_VALUES, display_label

SIZE_OPTIONS = list(SIZE_NAME_TO_PT)
ALIGN_OPTIONS = ["left", "center", "right", "justify"]
CELL_ALIGN_OPTIONS = ALIGN_OPTIONS + ["smart"]
FONT_ZH_OPTIONS = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑"]
FONT_EN_OPTIONS = ["Times New Roman", "Arial", "Calibri", "Cambria"]
POSITION_OPTIONS = ["above", "below"]
VALIGN_OPTIONS = ["center", "top"]
SIZE_MODE_OPTIONS = ["fit", "fixed"]
FIGURE_PREFIX_OPTIONS = ["图", "Figure", "Fig"]
TABLE_PREFIX_OPTIONS = ["表", "Table"]


def _heading_rows(level):
    prefix = ("headings", "h" + str(level))
    return [
        {"name": "中文字体", "key": (*prefix, "font"), "kind": "enum", "options": FONT_ZH_OPTIONS},
        {"name": "英文字体", "key": (*prefix, "asciiFont"), "kind": "enum", "options": FONT_EN_OPTIONS},
        {"name": "字号", "key": (*prefix, "size"), "kind": "size", "options": SIZE_OPTIONS},
        {
            "name": "行距",
            "key": (*prefix, "lineSpacing"),
            "kind": "number",
            "unit": "倍",
            "lo": 1.0,
            "hi": 3.0,
            "step": 0.5,
        },
        {"name": "加粗", "key": (*prefix, "bold"), "kind": "bool"},
        {
            "name": "段前间距",
            "key": (*prefix, "spaceBeforePt"),
            "kind": "number",
            "unit": "磅",
            "lo": 0,
            "hi": 40,
            "step": 1,
        },
        {
            "name": "段后间距",
            "key": (*prefix, "spaceAfterPt"),
            "kind": "number",
            "unit": "磅",
            "lo": 0,
            "hi": 40,
            "step": 1,
        },
        {"name": "对齐", "key": (*prefix, "align"), "kind": "enum", "options": ALIGN_OPTIONS},
    ]


# 高级设置信息架构：按用户目的分 11 组（卡片分类过程见 design/ia-regroup.md）。
# kind: number(数字+单位) / bool(开关) / enum(下拉) / size(下拉+磅值对照)。
SETTING_GROUPS = [
    {
        "id": "page",
        "title": "页面",
        "icon": "page",
        "about": "纸张四边的空白宽度。",
        "rows": [
            {
                "name": "上边距",
                "key": ("page", "margins", "top"),
                "kind": "number",
                "unit": "厘米",
                "lo": 0,
                "hi": 15,
                "step": 0.5,
            },
            {
                "name": "下边距",
                "key": ("page", "margins", "bottom"),
                "kind": "number",
                "unit": "厘米",
                "lo": 0,
                "hi": 15,
                "step": 0.5,
            },
            {
                "name": "左边距",
                "key": ("page", "margins", "left"),
                "kind": "number",
                "unit": "厘米",
                "lo": 0,
                "hi": 15,
                "step": 0.5,
            },
            {
                "name": "右边距",
                "key": ("page", "margins", "right"),
                "kind": "number",
                "unit": "厘米",
                "lo": 0,
                "hi": 15,
                "step": 0.5,
            },
        ],
    },
    {
        "id": "pageNumber",
        "title": "页码",
        "icon": "dialpad",
        "about": "页码放哪、多大、用什么字体。",
        "rows": [
            {"name": "显示页码", "key": ("pageNumber", "enabled"), "kind": "bool"},
            {"name": "只更新已有页码", "key": ("pageNumber", "onlyUpdateExisting"), "kind": "bool"},
            {"name": "页码位置", "key": ("pageNumber", "align"), "kind": "enum", "options": ALIGN_OPTIONS},
            {"name": "页码大小", "key": ("pageNumber", "size"), "kind": "size", "options": SIZE_OPTIONS},
            {"name": "中文字体", "key": ("pageNumber", "font"), "kind": "enum", "options": FONT_ZH_OPTIONS},
            {"name": "英文字体", "key": ("pageNumber", "asciiFont"), "kind": "enum", "options": FONT_EN_OPTIONS},
        ],
    },
    {
        "id": "body",
        "title": "正文",
        "icon": "font",
        "about": "论文主体文字的样子。",
        "rows": [
            {"name": "中文字体", "key": ("body", "font"), "kind": "enum", "options": FONT_ZH_OPTIONS},
            {"name": "英文字体", "key": ("body", "asciiFont"), "kind": "enum", "options": FONT_EN_OPTIONS},
            {"name": "字号", "key": ("body", "size"), "kind": "size", "options": SIZE_OPTIONS},
            {
                "name": "行距",
                "key": ("body", "lineSpacing"),
                "kind": "number",
                "unit": "倍",
                "lo": 1.0,
                "hi": 3.0,
                "step": 0.5,
            },
            {
                "name": "段首缩进",
                "key": ("body", "firstLineIndentChars"),
                "kind": "number",
                "unit": "字",
                "lo": 0,
                "hi": 10,
                "step": 1,
            },
            {"name": "对齐", "key": ("body", "align"), "kind": "enum", "options": ALIGN_OPTIONS},
            {
                "name": "段后间距",
                "key": ("body", "spaceAfterPt"),
                "kind": "number",
                "unit": "磅",
                "lo": 0,
                "hi": 30,
                "step": 1,
            },
        ],
    },
    {
        "id": "h1",
        "title": "章标题（第1章）",
        "icon": "bold",
        "about": "一级标题（第1章）的样式。",
        "rows": _heading_rows(1),
    },
    {
        "id": "h2",
        "title": "节标题（1.1）",
        "icon": "bold",
        "about": "二级标题（1.1）的样式。",
        "rows": _heading_rows(2),
    },
    {
        "id": "h3",
        "title": "条标题（1.1.1）",
        "icon": "bold",
        "about": "三级标题（1.1.1）的样式。",
        "rows": _heading_rows(3),
    },
    {
        "id": "h4",
        "title": "四级标题（1.1.1.1）",
        "icon": "bold",
        "about": "四级标题（1.1.1.1）的样式。",
        "rows": _heading_rows(4),
    },
    {
        "id": "figure",
        "title": "图片",
        "icon": "photo",
        "about": "图片统一成多大。",
        "rows": [
            {"name": "图片大小方式", "key": ("figure", "sizeMode"), "kind": "enum", "options": SIZE_MODE_OPTIONS},
            {
                "name": "等比最大宽度",
                "key": ("figure", "maxWidthCm"),
                "kind": "number",
                "unit": "厘米",
                "lo": 5,
                "hi": 25,
                "step": 0.5,
            },
            {
                "name": "等比最大高度",
                "key": ("figure", "maxHeightCm"),
                "kind": "number",
                "unit": "厘米",
                "lo": 5,
                "hi": 30,
                "step": 0.5,
            },
            {
                "name": "统一宽度",
                "key": ("figure", "fixedWidthCm"),
                "kind": "number",
                "unit": "厘米",
                "lo": 1,
                "hi": 25,
                "step": 0.5,
            },
            {
                "name": "统一高度",
                "key": ("figure", "fixedHeightCm"),
                "kind": "number",
                "unit": "厘米",
                "lo": 1,
                "hi": 30,
                "step": 0.5,
            },
            {"name": "统一时保持比例", "key": ("figure", "keepAspectRatio"), "kind": "bool"},
            {"name": "宽度不超正文", "key": ("figure", "clampToTextWidth"), "kind": "bool"},
            {"name": "跳过表格内图片", "key": ("figure", "skipInTables"), "kind": "bool"},
            {"name": "图和图题不分开", "key": ("figure", "keepWithCaption"), "kind": "bool"},
        ],
    },
    {
        "id": "figureCaption",
        "title": "图题与说明",
        "icon": "caption",
        "about": "图下面「图1 …」和说明文字的样式。",
        "rows": [
            {"name": "图题位置", "key": ("figure", "captionPosition"), "kind": "enum", "options": POSITION_OPTIONS},
            {"name": "图题字号", "key": ("figure", "captionSize"), "kind": "size", "options": SIZE_OPTIONS},
            {"name": "图题加粗", "key": ("figure", "captionBold"), "kind": "bool"},
            {"name": "图题对齐", "key": ("figure", "captionAlign"), "kind": "enum", "options": ALIGN_OPTIONS},
            {
                "name": "图题行距",
                "key": ("figure", "captionLineSpacing"),
                "kind": "number",
                "unit": "倍",
                "lo": 1.0,
                "hi": 3.0,
                "step": 0.5,
            },
            {"name": "注释字号", "key": ("figure", "noteSize"), "kind": "size", "options": SIZE_OPTIONS},
            {"name": "注释对齐", "key": ("figure", "noteAlign"), "kind": "enum", "options": ALIGN_OPTIONS},
            {"name": "编号开头", "key": ("figure", "prefix"), "kind": "enum", "options": FIGURE_PREFIX_OPTIONS},
        ],
    },
    {
        "id": "table",
        "title": "表格",
        "icon": "table",
        "about": "表格、表题与表内文字的样式。",
        "rows": [
            {"name": "表题位置", "key": ("table", "captionPosition"), "kind": "enum", "options": POSITION_OPTIONS},
            {"name": "表题字号", "key": ("table", "captionSize"), "kind": "size", "options": SIZE_OPTIONS},
            {"name": "表题加粗", "key": ("table", "captionBold"), "kind": "bool"},
            {"name": "表题对齐", "key": ("table", "captionAlign"), "kind": "enum", "options": ALIGN_OPTIONS},
            {
                "name": "表题行距",
                "key": ("table", "captionLineSpacing"),
                "kind": "number",
                "unit": "倍",
                "lo": 1.0,
                "hi": 3.0,
                "step": 0.5,
            },
            {"name": "注释字号", "key": ("table", "noteSize"), "kind": "size", "options": SIZE_OPTIONS},
            {"name": "注释对齐", "key": ("table", "noteAlign"), "kind": "enum", "options": ALIGN_OPTIONS},
            {"name": "内容字号", "key": ("table", "cellSize"), "kind": "size", "options": SIZE_OPTIONS},
            {"name": "内容对齐", "key": ("table", "cellAlign"), "kind": "enum", "options": CELL_ALIGN_OPTIONS},
            {"name": "上下位置", "key": ("table", "cellVAlign"), "kind": "enum", "options": VALIGN_OPTIONS},
            {
                "name": "内容行距",
                "key": ("table", "cellLineSpacing"),
                "kind": "number",
                "unit": "倍",
                "lo": 1.0,
                "hi": 3.0,
                "step": 0.5,
            },
            {"name": "表头加粗", "key": ("table", "headerBold"), "kind": "bool"},
            {"name": "编号开头", "key": ("table", "prefix"), "kind": "enum", "options": TABLE_PREFIX_OPTIONS},
        ],
    },
    {
        "id": "references",
        "title": "参考文献",
        "icon": "library",
        "about": "参考文献列表的样式；条目缩进指第二条起相对第一条缩进的距离。",
        "rows": [
            {"name": "中文字体", "key": ("references", "font"), "kind": "enum", "options": FONT_ZH_OPTIONS},
            {"name": "英文字体", "key": ("references", "asciiFont"), "kind": "enum", "options": FONT_EN_OPTIONS},
            {"name": "字号", "key": ("references", "size"), "kind": "size", "options": SIZE_OPTIONS},
            {
                "name": "行距",
                "key": ("references", "lineSpacing"),
                "kind": "number",
                "unit": "倍",
                "lo": 1.0,
                "hi": 3.0,
                "step": 0.5,
            },
            {"name": "对齐", "key": ("references", "align"), "kind": "enum", "options": ALIGN_OPTIONS},
            {
                "name": "条目缩进",
                "key": ("references", "hangingIndentPt"),
                "kind": "number",
                "unit": "磅",
                "lo": 0,
                "hi": 60,
                "step": 1,
            },
        ],
    },
]


def spec_default(spec):
    if spec["kind"] == "bool":
        return False
    if spec["kind"] == "number":
        return spec.get("lo", 0)
    if "options" in spec:
        return spec["options"][0]
    return ""


def fmt_number(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if number == int(number):
        return str(int(number))
    return str(number)


def text_to_config_value(raw):
    """界面文本转成配置值：中文标签还原、数字转数值、其余保留字符串。"""
    if raw in LABEL_VALUES:
        return LABEL_VALUES[raw]
    try:
        number = float(raw)
    except (TypeError, ValueError):
        return raw
    return int(number) if number == int(number) else number


def humanize_value(spec, raw):
    """当前生效值的人话展示：附单位说明（磅/厘米/倍/字），字号补磅值对照。"""
    kind = spec["kind"]
    if kind == "bool":
        return "开" if raw else "关"
    if kind == "number":
        return fmt_number(raw) + " " + spec.get("unit", "")
    if kind == "size":
        text = str(raw)
        pt = SIZE_NAME_TO_PT.get(text)
        if pt is not None:
            return text + " · " + fmt_number(pt) + " 磅"
        return text
    return display_label(str(raw))


def raw_equal(spec, first, second):
    """两个配置取值是否等价（数字按值比较，枚举按字符串比较）。"""
    if first is None or second is None:
        return first is None and second is None
    if spec["kind"] == "number":
        try:
            return abs(float(first) - float(second)) < 1e-9
        except (TypeError, ValueError):
            return str(first) == str(second)
    if spec["kind"] == "bool":
        return bool(first) == bool(second)
    return str(first) == str(second)


def cfg_get(cfg, chain, default=None):
    current = cfg
    for key in chain:
        if not isinstance(current, dict) or key not in current:
            return default
        current = current[key]
    return current
