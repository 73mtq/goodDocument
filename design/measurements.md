# macOS 系统设置结构量测 → ui_theme 令牌（v0.8）

标本：macOS「系统设置」（官方截图，量法与来源见 apple-measurements.md，@2x 像素 ÷2 = pt）。
量结构不量颜色：行高、列宽、尾部控件对齐、分组标题、分隔线节奏。
pt → px 换算按 tkinter 96dpi：px = pt × 4/3，再取整到 4px 网格。

## 量得值与换算

| 元素（系统设置） | 量得值（@2x → pt） | 换算 px（×4/3） | 落 4 网格 → 令牌 |
|---|---|---|---|
| 内容卡行距（相邻分隔线间距，即行高节奏） | 67-68px → 34pt | 45.3 | **44 `row_height`** |
| 侧栏行高 | 50px → 25pt | 33.3 | 32 `group_header_height` |
| 弹窗按钮 bezel 高 | 15px/37px 级 → 19-20pt | 26 | **28 `control_height`** |
| 分组卡片 | 白底浮在 #F7F7F7 内容底上 | - | 组卡白 `surface`，面板底 `bg` |
| 分组分隔（卡与卡） | 约 18pt 间隙 | 24 | **24 `group_gap`** |
| 分隔线 | 1px，#E4E4E4 一带，从文字起缩进 | 1 | 1px `border`，左缩进 44 `separator_inset` |
| 文字阶梯（HIG Typography） | Title2 17 / Body 13 / Subheadline 11 | 同 v0.7 | 字体令牌不变 |
| 控件列 | 右缘对齐、宽度随类型 | - | 控件统一右对齐；下拉/步进/开关各自定宽 |
| 名称列 | 标签列占卡宽约 40% | - | `name_column` 128（16px 图标 + 8 间距 + 6-7 个汉字） |
| 名称文字起点 | - | - | `indent_step` 36（12+16+8）；组头 `indent_group` 40（12+20+8） |

图标尺寸依微软官方字体文档建议档位（16/20/24/32/40/48）：
组头 20 `icon_group`、行内 16 `icon_row`、拖放区主图 40 `icon_hero`。

## 落成令牌（ui_theme.MEASUREMENTS）

```python
MEASUREMENTS = {
    "row_height": 44,          # 设置行高（系统设置行距 34pt）
    "group_header_height": 32, # 组头高（侧栏行 25pt）
    "control_height": 28,      # 尾部控件高（bezel 19-20pt）
    "panel_width": 480,        # 面板宽（原 460 取 4 网格）
    "name_column": 128,        # 行名称列宽
    "group_gap": 24,           # 组卡间距
    "separator_inset": 44,     # 行分隔线左缩进（图标 16 + 12 + 16）
    "drop_height": 96,         # 拖放区高（96 = 2 行 44 + 8）
    "wrap_main": 760,          # 主卡说明折行宽
    "wrap_panel": 440,         # 面板说明折行宽
    "icon_group": 20,
    "icon_row": 16,
    "icon_hero": 40,
}
```

密度口径：「密而整齐是节奏，不是留白」——行高 44 固定、列缘对齐、分隔线规律出现，
组内不加额外行距；紧凑感来自 13pt 正文配 44px 行高（约 2.4 倍文字高度）与
1px 分隔线的重复，不来自压缩行高。

## 配色调整（结构性，不新增颜色）

- 高级面板底由 #EDEDED 改 `bg` #F7F7F7，组卡 `surface` 白，向系统设置内容区看齐；
- 分段按钮选中态弃用 `accent_soft`（强调蓝只留给「开始整理」与聚焦态），
  改 `surface_pressed` 底 + 加粗近黑字；
- 菜单高亮、文本选区弃用蓝色系，改中性灰（`surface_hover` / `surface_pressed`）。
