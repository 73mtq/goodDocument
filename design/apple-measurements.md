# Apple 系统色与规格测量记录（一手材料）

日期：2026-10-25。方法：官方 HIG JSON 原文 + Apple 官方支持页截图，用 PIL 写脚本
量色与量尺寸（design/ref/measure.py 等）。未使用任何博客或 AI 生成图。

## 来源（全部一手）

1. HIG Foundations · Color（developer.apple.com/tutorials/data/design/human-interface-guidelines/color.json）
2. HIG Foundations · Typography（同路径 typography.json）
3. 官方截图（support.apple.com 原文内嵌资产，@2x，pt = px ÷ 2）：
   - 系统设置全窗 1144×1112：ipcdn-web.apple.com/assets/v2/web/40bd24b5-d964-4cb9-8248-1d8bad524aba
     （出处 support.apple.com/guide/mac-help/change-system-settings-mh15217/mac）
   - 预览主窗 818×592：…/852f6fc5-175c-43cd-8376-6f3708f277f8
   - 预览「打开」对话框 756×424：…/991c06e6-1fd8-46e4-8eed-180bd2e6a38e
     （出处 support.apple.com/guide/preview/preview-basics-prvw1011/mac）

## HIG Color 原文要点（macOS）

- 系统色（Default Light，HIG 表格原文 RGB）：systemBlue 0/122/255 = #007AFF；
  systemRed 255/59/48；systemOrange 255/149/0；systemGreen 40/205/65；systemGray 142/142/147 = #8E8E93；
  可及版 systemBlue 0/64/221 = #0040DD。
- 原则：桌面 App「有节制地用色，少量颜色只为唤起对重要信息的注意」；同一颜色含义唯一；
  「蓝色被理解为可交互」；强调色概念（controlAccentColor）；避免让颜色与文字抢注意力。

## HIG Typography 原文要点（macOS 内建文字样式，单位 pt）

| 样式 | 字重/字号/行高 | Semibold 变体 |
|---|---|---|
| Large Title | Regular 26 / 32 | Bold |
| Title 1 | Regular 22 / 26 | Bold |
| Title 2 | Regular 17 / 22 | Bold |
| Title 3 | Regular 15 / 20 | Semibold |
| Headline | Bold 13 / 16 | Heavy |
| Body | Regular 13 / 16 | Semibold |
| Callout | Regular 12 / 15 | Semibold |
| Subheadline | Regular 11 / 14 | Semibold |
| Footnote | Regular 10 / 13 | Semibold |
| Caption 1 / 2 | Regular 10 / 13 | Medium / Semibold |

- 「在 macOS 中的目标是使用最小字号 10 pt」；避免 Light 字重（选 Regular/Medium/Semibold/Bold）；
  字体族尽量少。系统字体 SF Pro（Windows 现实替代：Microsoft YaHei UI，维持不变）。

## 官方截图量得值（@2x → pt）

系统设置（macOS 26 渲染，浅色）：

| 元素 | 量得值 |
|---|---|
| 内容区分组底 | #F7F7F7（占 47%；同族近值 #F0F0F0/#F6F6F6，半透明材质所致） |
| 分组卡片（grouped card） | #FFFFFF |
| 侧栏底 | #EDEDED（侧栏宽 340px = 170pt） |
| 分隔线 | #E4E4E4 / #DEDEDE（卡内行分隔）；控件描边 #D7D7D7、#BEBEBE |
| 主文字（近黑） | #1F1F1F（纯 #000000 亦大量出现 = labelColor） |
| 次文字 | #333333 一带（50% 黑在白底的渲染色） |
| 侧栏行距 | 中位 50px = 25pt |
| 内容卡行距（分隔线间距） | 67-68px = 34pt |
| 弹窗按钮（pop-up）bezel 带 | 15px/37px 级 ≈ 19-20pt 高 |
| 选中项/开关填充蓝（高饱和众数） | **#2F6EED**（3173px 众数；#3375F0/#3E8DF7 为抗锯齿邻值） |

预览（Preview）：

| 元素 | 量得值 |
|---|---|
| 统一式标题栏/工具栏底 | #FCFDFF/#FFFFFF（近白） |
| 对话框底 | #FFFFFF；禁用区 #EDEDED；描边 #D7D7D7/#BEBEBE |

## 对比度计算（ui_theme.contrast_ratio 同公式）

| 前景/背景 | 比值 | 结论 |
|---|---|---|
| #1F1F1F / #F7F7F7 | 15.39 | 正文达标 |
| #1F1F1F / #FFFFFF | 16.48 | 卡片正文达标 |
| #333333 / #F7F7F7 | 11.79 | 次文字达标 |
| 白 / #2F6EED（量得蓝） | 4.59 | 蓝底白字达标 |
| #2F6EED / #F7F7F7 | 4.29 | 蓝字在分组底上不达标 → 蓝只用于白底/蓝底场景 |
| 白 / #007AFF（HIG 文档蓝） | 4.02 | **不达标，弃用为主按钮底色** |
| 白 / #0040DD（HIG 可及蓝） | 7.56 | 备选（更暗，观感偏重） |

## 令牌决策（依据以上测量）

- accent = #2F6EED：既是 macOS 26 实际渲染的系统蓝（截图众数），又是唯一与白字过 4.5 的候选。
  HIG 文档值 #007AFF 记录在案但因对比度 4.02 不采用。
- accent_pressed/hover = 在 #2F6EED 基础上加深（#2857C8 / #2860D8 一族），同样过 4.5（测试断言）。
- bg（分组底）= #F7F7F7，surface（卡片/输入底）= #FFFFFF，侧栏/面板 = #EDEDED，
  border = #D8D8D8（#D7D7D7 量化取 8 步进），border_strong = #BEBEBE，
  ink = #1F1F1F，ink_muted = #6E6E73（Apple 灰族，与 #333333 次文字区分层级），
  ink_disabled = #BEBEBE，success = #28CD41（HIG systemGreen 量值，深色文字场景另行断言）。
- 字号（tkinter pt）：display 17（Title2）/ section 13 bold（Headline）/ body 13（Body）/
  caption 11（Subheadline）/ mono 11；Windows 上 YaHei UI 渲染略宽，属平台现实。
- 间距令牌全部 4 的倍数：4/8/12/16/24/32；控件高度 32px 标准、36px 主按钮
  （对应 macOS ~24pt/27pt 控件，落 4 网格）。
