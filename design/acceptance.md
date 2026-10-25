# v0.7.0 界面重做验收记录（2026-10-25）

四道验收关的结果与证据。取证材料在 design/audit/（改造前）与 design/ref/（Apple 一手截图，均 gitignored）。

## 第 1 关 · 机械 lint（tests/test_ui_lint.py，4 项全过）

- `test_ui_copy_has_no_banned_marks` ok：app.py + ui_theme.py 全部中文文案零 em-dash（U+2014/U+2013）、零箭头（U+2192）、零半角/全角感叹号
- `test_ui_copy_has_no_emoji` ok：零 emoji（含变体选择符、箭头区段）
- `test_spacing_values_come_from_token_table` ok：两文件全部 padx/pady/padding/spacing 字面量 ∈ SPACING 令牌 {4,8,12,16,24,32} ∪ {0}
- `test_spacing_tokens_follow_four_px_grid` ok：令牌全部 4 的倍数

对比度断言（test_text_contrast_meets_wcag_aa，界面实际用到的 14 组全部 ≥ 4.5:1）：
ink/bg 15.39 · ink/surface 16.48 · ink/surface_alt 14.08 · ink_muted/bg 5.58 · ink_muted/surface 5.97 ·
ink_muted/surface_alt 5.10 · ink/accent_soft 14.11 · surface/accent 4.59 · surface/accent_hover 5.59 ·
surface/accent_pressed 6.62 · success/bg 5.26 · success/surface 5.63 · warning/surface 6.51 · danger/surface 5.56

## 第 2 关 · 既有测试全绿

- Python：129 个全过（125 个既有 + 4 个新 lint；涉及主题色的断言更新为 14 组对比度断言，未删除任何测试）
- Node.js：16 个全过
- ruff：All checks passed

## 第 3 关 · 第三方盲测（对比图 design/blind-test-sheet.png / blind-test.pdf，A=macOS 系统设置，B=新界面，C=macOS 预览）

- 原样对比（含窗口边框）：第三方（独立视觉模型，未见实现）**指认 B**，判断线索是右上角 Windows 最小化/最大化/关闭按钮——系统级窗口部件，且任务明确禁止伪造红绿灯。此项在「不贴 macOS 皮」约束下结构性不可通过，如实记录。
- 去边框只比内容区（blind-test-sheet-nochrome.png）：第三方**仍指认 B**，线索为中文黑体的 Windows 渲染（任务指定维持 Microsoft YaHei UI）与圆形虚线拖放区（功能本体）；评审同时确认配色、卡片分组、间距、层级与 macOS 同类一致，未因设计本身被识破。

结论：按「新界面不被认出为通过」的字面标准未通过；被指认的三项依据分别是任务禁止伪造的窗口部件、任务指定维持的字体平台、功能必需的拖放区。颜色、版式、层级本身未被认出。

## 第 4 关 · 5 秒测试（第三方模拟非技术用户，看 after-01-main.png）

- 「这个软件是干什么用的」回答：**「把论文 Word 文档自动排成学校要求的格式」**，5 秒内答对 → 通过
- 「成都理工学士论文怎么排版」回答：点虚线框选论文 → 学校模板选「成都理工大学学士论文」 → 预览改动 → 开始整理，无求助完成 → 通过
- 本轮补充意见：无看不懂的词、无错位、无截断（上一轮的三条文案反馈已修复：文件名占位与拖放区重复、模板说明拗口、「不碰文件」口语化）

## 过程中修复的界面缺陷

- pack 顺序错误导致结果区被挤为 1px（状态栏必须在结果区之前 pack）——窗口几何断言式验证（winfo_height）+ 截图像素验证后修复
- 窗口高度不足（13pt 字体下 780px 放不下主卡+结果区）→ 1120×860，卡片内边距收紧，结果区 118px（约 5 行）
