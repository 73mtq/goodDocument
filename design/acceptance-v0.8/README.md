# v0.8.0 验收报告（四道关卡）

日期：2026-10-25。对象：`dist\DocFormatter.exe`（v0.8.0，19.9MB，13:24 打包）。

## 关卡 1 · 机械关 —— 通过（结果全贴出）

### 1.1 全量测试

```
Python: 157 passed (pytest tests/ -q, 6.59s)    —— v0.7 为 129，新增 28，零删除
Node:   16 pass / 0 fail (node --test)
ruff:   All checks passed!
```

### 1.2 文案禁令（tests/test_ui_lint.py，扫描 app.py / ui_theme.py 全部字符串字面量）

| 规则 | 结果 |
|---|---|
| 零 emoji（U+1F000-1FAFF、U+2600-27BF、U+2B00-2BFF、U+FE0F） | 通过 |
| 零感叹号（半角/全角） | 通过 |
| 零 em-dash / en-dash / 箭头 | 通过 |
| 零字符图形当图标（U+25A0-25FF 几何、U+2190-21FF 箭头区等，含 ▸/▾/虚线矩形来源区段） | 通过 |
| 私用区字形（图标字体码点）只准出现在 `ui_icons.py` | 通过 |
| `.bind_all(` 全局滚轮劫持 | 已删除，通过 |
| `dash=` 虚线字符画 | 已删除，通过 |

### 1.3 令牌纪律

| 规则 | 结果 |
|---|---|
| `padx/pady/padding/spacing1/spacing3` 全部来自 SPACING 令牌表 | 通过 |
| 控件 `height`/`wraplength` 关键字实参全部来自 MEASUREMENTS/SPACING 令牌表（AST 级检查） | 通过 |
| SPACING 与 MEASUREMENTS 全部落在 4px 网格 | 通过 |

### 1.4 对比度（tests/test_ui_theme.py，WCAG 公式，全部 ≥ 4.5:1）

18 组：ink/bg、ink/surface、ink/surface_alt、ink/surface_hover、ink/surface_pressed、
ink_muted/bg、ink_muted/surface、ink_muted/surface_alt、ink/accent_soft、
surface/accent、surface/accent_hover、surface/accent_pressed、accent/surface、
success/bg、success/surface、warning/surface、danger/surface、danger/bg。全过。

### 1.5 GUI 冒烟（真 Tk 进程，非仅数据层）

构建 → 展开面板 → 过滤「页码」只剩 1 组 → 空结果显示占位 → 切换成都理工预设
→ 过滤「字号」→ 恢复默认 → 切回自定义 → 忙碌态开关 → 拖放区三种绘制 → 全部正常；
控件列对齐以几何数据验证（12 行控件右缘统一 x+w=468，行宽 480）；
本机（Win11）检测到图标字体 Segoe Fluent Icons，图标全部渲染（截图双人评审复核，
无豆腐块、无截断重叠）；exe 启动 6 秒进程存活。

## 关卡 2 · 任务关 —— 材料就绪，待执行

脚本：[task-script.md](task-script.md)。开场白、判定标准、记录表齐备。
要求：≥1 名（最好 3 名）非技术用户，无求助完成「打开高级设置 → 改页码对齐 →
恢复默认 → 收起回主流程」，计时贴出；有求助即回炉。

## 关卡 3 · 图标关 —— 材料就绪，待执行

脚本：[icon-quiz.md](icon-quiz.md)。12 个去标签图标的判定标准与记录表齐备，
含最弱项预判（Dialpad/Bold）与替换候选（Calendar E787 / FontSize E8E9）。
要求：≥5 人，总体正确率 ≥80%。

## 关卡 4 · 可用性关 —— 材料就绪，待执行

脚本：[sus-form.md](sus-form.md)。标准 10 题 SUS 量表与计分表齐备。
要求：3 位非技术用户均分 ≥70；<68 写明短板再改。

## 结论

机械关全部通过并已固化进测试（防回归）；三道人工关依赖真人参试，
脚本与记录表已备好，跑完把结果填进对应记录表并回填本文件即可。

---

## v0.8.1 增补 · 图标对齐关 —— 已机械化（2026-10-25）

用户报告「图标偏移严重」后，图标对齐从目测微调升级为测量驱动系统（取证与定案：
[../icon-metrics/DECISION.md](../icon-metrics/DECISION.md)）。本关并入机械关：

### 新增机械证据（全部固化进测试防回归）

| 项 | 结果 |
|---|---|
| `tests/test_icon_alignment.py` 屏上验收关：真实渲染管线逐格断言分量公差（≤0.5px@100% → ≤1px@150% 线性） | **198/198 达标**（22 图标 × 9 档 ppem；136 格残差恰为 0，最差 ±0.5px=量测粒度） |
| `tests/test_icon_slot.py`：数据形状 / 摆位公式 / 渲染墨迹不越格 / 数据新鲜度（与字体文件比对）/ PUA 纪律 | 12 项全过 |
| 校准器（design/icon-metrics/calibrate.py，396 格原始残差） | 唯一系统性离群 = warning 字形 5 档 ppem 纵向 +1px（grid-fit），已入 CORRECTIONS 校准表 |
| 全量测试 | 169 passed + ruff 全绿 |
| 打包产物复测 | dist\DocFormatter.exe（v0.8.1，19.9MB）启动截图核验：齿轮/拖放区主图/步骤行/白图标/分段/文件夹图标全部对齐（[../icon-metrics/exe_smoke.png](../icon-metrics/exe_smoke.png)） |

### 字样版（图标关人工复核材料）

[../icon-metrics/specimen.png](../icon-metrics/specimen.png)：全部 22 图标 × 9 档尺寸
（token 16/20/40 × 设计缩放 100/125/150%），逐格白框基准渲染，供人工关执行时对照。
偏差报告：[../icon-metrics/acceptance_report.json](../icon-metrics/acceptance_report.json)。

### 遗留观察（不计入本关，待人工关判断）

- 拖放区 document 主图墨迹与「把论文拖到这里」标题间距偏紧（墨迹现已精确居中于
  既定锚点，属既有布局参数，未在本次范围内调整）。
- 三道人工关（任务/图标/可用性）仍待真人执行，脚本不变。

## 结论（v0.8.1 更新）

机械关全部通过（含新增图标对齐关，已固化防回归）；三道人工关依赖真人参试，
脚本与记录表已备好，跑完把结果填进对应记录表并回填本文件即可。
