# icon-metrics：图标对齐测量链

图标对齐是「测量驱动的系统属性」：字体表取数 → 解析预测 → GDI 校准 → 屏上验收，
全程可复跑。决策与根因：[DECISION.md](DECISION.md)。运行时唯一入口：仓库根 `icon_slot.py`。

## 文件索引

| 文件 | 作用 | 依赖 |
|---|---|---|
| `measure_glyphs.py` | 字体表（OS/2·hhea·hmtx）+ 逐字形墨迹盒 → `glyph_metrics.json` / `metrics_report.md` | fontTools/Pillow |
| `probe_gdi_placement.py` | Route A 可行性探针：解析预测 vs 真屏 GDI 渲染残差（8 字形 × 4 档） | +tk+numpy |
| `build_icon_data.py` | 生成仓库根 `icon_slot_data.py`（墨迹表 + 校准表） | fontTools |
| `calibrate.py` | 全覆盖校准：22 图标 × 9 档 ppem × 2 字体原始残差 → `calibration_report.json` | +tk |
| `inklab.py` | 共用量具仪器：白框基准画布 + 截图墨迹检测（分量公差判定） | numpy/PIL/tk |
| `specimen.png` | 字样版：全图标全档位，人工关复核材料（tests/test_icon_alignment 产出） | — |
| `acceptance_report.json` | 屏上验收逐格偏差报告（tests/test_icon_alignment 产出） | — |
| `glyph_metrics.json` / `metrics_report.md` | 字体原料测量存档 | — |
| `calibration_report.json` / `calibration_screenshot.png` | 校准存档（残差、5 条离群） | — |
| `probe_results.json` / `probe_screenshot.png` | 探针存档 | — |
| `app_smoke.png` / `exe_smoke.png` | 开发态 / 打包产物冒烟截图 | — |

## 复跑顺序（改了码点表或字体环境变化时）

```
python design/icon-metrics/measure_glyphs.py        # 原料存档
python design/icon-metrics/calibrate.py             # 原始残差 → calibration_report.json
python design/icon-metrics/build_icon_data.py       # 并入 CORRECTIONS → icon_slot_data.py
python -m pytest tests/test_icon_slot.py tests/test_icon_alignment.py   # 数学关 + 屏上验收关
```

calibrate 与 build_icon_data 成对跑：calibrate 测「原始残差」（进程内清空校准防振荡），
build_icon_data 从 calibration_report.json 提取 |分量| ≥ 0.75px 的确定性离群。

## 纪律

- `icon_slot_data.py` 自动生成勿手改（tests/test_icon_slot 与字体文件比对守卫）。
- 码点唯一来源 `ui_icons.ICONS`；PUA 字面量只准在 ui_icons.py（lint + 测试双守卫）。
- 图标尺寸恒为整数像素 = round(设计 token × dpi/96)；token 按 points 传字体的旧做法禁止恢复。
- 调用方只许摆放图标槽本身；槽内微调一律走 `icon_slot.py`（数据驱动，不做手工 padY）。
