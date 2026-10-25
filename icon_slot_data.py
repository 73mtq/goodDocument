"""自动生成（design/icon-metrics/build_icon_data.py），勿手改。

INK：两套图标字体逐字形 glyf 墨迹盒（字体单位）；某字体缺失记 None。
CORRECTIONS：GDI grid-fit 确定性离群校准（分量残差 ≥ 0.75px，来源 calibration_report.json；当前 5 条）。
 regeneration 见 design/icon-metrics/DECISION.md。
"""

INK = {
 "Segoe Fluent Icons": {
  "upm": 2048,
  "glyphs": {
   "page": [
    256,
    0,
    1792,
    2048
   ],
   "dialpad": [
    384,
    128,
    1664,
    1920
   ],
   "font": [
    0,
    0,
    2048,
    2048
   ],
   "bold": [
    384,
    140,
    1613,
    1920
   ],
   "photo": [
    128,
    128,
    1920,
    1920
   ],
   "caption": [
    256,
    64,
    1792,
    2048
   ],
   "table": [
    0,
    448,
    2048,
    1728
   ],
   "library": [
    0,
    0,
    2048,
    2048
   ],
   "document": [
    256,
    0,
    1792,
    2048
   ],
   "save": [
    128,
    128,
    1920,
    1920
   ],
   "search": [
    128,
    128,
    1920,
    1920
   ],
   "settings": [
    56,
    0,
    1992,
    2048
   ],
   "undo": [
    256,
    0,
    1920,
    2048
   ],
   "play": [
    256,
    122,
    1913,
    1926
   ],
   "view": [
    30,
    384,
    2018,
    1664
   ],
   "edit": [
    24,
    24,
    2041,
    2048
   ],
   "warning": [
    15,
    126,
    1905,
    2016
   ],
   "list": [
    0,
    256,
    2048,
    1920
   ],
   "folder": [
    0,
    128,
    2019,
    1920
   ],
   "check": [
    256,
    448,
    1792,
    1536
   ],
   "chevron_right": [
    640,
    256,
    1472,
    1792
   ],
   "clear": [
    256,
    384,
    1664,
    1792
   ]
  }
 },
 "Segoe MDL2 Assets": {
  "upm": 2048,
  "glyphs": {
   "page": [
    128,
    0,
    1792,
    2048
   ],
   "dialpad": [
    384,
    128,
    1664,
    1920
   ],
   "font": [
    167,
    256,
    1898,
    1792
   ],
   "bold": [
    512,
    256,
    1701,
    1920
   ],
   "photo": [
    0,
    256,
    2048,
    1792
   ],
   "caption": [
    0,
    0,
    2048,
    1792
   ],
   "table": [
    0,
    384,
    2048,
    1664
   ],
   "library": [
    128,
    249,
    1975,
    1815
   ],
   "document": [
    256,
    0,
    1792,
    2048
   ],
   "save": [
    128,
    128,
    1920,
    1920
   ],
   "search": [
    0,
    0,
    2048,
    2048
   ],
   "settings": [
    6,
    6,
    2042,
    2042
   ],
   "undo": [
    256,
    19,
    1788,
    2048
   ],
   "play": [
    512,
    128,
    1792,
    1920
   ],
   "view": [
    0,
    512,
    2048,
    1792
   ],
   "edit": [
    0,
    0,
    2048,
    2048
   ],
   "warning": [
    0,
    128,
    1920,
    2048
   ],
   "list": [
    0,
    512,
    2048,
    1664
   ],
   "folder": [
    256,
    0,
    1792,
    2048
   ],
   "check": [
    211,
    421,
    1837,
    1581
   ],
   "chevron_right": [
    515,
    3,
    1627,
    2045
   ],
   "clear": [
    243,
    243,
    1805,
    1805
   ]
  }
 }
}

CORRECTIONS = {
 "Segoe Fluent Icons": {
  "warning": {
   "24": [
    0.0,
    1.0
   ]
  }
 },
 "Segoe MDL2 Assets": {
  "warning": {
   "16": [
    0.5,
    1.0
   ],
   "20": [
    0.0,
    1.0
   ],
   "24": [
    0.5,
    1.0
   ],
   "50": [
    0.0,
    1.0
   ]
  }
 }
}


def correction(family, name, size_px):
    """该 (字体, 字形, ppem) 的 grid-fit 校准 [dx, dy]；无则 (0.0, 0.0)。"""
    try:
        dx, dy = CORRECTIONS[family][name][str(size_px)]
        return dx, dy
    except KeyError:
        return 0.0, 0.0
