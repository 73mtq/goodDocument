"""Segoe 图标字体系统：字形表、运行时检测、纯文字回退。

字形码点全部来自微软官方字形表（learn.microsoft.com Segoe Fluent Icons），
且只选 Segoe Fluent Icons（Windows 11）与 Segoe MDL2 Assets（Windows 10）
共有的 E7xx-E9xx 区段，一套码点两种字体通吃。
运行时按字体可用性检测：Segoe Fluent Icons → Segoe MDL2 Assets → 无字体（纯文字回退）。
不捆绑字体文件：系统字体授权跟随 Windows，随包分发违约，所以缺失就退到纯文字。
图标统一 20px 网格（行内 16、组头 20、拖放区主图 40），颜色跟随所在文字色。
"""
ICON_FAMILIES = ("Segoe Fluent Icons", "Segoe MDL2 Assets")

# 名称 → 字形。用途注释标在右侧；禁止在此之外出现私用区字符。
ICONS = {
    "page": "\ue7c3",           # Page        设置组「页面」
    "dialpad": "\ue75f",        # Dialpad     设置组「页码」（数字盘）
    "font": "\ue8d2",           # Font        设置组「正文」
    "bold": "\ue8dd",           # Bold        设置组「标题」（加粗 B，四个标题级共用）
    "photo": "\ue91b",          # Photo       设置组「图片」
    "caption": "\ue8ba",        # Caption     设置组「图题与说明」
    "table": "\ue8fd",          # BulletedList 设置组「表格」
    "library": "\ue8f1",        # Library     设置组「参考文献」；主流程「学校模板」
    "document": "\ue8a5",       # Document    主流程「选文档」；拖放区主图
    "save": "\ue74e",           # Save        主流程「保存位置」
    "search": "\ue721",         # Search      过滤框
    "settings": "\ue713",       # Settings    「高级设置」入口
    "undo": "\ue7a7",           # Undo        组头「恢复默认」
    "play": "\ue768",           # Play        「开始整理」
    "view": "\ue890",           # View        「预览改动」
    "edit": "\ue70f",           # Edit        结果区「改动」
    "warning": "\ue7ba",        # Warning     结果区「提醒」
    "list": "\uea37",           # List        结果区「记录」
    "folder": "\ue838",         # FolderOpen  「打开所在文件夹」
    "check": "\ue73e",          # CheckMark   完成
    "chevron_right": "\ue76c",  # ChevronRight 收起面板
    "clear": "\ue711",          # Cancel      清除过滤
}

# 私用区（PUA）判定，供 lint 与测试用：E000-F8FF。
PUA_RANGE = (0xE000, 0xF8FF)


def select_icon_family(available_families):
    """从 tkinter font families() 里挑图标字体；都没有返回 None（纯文字回退）。"""
    available = set(available_families)
    for family in ICON_FAMILIES:
        if family in available:
            return family
    return None


def glyph(name):
    """取字形字符；未知名称返回 None，调用方按无图标排版。"""
    return ICONS.get(name)
