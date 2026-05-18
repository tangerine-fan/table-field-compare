"""
excel-styles — Excel 样式主题 + 写入辅助工具

用法:
    import excel_styles as xls

    # 列出可用主题
    for t in xls.list_themes():
        print(t["name"], t["description"])

    # 获取默认主题
    theme = xls.get_theme()

    # 获取指定主题
    theme = xls.get_theme("obsidian_dark")

    # 写表头
    xls.style_header_row(ws, ["列1", "列2", "列3"])

    # 写数据单元格
    xls.style_data_cell(ws, row=2, col=1, font=theme.BODY_FONT)

    # 设置列宽
    xls.set_column_widths(ws, [20, 30, 15])

    # 冻结表头
    xls.freeze_header(ws)

    # 写汇总行
    xls.style_summary_row(ws, row=6, col_count=6, text="共 5 条")
"""

from .helpers import (
    freeze_header,
    get_theme,
    list_themes,
    set_column_widths,
    style_data_cell,
    style_header_row,
    style_summary_row,
    update_font_name,
)

__all__ = [
    "freeze_header",
    "get_theme",
    "list_themes",
    "set_column_widths",
    "style_data_cell",
    "style_header_row",
    "style_summary_row",
    "update_font_name",
]
