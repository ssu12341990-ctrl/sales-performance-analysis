# -*- coding: utf-8 -*-
"""
电商业绩分析老板看板 - Excel 一键生成脚本
运行方式：python generate_excel.py
依赖：pip install openpyxl
生成文件：电商业绩分析老板看板_5月.xlsx
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────
# 样式定义
# ─────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", fgColor="1890FF")   # 蓝色表头
SUB_FILL      = PatternFill("solid", fgColor="E6F0FF")   # 浅蓝隔行
RED_FILL      = PatternFill("solid", fgColor="FF4D4F")   # 红色预警
ORANGE_FILL   = PatternFill("solid", fgColor="FA8C16")   # 橙色警告
GREEN_FILL    = PatternFill("solid", fgColor="52C41A")   # 绿色正常
KPI_FILL      = PatternFill("solid", fgColor="001529")   # 深蓝KPI背景
KPI_VAL_FILL  = PatternFill("solid", fgColor="0D2137")   # KPI数值背景
GRAY_FILL     = PatternFill("solid", fgColor="F5F5F5")   # 灰色背景

WHITE_BOLD    = Font(bold=True, color="FFFFFF", size=11)
BLUE_BOLD     = Font(bold=True, color="1890FF", size=11)
DARK_BOLD     = Font(bold=True, color="001529", size=11)
NORMAL        = Font(color="333333", size=10)
RED_FONT      = Font(bold=True, color="FF4D4F", size=10)
GREEN_FONT    = Font(bold=True, color="52C41A", size=10)
ORANGE_FONT   = Font(bold=True, color="FA8C16", size=10)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def set_header_row(ws, row, cols, fill=HEADER_FILL, font=WHITE_BOLD):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = CENTER
        cell.border = thin_border()

def set_data_row(ws, row, cols, fill=None):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        if fill:
            cell.fill = fill
        cell.font = NORMAL
        cell.alignment = CENTER
        cell.border = thin_border()

def col_width(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ─────────────────────────────────────────
# Sheet1：业绩流水表
# ─────────────────────────────────────────
def build_sheet1(wb):
    ws = wb.create_sheet("业绩流水表")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    headers = ["业务日期","公司主体","销售姓名","客户姓名","客户电话",
               "客户进线日期","类型","金额(元)","备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    data = [
        ("2026/5/19","鸿业","李","a","","2026/5/19","报名费",  1800, "首次报名"),
        ("2026/5/19","鸿业","李","d","","2026/5/18","尾款",   4000, "补尾款"),
        ("2026/5/19","鸿业","李","f","","2026/5/19","全款",   4800, "一次付清"),
        ("2026/5/19","鸿业","李","e","","2026/5/19","退款",  -4800, "退款"),
        ("2026/5/19","鸿业","王","b","","2026/5/18","报名费",  1800, "首次报名"),
        ("2026/5/19","鸿业","王","c","","2026/5/17","报名费",  1800, "首次报名"),
        ("2026/5/19","鸿业","王","g","","2026/5/17","全款",   4800, "一次付清"),
        ("2026/5/19","汉教","刘","h","","2026/5/19","报名费",  1800, "首次报名"),
        ("2026/5/20","鸿业","王","g","","2026/5/17","退款",  -1800, "部分退款"),
        ("2026/5/21","鸿业","王","b","","2026/5/18","退款",  -1800, "退款"),
        ("2026/5/22","鸿业","王","c","","2026/5/17","全款",   4800, "补全款"),
        ("2026/5/22","鸿业","李","a","","2026/5/19","全款",   4800, "补全款"),
        ("2026/5/23","鸿业","李","a","","2026/5/19","退款",  -4800, "退款"),
    ]

    for r, row in enumerate(data, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)
        fill = SUB_FILL if r % 2 == 0 else None
        set_data_row(ws, r, len(headers), fill)
        # 退款红色
        if row[6] == "退款":
            ws.cell(row=r, column=8).font = RED_FONT

    col_width(ws, [14,10,8,8,14,14,10,12,16])
    return ws


# ─────────────────────────────────────────
# Sheet2：客户订单汇总表
# ─────────────────────────────────────────
def build_sheet2(wb):
    ws = wb.create_sheet("客户订单汇总表")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    headers = ["开单日期","公司主体","销售姓名","客户姓名","客户电话",
               "客户进线日期","报名费","尾款","全款","退款",
               "实收业绩","净业绩","尾款状态","付款类型","退款状态","备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    data = [
        ("2026/5/19","鸿业","李","a","","2026/5/19",1800,0,4800,-4800,6600,1800,"无尾款","报名费+全款","已退款","报名后补全款再退款"),
        ("2026/5/19","鸿业","李","d","","2026/5/18",0,4000,0,0,4000,4000,"已补尾款","尾款","未退款","正常补款"),
        ("2026/5/19","鸿业","李","f","","2026/5/19",0,0,4800,0,4800,4800,"无尾款","全款","未退款","正常成交"),
        ("2026/5/19","鸿业","李","e","","2026/5/19",0,0,0,-4800,0,-4800,"无","退款","已退款","纯退款"),
        ("2026/5/19","鸿业","王","b","","2026/5/18",1800,0,0,-1800,1800,0,"未补尾款","报名费","已退款","报名后退款"),
        ("2026/5/19","鸿业","王","c","","2026/5/17",1800,0,4800,0,6600,6600,"无尾款","报名费+全款","未退款","报名后补全款"),
        ("2026/5/19","鸿业","王","g","","2026/5/17",0,0,4800,-1800,4800,3000,"无尾款","全款","部分退款","部分退款"),
        ("2026/5/19","汉教","刘","h","","2026/5/19",1800,0,0,0,1800,1800,"未补尾款","报名费","未退款","待追尾款"),
    ]

    for r, row in enumerate(data, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)
        fill = SUB_FILL if r % 2 == 0 else None
        set_data_row(ws, r, len(headers), fill)
        # 退款状态着色
        refund_status = row[14]
        cell = ws.cell(row=r, column=15)
        if refund_status == "已退款":
            cell.font = RED_FONT
        elif refund_status == "部分退款":
            cell.font = ORANGE_FONT
        else:
            cell.font = GREEN_FONT

    col_width(ws, [14,10,8,8,14,14,10,10,10,10,12,12,12,16,12,20])
    return ws


# ─────────────────────────────────────────
# Sheet3：销售业绩汇总表
# ─────────────────────────────────────────
def build_sheet3(wb):
    ws = wb.create_sheet("销售业绩汇总表")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    headers = ["统计周期","公司主体","销售姓名","成交客户数","报名费客户数",
               "全款客户数","已补尾款客户数","未补尾款客户数","退款客户数",
               "实收业绩","净业绩","退款金额","尾款金额","退款率","客单价","排名"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    data = [
        ("2026年5月","鸿业","王",3,2,1,0,1,2,13200,9600,-3600,0,"27.3%",4400,1),
        ("2026年5月","鸿业","李",4,1,2,1,0,2,15400,5800,-9600,4000,"62.3%",3850,2),
        ("2026年5月","汉教","刘",1,1,0,0,1,0,1800,1800,0,0,"0%",1800,3),
    ]

    for r, row in enumerate(data, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)
        fill = SUB_FILL if r % 2 == 0 else None
        set_data_row(ws, r, len(headers), fill)
        # 退款率着色
        rate_str = row[13]
        rate_val = float(rate_str.replace("%",""))
        cell = ws.cell(row=r, column=14)
        if rate_val > 50:
            cell.font = RED_FONT
        elif rate_val > 20:
            cell.font = ORANGE_FONT
        else:
            cell.font = GREEN_FONT

    col_width(ws, [12,10,10,12,14,12,16,16,12,12,12,12,10,10,10,8])
    return ws


# ─────────────────────────────────────────
# Sheet4：公司业绩汇总表
# ─────────────────────────────────────────
def build_sheet4(wb):
    ws = wb.create_sheet("公司业绩汇总表")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    headers = ["统计周期","公司主体","成交客户数","实收业绩","净业绩",
               "全款金额","报名费金额","尾款金额","退款金额","退款率",
               "尾款未收金额","尾款回收率","备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    data = [
        ("2026年5月","鸿业",7,28600,15400,19200,5400,4000,-13200,"46.2%",0,"100%","含李+王"),
        ("2026年5月","汉教",1,1800,1800,0,1800,0,0,"0%",0,"0%","刘销售"),
        ("2026年5月","合计",8,30400,17200,19200,7200,4000,-13200,"43.4%",0,"100%",""),
    ]

    for r, row in enumerate(data, 2):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)
        fill = SUB_FILL if r % 2 == 0 else None
        if row[1] == "合计":
            fill = PatternFill("solid", fgColor="D0E8FF")
            for c in range(1, len(headers)+1):
                ws.cell(row=r, column=c).font = DARK_BOLD
        set_data_row(ws, r, len(headers), fill)

    col_width(ws, [12,10,12,12,12,12,12,10,12,10,14,12,12])
    return ws


# ─────────────────────────────────────────
# Sheet5：老板看板（KPI + 图表）
# ─────────────────────────────────────────
def build_sheet5(wb):
    ws = wb.create_sheet("老板看板")
    ws.sheet_view.showGridLines = False

    # ── 标题
    ws.merge_cells("A1:L1")
    ws["A1"] = "📊 电商业绩分析 — 老板看板（2026年5月）"
    ws["A1"].fill   = KPI_FILL
    ws["A1"].font   = Font(bold=True, color="FFFFFF", size=16)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 40

    # ── KPI 卡片（第3-8行）
    kpi_data = [
        ("本期总实收业绩", "¥30,400"),
        ("本期净业绩",     "¥17,200"),
        ("总退款金额",     "-¥13,200"),
        ("退款率",         "43.4% ⚠️"),
        ("成交客户数",     "8人"),
        ("尾款回收率",     "100% ✅"),
    ]
    ws.row_dimensions[2].height = 10  # 间距

    kpi_colors = ["1890FF","52C41A","FF4D4F","FA8C16","1890FF","52C41A"]

    col_pairs = [(1,2),(3,4),(5,6),(7,8),(9,10),(11,12)]
    for idx, (label, value) in enumerate(kpi_data):
        c1, c2 = col_pairs[idx]
        row_label = 3
        row_value = 4

        # 标签行
        ws.merge_cells(
            start_row=row_label, start_column=c1,
            end_row=row_label,   end_column=c2
        )
        cell = ws.cell(row=row_label, column=c1, value=label)
        cell.fill = PatternFill("solid", fgColor=kpi_colors[idx])
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = CENTER
        ws.row_dimensions[row_label].height = 22

        # 数值行
        ws.merge_cells(
            start_row=row_value, start_column=c1,
            end_row=row_value,   end_column=c2
        )
        cell2 = ws.cell(row=row_value, column=c1, value=value)
        cell2.fill = KPI_VAL_FILL
        cell2.font = Font(bold=True, color="FFFFFF", size=14)
        cell2.alignment = CENTER
        ws.row_dimensions[row_value].height = 30

    # ── 空行
    ws.row_dimensions[5].height = 10

    # ── 风险预警标题
    ws.merge_cells("A6:L6")
    ws["A6"] = "⚠️ 风险预警区"
    ws["A6"].fill = PatternFill("solid", fgColor="FF4D4F")
    ws["A6"].font = Font(bold=True, color="FFFFFF", size=12)
    ws["A6"].alignment = CENTER
    ws.row_dimensions[6].height = 28

    # ── 风险表
    risk_headers = ["销售","实收业绩","净业绩","退款金额","退款率","预警等级"]
    for i, h in enumerate(risk_headers, 1):
        cell = ws.cell(row=7, column=i, value=h)
        cell.fill = PatternFill("solid", fgColor="FFD6D6")
        cell.font = Font(bold=True, color="CC0000", size=10)
        cell.alignment = CENTER
        cell.border = thin_border()
    ws.row_dimensions[7].height = 22

    risk_data = [
        ("李", 15400, 5800,  -9600, "62.3%", "🔴 高风险"),
        ("王", 13200, 9600,  -3600, "27.3%", "🟠 中风险"),
        ("刘",  1800, 1800,      0,    "0%", "🟢 正常"),
    ]
    risk_fonts = [RED_FONT, ORANGE_FONT, GREEN_FONT]
    for ridx, (r, row) in enumerate(zip(range(8, 11), risk_data)):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = CENTER
            cell.border = thin_border()
            cell.font = risk_fonts[ridx]

    # ── 列宽
    for i in range(1, 13):
        ws.column_dimensions[get_column_letter(i)].width = 14

    return ws


# ─────────────────────────────────────────
# Sheet6：图表页
# ─────────────────────────────────────────
def build_charts(wb):
    ws_chart = wb.create_sheet("图表分析")
    ws_chart.sheet_view.showGridLines = False
    ws_chart.column_dimensions["A"].width = 2  # 左边距

    # 标题
    ws_chart.merge_cells("B1:P1")
    ws_chart["B1"] = "📈 业绩图表分析（2026年5月）"
    ws_chart["B1"].fill = PatternFill("solid", fgColor="001529")
    ws_chart["B1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws_chart["B1"].alignment = CENTER
    ws_chart.row_dimensions[1].height = 36

    # ── 图表数据源（隐藏在本Sheet右侧）
    # 销售净业绩数据 → 放在 R列
    src = [
        ("R1", "销售"),  ("S1", "实收业绩"), ("T1", "净业绩"), ("U1", "退款金额"),
        ("R2", "王"),    ("S2", 13200),       ("T2", 9600),    ("U2", 3600),
        ("R3", "李"),    ("S3", 15400),       ("T3", 5800),    ("U3", 9600),
        ("R4", "刘"),    ("S4", 1800),        ("T4", 1800),    ("U4", 0),
    ]
    for addr, val in src:
        ws_chart[addr] = val

    # 付款结构数据
    pie_src = [
        ("W1", "类型"),   ("X1", "金额"),
        ("W2", "全款"),   ("X2", 19200),
        ("W3", "报名费"), ("X3", 7200),
        ("W4", "尾款"),   ("X4", 4000),
    ]
    for addr, val in pie_src:
        ws_chart[addr] = val

    # ── 图表1：销售实收 vs 净业绩（柱状图）
    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "销售实收业绩 vs 净业绩对比"
    chart1.y_axis.title = "金额（元）"
    chart1.x_axis.title = "销售"
    chart1.style = 10
    chart1.width  = 18
    chart1.height = 12

    cats = Reference(ws_chart, min_col=18, min_row=2, max_row=4)  # R列
    data1 = Reference(ws_chart, min_col=19, max_col=20, min_row=1, max_row=4)  # S-T列
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.series[0].graphicalProperties.solidFill = "1890FF"
    chart1.series[1].graphicalProperties.solidFill = "52C41A"
    ws_chart.add_chart(chart1, "B3")

    # ── 图表2：退款金额柱状图
    chart2 = BarChart()
    chart2.type = "col"
    chart2.title = "销售退款金额对比"
    chart2.y_axis.title = "退款金额（元）"
    chart2.x_axis.title = "销售"
    chart2.style = 10
    chart2.width  = 18
    chart2.height = 12

    cats2 = Reference(ws_chart, min_col=18, min_row=2, max_row=4)
    data2 = Reference(ws_chart, min_col=21, min_row=1, max_row=4)  # U列
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.series[0].graphicalProperties.solidFill = "FF4D4F"
    ws_chart.add_chart(chart2, "K3")

    # ── 图表3：付款结构饼图
    chart3 = PieChart()
    chart3.title = "付款结构占比（全款/报名费/尾款）"
    chart3.style = 10
    chart3.width  = 18
    chart3.height = 12

    pie_labels = Reference(ws_chart, min_col=23, min_row=2, max_row=4)  # W列
    pie_data   = Reference(ws_chart, min_col=24, min_row=1, max_row=4)  # X列
    chart3.add_data(pie_data, titles_from_data=True)
    chart3.set_categories(pie_labels)
    pie_colors = ["1890FF", "52C41A", "FA8C16"]
    for i, color in enumerate(pie_colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = color
        chart3.series[0].dPt.append(pt)
    ws_chart.add_chart(chart3, "B25")

    return ws_chart


# ─────────────────────────────────────────
# 主函数
# ─────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    # 删除默认Sheet
    default = wb.active
    wb.remove(default)

    build_sheet1(wb)   # 业绩流水表
    build_sheet2(wb)   # 客户订单汇总表
    build_sheet3(wb)   # 销售业绩汇总表
    build_sheet4(wb)   # 公司业绩汇总表
    build_sheet5(wb)   # 老板看板（KPI）
    build_charts(wb)   # 图表分析

    filename = "电商业绩分析老板看板_5月.xlsx"
    wb.save(filename)
    print(f"✅ Excel 已生成：{filename}")
    print("📂 包含：业绩流水表 / 客户订单汇总表 / 销售业绩汇总表 / 公司业绩汇总表 / 老板看板 / 图表分析")


if __name__ == "__main__":
    main()
