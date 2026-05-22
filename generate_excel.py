# -*- coding: utf-8 -*-
"""
电商业绩分析老板看板 - Excel 一键生成脚本 v3.0
读取 data/ 目录下的CSV文件，合并所有Sheet+图表到一个Excel
运行方式：python generate_excel.py
依赖：pip install openpyxl
生成文件：电商业绩分析老板看板_5月.xlsx
"""

import csv
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════
# 样式定义
# ══════════════════════════════════════════
HEADER_FILL  = PatternFill("solid", fgColor="1890FF")
SUB_FILL     = PatternFill("solid", fgColor="E6F0FF")
KPI_FILL     = PatternFill("solid", fgColor="001529")
KPI_VAL_FILL = PatternFill("solid", fgColor="0D2137")

WHITE_BOLD  = Font(bold=True, color="FFFFFF", size=11)
DARK_BOLD   = Font(bold=True, color="001529", size=11)
NORMAL      = Font(color="333333", size=10)
RED_FONT    = Font(bold=True, color="FF4D4F", size=10)
GREEN_FONT  = Font(bold=True, color="52C41A", size=10)
ORANGE_FONT = Font(bold=True, color="FA8C16", size=10)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def set_header_row(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = WHITE_BOLD
        cell.alignment = CENTER
        cell.border = thin_border()

def style_row(ws, row, cols, fill=None, font=NORMAL):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        if fill:
            cell.fill = fill
        cell.font = font
        cell.alignment = CENTER
        cell.border = thin_border()

def col_width(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def read_csv(filename):
    """读取 data/ 目录下的CSV文件"""
    path = os.path.join("data", filename)
    if not os.path.exists(path):
        print(f"⚠️  找不到文件：{path}，跳过")
        return [], []
    with open(path, encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return [], []
    return rows[0], rows[1:]  # headers, data


# ══════════════════════════════════════════
# Sheet1：业绩流水表（从CSV读取）
# ══════════════════════════════════════════
def build_sheet1(wb):
    ws = wb.create_sheet("业绩流水表")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 28

    headers, data = read_csv("sheet1-业绩流水表.csv")
    if not headers:
        return
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    for r, row in enumerate(data, 2):
        ws.row_dimensions[r].height = 20
        for c, val in enumerate(row, 1):
            # 数字转换
            try:
                val = int(val)
            except (ValueError, TypeError):
                pass
            ws.cell(row=r, column=c, value=val)
        fill = SUB_FILL if r % 2 == 0 else None
        style_row(ws, r, len(headers), fill)
        # 退款行红色
        if len(row) > 6 and row[6] == "退款":
            ws.cell(row=r, column=8).font = RED_FONT

    col_width(ws, [14,10,10,8,14,14,10,12,16])


# ══════════════════════════════════════════
# Sheet2：客户订单汇总表（从CSV读取）
# ══════════════════════════════════════════
def build_sheet2(wb):
    ws = wb.create_sheet("客户订单汇总表")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 28

    headers, data = read_csv("sheet2-客户订单汇总表.csv")
    if not headers:
        return
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    for r, row in enumerate(data, 2):
        ws.row_dimensions[r].height = 20
        for c, val in enumerate(row, 1):
            try:
                val = int(val)
            except (ValueError, TypeError):
                pass
            ws.cell(row=r, column=c, value=val)
        fill = SUB_FILL if r % 2 == 0 else None
        style_row(ws, r, len(headers), fill)
        # 退款状态着色（第15列）
        if len(row) >= 15:
            status = row[14]
            cell = ws.cell(row=r, column=15)
            if status == "已退款":
                cell.font = RED_FONT
            elif status == "部分退款":
                cell.font = ORANGE_FONT
            else:
                cell.font = GREEN_FONT

    col_width(ws, [14,10,10,8,14,14,10,10,10,10,12,12,12,16,12,20])


# ══════════════════════════════════════════
# Sheet3：销售业绩汇总表（从CSV读取）
# ══════════════════════════════════════════
def build_sheet3(wb):
    ws = wb.create_sheet("销售业绩汇总表")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 28

    headers, data = read_csv("sheet3-销售业绩汇总表.csv")
    if not headers:
        return
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    for r, row in enumerate(data, 2):
        ws.row_dimensions[r].height = 20
        for c, val in enumerate(row, 1):
            try:
                val = int(val)
            except (ValueError, TypeError):
                pass
            ws.cell(row=r, column=c, value=val)
        fill = SUB_FILL if r % 2 == 0 else None
        style_row(ws, r, len(headers), fill)
        # 退款率着色（第14列）
        if len(row) >= 14:
            try:
                rate_val = float(str(row[13]).replace("%",""))
                cell = ws.cell(row=r, column=14)
                if rate_val > 50:
                    cell.font = RED_FONT
                elif rate_val > 20:
                    cell.font = ORANGE_FONT
                else:
                    cell.font = GREEN_FONT
            except ValueError:
                pass

    col_width(ws, [12,10,10,12,14,12,16,16,12,12,12,12,10,10,10,8])


# ══════════════════════════════════════════
# Sheet4：公司业绩汇总表（从CSV读取）
# ══════════════════════════════════════════
def build_sheet4(wb):
    ws = wb.create_sheet("公司业绩汇总表")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 28

    headers, data = read_csv("sheet4-公司业绩汇总表.csv")
    if not headers:
        return
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    for r, row in enumerate(data, 2):
        ws.row_dimensions[r].height = 20
        for c, val in enumerate(row, 1):
            try:
                val = int(val)
            except (ValueError, TypeError):
                pass
            ws.cell(row=r, column=c, value=val)
        # 合计行加粗蓝底
        if len(row) > 1 and row[1] == "合计":
            style_row(ws, r, len(headers),
                      fill=PatternFill("solid", fgColor="D0E8FF"),
                      font=DARK_BOLD)
        else:
            fill = SUB_FILL if r % 2 == 0 else None
            style_row(ws, r, len(headers), fill)

    col_width(ws, [12,10,12,12,12,12,12,10,12,10,14,12,12])


# ══════════════════════════════════════════
# Sheet5：老板看板（KPI + 风险预警）
# ══════════════════════════════════════════
def build_sheet5(wb):
    ws = wb.create_sheet("老板看板")
    ws.sheet_view.showGridLines = False

    # 标题
    ws.merge_cells("A1:L1")
    ws["A1"] = "📊 电商业绩分析 — 老板看板（2026年5月）"
    ws["A1"].fill = KPI_FILL
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=16)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 42

    # KPI卡片
    kpi_data = [
        ("本期总实收业绩", "¥30,400"),
        ("本期净业绩",     "¥17,200"),
        ("总退款金额",     "-¥13,200"),
        ("退款率",         "43.4% ⚠️"),
        ("成交客户数",     "8人"),
        ("尾款回收率",     "100% ✅"),
    ]
    kpi_colors = ["1890FF","52C41A","FF4D4F","FA8C16","1890FF","52C41A"]
    col_pairs  = [(1,2),(3,4),(5,6),(7,8),(9,10),(11,12)]

    ws.row_dimensions[2].height = 8
    for (label, value), color, (c1, c2) in zip(kpi_data, kpi_colors, col_pairs):
        ws.merge_cells(start_row=3, start_column=c1, end_row=3, end_column=c2)
        cell = ws.cell(row=3, column=c1, value=label)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = CENTER
        ws.row_dimensions[3].height = 22

        ws.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c2)
        cell2 = ws.cell(row=4, column=c1, value=value)
        cell2.fill = KPI_VAL_FILL
        cell2.font = Font(bold=True, color="FFFFFF", size=14)
        cell2.alignment = CENTER
        ws.row_dimensions[4].height = 32

    ws.row_dimensions[5].height = 12

    # 销售排名标题
    ws.merge_cells("A6:L6")
    ws["A6"] = "🏆 销售净业绩排名"
    ws["A6"].fill = PatternFill("solid", fgColor="1890FF")
    ws["A6"].font = Font(bold=True, color="FFFFFF", size=12)
    ws["A6"].alignment = CENTER
    ws.row_dimensions[6].height = 26

    rank_headers = ["排名","销售","公司","实收业绩","净业绩","退款金额","退款率","客单价"]
    for i, h in enumerate(rank_headers, 1):
        cell = ws.cell(row=7, column=i, value=h)
        cell.fill = PatternFill("solid", fgColor="E6F0FF")
        cell.font = Font(bold=True, color="1890FF", size=10)
        cell.alignment = CENTER
        cell.border = thin_border()
    ws.row_dimensions[7].height = 22

    rank_data = [
        ("🥇 1","王","鸿业",13200,9600,-3600,"27.3%",4400),
        ("🥈 2","李","鸿业",15400,5800,-9600,"62.3%",3850),
        ("🥉 3","刘","汉教",1800,1800,0,"0%",1800),
    ]
    rank_fonts = [GREEN_FONT, ORANGE_FONT, GREEN_FONT]
    for ridx, (r, row) in enumerate(zip(range(8,11), rank_data)):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = CENTER
            cell.border = thin_border()
            cell.font = rank_fonts[ridx]

    ws.row_dimensions[11].height = 12

    # 风险预警
    ws.merge_cells("A12:L12")
    ws["A12"] = "⚠️ 风险预警区"
    ws["A12"].fill = PatternFill("solid", fgColor="FF4D4F")
    ws["A12"].font = Font(bold=True, color="FFFFFF", size=12)
    ws["A12"].alignment = CENTER
    ws.row_dimensions[12].height = 26

    risk_headers = ["销售","公司","实收业绩","净业绩","退款金额","退款率","预警等级"]
    for i, h in enumerate(risk_headers, 1):
        cell = ws.cell(row=13, column=i, value=h)
        cell.fill = PatternFill("solid", fgColor="FFD6D6")
        cell.font = Font(bold=True, color="CC0000", size=10)
        cell.alignment = CENTER
        cell.border = thin_border()
    ws.row_dimensions[13].height = 22

    risk_data = [
        ("李","鸿业",15400,5800,-9600,"62.3%","🔴 高风险"),
        ("王","鸿业",13200,9600,-3600,"27.3%","🟠 中风险"),
        ("刘","汉教",1800,1800,0,"0%","🟢 正常"),
    ]
    risk_fonts = [RED_FONT, ORANGE_FONT, GREEN_FONT]
    for ridx, (r, row) in enumerate(zip(range(14,17), risk_data)):
        ws.row_dimensions[r].height = 22
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = CENTER
            cell.border = thin_border()
            cell.font = risk_fonts[ridx]

    for i in range(1, 13):
        ws.column_dimensions[get_column_letter(i)].width = 14


# ══════════════════════════════════════════
# Sheet6：图表分析
# ══════════════════════════════════════════
def build_charts(wb):
    ws = wb.create_sheet("图表分析")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:P1")
    ws["A1"] = "📈 业绩图表分析（2026年5月）"
    ws["A1"].fill = PatternFill("solid", fgColor="001529")
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 36

    # 数据源（放右侧隐藏列）
    chart_data = [
        ("S1","销售"), ("T1","实收业绩"), ("U1","净业绩"), ("V1","退款金额"),
        ("S2","王"),   ("T2",13200),     ("U2",9600),    ("V2",3600),
        ("S3","李"),   ("T3",15400),     ("U3",5800),    ("V3",9600),
        ("S4","刘"),   ("T4",1800),      ("U4",1800),    ("V4",0),
        ("S6","公司"), ("T6","实收业绩"), ("U6","净业绩"),
        ("S7","鸿业"), ("T7",28600),     ("U7",15400),
        ("S8","汉教"), ("T8",1800),      ("U8",1800),
        ("S10","类型"),  ("T10","金额"),
        ("S11","全款"),  ("T11",19200),
        ("S12","报名费"),("T12",7200),
        ("S13","尾款"),  ("T13",4000),
    ]
    for addr, val in chart_data:
        ws[addr] = val

    # 图表1：销售实收 vs 净业绩
    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "销售实收业绩 vs 净业绩对比"
    chart1.y_axis.title = "金额（元）"
    chart1.x_axis.title = "销售"
    chart1.style = 10
    chart1.width = 20
    chart1.height = 13
    cats1 = Reference(ws, min_col=19, min_row=2, max_row=4)
    data1 = Reference(ws, min_col=20, max_col=21, min_row=1, max_row=4)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.series[0].graphicalProperties.solidFill = "1890FF"
    chart1.series[1].graphicalProperties.solidFill = "52C41A"
    ws.add_chart(chart1, "A3")

    # 图表2：销售退款金额
    chart2 = BarChart()
    chart2.type = "col"
    chart2.title = "销售退款金额对比"
    chart2.y_axis.title = "退款金额（元）"
    chart2.x_axis.title = "销售"
    chart2.style = 10
    chart2.width = 20
    chart2.height = 13
    cats2 = Reference(ws, min_col=19, min_row=2, max_row=4)
    data2 = Reference(ws, min_col=22, min_row=1, max_row=4)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.series[0].graphicalProperties.solidFill = "FF4D4F"
    ws.add_chart(chart2, "K3")

    # 图表3：公司业绩对比
    chart3 = BarChart()
    chart3.type = "col"
    chart3.title = "公司业绩对比（实收 vs 净业绩）"
    chart3.y_axis.title = "金额（元）"
    chart3.x_axis.title = "公司"
    chart3.style = 10
    chart3.width = 20
    chart3.height = 13
    cats3 = Reference(ws, min_col=19, min_row=7, max_row=8)
    data3 = Reference(ws, min_col=20, max_col=21, min_row=6, max_row=8)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    chart3.series[0].graphicalProperties.solidFill = "1890FF"
    chart3.series[1].graphicalProperties.solidFill = "52C41A"
    ws.add_chart(chart3, "A22")

    # 图表4：付款结构饼图
    chart4 = PieChart()
    chart4.title = "付款结构占比（全款/报名费/尾款）"
    chart4.style = 10
    chart4.width = 20
    chart4.height = 13
    pie_labels = Reference(ws, min_col=19, min_row=11, max_row=13)
    pie_data   = Reference(ws, min_col=20, min_row=10, max_row=13)
    chart4.add_data(pie_data, titles_from_data=True)
    chart4.set_categories(pie_labels)
    for i, color in enumerate(["1890FF","52C41A","FA8C16"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = color
        chart4.series[0].dPt.append(pt)
    ws.add_chart(chart4, "K22")

    for i in range(1, 17):
        ws.column_dimensions[get_column_letter(i)].width = 13


# ══════════════════════════════════════════
# 主函数
# ══════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_sheet1(wb)
    build_sheet2(wb)
    build_sheet3(wb)
    build_sheet4(wb)
    build_sheet5(wb)
    build_charts(wb)

    filename = "电商业绩分析老板看板_5月.xlsx"
    wb.save(filename)
    print(f"✅ Excel 已生成：{filename}")
    print("📂 共6个Sheet：业绩流水表 / 客户订单汇总表 / 销售业绩汇总表 / 公司业绩汇总表 / 老板看板 / 图表分析")

if __name__ == "__main__":
    main()
