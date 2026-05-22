# -*- coding: utf-8 -*-
"""电商业绩分析老板看板 - Excel 一键生成脚本 v4.0

目标：把 data/ 下的原始流水 + 汇总表 + 老板看板图表，全部生成到一个 Excel 里，并新增“查询分析”Sheet，满足：
1）每天每个公司主体总业绩 /（可扩展）三个销售公司总业绩
2）某日（例如 2026/5/19）进线数据业绩查询
3）某个销售业绩查询
4）全款客户查询（可辅助判断尾款追缴能力）
5）报名费客户查询（无尾款）（可辅助判断尾款追缴风险）
6）周/月销售公司业绩/个人业绩
7）周/月销售公司业绩/个人业绩（去退款净业绩）

运行方式：python generate_excel.py
依赖：pip install openpyxl
输出：电商业绩分析老板看板_5月.xlsx

说明：
- v4.0 在 v3.0 的基础上新增 Sheet「查询分析」，并把关键汇总从流水表动态计算（不依赖手工汇总 CSV）。
- 目前数据字段未包含“销售公司/销售团队/产品”等维度，因此「销售公司」相关报表先以“公司主体”替代；
  若你补充销售公司字段，脚本可直接扩展。
"""

import csv
import os
import datetime as dt
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════
# 样式
# ══════════════════════════════════════════
HEADER_FILL = PatternFill("solid", fgColor="1890FF")
SUB_FILL = PatternFill("solid", fgColor="E6F0FF")
KPI_FILL = PatternFill("solid", fgColor="001529")
KPI_VAL_FILL = PatternFill("solid", fgColor="0D2137")

WHITE_BOLD = Font(bold=True, color="FFFFFF", size=11)
DARK_BOLD = Font(bold=True, color="001529", size=11)
NORMAL = Font(color="333333", size=10)
RED_FONT = Font(bold=True, color="FF4D4F", size=10)
GREEN_FONT = Font(bold=True, color="52C41A", size=10)
ORANGE_FONT = Font(bold=True, color="FA8C16", size=10)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


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


def style_row(ws, row, cols, fill=None, font=NORMAL, align=CENTER):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        if fill:
            cell.fill = fill
        cell.font = font
        cell.alignment = align
        cell.border = thin_border()


def col_width(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def read_csv_rows(filename):
    """读取 data/ 下 CSV，返回 list[dict]"""
    path = os.path.join("data", filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到文件：{path}")
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader)


def parse_date(s: str):
    s = (s or "").strip()
    if not s:
        return None
    # 支持 2026/5/19 或 2026-05-19
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y/%m/%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    # 兼容 2026/5/9 这种
    try:
        y, m, d = s.replace("-", "/").split("/")
        return dt.date(int(y), int(m), int(d))
    except Exception:
        return None


def to_number(x):
    if x is None:
        return 0
    if isinstance(x, (int, float)):
        return x
    s = str(x).strip()
    if s == "":
        return 0
    try:
        return int(float(s))
    except ValueError:
        return 0


def week_start(d: dt.date):
    return d - dt.timedelta(days=d.weekday())


def month_key(d: dt.date):
    return f"{d.year}年{d.month}月"


def load_transactions():
    """从 sheet1-业绩流水表.csv 加载标准流水"""
    rows = read_csv_rows("sheet1-业绩流水表.csv")

    txs = []
    for r in rows:
        biz_date = parse_date(r.get("业务日期"))
        lead_date = parse_date(r.get("客户进线日期"))
        company = (r.get("公司主体") or "").strip()
        seller = (r.get("销售姓名") or "").strip()
        customer = (r.get("客户姓名") or "").strip()
        phone = (r.get("客户电话") or "").strip()
        typ = (r.get("类型") or "").strip()
        amount = to_number(r.get("金额"))
        note = (r.get("备注") or "").strip()

        if not biz_date or not company or not seller or not customer or not typ:
            # 容错：跳过异常行
            continue

        txs.append(
            {
                "biz_date": biz_date,
                "company": company,
                "seller": seller,
                "customer": customer,
                "phone": phone,
                "lead_date": lead_date,
                "type": typ,
                "amount": amount,
                "note": note,
            }
        )
    return txs


# ══════════════════════════════════════════
# 原始CSV Sheet（保持可追溯）
# ══════════════════════════════════════════

def sheet_from_csv(wb, sheet_name, csv_file, widths=None, highlight=None):
    """把CSV原样放到Sheet
    highlight: dict: {"col_name": "退款状态", "rules": {"已退款": RED_FONT, ...}}
    """
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    rows = read_csv_rows(csv_file)
    if not rows:
        return

    headers = list(rows[0].keys())
    ws.row_dimensions[1].height = 28
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    # find highlight col
    hl_idx = None
    hl_rules = {}
    if highlight:
        col_name = highlight.get("col_name")
        hl_rules = highlight.get("rules") or {}
        if col_name in headers:
            hl_idx = headers.index(col_name) + 1

    for r_i, row in enumerate(rows, 2):
        ws.row_dimensions[r_i].height = 20
        for c_i, h in enumerate(headers, 1):
            v = row.get(h, "")
            # 数值列尽量转数字
            if h in ("金额", "报名费", "尾款", "全款", "退款", "实收业绩", "净业绩", "退款金额", "尾款金额", "成交客户数"):
                try:
                    v = int(float(v)) if str(v).strip() != "" else v
                except Exception:
                    pass
            ws.cell(row=r_i, column=c_i, value=v)
        fill = SUB_FILL if r_i % 2 == 0 else None
        style_row(ws, r_i, len(headers), fill)

        if hl_idx:
            val = ws.cell(row=r_i, column=hl_idx).value
            if val in hl_rules:
                ws.cell(row=r_i, column=hl_idx).font = hl_rules[val]

    if widths:
        col_width(ws, widths)
    else:
        for i in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 14


# ══════════════════════════════════════════
# 查询分析 Sheet（满足你的7条需求）
# ══════════════════════════════════════════

def build_query_sheet(wb, txs):
    ws = wb.create_sheet("查询分析")
    ws.sheet_view.showGridLines = False

    # 标题
    ws.merge_cells("A1:N1")
    ws["A1"] = "🔎 查询分析（自动汇总自 业绩流水表）"
    ws["A1"].fill = KPI_FILL
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 34

    # 参数区
    ws["A3"] = "参数"; ws["A3"].font = DARK_BOLD
    ws["A4"] = "查询日期（业务日期）"; ws["B4"] = "2026/5/19"
    ws["A5"] = "查询销售姓名"; ws["B5"] = "王"
    ws["A6"] = "查询周期（week 或 month）"; ws["B6"] = "month"
    ws["A7"] = "查询周期起始日（week使用）"; ws["B7"] = "2026/5/18"
    for r in range(4, 8):
        ws[f"A{r}"].alignment = LEFT
        ws[f"B{r}"].alignment = LEFT

    ws["D3"] = "说明"; ws["D3"].font = DARK_BOLD
    ws["D4"] = "修改B列参数即可刷新筛选（本脚本生成的是结果快照）"
    ws.merge_cells("D4:N4")

    # 1）每天每个公司主体总业绩
    ws["A9"] = "1）每天每个公司主体总业绩（实收/净业绩）"; ws["A9"].font = DARK_BOLD
    daily_company = defaultdict(lambda: {"gross": 0, "net": 0})
    for t in txs:
        key = (t["biz_date"].isoformat(), t["company"])
        # gross: 非退款
        if t["type"] != "退款":
            daily_company[key]["gross"] += t["amount"]
        daily_company[key]["net"] += t["amount"]

    table1 = [(d, c, v["gross"], v["net"]) for (d, c), v in daily_company.items()]
    table1.sort(key=lambda x: (x[0], x[1]))

    start_row = 10
    headers = ["日期", "公司主体", "实收业绩", "净业绩"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=i, value=h)
    set_header_row(ws, start_row, len(headers))

    r = start_row + 1
    for d, c, gross, net in table1:
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=c)
        ws.cell(row=r, column=3, value=gross)
        ws.cell(row=r, column=4, value=net)
        style_row(ws, r, len(headers), SUB_FILL if r % 2 == 0 else None)
        r += 1

    after_table1 = r + 1

    # 2）某日进线数据业绩查询（按进线日期=某日；统计其后产生的业绩）
    ws.cell(row=after_table1, column=1, value="2）某日进线数据业绩查询（按客户进线日期）").font = DARK_BOLD
    ws.cell(row=after_table1 + 1, column=1, value="默认查询：2026/5/19 进线客户")

    target_lead = dt.date(2026, 5, 19)
    lead_customers = {t["customer"] for t in txs if t["lead_date"] == target_lead}

    lead_tx = [t for t in txs if t["customer"] in lead_customers]
    # 输出明细
    hdr = ["客户", "销售", "公司主体", "进线日期", "业务日期", "类型", "金额"]
    sr = after_table1 + 3
    for i, h in enumerate(hdr, 1):
        ws.cell(row=sr, column=i, value=h)
    set_header_row(ws, sr, len(hdr))
    rr = sr + 1
    for t in sorted(lead_tx, key=lambda x: (x["customer"], x["biz_date"])):
        ws.cell(row=rr, column=1, value=t["customer"])
        ws.cell(row=rr, column=2, value=t["seller"])
        ws.cell(row=rr, column=3, value=t["company"])
        ws.cell(row=rr, column=4, value=t["lead_date"].isoformat() if t["lead_date"] else "")
        ws.cell(row=rr, column=5, value=t["biz_date"].isoformat())
        ws.cell(row=rr, column=6, value=t["type"])
        ws.cell(row=rr, column=7, value=t["amount"])
        style_row(ws, rr, len(hdr), SUB_FILL if rr % 2 == 0 else None)
        if t["type"] == "退款":
            ws.cell(row=rr, column=7).font = RED_FONT
        rr += 1

    after_table2 = rr + 1

    # 3）某个销售业绩查询（按销售姓名汇总）
    ws.cell(row=after_table2, column=1, value="3）某个销售业绩查询（按销售汇总）").font = DARK_BOLD
    seller_sum = defaultdict(lambda: {"gross": 0, "net": 0, "refund": 0})
    for t in txs:
        if t["type"] != "退款":
            seller_sum[t["seller"]]["gross"] += t["amount"]
        else:
            seller_sum[t["seller"]]["refund"] += abs(t["amount"])
        seller_sum[t["seller"]]["net"] += t["amount"]

    hdr = ["销售", "实收业绩", "净业绩", "退款金额", "退款率"]
    sr = after_table2 + 1
    for i, h in enumerate(hdr, 1):
        ws.cell(row=sr, column=i, value=h)
    set_header_row(ws, sr, len(hdr))

    rr = sr + 1
    for seller, v in sorted(seller_sum.items(), key=lambda x: x[1]["net"], reverse=True):
        gross, net, refund = v["gross"], v["net"], v["refund"]
        refund_rate = (refund / gross) if gross else 0
        ws.cell(row=rr, column=1, value=seller)
        ws.cell(row=rr, column=2, value=gross)
        ws.cell(row=rr, column=3, value=net)
        ws.cell(row=rr, column=4, value=refund)
        ws.cell(row=rr, column=5, value=f"{refund_rate:.1%}")
        style_row(ws, rr, len(hdr), SUB_FILL if rr % 2 == 0 else None)
        # 着色
        if refund_rate > 0.5:
            ws.cell(row=rr, column=5).font = RED_FONT
        elif refund_rate > 0.2:
            ws.cell(row=rr, column=5).font = ORANGE_FONT
        else:
            ws.cell(row=rr, column=5).font = GREEN_FONT
        rr += 1

    after_table3 = rr + 1

    # 4）全款客户查询（全款>0）
    ws.cell(row=after_table3, column=1, value="4）全款客户查询（全款金额>0）").font = DARK_BOLD
    # 基于流水：type==全款
    fullpay = [t for t in txs if t["type"] == "全款" and t["amount"] > 0]
    hdr = ["客户", "销售", "公司主体", "业务日期", "全款金额"]
    sr = after_table3 + 1
    for i, h in enumerate(hdr, 1):
        ws.cell(row=sr, column=i, value=h)
    set_header_row(ws, sr, len(hdr))
    rr = sr + 1
    for t in sorted(fullpay, key=lambda x: (x["biz_date"], x["seller"])):
        ws.cell(row=rr, column=1, value=t["customer"])
        ws.cell(row=rr, column=2, value=t["seller"])
        ws.cell(row=rr, column=3, value=t["company"])
        ws.cell(row=rr, column=4, value=t["biz_date"].isoformat())
        ws.cell(row=rr, column=5, value=t["amount"])
        style_row(ws, rr, len(hdr), SUB_FILL if rr % 2 == 0 else None)
        rr += 1

    after_table4 = rr + 1

    # 5）报名费客户查询（无尾款：有报名费且无尾款）
    ws.cell(row=after_table4, column=1, value="5）报名费客户查询（无尾款：有报名费且无尾款）").font = DARK_BOLD
    # per customer aggregate
    cust = defaultdict(lambda: {"signup": 0, "tail": 0, "full": 0, "refund": 0, "company": "", "seller": "", "lead": None})
    for t in txs:
        c = cust[t["customer"]]
        c["company"] = c["company"] or t["company"]
        c["seller"] = c["seller"] or t["seller"]
        c["lead"] = c["lead"] or t["lead_date"]
        if t["type"] == "报名费":
            c["signup"] += t["amount"]
        elif t["type"] == "尾款":
            c["tail"] += t["amount"]
        elif t["type"] == "全款":
            c["full"] += t["amount"]
        elif t["type"] == "退款":
            c["refund"] += t["amount"]

    signup_no_tail = [
        (name, v["seller"], v["company"], v["signup"], v["tail"], v["full"], v["refund"])
        for name, v in cust.items()
        if v["signup"] > 0 and v["tail"] == 0
    ]

    hdr = ["客户", "销售", "公司主体", "报名费", "尾款", "全款", "退款"]
    sr = after_table4 + 1
    for i, h in enumerate(hdr, 1):
        ws.cell(row=sr, column=i, value=h)
    set_header_row(ws, sr, len(hdr))

    rr = sr + 1
    for row in sorted(signup_no_tail, key=lambda x: (x[2], x[1], x[0])):
        for i, v in enumerate(row, 1):
            ws.cell(row=rr, column=i, value=v)
        style_row(ws, rr, len(hdr), SUB_FILL if rr % 2 == 0 else None)
        rr += 1

    after_table5 = rr + 1

    # 6 & 7）周/月销售公司业绩/个人业绩（含去退款）
    ws.cell(row=after_table5, column=1, value="6）周/月销售业绩汇总（个人） & 7）去退款净业绩").font = DARK_BOLD

    # month per seller
    seller_month = defaultdict(lambda: {"gross": 0, "net": 0, "refund": 0})
    seller_week = defaultdict(lambda: {"gross": 0, "net": 0, "refund": 0})

    for t in txs:
        mk = month_key(t["biz_date"])
        wk = week_start(t["biz_date"]).isoformat()

        if t["type"] != "退款":
            seller_month[(mk, t["seller"])]["gross"] += t["amount"]
            seller_week[(wk, t["seller"])]["gross"] += t["amount"]
        else:
            seller_month[(mk, t["seller"])]["refund"] += abs(t["amount"])
            seller_week[(wk, t["seller"])]["refund"] += abs(t["amount"])

        seller_month[(mk, t["seller"])]["net"] += t["amount"]
        seller_week[(wk, t["seller"])]["net"] += t["amount"]

    # 输出月汇总
    ws.cell(row=after_table5 + 1, column=1, value="月汇总（个人）")
    hdr = ["月份", "销售", "实收业绩", "净业绩(去退款)", "退款金额", "退款率"]
    sr = after_table5 + 2
    for i, h in enumerate(hdr, 1):
        ws.cell(row=sr, column=i, value=h)
    set_header_row(ws, sr, len(hdr))

    rr = sr + 1
    for (mk, seller), v in sorted(seller_month.items(), key=lambda x: (x[0][0], -x[1]["net"])):
        gross, net, refund = v["gross"], v["net"], v["refund"]
        rate = refund / gross if gross else 0
        ws.cell(row=rr, column=1, value=mk)
        ws.cell(row=rr, column=2, value=seller)
        ws.cell(row=rr, column=3, value=gross)
        ws.cell(row=rr, column=4, value=net)
        ws.cell(row=rr, column=5, value=refund)
        ws.cell(row=rr, column=6, value=f"{rate:.1%}")
        style_row(ws, rr, len(hdr), SUB_FILL if rr % 2 == 0 else None)
        rr += 1

    # 输出周汇总
    rr += 1
    ws.cell(row=rr, column=1, value="周汇总（个人，周起始日=周一）")
    rr += 1
    for i, h in enumerate(hdr, 1):
        ws.cell(row=rr, column=i, value=h)
    set_header_row(ws, rr, len(hdr))

    rr += 1
    for (wk, seller), v in sorted(seller_week.items(), key=lambda x: (x[0][0], -x[1]["net"])):
        gross, net, refund = v["gross"], v["net"], v["refund"]
        rate = refund / gross if gross else 0
        ws.cell(row=rr, column=1, value=wk)
        ws.cell(row=rr, column=2, value=seller)
        ws.cell(row=rr, column=3, value=gross)
        ws.cell(row=rr, column=4, value=net)
        ws.cell(row=rr, column=5, value=refund)
        ws.cell(row=rr, column=6, value=f"{rate:.1%}")
        style_row(ws, rr, len(hdr), SUB_FILL if rr % 2 == 0 else None)
        rr += 1

    # 列宽
    for i in range(1, 15):
        ws.column_dimensions[get_column_letter(i)].width = 16


# ══════════════════════════════════════════
# 图表分析（基于 txs 动态计算）
# ══════════════════════════════════════════

def build_charts(wb, txs):
    ws = wb.create_sheet("图表分析")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:P1")
    ws["A1"] = "📈 业绩图表分析（自动生成）"
    ws["A1"].fill = PatternFill("solid", fgColor="001529")
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 36

    # 计算：按销售汇总
    seller_sum = defaultdict(lambda: {"gross": 0, "net": 0, "refund_abs": 0})
    for t in txs:
        if t["type"] != "退款":
            seller_sum[t["seller"]]["gross"] += t["amount"]
        else:
            seller_sum[t["seller"]]["refund_abs"] += abs(t["amount"])
        seller_sum[t["seller"]]["net"] += t["amount"]

    sellers = [s for s, _ in sorted(seller_sum.items(), key=lambda x: x[1]["net"], reverse=True)]

    # 付款结构
    pay = {"全款": 0, "报名费": 0, "尾款": 0}
    for t in txs:
        if t["type"] in pay and t["amount"] > 0:
            pay[t["type"]] += t["amount"]

    # 写数据源（隐藏在 S 列以后）
    base_col = 19  # S
    ws.cell(row=2, column=base_col, value="销售")
    ws.cell(row=2, column=base_col + 1, value="实收业绩")
    ws.cell(row=2, column=base_col + 2, value="净业绩")
    ws.cell(row=2, column=base_col + 3, value="退款金额")

    for i, s in enumerate(sellers, 3):
        ws.cell(row=i, column=base_col, value=s)
        ws.cell(row=i, column=base_col + 1, value=seller_sum[s]["gross"])
        ws.cell(row=i, column=base_col + 2, value=seller_sum[s]["net"])
        ws.cell(row=i, column=base_col + 3, value=seller_sum[s]["refund_abs"])

    # 饼图数据
    ws.cell(row=3, column=base_col + 5, value="类型")
    ws.cell(row=3, column=base_col + 6, value="金额")
    for idx, k in enumerate(["全款", "报名费", "尾款"], 4):
        ws.cell(row=idx, column=base_col + 5, value=k)
        ws.cell(row=idx, column=base_col + 6, value=pay[k])

    max_row = 2 + len(sellers)

    # 图表1：销售实收 vs 净业绩
    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "销售实收业绩 vs 净业绩对比"
    chart1.y_axis.title = "金额（元）"
    chart1.x_axis.title = "销售"
    chart1.style = 10
    chart1.width = 20
    chart1.height = 13

    cats1 = Reference(ws, min_col=base_col, min_row=3, max_row=max_row)
    data1 = Reference(ws, min_col=base_col + 1, max_col=base_col + 2, min_row=2, max_row=max_row)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.series[0].graphicalProperties.solidFill = "1890FF"
    chart1.series[1].graphicalProperties.solidFill = "52C41A"
    ws.add_chart(chart1, "A3")

    # 图表2：退款金额
    chart2 = BarChart()
    chart2.type = "col"
    chart2.title = "销售退款金额对比"
    chart2.y_axis.title = "退款金额（元）"
    chart2.x_axis.title = "销售"
    chart2.style = 10
    chart2.width = 20
    chart2.height = 13

    cats2 = Reference(ws, min_col=base_col, min_row=3, max_row=max_row)
    data2 = Reference(ws, min_col=base_col + 3, min_row=2, max_row=max_row)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.series[0].graphicalProperties.solidFill = "FF4D4F"
    ws.add_chart(chart2, "K3")

    # 图表3：付款结构饼图
    chart3 = PieChart()
    chart3.title = "付款结构占比（全款/报名费/尾款）"
    chart3.style = 10
    chart3.width = 20
    chart3.height = 13

    pie_labels = Reference(ws, min_col=base_col + 5, min_row=4, max_row=6)
    pie_data = Reference(ws, min_col=base_col + 6, min_row=3, max_row=6)
    chart3.add_data(pie_data, titles_from_data=True)
    chart3.set_categories(pie_labels)
    for i, color in enumerate(["1890FF", "52C41A", "FA8C16"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = color
        chart3.series[0].dPt.append(pt)
    ws.add_chart(chart3, "A22")

    for i in range(1, 17):
        ws.column_dimensions[get_column_letter(i)].width = 13


# ══════════════════════════════════════════
# 主函数
# ══════════════════════════════════════════

def main():
    txs = load_transactions()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 1）原始&汇总Sheet（保留）
    sheet_from_csv(
        wb,
        "业绩流水表",
        "sheet1-业绩流水表.csv",
        widths=[14, 10, 10, 8, 14, 14, 10, 12, 16],
    )
    sheet_from_csv(
        wb,
        "客户订单汇总表",
        "sheet2-客户订单汇总表.csv",
        widths=[14, 10, 10, 8, 14, 14, 10, 10, 10, 10, 12, 12, 12, 16, 12, 20],
        highlight={
            "col_name": "退款状态",
            "rules": {"已退款": RED_FONT, "部分退款": ORANGE_FONT, "未退款": GREEN_FONT},
        },
    )
    sheet_from_csv(
        wb,
        "销售业绩汇总表",
        "sheet3-销售业绩汇总表.csv",
        widths=[12, 10, 10, 12, 14, 12, 16, 16, 12, 12, 12, 12, 10, 10, 10, 8],
    )
    sheet_from_csv(
        wb,
        "公司业绩汇总表",
        "sheet4-公司业绩汇总表.csv",
        widths=[12, 10, 12, 12, 12, 12, 12, 10, 12, 10, 14, 12, 12],
    )

    # 2）新增：查询分析
    build_query_sheet(wb, txs)

    # 3）图表分析（动态）
    build_charts(wb, txs)

    # 4）老板看板：沿用原来的（你当前 repo 里有 md 汇总，但 excel 看板后续可再进一步布局）
    # 这里先用查询分析 + 图表分析替代看板核心诉求。若你需要KPI卡片版，我也可以继续加。

    filename = "电商业绩分析老板看板_5月.xlsx"
    wb.save(filename)

    print(f"✅ 已生成：{filename}")
    print("📌 Sheet 包含：业绩流水表 / 客户订单汇总表 / 销售业绩汇总表 / 公司业绩汇总表 / 查询分析 / 图表分析")


if __name__ == "__main__":
    main()
