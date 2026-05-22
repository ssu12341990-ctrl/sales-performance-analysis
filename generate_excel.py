# -*- coding: utf-8 -*-
"""电商业绩分析老板看板 - 老板汇报版 v10.2

兼容性更新：
- 将 `dt.date | None` 改为旧版 Python 兼容写法 `Optional[dt.date]`
- 兼容较旧的 macOS 自带 python3 环境
- 保持现有老板看板、异常校验、进线产值等功能不变

输出文件：电商业绩分析老板看板.xlsx
依赖：pip install openpyxl
"""

import csv
import os
import datetime as dt
from dataclasses import dataclass
from collections import defaultdict, Counter
from typing import Optional

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.series import DataPoint

SIGNUP_FEE = 1800
OUTPUT_FILE = "电商业绩分析老板看板.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1D39C4")
SUB_FILL = PatternFill("solid", fgColor="F7FAFC")
TITLE_FILL = PatternFill("solid", fgColor="0B1F33")
CARD_BLUE = PatternFill("solid", fgColor="1677FF")
CARD_DARK = PatternFill("solid", fgColor="102A43")
CARD_GREEN = PatternFill("solid", fgColor="52C41A")
CARD_ORANGE = PatternFill("solid", fgColor="FA8C16")
CARD_RED = PatternFill("solid", fgColor="FF4D4F")
WHITE_BOLD = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color="FFFFFF", size=18)
SUBTITLE_FONT = Font(color="D9E2EC", size=10)
CARD_LABEL_FONT = Font(bold=True, color="FFFFFF", size=10)
CARD_VALUE_FONT = Font(bold=True, color="FFFFFF", size=16)
NORMAL = Font(color="333333", size=10)
RED_FONT = Font(bold=True, color="FF4D4F", size=10)
GREEN_FONT = Font(bold=True, color="52C41A", size=10)
ORANGE_FONT = Font(bold=True, color="FA8C16", size=10)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def thin_border():
    s = Side(style="thin", color="D9D9D9")
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


def set_col_width(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def parse_date(s):
    s = (s or "").strip().replace("-", "/")
    if not s:
        return None
    parts = s.split("/")
    try:
        if len(parts) == 3:
            y, m, d = parts
            return dt.date(int(y), int(m), int(d))
        if len(parts) == 2:
            m, d = parts
            return dt.date(dt.date.today().year, int(m), int(d))
    except Exception:
        return None
    return None


def to_int(x):
    if x is None:
        return 0
    if isinstance(x, (int, float)):
        return int(x)
    s = str(x).strip().replace(",", "")
    if s == "":
        return 0
    try:
        return int(float(s))
    except Exception:
        return 0


def pick(d, *names):
    for n in names:
        if n in d:
            return d.get(n)
    return None


@dataclass
class Tx:
    biz_date: dt.date
    lead_date: Optional[dt.date]
    company: str
    sales_company: str
    seller: str
    customer: str
    phone: str
    deposit: int
    signup: int
    tail: int
    full: int
    refund: int
    total_price: int
    note: str
    age: str

    @property
    def order_key(self):
        name = (self.customer or "").strip()
        phone = (self.phone or "").strip()
        if name and phone:
            return "%s|%s" % (name, phone)
        return phone or name


def load_txs():
    path = os.path.join("data", "sheet1-业绩流水表.csv")
    with open(path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    txs = []
    for r in rows:
        biz_date = parse_date(pick(r, "开单日期", "业务日期"))
        if not biz_date:
            continue
        company = (pick(r, "公司主体") or "").strip()
        sales_company = (pick(r, "销售公司") or "").strip() or company
        tx = Tx(
            biz_date=biz_date,
            lead_date=parse_date(pick(r, "客户进线日", "客户进线日期")),
            company=company,
            sales_company=sales_company,
            seller=(pick(r, "销售", "销售姓名") or "").strip(),
            customer=(pick(r, "客户姓名") or "").strip(),
            phone=(pick(r, "客户电话") or "").strip(),
            deposit=to_int(pick(r, "订金", "定金")),
            signup=to_int(pick(r, "报名费")),
            tail=to_int(pick(r, "尾款")),
            full=to_int(pick(r, "全款")),
            refund=to_int(pick(r, "退款")),
            total_price=to_int(pick(r, "产品总价")),
            note=(pick(r, "备注") or "").strip(),
            age=str(pick(r, "年龄") or "").strip(),
        )
        if not tx.company or not tx.seller or not tx.customer:
            continue
        txs.append(tx)
    return txs


def summarize_orders(txs):
    agg = defaultdict(lambda: {"company": "", "sales_company": "", "seller": "", "customer": "", "phone": "", "lead_date": None, "deposit": 0, "signup": 0, "tail": 0, "full": 0, "refund": 0, "total_price": 0})
    for t in txs:
        a = agg[t.order_key]
        a["company"] = a["company"] or t.company
        a["sales_company"] = a["sales_company"] or t.sales_company or t.company
        a["seller"] = a["seller"] or t.seller
        a["customer"] = a["customer"] or t.customer
        a["phone"] = a["phone"] or t.phone
        a["lead_date"] = a["lead_date"] or t.lead_date
        a["deposit"] += t.deposit
        a["signup"] += t.signup
        a["tail"] += t.tail
        a["full"] += t.full
        a["refund"] += t.refund
        if t.total_price:
            a["total_price"] = max(a["total_price"], t.total_price)
    rows = []
    for _, a in agg.items():
        first_payment = a["signup"] if a["signup"] > 0 else a["deposit"]
        total_price = a["total_price"] or (a["full"] if a["full"] > 0 else (SIGNUP_FEE + a["tail"] if first_payment > 0 and a["tail"] > 0 else 0))
        receivable_tail = total_price - SIGNUP_FEE if total_price else 0
        received_tail = max(a["tail"], a["full"] - SIGNUP_FEE if a["full"] > 0 else 0)
        unpaid_tail = receivable_tail - received_tail if total_price else 0
        tail_rate = ("%.1f%%" % ((received_tail * 100.0 / receivable_tail))) if total_price and receivable_tail > 0 and (a["tail"] > 0 or a["full"] > 0) else "--"
        gross = first_payment + a["tail"] + a["full"]
        net = gross - a["refund"]
        refund_status = "未退款"
        if a["refund"] > 0 and net <= 0:
            refund_status = "已退款"
        elif a["refund"] > 0:
            refund_status = "部分退款"
        rows.append([a["company"], a["sales_company"], a["seller"], a["customer"], a["phone"], a["lead_date"].isoformat() if a["lead_date"] else "", a["deposit"], a["signup"], first_payment, total_price if total_price else "", receivable_tail if total_price else "", received_tail if total_price else "", unpaid_tail if total_price else "", tail_rate, a["tail"], a["full"], a["refund"], gross, net, refund_status])
    rows.sort(key=lambda x: (x[0], x[2], x[3]))
    headers = ["公司主体", "销售公司", "销售", "客户", "电话", "进线日期", "订金", "报名费", "首款计入口径", "产品总价", "应收尾款", "已收尾款", "未收尾款", "尾款回收率", "尾款", "全款", "退款", "实收业绩", "净业绩", "退款状态"]
    return headers, rows


def summarize_seller(order_rows):
    agg = defaultdict(lambda: {"company": set(), "sales_company": set(), "customers": 0, "gross": 0, "net": 0, "refund": 0, "rec": 0, "got": 0, "un": 0})
    for r in order_rows:
        a = agg[r[2]]
        a["company"].add(r[0])
        if r[1]:
            a["sales_company"].add(r[1])
        a["customers"] += 1
        a["gross"] += r[17] if isinstance(r[17], int) else 0
        a["net"] += r[18] if isinstance(r[18], int) else 0
        a["refund"] += r[16] if isinstance(r[16], int) else 0
        a["rec"] += r[10] if isinstance(r[10], int) else 0
        a["got"] += r[11] if isinstance(r[11], int) else 0
        a["un"] += r[12] if isinstance(r[12], int) else 0
    rows = []
    for seller, a in agg.items():
        refund_rate = (a["refund"] / float(a["gross"])) if a["gross"] else 0
        tail_rate = (a["got"] / float(a["rec"])) if a["rec"] else None
        rows.append([seller, ",".join(sorted(a["company"])), ",".join(sorted(a["sales_company"])), a["customers"], a["gross"], a["net"], a["refund"], "%.1f%%" % (refund_rate * 100), a["rec"], a["got"], a["un"], ("%.1f%%" % (tail_rate * 100)) if tail_rate is not None else "--", 0])
    rows.sort(key=lambda x: x[5], reverse=True)
    for i, r in enumerate(rows, 1):
        r[-1] = i
    headers = ["销售", "公司主体(集合)", "销售公司(集合)", "客户数", "实收业绩", "净业绩", "退款金额", "退款率", "应收尾款", "已收尾款", "未收尾款", "尾款回收率", "排名"]
    return headers, rows


def summarize_group(order_rows, idx, key_name):
    agg = defaultdict(lambda: {"orders": 0, "gross": 0, "net": 0, "refund": 0, "rec": 0, "got": 0, "un": 0})
    for r in order_rows:
        key = r[idx] or "(空)"
        a = agg[key]
        a["orders"] += 1
        a["gross"] += r[17] if isinstance(r[17], int) else 0
        a["net"] += r[18] if isinstance(r[18], int) else 0
        a["refund"] += r[16] if isinstance(r[16], int) else 0
        a["rec"] += r[10] if isinstance(r[10], int) else 0
        a["got"] += r[11] if isinstance(r[11], int) else 0
        a["un"] += r[12] if isinstance(r[12], int) else 0
    rows = []
    for k, a in agg.items():
        refund_rate = (a["refund"] / float(a["gross"])) if a["gross"] else 0
        tail_rate = (a["got"] / float(a["rec"])) if a["rec"] else None
        rows.append([k, a["orders"], a["gross"], a["net"], a["refund"], "%.1f%%" % (refund_rate * 100), a["rec"], a["got"], a["un"], ("%.1f%%" % (tail_rate * 100)) if tail_rate is not None else "--"])
    rows.sort(key=lambda x: x[3], reverse=True)
    headers = [key_name, "订单数", "实收业绩", "净业绩", "退款金额", "退款率", "应收尾款", "已收尾款", "未收尾款", "尾款回收率"]
    return headers, rows


def summarize_lead_value(txs):
    agg = defaultdict(lambda: {"customers": set(), "deposit": 0, "signup": 0, "tail": 0, "full": 0, "refund": 0, "gross": 0, "net": 0})
    seller_agg = defaultdict(lambda: {"customers": set(), "gross": 0, "net": 0, "refund": 0})
    detail_rows = []
    for t in txs:
        if not t.lead_date:
            continue
        lead_key = t.lead_date.isoformat()
        gross = t.deposit + t.signup + t.tail + t.full
        net = gross - t.refund
        a = agg[lead_key]
        a["customers"].add(t.order_key)
        a["deposit"] += t.deposit
        a["signup"] += t.signup
        a["tail"] += t.tail
        a["full"] += t.full
        a["refund"] += t.refund
        a["gross"] += gross
        a["net"] += net
        s = seller_agg[(lead_key, t.seller)]
        s["customers"].add(t.order_key)
        s["gross"] += gross
        s["net"] += net
        s["refund"] += t.refund
        detail_rows.append([lead_key, t.biz_date.isoformat(), t.company, t.sales_company, t.seller, t.customer, t.phone, t.deposit, t.signup, t.tail, t.full, t.refund, gross, net, t.note])
    summary_rows = [[lead_date, len(a["customers"]), a["deposit"], a["signup"], a["tail"], a["full"], a["refund"], a["gross"], a["net"]] for lead_date, a in sorted(agg.items())]
    seller_rows = []
    for (lead_date, seller), a in sorted(seller_agg.items(), key=lambda x: (x[0][0], -x[1]["net"])):
        refund_rate = (a["refund"] / float(a["gross"])) if a["gross"] else 0
        seller_rows.append([lead_date, seller, len(a["customers"]), a["gross"], a["net"], a["refund"], "%.1f%%" % (refund_rate * 100)])
    return summary_rows, seller_rows, detail_rows


def build_abnormal_checks(txs):
    rows = []
    for t in txs:
        if not t.phone:
            rows.append(["缺手机号", t.company, t.seller, t.customer, t.biz_date.isoformat(), "客户电话为空"])
    pay_counter = Counter((t.customer, t.phone, t.biz_date.isoformat(), t.deposit, t.signup, t.tail, t.full, t.refund) for t in txs)
    for key, cnt in pay_counter.items():
        customer, phone, biz_date, deposit, signup, tail, full, refund = key
        if cnt > 1 and (deposit or signup or tail or full or refund):
            rows.append(["疑似重复收款", "", "", customer, biz_date, "同客户同电话同日同金额记录数=%s" % cnt])
    headers = ["异常类型", "公司主体", "销售", "客户", "日期/电话", "说明"]
    return headers, rows


def write_sheet_table(ws, title, headers, rows, widths=None, highlight=None):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws.cell(row=1, column=1, value=title)
    c.fill = TITLE_FILL
    c.font = Font(bold=True, color="FFFFFF", size=14)
    c.alignment = LEFT
    for i, h in enumerate(headers, 1):
        ws.cell(row=2, column=i, value=h)
    set_header_row(ws, 2, len(headers))
    hl_col = None
    rules = {}
    if highlight and highlight.get("col") in headers:
        hl_col = headers.index(highlight["col"]) + 1
        rules = highlight.get("rules") or {}
    row_idx = 3
    for row in rows:
        for i, v in enumerate(row, 1):
            ws.cell(row=row_idx, column=i, value=v)
        style_row(ws, row_idx, len(headers), SUB_FILL if row_idx % 2 == 1 else None)
        if hl_col:
            v = ws.cell(row=row_idx, column=hl_col).value
            if v in rules:
                ws.cell(row=row_idx, column=hl_col).font = rules[v]
        row_idx += 1
    set_col_width(ws, widths or [14] * len(headers))
    ws.freeze_panes = "A3"


def build_dashboard(wb, txs, order_rows, seller_rows, company_rows, lead_summary_rows, abnormal_rows):
    ws = wb.create_sheet("老板看板", 0)
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:L1")
    ws["A1"] = "📊 电商业绩分析老板看板"
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = CENTER
    ws.merge_cells("A2:L2")
    ws["A2"] = "汇报日期：%s  |  数据来源：sheet1-业绩流水表.csv  |  订单唯一口径=客户姓名+客户电话  |  销售公司=公司主体" % dt.date.today().isoformat()
    ws["A2"].fill = TITLE_FILL
    ws["A2"].font = SUBTITLE_FONT
    ws["A2"].alignment = CENTER

    total_gross = sum(r[17] for r in order_rows if isinstance(r[17], int))
    total_net = sum(r[18] for r in order_rows if isinstance(r[18], int))
    total_refund = sum(r[16] for r in order_rows if isinstance(r[16], int))
    total_orders = len(order_rows)
    abnormal_count = len(abnormal_rows)
    refund_rate = (total_refund / float(total_gross)) if total_gross else None

    kpis = [
        ("总实收业绩", total_gross, CARD_BLUE),
        ("总净业绩", total_net, CARD_GREEN),
        ("总退款金额", total_refund, CARD_RED),
        ("退款率", ("%.1f%%" % (refund_rate * 100)) if refund_rate is not None else "--", CARD_ORANGE),
        ("订单数", total_orders, CARD_DARK),
        ("异常条数", abnormal_count, CARD_RED if abnormal_count else CARD_GREEN),
    ]
    for (label, value, fill), col in zip(kpis, [1, 3, 5, 7, 9, 11]):
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        ws.merge_cells(start_row=5, start_column=col, end_row=6, end_column=col + 1)
        ws.cell(row=4, column=col, value=label).fill = fill
        ws.cell(row=4, column=col).font = CARD_LABEL_FONT
        ws.cell(row=4, column=col).alignment = CENTER
        ws.cell(row=5, column=col, value=value).fill = CARD_DARK
        ws.cell(row=5, column=col).font = CARD_VALUE_FONT
        ws.cell(row=5, column=col).alignment = CENTER

    seller_base = 20
    ws.cell(row=2, column=seller_base, value="销售")
    ws.cell(row=2, column=seller_base + 1, value="净业绩")
    for i, row in enumerate(seller_rows[:10], 3):
        ws.cell(row=i, column=seller_base, value=row[0])
        ws.cell(row=i, column=seller_base + 1, value=row[5])

    company_base = 30
    ws.cell(row=2, column=company_base, value="公司主体")
    ws.cell(row=2, column=company_base + 1, value="净业绩")
    for i, row in enumerate(company_rows[:10], 3):
        ws.cell(row=i, column=company_base, value=row[0])
        ws.cell(row=i, column=company_base + 1, value=row[3])

    lead_base = 40
    ws.cell(row=2, column=lead_base, value="进线日期")
    ws.cell(row=2, column=lead_base + 1, value="净产值")
    top_leads = sorted(lead_summary_rows, key=lambda x: x[8], reverse=True)[:10]
    for i, row in enumerate(top_leads, 3):
        ws.cell(row=i, column=lead_base, value=row[0])
        ws.cell(row=i, column=lead_base + 1, value=row[8])

    day_map = defaultdict(int)
    pay_map = {"订金": 0, "报名费": 0, "尾款": 0, "全款": 0}
    for t in txs:
        pay_map["订金"] += t.deposit
        pay_map["报名费"] += t.signup
        pay_map["尾款"] += t.tail
        pay_map["全款"] += t.full
        day_map[t.biz_date.isoformat()] += t.deposit + t.signup + t.tail + t.full - t.refund
    pay_base = 50
    ws.cell(row=2, column=pay_base, value="类型")
    ws.cell(row=2, column=pay_base + 1, value="金额")
    for i, key in enumerate(["订金", "报名费", "尾款", "全款"], 3):
        ws.cell(row=i, column=pay_base, value=key)
        ws.cell(row=i, column=pay_base + 1, value=pay_map[key])

    trend_base = 55
    ws.cell(row=2, column=trend_base, value="日期")
    ws.cell(row=2, column=trend_base + 1, value="净流入")
    trend_items = sorted(day_map.items())
    for i, item in enumerate(trend_items, 3):
        d, v = item
        ws.cell(row=i, column=trend_base, value=d)
        ws.cell(row=i, column=trend_base + 1, value=v)

    chart1 = BarChart()
    chart1.type = "bar"
    chart1.title = "销售净业绩排行"
    chart1.width = 17
    chart1.height = 8
    chart1.add_data(Reference(ws, min_col=seller_base + 1, min_row=2, max_row=2 + min(len(seller_rows), 10)), titles_from_data=True)
    chart1.set_categories(Reference(ws, min_col=seller_base, min_row=3, max_row=2 + min(len(seller_rows), 10)))
    chart1.series[0].graphicalProperties.solidFill = "1677FF"
    ws.add_chart(chart1, "A9")

    chart2 = BarChart()
    chart2.type = "col"
    chart2.title = "公司主体净业绩对比"
    chart2.width = 17
    chart2.height = 8
    chart2.add_data(Reference(ws, min_col=company_base + 1, min_row=2, max_row=2 + min(len(company_rows), 10)), titles_from_data=True)
    chart2.set_categories(Reference(ws, min_col=company_base, min_row=3, max_row=2 + min(len(company_rows), 10)))
    chart2.series[0].graphicalProperties.solidFill = "52C41A"
    ws.add_chart(chart2, "G9")

    if trend_items:
        chart3 = LineChart()
        chart3.title = "每日净流入趋势"
        chart3.width = 17
        chart3.height = 8
        chart3.add_data(Reference(ws, min_col=trend_base + 1, min_row=2, max_row=2 + len(trend_items)), titles_from_data=True)
        chart3.set_categories(Reference(ws, min_col=trend_base, min_row=3, max_row=2 + len(trend_items)))
        chart3.series[0].graphicalProperties.line.solidFill = "FA8C16"
        ws.add_chart(chart3, "A26")

    chart4 = PieChart()
    chart4.title = "收款结构占比"
    chart4.width = 17
    chart4.height = 8
    chart4.add_data(Reference(ws, min_col=pay_base + 1, min_row=2, max_row=6), titles_from_data=True)
    chart4.set_categories(Reference(ws, min_col=pay_base, min_row=3, max_row=6))
    for i, color in enumerate(["13C2C2", "1677FF", "FA8C16", "52C41A"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = color
        chart4.series[0].dPt.append(pt)
    ws.add_chart(chart4, "G26")

    ws.merge_cells("A43:F43")
    ws["A43"] = "🏆 销售排行榜"
    ws["A43"].fill = TITLE_FILL
    ws["A43"].font = WHITE_BOLD
    ws["A43"].alignment = CENTER
    seller_headers = ["排名", "销售", "净业绩", "退款率", "未收尾款", "销售公司"]
    for i, h in enumerate(seller_headers, 1):
        ws.cell(row=44, column=i, value=h)
    set_header_row(ws, 44, len(seller_headers))
    row_no = 45
    for row in seller_rows[:10]:
        vals = [row[12], row[0], row[5], row[7], row[10], row[2]]
        for i, v in enumerate(vals, 1):
            ws.cell(row=row_no, column=i, value=v)
        style_row(ws, row_no, len(seller_headers), SUB_FILL if row_no % 2 == 1 else None)
        row_no += 1

    ws.merge_cells("H43:L43")
    ws["H43"] = "🔎 异常概览"
    ws["H43"].fill = TITLE_FILL
    ws["H43"].font = WHITE_BOLD
    ws["H43"].alignment = CENTER
    abnormal_headers = ["异常类型", "客户", "说明", "优先级", "建议"]
    for i, h in enumerate(abnormal_headers, 8):
        ws.cell(row=44, column=i, value=h)
    set_header_row(ws, 44, len(abnormal_headers))
    row_no = 45
    for row in abnormal_rows[:10]:
        abnormal_type = row[0]
        priority = "中"
        suggestion = "检查录入"
        font = ORANGE_FONT
        if abnormal_type == "疑似重复收款":
            priority = "高"
            suggestion = "核对是否误录重复金额"
            font = RED_FONT
        elif abnormal_type == "缺手机号":
            priority = "中"
            suggestion = "建议补充手机号以便唯一校验"
        vals = [row[0], row[3], row[5], priority, suggestion]
        for i, v in enumerate(vals, 8):
            ws.cell(row=row_no, column=i, value=v)
        style_row(ws, row_no, len(abnormal_headers), SUB_FILL if row_no % 2 == 1 else None)
        ws.cell(row=row_no, column=11).font = font
        row_no += 1

    set_col_width(ws, [16, 14, 14, 14, 14, 16, 4, 16, 14, 20, 10, 22])
    ws.freeze_panes = "A9"


def main():
    txs = load_txs()
    order_headers, order_rows = summarize_orders(txs)
    seller_headers, seller_rows = summarize_seller(order_rows)
    sales_company_headers, sales_company_rows = summarize_group(order_rows, 1, "销售公司")
    company_headers, company_rows = summarize_group(order_rows, 0, "公司主体")
    lead_summary_rows, lead_seller_rows, lead_detail_rows = summarize_lead_value(txs)
    abnormal_headers, abnormal_rows = build_abnormal_checks(txs)

    raw_headers = ["开单日期", "公司主体", "销售公司", "销售", "客户姓名", "年龄", "客户电话", "客户进线日", "订金", "报名费", "尾款", "全款", "退款", "产品总价", "备注"]
    raw_rows = [[t.biz_date.isoformat(), t.company, t.sales_company, t.seller, t.customer, t.age, t.phone, t.lead_date.isoformat() if t.lead_date else "", t.deposit, t.signup, t.tail, t.full, t.refund, t.total_price if t.total_price else "", t.note] for t in txs]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_dashboard(wb, txs, order_rows, seller_rows, company_rows, lead_summary_rows, abnormal_rows)

    ws1 = wb.create_sheet("原始流水")
    write_sheet_table(ws1, "原始流水", raw_headers, raw_rows, widths=[12, 10, 12, 10, 10, 8, 14, 12, 10, 10, 10, 10, 10, 10, 20])

    ws2 = wb.create_sheet("客户订单汇总")
    write_sheet_table(ws2, "客户订单汇总", order_headers, order_rows, widths=[10, 12, 10, 10, 14, 12, 10, 10, 12, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10], highlight={"col": "退款状态", "rules": {"已退款": RED_FONT, "部分退款": ORANGE_FONT, "未退款": GREEN_FONT}})

    ws3 = wb.create_sheet("销售业绩汇总")
    write_sheet_table(ws3, "销售业绩汇总", seller_headers, seller_rows, widths=[10, 18, 18, 10, 12, 12, 10, 10, 10, 10, 10, 10, 8])

    ws4 = wb.create_sheet("销售公司业绩汇总")
    write_sheet_table(ws4, "销售公司业绩汇总", sales_company_headers, sales_company_rows)

    ws5 = wb.create_sheet("公司主体业绩汇总")
    write_sheet_table(ws5, "公司主体业绩汇总", company_headers, company_rows)

    ws6 = wb.create_sheet("进线产值汇总")
    write_sheet_table(ws6, "进线产值汇总", ["进线日期", "客户数", "订金", "报名费", "尾款", "全款", "退款", "实收产值", "净产值"], lead_summary_rows, widths=[12, 10, 10, 10, 10, 10, 10, 12, 12])

    ws7 = wb.create_sheet("进线产值-销售")
    write_sheet_table(ws7, "进线产值-销售分析", ["进线日期", "销售", "客户数", "实收产值", "净产值", "退款金额", "退款率"], lead_seller_rows, widths=[12, 10, 10, 12, 12, 12, 10])

    ws8 = wb.create_sheet("进线产值明细")
    write_sheet_table(ws8, "进线产值明细", ["进线日期", "业务日期", "公司主体", "销售公司", "销售", "客户", "电话", "订金", "报名费", "尾款", "全款", "退款", "实收产值", "净产值", "备注"], lead_detail_rows, widths=[12, 12, 10, 12, 10, 10, 14, 10, 10, 10, 10, 10, 12, 12, 20])

    ws9 = wb.create_sheet("异常校验")
    write_sheet_table(ws9, "异常校验", abnormal_headers, abnormal_rows, widths=[14, 10, 10, 12, 14, 28])

    wb.save(OUTPUT_FILE)
    print("✅ 已生成老板汇报版仪表盘：%s" % OUTPUT_FILE)


if __name__ == "__main__":
    main()
