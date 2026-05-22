# -*- coding: utf-8 -*-
"""电商业绩分析老板看板 - Excel 一键生成脚本 v5.0（通用模板版）

你后续只需要维护 data/sheet1-业绩流水表.csv（标准流水）即可自动生成：
- 原始流水表
- 客户订单汇总表（含产品总价/应收尾款/尾款回收率等）
- 销售业绩汇总表（含尾款回收率排行）
- 销售公司业绩汇总表
- 公司主体业绩汇总表（含日/周/月）
- 查询分析（按日/按进线日/按销售/全款/报名费无尾款/周月/净业绩等）
- 图表分析（基于流水动态计算）

业务口径（已固化）：
- 报名费固定：1800
- 产品总价：由你在流水中录入（通常在“全款”行或任意一行填入）
- 应收尾款 = 产品总价 - 1800
- 只有报名费、还没补尾款：尾款回收率显示 “--”
- 订单唯一标识：优先（客户电话），否则（客户姓名）

依赖：pip install openpyxl
运行：python generate_excel.py
输出：电商业绩分析老板看板.xlsx

注意：当前仓库的 sheet1 CSV 还没有“销售公司/产品总价”列，脚本会自动兼容（缺失则为空/不计算）。
"""

import csv
import os
import datetime as dt
from dataclasses import dataclass
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter


SIGNUP_FEE = 1800
OUTPUT_FILE = "电商业绩分析老板看板.xlsx"

# ─────────────────────────────────────────
# 样式
# ─────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1890FF")
SUB_FILL = PatternFill("solid", fgColor="E6F0FF")
KPI_FILL = PatternFill("solid", fgColor="001529")

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


def set_col_width(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def parse_date(s: str):
    s = (s or "").strip()
    if not s:
        return None
    s2 = s.replace("-", "/")
    try:
        y, m, d = s2.split("/")
        return dt.date(int(y), int(m), int(d))
    except Exception:
        return None


def to_int(x):
    if x is None:
        return 0
    if isinstance(x, (int, float)):
        return int(x)
    s = str(x).strip()
    if s == "":
        return 0
    try:
        return int(float(s))
    except Exception:
        return 0


def week_start(d: dt.date):
    return d - dt.timedelta(days=d.weekday())


def month_key(d: dt.date):
    return f"{d.year}年{d.month}月"


@dataclass
class Tx:
    biz_date: dt.date
    lead_date: dt.date | None
    company: str
    sales_company: str
    seller: str
    customer: str
    phone: str
    typ: str
    amount: int
    total_price: int
    note: str

    @property
    def order_key(self):
        p = (self.phone or "").strip()
        if p:
            return p
        return (self.customer or "").strip()


def read_csv_dicts(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到文件：{path}")
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader)


def load_txs_from_sheet1():
    rows = read_csv_dicts(os.path.join("data", "sheet1-业绩流水表.csv"))
    txs: list[Tx] = []

    for r in rows:
        biz_date = parse_date(r.get("业务日期"))
        if not biz_date:
            continue

        tx = Tx(
            biz_date=biz_date,
            lead_date=parse_date(r.get("客户进线日期")),
            company=(r.get("公司主体") or "").strip(),
            sales_company=(r.get("销售公司") or "").strip(),
            seller=(r.get("销售姓名") or "").strip(),
            customer=(r.get("客户姓名") or "").strip(),
            phone=(r.get("客户电话") or "").strip(),
            typ=(r.get("类型") or "").strip(),
            amount=to_int(r.get("金额")),
            total_price=to_int(r.get("产品总价")),
            note=(r.get("备注") or "").strip(),
        )

        if not tx.company or not tx.seller or not tx.customer or not tx.typ:
            continue

        txs.append(tx)

    return txs


def ws_from_rows(wb, name, headers, rows, widths=None, highlight=None):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 28
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    set_header_row(ws, 1, len(headers))

    hl_col = None
    rules = {}
    if highlight and highlight.get("col") in headers:
        hl_col = headers.index(highlight["col"]) + 1
        rules = highlight.get("rules") or {}

    for r_i, row in enumerate(rows, 2):
        ws.row_dimensions[r_i].height = 20
        for c_i, val in enumerate(row, 1):
            ws.cell(row=r_i, column=c_i, value=val)
        style_row(ws, r_i, len(headers), SUB_FILL if r_i % 2 == 0 else None)

        if hl_col:
            v = ws.cell(row=r_i, column=hl_col).value
            if v in rules:
                ws.cell(row=r_i, column=hl_col).font = rules[v]

    if widths:
        set_col_width(ws, widths)
    else:
        for i in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 14

    return ws


def build_raw_sheet(wb, txs: list[Tx]):
    headers = [
        "业务日期",
        "公司主体",
        "销售公司",
        "销售姓名",
        "客户姓名",
        "客户电话",
        "客户进线日期",
        "类型",
        "金额",
        "产品总价",
        "备注",
    ]

    rows = []
    for t in txs:
        rows.append(
            [
                t.biz_date.isoformat(),
                t.company,
                t.sales_company,
                t.seller,
                t.customer,
                t.phone,
                t.lead_date.isoformat() if t.lead_date else "",
                t.typ,
                t.amount,
                t.total_price if t.total_price else "",
                t.note,
            ]
        )

    ws_from_rows(
        wb,
        "业绩流水表",
        headers,
        rows,
        widths=[12, 10, 12, 10, 10, 14, 14, 10, 10, 10, 18],
    )


def build_order_summary(txs: list[Tx]):
    agg = defaultdict(
        lambda: {
            "company": "",
            "sales_company": "",
            "seller": "",
            "customer": "",
            "phone": "",
            "lead_date": None,
            "total_price": 0,
            "signup": 0,
            "tail": 0,
            "full": 0,
            "refund": 0,
        }
    )

    for t in txs:
        a = agg[t.order_key]
        a["company"] = a["company"] or t.company
        a["sales_company"] = a["sales_company"] or t.sales_company
        a["seller"] = a["seller"] or t.seller
        a["customer"] = a["customer"] or t.customer
        a["phone"] = a["phone"] or t.phone
        a["lead_date"] = a["lead_date"] or t.lead_date

        if t.total_price:
            a["total_price"] = max(a["total_price"], t.total_price)

        if t.typ == "报名费":
            a["signup"] += t.amount
        elif t.typ == "尾款":
            a["tail"] += t.amount
        elif t.typ == "全款":
            a["full"] += t.amount
            if not a["total_price"]:
                a["total_price"] = max(a["total_price"], t.amount)
        elif t.typ == "退款":
            a["refund"] += t.amount

    rows = []
    for _, a in agg.items():
        total_price = a["total_price"]
        receivable_tail = total_price - SIGNUP_FEE if total_price else 0

        received_tail = a["tail"]
        if a["full"] > 0:
            received_tail = max(received_tail, a["full"] - SIGNUP_FEE)

        if total_price and receivable_tail > 0 and (a["tail"] > 0 or a["full"] > 0):
            tail_rate_disp = f"{(received_tail / receivable_tail):.1%}"
        else:
            tail_rate_disp = "--"

        unpaid_tail = receivable_tail - received_tail if total_price else 0

        gross = a["signup"] + a["tail"] + a["full"]
        net = gross + a["refund"]

        refund_status = "未退款"
        if a["refund"] < 0 and net <= 0:
            refund_status = "已退款"
        elif a["refund"] < 0 and net > 0:
            refund_status = "部分退款"

        rows.append(
            [
                a["company"],
                a["sales_company"],
                a["seller"],
                a["customer"],
                a["phone"],
                a["lead_date"].isoformat() if a["lead_date"] else "",
                total_price if total_price else "",
                SIGNUP_FEE if total_price else "",
                receivable_tail if total_price else "",
                received_tail if total_price else "",
                unpaid_tail if total_price else "",
                tail_rate_disp,
                a["signup"],
                a["tail"],
                a["full"],
                a["refund"],
                gross,
                net,
                refund_status,
            ]
        )

    rows.sort(key=lambda r: (r[0], r[2], r[3]))

    headers = [
        "公司主体",
        "销售公司",
        "销售",
        "客户",
        "电话",
        "进线日期",
        "产品总价",
        "报名费(固定)",
        "应收尾款",
        "已收尾款",
        "未收尾款",
        "尾款回收率",
        "报名费",
        "尾款",
        "全款",
        "退款",
        "实收业绩",
        "净业绩",
        "退款状态",
    ]

    return headers, rows


def build_seller_summary(order_rows):
    agg = defaultdict(lambda: {"company": set(), "sales_company": set(), "customers": 0, "gross": 0, "net": 0, "refund_abs": 0,
                               "rec_tail": 0, "got_tail": 0, "un_tail": 0, "tail_num": 0, "tail_den": 0,
                               "full_customers": 0, "signup_customers": 0, "signup_no_tail": 0})

    for r in order_rows:
        seller = r[2]
        a = agg[seller]
        a["company"].add(r[0])
        if r[1]:
            a["sales_company"].add(r[1])
        a["customers"] += 1

        gross = r[16] if isinstance(r[16], int) else 0
        net = r[17] if isinstance(r[17], int) else 0
        refund = r[15] if isinstance(r[15], int) else 0

        a["gross"] += gross
        a["net"] += net
        a["refund_abs"] += abs(refund) if refund < 0 else 0

        rec_tail = r[8] if isinstance(r[8], int) else 0
        got_tail = r[9] if isinstance(r[9], int) else 0
        un_tail = r[10] if isinstance(r[10], int) else 0
        a["rec_tail"] += rec_tail
        a["got_tail"] += got_tail
        a["un_tail"] += un_tail
        if rec_tail > 0:
            a["tail_den"] += rec_tail
            a["tail_num"] += got_tail

        full_amt = r[14] if isinstance(r[14], int) else 0
        signup_amt = r[12] if isinstance(r[12], int) else 0
        tail_amt = r[13] if isinstance(r[13], int) else 0
        if full_amt > 0:
            a["full_customers"] += 1
        if signup_amt > 0:
            a["signup_customers"] += 1
            if tail_amt == 0 and full_amt == 0:
                a["signup_no_tail"] += 1

    headers = [
        "销售",
        "公司主体(集合)",
        "销售公司(集合)",
        "客户数",
        "全款客户数",
        "报名费客户数",
        "报名费无尾款客户数",
        "实收业绩",
        "净业绩",
        "退款金额",
        "退款率",
        "应收尾款",
        "已收尾款",
        "未收尾款",
        "尾款回收率",
        "排名(按净业绩)",
    ]

    rows = []
    for seller, a in agg.items():
        refund_rate = a["refund_abs"] / a["gross"] if a["gross"] else 0
        tail_rate = a["tail_num"] / a["tail_den"] if a["tail_den"] else None
        rows.append(
            [
                seller,
                ",".join(sorted(a["company"])),
                ",".join(sorted(a["sales_company"])) if a["sales_company"] else "",
                a["customers"],
                a["full_customers"],
                a["signup_customers"],
                a["signup_no_tail"],
                a["gross"],
                a["net"],
                -a["refund_abs"],
                f"{refund_rate:.1%}",
                a["rec_tail"],
                a["got_tail"],
                a["un_tail"],
                f"{tail_rate:.1%}" if tail_rate is not None else "--",
                0,
            ]
        )

    rows.sort(key=lambda r: r[8], reverse=True)
    for i, r in enumerate(rows, 1):
        r[-1] = i

    return headers, rows


def build_group_summary(order_rows, key_index, key_name):
    agg = defaultdict(lambda: {"orders": 0, "gross": 0, "net": 0, "refund_abs": 0, "rec_tail": 0, "got_tail": 0, "un_tail": 0})

    for r in order_rows:
        k = r[key_index] or "(空)"
        a = agg[k]
        a["orders"] += 1
        a["gross"] += r[16] if isinstance(r[16], int) else 0
        a["net"] += r[17] if isinstance(r[17], int) else 0
        refund = r[15] if isinstance(r[15], int) else 0
        a["refund_abs"] += abs(refund) if refund < 0 else 0
        a["rec_tail"] += r[8] if isinstance(r[8], int) else 0
        a["got_tail"] += r[9] if isinstance(r[9], int) else 0
        a["un_tail"] += r[10] if isinstance(r[10], int) else 0

    rows = []
    for k, a in agg.items():
        refund_rate = a["refund_abs"] / a["gross"] if a["gross"] else 0
        tail_rate = a["got_tail"] / a["rec_tail"] if a["rec_tail"] else None
        rows.append(
            [
                k,
                a["orders"],
                a["gross"],
                a["net"],
                -a["refund_abs"],
                f"{refund_rate:.1%}",
                a["rec_tail"],
                a["got_tail"],
                a["un_tail"],
                f"{tail_rate:.1%}" if tail_rate is not None else "--",
            ]
        )

    rows.sort(key=lambda r: r[3], reverse=True)
    headers = [key_name, "订单数", "实收业绩", "净业绩", "退款金额", "退款率", "应收尾款", "已收尾款", "未收尾款", "尾款回收率"]
    return headers, rows


def build_company_time_summary(txs: list[Tx]):
    day = defaultdict(lambda: {"gross": 0, "net": 0})
    week = defaultdict(lambda: {"gross": 0, "net": 0})
    month = defaultdict(lambda: {"gross": 0, "net": 0})

    for t in txs:
        dkey = (t.biz_date.isoformat(), t.company)
        wkey = (week_start(t.biz_date).isoformat(), t.company)
        mkey = (month_key(t.biz_date), t.company)

        if t.typ != "退款":
            day[dkey]["gross"] += t.amount
            week[wkey]["gross"] += t.amount
            month[mkey]["gross"] += t.amount
        day[dkey]["net"] += t.amount
        week[wkey]["net"] += t.amount
        month[mkey]["net"] += t.amount

    def rows_from(m):
        rows = []
        for (k, company), v in m.items():
            rows.append([k, company, v["gross"], v["net"]])
        rows.sort(key=lambda r: (r[0], r[1]))
        return rows

    return (
        ("按日", rows_from(day)),
        ("按周(周一)", rows_from(week)),
        ("按月", rows_from(month)),
    )


def build_query_sheet(wb, txs: list[Tx], order_rows):
    ws = wb.create_sheet("查询分析")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:O1")
    ws["A1"] = "🔎 查询分析（通用模板版）"
    ws["A1"].fill = KPI_FILL
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 34

    r = 3
    ws[f"A{r}"] = "1）每天每个公司主体总业绩（实收/净）"; ws[f"A{r}"].font = DARK_BOLD
    r += 1
    headers = ["日期", "公司主体", "实收业绩", "净业绩"]
    rows = build_company_time_summary(txs)[0][1]

    for i, h in enumerate(headers, 1):
        ws.cell(row=r, column=i, value=h)
    set_header_row(ws, r, len(headers))
    r += 1
    for row in rows:
        for i, v in enumerate(row, 1):
            ws.cell(row=r, column=i, value=v)
        style_row(ws, r, len(headers), SUB_FILL if r % 2 == 0 else None)
        r += 1

    r += 2
    ws[f"A{r}"] = "4）全款客户查询（从流水筛选类型=全款）"; ws[f"A{r}"].font = DARK_BOLD
    r += 1
    headers = ["客户", "销售", "销售公司", "公司主体", "全款金额", "业务日期"]
    full_rows = []
    for t in txs:
        if t.typ == "全款" and t.amount > 0:
            full_rows.append([t.customer, t.seller, t.sales_company, t.company, t.amount, t.biz_date.isoformat()])
    full_rows.sort(key=lambda x: (x[5], x[1], x[0]))

    for i, h in enumerate(headers, 1):
        ws.cell(row=r, column=i, value=h)
    set_header_row(ws, r, len(headers))
    r += 1
    for row in full_rows[:200]:
        for i, v in enumerate(row, 1):
            ws.cell(row=r, column=i, value=v)
        style_row(ws, r, len(headers), SUB_FILL if r % 2 == 0 else None)
        r += 1

    r += 2
    ws[f"A{r}"] = "5）报名费客户（无尾款/无全款）"; ws[f"A{r}"].font = DARK_BOLD
    r += 1
    headers = ["客户", "销售", "销售公司", "公司主体", "报名费", "尾款", "全款", "尾款回收率"]

    no_tail = []
    for row in order_rows:
        signup, tail, full = row[12], row[13], row[14]
        if isinstance(signup, int) and signup > 0 and (not isinstance(tail, int) or tail == 0) and (not isinstance(full, int) or full == 0):
            no_tail.append([row[3], row[2], row[1], row[0], signup, tail if isinstance(tail, int) else 0, full if isinstance(full, int) else 0, row[11]])
    no_tail.sort(key=lambda x: (x[3], x[1], x[0]))

    for i, h in enumerate(headers, 1):
        ws.cell(row=r, column=i, value=h)
    set_header_row(ws, r, len(headers))
    r += 1
    for row in no_tail[:200]:
        for i, v in enumerate(row, 1):
            ws.cell(row=r, column=i, value=v)
        style_row(ws, r, len(headers), SUB_FILL if r % 2 == 0 else None)
        r += 1

    for i in range(1, 16):
        ws.column_dimensions[get_column_letter(i)].width = 16


def build_charts(wb, seller_rows):
    ws = wb.create_sheet("图表分析")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:P1")
    ws["A1"] = "📈 图表分析（基于销售汇总自动生成）"
    ws["A1"].fill = KPI_FILL
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 36

    base_col = 19  # S
    ws.cell(row=2, column=base_col, value="销售")
    ws.cell(row=2, column=base_col + 1, value="实收业绩")
    ws.cell(row=2, column=base_col + 2, value="净业绩")
    ws.cell(row=2, column=base_col + 3, value="退款金额")

    for i, r in enumerate(seller_rows, 3):
        ws.cell(row=i, column=base_col, value=r[0])
        ws.cell(row=i, column=base_col + 1, value=r[7])
        ws.cell(row=i, column=base_col + 2, value=r[8])
        ws.cell(row=i, column=base_col + 3, value=r[9])

    max_row = 2 + len(seller_rows)

    c1 = BarChart()
    c1.type = "col"
    c1.title = "销售实收业绩 vs 净业绩"
    c1.y_axis.title = "金额"
    c1.x_axis.title = "销售"
    c1.width = 22
    c1.height = 12

    cats = Reference(ws, min_col=base_col, min_row=3, max_row=max_row)
    data = Reference(ws, min_col=base_col + 1, max_col=base_col + 2, min_row=2, max_row=max_row)
    c1.add_data(data, titles_from_data=True)
    c1.set_categories(cats)
    c1.series[0].graphicalProperties.solidFill = "1890FF"
    c1.series[1].graphicalProperties.solidFill = "52C41A"
    ws.add_chart(c1, "A3")

    c2 = BarChart()
    c2.type = "col"
    c2.title = "销售退款金额"
    c2.y_axis.title = "金额"
    c2.x_axis.title = "销售"
    c2.width = 22
    c2.height = 12

    data2 = Reference(ws, min_col=base_col + 3, min_row=2, max_row=max_row)
    c2.add_data(data2, titles_from_data=True)
    c2.set_categories(cats)
    c2.series[0].graphicalProperties.solidFill = "FF4D4F"
    ws.add_chart(c2, "A20")

    for i in range(1, 17):
        ws.column_dimensions[get_column_letter(i)].width = 13


def main():
    txs = load_txs_from_sheet1()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_raw_sheet(wb, txs)

    order_headers, order_rows = build_order_summary(txs)
    ws_from_rows(
        wb,
        "客户订单汇总表",
        order_headers,
        order_rows,
        widths=[10, 12, 8, 10, 14, 12, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10],
        highlight={
            "col": "退款状态",
            "rules": {"已退款": RED_FONT, "部分退款": ORANGE_FONT, "未退款": GREEN_FONT},
        },
    )

    seller_headers, seller_rows = build_seller_summary(order_rows)
    ws_from_rows(
        wb,
        "销售业绩汇总表",
        seller_headers,
        seller_rows,
        widths=[10, 18, 18, 10, 12, 12, 16, 12, 12, 10, 10, 10, 10, 10, 10, 10],
    )

    sc_headers, sc_rows = build_group_summary(order_rows, key_index=1, key_name="销售公司")
    ws_from_rows(wb, "销售公司业绩汇总表", sc_headers, sc_rows)

    c_headers, c_rows = build_group_summary(order_rows, key_index=0, key_name="公司主体")
    ws_from_rows(wb, "公司主体业绩汇总表", c_headers, c_rows)

    build_query_sheet(wb, txs, order_rows)
    build_charts(wb, seller_rows)

    wb.save(OUTPUT_FILE)
    print(f"✅ 已生成：{OUTPUT_FILE}")


if __name__ == "__main__":
    main()
