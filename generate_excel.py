# -*- coding: utf-8 -*-
"""电商业绩分析老板看板 - 单Sheet版 v5.2

需求：把 data/ 下面的数据和分析结果都放到同一个 Excel 的同一个 Sheet 里，
不要拆成多个 Sheet。

当前实现：
- 仅生成 1 个工作表：总表
- 同一个 Sheet 内分区展示：
  1. 原始流水
  2. 客户订单汇总
  3. 销售业绩汇总
  4. 销售公司业绩汇总
  5. 公司主体业绩汇总
  6. 查询分析

数据口径：
- 支持 开单日期/业务日期、销售/销售姓名、客户进线日/客户进线日期、订金/定金
- 有订金且有报名费：订金包含在报名费里，不重复统计
- 有订金但没报名费：订金计入实收
- 产品总价优先：手填产品总价 > 全款 > 1800+尾款
- 尾款回收率：若仅有首款、未形成完整尾款口径，则显示 --
"""

import csv
import os
import datetime as dt
from dataclasses import dataclass
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SIGNUP_FEE = 1800
OUTPUT_FILE = "电商业绩分析老板看板.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1890FF")
SUB_FILL = PatternFill("solid", fgColor="E6F0FF")
TITLE_FILL = PatternFill("solid", fgColor="001529")
WHITE_BOLD = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color="FFFFFF", size=13)
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
            return dt.date(2026, int(m), int(d))
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
    lead_date: dt.date | None
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
        return (self.phone or "").strip() or (self.customer or "").strip()


def load_txs_from_sheet1():
    with open(os.path.join("data", "sheet1-业绩流水表.csv"), encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    txs = []
    for r in rows:
        biz_date = parse_date(pick(r, "开单日期", "业务日期"))
        if not biz_date:
            continue
        tx = Tx(
            biz_date=biz_date,
            lead_date=parse_date(pick(r, "客户进线日", "客户进线日期")),
            company=(pick(r, "公司主体") or "").strip(),
            sales_company=(pick(r, "销售公司") or "").strip(),
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


def build_order_summary(txs):
    agg = defaultdict(lambda: {
        "company":"", "sales_company":"", "seller":"", "customer":"", "phone":"", "lead_date":None,
        "deposit":0, "signup":0, "tail":0, "full":0, "refund":0, "total_price":0,
    })
    for t in txs:
        a = agg[t.order_key]
        a["company"] = a["company"] or t.company
        a["sales_company"] = a["sales_company"] or t.sales_company
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
        total_price = a["total_price"]
        if not total_price:
            if a["full"] > 0:
                total_price = a["full"]
            elif first_payment > 0 and a["tail"] > 0:
                total_price = SIGNUP_FEE + a["tail"]

        receivable_tail = total_price - SIGNUP_FEE if total_price else 0
        received_tail = a["tail"]
        if a["full"] > 0:
            received_tail = max(received_tail, a["full"] - SIGNUP_FEE)
        tail_rate = f"{(received_tail / receivable_tail):.1%}" if total_price and receivable_tail > 0 and (a["tail"] > 0 or a["full"] > 0) else "--"
        unpaid_tail = receivable_tail - received_tail if total_price else 0
        gross = first_payment + a["tail"] + a["full"]
        net = gross - a["refund"]

        refund_status = "未退款"
        if a["refund"] > 0 and net <= 0:
            refund_status = "已退款"
        elif a["refund"] > 0 and net > 0:
            refund_status = "部分退款"

        rows.append([
            a["company"], a["sales_company"], a["seller"], a["customer"], a["phone"],
            a["lead_date"].isoformat() if a["lead_date"] else "",
            a["deposit"], a["signup"], first_payment,
            total_price if total_price else "", receivable_tail if total_price else "", received_tail if total_price else "",
            unpaid_tail if total_price else "", tail_rate,
            a["tail"], a["full"], a["refund"], gross, net, refund_status,
        ])

    rows.sort(key=lambda r: (r[0], r[2], r[3]))
    headers = [
        "公司主体", "销售公司", "销售", "客户", "电话", "进线日期",
        "订金", "报名费", "首款计入口径", "产品总价", "应收尾款", "已收尾款", "未收尾款", "尾款回收率",
        "尾款", "全款", "退款", "实收业绩", "净业绩", "退款状态"
    ]
    return headers, rows


def build_seller_summary(order_rows):
    agg = defaultdict(lambda: {"company":set(), "sales_company":set(), "customers":0, "gross":0, "net":0, "refund":0, "rec":0, "got":0, "un":0, "tail_num":0, "tail_den":0})
    for r in order_rows:
        seller = r[2]
        a = agg[seller]
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
        if isinstance(r[10], int) and r[10] > 0:
            a["tail_den"] += r[10]
            a["tail_num"] += (r[11] if isinstance(r[11], int) else 0)

    rows = []
    for seller, a in agg.items():
        refund_rate = (a["refund"] / a["gross"]) if a["gross"] else 0
        tail_rate = (a["tail_num"] / a["tail_den"]) if a["tail_den"] else None
        rows.append([
            seller,
            ",".join(sorted(a["company"])),
            ",".join(sorted(a["sales_company"])) if a["sales_company"] else "",
            a["customers"], a["gross"], a["net"], a["refund"], f"{refund_rate:.1%}",
            a["rec"], a["got"], a["un"], f"{tail_rate:.1%}" if tail_rate is not None else "--",
            0,
        ])
    rows.sort(key=lambda r: r[5], reverse=True)
    for i, r in enumerate(rows, 1):
        r[-1] = i
    headers = ["销售", "公司主体(集合)", "销售公司(集合)", "客户数", "实收业绩", "净业绩", "退款金额", "退款率", "应收尾款", "已收尾款", "未收尾款", "尾款回收率", "排名"]
    return headers, rows


def build_group_summary(order_rows, key_index, key_name):
    agg = defaultdict(lambda: {"orders":0, "gross":0, "net":0, "refund":0, "rec":0, "got":0, "un":0})
    for r in order_rows:
        k = r[key_index] or "(空)"
        a = agg[k]
        a["orders"] += 1
        a["gross"] += r[17] if isinstance(r[17], int) else 0
        a["net"] += r[18] if isinstance(r[18], int) else 0
        a["refund"] += r[16] if isinstance(r[16], int) else 0
        a["rec"] += r[10] if isinstance(r[10], int) else 0
        a["got"] += r[11] if isinstance(r[11], int) else 0
        a["un"] += r[12] if isinstance(r[12], int) else 0
    rows = []
    for k, a in agg.items():
        refund_rate = (a["refund"] / a["gross"]) if a["gross"] else 0
        tail_rate = (a["got"] / a["rec"]) if a["rec"] else None
        rows.append([k, a["orders"], a["gross"], a["net"], a["refund"], f"{refund_rate:.1%}", a["rec"], a["got"], a["un"], f"{tail_rate:.1%}" if tail_rate is not None else "--"])
    rows.sort(key=lambda r: r[3], reverse=True)
    headers = [key_name, "订单数", "实收业绩", "净业绩", "退款金额", "退款率", "应收尾款", "已收尾款", "未收尾款", "尾款回收率"]
    return headers, rows


def write_section_title(ws, row, title, max_col):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    cell = ws.cell(row=row, column=1, value=title)
    cell.fill = TITLE_FILL
    cell.font = TITLE_FONT
    cell.alignment = LEFT
    for c in range(1, max_col + 1):
        ws.cell(row=row, column=c).border = thin_border()


def write_table(ws, start_row, title, headers, rows, widths=None, highlight=None):
    max_col = len(headers)
    write_section_title(ws, start_row, title, max_col)
    header_row = start_row + 1
    for i, h in enumerate(headers, 1):
        ws.cell(row=header_row, column=i, value=h)
    set_header_row(ws, header_row, max_col)

    hl_col = None
    rules = {}
    if highlight and highlight.get("col") in headers:
        hl_col = headers.index(highlight["col"]) + 1
        rules = highlight.get("rules") or {}

    r = header_row + 1
    for row in rows:
        for i, v in enumerate(row, 1):
            ws.cell(row=r, column=i, value=v)
        style_row(ws, r, max_col, SUB_FILL if r % 2 == 0 else None)
        if hl_col:
            v = ws.cell(row=r, column=hl_col).value
            if v in rules:
                ws.cell(row=r, column=hl_col).font = rules[v]
        r += 1

    if widths:
        set_col_width(ws, widths)

    return r + 2


def main():
    txs = load_txs_from_sheet1()
    order_headers, order_rows = build_order_summary(txs)
    seller_headers, seller_rows = build_seller_summary(order_rows)
    sales_company_headers, sales_company_rows = build_group_summary(order_rows, 1, "销售公司")
    company_headers, company_rows = build_group_summary(order_rows, 0, "公司主体")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "总表"
    ws.sheet_view.showGridLines = False

    row = 1

    raw_headers = ["开单日期", "公司主体", "销售公司", "销售", "客户姓名", "年龄", "客户电话", "客户进线日期", "订金", "报名费", "尾款", "全款", "退款", "产品总价", "备注"]
    raw_rows = []
    for t in txs:
        raw_rows.append([
            t.biz_date.isoformat(), t.company, t.sales_company, t.seller, t.customer, t.age, t.phone,
            t.lead_date.isoformat() if t.lead_date else "", t.deposit, t.signup, t.tail, t.full, t.refund,
            t.total_price if t.total_price else "", t.note,
        ])
    row = write_table(ws, row, "1. 原始流水", raw_headers, raw_rows, widths=[12,10,12,10,10,8,14,12,10,10,10,10,10,10,20])

    row = write_table(ws, row, "2. 客户订单汇总", order_headers, order_rows,
                      widths=[10,12,8,10,14,12,10,10,12,10,10,10,10,10,10,10,10,10,10,10],
                      highlight={"col":"退款状态", "rules":{"已退款":RED_FONT, "部分退款":ORANGE_FONT, "未退款":GREEN_FONT}})

    row = write_table(ws, row, "3. 销售业绩汇总", seller_headers, seller_rows,
                      widths=[10,18,18,10,12,12,10,10,10,10,10,10,8])

    row = write_table(ws, row, "4. 销售公司业绩汇总", sales_company_headers, sales_company_rows)
    row = write_table(ws, row, "5. 公司主体业绩汇总", company_headers, company_rows)

    query_headers = ["客户", "销售", "公司主体", "订金", "报名费", "首款计入口径", "尾款", "全款", "退款", "尾款回收率"]
    query_rows = [[r[3], r[2], r[0], r[6], r[7], r[8], r[14], r[15], r[16], r[13]] for r in order_rows]
    row = write_table(ws, row, "6. 查询分析", query_headers, query_rows)

    wb.save(OUTPUT_FILE)
    print(f"✅ 已生成单Sheet文件：{OUTPUT_FILE}")


if __name__ == "__main__":
    main()
