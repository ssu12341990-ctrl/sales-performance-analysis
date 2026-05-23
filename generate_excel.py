# -*- coding: utf-8 -*-
"""电商业绩分析老板看板 - 汇报版 v10.7

更新要点（按用户需求）：
- 进线相关 Sheet 排到最前面
- 第一张看板标题去掉“老板”字样
- 新增：一个总览 Sheet（图表形式）
  - 自然周 + 销售公司 汇总
  - 月 + 销售公司 汇总
  - 自然周/月 + 销售 汇总
- 新增：报名费无尾款客户清单（报名费>0 且 尾款=0 且 全款=0）
- 继续保留：问题明细、异常校验、原始流水、客户订单汇总、公司/销售汇总等

输出：电商业绩分析老板看板.xlsx
依赖：pip install openpyxl
"""

import csv
import os
import datetime as dt
from dataclasses import dataclass
from collections import defaultdict, Counter
from typing import Optional, Tuple, List, Dict, Any

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint

SIGNUP_FEE = 1800
OUTPUT_FILE = "电商业绩分析老板看板.xlsx"

REAL_DATA_FILE = os.path.join("data", "销售_业绩new.csv")
DEFAULT_DATA_FILE = os.path.join("data", "sheet1-业绩流水表.csv")

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

# 公司主体/企业颜色映射（可扩展）
COMPANY_COLORS = {
    "汉教": "1677FF",  # 蓝
    "鸿业": "52C41A",  # 绿
    "心途": "FA8C16",  # 橙
    "(空)": "BFBFBF",  # 灰
}
DEFAULT_COMPANY_COLOR = "722ED1"  # 紫


def _company_color(name: str) -> str:
    name = (name or "").strip() or "(空)"
    return COMPANY_COLORS.get(name, DEFAULT_COMPANY_COLOR)


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


def parse_date(s: Any) -> Optional[dt.date]:
    s = ("" if s is None else str(s)).strip()
    if not s:
        return None
    s = s.split(" ")[0]
    s = s.replace("-", "/")
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


def to_int(x: Any) -> int:
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


def pick(d: Dict[str, Any], *names: str) -> Any:
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


# ------------------ data loading ------------------

def _detect_data_path() -> str:
    if os.path.exists(REAL_DATA_FILE):
        return REAL_DATA_FILE
    return DEFAULT_DATA_FILE


def _read_csv_dict_rows(path: str) -> Tuple[List[str], List[Dict[str, Any]]]:
    last_err = None
    for enc in ("utf-8", "utf-8-sig", "gbk"):
        try:
            with open(path, encoding=enc) as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                return (reader.fieldnames or []), rows
        except Exception as e:
            last_err = e
    raise last_err


def load_txs_with_issues() -> Tuple[List[Tx], List[str], List[Dict[str, Any]]]:
    path = _detect_data_path()
    fieldnames, rows = _read_csv_dict_rows(path)

    txs: List[Tx] = []
    issues: List[Dict[str, Any]] = []

    for i, r in enumerate(rows, start=2):
        biz_raw = pick(r, "开单日期", "业务日期")
        biz_date = parse_date(biz_raw)
        if not biz_date:
            issues.append({"row_index": i, "reason": "开单日期解析失败", "raw": r})
            continue

        company = (pick(r, "公司主体") or "").strip()
        seller = (pick(r, "销售", "销售姓名") or "").strip()
        customer = (pick(r, "客户姓名") or "").strip()
        if not company or not seller or not customer:
            issues.append({"row_index": i, "reason": "必填字段为空(公司主体/销售/客户姓名)", "raw": r})
            continue

        sales_company = (pick(r, "销售公司") or "").strip() or company

        txs.append(
            Tx(
                biz_date=biz_date,
                lead_date=parse_date(pick(r, "客户进线日", "客户进线日期")),
                company=company,
                sales_company=sales_company,
                seller=seller,
                customer=customer,
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
        )

    return txs, fieldnames, issues


# ------------------ summarization ------------------

def summarize_orders(txs: List[Tx]):
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
        tail_rate = (received_tail / float(receivable_tail)) if (total_price and receivable_tail > 0) else None
        gross = first_payment + a["tail"] + a["full"]
        net = gross - a["refund"]
        refund_status = "未退款"
        if a["refund"] > 0 and net <= 0:
            refund_status = "已退款"
        elif a["refund"] > 0:
            refund_status = "部分退款"
        rows.append(
            [
                a["company"],
                a["sales_company"],
                a["seller"],
                a["customer"],
                a["phone"],
                a["lead_date"].isoformat() if a["lead_date"] else "",
                a["deposit"],
                a["signup"],
                first_payment,
                total_price if total_price else "",
                receivable_tail if total_price else "",
                received_tail if total_price else "",
                unpaid_tail if total_price else "",
                ("%.1f%%" % (tail_rate * 100)) if tail_rate is not None else "--",
                a["tail"],
                a["full"],
                a["refund"],
                gross,
                net,
                refund_status,
            ]
        )
    rows.sort(key=lambda x: (x[0], x[2], x[3]))
    headers = ["公司主体", "销售公司", "销售", "客户", "电话", "进线日期", "订金", "报名费", "首款计入口径", "产品总价", "应收尾款", "已收尾款", "未收尾款", "尾款回收率", "尾款", "全款", "退款", "实收业绩", "净业绩", "退款状态"]
    return headers, rows


def _week_key(d: dt.date) -> str:
    # 自然周：周一-周日；用周一作为周起始
    monday = d - dt.timedelta(days=d.weekday())
    sunday = monday + dt.timedelta(days=6)
    return "%s~%s" % (monday.isoformat(), sunday.isoformat())


def _month_key(d: dt.date) -> str:
    return "%04d-%02d" % (d.year, d.month)


def summarize_week_month_sales_company(order_rows: List[List[Any]]):
    # order_rows fields: [公司主体, 销售公司, 销售, 客户, 电话, 进线日期, ..., 实收业绩(17), 净业绩(18), ...]
    weekly = defaultdict(lambda: {"orders": 0, "gross": 0, "net": 0, "refund": 0})
    monthly = defaultdict(lambda: {"orders": 0, "gross": 0, "net": 0, "refund": 0})
    for r in order_rows:
        # 订单汇总没有开单日期，这里无法按开单日期周/月；改用“进线日期”会偏口径。
        # 因此：改用 Tx.biz_date 生成周/月更准确（下面会用 txs 生成）。
        pass
    return weekly, monthly


def summarize_week_month_from_txs(txs: List[Tx]):
    # 从流水 txs 以开单日期为时间维度做周/月汇总
    # 销售公司/销售粒度
    wk_sc = defaultdict(lambda: {"gross": 0, "net": 0, "refund": 0, "cnt": 0})
    mo_sc = defaultdict(lambda: {"gross": 0, "net": 0, "refund": 0, "cnt": 0})
    wk_seller = defaultdict(lambda: {"gross": 0, "net": 0, "refund": 0, "cnt": 0})
    mo_seller = defaultdict(lambda: {"gross": 0, "net": 0, "refund": 0, "cnt": 0})

    for t in txs:
        gross = t.deposit + t.signup + t.tail + t.full
        net = gross - t.refund
        wk = _week_key(t.biz_date)
        mo = _month_key(t.biz_date)

        k1 = (wk, t.sales_company or t.company)
        wk_sc[k1]["gross"] += gross
        wk_sc[k1]["net"] += net
        wk_sc[k1]["refund"] += t.refund
        wk_sc[k1]["cnt"] += 1

        k2 = (mo, t.sales_company or t.company)
        mo_sc[k2]["gross"] += gross
        mo_sc[k2]["net"] += net
        mo_sc[k2]["refund"] += t.refund
        mo_sc[k2]["cnt"] += 1

        k3 = (wk, t.seller)
        wk_seller[k3]["gross"] += gross
        wk_seller[k3]["net"] += net
        wk_seller[k3]["refund"] += t.refund
        wk_seller[k3]["cnt"] += 1

        k4 = (mo, t.seller)
        mo_seller[k4]["gross"] += gross
        mo_seller[k4]["net"] += net
        mo_seller[k4]["refund"] += t.refund
        mo_seller[k4]["cnt"] += 1

    return wk_sc, mo_sc, wk_seller, mo_seller


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
        company_str = ",".join(sorted(a["company"]))
        sales_company_str = ",".join(sorted(a["sales_company"]))
        primary_company = sorted(a["company"])[0] if a["company"] else "(空)"
        rows.append([seller, company_str, sales_company_str, a["customers"], a["gross"], a["net"], a["refund"], "%.1f%%" % (refund_rate * 100), a["rec"], a["got"], a["un"], ("%.1f%%" % (tail_rate * 100)) if tail_rate is not None else "--", 0, primary_company])
    rows.sort(key=lambda x: x[5], reverse=True)
    for i, r in enumerate(rows, 1):
        r[12] = i
    headers = ["销售", "公司主体(集合)", "销售公司(集合)", "客户数", "实收业绩", "净业绩", "退款金额", "退款率", "应收尾款", "已收尾款", "未收尾款", "尾款回收率", "排名", "主企业"]
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


def filter_signup_no_tail_no_full_customers(order_rows: List[List[Any]]):
    # 报名费>0 且 尾款=0 且 全款=0（按订单汇总口径）
    rows = []
    for r in order_rows:
        signup = r[7] if isinstance(r[7], int) else 0
        tail = r[14] if isinstance(r[14], int) else 0
        full = r[15] if isinstance(r[15], int) else 0
        if signup > 0 and tail == 0 and full == 0:
            rows.append(r)
    return rows


# ------------------ sheet writers ------------------

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


def write_import_issues_sheet(wb, data_path: str, fieldnames: List[str], issues: List[Dict[str, Any]]):
    ws = wb.create_sheet("问题明细")
    ws.sheet_view.showGridLines = False

    base_headers = ["行号(CSV)", "问题原因"]
    headers = base_headers + (fieldnames or [])

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.cell(row=1, column=1, value="导入问题明细（数据源：%s）" % data_path).fill = TITLE_FILL
    ws.cell(row=1, column=1).font = Font(bold=True, color="FFFFFF", size=14)
    ws.cell(row=1, column=1).alignment = LEFT

    for i, h in enumerate(headers, 1):
        ws.cell(row=2, column=i, value=h)
    set_header_row(ws, 2, len(headers))

    r = 3
    for item in issues:
        raw = item.get("raw") or {}
        ws.cell(row=r, column=1, value=item.get("row_index"))
        ws.cell(row=r, column=2, value=item.get("reason"))
        for j, fn in enumerate(fieldnames or [], start=3):
            ws.cell(row=r, column=j, value=raw.get(fn, ""))
        style_row(ws, r, len(headers), SUB_FILL if r % 2 == 1 else None, align=LEFT)
        r += 1

    widths = [10, 26] + [14] * max(1, len(headers) - 2)
    set_col_width(ws, widths[: len(headers)])
    ws.freeze_panes = "A3"


def build_dashboard(wb, txs, order_rows, seller_rows, company_rows, lead_summary_rows, abnormal_rows):
    ws = wb.create_sheet("业绩看板", 0)
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:L1")
    ws["A1"] = "📊 电商业绩分析看板"
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = CENTER
    ws.merge_cells("A2:L2")
    ws["A2"] = "汇报日期：%s  |  数据来源：%s  |  订单唯一口径=客户姓名+客户电话" % (dt.date.today().isoformat(), _detect_data_path())
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

    # Sales chart
    seller_base = 20
    ws.cell(row=2, column=seller_base, value="销售")
    ws.cell(row=2, column=seller_base + 1, value="净业绩")
    ws.cell(row=2, column=seller_base + 2, value="主企业")
    top_sellers = seller_rows[:10]
    for i, row in enumerate(top_sellers, 3):
        ws.cell(row=i, column=seller_base, value=row[0])
        ws.cell(row=i, column=seller_base + 1, value=row[5])
        ws.cell(row=i, column=seller_base + 2, value=row[13])

    company_base = 30
    ws.cell(row=2, column=company_base, value="公司主体")
    ws.cell(row=2, column=company_base + 1, value="净业绩")
    top_companies = company_rows[:10]
    for i, row in enumerate(top_companies, 3):
        ws.cell(row=i, column=company_base, value=row[0])
        ws.cell(row=i, column=company_base + 1, value=row[3])

    chart1 = BarChart()
    chart1.type = "bar"
    chart1.title = "销售净业绩排行（按企业分色）"
    chart1.width = 17
    chart1.height = 9
    chart1.add_data(Reference(ws, min_col=seller_base + 1, min_row=2, max_row=2 + len(top_sellers)), titles_from_data=True)
    chart1.set_categories(Reference(ws, min_col=seller_base, min_row=3, max_row=2 + len(top_sellers)))
    if chart1.series:
        s = chart1.series[0]
        s.dPt = []
        for idx, row in enumerate(top_sellers):
            company = row[13] or "(空)"
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = _company_color(company)
            s.dPt.append(pt)
    ws.add_chart(chart1, "A9")

    chart2 = BarChart()
    chart2.type = "col"
    chart2.title = "公司主体净业绩对比（按企业分色）"
    chart2.width = 17
    chart2.height = 9
    chart2.add_data(Reference(ws, min_col=company_base + 1, min_row=2, max_row=2 + len(top_companies)), titles_from_data=True)
    chart2.set_categories(Reference(ws, min_col=company_base, min_row=3, max_row=2 + len(top_companies)))
    if chart2.series:
        s = chart2.series[0]
        s.dPt = []
        for idx, row in enumerate(top_companies):
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = _company_color(row[0])
            s.dPt.append(pt)
    ws.add_chart(chart2, "G9")

    set_col_width(ws, [16, 14, 14, 14, 14, 16, 4, 16, 14, 20, 10, 22])
    ws.freeze_panes = "A9"


def write_trend_overview_sheet(wb, txs: List[Tx]):
    """一个表里放周/月 + 销售公司、周/月 + 销售的汇总，以及对应折线图。"""
    wk_sc, mo_sc, wk_seller, mo_seller = summarize_week_month_from_txs(txs)

    ws = wb.create_sheet("周月趋势总览")
    ws.sheet_view.showGridLines = False

    # 区块1：周 + 销售公司
    ws.merge_cells("A1:F1")
    ws["A1"] = "自然周-销售公司业绩（按开单日期）"
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = WHITE_BOLD
    ws["A1"].alignment = LEFT

    headers1 = ["自然周", "销售公司", "流水笔数", "实收业绩", "净业绩", "退款金额"]
    for i, h in enumerate(headers1, 1):
        ws.cell(row=2, column=i, value=h)
    set_header_row(ws, 2, len(headers1))

    rows1 = []
    for (wk, sc), v in wk_sc.items():
        rows1.append([wk, sc, v["cnt"], v["gross"], v["net"], v["refund"]])
    rows1.sort(key=lambda x: (x[0], x[1]))
    r = 3
    for row in rows1:
        for i, v in enumerate(row, 1):
            ws.cell(row=r, column=i, value=v)
        style_row(ws, r, len(headers1), SUB_FILL if r % 2 == 1 else None)
        r += 1

    # 周-销售公司净业绩折线（按销售公司分系列太复杂，这里先给总净业绩折线）
    # 先聚合每周总净业绩
    wk_total = defaultdict(int)
    for (wk, sc), v in wk_sc.items():
        wk_total[wk] += v["net"]
    wk_total_rows = sorted([[wk, net] for wk, net in wk_total.items()], key=lambda x: x[0])
    chart_start = r + 1
    ws.cell(row=chart_start, column=1, value="自然周")
    ws.cell(row=chart_start, column=2, value="总净业绩")
    set_header_row(ws, chart_start, 2)
    rr = chart_start + 1
    for wk, net in wk_total_rows:
        ws.cell(row=rr, column=1, value=wk)
        ws.cell(row=rr, column=2, value=net)
        rr += 1

    lc1 = LineChart()
    lc1.title = "周度总净业绩趋势"
    lc1.height = 8
    lc1.width = 20
    lc1.add_data(Reference(ws, min_col=2, min_row=chart_start, max_row=rr - 1), titles_from_data=True)
    lc1.set_categories(Reference(ws, min_col=1, min_row=chart_start + 1, max_row=rr - 1))
    ws.add_chart(lc1, "H2")

    # 区块2：月 + 销售公司
    base2 = rr + 2
    ws.merge_cells(start_row=base2, start_column=1, end_row=base2, end_column=6)
    ws.cell(row=base2, column=1, value="月度-销售公司业绩（按开单日期）").fill = TITLE_FILL
    ws.cell(row=base2, column=1).font = WHITE_BOLD
    ws.cell(row=base2, column=1).alignment = LEFT

    for i, h in enumerate(["月份", "销售公司", "流水笔数", "实收业绩", "净业绩", "退款金额"], 1):
        ws.cell(row=base2 + 1, column=i, value=h)
    set_header_row(ws, base2 + 1, 6)

    rows2 = []
    for (mo, sc), v in mo_sc.items():
        rows2.append([mo, sc, v["cnt"], v["gross"], v["net"], v["refund"]])
    rows2.sort(key=lambda x: (x[0], x[1]))
    r2 = base2 + 2
    for row in rows2:
        for i, v in enumerate(row, 1):
            ws.cell(row=r2, column=i, value=v)
        style_row(ws, r2, 6, SUB_FILL if r2 % 2 == 1 else None)
        r2 += 1

    # 区块3：周/月 + 销售（两张表）
    base3 = r2 + 2
    ws.merge_cells(start_row=base3, start_column=1, end_row=base3, end_column=6)
    ws.cell(row=base3, column=1, value="自然周-个人业绩（按开单日期）").fill = TITLE_FILL
    ws.cell(row=base3, column=1).font = WHITE_BOLD
    ws.cell(row=base3, column=1).alignment = LEFT

    for i, h in enumerate(["自然周", "销售", "流水笔数", "实收业绩", "净业绩", "退款金额"], 1):
        ws.cell(row=base3 + 1, column=i, value=h)
    set_header_row(ws, base3 + 1, 6)

    rows3 = []
    for (wk, seller), v in wk_seller.items():
        rows3.append([wk, seller, v["cnt"], v["gross"], v["net"], v["refund"]])
    rows3.sort(key=lambda x: (x[0], -x[4]))
    r3 = base3 + 2
    for row in rows3:
        for i, v in enumerate(row, 1):
            ws.cell(row=r3, column=i, value=v)
        style_row(ws, r3, 6, SUB_FILL if r3 % 2 == 1 else None)
        r3 += 1

    base4 = r3 + 2
    ws.merge_cells(start_row=base4, start_column=1, end_row=base4, end_column=6)
    ws.cell(row=base4, column=1, value="月度-个人业绩（按开单日期）").fill = TITLE_FILL
    ws.cell(row=base4, column=1).font = WHITE_BOLD
    ws.cell(row=base4, column=1).alignment = LEFT

    for i, h in enumerate(["月份", "销售", "流水笔数", "实收业绩", "净业绩", "退款金额"], 1):
        ws.cell(row=base4 + 1, column=i, value=h)
    set_header_row(ws, base4 + 1, 6)

    rows4 = []
    for (mo, seller), v in mo_seller.items():
        rows4.append([mo, seller, v["cnt"], v["gross"], v["net"], v["refund"]])
    rows4.sort(key=lambda x: (x[0], -x[4]))
    r4 = base4 + 2
    for row in rows4:
        for i, v in enumerate(row, 1):
            ws.cell(row=r4, column=i, value=v)
        style_row(ws, r4, 6, SUB_FILL if r4 % 2 == 1 else None)
        r4 += 1

    set_col_width(ws, [22, 16, 10, 12, 12, 12, 3, 18, 14, 14, 14, 14])
    ws.freeze_panes = "A3"


def reorder_sheets(wb: openpyxl.Workbook):
    """把进线相关 sheet 放前面，其次看板，再是周月趋势，再是其它。"""
    desired = [
        "进线产值汇总",
        "进线产值-销售",
        "进线产值明细",
        "业绩看板",
        "周月趋势总览",
    ]
    # keep any existing ordering for the rest
    name_to_sheet = {s.title: s for s in wb.worksheets}
    ordered = []
    for n in desired:
        if n in name_to_sheet:
            ordered.append(name_to_sheet[n])
    for s in wb.worksheets:
        if s not in ordered:
            ordered.append(s)
    wb._sheets = ordered


# ------------------ main ------------------

def main():
    data_path = _detect_data_path()
    txs, fieldnames, import_issues = load_txs_with_issues()

    order_headers, order_rows = summarize_orders(txs)
    seller_headers, seller_rows = summarize_seller(order_rows)
    sales_company_headers, sales_company_rows = summarize_group(order_rows, 1, "销售公司")
    company_headers, company_rows = summarize_group(order_rows, 0, "公司主体")
    lead_summary_rows, lead_seller_rows, lead_detail_rows = summarize_lead_value(txs)
    abnormal_headers, abnormal_rows = build_abnormal_checks(txs)

    # 报名费无尾款订单清单
    signup_no_tail = filter_signup_no_tail_no_full_customers(order_rows)
    signup_headers = order_headers

    raw_headers = ["开单日期", "公司主体", "销售公司", "销售", "客户姓名", "年龄", "客户电话", "客户进线日", "订金", "报名费", "尾款", "全款", "退款", "产品总价", "备注"]
    raw_rows = [[t.biz_date.isoformat(), t.company, t.sales_company, t.seller, t.customer, t.age, t.phone, t.lead_date.isoformat() if t.lead_date else "", t.deposit, t.signup, t.tail, t.full, t.refund, t.total_price if t.total_price else "", t.note] for t in txs]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 先创建进线相关 sheet（后续会 reorder）
    ws6 = wb.create_sheet("进线产值汇总")
    write_sheet_table(ws6, "进线产值汇总", ["进线日期", "客户数", "订金", "报名费", "尾款", "全款", "退款", "实收产值", "净产值"], lead_summary_rows, widths=[12, 10, 10, 10, 10, 10, 10, 12, 12])

    ws7 = wb.create_sheet("进线产值-销售")
    write_sheet_table(ws7, "进线产值-销售分析", ["进线日期", "销售", "客户数", "实收产值", "净产值", "退款金额", "退款率"], lead_seller_rows, widths=[12, 10, 10, 12, 12, 12, 10])

    ws8 = wb.create_sheet("进线产值明细")
    write_sheet_table(ws8, "进线产值明细", ["进线日期", "业务日期", "公司主体", "销售公司", "销售", "客户", "电话", "订金", "报名费", "尾款", "全款", "退款", "实收产值", "净产值", "备注"], lead_detail_rows, widths=[12, 12, 10, 12, 10, 10, 14, 10, 10, 10, 10, 10, 12, 12, 20])

    # 看板
    build_dashboard(wb, txs, order_rows, seller_rows, company_rows, lead_summary_rows, abnormal_rows)

    # 周/月趋势总览
    write_trend_overview_sheet(wb, txs)

    # 其它 sheet
    ws1 = wb.create_sheet("原始流水")
    write_sheet_table(ws1, "原始流水", raw_headers, raw_rows, widths=[12, 10, 12, 10, 10, 8, 14, 12, 10, 10, 10, 10, 10, 10, 20])

    ws2 = wb.create_sheet("客户订单汇总")
    write_sheet_table(ws2, "客户订单汇总", order_headers, order_rows, widths=[10, 12, 10, 10, 14, 12, 10, 10, 12, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10], highlight={"col": "退款状态", "rules": {"已退款": RED_FONT, "部分退款": ORANGE_FONT, "未退款": GREEN_FONT}})

    ws3 = wb.create_sheet("销售业绩汇总")
    write_sheet_table(ws3, "销售业绩汇总", seller_headers, seller_rows, widths=[10, 18, 18, 10, 12, 12, 10, 10, 10, 10, 10, 10, 8, 10])

    ws4 = wb.create_sheet("销售公司业绩汇总")
    write_sheet_table(ws4, "销售公司业绩汇总", sales_company_headers, sales_company_rows)

    ws5 = wb.create_sheet("公司主体业绩汇总")
    write_sheet_table(ws5, "公司主体业绩汇总", company_headers, company_rows)

    ws10 = wb.create_sheet("报名费无尾款客户")
    write_sheet_table(ws10, "报名费>0 且 尾款=0 且 全款=0（按订单汇总）", signup_headers, signup_no_tail)

    ws9 = wb.create_sheet("异常校验")
    write_sheet_table(ws9, "异常校验", abnormal_headers, abnormal_rows, widths=[14, 10, 10, 12, 14, 28])

    write_import_issues_sheet(wb, data_path, fieldnames, import_issues)

    reorder_sheets(wb)

    wb.save(OUTPUT_FILE)
    print("✅ 已生成汇报版仪表盘：%s" % OUTPUT_FILE)
    print("ℹ️ 使用数据源：%s" % data_path)
    print("ℹ️ 导入成功行数：%s，问题行数：%s" % (len(txs), len(import_issues)))


if __name__ == "__main__":
    main()
