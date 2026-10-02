# -*- coding: utf-8 -*-
"""acct-tools —— 会计工具箱（python38 技能）

协议（SkillRunner 约定）：
  输入：暂存目录 request.json = {"task":"acct-tools","inputs":["<xlsx/csv 绝对路径>"],"action":"...",...}
        —— 参数直接在顶层（可选再套一层 params，两种都支持）
  输出：stdout 最后一行 JSON = {"ok":true/false,"message":"...","data":{...}}

覆盖会计高频刚需（宿主原先都没有）：
  trial-balance  试算平衡 / 科目余额表：按科目汇总借贷，检查借贷是否相等
  aging          账龄分析：按到期日/发生日把应收应付分到 0-30/31-60/61-90/90+ 区间
  depreciation   折旧计算：直线法 / 双倍余额递减法，输出逐期折旧表

输入约定（表头名可用同义词，见下方 HEADER_ALIAS）：
  trial-balance 需要：科目、借方、贷方（至少一组）
  aging         需要：客户/单号、金额、日期（或账期天数）
  depreciation  不需要输入文件；参数给原值/残值/年限即可
"""
import json
import os
import sys
from datetime import datetime, date

try:
    import openpyxl
except Exception as e:  # pragma: no cover
    print(json.dumps({"ok": False, "message": "缺少 openpyxl: %s" % e}, ensure_ascii=False))
    sys.exit(0)


def stage_dir():
    return os.path.dirname(os.path.abspath(__file__))


def load_request():
    with open(os.path.join(stage_dir(), "request.json"), encoding="utf-8") as f:
        return json.load(f)


def parse_csv(text):
    rows, row, cur, inq = [], [], [], False
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if inq:
            if c == '"':
                if i + 1 < n and text[i + 1] == '"':
                    cur.append('"')
                    i += 1
                else:
                    inq = False
            else:
                cur.append(c)
        else:
            if c == '"' and not cur:
                inq = True
            elif c == ',':
                row.append("".join(cur))
                cur = []
            elif c in ('\r', '\n'):
                if c == '\r' and i + 1 < n and text[i + 1] == '\n':
                    i += 1
                row.append("".join(cur))
                rows.append(row)
                row, cur = [], []
            else:
                cur.append(c)
        i += 1
    if cur or row:
        row.append("".join(cur))
        rows.append(row)
    return rows


def read_table(path, sheet=None, header_row=1):
    """读 xlsx/csv → (表头 list, 数据行 list[list])。表头行 1-based。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, "rb") as f:
            raw = f.read()
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("gb18030")
        rows = parse_csv(text)
    else:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if (sheet and sheet in wb.sheetnames) else wb.worksheets[0]
        rows = []
        for r in ws.iter_rows(values_only=True):
            rows.append(["" if v is None else v for v in r])
        wb.close()
    if not rows:
        return [], []
    hi = max(0, int(header_row) - 1)
    if hi >= len(rows):
        return [], []
    header = [str(h).strip() if h is not None else "" for h in rows[hi]]
    return header, rows[hi + 1:]


def r2(x):
    """四舍五入到 2 位并消除浮点负零（round(-0.0,2) 会显示成 '-0'，报表里不该出现）。"""
    v = round(float(x or 0.0), 2)
    return 0.0 if v == 0 else v


def to_num(v):
    """把单元格值转成 float；空/非数字返回 None（绝不静默当 0）。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("，", "")
    if s == "" or s in ("-", "—", "/"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    s = s.replace("￥", "").replace("¥", "").replace("元", "")
    try:
        x = float(s)
        return -x if neg else x
    except ValueError:
        return None


def find_col(header, names):
    """按同义词找列下标；未命中返回 -1。"""
    for i, h in enumerate(header):
        hs = (h or "").strip().lower()
        for n in names:
            if hs == n.lower():
                return i
    # 退一步：包含匹配
    for i, h in enumerate(header):
        hs = (h or "").strip().lower()
        for n in names:
            if n.lower() in hs:
                return i
    return -1


def parse_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d",
                "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


# ---------------- 试算平衡 ----------------

def do_trial_balance(header, rows, params):
    c_subj = find_col(header, ["科目", "科目名称", "account", "subject", "会计科目"])
    c_debit = find_col(header, ["借方", "借方金额", "debit", "借"])
    c_credit = find_col(header, ["贷方", "贷方金额", "credit", "贷"])
    if c_subj < 0:
        return None, "未找到「科目」列（可用表头：科目/科目名称/account/subject）"
    if c_debit < 0 and c_credit < 0:
        return None, "未找到「借方」或「贷方」列"

    agg = {}
    order = []
    bad = 0
    for r in rows:
        subj = str(r[c_subj]).strip() if c_subj < len(r) and r[c_subj] is not None else ""
        if subj == "":
            continue
        d = to_num(r[c_debit]) if (c_debit >= 0 and c_debit < len(r)) else None
        c = to_num(r[c_credit]) if (c_credit >= 0 and c_credit < len(r)) else None
        if (d is None) and (c is None):
            bad += 1
            continue
        if subj not in agg:
            agg[subj] = [0.0, 0.0]
            order.append(subj)
        agg[subj][0] += (d or 0.0)
        agg[subj][1] += (c or 0.0)

    out = []
    td = tc = 0.0
    for s in order:
        d, c = agg[s]
        bal = d - c
        td += d
        tc += c
        out.append({"科目": s, "借方合计": round(d, 2), "贷方合计": round(c, 2),
                    "余额": round(bal, 2), "方向": "借" if bal >= 0 else "贷"})
    diff = round(td - tc, 2)
    return {
        "rows": out,
        "totalDebit": round(td, 2),
        "totalCredit": round(tc, 2),
        "difference": diff,
        "balanced": abs(diff) < 0.005,
        "skippedRows": bad,
        "subjectCount": len(out),
    }, None


# ---------------- 账龄分析 ----------------

def do_aging(header, rows, params):
    c_name = find_col(header, ["客户", "单位", "往来单位", "供应商", "customer", "名称", "对方"])
    c_no = find_col(header, ["单号", "凭证号", "发票号", "订单号", "编号", "no", "invoice"])
    c_amt = find_col(header, ["金额", "应收", "应付", "余额", "amount", "未收", "未付"])
    c_date = find_col(header, ["日期", "发生日", "开票日", "到期日", "date", "开票日期", "业务日期"])
    if c_amt < 0:
        return None, "未找到「金额」列（可用：金额/应收/应付/余额/amount）"
    if c_date < 0 and c_name < 0:
        return None, "至少需要「日期」或「客户」列之一"

    asof = parse_date(params.get("asOf")) or date.today()
    buckets = params.get("buckets") or [30, 60, 90]
    try:
        b1, b2, b3 = [int(x) for x in buckets[:3]]
    except Exception:
        b1, b2, b3 = 30, 60, 90
    labels = ["0-%d天" % b1, "%d-%d天" % (b1 + 1, b2), "%d-%d天" % (b2 + 1, b3), "%d天以上" % (b3 + 1)]

    detail = []
    tot = [0.0, 0.0, 0.0, 0.0]
    bad = 0
    for r in rows:
        amt = to_num(r[c_amt]) if c_amt < len(r) else None
        if amt is None or amt == 0:
            bad += 1
            continue
        d = parse_date(r[c_date]) if (c_date >= 0 and c_date < len(r)) else None
        if d is None:
            bad += 1
            continue
        days = (asof - d).days
        if days < 0:
            days = 0
        if days <= b1:
            bi = 0
        elif days <= b2:
            bi = 1
        elif days <= b3:
            bi = 2
        else:
            bi = 3
        tot[bi] += amt
        nm = ""
        if c_name >= 0 and c_name < len(r):
            nm = str(r[c_name]).strip()
        elif c_no >= 0 and c_no < len(r):
            nm = str(r[c_no]).strip()
        detail.append({"名称": nm, "日期": d.isoformat(), "天数": days,
                       "金额": round(amt, 2), "区间": labels[bi]})

    by_party = {}
    for it in detail:
        k = it["名称"] or "(未填名称)"
        if k not in by_party:
            by_party[k] = [0.0, 0.0, 0.0, 0.0]
        bi = labels.index(it["区间"])
        by_party[k][bi] += it["金额"]

    summary = []
    for k in sorted(by_party.keys()):
        v = by_party[k]
        summary.append({"名称": k, labels[0]: round(v[0], 2), labels[1]: round(v[1], 2),
                        labels[2]: round(v[2], 2), labels[3]: round(v[3], 2),
                        "合计": round(sum(v), 2)})

    return {
        "asOf": asof.isoformat(),
        "labels": labels,
        "totals": [round(x, 2) for x in tot],
        "grandTotal": round(sum(tot), 2),
        "summary": summary,
        "detail": detail[:500],
        "skippedRows": bad,
    }, None


# ---------------- 折旧 ----------------

def do_depreciation(params):
    try:
        cost = float(params.get("cost"))
    except (TypeError, ValueError):
        return None, "缺少 cost（原值）"
    salvage = float(params.get("salvage") or 0)
    years = int(params.get("years") or 5)
    method = (params.get("method") or "straight").strip().lower()
    if years <= 0:
        return None, "years 必须为正整数"
    if cost <= 0:
        return None, "cost 必须为正数"
    if salvage < 0 or salvage >= cost:
        return None, "salvage 必须 >= 0 且小于 cost"

    depreciable = cost - salvage
    rows = []
    book = cost
    if method in ("straight", "直线", "直线法"):
        per = depreciable / years
        for y in range(1, years + 1):
            d = per
            if y == years:                      # 末期兜平，消除浮点误差
                d = book - salvage
            book -= d
            rows.append({"年": y, "折旧额": round(d, 2), "累计折旧": round(cost - book, 2),
                         "账面净值": round(book, 2)})
        msg = "直线法：每年折旧 %.2f 元，共 %d 年" % (per, years)
    elif method in ("ddb", "双倍余额", "双倍余额递减", "double"):
        rate = 2.0 / years
        remain = years
        for y in range(1, years + 1):
            d = book * rate
            remain -= 1
            # 最后两年转直线，且不低于残值
            if remain <= 1:
                d = book - salvage
            if book - d < salvage:
                d = book - salvage
            if d < 0:
                d = 0.0
            book -= d
            rows.append({"年": y, "折旧额": round(d, 2), "累计折旧": round(cost - book, 2),
                         "账面净值": round(book, 2)})
        msg = "双倍余额递减法：年折旧率 %.2f%%，共 %d 年" % (rate * 100, years)
    else:
        return None, "未知 method: %s（支持 straight / ddb）" % method

    return {"method": method, "cost": cost, "salvage": salvage, "years": years,
            "rows": rows, "message": msg}, None


# ---------------- 增值税计算 ----------------

def do_vat(params):
    """增值税：一般计税 / 简易计税，按税率或征收率算销项、进项、应纳。"""
    try:
        sales = float(params.get("sales") or 0)          # 不含税销售额
    except (TypeError, ValueError):
        return None, "sales 必须为数字（不含税销售额）"
    if sales < 0:
        return None, "sales 不能为负"

    try:
        purchases = float(params.get("purchases") or 0)  # 不含税采购额
    except (TypeError, ValueError):
        return None, "purchases 必须为数字（不含税采购额）"

    mode = (params.get("mode") or "general").strip().lower()   # general=一般计税 simplified=简易
    try:
        rate = float(params.get("rate") if params.get("rate") is not None
                     else (0.13 if mode == "general" else 0.03))
    except (TypeError, ValueError):
        return None, "rate 必须为数字（如 0.13 表示 13%）"
    if rate < 0 or rate > 1:
        return None, "rate 应在 0~1 之间（0.13 表示 13%）"

    if mode == "simplified":
        input_rate = 0.0
        note = "简易计税：不得抵扣进项"
    else:
        input_rate = rate
        try:
            input_rate = float(params.get("inputRate") if params.get("inputRate") is not None else rate)
        except (TypeError, ValueError):
            return None, "inputRate 必须为数字"
        note = "一般计税：销项 - 进项"

    output_tax = sales * rate
    input_tax = purchases * input_rate
    payable = output_tax - input_tax
    rows = [
        ["不含税销售额", round(sales, 2)],
        ["销项税率", "%.2f%%" % (rate * 100)],
        ["销项税额", round(output_tax, 2)],
        ["不含税采购额", round(purchases, 2)],
        ["进项税率", "%.2f%%" % (input_rate * 100)],
        ["进项税额", round(input_tax, 2)],
        [("应纳增值税" if payable >= 0 else "留抵税额"), round(abs(payable), 2)],
        ["价税合计（销售）", round(sales + output_tax, 2)],
    ]
    return {
        "mode": mode, "rate": rate, "inputRate": input_rate,
        "sales": round(sales, 2), "outputTax": round(output_tax, 2),
        "purchases": round(purchases, 2), "inputTax": round(input_tax, 2),
        "payable": round(payable, 2),
        "isCredit": payable < 0,
        "rows": rows,
        "note": note,
    }, None


# ---------------- 财务报表（资产负债表 / 利润表） ----------------

def do_statements(params):
    """根据科目余额（或直接给数）生成利润表 / 资产负债表。
    输入方式二选一：
      A) 传一个含「科目」「余额」列的表（可带「方向」列，借/贷）——自动按科目名归集；
      B) 直接在 params 给 items: [{"科目":"主营业务收入","金额":100000,"方向":"贷"}, ...]
    """
    items = params.get("items")
    if not items:
        return None, "需要 items 数组（或传含科目/余额列的表）"

    def get(name_list):
        tot = 0.0
        for it in items:
            nm = str(it.get("科目") or it.get("name") or "").strip()
            for n in name_list:
                if n in nm:
                    amt = float(it.get("金额") if it.get("金额") is not None else (it.get("amount") or 0))
                    d = str(it.get("方向") or it.get("dir") or "借").strip()
                    # 统一成"借正贷负"的净额口径
                    tot += amt if d.startswith("借") else -amt
                    break
        return tot

    # 利润表（收入为贷方→负，费用为借方→正；这里按净额反推）
    revenue = -get(["主营业务收入", "营业收入", "收入"])
    cost = get(["主营业务成本", "营业成本", "成本"])
    tax_surcharge = get(["税金及附加", "营业税金"])
    selling = get(["销售费用"])
    admin = get(["管理费用"])
    finance_exp = get(["财务费用"])
    other_income = -get(["其他收益", "投资收益"])
    operating_profit = revenue - cost - tax_surcharge - selling - admin - finance_exp + other_income
    non_op_income = -get(["营业外收入"])
    non_op_expense = get(["营业外支出"])
    total_profit = operating_profit + non_op_income - non_op_expense
    income_tax = get(["所得税费用"])
    net_profit = total_profit - income_tax

    income_rows = [
        ["一、营业收入", r2(revenue)],
        ["减：营业成本", r2(cost)],
        ["    税金及附加", r2(tax_surcharge)],
        ["    销售费用", r2(selling)],
        ["    管理费用", r2(admin)],
        ["    财务费用", r2(finance_exp)],
        ["加：其他收益/投资收益", r2(other_income)],
        ["二、营业利润", r2(operating_profit)],
        ["加：营业外收入", r2(non_op_income)],
        ["减：营业外支出", r2(non_op_expense)],
        ["三、利润总额", r2(total_profit)],
        ["减：所得税费用", r2(income_tax)],
        ["四、净利润", r2(net_profit)],
    ]

    # 资产负债表。
    # ⚠ 符号约定：get() 返回的是"借正贷负"的净额。
    #   · 资产类（借方余额）→ get() 已是正数，**直接取用，不要再取负**
    #     （曾误写成 -get(...)，导致资产合计变成负数、报表不平，实测踩到）
    #   · 负债/权益类（贷方余额）→ get() 是负数，需取负转成正数
    cash = get(["库存现金", "银行存款", "货币资金"])
    ar = get(["应收账款"])
    other_recv = get(["其他应收款", "预付账款"])
    inventory = get(["库存商品", "原材料", "存货"])
    fixed = get(["固定资产"])
    accum_dep = -get(["累计折旧"])                      # 贷方 → 转正，作为资产的减项
    fixed_net = fixed - accum_dep
    current_assets = cash + ar + other_recv + inventory
    total_assets = current_assets + fixed_net

    # 负债与权益（贷方余额 → get() 为负，取负转正）
    ap = -get(["应付账款"])
    other_pay = -get(["其他应付款", "预收账款"])
    payroll = -get(["应付职工薪酬"])
    taxes_payable = -get(["应交税费"])
    total_liab = ap + other_pay + payroll + taxes_payable

    paid_in = -get(["实收资本", "股本"])
    surplus = -get(["盈余公积"])
    retained = -get(["本年利润", "未分配利润", "利润分配"])
    # 未分配利润是贷方余额；本期净利若未结转，也计入权益
    total_equity = paid_in + surplus + retained + net_profit

    balance_rows = [
        ["资产", ""],
        ["  货币资金", r2(cash)],
        ["  应收账款", r2(ar)],
        ["  其他应收款/预付", r2(other_recv)],
        ["  存货", r2(inventory)],
        ["  固定资产原值", r2(fixed)],
        ["  减：累计折旧", r2(accum_dep)],
        ["  固定资产净值", r2(fixed_net)],
        ["资产合计", r2(total_assets)],
        ["", ""],
        ["负债", ""],
        ["  应付账款", r2(ap)],
        ["  其他应付款/预收", r2(other_pay)],
        ["  应付职工薪酬", r2(payroll)],
        ["  应交税费", r2(taxes_payable)],
        ["负债合计", r2(total_liab)],
        ["", ""],
        ["所有者权益", ""],
        ["  实收资本", r2(paid_in)],
        ["  盈余公积", r2(surplus)],
        ["  未分配利润（含本期净利）", r2(retained + net_profit)],
        ["所有者权益合计", r2(total_equity)],
        ["负债和所有者权益合计", r2(total_liab + total_equity)],
    ]
    diff = round(total_assets - (total_liab + total_equity), 2)

    return {
        "income": income_rows,
        "balance": balance_rows,
        "revenue": r2(revenue),
        "netProfit": r2(net_profit),
        "totalAssets": r2(total_assets),
        "totalLiabEquity": r2(total_liab + total_equity),
        "balanceDiff": diff,
        "balanced": abs(diff) < 0.005,
    }, None


# ---------------- 写结果到 xlsx ----------------

def write_result(out_path, title, header, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "结果"
    ws.append([title])
    ws["A1"].font = openpyxl.styles.Font(bold=True, size=13)
    ws.append([])
    ws.append(list(header))
    for c in range(1, len(header) + 1):
        ws.cell(row=3, column=c).font = openpyxl.styles.Font(bold=True)
    for r in rows:
        ws.append(list(r))
    for col in range(1, len(header) + 1):
        from openpyxl.utils import get_column_letter
        best = 10
        for row in range(3, ws.max_row + 1):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            w = sum(2 if ord(ch) > 127 else 1 for ch in str(v)) + 2
            if w > best:
                best = w
        ws.column_dimensions[get_column_letter(col)].width = min(best, 40)
    if os.path.dirname(out_path) and not os.path.isdir(os.path.dirname(out_path)):
        os.makedirs(os.path.dirname(out_path))
    wb.save(out_path)


def main():
    try:
        req = load_request()
    except Exception as e:
        print(json.dumps({"ok": False, "message": "request.json 读取失败: %s" % e}, ensure_ascii=False))
        return 1

    params = dict(req)
    if isinstance(req.get("params"), dict):
        params.update(req["params"])

    action = (params.get("action") or "trial-balance").strip().lower()
    inputs = [p for p in (req.get("inputs") or []) if os.path.isfile(p)]

    # 输出路径
    out = (params.get("out") or "").strip()
    base_dir = os.path.dirname(inputs[0]) if inputs else os.getcwd()

    try:
        if action in ("vat", "增值税"):
            data, err = do_vat(params)
            if err:
                print(json.dumps({"ok": False, "message": err}, ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "增值税计算.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            write_result(out, "增值税计算表（%s）" % data["note"], ["项目", "金额"], data["rows"])
            if data["isCredit"]:
                verdict = "进项大于销项，本期形成留抵税额 %.2f 元" % abs(data["payable"])
            else:
                verdict = "本期应纳增值税 %.2f 元" % data["payable"]
            print(json.dumps({
                "ok": True,
                "message": "%s（销项 %.2f - 进项 %.2f）；已输出 %s" % (
                    verdict, data["outputTax"], data["inputTax"], out),
                "data": dict(data, out=out),
            }, ensure_ascii=False))
            return 0

        if action in ("statements", "报表", "fs"):
            items = params.get("items")
            if not items and inputs:
                # 没有直接给 items：从输入表读「科目/余额」列
                header, rows = read_table(inputs[0], params.get("sheet"), params.get("headerRow") or 1)
                c_subj = find_col(header, ["科目", "科目名称", "subject", "account"])
                c_amt = find_col(header, ["余额", "金额", "amount", "balance"])
                c_dir = find_col(header, ["方向", "借贷", "dir"])
                if c_subj < 0 or c_amt < 0:
                    print(json.dumps({"ok": False, "message": "需要「科目」与「余额」列（或直接传 items）",
                                      "data": {"header": header}}, ensure_ascii=False))
                    return 1
                items = []
                for r in rows:
                    subj = str(r[c_subj]).strip() if c_subj < len(r) and r[c_subj] is not None else ""
                    if not subj:
                        continue
                    amt = to_num(r[c_amt]) if c_amt < len(r) else None
                    if amt is None:
                        continue
                    d = "借"
                    if c_dir >= 0 and c_dir < len(r) and r[c_dir]:
                        d = str(r[c_dir]).strip()
                    items.append({"科目": subj, "金额": amt, "方向": d})
                params["items"] = items
            data, err = do_statements(params)
            if err:
                print(json.dumps({"ok": False, "message": err}, ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "财务报表.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            # 两张表写进同一工作簿的两个 sheet
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "利润表"
            ws1.append(["利润表"]); ws1["A1"].font = openpyxl.styles.Font(bold=True, size=13)
            ws1.append(["项目", "金额"])
            ws1["A2"].font = openpyxl.styles.Font(bold=True)
            ws1["B2"].font = openpyxl.styles.Font(bold=True)
            for r in data["income"]:
                ws1.append(list(r))
            ws2 = wb.create_sheet("资产负债表")
            ws2.append(["资产负债表"]); ws2["A1"].font = openpyxl.styles.Font(bold=True, size=13)
            for r in data["balance"]:
                ws2.append(list(r))
            for ws in (ws1, ws2):
                ws.column_dimensions["A"].width = 26
                ws.column_dimensions["B"].width = 16
            if os.path.dirname(out) and not os.path.isdir(os.path.dirname(out)):
                os.makedirs(os.path.dirname(out))
            wb.save(out)
            verdict = "资产 = 负债+权益 ✓" if data["balanced"] else \
                "资产与负债+权益不平 ✗（差 %.2f）" % data["balanceDiff"]
            print(json.dumps({
                "ok": True,
                "message": "营业收入 %.2f，净利润 %.2f，资产合计 %.2f；%s；已输出 %s" % (
                    data["revenue"], data["netProfit"], data["totalAssets"], verdict, out),
                "data": dict(data, out=out),
            }, ensure_ascii=False))
            return 0

        if action in ("depreciation", "dep"):
            data, err = do_depreciation(params)
            if err:
                print(json.dumps({"ok": False, "message": err}, ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "折旧表.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            write_result(out, "折旧计算表", ["年", "折旧额", "累计折旧", "账面净值"],
                         [[r["年"], r["折旧额"], r["累计折旧"], r["账面净值"]] for r in data["rows"]])
            print(json.dumps({"ok": True, "message": data["message"] + "；已输出 " + out,
                              "data": {"out": out, "rows": data["rows"]}}, ensure_ascii=False))
            return 0

        if not inputs:
            print(json.dumps({"ok": False, "message": "%s 需要一个输入文件（xlsx/csv）" % action},
                             ensure_ascii=False))
            return 1

        header, rows = read_table(inputs[0], params.get("sheet"), params.get("headerRow") or 1)
        if not header:
            print(json.dumps({"ok": False, "message": "读取不到表头（文件为空或 headerRow 不对）"},
                             ensure_ascii=False))
            return 1

        if action in ("trial-balance", "trial", "tb"):
            data, err = do_trial_balance(header, rows, params)
            if err:
                print(json.dumps({"ok": False, "message": err, "data": {"header": header}},
                                 ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "试算平衡表.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            write_result(out, "试算平衡表", ["科目", "借方合计", "贷方合计", "余额", "方向"],
                         [[r["科目"], r["借方合计"], r["贷方合计"], r["余额"], r["方向"]] for r in data["rows"]])
            verdict = "借贷平衡 ✓" if data["balanced"] else "借贷不平 ✗（差 %.2f）" % data["difference"]
            print(json.dumps({
                "ok": True,
                "message": "%s；%d 个科目，借方合计 %.2f，贷方合计 %.2f；已输出 %s" % (
                    verdict, data["subjectCount"], data["totalDebit"], data["totalCredit"], out),
                "data": dict(data, out=out),
            }, ensure_ascii=False))
            return 0

        if action in ("aging", "账龄"):
            data, err = do_aging(header, rows, params)
            if err:
                print(json.dumps({"ok": False, "message": err, "data": {"header": header}},
                                 ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "账龄分析.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            lbs = data["labels"]
            write_result(out, "账龄分析（截至 %s）" % data["asOf"],
                         ["名称"] + lbs + ["合计"],
                         [[s["名称"], s[lbs[0]], s[lbs[1]], s[lbs[2]], s[lbs[3]], s["合计"]]
                          for s in data["summary"]])
            print(json.dumps({
                "ok": True,
                "message": "账龄分析完成：%d 个往来单位，合计 %.2f 元；已输出 %s" % (
                    len(data["summary"]), data["grandTotal"], out),
                "data": dict(data, out=out),
            }, ensure_ascii=False))
            return 0

        print(json.dumps({"ok": False, "message": "未知 action: %s（支持 trial-balance / aging / depreciation）"
                          % action}, ensure_ascii=False))
        return 1

    except Exception as e:
        print(json.dumps({"ok": False, "message": "执行异常: %s" % e}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())
