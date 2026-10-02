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
  vat            增值税：一般计税 / 简易计税
  statements     财务报表：资产负债表 + 利润表（含平衡校验）
  payroll        工资表与个税：中国累计预扣法（2019 起口径），含社保公积金
  bank-recon     银行余额调节表：对账单 vs 日记账，输出调节表并验证调节后相等
  consolidation  合并报表抵消：多公司汇总 + 内部交易抵消分录
  journal        凭证/日记账生成：按分录生成记账凭证并登记日记账

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
    """按同义词找列下标；未命中返回 -1。

    ★ 匹配优先级（回归修复）：科目余额表常同时含「期初借方/期初贷方/本期借方/本期贷方/
      期末借方/期末贷方」六列，全部含「借方」或「贷方」字样。旧实现"退一步用包含匹配，
      返回第一个命中列"，于是 c_debit 命中「期初借方」、c_credit 命中「期初贷方」——
      技能悄悄读了期初列而不是用户以为的本期/期末列，符号与金额全错。
    现在的顺序：
      1) 完全相等（忽略大小写与首尾空白）—— 最可靠
      2) 包含匹配，但**按 names 的书写顺序**优先：调用方把最想要的写法放前面，
         如 ["本期借方","借方"] 会先找「本期借方」，找不到才退化到含「借方」的列
      3) 包含匹配时优先**更短的表头**：在多个候选里，"借方"（2字）比"期初借方"（4字）
         更可能是通用列；但这条只在同名前缀下生效，故放在第 2 条之后作为并列打破规则
    """
    def norm(s):
        return (s or "").strip().lower()

    # 1) 完全相等
    for i, h in enumerate(header):
        hs = norm(h)
        for n in names:
            if hs == norm(n):
                return i

    # 2) 按 names 顺序做包含匹配（先试最具体的写法）
    for n in names:
        nl = norm(n)
        if not nl:
            continue
        best = -1
        for i, h in enumerate(header):
            hs = norm(h)
            if nl in hs:
                # 同一次 names 尝试内出现多个候选：取表头更短的（更通用）
                if best < 0 or len(hs) < len(norm(header[best])):
                    best = i
        if best >= 0:
            return best
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
    c_subj = find_col(header, ["科目名称", "会计科目", "科目", "account", "subject"])
    # 列优先级（回归修复）：科目余额表常同时含期初/本期/期末三组借贷列。
    # 试算平衡的业务含义是"核对期末余额"，故优先取期末，其次取不带期间前缀的通用列，
    # 最后才退化到期初/本期。旧实现按"第一个包含命中"取列，会静默读到期初列，
    # 导致金额与方向全错（实测：成本/费用科目被算成负的借方合计、方向判成"贷"）。
    c_debit = find_col(header, ["期末借方", "期末余额借方", "借方余额", "借方", "借方金额", "debit", "借"])
    c_credit = find_col(header, ["期末贷方", "期末余额贷方", "贷方余额", "贷方", "贷方金额", "credit", "贷"])
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
    # ★ 本期净利只在**尚未结转**时计入权益（回归修复）。
    #   期末余额表若已含「本年利润」科目（本期净利已结转过去），再叠加 net_profit 就是
    #   重复计入 —— 实测：本年利润 35000 已入表，又加了一次净利 35000，导致
    #   负债+权益虚增 35000，资产负债表自报"不平（差 -35000）"，而数据本身是平的。
    #   判据：表里根本没出现利润类科目（retained == 0 且没有该科目的行）才补加。
    has_retained_subject = False
    for it in items:
        nm = str(it.get("科目") or it.get("name") or "").strip()
        if nm in ("本年利润", "未分配利润", "利润分配"):
            has_retained_subject = True
            break
    total_equity = paid_in + surplus + retained + (0.0 if has_retained_subject else net_profit)

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


# ---------------- 工资表与个税（累计预扣法） ----------------

# 中国居民工资薪金个税预扣率表（累计预扣法，2019 起）：
# (累计应纳税所得额上限, 税率, 速算扣除数)
IIT_BRACKETS = [
    (36000.0, 0.03, 0.0),
    (144000.0, 0.10, 2520.0),
    (300000.0, 0.20, 16920.0),
    (420000.0, 0.25, 31920.0),
    (660000.0, 0.30, 52920.0),
    (960000.0, 0.35, 85920.0),
    (float("inf"), 0.45, 181920.0),
]


def iit_of(cum_taxable):
    """按累计应纳税所得额查表得 (税率, 速算扣除数)。"""
    for cap, rate, deduct in IIT_BRACKETS:
        if cum_taxable <= cap:
            return rate, deduct
    return 0.45, 181920.0


# 年终奖（全年一次性奖金）单独计税用的月度税率表：
# 先把奖金 ÷12 查"按月换算后的综合所得税率表"，再对**全额**乘税率减速算扣除数。
# 注意这张表的速算扣除数是"月度"口径（与年度表的 2520/16920 等不同），
# 两者混用是最常见的算错原因。
BONUS_BRACKETS = [
    (3000.0, 0.03, 0.0),
    (12000.0, 0.10, 210.0),
    (25000.0, 0.20, 1410.0),
    (35000.0, 0.25, 2660.0),
    (55000.0, 0.30, 4410.0),
    (80000.0, 0.35, 7160.0),
    (float("inf"), 0.45, 15160.0),
]


def bonus_tax_of(bonus):
    """年终奖单独计税：奖金 ÷12 定档，再对全额计税。返回 (税额, 税率, 速算扣除数, 月均)。"""
    avg = bonus / 12.0
    for cap, rate, deduct in BONUS_BRACKETS:
        if avg <= cap:
            tax = bonus * rate - deduct
            if tax < 0:
                tax = 0.0
            return round(tax, 2), rate, deduct, round(avg, 2)
    tax = bonus * 0.45 - 15160.0
    return round(max(tax, 0.0), 2), 0.45, 15160.0, round(avg, 2)


# 社保公积金默认费率（个人部分）。各地差异大，这里的默认值仅作示例，
# 实际请用 params.socialRates 覆盖；单位部分另计，本工具只算个人扣缴。
DEFAULT_SOCIAL_RATES = {
    "pension": 0.08,     # 养老 8%
    "medical": 0.02,     # 医疗 2%
    "unemploy": 0.005,   # 失业 0.5%
    "housing": 0.12,     # 公积金 12%
}


def calc_social(base, params):
    """按缴费基数 × 费率算个人应缴社保与公积金。
    params.socialRates 可覆盖默认费率；params.socialBaseCap 可设基数上限。
    返回 (社保合计, 公积金, 明细dict)。
    """
    rates = dict(DEFAULT_SOCIAL_RATES)
    custom = params.get("socialRates")
    if isinstance(custom, dict):
        for k in list(rates.keys()):
            if k in custom:
                try:
                    rates[k] = float(custom[k])
                except (TypeError, ValueError):
                    pass
    cap = params.get("socialBaseCap")
    b = base
    if cap is not None:
        try:
            capv = float(cap)
            if capv > 0 and b > capv:
                b = capv
        except (TypeError, ValueError):
            pass
    detail = {}
    si = 0.0
    for k in ("pension", "medical", "unemploy"):
        v = round(b * rates[k], 2)
        detail[k] = v
        si += v
    hf = round(b * rates["housing"], 2)
    detail["housing"] = hf
    detail["base"] = round(b, 2)
    detail["rates"] = {k: rates[k] for k in rates}
    return round(si, 2), hf, detail


def do_bonus(params):
    """年终奖单独计税（全年一次性奖金）。
    参数：bonus（奖金金额）或 bonusList（多人：[{"姓名","奖金"}]）。
    可选：compareWithCombined=true 时同时给出"并入综合所得"的估算对比，
         需要提供年度应纳税所得额（annualTaxable，不含该奖金）。
    """
    lst = params.get("bonusList")
    rows = []
    if isinstance(lst, list) and lst:
        for it in lst:
            if not isinstance(it, dict):
                continue
            nm = str(it.get("姓名") or it.get("name") or "").strip()
            b = to_num(it.get("奖金") if it.get("奖金") is not None else it.get("bonus"))
            if b is None or b < 0:
                continue
            tax, rate, deduct, avg = bonus_tax_of(b)
            rows.append({"姓名": nm, "奖金": round(b, 2), "月均": avg,
                         "税率": "%.0f%%" % (rate * 100), "速算扣除数": deduct,
                         "税额": tax, "税后": round(b - tax, 2)})
    else:
        b = to_num(params.get("bonus"))
        if b is None or b < 0:
            return None, "需要 bonus（奖金金额）或 bonusList（多人列表）"
        tax, rate, deduct, avg = bonus_tax_of(b)
        rows.append({"姓名": "（单人）", "奖金": round(b, 2), "月均": avg,
                     "税率": "%.0f%%" % (rate * 100), "速算扣除数": deduct,
                     "税额": tax, "税后": round(b - tax, 2)})

    tot_b = sum(r["奖金"] for r in rows)
    tot_t = sum(r["税额"] for r in rows)

    # 可选：与"并入综合所得"对比
    compare = None
    if params.get("compareWithCombined"):
        at = to_num(params.get("annualTaxable"))
        if at is not None:
            # 单独计税：奖金按上面算；综合所得部分按年度表
            sep_bonus_tax = tot_t
            rate_a, quick_a = iit_of(max(at, 0.0))
            sep_other_tax = round(max(at * rate_a - quick_a, 0.0), 2)
            # 并入后：综合所得 + 全部奖金 一起查年度表
            rate_c, quick_c = iit_of(max(at + tot_b, 0.0))
            comb_tax = round(max((at + tot_b) * rate_c - quick_c, 0.0), 2)
            compare = {
                "annualTaxableExBonus": round(at, 2),
                "separate": {"bonusTax": sep_bonus_tax, "otherTax": sep_other_tax,
                             "total": round(sep_bonus_tax + sep_other_tax, 2)},
                "combined": {"totalTax": comb_tax},
                "savingBySeparate": round(comb_tax - (sep_bonus_tax + sep_other_tax), 2),
            }

    return {"rows": rows, "totalBonus": round(tot_b, 2), "totalTax": round(tot_t, 2),
            "totalNet": round(tot_b - tot_t, 2), "count": len(rows),
            "compare": compare}, None


def do_payroll(header, rows, params):
    """工资表与个税。
    输入列（同义词自动识别）：姓名、应发工资、社保、公积金、专项附加扣除；
    可选：月份（1-12，用于累计）。不给月份时按"本月即第 month 个月"处理，
    并把之前的累计视为 0（适合只算单月的场景）。
    """
    c_name = find_col(header, ["姓名", "员工", "名字", "name", "人员"])
    c_gross = find_col(header, ["应发工资", "应发", "工资", "gross", "应付工资", "应发合计"])
    c_si = find_col(header, ["社保", "社会保险", "五险", "social", "个人社保"])
    c_hf = find_col(header, ["公积金", "住房公积", "housing", "个人公积金"])
    c_extra = find_col(header, ["专项附加扣除", "专扣", "附加扣除", "extra", "专项附加"])
    c_month = find_col(header, ["月份", "月", "month"])
    # 社保缴费基数：给了就按基数×费率自动算社保公积金（见下方优先级说明）
    c_base = find_col(header, ["缴费基数", "社保基数", "基数", "base", "社保缴费基数"])

    if c_name < 0 or c_gross < 0:
        return None, "需要「姓名」与「应发工资」列（可用表头：姓名/员工/name；应发工资/工资/gross）"

    try:
        month = int(params.get("month") or 1)
    except (TypeError, ValueError):
        month = 1
    month = max(1, min(month, 12))
    basic = float(params.get("basicDeduction") if params.get("basicDeduction") is not None else 5000)

    out = []
    tot_gross = tot_tax = tot_net = tot_si = tot_hf = 0.0
    skipped = 0

    for r in rows:
        nm = str(r[c_name]).strip() if c_name < len(r) and r[c_name] is not None else ""
        if nm == "":
            continue
        gross = to_num(r[c_gross]) if c_gross < len(r) else None
        if gross is None:
            skipped += 1
            continue
        si = to_num(r[c_si]) if (c_si >= 0 and c_si < len(r)) else 0.0
        hf = to_num(r[c_hf]) if (c_hf >= 0 and c_hf < len(r)) else 0.0
        extra = to_num(r[c_extra]) if (c_extra >= 0 and c_extra < len(r)) else 0.0
        si = si or 0.0
        hf = hf or 0.0
        extra = extra or 0.0

        # 社保/公积金：给了金额就用金额；没给金额但给了缴费基数，就按基数×费率算。
        # 优先级：显式金额 > 基数×费率。两者都给时以金额为准（便于手工调整个别员工）。
        social_detail = None
        if si == 0 and hf == 0:
            base_v = None
            if c_base >= 0 and c_base < len(r):
                base_v = to_num(r[c_base])
            if base_v is None and params.get("socialBase") is not None:
                base_v = to_num(params.get("socialBase"))
            if base_v is None and params.get("useGrossAsBase"):
                base_v = gross
            if base_v is not None and base_v > 0:
                si, hf, social_detail = calc_social(base_v, params)
        m = month
        if c_month >= 0 and c_month < len(r):
            try:
                m = int(float(str(r[c_month]).strip()))
                m = max(1, min(m, 12))
            except (TypeError, ValueError):
                m = month

        # 累计预扣法。
        # 有两种用法，务必分清（口径不同，结果不同）：
        #   A) 算"某个月的这一期"（默认）：假设本月的收入/扣除与前几个月相同，
        #      于是累计 = 本月值 × 月份 m。适合快速估算或每月都一样的场景。
        #      ⚠ 这是**估算**：真实情况里月度收入与社保常有波动。
        #   B) 精确累计（推荐用于真实申报）：用 params.prior* 传入**截至上月的累计数**，
        #      本工具只负责加上本月并算出本期应预扣。此时月份 m 仅作展示。
        prior_income = params.get("priorIncome")
        prior_si = params.get("priorSocialInsurance")
        prior_hf = params.get("priorHousingFund")
        prior_extra = params.get("priorExtra")
        prior_tax = params.get("priorTaxPaid")
        has_prior = any(x is not None for x in (prior_income, prior_si, prior_hf, prior_extra, prior_tax))

        if has_prior:
            try:
                pi = float(prior_income or 0)
                psi = float(prior_si or 0)
                phf = float(prior_hf or 0)
                pex = float(prior_extra or 0)
                ptax = float(prior_tax or 0)
            except (TypeError, ValueError):
                return None, "prior* 参数必须为数字（截至上月的累计数）"
            cum_gross = pi + gross
            cum_deduct = (basic * m) + (psi + si) + (phf + hf) + (pex + extra)
            mode_note = "精确累计（含传入的截至上月累计）"
        else:
            cum_gross = gross * m
            cum_deduct = basic * m + (si + hf) * m + extra * m
            ptax = 0.0
            mode_note = "单月估算（假设前 %d 个月与本月相同）" % (m - 1) if m > 1 else "单月"

        cum_taxable = cum_gross - cum_deduct
        if cum_taxable < 0:
            cum_taxable = 0.0
        rate, quick = iit_of(cum_taxable)
        cum_tax = cum_taxable * rate - quick
        if cum_tax < 0:
            cum_tax = 0.0
        # 本期应预扣 = 累计应纳税额 - 已预扣（精确模式）；估算模式已预扣按 0
        tax = round(cum_tax - ptax, 2)
        if tax < 0:
            tax = 0.0
        net = round(gross - si - hf - tax, 2)

        out.append({
            "姓名": nm, "月份": m, "应发工资": round(gross, 2),
            "社保": round(si, 2), "公积金": round(hf, 2),
            "社保基数": (social_detail or {}).get("base", None),
            "专项附加扣除": round(extra, 2),
            "累计应纳税所得额": round(cum_taxable, 2),
            "累计应纳税额": round(cum_tax, 2),
            "已预扣": round(ptax, 2),
            "口径": mode_note,
            "税率": "%.0f%%" % (rate * 100),
            "速算扣除数": round(quick, 2),
            "个税": tax,
            "实发工资": net,
        })
        tot_gross += gross
        tot_si += si
        tot_hf += hf
        tot_tax += tax
        tot_net += net

    return {
        "month": month, "basicDeduction": basic,
        "rows": out,
        "totalGross": round(tot_gross, 2),
        "totalSocialInsurance": round(tot_si, 2),
        "totalHousingFund": round(tot_hf, 2),
        "totalTax": round(tot_tax, 2),
        "totalNet": round(tot_net, 2),
        "employeeCount": len(out),
        "skippedRows": skipped,
    }, None


# ---------------- 银行余额调节表 ----------------

def do_bank_recon(header, rows, params):
    """银行余额调节表。
    输入列：日期、摘要、金额、类型（收/付 或 借/贷）；可选：来源（银行/企业）。
    也可传两文件：inputs[0]=对账单, inputs[1]=日记账，分别读再合并。
    输出调节表：银行对账单余额 + 企业已收银行未收 - 企业已付银行未付
                = 企业日记账余额 + 银行已收企业未收 - 银行已付企业未付
    """
    try:
        bank_bal = float(params.get("bankBalance") or 0)
        book_bal = float(params.get("bookBalance") or 0)
    except (TypeError, ValueError):
        return None, "需要 bankBalance（对账单余额）与 bookBalance（日记账余额）"

    # 从传入的记录里挑出"未达账项"：按 params.unmatched 给的两组
    # 结构：unmatched = {"bankReceived": [...], "bankPaid": [...],
    #                    "bookReceived": [...], "bookPaid": [...]}
    um = params.get("unmatched") or {}

    def total(key):
        t = 0.0
        for x in (um.get(key) or []):
            if isinstance(x, dict):
                v = to_num(x.get("金额") or x.get("amount"))
            else:
                v = to_num(x)
            if v is not None:
                t += v
        return round(t, 2)

    br = total("bankReceived")    # 银行已收、企业未收 → 加在日记账侧
    bp = total("bankPaid")        # 银行已付、企业未付 → 减在日记账侧
    er = total("bookReceived")    # 企业已收、银行未收 → 加在对账单侧
    ep = total("bookPaid")        # 企业已付、银行未付 → 减在对账单侧

    adj_bank = round(bank_bal + er - ep, 2)
    adj_book = round(book_bal + br - bp, 2)
    diff = round(adj_bank - adj_book, 2)

    rows = [
        ["银行对账单余额", round(bank_bal, 2)],
        ["加：企业已收、银行未收", er],
        ["减：企业已付、银行未付", ep],
        ["调节后银行余额", adj_bank],
        ["", ""],
        ["企业日记账余额", round(book_bal, 2)],
        ["加：银行已收、企业未收", br],
        ["减：银行已付、企业未付", bp],
        ["调节后企业余额", adj_book],
        ["", ""],
        ["差额", diff],
    ]

    return {
        "bankBalance": round(bank_bal, 2), "bookBalance": round(book_bal, 2),
        "bookReceived": er, "bookPaid": ep,
        "bankReceived": br, "bankPaid": bp,
        "adjustedBank": adj_bank, "adjustedBook": adj_book,
        "difference": diff, "balanced": abs(diff) < 0.005,
        "rows": rows,
    }, None


# ---------------- 合并报表抵消 ----------------

def do_consolidation(params):
    """合并报表抵消。
    params.companies: [{"name":"A","assets":{...},"liabilities":{...},"equity":{...},
                        "revenue":{...},"expense":{...}}]
    params.eliminations: [{"desc":"内部销售抵消","debit":"主营业务收入","credit":"主营业务成本","amount":1000}]
    简化口径：只做"内部交易抵消分录"对损益与往来余额的影响汇总，
    输出各公司个别数 → 抵消分录 → 合并数 三栏对照。
    """
    companies = params.get("companies")
    if not companies:
        return None, "需要 companies 数组"
    elims = params.get("eliminations") or []

    def num(d, k):
        try:
            return float(d.get(k) or 0)
        except (TypeError, ValueError):
            return 0.0

    # 个别数汇总
    tot = {"资产": 0.0, "负债": 0.0, "权益": 0.0, "收入": 0.0, "费用": 0.0}
    detail = []
    for c in companies:
        a = sum(num(v, "金额") if isinstance(v, dict) else 0.0 for v in (c.get("assets") or {}).values()) \
            if isinstance(c.get("assets"), dict) else num(c, "assets")
        l = sum(num(v, "金额") if isinstance(v, dict) else 0.0 for v in (c.get("liabilities") or {}).values()) \
            if isinstance(c.get("liabilities"), dict) else num(c, "liabilities")
        e = sum(num(v, "金额") if isinstance(v, dict) else 0.0 for v in (c.get("equity") or {}).values()) \
            if isinstance(c.get("equity"), dict) else num(c, "equity")
        r = sum(num(v, "金额") if isinstance(v, dict) else 0.0 for v in (c.get("revenue") or {}).values()) \
            if isinstance(c.get("revenue"), dict) else num(c, "revenue")
        x = sum(num(v, "金额") if isinstance(v, dict) else 0.0 for v in (c.get("expense") or {}).values()) \
            if isinstance(c.get("expense"), dict) else num(c, "expense")
        detail.append({"name": c.get("name") or "?", "资产": a, "负债": l, "权益": e,
                       "收入": r, "费用": x, "利润": r - x})
        tot["资产"] += a
        tot["负债"] += l
        tot["权益"] += e
        tot["收入"] += r
        tot["费用"] += x

    # 抵消分录：按借贷科目归类，抵减对应大类
    elim_rows = []
    elim_effect = {"资产": 0.0, "负债": 0.0, "权益": 0.0, "收入": 0.0, "费用": 0.0}

    def classify(subj):
        s = str(subj or "")
        for kw, cat in (("收入", "收入"), ("成本", "费用"), ("费用", "费用"), ("资产", "资产"),
                        ("应收", "资产"), ("存货", "资产"), ("负债", "负债"),
                        ("应付", "负债"), ("权益", "权益"), ("资本", "权益")):
            if kw in s:
                return cat
        return None

    unmatched = []
    for el in elims:
        amt = to_num(el.get("amount"))
        if amt is None:
            continue
        dj = classify(el.get("debit"))
        cj = classify(el.get("credit"))
        if dj is None or cj is None:
            unmatched.append(el.get("desc") or str(el))
            continue
        # 抵消分录：借某科目 = 该类别减少（收入类借减）；贷某科目 = 该类别减少（费用/资产类贷减）
        # 简化：只在"同类相抵"（如内部收入 vs 内部成本）时把两边各自冲减
        elim_effect[dj] -= amt
        elim_effect[cj] -= amt
        elim_rows.append([el.get("desc") or "", el.get("debit") or "", el.get("credit") or "", amt])

    merged = {}
    for k in ("资产", "负债", "权益", "收入", "费用"):
        merged[k] = round(tot[k] + elim_effect[k], 2)
    merged["利润"] = round(merged["收入"] - merged["费用"], 2)

    # 合并口径勾稽：资产 = 负债 + 权益（利润并入权益）
    lhs = merged["资产"]
    rhs = round(merged["负债"] + merged["权益"] + merged["利润"], 2)
    diff = round(lhs - rhs, 2)

    return {
        "companies": detail,
        "eliminations": elim_rows,
        "unmatchedEliminations": unmatched,
        "individual": {k: round(v, 2) for k, v in tot.items()},
        "merged": merged,
        "balanceDiff": diff,
        "balanced": abs(diff) < 0.005,
    }, None


# ---------------- 凭证 / 日记账生成 ----------------

def do_journal(params):
    """按分录生成记账凭证并登记日记账。
    params.entries: [{"date":"2026-10-02","voucher":"记-001","summary":"收到货款",
                      "lines":[{"科目":"银行存款","借方":10000},{"科目":"应收账款","贷方":10000}]}]
    校验：每张凭证借贷必须相等；不等的凭证**不登记**并列入错误清单（绝不静默计入）。
    """
    entries = params.get("entries")
    if not entries:
        return None, "需要 entries 数组"

    vouchers = []
    journal = []
    errors = []
    vno = 0

    for e in entries:
        vno += 1
        lines = e.get("lines") or []
        if not lines:
            errors.append("第 %d 张凭证没有分录行" % vno)
            continue
        td = tc = 0.0
        norm = []
        for ln in lines:
            subj = str(ln.get("科目") or "").strip()
            d = to_num(ln.get("借方"))
            c = to_num(ln.get("贷方"))
            d = d or 0.0
            c = c or 0.0
            if subj == "":
                errors.append("第 %d 张凭证存在无科目的分录行" % vno)
                norm = None
                break
            if d != 0 and c != 0:
                errors.append("第 %d 张凭证科目「%s」借贷同时有值" % (vno, subj))
                norm = None
                break
            td += d
            tc += c
            norm.append({"科目": subj, "借方": round(d, 2), "贷方": round(c, 2)})
        if norm is None:
            continue
        if abs(td - tc) >= 0.005:
            errors.append("第 %d 张凭证借贷不平（借 %.2f / 贷 %.2f）— 未登记" % (vno, td, tc))
            continue

        dt = str(e.get("date") or "").strip()
        v = str(e.get("voucher") or ("记-%03d" % vno)).strip()
        summ = str(e.get("summary") or "").strip()
        vouchers.append({"序号": vno, "日期": dt, "凭证号": v, "摘要": summ,
                         "借方合计": round(td, 2), "贷方合计": round(tc, 2)})
        for ln in norm:
            journal.append({"日期": dt, "凭证号": v, "摘要": summ,
                            "科目": ln["科目"], "借方": ln["借方"], "贷方": ln["贷方"]})

    return {
        "vouchers": vouchers,
        "journal": journal,
        "errors": errors,
        "voucherCount": len(vouchers),
        "journalLineCount": len(journal),
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
        if action in ("bonus", "年终奖", "全年一次性奖金"):
            data, err = do_bonus(params)
            if err:
                print(json.dumps({"ok": False, "message": err}, ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "年终奖计税.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            write_result(out, "年终奖个人所得税（单独计税）",
                         ["姓名", "奖金", "月均", "税率", "速算扣除数", "税额", "税后"],
                         [[r["姓名"], r["奖金"], r["月均"], r["税率"], r["速算扣除数"],
                           r["税额"], r["税后"]] for r in data["rows"]])
            msg = "%d 人年终奖合计 %.2f，个税合计 %.2f，税后合计 %.2f" % (
                data["count"], data["totalBonus"], data["totalTax"], data["totalNet"])
            if data.get("compare"):
                c = data["compare"]
                saving = c["savingBySeparate"]
                if saving > 0:
                    msg += "；单独计税比并入综合所得**省 %.2f 元**" % saving
                elif saving < 0:
                    msg += "；并入综合所得更省 %.2f 元" % (-saving)
                else:
                    msg += "；两种方式税负相同"
            msg += "；已输出 " + out
            print(json.dumps({"ok": True, "message": msg, "data": dict(data, out=out)},
                             ensure_ascii=False))
            return 0

        if action in ("payroll", "工资", "个税"):
            if not inputs:
                print(json.dumps({"ok": False, "message": "payroll 需要输入文件（含姓名/应发工资列的 xlsx/csv）"},
                                 ensure_ascii=False))
                return 1
            header, rows = read_table(inputs[0], params.get("sheet"), params.get("headerRow") or 1)
            data, err = do_payroll(header, rows, params)
            if err:
                print(json.dumps({"ok": False, "message": err, "data": {"header": header}},
                                 ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "工资表.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            write_result(out, "工资表（%d 月，累计预扣法）" % data["month"],
                         ["姓名", "应发工资", "社保", "公积金", "专项附加扣除",
                          "累计应纳税所得额", "税率", "速算扣除数", "个税", "实发工资"],
                         [[r["姓名"], r["应发工资"], r["社保"], r["公积金"], r["专项附加扣除"],
                           r["累计应纳税所得额"], r["税率"], r["速算扣除数"], r["个税"], r["实发工资"]]
                          for r in data["rows"]])
            print(json.dumps({
                "ok": True,
                "message": "%d 人工资表：应发合计 %.2f，个税合计 %.2f，实发合计 %.2f；已输出 %s" % (
                    data["employeeCount"], data["totalGross"], data["totalTax"], data["totalNet"], out),
                "data": dict(data, out=out),
            }, ensure_ascii=False))
            return 0

        if action in ("bank-recon", "bankrecon", "银行余额调节表", "调节表"):
            data, err = do_bank_recon(None, None, params)
            if err:
                print(json.dumps({"ok": False, "message": err}, ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "银行余额调节表.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            write_result(out, "银行存款余额调节表", ["项目", "金额"], data["rows"])
            verdict = "调节后双方相等 ✓" if data["balanced"] else \
                "调节后仍差 %.2f ✗（请核对未达账项）" % data["difference"]
            print(json.dumps({
                "ok": True,
                "message": "调节后银行 %.2f / 企业 %.2f；%s；已输出 %s" % (
                    data["adjustedBank"], data["adjustedBook"], verdict, out),
                "data": dict(data, out=out),
            }, ensure_ascii=False))
            return 0

        if action in ("consolidation", "合并", "抵消"):
            data, err = do_consolidation(params)
            if err:
                print(json.dumps({"ok": False, "message": err}, ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "合并报表.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            # 三栏对照：个别数 / 抵消 / 合并数
            rows = []
            for k in ("资产", "负债", "权益", "收入", "费用", "利润"):
                indiv = data["individual"].get(k, 0.0) if k != "利润" else \
                    round(data["individual"]["收入"] - data["individual"]["费用"], 2)
                merged = data["merged"].get(k, 0.0)
                rows.append([k, indiv, r2(merged - indiv), merged])
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "合并报表"
            ws.append(["合并报表工作底稿"]); ws["A1"].font = openpyxl.styles.Font(bold=True, size=13)
            ws.append([])
            ws.append(["项目", "个别数合计", "抵消", "合并数"])
            for c in range(1, 5):
                ws.cell(row=3, column=c).font = openpyxl.styles.Font(bold=True)
            for r in rows:
                ws.append(r)
            if data["eliminations"]:
                ws.append([])
                ws.append(["抵消分录", "借方科目", "贷方科目", "金额"])
                for c in range(1, 5):
                    ws.cell(row=ws.max_row, column=c).font = openpyxl.styles.Font(bold=True)
                for e in data["eliminations"]:
                    ws.append(e)
            ws.column_dimensions["A"].width = 24
            for col in ("B", "C", "D"):
                ws.column_dimensions[col].width = 16
            if os.path.dirname(out) and not os.path.isdir(os.path.dirname(out)):
                os.makedirs(os.path.dirname(out))
            wb.save(out)
            verdict = "合并后资产 = 负债+权益+利润 ✓" if data["balanced"] else \
                "合并后不平 ✗（差 %.2f）" % data["balanceDiff"]
            extra = ""
            if data["unmatchedEliminations"]:
                extra = "；%d 条抵消分录科目无法归类，已跳过（见 unmatchedEliminations）" % \
                    len(data["unmatchedEliminations"])
            print(json.dumps({
                "ok": True,
                "message": "%d 家公司合并，抵消 %d 笔；%s%s；已输出 %s" % (
                    len(data["companies"]), len(data["eliminations"]), verdict, extra, out),
                "data": dict(data, out=out),
            }, ensure_ascii=False))
            return 0

        if action in ("journal", "凭证", "日记账"):
            data, err = do_journal(params)
            if err:
                print(json.dumps({"ok": False, "message": err}, ensure_ascii=False))
                return 1
            if not out:
                out = os.path.join(base_dir, "记账凭证.xlsx")
            elif not os.path.isabs(out):
                out = os.path.join(base_dir, out)
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "凭证汇总"
            ws1.append(["记账凭证汇总"]); ws1["A1"].font = openpyxl.styles.Font(bold=True, size=13)
            ws1.append([])
            ws1.append(["序号", "日期", "凭证号", "摘要", "借方合计", "贷方合计"])
            for c in range(1, 7):
                ws1.cell(row=3, column=c).font = openpyxl.styles.Font(bold=True)
            for v in data["vouchers"]:
                ws1.append([v["序号"], v["日期"], v["凭证号"], v["摘要"], v["借方合计"], v["贷方合计"]])
            ws2 = wb.create_sheet("日记账")
            ws2.append(["日期", "凭证号", "摘要", "科目", "借方", "贷方"])
            for c in range(1, 7):
                ws2.cell(row=1, column=c).font = openpyxl.styles.Font(bold=True)
            for j in data["journal"]:
                ws2.append([j["日期"], j["凭证号"], j["摘要"], j["科目"], j["借方"], j["贷方"]])
            for ws in (ws1, ws2):
                ws.column_dimensions["A"].width = 12
                ws.column_dimensions["C"].width = 24
                ws.column_dimensions["D"].width = 18
            if os.path.dirname(out) and not os.path.isdir(os.path.dirname(out)):
                os.makedirs(os.path.dirname(out))
            wb.save(out)
            msg = "生成 %d 张凭证、%d 行日记账" % (data["voucherCount"], data["journalLineCount"])
            if data["errors"]:
                msg += "；%d 张被拒（借贷不平或科目缺失）：%s" % (len(data["errors"]), data["errors"][0])
            msg += "；已输出 " + out
            print(json.dumps({"ok": True, "message": msg, "data": dict(data, out=out)},
                             ensure_ascii=False))
            return 0

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
                # 没有直接给 items：从输入表构造。
                # 支持两种真实表形态（回归修复）：
                #   A) 单一「余额」列 + 可选「方向」列
                #   B) 借贷分列（期末借方/期末贷方，或通用 借方/贷方）
                # 旧实现只认 A，且在没有「方向」列时**默认全部按借方**——真实余额表多为
                # 借贷分列，模型只好自己拼一张"标准表"，结果负债/权益科目全被当成借方，
                # 资产负债表出现负的负债合计（实测：资产 650000 vs 负债+权益 -650000，
                # 差 1300000，模型据此报告"报表不平衡"）。
                header, rows = read_table(inputs[0], params.get("sheet"), params.get("headerRow") or 1)
                c_subj = find_col(header, ["科目名称", "会计科目", "科目", "subject", "account"])
                # 科目代码列（可选）：用于识别损益类科目（6 开头），决定取期末余额还是本期发生额。
                # 「科目代码」必须排在「科目」之前试——否则包含匹配会命中的是"科目代码"之外的列。
                c_code = find_col(header, ["科目代码", "科目编码", "代码", "code"])
                c_amt = find_col(header, ["余额", "金额", "amount", "balance"])
                c_dir = find_col(header, ["方向", "借贷", "dir"])
                # 借贷分列：优先期末，退化到通用借贷列
                c_dr = find_col(header, ["期末借方", "借方余额", "借方", "debit"])
                c_cr = find_col(header, ["期末贷方", "贷方余额", "贷方", "credit"])
                # 本期发生额列：损益类科目期末余额为 0（已结转），利润表必须取本期发生额，
                # 否则利润表全是 0（实测：模型据此报告"利润表计算准确"但数字全空）。
                c_cur_dr = find_col(header, ["本期借方", "本期发生额借方", "发生额借方"])
                c_cur_cr = find_col(header, ["本期贷方", "本期发生额贷方", "发生额贷方"])
                split_mode = (c_amt < 0) and (c_dr >= 0 or c_cr >= 0)
                if c_subj < 0 or (c_amt < 0 and not split_mode):
                    print(json.dumps({"ok": False, "message": "需要「科目」加「余额」列，或「科目」加借贷分列（借方/贷方）",
                                      "data": {"header": header}}, ensure_ascii=False))
                    return 1
                items = []
                for r in rows:
                    subj = str(r[c_subj]).strip() if c_subj < len(r) and r[c_subj] is not None else ""
                    # 跳过表头重复行与常见合计行（否则会被当科目计入，虚增金额）
                    if not subj or subj in ("科目", "科目名称", "合计", "总计", "小计"):
                        continue
                    # 科目代码列（若存在）用于判断损益类；没有代码时按名称判断
                    code = ""
                    if c_code >= 0 and c_code < len(r) and r[c_code] is not None:
                        code = str(r[c_code]).strip()
                    is_pl = code.startswith("6") or any(
                        k in subj for k in ("收入", "成本", "费用", "税金及附加", "所得税"))
                    if split_mode:
                        if is_pl and (c_cur_dr >= 0 or c_cur_cr >= 0):
                            # 损益类取本期发生额
                            dv = to_num(r[c_cur_dr]) if (c_cur_dr >= 0 and c_cur_dr < len(r)) else None
                            cv = to_num(r[c_cur_cr]) if (c_cur_cr >= 0 and c_cur_cr < len(r)) else None
                        else:
                            dv = to_num(r[c_dr]) if (c_dr >= 0 and c_dr < len(r)) else None
                            cv = to_num(r[c_cr]) if (c_cr >= 0 and c_cr < len(r)) else None
                        dv = dv or 0.0
                        cv = cv or 0.0
                        if dv == 0.0 and cv == 0.0:
                            continue
                        # 净额口径：借正贷负
                        amt = dv - cv
                        if amt == 0.0:
                            continue
                        items.append({"科目": subj, "金额": abs(amt), "方向": "借" if amt > 0 else "贷"})
                    else:
                        amt = to_num(r[c_amt]) if c_amt < len(r) else None
                        if amt is None:
                            continue
                        d = ""
                        if c_dir >= 0 and c_dir < len(r) and r[c_dir]:
                            d = str(r[c_dir]).strip()
                        if not d:
                            # 没有方向列：按金额符号推断（负数为贷方），而不是一律当借方
                            d = "借" if amt >= 0 else "贷"
                            amt = abs(amt)
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
