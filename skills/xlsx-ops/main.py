# -*- coding: utf-8 -*-
"""xlsx-ops —— Excel 高级操作（python38 技能）

协议（SkillRunner 约定）：
  输入：暂存目录 request.json = {"task":"xlsx-ops","inputs":["<xlsx 绝对路径>"],"action":"...","...":...}
        —— 参数直接在**顶层**（可选再套一层 params 对象，两种都支持）
  输出：stdout 最后一行 JSON = {"ok":true/false,"message":"...","data":{...}}
  数据只经文件与 stdout，不过命令行。

支持的操作（action，默认 inspect）：
  inspect   —— 列出工作簿结构：工作表名、行列数、表头、公式单元格、合并区域、列宽
  add-sheet —— 用 CSV 文本新建一个工作表（params.name, params.csv）
  set-cell  —— 写入单元格（params.sheet, params.cell 如 B3, params.value）
  formula   —— 写入公式（params.sheet, params.cell, params.formula）
  freeze    —— 冻结窗格（params.sheet, params.cell 如 A2）
  autofit   —— 按内容估算列宽并应用（params.sheet，可选）

输出文件：原地保存到 <输入名>_ops.xlsx（原文件只读，绝不覆盖），
         或用 params.out 指定输出路径（相对路径落到输入文件同目录）。
"""
import json
import os
import sys

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
except Exception as e:  # pragma: no cover
    print(json.dumps({"ok": False, "message": "缺少 openpyxl: %s" % e}, ensure_ascii=False))
    sys.exit(0)


def stage_dir():
    return os.path.dirname(os.path.abspath(__file__))


def load_request():
    p = os.path.join(stage_dir(), "request.json")
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def parse_csv(text):
    """极简 CSV 解析（支持引号包裹与双引号转义），与 row-stat 同族。"""
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


def coerce(v):
    """把字符串尽量转成数字/布尔，便于 Excel 计算。"""
    s = (v or "").strip()
    if s == "":
        return None
    try:
        if "." in s or "e" in s.lower():
            return float(s)
        return int(s)
    except ValueError:
        return s


def do_inspect(wb):
    sheets = []
    for ws in wb.worksheets:
        header = []
        for c in range(1, min(ws.max_column, 20) + 1):
            v = ws.cell(row=1, column=c).value
            header.append("" if v is None else str(v))
        formulas = []
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas.append("%s!%s=%s" % (ws.title, cell.coordinate, cell.value))
        merged = [str(r) for r in ws.merged_cells.ranges]
        widths = {}
        for k, dim in ws.column_dimensions.items():
            if dim.width:
                widths[k] = round(dim.width, 1)
        sheets.append({
            "name": ws.title,
            "rows": ws.max_row,
            "cols": ws.max_column,
            "header": header,
            "formulaCount": len(formulas),
            "formulas": formulas[:20],
            "merged": merged[:20],
            "colWidths": widths,
            "freeze": str(ws.freeze_panes) if ws.freeze_panes else "",
        })
    return {"sheets": sheets}


def main():
    try:
        req = load_request()
    except Exception as e:
        print(json.dumps({"ok": False, "message": "request.json 读取失败: %s" % e}, ensure_ascii=False))
        return 1

    inputs = [p for p in req.get("inputs", []) if os.path.isfile(p)]
    if not inputs:
        print(json.dumps({"ok": False, "message": "没有可处理的输入文件"}, ensure_ascii=False))
        return 1

    # 参数取自 request.json 的**顶层**（与 row-stat 一致；宿主 /skillrun --request 也是平铺拼接，
    # 不套 params 包装）。为兼容两种写法，若存在 params 对象则并入。
    params = dict(req)
    if isinstance(req.get("params"), dict):
        params.update(req["params"])
    action = (params.get("action") or "inspect").strip().lower()
    src = inputs[0]

    try:
        wb = openpyxl.load_workbook(src)
    except Exception as e:
        print(json.dumps({"ok": False, "message": "无法打开工作簿: %s" % e}, ensure_ascii=False))
        return 1

    # ---------- inspect 只读，不写盘 ----------
    if action == "inspect":
        data = do_inspect(wb)
        names = [s["name"] for s in data["sheets"]]
        total_f = sum(s["formulaCount"] for s in data["sheets"])
        print(json.dumps({
            "ok": True,
            "message": "工作簿含 %d 个工作表（%s），公式单元格 %d 个" % (
                len(names), "、".join(names), total_f),
            "data": data,
        }, ensure_ascii=False))
        return 0

    # ---------- 以下动作都会改文件，先确定输出路径 ----------
    base = os.path.splitext(os.path.basename(src))[0]
    out = (params.get("out") or "").strip()
    if not out:
        out = os.path.join(os.path.dirname(src), base + "_ops.xlsx")
    elif not os.path.isabs(out):
        out = os.path.join(os.path.dirname(src), out)

    def sheet_or_default(name):
        if name and name in wb.sheetnames:
            return wb[name]
        return wb.worksheets[0]

    detail = ""

    if action == "add-sheet":
        sname = (params.get("name") or "新表").strip()
        if sname in wb.sheetnames:
            print(json.dumps({"ok": False, "message": "工作表已存在: %s" % sname}, ensure_ascii=False))
            return 1
        ws = wb.create_sheet(sname)
        rows = parse_csv(params.get("csv") or "")
        for row in rows:
            ws.append([coerce(v) for v in row])
        # 表头加粗
        if rows:
            for c in range(1, len(rows[0]) + 1):
                ws.cell(row=1, column=c).font = Font(bold=True)
        detail = "新增工作表「%s」，写入 %d 行" % (sname, len(rows))

    elif action == "set-cell":
        ws = sheet_or_default(params.get("sheet"))
        cell = (params.get("cell") or "").strip()
        if not cell:
            print(json.dumps({"ok": False, "message": "缺少 cell 参数（如 B3）"}, ensure_ascii=False))
            return 1
        ws[cell] = coerce(params.get("value"))
        detail = "已写入 %s!%s = %r" % (ws.title, cell, ws[cell].value)

    elif action == "formula":
        ws = sheet_or_default(params.get("sheet"))
        cell = (params.get("cell") or "").strip()
        f = (params.get("formula") or "").strip()
        if not cell or not f:
            print(json.dumps({"ok": False, "message": "需要 cell 与 formula 参数"}, ensure_ascii=False))
            return 1
        if not f.startswith("="):
            f = "=" + f
        ws[cell] = f
        detail = "已写入公式 %s!%s = %s" % (ws.title, cell, f)

    elif action == "freeze":
        ws = sheet_or_default(params.get("sheet"))
        cell = (params.get("cell") or "A2").strip().upper()
        ws.freeze_panes = cell
        detail = "已冻结 %s 的 %s 以上/以左区域" % (ws.title, cell)

    elif action == "autofit":
        ws = sheet_or_default(params.get("sheet"))
        for col in range(1, ws.max_column + 1):
            best = 8
            for row in range(1, min(ws.max_row, 200) + 1):
                v = ws.cell(row=row, column=col).value
                if v is None:
                    continue
                # 中文按 2 个字符宽度估算
                s = str(v)
                w = sum(2 if ord(ch) > 127 else 1 for ch in s) + 2
                if w > best:
                    best = w
            ws.column_dimensions[get_column_letter(col)].width = min(best, 60)
        detail = "已按内容估算并设置 %d 列列宽" % ws.max_column

    else:
        print(json.dumps({"ok": False, "message": "未知 action: %s" % action}, ensure_ascii=False))
        return 1

    try:
        wb.save(out)
    except Exception as e:
        print(json.dumps({"ok": False, "message": "保存失败（文件可能被 Excel 占用）: %s" % e}, ensure_ascii=False))
        return 1

    print(json.dumps({
        "ok": True,
        "message": "%s；已保存到 %s" % (detail, out),
        "data": {"out": out, "action": action},
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
