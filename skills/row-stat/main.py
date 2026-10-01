# -*- coding: utf-8 -*-
"""row-stat —— python38 内置示例技能：CSV 行列与数值列统计（纯 stdlib）

协议（SkillRunner 约定）：
  输入：暂存目录 request.json = {"task":"row-stat","inputs":["<csv 绝对路径>",...]}
  输出：stdout 最后一行 JSON = {"ok":true/false,"message":"...","data":{...}}
  数据只经文件与 stdout，不过命令行；编码探测 UTF-8(BOM)/GB18030（会计 CSV 现实）。
"""
import json
import os
import sys


def detect_read(path):
    data = open(path, "rb").read()
    if data[:3] == b"\xef\xbb\xbf":
        return data[3:].decode("utf-8"), "utf-8-sig"
    try:
        return data.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        return data.decode("gb18030"), "gb18030"


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
            elif c == '\r':
                if i + 1 < n and text[i + 1] == '\n':
                    i += 1
                row.append("".join(cur))
                rows.append(row)
                row, cur = [], []
            elif c == '\n':
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


def is_number(s):
    s = s.strip().replace(",", "")
    if not s:
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def stat_file(path):
    text, enc = detect_read(path)
    rows = parse_csv(text)
    if not rows:
        return {"file": os.path.basename(path), "rows": 0, "cols": 0, "encoding": enc}
    header = rows[0]
    cols = len(header)
    numeric = []
    for c in range(cols):
        vals = [r[c] for r in rows[1:] if c < len(r) and is_number(r[c])]
        if len(rows) > 1 and len(vals) >= max(1, len(rows) - 1):
            nums = [float(v.replace(",", "")) for v in vals]
            numeric.append({"col": header[c] if c < len(header) else str(c),
                            "sum": round(sum(nums), 4),
                            "min": round(min(nums), 4),
                            "max": round(max(nums), 4)})
    return {"file": os.path.basename(path), "rows": len(rows) - 1, "cols": cols,
            "encoding": enc, "numeric": numeric}


def main():
    stage = os.path.dirname(os.path.abspath(__file__))
    req_path = os.path.join(stage, "request.json")
    try:
        req = json.load(open(req_path, encoding="utf-8"))
    except Exception as e:
        print(json.dumps({"ok": False, "message": "request.json 读取失败: %s" % e}, ensure_ascii=False))
        return 1
    inputs = [p for p in req.get("inputs", []) if os.path.isfile(p)]
    if not inputs:
        print(json.dumps({"ok": False, "message": "没有可统计的输入文件"}, ensure_ascii=False))
        return 1
    files = [stat_file(p) for p in inputs]
    total_rows = sum(f["rows"] for f in files)
    print(json.dumps({"ok": True,
                      "message": "统计完成：%d 个文件，共 %d 行" % (len(files), total_rows),
                      "data": {"files": files, "totalRows": total_rows}},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
