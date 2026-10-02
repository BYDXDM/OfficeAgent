# -*- coding: utf-8 -*-
"""docx-report —— Word 文档生成（python38 技能）

协议（SkillRunner 约定）：
  输入：暂存目录 request.json = {"task":"docx-report","inputs":[...],"...":"..."}
        —— 参数直接在顶层（可选再套一层 params，两种都支持）
  输出：stdout 最后一行 JSON = {"ok":true/false,"message":"...","data":{...}}
  数据只经文件与 stdout，不过命令行。

用途：把结构化内容写成 .docx（报表/说明/通知/合同草稿）。
宿主原先只能**读** Word，不能生成 —— 本技能补上这一环。

参数：
  out      输出 .docx 路径（绝对路径；相对路径落到第一个输入文件所在目录，
           无输入文件时落到当前目录）
  title    文档标题（可选，会作为 Heading 0 大标题）
  blocks   内容块数组，每项是对象，支持：
             {"type":"heading", "text":"...", "level":1}
             {"type":"para",    "text":"..."}
             {"type":"bullet",  "items":["...","..."]}
             {"type":"number",  "items":["...","..."]}
             {"type":"table",   "header":["列1","列2"], "rows":[["a","b"],...]}
             {"type":"pagebreak"}
             {"type":"kv",      "pairs":[["项目","值"],...]}   # 两列无边框表，适合抬头
  csv      可选：若给 csv（CSV 文本），作为一张表插入（等价 blocks 里一个 table）

兼容便捷写法：
  text     纯文本正文（按空行分段），当 blocks 未给时使用
"""
import json
import os
import sys

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.table import WD_TABLE_ALIGNMENT
except Exception as e:  # pragma: no cover
    print(json.dumps({"ok": False, "message": "缺少 python-docx: %s" % e}, ensure_ascii=False))
    sys.exit(0)


def stage_dir():
    return os.path.dirname(os.path.abspath(__file__))


def load_request():
    with open(os.path.join(stage_dir(), "request.json"), encoding="utf-8") as f:
        return json.load(f)


def parse_csv(text):
    """极简 CSV 解析（支持引号包裹与双引号转义）。"""
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


def set_cjk_font(run, name="宋体", size=None, bold=None):
    """设置中文字体：需同时设 ascii/eastasia，否则 Word 里中文会回退成默认字体。"""
    run.font.name = name
    try:
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
        if rfonts is None:
            from docx.oxml.ns import qn
            rfonts = rpr.makeelement(qn('w:rFonts'), {})
            rpr.append(rfonts)
        from docx.oxml.ns import qn
        rfonts.set(qn('w:eastAsia'), name)
    except Exception:
        pass
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def add_table(doc, header, rows):
    ncols = max([len(header or [])] + [len(r) for r in (rows or [])] + [1])
    nrows = (1 if header else 0) + len(rows or [])
    t = doc.add_table(rows=nrows, cols=ncols)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    ri = 0
    if header:
        for ci in range(ncols):
            cell = t.cell(0, ci)
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(header[ci]) if ci < len(header) else "")
            set_cjk_font(run, "宋体", 10.5, True)
        ri = 1
    for r in (rows or []):
        for ci in range(ncols):
            cell = t.cell(ri, ci)
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(r[ci]) if ci < len(r) else "")
            set_cjk_font(run, "宋体", 10.5, False)
        ri += 1
    return t


def build(doc, req):
    title = (req.get("title") or "").strip()
    if title:
        h = doc.add_heading(title, level=0)
        for run in h.runs:
            set_cjk_font(run, "微软雅黑", 20, True)
            run.font.color.rgb = RGBColor(0x1C, 0x1E, 0x26)

    blocks = req.get("blocks")
    if not blocks:
        # 便捷写法：纯文本正文
        text = req.get("text") or ""
        blocks = []
        for para in [p.strip() for p in text.split("\n\n") if p.strip()]:
            blocks.append({"type": "para", "text": para})
        if req.get("csv"):
            blocks.append({"type": "table", "header": None, "rows": None, "_csv": req["csv"]})

    n_table = 0
    for b in blocks:
        typ = (b.get("type") or "para").strip().lower()

        if typ == "heading":
            lvl = int(b.get("level") or 1)
            lvl = max(1, min(lvl, 4))
            h = doc.add_heading(str(b.get("text") or ""), level=lvl)
            for run in h.runs:
                set_cjk_font(run, "微软雅黑", 16 - lvl, True)

        elif typ == "para":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(str(b.get("text") or ""))
            set_cjk_font(run, "宋体", 11)

        elif typ == "bullet":
            for it in (b.get("items") or []):
                p = doc.add_paragraph(style='List Bullet')
                run = p.add_run(str(it))
                set_cjk_font(run, "宋体", 11)

        elif typ == "number":
            for it in (b.get("items") or []):
                p = doc.add_paragraph(style='List Number')
                run = p.add_run(str(it))
                set_cjk_font(run, "宋体", 11)

        elif typ == "kv":
            rows = [[str(k), str(v)] for k, v in (b.get("pairs") or [])]
            t = add_table(doc, None, rows)
            # 首列加粗
            for r in t.rows:
                if r.cells:
                    for run in r.cells[0].paragraphs[0].runs:
                        run.font.bold = True
            n_table += 1

        elif typ == "table":
            csv_text = b.get("_csv")
            if csv_text:
                parsed = parse_csv(csv_text)
                if parsed:
                    add_table(doc, parsed[0], parsed[1:])
                    n_table += 1
            else:
                add_table(doc, b.get("header"), b.get("rows"))
                n_table += 1

        elif typ == "pagebreak":
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

        else:
            # 未知类型按普通段落处理，不静默丢弃
            p = doc.add_paragraph()
            run = p.add_run(str(b.get("text") or b))
            set_cjk_font(run, "宋体", 11)

    return n_table


def main():
    try:
        req = load_request()
    except Exception as e:
        print(json.dumps({"ok": False, "message": "request.json 读取失败: %s" % e}, ensure_ascii=False))
        return 1

    params = dict(req)
    if isinstance(req.get("params"), dict):
        params.update(req["params"])

    inputs = [p for p in (req.get("inputs") or []) if os.path.isfile(p)]

    out = (params.get("out") or "").strip()
    if not out:
        base_dir = os.path.dirname(inputs[0]) if inputs else os.getcwd()
        out = os.path.join(base_dir, "报告.docx")
    elif not os.path.isabs(out):
        base_dir = os.path.dirname(inputs[0]) if inputs else os.getcwd()
        out = os.path.join(base_dir, out)
    if not out.lower().endswith(".docx"):
        out += ".docx"

    try:
        doc = Document()
        # 页边距（A4 常规）
        for s in doc.sections:
            s.top_margin = Cm(2.54)
            s.bottom_margin = Cm(2.54)
            s.left_margin = Cm(3.17)
            s.right_margin = Cm(3.17)
        n_table = build(doc, params)
        d = os.path.dirname(out)
        if d and not os.path.isdir(d):
            os.makedirs(d)
        doc.save(out)
    except Exception as e:
        print(json.dumps({"ok": False, "message": "生成失败: %s" % e}, ensure_ascii=False))
        return 1

    n_para = len(doc.paragraphs)
    print(json.dumps({
        "ok": True,
        "message": "已生成 Word 文档（%d 段落、%d 表格）: %s" % (n_para, n_table, out),
        "data": {"out": out, "paragraphs": n_para, "tables": n_table},
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
