# -*- coding: utf-8 -*-
# 构造一个"模拟用户文件"的 xlsx：共享公式（master 带内容 + follower 自闭合 <f/>），
# 且**不写 <v> 缓存值** —— 正是系统/ERP 程序化生成文件的典型形态。
import os, sys, io, zipfile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

out = os.path.join(os.environ["TEMP"], "oa_shared_formula_test.xlsx")

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

WB = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>'''

WB_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''

def s(ref, txt):
    return '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, txt)

# B2 = master（带公式内容、无 <v>）；B3 = follower（自闭合 <f/>、无 <v>）
SHEET = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    '<row r="1">' + s("A1", "名称") + s("B1", "计算结果") + '</row>'
    '<row r="2">' + s("A2", "甲")
        + '<c r="B2"><f t="shared" ref="B2:B3" si="0">IF(A2=&quot;&quot;,&quot;&quot;,A2&amp;&quot;-x&quot;)</f></c>'
        + '</row>'
    '<row r="3">' + s("A3", "乙")
        + '<c r="B3"><f t="shared" si="0"/></c>'
        + '</row>'
    '</sheetData></worksheet>')

with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", SHEET)

print("已生成:", out, os.path.getsize(out), "bytes")
# 打印 sheet XML 供核对
print()
print(SHEET)
