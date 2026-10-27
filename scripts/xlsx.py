"""极简 XLSX 写出器：只用标准库，不依赖 openpyxl。

为什么不用 openpyxl
------------------
报表要能在**任何** python3 下直接跑（定时任务、别人的机器、临时 venv）。openpyxl
是第三方包，一旦 import 失败整份报表就出不来，而这个文件的全部用途不过是「把六张
表放进一个工作簿」。所以这里手写 OOXML：

    [Content_Types].xml
    _rels/.rels
    xl/workbook.xml
    xl/_rels/workbook.xml.rels
    xl/styles.xml
    xl/worksheets/sheetN.xml

支持的能力刚好够用：多 sheet、表头加粗+底色、**全部单元格水平垂直居中**、
首行冻结、按内容估算列宽、数字按数值写（能被 Excel 排序/求和）、字符串走
inlineStr（省掉 sharedStrings 那一套）。
"""

from __future__ import annotations

import re
import zipfile

# 表头样式 id / 正文样式 id，与 styles.xml 里的 cellXfs 顺序对应
ST_HEAD = 2
ST_BODY = 1

_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# 自己转义，不引 xml.sax.saxutils：后者会把 urllib.request / http.client /
# tempfile 整条链拉进来，光是导入就要 0.25s，而这里只需要替换四个字符。
# str.translate 一次扫描完成，也不会出现「转义后再把 & 转一遍」的二次转义。
_XML_ESC = str.maketrans({"&": "&amp;", "<": "&lt;", ">": "&gt;"})
_XML_ESC_ATTR = str.maketrans({"&": "&amp;", "<": "&lt;", ">": "&gt;",
                              '"': "&quot;"})


def escape(text):
    return text.translate(_XML_ESC)


def escape_attr(text):
    return text.translate(_XML_ESC_ATTR)


def _clean(text):
    return _ILLEGAL.sub("", str(text))


def _col_name(index):
    """1 -> A，27 -> AA。"""
    name = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def _disp_width(value):
    """中文按 2 列算，否则算出来的列宽会把中文挤成 ###。"""
    text = _clean(value) if value is not None else ""
    width = 0
    for ch in text:
        width += 2 if ord(ch) > 127 else 1
    return width


def _cell_xml(ref, value, style):
    if value is None or value == "":
        return '<c r="%s" s="%d"/>' % (ref, style)
    if isinstance(value, bool):
        value = "是" if value else "否"
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value != value:   # NaN
            return '<c r="%s" s="%d"/>' % (ref, style)
        return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, repr(value) if isinstance(value, float) else value)
    text = escape(_clean(value))
    return ('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, style, text))


def _sheet_xml(headers, rows):
    ncol = max([len(headers)] + [len(r) for r in rows]) if rows else len(headers)
    # 列宽：表头与前 200 行的最大显示宽度，夹在 8~42 之间
    widths = []
    for i in range(ncol):
        head = headers[i] if i < len(headers) else ""
        widest = _disp_width(head)
        for row in rows[:200]:
            if i < len(row):
                widest = max(widest, _disp_width(row[i]))
        widths.append(max(8, min(42, widest + 2)))

    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>',
        '<sheetFormatPr defaultRowHeight="16"/>',
        "<cols>",
    ]
    for i, width in enumerate(widths, start=1):
        out.append('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i, i, width))
    out.append("</cols><sheetData>")

    if headers:
        cells = "".join(_cell_xml("%s1" % _col_name(i), h, ST_HEAD)
                        for i, h in enumerate(headers, start=1))
        out.append('<row r="1" ht="22" customHeight="1">%s</row>' % cells)

    for r, row in enumerate(rows, start=2):
        cells = "".join(_cell_xml("%s%d" % (_col_name(i), r), v, ST_BODY)
                        for i, v in enumerate(row, start=1))
        out.append('<row r="%d">%s</row>' % (r, cells))

    out.append("</sheetData></worksheet>")
    return "".join(out)


_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="10"/><name val="Microsoft YaHei"/></font>
<font><b/><sz val="10"/><name val="Microsoft YaHei"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
<alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1"
 applyFill="1" applyAlignment="1">
<alignment horizontal="center" vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def write_workbook(path, sheets):
    """写出一个多 sheet 工作簿。

    ``sheets`` = [(sheet 名, 表头列表, 行列表), ...]，顺序即工作簿里的顺序。
    所有单元格（含表头）水平+垂直居中；表头加粗+浅蓝底+冻结首行。
    """
    if not sheets:
        raise ValueError("至少需要一张表")

    names = []
    for name, _, _ in sheets:
        # Excel 硬限制：≤31 字符，且不能含 : \ / ? * [ ]
        safe = re.sub(r"[:\\/?*\[\]]", "_", name)[:31] or "Sheet"
        base, suffix = safe, 2
        while safe in names:          # 重名会被 Excel 判为损坏
            tail = "_%d" % suffix
            safe = base[:31 - len(tail)] + tail
            suffix += 1
        names.append(safe)

    rel_styles = "rId%d" % (len(sheets) + 1)
    content = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for i in range(1, len(sheets) + 1):
        content.append(
            '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType='
            '"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
    content.append("</Types>")

    workbook = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheets>",
    ]
    for i, name in enumerate(names, start=1):
        workbook.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                        % (escape_attr(name), i, i))
    workbook.append("</sheets></workbook>")

    wb_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for i in range(1, len(sheets) + 1):
        wb_rels.append(
            '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>'
            % (i, i))
    wb_rels.append(
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/styles" Target="styles.xml"/>' % rel_styles)
    wb_rels.append("</Relationships>")

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "".join(content))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", "".join(workbook))
        zf.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
        zf.writestr("xl/styles.xml", _STYLES)
        for i, (_, headers, rows) in enumerate(sheets, start=1):
            zf.writestr("xl/worksheets/sheet%d.xml" % i, _sheet_xml(headers, rows))
    return path
