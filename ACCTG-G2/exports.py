"""Tiny dependency-free writers for .xlsx and .pdf files (standard library only)."""
import zipfile
from datetime import datetime
from decimal import Decimal
from xml.sax.saxutils import escape


def _col_letter(index):
    """0 -> A, 25 -> Z, 26 -> AA ..."""
    index += 1
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def write_xlsx(path, sheet_name, headers, rows):
    """Write a one-sheet Excel workbook.

    rows is a list of (values, bold); each value is a string, a number
    (int / float / Decimal) or None for an empty cell.
    """
    def cell(ref, value, bold=False):
        if value is None or value == "":
            return ""
        if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            text = format(value, "f") if isinstance(value, Decimal) else repr(value)
            return f'<c r="{ref}" s="{3 if bold else 2}"><v>{text}</v></c>'
        return (f'<c r="{ref}" t="inlineStr" s="{1 if bold else 0}">'
                f'<is><t xml:space="preserve">{escape(str(value))}</t></is></c>')

    xml_rows = []
    header_cells = "".join(cell(f"{_col_letter(c)}1", h, True) for c, h in enumerate(headers))
    xml_rows.append(f'<row r="1">{header_cells}</row>')
    for r, (values, bold) in enumerate(rows, start=2):
        cells = "".join(cell(f"{_col_letter(c)}{r}", v, bold) for c, v in enumerate(values))
        xml_rows.append(f'<row r="{r}">{cells}</row>')

    ncols = len(headers)
    cols = ('<cols><col min="1" max="1" width="36" customWidth="1"/>'
            f'<col min="2" max="{max(ncols, 2)}" width="13" customWidth="1"/></cols>')
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>'
        '</sheetView></sheetViews>'
        f'{cols}<sheetData>{"".join(xml_rows)}</sheetData></worksheet>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '<xf numFmtId="4" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )
    safe_name = escape(sheet_name[:31], {'"': "&quot;"})
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{safe_name}" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/styles" Target="styles.xml"/></Relationships>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.styles+xml"/></Types>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", styles)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


def _pdf_text(text):
    """Escape a string for a PDF literal (Latin-1 only; other characters become '?')."""
    data = str(text).encode("latin-1", "replace").decode("latin-1")
    return data.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def write_pdf(path, title, headers, rows):
    """Write a landscape A4 PDF table in a monospaced font.

    rows is a list of (values, bold); every value must already be a string.
    The font shrinks automatically so wide tables (like the 12-column
    worksheet) fit on the page width.
    """
    page_w, page_h, margin = 842, 595, 30
    gap = 2
    table = [([str(v) for v in values], bold) for values, bold in rows]
    ncols = len(headers)
    widths = [len(str(h)) for h in headers]
    for values, _ in table:
        for j, v in enumerate(values[:ncols]):
            widths[j] = max(widths[j], len(v))
    total_chars = sum(widths) + gap * (ncols - 1)
    size = max(4.0, min(9.0, (page_w - 2 * margin) / (0.6 * total_chars)))
    line_h = size * 1.5

    def build_line(values):
        parts = []
        for j in range(ncols):
            v = values[j] if j < len(values) else ""
            parts.append(v.ljust(widths[j]) if j == 0 else v.rjust(widths[j]))
        return (" " * gap).join(parts)

    header_y = page_h - 55
    first_row_y = header_y - line_h * 1.6
    per_page = max(1, int((first_row_y - 35) // line_h) + 1)
    chunks = [table[i:i + per_page] for i in range(0, len(table), per_page)] or [[]]
    stamp = datetime.now().strftime("%b %d, %Y")

    objects = [None] * (4 + 2 * len(chunks) + 1)  # 1-based ids; slot 0 unused
    objects[1] = "<< /Type /Catalog /Pages 2 0 R >>"
    kids = " ".join(f"{5 + 2 * i} 0 R" for i in range(len(chunks)))
    objects[2] = f"<< /Type /Pages /Kids [{kids}] /Count {len(chunks)} >>"
    objects[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Courier /Encoding /WinAnsiEncoding >>"
    objects[4] = "<< /Type /Font /Subtype /Type1 /BaseFont /Courier-Bold /Encoding /WinAnsiEncoding >>"

    for i, chunk in enumerate(chunks):
        ops = [
            f"BT /F2 12 Tf {margin} {page_h - 30} Td ({_pdf_text(title)}) Tj ET",
            f"BT /F1 8 Tf {page_w - margin - 120} {page_h - 30} Td ({_pdf_text(stamp)}) Tj ET",
            f"BT /F2 {size:.2f} Tf {margin} {header_y} Td ({_pdf_text(build_line(headers))}) Tj ET",
        ]
        y = first_row_y
        for values, bold in chunk:
            font = "F2" if bold else "F1"
            ops.append(f"BT /{font} {size:.2f} Tf {margin} {y:.2f} Td ({_pdf_text(build_line(values))}) Tj ET")
            y -= line_h
        ops.append(f"BT /F1 8 Tf {page_w / 2 - 30} 18 Td (Page {i + 1} of {len(chunks)}) Tj ET")
        stream = "\n".join(ops)
        page_id, content_id = 5 + 2 * i, 6 + 2 * i
        objects[page_id] = (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_w} {page_h}] "
                            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {content_id} 0 R >>")
        objects[content_id] = f"<< /Length {len(stream.encode('latin-1', 'replace'))} >>\nstream\n{stream}\nendstream"

    out = bytearray(b"%PDF-1.4\n")
    offsets = {}
    for obj_id in range(1, len(objects)):
        offsets[obj_id] = len(out)
        out += f"{obj_id} 0 obj\n".encode("latin-1")
        out += objects[obj_id].encode("latin-1", "replace")
        out += b"\nendobj\n"
    xref_pos = len(out)
    out += f"xref\n0 {len(objects)}\n".encode("latin-1")
    out += b"0000000000 65535 f \n"
    for obj_id in range(1, len(objects)):
        out += f"{offsets[obj_id]:010d} 00000 n \n".encode("latin-1")
    out += (f"trailer\n<< /Size {len(objects)} /Root 1 0 R >>\n"
            f"startxref\n{xref_pos}\n%%EOF\n").encode("latin-1")
    with open(path, "wb") as f:
        f.write(out)