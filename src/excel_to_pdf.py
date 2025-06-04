# coding: utf-8
"""Simple Excel to PDF converter using only the Python standard library.

This script reads data from the sheet ``ENTREGAS`` inside ``PROJECT MANAGEMENT.xlsx``
located in the same directory. The headers are defined in row 3 and each
subsequent row forms a separate PDF file containing the values for all columns.

Because third-party libraries are unavailable in this environment, the script
implements minimal XLSX parsing and a very small PDF generator from scratch.
"""

import sys
import os
import xml.etree.ElementTree as ET
import zipfile


# ---------------------------------------------------------------------------
# Utilities for XLSX parsing
# ---------------------------------------------------------------------------

def _load_shared_strings(z):
    """Return list of shared strings from an open ZipFile ``z``."""
    try:
        xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings = []
    for si in xml.findall("m:si", ns):
        t = si.find(".//m:t", ns)
        strings.append(t.text if t is not None else "")
    return strings


def _column_index(cell_ref):
    """Convert an Excel column letter reference (e.g. 'A', 'AB') to zero-based index."""
    col = ''.join(filter(str.isalpha, cell_ref))
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx - 1


def _read_sheet(z, sheet_path, shared_strings):
    """Return list of rows, each row is list of cell values."""
    xml = ET.fromstring(z.read(sheet_path))
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows = []
    for row in xml.findall(".//m:sheetData/m:row", ns):
        row_values = {}
        for c in row.findall("m:c", ns):
            ref = c.get("r")
            idx = _column_index(ref)
            v = c.find("m:v", ns)
            if v is None:
                val = ""
            elif c.get("t") == "s":
                val = shared_strings[int(v.text)]
            else:
                val = v.text
            row_values[idx] = val
        # convert to list while filling missing cells
        if row_values:
            max_idx = max(row_values.keys())
            row_list = [row_values.get(i, "") for i in range(max_idx + 1)]
            rows.append(row_list)
    return rows


def read_entregas_sheet(xlsx_path):
    """Read sheet 'ENTREGAS' from ``xlsx_path``. Return (headers, data_rows)."""
    with zipfile.ZipFile(xlsx_path) as z:
        shared_strings = _load_shared_strings(z)
        # find sheet path for ENTREGAS
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rid = None
        for sheet in wb.findall(".//m:sheets/m:sheet", ns):
            if sheet.get("name") == "ENTREGAS":
                rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                break
        if not rid:
            raise ValueError("Sheet 'ENTREGAS' not found")
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        nsr = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
        target = None
        for rel in rels.findall("r:Relationship", nsr):
            if rel.get("Id") == rid:
                target = rel.get("Target")
                break
        if not target:
            raise ValueError("Relationship for sheet 'ENTREGAS' not found")
        sheet_path = os.path.join("xl", target)
        rows = _read_sheet(z, sheet_path, shared_strings)
        if len(rows) < 4:
            return [], []
        headers = rows[2]  # header row is the third row (index 2)
        data_rows = rows[3:]  # rows with information start at row 4
        return headers, data_rows


# ---------------------------------------------------------------------------
# Minimal PDF generation
# ---------------------------------------------------------------------------

def _escape(text):
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)").replace("\r", "").replace("\n", "\\n")


def write_pdf(path, lines):
    """Create a very small PDF file at ``path`` containing ``lines`` of text."""
    # build page content
    y = 750
    content_lines = ["BT", "/F1 12 Tf"]
    for line in lines:
        content_lines.append(f"72 {y} Td ({_escape(line)}) Tj")
        y -= 14
    content_lines.append("ET")
    content_stream = "\n".join(content_lines) + "\n"

    objects = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        f"<< /Length {len(content_stream.encode('utf-8'))} >>\nstream\n{content_stream}endstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]

    offsets = []
    offset = len("%PDF-1.4\n")
    for obj in objects:
        offsets.append(offset)
        offset += len(f"{objects.index(obj)+1} 0 obj\n{obj}\nendobj\n".encode("utf-8"))
    xref_offset = offset

    with open(path, "wb") as f:
        f.write(b"%PDF-1.4\n")
        for i, obj in enumerate(objects, start=1):
            f.write(f"{i} 0 obj\n".encode("utf-8"))
            f.write(obj.encode("utf-8"))
            f.write(b"\nendobj\n")
        f.write(f"xref\n0 {len(objects)+1}\n".encode("utf-8"))
        f.write(b"0000000000 65535 f \n")
        for off in offsets:
            f.write(f"{off:010d} 00000 n \n".encode("utf-8"))
        f.write(b"trailer\n")
        f.write(f"<< /Size {len(objects)+1} /Root 1 0 R >>\n".encode("utf-8"))
        f.write(b"startxref\n")
        f.write(f"{xref_offset}\n".encode("utf-8"))
        f.write(b"%%EOF")


# ---------------------------------------------------------------------------
# Main functionality
# ---------------------------------------------------------------------------

def main():
    root_dir = os.path.dirname(os.path.dirname(__file__))
    xlsx_path = os.path.join(root_dir, "data", "PROJECT MANAGEMENT.xlsx")
    out_dir = os.path.join(root_dir, "pdf_output")
    os.makedirs(out_dir, exist_ok=True)

    headers, rows = read_entregas_sheet(xlsx_path)
    if not headers:
        print("No data found.")
        return

    for idx, row in enumerate(rows, start=1):
        # merge headers and row values into lines
        lines = []
        for h, v in zip(headers, row):
            lines.append(f"{h}: {v}")
        pdf_name = os.path.join(out_dir, f"row_{idx}.pdf")
        write_pdf(pdf_name, lines)
        print(f"Wrote {pdf_name}")


if __name__ == "__main__":
    main()
