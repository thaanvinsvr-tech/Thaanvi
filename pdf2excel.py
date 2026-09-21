"""
PDF -> Excel Converter
------------------------
Extracts every table pdfplumber can find in an uploaded PDF and writes each
one to its own sheet in a downloadable workbook. Pages with no detected
table fall back to a plain-text sheet so nothing is silently dropped.
"""

import io
import re

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
INVALID_SHEET_CHARS = str.maketrans({c: "-" for c in r'\/?*[]:'})


def safe_sheet_name(name):
    return name.translate(INVALID_SHEET_CHARS)[:31]


def extract_pdf_to_sheets(file_bytes, merge_tables_per_page=False):
    """
    Returns a list of (sheet_name, kind, rows) tuples.
    kind is "table" (2D grid of cell values) or "text" (list of single lines).
    """
    sheets = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if tables:
                if merge_tables_per_page and len(tables) > 1:
                    merged = []
                    for t in tables:
                        merged.extend(t)
                        merged.append([])  # blank separator row
                    sheets.append((f"Page{page_num}", "table", merged))
                else:
                    for t_num, table in enumerate(tables, start=1):
                        name = f"Page{page_num}" if len(tables) == 1 else f"Page{page_num}_T{t_num}"
                        sheets.append((name, "table", table))
            else:
                text = page.extract_text() or ""
                lines = [ln for ln in text.split("\n") if ln.strip()]
                if lines:
                    sheets.append((f"Page{page_num}_Text", "text", lines))
    return sheets


def build_workbook(all_file_sheets):
    """
    all_file_sheets: list of (filename, [(sheet_name, kind, rows), ...])
    Returns an in-memory xlsx buffer.
    """
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)
    multi_file = len(all_file_sheets) > 1

    used_names = set()

    def unique_name(base):
        name = safe_sheet_name(base)
        i = 2
        while name in used_names:
            suffix = f"_{i}"
            name = safe_sheet_name(base[:31 - len(suffix)] + suffix)
            i += 1
        used_names.add(name)
        return name

    for filename, sheets in all_file_sheets:
        prefix = re.sub(r"\.pdf$", "", filename, flags=re.I)
        for sheet_name, kind, rows in sheets:
            full_name = f"{prefix}_{sheet_name}" if multi_file else sheet_name
            ws = out_wb.create_sheet(unique_name(full_name))
            if kind == "table":
                for r, row in enumerate(rows, start=1):
                    for c, val in enumerate(row, start=1):
                        ws.cell(row=r, column=c, value=val)
                if rows:
                    for c in range(1, len(rows[0]) + 1):
                        cell = ws.cell(row=1, column=c)
                        cell.font = HEADER_FONT
                        cell.fill = HEADER_FILL
                    ws.freeze_panes = "A2"
            else:
                for r, line in enumerate(rows, start=1):
                    ws.cell(row=r, column=1, value=line)
                ws.column_dimensions["A"].width = 120

    if not out_wb.sheetnames:
        out_wb.create_sheet("Empty")

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf
