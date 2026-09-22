"""
PDF -> Excel Converter
------------------------
Extracts every table pdfplumber can find in an uploaded PDF. Tables that
continue across pages (the common case for bank statements, ledgers, and
other long tabular PDFs — same column count on consecutive pages, header
only repeated on the first page or not at all) are automatically merged
into ONE sheet instead of one sheet per page. Tables with a different
shape (different column count) start a new sheet, since that usually means
a genuinely different table. Pages with no detected table fall back to a
plain-text sheet so nothing is silently dropped.
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


def _raw_extract(file_bytes):
    """
    Returns an ordered list of items, one per page:
      {"page": n, "tables": [table_rows, ...]} or {"page": n, "text": [lines]}
    (tables is empty list and text is used only when no tables were found)
    """
    items = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if tables:
                items.append({"page": page_num, "tables": tables})
            else:
                text = page.extract_text() or ""
                lines = [ln for ln in text.split("\n") if ln.strip()]
                items.append({"page": page_num, "text": lines})
    return items


def extract_pdf_to_sheets(file_bytes, merge_tables_per_page=False, merge_across_pages=True):
    """
    Returns a list of (sheet_name, kind, rows) tuples.
    kind is "table" (2D grid of cell values) or "text" (list of single lines).

    merge_across_pages: when True (default), consecutive tables that share
    the same column count are treated as one continued table and merged
    into a single sheet — a repeated header row (identical to the first
    table's first row) is stripped from continuation tables so it doesn't
    appear twice.
    """
    items = _raw_extract(file_bytes)

    if not merge_across_pages:
        sheets = []
        for item in items:
            if "tables" in item:
                tables = item["tables"]
                if merge_tables_per_page and len(tables) > 1:
                    merged = []
                    for t in tables:
                        merged.extend(t)
                        merged.append([])
                    sheets.append((f"Page{item['page']}", "table", merged))
                else:
                    for t_num, table in enumerate(tables, start=1):
                        name = f"Page{item['page']}" if len(tables) == 1 else f"Page{item['page']}_T{t_num}"
                        sheets.append((name, "table", table))
            elif item.get("text"):
                sheets.append((f"Page{item['page']}_Text", "text", item["text"]))
        return sheets

    # --- merge_across_pages=True: build a flat ordered list of tables, each
    # tagged with its page and column count, then cluster consecutive
    # same-column-count tables together. ---
    flat_tables = []  # (page, table_rows, ncols)
    text_items = {}   # page -> lines, for pages with no table at all
    for item in items:
        if "tables" in item:
            for table in item["tables"]:
                ncols = max((len(r) for r in table), default=0)
                flat_tables.append((item["page"], table, ncols))
        elif item.get("text"):
            text_items[item["page"]] = item["text"]

    sheets = []
    cluster = []  # list of (page, table_rows)
    cluster_ncols = None

    def flush_cluster():
        if not cluster:
            return
        first_page = cluster[0][0]
        last_page = cluster[-1][0]
        if len(cluster) == 1:
            name = f"Page{first_page}"
            sheets.append((name, "table", cluster[0][1]))
            return

        name = f"Page{first_page}-{last_page}"
        header = cluster[0][1][0] if cluster[0][1] else None
        merged_rows = []
        for i, (page, rows) in enumerate(cluster):
            if i == 0:
                merged_rows.extend(rows)
            else:
                # strip a repeated header row if this table's first row
                # exactly matches the cluster's header
                if header is not None and rows and list(rows[0]) == list(header):
                    merged_rows.extend(rows[1:])
                else:
                    merged_rows.extend(rows)
        sheets.append((name, "table", merged_rows))

    for page, table, ncols in flat_tables:
        if cluster and ncols == cluster_ncols:
            cluster.append((page, table))
        else:
            flush_cluster()
            cluster = [(page, table)]
            cluster_ncols = ncols
    flush_cluster()

    # append text-only pages (pages with no table) as their own sheets, in
    # page order relative to everything else
    for page, lines in sorted(text_items.items()):
        sheets.append((f"Page{page}_Text", "text", lines))

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
