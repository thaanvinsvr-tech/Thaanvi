"""
GSTR-2B Consolidator
---------------------
Upload multiple GSTR-2B excel exports (as downloaded from the GST portal) and
get back ONE workbook with sheet-wise consolidation across all periods.

- "Read me" sheet is never consolidated.
- Every other sheet (B2B, B2BA, ITC Available, IMPG, etc.) is consolidated
  across all uploaded files into a single sheet of the same name.
- A "Tax Period" column (read from each file's "Read me" sheet) is added as
  the first column of every row, so you can tell which month each row
  belongs to after consolidation.
"""

import io
from datetime import datetime

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="GSTR-2B Consolidator", page_icon="\U0001F4CA", layout="wide")

# --------------------------------------------------------------------------
# Core extraction engine
# --------------------------------------------------------------------------

READ_ME_SHEET = "Read me"


def get_period_info(wb):
    """Read Tax Period + Financial Year from the 'Read me' sheet."""
    if READ_ME_SHEET not in wb.sheetnames:
        return None, None
    ws = wb[READ_ME_SHEET]
    fy, period = None, None
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                label = cell.value.strip().lower()
                if label == "financial year":
                    fy = ws.cell(row=cell.row, column=3).value
                elif label == "tax period":
                    period = ws.cell(row=cell.row, column=3).value
    return period, fy


def get_extra_info(wb):
    """Pull GSTIN / Legal name / Date of generation, purely for the summary tab."""
    info = {}
    if READ_ME_SHEET not in wb.sheetnames:
        return info
    ws = wb[READ_ME_SHEET]
    wanted = {"gstin", "legal name", "trade name (if any)", "date of generation"}
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                label = cell.value.strip().lower()
                if label in wanted:
                    info[cell.value.strip()] = ws.cell(row=cell.row, column=3).value
    return info


def detect_header_end_row(ws):
    """
    Header rows always start at row 5 in the GSTR-2B template (rows 1-4 are
    the sheet title / section description banners). Figure out how many rows
    the header actually spans (some sheets have 1, some 2, some 3 header
    rows) by looking at merged cells anchored in rows 5-7, ignoring any
    merge that spans the FULL width of the sheet (those are banner/section
    text, not column headers).
    """
    max_col = ws.max_column
    candidates = [
        m.max_row
        for m in ws.merged_cells.ranges
        if 5 <= m.min_row <= 7 and not (m.min_col == 1 and m.max_col == max_col)
    ]
    return max(candidates) if candidates else 6


def build_merge_map(ws, max_row):
    """Expand merged cells (within header rows only) into a {(row,col): value} map."""
    m = {}
    max_col = ws.max_column
    for rng in ws.merged_cells.ranges:
        if rng.min_row > max_row:
            continue
        if rng.min_col == 1 and rng.max_col == max_col:
            continue  # skip full-width banner merges
        anchor_val = ws.cell(row=rng.min_row, column=rng.min_col).value
        for r in range(rng.min_row, min(rng.max_row, max_row) + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                m[(r, c)] = anchor_val
    return m


def get_cell(ws, merge_map, row, col):
    if (row, col) in merge_map:
        return merge_map[(row, col)]
    return ws.cell(row=row, column=col).value


def extract_sheet(ws):
    """Return (headers, data_rows) for a single GSTR-2B data sheet."""
    header_end = detect_header_end_row(ws)
    merge_map = build_merge_map(ws, header_end)
    max_col = ws.max_column

    headers = []
    for c in range(1, max_col + 1):
        val = None
        for r in range(header_end, 4, -1):  # walk upward from most specific row
            v = get_cell(ws, merge_map, r, c)
            if v not in (None, ""):
                val = str(v).strip()
                break
        headers.append(val or f"Column {get_column_letter(c)}")

    # de-duplicate header names (pandas/Excel both choke on repeats) by
    # suffixing repeats with (2), (3), ...
    seen = {}
    unique_headers = []
    for h in headers:
        if h not in seen:
            seen[h] = 1
            unique_headers.append(h)
        else:
            seen[h] += 1
            unique_headers.append(f"{h} ({seen[h]})")
    headers = unique_headers

    data_rows = []
    for r in range(header_end + 1, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        if any(v not in (None, "") for v in row_vals):
            data_rows.append(row_vals)

    return headers, data_rows


def format_period_label(period, fy):
    if period and fy:
        return f"{period} {fy}"
    return period or fy or "Unknown Period"


def process_files(uploaded_files):
    """
    Returns:
      consolidated: { sheet_name: {"headers": [...], "rows": [...]} }
      file_summaries: [ {filename, period, gstin, legal_name, sheet_count}, ... ]
      warnings: [str, ...]
    """
    consolidated = {}   # sheet_name -> {"headers": [...], "rows": [[...]]}
    file_summaries = []
    warnings = []

    for uf in uploaded_files:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(uf.getvalue()), data_only=True)
        except Exception as e:
            warnings.append(f"Could not open '{uf.name}': {e}")
            continue

        period, fy = get_period_info(wb)
        period_label = format_period_label(period, fy)
        extra = get_extra_info(wb)

        sheet_names = [s for s in wb.sheetnames if s != READ_ME_SHEET]
        rows_added_this_file = 0

        for sheet_name in sheet_names:
            ws = wb[sheet_name]
            try:
                headers, data_rows = extract_sheet(ws)
            except Exception as e:
                warnings.append(f"Skipped sheet '{sheet_name}' in '{uf.name}': {e}")
                continue

            if not data_rows:
                # still register the sheet so it appears in the output even if empty
                consolidated.setdefault(sheet_name, {"headers": headers, "rows": []})
                continue

            if sheet_name not in consolidated:
                consolidated[sheet_name] = {"headers": headers, "rows": []}

            ref_headers = consolidated[sheet_name]["headers"]

            # Align this file's columns to the reference header order (by name).
            if headers == ref_headers:
                aligned_rows = data_rows
            else:
                # map header name -> index in this sheet's own headers
                name_to_idx = {}
                for idx, h in enumerate(headers):
                    name_to_idx.setdefault(h, idx)
                aligned_rows = []
                for row in data_rows:
                    new_row = []
                    for h in ref_headers:
                        idx = name_to_idx.get(h)
                        new_row.append(row[idx] if idx is not None and idx < len(row) else None)
                    aligned_rows.append(new_row)
                # any columns in this file not present in the reference get appended once
                extra_cols = [h for h in headers if h not in ref_headers]
                if extra_cols:
                    ref_headers = ref_headers + extra_cols
                    consolidated[sheet_name]["headers"] = ref_headers
                    # pad previously stored rows with blanks for the new columns
                    for r in consolidated[sheet_name]["rows"]:
                        r.extend([None] * len(extra_cols))
                    for i, row in enumerate(data_rows):
                        for h in extra_cols:
                            idx = name_to_idx.get(h)
                            aligned_rows[i].append(row[idx] if idx is not None and idx < len(row) else None)

            for row in aligned_rows:
                consolidated[sheet_name]["rows"].append([period_label] + list(row))
                rows_added_this_file += 1

        file_summaries.append({
            "filename": uf.name,
            "period": period_label,
            "gstin": extra.get("GSTIN", ""),
            "legal_name": extra.get("Legal Name", ""),
            "sheets": len(sheet_names),
            "rows": rows_added_this_file,
        })

    # Prefix "Tax Period" onto every sheet's header list
    for sheet_name, payload in consolidated.items():
        payload["headers"] = ["Tax Period"] + payload["headers"]

    return consolidated, file_summaries, warnings


# --------------------------------------------------------------------------
# Workbook builder
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
SUBTITLE_FONT = Font(italic=True, size=10, color="666666")


def build_output_workbook(consolidated, file_summaries):
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    # ---- Summary sheet first ----
    ws = out_wb.create_sheet("Consolidation Summary")
    ws["A1"] = "GSTR-2B Consolidation Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated on {datetime.now().strftime('%d-%b-%Y %H:%M')}"
    ws["A2"].font = SUBTITLE_FONT

    headers = ["File Name", "Tax Period", "GSTIN", "Legal Name", "Sheets Found", "Rows Consolidated"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for r, fs in enumerate(file_summaries, start=5):
        ws.cell(row=r, column=1, value=fs["filename"])
        ws.cell(row=r, column=2, value=fs["period"])
        ws.cell(row=r, column=3, value=fs["gstin"])
        ws.cell(row=r, column=4, value=fs["legal_name"])
        ws.cell(row=r, column=5, value=fs["sheets"])
        ws.cell(row=r, column=6, value=fs["rows"])
    for c, width in zip(range(1, 7), [40, 20, 20, 35, 14, 18]):
        ws.column_dimensions[get_column_letter(c)].width = width

    # ---- One sheet per GSTR-2B sheet name, consolidated ----
    for sheet_name, payload in consolidated.items():
        safe_name = sheet_name[:31]  # excel sheet name limit
        ws = out_wb.create_sheet(safe_name)
        headers = payload["headers"]
        rows = payload["rows"]

        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.freeze_panes = "A2"

        for r, row in enumerate(rows, start=2):
            for c, val in enumerate(row, start=1):
                ws.cell(row=r, column=c, value=val)

        # reasonable column widths
        ws.column_dimensions["A"].width = 22  # Tax Period
        for c in range(2, len(headers) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 18

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.title("\U0001F4CA GSTR-2B Consolidator")
st.caption(
    "Upload two or more GSTR-2B excel files (as exported from the GST portal). "
    "They'll be consolidated sheet-wise into a single workbook, with each row "
    "tagged with its Tax Period. The 'Read me' sheet is never consolidated."
)

uploaded_files = st.file_uploader(
    "Upload GSTR-2B excel files",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) < 2:
        st.info("Tip: upload two or more files to actually see consolidation across periods.")

    with st.spinner("Reading and consolidating files..."):
        consolidated, file_summaries, warnings = process_files(uploaded_files)

    if warnings:
        for w in warnings:
            st.warning(w)

    st.subheader("Files processed")
    st.dataframe(
        [
            {
                "File": fs["filename"],
                "Tax Period": fs["period"],
                "GSTIN": fs["gstin"],
                "Legal Name": fs["legal_name"],
                "Sheets": fs["sheets"],
                "Rows added": fs["rows"],
            }
            for fs in file_summaries
        ],
        use_container_width=True,
        hide_index=True,
    )

    total_rows = sum(len(p["rows"]) for p in consolidated.values())
    non_empty_sheets = sum(1 for p in consolidated.values() if p["rows"])
    c1, c2, c3 = st.columns(3)
    c1.metric("Files consolidated", len(file_summaries))
    c2.metric("Sheets with data", f"{non_empty_sheets} / {len(consolidated)}")
    c3.metric("Total data rows", total_rows)

    st.subheader("Preview")
    sheet_choice = st.selectbox(
        "Pick a sheet to preview",
        options=[s for s in consolidated.keys() if consolidated[s]["rows"]] or list(consolidated.keys()),
    )
    if sheet_choice:
        payload = consolidated[sheet_choice]
        if payload["rows"]:
            import pandas as pd
            df = pd.DataFrame(payload["rows"], columns=payload["headers"])
            st.dataframe(df, use_container_width=True, height=350)
        else:
            st.caption("No data rows in this sheet across the uploaded files.")

    output_buf = build_output_workbook(consolidated, file_summaries)
    st.download_button(
        label="\u2B07\uFE0F Download consolidated workbook",
        data=output_buf,
        file_name=f"GSTR2B_Consolidated_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload your GSTR-2B excel files to get started.")
