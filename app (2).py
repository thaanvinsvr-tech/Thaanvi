"""
GST Compliance Suite
----------------------
Single Streamlit app hosting multiple consolidator tools, picked from the
sidebar. Currently:
  - GSTR-2B Consolidator (Excel exports -> RCM/Net/Pivot + sheet-wise)
  - GSTR-3B Consolidator (PDF returns -> Output/Input bifurcated annexures)
"""

import io
from datetime import datetime

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import engine as engine2b
import engine3b
import pdf2excel
import excel2pdf
import json2excel

st.set_page_config(page_title="GST Compliance Suite", page_icon="\U0001F4CA", layout="wide")


# ==========================================================================
# Theme (nude palette, light/dark toggle)
# ==========================================================================

NUDE_LIGHT = {
    "bg": "#FBF3EA", "sidebar_bg": "#F1E1D0", "card_bg": "#FFFFFF",
    "text": "#4A372C", "accent": "#C99A76", "accent_text": "#FFFFFF",
    "border": "#E3CBB2",
}
NUDE_DARK = {
    "bg": "#2B2420", "sidebar_bg": "#3A2F28", "card_bg": "#3E332C",
    "text": "#F3E6D8", "accent": "#C99A76", "accent_text": "#2B2420",
    "border": "#54453B",
}


def inject_theme(mode):
    p = NUDE_DARK if mode == "Dark" else NUDE_LIGHT
    st.markdown(f"""
    <style>
    .stApp {{ background-color: {p['bg']}; }}
    .stApp, .stApp p, .stApp label, .stApp span, .stApp div,
    h1, h2, h3, h4, h5, h6 {{ color: {p['text']}; }}
    section[data-testid="stSidebar"] {{ background-color: {p['sidebar_bg']}; }}
    section[data-testid="stSidebar"] * {{ color: {p['text']} !important; }}
    .stButton>button, .stDownloadButton>button {{
        background-color: {p['accent']}; color: {p['accent_text']};
        border-radius: 10px; border: 1px solid {p['border']}; font-weight: 600;
    }}
    .stButton>button:hover, .stDownloadButton>button:hover {{
        opacity: 0.85; border-color: {p['accent']};
    }}
    div[data-testid="stMetric"] {{
        background-color: {p['card_bg']}; padding: 14px; border-radius: 12px;
        border: 1px solid {p['border']};
    }}
    div[data-testid="stFileUploaderDropzone"] {{
        background-color: {p['card_bg']}; border: 1.5px dashed {p['accent']};
        border-radius: 12px;
    }}
    .stTabs [data-baseweb="tab"] {{ color: {p['text']}; }}
    .stTabs [aria-selected="true"] {{
        color: {p['accent']} !important; border-bottom-color: {p['accent']} !important;
    }}
    </style>
    """, unsafe_allow_html=True)


theme_mode = st.session_state.get("theme_mode_radio", "Light")
inject_theme(theme_mode)


# ==========================================================================
# Shared helpers
# ==========================================================================

HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
SUBTITLE_FONT = Font(italic=True, size=10, color="666666")
INVALID_SHEET_CHARS = str.maketrans({c: "-" for c in r'\/?*[]:'})


def safe_sheet_name(name):
    return name.translate(INVALID_SHEET_CHARS)[:31]


# ==========================================================================
# GSTR-2B tool
# ==========================================================================

HEADER_FILL_2B = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")


def build_2b_workbook(result_sheets, file_summaries):
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    ws = out_wb.create_sheet("Consolidation Summary")
    ws["A1"] = "GSTR-2B Consolidation Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated on {datetime.now().strftime('%d-%b-%Y %H:%M')}"
    ws["A2"].font = SUBTITLE_FONT

    headers = ["File Name", "Month", "Financial Year", "GSTIN", "Legal Name", "Sheets Found", "Rows Consolidated"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL_2B
    for r, fs in enumerate(file_summaries, start=5):
        ws.cell(row=r, column=1, value=fs["filename"])
        ws.cell(row=r, column=2, value=fs["period"])
        ws.cell(row=r, column=3, value=fs["financial_year"])
        ws.cell(row=r, column=4, value=fs["gstin"])
        ws.cell(row=r, column=5, value=fs["legal_name"])
        ws.cell(row=r, column=6, value=fs["sheets"])
        ws.cell(row=r, column=7, value=fs["rows"])
    for c, width in zip(range(1, 8), [40, 14, 16, 20, 35, 14, 18]):
        ws.column_dimensions[get_column_letter(c)].width = width

    for sheet_name, payload in result_sheets.items():
        ws = out_wb.create_sheet(safe_sheet_name(sheet_name))
        headers = payload["headers"]
        rows = payload["rows"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL_2B
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.freeze_panes = "A2"
        for r, row in enumerate(rows, start=2):
            for c, val in enumerate(row, start=1):
                ws.cell(row=r, column=c, value=val)
        ws.column_dimensions["A"].width = 14
        for c in range(2, len(headers) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 18

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf


def render_gstr2b():
    st.title("\U0001F4CA GSTR-2B Consolidator")
    st.caption(
        "Upload two or more GSTR-2B excel files (as exported from the GST portal). "
        "They'll be consolidated into RCM / B2B / B2BA / B2B-CDNR / Pivot / Net / "
        "IMPG / IMPGSEZ / IMPGA sheets, with any other sheet types appended after. "
        "The 'Read me' sheet is never consolidated."
    )

    uploaded_files = st.file_uploader(
        "Upload GSTR-2B excel files", type=["xlsx"], accept_multiple_files=True, key="gstr2b_uploader",
    )

    if not uploaded_files:
        st.info("Upload your GSTR-2B excel files to get started.")
        return

    if len(uploaded_files) < 2:
        st.info("Tip: upload two or more files to actually see consolidation across periods.")

    with st.spinner("Reading and consolidating files..."):
        result_sheets, file_summaries, warnings = engine2b.process_files(uploaded_files)

    for w in warnings:
        st.warning(w)

    st.subheader("Files processed")
    st.dataframe(
        [
            {
                "File": fs["filename"], "Month": fs["period"], "FY": fs["financial_year"],
                "GSTIN": fs["gstin"], "Legal Name": fs["legal_name"],
                "Sheets": fs["sheets"], "Rows added": fs["rows"],
            }
            for fs in file_summaries
        ],
        width='stretch', hide_index=True,
    )

    total_rows = sum(len(p["rows"]) for p in result_sheets.values())
    non_empty_sheets = sum(1 for p in result_sheets.values() if p["rows"])
    c1, c2, c3 = st.columns(3)
    c1.metric("Files consolidated", len(file_summaries))
    c2.metric("Sheets with data", f"{non_empty_sheets} / {len(result_sheets)}")
    c3.metric("Total data rows", total_rows)

    import pandas as pd

    if "Pivot" in result_sheets and result_sheets["Pivot"]["rows"]:
        st.subheader("Month-wise ITC totals (Pivot)")
        pdf = pd.DataFrame(result_sheets["Pivot"]["rows"], columns=result_sheets["Pivot"]["headers"])
        st.dataframe(pdf, width='stretch', hide_index=True)

    st.subheader("Preview a sheet")
    sheet_choice = st.selectbox(
        "Pick a sheet to preview",
        options=[s for s in result_sheets.keys() if result_sheets[s]["rows"]] or list(result_sheets.keys()),
        key="gstr2b_sheet_choice",
    )
    if sheet_choice:
        payload = result_sheets[sheet_choice]
        if payload["rows"]:
            df = pd.DataFrame(payload["rows"], columns=payload["headers"])
            st.dataframe(df, width='stretch', height=350)
        else:
            st.caption("No data rows in this sheet across the uploaded files.")

    output_buf = build_2b_workbook(result_sheets, file_summaries)
    st.download_button(
        label="\u2B07\uFE0F Download consolidated workbook",
        data=output_buf,
        file_name=f"GSTR2B_Consolidated_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="gstr2b_download",
    )


# ==========================================================================
# GSTR-3B tool
# ==========================================================================

OUTPUT_FILL_3B = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
INPUT_FILL_3B = PatternFill(start_color="6E2C00", end_color="6E2C00", fill_type="solid")
OTHER_FILL_3B = PatternFill(start_color="4A4A4A", end_color="4A4A4A", fill_type="solid")

TAB_COLORS_3B = {}
for _name in engine3b.OUTPUT_SHEETS:
    TAB_COLORS_3B[_name] = "1F4E78"
for _name in engine3b.INPUT_SHEETS:
    TAB_COLORS_3B[_name] = "6E2C00"
for _name in engine3b.OTHER_SHEETS:
    TAB_COLORS_3B[_name] = "4A4A4A"


def build_3b_workbook(result_sheets, file_summaries):
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    ws = out_wb.create_sheet("Filing Summary")
    ws["A1"] = "GSTR-3B Consolidation Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated on {datetime.now().strftime('%d-%b-%Y %H:%M')}"
    ws["A2"].font = SUBTITLE_FONT

    headers = ["File Name", "Month", "Financial Year", "GSTIN", "Legal Name", "ARN", "Date of Filing", "Rows Extracted"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = OUTPUT_FILL_3B
    for r, fs in enumerate(file_summaries, start=5):
        ws.cell(row=r, column=1, value=fs["filename"])
        ws.cell(row=r, column=2, value=fs["period"])
        ws.cell(row=r, column=3, value=fs["financial_year"])
        ws.cell(row=r, column=4, value=fs["gstin"])
        ws.cell(row=r, column=5, value=fs["legal_name"])
        ws.cell(row=r, column=6, value=fs["arn"])
        ws.cell(row=r, column=7, value=fs["date_of_filing"])
        ws.cell(row=r, column=8, value=fs["rows"])
    for c, width in zip(range(1, 9), [40, 12, 14, 20, 35, 20, 16, 14]):
        ws.column_dimensions[get_column_letter(c)].width = width

    for sheet_name, payload in result_sheets.items():
        ws = out_wb.create_sheet(safe_sheet_name(sheet_name))
        ws.sheet_properties.tabColor = TAB_COLORS_3B.get(sheet_name, "1F4E78")
        fill = OUTPUT_FILL_3B if sheet_name in engine3b.OUTPUT_SHEETS else (
            INPUT_FILL_3B if sheet_name in engine3b.INPUT_SHEETS else OTHER_FILL_3B)

        headers = payload["headers"]
        rows = payload["rows"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.freeze_panes = "A2"
        for r, row in enumerate(rows, start=2):
            for c, val in enumerate(row, start=1):
                ws.cell(row=r, column=c, value=val)
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 42
        for c in range(3, len(headers) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 18

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf


def render_gstr3b():
    st.title("\U0001F4CB GSTR-3B Consolidator")
    st.caption(
        "Upload two or more GSTR-3B PDFs (as downloaded from the GST portal). "
        "Every table becomes its own annexure sheet — bifurcated into Output "
        "(outward supplies, export/SEZ, RCM liability) and Input (ITC) — with "
        "each row tagged by Month and consolidated across all uploaded periods."
    )

    uploaded_files = st.file_uploader(
        "Upload GSTR-3B PDF files", type=["pdf"], accept_multiple_files=True, key="gstr3b_uploader",
    )

    if not uploaded_files:
        st.info("Upload your GSTR-3B PDF files to get started.")
        return

    if len(uploaded_files) < 2:
        st.info("Tip: upload two or more months to actually see consolidation across periods.")

    with st.spinner("Reading and parsing PDFs..."):
        result_sheets, file_summaries, warnings = engine3b.process_files(uploaded_files)

    for w in warnings:
        st.warning(w)

    st.subheader("Files processed")
    st.dataframe(
        [
            {
                "File": fs["filename"], "Month": fs["period"], "FY": fs["financial_year"],
                "GSTIN": fs["gstin"], "Legal Name": fs["legal_name"],
                "ARN": fs["arn"], "Filed on": fs["date_of_filing"], "Rows extracted": fs["rows"],
            }
            for fs in file_summaries
        ],
        width='stretch', hide_index=True,
    )

    total_rows = sum(len(p["rows"]) for p in result_sheets.values())
    c1, c2, c3 = st.columns(3)
    c1.metric("Files consolidated", len(file_summaries))
    c2.metric("Annexure sheets", len(result_sheets))
    c3.metric("Total rows extracted", total_rows)

    import pandas as pd

    tab_output, tab_input, tab_other = st.tabs(["\U0001F4E4 Output", "\U0001F4E5 Input (ITC)", "\U0001F4C4 Other"])

    def render_group(container, sheet_names):
        with container:
            for name in sheet_names:
                payload = result_sheets[name]
                if payload["rows"]:
                    st.markdown(f"**{name}**")
                    df = pd.DataFrame(payload["rows"], columns=payload["headers"])
                    st.dataframe(df, width='stretch', hide_index=True)

    render_group(tab_output, engine3b.OUTPUT_SHEETS)
    render_group(tab_input, engine3b.INPUT_SHEETS)
    render_group(tab_other, engine3b.OTHER_SHEETS)

    output_buf = build_3b_workbook(result_sheets, file_summaries)
    st.download_button(
        label="\u2B07\uFE0F Download consolidated workbook",
        data=output_buf,
        file_name=f"GSTR3B_Consolidated_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="gstr3b_download",
    )


# ==========================================================================
# PDF -> Excel tool
# ==========================================================================

def render_pdf2excel():
    st.title("\U0001F4C4\u27A1\uFE0F\U0001F4CA PDF to Excel Converter")
    st.caption(
        "Upload one or more PDFs. Every table pdfplumber can detect is extracted "
        "onto its own sheet; pages with no table fall back to a plain-text sheet "
        "so nothing gets silently dropped."
    )

    uploaded_files = st.file_uploader(
        "Upload PDF files", type=["pdf"], accept_multiple_files=True, key="pdf2excel_uploader",
    )
    if not uploaded_files:
        st.info("Upload one or more PDF files to get started.")
        return

    col1, col2 = st.columns(2)
    with col1:
        merge_across = st.checkbox(
            "Merge multi-page tables into one sheet",
            value=True, key="pdf2excel_merge_across",
            help="On (default): consecutive tables with the same column count (e.g. a bank statement or ledger spanning many pages) are combined into a single sheet, with any repeated header row stripped. Off: every table gets its own per-page sheet.",
        )
    with col2:
        merge_per_page = st.checkbox(
            "Merge multiple tables on the same page",
            value=False, key="pdf2excel_merge_per_page",
            help="On: if a single page has more than one detected table, stack them into one sheet separated by a blank row. Off (default): each gets its own sheet.",
        )

    with st.spinner("Extracting tables..."):
        all_file_sheets = []
        total_tables = 0
        for uf in uploaded_files:
            sheets = pdf2excel.extract_pdf_to_sheets(
                uf.getvalue(), merge_tables_per_page=merge_per_page, merge_across_pages=merge_across,
            )
            all_file_sheets.append((uf.name, sheets))
            total_tables += len(sheets)

    c1, c2 = st.columns(2)
    c1.metric("Files processed", len(uploaded_files))
    c2.metric("Sheets extracted", total_tables)

    st.subheader("Preview")
    flat_options = []
    for filename, sheets in all_file_sheets:
        for sheet_name, kind, rows in sheets:
            flat_options.append((f"{filename} — {sheet_name}", kind, rows))
    if flat_options:
        labels = [o[0] for o in flat_options]
        choice = st.selectbox("Pick a sheet to preview", labels, key="pdf2excel_preview_choice")
        _, kind, rows = flat_options[labels.index(choice)]
        if kind == "table" and rows:
            import pandas as pd
            width = max(len(r) for r in rows)
            padded = [list(r) + [None] * (width - len(r)) for r in rows]
            df = pd.DataFrame(padded[1:], columns=[str(c) for c in padded[0]]) if len(padded) > 1 else pd.DataFrame(padded)
            st.dataframe(df, width='stretch', height=300)
        else:
            st.text("\n".join(rows[:50]))
    else:
        st.caption("No tables or text could be extracted from the uploaded PDF(s).")

    output_buf = pdf2excel.build_workbook(all_file_sheets)
    st.download_button(
        label="\u2B07\uFE0F Download Excel workbook",
        data=output_buf,
        file_name=f"PDF_to_Excel_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="pdf2excel_download",
    )


# ==========================================================================
# Excel -> PDF tool
# ==========================================================================

def render_excel2pdf():
    st.title("\U0001F4CA\u27A1\uFE0F\U0001F4C4 Excel to PDF Converter")
    st.caption(
        "Upload one or more Excel workbooks. Every sheet is rendered as a table "
        "in a landscape PDF, one section per sheet, auto-paginated for long sheets."
    )

    uploaded_files = st.file_uploader(
        "Upload Excel files", type=["xlsx"], accept_multiple_files=True, key="excel2pdf_uploader",
    )
    if not uploaded_files:
        st.info("Upload one or more .xlsx files to get started.")
        return

    for uf in uploaded_files:
        with st.spinner(f"Converting {uf.name}..."):
            try:
                pdf_buf = excel2pdf.workbook_to_pdf(uf.getvalue(), uf.name)
            except Exception as e:
                st.error(f"Could not convert '{uf.name}': {e}")
                continue
        st.download_button(
            label=f"\u2B07\uFE0F Download {uf.name.rsplit('.', 1)[0]}.pdf",
            data=pdf_buf,
            file_name=f"{uf.name.rsplit('.', 1)[0]}.pdf",
            mime="application/pdf",
            key=f"excel2pdf_download_{uf.name}",
        )


# ==========================================================================
# JSON -> Excel tool
# ==========================================================================

def render_json2excel():
    st.title("\U0001F5C2\uFE0F\u27A1\uFE0F\U0001F4CA JSON to Excel Converter")
    st.caption(
        "Upload one or more JSON files. Top-level keys become sheets; nested "
        "objects flatten into dotted column names; nested arrays of records "
        "(e.g. GST-style invoice -> item structures) are exploded into rows "
        "with their parent fields repeated."
    )

    uploaded_files = st.file_uploader(
        "Upload JSON files", type=["json"], accept_multiple_files=True, key="json2excel_uploader",
    )
    if not uploaded_files:
        st.info("Upload one or more .json files to get started.")
        return

    with st.spinner("Parsing JSON..."):
        all_file_sheets = []
        warnings = []
        for uf in uploaded_files:
            try:
                sheets = json2excel.parse_json_to_sheets(uf.getvalue())
                all_file_sheets.append((uf.name, sheets))
            except Exception as e:
                warnings.append(f"Could not parse '{uf.name}': {e}")

    for w in warnings:
        st.warning(w)

    if not all_file_sheets:
        return

    total_sheets = sum(len(s) for _, s in all_file_sheets)
    c1, c2 = st.columns(2)
    c1.metric("Files processed", len(all_file_sheets))
    c2.metric("Sheets produced", total_sheets)

    st.subheader("Preview")
    flat_options = []
    for filename, sheets in all_file_sheets:
        for sheet_name, df in sheets:
            flat_options.append((f"{filename} — {sheet_name}", df))
    if flat_options:
        labels = [o[0] for o in flat_options]
        choice = st.selectbox("Pick a sheet to preview", labels, key="json2excel_preview_choice")
        df = flat_options[labels.index(choice)][1]
        st.dataframe(df, width='stretch', height=300)

    output_buf = json2excel.build_workbook(all_file_sheets)
    st.download_button(
        label="\u2B07\uFE0F Download Excel workbook",
        data=output_buf,
        file_name=f"JSON_to_Excel_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="json2excel_download",
    )


# ==========================================================================
# Sidebar navigation
# ==========================================================================

st.sidebar.title("GST Compliance Suite")
tool = st.sidebar.radio(
    "Choose a tool",
    [
        "GSTR-2B Consolidator",
        "GSTR-3B Consolidator",
        "PDF to Excel Converter",
        "Excel to PDF Converter",
        "JSON to Excel Converter",
    ],
)
st.sidebar.markdown("---")
st.sidebar.caption("More tools coming soon.")

st.sidebar.markdown("### Appearance")
st.sidebar.radio("Theme", ["Light", "Dark"], horizontal=True, key="theme_mode_radio")

RENDERERS = {
    "GSTR-2B Consolidator": render_gstr2b,
    "GSTR-3B Consolidator": render_gstr3b,
    "PDF to Excel Converter": render_pdf2excel,
    "Excel to PDF Converter": render_excel2pdf,
    "JSON to Excel Converter": render_json2excel,
}
RENDERERS[tool]()
