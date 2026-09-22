"""
GSTR-3B PDF Consolidation Engine
----------------------------------
Parses GSTR-3B PDF returns (as downloaded from the GST portal) and produces
one consolidated workbook, in the same spirit as the GSTR-2B consolidator:
one sheet per table/annexure, each row tagged with its Month, stacked across
every uploaded period.

Output is grouped as:
  OUTPUT  : Outward Taxable Supplies, Zero Rated (Export/SEZ), Nil Rated &
            Exempted, RCM Inward Liability, Non-GST Outward Supplies,
            ECO Supplies u/s 9(5), Interstate Supplies Breakup
  INPUT   : ITC Available, ITC Reversed, Net ITC Available, Ineligible/Other
            ITC Details
  OTHER   : Exempt/Nil/Non-GST Inward, Interest & Late Fee, Tax Payment
            Summary, Liability Breakup, Filing Summary

The GSTR-3B PDF template stamps a "FILED" watermark whose individual
letters can bleed into table cells at unpredictable positions depending on
page layout — these are stripped during cleaning. Rows are matched by their
label text (not position), so the parser is resilient to the watermark and
to tables splitting across a page break.
"""

import io
import re

import pdfplumber

MONTH_ORDER = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]


# --------------------------------------------------------------------------
# Low-level cleaning
# --------------------------------------------------------------------------

def clean_cell(v):
    """Strip stray single-uppercase-letter watermark lines and normalise whitespace."""
    if v is None:
        return None
    v = str(v)
    lines = [ln for ln in v.split("\n") if not re.fullmatch(r"[A-Z]", ln.strip())]
    v = " ".join(lines)
    v = re.sub(r"\s+", " ", v).strip()
    return v


def to_number(v):
    """Parse a GSTR-3B numeric cell. '-' means not applicable -> 0.0."""
    if v is None:
        return None
    s = re.sub(r"\s+", "", str(v))
    if s in ("", "-", "\u2014"):
        return 0.0
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def find_row(rows, label_substr, start=0):
    """First row (from index `start`) whose first cell contains label_substr (case-insensitive)."""
    needle = label_substr.lower()
    for i in range(start, len(rows)):
        cell0 = rows[i][0]
        if cell0 and needle in cell0.lower():
            return i, rows[i]
    return None, None


# --------------------------------------------------------------------------
# Header info (GSTIN, legal name, period, ARN, filing date)
# --------------------------------------------------------------------------

def extract_header_info(all_rows, full_text):
    info = {}
    _, r = find_row(all_rows, "Period")
    if r:
        info["Month"] = r[1] if len(r) > 1 else None
    _, r = find_row(all_rows, "Year")
    if r:
        info["Financial Year"] = r[1] if len(r) > 1 else None
    _, r = find_row(all_rows, "GSTIN of the supplier")
    if r:
        info["GSTIN"] = r[1] if len(r) > 1 else None
    _, r = find_row(all_rows, "Legal name of the registered person")
    if r:
        info["Legal Name"] = r[1] if len(r) > 1 else None
    _, r = find_row(all_rows, "Trade name")
    if r:
        info["Trade Name"] = r[1] if len(r) > 1 else None
    _, r = find_row(all_rows, "ARN")
    if r and "Date" not in (r[0] or ""):
        info["ARN"] = r[1] if len(r) > 1 else None
    _, r = find_row(all_rows, "Date of ARN")
    if r:
        info["Date of Filing"] = r[1] if len(r) > 1 else None

    # Fallback for Month/Year from the free-text header line if the table
    # parse above didn't line up (e.g. "Year 2025-26" / "Period March").
    if not info.get("Month"):
        m = re.search(r"Period\s+([A-Za-z]+)", full_text)
        if m:
            info["Month"] = m.group(1)
    if not info.get("Financial Year"):
        m = re.search(r"Year\s+(\d{4}-\d{2})", full_text)
        if m:
            info["Financial Year"] = m.group(1)

    return info


# --------------------------------------------------------------------------
# Table 3.1 / 3.1.1 / 3.2 — OUTPUT side
# --------------------------------------------------------------------------

OUTPUT_31_COLS = ["Taxable Value (₹)", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)"]

OUTPUT_31_ROWS = [
    ("Outward Taxable Supplies", "(a) Outward taxable supplies"),
    ("Zero Rated (Export/SEZ)", "(b) Outward taxable supplies (zero rated)"),
    ("Nil Rated & Exempted", "Other outward supplies (nil rated"),
    ("RCM Inward Liability", "Inward supplies (liable to reverse charge)"),
    ("Non-GST Outward Supplies", "Non-GST outward supplies"),
]

OUTPUT_311_ROWS = [
    ("(i) ECO pays tax u/s 9(5)", "Taxable supplies on which electronic commerce operator pays tax"),
    ("(ii) Supplies through ECO u/s 9(5)", "Taxable supplies made by registered person through electronic commerce"),
]

INTERSTATE_32_ROWS = [
    ("Supplies to Unregistered Persons", "Supplies made to Unregistered Persons"),
    ("Supplies to Composition Taxable Persons", "Supplies made to Composition Taxable"),
    ("Supplies to UIN Holders", "Supplies made to UIN holders"),
]

# --------------------------------------------------------------------------
# Table 4 — INPUT side (Eligible ITC)
# --------------------------------------------------------------------------

ITC_COLS = ["Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)"]

ITC_AVAILABLE_ROWS = [
    ("(1) Import of goods", "(1) Import of goods"),
    ("(2) Import of services", "(2) Import of services"),
    ("(3) Inward supplies liable to RCM", "Inward supplies liable to reverse charge (other than"),
    ("(4) Inward supplies from ISD", "Inward supplies from ISD"),
    ("(5) All other ITC", "All other ITC"),
]

ITC_REVERSED_ROWS = [
    ("(1) As per Rules 38, 42 & 43 / Sec 17(5)", "As per rules 38,42 & 43"),
    ("(2) Others", "(2) Others"),
]

ITC_NET_ROWS = [
    ("Net ITC Available (A-B)", "Net ITC available"),
]

ITC_OTHER_ROWS = [
    ("(D) Other Details (subtotal)", "(D) Other Details"),
    ("(1) ITC reclaimed (reversed under 4(B)(2) earlier period)", "ITC reclaimed which was reversed"),
    ("(2) Ineligible ITC u/s 16(4) & PoS restricted", "Ineligible ITC under section 16(4)"),
]

# --------------------------------------------------------------------------
# Table 5 / 5.1
# --------------------------------------------------------------------------

EXEMPT_5_COLS = ["Inter-State Supplies (₹)", "Intra-State Supplies (₹)"]
EXEMPT_5_ROWS = [
    ("Composition/Exempt/Nil rated (from supplier)", "From a supplier under composition scheme"),
    ("Non-GST Supply", "Non GST supply"),
]

INTEREST_51_COLS = ["Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)"]
INTEREST_51_ROWS = [
    ("System Computed Interest", "System computed Interest"),
    ("Interest Paid", "Interest Paid"),
    ("Late Fee", "Late fee"),
]


def extract_simple_table(all_rows, row_specs, n_cols):
    """Generic: for each (label, anchor) pair, find the row and pull n_cols numbers."""
    out = []
    cursor = 0
    for label, anchor in row_specs:
        idx, row = find_row(all_rows, anchor, start=cursor)
        if row is None:
            out.append([label] + [None] * n_cols)
            continue
        vals = [to_number(v) for v in row[1:1 + n_cols]]
        while len(vals) < n_cols:
            vals.append(None)
        out.append([label] + vals)
        cursor = idx + 1
    return out


# --------------------------------------------------------------------------
# Table 6.1 — Tax Payment Summary (11 columns, 2 sections x 4 tax types)
# --------------------------------------------------------------------------

PAYMENT_61_COLS = [
    "Section", "Tax Type", "Tax Payable (₹)", "Adjustment of Negative Liability (₹)",
    "Net Tax Payable (₹)", "Paid via ITC - Integrated (₹)", "Paid via ITC - Central (₹)",
    "Paid via ITC - State/UT (₹)", "Paid via ITC - Cess (₹)", "Tax Paid in Cash (₹)",
    "Interest Paid in Cash (₹)", "Late Fee Paid in Cash (₹)",
]


def extract_payment_table(all_rows):
    out = []
    idx_a, _ = find_row(all_rows, "(A) Other than reverse charge")
    idx_b, _ = find_row(all_rows, "(B) Reverse charge and supplies made")
    if idx_a is None or idx_b is None:
        return out

    def four_rows(start, end, section_label):
        rows = []
        for label in ["Integrated tax", "Central tax", "State/UT tax", "Cess"]:
            idx, row = find_row(all_rows, label, start=start)
            if row is None or (idx is not None and idx >= end):
                rows.append([section_label, label] + [None] * 10)
                continue
            vals = [to_number(v) for v in row[1:11]]
            while len(vals) < 10:
                vals.append(None)
            rows.append([section_label, label] + vals)
            start = idx + 1
        return rows

    end_a = idx_b
    end_b = len(all_rows)
    out += four_rows(idx_a + 1, end_a, "(A) Other than reverse charge")
    out += four_rows(idx_b + 1, end_b, "(B) Reverse charge / Section 9(5)")
    return out


def extract_liability_breakup(all_rows):
    idx, _ = find_row(all_rows, "Period")
    # find the "Period" row that is followed by an actual month-year data row
    for i in range(len(all_rows)):
        row = all_rows[i]
        if row[0] and re.match(r"^[A-Za-z]+ \d{4}$", row[0].strip()):
            vals = [to_number(v) for v in row[1:5]]
            while len(vals) < 4:
                vals.append(None)
            return [[row[0]] + vals]
    return []


# --------------------------------------------------------------------------
# Top-level: parse one PDF
# --------------------------------------------------------------------------

def parse_gstr3b_pdf(file_bytes):
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        all_rows = []
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"
            for table in page.extract_tables():
                for row in table:
                    all_rows.append([clean_cell(c) for c in row])

    header = extract_header_info(all_rows, full_text)
    month = header.get("Month")

    tables = {}
    tables["Outward Taxable Supplies"] = extract_simple_table(all_rows, OUTPUT_31_ROWS[:1], 5)
    tables["Zero Rated (Export SEZ)"] = extract_simple_table(all_rows, OUTPUT_31_ROWS[1:2], 5)
    tables["Nil Rated & Exempted"] = extract_simple_table(all_rows, OUTPUT_31_ROWS[2:3], 5)
    tables["RCM Inward Liability"] = extract_simple_table(all_rows, OUTPUT_31_ROWS[3:4], 5)
    tables["Non-GST Outward Supplies"] = extract_simple_table(all_rows, OUTPUT_31_ROWS[4:5], 5)
    tables["ECO Supplies 9(5)"] = extract_simple_table(all_rows, OUTPUT_311_ROWS, 5)
    tables["Interstate Supplies Breakup"] = extract_simple_table(all_rows, INTERSTATE_32_ROWS, 2)

    tables["ITC Available"] = extract_simple_table(all_rows, ITC_AVAILABLE_ROWS, 4)
    tables["ITC Reversed"] = extract_simple_table(all_rows, ITC_REVERSED_ROWS, 4)
    tables["Net ITC Available"] = extract_simple_table(all_rows, ITC_NET_ROWS, 4)
    tables["Ineligible / Other ITC Details"] = extract_simple_table(all_rows, ITC_OTHER_ROWS, 4)

    tables["Exempt Nil NonGST Inward"] = extract_simple_table(all_rows, EXEMPT_5_ROWS, 2)
    tables["Interest & Late Fee"] = extract_simple_table(all_rows, INTEREST_51_ROWS, 4)
    tables["Tax Payment Summary"] = extract_payment_table(all_rows)
    tables["Liability Breakup"] = extract_liability_breakup(all_rows)

    return header, tables


# --------------------------------------------------------------------------
# Consolidation across multiple uploaded PDFs
# --------------------------------------------------------------------------

SHEET_COLUMN_SPECS = {
    "Outward Taxable Supplies": ["Category"] + OUTPUT_31_COLS,
    "Zero Rated (Export SEZ)": ["Category"] + OUTPUT_31_COLS,
    "Nil Rated & Exempted": ["Category"] + OUTPUT_31_COLS,
    "RCM Inward Liability": ["Category"] + OUTPUT_31_COLS,
    "Non-GST Outward Supplies": ["Category"] + OUTPUT_31_COLS,
    "ECO Supplies 9(5)": ["Category"] + OUTPUT_31_COLS,
    "Interstate Supplies Breakup": ["Category", "Taxable Value (₹)", "Integrated Tax (₹)"],
    "ITC Available": ["Category"] + ITC_COLS,
    "ITC Reversed": ["Category"] + ITC_COLS,
    "Net ITC Available": ["Category"] + ITC_COLS,
    "Ineligible / Other ITC Details": ["Category"] + ITC_COLS,
    "Exempt Nil NonGST Inward": ["Category"] + EXEMPT_5_COLS,
    "Interest & Late Fee": ["Category"] + INTEREST_51_COLS,
    "Tax Payment Summary": PAYMENT_61_COLS,
    "Liability Breakup": ["Period", "Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)"],
}

# Grouping + order for the final workbook.
OUTPUT_SHEETS = ["Outward Taxable Supplies", "Zero Rated (Export SEZ)", "Nil Rated & Exempted",
                 "RCM Inward Liability", "Non-GST Outward Supplies", "ECO Supplies 9(5)",
                 "Interstate Supplies Breakup"]
INPUT_SHEETS = ["ITC Available", "ITC Reversed", "Net ITC Available", "Ineligible / Other ITC Details"]
OTHER_SHEETS = ["Exempt Nil NonGST Inward", "Interest & Late Fee", "Tax Payment Summary", "Liability Breakup"]


def process_files(uploaded_files):
    """
    Returns:
      result_sheets: { sheet_name: {"headers": ["Month", ...], "rows": [...]} }
      file_summaries: [...]
      warnings: [...]
    """
    result_sheets = {name: {"headers": ["Month"] + SHEET_COLUMN_SPECS[name][1:] if False else None, "rows": []}
                     for name in SHEET_COLUMN_SPECS}
    # headers: Month + the sheet's own columns (Category/Section/etc. + values)
    for name in SHEET_COLUMN_SPECS:
        result_sheets[name]["headers"] = ["Month"] + SHEET_COLUMN_SPECS[name]

    file_summaries = []
    warnings = []

    for uf in uploaded_files:
        try:
            header, tables = parse_gstr3b_pdf(uf.getvalue())
        except Exception as e:
            warnings.append(f"Could not parse '{uf.name}': {e}")
            continue

        month = header.get("Month")
        if not month:
            warnings.append(f"Could not detect the tax period (Month) in '{uf.name}' — rows were still extracted.")

        rows_added = 0
        for sheet_name, rows in tables.items():
            for row in rows:
                result_sheets[sheet_name]["rows"].append([month] + row)
                rows_added += 1

        file_summaries.append({
            "filename": uf.name,
            "period": month,
            "financial_year": header.get("Financial Year"),
            "gstin": header.get("GSTIN"),
            "legal_name": header.get("Legal Name"),
            "arn": header.get("ARN"),
            "date_of_filing": header.get("Date of Filing"),
            "rows": rows_added,
        })

    ordered = {}
    for name in OUTPUT_SHEETS + INPUT_SHEETS + OTHER_SHEETS:
        ordered[name] = result_sheets[name]

    return ordered, file_summaries, warnings
