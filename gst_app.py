import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="GST Reconciliation Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
    <style>
    * {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    .main {
        padding-top: 0rem;
        background-color: #f8f9fa;
    }
    
    .header-box {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #7aa8d8 100%);
        padding: 40px;
        border-radius: 15px;
        color: white;
        margin-bottom: 40px;
        box-shadow: 0 8px 32px rgba(30, 60, 114, 0.3);
    }
    
    .header-box h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    
    .header-box p {
        margin: 10px 0 0 0;
        font-size: 1.1rem;
        opacity: 0.95;
    }
    
    .info-box {
        background-color: #e3f2fd;
        padding: 20px 25px;
        border-radius: 10px;
        border-left: 5px solid #1976d2;
        margin-bottom: 20px;
        color: #0d47a1;
        font-weight: 500;
    }
    
    .success-box {
        background-color: #e8f5e9;
        padding: 20px 25px;
        border-radius: 10px;
        border-left: 5px solid #43a047;
        color: #1b5e20;
        margin-bottom: 20px;
        font-weight: 500;
    }
    
    .error-box {
        background-color: #ffebee;
        padding: 20px 25px;
        border-radius: 10px;
        border-left: 5px solid #e53935;
        color: #b71c1c;
        margin-bottom: 20px;
        font-weight: 500;
    }
    
    .upload-section {
        background: linear-gradient(135deg, #ffffff 0%, #f5f7ff 100%);
        padding: 40px;
        border-radius: 15px;
        border: none;
        margin-bottom: 30px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    }
    
    .upload-section h3 {
        color: #1e3c72;
        margin-top: 0;
        font-weight: 700;
        font-size: 1.3rem;
    }
    
    /* File uploader styling */
    [data-testid="stFileUploadDropzone"] {
        border: 2px solid #1976d2 !important;
        border-radius: 10px !important;
        background-color: #f0f7ff !important;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 12px 30px;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(30, 60, 114, 0.2);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(30, 60, 114, 0.3);
    }
    
    /* Metric styling */
    [data-testid="metric-container"] {
        background: white !important;
        border-radius: 10px !important;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05) !important;
    }
    
    /* Dataframe styling */
    [data-testid="stDataFrame"] {
        border-radius: 10px !important;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05) !important;
    }
    
    /* Download button special styling */
    [data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #43a047 0%, #2e7d32 100%) !important;
        box-shadow: 0 4px 15px rgba(67, 160, 71, 0.2) !important;
    }
    
    [data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(67, 160, 71, 0.3) !important;
    }
    
    /* Subheader styling */
    h2 {
        color: #1e3c72 !important;
        font-weight: 700 !important;
        margin-top: 30px !important;
    }
    
    /* Divider styling */
    hr {
        margin: 30px 0 !important;
        border: 1px solid rgba(30, 60, 114, 0.1) !important;
    }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def norm_text(value):
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ").replace("\r", " ").strip().upper())

def norm_key(value):
    return re.sub(r"[^A-Z0-9]", "", norm_text(value))

def norm_party(value):
    s = norm_text(value)
    s = s.replace("&", "AND")
    s = re.sub(r"\bM\s*/\s*S\.?\b", " ", s)
    s = re.sub(r"\bPVT\.?\s*LTD\.?\b", "PRIVATE LIMITED", s)
    s = re.sub(r"\bPRIVATE\s+LIMITED\b", " ", s)
    s = re.sub(r"\bLIMITED\b", " ", s)
    return re.sub(r"[^A-Z0-9]", "", s)

def to_number(value):
    if value is None or pd.isna(value):
        return np.nan
    if isinstance(value, str):
        s = value.strip().replace(",", "").replace("₹", "")
        if not s:
            return np.nan
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
    else:
        s = value
    try:
        return float(s)
    except Exception:
        return np.nan

def to_date(value):
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, str) and not value.strip():
        return pd.NaT
    return pd.to_datetime(value, dayfirst=True, errors="coerce")

def diff_money(a, b):
    if pd.isna(a) and pd.isna(b):
        return np.nan
    a_val = 0 if pd.isna(a) else float(a)
    b_val = 0 if pd.isna(b) else float(b)
    return round(a_val - b_val, 2)

def diff_date(a, b):
    if pd.isna(a) or pd.isna(b):
        return np.nan
    return int((pd.Timestamp(a) - pd.Timestamp(b)).days)

def close(a, b, tolerance=1.00):
    if pd.isna(a) or pd.isna(b):
        return False
    return abs(float(a) - float(b)) <= tolerance

def find_column(df, names, required=False):
    mapping = {norm_key(c): c for c in df.columns}
    for name in names:
        k = norm_key(name)
        if k in mapping:
            return mapping[k]
    if required:
        raise KeyError(f"Required column not found: {names}")
    return None

# ============================================================
# HEADER DETECTION
# ============================================================

def detect_books_header(raw):
    """Multiple strategy approach for detecting Books header"""
    
    # Strategy 1: Look for Date + Particulars combination
    for i in range(min(30, len(raw))):
        vals = {norm_key(x) for x in raw.iloc[i].tolist()}
        if "DATE" in vals and "PARTICULARS" in vals:
            return i
    
    # Strategy 2: Look for Date + Party Name
    for i in range(min(30, len(raw))):
        vals = {norm_key(x) for x in raw.iloc[i].tolist()}
        if "DATE" in vals and ("PARTY" in vals or "SUPPLIER" in vals):
            return i
    
    # Strategy 3: Look for Date column alone
    for i in range(min(30, len(raw))):
        vals = {norm_key(x) for x in raw.iloc[i].tolist()}
        if "DATE" in vals and any(k in vals for k in ["INVOICE", "BILL", "AMOUNT"]):
            return i
    
    raise RuntimeError(
        "❌ Could not detect Books file header row.\n\n"
        "📋 File Format Issues:\n"
        "• Headers not in first visible row\n"
        "• Missing Date column\n"
        "• Missing Party/Particulars column\n\n"
        "✅ How to Fix:\n"
        "1. Open file in Excel\n"
        "2. Ensure first row has: Date, Particulars/Party, Invoice No, Amount\n"
        "3. Delete any empty rows above headers\n"
        "4. Save and try again"
    )

def row_tokens(raw, row_no):
    return {norm_key(x) for x in raw.iloc[row_no].tolist()}

def contains_token(tokens, options):
    return any(norm_key(x) in tokens for x in options)

def detect_gst_headers(raw):
    max_scan = min(40, len(raw))

    for i in range(max_scan - 1):
        r1 = row_tokens(raw, i)
        r2 = row_tokens(raw, i + 1)
        gstin = contains_token(r1, ["GSTIN of supplier", "GSTIN"])
        party = contains_token(r1, ["Trade/Legal name", "Trade Legal name", "Trade/Legal Name of Supplier", "Party Name"])
        invoice = contains_token(r2, ["Invoice number", "Invoice No", "Invoice Number"])
        date = contains_token(r2, ["Invoice Date", "Invoice date"])
        value = contains_token(r2, ["Invoice Value(₹)", "Invoice Value"])
        if gstin and party and invoice and date and value:
            return i, i + 1

    for i in range(max_scan):
        r = row_tokens(raw, i)
        if (contains_token(r, ["Invoice number", "Invoice No", "Invoice Number"])
            and contains_token(r, ["Invoice Date", "Invoice date"])
            and contains_token(r, ["Invoice Value(₹)", "Invoice Value"])):
            return max(0, i - 1), i

    for i in range(max_scan):
        r = row_tokens(raw, i)
        if (contains_token(r, ["GSTIN of supplier", "GSTIN"])
            and contains_token(r, ["Invoice number", "Invoice No", "Invoice Number"])
            and contains_token(r, ["Invoice Date", "Invoice date"])
            and contains_token(r, ["Invoice Value(₹)", "Invoice Value"])):
            return i, None

    return None, None

def build_gst_headers(raw, h1, h2):
    headers = []
    for c in range(raw.shape[1]):
        labels = []
        if h1 is not None and not pd.isna(raw.iloc[h1, c]):
            labels.append(str(raw.iloc[h1, c]).strip())
        if h2 is not None and not pd.isna(raw.iloc[h2, c]):
            labels.append(str(raw.iloc[h2, c]).strip())
        label = next((x for x in reversed(labels) if x), f"Unnamed_{c}")
        if label.lower() in {"nan", "none"}:
            label = f"Unnamed_{c}"
        headers.append(label)

    used = {}
    unique = []
    for h in headers:
        if h not in used:
            used[h] = 0
            unique.append(h)
        else:
            used[h] += 1
            unique.append(f"{h}_{used[h]}")
    return unique

# ============================================================
# PREPARE DATA
# ============================================================

def prepare_books(books):
    date_col = find_column(books, ["Date"], required=True)
    party_col = find_column(books, ["Particulars", "Party Name", "Supplier Name"], required=True)
    gstin_col = find_column(books, ["Sales Tax No.", "GSTIN", "GSTIN of Supplier"], required=False)
    invoice_col = find_column(books, ["Voucher Ref.", "Invoice No", "Bill No"], required=False)
    voucher_no_col = find_column(books, ["Vch No.", "Voucher No"], required=False)
    amount_col = find_column(books, ["Gross Total", "Invoice Value", "Total Amount"], required=True)
    taxable_col = find_column(books, ["Purchase", "Taxable Value"], required=False)
    igst_col = find_column(books, ["IGST"], required=False)
    cgst_col = find_column(books, ["CGST", "CGST.1"], required=False)
    sgst_col = find_column(books, ["SGST", "SGST.1"], required=False)

    books["_gstin"] = books[gstin_col].apply(norm_key) if gstin_col else ""
    books["_party"] = books[party_col].apply(norm_party)

    if invoice_col:
        inv = books[invoice_col].apply(norm_key)
    else:
        inv = pd.Series("", index=books.index)

    if voucher_no_col:
        fallback = books[voucher_no_col].apply(norm_key)
        inv = inv.where(inv != "", fallback)

    books["_invoice"] = inv
    books["_date"] = books[date_col].apply(to_date)
    books["_invoice_value"] = books[amount_col].apply(to_number)
    books["_taxable"] = books[taxable_col].apply(to_number) if taxable_col else np.nan
    books["_igst"] = books[igst_col].apply(to_number) if igst_col else np.nan
    books["_cgst"] = books[cgst_col].apply(to_number) if cgst_col else np.nan
    books["_sgst"] = books[sgst_col].apply(to_number) if sgst_col else np.nan
    books["_matched"] = False

    generic_accounts = {"PURCHASE", "SALES", "EXPENSE", "INCOME", "BANK", "CASH"}
    books = books[~books["_party"].isin(generic_accounts)].reset_index(drop=True)

    return books, {"gstin": gstin_col, "party": party_col, "invoice": invoice_col, "date": date_col, "amount": amount_col, "taxable": taxable_col, "igst": igst_col, "cgst": cgst_col, "sgst": sgst_col}

def prepare_gst(gst):
    gstin_col = find_column(gst, ["GSTIN of supplier", "GSTIN"], required=True)
    party_col = find_column(gst, ["Trade/Legal name", "Party Name", "Supplier Name"], required=True)
    invoice_col = find_column(gst, ["Invoice number", "Invoice No"], required=True)
    date_col = find_column(gst, ["Invoice Date"], required=True)
    amount_col = find_column(gst, ["Invoice Value(₹)", "Invoice Value"], required=True)
    taxable_col = find_column(gst, ["Taxable Value (₹)", "Taxable Value"], required=True)
    igst_col = find_column(gst, ["Integrated Tax(₹)", "IGST"], required=False)
    cgst_col = find_column(gst, ["Central Tax(₹)", "CGST"], required=False)
    sgst_col = find_column(gst, ["State/UT Tax(₹)", "SGST"], required=False)

    gst["_gstin"] = gst[gstin_col].apply(norm_key)
    gst["_party"] = gst[party_col].apply(norm_party)
    gst["_invoice"] = gst[invoice_col].apply(norm_key)
    gst["_date"] = gst[date_col].apply(to_date)
    gst["_invoice_value"] = gst[amount_col].apply(to_number)
    gst["_taxable"] = gst[taxable_col].apply(to_number)
    gst["_igst"] = gst[igst_col].apply(to_number) if igst_col else np.nan
    gst["_cgst"] = gst[cgst_col].apply(to_number) if cgst_col else np.nan
    gst["_sgst"] = gst[sgst_col].apply(to_number) if sgst_col else np.nan
    gst["_matched"] = False

    invoice_col_name = find_column(gst, ["Invoice number", "Invoice No"], required=True)
    mask = gst[invoice_col_name].notna() & (gst[invoice_col_name].astype(str).str.strip() != "")
    gst = gst.loc[mask].reset_index(drop=True)

    return gst, {"gstin": gstin_col, "party": party_col, "invoice": invoice_col, "date": date_col, "amount": amount_col, "taxable": taxable_col, "igst": igst_col, "cgst": cgst_col, "sgst": sgst_col}

# ============================================================
# MATCHING
# ============================================================

def candidate_score(book, gst):
    score = 0
    same_gstin = bool(book["_gstin"] and gst["_gstin"] and book["_gstin"] == gst["_gstin"])
    same_invoice = bool(book["_invoice"] and gst["_invoice"] and book["_invoice"] == gst["_invoice"])
    same_party = bool(book["_party"] and gst["_party"] and book["_party"] == gst["_party"])
    same_amount = close(book["_invoice_value"], gst["_invoice_value"])
    same_taxable = close(book["_taxable"], gst["_taxable"])
    same_date = (pd.notna(book["_date"]) and pd.notna(gst["_date"]) and pd.Timestamp(book["_date"]) == pd.Timestamp(gst["_date"]))

    if same_gstin:
        score += 1000
    if same_invoice:
        score += 1000
    if same_party:
        score += 300
    if same_amount:
        score += 150
    if same_taxable:
        score += 50
    if same_date:
        score += 40

    eligible = ((same_gstin and same_invoice) or (same_invoice and same_party) or (same_gstin and same_party and same_amount) or (same_party and same_invoice and same_amount) or (same_party and same_amount and same_date))
    return score, eligible

def find_best_gst_match(book, gst):
    available = gst[gst["_matched"] == False]
    candidates = []
    for idx, grow in available.iterrows():
        score, eligible = candidate_score(book, grow)
        if eligible:
            candidates.append((idx, score))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[1], reverse=True)
    if len(candidates) > 1 and candidates[0][1] == candidates[1][1]:
        top_score = candidates[0][1]
        tied = [x for x in candidates if x[1] == top_score]
        if len(tied) > 1:
            return None
    return candidates[0][0]

def mismatch_flags(ad, td, cd, sd, idf, dd):
    return {
        "Invoice value mismatch": pd.notna(ad) and abs(ad) > 1.00,
        "Invoice date mismatch": pd.notna(dd) and abs(int(dd)) > 0.0,
        "Taxable Value mismatch": pd.notna(td) and abs(td) > 1.00,
        "CGST mismatch": pd.notna(cd) and abs(cd) > 1.00,
        "SGST mismatch": pd.notna(sd) and abs(sd) > 1.00,
        "IGST mismatch": pd.notna(idf) and abs(idf) > 1.00,
    }

def calculate_status(flags):
    if flags["Invoice value mismatch"]:
        return "AMOUNT MISMATCH"
    if flags["Invoice date mismatch"]:
        return "DATE MISMATCH"
    if any(flags[x] for x in ["Taxable Value mismatch", "CGST mismatch", "SGST mismatch", "IGST mismatch"]):
        return "GST MISMATCH"
    return "MATCHED"

def create_remarks(status, flags):
    if status in ["NOT FOUND IN GST", "NOT FOUND IN BOOKS", "MATCHED"]:
        return status
    
    mismatches = []
    if flags.get("Invoice value mismatch"):
        mismatches.append("Amount Mismatch")
    if flags.get("Invoice date mismatch"):
        mismatches.append("Date Mismatch")
    if flags.get("Taxable Value mismatch"):
        mismatches.append("Taxable Value Mismatch")
    if flags.get("CGST mismatch"):
        mismatches.append("CGST Mismatch")
    if flags.get("SGST mismatch"):
        mismatches.append("SGST Mismatch")
    if flags.get("IGST mismatch"):
        mismatches.append("IGST Mismatch")
    
    return "; ".join(mismatches) if mismatches else status

def make_row(gstin, party, invoice, bd, gd, biv, giv, tb, tg, cb, cg, sb, sg, ib, ig, remarks):
    return {
        "GSTIN": gstin, "Party Name": party, "Invoice No": invoice,
        "Books Date": bd, "GST Date": gd, "Date Difference": diff_date(bd, gd),
        "Books Invoice Value": biv, "GST Invoice Value": giv, "Invoice Value Difference": diff_money(biv, giv),
        "Books Taxable Value": tb, "GST Taxable Value": tg, "Taxable Value Difference": diff_money(tb, tg),
        "Books CGST": cb, "GST CGST": cg, "CGST Difference": diff_money(cb, cg),
        "Books SGST": sb, "GST SGST": sg, "SGST Difference": diff_money(sb, sg),
        "Books IGST": ib, "GST IGST": ig, "IGST Difference": diff_money(ib, ig),
        "Remarks": remarks,
    }

def reconcile(books, gst, books_cols, gst_cols):
    rows = []
    
    for bidx, b in books.iterrows():
        gidx = find_best_gst_match(b, gst)
        
        if gidx is None:
            party_name = b[books_cols["party"]]
            if pd.isna(party_name) or not str(party_name).strip():
                party_name = b["_gstin"]
            rows.append(make_row(b["_gstin"], party_name, b["_invoice"], b["_date"], pd.NaT, b["_invoice_value"], np.nan, b["_taxable"], np.nan, b["_cgst"], np.nan, b["_sgst"], np.nan, b["_igst"], np.nan, "NOT FOUND IN GST"))
            continue

        g = gst.loc[gidx]
        books.loc[bidx, "_matched"] = True
        gst.loc[gidx, "_matched"] = True

        ad = diff_money(b["_invoice_value"], g["_invoice_value"])
        td = diff_money(b["_taxable"], g["_taxable"])
        cd = diff_money(b["_cgst"], g["_cgst"])
        sd = diff_money(b["_sgst"], g["_sgst"])
        idf = diff_money(b["_igst"], g["_igst"])
        dd = diff_date(b["_date"], g["_date"])

        flags = mismatch_flags(ad, td, cd, sd, idf, dd)
        status = calculate_status(flags)
        remarks = create_remarks(status, flags)

        gstin = g["_gstin"] or b["_gstin"]
        party_raw = g[gst_cols["party"]]
        if pd.notna(party_raw) and str(party_raw).strip():
            party = party_raw
        else:
            party = b[books_cols["party"]]
        if pd.isna(party) or not str(party).strip():
            party = gstin

        invoice_raw = g[gst_cols["invoice"]]
        invoice = invoice_raw if (pd.notna(invoice_raw) and str(invoice_raw).strip()) else b["_invoice"]

        rows.append(make_row(gstin, party, invoice, b["_date"], g["_date"], b["_invoice_value"], g["_invoice_value"], b["_taxable"], g["_taxable"], b["_cgst"], g["_cgst"], b["_sgst"], g["_sgst"], b["_igst"], g["_igst"], remarks))

    for _, g in gst.iterrows():
        if bool(g["_matched"]):
            continue
        party_name = g[gst_cols["party"]]
        if pd.isna(party_name) or not str(party_name).strip():
            party_name = g["_gstin"]
        rows.append(make_row(g["_gstin"], party_name, g[gst_cols["invoice"]], pd.NaT, g["_date"], np.nan, g["_invoice_value"], np.nan, g["_taxable"], np.nan, g["_cgst"], np.nan, g["_sgst"], np.nan, g["_igst"], "NOT FOUND IN BOOKS"))

    return pd.DataFrame(rows, columns=["GSTIN", "Party Name", "Invoice No", "Books Date", "GST Date", "Date Difference", "Books Invoice Value", "GST Invoice Value", "Invoice Value Difference", "Books Taxable Value", "GST Taxable Value", "Taxable Value Difference", "Books CGST", "GST CGST", "CGST Difference", "Books SGST", "GST SGST", "SGST Difference", "Books IGST", "GST IGST", "IGST Difference", "Remarks"])

def make_summary(result):
    totals = {
        "Total Invoices (Books)": result["Invoice No"].notna().sum() - (result["Books Date"].isna().sum()),
        "Total Invoices (GST)": result["GST Date"].notna().sum(),
        "Matched Invoices": (result["Remarks"] == "MATCHED").sum(),
        "Amount Mismatch": result["Remarks"].str.contains("Amount Mismatch", na=False).sum(),
        "Taxable Value Mismatch": result["Remarks"].str.contains("Taxable Value Mismatch", na=False).sum(),
        "CGST Mismatch": result["Remarks"].str.contains("CGST Mismatch", na=False).sum(),
        "SGST Mismatch": result["Remarks"].str.contains("SGST Mismatch", na=False).sum(),
        "IGST Mismatch": result["Remarks"].str.contains("IGST Mismatch", na=False).sum(),
        "Date Mismatch": result["Remarks"].str.contains("Date Mismatch", na=False).sum(),
        "Not in GST": (result["Remarks"] == "NOT FOUND IN GST").sum(),
        "Not in Books": (result["Remarks"] == "NOT FOUND IN BOOKS").sum(),
    }
    return pd.DataFrame(list(totals.items()), columns=["Metric", "Count"])

# ============================================================
# EXCEL GENERATION
# ============================================================

def create_excel_report(result, summary):
    numeric_cols = ["Books Invoice Value", "GST Invoice Value", "Invoice Value Difference", "Books Taxable Value", "GST Taxable Value", "Taxable Value Difference", "Books CGST", "GST CGST", "CGST Difference", "Books SGST", "GST SGST", "SGST Difference", "Books IGST", "GST IGST", "IGST Difference", "Date Difference"]
    
    for col in numeric_cols:
        result[col] = result[col].fillna(0)

    wb = Workbook()
    summary_ws = wb.active
    summary_ws.title = "Summary"
    ws = wb.create_sheet("Reconciliation")

    # Summary Sheet
    summary_ws["A1"] = "GST RECONCILIATION SUMMARY"
    summary_ws.merge_cells("A1:B1")
    summary_ws.row_dimensions[2].height = 5
    summary_ws["A3"] = "Metric"
    summary_ws["B3"] = "Count"

    for r, row in enumerate(summary.itertuples(index=False, name=None), start=4):
        summary_ws.cell(r, 1, row[0])
        summary_ws.cell(r, 2, row[1])

    # Styling
    header_fill = PatternFill("solid", fgColor="1F4E78")
    sub_fill = PatternFill("solid", fgColor="5B9BD5")
    white_bold = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    summary_ws["A1"].fill = header_fill
    summary_ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    summary_ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    for c in summary_ws[3]:
        if c.column <= 2:
            c.fill = sub_fill
            c.font = white_bold
            c.border = border
            c.alignment = Alignment(horizontal="center")

    summary_rows = len(summary) + 4
    for r in range(4, summary_rows):
        cell_metric = summary_ws.cell(r, 1)
        cell_metric.border = border
        cell_metric.alignment = Alignment(horizontal="left", vertical="center")
        cell_metric.font = Font(size=11)
        
        cell_count = summary_ws.cell(r, 2)
        cell_count.border = border
        cell_count.alignment = Alignment(horizontal="center", vertical="center")
        cell_count.number_format = "0"
        cell_count.font = Font(size=11)
        
        if (r - 4) % 2 == 1:
            alt_fill = PatternFill("solid", fgColor="F2F2F2")
            cell_metric.fill = alt_fill
            cell_count.fill = alt_fill

    summary_ws.column_dimensions["A"].width = 30
    summary_ws.column_dimensions["B"].width = 15
    summary_ws.row_dimensions[1].height = 28
    summary_ws.row_dimensions[3].height = 20
    for r in range(4, summary_rows):
        summary_ws.row_dimensions[r].height = 22
    summary_ws.freeze_panes = "A4"

    # Reconciliation Sheet
    groups = [("GSTIN", 1, 1), ("Party Name", 2, 2), ("Invoice No", 3, 3), ("Invoice date", 4, 6), ("Invoice Value", 7, 9), ("Taxable Value", 10, 12), ("CGST", 13, 15), ("SGST", 16, 18), ("IGST", 19, 21), ("Remarks", 22, 22)]

    for title, start, end in groups:
        ws.cell(1, start, title)
        if end > start:
            ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
        else:
            ws.merge_cells(start_row=1, start_column=start, end_row=2, end_column=end)

    subheaders = {4: "As per Books", 5: "As per GSTR - 2B", 6: "Difference in days", 7: "As per Books", 8: "As per GSTR - 2B", 9: "Difference", 10: "As per Books", 11: "As per GSTR - 2B", 12: "Difference", 13: "As per Books", 14: "As per GSTR - 2B", 15: "Difference", 16: "As per Books", 17: "As per GSTR - 2B", 18: "Difference", 19: "As per Books", 20: "As per GSTR - 2B", 21: "Difference"}

    for col, value in subheaders.items():
        ws.cell(2, col, value)

    money_cols = set(range(7, 22))
    for r, values in enumerate(result.itertuples(index=False, name=None), start=3):
        for c, value in enumerate(values, start=1):
            cell = ws.cell(r, c, value)
            if c in (4, 5) and pd.notna(value):
                cell.number_format = "dd-mm-yyyy"
            if c in money_cols and pd.notna(value):
                cell.number_format = "#,##0.00"

    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=22):
        for cell in row:
            cell.fill = header_fill if cell.row == 1 else sub_fill
            cell.font = white_bold
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in ws.iter_rows(min_row=3, max_row=max(3, ws.max_row), min_col=1, max_col=22):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    formulas = [
        ('ISNUMBER(SEARCH("Amount Mismatch",$V3))', "F4CCCC"),
        ('ISNUMBER(SEARCH("Taxable Value Mismatch",$V3))', "D9EAD3"),
        ('ISNUMBER(SEARCH("CGST Mismatch",$V3))', "D9D2E9"),
        ('ISNUMBER(SEARCH("SGST Mismatch",$V3))', "CFE2F3"),
        ('ISNUMBER(SEARCH("IGST Mismatch",$V3))', "FCE5CD"),
        ('ISNUMBER(SEARCH("Date Mismatch",$V3))', "FFF2CC"),
        ('$V3="NOT FOUND IN GST"', "D0E0E3"),
        ('$V3="NOT FOUND IN BOOKS"', "EAD1DC"),
    ]

    end_row = max(3, ws.max_row)
    for formula, color in formulas:
        ws.conditional_formatting.add(f"A3:V{end_row}", FormulaRule(formula=[formula], fill=PatternFill("solid", fgColor=color)))

    widths = {1: 20, 2: 35, 3: 22, 4: 14, 5: 16, 6: 18, 7: 16, 8: 18, 9: 16, 10: 16, 11: 18, 12: 16, 13: 14, 14: 18, 15: 14, 16: 14, 17: 18, 18: 14, 19: 14, 20: 18, 21: 14, 22: 45}

    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 42
    for r in range(3, ws.max_row + 1):
        ws.row_dimensions[r].height = 30

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:V{max(2, ws.max_row)}"

    return wb

# ============================================================
# STREAMLIT UI
# ============================================================

st.markdown("""
    <div class="header-box">
        <h1 style="margin-top: 0;">GST Reconciliation Tool</h1>
        <p style="margin-bottom: 0;">Automated GSTR-2B reconciliation system</p>
    </div>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="info-box">
        <strong>ℹ️ About This Tool:</strong><br>
        An automated solution that compares accounting records with GSTR-2B data to identify discrepancies and provide detailed insights into mismatches, tax differences and unreconciled invoices.
    </div>
""", unsafe_allow_html=True)

# Initialize session state
if "files_uploaded" not in st.session_state:
    st.session_state.files_uploaded = False
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "summary_df" not in st.session_state:
    st.session_state.summary_df = None

# File Upload Section
st.markdown("""
    <div class="upload-section">
        <h3>📁 Upload data</h3>
    </div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("<h4 style='color: #1e3c72; margin-bottom: 10px;'>📊 Data as per accounting records</h4>", unsafe_allow_html=True)
    books_file = st.file_uploader(
        "Select Books File (Excel format)",
        type=["xls", "xlsx"],
        key="books_upload",
        help="Upload Books.xls or Books.xlsx"
    )
    if books_file:
        st.success(f"✅ {books_file.name} uploaded ({books_file.size/1024:.1f} KB)")

with col2:
    st.markdown("<h4 style='color: #1e3c72; margin-bottom: 10px;'>📊 Data as per GSTR-2B</h4>", unsafe_allow_html=True)
    gst_file = st.file_uploader(
        "Select GSTR-2B File (Excel format)",
        type=["xlsx"],
        key="gst_upload",
        help="Upload Conso 2B.xlsx or GSTR-2B export"
    )
    if gst_file:
        st.success(f"✅ {gst_file.name} uploaded ({gst_file.size/1024:.1f} KB)")

# Check if files are uploaded
if books_file is None or gst_file is None:
    st.markdown("""
        <div class="error-box">
            <strong style="font-size: 1.1rem;">❌ Upload files first</strong><br>
            <p style="margin: 10px 0 0 0;">Please upload both <strong>Books File</strong> and <strong>GST B2B File</strong> to generate the report.</p>
        </div>
    """, unsafe_allow_html=True)
    st.markdown("""
        <div class="info-box">
            <strong>📌 Next Steps:</strong><br>
            1️⃣ Click the upload box for Books file (Books.xls or Books.xlsx)<br>
            2️⃣ Click the upload box for GST file (Conso 2B.xlsx)<br>
            3️⃣ Once both files are uploaded, click the Generate Report button
        </div>
    """, unsafe_allow_html=True)
    st.stop()

# Generate Report Button
st.markdown("<br>", unsafe_allow_html=True)
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    generate_report = st.button("🚀 Generate Report", key="generate_btn", use_container_width=True)

if generate_report:
    try:
        with st.spinner("⏳ Processing files... (Step 1/4)"):
            # Read Books
            books_raw = pd.read_excel(books_file, sheet_name=0, header=None)
            h = detect_books_header(books_raw)
            headers = [str(x).strip() if not pd.isna(x) else f"Unnamed_{i}" for i, x in enumerate(books_raw.iloc[h].tolist())]
            books = books_raw.iloc[h + 1:].copy()
            books.columns = headers
            books = books.dropna(how="all").reset_index(drop=True)
            st.success("✅ Books file loaded")

        with st.spinner("⏳ Processing GST file... (Step 2/4)"):
            # Read GST
            gst_raw = pd.read_excel(gst_file, sheet_name=0, header=None)
            h1, h2 = detect_gst_headers(gst_raw)
            if h1 is None:
                raise RuntimeError(
                    "❌ Could not detect GSTR-2B file format.\n\n"
                    "📋 File Format Issues:\n"
                    "• Not a standard GSTR-2B export from GST Portal\n"
                    "• File has been manually edited\n"
                    "• Missing standard columns (GSTIN, Invoice No, Date, Amount)\n\n"
                    "✅ How to Fix:\n"
                    "1. Download fresh GSTR-2B from: www.gstreturn.gov.in\n"
                    "2. Login → RETURNS → GSTR-2B → Download B2B\n"
                    "3. Do NOT edit the file\n"
                    "4. Use file as-is from portal"
                )
            headers = build_gst_headers(gst_raw, h1, h2)
            data_start = (h2 + 1) if h2 is not None else (h1 + 1)
            gst = gst_raw.iloc[data_start:].copy()
            gst.columns = headers
            gst = gst.dropna(how="all").reset_index(drop=True)
            st.success("✅ GST file loaded")

        with st.spinner("⏳ Preparing data... (Step 3/4)"):
            books, books_cols = prepare_books(books)
            gst, gst_cols = prepare_gst(gst)
            st.success("✅ Data prepared")

        with st.spinner("⏳ Running reconciliation... (Step 4/4)"):
            result = reconcile(books, gst, books_cols, gst_cols)
            summary = make_summary(result)
            st.session_state.result_df = result
            st.session_state.summary_df = summary
            st.session_state.files_uploaded = True
            st.success("✅ Reconciliation complete!")

    except KeyError as e:
        st.markdown(f"""
            <div class="error-box">
                <strong>❌ Column Not Found Error</strong><br>
                {str(e)}<br><br>
                <strong>How to Fix:</strong><br>
                • Ensure your Books file has columns: Date, Particulars, Invoice No<br>
                • Ensure GST file is standard GSTR-2B export from portal<br>
                • Check that headers are in the first visible row
            </div>
        """, unsafe_allow_html=True)
    except RuntimeError as e:
        st.markdown(f"""
            <div class="error-box">
                <strong>❌ File Format Error</strong><br>
                {str(e)}<br><br>
                <strong>How to Fix:</strong><br>
                • Books file must have: Date, Particulars, Invoice columns<br>
                • GST file must be standard GSTR-2B export<br>
                • File should not be corrupted
            </div>
        """, unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f"""
            <div class="error-box">
                <strong>❌ Unexpected Error</strong><br>
                {str(e)}<br><br>
                <strong>How to Fix:</strong><br>
                • Verify file format is correct<br>
                • Check file is not corrupted<br>
                • Close the file if open in Excel
            </div>
        """, unsafe_allow_html=True)

# Display Results
if st.session_state.files_uploaded and st.session_state.result_df is not None:
    st.markdown("""
        <div class="success-box">
            <strong style="font-size: 1.1rem;">✅ Reconciliation Successful!</strong><br>
            <p style="margin: 8px 0 0 0;">Review the summary and detailed results below.</p>
        </div>
    """, unsafe_allow_html=True)

    # Summary Metrics
    st.subheader("📈 Summary Statistics")
    summary = st.session_state.summary_df
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        matched = int(summary[summary["Metric"] == "Matched Invoices"]["Count"].values[0])
        st.metric("✅ Matched", matched)
    with col2:
        amount = int(summary[summary["Metric"] == "Amount Mismatch"]["Count"].values[0])
        st.metric("💰 Amount", amount)
    with col3:
        taxable = int(summary[summary["Metric"] == "Taxable Value Mismatch"]["Count"].values[0])
        st.metric("📋 Taxable", taxable)
    with col4:
        not_gst = int(summary[summary["Metric"] == "Not in GST"]["Count"].values[0])
        st.metric("🚫 Not in GST", not_gst)
    with col5:
        not_books = int(summary[summary["Metric"] == "Not in Books"]["Count"].values[0])
        st.metric("🚫 Not in Books", not_books)

    # Detailed Summary Table
    st.subheader("📋 Detailed Summary")
    st.dataframe(summary, use_container_width=True, hide_index=True)

    # Reconciliation Data
    st.subheader("🔍 Reconciliation Details")
    st.dataframe(st.session_state.result_df, use_container_width=True, hide_index=True)

    # Download Report
    st.divider()
    st.subheader("⬇️ Download Report")
    
    with st.spinner("📦 Generating Excel file..."):
        wb = create_excel_report(st.session_state.result_df, st.session_state.summary_df)
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Download Reconciliation Report (Excel)",
            data=output,
            file_name="GST_Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.markdown("""
        <div class="success-box">
            <strong style="font-size: 1.1rem;">✨ Report Ready!</strong><br>
            <p style="margin: 8px 0 0 0;">Click the button above to download your Excel report with professional formatting and color-coded mismatches.</p>
        </div>
    """, unsafe_allow_html=True)