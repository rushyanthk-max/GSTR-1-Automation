import streamlit as st
import pandas as pd
import re, io

st.set_page_config(
    page_title="BCPL GST Sanitizer & Auditor",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# HELPERS
# =============================================================================

def clean_df_columns(df):
    """Normalize headers; auto-detect header row when HSN/SKU columns missing."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)
    HSN_KW = {"hsn", "sac", "commodity", "nomenclature"}
    SKU_KW = {"sku", "fsn", "seller-sku", "item-code", "product-id", "article"}
    has_hsn = any(any(k in str(c).lower() for k in HSN_KW) for c in df.columns)
    has_sku = any(any(k in str(c).lower() for k in SKU_KW) for c in df.columns)
    if not (has_hsn or has_sku):
        TRIGGERS = {"hsn", "hsn code", "hsn/sac", "sku", "seller sku", "fsn"}
        for idx, row in df.iterrows():
            vals = {str(v).strip().lower() for v in row.values if pd.notna(v)}
            if vals & TRIGGERS:
                df.columns = [str(v).strip() for v in row.values]
                df = df.iloc[idx + 1:].reset_index(drop=True)
                break
    return df


def detect_columns(df):
    """
    Scan a DataFrame and return dict of key column names:
    sku, hsn, cgst_rate, sgst_rate, igst_rate, order_id, order_item_id, tx_type
    First match wins; None for any field not found.
    """
    c = {k: None for k in ("sku", "hsn", "cgst_rate", "sgst_rate",
                             "igst_rate", "order_id", "order_item_id", "tx_type")}
    SKIP = {"tcs", "shipping", "gift", "wrap", "delivery", "postage",
            "cst", "vat", "cess", "tds", "amount", "amt", "value"}
    for col in df.columns:
        cl = str(col).strip().lower()
        if cl in {"order id", "order_id"}       and not c["order_id"]:      c["order_id"]      = col
        if cl in {"order item id", "order_item_id"} and not c["order_item_id"]: c["order_item_id"] = col
        if not c["sku"]:
            if cl in {"sku", "seller-sku", "seller sku", "item-code", "article-code",
                      "wms_code", "fsn", "asin"}:
                c["sku"] = col
            elif any(k in cl for k in ("sku", "fsn", "product-id", "article", "item code")):
                c["sku"] = col
        if not c["hsn"]:
            if cl in {"hsn", "hsn/sac", "hsn_sac", "hsncode", "hsn code", "hsn_code",
                      "commodity", "hsn sac", "sac code", "hsn no"}:
                c["hsn"] = col
            elif any(k in cl for k in ("hsn", "sac", "commodity", "nomenclature", "tariff")):
                c["hsn"] = col
        if not c["tx_type"] and any(k in cl for k in (
            "transaction type", "transaction_type", "order status",
            "order_status", "event type", "event_type", "document type")):
            c["tx_type"] = col
        if any(x in cl for x in SKIP):
            continue
        if "cgst"  in cl and "rate" in cl and not c["cgst_rate"]:  c["cgst_rate"]  = col
        if ("sgst" in cl or "utgst" in cl) and "rate" in cl and not c["sgst_rate"]: c["sgst_rate"] = col
        if "igst"  in cl and "rate" in cl and not c["igst_rate"]:  c["igst_rate"]  = col
    return c


def deep_clean_sku(val):
    """Return lowercase alphanumeric-only SKU key for lookups."""
    if pd.isna(val): return ""
    s = str(val).strip().lower()
    s = re.sub(r'^[`"\'\s]+|[`"\'\s]+$', "", s)
    if s.startswith("sku:"): s = s[4:]
    elif s.startswith("sku"): s = s[3:]
    result = re.sub(r"[^a-z0-9]", "", s)
    return "" if result in ("nan", "none", "na") else result   # guard against literal 'nan'


def extract_rate(val):
    """
    Parse a raw GST rate cell → float percentage.
    Handles '18%', '18', '0.18', '0.018' (3-decimal Amazon format), etc.
    """
    if pd.isna(val) or str(val).strip() in {"", "nan", "None", "<NA>"}: return 0.0
    s = str(val).strip().replace("%", "").strip()
    s = re.sub(r"\.0+$", "", s)
    # Platforms like Amazon export 18% GST as 0.018
    EXACT = {"0.028": 28, "0.018": 18, "0.012": 12, "0.005": 5, "0.003": 3,
             "0.28":  28, "0.18":  18, "0.12":  12, "0.05":  5, "0.03":  3}
    if s in EXACT: return float(EXACT[s])
    digits = "".join(ch for ch in s if ch.isdigit() or ch == ".")
    try:
        n = float(digits) if digits else 0.0
        if 0 < n <= 1.0: n *= 100        # 0.18 → 18
        return n
    except Exception: return 0.0


def rate_to_str(r):
    """Convert a float rate to a clean string: 18.0 → '18', 2.5 → '2.5'."""
    try:
        f = float(r)
        return str(int(f)) if f == int(f) else str(f)
    except Exception: return str(r)


def normalize_hsn(val):
    """
    Clean any HSN representation → digit-only string.
    ✅ Strips Excel formula prefix =".."
    ✅ Removes spaces, hyphens, dots
    ✅ Pads 7-digit HSN → 8-digit (prepend 0)
    """
    if pd.isna(val) or str(val).strip() in {"", "nan", "None"}: return ""
    h = str(val).strip()
    if h.startswith('="') and h.endswith('"'): h = h[2:-1]   # strip =".." Excel format
    digits = "".join(filter(str.isdigit, re.sub(r"[\s\-\.\/]", "", h)))
    if len(digits) == 7: digits = "0" + digits                # ✅ pad 7-digit to 8
    return digits


def excel_hsn(h):
    """Format HSN for Excel output — preserves leading zeros via formula."""
    return f'="{h}"' if h and h != "MISSING HSN" else "MISSING HSN"


def majority_vote(lst):
    """Return the most common element in a non-empty list."""
    return max(set(lst), key=lst.count)


# =============================================================================
# UI — HEADER & UPLOADERS
# =============================================================================
st.title("📦 BCPL Universal E-commerce GST Sanitizer & Auditor")
st.caption("Upload files to clean multi-sheet workbooks and generate a side-by-side Audit Error Report.")

up1, up2 = st.columns(2)
with up1:
    st.subheader("1️⃣ Raw Transaction Report")
    uploaded_file = st.file_uploader("Sales sheet workbook (.xlsx / .xls / .csv)",
                                      type=["xlsx","xls","csv"], key="sales")
with up2:
    st.subheader("2️⃣ Master Product Catalog (Optional)")
    attribute_file = st.file_uploader("Item catalog — enables SKU/HSN/Tax audits & auto-healing",
                                       type=["xlsx","xls","csv"], key="attr")

if not uploaded_file:
    st.info("👆 Upload a sales / transaction report above to get started.")
    st.stop()

# =============================================================================
# STEP 1 — LOAD TRANSACTION SHEETS
# =============================================================================
pb = st.progress(0, text="📂 Loading transaction sheets…")
raw_sheets = {}
try:
    if uploaded_file.name.lower().endswith((".xlsx", ".xls")):
        uploaded_file.seek(0)
        xf = pd.ExcelFile(uploaded_file)           # parse once → use xf.parse() per sheet
        for sname in xf.sheet_names:
            raw_sheets[sname] = clean_df_columns(xf.parse(sname, dtype=str))
    else:
        uploaded_file.seek(0)
        raw_sheets["Sales Report"] = clean_df_columns(
            pd.read_csv(uploaded_file, dtype=str, low_memory=False))
except Exception as e:
    st.error(f"❌ Failed to load file: {e}")
    st.stop()

pb.progress(15, text="📋 Loading master product catalog…")

# =============================================================================
# STEP 2 — MASTER CATALOG
#   master_sku_hsn  : clean_sku  → hsn_digits
#   master_sku_tax  : clean_sku  → rate_str
#   master_hsn_tax  : hsn_digits → rate_str
# =============================================================================
master_sku_hsn = {}
master_sku_tax = {}
master_hsn_tax = {}
catalog_sku_col = catalog_hsn_col = None

if attribute_file:
    try:
        if attribute_file.name.lower().endswith((".xlsx", ".xls")):
            attribute_file.seek(0)
            xf2 = pd.ExcelFile(attribute_file)
            attr_df = clean_df_columns(xf2.parse(xf2.sheet_names[0], dtype=str))
        else:
            attribute_file.seek(0)
            attr_df = clean_df_columns(
                pd.read_csv(attribute_file, dtype=str, low_memory=False))

        a_sku = a_hsn = a_tax = None
        for col in attr_df.columns:
            cl = str(col).strip().lower()
            if not a_sku:
                if cl in {"sku", "seller-sku", "seller sku", "item-code", "item code",
                          "article-code", "product-sku", "fsn", "asin"}:
                    a_sku = col
                elif any(k in cl for k in ("sku", "fsn", "product id", "article")):
                    a_sku = col
            if not a_hsn:
                if cl in {"hsn", "hsn/sac", "hsn_sac", "hsncode", "hsn code",
                          "hsn_code", "commodity", "producttaxcode"}:
                    a_hsn = col
                elif any(k in cl for k in ("hsn", "sac", "taxcode", "commodity", "nomenclature")):
                    a_hsn = col
            if not a_tax and any(k in cl for k in (
                "producttaxrule", "tax rule", "gst rate", "tax percentage",
                "tax_rule", "tax rate", "rate")):
                a_tax = col

        catalog_sku_col, catalog_hsn_col = a_sku, a_hsn

        if a_sku and a_hsn:
            for _, row in attr_df.iterrows():
                r_hsn = normalize_hsn(row[a_hsn]) if pd.notna(row[a_hsn]) else ""
                r_sku = deep_clean_sku(row[a_sku])
                # ✅ FIXED: use extract_rate() not raw digit-filter (avoids "018" for "0.18")
                r_tax_num = extract_rate(row[a_tax]) if a_tax and pd.notna(row[a_tax]) else 0.0
                r_tax = rate_to_str(r_tax_num) if r_tax_num > 0 else ""
                if r_hsn and r_sku:
                    master_sku_hsn[r_sku] = r_hsn
                    if r_tax:
                        master_sku_tax[r_sku] = r_tax
                        master_hsn_tax[r_hsn] = r_tax
    except Exception as e:
        st.warning(f"⚠️ Could not fully load catalog: {e}")

pb.progress(28, text="🔗 Building cross-sheet order reference…")

# =============================================================================
# STEP 3 — CROSS-SHEET ORDER REFERENCE  (fill SKU/HSN gaps from other tabs)
# =============================================================================
order_ref = {}    # (order_id, order_item_id) → (sku_str, raw_hsn_str)
for sname, df_s in raw_sheets.items():
    c = detect_columns(df_s)
    if c["order_id"] and c["order_item_id"] and c["sku"]:
        for _, row in df_s.iterrows():
            oid  = str(row[c["order_id"]]).strip()      if pd.notna(row[c["order_id"]])      else ""
            oiid = str(row[c["order_item_id"]]).strip() if pd.notna(row[c["order_item_id"]]) else ""
            sku  = str(row[c["sku"]]).strip()  if c["sku"]  and pd.notna(row[c["sku"]]) else ""
            hsn  = str(row[c["hsn"]]).strip()  if c["hsn"]  and pd.notna(row[c["hsn"]]) else ""
            if oid and oiid and (oid, oiid) not in order_ref:
                order_ref[(oid, oiid)] = (sku, hsn)

pb.progress(40, text="🕵️ Building records & healing HSNs…")

# =============================================================================
# STEP 4 — BUILD RECORDS
#   ✅ HSN from master fills missing codes right here (healed_hsn field)
#   ✅ 7-digit HSN gets padded to 8 via normalize_hsn()
#   ✅ igst_type tracked per row for correct IGST vs CGST+SGST split later
# =============================================================================
records = []
col_maps = []

for sname, df_s in raw_sheets.items():
    c = detect_columns(df_s)
    col_maps.append((sname, c))

    for pos, (df_idx, row) in enumerate(df_s.iterrows()):
        rhsn = str(row[c["hsn"]]).strip() if c["hsn"] and pd.notna(row[c["hsn"]]) else ""
        rsku = str(row[c["sku"]]).strip() if c["sku"] and pd.notna(row[c["sku"]]) else ""

        # Fill from cross-sheet order reference
        if c["order_id"] and c["order_item_id"] and (not rhsn or not rsku):
            oid  = str(row[c["order_id"]]).strip()      if pd.notna(row[c["order_id"]])      else ""
            oiid = str(row[c["order_item_id"]]).strip() if pd.notna(row[c["order_item_id"]]) else ""
            ref  = order_ref.get((oid, oiid), ("", ""))
            if not rsku: rsku = ref[0]
            if not rhsn: rhsn = ref[1]

        # ✅ Normalize HSN: strip junk, pad 7-digit → 8-digit
        raw_hsn   = normalize_hsn(rhsn)
        clean_sku = deep_clean_sku(rsku)
        sku_disp  = re.sub(r'^["\'`\s]+|["\'`\s]+$', "", rsku)
        if sku_disp.upper().startswith("SKU:"): sku_disp = sku_disp[4:]

        # ✅ HEAL: fill missing HSN from master catalog using SKU lookup
        healed_hsn = raw_hsn
        healed_from_master = False
        if not healed_hsn and clean_sku and clean_sku in master_sku_hsn:
            healed_hsn = master_sku_hsn[clean_sku]
            healed_from_master = True

        # Tax rates
        cgst = extract_rate(row[c["cgst_rate"]]) if c["cgst_rate"] else 0.0
        sgst = extract_rate(row[c["sgst_rate"]]) if c["sgst_rate"] else 0.0
        igst = extract_rate(row[c["igst_rate"]]) if c["igst_rate"] else 0.0
        total = igst if igst > 0 else (cgst + sgst)
        rate_str = rate_to_str(total)

        # Track whether this row uses IGST (inter-state) or CGST+SGST (intra-state)
        if igst > 0: igst_type = True
        elif cgst > 0 or sgst > 0: igst_type = False
        else: igst_type = None   # unknown / zero-rate row

        tx_status = str(row[c["tx_type"]]).strip().lower() if c["tx_type"] else ""

        records.append({
            "sheet":              sname,
            "pos":                pos,
            "sku_display":        sku_disp,
            "clean_sku":          clean_sku,
            "raw_hsn":            raw_hsn,
            "healed_hsn":         healed_hsn,        # ← filled from master if was blank
            "healed_from_master": healed_from_master,
            "rate_str":           rate_str,
            "igst_type":          igst_type,
            "tx_status":          tx_status,
        })

pb.progress(55, text="📊 Computing global majority tax rates…")

# =============================================================================
# STEP 5 — GLOBAL MAJORITY RATE
#   ✅ Computed GLOBALLY across ALL sheets (not per-sheet)
#   ✅ Uses healed_hsn so filled codes benefit from majority too
# =============================================================================
_hsn_rate_pool = {}
for r in records:
    h, rt = r["healed_hsn"], r["rate_str"]
    if h and rt not in {"", "0", "0.0"}:
        _hsn_rate_pool.setdefault(h, []).append(rt)

global_majority_rate = {h: majority_vote(rates) for h, rates in _hsn_rate_pool.items() if rates}


def resolve_rate(healed_hsn, clean_sku, original_rate):
    """
    Priority chain for final tax rate:
    1. Global majority rate (most common non-zero rate seen for this HSN across all sheets)
    2. Master catalog HSN→rate
    3. Master catalog SKU→rate
    4. Original rate from the row
    """
    if healed_hsn in global_majority_rate:     return global_majority_rate[healed_hsn]
    if healed_hsn in master_hsn_tax:           return master_hsn_tax[healed_hsn]
    if clean_sku  in master_sku_tax:           return master_sku_tax[clean_sku]
    return original_rate


pb.progress(65, text="🔍 Detecting compliance errors…")

# =============================================================================
# STEP 6 — ERROR DETECTION
# =============================================================================
# Double-rate detection: same HSN → multiple distinct rates across all data
_hsn_multi = {}
for r in records:
    h, rt = r["healed_hsn"], r["rate_str"]
    if h and rt not in {"", "0", "0.0"}:
        _hsn_multi.setdefault(h, set()).add(rt)
double_rate_map = {h: ",".join(sorted(v)) for h, v in _hsn_multi.items() if len(v) > 1}

list_healed = []       # HSNs that were blank → filled from master
list_missing = []      # HSNs still blank after master lookup (no match found)
list_double = []       # Same HSN has conflicting tax rates
list_invalid_len = []  # HSN length ≠ 6 or 8
list_wrong_tax_hsn = []
list_wrong_hsn_sku = []
list_wrong_tax_sku = []

_seen = {k: set() for k in ("hld","miss","dbl","inv","wth","whs","wts")}

for r in records:
    h   = r["healed_hsn"]
    raw = r["raw_hsn"]
    rt  = r["rate_str"]
    sd  = r["sku_display"]
    cs  = r["clean_sku"]
    ts  = r["tx_status"]

    # Healed from master
    if r["healed_from_master"] and sd not in _seen["hld"]:
        list_healed.append({"SKU": sd, "Filled HSN": h})
        _seen["hld"].add(sd)

    # Still missing after master lookup
    if not h and "cancel" not in ts and sd and sd not in _seen["miss"]:
        list_missing.append(sd)
        _seen["miss"].add(sd)

    # Conflicting rates (will be resolved by majority in output)
    if h in double_rate_map:
        key = (h, sd)
        if key not in _seen["dbl"]:
            list_double.append({
                "HSN": h, "SKU": sd,
                "Rates Found": double_rate_map[h],
                "Majority Rate (Applied)": global_majority_rate.get(h, "—")
            })
            _seen["dbl"].add(key)

    # Invalid length — must be 6 or 8 digits
    if h and len(h) not in {6, 8} and h not in _seen["inv"]:
        list_invalid_len.append({"HSN Code": h, "Digit Count": len(h)})
        _seen["inv"].add(h)

    # Wrong rate vs master HSN tax
    if h in master_hsn_tax:
        m_tax = master_hsn_tax[h]
        key = (h, rt)
        if rt not in {"0",""} and rt != m_tax and key not in _seen["wth"]:
            list_wrong_tax_hsn.append({"HSN": h, "Input Rate": rt, "Master Rate": m_tax})
            _seen["wth"].add(key)

    # Wrong HSN vs master SKU map (only flag original, non-healed mismatches)
    if cs in master_sku_hsn and raw:
        m_hsn = master_sku_hsn[cs]
        key = (cs, raw)
        if raw != m_hsn and key not in _seen["whs"]:
            list_wrong_hsn_sku.append({"SKU": sd, "Input HSN": raw, "Master HSN": m_hsn})
            _seen["whs"].add(key)

    # Wrong rate vs master SKU tax
    if cs in master_sku_tax:
        m_tax = master_sku_tax[cs]
        key = (cs, rt)
        if rt not in {"0",""} and rt != m_tax and key not in _seen["wts"]:
            list_wrong_tax_sku.append({"SKU": sd, "Input Rate": rt, "Master Rate": m_tax})
            _seen["wts"].add(key)

pb.progress(78, text="🛠️ Sanitizing workbook…")

# =============================================================================
# STEP 7 — SANITIZATION
#   ✅ Writes healed_hsn (filled from master) to HSN column
#   ✅ Applies global majority rate (with master fallback) to all rate columns
#   ✅ Correctly splits IGST vs CGST+SGST based on original row type
# =============================================================================
sanitized = {}

for sname, df_s in raw_sheets.items():
    df_out = df_s.copy()
    c = detect_columns(df_out)
    sheet_recs = [r for r in records if r["sheet"] == sname]

    out_hsns  = []
    out_rates = []

    for r in sheet_recs:
        h  = r["healed_hsn"] or "MISSING HSN"
        rt = resolve_rate(r["healed_hsn"], r["clean_sku"], r["rate_str"])
        out_hsns.append(excel_hsn(h))   # ✅ preserves leading zeros via ="..." formula
        out_rates.append(rt)

    # Write healed HSN column and new Total Tax Rate column
    if c["hsn"]:
        df_out[c["hsn"]] = out_hsns
    df_out["Total Tax Rate"] = out_rates

    # Write IGST / CGST / SGST individual columns
    for pos, (df_idx, _) in enumerate(df_out.iterrows()):
        winner   = out_rates[pos]
        igst_type = sheet_recs[pos]["igst_type"]
        try:
            w = float(winner)
            if w == 0: continue

            if igst_type is True:
                # Inter-state transaction → IGST only
                if c["igst_rate"]: df_out.at[df_idx, c["igst_rate"]] = winner
                if c["cgst_rate"]: df_out.at[df_idx, c["cgst_rate"]] = "0"
                if c["sgst_rate"]: df_out.at[df_idx, c["sgst_rate"]] = "0"
            elif igst_type is False:
                # Intra-state → split equally into CGST + SGST
                half = rate_to_str(w / 2)
                if c["cgst_rate"]: df_out.at[df_idx, c["cgst_rate"]] = half
                if c["sgst_rate"]: df_out.at[df_idx, c["sgst_rate"]] = half
                if c["igst_rate"]: df_out.at[df_idx, c["igst_rate"]] = "0"
            # igst_type is None → zero-rate or unknown, leave columns untouched
        except Exception:
            pass

    sanitized[sname] = df_out

pb.progress(90, text="📊 Building output Excel files…")

# =============================================================================
# STEP 8 — BUILD OUTPUT FILES
# =============================================================================
max_r = max(len(list_healed), len(list_missing), len(list_double),
            len(list_invalid_len), len(list_wrong_tax_hsn),
            len(list_wrong_hsn_sku), len(list_wrong_tax_sku), 1)

def _pad(lst, key=None):
    src = [row[key] if (key and isinstance(row, dict)) else row for row in lst]
    return [src[i] if i < len(src) else "" for i in range(max_r)]

audit_df = pd.DataFrame({
    "Healed from Master — SKU":          _pad(list_healed, "SKU"),
    "Healed from Master — Filled HSN":   _pad(list_healed, "Filled HSN"),
    "Still Missing HSN (SKU)":           _pad(list_missing),
    "Invalid HSN Length — Code":         _pad(list_invalid_len, "HSN Code"),
    "Invalid HSN Length — Digit Count":  _pad(list_invalid_len, "Digit Count"),
    "Double Rate — HSN":                 _pad(list_double, "HSN"),
    "Double Rate — SKU":                 _pad(list_double, "SKU"),
    "Double Rate — Rates Found":         _pad(list_double, "Rates Found"),
    "Double Rate — Majority Applied":    _pad(list_double, "Majority Rate (Applied)"),
    "Wrong Tax (HSN) — HSN":             _pad(list_wrong_tax_hsn, "HSN"),
    "Wrong Tax (HSN) — Input Rate":      _pad(list_wrong_tax_hsn, "Input Rate"),
    "Wrong Tax (HSN) — Master Rate":     _pad(list_wrong_tax_hsn, "Master Rate"),
    "Wrong HSN (SKU) — SKU":             _pad(list_wrong_hsn_sku, "SKU"),
    "Wrong HSN (SKU) — Input HSN":       _pad(list_wrong_hsn_sku, "Input HSN"),
    "Wrong HSN (SKU) — Master HSN":      _pad(list_wrong_hsn_sku, "Master HSN"),
    "Incorrect GST Rate — SKU":          _pad(list_wrong_tax_sku, "SKU"),
    "Incorrect GST Rate — Input Rate":   _pad(list_wrong_tax_sku, "Input Rate"),
    "Incorrect GST Rate — Master Rate":  _pad(list_wrong_tax_sku, "Master Rate"),
})

# Color-coded audit Excel
GROUPS = [
    (range(0, 2),   "#1a6b45"),   # Healed (green)
    (range(2, 3),   "#c0392b"),   # Still missing (red)
    (range(3, 5),   "#e67e22"),   # Invalid length (orange)
    (range(5, 9),   "#8e44ad"),   # Double rate (purple)
    (range(9, 12),  "#2980b9"),   # Wrong tax HSN (blue)
    (range(12, 15), "#16a085"),   # Wrong HSN SKU (teal)
    (range(15, 18), "#d35400"),   # Wrong tax SKU (dark orange)
]

audit_buf = io.BytesIO()
with pd.ExcelWriter(audit_buf, engine="xlsxwriter") as writer:
    audit_df.to_excel(writer, sheet_name="HSN_GST_Audit", index=False)
    wb = writer.book
    ws = writer.sheets["HSN_GST_Audit"]
    ws.set_row(0, 36)
    for col_range, color in GROUPS:
        fmt = wb.add_format({"bold": True, "bg_color": color, "font_color": "#FFFFFF",
                              "border": 1, "text_wrap": True, "valign": "vcenter"})
        for ci in col_range:
            if ci < len(audit_df.columns):
                ws.write(0, ci, audit_df.columns[ci], fmt)
                ws.set_column(ci, ci, 26)
audit_buf.seek(0)
audit_bytes = audit_buf.getvalue()

clean_buf = io.BytesIO()
with pd.ExcelWriter(clean_buf, engine="xlsxwriter") as writer:
    for sname, df_out in sanitized.items():
        df_out.to_excel(writer, sheet_name=sname, index=False)
clean_buf.seek(0)
clean_bytes = clean_buf.getvalue()

pb.progress(100, text="✅ Done!")
pb.empty()

# =============================================================================
# UI — RESULTS DASHBOARD
# =============================================================================
with st.sidebar:
    st.header("📊 Audit Summary")
    st.metric("✅ HSNs Healed from Master", len(list_healed))
    st.metric("🔴 Still Missing HSN",       len(list_missing))
    st.metric("⚠️ Invalid HSN Lengths",     len(list_invalid_len))
    st.metric("🔁 Double Rate HSNs",         len(list_double))
    st.metric("❌ Wrong Tax (HSN)",          len(list_wrong_tax_hsn))
    st.metric("🔀 Wrong HSN (SKU)",          len(list_wrong_hsn_sku))
    st.metric("💸 Incorrect GST Rate",       len(list_wrong_tax_sku))
    st.divider()
    total = sum(map(len, [list_missing, list_invalid_len, list_double,
                           list_wrong_tax_hsn, list_wrong_hsn_sku, list_wrong_tax_sku]))
    st.metric("⚡ Total Issues", total)
    if total == 0: st.success("No compliance issues!")
    else: st.warning(f"{total} issues found.")
    st.divider()
    st.subheader("🎯 Column Detection")
    for sname, c in col_maps:
        st.markdown(f"**{sname}**  \nHSN: `{c['hsn'] or '—'}` | SKU: `{c['sku'] or '—'}`")
    if attribute_file:
        st.divider()
        st.subheader("📋 Catalog")
        if catalog_sku_col and catalog_hsn_col:
            st.success(f"SKU col: `{catalog_sku_col}`  \nHSN col: `{catalog_hsn_col}`")
            st.metric("SKU→HSN mappings", len(master_sku_hsn))
        else:
            st.error("Could not detect SKU/HSN columns in catalog.")

# Banner
st.success(
    f"✅ **{len(raw_sheets)}** sheet(s) · **{len(records)}** rows · "
    f"**{len(list_healed)}** HSNs healed from master · "
    f"**{len(global_majority_rate)}** unique HSNs rate-normalised"
)

# 7-column metrics
mc = st.columns(7)
mc[0].metric("✅ Healed HSNs",    len(list_healed),        help="HSN filled from master catalog")
mc[1].metric("🔴 Missing HSN",    len(list_missing))
mc[2].metric("⚠️ Invalid Len",   len(list_invalid_len))
mc[3].metric("🔁 Double Rate",    len(list_double))
mc[4].metric("❌ Wrong Tax/HSN",  len(list_wrong_tax_hsn))
mc[5].metric("🔀 Wrong HSN/SKU",  len(list_wrong_hsn_sku))
mc[6].metric("💸 Wrong GST Rate", len(list_wrong_tax_sku))

st.divider()
CAT_NOTE = "" if attribute_file else "Upload a master catalog (Step 2) to enable this check."

with st.expander(f"✅ HSNs Healed from Master Catalog — {len(list_healed)}", expanded=bool(list_healed)):
    if list_healed:
        st.info("These HSN codes were BLANK in the source but auto-filled from the master catalog in the sanitized output.")
        st.dataframe(pd.DataFrame(list_healed), use_container_width=True, height=220)
    elif attribute_file:
        st.success("No healing needed — all rows already had HSN codes.")
    else: st.info(CAT_NOTE)

with st.expander(f"🔴 Still Missing HSN (no master match) — {len(list_missing)}", expanded=bool(list_missing)):
    if list_missing:
        st.warning("These SKUs have no HSN in the data AND no match in the master catalog.")
        st.dataframe(pd.DataFrame({"SKU": list_missing}), use_container_width=True, height=220)
    else: st.success("All rows have an HSN (or were healed from master).")

with st.expander(f"🔁 Conflicting Tax Rates (Same HSN) — {len(list_double)}", expanded=bool(list_double)):
    if list_double:
        st.info("Majority rate shown in last column has been applied to the sanitized output.")
        st.dataframe(pd.DataFrame(list_double), use_container_width=True, height=220)
    else: st.success("No conflicting tax rates found.")

with st.expander(f"⚠️ Invalid HSN Lengths — {len(list_invalid_len)}", expanded=False):
    if list_invalid_len:
        st.dataframe(pd.DataFrame(list_invalid_len), use_container_width=True, height=220)
    else: st.success("All HSN codes are 6 or 8 digits.")

with st.expander(f"❌ Wrong Tax Rate (HSN-Based) — {len(list_wrong_tax_hsn)}", expanded=False):
    if list_wrong_tax_hsn:
        st.dataframe(pd.DataFrame(list_wrong_tax_hsn), use_container_width=True, height=220)
    elif attribute_file: st.success("No HSN-based rate mismatches.")
    else: st.info(CAT_NOTE)

with st.expander(f"🔀 Wrong HSN per SKU — {len(list_wrong_hsn_sku)}", expanded=False):
    if list_wrong_hsn_sku:
        st.dataframe(pd.DataFrame(list_wrong_hsn_sku), use_container_width=True, height=220)
    elif attribute_file: st.success("No SKU→HSN mismatches.")
    else: st.info(CAT_NOTE)

with st.expander(f"💸 Incorrect GST Rate (SKU-Based) — {len(list_wrong_tax_sku)}", expanded=False):
    if list_wrong_tax_sku:
        st.dataframe(pd.DataFrame(list_wrong_tax_sku), use_container_width=True, height=220)
    elif attribute_file: st.success("No SKU-based rate errors.")
    else: st.info(CAT_NOTE)

st.divider()

base = uploaded_file.name.rsplit(".", 1)[0]
d1, d2 = st.columns(2)
with d1:
    st.download_button("📥 Download Audit Error Report", audit_bytes,
                        f"ERROR_REPORT_{base}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, type="secondary")
with d2:
    st.download_button("📥 Download Sanitized Workbook", clean_bytes,
                        f"CLEANED_{base}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, type="primary")

st.write("### 📋 Sanitized Preview — first 50 rows")
first = list(sanitized.keys())[0]
st.dataframe(sanitized[first].head(50), use_container_width=True)
