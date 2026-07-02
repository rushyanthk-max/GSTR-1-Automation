import streamlit as st
import pandas as pd
import re
import io

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BCPL Universal GST Sanitizer & Auditor",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def clean_df_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column headers; auto-detect header row when missing HSN/SKU/ASIN cols."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)          # Always ensure clean 0-based RangeIndex

    HSN_KW = {"hsn", "sac", "commodity", "nomenclature"}
    SKU_KW = {"sku", "fsn", "seller-sku", "item-code", "product-id", "article", "asin"}

    has_hsn = any(any(k in str(c).lower() for k in HSN_KW) for c in df.columns)
    has_sku = any(any(k in str(c).lower() for k in SKU_KW) for c in df.columns)

    if not (has_hsn or has_sku):
        TRIGGERS = {"hsn", "hsn code", "hsn/sac", "sku", "seller sku", "fsn", "asin", "asin code"}
        for idx, row in df.iterrows():
            vals = {str(v).strip().lower() for v in row.values if pd.notna(v)}
            if vals & TRIGGERS:
                df.columns = [str(v).strip() for v in row.values]
                df = df.iloc[idx + 1:].reset_index(drop=True)
                break
    return df


def detect_columns_v2(df: pd.DataFrame) -> dict:
    """
    Two-pass prioritized column scanner. Ensures strong exact structural matches 
    override loose partial matches, and dynamically auto-detects ASIN columns.
    """
    c = {k: None for k in (
        "sku", "hsn", "cgst_rate", "sgst_rate",
        "igst_rate", "order_id", "order_item_id", "tx_type", "asin"
    )}
    
    # PASS 1: Direct structural locks & explicit template hits (Flipkart/Amazon/Master Catalog)
    for col in df.columns:
        cl = str(col).strip().lower()
        if cl == 'sku': c['sku'] = col
        if cl in ['hsn', 'hsn code', 'hsn/sac', 'producttaxcode', 'commodity code', 'hsn_code', 'hsncode']: c['hsn'] = col
        if cl in ['asin', 'asin code', 'product-asin', 'item-asin', 'asin_code']: c['asin'] = col
        if cl == 'order id': c['order_id'] = col
        if cl == 'order item id': c['order_item_id'] = col
        if cl in ['cgst rate', 'cgst_rate']: c['cgst_rate'] = col
        if cl in ['sgst rate', 'sgst_rate', 'sgst rate (or utgst as applicable)']: c['sgst_rate'] = col
        if cl in ['igst rate', 'igst_rate']: c['igst_rate'] = col
        if cl in ['event type', 'document type', 'transaction type']: c['tx_type'] = col

    # Heuristic: Scan first 5 rows to auto-detect any un-mapped custom ASIN data streams
    if not c['asin']:
        for col in df.columns:
            sample_vals = df[col].dropna().head(5).astype(str).str.strip().str.replace('`','').str.replace('"','')
            if any(re.match(r'^B0[A-Z0-9]{8}$', v, re.IGNORECASE) for v in sample_vals):
                c['asin'] = col
                break

    # PASS 2: Universal fallback scanner for custom formats
    SKIP = {"tcs", "shipping", "gift", "wrap", "delivery", "postage", "cst", "vat", "cess", "tds", "amount", "amt", "value"}
    for col in df.columns:
        cl = str(col).strip().lower()
        if not c["sku"]:
            if any(k == cl for k in ("sku", "seller-sku", "item-code", "article-code", "wms_code")):
                c["sku"] = col
            elif any(k in cl for k in ("sku", "fsn", "product-id", "article", "fsn code")) and not any(x in cl for x in ['type', 'parent', 'accounting', 'parent_sku']):
                c["sku"] = col
        if not c["hsn"]:
            if any(k in cl for k in ("hsn", "sac", "commodity", "nomenclature", "taxcode", "taxrule", "producttaxcode")):
                c["hsn"] = col
        if not c["tx_type"] and any(k in cl for k in ("type", "status", "order status", "transaction type")):
            c["tx_type"] = col
        if any(x in cl for x in SKIP):
            continue
        if not c["cgst_rate"] and "cgst" in cl and "rate" in cl: c["cgst_rate"] = col
        if not c["sgst_rate"] and ("sgst" in cl or "utgst" in cl) and "rate" in cl: c["sgst_rate"] = col
        if not c["igst_rate"] and "igst" in cl and "rate" in cl: c["igst_rate"] = col

    return c


def deep_clean_sku(val) -> str:
    """Return a lowercase, alphanumeric-only normalized SKU string."""
    if pd.isna(val):
        return ""
    s = str(val).strip().lower()
    s = s.replace('"', '').replace("'", "").replace("`", "")
    if s.startswith("sku:"):
        s = s[4:]
    elif s.startswith("sku"):
        s = s[3:]
    s = re.sub(r'^[^a-z0-9]+', '', s)
    s = re.sub(r'[^a-z0-9]', '', s)
    return s


def extract_rate_number(val) -> float:
    """Parse a raw GST rate cell value into a float percentage whole integer."""
    if pd.isna(val) or str(val).strip() in {"", "nan", "None", "<NA>"}:
        return 0.0
    s = str(val).strip().replace("%", "").strip()
    s = re.sub(r"\.0+$", "", s)

    EXACT: dict[str, float] = {
        "0.028": 28, "0.018": 18, "0.012": 12, "0.005": 5, "0.003": 3,
        "0.28":  28, "0.18":  18, "0.12":  12, "0.05":  5, "0.03":  3,
    }
    if s in EXACT:
        return float(EXACT[s])

    digits = "".join(ch for ch in s if ch.isdigit() or ch == ".")
    try:
        num = float(digits) if digits else 0.0
        if 0 < num <= 1.0:
            num *= 100
        return num
    except Exception:
        return 0.0


def normalize_hsn(val) -> str:
    """Normalize any HSN code representation to a clean digit-only string with 7-to-8 digit zero padding."""
    if pd.isna(val) or str(val).strip() in {"", "nan", "None"}:
        return ""
    h = str(val).strip()
    if h.startswith('="') and h.endswith('"'):
        h = h[2:-1]
    digits = "".join(filter(str.isdigit, re.sub(r"[\s\-\.\/]", "", h)))
    if len(digits) == 7:
        digits = "0" + digits
    return digits


# =============================================================================
# UI — HEADER & FILE UPLOADER COMPONENTS
# =============================================================================
st.title("📦 BCPL Universal E-commerce GST Sanitizer & Auditor")
st.caption(
    "Upload your files to clean multi-sheet workbooks and generate a unified "
    "side-by-side Audit Error Report based on raw inputs."
)

up_col1, up_col2, up_col3 = st.columns(3)
with up_col1:
    st.subheader("1️⃣ Raw Transaction Report")
    uploaded_file = st.file_uploader(
        "Sales workbook (.xlsx / .csv)",
        type=["xlsx", "xls", "csv"],
        key="sales_report",
    )
with up_col2:
    st.subheader("2️⃣ Master Catalog (ProductAttribute)")
    attribute_file = st.file_uploader(
        "Catalog — maps master HSN fields",
        type=["xlsx", "xls", "csv"],
        key="attribute_sheet",
    )
with up_col3:
    st.subheader("3️⃣ Indian SKU List (For STN Sheets)")
    stn_mapping_file = st.file_uploader(
        "Mapping — translates ASIN to SKU codes",
        type=["xlsx", "xls", "csv"],
        key="stn_mapping_sheet",
    )

if not uploaded_file:
    st.info("👆 Upload a sales / transaction report above to get started.")
    st.stop()

# =============================================================================
# LOAD TRANSACTION SHEETS
# =============================================================================
progress_bar = st.progress(0, text="📂 Loading transaction reports…")

raw_sheets_dict: dict[str, pd.DataFrame] = {}
try:
    if uploaded_file.name.endswith((".xlsx", ".xls")):
        xf = pd.ExcelFile(uploaded_file)
        for sname in xf.sheet_names:
            raw_sheets_dict[sname] = clean_df_columns(
                pd.read_excel(uploaded_file, sheet_name=sname, dtype=str)
            )
    else:
        raw_sheets_dict["Sales Report"] = clean_df_columns(
            pd.read_csv(uploaded_file, dtype=str, low_memory=False)
        )
except Exception as err:
    st.error(f"❌ Failed to read uploaded transaction file: {err}")
    st.stop()

progress_bar.progress(15, text="🔗 Building cross-sheet order reference map…")

# =============================================================================
# CROSS-SHEET ORDER ID MAPPING REGISTRY
# =============================================================================
sales_lookup_df = pd.DataFrame(columns=["Order ID", "Order Item ID", "SKU", "HSN Code"])
for sname, df_s in raw_sheets_dict.items():
    c = detect_columns_v2(df_s)
    if c["order_id"] and c["order_item_id"] and c["sku"]:
        sub = pd.DataFrame()
        sub["Order ID"] = df_s[c["order_id"]].fillna("").astype(str).str.strip()
        sub["Order Item ID"] = df_s[c["order_item_id"]].fillna("").astype(str).str.strip()
        sub["SKU"] = df_s[c["sku"]].fillna("").astype(str).str.strip()
        sub["HSN Code"] = df_s[c["hsn"]].fillna("").astype(str).str.strip() if c["hsn"] else ""
        
        sales_lookup_df = pd.concat(
            [sales_lookup_df, sub.dropna(subset=["Order ID", "Order Item ID"])],
            ignore_index=True,
        )
sales_lookup_df.drop_duplicates(subset=["Order ID", "Order Item ID"], inplace=True)

progress_bar.progress(25, text="📋 Loading custom Indian ASIN-SKU mapping list…")

# =============================================================================
# LOAD AND PARSE CUSTOM STN ASIN MAPPING FILE
# =============================================================================
custom_asin_sku_map: dict[str, str] = {}
if stn_mapping_file:
    try:
        if stn_mapping_file.name.endswith((".xlsx", ".xls")):
            stn_m_df = clean_df_columns(pd.read_excel(stn_mapping_file, dtype=str))
        else:
            stn_m_df = clean_df_columns(pd.read_csv(stn_mapping_file, dtype=str, low_memory=False))
        
        stn_c = detect_columns_v2(stn_m_df)
        if stn_c["asin"] and stn_c["sku"]:
            for _, row in stn_m_df.iterrows():
                m_asin = str(row[stn_c["asin"]]).strip().replace('`','').replace('"','').upper()
                m_sku = str(row[stn_c["sku"]]).strip().replace('`','').replace('"','')
                if m_asin and m_sku:
                    custom_asin_sku_map[m_asin] = m_sku
            st.sidebar.success(f"📌 Loaded {len(custom_asin_sku_map)} custom ASIN links from Indian SKU List!")
    except Exception as err:
        st.warning(f"⚠️ Failed to map custom STN index files: {err}")

progress_bar.progress(40, text="📋 Mapping main product catalog fields…")

# =============================================================================
# MASTER CATALOG LIBRARIES PARSING
# =============================================================================
master_sku_hsn: dict[str, str] = {}
master_sku_tax: dict[str, str] = {}
master_hsn_tax: dict[str, str] = {}

if attribute_file:
    try:
        if attribute_file.name.endswith((".xlsx", ".xls")):
            attr_df = clean_df_columns(pd.read_excel(attribute_file, dtype=str))
        else:
            attr_df = clean_df_columns(
                pd.read_csv(attribute_file, dtype=str, low_memory=False)
            )

        attr_c = detect_columns_v2(attr_df)

        if attr_c["hsn"] and attr_c["sku"]:
            a_tax = None
            for col in attr_df.columns:
                cl = str(col).strip().lower()
                if any(k in cl for k in ("producttaxrule", "tax rule", "gst rate", "tax percentage", "tax_rule")):
                    a_tax = col; break

            for _, row in attr_df.iterrows():
                r_hsn = normalize_hsn(row[attr_c["hsn"]]) if pd.notna(row[attr_c["hsn"]]) else ""
                r_sku = deep_clean_sku(row[attr_c["sku"]])
                
                r_tax = (
                    "".join(filter(str.isdigit, str(row[a_tax]).strip()))
                    if a_tax and pd.notna(row[a_tax])
                    else ""
                )
                if r_hsn and r_sku:
                    master_sku_hsn[r_sku] = r_hsn
                    if r_tax:
                        master_sku_tax[r_sku] = r_tax
                        master_hsn_tax[r_hsn] = r_tax

                # Internal model search to ensure standard ASIN strings cross-reference if present in master
                for col in attr_df.columns:
                    val_str = str(row[col]).strip().replace('`','').replace('"','').upper()
                    if re.match(r'^B0[A-Z0-9]{8}$', val_str) and r_sku:
                        # Fallback mapping if no custom sheet is uploaded
                        if val_str not in custom_asin_sku_map:
                            custom_asin_sku_map[val_str] = str(row[attr_c["sku"]]).strip().replace('`','').replace('"','')
    except Exception as err:
        st.warning(f"⚠️ Could not fully load master catalog mappings: {err}")

progress_bar.progress(60, text="🕵️ Evaluating raw layout data profiles for audit logs…")

# =============================================================================
# BUILD UNMUTATED RAW DATA MASTER ARRAYS
# =============================================================================
global_raw_records = []
discovered_mappings = []

for sname, df_s in raw_sheets_dict.items():
    c = detect_columns_v2(df_s)
    discovered_mappings.append(
        f"**{sname}** — HSN col: `{c['hsn'] or '—'}` | SKU col: `{c['sku'] or '—'}`"
    )

    df_work = df_s.copy()
    if c["order_id"] and c["order_item_id"] and (not c["sku"] or not c["hsn"]):
        df_work = df_work.merge(sales_lookup_df, on=["Order ID", "Order Item ID"], how="left")
        if not c["sku"] and "SKU" in df_work.columns: c["sku"] = "SKU"
        if not c["hsn"] and "HSN Code" in df_work.columns: c["hsn"] = "HSN Code"

    cgst_s = df_work[c["cgst_rate"]].apply(extract_rate_number) if c["cgst_rate"] else pd.Series(0.0, index=df_work.index)
    sgst_s = df_work[c["sgst_rate"]].apply(extract_rate_number) if c["sgst_rate"] else pd.Series(0.0, index=df_work.index)
    igst_s = df_work[c["igst_rate"]].apply(extract_rate_number) if c["igst_rate"] else pd.Series(0.0, index=df_work.index)

    for pos, (df_idx, row) in enumerate(df_work.iterrows()):
        rhsn = str(row[c["hsn"]]).strip() if c["hsn"] and pd.notna(row[c["hsn"]]) else ""
        rsku = str(row[c["sku"]]).strip() if c["sku"] and pd.notna(row[c["sku"]]) else ""
        rasin = str(row[c["asin"]]).strip() if c["asin"] and pd.notna(row[c["asin"]]) else ""

        rasin_clean = rasin.upper().replace('"', '').replace("'", "").replace("`","").strip()
        rsku_clean_upper = rsku.upper().replace('"', '').replace("'", "").replace("`","").strip()

        if re.match(r'^B0[A-Z0-9]{8}$', rsku_clean_upper):
            if not rasin_clean: rasin_clean = rsku_clean_upper

        # Resolve true custom mapping links to identify true catalog properties before evaluation
        if rasin_clean in custom_asin_sku_map:
            rsku = custom_asin_sku_map[rasin_clean]

        hsn_dig = normalize_hsn(rhsn)
        csku = deep_clean_sku(rsku)

        total = igst_s.loc[df_idx] if igst_s.loc[df_idx] > 0 else (cgst_s.loc[df_idx] + sgst_s.loc[df_idx])
        rate_str = str(int(total)) if total == int(total) else str(total)
        tx_status = str(row[c["tx_type"]]).strip().lower() if c["tx_type"] else ""

        sku_disp_clean = rsku.replace('"', '').replace("'", "").replace("`", "").strip() if rsku else ""

        global_raw_records.append({
            "sheet":       sname,
            "df_idx":      df_idx,
            "sku_display": sku_disp_clean,
            "clean_sku":   csku,
            "raw_asin_clean": rasin_clean,
            "raw_hsn":     hsn_dig,
            "rate_str":    rate_str,
            "tx_status":   tx_status,
        })

progress_bar.progress(75, text="🔍 Running global majority tax vote maps…")

# =============================================================================
# GLOBAL MAJORITY VOTE PRE-CALCULATION & COMPLIANCE RISK AUDITING
# =============================================================================
global_hsn_rates: dict[str, list[str]] = {}
for r in global_raw_records:
    h_healed = r["raw_hsn"]
    if not h_healed and r["clean_sku"] in master_sku_hsn: 
        h_healed = master_sku_hsn[r["clean_sku"]]
    
    rt = r["rate_str"]
    if h_healed and h_healed != "" and rt not in {"0", "0.0", ""}:
        global_hsn_rates.setdefault(h_healed, []).append(rt)

global_majority_tax: dict[str, str] = {}
for hsn, rates in global_hsn_rates.items():
    if rates: global_majority_tax[hsn] = max(set(rates), key=rates.count)

_raw_hsn_rates = {}
for r in global_raw_records:
    h, rt = r["raw_hsn"], r["rate_str"]
    if h and rt not in {"0", "0.0", ""}: _raw_hsn_rates.setdefault(h, set()).add(rt)
double_rate_hsns = {h: ",".join(sorted(v)) for h, v in _raw_hsn_rates.items() if len(v) > 1}

list_missing_hsn = []
list_double_rates = []
list_invalid_len = []
list_wrong_tax_hsn = []
list_wrong_hsn_sku = []
list_wrong_tax_sku = []
_seen = {k: set() for k in ("miss", "dbl", "inv", "wth", "whs", "wts")}

for r in global_raw_records:
    h, rt, sd, cs, ts, ac = r["raw_hsn"], r["rate_str"], r["sku_display"], r["clean_sku"], r["tx_status"], r["raw_asin_clean"]
    
    if not h and "cancel" not in ts and sd and sd not in _seen["miss"]:
        list_missing_hsn.append(sd)
        _seen["miss"].add(sd)

    if h in double_rate_hsns:
        key = (h, sd)
        if key not in _seen["dbl"]:
            list_double_rates.append({"HSN": h, "SKU": sd, "Tax Rates Found": double_rate_hsns[h]})
            _seen["dbl"].add(key)

    if h and len(h) not in {6, 8} and h not in _seen["inv"]:
        list_invalid_len.append({"HSN Code": h, "Digit Count": len(h)})
        _seen["inv"].add(h)

    if h in master_hsn_tax:
        m_tax = master_hsn_tax[h]
        key = (h, rt)
        if rt not in {"0", ""} and rt != m_tax and key not in _seen["wth"]:
            list_wrong_tax_hsn.append({"HSN": h, "Input Rate": rt, "Master Rate": m_tax})
            _seen["wth"].add(key)

    if cs in master_sku_hsn:
        m_hsn = master_sku_hsn[cs]
        key = (cs, h)
        if h and h != m_hsn and key not in _seen["whs"]:
            list_wrong_hsn_sku.append({"SKU": sd, "Input HSN": h, "Master HSN": m_hsn})
            _seen["whs"].add(key)

    if cs in master_sku_tax:
        m_tax = master_sku_tax[cs]
        key = (cs, rt)
        if rt not in {"0", ""} and rt != m_tax and key not in _seen["wts"]:
            list_wrong_tax_sku.append({"SKU": sd, "Input Rate": rt, "Master Rate": m_tax})
            _seen["wts"].add(key)

progress_bar.progress(85, text="🛠️ Commencing multi-sheet full healing pass routines…")

# =============================================================================
# PRODUCTION WORKBOOK CLEANING PHASE
# =============================================================================
sanitized_sheets: dict[str, pd.DataFrame] = {}

for sname, df_s in raw_sheets_dict.items():
    df_out = df_s.copy()
    c = detect_columns_v2(df_out)
    sheet_recs = [r for r in global_raw_records if r["sheet"] == sname]

    if (not c["sku"] or c["sku"] == "SKU" or c["sku"] == "asin") and c["asin"]:
        df_out['SKU'] = ""
        c["sku"] = 'SKU'
    if not c["cgst_rate"]:
        df_out['CGST Rate'] = "0"
        c["cgst_rate"] = 'CGST Rate'
    if not c["sgst_rate"]:
        df_out['SGST Rate (or UTGST as applicable)'] = "0"
        c["sgst_rate"] = 'SGST Rate (or UTGST as applicable)'
    if not c["igst_rate"]:
        df_out['IGST Rate'] = "0"
        c["igst_rate"] = 'IGST Rate'

    healed_hsns = []
    final_skus = []
    for r in sheet_recs:
        h_val = r["raw_hsn"]
        s_val = r["sku_display"]
        
        # Cross-reference ASIN targets back to true catalog SKU descriptors for the download workbook
        if r["raw_asin_clean"] in custom_asin_sku_map:
            s_val = custom_asin_sku_map[r["raw_asin_clean"]]
            
        # Complete full lookup heal targeting master files directly (Removes MISSING HSN values entirely)
        if not h_val or h_val == "":
            c_lookup_sku = deep_clean_sku(s_val)
            if c_lookup_sku in master_sku_hsn: 
                h_val = master_sku_hsn[c_lookup_sku]
            else: 
                h_val = "MISSING HSN"
            
        healed_hsns.append(h_val)
        final_skus.append(s_val)

    final_hsns  = [f'="{h}"' if h != "MISSING HSN" else "MISSING HSN" for h in healed_hsns]
    final_rates = [global_majority_tax.get(h, r["rate_str"]) for h, r in zip(healed_hsns, sheet_recs)]

    if c["sku"]: df_out[c["sku"]] = final_skus
    if c["hsn"]: df_out[c["hsn"]] = final_hsns
    else: df_out['HSN Code'] = final_hsns
    df_out["Total Tax Rate"] = final_rates

    for pos, (df_idx, _) in enumerate(df_out.iterrows()):
        winner = final_rates[pos]
        try:
            w = float(winner)
            if w == 0: continue
            orig_igst = "0"
            orig_igst_col = detect_columns_v2(df_s)["igst_rate"]
            if orig_igst_col: orig_igst = str(df_s.at[df_idx, orig_igst_col]).strip()
            if orig_igst not in {"", "0", "0.0"}:
                df_out.at[df_idx, c["igst_rate"]] = winner
                df_out.at[df_idx, c["cgst_rate"]] = "0"
                df_out.at[df_idx, c["sgst_rate"]] = "0"
            else:
                half = w / 2
                split = str(int(half)) if half == int(half) else str(half)
                df_out.at[df_idx, c["cgst_rate"]] = split
                df_out.at[df_idx, c["sgst_rate"]] = split
                df_out.at[df_idx, c["igst_rate"]] = "0"
        except Exception: pass
    sanitized_sheets[sname] = df_out

progress_bar.progress(95, text="📊 Compiling downloadable Excel packages…")
max_rows = max(len(list_missing_hsn), len(list_double_rates), len(list_invalid_len), len(list_wrong_tax_hsn), len(list_wrong_hsn_sku), len(list_wrong_tax_sku), 1)

def _pad(lst: list, key: str | None = None) -> list:
    src = [row[key] if (key and isinstance(row, dict)) else row for row in lst]
    return [src[i] if i < len(src) else "" for i in range(max_rows)]

audit_df = pd.DataFrame({
    "Missing HSN Codes (SKUs)":          _pad(list_missing_hsn),
    "Invalid Length HSNs":               _pad(list_invalid_len, "HSN Code"),
    "Invalid Length — Digit Count":      _pad(list_invalid_len, "Digit Count"),
    "Double Rate — HSN":                 _pad(list_double_rates, "HSN"),
    "Double Rate — SKU":                 _pad(list_double_rates, "SKU"),
    "Double Rate — Rates Found":         _pad(list_double_rates, "Tax Rates Found"),
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

audit_buf = io.BytesIO()
with pd.ExcelWriter(audit_buf, engine="xlsxwriter") as writer:
    audit_df.to_excel(writer, sheet_name="HSN_GST_Audit_Dashboard", index=False)
    wb = writer.book
    ws = writer.sheets["HSN_GST_Audit_Dashboard"]
    hdr_fmt = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF", "border": 1, "text_wrap": True, "valign": "vcenter"})
    ws.set_row(0, 32)
    for col_i, col_name in enumerate(audit_df.columns):
        ws.write(0, col_i, col_name, hdr_fmt)
        ws.set_column(col_i, col_i, max(20, len(col_name) // 2 + 4))
audit_buf.seek(0)
audit_bytes = audit_buf.getvalue()

clean_buf = io.BytesIO()
with pd.ExcelWriter(clean_buf, engine="xlsxwriter") as writer:
    for sname, df_out in sanitized_sheets.items(): df_out.to_excel(writer, sheet_name=sname, index=False)
clean_buf.seek(0)
clean_bytes = clean_buf.getvalue()

# =============================================================================
# RESULTS HUD HUB RENDERING
# =============================================================================
with st.sidebar:
    st.header("📊 Audit Summary")
    st.metric("🔴 Missing HSN (SKUs)",  len(list_missing_hsn))
    st.metric("⚠️ Invalid HSN Lengths", len(list_invalid_len))
    st.metric("🔁 Double Tax Rates",    len(list_double_rates))
    st.metric("❌ Wrong Tax (HSN)",     len(list_wrong_tax_hsn))
    st.metric("🔀 Wrong HSN (SKU)",     len(list_wrong_hsn_sku))
    st.metric("💸 Incorrect GST Rate",  len(list_wrong_tax_sku))
    st.divider()
    total_issues = sum(map(len, [
        list_missing_hsn, list_invalid_len, list_double_rates,
        list_wrong_tax_hsn, list_wrong_hsn_sku, list_wrong_tax_sku,
    ]))
    st.metric("⚡ Total Issues", total_issues)
    st.divider()
    st.subheader("🎯 Column Detection")
    for m in discovered_mappings: st.markdown(m)

st.success(f"✅ Processed **{len(raw_sheets_dict)}** sheet(s) · **{len(global_raw_records)}** rows audited · **{len(master_sku_hsn)}** SKUs in catalog")
cols = st.columns(6)
for col, lbl, cnt in zip(cols, ["🔴 Missing HSN", "⚠️ Invalid Len", "🔁 Double Rate", "❌ Wrong Tax/HSN", "🔀 Wrong HSN/SKU", "💸 Wrong GST"], [len(list_missing_hsn), len(list_invalid_len), len(list_double_rates), len(list_wrong_tax_hsn), len(list_wrong_hsn_sku), len(list_wrong_tax_sku)]):
    col.metric(lbl, cnt)

st.divider()
CAT_NOTE = "Upload a master catalog (Step 2) to enable this check."

with st.expander(f"🔴 Missing HSN Codes — {len(list_missing_hsn)} SKU(s)", expanded=bool(list_missing_hsn)):
    if list_missing_hsn: st.dataframe(pd.DataFrame({"SKU": list_missing_hsn}), use_container_width=True, height=220)
    else: st.success("No missing HSN codes found.")

with st.expander(f"⚠️ Invalid HSN Lengths — {len(list_invalid_len)}", expanded=False):
    if list_invalid_len: st.dataframe(pd.DataFrame(list_invalid_len), use_container_width=True, height=220)
    else: st.success("All HSN codes have valid lengths (6 or 8 digits).")

with st.expander(f"🔁 Conflicting Tax Rates (Same HSN) — {len(list_double_rates)}", expanded=False):
    if list_double_rates: st.dataframe(pd.DataFrame(list_double_rates), use_container_width=True, height=220)
    else: st.success("No conflicting tax rates detected.")

with st.expander(f"❌ Wrong Tax Rate (HSN-Based) — {len(list_wrong_tax_hsn)}", expanded=False):
    if list_wrong_tax_hsn: st.dataframe(pd.DataFrame(list_wrong_tax_hsn), use_container_width=True, height=220)
    elif attribute_file: st.success("No HSN-based tax rate mismatches found.")
    else: st.info(CAT_NOTE)

with st.expander(f"🔀 Wrong HSN per SKU — {len(list_wrong_hsn_sku)}", expanded=False):
    if list_wrong_hsn_sku: st.dataframe(pd.DataFrame(list_wrong_hsn_sku), use_container_width=True, height=220)
    elif attribute_file: st.success("No SKU→HSN mismatches found.")
    else: st.info(CAT_NOTE)

with st.expander(f"💸 Incorrect GST Rate (SKU-Based) — {len(list_wrong_tax_sku)}", expanded=False):
    if list_wrong_tax_sku: st.dataframe(pd.DataFrame(list_wrong_tax_sku), use_container_width=True, height=220)
    elif attribute_file: st.success("No SKU-based GST rate errors found.")
    else: st.info(CAT_NOTE)

st.divider()
base_name = uploaded_file.name.rsplit(".", 1)[0]
dl1, dl2 = st.columns(2)
with dl1: st.download_button(label="📥 Download Unified Side-by-Side Error Report", data=audit_bytes, file_name=f"ERROR_REPORT_{base_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")
with dl2: st.download_button(label="📥 Download Sanitized Workbook", data=clean_bytes, file_name=f"CLEANED_{base_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

st.write("### 📋 Sanitized Sheet Preview (first 50 rows)")
st.dataframe(sanitized_sheets[list(sanitized_sheets.keys())[0]].head(50), use_container_widt
