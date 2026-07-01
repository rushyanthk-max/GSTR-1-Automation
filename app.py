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
    """Normalize column headers; auto-detect header row when missing HSN/SKU cols."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)          # Always ensure clean 0-based RangeIndex

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


def detect_columns(df: pd.DataFrame) -> dict:
    """
    Scan a DataFrame and return a dict of detected column names for key fields:
      sku, hsn, cgst_rate, sgst_rate, igst_rate, order_id, order_item_id, tx_type
    First match wins; returns None for any field not found.
    """
    c = {k: None for k in (
        "sku", "hsn", "cgst_rate", "sgst_rate",
        "igst_rate", "order_id", "order_item_id", "tx_type",
    )}
    SKIP = {
        "tcs", "shipping", "gift", "wrap", "delivery", "postage",
        "cst", "vat", "cess", "tds", "amount", "amt", "value",
    }
    TX_KW = {
        "transaction type", "transaction_type", "order status",
        "order_status", "event type", "event_type", "document type",
    }

    for col in df.columns:
        cl = str(col).strip().lower()

        if cl == "order id"      and not c["order_id"]:      c["order_id"]      = col
        if cl == "order item id" and not c["order_item_id"]: c["order_item_id"] = col

        if not c["sku"]:
            if cl in {"sku", "seller-sku", "item-code", "article-code", "wms_code"}:
                c["sku"] = col
            elif any(k in cl for k in ("sku", "fsn", "product-id", "article")):
                c["sku"] = col

        if not c["hsn"]:
            if cl in {"hsn", "hsn/sac", "hsn_sac", "hsncode", "hsn code",
                      "hsn_code", "commodity", "hsn sac"}:
                c["hsn"] = col
            elif any(k in cl for k in ("hsn", "sac", "commodity", "nomenclature")):
                c["hsn"] = col

        if not c["tx_type"] and any(k in cl for k in TX_KW):
            c["tx_type"] = col

        # Skip columns that are definitely not GST rate columns
        if any(x in cl for x in SKIP):
            continue

        if "cgst"  in cl and "rate" in cl and not c["cgst_rate"]:  c["cgst_rate"]  = col
        if ("sgst" in cl or "utgst" in cl) and "rate" in cl and not c["sgst_rate"]: c["sgst_rate"] = col
        if "igst"  in cl and "rate" in cl and not c["igst_rate"]:  c["igst_rate"]  = col

    return c


def deep_clean_sku(val) -> str:
    """Return a lowercase, alphanumeric-only normalized SKU string."""
    if pd.isna(val):
        return ""
    s = str(val).strip().lower()
    s = re.sub(r'^[`"\'\s]+|[`"\'\s]+$', "", s)
    if s.startswith("sku:"):
        s = s[4:]
    elif s.startswith("sku"):
        s = s[3:]
    return re.sub(r"[^a-z0-9]", "", s)


def extract_rate_number(val) -> float:
    """
    Parse a raw GST rate cell value into a float percentage.
    Handles: '18%', '18', '0.18', '0.018' (3-decimal Amazon/Flipkart export), etc.
    """
    if pd.isna(val) or str(val).strip() in {"", "nan", "None", "<NA>"}:
        return 0.0
    s = str(val).strip().replace("%", "").strip()
    s = re.sub(r"\.0+$", "", s)  # Remove trailing .000

    # Platforms sometimes export rates as 3-decimal fractions (e.g. 0.018 = 18%)
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
            num *= 100      # e.g. 0.18 → 18
        return num
    except Exception:
        return 0.0


def normalize_hsn(val) -> str:
    """
    Normalize any HSN code representation to a clean digit-only string.
    Strips Excel formula prefix (="), non-digit chars, and pads 7-digit → 8.
    """
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
# UI — HEADER & FILE UPLOADERS
# =============================================================================
st.title("📦 BCPL Universal E-commerce GST Sanitizer & Auditor")
st.caption(
    "Upload your files to clean multi-sheet workbooks and generate a "
    "side-by-side Audit Error Report."
)

up_col1, up_col2 = st.columns(2)
with up_col1:
    st.subheader("1️⃣ Raw Transaction Report")
    uploaded_file = st.file_uploader(
        "Sales sheet workbook (.xlsx / .xls / .csv)",
        type=["xlsx", "xls", "csv"],
        key="sales_report",
    )
with up_col2:
    st.subheader("2️⃣ Master Product Catalog (Optional)")
    attribute_file = st.file_uploader(
        "Item catalog — enables SKU/HSN/Tax audits & auto-healing",
        type=["xlsx", "xls", "csv"],
        key="attribute_sheet",
    )

if not uploaded_file:
    st.info("👆 Upload a sales / transaction report above to get started.")
    st.stop()

# =============================================================================
# LOAD SHEETS
# =============================================================================
progress_bar = st.progress(0, text="📂 Loading sheets…")

raw_sheets_dict: dict[str, pd.DataFrame] = {}
try:
    if uploaded_file.name.lower().endswith((".xlsx", ".xls")):
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
    st.error(f"❌ Failed to read uploaded file: {err}")
    st.stop()

progress_bar.progress(15, text="🔗 Building cross-sheet order reference…")

# =============================================================================
# CROSS-SHEET ORDER REFERENCE LOOKUP
# =============================================================================
sales_lookup_df = pd.DataFrame(columns=["Order ID", "Order Item ID", "SKU", "HSN Code"])
for sname, df_s in raw_sheets_dict.items():
    c = detect_columns(df_s)
    if c["order_id"] and c["order_item_id"] and c["sku"]:
        sub = df_s[[c["order_id"], c["order_item_id"], c["sku"]]].copy()
        sub.columns = ["Order ID", "Order Item ID", "SKU"]
        sub["HSN Code"] = df_s[c["hsn"]].values if c["hsn"] else ""
        sales_lookup_df = pd.concat(
            [sales_lookup_df, sub.dropna(subset=["Order ID", "Order Item ID"])],
            ignore_index=True,
        )
sales_lookup_df.drop_duplicates(subset=["Order ID", "Order Item ID"], inplace=True)

progress_bar.progress(28, text="📋 Loading master catalog…")

# =============================================================================
# MASTER CATALOG
# =============================================================================
master_sku_hsn: dict[str, str] = {}
master_sku_tax: dict[str, str] = {}
master_hsn_tax: dict[str, str] = {}

if attribute_file:
    try:
        if attribute_file.name.lower().endswith((".xlsx", ".xls")):
            attr_df = clean_df_columns(pd.read_excel(attribute_file, dtype=str))
        else:
            attr_df = clean_df_columns(
                pd.read_csv(attribute_file, dtype=str, low_memory=False)
            )

        a_sku = a_hsn = a_tax = None
        for col in attr_df.columns:
            cl = str(col).strip().lower()
            if not a_sku and any(k in cl for k in ("sku", "item-code", "article-code", "product-sku")):
                a_sku = col
            if not a_hsn and any(k in cl for k in ("hsn", "sac", "taxcode", "commodity",
                                                     "nomenclature", "producttaxcode")):
                a_hsn = col
            if not a_tax and any(k in cl for k in ("producttaxrule", "tax rule",
                                                     "gst rate", "tax percentage", "tax_rule")):
                a_tax = col

        if a_hsn and a_sku:
            for _, row in attr_df.iterrows():
                r_hsn = normalize_hsn(row[a_hsn]) if pd.notna(row[a_hsn]) else ""
                r_sku = deep_clean_sku(row[a_sku])
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
    except Exception as err:
        st.warning(f"⚠️ Could not fully load master catalog: {err}")

progress_bar.progress(42, text="🕵️ Auditing records…")

# =============================================================================
# BUILD RAW RECORDS  (one entry per data row across all sheets)
# =============================================================================
global_raw_records: list[dict] = []
discovered_mappings: list[str] = []

for sname, df_s in raw_sheets_dict.items():
    c = detect_columns(df_s)
    discovered_mappings.append(
        f"**{sname}** — HSN col: `{c['hsn'] or '—'}` | SKU col: `{c['sku'] or '—'}`"
    )

    for pos, (df_idx, row) in enumerate(df_s.iterrows()):
        rhsn = str(row[c["hsn"]]).strip() if c["hsn"] and pd.notna(row[c["hsn"]]) else ""
        rsku = str(row[c["sku"]]).strip() if c["sku"] and pd.notna(row[c["sku"]]) else ""

        # Fill blanks from cross-sheet order reference
        if c["order_id"] and c["order_item_id"] and (not rhsn or not rsku):
            oid  = str(row[c["order_id"]]).strip()      if pd.notna(row[c["order_id"]])      else ""
            oiid = str(row[c["order_item_id"]]).strip() if pd.notna(row[c["order_item_id"]]) else ""
            match = sales_lookup_df[
                (sales_lookup_df["Order ID"] == oid) & (sales_lookup_df["Order Item ID"] == oiid)
            ]
            if not match.empty:
                if not rsku: rsku = str(match["SKU"].values[0])
                if not rhsn: rhsn = str(match["HSN Code"].values[0])

        hsn_dig  = normalize_hsn(rhsn)
        sku_disp = re.sub(r'^["\'`\s]+|["\'`\s]+$', "", rsku)
        if sku_disp.upper().startswith("SKU:"):
            sku_disp = sku_disp[4:]

        cgst = extract_rate_number(row[c["cgst_rate"]]) if c["cgst_rate"] else 0.0
        sgst = extract_rate_number(row[c["sgst_rate"]]) if c["sgst_rate"] else 0.0
        igst = extract_rate_number(row[c["igst_rate"]]) if c["igst_rate"] else 0.0
        total = igst if igst > 0 else (cgst + sgst)
        rate_str = str(int(total)) if total == int(total) else str(total)
        tx_status = str(row[c["tx_type"]]).strip().lower() if c["tx_type"] else ""

        global_raw_records.append({
            "sheet":       sname,
            "pos":         pos,          # positional index within sheet (0-based)
            "sku_display": sku_disp,
            "clean_sku":   deep_clean_sku(rsku),
            "raw_hsn":     hsn_dig,
            "rate_str":    rate_str,
            "tx_status":   tx_status,
        })

progress_bar.progress(58, text="🔍 Detecting compliance errors…")

# =============================================================================
# ERROR DETECTION  (6 categories)
# =============================================================================

# Pre-compute HSNs that appear with more than one distinct tax rate in this file
_hsn_rates: dict[str, set] = {}
for r in global_raw_records:
    h, rt = r["raw_hsn"], r["rate_str"]
    if h and rt not in {"0", "0.0", ""}:
        _hsn_rates.setdefault(h, set()).add(rt)
double_rate_hsns = {h: ",".join(sorted(v)) for h, v in _hsn_rates.items() if len(v) > 1}

list_missing_hsn:  list = []
list_double_rates: list = []
list_invalid_len:  list = []
list_wrong_tax_hsn: list = []
list_wrong_hsn_sku: list = []
list_wrong_tax_sku: list = []

# O(1) seen-sets to avoid O(n²) duplicate checks
_seen = {k: set() for k in ("miss", "dbl", "inv", "wth", "whs", "wts")}

for r in global_raw_records:
    h, rt, sd, cs, ts = (
        r["raw_hsn"], r["rate_str"],
        r["sku_display"], r["clean_sku"], r["tx_status"],
    )

    # 1. Missing HSN code
    if not h and "cancel" not in ts and sd and sd not in _seen["miss"]:
        list_missing_hsn.append(sd)
        _seen["miss"].add(sd)

    # 2. Conflicting tax rates for the same HSN
    if h in double_rate_hsns:
        key = (h, sd)
        if key not in _seen["dbl"]:
            list_double_rates.append({"HSN": h, "SKU": sd, "Tax Rates Found": double_rate_hsns[h]})
            _seen["dbl"].add(key)

    # 3. HSN digit length not 6 or 8
    if h and len(h) not in {6, 8} and h not in _seen["inv"]:
        list_invalid_len.append({"HSN Code": h, "Digit Count": len(h)})
        _seen["inv"].add(h)

    # 4. Wrong tax rate per HSN (vs master catalog)
    if h in master_hsn_tax:
        m_tax = master_hsn_tax[h]
        key = (h, rt)
        if rt not in {"0", ""} and rt != m_tax and key not in _seen["wth"]:
            list_wrong_tax_hsn.append({"HSN": h, "Input Rate": rt, "Master Rate": m_tax})
            _seen["wth"].add(key)

    # 5. Wrong HSN for SKU (vs master catalog)
    if cs in master_sku_hsn:
        m_hsn = master_sku_hsn[cs]
        key = (cs, h)
        if h and h != m_hsn and key not in _seen["whs"]:
            list_wrong_hsn_sku.append({"SKU": sd, "Input HSN": h, "Master HSN": m_hsn})
            _seen["whs"].add(key)

    # 6. Wrong GST rate for SKU (vs master catalog)
    if cs in master_sku_tax:
        m_tax = master_sku_tax[cs]
        key = (cs, rt)
        if rt not in {"0", ""} and rt != m_tax and key not in _seen["wts"]:
            list_wrong_tax_sku.append({"SKU": sd, "Input Rate": rt, "Master Rate": m_tax})
            _seen["wts"].add(key)

progress_bar.progress(72, text="🛠️ Sanitizing workbook…")

# =============================================================================
# SANITIZATION PHASE
# =============================================================================
sanitized_sheets: dict[str, pd.DataFrame] = {}

for sname, df_s in raw_sheets_dict.items():
    df_out = df_s.copy()
    c = detect_columns(df_out)
    sheet_recs = [r for r in global_raw_records if r["sheet"] == sname]

    # Step A — Heal missing HSNs from master catalog
    healed_hsns: list[str] = []
    for r in sheet_recs:
        if not r["raw_hsn"] and r["clean_sku"] in master_sku_hsn:
            healed_hsns.append(master_sku_hsn[r["clean_sku"]])
        else:
            healed_hsns.append(r["raw_hsn"] or "MISSING HSN")

    # Step B — Majority-vote tax rate per HSN within this sheet
    _hsn_pos: dict[str, list[int]] = {}
    for pos, h in enumerate(healed_hsns):
        if h != "MISSING HSN":
            _hsn_pos.setdefault(h, []).append(pos)

    majority_tax: dict[str, str] = {}
    for h, positions in _hsn_pos.items():
        rates = [
            sheet_recs[p]["rate_str"] for p in positions
            if sheet_recs[p]["rate_str"] not in {"", "0", "0.0"}
        ]
        if rates:
            majority_tax[h] = max(set(rates), key=rates.count)

    # Step C — Build final output vectors
    final_hsns  = [f'="{h}"' if h != "MISSING HSN" else "MISSING HSN" for h in healed_hsns]
    final_rates = [majority_tax.get(h, r["rate_str"]) for h, r in zip(healed_hsns, sheet_recs)]

    if c["hsn"]:
        df_out[c["hsn"]] = final_hsns
    df_out["Total Tax Rate"] = final_rates

    # Step D — Apply IGST vs CGST+SGST split
    for pos, (df_idx, _) in enumerate(df_out.iterrows()):
        winner = final_rates[pos]
        try:
            w = float(winner)
            if w == 0:
                continue
            igst_raw = str(df_out.at[df_idx, c["igst_rate"]]).strip() if c["igst_rate"] else "0"
            if c["igst_rate"] and igst_raw not in {"", "0", "0.0"}:
                df_out.at[df_idx, c["igst_rate"]] = winner
                if c["cgst_rate"]: df_out.at[df_idx, c["cgst_rate"]] = "0"
                if c["sgst_rate"]: df_out.at[df_idx, c["sgst_rate"]] = "0"
            else:
                half = w / 2
                split = str(int(half)) if half == int(half) else str(half)
                if c["cgst_rate"]: df_out.at[df_idx, c["cgst_rate"]] = split
                if c["sgst_rate"]: df_out.at[df_idx, c["sgst_rate"]] = split
                if c["igst_rate"]: df_out.at[df_idx, c["igst_rate"]] = "0"
        except Exception:
            pass

    sanitized_sheets[sname] = df_out

progress_bar.progress(88, text="📊 Generating Excel reports…")

# =============================================================================
# BUILD AUDIT REPORT EXCEL
# =============================================================================
max_rows = max(
    len(list_missing_hsn), len(list_double_rates), len(list_invalid_len),
    len(list_wrong_tax_hsn), len(list_wrong_hsn_sku), len(list_wrong_tax_sku), 1,
)


def _pad(lst: list, key: str | None = None) -> list:
    """Pad a list of strings or dicts to max_rows length."""
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
    hdr_fmt = wb.add_format({
        "bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF",
        "border": 1, "text_wrap": True, "valign": "vcenter",
    })
    ws.set_row(0, 32)
    for col_i, col_name in enumerate(audit_df.columns):
        ws.write(0, col_i, col_name, hdr_fmt)
        ws.set_column(col_i, col_i, max(20, len(col_name) // 2 + 4))
audit_buf.seek(0)
audit_bytes = audit_buf.getvalue()

# Sanitized workbook
clean_buf = io.BytesIO()
with pd.ExcelWriter(clean_buf, engine="xlsxwriter") as writer:
    for sname, df_out in sanitized_sheets.items():
        df_out.to_excel(writer, sheet_name=sname, index=False)
clean_buf.seek(0)
clean_bytes = clean_buf.getvalue()

progress_bar.progress(100, text="✅ All done!")
progress_bar.empty()

# =============================================================================
# UI — RESULTS DASHBOARD
# =============================================================================

# ── Sidebar ───────────────────────────────────────────────────────────────────
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
    if total_issues == 0:
        st.success("No compliance issues found!")
    else:
        st.warning(f"{total_issues} issues need attention.")
    st.divider()
    st.subheader("🎯 Column Detection")
    for m in discovered_mappings:
        st.markdown(m)

# ── Status banner ─────────────────────────────────────────────────────────────
st.success(
    f"✅ Processed **{len(raw_sheets_dict)}** sheet(s) · "
    f"**{len(global_raw_records)}** rows audited · "
    f"**{len(master_sku_hsn)}** SKUs in catalog"
)

# ── Metrics row ───────────────────────────────────────────────────────────────
cols = st.columns(6)
labels = ["🔴 Missing HSN", "⚠️ Invalid Len", "🔁 Double Rate",
          "❌ Wrong Tax/HSN", "🔀 Wrong HSN/SKU", "💸 Wrong GST"]
counts = [len(list_missing_hsn), len(list_invalid_len), len(list_double_rates),
          len(list_wrong_tax_hsn), len(list_wrong_hsn_sku), len(list_wrong_tax_sku)]
for col, lbl, cnt in zip(cols, labels, counts):
    col.metric(lbl, cnt)

st.divider()

# ── Error preview expanders (one per category) ────────────────────────────────
CAT_NOTE = "Upload a master catalog (Step 2) to enable this check."

with st.expander(f"🔴 Missing HSN Codes — {len(list_missing_hsn)} SKU(s)",
                 expanded=bool(list_missing_hsn)):
    if list_missing_hsn:
        st.dataframe(pd.DataFrame({"SKU": list_missing_hsn}),
                     use_container_width=True, height=220)
    else:
        st.success("No missing HSN codes found.")

with st.expander(f"⚠️ Invalid HSN Lengths — {len(list_invalid_len)}", expanded=False):
    if list_invalid_len:
        st.dataframe(pd.DataFrame(list_invalid_len),
                     use_container_width=True, height=220)
    else:
        st.success("All HSN codes have valid lengths (6 or 8 digits).")

with st.expander(f"🔁 Conflicting Tax Rates (Same HSN) — {len(list_double_rates)}",
                 expanded=False):
    if list_double_rates:
        st.dataframe(pd.DataFrame(list_double_rates),
                     use_container_width=True, height=220)
    else:
        st.success("No conflicting tax rates detected.")

with st.expander(f"❌ Wrong Tax Rate (HSN-Based) — {len(list_wrong_tax_hsn)}",
                 expanded=False):
    if list_wrong_tax_hsn:
        st.dataframe(pd.DataFrame(list_wrong_tax_hsn),
                     use_container_width=True, height=220)
    elif attribute_file:
        st.success("No HSN-based tax rate mismatches found.")
    else:
        st.info(CAT_NOTE)

with st.expander(f"🔀 Wrong HSN per SKU — {len(list_wrong_hsn_sku)}", expanded=False):
    if list_wrong_hsn_sku:
        st.dataframe(pd.DataFrame(list_wrong_hsn_sku),
                     use_container_width=True, height=220)
    elif attribute_file:
        st.success("No SKU→HSN mismatches found.")
    else:
        st.info(CAT_NOTE)

with st.expander(f"💸 Incorrect GST Rate (SKU-Based) — {len(list_wrong_tax_sku)}",
                 expanded=False):
    if list_wrong_tax_sku:
        st.dataframe(pd.DataFrame(list_wrong_tax_sku),
                     use_container_width=True, height=220)
    elif attribute_file:
        st.success("No SKU-based GST rate errors found.")
    else:
        st.info(CAT_NOTE)

st.divider()

# ── Download buttons ──────────────────────────────────────────────────────────
base_name = uploaded_file.name.rsplit(".", 1)[0]
dl1, dl2 = st.columns(2)
with dl1:
    st.download_button(
        label="📥 Download Audit Error Report",
        data=audit_bytes,
        file_name=f"ERROR_REPORT_{base_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="secondary",
    )
with dl2:
    st.download_button(
        label="📥 Download Sanitized Workbook",
        data=clean_bytes,
        file_name=f"CLEANED_{base_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

# ── Data preview ──────────────────────────────────────────────────────────────
st.write("### 📋 Sanitized Sheet Preview (first 50 rows)")
first_sname = list(sanitized_sheets.keys())[0]
st.dataframe(sanitized_sheets[first_sname].head(50), use_container_width=True)
