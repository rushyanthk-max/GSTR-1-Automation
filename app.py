import streamlit as st
import pandas as pd
import re
import io

# Set up clean browser tab and layout
st.set_page_config(page_title="BCPL Universal GST Sanitizer & Auditor", layout="centered")

st.title("📦 BCPL Universal E-commerce GST Sanitizer")
st.write("Upload your files to clean multi-sheet workbooks and generate a raw-data side-by-side Audit Error Report.")

# Robust helper function to handle column strings cleanup formatting safely
def clean_df_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    
    # Skip any trailing completely empty or unnamed edge spacers if found
    hsn_keywords = ['hsn', 'sac', 'commodity', 'nomenclature']
    sku_keywords = ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article']
    
    has_hsn = any(any(k in str(c).lower() for k in hsn_keywords) for c in df.columns)
    has_sku = any(any(k in str(c).lower() for k in sku_keywords) for c in df.columns)
    
    if not (has_hsn or has_sku):
        for idx, row in df.iterrows():
            row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            if any(v in ['hsn', 'hsn code', 'hsn/sac', 'sku', 'seller sku', 'fsn'] for v in row_vals):
                df.columns = [str(v).strip() for v in row.values]
                df = df.iloc[idx + 1:].reset_index(drop=True)
                break
    return df

# FLUENT DEEP CLEAN LAYER FOR SKUS
def deep_clean_sku(sku_val):
    if pd.isna(sku_val):
        return ""
    s = str(sku_val).strip().lower()
    s = re.sub(r'^[`"\'\s]+|[`"\'\s]+$', '', s)
    s = s.replace('"', '').replace("'", "").replace("`", "")
    if s.startswith("sku:"):
        s = s[4:]
    elif s.startswith("sku"):
        s = s[3:]
    s = re.sub(r'^[^a-z0-9]+', '', s)
    s = re.sub(r'[^a-z0-9]', '', s)
    return s

# EXTRACT NUMERIC % WHOLE VALUES FROM RAW STRINGS
def extract_rate_number(val):
    if pd.isna(val) or str(val).strip() in ['', 'nan', 'None', '<NA>']: return 0.0
    s = str(val).strip().replace('%', '')
    s = re.sub(r'\.0+$', '', s)
    if s in ['0.018', '0.005', '0.012', '0.028']: s = str(float(s) * 10)
    digits = "".join(c for c in s if c.isdigit() or c == '.')
    try: 
        num = float(digits) if digits else 0.0
        if num > 0 and num <= 1.0: num = num * 100
        return num
    except: return 0.0

# =========================================================================
# 1. DUAL FILE UPLOADER COMPONENTS
# =========================================================================
st.subheader("1️⃣ Step 1: Upload Raw Transaction Report")
uploaded_file = st.file_uploader("Drop your sales sheet workbook here", type=["xlsx", "xls", "csv"], key="sales_report")

st.subheader("2️⃣ Step 2: Upload Master Product Attribute / Catalog File (Optional)")
attribute_file = st.file_uploader("Drop your Master Item Catalog sheet here to enable SKU/Tax audits and auto-healing", type=["xlsx", "xls", "csv"], key="attribute_sheet")

if uploaded_file:
    # Dictionary tracking raw dataframes extracted per tab sheet
    raw_sheets_dict = {}
    
    if uploaded_file.name.endswith(('.xlsx', '.xls')):
        excel_file = pd.ExcelFile(uploaded_file)
        for sheet in excel_file.sheet_names:
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet, dtype=str)
            raw_sheets_dict[sheet] = clean_df_columns(df_sheet)
    else:
        df_csv = pd.read_csv(uploaded_file, dtype=str, low_memory=False)
        raw_sheets_dict["Sales Report"] = clean_df_columns(df_csv)

    # 3. BUILD CENTRALIZED INTER-SHEET MATCH REFERENCE LOOKUP
    sales_lookup_df = pd.DataFrame(columns=['Order ID', 'Order Item ID', 'SKU', 'HSN Code'])
    for sheet_name, df_s in raw_sheets_dict.items():
        cols_low = [str(c).lower() for c in df_s.columns]
        if 'order id' in cols_low and 'order item id' in cols_low and any('sku' in c for c in cols_low):
            o_col = [c for c in df_s.columns if str(c).lower() == 'order id'][0]
            oi_col = [c for c in df_s.columns if str(c).lower() == 'order item id'][0]
            s_col = [c for c in df_s.columns if 'sku' in str(c).lower()][0]
            h_col = [c for c in df_s.columns if 'hsn' in str(c).lower() or 'sac' in str(c).lower()][0] if any('hsn' in c or 'sac' in c for c in cols_low) else None
            
            if h_col:
                subset = df_s[[o_col, oi_col, s_col, h_col]].dropna(subset=[o_col, oi_col])
                subset.columns = ['Order ID', 'Order Item ID', 'SKU', 'HSN Code']
                sales_lookup_df = pd.concat([sales_lookup_df, subset], ignore_index=True)
    sales_lookup_df.drop_duplicates(subset=['Order ID', 'Order Item ID'], inplace=True)

    # 4. MAP MASTER PRODUCT CATALOG DICTIONARIES
    master_sku_hsn_map = {}
    master_sku_tax_map = {}
    master_hsn_tax_map = {}

    if attribute_file:
        if attribute_file.name.endswith(('.xlsx', '.xls')):
            attr_raw = pd.read_excel(attribute_file, dtype=str)
        else:
            attr_raw = pd.read_csv(attribute_file, dtype=str, low_memory=False)
        attr_df = clean_df_columns(attr_raw)
        
        attr_hsn_col, attr_sku_col, attr_tax_col = None, None, None
        for c in attr_df.columns:
            cl = str(c).strip().lower()
            if cl in ['sku', 'seller-sku', 'item-code', 'article-code', 'product-sku']: attr_sku_col = c; break
        if not attr_sku_col:
            for c in attr_df.columns:
                cl = str(c).strip().lower()
                if 'sku' in cl and not any(x in cl for x in ['type', 'parent', 'accounting']): attr_sku_col = c; break

        for c in attr_df.columns:
            cl = str(c).strip().lower()
            if cl in ['hsn', 'hsn/sac', 'hsn_sac', 'hsncode', 'hsn code', 'hsn_code', 'commoditycode', 'producttaxcode']: attr_hsn_col = c; break
        if not attr_hsn_col:
            for c in attr_df.columns:
                cl = str(c).strip().lower()
                if any(k in cl for k in ['hsn', 'sac', 'taxcode', 'commodity', 'nomenclature']): attr_hsn_col = c; break

        for c in attr_df.columns:
            cl = str(c).strip().lower()
            if any(k in cl for k in ['producttaxrule', 'tax rule', 'tax_rule', 'gst rate', 'tax percentage']): attr_tax_col = c; break
        
        if attr_hsn_col and attr_sku_col:
            for _, row in attr_df.iterrows():
                r_hsn = str(row[attr_hsn_col]).strip() if pd.notna(row[attr_hsn_col]) else ""
                r_sku_clean = deep_clean_sku(row[attr_sku_col])
                r_tax_raw = str(row[attr_tax_col]).strip() if attr_tax_col and pd.notna(row[attr_tax_col]) else ""
                
                if r_hsn.startswith('="') and r_hsn.endswith('"'): r_hsn = r_hsn[2:-1]
                r_digits = "".join(filter(str.isdigit, re.sub(r'[\s\-\.\/]', '', r_hsn)))
                if len(r_digits) == 7: r_digits = "0" + r_digits
                r_tax_digits = "".join(filter(str.isdigit, r_tax_raw))
                
                if r_digits and r_digits.lower() not in ["", "nan", "none"] and r_sku_clean:
                    master_sku_hsn_map[r_sku_clean] = r_digits
                    if r_tax_digits:
                        master_sku_tax_map[r_sku_clean] = r_tax_digits
                        master_hsn_tax_map[r_digits] = r_tax_digits

    # =========================================================================
    # 🕵️‍♂️ COMPUTE COMPLETE RAW AUDIT RECORDS AND COMPILE ERROR REPORT
    # =========================================================================
    global_raw_records = []
    
    for sheet_name, df_s in raw_sheets_dict.items():
        # Identify headers for current sheet
        sh_hsn_col, sh_sku_col, sh_cgst_col, sh_sgst_col, sh_igst_col, sh_tx_type_col = None, None, None, None, None, None
        sh_oid_col, sh_oiid_col = None, None

        for col in df_s.columns:
            c_low = str(col).strip().lower()
            if c_low == 'order id': sh_oid_col = col
            if c_low == 'order item id': sh_oiid_col = col
            if c_low in ['sku', 'seller-sku', 'item-code', 'article-code', 'wms_code']: sh_sku_col = col
            if c_low in ['hsn', 'hsn/sac', 'hsn_sac', 'hsncode', 'hsn code', 'hsn_code', 'commodity', 'hsn sac']: sh_hsn_col = col

        if not sh_sku_col:
            for col in df_s.columns:
                if any(k in str(col).lower() for k in ['sku', 'fsn', 'seller-sku', 'product-id', 'article']): sh_sku_col = col; break
        if not sh_hsn_col:
            for col in df_s.columns:
                if any(k in str(col).lower() for k in ['hsn', 'sac', 'commodity', 'nomenclature']): sh_hsn_col = col; break

        for col in df_s.columns:
            c_low = str(col).strip().lower()
            if any(k in c_low for k in ['transaction type', 'type', 'status', 'order status', 'transaction_type', 'order_status', 'event type', 'event_type', 'document type']):
                sh_tx_type_col = col; break

        # 🎯 CRITICAL STABILIZATION: Explicitly and exclusively fetch IGST, CGST, SGST Rates
        for col in df_s.columns:
            c_low = str(col).strip().lower()
            if any(x in c_low for x in ['tcs', 'shipping', 'gift', 'wrap', 'delivery', 'postage', 'cst', 'vat', 'cess', 'tds', 'amount', 'amt', 'value']):
                continue
            if 'cgst' in c_low and 'rate' in c_low: sh_cgst_col = col
            if ('sgst' in c_low or 'utgst' in c_low) and 'rate' in c_low: sh_sgst_col = col
            if 'igst' in c_low and 'rate' in c_low: sh_igst_col = col

        # Build raw details row by row for current sheet
        for index, row in df_s.iterrows():
            # Handle cross-tab lookup if sheet lacks HSN or SKU directly
            rhsn_field, rsku_field = "", ""
            if sh_hsn_col: rhsn_field = str(row[sh_hsn_col]).strip() if pd.notna(row[sh_hsn_col]) else ""
            if sh_sku_col: rsku_field = str(row[sh_sku_col]).strip() if pd.notna(row[sh_sku_col]) else ""
            
            if sh_oid_col and sh_oiid_col and (not rhsn_field or not rsku_field):
                oid_val = str(row[sh_oid_col]).strip() if pd.notna(row[sh_oid_col]) else ""
                oiid_val = str(row[sh_oiid_col]).strip() if pd.notna(row[sh_oiid_col]) else ""
                match_ref = sales_lookup_df[(sales_lookup_df['Order ID'] == oid_val) & (sales_lookup_df['Order Item ID'] == oiid_val)]
                if not match_ref.empty:
                    if not rsku_field: rsku_field = str(match_ref['SKU'].values[0])
                    if not rhsn_field: rhsn_field = str(match_ref['HSN Code'].values[0])

            if rhsn_field.startswith('="') and rhsn_field.endswith('"'): rhsn_field = rhsn_field[2:-1]
            hsn_digits = "".join(filter(str.isdigit, re.sub(r'[\s\-\.\/]', '', rhsn_field)))
            if len(hsn_digits) == 7: hsn_digits = "0" + hsn_digits

            sku_disp = re.sub(r'^["\'`\s]+|["\'`\s]+$', '', rsku_field)
            if sku_disp.startswith("SKU:"): sku_disp = sku_disp[4:]

            cgst_v = extract_rate_number(row[sh_cgst_col]) if sh_cgst_col else 0.0
            sgst_v = extract_rate_number(row[sh_sgst_col]) if sh_sgst_col else 0.0
            igst_v = extract_rate_number(row[sh_igst_col]) if sh_igst_col else 0.0
            
            total_math = igst_v if igst_v > 0 else (cgst_v + sgst_v)
            rate_str = str(int(total_math)) if total_math.is_integer() else str(total_math)
            tx_status = str(row[sh_tx_type_col]).strip().lower() if sh_tx_type_col else ""

            global_raw_records.append({
                "sheet": sheet_name,
                "index": index,
                "sku_display": sku_disp,
                "clean_sku": deep_clean_sku(rsku_field),
                "raw_hsn_digits": hsn_digits,
                "raw_rate_str": rate_str,
                "tx_status": tx_status
            })

    # Group by raw HSN digits across entire workbook to pinpoint dual tax brackets
    hsn_to_rates_dict = {}
    for r in global_raw_records:
        h = r["raw_hsn_digits"]
        rt = r["raw_rate_str"]
        if h and h != "MISSING HSN" and rt and rt != "0" and rt != "0.0" and rt != "":
            if h not in hsn_to_rates_dict: hsn_to_rates_dict[h] = set()
            hsn_to_rates_dict[h].add(rt)
    double_rate_raw_hsns = {h: ",".join(sorted(list(v))) for h, v in hsn_to_rates_dict.items() if len(v) > 1}

    # Populate 6-category error tracking lists based on raw details
    list_missing_hsn = []
    list_double_rates = []
    list_invalid_lengths = []
    list_wrong_tax_hsn = []
    list_wrong_hsn_sku = []
    list_wrong_tax_sku = []

    for r in global_raw_records:
        h = r["raw_hsn_digits"]
        rt = r["raw_rate_str"]
        sd = r["sku_display"]
        cs = r["clean_sku"]
        ts = r["tx_status"]

        if not h and "cancel" not in ts:
            if sd and sd not in list_missing_hsn: list_missing_hsn.append(sd)

        if h in double_rate_raw_hsns:
            entry = {"hsn": h, "sku": sd, "rates": double_rate_raw_hsns[h]}
            if entry not in list_double_rates: list_double_rates.append(entry)

        if h and len(h) not in [6, 8]:
            if h not in list_invalid_lengths: list_invalid_lengths.append(h)

        if h in master_hsn_tax_map:
            m_tax = master_hsn_tax_map[h]
            if rt != "0" and rt != "" and rt != m_tax:
                entry = {"hsn": h, "rate": rt, "correct": m_tax}
                if entry not in list_wrong_tax_hsn: list_wrong_tax_hsn.append(entry)

        if cs in master_sku_hsn_map:
            m_hsn = master_sku_hsn_map[cs]
            if h and h != m_hsn:
                entry = {"sku": sd, "wrong_hsn": h, "correct_hsn": m_hsn}
                if entry not in list_wrong_hsn_sku: list_wrong_hsn_sku.append(entry)

        if cs in master_sku_tax_map:
            m_sku_tax = master_sku_tax_map[cs]
            if rt != "0" and rt != "" and rt != m_sku_tax:
                entry = {"sku": sd, "rate": rt, "correct": m_sku_tax}
                if entry not in list_wrong_tax_sku: list_wrong_tax_sku.append(entry)

    # =========================================================================
    # 📥 ASSEMBLE UNIFIED DASHBOARD SUMMARY (Side-by-Side Excel View)
    # =========================================================================
    max_len = max(len(list_missing_hsn), len(list_double_rates), len(list_invalid_lengths), 
                  len(list_wrong_tax_hsn), len(list_wrong_hsn_sku), len(list_wrong_tax_sku), 1)

    audit_data = {
        "Missing HSN Codes (SKUs)": [list_missing_hsn[i] if i < len(list_missing_hsn) else "" for i in range(max_len)],
        "Invalid Length HSNs (Not 6 or 8)": [list_invalid_lengths[i] if i < len(list_invalid_lengths) else "" for i in range(max_len)],
        "Double Rate - HSN": [list_double_rates[i]["hsn"] if i < len(list_double_rates) else "" for i in range(max_len)],
        "Double Rate - SKU": [list_double_rates[i]["sku"] if i < len(list_double_rates) else "" for i in range(max_len)],
        "Double Rate - Tax Rates": [list_double_rates[i]["rates"] if i < len(list_double_rates) else "" for i in range(max_len)],
        "Wrong Tax by HSN - HSN": [list_wrong_tax_hsn[i]["hsn"] if i < len(list_wrong_tax_hsn) else "" for i in range(max_len)],
        "Wrong Tax by HSN - Input Rate": [list_wrong_tax_hsn[i]["rate"] if i < len(list_wrong_tax_hsn) else "" for i in range(max_len)],
        "Wrong Tax by HSN - Master Rate": [list_wrong_tax_hsn[i]["correct"] if i < len(list_wrong_tax_hsn) else "" for i in range(max_len)],
        "Wrong HSN by SKU - SKU": [list_wrong_hsn_sku[i]["sku"] if i < len(list_wrong_hsn_sku) else "" for i in range(max_len)],
        "Wrong HSN by SKU - Input HSN": [list_wrong_hsn_sku[i]["wrong_hsn"] if i < len(list_wrong_hsn_sku) else "" for i in range(max_len)],
        "Wrong HSN by SKU - Master HSN": [list_wrong_hsn_sku[i]["correct_hsn"] if i < len(list_wrong_hsn_sku) else "" for i in range(max_len)],
        "Incorrect GST Rates - SKU": [list_wrong_tax_sku[i]["sku"] if i < len(list_wrong_tax_sku) else "" for i in range(max_len)],
        "Incorrect GST Rates - Input Rate": [list_wrong_tax_sku[i]["rate"] if i < len(list_wrong_tax_sku) else "" for i in range(max_len)],
        "Incorrect GST Rates - Master Correct Rate": [list_wrong_tax_sku[i]["correct"] if i < len(list_wrong_tax_sku) else "" for i in range(max_len)]
    }
    df_audit_report = pd.DataFrame(audit_data)

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df_audit_report.to_excel(writer, sheet_name='HSN_GST_Audit_Dashboard', index=False)
    excel_buffer.seek(0)
    excel_binary_data = excel_buffer.getvalue()

    # =========================================================================
    # 🚀 RUN FINAL PRODUCTION SANITIZATION LAYER (APPLY VOTE WINNERS)
    # =========================================================================
    excel_clean_buffer = io.BytesIO()
    
    with pd.ExcelWriter(excel_clean_buffer, engine='xlsxwriter') as writer_clean:
        for sheet_name, df_s in raw_sheets_dict.items():
            # RE-ISOLATE TARGETS FOR CURRENT SHEET FOR ACCURATE MUTATION
            sh_hsn_col, sh_sku_col, sh_cgst_col, sh_sgst_col, sh_igst_col = None, None, None, None, None
            for col in df_s.columns:
                c_low = str(col).strip().lower()
                if c_low in ['sku', 'seller-sku', 'item-code', 'article-code', 'wms_code']: sh_sku_col = col; break
            if not sh_sku_col:
                for col in df_s.columns:
                    if any(k in str(col).lower() for k in ['sku', 'fsn', 'seller-sku', 'product-id', 'article']): sh_sku_col = col; break

            for col in df_s.columns:
                c_low = str(col).strip().lower()
                if c_low in ['hsn', 'hsn/sac', 'hsn_sac', 'hsncode', 'hsn code', 'hsn_code', 'commodity', 'hsn sac']: sh_hsn_col = col; break
            if not sh_hsn_col:
                for col in df_s.columns:
                    if any(k in str(col).lower() for k in ['hsn', 'sac', 'commodity', 'nomenclature']): sh_hsn_col = col; break

            for col in df_s.columns:
                c_low = str(col).strip().lower()
                if any(x in c_low for x in ['tcs', 'shipping', 'gift', 'wrap', 'delivery', 'postage', 'cst', 'vat', 'cess', 'tds', 'amount', 'amt', 'value']): continue
                if 'cgst' in c_low and 'rate' in c_low: sh_cgst_col = col
                if ('sgst' in c_low or 'utgst' in c_low) and 'rate' in c_low: sh_sgst_col = col
                if 'igst' in c_low and 'rate' in c_low: sh_igst_col = col

            sheet_records = [r for r in global_raw_records if r["sheet"] == sheet_name]
            
            sheet_healed_hsns = []
            for r in sheet_records:
                if not r["raw_hsn_digits"]:
                    if r["clean_sku"] in master_sku_hsn_map: sheet_healed_hsns.append(master_sku_hsn_map[r["clean_sku"]])
                    else: sheet_healed_hsns.append("MISSING HSN")
                else:
                    sheet_healed_hsns.append(r["raw_hsn_digits"])

            df_s['_temp_hsn_healed'] = sheet_healed_hsns

            sheet_majority_tax_map = {}
            for h_val, group in df_s.groupby('_temp_hsn_healed'):
                if h_val != "MISSING HSN":
                    g_indices = group.index
                    rates_series = pd.Series([sheet_records[df_s.index.get_loc(x)]["raw_rate_str"] for x in g_indices])
                    rates_series = rates_series[(rates_series != "") & (rates_series != "0") & (rates_series != "0.0")]
                    if not rates_series.empty:
                        sheet_majority_tax_map[h_val] = rates_series.value_counts().index[0]

            final_shielded_hsns = []
            final_total_rates = []

            for index, row in df_s.iterrows():
                loc_idx = df_s.index.get_loc(index)
                h_healed = sheet_healed_hsns[loc_idx]
                
                if h_healed in sheet_majority_tax_map: final_winner = sheet_majority_tax_map[h_healed]
                else: final_winner = sheet_records[loc_idx]["raw_rate_str"]
                
                final_shielded_hsns.append(f'="{h_healed}"' if h_healed != "MISSING HSN" else "MISSING HSN")
                final_total_rates.append(final_winner)

            if sh_hsn_col: df_s[sh_hsn_col] = final_shielded_hsns
            df_s['Total Tax Rate'] = final_total_rates

            for index, row in df_s.iterrows():
                loc_idx = df_s.index.get_loc(index)
                winner = final_total_rates[loc_idx]
                try:
                    w_num = float(winner)
                    orig_igst_str = str(df_s.loc[index, sh_igst_col]).strip() if sh_igst_col else '0'
                    
                    if sh_igst_col and orig_igst_str not in ['', '0', '0.0']:
                        df_s.loc[index, sh_igst_col] = winner
                        if sh_cgst_col: df_s.loc[index, sh_cgst_col] = '0'
                        if sh_sgst_col: df_s.loc[index, sh_sgst_col] = '0'
                    else:
                        split_str = str(int(w_num / 2)) if (w_num / 2).is_integer() else str(w_num / 2)
                        if sh_cgst_col: df_s.loc[index, sh_cgst_col] = split_str
                        if sh_sgst_col: df_s.loc[index, sh_sgst_col] = split_str
                        if sh_igst_col: df_s.loc[index, sh_igst_col] = '0'
                except: pass

            df_s.drop(columns=['_temp_hsn_healed'], inplace=True, errors='ignore')
            df_s.to_excel(writer_clean, sheet_name=sheet_name, index=False)
            
    excel_clean_buffer.seek(0)
    excel_clean_binary = excel_clean_buffer.getvalue()

    # =========================================================================
    # 🎨 RENDER INTERFACE SUCCESS DASHBOARD
    # =========================================================================
    st.success(f"✨ Multi-sheet Workbook Analysis Complete! Successfully parsed uploaded reports.")
    
    st.error("⚠️ COMPLIANCE RISK AUDIT REPORT DISCOVERED!")
    st.download_button(
        label="📥 Download Unified Side-by-Side Error Report",
        data=excel_binary_data,
        file_name=f"ERROR_REPORT_{uploaded_file.name.split('.')[0]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("✅ SANITIZATION PACKET WORKBOOK COMPLETED")
    st.download_button(
        label="📥 Download Sanitized Sales & CashBack Workbook",
        data=excel_clean_binary,
        file_name=f"CLEANED_WORKBOOK_{uploaded_file.name.split('.')[0]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.write("### Sanitized Master Report Grid Preview:")
    first_sheet_name = list(raw_sheets_dict.keys())[0]
    st.dataframe(raw_sheets_dict[first_sheet_name].head(50))
