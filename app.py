import streamlit as st
import pandas as pd
import re
import io

# Set up clean browser tab and layout
st.set_page_config(page_title="BCPL Universal GST Sanitizer & Auditor", layout="centered")

st.title("📦 BCPL Universal E-commerce GST Sanitizer")
st.write("Upload your files to instantly clean transaction sheets and generate your unified multi-column Error Report based on raw data.")

# Robust helper function to handle dynamic data loads safely without accidental row corruption
def load_data_safely(file_obj):
    try:
        if file_obj.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_obj, dtype=str)
        else:
            df = pd.read_csv(file_obj, dtype=str, low_memory=False)
        
        # Clean column names formatting whitespace
        df.columns = [str(c).strip() for c in df.columns]
        
        # Check if HSN or SKU indicators are already found right out of the box
        hsn_keywords = ['hsn', 'sac', 'commodity', 'nomenclature']
        sku_keywords = ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article']
        
        has_hsn = any(any(k in str(c).lower() for k in hsn_keywords) for c in df.columns)
        has_sku = any(any(k in str(c).lower() for k in sku_keywords) for c in df.columns)
        
        # 🎯 CRITICAL FIX: Only run row scanning heuristics if we TRULY cannot find HSN or SKU headers
        if not (has_hsn or has_sku):
            for idx, row in df.iterrows():
                row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                if any(v in ['hsn', 'hsn code', 'hsn/sac', 'sku', 'seller sku', 'fsn'] for v in row_vals):
                    df.columns = [str(v).strip() for v in row.values]
                    df = df.iloc[idx + 1:].reset_index(drop=True)
                    break
        return df
    except Exception as e:
        st.error(f"Error reading file '{file_obj.name}': {str(e)}")
        return None

# FLUENT DEEP CLEAN LAYER FOR SKUS (Insulated against Flipkart triple-quotes and 'SKU:' labels)
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

# =========================================================================
# 1. DUAL FILE UPLOADER COMPONENTS
# =========================================================================
st.subheader("1️⃣ Step 1: Upload Raw Transaction Report")
uploaded_file = st.file_uploader("Drop your sales report here (Amazon, Flipkart, Nykaa, Meesho, JioMart, WMS)", type=["xlsx", "xls", "csv"], key="sales_report")

st.subheader("2️⃣ Step 2: Upload Master Product Attribute / Catalog File (Optional)")
attribute_file = st.file_uploader("Drop your Master Item Catalog sheet here to enable SKU/Tax audits and auto-healing", type=["xlsx", "xls", "csv"], key="attribute_sheet")

if uploaded_file:
    df_raw = load_data_safely(uploaded_file)
    
    if df_raw is not None:
        df = df_raw.copy()
        initial_rows = len(df)
        df.dropna(how='all', inplace=True)
        blank_rows = initial_rows - len(df)

        # 2. SMART UNIVERSAL KEYWORD COLUMN SCANNER (Insulated for Flipkart layouts)
        hsn_col, sku_col, cgst_col, sgst_col, igst_col, tx_type_col = None, None, None, None, None, None
        
        for col in df.columns:
            c_low = str(col).strip().lower()
            if c_low in ['sku', 'seller-sku', 'item-code', 'article-code', 'wms_code']:
                sku_col = col
                break
        if not sku_col:
            for col in df.columns:
                c_low = str(col).strip().lower()
                if any(k in c_low for k in ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article', 'wms_code', 'fsn code']):
                    sku_col = col
                    break

        for col in df.columns:
            c_low = str(col).strip().lower()
            if c_low in ['hsn', 'hsn/sac', 'hsn_sac', 'hsncode', 'hsn code', 'hsn_code', 'commodity', 'hsn sac', 'commodity code', 'commodity_code']:
                hsn_col = col
                break
        if not hsn_col:
            for col in df.columns:
                c_low = str(col).strip().lower()
                if any(k in c_low for k in ['hsn', 'sac', 'commodity', 'nomenclature', 'hsn/sac', 'hsn code', 'hsn_sac', 'hsn_code', 'hsn sac']):
                    hsn_col = col
                    break

        for col in df.columns:
            c_low = str(col).strip().lower()
            if any(k in c_low for k in ['transaction type', 'type', 'status', 'order status', 'transaction_type', 'order_status', 'event type', 'event_type']):
                tx_type_col = col
                break

        for col in df.columns:
            c_low = str(col).strip().lower()
            if any(x in c_low for x in ['gift', 'wrap', 'shipping', 'delivery', 'ship', 'postage', 'tcs', 'amount', 'amt', 'value', 'tax paid', 'tax collected']):
                continue
            if 'cgst' in c_low: cgst_col = col
            if 'sgst' in c_low: sgst_col = col
            if 'igst' in c_low: igst_col = col

        # 3. EXECUTE RECONCILIATION AND REPORT BUILDING
        if hsn_col:
            df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
            if sku_col:
                df[sku_col] = df[sku_col].fillna("").astype(str).str.strip()

            # MASTER CATALOG MAPS
            master_sku_hsn_map = {}
            master_sku_tax_map = {}
            master_hsn_tax_map = {}

            if attribute_file:
                attr_df = load_data_safely(attribute_file)
                if attr_df is not None:
                    attr_hsn_col, attr_sku_col, attr_tax_col = None, None, None
                    
                    for c in attr_df.columns:
                        cl = str(c).strip().lower()
                        if cl in ['sku', 'seller-sku', 'item-code', 'article-code', 'product-sku']:
                            attr_sku_col = c
                            break
                    if not attr_sku_col:
                        for c in attr_df.columns:
                            cl = str(c).strip().lower()
                            if 'sku' in cl and not any(x in cl for x in ['type', 'parent', 'accounting']):
                                attr_sku_col = c
                                break

                    for c in attr_df.columns:
                        cl = str(c).strip().lower()
                        if cl in ['hsn', 'hsn/sac', 'hsn_code', 'hsncode', 'hsn code', 'commoditycode', 'producttaxcode']:
                            attr_hsn_col = c
                            break
                    if not attr_hsn_col:
                        for c in attr_df.columns:
                            cl = str(c).strip().lower()
                            if any(k in cl for k in ['hsn', 'sac', 'taxcode', 'commodity', 'nomenclature', 'hsn code']):
                                attr_hsn_col = c
                                break

                    for c in attr_df.columns:
                        cl = str(c).strip().lower()
                        if any(k in cl for k in ['producttaxrule', 'tax rule', 'tax_rule', 'gst rate', 'tax percentage']):
                            attr_tax_col = c
                            break
                    
                    if attr_hsn_col and attr_sku_col:
                        for _, row in attr_df.iterrows():
                            r_hsn = str(row[attr_hsn_col]).strip() if pd.notna(row[attr_hsn_col]) else ""
                            r_sku_clean = deep_clean_sku(row[attr_sku_col])
                            r_tax_raw = str(row[attr_tax_col]).strip() if attr_tax_col and pd.notna(row[attr_tax_col]) else ""
                            
                            if r_hsn.startswith('="') and r_hsn.endswith('"'): r_hsn = r_hsn[2:-1]
                            r_clean = re.sub(r'[\s\-\.\/]', '', r_hsn)
                            r_digits = "".join(filter(str.isdigit, r_clean))
                            if len(r_digits) == 7: r_digits = "0" + r_digits
                            r_tax_digits = "".join(filter(str.isdigit, r_tax_raw))
                            
                            if r_digits and r_digits.lower() not in ["", "nan", "none"] and r_sku_clean:
                                master_sku_hsn_map[r_sku_clean] = r_digits
                                if r_tax_digits:
                                    master_sku_tax_map[r_sku_clean] = r_tax_digits
                                    master_hsn_tax_map[r_digits] = r_tax_digits
            
            if sku_col:
                for _, row in df.iterrows():
                    raw_hsn = str(row[hsn_col]).strip() if pd.notna(row[hsn_col]) else ""
                    raw_sku_clean = deep_clean_sku(row[sku_col])
                    if raw_hsn.startswith('="') and raw_hsn.endswith('"'): raw_hsn = raw_hsn[2:-1]
                    clean_hsn = "".join(filter(str.isdigit, re.sub(r'[\s\-\.\/]', '', raw_hsn)))
                    if len(clean_hsn) == 7: clean_hsn = "0" + clean_hsn
                    if clean_hsn and clean_hsn not in ["", "nan", "none"] and raw_sku_clean:
                        if raw_sku_clean not in master_sku_hsn_map:
                            master_sku_hsn_map[raw_sku_clean] = clean_hsn

            # PASS A: EXTRACT RAW UN-MODIFIED TAX RATES AND HSN DIGITS ROW BY ROW
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

            cgst_series = df[cgst_col].apply(extract_rate_number) if cgst_col else pd.Series(0.0, index=df.index)
            sgst_series = df[sgst_col].apply(extract_rate_number) if sgst_col else pd.Series(0.0, index=df.index)
            igst_series = df[igst_col].apply(extract_rate_number) if igst_col else pd.Series(0.0, index=df.index)

            raw_total_rates = []
            raw_hsn_codes = []
            raw_sku_displays = []
            
            for index, row in df.iterrows():
                val = str(row[hsn_col]).strip() if pd.notna(row[hsn_col]) else ""
                if val.startswith('="') and val.endswith('"'): val = val[2:-1]
                val_clean = re.sub(r'[\s\-\.\/]', '', val)
                hsn_digits = "".join(filter(str.isdigit, val_clean))
                if len(hsn_digits) == 7: hsn_digits = "0" + hsn_digits
                raw_hsn_codes.append(hsn_digits)

                raw_sku_val = str(row[sku_col]).strip() if pd.notna(row[sku_col]) else ""
                raw_sku_display = re.sub(r'^["\'`]+|["\'`]+$', '', raw_sku_val)
                raw_sku_displays.append(raw_sku_display)
                
                if igst_series[index] > 0: total_math = igst_series[index]
                else: total_math = cgst_series[index] + sgst_series[index]
                raw_total_rates.append(str(int(total_math)) if total_math.is_integer() else str(total_math))

            df_raw['Calculated Total Tax Rate'] = raw_total_rates

            # PASS B: COMPUTE DUAL TAX RATE BRACKETS FROM THE RAW REPORT
            hsn_to_rates_map = {}
            for idx, h_code in enumerate(raw_hsn_codes):
                rt = raw_total_rates[idx]
                if h_code and rt and rt != "0" and rt != "0.0" and rt != "":
                    if h_code not in hsn_to_rates_map: hsn_to_rates_map[h_code] = set()
                    hsn_to_rates_map[h_code].add(rt)
            
            double_rate_raw_hsns = {h: ",".join(sorted(list(v))) for h, v in hsn_to_rates_map.items() if len(v) > 1}

            # =========================================================================
            # 🕵️‍♂️ RUN AUDIT CHECKS EXCLUSIVELY ON UN-MUTATED RAW DATA ARRAYS
            # =========================================================================
            list_missing_hsn = []
            list_double_rates = []
            list_invalid_lengths = []
            list_wrong_tax_hsn = []
            list_wrong_hsn_sku = []
            list_wrong_tax_sku = []

            for index, row in df.iterrows():
                loc_idx = df.index.get_loc(index)
                rhsn = raw_hsn_codes[loc_idx]
                rrate = raw_total_rates[loc_idx]
                rsku_disp = raw_sku_displays[loc_idx]
                csku = deep_clean_sku(str(row[sku_col]))
                tx_status = str(row[tx_type_col]).strip().lower() if tx_type_col else ""

                # Error 1: Missing HSN (No Cancelled lines)
                if not rhsn and "cancel" not in tx_status:
                    if rsku_disp and rsku_disp not in list_missing_hsn:
                        list_missing_hsn.append(rsku_disp)

                # Error 2: Double Tax Rates
                if rhsn in double_rate_raw_hsns:
                    entry = {"hsn": rhsn, "sku": rsku_disp, "rates": double_rate_raw_hsns[rhsn]}
                    if entry not in list_double_rates:
                        list_double_rates.append(entry)

                # Error 3: Invalid Digit Lengths (Not 6 or 8)
                if rhsn and len(rhsn) not in [6, 8]:
                    if rhsn not in list_invalid_lengths:
                        list_invalid_lengths.append(rhsn)

                # Error 4: Wrong Tax Rates by HSN Mapping
                if rhsn in master_hsn_tax_map:
                    m_tax = master_hsn_tax_map[rhsn]
                    if rrate != "0" and rrate != "" and rrate != m_tax:
                        entry = {"hsn": rhsn, "rate": rrate, "correct": m_tax}
                        if entry not in list_wrong_tax_hsn:
                            list_wrong_tax_hsn.append(entry)

                # Error 5: Wrong HSN Code by SKU mapping
                if csku in master_sku_hsn_map:
                    m_hsn = master_sku_hsn_map[csku]
                    if rhsn and rhsn != m_hsn:
                        entry = {"sku": rsku_disp, "wrong_hsn": rhsn, "correct_hsn": m_hsn}
                        if entry not in list_wrong_hsn_sku:
                            list_wrong_hsn_sku.append(entry)

                # Error 6: Wrong Tax Rate by SKU mapping
                if csku in master_sku_tax_map:
                    m_sku_tax = master_sku_tax_map[csku]
                    if rrate != "0" and rrate != "" and rrate != m_sku_tax:
                        entry = {"sku": rsku_disp, "rate": rrate, "correct": m_sku_tax}
                        if entry not in list_wrong_tax_sku:
                            list_wrong_tax_sku.append(entry)

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
            healed_hsns = []
            healed_total_rates = []

            # Step 1: Recover missing codes based on master directory values first
            for idx, h_code in enumerate(raw_hsn_codes):
                csku = deep_clean_sku(df.loc[idx, sku_col]) if sku_col else ""
                if not h_code:
                    if csku in master_sku_hsn_map: healed_hsns.append(master_sku_hsn_map[csku])
                    else: healed_hsns.append("MISSING HSN")
                else:
                    healed_hsns.append(h_code)

            df['_temp_hsn_healed'] = healed_hsns

            # Step 2: Establish commercial vote winners based on healed groups
            healed_hsn_majority_map = {}
            for h_val, group in df.groupby('_temp_hsn_healed'):
                if h_val != "MISSING HSN":
                    g_indices = group.index
                    rates_in_g = pd.Series([raw_total_rates[df.index.get_loc(x)] for x in g_indices])
                    rates_in_g = rates_in_g[(rates_in_g != "") & (rates_in_g != "0") & (rates_in_g != "0.0")]
                    if not rates_in_g.empty:
                        healed_hsn_majority_map[h_val] = rates_in_g.value_counts().index[0]

            # Step 3: Enforce final records and balance out sub-component splits
            sanitized_hsns = []
            for index, row in df.iterrows():
                loc_idx = df.index.get_loc(index)
                h_healed = healed_hsns[loc_idx]
                
                if h_healed in healed_hsn_majority_map: final_winner_rate = healed_hsn_majority_map[h_healed]
                else: final_winner_rate = raw_total_rates[loc_idx]
                
                sanitized_hsns.append(f'="{h_healed}"' if h_healed != "MISSING HSN" else "MISSING HSN")
                healed_total_rates.append(final_winner_rate)

            df[hsn_col] = sanitized_hsns
            df['Total Tax Rate'] = healed_total_rates

            # Overwrite split columns matching the winner rate
            for index, row in df.iterrows():
                loc_idx = df.index.get_loc(index)
                winner = healed_total_rates[loc_idx]
                try:
                    w_num = float(winner)
                    orig_igst = str(df_raw.loc[index, igst_col]).strip() if igst_col in df_raw.columns else '0'
                    
                    if igst_col and orig_igst not in ['', '0', '0.0']:
                        df.loc[index, igst_col] = winner
                        if cgst_col: df.loc[index, cgst_col] = '0'
                        if sgst_col: df.loc[index, sgst_col] = '0'
                    else:
                        split_str = str(int(w_num / 2)) if (w_num / 2).is_integer() else str(w_num / 2)
                        if cgst_col: df.loc[index, cgst_col] = split_str
                        if sgst_col: df.loc[index, sgst_col] = split_str
                        if igst_col: df.loc[index, igst_col] = '0'
                except:
                    pass

            df.drop(columns=['_temp_hsn_healed'], inplace=True, errors='ignore')

            # =========================================================================
            # 🎨 RENDER INTERFACE SUCCESS DASHBOARD
            # =========================================================================
            st.success(f"✨ Data Analytics Scan Complete! Processed {initial_rows} lines successfully.")
            st.info(f"🎯 **Target Matrix Locked Successfully:** \n* **HSN Column Locked:** '{hsn_col}' \n* **SKU Column Locked:** '{sku_col}'")
            
            st.error("⚠️ COMPLIANCE RISK AUDIT REPORT DISCOVERED!")
            st.download_button(
                label="📥 Download Unified Side-by-Side Error Report",
                data=excel_binary_data,
                file_name=f"ERROR_REPORT_{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.success("✅ SANITIZATION PACKET COMPLETED SUCCESSFULLY")
            csv_data = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Sanitized Sales File for Repotic",
                data=csv_data,
                file_name=f"CLEANED_{uploaded_file.name.split('.')[0]}.csv",
                mime="text/csv"
            )

            st.write("### Sanitized Data Preview Grid:")
            st.dataframe(df.head(50))
        else:
            st.error("❌ Column Detection Error: The script could not automatically identify an HSN column name in your file.")
