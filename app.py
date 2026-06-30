import streamlit as st
import pandas as pd
import re
import io

# Set up clean browser tab and layout
st.set_page_config(page_title="BCPL Universal GST Sanitizer & Auditor", layout="centered")

st.title("📦 BCPL Universal E-commerce GST Sanitizer")
st.write("Upload your files to instantly clean transaction sheets and generate your unified multi-column Error Report.")

# Helper function to read any layout safely into raw text strings
def load_data_safely(file_obj):
    try:
        if file_obj.name.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_obj, dtype=str)
        else:
            return pd.read_csv(file_obj, dtype=str, low_memory=False)
    except Exception as e:
        st.error(f"Error reading file '{file_obj.name}': {str(e)}")
        return None

# AGGRESSIVE DEEP CLEAN LAYER FOR SKUS
def deep_clean_sku(sku_val):
    if pd.isna(sku_val):
        return ""
    s = str(sku_val).strip().lower()
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

        # 2. SMART UNIVERSAL KEYWORD COLUMN SCANNER
        hsn_col, sku_col, cgst_col, sgst_col, igst_col, tx_type_col = None, None, None, None, None, None
        
        for col in df.columns:
            c_low = str(col).strip().lower()
            if c_low in ['sku', 'seller-sku', 'item-code', 'article-code', 'wms_code']:
                sku_col = col
                break
        if not sku_col:
            for col in df.columns:
                c_low = str(col).strip().lower()
                if any(k in c_low for k in ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article', 'wms_code']):
                    sku_col = col
                    break

        for col in df.columns:
            c_low = str(col).strip().lower()
            if c_low in ['hsn', 'hsn/sac', 'hsn_sac', 'hsncode', 'hsn_code', 'commodity']:
                hsn_col = col
                break
        if not hsn_col:
            for col in df.columns:
                c_low = str(col).strip().lower()
                if any(k in c_low for k in ['hsn', 'sac', 'commodity', 'nomenclature', 'hsn/sac', 'hsn_sac', 'hsn_code']):
                    hsn_col = col
                    break

        for col in df.columns:
            c_low = str(col).strip().lower()
            if any(k in c_low for k in ['transaction type', 'type', 'status', 'order status', 'transaction_type']):
                tx_type_col = col
                break

        for col in df.columns:
            c_low = str(col).strip().lower()
            if any(x in c_low for x in ['gift', 'wrap', 'shipping', 'delivery', 'ship', 'postage', 'tcs', 'amount', 'amt', 'value', 'tax paid', 'tax collected']):
                continue
            if any(r in c_low for r in ['rate', 'percentage', '%', 'code']):
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
                        if cl in ['hsn', 'hsn/sac', 'hsn_code', 'hsncode', 'commoditycode', 'producttaxcode']:
                            attr_hsn_col = c
                            break
                    if not attr_hsn_col:
                        for c in attr_df.columns:
                            cl = str(c).strip().lower()
                            if any(k in cl for k in ['hsn', 'sac', 'taxcode', 'commodity', 'nomenclature']):
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

            # PASS A: COMPUTE ORIGINAL COMBINED TAX RATES
            def extract_rate_number(val):
                if pd.isna(val) or str(val).strip() in ['', 'nan', 'None', '<NA>']: return 0.0
                s = str(val).strip().replace('%', '')
                s = re.sub(r'\.0+$', '', s)
                if s in ['0.018', '0.005', '0.012', '0.028']: s = str(float(s) * 10)
                digits = "".join(c for c in s if c.isdigit() or c == '.')
                try: return float(digits) if digits else 0.0
                except: return 0.0

            cgst_series = df[cgst_col].apply(extract_rate_number) if cgst_col else pd.Series(0.0, index=df.index)
            sgst_series = df[sgst_col].apply(extract_rate_number) if sgst_col else pd.Series(0.0, index=df.index)
            igst_series = df[igst_col].apply(extract_rate_number) if igst_col else pd.Series(0.0, index=df.index)

            if cgst_series.max() > 0 and cgst_series.max() <= 1.0: cgst_series = cgst_series * 100
            if sgst_series.max() > 0 and sgst_series.max() <= 1.0: sgst_series = sgst_series * 100
            if igst_series.max() > 0 and igst_series.max() <= 1.0: igst_series = igst_series * 100

            original_total_rates = []
            for idx in df.index:
                if igst_series[idx] > 0: total_math = igst_series[idx]
                else: total_math = cgst_series[idx] + sgst_series[idx]
                original_total_rates.append(str(int(total_math)) if total_math.is_integer() else str(total_math))

            # PASS B: REPAIR HSN ENCODINGS
            pure_numeric_hsns = []
            for index, row in df.iterrows():
                val = str(row[hsn_col]).strip() if pd.notna(row[hsn_col]) else ""
                sku_raw_val = str(row[sku_col]) if sku_col and pd.notna(row[sku_col]) else ""
                sku_clean_val = deep_clean_sku(sku_raw_val)
                
                if val.startswith('="') and val.endswith('"'): val = val[2:-1]
                val_clean = "".join(filter(str.isdigit, re.sub(r'[\s\-\.\/]', '', val)))
                
                if not val_clean or val_clean in ["", "nan", "none"]:
                    if sku_clean_val in master_sku_hsn_map:
                        val_clean = master_sku_hsn_map[sku_clean_val]
                    else:
                        val_clean = "MISSING HSN"
                
                if val_clean != "MISSING HSN":
                    if len(val_clean) == 7: val_clean = "0" + val_clean
                    pure_numeric_hsns.append(val_clean)
                else:
                    pure_numeric_hsns.append("MISSING HSN")
            df['_temp_hsn_pure'] = pure_numeric_hsns

            # PASS C: CAPTURE DUAL TAX RATE BRAKET GROUPS
            double_rate_hsn_list = {}
            hsn_majority_tax_map = {}
            
            for hsn_val, group in df.groupby('_temp_hsn_pure'):
                if hsn_val != "MISSING HSN":
                    group_indices = group.index
                    rates_in_group = [original_total_rates[df.index.get_loc(idx)] for idx in group_indices]
                    rates_series = pd.Series(rates_in_group)
                    rates_series = rates_series[(rates_series != "") & (rates_series != "0") & (rates_series != "0.0")]
                    
                    if not rates_series.empty:
                        unique_rates = rates_series.unique()
                        if len(unique_rates) > 1:
                            # Save all conflicting rates separated by commas (e.g., "5,18")
                            double_rate_hsn_list[hsn_val] = ",".join(sorted(unique_rates))
                        hsn_majority_tax_map[hsn_val] = rates_series.value_counts().index[0]

            # =========================================================================
            # 🕵️‍♂️ ISOLATE DISCREPANCIES FOR SIDE-BY-SIDE MERGE SUMMARY
            # =========================================================================
            list_missing_hsn = []
            list_double_rates = []
            list_invalid_lengths = []
            list_wrong_tax_hsn = []
            list_wrong_hsn_sku = []
            list_wrong_tax_sku = []

            for index, row in df.iterrows():
                loc_idx = df.index.get_loc(index)
                current_hsn = pure_numeric_hsns[loc_idx]
                current_rate = original_total_rates[loc_idx]
                raw_sku = str(row[sku_col]) if sku_col else ""
                clean_sku = deep_clean_sku(raw_sku)
                tx_status = str(row[tx_type_col]).strip().lower() if tx_type_col else ""

                # Error 1: Missing HSN (No Cancels)
                if current_hsn == "MISSING HSN" and "cancel" not in tx_status:
                    if raw_sku and raw_sku not in list_missing_hsn:
                        list_missing_hsn.append(raw_sku)

                # Error 2: Double Tax Rates
                if current_hsn in double_rate_hsn_list:
                    entry = {"hsn": current_hsn, "sku": raw_sku, "rates": double_rate_hsn_list[current_hsn]}
                    if entry not in list_double_rates:
                        list_double_rates.append(entry)

                # Error 3: Invalid Digit Lengths (Not 6 or 8)
                if current_hsn != "MISSING HSN" and len(current_hsn) not in [6, 8]:
                    if current_hsn not in list_invalid_lengths:
                        list_invalid_lengths.append(current_hsn)

                # Error 4: Wrong Tax Rates (by HSN Mapping)
                if current_hsn in master_hsn_tax_map:
                    master_expected_tax = master_hsn_tax_map[current_hsn]
                    if current_rate != "0" and current_rate != "" and current_rate != master_expected_tax:
                        entry = {"hsn": current_hsn, "rate": current_rate, "correct": master_expected_tax}
                        if entry not in list_wrong_tax_hsn:
                            list_wrong_tax_hsn.append(entry)

                # Error 5: Wrong HSN Code (by SKU mapping)
                if clean_sku in master_sku_hsn_map:
                    master_expected_hsn = master_sku_hsn_map[clean_sku]
                    # Verify original input column vs master
                    orig_hsn_val = "".join(filter(str.isdigit, str(row[hsn_col])))
                    if len(orig_hsn_val) == 7: orig_hsn_val = "0" + orig_hsn_val
                    if orig_hsn_val and orig_hsn_val != master_expected_hsn:
                        entry = {"sku": raw_sku, "wrong_hsn": orig_hsn_val, "correct_hsn": master_expected_hsn}
                        if entry not in list_wrong_hsn_sku:
                            list_wrong_hsn_sku.append(entry)

                # Error 6: Wrong Tax Rate (by SKU mapping)
                if clean_sku in master_sku_tax_map:
                    master_expected_sku_tax = master_sku_tax_map[clean_sku]
                    if current_rate != "0" and current_rate != "" and current_rate != master_expected_sku_tax:
                        entry = {"sku": raw_sku, "rate": current_rate, "correct": master_expected_sku_tax}
                        if entry not in list_wrong_tax_sku:
                            list_wrong_tax_sku.append(entry)

            # =========================================================================
            # 📥 ASSEMBLE UNIFIED SIDE-BY-SIDE SUMMARY SHEETS
            # =========================================================================
            max_len = max(len(list_missing_hsn), len(list_double_rates), len(list_invalid_lengths), 
                          len(list_wrong_tax_hsn), len(list_wrong_hsn_sku), len(list_wrong_tax_sku), 1)

            # Build synchronized multi-column frame structure
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

            # Compile into native binary stream packet 
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                df_audit_report.to_excel(writer, sheet_name='GSTR1_Audit_Error_Dashboard', index=False)
            excel_buffer.seek(0)
            excel_binary_data = excel_buffer.getvalue()

            # =========================================================================
            # 🚀 RUN PRODUCTION SANITIZATION (APPLY CLEAN VOTE OUTCOMES)
            # =========================================================================
            sanitized_hsns = []
            sanitized_rates = []

            for index, row in df.iterrows():
                loc_idx = df.index.get_loc(index)
                current_hsn = pure_numeric_hsns[loc_idx]
                
                if current_hsn in hsn_majority_tax_map:
                    final_rate = hsn_majority_tax_map[current_hsn]
                else:
                    final_rate = original_total_rates[loc_idx]
                
                sanitized_hsns.append(f'="{current_hsn}"' if current_hsn != "MISSING HSN" else "MISSING HSN")
                sanitized_rates.append(final_rate)

            df[hsn_col] = sanitized_hsns
            if cgst_col and sgst_col:
                splits = [str(int(float(r)/2)) if (float(r)/2).is_integer() else str(float(r)/2) for r in sanitized_rates]
                df[cgst_col] = splits
                df[sgst_col] = splits
            if igst_col:
                df[igst_col] = sanitized_rates

            df.drop(columns=['_temp_hsn_pure'], inplace=True, errors='ignore')

            # =========================================================================
            # 🎨 RENDER INTERFACE SUCCESS DASHBOARD
            # =========================================================================
            st.success(f"✨ Data Analytics Scan Complete! Processed {initial_rows} lines successfully.")
            
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
