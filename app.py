import streamlit as st
import pandas as pd
import re
import io

# Set up clean browser tab and layout
st.set_page_config(page_title="BCPL Universal GST Sanitizer & Auditor", layout="centered")

st.title("📦 BCPL Universal E-commerce GST Sanitizer")
st.write("Upload your reports to sanitize errors and instantly generate a **6-Sheet Compliance Audit Report**.")

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
        # Create a working copy for processing and keep an unmutated raw copy for error logging
        df = df_raw.copy()
        initial_rows = len(df)
        df.dropna(how='all', inplace=True)
        blank_rows = initial_rows - len(df)

        # 2. SMART UNIVERSAL KEYWORD COLUMN SCANNER (Main Sheet)
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

        # Isolate Transaction Type column for filtering cancelled orders
        for col in df.columns:
            c_low = str(col).strip().lower()
            if any(k in c_low for k in ['transaction type', 'type', 'status', 'order status', 'transaction_type']):
                tx_type_col = col
                break

        # Target explicit split tax percentage rate markers
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
            # Clean standard layout formats
            df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
            if sku_col:
                df[sku_col] = df[sku_col].fillna("").astype(str).str.strip()

            # MASTER CATALOG MAPS (Built from the uploaded attribute file)
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
                            
                            # Clean HSN formatting digits
                            if r_hsn.startswith('="') and r_hsn.endswith('"'): r_hsn = r_hsn[2:-1]
                            r_clean = re.sub(r'[\s\-\.\/]', '', r_hsn)
                            r_digits = "".join(filter(str.isdigit, r_clean))
                            if len(r_digits) == 7: r_digits = "0" + r_digits

                            # Extract numeric whole percentages out of text fields (e.g. "Tax_18%" -> "18")
                            r_tax_digits = "".join(filter(str.isdigit, r_tax_raw))
                            
                            if r_digits and r_digits.lower() not in ["", "nan", "none"] and r_sku_clean:
                                master_sku_hsn_map[r_sku_clean] = r_digits
                                if r_tax_digits:
                                    master_sku_tax_map[r_sku_clean] = r_tax_digits
                                    master_hsn_tax_map[r_digits] = r_tax_digits
            
            # Supplement reference map from clean entries already inside the raw sheet
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

            # PASS A: COMPUTE TARGET ORIGINAL COMBINED TAX RATES FROM SALES REPORT
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
            
            # Map original un-sanitized calculated rates directly to the original backup tracking frame
            df_raw['Calculated Total Tax Rate'] = original_total_rates

            # PASS B: REPAIR MISALIGNED HSNS AND DETECT VOTE GROUPS
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

            # PASS C: COMPUTE THE MATHEMATICAL VOTE MAP TO SPOT DOUBLE RATES
            double_rate_hsn_list = []
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
                            double_rate_hsn_list.append(hsn_val)
                        hsn_majority_tax_map[hsn_val] = rates_series.value_counts().index[0]

            # =========================================================================
            # 🕵️‍♂️ COMPILING THE 6 COMPLIANCE SHEET LOGS
            # =========================================================================
            error_indices_1 = [] # Missing HSNs (No Cancels)
            error_indices_2 = [] # Double Tax Rates
            error_indices_3 = [] # Invalid Digits (Not 6 or 8)
            error_indices_4 = [] # Wrong Tax Rates by HSN Mapping
            error_indices_5 = [] # Wrong HSN Codes mapped by SKU catalog
            error_indices_6 = [] # Wrong Tax Rates mapped by SKU catalog

            for index, row in df.iterrows():
                loc_idx = df.index.get_loc(index)
                current_hsn = pure_numeric_hsns[loc_idx]
                current_rate = original_total_rates[loc_idx]
                raw_sku = str(row[sku_col]) if sku_col else ""
                clean_sku = deep_clean_sku(raw_sku)
                tx_status = str(row[tx_type_col]).strip().lower() if tx_type_col else ""

                # Error 1: Missing HSN and transaction is not cancelled
                if current_hsn == "MISSING HSN" and "cancel" not in tx_status:
                    error_indices_1.append(index)

                # Error 2: HSN belongs to a dual tax bracket conflict group
                if current_hsn in double_rate_hsn_list:
                    error_indices_2.append(index)

                # Error 3: HSN length is non-compliant (neither 6 nor 8 digits)
                if current_hsn != "MISSING HSN" and len(current_hsn) not in [6, 8]:
                    error_indices_3.append(index)

                # Error 4: Cross-reference HSN to master catalog tax validation
                if current_hsn in master_hsn_tax_map:
                    master_expected_tax = master_hsn_tax_map[current_hsn]
                    if current_rate != "0" and current_rate != "" and current_rate != master_expected_tax:
                        error_indices_4.append(index)

                # Error 5: Check if the row's input HSN conflicts with master SKU registry
                if clean_sku in master_sku_hsn_map:
                    master_expected_hsn = master_sku_hsn_map[clean_sku]
                    if current_hsn != "MISSING HSN" and current_hsn != master_expected_hsn:
                        error_indices_5.append(index)

                # Error 6: Check if the row's tax rate percentage conflicts with master SKU registry
                if clean_sku in master_sku_tax_map:
                    master_expected_sku_tax = master_sku_tax_map[clean_sku]
                    if current_rate != "0" and current_rate != "" and current_rate != master_expected_sku_tax:
                        error_indices_6.append(index)

            # Slice logs neatly out of the pristine unmutated raw backup frame
            err_sheet_1 = df_raw.loc[error_indices_1]
            err_sheet_2 = df_raw.loc[error_indices_2]
            err_sheet_3 = df_raw.loc[error_indices_3]
            err_sheet_4 = df_raw.loc[error_indices_4]
            err_sheet_5 = df_raw.loc[error_indices_5]
            err_sheet_6 = df_raw.loc[error_indices_6]

            # =========================================================================
            # 📥 GENERATE MULTI-SHEET EXCEL PACKET IN STREAM MEMORY
            # =========================================================================
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                err_sheet_1.to_excel(writer, sheet_name='1_Missing_HSNs_No_Cancels', index=False)
                err_sheet_2.to_excel(writer, sheet_name='2_Double_Tax_Rates', index=False)
                err_sheet_3.to_excel(writer, sheet_name='3_Invalid_Lengths_Not_6_or_8', index=False)
                err_sheet_4.to_excel(writer, sheet_name='4_Wrong_Tax_Rates_by_HSN', index=False)
                err_sheet_5.to_excel(writer, sheet_name='5_Wrong_HSNs_by_SKU', index=False)
                err_sheet_6.to_excel(writer, sheet_name='6_Wrong_Tax_Rates_by_SKU', index=False)
            excel_buffer.seek(0)

            # =========================================================================
            # 🚀 RUN LIVE PRODUCTION SANITIZATION (APPLY VOTE WINNERS)
            # =========================================================================
            sanitized_hsns = []
            sanitized_rates = []

            for index, row in df.iterrows():
                loc_idx = df.index.get_loc(index)
                current_hsn = pure_numeric_hsns[loc_idx]
                
                # Apply vote map rule or keep baseline calculations
                if current_hsn in hsn_majority_tax_map:
                    final_rate = hsn_majority_tax_map[current_hsn]
                else:
                    final_rate = original_total_rates[loc_idx]
                
                sanitized_hsns.append(f'="{current_hsn}"' if current_hsn != "MISSING HSN" else "MISSING HSN")
                sanitized_rates.append(final_rate)

            # Apply final clean alignment onto the output file parameters
            df[hsn_col] = sanitized_hsns
            if cgst_col and sgst_col:
                splits = [str(int(float(r)/2)) if (float(r)/2).is_integer() else str(float(r)/2) for r in sanitized_rates]
                df[cgst_col] = splits
                df[sgst_col] = splits
            if igst_col:
                df[igst_col] = sanitized_rates

            # Strip working calculation markers safely before display
            df.drop(columns=['_temp_hsn_pure'], inplace=True, errors='ignore')

            # =========================================================================
            # 🎨 RENDER STREAMLIT INTERFACE HUD
            # =========================================================================
            st.success(f"✨ Data Analytics Scan Complete! Cleaned {blank_rows} trailing layout gaps.")
            
            st.info(f"📋 **Detected System Mappings:** \n* **Core Targets:** HSN (`{hsn_col}`) | SKU (`{sku_col}`) \n* **Tax Splits:** CGST (`{cgst_col}`), SGST (`{sgst_col}`), IGST (`{igst_col}`)")
            
            # Download Box 1: The Multi-Sheet Audit Log
            st.error("⚠️ COMPLIANCE RISK DETECTED: Audit logs generated below!")
            st.download_button(
                label="📥 Download 6-Sheet Audit Error Report",
                data=excel_buffer,
                file_name=f"AUDIT_LOG_{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Download Box 2: The Sanitized Final Sales Packet
            st.success("✅ SANITIZATION SUMMARY: Double rates resolved and text-shields applied successfully.")
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
