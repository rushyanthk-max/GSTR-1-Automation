import streamlit as st
import pandas as pd
import re

# Set up clean browser tab and layout
st.set_page_config(page_title="BCPL Universal GST Sanitizer", layout="centered")

st.title("📦 BCPL Universal E-commerce GST Sanitizer")
st.write("Upload raw sheets from **Amazon, Flipkart, Nykaa, Meesho, JioMart, or WMS** to instantly wipe out data errors.")

# 1. FILE UPLOADER COMPONENT
uploaded_file = st.file_uploader("Upload your raw Excel or CSV report", type=["xlsx", "xls", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(uploaded_file, dtype=str)
    else:
        df = pd.read_csv(uploaded_file, dtype=str, low_memory=False)
        
    initial_rows = len(df)
    
    # 2. REMOVE COMPLETELY BLANK ROWS
    df.dropna(how='all', inplace=True)
    blank_rows = initial_rows - len(df)

    # 3. SMART UNIVERSAL KEYWORD COLUMN SCANNER
    hsn_col = None
    sku_col = None
    cgst_col = None
    sgst_col = None
    igst_col = None
    
    for col in df.columns:
        c_low = str(col).strip().lower()
        
        # Isolate HSN column
        if any(k in c_low for k in ['hsn', 'sac', 'commodity', 'nomenclature', 'hsn/sac', 'hsn_sac', 'hsn_code']):
            hsn_col = col
            
        # Isolate SKU column
        if any(k in c_low for k in ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article', 'wms_code']):
            sku_col = col

        # Target explicit individual tax components (strictly avoiding currency or gift wrap columns)
        if any(x in c_low for x in ['gift', 'wrap', 'shipping', 'delivery', 'ship', 'postage', 'tcs', 'amount', 'amt', 'value', 'tax paid', 'tax collected']):
            continue

        if any(r in c_low for r in ['rate', 'percentage', '%', 'code']):
            if 'cgst' in c_low: cgst_col = col
            if 'sgst' in c_low: sgst_col = col
            if 'igst' in c_low: igst_col = col

    # 4. EXECUTE UNIVERSAL DATA RECONCILIATION
    if hsn_col:
        # Initial safety fill
        df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
        if sku_col:
            df[sku_col] = df[sku_col].fillna("").astype(str).str.strip()

        # STEP 1: SCAN THE SHEET TO BUILD A DYNAMIC SKU -> HSN DIRECTORY
        # This automatically learns your product catalog from rows that aren't broken!
        dynamic_sku_hsn_map = {}
        
        if sku_col:
            for index, row in df.iterrows():
                raw_hsn = str(row[hsn_col]).strip()
                raw_sku = str(row[sku_col]).strip()
                
                # Unbox existing Excel text shields if present during scanning
                if raw_hsn.startswith('="') and raw_hsn.endswith('"'):
                    raw_hsn = raw_hsn[2:-1]
                elif raw_hsn.startswith('='):
                    raw_hsn = raw_hsn.replace('=', '').replace('"', '')
                
                clean_hsn = re.sub(r'[\s\-\.\/]', '', raw_hsn)
                clean_hsn = "".join(filter(str.isdigit, clean_hsn))
                
                # If we find a valid HSN and a valid SKU, save it to our map
                if clean_hsn and clean_hsn not in ["", "nan", "none"] and raw_sku and raw_sku not in ["", "nan", "none"]:
                    # Ensure 7-digit HSNs get padded with a leading zero immediately
                    if len(clean_hsn) == 7:
                        clean_hsn = "0" + clean_hsn
                    dynamic_sku_hsn_map[raw_sku] = clean_hsn

        # STEP 2: AUTO-FILL MISSING HSNS AND NORMALIZE THE ENTIRE COLUMN
        pure_numeric_hsns = []
        auto_filled_count = 0
        
        for index, row in df.iterrows():
            val = str(row[hsn_col]).strip()
            sku_val = str(row[sku_col]).strip() if sku_col else ""
            
            # Clean off any existing equation formulas
            if val.startswith('="') and val.endswith('"'):
                val = val[2:-1]
            elif val.startswith('='):
                val = val.replace('=', '').replace('"', '')
                
            val_clean = re.sub(r'[\s\-\.\/]', '', val)
            val_clean = "".join(filter(str.isdigit, val_clean))
            
            # If the HSN is missing, look it up dynamically via its SKU
            if not val_clean or val_clean in ["", "nan", "none"]:
                if sku_val in dynamic_sku_hsn_map:
                    val_clean = dynamic_sku_hsn_map[sku_val]
                    auto_filled_count += 1
                else:
                    val_clean = "MISSING HSN"
            
            # Final digit checks and zero padding
            if val_clean != "MISSING HSN":
                clean_digits = "".join(filter(str.isdigit, val_clean))
                if len(clean_digits) == 7:
                    clean_digits = "0" + clean_digits
                pure_numeric_hsns.append(clean_digits)
            else:
                pure_numeric_hsns.append("MISSING HSN")
                
        df['_temp_hsn_pure'] = pure_numeric_hsns

        # STEP 3: DYNAMIC "TOTAL TAX RATE" RECONSTRUCTION & PERCENTAGE SCALING
        def extract_rate_number(val):
            if pd.isna(val) or str(val).strip() in ['', 'nan', 'None', '<NA>']:
                return 0.0
            s = str(val).strip().replace('%', '')
            s = re.sub(r'\.0+$', '', s)
            
            if s in ['0.018', '0.005', '0.012', '0.028']:
                s = str(float(s) * 10)  # Correct sub-decimal scaling artifacts
                
            digits = "".join(c for c in s if c.isdigit() or c == '.')
            try:
                return float(digits) if digits else 0.0
            except:
                return 0.0

        cgst_series = df[cgst_col].apply(extract_rate_number) if cgst_col else pd.Series(0.0, index=df.index)
        sgst_series = df[sgst_col].apply(extract_rate_number) if sgst_col else pd.Series(0.0, index=df.index)
        igst_series = df[igst_col].apply(extract_rate_number) if igst_col else pd.Series(0.0, index=df.index)

        if cgst_series.max() > 0 and cgst_series.max() <= 1.0: cgst_series = cgst_series * 100
        if sgst_series.max() > 0 and sgst_series.max() <= 1.0: sgst_series = sgst_series * 100
        if igst_series.max() > 0 and igst_series.max() <= 1.0: igst_series = igst_series * 100

        calculated_total_rates = []
        for idx in df.index:
            if igst_series[idx] > 0:
                total_math = igst_series[idx]
            else:
                total_math = cgst_series[idx] + sgst_series[idx]
            calculated_total_rates.append(str(int(total_math)) if total_math.is_integer() else str(total_math))

        df['Total Tax Rate'] = calculated_total_rates
        tax_col = 'Total Tax Rate'

        if cgst_col: df[cgst_col] = cgst_series.apply(lambda x: str(int(x)) if x.is_integer() else str(x))
        if sgst_col: df[sgst_col] = sgst_series.apply(lambda x: str(int(x)) if x.is_integer() else str(x))
        if igst_col: df[igst_col] = igst_series.apply(lambda x: str(int(x)) if x.is_integer() else str(x))

        # STEP 4: PURE MATHEMATICAL MAJORITY VOTE ENFORCER
        tax_corrections_made = 0
        for hsn_val, group in df.groupby('_temp_hsn_pure'):
            if hsn_val != "MISSING HSN" and not group[tax_col].empty:
                valid_taxes = group[tax_col].dropna().astype(str).str.strip()
                valid_taxes = valid_taxes[(valid_taxes != "") & (valid_taxes != "0") & (valid_taxes != "0.0")]
                
                if not valid_taxes.empty:
                    majority_tax_value = valid_taxes.value_counts().index[0]
                    mismatched_rows = group.index[df.loc[group.index, tax_col].astype(str).str.strip() != majority_tax_value]
                    
                    if len(mismatched_rows) > 0:
                        tax_corrections_made += len(mismatched_rows)
                        df.loc[mismatched_rows, tax_col] = majority_tax_value

                        try:
                            new_total_num = float(majority_tax_value)
                            if igst_col: df.loc[mismatched_rows, igst_col] = majority_tax_value
                            if cgst_col and sgst_col:
                                split_str = str(int(new_total_num / 2)) if (new_total_num / 2).is_integer() else str(new_total_num / 2)
                                df.loc[mismatched_rows, cgst_col] = split_str
                                df.loc[mismatched_rows, sgst_col] = split_str
                        except:
                            pass

        # STEP 5: FINALIZE EXCEL FORMULA PROTECTION SHIELD FOR HSNS
        final_shielded_hsns = []
        for val in df['_temp_hsn_pure']:
            if val == "MISSING HSN":
                final_shielded_hsns.append("MISSING HSN")
            else:
                final_shielded_hsns.append(f'="{val}"')
                
        df[hsn_col] = final_shielded_hsns
        df.drop(columns=['_temp_hsn_pure'], inplace=True)

        # 5. RENDER INTERFACE SUCCESS DASHBOARD
        st.success(f"✨ File parsed successfully! Cleaned up {blank_rows} blank formatting rows.")
        st.info(f"🧬 **Dynamic Reconstruction Engine Connected:** \n* **Discovered Rate Columns:** CGST ('{cgst_col if cgst_col else 'Not Found'}'), SGST ('{sgst_col if sgst_col else 'Not Found'}'), IGST ('{igst_col if igst_col else 'Not Found'}') \n* **Automated Self-Healing:** Successfully cross-referenced and repaired **{auto_filled_count} missing HSN rows** by matching their product SKUs dynamically!")
        
        if tax_corrections_made > 0:
            st.warning(f"⚖️ TAX AUTO-CORRECTION COMPLETE: Detected double tax rates! Overwrote **{tax_corrections_made} rows** inside your engineered **'Total Tax Rate'** column to perfectly align with majority group rules!")
        else:
            st.success("✅ Tax Rate Integrity: Checked all dynamic groups. Every matching item row perfectly aligns.")
            
    else:
        st.error("❌ Column Detection Error: The script could not automatically identify an HSN column name in your file.")

    st.write("### Data Preview Grid:")
    st.dataframe(df.head(50))
    
    csv_data = df.to_csv(index=False).encode('utf-8')
    
    # 6. DOWNLOAD COMPONENT BUTTON
    st.download_button(
        label="📥 Download Sanitized File for Repotic",
        data=csv_data,
        file_name=f"CLEANED_{uploaded_file.name.split('.')[0]}.csv",
        mime="text/csv"
    )
