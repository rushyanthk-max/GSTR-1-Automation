import streamlit as st
import pandas as pd
import re

# Set up clean browser tab and layout
st.set_page_config(page_title="BCPL Universal GST Sanitizer", layout="centered")

st.title("📦 BCPL Universal E-commerce GST Sanitizer")
st.write("Upload raw sheets from **Amazon, Flipkart, Nykaa, Meesho, JioMart, or WMS** to instantly wipe out data errors.")

# =========================================================================
# BCPL MASTER PRODUCT CATALOG (Add your SKUs and correct HSNs here!)
# =========================================================================
sku_hsn_catalog = {
    "SKU_SAMPLE_1": "33049910",  
    "MUG-BLUE-01": "69111011",   
}
# =========================================================================

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

    # 3. UNIVERSAL KEYWORD COLUMN SCANNER
    hsn_col = "Hsn/sac"
    sku_col = "Sku"
    cgst_col = "Cgst_Rate"
    sgst_col = "Sgst_Rate"
    igst_col = "Igst_Rate"
    
    for col in df.columns:
        c_low = str(col).strip().lower()
        
        # Isolate HSN column
        if any(k in c_low for k in ['hsn', 'sac', 'commodity', 'nomenclature', 'hsn/sac', 'hsn_sac', 'hsn_code']):
            hsn_col = col
            
        # Isolate SKU column
        if any(k in c_low for k in ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article', 'wms_code']):
            sku_col = col

        # Target explicit individual tax components (and reject absolute values/amounts)
        if 'amount' not in c_low and 'amt' not in c_low and 'value' not in c_low and 'tcs' not in c_low:
            if 'cgst' in c_low: cgst_col = ccol
            if 'sgst' in c_low: sgst_col = scol
            if 'igst' in c_low: igst_col = icol

        print (ccol)
        print (scol)
        print (icol)

    # 4. EXECUTE UNIVERSAL DATA RECONCILIATION
    if hsn_col:
        # Clear base HSN strings safely
        df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
        if sku_col:
            df[sku_col] = df[sku_col].fillna("").astype(str).str.strip()

        # PASS 1: Extract temporary pure numeric keys for calculation grouping
        pure_numeric_hsns = []
        for index, row in df.iterrows():
            val = str(row[hsn_col]).strip()
            sku_val = str(row[sku_col]).strip() if sku_col else ""
            
            if val.startswith('="') and val.endswith('"'):
                val = val[2:-1]
            elif val.startswith('='):
                val = val.replace('=', '').replace('"', '')
                
            val_clean = re.sub(r'[\s\-\.\/]', '', val)
            
            if not val_clean or val_clean in ["", "nan", "None"]:
                if sku_val in sku_hsn_catalog:
                    val_clean = sku_hsn_catalog[sku_val]
                else:
                    val_clean = "MISSING HSN"
            
            clean_digits = "".join(filter(str.isdigit, val_clean))
            if len(clean_digits) == 7:
                clean_digits = "0" + clean_digits
                
            if not clean_digits or val_clean == "MISSING HSN":
                pure_numeric_hsns.append("MISSING HSN")
            else:
                pure_numeric_hsns.append(clean_digits)
                
        df['_temp_hsn_pure'] = pure_numeric_hsns

        # PASS 2: DYNAMIC "TOTAL TAX RATE" RECONSTRUCTION MATHEMATICS
        # Mathematical string parser to extract digits cleanly
        def extract_rate_number(val):
            if pd.isna(val) or str(val).strip() in ['', 'nan', 'None', '<NA>']:
                return 0.0
            s = str(val).strip().replace('%', '')
            s = re.sub(r'\.0+$', '', s)
            digits = "".join(c for c in s if c.isdigit() or c == '.')
            try:
                return float(digits) if digits else 0.0
            except:
                return 0.0

        # Run extraction math on individual sub-component rows
        cgst_series = df[cgst_col].apply(extract_rate_number) if cgst_col else pd.Series(0.0, index=df.index)
        sgst_series = df[sgst_col].apply(extract_rate_number) if sgst_col else pd.Series(0.0, index=df.index)
        igst_series = df[igst_col].apply(extract_rate_number) if igst_col else pd.Series(0.0, index=df.index)

        # Build our mathematically verified Total Combined Tax column string!
        calculated_total_rates = []
        for idx in df.index:
            # If IGST has a value, that is our total tax rate. Otherwise, it is the sum of CGST + SGST
            if igst_series[idx] > 0:
                total_math = igst_series[idx]
            else:
                total_math = cgst_series[idx] + sgst_series[idx]
            
            # Reformat to clean string representation (e.g. 18 instead of 18.0)
            calculated_total_rates.append(str(int(total_math)) if total_math.is_integer() else str(total_math))

        # Inject our newly minted target column into the main dataframe
        df['Total Tax Rate'] = calculated_total_rates
        tax_col = 'Total Tax Rate'

        # PASS 3: PURE MATHEMATICAL MAJORITY VOTE ENFORCER
        tax_corrections_made = 0
        
        # Group rows by our clean numeric HSN codes
        for hsn_val, group in df.groupby('_temp_hsn_pure'):
            if hsn_val != "MISSING HSN" and not group[tax_col].empty:
                
                # Isolate rows containing values
                valid_taxes = group[tax_col].dropna().astype(str).str.strip()
                valid_taxes = valid_taxes[(valid_taxes != "") & (valid_taxes != "0")]
                
                if not valid_taxes.empty:
                    # Extract the absolute mathematical winner
                    majority_tax_value = valid_taxes.value_counts().index[0]
                    
                    # Target any rows in this group that contradict the majority vote
                    mismatched_rows = group.index[df.loc[group.index, tax_col].astype(str).str.strip() != majority_tax_value]
                    
                    # Count corrections and enforce the clean winner onto the calculated row column
                    tax_corrections_made += len(mismatched_rows)
                    df.loc[mismatched_rows, tax_col] = majority_tax_value

                    # SYNC BACK MATH: If split sub-columns exist, balance them out to prevent calculation mismatch errors downstream
                    if tax_corrections_made > 0:
                        try:
                            new_total_num = float(majority_tax_value)
                            if igst_col:
                                df.loc[mismatched_rows, igst_col] = majority_tax_value
                            if cgst_col and sgst_col:
                                split_value = str(new_total_num / 2)
                                split_str = str(int(new_total_num / 2)) if (new_total_num / 2).is_integer() else split_value
                                df.loc[mismatched_rows, cgst_col] = split_str
                                df.loc[mismatched_rows, sgst_col] = split_str
                        except:
                            pass

        # PASS 4: FINALIZE EXCEL FORMULA PROTECTION SHIELD FOR HSNS
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
        st.info(f"🧬 **Dynamic Reconstruction Engine Connected:** \n* **Discovered Split Columns:** CGST ('{cgst_col}'), SGST ('{sgst_col}'), IGST ('{igst_col}') \n* **Generated Verified Target:** 'Total Tax Rate' column successfully calculated from raw row rows!")
        
        if tax_corrections_made > 0:
            st.warning(f"⚖️ TAX AUTO-CORRECTION COMPLETE: Detected double tax rates! Overwrote **{tax_corrections_made} rows** inside your newly engineered **'Total Tax Rate'** column (and corresponding sub-components) to perfectly align with majority group rules!")
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
