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

    # 3. SMART UNIVERSAL KEYWORD COLUMN SCANNER (With strict TCS exclusion)
    hsn_col = None
    sku_col = None
    tax_col = None
    
    for col in df.columns:
        c_low = str(col).strip().lower()
        
        # Isolate HSN column
        if any(k in c_low for k in ['hsn', 'sac', 'commodity', 'nomenclature', 'hsn/sac', 'hsn_sac', 'hsn_code']):
            hsn_col = col
            
        # Isolate SKU column
        if any(k in c_low for k in ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article', 'wms_code']):
            sku_col = col

    # STRICT TOTAL TAX COLUMN LOCKOUT (Deliberately trashing any TCS markers)
    # Pass 1: Look for exact absolute total tax parameters
    for col in df.columns:
        c_low = str(col).strip().lower()
        if 'tcs' in c_low: 
            continue  # 🛑 Hard skip on any column tracking TCS rates
            
        if any(k in c_low for k in ['total tax', 'tax percentage', 'tax rate', 'rate%', 'gst rate', 'tax_rate', 'gst%']):
            tax_col = col
            break

    # Pass 2: Fallback to standard commercial rates (IGST/CGST/SGST total tracking markers) if no total header is present
    if not tax_col:
        for col in df.columns:
            c_low = str(col).strip().lower()
            if 'tcs' in c_low: 
                continue
            if 'igst' in c_low and 'rate' in c_low:
                tax_col = col
                break
            elif 'tax' in c_low and 'code' in c_low:
                tax_col = col
                break

    # Emergency fallback if headings are completely stripped
    if not tax_col:
        tax_candidates = [c for c in df.columns if 'rate' in str(c).lower() or 'tax' in str(c).lower() and 'tcs' not in str(c).lower()]
        if tax_candidates: tax_col = tax_candidates[0]

    # 4. EXECUTE UNIVERSAL DATA RECONCILIATION
    if hsn_col:
        # Clear base strings safely
        df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
        if tax_col:
            df[tax_col] = df[tax_col].fillna("0").astype(str).str.strip()
        
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

        # PASS 2: COMMERCIAL TAX RATE HARMONIZATION (Majority Wins Engine)
        tax_corrections_made = 0
        hsn_majority_tax_map = {}

        if tax_col:
            def standardize_tax_values(val):
                t = str(val).strip().upper().replace('%', '')
                t = re.sub(r'\.0+$', '', t)
                t = re.sub(r'[\s\-\.\/]', '', t)
                if '18' in t or 'STANDARD' in t: return "18"
                if '5' in t or 'REDUCED' in t or 'LOW' in t: return "5"
                if '12' in t: return "12"
                if '28' in t: return "28"
                return t

            df['_temp_tax_match'] = df[tax_col].apply(standardize_tax_values)

            # Map true commercial majority groups
            for hsn_val, group in df.groupby('_temp_hsn_pure'):
                if hsn_val != "MISSING HSN" and not group['_temp_tax_match'].empty:
                    valid_group = group[group['_temp_tax_match'] != "UNKNOWN"]
                    if not valid_group.empty:
                        majority_tax_value = valid_group['_temp_tax_match'].value_counts().index[0]
                        hsn_majority_tax_map[hsn_val] = majority_tax_value

            # Overwrite minority rate bugs
            def apply_harmonization(row):
                global tax_corrections_made
                hsn_pure = row['_temp_hsn_pure']
                current_raw = str(row[tax_col]).strip()
                current_clean = row['_temp_tax_match']
                
                if hsn_pure in hsn_majority_tax_map:
                    correct_majority = hsn_majority_tax_map[hsn_pure]
                    if current_clean != correct_majority and current_clean != "UNKNOWN":
                        tax_corrections_made += 1
                        
                        # Maintain original style layouts (.0 or %)
                        if '.' in current_raw: return correct_majority + ".0"
                        if '%' in current_raw: return correct_majority + "%"
                        return correct_majority
                return current_raw

            df[tax_col] = df.apply(apply_harmonization, axis=1)
            df.drop(columns=['_temp_tax_match'], inplace=True)

        # PASS 3: FINALIZE EXCEL FORMULA PROTECTION SHIELD FOR HSNs
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
        st.info(f"🎯 **Target Matrix Connected** \n* **HSN Column:** '{hsn_col}' \n* **Total Tax Column:** '{tax_col if tax_col else 'Not Found'}'")
        
        if tax_col and tax_corrections_made > 0:
            st.warning(f"⚖️ TAX AUTO-CORRECTION: Overwrote **{tax_corrections_made} anomalous rows** inside **'{tax_col}'** to resolve double tax rates via dominant majority rule!")
        elif tax_col:
            st.success(f"✅ Tax Rate Integrity: Checked all rows under total tax column '{tax_col}'. All matching item groups align.")
            
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
