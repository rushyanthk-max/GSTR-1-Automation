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

    # STRICT TOTAL TAX COLUMN LOCKOUT (Deliberately skipping any TCS columns)
    for col in df.columns:
        c_low = str(col).strip().lower()
        if 'tcs' in c_low: 
            continue  # Hard skip on TCS tracking columns
            
        if any(k in c_low for k in ['total tax', 'tax percentage', 'tax rate', 'rate%', 'gst rate', 'tax_rate', 'gst%']):
            tax_col = col
            break

    # Fallback to standard commercial rates if no explicit total header exists
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

    # Emergency fallback if all headers are non-standard
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

        # PASS 2: PURE MATHEMATICAL MAJORITY TAX RATE ENFORCER
        tax_corrections_made = 0
        
        if tax_col:
            # Group rows by our clean numeric HSN codes
            for hsn_val, group in df.groupby('_temp_hsn_pure'):
                if hsn_val != "MISSING HSN" and not group[tax_col].empty:
                    
                    # Drop completely blank cells or system NaNs in this specific group safely via string operations
                    valid_taxes = group[tax_col].dropna().astype(str).str.strip()
                    valid_taxes = valid_taxes[(valid_taxes != "") & (valid_taxes.str.lower() != "nan")]
                    
                    if not valid_taxes.empty:
                        # Extract the mathematical mode winner (the value that shows up most often)
                        majority_tax_value = valid_taxes.value_counts().index[0]
                        
                        # Correctly check for variations row-by-row using direct element indexing
                        mismatched_rows = group.index[df.loc[group.index, tax_col].astype(str).str.strip() != majority_tax_value]
                        
                        # Count the corrections and apply the winner to the original tax column
                        tax_corrections_made += len(mismatched_rows)
                        df.loc[mismatched_rows, tax_col] = majority_tax_value

        # PASS 3: FINALIZE EXCEL FORMULA PROTECTION SHIELD FOR HSNS
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
            st.warning(f"⚖️ TAX AUTO-CORRECTION: Overwrote **{tax_corrections_made} anomalous rows** inside **'{tax_col}'** to align with the dominant majority tax rate for their HSN groups!")
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
