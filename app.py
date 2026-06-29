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

    # 3. SMART UNIVERSAL KEYWORD COLUMN SCANNER
    hsn_col = None
    sku_col = None
    tax_col = None
    
    for col in df.columns:
        c_low = str(col).strip().lower()
        if any(k in c_low for k in ['hsn', 'sac', 'commodity', 'nomenclature', 'hsn/sac', 'hsn_sac']):
            hsn_col = col
        if any(k in c_low for k in ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article', 'wms_code']):
            sku_col = col
        if any(k in c_low for k in ['product tax code', 'total tax', 'tax percentage', 'tax rate', 'rate%', 'igst rate', 'gst rate', 'tax_rate', 'gst%']):
            tax_col = col

    # Emergency Fallbacks
    if not hsn_col:
        hsn_candidates = [c for c in df.columns if 'hsn' in str(c).lower()]
        if hsn_candidates: hsn_col = hsn_candidates[0]
        
    if not tax_col:
        tax_candidates = [c for c in df.columns if 'rate' in str(c).lower() or 'tax' in str(c).lower()]
        if tax_candidates: tax_col = tax_candidates[0]

    # 4. EXECUTE UNIVERSAL DATA RECONCILIATION
    if hsn_col:
        # Deep clean columns safely
        df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
        if tax_col:
            df[tax_col] = df[tax_col].fillna("0").astype(str).str.strip()
        
        # PASS 1: Generate a TEMPORARY pure-numeric HSN column for perfect grouping calculations
        pure_numeric_hsns = []
        for index, row in df.iterrows():
            val = str(row[hsn_col]).strip()
            sku_val = str(row[sku_col]).strip() if sku_col else ""
            
            # CRITICAL: Strip off any pre-existing Excel equation formulas completely!
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

        # PASS 2: TAX RATE HARMONIZATION (Groups using the guaranteed pure numeric list)
        if tax_col:
            # Clean up the tax strings inside the data loop to prevent "18" vs "18.0" vs "18%" issues
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

            # Group by pure numeric HSN strings and replace tax data with the dominant true majority value
            for hsn_val, group in df.groupby('_temp_hsn_pure'):
                if hsn_val != "MISSING HSN" and not group['_temp_tax_match'].empty:
                    majority_tax_value = group['_temp_tax_match'].value_counts().index[0]
                    
                    # Pull a sample style template of how the raw tax code looked for the majority
                    sample_rows = group[group['_temp_tax_match'] == majority_tax_value]
                    raw_style_template = str(df.loc[sample_rows.index[0], tax_col]).strip()
                    
                    # Force all rows in this group to match the raw format of the majority rate
                    df.loc[df['_temp_hsn_pure'] == hsn_val, tax_col] = raw_style_template

            df.drop(columns=['_temp_tax_match'], inplace=True)

        # PASS 3: WRAP THE PADDED HSNS IN EXCEL TEXT SHIELDS FOR THE FINAL OUTPUT
        final_shielded_hsns = []
        for val in df['_temp_hsn_pure']:
            if val == "MISSING HSN":
                final_shielded_hsns.append("MISSING HSN")
            else:
                final_shielded_hsns.append(f'="{val}"')
                
        df[hsn_col] = final_shielded_hsns
        df.drop(columns=['_temp_hsn_pure'], inplace=True) # Trash temp tracking column

        st.success(f"✨ File parsed successfully! Cleaned up {blank_rows} blank rows.")
        st.info(f"🎯 Target Matrix Connected -> HSN Column: '{hsn_col}' | Adjusted Tax Column: '{tax_col if tax_col else 'Not Found'}'")
            
    else:
        st.error("❌ Column Detection Error: The script could not automatically identify an HSN or Commodity column in this file layout.")

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
