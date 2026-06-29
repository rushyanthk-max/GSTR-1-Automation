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
    # Force Pandas to read every cell as raw text string to protect zeroes
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
        st.info(f"⚡ Universal Auto-Detection Matrix Loaded! Processing files...")
        
        # Deep clean columns safely
        df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
        
        # Process HSN codes cleanly without nested functions
        cleaned_hsns = []
        for index, row in df.iterrows():
            val = str(row[hsn_col]).strip()
            sku_val = str(row[sku_col]).strip() if sku_col else ""
            
            if val.startswith('="') and val.endswith('"'):
                val = val[2:-1]
                
            val_clean = re.sub(r'[\s\-\.\/]', '', val)
            
            if not val_clean or val_clean in ["", "nan", "None"]:
                if sku_val in sku_hsn_catalog:
                    val_clean = sku_hsn_catalog[sku_val]
                else:
                    val_clean = "MISSING HSN"
            
            clean_digits = "".join(filter(str.isdigit, val_clean))
            if len(clean_digits) == 7:
                clean_digits = "0" + clean_digits
                
            if clean_digits and clean_digits != "MISSING HSN":
                cleaned_hsns.append(f'="{clean_digits}"')
            else:
                cleaned_hsns.append("MISSING HSN")
                
        df[hsn_col] = cleaned_hsns

        # 5. DYNAMIC TAX RATE HARMONIZATION ENGINE (Flat majority logic execution)
        if tax_col:
            df[tax_col] = df[tax_col].fillna("0").astype(str).str.strip()
            # Group by HSN and replace tax data directly with the dominant group value
            for hsn_val, group in df.groupby(hsn_col):
                if hsn_val != "MISSING HSN" and not group[tax_col].empty:
                    majority_tax = group[tax_col].value_counts().index[0]
                    df.loc[df[hsn_col] == hsn_val, tax_col] = majority_tax

        st.success(f"✨ File parsed successfully! Cleaned up {blank_rows} blank rows.")
            
    else:
        st.error("❌ Column Detection Error: The script could not identify an HSN column name in your file.")

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
