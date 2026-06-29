import streamlit as st
import pandas as pd
import re

# Set up a clean browser tab and layout
st.set_page_config(page_title="BCPL Universal GST Sanitizer", layout="centered")

st.title("📦 BCPL Universal E-commerce GST Sanitizer")
st.write("Upload raw sheets from **Amazon, Flipkart, Nykaa, Meesho, JioMart, or WMS** to instantly wipe out data errors.")

# =========================================================================
# 🛑 BCPL MASTER PRODUCT CATALOG (Add your SKUs and correct HSNs here!)
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
        
        # Isolate HSN column
        if any(k in c_low for k in ['hsn', 'sac', 'commodity', 'nomenclature', 'hsn/sac', 'hsn_sac']):
            hsn_col = col
            
        # Isolate SKU column
        if any(k in c_low for k in ['sku', 'fsn', 'seller-sku', 'item-code', 'product-id', 'article', 'wms_code']):
            sku_col = col
            
        # Isolate Tax column (Looking for the final total rate, not splits)
        if any(k in c_low for k in ['product tax code', 'total tax', 'tax percentage', 'tax rate', 'rate%', 'igst rate', 'gst rate', 'tax_rate', 'gst%']):
            if any(t in c_low for t in ['total', 'percentage', 'code', 'rate', '%']):
                tax_col = col

    # Emergency Fallbacks if the smart search missed a weirdly named column
    if not hsn_col:
        hsn_candidates = [c for c in df.columns if 'hsn' in str(c).lower()]
        if hsn_candidates: hsn_col = hsn_candidates[0]
        
    if not tax_col:
        tax_candidates = [c for c in df.columns if 'rate' in str(c).lower() or 'tax' in str(c).lower()]
        if tax_candidates: tax_col = tax_candidates[0]
        
    if not sku_col:
        sku_candidates = [c for c in df.columns if any(k in str(c).lower() for k in ['sku', 'code', 'id'])]
        if sku_candidates: sku_col = sku_candidates[0]

    # 4. EXECUTE UNIVERSAL DATA RECONCILIATION
    if hsn_col:
        st.info(f"⚡ **Universal Auto-Detection Matrix:** \n* **HSN Column Located:** '{hsn_col}' \n* **Tax Column Located:** '{tax_col if tax_col else 'Not Found/Not Needed'}' \n* **SKU Column Located:** '{sku_col if sku_col else 'Not Found'}'")
        
        # Initial Deep Clean: Remove standard nulls and Excel decimal artifacts (.0)
        df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
        df[hsn_col] = df[hsn_col].replace(['nan', 'None', '<na>', '<NA>'], "")
        
        if sku_col:
            df[sku_col] = df[sku_col].fillna("").astype(str).str.strip()

        # PASS 1: Aggressive HSN Cleanup (Strips spaces, hyphens, and slashes completely)
        def initial_hsn_cleanup(row):
            val = str(row[hsn_col]).strip()
            sku_val = str(row[sku_col]).strip() if sku_col else ""

            if val.startswith('="') and val.endswith('"'):
                val = val[2:-1]

            # Strip out spaces, hyphens, dots, slashes entirely
            val_clean = re.sub(r'[\s\-\.\/]', '', val)

            if not val_clean or val_clean in ["", "nan", "None"]:
                if sku_val in sku_hsn_catalog:
                    return sku_hsn_catalog[sku_val]
                else:
                    return "MISSING HSN"
            
            clean_digits = "".join(filter(str.isdigit, val_clean))
            if len(clean_digits) == 7:
                return "0" + clean_digits
            return clean_digits

        df['_temp_hsn'] = df.apply(initial_hsn_cleanup, axis=1)

        # PASS 2: DYNAMIC TAX RATE HARMONIZATION ENGINE
        tax_corrections_made = 0
        hsn_majority_tax_map = {}

        if tax_col:
            # Cleans up numeric tax rates AND Amazon tax string codes simultaneously
            def standard_tax_extractor(val):
                if pd.isna(val) or str(val).strip() in ['nan', 'None', '', '<NA>']:
                    return "UNKNOWN"
                
                # Strip spaces and cast to upper case
                s = str(val).strip().upper().replace('%', '')
                s = re.sub(r'\.0+$', '', s)
                s = re.sub(r'[\s\-\.\/]', '', s) # Strip symbols and spaces from tax inputs
                
                if '18' in s or 'STANDARD' in s: return "18"
                if '5' in s or 'REDUCED' in s or 'LOW' in s: return "5"
                if '12' in s: return "12"
                if '28' in s: return "28"
                if '0' in s or 'EXEMPT' in s: return "0"
                
                s_digits = "".join(filter(str.isdigit, s))
                return s_digits if s_digits else s

            df['_temp_tax_clean'] = df[tax_col].apply(standard_tax_extractor)

            # Group by clean HSN and find the dominant majority tax value
            for hsn_code, group in df.groupby('_temp_hsn'):
                if hsn_code != "MISSING HSN" and not group['_temp_tax_clean'].empty:
                    # Filter out UNKNOWN values for tax evaluation if possible
                    valid_group = group[group['_temp_tax_clean'] != "UNKNOWN"]
                    if not valid_group.empty:
                        majority_tax = valid_group['_temp_tax_clean'].value_counts().index[0]
                        hsn_majority_tax_map[hsn_code] = majority_tax

            # Apply majority rule back to the tax column rows dynamically
            def harmonize_taxes(row):
                global tax_corrections_made
                hsn = row['_temp_hsn']
                current_raw_tax = str(row[tax_col]).strip()
                current_clean_tax = row['_temp_tax_clean']
                
                if hsn in hsn_majority_tax_map:
                    correct_majority_value = hsn_majority_tax_map[hsn]
                    
                    if current_clean_tax != correct_majority_value and correct_majority_value != "UNKNOWN":
                        tax_corrections_made += 1
                        
                        # Match original style formatting layout
                        sample_match = df[(df['_temp_hsn'] == hsn) & (df['_temp_tax_clean'] == correct_majority_value)]
                        if not sample_match.empty:
                            return str(sample_match[tax_col].iloc[0]).strip()
                        
                        if '.' in current_raw_tax: return correct_majority_value + ".0"
                        if '%' in current_raw_tax: return correct_majority_value + "%"
                        return correct_majority_value
                        
                return current_raw_tax

            df[tax_col] = df.apply(harmonize_taxes, axis=1)
            df.drop(columns=['_temp_tax_clean'], inplace=True)

        # PASS 3: FINALIZE EXCEL TEXT PROTECTION SHIELD FOR HSNs
        padded_count = 0
        filled_count = 0
        missing_count = 0

        def wrap_hsn_shield(val):
            global padded_count, filled_count, missing_count
            if val == "MISSING HSN":
                missing_count += 1
                return "MISSING HSN"
            if val.startswith('0'):
                padded_count += 1
            return f'="{val}"'

        df[hsn_col] = df['_temp_hsn'].apply(wrap_hsn_shield)
        df.drop(columns=['_temp_hsn'], inplace=True)

        # 5. RENDER SUCCESS DASHBOARD
        st.success(f"✨ File parsed successfully! Cleaned up {blank_rows} blank formatting rows.")
        st.info(f"🔢 HSN PADDING UPDATE: Processed and protected **{padded_count} HSN codes** with Excel text shields.")
        
        if tax_col and tax_corrections_made > 0:
            st.warning(f"⚖️ TAX AUTO-CORRECTION: Automatically aligned **{tax_corrections_made} rows** inside the column **'{tax_col}'** to resolve conflicting double tax rates using majority rule!")
        elif tax_col:
            st.success(f"✅ Tax Rate Integrity: Checked all rows under targeted tax column '{tax_col}'. No conflicting double tax rates found.")

        if missing_count > 0:
            st.warning(f"⚠️ Warning: Found {missing_count} fields that remain blank. Marked as 'MISSING HSN'.")
            
    else:
        st.error("❌ Column Detection Error: The script could not automatically identify an HSN or Commodity column in this file layout. Please verify the sheet headers.")

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
