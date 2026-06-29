import streamlit as st
import pandas as pd
import re

# Set up clean browser tab and layout
st.set_page_config(page_title="BCPL Universal GST Sanitizer", layout="centered")

st.title("📦 BCPL Universal E-commerce GST Sanitizer")
st.write("Upload your sales report **AND** your master product attribute sheet to instantly wipe out data errors.")

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

# AGGRESSIVE DEEP CLEAN LAYER
def deep_clean_sku(sku_val):
    if pd.isna(sku_val):
        return ""
    s = str(sku_val).strip().lower()
    # Strip away backticks, spaces, dashes, hyphens, and punctuation entirely
    s = re.sub(r'[^a-z0-9]', '', s)
    return s

# =========================================================================
# 1. DUAL FILE UPLOADER COMPONENTS
# =========================================================================
st.subheader("1️⃣ Step 1: Upload Raw Transaction Report")
uploaded_file = st.file_uploader("Drop your sales report here (Amazon, Flipkart, Nykaa, Meesho, JioMart, WMS)", type=["xlsx", "xls", "csv"], key="sales_report")

st.subheader("2️⃣ Step 2: Upload Master Product Attribute / Catalog File (Optional)")
attribute_file = st.file_uploader("Drop your Master Item Catalog sheet here to heal remaining missing HSN fields", type=["xlsx", "xls", "csv"], key="attribute_sheet")

if uploaded_file:
    df = load_data_safely(uploaded_file)
    
    if df is not None:
        initial_rows = len(df)
        df.dropna(how='all', inplace=True)
        blank_rows = initial_rows - len(df)

        # 2. SMART UNIVERSAL KEYWORD COLUMN SCANNER (Main Sheet - Prioritized)
        hsn_col = None
        sku_col = None
        cgst_col = None
        sgst_col = None
        igst_col = None
        
        # Pass 1: Look for exact structural SKU matching names
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

        # Pass 1: Look for exact structural HSN matching names
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

        # Target explicit split tax percentage rate markers
        for col in df.columns:
            c_low = str(col).strip().lower()
            if any(x in c_low for x in ['gift', 'wrap', 'shipping', 'delivery', 'ship', 'postage', 'tcs', 'amount', 'amt', 'value', 'tax paid', 'tax collected']):
                continue
            if any(r in c_low for r in ['rate', 'percentage', '%', 'code']):
                if 'cgst' in c_low: cgst_col = col
                if 'sgst' in c_low: sgst_col = col
                if 'igst' in c_low: igst_col = col

        # 3. EXECUTE MASTER DATA RECONCILIATION
        if hsn_col:
            df[hsn_col] = df[hsn_col].fillna("").astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
            if sku_col:
                df[sku_col] = df[sku_col].fillna("").astype(str).str.strip()

            # BASE DICTIONARY BUILDING
            master_sku_hsn_map = {}

            # PART A: Extract HSN mapping data from the external Product Attribute file
            if attribute_file:
                attr_df = load_data_safely(attribute_file)
                if attr_df is not None:
                    attr_hsn_col = None
                    attr_sku_col = None
                    
                    # 🎯 CRITICAL RECTIFICATION: Strict exact match locking for catalog headers
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
                            
                    if attr_hsn_col and attr_sku_col:
                        for _, row in attr_df.iterrows():
                            r_hsn = str(row[attr_hsn_col]).strip() if pd.notna(row[attr_hsn_col]) else ""
                            r_sku_clean = deep_clean_sku(row[attr_sku_col])
                            
                            if r_hsn.startswith('="') and r_hsn.endswith('"'): r_hsn = r_hsn[2:-1]
                            r_clean = re.sub(r'[\s\-\.\/]', '', r_hsn)
                            r_digits = "".join(filter(str.isdigit, r_clean))
                            
                            if r_digits and r_digits.lower() not in ["", "nan", "none"] and r_sku_clean:
                                if len(r_digits) == 7: r_digits = "0" + r_digits
                                master_sku_hsn_map[r_sku_clean] = r_digits
                    st.sidebar.success(f"📖 Loaded {len(master_sku_hsn_map)} reference links from Attribute sheet!")

            # PART B: Supplement mapping database using internal rows that aren't broken
            if sku_col:
                for _, row in df.iterrows():
                    raw_hsn = str(row[hsn_col]).strip() if pd.notna(row[hsn_col]) else ""
                    raw_sku_clean = deep_clean_sku(row[sku_col])
                    
                    if raw_hsn.startswith('="') and raw_hsn.endswith('"'): raw_hsn = raw_hsn[2:-1]
                    elif raw_hsn.startswith('='): raw_hsn = raw_hsn.replace('=', '').replace('"', '')
                    
                    clean_hsn = re.sub(r'[\s\-\.\/]', '', raw_hsn)
                    clean_hsn = "".join(filter(str.isdigit, clean_hsn))
                    
                    if clean_hsn and clean_hsn not in ["", "nan", "none"] and raw_sku_clean:
                        if len(clean_hsn) == 7: clean_hsn = "0" + clean_hsn
                        if raw_sku_clean not in master_sku_hsn_map:
                            master_sku_hsn_map[raw_sku_clean] = clean_hsn

            # PART C: AUTO-FILL MISSING HSNS WITH DEEP FUZZY MATCH TRACKING
            pure_numeric_hsns = []
            auto_filled_count = 0
            missing_unresolved = 0
            
            for index, row in df.iterrows():
                val = str(row[hsn_col]).strip() if pd.notna(row[hsn_col]) else ""
                sku_raw_val = str(row[sku_col]) if sku_col and pd.notna(row[sku_col]) else ""
                sku_clean_val = deep_clean_sku(sku_raw_val)
                
                if val.startswith('="') and val.endswith('"'): val = val[2:-1]
                elif val.startswith('='): val = val.replace('=', '').replace('"', '')
                    
                val_clean = re.sub(r'[\s\-\.\/]', '', val)
                val_clean = "".join(filter(str.isdigit, val_clean))
                
                if not val_clean or val_clean in ["", "nan", "none"]:
                    # Check Pass 1: Direct structural clean lookup
                    if sku_clean_val in master_sku_hsn_map:
                        val_clean = master_sku_hsn_map[sku_clean_val]
                        auto_filled_count += 1
                    # Check Pass 2: Fuzzy slice lookup (ignores trailing platform letters like _A or X)
                    elif len(sku_clean_val) > 1 and sku_clean_val[:-1] in master_sku_hsn_map:
                        val_clean = master_sku_hsn_map[sku_clean_val[:-1]]
                        auto_filled_count += 1
                    else:
                        val_clean = "MISSING HSN"
                
                if val_clean != "MISSING HSN":
                    clean_digits = "".join(filter(str.isdigit, val_clean))
                    if len(clean_digits) == 7: clean_digits = "0" + clean_digits
                    pure_numeric_hsns.append(clean_digits)
                else:
                    missing_unresolved += 1
                    pure_numeric_hsns.append("MISSING HSN")
                    
            df['_temp_hsn_pure'] = pure_numeric_hsns

            # PART D: DYNAMIC "TOTAL TAX RATE" RECONSTRUCTION & PERCENTAGE SCALING
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

            calculated_total_rates = []
            for idx in df.index:
                if igst_series[idx] > 0: total_math = igst_series[idx]
                else: total_math = cgst_series[idx] + sgst_series[idx]
                calculated_total_rates.append(str(int(total_math)) if total_math.is_integer() else str(total_math))

            df['Total Tax Rate'] = calculated_total_rates
            tax_col = 'Total Tax Rate'

            if cgst_col: df[cgst_col] = cgst_series.apply(lambda x: str(int(x)) if x.is_integer() else str(x))
            if sgst_col: df[sgst_col] = sgst_series.apply(lambda x: str(int(x)) if x.is_integer() else str(x))
            if igst_col: df[igst_col] = igst_series.apply(lambda x: str(int(x)) if x.is_integer() else str(x))

            # PART E: PURE MATHEMATICAL MAJORITY VOTE ENFORCER
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

            # PART F: FINALIZE EXCEL FORMULA PROTECTION SHIELD FOR HSNS
            final_shielded_hsns = []
            for val in df['_temp_hsn_pure']:
                if val == "MISSING HSN": final_shielded_hsns.append("MISSING HSN")
                else: final_shielded_hsns.append(f'="{val}"')
                    
            df[hsn_col] = final_shielded_hsns
            df.drop(columns=['_temp_hsn_pure'], inplace=True)

            # 4. RENDER INTERFACE SUCCESS DASHBOARD
            st.success(f"✨ File parsed successfully! Cleaned up {blank_rows} blank formatting rows.")
            st.info(f"🧬 **Universal Alphanumeric Deep Cleaning Connected:** \n* **Fuzzy Repair Status:** Successfully matched and healed **{auto_filled_count} missing HSN rows** by bypassing all spaces, dashes, punctuation marks, and structural encoding mismatches!")
            
            if tax_corrections_made > 0:
                st.warning(f"⚖️ TAX AUTO-CORRECTION COMPLETE: Overwrote **{tax_corrections_made} rows** inside your engineered **'Total Tax Rate'** column to resolve double tax rates!")
            else:
                st.success("✅ Tax Rate Integrity: Checked all dynamic groups. Every matching item row perfectly aligns.")

            if missing_unresolved > 0:
                st.warning(f"⚠️ Notice: {missing_unresolved} rows still could not be resolved. These remain labeled as 'MISSING HSN'.")

        else:
            st.error("❌ Column Detection Error: The script could not automatically identify an HSN column name in your file.")

        st.write("### Data Preview Grid:")
        st.dataframe(df.head(50))
        
        csv_data = df.to_csv(index=False).encode('utf-8')
        
        st.download_button(
            label="📥 Download Sanitized File for Repotic",
            data=csv_data,
            file_name=f"CLEANED_{uploaded_file.name.split('.')[0]}.csv",
            mime="text/csv"
        )
