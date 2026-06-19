import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="GSTR-1 Automation Suite", layout="wide")

st.title("🛍️ GSTR-1 Data Preparation & Compliance Suite")
st.caption("High-performance tax automation engine built for large marketplace datasets.")

# Create the Navigation Tabs
tab1, tab2 = st.tabs(["Phase 1: Marketplace Rectifier", "Phase 2: Error Engine"])

# =====================================================================
# PHASE 1: RECTIFIER MODULE
# =====================================================================
with tab1:
    st.header("Marketplace Report Rectifier")
    st.write("Upload your raw sales report to clean HSN codes, compute total tax rates, and eliminate double-tax rate anomalies.")
    
    p1_file = st.file_uploader("Upload Raw Marketplace Report (XLSX or CSV)", type=["xlsx", "csv"], key="p1_upload")
    
    if p1_file is not None:
        try:
            with st.spinner("Processing massive dataset..."):
                # Read file safely
                if p1_file.name.endswith('.csv'):
                    df = pd.read_csv(p1_file, dtype=str)
                else:
                    df = pd.read_excel(p1_file, dtype=str)
                
                df.columns = df.columns.str.strip()
                
                # Dynamic Case-Insensitive Header Matching
                hsn_col = next((c for c in df.columns if c.lower() in ['hsn code', 'hsn/sac', 'hsn']), None)
                igst_col = next((c for c in df.columns if c.lower() in ['igst', 'igst rate', 'igstamt']), None)
                let_cgst_col = next((c for c in df.columns if c.lower() in ['cgst', 'cgst rate', 'cgstamt']), None)
                let_sgst_col = next((c for c in df.columns if c.lower() in ['sgst', 'sgst rate', 'sgstamt']), None)
                
                if not hsn_col:
                    st.error("❌ Could not find an HSN column in this report. Please check your file headers.")
                else:
                    # Parse and convert found tax columns safely
                    igst_vals = pd.to_numeric(df[igst_col], errors='coerce').fillna(0.0) if igst_col else 0.0
                    cgst_vals = pd.to_numeric(df[let_cgst_col], errors='coerce').fillna(0.0) if let_cgst_col else 0.0
                    sgst_vals = pd.to_numeric(df[let_sgst_col], errors='coerce').fillna(0.0) if let_sgst_col else 0.0
                    
                    # Calculate and normalize Total Tax Rate
                    total_tax = igst_vals + cgst_vals + sgst_vals
                    df['Total Tax Rate'] = total_tax.apply(lambda x: round(x * 100) if 0 < x < 1 else round(x))
                    
                    # Clean and pad HSN codes
                    def clean_hsn_func(val):
                        if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "missing hsn":
                            return "Missing HSN"
                        cleaned = str(val).split('.')[0].strip()
                        return '0' + cleaned if len(cleaned) == 7 else cleaned
                        
                    df[hsn_col] = df[hsn_col].apply(clean_hsn_func)
                    
                    # Strict Majority Tax Calculation (Removes double tax completely)
                    valid_hsn_df = df[df[hsn_col] != "Missing HSN"]
                    if not valid_hsn_df.empty:
                        majority_tax = valid_hsn_df.groupby([hsn_col, 'Total Tax Rate']).size().reset_index(name='count')
                        majority_tax = majority_tax.sort_values(by=[hsn_col, 'count', 'Total Tax Rate'], ascending=[True, False, False])
                        majority_tax_map = majority_tax.drop_duplicates(subset=[hsn_col]).set_index(hsn_col)['Total Tax Rate'].to_dict()
                        
                        # Pre-convert columns to float to allow row assignments safely
                        if igst_col: df[igst_col] = pd.to_numeric(df[igst_col], errors='coerce').fillna(0.0)
                        if let_cgst_col: df[let_cgst_col] = pd.to_numeric(df[let_cgst_col], errors='coerce').fillna(0.0)
                        if let_sgst_col: df[let_sgst_col] = pd.to_numeric(df[let_sgst_col], errors='coerce').fillna(0.0)

                        def fix_row(row):
                            h = row[hsn_col]
                            curr = row['Total Tax Rate']
                            dom = majority_tax_map.get(h, curr)
                            if dom != curr:
                                if igst_col and row[igst_col] > 0:
                                    row[igst_col] = float(dom)
                                else:
                                    if let_cgst_col: row[let_cgst_col] = float(dom / 2)
                                    if let_sgst_col: row[let_sgst_col] = float(dom / 2)
                                row['Total Tax Rate'] = int(dom)
                            return row
                            
                        df = df.apply(fix_row, axis=1)
                    
                    # Success Summary Metrics
                    st.success(f"✨ Successfully parsed {len(df):,} rows! Double tax rates have been rectified.")
                    
                    # Convert to Excel layout stream and keep leading zeros intact
                    output_p1 = io.BytesIO()
                    with pd.ExcelWriter(output_p1, engine='xlsxwriter') as writer:
                        df.to_excel(writer, sheet_name='Rectified Data', index=False)
                        workbook = writer.book
                        worksheet = writer.sheets['Rectified Data']
                        hsn_letter = chr(65 + df.columns.get_loc(hsn_col))
                        worksheet.set_column(f'{hsn_letter}:{hsn_letter}', None, workbook.add_format({'num_format': '@'}))
                    
                    st.download_button(
                        label="📥 Download Rectified Report",
                        data=output_p1.getvalue(),
                        file_name=f"Rectified_Phase1_{p1_file.name.split('.')[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Error processing file: {e}")

# =====================================================================
# PHASE 2: ERROR ENGINE MODULE
# =====================================================================
with tab2:
    st.header("Multi-Sheet Error Diagnostic Engine")
    st.write("Upload BOTH your Marketplace Report and your Product Master to catch and map all six systematic data errors.")
    
    col1, col2 = st.columns(2)
    with col1:
        p2_market_file = st.file_uploader("1. Upload Marketplace Report", type=["xlsx", "csv"], key="p2_m")
    with col2:
        p2_master_file = st.file_uploader("2. Upload Product Master File", type=["xlsx", "csv"], key="p2_p")
        
    if p2_market_file and p2_master_file:
        try:
            with st.spinner("Executing relational database cross-checks..."):
                # Load marketplace data
                df_m = pd.read_csv(p2_market_file, dtype=str) if p2_market_file.name.endswith('.csv') else pd.read_excel(p2_market_file, dtype=str)
                # Load product master data
                df_p = pd.read_csv(p2_master_file, dtype=str) if p2_master_file.name.endswith('.csv') else pd.read_excel(p2_master_file, dtype=str)
                
                df_m.columns = df_m.columns.str.strip()
                df_p.columns = df_p.columns.str.strip()
                
                # Column Index Lookups (Case Insensitive)
                m_hsn = next((c for c in df_m.columns if c.lower() in ['hsn code', 'hsn/sac', 'hsn']), "HSN Code")
                m_sku = next((c for c in df_m.columns if c.lower() in ['sku', 'seller sku', 'product sku']), "SKU")
                m_type = next((c for c in df_m.columns if c.lower() in ['transaction type', 'type', 'order status']), "Transaction Type")
                m_igst = next((c for c in df_m.columns if c.lower() in ['igst', 'igst rate']), None)
                m_cgst = next((c for c in df_m.columns if c.lower() in ['cgst', 'cgst rate']), None)
                m_sgst = next((c for c in df_m.columns if c.lower() in ['sgst', 'sgst rate']), None)
                
                p_sku = next((c for c in df_p.columns if c.lower() in ['sku', 'product sku', 'item code']), "SKU")
                p_hsn = next((c for c in df_p.columns if c.lower() in ['correct hsn', 'hsn', 'hsn code']), "Correct HSN")
                p_tax = next((c for c in df_p.columns if c.lower() in ['correct tax rate', 'tax rate', 'tax', 'gst rate']), "Correct Tax Rate")
                
                # Real-time sum total tax processing safely
                m_igst_vals = pd.to_numeric(df_m[m_igst], errors='coerce').fillna(0.0) if m_igst else 0.0
                m_cgst_vals = pd.to_numeric(df_m[m_cgst], errors='coerce').fillna(0.0) if m_cgst else 0.0
                m_sgst_vals = pd.to_numeric(df_m[m_sgst], errors='coerce').fillna(0.0) if m_sgst else 0.0
                
                total_tax_m = m_igst_vals + m_cgst_vals + m_sgst_vals
                df_m['Total Tax Rate'] = total_tax_m.apply(lambda x: round(x * 100) if 0 < x < 1 else round(x))
                
                # Format master values cleanly
                df_p[p_sku] = df_p[p_sku].str.strip().str.lower()
                df_p[p_hsn] = df_p[p_hsn].str.split('.').str[0].str.strip()
                df_p[p_tax] = pd.to_numeric(df_p[p_tax], errors='coerce').fillna(0.0)
                df_p[p_tax] = df_p[p_tax].apply(lambda x: round(x * 100) if 0 < x < 1 else round(x))
                
                master_sku_map = df_p.set_index(p_sku)[[p_hsn, p_tax]].to_dict(orient='index')
                master_hsn_map = df_p.drop_duplicates(subset=[p_hsn]).set_index(p_hsn)[p_tax].to_dict()
                
                # Error arrays
                err1, err2, err3, err4, err5, err6 = [], [], [], [], [], []
                hsn_tax_sets, hsn_sku_sets = {}, {}
                
                for _, r in df_m.iterrows():
                    h = ''.join(filter(str.isdigit, str(r.get(m_hsn, "")).split('.')[0].strip())) if pd.notna(r.get(m_hsn)) else ""
                    s = str(r.get(m_sku, "")).strip()
                    t = int(r['Total Tax Rate'])
                    if h and h != "" and h.lower() != "missing hsn":
                        hsn_tax_sets.setdefault(h, set()).add(t)
                        hsn_sku_sets.setdefault(h, set()).add(s)
                
                for idx, r in df_m.iterrows():
                    hsn = ''.join(filter(str.isdigit, str(r.get(m_hsn, "")).split('.')[0].strip())) if pd.notna(r.get(m_hsn)) else ""
                    if hsn.lower() == "missing hsn" or hsn == "nan": hsn = ""
                    
                    sku = str(r.get(m_sku, "")).strip()
                    sku_lower = sku.lower()
                    tx_type = str(r.get(m_type, "")).lower() if pd.notna(r.get(m_type)) else ""
                    tax = int(r['Total Tax Rate'])
                    row_num = idx + 2
                    
                    is_cancelled = "cancel" in tx_type or "return" in tx_type
                    
                    if not hsn and not is_cancelled:
                        err1.append({"Row Index": row_num, "SKU": sku, "Transaction Type": r.get(m_type, "Unknown"), "Tax Rate": tax})
                    
                    if hsn:
                        if len(hsn) != 6 and len(hsn) != 8:
                            err6.append({"Row Index": row_num, "SKU": sku, "Invalid HSN Code": hsn, "Length": len(hsn), "Tax Rate": tax})
                        if hsn in master_hsn_map and int(master_hsn_map[hsn]) != tax:
                            err3.append({"SKU": sku, "HSN Code": hsn, "Marketplace Tax": tax, "Master Expected Tax": int(master_hsn_map[hsn])})
                            
                    if sku_lower in master_sku_map:
                        truth = master_sku_map[sku_lower]
                        if hsn and truth[p_hsn] != hsn:
                            err4.append({"SKU": sku, "Marketplace HSN": hsn, "Master Correct HSN": truth[p_hsn], "Tax Rate": tax})
                        if tax > 0 and int(truth[p_tax]) != tax:
                            err5.append({"SKU": sku, "HSN": hsn, "Marketplace Tax": tax, "Master Correct Tax Rate": int(truth[p_tax])})
                
                for hsn_code, tax_set in hsn_tax_sets.items():
                    if len(tax_set) > 1:
                        err2.append({"HSN Code": hsn_code, "Tax Rates Detected": ", ".join(map(str, tax_set)), "SKUs Implicated": ", ".join(hsn_sku_sets[hsn_code])})
                
                # Visual Metric Dashboard Summary
                m1, m2, m3, m4, m5, m6 = st.columns(6)
                m1.metric("S1: Blank HSN", len(err1))
                m2.metric("S2: Double Tax", len(err2))
                m3.metric("S3: Tax Mismatch", len(err3))
                m4.metric("S4: SKU vs HSN", len(err4))
                m5.metric("S5: SKU vs Tax", len(err5))
                m6.metric("S6: Bad Digits", len(err6))
                
                # Package up 6 Sheets workbook output stream
                output_p2 = io.BytesIO()
                with pd.ExcelWriter(output_p2, engine='xlsxwriter') as writer:
                    def pack(lst, cols, name):
                        pd.DataFrame(lst if lst else [], columns=cols).to_excel(writer, sheet_name=name, index=False)
                    pack(err1, ["Row Index", "SKU", "Transaction Type", "Tax Rate"], "Blank HSNs")
                    pack(err2, ["HSN Code", "Tax Rates Detected", "SKUs Implicated"], "Double GST Rates")
                    pack(err3, ["SKU", "HSN Code", "Marketplace Tax", "Master Expected Tax"], "GST Rate Mismatches")
                    pack(err4, ["SKU", "Marketplace HSN", "Master Correct HSN", "Tax Rate"], "HSN SKU Mismatches")
                    pack(err5, ["SKU", "HSN", "Marketplace Tax", "Master Correct Tax Rate"], "Tax SKU Mismatches")
                    pack(err6, ["Row Index", "SKU", "Invalid HSN Code", "Length", "Tax Rate"], "Invalid HSN Lengths")
                
                st.download_button(
                    label="📥 Download 6-Sheet Complete Error Report",
                    data=output_p2.getvalue(),
                    file_name="GSTR1_Comprehensive_Error_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error tracking metrics pipeline: {e}")
