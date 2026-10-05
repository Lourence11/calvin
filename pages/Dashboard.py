import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Worklist Reader", layout="wide")
set_premium_style()

def main():
    st.title("📋 Worklist Dashboard")
    st.markdown("Read-only dashboard to quickly inspect and review password-protected Worklist Excel files without launching Excel COM.")
    
    # Check for msoffcrypto-tool library
    try:
        import msoffcrypto
        MSOFFCRYPTO_AVAILABLE = True
    except ImportError:
        MSOFFCRYPTO_AVAILABLE = False
        st.error("⚠️ **Dependency Missing**: `msoffcrypto-tool` is required to decrypt password-protected Excel files without COM.")
        st.info("To fix this, please run the following command in your terminal:\n```bash\npip install msoffcrypto-tool\n```")
        st.stop()

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        worklist_file = st.file_uploader("📋 Drag & Drop Password-Protected Worklist File", type=["xlsx", "xls"])
    with col_u2:
        payments_file = st.file_uploader("💰 Drag & Drop Password-Protected Payments File", type=["xlsx", "xls"])

    files_to_process = []
    if worklist_file:
        files_to_process.append(("Worklist File", worklist_file, "📋"))
    if payments_file:
        files_to_process.append(("Payments File", payments_file, "💰"))

    if files_to_process:
        PASSWORD = "MAD_May2026"
        st.info(f"🔑 Automatically decrypting using password: `{PASSWORD}` (High-Speed In-Memory Decryption)")
        
        # Display each file in a major tab if both files are uploaded
        if len(files_to_process) > 1:
            file_tabs = st.tabs([f"{emoji} {label}" for label, _, emoji in files_to_process])
        else:
            file_tabs = [st.container()]
            
        for file_tab, (label, uploaded_file, emoji) in zip(file_tabs, files_to_process):
            with file_tab:
                st.markdown(f"## {emoji} {label}: **{uploaded_file.name}**")
                    
                with st.spinner(f"Reading {uploaded_file.name}..."):
                    try:
                        # Check if the file is encrypted using msoffcrypto-tool
                        uploaded_file.seek(0)
                        is_encrypted = False
                        try:
                            office_file = msoffcrypto.OfficeFile(uploaded_file)
                            is_encrypted = office_file.is_encrypted()
                        except Exception:
                            is_encrypted = False
                        
                        if is_encrypted:
                            decrypted_file = io.BytesIO()
                            office_file.load_key(password=PASSWORD)
                            office_file.decrypt(decrypted_file)
                            decrypted_file.seek(0)
                            xls_source = decrypted_file
                            st.success(f"✅ '{uploaded_file.name}' decrypted successfully using auto-password!")
                        else:
                            uploaded_file.seek(0)
                            xls_source = uploaded_file
                            st.success(f"✅ '{uploaded_file.name}' loaded successfully!")
                        
                        # Use pd.ExcelFile to read sheet names and check for sheet existence
                        xls = pd.ExcelFile(xls_source)
                        sheet_names = xls.sheet_names
                        
                        if not sheet_names:
                            st.error(f"❌ '{uploaded_file.name}' does not contain any sheets.")
                            continue
                        
                        # Visual Metrics for this file
                        col_m1, col_m2 = st.columns(2)
                        col_m1.metric("File Name", uploaded_file.name)
                        col_m2.metric("Total Sheets Detected", len(sheet_names))
                        
                        st.divider()
                        st.subheader("📋 Workbook Sheets & Columns")
                        
                        # Display all sheets using Streamlit Tabs
                        if sheet_names:
                            tabs = st.tabs([f"📄 {sheet}" for sheet in sheet_names])
                            for tab, sheet_name in zip(tabs, sheet_names):
                                with tab:
                                    # Load header row of the specific sheet with pandas (fast)
                                    xls_source.seek(0)
                                    try:
                                        df = pd.read_excel(xls_source, sheet_name=sheet_name, nrows=0)
                                        headers = list(df.columns)
                                        
                                        # Fuzzy match columns
                                        ob_col_name = None
                                        prin_col_name = None
                                        amt_col_name = None
                                        placement_col_name = None
                                        
                                        for h in headers:
                                            h_lower = str(h).lower()
                                            if label == "Worklist File":
                                                # Fuzzy match for Payoff Balance
                                                if "payoff balance" in h_lower or "payoffbal" in h_lower or ("payoff" in h_lower and "bal" in h_lower):
                                                    ob_col_name = h
                                                # Fuzzy match for Principal Balance
                                                elif "principal balance" in h_lower or "principalbal" in h_lower or "prin balance" in h_lower or ("prin" in h_lower and "bal" in h_lower):
                                                    prin_col_name = h
                                            elif label == "Payments File":
                                                # Exact match first for Amount (prevents wrong column pick)
                                                if h_lower == "amount":
                                                    amt_col_name = h  # exact win — stop looking
                                                elif amt_col_name is None and ("amount" in h_lower or h_lower == "amt"):
                                                    amt_col_name = h  # fuzzy fallback only if no exact found yet
                                                # Exact match first for Placement
                                                if h_lower == "placement":
                                                    placement_col_name = h
                                                elif placement_col_name is None and "placement" in h_lower:
                                                    placement_col_name = h
                                        
                                        if label == "Payments File":
                                            # Payments File: filter Amount by Placement campaign
                                            if amt_col_name and placement_col_name:
                                                xls_source.seek(0)
                                                try:
                                                    df_data = pd.read_excel(xls_source, sheet_name=sheet_name, usecols=[amt_col_name, placement_col_name])
                                                    
                                                    # Strictly do not return blank placements
                                                    df_data = df_data.dropna(subset=[placement_col_name])
                                                    df_data = df_data[df_data[placement_col_name].astype(str).str.strip() != ""]
                                                    
                                                    # Force Amount to numeric
                                                    df_data[amt_col_name] = pd.to_numeric(df_data[amt_col_name], errors='coerce')
                                                    df_data = df_data.dropna(subset=[amt_col_name])
                                                    
                                                    placement_series = df_data[placement_col_name].astype(str)
                                                    
                                                    # Use word-boundary regex so "150" never matches "1500", "2150", etc.
                                                    is_150 = placement_series.str.contains(r'\b150\b', case=False, na=False, regex=True)
                                                    is_180 = placement_series.str.contains(r'\b180\b', case=False, na=False, regex=True)
                                                    is_120 = placement_series.str.contains(r'\b120\b', case=False, na=False, regex=True)
                                                    
                                                    # Each group is independent — no exclusion cascading
                                                    df_150 = df_data[is_150 & ~is_180 & ~is_120]   # strictly 150 only
                                                    df_180 = df_data[is_180 & ~is_150 & ~is_120]   # strictly 180 only
                                                    
                                                    has_120_pattern = is_120.any()
                                                    if has_120_pattern:
                                                        df_120 = df_data[is_120 & ~is_150 & ~is_180]   # strictly 120 only
                                                    else:
                                                        # If no 120 pattern, return all rows not in 150 or 180
                                                        df_120 = df_data[~is_150 & ~is_180]
                                                        
                                                    st.markdown(f"### Sheet: **{sheet_name}** ({len(headers)} columns)")
                                                    st.markdown("#### 📦 Amount Breakdown by Placement Campaign")
                                                    
                                                    p_cols = st.columns(3)
                                                    
                                                    # 150 Card (Orange-Gold theme)
                                                    with p_cols[0]:
                                                        sum_150 = df_150[amt_col_name].sum()
                                                        avg_150 = df_150[amt_col_name].mean()
                                                        count_150 = len(df_150)
                                                        st.markdown(
                                                            f"""
                                                            <div style="background: linear-gradient(135deg, rgba(245, 158, 11, 0.1), rgba(217, 119, 6, 0.1)); padding: 15px; border-radius: 10px; border-left: 5px solid #f59e0b; margin-bottom: 20px; min-height: 150px;">
                                                                <h5 style="color: #f59e0b; margin-top: 0; margin-bottom: 5px; font-weight: 700; font-family: 'Inter', sans-serif;">🎯 150 Placement</h5>
                                                                <p style="font-size: 0.8rem; margin-bottom: 8px; color: #b0b0b0;">Accounts: <strong>{count_150:,}</strong></p>
                                                                <p style="font-size: 1.8rem; font-weight: 800; margin-bottom: 4px; color: #ffffff; font-family: 'Inter', sans-serif;">{f"₱{sum_150:,.2f}" if count_150 > 0 else "₱0.00"}</p>
                                                                <span style="font-size: 0.8rem; color: #888;">Average: {f"₱{avg_150:,.2f}" if count_150 > 0 else "₱0.00"}</span>
                                                            </div>
                                                            """,
                                                            unsafe_allow_html=True
                                                        )
                                                        
                                                    # 180 Card (Purple theme)
                                                    with p_cols[1]:
                                                        sum_180 = df_180[amt_col_name].sum()
                                                        avg_180 = df_180[amt_col_name].mean()
                                                        count_180 = len(df_180)
                                                        st.markdown(
                                                            f"""
                                                            <div style="background: linear-gradient(135deg, rgba(139, 92, 246, 0.1), rgba(109, 40, 217, 0.1)); padding: 15px; border-radius: 10px; border-left: 5px solid #8b5cf6; margin-bottom: 20px; min-height: 150px;">
                                                                <h5 style="color: #8b5cf6; margin-top: 0; margin-bottom: 5px; font-weight: 700; font-family: 'Inter', sans-serif;">🎯 180 Placement</h5>
                                                                <p style="font-size: 0.8rem; margin-bottom: 8px; color: #b0b0b0;">Accounts: <strong>{count_180:,}</strong></p>
                                                                <p style="font-size: 1.8rem; font-weight: 800; margin-bottom: 4px; color: #ffffff; font-family: 'Inter', sans-serif;">{f"₱{sum_180:,.2f}" if count_180 > 0 else "₱0.00"}</p>
                                                                <span style="font-size: 0.8rem; color: #888;">Average: {f"₱{avg_180:,.2f}" if count_180 > 0 else "₱0.00"}</span>
                                                            </div>
                                                            """,
                                                            unsafe_allow_html=True
                                                        )
                                                        
                                                    # 120 / Remaining Card (Teal theme)
                                                    with p_cols[2]:
                                                        sum_120 = df_120[amt_col_name].sum()
                                                        avg_120 = df_120[amt_col_name].mean()
                                                        count_120 = len(df_120)
                                                        card_title = "120 Placement" if has_120_pattern else "120 (Remaining) Placement"
                                                        st.markdown(
                                                            f"""
                                                            <div style="background: linear-gradient(135deg, rgba(20, 184, 166, 0.1), rgba(13, 148, 136, 0.1)); padding: 15px; border-radius: 10px; border-left: 5px solid #20b4a6; margin-bottom: 20px; min-height: 150px;">
                                                                <h5 style="color: #20b4a6; margin-top: 0; margin-bottom: 5px; font-weight: 700; font-family: 'Inter', sans-serif;">🎯 {card_title}</h5>
                                                                <p style="font-size: 0.8rem; margin-bottom: 8px; color: #b0b0b0;">Accounts: <strong>{count_120:,}</strong></p>
                                                                <p style="font-size: 1.8rem; font-weight: 800; margin-bottom: 4px; color: #ffffff; font-family: 'Inter', sans-serif;">{f"₱{sum_120:,.2f}" if count_120 > 0 else "₱0.00"}</p>
                                                                <span style="font-size: 0.8rem; color: #888;">Average: {f"₱{avg_120:,.2f}" if count_120 > 0 else "₱0.00"}</span>
                                                            </div>
                                                            """,
                                                            unsafe_allow_html=True
                                                        )

                                                    # Output AMOUNT column only (strictly ordered: 150 -> 180 -> 120/remaining)
                                                    df_150_amt = df_150[[amt_col_name]].copy()
                                                    df_180_amt = df_180[[amt_col_name]].copy()
                                                    df_120_amt = df_120[[amt_col_name]].copy()
                                                    
                                                    df_150_amt.columns = ["AMOUNT"]
                                                    df_180_amt.columns = ["AMOUNT"]
                                                    df_120_amt.columns = ["AMOUNT"]
                                                    
                                                    final_amounts_df = pd.concat([df_150_amt, df_180_amt, df_120_amt], ignore_index=True)
                                                    
                                                    st.markdown("---")
                                                    st.markdown("### 💎 Filtered Payment Amounts (AMOUNT Column Only)")
                                                    st.markdown("Below is the list of payment amounts matching the Placement criteria (150 first, then 180, then 120/remaining):")
                                                    
                                                    total_filtered_sum = final_amounts_df["AMOUNT"].sum()
                                                    st.metric(
                                                        label="Total Filtered Amount",
                                                        value=f"₱{total_filtered_sum:,.2f}",
                                                        delta=f"{len(final_amounts_df):,} payments returned"
                                                    )
                                                    
                                                    st.dataframe(
                                                        final_amounts_df,
                                                        use_container_width=True,
                                                        hide_index=True
                                                    )
                                                    
                                                    # CSV download button
                                                    csv_data = final_amounts_df.to_csv(index=False)
                                                    st.download_button(
                                                        label="📥 Download Filtered Amounts (CSV)",
                                                        data=csv_data,
                                                        file_name=f"filtered_payment_amounts_{sheet_name}.csv",
                                                        mime="text/csv",
                                                        key=f"dl_btn_{sheet_name}"
                                                    )

                                                except Exception as data_err:
                                                    st.error(f"❌ Failed to process Payments data: {data_err}")
                                            else:
                                                missing_cols = []
                                                if not amt_col_name: missing_cols.append("'AMOUNT'")
                                                if not placement_col_name: missing_cols.append("'PLACEMENT'")
                                                st.warning(f"⚠️ Could not perform placement breakdown because column(s) {', '.join(missing_cols)} were not detected in this sheet.")
                                        else:
                                            # Worklist File: Display OB and Prin general cards
                                            # Calculate sums & averages if columns are found
                                            total_ob, avg_ob = None, None
                                            total_prin, avg_prin = None, None
                                            
                                            cols_to_use = []
                                            if ob_col_name:
                                                cols_to_use.append(ob_col_name)
                                            if prin_col_name:
                                                cols_to_use.append(prin_col_name)
                                                
                                            if cols_to_use:
                                                xls_source.seek(0)
                                                try:
                                                    df_data = pd.read_excel(xls_source, sheet_name=sheet_name, usecols=cols_to_use)
                                                    
                                                    if ob_col_name and ob_col_name in df_data.columns:
                                                        ob_series = pd.to_numeric(df_data[ob_col_name], errors='coerce').dropna()
                                                        total_ob = ob_series.sum()
                                                        avg_ob = ob_series.mean()
                                                        
                                                    if prin_col_name and prin_col_name in df_data.columns:
                                                        prin_series = pd.to_numeric(df_data[prin_col_name], errors='coerce').dropna()
                                                        total_prin = prin_series.sum()
                                                        avg_prin = prin_series.mean()
                                                except Exception as data_err:
                                                    st.warning(f"Could not calculate totals for columns: {data_err}")
                                            
                                            st.markdown(f"### Sheet: **{sheet_name}** ({len(headers)} columns)")
                                            
                                            # Render Balance Cards if found
                                            active_matches = []
                                            if ob_col_name:
                                                active_matches.append(("ob", ob_col_name))
                                            if prin_col_name:
                                                active_matches.append(("prin", prin_col_name))
                                                
                                            if active_matches:
                                                st.markdown("#### 💰 Detected Financial Columns")
                                                num_cols = len(active_matches)
                                                bal_cols = st.columns(num_cols)
                                                
                                                for idx, (col_type, col_name) in enumerate(active_matches):
                                                    with bal_cols[idx]:
                                                        if col_type == "ob":
                                                            st.markdown(
                                                                f"""
                                                                <div style="background: linear-gradient(135deg, rgba(255, 75, 43, 0.1), rgba(255, 65, 108, 0.1)); padding: 15px; border-radius: 10px; border-left: 5px solid #ff4b2b; margin-bottom: 20px; min-height: 140px;">
                                                                    <h5 style="color: #ff4b2b; margin-top: 0; margin-bottom: 5px; font-weight: 700; font-family: 'Inter', sans-serif;">💳 Payoff Balance (OB)</h5>
                                                                    <p style="font-size: 0.85rem; margin-bottom: 8px; color: #b0b0b0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">Column: <strong>{ob_col_name}</strong></p>
                                                                    <p style="font-size: 1.8rem; font-weight: 800; margin-bottom: 4px; color: #ffffff; font-family: 'Inter', sans-serif;">{f"₱{total_ob:,.2f}" if total_ob is not None else "N/A"}</p>
                                                                    <span style="font-size: 0.8rem; color: #888;">Average: {f"₱{avg_ob:,.2f}" if avg_ob is not None else "N/A"}</span>
                                                                </div>
                                                                """,
                                                                unsafe_allow_html=True
                                                            )
                                                        elif col_type == "prin":
                                                            st.markdown(
                                                                f"""
                                                                <div style="background: linear-gradient(135deg, rgba(0, 180, 219, 0.1), rgba(0, 131, 176, 0.1)); padding: 15px; border-radius: 10px; border-left: 5px solid #00b4d8; margin-bottom: 20px; min-height: 140px;">
                                                                    <h5 style="color: #00b4d8; margin-top: 0; margin-bottom: 5px; font-weight: 700; font-family: 'Inter', sans-serif;">🏦 Principal Balance (Prin)</h5>
                                                                    <p style="font-size: 0.85rem; margin-bottom: 8px; color: #b0b0b0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">Column: <strong>{prin_col_name}</strong></p>
                                                                    <p style="font-size: 1.8rem; font-weight: 800; margin-bottom: 4px; color: #ffffff; font-family: 'Inter', sans-serif;">{f"₱{total_prin:,.2f}" if total_prin is not None else "N/A"}</p>
                                                                    <span style="font-size: 0.8rem; color: #888;">Average: {f"₱{avg_prin:,.2f}" if avg_prin is not None else "N/A"}</span>
                                                                </div>
                                                                """,
                                                                unsafe_allow_html=True
                                                            )
                                            else:
                                                st.warning("⚠️ No Payoff Balance (OB) or Principal Balance (Prin) columns were detected in this sheet.")
                                    except Exception as sheet_err:
                                        st.error(f"❌ Failed to read sheet '{sheet_name}': {sheet_err}")
                        else:
                            st.warning("⚠️ No sheets available in the decrypted file.")
                            
                    except Exception as e:
                        st.error(f"❌ Failed to decrypt or read file '{uploaded_file.name}': {e}")
                        st.markdown(
                            """
                            **Possible causes:**
                            - The file is not encrypted, or uses a different password than `MAD_May2026`.
                            - The file format is corrupt or invalid.
                            """
                        )
    else:
        st.info("Upload password-protected Excel files to inspect their headers and balances.")

if __name__ == "__main__":
    main()
