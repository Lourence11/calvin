import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO

# --- CONFIGURATION ---
st.set_page_config(page_title="Collection Efforts Automation (Combined)", layout="wide")

# Apply consistent premium design
set_premium_style()


PAGE_MAP = {
    "Home": "Main.py",
    "Daily Prod": "pages/1_Daily_Prod.py",
    "Collection Efforts Clean": "pages/2_CE_Clean.py",
    "Collection Efforts Combined": "pages/3_CE_COMBINED.py",
    "Daily Prods accu": "pages/dailyprod.py",
}

# --- STRICT LIST DEFINITIONS FOR CALL ---
POS_LIST_CALL = [

    "POS 3RD PARTY", "POS CLIENT", "PTP EPA", "RPC_INBOUND CALL", "PTP OLD", "PTP NEW NEGO", "TRANSFER", 

    "POS 3RD PARTY - CLIENT NOT AROUND", "POS 3RD PARTY - CLIENT OUT OF THE COUNTRY",
    "POS 3RD PARTY - DROPPED THE CALL", "POS 3RD PARTY - LEAVE MESSAGE FOR CALLBACK",
    "POS 3RD PARTY - NO LONGER CONNECTED", "POS 3RD PARTY - UNCOOPERATIVE",
    "POS CLIENT - CALL DROPPED (WITH PID)", "POS CLIENT - FOR MANUAL CALL BACK",
    "POS CLIENT - RFD 3RD PARTY UNDER MEDICATION", "POS CLIENT - RFD BUSINESS BANKRUPTCY",
    "POS CLIENT - RFD BUSINESS SLOWDOWN", "POS CLIENT - RFD BUSINESS SLOWDOWN/BANKRUPTCY",
    "POS CLIENT - RFD CLAIMING FULLY PAID", "POS CLIENT - RFD CLIENT IS UNDER MEDICATION",
    "POS CLIENT - RFD CLIENT WAS SCAMMED", "POS CLIENT - RFD FUNDS WAS DELAYED",
    "POS CLIENT - RFD UNEMPLOYED", "POS CLIENT - RFD VICTIM OF NATURAL CALAMITY",
    "POS CLIENT - RFD WITH OTHER LOANS TO PAY", "PTP EPA - COMPLYING",
    "PTP EPA - DEFAULTED", "PTP EPA - PRE TERMINATION", "PTP EPA - SPLIT DOWNPAYMENT",
    "PTP NEW NEGO - EPA NO DOWNPAYMENT", "PTP NEW NEGO - EPA WITH DOWNPAYMENT",
    "PTP NEW NEGO - ONE TIME (FULL)", "PTP NEW NEGO - ONE TIME (SPLIT)",
    "PTP NEW NEGO - PERENNIAL", "PTP OLD - RENEGO EPA WITH DOWNPAYMENT", "POS CLIENT - RFD CLIENT IS BREADWINNER",
    "PTP OLD - RENEGO OTP", "PTP OLD - RENEGO OTP SPLIT", "PTP OLD - RENEGO EPA WITHOUT DOWNPAYMENT",
    "RPC_INBOUND CALL (IC) - INQUIRY", "RPC_INBOUND CALL (IC) - NEGO",
    "RPC_INBOUND CALL (IC) - RETURN CALL", "TRANSFER",
    "PTP_FF UP - KEEPS ON RINGING (KOR)",
    "PTP_FF UP - ANSWERED WILL SETTLE",
    "PTP_FF UP - BUSY",

]

NEG_LIST_CALL = [
    "BUSY", "DROPPED"," NEG VIA CALL", "RNA", "NEG VIA CALL - BUSY", "NEG VIA CALL - DROPPED THE CALL",
    "NEG VIA CALL - KEEPS ON RINGING", "NEG VIA CALL - NO LONGER CONNECTED (NLC)",
    "NEG VIA CALL - NO SUCH PERSON (NSP)", "NEG VIA CALL - NOT IN SERVICE (NIS)", 
]
# --- STRICT LIST DEFINITIONS FOR SKIP ---
POS_LIST_SKIP = [
    "POS VIA SOCMED"
]

NEG_LIST_SKIP = [
    "NEG VIA SOCMED"
]

def main():
    # Ensure the save directory exists only when this page is active
    SAVE_PATH = r"C:\Users\SPM\Documents\SaveFolder\COLLECTION EFFORTS\Combined"
    os.makedirs(SAVE_PATH, exist_ok=True)

    st.title("Collection Efforts Automation (Combined CALL & SKIP)")
    st.success("Remember to add '000000' in account number and date format is short date, delete DECEASED")

    uploaded_files = st.file_uploader("Drop your files here", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

    if uploaded_files:
        all_dfs = []
        
        for file in uploaded_files:
            try:
                if file.name.endswith('.csv'):
                    temp_df = pd.read_csv(file)
                else:
                    temp_df = pd.read_excel(file)
                all_dfs.append(temp_df)
            except Exception as e:
                st.error(f"Error loading {file.name}: {e}")

        if all_dfs:
            df = pd.concat(all_dfs, ignore_index=True)
            st.success(f"Successfully merged {len(uploaded_files)} files!")

            # 2. Setup Columns
            date_col = 'Date' 
            status_col = 'Status' 

            if date_col in df.columns and status_col in df.columns:
                # Parse dates
                df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
                invalid_dates = df[date_col].isna().sum()
                if invalid_dates > 0:
                    st.warning(f"Skipped {invalid_dates} rows due to invalid date formats.")
                    df = df.dropna(subset=[date_col])

                # Calculate date ranges
                today = datetime.today()
                curr_m, curr_y = today.month, today.year
                last_month_dt = today - relativedelta(months=1)
                last_m, last_y = last_month_dt.month, last_month_dt.year

                # --- PROCESS CALL LOGIC ---
                pos_curr_call = df[(df[status_col].str.startswith(tuple(POS_LIST_CALL), na=False)) & (df[date_col].dt.month == curr_m) & (df[date_col].dt.year == curr_y)].sort_values(by=date_col, ascending=False)
                neg_curr_call = df[(df[status_col].str.startswith(tuple(NEG_LIST_CALL), na=False)) & (df[date_col].dt.month == curr_m) & (df[date_col].dt.year == curr_y)].sort_values(by=date_col, ascending=False)
                pos_last_call = df[(df[status_col].str.startswith(tuple(POS_LIST_CALL), na=False)) & (df[date_col].dt.month == last_m) & (df[date_col].dt.year == last_y)].sort_values(by=date_col, ascending=False)
                neg_last_call = df[(df[status_col].str.startswith(tuple(NEG_LIST_CALL), na=False)) & (df[date_col].dt.month == last_m) & (df[date_col].dt.year == last_y)].sort_values(by=date_col, ascending=False)
                df_call = pd.concat([pos_curr_call, neg_curr_call, pos_last_call, neg_last_call], ignore_index=True)

                # --- PROCESS SKIP LOGIC ---
                pos_curr_skip = df[(df[status_col].str.startswith("POS VIA SOCMED", na=False)) & (df[date_col].dt.month == curr_m) & (df[date_col].dt.year == curr_y)].sort_values(by=date_col, ascending=False)
                neg_curr_skip = df[(df[status_col].str.startswith("NEG VIA SOCMED", na=False)) & (df[date_col].dt.month == curr_m) & (df[date_col].dt.year == curr_y)].sort_values(by=date_col, ascending=False)
                pos_last_skip = df[(df[status_col].str.startswith("POS VIA SOCMED", na=False)) & (df[date_col].dt.month == last_m) & (df[date_col].dt.year == last_y)].sort_values(by=date_col, ascending=False)
                neg_last_skip = df[(df[status_col].str.startswith("NEG VIA SOCMED", na=False)) & (df[date_col].dt.month == last_m) & (df[date_col].dt.year == last_y)].sort_values(by=date_col, ascending=False)
                df_skip = pd.concat([pos_curr_skip, neg_curr_skip, pos_last_skip, neg_last_skip], ignore_index=True)

                # Display Results
                st.divider()
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("CALL POS (Curr)", len(pos_curr_call))
                col2.metric("CALL NEG (Curr)", len(neg_curr_call))
                col3.metric("SKIP POS (Curr)", len(pos_curr_skip))
                col4.metric("SKIP NEG (Curr)", len(neg_curr_skip))

                st.divider()
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("CALL POS (Last)", len(pos_last_call))
                col2.metric("CALL NEG (Last)", len(neg_last_call))
                col3.metric("SKIP POS (Last)", len(pos_last_skip))
                col4.metric("SKIP NEG (Last)", len(neg_last_skip))

                st.subheader("CALL Data Preview")
                st.dataframe(df_call.head(10), use_container_width=True)
                st.subheader("SKIP Data Preview")
                st.dataframe(df_skip.head(10), use_container_width=True)

                # Export Logic
                if not df_call.empty or not df_skip.empty:
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df_call.to_excel(writer, index=False, sheet_name="CALL")
                        df_skip.to_excel(writer, index=False, sheet_name="SKIP")
                    output.seek(0)

                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(" Save to Callouts Folder"):
                            ts = datetime.now().strftime("%Y%m%d")
                            save_filename = f"CE_COMBINED_{ts}.xlsx"
                            full_path = os.path.join(SAVE_PATH, save_filename)
                            try:
                                with open(full_path, "wb") as f:
                                    f.write(output.getvalue())
                                st.balloons()
                                st.success(f" Saved to: {full_path}")
                            except Exception as e:
                                st.error(f"Could not save file. Error: {e}")

                    with col2:
                        st.download_button(
                            label=" Download to Browser",
                            data=output.getvalue(),
                            file_name=f"CE_COMBINED_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.info("No records matched the criteria.")
                    
            else:
                st.error(f"Columns not found. Found: {list(df.columns)}. Expected: '{date_col}' and '{status_col}'")

main()