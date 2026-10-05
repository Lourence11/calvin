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
st.set_page_config(page_title="Collection Efforts Automation", layout="wide")

# Apply consistent premium design
set_premium_style()


PAGE_MAP = {
    "Home": "Main.py",
    "Daily Prod": "pages/1_Daily_Prod.py",
    "Collection Efforts Clean": "pages/2_CE_Clean.py",
    "Collection Efforts Callout": "pages/3_CE_CALL.py",
    "Collection Efforts Skip": "pages/3_CE_SKIP.py",
    "Daily Prods accu": "pages/dailyprod.py",
}

# --- STRICT LIST DEFINITIONS ---
POS_LIST = [
    "POS VIA SOCMED"
]

NEG_LIST = [
    "NEG VIA SOCMED"
]

def main():
    # Ensure the save directory exists only when this page is active
    SAVE_PATH = r"C:\Users\SPM\Downloads\BPI\Collection Efforts_Every Friday\CE Auto"
    os.makedirs(SAVE_PATH, exist_ok=True)

    st.title("Collection Efforts Automation (Skip)")
    st.success("Remember to add '000000' in account number and date format is short date")
    uploaded_file = st.file_uploader("Drop your file here", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        # 1. Load Data
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"Loaded: {uploaded_file.name}")
        except Exception as e:
            st.error(f"Error loading file: {e}")
            st.stop()

        # 2. Setup Columns
        date_col = 'Date' 
        status_col = 'Status' 

        if date_col in df.columns and status_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
            
            invalid_dates = df[date_col].isna().sum()
            if invalid_dates > 0:
                st.warning(f"Skipped {invalid_dates} rows due to invalid date formats.")
                df = df.dropna(subset=[date_col])

            # 3. Calculate Date Ranges
            today = datetime.today()
            curr_m, curr_y = today.month, today.year
            last_month_dt = today - relativedelta(months=1)
            last_m, last_y = last_month_dt.month, last_month_dt.year

            # 4. Strict Hierarchical Extraction
            pos_curr = df[(df[status_col].str.startswith("POS VIA SOCMED", na=False)) & (df[date_col].dt.month == curr_m) & (df[date_col].dt.year == curr_y)].sort_values(by=date_col, ascending=False)
            neg_curr = df[(df[status_col].str.startswith("NEG VIA SOCMED", na=False)) & (df[date_col].dt.month == curr_m) & (df[date_col].dt.year == curr_y)].sort_values(by=date_col, ascending=False)
            pos_last = df[(df[status_col].str.startswith("POS VIA SOCMED", na=False)) & (df[date_col].dt.month == last_m) & (df[date_col].dt.year == last_y)].sort_values(by=date_col, ascending=False)
            neg_last = df[(df[status_col].str.startswith("NEG VIA SOCMED", na=False)) & (df[date_col].dt.month == last_m) & (df[date_col].dt.year == last_y)].sort_values(by=date_col, ascending=False)

            # Combine in exact order (newest to oldest within each group)
            final_df = pd.concat([pos_curr, neg_curr, pos_last, neg_last], ignore_index=True)

            # 5. Display Results
            st.divider()
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("POS (Curr)", len(pos_curr))
            col2.metric("NEG (Curr)", len(neg_curr))
            col3.metric("POS (Last)", len(pos_last))
            col4.metric("NEG (Last)", len(neg_last))

            st.subheader("Final Hierarchical Data")
            st.dataframe(final_df, use_container_width=True)

            # 6. Export Logic
            if not final_df.empty:
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    final_df.to_excel(writer, index=False, sheet_name="Callouts")
                    pd.DataFrame({"POS": sorted(df[df[status_col].str.startswith("POS ", na=False)][status_col].unique())}).to_excel(writer, index=False, sheet_name="POS_LIST")
                    pd.DataFrame({"NEG": sorted(df[df[status_col].str.startswith("NEG ", na=False)][status_col].unique())}).to_excel(writer, index=False, sheet_name="NEG_LIST")
                output.seek(0)

                col1, col2 = st.columns(2)
                with col1:
                    if st.button(" Save to Callouts Folder"):
                        ts = datetime.now().strftime("%Y%m%d")
                        save_filename = f"CE_SKIP_{ts}.xlsx"
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
                        file_name=f"CE_SKIP{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("No records matched the POS/NEG criteria for the current or last month.")
                
        else:
            st.error(f"Columns not found. Found: {list(df.columns)}. Expected: '{date_col}' and '{status_col}'")

main()