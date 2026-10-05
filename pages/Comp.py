import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
from io import BytesIO

# Set page layout to wide to accommodate side-by-side comparison tables
st.set_page_config(page_title="Excel S.No Comparator", layout="wide")

# Apply consistent premium design
set_premium_style()


def compare_files(df1, df2):
    # Identify the S.No column (assumed to be the first column)
    sno_col = df1.columns[0]
    
    # Set the index to S.No for both dataframes to align them
    df1 = df1.set_index(sno_col)
    df2 = df2.set_index(sno_col)

    # 1. Identify Missing S.Nos
    missing_in_file2 = df1.index.difference(df2.index)
    missing_in_file1 = df2.index.difference(df1.index)

    # 2. Identify Data Mismatches (Only for S.Nos that exist in both)
    common_ids = df1.index.intersection(df2.index)
    # make sure we only compare columns present in both files
    common_cols = df1.columns.intersection(df2.columns)
    if len(common_cols) == 0:
        # nothing to compare
        mismatches = pd.DataFrame()
    else:
        # compare() returns only the cells that are different; restrict to common columns
        mismatches = df1.loc[common_ids, common_cols].compare(df2.loc[common_ids, common_cols])

    return missing_in_file2, missing_in_file1, mismatches, df1, df2

def main():
    st.warning("Only for S.No")

    # --- UI Header ---
    st.title(" S.No Comparator")
    st.markdown("""
    This tool compares two Excel files based on the **S.No (Column A)**. 
    It will isolate missing rows and highlight specific cell changes.
    """)

    # --- File Upload ---
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        file_a = st.file_uploader("Upload Original File (A)", type=["xlsx", "xls", "csv"])
    with col_u2 :
        file_b = st.file_uploader("Upload Comparison File (B)", type=["xlsx", "xls", "csv"])

    if file_a and file_b:
        # Read Excel Files
        df_a = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)

        # Run Comparison Logic
        missing_b, missing_a, diffs, raw_a, raw_b = compare_files(df_a, df_b)

        # --- Results Display ---
        st.divider()
        
        # Summary Metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("Missing in File B", len(missing_b))
        m2.metric("New in File B", len(missing_a))
        m3.metric("Data Mismatches", len(diffs))

        # Tabs for detailed view
        tab1, tab2, tab3 = st.tabs([" Missing in File B", "➕ New in File B", "🔄 Data Mismatches"])

        with tab1:
            if not missing_b.empty:
                st.error(f"The following S.Nos were found in File A but are missing in File B:")
                st.dataframe(raw_a.loc[missing_b], use_container_width=True)
            else:
                st.success("No S.Nos are missing in File B.")

        with tab2:
            if not missing_a.empty:
                st.warning(f"The following S.Nos are new in File B (not present in File A):")
                st.dataframe(raw_b.loc[missing_a], use_container_width=True)
            else:
                st.success("No new S.Nos detected in File B.")

        with tab3:
            if not diffs.empty:
                st.info("S.No exists in both, but specific values have changed:")
                # Renaming for clarity: 'self' is File A, 'other' is File B
                st.dataframe(diffs.rename(columns={'self': 'Original (A)', 'other': 'Modified (B)'}), use_container_width=True)
            else:
                st.success("All matching S.Nos have identical data across all columns.")

        # --- Export Feature ---
        if not diffs.empty or not missing_b.empty or not missing_a.empty:
            st.divider()
            st.subheader(" Download Discrepancy Report")
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                if not missing_b.empty:
                    raw_a.loc[missing_b].to_excel(writer, sheet_name='Missing_in_B')
                if not missing_a.empty:
                    raw_b.loc[missing_a].to_excel(writer, sheet_name='New_in_B')
                if not diffs.empty:
                    diffs.to_excel(writer, sheet_name='Data_Mismatches')
            
            st.download_button(
                label="Download Excel Report",
                data=output.getvalue(),
                file_name="comparison_report.xlsx",
                mime="application/vnd.ms-excel"
            )
    else:
        st.info("Please upload both Excel files to begin the comparison.")

main()