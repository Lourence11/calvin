import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
from io import BytesIO

def main():
    st.title("Production Efforts Clean")

    uploaded_file = st.file_uploader("Upload File", type=["xlsx", "csv"])

    if uploaded_file:
        # 1. Load Data
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        
        # --- FILTER ROWS BASED ON STATUS AND REMARK BY ---
        st.subheader(" Filtering Data")
        
        # Track removed rows for reporting
        total_rows_before = len(df)
        removed_status_rows = 0
        removed_remarkby_rows = 0
        
        # Filter Status: Remove rows with "BP"
        if 'Status' in df.columns:
            status_before = len(df)
            df = df[df['Status'].astype(str) != 'BP']
            removed_status_rows = status_before - len(df)
            if removed_status_rows > 0:
                st.warning(f" Removed {removed_status_rows} row(s) with Status = 'BP'")
        else:
            st.warning(" 'Status' column not found - cannot filter by Status")
        
        # Filter Remark By: Remove rows with specified names
        if 'Remark By' in df.columns:
            remark_before = len(df)
            excluded_remark_by = ["DCCAUNTE", "CMENRIQUEZ", "EDSUMAIT", "JSCANA", "NRPARAYAOAN", "CNOCON", "AORBASIDO", "OQUARINO"]
            
            # Create a boolean mask for rows to keep (not in excluded list)
            mask = ~df['Remark By'].astype(str).str.upper().isin(excluded_remark_by)
            df = df[mask]
            
            removed_remarkby_rows = remark_before - len(df)
            if removed_remarkby_rows > 0:
                st.warning(f" Removed {removed_remarkby_rows} row(s) with Remark By in: {', '.join(excluded_remark_by)}")
        else:
            st.warning(" 'Remark By' column not found - cannot filter by Remark By")
        
        # Summary of filtering
        total_removed = removed_status_rows + removed_remarkby_rows
        if total_removed > 0:
            st.success(f" Total rows removed: {total_removed} (from {total_rows_before} to {len(df)} rows)")
            
        # --- REMARK SANITIZATION ---
        if 'Remark' in df.columns:
            # Remove leading '=' to prevent Excel from treating remarks as formulas and showing #NAME?
            df['Remark'] = df['Remark'].astype(str).str.lstrip('=')
        
        # --- DELETE LAST 2 COLUMNS ---
        if len(df.columns) > 2:
            cols_to_remove = df.columns[-2:].tolist()
            df = df.iloc[:, :-2]
            st.warning(f" Deleted last 2 columns: {', '.join(cols_to_remove)}")
        
        # --- MOVE LOGIC ---
        column_to_move = "Next Call"
        target_position = 23  # Moves to Column X (index 23)

        # Try to find Next Call column
        next_call_col = None
        if column_to_move in df.columns:
            next_call_col = column_to_move
        else:
            for col in df.columns:
                if 'next' in str(col).lower() and 'call' in str(col).lower():
                    next_call_col = col
                    st.info(f"Found Next Call column: '{next_call_col}'")
                    break
        
        if next_call_col:
            actual_target = min(target_position, len(df.columns))
            
            st.write(f"Moving **{next_call_col}** to index **{actual_target}**...")
            
            col = df.pop(next_call_col)
            df.insert(actual_target, next_call_col, col)
            
            st.success(" Column Shifted and Cleanup Complete!")
            
            # Show preview
            st.subheader("Preview of processed data:")
            st.dataframe(df.head(10))
        else:
            st.error(f"Column '{column_to_move}' not found in this file.")

        # --- DOWNLOAD SECTION ---
        output = BytesIO()

        if uploaded_file.name.endswith('.csv'):
            df.to_csv(output, index=False)
            file_name = f"ProdE_Clean_{uploaded_file.name.split('.')[0]}.csv"
            mime_type = "text/csv"
        else:
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            file_name = f"ProdE_Clean_{uploaded_file.name.split('.')[0]}.xlsx"
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        output.seek(0)

        st.download_button(
            label=" Download Processed File",
            data=output,
            file_name=file_name,
            mime=mime_type
        )

main()
# Apply consistent premium design
set_premium_style()
