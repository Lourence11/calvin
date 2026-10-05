import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui_utils import set_premium_style

import streamlit as st
import pandas as pd
from io import BytesIO

st.title(" RFD Generator Automation")

def main():
    uploaded_file = st.file_uploader(
        "Upload CSV or Excel file",
        type=["csv", "xlsx"]
    )

    RFD_KEYWORDS = [
        "WITH OTHER EXPENSES",
        "SHORT OF FUNDS",
        "BUSINESS BANKRUPTCY",
        "WITH OTHER LOANS TO PAY",
        "CLIENT IS UNDER MEDICATION",
        "GOT SCAMMED",
        "SCAMMED",
        "GOT SCAM",
        "UNEMPLOYED",
        "NO JOB",
        "DELAYED FUNDS",
        "VICTIM OF CALAMITY",
        "OOTC",
        "BUSINESS SLOWSDOWN",
        "CLAIMING FULLY PAID",
        "FAMILY MEMBER IS UNDER MEDICATION",
        "CLAIMING NO OBLIGATION",
        "CLAIMING DISPUTE",
        "WITH RIDER",
        "DECEASED",
        "THIRD PARTY USED THE CREDIT CARD",
        "WITH UNAUTHORIZED TRANSACTION",
        "OVERLOOKED THE ACCOUNT",
        "GAVE BIRTH",
        "DEATH OF FAMILY MEMBER",
        
    ]

    ootc_aliases = [
        "OUT OF THE COUNTRY"
    ]

    def get_rfd(remark):
        for alias in ootc_aliases:
            if alias in remark:
                return "OOTC"
            
        for keyword in RFD_KEYWORDS:
            if keyword in remark:
                if keyword in ["SCAMMED", "GOT SCAM"]:
                    return "GOT SCAMMED"
                return keyword
        return ""

    if uploaded_file:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, low_memory=False)
        else:
            df = pd.read_excel(uploaded_file)

        # Optimization: Convert all columns to string to avoid Arrow conversion errors with mixed types
        # and prevent MessageSizeError for massive dataframes
        for col in df.columns:
            df[col] = df[col].astype(str)

        if "Remark" not in df.columns:
            st.error(" 'Remark' column not found.")
        else:
            df["Remark"] = df["Remark"].astype(str).str.upper()
            df["RFD"] = df["Remark"].apply(get_rfd)

            st.success(" RFD column generated successfully!")
            
            # Display limit to avoid MessageSizeError and browser lag
            if len(df) > 100:
                st.warning(f" Showing first 100 rows of {len(df)} total rows to save memory.")
                st.dataframe(df.head(100))
            else:
                st.dataframe(df)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Result")
            output.seek(0)

            st.download_button(
                label=" Download Updated File (Excel)",
                data=output,
                file_name="rfd-result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

main()
# Apply consistent premium design
set_premium_style()
