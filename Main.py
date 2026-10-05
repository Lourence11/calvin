import streamlit as st
import os
from ui_utils import set_premium_style

st.set_page_config(page_title="Automation Hub", layout="wide", initial_sidebar_state="collapsed")

# Apply consistent premium design
set_premium_style()

# Custom Dashboard Title Animation (kept in Main.py as it's unique)
st.markdown("""
<style>
    .main-title {
        font-size: 3.5rem;
        font-weight: 800;
        background: linear-gradient(to right, #ff4b2b, #ff416c);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 0.5rem;
        animation: fadeIn 1.5s ease-out;
    }
    
    .subtitle {
        color: #b0b0b0;
        text-align: center;
        font-size: 1.2rem;
        margin-bottom: 3rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-title">Automation Hub</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Select a tool to launch it efficiently "one by one".</p>', unsafe_allow_html=True)

# Dynamic Page Mapping
PAGE_MAP = {
    "Worklist Dashboard": {"path": "pages/Dashboard.py", "desc": "Upload and inspect password-protected worklists at lightning speed without COM."},
    "Daily Prod": {"path": "pages/1_Daily_Prodv3.py", "desc": "Process daily production reports with cleaning logic."},
    "Collection Efforts Clean": {"path": "pages/2_CE_Clean.py", "desc": "Clean and format collection effort data."},
    "Collection Efforts Callout": {"path": "pages/3_CE_CALL.py", "desc": "Generate callout lists for collections."},
    "Collection Efforts Skip": {"path": "pages/3_CE_SKIP.py", "desc": "Filter and process skip trace data."},
    "Daily Prods accu": {"path": "pages/dailyprod.py", "desc": "Accumulate and analyze daily production trends."},
    "Collection Efforts Combined": {"path": "pages/3_CE_COMBINED.py", "desc": "Merge multiple collection effort datasets."},
    "Compare": {"path": "pages/Comp.py", "desc": "Compare different report versions side-by-side."},
    "Productivity Efforts Clean": {"path": "pages/4_ProdE_Clean.py", "desc": "Refine productivity effort reports."},
    "RFD Generator": {"path": "pages/rfd-generator.py", "desc": "Automated Reason for Delinquency generation."},
}

# Create grid of buttons
cols = st.columns(3)
for i, (name, info) in enumerate(PAGE_MAP.items()):
    col = cols[i % 3]
    with col:
        st.markdown(f"""
        <div style="margin-bottom: 20px;">
        """, unsafe_allow_html=True)
        if st.button(name, key=name, use_container_width=True, type="secondary"):
             st.switch_page(info['path'])

        st.markdown(f"""
            <div style="color: #666; font-size: 0.8rem; text-align: center; margin-top: -10px; padding: 0 10px;">
                {info['desc']}
            </div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")
with st.sidebar:
    st.info("**Memory Optimization**: Tools are only loaded on demand. This 'one-by-one' approach saves system resources.")
    if st.button("Home", use_container_width=True):
        st.rerun()
