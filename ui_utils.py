import streamlit as st
import os
import inspect

def set_premium_style():
    """Applies the premium CSS design to the current Streamlit page."""
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
        
        .stApp {
            background: radial-gradient(circle at top left, #1a1c2c, #4a192c);
            font-family: 'Inter', sans-serif;
            color: white;
        }
        
        /* Glassmorphism containers */
        div[data-testid="stVerticalBlock"] > div:has(div.stMarkdown) {
            # background: rgba(255, 255, 255, 0.03);
            # backdrop-filter: blur(10px);
            # border-radius: 15px;
            # padding: 10px;
        }

        /* Titles */
        h1 {
            font-weight: 800 !important;
            background: linear-gradient(to right, #ff4b2b, #ff416c);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 1rem !important;
        }

        /* Sidebar styling */
        [data-testid="stSidebar"] {
            background-color: rgba(26, 28, 44, 0.95) !important;
            border-right: 1px solid rgba(255, 255, 255, 0.1);
        }
        
        /* AGGRESSIVE HIDE for default navigation */
        [data-testid="stSidebarNav"], 
        section[data-testid="stSidebar"] nav,
        div.st-emotion-cache-16ids61 {
            display: none !important;
            height: 0px !important;
            overflow: hidden !important;
        }


        /* Custom scrollbar for sidebar */
        [data-testid="stSidebarContent"]::-webkit-scrollbar {
            width: 4px;
        }
        [data-testid="stSidebarContent"]::-webkit-scrollbar-track {
            background: transparent;
        }
        [data-testid="stSidebarContent"]::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.2);
            border-radius: 10px;
        }

        /* Selectbox styling */
        div[data-baseweb="select"] {
            background: rgba(255, 255, 255, 0.05) !important;
            border-radius: 8px !important;
            border: 1px solid #ff4b2b !important;
        }
        
        div[data-baseweb="select"]:hover, div[data-baseweb="select"]:focus-within {
            border-color: #ff416c !important;
            box-shadow: 0 0 10px rgba(255, 75, 43, 0.3) !important;
        }

        /* Dropdown list styling */
        ul[data-testid="stSelectboxVirtualList"] {
            background-color: #1a1c2c !important;
            border: 1px solid rgba(255, 255, 255, 0.1) !important;
        }


        /* Button styling */
        .stButton>button {
            border-radius: 12px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            background: rgba(255, 255, 255, 0.05);
            color: white;
            transition: all 0.3s ease;
        }

        .stButton>button:hover {
            border-color: #ff4b2b;
            background: rgba(255, 255, 255, 0.1);
            transform: translateY(-2px);
        }

        /* Success/Error boxes */
        div[data-testid="stNotification"] {
            background: rgba(255, 255, 255, 0.05) !important;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.1) !important;
            border-radius: 15px !important;
        }

        /* Dataframes */
        div[data-testid="stDataFrame"] {
            background: rgba(255, 255, 255, 0.05);
            border-radius: 15px;
            padding: 10px;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .main {
            animation: fadeIn 0.8s ease-out;
        }
    </style>

    """, unsafe_allow_html=True)
    
    st.sidebar.markdown("### Choose Automation")
    
    # Custom dropdown navigation
    pages = {
        "Main Dashboard": "Main.py",
        "Worklist Dashboard": "pages/Dashboard.py",
        "Daily Prod": "pages/1_Daily_Prodv3.py",
        "CE Clean": "pages/2_CE_Clean.py",
        "CE CALL": "pages/3_CE_CALL.py",
        "CE COMBINED": "pages/3_CE_COMBINED.py",
        "CE SKIP": "pages/3_CE_SKIP.py",
        "ProdE Auto": "pages/4_ProdE_Auto.py",
        "ProdE Clean": "pages/4_ProdE_Clean.py",
        "FINNONE": "pages/5_FINNONE.py",
        "Comp": "pages/Comp.py",
        "dailyprod": "pages/dailyprod.py",
        "rfd-generator": "pages/rfd-generator.py"
    }
    
    # Find current index based on caller's filename
    try:
        current_file = os.path.basename(inspect.stack()[1].filename)
    except:
        current_file = "Main.py"

    current_idx = 0

    for i, (name, path) in enumerate(pages.items()):
        if os.path.basename(path) == current_file:
            current_idx = i
            break
    
    selected_page_name = st.sidebar.selectbox(
        "Choose Automation",
        options=list(pages.keys()),
        index=current_idx,
        label_visibility="collapsed"
    )
    
    # Only switch if a different page is selected
    if os.path.basename(pages[selected_page_name]) != current_file:
        st.switch_page(pages[selected_page_name])

    st.sidebar.markdown("---")
    if st.sidebar.button("Back to Dashboard", use_container_width=True):
        st.switch_page("Main.py")
    st.sidebar.markdown("---")



