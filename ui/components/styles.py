import streamlit as st

def apply_custom_css():
    """Apply minimal dark grey professional styling to Streamlit app"""
    st.markdown("""
    <style>
    /* ==========================================
       CLICKNGET - MINIMAL DARK GREY UI
       ========================================== */

    :root {
        --bg-1: #0a0a0a;
        --bg-2: #1a1a1a;
        --bg-3: #2a2a2a;
        --bg-hover: #333333;
        
        --text-primary: #ffffff;
        --text-secondary: #d0d0d0;
        --text-muted: #808080;
        
        --border-color: #3a3a3a;
        --shadow-sm: 0 2px 4px rgba(0, 0, 0, 0.3);
        --shadow-md: 0 4px 8px rgba(0, 0, 0, 0.4);
        
        --radius: 6px;
        --transition: all 0.15s ease;
    }

    /* ==========================================
       GLOBAL STYLES
       ========================================== */

    body, .stApp {
        background-color: var(--bg-1);
        color: var(--text-primary);
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: var(--bg-2);
        border-right: 1px solid var(--border-color);
    }

    /* ==========================================
       TYPOGRAPHY
       ========================================== */

    h1, h2, h3, h4, h5, h6 {
        color: var(--text-primary);
        font-weight: 600;
        margin-bottom: 1rem;
    }

    h1 {
        font-size: 2rem;
        margin-bottom: 1.5rem;
    }

    h2 {
        font-size: 1.5rem;
    }

    h3 {
        font-size: 1.25rem;
    }

    p {
        color: var(--text-secondary);
        line-height: 1.5;
    }

    /* ==========================================
       CARDS & CONTAINERS
       ========================================== */

    .metric-card {
        background: var(--bg-2);
        border: 1px solid var(--border-color);
        border-radius: var(--radius);
        padding: 1.25rem;
        transition: var(--transition);
        box-shadow: var(--shadow-sm);
    }

    .metric-card:hover {
        background: var(--bg-3);
        border-color: var(--text-secondary);
        box-shadow: var(--shadow-md);
    }

    .card-title {
        font-size: 0.8rem;
        color: var(--text-muted);
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 0.5rem;
    }

    .card-value {
        font-size: 1.5rem;
        color: var(--text-primary);
        font-weight: 700;
        margin-bottom: 0.25rem;
    }

    .card-description {
        font-size: 0.8rem;
        color: var(--text-muted);
    }

    .nav-card {
        background: var(--bg-2);
        border: 1px solid var(--border-color);
        border-radius: var(--radius);
        padding: 1rem;
        text-align: center;
        cursor: pointer;
        transition: var(--transition);
        box-shadow: var(--shadow-sm);
        display: inline-block;
        min-width: 150px;
        margin: 0.5rem;
        color: var(--text-primary);
        text-decoration: none;
    }

    .nav-card:hover {
        background: var(--bg-3);
        border-color: var(--text-secondary);
        box-shadow: var(--shadow-md);
        transform: translateY(-2px);
    }

    /* ==========================================
       BUTTONS & INPUTS
       ========================================== */

    .stButton > button {
        background: var(--bg-3);
        color: var(--text-primary);
        border: 1px solid var(--border-color);
        border-radius: var(--radius);
        padding: 0.6rem 1.2rem;
        font-weight: 500;
        cursor: pointer;
        transition: var(--transition);
    }

    .stButton > button:hover {
        background: var(--bg-hover);
        border-color: var(--text-secondary);
    }

    .stTextInput > div > div > input,
    .stSelectbox > div > div > select,
    .stTextArea > div > div > textarea,
    .stNumberInput > div > div > input {
        background-color: var(--bg-2) !important;
        border: 1px solid var(--border-color) !important;
        color: var(--text-primary) !important;
        border-radius: var(--radius) !important;
        padding: 0.6rem 0.8rem !important;
    }

    /* ==========================================
       DIVIDERS
       ========================================== */

    hr {
        border: none;
        border-top: 1px solid var(--border-color);
        margin: 1.5rem 0;
    }

    /* ==========================================
       RESPONSIVE
       ========================================== */

    @media (max-width: 768px) {
        h1 { font-size: 1.5rem; }
        h2 { font-size: 1.25rem; }
        .nav-card { min-width: 100px; padding: 0.75rem; }
    }

    </style>
    """, unsafe_allow_html=True)
