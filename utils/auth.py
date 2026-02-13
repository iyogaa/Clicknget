import streamlit as st

def authenticate(username, password):
    if "credentials" in st.secrets and username in st.secrets["credentials"]:
        user_data = st.secrets["credentials"][username]
        if password == user_data["password"]:
            return user_data["role"]
    return None

def init_session_state():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
        st.session_state["username"] = None
        st.session_state["role"] = None

def show_login():
    with st.sidebar:
        st.title("🔐 Login")
        username = st.text_input("Username", key="login_user")
        password = st.text_input("Password", type="password", key="login_pass")
        if st.button("Login"):
            role = authenticate(username, password)
            if role:
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.session_state["role"] = role
                st.rerun()
            else:
                st.error("Invalid username or password")

def get_menu_options(role):
    # Base tools (MVR and PDF)
    base = ["MVR All Trans", "HDVI-MVR", "PDF Maker", "PDF Play"]
    # GPT tools
    gpt = ["Body GPT", "Cause GPT", "Accident GPT", "Custom GPT Categorization"]
    
    if role == "ADMIN":
        return base + gpt
    elif role in "QA":
        return base + gpt
    elif role in ["MAKER", "TL"]:
        return "PDF Play"
    return []

def logout():
    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()
