import streamlit as st

def authenticate(username, password):
    if "credentials" in st.secrets:
        credentials = st.secrets["credentials"]
        # Case-insensitive lookup
        for cred_username, data in credentials.items():
            if cred_username.lower() == username.lower():
                if password == data["password"]:
                    return data["role"]
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
    base = ["MVR All Trans", "HDVI-MVR", "Riscom MVR", "PDF Maker", "PDF Play"]
    
    if role == "ADMIN":
        return base
    elif role == "QA":
        return base
    elif role in ["MAKER", "TL"]:
        return ["PDF Play"]
    elif role == "facilitator":
        return ["PDF Maker", "PDF Play"]
    return []

def logout():
    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()
