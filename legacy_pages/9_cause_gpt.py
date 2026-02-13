import constants
import streamlit as st
import streamlit_authenticator as stauth

from cause_gpt.gpt import run

st.set_page_config(page_title="Cause Hierarchy GPT", page_icon="🚓")

st.header("Cause Hierarchy GPT 🚓")
st.write("Use this tool to categorize incident descriptions into standardised cause hierarchy categories.")
st.info("""
1. Use for Worker Compensation Clients.
2. Upload only xlsx format file with lossrun_data sheet.
3. Logic will work only on selected column.
4. In output, Cause - Hierarchy 1 column will be added. Download as csv or excel.
""")

config = constants.config

# ✅ FIXED — removed preauthorized
authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"],
)

def main():

    # ✅ New login style
    authenticator.login(location="main")

    authentication_status = st.session_state.get("authentication_status")
    name = st.session_state.get("name")
    username = st.session_state.get("username")

    if authentication_status:

        with st.sidebar:
            st.write(f"Welcome *{name}*")
            authenticator.logout(location="sidebar")

        run()

    elif authentication_status is False:
        st.error("Username/password is incorrect")

    elif authentication_status is None:
        st.warning("Please enter your username and password")


if __name__ == "__main__":
    main()
