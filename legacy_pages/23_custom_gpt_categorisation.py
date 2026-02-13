import constants
import streamlit as st
import streamlit_authenticator as stauth

from custom_gpt_categorisation.gpt import run

st.set_page_config(page_title="Custom Categorisation GPT", page_icon="🚓")

st.header("Custom Categorisation GPT 🚓")
st.write("Use this tool to categorize incident descriptions into standardised injury type categories.")
st.info("""
1. Only for FCCI client.
2. Upload only xlsx format file with lossrun_data sheet.
3. Logic will work only on selected column.
4. In output, Injury Type column will be added. Download as csv or excel.
""")

config = constants.config

authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"],
)

def main():
    # New login method (stores values in session_state)
    authenticator.login(location="main")

    if st.session_state.get("authentication_status"):

        with st.sidebar:
            st.write(f"Welcome *{st.session_state.get('name')}*")
            authenticator.logout(location="sidebar")

        run()

    elif st.session_state.get("authentication_status") is False:
        st.error("Username/password is incorrect")

    elif st.session_state.get("authentication_status") is None:
        st.warning("Please enter your username and password")


if __name__ == "__main__":
    main()
