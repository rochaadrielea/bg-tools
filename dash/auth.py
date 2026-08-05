"""
auth.py — password gate for Streamlit Community Cloud.
Password is stored in the app's Settings -> Secrets (never in the repo).
"""
import streamlit as st


def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True

    stored = st.secrets.get("app_password", None)
    if stored is None:
        st.error("No password configured. Set 'app_password' in Settings -> Secrets.")
        return False

    st.title("Quality BRM Dashboard")
    st.caption("Please sign in to continue")
    entered = st.text_input("Password", type="password", key="pw_input")

    if st.button("Sign in"):
        if entered == stored:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Wrong password.")
    return False