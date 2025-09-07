import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth

# Page config
st.set_page_config(page_title="Login", page_icon="🔐", initial_sidebar_state="collapsed")

# Hide sidebar & default UI
st.markdown(
    """
    <style>
        div[data-testid="stSidebarNav"] {display: none;}
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# Load config
try:
    with open("config.yaml") as file:
        config = yaml.load(file, Loader=SafeLoader)
except FileNotFoundError:
    st.error("Configuration file 'config.yaml' not found.")
    st.stop()
except Exception as e:
    st.error(f"Error loading configuration: {e}")
    st.stop()

# Build authenticator
try:
    authenticator = stauth.Authenticate(
        config["credentials"],
        config["cookie"]["name"],
        config["cookie"]["key"],
        config["cookie"]["expiry_days"],
    )
except KeyError as e:
    st.error(f"Missing required configuration: {e}")
    st.stop()

# Initialize session state for authentication
if 'authentication_status' not in st.session_state:
    st.session_state.authentication_status = None
if 'username' not in st.session_state:
    st.session_state.username = None
if 'name' not in st.session_state:
    st.session_state.name = None

# -------------------------------
# LOGIN FORM
# -------------------------------
# Use the login method without unpacking
authenticator.login(location="main")

# Check if authentication was successful by examining session state
if (st.session_state.get("authentication_status") and
        st.session_state.get("username") and
        st.session_state.get("name")):

    # -------------------------------
    # DASHBOARD (only if authenticated)
    # -------------------------------
    st.sidebar.success(f"✅ Logged in as {st.session_state['name']}")

    try:
        user_email = config['credentials']['usernames'][st.session_state['username']]['email']
        st.sidebar.write(f"Email: {user_email}")
    except KeyError:
        st.sidebar.warning("Email not found in configuration")

    # Logout button
    if st.sidebar.button("🚪 Logout"):
        authenticator.logout()
        # Clear session state
        for key in ["authentication_status", "username", "name"]:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()  # Rerun to show login form

    # Hidden navigation to Mail Drafter dashboard
    try:
        pages = {
            "Main": [
                st.Page("pages/mail_drafter.py", title="Mail Drafter", icon="✉️"),
            ]
        }
        pg = st.navigation(pages, position="hidden")
        pg.run()
    except Exception as e:
        st.error(f"Error loading navigation: {e}")

elif st.session_state.get("authentication_status") is False:
    st.error("❌ Username/password is incorrect")

elif st.session_state.get("authentication_status") is None:
    st.warning("🔑 Please enter your username and password")