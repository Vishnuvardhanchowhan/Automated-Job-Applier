import streamlit as st
import os
import base64
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv
import pandas as pd
from textwrap import dedent
import webbrowser

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="LinkedIn Message Sender", page_icon="💬", layout="wide")
load_dotenv()
user = st.session_state.get("username")  # logged in user
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1k24qvkre4NQhs1gIWNxVkMV8vMPbj069j1OfuMrbNCQ'


# ---------------- AUTH ----------------
def authenticate_google_sheets():
    creds_base64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_BASE64")
    if not creds_base64:
        st.error("❌ Missing credentials. Please set GOOGLE_SERVICE_ACCOUNT_BASE64 in environment.")
        return None
    try:
        creds_json = base64.b64decode(creds_base64).decode("utf-8")
        creds_info = json.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        service = build("sheets", "v4", credentials=creds)
        return service
    except Exception as e:
        st.error(f"⚠️ Error loading credentials: {e}")
        return None


service = authenticate_google_sheets()
if not service:
    st.stop()

def get_sheet_id(spreadsheet, title):
    for sheet in spreadsheet.get("sheets", []):
        if sheet["properties"]["title"] == title:
            return sheet["properties"]["sheetId"]
    return None

ROLE_OPTIONS = {
    "vishnu": ['Data Analyst', 'Data Scientist', 'Data Engineer', 'Machine Learning Engineer',
               'Data Governance Analyst', 'Product Analyst', 'Python Developer'],
    "sakshi": ['Full Stack Developer', 'Frontend Developer', 'Backend Developer', 'Software Developer'],
    "sai": ['Full Stack Engineer', 'Android Developer', 'Frontend Developer', 'Mobile Developer',
            'Software Developer', 'Software Engineer'],
    "harsha": ['Data Analyst', 'Market Researcher', 'Project Manager'],
    "bhanu": ['Full Stack Developer', 'Software Developer', 'Backend Developer']
}

STAGE_OPTIONS = ["Start", "After Reply", "Referral Request", "Follow-up"]

# ---------------- CHECK OR CREATE SHEET ----------------
def ensure_user_sheet(service, spreadsheet_id, user):
    """Ensure a sheet exists for the user, create if not, with dropdown validation."""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        sheet_exists = any(sheet['properties']['title'] == user for sheet in sheets)

        if not sheet_exists:
            # Create sheet
            requests = [{"addSheet": {"properties": {"title": user}}}]
            service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()

        # Initialize headers if sheet is empty
        RANGE_NAME = f"{user}!A1:E1"
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=RANGE_NAME).execute()
        values = result.get("values", [])

        if not values:
            headers = ["Prospect Name", "Prospect Linkedin Profile Link", "Company", "Role", "Stage"]
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=RANGE_NAME,
                valueInputOption="RAW",
                body={"values": [headers]}
            ).execute()
            spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = get_sheet_id(spreadsheet, user)
            if sheet_id is None:
                st.error("❌ Could not get sheet ID for dropdown validation.")
                return False
            requests = [
                # Role column (D)
                {
                    "setDataValidation": {
                        "range": {"sheetId": sheet_id,
                                  "startRowIndex": 1,
                                  "endRowIndex": 1000,
                                  "startColumnIndex": 3,
                                  "endColumnIndex": 4},
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [{"userEnteredValue": v} for v in ROLE_OPTIONS.get(user, [])]
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                },
                # Stage column (E)
                {
                    "setDataValidation": {
                        "range": {"sheetId": sheet_id,
                                  "startRowIndex": 1,
                                  "endRowIndex": 1000,
                                  "startColumnIndex": 4,
                                  "endColumnIndex": 5},
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [{"userEnteredValue": v} for v in STAGE_OPTIONS]
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                }
            ]

            service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()

        return True
    except Exception as e:
        st.error(f"Error creating sheet: {e}")
        return False

def main():
    sheet_exists = ensure_user_sheet(service, SPREADSHEET_ID, user)

    # ---------------- LOAD DATA ----------------
    RANGE_NAME = f"{user}!A:E"
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get("values", [])

    # If sheet just created and empty, ask user to input initial prospect info
    if not sheet_exists or len(values) <= 1:
        st.info("📝 Your sheet is empty. Please add initial prospect info to start.")
        with st.form("new_prospect_form"):
            prospect_name = st.text_input("Prospect Name")
            linkedin_link = st.text_input("LinkedIn Profile Link")
            company_name = st.text_input("Company Name")
            if user == 'vishnu':
                role = st.selectbox("Select the Role",
                                    ['Data Analyst', 'Data Scientist', 'Data Engineer', 'Machine Learning Engineer',
                                     'Data Governance Analyst', 'Product Analyst', 'Python Developer'])
            elif user == 'sakshi':
                role = st.selectbox("Select the Role", ['Full Stack Developer', 'Frontend Developer', 'Backend Developer',
                                                        'Software Developer'])
            elif user == 'sai':
                role = st.selectbox("Select the Role",
                                    ['Full Stack Engineer', 'Android Developer', 'Frontend Developer', 'Mobile Developer',
                                     'Software Developer', 'Software Engineer'])
            elif user == 'harsha':
                role = st.selectbox("Select the Role", ['Data Analyst', 'Market Researcher', 'Project Manager'])
            elif user == 'bhanu':
                role = st.selectbox("Select the Role", ['Full Stack Developer', 'Software Developer', 'Backend Developer'])
            stage = st.selectbox("Stage", ["Start", "After Reply", "Referral Request", "Follow-up"])
            submitted = st.form_submit_button("Add Prospect")
            if submitted:
                if all([prospect_name, linkedin_link, company_name, role, stage]):
                    body = {"values": [[prospect_name, linkedin_link, company_name, role, stage]]}
                    service.spreadsheets().values().append(
                        spreadsheetId=SPREADSHEET_ID,
                        range=RANGE_NAME,
                        valueInputOption="RAW",
                        body=body
                    ).execute()
                    st.success(f"✅ Prospect {prospect_name} added successfully!")
                    st.rerun()
                else:
                    st.error("⚠️ Please fill all fields.")

    # ---------------- LOAD DATAFRAME ----------------
    headers = values[0]
    rows = values[1:] if len(values) > 1 else []
    df = pd.DataFrame(rows, columns=headers).dropna(subset=["Prospect Name"])
    df.set_index("Prospect Name", inplace=True)

    # ---------------- USER-SPECIFIC MESSAGE TEMPLATES ----------------
    user_templates = {
        "vishnu": dedent("""\
            Hi {Name},
    
            Thank you for connecting! I came across your profile and was impressed by the work at {Company}.
    
            I specialize in data analytics, Python, SQL, and automation, and I’m currently looking for opportunities where I can contribute my {Role} skills.
    
            I’d love to know if there are any openings or upcoming projects where I could be a good fit.
        """),

        "sakshi": dedent("""\
            Hi {Name},
    
            Thank you for accepting my connection! I noticed the amazing work being done at {Company}.
    
            I am experienced in full-stack development, including React, Node.js, and TypeScript, and I’m eager to contribute my {Role} skills.
    
            Are there any roles or projects where I could be helpful?
        """),

        "harsha": dedent("""\
            Hello {Name},
    
            Thank you for connecting! I am impressed by the initiatives at {Company}.
    
            I specialize in data analytics, market research, and project management, and I’d love to contribute my {Role} skills to your team.
    
            Could you let me know if there are any roles or opportunities I can assist with?
        """),

        "sai": dedent("""\
            Hi {Name},
    
            Thanks for connecting! I saw the work at {Company} and found it very inspiring.
    
            I am a versatile software engineer experienced in full-stack and mobile development, and I’d love to contribute my {Role} skills to your team.
    
            Are there any openings or upcoming projects that might be a good fit?
        """),

        "bhanu": dedent("""\
            Hello {Name},
    
            Thank you for connecting! I am impressed by {Company}’s tech initiatives.
    
            My expertise lies in backend and full-stack development, and I am eager to apply my {Role} skills to contribute effectively.
    
            Would you be able to guide me on any opportunities where I could add value?
        """)
    }


    common_dict = {
        "After Reply": dedent("""\
            Hi {Name},
    
            Thank you so much for getting back to me! 🙏
    
            I’d love to share my resume and explore if my experience aligns with any opportunities at {Company}. Your guidance or feedback would mean a lot!
        """),

        "Referral Request": dedent("""\
            Hi {Name},
    
            I hope you’re doing well! 🌟
    
            I noticed a {Role} position at {Company} that matches my skills and experience perfectly. If you think I could be a fit, I would greatly appreciate your referral. 
    
            I’m really excited about the chance to contribute and grow with your team!
        """),

        "Follow-up": dedent("""\
            Hi {Name},
    
            I hope all is well! Just following up on my previous message to see if you had a chance to review my profile.
    
            I remain very interested in opportunities at {Company} and would love to contribute my {Role} skills to your team. Looking forward to your thoughts! 🙂
        """)
    }


    # ---------------- SELECT PROSPECT ----------------
    st.sidebar.title("💬 LinkedIn Message Sender")
    # Option to add new prospect
    add_new = st.sidebar.checkbox("➕ Add New Prospect")

    if add_new:
        st.sidebar.subheader("Add New Prospect")
        new_name = st.sidebar.text_input("Prospect Name")
        new_link = st.sidebar.text_input("LinkedIn Profile Link")
        new_company = st.sidebar.text_input("Company Name")

        # Dynamic role options based on user
        role_options = ROLE_OPTIONS.get(user, [])
        new_role = st.sidebar.selectbox("Role", role_options)
        new_stage = st.sidebar.selectbox("Stage", STAGE_OPTIONS)

        if st.sidebar.button("Add Prospect"):
            if new_name not in df.index:
                if all([new_name, new_link, new_company, new_role, new_stage]):
                    # Append to Google Sheet
                    RANGE_NAME = f"{user}!A:E"
                    body = {"values": [[new_name, new_link, new_company, new_role, new_stage]]}
                    try:
                        service.spreadsheets().values().append(
                            spreadsheetId=SPREADSHEET_ID,
                            range=RANGE_NAME,
                            valueInputOption="RAW",
                            body=body
                        ).execute()

                        # Update local DataFrame immediately
                        df.loc[new_name] = [new_link, new_company, new_role, new_stage]
                        st.success(f"✅ Prospect '{new_name}' added successfully!")
                    except Exception as e:
                        st.error(f"⚠️ Failed to add prospect: {e}")
                else:
                    st.warning("⚠️ Please fill all fields.")
            else:
                st.warning("⚠️ Please choose different name as name you added already exist in sheet!")

    if not df.empty:
        name = st.sidebar.selectbox("Prospect Name", df.index)
        profile_link = df.loc[name, 'Prospect Linkedin Profile Link']
        company_name = df.loc[name, 'Company']
        role = df.loc[name, 'Role']
        stage = df.loc[name, 'Stage']

        # ---------------- MESSAGE LOGIC ----------------
        common_dict['Start'] = user_templates.get(user, user_templates["sakshi"])
        message_filled = common_dict.get(stage, common_dict["Start"]).format(Name=name, Company=company_name,
                                                                                   Role=role)

        # ---------------- UI ----------------
        st.markdown("<h1 style='text-align:center;'>💼 LinkedIn Message Sender</h1>", unsafe_allow_html=True)
        st.divider()
        col1, col2 = st.columns([1.1, 2])

        with col1:
            st.subheader("👤 Prospect Details")
            st.write(f"**Name:** {name}")
            st.write(f"**Company:** {company_name}")
            st.write(f"**Role:** {role}")
            st.write(f"**Stage:** {stage}")
            st.markdown(f"[🔗 View LinkedIn Profile]({profile_link})", unsafe_allow_html=True)

        with col2:
            st.subheader("💬 Auto-Generated Message")
            st.text_area("Generated Message", value=message_filled, height=260)
            colA, colB = st.columns(2)
            with colA:
                if st.button("📋 Copy Message"):
                    st.session_state["copied_message"] = message_filled
                    st.success("✅ Message copied! Paste into LinkedIn.")
            with colB:
                if st.button("🔗 Open LinkedIn Profile"):
                    webbrowser.open_new_tab(profile_link)
                    st.info("🌐 LinkedIn profile opened in browser.")
if __name__ == "__main__":
    main()