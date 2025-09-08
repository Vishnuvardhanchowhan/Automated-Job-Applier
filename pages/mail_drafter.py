import os
import smtplib
from datetime import date
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from textwrap import dedent
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
import os
import base64
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
load_dotenv()
creds_base64 = os.getenv('GOOGLE_CREDENTIALS_BASE64')

def authenticate_google_sheets():
    """Authenticate with Google Sheets API using environment variables."""
    creds = None
    token_base64 = os.getenv('GOOGLE_TOKEN_BASE64')
    if token_base64 and os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    creds_base64 = os.getenv('GOOGLE_CREDENTIALS_BASE64')
    if creds_base64:
        try:
            creds_json = base64.b64decode(creds_base64).decode('utf-8')
            creds_info = json.loads(creds_json)
            flow = InstalledAppFlow.from_client_config(creds_info, SCOPES)
            creds = flow.run_local_server(port=0)
            print("Using service account credentials from environment variable")
        except Exception as e:
            print(f"Error decoding service account credentials: {e}")
            return None
    elif not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            client_id = os.getenv('GOOGLE_CLIENT_ID')
            client_secret = os.getenv('GOOGLE_CLIENT_SECRET')
            if client_id and client_secret:
                client_config = {
                    "web": {
                        "client_id": client_id,
                        "client_secret": client_secret,
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token"
                    }
                }
                flow = InstalledAppFlow.from_client_config(
                    client_config,
                    SCOPES
                )
            else:
                if not os.path.exists("credentials.json"):
                    print(
                        "No credentials found. Please set GOOGLE_CREDENTIALS_BASE64 or GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET environment variables.")
                    return None
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )

            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    try:
        service = build("sheets", "v4", credentials=creds)
        return service
    except HttpError as err:
        print(f"Authentication error: {err}")
        return None

def ensure_user_sheet_exists(service, spreadsheet_id, user_name):
    try:
        # Get all sheets in the spreadsheet
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        sheet_exists = any(sheet['properties']['title'] == user_name for sheet in sheets)

        if not sheet_exists:
            # Create a new sheet for the user
            requests = [{
                "addSheet": {
                    "properties": {
                        "title": user_name
                    }
                }
            }]
            body = {'requests': requests}
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            print(f"Created new sheet for user: {user_name}")

            # Initialize headers in the new sheet
            headers = [
                "Date", "Company", "Role", "Job ID", "Recruiter Name",
                "Recruiter Emails", "Subject Line", "Why Join", "JD", "Status"
            ]
            body = {'values': [headers]}
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{user_name}!A1:J1",
                valueInputOption="RAW",
                body=body
            ).execute()
        else:
            print(f"Sheet for user {user_name} already exists.")
    except HttpError as err:
        print(f"Error ensuring sheet exists: {err}")

def log_application(service, spreadsheet_id, user_name, details):
    ensure_user_sheet_exists(service, spreadsheet_id, user_name)
    try:
        body = {'values': [details]}
        result = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{user_name}!A:J",
            valueInputOption="RAW",
            body=body
        ).execute()
        print(f"Logged application for user {user_name}: {result.get('updates').get('updatedCells')} cells updated.")
    except HttpError as err:
        print(f"Error logging application: {err}")

# ---------------- PDF GENERATOR ----------------
def generate_cover_letter_pdf(text, name, official_role, filename):
    doc = SimpleDocTemplate(
        filename, pagesize=LETTER,
        rightMargin=20, leftMargin=20,
        topMargin=20, bottomMargin=20
    )
    styles = getSampleStyleSheet()

    # Custom styles
    name_style = ParagraphStyle(
        "NameStyle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=20,
        textColor=colors.darkblue,
        alignment=0,  # left align
        spaceAfter=12
    )

    subtitle_style = ParagraphStyle(
        "SubtitleStyle",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=12,
        textColor=colors.grey,
        alignment=0,
        spaceAfter=18
    )

    body_style = ParagraphStyle(
        "BodyStyle",
        parent=styles["Normal"],
        fontName="Times-Roman",
        fontSize=12,
        leading=18,
        spaceAfter=12
    )

    bullet_style = ParagraphStyle(
        "BulletStyle",
        parent=body_style,
        leftIndent=14,
        bulletIndent=0,
        bulletFontName="Helvetica-Bold"
    )

    elements = []

    # Add name and subtitle
    elements.append(Paragraph(name, name_style))
    elements.append(Paragraph(official_role, subtitle_style))

    # Add a subtle line separator
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.grey, spaceBefore=6, spaceAfter=18))

    # Split the rest of text into paragraphs
    for para in text.split("\n\n"):
        para = para.strip().replace("\n", "<br/>")
        if para.startswith("• "):
            # Make bullet points nicer
            elements.append(Paragraph(para.replace("• ", ""), bullet_style, bulletText="•"))
        else:
            elements.append(Paragraph(para, body_style))
        elements.append(Spacer(1, 6))

    doc.build(elements)
    return filename

# def log_application(filename, details):
#     if not os.path.exists(filename):
#         wb = openpyxl.Workbook()
#         ws = wb.active
#         ws.title = "Applications"
#         ws.append([
#             "Date", "Company", "Role", "Job ID", "Recruiter Name",
#             "Recruiter Emails", "Subject Line", "Why Join", 'JD', "Status"
#         ])
#         wb.save(filename)
#
#     wb = openpyxl.load_workbook(filename)
#     ws = wb.active
#     ws.append(details)
#     wb.save(filename)

class EmailSender:
    def __init__(self, role, recruiter, role_name, company_name, why_company=None):
        self.role = role
        self.recruiter = recruiter
        self.role_name = role_name
        self.company_name = company_name
        self.why_company = why_company or "I admire your commitment to innovation and data-driven decision-making."
    def vishnu(self):
        if self.role == 'Data Analyst':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in Electrical Engineering from <b>IIT Bombay (2024)</b> and currently work 
            as a <b>Data Analytics and Automation Specialist</b> at Bintix. Over the past 1.5 years, I have:</p>

            <ul>
            <li>Built dashboards in Streamlit for KPI tracking and consumer insights.</li>
            <li>Automated SQL + Python reporting pipelines, reducing manual effort by 80% and improving data accuracy by 30%.</li>
            <li>Delivered scalable analytics solutions for clients including <b>L’Oréal, HUL, and ITC</b>.</li>
            </ul>

            <p>My expertise in <b>Python, SQL, and dashboarding</b> enables me to turn complex data into actionable insights, 
            and I am eager to bring the same impact to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Vishnuvardhan Chowhan</b><br>
            Ph: 7036363267<br>
            <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
            """

            self.bullet1 = "I bring experience designing Streamlit-based KPI and Consumer Journey Dashboards, demonstrating my ability to translate complex data into actionable insights for business teams."
            self.bullet2 = "I have automated SQL + Python reporting pipelines, showing my strength in reducing manual effort and improving data accuracy—skills I can apply to optimize any data process I take on."
            self.bullet3 = "I developed an Innovations Tracker Dashboard to identify and benchmark product launches, highlighting my capability to build analytics tools that uncover market opportunities."
            self.highlights = "Python, SQL, Power BI + Streamlit dashboarding, reporting automation"
            self.cta = f"I’d love the opportunity to discuss how I can contribute to {self.company_name}’s analytics team and help drive data-driven decision making."
        elif self.role == 'Data Scientist':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I graduated from <b>IIT Bombay</b> in 2024 with a B.Tech in Electrical Engineering 
            and currently work as a <b>Data Analytics and Automation Specialist</b> at Bintix. 
            Alongside my industry experience, I completed an Advanced Machine Learning course, gaining hands-on expertise in:</p>

            <ul>
            <li>Regression, clustering, and random forests</li>
            <li>CNNs, RNNs, LSTMs, and GANs</li>
            </ul>

            <p>At <b>Bintix</b>, I applied these skills to:</p>

            <ul>
            <li><strong>Built an AI-powered chatbot for KPI extraction</strong> and developed <strong>Streamlit dashboards</strong> to visualize KPI tracking and consumer insights, translating raw datasets into clear reports.</li>
            <li>Automate insights generation from large datasets</li>
            <li>Apply ML models for trend detection and product benchmarking</li>
            </ul>

            <p>My experience in Python, SQL, Machine learning, and Analytics has enabled global clients to improve efficiency and accuracy, and I am eager to bring the same impact to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Vishnuvardhan Chowhan</b><br>
            Ph: <a href="tel:+917036363267">7036363267</a><br>
            <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
            """

            self.bullet3 = "Completed an Advanced ML course, applying models such as regression, clustering, CNNs, RNNs, LSTMs, and GANs to real-world datasets—demonstrating strong foundations in both classical and deep learning."
            self.bullet2 = "Designed an AI-powered chatbot agent that extracts KPIs and custom data cuts from natural language queries, showcasing my ability to integrate NLP with business analytics."
            self.bullet1 = "Applied ML techniques for trend identification and product benchmarking, illustrating my capacity to convert raw data into actionable insights for strategic decision-making."
            self.highlights = "AI/ML, Python, advanced analytics, natural language data agents"
            self.cta = f"I’d welcome the chance to explore how my machine learning expertise and real-world project experience can support {self.company_name}’s data science initiatives."
        elif self.role == 'Data Engineer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in Electrical Engineering from <b>IIT Bombay (2024)</b> and currently work 
            as a <b>Data Analytics and Automation Specialist</b> at <b>Bintix</b>. 
            My experience includes:</p>

            <ul>
            <li>Designing and deploying <b>Python-based ETL pipelines</b> for large datasets</li>
            <li>Automating reporting workflows, reducing manual effort by 80%</li>
            <li>Building scalable SQL-driven data processes, improving data accuracy by 30%</li>
            </ul>

            <p>I have worked with global clients such as <b>L’Oréal, HUL, and ITC</b>, ensuring data pipelines are 
            both reliable and insight-driven. I believe my expertise in <b>Python, SQL, and workflow automation</b> 
            makes me a strong fit for this opportunity.</p>

            <p><b>Best regards,</b><br>
            <b>Vishnuvardhan Chowhan</b><br>
            Ph: <a href="tel:+917036363267">7036363267</a><br>
            <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
            """

            self.bullet1 = "Designed and deployed Python-based ETL pipelines to clean, transform, and standardize millions of rows of data, demonstrating my ability to manage and scale large datasets efficiently."
            self.bullet2 = "Automated end-to-end reporting workflows, reducing manual effort by 80% and highlighting my strength in streamlining repetitive processes through automation."
            self.bullet3 = "Built robust SQL and Python scripts for data cleaning, validation, and transformation, showcasing my focus on improving pipeline reliability and ensuring data accuracy."
            self.highlights = "Python, SQL, Pandas, Big Data, ETL pipelines, workflow automation"
            self.cta = f"I’d be glad to discuss how my background in building reliable data workflows can strengthen {self.company_name}’s data infrastructure and support downstream analytics."
        elif self.role == 'Data Governance Analyst':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in Electrical Engineering from <b>IIT Bombay (2024)</b> and currently work 
            as a <b>Data Analytics and Automation Specialist</b> at Bintix. My experience sits at the 
            intersection of <b>analytics and engineering</b>, with a strong emphasis on data quality and governance. 
            Some of my contributions include:</p>

            <ul>
            <li>Designing <b>Python + SQL pipelines</b> for data cleaning, validation, and transformation, ensuring accuracy and reliability across datasets.</li>
            <li>Automating reporting workflows that enforce <b>data consistency and auditability</b> across multiple stakeholders.</li>
            <li>Developing dashboards that combine <b>data lineage tracking and KPI insights</b> for clients such as <b>L’Oréal, HUL, and ITC</b>.</li>
            </ul>

            <p>My skills in <b>data governance, process automation, and pipeline reliability</b> enable me 
            to bridge the gap between engineering and analytics, ensuring data is not only insightful but 
            also trusted. I’m excited to bring this blend of expertise to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Vishnuvardhan Chowhan</b><br>
            Ph: <a href="tel:+917036363267">7036363267</a><br>
            <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
            """

            self.bullet1 = "Built automated pipelines that validate, transform, and reconcile data across multiple sources, ensuring governance and integrity in reporting."
            self.bullet2 = "Created dashboards that integrate data lineage and business KPIs, improving both transparency and decision-making."
            self.bullet3 = "Worked with global clients where ensuring data compliance, accuracy, and auditability was critical—strengthening my governance-first mindset."
            self.highlights = "Data governance, Python, SQL, ETL, pipeline validation, data lineage, compliance"
            self.cta = f"I’d be glad to discuss how my combined experience in analytics and engineering can help strengthen {self.company_name}’s data governance and reliability frameworks."
        elif self.role == 'Product Analyst':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in Electrical Engineering from <b>IIT Bombay (2024)</b> and currently work 
            as a <b>Data Analytics and Automation Specialist</b> at Bintix. 
            In my role, I’ve worked closely with business and product teams to derive insights that drive 
            decision-making. Some examples include:</p>

            <ul>
            <li>Building <b>KPI dashboards</b> in Streamlit and Power BI to track <b>user engagement, funnel conversion, and retention</b>.</li>
            <li>Automating SQL + Python workflows to deliver real-time insights on <b>product adoption and consumer behavior</b>.</li>
            <li>Developing a <b>Product Innovations Tracker</b> to benchmark competitor launches and identify whitespace opportunities.</li>
            </ul>

            <p>With expertise in <b>SQL, Python, dashboarding, and experiment analysis</b>, I specialize in translating 
            raw data into product insights that influence growth strategies. I’m excited to bring this impact-driven 
            mindset to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Vishnuvardhan Chowhan</b><br>
            Ph: <a href="tel:+917036363267">7036363267</a><br>
            <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
            """

            self.bullet1 = "Built KPI dashboards for funnels, retention, and adoption, enabling product teams to track performance and iterate faster."
            self.bullet2 = "Automated SQL + Python pipelines for real-time consumer insights, reducing latency between data collection and decision-making."
            self.bullet3 = "Developed an Innovations Tracker Dashboard that benchmarked competitor products, providing strategic inputs to product roadmaps."
            self.highlights = "Product analytics, SQL, Python, dashboarding, funnel analysis, retention, consumer insights"
            self.cta = f"I’d welcome the opportunity to show how my data-driven approach can support {self.company_name}’s product growth and decision-making."

        self.TEMPLATE = """{today}
        Hiring Manager
        {company}

        Dear {hiring_manager},

        I’m enthusiastic about applying for the <b>{role}</b> role at <b>{company}</b>, where I see a strong alignment between my skills and the firm’s data-driven culture.

        I specialize in {highlights}, and I’ve used these skills to deliver measurable impact.

        Why {company}? {why_company}

        What I bring:

        • <b>{bullet1}</b>

        • <b>{bullet2}</b>

        • <b>{bullet3}</b>

        {cta}

        <b>Best regards,</b>  
        <b>Vishnuvardhan Chowhan</b>  
        7036363267 | ✉ vishnuvardhan.chowhan@gmail.com  
        Portfolio: https://notion-sparkle-site.lovable.app/
        """

        self.text = self.TEMPLATE.format(
            today=date.today().strftime("%B %d, %Y"),
            hiring_manager=self.recruiter or "Hiring Manager",
            role=self.role_name,
            company=self.company_name,
            highlights=self.highlights,
            why_company=self.why_company,
            bullet1=self.bullet1,
            bullet2=self.bullet2,
            bullet3=self.bullet3,
            cta=self.cta
        )
        self.text = dedent(self.text)
        self.name = 'Vishnuvardhan Chowhan'
        self.official_role = 'Data Analytics & Automation Specialist | Hyderabad, India'
        self.pdf_filename = generate_cover_letter_pdf(self.text,self.name.capitalize(), self.official_role,  f"{st.session_state.get('username')} Cover Letter.pdf")
        self.official_name = "vishnuvardhan.chowhan@gmail.com"
        self.resume_path = r"VishnuvardhanChowhanResume.pdf"
        return self.email_body, self.pdf_filename, self.official_name, self.resume_path

    def sakshi(self):
        if self.role == 'Full Stack Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your profile on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in Computer Science and currently work as a 
            <b>Full Stack Developer</b>, where I’ve gained hands-on experience in 
            designing, building, and scaling end-to-end web applications. My recent projects include:</p>

            <ul>
            <li>Developing a profiling portal using <b>React.js, Node.js, and TypeScript</b> with role-based access control.</li>
            <li>Building RESTful APIs for barcode lifecycle management, including assignment, validation, and approval workflows.</li>
            <li>Optimizing backend queries and designing scalable data-driven UIs for real-time insights.</li>
            </ul>

            <p>My expertise across <b>frontend (React.js, TypeScript)</b> and <b>backend (Node.js, APIs, SQL)</b> 
            enables me to deliver production-ready solutions, and I’m excited about applying the same to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Sakshi Gawande</b><br>
            Ph: 7057634407<br>
            """

            self.bullet1 = "Built end-to-end web applications with React.js, Node.js, and TypeScript, gaining strong expertise in both frontend and backend."
            self.bullet2 = "Designed modular APIs and scalable database interactions for workflow automation and real-time insights."
            self.bullet3 = "Implemented secure role-based access systems and optimized backend performance for large-scale data operations."
            self.highlights = "React.js, Node.js, TypeScript, REST APIs, SQL, Full Stack Architecture"
            self.cta = f"I’d love to discuss how my full-stack expertise can help {self.company_name} build scalable and user-focused applications."

        elif self.role == 'Frontend Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your profile on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I specialize in creating responsive, user-centric web applications using 
            <b>React.js, TypeScript, and modern UI frameworks</b>. My recent contributions include:</p>

            <ul>
            <li>Developing interactive dashboards with dynamic dropdowns and search features connected to live databases.</li>
            <li>Implementing reusable React components for forms, validations, and visualization modules.</li>
            <li>Optimizing page performance and enhancing user experience with clean UI/UX design.</li>
            </ul>

            <p>With my focus on <b>UI/UX, frontend performance, and modern JavaScript frameworks</b>, 
            I am excited to bring impactful user experiences to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Sakshi Gawande</b><br>
            Ph: 7057634407<br>
            """

            self.bullet1 = "Designed and built interactive dashboards in React.js and TypeScript, focusing on responsive design and performance."
            self.bullet2 = "Developed reusable UI components, form handlers, and visualization modules for scalable frontend projects."
            self.bullet3 = "Enhanced user experience with optimized rendering, smooth navigation, and accessible design."
            self.highlights = "React.js, TypeScript, JavaScript ES6+, UI/UX, Frontend Development"
            self.cta = f"I’d be excited to contribute to {self.company_name} by delivering intuitive and performance-driven frontend applications."

        elif self.role == 'Backend Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your profile on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I bring strong experience in <b>backend development with Node.js, Express, and SQL</b>. 
            At my current role, I have:</p>

            <ul>
            <li>Designed and deployed RESTful APIs for barcode lifecycle management and workflow automation.</li>
            <li>Implemented authentication and role-based access control for secure applications.</li>
            <li>Optimized database queries and built scalable data pipelines for bulk operations and analytics.</li>
            </ul>

            <p>My focus on <b>API design, database optimization, and system scalability</b> 
            makes me eager to contribute to <b>{self.company_name}</b>’s backend engineering team.</p>

            <p><b>Best regards,</b><br>
            <b>Sakshi Gawande</b><br>
            Ph: 7057634407<br>
            """

            self.bullet1 = "Developed and optimized RESTful APIs with Node.js and Express for workflow automation and data management."
            self.bullet2 = "Built secure backend services with authentication, authorization, and role-based access control."
            self.bullet3 = "Improved system scalability by optimizing SQL queries and designing efficient database workflows."
            self.highlights = "Node.js, Express, SQL, Authentication, API Development, Backend Scalability"
            self.cta = f"I’d be glad to explore how my backend expertise can strengthen {self.company_name}’s engineering team and infrastructure."

        self.TEMPLATE = """{today}
                Hiring Manager
                {company}

                Dear {hiring_manager},

                I’m excited to apply for the <b>{role}</b> role at <b>{company}</b>. With a strong foundation in 
                full-stack development, I bring proven expertise in building scalable applications and crafting 
                seamless user experiences.

                I specialize in {highlights}, and I’ve applied these skills to deliver impactful projects.

                Why {company}? {why_company}

                What I bring:

                • <b>{bullet1}</b>

                • <b>{bullet2}</b>

                • <b>{bullet3}</b>

                {cta}

                <b>Best regards,</b>  
                <b>Sakshi Gawande</b>  
                7057634407 | ✉ sakshigawandecse@gmail.com
                """

        self.text = self.TEMPLATE.format(
            today=date.today().strftime("%B %d, %Y"),
            hiring_manager=self.recruiter or "Hiring Manager",
            role=self.role_name,
            company=self.company_name,
            highlights=self.highlights,
            why_company=self.why_company,
            bullet1=self.bullet1,
            bullet2=self.bullet2,
            bullet3=self.bullet3,
            cta=self.cta
        )
        self.text = dedent(self.text)
        self.name = 'Sakshi Gawande'
        self.official_role = 'FullStack Developer | Hyderabad, India'
        self.pdf_filename = generate_cover_letter_pdf(self.text, self.name.capitalize(), self.official_role,
                                                      f"{st.session_state.get('username')} Cover Letter.pdf")
        self.official_name = "sakshigawandecse@gmail.com"
        self.resume_path = r"Sakshi_Gawande_Resume.pdf"
        return self.email_body, self.pdf_filename, self.official_name, self.resume_path

def main():
    # ---------------- STREAMLIT UI ----------------
    st.set_page_config(page_title="Automated Job Application", page_icon="📧")

    st.title("📧 Automated Job Application Email Generator")
    user = st.session_state.get("username")
    st.markdown(
        "Fill in the details below to generate a professional email and cover letter for recruiters."
    )

    st.header("1️⃣ Role & Job Details")

    if user == 'vishnu':
        role = st.selectbox("Select the Role", ['Data Analyst', 'Data Scientist', 'Data Engineer', 'Data Governance Analyst', 'Product Analyst'])
    elif user == 'sakshi':
        role = st.selectbox("Select the Role", ['Full Stack Developer', 'Frontend Developer', 'Backend Developer'])
    role_name = st.text_input("Official Role Name (as per Job Posting)", placeholder="Type here...")
    job_id = st.text_input("Job ID / Reference Number", placeholder="Type here...")

    st.header("2️⃣ Recruiter & Company Info")
    recruiter_mail = st.text_input("Recruiter's Email(s)", placeholder="e.g., adarsh@company.com, rina@company.com")
    recipient_list = [email.strip() for email in recruiter_mail.split(",") if email.strip()]
    company_name = st.text_input("Company Name", placeholder="Type here...")

    st.header("3️⃣ Motivation & Customization")
    catchy_subject = st.text_input(
        "Write catchy subject to attract recruiters ✨",
        placeholder="Write witty subject..."
    )
    why_company = st.text_input(
        "Why do you want to join this company?",
        placeholder="Write 1–2 lines about your motivation..."
    )
    job_description = st.text_input(
        "Job description link for future follow ups?",
        placeholder="Job description link..."
    )

    st.markdown("---")
    st.caption("⚠️ Make sure all fields are filled before generating the email.")

    if st.button("Send"):
        pass_dict = {'sakshigawandecse@gmail.com':"illr ufri rqeo cwia", "vishnuvardhan.chowhan@gmail.com": "dipi cqsq sgvz ukof"}
        if not recruiter_mail or not company_name or not role_name:
            st.warning("⚠️ Please fill in all the fields before sending.")
        else:
            st.success("✅ All fields are filled. Ready to send!")
        # ---------------- EMAIL ----------------
        for recipient in recipient_list:
            email_username = recipient.split("@")[0]
            recruiter_name = email_username.replace(".", " ").title()
            email = EmailSender(
                role=role,
                recruiter=recruiter_name,
                role_name=role_name,
                company_name=company_name,
                why_company=why_company)
            method = getattr(email, user)
            email_body, pdf_filename, official_mail, resume_path = method()
            msg = MIMEMultipart()
            msg["From"] = official_mail
            msg["To"] = recipient
            if catchy_subject:
                msg["Subject"] = catchy_subject
            elif not job_id:
                msg["Subject"] = f"{role_name} Role Application – Resume & Cover Letter"
            else:
                msg["Subject"] = f"{role_name} Application [{job_id}] – Resume & Cover Letter"
            body = MIMEText(email_body, "html")
            msg.attach(body)
            if os.path.exists(resume_path):
                with open(resume_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(resume_path)}")
                msg.attach(part)

            # Attach Cover Letter PDF
            if os.path.exists(pdf_filename):
                with open(pdf_filename, "rb") as attachment:
                    part = MIMEBase("application", "pdf")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(pdf_filename)}")
                msg.attach(part)

            # Send email
            smtp = smtplib.SMTP("smtp.gmail.com", 587)
            smtp.starttls()
            smtp.login(official_mail, pass_dict[official_mail])
            smtp.sendmail(msg["From"], recipient, msg.as_string())
            smtp.quit()

            st.success(f"📧 Email sent successfully to {recruiter_name} with Resume + Cover Letter PDF!")
            service = authenticate_google_sheets()
            if not service:
                return
            SPREADSHEET_ID = "1bsD_uv_r1uNWn9JD85WWMpwTnxEmuP-Eqm-zlI2tp9U"
            # log_application(f"../{user}.xlsx", [
            #     date.today().strftime("%Y-%m-%d"),
            #     company_name,
            #     role_name,
            #     job_id,
            #     recruiter_name,
            #     recruiter_mail,
            #     msg["Subject"],
            #     why_company,
            #     job_description,
            #     "Sent"
            # ])
            application_details = [date.today().strftime("%Y-%m-%d"),company_name,role_name,job_id,recruiter_name,recruiter_mail,msg["Subject"],why_company,job_description,"Sent"]
            log_application(service, SPREADSHEET_ID, user, application_details)
            st.info(f"✅ Application logged in {user}.xlsx")
if __name__ == "__main__":
    main()
