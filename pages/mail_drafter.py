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
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
    ListFlowable
)
import os
import base64
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
load_dotenv()
creds_base64 = os.getenv('GOOGLE_CREDENTIALS_BASE64')


def authenticate_google_sheets():
    """Authenticate with Google Sheets API using service account from Base64 env variable."""
    creds_base64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_BASE64")

    if not creds_base64:
        st.write("No service account credentials found. Please set GOOGLE_SERVICE_ACCOUNT_BASE64 in your environment.")
        return None

    try:
        creds_json = base64.b64decode(creds_base64).decode("utf-8")
        creds_info = json.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        service = build("sheets", "v4", credentials=creds)
        return service

    except Exception as e:
        st.write(f"Error loading service account credentials: {e}")
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
            Ph: <a href="tel:+917036363267">7036363267</a><br>
            <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
            """

            self.bullet1 = "Designed Streamlit KPI & Consumer Dashboards to deliver actionable insights."
            self.bullet2 = "Automated SQL + Python pipelines, reducing manual effort and improving accuracy."
            self.bullet3 = "Built an Innovations Tracker Dashboard to benchmark launches and spot opportunities."
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

            self.bullet1 = "Applied ML for trend analysis and product benchmarking, turning raw data into strategic insights."
            self.bullet2 = "Built an AI-powered chatbot to extract KPIs from natural language, integrating NLP with analytics."
            self.bullet3 = "Completed an Advanced ML course, applying regression, clustering, CNNs, RNNs, LSTMs, and GANs to real-world datasets."
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

            self.bullet1 = "Scaled data management by designing Python ETL pipelines that cleaned and standardized millions of rows."
            self.bullet2 = "Reduced manual effort by 80% through automation of end-to-end reporting workflows."
            self.bullet3 = "Improved pipeline reliability and accuracy by building SQL + Python scripts for data cleaning and validation."
            self.highlights = "Python, SQL, Pandas, Big Data, ETL pipelines, workflow automation"
            self.cta = f"I’d be glad to discuss how my background in building reliable data workflows can strengthen {self.company_name}’s data infrastructure and support downstream analytics."
        elif self.role == 'Machine Learning Engineer':
            self.email_body = f"""
                    <p>Hi {self.recruiter},</p>

                    <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
                    I came across your contact information on LinkedIn and wanted to reach out directly. 
                    Thank you for taking the time to consider my application.</p>

                    <p>I hold a B.Tech in Electrical Engineering from <b>IIT Bombay (2024)</b> and currently work as a 
                    <b>Data Analytics and Automation Specialist</b> at <b>Bintix</b>. My role bridges data engineering and applied AI, where I have:</p>

                    <ul>
                    <li>Designed and deployed <b>Python-based ETL pipelines</b> to clean, transform, and scale datasets for downstream analytics</li>
                    <li>Developed the <b>Graahax AI agent</b>, an NLP-powered assistant that extracts KPIs and data cuts using prompt engineering and fuzzy matching</li>
                    <li>Automated reporting workflows and integrated AI outputs into <b>Streamlit dashboards</b>, reducing manual effort by 80%</li>
                    </ul>

                    <p>At <b>Bintix</b>, I have successfully combined <b>AI integration</b> with <b>data pipeline engineering</b> 
                    to deliver scalable, insight-driven solutions for global clients such as <b>L’Oréal, HUL, and ITC</b>. 
                    I believe this unique blend of skills makes me a strong fit for <b>{self.company_name}</b>’s data initiatives.</p>

                    <p><b>Best regards,</b><br>
                    <b>Vishnuvardhan Chowhan</b><br>
                    Ph: <a href="tel:+917036363267">7036363267</a><br>
                    <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
                    """

            self.bullet1 = "Built the Graahax AI agent using LLMs, prompt engineering, and fuzzy matching to turn natural language queries into structured KPI insights."
            self.bullet2 = "Developed scalable Python ETL pipelines for ingestion, cleaning, and transformation of large datasets, ensuring accuracy and reliability."
            self.bullet3 = "Automated reporting workflows and integrated AI-driven insights into Streamlit dashboards, cutting manual analyst effort by 80%."
            self.highlights = "AI agents, Prompt Engineering, Python, SQL, ETL pipelines, workflow automation"
            self.cta = f"I’d welcome the chance to discuss how my expertise in AI-powered agents and data engineering can strengthen {self.company_name}’s data science and analytics efforts."
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

            self.bullet1 = "Built automated pipelines to validate, transform, and reconcile multi-source data, ensuring governance and reporting integrity."
            self.bullet2 = "Developed dashboards combining data lineage with business KPIs, enhancing transparency and decision-making."
            self.bullet3 = "Partnered with global clients on compliance-critical projects, strengthening my focus on accuracy, auditability, and governance."
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
                In my role, I collaborate with product and business stakeholders to provide insights 
                that guide decision-making and feature prioritization. Some examples include:</p>

                <ul>
                <li>Building <b>KPI dashboards</b> in Streamlit and Power BI that track brand performance, consumer journeys, and adoption patterns.</li>
                <li>Automating SQL + Python workflows to generate real-time reporting on <b>user behavior and market trends</b>, reducing manual effort by 80%.</li>
                <li>Developing tools like an <b>Innovations Tracker</b> that benchmarked new product launches and highlighted whitespace opportunities for growth.</li>
                </ul>

                <p>With expertise in <b>SQL, Python, dashboarding, and analytics automation</b>, 
                I specialize in transforming raw datasets into insights that inform <b>product strategy and growth decisions</b>. 
                I’m excited to bring this impact-driven mindset to <b>{self.company_name}</b>.</p>

                <p><b>Best regards,</b><br>
                <b>Vishnuvardhan Chowhan</b><br>
                Ph: <a href="tel:+917036363267">7036363267</a><br>
                <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
                """

            self.bullet1 = "Designed KPI dashboards to track adoption, journeys, and performance, giving product teams clarity on growth levers."
            self.bullet2 = "Automated SQL + Python pipelines for reporting, enabling faster, more reliable insights on user and market behavior."
            self.bullet3 = "Built an Innovations Tracker Dashboard to benchmark competitor launches and spot market whitespace, shaping product strategy."
            self.highlights = "Product-focused analytics, SQL, Python, dashboarding, consumer journeys, growth insights"
            self.cta = f"I’d love the opportunity to discuss how my analytics background can support {self.company_name}’s product growth and help shape data-informed decisions."

        elif self.role == 'Python Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in Electrical Engineering from <b>IIT Bombay (2024)</b> and currently work 
            as a <b>Data Analytics and Automation Specialist</b> at Bintix. 
            My work revolves around designing scalable Python solutions and automations. Some highlights include:</p>

            <ul>
            <li>Developing <b>Python-based ETL pipelines</b> for ingestion, cleaning, and transformation of millions of rows of data.</li>
            <li>Building and deploying <b>automation scripts and APIs</b> to streamline reporting and integration with client systems.</li>
            <li>Implementing <b>AI-powered agents</b> and integrating them into dashboards using FastAPI and Streamlit.</li>
            </ul>

            <p>My expertise in <b>Python, SQL, APIs, and workflow automation</b> has enabled me to 
            deliver production-ready, scalable solutions for global clients such as <b>L’Oréal, HUL, and ITC</b>. 
            I’m excited about the opportunity to bring the same impact to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Vishnuvardhan Chowhan</b><br>
            Ph: <a href="tel:+917036363267">7036363267</a><br>
            <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
            """

            self.bullet1 = "Built Python ETL pipelines to clean, transform, and load large datasets for reliable workflows."
            self.bullet2 = "Developed automation scripts and APIs, streamlining reporting and system integrations."
            self.bullet3 = "Integrated AI-powered agents into dashboards with FastAPI and Streamlit for interactive analytics."
            self.highlights = "Python, APIs, ETL pipelines, FastAPI, workflow automation, data integration"
            self.cta = f"I’d welcome the chance to discuss how my Python expertise and automation background can strengthen {self.company_name}’s engineering efforts."

        self.TEMPLATE = """{today}
        Hiring Manager
        {company}

        Dear {hiring_manager},

        I’m enthusiastic about applying for the <b>{role}</b> role at <b>{company}</b>, where I see a strong alignment between my skills and the firm’s data-driven culture.

        I specialize in {highlights}, and I’ve used these skills to deliver measurable impact.

        Why {company}?
        
        {why_company}

        My experience has equipped me with skills directly relevant to this role:  
        
        • <b>{bullet1}</b>
        
        • <b>{bullet2}</b>
        
        • <b>{bullet3}</b>

        {cta}

        <b>Best regards,</b>
        <b>Vishnuvardhan Chowhan</b>  
        7036363267 | ✉ vishnuvardhan.chowhan@gmail.com 
        <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
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
            Ph: <a href="tel:+917057634407">7057634407</a><br>
            <a href="linkedin.com/in/sakshi-gawande-0095351ab">LinkedIn</a></p>
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
            Ph: <a href="tel:+917057634407">7057634407</a><br>
            <a href="linkedin.com/in/sakshi-gawande-0095351ab">LinkedIn</a></p>
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
            Ph: <a href="tel:+917057634407">7057634407</a><br>
            <a href="linkedin.com/in/sakshi-gawande-0095351ab">LinkedIn</a></p>
            """

            self.bullet1 = "Developed and optimized RESTful APIs with Node.js and Express for workflow automation and data management."
            self.bullet2 = "Built secure backend services with authentication, authorization, and role-based access control."
            self.bullet3 = "Improved system scalability by optimizing SQL queries and designing efficient database workflows."
            self.highlights = "Node.js, Express, SQL, Authentication, API Development, Backend Scalability"
            self.cta = f"I’d be glad to explore how my backend expertise can strengthen {self.company_name}’s engineering team and infrastructure."

        elif self.role == 'Software Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your profile on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>As a <b>Software Developer</b>, I bring experience in building reliable, scalable, and user-focused applications. 
            My work spans across <b>frontend, backend, and database systems</b>, allowing me to contribute to every layer of software development. 
            Key contributions include:</p>

            <ul>
            <li>Developed end-to-end web applications using <b>React.js, Node.js, and TypeScript</b> with modular, reusable components.</li>
            <li>Built RESTful APIs and optimized backend queries to support high-volume data workflows.</li>
            <li>Implemented secure role-based access and designed scalable databases for performance-driven applications.</li>
            </ul>

            <p>My expertise in <b>full-stack development, software architecture, and performance optimization</b> 
            equips me to deliver high-quality solutions, and I am eager to bring the same impact to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Sakshi Gawande</b><br>
            Ph: <a href="tel:+917057634407">7057634407</a><br>
            <a href="linkedin.com/in/sakshi-gawande-0095351ab">LinkedIn</a></p>
            """

            self.bullet1 = "Developed full-stack web applications with React.js, Node.js, and TypeScript, ensuring scalable and maintainable codebases."
            self.bullet2 = "Designed and deployed RESTful APIs, focusing on backend performance, data workflows, and integration reliability."
            self.bullet3 = "Implemented secure role-based access systems and optimized database queries to enhance application performance."
            self.highlights = "Full-Stack Development, React.js, Node.js, TypeScript, REST APIs, Software Architecture"
            self.cta = f"I’d be excited to discuss how my software development expertise can support {self.company_name} in building scalable and user-friendly applications."
        elif self.role == 'Process Associate':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> position at <b>{self.company_name}</b>. 
            I came across your profile on LinkedIn and wanted to connect directly. 
            Thank you for considering my application.</p>

            <p>I have experience in <b>data processing, workflow optimization, and process documentation</b> 
            to ensure smooth and accurate business operations. My recent contributions include:</p>

            <ul>
            <li>Processing large data sets and ensuring 100% accuracy through verification and validation steps.</li>
            <li>Maintaining detailed documentation for audits, performance tracking, and process improvements.</li>
            <li>Collaborating with cross-functional teams to resolve operational bottlenecks and streamline workflows.</li>
            </ul>

            <p>With my attention to detail, <b>process discipline, and commitment to efficiency</b>, 
            I look forward to supporting <b>{self.company_name}</b> in delivering seamless business operations.</p>

            <p><b>Best regards,</b><br>
            <b>Sakshi Gawande</b><br>
            Ph: <a href="tel:+917057634407">7057634407</a><br>
            <a href="linkedin.com/in/sakshi-gawande-0095351ab">LinkedIn</a></p>
            """

            self.bullet1 = "Processed and verified large datasets with a focus on accuracy and consistency across workflows."
            self.bullet2 = "Created and maintained process documentation for audits and operational transparency."
            self.bullet3 = "Collaborated with internal teams to identify and resolve process inefficiencies."
            self.highlights = "Data Processing, Process Optimization, Documentation, MS Excel, Workflow Management"
            self.cta = f"I’m eager to contribute to {self.company_name} by ensuring efficient and accurate execution of core business processes."

        if self.role in ['Full Stack Developer', 'Frontend Developer', 'Backend Developer', 'Software Developer']:
            self.TEMPLATE = """{today}
                    Hiring Manager
                    {company}
    
                    Dear {hiring_manager},
    
                    I’m excited to apply for the <b>{role}</b> role at <b>{company}</b>. With a strong foundation in 
                    full-stack development, I bring proven expertise in building scalable applications and crafting 
                    seamless user experiences.
    
                    I specialize in {highlights}, and I’ve applied these skills to deliver impactful projects.
    
                    Why {company}? {why_company}
    
                    My experience has equipped me with skills directly relevant to this role:  
                    
                    • <b>{bullet1}</b>
                    
                    • <b>{bullet2}</b>
                    
                    • <b>{bullet3}</b>
    
                    {cta}
    
                    <b>Best regards,</b>  
                    <b>Sakshi Gawande</b>  
                    7057634407 | ✉ sakshigawandecse@gmail.com
                    <a href="linkedin.com/in/sakshi-gawande-0095351ab">LinkedIn</a></p>
                    """
        else:
            self.TEMPLATE = """{today}
                            Hiring Manager
                            {company}
    
                            Dear {hiring_manager},
    
                            I’m excited to apply for the <b>{role}</b> position at <b>{company}</b>. With a strong foundation in 
                            operations, coordination, and communication, I bring a proven ability to manage processes efficiently 
                            and contribute to team success.
    
                            I specialize in {highlights}, and I’ve consistently applied these skills to improve workflows and support 
                            organizational goals.
    
                            Why {company}? {why_company}
    
                            My experience has equipped me with skills directly relevant to this role:  
    
                            • <b>{bullet1}</b>  
                            • <b>{bullet2}</b>  
                            • <b>{bullet3}</b>
    
                            {cta}
    
                            <b>Best regards,</b>  
                            <b>Sakshi Gawande</b>  
                            7057634407 | ✉ sakshigawandecse@gmail.com  
                            <a href="linkedin.com/in/sakshi-gawande-0095351ab">LinkedIn</a>
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
        if self.role in ['Full Stack Developer', 'Frontend Developer', 'Backend Developer', 'Software Developer']:
            self.resume_path = r"Sakshi_Gawande_Resume_tech.pdf"
        else:
            self.resume_path = r"Sakshi_Gawande_Resume_non_tech.pdf"
        return self.email_body, self.pdf_filename, self.official_name, self.resume_path

    def sai(self):
        if self.role == 'Full Stack Engineer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your profile on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in <b>Computer Science and Engineering (2024)</b> from <b>Adikavi Nannaya University</b> 
            and currently work as a <b>Full Stack Engineer & App Developer</b> at <b>Bintix (T-Hub, Hyderabad)</b>. 
            Over the past 1.5 years, I have:</p>

            <ul>
            <li>Developed and deployed Android & iOS apps with AI-based features such as image blur detection, Bluetooth weighing, and real-time barcode validation—used across 7 metro cities.</li>
            <li>Built dynamic web applications with <b>React</b> and <b>Material UI</b>, integrating <b>Highcharts</b> for interactive data visualizations and analytics dashboards.</li>
            <li>Developed and deployed scalable web and backend applications — including a Node.js REST API project, a responsive restaurant website, and a serverless Toy Store app using <b>Google Cloud (Cloud Run, Firebase, Vertex AI)</b> for intelligent, cloud-native product experiences.</li>
            </ul>

            <p>My expertise in <b>React Native, React, Node.js, and Python</b> allows me to develop full-stack solutions 
            that are both high-performing and user-centric. I’m excited about the opportunity to bring 
            this experience to <b>{self.company_name}</b> and contribute to impactful products.</p>

            <p><b>Best regards,</b><br>
            <b>Polloju Sai Kiran</b><br>
            Ph: 7093263001<br>
            ✉ <a href="mailto:pollojukiran06@gmail.com">pollojukiran06@gmail.com</a><br>
            <a href="https://www.linkedin.com/in/polloju-sai-kiran/">LinkedIn</a><br>
            <a href="https://my-digital-creations.lovable.app/">Portfolio</a></p>
            Hyderabad, Telangana</p>
            """

            self.bullet1 = "I’ve built and deployed production-ready mobile applications integrating AI-based features, demonstrating my ability to combine innovation with performance optimization."
            self.bullet2 = "I developed complex React + Material UI web apps with Highcharts visualizations, showcasing my strength in designing engaging, data-driven user experiences."
            self.bullet3 = "I developed scalable web and backend projects—including a Node.js REST API (TwitterClone), a responsive restaurant website, and a serverless Toy Store App using Google Cloud (Cloud Run, Firebase, Vertex AI)—demonstrating my expertise in building intelligent, cloud-native, and open-source integrated solutions."
            self.highlights = "React Native, React, Node.js, TypeScript, SQLite, Material UI, Python"
            self.cta = f"I’d love the opportunity to discuss how I can contribute to {self.company_name}’s engineering team and help build scalable, impactful digital products."

        elif self.role == 'Android Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your profile on LinkedIn and wanted to reach out directly.</p>

            <p>I hold a B.Tech in <b>Computer Science and Engineering (2024)</b> from <b>Adikavi Nannaya University</b> 
            and currently work as an <b>Android Developer</b> at <b>Bintix (T-Hub, Hyderabad)</b>. 
            Over the past 1.5 years, I have:</p>

            <ul>
            <li>Built and deployed Android applications with AI features like image blur detection, Bluetooth weighing, and real-time barcode validation.</li>
            <li>Worked with <b>Kotlin</b> and <b>Jetpack Compose</b> to deliver responsive and performant mobile apps.</li>
            <li>Integrated apps with backend services using <b>Firebase</b> and <b>REST APIs</b> for seamless cloud-native experiences.</li>
            </ul>

            <p>My expertise in <b>Kotlin, Java, Android SDK, and Firebase</b> enables me to create high-performing, user-centric applications. 
            I’m excited about the opportunity to bring this experience to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Polloju Sai Kiran</b><br>
            Ph: 7093263001<br>
            ✉ <a href="mailto:pollojukiran06@gmail.com">pollojukiran06@gmail.com</a><br>
            <a href="https://www.linkedin.com/in/polloju-sai-kiran/">LinkedIn</a><br>
            <a href="https://my-digital-creations.lovable.app/">Portfolio</a></p>
            Hyderabad, Telangana</p>
            """

            self.bullet1 = "Built and deployed Android applications integrating AI features, ensuring high performance and user engagement."
            self.bullet2 = "Proficient in Kotlin, Java, and Jetpack Compose for creating dynamic, responsive applications."
            self.bullet3 = "Integrated apps with Firebase and REST APIs for scalable cloud-native experiences."
            self.highlights = "Kotlin, Java, Android SDK, Firebase, Jetpack Compose, REST APIs"
            self.cta = f"I’d love to discuss how I can contribute to {self.company_name}’s Android development team and deliver impactful mobile solutions."

        elif self.role == 'Frontend Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’m interested in the <b>{self.role_name}</b> role at <b>{self.company_name}</b> and wanted to reach out directly.</p>

            <p>I hold a B.Tech in <b>Computer Science and Engineering (2024)</b> from <b>Adikavi Nannaya University</b> 
            and currently work as a <b>Frontend Developer</b> at <b>Bintix (T-Hub, Hyderabad)</b>. 
            Over the past 1.5 years, I have:</p>

            <ul>
            <li>Developed responsive web applications using <b>React</b> and <b>Material UI</b>.</li>
            <li>Integrated <b>Highcharts</b> and other visualization tools for interactive dashboards.</li>
            <li>Optimized UI performance, accessibility, and cross-browser compatibility for multiple client projects.</li>
            </ul>

            <p>My expertise in <b>React, JavaScript, Material UI</b> allows me to build dynamic, user-focused interfaces. 
            I’m eager to contribute to <b>{self.company_name}</b>’s frontend development efforts.</p>

            <p><b>Best regards,</b><br>
            <b>Polloju Sai Kiran</b><br>
            Ph: 7093263001<br>
            ✉ <a href="mailto:pollojukiran06@gmail.com">pollojukiran06@gmail.com</a><br>
            <a href="https://www.linkedin.com/in/polloju-sai-kiran/">LinkedIn</a><br>
            <a href="https://my-digital-creations.lovable.app/">Portfolio</a></p>
            Hyderabad, Telangana</p>
            """

            self.bullet1 = "Developed responsive and dynamic web applications with React and Material UI."
            self.bullet2 = "Integrated interactive charts and dashboards using Highcharts for better data insights."
            self.bullet3 = "Optimized frontend performance, accessibility, and cross-browser compatibility."
            self.highlights = "React, JavaScript, Material UI, HTML, CSS, Highcharts"
            self.cta = f"I’d love to discuss how I can contribute to {self.company_name}’s frontend development team and build intuitive user interfaces."

        elif self.role == 'Mobile Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’m interested in the <b>{self.role_name}</b> role at <b>{self.company_name}</b> and wanted to reach out directly.</p>

            <p>I hold a B.Tech in <b>Computer Science and Engineering (2024)</b> from <b>Adikavi Nannaya University</b> 
            and currently work as a <b>Mobile Developer</b> at <b>Bintix (T-Hub, Hyderabad)</b>. 
            Over the past 1.5 years, I have:</p>

            <ul>
            <li>Built and deployed Android & iOS applications with AI-powered features such as image blur detection and real-time barcode validation.</li>
            <li>Worked with <b>React Native</b>, <b>Flutter</b>, and native SDKs to create responsive mobile apps.</li>
            <li>Integrated mobile apps with cloud services like <b>Firebase</b> and <b>Google Cloud</b> for scalable deployments.</li>
            </ul>

            <p>My expertise in <b>React Native, Flutter, Android, iOS, and Firebase</b> allows me to develop efficient and high-quality mobile solutions. 
            I’m excited to bring this experience to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Polloju Sai Kiran</b><br>
            Ph: 7093263001<br>
            ✉ <a href="mailto:pollojukiran06@gmail.com">pollojukiran06@gmail.com</a><br>
            <a href="https://www.linkedin.com/in/polloju-sai-kiran/">LinkedIn</a><br>
            <a href="https://my-digital-creations.lovable.app/">Portfolio</a></p>
            Hyderabad, Telangana</p>
            """

            self.bullet1 = "Built and deployed mobile applications for Android and iOS integrating AI features."
            self.bullet2 = "Proficient in React Native, Flutter, and native mobile SDKs for creating responsive apps."
            self.bullet3 = "Integrated cloud services like Firebase and Google Cloud for scalable mobile deployments."
            self.highlights = "React Native, Flutter, Android SDK, iOS SDK, Firebase, Google Cloud"
            self.cta = f"I’d love to discuss how I can contribute to {self.company_name}’s mobile development team and deliver impactful apps."

        elif self.role == 'Software Developer' or self.role == 'Software Engineer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your profile on LinkedIn and wanted to reach out directly.</p>

            <p>I hold a B.Tech in <b>Computer Science and Engineering (2024)</b> from <b>Adikavi Nannaya University</b> 
            and currently work as a <b>Software Developer</b> at <b>Bintix (T-Hub, Hyderabad)</b>. 
            Over the past 1.5 years, I have:</p>

            <ul>
            <li>Developed and deployed full-stack applications with <b>React</b>, <b>Node.js</b>, and <b>Python</b>.</li>
            <li>Worked on scalable backend services and REST APIs for multiple client projects.</li>
            <li>Created automated solutions and cloud-integrated applications using <b>Google Cloud</b> and <b>Firebase</b>.</li>
            </ul>

            <p>My expertise in <b>React, Node.js, Python, and cloud technologies</b> enables me to build efficient, scalable software solutions. 
            I’m eager to contribute to <b>{self.company_name}</b>’s engineering initiatives.</p>

            <p><b>Best regards,</b><br>
            <b>Polloju Sai Kiran</b><br>
            Ph: 7093263001<br>
            ✉ <a href="mailto:pollojukiran06@gmail.com">pollojukiran06@gmail.com</a><br>
            <a href="https://www.linkedin.com/in/polloju-sai-kiran/">LinkedIn</a><br>
            <a href="https://my-digital-creations.lovable.app/">Portfolio</a></p>
            Hyderabad, Telangana</p>
            """

            self.bullet1 = "Built full-stack applications with React, Node.js, and Python, demonstrating end-to-end development capabilities."
            self.bullet2 = "Developed scalable backend services and REST APIs for multiple projects."
            self.bullet3 = "Integrated cloud solutions using Google Cloud and Firebase for efficient deployments."
            self.highlights = "React, Node.js, Python, Firebase, Google Cloud, REST APIs"
            self.cta = f"I’d love the opportunity to discuss how I can contribute to {self.company_name}’s software engineering team and deliver scalable solutions."

        self.TEMPLATE = """{today}
                Hiring Manager
                {company}

                Dear {hiring_manager},

                I’m enthusiastic about applying for the <b>{role}</b> role at <b>{company}</b>, where I see a strong alignment between my full-stack development experience and your team’s product vision.

                I specialize in {highlights}, and I’ve used these skills to deliver robust applications and data-driven interfaces.

                Why {company}? {why_company}

                What I bring:

                • <b>{bullet1}</b>

                • <b>{bullet2}</b>

                • <b>{bullet3}</b>

                {cta}

                <b>Best regards,</b>  
                <b>Polloju Sai Kiran</b>  
                7093263001 | ✉ pollojukiran06@gmail.com  
                Hyderabad, Telangana"""
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
        self.name = 'Polloju Sai Kiran'
        self.official_role = 'Full Stack Engineer | Hyderabad, India'
        self.pdf_filename = generate_cover_letter_pdf(self.text, self.name.capitalize(), self.official_role,
                                                      f"{st.session_state.get('username')} Cover Letter.pdf")
        self.official_name = "pollojukiran06@gmail.com"
        self.resume_path = r"polloju_SaiKiran_Resume.pdf"
        return self.email_body, self.pdf_filename, self.official_name, self.resume_path

    def harsha(self):
        if self.role == 'Data Analyst':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold an <b>MBA in Healthcare and Hospital Management</b> from <b>Hyderabad Central University</b> 
            and a <b>B.Sc. in Life Sciences</b> from the University of Delhi. 
            Over the past 3+ years, I’ve built expertise in market research, data analysis, and project management, working across sectors like FMCG, healthcare, and public policy. 
            I currently work as a <b>Data Analyst at Bintix Waste Research</b>, where I support consumer insights and decision-making for leading clients.</p>

            <ul>
            <li>Analyzed consumer and market data in the FMCG domain to uncover actionable insights for strategic marketing decisions.</li>
            <li>Designed and delivered teaser and performance reports that improved engagement and informed business strategies.</li>
            <li>Led end-to-end project delivery in previous roles — managing research, financial analysis, and stakeholder communication to drive impact.</li>
            </ul>

            <p>My expertise in <b>Advanced Excel, Power BI, MySQL (Basics)</b>, and <b>data-driven storytelling</b> enables me to transform complex data into clear insights — 
            a skill I’m excited to bring to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Harsha Jha</b><br>
            Ph: <a href="tel:+917836907197">7836907197</a><br>
            """

            self.bullet1 = "I bring hands-on experience in analyzing consumer and market data to generate actionable insights for strategic decision-making."
            self.bullet2 = "I have a strong foundation in reporting and dashboarding using Advanced Excel and Power BI, enabling data-driven storytelling for business stakeholders."
            self.bullet3 = "With experience in market research, financial analysis, and project delivery, I’m adept at managing end-to-end data initiatives that align with business goals."
            self.highlights = "Advanced Excel, Power BI, MySQL (Basics), Reporting & Analysis, Market Research"
            self.cta = f"I’d love the opportunity to discuss how my analytical skills and research experience can contribute to {self.company_name}’s data analytics team."
        elif self.role == 'Market Researcher':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold an <b>MBA in Healthcare and Hospital Management</b> from <b>Hyderabad Central University</b> 
            and a <b>B.Sc. in Life Sciences</b> from the University of Delhi. 
            Over the past 3+ years, I’ve developed expertise in consumer behavior, market research methodologies, and data-driven insights across FMCG, healthcare, and public policy sectors. 
            I currently work as a <b>Market Researcher at Bintix Waste Research</b>, helping clients make strategic decisions through research and insights.</p>

            <ul>
            <li>Conducted qualitative and quantitative research to analyze market trends, consumer behavior, and competitive landscapes.</li>
            <li>Developed reports and presentations that translated complex data into actionable business insights.</li>
            <li>Collaborated with cross-functional teams to deliver research projects on time and ensure stakeholder satisfaction.</li>
            </ul>

            <p>My expertise in <b>survey design, data collection, SPSS, Excel</b>, and <b>market analysis</b> enables me to provide precise and actionable insights — 
            a skill I’m eager to bring to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Harsha Jha</b><br>
            Ph: <a href="tel:+917836907197">7836907197</a><br>
            """

            self.bullet1 = "I have hands-on experience conducting qualitative and quantitative research to uncover consumer behavior and market trends."
            self.bullet2 = "Skilled in designing surveys, collecting data, and using SPSS and Excel for actionable insights."
            self.bullet3 = "Experienced in delivering market research projects end-to-end, ensuring accurate insights for strategic decision-making."
            self.highlights = "Market Research, Consumer Insights, SPSS, Excel, Survey Design, Data Analysis"
            self.cta = f"I’d love the opportunity to discuss how my market research expertise can help {self.company_name} make informed, data-driven decisions."
        elif self.role == 'Project Manager':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’d like to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact information on LinkedIn and wanted to reach out directly. 
            Thank you for taking the time to consider my application.</p>

            <p>I hold an <b>MBA in Healthcare and Hospital Management</b> from <b>Hyderabad Central University</b> 
            and a <b>B.Sc. in Life Sciences</b> from the University of Delhi. 
            Over the past 3+ years, I’ve led multiple projects across FMCG, healthcare, and public policy sectors, managing timelines, resources, and stakeholders efficiently. 
            I currently work as a <b>Project Manager at Bintix Waste Research</b>, delivering strategic initiatives that drive organizational impact.</p>

            <ul>
            <li>Planned, executed, and monitored projects from inception to completion, ensuring timely delivery and high-quality outcomes.</li>
            <li>Managed cross-functional teams, coordinated stakeholders, and streamlined workflows for project efficiency.</li>
            <li>Developed reports and dashboards to track progress, risks, and KPIs, facilitating data-driven decision-making.</li>
            </ul>

            <p>My expertise in <b>project planning, stakeholder management, MS Project, Excel, and reporting</b> equips me to deliver impactful projects — 
            and I’m excited to bring this to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Harsha Jha</b><br>
            Ph: <a href="tel:+917836907197">7836907197</a><br>
            """

            self.bullet1 = "Proven experience in planning, executing, and monitoring projects across diverse sectors."
            self.bullet2 = "Strong skills in stakeholder management and coordinating cross-functional teams to achieve project goals."
            self.bullet3 = "Proficient in project tracking, reporting, and using tools like MS Project and Excel for data-driven decisions."
            self.highlights = "Project Management, Stakeholder Management, MS Project, Excel, Reporting, Workflow Optimization"
            self.cta = f"I’d love the opportunity to discuss how my project management skills can drive successful initiatives at {self.company_name}."

        self.TEMPLATE = """{today}
        Hiring Manager
        {company}

        Dear {hiring_manager},

        I’m enthusiastic about applying for the <b>{role}</b> role at <b>{company}</b>. With over 4+ years of experience in 
        market research, data analysis, and project management, I bring a strong analytical mindset and a proven ability 
        to turn data into meaningful business insights.

        I specialize in {highlights}, and I’ve applied these skills to support decision-making in sectors like FMCG, 
        healthcare, and public policy.

        Why {company}? {why_company}

        What I bring:

        • <b>{bullet1}</b>

        • <b>{bullet2}</b>

        • <b>{bullet3}</b>

        {cta}

        <b>Best regards,</b>  
        <b>Harsha Jha</b>  
        📞+91-7836907197 | ✉ harshajha13@gmail.com
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
        self.name = 'Harsha Jha'
        self.official_role = 'Data Analyst'
        self.pdf_filename = generate_cover_letter_pdf(self.text, self.name.capitalize(), self.official_role,f"{st.session_state.get('username')} Cover Letter.pdf")
        self.official_name = "harshajha13@gmail.com"
        self.resume_path = r"Harsha Jha Resume.pdf"
        return self.email_body, self.pdf_filename, self.official_name, self.resume_path

    def bhanu(self):
        if self.role == 'Full Stack Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’m writing to express my keen interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact on LinkedIn and wanted to reach out directly regarding potential opportunities 
            in your engineering team. Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in Mechanical Engineering from <b>IIT Patna (2024)</b> and currently work as a 
            <b>Full Stack Developer</b> at <b>Bintix</b>. Over the past year, I’ve contributed to multiple end-to-end 
            web applications, focusing on scalable backend logic, responsive front-ends, and seamless data integration.</p>

            <ul>
            <li>Built and maintained web applications using <b>React.js, Node.js, and Express</b> with MongoDB and Sequelize ORM.</li>
            <li>Enhanced performance and maintainability by converting legacy PHP modules to Node.js and improving REST API efficiency.</li>
            <li>Implemented advanced UI features like infinite scrolling, modular grids, and pagination using <b>Material-UI</b> and CSS.</li>
            </ul>

            <p>My expertise in <b>MERN stack development</b> enables me to design performant APIs and intuitive front-ends, 
            and I’m eager to bring the same impact to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Macharla Venkata Bhanu</b><br>
            Ph: <a href="tel:+916302376836">6302376836</a><br>
            <a href="https://github.com/MacharlaBhanu">GitHub Profile</a> | 
            <a href="https://www.linkedin.com/in/macharla-venkata-bhanu/">LinkedIn Profile</a></p>
            """
            self.bullet1 = "At Bintix, I built and optimized full-stack applications by migrating PHP services to Node.js and improving UI performance using React and MUI."
            self.bullet2 = "I’ve implemented REST APIs with Express and MongoDB, ensuring scalability, modularity, and reliability across services."
            self.bullet3 = "I developed personal projects like a Chat Application and an Entertainment Hub integrating APIs, authentication, and responsive design."
            self.highlights = "React.js, Node.js, Express, MongoDB, RESTful APIs, MUI, and backend optimization"
            self.cta = f"I’d love the opportunity to discuss how I can contribute to {self.company_name}’s engineering team by building scalable, maintainable, and high-impact systems."
        elif self.role == 'Software Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’m writing to express my keen interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact on LinkedIn and wanted to reach out directly regarding potential opportunities 
            in your engineering team. Thank you for taking the time to consider my application.</p>

            <p>I hold a B.Tech in Mechanical Engineering from <b>IIT Patna (2024)</b> and currently work as a 
            <b>Software Development Engineer (SDE)</b> at <b>Bintix</b>. Over the past year, I’ve worked on building 
            scalable software solutions, optimizing algorithms, and improving system efficiency across multiple projects.</p>

            <ul>
            <li>Designed and implemented RESTful APIs and backend services in <b>Node.js and Python</b> for enterprise applications.</li>
            <li>Optimized data processing pipelines, reducing computation time and improving throughput for high-volume applications.</li>
            <li>Developed modular and reusable code for internal tools and client-facing applications, improving maintainability.</li>
            </ul>

            <p>My expertise in <b>backend architecture, API development, and algorithm optimization</b> allows me to deliver high-quality software solutions, 
            and I’m excited to bring this experience to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Macharla Venkata Bhanu</b><br>
            Ph: <a href="tel:+916302376836">6302376836</a><br>
            <a href="https://github.com/MacharlaBhanu">GitHub Profile</a> | 
            <a href="https://www.linkedin.com/in/macharla-venkata-bhanu/">LinkedIn Profile</a></p>
            """

            self.bullet1 = "Developed RESTful APIs and backend services in Node.js and Python, ensuring scalability and reliability."
            self.bullet2 = "Optimized data pipelines and algorithms to enhance system efficiency and throughput."
            self.bullet3 = "Created modular, reusable code for internal and client-facing applications, improving maintainability."
            self.highlights = "Node.js, Python, RESTful APIs, Backend Development, Data Pipelines, Algorithm Optimization"
            self.cta = f"I’d love the opportunity to discuss how I can contribute to {self.company_name}’s engineering team by building robust, scalable software solutions."


        elif self.role == 'Backend Developer':
            self.email_body = f"""
            <p>Hi {self.recruiter},</p>

            <p>I’m writing to express my interest in the <b>{self.role_name}</b> role at <b>{self.company_name}</b>. 
            I came across your contact on LinkedIn and wanted to connect regarding backend development opportunities. 
            Thank you for your time in considering my application.</p>

            <p>I hold a B.Tech in Mechanical Engineering from <b>IIT Patna (2024)</b> and currently work as a 
            <b>Backend Developer</b> at <b>Bintix</b>. Over the past year, I’ve focused on designing scalable backend architectures, 
            building REST APIs, and integrating databases for high-performance applications.</p>

            <ul>
            <li>Built and maintained backend services using <b>Node.js, Express, and MongoDB</b> for production applications.</li>
            <li>Optimized database queries and API responses to improve application speed and reduce server load.</li>
            <li>Implemented authentication, authorization, and data validation for secure and reliable applications.</li>
            </ul>

            <p>My expertise in <b>backend development, database management, and API design</b> allows me to create scalable and reliable systems, 
            and I’m eager to bring this expertise to <b>{self.company_name}</b>.</p>

            <p><b>Best regards,</b><br>
            <b>Macharla Venkata Bhanu</b><br>
            Ph: <a href="tel:+916302376836">6302376836</a><br>
            <a href="https://github.com/MacharlaBhanu">GitHub Profile</a> | 
            <a href="https://www.linkedin.com/in/macharla-venkata-bhanu/">LinkedIn Profile</a></p>
            """

            self.bullet1 = "Built and maintained backend services with Node.js, Express, and MongoDB for production applications."
            self.bullet2 = "Optimized database queries and API responses to enhance performance and reduce server load."
            self.bullet3 = "Implemented secure authentication, authorization, and data validation mechanisms for reliable systems."
            self.highlights = "Node.js, Express, MongoDB, RESTful APIs, Backend Architecture, Security, Performance Optimization"
            self.cta = f"I’d love the opportunity to discuss how I can contribute to {self.company_name}’s backend team by building scalable, secure, and high-performance systems."

        self.TEMPLATE = """{today}
                Hiring Manager
                {company}

                Dear {hiring_manager},

                I’m enthusiastic about applying for the <b>{role}</b> role at <b>{company}</b>, where I see a strong alignment between my
                experience and your product-focused engineering culture.

                I specialize in {highlights}, and I’ve used these skills to deliver measurable impact.

                Why {company}? {why_company}

                What I bring:

                • <b>{bullet1}</b>

                • <b>{bullet2}</b>

                • <b>{bullet3}</b>

                {cta}

                <b>Best regards,</b>  
                <b>Macharla Venkata Bhanu</b>  
                6302376836 | ✉ macharlabhanu169@gmail.com  
                GitHub: https://github.com/MacharlaBhanu  
                LinkedIn: https://www.linkedin.com/in/macharla-venkata-bhanu/"""
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
        self.name = 'Macharla Venkata Bhanu'
        self.official_role = 'Full Stack Developer'
        self.pdf_filename = generate_cover_letter_pdf(self.text, self.name.capitalize(), self.official_role,
                                                      f"{st.session_state.get('username')} Cover Letter.pdf")
        self.official_name = "macharlabhanu169@gmail.com"
        self.resume_path = r"Macharla Venkata Bhanu Resume.pdf"
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
        role = st.selectbox("Select the Role", ['Data Analyst', 'Data Scientist', 'Data Engineer','Machine Learning Engineer', 'Data Governance Analyst', 'Product Analyst', 'Python Developer'])
    elif user == 'sakshi':
        role = st.selectbox("Select the Role", ['Full Stack Developer', 'Frontend Developer', 'Backend Developer', 'Software Developer', 'Process Associate'])
    elif user == 'sai':
        role = st.selectbox("Select the Role", ['Full Stack Engineer', 'Android Developer', 'Frontend Developer', 'Mobile Developer', 'Software Developer', 'Software Engineer'])
    elif user == 'harsha':
        role = st.selectbox("Select the Role", ['Data Analyst', 'Market Researcher', 'Project Manager'])
    elif user == 'bhanu':
        role = st.selectbox("Select the Role", ['Full Stack Developer', 'Software Developer', 'Backend Developer'])

    role_name = st.text_input("Official Role Name (as per Job Posting)", placeholder="Type here...")
    if not role_name:
        st.warning("Role isn't specified in text box taking selected dropdown role name!")
        role_name = role
    job_id = st.text_input("Job ID / Reference Number", placeholder="Type here...")
    st.header("2️⃣ Recruiter & Company Info")
    recruiter_mail = st.text_area(
        "Recruiter's Email(s)",
        placeholder="e.g., adarsh@company.com, rina@company.com",
        height=150
    )

    recipient_list = [email.strip() for email in recruiter_mail.split(",") if email.strip()]

    name_sel = st.radio(
        "Do you want to write custom names for sending recruiters? (Optional)",
        ["no", "yes"],
        horizontal=True
    )
    if "names_dict" not in st.session_state:
        st.session_state["names_dict"] = {}

    recipient_name_list = []

    if name_sel == "yes":
        cols_list = st.columns(len(recipient_list))
        st.warning("If name isn't placed below, default name will be extracted from email!")

        for i, (col, recipient) in enumerate(zip(cols_list, recipient_list)):
            with col:
                st.write(f"Name for email: {recipient}")

                # Use existing session_state value as default
                default_name = st.session_state["names_dict"].get(recipient, "")
                name = st.text_input(
                    "Enter recipient name",
                    value=default_name,
                    key=f"name_{i}"
                )

                if name.strip() == "":
                    # Auto-generate name from email
                    email_username = recipient.split("@")[0]
                    name = email_username.replace(".", " ").replace("_", " ").replace("-", " ")
                    name = re.sub(r"\d+", "", name)
                    name = re.sub(r"\s+", " ", name).strip()
                st.session_state["names_dict"][recipient] = name
                recipient_name_list.append(name)

    else:
        for recipient in recipient_list:
            email_username = recipient.split("@")[0]
            name = email_username.replace(".", " ").replace("_", " ").replace("-", " ")
            name = re.sub(r"\d+", "", name)
            name = re.sub(r"\s+", " ", name).strip()
            recipient_name_list.append(name)
    company_name = st.text_input("Company Name", placeholder="Type here...")
    st.header("3️⃣ Motivation & Customization")
    catchy_subject = st.text_input(
        "Write catchy subject to attract recruiters ✨",
        placeholder="Write witty subject..."
    )
    why_company = st.text_area(
        "Why do you want to join this company?",
        placeholder="Write 2–3 short paragraphs about your motivation...",
        height=150  # you can adjust this value
    )
    job_description = st.text_area(
        "Job description link for future follow ups?",
        placeholder="Job description link...",
        height=100
    )

    st.markdown("---")
    st.caption("⚠️ Make sure all fields are filled before generating the email.")

    if st.button("Send"):
        has_error = False
        if not recruiter_mail:
            st.warning("⚠️ Please enter at least one recruiter's email.")
            has_error = True
        if not company_name:
            st.warning("⚠️ Please enter the company name.")
            has_error = True
        if not why_company:
            st.warning("⚠️ Please provide your reason for joining the company.")
            has_error = True
        if not job_description:
            st.warning("⚠️ Please paste the job description.")
            has_error = True
        if has_error:
            st.stop()
        pass_dict = {'sakshigawandecse@gmail.com':"illr ufri rqeo cwia", "vishnuvardhan.chowhan@gmail.com": "dipi cqsq sgvz ukof", "pollojukiran06@gmail.com": "hros iuyd sbhs ptgh", "harshajha13@gmail.com": "rnrq afrs yicb kuyg", "macharlabhanu169@gmail.com": "yzml pmkv bzqh ytxz"}
        if not recruiter_mail or not company_name or not role_name:
            st.warning("⚠️ Please fill in all the fields before sending.")
        else:
            st.success("✅ All fields are filled. Ready to send!")
        # ---------------- EMAIL ----------------
        for indx, recipient in enumerate(recipient_list):
            name = recipient_name_list[indx]
            recruiter_name = name.title()
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
            SPREADSHEET_ID = "1bsD_uv_r1uNWn9JD85WWMpwTnxEmuP-Eqm-zlI2tp9U"
            application_details = [date.today().strftime("%Y-%m-%d"),company_name,role_name,job_id,recruiter_name,recipient,msg["Subject"],why_company,job_description,"Sent"]
            log_application(service, SPREADSHEET_ID, user, application_details)
            st.info(f"✅ Application logged in {user}.xlsx")
if __name__ == "__main__":
    main()