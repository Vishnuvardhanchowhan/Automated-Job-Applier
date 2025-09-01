import streamlit as st
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import smtplib
from email import encoders
from datetime import date
from textwrap import dedent
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable


# ---------------- PDF GENERATOR ----------------
def generate_cover_letter_pdf(text, filename="vishnuvardhan_cover_letter.pdf"):
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
    elements.append(Paragraph("VISHNUVARDHAN CHOWHAN", name_style))
    elements.append(Paragraph("Data Analytics & Automation Specialist | Hyderabad, India", subtitle_style))

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


# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="Automated Job Application", page_icon="📧")

st.title("📧 Automated Job Application Email Generator")
st.markdown(
    "Fill in the details below to generate a professional email and cover letter for recruiters."
)

st.header("1️⃣ Role & Job Details")
role = st.selectbox("Select the Role", ['Data Analyst', 'Data Scientist', 'Data Engineer', 'Data Governance Analyst', 'Product Analyst'])
role_name = st.text_input("Official Role Name (as per Job Posting)", placeholder="Type here...")
job_id = st.text_input("Job ID / Reference Number", placeholder="Type here...")

st.header("2️⃣ Recruiter & Company Info")
recruiter = st.text_input("Recruiter's Name", placeholder="e.g., Adarsh sir")
recruiter_mail = st.text_input("Recruiter's Email(s)", placeholder="e.g., adarsh@company.com, rina@company.com")
recipient_list = [email.strip() for email in recruiter_mail.split(",") if email.strip()]
company_name = st.text_input("Company Name", placeholder="Type here...")

st.header("3️⃣ Motivation & Customization")
why_company = st.text_input(
    "Why do you want to join this company?",
    placeholder="Write 1–2 lines about your motivation..."
)

st.markdown("---")
st.caption("⚠️ Make sure all fields are filled before generating the email.")

if st.button("Send"):
    if not recruiter_mail or not company_name or not role_name:
        st.warning("⚠️ Please fill in all the fields before sending.")
    else:
        st.success("✅ All fields are filled. Ready to send!")

    if role == 'Data Analyst':
        email_body = f"""
        <p>Hi {recruiter},</p>

        <p>I’d like to express my interest in the <b>{role_name}</b> role at <b>{company_name}</b>. 
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
        and I am eager to bring the same impact to <b>{company_name}</b>.</p>

        <p><b>Best regards,</b><br>
        <b>Vishnuvardhan Chowhan</b><br>
        Ph: 7036363267<br>
        <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
        """

        bullet1 = "I bring experience designing Streamlit-based KPI and Consumer Journey Dashboards, demonstrating my ability to translate complex data into actionable insights for business teams."
        bullet2 = "I have automated SQL + Python reporting pipelines, showing my strength in reducing manual effort and improving data accuracy—skills I can apply to optimize any data process I take on."
        bullet3 = "I developed an Innovations Tracker Dashboard to identify and benchmark product launches, highlighting my capability to build analytics tools that uncover market opportunities."
        highlights = "Python, SQL, Power BI + Streamlit dashboarding, reporting automation"
        cta = f"I’d love the opportunity to discuss how I can contribute to {company_name}’s analytics team and help drive data-driven decision making."
    elif role == 'Data Scientist':
        email_body = f"""
        <p>Hi {recruiter},</p>

        <p>I’d like to express my interest in the <b>{role_name}</b> role at <b>{company_name}</b>. 
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

        <p>My experience in Python, SQL, Machine learning, and Analytics has enabled global clients to improve efficiency and accuracy, and I am eager to bring the same impact to <b>{company_name}</b>.</p>

        <p><b>Best regards,</b><br>
        <b>Vishnuvardhan Chowhan</b><br>
        Ph: <a href="tel:+917036363267">7036363267</a><br>
        <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
        """

        bullet3 = "Completed an Advanced ML course, applying models such as regression, clustering, CNNs, RNNs, LSTMs, and GANs to real-world datasets—demonstrating strong foundations in both classical and deep learning."
        bullet2 = "Designed an AI-powered chatbot agent that extracts KPIs and custom data cuts from natural language queries, showcasing my ability to integrate NLP with business analytics."
        bullet1 = "Applied ML techniques for trend identification and product benchmarking, illustrating my capacity to convert raw data into actionable insights for strategic decision-making."
        highlights = "AI/ML, Python, advanced analytics, natural language data agents"
        cta = f"I’d welcome the chance to explore how my machine learning expertise and real-world project experience can support {company_name}’s data science initiatives."
    elif role == 'Data Engineer':
        email_body = f"""
        <p>Hi {recruiter},</p>

        <p>I’d like to express my interest in the <b>{role_name}</b> role at <b>{company_name}</b>. 
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

        bullet1 = "Designed and deployed Python-based ETL pipelines to clean, transform, and standardize millions of rows of data, demonstrating my ability to manage and scale large datasets efficiently."
        bullet2 = "Automated end-to-end reporting workflows, reducing manual effort by 80% and highlighting my strength in streamlining repetitive processes through automation."
        bullet3 = "Built robust SQL and Python scripts for data cleaning, validation, and transformation, showcasing my focus on improving pipeline reliability and ensuring data accuracy."
        highlights = "Python, SQL, PySpark, Big Data, ETL pipelines, workflow automation"
        cta = f"I’d be glad to discuss how my background in building reliable data workflows can strengthen {company_name}’s data infrastructure and support downstream analytics."
    elif role == 'Data Governance Analyst':
        email_body = f"""
        <p>Hi {recruiter},</p>

        <p>I’d like to express my interest in the <b>{role_name}</b> role at <b>{company_name}</b>. 
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
        also trusted. I’m excited to bring this blend of expertise to <b>{company_name}</b>.</p>

        <p><b>Best regards,</b><br>
        <b>Vishnuvardhan Chowhan</b><br>
        Ph: <a href="tel:+917036363267">7036363267</a><br>
        <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
        """

        bullet1 = "Built automated pipelines that validate, transform, and reconcile data across multiple sources, ensuring governance and integrity in reporting."
        bullet2 = "Created dashboards that integrate data lineage and business KPIs, improving both transparency and decision-making."
        bullet3 = "Worked with global clients where ensuring data compliance, accuracy, and auditability was critical—strengthening my governance-first mindset."
        highlights = "Data governance, Python, SQL, ETL, pipeline validation, data lineage, compliance"
        cta = f"I’d be glad to discuss how my combined experience in analytics and engineering can help strengthen {company_name}’s data governance and reliability frameworks."
    elif role == 'Product Analyst':
        email_body = f"""
        <p>Hi {recruiter},</p>

        <p>I’d like to express my interest in the <b>{role_name}</b> role at <b>{company_name}</b>. 
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
        mindset to <b>{company_name}</b>.</p>

        <p><b>Best regards,</b><br>
        <b>Vishnuvardhan Chowhan</b><br>
        Ph: <a href="tel:+917036363267">7036363267</a><br>
        <a href="https://notion-sparkle-site.lovable.app/">Portfolio</a></p>
        """

        bullet1 = "Built KPI dashboards for funnels, retention, and adoption, enabling product teams to track performance and iterate faster."
        bullet2 = "Automated SQL + Python pipelines for real-time consumer insights, reducing latency between data collection and decision-making."
        bullet3 = "Developed an Innovations Tracker Dashboard that benchmarked competitor products, providing strategic inputs to product roadmaps."
        highlights = "Product analytics, SQL, Python, dashboarding, funnel analysis, retention, consumer insights"
        cta = f"I’d welcome the opportunity to show how my data-driven approach can support {company_name}’s product growth and decision-making."

    TEMPLATE = """{today}
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

    text = TEMPLATE.format(
        today=date.today().strftime("%B %d, %Y"),
        hiring_manager=recruiter or "Hiring Manager",
        role=role_name,
        company=company_name,
        highlights=highlights,
        why_company=why_company,
        bullet1=bullet1,
        bullet2=bullet2,
        bullet3=bullet3,
        cta=cta
    )

    text = dedent(text)

    # Generate PDF
    pdf_filename = generate_cover_letter_pdf(text)

    # ---------------- EMAIL ----------------
    msg = MIMEMultipart()
    if not job_id:
        msg["Subject"] = f"{role_name} Role Application – Resume & Cover Letter"
    else:
        msg["Subject"] = f"{role_name} Application [{job_id}] – Resume & Cover Letter"
    msg["From"] = "vishnuvardhan.chowhan@gmail.com"
    msg["To"] = ", ".join(recipient_list)

    # # Send the email
    #
    # msg["To"] = recruiter_mail
    body = MIMEText(email_body, "html")
    msg.attach(body)

    # Attach Resume
    resume_path = r"C:\Users\vishn\OneDrive\Documents\vishnu docs\Projects\VishnuvardhanChowhanResume.pdf"
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
    smtp.login("vishnuvardhan.chowhan@gmail.com", "dipi cqsq sgvz ukof")  # 🔑 Use app password
    # smtp.sendmail(msg["From"], msg["To"].split(","), msg.as_string())
    smtp.sendmail(msg["From"], recipient_list, msg.as_string())
    smtp.quit()

    st.success("📧 Email with Resume + Cover Letter PDF sent successfully!")
