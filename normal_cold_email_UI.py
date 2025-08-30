import streamlit as st
import pandas as pd
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="📧 Daily Cold Email Sender", layout="centered")
st.title("📧 Daily Cold Email Sender")

# --- Folder paths ---
CONTACTS_DIR = "input"
RESUME_DIR = os.getenv('RESUME_DIR')

# --- Load resume ---
resume_files = [f for f in os.listdir(RESUME_DIR) if f.endswith(".pdf")]
selected_resume = st.selectbox("Choose Resume PDF", resume_files)

resume_path = os.path.join(RESUME_DIR, selected_resume)
contacts_path = os.path.join(CONTACTS_DIR, "contacts_Daily.xlsx")

# Proceed if all files are selected
if selected_resume:
    subject = """
        Praveen Satya: Excited About the {role} at {company}
    """

    email_body = """
    <p>
        Hi {name},<br>
        I hope this note finds you well. I came across the {role} role at {company} and am excited about the opportunity to support your team with impactful data solutions.
    </p>

    <p>
        I’m Praveen Satya, a Data Analyst with 3+ years of experience and an MS in Business Analytics and AI from UT Dallas. A few highlights about me:
        <ul>
            <li><b>Prevented $100K+ quarterly revenue loss</b> by deploying a BigQuery/Airflow platform for real-time anomaly detection across 70+ business metrics.</li>
            <li><b>Saved $325K in 6 months</b> by benchmarking AI-driven logistics assignments, optimizing cost and efficiency, and presenting insights to executives.</li>
            <li><b>Eliminated manual investigation for 15 analysts</b> (saving 40+ hours per case) by rolling out instant, AI-powered root-cause analysis.</li>
            <li><b>Built end-to-end AWS data pipelines</b> and Tableau dashboards to automate logistics analytics, enabling real-time insights and 90% faster reporting.</li>
        </ul>
    </p>
    
    <p>
         I’m eager to bring my technical and analytical skills to your team. Would you be open to a quick 15-minute call to discuss how I can contribute?
    </p>

    <p>
        Thank you for your time. I look forward to connecting!
    </p>

    <p>
        Regards,<br>
        Praveen Satya Rajamanickam Vijayaraghavan<br>
        469-471-4540 | <a href="https://linkedin.com/in/praveen-satya-r-v">LinkedIn</a> | <a href="https://github.com/praveensatyarv/">GitHub</a> | <a href="https://praveensatyarv.github.io/portfolio/">Portfolio</a><br>
        Dallas, TX
    </p>
    """

    df = pd.read_excel(contacts_path)

    # Show contacts to be emailed
    st.subheader("📇 Contact Preview")
    st.dataframe(df[df['Send Email?'].str.lower() == "yes"])

    # Preview single email
    preview_index = st.number_input("Preview Email for Row Index", min_value=0, max_value=len(df)-1, value=0)
    row = df.iloc[preview_index]
    name = row["Full Name"].split(" ")[0]
    role = row["Role"].split(" ")[0]
    company = row["Company"].split(" ")[0]


    st.subheader("📬 Email Preview")
    st.markdown(f"**Subject:** {subject.format(role=role, company=company)}")
    st.markdown(email_body.format(name=name, role=role, company=company), unsafe_allow_html=True)

    # Email Sending
    if st.button("🚀 Send Emails"):
        import smtplib
        from email.mime.application import MIMEApplication

        smtp_username = os.getenv('SMTP_USERNAME')
        smtp_password = os.getenv('SMTP_PASSWORD')

        with open(resume_path, "rb") as f:
            pdf_bytes = f.read()

        for idx, row in df.iterrows():
            if row['Send Email?'].lower() != "yes":
                continue

            name = row["Full Name"].split(" ")[0]
            role = row["Role"].split(" ")[0]
            company = row["Company"].split(" ")[0]

            custom_subject = subject.format(role=role, company=company)

            message = MIMEMultipart()
            message['Subject'] = custom_subject
            message['From'] = smtp_username
            message['To'] = row['Email']


            custom_body = email_body.format(name=name, role=role, company=company)
            message.attach(MIMEText(custom_body, 'html'))

            # Attach PDF
            pdf_attachment = MIMEApplication(pdf_bytes, _subtype="pdf")
            pdf_attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(resume_path))
            message.attach(pdf_attachment)

            try:
                server = smtplib.SMTP('smtp.gmail.com:587')
                server.starttls()
                server.login(smtp_username, smtp_password)
                server.sendmail(smtp_username, row['Email'], message.as_string())
                server.quit()
                st.success(f"✅ Email sent to {row['Email']}")
            except Exception as e:
                st.error(f"❌ Error sending to {row['Email']}: {e}")