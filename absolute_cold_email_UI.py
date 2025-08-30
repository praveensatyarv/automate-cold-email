import streamlit as st
import pandas as pd
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="📧 Absolute Cold Email Sender", layout="centered")
st.title("📧 Absolute Cold Email Sender")

# --- Folder paths ---
TEMPLATE_DIR = "absolute_cold_templates/absolute_cold_template_bodies"
SUBJECT_DIR = "absolute_cold_templates/absolute_cold_template_subjects"
CONTACTS_DIR = "input"
RESUME_DIR = os.getenv('RESUME_DIR')

available_companies = sorted({f.split("_")[0] for f in os.listdir(TEMPLATE_DIR)})
company = st.selectbox("Choose Company", available_companies)
available_versions = sorted({f.split("_")[1].split(".")[0] for f in os.listdir(TEMPLATE_DIR) if f.startswith(company)})
version = st.selectbox("Choose Version", available_versions)

# --- Load available files ---
# template_files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".html")]
# subject_files = [f for f in os.listdir(SUBJECT_DIR) if f.endswith(".txt")]
contacts_files = [f for f in os.listdir(CONTACTS_DIR) if f.endswith((".xlsx", ".xls")) and (f.startswith(f"contacts_{company}") or f.startswith(f"contacts_AAA_DEFAULT"))]
resume_files = [f for f in os.listdir(RESUME_DIR) if f.endswith(".pdf")]

# --- File selection UI ---
# selected_template = st.selectbox("Choose Email Template", template_files)
# selected_subject = st.selectbox("Choose Email Subject", subject_files)
selected_contacts = st.selectbox("Choose Contacts File", contacts_files)
selected_resume = st.selectbox("Choose Resume PDF", resume_files)

# --- Construct file paths ---
template_path = os.path.join(TEMPLATE_DIR, f"{company}_{version}.html")
subject_path = os.path.join(SUBJECT_DIR, f"{company}_{version}.txt")
contacts_path = os.path.join(CONTACTS_DIR, selected_contacts)
resume_path = os.path.join(RESUME_DIR, selected_resume)

# Proceed if all files are selected
if company and version and selected_contacts and selected_resume:
    # Load files
    with open(template_path, 'r', encoding='utf-8') as f:
        template_html = f.read()

    df = pd.read_excel(contacts_path)

    # Show contacts to be emailed
    st.subheader("📇 Contact Preview")
    st.dataframe(df[df['Send Email?'].str.lower() == "yes"])

    with open(subject_path, 'r', encoding='utf-8') as file:
        subject = file.read()

    # Preview single email
    preview_index = st.number_input("Preview Email for Row Index", min_value=0, max_value=len(df)-1, value=0)
    row = df.iloc[preview_index]
    name = row["Full Name"].split(" ")[0]

    email_body = """
    Hi {name},<br>
    {template_html}
    <p>
        Regards,<br>
        Praveen Satya Rajamanickam Vijayaraghavan<br>
        469-471-4540 | <a href="https://linkedin.com/in/praveen-satya-r-v">LinkedIn</a> | <a href="https://github.com/praveensatyarv/">GitHub</a> | <a href="https://praveensatyarv.github.io/portfolio/">Portfolio</a><br>
        Dallas, TX
    </p>
    """

    st.subheader("📬 Email Preview")
    st.markdown(f"**Subject:** {subject}")
    st.markdown(email_body.format(name=name, template_html=template_html), unsafe_allow_html=True)

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

            message = MIMEMultipart()
            message['Subject'] = subject
            message['From'] = smtp_username
            message['To'] = row['Email']

            name = row["Full Name"].split(" ")[0]
            custom_body = email_body.format(name=row['Full Name'].split(" ")[0], template_html=template_html)
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