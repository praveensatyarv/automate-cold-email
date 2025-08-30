import pandas as pd
import smtplib
import os
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

load_dotenv()


smtp_username = os.getenv('SMTP_USERNAME')
smtp_password = os.getenv('SMTP_PASSWORD')
pdf_path_env = os.getenv('PDF_PATH')
# excel_file = 'input/contacts_AMD_supply_chain.csv'
excel_file = 'input/contacts.xlsx'
template_path = 'absolute_cold_templates/AMD.html'

# make sure that the column names containing name, company name and email are "First Name", "Company" and "Email" resp
df = pd.read_excel(excel_file)

with open(template_path, 'r', encoding='utf-8') as file:
    email_template = file.read()


email_body_template = """
Hi {name},
{email_body}
<p>
    Regards,<br>
    Praveen Satya Rajamanickam Vijayaraghavan<br>
    469-471-4540 | <a href="{linkedin_link}">LinkedIn</a> | <a href="{github_link}">GitHub</a> | <a href="{portfolio_link}">Portfolio</a><br>
    Dallas, TX
</p>
"""


try:
    for index, row in df.iterrows():
        if row['Send Email?'].lower() == "yes":
            message = MIMEMultipart()
            message['Subject'] = f"Praveen Satya | Cold Outreach – Helping AMD Navigate Supply Chain Disruptions with Data"
            message['From'] = smtp_username
            message['To'] = row['Email']

            email_body = email_body_template.format(
                name=row['Full Name'].split(" ")[0], 
                company=row['Company'], 
                linkedin_link="https://linkedin.com/in/praveen-satya-r-v",
                github_link="https://github.com/praveensatyarv/",
                portfolio_link="https://praveensatyarv.github.io/portfolio/",
                email_body=email_template
                ## add other links as needed
                )
            message.attach(MIMEText(email_body, 'html'))
            
            # pdf attach (resume or other doc)
            pdf_path = pdf_path_env
            with open(pdf_path, "rb") as pdf_file:
                pdf_attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
                pdf_attachment.add_header('Content-Disposition', f'attachment; filename=Resume_Praveen_Satya_R_V.pdf')
                message.attach(pdf_attachment)
            
            server = smtplib.SMTP('smtp.gmail.com:587')
            server.ehlo()
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(smtp_username, row['Email'], message.as_string())
            server.quit()
            print("email sent",index+1)
except Exception as e:
    print("error",e)

   