import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Configuration
gmail_user = os.getenv('GMAIL_USER')
gmail_password = os.getenv('GMAIL_PASSWORD')  # Use your regular Google account password here
subject = 'Subject of the Email'
excel_file_path = 'files/data.xlsx'
attachment_folder = 'files/attached_files/'

# New email body template
body_template = """
Dear Hiring Manager,

With extensive experience at Raymond Limited in Supply Chain management, Warehouse, Production, Quality control management, I am willing to express my interest in managerial roles at {company_name} suitable for my profile. Please find my cover letter and CV attached for reference.

Sincerely,
Channabasavaraj Banagar
LinkedIn: https://www.linkedin.com/in/channabasavaraj-banagar-5069441b0/
Contact no: 9920751247
"""

def send_email(to_email, company_name, files):
   # Setup the MIME
   message = MIMEMultipart()
   message['From'] = gmail_user
   message['To'] = to_email
   message['Subject'] = f"Application to suitable roles at {company_name}"

   # Substitute company_name in the body template
   body = body_template.format(company_name=company_name)

   # Attach the body with the msg instance
   message.attach(MIMEText(body, 'plain'))

   # Attach all files from the specified folder
   for file in files:
       with open(file, 'rb') as attachment:
           part = MIMEBase('application', 'octet-stream')
           part.set_payload(attachment.read())
           encoders.encode_base64(part)
           part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(file)}")
           message.attach(part)

   # Create SMTP session for sending the mail
   try:
       server = smtplib.SMTP('smtp.gmail.com', 587)
       server.starttls()  # Enable security
       server.login(gmail_user, gmail_password)  # Login with your regular Google account password
       text = message.as_string()
       server.sendmail(gmail_user, to_email, text)
       server.quit()
       print(f'Email sent to {company_name} at {to_email}')
   except Exception as e:
       print(f'Failed to send email to {company_name} at {to_email}. Error: {str(e)}')

# Read the Excel file from the "Current" sheet
try:
   xl = pd.ExcelFile(excel_file_path)
   df = xl.parse('Current')  # Specify the sheet name 'Current'
except FileNotFoundError:
   print(f"Excel file '{excel_file_path}' not found.")
   exit()
except Exception as e:
   print(f"Failed to read Excel file: {str(e)}")
   exit()

# Get list of files in the attachment folder
files_to_attach = [os.path.join(attachment_folder, file) for file in os.listdir(attachment_folder)]

# Send email to each company
for index, row in df.iterrows():
   send_email(row['Company email id'], row['Company Name'], files_to_attach)
