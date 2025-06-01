import pandas as pd
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import logging
import os
import traceback

# --- Load credentials from .env ---
load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

# --- Configuration ---
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EXCEL_FILE = 'C:/HTC-2025/PYTHON/EXCERCISE/q2/emi_schedule.xlsx'
LOG_FILE = 'email_errors.log'

# --- Logging configuration ---
logging.basicConfig(filename=LOG_FILE, level=logging.ERROR,
                    format='%(asctime)s:%(levelname)s:%(message)s')

# --- Read Excel File ---
try:
    df = pd.read_excel(EXCEL_FILE)
except Exception as e:
    logging.error(f"Failed to read Excel file: {e}")
    exit()

# --- Filter for EMIs due today ---
try:
    df['EMIDueDate'] = pd.to_datetime(df['EMIDueDate']).dt.date
    today = datetime.today().date()
    due_today = df[df['EMIDueDate'] == today]
except Exception as e:
    logging.error(f"Failed to filter due dates: {e}")
    exit()

# --- Email Sender Function ---
def send_email(to_email, customer_name, loan_number, amount, due_date):
    try:
        subject = f"EMI Due Alert - Loan #{loan_number}"
        body = f"""Dear {customer_name},

This is a reminder that your EMI of ₹{amount} for your Loan Account Number {loan_number} is due today ({due_date}).

Kindly ensure the payment is made to avoid penalties.

Regards,  
Bank Loan Department
"""

        message = MIMEMultipart()
        message['From'] = SENDER_EMAIL
        message['To'] = to_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(message)

        print(f"✅ Email sent to {customer_name} at {to_email}")

    except Exception as e:
        error_detail = traceback.format_exc()
        logging.error(f"Failed to send email to {customer_name} ({to_email}): {e}\n{error_detail}")

# --- Send Emails ---
if not due_today.empty:
    for _, row in due_today.iterrows():
        send_email(
            to_email=row['Email'],
            customer_name=row['CustomerName'],
            loan_number=row['LoanAccountNumber'],
            amount=row['EMIAmount'],
            due_date=row['EMIDueDate']
        )
else:
    print("📭 No EMIs due today.")
