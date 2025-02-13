import smtplib
import shutil
import openpyxl
import os
import time
from datetime import datetime
from email.message import EmailMessage

# Configuration
EXCEL_FILE = "email_quantification.xlsx"
OUTPUT_FOLDER = "Processed_Reports"
EMAIL_SENDER = "your_email@example.com"
EMAIL_PASSWORD = "your_password" # Use environment variables instead
EMAIL_RECIPIENTS = ["recipient1@example.com", "recipient2@example.com"]
SMTP_SERVER = "smtp.example.com" # Update with your email provider's SMTP
SMTP_PORT = 587 # Usually 587 for TLS

# Ensure output folder exists
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def copy_and_send_report():
    """Copy extracted email details to a new Excel file and send via email."""
    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M")
    loop_number = now.hour + 1 # Assuming 1-hour intervals; Loop 1 at 12 AM, Loop 2 at 1 AM, etc.
    loop_name = f"Loop {loop_number}"

    new_excel_name = f"{OUTPUT_FOLDER}/Email_Report_{timestamp}.xlsx"

    # Copy original Excel file
    shutil.copy(EXCEL_FILE, new_excel_name)

    # Append loop name to the new Excel file
    workbook = openpyxl.load_workbook(new_excel_name)
    sheet = workbook.active
    sheet.append(["", "", "", "", "", "", f"Run: {loop_name}"])
    workbook.save(new_excel_name)

    print(f"Report saved as {new_excel_name}")
    
    # Send email with attachment
    send_email(new_excel_name, loop_name)

def send_email(file_path, loop_name):
    """Send email with the generated report as an attachment."""
    msg = EmailMessage()
    msg["Subject"] = f"Email Report - {loop_name}"
    msg["From"] = EMAIL_SENDER
    msg["To"] = ", ".join(EMAIL_RECIPIENTS)
    msg.set_content(f"Attached is the email report for {loop_name}.\n\nBest Regards,\nAutomation Script")

    # Attach Excel file
    with open(file_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=os.path.basename(file_path))

    # Send email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
        print(f"Email sent successfully for {loop_name}.")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Call the function after email processing
copy_and_send_report()
