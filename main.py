import smtplib
import ssl
import os
import pandas as pd
import gspread
import json
from google.cloud import secretmanager
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from datetime import datetime
import uuid
from urllib.parse import quote

import os.path
from google.oauth2 import service_account
from googleapiclient.discovery import build


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


#--- CONFIG ---
GOOGLE_SHEET_NAME = "SBI General Interest Form (Responses)"
LOGO_FILE = "EmailSignature.gif"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
BOOKING_BASE_URL = "https://sbi-booking-400556956516.us-central1.run.app"




def get_secret(secret_id):
    """Retrieve a secret from Google Secret Manager."""
    client = secretmanager.SecretManagerServiceClient()
    project_id = os.environ.get('GOOGLE_CLOUD_PROJECT')
    name = f"projects/{project_id}/secrets/{secret_id}/versions/latest"
    response = client.access_secret_version(request={"name": name})
    return response.payload.data.decode("UTF-8")




def get_email_credentials():
    """Retrieve email credentials from Secret Manager."""
    sender_email = get_secret("EMAIL_USER")
    sender_password = get_secret("GOOGLE_PASS")
    return sender_email, sender_password




def get_new_signups():
    """Retrieve all records from Google Sheets Database."""
    try:
        # Get credentials from Secret Manager
        credentials_json = get_secret("SERVICE_ACCOUNT_FILE")
        credentials_dict = json.loads(credentials_json)
       
        # Create temporary file
        with open('/tmp/credentials.json', 'w') as f:
            json.dump(credentials_dict, f)
       
        gc = gspread.service_account(filename='/tmp/credentials.json')
       
        # Open the spreadsheet
        sh = gc.open(GOOGLE_SHEET_NAME).worksheet("Form Responses 1")

        # Get all records from the sheet and convert to a pandas DataFrame
        records = sh.get_all_records()
        print(f"Total records found: {len(records)}")
       
        df = pd.DataFrame(records)
        print(f"DataFrame shape: {df.shape}")
        print(f"Column names: {df.columns.tolist()}")
       
        df['original_row_index'] = df.index + 2
       
        return sh, df  # Return ALL records
   
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"Error: Spreadsheet '{GOOGLE_SHEET_NAME}' not found.")
        return None, pd.DataFrame()
   
    except Exception as e:
        print(f"An error occurred while fetching data: {e}")
        return None, pd.DataFrame()




def send_welcome_email(recipient_name, recipient_email, departments_str):
    """Create custom welcome email body based on user input."""
    try:
        sender_email, sender_password = get_email_credentials()
    except Exception as e:
        print(f"Error: Could not retrieve email credentials: {e}")
        return False
   
    try:
        message = MIMEMultipart("related")
        message["Subject"] = "Next Steps With SBI!"
        message["From"] = sender_email
        message["To"] = recipient_email
       
        # Department descriptions
        department_info = {
            'Research and Development': 'Researches and advises on cutting-edge sustainable materials, technologies, and methodologies, and conducts post-project analysis.',
            'Finance': 'Manages all financial aspects of projects, from initial budgeting and expense tracking, invoicing, and final financial reporting.',
            'Tech': 'Identifies, designs, and implements internal and external technologies with AI integration, managing software and system installation.',
            'Engineering': 'Designs and oversees the structural, mechanical (HVAC, plumbing), and electrical systems, including renewable energy integration and site planning.',
            'Architecture': 'Responsible for the aesthetic and functional design of projects, creating concepts, detailed drawings, and selecting sustainable materials.',
            'Public Relations': 'Handles recruitment, internal and external communications, project announcements, and public events, maintaining team morale.',
            'Legal': 'Manages contracts, ensures regulatory compliance, handles permitting, and oversees all legal aspects of the project.',
        }

        departments_list = [dep.strip() for dep in departments_str.split(',') if dep.strip()]
       
        # Generate the HTML for the department selections
        departments_html = ''
        for dep in departments_list:
            if dep in department_info:
                departments_html += f"<strong>{dep}:</strong> {department_info[dep]}<br><br>"
            else:
                print(f"Warning: Unknown department '{dep}'")
                departments_html += f"<strong>{dep}:</strong> Department information not available.<br><br>"
       
        # The HTML email content
        html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Next Steps With SBI!</title>
</head>
<body style="margin: 0; padding: 0; font-family: Arial, sans-serif;">
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px;">
       
        <p style="margin-bottom: 16px;">Hello {recipient_name},</p>
       
        <p style="margin-bottom: 16px;">Thank you so much for completing our form and for your interest in joining Sustainable Building Initiative!</p>
       
        <p style="margin-bottom: 16px;">We wanted to give you a quick update on what happens next:</p>
       
        <div style="margin-bottom: 20px;">
            <p style="margin-bottom: 8px;"><strong>Your Department Selections:</strong></p>
            <div style="margin-left: 20px; margin-bottom: 16px;">
                {departments_html}
            </div>
        </div>
       
        <div style="margin-bottom: 20px;">
            <p style="margin-bottom: 8px;"><strong>What to Expect Next:</strong></p>
            <p style="margin-left: 20px; margin-bottom: 16px;">
                Our team will review your responses and reach out to you with more details about each department you've selected. You'll also receive updates about opportunities, events, and important information for the upcoming semester.
            </p>
        </div>
       
        <p style="margin-bottom: 16px;">If you have any questions, feel free to reply to this email or contact our President at <a href="mailto:px.guzman@utexas.edu" style="color: #0066cc; text-decoration: none;">px.guzman@utexas.edu</a>.</p>
       
        <p style="margin-bottom: 16px;">Stay tuned—we're excited to connect with you soon!</p>
       
        <p style="margin-bottom: 20px;">
            <strong>Best regards,</strong><br>
            The Sustainable Building Initiative Team
        </p>
       
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">
            <img src="cid:logo" alt="SBI Logo" style="max-width: 120px; height: auto; display: block; margin-bottom: 10px;" />
            <p style="margin: 0; font-size: 12px; color: #666666;">
                Website: <a href="https://utsbi.org" target="_blank" style="color: #0066cc; text-decoration: none;">utsbi.org</a>
            </p>
        </div>
       
    </div>
</body>
</html>
"""

        html_part = MIMEText(html_content, "html")
        message.attach(html_part)
       
        # Add logo if file exists
        if os.path.exists(LOGO_FILE):
            with open(LOGO_FILE, "rb") as f:
                logo_data = f.read()
                logo = MIMEImage(logo_data)
                logo.add_header("Content-ID", "<logo>")
                message.attach(logo)
               
        # Create SMTP session and send email
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())
               
        print(f"Welcome email sent successfully to {recipient_email}")
        return True
   
    except Exception as e:
        print(f"Failed to send welcome email to {recipient_email}: {e}")
        return False


def send_interview_email(recipient_name, recipient_email, department, booking_link):
    """Send separate interview scheduling email."""
    try:
        sender_email, sender_password = get_email_credentials()
    except Exception as e:
        print(f"Error: Could not retrieve email credentials: {e}")
        return False
   
    try:
        message = MIMEMultipart("related")
        message["Subject"] = "Schedule Your SBI Interview"
        message["From"] = sender_email
        message["To"] = recipient_email
       
        html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Schedule Your SBI Interview</title>
</head>
<body style="margin: 0; padding: 0; font-family: Arial, sans-serif;">
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px;">
       
        <p style="margin-bottom: 16px;">Hello {recipient_name},</p>
       
        <p style="margin-bottom: 16px;">Great news! We'd like to invite you to interview for the <strong>{department}</strong> department at Sustainable Building Initiative.</p>
       
        <div style="margin-bottom: 20px; padding: 20px; background-color: #f0f8ff; border-left: 4px solid #0066cc; text-align: center;">
            <p style="margin-bottom: 16px; font-size: 16px;"><strong>📅 Schedule Your Interview</strong></p>
            <p style="margin-bottom: 20px;">
                Click the button below to choose a time that works best for you:
            </p>
            <a href="{booking_link}" style="display: inline-block; padding: 14px 28px; background-color: #0066cc; color: white; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;">Book Your Interview Time</a>
        </div>
       
        <p style="margin-bottom: 16px;">The interview will be approximately 30 minutes and will be held at:</p>
        <p style="margin-left: 20px; margin-bottom: 16px;"><strong>McCombs School of Business<br>2110 Speedway, Austin, TX 78705</strong></p>
        
        <p style="margin-bottom: 16px; font-size: 13px; color: #666;">We will message you through text beforehand about the exact location, or if the location changes.</p>
       
        <p style="margin-bottom: 16px;">We're looking forward to meeting you!</p>
       
        <p style="margin-bottom: 20px;">
            <strong>Best regards,</strong><br>
            The Sustainable Building Initiative Team
        </p>
       
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">
            <img src="cid:logo" alt="SBI Logo" style="max-width: 120px; height: auto; display: block; margin-bottom: 10px;" />
            <p style="margin: 0; font-size: 12px; color: #666666;">
                Website: <a href="https://utsbi.org" target="_blank" style="color: #0066cc; text-decoration: none;">utsbi.org</a>
            </p>
        </div>
       
    </div>
</body>
</html>
"""

        html_part = MIMEText(html_content, "html")
        message.attach(html_part)
       
        # Add logo if file exists
        if os.path.exists(LOGO_FILE):
            with open(LOGO_FILE, "rb") as f:
                logo_data = f.read()
                logo = MIMEImage(logo_data)
                logo.add_header("Content-ID", "<logo>")
                message.attach(logo)
               
        # Create SMTP session and send email
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())
               
        print(f"Interview email sent successfully to {recipient_email}")
        return True
   
    except Exception as e:
        print(f"Failed to send interview email to {recipient_email}: {e}")
        return False




def update_email_sent_status(sheet, row_index):
    """Updates Google Sheets to mark sent."""
    try:
        headers = sheet.row_values(1)
        email_sent_col = headers.index("Automated Email Sent") + 1
       
        sheet.update_cell(row_index, email_sent_col, "Yes")
        print(f"Updated row {row_index} - marked email as sent")
        return True
   
    except Exception as e:
        print(f"Failed to update row {row_index}: {e}")
        return False


def update_interview_sent_status(sheet, row_index, booking_id):
    """Updates Google Sheets with booking link ID."""
    try:
        headers = sheet.row_values(1)
        interview_sent_col = headers.index("Interview Sent") + 1
       
        sheet.update_cell(row_index, interview_sent_col, booking_id)
        print(f"Updated row {row_index} - marked interview link sent: {booking_id}")
        return True
   
    except Exception as e:
        print(f"Failed to update row {row_index}: {e}")
        return False




def main():  
    # Get all records from Google Sheets
    sheet, all_records_df = get_new_signups()
   
    if sheet is None or all_records_df.empty:
        print("No records found or error accessing sheet.")
        return
   
    # Process welcome emails (for those who haven't received it yet)
    new_signups_df = all_records_df[all_records_df["Automated Email Sent"] == ''].copy()
    print(f"Found {len(new_signups_df)} new signups for welcome emails.")
   
    for index, row in new_signups_df.iterrows():
        name = row.get("What is your name?", "")
        email = row.get("What is your email?", "")
        departments = row.get("Which department(s) do you want to be in? (Pick up to 2)", "")
        original_row = row.get("original_row_index")

        if not email or not name:
            print(f"Skipping row {original_row}: missing name or email")
            continue
       
        print(f"Processing welcome email: {name} ({email})")
        name = name.title()

        if send_welcome_email(name, email, departments):
            update_email_sent_status(sheet, original_row)
            print(f"Successfully sent welcome email to {name}")
        else:
            print(f"Failed to send welcome email to {name}")
    
    # Process interview emails (separate from welcome emails)
    # Filter for: Give Interview = "Yes" AND Interview Sent is empty
    interview_needed_df = all_records_df[
        (all_records_df["Give Interview"].astype(str).str.strip().str.lower() == "yes") &
        (all_records_df["Interview Sent"] == '')
    ].copy()
    
    print(f"\nFound {len(interview_needed_df)} people needing interview emails.")
    
    for index, row in interview_needed_df.iterrows():
        name = row.get("What is your name?", "")
        email = row.get("What is your email?", "")
        departments = row.get("Which department(s) do you want to be in? (Pick up to 2)", "")
        original_row = row.get("original_row_index")

        if not email or not name:
            print(f"Skipping row {original_row}: missing name or email")
            continue
       
        print(f"Processing interview email: {name} ({email})")
        name = name.title()
        
        booking_id = str(uuid.uuid4())
        dept_param = str(departments) if departments else "Tech"
        booking_link = f"{BOOKING_BASE_URL}?id={booking_id}&name={quote(name)}&email={quote(email)}&dept={quote(dept_param)}"
        
        if send_interview_email(name, email, dept_param, booking_link):
            update_interview_sent_status(sheet, original_row, booking_id)
            print(f"Successfully sent interview email to {name}")
        else:
            print(f"Failed to send interview email to {name}")
   
    print("\nEmail automation process completed.")


if __name__ == "__main__":
    main()