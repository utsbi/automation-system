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

# ESEOHE ------------------------------------------------------------
from google.oauth2 import service_account
from googleapiclient.discovery import build
# ESEOHE ------------------------------------------------------------

#--- CONFIG ---
GOOGLE_SHEET_NAME = "SBI General Interest Form (Responses)"
LOGO_FILE = "EmailSignature.gif"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# ESEOHE ------------------------------------------------------------
SCOPES = ["https://www.googleapis.com/auth/calendar"]
# ESEOHE ------------------------------------------------------------


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
    """Retrieve new signups on Google Sheets Database."""
    try:
        # Get credentials from Secret Manager
        credentials_json = get_secret("SERVICE_ACCOUNT_FILE")
        credentials_dict = json.loads(credentials_json)
        
        # Create temporary file
        with open('/tmp/credentials.json', 'w') as f:
            json.dump(credentials_dict, f)
        
        gc = gspread.service_account(filename='/tmp/credentials.json')
        
        # Open the spreadsheet
        sh = gc.open(GOOGLE_SHEET_NAME).sheet1
        
        # Get all records from the sheet and convert to a pandas DataFrame
        records = sh.get_all_records()
        # print(f"Total records found: {len(records)}")
        
        df = pd.DataFrame(records)
        print(f"DataFrame shape: {df.shape}")
        print(f"Column names: {df.columns.tolist()}")
        
        df['original_row_index'] = df.index + 2
        
        # Filter for rows where "Automated Email sent" is empty
        new_signups_df = df[df["Automated Email Sent"] == ''].copy()
        
        return sh, new_signups_df
    
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"Error: Spreadsheet '{GOOGLE_SHEET_NAME}' not found.")
        return None, pd.DataFrame()
    
    except Exception as e:
        print(f"An error occurred while fetching data: {e}")
        return None, pd.DataFrame()


def send_welcome_email(recipient_name, recipient_email, departments_str):
    """Create custom email body based on user input."""
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
                from email.mime.image import MIMEImage
                logo = MIMEImage(logo_data)
                logo.add_header("Content-ID", "<logo>")  # Matches cid:logo in HTML
                message.attach(logo)
                
        # Create SMTP session and send email
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())
                
        print(f"Email sent successfully to {recipient_email}")
        return True
    
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")
        return False


def update_email_sent_status(sheet, row_index):
    """Updates Google Sheets to mark sent."""
    try:
        # Find the column index for "Automated Email Sent"
        headers = sheet.row_values(1)
        email_sent_col = headers.index("Automated Email Sent") + 1
        
        # Update the cell
        sheet.update_cell(row_index, email_sent_col, "Yes")
        print(f"Updated row {row_index} - marked email as sent")
        return True
    
    except Exception as e:
        print(f"Failed to update row {row_index}: {e}")
        return False


def main():    

    # ESEOHE ------------------------------------------------------------
    print("TEST TEST TEST")
    credentials_json = get_secret("CALENDAR_SERVICE_ACCOUNT_FILE")
    creds_dict = json.loads(credentials_json)
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=SCOPES
    )
    service = build("calendar", "v3", credentials=creds)
    calendar_id = "eseoheaigberadion@gmail.com"
    print(f"Calendar service initialized for {calendar_id}")
    # ESEOHE ------------------------------------------------------------

    # Get new signups from Google Sheets
    sheet, new_signups_df = get_new_signups()
    
    if sheet is None or new_signups_df.empty:
        print("No new signups found or error accessing sheet.")
        return
    
    print(f"Found {len(new_signups_df)} new signups to process.")
    
    # Process each new signup
    for index, row in new_signups_df.iterrows():
        name = row.get("What is your name?", "")
        email = row.get("What is your email?", "")
        departments = row.get("Which department(s) do you want to be in? (Pick up to 2)", "")
        original_row = row.get("original_row_index")

        if not email or not name:
            print(f"Skipping row {original_row}: missing name or email")
            continue
        
        print(f"Processing: {name} ({email})")
        
        # Capitalize name of person
        name = name.title()

        # Send email
        if send_welcome_email(name, email, departments):
            # Mark as sent in Google Sheets
            if update_email_sent_status(sheet, original_row):
                print(f"Successfully processed {name}")
            else:
                print(f"Email sent to {name} but failed to update sheet")
        else:
            print(f"Failed to process {name}")

        # ESEOHE ------------------------------------------------------------
        print(f"Attempting to create calendar event for {email}")
        event = {
            "summary": "SBI Interview",
            "start": {"dateTime": "2025-01-15T12:00:00-06:00"},
            "end": {"dateTime": "2025-01-15T13:00:00-06:00"},
            "attendees": [{"email": email}],
        }
        print(f"Event details: {event}")
        created = service.events().insert(
            calendarId=calendar_id,
            body=event,
            sendUpdates="all"
        ).execute()
        print(f"Event created successfully: {created['htmlLink']}")
        # ESEOHE ------------------------------------------------------------

    print("Email automation process completed.")

if __name__ == "__main__":
    main()