from flask import Flask, request, render_template_string, redirect
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import os
import json
from google.cloud import secretmanager
from urllib.parse import quote
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ssl

app = Flask(__name__)

# Department to calendar mapping
DEPARTMENT_CALENDARS = {
    'Research and Development': 'eseoheaigberadion@gmail.com',
    'Finance': 'eseoheaigberadion@gmail.com',
    'Tech': 'eseoheaigberadion@gmail.com',
    'Engineering': 'eseoheaigberadion@gmail.com',
    'Architecture': 'eseoheaigberadion@gmail.com',
    'Public Relations': 'eseoheaigberadion@gmail.com',
    'Legal': 'eseoheaigberadion@gmail.com',
}

# Department-specific meeting locations
DEPARTMENT_LOCATIONS = {
    'Research and Development': 'McCombs School of Business, 2110 Speedway, Austin, TX 78705, USA',
    'Finance': 'McCombs School of Business, 2110 Speedway, Austin, TX 78705, USA',
    'Tech': 'McCombs School of Business, 2110 Speedway, Austin, TX 78705, USA',
    'Engineering': 'McCombs School of Business, 2110 Speedway, Austin, TX 78705, USA',
    'Architecture': 'McCombs School of Business, 2110 Speedway, Austin, TX 78705, USA',
    'Public Relations': 'McCombs School of Business, 2110 Speedway, Austin, TX 78705, USA',
    'Legal': 'McCombs School of Business, 2110 Speedway, Austin, TX 78705, USA',
}

SCOPES = ["https://www.googleapis.com/auth/calendar"]

def get_secret(secret_id):
    """Retrieve a secret from Google Secret Manager."""
    client = secretmanager.SecretManagerServiceClient()
    project_id = os.environ.get('GOOGLE_CLOUD_PROJECT')
    name = f"projects/{project_id}/secrets/{secret_id}/versions/latest"
    response = client.access_secret_version(request={"name": name})
    return response.payload.data.decode("UTF-8")

def get_calendar_service():
    """Initialize Google Calendar service."""
    credentials_json = get_secret("CALENDAR_SERVICE_ACCOUNT_FILE")
    creds_dict = json.loads(credentials_json)
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=SCOPES
    )
    return build("calendar", "v3", credentials=creds)

def generate_time_slots(date_str):
    """Generate 30-min time slots from 9am to 9pm."""
    slots = []
    date = datetime.strptime(date_str, '%Y-%m-%d')
    
    for hour in range(9, 21):  # 9am to 9pm
        for minute in [0, 30]:
            start = date.replace(hour=hour, minute=minute, second=0, microsecond=0)
            end = start + timedelta(minutes=30)
            slots.append({
                'start': start.isoformat(),
                'end': end.isoformat(),
                'display': start.strftime('%I:%M %p')
            })
    
    return slots

def create_ics_file(name, email, department, start_time, end_time, location):
    """Generate .ics calendar file content."""
    start_dt = datetime.fromisoformat(start_time)
    end_dt = datetime.fromisoformat(end_time)
    
    ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//SBI//Interview Scheduler//EN
BEGIN:VEVENT
UID:{start_time}-{email}@utsbi.org
DTSTAMP:{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}
DTSTART:{start_dt.strftime('%Y%m%dT%H%M%S')}
DTEND:{end_dt.strftime('%Y%m%dT%H%M%S')}
SUMMARY:SBI {department} Interview
DESCRIPTION:Your interview with Sustainable Building Initiative for the {department} department.\\n\\nWe will message you through text beforehand about the exact location, or if the location changes.
LOCATION:{location}
STATUS:CONFIRMED
SEQUENCE:0
BEGIN:VALARM
TRIGGER:-PT30M
ACTION:DISPLAY
DESCRIPTION:Interview reminder
END:VALARM
END:VEVENT
END:VCALENDAR"""
    
    return ics_content

def send_calendar_invite_email(name, email, department, start_time, end_time, ics_content, location):
    """Send email with .ics calendar attachment."""
    try:
        sender_email = get_secret("EMAIL_USER")
        sender_password = get_secret("GOOGLE_PASS")
        
        message = MIMEMultipart()
        message["Subject"] = f"Your SBI {department} Interview"
        message["From"] = sender_email
        message["To"] = email
        
        start_dt = datetime.fromisoformat(start_time)
        
        body = f"""
Hi {name},

Your interview with Sustainable Building Initiative has been confirmed!

📅 Date & Time: {start_dt.strftime('%B %d, %Y at %I:%M %p')} CST
🏢 Department: {department}
📍 Location: {location}

The calendar invite is attached to this email. Simply click on the attachment to add it to your calendar.

We will message you through text beforehand about the exact location, or if the location changes.

We look forward to meeting you!

Best regards,
The SBI Team
"""
        
        message.attach(MIMEText(body, "plain"))
        
        # Attach .ics file
        part = MIMEBase("text", "calendar", method="REQUEST")
        part.set_payload(ics_content.encode('utf-8'))
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="sbi-interview.ics"')
        message.attach(part)
        
        # Send email
        context = ssl.create_default_context()
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, email, message.as_string())
        
        print(f"Calendar invite email sent to {email}")
        return True
        
    except Exception as e:
        print(f"Failed to send email to {email}: {e}")
        return False

@app.route('/')
def booking_page():
    booking_id = request.args.get('id', '')
    name = request.args.get('name', '')
    email = request.args.get('email', '')
    department = request.args.get('dept', 'Tech')
    
    # Default to tomorrow's date
    tomorrow = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
    selected_date = request.args.get('date', tomorrow)
    
    time_slots = generate_time_slots(selected_date)
    
    html = '''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Book Your SBI Interview - {{ department }}</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                max-width: 800px;
                margin: 50px auto;
                padding: 20px;
                background-color: #f5f5f5;
            }
            .container {
                background: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            h1 {
                color: #0066cc;
                margin-bottom: 10px;
            }
            .info {
                background-color: #f0f8ff;
                padding: 15px;
                border-left: 4px solid #0066cc;
                margin-bottom: 20px;
            }
            .date-selector {
                margin: 20px 0;
            }
            .date-selector input {
                padding: 10px;
                font-size: 16px;
                border: 1px solid #ddd;
                border-radius: 5px;
            }
            .time-grid {
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
                gap: 10px;
                margin-top: 20px;
            }
            .time-slot {
                background-color: #0066cc;
                color: white;
                padding: 15px;
                text-align: center;
                border-radius: 5px;
                cursor: pointer;
                border: none;
                font-size: 16px;
                transition: background-color 0.3s;
            }
            .time-slot:hover {
                background-color: #0052a3;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Book Your SBI Interview</h1>
            <div class="info">
                <strong>Name:</strong> {{ name }}<br>
                <strong>Email:</strong> {{ email }}<br>
                <strong>Department:</strong> {{ department }}
            </div>
            
            <div class="date-selector">
                <label for="date"><strong>Select Date:</strong></label><br>
                <input type="date" id="date" value="{{ selected_date }}" 
                       min="{{ tomorrow }}"
                       onchange="window.location.href='?id={{ booking_id }}&name={{ name }}&email={{ email }}&dept={{ department }}&date=' + this.value">
            </div>
            
            <h2>Available Time Slots</h2>
            <div class="time-grid">
                {% for slot in time_slots %}
                <form method="POST" action="/book" style="margin: 0;">
                    <input type="hidden" name="booking_id" value="{{ booking_id }}">
                    <input type="hidden" name="name" value="{{ name }}">
                    <input type="hidden" name="email" value="{{ email }}">
                    <input type="hidden" name="department" value="{{ department }}">
                    <input type="hidden" name="start_time" value="{{ slot.start }}">
                    <input type="hidden" name="end_time" value="{{ slot.end }}">
                    <button type="submit" class="time-slot">{{ slot.display }}</button>
                </form>
                {% endfor %}
            </div>
        </div>
    </body>
    </html>
    '''
    
    from jinja2 import Template
    template = Template(html)
    return template.render(
        booking_id=booking_id,
        name=name,
        email=email,
        department=department,
        selected_date=selected_date,
        tomorrow=tomorrow,
        time_slots=time_slots
    )

@app.route('/book', methods=['POST'])
def create_booking():
    name = request.form.get('name')
    email = request.form.get('email')
    department = request.form.get('department')
    start_time = request.form.get('start_time')
    end_time = request.form.get('end_time')
    
    calendar_id = DEPARTMENT_CALENDARS.get(department, 'eseoheaigberadion@gmail.com')
    location = DEPARTMENT_LOCATIONS.get(department, 'McCombs School of Business, 2110 Speedway, Austin, TX 78705, USA')
    
    try:
        # Create calendar event on YOUR calendar
        service = get_calendar_service()
        
        event = {
            'summary': f'SBI {department} Interview - {name}',
            'description': f'Department: {department}\nCandidate: {name}\nEmail: {email}\n\nWe will message you through text beforehand about the exact location, or if the location changes.',
            'location': location,
            'start': {
                'dateTime': start_time,
                'timeZone': 'America/Chicago',
            },
            'end': {
                'dateTime': end_time,
                'timeZone': 'America/Chicago',
            },
        }
        
        created_event = service.events().insert(
            calendarId=calendar_id,
            body=event
        ).execute()
        
        # Create .ics file and send email to candidate
        ics_content = create_ics_file(name, email, department, start_time, end_time, location)
        email_sent = send_calendar_invite_email(name, email, department, start_time, end_time, ics_content, location)
        
        start_dt = datetime.fromisoformat(start_time)
        
        confirmation_html = f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Interview Booked!</title>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    max-width: 600px;
                    margin: 50px auto;
                    padding: 20px;
                    background-color: #f5f5f5;
                }}
                .container {{
                    background: white;
                    padding: 30px;
                    border-radius: 10px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    text-align: center;
                }}
                h1 {{
                    color: #0066cc;
                }}
                .success {{
                    background-color: #d4edda;
                    border: 1px solid #c3e6cb;
                    color: #155724;
                    padding: 15px;
                    border-radius: 5px;
                    margin: 20px 0;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>✅ Interview Booked!</h1>
                <div class="success">
                    <strong>Your interview has been scheduled!</strong><br><br>
                    <strong>Date & Time:</strong> {start_dt.strftime('%B %d, %Y at %I:%M %p')} CST<br>
                    <strong>Department:</strong> {department}<br>
                    <strong>Location:</strong> {location}<br><br>
                    {'✉️ A calendar invite has been sent to your email!' if email_sent else 'Check your email for confirmation.'}
                </div>
                <p>Check your inbox at <strong>{email}</strong> for the calendar invitation.<br>
                Simply click on the attachment to add it to your calendar!</p>
                <p style="font-size: 14px; color: #666; margin-top: 20px;">
                We will message you through text beforehand about the exact location, or if the location changes.
                </p>
                <p>We look forward to meeting you!</p>
            </div>
        </body>
        </html>
        '''
        
        return confirmation_html
        
    except Exception as e:
        print(f"ERROR creating calendar event: {str(e)}")
        import traceback
        traceback.print_exc()
        return f'<h1>Error booking interview</h1><p>{str(e)}</p>', 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)