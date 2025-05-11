import os
import requests
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

def get_credentials():
    load_dotenv(dotenv_path=".env", override=True)
    email = os.getenv("EMAIL_WORK")
    password = os.getenv("EMAIL_WORK_PASSWORD")
    if not email or not password or "example" in email:
        raise Exception("EMAIL_WORK / EMAIL_WORK_PASSWORD not set properly.")
    return email, password

def call_ai(prompt):
    from modules.settings import use_local_ai
    if use_local_ai():
        import ollama
        return ollama.chat(model="qwen2.5:0.5b", messages=[{"role": "user", "content": prompt}])["message"]["content"]
    else:
        headers = {"Authorization": "Bearer b2IuTN-YqsPpLG60zQRGdt_f-ERKSdS2U1PNVXZFeOk"}
        res = requests.post("https://dev.aimicromind.com/api/v1/prediction/03d297ba-bcd3-4a39-b3b2-7d51cac880ab",
                            headers=headers, json={"question": prompt})
        data = res.json()
        if "text" not in data:
            print("⚠️ AI API error:", data)
            raise Exception("AI server response missing 'text' key.")
        return data["text"]

def format_email_body(template, contact):
    """Format email body with dynamic variables from contact dictionary"""
    try:
        # Replace all placeholders dynamically
        body = template
        for key, value in contact.items():
            placeholder = "{" + key + "}"
            body = body.replace(placeholder, str(value))
        
        # Create HTML version
        html = f"<html><body>{body}</body></html>"
        return body, html
    except KeyError as e:
        # If a placeholder in template isn't found in contact dict
        raise Exception(f"Missing variable in data: {str(e)}")

def send_email_with_smtp(to_email, subject, body_text, body_html=None):
    """Send email using SMTP instead of yagmail"""
    sender_email, sender_password = get_credentials()
    
    # Create a multipart message
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = to_email
    
    # Add text part
    part1 = MIMEText(body_text, "plain")
    message.attach(part1)
    
    # Add HTML part if provided
    if body_html:
        part2 = MIMEText(body_html, "html")
        message.attach(part2)
    
    # Create secure SSL context
    context = ssl.create_default_context()
    
    try:
        with smtplib.SMTP_SSL("smtpout.secureserver.net", 465, context=context) as server:
            server.login(sender_email, sender_password)
            response = server.sendmail(sender_email, to_email, message.as_string())
            # Return empty dict on success or response dict with undelivered addresses
            return response
    except Exception as e:
        raise Exception(f"Failed to send email: {str(e)}")

def send_summary_report(sent, failed, retry_failed=None):
    if not sent and not failed and not retry_failed:
        return
    
    sender_email, _ = get_credentials()
    
    body_text = f"Sent: {len(sent)}\nFailed: {len(failed)}"
    if retry_failed:
        body_text += f"\nSuccessfully retried: {len(retry_failed)}"
    
    send_email_with_smtp(sender_email, "Email Campaign Summary", body_text)

def extract_placeholders(template):
    """Extract all placeholders from a template string"""
    import re
    # Find all strings like {name}, {email}, etc.
    placeholders = re.findall(r'\{([^}]+)\}', template)
    return placeholders