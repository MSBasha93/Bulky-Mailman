import os
import imaplib
import email
from email.header import decode_header
import pandas as pd
import yagmail
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import ollama

# Global variables
recipients = []  # For sending bulk emails
sent_emails = []
failed_emails = []
kill_process = False
csv_file_path = "emails.csv"  # Path to the CSV file for incoming emails

# Function to load recipients from a file (for bulk email sending)
def load_recipients():
    global recipients
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
    if file_path:
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            recipients = df.to_dict('records')
            messagebox.showinfo("Success", f"Loaded {len(recipients)} recipients.")
            progress_bar['maximum'] = len(recipients)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load recipients: {str(e)}")

# Function to stop the process
def stop_sending():
    global kill_process
    kill_process = True

# Function to send bulk emails
def send_email():
    def run():
        global sent_emails, failed_emails, kill_process
        try:
            # Retrieve credentials from environment variables
            email_address = os.getenv("EMAIL_WORK")
            password = os.getenv("EMAIL_WORK_PASSWORD")

            if not email_address or not password:
                raise ValueError("Environment variables EMAIL_WORK and EMAIL_WORK_PASSWORD are not set.")

            subject = subject_input.get()
            body_template = body_input.get("1.0", tk.END).strip()
            delay = int(delay_input.get()) if delay_input.get().isdigit() else 0

            if not recipients:
                messagebox.showwarning("Warning", "No recipients loaded.")
                return

            sent_emails = []
            failed_emails = []
            kill_process = False

            # Initialize yagmail SMTP connection
            yag = yagmail.SMTP(user=email_address, password=password, host='smtpout.secureserver.net', port=465)
            for index, contact in enumerate(recipients, start=1):
                if kill_process:
                    break

                formatted_body = body_template.replace('{name}', contact.get('name', '')) \
                                              .replace('{company}', contact.get('company_name', '')) \
                                              .replace('{info}', contact.get('company_info', ''))

                # HTML version of the email
                formatted_html_body = f"""
                <html>
                <body>
                <p>Hi <b>{contact.get('name', '')}</b>,</p>
                <p>{contact.get('company_info', '')}</p>
                <p>Best regards,<br><b>Your Company</b></p>
                </body>
                </html>
                """

                try:
                    yag.send(
                        to=contact.get('email', ''),
                        subject=subject,
                        contents=[formatted_body, formatted_html_body]
                    )
                    sent_emails.append(f"{contact.get('email', '')} ({contact.get('company_name', '')})")
                    print(f"Email sent to {contact.get('email', '')}")
                except Exception as e:
                    failed_emails.append(f"{contact.get('email', '')} ({contact.get('company_name', '')})")
                    print(f"Failed to send to {contact.get('email', '')}: {str(e)}")

                progress_bar['value'] = index
                window.update_idletasks()
                time.sleep(delay)

            send_summary_report(yag, email_address)
            messagebox.showinfo("Success", "All emails processed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send emails: {str(e)}")

    threading.Thread(target=run).start()

# Function to fetch emails, process with Qwen, and update CSV
def fetch_and_process_emails():
    try:
        # Retrieve credentials from environment variables
        email_address = os.getenv("EMAIL_WORK")
        password = os.getenv("EMAIL_WORK_PASSWORD")

        if not email_address or not password:
            raise ValueError("Environment variables EMAIL_WORK and EMAIL_WORK_PASSWORD are not set.")

        # Connect to the IMAP server
        mail = imaplib.IMAP4_SSL("imap.secureserver.net", 993)
        mail.login(email_address, password)
        mail.select("inbox")

        # Search for unseen emails
        status, messages = mail.search(None, "UNSEEN")
        if status == "OK":
            email_ids = messages[0].split()
            if not email_ids:
                print("No new emails found.")
                return

            # Load existing CSV file or create a new one if it doesn't exist
            if os.path.exists(csv_file_path):
                df = pd.read_csv(csv_file_path)
            else:
                df = pd.DataFrame(columns=["From", "Subject", "Body", "Summary", "Suggested Response"])

            for email_id in email_ids:
                # Fetch the email by ID
                status, msg_data = mail.fetch(email_id, "(RFC822)")
                if status == "OK":
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            subject, encoding = decode_header(msg["Subject"])[0]
                            if isinstance(subject, bytes):
                                subject = subject.decode(encoding or "utf-8")
                            from_ = msg.get("From")
                            body = ""
                            if msg.is_multipart():
                                for part in msg.walk():
                                    if part.get_content_type() == "text/plain":
                                        body += part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8")
                            else:
                                body = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8")

                            # Check for duplicates
                            if from_ in df["From"].values and subject in df["Subject"].values:
                                print(f"Duplicate email detected: {subject}")
                                continue

                            # Generate summary and response using Qwen
                            prompt_summary = f"Summarize the following email:\nSubject: {subject}\nBody: {body}"
                            response_summary = ollama.chat(model="qwen2.5:0.5b", messages=[{"role": "user", "content": prompt_summary}])
                            summary = response_summary["message"]["content"]

                            prompt_response = f"Suggest a polite response to the following email:\nSubject: {subject}\nBody: {body}"
                            response_suggestion = ollama.chat(model="qwen2.5:0.5b", messages=[{"role": "user", "content": prompt_response}])
                            suggested_response = response_suggestion["message"]["content"]

                            # Add email to DataFrame
                            new_email = {
                                "From": from_,
                                "Subject": subject,
                                "Body": body,
                                "Summary": summary,
                                "Suggested Response": suggested_response
                            }
                            df = pd.concat([df, pd.DataFrame([new_email])], ignore_index=True)

            # Save the updated DataFrame to the CSV file
            df.to_csv(csv_file_path, index=False)
            print(f"Updated CSV file with {len(email_ids)} new email(s).")
        mail.logout()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to fetch and process emails: {str(e)}")

# Function to periodically check for new emails and process them
def monitor_inbox():
    while not kill_process:
        fetch_and_process_emails()
        time.sleep(30)  # Check every 30 seconds

# Create the main window
window = tk.Tk()
window.title("Bulk Email Sender and Inbox Processor")
window.geometry("520x850")
window.configure(bg="#eef2f7")

# Sender Frame (for bulk email sending)
sender_frame = tk.Frame(window, bg="#eef2f7")
sender_frame.pack(pady=10)

# Subject
tk.Label(sender_frame, text="Subject:", bg="#eef2f7", font=("Arial", 10, "bold")).pack(pady=2)
subject_input = tk.Entry(sender_frame, width=50)
subject_input.pack(pady=2)

# Email Body
tk.Label(sender_frame, text="Body (use {name}, {company}, {info}):", bg="#eef2f7", font=("Arial", 10, "bold")).pack(pady=2)
body_input = tk.Text(sender_frame, height=10, width=50)
body_input.pack(pady=5)

# Delay between emails
tk.Label(sender_frame, text="Delay between emails (seconds):", bg="#eef2f7", font=("Arial", 10, "bold")).pack(pady=2)
delay_input = tk.Entry(sender_frame, width=50)
delay_input.pack(pady=2)

# Load Recipients Button (for bulk email sending)
tk.Button(window, text="Load Recipients from CSV/Excel", command=load_recipients, bg="#4caf50", fg="white", font=("Arial", 10, "bold")).pack(pady=10)

# Progress Bar
progress_bar = ttk.Progressbar(window, orient=tk.HORIZONTAL, length=400, mode='determinate')
progress_bar.pack(pady=10)

# Buttons Frame
button_frame = tk.Frame(window, bg="#eef2f7")
button_frame.pack(pady=20)

# Send and Kill Buttons (for bulk email sending)
tk.Button(button_frame, text="Send Emails", command=send_email, bg="#2196f3", fg="white", font=("Arial", 10, "bold"), width=15).grid(row=0, column=0, padx=10)
tk.Button(button_frame, text="Kill Process", command=stop_sending, bg="#f44336", fg="white", font=("Arial", 10, "bold"), width=15).grid(row=0, column=1, padx=10)

# Start monitoring inbox for new emails
threading.Thread(target=monitor_inbox, daemon=True).start()

# Run the application
window.mainloop()