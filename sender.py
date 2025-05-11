import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Frame, Label, Entry, Text, Button, Scrollbar
import pandas as pd
import time, threading
from modules.utils import get_credentials, format_email_body, send_summary_report, call_ai, send_email_with_smtp, extract_placeholders

class SenderModule:
    def __init__(self, parent):
        self.recipients = []
        self.sent_emails = []
        self.failed_emails = []
        self.kill_process = False
        self.parent = parent
        self.df = None  # Store DataFrame for column access
        self.available_fields = []  # Store column names
        self._init_gui()

    def _init_gui(self):
        # Main container for better centering
        main_container = Frame(self.parent)
        main_container.pack(padx=20, pady=10, fill="both", expand=True)
        
        ttk.Label(main_container, text="📨 Subject:", font=("Arial", 10, "bold")).pack(pady=2)
        self.subject_input = ttk.Entry(main_container, width=70)
        self.subject_input.pack()

        ttk.Label(main_container, text="📝 Email Body:", font=("Arial", 10, "bold")).pack(pady=2)
        
        # Container for body input with scrollbar
        body_frame = Frame(main_container)
        body_frame.pack(fill="both", expand=True, pady=2)
        self.body_input = Text(body_frame, height=8, width=70)
        body_scrollbar = Scrollbar(body_frame, command=self.body_input.yview)
        self.body_input.configure(yscrollcommand=body_scrollbar.set)
        self.body_input.pack(side=tk.LEFT, fill="both", expand=True)
        body_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Available fields display
        self.fields_frame = Frame(main_container)
        self.fields_frame.pack(pady=2)
        self.fields_label = Label(self.fields_frame, text="Available fields: None", font=("Arial", 8), fg="gray")
        self.fields_label.pack()

        ttk.Label(main_container, text="🎨 Tone / Style:", font=("Arial", 10, "bold")).pack(pady=2)
        self.tone_input = Text(main_container, height=2, width=70)
        self.tone_input.pack()

        self.set_placeholder(self.body_input, "Example: Hello {name}, we're excited to connect...")
        self.set_placeholder(self.tone_input, "Example: professional, friendly, concise")

        enhance_btn = ttk.Button(main_container, text="🧠 Enhance with AI", command=self.enhance_body)
        enhance_btn.pack(pady=5)

        delay_frame = Frame(main_container)
        delay_frame.pack(pady=2)
        ttk.Label(delay_frame, text="⏱ Delay (sec):").pack(side=tk.LEFT, padx=5)
        self.delay_input = ttk.Entry(delay_frame, width=10)
        self.delay_input.pack(side=tk.LEFT)
        self.delay_input.insert(0, "2")  # Default delay

        # File buttons frame
        file_frame = Frame(main_container)
        file_frame.pack(pady=5)
        ttk.Button(file_frame, text="📂 Load CSV", command=lambda: self.load_recipients('csv')).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="📊 Load Excel", command=lambda: self.load_recipients('excel')).pack(side=tk.LEFT, padx=5)

        # Recipients preview tree
        self.tree = ttk.Treeview(main_container, show="headings", height=6)
        tree_scroll = ttk.Scrollbar(main_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill="both", expand=True, pady=5)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

        self.progress_bar = ttk.Progressbar(main_container, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.progress_bar.pack(pady=5, fill=tk.X)

        # Action buttons frame
        btn_frame = Frame(main_container)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="🚀 Send Emails", command=self.send_email).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="🔄 Retry Failed", command=self.retry_failed).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="❌ Kill", command=self.stop_sending).pack(side=tk.LEFT, padx=10)

        # Status label
        self.status_label = Label(main_container, text="", font=("Arial", 9))
        self.status_label.pack(pady=5)

    def set_placeholder(self, text_widget, placeholder):
        def on_focus(event):
            if text_widget.get("1.0", tk.END).strip() == placeholder:
                text_widget.delete("1.0", tk.END)
        def on_blur(event):
            if not text_widget.get("1.0", tk.END).strip():
                text_widget.insert("1.0", placeholder)
        text_widget.insert("1.0", placeholder)
        text_widget.bind("<FocusIn>", on_focus)
        text_widget.bind("<FocusOut>", on_blur)

    def enhance_body(self):
        tone = self.tone_input.get("1.0", tk.END).strip()
        body = self.body_input.get("1.0", tk.END).strip()
        subject = self.subject_input.get().strip()

        prompt = f"""Please enhance the following email for clarity, tone, and effectiveness.
Keep all placeholders unchanged (like {{name}}, {{company}}, etc.).
Tone/style: {tone or 'neutral'}.

Subject:
{subject}

Body:
{body}
"""
        try:
            result = call_ai(prompt)
            if "Subject:" in result and "Body:" in result:
                subj = result.split("Subject:")[1].split("Body:")[0].strip()
                bod = result.split("Body:")[1].strip()
                self.subject_input.delete(0, tk.END)
                self.subject_input.insert(0, subj)
                self.body_input.delete("1.0", tk.END)
                self.body_input.insert("1.0", bod)
            else:
                self.body_input.delete("1.0", tk.END)
                self.body_input.insert("1.0", result)
            messagebox.showinfo("Enhanced", "Email enhanced by AI.")
        except Exception as e:
            messagebox.showerror("AI Error", str(e))

    def update_available_fields(self):
        """Update the available fields display based on loaded data"""
        if hasattr(self, 'available_fields') and self.available_fields:
            field_text = "Available fields: " + ", ".join([f"{{{field}}}" for field in self.available_fields])
            self.fields_label.config(text=field_text)
        else:
            self.fields_label.config(text="Available fields: None")

    def load_recipients(self, file_type):
        """Load recipients from CSV or Excel file"""
        file_types = []
        if file_type == 'csv':
            file_types = [("CSV", "*.csv")]
        elif file_type == 'excel':
            file_types = [("Excel", "*.xlsx"), ("Excel Legacy", "*.xls")]
        else:
            file_types = [("CSV", "*.csv"), ("Excel", "*.xlsx"), ("Excel Legacy", "*.xls")]
            
        file_path = filedialog.askopenfilename(filetypes=file_types)
        if not file_path:
            return
            
        try:
            if file_path.lower().endswith('.csv'):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)
                
            # Convert DataFrame to dict records for consistency
            self.recipients = self.df.to_dict('records')
            self.progress_bar['maximum'] = len(self.recipients)
            
            # Update tree with columns from file
            for row in self.tree.get_children():
                self.tree.delete(row)
                
            # Configure columns dynamically
            columns = list(self.df.columns)
            self.available_fields = columns
            self.update_available_fields()
            
            self.tree["columns"] = columns
            for col in columns:
                self.tree.heading(col, text=col)
                # Calculate width based on column name length
                width = max(100, min(200, len(col) * 10))
                self.tree.column(col, width=width)
            
            # Insert data
            for i, rec in enumerate(self.recipients[:50]):  # Show first 50 records
                values = [rec.get(col, '') for col in columns]
                self.tree.insert('', 'end', values=values)
                
            total = len(self.recipients)
            shown = min(50, total)
            status = f"Loaded {total} recipients{' (showing first 50)' if total > 50 else ''}."
            self.status_label.config(text=status, fg="green")
            messagebox.showinfo("Loaded", status)
            
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def stop_sending(self):
        self.kill_process = True
        self.status_label.config(text="Stopping process...", fg="orange")

    def send_email(self):
        def run():
            try:
                subject = self.subject_input.get()
                body = self.body_input.get("1.0", tk.END).strip()
                delay = int(self.delay_input.get()) if self.delay_input.get().isdigit() else 0
                
                # Validation
                if not self.recipients:
                    messagebox.showwarning("Warning", "No recipients loaded.")
                    return
                if not subject or not body:
                    messagebox.showwarning("Warning", "Subject or body empty.")
                    return
                
                # Reset tracking variables
                self.sent_emails = []
                self.failed_emails = []
                self.kill_process = False
                
                # Extract placeholders from template for validation
                placeholders = extract_placeholders(body)
                subject_placeholders = extract_placeholders(subject)
                all_placeholders = set(placeholders + subject_placeholders)
                
                # Validate all required fields are in data
                if all_placeholders:
                    missing_fields = [field for field in all_placeholders if field not in self.available_fields]
                    if missing_fields:
                        messagebox.showwarning("Warning", f"Missing fields in data: {', '.join(missing_fields)}")
                        return

                for idx, contact in enumerate(self.recipients, 1):
                    if self.kill_process:
                        self.status_label.config(text="Process stopped.", fg="orange")
                        break
                        
                    self.status_label.config(text=f"Sending to {contact.get('email', 'Unknown')}...", fg="blue")
                    
                    try:
                        # Format email with dynamic fields
                        text, html = format_email_body(body, contact)
                        formatted_subject = subject
                        for key, value in contact.items():
                            placeholder = "{" + key + "}"
                            formatted_subject = formatted_subject.replace(placeholder, str(value))
                        
                        # Send using SMTP
                        response = send_email_with_smtp(contact["email"], formatted_subject, text, html)
                        
                        # Empty response dict means successful delivery
                        if not response:
                            self.sent_emails.append(contact["email"])
                        else:
                            self.failed_emails.append(contact["email"])
                            print(f"SMTP Error: {response}")
                            
                    except Exception as e:
                        self.failed_emails.append(contact["email"])
                        print(f"Error: {e}")
                        
                    self.progress_bar['value'] = idx
                    self.parent.update_idletasks()
                    time.sleep(delay)

                send_summary_report(self.sent_emails, self.failed_emails)
                final_msg = f"Completed: Sent {len(self.sent_emails)}, Failed {len(self.failed_emails)}"
                self.status_label.config(text=final_msg, fg="green")
                messagebox.showinfo("Done", final_msg)
                
            except Exception as e:
                self.status_label.config(text=f"Error: {str(e)}", fg="red")
                messagebox.showerror("Send Failed", str(e))

        threading.Thread(target=run).start()
        
    def retry_failed(self):
        """Retry sending emails to failed recipients"""
        if not self.failed_emails:
            messagebox.showinfo("Retry", "No failed emails to retry.")
            return
            
        retry_count = len(self.failed_emails)
        
        def run():
            subject = self.subject_input.get()
            body = self.body_input.get("1.0", tk.END).strip()
            delay = int(self.delay_input.get()) if self.delay_input.get().isdigit() else 0
            
            self.kill_process = False
            retry_success = []
            
            # Reset progress bar for retry
            self.progress_bar['maximum'] = retry_count
            self.progress_bar['value'] = 0
            
            for idx, email in enumerate(self.failed_emails[:], 1):
                if self.kill_process:
                    break
                    
                self.status_label.config(text=f"Retrying {email}...", fg="blue")
                
                # Find the contact info for this email
                contact = next((c for c in self.recipients if c.get("email") == email), None)
                if not contact:
                    continue
                    
                try:
                    # Format email
                    text, html = format_email_body(body, contact)
                    formatted_subject = subject
                    for key, value in contact.items():
                        placeholder = "{" + key + "}"
                        formatted_subject = formatted_subject.replace(placeholder, str(value))
                    
                    # Send using SMTP
                    response = send_email_with_smtp(email, formatted_subject, text, html)
                    
                    if not response:  # Success
                        retry_success.append(email)
                        
                except Exception as e:
                    print(f"Retry Error for {email}: {e}")
                    
                self.progress_bar['value'] = idx
                self.parent.update_idletasks()
                time.sleep(delay)
            
            # Update failed emails list
            self.failed_emails = [e for e in self.failed_emails if e not in retry_success]
            
            # Send summary report
            send_summary_report(self.sent_emails, self.failed_emails, retry_success)
            
            final_msg = f"Retry completed: {len(retry_success)} of {retry_count} succeeded"
            self.status_label.config(text=final_msg, fg="green")
            messagebox.showinfo("Retry Complete", final_msg)
            
        threading.Thread(target=run).start()