import os
from tkinter import ttk, Label, Entry, StringVar, Button, Radiobutton, messagebox, Frame
from dotenv import load_dotenv, set_key
import smtplib
import ssl

ENV_PATH = ".env"

def load_settings_env():
    load_dotenv(ENV_PATH, override=True)

def use_local_ai():
    return os.getenv("AI_MODE", "api") == "local"

def save_settings_to_env(email, password, mode):
    if not os.path.exists(ENV_PATH):
        with open(ENV_PATH, "w") as f:
            f.write("")  # Create empty .env file

    set_key(ENV_PATH, "EMAIL_WORK", email)
    set_key(ENV_PATH, "EMAIL_WORK_PASSWORD", password)
    set_key(ENV_PATH, "AI_MODE", mode)
    load_dotenv(ENV_PATH, override=True)  # Refresh immediately after saving

class SettingsModule:
    def __init__(self, frame):
        self.frame = frame
        self.email_var = StringVar(value=os.getenv("EMAIL_WORK", ""))
        self.pass_var = StringVar(value=os.getenv("EMAIL_WORK_PASSWORD", ""))
        self.ai_mode = StringVar(value=os.getenv("AI_MODE", "api"))
        self._init_ui()

    def _init_ui(self):
        main_container = Frame(self.frame)
        main_container.pack(padx=20, pady=10, fill="both", expand=True)
        
        Label(main_container, text="✉️ Your Email:", font=("Arial", 10, "bold")).pack(pady=5)
        Entry(main_container, textvariable=self.email_var, width=40).pack()

        Label(main_container, text="🔒 Your Password:", font=("Arial", 10, "bold")).pack(pady=5)
        Entry(main_container, textvariable=self.pass_var, width=40, show="*").pack()

        # Button frame for test connection
        test_btn_frame = Frame(main_container)
        test_btn_frame.pack(pady=5)
        Button(test_btn_frame, text="🔄 Test Connection", command=self.test_connection).pack()

        Label(main_container, text="🧠 AI Mode:", font=("Arial", 10, "bold")).pack(pady=10)
        Radiobutton(main_container, text="API (Remote)", variable=self.ai_mode, value="api").pack()
        Radiobutton(main_container, text="Local (Qwen)", variable=self.ai_mode, value="local").pack()

        Button(main_container, text="💾 Save Settings", command=self.save_settings).pack(pady=15)

        self.status_label = Label(main_container, text="", font=("Arial", 10))
        self.status_label.pack()

    def test_connection(self):
        """Test SMTP and IMAP connections with current credentials"""
        email = self.email_var.get()
        password = self.pass_var.get()
        
        if not email or not password:
            self.status_label.config(text="❌ Missing email or password", foreground="red")
            return
            
        # Test SMTP connection
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtpout.secureserver.net", 465, context=context) as server:
                server.login(email, password)
                self.status_label.config(text="✅ SMTP Connection Successful", foreground="green")
                messagebox.showinfo("Connection Test", "SMTP Connection Successful!")
        except Exception as e:
            self.status_label.config(text=f"❌ SMTP Connection Failed", foreground="red")
            messagebox.showerror("Connection Test", f"SMTP Connection Failed: {str(e)}")
            return
            
        # Test IMAP connection
        try:
            import imaplib
            mail = imaplib.IMAP4_SSL("imap.secureserver.net", 993)
            mail.login(email, password)
            mail.logout()
            self.status_label.config(text="✅ Both Connections Successful", foreground="green")
            messagebox.showinfo("Connection Test", "IMAP Connection Successful!")
        except Exception as e:
            self.status_label.config(text=f"❌ IMAP Connection Failed", foreground="red")
            messagebox.showerror("Connection Test", f"IMAP Connection Failed: {str(e)}")
            return

    def save_settings(self):
        email = self.email_var.get()
        password = self.pass_var.get()
        mode = self.ai_mode.get()

        save_settings_to_env(email, password, mode)

        print("✔ .env updated:")
        print(open(ENV_PATH).read())

        self.status_label.config(text="✅ Settings Saved", foreground="green")