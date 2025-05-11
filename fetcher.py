import os
import time
import imaplib
import email
import pandas as pd
import re
from email.header import decode_header
from tkinter import ttk, filedialog, messagebox, Toplevel, Text, Scrollbar, RIGHT, Y, END, Menu, Frame, Button, Label
from modules.utils import call_ai, get_credentials

CSV_FILE = "emails.csv"

class FetcherModule:
    tree = None

    def __init__(self, frame):
        self.frame = frame
        self._init_ui()

    def _init_ui(self):
        main_container = Frame(self.frame)
        main_container.pack(pady=5, padx=20, fill="both", expand=True)

        top_frame = Frame(main_container)
        top_frame.pack(pady=5, fill="x")

        Label(top_frame, text="📥 Fetched Emails", font=("Arial", 11, "bold")).pack(side="left")

        buttons_frame = Frame(top_frame)
        buttons_frame.pack(side="right")
        Button(buttons_frame, text="⬇️ Download CSV", command=lambda: self.download_records('csv')).pack(side="left", padx=2)
        Button(buttons_frame, text="⬇️ Download Excel", command=lambda: self.download_records('excel')).pack(side="left", padx=2)
        Button(buttons_frame, text="🗑️ Clear Data", command=self.clear_data).pack(side="left", padx=2)
        Button(buttons_frame, text="🔄 Refresh", command=self.refresh_display).pack(side="left", padx=2)

        tree_frame = Frame(main_container)
        tree_frame.pack(fill="both", expand=True, pady=5)

        self.tree = ttk.Treeview(tree_frame, columns=("From", "Subject", "Summary", "Suggested Response"), show='headings')
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)

        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(column=0, row=0, sticky='nsew')
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')

        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            width = 180 if col != "Suggested Response" else 220
            self.tree.column(col, width=width, minwidth=100)

        FetcherModule.tree = self.tree

        self.tree.bind("<Double-1>", self.show_full_content)
        self.tree.bind("<Button-3>", self.show_context_menu)

        self.status_label = Label(main_container, text="", font=("Arial", 9))
        self.status_label.pack(pady=5, fill="x")

        self.refresh_display()

    def show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        menu = Menu(self.tree, tearoff=0)
        menu.add_command(label="📋 Copy Row", command=lambda: self.copy_row(row_id))
        menu.add_command(label="🗑️ Delete Row", command=lambda: self.delete_row(row_id))
        menu.post(event.x_root, event.y_root)

    def copy_row(self, row_id):
        values = self.tree.item(row_id, "values")
        row_str = "\\t".join(values)
        self.frame.clipboard_clear()
        self.frame.clipboard_append(row_str)
        self.status_label.config(text="Row copied to clipboard", fg="green")

    def delete_row(self, row_id):
        if messagebox.askyesno("Delete", "Delete this email record?"):
            values = self.tree.item(row_id, "values")
            sender = values[0]
            subject = values[1]

            try:
                df = pd.read_csv(CSV_FILE)
                mask = (df['From'] == sender) & (df['Subject'] == subject)
                if any(mask):
                    df = df[~mask]
                    df.to_csv(CSV_FILE, index=False)
                    self.tree.delete(row_id)
                    self.status_label.config(text="Record deleted", fg="green")
                else:
                    self.status_label.config(text="Record not found in CSV", fg="red")
            except Exception as e:
                messagebox.showerror("Delete Failed", str(e))
                self.status_label.config(text=f"Error: {str(e)}", fg="red")

    def clear_data(self):
        if messagebox.askyesno("Confirm", "Clear all fetched emails?"):
            try:
                if os.path.exists(CSV_FILE):
                    os.remove(CSV_FILE)
                for row in self.tree.get_children():
                    self.tree.delete(row)
                self.status_label.config(text="Inbox records cleared", fg="green")
            except Exception as e:
                messagebox.showerror("Clear Failed", str(e))
                self.status_label.config(text=f"Error clearing: {str(e)}", fg="red")

    def refresh_display(self):
        if os.path.exists(CSV_FILE):
            df = pd.read_csv(CSV_FILE)
            for row in self.tree.get_children():
                self.tree.delete(row)
            for _, row in df.iterrows():
                self.tree.insert('', 'end', values=(row['From'], row['Subject'], row['Summary'], row['Suggested Response']))

    @staticmethod
    def run_background_monitoring():
        while True:
            try:
                FetcherModule.fetch_and_process()
                if FetcherModule.tree:
                    for item in FetcherModule.tree.get_children():
                        FetcherModule.tree.delete(item)
                    df = pd.read_csv(CSV_FILE)
                    for _, row in df.iterrows():
                        FetcherModule.tree.insert('', 'end', values=(row['From'], row['Subject'], row['Summary'], row['Suggested Response']))
            except Exception as e:
                print("Monitoring error:", e)
            time.sleep(10)
    def show_full_content(self, event):
        selected = self.tree.focus()
        if not selected:
            return
        values = self.tree.item(selected, "values")
        if not values:
            return
        full_text = f"From: {values[0]}\nSubject: {values[1]}\n\nSummary:\n{values[2]}\n\nSuggested Response:\n{values[3]}"
        self.show_popup("Email Details", full_text)

    def show_popup(self, title, content):
        popup = Toplevel(self.frame)
        popup.title(title)
        popup.geometry("600x400")

        text = Text(popup, wrap="word")
        scrollbar = Scrollbar(popup, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.insert("1.0", content)
        text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

    def copy_all(event=None):
        popup.clipboard_clear()
        popup.clipboard_append(text.get("1.0", END))

        text.bind("<Control-c>", copy_all)
        text.bind("<Button-3>", lambda e: copy_all())
        
    def download_records(self, file_type='csv'):
        if not os.path.exists(CSV_FILE):
            messagebox.showwarning("No Data", "No records to download.")
            return

        filetypes = [("CSV Files", "*.csv")] if file_type == 'csv' else [("Excel Files", "*.xlsx")]
        default_ext = ".csv" if file_type == 'csv' else ".xlsx"
        
        save_path = filedialog.asksaveasfilename(defaultextension=default_ext, filetypes=filetypes)
        if save_path:
            try:
                df = pd.read_csv(CSV_FILE)
                if file_type == 'excel':
                    df.to_excel(save_path, index=False)
                else:
                    df.to_csv(save_path, index=False)
                messagebox.showinfo("Saved", "Records saved successfully.")
            except Exception as e:
                messagebox.showerror("Save Failed", str(e))


    @staticmethod
    def fetch_and_process():
        user, pwd = get_credentials()
        mail = imaplib.IMAP4_SSL("imap.secureserver.net", 993)
        mail.login(user, pwd)
        mail.select("inbox")

        status, messages = mail.search(None, "UNSEEN")
        email_ids = messages[0].split()

        df = pd.read_csv(CSV_FILE) if os.path.exists(CSV_FILE) else pd.DataFrame(columns=["From", "Subject", "Body", "Summary", "Suggested Response"])

        for eid in email_ids:
            _, msg_data = mail.fetch(eid, "(RFC822)")
            for part in msg_data:
                if isinstance(part, tuple) and len(part) == 2:
                    msg = email.message_from_bytes(part[1])
                    subject, enc = decode_header(msg["Subject"])[0]
                    subject = subject.decode(enc or "utf-8") if isinstance(subject, bytes) else subject
                    from_ = msg.get("From")
                    body = FetcherModule.get_email_body(msg)

                    if not from_:
                        continue
                    if user.lower() in from_.lower() or "mailer-daemon" in from_.lower() or "no-reply" in from_.lower():
                        continue
                    match = re.search(r'<(.+?)>', from_)
                    if match:
                        from_ = match.group(1)

                    summary = call_ai(f"Summarize this to be less than 20 words:\n{body}")
                    reply = call_ai(f"Suggest a reply that fits me as an HR representative and make it professional and warm:\n{body}")

                    df.loc[len(df)] = [from_, subject, body, summary, reply]

        df.to_csv(CSV_FILE, index=False)
        mail.logout()

    @staticmethod
    def get_email_body(msg):
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body += part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8")
        else:
            body = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8")
        return body