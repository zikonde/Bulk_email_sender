import smtplib
import os
import re
import ssl
import pandas as pd
import imaplib
import time
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


class BulkEmailSender(ctk.CTk):
    def __init__(self):
        ctk.CTk.__init__(self)
        ctk.set_appearance_mode('dark')
        ctk.set_default_color_theme('dark-blue')
        self.title("Bulk Email Sender - by Zikonde")
        self.iconbitmap(r'C:\Users\hp\Desktop\MastersApplication_研究生申请\bulk_email_sender\1\favicon.ico')
        self.frame = tk.Frame(master=self,background='#1A1A1A')
        self.frame.pack(fill="both",expand=True , padx=20 , pady=60)

        # Labels
        ctk.CTkLabel(self.frame, text="Sender Email:").grid(row=0, column=0, sticky=ctk.W)
        ctk.CTkLabel(self.frame, text="host:").grid(row=0, column=2, sticky=ctk.W)
        ctk.CTkLabel(self.frame, text="Password:").grid(row=2, column=0, sticky=ctk.W)
        ctk.CTkLabel(self.frame, text="Subject:").grid(row=6, column=0, sticky=ctk.W)
        ctk.CTkLabel(self.frame, text="Body:").grid(row=8, column=0, sticky=ctk.W)

        # Entry Fields
        self.entry_sender = ctk.CTkEntry(self.frame)
        self.entry_sender.grid(row=0, column=1)
        
        self.selected_option =ctk.StringVar()
        self.host = ctk.CTkOptionMenu(self.frame, variable=self.selected_option, values=( 'gmail.com','qq.com','163.com'),width=50)
        self.host.grid(row=0,column=3)

        self.entry_password = ctk.CTkEntry(self.frame, show='*')
        self.entry_password.grid(row=2, column=1)
        
        self.entry_subject = ctk.CTkEntry(self.frame)
        self.entry_subject.grid(row=6, column=1)

        # Text Widget for Body
        self.text_body = ctk.CTkTextbox(self.frame, height=300, width=250)
        self.text_body.grid(row=8, column=1)

        # UI Elements
        self.label_excel_file = ctk.CTkLabel(self.frame, text="Select Excel File:")
        self.label_excel_file.grid(row=12, column=0, pady=10)

        self.label_selected_excel_file = ctk.CTkLabel(self.frame, text="")
        self.label_selected_excel_file.grid(row=12, column=3)

        self.button_browse = ctk.CTkButton(self.frame, text="Browse", command=self.browse_excel_file)
        self.button_browse.grid(row=12, column=1, pady=10)

        self.label_attachment = ctk.CTkLabel(self.frame, text="Select Attachment:")
        self.label_attachment.grid(row=14, column=0, pady=10)

        self.label_selected_attachment = ctk.CTkLabel(self.frame, text="")
        self.label_selected_attachment.grid(row=14, column=3)

        self.button_attachment = ctk.CTkButton(self.frame, text="Attach File", command=self.attach_file)
        self.button_attachment.grid(row=14, column=1, pady=10)

        self.label_status = ctk.CTkLabel(self.frame, text="")
        self.label_status.grid(row=17, column=0, columnspan=2, pady=10)

        self.button_send_emails = ctk.CTkButton(self.frame, text="Send Emails", command=self.send_emails)
        self.button_send_emails.grid(row=19, column=0, columnspan=2, pady=10)

        # Variables
        
        self.excel_file_path = ""
        self.attachment_path = ""

    def extract_from_brackets(self,body):
        pattern = re.compile(r'\{([^\}]+)\}')
        matches = pattern.findall(body)
        return matches

    def browse_excel_file(self):
        self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.label_selected_excel_file.configure(text=f"{os.path.basename(self.excel_file_path)}")

    def attach_file(self):
        self.attachment_path = filedialog.askopenfilename()
        self.label_selected_attachment.configure(text=f"{os.path.basename(self.attachment_path)}")

    def send_email(self,to_email, subject, body, attachment_path=None):
        sender_email = self.entry_sender.get()
        password = self.entry_password.get()

        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = to_email
        message["Subject"] = subject

        message.attach(MIMEText(body, "plain"))

        # Attach file if selected
        if attachment_path:
            with open(attachment_path, "rb") as attachment:
                file_name = os.path.basename(attachment_path)
                part = MIMEApplication(attachment.read(), Name=file_name)
                part["Content-Disposition"] = f"attachment; filename={attachment_path}"

                message.attach(part)

        # Establishing connection to SMTP server and sending email
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp."+self.selected_option.get(), 465, context=context) as server:           #prt 465 or 587
                server.login(sender_email, password)
                server.sendmail(sender_email, to_email, message.as_string())
            self.label_status.configure(text=f"Emails successfully! sent to {to_email}")
        except Exception as e:
            self.label_status.configure(text=f"Error sending email to {to_email}: {str(e)}")
        else:
            try:
                with imaplib.IMAP4_SSL('imap.'+self.selected_option.get()) as mail:
                    mail.login(sender_email, password)
                    mail.select('sent')  # You might need to adjust the folder name
                    mail.append('sent', None, imaplib.Time2Internaldate(time.time()), message.as_bytes())
                self.label_status.configure(text=f"Email saved to {to_email}")
            except:
                self.label_status.configure(text=f"Error saving email to {to_email}: {str(e)}")

    def send_emails(self):
        # Read emails and other details from Excel file
        try:
            application_subject = self.entry_subject.get()
            body = self.text_body.get('1.0','end-1c')
            parameters = self.extract_from_brackets(body)

            xls = pd.ExcelFile(self.excel_file_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(self.excel_file_path,sheet_name)
                for index, row in df.iterrows():

                    body_parsed = body
                    for x in parameters:
                        body_parsed = body_parsed.replace('{'+f'{x}'+'}',f'{row[x]}')

                    self.send_email(row['Email'], application_subject, body_parsed, self.attachment_path)

            self.label_status.configure(text="Emails sent successfully!")

        except Exception as e:
            self.label_status.configure(text=f"An error occurred: {e}")

if __name__ == "__main__":
    app = BulkEmailSender()
    app.config()
    app.mainloop()