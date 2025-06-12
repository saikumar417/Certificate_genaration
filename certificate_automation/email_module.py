import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def save_email_details(output_excel_path, certificate_details):
    """
    Appends name, email, and certificate path to an Excel file.

    :param output_excel_path: Path to save the email details Excel file.
    :param certificate_details: List of tuples (name, email, certificate_path).
    """
    # Convert the input list to a DataFrame
    new_data = pd.DataFrame(certificate_details, columns=["Name", "Email", "Certificate Path"])

    # Check if file exists
    if os.path.exists(output_excel_path):
        # Read existing data
        existing_data = pd.read_excel(output_excel_path)

        # Append new data (avoid duplicates)
        updated_data = pd.concat([existing_data, new_data]).drop_duplicates()

        # Save back to Excel
        updated_data.to_excel(output_excel_path, index=False)
    else:
        # Create a new Excel file and save data
        new_data.to_excel(output_excel_path, index=False)

    print(f"Email details updated in: {output_excel_path}")


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Sender")
        self.root.geometry("400x350")

        # Email & Password Fields
        tk.Label(root, text="Enter Your Email:").pack()
        self.email_entry = tk.Entry(root, width=40)
        self.email_entry.pack()

        tk.Label(root, text="Enter Your App Password:").pack()
        self.password_entry = tk.Entry(root, width=40, show="*")
        self.password_entry.pack()

        # File Upload Button
        self.file_path = ""
        tk.Button(root, text="Upload Excel File", command=self.upload_file).pack()

        # Display Selected File
        self.file_label = tk.Label(root, text="", fg="green")
        self.file_label.pack()

        # Send Button
        tk.Button(root, text="Send Emails", command=self.send_emails).pack()

    def upload_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.file_label.config(text=f"File: {os.path.basename(self.file_path)} uploaded!")

    def send_emails(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        email = self.email_entry.get()
        password = self.password_entry.get()
        if not email or not password:
            messagebox.showerror("Error", "Please enter email and app password.")
            return

        df = pd.read_excel(self.file_path)
        skipped = []
        sent_count = 0

        for _, row in df.iterrows():
            recipient_email = row.get("Email", "").strip()
            cert_path = row.get("Certificate Path", "").strip()
            name = row.get("Name", "").strip()

            if not recipient_email:
                skipped.append(name)
                continue  # Skip sending email if missing

            try:
                self.send_email(email, password, recipient_email, cert_path, name)
                sent_count += 1
            except smtplib.SMTPAuthenticationError:
                messagebox.showerror("Error", "Authentication failed! Check your email or app password.")
                return
            except Exception as e:
                print(f"Error sending to {recipient_email}: {e}")

        messagebox.showinfo("Completed", f"Emails sent: {sent_count}\nSkipped: {len(skipped)}")
        print("Skipped recipients:", skipped)

    def send_email(self, sender_email, password, recipient_email, cert_path, name):
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "Your Certificate"
        
        with open(cert_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(cert_path)}")
            msg.attach(part)

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print(f"✅ Email sent to {recipient_email}")

def main():
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
