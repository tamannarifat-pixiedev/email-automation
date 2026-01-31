"""
Email Sender Module
Sends emails using Gmail SMTP
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import os

class EmailSender:
    def __init__(self, email_address, password):
        self.email = email_address
        self.password = password
        self.smtp = None
    
    def connect(self):
        """Gmail SMTP এর সাথে connect করে"""
        try:
            self.smtp = smtplib.SMTP("smtp.gmail.com", 587)
            self.smtp.starttls()
            self.smtp.login(self.email, self.password)
            return True, "Connected successfully!"
        except Exception as e:
            return False, f"Connection failed: {str(e)}"
    
    def send_email(self, to_email, subject, body, attachments=None, 
                   is_html=False):
        """
        একটা email send করে
        
        Parameters:
        - to_email: Recipient email
        - subject: Email subject
        - body: Email body (text or HTML)
        - attachments: List of file paths to attach
        - is_html: Whether body is HTML
        """
        try:
            # Message create করা
            msg = MIMEMultipart()
            msg["From"] = self.email
            msg["To"] = to_email
            msg["Subject"] = subject
            
            # Body attach করা
            if is_html:
                msg.attach(MIMEText(body, "html"))
            else:
                msg.attach(MIMEText(body, "plain"))
            
            # Attachments যোগ করা
            if attachments:
                for file_path in attachments:
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as attachment:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment.read())
                        
                        encoders.encode_base64(part)
                        
                        filename = os.path.basename(file_path)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename= {filename}",
                        )
                        
                        msg.attach(part)
            
            # Send করা
            self.smtp.send_message(msg)
            return True, "Email sent successfully!"
        
        except Exception as e:
            return False, f"Failed to send: {str(e)}"
    
    def send_bulk_emails(self, recipients_df, subject_template, 
                        body_template, attachments=None):
        """
        Multiple emails send করে (personalized)
        
        Parameters:
        - recipients_df: DataFrame with columns: name, email, and other custom fields
        - subject_template: Subject with placeholders like {name}
        - body_template: Body with placeholders
        - attachments: List of files to attach to all emails
        """
        results = []
        
        for index, row in recipients_df.iterrows():
            try:
                # Personalize subject & body
                subject = subject_template
                body = body_template
                
                # Replace placeholders
                for col in recipients_df.columns:
                    placeholder = "{" + col + "}"
                    subject = subject.replace(placeholder, str(row[col]))
                    body = body.replace(placeholder, str(row[col]))
                
                # Send email
                success, message = self.send_email(
                    row["email"],
                    subject,
                    body,
                    attachments
                )
                
                results.append({
                    "Email": row["email"],
                    "Name": row.get("name", "N/A"),
                    "Status": "Sent" if success else "Failed",
                    "Message": message
                })
            
            except Exception as e:
                results.append({
                    "Email": row.get("email", "Unknown"),
                    "Name": row.get("name", "N/A"),
                    "Status": "Failed",
                    "Message": str(e)
                })
        
        return pd.DataFrame(results)
    
    def disconnect(self):
        """Connection close করে"""
        if self.smtp:
            try:
                self.smtp.quit()
            except:
                pass