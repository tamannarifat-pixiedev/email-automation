"""
Email Scraper Module
Scrapes emails from Gmail using IMAP
"""

import imaplib
import email
from email.header import decode_header
import pandas as pd
from datetime import datetime, timedelta
import re

class EmailScraper:
    def __init__(self, email_address, password):
        self.email = email_address
        self.password = password
        self.imap = None
    
    def connect(self):
        """Gmail এর সাথে connect করে"""
        try:
            self.imap = imaplib.IMAP4_SSL("imap.gmail.com")
            self.imap.login(self.email, self.password)
            return True, "Connected successfully!"
        except Exception as e:
            return False, f"Connection failed: {str(e)}"
    
    def get_folders(self):
        """সব email folders list করে"""
        try:
            status, folders = self.imap.list()
            folder_list = []
            for folder in folders:
                folder_name = folder.decode().split('"')[-2]
                folder_list.append(folder_name)
            return folder_list
        except Exception as e:
            return ["INBOX"]
    
    def scrape_emails(self, folder="INBOX", days=7, keyword=None, 
                     sender=None, max_emails=100):
        """
        Emails scrape করে
        
        Parameters:
        - folder: যে folder থেকে scrape করবে (INBOX, Sent, etc.)
        - days: কত দিনের emails নিবে
        - keyword: Subject/body তে search করার জন্য
        - sender: Specific sender থেকে
        - max_emails: Maximum কত emails নিবে
        """
        try:
            self.imap.select(folder)
            
            # Search criteria build করা
            search_criteria = []
            
            # Date filter
            if days:
                date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
                search_criteria.append(f'SINCE {date}')
            
            # Keyword filter
            if keyword:
                search_criteria.append(f'SUBJECT "{keyword}"')
            
            # Sender filter
            if sender:
                search_criteria.append(f'FROM "{sender}"')
            
            # Search করা
            search_string = ' '.join(search_criteria) if search_criteria else 'ALL'
            status, messages = self.imap.search(None, search_string)
            
            email_ids = messages[0].split()
            
            # Limit emails
            email_ids = email_ids[-max_emails:]
            
            emails_data = []
            
            for email_id in email_ids:
                try:
                    # Email fetch করা
                    status, msg_data = self.imap.fetch(email_id, "(RFC822)")
                    
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            
                            # Subject decode করা
                            subject = self.decode_subject(msg["Subject"])
                            
                            # From address
                            from_ = msg.get("From")
                            
                            # Date
                            date = msg.get("Date")
                            
                            # Body extract করা
                            body = self.get_email_body(msg)
                            
                            # Attachments check করা
                            attachments = self.get_attachments(msg)
                            
                            emails_data.append({
                                "ID": email_id.decode(),
                                "From": from_,
                                "Subject": subject,
                                "Date": date,
                                "Body": body[:200] + "..." if len(body) > 200 else body,
                                "Attachments": ", ".join(attachments) if attachments else "None"
                            })
                
                except Exception as e:
                    print(f"Error processing email {email_id}: {e}")
                    continue
            
            return pd.DataFrame(emails_data)
        
        except Exception as e:
            print(f"Scraping error: {e}")
            return pd.DataFrame()
    
    def decode_subject(self, subject):
        """Subject decode করে"""
        if subject is None:
            return "No Subject"
        
        decoded_parts = decode_header(subject)
        subject_parts = []
        
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                try:
                    subject_parts.append(part.decode(encoding or 'utf-8'))
                except:
                    subject_parts.append(part.decode('utf-8', errors='ignore'))
            else:
                subject_parts.append(str(part))
        
        return ''.join(subject_parts)
    
    def get_email_body(self, msg):
        """Email body extract করে"""
        body = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    try:
                        body = part.get_payload(decode=True).decode()
                        break
                    except:
                        pass
        else:
            try:
                body = msg.get_payload(decode=True).decode()
            except:
                body = str(msg.get_payload())
        
        # Clean body
        body = re.sub(r'\s+', ' ', body).strip()
        return body
    
    def get_attachments(self, msg):
        """Attachment names list করে"""
        attachments = []
        
        for part in msg.walk():
            if part.get_content_disposition() == 'attachment':
                filename = part.get_filename()
                if filename:
                    attachments.append(filename)
        
        return attachments
    
    def disconnect(self):
        """Connection close করে"""
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
            except:
                pass