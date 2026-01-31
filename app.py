"""
Email Automation Dashboard
Complete email scraping & sending system
"""

import streamlit as st
import pandas as pd
from modules.email_scraper import EmailScraper
from modules.email_sender import EmailSender
import os
from dotenv import load_dotenv
from datetime import datetime

# Load environment variables
load_dotenv()

# Page config
st.set_page_config(
    page_title="Email Automation System",
    page_icon="📧",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stat-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f0f2f6;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<p class="main-header">📧 Email Automation System</p>', 
            unsafe_allow_html=True)

# Sidebar - Credentials
st.sidebar.title("⚙️ Configuration")

# Auto-fill from .env if available
default_email = os.getenv("GMAIL_ADDRESS", "")
default_password = os.getenv("GMAIL_APP_PASSWORD", "")

email_address = st.sidebar.text_input(
    "Gmail Address:",
    value=default_email,
    help="Your Gmail address"
)

app_password = st.sidebar.text_input(
    "App Password:",
    type="password",
    value=default_password,
    help="16-digit Gmail App Password (not regular password)"
)

st.sidebar.markdown("---")
st.sidebar.markdown("""
### 📌 Setup Guide:
1. Go to [Google Account](https://myaccount.google.com)
2. Security → 2-Step Verification
3. App Passwords → Generate
4. Use 16-digit password here
""")

# Main navigation
page = st.sidebar.radio(
    "Navigate to:",
    ["📥 Email Scraper", "📤 Email Sender", "📊 Dashboard"]
)

# =================== EMAIL SCRAPER PAGE ===================
if page == "📥 Email Scraper":
    st.title("📥 Email Scraper")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("Scraping Options")
        
        folder = st.selectbox(
            "Select Folder:",
            ["INBOX", "Sent", "[Gmail]/All Mail", "[Gmail]/Spam", "[Gmail]/Trash"],
            help="Which folder to scrape from"
        )
        
        col_a, col_b = st.columns(2)
        
        with col_a:
            days = st.number_input(
                "Last N days:",
                min_value=1,
                max_value=365,
                value=7,
                help="How many days back to scrape"
            )
        
        with col_b:
            max_emails = st.number_input(
                "Max emails:",
                min_value=10,
                max_value=1000,
                value=100,
                step=10
            )
        
        keyword = st.text_input(
            "Keyword (optional):",
            placeholder="Search in subject/body",
            help="Leave empty to get all emails"
        )
        
        sender = st.text_input(
            "From (optional):",
            placeholder="sender@example.com",
            help="Filter by sender email"
        )
    
    with col2:
        st.subheader("Quick Stats")
        if 'scraped_df' in st.session_state:
            df = st.session_state.scraped_df
            st.metric("Total Emails", len(df))
            st.metric("Unique Senders", df['From'].nunique() if len(df) > 0 else 0)
            st.metric("With Attachments", 
                     len(df[df['Attachments'] != 'None']) if len(df) > 0 else 0)
    
    st.markdown("---")
    
    # Scrape button
    if st.button("🔍 Start Scraping", type="primary", use_container_width=True):
        if not email_address or not app_password:
            st.error("⚠️ Please enter Gmail credentials in sidebar!")
        else:
            scraper = EmailScraper(email_address, app_password)
            
            with st.spinner("Connecting to Gmail..."):
                success, message = scraper.connect()
            
            if not success:
                st.error(f"❌ {message}")
                st.info("💡 Make sure you're using App Password, not regular password!")
            else:
                st.success(f"✅ {message}")
                
                with st.spinner(f"Scraping emails from {folder}..."):
                    df = scraper.scrape_emails(
                        folder=folder,
                        days=days,
                        keyword=keyword,
                        sender=sender,
                        max_emails=max_emails
                    )
                
                scraper.disconnect()
                
                if len(df) == 0:
                    st.warning("No emails found with the given criteria.")
                else:
                    st.success(f"✅ Found {len(df)} emails!")
                    
                    # Store in session state
                    st.session_state.scraped_df = df
                    
                    # Display results
                    st.subheader("📋 Scraped Emails")
                    st.dataframe(df, use_container_width=True)
                    
                    # Download options
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        csv = df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            "📥 Download CSV",
                            csv,
                            f"emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                            "text/csv",
                            use_container_width=True
                        )
                    
                    with col2:
                        # Excel download
                        excel_file = f"emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        df.to_excel(excel_file, index=False)
                        with open(excel_file, "rb") as file:
                            st.download_button(
                                "📥 Download Excel",
                                file,
                                excel_file,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        os.remove(excel_file)

# =================== EMAIL SENDER PAGE ===================
elif page == "📤 Email Sender":
    st.title("📤 Email Sender")
    
    # Tabs for single vs bulk
    tab1, tab2 = st.tabs(["📧 Single Email", "📬 Bulk Emails"])
    
    # ===== SINGLE EMAIL TAB =====
    with tab1:
        st.subheader("Send Single Email")
        
        to_email = st.text_input("To:", placeholder="recipient@example.com")
        subject = st.text_input("Subject:", placeholder="Email subject")
        
        body = st.text_area(
            "Message:",
            height=200,
            placeholder="Type your message here..."
        )
        
        is_html = st.checkbox("Send as HTML")
        
        attachments = st.file_uploader(
            "Attachments (optional):",
            accept_multiple_files=True
        )
        
        if st.button("📨 Send Email", type="primary"):
            if not email_address or not app_password:
                st.error("⚠️ Please enter Gmail credentials in sidebar!")
            elif not to_email or not subject or not body:
                st.warning("⚠️ Please fill all fields!")
            else:
                sender = EmailSender(email_address, app_password)
                
                with st.spinner("Connecting to Gmail..."):
                    success, message = sender.connect()
                
                if not success:
                    st.error(f"❌ {message}")
                else:
                    # Save attachments temporarily
                    attachment_paths = []
                    if attachments:
                        for uploaded_file in attachments:
                            file_path = f"temp_{uploaded_file.name}"
                            with open(file_path, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            attachment_paths.append(file_path)
                    
                    with st.spinner("Sending email..."):
                        success, message = sender.send_email(
                            to_email,
                            subject,
                            body,
                            attachment_paths if attachment_paths else None,
                            is_html
                        )
                    
                    sender.disconnect()
                    
                    # Clean up temp files
                    for file_path in attachment_paths:
                        os.remove(file_path)
                    
                    if success:
                        st.success(f"✅ {message}")
                        st.balloons()
                    else:
                        st.error(f"❌ {message}")
    
    # ===== BULK EMAIL TAB =====
    with tab2:
        st.subheader("Send Bulk Personalized Emails")
        
        st.info("""
        📋 **CSV Format Required:**
        - Must have columns: `email`, `name`
        - Optional: Add more columns for personalization
        - Use `{column_name}` in subject/body for personalization
        
        **Example:**
        | name | email | company |
        |------|-------|---------|
        | John | john@example.com | ABC Corp |
        """)
        
        recipients_file = st.file_uploader(
            "Upload Recipients (CSV):",
            type=["csv"],
            help="CSV file with email, name, and other columns"
        )
        
        if recipients_file:
            try:
                recipients_df = pd.read_csv(recipients_file)
                
                st.success(f"✅ Loaded {len(recipients_df)} recipients")
                st.dataframe(recipients_df.head(), use_container_width=True)
                
                # Check required columns
                if 'email' not in recipients_df.columns:
                    st.error("❌ CSV must have 'email' column!")
                else:
                    st.markdown("---")
                    
                    subject_template = st.text_input(
                        "Subject Template:",
                        placeholder="Hello {name}! Special offer from {company}",
                        help="Use {column_name} for personalization"
                    )
                    
                    body_template = st.text_area(
                        "Email Template:",
                        height=200,
                        placeholder="Dear {name},\n\nWe have a special offer for {company}...",
                        help="Use {column_name} for personalization"
                    )
                    
                    # Preview
                    if st.checkbox("📄 Preview First Email"):
                        first_row = recipients_df.iloc[0]
                        preview_subject = subject_template
                        preview_body = body_template
                        
                        for col in recipients_df.columns:
                            placeholder = "{" + col + "}"
                            preview_subject = preview_subject.replace(
                                placeholder, str(first_row[col])
                            )
                            preview_body = preview_body.replace(
                                placeholder, str(first_row[col])
                            )
                        
                        st.info("**Preview:**")
                        st.write(f"**To:** {first_row['email']}")
                        st.write(f"**Subject:** {preview_subject}")
                        st.text_area("Body:", preview_body, height=150, disabled=True)
                    
                    if st.button("📨 Send Bulk Emails", type="primary"):
                        if not email_address or not app_password:
                            st.error("⚠️ Please enter Gmail credentials!")
                        elif not subject_template or not body_template:
                            st.warning("⚠️ Please fill subject and body!")
                        else:
                            sender = EmailSender(email_address, app_password)
                            
                            with st.spinner("Connecting to Gmail..."):
                                success, message = sender.connect()
                            
                            if not success:
                                st.error(f"❌ {message}")
                            else:
                                progress_bar = st.progress(0)
                                status_text = st.empty()
                                
                                results = sender.send_bulk_emails(
                                    recipients_df,
                                    subject_template,
                                    body_template
                                )
                                
                                sender.disconnect()
                                
                                progress_bar.progress(100)
                                status_text.success("✅ All emails processed!")
                                
                                # Show results
                                st.subheader("📊 Results")
                                
                                col1, col2, col3 = st.columns(3)
                                col1.metric("Total", len(results))
                                col2.metric("Sent", 
                                           len(results[results['Status'] == 'Sent']))
                                col3.metric("Failed", 
                                           len(results[results['Status'] == 'Failed']))
                                
                                st.dataframe(results, use_container_width=True)
                                
                                # Download results
                                csv = results.to_csv(index=False).encode('utf-8')
                                st.download_button(
                                    "📥 Download Report",
                                    csv,
                                    f"email_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                    "text/csv"
                                )
                
            except Exception as e:
                st.error(f"❌ Error loading CSV: {e}")

# =================== DASHBOARD PAGE ===================
elif page == "📊 Dashboard":
    st.title("📊 Analytics Dashboard")
    
    if 'scraped_df' not in st.session_state or len(st.session_state.scraped_df) == 0:
        st.info("👈 Scrape some emails first to see analytics!")
    else:
        df = st.session_state.scraped_df
        
        # Metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown('<div class="stat-box">', unsafe_allow_html=True)
            st.metric("Total Emails", len(df))
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="stat-box">', unsafe_allow_html=True)
            st.metric("Unique Senders", df['From'].nunique())
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col3:
            st.markdown('<div class="stat-box">', unsafe_allow_html=True)
            st.metric("With Attachments", len(df[df['Attachments'] != 'None']))
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col4:
            st.markdown('<div class="stat-box">', unsafe_allow_html=True)
            avg_length = df['Body'].str.len().mean()
            st.metric("Avg Body Length", f"{int(avg_length)} chars")
            st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Top senders
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📧 Top 10 Senders")
            top_senders = df['From'].value_counts().head(10)
            st.bar_chart(top_senders)
        
        with col2:
            st.subheader("📎 Emails with Attachments")
            attachment_count = {
                'With Attachments': len(df[df['Attachments'] != 'None']),
                'Without Attachments': len(df[df['Attachments'] == 'None'])
            }
            st.bar_chart(attachment_count)
        
        st.markdown("---")
        
        # Recent emails
        st.subheader("📬 Recent Emails")
        st.dataframe(df.head(20), use_container_width=True)

# Footer
st.sidebar.markdown("---")
st.sidebar.markdown("""
### 📘 Features:
- ✅ Email Scraping
- ✅ Bulk Email Sending
- ✅ Personalization
- ✅ Analytics
- ✅ Export to CSV/Excel

**Made with ❤️ using Tamanna**
""")