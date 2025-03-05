import pandas as pd
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os

def send_email(sender_email, sender_password, recipient_email, company_name,receiver_name,pdf_files=None):
    """
    Send an email with HTML content and PDF attachments
    
    Args:
        sender_email (str): Sender's email address
        sender_password (str): Sender's email password or app-specific password
        recipient_email (str): Recipient's email address
        company_name (str): Name of the company to customize email content
        pdf_files (list): List of paths to PDF files to attach
    """
    html_content = f"""
        <p>Dear {recep_name},</p>

    <p>I hope you're doing well. I’m a final-year B.Tech student at reaching out to explore Software Development Engineer (SDE) opportunities at <strong>{company_name} </strong>. With a strong foundation in full-stack development, I am passionate about building scalable and efficient software solutions.</p>

    <h3>Technical Projects:</h3>
    <ul>
        <li><strong>Project 1:</strong> Project detail ......</li>
        <li><strong>Project 2:</strong>Project detail .......</li>
    </ul>

    <h3>Technical Expertise:</h3>
    <ul>
        <li><strong>Programming Languages:</strong> ******</li>
        <li><strong>Frontend:</strong> *****</li>
        <li><strong>Backend:</strong> ******</li>
        <li><strong>Development Tools:</strong> Git, *****</li>
        <li><strong>CS Fundamentals:</strong> ******</li>
    </ul>

    
    """

    msg = MIMEMultipart('alternative')
    msg['Subject'] = f"Inquiry About Software Development Engineer (SDE) Role"
    msg['From'] = sender_email
    msg['To'] = recipient_email

    # Attach HTML content
    msg.attach(MIMEText(html_content, 'html'))

    # Attach PDF files
    if pdf_files:
        for pdf_path in pdf_files:
            try:
                with open(pdf_path, 'rb') as f:
                    pdf_attachment = MIMEApplication(f.read(), _subtype='pdf')
                    pdf_attachment.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=os.path.basename(pdf_path)
                    )
                    msg.attach(pdf_attachment)
            except Exception as e:
                print(f"Error attaching file {pdf_path}: {e}")
                continue

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error sending email: {e}")

# Usage example
if __name__ == "__main__":
    # Replace these with your actual credentials
    sender_email = "yourEmailId@gmail.com"
    sender_password = "**** **** **** ****"  # Use App Password if using Gmail Just search goggle App password from their you will get 16 characters key
    pdf_files = [
        "Resume.pdf",
    ]
    # send_email(sender_email, sender_password, "xyz@gmail.com", "XYZ.AI",receiver_name="XYZ",pdf_files=pdf_files)
    df = pd.read_excel('input.xlsx')
    companies = list(df['Company'])
    receiver_name = list(df['Name'])
    emails = list(df['Email'])
    print(len(companies),len(emails))
    for i in range(len(companies)):
        company_name = companies[i]
        recep_emails = emails[i]
        recep_name = receiver_name[i]
        recep_emails = str(recep_emails).split(',')
        for recipient_email in recep_emails:
            print(company_name,recipient_email)
            send_email(sender_email, sender_password, recipient_email, company_name,receiver_name=receiver_name,pdf_files=pdf_files)
            time.sleep(5)
    # # send_email(sender_email, sender_password, recipient_email, company_name,pdf_files=pdf_files)
    
