import smtplib
from email.mime.text import MIMEText

def test_email_connection():
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    email_address = "hemanthgutta57@gmail.com"
    password = "pjhb yjrx jcvp dbie"  # Your current app password
    
    try:
        print("Testing email connection...")
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        
        print("Attempting login...")
        server.login(email_address, password)
        print("✅ Login successful!")
        
        # Send test email to yourself
        msg = MIMEText("This is a test email from your bulk email system.")
        msg['Subject'] = "Test Email - Bulk Email System"
        msg['From'] = email_address
        msg['To'] = email_address
        
        server.send_message(msg)
        print("✅ Test email sent successfully!")
        
        server.quit()
        return True
        
    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ Authentication failed: {e}")
        print("Please check your app password")
        return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False

if __name__ == "__main__":
    test_email_connection() 