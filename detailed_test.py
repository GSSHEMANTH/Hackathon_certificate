import smtplib
from email.mime.text import MIMEText
import ssl

def detailed_email_test():
    smtp_server = "smtp.gmail.com"
    smtp_port = 465  # Changed to SSL
    email_address = "hemanthgutta57@gmail.com"
    password = "pjhb yjrx jcvp dbie"
    
    print("=== DETAILED EMAIL TEST (SSL) ===")
    print(f"Server: {smtp_server}")
    print(f"Port: {smtp_port}")
    print(f"Email: {email_address}")
    print(f"Password: {password[:4]}****{password[-4:]}")
    
    try:
        print("\n1. Creating SSL SMTP connection...")
        context = ssl.create_default_context()
        server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
        print("✅ SSL SMTP connection created")
        
        print("\n2. Attempting login...")
        server.login(email_address, password)
        print("✅ Login successful!")
        
        print("\n3. Sending test email...")
        msg = MIMEText("This is a test email from your bulk email system.")
        msg['Subject'] = "Test Email - Bulk Email System"
        msg['From'] = email_address
        msg['To'] = email_address
        
        server.send_message(msg)
        print("✅ Test email sent successfully!")
        
        server.quit()
        print("\n🎉 All tests passed! Your email configuration is working.")
        return True
        
    except smtplib.SMTPAuthenticationError as e:
        print(f"\n❌ AUTHENTICATION ERROR: {e}")
        return False
    except smtplib.SMTPConnectError as e:
        print(f"\n❌ CONNECTION ERROR: {e}")
        return False
    except Exception as e:
        print(f"\n❌ UNEXPECTED ERROR: {e}")
        print(f"Error type: {type(e).__name__}")
        return False

if __name__ == "__main__":
    detailed_email_test() 