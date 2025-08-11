import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os
from datetime import datetime
import time
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io

class BulkEmailCertificateSystem:
    def __init__(self, excel_file_path, smtp_server, smtp_port, email_address, password):
        """
        Initialize the bulk email certificate system
        
        Args:
            excel_file_path (str): Path to the Excel file containing student details
            smtp_server (str): SMTP server (e.g., 'smtp.gmail.com')
            smtp_port (int): SMTP port (e.g., 587 for TLS)
            email_address (str): Your email address
            password (str): Your email password or app password
        """
        self.excel_file_path = excel_file_path
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email_address = email_address
        self.password = password
        self.certificates_dir = "certificates"
        self.template_dir = "templates"
        
        # Create necessary directories
        os.makedirs(self.certificates_dir, exist_ok=True)
        os.makedirs(self.template_dir, exist_ok=True)
        
    def read_student_data(self):
        """
        Read student details from Excel file
        
        Returns:
            pandas.DataFrame: DataFrame containing student information
        """
        try:
            df = pd.read_excel(self.excel_file_path)
            print(f"Successfully read {len(df)} student records from {self.excel_file_path}")
            return df
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return None
    
    def generate_certificate(self, student_name, course_name, completion_date, certificate_id):
        """
        Generate a certificate PDF for a student
        
        Args:
            student_name (str): Name of the student
            course_name (str): Name of the course
            completion_date (str): Date of completion
            certificate_id (str): Unique certificate ID
            
        Returns:
            str: Path to the generated certificate file
        """
        # Create certificate filename
        filename = f"{self.certificates_dir}/certificate_{certificate_id}_{student_name.replace(' ', '_')}.pdf"
        
        # Create PDF document
        doc = SimpleDocTemplate(filename, pagesize=A4)
        story = []
        
        # Get styles
        styles = getSampleStyleSheet()
        
        # Create custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            spaceAfter=30,
            alignment=TA_CENTER,
            textColor=colors.darkblue
        )
        
        name_style = ParagraphStyle(
            'CustomName',
            parent=styles['Heading2'],
            fontSize=20,
            spaceAfter=20,
            alignment=TA_CENTER,
            textColor=colors.black
        )
        
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['Normal'],
            fontSize=14,
            spaceAfter=12,
            alignment=TA_CENTER,
            textColor=colors.black
        )
        
        # Add certificate content
        story.append(Paragraph("CERTIFICATE OF COMPLETION", title_style))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph("This is to certify that", body_style))
        story.append(Spacer(1, 10))
        
        story.append(Paragraph(f"<b>{student_name}</b>", name_style))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph("has successfully completed the course", body_style))
        story.append(Spacer(1, 10))
        
        story.append(Paragraph(f"<b>{course_name}</b>", name_style))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph(f"on {completion_date}", body_style))
        story.append(Spacer(1, 30))
        
        story.append(Paragraph("Certificate ID: " + certificate_id, body_style))
        story.append(Spacer(1, 40))
        
        # Add signature line
        story.append(Paragraph("_________________________", body_style))
        story.append(Paragraph("Authorized Signature", body_style))
        
        # Build PDF
        doc.build(story)
        
        print(f"Generated certificate for {student_name}: {filename}")
        return filename
    
    def send_email(self, to_email, student_name, certificate_path, course_name):
        """
        Send email with certificate attachment
        
        Args:
            to_email (str): Recipient email address
            student_name (str): Name of the student
            certificate_path (str): Path to the certificate file
            course_name (str): Name of the course
            
        Returns:
            bool: True if email sent successfully, False otherwise
        """
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.email_address
            msg['To'] = to_email
            msg['Subject'] = f"Certificate of Completion - {course_name}"
            
            # Email body
            body = f"""
Dear {student_name},

Congratulations! You have successfully completed the course "{course_name}".

Please find your certificate of completion attached to this email.

Certificate Details:
- Student Name: {student_name}
- Course: {course_name}
- Issue Date: {datetime.now().strftime('%B %d, %Y')}

If you have any questions, please don't hesitate to contact us.

Best regards,
Certificate Team
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach certificate
            with open(certificate_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {os.path.basename(certificate_path)}'
            )
            msg.attach(part)
            
            # Send email using SSL
            import ssl
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, context=context)
            server.login(self.email_address, self.password)
            text = msg.as_string()
            server.sendmail(self.email_address, to_email, text)
            server.quit()
            
            print(f"Email sent successfully to {student_name} ({to_email})")
            return True
            
        except Exception as e:
            print(f"Error sending email to {student_name} ({to_email}): {e}")
            return False
    
    def add_name_to_certificate(self, student_name, certificate_id):
        """
        Add student name to the original PDF certificate
        """
        # Use your new certificate file
        original_certificate = r"C:\Users\heman\Downloads\ML Hackathon - Participation Certificate 1.pdf"
        
        # Create new filename for this student
        new_filename = f"{self.certificates_dir}/certificate_{certificate_id}_{student_name.replace(' ', '_')}.pdf"
        
        try:
            # Read the original PDF
            reader = PdfReader(original_certificate)
            writer = PdfWriter()
            page = reader.pages[0]
            
            # Create overlay with name
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            
            # Position for the student name (same coordinates as before)
            x_position = 200
            y_position = 480
            
            # Reduced font size from 20 to 16 (smaller)
            can.setFont("Helvetica-Bold", 16)
            can.setFillColorRGB(0, 0, 0)
            can.drawString(x_position, y_position, student_name)
            
            can.save()
            packet.seek(0)
            new_pdf = PdfReader(packet)
            
            # Merge overlay with original certificate
            page.merge_page(new_pdf.pages[0])
            writer.add_page(page)
            
            # Save the result
            with open(new_filename, "wb") as output_file:
                writer.write(output_file)
            
            print(f"Generated certificate for {student_name}: {new_filename}")
            return new_filename
            
        except Exception as e:
            print(f"Error adding name to certificate for {student_name}: {e}")
            return None
    
    def process_all_students(self):
        """
        Process all students: add names to certificates and send emails
        """
        # Read student data
        df = self.read_student_data()
        if df is None:
            return
        
        # Check required columns
        required_columns = ['Name', 'Email', 'Course']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Missing required columns: {missing_columns}")
            print("Please ensure your Excel file has columns: Name, Email, Course")
            return
        
        # Check if your new certificate exists
        original_certificate = r"C:\Users\heman\Downloads\ML Hackathon - Participation Certificate 1.pdf"
        if not os.path.exists(original_certificate):
            print(f"Certificate file not found: {original_certificate}")
            print("Please check the path to your PDF certificate")
            return
        
        # Process each student
        successful_emails = 0
        total_students = len(df)
        
        for index, row in df.iterrows():
            student_name = row['Name']
            email = row['Email']
            course_name = row['Course']
            
            # Generate certificate ID
            certificate_id = f"CERT-{datetime.now().strftime('%Y%m%d')}-{index+1:03d}"
            
            print(f"\nProcessing student {index+1}/{total_students}: {student_name}")
            
            # Add name to certificate
            certificate_path = self.add_name_to_certificate(student_name, certificate_id)
            
            if certificate_path:
                # Send email
                if self.send_email(email, student_name, certificate_path, course_name):
                    successful_emails += 1
            else:
                print(f"Failed to generate certificate for {student_name}")
            
            # Add delay to avoid overwhelming the email server
            time.sleep(2)
        
        print(f"\n=== PROCESSING COMPLETE ===")
        print(f"Total students processed: {total_students}")
        print(f"Successful emails sent: {successful_emails}")
        print(f"Failed emails: {total_students - successful_emails}")
        print(f"Original certificate: {original_certificate}")

def create_sample_excel():
    """
    Create a sample Excel file with student data
    """
    sample_data = {
        'Name': ['John Doe', 'Jane Smith', 'Mike Johnson', 'Sarah Wilson', 'David Brown'],
        'Email': ['john.doe@email.com', 'jane.smith@email.com', 'mike.johnson@email.com', 
                 'sarah.wilson@email.com', 'david.brown@email.com'],
        'Course': ['Python Programming', 'Data Science', 'Web Development', 
                  'Machine Learning', 'Database Management'],
        'Completion_Date': ['December 15, 2024', 'December 16, 2024', 'December 17, 2024',
                           'December 18, 2024', 'December 19, 2024']
    }
    
    df = pd.DataFrame(sample_data)
    df.to_excel('students_data.xlsx', index=False)
    print("Sample Excel file 'students_data.xlsx' created successfully!")

if __name__ == "__main__":
    # Create sample data if no Excel file exists
    if not os.path.exists('students_data.xlsx'):
        create_sample_excel()
    
    try:
        # Import configuration
        from config import *
        
        print("=== BULK EMAIL CERTIFICATE SYSTEM ===")
        print("Configuration loaded from config.py")
        print(f"Excel file: {EXCEL_FILE}")
        print(f"SMTP Server: {SMTP_SERVER}")
        print(f"SMTP Port: {SMTP_PORT}")
        print(f"Email Address: {EMAIL_ADDRESS}")
        
        # Check if credentials are set
        if EMAIL_ADDRESS == "your_email@gmail.com" or EMAIL_PASSWORD == "your_app_password":
            print("\n⚠️  WARNING: Please update your email credentials in config.py")
            print("1. Open config.py")
            print("2. Update EMAIL_ADDRESS and EMAIL_PASSWORD")
            print("3. For Gmail users, use an App Password")
            print("\nRun the system after updating credentials:")
            print("system = BulkEmailCertificateSystem(EXCEL_FILE, SMTP_SERVER, SMTP_PORT, EMAIL_ADDRESS, EMAIL_PASSWORD)")
            print("system.process_all_students()")
        else:
            # Run the system
            print("\n🚀 Starting bulk email process...")
            system = BulkEmailCertificateSystem(EXCEL_FILE, SMTP_SERVER, SMTP_PORT, EMAIL_ADDRESS, EMAIL_PASSWORD)
            system.process_all_students()
            
    except ImportError:
        print("=== BULK EMAIL CERTIFICATE SYSTEM ===")
        print("Configuration file not found. Please create config.py with your settings.")
        print("\nExample config.py:")
        print("SMTP_SERVER = 'smtp.gmail.com'")
        print("SMTP_PORT = 587")
        print("EMAIL_ADDRESS = 'your_email@gmail.com'")
        print("EMAIL_PASSWORD = 'your_app_password'")
        print("EXCEL_FILE = 'students_data.xlsx'")
    except Exception as e:
        print(f"Error: {e}")
        print("Please check your configuration and try again.") 