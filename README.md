# Bulk Email Certificate System

A Python script that automatically generates certificates and sends them via email to students based on data from an Excel file.

## Features

- 📊 Reads student details from Excel files
- 🎓 Generates professional PDF certificates
- 📧 Sends personalized emails with certificate attachments
- 🔄 Processes multiple students automatically
- 📈 Provides detailed progress tracking
- 🛡️ Includes error handling and validation

## Prerequisites

- Python 3.7 or higher
- Excel file with student data
- Email account with SMTP access

## Installation

1. **Clone or download this project**

2. **Install required packages:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up your email credentials:**
   - Edit `config.py` and update your email settings
   - For Gmail users, you'll need to:
     - Enable 2-factor authentication
     - Generate an App Password
     - Use the App Password instead of your regular password

## Excel File Format

Your Excel file should have the following columns:

| Column Name | Description | Required |
|-------------|-------------|----------|
| Name | Student's full name | ✅ Yes |
| Email | Student's email address | ✅ Yes |
| Course | Course name | ✅ Yes |
| Completion_Date | Date of completion (optional) | ❌ No |

### Sample Excel Structure:
```
Name            | Email                    | Course              | Completion_Date
John Doe        | john.doe@email.com      | Python Programming  | December 15, 2024
Jane Smith      | jane.smith@email.com    | Data Science        | December 16, 2024
```

## Configuration

1. **Update `config.py`:**
   ```python
   EMAIL_ADDRESS = "your_email@gmail.com"
   EMAIL_PASSWORD = "your_app_password"
   ```

2. **Email Provider Settings:**
   - **Gmail:** `smtp.gmail.com:587`
   - **Outlook:** `smtp-mail.outlook.com:587`
   - **Yahoo:** `smtp.mail.yahoo.com:587`

## Usage

### Method 1: Run the main script
```bash
python bulk_email_system.py
```

### Method 2: Use as a module
```python
from bulk_email_system import BulkEmailCertificateSystem
from config import *

# Initialize the system
system = BulkEmailCertificateSystem(
    EXCEL_FILE, 
    SMTP_SERVER, 
    SMTP_PORT, 
    EMAIL_ADDRESS, 
    EMAIL_PASSWORD
)

# Process all students
system.process_all_students()
```

## Certificate Template

The system generates professional certificates with:
- Student name
- Course name
- Completion date
- Unique certificate ID
- Professional formatting
- Signature line

## Output

The system creates:
- `certificates/` folder with generated PDF certificates
- Console output showing progress
- Summary of successful/failed emails

## Troubleshooting

### Common Issues:

1. **Email Authentication Error:**
   - Ensure 2-factor authentication is enabled
   - Use App Password instead of regular password
   - Check SMTP server and port settings

2. **Excel File Not Found:**
   - Ensure the Excel file exists in the project directory
   - Check the file path in `config.py`

3. **Missing Columns:**
   - Ensure your Excel file has required columns: Name, Email, Course
   - Column names are case-sensitive

4. **Rate Limiting:**
   - The system includes delays between emails
   - Increase `DELAY_BETWEEN_EMAILS` in `config.py` if needed

### Gmail Setup Guide:

1. Go to your Google Account settings
2. Enable 2-Step Verification
3. Generate an App Password:
   - Go to Security → App passwords
   - Select "Mail" and your device
   - Copy the generated password
4. Use this App Password in `config.py`

## Security Notes

- Never commit your email password to version control
- Use environment variables for production use
- Consider using OAuth2 for better security

## Customization

### Modify Certificate Template:
Edit the `generate_certificate()` method in `bulk_email_system.py` to customize:
- Certificate layout
- Colors and fonts
- Additional information
- Logo placement

### Customize Email Template:
Edit the `send_email()` method to modify:
- Email subject line
- Email body content
- Additional attachments

## License

This project is open source and available under the MIT License.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Verify your configuration settings
3. Ensure all dependencies are installed 