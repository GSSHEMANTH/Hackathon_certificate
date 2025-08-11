"""
Script to create an example Excel file with sample student data
"""

import pandas as pd
from datetime import datetime

def create_example_excel():
    """Create an example Excel file with sample student data"""
    
    # Sample student data
    sample_data = {
        'Name': [
            'John Doe',
            'Jane Smith', 
            'Michael Johnson',
            'Sarah Wilson',
            'David Brown',
            'Emily Davis',
            'Robert Miller',
            'Lisa Anderson',
            'James Taylor',
            'Maria Garcia'
        ],
        'Email': [
            'john.doe@email.com',
            'jane.smith@email.com',
            'michael.johnson@email.com',
            'sarah.wilson@email.com',
            'david.brown@email.com',
            'emily.davis@email.com',
            'robert.miller@email.com',
            'lisa.anderson@email.com',
            'james.taylor@email.com',
            'maria.garcia@email.com'
        ],
        'Course': [
            'Python Programming Fundamentals',
            'Data Science with Python',
            'Web Development with Django',
            'Machine Learning Basics',
            'Database Management Systems',
            'JavaScript for Beginners',
            'React.js Development',
            'Cybersecurity Fundamentals',
            'Cloud Computing with AWS',
            'Mobile App Development'
        ],
        'Completion_Date': [
            'December 15, 2024',
            'December 16, 2024',
            'December 17, 2024',
            'December 18, 2024',
            'December 19, 2024',
            'December 20, 2024',
            'December 21, 2024',
            'December 22, 2024',
            'December 23, 2024',
            'December 24, 2024'
        ]
    }
    
    # Create DataFrame
    df = pd.DataFrame(sample_data)
    
    # Save to Excel file
    filename = 'students_data.xlsx'
    df.to_excel(filename, index=False)
    
    print(f"✅ Example Excel file created: {filename}")
    print(f"📊 Contains {len(df)} student records")
    print("\nSample data:")
    print(df.head())
    
    print("\n📋 Required columns:")
    print("- Name: Student's full name")
    print("- Email: Student's email address")
    print("- Course: Course name")
    print("- Completion_Date: Date of completion (optional)")
    
    print("\n💡 You can:")
    print("1. Replace the sample data with your actual student data")
    print("2. Add more students by adding rows")
    print("3. Modify course names and completion dates")
    print("4. Keep the same column structure")

if __name__ == "__main__":
    create_example_excel() 