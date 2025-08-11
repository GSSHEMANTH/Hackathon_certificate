"""
Test script for the Bulk Email Certificate System
This script tests the certificate generation without sending emails
"""

import os
import pandas as pd
from datetime import datetime
from bulk_email_system import BulkEmailCertificateSystem

def test_certificate_generation():
    """Test certificate generation functionality"""
    print("=== TESTING CERTIFICATE GENERATION ===")
    
    # Create test data
    test_data = {
        'Name': ['Test Student 1', 'Test Student 2'],
        'Email': ['test1@example.com', 'test2@example.com'],
        'Course': ['Test Course 1', 'Test Course 2'],
        'Completion_Date': ['December 20, 2024', 'December 21, 2024']
    }
    
    # Create test Excel file
    df = pd.DataFrame(test_data)
    test_excel_file = "test_students.xlsx"
    df.to_excel(test_excel_file, index=False)
    print(f"Created test Excel file: {test_excel_file}")
    
    # Initialize system with dummy email credentials
    system = BulkEmailCertificateSystem(
        test_excel_file,
        "smtp.gmail.com",
        587,
        "test@example.com",
        "dummy_password"
    )
    
    # Test reading Excel file
    print("\n1. Testing Excel file reading...")
    df_read = system.read_student_data()
    if df_read is not None:
        print(f"✅ Successfully read {len(df_read)} records")
        print("Sample data:")
        print(df_read.head())
    else:
        print("❌ Failed to read Excel file")
        return
    
    # Test certificate generation
    print("\n2. Testing certificate generation...")
    try:
        certificate_path = system.generate_certificate(
            "Test Student",
            "Test Course",
            "December 20, 2024",
            "TEST-001"
        )
        if os.path.exists(certificate_path):
            print(f"✅ Certificate generated successfully: {certificate_path}")
        else:
            print("❌ Certificate file not found")
    except Exception as e:
        print(f"❌ Certificate generation failed: {e}")
    
    # Test processing without sending emails
    print("\n3. Testing full processing (without emails)...")
    try:
        # Override send_email method to prevent actual sending
        original_send_email = system.send_email
        system.send_email = lambda *args: True
        
        system.process_all_students()
        print("✅ Full processing test completed")
        
        # Restore original method
        system.send_email = original_send_email
    except Exception as e:
        print(f"❌ Processing test failed: {e}")
    
    # Clean up test files
    print("\n4. Cleaning up test files...")
    if os.path.exists(test_excel_file):
        os.remove(test_excel_file)
        print(f"✅ Removed {test_excel_file}")
    
    # Check generated certificates
    if os.path.exists("certificates"):
        certificate_files = os.listdir("certificates")
        print(f"✅ Generated {len(certificate_files)} certificate files in certificates/ folder")
        for cert_file in certificate_files:
            print(f"   - {cert_file}")

def test_excel_format():
    """Test different Excel formats"""
    print("\n=== TESTING EXCEL FORMATS ===")
    
    # Test with missing optional column
    test_data_minimal = {
        'Name': ['Minimal Student'],
        'Email': ['minimal@example.com'],
        'Course': ['Minimal Course']
    }
    
    df_minimal = pd.DataFrame(test_data_minimal)
    minimal_file = "test_minimal.xlsx"
    df_minimal.to_excel(minimal_file, index=False)
    
    system = BulkEmailCertificateSystem(
        minimal_file,
        "smtp.gmail.com",
        587,
        "test@example.com",
        "dummy_password"
    )
    
    df_read = system.read_student_data()
    if df_read is not None and 'Completion_Date' in df_read.columns:
        print("✅ System correctly handles missing Completion_Date column")
    else:
        print("❌ System failed to handle missing Completion_Date column")
    
    # Clean up
    if os.path.exists(minimal_file):
        os.remove(minimal_file)

if __name__ == "__main__":
    print("Bulk Email Certificate System - Test Suite")
    print("=" * 50)
    
    # Run tests
    test_certificate_generation()
    test_excel_format()
    
    print("\n" + "=" * 50)
    print("Test completed! Check the certificates/ folder for generated files.")
    print("\nTo run the actual system:")
    print("1. Update config.py with your email credentials")
    print("2. Prepare your Excel file with student data")
    print("3. Run: python bulk_email_system.py") 