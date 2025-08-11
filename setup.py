"""
Setup script for Bulk Email Certificate System
This script helps you set up the system quickly
"""

import os
import sys
import subprocess

def install_requirements():
    """Install required packages"""
    print("📦 Installing required packages...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Packages installed successfully!")
        return True
    except subprocess.CalledProcessError:
        print("❌ Failed to install packages. Please install manually:")
        print("pip install -r requirements.txt")
        return False

def create_example_files():
    """Create example files"""
    print("\n📄 Creating example files...")
    
    # Create example Excel file
    try:
        from create_example_excel import create_example_excel
        create_example_excel()
    except Exception as e:
        print(f"❌ Failed to create example Excel file: {e}")
    
    print("✅ Example files created!")

def check_config():
    """Check if config file exists and is properly configured"""
    print("\n⚙️  Checking configuration...")
    
    if not os.path.exists('config.py'):
        print("❌ config.py not found. Please create it with your email settings.")
        return False
    
    try:
        from config import EMAIL_ADDRESS, EMAIL_PASSWORD
        if EMAIL_ADDRESS == "your_email@gmail.com" or EMAIL_PASSWORD == "your_app_password":
            print("⚠️  Please update your email credentials in config.py")
            return False
        else:
            print("✅ Configuration looks good!")
            return True
    except ImportError:
        print("❌ Error reading config.py")
        return False

def run_test():
    """Run a test to verify everything works"""
    print("\n🧪 Running system test...")
    try:
        from test_system import test_certificate_generation
        test_certificate_generation()
        print("✅ System test completed successfully!")
        return True
    except Exception as e:
        print(f"❌ System test failed: {e}")
        return False

def main():
    """Main setup function"""
    print("🚀 Bulk Email Certificate System - Setup")
    print("=" * 50)
    
    # Step 1: Install requirements
    if not install_requirements():
        return
    
    # Step 2: Create example files
    create_example_files()
    
    # Step 3: Check configuration
    config_ok = check_config()
    
    # Step 4: Run test
    if config_ok:
        run_test()
    
    print("\n" + "=" * 50)
    print("🎉 Setup completed!")
    
    if not config_ok:
        print("\n📝 Next steps:")
        print("1. Edit config.py and add your email credentials")
        print("2. For Gmail users, enable 2FA and generate an App Password")
        print("3. Update students_data.xlsx with your student data")
        print("4. Run: python bulk_email_system.py")
    else:
        print("\n✅ System is ready to use!")
        print("Run: python bulk_email_system.py")

if __name__ == "__main__":
    main() 