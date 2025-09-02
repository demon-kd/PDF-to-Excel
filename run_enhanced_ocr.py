
"""
Enhanced Electoral Roll OCR Launcher
===================================

Simple launcher script to run the enhanced OCR converter with proper setup checking.
"""

import os
import sys
import subprocess

def check_python():
    """Check Python version"""
    if sys.version_info < (3, 7):
        print("❌ Python 3.7 or higher required")
        return False
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor} detected")
    return True

def check_packages():
    """Check required packages"""
    required_packages = [
        ('pandas', 'pandas'),
        ('openpyxl', 'openpyxl'), 
        ('pytesseract', 'pytesseract'),
        ('PIL', 'Pillow'),
        ('pdf2image', 'pdf2image')
    ]

    missing = []
    for import_name, package_name in required_packages:
        try:
            __import__(import_name)
            print(f"✅ {package_name}")
        except ImportError:
            print(f"❌ {package_name} - Missing")
            missing.append(package_name)

    if missing:
        print(f"\n📦 Install missing packages:")
        print(f"pip install {' '.join(missing)}")
        return False

    return True

def check_tesseract():
    """Check Tesseract OCR"""
    try:
        import pytesseract
        version = pytesseract.get_tesseract_version()
        print(f"✅ Tesseract OCR {version}")
        return True
    except Exception as e:
        print(f"❌ Tesseract OCR not found: {e}")
        print("\n📥 Install Tesseract OCR:")
        print("Windows: https://github.com/UB-Mannheim/tesseract/wiki")
        print("macOS: brew install tesseract") 
        print("Linux: sudo apt-get install tesseract-ocr")
        return False

def main():
    """Main launcher"""
    print("🚀 ENHANCED ELECTORAL ROLL OCR LAUNCHER")
    print("="*50)

    # Check requirements
    print("\n🔍 Checking system requirements...")
    if not check_python():
        input("Press Enter to exit...")
        return

    if not check_packages():
        input("Press Enter to exit...")
        return

    if not check_tesseract():
        input("Press Enter to exit...")
        return

    print("\n✅ All requirements satisfied!")
    print("\n🎯 Choose how to run the enhanced OCR converter:")
    print("1. GUI Version (Recommended)")
    print("2. Command Line Version")
    print("3. Exit")

    while True:
        try:
            choice = input("\nEnter choice (1-3): ").strip()

            if choice == '1':
                print("\n🖥️ Starting GUI version...")
                if os.path.exists('enhanced_electoral_roll_ocr_gui.py'):
                    subprocess.run([sys.executable, 'enhanced_electoral_roll_ocr_gui.py'])
                else:
                    print("❌ enhanced_electoral_roll_ocr_gui.py not found")
                break

            elif choice == '2':
                print("\n⌨️ Starting command line version...")
                pdf_file = input("Enter PDF file path: ").strip().strip('"')
                excel_file = input("Enter Excel output file path: ").strip().strip('"')

                if pdf_file and excel_file:
                    if os.path.exists('enhanced_electoral_roll_ocr.py'):
                        subprocess.run([sys.executable, 'enhanced_electoral_roll_ocr.py', pdf_file, excel_file])
                    else:
                        print("❌ enhanced_electoral_roll_ocr.py not found")
                else:
                    print("❌ Please provide both file paths")
                break

            elif choice == '3':
                print("👋 Goodbye!")
                break

            else:
                print("❌ Invalid choice. Please enter 1, 2, or 3.")

        except KeyboardInterrupt:
            print("\n👋 Goodbye!")
            break
        except Exception as e:
            print(f"❌ Error: {e}")
            break

    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
