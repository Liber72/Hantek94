#!/usr/bin/env python3
"""
Installation script for VinGroup Financial Analysis Tool
=======================================================

This script installs the required dependencies and sets up the tool.
"""

import subprocess
import sys
import os

def install_package(package):
    """Install a package using pip"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except subprocess.CalledProcessError:
        return False

def main():
    print("🔧 VinGroup Financial Analysis Tool - Installation")
    print("=" * 50)
    
    # Check Python version
    if sys.version_info < (3, 8):
        print("❌ Python 3.8+ required. Current version:", sys.version)
        return False
    
    print("✅ Python version check passed")
    
    # Required packages
    packages = [
        "openpyxl",
        "pandas", 
        "matplotlib",
        "numpy"
    ]
    
    print("\n📦 Installing required packages...")
    
    success_count = 0
    for package in packages:
        print(f"  Installing {package}...", end=" ")
        if install_package(package):
            print("✅")
            success_count += 1
        else:
            print("❌")
    
    print(f"\n📊 Installation summary: {success_count}/{len(packages)} packages installed")
    
    if success_count == len(packages):
        print("\n✅ All packages installed successfully!")
        print("\n🚀 You can now run:")
        print("   python demo.py              # Run demo")
        print("   python vingroup_financial_analyzer.py  # Generate CSV files")
        print("   python excel_generator.py   # Generate Excel file")
        return True
    else:
        print("\n⚠️  Some packages failed to install.")
        print("   You can still use the CSV generation features.")
        print("   Try installing missing packages manually:")
        for package in packages:
            print(f"   pip install {package}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)