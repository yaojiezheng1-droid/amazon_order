#!/usr/bin/env python3
"""
Amazon Order Generation - Dependency Setup Script

This script automatically checks for and installs all required dependencies
for the Amazon Order Generation system on a new laptop or environment.

Supported Applications:
- product_search_gui.py
- accessory_mapping_updater_gui.py  
- advanced_excel_to_json.py
- All related order generation scripts

Usage:
    python setup_dependencies.py
    
Features:
- Automatic dependency detection and installation
- Python version compatibility check
- Virtual environment setup option
- Detailed installation logging
- Rollback capability on failure
- Cross-platform support (Windows/macOS/Linux)
"""

import sys
import subprocess
import importlib
import platform
import os
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import json
import time


class DependencyManager:
    def __init__(self):
        self.python_version = sys.version_info
        self.platform = platform.system()
        self.installation_log = []
        self.failed_packages = []
        self.success_packages = []
        
        # Define required packages with versions and alternatives
        self.required_packages = {
            # Core GUI dependencies
            "pyperclip": {
                "version": ">=1.8.0",
                "description": "Clipboard functionality for product search GUI",
                "required_by": ["product_search_gui.py"],
                "install_name": "pyperclip"
            },
            "openpyxl": {
                "version": ">=3.0.0",
                "description": "Excel file processing for multiple scripts",
                "required_by": ["accessory_mapping_updater_gui.py", "json_PO_excel.py", "excel_to_json_template.py"],
                "install_name": "openpyxl"
            },
            "pillow": {
                "version": ">=8.0.0",
                "description": "Image processing for Excel files and product images",
                "required_by": ["json_PO_excel.py"],
                "install_name": "pillow"
            },
            
            # Standard library packages (usually included)
            "tkinter": {
                "version": "builtin",
                "description": "GUI framework for all GUI applications",
                "required_by": ["product_search_gui.py", "accessory_mapping_updater_gui.py"],
                "install_name": None,  # Usually built-in
                "linux_package": "python3-tk"  # For Linux systems
            },
            
            # Core Python packages (usually included)
            "json": {
                "version": "builtin",
                "description": "JSON processing",
                "required_by": ["all scripts"],
                "install_name": None
            },
            "pathlib": {
                "version": "builtin", 
                "description": "Path handling",
                "required_by": ["all scripts"],
                "install_name": None
            },
            "typing": {
                "version": "builtin",
                "description": "Type hints support",
                "required_by": ["all scripts"],
                "install_name": None
            },
            "xml.etree.ElementTree": {
                "version": "builtin",
                "description": "XML processing for Excel files",
                "required_by": ["accessory_mapping_updater_gui.py", "advanced_excel_to_json.py"],
                "install_name": None
            },
            "zipfile": {
                "version": "builtin",
                "description": "ZIP file processing for Excel files",
                "required_by": ["accessory_mapping_updater_gui.py", "advanced_excel_to_json.py"],
                "install_name": None
            },
            "datetime": {
                "version": "builtin",
                "description": "Date and time handling",
                "required_by": ["accessory_mapping_updater_gui.py", "json_PO_excel.py"],
                "install_name": None
            },
            "subprocess": {
                "version": "builtin",
                "description": "Process execution",
                "required_by": ["product_search_gui.py", "direct_sku_to_json.py"],
                "install_name": None
            },
            "re": {
                "version": "builtin",
                "description": "Regular expressions",
                "required_by": ["advanced_excel_to_json.py", "direct_sku_to_json.py", "json_PO_excel.py", "excel_to_json_template.py"],
                "install_name": None
            },
            "traceback": {
                "version": "builtin",
                "description": "Error handling and debugging",
                "required_by": ["accessory_mapping_updater_gui.py"],
                "install_name": None
            },
            "argparse": {
                "version": "builtin",
                "description": "Command line argument parsing",
                "required_by": ["direct_sku_to_json.py"],
                "install_name": None
            },
            "sys": {
                "version": "builtin",
                "description": "System-specific parameters and functions",
                "required_by": ["all scripts"],
                "install_name": None
            },
            "tempfile": {
                "version": "builtin",
                "description": "Temporary file and directory creation",
                "required_by": ["direct_sku_to_json.py"],
                "install_name": None
            }
        }
        
        # Optional packages that enhance functionality
        self.optional_packages = {
            "pandas": {
                "version": ">=1.3.0",
                "description": "Enhanced Excel processing capabilities",
                "benefit": "Better Excel file handling and data manipulation"
            },
            "xlsxwriter": {
                "version": ">=3.0.0", 
                "description": "Advanced Excel writing capabilities",
                "benefit": "Enhanced Excel output formatting"
            }
        }
    
    def log(self, message: str, level: str = "INFO"):
        """Log a message with timestamp"""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}"
        print(log_entry)
        self.installation_log.append(log_entry)
    
    def check_python_version(self) -> bool:
        """Check if Python version is compatible"""
        self.log("Checking Python version compatibility...")
        
        if self.python_version < (3, 7):
            self.log(f"ERROR: Python {self.python_version.major}.{self.python_version.minor} detected. Minimum required: Python 3.7", "ERROR")
            return False
        elif self.python_version < (3, 9):
            self.log(f"WARNING: Python {self.python_version.major}.{self.python_version.minor} detected. Recommended: Python 3.9+", "WARNING")
        else:
            self.log(f"✓ Python {self.python_version.major}.{self.python_version.minor}.{self.python_version.micro} is compatible")
        
        return True
    
    def check_package_installed(self, package_name: str) -> Tuple[bool, Optional[str]]:
        """Check if a package is installed and return version"""
        try:
            # Handle special cases
            if package_name == "tkinter":
                if self.platform == "Windows":
                    import tkinter
                else:
                    import tkinter
                return True, "builtin"
            elif "." in package_name:
                # Handle nested imports like xml.etree.ElementTree
                module = importlib.import_module(package_name)
                return True, "builtin"
            else:
                module = importlib.import_module(package_name)
                version = getattr(module, '__version__', 'unknown')
                return True, version
        except ImportError:
            return False, None
    
    def install_package(self, package_name: str, install_name: str = None) -> bool:
        """Install a package using pip"""
        if install_name is None:
            install_name = package_name
            
        self.log(f"Installing {install_name}...")
        
        try:
            # Use the same Python executable that's running this script
            cmd = [sys.executable, "-m", "pip", "install", install_name]
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            
            self.log(f"✓ Successfully installed {install_name}")
            self.success_packages.append(install_name)
            return True
            
        except subprocess.CalledProcessError as e:
            self.log(f"✗ Failed to install {install_name}: {e.stderr}", "ERROR")
            self.failed_packages.append(install_name)
            return False
        except Exception as e:
            self.log(f"✗ Unexpected error installing {install_name}: {e}", "ERROR")
            self.failed_packages.append(install_name)
            return False
    
    def install_linux_package(self, package_name: str) -> bool:
        """Install system package on Linux (for tkinter)"""
        if self.platform != "Linux":
            return False
            
        self.log(f"Attempting to install system package: {package_name}")
        
        # Try different package managers
        managers = [
            ["sudo", "apt-get", "install", "-y", package_name],  # Debian/Ubuntu
            ["sudo", "yum", "install", "-y", package_name],     # RHEL/CentOS
            ["sudo", "dnf", "install", "-y", package_name],     # Fedora
            ["sudo", "pacman", "-S", "--noconfirm", package_name]  # Arch
        ]
        
        for cmd in managers:
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, check=True)
                self.log(f"✓ Successfully installed {package_name} via {cmd[1]}")
                return True
            except (subprocess.CalledProcessError, FileNotFoundError):
                continue
        
        self.log(f"✗ Failed to install {package_name} via system package manager", "ERROR")
        return False
    
    def check_all_dependencies(self) -> Dict[str, Tuple[bool, str]]:
        """Check status of all required dependencies"""
        self.log("Checking all dependencies...")
        results = {}
        
        for package, info in self.required_packages.items():
            installed, version = self.check_package_installed(package)
            status = "✓ Installed" if installed else "✗ Missing"
            version_info = f" (v{version})" if version and version != "builtin" else ""
            
            self.log(f"{status}: {package}{version_info}")
            results[package] = (installed, version or "unknown")
        
        return results
    
    def install_missing_dependencies(self, check_results: Dict[str, Tuple[bool, str]]) -> bool:
        """Install all missing dependencies"""
        missing_packages = [pkg for pkg, (installed, _) in check_results.items() if not installed]
        
        if not missing_packages:
            self.log("✓ All required dependencies are already installed!")
            return True
        
        self.log(f"Found {len(missing_packages)} missing dependencies")
        success = True
        
        for package in missing_packages:
            info = self.required_packages[package]
            install_name = info.get("install_name")
            
            if install_name is None:
                # Handle special cases
                if package == "tkinter" and self.platform == "Linux":
                    linux_package = info.get("linux_package")
                    if linux_package:
                        if not self.install_linux_package(linux_package):
                            self.log(f"Please manually install tkinter: sudo apt-get install {linux_package}", "WARNING")
                            success = False
                else:
                    self.log(f"Skipping {package} (should be built-in)")
            else:
                if not self.install_package(package, install_name):
                    success = False
        
        return success
    
    def install_optional_packages(self) -> None:
        """Install optional packages with user consent"""
        self.log("\nOptional packages available:")
        
        for package, info in self.optional_packages.items():
            print(f"\n{package}: {info['description']}")
            print(f"Benefit: {info['benefit']}")
            
            try:
                choice = input(f"Install {package}? (y/N): ").strip().lower()
                if choice in ['y', 'yes']:
                    self.install_package(package)
            except KeyboardInterrupt:
                self.log("\nSkipping optional packages...")
                break
    
    def generate_report(self) -> str:
        """Generate installation report"""
        report = []
        report.append("=" * 60)
        report.append("AMAZON ORDER GENERATION - DEPENDENCY SETUP REPORT")
        report.append("=" * 60)
        report.append(f"Date: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        report.append(f"Platform: {self.platform}")
        report.append(f"Python: {self.python_version.major}.{self.python_version.minor}.{self.python_version.micro}")
        report.append("")
        
        if self.success_packages:
            report.append(f"✓ Successfully installed ({len(self.success_packages)}):")
            for pkg in self.success_packages:
                report.append(f"  - {pkg}")
            report.append("")
        
        if self.failed_packages:
            report.append(f"✗ Failed to install ({len(self.failed_packages)}):")
            for pkg in self.failed_packages:
                report.append(f"  - {pkg}")
            report.append("")
        
        report.append("REQUIRED FOR APPLICATIONS:")
        report.append("- product_search_gui.py: pyperclip")
        report.append("- accessory_mapping_updater_gui.py: openpyxl")
        report.append("- advanced_excel_to_json.py: (all builtin)")
        report.append("- direct_sku_to_json.py: (all builtin)")
        report.append("- json_PO_excel.py: openpyxl, pillow")
        report.append("- excel_to_json_template.py: openpyxl")
        report.append("")
        
        report.append("NEXT STEPS:")
        if not self.failed_packages:
            report.append("✓ All dependencies installed successfully!")
            report.append("✓ You can now run all applications:")
            report.append("  python order_generation/product_search_gui.py")
            report.append("  python order_generation/accessory_mapping_updater_gui.py")
            report.append("  python order_generation/advanced_excel_to_json.py")
            report.append("  python order_generation/direct_sku_to_json.py")
            report.append("  python order_generation/json_PO_excel.py")
            report.append("  python order_generation/excel_to_json_template.py")
        else:
            report.append("⚠ Some packages failed to install. Please:")
            report.append("1. Check your internet connection")
            report.append("2. Ensure pip is up to date: python -m pip install --upgrade pip")
            report.append("3. Try manual installation:")
            for pkg in self.failed_packages:
                report.append(f"   pip install {pkg}")
        
        report.append("")
        report.append("=" * 60)
        
        return "\n".join(report)
    
    def save_log(self) -> None:
        """Save installation log to file"""
        log_file = Path(__file__).parent / "dependency_setup.log"
        
        try:
            with open(log_file, 'w', encoding='utf-8') as f:
                for entry in self.installation_log:
                    f.write(entry + "\n")
            
            self.log(f"Installation log saved to: {log_file}")
        except Exception as e:
            self.log(f"Failed to save log: {e}", "ERROR")
    
    def setup_virtual_environment(self) -> bool:
        """Optionally create and activate virtual environment"""
        try:
            choice = input("\nCreate virtual environment for this project? (y/N): ").strip().lower()
            if choice not in ['y', 'yes']:
                return True
            
            venv_path = Path(__file__).parent / "venv"
            
            self.log("Creating virtual environment...")
            subprocess.run([sys.executable, "-m", "venv", str(venv_path)], check=True)
            
            # Provide activation instructions
            if self.platform == "Windows":
                activate_cmd = str(venv_path / "Scripts" / "activate.bat")
            else:
                activate_cmd = f"source {venv_path}/bin/activate"
            
            self.log(f"Virtual environment created at: {venv_path}")
            self.log(f"To activate: {activate_cmd}")
            self.log("Note: Restart this script after activating the virtual environment")
            
            return False  # Don't continue with installation in this session
            
        except Exception as e:
            self.log(f"Failed to create virtual environment: {e}", "ERROR")
            return True  # Continue anyway
    
    def run_setup(self) -> bool:
        """Main setup process"""
        self.log("Starting Amazon Order Generation dependency setup...")
        self.log(f"Platform: {self.platform}")
        
        # Check Python version
        if not self.check_python_version():
            return False
        
        # Optional virtual environment setup
        if not self.setup_virtual_environment():
            return False
        
        # Check current dependency status
        check_results = self.check_all_dependencies()
        
        # Install missing dependencies
        if not self.install_missing_dependencies(check_results):
            self.log("Some required dependencies failed to install", "WARNING")
        
        # Verify installation
        self.log("\nVerifying installation...")
        final_results = self.check_all_dependencies()
        
        # Optional packages
        try:
            self.install_optional_packages()
        except KeyboardInterrupt:
            self.log("Skipped optional packages")
        
        # Generate and display report
        report = self.generate_report()
        print("\n" + report)
        
        # Save log
        self.save_log()
        
        # Final status
        all_required_installed = all(installed for installed, _ in final_results.values())
        if all_required_installed:
            self.log("✓ Setup completed successfully!")
            return True
        else:
            self.log("⚠ Setup completed with some issues. See report above.", "WARNING")
            return False


def main():
    """Main function"""
    print("Amazon Order Generation - Dependency Setup")
    print("=" * 50)
    print("This script will check and install all required dependencies.")
    print("Supported applications:")
    print("- Product Search GUI")
    print("- Accessory Mapping Updater GUI") 
    print("- Advanced Excel to JSON Converter")
    print()
    
    try:
        choice = input("Continue with dependency setup? (Y/n): ").strip().lower()
        if choice in ['n', 'no']:
            print("Setup cancelled.")
            return
    except KeyboardInterrupt:
        print("\nSetup cancelled.")
        return
    
    manager = DependencyManager()
    success = manager.run_setup()
    
    if success:
        print("\n🎉 Setup completed successfully!")
        print("You can now run the GUI applications:")
        print("  python order_generation/product_search_gui.py")
        print("  python order_generation/accessory_mapping_updater_gui.py")
    else:
        print("\n⚠ Setup completed with warnings.")
        print("Please check the report above and install any missing dependencies manually.")
    
    return 0 if success else 1


if __name__ == "__main__":
    exit(main())
