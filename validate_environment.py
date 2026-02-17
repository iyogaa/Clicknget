"""
Environment Configuration Validator
Checks for missing dependencies and environment issues before app startup
"""

import sys
import importlib
from typing import List, Tuple

def check_python_version() -> Tuple[bool, str]:
    """Verify Python version is 3.10 or higher"""
    version = sys.version_info
    if version.major == 3 and version.minor >= 10:
        return True, f"✓ Python {version.major}.{version.minor}.{version.micro}"
    return False, f"✗ Python {version.major}.{version.minor}.{version.micro} (requires 3.10+)"

def check_dependencies() -> List[Tuple[str, bool, str]]:
    """Check if all required dependencies are installed"""
    required_packages = [
        ("streamlit", "Streamlit"),
        ("pandas", "Pandas"),
        ("numpy", "NumPy"),
        ("openpyxl", "OpenPyXL"),
        ("fitz", "PyMuPDF"),
        ("yaml", "PyYAML"),
        ("PIL", "Pillow"),
        ("reportlab", "ReportLab"),
        ("dateutil", "python-dateutil"),
        ("mammoth", "Mammoth"),
        ("fuzzywuzzy", "FuzzyWuzzy"),
    ]
    
    results = []
    for module_name, display_name in required_packages:
        try:
            importlib.import_module(module_name)
            results.append((display_name, True, "✓ Installed"))
        except ImportError:
            results.append((display_name, False, "✗ Missing"))
    
    return results

def check_secrets_file() -> Tuple[bool, str]:
    """Check if secrets.toml exists"""
    import os
    secrets_path = os.path.join(".streamlit", "secrets.toml")
    if os.path.exists(secrets_path):
        return True, "✓ secrets.toml found"
    return False, "✗ secrets.toml missing (authentication will fail)"

def check_template_file() -> Tuple[bool, str]:
    """Check if Template.xlsx exists"""
    import os
    if os.path.exists("Template.xlsx"):
        return True, "✓ Template.xlsx found"
    return False, "⚠ Template.xlsx missing (MVR processing may fail)"

def validate_environment() -> bool:
    """Run all validation checks and return overall status"""
    print("\n" + "="*60)
    print("CLICKNGET - Environment Validation")
    print("="*60 + "\n")
    
    all_ok = True
    
    # Python version
    py_ok, py_msg = check_python_version()
    print(f"Python Version: {py_msg}")
    if not py_ok:
        all_ok = False
    
    print("\nDependencies:")
    dep_results = check_dependencies()
    for name, ok, msg in dep_results:
        print(f"  {msg} {name}")
        if not ok:
            all_ok = False
    
    print("\nConfiguration Files:")
    secrets_ok, secrets_msg = check_secrets_file()
    print(f"  {secrets_msg}")
    if not secrets_ok:
        all_ok = False
    
    template_ok, template_msg = check_template_file()
    print(f"  {template_msg}")
    # Template is warning only, not critical
    
    print("\n" + "="*60)
    if all_ok:
        print("✓ Environment validation PASSED")
        print("="*60 + "\n")
        return True
    else:
        print("✗ Environment validation FAILED")
        print("\nTo fix:")
        print("  1. Run: setup_environment.bat (Windows)")
        print("  2. Or manually: pip install -r requirements.txt")
        print("  3. Ensure .streamlit/secrets.toml exists")
        print("="*60 + "\n")
        return False

if __name__ == "__main__":
    success = validate_environment()
    sys.exit(0 if success else 1)
