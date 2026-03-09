"""
Import Verification Script
Checks all Python files for imports and verifies they're in requirements.txt
"""

import os
import re
import sys
from pathlib import Path

def extract_imports_from_file(filepath):
    """Extract all import statements from a Python file"""
    imports = set()
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # Match: import module
        for match in re.finditer(r'^\s*import\s+(\w+)', content, re.MULTILINE):
            imports.add(match.group(1))
        
        # Match: from module import ...
        for match in re.finditer(r'^\s*from\s+(\w+)', content, re.MULTILINE):
            imports.add(match.group(1))
            
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
    
    return imports

def get_all_python_files(directory):
    """Get all Python files in directory"""
    python_files = []
    for root, dirs, files in os.walk(directory):
        # Skip venv and __pycache__
        dirs[:] = [d for d in dirs if d not in ['venv', '.venv', '__pycache__', '.git', 'app']]
        
        for file in files:
            if file.endswith('.py'):
                python_files.append(os.path.join(root, file))
    
    return python_files

def main():
    print("=" * 70)
    print("IMPORT VERIFICATION - Checking all Python files")
    print("=" * 70)
    print()
    
    # Standard library modules (don't need to be in requirements.txt)
    stdlib_modules = {
        'os', 'sys', 'io', 're', 'datetime', 'string', 'typing', 'traceback',
        'argparse', 'logging', 'copy', 'unittest', 'tempfile', 'zipfile',
        'collections', 'itertools', 'functools', 'pathlib', 'json', 'csv',
        'math', 'random', 'time', 'warnings', 'abc', 'enum', 'dataclasses',
        'difflib', 'unicodedata', 'importlib', 'email'
    }
    
    # Expected third-party packages
    expected_packages = {
        'streamlit', 'pandas', 'numpy', 'openpyxl', 'fitz', 'reportlab',
        'PIL', 'mammoth', 'fuzzywuzzy', 'yaml', 'dateutil'
    }
    
    # Get all Python files
    base_dir = os.path.dirname(os.path.abspath(__file__))
    python_files = get_all_python_files(base_dir)
    
    print(f"Found {len(python_files)} Python files to check\n")
    
    all_imports = set()
    file_imports = {}
    
    for filepath in python_files:
        rel_path = os.path.relpath(filepath, base_dir)
        imports = extract_imports_from_file(filepath)
        if imports:
            file_imports[rel_path] = imports
            all_imports.update(imports)
    
    # Filter out standard library and local modules
    third_party = all_imports - stdlib_modules
    third_party = {imp for imp in third_party if not imp.startswith('_')}
    
    # Remove local modules (features, utils, mvr_renewal, etc.)
    local_modules = {'features', 'utils', 'mvr_renewal', 'Alltran', 'Hdvi', 'Pdf_maker', 'pdf_play', 'app', 'validate_environment', 'verify_imports'}
    third_party = third_party - local_modules
    
    print("Third-Party Packages Used:")
    print("-" * 70)
    for pkg in sorted(third_party):
        status = "✅" if pkg in expected_packages else "⚠️ "
        print(f"{status} {pkg}")
    
    print()
    print("=" * 70)
    
    # Check for unexpected packages
    unexpected = third_party - expected_packages
    if unexpected:
        print("⚠️  UNEXPECTED PACKAGES (not in requirements.txt):")
        for pkg in sorted(unexpected):
            print(f"   - {pkg}")
            # Show which files use it
            for file, imports in file_imports.items():
                if pkg in imports:
                    print(f"     Used in: {file}")
        print()
        return 1
    else:
        print("✅ ALL IMPORTS VERIFIED - All packages are in requirements.txt")
        print()
        return 0

if __name__ == "__main__":
    sys.exit(main())
