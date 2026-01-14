"""
Test script to demonstrate DOB validation functionality
This script shows how the new _is_valid_dob() method works
"""

import pandas as pd
import re

def _is_valid_dob(val):
    """
    Validate if a DOB value is a valid, parseable date.
    Returns True only if the value can be successfully parsed as a date.
    Returns False for corrupted, malformed, or non-date values.
    """
    if val is None:
        return False
    try:
        # Handle string values
        v = val
        if isinstance(v, str):
            # Strip leading apostrophe if present
            if v.startswith("'"):
                v = v[1:]
            # Strip whitespace
            v = v.strip()
            # Check for empty string
            if v == "":
                return False
            # Check for obvious non-date patterns (e.g., "XX XX X003")
            # If it contains only X's, spaces, and maybe some digits but not a valid date format
            if re.match(r'^[X\s]+\d*$', v, re.IGNORECASE):
                return False
        
        # Try to parse as datetime
        ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts):
            # Try with dayfirst=True as fallback
            ts = pd.to_datetime(v, errors="coerce", dayfirst=True)
        
        # If still NaT (Not a Time), it's invalid
        if pd.isna(ts):
            return False
        
        # Successfully parsed
        return True
    except Exception:
        return False


# Test cases
test_cases = [
    # (value, expected_result, description)
    ("01/15/1990", True, "Valid US date format"),
    ("15/01/1990", True, "Valid UK date format"),
    ("1990-01-15", True, "Valid ISO date format"),
    ("XX XX X003", False, "Corrupted DOB with X's"),
    ("INVALID", False, "Text instead of date"),
    ("###", False, "Special characters"),
    ("", False, "Empty string"),
    (None, False, "None value"),
    ("'01/15/1990", True, "Date with leading apostrophe"),
    ("  01/15/1990  ", True, "Date with whitespace"),
    ("12/32/2020", False, "Invalid day (32)"),
    ("13/01/2020", True, "Valid day-first format"),
    ("XX XX 1990", False, "Partial corruption"),
]

print("=" * 80)
print("DOB VALIDATION TEST RESULTS")
print("=" * 80)
print()

passed = 0
failed = 0

for value, expected, description in test_cases:
    result = _is_valid_dob(value)
    status = "✓ PASS" if result == expected else "✗ FAIL"
    
    if result == expected:
        passed += 1
    else:
        failed += 1
    
    print(f"{status} | {description}")
    print(f"       Value: {repr(value)}")
    print(f"       Expected: {expected}, Got: {result}")
    print()

print("=" * 80)
print(f"SUMMARY: {passed} passed, {failed} failed out of {len(test_cases)} tests")
print("=" * 80)
