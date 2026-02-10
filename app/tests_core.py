import pytest
import io
from src.core.converter import ExcelToPdfConverter
from src.core.validators import validate_file
from src.core.exceptions import InvalidFileTypeError, FileTooLargeError

def test_validator_valid():
    # Should not raise
    validate_file("test.xlsx", 1024, {"xlsx"}, 50)

def test_validator_invalid_ext():
    with pytest.raises(InvalidFileTypeError):
        validate_file("test.exe", 1024, {"xlsx"}, 50)

def test_validator_too_large():
    with pytest.raises(FileTooLargeError):
        validate_file("test.xlsx", 100 * 1024 * 1024, {"xlsx"}, 50)

def test_converter_init():
    converter = ExcelToPdfConverter(dpi=300)
    assert converter.dpi == 300
    assert converter.styles is not None
