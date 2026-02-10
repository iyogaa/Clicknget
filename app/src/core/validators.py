import os
from typing import Set
from src.core.exceptions import InvalidFileTypeError, FileTooLargeError

def validate_file(filename: str, file_size: int, allowed_extensions: Set[str], max_size_mb: int):
    """
    Validate file extension and size.
    """
    ext = os.path.splitext(filename)[1].lower().lstrip('.')
    if ext not in allowed_extensions:
        raise InvalidFileTypeError(f"Extension .{ext} is not supported. Allowed: {', '.join(allowed_extensions)}")
    
    if file_size > max_size_mb * 1024 * 1024:
        raise FileTooLargeError(f"File size {file_size / (1024 * 1024):.2f}MB exceeds the {max_size_mb}MB limit.")
