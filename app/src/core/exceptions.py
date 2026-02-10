class AppError(Exception):
    """Base class for application exceptions."""
    pass

class ValidationError(AppError):
    """Raised when file validation fails."""
    pass

class ConversionError(AppError):
    """Raised when file conversion fails."""
    pass

class FileTooLargeError(ValidationError):
    """Raised when the uploaded file exceeds the size limit."""
    pass

class InvalidFileTypeError(ValidationError):
    """Raised when the file type is not supported."""
    pass
