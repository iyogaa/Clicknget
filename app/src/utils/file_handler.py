import os
import tempfile
import shutil
from contextlib import contextmanager
from typing import Generator
from src.utils.logger import logger

@contextmanager
def temporary_file(suffix: str = None) -> Generator[str, None, None]:
    """
    Context manager for creating and clean up a temporary file.
    """
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    try:
        yield path
    finally:
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception as e:
            logger.error(f"Failed to remove temp file {path}: {e}")

@contextmanager
def temporary_directory() -> Generator[str, None, None]:
    """
    Context manager for creating and clean up a temporary directory.
    """
    path = tempfile.mkdtemp()
    try:
        yield path
    finally:
        try:
            shutil.rmtree(path)
        except Exception as e:
            logger.error(f"Failed to remove temp dir {path}: {e}")
