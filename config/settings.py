from dataclasses import dataclass, field
from typing import Set

@dataclass(frozen=True)
class AppConfig:
    """Application configuration."""
    APP_NAME: str = "Excel-to-PDF Master"
    MAX_UPLOAD_SIZE_MB: int = 50
    ALLOWED_EXTENSIONS: Set[str] = field(default_factory=lambda: {"xlsx", "xls"})
    DEFAULT_DPI: int = 300
    CACHE_TTL_SECONDS: int = 3600
    THEME_PRIMARY_COLOR: str = "#00d4aa"
    HISTORY_LIMIT: int = 5

config = AppConfig()
