from __future__ import annotations
import os
from dataclasses import dataclass
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

@dataclass(frozen=True)
class Settings:
    secret_key: str = os.getenv("SECRET_KEY", "shalfa-dev-secret-change-me")
    database_url: str = os.getenv("DATABASE_URL", "").strip()
    database_path: Path = Path(os.getenv("DATABASE_PATH", BASE_DIR / "data" / "shalfa.db"))
    environment: str = os.getenv("ENVIRONMENT", "development")

    @property
    def database_engine(self) -> str:
        return "postgresql" if self.database_url.startswith(("postgres://", "postgresql://")) else "sqlite"

settings = Settings()
