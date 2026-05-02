from __future__ import annotations

from pathlib import Path

from pydantic import Field
from pydantic import field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


DEFAULT_DOCX_OUTPUT_DIR = Path(__file__).resolve().parents[1] / "var" / "docx-results"


class Settings(BaseSettings):
    model_config = SettingsConfigDict(extra="ignore")

    ai_api_key: str = Field(default="", alias="AI_API_KEY")
    ai_provider_api: str = Field(default="responses", alias="AI_PROVIDER_API")
    ai_profiles_json: str = Field(default="", alias="AI_PROFILES_JSON")
    openai_api_base_url: str = Field(default="https://api.openai.com/v1", alias="OPENAI_API_BASE_URL")
    openai_model: str = Field(default="gpt-4o-mini", alias="OPENAI_MODEL")
    ai_request_timeout_seconds: float = Field(default=30, alias="AI_REQUEST_TIMEOUT_SECONDS")
    ai_fast_max_tokens: int = Field(default=8192, alias="AI_FAST_MAX_TOKENS")
    ai_thinking_max_tokens: int = Field(default=16384, alias="AI_THINKING_MAX_TOKENS")
    backend_cors_origins: str = Field(
        default="https://localhost:3000,http://localhost:3000",
        alias="BACKEND_CORS_ORIGINS",
    )
    backend_log_level: str = Field(default="INFO", alias="BACKEND_LOG_LEVEL")
    docx_output_dir: Path = Field(default=DEFAULT_DOCX_OUTPUT_DIR, alias="DOCX_OUTPUT_DIR")
    docx_retention_days: int = Field(default=7, alias="DOCX_RETENTION_DAYS")

    @field_validator("docx_output_dir")
    @classmethod
    def relative_docx_output_dir_uses_backend_root(cls, value: Path) -> Path:
        return value if value.is_absolute() else Path(__file__).resolve().parents[1] / value

    @field_validator("docx_retention_days")
    @classmethod
    def docx_retention_days_must_be_at_least_seven(cls, value: int) -> int:
        return max(7, value)

    @property
    def backend_cors_origin_list(self) -> list[str]:
        return [origin.strip() for origin in self.backend_cors_origins.split(",") if origin.strip()]


def get_settings() -> Settings:
    return Settings()
