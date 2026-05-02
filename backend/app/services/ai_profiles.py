from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Literal

from pydantic import BaseModel, Field, ValidationError, field_validator

from app.settings import Settings, get_settings

ProviderAPI = Literal["responses", "chat"]
DEFAULT_PROFILE_ID = "default"


class AIProfileError(RuntimeError):
    """Raised when an AI profile cannot be resolved from environment config."""


@dataclass(frozen=True)
class AIProfile:
    id: str
    label: str
    api_base_url: str
    api_key: str
    api_key_env: str
    model: str
    default_api: ProviderAPI
    supported_apis: tuple[ProviderAPI, ...]


class AIProfileConfig(BaseModel):
    id: str = Field(..., min_length=1)
    label: str = Field(..., min_length=1)
    api_base_url: str = Field(..., min_length=1)
    api_key_env: str = Field(default="AI_API_KEY", min_length=1)
    model: str = Field(..., min_length=1)
    default_api: ProviderAPI = "responses"
    supported_apis: list[ProviderAPI] = Field(default_factory=lambda: ["responses", "chat"])

    @field_validator("id", "label", "api_base_url", "api_key_env", "model")
    @classmethod
    def string_fields_must_not_be_blank(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("field must not be blank")
        return normalized

    @field_validator("supported_apis")
    @classmethod
    def supported_apis_must_not_be_empty(cls, value: list[ProviderAPI]) -> list[ProviderAPI]:
        deduped = list(dict.fromkeys(value))
        if not deduped:
            raise ValueError("supported_apis must not be empty")
        return deduped

    def to_profile(self) -> AIProfile:
        if self.default_api not in self.supported_apis:
            raise AIProfileError(
                f"AI profile {self.id} default_api must be included in supported_apis"
            )

        return AIProfile(
            id=self.id,
            label=self.label,
            api_base_url=self.api_base_url,
            api_key=os.getenv(self.api_key_env, ""),
            api_key_env=self.api_key_env,
            model=self.model,
            default_api=self.default_api,
            supported_apis=tuple(self.supported_apis),
        )


@dataclass(frozen=True)
class AIProfilePublic:
    id: str
    label: str
    model: str
    default_api: ProviderAPI
    supported_apis: tuple[ProviderAPI, ...]
    configured: bool


def list_ai_profiles(settings: Settings | None = None) -> list[AIProfile]:
    settings = settings or get_settings()
    if not settings.ai_profiles_json.strip():
        return [_default_profile(settings)]

    try:
        parsed = json.loads(settings.ai_profiles_json)
    except json.JSONDecodeError as exc:
        raise AIProfileError("AI_PROFILES_JSON must be valid JSON") from exc

    if not isinstance(parsed, list):
        raise AIProfileError("AI_PROFILES_JSON must be a JSON array")

    profiles: list[AIProfile] = []
    seen_ids: set[str] = set()
    try:
        configs = [AIProfileConfig.model_validate(item) for item in parsed]
    except ValidationError as exc:
        raise AIProfileError("AI_PROFILES_JSON profile entries are invalid") from exc

    for config in configs:
        if config.id in seen_ids:
            raise AIProfileError(f"AI profile id is duplicated: {config.id}")
        seen_ids.add(config.id)
        profiles.append(config.to_profile())

    if not profiles:
        raise AIProfileError("AI_PROFILES_JSON must include at least one profile")

    return profiles


def list_public_ai_profiles(settings: Settings | None = None) -> list[AIProfilePublic]:
    return [
        AIProfilePublic(
            id=profile.id,
            label=profile.label,
            model=profile.model,
            default_api=profile.default_api,
            supported_apis=profile.supported_apis,
            configured=bool(profile.api_key),
        )
        for profile in list_ai_profiles(settings)
    ]


def resolve_ai_profile(
    settings: Settings | None = None,
    ai_profile_id: str | None = None,
) -> AIProfile:
    profiles = list_ai_profiles(settings)
    requested_id = ai_profile_id or profiles[0].id

    for profile in profiles:
        if profile.id == requested_id:
            return profile

    raise AIProfileError(f"AI profile not found: {requested_id}")


def resolve_provider_api(profile: AIProfile, provider_api: ProviderAPI | None = None) -> ProviderAPI:
    resolved = provider_api or profile.default_api
    if resolved not in profile.supported_apis:
        raise AIProfileError(
            f"AI profile {profile.id} does not support provider_api={resolved}"
        )
    return resolved


def _default_profile(settings: Settings) -> AIProfile:
    default_api = _resolve_default_api(settings.ai_provider_api)
    return AIProfile(
        id=DEFAULT_PROFILE_ID,
        label="Default AI (.env)",
        api_base_url=settings.openai_api_base_url,
        api_key=settings.ai_api_key,
        api_key_env="AI_API_KEY",
        model=settings.openai_model,
        default_api=default_api,
        supported_apis=("responses", "chat"),
    )


def _resolve_default_api(value: str) -> ProviderAPI:
    if value in {"responses", "chat"}:
        return value
    raise AIProfileError("AI_PROVIDER_API must be responses or chat")
