import json

from app.services.ai_profiles import list_public_ai_profiles, resolve_ai_profile, resolve_provider_api
from app.settings import Settings


def test_default_profile_uses_legacy_env_settings():
    settings = Settings(
        AI_API_KEY="legacy-key",
        AI_PROVIDER_API="chat",
        OPENAI_API_BASE_URL="https://legacy.example/v1",
        OPENAI_MODEL="legacy-model",
    )

    profile = resolve_ai_profile(settings)

    assert profile.id == "default"
    assert profile.label == "Default AI (.env)"
    assert profile.api_base_url == "https://legacy.example/v1"
    assert profile.api_key == "legacy-key"
    assert profile.model == "legacy-model"
    assert profile.default_api == "chat"
    assert resolve_provider_api(profile, None) == "chat"


def test_profiles_json_resolves_custom_key_env(monkeypatch):
    monkeypatch.setenv("OPENROUTER_API_KEY", "openrouter-key")
    settings = Settings(
        AI_PROFILES_JSON=json.dumps(
            [
                {
                    "id": "openrouter-qwen",
                    "label": "OpenRouter / Qwen",
                    "api_base_url": "https://openrouter.ai/api/v1",
                    "api_key_env": "OPENROUTER_API_KEY",
                    "model": "qwen/test",
                    "default_api": "chat",
                    "supported_apis": ["chat"],
                }
            ]
        )
    )

    profile = resolve_ai_profile(settings, "openrouter-qwen")
    public_profile = list_public_ai_profiles(settings)[0]

    assert profile.api_key == "openrouter-key"
    assert profile.supported_apis == ("chat",)
    assert public_profile.configured is True
    assert not hasattr(public_profile, "api_key")
