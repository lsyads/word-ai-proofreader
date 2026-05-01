from app.settings import Settings


def test_settings_defaults():
    settings = Settings()

    assert settings.ai_api_key == ""
    assert settings.ai_provider_api == "responses"
    assert settings.openai_api_base_url == "https://api.openai.com/v1"
    assert settings.openai_model == "gpt-4o-mini"
    assert settings.ai_request_timeout_seconds == 30
    assert settings.ai_max_tokens == 32768
    assert settings.ai_fast_max_tokens == 8192
    assert settings.ai_thinking_max_tokens == 16384
    assert settings.backend_cors_origins == "https://localhost:3000,http://localhost:3000"
    assert settings.backend_cors_origin_list == ["https://localhost:3000", "http://localhost:3000"]
    assert settings.backend_log_level == "INFO"


def test_settings_parses_comma_separated_cors_origins():
    settings = Settings(BACKEND_CORS_ORIGINS="https://localhost:3000, http://localhost:3000")

    assert settings.backend_cors_origin_list == ["https://localhost:3000", "http://localhost:3000"]


def test_settings_reads_comma_separated_cors_origins_from_env(monkeypatch):
    monkeypatch.setenv("BACKEND_CORS_ORIGINS", "https://localhost:3000,http://localhost:3000")

    settings = Settings()

    assert settings.backend_cors_origin_list == ["https://localhost:3000", "http://localhost:3000"]
