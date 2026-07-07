import json
from pathlib import Path

from dotenv import dotenv_values

from app.settings import Settings


def test_env_example_keeps_code_read_environment_variables():
    env_path = Path(__file__).resolve().parents[2] / ".env.example"
    values = dotenv_values(env_path)
    settings_aliases = {field.alias for field in Settings.model_fields.values() if field.alias}

    assert settings_aliases <= set(values)


def test_env_example_includes_profile_key_env_variables():
    env_path = Path(__file__).resolve().parents[2] / ".env.example"
    values = dotenv_values(env_path)
    profiles = json.loads(values["AI_PROFILES_JSON"])
    profile_key_envs = {profile["api_key_env"] for profile in profiles}

    assert profile_key_envs <= set(values)


def test_env_example_leaves_api_key_values_blank():
    env_path = Path(__file__).resolve().parents[2] / ".env.example"
    values = dotenv_values(env_path)

    api_key_vars = [name for name in values if name.endswith("_API_KEY")]

    assert api_key_vars
    assert all(values[name] == "" for name in api_key_vars)
