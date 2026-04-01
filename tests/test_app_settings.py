from pathlib import Path
from uuid import uuid4

from my_agent_tools.app_settings import DesktopPlannerSettings, load_desktop_settings, save_desktop_settings


def _settings_output(name: str) -> Path:
    output_dir = Path("out") / "test_outputs"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / f"{name}_{uuid4().hex}.json"


def test_load_desktop_settings_returns_default_when_missing():
    settings = load_desktop_settings(_settings_output("missing"))

    assert settings.provider_preset == "custom"
    assert settings.model == "gpt-5-mini"
    assert settings.remember_api_key is False


def test_save_desktop_settings_omits_api_key_when_not_remembered():
    path = _settings_output("desktop_settings")
    settings = DesktopPlannerSettings(
        provider_preset="openai",
        model="gpt-5-mini",
        base_url="https://api.openai.com/v1",
        api_key="sk-secret",
        remember_api_key=False,
    )

    save_desktop_settings(path, settings.model_copy(update={"api_key": ""}))
    loaded = load_desktop_settings(path)

    assert loaded.base_url == "https://api.openai.com/v1"
    assert loaded.api_key == ""
