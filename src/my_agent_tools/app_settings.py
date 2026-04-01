from __future__ import annotations

import json
from pathlib import Path

from pydantic import BaseModel


PROVIDER_PRESETS: dict[str, dict[str, str]] = {
    "custom": {
        "label": "自定义兼容接口",
        "model": "",
        "base_url": "",
    },
    "openai": {
        "label": "官方 OpenAI",
        "model": "gpt-5-mini",
        "base_url": "https://api.openai.com/v1",
    },
    "gateway": {
        "label": "本地代理 / 中转网关",
        "model": "",
        "base_url": "",
    },
}


class DesktopPlannerSettings(BaseModel):
    provider_preset: str = "custom"
    model: str = "gpt-5-mini"
    base_url: str = ""
    api_key: str = ""
    remember_api_key: bool = False


def load_desktop_settings(path: str | Path) -> DesktopPlannerSettings:
    target = Path(path)
    if not target.exists():
        return DesktopPlannerSettings()
    return DesktopPlannerSettings.model_validate_json(target.read_text(encoding="utf-8"))


def save_desktop_settings(path: str | Path, settings: DesktopPlannerSettings) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(
        json.dumps(settings.model_dump(mode="json"), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
