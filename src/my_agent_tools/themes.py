from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Theme:
    name: str
    font_family: str
    title_font_size: int
    subtitle_font_size: int
    body_font_size: int
    small_font_size: int
    code_font_size: int
    background_color: str
    title_color: str
    text_color: str
    muted_text_color: str
    accent_color: str
    accent_soft_color: str
    border_color: str
    card_background: str


DEFAULT_THEMES: dict[str, Theme] = {
    "corporate-clean": Theme(
        name="corporate-clean",
        font_family="Microsoft YaHei",
        title_font_size=28,
        subtitle_font_size=15,
        body_font_size=18,
        small_font_size=11,
        code_font_size=12,
        background_color="F7F8FA",
        title_color="162033",
        text_color="263246",
        muted_text_color="5E6B80",
        accent_color="1F6FEB",
        accent_soft_color="DCE9FF",
        border_color="D9E0EA",
        card_background="FFFFFF",
    )
}


def load_theme(name: str = "corporate-clean", custom_path: str | None = None) -> Theme:
    if custom_path:
        payload = json.loads(Path(custom_path).read_text(encoding="utf-8"))
        return Theme(**payload)
    try:
        return DEFAULT_THEMES[name]
    except KeyError as exc:
        available = ", ".join(sorted(DEFAULT_THEMES))
        raise ValueError(f"unknown theme '{name}'. Available: {available}") from exc
