from __future__ import annotations

import json
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field, model_validator


SlideKindName = Literal["title", "section", "content", "two_column", "table", "image", "metrics", "chart", "closing"]


class LayoutBinding(BaseModel):
    layout_index: int | None = None
    layout_name: str | None = None
    title_placeholder_idx: int | None = 0
    subtitle_placeholder_idx: int | None = None
    body_placeholder_idxs: list[int] = Field(default_factory=list)
    picture_placeholder_idx: int | None = None
    footer_placeholder_idx: int | None = None
    slide_number_placeholder_idx: int | None = None
    notes_placeholder_idx: int | None = None
    use_template_footer: bool = False

    @model_validator(mode="after")
    def validate_layout_reference(self) -> "LayoutBinding":
        if self.layout_index is None and self.layout_name is None:
            raise ValueError("layout binding requires either 'layout_index' or 'layout_name'")
        return self


class TemplateConfig(BaseModel):
    slide_layouts: dict[SlideKindName, LayoutBinding]

    def binding_for(self, slide_kind: str) -> LayoutBinding | None:
        return self.slide_layouts.get(slide_kind)


def load_template_config(path: str | Path | None) -> TemplateConfig | None:
    if not path:
        return None
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    return TemplateConfig.model_validate(payload)
