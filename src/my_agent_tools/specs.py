from __future__ import annotations

import json
from math import ceil
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, Field, model_validator


SlideKind = Literal["title", "section", "content", "two_column", "table", "image", "metrics", "chart", "closing"]
BlockType = Literal["paragraph", "bullet_list", "table", "image", "code", "metrics", "chart"]
ChartType = Literal["column_clustered", "bar_clustered", "line_marker", "pie"]


class MetricSpec(BaseModel):
    label: str
    value: str
    delta: str | None = None


class ChartSeriesSpec(BaseModel):
    name: str
    values: list[float]


class ChartSpec(BaseModel):
    chart_type: ChartType = "column_clustered"
    categories: list[str]
    series: list[ChartSeriesSpec]
    title: str | None = None

    @model_validator(mode="after")
    def validate_series(self) -> "ChartSpec":
        if not self.categories:
            raise ValueError("chart block requires 'categories'")
        if not self.series:
            raise ValueError("chart block requires at least one series")
        expected = len(self.categories)
        for series in self.series:
            if len(series.values) != expected:
                raise ValueError(
                    f"chart series '{series.name}' length {len(series.values)} does not match category count {expected}"
                )
        return self


class BlockSpec(BaseModel):
    type: BlockType
    text: str | None = None
    items: list[str] | None = None
    columns: list[str] | None = None
    rows: list[list[Any]] | None = None
    path: str | None = None
    alt: str | None = None
    language: str | None = None
    content: str | None = None
    metrics: list[MetricSpec] | None = None
    chart: ChartSpec | None = None

    @model_validator(mode="after")
    def validate_payload(self) -> "BlockSpec":
        if self.type == "paragraph" and not self.text:
            raise ValueError("paragraph block requires 'text'")
        if self.type == "bullet_list" and not self.items:
            raise ValueError("bullet_list block requires 'items'")
        if self.type == "table":
            if not self.columns or self.rows is None:
                raise ValueError("table block requires 'columns' and 'rows'")
        if self.type == "image" and not self.path:
            raise ValueError("image block requires 'path'")
        if self.type == "code" and not self.content:
            raise ValueError("code block requires 'content'")
        if self.type == "metrics" and not self.metrics:
            raise ValueError("metrics block requires 'metrics'")
        if self.type == "chart" and not self.chart:
            raise ValueError("chart block requires 'chart'")
        return self

    def estimated_line_units(self) -> int:
        if self.type == "paragraph":
            return max(1, ceil(len(self.text or "") / 80)) + 1
        if self.type == "bullet_list":
            return sum(max(1, ceil(len(item) / 70)) for item in (self.items or []))
        if self.type == "code":
            return max(3, ceil(len(self.content or "") / 55))
        if self.type == "table":
            return len(self.rows or []) + 2
        if self.type == "image":
            return 5
        if self.type == "metrics":
            return 4
        if self.type == "chart":
            return 5
        return 1


class SlideSpec(BaseModel):
    title: str
    kind: SlideKind = "content"
    subtitle: str | None = None
    blocks: list[BlockSpec] = Field(default_factory=list)
    speaker_notes: str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="after")
    def validate_blocks(self) -> "SlideSpec":
        if self.kind in {"title", "section"}:
            return self
        if self.kind == "two_column" and len(self.blocks) < 2:
            raise ValueError("two_column slide requires at least two blocks")
        if self.kind == "table" and not any(block.type == "table" for block in self.blocks):
            raise ValueError("table slide requires a table block")
        if self.kind == "image" and not any(block.type == "image" for block in self.blocks):
            raise ValueError("image slide requires an image block")
        if self.kind == "metrics" and not any(block.type == "metrics" for block in self.blocks):
            raise ValueError("metrics slide requires a metrics block")
        if self.kind == "chart" and not any(block.type == "chart" for block in self.blocks):
            raise ValueError("chart slide requires a chart block")
        return self


class DeckMeta(BaseModel):
    title: str
    subtitle: str | None = None
    author: str | None = None
    company: str | None = None
    language: str = "zh-CN"
    ratio: Literal["16:9", "4:3"] = "16:9"


class ThemeRef(BaseModel):
    name: str = "corporate-clean"
    path: str | None = None


class DeckConstraints(BaseModel):
    max_slides: int | None = None
    require_speaker_notes: bool = False
    template_path: str | None = None
    template_config_path: str | None = None
    max_text_line_units_per_slide: int = 12
    max_chars_per_text_slide: int = 520
    max_table_rows_per_slide: int = 12


class DeckSpec(BaseModel):
    meta: DeckMeta
    theme: ThemeRef = Field(default_factory=ThemeRef)
    constraints: DeckConstraints = Field(default_factory=DeckConstraints)
    slides: list[SlideSpec]

    @model_validator(mode="after")
    def validate_constraints(self) -> "DeckSpec":
        if self.constraints.max_slides is not None and len(self.slides) > self.constraints.max_slides:
            raise ValueError("slide count exceeds constraints.max_slides")
        if self.constraints.require_speaker_notes:
            missing = [slide.title for slide in self.slides if not slide.speaker_notes]
            if missing:
                raise ValueError(f"speaker notes required but missing on slides: {', '.join(missing)}")
        return self

    @classmethod
    def from_path(cls, path: str | Path) -> "DeckSpec":
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls.model_validate(raw)
