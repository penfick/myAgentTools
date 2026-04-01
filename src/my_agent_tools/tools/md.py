from __future__ import annotations

from pathlib import Path

from my_agent_tools.specs import DeckSpec
from my_agent_tools.tools.base import ToolResult, ensure_parent_dir


def generate_markdown(spec: DeckSpec, output_path: str | Path) -> ToolResult:
    """Stable interface placeholder for the future Markdown renderer."""
    ensure_parent_dir(output_path)
    return ToolResult(
        ok=False,
        message="Markdown generation is not implemented yet. PPT is the only completed renderer in this scaffold.",
        output_path=str(output_path),
        meta={"planned": True, "slides_defined": len(spec.slides)},
    )
