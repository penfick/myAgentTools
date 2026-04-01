# Architecture

## Goal

Build a local, deterministic toolchain for an agent that can eventually generate multiple artifact types while keeping `PPT` as the highest-quality path first.

## Layers

### 1. Planner layer

The LLM should convert user requests into structured specs instead of trying to author output formats directly.

Primary responsibility:

- infer slide/page structure,
- normalize business requirements,
- decide which local tool to call,
- emit valid JSON for the target schema.

### 2. Shared schema layer

`src/my_agent_tools/specs.py`

This module defines the contract between the planner and renderers:

- `DeckSpec`
- `SlideSpec`
- `BlockSpec`
- `MetricSpec`

The schema is intentionally presentation-first right now because PPT is the priority. Future Markdown and HTML generators can either:

- reuse `DeckSpec` directly for deck-like documents, or
- introduce a higher-level `DocumentSpec` that compiles into `DeckSpec` when a slide deck is needed.

### 3. Renderer / tool layer

`src/my_agent_tools/tools/`

- `ppt.py`: implemented
- `md.py`: reserved interface
- `html.py`: reserved interface

Each tool is deterministic and file-system based. The LLM should never be allowed to bypass these tools when the output must be stable.

### 4. CLI / execution layer

`src/my_agent_tools/cli.py`

This is the local execution surface for both manual use and future agent integration.

## PPT design

### Current strengths

- JSON spec input
- speaker notes support
- theme abstraction
- optional template input
- optional template config input
- multiple slide kinds with explicit render policies
- responsive scaling for `16:9` and `4:3`
- chart slide rendering
- automatic text continuation slides
- automatic table pagination
- footer page numbering
- template-layout and placeholder-driven rendering

### Current slide kinds

- `title`
- `section`
- `content`
- `two_column`
- `table`
- `image`
- `metrics`
- `closing`

### Why this design

- Layout logic is in code, not in prompts.
- The spec describes meaning; the renderer decides placement.
- Theme decisions are centralized instead of repeated slide by slide.

## Recommended next PPT upgrades

1. Add brand-template onboarding utilities for custom `.pptx` assets and mapping extraction.
2. Improve overflow detection from heuristic pagination to measured layout fitting against actual text frames.
3. Add richer chart styling and more chart families.
4. Add image fit modes: contain, cover, crop focal point.
5. Add export-to-PDF verification through local Office or LibreOffice.

## Future multi-format plan

When Markdown and HTML are added, keep the interface pattern the same:

- planner emits structured spec
- renderer writes file
- tool returns structured result

That keeps the agent orchestration stable even as output types expand.
