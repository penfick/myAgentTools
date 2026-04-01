# Repository Guidelines

## Project Structure & Ownership
All production code lives in `src/my_agent_tools/`. Keep shared contracts in `specs.py`, planner-side models in `ai_models.py`, LLM orchestration in `openai_planner.py`, desktop workflow in `desktop_app.py`, and deterministic file writers in `src/my_agent_tools/tools/`.

`ppt.py` is the primary production path. `md.py` and `html.py` are placeholders and must not become the main focus until PPT work is stable. Inputs and sample assets belong in `examples/`; docs belong in `docs/`; generated files belong in `out/`.

## Non-Negotiable Architecture Rules
- PPT is the highest-priority feature. Do not dilute work into Markdown or HTML unless explicitly requested.
- The LLM must not generate final PPT files directly. It may only produce structured data.
- All AI output must validate against `DeckSpec` or related Pydantic models before rendering.
- Do not bypass `python-pptx` renderer logic with ad hoc slide placement from prompts.
- The desktop flow must remain confirmation-first: requirement -> outline -> user confirmation -> spec -> PPT.

## Build, Test, and Run
- `pip install -e .[dev]`: install editable package and test dependencies.
- `python -m pytest`: run all tests.
- `python -m my_agent_tools.cli render-ppt --spec .\examples\sample_deck.json --output .\out\sample_deck.pptx`: render a sample deck.
- `pythonw .\launch_gui.pyw`: start the local desktop app.
- `python -m py_compile src\my_agent_tools\desktop_app.py`: syntax-check edited modules.

## Coding Standards
Use 4-space indentation, Python type hints, and small focused functions. Use `snake_case` for modules, functions, variables, and tests; use `PascalCase` for classes. Keep UI copy explicit and workflow-oriented.

When adding features, extend the schema first, then planner logic, then renderer, then GUI. Do not invert that order.

## Testing Requirements
Use `pytest` and name files `test_*.py`. Every behavior change must include a regression test when practical, especially for schema validation, planner fallbacks, renderer output, or desktop settings. Write temporary test artifacts under `out/test_outputs/`, not system temp directories.

## Commit & PR Expectations
Git history is unavailable here, so use short imperative commit titles such as `Add provider presets to desktop app`. Keep one logical change per commit.

PRs should include a behavior summary, affected commands or files, GUI screenshots when applicable, sample PPT input/output when renderer behavior changes, and passing test evidence.

## Security & Configuration
Model access uses `API Base URL`, `API Key`, `OPENAI_BASE_URL`, and `OPENAI_API_KEY`. If local settings persistence is enabled, `config/desktop_settings.json` may store API keys in plain text. Treat that file as sensitive.
