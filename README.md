# My Agent Tools

Python-first local agent toolchain scaffold. The architecture is organized around:

- a shared deck/document schema,
- deterministic local tools,
- a PPT-first renderer implemented with `python-pptx`,
- extension points for future Markdown and HTML generators.

## Current state

- `PPT` generation is implemented and usable.
- Template-config-driven PPT rendering is supported.
- A confirmation-first AI planning workflow is available in the desktop GUI.
- `Markdown` and `HTML` generators are defined as stubs so the interface is stable before those features are added.
- `CLI` entrypoints are included for local execution.

## Project layout

```text
myAgentTools/
  examples/
    sample_deck.json
    default_template_config.json
  src/
    my_agent_tools/
      cli.py
      ai_models.py
      openai_planner.py
      specs.py
      template_config.py
      themes.py
      tools/
        base.py
        html.py
        md.py
        ppt.py
  tests/
    test_specs.py
```

## Quick start

1. Install the package in editable mode:

```powershell
pip install -e .[dev]
```

2. Render the sample PPT:

```powershell
python -m my_agent_tools.cli render-ppt --spec .\examples\sample_deck.json --output .\out\sample_deck.pptx
```

3. Inspect the generated presentation under `out/sample_deck.pptx`.

Optional: render with the built-in template mapping config:

```powershell
python -m my_agent_tools.cli render-ppt --spec .\examples\sample_deck.json --output .\out\sample_deck_v3.pptx --template-config .\examples\default_template_config.json
```

4. Launch the desktop GUI:

```powershell
pythonw .\launch_gui.pyw
```

Or just double-click `launch_gui.bat`.

## Desktop GUI

The local desktop interface supports:

- entering natural-language requirements,
- configuring `Model`, `API Base URL`, and `API Key`,
- choosing interface presets and saving planner settings locally,
- generating an AI outline first,
- asking the user for clarifications before deck generation,
- showing pending clarification questions in the GUI and writing an answer template back into the feedback box,
- confirming or editing the outline before final spec generation,
- loading the sample deck spec,
- opening and saving JSON specs,
- editing deck JSON directly,
- validating the spec,
- browsing slides from a left-side outline,
- previewing the selected slide in a native desktop window,
- generating a `.pptx` file from the current editor content.

## AI workflow

The desktop GUI now follows a confirmation-first sequence:

1. Enter the user's requirement in the `AI 规划` tab.
2. Fill `API Key`, and optionally fill `API Base URL`, or set `OPENAI_API_KEY` / `OPENAI_BASE_URL`.
3. Optionally choose an interface preset and save the planner settings.
4. Click `生成大纲`.
5. If the model asks clarification questions, review them in the pending-questions area, write answers into `补充说明 / 用户修正`, and click `结合补充重生成大纲`.
6. Review or edit the outline JSON manually if needed.
7. Click `确认大纲并生成规格`.
8. Review the generated spec and only then click `生成 PPT`.

Default AI model:

- `gpt-5-mini`

You can override it in the GUI model field.

API compatibility notes:

- The planner tries OpenAI `Responses API` first.
- If the target provider only supports OpenAI-compatible `Chat Completions`, it will automatically fall back.
- This makes third-party OpenAI-compatible gateways easier to use through `API Base URL`.

## Design principles

- LLMs should produce structured specs, not raw slide coordinates.
- The renderer owns layout decisions, theme enforcement, and file generation.
- Future generators should reuse the same high-level content schema where possible.
- When precision matters, theme and layout policies should be encoded in code and template assets instead of prompts.

## PPT roadmap

The current renderer already supports:

- title, section, content, two-column, metrics, table, chart, and closing slides,
- speaker notes,
- theme-driven typography and colors,
- optional `.pptx` template loading,
- optional layout and placeholder mapping from JSON config,
- automatic text continuation and table pagination,
- JSON-based deck specs.

The next PPT upgrades should focus on:

- brand-template onboarding,
- richer image fitting policies,
- more precise measured layout fitting,
- brand asset packs and logos.
