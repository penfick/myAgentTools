from __future__ import annotations

import argparse
import json
import sys

from my_agent_tools.specs import DeckSpec
from my_agent_tools.tools.html import generate_html
from my_agent_tools.tools.md import generate_markdown
from my_agent_tools.tools.ppt import generate_ppt


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="my-agent-tools", description="Local agent tools with a PPT-first renderer.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    render_ppt = subparsers.add_parser("render-ppt", help="Render a PPTX deck from a JSON spec.")
    render_ppt.add_argument("--spec", required=True, help="Path to the input deck spec JSON file.")
    render_ppt.add_argument("--output", required=True, help="Path to the output .pptx file.")
    render_ppt.add_argument("--template", help="Optional .pptx template file.")
    render_ppt.add_argument("--template-config", help="Optional JSON template mapping config file.")
    render_ppt.add_argument("--theme", help="Optional custom theme JSON file.")

    render_md = subparsers.add_parser("render-md", help="Reserved interface for future Markdown generation.")
    render_md.add_argument("--spec", required=True)
    render_md.add_argument("--output", required=True)

    render_html = subparsers.add_parser("render-html", help="Reserved interface for future HTML generation.")
    render_html.add_argument("--spec", required=True)
    render_html.add_argument("--output", required=True)

    inspect_spec = subparsers.add_parser("inspect-spec", help="Validate and summarize a deck spec.")
    inspect_spec.add_argument("--spec", required=True)

    return parser


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    spec = DeckSpec.from_path(args.spec)

    if args.command == "render-ppt":
        result = generate_ppt(
            spec,
            output_path=args.output,
            template_path=args.template,
            theme_path=args.theme,
            template_config_path=args.template_config,
        )
    elif args.command == "render-md":
        result = generate_markdown(spec, output_path=args.output)
    elif args.command == "render-html":
        result = generate_html(spec, output_path=args.output)
    elif args.command == "inspect-spec":
        payload = {
            "title": spec.meta.title,
            "slides": len(spec.slides),
            "slide_titles": [slide.title for slide in spec.slides],
            "theme": spec.theme.name,
            "ratio": spec.meta.ratio,
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    else:
        raise ValueError(f"Unknown command: {args.command}")

    payload = result.model_dump()
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if result.ok else 1


if __name__ == "__main__":
    sys.exit(main())
