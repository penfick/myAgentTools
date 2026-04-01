from __future__ import annotations

import json
import os
from typing import Any, TypeVar

from pydantic import BaseModel

from my_agent_tools.ai_models import OutlinePlan, PlannerSettings
from my_agent_tools.specs import DeckSpec


SchemaT = TypeVar("SchemaT", bound=BaseModel)

OUTLINE_PROMPT = """你是一个严格的 PPT 规划助手。
你的职责不是直接生成最终 PPT，而是先判断信息是否足够。

规则：
1. 如果用户需求缺少关键条件，必须返回 `needs_clarification=true`，并提出 1 到 3 个明确问题。
2. 如果信息足够，返回 `needs_clarification=false`，再给出可确认的大纲。
3. 不要假装知道用户没有提供的业务事实。
4. 页面类型只能从系统允许的 slide kind 中选择。
5. 大纲要适合商务 PPT，不要写成长文章。
6. 只输出结构化 JSON。"""

DECKSPEC_PROMPT = """你是一个严格的 PPT 规格生成助手。
你的任务是根据已经确认的大纲生成最终 `DeckSpec`。

规则：
1. 必须严格遵守用户已经确认的大纲结构。
2. 每页都应有清晰的 `title` 和合理的 `kind`。
3. `metrics` 页必须使用 metrics block。
4. `chart` 页必须使用 chart block。
5. `table` 页必须使用 table block。
6. 没有真实图片路径时，不要凭空捏造 image block。
7. 尽量为每页提供 `speaker_notes`，便于后续汇报。
8. 只输出结构化 JSON。"""

JSON_ONLY_HINT = "只返回 JSON 对象，不要 Markdown 代码块，不要附加解释。"


class PlannerError(RuntimeError):
    pass


class OpenAIPlanner:
    def __init__(self, settings: PlannerSettings) -> None:
        self.settings = settings

    def _client(self):
        api_key = self.settings.api_key or os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise PlannerError("未找到 API Key。请在界面中填写，或设置环境变量 OPENAI_API_KEY。")

        base_url = self.settings.base_url or os.getenv("OPENAI_BASE_URL")
        try:
            from openai import OpenAI
        except ImportError as exc:
            raise PlannerError("当前环境未安装 openai SDK。请先安装项目依赖。") from exc

        client_kwargs: dict[str, Any] = {"api_key": api_key}
        if base_url:
            client_kwargs["base_url"] = base_url
        return OpenAI(**client_kwargs)

    def generate_outline(self, requirement: str, feedback: str | None = None) -> OutlinePlan:
        payload = {
            "requirement": requirement,
            "feedback": feedback or "",
        }
        try:
            return self._request_structured(OUTLINE_PROMPT, payload, OutlinePlan)
        except Exception as exc:
            if isinstance(exc, PlannerError):
                raise
            raise PlannerError(f"生成大纲失败：{exc}") from exc

    def generate_deck_spec(self, requirement: str, outline: OutlinePlan, feedback: str | None = None) -> DeckSpec:
        if outline.needs_clarification:
            raise PlannerError("当前大纲仍处于待确认状态，不能直接生成 DeckSpec。")

        payload = {
            "requirement": requirement,
            "feedback": feedback or "",
            "confirmed_outline": outline.model_dump(mode="json"),
        }
        try:
            return self._request_structured(DECKSPEC_PROMPT, payload, DeckSpec)
        except Exception as exc:
            if isinstance(exc, PlannerError):
                raise
            raise PlannerError(f"生成 DeckSpec 失败：{exc}") from exc

    def _request_structured(self, system_prompt: str, payload: dict[str, Any], schema: type[SchemaT]) -> SchemaT:
        client = self._client()
        messages = self._messages(system_prompt, payload)
        failures: list[str] = []

        try:
            response = client.responses.parse(
                model=self.settings.model,
                input=messages,
                text_format=schema,
            )
            if response.output_parsed is not None:
                return response.output_parsed
            failures.append("responses.parse 未返回结构化结果")
        except Exception as exc:
            failures.append(f"responses.parse: {exc}")

        try:
            response = client.chat.completions.create(
                model=self.settings.model,
                messages=self._chat_messages(system_prompt, payload, force_json=True),
                response_format={"type": "json_object"},
            )
            content = self._extract_chat_content(response)
            return self._parse_model_json(content, schema)
        except Exception as exc:
            failures.append(f"chat.completions.create(json_object): {exc}")

        try:
            response = client.chat.completions.create(
                model=self.settings.model,
                messages=self._chat_messages(system_prompt, payload, force_json=True),
            )
            content = self._extract_chat_content(response)
            return self._parse_model_json(content, schema)
        except Exception as exc:
            failures.append(f"chat.completions.create(text): {exc}")

        failure_text = "\n".join(f"- {item}" for item in failures)
        raise PlannerError(f"结构化生成失败。请检查模型兼容性、API Base URL 或返回格式。\n{failure_text}")

    def _messages(self, system_prompt: str, payload: dict[str, Any]) -> list[dict[str, str]]:
        return [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ]

    def _chat_messages(self, system_prompt: str, payload: dict[str, Any], force_json: bool) -> list[dict[str, str]]:
        prompt = system_prompt if not force_json else f"{system_prompt}\n\n{JSON_ONLY_HINT}"
        return [
            {"role": "system", "content": prompt},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ]

    def _extract_chat_content(self, response: Any) -> str:
        try:
            content = response.choices[0].message.content
        except Exception as exc:
            raise PlannerError(f"无法从聊天补全响应中提取内容：{exc}") from exc

        if isinstance(content, str):
            return content
        if isinstance(content, list):
            parts: list[str] = []
            for item in content:
                if isinstance(item, dict) and item.get("type") == "text":
                    parts.append(item.get("text", ""))
                elif hasattr(item, "text"):
                    parts.append(item.text)
            return "".join(parts)
        raise PlannerError("聊天补全响应内容格式无法识别。")

    def _parse_model_json(self, content: str, schema: type[SchemaT]) -> SchemaT:
        cleaned = content.strip()
        if cleaned.startswith("```"):
            cleaned = self._strip_code_fences(cleaned)

        try:
            return schema.model_validate_json(cleaned)
        except Exception:
            pass

        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start != -1 and end != -1 and end > start:
            return schema.model_validate_json(cleaned[start : end + 1])

        raise PlannerError("模型返回的内容不是有效 JSON。")

    def _strip_code_fences(self, content: str) -> str:
        lines = content.strip().splitlines()
        if not lines:
            return content
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        return "\n".join(lines).strip()
