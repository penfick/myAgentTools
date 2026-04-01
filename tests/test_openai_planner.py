from types import SimpleNamespace

from my_agent_tools.ai_models import OutlinePlan, PlannerSettings
from my_agent_tools.openai_planner import OpenAIPlanner


class _FakeResponses:
    def parse(self, **_kwargs):
        raise RuntimeError("responses endpoint unsupported")


class _FakeChatCompletions:
    def __init__(self, content: str) -> None:
        self.content = content

    def create(self, **_kwargs):
        return SimpleNamespace(
            choices=[
                SimpleNamespace(
                    message=SimpleNamespace(
                        content=self.content,
                    )
                )
            ]
        )


class _FakeClient:
    def __init__(self, content: str) -> None:
        self.responses = _FakeResponses()
        self.chat = SimpleNamespace(completions=_FakeChatCompletions(content))


def test_generate_outline_falls_back_to_chat_completions(monkeypatch):
    planner = OpenAIPlanner(PlannerSettings(model="test-model", api_key="sk-test", base_url="https://example.com/v1"))
    monkeypatch.setattr(planner, "_client", lambda: _FakeClient(
        """```json
{
  "needs_clarification": false,
  "clarification_questions": [],
  "deck_goal": "说明方案",
  "target_audience": "管理层",
  "tone": "专业",
  "estimated_slides": 4,
  "slides": [
    {
      "title": "项目背景",
      "kind": "content",
      "objective": "说明背景",
      "bullets": ["背景一", "背景二"],
      "notes_hint": "先讲背景"
    }
  ]
}
```"""
    ))

    outline = planner.generate_outline("做一个项目汇报")

    assert outline.needs_clarification is False
    assert outline.estimated_slides == 4
    assert outline.slides[0].title == "项目背景"


def test_parse_model_json_accepts_wrapped_json():
    planner = OpenAIPlanner(PlannerSettings(model="test-model", api_key="sk-test"))

    outline = planner._parse_model_json(
        '说明如下：{"needs_clarification":false,"clarification_questions":[],"deck_goal":"说明方案","target_audience":"管理层","tone":"专业","estimated_slides":3,"slides":[]}',
        schema=OutlinePlan,
    )

    assert outline.deck_goal == "说明方案"
