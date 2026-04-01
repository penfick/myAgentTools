from __future__ import annotations

from pydantic import BaseModel, Field

from my_agent_tools.specs import SlideKind


class ClarificationQuestion(BaseModel):
    id: str
    question: str
    reason: str | None = None


class OutlineSlidePlan(BaseModel):
    title: str
    kind: SlideKind
    objective: str
    bullets: list[str] = Field(default_factory=list)
    notes_hint: str | None = None


class OutlinePlan(BaseModel):
    needs_clarification: bool = False
    clarification_questions: list[ClarificationQuestion] = Field(default_factory=list)
    deck_goal: str
    target_audience: str
    tone: str
    estimated_slides: int
    slides: list[OutlineSlidePlan] = Field(default_factory=list)


class PlannerSettings(BaseModel):
    model: str = "gpt-5-mini"
    base_url: str | None = None
    api_key: str | None = None
