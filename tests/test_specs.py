from my_agent_tools.specs import DeckSpec


def test_sample_deck_is_valid():
    spec = DeckSpec.from_path("examples/sample_deck.json")
    assert spec.meta.title == "2026 Q2 Operating Review"
    assert len(spec.slides) == 8
    assert spec.constraints.require_speaker_notes is True
