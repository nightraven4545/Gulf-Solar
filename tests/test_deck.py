import re

from pptx import Presentation

from nrep.build_deck import build, build_onepager


def text_of(path):
    return [" ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame) for s in Presentation(path).slides]


def test_deck_builds_with_no_missing_numbers(tmp_path):
    slides = text_of(build(tmp_path / "d.pptx"))
    assert len(slides) == 18
    for i, t in enumerate(slides, start=1):
        assert not re.search(r"\bNone\b|\bnan\b", t), i


def test_onepager_is_one_slide(tmp_path):
    assert len(text_of(build_onepager(tmp_path / "o.pptx"))) == 1
