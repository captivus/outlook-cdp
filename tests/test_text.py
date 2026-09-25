"""Text handling, which is the only part testable without a browser."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp.text import normalise, to_html  # noqa: E402


def test_wrapped_prose_is_rejoined():
    text = ("This is a letter line that runs well past sixty characters before it wraps,\n"
            "and this is the continuation of the same paragraph after the wrap.")
    html = to_html(text)
    assert "before it wraps, and this is the continuation" in html


def test_the_sixty_character_heuristic_is_deliberate_and_has_a_cost():
    """A block whose every line is short is treated as an address and keeps its breaks.

    That is right for a signature and wrong for short-line prose, which will not be rejoined. It is a
    heuristic, and this test exists so the limitation is visible rather than surprising.
    """
    short_prose = "Thank you for your reply.\nWe accept the narrowed scope."
    assert "<br>" in to_html(short_prose)          # not rejoined - the known cost
    long_prose = ("Thank you for your reply, and we are content to accept the narrowed scope "
                  "you proposed,\nsince it still covers the unit counts we asked about.")
    assert "you proposed, since it still covers" in to_html(long_prose)


def test_short_lines_keep_their_breaks():
    html = to_html("Example Organization\nPO Box 11\nExample City, ST 12345")
    # one block, so the only paragraph separator is the trailing one; the two joins are the line breaks
    assert "Example Organization<br>PO Box 11<br>Example City, ST 12345" in html


def test_list_items_stay_on_their_own_lines():
    html = to_html("We request:\n\n  1. The register\n  2. The owner email\n  3. The unit count")
    assert "1. The register<br>2. The owner email<br>3. The unit count" in html


def test_long_list_items_are_not_rejoined_into_prose():
    """The branch that matters: items over 60 characters still keep one line each."""
    text = ("We request the following:\n\n"
            "  1. The register of every licensed rental property, including the address of each\n"
            "  2. The name of every owner recorded against each property, company or natural person")
    html = to_html(text)
    assert "address of each<br>2. The name of every owner" in html


def test_html_is_escaped():
    assert "&lt;script&gt;" in to_html("a <script> tag")


def test_normalise_collapses_whitespace():
    assert normalise("a  b\n\tc ") == "a b c"


def test_normalise_handles_none():
    assert normalise(None) == ""


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
