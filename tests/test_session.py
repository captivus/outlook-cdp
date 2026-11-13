"""Choosing the Outlook tab. No browser: `targets()` is stubbed with chrome-agent's status listing."""
import os
import sys
from contextlib import contextmanager

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp.session import OutlookError, Session  # noqa: E402

TABS = [
    {"id": "0671D8D3", "full_id": "0671D8D3AAAA", "url": "https://dashboard.example.com/", "title": "Other"},
    {"id": "2F214D46", "full_id": "2F214D46BBBB", "url": "https://outlook.cloud.microsoft/mail/",
     "title": "Mail - A - Outlook"},
    {"id": "FC24BF81", "full_id": "FC24BF81CCCC", "url": "https://outlook.cloud.microsoft/mail/",
     "title": "Mail - B - Outlook"},
]


class Stub(Session):
    def __init__(self, tabs, **kw):
        super().__init__("inst", **kw)
        self._tabs = tabs

    def targets(self):
        return self._tabs


@contextmanager
def raises(exc, needle=""):
    try:
        yield
    except exc as got:
        assert needle in str(got), f"{got!r} does not mention {needle!r}"
    else:
        raise AssertionError(f"expected {exc.__name__}")


def test_two_outlook_tabs_and_no_target_is_refused_with_the_choices_listed():
    for needle in ("2 Outlook tabs", "--target", "2F214D46", "FC24BF81", "Mail - B"):
        with raises(OutlookError, needle):
            Stub(TABS).host


def test_a_target_prefix_picks_one_tab_case_insensitively():
    s = Stub(TABS, target="fc24bf81")
    assert s.target == "FC24BF81CCCC"
    assert s.host == "outlook.cloud.microsoft"


def test_a_lone_outlook_tab_is_chosen_and_pinned_by_id():
    s = Stub(TABS[:2])
    assert s.target == "2F214D46BBBB"
    assert s._target == "2F214D46BBBB"


def test_an_unknown_target_is_refused():
    with raises(OutlookError, "matches 0 tabs"):
        Stub(TABS, target="DEADBEEF").tab()


def test_a_non_outlook_target_is_refused():
    with raises(OutlookError, "not an Outlook tab"):
        Stub(TABS, target="0671").host


def test_no_outlook_tab_is_refused():
    with raises(OutlookError, "no Outlook tab"):
        Stub(TABS[:1]).tab()


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
