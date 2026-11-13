"""Moving messages. No browser: a small fake mailbox answers the page-side REST calls."""
import json
import os
import re
import sys
import time
import urllib.parse
from contextlib import contextmanager

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp import attachments as A  # noqa: E402
from outlook_cdp.moving import Mover  # noqa: E402
from outlook_cdp.session import OutlookError  # noqa: E402

FOLDERS = {"inbox": "F-IN", "junkemail": "F-JUNK", "deleteditems": "F-DEL"}
NAMES = {"F-IN": "Inbox", "F-JUNK": "Junk Email", "F-DEL": "Deleted Items"}


class FakeMailbox:
    """Answers `js` like the page would: the token lookup, then same-origin REST GETs and POSTs."""

    def __init__(self, messages):
        self.msgs = {m["Id"]: dict(m) for m in messages}
        self.posts = []
        self.n = 0

    def mailbox_identity(self):
        return "me@example.com"

    def tab_url(self):
        return ""

    def js(self, expression, await_promise=False):
        if expression == A._TAB_ACCOUNT_JS:
            return None
        if "msal.cache.encryption" in expression:
            return {"secret": "jwt", "expires": int(time.time()) + 3600}
        path = json.loads(re.search(r'fetch\(("(?:[^"\\]|\\.)*")', expression).group(1))
        path = path.split("/users/me%40example.com", 1)[1]
        route, _, query = path.partition("?")
        if 'method:"POST"' in expression:
            self.posts.append(route)
            mid = urllib.parse.unquote(route.split("/")[2])
            body = json.loads(json.loads(re.search(r'body:("(?:[^"\\]|\\.)*")', expression).group(1)))
            if mid not in self.msgs:
                return {"status": 404, "body": {"error": {"code": "ErrorItemNotFound", "message": "gone"}}}
            m = self.msgs.pop(mid)
            self.n += 1
            m = dict(m, Id=f"{mid}-moved{self.n}", ParentFolderId=FOLDERS[body["DestinationId"]])
            self.msgs[m["Id"]] = m
            return {"status": 201, "body": m}
        if route.startswith("/mailfolders/") and route.count("/") == 2:
            return {"status": 200, "body": {"Id": FOLDERS[route.split("/")[2]]}}
        if route == "/mailfolders":
            return {"status": 200, "body": {"value": [{"Id": i, "DisplayName": n} for i, n in NAMES.items()]}}
        q = urllib.parse.parse_qs(query)
        if route.startswith("/messages/"):
            mid = urllib.parse.unquote(route.split("/")[2])
            if mid not in self.msgs:
                return {"status": 404, "body": {"error": {"code": "ErrorItemNotFound", "message": "gone"}}}
            return {"status": 200, "body": self.msgs[mid]}
        want = re.findall(r"Address eq '([^']+)'", q.get("$filter", [""])[0])
        folder = FOLDERS.get(route.split("/")[2]) if route.startswith("/mailfolders/") else None
        rows = [m for m in self.msgs.values() if m["From"]["EmailAddress"]["Address"] in want
                and (folder is None or m["ParentFolderId"] == folder)]
        return {"status": 200, "body": {"value": rows}}


def msg(mid, sender, folder, received, subject="test", read=False):
    return {"Id": mid, "Subject": subject, "From": {"EmailAddress": {"Address": sender, "Name": sender}},
            "ParentFolderId": FOLDERS[folder], "ReceivedDateTime": received, "IsRead": read}


MAIL = [
    msg("a1", "t@test.com", "inbox", "2026-11-12T10:00:00Z"),
    msg("a2", "t@test.com", "junkemail", "2026-11-13T01:00:00Z"),
    msg("a3", "t@test.com", "deleteditems", "2026-11-12T11:00:00Z"),   # already there
    msg("a4", "t@test.com", "inbox", "2026-11-10T09:00:00Z"),          # too old
    msg("b1", "aws@test.com", "inbox", "2026-11-11T09:00:00Z", "Email Address Verification Request"),
    msg("b2", "aws@test.com", "inbox", "2026-11-11T09:30:00Z", "Your support case"),
    msg("c1", "keep@test.com", "inbox", "2026-11-12T12:00:00Z"),
]


@contextmanager
def raises(exc, needle=""):
    try:
        yield
    except exc as got:
        assert needle in str(got), f"{got!r} does not mention {needle!r}"
    else:
        raise AssertionError(f"expected {exc.__name__}")


def ids(ms):
    return [m.id for m in ms]


def test_find_searches_every_folder_newest_first_and_skips_the_destination():
    mv = Mover(FakeMailbox(MAIL))
    assert ids(mv.find(from_addresses=["t@test.com"])) == ["a2", "a1", "a4"]


def test_find_applies_dates_subjects_and_folders():
    mv = Mover(FakeMailbox(MAIL))
    assert ids(mv.find(from_addresses=["t@test.com"], received_after="2026-11-11")) == ["a2", "a1"]
    assert ids(mv.find(from_addresses=["t@test.com"], received_after="2026-11-12",
                       received_before="2026-11-13")) == ["a1"]
    assert ids(mv.find(from_addresses=["aws@test.com"], subject_contains=["verification"])) == ["b1"]
    assert ids(mv.find(from_addresses=["t@test.com"], folders=["junkemail"])) == ["a2"]


def test_find_refuses_without_a_sender_or_ids():
    with raises(OutlookError, "--from or --ids"):
        Mover(FakeMailbox(MAIL)).find(subject_contains=["test"])


def test_find_by_ids_still_skips_what_is_already_in_the_destination():
    assert ids(Mover(FakeMailbox(MAIL)).find(ids=["a1", "a3"])) == ["a1"]


def test_move_posts_a_move_never_a_delete_and_reads_each_one_back():
    box = FakeMailbox(MAIL)
    mv = Mover(box)
    results = mv.move(mv.find(from_addresses=["t@test.com"], received_after="2026-11-11"))
    assert [r.ok for r in results] == [True, True]
    assert all(r.folder == "Deleted Items" and r.new_id and r.new_id != r.old_id for r in results)
    assert box.posts == ["/messages/a2/move", "/messages/a1/move"]
    # nothing else moved, and a moved message keeps its read state
    assert box.msgs["c1"]["ParentFolderId"] == "F-IN" and box.msgs["a4"]["ParentFolderId"] == "F-IN"
    assert not any(m["IsRead"] for m in box.msgs.values())


def test_one_failed_move_does_not_stop_the_rest():
    mv = Mover(FakeMailbox(MAIL))
    results = mv.move(["missing", "b1"])
    assert [r.ok for r in results] == [False, True]
    assert "404" in results[0].error


def test_move_refuses_more_than_the_limit():
    mv = Mover(FakeMailbox(MAIL))
    with raises(OutlookError, "limit is 1"):
        mv.move(["a1", "a2"], max_messages=1)


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
