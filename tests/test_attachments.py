"""Attachment logic. No browser: a stub stands in for the one JS call each function makes.

The parts worth testing without a mailbox are the ones that decide where bytes land and whether they
are believed: filename sanitising, collision handling, magic bytes, and the two mistakes that cost
the most to find live - reading `/me/` instead of the mailbox on screen, and treating the reading
pane's URL id as a message id when it is a conversation id.
"""
import hashlib
import os
import sys
import tempfile
from contextlib import contextmanager

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp import attachments as A  # noqa: E402
from outlook_cdp.session import OutlookError  # noqa: E402


class FakeSession:
    """Answers `js` from a queue and records what it was asked, so no CDP is involved."""

    def __init__(self, answers=(), mailbox="office@example.com", url=""):
        self.answers = list(answers)
        self._mailbox = mailbox
        self._url = url
        self.asked = []

    def js(self, expression, await_promise=False):
        self.asked.append(expression)
        return self.answers.pop(0) if self.answers else None

    def mailbox_identity(self):
        return self._mailbox

    def tab_url(self):
        return self._url


@contextmanager
def raises(exc, needle=""):
    """Local stand-in for pytest.raises, so these run under bare `python3` like the other tests."""
    try:
        yield
    except exc as got:
        assert needle in str(got), f"{got!r} does not mention {needle!r}"
    else:
        raise AssertionError(f"expected {exc.__name__}")


def _write(data: bytes) -> str:
    fh = tempfile.NamedTemporaryFile(delete=False)
    fh.write(data)
    fh.close()
    return fh.name


# ------------------------------------------------------------------ filenames

def test_safe_name_cannot_escape_the_destination_directory():
    # The name is whatever the sender put in the MIME part, so it is untrusted input.
    assert A._safe_name("../../etc/passwd") == "passwd"
    assert A._safe_name("a/b/c/report.pdf") == "report.pdf"
    assert A._safe_name(r"C:\Users\x\report.pdf") == "report.pdf"


def test_safe_name_strips_control_characters_and_leading_dots():
    assert A._safe_name("re\x00po\nrt.pdf") == "re_po_rt.pdf"
    assert A._safe_name("...hidden.pdf") == "hidden.pdf"
    assert A._safe_name("-rf.pdf") == "rf.pdf"


def test_safe_name_always_returns_something_usable():
    assert A._safe_name("") == "attachment"
    assert A._safe_name("/") == "attachment"
    assert len(A._safe_name("x" * 400 + ".pdf")) <= 180


def test_unique_name_suffixes_instead_of_overwriting():
    # Three public bodies will all send you "response.pdf"; none of them may clobber another.
    with tempfile.TemporaryDirectory() as d:
        assert A._unique_name(d, "response.pdf") == "response.pdf"
        open(os.path.join(d, "response.pdf"), "wb").close()
        assert A._unique_name(d, "response.pdf") == "response-2.pdf"
        open(os.path.join(d, "response-2.pdf"), "wb").close()
        assert A._unique_name(d, "response.pdf") == "response-3.pdf"


# ------------------------------------------------------------------ believing the bytes

def test_magic_ok_accepts_a_real_spreadsheet_and_pdf():
    xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert A.magic_ok(b"PK\x03\x04rest", xlsx) is True
    assert A.magic_ok(b"%PDF-1.7\n", "application/pdf") is True


def test_magic_ok_rejects_the_wrong_bytes():
    assert A.magic_ok(b"<!DOCTYPE html>", "application/pdf") is False
    assert A.magic_ok(b"%PDF-1.7", "image/png") is False


def test_magic_ok_is_undecided_for_types_it_has_no_signature_for():
    assert A.magic_ok(b"a,b,c\n1,2,3\n", "text/csv") is None


def test_error_page_detection_catches_html_and_json():
    assert A.looks_like_an_error_page(b"\n  <!DOCTYPE html><html>") is True
    assert A.looks_like_an_error_page(b'{"error":{"code":"Unauthorized"}}') is True
    assert A.looks_like_an_error_page(b"%PDF-1.7\n1 0 obj") is False


def test_verify_file_reports_size_hash_and_magic():
    body = b"%PDF-1.7\n" + b"x" * 5000
    path = _write(body)
    try:
        got = A.verify_file(path, "application/pdf", expected_bytes=len(body))
        assert got["bytes"] == len(body)
        assert got["sha256"] == hashlib.sha256(body).hexdigest()
        assert got["magic_ok"] is True
        assert got["error_page"] is False
        assert got["size_ok"] is True
    finally:
        os.unlink(path)


def test_verify_file_flags_a_short_write():
    path = _write(b"%PDF-1.7\nshort")
    try:
        assert A.verify_file(path, "application/pdf", expected_bytes=99999)["size_ok"] is False
    finally:
        os.unlink(path)


# ------------------------------------------------------------------ scoping

def test_mailbox_base_targets_the_mailbox_on_screen_not_me():
    # `/api/v2.0/me` is the SIGNED-IN user, who for a shared mailbox is a different account. Reading
    # it returns somebody else's mail and no attachments, which reads like an API that does not work.
    base = A.mailbox_base(FakeSession(mailbox="office@example.com"))
    assert base == "/api/v2.0/users/office%40example.com"
    assert "/me" not in base


def test_mailbox_base_refuses_rather_than_guessing():
    with raises(OutlookError, "which mailbox"):
        A.mailbox_base(FakeSession(mailbox=None))


def test_open_conversation_id_is_read_from_the_tab_url_and_unescaped():
    s = FakeSession(url="https://outlook.cloud.microsoft/mail/office@example.com/inbox/id/AAQk%2Fab%3D")
    assert A.open_conversation_id(s) == "AAQk/ab="


def test_open_conversation_id_is_none_when_no_message_is_open():
    assert A.open_conversation_id(FakeSession(url="https://host/mail/office@example.com/inbox")) is None


def test_open_conversation_messages_refuses_when_nothing_is_open():
    with raises(OutlookError, "reading pane"):
        A.open_conversation_messages(FakeSession(url="https://host/mail/x@y/inbox"))


# ------------------------------------------------------------------ tokens and REST

def _token(ttl_seconds):
    import time
    return {"secret": "jwt", "expires": int(time.time() + ttl_seconds)}


def test_access_token_returns_a_live_token():
    assert A.access_token(FakeSession([_token(3600)])) == "jwt"


def test_access_token_explains_an_expired_token_rather_than_failing_later():
    # Worth its own message: an expired token produces a 401 whose body lands on disk as a file, so
    # the useful place to notice is here.
    with raises(OutlookError, "expired"):
        A.access_token(FakeSession([_token(-120)]))


def test_access_token_explains_an_empty_msal_cache():
    with raises(OutlookError, "MSAL"):
        A.access_token(FakeSession([None]))


def test_rest_get_raises_with_the_status_and_code():
    s = FakeSession([{"status": 401, "body": {"error": {"code": "Unauthorized", "message": "no"}}}])
    with raises(OutlookError, "401"):
        A.rest_get(s, "/api/v2.0/users/x/messages", token="jwt")


def test_rest_get_does_not_ask_for_a_token_when_one_is_passed():
    s = FakeSession([{"status": 200, "body": {"value": []}}])
    A.rest_get(s, "/p", token="jwt")
    assert len(s.asked) == 1  # one fetch, no MSAL cache read


def test_list_attachments_drops_inline_images_by_default():
    # A signature block routinely contributes three or four logos; "the attachments" never means those.
    body = {"status": 200, "body": {"value": [
        {"Id": "a1", "Name": "logo.png", "Size": 3467, "ContentType": "image/png", "IsInline": True},
        {"Id": "a2", "Name": "report.pdf", "Size": 624607, "ContentType": "application/pdf",
         "IsInline": False},
    ]}}
    got = A.list_attachments(FakeSession([body]), "MID", token="jwt")
    assert [a.name for a in got] == ["report.pdf"]
    assert got[0].message_id == "MID" and got[0].attachment_id == "a2"
    assert got[0].is_file is True

    got = A.list_attachments(FakeSession([body]), "MID", token="jwt", include_inline=True)
    assert [a.name for a in got] == ["logo.png", "report.pdf"]


def test_reported_size_is_not_the_content_length():
    """Outlook's `Size` counts the MIME envelope, so it must never be asserted as the file size.

    Measured live: a PDF reported as 624607 bytes downloads as 624077, and a spreadsheet reported as
    22409319 downloads as 22408839. Comparing the two fails on every correct download.
    """
    att = A.Attachment("m", "a", "report.pdf", 624607, "application/pdf", False)
    assert att.size != 624077


def test_messages_with_attachments_sorts_server_side_and_filters_here():
    # Asking the server for $filter=HasAttachments AND $orderby=ReceivedDateTime is rejected as
    # InefficientFilter, so the sort is the server's job and the filter is ours.
    s = FakeSession([{"status": 200, "body": {"value": [
        {"Id": "1", "HasAttachments": False},
        {"Id": "2", "HasAttachments": True},
    ]}}])
    assert [m["Id"] for m in A.messages_with_attachments(s, token="jwt")] == ["2"]
    sent = s.asked[0]
    assert "$orderby=ReceivedDateTime%20desc" in sent
    assert "HasAttachments%20eq" not in sent


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
