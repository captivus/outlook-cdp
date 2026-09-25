"""Guard logic. No browser: every guard reads only a Draft snapshot."""
import csv
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp import guards  # noqa: E402


def _blocklist(rows):
    fh = tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, newline="")
    w = csv.DictWriter(fh, fieldnames=["address", "reason", "first_bounced"])
    w.writeheader()
    w.writerows(rows)
    fh.close()
    return fh.name


def test_send_lock_blocks_when_the_file_exists():
    fh = tempfile.NamedTemporaryFile("w", delete=False)
    fh.write("locked because a human said so")
    fh.close()
    problems = guards.SendLock(fh.name)(guards.Draft())
    assert problems and "locked because a human said so" in problems[0]


def test_send_lock_is_silent_when_absent():
    assert guards.SendLock("/nonexistent/SEND_LOCK")(guards.Draft()) == []


def test_blocklist_fails_closed_when_the_file_is_missing():
    problems = guards.Blocklist("/nonexistent/bounced.csv")(guards.Draft())
    assert problems and "missing" in problems[0]


def test_blocklist_catches_a_chip_address():
    path = _blocklist([{"address": "bad@example.gov", "reason": "bounced", "first_bounced": "2026-09-23"}])
    d = guards.Draft(chip_addresses=["bad@example.gov"])
    problems = guards.Blocklist(path)(d)
    assert problems and "bounced" in problems[0]


def test_blocklist_catches_an_address_hidden_behind_a_display_name():
    """The case the whole chip reader exists for: the well shows a name, the chip holds the address."""
    path = _blocklist([{"address": "bad@example.gov", "reason": "bounced", "first_bounced": ""}])
    d = guards.Draft(chip_addresses=[], well_text="County FOIA Office <bad@example.gov>")
    assert guards.Blocklist(path)(d)


def test_from_address_must_match_exactly():
    assert guards.FromAddress("From: a@b.com")(guards.Draft(from_label="From: c@d.com"))
    assert guards.FromAddress("From: a@b.com")(guards.Draft(from_label="From: a@b.com")) == []


def test_recipients_readable_fails_on_an_unreadable_chip():
    d = guards.Draft(chip_problems=["To chip 'Someone' carries no readable address"],
                     chip_addresses=["a@b.com"])
    assert guards.RecipientsReadable()(d)


def test_recipients_readable_refuses_to_send_blind():
    assert guards.RecipientsReadable()(guards.Draft())


def test_intended_recipient_must_be_visible():
    d = guards.Draft(intended_to="wanted@example.gov", well_text="someone.else@example.gov")
    assert guards.IntendedRecipientsPresent()(d)


def test_intended_recipient_alias_is_accepted():
    d = guards.Draft(intended_to="office@example.com", well_text="example office")
    assert guards.IntendedRecipientsPresent({"office@example.com": "example office"})(d) == []


def test_body_match_rejects_a_truncated_paste():
    d = guards.Draft(intended_body="one two three four", body="one two")
    problems = guards.BodyMatchesIntended()(d)
    assert problems and "first difference at" in problems[0]


def test_body_match_ignores_rewrapping():
    d = guards.Draft(intended_body="one two\nthree", body="one two three")
    assert guards.BodyMatchesIntended()(d) == []


def test_body_must_not_contain_catches_a_surviving_signature():
    d = guards.Draft(body="Yours,\nA Person\npersonal@example.com")
    assert guards.BodyMustNotContain(["personal@"])(d)


def test_subject_guard_skips_replies_but_wants_one():
    assert guards.SubjectMatchesIntended()(guards.Draft(is_reply=True, subject="RE: x")) == []
    assert guards.SubjectMatchesIntended()(guards.Draft(is_reply=True, subject=" "))


def test_run_prefixes_each_problem_with_its_guard():
    out = guards.run([guards.FromAddress("From: a@b.com")], guards.Draft(from_label=None))
    assert out and out[0].startswith("[from address]")




# ---------------------------------------------------------------- shape-independence
# Every test below pins a behaviour found by running against a real mailbox, where the same account in
# two configurations differed. A guard that works in one shape and fails in the other is the bug.

def test_sender_identity_accepts_a_from_control():
    d = guards.Draft(from_label="From: office@example.com")
    assert guards.SenderIdentity("office@example.com")(d) == []


def test_sender_identity_accepts_a_mailbox_url_when_there_is_no_from_control():
    """A single-mailbox session has no From chooser at all; demanding one refuses every send."""
    d = guards.Draft(from_label=None, mailbox_identity="office@example.com")
    assert guards.SenderIdentity("office@example.com")(d) == []


def test_sender_identity_rejects_the_wrong_mailbox():
    d = guards.Draft(from_label=None, mailbox_identity="someone@else.com")
    assert guards.SenderIdentity("office@example.com")(d)


def test_sender_identity_fails_closed_with_neither_source():
    assert guards.SenderIdentity("office@example.com")(guards.Draft())


def test_allowlist_matches_a_directory_resolved_recipient_by_display_name():
    """Outlook replaces an internal address with a contact; the address is then nowhere in the page."""
    d = guards.Draft(chip_addresses=[], chip_display_names=["Example Office"])
    g = guards.RecipientAllowlist(["office@example.com"], {"office@example.com": "Example Office"})
    assert g(d) == []


def test_allowlist_refuses_an_unknown_display_name():
    d = guards.Draft(chip_addresses=[], chip_display_names=["Someone Else"])
    g = guards.RecipientAllowlist(["office@example.com"], {"office@example.com": "Example Office"})
    assert g(d)


def test_allowlist_refuses_when_nothing_can_be_identified():
    assert guards.RecipientAllowlist(["office@example.com"])(guards.Draft())


def test_recipients_readable_accepts_a_name_without_an_address():
    d = guards.Draft(chip_addresses=[], chip_display_names=["Example Office"])
    assert guards.RecipientsReadable()(d) == []


def test_subject_guard_is_silent_when_a_reply_has_no_subject_field():
    """An inline reply has no editable subject: absent, not empty. Requiring one refuses every reply."""
    d = guards.Draft(is_reply=True, subject="", subject_editable=False)
    assert guards.SubjectMatchesIntended()(d) == []


def test_subject_guard_refuses_a_new_message_with_no_subject_field():
    d = guards.Draft(is_reply=False, subject="", subject_editable=False, intended_subject="x")
    assert guards.SubjectMatchesIntended()(d)


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
