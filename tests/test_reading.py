"""Read-path parsing and normalisation. No browser: every test here works on strings.

The live behaviours these back up - which selector, which id, which wake call - cannot be tested
without a mailbox, so they are documented in `reading.py` and in the README instead. What *can* be
pinned down offline is every place a string was misread, and each test below is one of those.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp import reading  # noqa: E402


# --------------------------------------------------------------------- row aria-labels

def test_row_flags_are_split_off_the_front():
    flags, rest = reading.parse_row_flags("Unread Collapsed Has attachments A Sender A subject Tue 7:33 PM")
    assert flags == {"Unread": True, "Collapsed": True, "Has attachments": True}
    assert rest.startswith("A Sender")


def test_collapsed_alone_does_not_make_a_row_a_conversation():
    """The mistake this prevents: `Collapsed` describes the row widget, not the thread.

    Trusting it reported nearly every row in a real folder as a multi-message thread.
    """
    flags, _ = reading.parse_row_flags("Collapsed A Sender A subject 9/25/2026")
    assert flags == {"Collapsed": True}
    assert "Has attachments" not in flags


def test_a_row_with_no_flags_keeps_its_whole_label():
    flags, rest = reading.parse_row_flags("A Sender A subject 9/24/2026")
    assert flags == {}
    assert rest == "A Sender A subject 9/24/2026"


def test_count_badge():
    assert reading.parse_count_badge("(2)") == 2
    assert reading.parse_count_badge("(13)") == 13
    assert reading.parse_count_badge("") is None
    assert reading.parse_count_badge("2") is None


def test_container_count_accepts_outlooks_own_grammar():
    """A single message is labelled "1 messages"; a reader that demanded "1 message" saw nothing."""
    assert reading.parse_container_count("1 messages") == 1
    assert reading.parse_container_count("4 messages") == 4
    assert reading.parse_container_count("1 message") == 1
    assert reading.parse_container_count("Email message") is None


# --------------------------------------------------------------------- senders

def test_sender_is_split_without_a_space_before_the_bracket():
    """The pane writes `Name<addr>` with no space, which a naive `split(" <")` misses entirely."""
    assert reading.split_sender("A Name<a@b.gov>") == ("A Name", "a@b.gov")
    assert reading.split_sender("A Name <a@b.gov>") == ("A Name", "a@b.gov")


def test_a_bare_address_is_both_name_and_address():
    assert reading.split_sender("a@b.gov") == ("a@b.gov", "a@b.gov")


def test_a_bare_display_name_has_no_address():
    assert reading.split_sender("Records Office") == ("Records Office", None)


def test_addresses_are_lowercased():
    assert reading.split_sender("A Name<A.Name@B.GOV>")[1] == "a.name@b.gov"


def test_on_behalf_of_returns_the_human_as_the_sender():
    """Delegated mail reads `X<a@b> on behalf of Y<c@d>`; the first identity actually sent it."""
    text = "A Person<person@body.gov> on behalf of Records Office<records@body.gov>"
    assert reading.split_sender(text) == ("A Person", "person@body.gov")
    assert reading.on_behalf_of(text) == ("Records Office", "records@body.gov")


def test_on_behalf_of_is_none_for_ordinary_mail():
    assert reading.on_behalf_of("A Person<person@body.gov>") is None


# --------------------------------------------------------------------- recipients

def test_recipients_drop_the_field_label():
    assert reading.parse_recipients("To: a@b.gov") == ["a@b.gov"]
    assert reading.parse_recipients("Cc:A Name <a@b.gov>") == ["A Name <a@b.gov>"]


def test_recipients_split_on_separators_outside_angle_brackets():
    """A display name can contain a comma - "Surname, Given" is how Outlook renders many senders - so
    splitting on every comma tore one recipient into two."""
    got = reading.parse_recipients("To: Surname, Given <a@b.gov>, Other Name <c@d.gov>")
    assert got == ["Surname", "Given <a@b.gov>", "Other Name <c@d.gov>"]
    assert reading.addresses_in(got) == ["a@b.gov", "c@d.gov"]


def test_recipients_of_an_empty_field():
    assert reading.parse_recipients("To:") == []
    assert reading.parse_recipients("") == []


def test_addresses_in_ignores_entries_with_no_address():
    assert reading.addresses_in(["A Name", "Other <o@b.gov>"]) == ["o@b.gov"]


# --------------------------------------------------------------------- glyphs

def test_private_use_icon_glyphs_are_stripped():
    """Outlook's icons are private-use font glyphs, and they land in innerText where they silently
    break every string comparison against a date or a sender."""
    assert reading.strip_glyphs("Tue 9/29/2026 7:33 PM") == "Tue 9/29/2026 7:33 PM"
    assert reading.strip_glyphs("") == ""


# --------------------------------------------------------------------- html bodies

def test_html_to_text_keeps_paragraph_breaks():
    got = reading.html_to_text("<html><body><p>First line.</p><p>Second line.</p></body></html>")
    assert got == "First line.\n\nSecond line."


def test_html_to_text_unescapes_and_folds_nbsp():
    """Non-breaking spaces are invisible, they are everywhere in mail pasted out of Word, and they
    make an otherwise identical body fail to compare equal."""
    assert reading.html_to_text("<div>a&nbsp;&amp;&nbsp;b</div>") == "a & b"


def test_html_to_text_drops_style_and_script():
    got = reading.html_to_text("<style>.x{color:red}</style><div>Body text</div>"
                               "<script>var x=1</script>")
    assert got == "Body text"


def test_html_to_text_separates_table_cells_with_space():
    """Signature blocks are laid out in table cells; joining them with nothing runs words together."""
    got = reading.html_to_text("<table><tr><td>Name</td><td>Title</td></tr></table>")
    assert "Name Title" in got


def test_html_to_text_collapses_runs_of_blank_lines():
    got = reading.html_to_text("<div>a</div><div><br></div><div><br></div><div><br></div><div>b</div>")
    assert got == "a\n\nb"


def test_html_to_text_keeps_image_alt_text():
    assert reading.html_to_text('<div><img alt="Logo"> after</div>') == "[Logo] after"


def test_html_to_text_passes_plain_text_through():
    assert reading.html_to_text("already plain") == "already plain"
    assert reading.html_to_text("") == ""
    assert reading.html_to_text(None) == ""


# --------------------------------------------------------------------- returned shapes

def test_conversation_newest_and_oldest_follow_list_order():
    conv = reading.Conversation(messages=[reading.Message(index=0, sender="First"),
                                          reading.Message(index=1, sender="Second")])
    assert conv.oldest.sender == "First"
    assert conv.newest.sender == "Second"


def test_conversation_is_incomplete_when_the_pane_rendered_fewer_than_it_claimed():
    """The guard that stops `newest` being trusted when part of the thread never arrived."""
    conv = reading.Conversation(expected_count=4, messages=[reading.Message()])
    assert not conv.complete
    assert reading.Conversation(expected_count=1, messages=[reading.Message()]).complete
    assert reading.Conversation(messages=[reading.Message()]).complete


def test_conversation_from_address_matches_case_insensitively():
    conv = reading.Conversation(messages=[reading.Message(sender_address="a@b.gov"),
                                          reading.Message(sender_address="c@d.gov")])
    assert len(conv.from_address("A@B.GOV")) == 1


def test_message_recipient_addresses_are_derived():
    m = reading.Message(to=["A Name <a@b.gov>"], cc=["c@d.gov"])
    assert m.to_addresses == ["a@b.gov"]
    assert m.cc_addresses == ["c@d.gov"]


def test_row_id_is_the_conversation_id():
    assert reading.Row(conversation_id="abc=").id == "abc="
    assert reading.Row().id is None


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
