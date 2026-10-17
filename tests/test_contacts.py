"""Contact shaping, read-back comparison and phone matching. No browser: every test works on dicts.

What can be pinned offline is that we send what we mean, refuse what would not fit, notice when what came
back is not what was sent, and recognise one number across the ways people type it.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp import contacts  # noqa: E402
from outlook_cdp.session import OutlookError  # noqa: E402


def _contact(**kw):
    base = dict(given_name="Pat", middle_name="Q", surname="Example", mobile_phone="+1 (555) 555-0100",
                home_phones=["555-555-0101"], emails=["pat@example.com"],
                home_address={"street": "1 Main St", "city": "Springfield", "state": "IL",
                              "postal_code": "60000"},
                personal_notes="Line one.\nLine two.")
    base.update(kw)
    return contacts.Contact(**base)


def _raises(fn, needle=""):
    try:
        fn()
    except OutlookError as e:
        assert needle in str(e), str(e)
    else:
        raise AssertionError("expected OutlookError")


def test_to_rest_shapes_names_phones_and_address():
    out = _contact().to_rest()
    assert out["GivenName"] == "Pat" and out["Surname"] == "Example"
    assert out["DisplayName"] == "Pat Q Example"
    assert out["MobilePhone1"] == "+1 (555) 555-0100" and out["HomePhones"] == ["555-555-0101"]
    assert out["EmailAddresses"] == [{"Name": "pat@example.com", "Address": "pat@example.com"}]
    assert out["HomeAddress"] == {"Street": "1 Main St", "City": "Springfield", "State": "IL",
                                  "PostalCode": "60000"}
    assert out["PersonalNotes"] == "Line one.\nLine two."
    assert "BusinessPhones" not in out and "Birthday" not in out


def test_an_explicit_display_name_wins():
    assert _contact(display_name="Pat (club)").to_rest()["DisplayName"] == "Pat (club)"


def test_an_empty_contact_is_refused():
    _raises(lambda: contacts.Contact().to_rest(), "empty")


def test_more_values_than_slots_are_refused_not_dropped():
    _raises(lambda: _contact(home_phones=["1", "2", "3"]).to_rest(), "home phones")
    _raises(lambda: _contact(emails=["a@x", "b@x", "c@x", "d@x"]).to_rest(), "email")


def test_unknown_address_keys_and_bad_birthdays_are_refused():
    _raises(lambda: _contact(home_address={"zip": "60000"}).to_rest(), "zip")
    _raises(lambda: _contact(birthday="1/2/1980").to_rest(), "YYYY-MM-DD")


def test_round_trip_through_the_server_shape():
    asked = _contact(birthday="1980-01-02")
    raw = asked.to_rest()
    raw["Birthday"] = "1980-01-02T11:59:00Z"
    raw["PersonalNotes"] = "Line one.\r\nLine two."           # Outlook hands notes back with CRLF
    raw["EmailAddresses"] = [{"Name": "Pat Example", "Address": "PAT@example.com"}]
    stored = contacts.Contact.from_rest(raw)
    assert contacts.differences(asked, stored) == []


def test_a_dropped_field_is_noticed():
    asked = _contact()
    raw = asked.to_rest()
    del raw["HomePhones"]
    raw["HomeAddress"].pop("PostalCode")
    probs = contacts.differences(asked, contacts.Contact.from_rest(raw))
    assert any(p.startswith("home_phones") for p in probs)
    assert any(p.startswith("home_address") for p in probs)


def test_phone_digits_ignores_formatting_and_country_code():
    want = "5555550100"
    for typed in ("+1 (555) 555-0100", "555.555.0100", "1-555-555-0100", "(555)\xa0555-0100"):
        assert contacts.phone_digits(typed) == want, typed


def test_find_by_phone_searches_every_slot():
    book = contacts.Contacts.__new__(contacts.Contacts)
    people = [contacts.Contact(display_name="A", mobile_phone="555-555-0199"),
              contacts.Contact(display_name="B", business_phones=["", "+1 555 555 0100"]),
              contacts.Contact(display_name="C", home_phones=["(555) 555-0100"])]
    assert [c.name() for c in book.find_by_phone("5555550100", people)] == ["B", "C"]
    _raises(lambda: book.find_by_phone("123", people), "too few digits")


class _FakeBook(contacts.Contacts):
    """`update` without a browser: GET returns `stored`, PATCH records its payload and applies it."""

    def __init__(self, stored):
        self.s, self.stored, self.sent, self.base, self._token = None, stored, None, "/c", "t"

    def get(self, contact_id):
        return contacts.Contact.from_rest(dict(self.stored))

    def _patch(self, session, method, path, payload, token=None):
        self.sent = payload
        self.stored.update({k: v for k, v in payload.items()})
        return {}


def _with_fake_patch(book, fn):
    real = contacts.rest_send
    contacts.rest_send = book._patch
    try:
        return fn()
    finally:
        contacts.rest_send = real


def test_update_sends_null_to_clear_and_checks_it_cleared():
    book = _FakeBook(_contact(nickname="Pat").to_rest())
    got = _with_fake_patch(book, lambda: book.update("ID", nickname="", middle_name=""))
    assert book.sent["NickName"] is None and book.sent["MiddleName"] is None
    assert got.nickname == "" and got.middle_name == ""
    assert book.sent["DisplayName"] == "Pat Q Example"   # a stored display name is kept unless passed too


def test_update_refuses_a_field_the_server_kept():
    book = _FakeBook(_contact(nickname="Pat").to_rest())
    book._patch = lambda *a, **k: {}                     # the server ignores the PATCH entirely
    _raises(lambda: _with_fake_patch(book, lambda: book.update("ID", nickname="")), "nickname")


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
