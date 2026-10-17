"""Duplicate detection and the no-information-lost check. No browser: every test works on dicts.

What can be pinned offline is that a contact is only ever planned for deletion when another one holds
everything it does, in both views of it; that "holds" survives the ways Outlook and people format the same
thing but not a change of meaning or of slot; and that anything unfamiliar keeps a contact rather than
losing it.
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp import dedupe  # noqa: E402
from outlook_cdp.session import OutlookError  # noqa: E402

_n = [0]


def _entry(photo=None, folder=None, ews=None, **record):
    _n[0] += 1
    base = {"Id": f"id{_n[0]}", "ChangeKey": f"ck{_n[0]}", "GivenName": "Pat", "Surname": "Example",
            "DisplayName": "Pat Example", "EmailAddresses": [], "HomePhones": [], "BusinessPhones": [],
            "Categories": [], "LastModifiedDateTime": f"2020-01-01T00:00:{_n[0] % 60:02d}Z"}
    base.update(record)
    item = {"ItemId": {"Id": f"ews{_n[0]}", "ChangeKey": base["ChangeKey"]}, "DisplayName": base["DisplayName"],
            "GivenName": base["GivenName"], "Surname": base["Surname"], "Sensitivity": "Normal"}
    item.update(ews or {})
    return {"record": base, "photo": photo, "folder": folder, "ews": item}


def _backup(*entries, groups=()):
    return {"taken_at": "T", "mailbox": "me@example.com", "folders": [{"id": None, "name": "Contacts"}],
            "contacts": list(entries), "groups": list(groups)}


def _rest(dup, keep):
    return dedupe.rest_gaps(dup, keep)


def _ews(dup, keep):
    return dedupe.ews_gaps(dup, keep)


# ------------------------------------------------------------------ what is held

def test_an_identical_copy_loses_nothing():
    a = _entry(EmailAddresses=[{"Name": "Pat", "Address": "pat@example.com"}])
    b = _entry(EmailAddresses=[{"Name": "Pat", "Address": "PAT@example.com"}])
    assert dedupe.coverage_gaps(a, b) == [] and dedupe.coverage_gaps(b, a) == []


def test_a_subset_is_held_but_the_superset_is_not():
    small = _entry(MobilePhone1="555-555-0100")
    big = _entry(MobilePhone1="+1 (555) 555-0100", JobTitle="Engineer")
    assert _rest(small, big) == []
    assert any("JobTitle" in g for g in _rest(big, small))


def test_case_counts_in_names_but_not_in_addresses():
    assert _rest(_entry(GivenName="McDonald"), _entry(GivenName="Mcdonald")) != []
    a = _entry(EmailAddresses=[{"Name": "", "Address": "Pat@Example.com"}])
    b = _entry(EmailAddresses=[{"Name": "", "Address": "pat@example.com"}])
    assert _rest(a, b) == []


def test_a_number_must_be_in_the_same_kind_of_slot():
    fax = _entry(SingleValueExtendedProperties=[{"PropertyId": "String 0x3A24", "Value": "555-555-0100"}])
    assert _rest(fax, _entry(MobilePhone1="5555550100")) != []
    assert _rest(fax, _entry(SingleValueExtendedProperties=[
        {"PropertyId": "String 0x3a24", "Value": "(555) 555-0100"}])) == []
    assert _rest(_entry(HomePhones=["555.555.0100"]), _entry(BusinessPhones=["5555550100"])) != []


def test_an_extension_is_part_of_the_number():
    assert _rest(_entry(MobilePhone1="555-555-0100 x12"), _entry(MobilePhone1="555555010012")) != []
    assert _rest(_entry(MobilePhone1="555-555-0100 ext. 12"), _entry(MobilePhone1="5555550100 x12")) == []


def test_an_extension_written_without_a_space_is_still_an_extension():
    assert _rest(_entry(MobilePhone1="555-0000x12"), _entry(MobilePhone1="555000012")) != []


def test_a_default_is_only_skipped_when_the_kept_contact_has_it_too():
    assert _ews(_entry(ews={"Sensitivity": "Normal"}), _entry(ews={"Sensitivity": "Private"})) != []
    assert _ews(_entry(ews={"FileAsMapping": "LastCommaFirst"}), _entry(ews={"FileAsMapping": "None"})) != []


def test_the_note_format_counts():
    dup = _entry(ews={"Body": {"BodyType": "HTML", "Value": "hello"}})
    assert _ews(dup, _entry(ews={"Body": {"BodyType": "Text", "Value": "hello"}})) != []
    assert _ews(dup, _entry(ews={"Body": {"BodyType": "HTML", "Value": "hello"}})) == []


def test_attachments_flagged_but_not_listed_are_not_held():
    assert _ews(_entry(ews={"HasAttachments": True}), _entry(ews={"HasAttachments": True})) != []


def test_third_pass_counterexamples_are_not_held():
    assert _ews(_entry(ews={"Culture": "fr-FR"}), _entry(ews={"Culture": "en-US"})) != []
    pid = "String {00062004-0000-0000-C000-000000000046} Id 0x804F"
    assert _rest(_entry(SingleValueExtendedProperties=[{"PropertyId": pid, "Value": "Route/"}]),
                 _entry(SingleValueExtendedProperties=[{"PropertyId": pid, "Value": "Route"}])) != []
    dup = _entry(EmailAddresses=[{"Address": "a@example.com", "Name": "A", "RoutingType": "EX"}])
    keep = _entry(EmailAddresses=[{"Address": "a@example.com", "Name": "A", "RoutingType": "SMTP"}])
    assert _rest(dup, keep) != []


def test_one_kept_email_entry_must_hold_the_whole_entry():
    dup = _entry(EmailAddresses=[{"Address": "a@example.com", "Name": "A", "RoutingType": "EX", "MailboxType": "One"}])
    keep = _entry(EmailAddresses=[
        {"Address": "a@example.com", "Name": "A", "RoutingType": "EX", "MailboxType": "Two"},
        {"Address": "a@example.com", "Name": "B", "RoutingType": "SMTP", "MailboxType": "One"}])
    assert _rest(dup, keep) != []


def test_dial_pauses_and_doubled_slashes_count():
    assert _rest(_entry(MobilePhone1="555-0100,123"), _entry(MobilePhone1="5550100123")) != []
    assert _rest(_entry(MobilePhone1="555-0100,123"), _entry(MobilePhone1="555 0100,123")) == []
    assert _rest(_entry(BusinessHomePage="https://example.com/a//"), _entry(BusinessHomePage="https://example.com/a/")) != []


def test_empty_extended_values_are_not_information():
    dup = _entry(SingleValueExtendedProperties=[{"PropertyId": "String 0x3A1F", "Value": ""},
                                                {"PropertyId": "Short 0x3A4D", "Value": "0"}])
    assert _rest(dup, _entry()) == []


def test_notes_are_held_only_as_whole_words():
    dup = _entry(PersonalNotes="Met at\r\nthe club.")
    assert _rest(dup, _entry(PersonalNotes="Old note.\nMet at the club. More.")) == []
    assert _rest(dup, _entry(PersonalNotes="Something else")) != []
    assert _rest(_entry(PersonalNotes="Ann"), _entry(PersonalNotes="Joanne")) != []


def test_an_address_is_held_in_the_same_slot_and_zip_only_by_zip_plus_four():
    dup = _entry(HomeAddress={"Street": "1 Main St", "PostalCode": "60000"})
    assert _rest(dup, _entry(HomeAddress={"Street": "1 Main St", "City": "X", "PostalCode": "60000-1234"})) == []
    assert _rest(dup, _entry(OtherAddress={"Street": "1 Main St", "PostalCode": "60000"})) != []
    assert _rest(_entry(HomeAddress={"PostalCode": "6000"}), _entry(HomeAddress={"PostalCode": "60001"})) != []


def test_web_pages_match_exactly_but_for_a_trailing_slash():
    assert _rest(_entry(BusinessHomePage="https://example.com/"), _entry(BusinessHomePage="https://example.com")) == []
    assert _rest(_entry(BusinessHomePage="https://www.example.com"), _entry(BusinessHomePage="http://example.com")) != []


def test_photos_must_match_byte_for_byte():
    pic = {"type": "image/jpeg", "b64": "", "sha256": "aa"}
    assert _rest(_entry(photo=pic), _entry(photo=dict(pic))) == []
    assert _rest(_entry(photo=pic), _entry()) == ["photo (kept has none)"]
    assert _rest(_entry(photo=pic), _entry(photo={**pic, "sha256": "bb"})) != []


def test_an_unknown_property_is_never_held():
    assert any("SomethingNew" in g for g in _rest(_entry(SomethingNew="x"), _entry(SomethingNew="x")))


def test_an_email_keeps_its_own_display_name():
    dup = _entry(EmailAddresses=[{"Address": "a@example.com", "Name": "CEO"},
                                 {"Address": "b@example.com", "Name": "Bob"}])
    keep = _entry(EmailAddresses=[{"Address": "a@example.com", "Name": "Bob"},
                                  {"Address": "b@example.com", "Name": "CEO"}])
    assert len(_rest(dup, keep)) == 2


def test_email_slot_details_are_compared():
    dup = _entry(EmailAddresses=[{"Name": "", "Address": "a@example.com"}],
                 Email1Address={"Address": "a@example.com", "RoutingType": "EX"})
    keep = _entry(EmailAddresses=[{"Name": "", "Address": "a@example.com"}],
                  Email1Address={"Address": "a@example.com", "RoutingType": "SMTP"})
    assert any("RoutingType" in g for g in _rest(dup, keep))
    keep["record"]["Email1Address"]["RoutingType"] = "EX"
    assert _rest(dup, keep) == []


def test_ignored_categories_are_not_information():
    dup = _entry(Categories=["Old Import Marker", "Business"])
    keep = _entry(Categories=["business"])
    assert _rest(dup, keep) != []
    assert dedupe.rest_gaps(dup, keep, ["old import marker"]) == []


# ------------------------------------------------------------------ the second view

def test_without_a_second_view_nothing_is_held():
    dup, keep = _entry(), _entry()
    dup["ews"] = None
    assert dedupe.coverage_gaps(dup, keep) != []


def test_the_second_view_keys_numbers_by_kind():
    dup = _entry(ews={"PhoneNumbers": [{"Key": "BusinessFax", "PhoneNumber": "555-555-0100"},
                                       {"Key": "Pager", "PhoneNumber": ""}]})
    keep = _entry(ews={"PhoneNumbers": [{"Key": "MobilePhone", "PhoneNumber": "5555550100"}]})
    assert [g for g in _ews(dup, keep) if "BusinessFax" in g]
    keep["ews"]["PhoneNumbers"].append({"Key": "BusinessFax", "PhoneNumber": "(555) 555-0100"})
    assert _ews(dup, keep) == []


def test_the_second_view_compares_what_rest_cannot_see():
    flagged = _entry(ews={"Flag": {"FlagStatus": "Flagged"}, "Sensitivity": "Private",
                          "DisplayNamePrefix": "Dr."})
    gaps = _ews(flagged, _entry(ews={"Flag": {"FlagStatus": "NotFlagged"}}))
    assert {g.split(":")[0] for g in gaps} == {"Flag", "Sensitivity", "DisplayNamePrefix"}


def test_the_second_view_holds_nothing_it_does_not_know():
    assert _ews(_entry(ews={"BrandNewField": "x"}), _entry()) != []


def test_attachments_other_than_the_photo_are_never_held():
    photo = {"IsContactPhoto": True, "Name": "ContactPicture.jpg"}
    assert _ews(_entry(ews={"Attachments": [photo]}), _entry()) == []
    att = _entry(ews={"Attachments": [photo, {"IsContactPhoto": False, "Name": "card.pdf"}]})
    assert _ews(att, _entry(ews={"Attachments": [photo, {"IsContactPhoto": False, "Name": "card.pdf"}]})) != []


# ------------------------------------------------------------------ planning

def test_plan_keeps_the_richest_and_deletes_only_what_it_holds():
    rich = _entry(MobilePhone1="555-555-0100", JobTitle="Engineer", ews={"JobTitle": "Engineer"})
    poor = _entry(MobilePhone1="5555550100")
    other = _entry(MobilePhone1="5555550100", GivenName="Sam", DisplayName="Sam Example")
    p = dedupe.plan(_backup(poor, rich, other))
    assert [(d["delete"]["id"], d["keep"]["id"]) for d in p["delete"]] == [(poor["record"]["Id"], rich["record"]["Id"])]
    assert len(p["kept_with_differences"]) == 1           # Sam shares the number but is someone else


def test_identical_copies_leave_exactly_one():
    a, b, c = (_entry(EmailAddresses=[{"Name": "", "Address": "pat@example.com"}]) for _ in range(3))
    p = dedupe.plan(_backup(a, b, c))
    assert len(p["delete"]) == 2
    kept = {d["keep"]["id"] for d in p["delete"]}
    assert len(kept) == 1 and kept.isdisjoint({d["delete"]["id"] for d in p["delete"]})


def test_name_only_matches_wait_for_approval():
    a, b = _entry(), _entry()
    p = dedupe.plan(_backup(a, b))
    assert p["delete"] == [] and len(p["name_only"]) == 1
    assert dedupe.approved(p) == []
    p["name_only"][0]["deletable"][0]["approved"] = True
    assert len(dedupe.approved(p)) == 1


def test_contact_group_members_are_never_planned_for_deletion():
    a = _entry(EmailAddresses=[{"Name": "", "Address": "pat@example.com"}])
    b = _entry(EmailAddresses=[{"Name": "", "Address": "pat@example.com"}])
    group = {"folder": None, "item": {"Members": [{"Mailbox": {"EmailAddress": "PAT@example.com"}}]}}
    p = dedupe.plan(_backup(a, b, groups=[group]))
    assert p["delete"] == [] and len(p["in_a_contact_group"]) == 1


def test_of_identical_copies_the_preferred_folders_is_kept():
    old = _entry(folder="F", MobilePhone1="5555550100", LastModifiedDateTime="2025-01-01T00:00:00Z")
    new = _entry(folder=None, MobilePhone1="5555550100", LastModifiedDateTime="2020-01-01T00:00:00Z")
    backup = _backup(old, new)
    backup["folders"].append({"id": "F", "name": "Old"})
    [d] = dedupe.plan(backup)["delete"]
    assert d["delete"]["id"] == old["record"]["Id"] and d["keep"]["id"] == new["record"]["Id"]
    [d] = dedupe.plan(backup, prefer_folder="F")["delete"]
    assert d["keep"]["id"] == old["record"]["Id"]


def test_a_fuller_copy_elsewhere_still_wins_over_the_preferred_folder():
    poor = _entry(folder=None, MobilePhone1="5555550100")
    rich = _entry(folder="F", MobilePhone1="5555550100", JobTitle="Engineer", ews={"JobTitle": "Engineer"})
    backup = _backup(poor, rich)
    backup["folders"].append({"id": "F", "name": "Old"})
    [d] = dedupe.plan(backup)["delete"]
    assert d["keep"]["id"] == rich["record"]["Id"]


def test_same_folder_only():
    a = _entry(folder=None, MobilePhone1="5555550100")
    b = _entry(folder="F", MobilePhone1="5555550100")
    backup = _backup(a, b)
    backup["folders"].append({"id": "F", "name": "Old"})
    assert len(dedupe.plan(backup)["delete"]) == 1
    assert dedupe.plan(backup, across_folders=False)["delete"] == []


# ------------------------------------------------------------------ applying

class _FakeBook:
    """The live book `apply` sees: the backup's contacts, unchanged unless a test changes them."""

    def __init__(self, backup):
        self.mailbox = backup["mailbox"]
        self.live = {e["record"]["Id"]: e for e in backup["contacts"]}
        self.deleted, self.live_groups = [], [g["item"] for g in backup["groups"]]

    def folders(self):
        return [{"id": None, "name": "Contacts"}]

    def groups(self, folder_id=None):
        return self.live_groups

    def get_raw(self, cid):
        return self.live[cid]["record"] if cid in self.live else None

    def get_items(self, ews_ids):
        return [e["ews"] for e in self.live.values() if e["ews"]["ItemId"]["Id"] in ews_ids]

    def photos(self, ids):
        return {i: self.live[i]["photo"] for i in ids}

    def delete(self, cid):
        self.deleted.append(cid)
        del self.live[cid]


def _apply(backup, plan, book=None):
    book = book or _FakeBook(backup)
    with tempfile.NamedTemporaryFile("r", suffix=".jsonl") as log:
        tally = dedupe.apply(book, plan, backup, log.name)
    return tally, book


def test_apply_deletes_what_the_plan_and_the_live_book_agree_on():
    a, b = (_entry(MobilePhone1="5555550100") for _ in range(2))
    backup = _backup(a, b)
    tally, book = _apply(backup, dedupe.plan(backup))
    assert tally["deleted"] == 1 and len(book.deleted) == 1


def test_apply_refuses_a_deletion_the_backup_does_not_support():
    a = _entry(MobilePhone1="5555550100")
    b = _entry()                                   # no shared number: not a match
    backup = _backup(a, b)
    plan = dedupe.plan(backup)
    plan["delete"].append({"delete": {"id": b["record"]["Id"], "name": "x", "folder": "Contacts"},
                           "keep": {"id": a["record"]["Id"], "name": "y", "folder": "Contacts"}})
    book = _FakeBook(backup)
    try:
        _apply(backup, plan, book)
    except OutlookError as e:
        assert "made again" in str(e)
    else:
        raise AssertionError("expected OutlookError")
    assert book.deleted == []


def test_apply_refuses_a_plan_whose_settings_were_edited():
    a = _entry(MobilePhone1="5555550100", Categories=["Marker"], ews={"Categories": ["Marker"]})
    b = _entry(MobilePhone1="5555550100", JobTitle="Engineer", ews={"JobTitle": "Engineer"})
    backup = _backup(a, b)
    plan = dedupe.plan(backup, ignore_categories=["Marker"])
    assert len(plan["delete"]) == 1
    plan["ignore_categories"] = []                 # edited after the fact
    try:
        _apply(backup, plan)
    except OutlookError as e:
        assert "own settings" in str(e)
    else:
        raise AssertionError("expected OutlookError")


def test_apply_skips_a_contact_that_changed_since_the_backup():
    a, b = (_entry(MobilePhone1="5555550100") for _ in range(2))
    backup = _backup(a, b)
    plan = dedupe.plan(backup)
    book = _FakeBook(backup)
    victim = plan["delete"][0]["delete"]["id"]
    book.live[victim] = {**book.live[victim], "record": {**book.live[victim]["record"], "ChangeKey": "new"}}
    tally, book = _apply(backup, plan, book)
    assert book.deleted == [] and tally["skipped"] == 1


def test_apply_protects_a_contact_added_to_a_group_after_the_backup():
    a, b = (_entry(EmailAddresses=[{"Name": "", "Address": "pat@example.com"}]) for _ in range(2))
    backup = _backup(a, b)
    plan = dedupe.plan(backup)
    book = _FakeBook(backup)
    book.live_groups = [{"Members": [{"Mailbox": {"EmailAddress": "pat@example.com"}}]}]
    tally, book = _apply(backup, plan, book)
    assert book.deleted == []


def test_apply_refuses_another_mailbox():
    a, b = (_entry(MobilePhone1="5555550100") for _ in range(2))
    backup = _backup(a, b)
    book = _FakeBook(backup)
    book.mailbox = "someone@example.com"
    try:
        _apply(backup, dedupe.plan(backup), book)
    except OutlookError as e:
        assert "someone@example.com" in str(e)
    else:
        raise AssertionError("expected OutlookError")


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
