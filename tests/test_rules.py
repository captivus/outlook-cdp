"""Inbox-rule shaping, read-back comparison and local matching. No browser: every test works on dicts.

What a live server does with a rule - drop an unknown condition, demand a Sequence - is documented in
`rules.py`; what can be pinned offline is that we send what we mean, notice when what came back is not
it, and evaluate conditions the way Outlook does.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outlook_cdp import rules  # noqa: E402
from outlook_cdp.session import OutlookError  # noqa: E402


def _rule(**kw):
    base = dict(name="Archive notices", from_addresses=["notify@example.com"],
                subject_contains=["Usage Warning"], move_to="archive")
    base.update(kw)
    return rules.Rule(**base)


def test_to_rest_shapes_addresses_and_actions():
    out = _rule(sequence=3).to_rest("FOLDER-ID")
    assert out["Conditions"] == {"FromAddresses": [{"EmailAddress": {"Address": "notify@example.com"}}],
                                 "SubjectContains": ["Usage Warning"]}
    assert out["Actions"] == {"StopProcessingRules": True, "MoveToFolder": "FOLDER-ID"}
    assert out["Sequence"] == 3 and out["IsEnabled"] is True


def test_a_rule_with_no_conditions_is_refused():
    """It would match every message - archiving a whole inbox is the failure this prevents."""
    try:
        rules.Rule("everything", move_to="archive").to_rest("FOLDER-ID")
    except OutlookError as e:
        assert "every message" in str(e)
    else:
        raise AssertionError("a condition-less rule was accepted")


def test_an_unresolved_folder_is_refused():
    try:
        _rule().to_rest(None)
    except OutlookError:
        pass
    else:
        raise AssertionError("move_to without a folder id was accepted")


def test_round_trip_through_the_server_shape():
    raw = _rule(sequence=1).to_rest("FOLDER-ID")
    raw["Id"] = "RULE-ID"
    back = rules.Rule.from_rest(raw)
    assert back.from_addresses == ["notify@example.com"]
    assert back.subject_contains == ["Usage Warning"]
    assert back.move_to == "FOLDER-ID" and back.id == "RULE-ID" and back.sequence == 1
    assert rules.differences(_rule(), back, "FOLDER-ID") == []


def test_a_dropped_condition_is_a_difference():
    """The server accepting a POST is not the server storing the rule that was sent."""
    stored = rules.Rule.from_rest({"DisplayName": "Archive notices", "IsEnabled": True,
                                   "Conditions": {"FromAddresses": [{"EmailAddress": {
                                       "Address": "notify@example.com"}}]},
                                   "Actions": {"MoveToFolder": "FOLDER-ID"}})
    probs = rules.differences(_rule(), stored, "FOLDER-ID")
    assert len(probs) == 1 and probs[0].startswith("subject_contains")


def test_empty_and_null_conditions_from_the_server_are_not_unknown():
    stored = rules.Rule.from_rest({"Conditions": {"SubjectContains": ["x"], "BodyContains": [],
                                                  "Importance": None, "IsAutomaticReply": False},
                                   "Actions": {"Delete": False, "CopyToFolder": None}})
    assert stored.unknown_conditions == [] and stored.unknown_actions == []


def test_a_condition_this_module_does_not_know_is_reported():
    stored = rules.Rule.from_rest({"Conditions": {"SubjectContains": ["x"], "Importance": "High"},
                                   "Actions": {"ForwardTo": [{"EmailAddress": {"Address": "a@b.c"}}]}})
    assert stored.unknown_conditions == ["Importance"]
    assert stored.unknown_actions == ["ForwardTo"]


def test_matching_ands_kinds_and_ors_values():
    r = _rule(subject_contains=["Usage Warning", "Quota Notice"])
    assert r.matches("Example Org API Usage Warning", "notify@example.com")
    assert r.matches("Quota notice for today", "NOTIFY@example.com")          # case-insensitive
    assert not r.matches("Example Org API Access Suspended", "notify@example.com")
    assert not r.matches("Example Org API Usage Warning", "someone@example.com")


def test_from_address_is_exact_not_substring():
    r = _rule(subject_contains=[])
    assert not r.matches("anything", "notify@example.com.attacker.test")


def test_body_conditions_see_the_body():
    r = _rule(from_addresses=[], subject_contains=[], body_contains=["requests per second"])
    assert r.matches("s", "x@y.z", body="Your hourly requests per second exceeded")
    assert not r.matches("requests per second", "x@y.z", body="")


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn(); n += 1
            print(f"  ok  {name}")
    print(f"{n} passed")
