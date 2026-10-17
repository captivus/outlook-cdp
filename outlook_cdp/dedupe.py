"""Duplicate contacts: find them, and delete only those whose every detail another contact already holds.

    from outlook_cdp import Session, Contacts, dedupe

    book = Contacts(Session("my-chrome-agent-instance"), mailbox="me@example.com")
    summary = book.backup("backups")                       # nothing is deleted without a backup
    backup = dedupe.load(summary["path"])
    plan = dedupe.plan(backup)                             # offline: reads the backup, not the server
    dedupe.write_plan(plan, "plan.json")                   # review it; approve name-only matches by hand
    log = dedupe.apply(book, dedupe.load("plan.json"), backup, "apply-log.jsonl")

Nothing here merges.  Merging is how a contact book fills with cruft: each merge piles another stale
number or note onto the survivor.  Instead a contact is deleted only when `coverage_gaps` finds nothing in
it that the kept contact lacks - every name part, number, address, note, category, extended property and
photo - so deleting it loses no information.  A duplicate holding even one thing the other lacks is left
alone and listed for review.

How things are compared, because "the same" has to be decided somewhere.  The rule throughout is that
formatting may differ but meaning may not, and that a value counts only where it is, not just anywhere:

- Text must match exactly but for runs of whitespace (Outlook stores notes with CRLF).  Case counts: a
  name spelled ``McDonald`` is not held by ``mcdonald``.  Email addresses and category names, which mail
  and Outlook treat case-insensitively, are the exceptions.
- A phone number must be in the *same kind of slot* - a business fax by a business fax - and match in
  digits, a leading NANP ``1`` dropped, with any extension (``x12``, ``ext. 12``) kept apart.
- An email address must be on the kept contact with the same display name (or none, or a name that is
  just the address or one of the kept contact's names), and the slot details beside it (routing type,
  mailbox type) must match.
- An address must be in the same slot (home, business, other) with every part this one has; a ZIP is held
  by itself or by the same ZIP+4, nothing else.
- Notes are held when the kept contact's notes contain them as whole words.
- A web page must match exactly, but for a trailing slash.
- A photo is held only by a byte-identical photo.  Any other attachment is never held.
- A property this module does not know, set to anything, is never held.  Unknown is not the same as empty.

Each contact is compared twice, through two independent views of it, and must pass both:
`coverage_gaps` runs `rest_gaps` on the REST record (with the MAPI properties in
`contacts.EXTENDED_PROPERTIES`) and `ews_gaps` on Outlook's own service.svc item, which keys numbers and
addresses by kind and carries properties REST does not.  What neither view can see - a MAPI property no
API exposes - is why deletion relies on its being soft: a deleted contact waits whole in Deleted Items.

Contacts are grouped when they share an email address or a phone number (seven digits or more).  Within a
group the richest contact is kept and each other one that it, or another kept one, fully holds is planned
for deletion.  Contacts that share only a name are never deleted automatically: they are listed with the
same verdict, for a person to approve by setting ``"approved": true`` in the plan.  A contact whose
email address is a member of a contact group is never deleted either, because the group may link to that
very contact; those are listed under ``in_a_contact_group``.

The plan file is the thing a person approves: its settings (``across_folders``, ``ignore_categories``,
``prefer_folder``) and its name-only approvals are trusted as that approval - ``ignore_categories`` in
particular decides which categories may be lost, so read it before applying.  Everything else `apply`
re-derives: it makes the plan again from the backup with those settings, refuses the whole plan if the
result differs, and only carries out deletions that recomputed plan contains.  It reads the contact groups live and protects their
members as they are now.  For each deletion it re-reads both contacts in both views, refuses if either has
changed since the backup, re-runs `coverage_gaps` on the live copies, deletes at once, and then confirms
the deleted one is gone and the kept one is still there.  The server takes no precondition on a delete, so
the gap between that last read and the delete can be narrowed but not closed; a contact edited inside it
is still whole in Deleted Items.  Every step is written to a JSON-lines log as it happens.

A person can also decide, on the review page (`dedupe_review`), to delete a contact although it holds
something its keeper lacks.  Such a decision names the keeper and lists exactly what will be lost;
`checked_decisions` refuses it unless that list is what the backup shows, and `apply` refuses it if the
live copies would lose anything more.
"""
from __future__ import annotations

import datetime
import json
import re

from .contacts import EXTENDED_PHONES, Contacts, extended
from .session import OutlookError

#: Server bookkeeping: not information about the person.
_IGNORED = {"Id", "ChangeKey", "CreatedDateTime", "LastModifiedDateTime", "ParentFolderId",
            "SingleValueExtendedProperties"}

#: Text fields that must match the kept contact's same field.
_TEXT = ["GivenName", "MiddleName", "Surname", "Generation", "NickName", "Title", "SpouseName", "JobTitle",
         "CompanyName", "Department", "OfficeLocation", "Profession", "Manager", "AssistantName",
         "YomiGivenName", "YomiSurname", "YomiCompanyName"]

_LISTS = ["ImAddresses", "Children", "Categories"]
_EMAIL_SLOTS = ["Email1Address", "Email2Address", "Email3Address"]
_ADDRESSES = ["HomeAddress", "BusinessAddress", "OtherAddress"]
_EXT_PAGES = {"personal_home_page", "web_page"}

#: Every top-level REST property `rest_gaps` knows how to compare.
_KNOWN = (_IGNORED | set(_TEXT) | set(_LISTS) | set(_ADDRESSES) | set(_EMAIL_SLOTS) |
          {"DisplayName", "FileAs", "Initials", "PersonalNotes", "Birthday", "EmailAddresses",
           "MobilePhone1", "HomePhones", "BusinessPhones", "BusinessHomePage"})

#: service.svc item properties that are bookkeeping, derived, or compared elsewhere (the photo).
_EWS_IGNORED = {"__type", "ItemId", "ParentFolderId", "ParentFolderDisplayName", "ConversationId",
                "ConversationGuid", "DateTimeCreated", "DateTimeReceived", "DateTimeSent", "LastModifiedTime",
                "LastModifiedName", "Size", "InstanceKey", "EffectiveRights", "WebClientReadFormQueryString",
                "PersonId", "SourceId", "ObjectId", "IsAssociated", "IsDraft", "IsFromMe", "IsResend",
                "IsSubmitted", "IsUnmodified", "IsLocked", "SupportsSideConversation",
                "IsDocumentCentricMessage", "IsDocumentCentricNotification", "Likes", "LikesPreview",
                "DisplayTo", "DisplayCc", "DisplayBcc", "EntityExtractionResult", "HasPicture",
                "BirthdayLocal", "WeddingAnniversaryLocal", "ItemClass"}

#: service.svc properties that are a way of writing the name, held by any of the kept contact's names.
_EWS_NAMES = {"Subject", "DisplayName", "FileAs", "Initials"}

#: Values service.svc reports for a property nobody set.
_EWS_DEFAULTS = {"Sensitivity": "Normal", "Importance": "Normal", "PostalAddressIndex": "None"}


def _s(v) -> str:
    """Text with runs of whitespace (and NBSP) collapsed; case kept."""
    return re.sub(r"\s+", " ", str(v or "").replace("\xa0", " ")).strip()


def _t(v) -> str:
    """Case-insensitive text: for email addresses, category names and grouping keys only."""
    return _s(v).casefold()


def _digits(number: str) -> str:
    d = re.sub(r"\D", "", number or "")
    return d[1:] if len(d) == 11 and d.startswith("1") else d


def _phone(number: str) -> str:
    """A number's meaning: its digits, any dial pauses (``,`` ``;``), then ``x`` and the extension if any."""
    low = (number or "").lower()
    m = re.search(r"(?:extension|ext\.?|x|#)\s*(\d+)\s*$", low)
    main, ext = (low[:m.start()], m.group(1)) if m else (low, "")
    head, *pauses = re.split(r"([,;])", main)          # a dial-string pause or wait is part of the meaning
    return _digits(head) + "".join(p if p in ",;" else re.sub(r"\D", "", p) for p in pauses) + \
        (f"x{ext}" if ext else "")


def _phones(entry: dict) -> dict[str, str]:
    """Every number on a contact, any slot: digits -> how it was written.  For grouping, not for holding."""
    r = entry["record"]
    raw = [r.get("MobilePhone1"), *(r.get("HomePhones") or []), *(r.get("BusinessPhones") or [])]
    raw += [v for k, v in extended(r).items() if k in EXTENDED_PHONES]
    return {_digits(p): p for p in raw if _digits(p)}


def _phone_slots(r: dict) -> dict[str, set[str]]:
    """Numbers by kind of slot, as `_phone` values."""
    slots = {"mobile": [r.get("MobilePhone1")], "home": r.get("HomePhones") or [],
             "business": r.get("BusinessPhones") or []}
    slots.update({k: [v] for k, v in extended(r).items() if k in EXTENDED_PHONES})
    return {k: {_phone(v) for v in vals if _phone(v)} for k, vals in slots.items()}


def _emails(entry: dict) -> set[str]:
    r = entry["record"]
    return {_t(e.get("Address")) for e in [*(r.get("EmailAddresses") or []), *(r.get(k) or {} for k in _EMAIL_SLOTS)]
            if e.get("Address")}


def _names(r: dict) -> set[str]:
    """The ways a contact's name can be written, for judging derived fields like DisplayName and FileAs."""
    g, m, s, gen = (r.get(k) or "" for k in ("GivenName", "MiddleName", "Surname", "Generation"))
    forms = {r.get("DisplayName"), r.get("FileAs"), r.get("NickName"), r.get("CompanyName"),
             " ".join(p for p in (g, m, s, gen) if p), " ".join(p for p in (g, s) if p),
             " ".join(p for p in (g, m, s) if p),
             f"{s}, {g}" if s and g else "", f"{s}, {g} {m}" if s and g and m else ""}
    return {_s(f) for f in forms if _s(f)}


def _initials(r: dict) -> set[str]:
    parts = [r.get(k) or "" for k in ("GivenName", "MiddleName", "Surname")]
    letters = [p[0] for p in parts if p]
    return {_s(r.get("Initials")), "".join(letters), ".".join(letters) + ".", ". ".join(letters) + "."}


def _zip_held(want: str, have: str) -> bool:
    return have == want or bool(re.fullmatch(re.escape(want) + r"-?\d{4}", have) and re.fullmatch(r"\d{5}", want))


def _addr_held(want: dict | None, have: dict | None) -> bool:
    have = {k: _s(v) for k, v in (have or {}).items()}
    for k, v in (want or {}).items():
        v = _s(v)
        if v and not (have.get(k) == v or (k == "PostalCode" and _zip_held(v, have.get(k, "")))):
            return False
    return True


def _notes_held(want: str, have: str) -> bool:
    want, have = _s(want), _s(have)
    return not want or bool(re.search(r"(?<!\w)" + re.escape(want) + r"(?!\w)", have))


def _page(url) -> str:
    u = _s(url)
    return u[:-1] if u.endswith("/") else u


def _name_held(value: str, names: set[str]) -> bool:
    return not _s(value) or _s(value) in names


def rest_gaps(dup: dict, keep: dict, ignore_categories=()) -> list[str]:
    """What `dup`'s REST record holds that `keep`'s does not.  See the module notes for the rules."""
    ignored = {_t(c) for c in ignore_categories}
    d, k = dup["record"], keep["record"]
    gaps: list[str] = []
    for key in sorted(set(d) - _KNOWN):
        if d[key] not in (None, "", [], {}):
            gaps.append(f"{key}: a property this check does not know, set to {str(d[key])[:60]!r}")
    for key in _TEXT:
        if _s(d.get(key)) and _s(d.get(key)) != _s(k.get(key)):
            gaps.append(f"{key}: {d.get(key)!r} (kept has {k.get(key)!r})")
    names = _names(k) | {_t(a) for a in _emails(keep)} | {_s(p) for p in _phones(keep).values()}
    for key in ("DisplayName", "FileAs"):
        if not _name_held(d.get(key), names | {_t(x) for x in [d.get(key)] if _t(x) in _emails(keep)}):
            gaps.append(f"{key}: {d.get(key)!r} is not one of the kept contact's names")
    if _s(d.get("Initials")) and _s(d.get("Initials")) not in _initials(k):
        gaps.append(f"Initials: {d.get('Initials')!r}")
    have = _phone_slots(k)
    for slot, numbers in _phone_slots(d).items():
        for n in sorted(numbers - have.get(slot, set())):
            gaps.append(f"{slot} phone {n!r}")
    for e in d.get("EmailAddresses") or []:
        addr, name = _t(e.get("Address")), _s(e.get("Name"))
        if not addr:
            continue
        if addr not in _emails(keep):
            gaps.append(f"email {e.get('Address')!r}")
            continue
        twins = [x for x in k.get("EmailAddresses") or [] if _t(x.get("Address")) == addr]

        def holds(x, e=e, name=name, addr=addr):
            """One kept entry holds the whole of this one: its name and every other filled field."""
            name_ok = not name or name == _s(x.get("Name")) or _t(name) == addr or name in names
            return name_ok and all(_s(v) == _s(x.get(f)) for f, v in e.items()
                                   if f not in ("Address", "Name") and _s(v))
        if not any(holds(x) for x in twins):
            gaps.append(f"email {e.get('Address')!r} as {({f: v for f, v in e.items() if _s(v)})}")
    kept_slots = [k.get(x) or {} for x in _EMAIL_SLOTS]
    for key in _EMAIL_SLOTS:
        slot = d.get(key) or {}
        if not slot.get("Address"):
            continue
        twin = next((x for x in kept_slots if _t(x.get("Address")) == _t(slot["Address"])), None)
        if twin is None:
            gaps.append(f"{key}: {slot['Address']!r}")
            continue
        for f, v in slot.items():
            if f == "Address" or not _s(v):
                continue
            if f in ("Name", "OriginalDisplayName"):
                if _s(v) != _s(twin.get(f)) and _t(v) not in _emails(keep) and _s(v) not in names:
                    gaps.append(f"{key} {f}: {v!r}")
            elif _s(v) != _s(twin.get(f)):
                gaps.append(f"{key} {f}: {v!r} (kept has {twin.get(f)!r})")
    for key in _LISTS:
        norm = _t if key == "Categories" else _s
        missing = {norm(v) for v in d.get(key) or [] if norm(v)} - {norm(v) for v in k.get(key) or []}
        if key == "Categories":
            missing -= ignored
        if missing:
            gaps.append(f"{key}: {sorted(missing)}")
    for key in _ADDRESSES:
        if not _addr_held(d.get(key), k.get(key)):
            gaps.append(f"{key}: {d.get(key)} (kept has {k.get(key)})")
    if not _notes_held(d.get("PersonalNotes"), k.get("PersonalNotes")):
        gaps.append(f"PersonalNotes: {str(d.get('PersonalNotes'))[:80]!r}")
    if d.get("Birthday") and (d.get("Birthday") or "")[:10] != (k.get("Birthday") or "")[:10]:
        gaps.append(f"Birthday: {d['Birthday'][:10]} (kept has {(k.get('Birthday') or '')[:10] or 'none'})")
    if _page(d.get("BusinessHomePage")) and _page(d.get("BusinessHomePage")) != _page(k.get("BusinessHomePage")):
        gaps.append(f"BusinessHomePage: {d.get('BusinessHomePage')!r}")
    dx, kx = extended(d), extended(k)
    for name, val in dx.items():
        if name in EXTENDED_PHONES:
            continue                                   # compared with the other numbers above
        norm = _page if name in _EXT_PAGES else _s
        if norm(val) != norm(kx.get(name)):
            gaps.append(f"{name}: {str(val)[:60]!r} (kept has {kx.get(name)!r})")
    if dup.get("photo") and (not keep.get("photo") or keep["photo"]["sha256"] != dup["photo"]["sha256"]):
        gaps.append("photo" + (" (kept has a different one)" if keep.get("photo") else " (kept has none)"))
    return gaps


def _ews_filled(v) -> bool:
    if isinstance(v, dict):
        return any(_ews_filled(x) for x in v.values())
    if isinstance(v, list):
        return any(_ews_filled(x) for x in v)
    return v not in (None, "", False) and _s(v) != ""


def _ews_entry_held(want: dict, have: list[dict], names: set[str], emails: set[str]) -> bool:
    """One keyed entry (a phone, address, email or IM address) is held by a kept entry with the same Key."""
    for h in have:
        if h.get("Key") != want.get("Key"):
            continue
        ok = True
        for f, v in want.items():
            if f == "Key" or not _s(v):
                continue
            got = h.get(f)
            if f == "PhoneNumber":
                ok = _phone(v) == _phone(got)
            elif f == "EmailAddress":
                ok = _t(v) == _t(got)
            elif f == "PostalCode":
                ok = _zip_held(_s(v), _s(got))
            elif f == "Name":
                ok = _s(v) == _s(got) or _t(v) in emails or _s(v) in names
            else:
                ok = _s(v) == _s(got)
            if not ok:
                break
        if ok:
            return True
    return False


def ews_gaps(dup: dict, keep: dict, ignore_categories=()) -> list[str]:
    """What `dup`'s service.svc item holds that `keep`'s does not.  Both entries need an ``"ews"`` item."""
    if not dup.get("ews"):
        return ["no service.svc view of the duplicate to compare"]
    if not keep.get("ews"):
        return ["no service.svc view of the kept contact to compare"]
    d, k = dup["ews"], keep["ews"]
    ignored = {_t(c) for c in ignore_categories}
    names = _names(keep["record"]) | {_s(v) for key in _EWS_NAMES for v in [k.get(key)] if _s(v)}
    names |= {_s(v) for v in (k.get("CompleteName") or {}).values() if isinstance(v, str) and _s(v)}
    emails = _emails(keep)
    names |= {_s(p) for p in _phones(keep).values()}
    gaps: list[str] = []
    for key, val in sorted(d.items()):
        kv = k.get(key)
        if key in _EWS_DEFAULTS:
            if (val or _EWS_DEFAULTS[key]) != (kv or _EWS_DEFAULTS[key]):
                gaps.append(f"{key}: {val!r} (kept has {kv!r})")
            continue
        if key == "HasAttachments":
            if val and not d.get("Attachments"):
                gaps.append("HasAttachments is set but no attachment details came back")
            continue
        if key in _EWS_IGNORED or not _ews_filled(val):
            continue
        if key in _EWS_NAMES:
            if not (_name_held(val, names) or _t(val) in emails) and _s(val) not in _initials(keep["record"]):
                gaps.append(f"{key}: {val!r} is not one of the kept contact's names")
        elif key == "CompleteName":
            for f, v in val.items():
                if isinstance(v, str) and _s(v) and _s(v) != _s((kv or {}).get(f)) and _s(v) not in names:
                    gaps.append(f"CompleteName {f}: {v!r}")
        elif key in ("PhoneNumbers", "PhysicalAddresses", "EmailAddresses", "ImAddresses"):
            for entry in val:
                if _ews_filled({f: v for f, v in entry.items() if f != "Key"}) and \
                        not _ews_entry_held(entry, kv or [], names, emails):
                    gaps.append(f"{key} {entry.get('Key')}: " + ", ".join(
                        f"{v}" for f, v in entry.items() if f != "Key" and _s(v))[:80])
        elif key == "Categories":
            missing = {_t(c) for c in val} - {_t(c) for c in kv or []} - ignored
            if missing:
                gaps.append(f"Categories: {sorted(missing)}")
        elif key == "Body":
            if not _notes_held(val.get("Value"), (kv or {}).get("Value")):
                gaps.append(f"Body: {str(val.get('Value'))[:80]!r}")
            for f, v in val.items():
                if f not in ("Value", "IsTruncated") and _s(v) and _s(v) != _s((kv or {}).get(f)):
                    gaps.append(f"Body {f}: {v!r} (kept has {(kv or {}).get(f)!r})")
            if val.get("IsTruncated"):
                gaps.append("Body came back truncated, so it cannot be compared")
        elif key == "PersonalNotes":
            if not _notes_held(val, kv):
                gaps.append(f"PersonalNotes: {str(val)[:80]!r}")
        elif key in ("Birthday", "WeddingAnniversary"):
            if str(val)[:10] != str(kv or "")[:10]:
                gaps.append(f"{key}: {str(val)[:10]} (kept has {str(kv or '')[:10] or 'none'})")
        elif key == "Attachments":
            others = [a.get("Name") for a in val if not a.get("IsContactPhoto")]
            if others:
                gaps.append(f"attachments other than the photo: {others}")
        elif key in ("BusinessHomePage", "PersonalHomePage"):
            if _page(val) != _page(kv):
                gaps.append(f"{key}: {val!r}")
        elif isinstance(val, list):
            missing = {_s(json.dumps(v, sort_keys=True)) for v in val if _ews_filled(v)} - \
                      {_s(json.dumps(v, sort_keys=True)) for v in kv or []}
            if missing:
                gaps.append(f"{key}: {sorted(missing)[:3]}")
        elif isinstance(val, dict):
            if json.dumps(val, sort_keys=True) != json.dumps(kv, sort_keys=True):
                gaps.append(f"{key}: {json.dumps(val)[:80]} (kept has {json.dumps(kv)[:80]})")
        elif key == "ReminderMinutesBeforeStart":
            if d.get("ReminderIsSet") and val != kv:
                gaps.append(f"{key}: {val}")
        elif _s(val) != _s(kv):
            gaps.append(f"{key}: {str(val)[:60]!r} (kept has {str(kv)[:60]!r})")
    return gaps


def coverage_gaps(dup: dict, keep: dict, ignore_categories=()) -> list[str]:
    """What `dup` holds that `keep` does not, through both views.  Empty means deleting `dup` loses nothing.

    Both are backup entries: ``{"record": <REST contact>, "photo": {...} | None, "ews": <service.svc item>}``.
    Categories named in `ignore_categories` are treated as not information - markers an earlier tool
    stamped on, say.
    """
    gaps = rest_gaps(dup, keep, ignore_categories)
    gaps += [f"(service.svc) {g}" for g in ews_gaps(dup, keep, ignore_categories)]
    return gaps


def richness(entry: dict) -> int:
    """How much a contact holds: the count of filled values, used to choose which of a group to keep."""
    r = entry["record"]
    n = sum(1 for key, v in r.items() if key not in _IGNORED and v not in (None, "", [], {}))
    n += sum(len(r.get(key) or []) for key in ("EmailAddresses", "HomePhones", "BusinessPhones"))
    return n + len(extended(r)) + (2 if entry.get("photo") else 0)


def display(entry: dict) -> str:
    r = entry["record"]
    return (r.get("DisplayName") or " ".join(p for p in (r.get("GivenName"), r.get("Surname")) if p)
            or r.get("CompanyName") or next(iter(sorted(_emails(entry))), "")
            or next(iter(_phones(entry).values()), "") or "(no name)")


def _union(entries: list[dict], keys) -> list[list[int]]:
    parent = list(range(len(entries)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    seen: dict[str, int] = {}
    for i, e in enumerate(entries):
        for key in keys(e):
            if key in seen:
                parent[find(i)] = find(seen[key])
            else:
                seen[key] = i
    groups: dict[int, list[int]] = {}
    for i in range(len(entries)):
        groups.setdefault(find(i), []).append(i)
    return [g for g in groups.values() if len(g) > 1]


def _strong_keys(e: dict) -> list[str]:
    return [f"email:{a}" for a in _emails(e)] + [f"phone:{d}" for d in _phones(e) if len(d) >= 7]


def _name_key(e: dict) -> list[str]:
    r = e["record"]
    name = _t(" ".join(p for p in (r.get("GivenName"), r.get("Surname")) if p)) or _t(r.get("DisplayName"))
    return [f"name:{name}"] if name and "@" not in name else []


def _settle(entries: list[dict], members: list[int], folders: dict, ignore=(),
            prefer_folder=None) -> tuple[list[dict], list[dict]]:
    """Choose what to keep in one group: (deletions, kept) - each deletion names the kept contact holding it.

    Richest first, newest on a tie.  When a contact is held by a kept one that it holds right back - the
    same information - and only it is in `prefer_folder`, the two swap, provided it also holds everything
    already planned against the other.  So of identical copies, the one in the preferred folder stays.
    """
    order = sorted(members, key=lambda i: entries[i]["record"].get("LastModifiedDateTime") or "", reverse=True)
    order = sorted(order, key=lambda i: -richness(entries[i]))      # richest first; ties: newest first
    gaps = lambda a, b: coverage_gaps(entries[a], entries[b], ignore)  # noqa: E731
    kept: list[int] = []
    held_by: dict[int, list[int]] = {}
    for i in order:
        holder = next((k for k in kept if not gaps(i, k)), None)
        if holder is None:
            kept.append(i)
            held_by[i] = []
        elif (entries[i].get("folder") == prefer_folder != entries[holder].get("folder") and not gaps(holder, i)
              and not any(gaps(d, i) for d in held_by[holder])):
            kept[kept.index(holder)] = i
            held_by[i] = held_by.pop(holder) + [holder]
        else:
            held_by[holder].append(i)
    deletions = [_item(entries[d], entries[k], folders) for k in kept for d in held_by[k]]
    return deletions, [entries[k] for k in kept]


def _ref(e: dict, folders: dict) -> dict:
    r = e["record"]
    return {"id": r["Id"], "change_key": r.get("ChangeKey"), "name": display(e),
            "folder": folders.get(e.get("folder"), "Contacts")}


def _item(dup: dict, keep: dict, folders: dict) -> dict:
    return {"delete": _ref(dup, folders), "keep": _ref(keep, folders)}


def group_members(backup: dict) -> set[str]:
    """Email addresses that are members of a contact group in the backup, lowercased."""
    return {_t((m.get("Mailbox") or {}).get("EmailAddress"))
            for g in backup.get("groups") or [] for m in g["item"].get("Members") or []
            if (m.get("Mailbox") or {}).get("EmailAddress")}


def load(path: str) -> dict:
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def plan(backup: dict, *, across_folders: bool = True, ignore_categories=(), prefer_folder=None) -> dict:
    """Plan deletions from a backup, without touching the server.

    With `across_folders` false, a contact is only ever held by one in the same folder.
    `ignore_categories` is passed to `coverage_gaps`, and recorded in the plan so `apply` uses it too.
    Of identical copies, the one in `prefer_folder` is kept: by default the default Contacts folder,
    which is the one phones syncing over Exchange usually show.
    """
    ignore = sorted(ignore_categories)
    entries = backup["contacts"]
    folders = {f["id"]: f["name"] for f in backup["folders"]}
    scope = (lambda e: [""]) if across_folders else (lambda e: [str(e.get("folder"))])
    keyed = lambda fn: (lambda e: [s + "|" + k for s in scope(e) for k in fn(e)])  # noqa: E731
    out = {"format": 1, "backup_taken_at": backup["taken_at"], "mailbox": backup["mailbox"],
           "across_folders": across_folders, "ignore_categories": ignore,
           "prefer_folder": folders.get(prefer_folder, "Contacts"), "prefer_folder_id": prefer_folder, "delete": [], "in_a_contact_group": [],
           "kept_with_differences": [], "name_only": []}
    members_of_groups = group_members(backup)
    in_group = {entries[i]["record"]["Id"] for i in range(len(entries)) if _emails(entries[i]) & members_of_groups}
    strong_of: dict[int, int] = {}
    for n, members in enumerate(_union(entries, keyed(_strong_keys))):
        strong_of.update({i: n for i in members})
        deletions, kept = _settle(entries, members, folders, ignore, prefer_folder)
        out["delete"] += [d for d in deletions if d["delete"]["id"] not in in_group]
        out["in_a_contact_group"] += [d for d in deletions if d["delete"]["id"] in in_group]
        if len(kept) > 1:
            out["kept_with_differences"].append({
                "contacts": [_ref(e, folders) for e in kept],
                "each_holds_that_the_first_lacks": {display(e): coverage_gaps(e, kept[0], ignore) for e in kept[1:]}})
    for members in _union(entries, keyed(_name_key)):
        if len({strong_of.get(m, -1 - m) for m in members}) == 1:
            continue                                   # already one email/phone group
        deletions, kept = _settle(entries, members, folders, ignore, prefer_folder)
        for d in deletions:
            d["approved"] = False
        out["name_only"].append({
            "deletable": deletions,
            "kept": [_ref(e, folders) for e in kept],
            "each_holds_that_the_first_lacks": {display(e): coverage_gaps(e, kept[0], ignore) for e in kept[1:]}})
    out["name_only"] = [g for g in out["name_only"] if g["deletable"] or len(g["kept"]) > 1]
    planned = {d["delete"]["id"] for d in out["delete"]}
    for g in out["name_only"]:
        g["deletable"] = [d for d in g["deletable"] if d["delete"]["id"] not in planned | in_group]
    out["summary"] = {"contacts": len(entries), "delete": len(out["delete"]),
                      "in_a_contact_group": len(out["in_a_contact_group"]),
                      "kept_with_differences": len(out["kept_with_differences"]),
                      "name_only_groups": len(out["name_only"]),
                      "name_only_deletable": sum(len(g["deletable"]) for g in out["name_only"])}
    return out


def write_plan(p: dict, path: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(p, fh, ensure_ascii=False, indent=1)


def approved(p: dict) -> list[dict]:
    """The plan's deletions: every email/phone match, plus the name-only ones a person approved."""
    return p["delete"] + [d for g in p["name_only"] for d in g["deletable"] if d.get("approved") is True]


def _settings(p: dict) -> dict:
    return {"across_folders": p["across_folders"], "ignore_categories": p["ignore_categories"],
            "prefer_folder": p.get("prefer_folder_id")}


def _pairs(p: dict) -> set[tuple[str, str]]:
    return {(d["delete"]["id"], d["keep"]["id"]) for d in p["delete"]}


def _name_pairs(p: dict) -> set[tuple[str, str]]:
    return {(d["delete"]["id"], d["keep"]["id"]) for g in p["name_only"] for d in g["deletable"]}


def review_groups(p: dict) -> list[list[str]]:
    """The groups a plan leaves for a person: each a list of contact ids, as the review page shows them."""
    return ([[c["id"] for c in x["contacts"]] for x in p["kept_with_differences"]] +
            [[c["id"] for c in g["kept"]] + [d["delete"]["id"] for d in g["deletable"]] for g in p["name_only"]])


def checked_decisions(p: dict, backup: dict, decisions: dict) -> list[dict]:
    """A person's decisions from the review page, checked against the backup, as deletions with accepted loss.

    Each decision deletes one contact in favour of another in the same review group, and lists exactly what
    the person accepted losing.  That list must equal what `coverage_gaps` finds in the backup now, so a
    decision cannot accept less than it loses, and a stale or edited decisions file is refused whole.
    """
    if decisions.get("backup_taken_at") != backup["taken_at"] or decisions.get("mailbox") != backup["mailbox"]:
        raise OutlookError("these decisions were not made from this backup")
    together = {}
    for n, ids in enumerate(review_groups(p)):
        for i in ids:
            together.setdefault(i, set()).add(n)
    by = {e["record"]["Id"]: e for e in backup["contacts"]}
    folders = {f["id"]: f["name"] for f in backup["folders"]}
    ignore = p.get("ignore_categories") or ()
    out = []
    for d in decisions.get("decisions") or []:
        dup, keep = d.get("delete"), d.get("keep")
        if dup == keep or dup not in by or keep not in by or not together.get(dup, set()) & together.get(keep, set()):
            raise OutlookError(f"a decision pairs contacts that are not in one review group: {dup} / {keep}")
        loss = coverage_gaps(by[dup], by[keep], ignore)
        if sorted(loss) != sorted(d.get("accepted_loss") or []):
            raise OutlookError(f"a decision about {display(by[dup])!r} accepts a different loss than the backup "
                               f"shows; make the review page again from this backup and plan")
        out.append({**_item(by[dup], by[keep], folders), "accepted_loss": loss})
    deleted = [d["delete"]["id"] for d in out]
    if len(set(deleted)) != len(deleted):
        raise OutlookError("the decisions delete one contact twice")
    return out


def apply(book: Contacts, p: dict, backup: dict, log_path: str, *, limit: int | None = None,
          decisions: dict | None = None) -> dict:
    """Carry out a reviewed plan, re-verifying every deletion against the live book.  See the module notes.

    `decisions` (from the review page) adds deletions a person chose despite a loss they accepted; each is
    carried out only if the live copies lose no more than that.
    """
    if p["backup_taken_at"] != backup["taken_at"] or p["mailbox"] != backup["mailbox"]:
        raise OutlookError("this plan was not made from this backup")
    if book.mailbox != backup["mailbox"]:
        raise OutlookError(f"the backup is of {backup['mailbox']}, not {book.mailbox}")
    again = plan(backup, **_settings(p))
    if _pairs(again) != _pairs(p) or _name_pairs(again) - _name_pairs(p):
        raise OutlookError("the plan does not match one made again from the backup with its own settings; "
                           "it was edited or made by other code. Make it again with dedupe-plan.")
    allowed = _pairs(again) | _name_pairs(again)
    ignore = p.get("ignore_categories") or ()
    saved = {e["record"]["Id"]: e for e in backup["contacts"]}
    members_now = {_t((m.get("Mailbox") or {}).get("EmailAddress"))
                   for f in book.folders() for g in book.groups(f["id"]) for m in g.get("Members") or []
                   if (m.get("Mailbox") or {}).get("EmailAddress")}
    chosen = checked_decisions(p, backup, decisions) if decisions else []
    allowed |= {(d["delete"]["id"], d["keep"]["id"]) for d in chosen}
    everything = approved(p) + chosen
    todo = everything[:limit] if limit else everything
    deleting = {d["delete"]["id"] for d in todo}
    tally = {"deleted": 0, "skipped": 0}
    with open(log_path, "a", encoding="utf-8") as log:
        def note(item, outcome, why=""):
            tally["deleted" if outcome == "deleted" else "skipped"] += 1
            log.write(json.dumps({"at": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
                                  "outcome": outcome, "why": why, **item}, ensure_ascii=False) + "\n")
            log.flush()

        def live(cid):
            """Both views of one contact as it is now, or None; and whether it is as the backup has it."""
            record = book.get_raw(cid)
            if record is None:
                return None, False
            was = saved[cid]
            items = book.get_items([was["ews"]["ItemId"]["Id"]]) if was.get("ews") else []
            ews = items[0] if items else None
            same = (record.get("ChangeKey") == was["record"].get("ChangeKey") and ews is not None
                    and ews["ItemId"].get("ChangeKey") == record.get("ChangeKey"))
            return {"record": record, "ews": ews, "folder": was.get("folder")}, same

        for item in todo:
            dup_id, keep_id = item["delete"]["id"], item["keep"]["id"]
            if (dup_id, keep_id) not in allowed:
                note(item, "skipped", "not a deletion the backup supports")
                continue
            if keep_id in deleting:
                note(item, "skipped", "the kept contact is itself planned for deletion")
                continue
            if dup_id not in saved or keep_id not in saved:
                note(item, "skipped", "not in the backup, so it could not be restored")
                continue
            if _emails(saved[dup_id]) & members_now:
                note(item, "skipped", "a member of a contact group")
                continue
            dup, dup_same = live(dup_id)
            keep, keep_same = live(keep_id)
            if dup is None:
                note(item, "skipped", "already gone")
                continue
            if keep is None:
                note(item, "skipped", "the kept contact is gone")
                continue
            if not (dup_same and keep_same):
                note(item, "skipped", "changed since the backup; take a new backup and plan")
                continue
            if _emails(dup) & members_now:
                note(item, "skipped", "a member of a contact group")
                continue
            pics = book.photos([dup_id, keep_id])
            dup["photo"], keep["photo"] = pics[dup_id], pics[keep_id]
            gaps = coverage_gaps(dup, keep, ignore)
            extra = sorted(set(gaps) - set(item.get("accepted_loss") or []))
            if extra:
                note(item, "skipped", "live copy holds more than the kept one: " + "; ".join(extra))
                continue
            book.delete(dup_id)
            if book.get_raw(dup_id) is not None:
                note(item, "failed", "still present after DELETE")
                continue
            if book.get_raw(keep_id) is None:
                raise OutlookError(f"the kept contact {item['keep']['name']!r} vanished after deleting its "
                                   f"duplicate; stopping. Restore from the backup.")
            note(item, "deleted")
    return tally
