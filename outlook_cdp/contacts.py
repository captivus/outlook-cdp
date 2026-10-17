"""Contacts: list, find, create and update with read-back verification, photos, delete - over REST.

    from outlook_cdp import Session, Contacts, Contact

    book = Contacts(Session("my-chrome-agent-instance"), mailbox="me@example.com")
    for c in book.find_by_phone("(555) 555-0100"):          # digits compared, formatting ignored
        print(c.display_name, c.mobile_phone)
    made = book.create(Contact(given_name="Pat", surname="Example", mobile_phone="+1 555-555-0100",
                               personal_notes="Met at the club."))
    book.update(made.id, nickname="")                  # "" or [] clears a field
    book.set_photo(made.id, open("face.jpg", "rb").read(), "image/jpeg")

Six things about this that are not obvious:

1. **The token already allows it.**  The access token Outlook caches for its own REST calls carries
   ``Contacts.ReadWrite``, so ``/contacts`` takes a POST from inside the page with no consent prompt and
   no People UI to drive.
2. **The fixed-size slots are small.**  A contact holds one mobile number, two home numbers, two
   business numbers and three email addresses.  `to_rest` refuses more than fit rather than letting the
   server keep some and drop the rest without saying so; put the overflow in `personal_notes`.
3. **A stored contact is not proof of the contact you sent**, for the same reason as rules: so `create`
   reads it back, compares every field it sent, and deletes it again on any difference.
4. **Phone numbers are free text.**  The same number is stored as ``+1 (555) 555-0100``,
   ``555.555.0100`` or with a non-breaking space in it.  `find_by_phone` compares the last ten digits,
   and `create` refuses a contact whose number is already on another one unless told otherwise, so
   re-running a script does not stack up copies.
5. **An empty value in a PATCH is ignored, not applied.**  Clearing a field takes an explicit ``null``
   (``[]`` for a list), which `update` sends for any field given as ``""`` or ``[]``, then checks it.
6. **A photo cannot be handed to the page in one CDP call.**  chrome-agent puts each call on its command
   line and one argument is capped at 128 KiB, so `set_photo` moves the bytes in base64 slices and
   confirms the upload by reading the stored photo's dimensions.

The notes field is ``PersonalNotes``, plain text; Outlook shows it under "Notes" on the contact card.
Like the rest of the REST path, every call is scoped to ``/api/v2.0/users/<mailbox>/``.
"""
from __future__ import annotations

import base64
import json
import re
import urllib.parse
from dataclasses import dataclass, field

from .attachments import access_token, rest_get
from .rules import rest_send
from .session import OutlookError, Session

#: The slots a contact has for each kind of value; the server does not take more.
MAX_HOME_PHONES = 2
MAX_BUSINESS_PHONES = 2
MAX_EMAILS = 3

#: Scalar fields: dataclass attribute -> REST property.
_SCALARS = {
    "given_name": "GivenName",
    "middle_name": "MiddleName",
    "surname": "Surname",
    "generation": "Generation",
    "nickname": "NickName",
    "display_name": "DisplayName",
    "spouse_name": "SpouseName",
    "job_title": "JobTitle",
    "company_name": "CompanyName",
    "mobile_phone": "MobilePhone1",
    "personal_notes": "PersonalNotes",
}

_ADDRESS_KEYS = {"street": "Street", "city": "City", "state": "State", "postal_code": "PostalCode",
                 "country": "CountryOrRegion"}

#: Base64 characters per CDP call when moving photo bytes into the page.
_SLICE = 60_000

#: PUT the parked photo slices as raw bytes; the global is released either way.
_PHOTO_JS = """(async()=>{try{
  const bin=atob(window.__ocdpPhoto.join('')); delete window.__ocdpPhoto;
  const bytes=new Uint8Array(bin.length); for(let i=0;i<bin.length;i++) bytes[i]=bin.charCodeAt(i);
  const r=await fetch(%s,{method:'PUT',headers:{authorization:'Bearer '+%s,'content-type':%s},body:bytes});
  return {status:r.status, body:await r.text()}}catch(e){delete window.__ocdpPhoto;return {status:0, body:String(e)}}})()"""

_SELECT = ",".join(["Id", "HomePhones", "BusinessPhones", "EmailAddresses", "HomeAddress",
                    "OtherAddress", "Categories", "Birthday", *_SCALARS.values()])


def phone_digits(number: str) -> str:
    """The comparable part of a phone number: its last ten digits (a NANP number without the 1)."""
    d = re.sub(r"\D", "", number or "")
    return d[-10:] if len(d) >= 10 else d


@dataclass
class Contact:
    """One contact, in the shape this module can send, read back and compare.

    Addresses are dicts with any of ``street``, ``city``, ``state``, ``postal_code``, ``country``.
    `birthday` is ``YYYY-MM-DD``.
    """
    given_name: str = ""
    middle_name: str = ""
    surname: str = ""
    generation: str = ""
    nickname: str = ""
    display_name: str = ""
    spouse_name: str = ""
    job_title: str = ""
    company_name: str = ""
    mobile_phone: str = ""
    home_phones: list[str] = field(default_factory=list)
    business_phones: list[str] = field(default_factory=list)
    emails: list[str] = field(default_factory=list)
    home_address: dict = field(default_factory=dict)
    other_address: dict = field(default_factory=dict)
    birthday: str = ""
    categories: list[str] = field(default_factory=list)
    personal_notes: str = ""
    id: str | None = None

    def phones(self) -> list[str]:
        return [p for p in [self.mobile_phone, *self.home_phones, *self.business_phones] if p]

    def name(self) -> str:
        return self.display_name or " ".join(
            p for p in (self.given_name, self.middle_name, self.surname, self.generation) if p)

    def to_rest(self) -> dict:
        if not (self.name() or self.phones() or self.emails):
            raise OutlookError("a contact with no name, phone or email would be an empty card. Refusing.")
        for label, vals, cap in (("home phones", self.home_phones, MAX_HOME_PHONES),
                                 ("business phones", self.business_phones, MAX_BUSINESS_PHONES),
                                 ("email addresses", self.emails, MAX_EMAILS)):
            if len(vals) > cap:
                raise OutlookError(f"{len(vals)} {label} given; a contact holds {cap}. Put the rest in "
                                   f"personal_notes rather than have the server drop them.")
        out: dict = {rest: getattr(self, key) for key, rest in _SCALARS.items() if getattr(self, key)}
        if not self.display_name and self.name():
            out["DisplayName"] = self.name()
        if self.home_phones:
            out["HomePhones"] = list(self.home_phones)
        if self.business_phones:
            out["BusinessPhones"] = list(self.business_phones)
        if self.emails:
            out["EmailAddresses"] = [{"Name": a, "Address": a} for a in self.emails]
        for key, rest in (("home_address", "HomeAddress"), ("other_address", "OtherAddress")):
            addr = getattr(self, key)
            unknown = set(addr) - set(_ADDRESS_KEYS)
            if unknown:
                raise OutlookError(f"{key} has keys this module does not send: {sorted(unknown)}")
            if addr:
                out[rest] = {_ADDRESS_KEYS[k]: v for k, v in addr.items() if v}
        if self.birthday:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", self.birthday):
                raise OutlookError(f"birthday {self.birthday!r} is not YYYY-MM-DD")
            out["Birthday"] = self.birthday + "T11:59:00Z"   # midday, so no timezone moves the date
        if self.categories:
            out["Categories"] = list(self.categories)
        return out

    @classmethod
    def from_rest(cls, raw: dict) -> "Contact":
        kw = {key: raw.get(rest) or "" for key, rest in _SCALARS.items()}
        addr = lambda a: {k: (a or {}).get(r) for k, r in _ADDRESS_KEYS.items() if (a or {}).get(r)}  # noqa: E731
        return cls(**kw,
                   home_phones=[p for p in raw.get("HomePhones") or [] if p],
                   business_phones=[p for p in raw.get("BusinessPhones") or [] if p],
                   emails=[e.get("Address") for e in raw.get("EmailAddresses") or [] if e.get("Address")],
                   home_address=addr(raw.get("HomeAddress")), other_address=addr(raw.get("OtherAddress")),
                   birthday=(raw.get("Birthday") or "")[:10], categories=list(raw.get("Categories") or []),
                   id=raw.get("Id"))


def _norm(s: str) -> str:
    """Whitespace-insensitive text: Outlook may store notes with CRLF and phones with NBSP."""
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()


def differences(asked: Contact, stored: Contact) -> list[str]:
    """Fields the caller set that the server did not store as sent.  Empty means it took them all."""
    probs = []
    for key in _SCALARS:
        want = getattr(asked, key)
        if key == "display_name":
            want = asked.name()
        if want and _norm(want) != _norm(getattr(stored, key)):
            probs.append(f"{key}: asked {want!r}, stored {getattr(stored, key)!r}")
    for key in ("home_phones", "business_phones", "categories"):
        want = [_norm(v) for v in getattr(asked, key)]
        if want and want != [_norm(v) for v in getattr(stored, key)]:
            probs.append(f"{key}: asked {getattr(asked, key)}, stored {getattr(stored, key)}")
    if asked.emails and [e.lower() for e in asked.emails] != [e.lower() for e in stored.emails]:
        probs.append(f"emails: asked {asked.emails}, stored {stored.emails}")
    for key in ("home_address", "other_address"):
        want = {k: _norm(v) for k, v in getattr(asked, key).items() if v}
        got = {k: _norm(v) for k, v in getattr(stored, key).items() if v}
        if want and want != got:
            probs.append(f"{key}: asked {getattr(asked, key)}, stored {getattr(stored, key)}")
    if asked.birthday and asked.birthday != stored.birthday:
        probs.append(f"birthday: asked {asked.birthday}, stored {stored.birthday}")
    return probs


class Contacts:
    """The default contacts folder of one mailbox.

    `mailbox` defaults to the one in the tab's URL and is refused when there is none, as in `Rules`;
    ``Rules.signed_in_address(session)`` gives the signed-in user's own address.
    """

    def __init__(self, session: Session, mailbox: str | None = None):
        self.s = session
        self.mailbox = (mailbox or session.mailbox_identity() or "").lower() or None
        if not self.mailbox:
            raise OutlookError(
                "cannot tell which mailbox this tab is showing: its URL names none. Pass mailbox= "
                "explicitly (Rules.signed_in_address(session) gives the signed-in user's address).")
        self.base = "/api/v2.0/users/" + urllib.parse.quote(self.mailbox, safe="") + "/contacts"
        self._token: str | None = None

    def token(self) -> str:
        if self._token is None:
            self._token = access_token(self.s)
        return self._token

    def list(self) -> list[Contact]:
        """Every contact in the default folder, following server paging."""
        out: list[Contact] = []
        path: str | None = f"{self.base}?$top=500&$select={_SELECT}"
        while path:
            page = rest_get(self.s, path, self.token())
            out += [Contact.from_rest(r) for r in page.get("value") or []]
            nxt = page.get("@odata.nextLink")
            path = ("/" + nxt.split("/", 3)[3]) if nxt else None
        return out

    def get(self, contact_id: str) -> Contact:
        return Contact.from_rest(rest_get(
            self.s, f"{self.base}/{urllib.parse.quote(contact_id, safe='')}?$select={_SELECT}", self.token()))

    def find_by_phone(self, number: str, contacts: list[Contact] | None = None) -> list[Contact]:
        """Contacts carrying `number` in any phone slot, compared on the last ten digits."""
        want = phone_digits(number)
        if len(want) < 7:
            raise OutlookError(f"{number!r} has too few digits to match on")
        return [c for c in (self.list() if contacts is None else contacts)
                if any(phone_digits(p) == want for p in c.phones())]

    def find_by_name(self, text: str, contacts: list[Contact] | None = None) -> list[Contact]:
        """Contacts whose display name contains `text`, case-insensitively."""
        return [c for c in (self.list() if contacts is None else contacts)
                if text.lower() in c.name().lower()]

    def create(self, contact: Contact, *, allow_duplicate_phone: bool = False) -> Contact:
        """Create `contact`, read it back, and refuse - deleting it again - if the server changed it.

        Refuses when any of its numbers is already on another contact, unless `allow_duplicate_phone`.
        """
        payload = contact.to_rest()
        if not allow_duplicate_phone and contact.phones():
            existing = self.list()
            clash = {c.name() or c.id for p in contact.phones() for c in self.find_by_phone(p, existing)}
            if clash:
                raise OutlookError(f"a number on this contact is already on: {sorted(clash)}")
        made = rest_send(self.s, "POST", self.base, payload, self.token())
        if not made.get("Id"):
            raise OutlookError(f"contact POST returned no Id: {json.dumps(made)[:300]}")
        stored = self.get(made["Id"])
        probs = differences(contact, stored)
        if probs:
            self.delete(made["Id"])
            raise OutlookError("the server stored a different contact than was asked for, so it was "
                               "deleted again: " + "; ".join(probs))
        return stored

    def update(self, contact_id: str, **changes) -> Contact:
        """Change the named fields of a stored contact, read it back, and refuse if any did not take.

        A field set to ``""`` or ``[]`` is cleared on the server, not left alone.  The stored display name
        is not re-derived from the name parts: pass `display_name` as well when changing them.
        """
        current = self.get(contact_id)
        for key, val in changes.items():
            if key == "id" or not hasattr(current, key):
                raise OutlookError(f"{key!r} is not a contact field")
            setattr(current, key, val)
        payload = current.to_rest()
        for key, val in changes.items():
            if val:
                continue
            if key in _SCALARS:
                payload[_SCALARS[key]] = None
            elif key in ("home_phones", "business_phones", "categories"):
                payload[{"home_phones": "HomePhones", "business_phones": "BusinessPhones",
                         "categories": "Categories"}[key]] = []
            elif key == "emails":
                payload["EmailAddresses"] = []
        rest_send(self.s, "PATCH", f"{self.base}/{urllib.parse.quote(contact_id, safe='')}", payload,
                  self.token())
        stored = self.get(contact_id)
        probs = differences(current, stored)
        probs += [f"{k}: asked to clear, stored {getattr(stored, k)!r}"
                  for k, v in changes.items() if not v and getattr(stored, k)]
        if probs:
            raise OutlookError("the server did not store the update as sent: " + "; ".join(probs))
        return stored

    def set_photo(self, contact_id: str, data: bytes, content_type: str = "image/jpeg") -> dict:
        """Upload `data` as the contact's picture and return the stored photo's metadata.

        The bytes go into the page in base64 slices parked on a page global, because chrome-agent
        passes each CDP call on its command line, and one argument is capped at 128 KiB.
        """
        if not data or content_type not in ("image/jpeg", "image/png", "image/gif"):
            raise OutlookError(f"refusing an empty photo or an unsupported type {content_type!r}")
        b64 = base64.b64encode(data).decode()
        self.s.js("window.__ocdpPhoto=[]")
        for i in range(0, len(b64), _SLICE):
            self.s.js(f"window.__ocdpPhoto.push({json.dumps(b64[i:i + _SLICE])})")
        path = f"{self.base}/{urllib.parse.quote(contact_id, safe='')}/photo/$value"
        out = self.s.js(_PHOTO_JS % (json.dumps(path), json.dumps(self.token()), json.dumps(content_type)),
                        await_promise=True) or {}
        if not 200 <= (out.get("status") or 0) < 300:
            raise OutlookError(f"photo upload returned {out.get('status')} {out.get('body', '')[:300]}")
        meta = rest_get(self.s, f"{self.base}/{urllib.parse.quote(contact_id, safe='')}/photo", self.token())
        if not meta.get("Width"):
            raise OutlookError(f"photo upload reported success but no photo is stored: {meta}")
        return meta

    def delete(self, contact_id: str) -> None:
        rest_send(self.s, "DELETE", f"{self.base}/{urllib.parse.quote(contact_id, safe='')}",
                  token=self.token())
