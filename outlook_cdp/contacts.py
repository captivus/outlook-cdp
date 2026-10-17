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

`Contact` is the shape for *writing*.  For keeping or comparing what is already there, use the raw
records from `list_raw` / `backup`, because the REST contact leaves out much of an Outlook contact:

7. **Half the phone slots are not on the REST contact.**  Other, fax, pager, car, callback, assistant,
   company-main and primary numbers, the personal web page, the anniversary and the four user fields
   exist only as MAPI properties.  `list_raw` asks for them (`EXTENDED_PROPERTIES`) as single-value
   extended properties, so a card holding nothing but a fax number does not look empty.
8. **Contacts live in more folders than the default one.**  ``/contacts`` is the default folder only;
   `folders` walks ``/contactfolders`` and every ``childfolders`` below it, and a backup covers them all.
9. **Contact groups are in the folder but not in ``/contacts``.**  A folder's ``$count`` includes its
   distribution lists; the REST contact list silently leaves them out, so the two never agree.  `groups`
   reads them through Outlook's own ``service.svc`` (FindItem, then GetItem), members and all, and a
   backup checks that contacts plus groups add up to the server's count.
10. **The REST contact is one view; Outlook's own is another.**  ``service.svc`` GetItem with
    ``AllProperties`` returns the contact in Exchange's schema: phone numbers keyed by kind
    (``BusinessFax``, ``Pager``...), addresses keyed Home/Business/Other, follow-up flags, sensitivity.
    REST ids are not accepted there, but the ChangeKey is shared, so `ews_items` joins the two by it.
    A backup holds both views, and the deduplication check needs both to agree that nothing is lost.
11. **DELETE takes no precondition.**  An ``If-Match`` with a stale ChangeKey is ignored and the contact
    deleted anyway, so "unchanged since I checked" can only be narrowed, never guaranteed, by re-reading
    just before deleting.  What makes deleting tolerable is that it is soft: a deleted contact moves to
    Deleted Items whole, every MAPI property included, until that folder is emptied.
"""
from __future__ import annotations

import base64
import datetime
import hashlib
import json
import os
import re
import time
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


#: MAPI properties an Outlook contact can carry that the REST contact does not expose: name -> PropertyId.
_ADDRESS_SET = "{00062004-0000-0000-C000-000000000046}"
EXTENDED_PROPERTIES = {
    "other_phone": "String 0x3A1F", "primary_phone": "String 0x3A1A", "company_main_phone": "String 0x3A57",
    "business_fax": "String 0x3A24", "home_fax": "String 0x3A25", "other_fax": "String 0x3A23",
    "pager": "String 0x3A21", "car_phone": "String 0x3A1E", "callback_phone": "String 0x3A02",
    "radio_phone": "String 0x3A1D", "telex": "String 0x3A2C", "isdn": "String 0x3A2D",
    "tty_tdd_phone": "String 0x3A4B", "assistant_phone": "String 0x3A2E",
    "personal_home_page": "String 0x3A50", "hobbies": "String 0x3A43", "language": "String 0x3A0C",
    "location": "String 0x3A0D", "account": "String 0x3A00", "customer_id": "String 0x3A4A",
    "government_id": "String 0x3A07", "organizational_id": "String 0x3A10",
    "referred_by": "String 0x3A47", "computer_network_name": "String 0x3A49", "ftp_site": "String 0x3A4C",
    "anniversary": "SystemTime 0x3A41", "gender": "Short 0x3A4D",
    "user_field_1": f"String {_ADDRESS_SET} Id 0x804F", "user_field_2": f"String {_ADDRESS_SET} Id 0x8050",
    "user_field_3": f"String {_ADDRESS_SET} Id 0x8051", "user_field_4": f"String {_ADDRESS_SET} Id 0x8052",
    "web_page": f"String {_ADDRESS_SET} Id 0x802B", "im_address": f"String {_ADDRESS_SET} Id 0x8062",
}

#: Which of those are phone numbers, so they are matched by digits like the REST slots.
EXTENDED_PHONES = {"other_phone", "primary_phone", "company_main_phone", "business_fax", "home_fax",
                   "other_fax", "pager", "car_phone", "callback_phone", "radio_phone", "telex", "isdn",
                   "tty_tdd_phone", "assistant_phone"}

#: Server-assigned properties a create must not send back.
_READ_ONLY = {"Id", "ChangeKey", "CreatedDateTime", "LastModifiedDateTime", "ParentFolderId"}

#: Items per service.svc GetItem call.
_GETITEM_BATCH = 100

#: Contacts per in-page call when fetching photos; each call runs its fetches one after another.
_PHOTO_BATCH = 50

#: Fetch each id's photo as base64, or the HTTP status when there is none (404).
_PHOTOS_JS = """(async()=>{const out={};for(const id of %s){try{
  const r=await fetch(%s+'/'+encodeURIComponent(id)+'/photo/$value',{headers:{authorization:'Bearer '+%s}});
  if(r.status!==200){out[id]=r.status;continue}
  const b=new Uint8Array(await r.arrayBuffer());let s='';
  for(let i=0;i<b.length;i+=8192) s+=String.fromCharCode.apply(null,b.subarray(i,i+8192));
  out[id]={type:r.headers.get('content-type'),b64:btoa(s)}}catch(e){out[id]=String(e)}}return out})()"""


#: POST one request to Outlook's own JSON service; returns its response messages, or the error body.
_OWA_JS = """(async()=>{const r=await fetch('/owa/service.svc?action='+%s+'&app=Mail',{method:'POST',
  headers:{authorization:'Bearer '+%s,action:%s,'content-type':'application/json; charset=utf-8'},body:%s});
  let d=null;try{d=await r.json()}catch(e){}
  const m=d&&d.Body&&d.Body.ResponseMessages&&d.Body.ResponseMessages.Items;
  return {status:r.status, messages:m||null, error:m?null:JSON.stringify(d).slice(0,400)}})()"""


def _ext_expand() -> str:
    flt = " or ".join(f"PropertyId eq '{pid}'" for pid in EXTENDED_PROPERTIES.values())
    return "$expand=SingleValueExtendedProperties($filter=" + urllib.parse.quote(flt) + ")"


def extended(raw: dict) -> dict:
    """A raw record's non-empty extended properties, by `EXTENDED_PROPERTIES` name."""
    by_id = {pid.lower(): name for name, pid in EXTENDED_PROPERTIES.items()}
    out = {}
    for p in raw.get("SingleValueExtendedProperties") or []:
        name = by_id.get((p.get("PropertyId") or "").lower(), p.get("PropertyId"))
        val = p.get("Value")
        if val not in (None, "") and not (name == "gender" and str(val) == "0"):
            out[name] = val
    return out


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
        self._token_at = 0.0
        self.root = self.base.rsplit("/", 1)[0]

    def token(self) -> str:
        """The page's token, re-read every ten minutes so a long backup or apply outlives one token."""
        at = getattr(self, "_token_at", 0.0)
        if self._token is None or (at and time.time() - at > 600):
            self._token, self._token_at = access_token(self.s), time.time()
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
        """Delete one contact.  It goes to Deleted Items whole; the server takes no ChangeKey precondition."""
        rest_send(self.s, "DELETE", f"{self.base}/{urllib.parse.quote(contact_id, safe='')}",
                  token=self.token())

    # ------------------------------------------------------------------ whole records

    def _folder_base(self, folder_id: str | None) -> str:
        return self.base if not folder_id else \
            f"{self.root}/contactfolders/{urllib.parse.quote(folder_id, safe='')}/contacts"

    def folders(self) -> list[dict]:
        """Every contact folder: the default one (id ``None``) first, then each folder below it, depth first."""
        out = [{"id": None, "name": "Contacts", "parent": None}]

        def walk(path: str, parent: str | None, prefix: str) -> None:
            while path:
                page = rest_get(self.s, path, self.token())
                for f in page.get("value") or []:
                    name = prefix + f.get("DisplayName", "")
                    out.append({"id": f["Id"], "name": name, "parent": parent})
                    walk(f"{self.root}/contactfolders/{urllib.parse.quote(f['Id'], safe='')}/childfolders"
                         f"?$top=100", f["Id"], name + "/")
                nxt = page.get("@odata.nextLink")
                path = ("/" + nxt.split("/", 3)[3]) if nxt else None

        walk(f"{self.root}/contactfolders?$top=100", None, "")
        return out

    def count(self, folder_id: str | None = None) -> int:
        """How many contacts the server says `folder_id` holds, independently of paging through them."""
        page = rest_get(self.s, f"{self._folder_base(folder_id)}?$top=1&$select=Id&$count=true", self.token())
        return int(page["@odata.count"])

    def list_raw(self, folder_id: str | None = None) -> list[dict]:
        """Every contact in one folder as the server's whole record, with `EXTENDED_PROPERTIES` expanded."""
        out: list[dict] = []
        path: str | None = f"{self._folder_base(folder_id)}?$top=500&{_ext_expand()}"
        while path:
            page = rest_get(self.s, path, self.token())
            out += [{k: v for k, v in r.items() if not k.startswith("@odata")} for r in page.get("value") or []]
            nxt = page.get("@odata.nextLink")
            path = ("/" + nxt.split("/", 3)[3]) if nxt else None
        return out

    def _owa(self, action: str, body: dict) -> list[dict]:
        """One service.svc call; its response messages (one per item asked for), all checked NoError."""
        req = {"__type": f"{action}JsonRequest:#Exchange",
               "Header": {"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": "V2018_01_08"},
               "Body": {"__type": f"{action}Request:#Exchange", **body}}
        out = self.s.js(_OWA_JS % (json.dumps(action), json.dumps(self.token()), json.dumps(action),
                                   json.dumps(json.dumps(req))), await_promise=True) or {}
        msgs = out.get("messages") or []
        bad = [m.get("ResponseCode") for m in msgs if m.get("ResponseCode") != "NoError"]
        if out.get("status") != 200 or not msgs or bad:
            raise OutlookError(f"service.svc {action} returned {out.get('status')} {bad} "
                               f"{out.get('error') or ''}"[:400])
        return msgs

    def _find_items(self, folder_id: str | None) -> list[dict]:
        """Every item in one folder as ``{"ItemId": {"Id", "ChangeKey"}, "ItemClass"}``, through service.svc."""
        parent = ({"__type": "DistinguishedFolderId:#Exchange", "Id": "contacts"} if not folder_id else
                  {"__type": "FolderId:#Exchange", "Id": folder_id})
        shape = {"__type": "ItemResponseShape:#Exchange", "BaseShape": "IdOnly",
                 "AdditionalProperties": [{"__type": "PropertyUri:#Exchange", "FieldURI": "ItemClass"}]}
        out, offset = [], 0
        while True:
            root = self._owa("FindItem", {"ItemShape": shape, "ParentFolderIds": [parent], "Traversal": "Shallow",
                                          "Paging": {"__type": "IndexedPageView:#Exchange", "BasePoint": "Beginning",
                                                     "Offset": offset, "MaxEntriesReturned": 1000}})[0]["RootFolder"]
            items = root.get("Items") or []
            out += items
            offset += len(items)
            if root.get("IncludesLastItemInRange", True) or not items:
                break
        return out

    def get_items(self, ews_ids: list[str]) -> list[dict]:
        """Outlook's whole items (GetItem, ``AllProperties``) for service.svc ids, in order."""
        out = []
        for i in range(0, len(ews_ids), _GETITEM_BATCH):
            got = self._owa("GetItem", {"ItemShape": {"__type": "ItemResponseShape:#Exchange",
                                                      "BaseShape": "AllProperties"},
                                        "ItemIds": [{"__type": "ItemId:#Exchange", "Id": x}
                                                    for x in ews_ids[i:i + _GETITEM_BATCH]]})
            out += [item for m in got for item in m.get("Items") or []]
        return out

    def groups(self, folder_id: str | None = None) -> list[dict]:
        """The contact groups (distribution lists) in one folder, each as Outlook's whole item with Members."""
        return self.get_items([i["ItemId"]["Id"] for i in self._find_items(folder_id)
                               if (i.get("ItemClass") or "").startswith("IPM.DistList")])

    def ews_items(self, folder_id: str | None = None, change_keys: set | None = None) -> dict[str, dict]:
        """Outlook's own view of the contacts in one folder, keyed by ChangeKey - the one key REST shares.

        With `change_keys`, only those contacts are fetched.
        """
        ids = [i["ItemId"]["Id"] for i in self._find_items(folder_id)
               if (i.get("ItemClass") or "").startswith("IPM.Contact")
               and (change_keys is None or i["ItemId"]["ChangeKey"] in change_keys)]
        return {item["ItemId"]["ChangeKey"]: item for item in self.get_items(ids)}

    def change_keys(self, folder_id: str | None = None) -> dict[str, str]:
        """Each contact's id -> ChangeKey in one folder: cheap, for proving nothing changed meanwhile."""
        out: dict[str, str] = {}
        path: str | None = f"{self._folder_base(folder_id)}?$top=500&$select=Id,ChangeKey"
        while path:
            page = rest_get(self.s, path, self.token())
            out.update({r["Id"]: r["ChangeKey"] for r in page.get("value") or []})
            nxt = page.get("@odata.nextLink")
            path = ("/" + nxt.split("/", 3)[3]) if nxt else None
        return out

    def get_raw(self, contact_id: str) -> dict | None:
        """One whole record, or ``None`` when the server no longer has it."""
        try:
            r = rest_get(self.s, f"{self.base}/{urllib.parse.quote(contact_id, safe='')}?{_ext_expand()}",
                         self.token())
        except OutlookError as e:
            if " 404 " in str(e) or "ErrorItemNotFound" in str(e):
                return None
            raise
        return {k: v for k, v in r.items() if not k.startswith("@odata")}

    def photos(self, contact_ids: list[str]) -> dict[str, dict | None]:
        """Each contact's photo as ``{"type", "b64", "sha256"}``, or ``None`` for a contact without one."""
        out: dict[str, dict | None] = {}
        for i in range(0, len(contact_ids), _PHOTO_BATCH):
            batch = contact_ids[i:i + _PHOTO_BATCH]
            got = self.s.js(_PHOTOS_JS % (json.dumps(batch), json.dumps(self.base), json.dumps(self.token())),
                            await_promise=True) or {}
            for cid in batch:
                v = got.get(cid)
                if isinstance(v, dict):
                    v["sha256"] = hashlib.sha256(base64.b64decode(v["b64"])).hexdigest()
                    out[cid] = v
                elif v == 404:
                    out[cid] = None
                else:
                    raise OutlookError(f"photo fetch for one contact failed: {v!r}")
        return out

    def backup(self, directory: str) -> dict:
        """Write every contact in every folder, with photos, to a timestamped JSON file, and prove it whole.

        The proof: each folder's contacts plus its contact groups equal the server's own ``$count`` taken
        before and after reading it; every contact's two views (REST and service.svc) are joined by
        ChangeKey; once everything is read, every folder's contact and group ids and ChangeKeys are read again
        and must be exactly those saved, so nothing was added, removed or edited meanwhile; every id is distinct; and
        the file read back from disk parses to the same records.
        Returns a summary with the file's path and SHA-256.
        """
        stamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        folders, records, groups = self.folders(), [], []
        for f in folders:
            before = self.count(f["id"])
            raw = self.list_raw(f["id"])
            grp = self.groups(f["id"])
            after = self.count(f["id"])
            if not before == after == len(raw) + len(grp):
                raise OutlookError(f"folder {f['name']!r}: server counted {before} then {after}, read "
                                   f"{len(raw)} contacts and {len(grp)} groups. The book changed or "
                                   f"paging lost records; not saved.")
            f["count"], f["groups"] = len(raw), len(grp)
            groups += [{"folder": f["id"], "item": g} for g in grp]
            pics = self.photos([r["Id"] for r in raw])
            ews = self.ews_items(f["id"])
            unjoined = [r["Id"] for r in raw if r["ChangeKey"] not in ews]
            if unjoined or len(ews) != len(raw):
                raise OutlookError(f"folder {f['name']!r}: {len(unjoined)} contacts have no service.svc view "
                                   f"({len(ews)} views for {len(raw)} contacts); not saved.")
            records += [{"folder": f["id"], "record": r, "photo": pics[r["Id"]], "ews": ews[r["ChangeKey"]]}
                        for r in raw]
        for f in folders:
            saved = {e["record"]["Id"]: e["record"]["ChangeKey"] for e in records if e["folder"] == f["id"]}
            saved_groups = {g["item"]["ItemId"]["Id"]: g["item"]["ItemId"]["ChangeKey"]
                            for g in groups if g["folder"] == f["id"]}
            now_groups = {i["ItemId"]["Id"]: i["ItemId"]["ChangeKey"] for i in self._find_items(f["id"])
                          if (i.get("ItemClass") or "").startswith("IPM.DistList")}
            if self.change_keys(f["id"]) != saved or now_groups != saved_groups:
                raise OutlookError(f"folder {f['name']!r} changed while it was being read; not saved. "
                                   f"Run the backup again.")
        ids = [r["record"]["Id"] for r in records]
        if len(set(ids)) != len(ids):
            raise OutlookError("the same contact id was read twice; not saved")
        doc = {"format": 1, "mailbox": self.mailbox, "taken_at": stamp, "folders": folders,
               "extended_properties": EXTENDED_PROPERTIES, "contacts": records, "groups": groups}
        os.makedirs(directory, exist_ok=True)
        path = os.path.join(directory, f"contacts-backup-{stamp}.json")
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False, indent=1)
        with open(path, "rb") as fh:
            data = fh.read()
        if json.loads(data) != doc:
            raise OutlookError(f"{path} does not read back as what was written")
        return {"path": path, "sha256": hashlib.sha256(data).hexdigest(), "contacts": len(records),
                "groups": len(groups), "photos": sum(1 for r in records if r["photo"]),
                "folders": {f["name"]: {"contacts": f["count"], "groups": f["groups"]} for f in folders}}

    def restore(self, entry: dict, folder_id: str | None = None) -> dict:
        """Recreate one backed-up contact (an item of a backup's ``contacts``) and prove nothing was lost.

        It goes back into its own folder unless `folder_id` is given.  The new contact gets a new id; it
        is read back and must hold everything the backed-up one held in both views, or it is deleted again
        and refused - as it is if anything fails after it was created.
        """
        from .dedupe import coverage_gaps   # dedupe imports this module
        raw = entry["record"]
        payload = {k: v for k, v in raw.items() if k not in _READ_ONLY and k != "SingleValueExtendedProperties"
                   and v not in (None, "", [], {})}
        ext = [p for p in raw.get("SingleValueExtendedProperties") or [] if p.get("Value") not in (None, "")]
        if ext:
            payload["SingleValueExtendedProperties"] = ext
        made = rest_send(self.s, "POST", self._folder_base(folder_id or entry.get("folder")), payload,
                         self.token())
        if not made.get("Id"):
            raise OutlookError(f"contact POST returned no Id: {json.dumps(made)[:300]}")
        try:
            if entry.get("photo"):
                self.set_photo(made["Id"], base64.b64decode(entry["photo"]["b64"]), entry["photo"]["type"])
            record = self.get_raw(made["Id"])
            back = {"record": record, "photo": self.photos([made["Id"]])[made["Id"]],
                    "ews": self.ews_items(folder_id or entry.get("folder"), {record["ChangeKey"]}).get(
                        record["ChangeKey"])}
            gaps = coverage_gaps(entry, back)
        except Exception:
            self.delete(made["Id"])
            raise
        if gaps:
            self.delete(made["Id"])
            raise OutlookError("the restored contact lacked some of the backup, so it was deleted again: "
                               + "; ".join(gaps))
        return back["record"]
