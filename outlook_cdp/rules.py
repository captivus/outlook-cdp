"""Inbox rules: list, create, verify, delete - over the mailbox's own REST API, without the settings UI.

    from outlook_cdp import Session, Rules, Rule

    rules = Rules(Session("my-chrome-agent-instance"))
    rule = Rule("Archive build notices", from_addresses=["ci@example.com"],
                subject_contains=["Build succeeded"], move_to="archive")

    for row in rules.preview(rule):          # what it would have caught, newest first
        print(row.received, row.subject)
    created = rules.create(rule)             # refuses a duplicate name, re-reads what the server stored

Three things about this that are not obvious:

1. **The token already allows it.**  The access token Outlook caches for its own REST calls carries
   ``MailboxSettings.ReadWrite``, which is what ``/mailfolders/inbox/messagerules`` needs, so a rule can be
   created from inside the page with no consent prompt and no settings dialog to click through.
2. **A rule is not proof of what it will do.**  The server can accept a POST and store fewer conditions
   than were sent - an unrecognised condition key is dropped, not refused - and a rule with no
   conditions matches *every* message.  So `create` reads the rule back and compares it with what was
   asked for, and deletes it again if they differ.  A rule that would archive the whole inbox is worse
   than no rule.
3. **A rule only acts on mail that arrives after it exists.**  Outlook's "run this rule now" is a UI
   feature with no REST equivalent.  `apply` does the same thing explicitly: it moves the messages
   `preview` returns, and only those.

`move_to` takes a well-known folder name (``archive``, ``deleteditems``, ``junkemail``, ...) or a folder
id; it is resolved to an id before the POST, because the server stores and returns ids.

Like the rest of the REST path, every call is scoped to ``/api/v2.0/users/<mailbox>/``.  The mailbox
comes from the tab URL when it carries one; a tab sitting on ``/mail/`` does not, and then the caller
must say whose rules these are - see `Rules`.
"""
from __future__ import annotations

import json
import urllib.parse
from dataclasses import dataclass, field

from .attachments import auth_header, rest_get, resolve_mailbox
from .reading import Reader, Row
from .session import OutlookError, Session

#: Conditions this module knows how to send, read back and evaluate locally.  Anything else is refused
#: rather than passed through, because the server would silently drop a key it does not recognise.
_CONDITION_KEYS = {
    "from_addresses": "FromAddresses",
    "sender_contains": "SenderContains",
    "subject_contains": "SubjectContains",
    "body_contains": "BodyContains",
    "subject_or_body_contains": "SubjectOrBodyContains",
}

#: Same-origin request with the borrowed bearer token, for the verbs `rest_get` deliberately lacks.
_SEND_JS = """(async()=>{try{
  const r=await fetch(%s,{method:%s,headers:{authorization:%s,accept:'application/json',
    'content-type':'application/json'},body:%s});
  let body=null; try{body=await r.json()}catch(e){}
  return {status:r.status, body:body}}catch(e){return {status:0, error:String(e)}}})()"""


def rest_send(session: Session, method: str, path: str, payload: dict | None = None,
              token: str | None = None) -> dict:
    """POST / PATCH / DELETE a same-origin Outlook REST path; return the decoded JSON body, if any.

    Raises on anything but 2xx, with the server's error code, so a refused write never reads as done.
    """
    token = token or auth_header(session)
    body = json.dumps(json.dumps(payload)) if payload is not None else "undefined"
    out = session.js(_SEND_JS % (json.dumps(path), json.dumps(method.upper()), json.dumps(token), body),
                     await_promise=True) or {}
    status = out.get("status") or 0
    if not 200 <= status < 300:
        got = out.get("body") or {}
        err = (got.get("error") or {}) if isinstance(got, dict) else {}
        raise OutlookError(
            f"REST {method.upper()} {path.split('?')[0]} returned {status} "
            f"{err.get('code', '')} {err.get('message', out.get('error', ''))}"[:400])
    return out.get("body") or {}


@dataclass
class Rule:
    """One inbox rule, in the shape this module can send, read back and evaluate.

    Conditions of different kinds are ANDed; the values inside one kind are ORed - which is how
    Outlook evaluates them.  `*_contains` matching is a case-insensitive substring; `from_addresses`
    is an exact address.
    """
    name: str
    from_addresses: list[str] = field(default_factory=list)
    sender_contains: list[str] = field(default_factory=list)
    subject_contains: list[str] = field(default_factory=list)
    body_contains: list[str] = field(default_factory=list)
    subject_or_body_contains: list[str] = field(default_factory=list)
    move_to: str | None = None          # well-known name or folder id
    mark_read: bool = False
    stop_processing: bool = True
    enabled: bool = True
    sequence: int | None = None
    id: str | None = None
    unknown_conditions: list[str] = field(default_factory=list)   # read back, not understood here
    unknown_actions: list[str] = field(default_factory=list)

    def conditions(self) -> dict[str, list[str]]:
        return {k: getattr(self, k) for k in _CONDITION_KEYS if getattr(self, k)}

    def to_rest(self, folder_id: str | None) -> dict:
        if not self.conditions():
            raise OutlookError(
                f"rule {self.name!r} has no conditions, so it would match every message. Refusing.")
        if self.move_to and not folder_id:
            raise OutlookError(f"rule {self.name!r}: move_to {self.move_to!r} was not resolved to a folder id")
        conds: dict = {}
        for key, rest_key in _CONDITION_KEYS.items():
            vals = getattr(self, key)
            if not vals:
                continue
            if key == "from_addresses":
                conds[rest_key] = [{"EmailAddress": {"Address": a}} for a in vals]
            else:
                conds[rest_key] = list(vals)
        actions: dict = {"StopProcessingRules": bool(self.stop_processing)}
        if folder_id:
            actions["MoveToFolder"] = folder_id
        if self.mark_read:
            actions["MarkAsRead"] = True
        out = {"DisplayName": self.name, "IsEnabled": bool(self.enabled),
               "Conditions": conds, "Actions": actions}
        if self.sequence is not None:
            out["Sequence"] = int(self.sequence)
        return out

    @classmethod
    def from_rest(cls, raw: dict) -> "Rule":
        conds = raw.get("Conditions") or {}
        acts = raw.get("Actions") or {}
        known = set(_CONDITION_KEYS.values())
        kw: dict = {}
        for key, rest_key in _CONDITION_KEYS.items():
            vals = conds.get(rest_key) or []
            if key == "from_addresses":
                vals = [((v.get("EmailAddress") or {}).get("Address") or "") for v in vals]
            kw[key] = [v for v in vals if v]
        return cls(
            name=raw.get("DisplayName") or "", **kw,
            move_to=acts.get("MoveToFolder"), mark_read=bool(acts.get("MarkAsRead")),
            stop_processing=bool(acts.get("StopProcessingRules")),
            enabled=bool(raw.get("IsEnabled")), sequence=raw.get("Sequence"), id=raw.get("Id"),
            # Outlook sends empty or null for conditions that are not set; only a real value counts.
            unknown_conditions=sorted(k for k, v in conds.items() if k not in known and v),
            unknown_actions=sorted(k for k, v in acts.items()
                                   if k not in ("MoveToFolder", "MarkAsRead", "StopProcessingRules")
                                   and v))

    def matches(self, subject: str, sender_address: str, sender_name: str = "", body: str = "") -> bool:
        """Evaluate the conditions against one message, as Outlook would."""
        if not self.conditions():
            return False
        low = lambda s: (s or "").lower()   # noqa: E731
        any_in = lambda needles, hay: any(low(n) in low(hay) for n in needles)   # noqa: E731
        if self.from_addresses and low(sender_address) not in {low(a) for a in self.from_addresses}:
            return False
        if self.sender_contains and not any_in(self.sender_contains, f"{sender_name} {sender_address}"):
            return False
        if self.subject_contains and not any_in(self.subject_contains, subject):
            return False
        if self.body_contains and not any_in(self.body_contains, body):
            return False
        if self.subject_or_body_contains and not any_in(self.subject_or_body_contains,
                                                        f"{subject}\n{body}"):
            return False
        return True


def differences(asked: Rule, stored: Rule, folder_id: str | None) -> list[str]:
    """What the server stored that differs from what was asked for.  Empty means it took the rule as sent."""
    probs = []
    for key in _CONDITION_KEYS:
        want = sorted(v.lower() for v in getattr(asked, key))
        got = sorted(v.lower() for v in getattr(stored, key))
        if want != got:
            probs.append(f"{key}: asked {want}, stored {got}")
    if stored.unknown_conditions:
        probs.append(f"stored conditions this module did not send: {stored.unknown_conditions}")
    if (folder_id or None) != (stored.move_to or None):
        probs.append(f"move_to: asked {folder_id}, stored {stored.move_to}")
    if asked.mark_read != stored.mark_read:
        probs.append(f"mark_read: asked {asked.mark_read}, stored {stored.mark_read}")
    if asked.enabled != stored.enabled:
        probs.append(f"enabled: asked {asked.enabled}, stored {stored.enabled}")
    return probs


class Rules:
    """The inbox rules of one mailbox.

    `mailbox` defaults to the one in the tab's URL.  A tab on plain ``/mail/`` has none there, and this
    refuses rather than guessing: on a shared mailbox, ``/me`` is a different person.  When the tab is
    the signed-in user's own mailbox, `signed_in_address()` gives the address to pass.
    """

    def __init__(self, session: Session, mailbox: str | None = None):
        self.s = session
        self.mailbox = resolve_mailbox(session, mailbox)
        self.base = "/api/v2.0/users/" + urllib.parse.quote(self.mailbox, safe="")
        self._token: str | None = None

    @staticmethod
    def signed_in_address(session: Session) -> str:
        """The signed-in user's primary address - which is not necessarily the mailbox on screen."""
        return (rest_get(session, "/api/v2.0/me").get("EmailAddress") or "").lower()

    # ------------------------------------------------------------------ plumbing

    def token(self) -> str:
        if self._token is None:
            self._token = auth_header(self.s)
        return self._token

    def _get(self, path: str) -> dict:
        return rest_get(self.s, path, self.token())

    def _send(self, method: str, path: str, payload: dict | None = None) -> dict:
        return rest_send(self.s, method, path, payload, self.token())

    def folder_id(self, folder: str) -> str:
        """Resolve a well-known name or id to the id the server stores."""
        got = self._get(f"{self.base}/mailfolders/{urllib.parse.quote(folder, safe='')}?$select=Id")
        if not got.get("Id"):
            raise OutlookError(f"no folder {folder!r} in {self.mailbox}")
        return got["Id"]

    def folder_name(self, folder_id: str) -> str:
        got = self._get(f"{self.base}/mailfolders/{urllib.parse.quote(folder_id, safe='')}"
                        f"?$select=DisplayName")
        return got.get("DisplayName") or folder_id

    # ------------------------------------------------------------------ rules

    @property
    def _rules_path(self) -> str:
        return f"{self.base}/mailfolders/inbox/messagerules"

    def list(self) -> list[Rule]:
        """Every inbox rule, in evaluation order."""
        raws = self._get(self._rules_path).get("value") or []
        return sorted((Rule.from_rest(r) for r in raws), key=lambda r: r.sequence or 0)

    def get(self, rule_id: str) -> Rule:
        return Rule.from_rest(self._get(f"{self._rules_path}/{urllib.parse.quote(rule_id, safe='')}"))

    def find(self, name: str) -> list[Rule]:
        return [r for r in self.list() if r.name == name]

    def create(self, rule: Rule) -> Rule:
        """Create `rule`, read it back, and refuse - deleting it again - if the server changed it.

        Refuses a name that is already taken, so re-running a script does not stack up copies.  With no
        `sequence` the rule goes last: the server requires one and answers an omitted value with
        ``InvalidValue ... Field: 'Sequence', Value: '0'``.
        """
        existing = self.list()
        if any(r.name == rule.name for r in existing):
            raise OutlookError(f"a rule named {rule.name!r} already exists; delete it or pick a new name")
        if rule.sequence is None:
            rule.sequence = max((r.sequence or 0 for r in existing), default=0) + 1
        folder_id = self.folder_id(rule.move_to) if rule.move_to else None
        made = self._send("POST", self._rules_path, rule.to_rest(folder_id))
        if not made.get("Id"):
            raise OutlookError(f"rule POST returned no Id: {json.dumps(made)[:300]}")
        stored = self.get(made["Id"])
        probs = differences(rule, stored, folder_id)
        if probs:
            self.delete(made["Id"])
            raise OutlookError("the server stored a different rule than was asked for, so it was "
                               "deleted again: " + "; ".join(probs))
        return stored

    def delete(self, rule_id: str) -> None:
        self._send("DELETE", f"{self._rules_path}/{urllib.parse.quote(rule_id, safe='')}")

    def set_enabled(self, rule_id: str, enabled: bool) -> Rule:
        self._send("PATCH", f"{self._rules_path}/{urllib.parse.quote(rule_id, safe='')}",
                   {"IsEnabled": bool(enabled)})
        return self.get(rule_id)

    # ------------------------------------------------------------------ existing mail

    def preview(self, rule: Rule, folder: str = "inbox", *, limit: int = 5000) -> list[Row]:
        """Messages already in `folder` that `rule` matches, newest first.  Changes nothing.

        When the rule names sender addresses the server pre-filters on them; otherwise the folder is
        scanned, newest first, up to `limit` messages.  The conditions are then evaluated here, so the
        answer is the rule's semantics rather than ``$search``'s, which tokenises and stems.
        """
        reader = Reader(self.s, wake=False)
        reader._base, reader._token = self.base, self.token()
        sel = "Id,Subject,From,ReceivedDateTime,IsRead,ConversationId,ParentFolderId"
        if rule.body_contains or rule.subject_or_body_contains:
            sel += ",Body"
        path = (f"{self.base}/mailfolders/{urllib.parse.quote(folder, safe='')}/messages"
                f"?$top=250&$select={sel}")
        if rule.from_addresses:
            ors = " or ".join(f"From/EmailAddress/Address eq '{a.replace(chr(39), chr(39) * 2)}'"
                              for a in rule.from_addresses)
            path += "&$filter=" + urllib.parse.quote(ors)
        else:
            path += "&$orderby=ReceivedDateTime%20desc"
        hits: list[Row] = []
        seen = 0
        while path and seen < limit:
            page = self._get(path)
            for raw in page.get("value") or []:
                seen += 1
                m = reader._message(raw)
                body = (raw.get("Body") or {}).get("Content") or ""
                if rule.matches(m.subject, m.sender_address or "", m.sender or "", body):
                    hits.append(Row(senders=[m.sender] if m.sender else [], sender=m.sender,
                                    sender_address=m.sender_address, subject=m.subject,
                                    date_text=m.received, received=m.received, unread=m.unread,
                                    conversation_id=m.conversation_id, message_id=m.id,
                                    message_ids=[m.id or ""], folder=folder, source="rest"))
            nxt = page.get("@odata.nextLink")
            path = ("/" + nxt.split("/", 3)[3]) if nxt else None
        hits.sort(key=lambda r: r.received or "", reverse=True)
        for i, r in enumerate(hits):
            r.index = i
        return hits

    def apply(self, rule: Rule, folder: str = "inbox", *, limit: int = 5000) -> list[str]:
        """Do to existing mail what `rule` will do to new mail: move what `preview` matches.

        Outlook rules act only on arrivals.  This is the REST stand-in for "run this rule now", and it
        changes the mailbox, so it is never called by `create`.  Returns the moved messages' new ids.
        """
        if not rule.move_to:
            raise OutlookError(f"rule {rule.name!r} moves nothing; there is nothing to apply")
        dest = self.folder_id(rule.move_to)
        moved = []
        for row in self.preview(rule, folder, limit=limit):
            got = self._send("POST", f"{self.base}/messages/{urllib.parse.quote(row.message_id or '', safe='')}"
                                     f"/move", {"DestinationId": dest})
            moved.append(got.get("Id") or "")
            if rule.mark_read and got.get("Id"):
                self._send("PATCH", f"{self.base}/messages/{urllib.parse.quote(got['Id'], safe='')}",
                           {"IsRead": True})
        return moved
