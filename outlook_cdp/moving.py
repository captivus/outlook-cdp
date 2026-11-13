"""Moving chosen messages to another folder - by default Deleted Items - over REST, and proving it.

    mv = Mover(Session("my-instance", target="2F214D46"))
    plan = mv.find(from_addresses=["noisy@example.com"], received_after="2026-11-12")
    for m in plan:
        print(m.received, m.folder, m.sender_address, m.subject)
    results = mv.move(plan)                     # each one read back from its new folder
    assert all(r.ok for r in results)

What this does and deliberately does not do:

1. **A move, never a delete.**  ``POST messages/{id}/move`` with a destination folder.  The default
   destination is ``deleteditems``, which is where Outlook's own Delete key puts a message, so it can
   be restored from there.  Nothing here calls ``DELETE``.
2. **Nothing else changes.**  A REST move keeps ``IsRead`` as it was, so nothing is marked read.  Moving
   a message out of Junk Email is not "Not junk": that is a separate report that also adds the sender
   to safe senders, and nothing here makes it.  No rule and no safe-senders entry is touched.
3. **A move gives the message a new id.**  Exchange ids encode the folder, so the id Outlook knew before
   the move no longer finds anything.  `move` returns the new one, and verification reads *that* back
   and checks its folder, rather than trusting the 201.
4. **Finding is mailbox-wide.**  ``/messages`` with no folder covers every folder, Junk Email and
   Deleted Items included, so a message that went to Junk is found without naming Junk.  Whatever is
   already in the destination is left out of the plan.
5. **The sender filter runs on the server; every other filter runs here.**  Combining a ``$filter`` on
   ``From`` with ``$orderby`` or with a ``ReceivedDateTime`` clause risks ``InefficientFilter``, so
   only the sender goes to the server and dates and subjects are applied to what comes back.
"""
from __future__ import annotations

import json
import urllib.parse
from dataclasses import dataclass

from .reading import MESSAGE_FIELDS, Message, Reader
from .rules import rest_send
from .session import OutlookError, Session

#: Refuse a move bigger than this unless the caller raises it.  A typo in a sender address finds
#: nothing; a too-broad one finds a mailbox, and this is the backstop for that.
DEFAULT_MAX = 50


@dataclass
class MoveResult:
    """What happened to one message."""

    old_id: str
    subject: str
    ok: bool
    new_id: str | None = None
    folder: str | None = None       #: where the read-back found it
    error: str = ""


def _in_range(received: str, after: str | None, before: str | None) -> bool:
    """`after` inclusive, `before` exclusive, both compared as UTC ISO strings.

    A bare date works as a bound because ``2026-11-12T08:00:00Z`` sorts after ``2026-11-12`` and before
    ``2026-11-13``: ``--received-after 2026-11-12 --received-before 2026-11-13`` is that UTC day.
    """
    return (not after or received >= after) and (not before or received < before)


def matches(m: Message, *, from_addresses=(), subject_contains=(), received_after=None,
            received_before=None) -> bool:
    """The local half of `Mover.find`'s filtering.  Values within a filter are OR, filters are AND."""
    addrs = {a.lower() for a in from_addresses}
    if addrs and (m.sender_address or "").lower() not in addrs:
        return False
    subj = (m.subject or "").lower()
    if subject_contains and not any(s.lower() in subj for s in subject_contains):
        return False
    return _in_range(m.received or "", received_after, received_before)


class Mover:
    """Moves messages in the mailbox a tab is showing (or `mailbox`).  See the module docstring."""

    def __init__(self, session: Session, mailbox: str | None = None):
        self.s = session
        self.r = Reader(session, wake=False, mailbox=mailbox)
        self._dest_ids: dict[str, str] = {}

    @property
    def base(self) -> str:
        return self.r.base

    def folder_id(self, folder: str) -> str:
        """The id of a well-known folder name (``deleteditems``, ``junkemail``...) or of an id."""
        if folder not in self._dest_ids:
            got = self.r._get(f"{self.base}/mailfolders/{urllib.parse.quote(folder, safe='')}?$select=Id")
            if not got.get("Id"):
                raise OutlookError(f"no folder {folder!r} in {self.base}")
            self._dest_ids[folder] = got["Id"]
        return self._dest_ids[folder]

    # ------------------------------------------------------------------ finding

    def find(self, *, from_addresses=(), ids=(), folders=(), subject_contains=(),
             received_after: str | None = None, received_before: str | None = None,
             destination: str | None = "deleteditems", limit: int = 1000) -> list[Message]:
        """The messages a move would take, newest first.  Changes nothing, marks nothing read.

        At least one of `from_addresses` or `ids` is required, so no call can sweep a whole folder.
        `folders` (well-known names or ids) narrows the search; without it, the whole mailbox is
        searched.  Messages already in `destination` are never included; pass ``destination=None``
        to exclude nothing, which makes this a plain sender search.
        """
        if not (from_addresses or ids):
            raise OutlookError("refusing to look for messages to move without --from or --ids")
        dest = self.folder_id(destination) if destination else None
        sel = MESSAGE_FIELDS
        found: dict[str, dict] = {}
        if ids:
            for mid in ids:
                found[mid] = self.r._get(f"{self.base}/messages/{urllib.parse.quote(mid, safe='')}"
                                         f"?$select={sel}")
        else:
            ors = " or ".join(f"From/EmailAddress/Address eq '{a.replace(chr(39), chr(39) * 2)}'"
                              for a in from_addresses)
            scopes = ([f"mailfolders/{urllib.parse.quote(f, safe='')}/messages" for f in folders]
                      or ["messages"])
            for scope in scopes:
                path = (f"{self.base}/{scope}?$top=250&$select={sel}"
                        f"&$filter={urllib.parse.quote(ors)}")
                while path and len(found) < limit:
                    page = self.r._get(path)
                    for raw in page.get("value") or []:
                        found[raw["Id"]] = raw
                    nxt = page.get("@odata.nextLink")
                    path = ("/" + nxt.split("/", 3)[3]) if nxt else None
        # With ids, `folders` is a check on them rather than a search scope.
        allowed = {self.folder_id(f) for f in folders} if (folders and ids) else None
        out = []
        for raw in found.values():
            if dest and raw.get("ParentFolderId") == dest:
                continue
            if allowed is not None and raw.get("ParentFolderId") not in allowed:
                continue
            m = self.r._message(raw)
            if matches(m, from_addresses=from_addresses, subject_contains=subject_contains,
                       received_after=received_after, received_before=received_before):
                out.append(m)
        out.sort(key=lambda m: m.received or "", reverse=True)
        return out

    # ------------------------------------------------------------------ moving

    def move(self, messages, destination: str = "deleteditems", *,
             max_messages: int = DEFAULT_MAX) -> list[MoveResult]:
        """Move each message (a `Message` or an id) to `destination`, then read each one back there.

        Every message gets its own result: one refusal does not stop the rest, and nothing counts as
        moved until the read-back has found it in `destination`.
        """
        items = [(m, "") if isinstance(m, str) else (m.id, m.subject) for m in messages]
        if len(items) > max_messages:
            raise OutlookError(f"refusing to move {len(items)} messages; the limit is {max_messages}. "
                               f"Narrow the filters, or raise the limit if this many is intended.")
        dest = self.folder_id(destination)
        dest_name = self.r._folder_names().get(dest, destination)
        results = []
        for mid, subject in items:
            try:
                moved = rest_send(self.s, "POST",
                                  f"{self.base}/messages/{urllib.parse.quote(mid, safe='')}/move",
                                  {"DestinationId": destination}, self.r.token())
                new_id = moved.get("Id")
                if not new_id:
                    raise OutlookError("the move answered without the message's new id")
                back = self.r._get(f"{self.base}/messages/{urllib.parse.quote(new_id, safe='')}"
                                   f"?$select=Id,ParentFolderId")
                where = back.get("ParentFolderId")
                ok = where == dest
                results.append(MoveResult(
                    old_id=mid, subject=subject, ok=ok, new_id=new_id,
                    folder=dest_name if ok else self.r._folder_names().get(where, where),
                    error="" if ok else f"read back from {where!r}, not {destination}"))
            except OutlookError as e:
                results.append(MoveResult(old_id=mid, subject=subject, ok=False, error=str(e)))
        return results


def describe(m: Message) -> dict:
    """One plan line, in the order a person checks it."""
    return {"received": m.received, "folder": m.folder, "from": m.sender_address,
            "subject": m.subject, "unread": m.unread, "id": m.id}


def dumps(rows) -> str:
    return json.dumps(rows, indent=1, ensure_ascii=False)
