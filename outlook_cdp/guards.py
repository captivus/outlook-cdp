"""Pre-send checks, as objects you compose rather than rules baked into the sender.

Every guard is a callable taking a `Draft` snapshot and returning a list of problem strings.  An empty
list means "no objection".  A sender runs all of them and refuses to send if any produced a problem, so
adding a rule never means editing the send path.

The rules that matter are project-specific — which mailbox may send, which addresses have bounced, where
the kill switch lives — so none of them is hardcoded here.  They are constructed by the caller.
"""
from __future__ import annotations

import csv
import os
from dataclasses import dataclass, field

from .text import normalise


@dataclass
class Draft:
    """What a guard is allowed to look at: a read-back snapshot of the open compose pane."""
    from_label: str | None = None
    mailbox_identity: str | None = None
    subject: str = ""
    subject_editable: bool = True
    body: str = ""
    chip_addresses: list[str] = field(default_factory=list)
    chip_display_names: list[str] = field(default_factory=list)
    chip_problems: list[str] = field(default_factory=list)
    well_text: str = ""
    intended_to: str = ""
    intended_cc: str = ""
    intended_bcc: str = ""
    intended_subject: str = ""
    intended_body: str = ""
    is_reply: bool = False


class Guard:
    name = "guard"

    def __call__(self, draft: Draft) -> list[str]:      # pragma: no cover - interface
        raise NotImplementedError


class SendLock(Guard):
    """A file whose existence forbids sending. The kill switch a human controls.

    A lock is better than a remembered rule: it is checked before anything is composed, it survives a
    restart, a new session and a different agent, and its contents explain themselves.
    """
    name = "send lock"

    def __init__(self, path: str):
        self.path = path

    def __call__(self, draft: Draft) -> list[str]:
        if os.path.exists(self.path):
            body = ""
            try:
                body = open(self.path, encoding="utf-8", errors="replace").read().strip()
            except OSError:
                pass
            return [f"sending is locked by {self.path}" + (f": {body}" if body else "")]
        return []


class FromAddress(Guard):
    """Require the compose pane's From control to read exactly this.

    Outlook will happily send from whichever identity the pane defaulted to. When a mailbox is shared,
    that default is not always the one you mean, and the recipient sees whoever actually sent it.
    """
    name = "from address"

    def __init__(self, expected_label: str):
        self.expected = expected_label

    def __call__(self, draft: Draft) -> list[str]:
        if (draft.from_label or "") != self.expected:
            return [f"From reads {draft.from_label!r}, not {self.expected!r}"]
        return []


class SenderIdentity(Guard):
    """Require positive proof of who is sending, from whichever source the session offers.

    Outlook renders a From chooser only when there is more than one identity available. A session
    signed in to a single mailbox has **no From control at all**, so `FromAddress` alone refuses every
    send in that case - which is exactly what happened the first time this library met one.

    This accepts either: a From control reading the expected label, or, when there is no From control,
    a tab URL scoped to the expected mailbox. It **fails closed**: if neither can be established it
    refuses, so "no From button" never silently becomes "any sender will do".
    """
    name = "sender identity"

    def __init__(self, expected_address: str, expected_from_label: str | None = None):
        self.address = expected_address.strip().lower()
        self.label = expected_from_label or f"From: {expected_address}"

    def __call__(self, draft: Draft) -> list[str]:
        if draft.from_label:
            if draft.from_label.strip() == self.label.strip():
                return []
            return [f"From control reads {draft.from_label!r}, not {self.label!r}"]
        if draft.mailbox_identity:
            if draft.mailbox_identity.strip().lower() == self.address:
                return []
            return [f"no From control, and the mailbox is {draft.mailbox_identity!r}, "
                    f"not {self.address!r}"]
        return ["neither a From control nor a mailbox identity could be established; "
                "refusing to send without knowing who it comes from"]


class RecipientAllowlist(Guard):
    """Refuse any recipient that is not on an explicit list. The inverse of Blocklist.

    A blocklist stops known-bad addresses; it cannot stop a typo reaching a stranger. When the point is
    that nothing may leave a closed circle - a live send test, a staging run - this is the only guard
    that guarantees it.

    It matches a chip's address when Outlook left one, and otherwise the display name Outlook resolved
    the address to, which the caller supplies via `aliases`. Without that an internal recipient has
    nothing in the page to compare against, so **every** chip must match something or the send is
    refused: an unidentifiable recipient is the case this exists for.
    """
    name = "recipient allowlist"

    def __init__(self, allowed, aliases: dict[str, str] | None = None):
        self.allowed = {str(a).strip().lower() for a in allowed}
        self.names = {str(v).strip().lower() for k, v in (aliases or {}).items()
                      if str(k).strip().lower() in self.allowed}

    def __call__(self, draft: Draft) -> list[str]:
        problems = []
        if not (draft.chip_addresses or draft.chip_display_names):
            return ["no recipient could be identified, so the allowlist cannot be enforced"]
        for addr in draft.chip_addresses:
            if addr.strip().lower() not in self.allowed:
                problems.append(f"recipient {addr} is not on the allowlist "
                                f"({', '.join(sorted(self.allowed))})")
        for name in draft.chip_display_names:
            if name.strip().lower() not in self.names:
                problems.append(
                    f"recipient shows as {name!r} with no address in the page, and it is not a known "
                    f"alias of an allowlisted address"
                    + (f" (known: {', '.join(sorted(self.names))})" if self.names else ""))
        return problems


class Blocklist(Guard):
    """Refuse any recipient on a CSV of addresses that have bounced or must never be written to.

    Fails closed: a missing file is not an empty list, it means the check cannot be performed, and a
    guard that silently passes when it cannot see is worse than no guard. The CSV needs an `address`
    column; `reason` and `first_bounced` are used in the message when present.
    """
    name = "blocklist"

    def __init__(self, path: str):
        self.path = path

    def _rows(self):
        if not os.path.exists(self.path):
            raise FileNotFoundError(self.path)
        with open(self.path, encoding="utf-8", errors="replace", newline="") as fh:
            return list(csv.DictReader(fh))

    def __call__(self, draft: Draft) -> list[str]:
        try:
            rows = self._rows()
        except FileNotFoundError:
            return [f"{self.path} is missing, so blocked recipients cannot be checked"]
        blocked = {(r.get("address") or "").strip().lower(): r for r in rows if r.get("address")}
        problems = []
        for addr in draft.chip_addresses:
            row = blocked.get(addr.strip().lower())
            if row:
                extra = ", ".join(x for x in (row.get("reason"), row.get("first_bounced")) if x)
                problems.append(f"recipient {addr} is blocked" + (f" ({extra})" if extra else ""))
        low = (draft.well_text or "").lower()
        for addr, row in blocked.items():
            if addr and addr in low and addr not in {a.lower() for a in draft.chip_addresses}:
                problems.append(f"blocked address {addr} appears in the recipient well text")
        return problems


class RecipientsReadable(Guard):
    """Refuse to send when any recipient could not be read. An unverifiable recipient is the risk."""
    name = "recipients readable"

    def __call__(self, draft: Draft) -> list[str]:
        problems = list(draft.chip_problems)
        if not (draft.chip_addresses or draft.chip_display_names) and not normalise(draft.well_text):
            problems.append("no recipient could be read at all; refusing to send blind")
        return problems


class IntendedRecipientsPresent(Guard):
    """Every address the caller asked for must be visible in the pane.

    Outlook's autocomplete substitutes cached contacts, and a recipient count rising is not evidence it
    rose to the address you typed. `aliases` maps an address to a display name it legitimately resolves
    to, for your own mailbox.
    """
    name = "intended recipients present"

    def __init__(self, aliases: dict[str, str] | None = None):
        self.aliases = {k.lower(): v.lower() for k, v in (aliases or {}).items()}

    def __call__(self, draft: Draft) -> list[str]:
        fields = (draft.intended_cc, draft.intended_bcc) if draft.is_reply else (
            draft.intended_to, draft.intended_cc, draft.intended_bcc)
        seen = " ".join([(draft.well_text or ""), *draft.chip_addresses,
                         *draft.chip_display_names]).lower()
        problems = []
        for group in fields:
            for addr in [x.strip() for x in (group or "").split(",") if x.strip()]:
                if addr.lower() not in seen and self.aliases.get(addr.lower(), "\0") not in seen:
                    problems.append(f"recipient {addr} is not visible in the pane")
        return problems


class BodyMatchesIntended(Guard):
    """The whole intended text must appear in the pane, whitespace-normalised.

    Not a prefix check: a prefix passes on a body that was truncated, pasted twice, or left on top of an
    older draft. On failure this reports the first differing offset with both sides, which is the only
    thing that makes a paste failure quick to diagnose.
    """
    name = "body matches intended"

    def __call__(self, draft: Draft) -> list[str]:
        want, got = normalise(draft.intended_body), normalise(draft.body)
        if not want or want in got:
            return []
        i = next((k for k in range(min(len(want), len(got))) if want[k] != got[k]),
                 min(len(want), len(got)))
        return [f"body does not match: {len(got)} chars in the pane vs {len(want)} intended; "
                f"first difference at {i}: intended {want[i:i + 60]!r} vs pane {got[i:i + 60]!r}"]


class BodyMustNotContain(Guard):
    """Refuse when the body still holds text that should have been removed.

    The case this was written for: Outlook auto-inserts a personal signature, and in reply mode the body
    is not wiped first (that would destroy the quoted thread), so the signature sits between your text
    and the quote and goes out over your own name.
    """
    name = "body must not contain"

    def __init__(self, needles):
        self.needles = list(needles)

    def __call__(self, draft: Draft) -> list[str]:
        low = (draft.body or "").lower()
        return [f"body still contains {n!r}" for n in self.needles if str(n).lower() in low]


class BodyMustContain(Guard):
    name = "body must contain"

    def __init__(self, needles):
        self.needles = list(needles)

    def __call__(self, draft: Draft) -> list[str]:
        low = (draft.body or "").lower()
        return [f"body does not contain {n!r}" for n in self.needles if str(n).lower() not in low]


class SubjectMatchesIntended(Guard):
    """Check the subject we set, and do not demand one we never set.

    A reply's subject is Outlook's own `RE: ...`. In some sessions it sits in an editable Subject input;
    in a single-mailbox session an inline reply has **no Subject input at all** and the subject is part of
    the thread header. So when the field is absent this checks nothing rather than reporting "reply has no
    subject" - which would refuse every reply in that shape.
    """
    name = "subject matches intended"

    def __call__(self, draft: Draft) -> list[str]:
        if draft.is_reply:
            if not draft.subject_editable:
                return []
            return [] if normalise(draft.subject) else ["reply has an empty subject field"]
        if not draft.subject_editable:
            return ["a new message has no editable subject field; refusing to send it unaddressed"]
        if draft.intended_subject and draft.subject != draft.intended_subject:
            return [f"subject reads {draft.subject!r}, not {draft.intended_subject!r}"]
        return []


def run(guards, draft: Draft) -> list[str]:
    """Every problem from every guard, each prefixed with the guard that raised it."""
    problems = []
    for guard in guards:
        for problem in guard(draft):
            problems.append(f"[{getattr(guard, 'name', type(guard).__name__)}] {problem}")
    return problems
