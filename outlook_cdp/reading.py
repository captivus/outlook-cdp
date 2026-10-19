"""Reading a mailbox: folders, message bodies and search.

This is the counterpart to `mail.py`, and it has two implementations of the same ideas.

**Use the REST path.**  Outlook on the web talks to its own REST API from inside the page, and every
read here can go through it: `list_folder`, `conversation`, `message`, `search`, `folders`.  It is
faster than the DOM, it returns typed fields instead of `innerText` guesses, it reaches messages in
other folders, and - the reason it matters most - **it does not mark anything read**.  The token and the
mailbox-scoping live in `attachments.py`; this module borrows them rather than keeping a second copy.

The **DOM path** (`*_dom`) is kept because it is the only thing that works when the page has no usable
token, and because what it learned about the reading pane is worth not losing.  Every DOM method opens
messages the way a human does, which **marks them read**.

Why conversation grouping defeated the obvious implementations, since the fix is not obvious either:

* The `/id/<...>` segment of the reading-pane URL is a **ConversationId, not a message id**.  Feeding it
  to `/messages/<id>` answers `404 ... doesn't belong to the targeted mailbox`, which reads like a
  permissions problem and is really a type error.  `$filter=ConversationId eq '<it>'` turns it back into
  individual messages, each with its own `From`, `ReceivedDateTime` and body.
* In the DOM, the pane renders the **whole thread**, so taking the largest block of text returns the
  oldest message - for a reply to something you sent, that is your own outgoing mail.
* A freshly opened thread expands only its newest message; the rest are collapsed one-line stubs with
  **no body in the DOM at all**.  A short thread, though, can come up fully expanded, so neither "the
  first block" nor "the biggest block" is the message you want.  Messages have to be addressed
  individually, which is what `read_open_conversation_dom` does.

DOM behaviours worth not rediscovering, all verified against a live mailbox:

* `div[aria-label="N messages"]` wraps the thread, and is present even for a single message, where it
  reads `"1 messages"` - Outlook's grammar, not a typo here.  It is the authoritative thread length.
* An expanded message is `div[aria-label="Email message"]`; a collapsed one is a plain sibling with a
  preview and no body, and a trusted click on it renders it in place.
* Expanded messages carry stable per-message ids: `MSG_<id>_FROM`, `_TO`, `_CC`, `_BCC`, `_DATETIME`,
  `_SUBJECT`, and the thread's subject is on `CONV_<id>_SUBJECT`.  Every class name in the pane is a
  build hash and will change; these ids and the aria-labels will not.
* **The wake sequence has to be re-asserted after every navigation.**  `Emulation.setFocusEmulation
  Enabled` plus `Page.setWebLifecycleState {"state":"active"}` is not a once-per-session call.  Run a
  search in a background tab and the results list is *genuinely* empty - no rows, no "no results" text -
  until the pair is sent again.  It is indistinguishable from a query that matched nothing.
* **`div[role=option]` is not only a message row.**  Attachment tiles in the reading pane share the
  role, so an unscoped `querySelectorAll` mixes attachments into a folder listing.  Every row query here
  is scoped to `[role=listbox][aria-label^="Message list"]`.
* `/mail/<mailbox>/search/id/<q>?q=<q>` does **not** filter.  It renders the unfiltered folder, which
  looks like a search that matched everything.
"""
from __future__ import annotations

import html as _html
import json
import re
import time
import urllib.parse
from dataclasses import dataclass, field
from html.parser import HTMLParser

from .attachments import (auth_header, list_attachments, mailbox_base,
                          open_conversation_id, rest_get)
from .session import OutlookError, Session

#: The message list, and the scope for every DOM row query.  See the note on `role=option` above.
LIST_SEL = '[role=listbox][aria-label^="Message list"]'
#: One expanded message in the reading pane.
MESSAGE_SEL = '[aria-label="Email message"]'
#: A rendered body.  The **compose** surface carries the same label plus `contenteditable="true"`, which
#: is why `mail.py` qualifies it; nothing here looks outside an `Email message`, so a compose pane cannot
#: be mistaken for a read one.
BODY_SEL = '[aria-label="Message body"]'
SEARCH_INPUT_SEL = "#topSearchInput"

#: Fields worth having on every message. `UniqueBody` is the point of the list: it is the message's own
#: content with the quoted trail already stripped, which is exactly the split the DOM path has to do by
#: hand. Selecting it costs nothing extra.
MESSAGE_FIELDS = ("Id,Subject,From,Sender,ToRecipients,CcRecipients,BccRecipients,ReceivedDateTime,"
                  "HasAttachments,IsRead,IsDraft,ConversationId,BodyPreview,ParentFolderId")
BODY_FIELDS = MESSAGE_FIELDS + ",Body,UniqueBody"

#: Row aria-labels open with zero or more of these flag words before the sender.
_ROW_FLAGS = ("Unread", "Read", "Collapsed", "Expanded", "Has attachments", "Flagged", "Pinned",
              "Draft", "Mentioned", "Encrypted", "Signed", "High importance", "Low importance")
_FLAG_RE = re.compile(r"^\s*(" + "|".join(re.escape(f) for f in _ROW_FLAGS) + r")\b")

_BLOCK_TAGS = {"p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote",
               "table", "section", "article", "header", "footer", "pre", "hr", "ul", "ol", "dd", "dt"}


# --------------------------------------------------------------------------- parsing (no browser)

class _TextExtractor(HTMLParser):
    """Turn a mail body's HTML into readable plain text."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out: list[str] = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "head"):
            self._skip += 1
        elif tag in _BLOCK_TAGS:
            self.out.append("\n")
        elif tag == "td":
            # A cell break has to become whitespace, not nothing: mail signatures and the tables that
            # public bodies paste in are laid out in cells, and joining them with no separator runs
            # words together into one unsearchable string.
            self.out.append(" ")
        elif tag == "img":
            alt = dict(attrs).get("alt")
            if alt:
                self.out.append(f"[{alt}]")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "head"):
            self._skip = max(0, self._skip - 1)
        elif tag in _BLOCK_TAGS:
            self.out.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self.out.append(data)


def html_to_text(content: str) -> str:
    """A mail body's HTML as plain text, with paragraph breaks kept and runs of blank lines collapsed.

    Outlook's REST API returns bodies as HTML unless you ask for text with a `Prefer` header, and the
    header cannot be set through the JSON helper this module borrows, so the conversion happens here.
    Non-breaking spaces are folded to ordinary ones: they are invisible, they are everywhere in mail
    pasted out of Word, and they make every string comparison against a body fail for no visible reason.
    """
    if not content:
        return ""
    if "<" not in content:
        text = content
    else:
        p = _TextExtractor()
        p.feed(content)
        p.close()
        text = "".join(p.out)
    text = _html.unescape(text).replace(" ", " ").replace("​", "")
    lines = [re.sub(r"[ \t]+", " ", l).strip() for l in text.replace("\r\n", "\n").split("\n")]
    out: list[str] = []
    for line in lines:
        if line or (out and out[-1]):
            out.append(line)
    return "\n".join(out).strip()


def parse_row_flags(aria_label: str) -> tuple[dict[str, bool], str]:
    """Split a DOM row's aria-label into its leading flag words and the remaining text.

    Returns `({flag: True, ...}, rest)`, reporting only the flags present.

    **`Collapsed` does not mean "this row is a conversation".**  Single messages carry it too - it
    describes the row widget, not the thread - so a listing that trusted it called almost every row a
    thread.
    """
    flags: dict[str, bool] = {}
    rest = aria_label or ""
    while True:
        m = _FLAG_RE.match(rest)
        if not m:
            break
        flags[m.group(1)] = True
        rest = rest[m.end():]
    return flags, rest.strip()


def parse_count_badge(text: str) -> int | None:
    """`"(2)"` -> 2.  The badge Outlook puts on some rows that stand for more than one message."""
    m = re.fullmatch(r"\((\d+)\)", (text or "").strip())
    return int(m.group(1)) if m else None


def parse_container_count(aria_label: str) -> int | None:
    """`"4 messages"` -> 4, `"1 messages"` -> 1.  The reading pane's own thread length."""
    m = re.fullmatch(r"(\d+)\s+messages?", (aria_label or "").strip())
    return int(m.group(1)) if m else None


def split_sender(text: str) -> tuple[str, str | None]:
    """`"A Name<a@b.gov>"` -> `("A Name", "a@b.gov")`.

    Outlook writes the pane's From line with no space before the angle bracket, and for delegated mail
    as `"Someone<a@b> on behalf of Something<c@d>"`.  In that shape the *first* identity is the human
    who actually sent it, so it is the one returned; `on_behalf_of()` gets the other.
    """
    s = " ".join((text or "").split())
    s = re.split(r"\s+on behalf of\s+", s)[0]
    m = re.match(r"^(.*?)\s*<([^<>]+)>\s*$", s)
    if m:
        return m.group(1).strip(), m.group(2).strip().lower()
    return s.strip(), (s.strip().lower() if "@" in s and " " not in s else None)


def on_behalf_of(text: str) -> tuple[str, str | None] | None:
    """The mailbox a delegated message was sent *for*, or None when it was not delegated."""
    parts = re.split(r"\s+on behalf of\s+", " ".join((text or "").split()))
    return split_sender(parts[1]) if len(parts) > 1 else None


def parse_recipients(text: str) -> list[str]:
    """Pull recipients out of a DOM `To:`/`Cc:`/`Bcc:` field's aria-label or text.

    The field holds addresses when Outlook has them, display names when it does not, and the two mixed
    in one field.  Nothing here resolves a name to an address; the REST path returns both and should be
    preferred whenever a caller needs the address.
    """
    s = " ".join((text or "").split())
    s = re.sub(r"^(To|Cc|Bcc|From)\s*:\s*", "", s, flags=re.I)
    return [p.strip() for p in re.split(r"[;,]\s*(?![^<]*>)", s) if p.strip()]


def strip_glyphs(text: str) -> str:
    """Drop Unicode private-use characters.

    Outlook's icons are glyphs in a private-use font, so they arrive in `innerText` as U+E000-U+F8FF
    characters that look like empty boxes and quietly defeat string comparison.
    """
    return "".join(ch for ch in (text or "") if not ("" <= ch <= ""))


def addresses_in(entries: list[str]) -> list[str]:
    """Just the email addresses out of a list of `"Name <addr>"` strings, lowercased."""
    out = []
    for e in entries or []:
        m = re.search(r"<([^<>@]+@[^<>]+)>", e) or re.fullmatch(r"\s*([^\s<>@]+@[^\s<>]+)\s*", e or "")
        if m:
            out.append(m.group(1).strip().lower())
    return out


def _fmt_address(entry: dict | None) -> str:
    """One REST `EmailAddress` as `"Name <addr>"`, or just the address when there is no name."""
    ea = ((entry or {}).get("EmailAddress") or {}) if isinstance(entry, dict) else {}
    name, addr = (ea.get("Name") or "").strip(), (ea.get("Address") or "").strip()
    if name and addr and name.lower() != addr.lower():
        return f"{name} <{addr}>"
    return addr or name


# --------------------------------------------------------------------------- returned shapes

@dataclass
class Message:
    """One message.  Built from REST, or from one block of the reading pane."""
    id: str | None = None                  # the REST message id; None on the DOM path
    conversation_id: str | None = None
    subject: str = ""
    sender: str = ""
    sender_address: str | None = None
    on_behalf_of: str | None = None
    to: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    bcc: list[str] = field(default_factory=list)
    received: str = ""                     # ISO 8601 UTC from REST; "" on the DOM path
    date_text: str = ""                    # as displayed, e.g. "Tue 9/29/2026 7:33 PM"
    unread: bool = False
    is_draft: bool = False
    has_attachments: bool = False
    attachment_names: list[str] = field(default_factory=list)
    body: str = ""                         # this message's own text, without the quoted trail
    full_body: str = ""                     # the body as sent, quoted trail included (REST only)
    quoted_history: str = ""               # the trail on its own (DOM path)
    preview: str = ""                      # BodyPreview, or a collapsed stub's preview line
    folder: str | None = None
    index: int = 0                         # position in the thread; 0 = oldest
    expanded: bool = True                  # DOM path: was a body actually rendered
    source: str = "rest"                   # "rest" or "dom"

    @property
    def to_addresses(self) -> list[str]:
        return addresses_in(self.to)

    @property
    def cc_addresses(self) -> list[str]:
        return addresses_in(self.cc)


@dataclass
class Row:
    """One row of a folder listing or a search result. Metadata only; reading it marks nothing read."""
    index: int = 0
    senders: list[str] = field(default_factory=list)
    sender: str = ""
    sender_address: str | None = None
    subject: str = ""
    date_text: str = ""            # as shown, e.g. "Tue 7:33 PM" or "9/28/2026"
    received: str = ""             # the full timestamp: ISO from REST, the cell's title from the DOM
    unread: bool = False
    has_attachments: bool = False
    attachment_names: list[str] = field(default_factory=list)
    count_badge: int | None = None  # DOM only: Outlook's raw `(N)` badge, or None when absent
    message_count: int = 1          # messages of this thread in this folder
    is_conversation: bool = False   # more than one message, here or in another folder
    expandable: bool = False        # DOM only: the row offers `Expand conversation`
    conversation_id: str | None = None
    message_id: str | None = None   # REST only; the newest message of the row's thread
    message_ids: list[str] = field(default_factory=list)   # REST only; every message in this folder
    folder: str | None = None
    aria_label: str = ""
    source: str = "rest"

    @property
    def id(self) -> str | None:
        """The stable handle: a conversation id, which both paths can produce."""
        return self.conversation_id


@dataclass
class Conversation:
    """A whole thread, oldest message first."""
    subject: str = ""
    conversation_id: str | None = None
    expected_count: int | None = None
    messages: list[Message] = field(default_factory=list)
    source: str = "rest"

    @property
    def newest(self) -> Message | None:
        """The last message in the thread - what a caller almost always means by "the reply"."""
        return self.messages[-1] if self.messages else None

    @property
    def oldest(self) -> Message | None:
        return self.messages[0] if self.messages else None

    @property
    def complete(self) -> bool:
        """Did every message the thread claims to hold actually arrive?

        On the REST path this is always true - the query *is* the thread.  On the DOM path False means
        part of the thread is elsewhere, or a block would not expand, and `newest` may therefore not be
        the newest message.  Check it before trusting `newest`; not quietly returning the wrong message
        is the whole point of this module.
        """
        return self.expected_count is None or len(self.messages) >= self.expected_count

    def from_address(self, address: str) -> list[Message]:
        a = address.lower()
        return [m for m in self.messages if (m.sender_address or "").lower() == a]


# --------------------------------------------------------------------------- the reader

class Reader:
    """Read-only operations on one Outlook tab.

        s = Session(os.environ["CHROME_AGENT_INSTANCE"])
        r = Reader(s)

        for row in r.list_folder():
            print(row.received, row.sender, row.subject, row.message_count)

        conv = r.conversation(rows[0].conversation_id)
        print(conv.newest.sender, conv.newest.date_text)
        print(conv.newest.body)

        for hit in r.search("some phrase"):
            print(hit.folder, hit.sender, hit.subject)

    Nothing here sends, replies, drafts, deletes, moves or flags.  The REST methods change nothing at
    all.  The `*_dom` methods drive the UI and therefore **mark what they open read**.
    """

    def __init__(self, session: Session, *, wake: bool = True, mailbox: str | None = None):
        """`mailbox` overrides the tab URL's, which a tab on plain ``/mail/`` does not carry."""
        self.s = session
        self._token: str | None = None
        self._base: str | None = ("/api/v2.0/users/" + urllib.parse.quote(mailbox.lower(), safe="")
                                  if mailbox else None)
        self._folders: dict[str, str] | None = None
        if wake:
            self.wake()

    # ------------------------------------------------------------------ plumbing

    def wake(self) -> None:
        """Make a backgrounded tab render.

        Deliberately not a call to `Session.focus_emulation()`: this module must not depend on a
        `session.py` that is being changed in parallel, and the lifecycle call matters as much as the
        focus one.  **This belongs in `Session`, with every navigation re-asserting it** - a tab that
        has just navigated or just run a search renders an empty list until both calls land again.
        """
        self.s.cdp("Emulation.setFocusEmulationEnabled", {"enabled": True})
        self.s.cdp("Page.setWebLifecycleState", {"state": "active"})

    def _js(self, expression: str):
        return self.s.js(expression)

    # ------------------------------------------------------------------ REST plumbing

    @property
    def base(self) -> str:
        """The REST prefix for the mailbox **on screen**.

        Not `/api/v2.0/me`: `me` is the signed-in account, which for a shared mailbox is somebody
        else's mail entirely, and the call succeeds while returning the wrong person's messages.
        """
        if self._base is None:
            self._base = mailbox_base(self.s)
        return self._base

    def token(self, refresh: bool = False) -> str:
        """The page's own REST access token, cached for this Reader.

        Tokens last about an hour and Outlook refreshes them silently while the tab is alive, so an
        expired one means the tab has been idle rather than that anything is broken.
        """
        if refresh or self._token is None:
            self._token = auth_header(self.s, refresh=refresh)
        return self._token

    def rest_available(self) -> bool:
        """Can the REST path be used at all? False means fall back to the `*_dom` methods."""
        try:
            self.token()
            return bool(self.base)
        except OutlookError:
            return False

    def _get(self, path: str) -> dict:
        try:
            return rest_get(self.s, path, self.token())
        except OutlookError as exc:
            # A token that expired mid-run is the common failure and is entirely recoverable, so retry
            # once with a fresh one before handing the caller an error they cannot act on.
            if "401" in str(exc) or "expired" in str(exc).lower():
                return rest_get(self.s, path, self.token(refresh=True))
            raise

    def folders(self) -> list[dict]:
        """Every mail folder, as `{"id", "name", "total", "unread"}`."""
        rows = self._get(f"{self.base}/mailfolders?$top=100"
                         f"&$select=Id,DisplayName,TotalItemCount,UnreadItemCount").get("value") or []
        return [{"id": r.get("Id"), "name": r.get("DisplayName"),
                 "total": r.get("TotalItemCount"), "unread": r.get("UnreadItemCount")} for r in rows]

    def _folder_names(self) -> dict[str, str]:
        if self._folders is None:
            self._folders = {f["id"]: f["name"] for f in self.folders() if f.get("id")}
        return self._folders

    # ------------------------------------------------------------------ REST: messages

    def _message(self, raw: dict, *, index: int = 0, folder: str | None = None) -> Message:
        # A thread spans folders, so which folder each message is in is part of reading it: it is how
        # "their reply" is told from "our own copy of what we sent".
        if folder is None and raw.get("ParentFolderId"):
            folder = self._folder_names().get(raw["ParentFolderId"])
        sender = _fmt_address(raw.get("From")) or _fmt_address(raw.get("Sender"))
        name, addr = split_sender(sender)
        on_behalf = None
        sender_entry = _fmt_address(raw.get("Sender"))
        if sender_entry and sender_entry != sender:
            # `From` is who the message says it is from and `Sender` is the mailbox that actually put
            # it on the wire. They differ for delegated and shared-mailbox mail, and the difference is
            # worth keeping rather than flattening: it is how "sent on behalf of" is represented.
            on_behalf = split_sender(sender_entry)[1] or sender_entry
        body_html = (raw.get("UniqueBody") or {}).get("Content") or ""
        full_html = (raw.get("Body") or {}).get("Content") or ""
        return Message(
            id=raw.get("Id"), conversation_id=raw.get("ConversationId"),
            subject=raw.get("Subject") or "", sender=name, sender_address=addr,
            on_behalf_of=on_behalf,
            to=[_fmt_address(e) for e in raw.get("ToRecipients") or []],
            cc=[_fmt_address(e) for e in raw.get("CcRecipients") or []],
            bcc=[_fmt_address(e) for e in raw.get("BccRecipients") or []],
            received=raw.get("ReceivedDateTime") or "",
            date_text=raw.get("ReceivedDateTime") or "",
            unread=not raw.get("IsRead", True), is_draft=bool(raw.get("IsDraft")),
            has_attachments=bool(raw.get("HasAttachments")),
            body=html_to_text(body_html or full_html),
            full_body=html_to_text(full_html),
            preview=(raw.get("BodyPreview") or "").strip(),
            folder=folder, index=index, source="rest")

    def message(self, message_id: str, *, body: bool = True,
                attachment_names: bool = False) -> Message:
        """One message by its REST id. Marks nothing read."""
        sel = BODY_FIELDS if body else MESSAGE_FIELDS
        raw = self._get(f"{self.base}/messages/{urllib.parse.quote(message_id, safe='')}?$select={sel}")
        m = self._message(raw)
        if attachment_names:
            self._name_attachments([m])
        return m

    def _name_attachments(self, messages: list[Message]) -> None:
        """Fill in `attachment_names` for the messages that have any.

        Off by default and done in a second pass, because it is one request per message: `HasAttachments`
        rides along with the message but the filenames do not. Inline signature logos are excluded, so a
        message whose only attachments are a sender's three logos comes back with an empty list and
        `has_attachments` still true - which is correct, and worth knowing before treating the pair as
        contradictory.
        """
        for m in messages:
            if not (m.has_attachments and m.id):
                continue
            try:
                m.attachment_names = [a.name for a in list_attachments(self.s, m.id, self.token())]
            except OutlookError:
                # A message whose attachment list cannot be read is still a readable message; leaving
                # the names empty is better than failing the whole thread over them.
                m.attachment_names = []

    def conversation(self, conversation_id: str, *, body: bool = True, top: int = 100,
                     attachment_names: bool = True) -> Conversation:
        """Every message of one thread, oldest first, wherever in the mailbox it lives.

        This is the answer to conversation grouping.  The reading pane may be painting any member of the
        thread, or only some of them; this returns all of them, individually, each with its own sender,
        timestamp and body, and **marks nothing read**.

        Threads span folders, so the reply in the Inbox and the request in Sent Items both come back
        here even though a folder listing shows only one of them.
        """
        sel = BODY_FIELDS if body else MESSAGE_FIELDS
        path = (f"{self.base}/messages?$top={int(top)}&$select={sel}"
                f"&$filter=ConversationId%20eq%20'{urllib.parse.quote(conversation_id, safe='')}'")
        rows = self._get(path).get("value") or []
        if not rows:
            raise OutlookError(
                f"no messages for conversation {conversation_id[:24]}... . If the id came from a row's "
                f"data-convid it should match; if it came from a URL, remember the /id/ segment is "
                f"URL-encoded and has to be unquoted first.")
        rows.sort(key=lambda r: r.get("ReceivedDateTime") or "")
        msgs = [self._message(r, index=i) for i, r in enumerate(rows)]
        if attachment_names:
            self._name_attachments(msgs)
        return Conversation(subject=next((m.subject for m in msgs if m.subject), ""),
                            conversation_id=conversation_id, expected_count=len(msgs),
                            messages=msgs, source="rest")

    def open_conversation(self, *, body: bool = True) -> Conversation:
        """The thread the reading pane is showing, read over REST.

        The pane's URL carries a **ConversationId**, not a message id - which is why navigating to
        `/messages/<that id>` answers 404 and why the id alone never produced the reply.
        """
        conv = open_conversation_id(self.s)
        if not conv:
            raise OutlookError("no message is open in the reading pane (the tab URL has no /id/ part)")
        return self.conversation(conv, body=body)

    def newest_in_conversation(self, conversation_id: str, *, body: bool = True, **kw) -> Message:
        """The newest message of a thread.

        The call most readers want, and the one the DOM made hard: with grouping on, the obvious
        implementation returns the *oldest* message - your own outgoing mail - and looks like it worked.
        """
        conv = self.conversation(conversation_id, body=body, **kw)
        newest = conv.newest
        if newest is None:
            raise OutlookError("the thread holds no messages")
        return newest

    def reply_to(self, conversation_id: str, *, mailbox: str | None = None,
                 body: bool = True, **kw) -> Message | None:
        """The newest message in a thread that was **not** sent from this mailbox, or None.

        For a thread that is "our request, their answer", this is the answer, and it stays right when
        the thread later grows another message of ours - which `newest` alone does not.
        """
        mine = (mailbox or self.s.mailbox_identity() or "").lower()
        conv = self.conversation(conversation_id, body=body, **kw)
        for m in reversed(conv.messages):
            if (m.sender_address or "").lower() != mine:
                return m
        return None

    # ------------------------------------------------------------------ REST: folder listing

    def list_folder(self, folder: str = "inbox", *, top: int = 100, group: bool = True,
                    body: bool = False) -> list[Row]:
        """A folder's contents, newest first.

        `folder` is a well-known name (`inbox`, `sentitems`, `drafts`, `archive`, `deleteditems`) or a
        folder id from `folders()`.

        REST returns **messages**; Outlook's own list shows **threads**.  With `group=True` messages are
        grouped by `ConversationId` so one row stands for one thread as the UI would show it, with
        `message_count` and `message_ids` covering the members that are in *this* folder - the rest of
        the thread lives in other folders and `conversation()` will fetch it.  With `group=False` every
        message gets its own row, which is usually what a program wants.

        `$orderby=ReceivedDateTime desc` is asked for here rather than sorted afterwards so that `top`
        means "the newest N".  Note it cannot be combined with `HasAttachments eq true` - the server
        answers `InefficientFilter` - so filter that in Python.
        """
        sel = BODY_FIELDS if body else MESSAGE_FIELDS
        path = (f"{self.base}/mailfolders/{urllib.parse.quote(folder)}/messages"
                f"?$top={int(top)}&$orderby=ReceivedDateTime%20desc&$select={sel}")
        raws = self._get(path).get("value") or []
        rows: list[Row] = []
        by_conv: dict[str, Row] = {}
        for raw in raws:
            m = self._message(raw)
            cid = m.conversation_id
            if group and cid and cid in by_conv:
                row = by_conv[cid]
                row.message_count += 1
                row.is_conversation = True
                row.message_ids.append(m.id or "")
                if m.sender and m.sender not in row.senders:
                    row.senders.append(m.sender)
                row.unread = row.unread or m.unread
                row.has_attachments = row.has_attachments or m.has_attachments
                continue
            row = Row(senders=[m.sender] if m.sender else [], sender=m.sender,
                      sender_address=m.sender_address, subject=m.subject,
                      date_text=m.received, received=m.received, unread=m.unread,
                      has_attachments=m.has_attachments, message_count=1,
                      conversation_id=cid, message_id=m.id, message_ids=[m.id or ""],
                      folder=m.folder or folder, source="rest")
            rows.append(row)
            if cid:
                by_conv[cid] = row
        for i, row in enumerate(rows):
            row.index = i
            row.sender = "; ".join(row.senders)
        return rows

    def unread(self, folder: str = "inbox", *, top: int = 100) -> list[Row]:
        """Unread messages in a folder, newest first, one row per message."""
        return [r for r in self.list_folder(folder, top=top, group=False) if r.unread]

    # ------------------------------------------------------------------ REST: search

    def search(self, query: str, *, folder: str | None = None, top: int = 25,
               body: bool = False) -> list[Row]:
        """Search the mailbox, or one folder, and get back **individual messages**.

        Each hit is one message with its own sender, timestamp and `folder`, so a thread that matches
        three times appears three times - which makes search the cheapest way to address one message
        inside a thread.

        Two traps: `$search` **cannot** be combined with `$orderby` (the server answers
        `SearchWithOrderBy`), so results are sorted here; and Outlook's own search *URL*,
        `/mail/<mailbox>/search/id/<q>?q=<q>`, does not filter at all - it renders the unfiltered
        folder, which looks like a query that matched everything.
        """
        sel = BODY_FIELDS if body else MESSAGE_FIELDS
        scope = f"mailfolders/{urllib.parse.quote(folder)}/messages" if folder else "messages"
        path = (f"{self.base}/{scope}?$top={int(top)}&$select={sel},ParentFolderId"
                f"&$search={urllib.parse.quote(json.dumps(query))}")
        raws = self._get(path).get("value") or []
        raws.sort(key=lambda r: r.get("ReceivedDateTime") or "", reverse=True)
        rows = []
        for i, raw in enumerate(raws):
            m = self._message(raw, index=i)
            rows.append(Row(index=i, senders=[m.sender] if m.sender else [], sender=m.sender,
                            sender_address=m.sender_address, subject=m.subject,
                            date_text=m.received, received=m.received, unread=m.unread,
                            has_attachments=m.has_attachments, conversation_id=m.conversation_id,
                            message_id=m.id, message_ids=[m.id or ""], folder=m.folder,
                            source="rest"))
        return rows

    # ================================================================== the DOM path
    # Everything below drives the UI.  It is the fallback for a page with no usable token, and it
    # **marks messages read**.  Prefer the REST methods above.

    def _click_element(self, selector_js: str, *, into_view: bool = True, y_frac: float = 0.5,
                       x_frac: float = 0.5, y_offset: float | None = None) -> bool:
        """Trusted click on an element given a JS expression for it.

        Scrolls it into view, then **re-reads the rectangle**, then refuses a centre that is still
        outside the viewport.  `Input.dispatchMouseEvent` takes viewport coordinates, and the reading
        pane scrolls independently of the page, so an element can have a perfectly healthy box at
        y=2100 in a 1321px viewport.  A click there is accepted by the protocol and lands on nothing -
        which is exactly the "the click succeeded and nothing happened" failure that cost the
        attachment work a day.
        """
        measure = ("(()=>{const e=%s;if(!e)return null;%s"
                   "const r=e.getBoundingClientRect();"
                   "if(r.width<2||r.height<2)return null;"
                   "const x=r.x+r.width*%r, y=r.y+%s;"
                   "return [x, y, innerWidth, innerHeight]})()")
        y_expr = repr(y_offset) if y_offset is not None else "r.height*%r" % y_frac
        got = self._js(measure % (selector_js,
                                  "e.scrollIntoView({block:'center'});" if into_view else "",
                                  x_frac, y_expr))
        if not got:
            return False
        if into_view:
            # The scroll is animated, so a rectangle measured in the same expression that requested it
            # points at where the element *was*.
            time.sleep(0.4)
            got = self._js(measure % (selector_js, "", x_frac, y_expr))
            if not got:
                return False
        x, y, vw, vh = got
        if not (0 <= x <= vw and 0 <= y <= vh):
            return False
        self.s.click(x, y)
        return True

    #: Classify every cell of every row in one pass.  Written as one expression because the list is
    #: re-rendered between chrome-agent's separate one-shot connections, so a second call can measure a
    #: different row than the first.
    _ROW_JS = r"""
(()=>{
 const lb=document.querySelector(%(list)s);
 if(!lb) return null;
 const isGlyph=t=>/^[-\s]*$/.test(t);
 const DATE=/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)?\s*\d{1,2}\/\d{1,2}\/\d{2,4}|\d{1,2}:\d{2}\s*(AM|PM)/i;
 // Inline elements Outlook puts *inside* a cell rather than beside it. MARK is the important one: in
 // search results the query terms come back wrapped in <mark>, which splits one subject into three
 // leaf nodes. Reading leaves directly therefore returned a subject of "Assessment", then "Roll", as
 // if they were separate columns, and shifted every later column along by two.
 const INLINE={MARK:1,I:1,B:1,EM:1,STRONG:1,BR:1,WBR:1,U:1};
 const cellOf=e=>{
   let c=e; const chain=[c];
   while(c.tagName==='MARK'&&c.parentElement){ c=c.parentElement; chain.push(c); }
   while(c.parentElement && c.parentElement!==lb &&
         [...c.parentElement.children].every(k=>k===c||INLINE[k.tagName])){ c=c.parentElement; chain.push(c); }
   // The title can sit anywhere on the way up - the sender's address is on the inner span while the
   // cell coalesces to its wrapper - so take the first one found rather than only the cell's own.
   // Reading it off the cell alone silently dropped every sender address.
   const title=(chain.map(k=>k.getAttribute('title')||'').find(t=>t)||'');
   return [c, title];
 };
 return [...lb.querySelectorAll('div[role=option]')].map((o,i)=>{
   let date='', received='', count=null, folder=null;
   const names=[], plain=[]; const seen=new Set();
   for(const leaf of [...o.querySelectorAll('*')].filter(e=>e.children.length===0)){
     if(!(leaf.textContent||'').trim()) continue;
     // The avatar monogram duplicates the sender and is not a column of its own.
     if(leaf.closest('[role=img]')) continue;
     const [c, title]=cellOf(leaf);
     if(seen.has(c)) continue; seen.add(c);
     const t=(c.textContent||'').trim();
     if(title && DATE.test(title)){ date=t; received=title; continue; }
     // An attachment chip carries the filename in a title. The `@` test is not decoration: the sender
     // cell's title is the sender's address, and an address also ends in a dot-extension, so without
     // it every row with a readable sender reported a phantom attachment named after the sender and
     // lost its sender and subject columns to the attachment list.
     if(title && title.indexOf('@')<0 && /\.[A-Za-z0-9]{1,6}$/.test(title)){
       if(!names.includes(title)) names.push(title); continue; }
     if(!t || isGlyph(t)) continue;
     if(/^\(\d+\)$/.test(t)){ count=parseInt(t.slice(1,-1),10); continue; }
     // Search results tag each hit with the folder it lives in, as a Fluent InteractionTag. Matched on
     // the component class rather than on position, because the position of the tag moves.
     if(c.closest('[class*="InteractionTag"]')){ folder=t; continue; }
     // Row hover actions are buttons; nothing inside one is data.
     if(c.closest('button,[role=button]')) continue;
     plain.push({t:t, title:title});
   }
   const btns=[...o.querySelectorAll('button,[role=button]')].map(b=>b.getAttribute('aria-label')||'');
   return {index:i, aria_label:o.getAttribute('aria-label')||'',
           convid:o.getAttribute('data-convid'), date_text:date, received:received,
           count:count, cells:plain, attachments:names, folder:folder,
           expandable:btns.some(b=>/^Expand conversation$/i.test(b))};
 });
})()"""

    def _read_rows_dom(self, in_search: bool) -> list[Row]:
        raw = self._js(self._ROW_JS % {"list": json.dumps(LIST_SEL)})
        if raw is None:
            raise OutlookError(
                "no message list in this tab. Either it is not on a mail folder, or the tab has not "
                "rendered - call wake() and retry; a backgrounded tab shows an empty list.")
        rows = []
        for r in raw:
            flags, _ = parse_row_flags(r["aria_label"])
            cells = [c["t"] for c in r["cells"]]
            titles = [c["title"] for c in r["cells"]]
            # Column order is sender, subject, and in search results a folder name. Anything past that
            # is preview text Outlook sometimes renders, and is ignored rather than guessed at.
            sender = cells[0] if cells else ""
            subject = cells[1] if len(cells) > 1 else ""
            folder = r["folder"] or (cells[2] if in_search and len(cells) > 2 else None)
            addr = next((t.strip().lower() for t in titles if t and "@" in t), None)
            senders = [p.strip() for p in sender.split(";") if p.strip()]
            rows.append(Row(
                index=r["index"], senders=senders, sender=sender, sender_address=addr,
                subject=subject, date_text=r["date_text"], received=r["received"],
                unread=bool(flags.get("Unread")),
                has_attachments=bool(flags.get("Has attachments")) or bool(r["attachments"]),
                attachment_names=r["attachments"], count_badge=r["count"],
                message_count=r["count"] or 1,
                is_conversation=bool(r["count"] and r["count"] > 1) or bool(r["expandable"]),
                expandable=bool(r["expandable"]), conversation_id=r["convid"], folder=folder,
                aria_label=r["aria_label"], source="dom"))
        return rows

    def list_folder_dom(self, *, scroll: bool = True, max_scrolls: int = 60) -> list[Row]:
        """The folder the tab is on, read out of the DOM. Marks nothing read; prefer `list_folder()`.

        **This returns only what the list has rendered - roughly one screenful.**  In one real mailbox
        that was 23 rows of 88 (`aria-setsize` on the first row gives the true total, so the shortfall
        is at least detectable).  It is not for want of trying: Outlook's message list is a custom
        virtual scroller, and

        * assigning `scrollTop` on it, or on any ancestor, does nothing - the value reads back as 0;
        * every ancestor of the listbox reports `scrollHeight == clientHeight`, so the element that
          scrolls is *inside* the listbox, not above it;
        * `lastRow.scrollIntoView()` repeated to the bottom adds no rows;
        * `Input.dispatchMouseEvent {"type":"mouseWheel"}` - the one thing that would work, being a
          trusted scroll - **never returns** over chrome-agent's one-shot CDP connection and times the
          call out;
        * arrow keys would page it, but they move the **selection**, which opens messages and marks
          them read, so they are not an option for a read path.

        `list_folder()` over REST has no such limit and is the answer.  The scroll attempt is kept
        because it costs nothing and would start working if Outlook's scroller ever becomes a real one.

        **`message_count` here is a hint, not a count.**  Outlook renders the `(N)` badge on only some
        multi-message rows - across one real folder it appeared on two threads and was absent from seven
        others that demonstrably were threads, and it vanished from both as soon as they were read.  On
        one of the two it disagreed with the reading pane, saying 3 where the pane said 4.  Threads are
        also counted across folders, so a row whose thread is one message here plus one in Sent Items
        shows no badge at all.  `expandable` is the dependable "this is a thread" signal; a real count
        needs `conversation()`.
        """
        merged: dict[str, Row] = {}
        order: list[str] = []

        def absorb(rows: list[Row]) -> None:
            for row in rows:
                key = row.conversation_id or f"{row.aria_label}|{row.received}"
                if key not in merged:
                    merged[key] = row
                    order.append(key)

        absorb(self._read_rows_dom(in_search=False))
        if scroll:
            # The scroller is **inside** the listbox, not an ancestor of it: every ancestor reports
            # scrollHeight == clientHeight, so walking upwards finds nothing that scrolls and the
            # listing silently stopped at one screenful - 23 rows out of 88. Walk up from a row
            # instead, which passes through the element that actually scrolls.
            scroller = ("(()=>{const l=document.querySelector(%s);if(!l)return null;"
                        "let e=l.querySelector('div[role=option]');if(!e)return null;"
                        "while(e&&e!==l&&e.scrollHeight<=e.clientHeight+4)e=e.parentElement;"
                        "return (e&&e.scrollHeight>e.clientHeight+4)?e:null})()"
                        % json.dumps(LIST_SEL))
            for _ in range(max_scrolls):
                moved = self._js("(()=>{const e=%s;if(!e)return false;const b=e.scrollTop;"
                                 "e.scrollTop=b+e.clientHeight*0.8;return e.scrollTop>b+2})()" % scroller)
                if not moved:
                    break
                time.sleep(0.5)
                absorb(self._read_rows_dom(in_search=False))
            self._js("(()=>{const e=%s;if(e)e.scrollTop=0;return 1})()" % scroller)
            time.sleep(0.3)

        rows = [merged[k] for k in order]
        for i, row in enumerate(rows):
            row.index = i
        return rows

    def _row_finder(self, *, conversation_id=None, match=None, index=None) -> str:
        base = "[...document.querySelector(%s).querySelectorAll('div[role=option]')]" % json.dumps(LIST_SEL)
        if conversation_id is not None:
            return "%s.find(o=>o.getAttribute('data-convid')===%s)" % (base, json.dumps(conversation_id))
        if match is not None:
            return ("%s.find(o=>new RegExp(%s,'i').test(o.getAttribute('aria-label')||''))"
                    % (base, json.dumps(match)))
        if index is not None:
            return "%s[%d]" % (base, int(index))
        raise OutlookError("need one of conversation_id, match or index")

    def open_conversation_dom(self, *, conversation_id: str | None = None, match: str | None = None,
                              index: int | None = None, expand_all: bool = True,
                              timeout: float = 25.0) -> Conversation:
        """Click a row open and read the whole thread out of the reading pane.

        **This marks the message read**, which the REST path does not; use `conversation()` unless you
        specifically need what is painted.  `match` is a case-insensitive regex against the row's
        aria-label, the convention `Mailbox.open_reply` uses; prefer `conversation_id` because a label
        is not unique when two senders use the same subject.
        """
        finder = self._row_finder(conversation_id=conversation_id, match=match, index=index)
        if not self._js("!!(%s)" % finder):
            raise OutlookError("no row matched. List the folder first and pass a conversation_id.")
        # Remember which thread was showing. Clicking an already-selected row is a no-op, and without
        # this the read below would return the previous thread and look like a success.
        before = self._js("(()=>{const c=document.querySelector('[id^=\"CONV_\"][id$=\"_SUBJECT\"]');"
                          "return c?c.id:''})()")
        if not self._click_element(finder, y_offset=22):
            raise OutlookError("the matched row has no clickable box inside the viewport")
        # Re-assert the wake sequence: the pane is a fresh render and a backgrounded tab leaves it blank.
        self.wake()
        want = conversation_id
        ok = self.s.wait_for(
            "(()=>{const c=document.querySelector('[id^=\"CONV_\"][id$=\"_SUBJECT\"]');"
            "return !!c && !!document.querySelector('%s') && %s})()"
            % (MESSAGE_SEL,
               ("c.id.indexOf(%s)>=0" % json.dumps(want[-12:])) if want else
               ("c.id!==%s" % json.dumps(before)) if before else "true"),
            timeout=timeout)
        if not ok and not self.s.wait_for("!!document.querySelector('%s')" % MESSAGE_SEL, timeout=5.0):
            raise OutlookError("clicked the row but no message rendered in the reading pane")
        time.sleep(0.6)
        return self.read_open_conversation_dom(expand_all=expand_all)

    _STRIP_JS = r"""
(()=>{
 const c=[...document.querySelectorAll('[aria-label]')]
   .find(e=>/^\d+ messages?$/.test(e.getAttribute('aria-label')));
 if(!c) return null;
 const b=c.querySelector('%(msg)s');
 if(!b) return null;
 // The per-message wrapper is the ancestor of a message block whose own parent is the container's
 // single child. Found by walking up rather than by class, because every class here is a build hash.
 let w=b; while(w.parentElement && w.parentElement.parentElement!==c) w=w.parentElement;
 return w.parentElement;
})()"""

    def _strip_js(self) -> str:
        return self._STRIP_JS % {"msg": MESSAGE_SEL}

    def _collapsed_stub_js(self) -> str:
        return ("(()=>{const s=%s;if(!s)return null;"
                "return [...s.children].find(k=>k.innerText.trim() && !k.querySelector('%s'))||null})()"
                % (self._strip_js(), MESSAGE_SEL))

    def expand_all(self, max_rounds: int = 25) -> int:
        """Expand every collapsed message in the open thread. Returns how many were expanded.

        A freshly opened thread expands only its newest message; the older ones are one-line stubs with
        **no body in the DOM at all**.  They are ordinary siblings of the expanded blocks, and a trusted
        click renders one in place.

        What does *not* work, so it is not tried again: the row's `Expand conversation` button in the
        message **list** flips the row's aria-label to "Expanded", adds no `div[role=option]` child
        rows, and the element it reveals has a zero-size box, so there is nothing to click.  The
        expansion that matters happens in the reading pane, not in the list.
        """
        expanded = 0
        for _ in range(max_rounds):
            if not self._js("!!(%s)" % self._collapsed_stub_js()):
                break
            before = self._js("document.querySelectorAll('%s').length" % MESSAGE_SEL)
            if not self._click_element(self._collapsed_stub_js(), x_frac=0.4):
                break
            if not self.s.wait_for("document.querySelectorAll('%s').length>%d" % (MESSAGE_SEL, before),
                                   timeout=8.0):
                # A stub that will not expand stays a stub and is reported through
                # Conversation.complete, rather than being silently skipped.
                break
            expanded += 1
        return expanded

    _READ_JS = r"""
(()=>{
 const container=[...document.querySelectorAll('[aria-label]')]
   .find(e=>/^\d+ messages?$/.test(e.getAttribute('aria-label')));
 const convEl=document.querySelector('[id^="CONV_"][id$="_SUBJECT"]');
 const strip=%(strip)s;
 const field=(mid,name)=>{
   const e=document.getElementById('MSG_'+mid+'_'+name);
   if(!e) return null;
   const inner=e.querySelector('[aria-label]');
   const al=inner?inner.getAttribute('aria-label'):null;
   return (al && al.trim()) ? al : e.innerText;
 };
 const out=[];
 const wrappers = strip ? [...strip.children].filter(k=>k.innerText.trim()) : [];
 wrappers.forEach(w=>{
   const m=w.querySelector('%(msg)s');
   if(!m){
     // A stub that would not expand: report its metadata and preview, never an empty body.
     const lines=w.innerText.split('\n').map(s=>s.trim()).filter(Boolean);
     out.push({expanded:false, sender:lines[1]||'', date_text:lines[2]||'',
               preview:lines.slice(3).join(' '), message_id:null,
               to:null, cc:null, bcc:null, body:'', history:'', attachments:[]});
     return;
   }
   const from=m.querySelector('[id^="MSG_"][id$="_FROM"]');
   const mid=from ? from.id.slice(4,-5) : null;
   const bodies=[...m.querySelectorAll('%(body)s')];
   // DOM order inside a message block is: the message's own body, then the quoted trail that
   // "Show message history" reveals. Concatenating them is what makes a two-line reply look like a
   // copy of the mail it answers.
   const own=bodies.length?bodies[0].innerText:'';
   const hist=bodies.slice(1).map(b=>b.innerText).join('\n');
   const att=[...m.querySelectorAll('[role=option]')]
     .map(o=>(o.getAttribute('aria-label')||'').replace(/\s+Open\s+[\d.]+\s*[KMGT]?B\s*$/i,'').trim())
     .filter(Boolean);
   out.push({expanded:true, message_id:mid,
             sender:from?from.innerText:'',
             date_text:mid?(field(mid,'DATETIME')||''):'',
             to:mid?field(mid,'TO'):null, cc:mid?field(mid,'CC'):null, bcc:mid?field(mid,'BCC'):null,
             body:own, history:hist, attachments:att, preview:''});
 });
 return {subject:convEl?convEl.innerText:'',
         conversation_id:convEl?convEl.id.slice(5,-8):null,
         expected:container?container.getAttribute('aria-label'):null,
         messages:out};
})()"""

    def read_open_conversation_dom(self, *, expand_all: bool = True) -> Conversation:
        """Read the thread the pane is already showing. Opens nothing, so marks nothing newly read."""
        if expand_all:
            self.expand_all()
        raw = self._js(self._READ_JS % {"strip": self._strip_js(), "msg": MESSAGE_SEL,
                                        "body": BODY_SEL})
        if not raw:
            raise OutlookError("the reading pane holds no message. Select a row first, and remember a "
                               "backgrounded tab renders nothing until wake() is re-asserted.")
        messages = []
        for i, m in enumerate(raw["messages"]):
            name, addr = split_sender(m["sender"])
            obo = on_behalf_of(m["sender"])
            messages.append(Message(
                index=i, id=None, subject=strip_glyphs(raw["subject"]).strip(),
                conversation_id=raw["conversation_id"],
                sender=name, sender_address=addr,
                on_behalf_of=(obo[1] or obo[0]) if obo else None,
                to=parse_recipients(m["to"] or ""), cc=parse_recipients(m["cc"] or ""),
                bcc=parse_recipients(m["bcc"] or ""),
                date_text=strip_glyphs(m["date_text"]).strip(),
                body=(m["body"] or "").strip(), quoted_history=(m["history"] or "").strip(),
                attachment_names=m["attachments"], has_attachments=bool(m["attachments"]),
                expanded=m["expanded"], preview=(m["preview"] or "").strip(), source="dom"))
        return Conversation(subject=strip_glyphs(raw["subject"]).strip(),
                            conversation_id=raw["conversation_id"],
                            expected_count=parse_container_count(raw["expected"] or ""),
                            messages=messages, source="dom")

    def read_newest_dom(self, **kw) -> Message:
        """Open a thread in the UI and return only its newest message, refusing to guess.

        It raises rather than returning a best effort because of the failure it was written for: the
        obvious implementation hands back the oldest message in the thread and looks like it worked.
        """
        conv = self.open_conversation_dom(**kw)
        newest = conv.newest
        if newest is None:
            raise OutlookError("the thread rendered no messages")
        if not conv.complete:
            raise OutlookError(
                f"this thread claims {conv.expected_count} messages but only {len(conv.messages)} "
                f"rendered, so the newest may not be here. Use conversation() over REST, which is not "
                f"limited to what the pane painted.")
        if not newest.expanded:
            raise OutlookError("the newest message would not expand; only its preview is available")
        return newest

    def search_dom(self, query: str, *, timeout: float = 30.0, settle: float = 2.0) -> list[Row]:
        """Drive Outlook's own search box. Prefer `search()`, which needs no UI and no wake dance.

        **The results list renders empty until the wake sequence is re-asserted after the query runs** -
        no rows, no "no results" text, an empty container, indistinguishable from a query that matched
        nothing.  This polls with a wake between attempts, so an empty list here is a real empty result.

        Outlook answers with a relevance-ranked "Top results" set rather than everything that matches,
        so treat a full list as a sample.  The `$search` REST call does not have that limit.
        """
        box = "document.querySelector(%s)" % json.dumps(SEARCH_INPUT_SEL)
        if not self._click_element(box, into_view=False, x_frac=0.3):
            raise OutlookError("no search box in this tab")
        time.sleep(0.8)
        # Select what is there first: Outlook keeps the last query and typing would append to it.
        self._js("(()=>{const e=document.querySelector(%s);e.focus();"
                 "e.setSelectionRange(0,e.value.length);return 1})()" % json.dumps(SEARCH_INPUT_SEL))
        self.s.type_text(query)
        time.sleep(0.8)
        self.s.key("Enter", "Enter", 13)

        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(settle)
            self.wake()          # the load-bearing call; see above
            time.sleep(0.6)
            if self._js("(()=>{const l=document.querySelector(%s);return l?"
                        "l.querySelectorAll('div[role=option]').length:0})()" % json.dumps(LIST_SEL)):
                return self._read_rows_dom(in_search=True)
        return []

    def in_search(self) -> bool:
        return bool(self._js("[...document.querySelectorAll('button,[role=button]')]"
                             ".some(b=>b.getAttribute('aria-label')==='Exit search')"))

    def exit_search(self, timeout: float = 15.0) -> bool:
        """Leave search mode and go back to the folder. Call it, or the tab stays in search."""
        if not self._click_element("[...document.querySelectorAll('button,[role=button]')]"
                                   ".find(b=>b.getAttribute('aria-label')==='Exit search')",
                                   into_view=False):
            return not self.in_search()
        ok = self.s.wait_for("![...document.querySelectorAll('button,[role=button]')]"
                             ".some(b=>b.getAttribute('aria-label')==='Exit search')", timeout=timeout)
        self.wake()
        return ok
