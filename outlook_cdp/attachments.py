"""Enumerating and downloading message attachments, without driving the attachment UI.

Read this before changing anything here, because almost every obvious approach fails.

**The attachment UI is not a usable download path.**  An attachment renders as
``div[role=option]`` whose ``aria-label`` is ``"<name> Open <size>"``, containing exactly one button,
``More actions``.  Clicking that button opens a menu with ``Preview / Copy / Download`` - but only if
the tile is inside the viewport.  The reading pane scrolls independently and puts the tile at a
y-coordinate well past ``innerHeight`` on any message with a normal amount of text; a trusted
``Input.dispatchMouseEvent`` at that coordinate is accepted, returns success, and lands on nothing.
That is the whole explanation for "the click worked and no menu opened", and it is not specific to
attachments - it applies to any control this library clicks by coordinate.  ``dom.box_center`` cannot
catch it either, because the box is a perfectly healthy 32x46; it is simply off-screen.

Even with the menu open, ``Download`` is a dead end for automation: Outlook fetches the bytes itself,
wraps them in a ``blob:`` URL and triggers an anchor download, so what CDP sees is
``Page.downloadWillBegin`` on a blob it cannot read, and whether a file lands depends on download
behaviour that a one-shot CDP connection cannot keep set (the state dies with the connection).

**What does work: be the authenticated client.**  The tab already holds the session, so the bytes can
be fetched from inside the page.  Two things about that are counter-intuitive:

1. **Cookies are not enough.**  ``fetch(url, {credentials: 'include'})`` returns **401**.  Outlook on
   the web is bearer-token authenticated, not cookie authenticated, even same-origin.  The tokens are
   in the page's MSAL cache in ``localStorage``, keyed by scope; see `access_token`.
2. **``/me/`` is the wrong mailbox.**  The signed-in user and the mailbox on screen are often not the
   same account - a shared mailbox is the normal case for this library.  ``/api/v2.0/me/...`` reads the
   *signed-in user's* mail and will happily return a hundred unrelated messages and no attachments,
   which reads exactly like "the API does not have them".  Every call here is scoped to
   ``/api/v2.0/users/<mailbox>/``, taking the mailbox from `Session.mailbox_identity`.

Everything in this module is **read-only**.  The token it borrows carries write and send scopes
because that is what Outlook asked for, so no function here issues anything but ``GET``; keep it that
way.

Nothing here clicks, so nothing here marks a message read - which the attachment UI path cannot say.
"""
from __future__ import annotations

import base64
import binascii
import hashlib
import json
import os
import re
import time
import urllib.parse
from dataclasses import dataclass

from .session import OutlookError, Session

#: Scope fragment of the token that authenticates the Outlook REST API.  Matched as a substring
#: against the MSAL cache entry's ``target``, which is a space-separated scope list.
REST_SCOPE = "outlook.office.com/Mail.ReadWrite"

#: Scope fragment of the separate, narrower token Outlook mints for ``attachments.office.net``.  Not
#: needed by anything here - the REST ``$value`` endpoint serves the same bytes from the page's own
#: origin - but kept documented because the UI uses it and a network trace will show it.
ATTACHMENT_SCOPE = "OwaAttachments.Read"

#: Leading bytes that prove a download is the file type it claims.  An expired token, a wrong mailbox
#: or a sign-in redirect all answer with HTML or JSON, which lands on disk as a plausible-looking file
#: of plausible-looking length; only the first few bytes tell you apart a real record from an error
#: page.  Keys are the prefix of the MIME type Outlook reports.
MAGIC: dict[str, tuple[bytes, ...]] = {
    "application/vnd.openxmlformats-officedocument": (b"PK\x03\x04",),  # xlsx / docx / pptx are zips
    "application/zip": (b"PK\x03\x04",),
    "application/pdf": (b"%PDF-",),
    "image/png": (b"\x89PNG\r\n\x1a\n",),
    "image/jpeg": (b"\xff\xd8\xff",),
    "image/gif": (b"GIF87a", b"GIF89a"),
    "application/msword": (b"\xd0\xcf\x11\xe0",),        # legacy OLE container
    "application/vnd.ms-excel": (b"\xd0\xcf\x11\xe0",),
}

#: Where a one-shot `Runtime.evaluate` parks fetched bytes between calls.  chrome-agent opens a new
#: CDP connection per command, but the *page* is the same page, so a global outlives the connection.
#: That is what makes chunked reads possible at all.
_STASH = "window.__outlook_cdp_attachment__"

#: Default slice size for moving bytes out of the page.  Base64 inflates by 4/3, and the whole slice
#: crosses a subprocess pipe as one JSON string, so a 21 MB attachment at 3 MiB a slice is eight
#: round-trips of about 4 MB each rather than one 28 MB string that may not survive the buffer.
DEFAULT_CHUNK = 3 * 1024 * 1024


@dataclass(frozen=True)
class Attachment:
    """One attachment, as the REST API describes it."""

    message_id: str
    attachment_id: str
    name: str
    size: int          #: Outlook's own figure. See `download` - it is NOT the content length.
    content_type: str
    is_inline: bool

    @property
    def is_file(self) -> bool:
        """False for the images a sender's signature drags along, which are rarely wanted."""
        return not self.is_inline


# --------------------------------------------------------------------------- credentials

#: Pulls one access token out of the page's MSAL cache.  MSAL stores each token as a JSON blob in
#: localStorage under a key containing ``|accesstoken|``; the blob's ``target`` is the granted scope
#: list and ``secret`` is the JWT.  Matching on scope rather than on the key means this keeps working
#: across MSAL's cache-key format changes, which have happened before.
_TOKEN_JS = """(()=>{const want=%s;let best=null;
  for(const k of Object.keys(localStorage)){
    if(!/\\|accesstoken\\|/.test(k)) continue;
    let v; try{v=JSON.parse(localStorage.getItem(k))}catch(e){continue}
    if(!v||!v.secret||!(v.target||'').includes(want)) continue;
    const exp=parseInt(v.expiresOn||'0',10)||0;
    if(!best||exp>best.expires) best={secret:v.secret,expires:exp};
  }
  return best})()"""


def access_token(session: Session, scope: str = REST_SCOPE, min_ttl: float = 30.0) -> str:
    """The page's own access token for `scope`, read from its MSAL cache.

    There is no way to *mint* a token from outside the page, so this borrows one Outlook already has.
    They are short-lived - an hour is typical - and Outlook refreshes them silently as long as the tab
    is alive, so a token that is expired here means the tab has been idle rather than that anything is
    broken; touching the mailbox in the UI, or simply waiting for Outlook's own next poll, replaces it.
    """
    got = session.js(_TOKEN_JS % json.dumps(scope))
    if not got:
        raise OutlookError(
            f"no access token for scope {scope!r} in the page's MSAL cache. Either the tab is not "
            f"signed in, or Outlook has changed where it caches tokens.")
    ttl = int(got.get("expires") or 0) - time.time()
    if ttl < min_ttl:
        raise OutlookError(
            f"the page's {scope!r} token expired {-int(ttl)}s ago. Outlook refreshes it on its own "
            f"while the tab is alive; give it a moment, or interact with the mailbox, then retry.")
    return got["secret"]


def mailbox_base(session: Session) -> str:
    """The REST path prefix for the mailbox this tab is showing.

    Deliberately not ``/api/v2.0/me``.  See the module docstring: ``me`` is the signed-in account,
    which for a shared mailbox is somebody else's mail entirely.
    """
    mailbox = session.mailbox_identity()
    if not mailbox:
        raise OutlookError(
            "cannot tell which mailbox this tab is showing - Session.mailbox_identity() read nothing "
            "from the tab URL, so a REST call would silently read the signed-in user's mail instead.")
    return "/api/v2.0/users/" + urllib.parse.quote(mailbox, safe="")


# --------------------------------------------------------------------------- REST

#: Same-origin GET with the borrowed bearer token, run inside the page.  Returns the parsed body and
#: the status separately so a caller can tell "no attachments" from "the request was refused".
_REST_JS = """(async()=>{try{
  const r=await fetch(%s,{headers:{authorization:'Bearer '+%s,accept:'application/json'}});
  let body=null; try{body=await r.json()}catch(e){}
  return {status:r.status, body:body}}catch(e){return {status:0, error:String(e)}}})()"""


def rest_get(session: Session, path: str, token: str | None = None) -> dict:
    """GET a same-origin Outlook REST path and return the decoded JSON body.

    `path` must already be query-encoded; OData's ``$filter`` and friends need literal ``$`` and the
    values inside them need percent-encoding, so encoding the whole thing here would corrupt it.
    """
    token = token or access_token(session)
    out = session.js(_REST_JS % (json.dumps(path), json.dumps(token)), await_promise=True) or {}
    if out.get("status") != 200:
        body = out.get("body") or {}
        err = (body.get("error") or {}) if isinstance(body, dict) else {}
        raise OutlookError(
            f"REST GET {path.split('?')[0]} returned {out.get('status')} "
            f"{err.get('code', '')} {err.get('message', out.get('error', ''))}"[:400])
    return out.get("body") or {}


def messages_with_attachments(session: Session, folder: str = "inbox", top: int = 50,
                              token: str | None = None) -> list[dict]:
    """Recent messages in `folder` that carry at least one attachment, newest first.

    The ``HasAttachments`` flag is filtered here in Python on purpose.  Asking the server for
    ``$filter=HasAttachments eq true`` *together with* ``$orderby=ReceivedDateTime desc`` is rejected
    with ``InefficientFilter``, and dropping the sort to keep the filter returns the mailbox in
    arrival order oldest-first, which is the opposite of what any caller wants.
    """
    token = token or access_token(session)
    path = (f"{mailbox_base(session)}/mailfolders/{urllib.parse.quote(folder)}/messages"
            f"?$top={int(top)}&$orderby=ReceivedDateTime%20desc"
            f"&$select=Id,Subject,From,ReceivedDateTime,HasAttachments,IsRead,ConversationId")
    rows = rest_get(session, path, token).get("value") or []
    return [r for r in rows if r.get("HasAttachments")]


def list_attachments(session: Session, message_id: str, token: str | None = None,
                     include_inline: bool = False) -> list[Attachment]:
    """Every attachment on one message.

    Inline images are excluded by default: a signature block routinely contributes three or four
    logos, and a caller asking for "the attachments" almost never means those.
    """
    token = token or access_token(session)
    path = (f"{mailbox_base(session)}/messages/{urllib.parse.quote(message_id, safe='')}/attachments"
            f"?$select=Id,Name,Size,ContentType,IsInline")
    out = []
    for row in rest_get(session, path, token).get("value") or []:
        att = Attachment(message_id=message_id, attachment_id=row.get("Id") or "",
                         name=row.get("Name") or "", size=int(row.get("Size") or 0),
                         content_type=row.get("ContentType") or "",
                         is_inline=bool(row.get("IsInline")))
        if include_inline or att.is_file:
            out.append(att)
    return out


# --------------------------------------------------------------------------- the open message

def open_conversation_id(session: Session) -> str | None:
    """The id in the reading pane's URL - which is a **conversation** id, not a message id.

    Worth being explicit about, because it looks like a message id and is not one: feeding it to
    ``/messages/<id>`` answers ``404 ... doesn't belong to the targeted mailbox``, which reads like a
    permissions problem and is really a type error.  Outlook groups by conversation, so the pane
    addresses the thread; ``ConversationId eq '<it>'`` is what turns it back into messages.
    """
    m = re.search(r"/id/([^/?#]+)", session.tab_url())
    return urllib.parse.unquote(m.group(1)) if m else None


def open_conversation_messages(session: Session, token: str | None = None) -> list[dict]:
    """Every message in the thread the reading pane is showing, newest first.

    This is the way past conversation grouping: the pane may be rendering the wrong member of the
    thread, but the thread's messages are all addressable here regardless of what is painted.
    """
    conv = open_conversation_id(session)
    if not conv:
        raise OutlookError("no message is open in the reading pane (the tab URL carries no /id/ part)")
    token = token or access_token(session)
    path = (f"{mailbox_base(session)}/messages"
            f"?$top=50&$select=Id,Subject,From,ReceivedDateTime,HasAttachments,IsRead"
            f"&$filter=ConversationId%20eq%20'{urllib.parse.quote(conv, safe='')}'")
    rows = rest_get(session, path, token).get("value") or []
    return sorted(rows, key=lambda r: r.get("ReceivedDateTime") or "", reverse=True)


def open_message_attachments(session: Session, token: str | None = None,
                             include_inline: bool = False) -> list[Attachment]:
    """Attachments on the open thread, across every message in it.

    Across the whole thread rather than one message because the reading pane does not reliably tell
    you which member it is showing, and an attachment that arrived on the reply is the one the caller
    came for.
    """
    token = token or access_token(session)
    found: list[Attachment] = []
    for msg in open_conversation_messages(session, token):
        if msg.get("HasAttachments"):
            found += list_attachments(session, msg["Id"], token, include_inline)
    return found


def tile_labels(session: Session) -> list[str]:
    """The attachment tiles the reading pane is painting, read from their aria-labels.

    A cross-check on the REST answer, and nothing more: it sees only what is rendered, it cannot see
    an attachment on a collapsed thread member, and the label's size is Outlook's rounded display
    string (``"21 MB"``).  It clicks nothing, so it marks nothing read.
    """
    labels = session.js(
        "[...document.querySelectorAll('[role=option]')]"
        ".map(o=>o.getAttribute('aria-label')||'').filter(l=>/ Open \\d+(\\.\\d+)? ?[KMG]?B/.test(l))")
    return labels or []


# --------------------------------------------------------------------------- bytes

#: Fetch the attachment's raw bytes and park them on the page.  `$value` on the REST attachment
#: resource serves the decoded file from the page's **own** origin, so no CORS preflight and no second
#: token - unlike the `attachments.office.net` URL the Download button uses, which is cross-origin and
#: needs the separate OwaAttachments.Read token.
_FETCH_JS = """(async()=>{try{
  const r=await fetch(%s,{headers:{authorization:'Bearer '+%s}});
  if(!r.ok) return {status:r.status};
  const b=await r.arrayBuffer();
  %s={bytes:new Uint8Array(b), type:r.headers.get('content-type')||''};
  return {status:r.status, length:b.byteLength, type:%s.type}}
 catch(e){return {status:0, error:String(e)}}})()"""

#: Base64 one window of the parked bytes.  `String.fromCharCode.apply` is fed 32 KiB at a time because
#: spreading a multi-megabyte array into an argument list overflows the call stack - the failure is a
#: RangeError from deep inside the runtime, with nothing to point at the real cause.
_SLICE_JS = """(()=>{const s=%s; if(!s||!s.bytes) return null;
  const a=s.bytes, from=%d, to=Math.min(%d, a.length); let out='';
  for(let i=from;i<to;i+=0x8000) out+=String.fromCharCode.apply(null, a.subarray(i, Math.min(i+0x8000, to)));
  return btoa(out)})()"""


def magic_ok(head: bytes, content_type: str) -> bool | None:
    """Do these leading bytes match the declared type?  None when the type has no known signature."""
    for prefix, signatures in MAGIC.items():
        if content_type.startswith(prefix):
            return head.startswith(signatures)
    return None


def looks_like_an_error_page(head: bytes) -> bool:
    """Is this HTML or JSON where a file was expected?

    The failure this catches is a download that succeeded, weighs a sensible number of kilobytes and
    is a sign-in page.  Checked independently of `magic_ok` so it also fires for types with no
    signature of their own, such as CSV.
    """
    start = head[:512].lstrip()
    return start[:1] in (b"<", b"{") or b"<html" in start.lower()


def verify_file(path: str, content_type: str = "", expected_bytes: int | None = None) -> dict:
    """Size, sha256 and a magic-byte check on a downloaded file.

    `expected_bytes` is the *content length*, not Outlook's ``Size``.  Do not compare against
    ``Attachment.size``: that figure is the attachment's size **as stored in the message**, including
    its MIME envelope and base64 expansion, and it runs a few hundred bytes to a third larger than the
    file.  Asserting equality against it fails on every correct download.
    """
    digest = hashlib.sha256()
    total = 0
    head = b""
    with open(path, "rb") as fh:
        while True:
            block = fh.read(1 << 20)
            if not block:
                break
            if not head:
                head = block[:512]
            digest.update(block)
            total += len(block)
    result = {
        "path": path, "bytes": total, "sha256": digest.hexdigest(),
        "magic_ok": magic_ok(head, content_type),
        "error_page": looks_like_an_error_page(head),
    }
    if expected_bytes is not None:
        result["size_ok"] = (total == expected_bytes)
    return result


def _release(session: Session) -> None:
    try:
        session.js(f"(()=>{{{_STASH}=null;return 1}})()")
    except OutlookError:
        pass  # A freed stash is a courtesy to the tab, not a correctness requirement.


def download(session: Session, att: Attachment, dest_dir: str, *, filename: str | None = None,
             chunk_bytes: int = DEFAULT_CHUNK, token: str | None = None,
             verify: bool = True) -> dict:
    """Write one attachment to `dest_dir` and return its verification record.

    The bytes make two hops on purpose.  The page fetches the whole file into a typed array parked on
    a global, and this then copies it out in base64 slices, one CDP round-trip each, appending to the
    file as it goes.  Returning the whole thing from a single ``Runtime.evaluate`` is what the obvious
    implementation does, and it puts a 28 MB JSON string through a subprocess pipe for a 21 MB
    attachment.  Slicing also means a transport hiccup costs one slice, not the download.

    Raises rather than leaving a half-written file in place, because a truncated record that looks
    like a record is worse than no record.
    """
    token = token or access_token(session)
    safe = _safe_name(filename or att.name or att.attachment_id)
    os.makedirs(dest_dir, exist_ok=True)
    path = os.path.join(dest_dir, safe)

    url = (f"{mailbox_base(session)}/messages/{urllib.parse.quote(att.message_id, safe='')}"
           f"/attachments/{urllib.parse.quote(att.attachment_id, safe='')}/$value")
    got = session.js(_FETCH_JS % (json.dumps(url), json.dumps(token), _STASH, _STASH),
                     await_promise=True) or {}
    if got.get("status") != 200:
        raise OutlookError(f"fetching {att.name!r} returned {got.get('status')} "
                           f"{got.get('error', '')}".strip())
    length = int(got.get("length") or 0)

    try:
        written = 0
        partial = path + ".part"
        with open(partial, "wb") as fh:
            step = max(64 * 1024, int(chunk_bytes))
            while written < length:
                b64 = session.js(_SLICE_JS % (_STASH, written, written + step))
                if b64 is None:
                    raise OutlookError(
                        f"the page lost the parked bytes for {att.name!r} after {written} of "
                        f"{length} - the tab probably navigated mid-download")
                try:
                    block = base64.b64decode(b64)
                except (binascii.Error, ValueError) as exc:
                    raise OutlookError(f"slice at {written} of {att.name!r} did not decode: {exc}")
                if not block:
                    raise OutlookError(f"empty slice at {written} of {length} for {att.name!r}")
                fh.write(block)
                written += len(block)
        if written != length:
            raise OutlookError(f"{att.name!r}: wrote {written} bytes, the response was {length}")
        os.replace(partial, path)
    finally:
        _release(session)
        if os.path.exists(path + ".part"):
            os.unlink(path + ".part")

    if not verify:
        return {"path": path, "bytes": written, "sha256": None}
    record = verify_file(path, att.content_type, expected_bytes=length)
    record["name"] = att.name
    record["content_type"] = att.content_type
    record["reported_size"] = att.size
    if record["error_page"] or record["magic_ok"] is False:
        raise OutlookError(
            f"{att.name!r} downloaded {written} bytes that are not a {att.content_type} "
            f"(error_page={record['error_page']}). The file was kept at {path} to be looked at.")
    return record


def download_all(session: Session, attachments, dest_dir: str, **kw) -> list[dict]:
    """Download several attachments, one token for the lot.

    Names collide - three public bodies will all send you ``response.pdf`` - so a name already taken
    in `dest_dir` gets a ``-2``, ``-3`` suffix rather than overwriting what is there.
    """
    token = kw.pop("token", None) or access_token(session)
    records = []
    for att in attachments:
        records.append(download(session, att, dest_dir,
                                filename=_unique_name(dest_dir, att.name or att.attachment_id),
                                token=token, **kw))
    return records


def _safe_name(name: str) -> str:
    """A filename that cannot escape `dest_dir`.

    The name comes from whoever sent the message, so it is untrusted input: it can hold slashes, ``..``
    or a leading dash, and on this path it would be joined straight onto a directory.
    """
    name = os.path.basename((name or "").replace("\\", "/")).strip()
    name = re.sub(r'[\x00-\x1f<>:"|?*]', "_", name)
    name = name.lstrip(".-") or "attachment"
    return name[:180]


def _unique_name(dest_dir: str, name: str) -> str:
    stem, ext = os.path.splitext(_safe_name(name))
    candidate, n = stem + ext, 1
    while os.path.exists(os.path.join(dest_dir, candidate)):
        n += 1
        candidate = f"{stem}-{n}{ext}"
    return candidate
