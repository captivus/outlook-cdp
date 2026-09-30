# AGENT-COMMS — outlook-cdp library

**Append only.** Never edit or delete an existing entry. Add new entries at the **bottom**, each opening
with a heading in exactly this form:

    ## YYYY-MM-DD HH:MM CDT — <agent name> — <one-line subject>

If you get something wrong, append a correction; do not rewrite history.

> **This repository is PUBLIC.** Keep every entry here **project-neutral and free of personal data**: no
> mailbox addresses, no real message subjects or senders, no third-party names, no customer or business
> detail. Describe browser behaviour, APIs and test outcomes only. Operational specifics — which
> chrome-agent instance, which mailbox, which messages — belong in the calling project's own private
> notes, not here. Refer to test material generically ("a two-message thread", "a 21 MB spreadsheet").

---

## 2026-09-30 13:45 CDT — Claude (lead) — Opening brief: grow this from send-only into a full read/write library

The task is to take this library from "compose and send" to everything needed for **reading, writing and
interacting** with Outlook on the web, built by several sub-agents coordinating through this file.

### What exists today

v0.1.1. 1,027 lines across seven modules, and **all of it is the send path**:

- `session.py` — `Session`, one-shot CDP per call, `js`, `click`, `key`, `type_text`, `focus_emulation`,
  `screenshot`, `wait_for`, host resolution across `outlook.cloud.microsoft` / `outlook.office.com`
- `mail.py` — `Mailbox`: `open_new`, `open_reply`, `add_recipients`, `set_subject`, `set_body`, `snapshot`,
  `send`, `discard`
- `guards.py` — pre-send checks; `recipients.py`; `dom.py` (closed-shadow-root piercing); `text.py`

**There is no read capability at all.** No listing, no message bodies, no attachments, no folders, no search.

### The three walls — these are the acceptance tests

Found while reading a live mailbox by hand-rolled CDP. If the finished library cannot do these, it is not
done.

1. **A restored tab renders nothing.** After a chrome-agent session restore, the message list is empty —
   `document.body.innerText.length` around 109 — until `Emulation.setFocusEmulationEnabled {"enabled":true}`
   is followed by `Page.setWebLifecycleState {"state":"active"}`. Then the rows appear. `focus_emulation()`
   exists but nothing sets the lifecycle state, and nothing does it automatically.
2. **Conversation grouping makes the reading pane serve the wrong message.** Opening a two-message thread
   returned the **older** message — the one we had sent — rather than the newer reply. Reply *metadata*
   (sender, date, attachment filenames) is readable from the row `aria-label` and from small text blocks in
   the pane; the reply **body** is not. Approaches already tried and **failed**: clicking the row and taking
   the largest text div; clicking the row's `Expand conversation` button (the aria-label does flip to
   "Expanded", but no child rows appear as `div[role=option]`); clicking the expanded child labelled with
   the sender, which has a **zero-size bounding box**; navigating to
   `/mail/<mailbox>/inbox/id/<urlencoded data-convid>`. This is the most important thing to fix.
3. **Attachments cannot be downloaded.** The tile is `div[role=option]` with
   `aria-label="<filename> Open <size>"` and exactly **one** child button, `More actions`. A trusted click
   on it reports success and **opens no menu** — enumerating `[role=menuitem]` afterwards returns only the
   page toolbar. `Page.setDownloadBehavior {"behavior":"allow","downloadPath":...}` was set beforehand and
   nothing arrived there or in `~/Downloads`. Real attachments to retrieve run up to **21 MB**, so design
   for streaming or chunking to disk and verify byte size, sha256 and magic bytes rather than trusting that
   a file appeared.

Also already known and worth not rediscovering: `Input.dispatchMouseEvent` produces trusted clicks where
synthetic `element.click()` silently no-ops; `DOM.getDocument {"depth":-1,"pierce":true}` traverses closed
shadow roots; one-shot CDP calls **cannot** intercept `Network` events because they detach immediately, so
anything needing events wants `chrome-agent attach <inst> +Event` or the Python CDP API; and a search URL of
the form `/mail/<mailbox>/search/id/<q>?q=<q>` did **not** filter — it returned the unfiltered inbox.

### Hard rules for every sub-agent

- **Send nothing. Draft nothing. Reply to nothing.** The operator of the mailbox under test has outbound
  email locked and has said explicitly not to respond to anything in it. Build and verify the **read**
  paths against messages already present. Any prior permission to self-test by sending is **suspended**. If
  you believe a test genuinely requires sending, stop and say so here.
- **Do not take focus.** CDP input reaches a background tab; never activate a tab or window, and leave
  whatever blank tab is in front where it is.
- **Opening a message marks it read** — a real change to someone's mailbox. Prefer metadata-only
  verification where it will do, and record here (generically) what you opened.
- **Stay project-neutral.** This library is shared across projects and agents; caller-specific rules live in
  the caller. Keep the existing public API working — other code imports `Mailbox` and `guards`.
- **Do not commit retrieved mail, attachments or any personal data to this repo.** Download test material
  to a gitignored scratch path.
- **Do not bump the version, tag, or publish.** The owner releases.
- Comments explain *why*, especially where a browser behaviour is counter-intuitive. The README's
  documented-behaviours list is part of the deliverable.

### The split

Claimed by Claude's sub-agents:

- **A — read path.** Folder listing (sender, subject, date, unread, attachment presence, ids); open a
  message; reliable body extraction **including a threaded reply**; working search. Wall 2 is the hard part.
- **B — attachments.** Enumerate and download, including the 21 MB case. First approach worth trying, having
  failed with the UI: enable the `Network` domain, click the tile's `Open` once, capture the request URL
  Outlook issues, then `fetch(url, {credentials:'include'})` **inside the page** and stream the bytes out —
  the tab already holds the session, so no token handling and no UI.

Offered to a second agent's sub-agents:

- **C — session and navigation robustness.** Fold the wake sequence into `Session` so a restored tab works
  without the caller knowing; folder navigation (Inbox/Sent/Drafts/Archive); conversation-view awareness,
  including whether grouping can be disabled per session rather than worked around.
- **D — adversarial review of A and B**, plus tests.

To avoid three-way edits, A and B each own a new module and **neither edits `session.py`**; session-level
helpers they need are noted here for C to consolidate. Claim, decline or re-cut in this file.

## 2026-09-30 13:27 CDT — Sub-agent B (attachments) — Claiming B; and a correction to an entry that is no longer in history

Claiming **B — attachments**. Delivered as `outlook_cdp/attachments.py`; `session.py` untouched.

**Correction first.** My opening entry, committed at 13:15 CDT before the public-repo warning was added,
named a county in one sentence describing the test material. That commit is no longer in this branch's
history — it was replaced when the lead rewrote this file — so nothing needs cleaning on my account beyond
what the lead has already said they will squash. Flagging it anyway rather than leaving it unsaid. Nothing
else I have written to this repo carries a mailbox address, a sender, a filename, an organisation or a
message subject; I checked the whole working tree and my own commits with a grep before committing.

## 2026-09-30 13:27 CDT — Sub-agent B (attachments) — Wall 3 solved: the UI was never the problem, the viewport was

**Wall 3 is down, and the first finding is not about attachments at all.**

The `More actions` button does open its menu, with *Preview / Copy / Download* in it. The reason a trusted
click on it "reported success and opened no menu" is that **the tile was outside the viewport**. The reading
pane scrolls independently of the page, so on a message with a normal amount of text the tile sits at
y ≈ 2100 in a 1321-px-tall viewport. `Input.dispatchMouseEvent` at that coordinate is accepted and lands on
nothing. `dom.box_center` cannot catch it either — the box is a healthy 32x46, it is simply off-screen.

**This is a general defect in how this library clicks, not an attachment quirk.** Every coordinate click in
`mail.py` and `dom.py` has the same hole. `dom.box_center` should scroll the node into view and re-measure
before returning a centre, or refuse when the centre is outside `innerWidth`/`innerHeight` — right now it
cheerfully returns a coordinate that cannot be clicked. Recommending that to whoever owns `dom.py`; I have
not changed it, since A is working in that area.

But the UI is still the wrong path, for three further reasons:

- The menu's `Download` produces a **`blob:` URL** — Outlook fetches the bytes itself and triggers an anchor
  download — so CDP sees `Page.downloadWillBegin` on a blob it cannot read.
- `Page.setDownloadBehavior` set from a **one-shot** connection dies with that connection, so by the time
  the click happens the path is set for nobody. That is the second half of "nothing arrived".
- Reaching the tile means opening the message, which **marks it read**.

## 2026-09-30 13:27 CDT — Sub-agent B (attachments) — Finished: `attachments.py`, no UI, no clicks, nothing marked read

`outlook_cdp/attachments.py` enumerates and downloads attachments through Outlook's own REST API from
inside the page. No clicking, no tiles, no download behaviour, and **it does not mark anything read** —
verified against a message that was unread before its attachment was downloaded in full and unread after.

**The lead's suggested `fetch(url, {credentials:'include'})` does not work, and the reason is worth
knowing: Outlook on the web is bearer-token authenticated, not cookie authenticated, even same-origin.**
That fetch returns **401**. The tokens are in the page's MSAL cache in `localStorage` under keys containing
`|accesstoken|`, one entry per scope; `access_token()` matches on the entry's scope list rather than its
key, because MSAL's key format has changed before, and it raises a specific error when the token has
expired (Outlook refreshes them itself while the tab lives).

**The second trap cost more time than the first: `/api/v2.0/me/` is the signed-in user, not the mailbox on
screen.** Those are different accounts whenever a shared mailbox is open. Reading `me` returned 154
messages, two of them with attachments, none of them the ones plainly visible in the UI — which reads
exactly like an API that does not hold them. Everything must be scoped to `/api/v2.0/users/<mailbox>/`,
with the mailbox from `Session.mailbox_identity()`.

**For A, and this may be your wall 2:** the `/id/<...>` segment of the reading pane's URL is a
**ConversationId**, not a message id — feeding it to `/messages/<id>` answers `404 ... doesn't belong to the
targeted mailbox`, which reads like a permissions problem and is a type error. `$filter=ConversationId eq
'<it>'` returns every message in the thread, each individually addressable with its own `Body`,
`From`, `ReceivedDateTime` and `IsRead`, regardless of which member the pane is painting. On the
two-message thread that defeated the DOM approach it returns both messages cleanly. It needs no clicking
and marks nothing read. `attachments.open_conversation_messages()` already does this and you are welcome to
lift it; the token and REST helpers underneath it are the parts that belong in `Session` (see below).

Other API behaviour worth not rediscovering:

- `$filter=HasAttachments eq true` **cannot** be combined with `$orderby=ReceivedDateTime desc` — the server
  answers `InefficientFilter`. Sort server-side, filter locally.
- An attachment's `Size` is **not** its content length: it counts the MIME envelope, running from a few
  hundred bytes to a third over the file (e.g. 624,607 reported for 624,077 bytes; 22,409,319 for
  22,408,839). Asserting equality against it fails on every correct download.
- `/messages/<id>/attachments/<aid>/$value` serves the decoded bytes from the page's **own** origin, so no
  CORS preflight and only one token. The `attachments.office.net` URL the Download button uses is
  cross-origin and needs a second, narrower `OwaAttachments.Read` token.

**Moving large files out of the page.** The whole file is fetched inside the page and parked on a page
global — the page outlives chrome-agent's per-command connections — then copied out in 3 MiB base64 slices
and appended to disk. Returning it from one `Runtime.evaluate` would put a 28 MB string through a
subprocess pipe for a 21 MB attachment. Inside the page, `String.fromCharCode.apply` is fed 32 KiB windows;
spreading a multi-megabyte array into an argument list overflows the call stack and the RangeError points
nowhere near the cause. A 21 MB spreadsheet takes about 2.7 s end to end.

**Verification, as asked.** Seven attachments retrieved live across five messages — a 21 MB spreadsheet, a
3 MB spreadsheet, three PDFs on one message, and two single PDFs — each checked for content length,
sha256, and magic bytes for its declared type. Both spreadsheets also open as valid zips with
`xl/workbook.xml` present and a clean `testzip()`, and every PDF ends in `%%EOF`. `download()` **raises**
rather than keeping a file whose first bytes are HTML or JSON, because that is what an expired token or the
wrong mailbox actually puts on disk: a plausible-looking file of plausible-looking length. Attachment names
come from the sender, so they are sanitised against `..` and separators before being joined to a path, and
a name already taken gets a `-2` suffix rather than overwriting. Files went to a gitignored scratch path;
`.gitignore` covers it and none of them is committed.

**For C, session-level helpers I put in my own module rather than editing `session.py`:** `access_token()`,
`mailbox_base()` and `rest_get()` all belong on `Session` — every read path in this library will want an
authenticated same-origin REST call scoped to the right mailbox, and `open_conversation_id()` belongs
wherever navigation lands. Fold them in whenever you touch `Session`; `attachments.py` should then import
them rather than define them. Also: whatever wakes a restored tab should be aware that re-activating the
web lifecycle appears to make Outlook re-select the top row of the list, which opens that message and marks
it read. That happened once during my run.

**Mailbox changes I caused:** exactly one message moved from unread to read, and not by a download — it was
selected by Outlook itself when the tab's lifecycle was reactivated, before I had worked out that the API
made the UI unnecessary. Nothing was sent, drafted or replied to; no tab or window was activated; the blank
tab is still in front. 23 new browser-free tests (56 in the suite), wired into the CI test command.
README behaviours 19–28 document all of the above.

---

## 2026-09-30 13:44 CDT — Sub-agent A (read path) — Wall 2 solved twice: `outlook_cdp/reading.py` is done

`outlook_cdp/reading.py` gives a caller folder listings, message bodies including threaded replies, and
search. Verified against a live mailbox of 99 inbox items in 88 threads. **Wall 2 is fixed.** Noting for
the record that I solved it in the DOM *before* the attachments agent's REST finding reached me, so there
are now two independent paths and both work; REST is the one callers should use.

### The REST path (primary — nothing is clicked, nothing is marked read)

Built on `attachments.py`'s `access_token`, `mailbox_base`, `rest_get` and `open_conversation_id` —
imported, not duplicated, as asked.

`Reader(session)` gives: `folders()`, `list_folder(folder, top, group)`, `unread()`,
`conversation(conversation_id)`, `open_conversation()`, `message(id)`,
`newest_in_conversation(conversation_id)`, `reply_to(conversation_id)`, `search(query, folder)`.
Everything returns dataclasses — `Row`, `Message`, `Conversation` — never CDP or OData payloads.

Three things worth adding to what the attachments agent found:

* **`UniqueBody`** is selectable alongside `Body`, and it is the message's own content with the quoted
  trail already stripped. That is exactly the split the DOM has to do by hand, for free. `Message.body`
  is `UniqueBody`, `Message.full_body` is `Body`.
* **`$search` cannot be combined with `$orderby`** — the server answers `SearchWithOrderBy` — so results
  are sorted client-side. `$search` works mailbox-wide and per-folder, returns individual messages, and
  has none of the UI's "Top results" truncation.
* **`ParentFolderId` plus `mailfolders`** gives every message a folder name, which is how "their reply"
  is told from "our own copy of what we sent". `reply_to()` uses it with the mailbox identity to return
  the newest message *not* from this mailbox — which stays right when the thread later grows another
  message of ours, and `newest` alone does not.

A caveat on counts: REST lists **messages**, the UI lists **threads**. `list_folder(group=True)` groups
on `ConversationId` to match the UI, and a thread can legitimately contain two copies of one outgoing
message (the Sent Items copy and an inbox copy where the mailbox was bcc'd), so a thread's message count
over REST can exceed what the reading pane shows. `Message.folder` disambiguates them.

### The DOM path (fallback, kept because it is the only thing that works without a token)

Same operations, suffixed `_dom`. These click, so they **mark messages read**. What the pane actually
offers, all verified:

* `div[aria-label="N messages"]` wraps the thread and is the authoritative length — present even for one
  message, where it reads `"1 messages"`.
* An expanded message is `div[aria-label="Email message"]`; a collapsed one is a plain sibling wrapper
  with a preview and **no body in the DOM at all**, and a trusted click renders it in place.
* Expanded messages carry `MSG_<id>_FROM`, `_TO`, `_CC`, `_BCC`, `_DATETIME`, `_SUBJECT`; the thread's
  subject is `CONV_<id>_SUBJECT`. These ids are the only stable anchors — every class in the pane is a
  build hash.
* **A freshly opened thread expands only its newest message.** A short thread can come up fully
  expanded, which is why "the biggest text block" returned the oldest message and looked like a success.

### What I could not make work, stated plainly

* **A DOM folder listing cannot be paged past one rendered screenful** — in the test mailbox, 23 rows of
  88. `scrollTop` is inert and reads back 0; every ancestor of the listbox has
  `scrollHeight == clientHeight`, so the scroller is *inside* it; `lastRow.scrollIntoView()` adds no
  rows; and `Input.dispatchMouseEvent {"type":"mouseWheel"}` — the trusted scroll that should work —
  **never returns over a one-shot CDP connection** and times the call out at 120s. Arrow keys would page
  it but move the selection, which marks messages read, so they are not available to a read path. The
  first row's `aria-setsize` gives the true total, so the shortfall is at least detectable. REST has no
  such limit, so I stopped here rather than going after an attached-connection scroll.
* **I did not disable conversation grouping.** It was my first plan and I dropped it: it is a persistent
  change to someone's mailbox settings, and REST makes it unnecessary. If C wants it anyway, the quick
  toggle is the folder header's filter/sort menu ("Show as: Messages"), not a URL parameter — no URL
  parameter for it turned up.
* **`(N)` on a row is a hint, not a count.** It appeared on two threads and was absent from seven others
  that demonstrably were threads, vanished from both as soon as they were read, and on one said 3 where
  the pane said 4. `Collapsed` in a row's aria-label is likewise not a thread signal — single messages
  carry it. `Expand conversation` being offered is the only dependable one.

### Changes outside my own module

* **`dom.py`** — `box_center` now takes `in_viewport=True` and refuses a centre outside the viewport, and
  there is a new `center_in_viewport()` and `scroll_into_view_center()`. This is the defect the
  attachments agent traced: the reading pane scrolls independently, so an element can have a healthy
  32x46 box at y=2100 in a 1321px viewport, and `Input.dispatchMouseEvent` accepts it, reports success
  and lands on nothing. `click_first_laid_out` now tries the box, then scrolls and **re-measures**.
  **Session/navigation agent: this is done, do not redo it.**
* **`__init__.py`** — exports `Reader`, `Row`, `Message`, `Conversation`, `html_to_text` and the parsing
  helpers; existing exports unchanged.
* **CI workflow** — `tests/test_reading.py` appended to the offline test line.
* **`session.py` — untouched**, as agreed. `Reader.wake()` duplicates focus emulation *plus*
  `Page.setWebLifecycleState {"state":"active"}` and re-asserts it after every navigation, because once
  per session is not enough: a search in a backgrounded tab returns an empty list — no rows, no "no
  results" text — until both calls land again, and that reads as a query that matched nothing.
  **This belongs in `Session`, with navigation re-asserting it.** Please fold it in and delete mine.

### Tests and state left behind

28 new offline tests in `tests/test_reading.py`, house runner, 84 across the suite, no browser needed.
They pin the things that were actually misread: `Name<addr>` with no space, "on behalf of", a display
name containing a comma, `"1 messages"`, private-use icon glyphs, `&nbsp;`, and the incomplete-thread
guard.

Mailbox state: I opened **three** threads through the UI while building the DOM path, so those are now
read that were unread. Everything after that went through REST and changed nothing. I sent nothing,
drafted nothing, replied to nothing, never activated a tab, and ran no git network operation.

---

## 2026-09-30 13:50 CDT — Sub-agent A (read path) — Addendum: attachment filenames, and a privacy check

Two things after the entry above.

**Attachment filenames on the REST path.** `Message.attachment_names` was empty: `HasAttachments` rides
along with the message but the filenames need one request each. `conversation()` now names them by
default and takes `attachment_names=False` to skip the pass; `message()` defaults the other way, since a
caller asking for one message by id usually already knows what it wants. Inline signature images are
excluded, so a message whose only attachments are a sender's logos returns an empty list with
`has_attachments` still true — correct, and worth knowing before treating the pair as contradictory.
Verified on both attachment cases, including the large-spreadsheet one.

**Privacy.** Checked my own commits and files against the standing warning: nothing I have written to
this repo contains a mailbox address, a real sender or subject, a third-party or organisation name, or
the chrome-agent instance name. The two `example.gov` strings a grep turns up in `README.md` and
`tests/test_guards.py` are pre-existing placeholders. Test material is described generically throughout
("a two-message thread", "one real mailbox", "the test mailbox"), and the one numeric detail I kept —
23 rows of 88 — is a scroller limit with nothing identifying in it. No mail, bodies or attachments were
written anywhere in the repo. I have run no git network operation at any point.
