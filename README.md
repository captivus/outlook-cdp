# outlook-cdp

![Outlook compose pane passing sender, recipient, and message checks before sending](https://raw.githubusercontent.com/captivus/outlook-cdp/main/assets/readme-banner.png)

Drive Outlook on the web through Chrome DevTools, with the pre-send checks as objects you compose
rather than rules baked into a script.

It exists because every one of the behaviours below cost a real mistake to learn, and the next agent
should inherit them instead of rediscovering them.

## Installation

```sh
pip install outlook-cdp
```

The `chrome-agent` command must also be installed and available on `PATH`.

```python
from outlook_cdp import Session, Mailbox, guards

s = Session("my-chrome-agent-instance")     # the instance holding the Outlook tab
m = Mailbox(s)

m.open_new()
m.add_recipients("To", "clerk@example.gov")
m.set_subject("Records request")
m.set_body(open("letter.txt").read())

draft = m.snapshot(intended_to="clerk@example.gov",
                  intended_subject="Records request",
                  intended_body=open("letter.txt").read())

problems = guards.run([
    guards.SendLock("SEND_LOCK"),                       # a file a human controls
    guards.FromAddress("From: office@example.com"),     # which identity may send
    guards.Blocklist("bounced.csv"),                    # addresses that must never be written to
    guards.RecipientsReadable(),                        # refuse if any recipient is unverifiable
    guards.IntendedRecipientsPresent(),                 # autocomplete did not swap anyone
    guards.SubjectMatchesIntended(),
    guards.BodyMatchesIntended(),                       # the WHOLE letter, not its opening words
    guards.BodyMustNotContain(["personal@example.com"]), # the auto signature did not survive
], draft)

if problems:
    raise SystemExit("NOT SENT: " + "; ".join(problems))
m.send()                                                # and confirm the pane actually closed
```

Or from a shell:

```
outlook-cdp --instance my-instance compose \
    --to clerk@example.gov --subject "Records request" --body-file letter.txt \
    --from-label "From: office@example.com" --send-lock SEND_LOCK --blocklist bounced.csv
# add --send to actually send; without it you get a verified draft left open
outlook-cdp --instance my-instance recipients     # what the open draft's chips really carry
outlook-cdp --instance my-instance discard
```

Reading is separate from sending, and needs no UI at all:

```python
from outlook_cdp import Session, attachments

s = Session("my-chrome-agent-instance")

# Every attachment on the thread the reading pane is showing, inline signature images excluded.
for att in attachments.open_message_attachments(s):
    print(att.name, att.content_type, att.size)

# Or sweep the folder, newest first, and write everything to disk.
for msg in attachments.messages_with_attachments(s, folder="inbox", top=50):
    got = attachments.download_all(s, attachments.list_attachments(s, msg["Id"]), "out/")
    for rec in got:
        print(rec["bytes"], rec["sha256"], rec["magic_ok"], rec["path"])
```

Listing, message bodies and search work the same way — over the mailbox's own REST API, from inside
the page, with no clicking and nothing marked read:

```python
from outlook_cdp import Session, Reader

r = Reader(Session("my-chrome-agent-instance"))

for row in r.list_folder("inbox", top=50):
    print(row.received, row.sender, row.subject, row.message_count, row.unread)

# The whole thread, oldest first, across every folder it touches.
conv = r.conversation(row.conversation_id)
for m in conv.messages:
    print(m.index, m.received, m.folder, m.sender_address, len(m.body))

# The newest message that did not come from this mailbox: "their reply", not our own last word.
reply = r.reply_to(row.conversation_id)
print(reply.sender, reply.received, reply.attachment_names)
print(reply.body)                  # this message only; reply.full_body keeps the quoted trail

for hit in r.search("some phrase"):          # individual messages, not collapsed threads
    print(hit.folder, hit.received, hit.sender, hit.subject)
```

Inbox rules are REST too — created, read back and checked, without the settings dialog:

```python
from outlook_cdp import Session, Rules, Rule

s = Session("my-chrome-agent-instance")
rules = Rules(s)                     # a tab on plain /mail/ names no mailbox: pass mailbox= explicitly

rule = Rule("Archive build notices", from_addresses=["ci@example.com"],
            subject_contains=["Build succeeded"], move_to="archive")
print(len(rules.preview(rule)))      # what it would have caught already; changes nothing
rules.create(rule)                   # refuses a duplicate name; deletes it again if the server altered it
rules.apply(rule)                    # optional: rules act only on new mail, this moves the existing matches
```

```
outlook-cdp --instance my-instance rules preview --mailbox me \
    --from ci@example.com --subject-contains "Build succeeded"
outlook-cdp --instance my-instance rules add --mailbox me --name "Archive build notices" \
    --from ci@example.com --subject-contains "Build succeeded" --move-to archive
outlook-cdp --instance my-instance rules list --mailbox me
outlook-cdp --instance my-instance rules apply --mailbox me --name "Archive build notices"
    # --name alone uses the rule exactly as stored, so what is applied is what is live

outlook-cdp --instance my-instance search "some phrase" --mailbox me --top 50
outlook-cdp --instance my-instance search --from ci@example.com --folder archive --top 1000 --summary
```

Contacts are REST as well — created, read back and checked, without the People UI:

```python
from outlook_cdp import Session, Contacts, Contact

book = Contacts(Session("my-chrome-agent-instance"), mailbox="me@example.com")
print([c.name() for c in book.find_by_phone("(555) 555-0100")])   # digits compared, formatting ignored
book.create(Contact(given_name="Pat", surname="Example", mobile_phone="+1 (555) 555-0100",
                    home_address={"street": "1 Main St", "city": "Springfield", "state": "IL"},
                    personal_notes="Met at the club."))   # refuses a number already on a contact

pat = book.find_by_phone("5555550100")[0]
book.update(pat.id, nickname="", middle_name="")       # "" clears the field; read back and checked
book.set_photo(pat.id, open("face.jpg", "rb").read())  # verified by reading the stored photo's size
```

```
outlook-cdp --instance my-instance contacts find --mailbox me --phone 5555550100
outlook-cdp --instance my-instance contacts add --mailbox me --json people.json   # one object or a list
outlook-cdp --instance my-instance contacts update --mailbox me --id ID --json changes.json  # "" clears
outlook-cdp --instance my-instance contacts photo --mailbox me --id ID --photo face.jpg
```

`Reader(session, mailbox=...)` takes the mailbox explicitly for a tab whose URL names none.

`Reader` also has a `*_dom` method for each of those, which drives the UI instead. They exist for a page
with no usable token and they **mark what they open read**, so the REST calls are the ones to use.

`download` refuses rather than leaving a file that is not the type it claims: it checks the response
length, the magic bytes for the declared content type, and whether the first bytes are HTML or JSON —
which is what an expired token or the wrong mailbox actually puts on disk. It clicks nothing, so it does
not mark anything read.

## What it needs

`chrome-agent` on the PATH, a running instance with Outlook on the web signed in, and Python 3.10+.
No third-party packages.

## The browser behaviours it encodes

These are the whole value of the library. Each is a thing Outlook does that makes naive automation
quietly wrong rather than loudly broken.

1. **A JS `element.click()` silently does nothing** on several Outlook controls, which check
   `event.isTrusted`. Input must be dispatched through CDP's Input domain. Symptom: the click "works"
   and nothing happens.
2. **Reply / Reply all / Forward live in closed shadow roots.** `document.querySelectorAll` cannot see
   them, so any selector work reports the button does not exist. `DOM.getDocument` with `pierce: true`
   walks through, and its `backendNodeId`s stay valid across chrome-agent's separate connections.
3. **A node can exist and have no box** — collapsed, hidden, in an unrendered pane. Clicking a
   zero-area box does nothing. Try every candidate with the same label and take the first laid-out one.
4. **Composing a fresh message with `RE:` in the subject is not a reply.** It carries no `In-Reply-To`
   or `References` header, so it does not thread, and its recipients are whoever you retype rather than
   whoever wrote to you. This put one real answer in a generic intake queue instead of in front of the
   officer who asked the question.
5. **Recipient chips show a display name; the address is in the chip's `aria-label`** as
   `Display Name <a@b.com>`. A blocklist that greps the well's visible text will happily pass a draft
   addressed to exactly the address you blocked, because the well shows `County FOIA Office`.
6. **Outlook's autocomplete substitutes cached contacts**, and the recipient count rising is not
   evidence it rose to the address you typed. Require the address itself to appear.
7. **In reply mode the body must not be wiped** — that destroys the quoted thread — so the auto-inserted
   personal signature survives between your text and the quote. Remove the one node, `div#Signature`.
8. **A background tab stalls on dialogs and pane transitions.**
   `Emulation.setFocusEmulationEnabled` plus `Page.setWebLifecycleState: active` unstalls it without
   raising the window, so the human keeps whatever they had in front.
9. **The host moved.** Outlook on the web has been served from `outlook.office.com` and
   `outlook.cloud.microsoft`, and which one a session lands on varies. Resolve it at runtime from the
   live target list; hardcoding it broke the first version of this code.
10. **Set the body by pasting `text/html` through a synthetic ClipboardEvent.** Typing a multi-line
    string inserts breaks Outlook then reflows; assigning `innerHTML` loses the undo stack and sometimes
    the send button's dirty-state tracking.
11. **Verify the whole body, not its first words.** A prefix check passes on a body that was truncated,
    pasted twice, or left sitting on top of an older draft.
12. **A clicked Send that leaves the pane open has not sent.** Reporting success on the click alone is
    how a message goes out twice. Wait for the pane to disappear, and if it does not, say the outcome is
    uncertain rather than guessing either way.

13. **A single-mailbox session has no From control at all.** Outlook renders the From chooser only when
    there is more than one identity to choose from. A guard that demands one refuses every send in that
    shape. `SenderIdentity` accepts the From label when it exists and otherwise the mailbox in the tab's
    own URL (`/mail/office@example.com/`), and **fails closed** when it can establish neither.
14. **A directory-resolved recipient has its address nowhere in the page.** An external address keeps it
    in the chip's `aria-label` as `Display Name <a@b.com>`. Your own mailbox, or a colleague, is replaced
    by a contact whose label reads `unknownExample Office` — and searching every attribute of
    every descendant for the domain returns nothing. So the chip reader returns `display_names` as well as
    `addresses`, and a caller that must act on an internal recipient supplies the display name it expects.
15. **An inline reply has no editable Subject field.** Not empty — absent. Any check that requires one
    refuses every reply.
16. **The reading pane shares `aria-label="Message body"` with the composer.** Only the compose surface is
    `contenteditable="true"`, so that is the only safe test for "is something being composed", and the only
    safe target for setting a body. Without it, "a pane is open" is true whenever a message is selected.
17. **The discard confirmation ignores a coordinate click.** The dialog reads *Discard message / Are you
    sure you want to discard this draft?* with OK and Cancel. OK reports a real 96x32 box and a truthy
    `offsetParent`, and a trusted mouse click at its centre does nothing. **Enter** dismisses it, OK being
    the default; a JS `.click()` also works. The original script clicked the coordinates, which is why a
    draft with content was never really discarded and one sat open in a live mailbox for two days.
18. **Focus emulation must be re-asserted immediately before a click**, not once when the session is
    created. The same Discard click left the pane open for twenty seconds without it and closed in under
    one with it.

19. **A coordinate click outside the viewport is accepted and lands on nothing.** The reading pane scrolls
    independently of the page, so on any message with a normal amount of text an attachment tile sits at a
    y-coordinate well past `innerHeight`. `Input.dispatchMouseEvent` there returns success and does
    nothing, and `dom.box_center` cannot catch it because the box is a perfectly healthy 32x46 — it is
    simply off-screen. **This is not specific to attachments**: it applies to every control this library
    clicks by coordinate. Scroll the element into view and re-measure first. It is the whole explanation
    for "the click worked and no menu opened".
20. **An attachment tile carries exactly one button, `More actions`.** The menu behind it does hold
    *Preview / Copy / Download* — but only opens when the tile is on screen, per 19. There is no
    `Download` control in the light DOM to find.
21. **That Download is a `blob:` URL, so CDP cannot follow it.** Outlook fetches the bytes itself, wraps
    them in a blob and triggers an anchor download; what CDP observes is `Page.downloadWillBegin` on a
    blob it cannot read. Worse, `Page.setDownloadBehavior` set from a one-shot connection dies **with that
    connection**, so the path is set for nobody by the time the click happens. Driving the UI is not a
    usable way to get an attachment out.
22. **Cookies do not authenticate Outlook's own API.** `fetch(url, {credentials: 'include'})` against the
    page's own origin returns **401**. Outlook on the web is bearer-token authenticated throughout, and
    the tokens are in the page's MSAL cache in `localStorage`, one per scope, under keys containing
    `|accesstoken|`. `attachments.access_token` matches on the scope list rather than the key, because
    MSAL's key format has changed before.
23. **`/api/v2.0/me/` is the signed-in user, not the mailbox on screen.** Those are different accounts
    whenever a shared mailbox is open — the normal case here. Reading `me` returns a hundred unrelated
    messages and none of the attachments you can see, which reads exactly like an API that does not have
    them. Every call must be scoped to `/api/v2.0/users/<mailbox>/`, taking the mailbox from the tab URL.
24. **The reading pane's URL id is a *conversation* id, not a message id.** `/mail/<mailbox>/inbox/id/<x>`
    addresses the thread, because Outlook groups by conversation. Feeding it to `/messages/<id>` answers
    `404 … doesn't belong to the targeted mailbox`, which reads like a permissions problem and is a type
    error. `$filter=ConversationId eq '<x>'` turns it back into messages — which is also the way past
    conversation grouping generally, since the thread's members are all addressable even when the pane is
    painting the wrong one.
25. **An attachment's `Size` is not its content length.** It counts the MIME envelope and base64
    expansion, and runs a few hundred bytes to a third over the file: 624,607 reported for a 624,077-byte
    PDF, 22,409,319 for a 22,408,839-byte spreadsheet. Asserting equality against it fails on every
    correct download; verify against the response's own length.
26. **`$filter=HasAttachments eq true` and `$orderby=ReceivedDateTime desc` cannot be combined** — the
    server answers `InefficientFilter`. Dropping the sort to keep the filter returns the mailbox
    oldest-first, which is never what the caller wants, so sort server-side and filter locally.
27. **Reading an attachment over the API does not mark the message read.** Verified: an unread message
    still read unread after its attachment was downloaded in full. Opening the message in the UI to reach
    the tile *does* mark it read, which is a real change to someone's mailbox — a second reason not to
    drive the UI.
28. **Bytes have to leave the page in slices.** The whole file can be fetched inside the page in one go,
    but returning it from a single `Runtime.evaluate` puts a 28 MB base64 string through a subprocess pipe
    for a 21 MB attachment. Park it on a page global — the page survives chrome-agent's per-command
    connections — and copy it out a few MiB at a time. Inside the page, feed
    `String.fromCharCode.apply` 32 KiB windows; spreading a multi-megabyte array into an argument list
    overflows the call stack, and the RangeError points nowhere near the cause.

29. **The reading pane's `/id/<…>` URL segment is a ConversationId, not a message id.** It looks like
    a message id and is not one: `/messages/<it>` answers `404 … doesn't belong to the targeted mailbox`,
    which reads like a permissions problem and is really a type error. `$filter=ConversationId eq '<it>'`
    turns it back into individual messages, each with its own sender, timestamp and body.
30. **Conversation grouping makes the reading pane serve the wrong message.** The pane renders the whole
    thread, so taking the largest block of text returns the *oldest* message — for a reply to something
    you sent, that is your own outgoing mail, and it looks exactly like a success. Three things that do
    **not** fix it: the row's `Expand conversation` button, which flips the row's aria-label to
    "Expanded", adds no `div[role=option]` child rows and reveals only a zero-size element; clicking the
    expanded child labelled with the sender, which has no box to click; and navigating to the
    `data-convid` URL, which is behaviour 29. What fixes it is treating the pane as a list of messages
    and addressing them individually — or not using the pane at all.
31. **A freshly opened thread expands only its newest message.** The others are one-line stubs with
    sender, date, a text preview and **no body in the DOM at all**. A short thread, though, can come up
    fully expanded, so neither "the first block" nor "the biggest block" is the message you want. A
    trusted click on a stub renders it in place; the expansion that matters happens in the reading pane,
    not in the message list.
32. **The pane's per-message anchors are ids, not classes.** An expanded message is
    `div[aria-label="Email message"]`; inside it Outlook emits `MSG_<id>_FROM`, `_TO`, `_CC`, `_BCC`,
    `_DATETIME` and `_SUBJECT`, and the thread's subject is on `CONV_<id>_SUBJECT`. The thread is wrapped
    in `div[aria-label="N messages"]`, present even for a single message, where it reads `"1 messages"` —
    Outlook's grammar, not a typo. Every class name in the pane is a build hash and will change; these
    do not.
33. **`div[role=option]` is not only a message row.** Attachment tiles in the reading pane share the
    role, so an unscoped `querySelectorAll('div[role=option]')` mixes attachments into a folder listing.
    Scope row queries to `[role=listbox][aria-label^="Message list"]`.
34. **The wake sequence is not a once-per-session call.** `Emulation.setFocusEmulationEnabled` plus
    `Page.setWebLifecycleState {"state":"active"}` has to be re-asserted after every navigation. Run a
    search in a backgrounded tab and the results list is *genuinely* empty — no rows, no "no results"
    text, an empty container — until both land again. It is indistinguishable from a query that matched
    nothing, and it was read as one.
35. **The row `(N)` badge is a hint, not a thread length.** In one real folder it appeared on two threads
    and was absent from seven others that demonstrably were threads, it vanished from both as soon as
    they were read, and on one of the two it said 3 where the pane said 4. Threads are also counted
    across folders, so a row whose thread is one message here plus one in Sent Items shows no badge at
    all. `Expand conversation` being offered is the dependable "this is a thread" signal; a real count
    needs the thread itself. Nor does the aria-label help: **`Collapsed` appears on single messages
    too** — it describes the row widget, not the thread — and trusting it called nearly every row a
    thread.
36. **A DOM folder listing cannot be paged past one screenful.** Outlook's message list is a custom
    virtual scroller: assigning `scrollTop` does nothing and reads back as 0, every ancestor of the
    listbox reports `scrollHeight == clientHeight` so the element that scrolls is *inside* it,
    `lastRow.scrollIntoView()` adds no rows, and `Input.dispatchMouseEvent {"type":"mouseWheel"}` —
    the trusted scroll that would work — **never returns** over a one-shot CDP connection and times the
    call out. Arrow keys would page it but move the selection, which marks messages read. In one mailbox
    that ceiling was 23 rows of 88; the first row's `aria-setsize` gives the true total, so at least the
    shortfall is detectable. Listing over REST has no such limit.
37. **Outlook's search URL does not search.** `/mail/<mailbox>/search/id/<q>?q=<q>` renders the
    unfiltered folder, which looks like a query that matched everything. Driving the search box does
    work, and returns individual messages tagged with the folder each lives in, but only a
    relevance-ranked "Top results" set. `$search` over REST returns the lot — though it **cannot** be
    combined with `$orderby` (the server answers `SearchWithOrderBy`), so sort the results yourself.
38. **Query results are HTML, and highlighted terms fragment the DOM.** In search results Outlook wraps
    each matching term in `<mark>`, which splits one subject cell into several leaf nodes; reading leaves
    directly returned a subject of "Assessment", then "Roll", as though they were separate columns, and
    shifted every later column along by two. Bodies come back from REST as HTML unless you set a `Prefer`
    header, and they are full of `&nbsp;` — invisible, everywhere in mail pasted out of Word, and enough
    to make an otherwise identical body compare unequal.
39. **A healthy box can still be unclickable.** The reading pane scrolls independently of the page, so an
    element can have a 32x46 box at y=2100 in a 1321px viewport. `Input.dispatchMouseEvent` accepts those
    coordinates, reports success, and the click lands on nothing — the whole of the "the click ran and
    nothing happened" mystery. A box model cannot show it; compare the centre against `innerWidth` and
    `innerHeight`, and scroll and **re-measure** before clicking, because the scroll is animated and a
    rectangle taken in the same breath as the request to scroll describes where the element was.

40. **The cached REST token can write inbox rules.** It carries `MailboxSettings.ReadWrite`, so
    `/mailfolders/inbox/messagerules` takes a POST from inside the page with no consent prompt and no
    settings UI to drive.
41. **A rule needs a `Sequence`, and 0 is not one.** Omit it and the server answers
    `MessageRuleValidationError … Field: 'Sequence', Value: '0'`. `Rules.create` puts a new rule last.
42. **A stored rule is not proof of the rule you sent.** The server ignores condition keys it does not
    recognise rather than refusing them, and a rule left with no conditions matches every message. Read
    the rule back and compare; `Rules.create` deletes it again on any difference.
43. **The well-known `archive` folder need not be called "Archive".** In one mailbox it was "Archive1",
    with an ordinary, empty folder called "Archive" beside it. Moving by display name would have filed
    mail into the folder Outlook's Archive button never uses. Resolve `mailfolders/archive` to its id.
44. **Rules only act on arrivals.** Outlook's "run rule now" has no REST call; `Rules.apply` is the
    explicit stand-in and is never run implicitly.
45. **The tab URL does not always name the mailbox.** A tab on `/mail/` or a deep link carries none, so
    `mailbox_identity()` returns `None` and every REST helper refuses. For the signed-in user's own
    mailbox, `/api/v2.0/me`'s `EmailAddress` is the address to pass — never a silent fallback, because
    on a shared mailbox `me` is someone else.
46. **The cached REST token can write contacts.** It carries `Contacts.ReadWrite`, so `/contacts` takes a
    POST from inside the page with no consent prompt.
47. **A contact's phone and email slots are fixed and small**: one mobile, two home, two business numbers
    and three email addresses. `Contact.to_rest` refuses more than fit rather than trusting the server
    to say what it dropped; overflow belongs in `personal_notes`, which is the card's "Notes".
48. **Phone numbers are free text.** One contact book held the same number as `+1 (555) …`, `555.…` and
    with a non-breaking space inside it. Matching compares the last ten digits, and
    `Contacts.create` refuses a number already on another contact so a re-run does not duplicate people.
49. **A contact photo cannot go into the page in one call.** chrome-agent passes each CDP call on its
    command line, and Linux caps one argument at 128 KiB, so even a modest JPEG as base64 fails before
    it reaches Chrome. `Contacts.set_photo` parks 60 KB base64 slices on a page global, PUTs
    `/photo/$value` as raw bytes, and then reads `/photo` back for a width, because a 2xx on its own
    does not prove a picture was stored.
50. **An empty field in a PATCH is left alone, not cleared.** Clearing a nickname or middle name needs
    an explicit `null` (an empty list for phones), which `Contacts.update` sends for any field passed as
    `""` or `[]`, and then checks that the field really came back empty. The stored display name is not
    re-derived when its parts change, so dropping a middle initial means passing `display_name` too.

## Design rules

**Discover what the page offers; never assume one shape.** The same mailbox in two configurations differs
in whether there is a From chooser, whether a recipient resolves to an address or a contact, and whether a
reply has an editable subject. Every one of those was found by running the library against a real tab, and
each is handled by looking rather than by assuming. A feature that works in one shape and silently fails in
the other is worse than one that refuses.


**No safety rule is built in.** Which mailbox may send, which addresses are burnt, where the kill switch
lives — all of that is specific to whose mailbox it is, so it is passed in as guards. Adding a rule never
means editing the send path.

**Guards fail closed.** A blocklist file that is missing is not an empty blocklist; it means the check
could not be performed, and a guard that passes when it cannot see is worse than no guard. Same for a
recipient well that cannot be read.

**Nothing sends unless you ask twice** — `--send` on top of every guard passing. The default is a
verified draft left open for a human to look at.

## Layout

| File | What is in it |
|---|---|
| `outlook_cdp/session.py` | the CDP connection, trusted input, screenshots, waiting, host resolution |
| `outlook_cdp/dom.py` | the shadow-root piercing walk, box measurement, clicking the laid-out candidate |
| `outlook_cdp/recipients.py` | reading addresses out of recipient chips, and reporting what could not be read |
| `outlook_cdp/mail.py` | open / fill / read back / send, and discard |
| `outlook_cdp/attachments.py` | enumerating and downloading attachments, read-only, without the attachment UI |
| `outlook_cdp/reading.py` | folder listings, message bodies and search — over REST, with a DOM fallback |
| `outlook_cdp/rules.py` | inbox rules: list, preview, create with read-back verification, apply, delete |
| `outlook_cdp/contacts.py` | contacts: list, find by phone or name, create and update with read-back verification, photos, delete |
| `outlook_cdp/guards.py` | `Draft` plus the composable checks |
| `outlook_cdp/text.py` | plain text to Outlook HTML, and whitespace-normalised comparison |
| `outlook_cdp/cli.py` | the command line |
| `examples/foia_send.py` | a configuration template: one shared mailbox, a send lock, a bounce list |

## Tests

```
python3 tests/test_text.py && python3 tests/test_guards.py && \
    python3 tests/test_attachments.py && python3 tests/test_reading.py && python3 tests/test_rules.py && \
    python3 tests/test_contacts.py
```

105 tests, no browser required. They cover the text conversion, every guard, the read-path parsing, and the attachment logic
that decides where bytes land and whether they are believed — including the three cases that matter most:
a blocked address hidden behind a display name, a truncated paste, and an attachment filename that tries
to escape its destination directory.

**Verified end to end against a real mailbox.** `tests/live_send_test.py` opens a compose pane, resolves a
recipient, pastes a body, reads the whole thing back, runs every guard, **sends**, then finds the sent
message, opens a real threaded reply through the closed shadow roots, and discards it. All seven stages
pass. It sends only to addresses given with `--allow`, checked against each chip, and refuses if any
recipient cannot be identified:

```
python3 tests/live_send_test.py --instance NAME \
    --to office@example.com --allow office@example.com \
    --alias "office@example.com=Example Office"      # needed when the directory resolves the address
```

Add `--skip-send` to verify everything and discard instead of sending.

**Behaviours 13 to 18 above were all found by that test**, not by reading the code. Five of them were live
defects: a session with no From chooser refused every send; a directory-resolved recipient looked
unreadable; an inline reply looked like no pane at all, so `send()` would have reported success without
sending; the reading pane looked like an open composer; and discard silently did nothing on any draft with
content.

## Releasing

Update the version in `pyproject.toml`, then publish a GitHub Release with a matching `v` tag (for
example, `v0.1.0`). The release starts `.github/workflows/publish.yml`, which runs the offline tests,
builds the package, and publishes through PyPI Trusted Publishing. The GitHub `pypi` environment and
PyPI publisher must both be configured before the first release.

## Independence

This is an independent project and is not affiliated with, endorsed by, or sponsored by Microsoft.
Microsoft Outlook and its logo are trademarks of Microsoft Corporation.
