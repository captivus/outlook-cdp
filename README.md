# outlook-cdp

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
| `outlook_cdp/guards.py` | `Draft` plus the composable checks |
| `outlook_cdp/text.py` | plain text to Outlook HTML, and whitespace-normalised comparison |
| `outlook_cdp/cli.py` | the command line |
| `examples/foia_send.py` | a configuration template: one shared mailbox, a send lock, a bounce list |

## Tests

```
python3 tests/test_text.py && python3 tests/test_guards.py
```

33 tests, no browser required. They cover the text conversion and every guard, including the two cases
that matter most: a blocked address hidden behind a display name, and a truncated paste.

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
