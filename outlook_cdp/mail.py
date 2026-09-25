"""Composing, verifying and sending one message in Outlook on the web.

The order matters and is not arbitrary:

1. open the pane (new, or a real reply),
2. fill it,
3. **read it all back**,
4. run the guards,
5. only then send, and confirm the pane closed.

Step 3 is the point of the whole module.  Every failure worth preventing here was a case where the page
accepted an action and did something other than what was asked: autocomplete swapped a recipient, a paste
landed on top of an old draft, a signature survived, the From identity was not the shared mailbox.
"""
from __future__ import annotations

import json
import time

from . import guards as _guards
from .dom import center_of, click_first_laid_out
from .recipients import chip_addresses, resolved_recipients, well_text
from .session import OutlookError, Session
from .text import to_html

#: The compose body, and the `contenteditable` is not decoration. The **reading** pane carries the same
#: `aria-label="Message body"`, so the bare selector is true whenever a message is merely selected: it
#: made "is a compose pane open" true while nothing was being composed, and it would let set_body target
#: the message you are reading. Only the compose surface is editable.
BODY_SEL = 'div[aria-label="Message body"][contenteditable="true"]'
SUBJECT_SEL = 'input[aria-label="Subject"]'


class Mailbox:
    """A high-level handle on one Outlook tab."""

    def __init__(self, session: Session):
        self.s = session
        self.s.focus_emulation()

    # ------------------------------------------------------------------ state

    def pane_open(self) -> bool:
        """Is any compose surface open - a new message or an inline reply?

        Tested on the **editable** message body, which exists in both and nowhere else.

        Two wrong answers were tried first. Testing the Subject input reports "nothing open" during an
        inline reply, because in a single-mailbox session a reply has no editable subject at all - and
        `send()` confirmed a send by that very absence, so it would have reported success without sending.
        Testing the body without `contenteditable` reports "something open" whenever a message is merely
        selected, because the reading pane uses the same aria-label.
        """
        return bool(self.s.js(f"!!document.querySelector('{BODY_SEL}')"))

    def subject_editable(self) -> bool:
        """Is there an editable Subject field? True for a new message, false for an inline reply."""
        return bool(self.s.js(f"!!document.querySelector('{SUBJECT_SEL}')"))

    # Kept as the old name, now meaning what it always should have: any compose surface.
    def compose_open(self) -> bool:
        return self.pane_open()

    def body_present(self) -> bool:
        return self.pane_open()

    def from_label(self) -> str | None:
        return self.s.js(
            "(()=>{const b=[...document.querySelectorAll('button')]"
            ".find(b=>/^From: /.test(b.getAttribute('aria-label')||''));"
            "return b?b.getAttribute('aria-label'):null})()")

    # ------------------------------------------------------------------ opening

    def open_new(self, timeout: float = 10.0) -> None:
        if self.compose_open():
            raise OutlookError("a compose pane is already open; close or discard it first")
        xy = center_of(self.s, "[...document.querySelectorAll('button')]"
                               ".find(b=>b.getAttribute('aria-label')==='New mail')")
        if not xy:
            raise OutlookError("New mail button not found")
        self.s.click(*xy)
        if not self.s.wait_for(f"!!document.querySelector('{BODY_SEL}')", timeout=timeout):
            raise OutlookError("compose pane did not open")
        time.sleep(1.0)

    def open_reply(self, pattern: str, reply_all: bool = True, timeout: float = 12.0) -> str:
        """Select the message whose aria-label matches `pattern` and click Outlook's own Reply control.

        Composing a fresh message with `RE:` typed into the subject is **not** a reply: it carries no
        In-Reply-To or References header so it does not thread, and its recipients are whatever you
        retype rather than whoever wrote to you.  That difference is why this method exists.
        """
        if self.compose_open():
            raise OutlookError("a compose pane is already open; close or discard it first")
        finder = ("[...document.querySelectorAll('[role=option]')]"
                  ".find(o=>new RegExp(%s,'i').test(o.getAttribute('aria-label')||''))"
                  % json.dumps(pattern))
        if not self.s.js(f"(()=>{{const o={finder};if(!o)return false;"
                         f"o.scrollIntoView({{block:'center'}});return true}})()"):
            raise OutlookError(f"no message matches {pattern!r}")
        time.sleep(1.0)
        xy = self.s.js(f"(()=>{{const o={finder};const r=o.getBoundingClientRect();"
                       f"return [r.x+r.width/2,r.y+r.height/2]}})()")
        self.s.click(*xy)
        time.sleep(2.5)
        labels = ["Reply all", "Reply"] if reply_all else ["Reply"]
        used = click_first_laid_out(self.s, labels, f"!!document.querySelector('{BODY_SEL}')", timeout)
        if not used:
            raise OutlookError("found no laid-out Reply control on the selected message "
                               "(the controls are in closed shadow roots; see dom.pierce_find)")
        return used

    # ------------------------------------------------------------------ filling

    def reveal_bcc(self) -> None:
        xy = center_of(self.s, "[...document.querySelectorAll('button,[role=button],span,div')]"
                               ".find(e=>e.children.length===0&&(e.innerText||'').trim()==='Bcc')")
        if xy:
            self.s.click(*xy)
            time.sleep(0.8)
        if not self.s.js('!!document.querySelector(\'div[aria-label="Bcc"]\')'):
            raise OutlookError("could not open the Bcc field")

    def add_recipients(self, well: str, addresses: str, attempts: int = 12) -> None:
        """Type each address and press Enter until Outlook turns it into a chip carrying that address.

        Two checks, because one is not enough: the chip count must rise, **and** the intended address
        must be present in the well.

        The presence check reads the **chips**, not the well's text. Outlook resolves an address it
        recognises - your own mailbox, anyone in the directory - to a display name, so the address never
        appears in the visible text. Checking the text there rejects a perfectly good recipient: this
        method can refuse `office@example.com` because the well reads `Example Office added.`
        The chip's `aria-label` carries the address in every case, which is why recipients.py exists, and
        it is the right instrument here too. The text is kept only as a fallback for markup changes.
        """
        for one in [x.strip() for x in (addresses or "").split(",") if x.strip()]:
            before = well_text(self.s, well).count("added")
            xy = center_of(self.s, f'document.querySelector(\'div[aria-label="{well}"]\')')
            if not xy:
                raise OutlookError(f"{well} field not found")
            self.s.click(*xy)
            time.sleep(0.4)
            self.s.type_text(one)
            for _ in range(attempts):
                time.sleep(0.7)
                self.s.key("Enter", "Enter", 13)
                time.sleep(0.5)
                if well_text(self.s, well).count("added") > before:
                    break
            else:
                raise OutlookError(f"{one} did not resolve into a {well} recipient")
            read = chip_addresses(self.s, well) or {}
            got = {a.lower() for a in read.get("addresses", [])}
            names = read.get("display_names") or []
            if one.lower() in got or one.lower() in well_text(self.s, well).lower():
                continue
            # Outlook resolved it against the directory, so the address is nowhere in the page. A chip
            # appeared and carries a name: that is a resolved recipient, not a failure. The caller
            # verifies WHICH one at the guard stage, where it knows the expected display name.
            if names:
                continue
            raise OutlookError(
                f"{well} reported a recipient was added but no chip carries {one} or any name. "
                f"Well text {well_text(self.s, well)[:160]!r}")

    def set_subject(self, subject: str) -> None:
        xy = center_of(self.s, f"document.querySelector('{SUBJECT_SEL}')")
        if not xy:
            raise OutlookError("Subject field not found")
        self.s.click(*xy)
        time.sleep(0.3)
        focused = f"document.activeElement===document.querySelector('{SUBJECT_SEL}')"
        if not self.s.js(focused):
            self.s.js(f"document.querySelector('{SUBJECT_SEL}').focus()")
        if not self.s.js(focused):
            raise OutlookError("could not focus Subject")
        self.s.type_text(subject)

    def set_body(self, text: str, replace: bool = True, drop_signature: bool = False) -> None:
        """Paste `text` as HTML.

        `replace=True` selects everything and deletes it first — right for a new message, wrong for a
        reply, where it would destroy the quoted thread.  In reply mode the cursor is collapsed to the
        top instead, which leaves Outlook's auto-inserted signature sitting between the new text and the
        quote; `drop_signature` removes that one node (`div#Signature`) rather than wiping the body.
        """
        body_html = to_html(text)
        self.s.js("""(()=>{const el=document.querySelector('%s');el.focus();
          const r=document.createRange();r.selectNodeContents(el);
          const s=getSelection();s.removeAllRanges();s.addRange(r);
          if(%s) document.execCommand('delete');
          else {const c=document.createRange();c.selectNodeContents(el);c.collapse(true);
                s.removeAllRanges();s.addRange(c);}
          const dt=new DataTransfer();dt.setData('text/html',%s);
          el.dispatchEvent(new ClipboardEvent('paste',{clipboardData:dt,bubbles:true,cancelable:true}));
          return true})()""" % (BODY_SEL, "true" if replace else "false", json.dumps(body_html)))
        time.sleep(1.0)
        if drop_signature:
            if self.s.js("(()=>{const n=document.querySelector('div#Signature');"
                         "if(!n)return 0;n.remove();return 1})()"):
                time.sleep(0.4)

    # ------------------------------------------------------------------ reading

    def read_back(self) -> dict:
        return self.s.js("""(()=>{const b=document.querySelector('%s');
          const well=f=>{const e=document.querySelector('div[aria-label="'+f+'"]');if(!e)return '';
            const c=e.closest('[role="presentation"]')||e.parentElement.parentElement;
            return (c.innerText||'')};
          const si=document.querySelector('%s');
          return {subject:si?(si.value||''):'', subject_editable:!!si, body:b?b.innerText:'',
                  to:well('To'), cc:well('Cc'), bcc:well('Bcc')}})()""" % (BODY_SEL, SUBJECT_SEL))

    def snapshot(self, *, intended_to="", intended_cc="", intended_bcc="", intended_subject="",
                 intended_body="", is_reply=False) -> _guards.Draft:
        """Everything the guards are allowed to see, read from the live pane."""
        state = self.read_back()
        addresses, names, problems = resolved_recipients(self.s)
        return _guards.Draft(
            from_label=self.from_label(), mailbox_identity=self.s.mailbox_identity(),
            subject=state["subject"], subject_editable=bool(state.get("subject_editable")),
            body=state["body"],
            chip_addresses=addresses, chip_display_names=names, chip_problems=problems,
            well_text=" ".join((state["to"], state["cc"], state["bcc"])),
            intended_to=intended_to, intended_cc=intended_cc, intended_bcc=intended_bcc,
            intended_subject=intended_subject, intended_body=intended_body, is_reply=is_reply)

    # ------------------------------------------------------------------ finishing

    def send(self, timeout: float = 10.0) -> None:
        """Click Send and require the compose pane to disappear.

        A clicked Send that leaves the pane open has not sent. Reporting success on the click alone is how
        a message gets sent twice.
        """
        xy = center_of(self.s, "[...document.querySelectorAll('button')]"
                               ".find(b=>b.getAttribute('aria-label')==='Send')")
        if not xy:
            raise OutlookError("Send button not found")
        self.s.click(*xy)
        # Confirm on the BODY disappearing. The Subject input is absent throughout an inline reply, so
        # waiting for it to vanish would succeed instantly and report a send that never happened.
        if not self.s.wait_for(f"!document.querySelector('{BODY_SEL}')", timeout=timeout):
            raise OutlookError("Send was clicked but the compose pane is still open; "
                               "check the mailbox before retrying, it may have sent")

    def discard(self, timeout: float = 20.0) -> str:
        """Discard the open compose pane, confirming with the keyboard if Outlook asks.

        Two behaviours, both verified live:

        * A draft with no typed content is discarded straight away with no confirmation.
        * A draft with content raises a modal reading "Discard message / Are you sure you want to
          discard this draft?" with OK and Cancel.

        **A coordinate click on that OK does nothing.** The button reports a real box - 96x32 at
        (935, 690) - and a truthy `offsetParent`, and a trusted mouse click at its centre leaves the
        dialog open. Pressing **Enter** dismisses it immediately, OK being the default button. The
        original script this library came from clicked the coordinates, which is why a draft with content
        was never actually discarded and one sat open in a real mailbox for two days.

        The order is therefore: Enter first, then a coordinate click, then a JS click, so a change in any
        one of them still leaves a working path.
        """
        xy = center_of(self.s, "[...document.querySelectorAll('button')]"
                               ".find(b=>b.getAttribute('aria-label')==='Discard')")
        if not xy:
            return "no compose open"
        # Re-assert focus emulation immediately before the click: asserting it once when the Mailbox was
        # constructed is not enough, and without it the same click left the pane open for twenty seconds.
        self.s.focus_emulation()
        self.s.click(*xy)

        ok_js = ("[...document.querySelectorAll('[role=dialog] button')]"
                 ".find(b=>b.innerText.trim()==='OK')")
        deadline = time.time() + timeout
        tried: list[str] = []
        while time.time() < deadline:
            if not self.compose_open():
                return "discarded" + (f" after {tried[-1]}" if tried else " with no confirmation")
            if self.s.js("!!document.querySelector('[role=dialog]')"):
                self.s.focus_emulation()
                if "Enter" not in tried:
                    self.s.key("Enter", "Enter", 13)
                    tried.append("Enter")
                elif "a coordinate click" not in tried:
                    hit = center_of(self.s, ok_js)
                    if hit:
                        self.s.click(*hit)
                    tried.append("a coordinate click")
                elif "a JS click" not in tried:
                    self.s.js(f"(()=>{{const e={ok_js};if(e){{e.click();return 1}}return 0}})()")
                    tried.append("a JS click")
            time.sleep(0.5)
        return ("STILL OPEN after trying " + ", ".join(tried)) if tried else \
               "STILL OPEN and no confirm dialog appeared"
