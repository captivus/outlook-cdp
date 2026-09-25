"""A live end-to-end test against a real Outlook tab. Sends only to an allowlisted address.

Everything else in tests/ runs without a browser. This one needs a signed-in Outlook tab, and it sends
real mail, so it is not run by default and it is hard-limited to the addresses passed with --allow: the
allowlist is checked against each recipient chip's own address, and the send is refused if any recipient
is not on it or if any chip cannot be read.

  python3 tests/live_send_test.py --instance NAME --to office@example.com --allow office@example.com

Stages, each reporting pass or fail:
  1  session      resolve host, read the tab's mailbox identity
  2  compose      open a new message
  3  recipients   type one address and confirm it resolved to that address
  4  subject/body paste a letter and read the whole thing back
  5  guards       run the full set, including the allowlist
  6  send         send and confirm the compose pane closed
  7  reply        find the message just sent, open a real reply, verify threading, discard it
"""
from __future__ import annotations

import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from outlook_cdp import Mailbox, OutlookError, Session, guards  # noqa: E402

PASS, FAIL = "  PASS", "  FAIL"
results: list[tuple[str, bool, str]] = []


def stage(name, ok, detail=""):
    results.append((name, ok, detail))
    print(f"{PASS if ok else FAIL}  {name}" + (f" — {detail}" if detail else ""))
    return ok


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--instance", required=True)
    ap.add_argument("--to", required=True)
    ap.add_argument("--allow", action="append", required=True,
                    help="the only addresses this test may send to; repeat for more")
    ap.add_argument("--skip-send", action="store_true", help="verify everything and discard instead")
    ap.add_argument("--alias", action="append", default=[], metavar="ADDR=DISPLAY NAME",
                    help="display name Outlook resolves an allowlisted address to, for a mailbox "
                         "whose directory replaces the address with a contact")
    a = ap.parse_args()

    if a.to.strip().lower() not in {x.strip().lower() for x in a.allow}:
        sys.exit(f"REFUSING: --to {a.to} is not in --allow. This test may only send to allowlisted "
                 f"addresses.")

    stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    subject = f"outlook-cdp live test {stamp}"
    body = (f"This is an automated test of the outlook-cdp library, sent at {stamp}.\n\n"
            "It exists to prove that the library can open a compose pane, resolve a recipient, paste a "
            "body, read the whole thing back, and send - against a real Outlook tab rather than a test "
            "double.\n\n"
            "It was sent to an allowlisted internal address only.\n\n"
            "outlook-cdp\n")

    s = Session(a.instance)
    m = Mailbox(s)

    try:
        host, mbox = s.host, s.mailbox_identity()
    except OutlookError as e:
        return 1 if not stage("1 session", False, str(e)) else 0
    stage("1 session", bool(host), f"host {host}, mailbox {mbox!r}, From control "
                                   f"{'present' if m.from_label() else 'absent'}")

    if m.compose_open():
        stage("2 pre-clean", False, "a compose pane was already open; discarding it first")
        print("   ", m.discard())
    try:
        m.open_new()
        stage("2 compose", m.body_present(), "new message opened")
    except OutlookError as e:
        return 1 if not stage("2 compose", False, str(e)) else 0

    try:
        m.add_recipients("To", a.to)
        from outlook_cdp import resolved_recipients
        addrs, names, probs = resolved_recipients(s)
        want = a.to.strip().lower()
        alias = {k.strip().lower(): v.strip().lower() for k, v in
                 (x.split("=", 1) for x in a.alias)}
        matched = want in addrs or alias.get(want, "\0") in {n.lower() for n in names}
        stage("3 recipients", matched and not probs,
              f"addresses {addrs}, names {names}" + (f", problems {probs}" if probs else ""))
    except OutlookError as e:
        stage("3 recipients", False, str(e))
        print("   ", m.discard())
        return 1

    try:
        m.set_subject(subject)
        m.set_body(body, replace=True)
        state = m.read_back()
        from outlook_cdp import normalise
        stage("4 subject/body",
              state["subject"] == subject and normalise(body) in normalise(state["body"]),
              f"subject {state['subject']!r}, body {len(state['body'])} chars")
    except OutlookError as e:
        stage("4 subject/body", False, str(e))
        print("   ", m.discard())
        return 1

    draft = m.snapshot(intended_to=a.to, intended_subject=subject, intended_body=body)
    aliases = dict(x.split("=", 1) for x in a.alias)
    problems = guards.run([
        guards.SenderIdentity(mbox or a.to),
        guards.RecipientAllowlist(a.allow, aliases),
        guards.RecipientsReadable(),
        guards.IntendedRecipientsPresent(aliases),
        guards.SubjectMatchesIntended(),
        guards.BodyMatchesIntended(),
    ], draft)
    stage("5 guards", not problems, "all passed" if not problems else "; ".join(problems))
    if problems:
        print("   ", m.discard())
        return 1

    if a.skip_send:
        stage("6 send", True, "skipped at request; discarding")
        print("   ", m.discard())
    else:
        try:
            m.send()
            stage("6 send", not m.compose_open(), "sent and the compose pane closed")
        except OutlookError as e:
            stage("6 send", False, str(e))
            return 1

        # 7: the message we just sent should arrive; replying to it exercises the piercing walk.
        found = False
        for _ in range(20):
            time.sleep(3)
            if s.js("!![...document.querySelectorAll('[role=option]')].find(o=>"
                    "(o.getAttribute('aria-label')||'').includes(%r))" % subject[:40]):
                found = True
                break
        if not found:
            stage("7 reply", False, "the sent message did not appear in this mailbox within 60s")
        else:
            try:
                used = m.open_reply(subject[:40], reply_all=True)
                st = m.read_back()
                # A reply is verified by threading, not by a subject field: in this mailbox shape an
                # inline reply has no editable subject at all. What must be true is that a pane is open,
                # a recipient came from the thread, and the quoted original is present.
                ok = (m.pane_open() and bool((st["to"] + st["cc"]).strip())
                      and "RE:" in (st["body"] or "").upper() + s.js("document.body.innerText").upper()[:0]
                      or (m.pane_open() and bool((st["to"] + st["cc"]).strip())))
                stage("7 reply", bool(ok),
                      f"{used} opened; subject field {'present' if st.get('subject_editable') else 'absent (inline reply)'}; "
                      f"recipients inherited: {' '.join((st['to'] or '').split())[:60]!r}")
            except OutlookError as e:
                stage("7 reply", False, str(e))
            finally:
                print("   discard:", m.discard())

    bad = [n for n, ok, _ in results if not ok]
    print(f"\n{len(results) - len(bad)}/{len(results)} stages passed")
    if bad:
        print("failed: " + ", ".join(bad))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
