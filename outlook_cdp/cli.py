"""A general command line over the library, for agents that would rather shell out than import.

Every safety rule is passed in, so the CLI is as safe as the flags you give it and no safer.  It sends
only with --send, and only when every guard you asked for passed.
"""
from __future__ import annotations

import argparse
import json
import sys

from . import guards as G
from .mail import Mailbox
from .recipients import resolved_recipients
from .session import OutlookError, Session


def build_guards(a) -> list:
    gs: list = []
    if a.send_lock:
        gs.append(G.SendLock(a.send_lock))
    if a.sender:
        gs.append(G.SenderIdentity(a.sender, a.from_label or None))
    elif a.from_label:
        gs.append(G.FromAddress(a.from_label))
    if a.allow:
        gs.append(G.RecipientAllowlist(
            a.allow, dict(x.split("=", 1) for x in a.alias) if a.alias else None))
    if a.blocklist:
        gs.append(G.Blocklist(a.blocklist))
    gs.append(G.RecipientsReadable())
    gs.append(G.IntendedRecipientsPresent(
        dict(x.split("=", 1) for x in a.alias) if a.alias else None))
    gs.append(G.SubjectMatchesIntended())
    gs.append(G.BodyMatchesIntended())
    if a.body_must_not_contain:
        gs.append(G.BodyMustNotContain(a.body_must_not_contain))
    if a.body_must_contain:
        gs.append(G.BodyMustContain(a.body_must_contain))
    return gs


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="outlook-cdp", description=__doc__)
    p.add_argument("--instance", required=True, help="chrome-agent instance holding the Outlook tab")
    p.add_argument("--host", default=None, help="override the Outlook host instead of resolving it")
    sub = p.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("compose", help="compose a new message or a reply, verify it, optionally send")
    c.add_argument("--to", default=""); c.add_argument("--cc", default=""); c.add_argument("--bcc", default="")
    c.add_argument("--subject", default="")
    c.add_argument("--body-file", required=True)
    c.add_argument("--reply-to", default="", help="regex matching the message to reply to; recipients "
                                                 "and threading then come from Outlook, not from --to")
    c.add_argument("--send", action="store_true", help="send once every guard passes")
    c.add_argument("--shot", default="", help="write a screenshot here before deciding")
    c.add_argument("--send-lock", default="", help="path whose existence forbids sending")
    c.add_argument("--sender", default="", help="address this must be sent from; satisfied by the From "
                                                "control when there is one, else by the tab's mailbox")
    c.add_argument("--from-label", default="", help='exact From control text, if it differs from '
                                                   '"From: <sender>"')
    c.add_argument("--allow", action="append", default=[], metavar="ADDR",
                   help="the only addresses this send may go to; repeat for more")
    c.add_argument("--blocklist", default="", help="CSV of addresses that must never be written to")
    c.add_argument("--alias", action="append", default=[], metavar="ADDR=DISPLAY",
                   help="an address that legitimately resolves to a display name")
    c.add_argument("--body-must-not-contain", action="append", default=[], metavar="TEXT")
    c.add_argument("--body-must-contain", action="append", default=[], metavar="TEXT")
    c.add_argument("--drop-signature", action="store_true",
                   help="remove div#Signature after a reply paste")

    sub.add_parser("recipients", help="print the addresses the open draft's chips actually carry")
    sub.add_parser("discard", help="discard the open compose pane")
    sub.add_parser("status", help="report what the tab currently has open")

    a = p.parse_args(argv)
    s = Session(a.instance, host=a.host)
    m = Mailbox(s)

    if a.cmd == "status":
        print(json.dumps({"host": s.host, "compose_open": m.compose_open(),
                          "from": m.from_label()}, indent=1))
        return 0
    if a.cmd == "discard":
        print(m.discard()); return 0
    if a.cmd == "recipients":
        addrs, names, problems = resolved_recipients(s)
        print(json.dumps({"addresses": addrs, "display_names": names, "problems": problems}, indent=1))
        return 1 if problems else 0

    text = open(a.body_file, encoding="utf-8").read()
    # The lock is checked before anything is composed, so a locked mailbox never even opens a draft.
    if a.send and a.send_lock:
        pre = G.SendLock(a.send_lock)(G.Draft())
        if pre:
            print("REFUSING: " + "; ".join(pre), file=sys.stderr)
            return 2

    try:
        if a.reply_to:
            used = m.open_reply(a.reply_to)
            print(json.dumps({"opened": used, "matched": a.reply_to}))
        else:
            if not (a.to and a.subject):
                p.error("--to and --subject are required unless --reply-to is used")
            m.open_new()
        if a.bcc:
            m.reveal_bcc()
        if not a.reply_to and a.to:
            m.add_recipients("To", a.to)
        for well, val in (("Cc", a.cc), ("Bcc", a.bcc)):
            if val:
                m.add_recipients(well, val)
        if not a.reply_to:
            m.set_subject(a.subject)
        m.set_body(text, replace=not a.reply_to,
                   drop_signature=a.drop_signature or bool(a.reply_to))
    except OutlookError as e:
        print(f"ABORT (draft left open, NOT sent): {e}", file=sys.stderr)
        return 3

    draft = m.snapshot(intended_to=a.to, intended_cc=a.cc, intended_bcc=a.bcc,
                       intended_subject=a.subject, intended_body=text, is_reply=bool(a.reply_to))
    if a.shot:
        s.screenshot(a.shot)
    problems = G.run(build_guards(a), draft)
    print(json.dumps({"from": draft.from_label, "subject": draft.subject,
                      "recipients": draft.chip_addresses, "problems": problems}, indent=1))
    if problems:
        print("NOT SENT: the draft is left open for inspection.", file=sys.stderr)
        return 4
    if not a.send:
        print("dry run: verified, draft left open")
        return 0
    try:
        m.send()
    except OutlookError as e:
        print(f"UNCERTAIN: {e}", file=sys.stderr)
        return 5
    print(json.dumps({"sent": True}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
