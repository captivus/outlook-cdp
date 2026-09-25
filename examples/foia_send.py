"""A configuration template showing how to combine project-specific guards.

This is the shape of the script outlook-cdp was extracted from: one shared mailbox that must be the
sender, a kill-switch file a human controls, a list of addresses that have bounced, and a signature the
body must and must not contain.

Run it with --send only when the lock is lifted. Without --send it composes, verifies and leaves the
draft open, which is what you want while a human is still reviewing the letter.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from outlook_cdp import Mailbox, Session, guards  # noqa: E402

INSTANCE = "example-instance"  # replace with your chrome-agent instance
FROM_LABEL = "From: office@example.com"  # replace with your mailbox
PROJECT = os.path.expanduser("~/mail-project")  # replace with your project directory
SEND_LOCK = os.path.join(PROJECT, "data/foia/SEND_LOCK")
BLOCKLIST = os.path.join(PROJECT, "data/foia/known_bad_addresses.csv")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--to", default="")
    ap.add_argument("--cc", default="")
    ap.add_argument("--subject", default="")
    ap.add_argument("--body-file", required=True)
    ap.add_argument("--reply-to", default="")
    ap.add_argument("--send", action="store_true")
    a = ap.parse_args()
    text = open(a.body_file, encoding="utf-8").read()

    # Checked before a draft is even opened: a locked mailbox should not produce a draft to be tempted by.
    if a.send:
        locked = guards.SendLock(SEND_LOCK)(guards.Draft())
        if locked:
            print("REFUSING: " + "; ".join(locked), file=sys.stderr)
            return 2

    m = Mailbox(Session(INSTANCE))
    if a.reply_to:
        print("opened:", m.open_reply(a.reply_to))
    else:
        m.open_new()
        m.add_recipients("To", a.to)
        m.set_subject(a.subject)
    if a.cc:
        m.add_recipients("Cc", a.cc)
    m.set_body(text, replace=not a.reply_to, drop_signature=bool(a.reply_to))

    draft = m.snapshot(intended_to=a.to, intended_cc=a.cc, intended_subject=a.subject,
                       intended_body=text, is_reply=bool(a.reply_to))
    problems = guards.run([
        guards.SendLock(SEND_LOCK),
        guards.FromAddress(FROM_LABEL),
        guards.Blocklist(BLOCKLIST),
        guards.RecipientsReadable(),
        guards.IntendedRecipientsPresent({"office@example.com": "Example Office"}),
        guards.SubjectMatchesIntended(),
        guards.BodyMatchesIntended(),
        # The auto-inserted personal signature must not go out over the shared mailbox's name.
        guards.BodyMustNotContain(["personal@example.com"]),
        guards.BodyMustContain(["office@example.com"]),
    ], draft)

    for p in problems:
        print("  " + p, file=sys.stderr)
    if problems:
        print("NOT SENT: draft left open for inspection.", file=sys.stderr)
        return 4
    if not a.send:
        print("dry run: verified, draft left open")
        return 0
    m.send()
    print("sent")
    return 0


if __name__ == "__main__":
    sys.exit(main())
