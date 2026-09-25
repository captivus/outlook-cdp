"""Drive Outlook on the web through Chrome DevTools, safely.

    from outlook_cdp import Session, Mailbox, guards

    s = Session("my-chrome-agent-instance")
    m = Mailbox(s)
    m.open_new()
    m.add_recipients("To", "clerk@example.gov")
    m.set_subject("Records request")
    m.set_body(open("letter.txt").read())

    draft = m.snapshot(intended_to="clerk@example.gov", intended_subject="Records request",
                       intended_body=open("letter.txt").read())
    problems = guards.run([
        guards.SendLock("SEND_LOCK"),
        guards.FromAddress("From: office@example.com"),
        guards.Blocklist("bounced.csv"),
        guards.RecipientsReadable(),
        guards.IntendedRecipientsPresent(),
        guards.SubjectMatchesIntended(),
        guards.BodyMatchesIntended(),
    ], draft)
    if problems:
        raise SystemExit("NOT SENT: " + "; ".join(problems))
    m.send()

Nothing in the library sends by itself and no safety rule is built in: the rules that matter are
specific to whose mailbox it is, so you pass them.  See README.md for the browser behaviours this
encodes, each of which cost a real mistake to learn.
"""
from .session import Session, OutlookError, OUTLOOK_HOSTS          # noqa: F401
from .mail import Mailbox                                          # noqa: F401
from .recipients import chip_addresses, resolved_recipients, well_text   # noqa: F401
from .dom import pierce_find, box_center, center_of                 # noqa: F401
from .text import to_html, normalise                               # noqa: F401
from . import guards                                               # noqa: F401

__all__ = ["Session", "Mailbox", "OutlookError", "OUTLOOK_HOSTS", "guards",
           "chip_addresses", "resolved_recipients", "well_text",
           "pierce_find", "box_center", "center_of", "to_html", "normalise"]
__version__ = "0.1.0"
