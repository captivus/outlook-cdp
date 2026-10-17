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
from .reading import (Reader, Row, Message, Conversation, html_to_text,   # noqa: F401
                      parse_row_flags, parse_count_badge, parse_container_count,
                      split_sender, on_behalf_of, parse_recipients, strip_glyphs,
                      addresses_in)
from .recipients import chip_addresses, resolved_recipients, well_text   # noqa: F401
from .dom import (pierce_find, box_center, center_of,                # noqa: F401
                  center_in_viewport, scroll_into_view_center)
from .text import to_html, normalise                               # noqa: F401
from .attachments import (Attachment, access_token, list_attachments,      # noqa: F401
                          messages_with_attachments, open_message_attachments,
                          open_conversation_messages, download, download_all, verify_file)
from .rules import Rule, Rules, rest_send                          # noqa: F401
from .contacts import Contact, Contacts, phone_digits              # noqa: F401
from . import guards                                               # noqa: F401

__all__ = ["Session", "Mailbox", "OutlookError", "OUTLOOK_HOSTS", "guards",
           "Reader", "Row", "Message", "Conversation",
           "parse_row_flags", "parse_count_badge", "parse_container_count",
           "split_sender", "on_behalf_of", "parse_recipients", "strip_glyphs",
           "html_to_text", "addresses_in",
           "chip_addresses", "resolved_recipients", "well_text",
           "pierce_find", "box_center", "center_of", "to_html", "normalise",
           "Attachment", "access_token", "list_attachments", "messages_with_attachments",
           "open_message_attachments", "open_conversation_messages", "download",
           "download_all", "verify_file", "Rule", "Rules", "rest_send",
           "Contact", "Contacts", "phone_digits"]
__version__ = "0.1.0"
