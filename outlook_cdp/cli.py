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
from .reading import Reader
from .recipients import resolved_recipients
from . import dedupe
from .contacts import Contact, Contacts
from .rules import Rule, Rules
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

    q = sub.add_parser("search", help="search messages over REST; changes nothing, marks nothing read")
    q.add_argument("query", nargs="?", default="", help="$search text; omit to list with --from only")
    q.add_argument("--mailbox", default="", help="defaults to the tab URL's; 'me' is the signed-in user")
    q.add_argument("--folder", default="", help="limit to one folder (well-known name or id)")
    q.add_argument("--from", dest="from_address", default="", help="exact sender address")
    q.add_argument("--top", type=int, default=25)
    q.add_argument("--summary", action="store_true", help="count by sender, subject and folder instead")

    r = sub.add_parser("rules", help="list, preview, add, apply or delete inbox rules over REST")
    r.add_argument("action", choices=["list", "preview", "add", "apply", "delete"])
    r.add_argument("--mailbox", default="", help="whose rules; defaults to the tab URL's mailbox, and "
                                                "'me' means the signed-in user")
    r.add_argument("--name", default="", help="rule name (add, apply, delete)")
    r.add_argument("--id", default="", help="rule id (delete)")
    r.add_argument("--from", dest="from_addresses", action="append", default=[], metavar="ADDR")
    r.add_argument("--sender-contains", action="append", default=[], metavar="TEXT")
    r.add_argument("--subject-contains", action="append", default=[], metavar="TEXT")
    r.add_argument("--body-contains", action="append", default=[], metavar="TEXT")
    r.add_argument("--move-to", default="", help="well-known folder name (archive, ...) or folder id")
    r.add_argument("--mark-read", action="store_true")
    r.add_argument("--folder", default="inbox", help="folder that preview / apply look in")

    k = sub.add_parser("contacts", help="list, find, add, update, photograph, delete, back up, restore or "
                                        "deduplicate contacts over REST")
    k.add_argument("action", choices=["list", "find", "add", "update", "photo", "delete", "backup", "restore",
                                      "dedupe-plan", "dedupe-apply"])
    k.add_argument("--mailbox", default="", help="whose contacts; defaults to the tab URL's mailbox, and "
                                                "'me' means the signed-in user")
    k.add_argument("--phone", default="", help="find: number to match, formatting ignored")
    k.add_argument("--name", default="", help="find: display-name substring")
    k.add_argument("--json", dest="json_file", default="",
                   help="add: one Contact's fields as a JSON object, or a list of them; "
                        "update: the fields to change, \"\" to clear")
    k.add_argument("--allow-duplicate-phone", action="store_true",
                   help="add: create even when a number is already on another contact")
    k.add_argument("--id", default="", help="update, photo, delete: contact id")
    k.add_argument("--photo", default="", help="photo: a JPEG, PNG or GIF file to set as the picture")
    k.add_argument("--out", default="", help="backup: directory for the backup file; dedupe-plan: plan file")
    k.add_argument("--backup", default="", help="restore, dedupe-plan, dedupe-apply: a backup file")
    k.add_argument("--plan", default="", help="dedupe-apply: the reviewed plan file")
    k.add_argument("--same-folder", action="store_true",
                   help="dedupe-plan: only treat contacts in the same folder as duplicates")
    k.add_argument("--ignore-category", action="append", default=[], metavar="NAME",
                   help="dedupe-plan: a category that is not information (repeatable)")
    k.add_argument("--log", default="", help="dedupe-apply: JSON-lines log, appended to")
    k.add_argument("--limit", type=int, default=0, help="dedupe-apply: stop after this many deletions")
    k.add_argument("--delete", action="store_true",
                   help="dedupe-apply: actually delete; without it, only report what would be deleted")

    a = p.parse_args(argv)
    s = Session(a.instance, host=a.host)
    if a.cmd == "rules":
        return rules_cmd(a, s)
    if a.cmd == "contacts":
        return contacts_cmd(a, s)
    if a.cmd == "search":
        return search_cmd(a, s)
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


def search_cmd(a, s: Session) -> int:
    mailbox = Rules.signed_in_address(s) if a.mailbox == "me" else (a.mailbox or None)
    try:
        r = Reader(s, wake=False, mailbox=mailbox)
        if a.from_address:
            # Exact sender, every match: a sender filter needs no $search and is not capped by it.
            sender = Rules(s, mailbox or r.base.rsplit("/", 1)[1])
            hits = sender.preview(Rule("search", from_addresses=[a.from_address]),
                                  a.folder or "inbox", limit=max(a.top, 1))
            if a.query:
                hits = [h for h in hits if a.query.lower() in (h.subject or "").lower()]
        else:
            if not a.query:
                print("search needs a query or --from", file=sys.stderr)
                return 2
            hits = r.search(a.query, folder=a.folder or None, top=a.top)
    except OutlookError as e:
        print(f"REFUSED: {e}", file=sys.stderr)
        return 3
    if a.summary:
        from collections import Counter
        counts = Counter((h.sender_address, h.subject, h.folder) for h in hits)
        print(json.dumps({"messages": len(hits), "groups": [[n, *k] for k, n in counts.most_common()]},
                         indent=1))
    else:
        print(json.dumps([[h.received, h.folder, h.sender_address, h.subject] for h in hits], indent=1))
    return 0


def rules_cmd(a, s: Session) -> int:
    mailbox = Rules.signed_in_address(s) if a.mailbox == "me" else (a.mailbox or None)
    try:
        rules = Rules(s, mailbox)
        if a.action == "list":
            print(json.dumps([vars(x) for x in rules.list()], indent=1))
            return 0
        if a.action == "delete":
            ids = [a.id] if a.id else [x.id for x in rules.find(a.name)] if a.name else []
            if not ids:
                print("delete needs --id or a --name that exists", file=sys.stderr)
                return 2
            for rid in ids:
                rules.delete(rid)
            print(json.dumps({"deleted": ids}))
            return 0
        rule = Rule(a.name or "unnamed", from_addresses=a.from_addresses,
                    sender_contains=a.sender_contains, subject_contains=a.subject_contains,
                    body_contains=a.body_contains, move_to=a.move_to or None, mark_read=a.mark_read)
        if a.action in ("preview", "apply") and a.name and not rule.conditions():
            # Use the rule as the server stores it, so what is applied is exactly what is live.
            stored = rules.find(a.name)
            if len(stored) != 1:
                print(f"{len(stored)} stored rules are named {a.name!r}; need exactly one",
                      file=sys.stderr)
                return 2
            rule = stored[0]
        if a.action == "preview":
            hits = rules.preview(rule, a.folder)
            print(json.dumps({"matches": len(hits), "newest": [[h.received, h.sender_address, h.subject]
                                                                for h in hits[:20]]}, indent=1))
            return 0
        if a.action == "add":
            if not a.name:
                print("add needs --name", file=sys.stderr)
                return 2
            print(json.dumps(vars(rules.create(rule)), indent=1))
            return 0
        if a.action == "apply":
            moved = rules.apply(rule, a.folder)
            print(json.dumps({"moved": len(moved)}))
            return 0
    except OutlookError as e:
        print(f"REFUSED: {e}", file=sys.stderr)
        return 3
    return 2


def contacts_cmd(a, s: Session) -> int:
    if a.action == "dedupe-plan":                           # offline: reads only the backup file
        if not (a.backup and a.out):
            print("dedupe-plan needs --backup and --out", file=sys.stderr)
            return 2
        plan = dedupe.plan(dedupe.load(a.backup), across_folders=not a.same_folder,
                           ignore_categories=a.ignore_category)
        dedupe.write_plan(plan, a.out)
        print(json.dumps(plan["summary"], indent=1))
        return 0
    mailbox = Rules.signed_in_address(s) if a.mailbox == "me" else (a.mailbox or None)
    show = lambda c: {k: v for k, v in vars(c).items() if v}   # noqa: E731
    try:
        book = Contacts(s, mailbox)
        if a.action == "backup":
            if not a.out:
                print("backup needs --out DIRECTORY", file=sys.stderr)
                return 2
            print(json.dumps(book.backup(a.out), indent=1))
            return 0
        if a.action == "restore":
            if not (a.backup and a.id):
                print("restore needs --backup and --id (the contact's id in the backup)", file=sys.stderr)
                return 2
            entry = next((e for e in dedupe.load(a.backup)["contacts"] if e["record"]["Id"] == a.id), None)
            if entry is None:
                print(f"no contact with id {a.id} in that backup", file=sys.stderr)
                return 2
            made = book.restore(entry)
            print(json.dumps({"restored": dedupe.display(entry), "new_id": made["Id"]}))
            return 0
        if a.action == "dedupe-apply":
            if not (a.backup and a.plan and a.log):
                print("dedupe-apply needs --backup, --plan and --log", file=sys.stderr)
                return 2
            plan, backup = dedupe.load(a.plan), dedupe.load(a.backup)
            if not a.delete:
                todo = dedupe.approved(plan)[:a.limit or None]
                print(json.dumps({"would_delete": len(todo), "first": [
                    {"delete": d["delete"]["name"], "keep": d["keep"]["name"]} for d in todo[:20]]}, indent=1))
                print("Nothing deleted. Add --delete to carry out the plan.", file=sys.stderr)
                return 0
            print(json.dumps(dedupe.apply(book, plan, backup, a.log, limit=a.limit or None)))
            return 0
        if a.action == "list":
            print(json.dumps([show(c) for c in book.list()], indent=1))
            return 0
        if a.action == "find":
            if not (a.phone or a.name):
                print("find needs --phone or --name", file=sys.stderr)
                return 2
            everyone = book.list()
            hits = book.find_by_phone(a.phone, everyone) if a.phone else book.find_by_name(a.name, everyone)
            print(json.dumps([show(c) for c in hits], indent=1))
            return 0
        if a.action == "delete":
            if not a.id:
                print("delete needs --id", file=sys.stderr)
                return 2
            book.delete(a.id)
            print(json.dumps({"deleted": a.id}))
            return 0
        if a.action == "photo":
            if not (a.id and a.photo):
                print("photo needs --id and --photo", file=sys.stderr)
                return 2
            kind = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "gif": "image/gif"}
            ct = kind.get(a.photo.rsplit(".", 1)[-1].lower(), "")
            print(json.dumps(book.set_photo(a.id, open(a.photo, "rb").read(), ct)))
            return 0
        if not a.json_file:
            print(f"{a.action} needs --json", file=sys.stderr)
            return 2
        if a.action == "update":
            if not a.id:
                print("update needs --id", file=sys.stderr)
                return 2
            changes = json.load(open(a.json_file, encoding="utf-8"))
            print(json.dumps(show(book.update(a.id, **changes)), indent=1))
            return 0
        raw = json.load(open(a.json_file, encoding="utf-8"))
        made, failed = [], 0
        for fields in raw if isinstance(raw, list) else [raw]:
            try:
                c = book.create(Contact(**fields), allow_duplicate_phone=a.allow_duplicate_phone)
                made.append({"created": c.name(), "id": c.id})
            except (OutlookError, TypeError) as e:
                failed += 1
                made.append({"refused": Contact(**fields).name() if isinstance(e, OutlookError) else
                             str(fields)[:80], "why": str(e)})
        print(json.dumps(made, indent=1))
        return 3 if failed else 0
    except OutlookError as e:
        print(f"REFUSED: {e}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    sys.exit(main())
