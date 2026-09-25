"""Turning plain text into the HTML Outlook's composer accepts.

Outlook's body is a contenteditable div.  Pasting `text/html` through a synthetic ClipboardEvent is the
only approach that reliably preserves paragraphs; typing a multi-line string inserts line breaks that
Outlook then reflows, and setting innerHTML directly loses the undo stack and sometimes the send button's
dirty-state tracking.
"""
from __future__ import annotations

import html
import re


def to_html(text: str) -> str:
    """Rejoin hard-wrapped prose into paragraphs while keeping lists and address blocks line-for-line.

    Letters are usually wrapped at 80-100 columns for the file, and those wraps are not meaningful line
    breaks.  Rejoining them is what makes the sent mail look written rather than pre-formatted.  Three
    shapes are treated differently:

    * a block where every line is short (< 60 chars) is an address or signature - keep every line;
    * a block containing list items (indented, `1.`, `(a)`, `- `) keeps one line per item;
    * anything else is prose - join it into one paragraph.
    """
    out = []
    for block in re.split(r"\n\s*\n", text.strip("\n")):
        lines = block.split("\n")
        listy = [bool(re.match(r"\s+\S|\(\w\)\s|\d+\.\s|-\s", l)) for l in lines]
        if all(len(l) < 60 for l in lines):
            out.append("<br>".join(html.escape(l.strip()) for l in lines))
        elif any(listy):
            paras: list[str] = []
            for line, is_item in zip(lines, listy):
                if is_item or not paras:
                    paras.append(line.strip())
                else:
                    paras[-1] += " " + line.strip()
            out.append("<br>".join(html.escape(p) for p in paras))
        else:
            out.append(html.escape(" ".join(l.strip() for l in lines)))
    return "".join(f"<div>{p}</div><div><br></div>" for p in out)


def normalise(s: str) -> str:
    """Collapse all whitespace, for comparing intended text against what the pane holds.

    to_html rejoins wrapped lines, so the body in the pane differs from the source file in line breaks
    and nothing else.  Comparing normalised strings is what lets a verification step check the *whole*
    letter instead of its opening words - a prefix check passes on a body that was truncated, pasted
    twice, or left sitting on top of an older draft.
    """
    return " ".join((s or "").split())
