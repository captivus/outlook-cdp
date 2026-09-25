"""Reading who a draft is actually addressed to.

The To/Cc/Bcc wells render each recipient as a chip showing a **display name**, so any check that greps
the well's `innerText` for a blocked address will pass a draft addressed to exactly that address.

Chips come in two kinds and the difference decides what any guard can conclude:

* **An address Outlook did not resolve** - an external recipient - keeps its address in the chip's
  `aria-label`, formatted `Display Name <a@b.com>`.
* **An address Outlook resolved against the directory** - your own mailbox, a colleague - is replaced by
  a contact, and **the address is then in no attribute anywhere under the well.** Verified live: sending
  to `office@example.com` from that mailbox can produce a chip whose aria-label reads
  `unknownExample Office` and nothing else. Searching every attribute of every descendant for the
  address returns nothing.

So this returns **both** the addresses it could read and the display names it could not resolve to an
address, and only calls a chip unreadable when it yields neither. A caller that needs to act on a
directory-resolved recipient must supply the display name it expects; there is nothing in the page to
compare against otherwise.
"""
from __future__ import annotations

import json

from .session import Session

WELLS = ("To", "Cc", "Bcc")
#: Chips are spans whose id begins REK, or carry role=option / data-lp-id=recipientWell-chip.
#: Outlook has used all three; matching any of them survives the next rename.
_CHIP_JS = r"""(()=>{
  const root = document.querySelector('div[aria-label=%s]');
  if (!root) return null;
  const well = root.parentElement.parentElement.parentElement;
  const addresses = new Set(), names = new Set(), unreadable = [];
  let chips = 0;
  const rx = /[A-Za-z0-9._%%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
  const walk = (n) => {
    if (!n) return;
    if (n.nodeType === 1) {
      const id = n.id || '';
      const isChip = id.startsWith('REK') ||
                     (n.getAttribute && (n.getAttribute('role') === 'option' ||
                                         n.getAttribute('data-lp-id') === 'recipientWell-chip'));
      if (isChip) {
        chips++;
        const label = n.getAttribute('aria-label') || '';
        const blob = [label, n.getAttribute('title'), n.textContent].filter(Boolean).join(' ');
        const found = blob.match(rx) || [];
        found.forEach(a => addresses.add(a.toLowerCase()));
        if (!found.length) {
          // A directory-resolved contact. Outlook prefixes the label with a presence word such as
          // "unknown"/"available"; strip a leading lowercase run so the name is usable.
          const name = (label || n.textContent || '').trim().replace(/^[a-z]+(?=[A-Z])/, '').trim();
          if (name) names.add(name); else unreadable.push((n.textContent || '').trim().slice(0, 60));
        }
      }
      if (n.shadowRoot) walk(n.shadowRoot);
    }
    for (const c of (n.childNodes || [])) walk(c);
  };
  walk(well);
  // Fallback: any address visible anywhere in the well, in case the chip markup changes again.
  // A superset is the safe direction for a blocklist check.
  ((well.innerText || '').match(rx) || []).forEach(a => addresses.add(a.toLowerCase()));
  return {addresses: [...addresses], display_names: [...names], chips: chips, unreadable: unreadable};
})()"""


def chip_addresses(session: Session, well: str) -> dict | None:
    """`{addresses, display_names, chips, unreadable}`, or None if the well is not on the page."""
    return session.js(_CHIP_JS % json.dumps(well))


def resolved_recipients(session: Session, wells=WELLS, required=("To",)):
    """`(addresses, display_names, problems)` across the wells.

    Fails closed on what it genuinely cannot see: a well that is missing when it was required, and any
    chip that yields neither an address nor a name. It does **not** treat a directory-resolved contact as
    a failure - that is a normal, correct state, and calling it unreadable would refuse every internal
    recipient. The caller decides whether a name without an address is good enough for what it is doing.
    """
    addresses: set[str] = set()
    names: set[str] = set()
    problems: list[str] = []
    for well in wells:
        result = chip_addresses(session, well)
        if result is None:
            if well in required:
                problems.append(f"could not locate the {well} well, so its recipients are unverified")
            continue
        addresses.update(result["addresses"])
        names.update(result.get("display_names") or [])
        for label in result["unreadable"]:
            problems.append(f"{well} chip {label!r} yields neither an address nor a name")
        if result["chips"] and not (result["addresses"] or result.get("display_names")):
            problems.append(f"{well} holds {result['chips']} chip(s) and none could be identified")
    return sorted(addresses), sorted(names), problems


def well_text(session: Session, well: str) -> str:
    """The well's visible text. Useful for messages to a human; never for a safety check."""
    return session.js(
        '(()=>{const e=document.querySelector(\'div[aria-label="%s"]\');'
        'if(!e)return "";const c=e.closest(\'[role="presentation"]\')||e.parentElement.parentElement;'
        'return c.innerText||""})()' % well) or ""
