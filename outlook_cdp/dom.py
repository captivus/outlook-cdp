"""Finding controls Outlook does not put in the light DOM.

Outlook renders the reading pane's Reply / Reply all / Forward buttons inside **closed** shadow roots.
`document.querySelectorAll` cannot see into a closed root from page script, so any amount of JS
selector work will report the button does not exist.  It does exist; it is just not reachable that way.

`DOM.getDocument` with `pierce: true` walks through closed roots, and the `backendNodeId`s it returns
stay valid across chrome-agent's separate one-shot connections, so you can find a node in one call and
measure it in the next.
"""
from __future__ import annotations

import time

from .session import Session


def pierce_find(session: Session, labels, node_name: str = "BUTTON") -> list[int]:
    """Every `backendNodeId` whose aria-label matches one of `labels`, through closed shadow roots.

    Matching is case-insensitive on the trimmed label, and restricted to one tag name because Outlook
    labels several wrapper elements identically to the control inside them.
    """
    want = {str(l).strip().lower() for l in labels}
    root = session.cdp("DOM.getDocument", {"depth": -1, "pierce": True}).get("root", {})
    found: list[int] = []

    def walk(node):
        attrs = node.get("attributes") or []
        d = {attrs[i]: attrs[i + 1] for i in range(0, len(attrs) - 1, 2)}
        if (d.get("aria-label") or "").strip().lower() in want and node.get("nodeName") == node_name:
            found.append(node.get("backendNodeId"))
        for key in ("children", "shadowRoots", "contentDocument"):
            v = node.get(key)
            if isinstance(v, dict):
                walk(v)
            elif isinstance(v, list):
                for child in v:
                    walk(child)

    walk(root)
    return found


def box_center(session: Session, backend_node_id: int, in_viewport: bool = True):
    """Viewport centre of a pierced node, or None when it cannot be clicked there.

    A node can exist and have no box: collapsed, display:none, or in a pane that has not rendered.
    Clicking a zero-area box does nothing and looks like a mystery, so degenerate boxes return None and
    the caller should try the next candidate rather than the first.

    A node can also have a **perfectly healthy box that is outside the viewport**, and that is the more
    expensive mistake.  The reading pane scrolls independently of the page, so an element can sit at
    y=2100 in a 1321px viewport with a 32x46 box; `Input.dispatchMouseEvent` accepts those coordinates,
    reports success, and the click lands on nothing.  That is the whole of the "the click ran and
    nothing happened" mystery that held up the attachment work, and a box model alone cannot show it.
    With `in_viewport=True` such a centre returns None, so the caller scrolls and re-measures - or at
    least fails loudly - instead of clicking into space.
    """
    try:
        model = session.cdp("DOM.getBoxModel", {"backendNodeId": backend_node_id}).get("model", {})
    except Exception:
        return None
    quad = model.get("content")
    if not quad or len(quad) < 8:
        return None
    xs, ys = quad[0::2], quad[1::2]
    if max(xs) - min(xs) < 2 or max(ys) - min(ys) < 2:
        return None
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    if in_viewport and not center_in_viewport(session, cx, cy):
        return None
    return cx, cy


def center_in_viewport(session: Session, x: float, y: float) -> bool:
    """Is (x, y) somewhere a dispatched mouse event can actually land?"""
    size = session.js("[innerWidth, innerHeight]") or [0, 0]
    return 0 <= x <= (size[0] or 0) and 0 <= y <= (size[1] or 0)


def scroll_into_view_center(session: Session, backend_node_id: int, settle: float = 0.4):
    """Scroll a pierced node into view and then return its centre. None if it still cannot be clicked.

    The two steps have to be separate calls: the scroll is animated, so a box measured in the same
    breath as the request to scroll describes where the element *was*.
    """
    try:
        session.cdp("DOM.scrollIntoViewIfNeeded", {"backendNodeId": backend_node_id})
    except Exception:
        return None
    time.sleep(settle)
    return box_center(session, backend_node_id, in_viewport=True)


def center_of(session: Session, selector_js: str):
    """Centre of a light-DOM element, given a JS expression that evaluates to it (or null)."""
    return session.js(
        f"(()=>{{const e={selector_js};if(!e)return null;const r=e.getBoundingClientRect();"
        f"return [r.x+r.width/2,r.y+r.height/2]}})()")


def click_first_laid_out(session: Session, labels, after_js: str, timeout: float = 10.0) -> str | None:
    """Click the first pierced control with a real box, and wait for `after_js` to become true.

    Returns the label that worked, or None.  Written as a loop over candidates because Outlook often
    has several nodes with the same aria-label and only one of them is the live control.
    """
    for label in labels:
        for node_id in pierce_find(session, [label]):
            # Try where it is, then try scrolling it into view: a control below the fold has a real box
            # and is still not clickable, and skipping straight to the next candidate would pass over
            # the one live control on the page.
            xy = box_center(session, node_id) or scroll_into_view_center(session, node_id)
            if not xy:
                continue
            session.click(*xy)
            if session.wait_for(after_js, timeout=timeout):
                return label
    return None
