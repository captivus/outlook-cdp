"""A connection to one Outlook-on-the-web tab driven through chrome-agent and CDP.

Everything here is deliberately stateless between calls: chrome-agent makes a one-shot CDP connection
per invocation, which is slower than holding a socket but means a crashed script leaves nothing behind.
`backendNodeId`s survive across those one-shots, which is what makes the piercing walk in dom.py usable.
"""
from __future__ import annotations

import base64
import json
import re
import subprocess
import time

#: Hosts Outlook on the web has been served from.  Microsoft moved from outlook.office.com to
#: outlook.cloud.microsoft mid-2026 and both still answer, so the host is resolved at runtime from the
#: live target list rather than hardcoded.  Hardcoding it is what broke the first version of this code.
OUTLOOK_HOSTS = ("outlook.cloud.microsoft", "outlook.office.com", "outlook.live.com")


class OutlookError(RuntimeError):
    """Anything that means the page is not in the state the caller assumed."""


class Session:
    """One Outlook tab.

    >>> s = Session("my-instance")
    >>> s.js("document.title")

    The tab is never activated.  CDP input events reach a background tab, so the human keeps whatever
    window they had in front; `focus_emulation()` lets the background tab still complete dialogs and
    transitions.
    """

    def __init__(self, instance: str, host: str | None = None, timeout: float = 120.0):
        self.instance = instance
        self.timeout = timeout
        self._host = host

    # ---------------------------------------------------------------- plumbing

    @property
    def host(self) -> str:
        if self._host is None:
            r = subprocess.run(["chrome-agent", "status", self.instance],
                               capture_output=True, text=True, timeout=self.timeout)
            listing = r.stdout + r.stderr
            for cand in OUTLOOK_HOSTS:
                if cand in listing:
                    self._host = cand
                    break
            else:
                raise OutlookError(
                    f"no Outlook tab in chrome-agent instance {self.instance!r}. Targets were:\n"
                    + listing[:800])
        return self._host

    def tab_url(self) -> str:
        """The Outlook tab's own URL, from chrome-agent's target list."""
        r = subprocess.run(["chrome-agent", "status", self.instance],
                           capture_output=True, text=True, timeout=self.timeout)
        for line in (r.stdout + r.stderr).splitlines():
            if '"url"' in line and self.host in line:
                return line.split('"')[3] if line.count('"') >= 4 else line.strip()
        return ""

    def mailbox_identity(self) -> str | None:
        """The mailbox this tab is scoped to, read from its URL.

        Outlook shows a From chooser only when the session has more than one identity to choose from.
        Signed in to a single mailbox there is no From control at all, so a guard that demands one
        refuses every send - which is what happened the first time this library met a single-mailbox
        session.  The URL still says whose mailbox it is: `/mail/office@example.com/`.
        """
        m = re.search(r"/mail/([^/@]+@[^/]+)/", self.tab_url())
        return m.group(1).lower() if m else None

    def cdp(self, method: str, params: dict | None = None):
        r = subprocess.run(
            ["chrome-agent", self.instance, "--url", self.host, method, json.dumps(params or {})],
            capture_output=True, text=True, timeout=self.timeout)
        try:
            return json.loads(r.stdout)
        except Exception:
            raise OutlookError(f"{method} returned no JSON.\nstdout: {r.stdout[:400]}\n"
                               f"stderr: {r.stderr[:400]}")

    def js(self, expression: str, await_promise: bool = False):
        d = self.cdp("Runtime.evaluate", {"expression": expression, "returnByValue": True,
                                          "awaitPromise": await_promise})
        if "exceptionDetails" in d:
            raise OutlookError("JS threw: " + json.dumps(d["exceptionDetails"])[:500])
        return d.get("result", {}).get("value")

    # ------------------------------------------------------------------ input

    def click(self, x: float, y: float) -> None:
        """A *trusted* click.

        `element.click()` from JS silently does nothing on several Outlook controls, because they check
        `event.isTrusted`.  Dispatching through the Input domain produces a real one.  Anything that
        looks like "the click ran but nothing happened" is almost always this.
        """
        for kind in ("mousePressed", "mouseReleased"):
            self.cdp("Input.dispatchMouseEvent",
                     {"type": kind, "x": x, "y": y, "button": "left", "clickCount": 1})

    def key(self, key: str, code: str, vk: int) -> None:
        for kind in ("keyDown", "keyUp"):
            self.cdp("Input.dispatchKeyEvent",
                     {"type": kind, "key": key, "code": code, "windowsVirtualKeyCode": vk})

    def type_text(self, text: str) -> None:
        """Insert text without per-character key events. Faster and avoids autocomplete races."""
        self.cdp("Input.insertText", {"text": text})

    def focus_emulation(self) -> None:
        """Let a background tab behave as if focused, without raising the window.

        Outlook's confirmation dialogs and pane transitions stall in a backgrounded tab. These two
        calls unstall them while leaving the human's foreground window alone.
        """
        self.cdp("Emulation.setFocusEmulationEnabled", {"enabled": True})
        self.cdp("Page.setWebLifecycleState", {"state": "active"})

    # ------------------------------------------------------------------ output

    def screenshot(self, path: str) -> str:
        d = self.cdp("Page.captureScreenshot", {"format": "png"})
        with open(path, "wb") as fh:
            fh.write(base64.b64decode(d["data"]))
        return path

    # ------------------------------------------------------------------ waiting

    def wait_for(self, js_predicate: str, timeout: float = 10.0, interval: float = 0.25) -> bool:
        """Poll a JS boolean expression until it is true. Returns False on timeout, never raises."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                if self.js(js_predicate):
                    return True
            except OutlookError:
                pass
            time.sleep(interval)
        return False
