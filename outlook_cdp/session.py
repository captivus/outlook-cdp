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
import urllib.parse

#: Hosts Outlook on the web has been served from.  Microsoft moved from outlook.office.com to
#: outlook.cloud.microsoft mid-2026 and both still answer, so the host is resolved at runtime from the
#: live target list rather than hardcoded.  Hardcoding it is what broke the first version of this code.
OUTLOOK_HOSTS = ("outlook.cloud.microsoft", "outlook.office.com", "outlook.live.com")


class OutlookError(RuntimeError):
    """Anything that means the page is not in the state the caller assumed."""


def _describe(tabs: list[dict]) -> str:
    return "\n".join(f"  {t.get('id', '?')}  {t.get('url', '')}  {t.get('title', '')!r}"
                     for t in tabs) or "  (none)"


class Session:
    """One Outlook tab.

    >>> s = Session("my-instance")
    >>> s.js("document.title")

    The tab is never activated.  CDP input events reach a background tab, so the human keeps whatever
    window they had in front; `focus_emulation()` lets the background tab still complete dialogs and
    transitions.
    """

    def __init__(self, instance: str, host: str | None = None, timeout: float = 120.0,
                 target: str | None = None):
        """`target` is a chrome-agent target id (or a unique prefix of one, as `chrome-agent status`
        prints it).  It is only needed when the instance holds more than one Outlook tab, which it
        does whenever two accounts are open side by side; then every tab sits on the same host and
        the same bare ``/mail/`` URL, so nothing but the id can tell them apart."""
        self.instance = instance
        self.timeout = timeout
        self._host = host
        self._target = target

    # ---------------------------------------------------------------- plumbing

    def targets(self) -> list[dict]:
        """Every tab in the instance, as chrome-agent's ``status`` lists them (id, url, title...)."""
        r = subprocess.run(["chrome-agent", "status", self.instance],
                           capture_output=True, text=True, timeout=self.timeout)
        try:
            listing = json.loads(r.stdout)
        except ValueError:
            raise OutlookError(f"chrome-agent status {self.instance!r} returned no JSON.\n"
                               f"stdout: {r.stdout[:400]}\nstderr: {r.stderr[:400]}")
        for inst in listing if isinstance(listing, list) else [listing]:
            if inst.get("name") == self.instance:
                return inst.get("targets") or []
        raise OutlookError(f"no chrome-agent instance named {self.instance!r}")

    def outlook_tabs(self) -> list[dict]:
        """The instance's Outlook tabs, narrowed to `host` when one was given."""
        hosts = (self._host,) if self._host else OUTLOOK_HOSTS
        return [t for t in self.targets()
                if any(urllib.parse.urlsplit(t.get("url", "")).hostname == h for h in hosts)]

    def tab(self) -> dict:
        """This session's tab, as chrome-agent lists it.

        Read afresh every call, since the tab's URL changes as Outlook navigates; only its id is
        fixed.  With no `target` given there must be exactly one Outlook tab, and the first call pins
        to it by id.  Picking "the first match" instead would quietly drive the wrong account."""
        if self._target:
            want = self._target.upper()
            hits = [t for t in self.targets()
                    if str(t.get("full_id", t.get("id", ""))).upper().startswith(want)]
            if len(hits) != 1:
                raise OutlookError(
                    f"target {self._target!r} matches {len(hits)} tabs in chrome-agent instance "
                    f"{self.instance!r}. Outlook tabs are:\n" + _describe(self.outlook_tabs()))
            return hits[0]
        tabs = self.outlook_tabs()
        if not tabs:
            raise OutlookError(
                f"no Outlook tab in chrome-agent instance {self.instance!r}. Tabs are:\n"
                + _describe(self.targets()))
        if len(tabs) > 1:
            raise OutlookError(
                f"chrome-agent instance {self.instance!r} has {len(tabs)} Outlook tabs; choose one "
                f"with --target ID (Session(target=ID)):\n" + _describe(tabs))
        self._target = tabs[0].get("full_id") or tabs[0]["id"]
        return tabs[0]

    @property
    def target(self) -> str:
        """The chrome-agent target id of this session's tab."""
        t = self.tab()
        return t.get("full_id") or t["id"]

    @property
    def host(self) -> str:
        if self._host is None:
            host = urllib.parse.urlsplit(self.tab().get("url", "")).hostname or ""
            if host not in OUTLOOK_HOSTS:
                raise OutlookError(f"target {self._target!r} is not an Outlook tab: "
                                   f"{self.tab().get('url', '')}")
            self._host = host
        return self._host

    def tab_url(self) -> str:
        """The Outlook tab's own URL, from chrome-agent's target list."""
        return self.tab().get("url", "")

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
            ["chrome-agent", self.instance, "--target-id", self._target or self.target, method,
             json.dumps(params or {})],
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
