"""Local Teams meeting detection via window enumeration.

Strategy: when Teams puts you in a meeting / call, it opens a top-level
window whose title contains "Microsoft Teams" plus the meeting subject
(e.g. "Sprint planning | Microsoft Teams"). The main Teams shell window
has title "Microsoft Teams" or "Chat | Microsoft Teams" — we ignore those.

Limitations:
- Cannot detect meetings joined on phone / web / other devices. The
  user is expected to click the "Joined elsewhere" toast button for those.
- If the user types the meeting subject into Teams chat search, that chat
  window's title might match. Acceptable tradeoff.
"""
from __future__ import annotations

import logging
from typing import List, Optional, Tuple

LOG = logging.getLogger(__name__)

# Windows whose titles equal these (case-insensitive) are the Teams shell,
# not a meeting window.
_SHELL_TITLES = {
    "microsoft teams",
    "microsoft teams (preview)",
    "chat | microsoft teams",
    "activity | microsoft teams",
    "calendar | microsoft teams",
    "calls | microsoft teams",
    "teams | microsoft teams",
    "files | microsoft teams",
    "apps | microsoft teams",
}


def _enum_teams_windows() -> List[Tuple[int, str]]:
    """Return [(hwnd, title), ...] for visible windows whose title mentions Teams."""
    try:
        import win32gui  # type: ignore
    except ImportError:
        LOG.warning("pywin32 missing; cannot detect local Teams windows.")
        return []

    results: List[Tuple[int, str]] = []

    def cb(hwnd, _ctx):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd) or ""
            if "microsoft teams" in title.lower():
                results.append((hwnd, title))
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(cb, None)
    except Exception as e:
        LOG.warning("EnumWindows failed: %s", e)
    return results


def _is_meeting_window(title: str, subject: Optional[str]) -> bool:
    t = (title or "").strip().lower()
    if not t or t in _SHELL_TITLES:
        return False
    # Skip chat-tab windows. Teams formats chat threads as "Chat | <name> | ... | Microsoft Teams"
    # which would otherwise false-match the meeting subject.
    if t.startswith("chat | ") or t.startswith("chat: "):
        return False
    # Subject match: Teams meeting windows include the meeting subject.
    if subject:
        s = subject.strip().lower()
        if len(s) >= 4:
            # Truncate to handle title clipping by Windows.
            needle = s[:40]
            if needle in t:
                return True
    return False


def is_in_local_teams_meeting(subject: Optional[str] = None) -> bool:
    """True if a Teams meeting window for *subject* appears to be open on this PC.

    Without a subject, we cannot reliably distinguish a meeting window from
    a normal Teams chat window, so we return False (= keep nagging).
    """
    if not subject:
        return False
    windows = _enum_teams_windows()
    for _hwnd, title in windows:
        if _is_meeting_window(title, subject):
            LOG.debug("Detected local Teams meeting window: %r", title)
            return True
    return False


def debug_list_teams_windows() -> List[str]:
    """Return all Teams-related window titles (for diagnostics)."""
    return [t for _h, t in _enum_teams_windows()]
