"""Toast + sound alerter for Teams Meeting Nag.

Uses windows-toasts for Windows 10/11 actionable notifications.
Falls back gracefully when running headless (logs only).
"""
from __future__ import annotations

import logging
import os
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

LOG = logging.getLogger(__name__)

AUMID = "TeamsMeetingNag.App"
APP_DISPLAY_NAME = "Teams Meeting Nag"


def _ensure_aumid_shortcut() -> bool:
    """Create a Start-Menu .lnk tagged with our AUMID if one doesn't exist.

    Win11 silently drops toasts from InteractableWindowsToaster unless the
    AUMID is registered via a shortcut. The shortcut target doesn't matter
    for our use case — Windows just needs to know the AUMID exists.

    Returns True on success (or if shortcut already existed).
    """
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return False
    start_menu = Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs"
    shortcut_path = start_menu / f"{APP_DISPLAY_NAME}.lnk"
    if shortcut_path.exists():
        return True
    try:
        start_menu.mkdir(parents=True, exist_ok=True)
        import pythoncom  # type: ignore
        from win32com.shell import shell  # type: ignore
        from win32com.propsys import propsys, pscon  # type: ignore

        # Point the shortcut at the Python launcher; never actually invoked.
        target = sys.executable
        link = pythoncom.CoCreateInstance(
            shell.CLSID_ShellLink, None,
            pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink,
        )
        link.SetPath(target)
        link.SetDescription(APP_DISPLAY_NAME)

        store = link.QueryInterface(propsys.IID_IPropertyStore)
        store.SetValue(
            pscon.PKEY_AppUserModel_ID,
            propsys.PROPVARIANTType(AUMID, pythoncom.VT_LPWSTR),
        )
        store.Commit()

        persist = link.QueryInterface(pythoncom.IID_IPersistFile)
        persist.Save(str(shortcut_path), True)
        LOG.info("Registered AUMID shortcut at %s", shortcut_path)
        return True
    except Exception as e:
        LOG.warning("Could not create AUMID shortcut (%s); toasts may not appear.", e)
        return False


@dataclass
class AlertAction:
    join: Callable[[], None] = lambda: None
    snooze: Callable[[], None] = lambda: None
    dismiss: Callable[[], None] = lambda: None
    joined_elsewhere: Callable[[], None] = lambda: None


class Alerter:
    """Fires actionable toasts and plays a sound."""

    def __init__(self, sound: str = "beep") -> None:
        self._sound = sound or ""
        self._toaster = None
        self._toaster_cls = None
        self._toast_cls = None
        self._button_cls = None
        self._init_toaster()

    def _init_toaster(self) -> None:
        try:
            from windows_toasts import (  # type: ignore
                InteractableWindowsToaster,
                Toast,
                ToastButton,
            )
        except Exception as e:  # pragma: no cover
            LOG.warning("windows-toasts unavailable: %s — falling back to console.", e)
            return
        # Win11 requires a registered AUMID for interactable toasts.
        _ensure_aumid_shortcut()
        try:
            self._toaster = InteractableWindowsToaster(
                applicationText=APP_DISPLAY_NAME,
                notifierAUMID=AUMID,
            )
            self._toast_cls = Toast
            self._button_cls = ToastButton
        except Exception as e:  # pragma: no cover
            LOG.warning("Failed to init toaster: %s", e)
            self._toaster = None

    def alert(
        self,
        title: str,
        body: str,
        actions: AlertAction,
        join_url: Optional[str] = None,
    ) -> None:
        """Show toast and play sound. Non-blocking."""
        self._play_sound_async()
        if not self._toaster:
            print(f"[NAG] {title} — {body}", flush=True)
            return
        try:
            toast = self._toast_cls()
            toast.text_fields = [title, body]
            if join_url:
                # 'Join' opens the meeting link directly via protocol activation.
                toast.AddAction(self._button_cls(content="Join", launch=join_url))
            toast.AddAction(self._button_cls(content="I'm in it", arguments="joined_elsewhere"))
            toast.AddAction(self._button_cls(content="Snooze 1 min", arguments="snooze"))
            toast.AddAction(self._button_cls(content="Dismiss", arguments="dismiss"))

            def _on_activated(event) -> None:
                arg = getattr(event, "arguments", "") or ""
                if arg == "snooze":
                    actions.snooze()
                elif arg == "dismiss":
                    actions.dismiss()
                elif arg == "joined_elsewhere":
                    actions.joined_elsewhere()
                else:
                    # 'Join' uses launch URL, so no argument arrives here;
                    # treat any other activation as a join intent.
                    actions.join()

            # Intentionally do NOT wire on_dismissed: Windows fires it on
            # auto-timeout / moved-to-Action-Center, which we want to treat
            # as "let it re-nag", not as a permanent dismissal. The explicit
            # "Dismiss" button is the only path to actions.dismiss().
            toast.on_activated = _on_activated
            self._toaster.show_toast(toast)
        except Exception as e:
            LOG.warning("Toast failed (%s); printing instead.", e)
            print(f"[NAG] {title} — {body}", flush=True)

    def _play_sound_async(self) -> None:
        if not self._sound:
            return
        threading.Thread(target=self._play_sound, daemon=True).start()

    def _play_sound(self) -> None:
        try:
            import winsound  # type: ignore
        except Exception:
            return
        sound = self._sound
        try:
            if sound.lower() == "beep":
                # Two-tone urgency beep.
                winsound.Beep(880, 250)
                winsound.Beep(660, 250)
                winsound.Beep(880, 350)
                return
            p = Path(sound)
            if p.exists() and p.suffix.lower() == ".wav":
                winsound.PlaySound(str(p), winsound.SND_FILENAME | winsound.SND_ASYNC)
            else:
                LOG.warning("Sound '%s' not found or not .wav; using beep.", sound)
                winsound.Beep(880, 300)
        except Exception as e:  # pragma: no cover
            LOG.warning("Sound playback failed: %s", e)
