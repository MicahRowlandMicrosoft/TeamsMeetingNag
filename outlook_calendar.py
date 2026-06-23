"""Calendar source using Outlook desktop via COM automation.

Requires Outlook to be installed and a profile configured. No auth flow —
reuses your already-signed-in Outlook. Works fully offline (against cached
items) when Outlook is in cached-Exchange mode.

If Outlook isn't running, this will start it (per COM activation behavior),
which is usually fine; the user almost always has it running.
"""
from __future__ import annotations

import datetime as _dt
import logging
from typing import List, Optional

from models import MeetingEvent, extract_teams_url

LOG = logging.getLogger(__name__)

# Outlook OlDefaultFolders.olFolderCalendar
_OL_FOLDER_CALENDAR = 9


class OutlookCalendar:
    """Thin wrapper around the Outlook MAPI calendar folder."""

    def __init__(self) -> None:
        self._outlook = None
        self._namespace = None

    def _connect(self) -> None:
        if self._outlook is not None:
            return
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except ImportError as e:
            raise RuntimeError(
                "pywin32 is required for Outlook integration. "
                "Install it with: pip install pywin32"
            ) from e

        # CoInitialize on this thread. Safe to call repeatedly.
        pythoncom.CoInitialize()
        self._outlook = win32com.client.Dispatch("Outlook.Application")
        self._namespace = self._outlook.GetNamespace("MAPI")
        LOG.debug("Connected to Outlook via COM.")

    def get_events(
        self,
        window_start_utc: _dt.datetime,
        window_end_utc: _dt.datetime,
    ) -> List[MeetingEvent]:
        """Return events overlapping [window_start_utc, window_end_utc]."""
        self._connect()
        cal = self._namespace.GetDefaultFolder(_OL_FOLDER_CALENDAR)
        items = cal.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        # Outlook's Restrict() uses local time, formatted in the user's
        # short date/time format. Convert UTC window to local.
        local_tz = _dt.datetime.now().astimezone().tzinfo
        start_local = _aware_utc(window_start_utc).astimezone(local_tz)
        end_local = _aware_utc(window_end_utc).astimezone(local_tz)
        fmt = "%m/%d/%Y %I:%M %p"
        restriction = (
            f"[Start] <= '{end_local.strftime(fmt)}' AND "
            f"[End] >= '{start_local.strftime(fmt)}'"
        )

        try:
            filtered = items.Restrict(restriction)
        except Exception as e:
            LOG.warning("Outlook Restrict() failed (%s); iterating fully.", e)
            filtered = items

        out: List[MeetingEvent] = []
        # Restrict() returns a collection whose .Count is INT_MAX (a lazy-eval
        # sentinel). We must iterate by index and stop when Item(i) returns None.
        i = 1
        while True:
            try:
                item = filtered.Item(i)
            except Exception:
                break
            if item is None:
                break
            try:
                ev = _parse_item(item)
                if ev is not None:
                    out.append(ev)
            except Exception as e:
                LOG.debug("Skipping calendar item %d: %s", i, e)
            i += 1
            if i > 5000:  # safety net against an infinite loop
                LOG.warning("Calendar iteration hit safety limit at 5000 items.")
                break
        return out


def _aware_utc(dt: _dt.datetime) -> _dt.datetime:
    if dt.tzinfo is None:
        return dt.replace(tzinfo=_dt.timezone.utc)
    return dt.astimezone(_dt.timezone.utc)


def _parse_item(item) -> Optional[MeetingEvent]:
    # MessageClass is typically "IPM.Appointment" or "IPM.Schedule.Meeting.*".
    msg_cls = ""
    try:
        msg_cls = str(item.MessageClass or "")
    except Exception:
        pass
    if msg_cls and not msg_cls.startswith("IPM.Appointment") and not msg_cls.startswith(
        "IPM.Schedule.Meeting"
    ):
        return None

    subject = _safe_attr(item, "Subject", "") or ""
    start = _safe_attr(item, "Start", None)
    end = _safe_attr(item, "End", None)
    if start is None or end is None:
        return None
    all_day = bool(_safe_attr(item, "AllDayEvent", False))
    categories_raw = _safe_attr(item, "Categories", "") or ""
    categories = [c.strip() for c in categories_raw.split(",") if c.strip()]
    body = _safe_attr(item, "Body", "") or ""
    location = _safe_attr(item, "Location", "") or ""
    entry_id = _safe_attr(item, "EntryID", None) or _safe_attr(
        item, "GlobalAppointmentID", None
    )
    if not entry_id:
        # Last-resort identity.
        entry_id = f"{subject}|{start}"

    join_url = extract_teams_url(body, location)
    # Treat as an online meeting if we found a URL, OR if Outlook stamped the
    # location with "Microsoft Teams Meeting" (the URL may live in RTF/MAPI
    # named props we can't easily read — we still want to nag).
    is_online_meeting = bool(join_url) or ("microsoft teams" in (location or "").lower())

    return MeetingEvent(
        event_id=str(entry_id),
        subject=subject,
        start=_to_utc_naive(start),
        end=_to_utc_naive(end),
        is_all_day=all_day,
        is_online_meeting=is_online_meeting,
        join_url=join_url,
        categories=categories,
    )


def _safe_attr(item, name: str, default):
    try:
        return getattr(item, name)
    except Exception:
        return default


def _to_utc_naive(value) -> _dt.datetime:
    """Convert a pywintypes.datetime / datetime / date to naive UTC datetime.

    pywin32 returns COM DATE values tagged with `+00:00` tzinfo even though the
    underlying value is local time. We strip the bogus tzinfo, reattach the
    system local timezone, then convert to UTC.
    """
    if isinstance(value, _dt.datetime):
        naive_local = value.replace(tzinfo=None)
    elif isinstance(value, _dt.date):
        naive_local = _dt.datetime(value.year, value.month, value.day)
    else:
        # Last-resort coercion for unexpected types.
        naive_local = _dt.datetime.fromisoformat(str(value)).replace(tzinfo=None)
    local_tz = _dt.datetime.now().astimezone().tzinfo
    return (
        naive_local.replace(tzinfo=local_tz)
        .astimezone(_dt.timezone.utc)
        .replace(tzinfo=None)
    )


def filter_events(
    events: List[MeetingEvent],
    *,
    only_online_meetings: bool,
    ignore_all_day: bool,
    ignore_categories: List[str],
    ignore_subjects_contains: List[str],
) -> List[MeetingEvent]:
    cats_lower = {c.lower() for c in ignore_categories or []}
    subs_lower = [s.lower() for s in ignore_subjects_contains or []]
    out: List[MeetingEvent] = []
    for ev in events:
        if ignore_all_day and ev.is_all_day:
            continue
        if only_online_meetings and not ev.is_online_meeting:
            continue
        if cats_lower and any((c or "").lower() in cats_lower for c in ev.categories):
            continue
        if subs_lower:
            s = (ev.subject or "").lower()
            if any(needle in s for needle in subs_lower):
                continue
        out.append(ev)
    return out
