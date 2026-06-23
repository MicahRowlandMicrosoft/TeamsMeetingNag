"""Teams Meeting Nag — main loop (Plan B: local-only).

Every poll_interval_seconds:
  1. Read upcoming/current calendar events via Outlook COM.
  2. For each event in the nag window:
       - If a Teams meeting window for that subject is open on this PC → silent.
       - Else fire (or re-fire) a toast + sound until the user joins, snoozes,
         dismisses, or marks "I'm in it" (joined on another device).

Run:
    python nag.py
No sign-in needed; uses your already-signed-in Outlook desktop client.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import logging
import signal
import sys
import time
import webbrowser
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Set

from models import MeetingEvent
from notifier import AlertAction, Alerter
from outlook_calendar import OutlookCalendar, filter_events
from local_presence import is_in_local_teams_meeting

LOG = logging.getLogger("nag")
CONFIG_PATH = Path(__file__).with_name("config.json")


def _utc_now() -> _dt.datetime:
    """Naive UTC datetime (tz-aware UTC stripped of tzinfo)."""
    return _dt.datetime.now(_dt.timezone.utc).replace(tzinfo=None)


@dataclass
class MeetingState:
    last_alerted_at: float = 0.0
    snooze_until: float = 0.0
    dismissed: bool = False


@dataclass
class AppState:
    meetings: Dict[str, MeetingState] = field(default_factory=dict)

    def get(self, event_id: str) -> MeetingState:
        st = self.meetings.get(event_id)
        if st is None:
            st = MeetingState()
            self.meetings[event_id] = st
        return st

    def prune(self, active_ids: Set[str]) -> None:
        for key in list(self.meetings.keys()):
            if key not in active_ids:
                del self.meetings[key]


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"Missing config at {CONFIG_PATH}")
    raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    # Strip comment keys ("// ...").
    return {k: v for k, v in raw.items() if not k.startswith("//")}


def format_alert(ev: MeetingEvent, now: _dt.datetime) -> tuple[str, str]:
    delta = (now - ev.start).total_seconds()
    if delta < 0:
        when = f"starts in {int(-delta // 60)} min"
    elif delta < 60:
        when = "starting now"
    else:
        when = f"started {int(delta // 60)} min ago"
    title = f"🔔 Meeting: {ev.subject}"
    body = f"{when} — you haven't joined yet."
    return title, body


def run(config: dict) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    poll_s = int(config.get("poll_interval_seconds", 15))
    prealert_s = int(config.get("prealert_seconds", 0))
    renag_s = int(config.get("renag_seconds", 30))
    stop_after_min = int(config.get("stop_after_minutes", 10))
    snooze_s = int(config.get("snooze_seconds", 60))
    only_online = bool(config.get("only_online_meetings", True))
    ignore_all_day = bool(config.get("ignore_all_day", True))
    ignore_cats = list(config.get("ignore_categories", []))
    ignore_subs = list(config.get("ignore_subjects_contains", []))
    local_presence_enabled = bool(config.get("local_presence_enabled", True))

    calendar = OutlookCalendar()
    LOG.info("Reading calendar via Outlook COM...")
    # Probe once so we fail fast if Outlook isn't installed/running.
    try:
        now_utc = _utc_now()
        calendar.get_events(
            now_utc - _dt.timedelta(minutes=5),
            now_utc + _dt.timedelta(minutes=5),
        )
        LOG.info("Outlook calendar OK.")
    except Exception as e:
        LOG.error("Could not read Outlook calendar: %s", e)
        LOG.error("Make sure Outlook is installed and a mail profile is configured.")
        raise

    alerter = Alerter(sound=str(config.get("sound", "beep")))
    state = AppState()

    stop = {"flag": False}

    def _handle_signal(_sig, _frame):
        LOG.info("Stopping...")
        stop["flag"] = True

    signal.signal(signal.SIGINT, _handle_signal)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, _handle_signal)

    LOG.info(
        "Watching: poll=%ss prealert=%ss renag=%ss stop_after=%smin local_presence=%s",
        poll_s, prealert_s, renag_s, stop_after_min, local_presence_enabled,
    )

    while not stop["flag"]:
        try:
            tick(
                calendar=calendar,
                alerter=alerter,
                state=state,
                only_online=only_online,
                ignore_all_day=ignore_all_day,
                ignore_cats=ignore_cats,
                ignore_subs=ignore_subs,
                local_presence_enabled=local_presence_enabled,
                prealert_s=prealert_s,
                renag_s=renag_s,
                stop_after_min=stop_after_min,
                snooze_s=snooze_s,
            )
        except Exception as e:
            LOG.exception("Tick failed: %s", e)
        # Sleep in small slices so Ctrl+C is responsive.
        for _ in range(poll_s * 10):
            if stop["flag"]:
                break
            time.sleep(0.1)


def tick(
    *,
    calendar: OutlookCalendar,
    alerter: Alerter,
    state: AppState,
    only_online: bool,
    ignore_all_day: bool,
    ignore_cats: list,
    ignore_subs: list,
    local_presence_enabled: bool,
    prealert_s: int,
    renag_s: int,
    stop_after_min: int,
    snooze_s: int,
) -> None:
    now_utc = _utc_now()
    # Pull a small window: from 15 min ago (catch in-progress) to 15 min ahead.
    window_start = now_utc - _dt.timedelta(minutes=15)
    window_end = now_utc + _dt.timedelta(minutes=15)

    raw_events = calendar.get_events(window_start, window_end)
    events = filter_events(
        raw_events,
        only_online_meetings=only_online,
        ignore_all_day=ignore_all_day,
        ignore_categories=ignore_cats,
        ignore_subjects_contains=ignore_subs,
    )
    state.prune({e.event_id for e in events})

    if not events:
        return

    # ev.start / ev.end are naive UTC, so compare against now_utc (also naive UTC).
    actionable = [
        ev for ev in events
        if _in_nag_window(ev, now_utc, prealert_s, stop_after_min)
    ]
    if not actionable:
        return

    now_ts = time.time()
    for ev in actionable:
        st = state.get(ev.event_id)
        if st.dismissed:
            continue
        if now_ts < st.snooze_until:
            continue
        if (now_ts - st.last_alerted_at) < renag_s and st.last_alerted_at > 0:
            continue
        # Local presence: if a Teams meeting window for this subject is open
        # here, stay quiet.
        if local_presence_enabled and is_in_local_teams_meeting(ev.subject):
            LOG.info("Teams meeting window detected for '%s' — suppressing.", ev.subject)
            # Re-check after one renag interval rather than spamming the
            # window enum every poll.
            st.snooze_until = now_ts + renag_s
            continue
        st.last_alerted_at = now_ts
        title, body = format_alert(ev, now_utc)
        LOG.info("Alerting: %s", ev.subject)
        actions = _build_actions(ev=ev, st=st, snooze_s=snooze_s)
        alerter.alert(title=title, body=body, actions=actions, join_url=ev.join_url)


def _in_nag_window(
    ev: MeetingEvent,
    now: _dt.datetime,
    prealert_s: int,
    stop_after_min: int,
) -> bool:
    fire_start = ev.start - _dt.timedelta(seconds=prealert_s)
    fire_end = ev.start + _dt.timedelta(minutes=stop_after_min)
    # Also bound by meeting end so we don't nag for an event that already ended.
    fire_end = min(fire_end, ev.end)
    return fire_start <= now <= fire_end


def _build_actions(*, ev: MeetingEvent, st: MeetingState, snooze_s: int) -> AlertAction:
    def _join() -> None:
        if ev.join_url:
            try:
                webbrowser.open(ev.join_url)
            except Exception as e:
                LOG.warning("Failed to open join URL: %s", e)
        # Mark as alerted-recent so we don't re-fire immediately while Teams
        # spins up; local presence will normally take over within ~10s.
        st.snooze_until = time.time() + 30

    def _snooze() -> None:
        st.snooze_until = time.time() + snooze_s
        LOG.info("Snoozed '%s' for %ds.", ev.subject, snooze_s)

    def _dismiss() -> None:
        st.dismissed = True
        LOG.info("Dismissed '%s'.", ev.subject)

    def _joined_elsewhere() -> None:
        # Suppress for the remainder of the meeting plus a small buffer.
        end_ts = ev.end.replace(tzinfo=_dt.timezone.utc).timestamp() + 60
        st.snooze_until = max(st.snooze_until, end_ts)
        LOG.info("Marked '%s' as joined elsewhere — suppressing until meeting end.", ev.subject)

    return AlertAction(
        join=_join, snooze=_snooze, dismiss=_dismiss, joined_elsewhere=_joined_elsewhere
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Teams Meeting Nag (local mode)")
    parser.add_argument(
        "--probe", action="store_true",
        help="Connect to Outlook, list upcoming events, and exit.",
    )
    parser.add_argument(
        "--list-teams-windows", action="store_true",
        help="Print every visible window that mentions 'Microsoft Teams' and exit.",
    )
    args = parser.parse_args(argv)

    if args.list_teams_windows:
        from local_presence import debug_list_teams_windows
        titles = debug_list_teams_windows()
        if not titles:
            print("(no Teams windows found)")
        for t in titles:
            print(repr(t))
        return 0

    config = load_config()

    if args.probe:
        logging.basicConfig(level=logging.INFO, format="%(message)s")
        cal = OutlookCalendar()
        now_utc = _utc_now()
        events = cal.get_events(
            now_utc - _dt.timedelta(hours=1),
            now_utc + _dt.timedelta(hours=8),
        )
        if not events:
            print("(no events in the next 8 hours)")
            return 0
        for ev in events:
            online = " [online]" if ev.is_online_meeting else ""
            print(f"{ev.start:%H:%M}-{ev.end:%H:%M}  {ev.subject}{online}")
        return 0

    run(config)
    return 0


if __name__ == "__main__":
    sys.exit(main())
