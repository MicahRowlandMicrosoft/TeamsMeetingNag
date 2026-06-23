"""Shared data types for Teams Meeting Nag."""
from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import List, Optional

# Teams join URLs come in two main shapes:
#   - https://teams.microsoft.com/l/meetup-join/...   (legacy / federated)
#   - https://teams.microsoft.com/meet/<id>?p=<token>  (modern Teams)
# Match either.
TEAMS_URL_RE = re.compile(
    r"https://teams\.microsoft\.com/(?:l/meetup-join/|meet/)[^\s\"'<>]+",
    re.IGNORECASE,
)


@dataclass
class MeetingEvent:
    """A calendar event we might nag about. Times are naive UTC datetimes."""

    event_id: str
    subject: str
    start: _dt.datetime
    end: _dt.datetime
    is_all_day: bool = False
    is_online_meeting: bool = False
    join_url: Optional[str] = None
    categories: List[str] = field(default_factory=list)


def extract_teams_url(*texts: str) -> Optional[str]:
    """Return the first Teams meet-up URL found in any of the given texts."""
    for t in texts:
        if not t:
            continue
        m = TEAMS_URL_RE.search(t)
        if m:
            return m.group(0)
    return None
