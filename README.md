# Teams Meeting Nag

ADHD-friendly Windows nag for Microsoft Teams meetings. If a meeting has
started and you're not joined yet, it fires an actionable toast + sound every
30 seconds until you join, mark "I'm in it", snooze, or dismiss.

## How it works

- **Calendar:** reads your Outlook desktop calendar via COM (no sign-in - uses
  your already-signed-in Outlook profile).
- **"Am I joined?" detection:**
  - **This PC:** enumerates Teams windows; if a window titled with the
    meeting subject (e.g. `Sprint planning | Microsoft Teams`) is open, the
    nag stays quiet.
  - **Other devices (phone, web, another PC):** click the **I'm in it**
    button on the toast - that suppresses alerts for the rest of the
    meeting.

## Why no Microsoft Graph?

The original plan used Graph's `/me/presence` for true cross-device
detection. That requires either an app registration or a preauthorized
first-party client. In the Microsoft corp tenant every public client we
tried (Graph PowerShell, VS Code, Office, Graph Explorer, Azure CLI) is
either blocked or lacks the right scopes, and self-service app registration
isn't available. Local-only mode is the workaround.

## Requirements

- Windows 10 / 11
- Python 3.11+ (**3.14 recommended** - on win-arm64 it's the only build that
  bundles `tkinter`, which the settings UI needs)
- Outlook desktop with a configured mail profile (the same one you use
  every day)
- Teams desktop (new or classic)

## Run

```powershell
cd C:\source\TeamsMeetingNag
.\run.ps1
```

First run creates a venv and installs `windows-toasts` + `pywin32`. After
that it's instant.

Useful flags:

```powershell
.\run.ps1 -Config             # Open the settings UI (see below)
.\run.ps1 -Reinstall          # Force pip reinstall
```

The underlying script also accepts:

```powershell
.\.venv\Scripts\python.exe nag.py --probe                # List upcoming events and exit
.\.venv\Scripts\python.exe nag.py --list-teams-windows   # Dump Teams window titles
```

### Auto-start at login (optional)

1. `Win`+`R` → `shell:startup` → Enter.
2. Right-click → **New → Shortcut**.
3. Target:
   ```
   powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\source\TeamsMeetingNag\run.ps1"
   ```

## Configuration

Easiest: run `.\run.ps1 -Config` to open a small Tk window with every
setting, inline help, validation, and a sound picker. Changes are written
back to [config.json](config.json) (your `// ...` comment keys are
preserved) and take effect the next time `nag.py` starts.

Or edit [config.json](config.json) by hand. Keys starting with `//` are
comments and are ignored.

| Key | Default | What it does |
|---|---|---|
| `poll_interval_seconds` | `15` | How often to check Outlook + Teams windows |
| `prealert_seconds` | `0` | Start nagging this many seconds before scheduled start |
| `renag_seconds` | `30` | Re-fire toast + sound every N seconds |
| `stop_after_minutes` | `10` | Give up nagging this long after start |
| `snooze_seconds` | `60` | How long the **Snooze 1 min** button suppresses alerts |
| `only_online_meetings` | `true` | Skip events without a Teams join link |
| `ignore_all_day` | `true` | Skip all-day events |
| `ignore_categories` | `["No Nag"]` | Skip events with these Outlook categories |
| `ignore_subjects_contains` | `["Focus time", "Lunch"]` | Skip events whose subject contains these substrings |
| `local_presence_enabled` | `true` | Suppress alerts when a Teams meeting window for the subject is open |
| `sound` | `"beep"` | `"beep"`, a path to a `.wav` file, or `""` to mute |

### Pro tip for ADHD

- Add an Outlook category called **No Nag** and apply it to tentative or
  optional meetings you don't want to be pestered about.
- Drop a louder `.wav` in this folder (try a doorbell or alarm clip) and set
  `"sound": "alarm.wav"` for higher-urgency events.

## Toast buttons

| Button | What it does |
|---|---|
| **Join** | Opens the meeting's Teams join URL |
| **I'm in it** | Marks the meeting as joined elsewhere - silences until end + 1 min |
| **Snooze 1 min** | Stops nagging for `snooze_seconds` |
| **Dismiss** | Stops nagging for this meeting permanently (until next occurrence) |

## Files

- [nag.py](nag.py) - main loop
- [outlook_calendar.py](outlook_calendar.py) - Outlook COM calendar reader
- [local_presence.py](local_presence.py) - Teams window detection
- [notifier.py](notifier.py) - toast + sound
- [models.py](models.py) - shared `MeetingEvent` dataclass
- [config.json](config.json) - tunables
- [config_ui.py](config_ui.py) - Tk settings editor (`run.ps1 -Config`)
- [run.ps1](run.ps1) - venv bootstrap + launcher

## Troubleshooting

- **Outlook prompts you with a security dialog about programmatic access**:
  this happens if your antivirus / Outlook security policy intercepts COM.
  Allow the access (you can pick "Allow for 10 minutes" repeatedly, or
  configure Outlook trust center to allow programmatic calendar reads).
- **Nag fires while I'm already in the meeting on this PC**: run
  `.\run.ps1 -ListTeamsWindows` while in the meeting and tell me what titles
  Teams is showing. The window-title heuristic in
  [local_presence.py](local_presence.py) may need tuning.
- **Nag fires while I'm on my phone**: that's the inherent local-only
  limitation. Click **I'm in it** on the toast.
- **Toast doesn't show buttons**: Windows Focus Assist may be silencing
  notifications. Turn off "Do not disturb" or whitelist the app.
- **`No usable Python found`**: install Python 3.14 from
  <https://python.org/downloads/windows/> (ARM64 build for Snapdragon PCs).
- **`-Config` says it's rebuilding the venv**: that's expected if your
  current venv was built on a Python without `tkinter` (notably 3.13
  win-arm64). It will re-create `.venv` on a Python that has it and
  reinstall the deps. Subsequent runs are fast.
