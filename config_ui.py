"""Tkinter UI for editing config.json.

Run:
    python config_ui.py
or:
    .\\run.ps1 -Config

Preserves any "// ..." comment keys in config.json so the on-disk file
stays self-documenting.
"""
from __future__ import annotations

import json
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict

CONFIG_PATH = Path(__file__).with_name("config.json")

# (key, label, kind, help) - kind in {"int","bool","str","list","sound"}.
FIELDS = [
    ("poll_interval_seconds", "Poll interval (seconds)", "int",
     "How often to check the Outlook calendar and local Teams windows."),
    ("prealert_seconds", "Pre-alert (seconds before start)", "int",
     "Start nagging this many seconds BEFORE meeting start. 0 = exactly at start."),
    ("renag_seconds", "Re-nag interval (seconds)", "int",
     "Re-fire toast + sound every N seconds while you remain un-joined."),
    ("stop_after_minutes", "Stop after (minutes)", "int",
     "Give up nagging this many minutes after scheduled start."),
    ("snooze_seconds", "Snooze duration (seconds)", "int",
     "If you click Snooze, suppress alerts for this meeting for this long."),
    ("only_online_meetings", "Only nag for Teams online meetings", "bool",
     "If off, nag for ALL calendar events."),
    ("ignore_all_day", "Ignore all-day events", "bool", ""),
    ("local_presence_enabled", "Suppress when Teams window is open here", "bool",
     "If on, skip alerts when a Teams meeting window for the event's subject "
     "is open on this PC."),
    ("ignore_categories", "Ignore Outlook categories (one per line)", "list",
     "Skip events tagged with any of these Outlook categories."),
    ("ignore_subjects_contains", "Ignore subjects containing (one per line)", "list",
     "Skip events whose subject contains any of these substrings."),
    ("sound", "Sound", "sound",
     "Path to a .wav file, 'beep' for a system beep, or empty to disable sound."),
]

DEFAULTS: Dict[str, Any] = {
    "poll_interval_seconds": 15,
    "prealert_seconds": 120,
    "renag_seconds": 30,
    "stop_after_minutes": 10,
    "snooze_seconds": 60,
    "only_online_meetings": True,
    "ignore_all_day": True,
    "ignore_categories": ["No Nag"],
    "ignore_subjects_contains": ["Focus time", "Lunch"],
    "local_presence_enabled": True,
    "sound": "beep",
}


def load_raw() -> Dict[str, Any]:
    if not CONFIG_PATH.exists():
        return dict(DEFAULTS)
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        messagebox.showerror(
            "Invalid config.json",
            f"Could not parse {CONFIG_PATH.name}:\n\n{e}\n\nLoading defaults.",
        )
        return dict(DEFAULTS)


def write_raw(raw: Dict[str, Any]) -> None:
    CONFIG_PATH.write_text(
        json.dumps(raw, indent=4, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


class ConfigApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Teams Meeting Nag - Configuration")
        self.geometry("640x720")
        self.minsize(560, 600)

        self._raw = load_raw()
        self._vars: Dict[str, Any] = {}
        self._list_widgets: Dict[str, tk.Text] = {}

        self._build_ui()
        self._load_values()

    def _build_ui(self) -> None:
        try:
            ttk.Style().theme_use("vista")
        except tk.TclError:
            pass

        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        header = ttk.Label(
            outer,
            text="Edit settings, then click Save. Changes apply next time nag.py starts.",
            wraplength=600,
        )
        header.pack(anchor="w", pady=(0, 8))

        # Scrollable area for the fields.
        canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="top", fill="both", expand=True)

        body = ttk.Frame(canvas)
        body_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_body_configure(_evt: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(evt: tk.Event) -> None:
            canvas.itemconfigure(body_id, width=evt.width)

        body.bind("<Configure>", _on_body_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # Mouse wheel scroll while the cursor is over the canvas.
        def _on_mousewheel(evt: tk.Event) -> None:
            canvas.yview_scroll(int(-1 * (evt.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        for key, label, kind, help_text in FIELDS:
            self._add_field(body, key, label, kind, help_text)

        # Buttons row.
        btn_row = ttk.Frame(outer)
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="Restore defaults", command=self._restore_defaults).pack(side="left")
        ttk.Button(btn_row, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(btn_row, text="Save", command=self._save).pack(side="right", padx=(0, 6))

        self.bind("<Control-s>", lambda _e: self._save())
        self.bind("<Escape>", lambda _e: self.destroy())

    def _add_field(
        self, parent: ttk.Frame, key: str, label: str, kind: str, help_text: str
    ) -> None:
        frame = ttk.Frame(parent, padding=(0, 6))
        frame.pack(fill="x", expand=True)

        ttk.Label(frame, text=label, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        if help_text:
            ttk.Label(
                frame, text=help_text, foreground="#555", wraplength=580,
            ).pack(anchor="w", pady=(0, 2))

        if kind == "int":
            var = tk.StringVar()
            ttk.Spinbox(
                frame, from_=0, to=86400, increment=1, textvariable=var, width=10,
            ).pack(anchor="w")
            self._vars[key] = ("int", var)
        elif kind == "bool":
            var = tk.BooleanVar()
            ttk.Checkbutton(frame, variable=var, text="Enabled").pack(anchor="w")
            self._vars[key] = ("bool", var)
        elif kind == "str":
            var = tk.StringVar()
            ttk.Entry(frame, textvariable=var).pack(fill="x")
            self._vars[key] = ("str", var)
        elif kind == "list":
            txt = tk.Text(frame, height=4, width=40, wrap="none")
            txt.pack(fill="x")
            self._list_widgets[key] = txt
            self._vars[key] = ("list", txt)
        elif kind == "sound":
            self._build_sound_field(frame, key)
        else:
            raise ValueError(f"Unknown field kind: {kind}")

    def _build_sound_field(self, frame: ttk.Frame, key: str) -> None:
        mode_var = tk.StringVar(value="beep")
        path_var = tk.StringVar(value="")

        mode_row = ttk.Frame(frame)
        mode_row.pack(fill="x")
        for label, val in (("System beep", "beep"), ("Silent", "silent"), ("Custom .wav", "file")):
            ttk.Radiobutton(
                mode_row, text=label, value=val, variable=mode_var,
                command=lambda: self._sync_sound_state(mode_var, path_var, file_entry, browse_btn),
            ).pack(side="left", padx=(0, 8))

        path_row = ttk.Frame(frame)
        path_row.pack(fill="x", pady=(4, 0))
        file_entry = ttk.Entry(path_row, textvariable=path_var)
        file_entry.pack(side="left", fill="x", expand=True)
        browse_btn = ttk.Button(path_row, text="Browse...", command=lambda: self._browse_wav(path_var))
        browse_btn.pack(side="left", padx=(6, 0))

        self._vars[key] = ("sound", mode_var, path_var, file_entry, browse_btn)
        self._sync_sound_state(mode_var, path_var, file_entry, browse_btn)

    def _sync_sound_state(
        self,
        mode_var: tk.StringVar,
        _path_var: tk.StringVar,
        file_entry: ttk.Entry,
        browse_btn: ttk.Button,
    ) -> None:
        state = "normal" if mode_var.get() == "file" else "disabled"
        file_entry.configure(state=state)
        browse_btn.configure(state=state)

    def _browse_wav(self, path_var: tk.StringVar) -> None:
        initial = path_var.get() or str(Path.home())
        path = filedialog.askopenfilename(
            title="Pick a .wav file",
            initialdir=str(Path(initial).parent if Path(initial).parent.exists() else Path.home()),
            filetypes=[("WAV audio", "*.wav"), ("All files", "*.*")],
        )
        if path:
            path_var.set(path)

    def _load_values(self) -> None:
        for key, _label, _kind, _help in FIELDS:
            current = self._raw.get(key, DEFAULTS.get(key))
            self._set_field(key, current)

    def _set_field(self, key: str, value: Any) -> None:
        meta = self._vars[key]
        kind = meta[0]
        if kind == "int":
            meta[1].set(str(int(value or 0)))
        elif kind == "bool":
            meta[1].set(bool(value))
        elif kind == "str":
            meta[1].set("" if value is None else str(value))
        elif kind == "list":
            txt = meta[1]
            txt.delete("1.0", "end")
            if isinstance(value, list):
                txt.insert("1.0", "\n".join(str(x) for x in value))
        elif kind == "sound":
            _, mode_var, path_var, file_entry, browse_btn = meta
            v = "" if value is None else str(value)
            if v == "" or v.lower() in {"none", "off", "silent"}:
                mode_var.set("silent")
                path_var.set("")
            elif v.lower() == "beep":
                mode_var.set("beep")
                path_var.set("")
            else:
                mode_var.set("file")
                path_var.set(v)
            self._sync_sound_state(mode_var, path_var, file_entry, browse_btn)

    def _collect_values(self) -> Dict[str, Any]:
        out: Dict[str, Any] = {}
        for key, label, _kind, _help in FIELDS:
            meta = self._vars[key]
            kind = meta[0]
            if kind == "int":
                raw = meta[1].get().strip()
                try:
                    out[key] = int(raw) if raw else 0
                except ValueError:
                    raise ValueError(f"'{label}' must be a whole number (got {raw!r}).")
                if out[key] < 0:
                    raise ValueError(f"'{label}' must be 0 or greater.")
            elif kind == "bool":
                out[key] = bool(meta[1].get())
            elif kind == "str":
                out[key] = meta[1].get().strip()
            elif kind == "list":
                lines = [ln.strip() for ln in meta[1].get("1.0", "end").splitlines()]
                out[key] = [ln for ln in lines if ln]
            elif kind == "sound":
                _, mode_var, path_var, *_ = meta
                mode = mode_var.get()
                if mode == "beep":
                    out[key] = "beep"
                elif mode == "silent":
                    out[key] = ""
                else:
                    path = path_var.get().strip()
                    if not path:
                        raise ValueError("Pick a .wav file or choose System beep / Silent.")
                    out[key] = path
        return out

    def _restore_defaults(self) -> None:
        if not messagebox.askyesno(
            "Restore defaults",
            "Reset every field to its default value? (Not saved until you click Save.)",
        ):
            return
        for key, _label, _kind, _help in FIELDS:
            self._set_field(key, DEFAULTS.get(key))

    def _save(self) -> None:
        try:
            values = self._collect_values()
        except ValueError as e:
            messagebox.showerror("Invalid value", str(e))
            return

        # Merge into raw to preserve comment keys ("// ...") and any unknown keys.
        merged = dict(self._raw)
        for k, v in values.items():
            merged[k] = v
        try:
            write_raw(merged)
        except OSError as e:
            messagebox.showerror("Save failed", f"Could not write {CONFIG_PATH}:\n\n{e}")
            return
        self._raw = merged
        messagebox.showinfo(
            "Saved",
            f"Saved to {CONFIG_PATH.name}.\nRestart nag.py for changes to take effect.",
        )


def main() -> int:
    app = ConfigApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
