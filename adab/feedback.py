#!/usr/bin/env python3
"""feedback.py — a drop-in "Send feedback" button for any tkinter tool.

Goal: never take notes by hand again. Every tool gets a small feedback button;
when a user clicks it they type a bug / idea / wrong-result note, and it is
stored in ONE shared database. Users can only SUBMIT — they never see anyone
else's feedback. You (the developer) read the database later to plan the next
version.

Use it in three lines:

    from feedback import add_feedback_button
    ...
    add_feedback_button(root, app="DocuBOM", version="1.0")

Where the feedback goes (first that works wins):
    1. the db_path you pass to add_feedback_button(...)
    2. the environment variable  BG_FEEDBACK_DB  (point this at a network share
       so ALL machines log to the same file, e.g. \\\\server\\quality\\feedback.db)
    3. ~/.beyondgravity/feedback.db   (per-user local fallback)

Nothing leaves the machine except to that database file. No internet needed.
Every entry is ALSO appended to a .jsonl next to the DB as a belt-and-braces
backup, so a note is never lost even if the DB is briefly locked or offline.
"""
import getpass
import json
import os
import socket
import sqlite3
import datetime

APP_ENV = "BG_FEEDBACK_DB"
CATEGORIES = ["Bug / something broke", "Idea / improvement",
              "Wrong result", "Question", "Other"]


# ----------------------------------------------------------------------------
# storage — plain SQLite, no dependencies
# ----------------------------------------------------------------------------
def default_db_path():
    env = os.environ.get(APP_ENV)
    if env:
        return env
    return os.path.join(os.path.expanduser("~"), ".beyondgravity",
                        "feedback.db")


def _connect(db_path):
    os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
    conn = sqlite3.connect(db_path, timeout=10)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS feedback (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            ts         TEXT,
            app        TEXT,
            version    TEXT,
            category   TEXT,
            message    TEXT,
            user       TEXT,
            machine    TEXT,
            context    TEXT
        )""")
    return conn


def submit_feedback(app, version, category, message, context=None,
                    db_path=None):
    """Store one feedback entry. Returns the path it was written to.
    Falls back to a local DB if the primary path can't be written, and always
    also appends a JSONL backup line so nothing is ever lost."""
    ts = datetime.datetime.now().isoformat(timespec="seconds")
    user = _safe(getpass.getuser)
    machine = _safe(socket.gethostname)
    ctx = json.dumps(context, ensure_ascii=False) if context else ""
    row = (ts, str(app), str(version or ""), str(category or ""),
           str(message or ""), user, machine, ctx)

    target = db_path or default_db_path()
    written = None
    for path in [target, os.path.join(os.path.expanduser("~"),
                                      ".beyondgravity", "feedback.db")]:
        try:
            conn = _connect(path)
            conn.execute(
                "INSERT INTO feedback (ts,app,version,category,message,user,"
                "machine,context) VALUES (?,?,?,?,?,?,?,?)", row)
            conn.commit(); conn.close()
            written = path
            break
        except Exception:
            continue

    # JSONL backup next to whichever DB we used (or the target if none worked)
    try:
        base = written or target
        jsonl = os.path.splitext(base)[0] + ".jsonl"
        os.makedirs(os.path.dirname(os.path.abspath(jsonl)), exist_ok=True)
        with open(jsonl, "a", encoding="utf-8") as f:
            f.write(json.dumps({
                "ts": ts, "app": app, "version": version,
                "category": category, "message": message, "user": user,
                "machine": machine, "context": context}, ensure_ascii=False)
                + "\n")
    except Exception:
        pass

    if written is None:
        raise IOError("Could not save feedback to any database location.")
    return written


def _safe(fn, default="unknown"):
    try:
        return str(fn())
    except Exception:
        return default


# ----------------------------------------------------------------------------
# the button + dialog (tkinter)
# ----------------------------------------------------------------------------
def add_feedback_button(parent, app, version="", db_path=None,
                        context_provider=None, text="Feedback"):
    """Add a small feedback button to a tkinter window.

    parent            : the Tk root or a Frame to place the button in
    app               : tool name, e.g. "DocuBOM"
    version           : optional version string
    db_path           : optional explicit DB path (else env / local default)
    context_provider  : optional callable returning a dict of current state
                        (e.g. the file being processed) attached to the note
    """
    import tkinter as tk

    def open_dialog():
        ctx = None
        if context_provider:
            try:
                ctx = context_provider()
            except Exception:
                ctx = None
        _FeedbackDialog(parent, app, version, db_path, ctx)

    btn = tk.Button(parent, text="  " + text + "  ", command=open_dialog,
                    relief="groove")
    return btn


class _FeedbackDialog:
    def __init__(self, parent, app, version, db_path, context):
        import tkinter as tk
        self.app, self.version, self.db_path = app, version, db_path
        self.context = context
        self.win = tk.Toplevel(parent)
        self.win.title(f"Send feedback — {app}")
        self.win.geometry("460x400"); self.win.resizable(False, False)
        self.win.transient(parent); self.win.grab_set()

        tk.Label(self.win, text=f"Help improve {app}",
                 font=("Segoe UI", 13, "bold"), fg="#1F3864").pack(
            anchor="w", padx=14, pady=(12, 0))
        tk.Label(self.win, text="Found a bug, a wrong result, or have an idea? "
                 "Write it here — it goes straight to the team for the next "
                 "version. You won't need to email anyone.",
                 font=("Segoe UI", 9), fg="#555", justify="left",
                 wraplength=430).pack(anchor="w", padx=14, pady=(2, 8))

        tk.Label(self.win, text="Type:", font=("Segoe UI", 9, "bold")).pack(
            anchor="w", padx=14)
        self.cat = tk.StringVar(value=CATEGORIES[0])
        tk.OptionMenu(self.win, self.cat, *CATEGORIES).pack(
            anchor="w", padx=12, fill="x")

        tk.Label(self.win, text="Your message:",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=14,
                                                    pady=(8, 0))
        self.text = tk.Text(self.win, height=7, wrap="word",
                            font=("Segoe UI", 10))
        self.text.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        self.text.focus_set()

        row = tk.Frame(self.win); row.pack(fill="x", padx=14, pady=(0, 12))
        self.msg = tk.Label(row, text="", font=("Segoe UI", 8), fg="#888")
        self.msg.pack(side="left")
        tk.Button(row, text="Cancel", command=self.win.destroy).pack(
            side="right")
        tk.Button(row, text="Send", bg="#1F3864", fg="white",
                  font=("Segoe UI", 10, "bold"), command=self.send).pack(
            side="right", padx=6)

    def send(self):
        import tkinter as tk
        body = self.text.get("1.0", "end").strip()
        if not body:
            self.msg.config(text="Please type something first.", fg="#c00")
            return
        try:
            where = submit_feedback(self.app, self.version, self.cat.get(),
                                    body, context=self.context,
                                    db_path=self.db_path)
        except Exception as e:
            self.msg.config(text="Could not save: " + str(e), fg="#c00")
            return
        # thank-you, then close
        for w in self.win.winfo_children():
            w.destroy()
        tk.Label(self.win, text="✓  Thank you!",
                 font=("Segoe UI", 15, "bold"), fg="#2d7d46").pack(pady=(60, 4))
        tk.Label(self.win, text="Your feedback was saved. It'll help shape the "
                 "next version.", font=("Segoe UI", 9), fg="#555",
                 wraplength=380).pack(padx=20)
        tk.Button(self.win, text="Close", command=self.win.destroy).pack(
            pady=16)
        self.win.after(2500, self.win.destroy)


if __name__ == "__main__":
    # tiny self-test window
    import tkinter as tk
    root = tk.Tk(); root.title("feedback.py demo"); root.geometry("360x160")
    tk.Label(root, text="Demo tool", font=("Segoe UI", 14)).pack(pady=20)
    add_feedback_button(root, app="DemoTool", version="1.0").pack()
    root.mainloop()
