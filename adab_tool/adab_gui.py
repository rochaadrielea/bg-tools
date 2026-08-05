#!/usr/bin/env python3
"""ADAB Compare - desktop app (As-Design vs an As-Built source).

Beyond Gravity theme (v2.1): dark "space" look, white text, bright-blue accent.
Card layout, a live progress bar with a percentage, a plain-words status line,
and a tidy (collapsible) log. Same engine and options as before.

Pick the As-Design and an As-Built source, say what KIND of source it is
(Scan / Manual list / Reserved / mb51 / Teamcenter / Find Batch), choose file or
folder, tick Combine if the files together form one list, then Run.

The source type names the two "unmatched" tabs in the report, e.g. for Reserved:
    "In Design, not in Reserved"   and   "In Reserved, not in Design"

"Find Batch" source type: instead of reconciling counts, ADAB looks up each
material of the As-Design list in the As-Built (label/scan) source and fills the
Charge (batch) + as-built Revision. Flags FOUND / MULTIPLE / NONE.

Run it with run_adab.bat, or:
    conda activate quality
    python adab_gui.py

(adab_gui.py, adab_batch_compare.py and batch_finder.py must be in the same folder.)
"""
import os
import glob
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, ttk

import adab_batch_compare as core

try:
    from feedback import add_feedback_button      # every UI tool gets this
except Exception:
    add_feedback_button = None

try:
    import batch_finder                           # the Find Batch engine
except Exception:
    batch_finder = None

# Source-type label shown in the dropdown -> short label used in the report tabs
SOURCE_TYPES = {
    "Label scanner": "Scanner",
    "Manual List": "Manual",
    "Reserved Logistic": "Reserved",
    "Team center": "Teamcenter",
    "Find Batch": "FindBatch",     # fill Charge/Rev by lookup
}

# ---------------------------------------------- Beyond Gravity palette (dark) --
BG      = "#081521"   # deep space navy (window)
CARD    = "#0F2536"   # panel / card
BORDER  = "#21455E"   # card border
INK     = "#EAF2F8"   # primary text (near white)
SUB     = "#8FA9BC"   # secondary text
ACCENT  = "#1E9BE0"   # Beyond Gravity blue
ACCENTB = "#4FC3F7"   # bright blue (the logo dot / hover-up)
ACCENTD = "#1685C4"   # darker blue (button hover)
ACCENTL = "#123246"   # blue tint (subtle highlight)
FIELD   = "#0A1B29"   # entry background
TRACK   = "#17334A"   # progress trough
LOGBG   = "#06111C"   # log background
LOGINK  = "#C7D6E2"   # log text

# header gradient stops (top -> bottom): dark space to a blue horizon glow
GRAD_TOP = (5, 12, 20)      # #050C14
GRAD_BOT = (12, 58, 92)     # #0C3A5C

FONT      = "Segoe UI"
FONT_MONO = "Consolas"


def _milestone(line):
    """Return (status_text, target_fraction 0..1) for a known log line, else None."""
    s = line.strip()
    low = s.lower()
    if low.startswith("as-design"):
        return ("Reading the As-Design baseline…", 0.10)
    if low.startswith("combine:"):
        return ("Merging the As-Built files…", 0.15)
    if low.startswith("mode: find batch"):
        return ("Looking up batches…", 0.20)
    if "roles ->" in low or low.startswith("roles"):
        return ("Checking which file is which…", 0.30)
    if "match engines" in low:
        return ("Matching part numbers…", 0.45)
    if "name engine" in low or "description engine" in low:
        return ("Comparing names…", 0.70)
    if "line conservation" in low:
        return ("Keeping every line…", 0.82)
    if "distinct parts ->" in low:
        return ("Writing the report…", 0.96)
    if low.startswith("need rows") or low.startswith("report:"):
        return ("Writing the report…", 0.96)
    if low.startswith("done") or low.startswith("finished"):
        return ("Finished", 1.0)
    return None


class App:
    def __init__(self, root):
        self.root = root
        root.title("ADAB Compare  —  Beyond Gravity")
        root.geometry("980x830")
        root.minsize(860, 730)
        root.configure(bg=BG)

        self.design_var = tk.StringVar()
        self.built_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.combine_var = tk.BooleanVar(value=False)
        self.source_var = tk.StringVar(value="Team center")

        self._pct = 0.0
        self._target = 0.0
        self._files_total = 1
        self._files_done = 0
        self._running = False
        self._log_visible = True

        self._init_styles()
        self._build()
        self._tick()

    # ------------------------------------------------------------ styling --
    def _init_styles(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("BG.Horizontal.TProgressbar",
                     troughcolor=TRACK, bordercolor=TRACK,
                     background=ACCENT, lightcolor=ACCENTB, darkcolor=ACCENT,
                     thickness=16)

    # -------------------------------------------------------------- layout --
    def _card(self, parent, pady=(0, 12)):
        outer = tk.Frame(parent, bg=BG)
        outer.pack(fill="x", padx=22, pady=pady)
        card = tk.Frame(outer, bg=CARD, highlightbackground=BORDER,
                        highlightthickness=1, bd=0)
        card.pack(fill="x")
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="x", padx=16, pady=14)
        return inner

    def _label(self, parent, text, size=9, bold=False, fg=INK, **kw):
        return tk.Label(parent, text=text, bg=parent["bg"], fg=fg,
                        font=(FONT, size, "bold" if bold else "normal"), **kw)

    def _entry(self, parent, var):
        return tk.Entry(parent, textvariable=var, font=(FONT, 10), bg=FIELD,
                        fg=INK, relief="flat", highlightbackground=BORDER,
                        highlightcolor=ACCENT, highlightthickness=1,
                        insertbackground=INK, disabledbackground=FIELD)

    def _ghost_btn(self, parent, text, cmd):
        return tk.Button(parent, text=text, command=cmd, font=(FONT, 9),
                         bg=CARD, fg=INK, activebackground=ACCENTL,
                         activeforeground=INK, relief="flat", bd=0,
                         highlightbackground=BORDER, highlightthickness=1,
                         padx=12, pady=5, cursor="hand2")

    # ------------------------------------------------------- header canvas --
    def _draw_header(self, event=None):
        c = self.header
        w = c.winfo_width() or 980
        h = int(c["height"])
        c.delete("all")
        # vertical gradient: dark space -> blue horizon glow
        steps = h
        for i in range(steps):
            t = i / max(1, steps - 1)
            r = int(GRAD_TOP[0] + (GRAD_BOT[0] - GRAD_TOP[0]) * t)
            g = int(GRAD_TOP[1] + (GRAD_BOT[1] - GRAD_TOP[1]) * t)
            b = int(GRAD_TOP[2] + (GRAD_BOT[2] - GRAD_TOP[2]) * t)
            c.create_line(0, i, w, i, fill=f"#{r:02x}{g:02x}{b:02x}")
        # soft horizon glow bottom-centre
        for rad, col in ((190, "#0E4A72"), (130, "#12688f"), (70, "#1c86b8")):
            c.create_oval(w / 2 - rad, h - rad // 2, w / 2 + rad, h + rad,
                          outline="", fill=col)
        c.create_line(0, h - 1, w, h - 1, fill=ACCENT)
        # title + brand dot
        c.create_text(24, 30, anchor="w", text="ADAB Compare",
                      fill="white", font=(FONT, 22, "bold"))
        tw = 24 + 12.2 * len("ADAB Compare")
        c.create_oval(tw + 6, 22, tw + 20, 36, outline="", fill=ACCENTB)
        c.create_text(25, 62, anchor="w",
                      text="As-Design vs As-Built traceability  ·  beyond gravity",
                      fill="#BBD4E6", font=(FONT, 10))

    def _build(self):
        # ---- header (canvas gradient) ----
        self.header = tk.Canvas(self.root, height=92, highlightthickness=0, bd=0)
        self.header.pack(fill="x")
        self.header.bind("<Configure>", self._draw_header)
        if add_feedback_button is not None:
            fb = add_feedback_button(
                self.root, app="ADAB Compare", version="2.1",
                context_provider=lambda: {"design": self.design_var.get(),
                                          "built": self.built_var.get(),
                                          "source": self.source_var.get(),
                                          "combine": self.combine_var.get()})
            fb.configure(bg=ACCENT, fg="white", relief="flat",
                         font=(FONT, 9, "bold"), cursor="hand2", bd=0,
                         padx=12, pady=6, activebackground=ACCENTD,
                         activeforeground="white")
            self.header.create_window(0, 0, window=fb, anchor="ne", tags="fb")
            self.header.bind("<Configure>",
                             lambda e: (self._draw_header(e),
                                        self.header.coords(
                                            "fb", self.header.winfo_width() - 18, 28)),
                             add="+")

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, pady=(14, 0))

        # ---- As-Design ----
        c = self._card(body)
        self._label(c, "1  ·  As-Design (the F- baseline)", bold=True,
                    size=10, fg=ACCENTB).pack(anchor="w")
        self._path_row(c, self.design_var, role="As-DESIGN")

        # ---- As-Built ----
        c = self._card(body)
        self._label(c, "2  ·  As-Built source (file or folder)", bold=True,
                    size=10, fg=ACCENTB).pack(anchor="w")
        self._path_row(c, self.built_var, role="As-BUILT source")

        self._label(c, "Report tabs are always “In As-Built, not in Design” / "
                    "“In Design, not in As-Built”.", fg=SUB, size=8).pack(
            anchor="w", pady=(8, 0))

        cbrow = tk.Frame(c, bg=CARD)
        cbrow.pack(fill="x", pady=(6, 0))
        chk = tk.Checkbutton(cbrow, variable=self.combine_var,
                             text="  Combine all files into one list (one report)",
                             bg=CARD, fg=INK, activebackground=CARD,
                             activeforeground=INK, selectcolor=FIELD,
                             font=(FONT, 10), bd=0, highlightthickness=0,
                             cursor="hand2")
        chk.pack(side="left")
        self._label(cbrow, "← tick when several files together are ONE list",
                    fg=SUB, size=8).pack(side="left", padx=(6, 0))

        # ---- Output ----
        c = self._card(body)
        self._label(c, "3  ·  Output folder (reports go here)", bold=True,
                    size=10, fg=ACCENTB).pack(anchor="w")
        self._path_row(c, self.out_var, role="OUTPUT", folder_only=True)

        # ---- Run ----
        runwrap = tk.Frame(body, bg=BG)
        runwrap.pack(fill="x", padx=22, pady=(2, 8))
        self.run_btn = tk.Button(runwrap, text="Run  ·  As-Built vs As-Design",
                                 font=(FONT, 13, "bold"), bg=ACCENT, fg="white",
                                 activebackground=ACCENTD, activeforeground="white",
                                 relief="flat", bd=0, height=2, cursor="hand2",
                                 command=self.run)
        self.run_btn.pack(fill="x")
        self.run_btn.bind("<Enter>", lambda e: (not self._running)
                          and self.run_btn.config(bg=ACCENTD))
        self.run_btn.bind("<Leave>", lambda e: (not self._running)
                          and self.run_btn.config(bg=ACCENT))

        # ---- Progress ----
        c = self._card(body)
        prow = tk.Frame(c, bg=CARD)
        prow.pack(fill="x")
        self.status_lbl = self._label(prow, "Ready.", bold=True)
        self.status_lbl.pack(side="left")
        self.pct_lbl = self._label(prow, "0%", bold=True, fg=ACCENTB, size=11)
        self.pct_lbl.pack(side="right")
        self.bar = ttk.Progressbar(c, style="BG.Horizontal.TProgressbar",
                                   mode="determinate", maximum=100, value=0)
        self.bar.pack(fill="x", pady=(8, 0))

        # ---- Log (collapsible) ----
        logcard = tk.Frame(body, bg=BG)
        logcard.pack(fill="both", expand=True, padx=22, pady=(0, 16))
        loghead = tk.Frame(logcard, bg=BG)
        loghead.pack(fill="x")
        self.log_toggle = tk.Button(loghead, text="▾  Log", command=self._toggle_log,
                                    font=(FONT, 9, "bold"), bg=BG, fg=SUB,
                                    relief="flat", bd=0, cursor="hand2",
                                    activebackground=BG, activeforeground=INK)
        self.log_toggle.pack(side="left")
        tk.Button(loghead, text="Copy log", command=self._copy_log,
                  font=(FONT, 8), bg=BG, fg=SUB, relief="flat", bd=0,
                  cursor="hand2", activebackground=BG,
                  activeforeground=INK).pack(side="right")

        self.log_frame = tk.Frame(logcard, bg=LOGBG, highlightbackground=BORDER,
                                  highlightthickness=1)
        self.log_frame.pack(fill="both", expand=True, pady=(4, 0))
        self.log = tk.Text(self.log_frame, height=10, wrap="word",
                           font=(FONT_MONO, 9), bg=LOGBG, fg=LOGINK,
                           relief="flat", bd=0, padx=10, pady=8,
                           highlightthickness=0, insertbackground=INK)
        sb = tk.Scrollbar(self.log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        self.log.tag_configure("ok", foreground=ACCENTB)
        self.log.tag_configure("warn", foreground="#F0B429")
        self.log.tag_configure("err", foreground="#FF6B6B")
        self.log.tag_configure("dim", foreground=SUB)

    def _path_row(self, parent, var, role="file", folder_only=False):
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", pady=(8, 0))
        e = self._entry(row, var)
        e.pack(side="left", fill="x", expand=True, ipady=5)
        if not folder_only:
            self._ghost_btn(row, "Browse file…",
                            lambda: self._pick_file(var, role)).pack(
                side="left", padx=(8, 0))
        self._ghost_btn(row, "Browse folder…",
                        lambda: self._pick_folder(var, role)).pack(
            side="left", padx=(8, 0))

    # ---------------------------------------------------------- callbacks --
    def _on_source_change(self, *_):
        if self.source_var.get() == "Find Batch":
            self.mode_hint.config(
                text="Find Batch mode: As-Design = the list MISSING batches (e.g. RED)  |  "
                     "As-Built = the label/scan list that HAS batches. Fills Charge + Rev.")
        else:
            self.mode_hint.config(text="")

    def _pick_file(self, var, role="file"):
        f = filedialog.askopenfilename(
            title=f">>> SELECT THE {role.upper()} FILE <<<",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if f:
            var.set(f)

    def _pick_folder(self, var, role="folder"):
        d = filedialog.askdirectory(title=f">>> SELECT THE {role.upper()} FOLDER <<<")
        if d:
            var.set(d)

    def _toggle_log(self):
        self._log_visible = not self._log_visible
        if self._log_visible:
            self.log_frame.pack(fill="both", expand=True, pady=(4, 0))
            self.log_toggle.config(text="▾  Log")
        else:
            self.log_frame.pack_forget()
            self.log_toggle.config(text="▸  Log")

    def _copy_log(self):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.log.get("1.0", "end"))
        except Exception:
            pass

    # -------------------------------------------------------- log + progress
    def write_log(self, msg):
        self.root.after(0, self._write_log_ui, str(msg))

    def _write_log_ui(self, msg):
        low = msg.lower()
        tag = None
        if "error" in low or "could not" in low or "failed" in low:
            tag = "err"
        elif "warning" in low or "!!!" in low or "skipped" in low:
            tag = "warn"
        elif msg.strip().lower().startswith(("done", "finished")) or "conservation" in low:
            tag = "ok"
        elif msg.strip().startswith(" "):
            tag = "dim"
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        ms = _milestone(msg)
        if ms:
            status, frac = ms
            if "distinct parts ->" in low:
                self._files_done = min(self._files_done + 1, self._files_total)
            base = self._files_done / max(1, self._files_total)
            span = 1.0 / max(1, self._files_total)
            overall = (base + frac * span) * 100.0
            if frac >= 1.0:
                overall = 100.0
            self._target = max(self._target, min(100.0, overall))
            self.status_lbl.config(text=status)

    def _tick(self):
        if self._pct < self._target:
            self._pct += max(0.4, (self._target - self._pct) * 0.18)
            if self._pct > self._target:
                self._pct = self._target
        self.bar["value"] = self._pct
        self.pct_lbl.config(text=f"{int(round(self._pct))}%")
        self.root.after(33, self._tick)

    # ----------------------------------------------------------------- run --
    def run(self):
        if self._running:
            return
        design = self.design_var.get().strip()
        built = self.built_var.get().strip()
        out = self.out_var.get().strip()
        self.log.delete("1.0", "end")
        self._pct = 0.0
        self._target = 0.0
        self._files_done = 0
        if not design or not built or not out:
            self._write_log_ui("Please pick the As-Design, the As-Built source, "
                               "and the Output folder.")
            self.status_lbl.config(text="Missing inputs.")
            return
        if not os.path.exists(design):
            self._write_log_ui(f"As-Design not found: {design}")
            return
        if not os.path.exists(built):
            self._write_log_ui(f"As-Built source not found: {built}")
            return

        combine = bool(self.combine_var.get())
        if combine or os.path.isfile(built):
            self._files_total = 1
        else:
            files = [f for f in glob.glob(os.path.join(built, "*.xls*"))
                     if not os.path.basename(f).startswith("~$")]
            self._files_total = max(1, len(files))

        label = "As Built"          # fixed: tabs are always As-Built / As-Design
        self._running = True
        self.run_btn.config(state="disabled", text="Running…", bg=SUB)
        self.status_lbl.config(text="Starting…")
        self._target = 3.0
        threading.Thread(target=self._worker,
                         args=(design, built, out, combine, label),
                         daemon=True).start()

    def _worker(self, design, built, out, combine, label):
        try:
            if self.source_var.get() == "Find Batch":
                if batch_finder is None:
                    self.write_log("batch_finder.py not found next to adab_gui.py.")
                    return
                self.write_log("Mode: FIND BATCH - fill Charge/Rev by lookup\n")
                self.write_log("  Need-batch list = As-Design field")
                self.write_log("  Label/scan source = As-Built field\n")
                out_path, (n, f, m, z) = batch_finder.run(
                    design, built, out, progress=self.write_log)
                self.write_log(
                    f"\nFinished. FOUND {f}, MULTIPLE {m}, NONE {z} of {n}."
                    f"\nReport: {out_path}")
                self._done_ok()
                return
            mode = "COMBINED (one report)" if combine else "one report per file"
            self.write_log(f"Source type: {label}   |   Mode: {mode}\n")
            results = core.run_compare(design, built, out, combine=combine,
                                       progress=self.write_log,
                                       built_label=label)
            ok = [r for r in results if "matched" in r]
            self.write_log(f"\nFinished. {len(ok)} report(s) written to:\n{out}")
            self._done_ok()
        except Exception as e:
            self.write_log("\nERROR: " + str(e))
            self.write_log(traceback.format_exc())
            self.root.after(0, lambda: self.status_lbl.config(text="Error — see log."))
        finally:
            self.root.after(0, self._reset_run_btn)

    def _done_ok(self):
        self.root.after(0, lambda: (setattr(self, "_target", 100.0),
                                    self.status_lbl.config(text="Finished ✓")))

    def _reset_run_btn(self):
        self._running = False
        self.run_btn.config(state="normal", text="Run  ·  As-Built vs As-Design",
                            bg=ACCENT)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()