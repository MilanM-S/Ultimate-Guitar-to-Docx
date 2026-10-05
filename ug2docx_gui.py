#!/usr/bin/env python3
"""
Chord Sheet Maker - a small window around ug2docx.py.
Keep this file in the same folder as ug2docx.py.

Run directly:   python3 ug2docx_gui.py
Or build a Mac app with build_mac_app.sh (see instructions).
"""

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ug2docx as core


def open_file(path):
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Chord Sheet Maker")
        self.geometry("560x330")
        self.minsize(520, 320)
        self.html_path = None
        self.doc_path = None

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)
        self.make_tab = ttk.Frame(nb, padding=12)
        self.trans_tab = ttk.Frame(nb, padding=12)
        nb.add(self.make_tab, text="New chord sheet")
        nb.add(self.trans_tab, text="Transpose a document")
        self._build_make_tab()
        self._build_transpose_tab()

        self.status = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status, foreground="#555").pack(
            fill="x", padx=12, pady=(0, 8))

    # ---------------- shared widgets ----------------
    def _transpose_controls(self, parent, row, var_steps, var_acc):
        ttk.Label(parent, text="Transpose (semitones):").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Spinbox(parent, from_=-11, to=11, width=5, textvariable=var_steps,
                    wrap=True).grid(row=row, column=1, sticky="w")
        ttk.Label(parent, text="0 = leave as is", foreground="#777").grid(
            row=row, column=2, sticky="w", padx=8)
        ttk.Label(parent, text="Chord names:").grid(row=row + 1, column=0, sticky="w")
        f = ttk.Frame(parent)
        f.grid(row=row + 1, column=1, columnspan=2, sticky="w")
        for text, val in (("Automatic", "auto"), ("Sharps (A#)", "sharps"), ("Flats (Bb)", "flats")):
            ttk.Radiobutton(f, text=text, value=val, variable=var_acc).pack(side="left", padx=(0, 10))

    @staticmethod
    def _flats_value(acc):
        return {"auto": None, "sharps": False, "flats": True}[acc]

    # ---------------- "make" tab ----------------
    def _build_make_tab(self):
        t = self.make_tab
        t.columnconfigure(1, weight=1)
        ttk.Label(t, text="Ultimate Guitar URL:").grid(row=0, column=0, sticky="w", pady=6)
        self.url = tk.StringVar()
        entry = ttk.Entry(t, textvariable=self.url)
        entry.grid(row=0, column=1, columnspan=2, sticky="ew")
        entry.focus()

        self.html_label = tk.StringVar(value="Download blocked? Use a saved page instead:")
        ttk.Label(t, textvariable=self.html_label, foreground="#777").grid(
            row=1, column=0, columnspan=2, sticky="w")
        ttk.Button(t, text="Choose file…", command=self.pick_html).grid(row=1, column=2, sticky="e")

        self.steps = tk.StringVar(value="0")
        self.acc = tk.StringVar(value="auto")
        self._transpose_controls(t, 2, self.steps, self.acc)

        self.go = ttk.Button(t, text="Create Word document", command=self.make)
        self.go.grid(row=5, column=0, columnspan=3, pady=(18, 0), ipadx=12, ipady=4)

    def pick_html(self):
        p = filedialog.askopenfilename(title="Saved Ultimate Guitar page",
                                       filetypes=[("Web page", "*.html *.htm"), ("All files", "*.*")])
        if p:
            self.html_path = p
            self.html_label.set("Using saved page: " + os.path.basename(p))

    def make(self):
        url = self.url.get().strip()
        if not url and not self.html_path:
            messagebox.showwarning("Missing URL", "Paste an Ultimate Guitar chords URL first.")
            return
        try:
            steps = int(self.steps.get())
        except ValueError:
            messagebox.showwarning("Transpose", "Transpose must be a whole number, e.g. -2 or 3.")
            return
        use_url = bool(url)
        self.go.config(state="disabled")
        self.status.set("Downloading chord sheet…")

        def work():
            try:
                if use_url:
                    page = core.fetch_html(url)
                else:
                    with open(self.html_path, encoding="utf-8", errors="replace") as f:
                        page = f.read()
                result = core.parse_ug_html(page)
                self.after(0, lambda: self.finish_make(result, steps))
            except Exception as e:
                self.after(0, lambda: self.fail(str(e)))

        threading.Thread(target=work, daemon=True).start()

    def finish_make(self, result, steps):
        title, artist, content = result
        default = f"{core.safe_filename(title)} - {core.safe_filename(artist)}.docx"
        path = filedialog.asksaveasfilename(
            title="Save Word document", initialfile=default, defaultextension=".docx",
            filetypes=[("Word document", "*.docx")])
        self.go.config(state="normal")
        if not path:
            self.status.set("Cancelled.")
            return
        try:
            core.make_docx(title, artist, core.build_sections(content), path,
                           transpose=steps, flats=self._flats_value(self.acc.get()))
        except Exception as e:
            self.fail(str(e))
            return
        self.status.set(f"Saved: {os.path.basename(path)}")
        if messagebox.askyesno("Done", f"Saved {os.path.basename(path)}.\n\nOpen it now?"):
            open_file(path)

    def fail(self, msg):
        self.go.config(state="normal")
        self.status.set("Something went wrong.")
        messagebox.showerror("Error", msg)

    # ---------------- "transpose" tab ----------------
    def _build_transpose_tab(self):
        t = self.trans_tab
        t.columnconfigure(1, weight=1)
        ttk.Label(t, text="Word document:").grid(row=0, column=0, sticky="w", pady=6)
        self.doc_label = tk.StringVar(value="(none chosen)")
        ttk.Label(t, textvariable=self.doc_label).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(t, text="Choose file…", command=self.pick_doc).grid(row=0, column=2, sticky="e")

        self.t_steps = tk.StringVar(value="2")
        self.t_acc = tk.StringVar(value="auto")
        self._transpose_controls(t, 1, self.t_steps, self.t_acc)

        ttk.Button(t, text="Transpose and save a copy", command=self.transpose).grid(
            row=4, column=0, columnspan=3, pady=(18, 0), ipadx=12, ipady=4)

    def pick_doc(self):
        p = filedialog.askopenfilename(title="Chord sheet (.docx)",
                                       filetypes=[("Word document", "*.docx")])
        if p:
            self.doc_path = p
            self.doc_label.set(os.path.basename(p))

    def transpose(self):
        if not self.doc_path:
            messagebox.showwarning("No file", "Choose a Word document first.")
            return
        try:
            steps = int(self.t_steps.get())
        except ValueError:
            messagebox.showwarning("Transpose", "Transpose must be a whole number, e.g. -2 or 3.")
            return
        base = os.path.splitext(os.path.basename(self.doc_path))[0]
        path = filedialog.asksaveasfilename(
            title="Save transposed copy", initialfile=f"{base} ({steps:+d}).docx",
            initialdir=os.path.dirname(self.doc_path), defaultextension=".docx",
            filetypes=[("Word document", "*.docx")])
        if not path:
            return
        try:
            n = core.transpose_docx(self.doc_path, steps, path,
                                    flats=self._flats_value(self.t_acc.get()))
        except Exception as e:
            self.fail(str(e))
            return
        self.status.set(f"Transposed {n} chord lines by {steps:+d}.")
        if messagebox.askyesno("Done", f"Transposed {n} chord lines.\n\nOpen the copy now?"):
            open_file(path)


if __name__ == "__main__":
    App().mainloop()
