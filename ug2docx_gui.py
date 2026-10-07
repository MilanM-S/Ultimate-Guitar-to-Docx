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


def resource_path(rel):
    """Find a bundled file, both when run as a script and inside the packaged app."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def system_is_dark():
    """True if the operating system is set to dark mode."""
    forced = os.environ.get("CSM_THEME")          # for testing: CSM_THEME=dark / light
    if forced:
        return forced == "dark"
    try:
        if sys.platform.startswith("win"):
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
        if sys.platform == "darwin":
            r = subprocess.run(["defaults", "read", "-g", "AppleInterfaceStyle"],
                               capture_output=True, text=True)
            return "Dark" in r.stdout
    except Exception:
        pass
    return False


DARK = dict(bg="#202020", fg="#f0f0f0", field="#2b2b2b", border="#454545", btn="#333333",
            btn_hot="#404040", accent="#4c9aff", muted="#a0a0a0", tab="#2a2a2a")


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
        if sys.platform.startswith("win"):
            try:   # own taskbar identity, so Windows shows our icon and not Python's feather
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ChordSheetMaker.App")
            except Exception:
                pass
        super().__init__()
        self.style = ttk.Style()
        self._native_theme = self.style.theme_use()
        self._default_bg = self.cget("bg")
        self._dark = None
        self._themed = sys.platform.startswith("win") or bool(os.environ.get("CSM_THEME"))
        self._set_icon()
        self.title("Chord Sheet Maker")
        self.geometry("620x430")
        self.minsize(600, 420)
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
        ttk.Label(self, textvariable=self.status, style="Muted.TLabel").pack(
            fill="x", padx=12, pady=(0, 8))

        self._poll_theme()

    # ---------------- icon and theme ----------------
    def _set_icon(self):
        try:
            ico = resource_path(os.path.join("assets", "AppIcon.ico"))
            png = resource_path(os.path.join("assets", "AppIcon.iconset", "icon_256x256.png"))
            if sys.platform.startswith("win") and os.path.exists(ico):
                self.iconbitmap(default=ico)
            elif os.path.exists(png):
                self._icon_img = tk.PhotoImage(file=png)
                self.iconphoto(True, self._icon_img)
        except Exception:
            pass

    def _poll_theme(self):
        """Follow the system light/dark setting, also while the app is open."""
        if self._themed:
            dark = system_is_dark()
            if dark != self._dark:
                self._dark = dark
                self.apply_theme(dark)
            self.after(2000, self._poll_theme)
        else:
            self.apply_theme(False)

    def apply_theme(self, dark):
        s = self.style
        muted = DARK["muted"] if dark else "#666666"
        if dark:
            d = DARK
            s.theme_use("clam")
            self.configure(bg=d["bg"])
            s.configure(".", background=d["bg"], foreground=d["fg"], fieldbackground=d["field"],
                        bordercolor=d["border"], lightcolor=d["bg"], darkcolor=d["bg"],
                        troughcolor=d["field"], focuscolor=d["bg"], insertcolor=d["fg"])
            s.configure("TNotebook", background=d["bg"], bordercolor=d["border"])
            s.configure("TNotebook.Tab", background=d["tab"], foreground=d["muted"], padding=(12, 4))
            s.map("TNotebook.Tab", background=[("selected", d["bg"])], foreground=[("selected", d["fg"])])
            s.configure("TButton", background=d["btn"], foreground=d["fg"], bordercolor=d["border"], padding=6)
            s.map("TButton", background=[("active", d["btn_hot"]), ("disabled", d["bg"])],
                  foreground=[("disabled", d["muted"])])
            s.configure("TEntry", fieldbackground=d["field"], foreground=d["fg"], insertcolor=d["fg"])
            s.configure("TSpinbox", fieldbackground=d["field"], foreground=d["fg"], background=d["btn"],
                        arrowcolor=d["fg"], insertcolor=d["fg"], bordercolor=d["border"])
            s.map("TSpinbox", fieldbackground=[("disabled", d["bg"])], foreground=[("disabled", d["muted"])])
            for w in ("TCheckbutton", "TRadiobutton"):
                s.configure(w, background=d["bg"], foreground=d["fg"], indicatorbackground=d["field"],
                            indicatorforeground=d["fg"], upperbordercolor=d["border"],
                            lowerbordercolor=d["border"])
                s.map(w, background=[("active", d["bg"])], foreground=[("disabled", d["muted"])],
                      indicatorcolor=[("selected", d["accent"]), ("!selected", d["field"])])
        else:
            s.theme_use(self._native_theme)
            self.configure(bg=self._default_bg)
        s.configure("Muted.TLabel", foreground=muted)
        self._dark_titlebar(dark)

    def _dark_titlebar(self, dark):
        """Dark title bar on Windows 10/11."""
        if not sys.platform.startswith("win"):
            return
        try:
            import ctypes
            self.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            val = ctypes.c_int(1 if dark else 0)
            for attr in (20, 19):                 # 20 = Windows 11 / newer 10, 19 = older 10
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, attr, ctypes.byref(val), ctypes.sizeof(val)) == 0:
                    break
        except Exception:
            pass

    # ---------------- shared widgets ----------------
    def _transpose_controls(self, parent, row, var_steps, var_acc):
        ttk.Label(parent, text="Transpose (semitones):").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Spinbox(parent, from_=-11, to=11, width=5, textvariable=var_steps,
                    wrap=True).grid(row=row, column=1, sticky="w")
        ttk.Label(parent, text="0 = leave as is", style="Muted.TLabel").grid(
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
        ttk.Label(t, textvariable=self.html_label, style="Muted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w")
        ttk.Button(t, text="Choose file…", command=self.pick_html).grid(row=1, column=2, sticky="e")

        self.steps = tk.StringVar(value="0")
        self.acc = tk.StringVar(value="auto")
        self._transpose_controls(t, 2, self.steps, self.acc)

        self.fix_capo = tk.BooleanVar(value=True)
        self.want_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(t, text="Convert capo chords to the original key",
                        variable=self.fix_capo).grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.autofit = tk.BooleanVar(value=True)
        self.size = tk.StringVar(value="12")
        fit = ttk.Frame(t)
        fit.grid(row=5, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(fit, text="Auto-fit text size (8-12 pt) to use the fewest pages",
                        variable=self.autofit, command=self._toggle_size).pack(side="left")
        ttk.Label(fit, text="   Text size:").pack(side="left")
        self.size_box = ttk.Spinbox(fit, from_=6, to=16, increment=0.5, width=5,
                                    textvariable=self.size)
        self.size_box.pack(side="left", padx=(4, 2))
        ttk.Label(fit, text="pt").pack(side="left")
        self._toggle_size()
        ttk.Checkbutton(t, text="Also save a PDF", variable=self.want_pdf).grid(
            row=6, column=0, columnspan=3, sticky="w")

        self.go = ttk.Button(t, text="Create Word document", command=self.make)
        self.go.grid(row=7, column=0, columnspan=3, pady=(16, 0), ipadx=12, ipady=4)

    def _toggle_size(self):
        """The manual size box is only usable when auto-fit is off."""
        self.size_box.config(state="disabled" if self.autofit.get() else "normal")

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
        if not self.autofit.get():
            try:
                if not 6 <= float(self.size.get()) <= 16:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Text size", "Text size must be a number between 6 and 16.")
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
        title, artist, content, meta_capo = result
        capo = core.detect_capo(content, meta_capo)
        if self.fix_capo.get():
            steps += capo
        default = f"{core.safe_filename(title)} - {core.safe_filename(artist)}.docx"
        path = filedialog.asksaveasfilename(
            title="Save Word document", initialfile=default, defaultextension=".docx",
            filetypes=[("Word document", "*.docx")])
        self.go.config(state="normal")
        if not path:
            self.status.set("Cancelled.")
            return
        try:
            sections = core.build_sections(content)
            flats = self._flats_value(self.acc.get())
            manual = 12.0 if self.autofit.get() else float(self.size.get())
            size, pages = core.plan_layout(title, artist, sections, steps, flats,
                                           autofit=self.autofit.get(), size=manual)
            core.make_docx(title, artist, sections, path, transpose=steps, flats=flats, size=size)
            saved = [os.path.basename(path)]
            if self.want_pdf.get():
                pdf = os.path.splitext(path)[0] + ".pdf"
                core.make_pdf(title, artist, sections, pdf, transpose=steps, flats=flats, size=size)
                saved.append(os.path.basename(pdf))
        except Exception as e:
            self.fail(str(e))
            return
        note = f"\nText size {size:g} pt, {pages} page{'s' if pages != 1 else ''}."
        if capo:
            note += (f"\nCapo {capo} found: " + ("chords raised to the original key." if self.fix_capo.get()
                    else "kept as on the page."))
        self.status.set("Saved: " + ", ".join(saved))
        if messagebox.askyesno("Done", "Saved " + " and ".join(saved) + "." + note + "\n\nOpen the Word file now?"):
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

        self.t_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(t, text="Also save a PDF", variable=self.t_pdf).grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Button(t, text="Transpose and save a copy", command=self.transpose).grid(
            row=4, column=0, columnspan=3, pady=(16, 0), ipadx=12, ipady=4)

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
            pdf = None
            if self.t_pdf.get():
                pdf = os.path.splitext(path)[0] + ".pdf"
                core.docx_to_pdf(path, pdf)
        except Exception as e:
            self.fail(str(e))
            return
        self.status.set(f"Transposed {n} chord lines by {steps:+d}.")
        saved = os.path.basename(path) + (f" and {os.path.basename(pdf)}" if pdf else "")
        if messagebox.askyesno("Done", f"Transposed {n} chord lines.\nSaved {saved}.\n\nOpen the Word copy now?"):
            open_file(path)


if __name__ == "__main__":
    App().mainloop()