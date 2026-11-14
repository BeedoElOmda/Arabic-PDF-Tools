"""Arabic PDF Tools - convert PDFs to editable PowerPoint / Word with correct Arabic.

Run with no arguments for the window, or from the right-click menu:
    ArabicPDFTools.pyw --pptx file1.pdf [file2.pdf ...]
    ArabicPDFTools.pyw --docx file1.pdf ...
    ArabicPDFTools.pyw --fix  file1.pdf ...
"""
import os
import sys
import threading
import traceback

APP_NAME = "Arabic PDF Tools"
AUTHOR = "Beedo El-Omda"
SIGNOFF = "Enjoy , Beedo El-Omda"
APP_VERSION = "dev"
try:                                   # build.sh writes the version next to the app
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.txt"), encoding="utf-8") as _f:
        APP_VERSION = _f.read().strip() or APP_VERSION
except OSError:
    pass
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

# pythonw has no console: send any output to a log file so nothing breaks.
LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ArabicPDFTools")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "log.txt")
if sys.stdout is None or sys.stderr is None or os.path.basename(sys.executable).lower() == "pythonw.exe":
    _log = open(LOG_FILE, "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stderr = _log

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

MODES = {
    "pptx": ("PowerPoint", "Convert to PowerPoint (editable)"),
    "docx": ("Word", "Convert to Word (editable)"),
    "fix": ("Repaired PDF", "Repair Arabic in PDF"),
}


def convert_one(mode, path, pptx_images=False):
    """Convert one PDF. Returns (output_path, count)."""
    if mode == "pptx":
        import pdf2pptx
        return pdf2pptx.convert(path, pptx_images)
    if mode == "docx":
        import pdf2docx_arabic
        return pdf2docx_arabic.convert(path)
    if mode == "fix":
        import fix_arabic_pdf
        out, _fonts, codes = fix_arabic_pdf.fix_pdf(path)
        return out, codes
    raise ValueError(mode)


def open_path(p):
    try:
        os.startfile(p)          # Windows
    except AttributeError:
        import subprocess
        subprocess.Popen(["xdg-open", p])
    except OSError:
        pass


def show_in_folder(p):
    try:
        import subprocess
        subprocess.Popen(["explorer", "/select,", os.path.normpath(p)])
    except Exception:
        open_path(os.path.dirname(p))


def set_icon(win):
    ico = os.path.join(HERE, "icon.ico")
    if os.path.exists(ico):
        try:
            win.iconbitmap(ico)
        except tk.TclError:
            pass


# ------------------------------------------------------------ main window ---
class App:
    def __init__(self, root, files=(), mode=None):
        self.root = root
        self.busy = False
        self.outputs = []
        root.title(f"{APP_NAME} {APP_VERSION}  -  by {AUTHOR}")
        set_icon(root)
        root.geometry("700x600")
        root.minsize(600, 560)

        style = ttk.Style(root)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Sub.TLabel", foreground="#555555")
        style.configure("Big.TButton", padding=(10, 8))

        top = ttk.Frame(root, padding=(16, 14, 16, 6))
        top.pack(fill="x")
        ttk.Label(top, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(top, text="Turn PDFs into editable PowerPoint or Word files with correct Arabic text.",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 0))
        ttk.Label(top, text=f"by {AUTHOR}", style="Sub.TLabel").pack(anchor="w")

        # file list
        mid = ttk.Frame(root, padding=(16, 8))
        mid.pack(fill="both", expand=True)
        bar = ttk.Frame(mid)
        bar.pack(fill="x")
        ttk.Label(bar, text="PDF files").pack(side="left")
        ttk.Button(bar, text="Clear", command=self.clear).pack(side="right")
        ttk.Button(bar, text="Remove", command=self.remove).pack(side="right", padx=4)
        ttk.Button(bar, text="Add PDFs...", command=self.add).pack(side="right")
        lf = ttk.Frame(mid)
        lf.pack(fill="both", expand=True, pady=(6, 0))
        self.listbox = tk.Listbox(lf, selectmode="extended", activestyle="none", height=6,
                                  font=("Segoe UI", 10), borderwidth=1, relief="solid")
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # actions
        act = ttk.Frame(root, padding=(16, 6))
        act.pack(fill="x")
        self.btns = []
        for m in ("pptx", "docx", "fix"):
            b = ttk.Button(act, text=MODES[m][1], style="Big.TButton",
                           command=lambda m=m: self.start(m))
            b.pack(side="left", padx=(0, 8))
            self.btns.append(b)
        self.pics = tk.BooleanVar(value=False)
        ttk.Checkbutton(root, text="PowerPoint: pictures only (exact look, text not editable)",
                        variable=self.pics).pack(anchor="w", padx=16)

        # progress + log
        bot = ttk.Frame(root, padding=(16, 8, 16, 14))
        bot.pack(fill="both")
        self.prog = ttk.Progressbar(bot, mode="determinate")
        self.prog.pack(fill="x")
        self.status = ttk.Label(bot, text="Add PDF files, then choose a conversion.")
        self.status.pack(anchor="w", pady=(6, 4))
        self.log = tk.Text(bot, height=6, font=("Consolas", 9), state="disabled",
                           borderwidth=1, relief="solid", wrap="word")
        self.log.pack(fill="both", expand=True)
        row = ttk.Frame(bot)
        row.pack(fill="x", pady=(8, 0))
        self.open_btn = ttk.Button(row, text="Open result", command=self.open_last, state="disabled")
        self.open_btn.pack(side="left")
        self.folder_btn = ttk.Button(row, text="Show in folder", command=self.folder_last, state="disabled")
        self.folder_btn.pack(side="left", padx=6)
        ttk.Label(row, text="Output is saved next to each PDF.", style="Sub.TLabel").pack(side="right")

        for f in files:
            self._add_file(f)
        if mode and files:
            root.after(300, lambda: self.start(mode))

    # --- file list
    def _add_file(self, f):
        f = os.path.abspath(f)
        if f.lower().endswith(".pdf") and f not in self.listbox.get(0, "end"):
            self.listbox.insert("end", f)

    def add(self):
        for f in filedialog.askopenfilenames(title="Choose PDF files",
                                             filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]):
            self._add_file(f)

    def remove(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)

    def clear(self):
        self.listbox.delete(0, "end")

    # --- log helpers (thread-safe through after())
    def write(self, text):
        def _w():
            self.log.configure(state="normal")
            self.log.insert("end", text + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.root.after(0, _w)

    def set_status(self, text, value=None):
        def _s():
            self.status.configure(text=text)
            if value is not None:
                self.prog["value"] = value
        self.root.after(0, _s)

    # --- conversion
    def start(self, mode):
        if self.busy:
            return
        files = list(self.listbox.get(0, "end"))
        if not files:
            messagebox.showinfo(APP_NAME, "Add one or more PDF files first.")
            return
        self.busy = True
        for b in self.btns:
            b.configure(state="disabled")
        self.prog.configure(maximum=len(files), value=0)
        threading.Thread(target=self._run, args=(mode, files, self.pics.get()), daemon=True).start()

    def _run(self, mode, files, pics):
        label = MODES[mode][0]
        ok, bad = [], []
        for i, f in enumerate(files):
            name = os.path.basename(f)
            self.set_status(f"Converting {i + 1} of {len(files)} to {label}: {name}", i)
            try:
                out, n = convert_one(mode, f, pics)
                ok.append(out)
                self.write(f"OK   {name}  ->  {os.path.basename(out)}")
            except Exception as e:
                bad.append(name)
                self.write(f"FAIL {name}: {e}")
                print(traceback.format_exc())
        self.outputs = ok
        msg = f"Done: {len(ok)} converted" + (f", {len(bad)} failed (see log)" if bad else "") + "."
        self.set_status(msg, len(files))
        self.root.after(0, lambda: self._finish(ok, bad))

    def _finish(self, ok, bad):
        self.busy = False
        for b in self.btns:
            b.configure(state="normal")
        state = "normal" if ok else "disabled"
        self.open_btn.configure(state=state)
        self.folder_btn.configure(state=state)
        if bad:
            messagebox.showwarning(APP_NAME, f"{len(bad)} file(s) could not be converted:\n" + "\n".join(bad)
                                   + f"\n\nDetails: {LOG_FILE}")
        if ok:
            messagebox.showinfo(APP_NAME, f"{len(ok)} file(s) converted - saved next to each PDF.\n\n{SIGNOFF}")

    def open_last(self):
        if self.outputs:
            open_path(self.outputs[-1])

    def folder_last(self):
        if self.outputs:
            show_in_folder(self.outputs[-1])


# --------------------------------------------- right-click (quick) window ---
class Quick:
    """Small progress window for right-click conversions."""

    def __init__(self, root, mode, files):
        self.root = root
        root.title(APP_NAME)
        set_icon(root)
        root.resizable(False, False)
        root.attributes("-topmost", True)
        f = ttk.Frame(root, padding=18)
        f.pack(fill="both", expand=True)
        self.msg = ttk.Label(f, text="Starting...", width=60, justify="left")
        self.msg.pack(anchor="w")
        self.prog = ttk.Progressbar(f, mode="determinate", maximum=len(files), length=420)
        self.prog.pack(fill="x", pady=10)
        self.row = ttk.Frame(f)
        self.row.pack(fill="x")
        self.files, self.mode, self.out = files, mode, []
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        label = MODES[self.mode][0]
        bad = []
        for i, path in enumerate(self.files):
            name = os.path.basename(path)
            self.root.after(0, lambda i=i, name=name: (
                self.msg.configure(text=f"Converting to {label}: {name}"), self.prog.configure(value=i)))
            try:
                out, _n = convert_one(self.mode, path)
                self.out.append(out)
            except Exception as e:
                bad.append(f"{name}: {e}")
                print(traceback.format_exc())
        self.root.after(0, lambda: self._done(bad))

    def _done(self, bad):
        self.prog.configure(value=len(self.files))
        if self.out:
            names = ", ".join(os.path.basename(o) for o in self.out[:3]) + (" ..." if len(self.out) > 3 else "")
            self.msg.configure(text=f"Created: {names}\n\n{SIGNOFF}")
            ttk.Button(self.row, text="Open", command=lambda: (open_path(self.out[-1]), self.root.destroy())
                       ).pack(side="left")
            ttk.Button(self.row, text="Show in folder",
                       command=lambda: (show_in_folder(self.out[-1]), self.root.destroy())).pack(side="left", padx=6)
        else:
            self.msg.configure(text="Nothing was converted.")
        ttk.Button(self.row, text="Close", command=self.root.destroy).pack(side="right")
        if bad:
            messagebox.showwarning(APP_NAME, "Could not convert:\n" + "\n".join(bad) + f"\n\nDetails: {LOG_FILE}")


def main():
    args = sys.argv[1:]
    mode = None
    for m in MODES:
        if f"--{m}" in args:
            mode = m
    files = [a for a in args if a.lower().endswith(".pdf") and os.path.exists(a)]
    root = tk.Tk()
    if mode and files and "--window" not in args:
        Quick(root, mode, files)
    else:
        App(root, files, None)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print(traceback.format_exc())
        try:
            messagebox.showerror(APP_NAME, traceback.format_exc(limit=3))
        except Exception:
            pass
