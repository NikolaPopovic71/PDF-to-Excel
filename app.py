"""PDF -> Excel konverter za rezultate plivanja — paketna obrada."""
import glob
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import convert

try:  # prevlačenje fajlova mišem (opciono)
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
except Exception:  # noqa: BLE001
    DND_FILES = None
    BaseTk = tk.Tk

BG = "#F4F6FA"
ACCENT = "#1F4E78"
FONT = "Segoe UI"


def open_path(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:  # noqa: BLE001
        pass


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("Rezultati plivanja — PDF u Excel")
        self.geometry("900x620")
        self.minsize(720, 480)
        self.configure(bg=BG)

        self.files = {}          # item_id -> {"pdf": ..., "xlsx": ...}
        self.out_mode = tk.StringVar(value="same")
        self.out_dir = tk.StringVar()
        self.open_folder = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Dodajte PDF fajlove (dugmetom, celim folderom ili prevlačenjem u listu).")
        self.running = False

        self._style()
        self._build()

    # ---------- izgled ----------
    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame", background=BG)
        s.configure("TLabelframe", background=BG)
        s.configure("TLabelframe.Label", background=BG, font=(FONT, 10, "bold"), foreground=ACCENT)
        s.configure("TLabel", background=BG, font=(FONT, 10))
        s.configure("Title.TLabel", font=(FONT, 18, "bold"), foreground=ACCENT)
        s.configure("TCheckbutton", background=BG, font=(FONT, 10))
        s.configure("TRadiobutton", background=BG, font=(FONT, 10))
        s.configure("TButton", font=(FONT, 10), padding=6)
        s.configure("Big.TButton", font=(FONT, 12, "bold"), padding=12, foreground="white", background=ACCENT)
        s.map("Big.TButton", background=[("active", "#2E6DA4"), ("disabled", "#9AA9BC")])
        s.configure("Treeview", font=(FONT, 10), rowheight=26)
        s.configure("Treeview.Heading", font=(FONT, 10, "bold"))

    def _build(self):
        root = ttk.Frame(self, padding=20)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="PDF ➜ Excel", style="Title.TLabel").pack(anchor="w")
        ttk.Label(root, text="Pretvara izveštaje „Medalists by event“ u Excel tabele — više fajlova odjednom.").pack(anchor="w", pady=(0, 12))

        # dugmad za listu
        bar = ttk.Frame(root)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="➕ Dodaj PDF fajlove…", command=self.add_files).pack(side="left")
        ttk.Button(bar, text="📁 Dodaj ceo folder…", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(bar, text="Ukloni izabrane", command=self.remove_selected).pack(side="left")
        ttk.Button(bar, text="🗑 Očisti sve", command=self.clear_all).pack(side="right")

        # lista fajlova
        lf = ttk.Frame(root)
        lf.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(lf, columns=("file", "status"), show="headings", selectmode="extended")
        self.tree.heading("file", text="PDF fajl")
        self.tree.heading("status", text="Status")
        self.tree.column("file", width=520, anchor="w")
        self.tree.column("status", width=280, anchor="w")
        self.tree.tag_configure("ok", foreground="#1E7B34")
        self.tree.tag_configure("err", foreground="#B3261E")
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self.open_result)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        if DND_FILES:
            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind("<<Drop>>", self.on_drop)

        # gde se čuva
        out = ttk.Labelframe(root, text="Gde sačuvati Excel fajlove", padding=10)
        out.pack(fill="x", pady=10)
        ttk.Radiobutton(out, text="U isti folder gde je PDF", variable=self.out_mode, value="same").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(out, text="U folder:", variable=self.out_mode, value="custom").grid(row=1, column=0, sticky="w")
        ttk.Entry(out, textvariable=self.out_dir).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Button(out, text="Izaberi…", command=self.pick_out).grid(row=1, column=2)
        out.columnconfigure(1, weight=1)
        ttk.Checkbutton(out, text="Posle konverzije otvori rezultat (jedan fajl → Excel, više fajlova → folder)", variable=self.open_folder).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        self.btn = ttk.Button(root, text="Konvertuj sve u Excel", style="Big.TButton", command=self.run)
        self.btn.pack(fill="x")
        self.bar = ttk.Progressbar(root, mode="determinate")
        self.bar.pack(fill="x", pady=(10, 4))
        ttk.Label(root, textvariable=self.status, wraplength=840).pack(anchor="w")

    # ---------- lista ----------
    def _add(self, paths):
        existing = {v["pdf"] for v in self.files.values()}
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            if p.lower().endswith(".pdf") and os.path.isfile(p) and p not in existing:
                iid = self.tree.insert("", "end", values=(p, "Čeka"))
                self.files[iid] = {"pdf": p, "xlsx": None}
                existing.add(p)
                added += 1
        self._count(f"Dodato: {added}.")

    def _count(self, prefix=""):
        self.status.set(f"{prefix} U listi: {len(self.files)} fajlova.".strip())

    def add_files(self):
        self._add(filedialog.askopenfilenames(title="Izaberite PDF fajlove", filetypes=[("PDF fajlovi", "*.pdf")]))

    def add_folder(self):
        d = filedialog.askdirectory(title="Izaberite folder sa PDF fajlovima")
        if d:
            self._add(sorted(glob.glob(os.path.join(d, "**", "*.pdf"), recursive=True)))

    def on_drop(self, event):
        paths = []
        for p in self.tk.splitlist(event.data):
            if os.path.isdir(p):
                paths += sorted(glob.glob(os.path.join(p, "**", "*.pdf"), recursive=True))
            else:
                paths.append(p)
        self._add(paths)

    def remove_selected(self):
        if self.running:
            return
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.files.pop(iid, None)
        self._count()

    def clear_all(self):
        if self.running:
            return
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self.bar["value"] = 0
        self._count("Lista je očišćena.")

    def pick_out(self):
        d = filedialog.askdirectory(title="Folder za Excel fajlove")
        if d:
            self.out_dir.set(d)
            self.out_mode.set("custom")

    def open_result(self, _event):
        sel = self.tree.selection()
        if sel and self.files.get(sel[0], {}).get("xlsx"):
            open_path(self.files[sel[0]]["xlsx"])

    # ---------- konverzija ----------
    def run(self):
        if not self.files:
            messagebox.showwarning("Prazna lista", "Prvo dodajte PDF fajlove.")
            return
        if self.out_mode.get() == "custom":
            d = self.out_dir.get().strip()
            if not d:
                messagebox.showwarning("Folder", "Izaberite folder za Excel fajlove.")
                return
            os.makedirs(d, exist_ok=True)
        self.running = True
        self.btn.state(["disabled"])
        self.bar.configure(maximum=len(self.files), value=0)
        threading.Thread(target=self._work, args=(list(self.files.items()),), daemon=True).start()

    def _target(self, pdf):
        name = os.path.splitext(os.path.basename(pdf))[0] + ".xlsx"
        folder = os.path.dirname(pdf) if self.out_mode.get() == "same" else self.out_dir.get().strip()
        return os.path.join(folder, name)

    def _work(self, items):
        ok = err = 0
        folders = set()
        produced = []
        for i, (iid, info) in enumerate(items, 1):
            dst = self._target(info["pdf"])
            self.after(0, self._set, iid, "Konvertujem…", None)
            try:
                n_ev, n_rows = convert(info["pdf"], dst)
                info["xlsx"] = dst
                folders.add(os.path.dirname(dst))
                produced.append(dst)
                msg, tag = f"✓ {n_ev} disciplina, {n_rows} redova", "ok"
                ok += 1
            except PermissionError:
                msg, tag = "✗ Excel fajl je otvoren — zatvorite ga", "err"
                err += 1
            except Exception as e:  # noqa: BLE001
                msg, tag = f"✗ {e}", "err"
                err += 1
            self.after(0, self._set, iid, msg, tag, i, len(items))
        self.after(0, self._finish, ok, err, folders, produced)

    def _set(self, iid, msg, tag, i=None, total=None):
        if self.tree.exists(iid):
            self.tree.item(iid, values=(self.files[iid]["pdf"], msg), tags=(tag,) if tag else ())
            self.tree.see(iid)
        if i is not None:
            self.bar["value"] = i
            self.status.set(f"Obrađeno {i} od {total}…")

    def _finish(self, ok, err, folders, produced):
        self.running = False
        self.btn.state(["!disabled"])
        self.status.set(f"Gotovo! Uspešno: {ok}, greške: {err}. Dvoklik na red otvara Excel fajl.")
        if self.open_folder.get() and ok:
            if ok == 1:
                open_path(produced[0])            # jedan fajl -> otvori Excel
            else:
                for f in sorted(folders)[:5]:  # više fajlova -> otvori folder(e)
                    open_path(f)
        if err:
            messagebox.showwarning("Završeno uz greške", f"Uspešno: {ok}\nGreške: {err}\n\nDetalji su u koloni Status.")


if __name__ == "__main__":
    App().mainloop()
