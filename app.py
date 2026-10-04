"""PDF -> Excel konverter za rezultate plivanja (grafički interfejs)."""
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import convert

BG = "#F4F6FA"
ACCENT = "#1F4E78"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rezultati plivanja — PDF u Excel")
        self.geometry("560x330")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.pdf_path = tk.StringVar()
        self.open_after = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Izaberite PDF fajl sa rezultatima.")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), foreground=ACCENT)
        style.configure("TCheckbutton", background=BG, font=("Segoe UI", 10))
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=10,
                        foreground="white", background=ACCENT)
        style.map("Big.TButton", background=[("active", "#2E6DA4"), ("disabled", "#9AA9BC")])
        style.configure("TButton", font=("Segoe UI", 10), padding=6)

        f = ttk.Frame(self, padding=24)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text="PDF ➜ Excel", style="Title.TLabel").pack(anchor="w")
        ttk.Label(f, text="Pretvara izveštaj „Medalists by event“ u Excel tabelu.").pack(anchor="w", pady=(0, 16))

        row = ttk.Frame(f)
        row.pack(fill="x")
        ttk.Entry(row, textvariable=self.pdf_path, font=("Segoe UI", 10)).pack(side="left", fill="x", expand=True, ipady=4)
        ttk.Button(row, text="Izaberi PDF…", command=self.pick).pack(side="left", padx=(8, 0))

        ttk.Checkbutton(f, text="Otvori Excel posle konverzije", variable=self.open_after).pack(anchor="w", pady=12)

        self.btn = ttk.Button(f, text="Konvertuj u Excel", style="Big.TButton", command=self.run)
        self.btn.pack(fill="x")

        self.bar = ttk.Progressbar(f, mode="indeterminate")
        self.bar.pack(fill="x", pady=(14, 4))
        ttk.Label(f, textvariable=self.status, wraplength=500).pack(anchor="w")

    def pick(self):
        p = filedialog.askopenfilename(title="Izaberite PDF", filetypes=[("PDF fajlovi", "*.pdf")])
        if p:
            self.pdf_path.set(p)
            self.status.set("Spremno. Kliknite „Konvertuj u Excel“.")

    def run(self):
        src = self.pdf_path.get().strip().strip('"')
        if not src or not os.path.isfile(src):
            messagebox.showwarning("Nema fajla", "Prvo izaberite PDF fajl.")
            return
        dst = filedialog.asksaveasfilename(
            title="Sačuvaj Excel kao", defaultextension=".xlsx",
            initialdir=os.path.dirname(src),
            initialfile=os.path.splitext(os.path.basename(src))[0] + ".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not dst:
            return
        self.btn.state(["disabled"])
        self.bar.start(12)
        self.status.set("Konvertujem…")
        threading.Thread(target=self._work, args=(src, dst), daemon=True).start()

    def _work(self, src, dst):
        try:
            n_ev, n_rows = convert(src, dst)
            self.after(0, self._done, dst, n_ev, n_rows, None)
        except PermissionError:
            self.after(0, self._done, dst, 0, 0, "Excel fajl je verovatno otvoren. Zatvorite ga i pokušajte ponovo.")
        except Exception as e:  # noqa: BLE001
            self.after(0, self._done, dst, 0, 0, str(e))

    def _done(self, dst, n_ev, n_rows, err):
        self.bar.stop()
        self.btn.state(["!disabled"])
        if err:
            self.status.set("Greška.")
            messagebox.showerror("Greška", err)
            return
        self.status.set(f"Gotovo: {n_ev} disciplina, {n_rows} redova.\n{dst}")
        if self.open_after.get():
            try:
                if sys.platform.startswith("win"):
                    os.startfile(dst)  # noqa
                elif sys.platform == "darwin":
                    os.system(f'open "{dst}"')
                else:
                    os.system(f'xdg-open "{dst}"')
            except Exception:  # noqa: BLE001
                pass


if __name__ == "__main__":
    App().mainloop()
