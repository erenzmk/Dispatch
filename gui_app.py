#!/usr/bin/env python
from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import datetime as dt
import logging

from dispatch import process_reports, aggregate_warnings
import assign_gui

class TextHandler(logging.Handler):
    def __init__(self, widget: tk.Text) -> None:
        super().__init__()
        self.widget = widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        self.widget.after(0, lambda: self.widget.insert(tk.END, msg + "\n"))


class DispatchApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Dispatch Interface")
        self._worker: threading.Thread | None = None
        self._pause = threading.Event()
        self._stop = threading.Event()

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frm, text="Tagesordner").grid(row=0, column=0, sticky="w")
        self.day_dir_var = tk.StringVar()
        day_entry = ttk.Entry(frm, textvariable=self.day_dir_var, width=40)
        day_entry.grid(row=0, column=1, sticky="we")
        ttk.Button(frm, text="...", command=self._choose_day_dir).grid(row=0, column=2)

        ttk.Label(frm, text="Liste.xlsx").grid(row=1, column=0, sticky="w")
        self.liste_var = tk.StringVar()
        liste_entry = ttk.Entry(frm, textvariable=self.liste_var, width=40)
        liste_entry.grid(row=1, column=1, sticky="we")
        ttk.Button(frm, text="...", command=self._choose_liste).grid(row=1, column=2)

        ttk.Label(frm, text="Datum (dd.mm.yyyy)").grid(row=2, column=0, sticky="w")
        self.date_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.date_var, width=20).grid(row=2, column=1, sticky="w")

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=3, pady=5)
        ttk.Button(btns, text="Start", command=self._start).pack(side="left", padx=5)
        ttk.Button(btns, text="Pause", command=self._pause_toggle).pack(side="left", padx=5)
        ttk.Button(btns, text="Stop", command=self._stop_worker).pack(side="left", padx=5)
        ttk.Button(btns, text="Namen prüfen", command=self._check_names).pack(side="left", padx=5)
        ttk.Button(btns, text="Zuordnen", command=self._open_assign_gui).pack(side="left", padx=5)

        self.log = tk.Text(frm, height=15)
        self.log.grid(row=4, column=0, columnspan=3, sticky="nsew")
        frm.rowconfigure(4, weight=1)

        handler = TextHandler(self.log)
        logging.basicConfig(level=logging.INFO, handlers=[handler])

    def _choose_day_dir(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.day_dir_var.set(path)

    def _choose_liste(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.liste_var.set(path)

    def _parse_date(self) -> dt.date | None:
        if not self.date_var.get():
            return None
        try:
            return dt.datetime.strptime(self.date_var.get(), "%d.%m.%Y").date()
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Datum")
            return None

    def _worker_main(self, day_dir: Path, liste: Path, date: dt.date | None) -> None:
        argv = [str(day_dir), str(liste)]
        if date:
            argv += ["--date", date.strftime("%d.%m.%Y")]
        try:
            process_reports.main(argv)
        except Exception as exc:  # pragma: no cover - interactive
            logging.exception("Verarbeitung fehlgeschlagen: %s", exc)

    def _start(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("Info", "Laufende Verarbeitung")
            return
        day_dir = Path(self.day_dir_var.get())
        liste = Path(self.liste_var.get())
        if not day_dir.is_dir() or not liste.is_file():
            messagebox.showerror("Fehler", "Pfad prüfen")
            return
        date = self._parse_date()
        self._pause.clear()
        self._stop.clear()
        self._worker = threading.Thread(
            target=self._worker_main, args=(day_dir, liste, date), daemon=True
        )
        self._worker.start()

    def _pause_toggle(self) -> None:
        if not self._worker:
            return
        if self._pause.is_set():
            self._pause.clear()
        else:
            self._pause.set()

    def _stop_worker(self) -> None:
        self._stop.set()

    def _check_names(self) -> None:
        day_dir = Path(self.day_dir_var.get())
        liste = Path(self.liste_var.get())
        if not day_dir.is_dir() or not liste.is_file():
            messagebox.showerror("Fehler", "Pfad prüfen")
            return
        try:
            valid = aggregate_warnings.gather_valid_names(liste)
        except ValueError as exc:
            messagebox.showerror("Fehler", str(exc))
            return
        unknown = aggregate_warnings.aggregate_warnings(day_dir, valid)
        for name, count in unknown.items():
            logging.info("%s: %s", name, count)

    def _open_assign_gui(self) -> None:
        day_dir = Path(self.day_dir_var.get())
        liste = Path(self.liste_var.get())
        if not day_dir.is_dir() or not liste.is_file():
            messagebox.showerror("Fehler", "Pfad prüfen")
            return
        assign_gui.main([str(day_dir), "--liste", str(liste)])


if __name__ == "__main__":  # pragma: no cover - GUI entry
    app = DispatchApp()
    app.mainloop()
