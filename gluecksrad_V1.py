#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import random
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd


# ============================================================
# Excel handling
# ============================================================
def load_namelist(xls_path: Path) -> pd.DataFrame:
    df = pd.read_excel(xls_path)
    if df.shape[1] < 1:
        raise ValueError("Spalte A (Namen) fehlt.")
    if "A" in df.columns and "B" in df.columns:
        df = df.rename(columns={"A": "Name", "B": "Counter"})
    else:
        cols = list(df.columns)
        df = df.rename(columns={cols[0]: "Name"})
        if len(cols) >= 2:
            df = df.rename(columns={cols[1]: "Counter"})
    if "Counter" not in df.columns:
        df["Counter"] = 0
    df["Name"] = df["Name"].astype(str).str.strip()
    df = df[~df["Name"].isna() & (df["Name"] != "")]
    df["Counter"] = pd.to_numeric(df["Counter"], errors="coerce").fillna(0).astype(int)
    return df.reset_index(drop=True)


def save_namelist(df: pd.DataFrame, xls_path: Path):
    with pd.ExcelWriter(xls_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)


# ============================================================
# Fairness logic
# ============================================================
def eligible_indices(df: pd.DataFrame):
    m = df["Counter"].min()
    return df.index[df["Counter"] == m].tolist()


def normalize_if_all_equal(df: pd.DataFrame):
    if df["Counter"].nunique() == 1:
        df["Counter"] = df["Counter"] - df["Counter"].min()


# ============================================================
# GUI Application (soundless)
# ============================================================
class GluecksradApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Glücksrad – Faire Zufallsauswahl")
        self.geometry("700x600")

        self.df: pd.DataFrame | None = None
        self.xls_path: Path | None = None

        # Configurable animation parameters
        self.SPIN_FAST_MS = 18            # start delay per step (smaller = faster)
        self.SPIN_SLOW_MS = 240           # maximum (slowest) delay per step
        self.SPIN_GROW = 1.12             # multiplicative delay growth (>1.0)
        self.SPIN_ROUNDS = 3              # fast full passes over eligible list
        self.SPIN_ROLLOUT_FACTOR = 1.5    # rollout length ≈ factor × len(eligible)
        self.BLINK_TIMES = 3              # winner blink pairs
        self.BLINK_MS = 180               # ms per blink toggle

        self._build_ui()
        self._make_styles()
        self._build_menubar()

        # Animation / round state
        self.anim_running = False
        self.current_highlight_row = None
        self.to_draw_total = 0
        self.drawn_count = 0

        # NEW: Track picks within the current multi-pick round
        self.round_selected_idx: set[int] = set()  # indices picked this "batch"
        self.round_excluded_idx: set[int] = set()  # excluded from eligibility during this batch

    # ---------------- UI ----------------
    def _build_ui(self):
        topbar = ttk.Frame(self)
        topbar.pack(side=tk.TOP, fill=tk.X, padx=8, pady=8)

        ttk.Button(topbar, text="Excel laden …", command=self.on_load_excel).pack(side=tk.LEFT)
        self.path_label = ttk.Label(topbar, text="Keine Datei geladen")
        self.path_label.pack(side=tk.LEFT, padx=10)

        middle = ttk.Frame(self)
        middle.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        columns = ("Name", "Counter")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings", height=18)
        self.tree.heading("Name", text="Name")
        self.tree.heading("Counter", text="Counter")
        self.tree.column("Name", width=460, anchor=tk.W)
        self.tree.column("Counter", width=120, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(middle, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        control = ttk.Frame(self)
        control.pack(fill=tk.X, padx=8, pady=(0, 8))

        ttk.Label(control, text="Anzahl ziehen:").pack(side=tk.LEFT)
        self.entry_n = ttk.Entry(control, width=6)
        self.entry_n.insert(0, "1")
        self.entry_n.pack(side=tk.LEFT, padx=(4, 10))

        self.btn_draw = ttk.Button(control, text="Ziehung starten", command=self.on_draw_clicked, state=tk.DISABLED)
        self.btn_draw.pack(side=tk.LEFT)

        self.btn_clear = ttk.Button(control, text="Markierungen zurücksetzen", command=self.clear_winner_highlights, state=tk.DISABLED)
        self.btn_clear.pack(side=tk.LEFT, padx=(10, 0))

        self.status = ttk.Label(self, text="Bereit.")
        self.status.pack(fill=tk.X, padx=8, pady=(0, 8))

    def _make_styles(self):
        self.tree.tag_configure("scan", background="#FFE082")   # yellow
        self.tree.tag_configure("winner", background="#C8E6C9") # green
        self.tree.tag_configure("normal", background="")

    def _build_menubar(self):
        menubar = tk.Menu(self)

        menu_settings = tk.Menu(menubar, tearoff=False)
        menu_settings.add_command(label="Spin-Parameter …", command=self.open_config_dialog)
        menubar.add_cascade(label="Einstellungen", menu=menu_settings)

        menu_help = tk.Menu(menubar, tearoff=False)
        menu_help.add_command(label="Über …", command=lambda: messagebox.showinfo(
            "Über",
            "Glücksrad – faire Zufallsauswahl\nSpin-Dynamik, Blinken, Excel-Speicherung.\n(ohne Soundeffekte)"
        ))
        menubar.add_cascade(label="Hilfe", menu=menu_help)

        self.config(menu=menubar)

    # ---------------- Config Dialog ----------------
    def open_config_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Spin-Parameter")
        dlg.resizable(False, False)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        row = 0
        def field(label, val):
            nonlocal row
            ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w")
            e = ttk.Entry(frm, width=10)
            e.insert(0, str(val))
            e.grid(row=row, column=1, sticky="e")
            row += 1
            return e

        e_fast   = field("Startgeschwindigkeit (ms pro Schritt):", self.SPIN_FAST_MS)
        e_slow   = field("Endgeschwindigkeit (ms pro Schritt):", self.SPIN_SLOW_MS)
        e_grow   = field("Verzögerungsfaktor (>1.0):", self.SPIN_GROW)
        e_rounds = field("Spin-Runden:", self.SPIN_ROUNDS)
        e_roll   = field("Ausroll-Faktor:", self.SPIN_ROLLOUT_FACTOR)
        e_blink  = field("Blinkanzahl:", self.BLINK_TIMES)
        e_bms    = field("Blinktempo (ms):", self.BLINK_MS)

        row += 1
        btns = ttk.Frame(frm); btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(10,0))
        ttk.Button(btns, text="Abbrechen", command=dlg.destroy).pack(side=tk.RIGHT, padx=(10,0))

        def apply():
            try:
                self.SPIN_FAST_MS = max(5, int(float(e_fast.get())))
                self.SPIN_SLOW_MS = max(self.SPIN_FAST_MS, int(float(e_slow.get())))
                self.SPIN_GROW = max(1.01, float(e_grow.get()))
                self.SPIN_ROUNDS = max(0, int(e_rounds.get()))
                self.SPIN_ROLLOUT_FACTOR = max(0.0, float(e_roll.get()))
                self.BLINK_TIMES = max(0, int(e_blink.get()))
                self.BLINK_MS = max(20, int(e_bms.get()))
            except Exception as ex:
                messagebox.showerror("Fehler", str(ex))
                return
            dlg.destroy()

        ttk.Button(btns, text="Übernehmen", command=apply).pack(side=tk.RIGHT)

        for c in (0,1):
            frm.grid_columnconfigure(c, weight=1)

    # ---------------- Excel ----------------
    def on_load_excel(self):
        file_path = filedialog.askopenfilename(
            title="Excel mit Namensliste auswählen",
            filetypes=[("Excel-Dateien","*.xlsx *.xls")]
        )
        if not file_path:
            return
        try:
            self.df = load_namelist(Path(file_path))
        except Exception as e:
            messagebox.showerror("Fehler beim Laden", str(e))
            return
        self.xls_path = Path(file_path)
        self.path_label.config(text=str(self.xls_path))
        self.populate_tree()
        self.btn_draw.config(state=tk.NORMAL)
        self.btn_clear.config(state=tk.NORMAL)
        self.status.config(text=f"Geladen: {len(self.df)} Einträge.")

    # ---------------- Treeview ----------------
    def populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        if self.df is None:
            return
        for i,row in self.df.iterrows():
            self.tree.insert("", "end", iid=f"row-{i}", values=(row["Name"], int(row["Counter"])), tags=("normal",))
        self.current_highlight_row = None

    def refresh_counters(self):
        if self.df is None:
            return
        for i,row in self.df.iterrows():
            iid=f"row-{i}"
            if iid in self.tree.get_children(""):
                self.tree.item(iid, values=(row["Name"], int(row["Counter"])))

    # ---------------- Drawing ----------------
    def on_draw_clicked(self):
        if self.df is None or self.anim_running:
            return
        try:
            n = int(self.entry_n.get())
        except ValueError:
            messagebox.showwarning("Eingabe","Bitte Zahl eingeben.")
            return
        n = max(1, min(n, len(self.df)))
        self.to_draw_total = n
        self.drawn_count = 0
        # NEW: reset per-batch tracking
        self.round_selected_idx.clear()
        self.round_excluded_idx.clear()
        self.status.config(text=f"Ziehe {n} Person(en)…")
        self.btn_draw.config(state=tk.DISABLED)
        self.anim_running = True
        self.draw_next_one()

    def draw_next_one(self):
        if self.df is None:
            return

        # base eligible set = all with minimal counter
        elig = eligible_indices(self.df)
        if not elig:
            normalize_if_all_equal(self.df)
            elig = eligible_indices(self.df)

        # NEW: exclude indices already selected this batch, *if* others remain
        elig_filtered = [i for i in elig if i not in self.round_excluded_idx]
        if len(elig_filtered) == 0 and len(elig) > 0:
            # everyone at minimal counter is already in the excluded set:
            # allow repeats now (we've effectively covered everyone in this batch)
            elig_filtered = list(elig)

        # choose winner from filtered list
        winner_idx = random.choice(elig_filtered)

        # Build animation path: fast rounds + rollout ending exactly at winner
        path = []
        for _ in range(self.SPIN_ROUNDS):
            path.extend(elig_filtered)  # spin over visible candidates this step

        rollout = max(1, int(len(elig_filtered) * self.SPIN_ROLLOUT_FACTOR))
        start = random.randrange(len(elig_filtered))
        cur = start
        steps = 0
        while True:
            idx = elig_filtered[cur]
            path.append(idx)
            steps += 1
            if steps >= rollout and idx == winner_idx:
                break
            cur = (cur + 1) % len(elig_filtered)

        self.animate_scan_path(path, winner_idx, self.SPIN_FAST_MS)

    def animate_scan_path(self, path, winner_idx, delay):
        if not path:
            self.finish_one_draw(winner_idx)
            return
        idx = path.pop(0)
        self.highlight_scan_row(idx)
        next_delay = min(int(delay * self.SPIN_GROW), self.SPIN_SLOW_MS)
        self.after(max(5, next_delay), lambda: self.animate_scan_path(path, winner_idx, next_delay))

    def highlight_scan_row(self, idx):
        self.clear_scan_highlight()
        iid = f"row-{idx}"
        try:
            self.tree.item(iid, tags=("scan",))
            self.tree.see(iid)
            self.current_highlight_row = iid
        except tk.TclError:
            pass

    def clear_scan_highlight(self):
        if self.current_highlight_row:
            try:
                tags = self.tree.item(self.current_highlight_row, "tags")
                if "winner" in tags:
                    self.tree.item(self.current_highlight_row, tags=("winner",))
                else:
                    self.tree.item(self.current_highlight_row, tags=("normal",))
            except tk.TclError:
                pass
        self.current_highlight_row = None

    # ---------------- Blink & Finish ----------------
    def blink_row(self, iid, times_left):
        if times_left <= 0:
            try:
                self.tree.item(iid, tags=("winner",))
            except tk.TclError:
                pass
            return
        try:
            cur = self.tree.item(iid, "tags")
            new_tag = "normal" if "winner" in cur else "winner"
            self.tree.item(iid, tags=(new_tag,))
        except tk.TclError:
            pass
        self.after(self.BLINK_MS, lambda: self.blink_row(iid, times_left - 1))

    def finish_one_draw(self, idx):
        if self.df is None:
            return

        # mark visually
        iid = f"row-{idx}"
        try:
            self.tree.item(iid, tags=("winner",))
            self.tree.see(iid)
        except tk.TclError:
            pass

        # NEW: remember this winner for the rest of the batch
        self.round_selected_idx.add(idx)
        self.round_excluded_idx.add(idx)

        # blink then apply counter & continue
        self.blink_row(iid, self.BLINK_TIMES * 2)
        t = (self.BLINK_TIMES * 2) * self.BLINK_MS + 80
        self.after(t, lambda: self._apply_winner_and_continue(idx))

    def _apply_winner_and_continue(self, idx):
        # update counters (logic may trigger normalization)
        self.df.at[idx, "Counter"] += 1
        normalize_if_all_equal(self.df)
        self.refresh_counters()

        self.drawn_count += 1
        name = self.df.at[idx, "Name"]
        self.status.config(text=f"Gezogen: {self.drawn_count}/{self.to_draw_total} – Gewinner: {name}")

        if self.drawn_count < self.to_draw_total:
            self.after(250, self.draw_next_one)
        else:
            self.after(150, self.finish_round)

    def finish_round(self):
        # save Excel and keep highlights until user resets manually
        try:
            if self.xls_path and self.df is not None:
                save_namelist(self.df, self.xls_path)
        except Exception as e:
            messagebox.showerror("Speicherfehler", str(e))
        self.status.config(text="Runde beendet.")
        self.anim_running = False
        self.btn_draw.config(state=tk.NORMAL)
        # batch done → clear only the internal tracking (keep visual highlights)
        self.round_selected_idx.clear()
        self.round_excluded_idx.clear()

    # ---------------- Visual marks ----------------
    def clear_winner_highlights(self):
        """Entfernt grüne Gewinner-Markierungen (nur Optik)."""
        for iid in self.tree.get_children(""):
            try:
                self.tree.item(iid, tags=("normal",))
            except tk.TclError:
                pass
        self.clear_scan_highlight()
        self.status.config(text="Markierungen zurückgesetzt.")


# ============================================================
# Main
# ============================================================
if __name__ == "__main__":
    random.seed()
    app = GluecksradApp()
    app.mainloop()
