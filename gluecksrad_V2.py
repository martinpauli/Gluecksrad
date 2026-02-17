#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Glücksrad – Faire Zufallsauswahl V2
Bug-fixes and improvements over V1.
"""

import random
import logging
from pathlib import Path
from typing import Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")
log = logging.getLogger(__name__)


# ============================================================
# Excel handling
# ============================================================
def load_namelist(xls_path: Path) -> pd.DataFrame:
    """Load names (col A) and optional counters (col B) from an Excel file.

    Bug-fix V2: after converting Name to str, rows that became the literal
    string 'nan' (from NaN cells) are now also filtered out.
    Bug-fix V2: empty DataFrames are caught early with a clear error message.
    """
    df = pd.read_excel(xls_path)
    if df.shape[1] < 1:
        raise ValueError("Spalte A (Namen) fehlt.")

    # Rename columns robustly
    if "A" in df.columns and "B" in df.columns:
        df = df.rename(columns={"A": "Name", "B": "Counter"})
    else:
        cols = list(df.columns)
        rename_map = {cols[0]: "Name"}
        if len(cols) >= 2:
            rename_map[cols[1]] = "Counter"
        df = df.rename(columns=rename_map)

    if "Counter" not in df.columns:
        df["Counter"] = 0

    df["Name"] = df["Name"].astype(str).str.strip()
    # BUG-FIX: astype(str) turns NaN → "nan"; filter that out too
    df = df[~df["Name"].isin(["", "nan", "None"])]
    df["Counter"] = pd.to_numeric(df["Counter"], errors="coerce").fillna(0).astype(int)
    df = df.reset_index(drop=True)

    if df.empty:
        raise ValueError("Die Datei enthält keine gültigen Namen.")
    return df


def save_namelist(df: pd.DataFrame, xls_path: Path) -> None:
    """Save the name/counter DataFrame back to Excel."""
    # BUG-FIX V2: write to a temp file first, then replace – prevents data
    # loss if the write is interrupted or the file is locked.
    tmp_path = xls_path.with_suffix(".tmp.xlsx")
    try:
        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        tmp_path.replace(xls_path)
    except Exception:
        # Clean up temp file on failure
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        raise


# ============================================================
# Fairness logic
# ============================================================
def eligible_indices(df: pd.DataFrame) -> list[int]:
    """Return indices of all rows whose Counter equals the global minimum."""
    m = df["Counter"].min()
    return df.index[df["Counter"] == m].tolist()


def normalize_counters(df: pd.DataFrame) -> None:
    """Subtract the minimum counter from all rows so the lowest counter is 0.

    Improvement V2: the V1 version only normalised when ALL counters were
    equal, which meant counters grew unboundedly. This version always
    normalises relative to the minimum, keeping numbers small and
    preventing potential integer-overflow-style drift over many rounds.
    """
    m = df["Counter"].min()
    if m > 0:
        df["Counter"] = df["Counter"] - m


# ============================================================
# GUI Application
# ============================================================
class GluecksradApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Glücksrad – Faire Zufallsauswahl")
        self.geometry("740x620")
        self.minsize(500, 400)

        self.df: Optional[pd.DataFrame] = None
        self.xls_path: Optional[Path] = None

        # ----- Configurable animation parameters -----
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

        # ----- Animation / round state -----
        self.anim_running: bool = False
        self.current_highlight_row: Optional[str] = None
        self.to_draw_total: int = 0
        self.drawn_count: int = 0
        # BUG-FIX V2: store pending `after` ids so we can cancel on reload/close
        self._pending_after_ids: list[str] = []

        # Track picks within the current multi-pick round
        self.round_selected_idx: set[int] = set()
        self.round_excluded_idx: set[int] = set()

        # Properly handle window close
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ================================================================
    # Safe `after` wrapper – keeps track of pending callbacks
    # ================================================================
    def _safe_after(self, ms: int, func) -> str:
        """Schedule *func* after *ms* and track the id for cancellation."""
        aid = self.after(ms, func)
        self._pending_after_ids.append(aid)
        return aid

    def _cancel_pending(self) -> None:
        """Cancel all scheduled `after` callbacks."""
        for aid in self._pending_after_ids:
            try:
                self.after_cancel(aid)
            except (tk.TclError, ValueError):
                pass
        self._pending_after_ids.clear()

    def _on_close(self) -> None:
        """Graceful shutdown: cancel animations, then destroy."""
        self._cancel_pending()
        self.destroy()

    # ================================================================
    # UI construction
    # ================================================================
    def _build_ui(self) -> None:
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
        self.tree.heading("Counter", text="Gezogen")
        self.tree.column("Name", width=480, anchor=tk.W)
        self.tree.column("Counter", width=120, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(middle, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        control = ttk.Frame(self)
        control.pack(fill=tk.X, padx=8, pady=(0, 8))

        ttk.Label(control, text="Anzahl ziehen:").pack(side=tk.LEFT)
        self.spin_n = ttk.Spinbox(control, from_=1, to=999, width=5)
        self.spin_n.set(1)
        self.spin_n.pack(side=tk.LEFT, padx=(4, 10))

        self.btn_draw = ttk.Button(
            control, text="Ziehung starten", command=self.on_draw_clicked, state=tk.DISABLED
        )
        self.btn_draw.pack(side=tk.LEFT)

        self.btn_clear = ttk.Button(
            control, text="Markierungen zurücksetzen",
            command=self.clear_winner_highlights, state=tk.DISABLED,
        )
        self.btn_clear.pack(side=tk.LEFT, padx=(10, 0))

        # IMPROVEMENT V2: reload button to re-read the file without file-dialog
        self.btn_reload = ttk.Button(
            control, text="⟳ Neu laden", command=self.on_reload_excel, state=tk.DISABLED
        )
        self.btn_reload.pack(side=tk.LEFT, padx=(10, 0))

        self.status = ttk.Label(self, text="Bereit.")
        self.status.pack(fill=tk.X, padx=8, pady=(0, 8))

    def _make_styles(self) -> None:
        self.tree.tag_configure("scan", background="#FFE082")    # yellow highlight
        self.tree.tag_configure("winner", background="#C8E6C9")  # green winner
        self.tree.tag_configure("normal", background="")

    def _build_menubar(self) -> None:
        menubar = tk.Menu(self)

        menu_settings = tk.Menu(menubar, tearoff=False)
        menu_settings.add_command(label="Spin-Parameter …", command=self.open_config_dialog)
        menubar.add_cascade(label="Einstellungen", menu=menu_settings)

        menu_help = tk.Menu(menubar, tearoff=False)
        menu_help.add_command(
            label="Über …",
            command=lambda: messagebox.showinfo(
                "Über",
                "Glücksrad V2 – faire Zufallsauswahl\n"
                "Spin-Dynamik, Blinken, Excel-Speicherung.\n"
                "Bug-fixes & Verbesserungen.",
            ),
        )
        menubar.add_cascade(label="Hilfe", menu=menu_help)

        self.config(menu=menubar)

    # ================================================================
    # Config Dialog
    # ================================================================
    def open_config_dialog(self) -> None:
        dlg = tk.Toplevel(self)
        dlg.title("Spin-Parameter")
        dlg.resizable(False, False)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        row = 0

        def field(label: str, val) -> ttk.Entry:
            nonlocal row
            ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8))
            e = ttk.Entry(frm, width=10)
            e.insert(0, str(val))
            e.grid(row=row, column=1, sticky="e")
            row += 1
            return e

        e_fast = field("Startgeschwindigkeit (ms pro Schritt):", self.SPIN_FAST_MS)
        e_slow = field("Endgeschwindigkeit (ms pro Schritt):", self.SPIN_SLOW_MS)
        e_grow = field("Verzögerungsfaktor (>1.0):", self.SPIN_GROW)
        e_rounds = field("Spin-Runden:", self.SPIN_ROUNDS)
        e_roll = field("Ausroll-Faktor:", self.SPIN_ROLLOUT_FACTOR)
        e_blink = field("Blinkanzahl:", self.BLINK_TIMES)
        e_bms = field("Blinktempo (ms):", self.BLINK_MS)

        row += 1
        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Abbrechen", command=dlg.destroy).pack(side=tk.RIGHT, padx=(10, 0))

        def apply() -> None:
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

        for c in (0, 1):
            frm.grid_columnconfigure(c, weight=1)

    # ================================================================
    # Excel loading / reloading
    # ================================================================
    def on_load_excel(self) -> None:
        # BUG-FIX V2: block loading while animation is running
        if self.anim_running:
            messagebox.showwarning("Bitte warten", "Bitte warten, bis die aktuelle Ziehung beendet ist.")
            return

        file_path = filedialog.askopenfilename(
            title="Excel mit Namensliste auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls")],
        )
        if not file_path:
            return
        self._load_file(Path(file_path))

    def on_reload_excel(self) -> None:
        """Re-read the currently loaded file (e.g. after external edits)."""
        if self.anim_running:
            messagebox.showwarning("Bitte warten", "Bitte warten, bis die aktuelle Ziehung beendet ist.")
            return
        if self.xls_path is None:
            return
        self._load_file(self.xls_path)

    def _load_file(self, path: Path) -> None:
        try:
            self.df = load_namelist(path)
        except Exception as e:
            log.exception("Fehler beim Laden von %s", path)
            messagebox.showerror("Fehler beim Laden", str(e))
            return
        self.xls_path = path
        self.path_label.config(text=str(self.xls_path.name))
        self.populate_tree()
        self.btn_draw.config(state=tk.NORMAL)
        self.btn_clear.config(state=tk.NORMAL)
        self.btn_reload.config(state=tk.NORMAL)
        # IMPROVEMENT V2: update Spinbox max to number of entries
        self.spin_n.config(to=len(self.df))
        self.status.config(text=f"Geladen: {len(self.df)} Einträge aus {self.xls_path.name}")
        log.info("Loaded %d names from %s", len(self.df), path)

    # ================================================================
    # Treeview helpers
    # ================================================================
    def populate_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        if self.df is None:
            return
        for i, row in self.df.iterrows():
            self.tree.insert(
                "", "end", iid=f"row-{i}",
                values=(row["Name"], int(row["Counter"])),
                tags=("normal",),
            )
        self.current_highlight_row = None

    def refresh_counters(self) -> None:
        """Update counter values in the tree without full repopulation."""
        if self.df is None:
            return
        children = set(self.tree.get_children(""))
        for i, row in self.df.iterrows():
            iid = f"row-{i}"
            if iid in children:
                self.tree.item(iid, values=(row["Name"], int(row["Counter"])))

    # ================================================================
    # Drawing logic
    # ================================================================
    def on_draw_clicked(self) -> None:
        if self.df is None or self.anim_running:
            return
        try:
            n = int(self.spin_n.get())
        except ValueError:
            messagebox.showwarning("Eingabe", "Bitte eine gültige Zahl eingeben.")
            return

        if n < 1:
            messagebox.showwarning("Eingabe", "Mindestens 1 Person muss gezogen werden.")
            return

        n = min(n, len(self.df))
        self.to_draw_total = n
        self.drawn_count = 0
        self.round_selected_idx.clear()
        self.round_excluded_idx.clear()
        self.status.config(text=f"Ziehe {n} Person(en) …")
        self.btn_draw.config(state=tk.DISABLED)
        self.anim_running = True
        self.draw_next_one()

    def draw_next_one(self) -> None:
        if self.df is None:
            self._end_round_early()
            return

        # base eligible set = all with minimal counter
        elig = eligible_indices(self.df)
        if not elig:
            normalize_counters(self.df)
            elig = eligible_indices(self.df)

        # BUG-FIX V2 safety: if still empty after normalising, bail out
        if not elig:
            log.error("No eligible indices even after normalisation.")
            self._end_round_early()
            return

        # Exclude indices already selected this batch, if alternatives remain
        elig_filtered = [i for i in elig if i not in self.round_excluded_idx]
        if not elig_filtered:
            # everyone at minimal counter already picked – allow repeats
            elig_filtered = list(elig)

        winner_idx = random.choice(elig_filtered)

        # Build animation path: fast rounds + rollout ending at winner
        path: list[int] = []
        for _ in range(self.SPIN_ROUNDS):
            path.extend(elig_filtered)

        rollout = max(1, int(len(elig_filtered) * self.SPIN_ROLLOUT_FACTOR))
        start = random.randrange(len(elig_filtered))
        cur = start
        steps = 0
        # BUG-FIX V2: add a hard upper-bound to prevent infinite loop when
        # elig_filtered has only 1 element and rollout > 1
        max_steps = rollout + len(elig_filtered) * 2
        while True:
            idx = elig_filtered[cur]
            path.append(idx)
            steps += 1
            if steps >= rollout and idx == winner_idx:
                break
            if steps >= max_steps:
                # Safety exit – append winner explicitly
                path.append(winner_idx)
                break
            cur = (cur + 1) % len(elig_filtered)

        self._animate_scan(path, winner_idx, float(self.SPIN_FAST_MS))

    def _animate_scan(self, path: list[int], winner_idx: int, delay: float) -> None:
        """Recursive animation: highlight each row in *path* with increasing delay."""
        if not path:
            self.finish_one_draw(winner_idx)
            return
        idx = path.pop(0)
        self._highlight_scan_row(idx)
        next_delay = min(delay * self.SPIN_GROW, float(self.SPIN_SLOW_MS))
        self._safe_after(
            max(5, int(next_delay)),
            lambda: self._animate_scan(path, winner_idx, next_delay),
        )

    def _highlight_scan_row(self, idx: int) -> None:
        self._clear_scan_highlight()
        iid = f"row-{idx}"
        try:
            self.tree.item(iid, tags=("scan",))
            self.tree.see(iid)
            self.current_highlight_row = iid
        except tk.TclError:
            pass

    def _clear_scan_highlight(self) -> None:
        if self.current_highlight_row:
            try:
                tags = self.tree.item(self.current_highlight_row, "tags")
                new_tag = "winner" if "winner" in tags else "normal"
                self.tree.item(self.current_highlight_row, tags=(new_tag,))
            except tk.TclError:
                pass
        self.current_highlight_row = None

    # ================================================================
    # Blink & Finish
    # ================================================================
    def _blink_row(self, iid: str, times_left: int) -> None:
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
        self._safe_after(self.BLINK_MS, lambda: self._blink_row(iid, times_left - 1))

    def finish_one_draw(self, idx: int) -> None:
        if self.df is None:
            self._end_round_early()
            return

        iid = f"row-{idx}"
        try:
            self.tree.item(iid, tags=("winner",))
            self.tree.see(iid)
        except tk.TclError:
            pass

        self.round_selected_idx.add(idx)
        self.round_excluded_idx.add(idx)

        # Blink, then apply counter & continue
        self._blink_row(iid, self.BLINK_TIMES * 2)
        blink_duration = (self.BLINK_TIMES * 2) * self.BLINK_MS + 100
        self._safe_after(blink_duration, lambda: self._apply_winner_and_continue(idx))

    def _apply_winner_and_continue(self, idx: int) -> None:
        if self.df is None:
            self._end_round_early()
            return

        self.df.at[idx, "Counter"] += 1
        normalize_counters(self.df)
        self.refresh_counters()

        self.drawn_count += 1
        name = self.df.at[idx, "Name"]
        self.status.config(
            text=f"Gezogen: {self.drawn_count}/{self.to_draw_total} – Gewinner: {name}"
        )
        log.info("Pick %d/%d: %s (idx=%d)", self.drawn_count, self.to_draw_total, name, idx)

        if self.drawn_count < self.to_draw_total:
            self._safe_after(300, self.draw_next_one)
        else:
            self._safe_after(150, self.finish_round)

    def finish_round(self) -> None:
        # Save Excel
        try:
            if self.xls_path and self.df is not None:
                save_namelist(self.df, self.xls_path)
                log.info("Saved to %s", self.xls_path)
        except Exception as e:
            log.exception("Speicherfehler")
            messagebox.showerror("Speicherfehler", str(e))

        # Summarise all winners
        winner_names = [
            self.df.at[i, "Name"] for i in sorted(self.round_selected_idx) if self.df is not None
        ]
        summary = ", ".join(winner_names) if winner_names else ""
        self.status.config(text=f"Runde beendet – Gewinner: {summary}")

        self.anim_running = False
        self.btn_draw.config(state=tk.NORMAL)
        self.round_selected_idx.clear()
        self.round_excluded_idx.clear()

    def _end_round_early(self) -> None:
        """Abort the current round gracefully (e.g. on unexpected state)."""
        self.anim_running = False
        self.btn_draw.config(state=tk.NORMAL)
        self.round_selected_idx.clear()
        self.round_excluded_idx.clear()
        self.status.config(text="Ziehung abgebrochen.")
        log.warning("Round ended early.")

    # ================================================================
    # Visual helpers
    # ================================================================
    def clear_winner_highlights(self) -> None:
        """Remove green winner highlights (visual only)."""
        for iid in self.tree.get_children(""):
            try:
                self.tree.item(iid, tags=("normal",))
            except tk.TclError:
                pass
        self._clear_scan_highlight()
        self.status.config(text="Markierungen zurückgesetzt.")


# ============================================================
# Main
# ============================================================
if __name__ == "__main__":
    random.seed()
    app = GluecksradApp()
    app.mainloop()
