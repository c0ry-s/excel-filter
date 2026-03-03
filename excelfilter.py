"""Excel/CSV Hostname Filter Tool.

Polish notes for GitHub:
- No org-specific paths are hard-coded; set optional update paths via environment variables.
- Debug output is gated behind EXCELFILTER_DEBUG=1.
- Uses standard logging instead of scattered print() calls.

Environment variables (optional):
- EXCELFILTER_REMOTE_VERSION_PATH: path to version.txt (UNC share, URL-mounted path, etc.)
- EXCELFILTER_UPDATE_EXE_PATH: path to the latest excelfilter.exe to copy/download
- EXCELFILTER_DEBUG: set to '1' to enable debug logging

Dependencies:
- pandas
- openpyxl
- tkinterdnd2
- packaging
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import logging
import os
import subprocess
import sys
import threading
from pathlib import Path


def _set_tkdnd_path_for_frozen_app() -> None:
    """Ensure tkdnd (used by tkinterdnd2) is discoverable in PyInstaller builds.

    PyInstaller extracts bundled resources to a temp folder referenced by sys._MEIPASS.
    Tcl looks for the tkdnd package via its auto_path; setting TCLLIBPATH points Tcl
    at the bundled tkdnd directory so `package require tkdnd` succeeds.
    """
    base = getattr(sys, "_MEIPASS", None)
    if not base:
        return

    tkdnd_dir = os.path.join(base, "tkinterdnd2", "tkdnd")
    if os.path.isdir(tkdnd_dir):
        os.environ["TCLLIBPATH"] = tkdnd_dir


_set_tkdnd_path_for_frozen_app()

import pandas as pd
import re
import tkinter as tk
from openpyxl.utils import get_column_letter
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import json
import urllib.request

try:
    from packaging import version as pkg_version
except Exception:  # packaging may be missing in some frozen builds
    pkg_version = None

CONFIG_PATH = os.path.expanduser("~/.excel_filter_config")

APP_NAME = "Excel Hostname Filter"
SCRIPT_VERSION = "1.0.1"

# Optional update/version check paths. Keep these OUT of source control.
# Set these via env vars in your deployment environment.
REMOTE_VERSION_PATH = os.environ.get("EXCELFILTER_REMOTE_VERSION_PATH", "").strip()
UPDATE_EXE_SOURCE_PATH = os.environ.get("EXCELFILTER_UPDATE_EXE_PATH", "").strip()
LOGO_PATH = os.environ.get("EXCELFILTER_LOGO_PATH", "").strip()
FILTERS_URL = os.environ.get("EXCELFILTER_FILTERS_URL", "").strip()

DEBUG = os.environ.get("EXCELFILTER_DEBUG", "0").strip() == "1"

logging.basicConfig(
    level=logging.DEBUG if DEBUG else logging.INFO,
    format="[%(levelname)s] %(message)s",
)
log = logging.getLogger(APP_NAME)


# --- Logo/resource helpers ---
def resource_path(relative_path: str) -> str:
    """Return absolute path to a resource, compatible with frozen builds."""
    try:
        base_path = getattr(sys, "_MEIPASS")  # PyInstaller
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


def find_logo_path() -> Optional[str]:
    """Best-effort lookup for a bundled logo image."""
    candidates: List[str] = []

    # 1) explicit override
    if LOGO_PATH:
        candidates.append(LOGO_PATH)

    # 2) alongside script/exe resources (be tolerant of casing)
    candidates.extend([
        resource_path("rexie.png"),
        resource_path("Rexie.png"),
        resource_path("assets/rexie.png"),
        resource_path("assets/Rexie.png"),
        resource_path("Assets/rexie.png"),
        resource_path("Assets/Rexie.png"),
    ])

    # 2b) best-effort scan for any rexie*.png in assets folders
    try:
        for folder in (resource_path("assets"), resource_path("Assets")):
            if os.path.isdir(folder):
                for name in os.listdir(folder):
                    low = name.lower()
                    if low.startswith("rexie") and low.endswith(".png"):
                        candidates.append(os.path.join(folder, name))
    except Exception:
        pass

    # 3) next to the executable (common for some packagers)
    try:
        exe_dir = os.path.dirname(sys.executable)
        candidates.extend([
            os.path.join(exe_dir, "rexie.png"),
            os.path.join(exe_dir, "Rexie.png"),
            os.path.join(exe_dir, "assets", "rexie.png"),
            os.path.join(exe_dir, "assets", "Rexie.png"),
            os.path.join(exe_dir, "Assets", "rexie.png"),
            os.path.join(exe_dir, "Assets", "Rexie.png"),
        ])
    except Exception:
        pass

    for p in candidates:
        try:
            if p and os.path.exists(p):
                return p
        except Exception:
            continue

    return None

# --- Remote filter groups (optional) ---

def _validate_filter_groups_payload(payload: object) -> Optional[Dict[str, List[str]]]:
    """Return normalized filter groups dict if payload is valid, else None."""
    try:
        if not isinstance(payload, dict):
            return None

        # Accept either {"groups": {...}} or a raw groups dict
        groups_obj = payload.get("groups") if "groups" in payload else payload
        if not isinstance(groups_obj, dict):
            return None

        normalized: Dict[str, List[str]] = {}
        for group_name, codes in groups_obj.items():
            if not isinstance(group_name, str) or not group_name.strip():
                continue
            if not isinstance(codes, list):
                continue
            cleaned: List[str] = []
            for c in codes:
                if not isinstance(c, str):
                    continue
                s = c.strip().upper()
                if s:
                    cleaned.append(s)
            # de-dupe while preserving order
            seen = set()
            deduped = [x for x in cleaned if not (x in seen or seen.add(x))]
            if deduped:
                normalized[group_name.strip()] = deduped

        return normalized or None
    except Exception:
        return None


def fetch_remote_filter_groups(url: str, timeout_sec: float = 1.5) -> Optional[Dict[str, List[str]]]:
    """Fetch filter groups JSON from a URL. Returns groups dict or None."""
    if not url:
        return None

    try:
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": f"{APP_NAME}/{SCRIPT_VERSION}",
                "Accept": "application/json",
            },
        )
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            data = resp.read()
        payload = json.loads(data.decode("utf-8"))
        return _validate_filter_groups_payload(payload)
    except Exception as e:
        log.debug(f"Remote filters fetch skipped/failed: {e}")
        return None

# --- Filter Groups (embedded defaults) ---
# Keep these modular so adding a new group is just: add a list + add an entry in `get_default_filter_groups()`.

AFFILIATE_TAB_CODES: List[str] = [
    "AMCRY", "AMGRN", "AMKDF", "FPWKB", "MCCCC", "MCCWC", "MCPTH", "MCRAL", "MCRTP",
    "MCSFR", "PN011", "PN012", "PN013", "PN014", "PN015", "PN01F", "PN021", "PN025",
    "PN043", "PNAPX", "PNBRC", "PNCOM", "PNCRY", "PNFON", "PNFQV", "PNGRN", "PNHED",
    "PNHSP", "PNKDF", "PNKND", "PNMID", "PNNHS", "PNOBR", "PNPOY", "PNPTH", "PNRAL",
    "PNWAK", "PNWKF", "RXAPX", "RXBRC", "RXCRY", "RXGRN", "RXKDF", "RXKND", "RXLIL",
    "RXMOV", "RXNRA", "RXPTH", "RXSFR", "RXSNY", "RXSUN", "RXWKF",
]

REX_MAIN_CODES: List[str] = [
    "MCHMB", "MCPHC", "MCQNC", "PNBRR", "PNPOY", "RXATR", "RXATU", "RXBCC", "RXBLU",
    "RXBRR", "RXCAN", "RXDRA", "RXDUR", "RXHAV", "RXLBT", "RXLKB", "RXMAI", "RXMOB",
    "RXMOP", "RXPNC", "RXSUN", "RXWFH",
]

NASH_CODES: List[str] = [
    "MCRMT", "NDEND", "NH100", "NH300", "NHART", "NHBTR", "NHBWC", "NHCAR", "NHCOM",
    "NHCPH", "NHDAY", "NHDGR", "NHGEN", "NHHCT", "NHHRB", "NHHRT", "NHMDS", "NHMMB",
    "NHMOB", "NHNUC", "NHNWC", "NHPAV", "NHRMT", "NHWFH", "NSNHC", "PNNVL", "PNRMT",
    "PNWDL", "RXLBG", "RXWIL", "RXWLS",
]

def get_default_filter_groups() -> Dict[str, List[str]]:
    return {
        "Rex Affiliate": AFFILIATE_TAB_CODES,
        "Rex Main": REX_MAIN_CODES,
        "Nash": NASH_CODES,
    }


FILTER_GROUPS = get_default_filter_groups()

# --- Release Notes ---
# v1.0.1 (2025-08-18)
# - Fix: Per-sheet header detection during Excel processing (don’t assume row 0 headers for all sheets).
# - Fix: Smarter hostname column detection with normalized header matching and 15-char validation using a larger sample.
# - UX: If no matches are found, print a concise scan summary (files, sheets, and column checked) for easier debugging.

# - Tech: Read a fuller preview slice for validation; bumped SCRIPT_VERSION to 1.0.1.

# --- Core (UI-agnostic) helpers ---

def _norm_header(h: object) -> str:
    return str(h).strip().lower().replace(" ", "")


def detect_header_row_from_preview(df_raw: pd.DataFrame, min_non_null: int = 3) -> Optional[int]:
    """Return the likely header row index from a headerless preview dataframe."""
    try:
        for i, row in df_raw.iterrows():
            if row.count() >= min_non_null:
                return int(i)
    except Exception:
        return None
    return None


def read_csv_headers(path: str) -> List[str]:
    df_preview = pd.read_csv(path, dtype=str, nrows=0)
    return sorted([str(h).strip() for h in df_preview.columns if pd.notna(h)])


def read_excel_preview(
    xls: pd.ExcelFile,
    sheet_name: str,
    preview_rows: int = 10,
    sample_rows: int = 200,
) -> Tuple[pd.DataFrame, List[str], int]:
    """Read an Excel sheet, auto-detect header row, then return (sample_df, headers, header_row_index)."""
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, header=None, nrows=preview_rows)
    header_row_index = detect_header_row_from_preview(df_raw)
    if header_row_index is None:
        raise ValueError("Could not detect a valid header row")

    df_full = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, header=header_row_index)
    df_sample = df_full.head(sample_rows).copy()
    headers = sorted([str(h).strip() for h in df_sample.columns if pd.notna(h)])
    return df_sample, headers, header_row_index


def auto_detect_hostname_column(headers: Sequence[str], df_sample: Optional[pd.DataFrame] = None) -> Optional[str]:
    """Attempt to find the hostname/device-name column from headers (and optional sample data)."""
    header_map = {_norm_header(h): h for h in headers}

    # canonical names first
    for key in ("hostname", "hostName", "host name", "computername", "computer name", "devicename", "device name", "host", "name"):
        k = _norm_header(key)
        if k in header_map:
            return header_map[k]

    # fallback: partial substring match; validate generic Name columns via 15-char heuristic if sample provided
    for h in headers:
        lowered = _norm_header(h)
        if "host" in lowered or lowered.endswith("name"):
            if lowered == "name" and df_sample is not None and h in df_sample.columns:
                try:
                    sample_vals = df_sample[h].dropna().astype(str)
                    n = len(sample_vals)
                    if n > 0:
                        count_15 = sample_vals.map(lambda s: len(s.strip())).eq(15).sum()
                        if count_15 / n >= 0.6:
                            return h
                except Exception as e:
                    log.debug(f"Failed to validate 'Name' column: {e}")
            else:
                return h

    return None


@dataclass
class ScanResult:
    matches: List[Dict[str, object]]
    files_scanned: int
    sheets_scanned: int


def scan_files_for_matches(
    file_paths: Sequence[str],
    col_name: str,
    targets: Sequence[str],
    excel_preview_rows: int = 10,
) -> ScanResult:
    """Scan CSV/Excel files for rows where `col_name` contains any of `targets` (case-insensitive)."""
    matches: List[Dict[str, object]] = []
    sheets_scanned = 0

    # Build a safe regex that matches any of the target tokens (case-insensitive)
    targets_l = [t.lower() for t in targets]
    pattern = "|".join(re.escape(t) for t in targets_l if t)

    for file_path in file_paths:
        try:
            base = os.path.basename(file_path)

            if file_path.lower().endswith(".csv"):
                # Cheap pass: read only the target column to see if there are any matches.
                try:
                    df_col = pd.read_csv(file_path, dtype=str, usecols=[col_name])
                except Exception:
                    # Fallback: if usecols fails (header variance), read full.
                    df_col = pd.read_csv(file_path, dtype=str)
                    if col_name not in df_col.columns:
                        continue
                    df_col = df_col[[col_name]]

                series = df_col[col_name].fillna("").astype(str).str.lower()
                mask = series.str.contains(pattern, regex=True, na=False) if pattern else series.astype(bool)

                if not mask.any():
                    continue

                match_idx = df_col.index[mask]

                # Full read only if needed, then extract matching rows
                df_full = pd.read_csv(file_path, dtype=str)
                if col_name not in df_full.columns:
                    continue

                rows = df_full.loc[match_idx].copy()
                rows["File"] = base
                rows["Sheet"] = "CSV"
                matches.extend(rows.to_dict("records"))

            else:
                xls = pd.ExcelFile(file_path)
                for sheet in xls.sheet_names:
                    sheets_scanned += 1
                    try:
                        df_head = pd.read_excel(
                            xls,
                            sheet_name=sheet,
                            dtype=str,
                            header=None,
                            nrows=excel_preview_rows,
                        )
                        header_row_index = detect_header_row_from_preview(df_head)
                        if header_row_index is None:
                            continue

                        # Cheap pass: read only the target column for this sheet (huge win when most sheets don't match)
                        try:
                            df_col = pd.read_excel(
                                xls,
                                sheet_name=sheet,
                                dtype=str,
                                header=header_row_index,
                                usecols=[col_name],
                            )
                        except Exception:
                            # Fallback if the column name isn't resolvable via usecols in this file
                            df_col = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=header_row_index)
                            if col_name not in df_col.columns:
                                continue
                            df_col = df_col[[col_name]]

                        series = df_col[col_name].fillna("").astype(str).str.lower()
                        mask = series.str.contains(pattern, regex=True, na=False) if pattern else series.astype(bool)

                        if not mask.any():
                            continue

                        match_idx = df_col.index[mask]

                        # Full read only when we know there are matches, then extract matching rows
                        df_full = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=header_row_index)
                        if col_name not in df_full.columns:
                            continue

                        rows = df_full.loc[match_idx].copy()
                        rows["File"] = base
                        rows["Sheet"] = sheet
                        matches.extend(rows.to_dict("records"))

                    except Exception as se:
                        log.debug(f"Skipping sheet '{sheet}' due to read error: {se}")

        except Exception as e:
            log.warning(f"Error reading {file_path}: {e}")

    return ScanResult(matches=matches, files_scanned=len(file_paths), sheets_scanned=sheets_scanned)


def write_matches_to_excel(out_path: str, matches: Sequence[Dict[str, object]]) -> None:
    """Write matches grouped by source file+sheet into an Excel workbook."""
    from collections import defaultdict

    grouped = defaultdict(list)
    for row in matches:
        key = f"{str(row.get('File', '')).replace('.xlsx', '').replace('.csv', '')}_{row.get('Sheet', '')}"
        grouped[key].append(row)

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        for sheet_name, rows in grouped.items():
            df = pd.DataFrame(rows).drop(columns=["File", "Sheet"], errors="ignore")
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
            worksheet = writer.sheets[safe_name]
            for col_idx, column in enumerate(df.columns, 1):
                max_length = max(df[column].astype(str).map(len).max(), len(str(column))) + 2
                worksheet.column_dimensions[get_column_letter(col_idx)].width = max_length

def check_version(root):
    try:
        # If no update channel configured, skip quietly (GitHub-safe default).
        if not REMOTE_VERSION_PATH:
            log.debug("No REMOTE_VERSION_PATH set; skipping version check.")
            return

        log.debug(f"Checking version at: {REMOTE_VERSION_PATH}")
        if not os.path.exists(REMOTE_VERSION_PATH):
            log.warning("Version file not found; skipping version check.")
            return

        with open(REMOTE_VERSION_PATH, 'r') as f:
            latest_version_str = f.readline().strip()

        if not latest_version_str:
            log.error("Version file is empty.")
            return

        if pkg_version is None:
            log.debug("packaging not available; skipping semantic version compare.")
            return

        current = pkg_version.parse(SCRIPT_VERSION)
        latest = pkg_version.parse(latest_version_str)
        log.debug(f"Current version: {current}, Latest version: {latest}")

        if latest > current:
            response = messagebox.askyesno(
                "Update Available",
                f"A new version ({latest_version_str}) is available.\n\nDownload and run it now?"
            )
            if response:
                try:
                    if not UPDATE_EXE_SOURCE_PATH:
                        messagebox.showerror(
                            "Update Error",
                            "Update is available, but EXCELFILTER_UPDATE_EXE_PATH is not set.\n\n"
                            "Ask your admin to configure the update source path.",
                        )
                        return

                    shared_file = UPDATE_EXE_SOURCE_PATH
                    save_path = filedialog.asksaveasfilename(
                        title="Save Updated File As",
                        defaultextension=".exe",
                        filetypes=[("Executable Files", "*.exe")],
                        initialfile="excelfilter.exe"
                    )
                    if not save_path:
                        return  # user cancelled

                    popup = tk.Toplevel(root)
                    popup.title("Downloading Update")
                    popup.geometry("320x100")
                    popup.resizable(False, False)
                    popup.grab_set()
                    popup.configure(bg="#f2f2f2")

                    tk.Label(popup, text="Downloading new version...", bg="#f2f2f2", fg="#333333",
                             font=("SF Pro Text", 11)).pack(pady=(15, 5))
                    progress_bar = ttk.Progressbar(popup, mode="indeterminate", length=250)
                    progress_bar.pack(pady=(0, 15))
                    progress_bar.start(10)

                    def threaded_download():
                        try:
                            with open(shared_file, 'rb') as src, open(save_path, 'wb') as dst:
                                while True:
                                    chunk = src.read(4096)
                                    if not chunk:
                                        break
                                    dst.write(chunk)
                            popup.destroy()
                            messagebox.showinfo("Download Complete", f"New version saved to:\n{save_path}")
                            if sys.platform == "win32":
                                subprocess.Popen([save_path], shell=True)
                            sys.exit(0)
                        except Exception as e:
                            popup.destroy()
                            log.error(f"Failed to download: {e}")
                            messagebox.showerror("Update Error", f"Failed to download update:\n{e}")

                    threading.Thread(target=threaded_download, daemon=True).start()
                except Exception as e:
                    log.error(f"Failed to launch threaded downloader: {e}")
                    messagebox.showerror("Update Error", f"Unexpected error:\n{e}")
        else:
            log.debug("Script is up to date.")
    except Exception:
        log.exception("Version check failed")

log.info(f"Running {APP_NAME} version {SCRIPT_VERSION}")

# --- Main Application Class ---
class HostnameFilterApp:
    def _rebuild_filter_menu(self, groups: Dict[str, List[str]]) -> None:
        """Rebuild the OptionMenu items from `groups` and preserve current selection when possible."""
        try:
            current = self.filter_var.get().strip()
            keys = list(groups.keys())
            if not keys:
                return

            if current not in groups:
                current = keys[0]
                self.filter_var.set(current)

            self._filter_menu_menu.delete(0, "end")
            for k in keys:
                self._filter_menu_menu.add_command(label=k, command=lambda v=k: self.filter_var.set(v))
        except Exception as e:
            log.debug(f"Failed to rebuild filter menu: {e}")


    def refresh_filters_async(self) -> None:
        """Best-effort: fetch remote filter groups and update UI if valid. No local cache."""
        if not FILTERS_URL:
            log.debug("No FILTERS_URL set; using embedded filter groups.")
            return

        def apply(groups: Optional[Dict[str, List[str]]]):
            if not groups:
                return

            # Update global FILTER_GROUPS in-place so existing references remain valid.
            try:
                FILTER_GROUPS.clear()
                FILTER_GROUPS.update(groups)
            except Exception:
                pass

            self._rebuild_filter_menu(FILTER_GROUPS)
            log.info(f"Loaded remote filter groups from {FILTERS_URL} ({len(FILTER_GROUPS)} group(s)).")

        def thread_target():
            groups = fetch_remote_filter_groups(FILTERS_URL)
            self.root.after(0, lambda: apply(groups))

        threading.Thread(target=thread_target, daemon=True).start()
    def __init__(self, root):
        # Title/Window settings
        self.root = root
        self.root.title(f"Rex Excel Filter v{SCRIPT_VERSION} • c0ry_s")
        self.root.geometry("720x560")
        self.root.minsize(720, 560)
        self.root.configure(bg="#f2f2f2")

        modern_font = ("SF Pro Text", 12)

        # Header (optional logo + title)
        self.header_frame = tk.Frame(root, bg=self.root["bg"])
        self.header_frame.pack(pady=(12, 4))

        self._logo_img = None  # keep reference; Tk will garbage-collect images otherwise
        self.logo_label = None
        try:
            logo_file = find_logo_path()
            if logo_file:
                img = tk.PhotoImage(file=logo_file)

                # --- Auto-scale logo to a reasonable max height ---
                max_height = 120 # tweak if you want slightly larger/smaller
                img_height = img.height()

                if img_height > max_height and img_height > 0:
                    scale_factor = max(1, img_height // max_height)
                    img = img.subsample(scale_factor, scale_factor)

                self._logo_img = img  # keep reference
                self.logo_label = tk.Label(self.header_frame, image=self._logo_img, bg=self.root["bg"])
                self.logo_label.pack(pady=(0, 4))
        except Exception as e:
            log.debug(f"Logo load skipped: {e}")

        self.title_label = tk.Label(
            self.header_frame,
            text=f"{APP_NAME}",
            bg=self.root["bg"],
            fg="#333333",
            font=("SF Pro Text", 16, "bold"),
        )
        self.title_label.pack()

        # Instruction label
        self.instructions = tk.Label(
            root,
            text="📂 Drag and drop one or more Excel or CSV files into the box below.\n1) Choose a filter group.  2) Click 🚀 Process Files.  3) Save the results.",
            bg="#f2f2f2",
            fg="#333333",
            font=modern_font,
            justify="center",
        )
        self.instructions.pack(pady=(6, 10))

        # Filter Group Selector Label
        self.group_label = tk.Label(root, text="Select Filter Group:",
                                    bg="#f2f2f2", fg="#333333", font=modern_font)
        self.group_label.pack(pady=(6, 0))

        # Filter Group Dropdown
        self.filter_var = tk.StringVar(value="Rex Affiliate")
        self.filter_menu = tk.OptionMenu(root, self.filter_var, *FILTER_GROUPS.keys())
        self.filter_menu.config(bg="#ffffff", fg="#333333", activebackground="#dddddd",
                                relief="flat", highlightthickness=1, font=modern_font)
        self.filter_menu.pack(pady=4)
        # Keep a handle to the OptionMenu's internal menu so we can theme it too (macOS needs this).
        self._filter_menu_widget = self.filter_menu
        self._filter_menu_menu = self._filter_menu_widget["menu"]
        self._filter_menu_menu.delete(0, "end")  # ensure clean state
        self._rebuild_filter_menu(FILTER_GROUPS)

        self.dark_mode = False

        # macOS Aqua buttons can ignore bg/fg and render a bright face. Use a custom pill control.
        self.toggle_pill = tk.Frame(
            root,
            bg="#e0e0e0",
            highlightthickness=0,
            bd=0,
        )
        self.toggle_pill.pack(pady=(0, 8))

        self.toggle_label = tk.Label(
            self.toggle_pill,
            text="🌓 Toggle Dark Mode",
            bg="#e0e0e0",
            fg="#111111",
            font=("SF Pro Text", 10),
            padx=10,
            pady=3,
        )
        self.toggle_label.pack()

        # Click anywhere on the pill to toggle
        self.toggle_pill.bind("<Button-1>", lambda _e: self.toggle_theme())
        self.toggle_label.bind("<Button-1>", lambda _e: self.toggle_theme())

        # Optional hover effect
        def _pill_enter(_e):
            self.toggle_label.configure(cursor="hand2")

        def _pill_leave(_e):
            self.toggle_label.configure(cursor="")

        self.toggle_pill.bind("<Enter>", _pill_enter)
        self.toggle_pill.bind("<Leave>", _pill_leave)
        self.toggle_label.bind("<Enter>", _pill_enter)
        self.toggle_label.bind("<Leave>", _pill_leave)

        # Drop area frame and listbox
        self.drop_frame = tk.Frame(root, bg=self.root["bg"])
        self.drop_frame.pack(pady=(10, 14), anchor="center")
        self.drop_area = tk.Listbox(self.drop_frame, width=60, height=9, selectmode=tk.SINGLE,
                                    bg="#ffffff", fg="#444444", relief="flat", borderwidth=1,
                                    font=("Menlo", 11))
        self.drop_area.pack()
        placeholder = "⬇️ Drop Files Here ⬇️"
        padding = (self.drop_area["width"] - len(placeholder)) // 2 
        # Vertically and horizontally center the placeholder
        blank_lines = [""] * ((self.drop_area["height"] // 2) - 1)
        for line in blank_lines:
            self.drop_area.insert(tk.END, line)
        self.drop_area.insert(tk.END, f"{' ' * padding}{placeholder}")
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind("<<Drop>>", self.handle_drop)

        # Process button setup
        self.process_button = tk.Button(
            root,
            text="🚀 Process Files",
            command=self.process_files,
            state="disabled",
            bg="#007AFF",
            fg="white",
            activebackground="#005BBB",
            disabledforeground="#bdbdbd",
            relief="flat",
            font=("SF Pro Text", 12, "bold"),
            padx=10,
            pady=5,
        )
        self.process_button.pack(pady=(14, 14))

        self.status_label = tk.Label(root, text="", bg="#f2f2f2", fg="#333333", font=("SF Pro Text", 10))
        self.status_label.pack(pady=(0, 8))

        # Dropped files list initialization
        self.dropped_files = []

        # Progress popup state
        self._progress_popup = None
        self._progress_bar = None
        self._progress_label = None

        self.load_theme_preference()
        self.root.after(250, self.refresh_filters_async)

        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def load_theme_preference(self):
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, 'r') as f:
                    mode = f.read().strip().lower()
                    if mode == "dark":
                        self.dark_mode = False  # flip to trigger update
                        self.toggle_theme()
            except Exception as e:
                log.debug(f"Failed to read theme config: {e}")

    def save_theme_preference(self):
        try:
            with open(CONFIG_PATH, 'w') as f:
                f.write("dark" if self.dark_mode else "light")
        except Exception as e:
            log.debug(f"Failed to write theme config: {e}")

    def handle_drop(self, event):
        # Processing dropped files
        files = self.root.tk.splitlist(event.data)
        # Filtering for Excel and CSV files
        self.dropped_files = [f for f in files if f.lower().endswith((".xlsx", ".xls", ".csv"))]
        self.drop_area.delete(0, tk.END)
        for f in self.dropped_files:
            self.drop_area.insert(tk.END, os.path.basename(f))
        # Updating the listbox and enabling the button
        if self.dropped_files:
            self.process_button.config(state="normal")
        else:
            self.drop_area.insert(tk.END, "No valid Excel or CSV files found.")
            self.process_button.config(state="disabled")

    def process_files(self):
        if not self.dropped_files:
            messagebox.showerror("Error", "No files to process.")
            return

        first_file = self.dropped_files[0]

        # ---- UI step: determine headers (and a sample df when possible) ----
        df_sample = None
        headers: List[str] = []
        selected_sheet = None

        try:
            if first_file.lower().endswith(".csv"):
                headers = read_csv_headers(first_file)
            else:
                xls_preview = pd.ExcelFile(first_file)
                selected_sheet = xls_preview.sheet_names[0]

                if len(xls_preview.sheet_names) > 1:
                    sheet_selection = tk.Toplevel(self.root)
                    sheet_selection.title("Select Sheet")
                    sheet_selection.geometry("300x200")
                    sheet_selection.grab_set()
                    tk.Label(
                        sheet_selection,
                        text="Multiple sheets found.\nSelect the one with header data:",
                        font=("SF Pro Text", 10),
                    ).pack(pady=10)

                    sheet_var = tk.StringVar(value=selected_sheet)
                    sheet_dropdown = ttk.Combobox(
                        sheet_selection,
                        textvariable=sheet_var,
                        values=xls_preview.sheet_names,
                        state="readonly",
                    )
                    sheet_dropdown.pack(pady=10)

                    confirmed_sheet: List[str] = []

                    def on_confirm_sheet():
                        confirmed_sheet.append(sheet_var.get())
                        sheet_selection.destroy()

                    tk.Button(sheet_selection, text="OK", command=on_confirm_sheet).pack(pady=10)
                    sheet_selection.wait_window()

                    if confirmed_sheet:
                        selected_sheet = confirmed_sheet[0]

                df_sample, headers, _ = read_excel_preview(xls_preview, selected_sheet)

            log.debug(f"Detected headers: {headers}")
            messagebox.showinfo("Detected Columns", "Columns in Row 1:\n\n" + "\n".join(headers))
        except Exception as e:
            messagebox.showerror("Error", f"Could not read headers from the first file:\n{e}")
            return

        if not headers:
            messagebox.showerror("Error", "No columns found in the file. Please check that the file has headers.")
            return

        # ---- Loop: choose column -> scan -> optional retry ----
        while True:
            # ---- Core logic: auto-detect hostname column ----
            col_name = auto_detect_hostname_column(headers, df_sample=df_sample)

            # ---- UI fallback: manual column selection ----
            if not col_name:
                col_selection = tk.Toplevel(self.root)
                col_selection.title("Select Column")
                col_selection.geometry("320x220")
                col_selection.grab_set()
                tk.Label(
                    col_selection,
                    text="Could not detect hostname column automatically.\nSelect one from the available columns below:",
                    font=("SF Pro Text", 10),
                    justify="center",
                    wraplength=280,
                ).pack(pady=(15, 5))

                columns_frame = tk.Frame(col_selection)
                columns_frame.pack(pady=(0, 7))
                tk.Label(columns_frame, text="Available columns:", font=("SF Pro Text", 9, "italic"), fg="#666666").pack(anchor="w")
                columns_listbox = tk.Listbox(
                    columns_frame,
                    height=min(len(headers), 5),
                    width=38,
                    font=("Menlo", 9),
                    bg="#f7f7f7",
                    fg="#333333",
                    borderwidth=0,
                    highlightthickness=0,
                )
                for h in headers:
                    columns_listbox.insert(tk.END, h)
                columns_listbox.pack()

                col_var = tk.StringVar(value=headers[0])
                col_dropdown = ttk.Combobox(col_selection, textvariable=col_var, values=headers, state="readonly", width=35)
                col_dropdown.pack(pady=(5, 15))

                confirmed_col: List[str] = []

                def on_confirm():
                    confirmed_col.append(col_var.get())
                    col_selection.destroy()

                tk.Button(col_selection, text="OK", command=on_confirm).pack()
                col_selection.update_idletasks()
                col_selection.wait_window()

                if not confirmed_col:
                    # user closed/cancelled
                    self.status_label.config(text="⚠️ Column selection cancelled.")
                    return

                if confirmed_col[0] not in headers:
                    messagebox.showerror("Error", "Invalid or unknown column name selected.")
                    return

                col_name = confirmed_col[0]

                # Heuristic hostname validation (UI decision)
                sample_vals = None
                try:
                    if df_sample is not None and col_name in df_sample.columns:
                        sample_vals = df_sample[col_name].dropna().astype(str)
                except Exception:
                    sample_vals = None

                if sample_vals is not None:
                    non_null_count = len(sample_vals)
                    exact_15_char_count = sample_vals.map(len).eq(15).sum() if non_null_count else 0

                    if non_null_count == 0:
                        messagebox.showerror("Error", f"The selected column '{col_name}' has no data.")
                        return

                    if exact_15_char_count / non_null_count < 0.6:
                        proceed = messagebox.askyesno(
                            "Column Validation Warning",
                            f"In column '{col_name}', only {exact_15_char_count} out of {non_null_count} rows are exactly 15 characters long.\n\n"
                            "This may not be a valid hostname column.\n\nContinue anyway?",
                        )
                        if not proceed:
                            # loop back to select a different column
                            continue

            # ---- UI: run scan + save results ----
            selected_group = self.filter_var.get()
            targets = FILTER_GROUPS[selected_group]

            try:
                result = self._run_bg_with_progress(
                    "Scanning",
                    "Scanning files for matches...",
                    lambda: scan_files_for_matches(self.dropped_files, col_name, targets),
                )
            except Exception as e:
                messagebox.showerror("Error", f"Scan failed:\n{e}")
                return

            if result.matches:
                default_dir = os.path.dirname(self.dropped_files[0])
                out_path = filedialog.asksaveasfilename(
                    title="Save Results As",
                    initialdir=default_dir,
                    initialfile="Filtered_Hostnames.xlsx",
                    defaultextension=".xlsx",
                    filetypes=[("Excel Workbook", "*.xlsx")],
                )
                if not out_path:
                    self.status_label.config(text="⚠️ Save cancelled. No output file was written.")
                    return

                try:
                    self._run_bg_with_progress(
                        "Writing Results",
                        "Writing output workbook...",
                        lambda: write_matches_to_excel(out_path, result.matches),
                    )
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to write output file:\n{e}")
                    return

                self.status_label.config(text=f"✅ {len(result.matches)} matches found across {result.files_scanned} file(s).")
                messagebox.showinfo("Done", f"Results saved to:\n{out_path}")
                return

            # no matches
            self.status_label.config(text="⚠️ No matching hostnames were found.")
            try:
                scanned_summary = []
                for f in self.dropped_files:
                    if f.lower().endswith(".csv"):
                        scanned_summary.append(f"{os.path.basename(f)}: CSV — column '{col_name}'")
                    else:
                        xls_tmp = pd.ExcelFile(f)
                        scanned_summary.append(
                            f"{os.path.basename(f)}: sheets → {', '.join(xls_tmp.sheet_names)}; column '{col_name}'"
                        )
                log.debug("Scan summary:\n" + "\n".join(scanned_summary))
            except Exception as dbg_e:
                log.debug(f"Failed to build scan summary: {dbg_e}")

            retry = messagebox.askyesno(
                "No Matches Found",
                "No matching hostnames were found.\n\nWould you like to try a different column?",
            )
            if not retry:
                return

            # loop continues to select a different column

    def _show_progress(self, title: str, message: str) -> None:
        """Show a small modal progress popup with an indeterminate bar."""
        try:
            if self._progress_popup is not None:
                return

            popup = tk.Toplevel(self.root)
            popup.title(title)
            popup.geometry("360x120")
            popup.resizable(False, False)
            popup.transient(self.root)
            popup.grab_set()

            # Theme the popup background to match current mode
            bg = "#262626" if self.dark_mode else "#f2f2f2"
            fg = "#f2f2f2" if self.dark_mode else "#111111"
            popup.configure(bg=bg)

            lbl = tk.Label(popup, text=message, bg=bg, fg=fg, font=("SF Pro Text", 11))
            lbl.pack(pady=(18, 10))

            bar = ttk.Progressbar(popup, mode="indeterminate", length=280)
            bar.pack(pady=(0, 18))
            bar.start(10)

            # Disable main controls during work
            try:
                self.process_button.config(state="disabled")
            except Exception:
                pass

            self._progress_popup = popup
            self._progress_bar = bar
            self._progress_label = lbl
        except Exception as e:
            log.debug(f"Failed to show progress popup: {e}")


    def _hide_progress(self) -> None:
        """Close the progress popup if present."""
        try:
            if self._progress_bar is not None:
                try:
                    self._progress_bar.stop()
                except Exception:
                    pass
            if self._progress_popup is not None:
                try:
                    self._progress_popup.grab_release()
                except Exception:
                    pass
                try:
                    self._progress_popup.destroy()
                except Exception:
                    pass
        finally:
            self._progress_popup = None
            self._progress_bar = None
            self._progress_label = None

            # Re-enable main controls if files are present
            try:
                if self.dropped_files:
                    self.process_button.config(state="normal")
            except Exception:
                pass


    def _run_bg_with_progress(self, title: str, message: str, fn):
        """Run `fn()` in a worker thread while showing a modal progress popup. Returns fn()'s result."""
        done = tk.BooleanVar(value=False)
        result_box = {"result": None, "error": None}

        self._show_progress(title, message)

        def worker():
            try:
                result_box["result"] = fn()
            except Exception as e:
                result_box["error"] = e
            finally:
                # Signal completion on the main thread
                self.root.after(0, lambda: done.set(True))

        threading.Thread(target=worker, daemon=True).start()

        # This keeps the UI responsive while we wait for the worker.
        self.root.wait_variable(done)

        self._hide_progress()

        if result_box["error"] is not None:
            raise result_box["error"]

        return result_box["result"]


    def toggle_theme(self):
        self.dark_mode = not self.dark_mode

        # High-contrast theme tokens (macOS Tk can render grays darker than expected)
        if self.dark_mode:
            bg_color = "#1f1f1f"
            text_primary = "#f2f2f2"
            text_secondary = "#d0d0d0"
            menu_bg = "#333333"
            menu_active_bg = "#4a4a4a"
            btn_bg = "#2f6fed"
            btn_active_bg = "#2459bf"
            toggle_bg = "#444444"
            list_bg = "#262626"
        else:
            bg_color = "#f2f2f2"
            text_primary = "#111111"
            text_secondary = "#333333"
            menu_bg = "#ffffff"
            menu_active_bg = "#dddddd"
            btn_bg = "#007AFF"
            btn_active_bg = "#005BBB"
            toggle_bg = "#e0e0e0"
            list_bg = "#ffffff"

        # Window + labels
        self.root.configure(bg=bg_color)
        # Header
        self.header_frame.configure(bg=bg_color)
        self.title_label.configure(bg=bg_color, fg=text_primary)
        if self.logo_label is not None:
            self.logo_label.configure(bg=bg_color)

        # Frames that previously inherited the old bg can show as light rectangles
        self.drop_frame.configure(bg=bg_color)

        # Header uses root bg via frames/labels; instructions + other labels updated below
        self.instructions.configure(bg=bg_color, fg=text_primary)
        self.group_label.configure(bg=bg_color, fg=text_primary)
        self.status_label.configure(bg=bg_color, fg=text_secondary)

        # OptionMenu (widget + its internal menu)
        self.filter_menu.config(
            bg=menu_bg,
            fg=text_primary,
            activebackground=menu_active_bg,
            highlightthickness=1,
        )
        try:
            self._filter_menu_menu.config(
                bg=menu_bg,
                fg=text_primary,
                activebackground=menu_active_bg,
                activeforeground=text_primary,
            )
        except Exception:
            pass

        # Drop listbox
        self.drop_area.configure(bg=list_bg, fg=text_primary)

        # Buttons
        self.process_button.configure(bg=btn_bg, fg="white", activebackground=btn_active_bg)

        # Theme the custom toggle pill
        self.toggle_pill.configure(bg=toggle_bg)
        self.toggle_label.configure(bg=toggle_bg, fg=text_primary)

        self.save_theme_preference()

        # Best-effort: if a progress popup is open, re-theme it
        try:
            if self._progress_popup is not None and self._progress_label is not None:
                bg = "#262626" if self.dark_mode else "#f2f2f2"
                fg = "#f2f2f2" if self.dark_mode else "#111111"
                self._progress_popup.configure(bg=bg)
                self._progress_label.configure(bg=bg, fg=fg)
        except Exception:
            pass

# --- Run the Application ---
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.withdraw()  # Hide the main window until version check completes
    try:
        check_version(root)
    except Exception:
        log.exception("Unhandled error during version check")
    root.deiconify()  # Show main window if no update occurs
    app = HostnameFilterApp(root)
    root.mainloop()