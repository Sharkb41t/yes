"""
export_gui.py

Run this manually (in your Anaconda terminal / Anaconda Prompt) to open a
point-and-click window for filtering and exporting the Order Audit data:

    python export_gui.py

For each filterable column, a scrollable list of every distinct value in the
database is shown. Select one or more values in a column's list to filter on
it (leave a column with nothing selected to not filter on it at all).
Selections within a column are OR'd together; selections across different
columns are AND'd together. Click "Export to Excel" to save the matching
rows as a spreadsheet in the exports/ folder.

If the number of matching rows exceeds MAX_ROWS_PER_FILE, the export is
automatically split into multiple .xlsx files (part1, part2, ...), each
capped at MAX_ROWS_PER_FILE rows, so you stay safely under Excel's hard
limit of 1,048,576 rows per sheet while leaving headroom to add formulas,
pivot tables, or other analytics on top of the exported data.

To keep the filter dropdowns fast even on a very large database, each
filter column's distinct values are indexed and cached in a small side
table the first time you run this. After that, startup reads from the
cache instead of re-scanning the whole database. If you import new data
and want new values to show up in the filter lists, click
"Rebuild Value Cache (after new import)".

Exporting runs in a background thread and streams rows straight to disk
with xlsxwriter (constant-memory mode), instead of loading everything
into memory with pandas/openpyxl. This keeps the window responsive
("Not Responding") and avoids crashes on exports of 1,000,000+ rows.
Requires the xlsxwriter package (pip install xlsxwriter, or
conda install xlsxwriter).
"""

import os
import sys
import datetime
import traceback
import threading
import queue
import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

from db_utils import (
    TABLE_NAME,
    EXPORTS_DIR,
    DATA_START_IDX,
    FILTER_COL_INDICES,
    TRADE_DATE_COL,
    get_connection,
    get_table_columns,
    table_exists,
    idx_to_colletter,
)

# Excel's actual hard limit is 1,048,576 rows per sheet. We cap well below
# that (1,000,000 data rows per file) to leave headroom for a header row
# plus room to add your own analytics/formulas without bumping the ceiling.
MAX_ROWS_PER_FILE = 1_000_000

# Rows pulled from SQLite per batch while streaming to disk. Keeps memory
# use flat regardless of how many rows match the filters overall.
QUERY_CHUNK_SIZE = 50_000


# Name of the small side table used to cache each filter column's distinct
# values, so the GUI doesn't have to re-scan the (potentially 100M+ row)
# main table every time it starts up.
CACHE_TABLE_NAME = f"{TABLE_NAME}_filter_value_cache"


def log(msg):
    """Print a timestamped status line to the terminal and flush immediately
    so messages show up right away even if the window looks frozen."""
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def get_filter_columns(conn):
    """Return list of (db_column_name, display_label) for all filterable
    columns, including Trade Date."""
    if not table_exists(conn, TABLE_NAME):
        return []

    db_columns = get_table_columns(conn, TABLE_NAME)  # ordered, matches DATA_COL_INDICES
    filter_cols = []
    for idx in FILTER_COL_INDICES:
        pos = idx - DATA_START_IDX
        if 0 <= pos < len(db_columns):
            col_name = db_columns[pos]
            label = f"{col_name} ({idx_to_colletter(idx)})"
            filter_cols.append((col_name, label))

    filter_cols.append((TRADE_DATE_COL, TRADE_DATE_COL))
    return filter_cols


class OrderAuditExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Order Audit - Filter & Export")
        self.root.geometry("1000x650")

        log("Connecting to database...")
        self.conn = get_connection()
        log("Connected.")
        self.listboxes = {}  # col_name -> Listbox widget
        self.export_queue = queue.Queue()
        self.export_button = None  # set in _build_layout
        self.export_running = False

        self._build_layout()
        self._ensure_value_cache()
        self.refresh_data()

    def _build_layout(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=8)

        ttk.Button(top_frame, text="Refresh Values", command=self.refresh_data).pack(side="left")
        ttk.Button(top_frame, text="Clear All Filters", command=self.clear_filters).pack(side="left", padx=6)
        ttk.Button(
            top_frame, text="Rebuild Value Cache (after new import)",
            command=self.rebuild_value_cache_and_refresh,
        ).pack(side="left", padx=6)
        self.export_button = ttk.Button(top_frame, text="Export to Excel", command=self.export_to_excel)
        self.export_button.pack(side="right")

        self.status_label = ttk.Label(top_frame, text="")
        self.status_label.pack(side="right", padx=10)

        # Scrollable canvas holding one column-of-checkboxes-per-filter-column
        container = ttk.Frame(self.root)
        container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        canvas = tk.Canvas(container, borderwidth=0)
        h_scroll = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.inner_frame = ttk.Frame(canvas)

        self.inner_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def refresh_data(self):
        log("Refreshing filter values from database...")
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        self.listboxes = {}

        filter_cols = get_filter_columns(self.conn)

        if not filter_cols:
            log("No table/data found. Run import_data.py first.")
            ttk.Label(
                self.inner_frame,
                text="No data found yet. Run import_data.py first, then click Refresh Values.",
            ).pack(padx=10, pady=10)
            self.status_label.config(text="No data")
            return

        for i, (col_name, label) in enumerate(filter_cols):
            col_frame = ttk.Frame(self.inner_frame, borderwidth=1, relief="groove")
            col_frame.grid(row=0, column=i, sticky="ns", padx=4, pady=4)

            ttk.Label(col_frame, text=label, wraplength=140, justify="center").pack(pady=(4, 2))

            scrollbar = ttk.Scrollbar(col_frame, orient="vertical")
            listbox = tk.Listbox(
                col_frame,
                selectmode=tk.EXTENDED,
                exportselection=False,
                width=18,
                height=20,
                yscrollcommand=scrollbar.set,
            )
            scrollbar.config(command=listbox.yview)
            listbox.pack(side="left", fill="y", padx=(4, 0), pady=(0, 4))
            scrollbar.pack(side="left", fill="y", pady=(0, 4))

            distinct_values = self._get_cached_distinct_values(col_name)
            log(f"  Loaded {len(distinct_values)} distinct values for '{col_name}' (from cache)")
            for val in distinct_values:
                listbox.insert(tk.END, val)

            self.listboxes[col_name] = listbox

        row_count = self._get_row_count()
        log(f"Refresh complete. {row_count} total rows in database.")
        self.status_label.config(text=f"{row_count} rows in database")

    def _get_distinct_values(self, col_name):
        try:
            cur = self.conn.execute(
                f'SELECT DISTINCT "{col_name}" FROM {TABLE_NAME} '
                f'WHERE "{col_name}" IS NOT NULL AND "{col_name}" != "" '
                f'ORDER BY "{col_name}"'
            )
            return [str(r[0]) for r in cur.fetchall()]
        except Exception:
            return []

    def _get_row_count(self):
        try:
            cur = self.conn.execute(f"SELECT COUNT(*) FROM {TABLE_NAME}")
            return cur.fetchone()[0]
        except Exception:
            return 0

    # ---------------------------------------------------------------
    # Filter-value cache: avoids re-scanning a huge main table every
    # time the GUI starts. Distinct values per filter column are
    # computed once (with the help of an index on each column) and
    # stored in a small side table. Use "Rebuild Value Cache" after
    # importing new data so new values show up.
    # ---------------------------------------------------------------

    def _ensure_indexes(self, filter_cols):
        """Create an index on each filter column (if missing) so distinct
        value / lookup queries don't require a full table scan."""
        for col_name, _label in filter_cols:
            index_name = f"idx_{TABLE_NAME}_{col_name}".replace(" ", "_")
            try:
                log(f"  Ensuring index on '{col_name}'...")
                self.conn.execute(
                    f'CREATE INDEX IF NOT EXISTS "{index_name}" '
                    f'ON {TABLE_NAME} ("{col_name}")'
                )
            except Exception:
                log(f"  WARNING: could not create index on '{col_name}':")
                log(traceback.format_exc())
        self.conn.commit()

    def _cache_table_exists(self):
        cur = self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (CACHE_TABLE_NAME,),
        )
        return cur.fetchone() is not None

    def _cache_is_populated(self, filter_cols):
        """True only if the cache table exists AND already has rows for
        every current filter column (handles the case where filter
        columns changed since the cache was last built)."""
        if not self._cache_table_exists():
            return False
        cur = self.conn.execute(f"SELECT DISTINCT col_name FROM {CACHE_TABLE_NAME}")
        cached_cols = {r[0] for r in cur.fetchall()}
        return all(col_name in cached_cols for col_name, _label in filter_cols)

    def _ensure_value_cache(self):
        """Build the cache on first run only. If it already exists and
        covers all filter columns, leave it alone (fast path)."""
        filter_cols = get_filter_columns(self.conn)
        if not filter_cols:
            return
        if self._cache_is_populated(filter_cols):
            log("Filter value cache found. Using cached distinct values.")
            return
        log("No filter value cache found yet - building it now (one-time cost)...")
        self._rebuild_value_cache(filter_cols)

    def rebuild_value_cache_and_refresh(self):
        """Manually triggered: recompute the cache from the live table
        (use this after importing new data) and reload the UI."""
        filter_cols = get_filter_columns(self.conn)
        if not filter_cols:
            messagebox.showwarning("No data", "No data has been imported yet.")
            return
        self._rebuild_value_cache(filter_cols)
        self.refresh_data()
        messagebox.showinfo("Cache rebuilt", "Filter value cache has been rebuilt from the latest data.")

    def _rebuild_value_cache(self, filter_cols):
        start = datetime.datetime.now()
        log("Rebuilding filter value cache...")

        self.conn.execute(
            f"CREATE TABLE IF NOT EXISTS {CACHE_TABLE_NAME} (col_name TEXT, value TEXT)"
        )
        self._ensure_indexes(filter_cols)

        for col_name, _label in filter_cols:
            col_start = datetime.datetime.now()
            log(f"  Scanning distinct values for '{col_name}'...")
            values = self._get_distinct_values(col_name)  # live scan (indexed now)

            self.conn.execute(f"DELETE FROM {CACHE_TABLE_NAME} WHERE col_name = ?", (col_name,))
            self.conn.executemany(
                f"INSERT INTO {CACHE_TABLE_NAME} (col_name, value) VALUES (?, ?)",
                [(col_name, v) for v in values],
            )
            self.conn.commit()

            elapsed = (datetime.datetime.now() - col_start).total_seconds()
            log(f"    Cached {len(values)} values for '{col_name}' ({elapsed:.1f}s)")

        total_elapsed = (datetime.datetime.now() - start).total_seconds()
        log(f"Cache rebuild complete in {total_elapsed:.1f}s.")

    def _get_cached_distinct_values(self, col_name):
        try:
            cur = self.conn.execute(
                f"SELECT value FROM {CACHE_TABLE_NAME} WHERE col_name = ? ORDER BY value",
                (col_name,),
            )
            return [r[0] for r in cur.fetchall()]
        except Exception:
            log(f"  WARNING: could not read cache for '{col_name}', falling back to live scan.")
            return self._get_distinct_values(col_name)

    def clear_filters(self):
        for listbox in self.listboxes.values():
            listbox.selection_clear(0, tk.END)

    def _build_filter_query(self):
        where_clauses = []
        params = []

        for col_name, listbox in self.listboxes.items():
            selected_indices = listbox.curselection()
            if not selected_indices:
                continue
            selected_values = [listbox.get(i) for i in selected_indices]
            placeholders = ", ".join(["?"] * len(selected_values))
            where_clauses.append(f'"{col_name}" IN ({placeholders})')
            params.extend(selected_values)

        query = f"SELECT * FROM {TABLE_NAME}"
        if where_clauses:
            query += " WHERE " + " AND ".join(where_clauses)
        return query, params

    def export_to_excel(self):
        if self.export_running:
            messagebox.showinfo("Export in progress", "An export is already running - please wait for it to finish.")
            return

        log("Export requested.")

        if not table_exists(self.conn, TABLE_NAME):
            log("ERROR: table does not exist. No data has been imported yet.")
            messagebox.showwarning("No data", "No data has been imported yet.")
            return

        if xlsxwriter is None:
            log("ERROR: the 'xlsxwriter' package is not installed.")
            messagebox.showerror(
                "Missing package",
                "This export requires the 'xlsxwriter' package.\n\n"
                "Install it with:\n  pip install xlsxwriter\nor\n  conda install xlsxwriter\n"
                "then restart this tool.",
            )
            return

        # Build the filter query on the main thread (reads Tkinter listbox
        # selections, which isn't safe to do from a background thread).
        query, params = self._build_filter_query()
        log(f"Query: {query}")
        if params:
            log(f"  with {len(params)} filter value(s): {params}")

        os.makedirs(EXPORTS_DIR, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        self.export_running = True
        self.export_button.config(state="disabled", text="Exporting...")
        self.status_label.config(text="Exporting... 0 rows written")

        worker = threading.Thread(
            target=self._export_worker,
            args=(query, params, timestamp),
            daemon=True,
        )
        worker.start()
        self.root.after(150, self._poll_export_queue)

    def _export_worker(self, query, params, timestamp):
        """Runs on a background thread: counts matching rows, then streams
        them straight to disk in chunks via xlsxwriter (constant-memory
        mode), splitting into multiple files if needed. Never holds more
        than QUERY_CHUNK_SIZE rows in memory at once."""
        q = self.export_queue
        try:
            # Own connection: SQLite connections shouldn't be shared across
            # threads, and this keeps the main thread's connection free for
            # the UI.
            worker_conn = get_connection()

            count_query = f"SELECT COUNT(*) FROM ({query})"
            q.put(("log", f"Counting matching rows..."))
            total_rows = worker_conn.execute(count_query, params).fetchone()[0]
            q.put(("log", f"Query matches {total_rows:,} row(s)."))

            if total_rows == 0:
                q.put(("done", {"written_paths": [], "total_rows": 0}))
                return

            num_parts = max(1, -(-total_rows // MAX_ROWS_PER_FILE))  # ceil division
            if num_parts > 1:
                q.put((
                    "log",
                    f"{total_rows:,} rows exceeds the {MAX_ROWS_PER_FILE:,}-row cap per file. "
                    f"Splitting into {num_parts} files...",
                ))

            written_paths = []
            workbook = None
            worksheet = None
            columns = None
            rows_in_current_file = 0
            part_num = 0
            total_written = 0
            start_time = datetime.datetime.now()

            def open_new_part():
                nonlocal workbook, worksheet, part_num, rows_in_current_file
                part_num += 1
                if num_parts == 1:
                    out_path = os.path.join(EXPORTS_DIR, f"order_audit_export_{timestamp}.xlsx")
                else:
                    out_path = os.path.join(
                        EXPORTS_DIR,
                        f"order_audit_export_{timestamp}_part{part_num}of{num_parts}.xlsx",
                    )
                q.put(("log", f"Opening file {part_num}/{num_parts}: {out_path}"))
                wb = xlsxwriter.Workbook(out_path, {"constant_memory": True})
                ws = wb.add_worksheet()
                for col_idx, col_name in enumerate(columns):
                    ws.write(0, col_idx, col_name)
                rows_in_current_file = 0
                written_paths.append(out_path)
                return wb, ws

            for chunk in pd.read_sql_query(query, worker_conn, params=params, chunksize=QUERY_CHUNK_SIZE):
                if columns is None:
                    columns = list(chunk.columns)

                for row in chunk.itertuples(index=False, name=None):
                    if workbook is None or rows_in_current_file >= MAX_ROWS_PER_FILE:
                        if workbook is not None:
                            workbook.close()
                            q.put(("log", f"  Closed part {part_num} with {rows_in_current_file:,} data rows."))
                        workbook, worksheet = open_new_part()

                    excel_row = rows_in_current_file + 1  # +1 to skip header row
                    for col_idx, value in enumerate(row):
                        if pd.isna(value):
                            worksheet.write_blank(excel_row, col_idx, None)
                        else:
                            worksheet.write(excel_row, col_idx, value)
                    rows_in_current_file += 1
                    total_written += 1

                    if total_written % 100_000 == 0:
                        elapsed = (datetime.datetime.now() - start_time).total_seconds()
                        q.put(("progress", total_written, total_rows))
                        q.put(("log", f"  ...{total_written:,}/{total_rows:,} rows written ({elapsed:.0f}s elapsed)"))

            if workbook is not None:
                workbook.close()
                q.put(("log", f"  Closed part {part_num} with {rows_in_current_file:,} data rows."))

            total_elapsed = (datetime.datetime.now() - start_time).total_seconds()
            q.put(("log", f"Export finished in {total_elapsed:.1f}s."))
            q.put(("done", {"written_paths": written_paths, "total_rows": total_written}))

        except Exception as e:
            q.put(("log", "ERROR: export failed."))
            q.put(("log", traceback.format_exc()))
            q.put(("error", str(e)))

    def _poll_export_queue(self):
        """Runs on the main thread: drains messages from the worker thread
        and updates the UI. Tkinter widgets must only be touched here, not
        from _export_worker directly."""
        try:
            while True:
                message = self.export_queue.get_nowait()
                kind = message[0]

                if kind == "log":
                    log(message[1])
                elif kind == "progress":
                    _, written, total = message
                    self.status_label.config(text=f"Exporting... {written:,}/{total:,} rows written")
                elif kind == "done":
                    self._finish_export(message[1])
                    return  # stop polling
                elif kind == "error":
                    self._finish_export(None, error=message[1])
                    return  # stop polling
        except queue.Empty:
            pass

        if self.export_running:
            self.root.after(150, self._poll_export_queue)

    def _finish_export(self, result, error=None):
        self.export_running = False
        self.export_button.config(state="normal", text="Export to Excel")

        if error is not None:
            self.status_label.config(text="Export failed")
            messagebox.showerror("Export failed", error)
            return

        written_paths = result["written_paths"]
        total_rows = result["total_rows"]

        if not written_paths:
            self.status_label.config(text="No results")
            messagebox.showinfo("No results", "No rows matched the selected filters.")
            return

        self.status_label.config(text=f"Export complete: {total_rows:,} rows")
        log(f"Export complete. {len(written_paths)} file(s) written.")

        if len(written_paths) == 1:
            messagebox.showinfo(
                "Export complete",
                f"Exported {total_rows:,} rows to:\n{written_paths[0]}",
            )
        else:
            file_list = "\n".join(written_paths)
            messagebox.showinfo(
                "Export complete",
                f"Exported {total_rows:,} rows across {len(written_paths)} files "
                f"(max {MAX_ROWS_PER_FILE:,} rows each):\n\n{file_list}",
            )


def main():
    log("Starting Order Audit export tool...")
    try:
        root = tk.Tk()
        app = OrderAuditExportApp(root)
        log("Window ready.")
        root.mainloop()
        log("Window closed. Exiting.")
    except Exception:
        log("FATAL ERROR during startup:")
        log(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()
