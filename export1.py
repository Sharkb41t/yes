
import os
import sys
import datetime
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd

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
        ttk.Button(top_frame, text="Export to Excel", command=self.export_to_excel).pack(side="right")

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
        log("Export requested.")

        if not table_exists(self.conn, TABLE_NAME):
            log("ERROR: table does not exist. No data has been imported yet.")
            messagebox.showwarning("No data", "No data has been imported yet.")
            return

        query, params = self._build_filter_query()
        log(f"Running query: {query}")
        if params:
            log(f"  with {len(params)} filter value(s): {params}")

        try:
            df = pd.read_sql_query(query, self.conn, params=params)
        except Exception as e:
            log("ERROR: query failed.")
            log(traceback.format_exc())
            messagebox.showerror("Query failed", str(e))
            return

        total_rows = len(df)
        log(f"Query returned {total_rows} row(s).")

        if df.empty:
            log("No rows matched the selected filters. Nothing to export.")
            messagebox.showinfo("No results", "No rows matched the selected filters.")
            return

        os.makedirs(EXPORTS_DIR, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        try:
            if total_rows <= MAX_ROWS_PER_FILE:
                # Fits comfortably in a single file - no "part" suffix needed.
                out_path = os.path.join(EXPORTS_DIR, f"order_audit_export_{timestamp}.xlsx")
                log(f"Writing {total_rows} rows to {out_path} ...")
                df.to_excel(out_path, index=False, engine="openpyxl")
                log(f"  Done. ({os.path.getsize(out_path):,} bytes)")
                written_paths = [out_path]
            else:
                # Split into multiple files, each capped at MAX_ROWS_PER_FILE rows.
                written_paths = []
                num_parts = -(-total_rows // MAX_ROWS_PER_FILE)  # ceil division
                log(
                    f"{total_rows} rows exceeds the {MAX_ROWS_PER_FILE:,}-row cap per file. "
                    f"Splitting into {num_parts} files..."
                )
                for part_num in range(num_parts):
                    start = part_num * MAX_ROWS_PER_FILE
                    end = start + MAX_ROWS_PER_FILE
                    chunk = df.iloc[start:end]
                    out_path = os.path.join(
                        EXPORTS_DIR,
                        f"order_audit_export_{timestamp}_part{part_num + 1}of{num_parts}.xlsx",
                    )
                    log(f"  Writing part {part_num + 1}/{num_parts}: {len(chunk)} rows -> {out_path} ...")
                    chunk.to_excel(out_path, index=False, engine="openpyxl")
                    log(f"    Done. ({os.path.getsize(out_path):,} bytes)")
                    written_paths.append(out_path)
        except Exception as e:
            log("ERROR: export failed.")
            log(traceback.format_exc())
            messagebox.showerror("Export failed", str(e))
            return

        log(f"Export complete. {len(written_paths)} file(s) written.")

        if len(written_paths) == 1:
            messagebox.showinfo(
                "Export complete",
                f"Exported {total_rows} rows to:\n{written_paths[0]}",
            )
        else:
            file_list = "\n".join(written_paths)
            messagebox.showinfo(
                "Export complete",
                f"Exported {total_rows} rows across {len(written_paths)} files "
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
