
import os
import datetime
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

MAX_ROWS_PER_FILE = 1_000_000

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

        self.conn = get_connection()
        self.listboxes = {}  # col_name -> Listbox widget

        self._build_layout()
        self.refresh_data()

    def _build_layout(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=8)

        ttk.Button(top_frame, text="Refresh Values", command=self.refresh_data).pack(side="left")
        ttk.Button(top_frame, text="Clear All Filters", command=self.clear_filters).pack(side="left", padx=6)
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
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        self.listboxes = {}

        filter_cols = get_filter_columns(self.conn)

        if not filter_cols:
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

            distinct_values = self._get_distinct_values(col_name)
            for val in distinct_values:
                listbox.insert(tk.END, val)

            self.listboxes[col_name] = listbox

        self.status_label.config(text=f"{self._get_row_count()} rows in database")

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
        if not table_exists(self.conn, TABLE_NAME):
            messagebox.showwarning("No data", "No data has been imported yet.")
            return

        query, params = self._build_filter_query()
        try:
            df = pd.read_sql_query(query, self.conn, params=params)
        except Exception as e:
            messagebox.showerror("Query failed", str(e))
            return

        if df.empty:
            messagebox.showinfo("No results", "No rows matched the selected filters.")
            return

        os.makedirs(EXPORTS_DIR, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        total_rows = len(df)

        try:
            if total_rows <= MAX_ROWS_PER_FILE:
                # Fits comfortably in a single file - no "part" suffix needed.
                out_path = os.path.join(EXPORTS_DIR, f"order_audit_export_{timestamp}.xlsx")
                df.to_excel(out_path, index=False, engine="openpyxl")
                written_paths = [out_path]
            else:
                # Split into multiple files, each capped at MAX_ROWS_PER_FILE rows.
                written_paths = []
                num_parts = -(-total_rows // MAX_ROWS_PER_FILE)  # ceil division
                for part_num in range(num_parts):
                    start = part_num * MAX_ROWS_PER_FILE
                    end = start + MAX_ROWS_PER_FILE
                    chunk = df.iloc[start:end]
                    out_path = os.path.join(
                        EXPORTS_DIR,
                        f"order_audit_export_{timestamp}_part{part_num + 1}of{num_parts}.xlsx",
                    )
                    chunk.to_excel(out_path, index=False, engine="openpyxl")
                    written_paths.append(out_path)
        except Exception as e:
            messagebox.showerror("Export failed", str(e))
            return

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
    root = tk.Tk()
    app = OrderAuditExportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
