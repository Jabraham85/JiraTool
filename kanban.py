"""
KanbanMixin — Kanban board view for fetched Jira tickets.
"""
import copy
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

from config import DEFAULT_KANBAN_COLUMNS
from storage import save_storage
from utils import debug_log, _bind_mousewheel

# Priority → colour mapping for card badges
_PRIORITY_COLORS = {
    "highest":  "#ff5630",
    "high":     "#ff7452",
    "medium":   "#ffab00",
    "low":      "#36b37e",
    "lowest":   "#00875a",
}

_CARD_BG      = "#2d2d2d"
_CARD_HOVER   = "#353535"
_COL_BG       = "#252526"
_COL_HEADER   = "#1e1e1e"
_BOARD_BG     = "#1e1e1e"
_FG           = "#d4d4d4"
_FG_DIM       = "#888888"
_ACCENT       = "#0e639c"


class KanbanMixin:
    """Mixin providing a Kanban board view."""

    # ── public view switching ──────────────────────────────────────────────

    def show_kanban_view(self):
        try:
            self.notebook.pack_forget()
        except Exception:
            pass
        try:
            self.list_frame.pack_forget()
        except Exception:
            pass
        try:
            self._list_filter_bar.pack_forget()
        except Exception:
            pass
        if not hasattr(self, "kanban_frame") or self.kanban_frame is None:
            self._build_kanban_frame()
        self.kanban_frame.pack(fill="both", expand=True, padx=8, pady=8)
        self.view_mode = "kanban"
        if hasattr(self, "_view_var"):
            self._view_var.set("kanban")
        self._populate_kanban()

    def _hide_kanban(self):
        try:
            if hasattr(self, "kanban_frame") and self.kanban_frame is not None:
                self.kanban_frame.pack_forget()
        except Exception:
            pass

    # ── frame construction ─────────────────────────────────────────────────

    def _build_kanban_frame(self):
        parent = self.notebook.master
        self.kanban_frame = ttk.Frame(parent)

        toolbar = ttk.Frame(self.kanban_frame)
        toolbar.pack(fill="x", pady=(0, 4))

        ttk.Label(toolbar, text="Kanban Board", font=("Segoe UI", 11, "bold")).pack(side="left")

        ttk.Button(toolbar, text="Configure Columns",
                   command=self._kanban_column_config_dialog).pack(side="right", padx=4)

        folder_lbl = ttk.Label(toolbar, text="Folder:")
        folder_lbl.pack(side="right", padx=(8, 2))
        self._kanban_folder_var = tk.StringVar(value="All")
        self._kanban_folder_combo = ttk.Combobox(
            toolbar, textvariable=self._kanban_folder_var,
            state="readonly", width=18,
        )
        self._kanban_folder_combo.pack(side="right", padx=(0, 4))
        self._kanban_folder_combo.bind(
            "<<ComboboxSelected>>", lambda e: self._populate_kanban())
        self._refresh_kanban_folder_combo()

        self._kanban_scroll_canvas = tk.Canvas(
            self.kanban_frame, bg=_BOARD_BG, highlightthickness=0)
        self._kanban_hscroll = ttk.Scrollbar(
            self.kanban_frame, orient="horizontal",
            command=self._kanban_scroll_canvas.xview)
        self._kanban_scroll_canvas.configure(
            xscrollcommand=self._kanban_hscroll.set)
        self._kanban_hscroll.pack(side="bottom", fill="x")
        self._kanban_scroll_canvas.pack(side="top", fill="both", expand=True)

        self._kanban_inner = ttk.Frame(self._kanban_scroll_canvas)
        self._kanban_inner_id = self._kanban_scroll_canvas.create_window(
            (0, 0), window=self._kanban_inner, anchor="nw")

        def _on_inner_cfg(e):
            self._kanban_scroll_canvas.configure(
                scrollregion=self._kanban_scroll_canvas.bbox("all"))
        self._kanban_inner.bind("<Configure>", _on_inner_cfg)

        def _on_canvas_cfg(e):
            self._kanban_scroll_canvas.itemconfig(
                self._kanban_inner_id, height=max(e.height, 1))
        self._kanban_scroll_canvas.bind("<Configure>", _on_canvas_cfg)

        self._kanban_columns_widgets = []
        self._kanban_drag_data = {}

    def _refresh_kanban_folder_combo(self):
        folders = ["All", "Unfiled"] + sorted(self.meta.get("folders", []))
        try:
            self._kanban_folder_combo["values"] = folders
        except Exception:
            pass

    # ── populate / refresh ─────────────────────────────────────────────────

    def _get_kanban_columns(self):
        cols = self.meta.get("kanban_columns")
        if not cols:
            cols = copy.deepcopy(DEFAULT_KANBAN_COLUMNS)
            self.meta["kanban_columns"] = cols
        return cols

    def _populate_kanban(self):
        for w in self._kanban_columns_widgets:
            try:
                w.destroy()
            except Exception:
                pass
        self._kanban_columns_widgets.clear()

        columns = self._get_kanban_columns()
        tickets = list(self.list_items) if hasattr(self, "list_items") else []

        # Folder filter
        folder_filter = self._kanban_folder_var.get() if hasattr(self, "_kanban_folder_var") else "All"
        if folder_filter and folder_filter != "All":
            tf = self.meta.get("ticket_folders", {})
            if folder_filter == "Unfiled":
                tickets = [t for t in tickets
                           if not tf.get(t.get("Issue key", ""))
                           and not tf.get(t.get("Issue id", ""))]
            else:
                tickets = [t for t in tickets
                           if tf.get(t.get("Issue key", "")) == folder_filter
                           or tf.get(t.get("Issue id", "")) == folder_filter]

        # Build a status → column index map
        assigned_statuses = {}
        for ci, col in enumerate(columns):
            for s in col.get("statuses", []):
                assigned_statuses[s.lower()] = ci

        # Bucket tickets
        buckets = [[] for _ in columns]
        other_bucket = []
        for t in tickets:
            st = (t.get("Status") or "").strip()
            ci = assigned_statuses.get(st.lower())
            if ci is not None:
                buckets[ci].append(t)
            else:
                other_bucket.append(t)

        # Render columns
        for ci, col in enumerate(columns):
            self._render_kanban_column(ci, col["name"], buckets[ci])

        if other_bucket:
            self._render_kanban_column(len(columns), "Other", other_bucket)

    # ── single column rendering ────────────────────────────────────────────

    def _render_kanban_column(self, col_idx, col_name, tickets):
        col_frame = tk.Frame(self._kanban_inner, bg=_COL_BG, width=290,
                             bd=0, highlightthickness=1,
                             highlightbackground="#3c3c3c")
        col_frame.pack(side="left", fill="y", padx=4, pady=4)
        col_frame.pack_propagate(False)
        col_frame.configure(width=290)
        self._kanban_columns_widgets.append(col_frame)

        header = tk.Frame(col_frame, bg=_COL_HEADER, padx=8, pady=6)
        header.pack(fill="x")
        tk.Label(header, text=f"{col_name}  ({len(tickets)})",
                 bg=_COL_HEADER, fg=_FG,
                 font=("Segoe UI", 10, "bold"),
                 anchor="w").pack(fill="x")

        canvas = tk.Canvas(col_frame, bg=_COL_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(col_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        cards_frame = tk.Frame(canvas, bg=_COL_BG)
        cards_win = canvas.create_window((0, 0), window=cards_frame, anchor="nw")

        def _on_cards_cfg(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        cards_frame.bind("<Configure>", _on_cards_cfg)

        def _on_canvas_cfg(e):
            canvas.itemconfig(cards_win, width=max(e.width, 1))
        canvas.bind("<Configure>", _on_canvas_cfg)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)

        col_frame._kanban_col_idx = col_idx
        col_frame._kanban_col_name = col_name
        canvas._kanban_col_idx = col_idx
        cards_frame._kanban_col_idx = col_idx

        for ticket in tickets:
            self._render_kanban_card(cards_frame, ticket, col_idx, canvas)

    # ── card rendering ─────────────────────────────────────────────────────

    def _render_kanban_card(self, parent, ticket, col_idx, scroll_canvas):
        key      = ticket.get("Issue key", "")
        summary  = ticket.get("Summary", "")
        priority = ticket.get("Priority", "")
        assignee = ticket.get("Assignee", "")
        sprint   = ticket.get("Sprint", "")
        labels   = ticket.get("Labels", "")
        issue_type = ticket.get("Issue Type", "")

        card = tk.Frame(parent, bg=_CARD_BG, padx=8, pady=6,
                        bd=0, highlightthickness=1,
                        highlightbackground="#3c3c3c",
                        cursor="hand2")
        card.pack(fill="x", padx=4, pady=3)
        card._ticket = ticket
        card._col_idx = col_idx

        # Row 1: key + priority badge
        row1 = tk.Frame(card, bg=_CARD_BG)
        row1.pack(fill="x")
        tk.Label(row1, text=key, bg=_CARD_BG, fg=_ACCENT,
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left")
        if priority:
            pc = _PRIORITY_COLORS.get(priority.lower(), "#888888")
            tk.Label(row1, text=f" {priority} ", bg=pc, fg="#ffffff",
                     font=("Segoe UI", 7, "bold"),
                     padx=4, pady=0).pack(side="right")
        if issue_type:
            tk.Label(row1, text=issue_type, bg=_CARD_BG, fg=_FG_DIM,
                     font=("Segoe UI", 7)).pack(side="right", padx=(0, 6))

        # Row 2: summary
        sum_lbl = tk.Label(card, text=summary, bg=_CARD_BG, fg=_FG,
                           font=("Segoe UI", 9), anchor="w",
                           wraplength=250, justify="left")
        sum_lbl.pack(fill="x", pady=(2, 2))

        # Row 3: assignee + sprint
        row3 = tk.Frame(card, bg=_CARD_BG)
        row3.pack(fill="x")
        if assignee:
            tk.Label(row3, text=assignee, bg=_CARD_BG, fg=_FG_DIM,
                     font=("Segoe UI", 8), anchor="w").pack(side="left")
        if sprint:
            tk.Label(row3, text=sprint, bg=_CARD_BG, fg=_FG_DIM,
                     font=("Segoe UI", 8), anchor="e").pack(side="right")

        # Row 4: labels
        if labels:
            tk.Label(card, text=labels, bg=_CARD_BG, fg="#569cd6",
                     font=("Segoe UI", 7), anchor="w",
                     wraplength=250, justify="left").pack(fill="x")

        # Hover effect
        all_widgets = [card] + list(card.winfo_children())
        for r in card.winfo_children():
            all_widgets.extend(r.winfo_children())

        def _enter(e):
            for w in all_widgets:
                try:
                    w.configure(bg=_CARD_HOVER)
                except Exception:
                    pass
        def _leave(e):
            for w in all_widgets:
                try:
                    w.configure(bg=_CARD_BG)
                except Exception:
                    pass

        for w in all_widgets:
            w.bind("<Enter>", _enter)
            w.bind("<Leave>", _leave)

        # Click to open ticket in tab
        def _open(e, t=ticket):
            self._kanban_open_ticket(t)
        for w in all_widgets:
            w.bind("<Button-1>", _open)

        # Right-click menu
        def _show_menu(e, t=ticket, ci=col_idx):
            self._kanban_card_context_menu(e, t, ci)
        for w in all_widgets:
            w.bind("<Button-3>", _show_menu)

        # Drag-and-drop bindings
        for w in all_widgets:
            w.bind("<ButtonPress-1>", lambda e, c=card: self._drag_start(e, c))
            w.bind("<B1-Motion>", self._drag_motion)
            w.bind("<ButtonRelease-1>", lambda e, c=card: self._drag_end(e, c))

        # Forward mousewheel to column scroll canvas
        def _fwd_scroll(e, cv=scroll_canvas):
            cv.yview_scroll(int(-1 * (e.delta / 120)), "units")
        for w in all_widgets:
            w.bind("<MouseWheel>", _fwd_scroll)

    # ── card actions ───────────────────────────────────────────────────────

    def _kanban_open_ticket(self, ticket):
        key = ticket.get("Issue key", "")
        # Check if already open in a tab
        for frame, tf in getattr(self, "tabs", {}).items():
            try:
                d = tf.read_to_dict()
                if d.get("Issue key") == key:
                    self.show_tabs_view()
                    self.notebook.select(frame)
                    return
            except Exception:
                pass
        self.show_tabs_view()
        self.new_tab(initial_data=copy.deepcopy(ticket))

    def _kanban_card_context_menu(self, event, ticket, current_col_idx):
        menu = tk.Menu(self, tearoff=0)
        columns = self._get_kanban_columns()

        move_menu = tk.Menu(menu, tearoff=0)
        for ci, col in enumerate(columns):
            if ci == current_col_idx:
                continue
            move_menu.add_command(
                label=col["name"],
                command=lambda t=ticket, c=col: self._kanban_move_ticket(t, c))
        menu.add_cascade(label="Move to...", menu=move_menu)
        menu.add_separator()
        menu.add_command(label=f"Open {ticket.get('Issue key', '')}",
                         command=lambda: self._kanban_open_ticket(ticket))
        menu.add_command(label="Add to Bundle",
                         command=lambda: self._kanban_add_to_bundle(ticket))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _kanban_move_ticket(self, ticket, target_column):
        new_status = target_column["statuses"][0] if target_column["statuses"] else ""
        if not new_status:
            return

        key = ticket.get("Issue key", "")
        ticket["Status"] = new_status

        # Update in list_items
        for item in self.list_items:
            if item.get("Issue key") == key:
                item["Status"] = new_status
                break

        # Add to bundle so the change gets uploaded
        bundle_ticket = copy.deepcopy(ticket)
        already = any(b.get("Issue key") == key for b in self.bundle)
        if already:
            for i, b in enumerate(self.bundle):
                if b.get("Issue key") == key:
                    self.bundle[i] = bundle_ticket
                    break
        else:
            self.bundle.append(bundle_ticket)
        self.update_bundle_listbox()

        try:
            save_storage(self.templates, self.meta)
        except Exception:
            pass

        self._populate_kanban()

    def _kanban_add_to_bundle(self, ticket):
        key = ticket.get("Issue key", "")
        already = any(b.get("Issue key") == key for b in self.bundle)
        if not already:
            self.bundle.append(copy.deepcopy(ticket))
            self.update_bundle_listbox()

    # ── drag-and-drop ──────────────────────────────────────────────────────

    def _drag_start(self, event, card):
        self._kanban_drag_data = {
            "card": card,
            "ticket": card._ticket,
            "start_x": event.x_root,
            "start_y": event.y_root,
            "ghost": None,
            "dragging": False,
        }

    def _drag_motion(self, event):
        dd = self._kanban_drag_data
        if not dd:
            return
        dx = abs(event.x_root - dd.get("start_x", 0))
        dy = abs(event.y_root - dd.get("start_y", 0))
        if not dd.get("dragging") and (dx > 8 or dy > 8):
            dd["dragging"] = True
            ghost = tk.Toplevel(self)
            ghost.overrideredirect(True)
            ghost.attributes("-alpha", 0.75)
            ghost.attributes("-topmost", True)
            key = dd["ticket"].get("Issue key", "")
            summary = dd["ticket"].get("Summary", "")[:40]
            lbl = tk.Label(ghost, text=f"{key}: {summary}",
                           bg="#094771", fg="#ffffff",
                           font=("Segoe UI", 9, "bold"),
                           padx=10, pady=6)
            lbl.pack()
            dd["ghost"] = ghost

        if dd.get("dragging") and dd.get("ghost"):
            dd["ghost"].geometry(f"+{event.x_root + 12}+{event.y_root + 8}")

    def _drag_end(self, event, card):
        dd = self._kanban_drag_data
        if not dd:
            return

        ghost = dd.get("ghost")
        if ghost:
            try:
                ghost.destroy()
            except Exception:
                pass

        if not dd.get("dragging"):
            self._kanban_drag_data = {}
            return

        target_col = self._find_kanban_column_at(event.x_root, event.y_root)
        if target_col is not None and target_col != card._col_idx:
            columns = self._get_kanban_columns()
            # target_col might be len(columns) for "Other" — skip that
            if target_col < len(columns):
                self._kanban_move_ticket(dd["ticket"], columns[target_col])

        self._kanban_drag_data = {}

    def _find_kanban_column_at(self, x_root, y_root):
        for w in self._kanban_columns_widgets:
            try:
                wx = w.winfo_rootx()
                wy = w.winfo_rooty()
                ww = w.winfo_width()
                wh = w.winfo_height()
                if wx <= x_root <= wx + ww and wy <= y_root <= wy + wh:
                    return getattr(w, "_kanban_col_idx", None)
            except Exception:
                pass
        return None

    # ── column configuration dialog ────────────────────────────────────────

    def _kanban_column_config_dialog(self):
        if self._focus_existing_app_dialog("kanban_config"):
            return
        columns = copy.deepcopy(self._get_kanban_columns())
        all_statuses = sorted(self.meta.get("options", {}).get("Status", []))

        dlg = tk.Toplevel(self)
        self._track_app_dialog("kanban_config", dlg)
        try:
            self._register_toplevel(dlg)
        except Exception:
            pass
        dlg.title("Kanban Column Configuration")
        dlg.geometry("650x520")
        dlg.minsize(550, 400)
        dlg.resizable(True, True)
        try:
            dlg.attributes("-topmost", True)
        except Exception:
            pass

        _BG = "#1e1e1e"
        _PANEL = "#252526"
        _BORDER = "#3c3c3c"
        dlg.configure(bg=_BG)

        tk.Label(dlg, text="Configure Kanban Columns",
                 bg=_BG, fg=_FG, font=("Segoe UI", 12, "bold")).pack(
                     anchor="w", padx=12, pady=(12, 4))
        tk.Label(dlg, text="Each column has a name and one or more Jira statuses assigned to it.",
                 bg=_BG, fg=_FG_DIM, font=("Segoe UI", 9)).pack(
                     anchor="w", padx=12, pady=(0, 8))

        # Column list
        list_frame = tk.Frame(dlg, bg=_BG)
        list_frame.pack(fill="both", expand=True, padx=12, pady=4)

        col_lb = tk.Listbox(list_frame, bg=_PANEL, fg=_FG,
                            selectbackground=_ACCENT, selectforeground="#ffffff",
                            font=("Segoe UI", 10), height=8,
                            activestyle="none", bd=0,
                            highlightthickness=1, highlightbackground=_BORDER)
        col_lb.pack(side="left", fill="both", expand=True)
        col_sb = ttk.Scrollbar(list_frame, orient="vertical", command=col_lb.yview)
        col_sb.pack(side="right", fill="y")
        col_lb.configure(yscrollcommand=col_sb.set)

        detail_frame = tk.Frame(dlg, bg=_BG)
        detail_frame.pack(fill="x", padx=12, pady=4)
        tk.Label(detail_frame, text="Assigned statuses:",
                 bg=_BG, fg=_FG, font=("Segoe UI", 9)).pack(anchor="w")
        status_lbl = tk.Label(detail_frame, text="(select a column above)",
                              bg=_BG, fg=_FG_DIM, font=("Segoe UI", 9),
                              wraplength=600, justify="left")
        status_lbl.pack(fill="x")

        def _refresh_lb():
            col_lb.delete(0, "end")
            for c in columns:
                col_lb.insert("end", f"  {c['name']}  —  {', '.join(c['statuses'])}")

        def _on_select(e=None):
            sel = col_lb.curselection()
            if not sel:
                status_lbl.config(text="(select a column above)")
                return
            idx = sel[0]
            sts = columns[idx].get("statuses", [])
            status_lbl.config(text=", ".join(sts) if sts else "(no statuses)")

        col_lb.bind("<<ListboxSelect>>", _on_select)
        _refresh_lb()

        # Buttons
        btn_row = tk.Frame(dlg, bg=_BG)
        btn_row.pack(fill="x", padx=12, pady=4)

        def _add_col():
            name = simpledialog.askstring("New Column", "Column name:", parent=dlg)
            if not name or not name.strip():
                return
            columns.append({"name": name.strip(), "statuses": []})
            _refresh_lb()

        def _rename_col():
            sel = col_lb.curselection()
            if not sel:
                return
            idx = sel[0]
            name = simpledialog.askstring(
                "Rename Column", "New name:",
                parent=dlg, initialvalue=columns[idx]["name"])
            if not name or not name.strip():
                return
            columns[idx]["name"] = name.strip()
            _refresh_lb()

        def _remove_col():
            sel = col_lb.curselection()
            if not sel:
                return
            idx = sel[0]
            columns.pop(idx)
            _refresh_lb()

        def _move_up():
            sel = col_lb.curselection()
            if not sel or sel[0] == 0:
                return
            idx = sel[0]
            columns[idx - 1], columns[idx] = columns[idx], columns[idx - 1]
            _refresh_lb()
            col_lb.selection_set(idx - 1)

        def _move_down():
            sel = col_lb.curselection()
            if not sel or sel[0] >= len(columns) - 1:
                return
            idx = sel[0]
            columns[idx + 1], columns[idx] = columns[idx], columns[idx + 1]
            _refresh_lb()
            col_lb.selection_set(idx + 1)

        def _edit_statuses():
            sel = col_lb.curselection()
            if not sel:
                return
            idx = sel[0]
            self._kanban_edit_column_statuses(dlg, columns[idx], all_statuses)
            _refresh_lb()
            col_lb.selection_set(idx)
            _on_select()

        tk.Button(btn_row, text="Add Column", command=_add_col,
                  bg="#3c3c3c", fg="#ffffff", relief="flat", padx=8).pack(side="left", padx=2)
        tk.Button(btn_row, text="Rename", command=_rename_col,
                  bg="#3c3c3c", fg="#ffffff", relief="flat", padx=8).pack(side="left", padx=2)
        tk.Button(btn_row, text="Remove", command=_remove_col,
                  bg="#3c3c3c", fg="#ffffff", relief="flat", padx=8).pack(side="left", padx=2)
        tk.Button(btn_row, text="Edit Statuses", command=_edit_statuses,
                  bg=_ACCENT, fg="#ffffff", relief="flat", padx=8).pack(side="left", padx=2)
        tk.Button(btn_row, text="Move Up", command=_move_up,
                  bg="#3c3c3c", fg="#ffffff", relief="flat", padx=8).pack(side="left", padx=2)
        tk.Button(btn_row, text="Move Down", command=_move_down,
                  bg="#3c3c3c", fg="#ffffff", relief="flat", padx=8).pack(side="left", padx=2)

        # Save / Cancel
        footer = tk.Frame(dlg, bg=_BG)
        footer.pack(fill="x", padx=12, pady=(8, 12))

        def _save():
            self.meta["kanban_columns"] = columns
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass
            dlg.destroy()
            if self.view_mode == "kanban":
                self._populate_kanban()

        tk.Button(footer, text="Save", command=_save,
                  bg=_ACCENT, fg="#ffffff", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=16, pady=4).pack(side="right", padx=4)
        tk.Button(footer, text="Cancel", command=dlg.destroy,
                  bg="#3c3c3c", fg="#ffffff", relief="flat",
                  padx=16, pady=4).pack(side="right")

    def _kanban_edit_column_statuses(self, parent_dlg, column, all_statuses):
        current = set(column.get("statuses", []))

        dlg = tk.Toplevel(parent_dlg)
        dlg.title(f"Statuses for '{column['name']}'")
        dlg.geometry("360x420")
        dlg.minsize(300, 300)
        dlg.transient(parent_dlg)
        dlg.grab_set()

        _BG = "#1e1e1e"
        _PANEL = "#252526"
        dlg.configure(bg=_BG)

        tk.Label(dlg, text=f"Select statuses for: {column['name']}",
                 bg=_BG, fg=_FG, font=("Segoe UI", 10, "bold")).pack(
                     anchor="w", padx=10, pady=(10, 4))
        tk.Label(dlg, text="Click to toggle. Checked statuses belong to this column.",
                 bg=_BG, fg=_FG_DIM, font=("Segoe UI", 8)).pack(
                     anchor="w", padx=10, pady=(0, 6))

        checks_frame = tk.Frame(dlg, bg=_BG)
        checks_frame.pack(fill="both", expand=True, padx=10, pady=4)

        canvas = tk.Canvas(checks_frame, bg=_BG, highlightthickness=0)
        sb = ttk.Scrollbar(checks_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=_BG)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        vars_map = {}
        for s in all_statuses:
            v = tk.BooleanVar(value=(s in current))
            vars_map[s] = v
            tk.Checkbutton(inner, text=s, variable=v,
                           bg=_BG, fg=_FG, selectcolor=_PANEL,
                           activebackground=_BG, activeforeground=_FG,
                           font=("Segoe UI", 10),
                           anchor="w").pack(fill="x", padx=4, pady=1)

        # Also allow adding statuses not yet in the global list
        for s in current:
            if s not in vars_map:
                v = tk.BooleanVar(value=True)
                vars_map[s] = v
                tk.Checkbutton(inner, text=f"{s} (custom)",
                               variable=v,
                               bg=_BG, fg=_FG, selectcolor=_PANEL,
                               activebackground=_BG, activeforeground=_FG,
                               font=("Segoe UI", 10),
                               anchor="w").pack(fill="x", padx=4, pady=1)

        def _ok():
            column["statuses"] = [s for s, v in vars_map.items() if v.get()]
            dlg.destroy()

        btn_f = tk.Frame(dlg, bg=_BG)
        btn_f.pack(fill="x", padx=10, pady=(4, 10))
        tk.Button(btn_f, text="OK", command=_ok,
                  bg=_ACCENT, fg="#ffffff", relief="flat",
                  padx=16, pady=4).pack(side="right", padx=4)
        tk.Button(btn_f, text="Cancel", command=dlg.destroy,
                  bg="#3c3c3c", fg="#ffffff", relief="flat",
                  padx=16, pady=4).pack(side="right")
