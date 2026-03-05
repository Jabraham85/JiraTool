"""
List view mixin: build list view UI, scope/filter, CSV import, helpers, event handlers,
checkbox system, refresh, actions on selected, folder management.
"""
import os
import json
import csv
import copy
import traceback
import webbrowser
import uuid
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from config import HEADERS, FETCH_FIELDS
from storage import perform_jira_request, save_storage
from utils import debug_log, _bind_mousewheel, _dedup_list_items


class ListViewMixin:
    """Mixin providing list view UI and behavior for AvalancheApp."""

    # ---------------- Build ----------------
    def _build_list_view(self):
        parent = self.notebook.master
        self.list_frame = ttk.Frame(parent)
        top = ttk.Frame(self.list_frame)
        top.pack(fill="x", padx=8, pady=(8, 0))
        ttk.Label(top, text="Search list:").pack(side="left")
        self.list_search_var = tk.StringVar()
        self.list_search_var.trace_add("write", lambda *a: self._filter_listview())
        ttk.Entry(top, textvariable=self.list_search_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Label(top, text="Scope:").pack(side="left", padx=(12, 4))
        self.list_scope_var = tk.StringVar(value="All")
        self.list_scope_var.trace_add("write", lambda *a: self._filter_listview())
        self.list_scope_combo = ttk.Combobox(top, textvariable=self.list_scope_var, width=18, state="readonly")
        self.list_scope_combo["values"] = ("All", "Assigned to me", "Created by me", "Done")
        self.list_scope_combo.pack(side="left", padx=(0, 4))
        ttk.Label(top, text="Folder:").pack(side="left", padx=(12, 4))
        self.list_folder_var = tk.StringVar(value="All")
        self.list_folder_var.trace_add("write", lambda *a: self._filter_listview())
        self.list_folder_combo = ttk.Combobox(top, textvariable=self.list_folder_var, width=14, state="readonly")
        self.list_folder_combo.pack(side="left", padx=(0, 4))
        ttk.Button(top, text="Manage Folders", width=14, command=self.manage_folders).pack(side="left", padx=2)
        btns = ttk.Frame(top)
        btns.pack(side="right")
        ttk.Button(btns, text="Import CSV into list", command=self.import_csv_to_list).pack(side="left", padx=4)
        ttk.Button(btns, text="Move to Folder", command=self.move_selected_to_folder).pack(side="left", padx=4)
        ttk.Button(btns, text="Mass Edit", command=self._mass_edit_selected).pack(side="left", padx=4)
        ttk.Button(btns, text="Open selected as Tabs", command=self.open_selected_list_as_tabs).pack(side="left", padx=4)
        ttk.Button(btns, text="Add selected to Bundle", command=self.add_selected_list_to_bundle).pack(side="left", padx=4)
        cols = ["↻", "Issue key", "Summary", "Status", "Priority", "Internal", "Issue id", "Issue Type", "Project key", "Assignee", "Labels"]
        tree_frame = ttk.Frame(self.list_frame)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(6, 8))
        self.list_tree = ttk.Treeview(tree_frame, columns=cols, show="tree headings", selectmode="extended", height=12)
        style = ttk.Style()
        style.configure("Treeview", indent=10)
        self.list_tree.column("#0", width=60, anchor="w", stretch=False, minwidth=60)
        self.list_tree.heading("#0", text="☑", command=self._toggle_all_checks)
        self._checked_tickets = set()
        _col_widths = {
            "↻": 28, "Issue key": 140, "Summary": 320, "Status": 90,
            "Priority": 80, "Internal": 80, "Issue id": 70, "Issue Type": 90,
            "Project key": 90, "Assignee": 120, "Labels": 140,
        }
        for c in cols:
            self.list_tree.heading(c, text=c)
            w = _col_widths.get(c, 100)
            self.list_tree.column(c, width=w, anchor="w")
        self.list_tree.tag_configure("folder", background="#3d5a80", foreground="#e0e0e0")
        vs = ttk.Scrollbar(tree_frame, orient="vertical", command=self.list_tree.yview)
        hs = ttk.Scrollbar(self.list_frame, orient="horizontal", command=self.list_tree.xview)
        self.list_tree.configure(yscroll=vs.set, xscroll=hs.set)
        self.list_tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        hs.pack(fill="x", padx=8)
        _bind_mousewheel(self.list_tree, "vertical")
        bottom = ttk.Frame(self.list_frame)
        bottom.pack(fill="x", padx=8, pady=6)
        ttk.Button(bottom, text="Remove Selected", command=self.remove_selected_list).pack(side="left", padx=4)
        ttk.Button(bottom, text="Clear List", command=self.clear_list).pack(side="left", padx=4)
        ttk.Button(bottom, text="Jira", command=self._open_selected_list_in_jira).pack(side="right", padx=4)
        ttk.Button(bottom, text="Open All as Tabs", command=self.open_all_list_as_tabs).pack(side="right", padx=4)
        self.list_tree.bind("<Double-1>", lambda e: self._on_list_double_click(e))
        self.list_tree.bind("<Button-1>", lambda e: self._on_list_click(e))
        self.list_tree.bind("<Button-3>", lambda e: self._on_list_right_click(e))
        self._refresh_folder_combo()

    # ---------------- Scope/filter/toggle ----------------
    def _set_list_scope(self, scope):
        """Set list view scope filter from top bar buttons."""
        try:
            if hasattr(self, "list_scope_var"):
                self.list_scope_var.set(scope)
                self._filter_listview()
        except Exception:
            pass

    def toggle_list_view(self):
        if self.view_mode == "tabs":
            self.show_list_view()
        else:
            self.show_tabs_view()

    def show_list_view(self):
        try:
            self.notebook.pack_forget()
        except Exception:
            pass
        self.list_frame.pack(fill="both", expand=True, padx=8, pady=8)
        self._list_filter_bar.pack(side="left", padx=(0, 4))
        self.view_mode = "list"
        self._populate_listview()

    def show_tabs_view(self):
        try:
            self.list_frame.pack_forget()
        except Exception:
            pass
        try:
            self._list_filter_bar.pack_forget()
        except Exception:
            pass
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.view_mode = "tabs"

    # ---------------- CSV import to list ----------------
    def import_csv_to_list(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, newline="", encoding="utf-8") as f:
                rows = list(csv.DictReader(f))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV: {e}")
            return
        if not rows:
            messagebox.showinfo("Info", "CSV has no rows.")
            return
        for r in rows:
            item = {h: (r.get(h, "") or "") for h in HEADERS}
            if "Description ADF" in r and r.get("Description ADF"):
                try:
                    item["Description ADF"] = json.loads(r.get("Description ADF"))
                except Exception:
                    pass
            if not item.get("Issue id"):
                item["Issue id"] = str(uuid.uuid4())
            if not item.get("Issue key"):
                item["Issue key"] = "LOCAL-" + item["Issue id"][:8]
            self.list_items.append(item)
        self.list_items = _dedup_list_items(self.list_items)
        self._populate_listview()
        messagebox.showinfo("Imported", f"Imported {len(rows)} rows into list view.")
        self.show_list_view()

    # ---------------- Helpers ----------------
    def _get_item_ticket_key(self, item):
        """Stable key for folder lookup (Issue id or Issue key)."""
        return item.get("Issue id") or item.get("Issue key") or ""

    def _get_item_folder(self, item):
        """Folder name for this ticket."""
        key = self._get_item_ticket_key(item)
        return self.meta.get("ticket_folders", {}).get(key, "")

    def _refresh_folder_combo(self):
        """Refresh folder dropdown with current folders."""
        current = self.list_folder_var.get() or "All"
        folders = ["All", "Unfiled"] + sorted(self.meta.get("folders", []))
        self.list_folder_combo["values"] = folders
        if current in folders:
            self.list_folder_combo.set(current)
        else:
            self.list_folder_combo.set("All")

    def _get_filtered_rows_for_listview(self):
        """Apply search, scope, and folder filter, return list of (index, item) to display."""
        q = (self.list_search_var.get() or "").strip().lower()
        folder_filter = (self.list_folder_var.get() or "All").strip()
        try:
            scope_filter = (self.list_scope_var.get() or "All").strip()
        except Exception:
            scope_filter = "All"
        cur = self.meta.get("jira_current_user") or {}
        cur_display = (cur.get("displayName") or "").strip().lower()
        cur_email = (cur.get("emailAddress") or "").strip().lower()
        result = []
        for i, r in enumerate(self.list_items):
            if q:
                vals = " ".join(str(r.get(c, "")) for c in ["Issue key", "Summary", "Status", "Priority", "Issue id", "Issue Type", "Project key", "Assignee", "Labels"])
                if q not in vals.lower():
                    continue
            if scope_filter == "Assigned to me" and (cur_display or cur_email):
                assignee = (r.get("Assignee") or "").strip().lower()
                if not assignee or not ((cur_display and cur_display in assignee) or (cur_email and cur_email in assignee)):
                    continue
            elif scope_filter == "Created by me" and (cur_display or cur_email):
                reporter = (r.get("Reporter") or "").strip().lower()
                if not reporter or not ((cur_display and cur_display in reporter) or (cur_email and cur_email in reporter)):
                    continue
            elif scope_filter == "Done":
                status = (r.get("Status") or "").strip().lower()
                status_cat = (r.get("Status Category") or "").strip().lower()
                done_statuses = ("done", "closed", "resolved", "complete", "completed", "cancelled")
                if not status and not status_cat:
                    continue
                if status_cat == "done":
                    pass
                elif status in done_statuses:
                    pass
                else:
                    continue
            item_folder = self._get_item_folder(r)
            if folder_filter and folder_filter not in ("All", ""):
                if folder_filter == "Unfiled":
                    if item_folder:
                        continue
                elif item_folder != folder_filter:
                    continue
            result.append((i, r))
        return result

    def _populate_listview(self):
        """Build tree with folder headers at top, tickets grouped underneath."""
        try:
            self.list_tree.delete(*self.list_tree.get_children())
        except Exception:
            pass
        filtered = self._get_filtered_rows_for_listview()
        cols = self.list_tree["columns"]
        folder_filter = (self.list_folder_var.get() or "All").strip()
        # Group by folder: (folder_name, [(idx, item), ...])
        groups = {}
        for idx, r in filtered:
            folder = self._get_item_folder(r) or "Unfiled"
            groups.setdefault(folder, []).append((idx, r))
        # Use same folder order as dropdown when showing All; otherwise only the selected folder
        if folder_filter in ("All", ""):
            custom_folders = sorted(self.meta.get("folders", []))
            folder_order = custom_folders + ["Unfiled"]
        else:
            folder_order = [folder_filter]
        for i, folder_name in enumerate(folder_order):
            items = groups.get(folder_name, [])
            fid = f"folder_{i}"
            self.list_tree.insert("", "end", iid=fid, text=f"📁 {folder_name}", values=("",) * len(cols), tags=("folder",))
            for idx, r in items:
                internal = self.meta.get("internal_priorities", {}).get(r.get("Issue key") or r.get("Issue id"), "None")
                tid = f"ticket_{idx}"
                chk = "☑" if tid in self._checked_tickets else "☐"
                vals = ["↻"] + [(internal if c == "Internal" else r.get(c, "")) for c in cols if c != "↻"]
                self.list_tree.insert(fid, "end", iid=tid, text=chk, values=vals)
            self.list_tree.item(fid, open=True)

    def _filter_listview(self):
        """Rebuild list view with current search and folder filter."""
        self._populate_listview()

    def _list_iid_to_index(self, iid):
        """Map tree iid to list_items index. Returns -1 for folder rows."""
        if not iid:
            return -1
        if isinstance(iid, str) and iid.startswith("ticket_"):
            try:
                return int(iid.split("_", 1)[1])
            except Exception:
                return -1
        try:
            return int(iid)
        except Exception:
            pass
        try:
            vals = self.list_tree.item(iid, "values") or []
            cols = self.list_tree["columns"]
            if cols and "Issue id" in cols:
                idx_id = cols.index("Issue id")
                issue_id = vals[idx_id] if len(vals) > idx_id else None
            else:
                issue_id = vals[5] if len(vals) > 5 else (vals[2] if len(vals) > 2 else None)
            if issue_id:
                    for i, it in enumerate(self.list_items):
                        if (it.get("Issue id") or it.get("Issue key")) == issue_id:
                            return i
        except Exception:
            pass
        return -1

    # ---------------- Event handlers ----------------
    def _on_list_click(self, event):
        """Handle click on checkbox (tree #0) or refresh column (↻ = #1)."""
        try:
            region = self.list_tree.identify_region(event.x, event.y)
            col = self.list_tree.identify_column(event.x)
            iid = self.list_tree.identify_row(event.y)
            if not iid or str(iid).startswith("folder_"):
                return
            if region == "tree" or col == "#0":
                # Checkbox toggle (tree column)
                if iid in self._checked_tickets:
                    self._checked_tickets.discard(iid)
                else:
                    self._checked_tickets.add(iid)
                chk = "☑" if iid in self._checked_tickets else "☐"
                self.list_tree.item(iid, text=chk)
                return "break"
            if col == "#1" and region == "cell":
                # Refresh column
                idx = self._list_iid_to_index(iid)
                if idx < 0 or idx >= len(self.list_items):
                    return
                self._refresh_list_item_from_jira(idx)
        except Exception:
            pass

    def _on_list_right_click(self, event):
        sel = self.list_tree.selection()
        if not sel:
            return
        iids = [s for s in sel if s and not str(s).startswith("folder_")]
        if not iids:
            return
        menu = tk.Menu(self, tearoff=0)
        sub = tk.Menu(menu, tearoff=0)
        for lvl in self.meta.get("internal_priority_levels", ["High", "Medium", "Low", "None"]):
            sub.add_command(label=lvl, command=lambda l=lvl: self._set_internal_priority_for_selected(l))
        menu.add_cascade(label="Set Internal Priority", menu=sub)
        menu.add_separator()
        folder_sub = tk.Menu(menu, tearoff=0)
        for fname in ["Unfiled"] + sorted(self.meta.get("folders", [])):
            folder_sub.add_command(label=fname, command=lambda f=fname: self._move_selected_to_folder_by_name(f))
        menu.add_cascade(label="Move to Folder", menu=folder_sub)
        menu.add_command(label="Create folder and move here...", command=self._create_folder_and_move_selected)
        menu.add_separator()
        menu.add_command(label="Mass Edit...", command=self._mass_edit_selected)
        menu.add_separator()
        menu.add_command(label="Jira", command=self._open_selected_list_in_jira)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _on_list_double_click(self, event):
        sel = self.list_tree.selection()
        if not sel:
            return
        iid = sel[0]
        if str(iid).startswith("folder_"):
            return
        idx = self._list_iid_to_index(iid)
        if idx < 0 or idx >= len(self.list_items):
            return
        item = self.list_items[idx]
        self.show_tabs_view()
        self.new_tab(initial_data=item)

    # ---------------- Checkbox system ----------------
    def _toggle_all_checks(self):
        """Toggle all ticket checkboxes in the list view."""
        all_tickets = []
        for fid in self.list_tree.get_children():
            for tid in self.list_tree.get_children(fid):
                if not str(tid).startswith("folder_"):
                    all_tickets.append(tid)
        if not all_tickets:
            return
        all_checked = all(t in self._checked_tickets for t in all_tickets)
        for tid in all_tickets:
            if all_checked:
                self._checked_tickets.discard(tid)
            else:
                self._checked_tickets.add(tid)
            chk = "☑" if tid in self._checked_tickets else "☐"
            self.list_tree.item(tid, text=chk)

    def _get_checked_issues(self):
        """Return list of (list_items_index, item_dict) for all checked tickets."""
        result = []
        for iid in self._checked_tickets:
            idx = self._list_iid_to_index(iid)
            if 0 <= idx < len(self.list_items):
                result.append((idx, self.list_items[idx]))
        return result

    def _get_selected_or_checked_iids(self):
        """Return ticket iids from checkboxes first, falling back to tree selection."""
        checked = [iid for iid in self._checked_tickets
                   if not str(iid).startswith("folder_")]
        if checked:
            return checked
        sel = self.list_tree.selection()
        return [s for s in sel if s and not str(s).startswith("folder_")]

    # ---------------- Refresh ----------------
    def _map_issue_json_to_dict(self, issue_json: dict, base: dict = None) -> dict:
        """Canonical mapping from a Jira issue JSON response to our internal dict.

        Used by both auto-refresh and manual refresh so the two code paths are
        always in sync.  Pass ``base`` to merge on top of existing data (e.g.
        to preserve local-only keys like Variables, internal priority, etc.).
        """
        fields     = issue_json.get("fields", {}) or {}
        status_obj = fields.get("status") or {}

        result = dict(base) if base else {}

        result["Issue key"]       = issue_json.get("key", "") or result.get("Issue key", "")
        result["Issue id"]        = issue_json.get("id",  "") or result.get("Issue id",  "")
        result["Summary"]         = fields.get("summary", "") or ""
        result["Issue Type"]      = (fields.get("issuetype") or {}).get("name", "")
        result["Status"]          = status_obj.get("name", "")
        result["Status Category"] = (status_obj.get("statusCategory") or {}).get("name", "")
        result["Project key"]     = (fields.get("project") or {}).get("key", "")
        result["Project name"]    = (fields.get("project") or {}).get("name", "")
        result["Priority"]        = (fields.get("priority") or {}).get("name", "")
        result["Assignee"]        = ((fields.get("assignee") or {}).get("displayName", "")
                                     or (fields.get("assignee") or {}).get("emailAddress", "") or "")
        result["Reporter"]        = ((fields.get("reporter") or {}).get("displayName", "")
                                     or (fields.get("reporter") or {}).get("emailAddress", "") or "")
        result["Creator"]         = ((fields.get("creator") or {}).get("displayName", "")
                                     or (fields.get("creator") or {}).get("emailAddress", "") or "")
        result["Created"]         = fields.get("created", "")
        result["Updated"]         = fields.get("updated", "")
        result["Labels"]          = "; ".join(fields.get("labels") or [])
        result["Components"]      = "; ".join(c.get("name", "") for c in (fields.get("components") or []))

        # Environment (may be plain string or ADF doc)
        env = fields.get("environment")
        if isinstance(env, dict):
            try:
                result["Environment"] = self._extract_text_from_adf(env)
            except Exception:
                result["Environment"] = ""
        elif isinstance(env, str):
            result["Environment"] = env
        else:
            result["Environment"] = ""

        # Description
        rendered_html = (issue_json.get("renderedFields") or {}).get("description")
        if rendered_html:
            result["Description Rendered"] = rendered_html
        desc = fields.get("description", "")
        if isinstance(desc, dict):
            result["Description ADF"] = desc
            try:
                result["Description"] = self._extract_text_from_adf(desc)
            except Exception:
                result.setdefault("Description", "")
        elif isinstance(desc, str):
            result["Description"] = desc
        else:
            result.setdefault("Description", "")

        # Attachments, comments, epic/parent/issue-links
        result["Attachment"] = self._jira_attachments_to_field(fields.get("attachment")) or ""
        result["Comment"]    = self._parse_jira_comments(fields.get("comment"))
        self._map_epic_and_link_fields(fields, result)

        return result

    def _auto_refresh_from_jira(self, key, initial_data, tabform=None):
        """Silently refresh a ticket from Jira in a background thread.

        The tab is shown immediately with cached data; once the Jira fetch
        completes the tab form and list entry are updated on the main thread
        without any UI blocking.

        ``tabform`` should be the TabForm instance for this ticket so the
        background thread can repopulate it directly.  If omitted the method
        falls back to searching by key (works only if the tab is already
        registered in self.tabs).
        """
        self._session_refreshed_keys.add(key)
        jira = self.meta.get("jira", {})
        if not ((jira.get("base") or "").strip()
                and (jira.get("email") or "").strip()
                and (jira.get("token") or "").strip()):
            return
        s = self.get_jira_session()
        if not s:
            return

        import threading

        def _worker():
            try:
                issue_json = self.fetch_issue_details(s, key, fields=FETCH_FIELDS)
                refreshed  = self._map_issue_json_to_dict(issue_json, base=dict(initial_data))

                def _apply():
                    # Update the stored list entry
                    for i, it in enumerate(self.list_items):
                        if (it.get("Issue key") or it.get("Issue id")) == key:
                            self.list_items[i] = dict(refreshed)
                            break
                    try:
                        self.meta["fetched_issues"] = list(self.list_items)
                        save_storage(self.templates, self.meta)
                    except Exception:
                        pass
                    # Update the tab form — prefer the directly-passed reference,
                    # fall back to a key search for the re-open-existing-tab path.
                    tf = tabform
                    if tf is None:
                        try:
                            existing = self._find_tab_by_ticket_key(key)
                            if existing:
                                tf = existing[1]
                        except Exception:
                            pass
                    if tf is not None:
                        try:
                            self._enrich_with_internal_priority(refreshed)
                            tf.populate_from_dict(refreshed)
                        except Exception:
                            pass
                    debug_log(f"Auto-refreshed {key} from Jira (background)")

                self.after(0, _apply)
            except Exception as e:
                debug_log(f"Auto-refresh failed for {key}: {e}")

        threading.Thread(target=_worker, daemon=True).start()

    def _refresh_list_item_from_jira(self, idx):
        """Refresh a single list item from Jira by index."""
        if idx < 0 or idx >= len(self.list_items):
            return
        item = self.list_items[idx]
        key_or_id = item.get("Issue key") or item.get("Issue id")
        if not key_or_id or str(key_or_id).strip().startswith("LOCAL-"):
            messagebox.showinfo("Info", "This ticket has no Jira key (local only).")
            return
        s = self.get_jira_session()
        if not s:
            messagebox.showinfo("Info", "Set Jira API credentials first.")
            return
        try:
            issue_json = self.fetch_issue_details(s, key_or_id, fields=FETCH_FIELDS)
        except Exception as e:
            messagebox.showerror("Refresh failed", str(e))
            return
        # Merge on top of existing item to preserve local-only fields
        issue_dict = self._map_issue_json_to_dict(issue_json, base=dict(item))
        self.list_items[idx] = issue_dict
        self._session_refreshed_keys.add(key_or_id)
        try:
            self.meta["fetched_issues"] = list(self.list_items)
            save_storage(self.templates, self.meta)
        except Exception:
            pass
        self._populate_listview()
        messagebox.showinfo("Refreshed", f"Updated {issue_dict.get('Issue key', '')} from Jira.")

    # ---------------- Actions on selected ----------------
    # ── Ticket link open helpers (used as callbacks for TabForm) ─────────────

    def _open_ticket_link_in_app(self, key: str):
        """Open a ticket by key: switch to existing tab or fetch from Jira."""
        try:
            item = next(
                (it for it in getattr(self, "list_items", [])
                 if (it.get("Issue key") or it.get("Issue id")) == key),
                None,
            )
            if item:
                self.show_tabs_view()
                self.new_tab(initial_data=item)
                return
            s = self.get_jira_session()
            if not s:
                messagebox.showinfo("Not found",
                    f"{key} is not in your downloaded tickets and no Jira session is active.")
                return
            try:
                issue_json = self.fetch_issue_details(s, key, fields=FETCH_FIELDS)
            except Exception as e:
                messagebox.showerror("Fetch failed", f"Could not fetch {key}:\n{e}")
                return
            refreshed = self._map_issue_json_to_dict(issue_json)
            self.show_tabs_view()
            self.new_tab(initial_data=refreshed)
        except Exception:
            import traceback
            debug_log("_open_ticket_link_in_app error: " + traceback.format_exc())

    def _open_ticket_link_in_jira(self, key: str):
        """Open the Jira browse URL for *key* in the default browser."""
        try:
            import webbrowser
            s = self.get_jira_session()
            base = getattr(s, "_jira_base", "") if s else ""
            if base:
                webbrowser.open(f"{base}/browse/{key}")
            else:
                jira_cfg = self.meta.get("jira", {})
                base_url = (jira_cfg.get("base") or "").strip()
                if base_url:
                    webbrowser.open(f"{base_url}/browse/{key}")
        except Exception:
            pass

    def _open_selected_list_in_jira(self):
        """Open checked/selected list items in Jira (browser)."""
        sel = self._get_selected_or_checked_iids()
        keys = []
        for s in sel:
            if str(s).startswith("folder_"):
                continue
            idx = self._list_iid_to_index(s)
            if 0 <= idx < len(self.list_items):
                key = self.list_items[idx].get("Issue key") or self.list_items[idx].get("Issue id")
                if key and str(key).strip() and not str(key).strip().startswith("LOCAL-"):
                    keys.append(str(key).strip())
        if not keys:
            messagebox.showinfo("Info", "Select ticket(s) with a Jira key to open.")
            return
        for k in keys:
            self._open_in_jira_browser(k)

    def _set_internal_priority_for_selected(self, priority):
        sel = self._get_selected_or_checked_iids()
        for s in sel:
            if str(s).startswith("folder_"):
                continue
            idx = self._list_iid_to_index(s)
            if idx < 0 or idx >= len(self.list_items):
                continue
            key = self.list_items[idx].get("Issue key") or self.list_items[idx].get("Issue id")
            if key:
                self._set_internal_priority(key, priority, refresh_list=False)
        self._populate_listview()

    def open_selected_list_as_tabs(self):
        sel = self._get_selected_or_checked_iids()
        if not sel:
            messagebox.showinfo("Info", "Check or select rows to open as tabs.")
            return
        count = 0
        for s in sel:
            idx = self._list_iid_to_index(s)
            if idx < 0 or idx >= len(self.list_items):
                continue
            item = self.list_items[idx]
            self.new_tab(initial_data=item)
            count += 1
        messagebox.showinfo("Opened", f"Opened {count} rows as tabs.")
        self.show_tabs_view()

    def open_all_list_as_tabs(self):
        if not self.list_items:
            messagebox.showinfo("Info", "List is empty.")
            return
        for item in list(self.list_items):
            self.new_tab(initial_data=item)
        messagebox.showinfo("Opened", f"Opened {len(self.list_items)} rows as tabs.")
        self.show_tabs_view()

    def add_selected_list_to_bundle(self):
        sel = self._get_selected_or_checked_iids()
        if not sel:
            messagebox.showinfo("Info", "Check or select rows to add to bundle.")
            return
        added = 0
        for s in sel:
            idx = self._list_iid_to_index(s)
            if idx < 0 or idx >= len(self.list_items):
                continue
            item = dict(self.list_items[idx])
            if not item.get("Issue id"):
                item["Issue id"] = str(uuid.uuid4())
            if not item.get("Issue key"):
                item["Issue key"] = "LOCAL-" + item["Issue id"][:8]
            adf = item.get("Description ADF")
            if not adf or not (isinstance(adf, dict) and adf.get("content")):
                recovered = self._recover_adf_for_ticket(item)
                if recovered:
                    item["Description ADF"] = copy.deepcopy(recovered)
            self.bundle.append(item)
            added += 1
        self.update_bundle_listbox()
        messagebox.showinfo("Bundled", f"Added {added} items to bundle.")

    def remove_selected_list(self):
        sel = self._get_selected_or_checked_iids()
        if not sel:
            messagebox.showinfo("Info", "Check or select rows to remove.")
            return
        idxs = sorted((self._list_iid_to_index(s) for s in sel), reverse=True)
        removed = 0
        for idx in idxs:
            if 0 <= idx < len(self.list_items):
                try:
                    del self.list_items[idx]
                    removed += 1
                except Exception:
                    pass
        self._populate_listview()
        messagebox.showinfo("Removed", f"Removed {removed} rows.")

    def clear_list(self):
        if not self.list_items:
            return
        if not messagebox.askyesno("Confirm", f"Clear {len(self.list_items)} rows from list?"):
            return
        self.list_items = []
        self._populate_listview()

    def move_selected_to_folder(self):
        """Move selected tickets to a folder."""
        sel = self._get_selected_or_checked_iids()
        if not sel:
            messagebox.showinfo("Info", "Check or select rows to move to a folder.")
            return
        folders = ["Unfiled"] + sorted(self.meta.get("folders", []))
        win = tk.Toplevel(self)
        self._register_toplevel(win)
        win.title("Move to Folder")
        win.minsize(320, 100)
        win.geometry("360x120")
        win.resizable(True, True)
        ttk.Label(win, text=f"Move {len(sel)} selected ticket(s) to:").pack(anchor="w", padx=8, pady=(8, 4))
        default = folders[1] if len(folders) > 1 else "Unfiled"
        folder_var = tk.StringVar(value=default)
        cb = ttk.Combobox(win, textvariable=folder_var, values=folders, state="readonly", width=30)
        cb.pack(fill="x", padx=8, pady=4)
        cb.set(default)
        def do_move():
            folder = folder_var.get().strip()
            if not folder:
                messagebox.showwarning("No folder", "Select a folder.")
                return
            self.meta.setdefault("ticket_folders", {})
            moved = 0
            for s in sel:
                idx = self._list_iid_to_index(s)
                if idx < 0 or idx >= len(self.list_items):
                    continue
                item = self.list_items[idx]
                key = self._get_item_ticket_key(item)
                if not key:
                    item["Issue id"] = str(uuid.uuid4())
                    item["Issue key"] = "LOCAL-" + item["Issue id"][:8]
                    key = item["Issue id"]
                if folder == "Unfiled":
                    self.meta["ticket_folders"].pop(key, None)
                else:
                    self.meta["ticket_folders"][key] = folder
                moved += 1
            save_storage(self.templates, self.meta)
            self._populate_listview()
            win.destroy()
            msg = f"Moved {moved} ticket(s) to folder '{folder}'." if folder != "Unfiled" else f"Removed {moved} ticket(s) from folder."
            messagebox.showinfo("Moved", msg)
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill="x", padx=8, pady=8)
        ttk.Button(btn_frame, text="Move", command=do_move).pack(side="right", padx=4)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side="right")

    def _move_selected_to_folder_by_name(self, folder_name):
        """Move selected list items to the given folder."""
        sel = self.list_tree.selection()
        if not sel:
            return
        self.meta.setdefault("ticket_folders", {})
        moved = 0
        for s in sel:
            if str(s).startswith("folder_"):
                continue
            idx = self._list_iid_to_index(s)
            if idx < 0 or idx >= len(self.list_items):
                continue
            item = self.list_items[idx]
            key = item.get("Issue id") or item.get("Issue key")
            if not key:
                continue
            if folder_name == "Unfiled":
                self.meta["ticket_folders"].pop(key, None)
            else:
                self.meta["ticket_folders"][key] = folder_name
            moved += 1
        save_storage(self.templates, self.meta)
        self._populate_listview()
        self._refresh_folder_combo()
        if moved:
            self.list_folder_var.set(folder_name)
            self._filter_listview()
            self.show_list_view()
            msg = f"Moved {moved} ticket(s) to '{folder_name}'." if folder_name != "Unfiled" else f"Removed {moved} ticket(s) from folder."
            messagebox.showinfo("Moved", msg)

    def _create_folder_and_move_selected(self):
        """Prompt for folder name, create it, and move selected tickets there."""
        sel = self.list_tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Select rows to move.")
            return
        iids = [s for s in sel if s and not str(s).startswith("folder_")]
        if not iids:
            return
        name = simpledialog.askstring("Create Folder", "Enter folder name:")
        if not name or not name.strip():
            return
        name = name.strip()
        if name in self.meta.get("folders", []):
            messagebox.showinfo("Exists", f"Folder '{name}' already exists.")
            return
        self.meta.setdefault("folders", []).append(name)
        save_storage(self.templates, self.meta)
        self._refresh_folder_combo()
        self._move_selected_to_folder_by_name(name)

    # ---------------- Folder management ----------------
    def manage_folders(self):
        """Create, rename, or delete folders."""
        win = tk.Toplevel(self)
        self._register_toplevel(win)
        win.title("Manage Folders")
        win.minsize(380, 320)
        win.geometry("420x380")
        win.resizable(True, True)
        ttk.Label(win, text="Custom folders for organizing fetched tickets:").pack(anchor="w", padx=8, pady=(8, 4))
        lb_frame = ttk.Frame(win)
        lb_frame.pack(fill="both", expand=True, padx=8, pady=6)
        lb = tk.Listbox(lb_frame, height=12)
        lb.pack(side="left", fill="both", expand=True)
        vs = ttk.Scrollbar(lb_frame, orient="vertical", command=lb.yview)
        vs.pack(side="right", fill="y")
        lb.configure(yscrollcommand=vs.set)
        for f in sorted(self.meta.get("folders", [])):
            lb.insert(tk.END, f)
        def refresh_list():
            lb.delete(0, tk.END)
            for f in sorted(self.meta.get("folders", [])):
                lb.insert(tk.END, f)
        def add_folder():
            name = simpledialog.askstring("New Folder", "Enter folder name:")
            if not name or not name.strip():
                return
            name = name.strip()
            if name in self.meta.get("folders", []):
                messagebox.showinfo("Exists", f"Folder '{name}' already exists.")
                return
            self.meta.setdefault("folders", []).append(name)
            save_storage(self.templates, self.meta)
            refresh_list()
            self._refresh_folder_combo()
            messagebox.showinfo("Created", f"Created folder '{name}'.")
        def rename_folder():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Info", "Select a folder to rename.")
                return
            old = lb.get(sel[0])
            new = simpledialog.askstring("Rename Folder", f"New name for '{old}':", initialvalue=old)
            if not new or not new.strip():
                return
            new = new.strip()
            if new == old:
                return
            if new in self.meta.get("folders", []):
                messagebox.showerror("Exists", f"Folder '{new}' already exists.")
                return
            folders = self.meta.get("folders", [])
            if old in folders:
                folders[folders.index(old)] = new
            tf = self.meta.get("ticket_folders", {})
            for k, v in list(tf.items()):
                if v == old:
                    tf[k] = new
            save_storage(self.templates, self.meta)
            refresh_list()
            self._refresh_folder_combo()
            self._populate_listview()
            messagebox.showinfo("Renamed", f"Renamed '{old}' to '{new}'.")
        def delete_folder():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Info", "Select a folder to delete.")
                return
            name = lb.get(sel[0])
            count = sum(1 for v in self.meta.get("ticket_folders", {}).values() if v == name)
            if count and not messagebox.askyesno("Confirm", f"Delete folder '{name}'? {count} ticket(s) will become unfiled."):
                return
            self.meta.setdefault("folders", []).remove(name)
            tf = self.meta.get("ticket_folders", {})
            for k in list(tf.keys()):
                if tf[k] == name:
                    del tf[k]
            save_storage(self.templates, self.meta)
            refresh_list()
            self._refresh_folder_combo()
            self._populate_listview()
            win.destroy()
            messagebox.showinfo("Deleted", f"Deleted folder '{name}'.")
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill="x", padx=8, pady=6)
        ttk.Button(btn_frame, text="New Folder", command=add_folder).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Rename", command=rename_folder).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Delete", command=delete_folder).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side="right")
