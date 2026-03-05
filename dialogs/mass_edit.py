"""
MassEditMixin — Mass edit selected Jira issues.
"""
import tkinter as tk
from tkinter import ttk, messagebox

from storage import perform_jira_request, save_storage
from utils import debug_log


class MassEditMixin:
    """Mixin providing _mass_edit_selected."""

    _MASS_EDIT_FIELDS = [
        ("Summary",     "summary",     "text"),
        ("Priority",    "priority",    "select"),
        ("Labels",      "labels",      "text"),
        ("Components",  "components",  "text"),
        ("Assignee",    "assignee",    "text"),
        ("Status",      "status",      "transition"),
    ]

    def _mass_edit_selected(self):
        """Open mass-edit dialog for checked/selected list-view issues."""
        if not self.list_tree:
            return
        iids = self._get_selected_or_checked_iids()
        if not iids:
            messagebox.showinfo("Mass Edit", "Check or select one or more issues in the list first.")
            return
        issues = []
        for iid in iids:
            idx = self._list_iid_to_index(iid)
            if 0 <= idx < len(self.list_items):
                it = self.list_items[idx]
                key = it.get("Issue key") or it.get("Issue id") or ""
                if key and not key.startswith("LOCAL-"):
                    issues.append((idx, it))
        if not issues:
            messagebox.showinfo("Mass Edit", "No Jira issues found (local-only tickets cannot be mass-edited).")
            return

        dlg = tk.Toplevel(self)
        self._register_toplevel(dlg)
        dlg.title(f"Mass Edit — {len(issues)} issue(s)")
        dlg.minsize(520, 400)
        dlg.geometry("600x480")
        dlg.resizable(True, True)

        main = ttk.Frame(dlg, padding=12)
        main.pack(fill="both", expand=True)

        # Issue list
        ttk.Label(main, text=f"Selected issues ({len(issues)}):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        issue_frame = ttk.Frame(main)
        issue_frame.pack(fill="x", pady=(2, 8))
        issue_lb = tk.Listbox(issue_frame, height=min(6, len(issues)), bg="#1e1e1e", fg="#dcdcdc",
                              selectbackground="#3c3c3c", font=("Segoe UI", 9))
        issue_sb = ttk.Scrollbar(issue_frame, orient="vertical", command=issue_lb.yview)
        issue_lb.configure(yscrollcommand=issue_sb.set)
        issue_lb.pack(side="left", fill="x", expand=True)
        issue_sb.pack(side="right", fill="y")
        for _, it in issues:
            key = it.get("Issue key", "")
            summ = (it.get("Summary") or "")[:60]
            issue_lb.insert("end", f"{key}  —  {summ}")

        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=8)

        # Field selector
        field_frame = ttk.Frame(main)
        field_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(field_frame, text="Field to edit:").pack(side="left", padx=(0, 8))
        field_names = [f[0] for f in self._MASS_EDIT_FIELDS]
        field_var = tk.StringVar(value=field_names[0])
        field_combo = ttk.Combobox(field_frame, textvariable=field_var, values=field_names,
                                   width=20, state="readonly")
        field_combo.pack(side="left")

        # Value input area
        val_frame = ttk.LabelFrame(main, text="New value", padding=8)
        val_frame.pack(fill="both", expand=True, pady=(0, 8))

        val_var = tk.StringVar()
        val_entry = tk.Entry(val_frame, textvariable=val_var, font=("Segoe UI", 10),
                             bg="#1e1e1e", fg="#dcdcdc", insertbackground="#dcdcdc")
        val_combo = ttk.Combobox(val_frame, textvariable=val_var, font=("Segoe UI", 10), width=40)
        _active_input = [val_entry]

        # Search + multi-select listbox for Labels / Components
        multi_frame = ttk.Frame(val_frame)
        search_var = tk.StringVar()
        _SEARCH_PLACEHOLDER = "🔍  Type to filter..."
        search_entry = tk.Entry(multi_frame, textvariable=search_var, font=("Segoe UI", 10),
                                bg="#2a2a2a", fg="#777777", insertbackground="#dcdcdc")
        search_entry.pack(fill="x", pady=(0, 4))
        search_entry.insert(0, _SEARCH_PLACEHOLDER)
        def _search_focus_in(e):
            if search_entry.get() == _SEARCH_PLACEHOLDER:
                search_entry.delete(0, "end")
                search_entry.config(fg="#dcdcdc")
        def _search_focus_out(e):
            if not search_entry.get().strip():
                search_entry.config(fg="#777777")
                search_entry.delete(0, "end")
                search_entry.insert(0, _SEARCH_PLACEHOLDER)
        search_entry.bind("<FocusIn>", _search_focus_in)
        search_entry.bind("<FocusOut>", _search_focus_out)
        _all_options = []
        _selected_set = set()
        _visible_items = []
        _rebuilding = [False]

        selected_lbl = ttk.Label(multi_frame, text="0 selected", foreground="#888888")
        selected_lbl.pack(anchor="w")

        lb_inner = ttk.Frame(multi_frame)
        lb_inner.pack(fill="both", expand=True)
        multi_lb = tk.Listbox(lb_inner, selectmode="multiple", height=6, bg="#1e1e1e", fg="#dcdcdc",
                              selectbackground="#264f78", font=("Segoe UI", 9),
                              activestyle="none", cursor="hand2")
        multi_sb = ttk.Scrollbar(lb_inner, orient="vertical", command=multi_lb.yview)
        multi_lb.configure(yscrollcommand=multi_sb.set)

        def _update_selected_label():
            selected_lbl.config(text=f"{len(_selected_set)} selected")
            val_var.set("; ".join(sorted(_selected_set, key=str.lower)))

        def _rebuild_multi_list(query=""):
            _rebuilding[0] = True
            q = query.lower().strip()
            _visible_items[:] = [x for x in _all_options if q in x.lower()] if q else list(_all_options)
            multi_lb.delete(0, "end")
            for item in _visible_items:
                multi_lb.insert("end", "  " + item)
            for i, item in enumerate(_visible_items):
                if item in _selected_set:
                    multi_lb.selection_set(i)
            _rebuilding[0] = False
            _update_selected_label()

        def _on_multi_lb_select(*_):
            if _rebuilding[0]:
                return
            new_sel = {_visible_items[i] for i in range(len(_visible_items))
                       if multi_lb.selection_includes(i)}
            if not new_sel and _selected_set:
                for i, item in enumerate(_visible_items):
                    if item in _selected_set:
                        multi_lb.selection_set(i)
                return
            for i, item in enumerate(_visible_items):
                if multi_lb.selection_includes(i):
                    _selected_set.add(item)
                else:
                    _selected_set.discard(item)
            _update_selected_label()

        multi_lb.bind("<<ListboxSelect>>", _on_multi_lb_select)
        _is_multiselect_mode = [False]

        def _on_search(*_):
            raw = search_var.get()
            if raw == _SEARCH_PLACEHOLDER:
                return
            term = raw.strip()
            if _is_multiselect_mode[0]:
                _rebuild_multi_list(term)
            else:
                q = term.lower()
                multi_lb.delete(0, "end")
                for v in _all_options:
                    if not q or q in v.lower():
                        multi_lb.insert("end", v)
        search_var.trace_add("write", _on_search)
        multi_lb.pack(side="left", fill="both", expand=True)
        multi_sb.pack(side="right", fill="y")

        hint_lbl = ttk.Label(val_frame, text="", wraplength=500, foreground="#888888")
        hint_lbl.pack(anchor="w", pady=(4, 0))

        # Mode for appending labels/components
        mode_frame = ttk.Frame(val_frame)
        mode_frame.pack(fill="x", pady=(6, 0))
        mode_var = tk.StringVar(value="replace")
        ttk.Radiobutton(mode_frame, text="Replace", variable=mode_var, value="replace").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(mode_frame, text="Add to existing", variable=mode_var, value="add").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(mode_frame, text="Remove from existing", variable=mode_var, value="remove").pack(side="left")
        mode_frame.pack_forget()

        # Fetch button
        fetch_btn = ttk.Button(val_frame, text="↻ Load options from Jira")
        fetch_btn.pack_forget()

        _field_to_options_key = {
            "Priority": "Priority", "Labels": "Labels", "Components": "Components",
            "Status": "Status", "Assignee": None,
        }

        def _load_options_for_field(field_display_name):
            """Fetch Jira options for the given field and populate the input."""
            s = self.get_jira_session()
            if not s:
                return
            prog_var.set(f"Fetching {field_display_name} options from Jira...")
            dlg.update_idletasks()
            project_key = ""
            if issues:
                project_key = (issues[0][1].get("Project key") or "").strip()
            try:
                if field_display_name == "Priority":
                    vals = self._fetch_priorities(s)
                elif field_display_name == "Labels":
                    vals = self._fetch_labels(s)
                elif field_display_name == "Components":
                    vals = self._fetch_components(s, project_key)
                elif field_display_name == "Status":
                    vals = self._fetch_statuses(s, project_key)
                elif field_display_name == "Assignee":
                    vals = self._fetch_assignable_users(s, project_key)
                else:
                    vals = []
            except Exception as e:
                prog_var.set(f"Failed: {e}")
                return
            self.meta.setdefault("options", {})[field_display_name] = vals
            prog_var.set(f"Loaded {len(vals)} option(s).")
            _refresh_input_widget(field_display_name)

        def _refresh_input_widget(name):
            """Show the right input widget and populate options."""
            val_entry.pack_forget()
            val_combo.pack_forget()
            multi_frame.pack_forget()
            multi_lb.unbind("<<ListboxSelect>>")
            _selected_set.clear()
            _visible_items.clear()
            search_entry.delete(0, "end")
            search_entry.insert(0, _SEARCH_PLACEHOLDER)
            search_entry.config(fg="#777777")

            cached = self.meta.get("options", {}).get(name, [])

            if name in ("Labels", "Components"):
                _is_multiselect_mode[0] = True
                mode_frame.pack(fill="x", pady=(6, 0))
                multi_lb.config(selectmode="multiple")
                if cached:
                    _all_options.clear()
                    _all_options.extend(cached)
                    _rebuild_multi_list()
                    multi_frame.pack(fill="both", expand=True, pady=(4, 0))
                    hint_lbl.config(text="Click items to toggle selection. Use the search bar to filter.")
                    multi_lb.bind("<<ListboxSelect>>", _on_multi_lb_select)
                    _active_input[0] = multi_lb
                else:
                    val_entry.pack(fill="x")
                    _active_input[0] = val_entry
                    hint_lbl.config(text="Separate values with semicolons (;).  Click ↻ to load options from Jira.")
            elif name in ("Priority", "Status", "Assignee"):
                _is_multiselect_mode[0] = False
                mode_frame.pack_forget()
                mode_var.set("replace")
                if cached:
                    _all_options.clear()
                    _all_options.extend(cached)
                    search_var.set("")
                    multi_lb.delete(0, "end")
                    for v in cached:
                        multi_lb.insert("end", v)
                    multi_lb.config(selectmode="browse")
                    multi_frame.pack(fill="both", expand=True, pady=(4, 0))
                    _active_input[0] = val_combo
                    def _on_single_select(e):
                        sel = multi_lb.curselection()
                        if sel:
                            val_var.set(multi_lb.get(sel[0]).strip())
                    multi_lb.bind("<<ListboxSelect>>", _on_single_select)
                else:
                    val_entry.pack(fill="x")
                    _active_input[0] = val_entry
                if name == "Priority":
                    hint_lbl.config(text="Search and select a priority." if cached else "Enter priority name, or click ↻ to load from Jira.")
                elif name == "Status":
                    hint_lbl.config(text="Search and select a status." if cached else "Enter status name, or click ↻ to load from Jira.")
                elif name == "Assignee":
                    hint_lbl.config(text="Search and select an assignee." if cached else "Enter display name or email, or click ↻ to load from Jira.")
            else:
                _is_multiselect_mode[0] = False
                mode_frame.pack_forget()
                mode_var.set("replace")
                val_entry.pack(fill="x")
                _active_input[0] = val_entry
                hint_lbl.config(text="This will replace the value on all selected issues.")

            if name != "Summary":
                fetch_btn.config(command=lambda: _load_options_for_field(name))
                fetch_btn.pack(anchor="w", pady=(6, 0))
            else:
                fetch_btn.pack_forget()

        def _on_field_change(*_):
            name = field_var.get()
            val_var.set("")
            _refresh_input_widget(name)

        field_var.trace_add("write", _on_field_change)
        _on_field_change()

        # Progress
        prog_var = tk.StringVar(value="")
        prog_lbl = ttk.Label(main, textvariable=prog_var, foreground="#aaaaaa")
        prog_lbl.pack(anchor="w")

        # Buttons
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(8, 0))

        def _apply():
            field_name = field_var.get()
            spec = next((f for f in self._MASS_EDIT_FIELDS if f[0] == field_name), None)
            if not spec:
                return
            if field_name in ("Labels", "Components") and _selected_set:
                new_val = "; ".join(sorted(_selected_set, key=str.lower))
            else:
                new_val = val_var.get().strip()
            if not new_val:
                messagebox.showinfo("Mass Edit", "Select or enter a value first.")
                return
            _, jira_field, kind = spec
            mode = mode_var.get()

            session = self.get_jira_session()
            if not session:
                return

            mode_label = {"replace": "Replace with", "add": "Add", "remove": "Remove"}.get(mode, mode)
            if not messagebox.askyesno("⚠ Confirm — Push Changes to Jira",
                    f"You are about to update {len(issues)} issue(s) directly in Jira.\n\n"
                    f"  Field:   {field_name}\n"
                    f"  Action:  {mode_label}\n"
                    f"  Value:   {new_val}\n\n"
                    "This cannot be undone. Proceed?"):
                return

            apply_btn.config(state="disabled")
            successes = []
            failures = []

            for i, (li_idx, it) in enumerate(issues, 1):
                issue_key = it.get("Issue key") or ""
                prog_var.set(f"Processing {i}/{len(issues)}: {issue_key}...")
                dlg.update_idletasks()

                # 1) Refresh issue from Jira
                try:
                    fresh = self.fetch_issue_details(session, issue_key)
                    fresh_fields = fresh.get("fields") or {}
                except Exception as e:
                    failures.append((issue_key, f"Fetch failed: {e}"))
                    continue

                # 2) Build update payload
                payload_fields = {}
                try:
                    if jira_field == "summary":
                        payload_fields["summary"] = new_val
                    elif jira_field == "priority":
                        payload_fields["priority"] = {"name": new_val}
                    elif jira_field == "assignee":
                        project_key = (fresh_fields.get("project") or {}).get("key", "")
                        acct = self._resolve_assignee(session, new_val, project_key=project_key)
                        if not acct:
                            failures.append((issue_key, f"Could not resolve assignee '{new_val}'"))
                            continue
                        payload_fields["assignee"] = {"accountId": acct}
                    elif jira_field == "labels":
                        existing = fresh_fields.get("labels") or []
                        sep = ";" if ";" in new_val else ","
                        new_items = [v.strip() for v in new_val.split(sep) if v.strip()]
                        if mode == "replace":
                            payload_fields["labels"] = new_items
                        elif mode == "add":
                            payload_fields["labels"] = list(dict.fromkeys(existing + new_items))
                        elif mode == "remove":
                            remove_set = set(v.lower() for v in new_items)
                            payload_fields["labels"] = [l for l in existing if l.lower() not in remove_set]
                    elif jira_field == "components":
                        existing = [c.get("name", "") for c in (fresh_fields.get("components") or [])]
                        new_items = [v.strip() for v in new_val.split(";") if v.strip()]
                        if mode == "replace":
                            payload_fields["components"] = [{"name": n} for n in new_items]
                        elif mode == "add":
                            merged = list(dict.fromkeys(existing + new_items))
                            payload_fields["components"] = [{"name": n} for n in merged]
                        elif mode == "remove":
                            remove_set = set(v.lower() for v in new_items)
                            payload_fields["components"] = [{"name": n} for n in existing if n.lower() not in remove_set]
                    elif jira_field == "status":
                        # Status changes use transitions API
                        trans_url = f"{session._jira_base}/rest/api/3/issue/{issue_key}/transitions"
                        try:
                            tr = perform_jira_request(session, "GET", trans_url, timeout=15)
                            transitions = (tr.json() if tr.status_code == 200 else {}).get("transitions", [])
                        except Exception:
                            transitions = []
                        match = next((t for t in transitions if t.get("name", "").lower() == new_val.lower()), None)
                        if not match:
                            avail = ", ".join(t.get("name", "") for t in transitions)
                            failures.append((issue_key, f"No transition to '{new_val}'. Available: {avail}"))
                            continue
                        try:
                            resp = perform_jira_request(session, "POST", trans_url,
                                                        json_body={"transition": {"id": match["id"]}}, timeout=30)
                            if resp.status_code in (200, 204):
                                successes.append(issue_key)
                            else:
                                failures.append((issue_key, f"Transition failed: {resp.status_code}"))
                        except Exception as e:
                            failures.append((issue_key, str(e)))
                        continue
                except Exception as e:
                    failures.append((issue_key, str(e)))
                    continue

                # 3) PUT update
                if payload_fields:
                    update_url = f"{session._jira_base}/rest/api/3/issue/{issue_key}"
                    try:
                        resp = perform_jira_request(session, "PUT", update_url,
                                                    json_body={"fields": payload_fields}, timeout=30)
                        if resp.status_code in (200, 204):
                            successes.append(issue_key)
                        else:
                            err_text = (resp.text or "")[:200]
                            failures.append((issue_key, f"HTTP {resp.status_code}: {err_text}"))
                    except Exception as e:
                        failures.append((issue_key, str(e)))

            # 4) Refresh list_items from Jira for updated issues
            prog_var.set("Refreshing updated issues...")
            dlg.update_idletasks()
            for issue_key in successes:
                try:
                    fresh = self.fetch_issue_details(session, issue_key)
                    fresh_fields = fresh.get("fields") or {}
                    for li_idx, it in issues:
                        if (it.get("Issue key") or "") == issue_key:
                            it["Summary"] = fresh_fields.get("summary") or it.get("Summary", "")
                            it["Status"] = (fresh_fields.get("status") or {}).get("name", it.get("Status", ""))
                            it["Priority"] = (fresh_fields.get("priority") or {}).get("name", it.get("Priority", ""))
                            it["Assignee"] = (fresh_fields.get("assignee") or {}).get("displayName", "") or \
                                             (fresh_fields.get("assignee") or {}).get("emailAddress", "")
                            it["Labels"] = "; ".join(fresh_fields.get("labels") or [])
                            comps = fresh_fields.get("components") or []
                            it["Components"] = "; ".join(c.get("name", "") for c in comps)
                            it["Updated"] = fresh_fields.get("updated", it.get("Updated", ""))
                            break
                except Exception:
                    pass

            save_storage(self.templates, self.meta)
            self._populate_listview()

            # Show results
            msg = f"Done.  {len(successes)} succeeded,  {len(failures)} failed."
            if failures:
                detail = "\n".join(f"  {k}: {reason}" for k, reason in failures[:20])
                msg += f"\n\nFailures:\n{detail}"
            prog_var.set(msg)
            apply_btn.config(state="normal")

        apply_btn = ttk.Button(btn_frame, text="⚠ Submit Changes to Jira", command=_apply)
        apply_btn.pack(side="right", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side="right", padx=6)
