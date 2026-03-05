"""
Tab and template management mixin for AvalancheApp.
"""

import os
import re
import json
import csv
import copy
import traceback
import threading
import tkinter as tk
from tkinter import ttk, simpledialog, filedialog, messagebox

from config import HEADERS, FETCH_FIELDS, FETCHABLE_OPTION_FIELDS, MULTISELECT_FIELDS
from storage import save_storage
from utils import debug_log, _dedup_list_items
from tab_form import TabForm


class TabManagementMixin:
    """Mixin containing all tab and template management methods from AvalancheApp."""

    # ---------------- Template methods ----------------
    def refresh_templates(self):
        self.template_list.delete(0, tk.END)
        for name in sorted(self.templates.keys()):
            self.template_list.insert(tk.END, name)

    def on_template_select(self):
        sel = self.template_list.curselection()
        if not sel:
            return
        name = self.template_list.get(sel[0])
        tpl = self.templates.get(name, {})
        # Recover ADF if template was saved with empty ADF (pre-fix)
        adf = tpl.get("Description ADF")
        adf_empty = not adf or (isinstance(adf, dict) and not adf.get("content")) or (isinstance(adf, str) and not adf.strip())
        if adf_empty:
            recovered = self._recover_template_adf(name)
            if recovered:
                tpl["Description ADF"] = recovered
                self.templates[name] = tpl
                save_storage(self.templates, self.meta)
        self.show_tabs_view()
        # If this template is already opened in a tab, just focus it
        tab_frame = self._template_to_tab.get(name)
        if tab_frame and tab_frame in self.tabs:
            self.notebook.select(tab_frame)
            self.focus_set()
            self.lift()
            return
        # Clean stale entry if tab was closed
        if tab_frame:
            self._template_to_tab.pop(name, None)
        # Open new tab with template and associate it
        tf = self.new_tab(initial_data=tpl, select_tab=True)
        if tf and tf.frame:
            self._template_to_tab[name] = getattr(tf, "_tab_container", tf.frame)
            self.focus_set()
            self.lift()

    def _strip_identity_fields(self, data):
        """Remove Issue key/id so templates are reusable and don't accidentally update existing tickets."""
        for field in ("Issue key", "Issue id"):
            data.pop(field, None)
        return data

    def save_template_with_prompt(self, on_close=None):
        tf = self.get_active_tabform()
        if not tf:
            messagebox.showinfo("Info", "No active tab.")
            if on_close:
                self.after(50, on_close)
            return
        data = tf.read_to_dict()
        self._strip_identity_fields(data)
        name = simpledialog.askstring("Save Template", "Enter template name:", initialvalue="")
        if not name:
            if on_close:
                self.after(50, on_close)
            return
        name = name.strip()
        if name in self.templates:
            if messagebox.askyesno("Overwrite?", f"Template '{name}' already exists. Overwrite?"):
                self.templates[name] = data
                save_storage(self.templates, self.meta)
                self.refresh_templates()
                messagebox.showinfo("Saved", f"Overwrote template '{name}'.")
            else:
                new_name = simpledialog.askstring("New Name", "Enter a new template name:", initialvalue=name + " (copy)")
                if not new_name:
                    if on_close:
                        self.after(50, on_close)
                    return
                new_name = new_name.strip()
                if new_name in self.templates:
                    messagebox.showerror("Error", f"Template '{new_name}' already exists. Aborting.")
                    if on_close:
                        self.after(50, on_close)
                    return
                self.templates[new_name] = data
                save_storage(self.templates, self.meta)
                self.refresh_templates()
                messagebox.showinfo("Saved", f"Saved as new template '{new_name}'.")
        else:
            self.templates[name] = data
            save_storage(self.templates, self.meta)
            self.refresh_templates()
            messagebox.showinfo("Saved", f"Saved template '{name}'.")
        if on_close:
            self.after(50, on_close)

    def save_all(self):
        self.meta["user_cache"] = self._user_cache
        self.meta["fetched_issues"] = list(self.list_items)
        save_storage(self.templates, self.meta)
        messagebox.showinfo("Saved", "All templates, options, fetched tickets and caches saved.")

    def new_template(self):
        name = simpledialog.askstring("New Template", "Enter template name:")
        if not name:
            return
        name = name.strip()
        if name in self.templates:
            messagebox.showerror("Error", "Template with that name already exists.")
            return
        self.templates[name] = {}
        save_storage(self.templates, self.meta)
        self.refresh_templates()

    def duplicate_template(self):
        sel = self.template_list.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select a template to duplicate.")
            return
        name = self.template_list.get(sel[0])
        new_name = simpledialog.askstring("Duplicate Template", "Enter new template name:", initialvalue=name + " (copy)")
        if not new_name:
            return
        new_name = new_name.strip()
        if new_name in self.templates:
            messagebox.showerror("Error", "Template with that name already exists.")
            return
        self.templates[new_name] = dict(self.templates[name])
        save_storage(self.templates, self.meta)
        self.refresh_templates()

    def delete_template(self):
        sel = self.template_list.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select a template to delete.")
            return
        name = self.template_list.get(sel[0])
        if messagebox.askyesno("Delete", f"Delete template '{name}'?"):
            del self.templates[name]
            save_storage(self.templates, self.meta)
            self.refresh_templates()

    def remove_field_from_template(self):
        sel = self.template_list.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select a template first.")
            return
        name = self.template_list.get(sel[0])
        tpl = self.templates.get(name, {})
        keys = sorted(k for k in tpl.keys() if k in HEADERS)
        if not keys:
            messagebox.showinfo("Info", "No saved fields to remove.")
            return
        win = tk.Toplevel(self)
        self._register_toplevel(win)
        win.title("Remove Field From Template")
        win.minsize(380, 260)
        win.geometry("420x300")
        win.resizable(True, True)
        ttk.Label(win, text=f"Template: {name}").pack(anchor="w", padx=8, pady=6)
        lb = tk.Listbox(win)
        lb.pack(fill="both", expand=True, padx=8, pady=6)
        for k in keys:
            lb.insert(tk.END, k)
        def do_remove():
            s = lb.curselection()
            if not s:
                messagebox.showinfo("Info", "Select a field.")
                return
            fld = lb.get(s[0])
            if messagebox.askyesno("Confirm", f"Remove '{fld}' from template '{name}'?"):
                tpl.pop(fld, None)
                self.templates[name] = tpl
                save_storage(self.templates, self.meta)
                win.destroy()
        ttk.Button(win, text="Remove", command=do_remove).pack(side="left", padx=8, pady=6)
        ttk.Button(win, text="Cancel", command=win.destroy).pack(side="right", padx=8, pady=6)

    # ---------------- Tabs ----------------
    def _find_tab_by_ticket_key(self, key):
        """Return (tab_frame, tabform) if a tab exists for this ticket key, else None."""
        if not key:
            return None
        key = str(key).strip()
        welcome = getattr(self, "_welcome_frame", None)
        for tab_id in self.notebook.tabs():
            try:
                w = self.nametowidget(tab_id)
                if w is welcome:
                    continue
                tf = self.tabs.get(w)
                if tf:
                    identifiers = set()
                    tkey = getattr(tf, "_last_ticket_key", None)
                    if tkey:
                        identifiers.add(str(tkey).strip())
                    try:
                        d = tf.read_to_dict()
                        for k in ("Issue key", "Issue id"):
                            v = d.get(k)
                            if v:
                                identifiers.add(str(v).strip())
                    except Exception:
                        pass
                    if key in identifiers:
                        return (w, tf)
            except Exception:
                pass
        return None

    def new_tab(self, initial_data=None, select_tab=True):
        key = (initial_data or {}).get("Issue key") or (initial_data or {}).get("Issue id")
        needs_refresh = (
            key
            and not str(key).strip().startswith("LOCAL-")
            and key not in self._session_refreshed_keys
        )
        if key and not str(key).strip().startswith("LOCAL-"):
            existing = self._find_tab_by_ticket_key(key)
            if existing:
                tab_frame, tf = existing
                self._enrich_with_internal_priority(initial_data)
                tf.populate_from_dict(initial_data)
                if select_tab:
                    self.notebook.select(tab_frame)
                for h in HEADERS:
                    tf.set_option_values(h, self.meta["options"].get(h, []))
                self._schedule_on_open_reminder(key)
                if needs_refresh:
                    self.after(50, lambda k=key, d=dict(initial_data or {}), t=tf:
                               self._auto_refresh_from_jira(k, d, tabform=t))
                # Auto-fetch children if this is an Epic
                if str((initial_data or {}).get("Issue Type") or "").strip().lower() == "epic":
                    _em = str((initial_data or {}).get("_epic_mode") or "nextgen")
                    self.after(300, lambda k=key, m=_em, t=tf:
                               self._auto_fetch_epic_children(k, m, t))
                return tf
        container = ttk.Frame(self.notebook)
        tf = TabForm(
            container,
            self.meta["options"],
            field_menu_cb=self._field_context_menu,
            add_option_cb=self.add_option_from_tab_widget,
            convert_text_to_adf_cb=self._text_to_adf,
            extract_text_from_adf_cb=self._extract_text_from_adf,
            internal_priority_set_cb=self._set_internal_priority,
            open_attachments_cb=self._open_attachments_dialog,
            open_in_jira_cb=self._open_in_jira_browser,
            fetch_assignees_cb=self._on_fetch_assignees,
            fetch_options_cb=self._on_fetch_field_options,
            resolve_variables_in_adf_cb=self._resolve_vars_for_preview,
            collect_vars_cb=self._collect_defined_var_keys,
            notify_vars_changed_cb=self._refresh_var_previews,
            open_ticket_in_app_cb=self._open_ticket_link_in_app,
            open_ticket_in_jira_cb=self._open_ticket_link_in_jira
        )
        opts_cfg = self.meta.get("internal_priority_options", {})
        levels = self.meta.get("internal_priority_levels", ["High", "Medium", "Low", "None"])
        flat_opts = []
        for lvl in levels:
            flat_opts.extend(opts_cfg.get(lvl, [lvl]) if isinstance(opts_cfg.get(lvl), list) else [lvl])
        tf._internal_priority_options = flat_opts if flat_opts else ["High", "Medium", "Low", "None"]
        summary = (initial_data or {}).get("Summary", "") or ""
        short = " ".join(summary.split()[:2]) if summary.strip() else f"Ticket {len(self.tabs) + 1}"
        if len(short) > 25:
            short = short[:22] + "..."
        tf.frame.pack(side="top", fill="both", expand=True)
        tf.close_tab_cb = lambda: self.close_tab_by_frame(container)
        tab_frame = container
        self._tab_summaries[tab_frame] = summary
        self.notebook.add(tab_frame, text=short)
        self.tabs[tab_frame] = tf
        tf._tab_container = container
        if initial_data:
            self._enrich_with_internal_priority(initial_data)
            tf.populate_from_dict(initial_data)
        if select_tab:
            self.notebook.select(tab_frame)
        for h in HEADERS:
            tf.set_option_values(h, self.meta["options"].get(h, []))
        self._rename_tabs()
        self._apply_text_widget_colors_for_tabform(tf)
        self._setup_var_preview_for_tabform(tf)
        tf._pre_read_hook = lambda t: self._revert_all_var_previews(t)
        self._setup_tab_tooltip()
        if key and not str(key).strip().startswith("LOCAL-"):
            self._schedule_on_open_reminder(key)
        # Fire background refresh after tab is fully registered
        if needs_refresh:
            self.after(50, lambda k=key, d=dict(initial_data or {}), t=tf:
                       self._auto_refresh_from_jira(k, d, tabform=t))
        # Auto-fetch children if this is an Epic (deferred so tab fully renders first)
        if initial_data and str(initial_data.get("Issue Type") or "").strip().lower() == "epic":
            _ek = str(initial_data.get("Issue key") or "").strip()
            _em = str(initial_data.get("_epic_mode") or "nextgen")
            if _ek and not _ek.startswith("LOCAL-"):
                self.after(300, lambda k=_ek, m=_em, t=tf:
                           self._auto_fetch_epic_children(k, m, t))
        return tf

    def _auto_fetch_epic_children(self, epic_key: str, epic_mode: str, tf):
        """Background-fetch any child tickets of an epic that aren't in list_items.

        Called when an Epic tab is opened.  Tries both classic and next-gen JQL
        queries so children are found regardless of the stored epic mode.
        Downloads any that are missing locally, adds them to list_items, saves
        storage, and refreshes the tab's Child Issues view — all without blocking
        the UI.
        """
        # Show a loading indicator immediately so the user knows children
        # are being fetched rather than the epic being empty.
        try:
            tf._epic_children_loading = True
            tf._refresh_epic_view()
        except Exception:
            pass

        def worker():
            try:
                s = self.get_jira_session()
                if not s:
                    return

                _SEARCH_FIELDS = ["summary", "status", "issuetype"]

                # Try both JQL forms — classic and next-gen — and merge results.
                # One will typically return results and the other will be empty
                # (or error out silently), so this works for any project type.
                all_children_map = {}   # key -> issue dict (dedup by key)

                jqls = [
                    f'parent = "{epic_key}" ORDER BY created ASC',
                ]
                # "Epic Link" CF may not exist on next-gen; try it but tolerate failure
                jqls.append(f'"Epic Link" = "{epic_key}" ORDER BY created ASC')

                for jql in jqls:
                    try:
                        sr = self.jira_search_jql_simple(
                            s, jql=jql, max_results=200,
                            fields=_SEARCH_FIELDS)
                        for iss in sr.get("issues", []):
                            k = iss.get("key")
                            if k and k != epic_key and k not in all_children_map:
                                all_children_map[k] = iss
                    except Exception:
                        debug_log(f"Epic child JQL '{jql}' failed: "
                                  + traceback.format_exc())

                if not all_children_map:
                    def _clear_loading():
                        try:
                            tf._epic_children_loading = False
                            tf._refresh_epic_view()
                        except Exception:
                            pass
                    self.after(0, _clear_loading)
                    return

                all_children = list(all_children_map.values())

                # Build the compact summary list for the tab display
                children_summary = []
                for iss in all_children:
                    f = iss.get("fields") or {}
                    children_summary.append({
                        "key":     iss.get("key", ""),
                        "summary": f.get("summary", ""),
                        "status":  (f.get("status") or {}).get("name", ""),
                    })

                # Determine which children we don't have stored locally
                existing_keys = {
                    str(it.get("Issue key") or "").strip()
                    for it in self.list_items if it.get("Issue key")
                }
                missing = [
                    iss for iss in all_children
                    if iss.get("key") and iss["key"] not in existing_keys
                ]

                new_items = []
                for iss in missing:
                    child_key = iss.get("key")
                    if not child_key:
                        continue
                    try:
                        issue_json = self.fetch_issue_details(
                            s, child_key, fields=FETCH_FIELDS)
                        issue_dict = self._map_issue_json_to_dict(issue_json)
                        new_items.append(issue_dict)
                        debug_log(f"Auto-fetched epic child {child_key} for {epic_key}")
                    except Exception:
                        debug_log(f"Failed to fetch epic child {child_key}: "
                                  + traceback.format_exc())

                def _on_done():
                    if new_items:
                        self.list_items.extend(new_items)
                        self.list_items = _dedup_list_items(self.list_items)
                        self.meta["fetched_issues"] = list(self.list_items)
                        try:
                            save_storage(self.templates, self.meta)
                        except Exception:
                            pass
                        try:
                            self._rebuild_list_view()
                        except Exception:
                            pass
                        debug_log(
                            f"Added {len(new_items)} child ticket(s) for epic {epic_key}")

                    # Always refresh the tab's Child Issues display
                    try:
                        tf._epic_children_loading = False
                        tf._epic_children_data = children_summary
                        tf._refresh_epic_view()
                    except Exception:
                        pass

                self.after(0, _on_done)

            except Exception:
                debug_log(f"_auto_fetch_epic_children failed for {epic_key}: "
                          + traceback.format_exc())
                def _clear_on_error():
                    try:
                        tf._epic_children_loading = False
                        tf._refresh_epic_view()
                    except Exception:
                        pass
                self.after(0, _clear_on_error)

        threading.Thread(target=worker, daemon=True).start()

    def _apply_text_widget_colors_for_tabform(self, tabform):
        for hdr in HEADERS:
            info = tabform.field_widgets.get(hdr)
            if not info:
                continue
            w = info["widget"]
            if isinstance(w, tk.Text):
                try:
                    w.configure(bg="#1e1e1e", fg="#dcdcdc", insertbackground="#dcdcdc")
                except Exception:
                    pass

    def _setup_var_preview_for_tabform(self, tabform):
        """Wire up live variable preview on plain-text widgets in a tab.
        NOT applied to Description ADF — previewing {KEY}→value inside raw JSON
        corrupts the JSON structure when _update_preview_from_adf reads it back."""
        for hdr in ("Summary",):
            info = tabform.field_widgets.get(hdr)
            if not info:
                continue
            w = info["widget"]
            if isinstance(w, tk.Text):
                self._setup_var_preview(w)

    def _resolve_vars_for_preview(self, adf_dict):
        """Return a copy of the ADF with variable definitions stripped and references resolved."""
        vars_dict = self._collect_defined_var_keys()
        if not vars_dict:
            return adf_dict
        return self._apply_variables_to_adf(adf_dict, vars_dict)

    def close_current_tab(self):
        sel = self.notebook.select()
        if not sel:
            return
        self.close_tab_by_frame(self.nametowidget(sel))

    def close_tab_by_frame(self, tab_widget):
        """Close a specific tab. Prompts to save if unsaved."""
        if tab_widget is getattr(self, "_welcome_frame", None):
            messagebox.showinfo("Info", "The Welcome tab cannot be closed.")
            return
        tf = self.tabs.get(tab_widget)
        if not tf:
            return
        if self._is_tab_dirty(tf):
            r = messagebox.askyesnocancel("Unsaved Changes", "This tab has unsaved changes. Save before closing?")
            if r is None:
                return
            if r:
                self._save_tab_as(tf)
        self.notebook.forget(tab_widget)
        self.tabs.pop(tab_widget, None)
        self._tab_summaries.pop(tab_widget, None)
        for k, v in list(self._template_to_tab.items()):
            if v == tab_widget:
                del self._template_to_tab[k]
                break
        self._rename_tabs()

    def _is_tab_dirty(self, tf):
        """Check if tab has unsaved changes."""
        try:
            current = json.dumps(tf.read_to_dict(), sort_keys=True, default=str)
            return tf._last_saved_state != current
        except Exception:
            return False

    def save_current_tab(self):
        """Save the active tab with a user-chosen name."""
        tf = self.get_active_tabform()
        if not tf:
            messagebox.showinfo("Info", "No active tab.")
            return
        self._save_tab_as(tf)

    def _save_tab_as(self, tf):
        """Save tab to a template with user-chosen name. Prompts for name."""
        name = simpledialog.askstring("Save Ticket", "Name for this saved ticket (saved as template; tab name stays as-is):", parent=self)
        if not name or not name.strip():
            return
        name = name.strip()
        data = tf.read_to_dict()
        self._strip_identity_fields(data)
        self.templates[name] = data
        save_storage(self.templates, self.meta)
        try:
            tf._last_saved_state = json.dumps(data, sort_keys=True, default=str)
        except Exception:
            pass
        self.refresh_templates()
        messagebox.showinfo("Saved", f"Saved as template '{name}'.")

    def duplicate_current_tab(self):
        tf = self.get_active_tabform()
        if not tf:
            return
        data = tf.read_to_dict()
        self.new_tab(initial_data=data)

    def _rename_tabs(self):
        welcome = getattr(self, "_welcome_frame", None)
        for tab_id in self.notebook.tabs():
            try:
                w = self.nametowidget(tab_id)
                if w is welcome:
                    self.notebook.tab(tab_id, text="Welcome")
                else:
                    summary = self._tab_summaries.get(w, "")
                    short = " ".join(summary.split()[:2]) if summary.strip() else "Ticket"
                    if len(short) > 25:
                        short = short[:22] + "..."
                    self.notebook.tab(tab_id, text=short)
            except Exception:
                pass

    def _setup_tab_tooltip(self):
        """Bind motion to show full summary when hovering over a tab."""
        self._tab_tooltip_job = None
        def on_motion(event):
            try:
                ident = self.notebook.identify(event.x, event.y)
                if ident not in ("tab", "label"):
                    self._hide_tab_tooltip()
                    return
                sel = self.notebook.select()
                if sel:
                    w = self.nametowidget(sel)
                    if w is not getattr(self, "_welcome_frame", None):
                        summary = self._tab_summaries.get(w, "")
                        if summary:
                            self._schedule_tab_tooltip(event, summary)
                            return
                self._hide_tab_tooltip()
            except Exception:
                self._hide_tab_tooltip()
        def on_leave(event):
            self._hide_tab_tooltip()
        if not hasattr(self, "_tab_tooltip_bound") or not self._tab_tooltip_bound:
            self.notebook.bind("<Motion>", on_motion, add="+")
            self.notebook.bind("<Leave>", on_leave, add="+")
            self._tab_tooltip_bound = True

    def _schedule_tab_tooltip(self, event, text):
        if self._tab_tooltip_job:
            self.after_cancel(self._tab_tooltip_job)
        def show():
            self._tab_tooltip_job = None
            self._show_tab_tooltip(event, text)
        self._tab_tooltip_job = self.after(600, show)

    def _show_tab_tooltip(self, event, text):
        if not text:
            return
        if self._tab_tooltip_win and self._tab_tooltip_win.winfo_exists():
            try:
                self._tab_tooltip_win.destroy()
            except Exception:
                pass
        self._tab_tooltip_win = tk.Toplevel(self)
        self._tab_tooltip_win.wm_overrideredirect(True)
        self._tab_tooltip_win.wm_geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
        lbl = tk.Label(self._tab_tooltip_win, text=text[:200] + ("..." if len(text) > 200 else ""), bg="#3c3c3c", fg="#dcdcdc", font=("Segoe UI", 9), padx=8, pady=4, wraplength=400)
        lbl.pack()

    def _hide_tab_tooltip(self):
        if hasattr(self, "_tab_tooltip_job") and self._tab_tooltip_job:
            try:
                self.after_cancel(self._tab_tooltip_job)
            except Exception:
                pass
            self._tab_tooltip_job = None
        if hasattr(self, "_tab_tooltip_win") and self._tab_tooltip_win and self._tab_tooltip_win.winfo_exists():
            try:
                self._tab_tooltip_win.destroy()
            except Exception:
                pass
        self._tab_tooltip_win = None

    def get_active_tabform(self):
        sel = self.notebook.select()
        if not sel:
            return None
        tab_widget = self.nametowidget(sel)
        return self.tabs.get(tab_widget)

    def _enrich_with_internal_priority(self, data):
        key = data.get("Issue key") or data.get("Issue id")
        if key:
            data["Internal Priority"] = self.meta.get("internal_priorities", {}).get(str(key), "None")

    def _set_internal_priority(self, ticket_key, priority, refresh_list=True):
        if not ticket_key:
            return
        self.meta.setdefault("internal_priorities", {})[str(ticket_key)] = str(priority)
        for i, it in enumerate(self.list_items):
            if (it.get("Issue key") or it.get("Issue id")) == ticket_key:
                it["Internal Priority"] = priority
                break
        save_storage(self.templates, self.meta)
        if refresh_list:
            self._populate_listview()

    def on_tab_changed(self):
        tf = self.get_active_tabform()
        if not tf:
            return
        for h in HEADERS:
            tf.set_option_values(h, self.meta["options"].get(h, []))
        self.update_filter_for_active_tab()
        ticket_key = getattr(tf, "_last_ticket_key", None)
        if ticket_key:
            self._schedule_on_open_reminder(ticket_key)

    # ---------------- CSV Import/Export ----------------
    def import_from_csv_rows_open_tabs(self):
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
        win = tk.Toplevel(self)
        self._register_toplevel(win)
        win.title("Select rows to open as tabs")
        win.minsize(500, 400)
        win.geometry("1000x600")
        win.resizable(True, True)
        ttk.Label(win, text=f"CSV: {os.path.basename(path)} — select rows to open as tabs").pack(anchor="w", padx=8, pady=6)
        lb = tk.Listbox(win, selectmode="extended")
        lb.pack(fill="both", expand=True, padx=8, pady=6)
        for i, r in enumerate(rows, start=1):
            s = (r.get("Issue key") or "") + " — " + ((r.get("Summary") or "")[:200].replace("\n", " "))
            lb.insert(tk.END, f"{i}. {s}")
        def do_open():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Info", "No rows selected.")
                return
            for idx in sel:
                row = rows[idx]
                data = {h: row.get(h, "") for h in HEADERS}
                if "Description ADF" in row and row.get("Description ADF"):
                    try:
                        data["Description ADF"] = json.loads(row.get("Description ADF"))
                    except Exception:
                        pass
                self.new_tab(initial_data=data)
            win.destroy()
        ttk.Button(win, text="Open Selected as Tabs", command=do_open).pack(side="left", padx=8, pady=6)
        ttk.Button(win, text="Cancel", command=win.destroy).pack(side="right", padx=8, pady=6)

    def import_rows_into_current_tab(self):
        tf = self.get_active_tabform()
        if not tf:
            messagebox.showinfo("Info", "No active tab.")
            return
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
        win = tk.Toplevel(self)
        self._register_toplevel(win)
        win.title("Select row to import into active tab")
        win.minsize(500, 350)
        win.geometry("900x400")
        win.resizable(True, True)
        ttk.Label(win, text=f"CSV: {os.path.basename(path)} — select a row to import into this tab").pack(anchor="w", padx=8, pady=6)
        lb = tk.Listbox(win)
        lb.pack(fill="both", expand=True, padx=8, pady=6)
        for i, r in enumerate(rows, start=1):
            s = (r.get("Issue key") or "") + " — " + ((r.get("Summary") or "")[:200].replace("\n", " "))
            lb.insert(tk.END, f"{i}. {s}")
        def do_load():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Info", "No row selected.")
                return
            row = rows[sel[0]]
            data = {h: row.get(h, "") for h in HEADERS}
            if "Description ADF" in row and row.get("Description ADF"):
                try:
                    data["Description ADF"] = json.loads(row.get("Description ADF"))
                except Exception:
                    pass
            self._enrich_with_internal_priority(data)
            tf.populate_from_dict(data)
            win.destroy()
        ttk.Button(win, text="Load Selected into Active Tab", command=do_load).pack(side="left", padx=8, pady=6)
        ttk.Button(win, text="Cancel", command=win.destroy).pack(side="right", padx=8, pady=6)

    # ---------------- Combobox persistence ----------------
    def add_option_from_tab_widget(self, tabform, field):
        info = tabform.field_widgets.get(field)
        if not info:
            return
        widget = info["widget"]
        try:
            if isinstance(widget, tk.Text):
                val = widget.get("1.0", "end").strip()
            else:
                val = widget.get().strip()
        except Exception:
            val = ""
        if not val:
            messagebox.showinfo("Info", "No value to add.")
            return
        opts = self.meta.setdefault("options", {}).setdefault(field, [])
        if val in opts:
            messagebox.showinfo("Info", "Value already stored.")
            return
        opts.insert(0, val)
        self.meta["options"][field] = opts[:500]
        save_storage(self.templates, self.meta)
        for tf in self.tabs.values():
            tf.set_option_values(field, opts)
        messagebox.showinfo("Saved", f"Added '{val}' to saved options for {field}.")

    def _on_fetch_assignees(self, tabform):
        """Fetch assignable users from Jira and populate the Assignee dropdown."""
        s = self.get_jira_session()
        if not s:
            messagebox.showinfo("Info", "Set Jira API credentials first to load assignees.")
            return
        data = tabform.read_to_dict()
        project_key = (data.get("Project key") or "").strip()
        names = self._fetch_assignable_users(s, project_key)
        if not names:
            msg = "No assignable users found."
            if project_key:
                msg += f" (Project: {project_key})"
            messagebox.showinfo("Assignees", msg)
            return
        self.meta.setdefault("options", {})["Assignee"] = names
        save_storage(self.templates, self.meta)
        tabform.set_option_values("Assignee", names)
        messagebox.showinfo("Assignees", f"Loaded {len(names)} assignable user(s) into the dropdown.")

    def _on_fetch_field_options(self, tabform, field_name):
        """Fetch options for Project key, Issue Type, Status, Priority, Components, Labels, or Reporter from Jira."""
        s = self.get_jira_session()
        if not s:
            messagebox.showinfo("Info", "Set Jira API credentials first to load options.")
            return
        data = tabform.read_to_dict()
        project_key = (data.get("Project key") or "").strip()
        values = []
        try:
            if field_name == "Project key":
                values = self._fetch_projects(s)
            elif field_name == "Issue Type":
                values = self._fetch_issue_types(s, project_key)
            elif field_name == "Status":
                values = self._fetch_statuses(s, project_key)
            elif field_name == "Priority":
                values = self._fetch_priorities(s)
            elif field_name == "Components":
                values = self._fetch_components(s, project_key)
            elif field_name == "Labels":
                values = self._fetch_labels(s)
            elif field_name == "Reporter":
                values = self._fetch_assignable_users(s, project_key)
            else:
                messagebox.showinfo("Info", f"No fetch implemented for {field_name}.")
                return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch {field_name}: {e}")
            return
        if not values:
            msg = f"No {field_name} values found."
            if project_key and field_name in ("Issue Type", "Status", "Components"):
                msg += f" (Project: {project_key})"
            messagebox.showinfo(field_name, msg)
            return
        # Labels are already merged inside _fetch_labels; for other fields just replace
        self.meta.setdefault("options", {})[field_name] = values
        save_storage(self.templates, self.meta)
        tabform.set_option_values(field_name, values)
        messagebox.showinfo(field_name, f"Loaded {len(values)} option(s) into the dropdown.")

    def _field_context_menu(self, event, field, tabform):
        menu = tk.Menu(self, tearoff=0)
        w = event.widget

        # For tkinterweb widgets, the event.widget may be a child — walk up to the html widget
        html_w = None
        if hasattr(w, "selection_manager") or hasattr(w, "get_selection"):
            html_w = w
        elif hasattr(getattr(w, "master", None), "selection_manager") or hasattr(getattr(w, "master", None), "get_selection"):
            html_w = w.master

        # Check for a text selection
        has_sel = False
        try:
            if html_w:
                sm = getattr(html_w, "selection_manager", None)
                sel = sm.get_selection() if sm else html_w.get_selection()
                has_sel = bool(sel and sel.strip())
            elif isinstance(w, tk.Text):
                has_sel = bool(w.get("sel.first", "sel.last").strip())
            elif isinstance(w, (tk.Entry, ttk.Combobox)):
                has_sel = w.selection_present()
        except (tk.TclError, Exception):
            pass

        capture_w = html_w or w
        insert_to_adf_target = not isinstance(capture_w, (tk.Text, tk.Entry, ttk.Combobox))
        if has_sel:
            menu.add_command(
                label="📌 Define Variable (from selection)",
                command=lambda: (self._snapshot_var_selection_from(capture_w), self.define_variable_dialog()),
            )
        else:
            menu.add_command(label="📌 Define Variable (select text first)", state="disabled")

        # Position the caret at the right-click coordinates so that
        # get_caret_position() reflects WHERE the user right-clicked,
        # not wherever they last left-clicked.  Uses the same internal
        # tkhtml node(True, x, y) → caret_manager.set() path that
        # tkinterweb's own _on_click uses.
        _caret_text = None
        if insert_to_adf_target:
            try:
                hw = tabform._adf_html_widget
                inner = getattr(hw, '_html', None)
                if inner:
                    # event.widget is the inner canvas; x/y are relative to it
                    ex, ey = event.x, event.y
                    # If the event came through the outer wrapper, convert coords
                    if event.widget is not inner:
                        ex = event.x_root - inner.winfo_rootx()
                        ey = event.y_root - inner.winfo_rooty()
                    node_h, offset = inner.node(True, ex, ey)
                    if node_h:
                        inner.caret_manager.set(node_h, offset)
            except Exception:
                pass
            try:
                hw = tabform._adf_html_widget
                if hw and hasattr(hw, 'get_caret_position'):
                    cp = hw.get_caret_position()
                    if cp:
                        _, _caret_text, _ = cp
            except Exception:
                pass

        # List existing variables for insertion
        defined = self._collect_defined_var_keys()
        if defined:
            var_menu = tk.Menu(menu, tearoff=0)
            insert_to_adf = insert_to_adf_target
            for key, val in sorted(defined.items()):
                display_val = val[:40] + "…" if len(val) > 40 else val
                if insert_to_adf:
                    if has_sel:
                        var_menu.add_command(
                            label=f"{{{key}}}  =  {display_val}",
                            command=lambda k=key, ww=capture_w, tf=tabform: self._insert_var_reference_from_html_selection(k, ww, tf),
                        )
                    else:
                        var_menu.add_command(
                            label=f"{{{key}}}  =  {display_val}",
                            command=lambda k=key, tf=tabform, ct=_caret_text: self._insert_var_ref_into_adf(tf, k, caret_text=ct),
                        )
                else:
                    var_menu.add_command(
                        label=f"{{{key}}}  =  {display_val}",
                        command=lambda k=key, ww=capture_w: self._insert_var_reference(k, ww),
                    )
            menu.add_cascade(label="📎 Insert Variable Reference", menu=var_menu)

        menu.add_separator()
        if field in ("Summary", "Description", "Description ADF"):
            menu.add_command(label="(No persistent options for multi-line fields)", state="disabled")
        else:
            menu.add_command(label="Add current value to saved options", command=lambda f=field, tf=tabform: self.add_option_from_tab_widget(tf, f))
            menu.add_command(label="Clear saved options for this field", command=lambda f=field: self.clear_options_for_field(f))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def clear_options_for_field(self, field):
        if messagebox.askyesno("Confirm", f"Clear saved options for {field}?"):
            self.meta.setdefault("options", {})[field] = []
            save_storage(self.templates, self.meta)
            for tf in self.tabs.values():
                tf.set_option_values(field, [])
            messagebox.showinfo("Cleared", f"Cleared saved options for {field}.")

    # ---------------- Export active tab ----------------
    # ---------------- Filter / include ----------------
    def update_filter_for_active_tab(self):
        q = self.filter_var.get().strip().lower()
        tf = self.get_active_tabform()
        if not tf:
            return
        tf.collapse_unincluded(collapse_on=tf.collapse_mode, filter_q=q)

    def toggle_collapse_current_tab(self):
        tf = self.get_active_tabform()
        if not tf:
            return
        tf.collapse_mode = not tf.collapse_mode
        self.collapse_btn.config(text="Expand All Fields" if tf.collapse_mode else "Collapse Unincluded Fields")
        tf.collapse_unincluded(collapse_on=tf.collapse_mode, filter_q=self.filter_var.get().strip().lower())

    def check_all_current_tab(self):
        tf = self.get_active_tabform()
        if not tf:
            return
        tf.check_all()

    def uncheck_all_current_tab(self):
        tf = self.get_active_tabform()
        if not tf:
            return
        tf.uncheck_all()
