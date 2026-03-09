"""
FetchMixin — Fetch issues, refresh tab, refresh fetched tickets dialogs.
"""
import json
import copy
import traceback
import threading
import tkinter as tk
from tkinter import ttk, messagebox

from config import HEADERS, FETCH_FIELDS, DEBUG_LOG
from storage import perform_jira_request, save_storage
from utils import debug_log, _bind_mousewheel, _dedup_list_items


class FetchMixin:
    """Mixin providing fetch_my_issues_dialog, _start_fetch_issues, _on_fetch_complete, refresh_active_tab_from_jira, refresh_fetched_tickets, _on_refresh_complete."""

    def _map_epic_and_link_fields(self, fields: dict, issue_dict: dict) -> None:
        """Populate epic, parent, and issue-link fields on issue_dict in-place.

        Handles both classic Jira (customfield_10014) and next-gen (parent of
        type Epic). Also extracts the true sub-task parent and formal issue links.
        Mutates issue_dict directly; returns None.
        """
        # ── Epic relationship ───────────────────────────────────────────────
        epic_mode, epic_key = self._detect_epic_link_field(fields)
        issue_dict["Epic Link"]  = epic_key
        issue_dict["_epic_mode"] = epic_mode or ""   # remembered for upload

        # Epic Name: classic projects store it in customfield_10011;
        # next-gen projects store the epic's summary in the parent object.
        parent_obj = fields.get("parent") or {}
        if epic_mode == "nextgen" and parent_obj.get("key"):
            issue_dict["Epic Name"] = (
                (parent_obj.get("fields") or {}).get("summary", "")
                or fields.get("customfield_10011") or "")
        else:
            issue_dict["Epic Name"] = fields.get("customfield_10011") or ""

        # ── Sub-task / display parent ────────────────────────────────────────
        parent_key     = parent_obj.get("key", "")
        parent_fields  = parent_obj.get("fields") or {}
        parent_type    = parent_fields.get("issuetype", {}).get("name", "")
        parent_summary = parent_fields.get("summary", "")

        # "Parent" is the visible display field — show the key whenever any
        # kind of parent exists (epic or sub-task) so the row is never blank.
        issue_dict["Parent"] = parent_key

        if parent_key and parent_type.lower() != "epic":
            # True sub-task parent — store key + summary for upload and display
            issue_dict["Parent key"]     = parent_key
            issue_dict["Parent summary"] = parent_summary
        else:
            # Always write (not setdefault) so stale values from a previous
            # fetch are cleared when the parent relationship no longer exists.
            issue_dict["Parent key"]     = ""
            issue_dict["Parent summary"] = ""

        # ── Formal issue links ──────────────────────────────────────────────
        issue_dict["Issue Links"] = self._parse_jira_issue_links(
            fields.get("issuelinks"))

    def fetch_my_issues_dialog(self, on_close=None):
        s = self.get_jira_session()
        if not s:
            if on_close:
                self.after(50, on_close)
            return
        dlg = tk.Toplevel(self)

        def _close_and_callback():
            dlg.destroy()
            if on_close:
                self.after(50, on_close)
        self._register_toplevel(dlg)
        dlg.title("Fetch My Issues")
        dlg.minsize(560, 700)
        try:
            sw = dlg.winfo_screenwidth()
            sh = dlg.winfo_screenheight()
            w = min(700, max(620, int(sw * 0.45)))
            h = min(sh - 60, max(820, int(sh * 0.92)))
        except Exception:
            w, h = 640, 860
        dlg.geometry(f"{w}x{h}")
        dlg.resizable(True, True)
        dlg

        # Scrollable main area
        outer = tk.Frame(dlg, bg="#2b2b2b")
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, bg="#2b2b2b", highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        main = ttk.Frame(canvas, padding=16)
        cwin = canvas.create_window((0, 0), window=main, anchor="nw")
        def _on_cfg(e): canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_cv_cfg(e): canvas.itemconfig(cwin, width=max(e.width, 1))
        main.bind("<Configure>", _on_cfg)
        canvas.bind("<Configure>", _on_cv_cfg)
        _bind_mousewheel(canvas, "vertical")

        def _sep(): ttk.Separator(main, orient="horizontal").pack(fill="x", pady=(6, 10))

        # ── Shared dark picker helper ──────────────────────────────────────
        def _make_filter_row(label_text, available_items, selected_list, disp_var, disp_entry, field_name):
            """Open a dark search-picker for available_items; syncs into selected_list and disp_var."""
            if not available_items:
                messagebox.showinfo(field_name,
                    f"No {field_name} cached yet. Use ↻ on a ticket's {field_name} field first.")
                return
            _BG = "#1a1a1a"; _PANEL = "#252526"; _BORDER = "#3c3c3c"
            _FG = "#d4d4d4"; _SEL_BG = "#0e639c"; _SEARCH_BG = "#2d2d2d"

            all_items = sorted(available_items, key=str.lower)
            sel_set   = set(selected_list)
            vis       = []

            pw = tk.Toplevel(dlg)
            pw.title(f"Select {field_name}")
            pw.configure(bg=_BG)
            pw.minsize(300, 380)
            pw.geometry("340x440")
            pw.resizable(True, True)
            pw

            tb = tk.Frame(pw, bg=_PANEL, pady=10)
            tb.pack(fill="x")
            tk.Label(tb, text=f"Select {field_name} to filter by", bg=_PANEL, fg=_FG,
                     font=("Segoe UI", 11, "bold"), padx=16).pack(side="left")
            tk.Frame(pw, bg=_BORDER, height=1).pack(fill="x")

            sf = tk.Frame(pw, bg=_BG, padx=12)
            sf.pack(fill="x", pady=(10, 4))
            sv = tk.StringVar()
            se = tk.Entry(sf, textvariable=sv, bg=_SEARCH_BG, fg=_FG,
                          insertbackground=_FG, relief="flat", bd=0, font=("Segoe UI", 10),
                          highlightthickness=1, highlightbackground=_BORDER, highlightcolor="#007acc")
            se.pack(fill="x", ipady=6, padx=1)
            se.insert(0, "Search...")
            se.configure(fg="#666666")
            se.bind("<FocusIn>",  lambda e: (se.delete(0, "end"), se.configure(fg=_FG))
                                             if sv.get() == "Search..." else None)
            se.bind("<FocusOut>", lambda e: (se.insert(0, "Search..."), se.configure(fg="#666666"))
                                             if not sv.get().strip() else None)

            lo = tk.Frame(pw, bg=_BG, padx=12)
            lo.pack(fill="both", expand=True, pady=(4, 0))
            lbf = tk.Frame(lo, bg=_BORDER)
            lbf.pack(fill="both", expand=True)
            lb = tk.Listbox(lbf, selectmode="multiple", activestyle="none",
                            bg="#1e1e1e", fg=_FG, selectbackground=_SEL_BG,
                            selectforeground="#ffffff", font=("Segoe UI", 10),
                            bd=0, highlightthickness=1, highlightbackground=_BORDER,
                            highlightcolor="#007acc", relief="flat", cursor="hand2")
            sb2 = tk.Scrollbar(lbf, orient="vertical", command=lb.yview,
                               bg=_PANEL, troughcolor=_BG, width=10)
            lb.pack(side="left", fill="both", expand=True)
            sb2.pack(side="right", fill="y")
            lb.configure(yscrollcommand=sb2.set)
            _bind_mousewheel(lb, "vertical")

            cnt_v = tk.StringVar(value="0 selected")
            _rebuilding = [False]
            def _upd_cnt(): cnt_v.set(f"{len(sel_set)} selected")
            def _sync():
                # Only called from _ok() — never inside _rebuild() so sel_set
                # stays authoritative and Windows focus-loss deselects can't wipe it.
                for i, x in enumerate(vis):
                    if lb.selection_includes(i): sel_set.add(x)
                    else: sel_set.discard(x)
            def _rebuild(q=""):
                # sel_set is the single source of truth — do NOT call _sync() here.
                # On Windows the listbox fires <<ListboxSelect>> with nothing selected
                # when it loses focus; reading back from it here would wipe all picks.
                _rebuilding[0] = True
                vis[:] = [x for x in all_items if q.lower() in x.lower()] if q else list(all_items)
                lb.delete(0, "end")
                for x in vis: lb.insert("end", "  " + x)
                for i, x in enumerate(vis):
                    if x in sel_set: lb.selection_set(i)
                _rebuilding[0] = False
                _upd_cnt()
            def _on_sel(*_):
                if _rebuilding[0]:
                    return
                new_sel = {vis[i] for i in range(len(vis)) if lb.selection_includes(i)}
                # Guard: Windows fires <<ListboxSelect>> with nothing selected on
                # focus loss even though the user never deselected anything.
                if not new_sel and sel_set:
                    for i, x in enumerate(vis):
                        if x in sel_set: lb.selection_set(i)
                    return
                for i, x in enumerate(vis):
                    if lb.selection_includes(i): sel_set.add(x)
                    else: sel_set.discard(x)
                _upd_cnt()
                disp_var.set("; ".join(sorted(sel_set)) if sel_set else "None selected")
                disp_entry.configure(fg="#dcdcdc" if sel_set else "#aaaaaa")
            sv.trace_add("write", lambda *_: _rebuild(sv.get() if sv.get() != "Search..." else ""))
            lb.bind("<<ListboxSelect>>", _on_sel)
            _rebuild()

            tk.Label(pw, textvariable=cnt_v, bg=_BG, fg="#888888",
                     font=("Segoe UI", 9), anchor="w", padx=14).pack(fill="x", pady=(4, 0))
            tk.Frame(pw, bg=_BORDER, height=1).pack(fill="x", pady=(6, 0))
            bf = tk.Frame(pw, bg=_PANEL, pady=10, padx=12)
            bf.pack(fill="x")
            def _ok():
                _sync()
                selected_list[:] = sorted(sel_set)
                disp_var.set("; ".join(selected_list) if selected_list else "None selected")
                disp_entry.configure(fg="#dcdcdc" if selected_list else "#aaaaaa")
                pw.destroy()
            def _mkb(p, t, c, bg, hv):
                b = tk.Button(p, text=t, command=c, bg=bg, fg="#fff",
                              font=("Segoe UI", 9, "bold"), relief="flat", bd=0,
                              padx=14, pady=6, cursor="hand2",
                              activebackground=hv, activeforeground="#fff")
                b.bind("<Enter>", lambda e: b.configure(bg=hv))
                b.bind("<Leave>", lambda e: b.configure(bg=bg))
                return b
            def _clear_all():
                sel_set.clear()
                lb.selection_clear(0, "end")
                _upd_cnt()
                disp_var.set("None selected")
                disp_entry.configure(fg="#aaaaaa")

            _mkb(bf, "OK", _ok, "#0e639c", "#1177bb").pack(side="right", padx=(4, 0))
            _mkb(bf, "Cancel", pw.destroy, "#3c3c3c", "#505050").pack(side="right")
            _mkb(bf, "Clear", _clear_all, "#5a3030", "#7a4040").pack(side="left")
            pw.bind("<Return>", lambda e: _ok())
            pw.bind("<Escape>", lambda e: pw.destroy())
            pw.after(50, se.focus_set)

        def _filter_disp_entry(parent, disp_var, available_list, selected_list, field_name):
            """Build a clickable display entry + Choose button row. Returns the entry widget."""
            row = ttk.Frame(parent)
            row.pack(fill="x", pady=(0, 4))
            ent = tk.Entry(row, textvariable=disp_var, state="readonly",
                           readonlybackground="#3c3c3c", fg="#aaaaaa",
                           disabledforeground="#aaaaaa", font=("Segoe UI", 9),
                           relief="flat", bd=1, highlightthickness=1,
                           highlightbackground="#555555", cursor="hand2")
            ent.pack(side="left", fill="x", expand=True, padx=(0, 8))
            def _open(e=None):
                dlg.after_idle(lambda: _make_filter_row(
                    field_name, available_list, selected_list, disp_var, ent, field_name))
                return "break"
            ent.bind("<Button-1>", _open)
            ent.bind("<ButtonPress-1>", _open)
            ttk.Button(row, text="Choose...", command=lambda: _make_filter_row(
                field_name, available_list, selected_list, disp_var, ent, field_name)).pack(side="left")
            return ent

        opts = self.meta.get("options", {})

        # ── Project key ────────────────────────────────────────────────────
        last_proj = self.meta.get("last_project_key", "")
        default_proj = last_proj if last_proj else "SUNDANCE"
        ttk.Label(main, text="Project key:").pack(anchor="w", pady=(0, 4))
        proj_row = ttk.Frame(main)
        proj_row.pack(fill="x", pady=(0, 8))
        proj_var = tk.StringVar(value=default_proj)
        tk.Entry(proj_row, textvariable=proj_var, width=18,
                 bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc",
                 relief="flat", bd=1, font=("Segoe UI", 10)).pack(side="left", padx=(0, 8))
        ttk.Label(proj_row, text="(blank = all projects)").pack(side="left")
        _sep()

        # ── Assignee scope (optional) ──────────────────────────────────────
        ttk.Label(main, text="Assignee / Reporter:").pack(anchor="w", pady=(0, 4))
        scope_row = ttk.Frame(main)
        scope_row.pack(fill="x", pady=(0, 8))
        scope_var = tk.StringVar(value="assigned")
        scopes = [("Assigned to me", "assigned"), ("Created by me", "created"),
                  ("Assigned OR Created by me", "both"), ("Anyone (no restriction)", "any")]
        for label, val in scopes:
            ttk.Radiobutton(scope_row, text=label, variable=scope_var, value=val).pack(anchor="w", pady=1)
        _sep()

        # ── Label filter ───────────────────────────────────────────────────
        available_labels      = list(opts.get("Labels", []))
        selected_label_filter = []
        label_disp_var = tk.StringVar(value="None selected")
        ttk.Label(main, text="Filter by Labels (optional):").pack(anchor="w", pady=(0, 4))
        _filter_disp_entry(main, label_disp_var, available_labels, selected_label_filter, "Labels")
        lm_row = ttk.Frame(main)
        lm_row.pack(fill="x", pady=(0, 10))
        ttk.Label(lm_row, text="Match:").pack(side="left", padx=(0, 8))
        label_mode_var = tk.StringVar(value="any")
        ttk.Radiobutton(lm_row, text="At least one (OR)", variable=label_mode_var, value="any").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(lm_row, text="All selected (AND)", variable=label_mode_var, value="all").pack(side="left")
        _sep()

        # ── Component filter ───────────────────────────────────────────────
        available_comps      = list(opts.get("Components", []))
        selected_comp_filter = []
        comp_disp_var = tk.StringVar(value="None selected")
        ttk.Label(main, text="Filter by Components (optional):").pack(anchor="w", pady=(0, 4))
        _filter_disp_entry(main, comp_disp_var, available_comps, selected_comp_filter, "Components")
        cm_row = ttk.Frame(main)
        cm_row.pack(fill="x", pady=(0, 10))
        ttk.Label(cm_row, text="Match:").pack(side="left", padx=(0, 8))
        comp_mode_var = tk.StringVar(value="any")
        ttk.Radiobutton(cm_row, text="At least one (OR)", variable=comp_mode_var, value="any").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(cm_row, text="All selected (AND)", variable=comp_mode_var, value="all").pack(side="left")
        _sep()

        # ── Issue Type filter ──────────────────────────────────────────────
        available_types      = list(opts.get("Issue Type", []))
        selected_type_filter = []
        type_disp_var = tk.StringVar(value="None selected")
        ttk.Label(main, text="Filter by Issue Type (optional):").pack(anchor="w", pady=(0, 4))
        _filter_disp_entry(main, type_disp_var, available_types, selected_type_filter, "Issue Type")
        _sep()

        # ── Status filter ──────────────────────────────────────────────────
        available_statuses      = list(opts.get("Status", []))
        selected_status_filter  = []
        status_disp_var = tk.StringVar(value="None selected")
        ttk.Label(main, text="Filter by Status (optional):").pack(anchor="w", pady=(0, 4))
        _filter_disp_entry(main, status_disp_var, available_statuses, selected_status_filter, "Status")
        _sep()

        # ── Priority filter ────────────────────────────────────────────────
        available_priorities    = list(opts.get("Priority", []))
        selected_priority_filter = []
        priority_disp_var = tk.StringVar(value="None selected")
        ttk.Label(main, text="Filter by Priority (optional):").pack(anchor="w", pady=(0, 4))
        _filter_disp_entry(main, priority_disp_var, available_priorities, selected_priority_filter, "Priority")
        _sep()

        # ── Folder name ────────────────────────────────────────────────────
        ttk.Label(main, text="Save to folder (optional):").pack(anchor="w", pady=(0, 4))
        folder_row = ttk.Frame(main)
        folder_row.pack(fill="x", pady=(0, 4))
        folder_var = tk.StringVar()
        folder_entry = tk.Entry(folder_row, textvariable=folder_var, width=30,
                                bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc",
                                relief="flat", bd=1, font=("Segoe UI", 10))
        folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        def _auto_folder_name():
            parts = []
            pk = (proj_var.get() or "").strip()
            if pk: parts.append(pk)
            if selected_label_filter:    parts.append("; ".join(selected_label_filter[:2]) + ("…" if len(selected_label_filter) > 2 else ""))
            if selected_comp_filter:     parts.append("; ".join(selected_comp_filter[:2]))
            if selected_type_filter:     parts.append("; ".join(selected_type_filter[:2]))
            if selected_status_filter:   parts.append("; ".join(selected_status_filter[:2]))
            if selected_priority_filter: parts.append("; ".join(selected_priority_filter[:2]))
            folder_var.set(" | ".join(parts) if parts else "")

        ttk.Button(folder_row, text="Auto-name", command=_auto_folder_name).pack(side="left")
        ttk.Label(main, text="(leave blank to not create a folder)").pack(anchor="w", pady=(0, 10))
        _sep()

        # ── Max results ────────────────────────────────────────────────────
        mr_row = ttk.Frame(main)
        mr_row.pack(fill="x", pady=(0, 12))
        ttk.Label(mr_row, text="Max results (1–500):").pack(side="left", padx=(0, 8))
        max_var = tk.StringVar(value="50")
        ttk.Entry(mr_row, textvariable=max_var, width=8).pack(side="left")

        def on_fetch_click():
            try:
                cnt = int(max_var.get() or 50)
            except Exception:
                messagebox.showerror("Error", "Max results must be a number.")
                return
            cnt = max(1, min(500, cnt))
            pk = (proj_var.get() or "").strip()
            if pk:
                self.meta["last_project_key"] = pk
            folder_name = (folder_var.get() or "").strip()
            _close_and_callback()
            self._start_fetch_issues(
                scope=scope_var.get(),
                max_results=cnt,
                project_key=pk,
                label_filter=list(selected_label_filter),
                label_mode=label_mode_var.get(),
                component_filter=list(selected_comp_filter),
                component_mode=comp_mode_var.get(),
                type_filter=list(selected_type_filter),
                status_filter=list(selected_status_filter),
                priority_filter=list(selected_priority_filter),
                folder_name=folder_name,
            )

        btns = ttk.Frame(main)
        btns.pack(fill="x", pady=(4, 0))
        ttk.Button(btns, text="Fetch", command=on_fetch_click).pack(side="right", padx=6)
        ttk.Button(btns, text="Cancel", command=_close_and_callback).pack(side="right", padx=6)
        dlg.protocol("WM_DELETE_WINDOW", _close_and_callback)
        dlg.update_idletasks()

    def auto_fetch_settings_dialog(self):
        """Dialog to configure and save the automatic fetch-on-startup criteria."""
        cfg = dict(self.meta.get("auto_fetch_config") or {})
        opts = self.meta.get("options", {})

        _BG     = "#1e1e1e"
        _PANEL  = "#252526"
        _BORDER = "#3c3c3c"
        _FG     = "#dcdcdc"
        _ACCENT = "#0e639c"
        _ACCENT_HOV = "#1177bb"

        dlg = tk.Toplevel(self)
        self._register_toplevel(dlg)
        dlg.title("Auto-Fetch Settings")
        dlg.configure(bg=_BG)
        try:
            sw = dlg.winfo_screenwidth()
            sh = dlg.winfo_screenheight()
            w = min(680, max(560, int(sw * 0.42)))
            h = min(sh - 60, max(760, int(sh * 0.88)))
        except Exception:
            w, h = 620, 800
        dlg.geometry(f"{w}x{h}")
        dlg.resizable(True, True)
        try:
            dlg.grab_set()
        except Exception:
            pass

        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(dlg, bg=_PANEL, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Auto-Fetch Settings", bg=_PANEL, fg=_FG,
                 font=("Segoe UI", 13, "bold"), padx=20).pack(side="left")
        tk.Frame(dlg, bg=_BORDER, height=1).pack(fill="x")

        # ── Scrollable body ───────────────────────────────────────────────────
        body_canvas = tk.Canvas(dlg, bg=_BG, highlightthickness=0)
        body_sb = tk.Scrollbar(dlg, orient="vertical", command=body_canvas.yview,
                               bg=_PANEL, troughcolor=_BG)
        body_canvas.configure(yscrollcommand=body_sb.set)
        body_sb.pack(side="right", fill="y")
        body_canvas.pack(side="left", fill="both", expand=True)
        main = tk.Frame(body_canvas, bg=_BG, padx=24, pady=16)
        body_win = body_canvas.create_window((0, 0), window=main, anchor="nw")
        def _on_main_cfg(e):
            body_canvas.configure(scrollregion=body_canvas.bbox("all"))
        def _on_canvas_cfg(e):
            body_canvas.itemconfig(body_win, width=e.width)
        main.bind("<Configure>", _on_main_cfg)
        body_canvas.bind("<Configure>", _on_canvas_cfg)
        _bind_mousewheel(body_canvas, "vertical")

        def _sep():
            tk.Frame(main, bg=_BORDER, height=1).pack(fill="x", pady=(8, 8))
        def _lbl(text, **kw):
            tk.Label(main, text=text, bg=_BG, fg=_FG,
                     font=("Segoe UI", 9), **kw).pack(anchor="w", pady=(0, 4))
        def _mkb(parent, text, cmd, bg, hv):
            b = tk.Button(parent, text=text, command=cmd, bg=bg, fg="#fff",
                          font=("Segoe UI", 9, "bold"), relief="flat", bd=0,
                          padx=14, pady=6, cursor="hand2",
                          activebackground=hv, activeforeground="#fff")
            b.bind("<Enter>", lambda e: b.configure(bg=hv))
            b.bind("<Leave>", lambda e: b.configure(bg=bg))
            return b

        # ── Enable toggle ──────────────────────────────────────────────────────
        enabled_var = tk.BooleanVar(value=bool(cfg.get("enabled", False)))
        en_frame = tk.Frame(main, bg=_BG)
        en_frame.pack(fill="x", pady=(0, 10))
        en_cb = tk.Checkbutton(
            en_frame, variable=enabled_var,
            text="  Enable automatic fetch on startup",
            bg=_BG, fg=_FG, selectcolor=_PANEL,
            activebackground=_BG, activeforeground=_FG,
            font=("Segoe UI", 10, "bold"), cursor="hand2")
        en_cb.pack(side="left")
        tk.Label(en_frame,
                 text="  When enabled, Avalanche will automatically fetch matching tickets\n"
                      "  each time the app starts (once per calendar day).",
                 bg=_BG, fg="#888888", font=("Segoe UI", 8),
                 justify="left").pack(anchor="w", pady=(4, 0))
        _sep()

        # ── Inner _make_filter_row (needs dlg reference) ─────────────────────
        def _make_filter_row(label_text, available_items, selected_list, disp_var, disp_entry, field_name):
            if not available_items:
                messagebox.showinfo(field_name,
                    f"No {field_name} cached yet. Fetch some tickets first.")
                return
            all_items = sorted(available_items, key=str.lower)
            sel_set   = set(selected_list)
            vis       = []
            pw = tk.Toplevel(dlg)
            pw.title(f"Select {field_name}")
            pw.configure(bg=_BG)
            pw.minsize(300, 380); pw.geometry("340x440"); pw.resizable(True, True)
            tb = tk.Frame(pw, bg=_PANEL, pady=10); tb.pack(fill="x")
            tk.Label(tb, text=f"Select {field_name}", bg=_PANEL, fg=_FG,
                     font=("Segoe UI", 11, "bold"), padx=16).pack(side="left")
            tk.Frame(pw, bg=_BORDER, height=1).pack(fill="x")
            sf = tk.Frame(pw, bg=_BG, padx=12); sf.pack(fill="x", pady=(10, 4))
            sv2 = tk.StringVar()
            se = tk.Entry(sf, textvariable=sv2, bg="#2d2d2d", fg=_FG,
                          insertbackground=_FG, relief="flat", bd=0, font=("Segoe UI", 10),
                          highlightthickness=1, highlightbackground=_BORDER)
            se.pack(fill="x", ipady=6, padx=1)
            lo = tk.Frame(pw, bg=_BG, padx=12); lo.pack(fill="both", expand=True, pady=(4, 0))
            lbf = tk.Frame(lo, bg=_BORDER); lbf.pack(fill="both", expand=True)
            lb2 = tk.Listbox(lbf, selectmode="multiple", activestyle="none",
                             bg="#1e1e1e", fg=_FG, selectbackground=_ACCENT,
                             selectforeground="#ffffff", font=("Segoe UI", 10),
                             bd=0, highlightthickness=1, highlightbackground=_BORDER,
                             relief="flat", cursor="hand2")
            sb3 = tk.Scrollbar(lbf, orient="vertical", command=lb2.yview,
                               bg=_PANEL, troughcolor=_BG, width=10)
            lb2.pack(side="left", fill="both", expand=True); sb3.pack(side="right", fill="y")
            lb2.configure(yscrollcommand=sb3.set)
            _bind_mousewheel(lb2, "vertical")
            cnt_v2 = tk.StringVar(value="0 selected")
            _rebuilding2 = [False]
            def _upd(): cnt_v2.set(f"{len(sel_set)} selected")
            def _reb(q=""):
                _rebuilding2[0] = True
                vis[:] = [x for x in all_items if q.lower() in x.lower()] if q else list(all_items)
                lb2.delete(0, "end")
                for x in vis: lb2.insert("end", "  " + x)
                for i, x in enumerate(vis):
                    if x in sel_set: lb2.selection_set(i)
                _rebuilding2[0] = False; _upd()
            def _onsel(*_):
                if _rebuilding2[0]: return
                ns = {vis[i] for i in range(len(vis)) if lb2.selection_includes(i)}
                if not ns and sel_set:
                    for i, x in enumerate(vis):
                        if x in sel_set: lb2.selection_set(i)
                    return
                for i, x in enumerate(vis):
                    if lb2.selection_includes(i): sel_set.add(x)
                    else: sel_set.discard(x)
                _upd()
                disp_var.set("; ".join(sorted(sel_set)) if sel_set else "None selected")
                disp_entry.configure(fg=_FG if sel_set else "#aaaaaa")
            sv2.trace_add("write", lambda *_: _reb(sv2.get()))
            lb2.bind("<<ListboxSelect>>", _onsel)
            _reb()
            tk.Label(pw, textvariable=cnt_v2, bg=_BG, fg="#888888",
                     font=("Segoe UI", 9), anchor="w", padx=14).pack(fill="x", pady=(4, 0))
            tk.Frame(pw, bg=_BORDER, height=1).pack(fill="x", pady=(6, 0))
            bf2 = tk.Frame(pw, bg=_PANEL, pady=10, padx=12); bf2.pack(fill="x")
            def _ok2():
                selected_list[:] = sorted(sel_set)
                disp_var.set("; ".join(selected_list) if selected_list else "None selected")
                disp_entry.configure(fg=_FG if selected_list else "#aaaaaa")
                pw.destroy()
            def _clr():
                sel_set.clear(); lb2.selection_clear(0, "end"); _upd()
                disp_var.set("None selected"); disp_entry.configure(fg="#aaaaaa")
            _mkb(bf2, "OK", _ok2, _ACCENT, _ACCENT_HOV).pack(side="right", padx=(4, 0))
            _mkb(bf2, "Cancel", pw.destroy, "#3c3c3c", "#505050").pack(side="right")
            _mkb(bf2, "Clear", _clr, "#5a3030", "#7a4040").pack(side="left")
            pw.bind("<Return>", lambda e: _ok2()); pw.bind("<Escape>", lambda e: pw.destroy())
            pw.after(50, se.focus_set)

        def _filter_entry(disp_var, available_list, selected_list, field_name):
            row = tk.Frame(main, bg=_BG); row.pack(fill="x", pady=(0, 4))
            ent = tk.Entry(row, textvariable=disp_var, state="readonly",
                           readonlybackground="#3c3c3c", fg="#aaaaaa",
                           font=("Segoe UI", 9), relief="flat", bd=1,
                           highlightthickness=1, highlightbackground="#555555", cursor="hand2")
            ent.pack(side="left", fill="x", expand=True, padx=(0, 8))
            def _open(e=None):
                dlg.after_idle(lambda: _make_filter_row(
                    field_name, available_list, selected_list, disp_var, ent, field_name))
                return "break"
            ent.bind("<Button-1>", _open); ent.bind("<ButtonPress-1>", _open)
            tk.Button(row, text="Choose…", command=lambda: _make_filter_row(
                field_name, available_list, selected_list, disp_var, ent, field_name),
                bg="#3c3c3c", fg=_FG, relief="flat", bd=0, padx=10, pady=3,
                font=("Segoe UI", 9), cursor="hand2",
                activebackground="#505050").pack(side="left")
            return ent

        # ── Project key ────────────────────────────────────────────────────────
        _lbl("Project key:")
        proj_row = tk.Frame(main, bg=_BG); proj_row.pack(fill="x", pady=(0, 8))
        proj_var = tk.StringVar(value=cfg.get("project_key") or "SUNDANCE")
        tk.Entry(proj_row, textvariable=proj_var, width=18,
                 bg="#3c3c3c", fg=_FG, insertbackground=_FG,
                 relief="flat", bd=1, font=("Segoe UI", 10)).pack(side="left", padx=(0, 8))
        tk.Label(proj_row, text="(default: SUNDANCE)", bg=_BG, fg="#888888",
                 font=("Segoe UI", 9)).pack(side="left")
        _sep()

        # ── Scope ─────────────────────────────────────────────────────────────
        _lbl("Assignee / Reporter:")
        scope_row = tk.Frame(main, bg=_BG); scope_row.pack(fill="x", pady=(0, 8))
        scope_var = tk.StringVar(value=cfg.get("scope") or "assigned")
        for lbl_txt, val in [("Assigned to me", "assigned"), ("Created by me", "created"),
                              ("Assigned OR Created by me", "both"), ("Anyone (no restriction)", "any")]:
            tk.Radiobutton(scope_row, text=lbl_txt, variable=scope_var, value=val,
                           bg=_BG, fg=_FG, selectcolor=_PANEL,
                           activebackground=_BG, activeforeground=_FG,
                           font=("Segoe UI", 9)).pack(anchor="w", pady=1)
        _sep()

        # ── Max results ────────────────────────────────────────────────────────
        _lbl("Max results (1–500):")
        max_row = tk.Frame(main, bg=_BG); max_row.pack(fill="x", pady=(0, 8))
        max_var = tk.StringVar(value=str(cfg.get("max_results") or 50))
        tk.Entry(max_row, textvariable=max_var, width=8,
                 bg="#3c3c3c", fg=_FG, insertbackground=_FG,
                 relief="flat", bd=1, font=("Segoe UI", 10)).pack(side="left", padx=(0, 8))
        _sep()

        # ── Label filter ──────────────────────────────────────────────────────
        selected_label_filter = list(cfg.get("label_filter") or [])
        label_disp_var = tk.StringVar(value="; ".join(selected_label_filter) if selected_label_filter else "None selected")
        _lbl("Labels:")
        _filter_entry(label_disp_var, opts.get("Labels", []), selected_label_filter, "Labels")
        label_mode_row = tk.Frame(main, bg=_BG); label_mode_row.pack(fill="x", pady=(2, 8))
        label_mode_var = tk.StringVar(value=cfg.get("label_mode") or "any")
        tk.Label(label_mode_row, text="Match:", bg=_BG, fg="#888888",
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
        for m in ("any", "all"):
            tk.Radiobutton(label_mode_row, text=m, variable=label_mode_var, value=m,
                           bg=_BG, fg=_FG, selectcolor=_PANEL,
                           activebackground=_BG, activeforeground=_FG,
                           font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        _sep()

        # ── Component filter ──────────────────────────────────────────────────
        selected_comp_filter = list(cfg.get("component_filter") or [])
        comp_disp_var = tk.StringVar(value="; ".join(selected_comp_filter) if selected_comp_filter else "None selected")
        _lbl("Components:")
        _filter_entry(comp_disp_var, opts.get("Components", []), selected_comp_filter, "Components")
        comp_mode_row = tk.Frame(main, bg=_BG); comp_mode_row.pack(fill="x", pady=(2, 8))
        comp_mode_var = tk.StringVar(value=cfg.get("component_mode") or "any")
        tk.Label(comp_mode_row, text="Match:", bg=_BG, fg="#888888",
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
        for m in ("any", "all"):
            tk.Radiobutton(comp_mode_row, text=m, variable=comp_mode_var, value=m,
                           bg=_BG, fg=_FG, selectcolor=_PANEL,
                           activebackground=_BG, activeforeground=_FG,
                           font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        _sep()

        # ── Issue Type filter ─────────────────────────────────────────────────
        selected_type_filter = list(cfg.get("type_filter") or [])
        type_disp_var = tk.StringVar(value="; ".join(selected_type_filter) if selected_type_filter else "None selected")
        _lbl("Issue Types:")
        _filter_entry(type_disp_var, opts.get("Issue Type", []), selected_type_filter, "Issue Type")
        _sep()

        # ── Status filter ─────────────────────────────────────────────────────
        selected_status_filter = list(cfg.get("status_filter") or [])
        status_disp_var = tk.StringVar(value="; ".join(selected_status_filter) if selected_status_filter else "None selected")
        _lbl("Statuses:")
        _filter_entry(status_disp_var, opts.get("Status", []), selected_status_filter, "Status")
        _sep()

        # ── Priority filter ───────────────────────────────────────────────────
        selected_priority_filter = list(cfg.get("priority_filter") or [])
        priority_disp_var = tk.StringVar(value="; ".join(selected_priority_filter) if selected_priority_filter else "None selected")
        _lbl("Priorities:")
        _filter_entry(priority_disp_var, opts.get("Priority", []), selected_priority_filter, "Priority")
        _sep()

        # ── Folder ────────────────────────────────────────────────────────────
        _lbl("Save into folder (optional):")
        folder_row = tk.Frame(main, bg=_BG); folder_row.pack(fill="x", pady=(0, 8))
        folder_var = tk.StringVar(value=cfg.get("folder_name") or "")
        tk.Entry(folder_row, textvariable=folder_var, width=28,
                 bg="#3c3c3c", fg=_FG, insertbackground=_FG,
                 relief="flat", bd=1, font=("Segoe UI", 10)).pack(side="left")

        # ── Footer ────────────────────────────────────────────────────────────
        tk.Frame(dlg, bg=_BORDER, height=1).pack(fill="x", side="bottom")
        footer = tk.Frame(dlg, bg=_PANEL, pady=12, padx=20)
        footer.pack(fill="x", side="bottom")

        last_run = cfg.get("last_run_date") or "never"
        tk.Label(footer, text=f"Last auto-fetch: {last_run}", bg=_PANEL,
                 fg="#888888", font=("Segoe UI", 8)).pack(side="left")

        def _save():
            try:
                cnt = max(1, min(500, int(max_var.get() or 50)))
            except Exception:
                cnt = 50
            new_cfg = {
                "enabled":          enabled_var.get(),
                "scope":            scope_var.get(),
                "project_key":      proj_var.get().strip(),
                "label_filter":     list(selected_label_filter),
                "label_mode":       label_mode_var.get(),
                "component_filter": list(selected_comp_filter),
                "component_mode":   comp_mode_var.get(),
                "type_filter":      list(selected_type_filter),
                "status_filter":    list(selected_status_filter),
                "priority_filter":  list(selected_priority_filter),
                "max_results":      cnt,
                "folder_name":      folder_var.get().strip(),
                "last_run_date":    cfg.get("last_run_date") or "",
            }
            self.meta["auto_fetch_config"] = new_cfg
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass
            dlg.destroy()

        def _save_and_run():
            _save()
            self._run_auto_fetch(force=True)

        _mkb(footer, "Save", _save, _ACCENT, _ACCENT_HOV).pack(side="right", padx=(8, 0))
        _mkb(footer, "Run Now", _save_and_run, "#2a6a2a", "#3a8a3a").pack(side="right")
        _mkb(footer, "Cancel", dlg.destroy, "#3c3c3c", "#505050").pack(side="right")

        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def _run_auto_fetch(self, force: bool = False):
        """Run the configured auto-fetch if enabled and not already run today
        (unless *force* is True).

        When called from ``_startup_sync`` the results are logged into the
        startup log window.  When called via the "Run Now" button in the
        Auto-Fetch Settings dialog it opens the normal fetch progress dialog
        instead.
        """
        cfg = self.meta.get("auto_fetch_config") or {}
        if not cfg.get("enabled"):
            return
        if not force:
            import datetime
            today = datetime.date.today().isoformat()
            if cfg.get("last_run_date") == today:
                return

        # If the startup log window is open, run silently inside it.
        startup_win = getattr(self, "_startup_log_win", None)
        if startup_win and startup_win.winfo_exists() and not force:
            self._run_auto_fetch_silent(cfg)
            return

        # Otherwise (e.g. "Run Now" button) open the full dialog.
        self._start_fetch_issues(
            scope            = cfg.get("scope") or "assigned",
            max_results      = int(cfg.get("max_results") or 50),
            project_key      = (cfg.get("project_key") or "").strip() or "SUNDANCE",
            label_filter     = list(cfg.get("label_filter") or []),
            label_mode       = cfg.get("label_mode") or "any",
            component_filter = list(cfg.get("component_filter") or []),
            component_mode   = cfg.get("component_mode") or "any",
            type_filter      = list(cfg.get("type_filter") or []),
            status_filter    = list(cfg.get("status_filter") or []),
            priority_filter  = list(cfg.get("priority_filter") or []),
            folder_name      = cfg.get("folder_name") or "",
        )
        import datetime
        self.meta["auto_fetch_config"]["last_run_date"] = datetime.date.today().isoformat()
        try:
            save_storage(self.templates, self.meta)
        except Exception:
            pass

    def _run_auto_fetch_silent(self, cfg):
        """Execute the auto-fetch in the background, logging results to the
        startup log window instead of opening a separate dialog."""
        session = self.get_jira_session()
        if not session:
            self._log_startup("Auto-fetch: no Jira session — skipped.", "done")
            return

        scope            = cfg.get("scope") or "assigned"
        max_results      = int(cfg.get("max_results") or 50)
        project_key      = (cfg.get("project_key") or "").strip() or "SUNDANCE"
        label_filter     = list(cfg.get("label_filter") or [])
        label_mode       = cfg.get("label_mode") or "any"
        component_filter = list(cfg.get("component_filter") or [])
        component_mode   = cfg.get("component_mode") or "any"
        type_filter      = list(cfg.get("type_filter") or [])
        status_filter    = list(cfg.get("status_filter") or [])
        priority_filter  = list(cfg.get("priority_filter") or [])
        folder_name      = cfg.get("folder_name") or ""

        self._log_startup("Auto-fetch: running configured fetch...", "step")

        def worker():
            new_added = 0
            failed = 0
            try:
                existing_keys = {str(it.get("Issue key") or "").strip()
                                 for it in self.list_items if it.get("Issue key")}

                clauses = []
                if scope == "assigned":
                    clauses.append("assignee = currentUser()")
                elif scope == "created":
                    clauses.append("reporter = currentUser()")
                elif scope == "both":
                    clauses.append("(assignee = currentUser() OR reporter = currentUser())")

                pk = project_key or "SUNDANCE"
                clauses.append(f'project = "{pk}"')

                def _in_clause(field, items, mode):
                    items = [x.strip() for x in (items or []) if x.strip()]
                    if not items:
                        return None
                    q = lambda x: f'"{x}"'
                    if mode == "all":
                        return " AND ".join(f'{field} = {q(x)}' for x in items)
                    return f'{field} in ({", ".join(q(x) for x in items)})'

                for clause in [
                    _in_clause("labels",    label_filter,    label_mode),
                    _in_clause("component", component_filter, component_mode),
                    _in_clause("issuetype", type_filter,     "any"),
                    _in_clause("status",    status_filter,   "any"),
                    _in_clause("priority",  priority_filter, "any"),
                ]:
                    if clause:
                        clauses.append(clause)

                jql = (" AND ".join(clauses) if clauses else "ORDER BY created DESC").rstrip()
                if "ORDER BY" not in jql.upper():
                    jql += " ORDER BY created DESC"

                self.after(0, lambda: self._log_startup(
                    f"Auto-fetch: searching Jira (excluding {len(existing_keys)} existing)...", "step"))

                sr = self.jira_search_jql_simple(session, jql=jql,
                                                 max_results=max_results,
                                                 exclude_keys=existing_keys)
                issues_list = sr.get("issues", [])
                total = len(issues_list)

                if total == 0:
                    self.after(0, lambda: self._log_startup(
                        "Auto-fetch: no new tickets found.", "done"))
                else:
                    self.after(0, lambda t=total: self._log_startup(
                        f"Auto-fetch: found {t} candidate ticket(s), fetching...", "step"))

                existing_ids = {str(it.get("Issue id") or it.get("Issue key") or "").strip()
                                for it in self.list_items}
                new_issues = []

                for entry in issues_list:
                    issue_id = entry.get("id") or entry.get("key")
                    if not issue_id or str(issue_id) in existing_ids:
                        continue
                    try:
                        issue_json = self.fetch_issue_details(session, issue_id,
                                                             fields=FETCH_FIELDS)
                    except Exception:
                        failed += 1
                        continue
                    fields = issue_json.get("fields", {}) or {}
                    status_obj = fields.get("status") or {}
                    issue_dict = {
                        "Issue key": issue_json.get("key", ""),
                        "Issue id": issue_json.get("id", ""),
                        "Summary": fields.get("summary", "") or "",
                        "Description": "",
                        "Issue Type": (fields.get("issuetype") or {}).get("name", ""),
                        "Status": status_obj.get("name", ""),
                        "Status Category": (status_obj.get("statusCategory") or {}).get("name", ""),
                        "Project key": (fields.get("project") or {}).get("key", ""),
                        "Project name": (fields.get("project") or {}).get("name", ""),
                        "Priority": (fields.get("priority") or {}).get("name", ""),
                        "Assignee": (fields.get("assignee") or {}).get("displayName", "") or
                                    (fields.get("assignee") or {}).get("emailAddress", "") or "",
                        "Reporter": (fields.get("reporter") or {}).get("displayName", "") or
                                    (fields.get("reporter") or {}).get("emailAddress", "") or "",
                        "Created": fields.get("created", ""),
                        "Updated": fields.get("updated", ""),
                        "Labels": "; ".join(fields.get("labels") or []),
                        "Components": "; ".join(c.get("name", "") for c in (fields.get("components") or []))
                    }
                    rendered_html = (issue_json.get("renderedFields") or {}).get("description")
                    if rendered_html:
                        issue_dict["Description Rendered"] = rendered_html
                    desc = fields.get("description", "")
                    if isinstance(desc, str):
                        issue_dict["Description"] = desc
                    elif isinstance(desc, dict):
                        issue_dict["Description ADF"] = desc
                        try:
                            issue_dict["Description"] = self._extract_text_from_adf(desc)
                        except Exception:
                            issue_dict["Description"] = ""
                    else:
                        issue_dict["Description"] = str(desc)
                    issue_dict["Attachment"] = self._jira_attachments_to_field(
                        fields.get("attachment")) or ""
                    issue_dict["Comment"] = self._parse_jira_comments(fields.get("comment"))
                    self._map_epic_and_link_fields(fields, issue_dict)
                    new_issues.append(issue_dict)
                    new_added += 1

                if new_issues:
                    self.list_items.extend(new_issues)
                    self.list_items = _dedup_list_items(self.list_items)
                    if folder_name:
                        tf = self.meta.setdefault("ticket_folders", {})
                        for ni in new_issues:
                            k = ni.get("Issue key") or ni.get("Issue id") or ""
                            if k:
                                tf[k] = folder_name
                        # Also move existing matching tickets into the folder
                        self._assign_folder_to_matching(
                            folder_name, tf,
                            label_filter=label_filter,
                            component_filter=component_filter,
                            type_filter=type_filter,
                            status_filter=status_filter,
                            priority_filter=priority_filter,
                        )
                        folders_list = list(self.meta.get("folders") or [])
                        if folder_name not in folders_list:
                            folders_list.append(folder_name)
                            self.meta["folders"] = folders_list
                    self.meta["fetched_issues"] = list(self.list_items)
                    save_storage(self.templates, self.meta)

                import datetime
                self.meta["auto_fetch_config"]["last_run_date"] = datetime.date.today().isoformat()
                try:
                    save_storage(self.templates, self.meta)
                except Exception:
                    pass

            except Exception:
                debug_log(f"Auto-fetch silent failed: {traceback.format_exc()}")
                self.after(0, lambda: self._log_startup(
                    "Auto-fetch: failed — check debug log.", "done"))
                return

            def done():
                if new_added or failed:
                    parts = []
                    if new_added:
                        parts.append(f"{new_added} new ticket(s) added")
                    if failed:
                        parts.append(f"{failed} failed")
                    self._log_startup(f"Auto-fetch complete: {', '.join(parts)}.", "step")
                else:
                    self._log_startup("Auto-fetch complete: already up to date.", "done")
                if new_added:
                    self._populate_listview()
                    self._update_welcome_text()
            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _start_fetch_issues(self, scope="assigned", max_results=50,
                            project_key="SUNDANCE",
                            label_filter=None, label_mode="any",
                            component_filter=None, component_mode="any",
                            type_filter=None, status_filter=None, priority_filter=None,
                            folder_name=""):
        session = self.get_jira_session()
        if not session:
            return
        progress = tk.Toplevel(self)
        self._register_toplevel(progress)
        progress.title("Fetching issues...")
        progress.minsize(400, 280)
        progress.geometry("560x320")
        progress.resizable(True, True)
        progress.resizable(True, True)
        prog_main = ttk.Frame(progress, padding=16)
        prog_main.pack(fill="both", expand=True)
        ttk.Label(prog_main, text="Fetching issues from Jira — please wait...").pack(fill="x", pady=(0, 8))
        pb = ttk.Progressbar(prog_main, mode="indeterminate")
        pb.pack(fill="x", pady=(0, 12))
        logbox = tk.Text(prog_main, height=8, wrap="word", bg="#1e1e1e", fg="#dcdcdc", insertbackground="#dcdcdc")
        logbox.pack(fill="both", expand=True, pady=(0, 8))
        counts_lbl = ttk.Label(prog_main, text="0 fetched, 0 new, 0 failed")
        counts_lbl.pack(fill="x", pady=(0, 8))
        cancel_flag = {"cancel": False}
        def on_cancel():
            cancel_flag["cancel"] = True
            logbox.insert("end", "Cancel requested; will stop after current request...\n")
            logbox.see("end")
        btn_frame = ttk.Frame(prog_main)
        btn_frame.pack(fill="x", pady=(0, 0))
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side="left")
        try:
            pb.start(10)
        except Exception:
            pass
        try:
            progress
        except Exception:
            pass

        def worker():
            results = {"ok": False, "issues": [], "error": None}
            try:
                try:
                    r = perform_jira_request(session, "GET", f"{session._jira_base}/rest/api/3/myself", timeout=15)
                    if r.status_code == 200:
                        info = r.json()
                        self.meta["jira_current_user"] = {
                            "displayName": (info.get("displayName") or "").strip(),
                            "emailAddress": (info.get("emailAddress") or "").strip()
                        }
                except Exception:
                    pass
                existing_keys = {str(it.get("Issue key") or "").strip() for it in self.list_items if it.get("Issue key")}
                # Scope (who) clause
                clauses = []
                if scope == "assigned":
                    clauses.append("assignee = currentUser()")
                elif scope == "created":
                    clauses.append("reporter = currentUser()")
                elif scope == "both":
                    clauses.append("(assignee = currentUser() OR reporter = currentUser())")
                # "any" scope = no user restriction
                pk = (project_key or "").strip() or "SUNDANCE"
                clauses.append(f'project = "{pk}"')
                def _in_clause(field, items, mode):
                    items = [x.strip() for x in (items or []) if x.strip()]
                    if not items: return None
                    q = lambda x: f'"{x}"'
                    if mode == "all":
                        return " AND ".join(f'{field} = {q(x)}' for x in items)
                    return f'{field} in ({", ".join(q(x) for x in items)})'
                for clause in [
                    _in_clause("labels",    label_filter,    label_mode or "any"),
                    _in_clause("component", component_filter, component_mode or "any"),
                    _in_clause("issuetype", type_filter,     "any"),
                    _in_clause("status",    status_filter,   "any"),
                    _in_clause("priority",  priority_filter, "any"),
                ]:
                    if clause: clauses.append(clause)
                jql = (" AND ".join(clauses) if clauses else "order by created DESC").rstrip()
                if "ORDER BY" not in jql.upper():
                    jql += " ORDER BY created DESC"
                self.after(0, lambda: logbox.insert("end", f"Running JQL (excluding {len(existing_keys)} existing)...\n"))
                sr = self.jira_search_jql_simple(session, jql=jql, max_results=max_results, exclude_keys=existing_keys)
                issues_list = sr.get("issues", [])
                total = len(issues_list)
                fetched = 0
                new_added = 0
                failed = 0
                self.after(0, lambda: counts_lbl.config(text=f"0/{total} fetched, 0 new, 0 failed"))
                existing_ids = {str(it.get("Issue id") or it.get("Issue key") or "").strip() for it in self.list_items}
                for idx, entry in enumerate(issues_list):
                    if cancel_flag["cancel"]:
                        self.after(0, lambda: logbox.insert("end", "Cancelled by user.\n"))
                        break
                    issue_id = entry.get("id") or entry.get("key")
                    if not issue_id:
                        continue
                    if str(issue_id) in existing_ids:
                        continue
                    try:
                        self.after(0, lambda iid=issue_id: logbox.insert("end", f"Fetching details for {iid}...\n"))
                        issue_json = self.fetch_issue_details(session, issue_id, fields=FETCH_FIELDS)
                    except Exception:
                        debug_log(f"Failed to fetch details for {issue_id}: {traceback.format_exc()}")
                        failed += 1
                        fetched += 1
                        self.after(0, lambda: counts_lbl.config(text=f"{fetched}/{total} fetched, {new_added} new, {failed} failed"))
                        continue
                    fields = issue_json.get("fields", {}) or {}
                    status_obj = fields.get("status") or {}
                    issue_dict = {
                        "Issue key": issue_json.get("key", ""),
                        "Issue id": issue_json.get("id", ""),
                        "Summary": fields.get("summary", "") or "",
                        "Description": "",
                        "Issue Type": (fields.get("issuetype") or {}).get("name", ""),
                        "Status": status_obj.get("name", ""),
                        "Status Category": (status_obj.get("statusCategory") or {}).get("name", ""),
                        "Project key": (fields.get("project") or {}).get("key", ""),
                        "Project name": (fields.get("project") or {}).get("name", ""),
                        "Priority": (fields.get("priority") or {}).get("name", ""),
                        "Assignee": (fields.get("assignee") or {}).get("displayName", "") or (fields.get("assignee") or {}).get("emailAddress", "") or "",
                        "Reporter": (fields.get("reporter") or {}).get("displayName", "") or (fields.get("reporter") or {}).get("emailAddress", "") or "",
                        "Created": fields.get("created", ""),
                        "Updated": fields.get("updated", ""),
                        "Labels": "; ".join(fields.get("labels") or []),
                        "Components": "; ".join([c.get("name", "") for c in (fields.get("components") or [])])
                    }
                    # Capture server-rendered HTML if Jira returned it
                    rendered_html = (issue_json.get("renderedFields") or {}).get("description")
                    if rendered_html:
                        issue_dict["Description Rendered"] = rendered_html
                    desc = fields.get("description", "")
                    if isinstance(desc, str):
                        issue_dict["Description"] = desc
                    elif isinstance(desc, dict):
                        issue_dict["Description ADF"] = desc
                        try:
                            issue_dict["Description"] = self._extract_text_from_adf(desc)
                        except Exception:
                            issue_dict["Description"] = ""
                    else:
                        issue_dict["Description"] = str(desc)
                    issue_dict["Attachment"] = self._jira_attachments_to_field(fields.get("attachment")) or ""
                    issue_dict["Comment"] = self._parse_jira_comments(fields.get("comment"))
                    self._map_epic_and_link_fields(fields, issue_dict)
                    results["issues"].append(issue_dict)
                    fetched += 1
                    existing_idx = next((i for i, it in enumerate(self.list_items) if it.get("Issue id") == issue_dict.get("Issue id")), None)
                    if existing_idx is not None:
                        pass
                    else:
                        new_added += 1
                    self.after(0, lambda f=fetched, t=total, n=new_added, fa=failed: counts_lbl.config(text=f"{f}/{t} fetched, {n} new, {fa} failed"))
                results["ok"] = True
            except Exception:
                results["error"] = traceback.format_exc()
            self.after(0, lambda fn=folder_name: self._on_fetch_complete(
                results, progress, logbox, folder_name=fn,
                label_filter=label_filter, component_filter=component_filter,
                type_filter=type_filter, status_filter=status_filter,
                priority_filter=priority_filter))

        thr = threading.Thread(target=worker, daemon=True)
        thr.start()

    def refresh_active_tab_from_jira(self):
        """Refresh the active tab's ticket from Jira."""
        tf = self.get_active_tabform()
        if not tf:
            messagebox.showinfo("Info", "No active tab.")
            return
        data = tf.read_to_dict()
        key_or_id = data.get("Issue key") or data.get("Issue id") or ""
        key_or_id = str(key_or_id).strip()
        if not key_or_id or key_or_id.startswith("LOCAL-"):
            messagebox.showinfo("Info", "No Jira issue in this tab (Issue key/id required).")
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
        fields = issue_json.get("fields", {}) or {}
        status_obj = fields.get("status") or {}
        issue_dict = {
            "Issue key": issue_json.get("key", ""),
            "Issue id": issue_json.get("id", ""),
            "Summary": fields.get("summary", "") or "",
            "Description": "",
            "Issue Type": (fields.get("issuetype") or {}).get("name", ""),
            "Status": status_obj.get("name", ""),
            "Status Category": (status_obj.get("statusCategory") or {}).get("name", ""),
            "Project key": (fields.get("project") or {}).get("key", ""),
            "Project name": (fields.get("project") or {}).get("name", ""),
            "Priority": (fields.get("priority") or {}).get("name", ""),
            "Assignee": (fields.get("assignee") or {}).get("displayName", "") or (fields.get("assignee") or {}).get("emailAddress", "") or "",
            "Reporter": (fields.get("reporter") or {}).get("displayName", "") or (fields.get("reporter") or {}).get("emailAddress", "") or "",
            "Created": fields.get("created", ""),
            "Updated": fields.get("updated", ""),
            "Labels": "; ".join(fields.get("labels") or []),
            "Components": "; ".join([c.get("name", "") for c in (fields.get("components") or [])])
        }
        rendered_html = (issue_json.get("renderedFields") or {}).get("description")
        if rendered_html:
            issue_dict["Description Rendered"] = rendered_html
        desc = fields.get("description", "")
        if isinstance(desc, str):
            issue_dict["Description"] = desc
        elif isinstance(desc, dict):
            issue_dict["Description ADF"] = desc
            try:
                issue_dict["Description"] = self._extract_text_from_adf(desc)
            except Exception:
                issue_dict["Description"] = ""
        else:
            issue_dict["Description"] = str(desc)
        issue_dict["Attachment"] = self._jira_attachments_to_field(fields.get("attachment")) or ""
        issue_dict["Comment"] = self._parse_jira_comments(fields.get("comment"))
        self._map_epic_and_link_fields(fields, issue_dict)
        self._enrich_with_internal_priority(issue_dict)
        tf.populate_from_dict(issue_dict)
        tab_frame = tf.frame
        self._tab_summaries[tab_frame] = issue_dict.get("Summary", "")
        self._rename_tabs()
        # Update list_items if this ticket is in the list
        for i, it in enumerate(self.list_items):
            if (it.get("Issue key") or it.get("Issue id")) == (issue_dict.get("Issue key") or issue_dict.get("Issue id")):
                self.list_items[i] = issue_dict
                break
        self._populate_listview()
        messagebox.showinfo("Refreshed", f"Updated {issue_dict.get('Issue key', '')} from Jira.")

    def _on_fetch_complete(self, results, progress_win, logbox, folder_name="",
                           label_filter=None, component_filter=None,
                           type_filter=None, status_filter=None,
                           priority_filter=None):
        try:
            progress_win.grab_release()
        except Exception:
            pass
        try:
            progress_win.destroy()
        except Exception:
            pass
        if not results.get("ok"):
            err = results.get("error") or "Unknown error"
            messagebox.showerror("Fetch failed", f"Error during fetch:\n{err}\n\nSee {DEBUG_LOG}")
            return
        new_count = 0
        for issue_dict in results.get("issues", []):
            existing_idx = next((i for i, it in enumerate(self.list_items)
                                 if (it.get("Issue id") == issue_dict.get("Issue id") and issue_dict.get("Issue id"))
                                 or (it.get("Issue key") == issue_dict.get("Issue key") and issue_dict.get("Issue key"))), None)
            if existing_idx is not None:
                self.list_items[existing_idx] = issue_dict
            else:
                self.list_items.append(issue_dict)
                new_count += 1
        self.list_items = _dedup_list_items(self.list_items)

        fn = (folder_name or "").strip()
        moved_existing = 0
        if fn:
            if fn not in self.meta.setdefault("folders", []):
                self.meta["folders"].append(fn)
            tf = self.meta.setdefault("ticket_folders", {})

            # Assign folder to newly fetched tickets
            for issue_dict in results.get("issues", []):
                tk_key = str(issue_dict.get("Issue key") or issue_dict.get("Issue id") or "").strip()
                if tk_key:
                    tf[tk_key] = fn

            # Also assign folder to ALL existing tickets that match the
            # same filter criteria, so previously-fetched tickets that
            # belong to the same logical group are included.
            moved_existing = self._assign_folder_to_matching(
                fn, tf,
                label_filter=label_filter,
                component_filter=component_filter,
                type_filter=type_filter,
                status_filter=status_filter,
                priority_filter=priority_filter,
            )
            self._refresh_folder_combo()

        self.meta["fetched_issues"] = list(self.list_items)
        try:
            save_storage(self.templates, self.meta)
        except Exception:
            pass
        self._populate_listview()
        self.show_tabs_view()
        self.notebook.select(self._welcome_frame)
        all_keys = [it.get("Issue key") or it.get("Issue id") for it in results.get("issues", []) if (it.get("Issue key") or it.get("Issue id"))]
        self.meta["welcome_updates"] = {"new": new_count, "refreshed": len(results.get("issues", [])) - new_count, "new_ticket_keys": all_keys}
        self._update_welcome_text()
        if fn:
            extra = f"\nSaved to folder: '{fn}'"
            if moved_existing:
                extra += f" ({moved_existing} existing ticket(s) also moved)"
        else:
            extra = ""
        messagebox.showinfo("Fetched", f"Added/updated {len(results.get('issues',[]))} issues (new: {new_count}).{extra}")
        cb = getattr(self, "_post_fetch_callback", None)
        if cb:
            self._post_fetch_callback = None
            try:
                cb()
            except Exception:
                pass

    def _assign_folder_to_matching(self, folder_name, ticket_folders,
                                    label_filter=None, component_filter=None,
                                    type_filter=None, status_filter=None,
                                    priority_filter=None):
        """Scan all existing list_items and assign the folder to any ticket
        that matches the given filter criteria.  Returns the count of
        existing tickets that were moved."""
        label_set = set(x.strip().lower() for x in (label_filter or []) if x.strip())
        comp_set = set(x.strip().lower() for x in (component_filter or []) if x.strip())
        type_set = set(x.strip().lower() for x in (type_filter or []) if x.strip())
        status_set = set(x.strip().lower() for x in (status_filter or []) if x.strip())
        priority_set = set(x.strip().lower() for x in (priority_filter or []) if x.strip())

        # If no filters were active, don't mass-move everything
        if not any([label_set, comp_set, type_set, status_set, priority_set]):
            return 0

        moved = 0
        for item in self.list_items:
            tk_key = str(item.get("Issue key") or item.get("Issue id") or "").strip()
            if not tk_key:
                continue
            if ticket_folders.get(tk_key) == folder_name:
                continue

            # Check each active filter against the ticket's values
            if label_set:
                item_labels = set(
                    p.strip().lower()
                    for p in (item.get("Labels") or "").replace(",", ";").split(";")
                    if p.strip()
                )
                if not (label_set & item_labels):
                    continue
            if comp_set:
                item_comps = set(
                    p.strip().lower()
                    for p in (item.get("Components") or "").replace(",", ";").split(";")
                    if p.strip()
                )
                if not (comp_set & item_comps):
                    continue
            if type_set:
                if (item.get("Issue Type") or "").strip().lower() not in type_set:
                    continue
            if status_set:
                if (item.get("Status") or "").strip().lower() not in status_set:
                    continue
            if priority_set:
                if (item.get("Priority") or "").strip().lower() not in priority_set:
                    continue

            ticket_folders[tk_key] = folder_name
            moved += 1
        return moved

    def refresh_fetched_tickets(self):
        if not self.list_items:
            messagebox.showinfo("Info", "No fetched tickets to refresh.")
            return
        s = self.get_jira_session()
        if not s:
            return
        progress = tk.Toplevel(self)
        self._register_toplevel(progress)
        progress.title("Refreshing fetched tickets...")
        progress.minsize(400, 300)
        progress.geometry("560x340")
        progress.resizable(True, True)
        progress.resizable(True, True)
        ttk.Label(progress, text="Refreshing fetched tickets from Jira — please wait...").pack(fill="x", padx=12, pady=(12,6))
        pb = ttk.Progressbar(progress, mode="determinate")
        pb.pack(fill="x", padx=12, pady=(0, 12))
        logbox = tk.Text(progress, height=10, wrap="word")
        logbox.pack(fill="both", expand=True, padx=12, pady=(0,8))
        counts_lbl = ttk.Label(progress, text="0/0 checked, 0 updated, 0 failed")
        counts_lbl.pack(fill="x", padx=12, pady=(0,8))
        cancel_flag = {"cancel": False}
        def on_cancel():
            cancel_flag["cancel"] = True
            logbox.insert("end", "Cancel requested; will stop after current request...\n")
            logbox.see("end")
        btn_frame = ttk.Frame(progress)
        btn_frame.pack(fill="x", padx=12, pady=(0,12))
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side="left")
        try:
            progress
        except Exception:
            pass

        def worker():
            total = len(self.list_items)
            checked = 0
            updated = 0
            failed = 0
            self.after(0, lambda: pb.configure(maximum=total, value=0))
            for idx, stored in enumerate(list(self.list_items)):
                if cancel_flag["cancel"]:
                    self.after(0, lambda: logbox.insert("end", "Cancelled by user.\n"))
                    break
                issue_id = stored.get("Issue id") or stored.get("Issue key")
                if not issue_id:
                    checked += 1
                    self.after(0, lambda c=checked, t=total, u=updated, f=failed: counts_lbl.config(text=f"{c}/{t} checked, {u} updated, {f} failed"))
                    self.after(0, lambda v=checked: pb.configure(value=v))
                    continue
                try:
                    self.after(0, lambda iid=issue_id: logbox.insert("end", f"Refreshing {iid}...\n"))
                    issue_json = self.fetch_issue_details(
                        s,
                        issue_id,
                        fields=FETCH_FIELDS
                    )
                except Exception:
                    debug_log(f"Failed to refresh details for {issue_id}: {traceback.format_exc()}")
                    failed += 1
                    checked += 1
                    self.after(0, lambda c=checked, t=total, u=updated, f=failed: counts_lbl.config(text=f"{c}/{t} checked, {u} updated, {f} failed"))
                    self.after(0, lambda v=checked: pb.configure(value=v))
                    continue
                fields = issue_json.get("fields", {}) or {}
                new_updated = fields.get("updated", "")
                if new_updated and new_updated == stored.get("Updated"):
                    checked += 1
                    self.after(0, lambda c=checked, t=total, u=updated, f=failed: counts_lbl.config(text=f"{c}/{t} checked, {u} updated, {f} failed"))
                    self.after(0, lambda v=checked: pb.configure(value=v))
                    continue
                status_obj = fields.get("status") or {}
                new_dict = {
                    "Issue key": issue_json.get("key", ""),
                    "Issue id": issue_json.get("id", ""),
                    "Summary": fields.get("summary", "") or "",
                    "Description": "",
                    "Issue Type": (fields.get("issuetype") or {}).get("name", ""),
                    "Status": status_obj.get("name", ""),
                    "Status Category": (status_obj.get("statusCategory") or {}).get("name", ""),
                    "Project key": (fields.get("project") or {}).get("key", ""),
                    "Project name": (fields.get("project") or {}).get("name", ""),
                    "Priority": (fields.get("priority") or {}).get("name", ""),
                    "Assignee": (fields.get("assignee") or {}).get("displayName", "") or (fields.get("assignee") or {}).get("emailAddress", "") or "",
                    "Reporter": (fields.get("reporter") or {}).get("displayName", "") or (fields.get("reporter") or {}).get("emailAddress", "") or "",
                    "Created": fields.get("created", ""),
                    "Updated": fields.get("updated", ""),
                    "Labels": "; ".join(fields.get("labels") or []),
                    "Components": "; ".join([c.get("name", "") for c in (fields.get("components") or [])])
                }
                rendered_html = (issue_json.get("renderedFields") or {}).get("description")
                if rendered_html:
                    new_dict["Description Rendered"] = rendered_html
                desc = fields.get("description", "")
                if isinstance(desc, str):
                    new_dict["Description"] = desc
                elif isinstance(desc, dict):
                    new_dict["Description ADF"] = desc
                    try:
                        new_dict["Description"] = self._extract_text_from_adf(desc)
                    except Exception:
                        new_dict["Description"] = ""
                else:
                    new_dict["Description"] = str(desc)
                new_dict["Attachment"] = self._jira_attachments_to_field(fields.get("attachment")) or ""
                new_dict["Comment"] = self._parse_jira_comments(fields.get("comment"))
                self._map_epic_and_link_fields(fields, new_dict)
                self.list_items[idx] = new_dict
                updated += 1
                checked += 1
                self.after(0, lambda c=checked, t=total, u=updated, f=failed: counts_lbl.config(text=f"{c}/{t} checked, {u} updated, {f} failed"))
                self.after(0, lambda v=checked: pb.configure(value=v))
            self.meta["fetched_issues"] = list(self.list_items)
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass
            self.after(0, lambda: self._on_refresh_complete(updated, failed, progress))

        thr = threading.Thread(target=worker, daemon=True)
        thr.start()

    def _on_refresh_complete(self, updated_count, failed_count, progress_win):
        try:
            progress_win.grab_release()
        except Exception:
            pass
        try:
            progress_win.destroy()
        except Exception:
            pass
        self._populate_listview()
        self.meta["welcome_updates"] = {"refreshed": updated_count, "sync_status": f"{updated_count} refreshed, {failed_count} failed"}
        self._update_welcome_text()
        messagebox.showinfo("Refreshed", f"Refreshed {updated_count} tickets. Failed: {failed_count}")
