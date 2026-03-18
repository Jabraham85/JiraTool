"""
RemindersMixin — Configure reminders, on-open reminder, stale/blocked checks.
"""
import json
import datetime
import traceback
import time
import tkinter as tk
from tkinter import ttk, messagebox

from config import _JIRA_FIELDS_FALLBACK
from storage import perform_jira_request, save_storage
from utils import debug_log, _bind_mousewheel


class RemindersMixin:
    """Mixin providing configure_reminders_dialog, _schedule_on_open_reminder, _do_on_open_reminder, _parse_jira_updated_date, _get_last_substantive_update_date, _check_reminders, _schedule_reminder_tick."""

    def configure_reminders_dialog(self, on_close=None):
        """Configure internal priority levels, options per level, and reminder rules."""
        if self._focus_existing_app_dialog("reminders"):
            return
        win = tk.Toplevel(self)
        self._track_app_dialog("reminders", win)
        self._register_toplevel(win)
        win.title("Configure Reminders")
        win.minsize(520, 520)
        win.geometry("660x640")
        win.resizable(True, True)
        main = ttk.Frame(win, padding=12)
        main.pack(fill="both", expand=True)
        single_popup_var = tk.BooleanVar(value=self.meta.get("reminder_single_popup", True))
        single_popup_cb = tk.Checkbutton(main, text="Show all reminders in a single popup (recommended)", variable=single_popup_var, bg="#2b2b2b", fg="#dcdcdc", selectcolor="#3c3c3c", activebackground="#2b2b2b", activeforeground="#dcdcdc", highlightthickness=0)
        single_popup_cb.pack(anchor="w", pady=(0, 12))
        tutorial_enabled_var = tk.BooleanVar(value=self.meta.get("tutorial_enabled", True))
        tk.Checkbutton(main, text="Show tutorial on first startup", variable=tutorial_enabled_var, bg="#2b2b2b", fg="#dcdcdc", selectcolor="#3c3c3c", activebackground="#2b2b2b", activeforeground="#dcdcdc", highlightthickness=0).pack(anchor="w", pady=(0, 12))
        ttk.Label(main, text="Internal priority levels (comma-separated, order = high to low):").pack(anchor="w", pady=(0, 4))
        levels_var = tk.StringVar(value=", ".join(self.meta.get("internal_priority_levels", ["High", "Medium", "Low", "None"])))
        levels_entry = tk.Entry(main, textvariable=levels_var, width=50, bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc", relief="flat", bd=2)
        levels_entry.pack(fill="x", pady=(0, 12))
        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=(0, 12))
        stale_var = tk.BooleanVar(value=self.meta.get("stale_ticket_enabled", False))
        stale_days_var = tk.StringVar(value=str(self.meta.get("stale_ticket_days", 14)))
        stale_row = ttk.Frame(main)
        stale_row.pack(fill="x", pady=(0, 8))
        stale_cb = tk.Checkbutton(stale_row, text="Stale ticket reminder:", variable=stale_var, bg="#2b2b2b", fg="#dcdcdc", selectcolor="#3c3c3c", activebackground="#2b2b2b", activeforeground="#dcdcdc", highlightthickness=0)
        stale_cb.pack(side="left", padx=(0, 8))
        ttk.Label(stale_row, text="Warn when an open ticket has no Jira updates for").pack(side="left", padx=(0, 4))
        stale_days_entry = tk.Entry(stale_row, textvariable=stale_days_var, width=4, bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc", relief="flat", bd=2)
        stale_days_entry.pack(side="left", padx=(0, 4))
        ttk.Label(stale_row, text="days").pack(side="left")
        stale_ignore_row = ttk.Frame(main)
        stale_ignore_row.pack(fill="x", pady=(4, 0))
        ttk.Label(stale_ignore_row, text="Ignored fields:").pack(side="left", padx=(0, 8))
        ignored_fields_raw = [x.strip().lower() for x in self.meta.get("stale_ticket_ignored_fields", []) if x]
        # Mutable state shared with the popup
        stale_ignore_field_list = []
        selected_ignored_ids = list(ignored_fields_raw)

        ignore_summary_var = tk.StringVar()
        def _refresh_ignore_summary():
            n = len(selected_ignored_ids)
            ignore_summary_var.set(f"{n} field{'s' if n != 1 else ''} ignored" if n else "None")
        _refresh_ignore_summary()

        ignore_disp = tk.Entry(stale_ignore_row, textvariable=ignore_summary_var, state="readonly",
                               width=22, bg="#3c3c3c", fg="#aaaaaa", disabledforeground="#aaaaaa",
                               relief="flat", bd=2, readonlybackground="#3c3c3c")
        ignore_disp.pack(side="left", padx=(0, 8))

        def _open_ignore_fields_dialog():
            _BG         = "#1a1a1a"
            _PANEL      = "#252526"
            _BORDER     = "#3c3c3c"
            _FG         = "#d4d4d4"
            _SEL_BG     = "#0e639c"
            _BTN_BG     = "#3c3c3c"
            _BTN_HOV    = "#505050"
            _BTN_OK     = "#0e639c"
            _BTN_OK_HOV = "#1177bb"
            _SEARCH_BG  = "#2d2d2d"

            popup = tk.Toplevel(win)
            popup.title("Select ignored fields")
            popup.configure(bg=_BG)
            popup.minsize(320, 400)
            popup.geometry("400x500")
            popup.resizable(True, True)
            popup

            # Track selections by field id across filter changes
            sel_ids_set = set(selected_ignored_ids)
            visible_pairs = []   # (fid, fname) currently shown

            # Title bar
            title_bar = tk.Frame(popup, bg=_PANEL, pady=10)
            title_bar.pack(fill="x")
            tk.Label(title_bar, text="Select ignored fields", bg=_PANEL, fg=_FG,
                     font=("Segoe UI", 11, "bold"), padx=16).pack(side="left")
            tk.Frame(popup, bg=_BORDER, height=1).pack(fill="x")

            # Search bar
            search_frame = tk.Frame(popup, bg=_BG, padx=12)
            search_frame.pack(fill="x", pady=(10, 4))
            search_var = tk.StringVar()
            search_entry = tk.Entry(
                search_frame, textvariable=search_var, bg=_SEARCH_BG, fg=_FG,
                insertbackground=_FG, relief="flat", bd=0, font=("Segoe UI", 10),
                highlightthickness=1, highlightbackground=_BORDER, highlightcolor="#007acc"
            )
            search_entry.pack(fill="x", ipady=6, padx=1)
            search_entry.insert(0, "Search...")
            search_entry.configure(fg="#666666")
            def _sf_in(e):
                if search_var.get() == "Search...":
                    search_entry.delete(0, tk.END)
                    search_entry.configure(fg=_FG)
            def _sf_out(e):
                if not search_var.get().strip():
                    search_entry.delete(0, tk.END)
                    search_entry.insert(0, "Search...")
                    search_entry.configure(fg="#666666")
            search_entry.bind("<FocusIn>", _sf_in)
            search_entry.bind("<FocusOut>", _sf_out)

            # List
            list_outer = tk.Frame(popup, bg=_BG, padx=12)
            list_outer.pack(fill="both", expand=True, pady=(4, 0))
            lb_frame = tk.Frame(list_outer, bg=_BORDER)
            lb_frame.pack(fill="both", expand=True)
            lb = tk.Listbox(
                lb_frame, selectmode="multiple", activestyle="none",
                bg="#1e1e1e", fg=_FG, selectbackground=_SEL_BG, selectforeground="#ffffff",
                font=("Segoe UI", 10), bd=0, highlightthickness=1,
                highlightbackground=_BORDER, highlightcolor="#007acc",
                relief="flat", cursor="hand2"
            )
            sb = tk.Scrollbar(lb_frame, orient="vertical", command=lb.yview,
                              bg=_PANEL, troughcolor=_BG, width=10)
            lb.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            lb.configure(yscrollcommand=sb.set)
            _bind_mousewheel(lb, "vertical")

            # Count label — defined first so helpers below can call it
            count_var = tk.StringVar()
            _rebuilding = [False]

            def _upd_count(*_):
                count_var.set(f"{len(sel_ids_set)} selected")

            def _sync_lb_to_set():
                for i, (fid, _) in enumerate(visible_pairs):
                    if lb.selection_includes(i):
                        sel_ids_set.add(fid.lower())
                    else:
                        sel_ids_set.discard(fid.lower())

            def _rebuild(query=""):
                _rebuilding[0] = True
                _sync_lb_to_set()
                q = query.lower().strip()
                all_pairs = stale_ignore_field_list
                visible_pairs[:] = [(fid, fn) for fid, fn in all_pairs
                                    if q in fn.lower() or q in fid.lower()] if q else list(all_pairs)
                lb.delete(0, tk.END)
                for fid, fname in visible_pairs:
                    lb.insert(tk.END, f"  {fname}  ({fid})")
                for i, (fid, _) in enumerate(visible_pairs):
                    if fid.lower() in sel_ids_set:
                        lb.selection_set(i)
                _rebuilding[0] = False
                _upd_count()

            def _on_search(*_):
                q = search_var.get()
                if q == "Search...":
                    q = ""
                _rebuild(q)

            def _on_lb_select(*_):
                if _rebuilding[0]:
                    return
                _sync_lb_to_set()
                _upd_count()

            search_var.trace_add("write", lambda *_: _on_search())
            lb.bind("<<ListboxSelect>>", _on_lb_select)

            tk.Label(popup, textvariable=count_var, bg=_BG, fg="#888888",
                     font=("Segoe UI", 9), anchor="w", padx=14).pack(fill="x", pady=(4, 0))

            tk.Frame(popup, bg=_BORDER, height=1).pack(fill="x", pady=(6, 0))
            btn_frame = tk.Frame(popup, bg=_PANEL, pady=10, padx=12)
            btn_frame.pack(fill="x")

            def _make_btn(parent, text, cmd, bg, hov):
                b = tk.Button(parent, text=text, command=cmd, bg=bg, fg="#ffffff",
                              font=("Segoe UI", 9, "bold"), relief="flat", bd=0,
                              padx=16, pady=6, cursor="hand2", activebackground=hov,
                              activeforeground="#ffffff")
                b.bind("<Enter>", lambda e: b.configure(bg=hov))
                b.bind("<Leave>", lambda e: b.configure(bg=bg))
                return b

            def ok():
                _sync_lb_to_set()
                selected_ignored_ids[:] = list(sel_ids_set)
                _refresh_ignore_summary()
                popup.destroy()

            def _clear_all():
                sel_ids_set.clear()
                lb.selection_clear(0, tk.END)
                _upd_count()

            _make_btn(btn_frame, "OK", ok, _BTN_OK, _BTN_OK_HOV).pack(side="right", padx=(4, 0))
            _make_btn(btn_frame, "Cancel", popup.destroy, _BTN_BG, _BTN_HOV).pack(side="right")
            _make_btn(btn_frame, "Clear", _clear_all, "#5a3030", "#7a4040").pack(side="left")
            popup.bind("<Return>", lambda e: ok())
            popup.bind("<Escape>", lambda e: popup.destroy())

            # Load fields — use cached list or fetch in background
            def _fill_lb():
                _rebuild()
                _upd_count()
                popup.after(50, search_entry.focus_set)

            if stale_ignore_field_list:
                _fill_lb()
            else:
                import threading as _thr
                def _load_bg():
                    s = self.get_jira_session()
                    fields = self._fetch_jira_fields(s) if s else _JIRA_FIELDS_FALLBACK
                    stale_ignore_field_list[:] = fields
                    try:
                        popup.after(0, _fill_lb)
                    except Exception:
                        pass
                _thr.Thread(target=_load_bg, daemon=True).start()

        ttk.Button(stale_ignore_row, text="Choose fields...", command=_open_ignore_fields_dialog).pack(side="left")
        ttk.Label(main, text="(Uses changelog when ignored fields set; else ticket 'Updated' date; checked on startup and every minute)").pack(anchor="w", pady=(4, 12))
        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=(0, 12))
        ttk.Label(main, text="Blocked status reminder (Jira statuses like Blocked, Blocked by, etc.):").pack(anchor="w", pady=(0, 4))
        blocked_row = ttk.Frame(main)
        blocked_row.pack(fill="x", pady=(0, 8))
        blocked_names_var = tk.StringVar(value=", ".join(self.meta.get("blocked_status_names", ["Blocked"])))
        ttk.Label(blocked_row, text="Status names:").pack(side="left", padx=(0, 4))
        blocked_names_e = tk.Entry(blocked_row, textvariable=blocked_names_var, width=28, bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc", relief="flat", bd=2)
        blocked_names_e.pack(side="left", padx=(0, 8))
        ttk.Label(blocked_row, text="Rule:").pack(side="left", padx=(0, 4))
        blocked_cfg = self.meta.get("blocked_reminder_config", {})
        if isinstance(blocked_cfg, dict):
            bt = blocked_cfg.get("type", "daily")
            btm = blocked_cfg.get("time", "09:00")
            blocked_rule_val = f"time:{btm}" if bt == "time" else bt
        else:
            blocked_rule_val = str(blocked_cfg)
        blocked_rule_var = tk.StringVar(value=blocked_rule_val)
        blocked_rule_e = tk.Entry(blocked_row, textvariable=blocked_rule_var, width=12, bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc", relief="flat", bd=2)
        blocked_rule_e.pack(side="left", fill="x", expand=True)
        ttk.Label(main, text="(daily/weekly/on_open/time:HH:MM/never — checked on startup and every minute)").pack(anchor="w", pady=(0, 12))
        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=(0, 12))
        ttk.Label(main, text="Per level: options (comma-separated) and reminder rule (daily/weekly/on_open/time:HH:MM/never):").pack(anchor="w", pady=(0, 4))
        cfg = self.meta.get("reminder_config", {})
        opts_cfg = self.meta.get("internal_priority_options", {})
        entries = {}
        opt_entries = {}
        for lvl in self.meta.get("internal_priority_levels", ["High", "Medium", "Low", "None"]):
            row = ttk.Frame(main)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"{lvl}:", width=10, anchor="w").pack(side="left", padx=(0, 8))
            r = cfg.get(lvl, {})
            if isinstance(r, dict):
                typ = r.get("type", "never")
                tm = r.get("time", "09:00")
                val = f"time:{tm}" if typ == "time" else typ
            else:
                val = str(r)
            entries[lvl] = tk.StringVar(value=val)
            opt_list = opts_cfg.get(lvl, [lvl])
            opt_entries[lvl] = tk.StringVar(value=", ".join(opt_list) if isinstance(opt_list, list) else str(opt_list))
            ttk.Label(row, text="Options:").pack(side="left", padx=(0, 4))
            opt_e = tk.Entry(row, textvariable=opt_entries[lvl], width=18, bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc", relief="flat", bd=2)
            opt_e.pack(side="left", padx=(0, 8))
            ttk.Label(row, text="Rule:").pack(side="left", padx=(0, 4))
            rule_e = tk.Entry(row, textvariable=entries[lvl], width=12, bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc", relief="flat", bd=2)
            rule_e.pack(side="left", fill="x", expand=True)
        def save():
            levels = [x.strip() for x in levels_var.get().split(",") if x.strip()]
            if not levels:
                messagebox.showerror("Error", "At least one priority level required.")
                return
            self.meta["internal_priority_levels"] = levels
            new_cfg = {}
            new_opts = {}
            all_options = []
            for lvl in levels:
                ent = entries.get(lvl)
                v = ent.get().strip().lower() if ent else "never"
                if v.startswith("time:"):
                    new_cfg[lvl] = {"type": "time", "time": v[5:].strip() or "09:00"}
                elif v in ("daily", "weekly", "on_open", "never"):
                    new_cfg[lvl] = {"type": v}
                else:
                    new_cfg[lvl] = {"type": "never"}
                opt_ent = opt_entries.get(lvl)
                opt_str = opt_ent.get().strip() if opt_ent else lvl
                opt_list = [x.strip() for x in opt_str.split(",") if x.strip()] or [lvl]
                new_opts[lvl] = opt_list
                all_options.extend(opt_list)
            self.meta["reminder_config"] = new_cfg
            self.meta["internal_priority_options"] = new_opts
            try:
                self.meta["stale_ticket_enabled"] = stale_var.get()
                self.meta["stale_ticket_days"] = max(1, min(365, int(stale_days_var.get() or 14)))
            except (ValueError, TypeError):
                self.meta["stale_ticket_days"] = 14
            self.meta["stale_ticket_ignored_fields"] = list(selected_ignored_ids)
            blocked_names = [x.strip() for x in blocked_names_var.get().split(",") if x.strip()]
            self.meta["blocked_status_names"] = blocked_names if blocked_names else ["Blocked"]
            br = blocked_rule_var.get().strip().lower()
            if br.startswith("time:"):
                self.meta["blocked_reminder_config"] = {"type": "time", "time": br[5:].strip() or "09:00"}
            elif br in ("daily", "weekly", "on_open", "never"):
                self.meta["blocked_reminder_config"] = {"type": br}
            else:
                self.meta["blocked_reminder_config"] = {"type": "never"}
            opt_to_level = {}
            for lvl, opts in new_opts.items():
                for o in opts:
                    opt_to_level[o] = lvl
            self.meta["internal_priority_option_to_level"] = opt_to_level
            self.meta["reminder_single_popup"] = single_popup_var.get()
            self.meta["tutorial_enabled"] = tutorial_enabled_var.get()
            save_storage(self.templates, self.meta)
            for tf in self.tabs.values():
                tf._internal_priority_options = all_options
            messagebox.showinfo("Saved", "Reminder configuration saved.")
            win.destroy()
            if on_close:
                self.after(50, on_close)
        def _cancel():
            win.destroy()
            if on_close:
                self.after(50, on_close)
        ttk.Button(main, text="Save", command=save).pack(anchor="w", pady=(16, 0))
        ttk.Button(main, text="Cancel", command=_cancel).pack(anchor="w", pady=(4, 0))
        win.protocol("WM_DELETE_WINDOW", _cancel)

    def _schedule_on_open_reminder(self, ticket_key=None):
        """
        Schedule an on_open reminder check. When reminder_single_popup is True (default),
        debounces multiple calls and shows all reminders in one popup.
        When False, shows a popup per ticket (legacy behavior).
        """
        single = self.meta.get("reminder_single_popup", True)
        if single:
            if self._reminder_on_open_after_id:
                try:
                    self.after_cancel(self._reminder_on_open_after_id)
                except Exception:
                    pass
            self._reminder_on_open_after_id = self.after(450, lambda: self._do_on_open_reminder())
        else:
            self.after(300, lambda k=ticket_key: self._check_reminders("on_open", k))

    def _do_on_open_reminder(self):
        """Run consolidated on_open reminder check (all reminders in one popup)."""
        self._reminder_on_open_after_id = None
        self._check_reminders("on_open", None)

    def _parse_jira_updated_date(self, s):
        """Parse Jira 'Updated' field (e.g. 2025-02-26T10:30:00.000+0000) to date, or None."""
        if not s:
            return None
        try:
            s = str(s).strip()
            if "T" in s:
                s = s.split("T")[0]
            return datetime.datetime.strptime(s[:10], "%Y-%m-%d").date()
        except Exception:
            return None

    def _get_last_substantive_update_date(self, session, issue_key, ignored_fields):
        """Fetch changelog and return the date of the most recent update that changed a non-ignored field.
        ignored_fields: set of field names (lowercase) to ignore, e.g. {'labels', 'status'}.
        Returns date or None if fetch fails or all updates are to ignored fields."""
        if not ignored_fields:
            return None
        try:
            url = f"{session._jira_base}/rest/api/3/issue/{issue_key}"
            resp = perform_jira_request(session, "GET", url, params={"expand": "changelog", "fields": "summary"}, timeout=30)
            if resp.status_code != 200:
                return None
            data = resp.json()
            changelog = data.get("changelog") or {}
            histories = changelog.get("histories") or []
            ignored = set(f.strip().lower() for f in ignored_fields if f)
            for h in reversed(histories):
                items = h.get("items") or []
                has_substantive = False
                for item in items:
                    field = (item.get("field") or "").strip().lower()
                    if field and field not in ignored:
                        has_substantive = True
                        break
                if has_substantive:
                    created = h.get("created") or ""
                    return self._parse_jira_updated_date(created)
            return None
        except Exception:
            return None

    def _check_reminders(self, trigger="on_open", ticket_key=None):
        """Check and show reminders. trigger: 'on_open'|'startup'|'tick'."""
        try:
            cfg = self.meta.get("reminder_config", {})
            opt_to_level = self.meta.get("internal_priority_option_to_level", {})
            priorities = self.meta.get("internal_priorities", {})
            last = self.meta.get("last_reminder", {})
            shown = getattr(self, "_reminder_shown_session", set())
            now = datetime.datetime.now()
            to_show = []
            for key, prio in priorities.items():
                if prio == "None" or not prio:
                    continue
                level = opt_to_level.get(prio, prio)
                r = cfg.get(level, {"type": "never"})
                if isinstance(r, str):
                    r = {"type": r}
                typ = r.get("type", "never")
                if typ == "never":
                    continue
                if trigger == "on_open" and ticket_key and key != ticket_key:
                    continue
                if key in shown:
                    continue
                if typ == "on_open":
                    if trigger in ("on_open", "startup"):
                        to_show.append((key, prio))
                    continue
                last_t = last.get(key)
                if typ == "daily":
                    if not last_t or (now - datetime.datetime.fromisoformat(last_t)).days >= 1:
                        to_show.append((key, prio))
                elif typ == "weekly":
                    if not last_t or (now - datetime.datetime.fromisoformat(last_t)).days >= 7:
                        to_show.append((key, prio))
                elif typ == "time":
                    tm = r.get("time", "09:00")
                    try:
                        h, m = map(int, tm.split(":")[:2])
                        target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                        if now >= target and (not last_t or last_t[:10] < now.strftime("%Y-%m-%d")):
                            to_show.append((key, prio))
                    except Exception:
                        pass
            # Blocked status check: tickets whose Status is in blocked_status_names
            blocked_names = [s.strip().lower() for s in self.meta.get("blocked_status_names", ["Blocked"]) if s.strip()]
            blocked_cfg = self.meta.get("blocked_reminder_config", {})
            if isinstance(blocked_cfg, str):
                blocked_cfg = {"type": blocked_cfg}
            blocked_typ = blocked_cfg.get("type", "never")
            if blocked_names and blocked_typ != "never":
                for it in self.list_items:
                    key = it.get("Issue key") or it.get("Issue id")
                    if not key or str(key).strip().startswith("LOCAL-"):
                        continue
                    status = (it.get("Status") or "").strip().lower()
                    if status not in blocked_names:
                        continue
                    blocked_key = f"blocked:{key}"
                    if blocked_key in shown:
                        continue
                    if trigger == "on_open" and ticket_key and key != ticket_key:
                        continue
                    if blocked_typ == "on_open":
                        if trigger in ("on_open", "startup"):
                            to_show.append((key, "(blocked)"))
                        continue
                    last_t = last.get(blocked_key)
                    if blocked_typ == "daily":
                        if not last_t or (now - datetime.datetime.fromisoformat(last_t)).days >= 1:
                            to_show.append((key, "(blocked)"))
                    elif blocked_typ == "weekly":
                        if not last_t or (now - datetime.datetime.fromisoformat(last_t)).days >= 7:
                            to_show.append((key, "(blocked)"))
                    elif blocked_typ == "time":
                        tm = blocked_cfg.get("time", "09:00")
                        try:
                            h, m = map(int, tm.split(":")[:2])
                            target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                            if now >= target and (not last_t or last_t[:10] < now.strftime("%Y-%m-%d")):
                                to_show.append((key, "(blocked)"))
                        except Exception:
                            pass
            # Stale ticket check: open tabs with no Jira updates in X days
            if self.meta.get("stale_ticket_enabled") and trigger in ("startup", "tick"):
                stale_days = max(1, self.meta.get("stale_ticket_days", 14))
                cutoff = now.date() - datetime.timedelta(days=stale_days)
                welcome = getattr(self, "_welcome_frame", None)
                for tab_id in self.notebook.tabs():
                    try:
                        w = self.nametowidget(tab_id)
                        if w is welcome:
                            continue
                        tf = self.tabs.get(w)
                        if not tf:
                            continue
                        d = tf.read_to_dict()
                        key = d.get("Issue key") or d.get("Issue id")
                        if not key or str(key).strip().startswith("LOCAL-"):
                            continue
                        stale_key = f"stale:{key}"
                        if stale_key in shown:
                            continue
                        last_t = last.get(stale_key)
                        if last_t and (now - datetime.datetime.fromisoformat(last_t)).days < 1:
                            continue
                        ignored_fields = self.meta.get("stale_ticket_ignored_fields", [])
                        if ignored_fields:
                            s = self.get_jira_session()
                            if s:
                                substantive_date = self._get_last_substantive_update_date(s, key, ignored_fields)
                                if substantive_date is not None:
                                    updated_date = substantive_date
                                else:
                                    updated_date = self._parse_jira_updated_date(d.get("Updated", ""))
                            else:
                                updated_date = self._parse_jira_updated_date(d.get("Updated", ""))
                        else:
                            updated_str = d.get("Updated", "")
                            updated_date = self._parse_jira_updated_date(updated_str)
                        if updated_date is None:
                            continue
                        if updated_date <= cutoff:
                            to_show.append((key, "(stale)"))
                    except Exception:
                        pass
            if to_show:
                # Deduplicate: each ticket gets at most one reminder
                seen_keys = set()
                to_show_dedup = []
                for k, p in to_show:
                    if k not in seen_keys:
                        seen_keys.add(k)
                        to_show_dedup.append((k, p))
                to_show = to_show_dedup
                keys_to_show = [k for k, _ in to_show]
                items = [it for it in self.list_items if (it.get("Issue key") or it.get("Issue id")) in keys_to_show]
                if trigger == "startup" and items:
                    self.show_tabs_view()
                    self.new_tab(initial_data=items[0], select_tab=False)
                    self.after(0, lambda: self.notebook.select(self._welcome_frame))
                has_stale = any(p == "(stale)" for _, p in to_show)
                has_blocked = any(p == "(blocked)" for _, p in to_show)
                names = []
                for k in keys_to_show[:5]:
                    it = next((x for x in self.list_items if (x.get("Issue key") or x.get("Issue id")) == k), None)
                    names.append(f"{it.get('Issue key', k)} — {(it.get('Summary') or '')[:50]}" if it else str(k))
                if has_stale and len(to_show) == sum(1 for _, p in to_show if p == "(stale)"):
                    msg = "Stale ticket reminder:\n"
                elif has_blocked and len(to_show) == sum(1 for _, p in to_show if p == "(blocked)"):
                    msg = "Blocked ticket reminder:\n"
                else:
                    msg = "Reminder:\n"
                msg += "\n".join(names)
                if len(items) > 5 or len(keys_to_show) > 5:
                    msg += f"\n... and {max(len(items), len(keys_to_show)) - 5} more"
                now_secs = time.time()
                if now_secs - getattr(self, "_last_reminder_popup_time", 0) < 3.0:
                    pass
                else:
                    self._last_reminder_popup_time = now_secs
                    win = getattr(self, "_startup_log_win", None)
                    if win and win.winfo_exists():
                        self._set_startup_title("Reminders")
                        self._log_startup("")
                        for line in msg.splitlines():
                            self._log_startup(line, "reminder")
                    else:
                        messagebox.showinfo("Reminders", msg)
                for key, prio in to_show:
                    rk = f"stale:{key}" if prio == "(stale)" else (f"blocked:{key}" if prio == "(blocked)" else key)
                    self.meta.setdefault("last_reminder", {})[rk] = now.isoformat()
                    shown.add(rk)
                self._reminder_shown_session = shown
                save_storage(self.templates, self.meta)
        except Exception:
            pass

    def _schedule_reminder_tick(self):
        """Run reminder check periodically for daily/weekly/time-based."""
        try:
            self._check_reminders("tick")
        except Exception:
            pass
        self.after(60000, self._schedule_reminder_tick)
