"""
AvalancheApp — main application class inheriting from all mixins.
"""
import os
import sys
import json
import traceback
import threading
import tkinter as tk
from tkinter import ttk, messagebox

from config import TEMPLATES_FILE, HEADERS, DEBUG_LOG, FETCH_FIELDS, APP_VERSION
from storage import load_storage, save_storage
from utils import debug_log, _bind_mousewheel, _bind_mousewheel_to_target, _dedup_list_items
from tab_form import TabForm
from jira_api import JiraAPIMixin
from ui_builder import UIBuilderMixin
from tab_management import TabManagementMixin
from list_view import ListViewMixin
from variables import VariablesMixin
from upload import UploadMixin
from tutorial import TutorialMixin
from dialogs.credentials import CredentialsMixin
from dialogs.attachments import AttachmentsMixin
from dialogs.fetch import FetchMixin
from dialogs.bulk_import import BulkImportMixin
from dialogs.mass_edit import MassEditMixin
from dialogs.reminders import RemindersMixin
from dialogs.upload_dialog import UploadDialogMixin
from dialogs.bundle_share import BundleShareMixin
from dialogs.updater import UpdaterMixin
from kanban import KanbanMixin


class AvalancheApp(
    UIBuilderMixin,
    TabManagementMixin,
    JiraAPIMixin,
    ListViewMixin,
    KanbanMixin,
    VariablesMixin,
    UploadMixin,
    TutorialMixin,
    CredentialsMixin,
    AttachmentsMixin,
    FetchMixin,
    BulkImportMixin,
    MassEditMixin,
    RemindersMixin,
    UploadDialogMixin,
    BundleShareMixin,
    UpdaterMixin,
    tk.Tk,
):
    def __init__(self):
        super().__init__()
        self.title(f"Avalanche Jira Template Creator  v{APP_VERSION}")
        self.minsize(900, 600)
        self.resizable(True, True)
        self.geometry("1280x900")
        self.after(100, self._apply_initial_geometry)
        self.templates, self.meta = load_storage()
        self.meta.setdefault("options", {})
        self.meta.setdefault("jira", {})
        self.meta.setdefault("variables", {})
        for h in HEADERS:
            self.meta["options"].setdefault(h, [])
        self.meta.setdefault("fetched_issues", [])
        self.meta.setdefault("user_cache", {})
        self.meta.setdefault("folders", [])
        self.meta.setdefault("ticket_folders", {})
        self.meta.setdefault("auto_fetch_config", {
            "enabled":          False,
            "scope":            "assigned",
            "project_key":      "SUNDANCE",
            "label_filter":     [],
            "label_mode":       "any",
            "component_filter": [],
            "component_mode":   "any",
            "type_filter":      [],
            "status_filter":    [],
            "priority_filter":  [],
            "max_results":      50,
            "folder_name":      "",
            "last_run_date":    "",   # ISO date "YYYY-MM-DD" of last auto-fetch
        })
        self.tabs = {}
        self._tab_summaries = {}  # tab_frame -> full summary for tooltip
        self._template_to_tab = {}  # template_name -> tab_frame (for focus-if-open behavior)
        self._tab_tooltip_win = None
        self.bundle = []
        self.bundle_name = None
        self.view_mode = "tabs"
        self.list_items = _dedup_list_items(self.meta.get("fetched_issues", []))
        self.list_frame = None
        self.list_tree = None
        self.kanban_frame = None
        self._checked_tickets = set()
        self._session_refreshed_keys = set()
        self.list_search_var = None
        self._user_cache = dict(self.meta.get("user_cache", {}))
        self._toplevels = set()
        self._style = ttk.Style()
        try:
            self._orig_theme = self._style.theme_use()
        except Exception:
            self._orig_theme = None
        # Apply dark-mode styles BEFORE building widgets so nothing
        # flashes with default white/grey/black backgrounds on first paint.
        try:
            self.apply_dark_mode()
        except Exception:
            pass
        # Build UI (all widgets inherit the already-configured styles)
        self._build_ui()
        # Second theme pass to catch tk.* widgets the recursive theme missed
        try:
            self._apply_recursive_theme(self, dark=True)
        except Exception:
            pass
        # Flush all pending geometry so the window is fully painted before
        # any background work starts creating more widgets.
        self.update_idletasks()
        # Clean up leftover files from a previous update
        self._cleanup_old_update()
        if self.meta.get("first_run_done"):
            self._open_startup_log()
            n = len(self.list_items)
            self._log_startup(f"Loaded {n} ticket(s) from storage.")
            self._log_startup(f"Loaded {len(self.templates)} template(s).")
        self._begin_startup()

    def _cleanup_old_update(self):
        """Remove leftover .old files and stale PyInstaller _MEI* temp dirs."""
        import shutil, time
        try:
            if getattr(sys, "frozen", False):
                exe = sys.executable
            else:
                exe = os.path.abspath(sys.argv[0]) if sys.argv else ""
            if not exe:
                return
            d = os.path.dirname(exe)
            base, ext = os.path.splitext(os.path.basename(exe))
            old = os.path.join(d, f"{base}.old{ext}")
            if os.path.exists(old):
                try:
                    os.remove(old)
                    debug_log(f"Cleaned up old update file: {old}")
                except Exception:
                    pass
        except Exception:
            pass

        # Clean up stale _MEI* temp directories from previous PyInstaller runs
        if getattr(sys, "frozen", False):
            try:
                tmp = os.path.dirname(sys._MEIPASS)
                current_mei = os.path.basename(sys._MEIPASS)
                now = time.time()
                for name in os.listdir(tmp):
                    if not name.startswith("_MEI") or name == current_mei:
                        continue
                    p = os.path.join(tmp, name)
                    if not os.path.isdir(p):
                        continue
                    try:
                        age = now - os.path.getctime(p)
                        if age > 30:
                            shutil.rmtree(p, ignore_errors=True)
                            debug_log(f"Cleaned stale temp dir: {p}")
                    except Exception:
                        pass
            except Exception:
                pass

    def _build_welcome_tab(self):
        """Create the Welcome landing pad tab."""
        self._welcome_frame = ttk.Frame(self.notebook, padding=24)
        self.notebook.add(self._welcome_frame, text="Welcome")
        try:
            self.notebook.tab(self._welcome_frame, style="NotebookNoClose.Tab")
        except tk.TclError:
            pass
        self.notebook.select(self._welcome_frame)
        ttk.Label(self._welcome_frame, text="Welcome to Avalanche Jira Template Creator", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 8))
        self._welcome_text = tk.Text(self._welcome_frame, height=6, wrap="word", bg="#1e1e1e", fg="#dcdcdc", font=("Segoe UI", 10), state="disabled", cursor="arrow")
        self._welcome_text.pack(fill="x", pady=(0, 8))
        self._welcome_new_frame = tk.LabelFrame(self._welcome_frame, text="New / Updated Tickets — click to open", bg="#2b2b2b", fg="#dcdcdc", highlightbackground="#3c3c3c", highlightcolor="#4a4a4a")
        self._welcome_new_frame.pack(fill="both", expand=True, pady=(0, 8))
        self._welcome_new_list = tk.Listbox(self._welcome_new_frame, height=6, selectmode="single", font=("Segoe UI", 10), bg="#1e1e1e", fg="#dcdcdc", selectbackground="#3d5a80", selectforeground="#ffffff", highlightthickness=0)
        self._welcome_new_list.pack(fill="both", expand=True, padx=4, pady=4)
        _bind_mousewheel(self._welcome_new_list, "vertical")
        self._welcome_new_list.bind("<Double-1>", self._on_welcome_new_double_click)
        self._welcome_high_frame = tk.LabelFrame(self._welcome_frame, text="High Internal Priority", bg="#2b2b2b", fg="#dcdcdc", highlightbackground="#3c3c3c", highlightcolor="#4a4a4a")
        self._welcome_high_frame.pack(fill="both", expand=True, pady=(0, 8))
        self._welcome_high_var = tk.BooleanVar(value=self.meta.get("welcome_show_high_priority", True))
        high_cb = tk.Checkbutton(self._welcome_high_frame, text="Show high-priority tickets on Welcome", variable=self._welcome_high_var, command=self._on_welcome_high_toggle, bg="#2b2b2b", fg="#dcdcdc", selectcolor="#3c3c3c", activebackground="#2b2b2b", activeforeground="#dcdcdc", highlightthickness=0)
        high_cb.pack(anchor="w", padx=4, pady=4)
        self._welcome_high_list = tk.Listbox(self._welcome_high_frame, height=4, selectmode="single", font=("Segoe UI", 10), bg="#1e1e1e", fg="#dcdcdc", selectbackground="#3d5a80", selectforeground="#ffffff", highlightthickness=0)
        self._welcome_high_list.pack(fill="both", expand=True, padx=4, pady=4)
        _bind_mousewheel(self._welcome_high_list, "vertical")
        self._welcome_high_list.bind("<Double-1>", self._on_welcome_high_double_click)
        # Blocked issues section
        self._welcome_blocked_frame = tk.LabelFrame(
            self._welcome_frame, text="Blocked Issues",
            bg="#2b2b2b", fg="#dcdcdc",
            highlightbackground="#3c3c3c", highlightcolor="#4a4a4a")
        self._welcome_blocked_frame.pack(fill="both", expand=True, pady=(0, 8))
        self._welcome_blocked_list = tk.Listbox(
            self._welcome_blocked_frame, height=4, selectmode="single",
            font=("Segoe UI", 10), bg="#1e1e1e", fg="#dcdcdc",
            selectbackground="#3d5a80", selectforeground="#ffffff",
            highlightthickness=0)
        self._welcome_blocked_list.pack(fill="both", expand=True, padx=4, pady=4)
        _bind_mousewheel(self._welcome_blocked_list, "vertical")
        self._welcome_blocked_list.bind("<Double-1>", self._on_welcome_blocked_double_click)
        self._welcome_blocked_keys: list = []
        btn_frame = ttk.Frame(self._welcome_frame)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Fetch My Issues", command=self.fetch_my_issues_dialog).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frame, text="Refresh All Tickets", command=self.refresh_fetched_tickets).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frame, text="Jira", command=self._open_welcome_selected_in_jira).pack(side="left", padx=(0, 8))
        self._update_welcome_text()

    def _open_startup_log(self):
        """Create the startup/reminders log window."""
        win = tk.Toplevel(self)
        win.title("Avalanche")
        w, h = 460, 400
        win.update_idletasks()
        sx = win.winfo_screenwidth()
        sy = win.winfo_screenheight()
        win.geometry(f"{w}x{h}+{(sx - w) // 2}+{(sy - h) // 2}")
        win.resizable(True, True)
        win.configure(bg="#2b2b2b")
        win.attributes("-topmost", True)
        try:
            win.transient(self)
        except Exception:
            pass
        self._startup_log_win = win
        title_frame = tk.Frame(win, bg="#2b2b2b")
        title_frame.pack(side="top", fill="x", padx=12, pady=(10, 4))
        sz = 28
        self._startup_spinner_canvas = tk.Canvas(
            title_frame, width=sz, height=sz, bg="#2b2b2b",
            highlightthickness=0)
        self._startup_spinner_canvas.pack(side="left", padx=(0, 8))
        self._startup_spinner_angle = 0
        self._startup_spinner_active = True
        self._startup_log_title = tk.Label(
            title_frame, text="Starting up...", font=("Segoe UI", 12, "bold"),
            bg="#2b2b2b", fg="#dcdcdc", anchor="w")
        self._startup_log_title.pack(side="left", fill="x", expand=True)
        self._tick_startup_spinner()
        self.after(10000, self._stop_startup_spinner)
        self._startup_dismiss_btn = tk.Button(
            win, text="Dismiss", bg="#3d5a80", fg="#dcdcdc",
            font=("Segoe UI", 10), relief="flat", padx=16, pady=4,
            command=self._close_startup_log)
        self._startup_dismiss_btn.pack(side="bottom", pady=(0, 10))
        self._startup_log_text = tk.Text(
            win, wrap="word", bg="#1e1e1e", fg="#dcdcdc",
            font=("Segoe UI", 10), state="disabled", height=4,
            highlightthickness=0, borderwidth=0, padx=8, pady=6)
        self._startup_log_text.pack(side="top", fill="both", expand=True, padx=12, pady=(0, 6))
        self._startup_log_text.tag_configure("step", foreground="#4ec9b0")
        self._startup_log_text.tag_configure("done", foreground="#888888")
        self._startup_log_text.tag_configure("reminder", foreground="#e0c872")

    def _tick_startup_spinner(self):
        """Draw a rotating arc on the canvas."""
        if not getattr(self, "_startup_spinner_active", False):
            return
        try:
            win = getattr(self, "_startup_log_win", None)
            if not win or not win.winfo_exists():
                return
            c = self._startup_spinner_canvas
            c.delete("all")
            pad = 3
            c.create_arc(pad, pad, 28 - pad, 28 - pad,
                         start=self._startup_spinner_angle, extent=270,
                         outline="#4ec9b0", width=3, style="arc")
            self._startup_spinner_angle = (self._startup_spinner_angle + 60) % 360
            c.update_idletasks()
        except Exception:
            pass
        self.after(50, self._tick_startup_spinner)

    def _stop_startup_spinner(self):
        """Stop the spinner and show a checkmark."""
        self._startup_spinner_active = False
        try:
            c = self._startup_spinner_canvas
            c.delete("all")
            c.create_text(14, 14, text="\u2714", fill="#4ec9b0",
                          font=("Segoe UI", 16, "bold"))
        except Exception:
            pass

    def _log_startup(self, msg, tag="step"):
        """Append a line to the startup log window."""
        try:
            win = getattr(self, "_startup_log_win", None)
            if not win or not win.winfo_exists():
                return
            tx = self._startup_log_text
            tx.config(state="normal")
            tx.insert("end", msg + "\n", tag)
            tx.see("end")
            tx.config(state="disabled")
            self.update_idletasks()
        except Exception:
            pass

    def _set_startup_title(self, title):
        """Update the startup log window title label."""
        try:
            win = getattr(self, "_startup_log_win", None)
            if not win or not win.winfo_exists():
                return
            self._startup_log_title.config(text=title)
        except Exception:
            pass

    def _close_startup_log(self):
        """Dismiss the startup log window."""
        self._startup_spinner_active = False
        try:
            win = getattr(self, "_startup_log_win", None)
            if win and win.winfo_exists():
                win.destroy()
            self._startup_log_win = None
        except Exception:
            pass


    def _open_welcome_selected_in_jira(self):
        """Open selected ticket from Welcome new or high list in Jira."""
        keys = []
        for lst, key_attr in [
            (self._welcome_new_list, "_welcome_new_keys"),
            (self._welcome_high_list, "_welcome_high_keys"),
            (self._welcome_blocked_list, "_welcome_blocked_keys"),
        ]:
            sel = lst.curselection()
            if sel:
                idx = int(sel[0])
                klist = getattr(self, key_attr, [])
                if 0 <= idx < len(klist):
                    keys.append(klist[idx])
                break
        if not keys:
            messagebox.showinfo("Info", "Select a ticket from one of the lists.")
            return
        for k in keys:
            self._open_in_jira_browser(k)

    def _on_welcome_new_double_click(self, event):
        """Open ticket from new/updated list."""
        try:
            sel = self._welcome_new_list.curselection()
            if not sel:
                return
            idx = int(sel[0])
            keys = getattr(self, "_welcome_new_keys", [])
            if 0 <= idx < len(keys):
                self._open_ticket_by_key(keys[idx])
        except Exception:
            pass

    def _on_welcome_high_double_click(self, event):
        """Open ticket from high-priority list."""
        try:
            sel = self._welcome_high_list.curselection()
            if not sel:
                return
            idx = int(sel[0])
            keys = getattr(self, "_welcome_high_keys", [])
            if 0 <= idx < len(keys):
                self._open_ticket_by_key(keys[idx])
        except Exception:
            pass

    def _on_welcome_blocked_double_click(self, event):
        """Open ticket from blocked issues list."""
        try:
            sel = self._welcome_blocked_list.curselection()
            if not sel:
                return
            idx = int(sel[0])
            keys = getattr(self, "_welcome_blocked_keys", [])
            if 0 <= idx < len(keys):
                self._open_ticket_by_key(keys[idx])
        except Exception:
            pass

    def _on_welcome_high_toggle(self):
        """Toggle visibility of high-priority section."""
        self.meta["welcome_show_high_priority"] = self._welcome_high_var.get()
        save_storage(self.templates, self.meta)
        self._update_welcome_lists()

    def _open_ticket_by_key(self, key):
        """Open a ticket tab by issue key/id. Switches to tabs view and focuses the ticket (or existing tab)."""
        item = next((it for it in self.list_items if (it.get("Issue key") or it.get("Issue id")) == key), None)
        if item:
            self.show_tabs_view()
            self.new_tab(initial_data=item)

    def _update_welcome_text(self, updates=None):
        """Update the welcome tab with current status and updates."""
        try:
            if not hasattr(self, "_welcome_text"):
                return
            self._welcome_text.config(state="normal")
            self._welcome_text.delete("1.0", "end")
            lines = []
            n = len(self.list_items)
            lines.append(f"• {n} ticket(s) in your list")
            u = updates or self.meta.get("welcome_updates", {})
            if u.get("refreshed", 0) or u.get("new", 0):
                lines.append(f"• Last sync: {u.get('refreshed', 0)} refreshed, {u.get('new', 0)} new")
            if u.get("sync_status"):
                lines.append(f"• {u['sync_status']}")
            lines.append("")
            lines.append("Use 'Fetch My Issues' to download tickets from Jira.")
            lines.append("Use 'Refresh All Tickets' to update existing tickets.")
            self._welcome_text.insert("1.0", "\n".join(lines))
            self._welcome_text.config(state="disabled")
            self._update_welcome_lists()
        except Exception:
            pass

    def _update_welcome_lists(self):
        """Populate new tickets and high-priority lists on Welcome tab."""
        try:
            if not hasattr(self, "_welcome_new_list"):
                return
            self._welcome_new_list.delete(0, tk.END)
            self._welcome_high_list.delete(0, tk.END)
            self._welcome_new_keys = []
            self._welcome_high_keys = []
            new_keys = list(self.meta.get("welcome_updates", {}).get("new_ticket_keys", []))[:50]
            for key in new_keys:
                item = next((it for it in self.list_items if (it.get("Issue key") or it.get("Issue id")) == key), None)
                if item:
                    s = (item.get("Issue key") or "") + " — " + ((item.get("Summary") or "")[:60])
                    self._welcome_new_list.insert(tk.END, s)
                    self._welcome_new_keys.append(key)
            high_levels = self._get_high_priority_levels()
            if self.meta.get("welcome_show_high_priority", True) and high_levels:
                for it in self.list_items:
                    ip = (it.get("Internal Priority") or "").strip()
                    if ip in high_levels:
                        key = it.get("Issue key") or it.get("Issue id")
                        if key:
                            s = (key or "") + " — " + ((it.get("Summary") or "")[:50])
                            self._welcome_high_list.insert(tk.END, s)
                            self._welcome_high_keys.append(key)
            self._welcome_high_frame.pack(fill="both", expand=True, pady=(0, 8)) if self.meta.get("welcome_show_high_priority", True) else self._welcome_high_frame.pack_forget()
            # Blocked issues — tickets that have an inward "is blocked by" link
            self._welcome_blocked_list.delete(0, tk.END)
            self._welcome_blocked_keys = []
            for it in self.list_items:
                links_raw = it.get("Issue Links") or ""
                if not links_raw:
                    continue
                try:
                    links = json.loads(links_raw) if isinstance(links_raw, str) else links_raw
                except Exception:
                    continue
                is_blocked = False
                blocker_keys = []
                for lnk in (links if isinstance(links, list) else []):
                    dl = (lnk.get("direction_label") or "").lower()
                    if "is blocked by" in dl:
                        is_blocked = True
                        bk = lnk.get("key", "")
                        if bk:
                            blocker_keys.append(bk)
                if is_blocked:
                    key = it.get("Issue key") or it.get("Issue id")
                    if key:
                        status = (it.get("Status") or "").strip().lower()
                        if status in ("done", "closed", "resolved"):
                            continue
                        blockers = ", ".join(blocker_keys[:3])
                        s = f"{key} — {(it.get('Summary') or '')[:45]}  ⛔ blocked by {blockers}"
                        self._welcome_blocked_list.insert(tk.END, s)
                        self._welcome_blocked_keys.append(key)
            if self._welcome_blocked_keys:
                self._welcome_blocked_frame.pack(fill="both", expand=True, pady=(0, 8))
            else:
                self._welcome_blocked_frame.pack_forget()
        except Exception:
            pass

    def _get_high_priority_levels(self):
        """Return set of internal priority values considered 'high' (first level or configured)."""
        levels = self.meta.get("internal_priority_levels", ["High", "Medium", "Low", "None"])
        opts = self.meta.get("internal_priority_options", {})
        if levels and levels[0] != "None":
            base = levels[0]
            return set(opts.get(base, [base])) if opts.get(base) else {base}
        return set()

    def _restore_open_tickets(self):
        """Restore tabs that were open when the app was last closed."""
        self._log_startup("Restoring open tickets...")
        keys = self.meta.get("open_ticket_keys", [])
        if not keys or not self.list_items:
            self.notebook.select(self._welcome_frame)
            return
        self._restore_queue = list(keys[:20])
        self._restore_opened = 0
        self._restore_next_ticket()

    def _restore_next_ticket(self):
        """Open one queued ticket tab, then yield to the event loop for the next."""
        if not self._restore_queue:
            self.after(0, lambda: self.notebook.select(self._welcome_frame))
            if self._restore_opened and getattr(self, "_reminder_startup_after_id", None):
                try:
                    self.after_cancel(self._reminder_startup_after_id)
                    self._reminder_startup_after_id = None
                except Exception:
                    pass
            return
        key = self._restore_queue.pop(0)
        item = next((it for it in self.list_items if (it.get("Issue key") or it.get("Issue id")) == key), None)
        if item:
            try:
                self.new_tab(initial_data=item, select_tab=False)
                self._restore_opened += 1
            except Exception:
                pass
        self.after(80, self._restore_next_ticket)

    def _on_app_close(self):
        """Save open ticket keys and close the app."""
        keys = []
        welcome = getattr(self, "_welcome_frame", None)
        for tab_id in self.notebook.tabs():
            try:
                w = self.nametowidget(tab_id)
                if w is welcome:
                    continue
                tf = self.tabs.get(w)
                if tf:
                    data = tf.read_to_dict()
                    k = data.get("Issue key") or data.get("Issue id")
                    if k:
                        keys.append(k)
            except Exception:
                pass
        self.meta["open_ticket_keys"] = keys
        save_storage(self.templates, self.meta)
        self.destroy()

    def _ensure_field_options(self, session):
        """Refresh ALL Jira field options on every startup so dropdowns
        always reflect the latest values in Jira."""
        self._log_startup("Syncing Jira field options...", "step")

        def worker():
            fetched = {}
            pk = "SUNDANCE"
            try:
                vals = self._fetch_projects(session)
                if vals:
                    fetched["Project key"] = vals
                    if pk not in vals and vals:
                        pk = vals[0]

                for name, fn in [
                    ("Priority",    lambda: self._fetch_priorities(session)),
                    ("Labels",      lambda: self._fetch_labels(session)),
                    ("Issue Type",  lambda: self._fetch_issue_types(session, pk)),
                    ("Status",      lambda: self._fetch_statuses(session, pk)),
                    ("Components",  lambda: self._fetch_components(session, pk)),
                    ("Sprint",      lambda: self._fetch_sprints(session, pk)),
                    ("Fix Version", lambda: self._fetch_versions(session, pk)),
                ]:
                    try:
                        v = fn()
                        if v:
                            fetched[name] = v
                    except Exception:
                        debug_log(f"_ensure_field_options [{name}]: "
                                  + traceback.format_exc())

                # Users (Assignee / Reporter)
                try:
                    users = self._fetch_assignable_users(session, pk)
                    if users:
                        fetched["Assignee"] = users
                        fetched["Reporter"] = users
                except Exception:
                    debug_log("_ensure_field_options [users]: "
                              + traceback.format_exc())
            except Exception:
                debug_log("_ensure_field_options failed: " + traceback.format_exc())

            def apply():
                opts = self.meta.setdefault("options", {})
                for name, vals in fetched.items():
                    existing = set(opts.get(name) or [])
                    merged = sorted(existing | set(vals), key=lambda x: x.lower())
                    opts[name] = merged
                if fetched:
                    save_storage(self.templates, self.meta)
                    self._log_startup(
                        f"Updated: {', '.join(fetched.keys())}.", "step")
            self.after(0, apply)
        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _startup_sync(self):
        """First-run prompt, then background fetch new + refresh non-done tickets."""
        self._log_startup("Checking Jira connection...")
        if not self.meta.get("first_run_done") and not self.list_items:
            self._update_welcome_text({"sync_status": "No tickets yet. Use Fetch My Issues to get started."})
            return
        if not self.list_items:
            self._update_welcome_text({"sync_status": "No tickets yet. Use Fetch My Issues to get started."})
            self._log_startup("No tickets to sync.", "done")
            self.after(200, self._run_auto_fetch)
            return
        s = self.get_jira_session()
        if not s:
            self._update_welcome_text({"sync_status": "Set Jira API credentials to sync."})
            self._log_startup("No Jira credentials — skipping sync.", "done")
            return
        # Ensure all field options are populated (Status, Components, etc.)
        self._ensure_field_options(s)
        self._log_startup("Syncing tickets with Jira...")
        self._update_welcome_text({"sync_status": "Syncing in background..."})
        def worker():
            updates = {"refreshed": 0, "new": 0, "failed": 0}
            try:
                done_statuses = ("done", "closed", "resolved", "complete", "completed", "cancelled")
                to_refresh = []
                for idx, it in enumerate(self.list_items):
                    status = (it.get("Status") or "").strip().lower()
                    status_cat = (it.get("Status Category") or "").strip().lower()
                    if status_cat == "done" or status in done_statuses:
                        continue
                    key = it.get("Issue key") or it.get("Issue id")
                    if key and not str(key).startswith("LOCAL-"):
                        to_refresh.append((idx, key))
                consecutive_404 = 0
                for i, (idx, key) in enumerate(to_refresh[:15]):
                    if consecutive_404 >= 3:
                        debug_log(f"Background sync: aborting after {consecutive_404} consecutive failures")
                        break
                    try:
                        issue_json = self.fetch_issue_details(s, key, fields=FETCH_FIELDS)
                        consecutive_404 = 0
                    except Exception as e:
                        if "404" in str(e) or "not found" in str(e).lower():
                            consecutive_404 += 1
                        updates["failed"] += 1
                        continue
                    issue_dict = self._map_issue_json_to_dict(issue_json)
                    self._enrich_with_internal_priority(issue_dict)
                    self.list_items[idx] = issue_dict
                    updates["refreshed"] += 1
                existing_keys = {str(it.get("Issue key") or "").strip() for it in self.list_items if it.get("Issue key")}
                jql = 'project = "SUNDANCE" AND (assignee = currentUser() OR reporter = currentUser()) ORDER BY created DESC'
                try:
                    sr = self.jira_search_jql_simple(s, jql=jql, max_results=20, exclude_keys=existing_keys)
                    for entry in (sr.get("issues") or [])[:10]:
                        issue_id = entry.get("id") or entry.get("key")
                        if not issue_id:
                            continue
                        try:
                            issue_json = self.fetch_issue_details(s, issue_id, fields=FETCH_FIELDS)
                        except Exception:
                            continue
                        issue_dict = self._map_issue_json_to_dict(issue_json)
                        self._enrich_with_internal_priority(issue_dict)
                        self.list_items.append(issue_dict)
                        updates["new"] += 1
                        updates.setdefault("new_ticket_keys", []).append(issue_dict.get("Issue key") or issue_dict.get("Issue id"))
                except Exception:
                    pass
                self.list_items = _dedup_list_items(self.list_items)
                self.meta["fetched_issues"] = list(self.list_items)
                self.meta["welcome_updates"] = updates
                save_storage(self.templates, self.meta)
            except Exception:
                updates["sync_status"] = "Sync failed."
                self.meta["welcome_updates"] = updates
            def done():
                self._populate_listview()
                status = f"{updates.get('refreshed', 0)} refreshed, {updates.get('new', 0)} new"
                if updates.get("failed"):
                    status += f", {updates['failed']} failed"
                self._update_welcome_text({**updates, "sync_status": status})
                self._log_startup(f"Sync complete: {status}.", "done")
                # Phase 2: run auto-fetch inside the same startup window
                self.after(200, self._run_auto_fetch)
            self.after(0, done)
        threading.Thread(target=worker, daemon=True).start()

    def run(self):
        self.mainloop()
