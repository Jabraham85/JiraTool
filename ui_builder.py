"""
UI construction and theme mixin for Avalanche.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from utils import _bind_mousewheel, _bind_mousewheel_to_target, _NotebookWithCloseTabs, debug_log
from storage import save_storage
from config import HEADERS


class UIBuilderMixin:
    def _apply_initial_geometry(self):
        """Open main window fullscreen (maximized) on first show (runs once)."""
        if getattr(self, "_initial_geometry_applied", False):
            return
        try:
            self.update_idletasks()
            try:
                self.state("zoomed")
            except Exception:
                sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
                if sw > 100 and sh > 100:
                    self.geometry(f"{sw}x{sh}+0+0")
            self._initial_geometry_applied = True
        except Exception:
            pass

    def _register_toplevel(self, win):
        try:
            self._toplevels.add(win)
            win.bind("<Destroy>", lambda e, w=win: self._toplevels.discard(w))
            # Defer theme so it runs after caller finishes adding widgets
            def _apply():
                try:
                    win.configure(bg="#2b2b2b")
                except Exception:
                    pass
                try:
                    self._apply_recursive_theme(win, dark=True)
                except Exception:
                    pass
            win.after_idle(_apply)
        except Exception:
            pass

    def _build_ui(self):
        paned = ttk.Panedwindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)
        # LEFT - scrollable panel
        left_outer = ttk.Frame(paned, width=320)
        paned.add(left_outer, weight=1)
        left_canvas = tk.Canvas(left_outer, highlightthickness=0, bg="#2b2b2b")
        left_sb = ttk.Scrollbar(left_outer, orient="vertical", command=left_canvas.yview)
        left_canvas.pack(side="left", fill="both", expand=True)
        left_sb.pack(side="right", fill="y")
        left_canvas.configure(yscrollcommand=left_sb.set)
        left = ttk.Frame(left_canvas)
        _left_win_id = left_canvas.create_window((0, 0), window=left, anchor="nw")
        def _left_configure(e):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        def _left_canvas_configure(e):
            left_canvas.itemconfig(_left_win_id, width=max(e.width, 1))
        left.bind("<Configure>", _left_configure)
        left_canvas.bind("<Configure>", _left_canvas_configure)
        _bind_mousewheel(left_canvas, "vertical")
        left_canvas.after_idle(lambda: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        # Templates header
        ttk.Label(left, text="Templates", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(12, 0))
        tpl_frame = ttk.Frame(left)
        self._tut_tpl_frame = tpl_frame
        tpl_frame.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        tpl_sb = ttk.Scrollbar(tpl_frame, orient="vertical")
        self.template_list = tk.Listbox(tpl_frame, height=10, yscrollcommand=tpl_sb.set)
        self.template_list.pack(side="left", fill="both", expand=True)
        tpl_sb.pack(side="right", fill="y")
        tpl_sb.config(command=self.template_list.yview)
        _bind_mousewheel(self.template_list, "vertical")
        self.template_list.bind("<<ListboxSelect>>", lambda e: self.on_template_select())
        self.template_list.bind("<Double-1>", lambda e: self.on_template_select())
        self.refresh_templates()
        btns = ttk.Frame(left)
        btns.pack(fill="x", padx=8, pady=4)
        ttk.Button(btns, text="New", command=self.new_template).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btns, text="Duplicate", command=self.duplicate_template).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btns, text="Delete", command=self.delete_template).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Separator(left, orient="horizontal").pack(fill="x", padx=8, pady=6)
        self._tut_save_template_btn = ttk.Button(left, text="Save Template", command=self.save_template_with_prompt)
        self._tut_save_template_btn.pack(fill="x", padx=8, pady=2)
        ttk.Button(left, text="Save All", command=self.save_all).pack(fill="x", padx=8, pady=2)
        ttk.Button(left, text="Import Template from CSV Row(s)", command=self.import_from_csv_rows_open_tabs).pack(fill="x", padx=8, pady=2)
        chk = ttk.Frame(left)
        chk.pack(fill="x", padx=8, pady=4)
        ttk.Button(chk, text="Check All", command=self.check_all_current_tab).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(chk, text="Uncheck All", command=self.uncheck_all_current_tab).pack(side="left", fill="x", expand=True, padx=2)
        self.collapse_btn = ttk.Button(left, text="Collapse Unincluded Fields", command=self.toggle_collapse_current_tab)
        self.collapse_btn.pack(fill="x", padx=8, pady=4)
        ttk.Separator(left, orient="horizontal").pack(fill="x", padx=8, pady=6)
        ttk.Button(left, text="Import CSV rows into current tab", command=self.import_rows_into_current_tab).pack(fill="x", padx=8, pady=2)
        # Jira
        ttk.Separator(left, orient="horizontal").pack(fill="x", padx=8, pady=8)
        ttk.Label(left, text="Jira", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(4, 0))
        jira_frame = ttk.Frame(left)
        jira_frame.pack(fill="x", padx=8, pady=6)
        self._tut_set_jira_btn = ttk.Button(jira_frame, text="Set Jira API...", command=self.set_jira_credentials)
        self._tut_set_jira_btn.pack(fill="x", pady=2)
        ttk.Button(jira_frame, text="Test Connection", command=self.test_jira_connection).pack(fill="x", pady=2)
        self._tut_fetch_btn = ttk.Button(jira_frame, text="Fetch My Issues...", command=self.fetch_my_issues_dialog)
        self._tut_fetch_btn.pack(fill="x", pady=2)
        ttk.Button(jira_frame, text="Auto-Fetch Settings...", command=self.auto_fetch_settings_dialog).pack(fill="x", pady=2)
        ttk.Button(jira_frame, text="Refresh Fetched Tickets", command=self.refresh_fetched_tickets).pack(fill="x", pady=2)
        self._tut_config_btn = ttk.Button(jira_frame, text="Configure Reminders...", command=self.configure_reminders_dialog)
        self._tut_config_btn.pack(fill="x", pady=2)
        ttk.Separator(jira_frame, orient="horizontal").pack(fill="x", pady=4)
        ttk.Button(jira_frame, text="Export Bundle…",
                   command=self._share_bundle_dialog).pack(fill="x", pady=2)
        ttk.Button(jira_frame, text="Import Bundle…",
                   command=self._import_bundle_dialog).pack(fill="x", pady=2)
        # Help
        ttk.Separator(left, orient="horizontal").pack(fill="x", padx=8, pady=8)
        ttk.Label(left, text="Help", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(4, 0))
        help_frame = ttk.Frame(left)
        help_frame.pack(fill="x", padx=8, pady=6)
        ttk.Button(help_frame, text="Show Tutorial", command=lambda: self._run_tutorial(force=True)).pack(fill="x", pady=2)
        ttk.Button(help_frame, text="Check for Updates",
                   command=lambda: self._check_for_updates(manual=True)).pack(fill="x", pady=2)
        from config import APP_VERSION as _v
        ttk.Label(help_frame, text=f"v{_v}", foreground="#888888",
                  font=("Segoe UI", 8)).pack(anchor="e", pady=(2, 0))
        # Bundle
        ttk.Separator(left, orient="horizontal").pack(fill="x", padx=8, pady=8)
        ttk.Label(left, text="Bundle", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(4, 0))
        bundle_frame = ttk.Frame(left)
        bundle_frame.pack(fill="both", expand=True, padx=8, pady=6)
        bnd_inner = ttk.Frame(bundle_frame)
        bnd_inner.pack(side="left", fill="both", expand=True)
        bnd_sb = ttk.Scrollbar(bnd_inner, orient="vertical")
        self.bundle_listbox = tk.Listbox(bnd_inner, height=8, yscrollcommand=bnd_sb.set)
        self.bundle_listbox.pack(side="left", fill="both", expand=True)
        bnd_sb.pack(side="right", fill="y")
        bnd_sb.config(command=self.bundle_listbox.yview)
        _bind_mousewheel(self.bundle_listbox, "vertical")
        bbtn_frame = ttk.Frame(bundle_frame)
        bbtn_frame.pack(side="right", fill="y", padx=(6, 0))
        self._tut_add_bundle_btn = ttk.Button(bbtn_frame, text="Add ▶", command=self.add_active_tab_to_bundle)
        self._tut_add_bundle_btn.pack(fill="x", pady=2)
        ttk.Button(bbtn_frame, text="Jira",   command=self._open_selected_bundle_in_jira).pack(fill="x", pady=2)
        ttk.Button(bbtn_frame, text="Export", command=self._share_bundle_dialog).pack(fill="x", pady=2)
        ttk.Button(bbtn_frame, text="Remove", command=self.remove_selected_from_bundle).pack(fill="x", pady=2)
        ttk.Button(bbtn_frame, text="Clear", command=self.clear_bundle).pack(fill="x", pady=2)
        ttk.Button(bbtn_frame, text="Rename", command=self.rename_bundle).pack(fill="x", pady=2)
        bundle_export_frame = ttk.Frame(left)
        bundle_export_frame.pack(fill="x", padx=8, pady=6)
        self._tut_upload_btn = ttk.Button(bundle_export_frame, text="Upload Bundle to Jira...", command=self.upload_bundle_to_jira_dialog)
        self._tut_upload_btn.pack(fill="x", pady=6)
        self._tut_bulk_import_btn = ttk.Button(bundle_export_frame, text="Bulk Import...", command=self.bulk_import_dialog)
        self._tut_bulk_import_btn.pack(fill="x", pady=2)
        # RIGHT
        right = ttk.Frame(paned)
        paned.add(right, weight=3)
        top = ttk.Frame(right)
        top.pack(fill="x", padx=12, pady=(12, 6))
        ttk.Label(top, text="Search fields:", font=("Segoe UI", 10)).pack(side="left")
        self.filter_var = tk.StringVar()
        self._filter_after_id = None
        def _debounced_filter(*a):
            if self._filter_after_id:
                self.after_cancel(self._filter_after_id)
            def _run():
                self._filter_after_id = None
                self.update_filter_for_active_tab()
            self._filter_after_id = self.after(120, _run)
        self.filter_var.trace_add("write", _debounced_filter)
        ttk.Entry(top, textvariable=self.filter_var, width=28).pack(side="left", fill="x", expand=True, padx=8)
        nb_controls = ttk.Frame(right)
        nb_controls.pack(fill="x", padx=12, pady=(4, 4))
        ttk.Button(nb_controls, text="New Tab", command=self.new_tab).pack(side="left")
        ttk.Button(nb_controls, text="Close Tab", command=self.close_current_tab).pack(side="left", padx=6)
        ttk.Button(nb_controls, text="Duplicate Tab", command=self.duplicate_current_tab).pack(side="left", padx=6)
        ttk.Button(nb_controls, text="Save", command=self.save_current_tab).pack(side="left", padx=6)
        self._tut_toggle_list_btn = ttk.Button(nb_controls, text="Toggle List View", command=self.toggle_list_view)
        self._tut_toggle_list_btn.pack(side="left", padx=6)
        _var_btn = ttk.Button(nb_controls, text="📌 Variable", command=self.define_variable_dialog)
        _var_btn.bind("<ButtonPress-1>", lambda e: self._snapshot_var_selection())
        _var_btn.pack(side="left", padx=6)
        ttk.Button(nb_controls, text="Refresh from Jira", command=self.refresh_active_tab_from_jira).pack(side="left", padx=6)
        self._list_filter_bar = ttk.Frame(nb_controls)
        ttk.Separator(self._list_filter_bar, orient="vertical").pack(side="left", fill="y", padx=8, pady=2)
        ttk.Label(self._list_filter_bar, text="Filter:").pack(side="left", padx=(0, 4))
        ttk.Button(self._list_filter_bar, text="All", width=10, command=lambda: self._set_list_scope("All")).pack(side="left", padx=2)
        ttk.Button(self._list_filter_bar, text="Assigned to me", width=14, command=lambda: self._set_list_scope("Assigned to me")).pack(side="left", padx=2)
        ttk.Button(self._list_filter_bar, text="Created by me", width=14, command=lambda: self._set_list_scope("Created by me")).pack(side="left", padx=2)
        ttk.Button(self._list_filter_bar, text="Done", width=10, command=lambda: self._set_list_scope("Done")).pack(side="left", padx=2)
        self.notebook = _NotebookWithCloseTabs(right, on_tab_close=self.close_tab_by_frame)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.on_tab_changed())
        # Build list view (hidden)
        self._build_list_view()
        # Welcome tab (landing pad)
        self._build_welcome_tab()
        self.update_bundle_listbox()
        self._reminder_shown_session = set()
        self._reminder_on_open_after_id = None
        self._reminder_startup_after_id = None
        self._last_reminder_popup_time = 0.0
        self.protocol("WM_DELETE_WINDOW", self._on_app_close)

    def _begin_startup(self):
        """Kick off all background startup tasks.  Called from __init__
        *after* the UI is fully built and themed so no widgets flash
        black or grey during the initial paint."""
        if self.meta.get("first_run_done"):
            self._reminder_startup_after_id = self.after(2000, lambda: self._check_reminders("startup"))
        self.after(60000, self._schedule_reminder_tick)
        self.after(100, self._startup_sync)
        self.after(500, self._restore_open_tickets)
        self.after(800, self._maybe_show_tutorial)
        self.after(1500, self._check_for_updates)

    def apply_dark_mode(self):
        s = self._style
        try:
            s.theme_use("clam")
        except Exception:
            pass
        try:
            _NotebookWithCloseTabs._apply_close_layouts(s)
        except Exception:
            pass
        bg = "#2b2b2b"
        fg = "#dcdcdc"
        entry_bg = "#3c3c3c"
        btn_bg = "#3d5a80"
        text_bg = "#1e1e1e"
        tree_bg = "#222222"
        sel_bg = "#555555"
        try:
            self.configure(bg=bg)
        except Exception:
            pass
        try:
            s.configure("TFrame", background=bg)
            s.configure("TLabel", background=bg, foreground=fg)
            s.configure("TButton", background=btn_bg, foreground=fg)
            try:
                s.map("TButton", background=[("active", "#4a6fa5"), ("pressed", "#2d4562")])
            except Exception:
                pass
            s.configure("TEntry", fieldbackground=entry_bg, foreground=fg)
            s.configure("TCombobox", fieldbackground=entry_bg, foreground=fg)
            s.configure("TRadiobutton", background=bg, foreground=fg)
            s.configure("TCheckbutton", background=bg, foreground=fg)
            s.configure("TProgressbar", background=entry_bg, troughcolor=tree_bg)
            s.configure("Treeview", background=tree_bg, foreground=fg, fieldbackground=tree_bg)
            s.map("Treeview", background=[("selected", sel_bg)])
            s.configure("TNotebook", background=bg)
            s.configure("TNotebook.Tab", background=entry_bg, foreground=fg)
            s.map("TNotebook.Tab", background=[("selected", bg)])
            s.configure("TPanedwindow", background=bg)
            s.configure("TSeparator", background=entry_bg)
            s.configure("Vertical.TScrollbar", background=entry_bg,
                         troughcolor=bg, arrowcolor=fg)
            s.configure("Horizontal.TScrollbar", background=entry_bg,
                         troughcolor=bg, arrowcolor=fg)
            s.configure("TLabelframe", background=bg)
            s.configure("TLabelframe.Label", background=bg, foreground=fg)
            s.configure("Link.TLabel", background=bg, foreground="#4a9eff")
        except Exception:
            pass
        for tf in list(self.tabs.values()):
            for hdr in HEADERS:
                info = tf.field_widgets.get(hdr)
                if not info:
                    continue
                w = info["widget"]
                if isinstance(w, tk.Text):
                    try:
                        w.configure(bg=text_bg, fg=fg, insertbackground=fg)
                    except Exception:
                        pass
        for win in list(self._toplevels) + [self]:
            try:
                self._apply_recursive_theme(win, dark=True)
            except Exception:
                pass

    def _apply_recursive_theme(self, widget, dark=False):
        try:
            # Configure Toplevel/Tk root background
            if isinstance(widget, (tk.Tk, tk.Toplevel)):
                try:
                    if dark:
                        widget.configure(bg="#2b2b2b")
                    else:
                        widget.configure(bg=None)
                except Exception:
                    pass
            for child in widget.winfo_children():
                if isinstance(child, tk.Frame):
                    try:
                        if dark:
                            child.configure(bg="#2b2b2b")
                        else:
                            child.configure(bg=None)
                    except Exception:
                        pass
                elif isinstance(child, tk.Canvas):
                    try:
                        if dark:
                            child.configure(bg="#2b2b2b", highlightthickness=0)
                        else:
                            child.configure(bg=None)
                    except Exception:
                        pass
                elif isinstance(child, tk.Label):
                    try:
                        if dark:
                            child.configure(bg="#2b2b2b", fg="#dcdcdc")
                        else:
                            child.configure(bg=None, fg=None)
                    except Exception:
                        pass
                elif isinstance(child, tk.Entry):
                    try:
                        if dark:
                            child.configure(bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc")
                        else:
                            child.configure(bg=None, fg=None, insertbackground=None)
                    except Exception:
                        pass
                elif isinstance(child, tk.Text):
                    try:
                        if dark:
                            child.configure(bg="#1e1e1e", fg="#dcdcdc", insertbackground="#dcdcdc")
                        else:
                            child.configure(bg="white", fg="black", insertbackground="black")
                    except Exception:
                        pass
                elif isinstance(child, tk.LabelFrame):
                    try:
                        if dark:
                            child.configure(bg="#2b2b2b", fg="#dcdcdc", highlightbackground="#3c3c3c", highlightcolor="#4a4a4a")
                        else:
                            child.configure(bg=None, fg=None, highlightbackground=None, highlightcolor=None)
                    except Exception:
                        pass
                elif isinstance(child, tk.Checkbutton):
                    try:
                        if dark:
                            child.configure(bg="#2b2b2b", fg="#dcdcdc", selectcolor="#3c3c3c", activebackground="#2b2b2b", activeforeground="#dcdcdc", highlightthickness=0)
                        else:
                            child.configure(bg=None, fg=None, selectcolor=None, activebackground=None, activeforeground=None, highlightthickness=0)
                    except Exception:
                        pass
                elif isinstance(child, tk.Listbox):
                    try:
                        if dark:
                            child.configure(bg="#1e1e1e", fg="#dcdcdc", selectbackground="#3d5a80", selectforeground="#ffffff")
                        else:
                            child.configure(bg=None, fg=None, selectbackground=None, selectforeground=None)
                    except Exception:
                        pass
                try:
                    self._apply_recursive_theme(child, dark=dark)
                except Exception:
                    pass
        except Exception:
            pass
