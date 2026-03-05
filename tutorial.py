"""
Tutorial system: guided tour with spotlight overlays and step-by-step instructions.
"""
import json
import tkinter as tk
from tkinter import ttk, messagebox

from .utils import debug_log, _bind_mousewheel
from .storage import save_storage


class TutorialMixin:
    """Mixin providing _maybe_show_tutorial and _run_tutorial for the main app."""

    def _maybe_show_tutorial(self):
        """Show tutorial on first startup if enabled."""
        if self.meta.get("first_run_done", False):
            return
        if not self.meta.get("tutorial_enabled", True):
            self.meta["first_run_done"] = True
            save_storage(self.templates, self.meta)
            return
        self._run_tutorial(force=False)

    def _run_tutorial(self, force=False):
        """Run the guided tutorial. Each step spotlights its target widget and auto-advances on completion."""
        if not force and self.meta.get("first_run_done", False):
            return
        self._tutorial_running = True
        _BG = "#1a1a1a"
        _PANEL = "#252526"
        _BORDER = "#3c3c3c"
        _FG = "#d4d4d4"
        _BTN_BG = "#3c3c3c"
        _BTN_HOV = "#505050"
        _BTN_OK = "#0e639c"
        _BTN_OK_HOV = "#1177bb"
        _DO_IT_BG = "#1a6b35"
        _DO_IT_HOV = "#228b44"
        _DEMO_TPL_NAME = "Tutorial Demo Ticket"
        _EXCL_DEMO_TPL_NAME = "Tutorial Bug Report (!)"

        # Step 10 — plain titles only (one per blank-line-separated block)
        _BASIC_IMPORT_EXAMPLE = (
            "Fix login page timeout\n"
            "\n"
            "Update dashboard layout\n"
            "\n"
            "Add CSV export to reports"
        )
        # Step 12 — structured blocks with title + detail lines for ! replacement
        _EXCL_IMPORT_EXAMPLE = (
            "Fix login page timeout\n"
            "The login page hangs after 30 seconds\n"
            "Users see a blank screen instead of an error\n"
            "\n"
            "Update dashboard layout\n"
            "Move the stats panel to the top\n"
            "Add a quick-actions sidebar\n"
            "\n"
            "Add CSV export to reports\n"
            "Support date-range filtering\n"
            "Include column headers in output"
        )

        # (title, instruction_line, detail_body, action_type, get_target_fn)
        STEPS = [
            # 0 — Welcome
            ("Welcome",
             None,
             "Welcome to Avalanche Jira Template Creator!\n\n"
             "This guided tour will walk you through the key features step by step. "
             "Each step will highlight the relevant part of the app and tell you exactly what to do.\n\n"
             "The sidebar on the left has Templates, Jira, and Bundle sections. "
             "The main area on the right shows ticket tabs or a list view.\n\n"
             "Click  Next →  to begin.",
             None, None),

            # 1 — Set Jira API
            ("Step 1 — Connect to Jira",
             "▶  Click  \"Set Jira API...\"  in the Jira section of the sidebar",
             "Before you can fetch or upload tickets you need to provide three things:\n\n"
             "  • Base URL — your Atlassian domain, e.g. https://yourcompany.atlassian.net\n"
             "  • Email — the email address of your Jira account\n"
             "  • API Token — generate one at id.atlassian.com → Security → API tokens\n\n"
             "Click Save to store the credentials, or Cancel to skip for now. "
             "This step will complete as soon as the dialog closes.",
             "set_jira", lambda: getattr(self, "_tut_set_jira_btn", None)),

            # 2 — Create a demo ticket
            ("Step 2 — Your First Ticket",
             "▶  A demo ticket has been opened for you — take a look at the fields",
             "A ticket has been automatically created and opened in a tab on the right.\n\n"
             "Each field in the form maps to a Jira issue field:\n"
             "  • Summary — the ticket title\n"
             "  • Description — rich ADF content (tables, lists, links)\n"
             "  • Issue Type, Priority, Labels, Components, Assignee…\n\n"
             "You can edit any text field directly. Right-click in any field to access "
             "variable options — more on that in Step 5.\n\n"
             "The Description area has its own toolbar with a  ✏ Edit  button that opens "
             "a full rich-text editor popup.\n\n"
             "Click  Next →  when you're ready to save this as a template.",
             "open_demo_ticket", None),

            # 3 — Save as template
            ("Step 3 — Save as a Template",
             "▶  Click  \"Save Template\"  in the sidebar (or the green button below)",
             "Templates are reusable blueprints. After saving, this ticket's fields "
             "become the starting point whenever you create a new ticket from this template.\n\n"
             "A dialog will appear with the name  \"" + _DEMO_TPL_NAME + "\"  "
             "already filled in — just click  OK  to confirm.\n\n"
             "Click the button to open the save dialog.",
             "save_template", lambda: getattr(self, "_tut_save_template_btn", None)),

            # 4 — Use the saved template
            ("Step 4 — Use a Template",
             "▶  Click the green button below to open a ticket from your demo template",
             "Your newly saved template now appears in the Templates list on the left.\n\n"
             "When you select a template and click  New Tab, a fresh ticket pre-filled "
             "with all the template's fields will open.\n\n"
             "We'll auto-select your demo template and open it for you now.",
             "use_template", lambda: getattr(self, "_tut_tpl_frame", None)),

            # 5 — Variables (info-only)
            ("Step 5 — Variables",
             None,
             "Variables let you turn any piece of text into a reusable placeholder.\n\n"
             "How it works:\n"
             "  1. Select text in any field (Summary, Description, etc.)\n"
             "  2. Right-click → Define Variable\n"
             "  3. The text is wrapped as  {A=your text}  automatically\n"
             "  4. Right-click → Insert Variable  to place  {A}  anywhere\n\n"
             "Variables are auto-assigned letters: A, B, C, and so on.\n"
             "On upload, {A=value} is stripped to just the value, and {A} references "
             "are replaced with the value.\n\n"
             "Variables work across the whole ticket:\n"
             "  • Define in Summary → instantly available in the description editor\n"
             "  • Define inside the  ✏ Edit  popup → updates the main window on save\n"
             "  • If the editor is already open when you define a variable elsewhere,\n"
             "    it syncs automatically within a couple of seconds\n"
             "  • Click  Jira Preview  in the description toolbar to see the description\n"
             "    exactly as Jira will display it, with all variable values filled in\n\n"
             "Variables persist in templates — define them once, reuse everywhere.\n\n"
             "Click  Next →  to continue.",
             None, None),

            # 6 — Internal Priorities — set one
            ("Step 6 — Set an Internal Priority",
             "▶  Use the  \"Internal:\"  dropdown at the top of the ticket to set a priority",
             "Internal Priorities are a local-only system for organising your work.\n\n"
             "  • Each ticket has an \"Internal:\" dropdown in its info bar\n"
             "  • Default levels: High, Medium, Low, None\n"
             "  • These are separate from Jira's Priority — they never get uploaded\n"
             "  • In List View you can sort and filter by Internal Priority\n\n"
             "Try it now — click the  \"Internal:\"  dropdown (highlighted) "
             "and change it to any value other than \"None\".",
             "set_internal_priority",
             lambda: getattr(self.get_active_tabform(), "_info_bar", None)
                     if self.get_active_tabform() else None),

            # 7 — Configure Reminders
            ("Step 7 — Configure Reminders",
             "▶  Click  \"Configure Reminders...\"  in the Jira section of the sidebar",
             "The Configure Reminders dialog lets you customise the priority system:\n\n"
             "  • Change priority levels (comma-separated, highest first)\n"
             "  • Enable stale-ticket warnings — flag tickets with no Jira updates\n"
             "  • Choose which fields to ignore for stale detection\n"
             "  • Toggle the tutorial on/off for future startups\n\n"
             "Open the dialog, take a look at the options, then Save or Cancel to continue.",
             "config_reminders", lambda: getattr(self, "_tut_config_btn", None)),

            # 8 — Fetch Issues
            ("Step 8 — Fetch Issues from Jira",
             "▶  Click  \"Fetch My Issues...\"  in the Jira section of the sidebar",
             "This dialog lets you pull existing tickets from Jira into the app.\n\n"
             "Options you can set:\n"
             "  • Scope — tickets assigned to you, created by you, or both\n"
             "  • Project key — e.g. SUNDANCE\n"
             "  • Filters — Labels, Components, Issue Type, Status, Priority\n"
             "  • Folder — organise fetched tickets into a named folder\n\n"
             "Click  Fetch  to download, or  Cancel  to close. "
             "This step completes as soon as the dialog closes.",
             "fetch", lambda: getattr(self, "_tut_fetch_btn", None)),

            # 9 — List view
            ("Step 9 — List View",
             "▶  Click  \"Toggle List View\"  in the toolbar above the tabs",
             "Toggle List View switches between two ways of working with tickets:\n\n"
             "  • Tabs view — each ticket opens in its own editable tab\n"
             "  • List view — all tickets shown as a sortable, filterable table\n\n"
             "In list view you can double-click any row to open it in a tab, "
             "use the search bar to filter, and organise tickets into folders.\n\n"
             "Click the button once to switch modes — the step completes immediately.",
             "toggle_list", lambda: getattr(self, "_tut_toggle_list_btn", None)),

            # 10 — Basic Bulk Import
            ("Step 10 — Basic Bulk Import",
             "▶  Click  \"Bulk Import...\"  in the Bundle section (or the green button below)",
             "Bulk Import creates many tickets at once from a template.\n\n"
             "The simplest use: paste one title per block (blank line = new ticket) "
             "and every ticket gets that title with the rest of the template fields inherited.\n\n"
             "We've pre-filled 3 ticket titles for you.\n"
             "Click the button, confirm the template is  \"" + _DEMO_TPL_NAME + "\",  "
             "and hit  Import  — three tickets will be created instantly.\n\n"
             "Once imported you can explore the tabs before clicking  Next →.",
             "bulk_import", lambda: getattr(self, "_tut_bulk_import_btn", None)),

            # 11 — Review imported tickets (info only)
            ("Step 11 — Review Your Imported Tickets",
             None,
             "Three new tickets just appeared as tabs — take a look!\n\n"
             "Each one inherited all the fields from the  \"" + _DEMO_TPL_NAME + "\"  "
             "template (Issue Type, Priority, Labels, Description…) but got its own unique "
             "summary from the text you pasted.\n\n"
             "Things to notice:\n"
             "  • Tabs are labelled with the ticket summary\n"
             "  • Every field is editable — tweak individual tickets before uploading\n"
             "  • The tickets were also added to the Bundle list automatically\n\n"
             "When you're done exploring, click  Next →  to learn about the  !  "
             "placeholder system.",
             None, None),

            # 12 — Bulk Import with ! placeholders
            ("Step 12 — Bulk Import with ! Placeholders",
             "▶  Click  \"Bulk Import...\"  again (or the green button below)",
             "The  !  placeholder makes bulk import really powerful.\n\n"
             "How it works:\n"
             "  • Put  !  anywhere in a template's description\n"
             "  • Each block's first line becomes the ticket title\n"
             "  • The remaining lines replace each  !  in the description\n\n"
             "A template called  \"" + _EXCL_DEMO_TPL_NAME + "\"  has been created "
             "with  !  in its description.  The pre-filled text has 3 blocks — "
             "each block's first line is the title, and the detail lines fill the  !.\n\n"
             "Click the button, make sure the template dropdown shows  \""
             + _EXCL_DEMO_TPL_NAME + "\",  then hit  Import.\n\n"
             "Open any imported ticket and check its Description — "
             "the detail lines will be in place of the  !  marker.",
             "bulk_import_excl", lambda: getattr(self, "_tut_bulk_import_btn", None)),

            # 13 — Add to Bundle
            ("Step 13 — Build a Bundle",
             "▶  Click  \"Add ▶\"  in the Bundle section on the left",
             "A bundle is a collection of tickets you want to upload to Jira together.\n\n"
             "Your demo ticket is now the active tab — we'll add it to the bundle.\n\n"
             "How bundles work:\n"
             "  1. Open one or more tickets in tabs\n"
             "  2. Click  Add ▶  to add the active tab to the bundle list\n"
             "  3. Repeat for any other tickets you want to include\n"
             "  4. Use  Remove  or  Clear  to manage the list\n\n"
             "Click  Add ▶  now (or the green button below) to add the current tab.",
             "add_bundle", lambda: getattr(self, "_tut_add_bundle_btn", None)),

            # 14 — Upload Bundle (demo only)
            ("Step 14 — Upload to Jira (Demo)",
             "▶  Click the green button below to simulate an upload",
             "When you're ready for real work,  Upload Bundle to Jira…  sends every ticket "
             "in the bundle to Jira in one go.\n\n"
             "For each ticket the app will:\n"
             "  • Create a new issue if the ticket has no Issue Key\n"
             "  • Update the existing issue if it already has an Issue Key\n"
             "  • Apply variable substitutions and sanitise ADF content\n\n"
             "For this tutorial we won't actually send anything — clicking the button "
             "below will clear the bundle and move on.",
             "clear_bundle", lambda: getattr(self, "_tut_upload_btn", None)),

            # 15 — Congratulations
            ("🎉  All done!",
             None,
             "You've completed the Avalanche Jira Template Creator tutorial.\n\n"
             "Here's a quick recap:\n\n"
             "  ✓ Connect to Jira with your API credentials\n"
             "  ✓ Create tickets and save reusable templates\n"
             "  ✓ Define variables — select text, right-click, done\n"
             "  ✓ Use Internal Priorities to organise your workflow\n"
             "  ✓ Fetch existing tickets from Jira with filters\n"
             "  ✓ Use List View to browse, sort, and filter tickets\n"
             "  ✓ Bulk Import by title — one paste, many tickets\n"
             "  ✓ Bulk Import with  !  — fill template placeholders in bulk\n"
             "  ✓ Build a bundle and upload tickets in bulk\n\n"
             "You can re-run this tutorial any time from the Help section in the sidebar. "
             "Happy ticketing!",
             None, None),
        ]

        win = tk.Toplevel(self)
        self._register_toplevel(win)
        win.title("Tutorial")
        win.configure(bg=_BG)
        win.minsize(400, 360)
        win.geometry("440x460")
        win.resizable(True, True)
        win.attributes("-topmost", True)

        step_idx = [0]
        _prev_action = [None]
        _OFF = "-500+-500"    # off-screen position for hiding an overlay panel
        _overlays_active = [True]

        def _relift_all():
            """Re-lift overlays and tutorial window to maintain stacking order."""
            if not _overlays_active[0]:
                return
            try:
                if not win.winfo_exists():
                    return
            except Exception:
                return
            for _p in _ov:
                try:
                    if _p.winfo_exists():
                        _p.lift()
                except Exception:
                    pass
            try:
                win.lift()
            except Exception:
                pass

        # ── Persistent overlay panels (created once, geometry updated per step) ──
        def _mk_panel():
            o = tk.Toplevel(self)
            o.overrideredirect(True)
            try:
                o.attributes("-alpha", 0.60)
            except Exception:
                pass
            try:
                o.attributes("-topmost", True)
            except Exception:
                pass
            o.configure(bg="black")
            o.geometry(f"1x1{_OFF}")
            # Absorb clicks AND immediately return focus to the tutorial window
            blocker = tk.Frame(o, bg="black", cursor="arrow")
            blocker.place(relx=0, rely=0, relwidth=1, relheight=1)
            def _absorb(e):
                _relift_all()
                try:
                    win.focus_force()
                except Exception:
                    pass
                return "break"
            for ev in ("<Button-1>", "<Button-2>", "<Button-3>",
                       "<Double-Button-1>", "<ButtonPress>"):
                blocker.bind(ev, _absorb)
            return o

        _ov = [_mk_panel() for _ in range(4)]   # top, bottom, left, right

        _focus_bind_id = self.bind("<FocusIn>", lambda e: _relift_all())

        def _hide_all_panels():
            for o in _ov:
                try:
                    o.geometry(f"1x1{_OFF}")
                except Exception:
                    pass

        def _set_panel(idx, x, y, w, h):
            try:
                if w > 0 and h > 0:
                    _ov[idx].geometry(f"{w}x{h}+{x}+{y}")
                else:
                    _ov[idx].geometry(f"1x1{_OFF}")
            except Exception:
                pass

        def _destroy_overlays():
            for o in _ov:
                try:
                    o.destroy()
                except Exception:
                    pass

        def _update_spotlight(target=None, padding=10):
            """Update the 4 overlay panels to cover only the main app window."""
            try:
                self.update_idletasks()
                ax = self.winfo_rootx()
                ay = self.winfo_rooty()
                aw = self.winfo_width()
                ah = self.winfo_height()
                ax2 = ax + aw
                ay2 = ay + ah
            except Exception:
                return

            if target is None:
                _set_panel(0, ax, ay, aw, ah)
                _set_panel(1, 0, 0, 0, 0)
                _set_panel(2, 0, 0, 0, 0)
                _set_panel(3, 0, 0, 0, 0)
            else:
                try:
                    tx = max(ax, target.winfo_rootx() - padding)
                    ty = max(ay, target.winfo_rooty() - padding)
                    tx2 = min(ax2, target.winfo_rootx() + target.winfo_width() + padding)
                    ty2 = min(ay2, target.winfo_rooty() + target.winfo_height() + padding)
                except Exception:
                    return
                _set_panel(0, ax,  ay,  aw,         ty - ay)       # top
                _set_panel(1, ax,  ty2, aw,         ay2 - ty2)     # bottom
                _set_panel(2, ax,  ty,  tx - ax,    ty2 - ty)      # left
                _set_panel(3, tx2, ty,  ax2 - tx2,  ty2 - ty)      # right

            for _p in _ov:
                try:
                    _p.deiconify()
                    _p.attributes("-topmost", True)
                    _p.lift()
                except Exception:
                    pass
            try:
                win.attributes("-topmost", True)
                win.lift()
            except Exception:
                pass

        def _position_win(target=None):
            """Move tutorial window beside target, or center on app if no target."""
            try:
                win.update_idletasks()
                sw = self.winfo_screenwidth()
                sh = self.winfo_screenheight()
                ww, wh = win.winfo_width(), win.winfo_height()
                ax = self.winfo_rootx()
                ay = self.winfo_rooty()
                aw = self.winfo_width()
                ah = self.winfo_height()
                if target is not None and target.winfo_width() > 1:
                    tx, ty = target.winfo_rootx(), target.winfo_rooty()
                    tw = target.winfo_width()
                    gap = 18
                    x = tx + tw + gap
                    if x + ww > sw - 10:
                        x = tx - ww - gap
                    if x < 10 or x + ww > sw - 10:
                        x = ax + (aw - ww) // 2
                    y = max(10, min(ty - wh // 4, sh - wh - 10))
                else:
                    x = ax + (aw - ww) // 2
                    y = ay + (ah - wh) // 3
                x = max(10, min(x, sw - ww - 10))
                y = max(10, min(y, sh - wh - 10))
                win.geometry(f"{ww}x{wh}+{x}+{y}")
            except Exception:
                pass

        def _refresh_spotlight(idx):
            _, _, _, _, get_target = STEPS[idx]
            target = get_target() if get_target else None
            try:
                self.update_idletasks()
            except Exception:
                pass
            _update_spotlight(target)
            _position_win(target)

        # ── Build window layout ────────────────────────────────────────────
        title_bar = tk.Frame(win, bg=_PANEL, pady=8)
        title_bar.pack(fill="x")
        title_lbl = tk.Label(title_bar, text="", bg=_PANEL, fg=_FG,
                             font=("Segoe UI", 11, "bold"), padx=14)
        title_lbl.pack(side="left")
        step_lbl = tk.Label(title_bar, text="", bg=_PANEL, fg="#888888",
                            font=("Segoe UI", 9), padx=14)
        step_lbl.pack(side="right")
        tk.Frame(win, bg=_BORDER, height=1).pack(fill="x")

        instr_frame = tk.Frame(win, bg="#0e3d1f", pady=8, padx=14)
        instr_lbl = tk.Label(instr_frame, text="", bg="#0e3d1f", fg="#7affa0",
                             font=("Segoe UI", 10, "bold"), wraplength=400,
                             justify="left", anchor="w")
        instr_lbl.pack(fill="x")

        btn_container = tk.Frame(win, bg=_PANEL)
        btn_container.pack(side="bottom", fill="x")

        text_frame = tk.Frame(win, bg=_BG, padx=14, pady=10)
        text_frame.pack(fill="both", expand=True)
        txt = tk.Text(text_frame, wrap="word", bg="#1e1e1e", fg=_FG,
                      font=("Segoe UI", 9), relief="flat", bd=0, padx=10, pady=10,
                      cursor="arrow", highlightthickness=0)
        sb_txt = tk.Scrollbar(text_frame, orient="vertical", command=txt.yview,
                              bg=_PANEL, troughcolor=_BG, width=10)
        txt.pack(side="left", fill="both", expand=True)
        sb_txt.pack(side="right", fill="y")
        txt.configure(yscrollcommand=sb_txt.set)
        _bind_mousewheel(txt, "vertical")

        # ── Button-patching helpers ────────────────────────────────────────
        _BTN_ORIGINALS = {
            "set_jira":          ("_tut_set_jira_btn",       self.set_jira_credentials),
            "fetch":             ("_tut_fetch_btn",           self.fetch_my_issues_dialog),
            "toggle_list":       ("_tut_toggle_list_btn",     self.toggle_list_view),
            "save_template":     ("_tut_save_template_btn",   self.save_template_with_prompt),
            "add_bundle":        ("_tut_add_bundle_btn",      self.add_active_tab_to_bundle),
            "clear_bundle":      ("_tut_upload_btn",          self.upload_bundle_to_jira_dialog),
            "config_reminders":  ("_tut_config_btn",          self.configure_reminders_dialog),
            "bulk_import":       ("_tut_bulk_import_btn",     self.bulk_import_dialog),
            "bulk_import_excl":  ("_tut_bulk_import_btn",     self.bulk_import_dialog),
        }

        def _restore_sidebar_btn():
            prev = _prev_action[0]
            if prev and prev in _BTN_ORIGINALS:
                attr, orig_cmd = _BTN_ORIGINALS[prev]
                try:
                    getattr(self, attr).config(command=orig_cmd)
                except Exception:
                    pass
            _prev_action[0] = None

        def _patch_sidebar_btn(action_type, on_complete):
            if action_type not in _BTN_ORIGINALS:
                return
            attr, _ = _BTN_ORIGINALS[action_type]
            btn = getattr(self, attr, None)
            if btn is None:
                return
            # For dialog-opening actions, lower tutorial first then raise on close
            if action_type == "set_jira":
                def _cmd(): _tut_lower(); self.set_jira_credentials(on_close=lambda: (_tut_raise(), on_complete()))
                btn.config(command=_cmd)
            elif action_type == "fetch":
                def _cmd():
                    _tut_lower()
                    _fetch_started = [False]
                    _orig_start = self._start_fetch_issues
                    def _wrapped_start(*a, **kw):
                        _fetch_started[0] = True
                        return _orig_start(*a, **kw)
                    self._start_fetch_issues = _wrapped_start
                    # Raise tutorial only AFTER the "Fetched X issues" messagebox is dismissed
                    self._post_fetch_callback = lambda: (_tut_raise(), on_complete())
                    def _on_dlg_close():
                        self._start_fetch_issues = _orig_start
                        if not _fetch_started[0]:
                            # User cancelled — no fetch running, raise immediately
                            self._post_fetch_callback = None
                            _tut_raise()
                            on_complete()
                    self.fetch_my_issues_dialog(on_close=_on_dlg_close)
                btn.config(command=_cmd)
            elif action_type == "save_template":
                def _cmd():
                    _tut_lower()
                    dlg = tk.Toplevel(self)
                    dlg.title("Save Template")
                    dlg.resizable(False, False)
                    dlg.attributes("-topmost", True)
                    dlg.geometry("360x140")
                    frm = ttk.Frame(dlg, padding=16)
                    frm.pack(fill="both", expand=True)
                    ttk.Label(frm, text="Template name:").pack(anchor="w")
                    _ne = tk.Entry(frm, font=("Segoe UI", 10))
                    _ne.insert(0, _DEMO_TPL_NAME)
                    _ne.config(state="readonly")
                    _ne.pack(fill="x", pady=(4, 12))
                    def _ok():
                        tf = self.get_active_tabform()
                        if tf:
                            data = tf.read_to_dict()
                            self._strip_identity_fields(data)
                            self.templates[_DEMO_TPL_NAME] = data
                            save_storage(self.templates, self.meta)
                            self.refresh_templates()
                        dlg.destroy()
                        _tut_raise()
                        on_complete()
                    ttk.Button(frm, text="OK", command=_ok).pack()
                    dlg.protocol("WM_DELETE_WINDOW", lambda: (dlg.destroy(), _tut_raise()))
                    dlg.after(50, lambda: dlg.focus_force())
                btn.config(command=_cmd)
            elif action_type == "toggle_list":
                def _tog(): self.toggle_list_view(); on_complete()
                btn.config(command=_tog)
            elif action_type == "add_bundle":
                def _add(): self.add_active_tab_to_bundle(); on_complete()
                btn.config(command=_add)
            elif action_type == "clear_bundle":
                def _clr(): self.clear_bundle(); on_complete()
                btn.config(command=_clr)
            elif action_type == "config_reminders":
                def _cmd(): _tut_lower(); self.configure_reminders_dialog(on_close=lambda: (_tut_raise(), on_complete()))
                btn.config(command=_cmd)
            elif action_type == "bulk_import":
                def _cmd():
                    _tut_lower()
                    self.bulk_import_dialog(
                        on_close=lambda: (_tut_raise(), on_complete()),
                        prefill_text=_BASIC_IMPORT_EXAMPLE,
                        prefill_template=_DEMO_TPL_NAME)
                btn.config(command=_cmd)
            elif action_type == "bulk_import_excl":
                def _cmd():
                    _tut_lower()
                    _ensure_excl_demo_template()
                    self.bulk_import_dialog(
                        on_close=lambda: (_tut_raise(), on_complete()),
                        prefill_text=_EXCL_IMPORT_EXAMPLE,
                        prefill_template=_EXCL_DEMO_TPL_NAME)
                btn.config(command=_cmd)
            _prev_action[0] = action_type

        # ── Demo template ─────────────────────────────────────────────────
        def _ensure_demo_template():
            """Create a demo template with rich ADF content if no templates exist."""
            if self.templates:
                return
            demo_adf = json.dumps({
                "version": 1, "type": "doc",
                "content": [
                    {"type": "paragraph", "content": [
                        {"type": "text", "text": "Welcome to the ticket editor! ",
                         "marks": [{"type": "strong"}]},
                        {"type": "text", "text": "This description was created by the tutorial. "
                         "Each ticket can have rich ADF content — bold text, lists, tables, links and more."}
                    ]},
                    {"type": "bulletList", "content": [
                        {"type": "listItem", "content": [{"type": "paragraph", "content": [
                            {"type": "text", "text": "Use "},
                            {"type": "text", "text": "Templates", "marks": [{"type": "strong"}]},
                            {"type": "text", "text": " in the sidebar as reusable blueprints"}
                        ]}]},
                        {"type": "listItem", "content": [{"type": "paragraph", "content": [
                            {"type": "text", "text": "Use "},
                            {"type": "text", "text": "Variables", "marks": [{"type": "strong"}]},
                            {"type": "text", "text": " like {project} as dynamic placeholders"}
                        ]}]},
                        {"type": "listItem", "content": [{"type": "paragraph", "content": [
                            {"type": "text", "text": "Add to "},
                            {"type": "text", "text": "Bundle", "marks": [{"type": "strong"}]},
                            {"type": "text", "text": " then upload multiple tickets to Jira in one go"}
                        ]}]},
                    ]},
                    {"type": "table",
                     "attrs": {"isNumberColumnEnabled": False, "layout": "default"},
                     "content": [
                        {"type": "tableRow", "content": [
                            {"type": "tableHeader", "attrs": {}, "content": [{"type": "paragraph", "content": [
                                {"type": "text", "text": "Field", "marks": [{"type": "strong"}]}]}]},
                            {"type": "tableHeader", "attrs": {}, "content": [{"type": "paragraph", "content": [
                                {"type": "text", "text": "Example value", "marks": [{"type": "strong"}]}]}]},
                        ]},
                        {"type": "tableRow", "content": [
                            {"type": "tableCell", "attrs": {}, "content": [{"type": "paragraph", "content": [
                                {"type": "text", "text": "Summary"}]}]},
                            {"type": "tableCell", "attrs": {}, "content": [{"type": "paragraph", "content": [
                                {"type": "text", "text": "Tutorial Demo: {project} ticket"}]}]},
                        ]},
                        {"type": "tableRow", "content": [
                            {"type": "tableCell", "attrs": {}, "content": [{"type": "paragraph", "content": [
                                {"type": "text", "text": "Priority"}]}]},
                            {"type": "tableCell", "attrs": {}, "content": [{"type": "paragraph", "content": [
                                {"type": "text", "text": "Medium"}]}]},
                        ]},
                        {"type": "tableRow", "content": [
                            {"type": "tableCell", "attrs": {}, "content": [{"type": "paragraph", "content": [
                                {"type": "text", "text": "Labels"}]}]},
                            {"type": "tableCell", "attrs": {}, "content": [{"type": "paragraph", "content": [
                                {"type": "text", "text": "tutorial; demo"}]}]},
                        ]},
                    ]},
                    {"type": "paragraph", "content": [
                        {"type": "text",
                         "text": "This template is safe to delete once you're familiar with the app.",
                         "marks": [{"type": "em"}]}
                    ]},
                ]
            })
            self.templates[_DEMO_TPL_NAME] = {
                "Summary": "Tutorial Demo: {project} ticket",
                "Description": "This ticket was created by the tutorial. Edit as needed.",
                "Description ADF": demo_adf,
                "Issue Type": "Task",
                "Priority": "Medium",
                "Labels": "tutorial; demo",
            }
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass
            self.refresh_templates()

        def _ensure_excl_demo_template():
            """Create the ! placeholder demo template if it doesn't already exist."""
            if _EXCL_DEMO_TPL_NAME in self.templates:
                return
            excl_adf = json.dumps({
                "version": 1, "type": "doc",
                "content": [
                    {"type": "paragraph", "content": [
                        {"type": "text", "text": "Steps to reproduce:",
                         "marks": [{"type": "strong"}]}
                    ]},
                    {"type": "paragraph", "content": [
                        {"type": "text", "text": "!"}
                    ]},
                    {"type": "paragraph", "content": [
                        {"type": "text",
                         "text": "Expected outcome: issue is resolved.",
                         "marks": [{"type": "em"}]}
                    ]},
                ]
            })
            self.templates[_EXCL_DEMO_TPL_NAME] = {
                "Summary": "Bug Report",
                "Description": "Steps to reproduce:\n!\n\nExpected outcome: issue is resolved.",
                "Description ADF": excl_adf,
                "Issue Type": "Bug",
                "Priority": "Medium",
            }
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass
            self.refresh_templates()

        # ── Per-step extra cleanup (e.g. temporary event bindings) ─────────
        _extra_cleanup = [[]]

        def _run_extra_cleanup():
            for fn in _extra_cleanup[0]:
                try:
                    fn()
                except Exception:
                    pass
            _extra_cleanup[0] = []

        # ── Topmost management ────────────────────────────────────────────
        # While a child dialog is open, withdraw the tutorial and hide
        # overlays so they don't overlap the dialog at all.
        _OFF_SCREEN = "+{x}+{y}".format(x=-3000, y=-3000)

        def _tut_lower():
            """Move tutorial to a screen corner and hide overlays so child dialogs can be used."""
            _overlays_active[0] = False
            try:
                win.attributes("-topmost", False)
                win.update_idletasks()
                sw = win.winfo_screenwidth()
                sh = win.winfo_screenheight()
                ww = max(win.winfo_width(), 440)
                wh = max(win.winfo_height(), 360)
                win.geometry(f"{ww}x{wh}+{sw - ww - 24}+{sh - wh - 60}")
                _hide_all_panels()
            except Exception:
                pass

        def _tut_raise():
            """Restore tutorial after child dialog closes."""
            _overlays_active[0] = True
            try:
                win.attributes("-topmost", True)
            except Exception:
                pass
            self.after(200, lambda: _refresh_spotlight(step_idx[0]))

        def _transition_to(idx):
            """Withdraw tutorial, move overlays off-screen, then show next step at its position."""
            try:
                win.withdraw()
            except Exception:
                pass
            for o in _ov:
                try:
                    o.geometry(f"1x1{_OFF_SCREEN}")
                except Exception:
                    pass
            win.after(120, lambda: _show_step(idx))

        # ── Step / navigation logic ────────────────────────────────────────
        def _step_complete():
            """Called when the step's required action is done.
            Unlocks Next so the user can explore before proceeding."""
            _tut_raise()
            do_it_btn.pack_forget()
            is_last = (step_idx[0] == len(STEPS) - 1)
            _set_next_enabled(True, is_last=is_last)

        def _show_step(idx):
            _run_extra_cleanup()
            _restore_sidebar_btn()
            step_idx[0] = idx
            title, instr, body, action_type, get_target = STEPS[idx]
            title_lbl.config(text=title)
            is_last = (idx == len(STEPS) - 1)
            step_lbl.config(text=f"{idx + 1} / {len(STEPS)}")

            if instr:
                instr_lbl.config(text=instr)
                instr_frame.pack(fill="x", after=title_bar)
            else:
                instr_frame.pack_forget()

            txt.config(state="normal")
            txt.delete("1.0", "end")
            txt.insert("1.0", body)
            txt.config(state="disabled")

            back_btn.config(state="normal" if idx > 0 else "disabled")

            if action_type is None or action_type == "open_demo_ticket":
                _set_next_enabled(True, is_last=is_last)
                do_it_btn.pack_forget()
                if action_type == "open_demo_ticket":
                    # Auto-open demo ticket in a tab so user can see it
                    _ensure_demo_template()
                    try:
                        self.show_tabs_view()
                        tpl = self.templates.get(_DEMO_TPL_NAME) or (
                              next(iter(self.templates.values()), None))
                        if tpl:
                            self.new_tab(initial_data=dict(tpl))
                    except Exception:
                        pass
            else:
                _set_next_enabled(False)
                _patch_sidebar_btn(action_type, _step_complete)

                if action_type == "set_jira":
                    def _do_set_jira():
                        _tut_lower()
                        self.set_jira_credentials(on_close=lambda: (_tut_raise(), _step_complete()))
                    do_it_btn.config(command=_do_set_jira, text='Open "Set Jira API..."')

                elif action_type == "open_demo_ticket":
                    pass  # handled above; Next already enabled

                elif action_type == "use_template":
                    _use_done = [False]

                    def _do_use():
                        if _use_done[0]:
                            return
                        _use_done[0] = True
                        tpl_idx = None
                        for i in range(self.template_list.size()):
                            if self.template_list.get(i) == _DEMO_TPL_NAME:
                                tpl_idx = i
                                break
                        if tpl_idx is None and self.template_list.size() > 0:
                            tpl_idx = 0
                        if tpl_idx is not None:
                            self.template_list.selection_clear(0, "end")
                            self.template_list.selection_set(tpl_idx)
                            self.template_list.see(tpl_idx)
                            self.on_template_select()
                        _step_complete()

                    def _on_list_click_use(e):
                        if _use_done[0]:
                            return
                        win.after(150, _do_use)

                    self.template_list.bind("<<ListboxSelect>>",
                                            lambda e: _on_list_click_use(e), add=True)
                    _extra_cleanup[0].append(lambda: self.template_list.bind(
                        "<<ListboxSelect>>", lambda e: self.on_template_select()))
                    do_it_btn.config(command=_do_use, text="Open Demo Template")

                elif action_type == "save_template":
                    def _do_save_tpl():
                        _tut_lower()
                        dlg = tk.Toplevel(self)
                        dlg.title("Save Template")
                        dlg.resizable(False, False)
                        dlg.attributes("-topmost", True)
                        dlg.geometry("360x140")
                        frm = ttk.Frame(dlg, padding=16)
                        frm.pack(fill="both", expand=True)
                        ttk.Label(frm, text="Template name:").pack(anchor="w")
                        name_entry = tk.Entry(frm, font=("Segoe UI", 10))
                        name_entry.insert(0, _DEMO_TPL_NAME)
                        name_entry.config(state="readonly")
                        name_entry.pack(fill="x", pady=(4, 12))
                        def _ok():
                            tf = self.get_active_tabform()
                            if tf:
                                data = tf.read_to_dict()
                                self._strip_identity_fields(data)
                                self.templates[_DEMO_TPL_NAME] = data
                                save_storage(self.templates, self.meta)
                                self.refresh_templates()
                            dlg.destroy()
                            _tut_raise()
                            _step_complete()
                        ttk.Button(frm, text="OK", command=_ok).pack()
                        dlg.protocol("WM_DELETE_WINDOW", lambda: (dlg.destroy(), _tut_raise()))
                        dlg.after(50, lambda: dlg.focus_force())
                    do_it_btn.config(command=_do_save_tpl, text='Save as "' + _DEMO_TPL_NAME + '"')
                elif action_type == "fetch":
                    def _do_fetch():
                        _tut_lower()
                        _fetch_started = [False]
                        _orig_start = self._start_fetch_issues
                        def _wrapped_start(*a, **kw):
                            _fetch_started[0] = True
                            return _orig_start(*a, **kw)
                        self._start_fetch_issues = _wrapped_start
                        self._post_fetch_callback = lambda: (_tut_raise(), _step_complete())
                        def _on_dlg_close():
                            self._start_fetch_issues = _orig_start
                            if not _fetch_started[0]:
                                self._post_fetch_callback = None
                                _tut_raise()
                                _step_complete()
                        self.fetch_my_issues_dialog(on_close=_on_dlg_close)
                    do_it_btn.config(command=_do_fetch, text='Open "Fetch My Issues..."')
                elif action_type == "toggle_list":
                    def _do_tog():
                        self.toggle_list_view()
                        _step_complete()
                    do_it_btn.config(command=_do_tog, text="Toggle List View")
                elif action_type == "add_bundle":
                    # Switch back to tabs view and focus the demo ticket
                    try:
                        self.show_tabs_view()
                        # Try to select the first tab (demo ticket)
                        tabs = list(self.tabs.keys())
                        if tabs:
                            self.notebook.select(tabs[0])
                    except Exception:
                        pass
                    def _do_add():
                        self.add_active_tab_to_bundle()
                        _step_complete()
                    do_it_btn.config(command=_do_add, text='Click "Add ▶"')
                elif action_type == "clear_bundle":
                    def _do_clr():
                        self.bundle = []
                        self.update_bundle_listbox()
                        _step_complete()
                    do_it_btn.config(command=_do_clr, text="Clear bundle (demo)")
                elif action_type == "set_internal_priority":
                    try:
                        self.show_tabs_view()
                        tabs = list(self.tabs.keys())
                        if tabs:
                            self.notebook.select(tabs[0])
                    except Exception:
                        pass
                    _ip_done = [False]
                    def _on_ip_changed(e=None):
                        if _ip_done[0]:
                            return
                        tf = self.get_active_tabform()
                        if tf:
                            val = tf._info_internal_var.get()
                            if val and val != "None":
                                _ip_done[0] = True
                                _step_complete()
                    _ip_bindings = []
                    for tf in self.tabs.values():
                        tf._info_internal_combo.bind("<<ComboboxSelected>>",
                                                     lambda e: _on_ip_changed(e), add=True)
                        _ip_bindings.append(tf)
                    def _restore_ip():
                        for tf2 in _ip_bindings:
                            try:
                                tf2._info_internal_combo.bind("<<ComboboxSelected>>",
                                                              lambda e, t=tf2: t._on_internal_priority_changed())
                            except Exception:
                                pass
                    _extra_cleanup[0].append(_restore_ip)
                    do_it_btn.pack_forget()
                elif action_type == "config_reminders":
                    def _do_config():
                        _tut_lower()
                        self.configure_reminders_dialog(on_close=lambda: (_tut_raise(), _step_complete()))
                    do_it_btn.config(command=_do_config, text='Open "Configure Reminders..."')
                elif action_type == "bulk_import":
                    def _do_bulk():
                        _tut_lower()
                        self.bulk_import_dialog(
                            on_close=lambda: (_tut_raise(), _step_complete()),
                            prefill_text=_BASIC_IMPORT_EXAMPLE,
                            prefill_template=_DEMO_TPL_NAME)
                    do_it_btn.config(command=_do_bulk, text='Open "Bulk Import..." (basic)')
                elif action_type == "bulk_import_excl":
                    def _do_bulk_excl():
                        _tut_lower()
                        _ensure_excl_demo_template()
                        self.bulk_import_dialog(
                            on_close=lambda: (_tut_raise(), _step_complete()),
                            prefill_text=_EXCL_IMPORT_EXAMPLE,
                            prefill_template=_EXCL_DEMO_TPL_NAME)
                    do_it_btn.config(command=_do_bulk_excl, text='Open "Bulk Import..." (with !)')
                do_it_btn.pack(side="left", padx=(0, 8))

            try:
                win.deiconify()
                self.update_idletasks()
            except Exception:
                pass
            _refresh_spotlight(idx)
            try:
                win.lift()
                win.focus_force()
            except Exception:
                pass

        def _cleanup():
            _overlays_active[0] = False
            try:
                self.unbind("<FocusIn>", _focus_bind_id)
            except Exception:
                pass
            _run_extra_cleanup()
            _destroy_overlays()
            _restore_sidebar_btn()
            self._tutorial_running = False

        def _skip():
            _cleanup()
            self.meta["first_run_done"] = True
            save_storage(self.templates, self.meta)
            win.destroy()

        def _next():
            if step_idx[0] < len(STEPS) - 1:
                _transition_to(step_idx[0] + 1)
            else:
                _cleanup()
                self.meta["first_run_done"] = True
                save_storage(self.templates, self.meta)
                win.destroy()

        def _back():
            if step_idx[0] > 0:
                _transition_to(step_idx[0] - 1)

        tk.Frame(btn_container, bg=_BORDER, height=1).pack(fill="x")
        btn_frame = tk.Frame(btn_container, bg=_PANEL, pady=10, padx=14)
        btn_frame.pack(fill="x")

        def _mkb(parent, text, cmd, bg, hv):
            b = tk.Button(parent, text=text, command=cmd, bg=bg, fg="#ffffff",
                          font=("Segoe UI", 9, "bold"), relief="flat", bd=0,
                          padx=12, pady=6, cursor="hand2",
                          activebackground=hv, activeforeground="#ffffff")
            b.bind("<Enter>", lambda e: b.configure(bg=hv))
            b.bind("<Leave>", lambda e: b.configure(bg=bg))
            return b

        def _set_next_enabled(enabled, is_last=False):
            """Toggle Next button between a bright enabled state and a clearly dimmed disabled state."""
            label = ("Finish  ✓" if is_last else "Next →") if enabled else "Complete task first…"
            if enabled:
                next_btn.config(
                    state="normal", text=label,
                    bg=_BTN_OK, fg="#ffffff", cursor="hand2",
                    activebackground=_BTN_OK_HOV,
                )
                next_btn.bind("<Enter>", lambda e: next_btn.configure(bg=_BTN_OK_HOV))
                next_btn.bind("<Leave>", lambda e: next_btn.configure(bg=_BTN_OK))
            else:
                next_btn.config(
                    state="disabled", text=label,
                    bg="#2a2a2a", fg="#555555", cursor="arrow",
                    activebackground="#2a2a2a",
                )
                next_btn.unbind("<Enter>")
                next_btn.unbind("<Leave>")

        back_btn = _mkb(btn_frame, "← Back", _back, _BTN_BG, _BTN_HOV)
        back_btn.pack(side="left")
        next_btn = _mkb(btn_frame, "Next →", _next, _BTN_OK, _BTN_OK_HOV)
        next_btn.pack(side="right", padx=(6, 0))
        _mkb(btn_frame, "Exit Tutorial", _skip, "#5a3030", "#7a4040").pack(side="right")
        do_it_btn = _mkb(btn_frame, "Do it", lambda: None, _DO_IT_BG, _DO_IT_HOV)

        win.bind("<Return>", lambda e: _next() if str(next_btn["state"]) != "disabled" else None)
        win.bind("<Escape>", lambda e: _skip())
        win.protocol("WM_DELETE_WINDOW", _skip)

        _show_step(0)
