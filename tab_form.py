"""
TabForm — ticket editor panel with ADF sync, field widgets, and preview.
"""
import os
import re
import json
import copy
import uuid
import traceback
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import html as _html
from html.parser import HTMLParser

from config import HEADERS, MULTISELECT_FIELDS, FETCHABLE_OPTION_FIELDS
from utils import _bind_mousewheel, _bind_mousewheel_to_target, _bind_mousewheel_to_target_recursive, debug_log
from desc_mixin import DescMixin


def _patch_tkinterweb_caret_hide():
    """Guard against tkinterweb caret hide crash when caret_frame is None."""
    try:
        from tkinterweb import extensions as _tkweb_ext  # type: ignore
        cls = getattr(_tkweb_ext, "CaretManager", None)
        if not cls:
            return
        if getattr(cls, "_avalanche_hide_patched", False):
            return
        original_hide = cls.hide

        def _safe_hide(self, *args, **kwargs):
            if getattr(self, "caret_frame", None) is None:
                return
            try:
                return original_hide(self, *args, **kwargs)
            except AttributeError as ex:
                if "place_forget" in str(ex):
                    return
                raise

        cls.hide = _safe_hide
        cls._avalanche_hide_patched = True
    except Exception:
        pass


def _patch_tkinterweb_focusout():
    """Guard tkinterweb focus-out callback from crashing on caret hide."""
    try:
        from tkinterweb import bindings as _tkweb_bindings  # type: ignore
        cls = getattr(_tkweb_bindings, "TkinterWeb", None)
        if not cls:
            return
        if getattr(cls, "_avalanche_focusout_patched", False):
            return
        original_focusout = cls._on_focusout

        def _safe_focusout(self, event):
            try:
                return original_focusout(self, event)
            except AttributeError as ex:
                if "place_forget" in str(ex):
                    return
                raise

        cls._on_focusout = _safe_focusout
        cls._avalanche_focusout_patched = True
    except Exception:
        pass


_patch_tkinterweb_caret_hide()
_patch_tkinterweb_focusout()


# ---------------- TabForm (with ADF sync) ----------------
class TabForm(DescMixin):
    """
    Ticket editor panel (one tab). Includes:
      - Summary (text)
      - Description (text)
      - Description ADF (JSON editor)
      - In-app rendered preview (HTML widget if available, otherwise text fallback)
    """
    def __init__(self, parent_frame, meta_options, field_menu_cb=None, add_option_cb=None, convert_text_to_adf_cb=None, extract_text_from_adf_cb=None, internal_priority_set_cb=None, close_tab_cb=None, open_attachments_cb=None, open_in_jira_cb=None, fetch_assignees_cb=None, fetch_options_cb=None, resolve_variables_in_adf_cb=None, collect_vars_cb=None, notify_vars_changed_cb=None, open_ticket_in_app_cb=None, open_ticket_in_jira_cb=None):
        self.frame = ttk.Frame(parent_frame)
        self._scroll_canvas = tk.Canvas(self.frame, highlightthickness=0, bg="#2b2b2b")
        self._scroll_sb = ttk.Scrollbar(self.frame, orient="vertical", command=self._scroll_canvas.yview)
        self._scroll_canvas.pack(side="left", fill="both", expand=True)
        self._scroll_sb.pack(side="right", fill="y")
        self._scroll_canvas.configure(yscrollcommand=self._scroll_sb.set)
        self._content = ttk.Frame(self._scroll_canvas)
        self._canvas_win_id = self._scroll_canvas.create_window((0, 0), window=self._content, anchor="nw")
        def _on_content_configure(e):
            self._scroll_canvas.configure(scrollregion=self._scroll_canvas.bbox("all"))
        def _on_canvas_configure(e):
            self._scroll_canvas.itemconfig(self._canvas_win_id, width=max(e.width, 1))
        self._content.bind("<Configure>", _on_content_configure)
        self._scroll_canvas.bind("<Configure>", _on_canvas_configure)
        _bind_mousewheel(self._scroll_canvas, "vertical")
        self.meta_options = meta_options
        self.field_menu_cb = field_menu_cb
        self.add_option_cb = add_option_cb
        self.open_attachments_cb = open_attachments_cb
        self.open_in_jira_cb = open_in_jira_cb
        self.fetch_assignees_cb = fetch_assignees_cb
        self.fetch_options_cb = fetch_options_cb
        self.resolve_variables_in_adf_cb = resolve_variables_in_adf_cb
        self.collect_vars_cb = collect_vars_cb
        self.notify_vars_changed_cb = notify_vars_changed_cb
        self.convert_text_to_adf_cb = convert_text_to_adf_cb
        self.extract_text_from_adf_cb = extract_text_from_adf_cb
        self.internal_priority_set_cb = internal_priority_set_cb
        self.close_tab_cb = close_tab_cb
        self.open_ticket_in_app_cb = open_ticket_in_app_cb
        self.open_ticket_in_jira_cb = open_ticket_in_jira_cb
        self._last_ticket_key = None
        self.field_widgets = {}
        self.collapse_mode = False

        # sync control
        self._suppress_sync = False
        self._desc_after_id = None
        self._adf_after_id = None
        self._editor_args_path = None   # set while editor subprocess is running
        self._sync_delay = 800  # ms debounce
        self._adf_preview_widget = None
        self._adf_html_widget = None
        self._adf_preview_editable = False
        self._adf_preview_after_id = None
        # Info bar for Status/Priority (updated on populate)
        self._info_bar = ttk.Frame(self._content)
        self._info_bar.pack(fill="x", padx=4, pady=(4, 8))
        ttk.Label(self._info_bar, text="Key:", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 4))
        self._info_key_lbl = ttk.Label(self._info_bar, text="—", font=("Segoe UI", 9))
        self._info_key_lbl.pack(side="left", padx=(0, 4))
        def _do_open_in_jira():
            if self.open_in_jira_cb and self._last_ticket_key:
                self.open_in_jira_cb(self._last_ticket_key)
        self._info_open_jira_btn = ttk.Button(self._info_bar, text="Jira", width=5, command=_do_open_in_jira, state="disabled")
        self._info_open_jira_btn.pack(side="left", padx=(0, 16))
        ttk.Label(self._info_bar, text="Status:", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 4))
        self._info_status_lbl = ttk.Label(self._info_bar, text="—", font=("Segoe UI", 9))
        self._info_status_lbl.pack(side="left", padx=(0, 16))
        ttk.Label(self._info_bar, text="Priority:", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 4))
        self._info_priority_lbl = ttk.Label(self._info_bar, text="—", font=("Segoe UI", 9))
        self._info_priority_lbl.pack(side="left", padx=(0, 16))
        ttk.Label(self._info_bar, text="Internal:", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 4))
        self._info_internal_var = tk.StringVar(value="None")
        self._info_internal_combo = ttk.Combobox(self._info_bar, textvariable=self._info_internal_var, width=10, state="readonly")
        self._info_internal_combo.pack(side="left")
        self._internal_priority_options = ["High", "Medium", "Low", "None"]
        self._suppress_internal_cb = False
        self._info_internal_combo.bind("<<ComboboxSelected>>", lambda e: self._on_internal_priority_changed())
        # Close tab button (×) — only on ticket tabs
        def _do_close_tab():
            if self.close_tab_cb:
                self.close_tab_cb()
        self._info_close_btn = ttk.Button(self._info_bar, text="×", width=3, command=_do_close_tab)
        self._info_close_btn.pack(side="right", padx=(8, 0))

        ttk.Separator(self._content, orient="horizontal").pack(fill="x", padx=4, pady=(0, 4))
        self._last_saved_state = None
        # Methods below may reference other TabForm helpers (not fully shown in the snippet)
        self._build_fields()
        # Sync embedded editor → ADF JSON store before any read/save
        self._pre_read_hook = lambda tf: tf._sync_htmltext_to_adf()
        self._rebind_scroll()
        # Load the dark-theme placeholder into the description viewer even for
        # blank new tabs where populate_from_dict is never called.
        self.frame.after(80, self._refresh_html_viewer)

    def _rebind_scroll(self):
        """Walk all descendants of _content and bind mousewheel to scroll
        the main canvas.  Skips Text widgets so they keep their own scroll.
        Safe to call repeatedly — newer bindings replace old ones."""
        _bind_mousewheel_to_target_recursive(
            self._content, self._scroll_canvas, "vertical")
        # HtmlFrame creates internal sub-widgets asynchronously; schedule
        # a second pass so those get bound too.
        try:
            self.frame.after(300, lambda: _bind_mousewheel_to_target_recursive(
                self._content, self._scroll_canvas, "vertical"))
        except Exception:
            pass

    def _build_fields(self):
        for hdr in HEADERS:
            if hdr == "Description":
                continue
            if hdr == "Comment":
                self._build_comments_section()
                continue
            if hdr == "Epic Link":
                self._build_epic_section()
                continue
            # Epic Name and Epic Children are rendered inside the epic section
            if hdr in ("Epic Name", "Epic Children"):
                continue
            if hdr == "Issue Links":
                self._build_issue_links_section()
                continue
            row = ttk.Frame(self._content)
            row.pack(fill="x", padx=4, pady=4)
            lbl = ttk.Label(row, text=hdr, width=28, anchor="w")
            lbl.pack(side="left", padx=(4, 8))
            if self.field_menu_cb:
                lbl.bind("<Button-3>", lambda e, h=hdr: self.field_menu_cb(e, h, self))
            _ALWAYS_INCLUDED = ("Description ADF",)
            include_var = tk.BooleanVar(value=True)
            chk = ttk.Checkbutton(row, variable=include_var)
            if hdr in _ALWAYS_INCLUDED:
                chk.pack_forget()
            else:
                chk.pack(side="left", padx=(0, 8))
            widget = None
            var = None
            add_btn = None
            adf_row = None
            if hdr in ("Summary", "Description ADF"):
                if hdr == "Summary":
                    height = 4
                else:
                    height = 18
                if hdr == "Description ADF":
                    adf_row = ttk.Frame(self._content, height=520)
                    adf_row.pack_propagate(False)
                    adf_row.pack(fill="both", padx=4, pady=(0, 6), expand=True)
                    widget = self._build_desc_field(adf_row, hdr)
                else:
                    text_frame = ttk.Frame(row)
                    text_frame.pack(side="left", fill="both", expand=True)
                    txt = tk.Text(text_frame, height=height, wrap="word", bg="#1e1e1e", fg="#dcdcdc", insertbackground="#dcdcdc")
                    vs = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
                    txt.configure(yscrollcommand=vs.set)
                    txt.pack(side="left", fill="both", expand=True)
                    vs.pack(side="right", fill="y")
                    _bind_mousewheel(txt, "vertical")
                    if self.field_menu_cb:
                        txt.bind("<Button-3>", lambda e, h=hdr: self.field_menu_cb(e, h, self))
                    widget = txt
            else:
                var = tk.StringVar()
                if hdr in MULTISELECT_FIELDS:
                    # Plain Entry for multiselect — displays any text reliably, picker opens on click
                    ms_ent = tk.Entry(row, textvariable=var, state="readonly",
                                      readonlybackground="#3c3c3c", fg="#dcdcdc",
                                      disabledforeground="#dcdcdc",
                                      font=("Segoe UI", 10), relief="flat", bd=1,
                                      highlightthickness=1, highlightbackground="#555555",
                                      cursor="hand2")
                    def _ms_click(e, h=hdr):
                        self.frame.after_idle(lambda: self._open_multiselect_dialog(h))
                        return "break"
                    ms_ent.bind("<Button-1>", _ms_click)
                    ms_ent.bind("<ButtonPress-1>", _ms_click)
                    ms_ent.bind("<space>", lambda e, h=hdr: self._open_multiselect_dialog(h))
                    ms_ent.pack(side="left", fill="x", expand=True)
                    if self.field_menu_cb:
                        ms_ent.bind("<Button-3>", lambda e, h=hdr: self.field_menu_cb(e, h, self))
                    combo = ms_ent
                else:
                    combo = ttk.Combobox(row, textvariable=var, width=60)
                    combo['values'] = self.meta_options.get(hdr, [])
                    combo.pack(side="left", fill="x", expand=True)
                    if self.field_menu_cb:
                        combo.bind("<Button-3>", lambda e, h=hdr: self.field_menu_cb(e, h, self))
                    self._setup_combo_autocomplete(combo, hdr, var)
                add_btn = ttk.Button(row, text="+", width=2, command=lambda h=hdr: self._add_option(h))
                add_btn.pack(side="left", padx=(6, 0))
                if hdr == "Attachment":
                    attach_btn = ttk.Button(row, text="Attach", width=6, command=self._on_attach_files)
                    attach_btn.pack(side="left", padx=(4, 0))
                    if self.open_attachments_cb:
                        open_btn = ttk.Button(row, text="Open", width=6, command=lambda: self.open_attachments_cb(self))
                        open_btn.pack(side="left", padx=(4, 0))
                    else:
                        open_btn = None
                elif hdr == "Assignee" and self.fetch_assignees_cb:
                    refresh_btn = ttk.Button(row, text="↻", width=2, command=self._on_refresh_assignees)
                    refresh_btn.pack(side="left", padx=(4, 0))
                    open_btn = None
                elif hdr in FETCHABLE_OPTION_FIELDS and self.fetch_options_cb:
                    refresh_btn = ttk.Button(row, text="↻", width=2, command=lambda h=hdr: self._on_refresh_field_options(h))
                    refresh_btn.pack(side="left", padx=(4, 0))
                    open_btn = None
                else:
                    open_btn = None
                widget = combo
            self.field_widgets[hdr] = {
                "var": var,
                "widget": widget,
                "include_var": include_var,
                "add_btn": add_btn,
                "row": row,
                "label": lbl,
                "adf_row": adf_row
            }

    def populate_from_dict(self, data: dict):
        if not isinstance(data, dict):
            return
        # Suppress all sync callbacks during populate (and for 2s after) to prevent
        # _on_preview_edited from overwriting the JSON editor when load_html fires.
        self._suppress_sync = True
        def _unsuppress():
            self._suppress_sync = False
        try:
            self.frame.after_cancel(getattr(self, "_unsuppress_id", None) or "")
        except Exception:
            pass
        self._unsuppress_id = self.frame.after(2000, _unsuppress)
        # Stash rendered HTML for media UUID resolution
        self._description_rendered_html = data.get("Description Rendered") or ""
        # Load comments separately — they use a custom widget
        if "Comment" in data and hasattr(self, "_comments_data"):
            self._load_comments_from_json(data.get("Comment") or [])
        # Load epic relationship
        if hasattr(self, "_epic_link_var"):
            self._load_epic_from_data(data)
        # Load issue links
        if hasattr(self, "_issue_links_data"):
            self._load_issue_links_from_json(data.get("Issue Links") or "[]")

        for h in HEADERS:
            if h in ("Comment", "Epic Link", "Epic Name", "Epic Children", "Issue Links"):
                continue   # handled by custom section builders
            info = self.field_widgets.get(h)
            if not info:
                continue
            widget = info.get("widget")
            include_var = info.get("include_var")
            try:
                include_var.set(bool(data.get(h) or (h == "Description ADF" and data.get("Description ADF"))))
            except Exception:
                pass
            val = data.get(h, "")
            try:
                if isinstance(widget, tk.Text):
                    widget.config(state="normal")
                    widget.delete("1.0", "end")
                    if isinstance(val, (dict, list)):
                        widget.insert("1.0", json.dumps(val, ensure_ascii=False, indent=2))
                    elif h == "Description ADF" and isinstance(val, str) and val.strip():
                        try:
                            adf = json.loads(val)
                            if isinstance(adf, dict):
                                widget.insert("1.0", json.dumps(adf, ensure_ascii=False, indent=2))
                            else:
                                widget.insert("1.0", str(val))
                        except Exception:
                            widget.insert("1.0", str(val))
                    else:
                        widget.insert("1.0", str(val or ""))
                    widget.config(state="normal")
                elif isinstance(widget, ttk.Combobox):
                    if h == "Attachment" and val and str(val).strip().startswith("["):
                        try:
                            items = json.loads(val)
                            if isinstance(items, list):
                                setattr(self, "_attachment_raw_json", val)
                                display = "; ".join((x.get("filename") or x.get("name") or "") for x in items if isinstance(x, dict))
                                widget.set(display or val)
                            else:
                                setattr(self, "_attachment_raw_json", None)
                                widget.set(val or "")
                        except Exception:
                            setattr(self, "_attachment_raw_json", None)
                            widget.set(val or "")
                    else:
                        setattr(self, "_attachment_raw_json", None)
                        widget.set(val or "")
                else:
                    try:
                        display_val = val or ""
                        if h in MULTISELECT_FIELDS and isinstance(display_val, str) and display_val:
                            parts = [p.strip() for p in display_val.split(";") if p.strip()]
                            display_val = "; ".join(parts)
                        sv = info.get("var")
                        if sv and isinstance(sv, tk.StringVar):
                            sv.set(display_val)
                        else:
                            prev_state = str(widget.cget("state"))
                            widget.config(state="normal")
                            widget.delete(0, "end")
                            widget.insert(0, display_val)
                            widget.config(state=prev_state)
                    except Exception:
                        pass
            except Exception:
                pass
        # Refresh the description viewer once the JSON store is populated
        self.frame.after(50, self._refresh_html_viewer)
        # Track as saved state for dirty check
        try:
            self._last_saved_state = json.dumps(self.read_to_dict(), sort_keys=True, default=str)
        except Exception:
            self._last_saved_state = None
        # Update info bar (Status, Priority, Internal)
        try:
            self._last_ticket_key = data.get("Issue key") or data.get("Issue id")
            self._info_key_lbl.config(text=data.get("Issue key") or "—")
            key = str(self._last_ticket_key or "").strip()
            if hasattr(self, "_info_open_jira_btn"):
                self._info_open_jira_btn.config(state="normal" if key and not key.startswith("LOCAL-") else "disabled")
            self._info_status_lbl.config(text=data.get("Status") or "—")
            self._info_priority_lbl.config(text=data.get("Priority") or "—")
            internal = data.get("Internal Priority", "None")
            if hasattr(self, "_info_internal_combo"):
                opts = getattr(self, "_internal_priority_options", ["High", "Medium", "Low", "None"])
                self._info_internal_combo["values"] = opts
                self._info_internal_var.set(internal if internal in opts else "None")
        except Exception:
            pass

        # Re-order sections for Epics: children + links near the top
        self._reorder_for_epic_if_needed(data)

    def _reorder_for_epic_if_needed(self, data: dict):
        """When the ticket is an Epic, move the Epic section and Issue Links
        section right below the Description so child tickets are prominent."""
        is_epic = str(data.get("Issue Type") or "").strip().lower() == "epic"
        epic_outer = getattr(self, "_epic_section_outer", None)
        links_outer = getattr(self, "_issue_links_section_outer", None)

        if not is_epic or not epic_outer:
            return

        # Find the Description ADF row as the anchor — sections go right after it
        adf_info = self.field_widgets.get("Description ADF")
        anchor = adf_info.get("adf_row") if adf_info else None
        if not anchor:
            return

        try:
            # Re-pack epic section right after the description
            epic_outer.pack_forget()
            epic_outer.pack(fill="x", padx=4, pady=(2, 4), after=anchor)
            # Re-pack issue links right after the epic section
            if links_outer:
                links_outer.pack_forget()
                links_outer.pack(fill="x", padx=4, pady=(2, 4), after=epic_outer)
        except Exception:
            pass

    def _open_multiselect_dialog(self, hdr):
        """Open a clean dark-mode picker with live search. Single click toggles selection."""
        info = self.field_widgets.get(hdr)
        if not info:
            return
        var = info.get("var")
        widget = info.get("widget")
        if not var or not widget:
            return
        options = list(self.meta_options.get(hdr, []))
        current_val = (var.get() or "").strip()
        current_items = [x.strip() for x in current_val.split(";") if x.strip()]
        all_items = sorted(set(options) | set(current_items), key=lambda x: x.lower())
        if not all_items:
            messagebox.showinfo("Select", f"No options available for {hdr}. Use ↻ to fetch from Jira first.")
            return

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

        # Track selections by value so they survive filtering
        selected_set = set(current_items)
        visible_items = []   # starts empty; _rebuild_list() populates it

        win = tk.Toplevel(self.frame.winfo_toplevel())
        win.title(f"Select {hdr}")
        win.configure(bg=_BG)
        win.minsize(300, 400)
        win.geometry("360x480")
        win.resizable(True, True)
        win

        # Title bar
        title_bar = tk.Frame(win, bg=_PANEL, pady=10)
        title_bar.pack(fill="x")
        tk.Label(title_bar, text=f"Select {hdr}", bg=_PANEL, fg=_FG,
                 font=("Segoe UI", 11, "bold"), padx=16).pack(side="left")
        tk.Frame(win, bg=_BORDER, height=1).pack(fill="x")

        # Search bar
        search_frame = tk.Frame(win, bg=_BG, padx=12)
        search_frame.pack(fill="x", pady=(10, 4))
        search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_frame, textvariable=search_var, bg=_SEARCH_BG, fg=_FG,
            insertbackground=_FG, relief="flat", bd=0, font=("Segoe UI", 10),
            highlightthickness=1, highlightbackground=_BORDER, highlightcolor="#007acc"
        )
        search_entry.pack(fill="x", ipady=6, padx=1)
        # Placeholder
        search_entry.insert(0, "Search...")
        search_entry.configure(fg="#666666")
        def _search_focus_in(e):
            if search_var.get() == "Search...":
                search_entry.delete(0, tk.END)
                search_entry.configure(fg=_FG)
        def _search_focus_out(e):
            if not search_var.get().strip():
                search_entry.delete(0, tk.END)
                search_entry.insert(0, "Search...")
                search_entry.configure(fg="#666666")
        search_entry.bind("<FocusIn>", _search_focus_in)
        search_entry.bind("<FocusOut>", _search_focus_out)

        # List
        list_outer = tk.Frame(win, bg=_BG, padx=12)
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
        _rebuilding = [False]   # guard: prevents <<ListboxSelect>> from wiping sel during rebuild

        def _update_count(*_):
            count_var.set(f"{len(selected_set)} selected")

        def _rebuild_list(query=""):
            # selected_set is the single source of truth — do NOT read back from
            # the listbox here.  On Windows the listbox fires <<ListboxSelect>>
            # with nothing selected when it loses focus, which would wipe all picks
            # before _rebuilding could be set to True.
            _rebuilding[0] = True
            q = query.lower().strip()
            visible_items[:] = [x for x in all_items if q in x.lower()] if q else list(all_items)
            lb.delete(0, tk.END)
            for item in visible_items:
                lb.insert(tk.END, "  " + item)
            for i, item in enumerate(visible_items):
                if item in selected_set:
                    lb.selection_set(i)
            _rebuilding[0] = False
            _update_count()

        def _on_search(*_):
            q = search_var.get()
            if q == "Search...":
                q = ""
            _rebuild_list(q)

        def _on_lb_select(*_):
            if _rebuilding[0]:
                return
            new_sel = {visible_items[i] for i in range(len(visible_items))
                       if lb.selection_includes(i)}
            # On Windows, clicking away from the listbox fires <<ListboxSelect>>
            # with nothing selected even though the user never deselected anything.
            # Detect this: if every visible item just became unselected but
            # selected_set is non-empty, restore the visual state and bail out.
            if not new_sel and selected_set:
                for i, item in enumerate(visible_items):
                    if item in selected_set:
                        lb.selection_set(i)
                return
            for i, item in enumerate(visible_items):
                if lb.selection_includes(i):
                    selected_set.add(item)
                else:
                    selected_set.discard(item)
            _update_count()
            # Live-update the field immediately
            var.set("; ".join(x for x in all_items if x in selected_set))

        search_var.trace_add("write", lambda *_: _on_search())
        lb.bind("<<ListboxSelect>>", _on_lb_select)
        _rebuild_list()

        tk.Label(win, textvariable=count_var, bg=_BG, fg="#888888",
                 font=("Segoe UI", 9), anchor="w", padx=14).pack(fill="x", pady=(4, 0))

        # Divider + buttons
        tk.Frame(win, bg=_BORDER, height=1).pack(fill="x", pady=(6, 0))
        btn_frame = tk.Frame(win, bg=_PANEL, pady=10, padx=12)
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
            _on_lb_select()   # flush final state
            var.set("; ".join(x for x in all_items if x in selected_set))
            win.destroy()

        def _clear_all():
            selected_set.clear()
            lb.selection_clear(0, tk.END)
            _update_count()
            var.set("")

        _make_btn(btn_frame, "OK", ok, _BTN_OK, _BTN_OK_HOV).pack(side="right", padx=(4, 0))
        _make_btn(btn_frame, "Cancel", win.destroy, _BTN_BG, _BTN_HOV).pack(side="right")
        _make_btn(btn_frame, "Clear", _clear_all, "#5a3030", "#7a4040").pack(side="left")
        win.bind("<Return>", lambda e: ok())
        win.bind("<Escape>", lambda e: win.destroy())
        win.after(50, search_entry.focus_set)

    # ── Epic section ──────────────────────────────────────────────────────────

    def _build_epic_section(self):
        """Build the Epic UI: colored badge + Search / Clear buttons."""
        _BG      = "#1a1a1a"
        _PANEL   = "#252526"
        _BORDER  = "#3c3c3c"
        _FG      = "#d4d4d4"
        _META_FG = "#888888"
        _EPIC_BG = "#3b1f6e"   # Jira-style purple epic badge
        _EPIC_FG = "#c7a3ff"

        outer = tk.Frame(self._content, bg=_PANEL, bd=1, relief="flat",
                         highlightthickness=1, highlightbackground=_BORDER)
        outer.pack(fill="x", padx=4, pady=(2, 4))

        # Header row
        hdr_bar = tk.Frame(outer, bg=_PANEL, pady=5, padx=10)
        hdr_bar.pack(fill="x")
        tk.Label(hdr_bar, text="⬡  Epic", bg=_PANEL, fg=_FG,
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        def _do_search_epic():
            self._open_epic_picker()

        def _do_clear_epic():
            self._epic_link_var.set("")
            self._epic_name_var.set("")
            self._refresh_epic_view()

        btn_frame = tk.Frame(hdr_bar, bg=_PANEL)
        btn_frame.pack(side="right")
        clear_btn = tk.Button(btn_frame, text="×", width=2,
                              command=_do_clear_epic,
                              bg=_PANEL, fg=_META_FG, relief="flat", bd=0,
                              font=("Segoe UI", 9), cursor="hand2")
        clear_btn.pack(side="right", padx=(4, 0))
        search_btn = tk.Button(btn_frame, text="Search…",
                               command=_do_search_epic,
                               bg="#0e639c", fg="#ffffff", relief="flat", bd=0,
                               font=("Segoe UI", 9), padx=10, pady=3,
                               cursor="hand2",
                               activebackground="#1177bb", activeforeground="#ffffff")
        search_btn.bind("<Enter>", lambda e: search_btn.configure(bg="#1177bb"))
        search_btn.bind("<Leave>", lambda e: search_btn.configure(bg="#0e639c"))
        search_btn.pack(side="right")

        tk.Frame(outer, bg=_BORDER, height=1).pack(fill="x")

        # Badge display area
        badge_frame = tk.Frame(outer, bg=_BG, padx=10, pady=6)
        badge_frame.pack(fill="x")
        self._epic_badge_lbl = tk.Label(
            badge_frame, text="No epic assigned",
            bg=_BG, fg=_META_FG,
            font=("Segoe UI", 9, "italic"))
        self._epic_badge_lbl.pack(side="left")

        # StringVars to hold epic key + name
        self._epic_link_var = tk.StringVar()
        self._epic_name_var = tk.StringVar()

        # Child Issues section (shown dynamically when Issue Type = Epic)
        self._children_outer = tk.Frame(outer, bg=_BG)
        # (packed/hidden by _refresh_epic_view)

        self.field_widgets["Epic Link"] = {
            "widget": None, "var": self._epic_link_var,
            "include_var": tk.BooleanVar(value=True),
            "row": outer, "label": None, "adf_row": None,
        }
        self._epic_section_outer = outer

        self._epic_children_data: list = []
        self._epic_children_loading = False
        self._refresh_epic_view()

    def _refresh_epic_view(self):
        """Redraw the epic badge and optionally show child issues."""
        _BG      = "#1a1a1a"
        _PANEL   = "#252526"
        _BORDER  = "#3c3c3c"
        _META_FG = "#888888"
        _EPIC_BG = "#3b1f6e"
        _EPIC_FG = "#c7a3ff"
        _FG      = "#d4d4d4"

        key  = getattr(self, "_epic_link_var", None)
        key  = key.get().strip() if key else ""
        name = getattr(self, "_epic_name_var", None)
        name = name.get().strip() if name else ""

        lbl = getattr(self, "_epic_badge_lbl", None)
        if lbl is None:
            return

        if key:
            display = f"  {key}  ·  {name}" if name else f"  {key}  "
            lbl.config(
                text=display, bg=_EPIC_BG, fg=_EPIC_FG,
                font=("Segoe UI", 9, "bold"), relief="flat",
                padx=8, pady=3, cursor="hand2")
            lbl.bind("<Button-1>", lambda e: self._show_ticket_link_menu_at(key, e))
            lbl.bind("<Button-3>", lambda e: self._show_ticket_link_menu_at(key, e))
        else:
            lbl.config(
                text="No epic assigned", bg=_BG, fg=_META_FG,
                font=("Segoe UI", 9, "italic"), relief="flat",
                padx=0, pady=0, cursor="")
            lbl.unbind("<Button-1>")
            lbl.unbind("<Button-3>")

        # Child issues section — shown when this ticket IS an Epic
        outer_children = getattr(self, "_children_outer", None)
        if outer_children is None:
            return
        children = getattr(self, "_epic_children_data", [])
        issue_type = ""
        try:
            info = self.field_widgets.get("Issue Type")
            if info:
                v = info.get("var")
                if v:
                    issue_type = v.get().strip()
        except Exception:
            pass

        is_epic = issue_type.lower() == "epic"
        loading = getattr(self, "_epic_children_loading", False)

        if is_epic:
            outer_children.pack(fill="x")
            for w in outer_children.winfo_children():
                w.destroy()
            tk.Frame(outer_children, bg=_BORDER, height=1).pack(fill="x")
            hdr = tk.Frame(outer_children, bg=_PANEL, padx=10, pady=4)
            hdr.pack(fill="x")
            tk.Label(hdr, text="Child Issues", bg=_PANEL, fg=_META_FG,
                     font=("Segoe UI", 8, "bold")).pack(side="left")
            add_btn = tk.Button(
                hdr, text="+ Add Child", command=self._open_child_issue_picker,
                bg="#0e639c", fg="#ffffff", relief="flat", bd=0,
                font=("Segoe UI", 8), padx=8, pady=2, cursor="hand2",
                activebackground="#1177bb", activeforeground="#ffffff")
            add_btn.pack(side="right")
            add_btn.bind("<Enter>", lambda e: add_btn.configure(bg="#1177bb"))
            add_btn.bind("<Leave>", lambda e: add_btn.configure(bg="#0e639c"))
            if loading and not children:
                load_row = tk.Frame(outer_children, bg=_BG, padx=12, pady=8)
                load_row.pack(fill="x")
                tk.Label(load_row, text="Loading child issues from Jira\u2026",
                         bg=_BG, fg="#4ec9b0",
                         font=("Segoe UI", 9, "italic")).pack(side="left")
            elif children:
                for child in children:
                    child_key = child.get("key", "")
                    child_sum = child.get("summary", "")
                    child_st  = child.get("status", "")
                    row = tk.Frame(outer_children, bg=_BG, padx=10, pady=3)
                    row.pack(fill="x")
                    status_badge = tk.Label(row, text=f" {child_st} ",
                                            bg="#2a4a2a", fg="#6dbf67",
                                            font=("Segoe UI", 8), relief="flat")
                    status_badge.pack(side="left", padx=(0, 6))
                    key_lbl = tk.Label(row, text=child_key, bg=_BG, fg="#4a9eff",
                                       font=("Segoe UI", 9, "bold"), cursor="hand2")
                    key_lbl.pack(side="left")
                    key_lbl.bind("<Button-1>", lambda e, k=child_key: self._show_ticket_link_menu_at(k, e))
                    key_lbl.bind("<Button-3>", lambda e, k=child_key: self._show_ticket_link_menu_at(k, e))
                    tk.Label(row, text=f"  {child_sum}", bg=_BG, fg=_FG,
                             font=("Segoe UI", 9)).pack(side="left")
            else:
                empty_row = tk.Frame(outer_children, bg=_BG, padx=12, pady=6)
                empty_row.pack(fill="x")
                tk.Label(empty_row, text="No child issues — use \"+ Add Child\" to assign tickets",
                         bg=_BG, fg=_META_FG,
                         font=("Segoe UI", 8, "italic")).pack(side="left")
        else:
            outer_children.pack_forget()
        self._rebind_scroll()

    def _show_ticket_link_menu_at(self, key: str, event=None):
        """Show the open-in-app / open-in-jira popup at the event position."""
        try:
            menu = tk.Menu(self.frame, tearoff=0,
                           bg="#2d2d2d", fg="#d4d4d4",
                           activebackground="#094771",
                           activeforeground="#ffffff",
                           font=("Segoe UI", 9))
            if self.open_ticket_in_app_cb:
                menu.add_command(
                    label=f"Open  {key}  locally",
                    command=lambda: self.open_ticket_in_app_cb(key),
                )
            if self.open_ticket_in_jira_cb:
                menu.add_command(
                    label=f"Open  {key}  in Jira",
                    command=lambda: self.open_ticket_in_jira_cb(key),
                )
            x = event.x_root if event else self.frame.winfo_rootx() + 40
            y = event.y_root if event else self.frame.winfo_rooty() + 40
            menu.tk_popup(x, y)
        except Exception:
            pass

    # Class-level epic cache shared across all TabForm instances.
    # Keyed by project_key (or "" for all-projects).
    _epic_cache: dict = {}   # {project_key: [epic_dicts]}

    def _open_epic_picker(self):
        """Open a dialog to search for and select an epic from Jira.

        Uses a session-level cache so the full epic list is only fetched once
        per project.  A ↻ Refresh button lets the user re-fetch when needed.
        """
        project_key = "SUNDANCE"
        try:
            info = self.field_widgets.get("Project key")
            if info:
                v = info.get("var")
                if v and v.get().strip():
                    project_key = v.get().strip()
        except Exception:
            pass

        root = self.frame.winfo_toplevel()
        if not hasattr(root, "get_jira_session"):
            return
        session = root.get_jira_session()
        if not session:
            return

        win = tk.Toplevel(self.frame.winfo_toplevel())
        win.title("Select Epic")
        win.geometry("620x520")
        win.resizable(True, True)
        try:
            win.grab_set()
        except Exception:
            pass

        _BG     = "#1e1e1e"
        _PANEL  = "#252526"
        _FG     = "#d4d4d4"
        _BORDER = "#3c3c3c"

        win.configure(bg=_BG)
        main = tk.Frame(win, bg=_BG, padx=12, pady=10)
        main.pack(fill="both", expand=True)

        # ── search bar ────────────────────────────────────────────────────────
        search_hdr = tk.Frame(main, bg=_BG)
        search_hdr.pack(fill="x")
        tk.Label(search_hdr, text="Filter epics (type to search):",
                 bg=_BG, fg=_FG, font=("Segoe UI", 9)).pack(side="left")
        refresh_btn = tk.Button(
            search_hdr, text="↻ Refresh from Jira",
            bg="#3c3c3c", fg=_FG, relief="flat", bd=0,
            font=("Segoe UI", 8), padx=8, pady=2, cursor="hand2",
            activebackground="#505050", activeforeground="#ffffff")
        refresh_btn.pack(side="right")
        refresh_btn.bind("<Enter>", lambda e: refresh_btn.configure(bg="#505050"))
        refresh_btn.bind("<Leave>", lambda e: refresh_btn.configure(bg="#3c3c3c"))

        search_var = tk.StringVar()
        search_frame = tk.Frame(main, bg=_BORDER, bd=1, relief="flat",
                                highlightthickness=1, highlightbackground=_BORDER)
        search_frame.pack(fill="x", pady=(4, 8))
        search_ent = tk.Entry(search_frame, textvariable=search_var,
                              bg=_BG, fg=_FG, insertbackground=_FG,
                              font=("Segoe UI", 10), relief="flat", bd=0)
        search_ent.pack(fill="x", padx=6, pady=5)

        # ── result list ───────────────────────────────────────────────────────
        list_frame = tk.Frame(main, bg=_BG)
        list_frame.pack(fill="both", expand=True, pady=(0, 6))
        lb = tk.Listbox(list_frame, bg=_PANEL, fg=_FG,
                        selectbackground="#094771", selectforeground="#ffffff",
                        font=("Segoe UI", 9), relief="flat", bd=0,
                        activestyle="none", highlightthickness=0)
        vsb = tk.Scrollbar(list_frame, orient="vertical", command=lb.yview,
                           bg=_PANEL, troughcolor=_BG)
        lb.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        lb.pack(side="left", fill="both", expand=True)

        status_lbl = tk.Label(main, text="", bg=_BG, fg="#888888",
                              font=("Segoe UI", 8, "italic"))
        status_lbl.pack(anchor="w", pady=(0, 6))

        _all_epics: list = []    # [{key, name, summary}]
        _filtered:  list = []    # subset currently shown in lb
        _after_id = [None]

        def _apply_filter(*_):
            if _after_id[0]:
                win.after_cancel(_after_id[0])
            _after_id[0] = win.after(120, _do_filter)

        def _do_filter():
            q = search_var.get().strip().lower()
            _filtered.clear()
            lb.delete(0, tk.END)
            for ep in _all_epics:
                if (not q
                        or q in ep["key"].lower()
                        or q in ep["name"].lower()
                        or q in ep["summary"].lower()):
                    _filtered.append(ep)
                    label = ep["key"] + ("  ·  " + (ep["name"] or ep["summary"])
                                         if (ep["name"] or ep["summary"]) else "")
                    lb.insert(tk.END, f"  {label}")
            visible = len(_filtered)
            total   = len(_all_epics)
            status_lbl.config(
                text=f"Showing {visible} of {total} epic(s)."
                if q else f"{total} epic(s) loaded.")

        search_var.trace_add("write", _apply_filter)

        # ── loader (uses cache, fetches only if needed) ───────────────────────
        import threading

        def _populate_from_cache():
            cached = TabForm._epic_cache.get(project_key)
            if cached:
                _all_epics.clear()
                _all_epics.extend(cached)
                _do_filter()
                return True
            return False

        def _fetch_from_jira():
            status_lbl.config(text="Loading epics from Jira…")
            refresh_btn.config(state="disabled", text="Loading…")

            def _load_all():
                try:
                    base_jql = (
                        f'issuetype = Epic AND project = "{project_key}" ORDER BY updated DESC'
                        if project_key else
                        'issuetype = Epic ORDER BY updated DESC'
                    )
                    _FIELDS = ["summary", "customfield_10011", "issuetype", "status"]
                    PAGE = 100
                    collected = []
                    token = None

                    while True:
                        body = {
                            "jql":        base_jql,
                            "maxResults": PAGE,
                            "fields":     _FIELDS,
                        }
                        if token:
                            body["nextPageToken"] = token
                        try:
                            url = f"{session._jira_base}/rest/api/3/search/jql"
                            from jira_api import perform_jira_request
                            resp = perform_jira_request(session, "POST", url,
                                                        json_body=body, timeout=60)
                            resp.raise_for_status()
                            data = resp.json()
                        except Exception:
                            debug_log("Epic picker load page error: " + traceback.format_exc())
                            break

                        issues = data.get("issues", [])
                        for iss in issues:
                            f   = iss.get("fields") or {}
                            k   = iss.get("key", "")
                            name = f.get("customfield_10011") or ""
                            summ = f.get("summary") or ""
                            collected.append({"key": k, "name": name, "summary": summ})

                        snapshot = list(collected)
                        win.after(0, lambda s=snapshot: _update_list(s))

                        token = data.get("nextPageToken")
                        if not token or len(issues) < PAGE:
                            break

                    TabForm._epic_cache[project_key] = list(collected)

                except Exception:
                    debug_log("Epic picker loader error: " + traceback.format_exc())
                    win.after(0, lambda: status_lbl.config(text="Load failed — see debug log."))
                finally:
                    win.after(0, lambda: refresh_btn.config(state="normal", text="↻ Refresh from Jira"))

            threading.Thread(target=_load_all, daemon=True).start()

        def _update_list(snapshot):
            _all_epics.clear()
            _all_epics.extend(snapshot)
            _do_filter()

        def _on_refresh():
            TabForm._epic_cache.pop(project_key, None)
            _fetch_from_jira()

        refresh_btn.config(command=_on_refresh)

        if not _populate_from_cache():
            _fetch_from_jira()

        # ── confirm / cancel ──────────────────────────────────────────────────
        def _confirm(*_):
            sel = lb.curselection()
            if not sel:
                return
            idx = sel[0]
            if idx < len(_filtered):
                chosen = _filtered[idx]
                self._epic_link_var.set(chosen["key"])
                self._epic_name_var.set(chosen["name"] or chosen["summary"])
                self._refresh_epic_view()
            win.destroy()

        btn_row = tk.Frame(main, bg=_BG)
        btn_row.pack(fill="x")
        tk.Button(btn_row, text="Select", command=_confirm,
                  bg="#0e639c", fg="#ffffff", relief="flat", bd=0,
                  font=("Segoe UI", 9, "bold"), padx=14, pady=5,
                  cursor="hand2", activebackground="#1177bb").pack(side="right", padx=(4, 0))
        tk.Button(btn_row, text="Cancel", command=win.destroy,
                  bg="#3c3c3c", fg=_FG, relief="flat", bd=0,
                  font=("Segoe UI", 9), padx=14, pady=5,
                  cursor="hand2").pack(side="right")

        lb.bind("<Double-Button-1>", _confirm)
        lb.bind("<Return>",     _confirm)
        search_ent.bind("<Return>",    lambda e: (_do_filter() or None))
        search_ent.bind("<KP_Enter>",  lambda e: (_do_filter() or None))
        win.bind("<Escape>", lambda e: win.destroy())

        search_ent.focus_set()

    # ── Child Issue Picker ─────────────────────────────────────────────────
    def _open_child_issue_picker(self):
        """Open a dialog to search ALL Jira tickets and assign selected ones
        as children of this epic.  The search queries Jira live (debounced)
        so even tickets not fetched locally appear in the results."""
        epic_key = ""
        try:
            v = getattr(self, "_epic_link_var", None)
            if v:
                epic_key = v.get().strip()
        except Exception:
            pass
        if not epic_key:
            info = self.field_widgets.get("Issue key")
            if info:
                v = info.get("var")
                if v:
                    epic_key = v.get().strip()
        if not epic_key or epic_key.startswith("LOCAL-"):
            from tkinter import messagebox
            messagebox.showwarning("Cannot add children",
                                   "Save / upload this epic to Jira first.",
                                   parent=self.frame.winfo_toplevel())
            return

        root = self.frame.winfo_toplevel()
        if not hasattr(root, "get_jira_session"):
            return
        session = root.get_jira_session()
        if not session:
            return

        project_key = "SUNDANCE"
        try:
            info = self.field_widgets.get("Project key")
            if info:
                pv = info.get("var")
                if pv and pv.get().strip():
                    project_key = pv.get().strip()
        except Exception:
            pass

        win = tk.Toplevel(self.frame.winfo_toplevel())
        win.title(f"Add Child Issues to {epic_key}")
        win.geometry("700x560")
        win.resizable(True, True)
        try:
            win.grab_set()
        except Exception:
            pass

        _BG     = "#1e1e1e"
        _PANEL  = "#252526"
        _FG     = "#d4d4d4"
        _BORDER = "#3c3c3c"
        _ACCENT = "#4a9eff"

        win.configure(bg=_BG)
        main = tk.Frame(win, bg=_BG, padx=12, pady=10)
        main.pack(fill="both", expand=True)

        tk.Label(main, text="Search Jira for tickets to add as children:",
                 bg=_BG, fg=_FG, font=("Segoe UI", 9)).pack(anchor="w")
        search_var = tk.StringVar()
        search_frame = tk.Frame(main, bg=_BORDER, bd=1, relief="flat",
                                highlightthickness=1, highlightbackground=_BORDER)
        search_frame.pack(fill="x", pady=(4, 8))
        search_ent = tk.Entry(search_frame, textvariable=search_var,
                              bg=_BG, fg=_FG, insertbackground=_FG,
                              font=("Segoe UI", 10), relief="flat", bd=0)
        search_ent.pack(fill="x", padx=6, pady=5)

        list_frame = tk.Frame(main, bg=_BG)
        list_frame.pack(fill="both", expand=True, pady=(0, 6))
        lb = tk.Listbox(list_frame, bg=_PANEL, fg=_FG,
                        selectbackground="#094771", selectforeground="#ffffff",
                        font=("Segoe UI", 9), relief="flat", bd=0,
                        activestyle="none", highlightthickness=0,
                        selectmode="multiple")
        vsb = tk.Scrollbar(list_frame, orient="vertical", command=lb.yview,
                           bg=_PANEL, troughcolor=_BG)
        lb.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        lb.pack(side="left", fill="both", expand=True)

        status_lbl = tk.Label(main, text="Loading recent tickets…", bg=_BG,
                              fg="#888888", font=("Segoe UI", 8, "italic"))
        status_lbl.pack(anchor="w", pady=(0, 6))

        _results: list = []   # [{key, summary, status, type}]
        _after_id = [None]
        _search_gen = [0]     # generation counter to discard stale results

        existing_children_keys = set()
        for c in getattr(self, "_epic_children_data", []):
            k = c.get("key", "").strip()
            if k:
                existing_children_keys.add(k)

        def _build_jql(q):
            """Build a JQL string that searches by key, summary, and text."""
            base_parts = ["issuetype != Epic"]
            if project_key:
                base_parts.insert(0, f'project = "{project_key}"')
            base = " AND ".join(base_parts)

            if not q:
                return f"{base} ORDER BY updated DESC"

            clauses = []
            q_escaped = q.replace('"', '\\"')

            # If purely numeric, match the ticket number portion of the key
            if q.isdigit() and project_key:
                clauses.append(f'key = "{project_key}-{q}"')

            # If it looks like a partial or full issue key (e.g. "PROJ-123")
            if "-" in q:
                clauses.append(f'key = "{q_escaped}"')

            # Always try summary and text search
            clauses.append(f'summary ~ "{q_escaped}"')
            clauses.append(f'text ~ "{q_escaped}"')

            search_clause = " OR ".join(clauses)
            return f"{base} AND ({search_clause}) ORDER BY updated DESC"

        def _do_search():
            q = search_var.get().strip()
            _search_gen[0] += 1
            gen = _search_gen[0]
            status_lbl.config(text="Searching…" if q else "Loading…")

            import threading

            def _worker():
                try:
                    jql = _build_jql(q)
                    from jira_api import perform_jira_request
                    url = f"{session._jira_base}/rest/api/3/search/jql"
                    body = {
                        "jql": jql,
                        "maxResults": 50,
                        "fields": ["summary", "status", "issuetype"],
                    }
                    resp = perform_jira_request(session, "POST", url,
                                                json_body=body, timeout=30)
                    resp.raise_for_status()
                    data = resp.json()
                    issues = data.get("issues", [])
                    hits = []
                    for iss in issues:
                        k = iss.get("key", "")
                        if k == epic_key:
                            continue
                        f = iss.get("fields") or {}
                        hits.append({
                            "key":     k,
                            "summary": f.get("summary", ""),
                            "status":  (f.get("status") or {}).get("name", ""),
                            "type":    (f.get("issuetype") or {}).get("name", ""),
                        })
                    if gen == _search_gen[0]:
                        win.after(0, lambda h=hits, t=len(issues): _show_results(h, t))
                except Exception:
                    debug_log("Child picker search failed: " + traceback.format_exc())
                    if gen == _search_gen[0]:
                        win.after(0, lambda: status_lbl.config(
                            text="Search failed — see debug log."))

            threading.Thread(target=_worker, daemon=True).start()

        # Load default list immediately on open
        _do_search()

        def _show_results(hits, total):
            _results.clear()
            _results.extend(hits)
            lb.delete(0, tk.END)
            for h in hits:
                already = " (already child)" if h["key"] in existing_children_keys else ""
                line = f"  {h['key']}  [{h['type']}]  {h['status']}  —  {h['summary']}{already}"
                lb.insert(tk.END, line)
            q = search_var.get().strip()
            if q:
                status_lbl.config(text=f"{len(hits)} result(s) for \"{q}\" (of {total} matched)")
            else:
                status_lbl.config(text=f"{len(hits)} recent ticket(s) shown")

        def _on_search_change(*_):
            if _after_id[0]:
                win.after_cancel(_after_id[0])
            _after_id[0] = win.after(400, _do_search)

        search_var.trace_add("write", _on_search_change)

        def _confirm(*_):
            sel = lb.curselection()
            if not sel:
                return
            chosen = [_results[i] for i in sel if i < len(_results)]
            if not chosen:
                return
            chosen = [c for c in chosen if c["key"] not in existing_children_keys]
            if not chosen:
                from tkinter import messagebox
                messagebox.showinfo("Already linked",
                                    "All selected tickets are already children of this epic.",
                                    parent=win)
                return

            status_lbl.config(text=f"Assigning {len(chosen)} ticket(s) as children…")
            for w in (search_ent, lb):
                w.config(state="disabled")

            import threading

            def _assign_worker():
                epic_mode = getattr(self, "_stored_epic_mode", "nextgen")
                successes = []
                failures = []
                for ticket in chosen:
                    child_key = ticket["key"]
                    try:
                        if epic_mode == "classic":
                            payload = {"fields": {"customfield_10014": epic_key}}
                        else:
                            payload = {"fields": {"parent": {"key": epic_key}}}
                        from jira_api import perform_jira_request
                        url = f"{session._jira_base}/rest/api/3/issue/{child_key}"
                        resp = perform_jira_request(session, "PUT", url,
                                                    json_body=payload, timeout=30)
                        if resp.status_code in (200, 204):
                            successes.append(ticket)
                            debug_log(f"Assigned {child_key} as child of {epic_key}")
                        else:
                            failures.append((child_key, f"{resp.status_code} {resp.text[:200]}"))
                            debug_log(f"Failed to assign {child_key}: {resp.status_code} {resp.text[:200]}")
                    except Exception:
                        failures.append((child_key, traceback.format_exc()[:200]))
                        debug_log(f"Exception assigning {child_key}: {traceback.format_exc()}")

                win.after(0, lambda: _on_assign_done(successes, failures))

            def _on_assign_done(successes, failures):
                if failures:
                    from tkinter import messagebox
                    fail_lines = "\n".join(f"  {k}: {e}" for k, e in failures)
                    messagebox.showwarning(
                        "Some assignments failed",
                        f"Failed to link {len(failures)} ticket(s):\n{fail_lines}",
                        parent=win)

                if successes:
                    new_children = []
                    for t in successes:
                        new_children.append({
                            "key":     t["key"],
                            "summary": t["summary"],
                            "status":  t["status"],
                        })
                        existing_children_keys.add(t["key"])

                    current = getattr(self, "_epic_children_data", [])
                    current.extend(new_children)
                    self._epic_children_data = current
                    self._refresh_epic_view()

                    self._fetch_missing_children_locally(
                        session, epic_key, [t["key"] for t in successes])

                win.destroy()

            threading.Thread(target=_assign_worker, daemon=True).start()

        btn_row = tk.Frame(main, bg=_BG)
        btn_row.pack(fill="x")
        tk.Button(btn_row, text="Assign as Children", command=_confirm,
                  bg="#0e639c", fg="#ffffff", relief="flat", bd=0,
                  font=("Segoe UI", 9, "bold"), padx=14, pady=5,
                  cursor="hand2", activebackground="#1177bb").pack(side="right", padx=(4, 0))
        tk.Button(btn_row, text="Cancel", command=win.destroy,
                  bg="#3c3c3c", fg=_FG, relief="flat", bd=0,
                  font=("Segoe UI", 9), padx=14, pady=5,
                  cursor="hand2").pack(side="right")

        lb.bind("<Double-Button-1>", _confirm)
        win.bind("<Escape>", lambda e: win.destroy())
        search_ent.focus_set()

    def _fetch_missing_children_locally(self, session, epic_key, child_keys):
        """Background-fetch full details for newly assigned children
        and add them to list_items if not already present."""
        import threading

        root = self.frame.winfo_toplevel()

        def _worker():
            try:
                existing_keys = set()
                if hasattr(root, "list_items"):
                    existing_keys = {
                        str(it.get("Issue key") or "").strip()
                        for it in root.list_items if it.get("Issue key")
                    }

                new_items = []
                for ck in child_keys:
                    if ck in existing_keys:
                        continue
                    try:
                        from config import FETCH_FIELDS
                        issue_json = root.fetch_issue_details(
                            session, ck, fields=FETCH_FIELDS)
                        issue_dict = root._map_issue_json_to_dict(issue_json)
                        new_items.append(issue_dict)
                        debug_log(f"Fetched newly assigned child {ck}")
                    except Exception:
                        debug_log(f"Failed to fetch child {ck}: "
                                  + traceback.format_exc())

                if new_items:
                    def _add():
                        try:
                            root.list_items.extend(new_items)
                            from utils import _dedup_list_items
                            root.list_items = _dedup_list_items(root.list_items)
                            root.meta["fetched_issues"] = list(root.list_items)
                            from storage import save_storage
                            save_storage(root.templates, root.meta)
                            root._rebuild_list_view()
                        except Exception:
                            debug_log("Failed to persist children: "
                                      + traceback.format_exc())
                    root.after(0, _add)
            except Exception:
                debug_log("_fetch_missing_children_locally failed: "
                          + traceback.format_exc())

        threading.Thread(target=_worker, daemon=True).start()

    def _load_epic_from_data(self, data: dict):
        """Populate epic vars from a data dict and refresh the view."""
        if hasattr(self, "_epic_link_var"):
            self._epic_link_var.set(data.get("Epic Link") or "")
        if hasattr(self, "_epic_name_var"):
            self._epic_name_var.set(data.get("Epic Name") or "")
        # Remember which Jira mechanism was used (classic vs next-gen)
        self._stored_epic_mode = data.get("_epic_mode") or "nextgen"
        # Load child issues (display-only).
        # Preserve any children already fetched by _auto_fetch_epic_children
        # if the incoming data has no children (avoids race-condition wipe).
        children_raw = data.get("Epic Children")
        incoming = []
        try:
            if isinstance(children_raw, list) and children_raw:
                incoming = children_raw
            elif isinstance(children_raw, str) and children_raw.strip() and children_raw.strip() != "[]":
                incoming = json.loads(children_raw)
        except Exception:
            incoming = []
        if incoming:
            self._epic_children_data = incoming
        elif not getattr(self, "_epic_children_data", None):
            self._epic_children_data = []
        self._refresh_epic_view()

    # ── Issue Links section ───────────────────────────────────────────────────

    def _build_issue_links_section(self):
        """Build the Issue Links UI: scrollable list + add-link form."""
        _BG      = "#1a1a1a"
        _PANEL   = "#252526"
        _BORDER  = "#3c3c3c"
        _FG      = "#d4d4d4"
        _META_FG = "#888888"
        _NEW_BG  = "#1e1e1e"

        outer = tk.Frame(self._content, bg=_PANEL, bd=1, relief="flat",
                         highlightthickness=1, highlightbackground=_BORDER)
        outer.pack(fill="x", padx=4, pady=(2, 4))

        # Header
        hdr_bar = tk.Frame(outer, bg=_PANEL, pady=5, padx=10)
        hdr_bar.pack(fill="x")
        tk.Label(hdr_bar, text="🔗  Issue Links", bg=_PANEL, fg=_FG,
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        self._issue_link_count_lbl = tk.Label(hdr_bar, text="", bg=_PANEL,
                                              fg=_META_FG, font=("Segoe UI", 9))
        self._issue_link_count_lbl.pack(side="left", padx=(6, 0))
        tk.Frame(outer, bg=_BORDER, height=1).pack(fill="x")

        # Scrollable list
        list_frame = tk.Frame(outer, bg=_BG)
        list_frame.pack(fill="x")
        self._issue_links_canvas = tk.Canvas(list_frame, bg=_BG,
                                             highlightthickness=0, height=120)
        vsb = tk.Scrollbar(list_frame, orient="vertical",
                           command=self._issue_links_canvas.yview,
                           bg=_PANEL, troughcolor=_BG, width=10)
        self._issue_links_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._issue_links_canvas.pack(side="left", fill="both", expand=True)
        self._issue_links_inner = tk.Frame(self._issue_links_canvas, bg=_BG)
        self._issue_links_win = self._issue_links_canvas.create_window(
            (0, 0), window=self._issue_links_inner, anchor="nw")

        def _on_inner_cfg(e):
            self._issue_links_canvas.configure(
                scrollregion=self._issue_links_canvas.bbox("all"))
        def _on_canvas_cfg(e):
            self._issue_links_canvas.itemconfig(
                self._issue_links_win, width=e.width)
        self._issue_links_inner.bind("<Configure>", _on_inner_cfg)
        self._issue_links_canvas.bind("<Configure>", _on_canvas_cfg)

        # Divider + add-link form
        tk.Frame(outer, bg=_BORDER, height=1).pack(fill="x", pady=(4, 0))
        add_frame = tk.Frame(outer, bg=_PANEL, padx=10, pady=8)
        add_frame.pack(fill="x")
        tk.Label(add_frame, text="Add link:", bg=_PANEL, fg=_FG,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 4))

        form_row = tk.Frame(add_frame, bg=_PANEL)
        form_row.pack(fill="x")

        _LINK_TYPES = [
            "blocks", "is blocked by",
            "clones", "is cloned by",
            "duplicates", "is duplicated by",
            "relates to",
        ]
        self._new_link_type_var = tk.StringVar(value="relates to")
        type_combo = ttk.Combobox(form_row, textvariable=self._new_link_type_var,
                                  values=_LINK_TYPES, width=18, state="readonly")
        type_combo.pack(side="left", padx=(0, 6))

        self._new_link_key_var = tk.StringVar()
        key_ent = tk.Entry(form_row, textvariable=self._new_link_key_var,
                           bg=_NEW_BG, fg=_FG, insertbackground=_FG,
                           font=("Segoe UI", 10), relief="flat", bd=1,
                           highlightthickness=1, highlightbackground=_BORDER,
                           width=16)
        key_ent.pack(side="left", padx=(0, 6), ipady=3)

        add_btn = tk.Button(form_row, text="Add", command=self._add_issue_link,
                            bg="#0e639c", fg="#ffffff", font=("Segoe UI", 9, "bold"),
                            relief="flat", bd=0, padx=12, pady=4, cursor="hand2",
                            activebackground="#1177bb", activeforeground="#ffffff")
        add_btn.bind("<Enter>", lambda e: add_btn.configure(bg="#1177bb"))
        add_btn.bind("<Leave>", lambda e: add_btn.configure(bg="#0e639c"))
        add_btn.pack(side="left")
        key_ent.bind("<Return>", lambda e: self._add_issue_link())

        self.field_widgets["Issue Links"] = {
            "widget": None, "var": None,
            "include_var": tk.BooleanVar(value=True),
            "row": outer, "label": None, "adf_row": None,
        }
        self._issue_links_section_outer = outer
        self._issue_links_data: list = []
        self._refresh_issue_links_view()

    def _refresh_issue_links_view(self):
        """Redraw the issue links list from self._issue_links_data."""
        _BG      = "#1a1a1a"
        _PANEL   = "#252526"
        _BORDER  = "#3c3c3c"
        _FG      = "#d4d4d4"
        _META_FG = "#888888"
        _PENDING = "#4ec9b0"

        for w in self._issue_links_inner.winfo_children():
            w.destroy()

        if not self._issue_links_data:
            tk.Label(self._issue_links_inner, text="No issue links.",
                     bg=_BG, fg=_META_FG,
                     font=("Segoe UI", 9, "italic"), padx=12, pady=8
                     ).pack(anchor="w")
        else:
            for i, lnk in enumerate(self._issue_links_data):
                if i:
                    tk.Frame(self._issue_links_inner, bg=_BORDER, height=1).pack(fill="x", padx=8)
                row = tk.Frame(self._issue_links_inner, bg=_BG, padx=12, pady=5)
                row.pack(fill="x")
                # Direction label badge
                dir_label = lnk.get("direction_label") or lnk.get("type_name") or "relates to"
                dir_badge = tk.Label(row, text=f" {dir_label} ",
                                     bg="#2a3a4a", fg="#7eb8d4",
                                     font=("Segoe UI", 8), relief="flat")
                dir_badge.pack(side="left", padx=(0, 8))
                # Key (clickable)
                key = lnk.get("key") or ""
                key_lbl = tk.Label(row, text=key, bg=_BG, fg="#4a9eff",
                                   font=("Segoe UI", 9, "bold"), cursor="hand2")
                key_lbl.pack(side="left")
                key_lbl.bind("<Button-1>", lambda e, k=key: self._show_ticket_link_menu_at(k, e))
                key_lbl.bind("<Button-3>", lambda e, k=key: self._show_ticket_link_menu_at(k, e))
                # Summary
                summary = lnk.get("summary") or ""
                if summary:
                    tk.Label(row, text=f"  {summary}", bg=_BG, fg=_FG,
                             font=("Segoe UI", 9)).pack(side="left")
                # Status
                status = lnk.get("status") or ""
                if status:
                    tk.Label(row, text=f"  [{status}]", bg=_BG, fg=_META_FG,
                             font=("Segoe UI", 8)).pack(side="left")
                # Pending badge
                if not lnk.get("posted", True):
                    tk.Label(row, text="  ● pending", bg=_BG, fg=_PENDING,
                             font=("Segoe UI", 8)).pack(side="left")
                # Remove button (local unposted only)
                if not lnk.get("posted", True):
                    idx_capture = i
                    rm_btn = tk.Button(row, text="×", width=2,
                                       bg=_BG, fg=_META_FG, relief="flat", bd=0,
                                       font=("Segoe UI", 9), cursor="hand2",
                                       command=lambda ix=idx_capture: self._remove_issue_link(ix))
                    rm_btn.pack(side="right")

        count = len(self._issue_links_data)
        try:
            self._issue_link_count_lbl.config(
                text=f"({count})" if count else "")
        except Exception:
            pass
        self._rebind_scroll()

    def _add_issue_link(self):
        """Add a new local (unposted) issue link."""
        link_type = self._new_link_type_var.get().strip()
        key       = self._new_link_key_var.get().strip().upper()
        if not key:
            return
        # Normalise direction: if the label starts with "is " it's typically inward
        direction = "inward" if link_type.startswith("is ") else "outward"
        self._issue_links_data.append({
            "id":              "",
            "type_name":       link_type,
            "direction":       direction,
            "direction_label": link_type,
            "key":             key,
            "summary":         "",
            "status":          "",
            "posted":          False,
        })
        self._new_link_key_var.set("")
        self._refresh_issue_links_view()

    def _remove_issue_link(self, idx: int):
        """Remove a locally-added (unposted) issue link by index."""
        try:
            del self._issue_links_data[idx]
        except IndexError:
            pass
        self._refresh_issue_links_view()

    def _get_issue_links_json(self) -> str:
        try:
            return json.dumps(self._issue_links_data, ensure_ascii=False)
        except Exception:
            return "[]"

    def _load_issue_links_from_json(self, raw):
        try:
            if isinstance(raw, list):
                self._issue_links_data = raw
            elif isinstance(raw, str) and raw.strip():
                parsed = json.loads(raw)
                self._issue_links_data = parsed if isinstance(parsed, list) else []
            else:
                self._issue_links_data = []
        except Exception:
            self._issue_links_data = []
        self._refresh_issue_links_view()

    # ── Comments section ──────────────────────────────────────────────────────

    def _build_comments_section(self):
        """Build the Comments UI: read-only thread view + new-comment input."""
        _BG       = "#1a1a1a"
        _PANEL    = "#252526"
        _BORDER   = "#3c3c3c"
        _FG       = "#d4d4d4"
        _META_FG  = "#888888"
        _NEW_BG   = "#1e1e1e"

        outer = tk.Frame(self._content, bg=_PANEL, bd=1, relief="flat",
                         highlightthickness=1, highlightbackground=_BORDER)
        outer.pack(fill="x", padx=4, pady=(2, 8))

        # ── Header bar ──
        hdr_bar = tk.Frame(outer, bg=_PANEL, pady=6, padx=10)
        hdr_bar.pack(fill="x")
        tk.Label(hdr_bar, text="💬  Comments", bg=_PANEL, fg=_FG,
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        self._comment_count_lbl = tk.Label(hdr_bar, text="", bg=_PANEL,
                                           fg=_META_FG, font=("Segoe UI", 9))
        self._comment_count_lbl.pack(side="left", padx=(6, 0))
        tk.Frame(outer, bg=_BORDER, height=1).pack(fill="x")

        # ── Scrollable comment thread ──
        thread_frame = tk.Frame(outer, bg=_BG)
        thread_frame.pack(fill="x")

        self._comments_canvas = tk.Canvas(thread_frame, bg=_BG,
                                          highlightthickness=0, height=180)
        vsb = tk.Scrollbar(thread_frame, orient="vertical",
                           command=self._comments_canvas.yview,
                           bg=_PANEL, troughcolor=_BG, width=10)
        self._comments_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._comments_canvas.pack(side="left", fill="both", expand=True)

        self._comments_inner = tk.Frame(self._comments_canvas, bg=_BG)
        self._comments_window = self._comments_canvas.create_window(
            (0, 0), window=self._comments_inner, anchor="nw")

        def _on_inner_cfg(e):
            self._comments_canvas.configure(
                scrollregion=self._comments_canvas.bbox("all"))
        def _on_canvas_cfg(e):
            self._comments_canvas.itemconfig(
                self._comments_window, width=e.width)
        self._comments_inner.bind("<Configure>", _on_inner_cfg)
        self._comments_canvas.bind("<Configure>", _on_canvas_cfg)

        # ── Divider ──
        tk.Frame(outer, bg=_BORDER, height=1).pack(fill="x", pady=(4, 0))

        # ── New-comment input ──
        new_frame = tk.Frame(outer, bg=_PANEL, padx=10, pady=8)
        new_frame.pack(fill="x")
        tk.Label(new_frame, text="Add a comment:", bg=_PANEL, fg=_FG,
                 font=("Segoe UI", 9)).pack(anchor="w")
        input_frame = tk.Frame(new_frame, bg=_NEW_BG, bd=1, relief="flat",
                               highlightthickness=1, highlightbackground=_BORDER)
        input_frame.pack(fill="x", pady=(4, 6))
        self._new_comment_txt = tk.Text(
            input_frame, height=3, wrap="word",
            bg=_NEW_BG, fg=_FG, insertbackground=_FG,
            font=("Segoe UI", 10), relief="flat", bd=0, padx=8, pady=6)
        self._new_comment_txt.pack(fill="both", expand=True)

        btn_row = tk.Frame(new_frame, bg=_PANEL)
        btn_row.pack(fill="x")
        add_btn = tk.Button(
            btn_row, text="Add Comment", command=self._add_comment,
            bg="#0e639c", fg="#ffffff", font=("Segoe UI", 9, "bold"),
            relief="flat", bd=0, padx=14, pady=5, cursor="hand2",
            activebackground="#1177bb", activeforeground="#ffffff")
        add_btn.bind("<Enter>", lambda e: add_btn.configure(bg="#1177bb"))
        add_btn.bind("<Leave>", lambda e: add_btn.configure(bg="#0e639c"))
        add_btn.pack(side="right")

        # Store in field_widgets so populate_from_dict / read_to_dict work
        self.field_widgets["Comment"] = {
            "widget": None, "var": None,
            "include_var": tk.BooleanVar(value=True),
            "add_btn": None, "row": outer, "label": None, "adf_row": None,
        }
        self._comments_data: list = []   # list of dicts: author/date/body/posted
        self._refresh_comments_view()

    def _refresh_comments_view(self):
        """Redraw the comment thread from self._comments_data."""
        _BG      = "#1a1a1a"
        _PANEL   = "#252526"
        _BORDER  = "#3c3c3c"
        _FG      = "#d4d4d4"
        _META_FG = "#888888"
        _NEW_FG  = "#4ec9b0"     # teal badge for unsent comments

        # Clear existing widgets
        for w in self._comments_inner.winfo_children():
            w.destroy()

        if not self._comments_data:
            tk.Label(self._comments_inner, text="No comments yet.",
                     bg=_BG, fg=_META_FG,
                     font=("Segoe UI", 9, "italic"), padx=12, pady=10
                     ).pack(anchor="w")
        else:
            for i, c in enumerate(self._comments_data):
                if i:
                    tk.Frame(self._comments_inner, bg=_BORDER,
                             height=1).pack(fill="x", padx=8)
                item = tk.Frame(self._comments_inner, bg=_BG, pady=6, padx=12)
                item.pack(fill="x")
                # Author + date row
                meta = tk.Frame(item, bg=_BG)
                meta.pack(fill="x")
                author = c.get("author") or "Unknown"
                date_raw = c.get("date") or ""
                # Shorten ISO date to "YYYY-MM-DD HH:MM"
                date_short = date_raw[:16].replace("T", "  ") if date_raw else ""
                tk.Label(meta, text=author, bg=_BG, fg=_FG,
                         font=("Segoe UI", 9, "bold")).pack(side="left")
                if date_short:
                    tk.Label(meta, text=f"  {date_short}", bg=_BG,
                             fg=_META_FG, font=("Segoe UI", 9)).pack(side="left")
                if not c.get("posted", True):
                    tk.Label(meta, text="  ● pending upload", bg=_BG,
                             fg=_NEW_FG, font=("Segoe UI", 8)).pack(side="left")
                # Body
                body = c.get("body") or ""
                body_lbl = tk.Label(item, text=body, bg=_BG, fg=_FG,
                                    font=("Segoe UI", 9), wraplength=560,
                                    justify="left", anchor="w")
                body_lbl.pack(fill="x", pady=(3, 0))

        count = len(self._comments_data)
        try:
            self._comment_count_lbl.config(
                text=f"({count})" if count else "")
        except Exception:
            pass
        self._rebind_scroll()

    def _add_comment(self):
        """Append a new local comment and refresh the view."""
        txt = self._new_comment_txt.get("1.0", "end").strip()
        if not txt:
            return
        import datetime
        now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        # Use the logged-in Jira display name if available
        author = "Me"
        try:
            user_info = getattr(self, "_app_ref", None)
            if user_info is None:
                # Walk up to app via frame
                root = self.frame.winfo_toplevel()
                if hasattr(root, "meta"):
                    author = (root.meta.get("jira_current_user") or {}).get(
                        "displayName") or "Me"
        except Exception:
            pass
        self._comments_data.append({
            "id": "",
            "author": author,
            "date": now,
            "body": txt,
            "posted": False,
        })
        self._new_comment_txt.delete("1.0", "end")
        self._refresh_comments_view()

    def _get_comments_json(self) -> str:
        """Serialise current comments list to JSON string for storage."""
        try:
            return json.dumps(self._comments_data, ensure_ascii=False)
        except Exception:
            return "[]"

    def _load_comments_from_json(self, raw):
        """Load comments from the stored JSON string or list."""
        try:
            if isinstance(raw, list):
                self._comments_data = raw
            elif isinstance(raw, str) and raw.strip():
                parsed = json.loads(raw)
                self._comments_data = parsed if isinstance(parsed, list) else []
            else:
                self._comments_data = []
        except Exception:
            self._comments_data = []
        self._refresh_comments_view()

    def _add_option(self, h):
        """
        Delegate to the app-level add_option callback if provided.
        """
        try:
            if self.add_option_cb:
                self.add_option_cb(self, h)
        except Exception:
            debug_log("Add option failed: " + traceback.format_exc())

    def _on_refresh_assignees(self):
        """Fetch assignable users from Jira and populate the Assignee dropdown."""
        if self.fetch_assignees_cb:
            self.fetch_assignees_cb(self)

    def _on_refresh_field_options(self, field_name):
        """Fetch options for a field from Jira and populate the dropdown."""
        if self.fetch_options_cb:
            self.fetch_options_cb(self, field_name)

    def set_option_values(self, field, values):
        """
        Set the combobox values for a field (used to update saved options).
        Also resets the autocomplete list so the fresh options are used next time.
        """
        try:
            info = self.field_widgets.get(field)
            if not info:
                return
            widget = info.get("widget")
            if isinstance(widget, ttk.Combobox):
                try:
                    widget['values'] = values
                except Exception:
                    widget.configure(values=values)
        except Exception:
            debug_log("set_option_values failed: " + traceback.format_exc())

    # Keys that should not trigger autocomplete filtering
    _AUTOCOMPLETE_SKIP = frozenset({
        "Return", "KP_Enter", "Escape", "Tab",
        "Up", "Down", "Left", "Right",
        "Home", "End", "Prior", "Next",
        "Shift_L", "Shift_R", "Control_L", "Control_R",
        "Alt_L", "Alt_R", "Caps_Lock", "Super_L", "Super_R",
        "F1", "F2", "F3", "F4", "F5", "F6",
        "F7", "F8", "F9", "F10", "F11", "F12",
    })

    def _setup_combo_autocomplete(self, combo: "ttk.Combobox", hdr: str,
                                   var: "tk.StringVar") -> None:
        """Wire live-filter autocomplete onto a ttk.Combobox.

        As the user types, the dropdown is filtered to entries that contain the
        typed text (case-insensitive substring match) and the popup opens
        automatically to show matches.  Selecting an option or pressing Escape
        restores the full list so the next interaction starts clean.
        """
        _after_id: list = [None]

        def _filter() -> None:
            _after_id[0] = None
            all_opts = list(self.meta_options.get(hdr, []))
            q = var.get().lower().strip()

            try:
                cursor_pos = combo.index(tk.INSERT)
            except Exception:
                cursor_pos = tk.END

            if q:
                filtered = [v for v in all_opts if q in v.lower()]
                combo["values"] = filtered if filtered else all_opts
            else:
                combo["values"] = all_opts

            try:
                combo.icursor(cursor_pos)
                combo.selection_clear()
            except Exception:
                pass

        def _on_key(event) -> None:
            if event.keysym in self._AUTOCOMPLETE_SKIP:
                return
            if _after_id[0]:
                combo.after_cancel(_after_id[0])
            _after_id[0] = combo.after(150, _filter)

        def _on_selected(event=None) -> None:
            # Restore full list after a selection so the next open shows everything
            combo["values"] = list(self.meta_options.get(hdr, []))

        def _on_escape(event=None) -> None:
            combo["values"] = list(self.meta_options.get(hdr, []))

        combo.bind("<KeyRelease>", _on_key)
        combo.bind("<<ComboboxSelected>>", _on_selected)
        combo.bind("<Escape>", _on_escape)

    def _on_internal_priority_changed(self):
        if self.internal_priority_set_cb and self._last_ticket_key:
            try:
                val = self._info_internal_var.get() or "None"
                self.internal_priority_set_cb(self._last_ticket_key, val)
            except Exception:
                pass

    def read_to_dict(self):
        """
        Read the current tab fields into a dict. Attempts to preserve Description ADF as JSON object if present.
        """
        if hasattr(self, "_pre_read_hook") and self._pre_read_hook:
            try:
                self._pre_read_hook(self)
            except Exception:
                pass
        out = {}
        try:
            for h in HEADERS:
                # Comments use a custom widget; serialise from internal list
                if h == "Comment":
                    out[h] = self._get_comments_json() if hasattr(self, "_comments_data") else "[]"
                    continue
                # Epic fields come from the dedicated epic section
                if h == "Epic Link":
                    out[h] = getattr(self, "_epic_link_var", tk.StringVar()).get()
                    out["_epic_mode"] = getattr(self, "_stored_epic_mode", "nextgen")
                    continue
                if h == "Epic Name":
                    out[h] = getattr(self, "_epic_name_var", tk.StringVar()).get()
                    continue
                if h == "Epic Children":
                    out[h] = ""   # display-only, not stored
                    continue
                # Issue Links from the custom section
                if h == "Issue Links":
                    out[h] = self._get_issue_links_json() if hasattr(self, "_issue_links_data") else "[]"
                    continue
                info = self.field_widgets.get(h)
                if not info:
                    if h == "Description":
                        adf_info = self.field_widgets.get("Description ADF")
                        if adf_info and self.extract_text_from_adf_cb:
                            try:
                                w = adf_info.get("widget")
                                raw = w.get("1.0", "end").strip() if w else ""
                                node = json.loads(raw) if raw else None
                                out[h] = self.extract_text_from_adf_cb(node) or "" if node else ""
                            except Exception:
                                out[h] = ""
                        else:
                            out[h] = ""
                    else:
                        out[h] = ""
                    continue
                widget = info.get("widget")
                if isinstance(widget, tk.Text):
                    try:
                        text = widget.get("1.0", "end").rstrip("\n")
                    except Exception:
                        text = ""
                    if h == "Description ADF":
                        try:
                            node = json.loads(text) if text else None
                            out[h] = node if node is not None else ""
                        except Exception:
                            out[h] = text
                    else:
                        out[h] = text
                elif isinstance(widget, ttk.Combobox):
                    try:
                        widget_val = widget.get()
                        if h == "Attachment":
                            raw = getattr(self, "_attachment_raw_json", None)
                            if raw:
                                try:
                                    items = json.loads(raw)
                                    display = "; ".join((x.get("filename") or x.get("name") or "") for x in (items if isinstance(items, list) else []) if isinstance(x, dict))
                                    out[h] = raw if widget_val == display else widget_val
                                    if widget_val != display:
                                        setattr(self, "_attachment_raw_json", None)
                                except Exception:
                                    out[h] = widget_val
                            else:
                                out[h] = widget_val
                        else:
                            out[h] = widget_val
                    except Exception:
                        out[h] = ""
                else:
                    try:
                        # Entry-like widgets
                        out[h] = widget.get()
                    except Exception:
                        out[h] = ""
        except Exception:
            debug_log("read_to_dict failed: " + traceback.format_exc())
        return out

    def check_all(self):
        """
        Check (include) all fields in this tab.
        """
        try:
            for info in self.field_widgets.values():
                iv = info.get("include_var")
                if iv is not None:
                    try:
                        iv.set(True)
                    except Exception:
                        pass
        except Exception:
            debug_log("check_all failed: " + traceback.format_exc())

    def uncheck_all(self):
        """
        Uncheck (exclude) all fields in this tab (except always-on fields like Description ADF).
        """
        try:
            for hdr, info in self.field_widgets.items():
                if hdr == "Description ADF":
                    continue
                iv = info.get("include_var")
                if iv is not None:
                    try:
                        iv.set(False)
                    except Exception:
                        pass
        except Exception:
            debug_log("uncheck_all failed: " + traceback.format_exc())

    def collapse_unincluded(self, collapse_on=False, filter_q=""):
        """
        Show/hide field rows based on include_var and a text filter.
        If collapse_on is True, rows where include_var is False will be hidden (pack_forget).
        filter_q (lowercase) will cause fields to be shown if header or field value contains the query.
        """
        try:
            q = (filter_q or "").strip().lower()
            for hdr, info in self.field_widgets.items():
                row = info.get("row")
                adf_row = info.get("adf_row")
                include_var = info.get("include_var")
                widget = info.get("widget")
                # Determine if this row matches filter
                matches_filter = False
                if q:
                    if q in (hdr or "").lower():
                        matches_filter = True
                    else:
                        try:
                            # extract widget text for searching
                            if isinstance(widget, tk.Text):
                                txt = widget.get("1.0", "end").strip().lower()
                                if q in txt:
                                    matches_filter = True
                            elif isinstance(widget, ttk.Combobox):
                                vv = widget.get() or ""
                                if q in str(vv).lower():
                                    matches_filter = True
                            else:
                                try:
                                    vv = widget.get()
                                    if q in str(vv).lower():
                                        matches_filter = True
                                except Exception:
                                    pass
                        except Exception:
                            pass
                # Determine visibility
                should_hide = False
                if collapse_on:
                    try:
                        included = bool(include_var.get())
                    except Exception:
                        included = True
                    if not included and not matches_filter:
                        should_hide = True
                # Apply visibility to main row
                try:
                    if row:
                        if should_hide:
                            row.pack_forget()
                        else:
                            # if not currently visible, re-pack it
                            # Use same pack options as original build_fields
                            try:
                                row.pack(fill="x", padx=4, pady=4)
                            except Exception:
                                pass
                    # For adf_row (the larger editor pane), keep it visible if its field visible
                    if adf_row:
                        if should_hide:
                            adf_row.pack_forget()
                        else:
                            try:
                                adf_row.pack(fill="both", padx=4, pady=(0, 6), expand=True)
                            except Exception:
                                pass
                except Exception:
                    debug_log("collapse_unincluded row pack error: " + traceback.format_exc())
        except Exception:
            debug_log("collapse_unincluded failed: " + traceback.format_exc())
