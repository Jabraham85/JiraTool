"""
Variable helpers: define variables, live preview, apply to ticket/ADF.
"""
import re
import json
import copy
import tkinter as tk
from tkinter import ttk, messagebox

from .utils import debug_log


class VariablesMixin:
    """Mixin providing variable definition, preview, and application for tickets."""

    # Per-ticket inline variable definition:  {KEY=value}
    # Reference (used anywhere in ticket):    {KEY}
    # On upload: definitions → keep only the value, references → replaced with value
    _VAR_DEF_RE = re.compile(r'\{([A-Z])=([^}]+)\}')
    _VAR_REF_RE = re.compile(r'\{([A-Za-z][A-Za-z0-9_]*)\}')

    # ── Live variable preview in tk.Text ──────────────────────────────
    _VAR_PREVIEW_TAG = "var_preview"
    _VAR_EDITING_TAG = "var_editing"

    def _snapshot_var_selection(self):
        """Capture the currently-focused widget's selection before focus moves (call on ButtonPress)."""
        self._var_pending_sel = None
        focused = self.focus_get()
        if focused is None:
            return
        self._snapshot_var_selection_from(focused)

    def _snapshot_var_selection_from(self, widget):
        """Capture selection from a specific widget and store in _var_pending_sel."""
        self._var_pending_sel = None
        try:
            if isinstance(widget, tk.Text):
                sel_text = widget.get("sel.first", "sel.last")
                if sel_text.strip():
                    self._var_pending_sel = (
                        "text", widget,
                        widget.index("sel.first"),
                        widget.index("sel.last"),
                        sel_text,
                    )
            elif isinstance(widget, (tk.Entry, ttk.Combobox)):
                if widget.selection_present():
                    start = widget.index("sel.first")
                    end = widget.index("sel.last")
                    sel_text = widget.selection_get()
                    if sel_text.strip():
                        self._var_pending_sel = ("entry", widget, start, end, sel_text)
        except tk.TclError:
            pass
        # Also try tkinterweb HtmlText / HtmlFrame via selection_manager (preferred) or get_selection
        if self._var_pending_sel is None:
            try:
                sm = getattr(widget, "selection_manager", None)
                sel_text = sm.get_selection() if sm else widget.get_selection()
                if sel_text and sel_text.strip():
                    self._var_pending_sel = ("html", widget, None, None, sel_text)
            except (AttributeError, Exception):
                pass
        # Walk up to parent widget if the click was on a child frame of the html widget
        if self._var_pending_sel is None:
            try:
                parent = widget.master
                sm = getattr(parent, "selection_manager", None)
                if parent and (sm or hasattr(parent, "get_selection")):
                    sel_text = sm.get_selection() if sm else parent.get_selection()
                    if sel_text and sel_text.strip():
                        self._var_pending_sel = ("html", parent, None, None, sel_text)
            except (AttributeError, Exception):
                pass

    def _collect_defined_var_keys(self):
        """Scan all text widgets of the active tab and return dict {KEY: value} for existing {KEY=value} defs."""
        found = {}
        tf = self.get_active_tabform()
        if not tf:
            return found
        for fw in tf.field_widgets.values():
            w = fw.get("widget")
            if w is None:
                continue
            try:
                if isinstance(w, tk.Text):
                    content = w.get("1.0", "end")
                elif isinstance(w, (tk.Entry, ttk.Combobox)):
                    content = w.get()
                else:
                    continue
            except tk.TclError:
                continue
            for m in self._VAR_DEF_RE.finditer(content):
                found[m.group(1).upper()] = m.group(2).strip()
        return found

    def _next_var_letter(self):
        """Return the next unused uppercase letter for auto-assignment (A, B, C, ...)."""
        used = set(self._collect_defined_var_keys().keys())
        for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            if c not in used:
                return c
        return "V" + str(len(used) + 1)

    def define_variable_dialog(self):
        """Highlight text → right-click → Define Variable. No dialog — instant replacement."""
        pending = getattr(self, "_var_pending_sel", None)
        if pending is None:
            focused = self.focus_get()
            if focused is not None:
                self._snapshot_var_selection_from(focused)
                pending = getattr(self, "_var_pending_sel", None)

        if pending is None:
            messagebox.showinfo("Define Variable",
                "Select some text in a ticket field first, then right-click → Define Variable.")
            return

        kind, src_widget, start, end, sel_text = pending
        self._var_pending_sel = None

        if not sel_text.strip():
            messagebox.showinfo("Define Variable", "Selected text is empty.")
            return

        auto_key = self._next_var_letter()
        marker = "{" + auto_key + "=" + sel_text + "}"

        try:
            if kind == "html":
                self._insert_var_into_adf_json(sel_text, marker)
            else:
                src_widget.delete(start, end)
                src_widget.insert(start, marker)
        except Exception as ex:
            messagebox.showerror("Error", str(ex))
            return
        self._refresh_var_previews()

    def _insert_var_reference(self, key, widget=None):
        """Insert a {KEY} reference at the cursor position of the given or focused widget."""
        if widget is None:
            widget = self.focus_get()
        if widget is None:
            return
        ref = "{" + key + "}"
        try:
            if isinstance(widget, tk.Text):
                widget.insert("insert", ref)
                self._schedule_var_preview(widget)
            elif isinstance(widget, (tk.Entry, ttk.Combobox)):
                pos = widget.index("insert")
                widget.insert(pos, ref)
            else:
                tf = self.get_active_tabform()
                if tf:
                    self._insert_var_ref_into_adf(tf, key)
        except tk.TclError:
            pass

    def _insert_var_reference_from_html_selection(self, key, widget=None, tabform=None):
        """Insert {KEY} at a selected location in HTML preview; fallback to ADF append."""
        if widget is not None:
            self._snapshot_var_selection_from(widget)
        pending = getattr(self, "_var_pending_sel", None)
        ref = "{" + key + "}"
        if pending:
            kind, _src_widget, _start, _end, sel_text = pending
            self._var_pending_sel = None
            if kind == "html" and sel_text and sel_text.strip():
                inserted = self._insert_var_into_adf_json(sel_text, sel_text + ref)
                if inserted:
                    self._refresh_var_previews()
                    return
        if tabform is None:
            tabform = self.get_active_tabform()
        if tabform:
            self._insert_var_ref_into_adf(tabform, key)

    @staticmethod
    def _find_text_nodes(node):
        """Yield all ADF text-node dicts from *node* in document order."""
        if not isinstance(node, dict):
            return
        if node.get("type") == "text":
            yield node
        for child in node.get("content", []):
            yield from VariablesMixin._find_text_nodes(child)

    def _write_adf_and_refresh(self, tabform, doc):
        """Serialize *doc* into the ADF editor, trigger preview update."""
        info = tabform.field_widgets.get("Description ADF")
        if not info:
            return
        adf_widget = info.get("widget")
        adf_widget.delete("1.0", "end")
        adf_widget.insert("1.0", json.dumps(doc, ensure_ascii=False, indent=2))
        try:
            tabform._on_adf_key()
        except Exception:
            pass
        self._refresh_var_previews()

    def _insert_var_ref_into_adf(self, tabform, key, caret_text=None):
        """Insert a {KEY} variable reference into the ADF JSON.

        Always operates on the *parsed* ADF tree so it can never corrupt
        JSON structure.  Three strategies (in priority order):

        1. Match *caret_text* (from the HTML preview) against resolved
           text-node content and append the ref to the best match.
        2. Append to the last text node in the document.
        3. Create a new paragraph with the ref.
        """
        info = tabform.field_widgets.get("Description ADF")
        if not info:
            return
        adf_widget = info.get("widget")
        if not isinstance(adf_widget, tk.Text):
            return

        ref = "{" + key + "}"
        raw = adf_widget.get("1.0", "end-1c")

        try:
            doc = json.loads(raw) if raw.strip() else {"type": "doc", "version": 1, "content": []}
        except Exception:
            doc = {"type": "doc", "version": 1, "content": []}

        text_nodes = list(self._find_text_nodes(doc))
        target = None

        # Strategy 1 — match caret_text against resolved node content
        if caret_text and text_nodes:
            vars_dict = self._collect_defined_var_keys()
            ct = caret_text.replace("\xa0", " ").strip()
            if ct:
                best_score = -1
                for tn in text_nodes:
                    raw_t = (tn.get("text") or "").replace("\xa0", " ").strip()
                    resolved = (self._apply_variables_to_text(tn.get("text", ""), vars_dict)
                                if vars_dict else tn.get("text", ""))
                    rc = resolved.replace("\xa0", " ").strip()
                    if not rc:
                        continue
                    if rc == ct:
                        score = 100
                    elif ct in rc:
                        score = 25 + 50 * len(ct) / max(len(rc), 1)
                    elif rc in ct:
                        score = 25 + 50 * len(rc) / max(len(ct), 1)
                    else:
                        continue
                    if score > best_score:
                        best_score = score
                        target = tn

        # Strategy 2 — append to the last text node
        if target is None and text_nodes:
            target = text_nodes[-1]

        # Strategy 3 — create a new paragraph
        if target is None:
            content = doc.setdefault("content", [])
            content.append({"type": "paragraph", "content": [{"type": "text", "text": ref}]})
        else:
            target["text"] = (target.get("text") or "") + ref

        self._write_adf_and_refresh(tabform, doc)

    def _insert_var_into_adf_json(self, original_text, replacement):
        """Replace *original_text* with *replacement* inside an ADF text node.

        Works on the parsed ADF tree so it can never corrupt JSON structure.
        Used when a variable is defined from the HTML preview pane.
        """
        tf = self.get_active_tabform()
        if not tf:
            return False
        info = tf.field_widgets.get("Description ADF")
        if not info:
            return False
        adf_widget = info.get("widget")
        if not isinstance(adf_widget, tk.Text):
            return False
        raw = adf_widget.get("1.0", "end-1c")

        try:
            doc = json.loads(raw) if raw.strip() else None
        except Exception:
            doc = None
        if not doc:
            return False

        for tn in self._find_text_nodes(doc):
            txt = tn.get("text", "")
            if original_text in txt:
                tn["text"] = txt.replace(original_text, replacement, 1)
                self._write_adf_and_refresh(tf, doc)
                return True
        return False

    def _setup_var_preview(self, text_widget):
        """Configure tags and bindings for live variable preview on a tk.Text widget."""
        text_widget._var_preview_enabled = True
        text_widget.tag_configure(self._VAR_PREVIEW_TAG,
                                  background="#1a3a2a", foreground="#4ec9b0",
                                  font=("Segoe UI", 10, "italic"))
        text_widget.tag_configure(self._VAR_EDITING_TAG,
                                  background="#2d2d00", foreground="#dcdcaa",
                                  font=("Consolas", 10, "bold"))
        text_widget.tag_bind(self._VAR_PREVIEW_TAG, "<Button-1>",
                             lambda e: self._on_var_preview_click(text_widget, e))
        text_widget.bind("<KeyRelease>", lambda e, w=text_widget: self._on_var_key(w, e), add=True)
        text_widget.bind("<FocusOut>", lambda e, w=text_widget: self._on_var_focus_out(w), add=True)

    def _schedule_var_preview(self, text_widget):
        """Debounced refresh of variable previews for a text widget."""
        aid = getattr(text_widget, "_var_preview_after", None)
        if aid:
            try:
                text_widget.after_cancel(aid)
            except Exception:
                pass
        text_widget._var_preview_after = text_widget.after(150, lambda: self._refresh_var_preview_widget(text_widget))

    def _on_var_key(self, text_widget, event):
        """After each key release, check if we need to update previews."""
        if event.keysym in ("Return", "braceright", "BackSpace", "Delete", "space") or len(event.char) == 1:
            self._schedule_var_preview(text_widget)

    def _on_var_focus_out(self, text_widget):
        """When focus leaves, clear editing tags and collapse references back to previews."""
        text_widget.tag_remove(self._VAR_EDITING_TAG, "1.0", "end")
        self._schedule_var_preview(text_widget)

    def _on_var_preview_click(self, text_widget, event):
        """Click on a previewed variable value → revert to showing {KEY} for editing."""
        try:
            idx = text_widget.index(f"@{event.x},{event.y}")
            tag_ranges = text_widget.tag_ranges(self._VAR_PREVIEW_TAG)
            for i in range(0, len(tag_ranges), 2):
                start, end = str(tag_ranges[i]), str(tag_ranges[i + 1])
                if text_widget.compare(idx, ">=", start) and text_widget.compare(idx, "<=", end):
                    raw_key = text_widget.get(start, end)
                    stored = getattr(text_widget, "_var_preview_map", {})
                    ref_text = stored.get(raw_key, raw_key)
                    text_widget.delete(start, end)
                    text_widget.insert(start, ref_text)
                    text_widget.tag_add(self._VAR_EDITING_TAG, start, f"{start}+{len(ref_text)}c")
                    text_widget.mark_set("insert", f"{start}+{len(ref_text)}c")
                    text_widget.focus_set()
                    return "break"
        except (tk.TclError, StopIteration):
            pass
        return "break"

    def _refresh_var_previews(self):
        """Refresh variable previews on widgets that opted in, then update the HTML preview."""
        tf = self.get_active_tabform()
        if not tf:
            return
        for fw in tf.field_widgets.values():
            w = fw.get("widget")
            if isinstance(w, tk.Text) and getattr(w, "_var_preview_enabled", False):
                self._refresh_var_preview_widget(w)
        try:
            tf._update_preview_from_adf()
        except Exception:
            pass
        try:
            tf._push_vars_to_editor()
        except Exception:
            pass

    def _revert_all_var_previews(self, tabform):
        """Revert all live previews in a tabform back to raw {KEY} text. Call before reading data."""
        for fw in tabform.field_widgets.values():
            w = fw.get("widget")
            if not isinstance(w, tk.Text) or not getattr(w, "_var_preview_enabled", False):
                continue
            old_map = getattr(w, "_var_preview_map", {})
            if not old_map:
                continue
            ranges = w.tag_ranges(self._VAR_PREVIEW_TAG)
            pairs = []
            for i in range(0, len(ranges), 2):
                pairs.append((str(ranges[i]), str(ranges[i + 1])))
            for start, end in reversed(pairs):
                display = w.get(start, end)
                ref = old_map.get(display, display)
                w.delete(start, end)
                w.insert(start, ref)
            w.tag_remove(self._VAR_PREVIEW_TAG, "1.0", "end")
            w.tag_remove(self._VAR_EDITING_TAG, "1.0", "end")
            w._var_preview_map = {}

    def _refresh_var_preview_widget(self, text_widget):
        """Scan a tk.Text for {KEY} references and replace them with previewed values.
        Also reverts any previously previewed values back to {KEY} first so raw text stays clean."""
        defined = self._collect_defined_var_keys()

        # First: revert any existing previews back to their raw {KEY} form
        old_map = getattr(text_widget, "_var_preview_map", {})
        if old_map:
            ranges = text_widget.tag_ranges(self._VAR_PREVIEW_TAG)
            pairs = []
            for i in range(0, len(ranges), 2):
                pairs.append((str(ranges[i]), str(ranges[i + 1])))
            for start, end in reversed(pairs):
                display = text_widget.get(start, end)
                ref = old_map.get(display, display)
                text_widget.delete(start, end)
                text_widget.insert(start, ref)
            text_widget.tag_remove(self._VAR_PREVIEW_TAG, "1.0", "end")
        text_widget._var_preview_map = {}

        if not defined:
            text_widget.tag_remove(self._VAR_EDITING_TAG, "1.0", "end")
            return

        try:
            cursor = text_widget.index("insert")
        except tk.TclError:
            cursor = "1.0"

        editing_ranges = []
        er = text_widget.tag_ranges(self._VAR_EDITING_TAG)
        for i in range(0, len(er), 2):
            editing_ranges.append((str(er[i]), str(er[i + 1])))

        def _in_editing(pos):
            for s, e in editing_ranges:
                if text_widget.compare(pos, ">=", s) and text_widget.compare(pos, "<=", e):
                    return True
            return False

        content = text_widget.get("1.0", "end-1c")
        preview_map = {}

        # Collect matches, then process in reverse so char offsets don't shift
        matches = list(re.finditer(r'\{([A-Z])\}', content))
        for m in reversed(matches):
            key = m.group(1)
            if key not in defined:
                continue
            val = defined[key]
            ref = m.group(0)
            start_idx = f"1.0+{m.start()}c"
            end_idx = f"1.0+{m.end()}c"

            if text_widget.compare(cursor, ">=", start_idx) and text_widget.compare(cursor, "<=", end_idx):
                continue
            if _in_editing(start_idx):
                continue

            actual = text_widget.get(start_idx, end_idx)
            if actual != ref:
                continue

            display = val
            preview_map[display] = ref
            text_widget.delete(start_idx, end_idx)
            text_widget.insert(start_idx, display)
            new_end = f"{start_idx}+{len(display)}c"
            text_widget.tag_add(self._VAR_PREVIEW_TAG, start_idx, new_end)

        text_widget._var_preview_map = preview_map

    def _show_summary_mismatch_excluded(self, excluded_keys):
        """Show warning that tickets were excluded due to summary mismatch."""
        if not excluded_keys:
            return
        keys_str = ", ".join(excluded_keys)
        messagebox.showwarning(
            "Summary mismatch - excluded",
            f"The following ticket(s) were excluded from the upload because their summary "
            f"doesn't match Jira:\n\n{keys_str}\n\n"
            "To create a NEW ticket instead, remove Issue key and Issue id from those tickets first."
        )

    def _apply_keyable_fields_from_jira(self, ticket, issue_json):
        """Update ticket with keyable fields from Jira (Issue key, Issue id, Project key, Issue Type, Status, etc.)."""
        if not issue_json:
            return
        fields = issue_json.get("fields") or {}
        status_obj = fields.get("status") or {}
        ticket["Issue key"] = issue_json.get("key", "")
        ticket["Issue id"] = issue_json.get("id", "")
        ticket["Project key"] = (fields.get("project") or {}).get("key", "")
        ticket["Project name"] = (fields.get("project") or {}).get("name", "")
        ticket["Issue Type"] = (fields.get("issuetype") or {}).get("name", "")
        ticket["Status"] = status_obj.get("name", "")
        ticket["Status Category"] = (status_obj.get("statusCategory") or {}).get("name", "")
        ticket["Reporter"] = (fields.get("reporter") or {}).get("displayName", "") or (fields.get("reporter") or {}).get("emailAddress", "") or ""
        ticket["Created"] = fields.get("created", "")
        ticket["Updated"] = fields.get("updated", "")

    def _collect_variables(self, ticket):
        """Collect variables from {KEY=value} definitions across all string fields."""
        vars_dict = {}
        for val in ticket.values():
            if isinstance(val, str):
                for m in self._VAR_DEF_RE.finditer(val):
                    vars_dict[m.group(1)] = m.group(2).strip()
        return vars_dict

    def _apply_variables_to_text(self, text, vars_dict):
        """Strip variable definitions and replace {KEY} references. Returns processed string."""
        if not text or not isinstance(text, str):
            return text
        # Strip {KEY=value} → keep only the value
        out = self._VAR_DEF_RE.sub(lambda m: m.group(2).strip(), text)
        # Replace {KEY} references
        def _repl(m):
            return str(vars_dict[m.group(1)]) if m.group(1) in vars_dict else m.group(0)
        out = self._VAR_REF_RE.sub(_repl, out)
        return out

    def _apply_variables_to_adf(self, adf, vars_dict):
        """Recursively apply variable substitution to ADF (replace {var} in text nodes)."""
        if not vars_dict:
            return adf
        if isinstance(adf, dict):
            if adf.get("type") == "text" and "text" in adf:
                adf = dict(adf)
                adf["text"] = self._apply_variables_to_text(adf["text"], vars_dict)
            result = {}
            for k, v in adf.items():
                result[k] = self._apply_variables_to_adf(v, vars_dict)
            return result
        if isinstance(adf, list):
            return [self._apply_variables_to_adf(x, vars_dict) for x in adf]
        return adf

    def _apply_variables_to_ticket(self, ticket):
        """Apply variable substitution to ticket content. Modifies ticket in place. Definitions removed, {var} replaced."""
        vars_dict = self._collect_variables(ticket)
        if not vars_dict:
            return ticket
        text_headers = ("Summary", "Description", "Variables", "Labels", "Components", "Environment", "Comment", "Parent summary")
        for h in text_headers:
            if h in ticket and isinstance(ticket[h], str):
                ticket[h] = self._apply_variables_to_text(ticket[h], vars_dict)
        if ticket.get("Description ADF") and isinstance(ticket["Description ADF"], (dict, list)):
            ticket["Description ADF"] = self._apply_variables_to_adf(copy.deepcopy(ticket["Description ADF"]), vars_dict)
        return ticket
