"""
BulkImportMixin — Bulk import tickets from pasted text using a template.
"""
import json
import copy
import re
import uuid
import tkinter as tk
from tkinter import ttk, messagebox

from ..config import HEADERS
from ..storage import save_storage
from ..utils import debug_log, _bind_mousewheel


class BulkImportMixin:
    """Mixin providing bulk_import_dialog."""

    def bulk_import_dialog(self, on_close=None, prefill_text=None, prefill_template=None):
        """Open dialog to bulk import tickets from pasted summaries using a template."""
        if not self.templates:
            messagebox.showinfo("Info", "Create at least one template first.")
            if on_close:
                self.after(50, on_close)
            return
        opts = self.meta.setdefault("bulk_import", {})
        dlg = tk.Toplevel(self)
        self._register_toplevel(dlg)
        dlg.title("Bulk Import Tickets")
        dlg.minsize(520, 520)
        dlg.geometry("640x620")
        dlg.resizable(True, True)
        # Template
        tpl_frame = ttk.Frame(dlg)
        tpl_frame.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Label(tpl_frame, text="Template:").pack(side="left", padx=(0, 8))
        if prefill_template and prefill_template in self.templates:
            default_tpl = prefill_template
        else:
            default_tpl = opts.get("template") or (sorted(self.templates.keys())[0] if self.templates else "")
            if default_tpl not in self.templates:
                default_tpl = sorted(self.templates.keys())[0] if self.templates else ""
        sorted_keys = sorted(self.templates.keys())
        tpl_var = tk.StringVar(value=default_tpl)
        tpl_combo = ttk.Combobox(tpl_frame, textvariable=tpl_var, values=sorted_keys, width=36, state="readonly")
        if default_tpl in sorted_keys:
            tpl_combo.current(sorted_keys.index(default_tpl))
        tpl_combo.pack(side="left", fill="x", expand=True)
        # Options frame
        opt_frame = ttk.LabelFrame(dlg, text="Options", padding=8)
        opt_frame.pack(fill="x", padx=8, pady=6)
        # Line delimiter
        row1 = ttk.Frame(opt_frame)
        row1.pack(fill="x", pady=2)
        ttk.Label(row1, text="Line delimiter:").pack(side="left", padx=(0, 8))
        delim_var = tk.StringVar(value=opts.get("delimiter", "newline"))
        for val, lbl in [("newline", "Newline"), ("comma", "Comma"), ("semicolon", "Semicolon"), ("tab", "Tab"), ("pipe", "Pipe (|)")]:
            ttk.Radiobutton(row1, text=lbl, variable=delim_var, value=val).pack(side="left", padx=(0, 12))
        # Exclude
        row2 = ttk.Frame(opt_frame)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="Exclude lines:").pack(side="left", padx=(0, 8))
        excl_empty_var = tk.BooleanVar(value=opts.get("exclude_empty", True))
        ttk.Checkbutton(row2, text="Empty", variable=excl_empty_var).pack(side="left", padx=(0, 12))
        excl_starts_var = tk.StringVar(value=opts.get("exclude_starts_with", ""))
        ttk.Label(row2, text="Starting with:").pack(side="left", padx=(12, 4))
        ttk.Entry(row2, textvariable=excl_starts_var, width=12).pack(side="left", padx=(0, 8))
        ttk.Label(row2, text="(e.g. # or //)").pack(side="left")
        # Summary mode
        row3 = ttk.Frame(opt_frame)
        row3.pack(fill="x", pady=2)
        ttk.Label(row3, text="Parse mode:").pack(side="left", padx=(0, 8))
        mode_var = tk.StringVar(value=opts.get("summary_mode", "structured"))
        for val, lbl in [
            ("replace", "Replace template summary"),
            ("prepend", "Prepend to template"),
            ("append", "Append to template"),
            ("structured", "Structured: blank line = new ticket; 1st line = title; rest = details for !")
        ]:
            ttk.Radiobutton(row3, text=lbl, variable=mode_var, value=val).pack(side="left", padx=(0, 12))
        row4 = ttk.Frame(opt_frame)
        row4.pack(fill="x", pady=2)
        sep_var = tk.StringVar(value=opts.get("separator", " "))
        ttk.Label(row4, text="Separator (for prepend/append):").pack(side="left", padx=(0, 8))
        ttk.Entry(row4, textvariable=sep_var, width=8).pack(side="left", padx=(0, 8))
        ttk.Label(row4, text="(e.g. space, ' - ', ' | ')").pack(side="left")
        row4b = ttk.Frame(opt_frame)
        row4b.pack(fill="x", pady=2)
        ttk.Label(row4b, text="Structured: Blank lines separate tickets. 1st line = title; following lines = details for ! in template (or appended if no !).").pack(side="left", padx=(0, 8))
        # Content format for detail lines
        row5 = ttk.Frame(opt_frame)
        row5.pack(fill="x", pady=2)
        ttk.Label(row5, text="Content format:").pack(side="left", padx=(0, 8))
        cfmt_var = tk.StringVar(value=opts.get("content_format", "smart_list"))
        for val, lbl in [("smart_list", "Smart list"), ("paragraphs", "Paragraphs"), ("bullet_list", "Bullet list"), ("ordered_list", "Ordered list")]:
            ttk.Radiobutton(row5, text=lbl, variable=cfmt_var, value=val).pack(side="left", padx=(0, 12))
        ttk.Label(row5, text="(how detail lines are inserted for !)").pack(side="left")
        # Paste box
        ttk.Label(dlg, text="Paste below. Structured: blank line = new ticket. First line = title; lines after = details for ! placeholders (or end of ticket if none).").pack(anchor="w", padx=8, pady=(8, 4))
        paste_frame = ttk.Frame(dlg)
        paste_frame.pack(fill="both", expand=True, padx=8, pady=4)
        paste_txt = tk.Text(paste_frame, height=14, wrap="word", font=("Segoe UI", 10), bg="#1e1e1e", fg="#dcdcdc", insertbackground="#dcdcdc")
        paste_sb = ttk.Scrollbar(paste_frame, orient="vertical", command=paste_txt.yview)
        paste_txt.pack(side="left", fill="both", expand=True)
        paste_sb.pack(side="right", fill="y")
        paste_txt.configure(yscrollcommand=paste_sb.set)
        _bind_mousewheel(paste_txt, "vertical")
        # Buttons
        def do_import():
            template_name = tpl_var.get().strip()
            if not template_name or template_name not in self.templates:
                messagebox.showerror("Error", "Select a valid template.")
                return
            text = paste_txt.get("1.0", "end").strip()
            if not text:
                messagebox.showinfo("Info", "Paste some text first.")
                return
            # Normalize line endings
            text = text.replace("\r\n", "\n").replace("\r", "\n")
            # Auto-detect block format: blank lines separate tickets → use structured parsing
            use_structured = mode_var.get() == "structured" or "\n\n" in text
            # Save options
            opts["template"] = template_name
            opts["delimiter"] = delim_var.get()
            opts["exclude_empty"] = excl_empty_var.get()
            opts["exclude_starts_with"] = excl_starts_var.get().strip()
            opts["summary_mode"] = mode_var.get()
            opts["separator"] = sep_var.get() or " "
            opts["content_format"] = cfmt_var.get()
            self.meta["bulk_import"] = opts
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass
            mode = mode_var.get()
            content_format = cfmt_var.get()
            excl_start = excl_starts_var.get().strip()
            items = []
            if use_structured:
                # Blank lines separate tickets. 1st line of block = title; rest = details for ! in template
                blocks = re.split(r"\n\s*\n", text)
                items = []
                for bi, block in enumerate(blocks):
                    lines = [raw.strip() for raw in block.split("\n")]
                    if excl_start:
                        lines = [s for s in lines if not (s and s.startswith(excl_start))]
                    if excl_empty_var.get():
                        lines = [s for s in lines if s]
                    if not lines:
                        continue
                    # Skip first block if it looks like a header (single line ending with : or common header phrases)
                    if bi == 0 and len(lines) == 1:
                        first = lines[0]
                        if first.endswith(":") or any(p in first.lower() for p in ("here are", "below are", "the following", "list of")):
                            continue
                    title = lines[0]
                    details = "\n".join(lines[1:]) if len(lines) > 1 else ""
                    items.append((title, details))
            else:
                delim_map = {"newline": "\n", "comma": ",", "semicolon": ";", "tab": "\t", "pipe": "|"}
                sep = delim_map.get(delim_var.get(), "\n")
                raw_items = [s.strip() for s in text.split(sep)]
                for s in raw_items:
                    if excl_empty_var.get() and not s:
                        continue
                    if excl_start and s.startswith(excl_start):
                        continue
                    items.append((s, ""))
            if not items:
                messagebox.showinfo("Info", "No items after applying exclude rules.")
                return
            # Create tickets
            base_summary = (self.templates[template_name].get("Summary") or "").strip()
            sep_str = (sep_var.get() or " ").strip() or " "
            created = 0
            for i, (summary, field_content_lines) in enumerate(items):
                ticket = copy.deepcopy(self.templates[template_name])
                # Apply summary mode (replace/prepend/append) for all modes including structured
                if mode == "replace" or mode == "structured":
                    ticket["Summary"] = summary.strip()
                elif mode == "prepend":
                    ticket["Summary"] = (summary.strip() + sep_str + base_summary) if base_summary else summary.strip()
                else:  # append
                    ticket["Summary"] = (base_summary + sep_str + summary.strip()) if base_summary else summary.strip()
                if use_structured:
                    if field_content_lines:
                        content_list = [s.strip() for s in field_content_lines.split("\n") if s.strip()]
                    else:
                        content_list = []
                    if content_list:
                        found_any = False
                        insert_after = [None, None]  # (content_list, index) for inserting leftovers after last !
                        # Process ADF first so table-cell ! get replacements before string fields consume them
                        adf = ticket.get("Description ADF")
                        if isinstance(adf, str):
                            try:
                                adf = json.loads(adf)
                            except Exception:
                                adf = self._text_to_adf(adf) if adf.strip() else None
                        # If ADF is empty/missing, try to recover it from the
                        # fetched issue list (which stores the original Jira ADF).
                        if not adf or (isinstance(adf, dict) and not adf.get("content")):
                            recovered = self._recover_template_adf(template_name)
                            if recovered:
                                adf = copy.deepcopy(recovered)
                                ticket["Description ADF"] = adf
                        # If ADF still has no content with !, fall back to the plain
                        # Description text field (which the user can edit directly).
                        # This handles templates where the ADF was never populated but
                        # the plain Description contains ! placeholders.
                        plain_desc = ticket.get("Description", "")
                        if isinstance(plain_desc, str) and "!" in plain_desc:
                            if not isinstance(adf, dict) or not self._adf_contains_exclamation(adf):
                                fallback_adf = self._text_to_adf(plain_desc)
                                if self._adf_contains_exclamation(fallback_adf):
                                    adf = fallback_adf
                                    ticket["Description ADF"] = adf

                        # Count total ! across ADF and string fields to decide
                        # whether a single ! should receive ALL lines joined.
                        # Skip "Description" — it mirrors ADF text and would double-count.
                        _SKIP_EXCL_FIELDS = {"Description", "Description ADF"}
                        total_excl = 0
                        if isinstance(adf, dict):
                            total_excl += self._count_exclamations_in_adf(adf)
                        for h in HEADERS:
                            if h in _SKIP_EXCL_FIELDS:
                                continue
                            val = ticket.get(h, "")
                            if isinstance(val, str):
                                total_excl += val.count("!")
                        _list_type_map = {"bullet_list": "bulletList", "ordered_list": "orderedList"}
                        use_list_fmt = content_format in _list_type_map or content_format == "smart_list"
                        if total_excl == 1 and len(content_list) > 1:
                            if use_list_fmt:
                                replacements = [""]
                            else:
                                replacements = ["\n".join(content_list)]
                        else:
                            replacements = content_list[:]
                        if isinstance(adf, dict) and (replacements or use_list_fmt):
                            adf = copy.deepcopy(adf)
                            if self._adf_contains_exclamation(adf):
                                found_any = True
                                self._replace_in_adf(adf, replacements, insert_after=insert_after)
                                if use_list_fmt and total_excl == 1 and content_list:
                                    if content_format == "smart_list":
                                        new_nodes = self._build_adf_smart_list(content_list)
                                    else:
                                        ln = self._build_adf_list(content_list, _list_type_map[content_format])
                                        new_nodes = [ln] if ln else []
                                    if new_nodes:
                                        p_list, p_idx = insert_after[0], insert_after[1]
                                        if p_list is not None and p_idx is not None:
                                            para = p_list[p_idx] if p_idx < len(p_list) else None
                                            if para and isinstance(para, dict) and para.get("type") == "paragraph":
                                                texts = [c.get("text", "") for c in (para.get("content") or []) if c.get("type") == "text"]
                                                if all(not t.strip() for t in texts):
                                                    p_list.pop(p_idx)
                                                    for ni, n in enumerate(new_nodes):
                                                        p_list.insert(p_idx + ni, n)
                                                else:
                                                    for ni, n in enumerate(new_nodes):
                                                        p_list.insert(p_idx + 1 + ni, n)
                                            else:
                                                for ni, n in enumerate(new_nodes):
                                                    p_list.insert(p_idx + 1 + ni, n)
                                        elif isinstance(adf.get("content"), list):
                                            adf["content"].extend(new_nodes)
                                ticket["Description ADF"] = adf
                        # Replace ! in string fields (skip Description — it mirrors ADF)
                        for h in HEADERS:
                            if h in _SKIP_EXCL_FIELDS:
                                continue
                            val = ticket.get(h, "")
                            if isinstance(val, str) and "!" in val:
                                found_any = True
                                while "!" in val and replacements:
                                    val = val.replace("!", replacements.pop(0), 1)
                                ticket[h] = val
                        # Insert leftovers after last ! (or append to end if no ! found)
                        if replacements or not found_any:
                            leftover_items = content_list if not found_any else replacements
                            bottom_text = "\n".join(leftover_items)
                            if bottom_text:
                                existing_plain = (ticket.get("Description") or "").strip()
                                combined_plain = (existing_plain + "\n\n" + bottom_text).strip() if existing_plain else bottom_text
                                ticket["Description"] = combined_plain
                                adf = ticket.get("Description ADF")
                                if isinstance(adf, str):
                                    try:
                                        adf = json.loads(adf)
                                    except Exception:
                                        adf = None
                                if content_format == "smart_list":
                                    new_nodes = self._build_adf_smart_list(leftover_items)
                                elif use_list_fmt:
                                    new_block = self._build_adf_list(leftover_items, _list_type_map[content_format])
                                    new_nodes = [new_block] if new_block else []
                                else:
                                    new_paras = self._text_to_adf(bottom_text)
                                    new_nodes = list(new_paras.get("content", [])) if new_paras else []
                                if new_nodes:
                                    if isinstance(adf, dict):
                                        parent_list, idx = insert_after[0], insert_after[1]
                                        if parent_list is not None and idx is not None:
                                            for p in reversed(new_nodes):
                                                parent_list.insert(idx + 1, copy.deepcopy(p))
                                        else:
                                            adf = copy.deepcopy(adf)
                                            adf.setdefault("content", [])
                                            adf["content"] = list(adf["content"]) + [copy.deepcopy(n) for n in new_nodes]
                                        ticket["Description ADF"] = adf
                                    else:
                                        if use_list_fmt and new_nodes:
                                            ticket["Description ADF"] = {"type": "doc", "version": 1, "content": new_nodes}
                                        else:
                                            ticket["Description ADF"] = self._text_to_adf(bottom_text)
                ticket["Issue id"] = str(uuid.uuid4())
                ticket["Issue key"] = "LOCAL-" + ticket["Issue id"][:8]
                self.bundle.append(ticket)
                self.new_tab(initial_data=ticket, select_tab=(i == 0))
                created += 1
            dlg.destroy()
            self.update_bundle_listbox()
            self.show_tabs_view()
            if not getattr(self, "_tutorial_running", False):
                messagebox.showinfo("Bulk Import", f"Created {created} tickets from template '{template_name}' and added to bundle.")
            if on_close:
                self.after(50, on_close)

        def _cancel():
            dlg.destroy()
            if on_close:
                self.after(50, on_close)

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(fill="x", padx=8, pady=8)
        ttk.Button(btn_frame, text="Import", command=do_import).pack(side="right", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=_cancel).pack(side="right", padx=6)
        dlg.protocol("WM_DELETE_WINDOW", _cancel)
        if prefill_text:
            paste_txt.insert("1.0", prefill_text)
