"""
BundleShareMixin — export and import Avalanche bundles (.avl files).

A "bundle" is a portable JSON snapshot of one or more tickets that can be
e-mailed, dropped in a shared folder, or sent via chat to a colleague.
The recipient opens Avalanche and imports the file — tickets land directly
in their list, ready to edit or upload.

Bundle format (avalanche_bundle v1.0):
{
    "avalanche_bundle": "1.0",
    "exported_by":  "alice@example.com",
    "exported_at":  "2026-03-01T14:32:00",
    "message":      "Hey, here are the tickets for the sprint",
    "tickets": [ { ...ticket dict... }, ... ]
}
"""
import json
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from utils import debug_log, _dedup_list_items
from storage import save_storage

_BUNDLE_VERSION = "1.0"
_FILE_EXT = ".avl"
_FILE_TYPE = [("Avalanche Bundle", "*.avl"), ("JSON", "*.json"), ("All files", "*.*")]


class BundleShareMixin:
    """Adds _share_bundle_dialog and _import_bundle_dialog to AvalancheApp."""

    # ──────────────────────────────────────────────────────────────────────────
    # SHARE / EXPORT
    # ──────────────────────────────────────────────────────────────────────────

    def _share_bundle_dialog(self):
        """Export the current bundle to a .avl file to share with a colleague.

        Only tickets already in the bundle panel (self.bundle) are shown.
        Compose the bundle first using the 'Add ▶' button, then export here.
        """
        bundle_items = getattr(self, "bundle", [])
        if not bundle_items:
            messagebox.showinfo(
                "Export Bundle",
                "The bundle is empty.\n\n"
                "Add tickets to the bundle first using the 'Add ▶' button, "
                "then export.",
            )
            return

        bundle_name = getattr(self, "bundle_name", None) or "bundle"

        if self._focus_existing_app_dialog("export_bundle"):
            return
        win = tk.Toplevel(self)
        self._track_app_dialog("export_bundle", win)
        self._register_toplevel(win)
        win.title("Export Bundle")
        win.geometry("640x540")
        win.minsize(520, 400)
        win.resizable(True, True)

        # ── header ────────────────────────────────────────────────────────────
        hdr = ttk.Frame(win)
        hdr.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(hdr, text="Export Bundle",
                  font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(hdr,
                  text="The tickets below are your current bundle. "
                       "Untick any you want to exclude, add an optional message, "
                       "then save the .avl file to share with a colleague.",
                  wraplength=600, font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))
        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=8, pady=6)

        # ── ticket list with checkboxes ───────────────────────────────────────
        ttk.Label(win, text=f"Bundle contents  ({len(bundle_items)} ticket(s)):").pack(
            anchor="w", padx=12)

        list_frame = ttk.Frame(win)
        list_frame.pack(fill="both", expand=True, padx=12, pady=(2, 4))

        columns = ("sel", "key", "summary", "status")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings",
                            selectmode="browse", height=12)
        tree.heading("sel",     text="✓",       anchor="center")
        tree.heading("key",     text="Key")
        tree.heading("summary", text="Summary")
        tree.heading("status",  text="Status")
        tree.column("sel",     width=30,  stretch=False, anchor="center")
        tree.column("key",     width=100, stretch=False)
        tree.column("summary", width=280, stretch=True)
        tree.column("status",  width=100, stretch=False)

        vs = ttk.Scrollbar(list_frame, orient="vertical",   command=tree.yview)
        hs = ttk.Scrollbar(list_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        # Track selection state per row
        _checked = {}  # iid -> BooleanVar

        for i, item in enumerate(bundle_items):
            key    = item.get("Issue key") or item.get("Issue id") or f"#{i+1}"
            summ   = (item.get("Summary") or "")[:80]
            status = item.get("Status") or ""
            var    = tk.BooleanVar(value=True)
            _checked[str(i)] = var
            tree.insert("", tk.END, iid=str(i),
                        values=("☑", key, summ, status))

        def _toggle(event):
            iid = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            if not iid:
                return
            var = _checked.get(iid)
            if var is None:
                return
            var.set(not var.get())
            tree.set(iid, "sel", "☑" if var.get() else "☐")
            _update_count()

        tree.bind("<Button-1>", _toggle)

        # select/deselect all bar
        sel_bar = ttk.Frame(win)
        sel_bar.pack(fill="x", padx=12, pady=(0, 4))

        def _set_all(val):
            for iid, var in _checked.items():
                var.set(val)
                tree.set(iid, "sel", "☑" if val else "☐")
            _update_count()

        ttk.Button(sel_bar, text="Select All",
                   command=lambda: _set_all(True)).pack(side="left", padx=(0, 4))
        ttk.Button(sel_bar, text="Deselect All",
                   command=lambda: _set_all(False)).pack(side="left")
        count_lbl = ttk.Label(sel_bar, text="")
        count_lbl.pack(side="right")

        def _update_count():
            n = sum(1 for v in _checked.values() if v.get())
            count_lbl.config(text=f"{n} ticket(s) selected")

        _update_count()

        # ── optional message ──────────────────────────────────────────────────
        ttk.Label(win, text="Optional message for recipient:").pack(
            anchor="w", padx=12, pady=(4, 0))
        msg_var = tk.StringVar()
        ttk.Entry(win, textvariable=msg_var, font=("Segoe UI", 10)).pack(
            fill="x", padx=12, pady=(2, 8))

        # ── saved path readout ────────────────────────────────────────────────
        saved_frame = ttk.Frame(win)
        saved_frame.pack(fill="x", padx=12, pady=(0, 4))
        saved_path_var = tk.StringVar()
        saved_entry = ttk.Entry(saved_frame, textvariable=saved_path_var,
                                state="readonly", font=("Segoe UI", 9))
        saved_entry.pack(side="left", fill="x", expand=True)

        def _copy_path():
            p = saved_path_var.get()
            if p:
                win.clipboard_clear()
                win.clipboard_append(p)
                messagebox.showinfo("Copied", "File path copied to clipboard.")

        ttk.Button(saved_frame, text="Copy Path",
                   command=_copy_path).pack(side="left", padx=(4, 0))

        # ── action buttons ────────────────────────────────────────────────────
        btn_bar = ttk.Frame(win)
        btn_bar.pack(fill="x", padx=12, pady=(4, 12))

        def do_save():
            selected = [bundle_items[int(iid)]
                        for iid, var in _checked.items() if var.get()
                        and int(iid) < len(bundle_items)]
            if not selected:
                messagebox.showwarning("No Tickets", "Select at least one ticket to export.")
                return

            dest = filedialog.asksaveasfilename(
                title="Save Bundle As",
                defaultextension=_FILE_EXT,
                filetypes=_FILE_TYPE,
                initialfile=bundle_name,
            )
            if not dest:
                return

            jira_cfg = self.meta.get("jira", {})
            exported_by = (jira_cfg.get("email") or "").strip() or "unknown"
            bundle = {
                "avalanche_bundle": _BUNDLE_VERSION,
                "exported_by":  exported_by,
                "exported_at":  datetime.datetime.now().isoformat(timespec="seconds"),
                "message":      msg_var.get().strip(),
                "tickets":      selected,
            }
            try:
                with open(dest, "w", encoding="utf-8") as f:
                    json.dump(bundle, f, indent=2, ensure_ascii=False, default=str)
                saved_path_var.set(dest)
                messagebox.showinfo(
                    "Bundle Saved",
                    f"Bundle with {len(selected)} ticket(s) saved to:\n{dest}\n\n"
                    "Share this file with your colleague — they can import it via "
                    "Import Bundle in Avalanche.",
                )
            except Exception as e:
                messagebox.showerror("Save Failed", str(e))
                debug_log(f"Bundle export error: {e}")

        ttk.Button(btn_bar, text="💾  Save Bundle File…",
                   command=do_save).pack(side="left", padx=(0, 6))
        ttk.Button(btn_bar, text="Close",
                   command=win.destroy).pack(side="right")

    # ──────────────────────────────────────────────────────────────────────────
    # IMPORT
    # ──────────────────────────────────────────────────────────────────────────

    def _import_bundle_dialog(self):
        """Open the Import Bundle dialog — browse for a .avl file and import tickets."""
        path = filedialog.askopenfilename(
            title="Open Bundle File",
            filetypes=_FILE_TYPE,
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                bundle = json.load(f)
        except Exception as e:
            messagebox.showerror("Import Failed", f"Could not read bundle file:\n{e}")
            return

        if not isinstance(bundle, dict) or bundle.get("avalanche_bundle") != _BUNDLE_VERSION:
            messagebox.showerror(
                "Invalid File",
                "This file does not appear to be a valid Avalanche bundle.",
            )
            return

        tickets = bundle.get("tickets", [])
        if not isinstance(tickets, list) or not tickets:
            messagebox.showinfo("Empty Bundle", "The bundle contains no tickets.")
            return

        # ── preview dialog ────────────────────────────────────────────────────
        if self._focus_existing_app_dialog("import_bundle"):
            return
        win = tk.Toplevel(self)
        self._track_app_dialog("import_bundle", win)
        self._register_toplevel(win)
        win.title("Import Bundle")
        win.geometry("680x560")
        win.minsize(520, 400)
        win.resizable(True, True)

        ttk.Label(win, text="Import Bundle",
                  font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=12, pady=(12, 2))
        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=8, pady=4)

        # meta info
        meta_frame = ttk.Frame(win)
        meta_frame.pack(fill="x", padx=12, pady=(0, 6))

        def _meta_row(label, value):
            r = ttk.Frame(meta_frame)
            r.pack(fill="x", pady=1)
            ttk.Label(r, text=label, width=14, anchor="e",
                      font=("Segoe UI", 9, "bold")).pack(side="left")
            ttk.Label(r, text=value or "—",
                      font=("Segoe UI", 9)).pack(side="left", padx=(4, 0))

        _meta_row("From:",    bundle.get("exported_by", ""))
        _meta_row("Exported:", bundle.get("exported_at", ""))
        _meta_row("Tickets:", str(len(tickets)))
        msg = bundle.get("message", "").strip()
        if msg:
            _meta_row("Message:", msg)

        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=8, pady=4)
        ttk.Label(win, text="Tickets in this bundle (select which to import):").pack(
            anchor="w", padx=12)

        # ticket list
        list_frame = ttk.Frame(win)
        list_frame.pack(fill="both", expand=True, padx=12, pady=(4, 4))

        cols = ("sel", "key", "summary", "type", "status")
        tree = ttk.Treeview(list_frame, columns=cols, show="headings",
                            selectmode="browse", height=12)
        tree.heading("sel",     text="✓",          anchor="center")
        tree.heading("key",     text="Key")
        tree.heading("summary", text="Summary")
        tree.heading("type",    text="Type")
        tree.heading("status",  text="Status")
        tree.column("sel",     width=30,  stretch=False, anchor="center")
        tree.column("key",     width=90,  stretch=False)
        tree.column("summary", width=260, stretch=True)
        tree.column("type",    width=90,  stretch=False)
        tree.column("status",  width=90,  stretch=False)

        vs = ttk.Scrollbar(list_frame, orient="vertical",   command=tree.yview)
        hs = ttk.Scrollbar(list_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        _checked = {}
        existing_keys = {
            str(it.get("Issue key") or "").strip()
            for it in self.list_items if it.get("Issue key")
        }

        for i, t in enumerate(tickets):
            if not isinstance(t, dict):
                continue
            key    = t.get("Issue key") or t.get("Issue id") or f"#{i+1}"
            summ   = (t.get("Summary") or "")[:80]
            itype  = t.get("Issue Type") or ""
            status = t.get("Status") or ""
            already = str(key).strip() in existing_keys
            var = tk.BooleanVar(value=not already)
            _checked[str(i)] = var
            mark = "☑" if var.get() else "☐"
            tag  = "exists" if already else ""
            tree.insert("", tk.END, iid=str(i),
                        values=(mark, key, summ, itype, status), tags=(tag,))

        tree.tag_configure("exists", foreground="#888")

        def _toggle(event):
            iid = tree.identify_row(event.y)
            if not iid:
                return
            var = _checked.get(iid)
            if var is None:
                return
            var.set(not var.get())
            tree.set(iid, "sel", "☑" if var.get() else "☐")
            _update_count()

        tree.bind("<Button-1>", _toggle)

        # select/deselect all
        sel_bar = ttk.Frame(win)
        sel_bar.pack(fill="x", padx=12, pady=(0, 2))
        ttk.Label(sel_bar, text="(grey = already in your list)",
                  font=("Segoe UI", 8), foreground="#888").pack(side="left")
        count_lbl = ttk.Label(sel_bar, text="")
        count_lbl.pack(side="right")

        def _update_count():
            n = sum(1 for v in _checked.values() if v.get())
            count_lbl.config(text=f"{n} ticket(s) will be imported")

        _update_count()

        def _set_all(val):
            for iid, var in _checked.items():
                var.set(val)
                tree.set(iid, "sel", "☑" if val else "☐")
            _update_count()

        btn_sel_all = ttk.Frame(win)
        btn_sel_all.pack(fill="x", padx=12, pady=(0, 6))
        ttk.Button(btn_sel_all, text="Select All",
                   command=lambda: _set_all(True)).pack(side="left", padx=(0, 4))
        ttk.Button(btn_sel_all, text="Deselect All",
                   command=lambda: _set_all(False)).pack(side="left")

        # ── action buttons ────────────────────────────────────────────────────
        btn_bar = ttk.Frame(win)
        btn_bar.pack(fill="x", padx=12, pady=(4, 12))

        def do_import():
            to_import = [tickets[int(iid)]
                         for iid, var in _checked.items() if var.get()
                         and int(iid) < len(tickets)]
            if not to_import:
                messagebox.showwarning("Nothing Selected",
                                       "Select at least one ticket to import.")
                return
            self.list_items.extend(to_import)
            self.list_items = _dedup_list_items(self.list_items)
            self.meta["fetched_issues"] = list(self.list_items)
            try:
                save_storage(self.templates, self.meta)
            except Exception as e:
                debug_log(f"Bundle import save error: {e}")
            try:
                self._rebuild_list_view()
            except Exception:
                pass
            win.destroy()
            messagebox.showinfo(
                "Imported",
                f"Added {len(to_import)} ticket(s) to your list.\n\n"
                "They are now available in the ticket list on the left.",
            )

        ttk.Button(btn_bar, text="⬇  Import Selected",
                   command=do_import).pack(side="left", padx=(0, 6))
        ttk.Button(btn_bar, text="Cancel",
                   command=win.destroy).pack(side="right")
