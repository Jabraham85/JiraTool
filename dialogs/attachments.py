"""
AttachmentsMixin — Jira attachment gallery dialog.
Shows image thumbnails alongside a file list, lets users preview full size
or open/save any attachment.
"""
import io
import os
import json
import tempfile
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from ..utils import debug_log

_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif", ".ico"}


def _is_image(attachment: dict) -> bool:
    mime = (attachment.get("mimeType") or "").lower()
    if mime.startswith("image/"):
        return True
    ext = os.path.splitext(attachment.get("filename") or "")[1].lower()
    return ext in _IMAGE_EXTS


def _format_size(size: int) -> str:
    if size >= 1_048_576:
        return f"{size / 1_048_576:.1f} MB"
    if size >= 1_024:
        return f"{size / 1_024:.0f} KB"
    return f"{size} B"


class AttachmentsMixin:
    """Mixin providing _open_attachments_dialog."""

    def _open_attachments_dialog(self, tabform):
        """Open a gallery dialog for all Jira attachments on the current tab."""
        data = tabform.read_to_dict()
        val = data.get("Attachment", "") or ""
        if not val or not str(val).strip().startswith("["):
            messagebox.showinfo(
                "Attachments",
                "No Jira attachments on this ticket.\n\n"
                "Use the Attachment field (semicolon-separated local paths) to attach "
                "files when uploading a new ticket.",
            )
            return
        try:
            items = json.loads(val)
        except Exception:
            messagebox.showinfo("Attachments", "Could not parse attachment data.")
            return
        if not isinstance(items, list) or not items:
            messagebox.showinfo("Attachments", "No attachments to show.")
            return
        s = self.get_jira_session()
        if not s:
            messagebox.showinfo("Info", "Set Jira API credentials to view attachments.")
            return

        # ── window ──────────────────────────────────────────────────────────
        win = tk.Toplevel(self)
        self._register_toplevel(win)
        ticket_key = data.get("Issue key") or data.get("Summary") or "Ticket"
        win.title(f"Attachments — {ticket_key}")
        win.minsize(700, 480)
        win.geometry("860x560")
        win.resizable(True, True)

        # top bar
        top = ttk.Frame(win)
        top.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(top, text=f"{len(items)} attachment(s)", font=("Segoe UI", 9)).pack(side="left")

        # main pane: left list | right preview
        pane = ttk.PanedWindow(win, orient="horizontal")
        pane.pack(fill="both", expand=True, padx=8, pady=4)

        # ── left: file list ──────────────────────────────────────────────────
        left = ttk.Frame(pane)
        pane.add(left, weight=1)

        cols = ("name", "size", "type")
        tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse", height=18)
        tree.heading("name", text="Filename")
        tree.heading("size", text="Size")
        tree.heading("type", text="Type")
        tree.column("name", width=180, stretch=True)
        tree.column("size", width=68,  anchor="e", stretch=False)
        tree.column("type", width=80,  stretch=False)
        vs = ttk.Scrollbar(left, orient="vertical",   command=tree.yview)
        hs = ttk.Scrollbar(left, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)

        for i, a in enumerate(items):
            if not isinstance(a, dict):
                continue
            fn   = a.get("filename") or a.get("name") or "(unknown)"
            sz   = _format_size(int(a.get("size") or 0))
            mime = a.get("mimeType") or os.path.splitext(fn)[1].lstrip(".").upper() or "—"
            tag  = "image" if _is_image(a) else "file"
            tree.insert("", tk.END, iid=str(i), values=(fn, sz, mime), tags=(tag,))

        tree.tag_configure("image", foreground="#4fc3f7")

        # ── right: preview ───────────────────────────────────────────────────
        right = ttk.Frame(pane)
        pane.add(right, weight=2)

        PREVIEW_W, PREVIEW_H = 520, 380

        preview_canvas = tk.Canvas(right, bg="#1e1e1e", bd=0, highlightthickness=0,
                                   width=PREVIEW_W, height=PREVIEW_H)
        preview_canvas.pack(fill="both", expand=True, padx=4, pady=4)

        info_lbl = ttk.Label(right, text="", font=("Segoe UI", 9), wraplength=500)
        info_lbl.pack(anchor="w", padx=8, pady=(0, 4))

        # state for the currently displayed image
        _state = {"photo": None, "loading": False}

        def _clear_preview(msg="Select an attachment to preview"):
            preview_canvas.delete("all")
            cw = preview_canvas.winfo_width()  or PREVIEW_W
            ch = preview_canvas.winfo_height() or PREVIEW_H
            preview_canvas.create_text(
                cw // 2, ch // 2, text=msg,
                fill="#555", font=("Segoe UI", 11), anchor="center"
            )
            _state["photo"] = None

        _clear_preview()

        def _show_image_bytes(raw: bytes, filename: str):
            """Decode raw bytes with Pillow and display in the canvas."""
            try:
                from PIL import Image, ImageTk
                img = Image.open(io.BytesIO(raw))
                img.thumbnail((PREVIEW_W, PREVIEW_H), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                _state["photo"] = photo          # keep a reference
                cw = preview_canvas.winfo_width()  or PREVIEW_W
                ch = preview_canvas.winfo_height() or PREVIEW_H
                preview_canvas.delete("all")
                preview_canvas.create_image(cw // 2, ch // 2, image=photo, anchor="center")
            except Exception:
                _clear_preview("Could not render image.")
                debug_log("Attachment preview error: " + traceback.format_exc())

        def _load_preview_bg(attachment: dict):
            """Background thread: download thumbnail (or full image) and display."""
            _state["loading"] = True
            url = attachment.get("thumbnail") or attachment.get("content") or ""
            if not url:
                win.after(0, lambda: _clear_preview("No URL available."))
                _state["loading"] = False
                return
            try:
                r = s.get(url, timeout=20)
                if r.status_code != 200:
                    win.after(0, lambda: _clear_preview(f"Download failed ({r.status_code})"))
                    return
                raw = r.content
                win.after(0, lambda raw=raw, fn=attachment.get("filename", ""):
                          _show_image_bytes(raw, fn))
            except Exception:
                win.after(0, lambda: _clear_preview("Download error."))
                debug_log("Attachment preview download error: " + traceback.format_exc())
            finally:
                _state["loading"] = False

        def _on_select(event=None):
            sel = tree.selection()
            if not sel:
                return
            idx = int(sel[0])
            if idx < 0 or idx >= len(items):
                return
            a = items[idx]
            if not isinstance(a, dict):
                return
            fn   = a.get("filename") or a.get("name") or "(unknown)"
            sz   = _format_size(int(a.get("size") or 0))
            mime = a.get("mimeType") or "—"
            info_lbl.config(text=f"{fn}  ·  {sz}  ·  {mime}")
            if _is_image(a):
                _clear_preview("Loading…")
                threading.Thread(target=_load_preview_bg, args=(a,), daemon=True).start()
            else:
                _clear_preview(f"📄  {fn}\n\nNot an image — use Open to view this file.")

        tree.bind("<<TreeviewSelect>>", _on_select)

        # ── bottom buttons ───────────────────────────────────────────────────
        btn_bar = ttk.Frame(win)
        btn_bar.pack(fill="x", padx=8, pady=(4, 8))

        def _selected_attachment():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Info", "Select an attachment first.")
                return None
            idx = int(sel[0])
            return items[idx] if 0 <= idx < len(items) else None

        def do_open():
            a = _selected_attachment()
            if not a:
                return
            url = a.get("content") or ""
            fn  = a.get("filename") or a.get("name") or "attachment"
            if not url:
                messagebox.showinfo("Info", "No download URL for this attachment.")
                return
            try:
                r = s.get(url, timeout=30)
                if r.status_code != 200:
                    messagebox.showerror("Error", f"Download failed: {r.status_code}")
                    return
                ext = os.path.splitext(fn)[1] or ""
                fd, path = tempfile.mkstemp(suffix=ext)
                try:
                    os.write(fd, r.content)
                finally:
                    os.close(fd)
                if os.name == "nt":
                    os.startfile(path)
                else:
                    import subprocess
                    subprocess.run(
                        ["xdg-open", path] if os.path.exists("/usr/bin/xdg-open") else ["open", path],
                        check=False,
                    )
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open: {e}")

        def do_save_as():
            a = _selected_attachment()
            if not a:
                return
            url = a.get("content") or ""
            fn  = a.get("filename") or a.get("name") or "attachment"
            if not url:
                messagebox.showinfo("Info", "No download URL for this attachment.")
                return
            dest = filedialog.asksaveasfilename(initialfile=fn, title="Save attachment as")
            if not dest:
                return
            try:
                r = s.get(url, timeout=60)
                if r.status_code != 200:
                    messagebox.showerror("Error", f"Download failed: {r.status_code}")
                    return
                with open(dest, "wb") as f:
                    f.write(r.content)
                messagebox.showinfo("Saved", f"Saved to {dest}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {e}")

        def do_view_full():
            """Open the full-resolution image in its own resizable window."""
            a = _selected_attachment()
            if not a:
                return
            if not _is_image(a):
                messagebox.showinfo("Info", "Select an image attachment to view full size.")
                return
            url = a.get("content") or ""
            fn  = a.get("filename") or a.get("name") or "image"
            if not url:
                messagebox.showinfo("Info", "No download URL for this attachment.")
                return

            def _open_full():
                try:
                    from PIL import Image, ImageTk
                    r = s.get(url, timeout=60)
                    if r.status_code != 200:
                        win.after(0, lambda: messagebox.showerror(
                            "Error", f"Download failed: {r.status_code}"))
                        return
                    img = Image.open(io.BytesIO(r.content))
                    # Fit inside 80 % of screen size
                    sw = int(win.winfo_screenwidth()  * 0.8)
                    sh = int(win.winfo_screenheight() * 0.8)
                    img.thumbnail((sw, sh), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)

                    def _show():
                        fw = tk.Toplevel(win)
                        fw.title(fn)
                        fw.resizable(True, True)
                        c = tk.Canvas(fw, bg="#1e1e1e", bd=0, highlightthickness=0,
                                      width=img.width, height=img.height)
                        c.pack(fill="both", expand=True)
                        c._photo = photo            # keep reference
                        c.create_image(img.width // 2, img.height // 2,
                                       image=photo, anchor="center")
                        ttk.Button(fw, text="Close", command=fw.destroy).pack(pady=6)

                    win.after(0, _show)
                except Exception:
                    win.after(0, lambda: messagebox.showerror(
                        "Error", "Could not load full image."))
                    debug_log("Full image view error: " + traceback.format_exc())

            threading.Thread(target=_open_full, daemon=True).start()

        ttk.Button(btn_bar, text="Open in App",  command=do_open).pack(side="left", padx=4)
        ttk.Button(btn_bar, text="Save As…",     command=do_save_as).pack(side="left", padx=4)
        ttk.Button(btn_bar, text="View Full Size", command=do_view_full).pack(side="left", padx=4)
        ttk.Button(btn_bar, text="Close",         command=win.destroy).pack(side="right", padx=4)
