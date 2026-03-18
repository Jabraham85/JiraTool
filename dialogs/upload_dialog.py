"""
UploadDialogMixin — Upload bundle to Jira dialog.
"""
import tkinter as tk
from tkinter import ttk, messagebox


class UploadDialogMixin:
    """Mixin providing upload_bundle_to_jira_dialog."""

    def upload_bundle_to_jira_dialog(self):
        if self._focus_existing_app_dialog("upload_bundle"):
            return
        if not self.bundle:
            messagebox.showinfo("Info", "Bundle is empty.")
            return
        s = self.get_jira_session()
        if not s:
            return
        dlg = tk.Toplevel(self)
        self._track_app_dialog("upload_bundle", dlg)
        self._register_toplevel(dlg)
        dlg.title("Upload Bundle to Jira")
        dlg.minsize(450, 160)
        dlg.geometry("480x180")
        dlg.resizable(True, True)
        dlg
        ttk.Label(dlg, text="Project key, Issue Type, Status, Assignee, etc. are pulled from each ticket's Jira data.").pack(anchor="w", padx=8, pady=(8, 4))
        ttk.Label(dlg, text="Upload attachments? (attachments paths read from 'Attachment' field, semicolon-separated)").pack(anchor="w", padx=8, pady=(6,0))
        attach_choice = tk.BooleanVar(value=True)
        ttk.Checkbutton(dlg, text="Upload attachments", variable=attach_choice).pack(anchor="w", padx=8, pady=(0,6))
        def do_upload():
            upload_attachments = attach_choice.get()
            dlg.destroy()
            self.upload_bundle_to_jira(upload_attachments=upload_attachments)
        btns = ttk.Frame(dlg)
        btns.pack(fill="x", padx=8, pady=8)
        ttk.Button(btns, text="Upload", command=do_upload).pack(side="right", padx=6)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=6)
