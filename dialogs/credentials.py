"""
CredentialsMixin — Jira API credential dialogs.
"""
import threading
import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser

from storage import save_storage
from utils import debug_log


class CredentialsMixin:
    """Mixin providing set_jira_credentials and clear_jira_credentials."""

    def set_jira_credentials(self, on_close=None):
        if self._focus_existing_app_dialog("credentials"):
            return
        dlg = tk.Toplevel(self)
        self._track_app_dialog("credentials", dlg)
        self._register_toplevel(dlg)
        dlg.title("Set Jira API Credentials")
        dlg.minsize(500, 220)
        dlg.geometry("880x260")
        dlg.resizable(True, True)

        def _close_and_callback():
            dlg.destroy()
            if on_close:
                self.after(50, on_close)
        ttk.Label(dlg, text="Base URL (e.g. https://your-domain.atlassian.net)").pack(anchor="w", padx=8, pady=(8, 0))
        base_var = tk.StringVar(value=(self.meta.get("jira", {}).get("base") or "https://wbg-avalanche.atlassian.net"))
        ttk.Entry(dlg, textvariable=base_var).pack(fill="x", padx=8, pady=4)
        ttk.Label(dlg, text="Email (Jira account email)").pack(anchor="w", padx=8, pady=(4, 0))
        email_var = tk.StringVar(value=self.meta.get("jira", {}).get("email", ""))
        ttk.Entry(dlg, textvariable=email_var).pack(fill="x", padx=8, pady=4)
        token_url = "https://id.atlassian.com/manage-profile/security/api-tokens"
        token_link_frame = ttk.Frame(dlg)
        token_link_frame.pack(anchor="w", padx=8, pady=(4, 0))
        ttk.Label(token_link_frame, text="API token — get yours at:").pack(side="left")
        link_lbl = ttk.Label(token_link_frame, text=token_url, style="Link.TLabel", cursor="hand2")
        link_lbl.pack(side="left", padx=(4, 0))
        link_lbl.bind("<Button-1>", lambda e: webbrowser.open(token_url))
        ttk.Label(dlg, text="API token:").pack(anchor="w", padx=8, pady=(4, 0))
        token_var = tk.StringVar(value=self.meta.get("jira", {}).get("token", ""))
        ttk.Entry(dlg, textvariable=token_var, show="*").pack(fill="x", padx=8, pady=4)
        def do_save():
            base = base_var.get().strip()
            email = email_var.get().strip()
            token = token_var.get().strip()
            if not base or not email or not token:
                messagebox.showerror("Error", "Please fill base URL, email and API token.")
                return
            self.meta.setdefault("jira", {})["base"] = base
            self.meta["jira"]["email"] = email
            self.meta["jira"]["token"] = token
            save_storage(self.templates, self.meta)
            messagebox.showinfo("Saved", "Jira credentials saved.")
            self.after(100, self._auto_fetch_jira_options)
            _close_and_callback()
        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(fill="x", pady=8, padx=8)
        ttk.Button(btn_frame, text="Save", command=do_save).pack(side="right", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=_close_and_callback).pack(side="right", padx=6)
        dlg.protocol("WM_DELETE_WINDOW", _close_and_callback)

    def _auto_fetch_jira_options(self):
        """Background-fetch common Jira field options after credentials are saved."""
        s = self.get_jira_session()
        if not s:
            return
        def worker():
            fetched = {}
            # Fetch project-independent options first
            for name, fn in [
                ("Project key", lambda: self._fetch_projects(s)),
                ("Priority", lambda: self._fetch_priorities(s)),
                ("Labels", lambda: self._fetch_labels(s)),
            ]:
                try:
                    vals = fn()
                    if vals:
                        fetched[name] = vals
                except Exception:
                    pass

            # Determine the primary project for project-scoped fields
            projects = fetched.get("Project key", [])
            primary_project = "SUNDANCE"
            if projects and primary_project not in projects:
                primary_project = projects[0]

            # Fetch project-scoped options using the primary project
            for name, fn in [
                ("Issue Type", lambda: self._fetch_issue_types(s, primary_project)),
                ("Status", lambda: self._fetch_statuses(s, primary_project)),
            ]:
                try:
                    vals = fn()
                    if vals:
                        fetched[name] = vals
                except Exception:
                    pass

            # Fetch Sprint and Fix Version (project-scoped)
            for name, fn in [
                ("Sprint", lambda: self._fetch_sprints(s, primary_project)),
                ("Fix Version", lambda: self._fetch_versions(s, primary_project)),
            ]:
                try:
                    vals = fn()
                    if vals:
                        fetched[name] = vals
                except Exception:
                    pass

            if projects:
                all_components = set()
                all_reporters = set()
                pks_to_scan = [primary_project] if primary_project in projects else []
                for pk in projects:
                    if pk not in pks_to_scan and len(pks_to_scan) < 5:
                        pks_to_scan.append(pk)
                for pk in pks_to_scan:
                    try:
                        comps = self._fetch_components(s, pk)
                        all_components.update(comps)
                    except Exception:
                        pass
                    try:
                        users = self._fetch_assignable_users(s, pk)
                        all_reporters.update(users)
                    except Exception:
                        pass
                if all_components:
                    fetched["Components"] = sorted(all_components, key=str.lower)
                if all_reporters:
                    fetched["Reporter"] = sorted(all_reporters, key=str.lower)
                    fetched["Assignee"] = sorted(all_reporters, key=str.lower)
            def apply():
                for name, vals in fetched.items():
                    self.meta.setdefault("options", {})[name] = vals
                if fetched:
                    save_storage(self.templates, self.meta)
                    debug_log(f"Auto-fetched Jira options: {list(fetched.keys())}")
            self.after(0, apply)
        threading.Thread(target=worker, daemon=True).start()

    def clear_jira_credentials(self):
        if messagebox.askyesno("Clear Jira API", "Clear stored Jira credentials?"):
            self.meta.setdefault("jira", {}).pop("base", None)
            self.meta.setdefault("jira", {}).pop("email", None)
            self.meta.setdefault("jira", {}).pop("token", None)
            save_storage(self.templates, self.meta)
            messagebox.showinfo("Cleared", "Stored Jira credentials cleared.")
