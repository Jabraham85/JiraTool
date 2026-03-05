"""
UpdaterMixin — Check for app updates against the GitHub repository.
"""
import os
import sys
import json
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

import requests

from config import APP_VERSION, GITHUB_VERSION_URL
from storage import save_storage
from utils import debug_log


def _parse_version(v: str):
    """Convert '1.2.3' to a comparable tuple of ints."""
    try:
        return tuple(int(x) for x in v.strip().split("."))
    except Exception:
        return (0, 0, 0)


class UpdaterMixin:
    """Mixin providing GitHub-based update checking and downloading."""

    def _check_for_updates(self, manual: bool = False):
        """Check GitHub for a newer version.  Runs the network call in a
        background thread so the UI stays responsive.

        *manual* — when True (user clicked "Check for Updates"), ignore the
        skipped-version preference and always show the prompt.
        """
        self._log_startup("Checking for updates...", "step")

        def worker():
            try:
                resp = requests.get(GITHUB_VERSION_URL, timeout=10)
                resp.raise_for_status()
                data = resp.json()
            except Exception:
                debug_log("Update check failed: " + traceback.format_exc())
                self.after(0, lambda: self._log_startup(
                    "Update check: could not reach GitHub.", "done"))
                return

            remote_version = (data.get("version") or "").strip()
            download_url = (data.get("download_url") or "").strip()
            changelog = (data.get("changelog") or "").strip()

            if not remote_version:
                self.after(0, lambda: self._log_startup(
                    "Update check: invalid version data.", "done"))
                return

            local = _parse_version(APP_VERSION)
            remote = _parse_version(remote_version)

            if remote <= local:
                self.after(0, lambda: self._log_startup(
                    f"Up to date (v{APP_VERSION}).", "done"))
                if manual:
                    self.after(0, lambda: messagebox.showinfo(
                        "No Update",
                        f"You are running the latest version (v{APP_VERSION})."))
                return

            skipped = self.meta.get("skipped_update_version", "")
            if not manual and skipped == remote_version:
                self.after(0, lambda: self._log_startup(
                    f"Update v{remote_version} available (skipped by user).", "done"))
                return

            self.after(0, lambda rv=remote_version, dl=download_url, cl=changelog:
                       self._prompt_and_apply_update(rv, dl, cl))

        threading.Thread(target=worker, daemon=True).start()

    def _prompt_and_apply_update(self, remote_version, download_url, changelog):
        """Show an update dialog and, if accepted, download and replace the
        current executable."""
        self._log_startup(
            f"Update available: v{APP_VERSION} → v{remote_version}", "step")

        _BG     = "#1e1e1e"
        _PANEL  = "#252526"
        _BORDER = "#3c3c3c"
        _FG     = "#d4d4d4"
        _ACCENT = "#0e639c"

        dlg = tk.Toplevel(self)
        try:
            self._register_toplevel(dlg)
        except Exception:
            pass
        dlg.title("Update Available")
        dlg.configure(bg=_BG)
        dlg.minsize(460, 300)
        dlg.geometry("520x380")
        dlg.resizable(True, True)
        try:
            dlg.attributes("-topmost", True)
        except Exception:
            pass

        # Header
        hdr = tk.Frame(dlg, bg=_PANEL, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="A new version is available!",
                 bg=_PANEL, fg=_FG, font=("Segoe UI", 13, "bold"),
                 padx=16).pack(anchor="w")
        tk.Frame(dlg, bg=_BORDER, height=1).pack(fill="x")

        # Body
        body = tk.Frame(dlg, bg=_BG, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        tk.Label(body, text=f"Current version:   v{APP_VERSION}",
                 bg=_BG, fg="#888888", font=("Segoe UI", 10),
                 anchor="w").pack(fill="x")
        tk.Label(body, text=f"New version:        v{remote_version}",
                 bg=_BG, fg="#4ec9b0", font=("Segoe UI", 10, "bold"),
                 anchor="w").pack(fill="x", pady=(0, 10))

        tk.Label(body, text="What's new:", bg=_BG, fg=_FG,
                 font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x")
        cl_text = tk.Text(body, height=6, wrap="word", bg="#2d2d2d", fg=_FG,
                          font=("Segoe UI", 10), relief="flat", bd=0,
                          highlightthickness=1, highlightbackground=_BORDER,
                          padx=8, pady=6)
        cl_text.pack(fill="both", expand=True, pady=(4, 0))
        cl_text.insert("1.0", changelog or "No changelog provided.")
        cl_text.config(state="disabled")

        # Progress area (hidden until download starts)
        prog_frame = tk.Frame(dlg, bg=_BG, padx=16)
        prog_var = tk.StringVar(value="")
        prog_lbl = tk.Label(prog_frame, textvariable=prog_var, bg=_BG,
                            fg="#888888", font=("Segoe UI", 9), anchor="w")
        prog_lbl.pack(fill="x")
        prog_bar = ttk.Progressbar(prog_frame, mode="indeterminate")
        prog_bar.pack(fill="x", pady=(4, 0))

        # Footer
        tk.Frame(dlg, bg=_BORDER, height=1).pack(fill="x")
        footer = tk.Frame(dlg, bg=_PANEL, pady=10, padx=12)
        footer.pack(fill="x")

        def _make_btn(parent, text, cmd, bg_c, hover_c):
            b = tk.Button(parent, text=text, command=cmd, bg=bg_c,
                          fg="#ffffff", font=("Segoe UI", 9, "bold"),
                          relief="flat", bd=0, padx=16, pady=6,
                          cursor="hand2", activebackground=hover_c,
                          activeforeground="#ffffff")
            b.bind("<Enter>", lambda e: b.configure(bg=hover_c))
            b.bind("<Leave>", lambda e: b.configure(bg=bg_c))
            return b

        def _do_update():
            if not download_url:
                messagebox.showerror("Update Error",
                                     "No download URL provided in the update.")
                return
            update_btn.config(state="disabled")
            skip_btn.config(state="disabled")
            prog_frame.pack(fill="x", pady=(0, 4))
            prog_bar.start(10)
            prog_var.set("Downloading update...")
            dlg.update_idletasks()

            def download_worker():
                try:
                    self._download_and_replace(
                        download_url, remote_version, prog_var, dlg)
                except Exception:
                    debug_log("Update download failed: "
                              + traceback.format_exc())
                    def _show_err():
                        prog_bar.stop()
                        prog_var.set("Download failed.")
                        update_btn.config(state="normal")
                        skip_btn.config(state="normal")
                        messagebox.showerror(
                            "Update Failed",
                            "Could not download the update. "
                            "Check your internet connection and try again.")
                    self.after(0, _show_err)

            threading.Thread(target=download_worker, daemon=True).start()

        def _do_skip():
            self.meta["skipped_update_version"] = remote_version
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass
            self._log_startup(
                f"Update v{remote_version} skipped.", "done")
            dlg.destroy()

        update_btn = _make_btn(footer, "Download & Install",
                               _do_update, _ACCENT, "#1177bb")
        update_btn.pack(side="right", padx=(8, 0))
        skip_btn = _make_btn(footer, "Skip This Version",
                             _do_skip, "#3c3c3c", "#505050")
        skip_btn.pack(side="right")
        _make_btn(footer, "Remind Me Later",
                  dlg.destroy, "#3c3c3c", "#505050").pack(side="right",
                                                           padx=(0, 4))

        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def _download_and_replace(self, download_url, remote_version,
                              prog_var, dlg):
        """Download the update file and replace the current executable."""
        current_exe = sys.executable
        is_frozen = getattr(sys, "frozen", False)

        if is_frozen:
            target_path = current_exe
        else:
            target_path = os.path.abspath(sys.argv[0]) if sys.argv else ""

        if not target_path:
            self.after(0, lambda: messagebox.showerror(
                "Update Error", "Could not determine the application path."))
            return

        target_dir = os.path.dirname(target_path)
        target_name = os.path.basename(target_path)
        base, ext = os.path.splitext(target_name)

        new_path = os.path.join(target_dir, f"{base}.new{ext}")
        old_path = os.path.join(target_dir, f"{base}.old{ext}")

        # Clean up leftover files from a previous update
        for leftover in (new_path, old_path):
            try:
                if os.path.exists(leftover):
                    os.remove(leftover)
            except Exception:
                pass

        self.after(0, lambda: prog_var.set("Downloading..."))

        resp = requests.get(download_url, stream=True, timeout=120)
        resp.raise_for_status()
        total = int(resp.headers.get("Content-Length", 0))
        downloaded = 0

        with open(new_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total:
                        pct = int(downloaded / total * 100)
                        self.after(0, lambda p=pct: prog_var.set(
                            f"Downloading... {p}%"))

        self.after(0, lambda: prog_var.set("Download complete. Applying..."))

        if is_frozen:
            # Swap: current -> .old, new -> current
            try:
                if os.path.exists(old_path):
                    os.remove(old_path)
                os.rename(target_path, old_path)
                os.rename(new_path, target_path)
            except Exception:
                debug_log("EXE swap failed: " + traceback.format_exc())
                # Try to restore
                try:
                    if not os.path.exists(target_path) and os.path.exists(old_path):
                        os.rename(old_path, target_path)
                except Exception:
                    pass
                self.after(0, lambda: messagebox.showerror(
                    "Update Failed",
                    "Could not replace the executable. "
                    "The downloaded file is saved as:\n"
                    f"{new_path}\n\n"
                    "You can manually replace the EXE."))
                return

            self.meta["skipped_update_version"] = ""
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass

            def _auto_restart():
                try:
                    dlg.destroy()
                except Exception:
                    pass

                # Fully shut down the old app before launching the new one
                try:
                    self.destroy()
                except Exception:
                    pass

                import subprocess
                try:
                    subprocess.Popen([target_path], close_fds=True,
                                     creationflags=subprocess.DETACHED_PROCESS)
                except Exception:
                    try:
                        os.startfile(target_path)
                    except Exception:
                        pass

                os._exit(0)
            self.after(0, _auto_restart)
        else:
            # Running from source — save the downloaded file and inform user
            self.meta["skipped_update_version"] = ""
            try:
                save_storage(self.templates, self.meta)
            except Exception:
                pass

            def _notify():
                try:
                    dlg.destroy()
                except Exception:
                    pass
                messagebox.showinfo(
                    "Update Downloaded",
                    f"Version v{remote_version} has been downloaded to:\n"
                    f"{new_path}\n\n"
                    "Replace the current files manually, or pull the "
                    "latest changes from GitHub.")
            self.after(0, _notify)
