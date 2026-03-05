"""
Upload mixin: bundle management and Jira upload.
"""
import json
import copy
import os
import uuid
import traceback
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

from config import FETCH_FIELDS, HEADERS
from storage import perform_jira_request, save_storage
from utils import debug_log, _dedup_list_items


class UploadMixin:
    # ---------------- Bundle features ----------------
    def update_bundle_listbox(self):
        if not hasattr(self, "bundle_listbox"):
            return
        self.bundle_listbox.delete(0, tk.END)
        for i, t in enumerate(self.bundle, start=1):
            summary = (t.get("Summary") or "")[:80].replace("\n", " ")
            key = t.get("Issue key", "")
            label = f"{i}. {key} — {summary}" if key else f"{i}. {summary}"
            self.bundle_listbox.insert(tk.END, label)
        count = len(self.bundle)
        self.title(f"Avalanche Jira Template Creator ({count} in bundle)")

    def add_active_tab_to_bundle(self):
        tf = self.get_active_tabform()
        if not tf:
            messagebox.showinfo("Info", "No active tab to add.")
            return
        ticket = tf.read_to_dict()
        if not ticket.get("Issue id"):
            ticket["Issue id"] = str(uuid.uuid4())
        if not ticket.get("Issue key"):
            ticket["Issue key"] = "LOCAL-" + ticket["Issue id"][:8]
        for h in HEADERS:
            if h == "Description ADF":
                continue
            info = tf.field_widgets.get(h)
            if info and not info.get("include_var", tk.BooleanVar(value=True)).get():
                ticket[h] = ""
        self.bundle.append(dict(ticket))
        self.update_bundle_listbox()
        if not getattr(self, "_tutorial_running", False):
            messagebox.showinfo("Bundled", f"Added ticket to bundle ({len(self.bundle)} total)")

    def _open_selected_bundle_in_jira(self):
        """Open selected bundle item in Jira (browser)."""
        sel = self.bundle_listbox.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select a bundle item to open in Jira.")
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.bundle):
            ticket = self.bundle[idx]
            key = ticket.get("Issue key") or ticket.get("Issue id")
            if key and str(key).strip() and not str(key).strip().startswith("LOCAL-"):
                self._open_in_jira_browser(key)
            else:
                messagebox.showinfo("Info", "This bundle item has no Jira key (local only).")

    def remove_selected_from_bundle(self):
        sel = self.bundle_listbox.curselection()
        if not sel:
            messagebox.showinfo("Info", "Select an item to remove from bundle.")
            return
        idx = sel[0]
        del self.bundle[idx]
        self.update_bundle_listbox()

    def clear_bundle(self):
        if not self.bundle:
            return
        if messagebox.askyesno("Confirm", f"Clear {len(self.bundle)} tickets from bundle?"):
            self.bundle = []
            self.update_bundle_listbox()

    def rename_bundle(self):
        name = simpledialog.askstring("Bundle Name", "Enter a name for this bundle (optional):", initialvalue=self.bundle_name or "")
        if name is None:
            return
        self.bundle_name = name.strip() or None
        self.update_bundle_listbox()

    def _post_new_comments(self, s, issue_key: str, ticket: dict) -> None:
        """POST any local (unposted) comments for a ticket to Jira."""
        comment_raw = ticket.get("Comment") or ""
        if not comment_raw:
            return
        try:
            comments = json.loads(comment_raw) if isinstance(comment_raw, str) else comment_raw
            if not isinstance(comments, list):
                return
            comment_url = f"{s._jira_base}/rest/api/3/issue/{issue_key}/comment"
            for c in comments:
                if c.get("posted"):
                    continue
                body_text = (c.get("body") or "").strip()
                if not body_text:
                    continue
                payload = {"body": self._text_to_adf(body_text)}
                try:
                    cr = perform_jira_request(s, "POST", comment_url,
                                              json_body=payload, timeout=30)
                    if cr.status_code == 201:
                        c["id"]     = cr.json().get("id", "")
                        c["posted"] = True
                        debug_log(f"Posted comment to {issue_key}: {body_text[:60]}")
                    else:
                        debug_log(f"Comment POST failed for {issue_key}: "
                                  f"{cr.status_code} {cr.text[:200]}")
                except Exception:
                    debug_log(f"Comment POST exception for {issue_key}: "
                              + traceback.format_exc())
            # Write back the updated (now-posted) comment list
            ticket["Comment"] = json.dumps(comments, ensure_ascii=False)
        except Exception:
            debug_log("_post_new_comments failed: " + traceback.format_exc())

    def _post_new_issue_links(self, s, issue_key: str, ticket: dict) -> None:
        """POST any locally-added (unposted) issue links for a ticket to Jira."""
        links_raw = ticket.get("Issue Links") or ""
        if not links_raw:
            return
        try:
            links = json.loads(links_raw) if isinstance(links_raw, str) else links_raw
            if not isinstance(links, list):
                return
            link_url = f"{s._jira_base}/rest/api/3/issueLink"
            for lnk in links:
                if lnk.get("posted"):
                    continue
                key_val = (lnk.get("key") or "").strip()
                if not key_val:
                    continue
                type_name = lnk.get("type_name") or "relates to"
                direction = lnk.get("direction") or "outward"
                if direction == "inward":
                    payload = {
                        "type":         {"name": type_name},
                        "inwardIssue":  {"key": issue_key},
                        "outwardIssue": {"key": key_val},
                    }
                else:
                    payload = {
                        "type":         {"name": type_name},
                        "outwardIssue": {"key": issue_key},
                        "inwardIssue":  {"key": key_val},
                    }
                try:
                    lr = perform_jira_request(s, "POST", link_url,
                                              json_body=payload, timeout=30)
                    if lr.status_code == 201:
                        lnk["posted"] = True
                        lnk["id"] = (lr.json() or {}).get("id", "")
                        debug_log(f"Posted issue link {issue_key} {type_name} {key_val}")
                    else:
                        debug_log(f"Issue link POST failed for {issue_key}: "
                                  f"{lr.status_code} {lr.text[:200]}")
                except Exception:
                    debug_log(f"Issue link POST exception for {issue_key}: "
                              + traceback.format_exc())
            ticket["Issue Links"] = json.dumps(links, ensure_ascii=False)
        except Exception:
            debug_log("_post_new_issue_links failed: " + traceback.format_exc())

    def upload_bundle_to_jira(self, project_key_override=None, issue_type_override=None, parent_issue_key_override=None, assignee_override=None, upload_attachments=True):
        s = self.get_jira_session()
        if not s:
            return
        # Count creates vs updates for confirm
        create_count = 0
        update_count = 0
        for ticket in self.bundle:
            key = str(ticket.get("Issue key") or ticket.get("Issue id") or "").strip()
            if key and not key.startswith("LOCAL-"):
                update_count += 1
            else:
                create_count += 1
        confirm_msg = []
        if create_count:
            confirm_msg.append(f"create {create_count} new")
        if update_count:
            confirm_msg.append(f"update {update_count} existing")
        if not confirm_msg:
            return
        if not messagebox.askyesno("Confirm upload", f"This will {' and '.join(confirm_msg)} issue(s) in Jira. Proceed?"):
            return
        successes = []
        failures = []
        summary_mismatch_excluded = []
        for i, ticket in enumerate(self.bundle, start=1):
            ticket_for_upload = copy.deepcopy(ticket)
            # Recover ADF if missing (e.g. ticket added from list with empty ADF)
            adf_check = ticket_for_upload.get("Description ADF")
            if not adf_check or not (isinstance(adf_check, dict) and adf_check.get("content")):
                recovered = self._recover_adf_for_ticket(ticket_for_upload)
                if recovered:
                    ticket_for_upload["Description ADF"] = copy.deepcopy(recovered)
                    ticket["Description ADF"] = copy.deepcopy(recovered)
            self._apply_variables_to_ticket(ticket_for_upload)
            summary = ticket_for_upload.get("Summary") or ticket_for_upload.get("Issue key") or f"Bundle Ticket {i}"
            description_plain = ticket_for_upload.get("Description") or ""
            project_key = (ticket_for_upload.get("Project key") or project_key_override)
            issuetype = ticket_for_upload.get("Issue Type") or issue_type_override or "Task"
            labels_raw = (ticket_for_upload.get("Labels") or "").strip()
            if ";" in labels_raw:
                labels = [l.strip() for l in labels_raw.split(";") if l.strip()]
            elif "," in labels_raw:
                labels = [l.strip() for l in labels_raw.split(",") if l.strip()]
            else:
                labels = [labels_raw] if labels_raw else []
            debug_log(f"Parsed labels for ticket {i}: {labels} (raw={labels_raw!r})")
            comps = [ {"name": c.strip()} for c in (ticket_for_upload.get("Components") or "").split(";") if c.strip() ]
            priority = ticket_for_upload.get("Priority") or None
            assignee_val = ticket_for_upload.get("Assignee") or assignee_override or None
            parent_key = ticket_for_upload.get("Parent key") or parent_issue_key_override or None
            epic_key   = (ticket_for_upload.get("Epic Link") or "").strip()
            epic_mode  = (ticket_for_upload.get("_epic_mode") or "nextgen").strip()
            issue_key = str(ticket.get("Issue key") or ticket.get("Issue id") or "").strip()
            is_update = bool(issue_key and not issue_key.startswith("LOCAL-"))
            existing = None
            if is_update:
                # Verify issue exists in Jira before updating
                try:
                    existing = self.fetch_issue_details(s, issue_key, fields=["summary", "project", "issuetype", "status", "reporter", "created", "updated"])
                    jira_summary = (existing.get("fields") or {}).get("summary") or ""
                    if jira_summary.strip() != summary.strip():
                        summary_mismatch_excluded.append(issue_key)
                        failures.append((ticket, "skipped", "Summary mismatch - excluded"))
                        continue
                except Exception as e:
                    if "404" in str(e) or "not found" in str(e).lower():
                        debug_log(f"Issue {issue_key} not found (404), falling back to create new")
                        is_update = False
                    else:
                        debug_log(f"Fetch failed for {issue_key}: {e}")
                        failures.append((ticket, "fetch_failed", str(e)))
                        continue
            if is_update:
                # Build update payload (only editable fields)
                update_fields = {}
                update_fields["summary"] = summary
                _resolved_adf_for_ticket = None
                # Resolve any pending inline images before building the ADF payload
                adf_raw = ticket_for_upload.get("Description ADF")
                if adf_raw:
                    adf_copy = copy.deepcopy(adf_raw) if isinstance(adf_raw, dict) else adf_raw
                    if isinstance(adf_copy, str):
                        try:
                            adf_copy = json.loads(adf_copy)
                        except Exception:
                            adf_copy = None
                    if isinstance(adf_copy, dict):
                        adf_copy, _ = self._resolve_pending_media(s, issue_key, adf_copy)
                        self._strip_custom_media_attrs(adf_copy)
                        self._remove_invalid_media_nodes(adf_copy)
                    if adf_copy:
                        sanitized_desc = self._sanitize_adf_for_upload(adf_copy)
                        debug_log(f"ADF description for update {issue_key}: {json.dumps(sanitized_desc, default=str)[:2000]}")
                        update_fields["description"] = sanitized_desc
                        _resolved_adf_for_ticket = sanitized_desc
                    else:
                        _resolved_adf_for_ticket = None
                elif description_plain and str(description_plain).strip():
                    update_fields["description"] = self._text_to_adf(description_plain)
                    _resolved_adf_for_ticket = None
                if labels:
                    update_fields["labels"] = labels
                if comps:
                    update_fields["components"] = comps
                if priority:
                    update_fields["priority"] = {"name": priority}
                if assignee_val:
                    jira_project = (existing.get("fields") or {}).get("project", {}).get("key") or project_key
                    acct = self._resolve_assignee(s, assignee_val, project_key=jira_project)
                    if acct:
                        update_fields["assignee"] = {"accountId": acct}
                # Epic relationship on update
                if epic_key:
                    if epic_mode == "classic":
                        update_fields["customfield_10014"] = epic_key
                    else:
                        update_fields["parent"] = {"key": epic_key}
                elif parent_key:
                    # True sub-task parent (non-epic)
                    update_fields["parent"] = {"key": parent_key}
                update_url = f"{s._jira_base}/rest/api/3/issue/{issue_key}"
                payload = {"fields": update_fields}
                try:
                    resp = perform_jira_request(s, "PUT", update_url, json_body=payload, timeout=60)
                    if resp.status_code in (200, 204):
                        self._apply_keyable_fields_from_jira(ticket, existing)
                        if _resolved_adf_for_ticket:
                            ticket["Description ADF"] = _resolved_adf_for_ticket
                        successes.append((ticket, issue_key))
                        debug_log(f"Updated issue {issue_key} for bundle item {i}")
                        self._post_new_comments(s, issue_key, ticket)
                        self._post_new_issue_links(s, issue_key, ticket)
                        if upload_attachments:
                            attachment_field = ticket.get("Attachment", "") or ""
                            if attachment_field and not str(attachment_field).strip().startswith("["):
                                paths = [p.strip() for p in attachment_field.split(";") if p.strip()]
                                for p in paths:
                                    if not os.path.exists(p):
                                        debug_log(f"Attachment not found: {p}")
                                        continue
                                    try:
                                        attach_url = f"{s._jira_base}/rest/api/3/issue/{issue_key}/attachments"
                                        with open(p, "rb") as fh:
                                            files = {"file": (os.path.basename(p), fh)}
                                            headers = {"X-Atlassian-Token": "no-check"}
                                            r2 = s.post(attach_url, files=files, headers=headers)
                                            if getattr(r2, "status_code", None) not in (200, 201):
                                                debug_log(f"Attachment failed: {getattr(r2, 'status_code', '')} {getattr(r2, 'text', '')}")
                                    except Exception:
                                        debug_log(f"Attachment upload exception: {traceback.format_exc()}")
                    else:
                        debug_log(f"Update issue failed for bundle item {i}: {getattr(resp, 'status_code', '')} {getattr(resp, 'text', '')}")
                        failures.append((ticket, getattr(resp, 'status_code', 'N/A'), getattr(resp, 'text', '')))
                except Exception:
                    debug_log(f"Exception updating issue for bundle item {i}: {traceback.format_exc()}")
                    failures.append((ticket, "exception", traceback.format_exc()))
            if not is_update:
                # Create new issue
                fields = {}
                if project_key:
                    fields["project"] = {"key": project_key}
                fields["summary"] = summary
                fields["issuetype"] = {"name": issuetype}
                # For create: check if ADF has pending images — if so, strip
                # them initially and we'll upload + patch after creation.
                has_pending_media = False
                adf_for_create = None
                if ticket_for_upload.get("Description ADF"):
                    adf_for_create = copy.deepcopy(ticket_for_upload["Description ADF"]) if isinstance(ticket_for_upload["Description ADF"], dict) else ticket_for_upload["Description ADF"]
                    if isinstance(adf_for_create, str):
                        try:
                            adf_for_create = json.loads(adf_for_create)
                        except Exception:
                            adf_for_create = None
                    if isinstance(adf_for_create, dict):
                        has_pending_media = self._adf_has_pending_media(adf_for_create)
                        self._strip_custom_media_attrs(adf_for_create)
                        # Remove media nodes entirely for the initial create
                        # (they have no valid id yet; we'll patch after uploading)
                        self._remove_invalid_media_nodes(adf_for_create)
                    if adf_for_create:
                        fields["description"] = self._sanitize_adf_for_upload(adf_for_create)
                else:
                    if description_plain and str(description_plain).strip():
                        fields["description"] = self._text_to_adf(description_plain)
                if labels:
                    fields["labels"] = labels
                if comps:
                    fields["components"] = comps
                if priority:
                    fields["priority"] = {"name": priority}
                if assignee_val:
                    acct = self._resolve_assignee(s, assignee_val, project_key=project_key)
                    if acct:
                        fields["assignee"] = {"accountId": acct}
                    else:
                        debug_log(f"Assignee resolution failed for ticket {i}: '{assignee_val}' - will omit assignee field")
                # Epic relationship on create (takes priority over plain parent)
                if epic_key:
                    if epic_mode == "classic":
                        fields["customfield_10014"] = epic_key
                    else:
                        fields["parent"] = {"key": epic_key}
                elif parent_key:
                    fields["parent"] = {"key": parent_key}
                create_url = f"{s._jira_base}/rest/api/3/issue"
                payload = {"fields": fields}
                try:
                    resp = perform_jira_request(s, "POST", create_url, json_body=payload, timeout=60)
                    if resp.status_code in (201,):
                        created = resp.json()
                        new_key = created.get("key") or created.get("id")
                        try:
                            fresh = self.fetch_issue_details(s, new_key, fields=["summary", "project", "issuetype", "status", "reporter", "created", "updated"])
                            self._apply_keyable_fields_from_jira(ticket, fresh)
                        except Exception as e:
                            debug_log(f"Fetch after create failed for {new_key}: {e}")
                            ticket["Issue key"] = new_key
                            ticket["Issue id"] = created.get("id", ticket.get("Issue id"))
                        successes.append((ticket, new_key))
                        debug_log(f"Created issue {new_key} for bundle item {i}")
                        # Upload pending inline images and update description
                        if has_pending_media:
                            try:
                                adf_src = ticket_for_upload.get("Description ADF")
                                if isinstance(adf_src, str):
                                    adf_src = json.loads(adf_src)
                                adf_resolved = copy.deepcopy(adf_src) if isinstance(adf_src, dict) else None
                                if adf_resolved:
                                    adf_resolved, did_upload = self._resolve_pending_media(s, new_key, adf_resolved)
                                    if did_upload:
                                        self._strip_custom_media_attrs(adf_resolved)
                                        self._remove_invalid_media_nodes(adf_resolved)
                                        sanitized = self._sanitize_adf_for_upload(adf_resolved)
                                        desc_payload = {"fields": {"description": sanitized}}
                                        patch_url = f"{s._jira_base}/rest/api/3/issue/{new_key}"
                                        perform_jira_request(s, "PUT", patch_url, json_body=desc_payload, timeout=60)
                                        ticket["Description ADF"] = sanitized
                                        debug_log(f"Patched description with resolved images for {new_key}")
                            except Exception:
                                debug_log(f"Failed to patch images for {new_key}: {traceback.format_exc()}")
                        self._post_new_comments(s, new_key, ticket)
                        self._post_new_issue_links(s, new_key, ticket)
                        if upload_attachments:
                            attachment_field = ticket.get("Attachment", "") or ""
                            if attachment_field and not str(attachment_field).strip().startswith("["):
                                paths = [p.strip() for p in attachment_field.split(";") if p.strip()]
                                for p in paths:
                                    if not os.path.exists(p):
                                        debug_log(f"Attachment not found: {p}")
                                        continue
                                    try:
                                        attach_url = f"{s._jira_base}/rest/api/3/issue/{new_key}/attachments"
                                        with open(p, "rb") as fh:
                                            files = {"file": (os.path.basename(p), fh)}
                                            headers = {"X-Atlassian-Token": "no-check"}
                                            r2 = s.post(attach_url, files=files, headers=headers)
                                            if getattr(r2, "status_code", None) not in (200, 201):
                                                debug_log(f"Attachment failed: {getattr(r2, 'status_code', '')} {getattr(r2, 'text', '')}")
                                    except Exception:
                                        debug_log(f"Attachment upload exception: {traceback.format_exc()}")
                    else:
                        debug_log(f"Create issue failed for bundle item {i}: {getattr(resp, 'status_code', '')} {getattr(resp, 'text', '')}")
                        failures.append((ticket, getattr(resp, 'status_code', 'N/A'), getattr(resp, 'text', '')))
                except Exception:
                    debug_log(f"Exception creating issue for bundle item {i}: {traceback.format_exc()}")
                    failures.append((ticket, "exception", traceback.format_exc()))
        try:
            self.meta.setdefault("user_cache", {}).update(self._user_cache)
            save_storage(self.templates, self.meta)
        except Exception:
            pass
        # Add created tickets to list_items and switch to Welcome
        created_items = []
        existing_keys = {str(it.get("Issue key") or it.get("Issue id") or "") for it in self.list_items}
        for ticket, issue_key in successes:
            tkey = str(ticket.get("Issue key") or ticket.get("Issue id") or "")
            if tkey not in existing_keys:
                self.list_items.append(ticket)
                created_items.append(ticket)
                existing_keys.add(tkey)
        if created_items:
            self.list_items = _dedup_list_items(self.list_items)
            self.meta["fetched_issues"] = list(self.list_items)
            self.meta["welcome_updates"] = {
                "new": len(created_items),
                "new_ticket_keys": [t.get("Issue key") or t.get("Issue id") for t in created_items]
            }
            save_storage(self.templates, self.meta)
            self._populate_listview()
            self._update_welcome_text()
        self.show_tabs_view()
        self.notebook.select(self._welcome_frame)
        # Warn if any tickets were excluded due to summary mismatch
        self._show_summary_mismatch_excluded(summary_mismatch_excluded)
        # Show dialog with direct links to each uploaded ticket
        jira_base = s._jira_base if s else ""
        self._show_upload_complete_dialog(successes, failures, jira_base)
        debug_log("Upload bundle completed. Summary: " + str(len(successes)) + " successes, " + str(len(failures)) + " failures.")

    def _show_upload_complete_dialog(self, successes, failures, jira_base):
        """Show dialog with direct links to each uploaded ticket."""
        win = tk.Toplevel(self)
        self._register_toplevel(win)
        win.title("Upload complete")
        win.minsize(450, 320)
        win.geometry("580x380")
        win.resizable(True, True)
        msg = f"Successes: {len(successes)}. Failures: {len(failures)}."
        if failures:
            has_perm_error = False
            for f_entry in failures:
                f_str = " ".join(str(x) for x in f_entry)
                if any(k in f_str.lower() for k in ("401", "403", "permission", "not found")):
                    has_perm_error = True
                    break
            if has_perm_error:
                msg += "\n\nPermission issue detected — your API token may not have access."
                msg += "\nTry regenerating at: id.atlassian.com/manage-profile/security/api-tokens"
            msg += "\nSee jira_debug.log for failure details."
        ttk.Label(win, text=msg).pack(anchor="w", padx=8, pady=(8, 4))
        ttk.Label(win, text="Created/updated tickets — double-click to open in Jira (browser), or use buttons below:").pack(anchor="w", padx=8, pady=(4, 0))
        lb_frame = ttk.Frame(win)
        lb_frame.pack(fill="both", expand=True, padx=8, pady=6)
        lb = tk.Listbox(lb_frame, height=12, selectmode="single", font=("Segoe UI", 10))
        lb.pack(side="left", fill="both", expand=True)
        vs = ttk.Scrollbar(lb_frame, orient="vertical", command=lb.yview)
        vs.pack(side="right", fill="y")
        lb.configure(yscrollcommand=vs.set)
        keys = []
        for ticket, issue_key in successes:
            summary = (ticket.get("Summary") or "")[:50]
            lb.insert(tk.END, f"{issue_key} — {summary}")
            keys.append(issue_key)
        if not keys:
            lb.insert(tk.END, "(No tickets created)")
        def open_in_app():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Info", "Select a ticket to open.")
                return
            idx = int(sel[0])
            if 0 <= idx < len(keys):
                self._open_ticket_by_key(keys[idx])
            win.destroy()
        def open_in_jira():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Info", "Select a ticket to open in Jira.")
                return
            idx = int(sel[0])
            if 0 <= idx < len(keys) and jira_base:
                url = f"{jira_base.rstrip('/')}/browse/{keys[idx]}"
                webbrowser.open(url)
        lb.bind("<Double-1>", lambda e: open_in_jira() if jira_base else open_in_app())
        if failures:
            fail_frame = ttk.LabelFrame(win, text="Failures", padding=4)
            fail_frame.pack(fill="both", expand=False, padx=8, pady=(0, 4))
            fail_txt = tk.Text(fail_frame, height=5, wrap="word", font=("Consolas", 9), bg="#1e1e1e", fg="#ff6b6b", insertbackground="#dcdcdc")
            fail_txt.pack(fill="both", expand=True)
            for f_entry in failures:
                summary_str = (f_entry[0].get("Summary") or "?")[:40] if isinstance(f_entry[0], dict) else "?"
                detail = " | ".join(str(x) for x in f_entry[1:])
                fail_txt.insert("end", f"{summary_str}: {detail}\n")
            fail_txt.configure(state="disabled")
            def copy_failures():
                lines = []
                for f_entry in failures:
                    lines.append(" | ".join(str(x) for x in f_entry[1:]))
                win.clipboard_clear()
                win.clipboard_append("\n".join(lines))
                messagebox.showinfo("Copied", "Failure details copied to clipboard.")
            ttk.Button(fail_frame, text="Copy errors to clipboard", command=copy_failures).pack(anchor="e", pady=(4, 0))
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill="x", padx=8, pady=6)
        ttk.Button(btn_frame, text="Open selected in app", command=open_in_app).pack(side="left", padx=(0, 8))
        if jira_base:
            ttk.Button(btn_frame, text="Jira", command=open_in_jira).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side="right")

    # ── Inline-image media resolution ──────────────────────────────────────

    def _resolve_pending_media(self, s, issue_key, adf_dict):
        """Upload pending inline images and patch ADF media nodes with
        real Jira attachment IDs.  Returns the updated ADF dict (mutated
        in place) and a bool indicating whether any images were uploaded."""
        if not isinstance(adf_dict, dict):
            return adf_dict, False

        pending = []  # collect (node_attrs, filepath) tuples

        def _collect(node):
            if not isinstance(node, dict):
                return
            if node.get("type") == "media":
                attrs = node.get("attrs") or {}
                fpath = attrs.get("__pendingPath") or ""
                fname = attrs.get("__fileName") or ""
                if fname or fpath:
                    real_path = fpath if (fpath and os.path.isfile(fpath)) else None
                    pending.append((attrs, real_path))
            for child in node.get("content", []):
                _collect(child)

        _collect(adf_dict)

        if not pending:
            return adf_dict, False

        uploaded = False
        # Deduplicate: upload each unique file only once, reuse the
        # attachment id for all media nodes that reference the same file.
        upload_cache = {}  # fpath -> attachment result dict
        for attrs, fpath in pending:
            if not fpath:
                attrs.pop("__fileName", None)
                attrs.pop("__pendingPath", None)
                continue
            try:
                if fpath in upload_cache:
                    result = upload_cache[fpath]
                else:
                    result = self.upload_attachment(s, issue_key, fpath)
                    upload_cache[fpath] = result
                if result:
                    content_url = result.get("content") or ""
                    att_id = result.get("id", "")
                    # Use 'external' type with the content URL — avoids
                    # needing Jira Media Services IDs and collection names.
                    attrs.clear()
                    attrs["type"] = "external"
                    attrs["url"] = content_url
                    uploaded = True
                    debug_log(f"Uploaded inline image {os.path.basename(fpath)} -> attachment {att_id} url={content_url}")
                else:
                    debug_log(f"Failed to upload inline image: {fpath}")
                    attrs.pop("__fileName", None)
                    attrs.pop("__pendingPath", None)
            except Exception:
                debug_log(f"Exception uploading inline image: {traceback.format_exc()}")
                attrs.pop("__fileName", None)
                attrs.pop("__pendingPath", None)

        return adf_dict, uploaded

    def _remove_invalid_media_nodes(self, node):
        """Remove mediaSingle/mediaGroup nodes whose child media has no valid
        reference (id or url) — Jira rejects these with INVALID_INPUT."""
        if not isinstance(node, dict):
            return
        content = node.get("content")
        if isinstance(content, list):
            cleaned = []
            for child in content:
                if isinstance(child, dict) and child.get("type") in ("mediaSingle", "mediaGroup"):
                    media_children = child.get("content") or []
                    has_valid = any(
                        isinstance(m, dict) and m.get("type") == "media"
                        and (
                            (m.get("attrs") or {}).get("id")
                            or (m.get("attrs") or {}).get("url")
                        )
                        for m in media_children
                    )
                    if not has_valid:
                        debug_log(f"Stripped invalid mediaSingle (no attachment id or url)")
                        continue
                cleaned.append(child)
                self._remove_invalid_media_nodes(child)
            node["content"] = cleaned

    def _adf_has_pending_media(self, node):
        """Return True if the ADF tree has any media nodes with __pendingPath."""
        if isinstance(node, dict):
            if node.get("type") == "media":
                attrs = node.get("attrs") or {}
                if attrs.get("__pendingPath") or attrs.get("__fileName"):
                    return True
            for child in node.get("content", []):
                if self._adf_has_pending_media(child):
                    return True
        elif isinstance(node, list):
            for item in node:
                if self._adf_has_pending_media(item):
                    return True
        return False

    def _strip_custom_media_attrs(self, node):
        """Remove __fileName and __pendingPath from all media nodes in an ADF
        tree so Jira doesn't reject unknown attributes."""
        if isinstance(node, dict):
            if node.get("type") == "media":
                attrs = node.get("attrs") or {}
                attrs.pop("__fileName", None)
                attrs.pop("__pendingPath", None)
            for child in node.get("content", []):
                self._strip_custom_media_attrs(child)
        elif isinstance(node, list):
            for item in node:
                self._strip_custom_media_attrs(item)
