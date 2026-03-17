"""
Jira API methods mixin for AvalancheApp.
"""
import re
import json
import copy
import uuid
import traceback
import webbrowser
import tkinter as tk
from tkinter import messagebox

try:
    import requests
except Exception:
    requests = None

from config import FETCH_FIELDS, _JIRA_FIELDS_FALLBACK, DEBUG_LOG
from storage import perform_jira_request, save_storage
from utils import debug_log


class JiraAPIMixin:
    """Jira API methods mixed into AvalancheApp."""

    def get_jira_session(self):
        if requests is None:
            messagebox.showerror("Missing dependency", "Install requests: py -m pip install requests")
            return None
        jira = self.meta.get("jira", {})
        base = (jira.get("base") or "").strip()
        email = (jira.get("email") or "").strip()
        token = (jira.get("token") or "").strip()
        if not base or not email or not token:
            messagebox.showinfo("No credentials", "Jira credentials not set. Use 'Set Jira API...' to configure.")
            return None
        s = requests.Session()
        s.auth = (email, token)
        s.headers.update({"Accept": "application/json"})
        s._jira_base = base.rstrip("/")
        debug_log(f"Created Jira session for base={s._jira_base}, user={email}")
        return s

    def _open_in_jira_browser(self, issue_key):
        """Open issue in default browser. Returns True if opened, False if no base/key."""
        key = str(issue_key or "").strip()
        if not key or key.startswith("LOCAL-"):
            messagebox.showinfo("Info", "No Jira issue key to open.")
            return False
        base = (self.meta.get("jira", {}).get("base") or "").strip().rstrip("/")
        if not base:
            messagebox.showinfo("Info", "Set Jira API credentials first.")
            return False
        url = f"{base}/browse/{key}"
        webbrowser.open(url)
        return True

    def test_jira_connection(self):
        s = self.get_jira_session()
        if not s:
            return
        auth_ok = False
        display = ""
        for ver in ("3", "2"):
            url = f"{s._jira_base}/rest/api/{ver}/myself"
            try:
                r = perform_jira_request(s, "GET", url, timeout=15)
            except Exception as e:
                messagebox.showerror("Connection failed", f"Request to {url} failed: {e}\n\nSee {DEBUG_LOG}")
                return
            if r.status_code == 200:
                try:
                    info = r.json()
                    display = info.get("displayName") or info.get("emailAddress") or "(OK)"
                    self.meta["jira_current_user"] = {
                        "displayName": (info.get("displayName") or "").strip(),
                        "emailAddress": (info.get("emailAddress") or "").strip()
                    }
                except Exception:
                    display = "(OK - non-JSON)"
                auth_ok = True
                break
            else:
                messagebox.showwarning("Connection check", f"Status {r.status_code} from {url}\n\nResponse (truncated):\n{(r.text or '')[:2000]}\n\nSee {DEBUG_LOG}")
        if not auth_ok:
            messagebox.showerror("Connection failed", "All tests failed. Check credentials and base URL.")
            return
        diag_lines = [f"Authenticated as: {display}"]
        try:
            proj_r = perform_jira_request(s, "GET", f"{s._jira_base}/rest/api/3/project/search", params={"maxResults": 5}, timeout=15)
            if proj_r.status_code == 200:
                projects = [p.get("key", "?") for p in (proj_r.json().get("values") or proj_r.json().get("projects") or [])[:5]]
                if projects:
                    diag_lines.append(f"Visible projects: {', '.join(projects)}")
                else:
                    diag_lines.append("WARNING: No projects visible to this token!")
            else:
                diag_lines.append(f"Project check: status {proj_r.status_code}")
        except Exception:
            diag_lines.append("Project check: failed")
        try:
            jql_r = perform_jira_request(s, "POST", f"{s._jira_base}/rest/api/3/search/jql",
                                         json_body={"jql": "assignee = currentUser() ORDER BY created DESC", "maxResults": 1}, timeout=15)
            if jql_r.status_code == 200:
                total = len(jql_r.json().get("issues") or [])
                diag_lines.append(f"Your assigned issues: {'found' if total else 'NONE visible'}")
            else:
                diag_lines.append(f"Issue search: status {jql_r.status_code}")
        except Exception:
            diag_lines.append("Issue search: failed")
        sample_key = ""
        for it in self.list_items[:5]:
            k = str(it.get("Issue key") or "").strip()
            if k and not k.startswith("LOCAL-"):
                sample_key = k
                break
        if sample_key:
            try:
                test_r = perform_jira_request(s, "GET", f"{s._jira_base}/rest/api/3/issue/{sample_key}",
                                              params={"fields": "summary"}, timeout=15)
                if test_r.status_code == 200:
                    diag_lines.append(f"Issue access ({sample_key}): OK")
                else:
                    diag_lines.append(f"Issue access ({sample_key}): {test_r.status_code} — token may lack permissions")
            except Exception:
                diag_lines.append(f"Issue access ({sample_key}): failed")
        messagebox.showinfo("Connection Test", "\n".join(diag_lines))

    def _resolve_assignee(self, session, value, project_key=None):
        if not value:
            return None
        val = str(value).strip()
        if len(val) >= 20 and all(c.isalnum() or c in "-_:" for c in val):
            return val
        cached = self._user_cache.get(val) or self.meta.get("user_cache", {}).get(val)
        if cached:
            return cached
        try:
            url = f"{session._jira_base}/rest/api/3/user/search"
            params = {"query": val}
            resp = perform_jira_request(session, "GET", url, params=params, timeout=20)
            if resp.status_code == 200:
                arr = resp.json()
                if isinstance(arr, list) and arr:
                    chosen = None
                    for u in arr:
                        if u.get("emailAddress") == val:
                            chosen = u
                            break
                    if not chosen:
                        for u in arr:
                            if u.get("displayName") == val:
                                chosen = u
                                break
                    if not chosen:
                        chosen = arr[0]
                    accountId = chosen.get("accountId") or chosen.get("key") or None
                    if accountId:
                        self._user_cache[val] = accountId
                        self.meta.setdefault("user_cache", {})[val] = accountId
                        save_storage(self.templates, self.meta)
                        return accountId
        except Exception:
            debug_log("User search failed: " + traceback.format_exc())
        return None

    def _fetch_assignable_users(self, session, project_key=None):
        """
        Fetch assignable users from Jira. Returns list of display names for the Assignee dropdown.
        Uses /rest/api/3/user/assignable/search when project_key is set, else /rest/api/3/users/search.
        """
        names = []
        seen = set()
        try:
            if project_key and str(project_key).strip():
                url = f"{session._jira_base}/rest/api/3/user/assignable/search"
                params = {"project": str(project_key).strip(), "maxResults": 1000}
                resp = perform_jira_request(session, "GET", url, params=params, timeout=30)
            else:
                url = f"{session._jira_base}/rest/api/3/users/search"
                params = {"maxResults": 1000}
                resp = perform_jira_request(session, "GET", url, params=params, timeout=30)
            if resp.status_code != 200:
                return []
            arr = resp.json()
            if not isinstance(arr, list):
                return []
            for u in arr:
                dn = (u.get("displayName") or "").strip()
                em = (u.get("emailAddress") or "").strip()
                key = dn or em or (u.get("accountId") or "")
                if not key:
                    continue
                label = dn or em or key
                if label not in seen:
                    seen.add(label)
                    names.append(label)
            names.sort(key=lambda x: x.lower())
        except Exception:
            debug_log("Fetch assignable users failed: " + traceback.format_exc())
        return names

    def _fetch_projects(self, session):
        """Fetch project keys from Jira."""
        names = []
        try:
            url = f"{session._jira_base}/rest/api/3/project/search"
            params = {"maxResults": 200}
            resp = perform_jira_request(session, "GET", url, params=params, timeout=30)
            if resp.status_code != 200:
                return []
            data = resp.json()
            values = data.get("values") if isinstance(data, dict) else (data if isinstance(data, list) else [])
            for p in values:
                if isinstance(p, dict):
                    key = (p.get("key") or "").strip()
                    if key:
                        names.append(key)
            names.sort(key=lambda x: x.upper())
        except Exception:
            debug_log("Fetch projects failed: " + traceback.format_exc())
        return names

    def _fetch_issue_types(self, session, project_key=None):
        """Fetch issue type names from Jira."""
        names = []
        try:
            resp = None
            if project_key:
                try:
                    proj_url = f"{session._jira_base}/rest/api/3/project/{project_key}"
                    proj_resp = perform_jira_request(session, "GET", proj_url, timeout=10)
                    if proj_resp.status_code == 200:
                        proj = proj_resp.json()
                        pid = proj.get("id")
                        if pid:
                            url = f"{session._jira_base}/rest/api/3/issuetype/project"
                            params = {"projectId": pid}
                            resp = perform_jira_request(session, "GET", url, params=params, timeout=20)
                except Exception:
                    pass
            if resp is None or resp.status_code != 200:
                url = f"{session._jira_base}/rest/api/3/issuetype"
                resp = perform_jira_request(session, "GET", url, timeout=20)
            if resp.status_code != 200:
                return []
            arr = resp.json()
            if not isinstance(arr, list):
                return []
            for it in arr:
                if isinstance(it, dict):
                    n = (it.get("name") or "").strip()
                    if n:
                        names.append(n)
            names.sort(key=lambda x: x.lower())
        except Exception:
            debug_log("Fetch issue types failed: " + traceback.format_exc())
        return names

    def _fetch_statuses(self, session, project_key=None):
        """Fetch status names for a project from Jira."""
        names = []
        if not project_key:
            project_key = "SUNDANCE"
        try:
            url = f"{session._jira_base}/rest/api/3/project/{project_key}/statuses"
            resp = perform_jira_request(session, "GET", url, timeout=20)
            if resp.status_code != 200:
                return []
            arr = resp.json()
            if not isinstance(arr, list):
                return []
            seen = set()
            for wf in arr:
                if isinstance(wf, dict):
                    for s in wf.get("statuses") or []:
                        if isinstance(s, dict):
                            n = (s.get("name") or "").strip()
                            if n and n not in seen:
                                seen.add(n)
                                names.append(n)
            names.sort(key=lambda x: x.lower())
        except Exception:
            debug_log("Fetch statuses failed: " + traceback.format_exc())
        return names

    def _fetch_priorities(self, session):
        """Fetch priority names from Jira."""
        names = []
        try:
            url = f"{session._jira_base}/rest/api/3/priority"
            resp = perform_jira_request(session, "GET", url, timeout=20)
            if resp.status_code != 200:
                return []
            arr = resp.json()
            if not isinstance(arr, list):
                return []
            for p in arr:
                if isinstance(p, dict):
                    n = (p.get("name") or "").strip()
                    if n:
                        names.append(n)
            names.sort(key=lambda x: x.lower())
        except Exception:
            debug_log("Fetch priorities failed: " + traceback.format_exc())
        return names

    def _fetch_components(self, session, project_key=None):
        """Fetch component names for a project from Jira."""
        names = []
        if not project_key:
            project_key = "SUNDANCE"
        try:
            url = f"{session._jira_base}/rest/api/3/project/{project_key}/components"
            resp = perform_jira_request(session, "GET", url, timeout=20)
            if resp.status_code != 200:
                return []
            arr = resp.json()
            if not isinstance(arr, list):
                return []
            for c in arr:
                if isinstance(c, dict):
                    n = (c.get("name") or "").strip()
                    if n:
                        names.append(n)
            names.sort(key=lambda x: x.lower())
        except Exception:
            debug_log("Fetch components failed: " + traceback.format_exc())
        return names

    def _fetch_labels(self, session):
        """Fetch ALL label values from Jira (fully paginated, merged with existing cache)."""
        existing = set(self.meta.get("options", {}).get("Labels", []))
        new_vals = set()
        try:
            url = f"{session._jira_base}/rest/api/3/label"
            start_at  = 0
            page_size = 1000   # max page the API accepts
            while True:
                params = {"startAt": start_at, "maxResults": page_size}
                resp = perform_jira_request(session, "GET", url, params=params, timeout=30)
                if resp.status_code != 200:
                    break
                data = resp.json()
                vals = data.get("values") if isinstance(data, dict) else []
                for v in (vals or []):
                    if isinstance(v, str) and v.strip():
                        new_vals.add(v.strip())
                if not vals or data.get("isLast", True):
                    break
                start_at += page_size
        except Exception:
            debug_log("Fetch labels failed: " + traceback.format_exc())
        # Merge new with existing so previously cached labels aren't lost
        merged = sorted(existing | new_vals, key=lambda x: x.lower())
        return merged

    def _fetch_sprints(self, session, project_key=None):
        """Fetch sprint names for a project's boards from Jira Agile API."""
        if not project_key:
            project_key = "SUNDANCE"
        names = set()
        try:
            board_url = f"{session._jira_base}/rest/agile/1.0/board"
            resp = perform_jira_request(session, "GET", board_url,
                                        params={"projectKeyOrId": project_key, "maxResults": 50},
                                        timeout=20)
            boards = []
            if resp.status_code == 200:
                boards = (resp.json() or {}).get("values", [])
            for board in boards:
                bid = board.get("id")
                if not bid:
                    continue
                sprint_url = f"{session._jira_base}/rest/agile/1.0/board/{bid}/sprint"
                sr = perform_jira_request(session, "GET", sprint_url,
                                          params={"maxResults": 100, "state": "active,future"},
                                          timeout=20)
                if sr.status_code != 200:
                    continue
                for sp in (sr.json() or {}).get("values", []):
                    n = (sp.get("name") or "").strip()
                    if n:
                        names.add(n)
        except Exception:
            debug_log("Fetch sprints failed: " + traceback.format_exc())
        return sorted(names, key=lambda x: x.lower())

    def _fetch_versions(self, session, project_key=None):
        """Fetch fix version / milestone names for a project."""
        if not project_key:
            project_key = "SUNDANCE"
        names = []
        try:
            url = f"{session._jira_base}/rest/api/3/project/{project_key}/versions"
            resp = perform_jira_request(session, "GET", url, timeout=20)
            if resp.status_code == 200:
                for v in (resp.json() if isinstance(resp.json(), list) else []):
                    n = (v.get("name") or "").strip()
                    if n:
                        names.append(n)
        except Exception:
            debug_log("Fetch versions failed: " + traceback.format_exc())
        return sorted(names, key=lambda x: x.lower())

    def _fetch_jira_fields(self, session):
        """Fetch all Jira fields (id, name) for changelog/ignored-fields. Returns [(id, name), ...]."""
        try:
            url = f"{session._jira_base}/rest/api/3/field"
            resp = perform_jira_request(session, "GET", url, timeout=20)
            if resp.status_code != 200:
                return _JIRA_FIELDS_FALLBACK
            data = resp.json()
            out = []
            for f in (data if isinstance(data, list) else []):
                fid = (f.get("id") or "").strip()
                name = (f.get("name") or fid).strip()
                if fid:
                    out.append((fid, name))
            return sorted(out, key=lambda x: (x[1].lower(), x[0])) if out else _JIRA_FIELDS_FALLBACK
        except Exception:
            debug_log("Fetch Jira fields failed: " + traceback.format_exc())
            return _JIRA_FIELDS_FALLBACK

    def jira_search_jql_simple(self, session, jql, max_results=50, start_at=0,
                               exclude_keys=None, fields=None):
        # Use /rest/api/3/search/jql (POST /rest/api/3/search returns 410 Gone as of 2025)
        url = f"{session._jira_base}/rest/api/3/search/jql"
        body = {"jql": jql}
        if max_results:
            body["maxResults"] = max_results
        if fields:
            body["fields"] = list(fields)
        # /search/jql uses nextPageToken for pagination (cursor-based); omit for first page
        if exclude_keys:
            keys = [k for k in exclude_keys if k and str(k).strip()]
            if keys and len(keys) <= 100:
                try:
                    quoted = ", ".join(f'"{k}"' for k in keys)
                    not_in = f" AND key NOT IN ({quoted})"
                    if " ORDER BY " in jql.upper():
                        base, _, order = jql.upper().partition(" ORDER BY ")
                        body["jql"] = jql[:len(base)] + not_in + " " + jql[jql.upper().find(" ORDER BY "):]
                    else:
                        body["jql"] = jql + not_in
                except Exception:
                    pass
        resp = perform_jira_request(session, "POST", url, json_body=body, timeout=60)
        resp.raise_for_status()
        return resp.json()

    def _jira_attachments_to_field(self, attachment_list):
        """Convert Jira attachment array to Attachment field value. Returns JSON string for Jira attachments, else empty."""
        if not attachment_list or not isinstance(attachment_list, list):
            return ""
        items = []
        for a in attachment_list:
            if isinstance(a, dict):
                fn = a.get("filename") or a.get("name") or ""
                url = a.get("content") or ""
                size = a.get("size", 0)
                mime = a.get("mimeType") or ""
                thumb = a.get("thumbnail") or ""
                if fn or url:
                    items.append({
                        "filename": fn,
                        "content": url,
                        "size": size,
                        "mimeType": mime,
                        "thumbnail": thumb,
                    })
        if not items:
            return ""
        return json.dumps(items, ensure_ascii=False)

    def fetch_issue_details(self, session, issue_id_or_key, fields=None):
        if fields is None:
            fields_param = ",".join(FETCH_FIELDS)
        else:
            fields_param = ",".join(fields) if isinstance(fields, (list, tuple)) else str(fields)
        url = f"{session._jira_base}/rest/api/3/issue/{issue_id_or_key}"
        # Request Jira's server-rendered HTML via expand=renderedFields
        resp = perform_jira_request(session, "GET", url, params={"fields": fields_param, "expand": "renderedFields"}, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def upload_attachment(self, session, issue_key, filepath):
        """Upload a single file as an attachment and return the Jira
        attachment metadata dict (with 'id', 'filename', 'content', etc.).
        Returns None on failure."""
        import os
        url = f"{session._jira_base}/rest/api/3/issue/{issue_key}/attachments"
        try:
            with open(filepath, "rb") as fh:
                resp = session.post(
                    url,
                    files={"file": (os.path.basename(filepath), fh)},
                    headers={"X-Atlassian-Token": "no-check"},
                )
            if resp.status_code in (200, 201):
                data = resp.json()
                if isinstance(data, list) and data:
                    return data[0]
                return data
        except Exception:
            pass
        return None

    def _sanitize_adf_for_upload(self, node):
        """Fix ADF issues that cause Jira API to reject or silently drop nodes.
        - taskList must have attrs.localId
        - taskItem must have attrs.localId/state and contain inline nodes (not paragraphs)
        - media nodes without a valid id are removed (their parent mediaSingle is dropped)
        Returns a cleaned copy."""
        if isinstance(node, dict):
            result = {}
            for k, v in node.items():
                result[k] = self._sanitize_adf_for_upload(v)
            ntype = result.get("type")
            # media: strip custom internal attrs and ensure required fields
            if ntype == "media":
                attrs = result.get("attrs") or {}
                attrs.pop("__fileName", None)
                attrs.pop("__pendingPath", None)
                # Jira requires 'collection' on media nodes with an id
                if attrs.get("id") and "collection" not in attrs:
                    attrs["collection"] = ""
                result["attrs"] = attrs
            # mediaSingle/mediaGroup: drop entirely if no child has a valid
            # media reference (either id for 'file' type or url for 'external')
            if ntype in ("mediaSingle", "mediaGroup"):
                children = result.get("content") or []
                has_valid = any(
                    isinstance(m, dict) and m.get("type") == "media"
                    and (
                        (m.get("attrs") or {}).get("id")
                        or (m.get("attrs") or {}).get("url")
                    )
                    for m in children
                )
                if not has_valid:
                    return None
            # Filter out None children (removed nodes) from content lists
            if "content" in result and isinstance(result["content"], list):
                result["content"] = [c for c in result["content"] if c is not None]
            # taskList: ensure attrs.localId exists
            if ntype == "taskList":
                attrs = result.setdefault("attrs", {})
                if not attrs.get("localId"):
                    attrs["localId"] = str(uuid.uuid4())
            # taskItem: ensure attrs, unwrap paragraph wrappers → inline content
            if ntype == "taskItem":
                attrs = result.setdefault("attrs", {})
                if not attrs.get("localId"):
                    attrs["localId"] = str(uuid.uuid4())
                if not attrs.get("state"):
                    attrs["state"] = "TODO"
                if "content" in result:
                    new_content = []
                    for child in (result["content"] or []):
                        if isinstance(child, dict) and child.get("type") == "paragraph":
                            for inline in (child.get("content") or []):
                                new_content.append(inline)
                        else:
                            new_content.append(child)
                    result["content"] = new_content
            return result
        if isinstance(node, list):
            return [r for r in (self._sanitize_adf_for_upload(x) for x in node) if r is not None]
        return node

    # ── Comment helpers ───────────────────────────────────────────────────────

    def _parse_jira_comments(self, comment_field) -> str:
        """Convert the Jira REST 'comment' field to a JSON string list.

        Jira returns:  {"comments": [...], "total": N, ...}
        Each comment:  {"id": "...", "author": {...}, "created": "...", "body": <ADF or str>}
        Returns a JSON-encoded list ready to store in the ticket's Comment field.
        """
        if not comment_field:
            return "[]"
        raw_list = []
        if isinstance(comment_field, dict):
            raw_list = comment_field.get("comments") or []
        elif isinstance(comment_field, list):
            raw_list = comment_field
        parsed = []
        for c in raw_list:
            author_obj = c.get("author") or {}
            author = (author_obj.get("displayName") or
                      author_obj.get("emailAddress") or "Unknown")
            body_node = c.get("body")
            if isinstance(body_node, dict):
                try:
                    body = self._extract_text_from_adf(body_node)
                except Exception:
                    body = str(body_node)
            else:
                body = str(body_node or "")
            parsed.append({
                "id":     c.get("id", ""),
                "author": author,
                "date":   c.get("created", ""),
                "body":   body,
                "posted": True,
            })
        try:
            return json.dumps(parsed, ensure_ascii=False)
        except Exception:
            return "[]"

    def _detect_epic_link_field(self, fields: dict) -> tuple:
        """Return (mode, epic_key) for the epic relationship on this ticket.

        Jira has two different mechanisms depending on project type:
        - Classic (company-managed): customfield_10014 holds the epic key as a string.
        - Next-gen (team-managed): the 'parent' field holds the epic object when
          the parent's issue type is "Epic".

        Returns:
            ("classic", key)   — classic epic link
            ("nextgen", key)   — next-gen parent-is-epic
            (None, "")         — no epic relationship found
        """
        # Classic Jira: customfield_10014 is the Epic Link field (plain key string)
        cf_14 = fields.get("customfield_10014")
        if cf_14 and isinstance(cf_14, str) and cf_14.strip():
            return ("classic", cf_14.strip())

        # Next-gen Jira: parent field exists and its issue type is "Epic"
        parent_obj = fields.get("parent") or {}
        if parent_obj.get("key"):
            parent_type = (
                (parent_obj.get("fields") or {})
                .get("issuetype", {})
                .get("name", "")
            )
            if parent_type.lower() == "epic":
                return ("nextgen", parent_obj["key"])

        return (None, "")

    def _parse_jira_issue_links(self, links_field) -> str:
        """Convert the Jira REST 'issuelinks' field to a JSON string list.

        Each Jira issue link has an inwardIssue or outwardIssue and a type with
        inward/outward labels (e.g. "blocks" / "is blocked by").

        Returns a JSON-encoded list:
          [{id, type_name, direction, direction_label, key, summary, status, posted}]
        where posted=True for links already in Jira, False for locally added ones.
        """
        if not links_field:
            return "[]"
        raw_list = links_field if isinstance(links_field, list) else []
        parsed = []
        for lnk in raw_list:
            try:
                ltype = lnk.get("type") or {}
                if lnk.get("outwardIssue"):
                    direction      = "outward"
                    direction_label = ltype.get("outward", "")
                    issue_obj       = lnk["outwardIssue"]
                else:
                    direction      = "inward"
                    direction_label = ltype.get("inward", "")
                    issue_obj       = lnk.get("inwardIssue") or {}
                parsed.append({
                    "id":              lnk.get("id", ""),
                    "type_name":       ltype.get("name", ""),
                    "direction":       direction,
                    "direction_label": direction_label,
                    "key":             issue_obj.get("key", ""),
                    "summary":         (issue_obj.get("fields") or {}).get("summary", ""),
                    "status":          ((issue_obj.get("fields") or {})
                                        .get("status", {}).get("name", "")),
                    "posted":          True,
                })
            except Exception:
                continue
        try:
            return json.dumps(parsed, ensure_ascii=False)
        except Exception:
            return "[]"

    def _text_to_adf(self, text: str):
        if text is None:
            return None
        text = str(text)
        if text.endswith("\n"):
            text = text.rstrip("\n")
        if text == "":
            return {"type": "doc", "version": 1, "content": []}
        lines = text.split("\n")
        content = []
        for line in lines:
            if line == "":
                content.append({"type": "paragraph", "content": []})
            else:
                content.append({"type": "paragraph", "content": [{"type": "text", "text": line}]})
        return {"type": "doc", "version": 1, "content": content}

    def _replace_in_adf(self, node, replacements, parent_content_info=None, insert_after=None):
        """Recursively replace '!' in ADF text nodes. Extra lines (more than !) are inserted right after last !.
        parent_content_info = (content_list, index) for current node; insert_after receives (list, index) on replace.
        Multi-line replacements (containing \\n) are split into separate paragraphs in the nearest block container."""
        if not replacements:
            return
        if isinstance(node, dict):
            if node.get("type") == "text" and "text" in node:
                text = node["text"]
                if "!" in text:
                    while "!" in text and replacements:
                        text = text.replace("!", replacements.pop(0), 1)
                    # If the replaced text has newlines, split into first line (this node)
                    # and extra paragraphs inserted into the parent block container.
                    if "\n" in text and parent_content_info is not None:
                        lines = text.split("\n")
                        node["text"] = lines[0]
                        extra_paras = []
                        for line in lines[1:]:
                            if line.strip():
                                extra_paras.append({"type": "paragraph", "content": [{"type": "text", "text": line}]})
                            else:
                                extra_paras.append({"type": "paragraph", "content": []})
                        if extra_paras:
                            block_list, block_idx = parent_content_info
                            if block_list is not None and block_idx is not None:
                                for j, p in enumerate(extra_paras):
                                    block_list.insert(block_idx + 1 + j, p)
                    else:
                        node["text"] = text
                    if insert_after is not None and parent_content_info is not None:
                        insert_after[0], insert_after[1] = parent_content_info
            # Process "content" first (document order)
            if "content" in node:
                # ADF: paragraph, heading, codeBlock have INLINE-only content; we must insert blocks
                # after them (into their parent), not inside. tableCell, listItem, etc. have block content.
                INLINE_ONLY_TYPES = ("paragraph", "heading", "codeBlock")
                for i, c in enumerate(node["content"]):
                    next_parent = parent_content_info if node.get("type") in INLINE_ONLY_TYPES else (node["content"], i)
                    self._replace_in_adf(c, replacements, parent_content_info=next_parent, insert_after=insert_after)
                    if not replacements:
                        return
            for k, v in node.items():
                if k == "content":
                    continue
                if isinstance(v, list):
                    for i, c in enumerate(v):
                        self._replace_in_adf(c, replacements, parent_content_info=(v, i), insert_after=insert_after)
                        if not replacements:
                            return
                elif isinstance(v, dict):
                    self._replace_in_adf(v, replacements, parent_content_info=parent_content_info, insert_after=insert_after)
                    if not replacements:
                        return
        elif isinstance(node, list):
            for i, c in enumerate(node):
                self._replace_in_adf(c, replacements, parent_content_info=(node, i), insert_after=insert_after)
                if not replacements:
                    return

    def _adf_contains_exclamation(self, node):
        """Return True if ADF contains '!' in any text node (doc, table, tableRow, tableCell, paragraph, text)."""
        if isinstance(node, dict):
            if node.get("type") == "text" and "!" in (node.get("text") or ""):
                return True
            if "content" in node and self._adf_contains_exclamation(node["content"]):
                return True
            for k, v in node.items():
                if k != "content" and isinstance(v, (list, dict)) and self._adf_contains_exclamation(v):
                    return True
        elif isinstance(node, list):
            return any(self._adf_contains_exclamation(c) for c in node)
        return False

    def _count_exclamations_in_adf(self, node):
        """Count total '!' occurrences across all text nodes in ADF."""
        count = 0
        if isinstance(node, dict):
            if node.get("type") == "text":
                count += (node.get("text") or "").count("!")
            if "content" in node:
                count += self._count_exclamations_in_adf(node["content"])
            for k, v in node.items():
                if k != "content" and isinstance(v, (list, dict)):
                    count += self._count_exclamations_in_adf(v)
        elif isinstance(node, list):
            for c in node:
                count += self._count_exclamations_in_adf(c)
        return count

    def _build_adf_list(self, items, list_type="bulletList"):
        """Build an ADF bulletList or orderedList node from a list of strings."""
        list_items = []
        for item in items:
            if item.strip():
                list_items.append({
                    "type": "listItem",
                    "content": [{"type": "paragraph", "content": [{"type": "text", "text": item}]}]
                })
        if not list_items:
            return None
        return {"type": list_type, "content": list_items}

    def _build_adf_smart_list(self, items):
        """Build an ADF taskList node from lines (same as the Smart list button)."""
        import uuid as _uuid
        task_items = []
        for item in items:
            text = item.strip()
            if not text:
                continue
            task_items.append({
                "type": "taskItem",
                "attrs": {"localId": str(_uuid.uuid4()), "state": "TODO"},
                "content": [{"type": "text", "text": text}]
            })
        if not task_items:
            return []
        return [{
            "type": "taskList",
            "attrs": {"localId": str(_uuid.uuid4())},
            "content": task_items
        }]

    def _extract_text_from_adf(self, node):
        out = []
        def walk(n):
            if n is None:
                return
            if isinstance(n, str):
                out.append(n)
                return
            if isinstance(n, dict):
                t = n.get("type")
                if t == "text" and "text" in n:
                    out.append(n.get("text", ""))
                elif t == "mention":
                    out.append((n.get("attrs") or {}).get("text", ""))
                elif t == "emoji":
                    out.append((n.get("attrs") or {}).get("text", "") or (n.get("attrs") or {}).get("shortName", ""))
                elif t == "status":
                    out.append((n.get("attrs") or {}).get("text", ""))
                elif t == "date":
                    out.append((n.get("attrs") or {}).get("timestamp", ""))
                elif t == "hardBreak":
                    out.append("\n")
                else:
                    for k in ("content",):
                        if k in n:
                            for c in n[k]:
                                walk(c)
            elif isinstance(n, list):
                for c in n:
                    walk(c)
        walk(node)
        return " ".join([s for s in (o.strip() for o in out) if s])

    def _recover_adf_for_ticket(self, ticket):
        """Try to recover valid ADF for a ticket from open tabs, stored list items, or rendered HTML."""
        issue_key = str(ticket.get("Issue key") or "").strip()
        issue_id = str(ticket.get("Issue id") or "").strip()
        # 1. Check open tabs — the JSON editor has the most up-to-date ADF
        for tab_id, tf in self.tabs.items():
            try:
                tab_data = tf.read_to_dict()
                tab_key = str(tab_data.get("Issue key") or "").strip()
                tab_id_val = str(tab_data.get("Issue id") or "").strip()
                if (issue_key and tab_key == issue_key) or (issue_id and tab_id_val == issue_id):
                    adf = tab_data.get("Description ADF")
                    if isinstance(adf, dict) and adf.get("content"):
                        return adf
            except Exception:
                continue
        # 2. Check stored list items
        for it in self.list_items:
            it_key = str(it.get("Issue key") or "").strip()
            it_id = str(it.get("Issue id") or "").strip()
            if (issue_key and it_key == issue_key) or (issue_id and it_id == issue_id):
                adf = it.get("Description ADF")
                if isinstance(adf, dict) and adf.get("content"):
                    return adf
                if isinstance(adf, str) and adf.strip():
                    try:
                        parsed = json.loads(adf)
                        if isinstance(parsed, dict) and parsed.get("content"):
                            return parsed
                    except Exception:
                        pass
                rendered = it.get("Description Rendered") or ""
                if rendered and ("<table" in rendered.lower() or "<ul" in rendered.lower() or "<ol" in rendered.lower()):
                    try:
                        tf = next(iter(self.tabs.values()), None)
                        if tf and hasattr(tf, "_convert_html_to_adf"):
                            return tf._convert_html_to_adf(rendered)
                    except Exception:
                        pass
                break
        return None

    def _recover_template_adf(self, template_name):
        """If a template has empty ADF (saved while JSON editor was broken), try to
        recover the ADF from the fetched issue list or by converting HTML to ADF."""
        tpl = self.templates.get(template_name, {})
        # Check if any fetched issue matches by Issue key/id and has ADF
        tpl_key = tpl.get("Issue key") or tpl.get("Issue id")
        if tpl_key:
            for it in self.list_items:
                it_key = it.get("Issue key") or it.get("Issue id")
                if it_key == tpl_key:
                    adf = it.get("Description ADF")
                    if isinstance(adf, dict) and adf.get("content"):
                        return adf
                    if isinstance(adf, str) and adf.strip():
                        try:
                            parsed = json.loads(adf)
                            if isinstance(parsed, dict) and parsed.get("content"):
                                return parsed
                        except Exception:
                            pass
                    # Try converting rendered HTML from fetched issue
                    rendered = it.get("Description Rendered") or ""
                    if rendered and ("<table" in rendered.lower()):
                        try:
                            tf = next(iter(self.tabs.values()), None)
                            if tf and hasattr(tf, "_convert_html_to_adf"):
                                return tf._convert_html_to_adf(rendered)
                        except Exception:
                            pass
                    break
        return None
