"""
Storage and Jira request helpers.
"""
import os
import json
import traceback
import time

from config import TEMPLATES_FILE, DEBUG_LOG, HEADERS
from utils import debug_log

try:
    import requests
except Exception:
    requests = None


def perform_jira_request(session, method, url, params=None, json_body=None, extra_headers=None, timeout=30):
    try:
        debug_log("=== JIRA REQUEST BEGIN ===")
        debug_log(f"Method: {method} URL: {url}")
        if params:
            try:
                debug_log(f"Params: {json.dumps(params, ensure_ascii=False)}")
            except Exception:
                debug_log(f"Params: {repr(params)}")
        if json_body:
            try:
                debug_log(f"JSON body (truncated): {json.dumps(json_body, ensure_ascii=False)[:8000]}")
            except Exception:
                debug_log("JSON body present (unserializable)")
        headers = dict(getattr(session, "headers", {}) or {})
        if extra_headers:
            headers.update(extra_headers)
            debug_log("Extra headers applied.")
        resp = session.request(method, url, params=params, json=json_body, headers=headers, timeout=timeout)
        debug_log(f"Response status: {getattr(resp, 'status_code', 'N/A')}")
        try:
            debug_log("Response text (truncated):")
            debug_log((getattr(resp, "text", "") or "")[:16000])
        except Exception:
            debug_log("Response body unable to read")
        debug_log("=== JIRA REQUEST END ===\n")
        return resp
    except Exception:
        debug_log("=== JIRA REQUEST EXCEPTION ===")
        debug_log(traceback.format_exc())
        debug_log("=== END EXCEPTION ===\n")
        raise


def load_storage():
    if os.path.exists(TEMPLATES_FILE):
        try:
            with open(TEMPLATES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and ("templates" in data or "meta" in data):
                templates = data.get("templates", {})
                meta = data.get("meta", {})
            elif isinstance(data, dict):
                templates = data
                meta = {"options": {}}
            else:
                templates = {}
                meta = {"options": {}}
        except Exception:
            templates = {}
            meta = {"options": {}}
    else:
        templates = {
            "Default Task": {
                "Summary": "[TASK] Short description",
                "Issue Type": "Task",
                "Priority": "Medium",
                "Assignee": "",
                "Labels": "",
                "Description": ""
            }
        }
        meta = {"options": {}}
    meta.setdefault("options", {})
    for h in HEADERS:
        meta["options"].setdefault(h, [])
    meta.setdefault("jira", {})
    meta.setdefault("fetched_issues", [])
    meta.setdefault("user_cache", {})
    meta.setdefault("folders", [])
    meta.setdefault("ticket_folders", {})
    meta.setdefault("internal_priorities", {})
    meta.setdefault("internal_priority_levels", ["High", "Medium", "Low", "None"])
    meta.setdefault("reminder_config", {
        "High": {"type": "daily"},
        "Medium": {"type": "weekly"},
        "Low": {"type": "on_open"},
        "None": {"type": "never"}
    })
    meta.setdefault("last_reminder", {})
    meta.setdefault("first_run_done", False)
    meta.setdefault("tutorial_enabled", True)  # Show tutorial on first startup
    meta.setdefault("welcome_updates", {})
    meta.setdefault("open_ticket_keys", [])
    meta.setdefault("welcome_show_high_priority", True)
    meta.setdefault("stale_ticket_enabled", False)
    meta.setdefault("stale_ticket_days", 14)
    meta.setdefault("stale_ticket_ignored_fields", [])
    meta.setdefault("blocked_status_names", ["Blocked"])  # Jira status names that count as blocked
    meta.setdefault("blocked_reminder_config", {"type": "daily"})  # daily/weekly/on_open/time:HH:MM/never
    meta.setdefault("reminder_single_popup", True)  # True = all reminders in one popup; False = separate popup per ticket
    meta.setdefault("internal_priority_options", {})  # level -> [option1, option2, ...]
    meta.setdefault("internal_priority_option_to_level", {})
    if not meta.get("internal_priority_option_to_level") and meta.get("internal_priority_levels"):
        for lvl in meta["internal_priority_levels"]:
            opts = meta.get("internal_priority_options", {}).get(lvl, [lvl])
            for o in (opts if isinstance(opts, list) else [opts]):
                meta["internal_priority_option_to_level"][o] = lvl
    return templates, meta


def save_storage(templates, meta):
    data = {"templates": templates, "meta": meta}
    try:
        with open(TEMPLATES_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        debug_log("Failed to save storage:\n" + traceback.format_exc())
