"""DescMixin — all description-field logic for TabForm.

Extracted from tab_form.py so the main form-builder stays focused on
widget layout.  TabForm inherits this mixin; every method uses *self*
just like it was still part of TabForm.
"""
import html as _html
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import uuid
import webbrowser
from html.parser import HTMLParser
from tkinter import filedialog, messagebox, simpledialog, ttk

from utils import _bind_mousewheel, _bind_mousewheel_to_target, debug_log


# ── Issue-key link helpers ────────────────────────────────────────────────────

_ISSUE_KEY_RE = re.compile(r'\b([A-Z][A-Z0-9]+-\d+)\b')

def _extract_jira_key_from_url(url: str) -> str:
    """Return the Jira issue key embedded in a /browse/KEY URL, else ''."""
    m = re.search(r'/browse/([A-Z][A-Z0-9]+-\d+)', url or "")
    return m.group(1) if m else ""

def _jira_key_badge(key: str) -> str:
    """Render a Jira issue key as a styled adf-ticket:// link badge."""
    return (
        f'<a href="adf-ticket://{_html.escape(key)}" '
        f'style="color:#4a9eff;text-decoration:none;padding:1px 6px;'
        f'background:#1e3a5f;border:1px solid #4a9eff;border-radius:3px;'
        f'font-size:12px;font-family:Consolas,monospace;white-space:nowrap">'
        f'{_html.escape(key)}</a>'
    )

def _autolink_issue_keys(html_text: str) -> str:
    """Replace bare issue keys (PROJ-123) with ticket badge links in escaped HTML."""
    return _ISSUE_KEY_RE.sub(lambda m: _jira_key_badge(m.group(1)), html_text)


# ── Dark-theme CSS injected into every viewer frame ──────────────────────────

_VIEWER_CSS = """
body {
    background: #1e1e1e;
    color: #dcdcdc;
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 13px;
    padding: 8px 12px;
    margin: 0;
}
p { margin: 4px 0; }
h1,h2,h3,h4,h5,h6 { color: #dcdcdc; margin: 8px 0 4px 0; }
a { color: #4a9eff; text-decoration: none; }
a:hover { text-decoration: underline; }
code {
    background: #2d2d2d;
    padding: 2px 4px;
    border-radius: 3px;
    font-family: Consolas, monospace;
    font-size: 12px;
}
pre {
    background: #2d2d2d;
    padding: 8px;
    border-radius: 4px;
    overflow-x: auto;
    margin: 4px 0;
}
pre code { background: none; padding: 0; }
blockquote {
    border-left: 4px solid #555;
    margin: 8px 0;
    padding-left: 12px;
    color: #aaa;
}
hr { border: none; border-top: 1px solid #555; margin: 12px 0; }
ul, ol { padding-left: 20px; margin: 4px 0; }
li { margin: 2px 0; }
table {
    border-collapse: collapse;
    width: 100%;
    margin: 8px 0 16px 0;
}
th {
    background: #3d5a80;
    color: #e0e0e0;
    border: 1px solid #555;
    padding: 8px 10px;
    text-align: left;
    font-weight: bold;
}
td {
    border: 1px solid #555;
    padding: 8px 10px;
    vertical-align: top;
}
"""


class DescMixin:
    """Mixin providing ADF/HTML editor logic for TabForm.

    Expects TabForm to provide the following instance attributes (set in
    TabForm.__init__ before any DescMixin method is called):
        self.frame              — the root Tkinter frame
        self.field_widgets      — dict[header -> {widget, ...}]
        self.field_menu_cb      — right-click menu callback or None
        self.collect_vars_cb         — returns {KEY: value} dict or None
        self.notify_vars_changed_cb  — called after editor saves so the app
                                       can refresh live variable previews; None ok
        self.extract_text_from_adf_cb — callback or None
        self._suppress_sync     — bool (debounce guard)
        self._adf_after_id      — after() cancel token
    """

    # ══════════════════════════════════════════════════════════════════════════
    # Field setup — called from TabForm._build_fields
    # ══════════════════════════════════════════════════════════════════════════

    def _build_desc_field(self, parent, hdr):
        """Build the Description ADF area inside *parent*.

        Layout:
            [toolbar row]
            [hidden JSON store — not packed, canonical ADF]
            [HtmlFrame read-only viewer — fills remaining space]

        Returns the hidden tk.Text widget that stores the canonical ADF
        JSON (registered under field_widgets[hdr]["widget"]).
        """
        # ── Toolbar ──────────────────────────────────────────────────────────
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", pady=(0, 4))
        ttk.Button(toolbar, text="Table",      width=8,  command=self._adf_insert_table).pack(side="left", padx=(0, 4))
        ttk.Button(toolbar, text="Smart list", width=10, command=self._adf_insert_task_list).pack(side="left", padx=(0, 4))
        ttk.Button(toolbar, text="Link",       width=6,  command=self._adf_insert_link).pack(side="left", padx=(0, 4))
        ttk.Button(toolbar, text="Smart link", width=10, command=self._adf_insert_smart_link).pack(side="left", padx=(0, 4))
        ttk.Button(toolbar, text="Attach",     width=7,  command=self._on_attach_files).pack(side="left", padx=(0, 4))
        ttk.Button(toolbar, text="Jira Preview", width=12, command=self._show_jira_preview).pack(side="right", padx=(4, 0))
        ttk.Button(toolbar, text="✏ Edit",    width=8,  command=self._open_wysiwyg_editor).pack(side="right", padx=(0, 4))

        # ── Hidden JSON store (canonical ADF, never shown to user) ────────────
        json_frame = ttk.Frame(parent)   # NOT packed — lives off-screen
        txt_json = tk.Text(
            json_frame, height=12, wrap="none",
            bg="#1e1e1e", fg="#dcdcdc", insertbackground="#dcdcdc",
        )
        xj = ttk.Scrollbar(json_frame, orient="horizontal", command=txt_json.xview)
        yj = ttk.Scrollbar(json_frame, orient="vertical",   command=txt_json.yview)
        txt_json.configure(xscrollcommand=xj.set, yscrollcommand=yj.set)
        yj.pack(side="right",  fill="y")
        xj.pack(side="bottom", fill="x")
        txt_json.pack(fill="both", expand=True)
        _bind_mousewheel(txt_json, "vertical")
        if self.field_menu_cb:
            txt_json.bind("<Button-3>", lambda e, h=hdr: self.field_menu_cb(e, h, self))
        txt_json.bind("<KeyRelease>", lambda e: self._on_adf_key())

        # ── HtmlFrame read-only viewer ────────────────────────────────────────
        viewer_frame = ttk.Frame(parent)
        viewer_frame.pack(fill="both", expand=True)
        self._desc_html_viewer = None
        try:
            from tkinterweb import HtmlFrame  # type: ignore
            viewer = HtmlFrame(viewer_frame, messages_enabled=False, dark_theme_enabled=True)
            viewer.pack(fill="both", expand=True)
            # Wire link clicks (adf-toggle://, regular URLs)
            try:
                viewer.on_link_click = lambda url: self._on_viewer_link_click(url)
            except Exception:
                pass
            try:
                viewer.html.on_link_click = lambda url: self._on_viewer_link_click(url)
            except Exception:
                pass
            self._desc_html_viewer = viewer
        except Exception:
            import tkinter.scrolledtext as _st
            fb = _st.ScrolledText(
                viewer_frame, wrap="word",
                bg="#1e1e1e", fg="#dcdcdc",
                font=("Segoe UI", 10), padx=8, pady=8,
                state="disabled",
            )
            fb.pack(fill="both", expand=True)
            self._desc_html_viewer = fb

        # ── Initialise state attrs referenced by other methods ────────────────
        self._desc_browser          = None   # no embedded browser
        self._desc_editor           = None   # no plain-text editor
        self._editor_page_loaded    = False
        self._desc_view_frame       = viewer_frame
        self._desc_edit_frame       = None
        self._desc_edit_bar         = None
        self._desc_edit_toggle_btn  = None
        self._desc_edit_mode        = False
        self._adf_html_widget       = None
        self._adf_html_api          = None
        self._adf_preview_editable  = False
        self._adf_preview_widget    = None

        return txt_json

    # ══════════════════════════════════════════════════════════════════════════
    # ADF ↔ HTML conversion
    # ══════════════════════════════════════════════════════════════════════════

    _jira_image_cache: dict = {}  # url -> base64 data URL (session-level cache)

    def _fetch_jira_image_as_data_url(self, url):
        """Download a Jira attachment URL via the authenticated session and
        return a base64 data URL.  Results are cached for the session."""
        cached = self._jira_image_cache.get(url)
        if cached is not None:
            return cached
        try:
            root = self.frame.winfo_toplevel()
            session = root.get_jira_session() if hasattr(root, "get_jira_session") else None
            if not session:
                debug_log("Image fetch: no session available")
                return ""
            import requests as _req
            # Jira attachment content endpoints return 303 redirects to signed
            # URLs. Disable redirect-following so we can fetch the signed URL
            # without sending Jira auth headers (which some CDNs reject).
            resp = session.get(url, timeout=30, allow_redirects=False)
            if resp.status_code in (301, 302, 303, 307, 308):
                redirect_url = resp.headers.get("Location", "")
                if redirect_url:
                    debug_log(f"Image redirect → {redirect_url[:120]}")
                    resp = _req.get(redirect_url, timeout=30)
            if resp.status_code != 200:
                debug_log(f"Image fetch failed ({resp.status_code}): {url[:120]}")
                self._jira_image_cache[url] = ""
                return ""
            ct = resp.headers.get("Content-Type", "image/png")
            if not ct.startswith("image/"):
                debug_log(f"Image fetch non-image Content-Type ({ct}): {url[:120]}")
                self._jira_image_cache[url] = ""
                return ""
            import base64 as _b64
            data_url = f"data:{ct};base64,{_b64.b64encode(resp.content).decode('ascii')}"
            self._jira_image_cache[url] = data_url
            debug_log(f"Image fetch OK ({len(resp.content)} bytes): {url[:120]}")
            return data_url
        except Exception:
            debug_log(f"Image fetch exception for {url[:120]}: " + traceback.format_exc())
            self._jira_image_cache[url] = ""
            return ""

    def _build_media_uuid_to_url_map(self):
        """Build a mapping of Jira media UUIDs → attachment content URLs.

        Uses the rendered HTML (which Jira resolves with real URLs) to extract
        media-services-id → src mappings.  Falls back to the attachment list
        for positional matching.
        """
        uuid_map = {}

        # Strategy 1: Parse rendered HTML for image URLs
        rendered = getattr(self, "_description_rendered_html", "") or ""
        debug_log(f"Media UUID map: rendered HTML length = {len(rendered)}")

        if rendered:
            import re
            # Try both attribute orderings for data-media-services-id
            for m in re.finditer(
                r'data-media-services-id=["\']([^"\']+)["\'][^>]*src=["\']([^"\']+)["\']',
                rendered, re.IGNORECASE
            ):
                uuid_map[m.group(1)] = m.group(2)
            for m in re.finditer(
                r'src=["\']([^"\']+)["\'][^>]*data-media-services-id=["\']([^"\']+)["\']',
                rendered, re.IGNORECASE
            ):
                if m.group(2) not in uuid_map:
                    uuid_map[m.group(2)] = m.group(1)
            for m in re.finditer(
                r'data-media-id=["\']([^"\']+)["\'][^>]*src=["\']([^"\']+)["\']',
                rendered, re.IGNORECASE
            ):
                if m.group(1) not in uuid_map:
                    uuid_map[m.group(1)] = m.group(2)
            for m in re.finditer(
                r'src=["\']([^"\']+)["\'][^>]*data-media-id=["\']([^"\']+)["\']',
                rendered, re.IGNORECASE
            ):
                if m.group(2) not in uuid_map:
                    uuid_map[m.group(2)] = m.group(1)

            # Log a snippet of the rendered HTML to see what format Jira uses
            if not uuid_map:
                # Look for ANY img tags to understand the format
                imgs = re.findall(r'<img[^>]+>', rendered, re.IGNORECASE)
                debug_log(f"Media UUID map: no UUID matches found. img tags in rendered HTML: {imgs[:3]}")
                # Also look for any attachment-related patterns
                att_patterns = re.findall(r'attachment[^"\'>\s]{0,80}', rendered, re.IGNORECASE)
                debug_log(f"Media UUID map: attachment patterns: {att_patterns[:5]}")

        debug_log(f"Media UUID map result: {len(uuid_map)} UUID mappings found")

        # Strategy 2: Use attachment list — collect content URLs for image files
        att_urls = []
        try:
            raw = getattr(self, "_attachment_raw_json", None) or ""
            if raw:
                items = json.loads(raw) if isinstance(raw, str) else raw
                if isinstance(items, list):
                    _img_exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg"}
                    for att in items:
                        if not isinstance(att, dict):
                            continue
                        fn = (att.get("filename") or att.get("name") or "").lower()
                        mime = (att.get("mimeType") or "").lower()
                        url = att.get("content") or ""
                        if url and (any(fn.endswith(e) for e in _img_exts)
                                    or mime.startswith("image/")):
                            att_urls.append(url)
                            debug_log(f"Media attachment URL: {fn} → {url[:100]}")
        except Exception:
            debug_log(f"Media attachment parse error: {traceback.format_exc()}")

        debug_log(f"Media UUID map: {len(att_urls)} image attachment URL(s) for positional fallback")
        return uuid_map, att_urls

    def _convert_adf_to_html(self, node):
        """Render an ADF node tree to an HTML fragment (body content only)."""
        try:
            _media_uuid_map, _media_att_urls = self._build_media_uuid_to_url_map()
        except Exception:
            _media_uuid_map, _media_att_urls = {}, []
        _att_url_idx = [0]  # mutable counter for positional fallback

        _table_attrs = (
            'border="1" cellpadding="6" cellspacing="0" '
            'style="border-collapse:collapse;width:100%;margin:8px 0 16px 0"'
        )
        _cell_attrs = 'style="border:1px solid #555;padding:8px 10px;min-height:44px"'
        _th_attrs   = (
            'style="border:1px solid #555;padding:8px 10px;min-height:44px;'
            'background:#3d5a80;color:#e0e0e0;font-weight:bold"'
        )

        def node_to_html(n, path=""):
            if n is None:
                return ""
            if isinstance(n, str):
                return _html.escape(n)
            if isinstance(n, dict):
                t = n.get("type")
                if t == "doc":
                    return "".join(node_to_html(c, f"{path}.{i}" if path else str(i))
                                   for i, c in enumerate(n.get("content", [])))
                if t == "paragraph":
                    inner = "".join(node_to_html(c, path) for c in n.get("content", []))
                    if not inner or not inner.strip():
                        inner = "&nbsp;"
                        style = "margin:4px 0;min-height:2.2em"
                    else:
                        style = "margin:4px 0"
                    return f"<p style='{style}'>{inner}</p>"
                if t == "text":
                    raw_text = n.get("text", "")
                    text = _html.escape(raw_text)
                    for m in (n.get("marks") or []):
                        mt = m.get("type")
                        if mt == "strong":
                            text = f"<strong>{text}</strong>"
                        elif mt in ("em", "emphasis"):
                            text = f"<em>{text}</em>"
                        elif mt == "code":
                            text = f"<code>{text}</code>"
                        elif mt == "link":
                            href = (m.get("attrs") or {}).get("href", "")
                            key = _extract_jira_key_from_url(href)
                            if key:
                                text = _jira_key_badge(key)
                            else:
                                text = f'<a href="{_html.escape(href)}">{text}</a>'
                    # Auto-link bare issue keys (e.g. PROJ-123) not already wrapped
                    if "<a " not in text:
                        text = _autolink_issue_keys(text)
                    return text
                if t == "inlineCard":
                    url = (n.get("attrs") or {}).get("url", "")
                    if url:
                        key = _extract_jira_key_from_url(url)
                        if key:
                            return _jira_key_badge(key)
                        return (
                            f'<a href="{_html.escape(url)}" '
                            f'style="color:#4a9eff;text-decoration:underline;padding:2px 6px;'
                            f'background:#2d3748;border-radius:4px;display:inline-block;margin:2px 0">'
                            f'{_html.escape(url)}</a>'
                        )
                    return ""
                if t == "mention":
                    text = (n.get("attrs") or {}).get("text", "")
                    if not text:
                        text = "@" + ((n.get("attrs") or {}).get("id", "unknown"))
                    return (
                        f'<span style="color:#4a9eff;background:#1a365d;'
                        f'padding:1px 4px;border-radius:3px">'
                        f'{_html.escape(text)}</span>'
                    )
                if t == "emoji":
                    shortName = (n.get("attrs") or {}).get("shortName", "")
                    text = (n.get("attrs") or {}).get("text", shortName)
                    return _html.escape(text) if text else ""
                if t == "status":
                    text = (n.get("attrs") or {}).get("text", "")
                    color = (n.get("attrs") or {}).get("color", "neutral")
                    _status_colors = {
                        "neutral": "#555", "purple": "#6554c0",
                        "blue": "#0065ff", "red": "#de350b",
                        "yellow": "#ff991f", "green": "#36b37e",
                    }
                    bg = _status_colors.get(color, "#555")
                    return (
                        f'<span style="background:{bg};color:#fff;'
                        f'padding:2px 6px;border-radius:3px;font-size:0.85em;'
                        f'font-weight:bold">{_html.escape(text)}</span>'
                    )
                if t == "date":
                    ts = (n.get("attrs") or {}).get("timestamp", "")
                    return _html.escape(ts)
                if t == "heading":
                    lvl  = min(6, max(1, int((n.get("attrs") or {}).get("level", 1))))
                    inner = "".join(node_to_html(c, path) for c in n.get("content", []))
                    return f"<h{lvl} style='margin:8px 0 4px 0'>{inner}</h{lvl}>"
                if t == "bulletList":
                    items = "".join(
                        "<li style='margin:2px 0'>{}</li>".format(
                            "".join(node_to_html(c, path) for c in li.get("content", []))
                        ) for li in n.get("content", [])
                    )
                    return f"<ul style='margin:4px 0;padding-left:20px'>{items}</ul>"
                if t == "orderedList":
                    items = "".join(
                        "<li style='margin:2px 0'>{}</li>".format(
                            "".join(node_to_html(c, path) for c in li.get("content", []))
                        ) for li in n.get("content", [])
                    )
                    return f"<ol style='margin:4px 0;padding-left:20px'>{items}</ol>"
                if t == "listItem":
                    inner = "".join(node_to_html(c, path) for c in n.get("content", []))
                    return f"<li>{inner}</li>"
                if t in ("taskList", "actionList"):
                    items = "".join(
                        node_to_html(c, f"{path}.{i}" if path else str(i))
                        for i, c in enumerate(n.get("content", []))
                    )
                    return f"<ul class='task-list' style='list-style:none;padding-left:0;margin:4px 0'>{items}</ul>"
                if t in ("taskItem", "action"):
                    state    = (n.get("attrs") or {}).get("state", "TODO")
                    chk_char = "&#9745;" if state == "DONE" else "&#9744;"
                    inner    = "".join(node_to_html(c, path) for c in n.get("content", []))
                    href     = f"adf-toggle://{_html.escape(path)}" if path else ""
                    if href:
                        return (
                            f"<li style='margin:6px 0;display:flex;align-items:flex-start'>"
                            f"<a href=\"{href}\" style='margin-right:8px;margin-top:3px;"
                            f"cursor:pointer;text-decoration:none;font-size:1.1em;color:inherit' "
                            f"title='Click to toggle'>{chk_char}</a>"
                            f"<span>{inner}</span></li>"
                        )
                    return (
                        f"<li style='margin:6px 0;display:flex;align-items:flex-start'>"
                        f"<span style='margin-right:8px'>{chk_char}</span>"
                        f"<span>{inner}</span></li>"
                    )
                if t == "codeBlock":
                    code = "".join(c.get("text", "") for c in n.get("content", []) if c.get("type") == "text")
                    return (
                        f"<pre style='background:#2d2d2d;padding:8px;border-radius:4px;"
                        f"overflow-x:auto;margin:4px 0'><code>{_html.escape(code)}</code></pre>"
                    )
                if t == "blockquote":
                    inner = "".join(node_to_html(c, path) for c in n.get("content", []))
                    return f"<blockquote style='border-left:4px solid #555;margin:8px 0;padding-left:12px;color:#aaa'>{inner}</blockquote>"
                if t == "rule":
                    return "<hr style='border:none;border-top:1px solid #555;margin:12px 0'>"
                if t == "hardBreak":
                    return "<br>"
                if t == "table":
                    rows_html = ""
                    for rown in n.get("content", []):
                        cols_html = ""
                        for cell in (rown.get("content") or []):
                            cell_html = "".join(node_to_html(c, path) for c in (cell.get("content") or []))
                            if not cell_html or not cell_html.strip():
                                cell_html = "&nbsp;"
                            cell_html = f"<div style='min-height:2.2em'>{cell_html}</div>"
                            tag   = "th" if cell.get("type") == "tableHeader" else "td"
                            attrs = _th_attrs if tag == "th" else _cell_attrs
                            cols_html += f"<{tag} {attrs}>{cell_html}</{tag}>"
                        rows_html += f"<tr>{cols_html}</tr>"
                    return f"<table {_table_attrs}><tbody>{rows_html}</tbody></table>"
                if t == "panel":
                    inner = "".join(node_to_html(c, path) for c in n.get("content", []))
                    ptype = (n.get("attrs") or {}).get("panelType", "info")
                    colors = {"info": "#2d3748", "note": "#2c5282", "success": "#276749",
                              "warning": "#744210", "error": "#742a2a"}
                    bg = colors.get(str(ptype), "#2d3748")
                    return (
                        f"<div style='background:{bg};border-left:4px solid #63b3ed;"
                        f"padding:8px 12px;margin:8px 0;border-radius:4px'>{inner}</div>"
                    )
                if t in ("mediaSingle", "mediaGroup"):
                    return "".join(node_to_html(c, path) for c in n.get("content", []))
                if t == "media":
                    attrs = n.get("attrs") or {}
                    fname = attrs.get("__fileName") or ""
                    pending = attrs.get("__pendingPath") or ""
                    jira_url = attrs.get("url") or ""
                    att_id = attrs.get("id") or ""
                    alt = _html.escape(fname) if fname else "image"
                    if pending and os.path.isfile(pending):
                        # Embed as base64 so both the viewer and the editor
                        # can display the image (file:// is blocked by WebView2).
                        try:
                            import base64 as _b64
                            with open(pending, "rb") as _fh:
                                _raw = _fh.read()
                            ext = os.path.splitext(fname)[1].lower().lstrip(".") or "png"
                            _mime_map = {"jpg": "jpeg", "svg": "svg+xml"}
                            _mime = f"image/{_mime_map.get(ext, ext)}"
                            _data_url = f"data:{_mime};base64,{_b64.b64encode(_raw).decode('ascii')}"
                        except Exception:
                            _data_url = ""
                        if _data_url:
                            # Single <img> tag — no wrapper divs or captions.
                            # The editor wraps it visually; the parser only
                            # sees one <img> and produces one mediaSingle.
                            return (
                                f'<img src="{_data_url}" alt="{alt}" '
                                f'data-avalanche-file="{alt}" '
                                f'style="max-width:100%;border:1px solid #444;'
                                f'border-radius:4px;display:block;margin:8px auto">'
                            )
                        return (
                            f'<div style="text-align:center;margin:8px 0;padding:12px;'
                            f'background:#2d2d2d;border:1px solid #444;border-radius:4px">'
                            f'📎 {alt} (file not readable)</div>'
                        )
                    if jira_url:
                        data_url = self._fetch_jira_image_as_data_url(jira_url)
                        if data_url:
                            return (
                                f'<img src="{data_url}" alt="{alt}" '
                                f'style="max-width:100%;border:1px solid #444;'
                                f'border-radius:4px;display:block;margin:8px auto">'
                            )
                        src = _html.escape(jira_url)
                        return (
                            f'<div style="text-align:center;margin:8px 0;padding:12px;'
                            f'background:#2d2d2d;border:1px solid #444;border-radius:4px">'
                            f'📎 {alt} (could not load from Jira)</div>'
                        )
                    if att_id:
                        debug_log(f"Media node: id={att_id}, fname={fname}, url={jira_url}")
                        resolved_url = ""
                        # 1) Check the UUID→URL map (from rendered HTML)
                        if att_id in _media_uuid_map:
                            resolved_url = _media_uuid_map[att_id]
                            debug_log(f"  → UUID map hit: {resolved_url[:100]}")
                        # 2) Positional fallback from the attachment list
                        if not resolved_url and _media_att_urls:
                            idx = _att_url_idx[0]
                            if idx < len(_media_att_urls):
                                resolved_url = _media_att_urls[idx]
                                debug_log(f"  → Positional fallback [{idx}]: {resolved_url[:100]}")
                            _att_url_idx[0] = idx + 1
                        if not resolved_url:
                            debug_log(f"  → No URL resolved for media {att_id}")
                        if resolved_url:
                            data_url = self._fetch_jira_image_as_data_url(resolved_url)
                            if data_url:
                                return (
                                    f'<img src="{data_url}" alt="{alt}" '
                                    f'style="max-width:100%;border:1px solid #444;'
                                    f'border-radius:4px;display:block;margin:8px auto">'
                                )
                            debug_log(f"  → Fetch failed for resolved URL")
                        # 3) Last resort: try the REST endpoint directly
                        try:
                            root = self.frame.winfo_toplevel()
                            s = root.get_jira_session() if hasattr(root, "get_jira_session") else None
                            if s:
                                direct_url = f"{s._jira_base}/rest/api/3/attachment/content/{att_id}"
                                debug_log(f"  → Trying direct REST: {direct_url[:100]}")
                                data_url = self._fetch_jira_image_as_data_url(direct_url)
                                if data_url:
                                    return (
                                        f'<img src="{data_url}" alt="{alt}" '
                                        f'style="max-width:100%;border:1px solid #444;'
                                        f'border-radius:4px;display:block;margin:8px auto">'
                                    )
                        except Exception:
                            pass
                        debug_log(f"  → All strategies failed for media {att_id}")
                        return (
                            f'<div style="text-align:center;margin:8px 0;padding:12px;'
                            f'background:#2d2d2d;border:1px solid #444;border-radius:4px">'
                            f'📎 Attachment (id: {_html.escape(att_id)})</div>'
                        )
                    if fname:
                        return (
                            f'<div style="text-align:center;margin:8px 0;padding:12px;'
                            f'background:#2d2d2d;border:1px solid #444;border-radius:4px">'
                            f'📎 {alt}</div>'
                        )
                    return ""
                return "".join(node_to_html(c, path) for c in n.get("content", []))
            if isinstance(n, list):
                return "".join(node_to_html(c, path) for c in n)
            return ""

        return node_to_html(node)

    def _convert_html_to_adf(self, html_source):
        """Convert an HTML string to an ADF document dict."""
        if not html_source or not html_source.strip():
            return {"type": "doc", "version": 1, "content": [{"type": "paragraph", "content": []}]}
        src = html_source.strip().lower()
        if "<body" in src:
            start   = html_source.lower().find("<body")
            end_tag = html_source.find(">", start) + 1
            close   = html_source.lower().find("</body>", end_tag)
            if close > end_tag:
                html_source = html_source[end_tag:close]

        class _ADFBuilder(HTMLParser):
            def __init__(self):
                super().__init__()
                self.stack       = []
                self.doc_content = []
                self._cur_inline = []
                self._cur_marks  = []

            def _emit_block(self, block_type, attrs=None):
                runs = []
                for text, marks in self._cur_inline:
                    if not text:
                        continue
                    if runs and runs[-1][1] == marks:
                        runs[-1] = (runs[-1][0] + text, marks)
                    else:
                        runs.append((text, list(marks)))
                self._cur_inline = []
                content = []
                for t, m in runs:
                    n = {"type": "text", "text": t}
                    if m:
                        n["marks"] = m
                    content.append(n)
                if attrs:
                    return {"type": block_type, "attrs": attrs, "content": content}
                return {"type": block_type, "content": content}

            def handle_starttag(self, tag, attrs):
                adict = dict(attrs)
                if tag in ("p", "div"):
                    self.stack.append((tag, [], []))
                elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                    self.stack.append((tag, {"level": int(tag[1])}, []))
                elif tag in ("ul", "ol", "blockquote", "pre", "table", "tbody", "thead", "tr", "li", "th", "td"):
                    self.stack.append((tag, [], []))
                elif tag in ("strong", "b"):
                    self._cur_marks.append({"type": "strong"})
                elif tag in ("em", "i"):
                    self._cur_marks.append({"type": "em"})
                elif tag == "code":
                    if not (self.stack and self.stack[-1][0] == "pre"):
                        self._cur_marks.append({"type": "code"})
                elif tag == "a":
                    self._cur_marks.append({"type": "link", "attrs": {"href": adict.get("href", "")}})
                elif tag == "img":
                    src = adict.get("src", "")
                    fname = adict.get("data-avalanche-file", "")
                    media_attrs = {"type": "file"}
                    if src.startswith("avalanche-pending://"):
                        media_attrs["__fileName"] = src.replace("avalanche-pending://", "")
                    elif fname:
                        media_attrs["__fileName"] = fname
                    else:
                        media_attrs["__fileName"] = os.path.basename(src) if src else "image"
                        if src and not src.startswith("data:"):
                            media_attrs["url"] = src
                    # Flush any accumulated inline text as a paragraph before
                    # inserting the block-level mediaSingle node.
                    dest = self.stack[-1][2] if self.stack else self.doc_content
                    if self._cur_inline:
                        p_node = self._emit_block("paragraph")
                        if p_node.get("content"):
                            dest.append(p_node)
                    media_node = {
                        "type": "mediaSingle",
                        "attrs": {"layout": "center"},
                        "content": [{"type": "media", "attrs": media_attrs}],
                    }
                    dest.append(media_node)
                elif tag == "br":
                    self._cur_inline.append(("\n", list(self._cur_marks)))

            def _pop_mark(self, tag):
                mtype = {"strong": "strong", "b": "strong", "em": "em", "i": "em",
                         "code": "code", "a": "link"}.get(tag)
                if not mtype:
                    return
                for i in range(len(self._cur_marks) - 1, -1, -1):
                    if self._cur_marks[i].get("type") == mtype:
                        self._cur_marks.pop(i)
                        return

            def handle_endtag(self, tag):
                if tag in ("strong", "b", "em", "i", "code", "a"):
                    self._pop_mark(tag)
                    return
                if not self.stack:
                    return
                t, a, children = self.stack.pop()
                dest = self.stack[-1][2] if self.stack else self.doc_content
                if tag in ("p", "div"):
                    if children:
                        for c in children:
                            dest.append(c)
                    else:
                        dest.append(self._emit_block("paragraph"))
                elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                    dest.append(self._emit_block("heading", {"level": int(tag[1])}))
                elif tag == "ul":
                    dest.append({"type": "bulletList", "content": children})
                elif tag == "ol":
                    dest.append({"type": "orderedList", "content": children})
                elif tag == "li":
                    blk = self._emit_block("paragraph")
                    if self.stack:
                        self.stack[-1][2].append({"type": "listItem", "content": [blk]})
                elif tag == "blockquote":
                    dest.append({"type": "blockquote", "content": children})
                elif tag == "pre":
                    raw = "".join(tx for tx, _ in self._cur_inline).strip()
                    self._cur_inline = []
                    dest.append({"type": "codeBlock", "content": [{"type": "text", "text": raw}]})
                elif tag == "hr":
                    dest.append({"type": "rule"})
                elif tag == "tr":
                    if self.stack:
                        self.stack[-1][2].append({"type": "tableRow", "content": children})
                elif tag in ("tbody", "thead"):
                    if self.stack:
                        self.stack[-1][2].extend(children)
                elif tag in ("th", "td"):
                    cell_type = "tableHeader" if tag == "th" else "tableCell"
                    if children:
                        block_children, inline_buf = [], []
                        block_node_types = {
                            "paragraph", "heading", "bulletList", "orderedList",
                            "codeBlock", "blockquote", "table", "rule",
                        }
                        for child in children:
                            if isinstance(child, dict) and child.get("type") in block_node_types:
                                if inline_buf:
                                    block_children.append({"type": "paragraph", "content": inline_buf})
                                    inline_buf = []
                                block_children.append(child)
                            else:
                                inline_buf.append(child)
                        if inline_buf:
                            block_children.append({"type": "paragraph", "content": inline_buf})
                        if not block_children:
                            block_children = [{"type": "paragraph", "content": []}]
                        blk = {"type": cell_type, "content": block_children}
                    else:
                        blk = {"type": cell_type, "content": [self._emit_block("paragraph")]}
                    if self.stack:
                        self.stack[-1][2].append(blk)
                elif tag == "table":
                    node = {"type": "table", "content": children}
                    dest.append(node)

            def handle_data(self, data):
                self._cur_inline.append((data, list(self._cur_marks)))

        try:
            parser = _ADFBuilder()
            parser.feed(html_source)
            if parser._cur_inline:
                parser.doc_content.append(parser._emit_block("paragraph"))
            content = parser.doc_content or [{"type": "paragraph", "content": []}]
            return {"type": "doc", "version": 1, "content": content}
        except Exception:
            return {"type": "doc", "version": 1, "content": [{"type": "paragraph", "content": []}]}

    # ══════════════════════════════════════════════════════════════════════════
    # Variable styling
    # ══════════════════════════════════════════════════════════════════════════

    def _style_variables_in_html(self, html_body):
        """Post-process HTML body: strip {KEY=value} definitions, style {KEY} refs."""
        if not html_body:
            return html_body
        vars_dict = {}
        if self.collect_vars_cb:
            try:
                vars_dict = self.collect_vars_cb() or {}
            except Exception:
                pass
        html_body = re.sub(
            r'\{([A-Z])=([^}]+)\}',
            lambda m: _html.escape(m.group(2).strip()),
            html_body,
        )

        def _style_ref(m):
            k = m.group(1)
            if k not in vars_dict:
                return m.group(0)
            esc_val = _html.escape(str(vars_dict[k]))
            return (
                f'<span style="background:#1a3a2a;color:#4ec9b0;padding:1px 5px;'
                f'border-radius:3px;font-style:italic;cursor:help;white-space:nowrap"'
                f' title="{k} = {esc_val}">{{{k}}}</span>'
            )

        return re.sub(r'\{([A-Z])\}', _style_ref, html_body)

    # ══════════════════════════════════════════════════════════════════════════
    # Viewer refresh
    # ══════════════════════════════════════════════════════════════════════════

    def _refresh_html_viewer(self):
        """Render the current ADF JSON store as HTML into the read-only HtmlFrame."""
        viewer = getattr(self, "_desc_html_viewer", None)
        if not viewer:
            return
        try:
            info = self.field_widgets.get("Description ADF")
            if not info:
                return
            raw = (info["widget"].get("1.0", "end") or "").strip()
            if raw:
                try:
                    node     = json.loads(raw)
                    body_html = self._convert_adf_to_html(node)
                    body_html = self._style_variables_in_html(body_html)
                except Exception:
                    body_html = ""
            else:
                body_html = ""

            if not body_html:
                body_html = (
                    '<p style="color:#666;font-style:italic">'
                    'Click <b>✏ Edit</b> to add a description.</p>'
                )

            full_html = (
                f'<!doctype html><html><head><meta charset="utf-8">'
                f'<style>{_VIEWER_CSS}</style></head>'
                f'<body>{body_html}</body></html>'
            )

            try:
                viewer.load_html(full_html)
                self._bind_viewer_scroll()
            except AttributeError:
                # Fallback: plain tk.Text widget
                try:
                    import re as _re
                    plain = _re.sub(r"<[^>]+>", " ", body_html).strip()
                    viewer.config(state="normal")
                    viewer.delete("1.0", "end")
                    viewer.insert("1.0", plain)
                    viewer.config(state="disabled")
                except Exception:
                    pass
        except Exception:
            pass

    def _bind_viewer_scroll(self):
        """Bind mousewheel on the HtmlFrame viewer and its internal widgets
        to scroll the main ticket canvas instead of being swallowed."""
        viewer = getattr(self, "_desc_html_viewer", None)
        canvas = getattr(self, "_scroll_canvas", None)
        if not viewer or not canvas:
            return

        def _deep_bind(w):
            try:
                if not isinstance(w, tk.Text):
                    _bind_mousewheel_to_target(w, canvas, "vertical")
                for child in w.winfo_children():
                    _deep_bind(child)
            except Exception:
                pass

        # Run immediately and again after a delay (HtmlFrame creates
        # internal widgets asynchronously after load_html returns).
        _deep_bind(viewer)
        try:
            self.frame.after(200, lambda: _deep_bind(viewer))
            self.frame.after(600, lambda: _deep_bind(viewer))
        except Exception:
            pass

    def _update_preview_from_adf(self):
        """Refresh the viewer after a toolbar action or ADF store edit."""
        self._refresh_html_viewer()

    def _load_editor_html(self, body_html: str):
        """Called by populate_from_dict — just refreshes the viewer.

        The *body_html* argument is ignored because _refresh_html_viewer
        reads directly from the canonical JSON store, ensuring consistency.
        """
        self._refresh_html_viewer()

    def _sync_htmltext_to_adf(self, *_):
        """No-op: the HtmlFrame viewer is read-only; the JSON store is always
        the canonical source and never needs syncing from the viewer."""
        pass

    def _sync_text_to_adf(self, *_):
        """Legacy alias."""
        self._sync_htmltext_to_adf()

    def _on_preview_edited(self):
        """Legacy no-op (HtmlFrame is read-only)."""
        pass

    # ══════════════════════════════════════════════════════════════════════════
    # ADF JSON store access
    # ══════════════════════════════════════════════════════════════════════════

    def _adf_get_doc(self):
        """Return the current ADF document from the hidden JSON store."""
        info = self.field_widgets.get("Description ADF")
        if not info:
            return None
        w = info.get("widget")
        if not w:
            return None
        try:
            raw = w.get("1.0", "end").strip()
            if not raw:
                return {"type": "doc", "version": 1, "content": []}
            return json.loads(raw)
        except Exception:
            return {"type": "doc", "version": 1, "content": []}

    def _adf_set_doc(self, doc):
        """Write *doc* to the JSON store and refresh the viewer."""
        info = self.field_widgets.get("Description ADF")
        if not info:
            return
        w = info.get("widget")
        if not w:
            return
        try:
            w.config(state="normal")
            w.delete("1.0", "end")
            w.insert("1.0", json.dumps(doc, ensure_ascii=False, indent=2))
            self._on_adf_key()
        except Exception:
            pass

    def _adf_toggle_task(self, path_or_local_id):
        """Toggle taskItem/action state (TODO↔DONE). Returns True if toggled."""
        doc = self._adf_get_doc()
        if not doc or not path_or_local_id:
            return False

        if "." in str(path_or_local_id) or str(path_or_local_id).isdigit():
            try:
                indices = [int(x) for x in str(path_or_local_id).split(".") if x.strip()]
                node = doc
                for i in indices:
                    node = node.get("content", [])[i]
                if node.get("type") in ("taskItem", "action"):
                    attrs = node.get("attrs") or {}
                    attrs["state"] = "DONE" if attrs.get("state", "TODO") == "TODO" else "TODO"
                    node["attrs"] = attrs
                    self._adf_set_doc(doc)
                    return True
            except (IndexError, KeyError, ValueError):
                pass

        local_id = str(path_or_local_id).strip()

        def _find_and_toggle(nodes):
            for n in (nodes or []):
                if isinstance(n, dict):
                    if n.get("type") in ("taskItem", "action"):
                        attrs = n.get("attrs") or {}
                        if attrs.get("localId") == local_id:
                            attrs["state"] = "DONE" if attrs.get("state", "TODO") == "TODO" else "TODO"
                            n["attrs"] = attrs
                            self._adf_set_doc(doc)
                            return True
                    if _find_and_toggle(n.get("content", [])):
                        return True
            return False

        return _find_and_toggle(doc.get("content", []))

    # ══════════════════════════════════════════════════════════════════════════
    # Debounced sync: hidden JSON store → Description field + viewer refresh
    # ══════════════════════════════════════════════════════════════════════════

    def _on_adf_key(self):
        """Debounced: update Description plain-text shadow and refresh viewer."""
        try:
            if self._suppress_sync:
                return
            if self._adf_after_id:
                try:
                    self.frame.after_cancel(self._adf_after_id)
                except Exception:
                    pass

            def do_sync():
                try:
                    info = self.field_widgets.get("Description ADF")
                    if not info:
                        return
                    adf_widget = info.get("widget")
                    raw = adf_widget.get("1.0", "end").strip()
                    if not raw:
                        self._update_preview_from_adf()
                        return
                    try:
                        node = json.loads(raw)
                    except Exception:
                        self._update_preview_from_adf()
                        return

                    # Sync plain-text Description shadow field
                    desc_info = self.field_widgets.get("Description")
                    if desc_info:
                        if self.extract_text_from_adf_cb:
                            text = self.extract_text_from_adf_cb(node)
                        else:
                            def _walk_text(n):
                                out = []
                                if isinstance(n, dict):
                                    if n.get("type") == "text" and "text" in n:
                                        out.append(n["text"])
                                    for c in n.get("content", []):
                                        out.extend(_walk_text(c))
                                elif isinstance(n, list):
                                    for c in n:
                                        out.extend(_walk_text(c))
                                return out
                            text = " ".join(s.strip() for s in _walk_text(node) if s.strip())
                        dw = desc_info.get("widget")
                        dw.config(state="normal")
                        dw.delete("1.0", "end")
                        dw.insert("1.0", text)

                    self._update_preview_from_adf()
                except Exception:
                    pass

            self._adf_after_id = self.frame.after(400, do_sync)
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════════════
    # Viewer link-click handler
    # ══════════════════════════════════════════════════════════════════════════

    def _on_viewer_link_click(self, url: str):
        """Route link clicks from the HtmlFrame viewer."""
        try:
            if url.startswith("adf-edit://"):
                self._open_wysiwyg_editor()
            elif url.startswith("adf-toggle://"):
                import urllib.parse as _up
                path = _up.unquote(url[len("adf-toggle://"):])
                self._adf_toggle_task(path)
                self._update_preview_from_adf()
            elif url.startswith("adf-ticket://"):
                key = url[len("adf-ticket://"):]
                self._show_ticket_link_menu(key)
            elif url.startswith(("http://", "https://")):
                # Check if the URL is a Jira browse link — offer the open menu
                key = _extract_jira_key_from_url(url)
                if key:
                    self._show_ticket_link_menu(key)
                else:
                    webbrowser.open(url)
        except Exception:
            pass

    def _show_ticket_link_menu(self, key: str):
        """Show a small popup menu to open *key* in-app or in Jira."""
        try:
            menu = tk.Menu(self.frame, tearoff=0)
            if getattr(self, "open_ticket_in_app_cb", None):
                menu.add_command(
                    label=f"Open  {key}  locally",
                    command=lambda: self.open_ticket_in_app_cb(key),
                )
            if getattr(self, "open_ticket_in_jira_cb", None):
                menu.add_command(
                    label=f"Open  {key}  in Jira",
                    command=lambda: self.open_ticket_in_jira_cb(key),
                )
            # Post at current pointer position
            x = self.frame.winfo_pointerx()
            y = self.frame.winfo_pointery()
            menu.tk_popup(x, y)
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════════════
    # pywebview subprocess editor
    # ══════════════════════════════════════════════════════════════════════════

    def _open_wysiwyg_editor(self):
        """Launch the WYSIWYG editor (editor_window.py) as a subprocess.

        The editor reads body HTML from a temp args file, writes the result
        to a temp output file, and exits.  We watch for the result in a
        background thread and call _apply_html_from_editor on the main thread.
        """
        if getattr(self, "_editor_args_path", None):
            return  # editor subprocess already running — don't launch a second
        doc       = self._adf_get_doc()
        body_html = self._convert_adf_to_html(doc) if doc else ""

        # Collect current variable definitions so the editor can offer them
        vars_dict = {}
        if self.collect_vars_cb:
            try:
                vars_dict = self.collect_vars_cb() or {}
            except Exception:
                pass

        try:
            args_fd, args_path = tempfile.mkstemp(suffix="_editor_args.json")
            out_fd,  out_path  = tempfile.mkstemp(suffix="_editor_out.html")
            os.close(args_fd)
            os.close(out_fd)
            images_dir = tempfile.mkdtemp(prefix="aval_img_")
            try:
                _s = self.get_jira_session() if hasattr(self, "get_jira_session") else None
                _jira_base = getattr(_s, "_jira_base", "") if _s else ""
            except Exception:
                _jira_base = ""
            with open(args_path, "w", encoding="utf-8") as f:
                json.dump(
                    {"html": body_html, "output": out_path, "vars": vars_dict,
                     "jira_base": _jira_base, "images_dir": images_dir},
                    f, ensure_ascii=False,
                )
        except Exception:
            return

        # Keep the path so _push_vars_to_editor can update it while editor runs
        self._editor_args_path = args_path

        pkg_dir       = os.path.dirname(os.path.abspath(__file__))
        editor_script = os.path.join(pkg_dir, "editor_window.py")

        # When frozen as a PyInstaller EXE, sys.executable is the app's own EXE.
        # Re-launch it with --editor so the entry point routes to editor_window.main()
        # instead of starting the full application again.
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--editor", args_path]
        else:
            cmd = [sys.executable, editor_script, args_path]

        def _run():
            try:
                proc = subprocess.Popen(
                    cmd,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                proc.wait()
            except Exception:
                proc = None
            finally:
                try:
                    self.frame.after(0, lambda: _on_done(proc))
                except Exception:
                    pass

        def _on_done(proc):
            self._editor_args_path = None
            try:
                os.unlink(args_path)
            except Exception:
                pass
            result_html = None
            if os.path.exists(out_path):
                try:
                    with open(out_path, encoding="utf-8") as f:
                        content = f.read().strip()
                    if content:
                        result_html = content
                except Exception:
                    pass
                try:
                    os.unlink(out_path)
                except Exception:
                    pass
            # Collect any pending images the editor wrote to images_dir
            pending_imgs = {}  # filename -> absolute path
            try:
                if os.path.isdir(images_dir):
                    for fname in os.listdir(images_dir):
                        fpath = os.path.join(images_dir, fname)
                        if os.path.isfile(fpath) and os.path.getsize(fpath) > 0:
                            pending_imgs[fname] = fpath
            except Exception:
                pass

            if result_html is not None:
                self._apply_html_from_editor(result_html)
                # Patch the ADF with __pendingPath so the viewer can show
                # local previews and upload.py can find the files later.
                if pending_imgs:
                    self._patch_adf_pending_images(pending_imgs)
            else:
                try:
                    os.unlink(out_path)
                except Exception:
                    pass
                # Clean up images dir if editor was cancelled
                try:
                    import shutil
                    shutil.rmtree(images_dir, ignore_errors=True)
                except Exception:
                    pass
            # Check if the editor requested to open a ticket in-app
            open_ticket_path = args_path + ".open_ticket"
            if os.path.exists(open_ticket_path):
                try:
                    with open(open_ticket_path, encoding="utf-8") as f:
                        ot = json.load(f)
                    key = (ot.get("key") or "").strip()
                    if key and getattr(self, "open_ticket_in_app_cb", None):
                        self.open_ticket_in_app_cb(key)
                except Exception:
                    pass
                try:
                    os.unlink(open_ticket_path)
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()

    def _push_vars_to_editor(self):
        """Rewrite the live args file with fresh variable definitions.

        Called whenever the main window's variables change while the editor
        subprocess is open.  The editor polls get_vars() periodically and will
        pick up the new entries within a couple of seconds.
        """
        args_path = getattr(self, "_editor_args_path", None)
        if not args_path or not os.path.exists(args_path):
            return
        if not self.collect_vars_cb:
            return
        try:
            vars_dict = self.collect_vars_cb() or {}
            with open(args_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            data["vars"] = vars_dict
            with open(args_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception:
            pass

    def _apply_html_from_editor(self, html: str):
        """Convert editor HTML → ADF, store it, refresh the HtmlFrame viewer."""
        try:
            adf = self._convert_html_to_adf(html)
            if not adf:
                return
            info = self.field_widgets.get("Description ADF")
            if not info:
                return
            w = info.get("widget")
            if not isinstance(w, tk.Text):
                return
            self._suppress_sync = True
            try:
                w.config(state="normal")
                w.delete("1.0", "end")
                w.insert("1.0", json.dumps(adf, ensure_ascii=False, indent=2))
            finally:
                self._suppress_sync = False
            self._refresh_html_viewer()
            notify = getattr(self, "notify_vars_changed_cb", None)
            if callable(notify):
                self.frame.after(50, notify)
        except Exception:
            pass

    def _patch_adf_pending_images(self, pending_imgs: dict):
        """Walk the ADF tree and set __pendingPath on media nodes whose
        __fileName matches a file the editor just wrote to disk."""
        doc = self._adf_get_doc()
        if not doc:
            return
        changed = False

        def _walk(node):
            nonlocal changed
            if not isinstance(node, dict):
                return
            if node.get("type") == "media":
                attrs = node.get("attrs") or {}
                fname = attrs.get("__fileName", "")
                if fname and fname in pending_imgs:
                    attrs["__pendingPath"] = pending_imgs[fname]
                    node["attrs"] = attrs
                    changed = True
            for child in node.get("content", []):
                _walk(child)

        _walk(doc)
        if changed:
            self._adf_set_doc(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # Jira light-theme preview popup
    # ══════════════════════════════════════════════════════════════════════════

    def _resolve_variables_in_html(self, html_body):
        """Fully substitute variable markers for the Jira preview.

        {A=value} definitions are stripped (replaced with just their value).
        {A} references are replaced with the resolved value from collect_vars_cb,
        or left as plain text if no value is known.
        """
        if not html_body:
            return html_body
        vars_dict = {}
        if self.collect_vars_cb:
            try:
                vars_dict = self.collect_vars_cb() or {}
            except Exception:
                pass
        # Strip {A=value} → value
        html_body = re.sub(
            r'\{([A-Z])=([^}]+)\}',
            lambda m: _html.escape(m.group(2).strip()),
            html_body,
        )
        # Replace {A} → resolved value (or keep as-is if unknown)
        def _resolve(m):
            k = m.group(1)
            if k in vars_dict:
                return _html.escape(str(vars_dict[k]))
            return m.group(0)          # unknown var — leave unchanged

        return re.sub(r'\{([A-Z])\}', _resolve, html_body)

    def _show_jira_preview(self):  # noqa: C901
        """Open a Toplevel showing the description in Jira's light-theme styling."""
        if self._focus_existing_dialog("jira_preview"):
            return
        doc       = self._adf_get_doc()
        body_html = self._convert_adf_to_html(doc) if doc else ""
        body_html = self._resolve_variables_in_html(body_html)
        if not body_html:
            body_html = "<p><em>No description yet.</em></p>"

        jira_css = """
        body {
            background:#ffffff !important; color:#172b4d !important;
            font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
            font-size:14px; line-height:1.714; padding:20px 24px; margin:0;
        }
        p { color:#172b4d !important; margin:6px 0; }
        h1,h2,h3,h4,h5,h6 { color:#172b4d !important; font-weight:600; margin:16px 0 6px 0; }
        h1{font-size:24px} h2{font-size:20px} h3{font-size:16px}
        a { color:#0052cc !important; text-decoration:none; }
        a:hover { text-decoration:underline; }
        code {
            background:#f4f5f7 !important; color:#172b4d !important;
            padding:2px 4px; border-radius:3px;
            font-family:Consolas,'Liberation Mono',Menlo,monospace; font-size:12px;
        }
        pre {
            background:#f4f5f7 !important; color:#172b4d !important;
            padding:12px 16px; border-radius:3px; margin:8px 0; font-size:12px;
        }
        pre code { background:none !important; padding:0; }
        blockquote {
            border-left:3px solid #dfe1e6 !important; color:#5e6c84 !important;
            margin:10px 0; padding:4px 12px;
        }
        hr { border:none; border-top:1px solid #dfe1e6; margin:14px 0; }
        ul,ol { padding-left:24px; margin:6px 0; }
        li { color:#172b4d !important; margin:3px 0; }
        li p { margin:0; }
        ul.task-list, ul.action-list { list-style:none; padding-left:2px; }
        table { border-collapse:collapse !important; width:100%; margin:10px 0; font-size:14px; }
        th {
            background:#f4f5f7 !important; color:#172b4d !important;
            border:1px solid #dfe1e6 !important; padding:8px 12px;
            text-align:left; font-weight:600;
        }
        td {
            background:#ffffff !important; color:#172b4d !important;
            border:1px solid #dfe1e6 !important; padding:8px 12px; vertical-align:top;
        }
        td div, th div { color:#172b4d !important; }
        span[title] { background:#e3fcef !important; color:#006644 !important; }
        """

        full_html = (
            f"<!doctype html><html><head><meta charset='utf-8'>"
            f"<style>{jira_css}</style></head>"
            f"<body>{body_html}</body></html>"
        )

        win = tk.Toplevel(self.frame.winfo_toplevel())
        self._track_dialog("jira_preview", win)
        win.title("Jira Preview")
        win.configure(bg="#ffffff")
        win.update_idletasks()
        ww, wh = 920, 650
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{ww}x{wh}+{(sw - ww) // 2}+{(sh - wh) // 2}")

        hdr = tk.Frame(win, bg="#0052cc", pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Jira Preview  —  how your description will look",
                 bg="#0052cc", fg="#ffffff", font=("Segoe UI", 10, "bold"),
                 padx=16).pack(side="left")
        tk.Button(hdr, text="✕  Close", command=win.destroy,
                  bg="#0052cc", fg="#ffffff", relief="flat", bd=0, padx=12,
                  font=("Segoe UI", 9), activebackground="#0065ff",
                  activeforeground="#ffffff", cursor="hand2").pack(side="right", padx=8)

        try:
            from tkinterweb import HtmlFrame  # type: ignore
            pane    = tk.Frame(win, bg="#ffffff")
            pane.pack(fill="both", expand=True)
            preview = HtmlFrame(pane, messages_enabled=False, dark_theme_enabled=False)
            preview.pack(fill="both", expand=True)
            preview.load_html(full_html)
        except Exception:
            import tkinter.scrolledtext as _st
            txt = _st.ScrolledText(win, wrap="word", bg="#ffffff", fg="#172b4d",
                                   font=("Segoe UI", 10), padx=16, pady=12)
            txt.pack(fill="both", expand=True)
            txt.insert("1.0", body_html)
            txt.config(state="disabled")

        win.bind("<Escape>", lambda e: win.destroy())
        win.focus_force()

    # ══════════════════════════════════════════════════════════════════════════
    # Toolbar insert actions
    # ══════════════════════════════════════════════════════════════════════════

    def _adf_insert_table(self):
        """Append a 2×2 table to the ADF document."""
        doc = self._adf_get_doc()
        if not doc:
            return
        table = {
            "type": "table",
            "attrs": {"isNumberColumnEnabled": False, "layout": "default"},
            "content": [
                {"type": "tableRow", "content": [
                    {"type": "tableHeader", "content": [{"type": "paragraph",
                        "content": [{"type": "text", "text": "Header 1"}]}]},
                    {"type": "tableHeader", "content": [{"type": "paragraph",
                        "content": [{"type": "text", "text": "Header 2"}]}]},
                ]},
                {"type": "tableRow", "content": [
                    {"type": "tableCell", "content": [{"type": "paragraph",
                        "content": [{"type": "text", "text": "Cell 1"}]}]},
                    {"type": "tableCell", "content": [{"type": "paragraph",
                        "content": [{"type": "text", "text": "Cell 2"}]}]},
                ]},
            ],
        }
        doc.setdefault("content", []).append(table)
        self._adf_set_doc(doc)

    def _adf_insert_task_list(self):
        """Append a two-item smart task list to the ADF document."""
        doc = self._adf_get_doc()
        if not doc:
            return
        task_list = {
            "type": "taskList",
            "attrs": {"localId": str(uuid.uuid4())},
            "content": [
                {"type": "taskItem",
                 "attrs": {"localId": str(uuid.uuid4()), "state": "TODO"},
                 "content": [{"type": "text", "text": "Task 1"}]},
                {"type": "taskItem",
                 "attrs": {"localId": str(uuid.uuid4()), "state": "TODO"},
                 "content": [{"type": "text", "text": "Task 2"}]},
            ],
        }
        doc.setdefault("content", []).append(task_list)
        self._adf_set_doc(doc)

    def _adf_insert_link(self):
        """Prompt for URL and text, then append a link paragraph."""
        url = simpledialog.askstring("Insert Link", "URL:", parent=self.frame)
        if not url or not url.strip():
            return
        text = simpledialog.askstring(
            "Insert Link", "Link text (leave empty to use URL):",
            parent=self.frame, initialvalue=url.strip(),
        ) or url
        doc = self._adf_get_doc()
        if not doc:
            return
        para = {
            "type": "paragraph",
            "content": [{"type": "text", "text": text.strip(),
                          "marks": [{"type": "link", "attrs": {"href": url.strip()}}]}],
        }
        doc.setdefault("content", []).append(para)
        self._adf_set_doc(doc)

    def _adf_insert_smart_link(self):
        """Prompt for a URL, then append an inlineCard (smart link) paragraph."""
        url = simpledialog.askstring(
            "Insert Smart Link",
            "URL (Jira issue, Confluence page, etc.):",
            parent=self.frame,
        )
        if not url or not url.strip():
            return
        doc = self._adf_get_doc()
        if not doc:
            return
        para = {
            "type": "paragraph",
            "content": [{"type": "inlineCard", "attrs": {"url": url.strip()}}],
        }
        doc.setdefault("content", []).append(para)
        self._adf_set_doc(doc)

    def _on_attach_files(self):
        """Open a file picker and add selected paths to the Attachment field."""
        info = self.field_widgets.get("Attachment")
        if not info:
            return
        paths = filedialog.askopenfilenames(title="Attach files to ticket")
        if not paths:
            return
        widget = info.get("widget")
        var    = info.get("var")
        if not (widget and var):
            return
        try:
            current   = var.get().strip()
            raw_json  = getattr(self, "_attachment_raw_json", None)
            if raw_json and str(raw_json).strip().startswith("["):
                if not messagebox.askyesno(
                    "Attach",
                    "This ticket has Jira attachments. Add local files?\n"
                    "(They will be uploaded when you save to Jira.)",
                ):
                    return
                current = ""
            elif current and current.startswith("["):
                current = ""
            existing  = [p.strip() for p in current.split(";") if p.strip()] if current else []
            new_val   = "; ".join(existing + list(paths))
            var.set(new_val)
            if isinstance(widget, ttk.Combobox):
                widget.set(new_val)
            setattr(self, "_attachment_raw_json", None)
            inc = info.get("include_var")
            if inc is not None:
                inc.set(True)
        except Exception as e:
            debug_log(f"Attach failed: {e}")
            messagebox.showerror("Attach", f"Failed to add files: {e}")

    # ══════════════════════════════════════════════════════════════════════════
    # ADF ↔ plain-text (kept for completeness; not used in normal flow)
    # ══════════════════════════════════════════════════════════════════════════

    def _adf_to_display_text(self, node):
        """Render ADF node to a markdown-ish plain text (round-trippable)."""
        if not node or not isinstance(node, dict):
            return ""
        lines = []

        def _txt(n):
            if isinstance(n, dict):
                if n.get("type") == "text":
                    return n.get("text", "")
                return "".join(_txt(c) for c in n.get("content", []))
            return ""

        def walk(n):
            if not isinstance(n, dict):
                return
            t = n.get("type")
            if t == "doc":
                for c in n.get("content", []):
                    walk(c)
            elif t == "paragraph":
                lines.append("".join(_txt(c) for c in n.get("content", [])))
            elif t == "heading":
                lvl = min(6, max(1, int((n.get("attrs") or {}).get("level", 1))))
                lines.append("#" * lvl + " " + "".join(_txt(c) for c in n.get("content", [])))
            elif t == "bulletList":
                for item in n.get("content", []):
                    for sub in item.get("content", []):
                        lines.append("• " + "".join(_txt(c) for c in sub.get("content", [])))
            elif t == "orderedList":
                for i, item in enumerate(n.get("content", []), 1):
                    for sub in item.get("content", []):
                        lines.append(f"{i}. " + "".join(_txt(c) for c in sub.get("content", [])))
            elif t in ("taskList", "actionList"):
                for item in n.get("content", []):
                    state = (item.get("attrs") or {}).get("state", "TODO")
                    chk   = "☑" if state == "DONE" else "☐"
                    lines.append(f"{chk} " + "".join(_txt(c) for c in item.get("content", [])))
            elif t == "table":
                rows = n.get("content", [])
                if not rows:
                    return
                grid, header_row = [], set()
                for ri, row in enumerate(rows):
                    cells, is_hdr = [], False
                    for cell in (row.get("content") or []):
                        if cell.get("type") == "tableHeader":
                            is_hdr = True
                        txt = ""
                        for para in (cell.get("content") or []):
                            txt += "".join(_txt(c) for c in para.get("content", []))
                        cells.append(txt)
                    grid.append(cells)
                    if is_hdr:
                        header_row.add(ri)
                num_cols = max(len(r) for r in grid) if grid else 1
                widths = [max(3, max((len(grid[ri][ci]) if ci < len(grid[ri]) else 0)
                              for ri in range(len(grid)))) for ci in range(num_cols)]

                def _row_line(cells):
                    padded = [(cells[ci] if ci < len(cells) else "").ljust(widths[ci])
                               for ci in range(num_cols)]
                    return "| " + " | ".join(padded) + " |"

                for ri, cells in enumerate(grid):
                    lines.append(_row_line(cells))
                    if ri == 0:
                        lines.append("|" + "|".join("-" * (w + 2) for w in widths) + "|")
                lines.append("")
            elif t == "codeBlock":
                code = "".join(c.get("text", "") for c in n.get("content", []) if c.get("type") == "text")
                lines.extend(["```", *code.split("\n"), "```"])
            elif t == "blockquote":
                before = len(lines)
                for c in n.get("content", []):
                    walk(c)
                for i in range(before, len(lines)):
                    lines[i] = "> " + lines[i]
            elif t == "rule":
                lines.append("---")
            elif t == "hardBreak":
                lines.append("")
            else:
                for c in n.get("content", []):
                    walk(c)

        walk(node)
        return "\n".join(lines)

    def _display_text_to_adf(self, text):
        """Parse the markdown-ish plain text produced by _adf_to_display_text back to ADF."""
        if not text or not text.strip():
            return {"type": "doc", "version": 1, "content": [{"type": "paragraph", "content": []}]}

        def _para(t):
            return {"type": "paragraph", "content": [{"type": "text", "text": t}] if t else []}

        def _is_table_row(line):
            s = line.strip()
            return s.startswith("|") and s.endswith("|") and len(s) > 2

        def _is_separator(line):
            return bool(re.match(r'^\|[-| ]+\|$', line.strip()))

        def _parse_cells(line):
            s = line.strip().strip("|")
            return [c.strip() for c in s.split("|")]

        content = []
        bullet_buf, ordered_buf, task_buf = [], [], []
        in_code, code_buf = False, []
        table_rows, in_table = [], False

        def flush_lists():
            nonlocal bullet_buf, ordered_buf, task_buf
            if bullet_buf:
                content.append({"type": "bulletList", "content": [
                    {"type": "listItem", "content": [_para(s)]} for s in bullet_buf]})
                bullet_buf.clear()
            if ordered_buf:
                content.append({"type": "orderedList", "content": [
                    {"type": "listItem", "content": [_para(s)]} for s in ordered_buf]})
                ordered_buf.clear()
            if task_buf:
                items = [{"type": "taskItem",
                           "attrs": {"localId": str(uuid.uuid4()), "state": st},
                           "content": [{"type": "text", "text": tx}] if tx else []}
                          for st, tx in task_buf]
                content.append({"type": "taskList",
                                  "attrs": {"localId": str(uuid.uuid4())},
                                  "content": items})
                task_buf.clear()

        def flush_table():
            nonlocal table_rows, in_table
            if table_rows:
                adf_rows = []
                for cells, is_hdr in table_rows:
                    cell_type = "tableHeader" if is_hdr else "tableCell"
                    adf_rows.append({"type": "tableRow", "content": [
                        {"type": cell_type, "attrs": {}, "content": [_para(ct)]}
                        for ct in cells
                    ]})
                content.append({"type": "table",
                                  "attrs": {"isNumberColumnEnabled": False, "layout": "default"},
                                  "content": adf_rows})
            table_rows.clear()
            in_table = False

        lines = text.split("\n")
        i = 0
        while i < len(lines):
            line = lines[i]
            if line.strip() == "```":
                flush_lists(); flush_table()
                if in_code:
                    content.append({"type": "codeBlock",
                                    "content": [{"type": "text", "text": "\n".join(code_buf)}]})
                    code_buf.clear(); in_code = False
                else:
                    in_code = True
                i += 1; continue
            if in_code:
                code_buf.append(line); i += 1; continue
            if _is_table_row(line):
                flush_lists()
                if not _is_separator(line):
                    cells  = _parse_cells(line)
                    is_hdr = len(table_rows) == 0
                    table_rows.append((cells, is_hdr))
                    in_table = True
                i += 1; continue
            elif in_table:
                flush_table()
            if line.strip() == "---":
                flush_lists(); content.append({"type": "rule"}); i += 1; continue
            m = re.match(r'^(#{1,6})\s+(.*)', line)
            if m:
                flush_lists()
                content.append({"type": "heading", "attrs": {"level": len(m.group(1))},
                                  "content": [{"type": "text", "text": m.group(2)}] if m.group(2) else []})
                i += 1; continue
            if line.startswith("• "):
                if ordered_buf or task_buf:
                    flush_lists()
                bullet_buf.append(line[2:]); i += 1; continue
            m = re.match(r'^\d+\.\s+(.*)', line)
            if m:
                if bullet_buf or task_buf:
                    flush_lists()
                ordered_buf.append(m.group(1)); i += 1; continue
            m = re.match(r'^([☐☑])\s+(.*)', line)
            if m:
                if bullet_buf or ordered_buf:
                    flush_lists()
                task_buf.append(("DONE" if m.group(1) == "☑" else "TODO", m.group(2)))
                i += 1; continue
            if line.startswith("> "):
                flush_lists()
                content.append({"type": "blockquote", "content": [_para(line[2:])]}); i += 1; continue
            flush_lists()
            content.append(_para(line))
            i += 1

        flush_lists()
        flush_table()
        return {"type": "doc", "version": 1, "content": content or [_para("")]}

    # ══════════════════════════════════════════════════════════════════════════
    # Legacy stubs (called by older code paths; kept for compatibility)
    # ══════════════════════════════════════════════════════════════════════════

    def _set_preview_html(self, html_doc, fallback_text=""):
        """Legacy: no-op in the HtmlFrame architecture."""
        pass

    def _toggle_desc_edit_mode(self):
        """Legacy: opens the WYSIWYG editor."""
        self._open_wysiwyg_editor()

    def _enter_desc_edit_mode(self):
        """Legacy no-op."""
        pass

    def _exit_desc_edit_mode(self):
        """Legacy no-op."""
        pass
