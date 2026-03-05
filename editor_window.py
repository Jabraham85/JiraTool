"""
Standalone WYSIWYG description editor using pywebview + Edge WebView2.

Called as a subprocess by tab_form.py:
    python -m avalanche.editor_window <json_args_file>

json_args_file contains {"html": "<initial html>", "output": "<output path>"}
Result HTML is written to the output path on save; nothing written on cancel.
"""

import sys
import os
import json
import base64
import tempfile

# ── Embedded editor HTML ──────────────────────────────────────────────────────
_EDITOR_HTML = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Description Editor</title>
<style>
  *, *::before, *::after { box-sizing: border-box; }

  body {
    margin: 0; padding: 0;
    background: #1e1e1e; color: #dcdcdc;
    font-family: 'Segoe UI', Roboto, sans-serif;
    font-size: 14px; line-height: 1.6;
    display: flex; flex-direction: column; height: 100vh; overflow: hidden;
  }

  /* ── Toolbar ── */
  #toolbar {
    background: #2d2d2d;
    border-bottom: 1px solid #444;
    padding: 5px 10px;
    display: flex; align-items: center; flex-wrap: wrap; gap: 3px;
    flex-shrink: 0;
  }
  .tb-sep { width: 1px; height: 22px; background: #555; margin: 0 4px; }
  .tb-btn {
    background: transparent; border: 1px solid transparent;
    color: #dcdcdc; cursor: pointer;
    padding: 3px 8px; border-radius: 3px;
    font-size: 13px; line-height: 1.4;
    transition: background 0.1s, border-color 0.1s;
    white-space: nowrap;
  }
  .tb-btn:hover { background: #3c3c3c; border-color: #555; }
  .tb-btn:active, .tb-btn.active { background: #264f78; border-color: #4a9eff; }
  .tb-btn b, .tb-btn i, .tb-btn u { color: inherit; }
  select.tb-sel {
    background: #3c3c3c; border: 1px solid #555; color: #dcdcdc;
    padding: 3px 6px; border-radius: 3px; font-size: 13px; cursor: pointer;
  }
  select.tb-sel:focus { outline: none; border-color: #4a9eff; }

  /* ── Editor area ── */
  #editor-wrap {
    flex: 1; overflow-y: auto; padding: 16px 24px;
  }
  #editor {
    min-height: 100%;
    outline: none;
    caret-color: #ffffff;
  }
  #editor:empty::before {
    content: 'Start typing here...';
    color: #555; pointer-events: none;
  }

  /* ── Inline styles ── */
  #editor p   { margin: 5px 0; }
  #editor h1  { font-size: 22px; margin: 14px 0 5px; color: #fff; }
  #editor h2  { font-size: 18px; margin: 12px 0 4px; color: #fff; }
  #editor h3  { font-size: 15px; margin: 10px 0 4px; color: #fff; }
  #editor h4,
  #editor h5,
  #editor h6  { margin: 8px 0 4px; color: #fff; }
  #editor a   { color: #4da6ff; text-decoration: none; }
  #editor a:hover { text-decoration: underline; }
  #editor code {
    background: #2d2d2d; color: #ce9178;
    padding: 1px 4px; border-radius: 3px;
    font-family: Consolas, monospace; font-size: 12px;
  }
  #editor pre {
    background: #2d2d2d; padding: 10px 14px;
    border-radius: 3px; margin: 8px 0; overflow-x: auto;
  }
  #editor pre code { background: none; padding: 0; }
  #editor blockquote {
    border-left: 3px solid #555; color: #aaa;
    margin: 8px 0; padding: 4px 12px;
  }
  #editor hr { border: none; border-top: 1px solid #444; margin: 12px 0; }
  #editor ul, #editor ol { padding-left: 22px; margin: 6px 0; }
  #editor li { margin: 3px 0; }

  /* ── Tables ── */
  #editor table {
    border-collapse: collapse;
    width: 100%;
    margin: 10px 0;
    font-size: 14px;
  }
  #editor th {
    background: #3d5a80; color: #e0e0e0;
    border: 1px solid #555;
    padding: 7px 10px;
    text-align: left; font-weight: 600;
    min-width: 80px;
  }
  #editor td {
    background: #252525; color: #dcdcdc;
    border: 1px solid #444;
    padding: 7px 10px;
    vertical-align: top;
    min-width: 60px; min-height: 32px;
  }
  #editor th:focus-within,
  #editor td:focus-within {
    outline: 2px solid #4a9eff;
    outline-offset: -2px;
  }
  /* selected cell highlight */
  #editor th.cell-selected,
  #editor td.cell-selected {
    background: #264f78 !important;
  }

  /* ── Footer ── */
  #footer {
    background: #2d2d2d;
    border-top: 1px solid #444;
    padding: 8px 16px;
    display: flex; align-items: center; justify-content: flex-end; gap: 8px;
    flex-shrink: 0;
  }
  #status { color: #888; font-size: 12px; margin-right: auto; }
  .foot-btn {
    padding: 6px 18px; border-radius: 3px;
    border: 1px solid #555; cursor: pointer;
    font-size: 13px; font-weight: 600;
    transition: background 0.15s;
  }
  #cancel-btn { background: #3c3c3c; color: #dcdcdc; }
  #cancel-btn:hover { background: #505050; }
  #save-btn { background: #0e639c; color: #fff; border-color: #0e639c; }
  #save-btn:hover { background: #1177bb; }

  /* Insert-table dialog */
  #table-dialog {
    display: none;
    position: fixed; top: 50%; left: 50%;
    transform: translate(-50%, -50%);
    background: #2d2d2d; border: 1px solid #555;
    padding: 20px 24px; border-radius: 6px;
    z-index: 999; box-shadow: 0 8px 32px rgba(0,0,0,0.6);
    min-width: 240px;
  }
  #table-dialog label { display: block; margin: 8px 0 4px; color: #bbb; font-size: 13px; }
  #table-dialog input[type=number] {
    width: 100%; padding: 5px 8px; background: #1e1e1e;
    border: 1px solid #555; border-radius: 3px;
    color: #dcdcdc; font-size: 13px;
  }
  #table-dialog .dlg-btns { display: flex; gap: 8px; margin-top: 16px; justify-content: flex-end; }
  #overlay { display: none; position: fixed; inset: 0; z-index: 998; }

  /* ── Inline images ── */
  #editor img.aval-img {
    max-width: 100%; height: auto;
    border: 1px solid #444; border-radius: 4px;
    margin: 8px 0; display: block;
    cursor: default;
  }
  #editor .aval-img-wrap {
    text-align: center; margin: 8px 0;
    position: relative;
  }
  #editor .aval-img-wrap .aval-img-name {
    font-size: 11px; color: #888; margin-top: 2px;
  }

  /* Custom context menu */
  .ctx-item {
    padding: 5px 16px; color: #dcdcdc; cursor: pointer; white-space: nowrap;
  }
  .ctx-item:hover { background: #264f78; }
  .ctx-sep { height: 1px; background: #444; margin: 4px 0; }
</style>
</head>
<body>

<div id="toolbar">
  <!-- Text format -->
  <button class="tb-btn" id="btn-bold"   title="Bold (Ctrl+B)"      onclick="exec('bold')"><b>B</b></button>
  <button class="tb-btn" id="btn-italic" title="Italic (Ctrl+I)"    onclick="exec('italic')"><i>I</i></button>
  <button class="tb-btn" id="btn-under"  title="Underline (Ctrl+U)" onclick="exec('underline')"><u>U</u></button>
  <button class="tb-btn" id="btn-strike" title="Strikethrough"      onclick="exec('strikeThrough')"><s>S</s></button>
  <button class="tb-btn" id="btn-code"   title="Inline code"        onclick="wrapCode()"><code style="font-size:11px">code</code></button>

  <div class="tb-sep"></div>

  <!-- Headings -->
  <select class="tb-sel" id="heading-sel" title="Block format" onchange="applyBlock(this.value)">
    <option value="p">Paragraph</option>
    <option value="h1">Heading 1</option>
    <option value="h2">Heading 2</option>
    <option value="h3">Heading 3</option>
    <option value="pre">Code block</option>
    <option value="blockquote">Quote</option>
  </select>

  <div class="tb-sep"></div>

  <!-- Lists -->
  <button class="tb-btn" title="Bullet list"   onclick="exec('insertUnorderedList')">• List</button>
  <button class="tb-btn" title="Numbered list" onclick="exec('insertOrderedList')">1. List</button>

  <div class="tb-sep"></div>

  <!-- Table ops -->
  <button class="tb-btn" title="Insert table"      onclick="openTableDialog()">⊞ Table</button>
  <button class="tb-btn" title="Add row below"     onclick="addRow()">+ Row</button>
  <button class="tb-btn" title="Delete current row" onclick="delRow()">− Row</button>
  <button class="tb-btn" title="Add column right"  onclick="addCol()">+ Col</button>
  <button class="tb-btn" title="Delete current column" onclick="delCol()">− Col</button>

  <div class="tb-sep"></div>

  <!-- Misc -->
  <button class="tb-btn" title="Horizontal rule" onclick="insertHR()">—</button>
  <button class="tb-btn" title="Link"            onclick="insertLink()">🔗</button>
  <button class="tb-btn" title="Insert image"    onclick="pickImage()">🖼 Image</button>

  <div class="tb-sep" id="vars-sep" style="display:none"></div>

  <!-- Variables (populated dynamically if vars are available) -->
  <select class="tb-sel" id="vars-sel" title="Insert variable reference" style="display:none"
          onchange="insertVarRef(this)">
    <option value="">⬡ Variables</option>
  </select>
</div>

<!-- Custom context menu -->
<div id="ctx-menu" style="display:none;position:fixed;z-index:9999;background:#2d2d2d;
     border:1px solid #555;border-radius:4px;padding:4px 0;min-width:180px;
     box-shadow:0 4px 16px rgba(0,0,0,0.5);font-size:13px;">
  <div class="ctx-item" onclick="ctxExec('cut')">Cut</div>
  <div class="ctx-item" onclick="ctxExec('copy')">Copy</div>
  <div class="ctx-item" onclick="ctxExec('paste')">Paste</div>
  <div class="ctx-sep"></div>
  <div class="ctx-item" onclick="ctxExec('selectAll')">Select All</div>
  <div class="ctx-sep"></div>
  <div class="ctx-item" onclick="ctxExec('bold')"><b>Bold</b></div>
  <div class="ctx-item" onclick="ctxExec('italic')"><i>Italic</i></div>
  <!-- Variable section — populated dynamically on each right-click -->
  <div id="ctx-var-section"></div>
</div>

<div id="editor-wrap">
  <div id="editor" contenteditable="true" spellcheck="true"></div>
</div>

<div id="footer">
  <span id="status">Tip: Tab moves between cells · Ctrl+Enter to save</span>
  <button class="foot-btn" id="cancel-btn" onclick="doCancel()">Cancel</button>
  <button class="foot-btn" id="save-btn"   onclick="doSave()">✓ Done</button>
</div>

<!-- Insert-table dialog -->
<div id="overlay"  onclick="closeTableDialog()"></div>
<div id="table-dialog">
  <div style="font-weight:600;margin-bottom:12px;color:#dcdcdc">Insert Table</div>
  <label>Rows</label>
  <input type="number" id="tbl-rows" value="3" min="1" max="20">
  <label>Columns</label>
  <input type="number" id="tbl-cols" value="3" min="1" max="10">
  <div class="dlg-btns">
    <button class="foot-btn" style="padding:5px 14px;" onclick="closeTableDialog()">Cancel</button>
    <button class="foot-btn" style="padding:5px 14px;background:#0e639c;color:#fff;border-color:#0e639c;" onclick="confirmInsertTable()">Insert</button>
  </div>
</div>

<script>
"use strict";

var ed = document.getElementById('editor');
var _ready = false;
var _loadedVars = {};   // vars passed in from Python {KEY: value}

// ── Init: load content from Python ───────────────────────────────────────────
window.addEventListener('pywebviewready', function () {
  _ready = true;
  window.pywebview.api.get_content().then(function (html) {
    ed.innerHTML = cleanIncoming(html);
    if (!ed.innerHTML.trim()) {
      ed.innerHTML = '<p><br></p>';
    }
    ensureTrailingParagraph();
    placeCursorAt(ed, true);
    ed.focus();
  });
  // Load variables for toolbar dropdown + right-click menu
  if (window.pywebview.api.get_vars) {
    window.pywebview.api.get_vars().then(function (vars) {
      if (!vars || typeof vars !== 'object') return;
      _loadedVars = vars;
      var keys = Object.keys(vars).sort();
      if (!keys.length) return;
      var sel = document.getElementById('vars-sel');
      var sep = document.getElementById('vars-sep');
      keys.forEach(function (k) {
        var opt = document.createElement('option');
        opt.value = k;
        opt.textContent = '{' + k + '} = ' + vars[k];
        sel.appendChild(opt);
      });
      sel.style.display = '';
      sep.style.display = '';
    });
  }
});

// ── Live variable sync: poll for new vars from the main window ────────────────
setInterval(function () {
  if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.get_vars) return;
  window.pywebview.api.get_vars().then(function (vars) {
    if (!vars || typeof vars !== 'object') return;
    Object.keys(vars).forEach(function (k) {
      if (!_loadedVars[k] || _loadedVars[k] !== vars[k]) {
        _loadedVars[k] = vars[k];
        registerVarInDropdown(k, vars[k]);
      }
    });
  });
}, 2000);

// ── Variable helpers ──────────────────────────────────────────────────────────

/** Return a Set of variable keys already used in the document ({A=…} markers). */
function usedVarKeys() {
  var keys = new Set(Object.keys(_loadedVars));
  var re = /\{([A-Z])=/g, m;
  var text = ed.textContent || '';
  while ((m = re.exec(text)) !== null) keys.add(m[1]);
  return keys;
}

/** Return the next free letter A-Z, or null if all 26 are taken. */
function nextFreeKey() {
  var used = usedVarKeys();
  for (var c = 65; c <= 90; c++) {
    var k = String.fromCharCode(c);
    if (!used.has(k)) return k;
  }
  return null;
}

/** Register a key/value pair in the toolbar dropdown (idempotent). */
function registerVarInDropdown(key, value) {
  var sel = document.getElementById('vars-sel');
  var sep = document.getElementById('vars-sep');
  for (var i = 0; i < sel.options.length; i++) {
    if (sel.options[i].value === key) {
      sel.options[i].textContent = '{' + key + '} = ' + value;
      return;
    }
  }
  var opt = document.createElement('option');
  opt.value = key;
  opt.textContent = '{' + key + '} = ' + value;
  sel.appendChild(opt);
  sel.style.display = '';
  sep.style.display = '';
}

/** Define a variable from the current selection. */
function defineVariable() {
  ctxMenu.style.display = 'none';
  // Restore the selection that was active at right-click time
  var selText = '';
  if (_savedRange) {
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(_savedRange);
    selText = sel.toString().trim();
    _savedRange = null;
  }
  if (!selText) {
    setStatus('Select some text first, then right-click → Define Variable.');
    return;
  }
  var key = nextFreeKey();
  if (!key) {
    setStatus('All variable slots A-Z are used.');
    return;
  }
  var marker = '{' + key + '=' + selText + '}';
  ed.focus();
  document.execCommand('insertText', false, marker);
  _loadedVars[key] = selText;
  registerVarInDropdown(key, selText);
  setStatus('Variable {' + key + '} defined as "' + selText + '"');
}

/** Insert a {KEY} reference at cursor. */
function insertVarAtCursor(key) {
  ctxMenu.style.display = 'none';
  if (_savedRange) {
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(_savedRange);
    _savedRange = null;
  }
  ed.focus();
  document.execCommand('insertText', false, '{' + key + '}');
}

// ── Variables toolbar ─────────────────────────────────────────────────────────
function insertVarRef(sel) {
  var key = sel.value;
  sel.value = '';   // reset to placeholder
  if (!key) return;
  ed.focus();
  var ref = '{' + key + '}';
  document.execCommand('insertText', false, ref);
}

// ── Custom right-click context menu ──────────────────────────────────────────
var _savedRange = null;
var ctxMenu = document.getElementById('ctx-menu');

document.addEventListener('contextmenu', function (e) {
  e.preventDefault();
  // Save selection so subsequent actions act on the right spot
  var sel = window.getSelection();
  var selText = '';
  if (sel && sel.rangeCount) {
    _savedRange = sel.getRangeAt(0).cloneRange();
    selText = sel.toString().trim();
  }

  // ── Rebuild the variable section each time the menu opens ────────────────
  var varSection = document.getElementById('ctx-var-section');
  varSection.innerHTML = '';

  var frag = document.createDocumentFragment();

  // "Define Variable" — only useful when text is selected
  var sep1 = document.createElement('div');
  sep1.className = 'ctx-sep';
  frag.appendChild(sep1);

  var defItem = document.createElement('div');
  defItem.className = 'ctx-item';
  if (selText) {
    var preview = selText.length > 20 ? selText.slice(0, 20) + '…' : selText;
    defItem.textContent = '⬡ Define Variable: "' + preview + '"';
    defItem.onclick = defineVariable;
  } else {
    defItem.textContent = '⬡ Define Variable';
    defItem.style.color = '#666';
    defItem.title = 'Select text first';
  }
  frag.appendChild(defItem);

  // "Insert {KEY}" entries for every known variable
  var allKeys = Object.keys(_loadedVars);
  // Also include any {KEY=…} defined inline in the document this session
  var re = /\{([A-Z])=([^}]+)\}/g, m;
  var text = ed.textContent || '';
  var inlineVars = {};
  while ((m = re.exec(text)) !== null) {
    if (!_loadedVars[m[1]]) inlineVars[m[1]] = m[2];
  }
  Object.keys(inlineVars).forEach(function (k) {
    if (allKeys.indexOf(k) === -1) allKeys.push(k);
    _loadedVars[k] = _loadedVars[k] || inlineVars[k];
  });
  allKeys.sort();

  if (allKeys.length) {
    var sep2 = document.createElement('div');
    sep2.className = 'ctx-sep';
    frag.appendChild(sep2);

    allKeys.forEach(function (k) {
      var item = document.createElement('div');
      item.className = 'ctx-item';
      var val = _loadedVars[k] || '';
      var preview = val.length > 18 ? val.slice(0, 18) + '…' : val;
      item.textContent = 'Insert {' + k + '} — ' + preview;
      item.onclick = (function (key) { return function () { insertVarAtCursor(key); }; })(k);
      frag.appendChild(item);
    });
  }

  varSection.appendChild(frag);

  // ── Ticket link detection ────────────────────────────────────────────────
  // Walk up from the click target to find the nearest <a> element
  var linkEl = e.target;
  while (linkEl && linkEl !== document.body) {
    if (linkEl.tagName === 'A' && linkEl.href) break;
    linkEl = linkEl.parentElement;
  }
  var ticketSection = document.getElementById('ctx-ticket-section');
  if (!ticketSection) {
    ticketSection = document.createElement('div');
    ticketSection.id = 'ctx-ticket-section';
    ctxMenu.appendChild(ticketSection);
  }
  ticketSection.innerHTML = '';
  if (linkEl && linkEl.tagName === 'A' && linkEl.href) {
    var href = linkEl.href || '';
    // Match /browse/KEY or just a bare KEY pattern
    var keyMatch = href.match(/\/browse\/([A-Z][A-Z0-9]+-\d+)/) ||
                   linkEl.textContent.match(/\b([A-Z][A-Z0-9]+-\d+)\b/);
    if (keyMatch) {
      var ticketKey = keyMatch[1];
      var tickSep = document.createElement('div');
      tickSep.className = 'ctx-sep';
      ticketSection.appendChild(tickSep);

      var openAppItem = document.createElement('div');
      openAppItem.className = 'ctx-item';
      openAppItem.textContent = '⧉ Open ' + ticketKey + ' locally';
      openAppItem.onclick = (function (k) { return function () {
        ctxMenu.style.display = 'none';
        if (window.pywebview && window.pywebview.api && window.pywebview.api.open_ticket_link) {
          window.pywebview.api.open_ticket_link(k);
        }
      }; })(ticketKey);
      ticketSection.appendChild(openAppItem);

      var openJiraItem = document.createElement('div');
      openJiraItem.className = 'ctx-item';
      openJiraItem.textContent = '↗ Open ' + ticketKey + ' in Jira';
      openJiraItem.onclick = (function (k, h) { return function () {
        ctxMenu.style.display = 'none';
        if (window.pywebview && window.pywebview.api && window.pywebview.api.open_in_jira) {
          window.pywebview.api.open_in_jira(k);
        } else {
          window.open(h, '_blank');
        }
      }; })(ticketKey, href);
      ticketSection.appendChild(openJiraItem);
    }
  }

  ctxMenu.style.left = Math.min(e.clientX, window.innerWidth  - 210) + 'px';
  ctxMenu.style.top  = Math.min(e.clientY, window.innerHeight - 260) + 'px';
  ctxMenu.style.display = 'block';
});

document.addEventListener('mousedown', function (e) {
  if (!ctxMenu.contains(e.target)) ctxMenu.style.display = 'none';
});

document.addEventListener('keydown', function (e) {
  if (e.key === 'Escape') ctxMenu.style.display = 'none';
}, true);

function ctxExec(cmd) {
  ctxMenu.style.display = 'none';
  // Restore selection before executing
  if (_savedRange) {
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(_savedRange);
    _savedRange = null;
  }
  ed.focus();
  if (cmd === 'paste') {
    // Clipboard API paste (requires focus + user gesture — works in Edge)
    if (navigator.clipboard && navigator.clipboard.readText) {
      navigator.clipboard.readText().then(function (txt) {
        document.execCommand('insertText', false, txt);
      }).catch(function () {
        document.execCommand('paste');
      });
    } else {
      document.execCommand('paste');
    }
  } else {
    document.execCommand(cmd, false, null);
  }
  updateToolbarState();
}

// ── Trailing-paragraph guard ──────────────────────────────────────────────────
// Block elements that cannot receive a cursor directly. After any of these
// sits at the very end of the editor, typing is impossible unless there is a
// paragraph below. We auto-insert one whenever this situation arises.
var UNCURSORABLE = ['TABLE', 'HR', 'FIGURE', 'PRE', 'BLOCKQUOTE', 'OL', 'UL'];

function ensureTrailingParagraph() {
  var last = ed.lastElementChild;
  if (!last || UNCURSORABLE.indexOf(last.tagName) !== -1) {
    var p = document.createElement('p');
    p.innerHTML = '<br>';
    ed.appendChild(p);
    return p;
  }
  return null;  // already fine
}

// Re-check after every edit (debounced so it doesn't interrupt typing)
var _etpTimer = null;
ed.addEventListener('input', function () {
  clearTimeout(_etpTimer);
  _etpTimer = setTimeout(ensureTrailingParagraph, 200);
});

// Clicking below all editor content in the wrapper area → jump to last paragraph
document.getElementById('editor-wrap').addEventListener('mousedown', function (e) {
  if (e.target !== this && e.target !== ed) return;
  var last = ed.lastElementChild;
  if (!last) return;
  var rect = last.getBoundingClientRect();
  if (e.clientY > rect.bottom) {
    e.preventDefault();
    var target = ensureTrailingParagraph() || ed.lastElementChild;
    moveCursorTo(target);
    ed.focus();
  }
});

function cleanIncoming(html) {
  // Strip adf-end / style-only sentinels, normalise &nbsp; to spaces in cells
  var tmp = document.createElement('div');
  tmp.innerHTML = html || '';
  tmp.querySelectorAll('[class*="adf-end"], [class*="adf_end"]').forEach(function (el) { el.remove(); });
  // Remove div wrappers inside cells (ADF renderer wraps cell content in divs)
  tmp.querySelectorAll('td > div, th > div').forEach(function (wrapper) {
    while (wrapper.firstChild) wrapper.parentNode.insertBefore(wrapper.firstChild, wrapper);
    wrapper.remove();
  });
  return tmp.innerHTML;
}

// ── execCommand wrapper ───────────────────────────────────────────────────────
function exec(cmd, val) {
  ed.focus();
  document.execCommand(cmd, false, val !== undefined ? val : null);
  updateToolbarState();
}

function applyBlock(tag) {
  ed.focus();
  if (tag === 'pre') {
    document.execCommand('formatBlock', false, '<pre>');
  } else if (tag === 'blockquote') {
    document.execCommand('formatBlock', false, '<blockquote>');
  } else {
    document.execCommand('formatBlock', false, '<' + tag + '>');
  }
  document.getElementById('heading-sel').value = tag;
  updateToolbarState();
}

function wrapCode() {
  var sel = window.getSelection();
  if (!sel || !sel.rangeCount) return;
  var range = sel.getRangeAt(0);
  var text = range.toString();
  if (!text) return;
  range.deleteContents();
  var code = document.createElement('code');
  code.textContent = text;
  range.insertNode(code);
  sel.removeAllRanges();
}

// ── Toolbar state update ──────────────────────────────────────────────────────
function updateToolbarState() {
  try {
    var cmds = {bold:'btn-bold', italic:'btn-italic', underline:'btn-under', strikeThrough:'btn-strike'};
    Object.keys(cmds).forEach(function(cmd) {
      var el = document.getElementById(cmds[cmd]);
      if (el) {
        if (document.queryCommandState(cmd)) el.classList.add('active');
        else el.classList.remove('active');
      }
    });
    // Update heading dropdown
    var sel = document.getElementById('heading-sel');
    var block = document.queryCommandValue('formatBlock').toLowerCase().replace(/[<>]/g, '');
    var validBlocks = ['p','h1','h2','h3','pre','blockquote'];
    sel.value = validBlocks.includes(block) ? block : 'p';
  } catch(e) {}
}

ed.addEventListener('keyup', updateToolbarState);
ed.addEventListener('click', updateToolbarState);

// ── Table insert dialog ───────────────────────────────────────────────────────
function openTableDialog() {
  document.getElementById('overlay').style.display = 'block';
  document.getElementById('table-dialog').style.display = 'block';
  document.getElementById('tbl-rows').focus();
}
function closeTableDialog() {
  document.getElementById('overlay').style.display = 'none';
  document.getElementById('table-dialog').style.display = 'none';
  ed.focus();
}
function confirmInsertTable() {
  var rows = Math.max(1, parseInt(document.getElementById('tbl-rows').value) || 3);
  var cols = Math.max(1, parseInt(document.getElementById('tbl-cols').value) || 3);
  closeTableDialog();
  doInsertTable(rows, cols);
}

function doInsertTable(rows, cols) {
  var html = '<table><tbody>';
  // Header row
  html += '<tr>';
  for (var c = 0; c < cols; c++) html += '<th>Header ' + (c + 1) + '</th>';
  html += '</tr>';
  // Data rows
  for (var r = 1; r < rows; r++) {
    html += '<tr>';
    for (var c = 0; c < cols; c++) html += '<td><br></td>';
    html += '</tr>';
  }
  html += '</tbody></table><p><br></p>';
  exec('insertHTML', html);
  // Move cursor into first data cell
  var tables = ed.querySelectorAll('table');
  var last = tables[tables.length - 1];
  if (last) {
    var first = last.querySelector('td, th');
    if (first) moveCursorTo(first);
  }
}

// ── Table helpers ─────────────────────────────────────────────────────────────
function getCurrentCell() {
  var sel = window.getSelection();
  if (!sel || !sel.rangeCount) return null;
  var n = sel.anchorNode;
  while (n && n !== ed) {
    if (n.nodeType === 1 && (n.tagName === 'TD' || n.tagName === 'TH')) return n;
    n = n.parentNode;
  }
  return null;
}

function getCurrentTable() {
  var c = getCurrentCell();
  return c ? c.closest('table') : null;
}

function moveCursorTo(cell) {
  if (!cell) return;
  // Make sure cell has at least a <br>
  if (!cell.firstChild) cell.appendChild(document.createElement('br'));
  var range = document.createRange();
  range.selectNodeContents(cell);
  range.collapse(false);
  var sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  cell.scrollIntoView({ block: 'nearest' });
}

function addRow() {
  var tbl = getCurrentTable();
  if (!tbl) { setStatus('Click inside a table first'); return; }
  var cell = getCurrentCell();
  var refRow = cell ? cell.parentElement : tbl.querySelector('tr:last-child');
  var cols = refRow ? refRow.cells.length : 2;
  var row = document.createElement('tr');
  for (var i = 0; i < cols; i++) {
    var td = document.createElement('td');
    td.innerHTML = '<br>';
    row.appendChild(td);
  }
  if (refRow) refRow.insertAdjacentElement('afterend', row);
  else (tbl.querySelector('tbody') || tbl).appendChild(row);
  moveCursorTo(row.cells[0]);
  setStatus('Row added');
}

function delRow() {
  var cell = getCurrentCell();
  if (!cell) { setStatus('Click inside a table first'); return; }
  var row = cell.parentElement;
  var tbl = row.closest('table');
  var allRows = tbl.querySelectorAll('tr');
  if (allRows.length <= 1) { setStatus('Cannot delete the last row'); return; }
  // Move cursor to previous row if possible
  var prev = row.previousElementSibling;
  if (!prev) prev = row.nextElementSibling;
  row.remove();
  if (prev && prev.cells[0]) moveCursorTo(prev.cells[0]);
  setStatus('Row deleted');
}

function addCol() {
  var tbl = getCurrentTable();
  if (!tbl) { setStatus('Click inside a table first'); return; }
  var cell = getCurrentCell();
  var colIdx = cell ? Array.from(cell.parentElement.cells).indexOf(cell) + 1 : -1;
  var rows = tbl.querySelectorAll('tr');
  rows.forEach(function (row, ri) {
    var isHeader = ri === 0;
    var newCell = document.createElement(isHeader ? 'th' : 'td');
    newCell.innerHTML = isHeader ? 'Header' : '<br>';
    if (colIdx >= 0 && colIdx < row.cells.length) {
      row.cells[colIdx].insertAdjacentElement('afterend', newCell);
    } else {
      row.appendChild(newCell);
    }
  });
  setStatus('Column added');
}

function delCol() {
  var cell = getCurrentCell();
  if (!cell) { setStatus('Click inside a table first'); return; }
  var tbl = cell.closest('table');
  var colIdx = Array.from(cell.parentElement.cells).indexOf(cell);
  var rows = tbl.querySelectorAll('tr');
  if (rows.length > 0 && rows[0].cells.length <= 1) { setStatus('Cannot delete the last column'); return; }
  rows.forEach(function (row) {
    if (row.cells[colIdx]) row.cells[colIdx].remove();
  });
  setStatus('Column deleted');
}

// ── Tab navigation in tables ──────────────────────────────────────────────────
ed.addEventListener('keydown', function (e) {
  if (e.key === 'Tab') {
    var cell = getCurrentCell();
    if (cell) {
      e.preventDefault();
      var tbl = cell.closest('table');
      var allCells = Array.from(tbl.querySelectorAll('td, th'));
      var idx = allCells.indexOf(cell);
      if (e.shiftKey) {
        idx--;
      } else {
        idx++;
      }
      if (idx >= 0 && idx < allCells.length) {
        moveCursorTo(allCells[idx]);
      } else if (!e.shiftKey) {
        // Add new row after last row
        var lastRow = tbl.querySelector('tr:last-child');
        var cols = lastRow ? lastRow.cells.length : 1;
        var newRow = document.createElement('tr');
        for (var i = 0; i < cols; i++) {
          var td = document.createElement('td');
          td.innerHTML = '<br>';
          newRow.appendChild(td);
        }
        (tbl.querySelector('tbody') || tbl).appendChild(newRow);
        moveCursorTo(newRow.cells[0]);
      }
    }
  }
});

// ── Misc inserts ──────────────────────────────────────────────────────────────
function insertHR() {
  exec('insertHTML', '<hr>');
}

function insertLink() {
  var url = prompt('Enter URL:');
  if (!url || !url.trim()) return;
  var text = window.getSelection().toString().trim() || url.trim();
  exec('insertHTML', '<a href="' + escHtml(url.trim()) + '">' + escHtml(text) + '</a>');
}

function escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Status bar ────────────────────────────────────────────────────────────────
function setStatus(msg) {
  var el = document.getElementById('status');
  el.textContent = msg;
  setTimeout(function () {
    el.textContent = 'Tip: Tab moves between cells · Ctrl+Enter to save';
  }, 2500);
}

// ── Image support ─────────────────────────────────────────────────────────────
var _imgCounter = 0;

function insertImageElement(name, dataUrl) {
  var wrap = document.createElement('div');
  wrap.className = 'aval-img-wrap';
  wrap.contentEditable = 'false';
  var img = document.createElement('img');
  img.className = 'aval-img';
  img.src = dataUrl;
  img.setAttribute('data-avalanche-file', name);
  wrap.appendChild(img);
  var caption = document.createElement('div');
  caption.className = 'aval-img-name';
  caption.textContent = name;
  wrap.appendChild(caption);
  // Insert at cursor position
  var sel = window.getSelection();
  if (sel && sel.rangeCount) {
    var range = sel.getRangeAt(0);
    range.deleteContents();
    range.insertNode(wrap);
    // Move cursor after the image
    range.setStartAfter(wrap);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  } else {
    ed.appendChild(wrap);
  }
  ensureTrailingParagraph();
  setStatus('Image inserted: ' + name);
}

function pickImage() {
  if (!_ready) return;
  window.pywebview.api.pick_image().then(function (result) {
    if (!result) return;
    try {
      var data = typeof result === 'string' ? JSON.parse(result) : result;
      if (data.name && data.dataUrl) {
        ed.focus();
        insertImageElement(data.name, data.dataUrl);
      }
    } catch (e) {
      setStatus('Failed to load image');
    }
  });
}

// Paste handler — intercept clipboard images
ed.addEventListener('paste', function (e) {
  var items = (e.clipboardData || {}).items;
  if (!items) return;
  for (var i = 0; i < items.length; i++) {
    if (items[i].type.indexOf('image/') === 0) {
      e.preventDefault();
      var blob = items[i].getAsFile();
      if (!blob) return;
      _imgCounter++;
      var ext = blob.type.split('/')[1] || 'png';
      var name = 'pasted_image_' + _imgCounter + '.' + ext;
      var reader = new FileReader();
      reader.onload = function (evt) {
        insertImageElement(name, evt.target.result);
      };
      reader.readAsDataURL(blob);
      return;
    }
  }
});

// ── Save / Cancel ─────────────────────────────────────────────────────────────
function getCleanHTML() {
  // Clone, clean up browser artifacts before sending back
  var clone = ed.cloneNode(true);
  // Remove empty <br>-only paragraphs at the very end
  var children = Array.from(clone.children);
  while (children.length > 0) {
    var last = children[children.length - 1];
    if (last.tagName === 'P' && last.innerHTML.trim() === '<br>') {
      clone.removeChild(last);
      children.pop();
    } else break;
  }
  // Replace base64 image srcs with placeholder URLs and collect image data
  var imgs = clone.querySelectorAll('img[data-avalanche-file]');
  var pending = [];
  imgs.forEach(function (img) {
    var name = img.getAttribute('data-avalanche-file');
    var src = img.src || '';
    if (src.indexOf('data:') === 0) {
      pending.push({ name: name, dataUrl: src });
      img.src = 'avalanche-pending://' + name;
    }
    // Unwrap from .aval-img-wrap div for clean HTML output
    var wrap = img.closest('.aval-img-wrap');
    if (wrap) {
      var parent = wrap.parentNode;
      parent.insertBefore(img, wrap);
      wrap.remove();
    }
  });
  return { html: clone.innerHTML, pendingImages: pending };
}

function doSave() {
  var result = getCleanHTML();
  document.getElementById('save-btn').textContent = 'Saving…';
  document.getElementById('save-btn').disabled = true;
  if (_ready) {
    // Save pending images to disk first, then save the HTML
    var imagesJson = JSON.stringify(result.pendingImages || []);
    window.pywebview.api.save_pending_images(imagesJson).then(function () {
      window.pywebview.api.save(result.html);
    });
  } else {
    window.close();
  }
}

function doCancel() {
  if (_ready) {
    window.pywebview.api.cancel();
  } else {
    window.close();
  }
}

// ── Keyboard shortcuts ────────────────────────────────────────────────────────
document.addEventListener('keydown', function (e) {
  if (e.ctrlKey && e.key === 'Enter') { e.preventDefault(); doSave(); }
  if (e.key === 'Escape') { doCancel(); }
});
</script>
</body>
</html>
"""


def main():
    if len(sys.argv) < 2:
        sys.exit(1)

    args_file = sys.argv[1]
    try:
        with open(args_file, encoding="utf-8") as f:
            args = json.load(f)
    except Exception as e:
        sys.exit(f"Cannot read args file: {e}")

    initial_html = args.get("html", "")
    output_path  = args.get("output", "")
    images_dir   = args.get("images_dir", "")

    if images_dir and not os.path.isdir(images_dir):
        try:
            os.makedirs(images_dir, exist_ok=True)
        except Exception:
            pass

    import webview  # type: ignore

    result = {"html": None}

    class Api:
        def get_content(self):
            return initial_html

        def get_vars(self):
            """Return the latest variable definitions by re-reading the args file."""
            try:
                with open(args_file, encoding="utf-8") as f:
                    fresh = json.load(f)
                return fresh.get("vars", {}) or {}
            except Exception:
                return args.get("vars", {}) or {}

        def open_ticket_link(self, key):
            """Write a ticket key to a side-channel file for the main app."""
            try:
                link_path = args_file + ".open_ticket"
                with open(link_path, "w", encoding="utf-8") as f:
                    json.dump({"key": key}, f)
            except Exception:
                pass

        def open_in_jira(self, key):
            """Open the Jira browse URL for *key* in the default browser."""
            try:
                jira_base = args.get("jira_base") or ""
                if jira_base:
                    import webbrowser
                    webbrowser.open(f"{jira_base.rstrip('/')}/browse/{key}")
            except Exception:
                pass

        def pick_image(self):
            """Open a native file dialog and return the image as a base64 data URL."""
            try:
                file_types = ('Image Files (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp)',)
                paths = win.create_file_dialog(
                    webview.OPEN_DIALOG,
                    allow_multiple=False,
                    file_types=file_types,
                )
                if not paths:
                    return None
                path = paths[0] if isinstance(paths, (list, tuple)) else paths
                if not path or not os.path.isfile(path):
                    return None
                name = os.path.basename(path)
                with open(path, "rb") as f:
                    data = f.read()
                ext = os.path.splitext(name)[1].lower().lstrip(".")
                mime_map = {"jpg": "jpeg", "svg": "svg+xml"}
                mime = f"image/{mime_map.get(ext, ext)}"
                b64 = base64.b64encode(data).decode("ascii")
                return json.dumps({"name": name, "dataUrl": f"data:{mime};base64,{b64}"})
            except Exception:
                return None

        def save_pending_images(self, images_json):
            """Write base64 images from the editor to the images_dir on disk."""
            if not images_dir:
                return True
            try:
                items = json.loads(images_json) if isinstance(images_json, str) else images_json
                for item in (items or []):
                    name = item.get("name", "")
                    data_url = item.get("dataUrl", "")
                    if not name or not data_url:
                        continue
                    if ";base64," in data_url:
                        raw = data_url.split(";base64,", 1)[1]
                    else:
                        continue
                    dest = os.path.join(images_dir, name)
                    with open(dest, "wb") as f:
                        f.write(base64.b64decode(raw))
            except Exception:
                pass
            return True

        def save(self, html):
            result["html"] = html
            win.destroy()

        def cancel(self):
            win.destroy()

    api = Api()
    win = webview.create_window(
        "Edit Description",
        html=_EDITOR_HTML,
        js_api=api,
        width=980,
        height=720,
        min_size=(500, 400),
        text_select=True,
    )

    webview.start(debug=False)

    # After the window closes, write result if saved
    if result["html"] is not None and output_path:
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(result["html"])
        except Exception as e:
            sys.exit(f"Cannot write output: {e}")


if __name__ == "__main__":
    main()
