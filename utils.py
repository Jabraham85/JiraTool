"""
Utility helpers: mouse wheel binding, close-tab notebook, debug logging, list deduplication.
"""
import os
import datetime
import tkinter as tk
from tkinter import ttk

from .config import DEBUG_LOG


# ---------------- Mouse wheel scroll helper ----------------
def _bind_mousewheel(widget, orient="vertical"):
    """Bind mouse wheel to scroll widget. orient: 'vertical' or 'horizontal'."""
    def _scroll(event):
        try:
            if orient == "vertical":
                if hasattr(event, "delta"):
                    widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    widget.yview_scroll(-3, "units")
                elif event.num == 5:
                    widget.yview_scroll(3, "units")
            else:
                if hasattr(event, "delta"):
                    widget.xview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    widget.xview_scroll(-3, "units")
                elif event.num == 5:
                    widget.xview_scroll(3, "units")
        except Exception:
            pass
    widget.bind("<MouseWheel>", _scroll)
    widget.bind("<Button-4>", _scroll)
    widget.bind("<Button-5>", _scroll)

def _bind_mousewheel_to_target(widget, target, orient="vertical"):
    """Bind mouse wheel on widget to scroll target (e.g. when hovering over content)."""
    def _scroll(event):
        try:
            if orient == "vertical":
                if hasattr(event, "delta"):
                    target.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    target.yview_scroll(-3, "units")
                elif event.num == 5:
                    target.yview_scroll(3, "units")
            else:
                if hasattr(event, "delta"):
                    target.xview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    target.xview_scroll(-3, "units")
                elif event.num == 5:
                    target.xview_scroll(3, "units")
        except Exception:
            pass
    widget.bind("<MouseWheel>", _scroll)
    widget.bind("<Button-4>", _scroll)
    widget.bind("<Button-5>", _scroll)


def _bind_combobox_scroll_passthrough(combo, target, orient="vertical"):
    """Prevent a Combobox from changing value on scroll when the dropdown is
    closed.  If the dropdown IS open, let the listbox scroll normally.
    Otherwise forward the wheel event to *target* (the main scroll canvas)."""
    def _scroll(event):
        try:
            # Check if the Combobox popdown listbox is visible.
            # When the dropdown is open, Tk creates a toplevel with a listbox;
            # we can detect it via the popdown widget path.
            popdown_visible = False
            try:
                popdown = combo.tk.call("ttk::combobox::PopdownWindow", combo)
                if popdown:
                    popdown_visible = bool(combo.tk.call("wm", "state", popdown) == "normal")
            except Exception:
                pass

            if popdown_visible:
                return  # let the dropdown list scroll naturally

            # Dropdown is closed — forward scroll to main canvas
            if orient == "vertical":
                if hasattr(event, "delta"):
                    target.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    target.yview_scroll(-3, "units")
                elif event.num == 5:
                    target.yview_scroll(3, "units")
            else:
                if hasattr(event, "delta"):
                    target.xview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    target.xview_scroll(-3, "units")
                elif event.num == 5:
                    target.xview_scroll(3, "units")
            return "break"  # prevent Combobox from changing its value
        except Exception:
            pass
    combo.bind("<MouseWheel>", _scroll, add=False)
    combo.bind("<Button-4>", _scroll, add=False)
    combo.bind("<Button-5>", _scroll, add=False)


def _bind_mousewheel_to_target_recursive(widget, target, orient="vertical"):
    """Bind mouse wheel on widget and descendants to scroll target;
    skip Text widgets (they scroll themselves).
    Comboboxes get special handling so scrolling doesn't change their value."""
    if isinstance(widget, tk.Text):
        return
    if isinstance(widget, ttk.Combobox):
        _bind_combobox_scroll_passthrough(widget, target, orient)
    else:
        _bind_mousewheel_to_target(widget, target, orient)
    for child in widget.winfo_children():
        _bind_mousewheel_to_target_recursive(child, target, orient)


# ---------------- Custom Notebook with close buttons on tabs ----------------
class _NotebookWithCloseTabs(ttk.Notebook):
    """ttk.Notebook with a close button on each tab (ticket tabs only; Welcome has no close)."""

    _style_initialized = False

    def __init__(self, parent, on_tab_close=None, **kwargs):
        if not _NotebookWithCloseTabs._style_initialized:
            root = parent.winfo_toplevel() if parent else None
            _NotebookWithCloseTabs._init_close_style(root)
            _NotebookWithCloseTabs._style_initialized = True
        kwargs["style"] = "NotebookWithClose"
        super().__init__(parent, **kwargs)
        self._on_tab_close = on_tab_close
        self._active_close_index = None
        self.bind("<ButtonPress-1>", self._on_close_press, True)
        self.bind("<ButtonRelease-1>", self._on_close_release)

    def _on_close_press(self, event):
        ident = self.identify(event.x, event.y)
        if "close" in str(ident):
            idx = self.index("@%d,%d" % (event.x, event.y))
            self._active_close_index = idx
            return "break"
        return None

    def _on_close_release(self, event):
        if self._active_close_index is None:
            return
        ident = self.identify(event.x, event.y)
        if "close" not in str(ident):
            self._active_close_index = None
            return
        idx = self.index("@%d,%d" % (event.x, event.y))
        if idx == self._active_close_index:
            tab_id = self.tabs()[idx]
            widget = self.nametowidget(tab_id)
            if self._on_tab_close:
                self._on_tab_close(widget)
            else:
                self.forget(idx)
            self.event_generate("<<NotebookTabClosed>>")
        self._active_close_index = None

    @classmethod
    def _init_close_style(cls, master=None):
        style = ttk.Style()
        root = master or getattr(tk, "_default_root", None)
        img_data = (
            "R0lGODlhCAAIAMIBAAAAADs7O4+Pj9nZ2Ts7Ozs7Ozs7Ozs7OyH+EUNyZWF0ZWQg"
            "d2l0aCBHSU1QACH5BAEKAAQALAAAAAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU"
            "5kEJADs="
        )
        try:
            if root:
                tk.PhotoImage(name="img_close", data=img_data, master=root)
                tk.PhotoImage(name="img_closeactive", data=img_data, master=root)
                tk.PhotoImage(name="img_closepressed", data=img_data, master=root)
        except Exception:
            pass
        try:
            style.element_create(
                "NotebookWithClose.close", "image", "img_close",
                ("active", "pressed", "!disabled", "img_closepressed"),
                ("active", "!disabled", "img_closeactive"),
                border=8, sticky=""
            )
        except tk.TclError:
            pass
        cls._apply_close_layouts(style)

    @classmethod
    def _apply_close_layouts(cls, style=None):
        """Apply tab layouts (call after theme_use to restore close buttons)."""
        if style is None:
            style = ttk.Style()
        root = getattr(tk, "_default_root", None)
        img_data = (
            "R0lGODlhCAAIAMIBAAAAADs7O4+Pj9nZ2Ts7Ozs7Ozs7OyH+EUNyZWF0ZWQg"
            "d2l0aCBHSU1QACH5BAEKAAQALAAAAAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU"
            "5kEJADs="
        )
        try:
            if root:
                tk.PhotoImage(name="img_close", data=img_data, master=root)
                tk.PhotoImage(name="img_closeactive", data=img_data, master=root)
                tk.PhotoImage(name="img_closepressed", data=img_data, master=root)
        except Exception:
            pass
        try:
            style.element_create(
                "NotebookWithClose.close", "image", "img_close",
                ("active", "pressed", "!disabled", "img_closepressed"),
                ("active", "!disabled", "img_closeactive"),
                border=8, sticky=""
            )
        except tk.TclError:
            pass
        try:
            style.layout("NotebookWithClose", [("Notebook.client", {"sticky": "nswe"})])
            style.layout("NotebookWithClose.Tab", [
                ("Notebook.tab", {
                    "sticky": "nswe",
                    "children": [
                        ("Notebook.padding", {
                            "side": "top",
                            "sticky": "nswe",
                            "children": [
                                ("Notebook.focus", {
                                    "side": "top",
                                    "sticky": "nswe",
                                    "children": [
                                        ("Notebook.label", {"side": "left", "sticky": ""}),
                                        ("NotebookWithClose.close", {"side": "left", "sticky": ""}),
                                    ]
                                })
                            ]
                        })
                    ]
                })
            ])
            style.layout("NotebookNoClose.Tab", [
                ("Notebook.tab", {
                    "sticky": "nswe",
                    "children": [
                        ("Notebook.padding", {
                            "side": "top",
                            "sticky": "nswe",
                            "children": [
                                ("Notebook.focus", {
                                    "side": "top",
                                    "sticky": "nswe",
                                    "children": [
                                        ("Notebook.label", {"side": "left", "sticky": ""}),
                                    ]
                                })
                            ]
                        })
                    ]
                })
            ])
        except tk.TclError:
            pass


# ---------------- Debug helpers ----------------
def debug_log(msg: str):
    ts = datetime.datetime.now().isoformat(sep=" ", timespec="seconds")
    try:
        with open(DEBUG_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass

def _dedup_list_items(items):
    """Remove duplicate tickets from list_items, keeping last occurrence (most recent data).
    Deduplicates by Issue key (preferred), then Issue id."""
    seen = {}
    for item in items:
        key = str(item.get("Issue key") or "").strip()
        iid = str(item.get("Issue id") or "").strip()
        dedup_key = key if key and not key.startswith("LOCAL-") else iid
        if dedup_key:
            seen[dedup_key] = item
        else:
            seen[id(item)] = item
    return list(seen.values())
