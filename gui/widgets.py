"""
gui/widgets.py — Reusable custom Tkinter / Matplotlib widgets.

Currently contains:

_AutoScaleToolbar
    A subclass of NavigationToolbar2Tk whose "Home" button restores the
    data bounds (with a small margin) rather than matplotlib's default
    full reset.  Left-click-only zoom is also enforced so that right-click
    can be reused for panning by the plotter.

Usage
-----
    from gui.widgets import AutoScaleToolbar

    toolbar = AutoScaleToolbar(
        canvas,
        toolbar_frame,
        get_bounds=plotter_tab._get_plot_data_bounds,
    )
"""

import tkinter as tk
from tkinter import ttk

from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk


class AutoScaleToolbar(NavigationToolbar2Tk):
    """Matplotlib navigation toolbar with smarter Home behaviour.

    Parameters
    ----------
    canvas:
        The ``FigureCanvasTkAgg`` instance.
    window:
        The Tkinter frame that hosts the toolbar.
    get_bounds:
        Optional callable ``() → (x_min, x_max, y_min, y_max) | None``.
        When supplied, pressing *Home* restores exactly these bounds
        (with a 5 % margin already baked in by the plotter).  If it
        returns ``None``, the default autoscale behaviour is used.
    """

    def __init__(self, canvas, window, *, get_bounds=None):
        self._get_bounds = get_bounds
        super().__init__(canvas, window)

    # ── Restrict zoom to left-click only ──────────────────────────────────────

    def press_zoom(self, event):
        if event.button != 1:
            return
        return super().press_zoom(event)

    def release_zoom(self, event):
        if event.button != 1:
            return
        return super().release_zoom(event)

    # ── Smart Home ────────────────────────────────────────────────────────────

    def home(self, *args):
        axes = self.canvas.figure.axes
        if not axes:
            return
        ax = axes[0]

        bounds = None
        if self._get_bounds is not None:
            try:
                bounds = self._get_bounds()
            except Exception:
                bounds = None

        if bounds is not None:
            x_min, x_max, y_min, y_max = bounds
            ax.set_xlim(x_min, x_max)
            ax.set_ylim(y_min, y_max)
            self.canvas.draw_idle()
            return

        # Fallback: standard autoscale with a small margin
        ax.relim()
        ax.autoscale_view(tight=True)
        ax.margins(x=0.05, y=0.05)
        self.canvas.draw_idle()


class ScrollableFrame(ttk.Frame):
    """A reusable frame with vertical and horizontal overflow handling."""

    def __init__(self, parent, *, min_width=0, fit_width=True, **kwargs):
        super().__init__(parent, **kwargs)
        self._min_width = int(min_width or 0)
        self._fit_width = bool(fit_width)

        self._canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self._vscroll = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._hscroll = ttk.Scrollbar(self, orient="horizontal", command=self._canvas.xview)
        self.content = ttk.Frame(self._canvas)
        self._window = self._canvas.create_window((0, 0), window=self.content, anchor="nw")

        self._canvas.configure(yscrollcommand=self._vscroll.set, xscrollcommand=self._hscroll.set)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        self._vscroll.grid(row=0, column=1, sticky="ns")
        self._hscroll.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.content.bind("<Configure>", self._on_content_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)
        self._bind_mousewheel(self._canvas)
        self._bind_mousewheel(self.content)

    def _on_content_configure(self, _event=None):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        if not self._fit_width:
            return
        width = max(int(event.width), self._min_width)
        self._canvas.itemconfigure(self._window, width=width)

    def _bind_mousewheel(self, widget):
        widget.bind("<Enter>", lambda _e: self._canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        widget.bind("<Leave>", lambda _e: self._canvas.unbind_all("<MouseWheel>"))

    def _on_mousewheel(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def show_guide_dialog(parent, title: str, sections, *, width=760, height=620):
    """Open a scrollable step-by-step guide dialog."""
    root = parent.winfo_toplevel()
    win = tk.Toplevel(root)
    win.title(title)
    win.transient(root)
    win.resizable(True, True)
    win.minsize(620, 460)

    container = ttk.Frame(win, padding=14)
    container.pack(fill="both", expand=True)
    container.rowconfigure(1, weight=1)
    container.columnconfigure(0, weight=1)

    header = ttk.Frame(container)
    header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    header.columnconfigure(0, weight=1)
    ttk.Label(
        header,
        text=title,
        font=("Arial", 15, "bold"),
    ).pack(side="left")
    ttk.Label(
        header,
        text="Help only — this window does not change settings or start a run.",
        foreground="#555555",
        font=("Arial", 10),
    ).pack(side="right")

    text_frame = ttk.Frame(container)
    text_frame.grid(row=1, column=0, sticky="nsew")
    text_frame.rowconfigure(0, weight=1)
    text_frame.columnconfigure(0, weight=1)

    guide_text = tk.Text(
        text_frame,
        wrap="word",
        font=("Arial", 11),
        padx=14,
        pady=12,
        height=24,
        spacing1=1,
        spacing3=4,
    )
    guide_text.grid(row=0, column=0, sticky="nsew")
    yscroll = ttk.Scrollbar(text_frame, orient="vertical", command=guide_text.yview)
    yscroll.grid(row=0, column=1, sticky="ns")
    guide_text.configure(yscrollcommand=yscroll.set)
    guide_text.tag_configure("section", font=("Arial", 13, "bold"), spacing1=12, spacing3=6)
    guide_text.tag_configure("body", font=("Arial", 11), spacing3=4)
    guide_text.tag_configure("step", font=("Arial", 11), lmargin1=8, lmargin2=22, spacing3=4)

    for section_title, lines in sections:
        guide_text.insert("end", f"{section_title}\n", "section")
        for line in lines:
            tag = "step" if str(line).lstrip()[:1].isdigit() else "body"
            guide_text.insert("end", f"{line}\n", tag)
        guide_text.insert("end", "\n")
    guide_text.configure(state="disabled")

    ttk.Button(container, text="Close", command=win.destroy).grid(
        row=2, column=0, sticky="e", pady=(10, 0)
    )
    _center_child_window(root, win, width=width, height=height)
    return win


def _center_child_window(root, win, *, width=760, height=620):
    try:
        root.update_idletasks()
        root_x = root.winfo_rootx()
        root_y = root.winfo_rooty()
        root_w = max(1, root.winfo_width())
        root_h = max(1, root.winfo_height())
        x = root_x + max(0, (root_w - width) // 2)
        y = root_y + max(0, (root_h - height) // 2)
        win.geometry(f"{width}x{height}+{x}+{y}")
    except Exception:
        win.geometry(f"{width}x{height}")


class InfoButton(tk.Canvas):
    """A compact circular info icon that opens help when clicked."""

    def __init__(self, parent, *, size=20, command=None, **kwargs):
        super().__init__(
            parent,
            width=size,
            height=size,
            highlightthickness=0,
            borderwidth=0,
            cursor="hand2" if command is not None else "question_arrow",
            **kwargs,
        )
        self._size = int(size)
        self._command = command
        self._fill = "#eef6ff"
        self._outline = "#2f6f9f"
        self.configure(background=self._background(parent))
        self._draw()
        self.bind("<Enter>", lambda _e: self._redraw(fill="#dceeff"), add="+")
        self.bind("<Leave>", lambda _e: self._redraw(fill="#eef6ff"), add="+")
        self.bind("<ButtonRelease-1>", self._on_click, add="+")

    def _draw(self):
        self.delete("all")
        pad = 2
        self.create_oval(
            pad,
            pad,
            self._size - pad,
            self._size - pad,
            fill=self._fill,
            outline=self._outline,
            width=1,
        )
        self.create_text(
            self._size // 2,
            self._size // 2,
            text="i",
            fill="#2f6f9f",
            font=("TkDefaultFont", 9, "bold"),
        )

    def _redraw(self, *, fill):
        self._fill = fill
        self._draw()

    def _on_click(self, _event=None):
        if self._command is not None:
            self._command()

    @staticmethod
    def _background(parent):
        try:
            return parent.cget("background")
        except Exception:
            return "#f0f0f0"


class FlowFrame(ttk.Frame):
    """A simple wrapping container for rows of buttons and compact controls."""

    def __init__(self, parent, *, hgap=4, vgap=4, **kwargs):
        super().__init__(parent, **kwargs)
        self._hgap = hgap
        self._vgap = vgap
        self.bind("<Configure>", lambda _e: self._layout())

    def add(self, widget):
        widget.place(in_=self)
        widget.bind("<Configure>", lambda _e: self._layout())
        self.after_idle(self._layout)
        return widget

    def separator(self):
        sep = ttk.Separator(self, orient="vertical")
        return self.add(sep)

    def _layout(self):
        width = max(1, self.winfo_width())
        x = 0
        y = 0
        row_h = 0
        for child in self.winfo_children():
            req_w = child.winfo_reqwidth()
            req_h = child.winfo_reqheight()
            if x and x + req_w > width:
                x = 0
                y += row_h + self._vgap
                row_h = 0
            child.place(x=x, y=y, width=req_w, height=req_h)
            x += req_w + self._hgap
            row_h = max(row_h, req_h)
        self.configure(height=y + row_h)


def attach_info_button(
    parent,
    title,
    sections,
    *,
    side="right",
    padx=4,
    pady=2,
    size=18,
):
    """Pack a compact click-only i that opens the guide dialog."""

    def _open():
        show_guide_dialog(parent, title, sections)

    btn = InfoButton(parent, size=size, command=_open)
    btn.pack(side=side, padx=padx, pady=pady)
    return btn


def grid_info_button(
    parent,
    title,
    sections,
    *,
    row=0,
    column=99,
    sticky="ne",
    padx=2,
    pady=2,
    size=18,
):
    """Grid a compact click-only i that opens the guide dialog."""

    def _open():
        show_guide_dialog(parent, title, sections)

    btn = InfoButton(parent, size=size, command=_open)
    btn.grid(row=row, column=column, sticky=sticky, padx=padx, pady=pady)
    return btn


def attach_step_strip(parent, steps, *, sections=None, info_title="Help"):
    """Numbered on-screen order, with an optional i for the longer guide."""
    wrap = ttk.LabelFrame(parent, text="Do this in order", padding=6)
    wrap.pack(fill="x", padx=4, pady=(4, 8))
    row = FlowFrame(wrap)
    row.pack(fill="x")
    for index, step in enumerate(steps, 1):
        row.add(ttk.Label(row, text=f"{index}. {step}", font=("Arial", 10, "bold")))
    if sections:
        def _open():
            show_guide_dialog(parent, info_title, sections)

        row.add(InfoButton(row, size=18, command=_open))
    return wrap
