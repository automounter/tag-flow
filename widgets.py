# -*- coding: utf-8 -*-
import re
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO


# 行の頭に来てはいけない文字（この前では改行しない）
NO_LINE_START = "ぁぃぅぇぉっゃゅょァィゥェォッャュョ、。，．,.)）」』】]！？!?ー・：:；;%％"


def wrap_text(text, font, width):
    # Tk は空白でしか改行しないので、日本語は字体の幅を測って改行を入れる（Linux・Windows で崩れるため）
    f = tkfont.Font(font=font)
    out = []
    for para in text.split("\n"):
        line = ""
        for ch in para:
            if line and ch != " " and ch not in NO_LINE_START and f.measure(line + ch) > width:
                m = re.search(r"[A-Za-z0-9]+$", line) if ch.isascii() and ch.isalnum() else None
                if m and m.start() > 0:
                    out.append(line[:m.start()].rstrip())
                    line = line[m.start():] + ch
                else:
                    out.append(line.rstrip())
                    line = ch
            elif line or ch != " ":
                line += ch
        out.append(line)
    return "\n".join(out)


def card(parent, theme, **kwargs):
    f = tk.Frame(parent, bg=theme["bg_elevated"], highlightbackground=theme["border"],
                 highlightthickness=1, bd=0, **kwargs)
    return f


def h_line(parent, theme):
    return tk.Frame(parent, bg=theme["border"], height=1)


def label(parent, text, theme, size=13, color_key="text", bold=False, mono=False, wraplength=0, anchor="w"):
    font_family = FONT_MONO if mono else FONT_BODY
    weight = "bold" if bold else "normal"
    font = (font_family, size, weight)
    if wraplength:
        text = wrap_text(text, font, wraplength)
    return tk.Label(parent, text=text, bg=parent["bg"], fg=theme[color_key],
                     font=font, justify="left", anchor=anchor)


def title_label(parent, text, theme, size=16):
    return tk.Label(parent, text=text, bg=parent["bg"], fg=theme["text"],
                     font=(FONT_DISPLAY, size, "bold"), anchor="w")


def pill(parent, text, theme, bg_key="bg_elevated_2", fg_key="text_muted", size=11):
    f = tk.Frame(parent, bg=theme[bg_key])
    lbl = tk.Label(f, text=text, bg=theme[bg_key], fg=theme[fg_key], font=(FONT_MONO, size),
                    padx=10, pady=4)
    lbl.pack()
    return f


def status_badge(parent, text, theme, color_key):
    f = tk.Frame(parent, bg=parent["bg"])
    dot = tk.Canvas(f, width=8, height=8, bg=parent["bg"], highlightthickness=0)
    dot.create_oval(1, 1, 7, 7, fill=theme[color_key], outline="")
    dot.pack(side="left", padx=(0, 6))
    lbl = tk.Label(f, text=text, bg=parent["bg"], fg=theme[color_key], font=(FONT_BODY, 11))
    lbl.pack(side="left")
    return f


class ClickButton(tk.Frame):
    """A colored, clickable button. Plain tk.Button ignores bg/fg colors on
    macOS Aqua, so we draw our own using a Frame + Label and a click binding."""

    def __init__(self, parent, text, command, theme, kind="primary"):
        self._theme = theme
        self._kind = kind
        self._command = command
        self._enabled = True
        bg, fg = self._colors()
        super().__init__(parent, bg=bg, cursor="pointinghand" if _is_mac() else "hand2")
        self._label = tk.Label(self, text=text, bg=bg, fg=fg, font=(FONT_BODY, 12), padx=16, pady=8)
        self._label.pack()
        self.bind("<Button-1>", self._on_click)
        self._label.bind("<Button-1>", self._on_click)

    def _colors(self):
        theme = self._theme
        if not self._enabled:
            return theme["bg_elevated_2"], theme["text_faint"]
        if self._kind == "primary":
            return theme["accent"], theme["accent_text"]
        if self._kind == "danger":
            return theme["red"], "#FFFFFF"
        return theme["bg_elevated_2"], theme["text"]

    def _on_click(self, event=None):
        if self._enabled and self._command:
            self._command()

    def configure(self, **kwargs):
        state = kwargs.pop("state", None)
        if state is not None:
            self._enabled = (state != "disabled")
            bg, fg = self._colors()
            self.configure_colors(bg, fg)
        if kwargs:
            super().configure(**kwargs)

    config = configure

    def configure_colors(self, bg, fg):
        tk.Frame.configure(self, bg=bg, cursor=("pointinghand" if (self._enabled and _is_mac()) else "arrow"))
        self._label.configure(bg=bg, fg=fg)


def button(parent, text, command, theme, kind="primary"):
    return ClickButton(parent, text, command, theme, kind=kind)


def _is_mac():
    import platform
    return platform.system() == "Darwin"


def run_async(widget, work_fn, on_done, interval=150):
    """Runs work_fn() in a background thread, then delivers the result back on
    the main thread via a self-rescheduling `widget.after()` poll.

    Tkinter widgets may only be touched from the thread that created them,
    and on this bundled Tk build a `.after()` call made directly from a
    background thread is silently never invoked (the Tcl notifier isn't
    woken across threads) -- so any screen doing background work (a network
    call, a slow filesystem scan) must schedule its callback from a poll
    that was itself first scheduled on the main thread, mirroring
    TagFlowApp.run_task. on_done(result, error) is called with error=None
    on success, or the caught exception otherwise."""
    import threading

    box = {"result": None, "error": None, "done": False}

    def worker():
        try:
            box["result"] = work_fn()
        except Exception as e:
            box["error"] = e
        box["done"] = True

    threading.Thread(target=worker, daemon=True).start()

    def poll():
        if not box["done"]:
            widget.after(interval, poll)
            return
        on_done(box["result"], box["error"])

    widget.after(interval, poll)


class Tooltip:
    """Shows `text` in a small borderless popup while the pointer hovers
    over `widget`. Attach and forget -- e.g. `Tooltip(help_badge, "...")`."""

    def __init__(self, widget, text, wraplength=280):
        self.widget = widget
        self.text = text
        self.wraplength = wraplength
        self._tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, event=None):
        if self._tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 4
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        try:
            self._tip.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self._tip, text=self.text, justify="left", bg="#2B2B2B", fg="#F2F2F2",
                 font=(FONT_BODY, 10), padx=10, pady=8, wraplength=self.wraplength).pack()
        # this Tk build doesn't always paint an overrideredirect Toplevel's
        # content on its first appearance without a forced redraw
        self._tip.update_idletasks()
        self._tip.lift()

    def _hide(self, event=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


def help_badge(parent, theme, text, wraplength=280):
    """A small circular '?' badge that shows `text` in a tooltip on hover."""
    badge = tk.Label(parent, text="?", bg=theme["bg_elevated_2"], fg=theme["text_muted"],
                      font=(FONT_BODY, 9, "bold"), width=2, cursor="hand2")
    Tooltip(badge, text, wraplength=wraplength)
    return badge


def enable_mousewheel(canvas):
    """Lets trackpad/mouse-wheel scrolling drive `canvas`, active only while
    the pointer is over it (bind_all is scoped to hover so it doesn't hijack
    scrolling on whatever screen is shown next)."""
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * event.delta), "units")

    def _bind(event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _unbind(event):
        canvas.unbind_all("<MouseWheel>")

    canvas.bind("<Enter>", _bind)
    canvas.bind("<Leave>", _unbind)


def progress_bar(parent, theme):
    style_name = "TagFlow.Horizontal.TProgressbar"
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(style_name, troughcolor=theme["bg_elevated_2"], background=theme["accent"],
                     bordercolor=theme["bg_elevated_2"], lightcolor=theme["accent"], darkcolor=theme["accent"])
    pb = ttk.Progressbar(parent, style=style_name, orient="horizontal", mode="determinate")
    return pb
