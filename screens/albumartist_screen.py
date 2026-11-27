# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import albumartist as backend


def build(parent, app):
    theme = app.theme
    parent.configure(bg=theme["bg"])

    outer = tk.Frame(parent, bg=theme["bg"])
    outer.pack(fill="both", expand=True, padx=36, pady=28)

    header = tk.Frame(outer, bg=theme["bg"])
    header.pack(fill="x", pady=(0, 12))
    back = tk.Frame(header, bg=theme["bg_elevated_2"], width=32, height=32,
                     highlightbackground=theme["border"], highlightthickness=1)
    back.pack(side="left")
    back.pack_propagate(False)
    back_lbl = tk.Label(back, text="←", bg=theme["bg_elevated_2"], fg=theme["text_muted"], font=(FONT_BODY, 14))
    back_lbl.pack(expand=True)
    for widget in (back, back_lbl):
        widget.bind("<Button-1>", lambda e: app.show_dashboard())
        widget.configure(cursor="hand2")

    title_box = tk.Frame(header, bg=theme["bg"])
    title_box.pack(side="left", padx=(12, 0))
    title_row = tk.Frame(title_box, bg=theme["bg"])
    title_row.pack(anchor="w")
    tk.Label(title_row, text="アルバムアーティスト推定", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "アルバム(フォルダ)ごとに判定します。\n"
                 "・単独/グループ名/主要アーティスト(過半数の曲に参加)は自動で決定\n"
                 "・複数並記は、各アーティストか並記のままかを選択\n"
                 "・過半数の曲に参加するアーティストがいない場合は Various Artists か参加者から選択\n"
                 "「Excelに反映」を押すまで書き込まれません。", wraplength=420
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text="STEP 04 ・ アルバムごとにアルバムアーティストを決めます", bg=theme["bg"],
              fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    excel_path = app.get_excel_path()
    if not excel_path:
        tk.Label(outer, text="先に STEP 02 (MP3タグ抽出) を実行してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return

    choices = app.albumartist_choices  # item key -> chosen album artist
    state = {"items": [], "by_key": {}, "selected": None}

    toolbar = tk.Frame(outer, bg=theme["bg"])
    toolbar.pack(fill="x", pady=(0, 8))
    review_only_var = tk.BooleanVar(value=True)
    tk.Checkbutton(toolbar, text="要確認のみ表示", variable=review_only_var,
                    bg=theme["bg"], fg=theme["text_muted"], selectcolor=theme["bg_elevated_2"],
                    activebackground=theme["bg"], font=(FONT_BODY, 11),
                    command=lambda: render_tree()).pack(side="left")
    summary_var = tk.StringVar(value="")
    tk.Label(toolbar, textvariable=summary_var, bg=theme["bg"], fg=theme["text_faint"],
              font=(FONT_BODY, 10)).pack(side="left", padx=(14, 0))
    w.button(toolbar, "Excelに反映", lambda: do_apply(), theme, kind="primary").pack(side="right")

    body = tk.Frame(outer, bg=theme["bg"])
    body.pack(fill="both", expand=True)

    # ---- left: album list ----
    style = ttk.Style()
    style.configure("AA.Treeview", background=theme["bg_elevated"], fieldbackground=theme["bg_elevated"],
                    foreground=theme["text"], rowheight=24, font=(FONT_BODY, 11))
    style.configure("AA.Treeview.Heading", font=(FONT_BODY, 10, "bold"))
    list_frame = tk.Frame(body, bg=theme["bg"])
    list_frame.pack(side="left", fill="both", expand=True)
    tree = ttk.Treeview(list_frame, columns=("cat", "album", "n", "value"), show="headings",
                         style="AA.Treeview", selectmode="browse")
    for col, text, width, stretch in (("cat", "区分", 110, False), ("album", "アルバム", 320, True),
                                      ("n", "曲数", 50, False), ("value", "アルバムアーティスト", 260, True)):
        tree.heading(col, text=text)
        tree.column(col, width=width, stretch=stretch, anchor="e" if col == "n" else "w")
    tree.tag_configure("review", foreground=theme["red"])
    tree.tag_configure("decided", foreground=theme["green"])
    sb = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="left", fill="y")

    # ---- right: decision panel ----
    panel_card = w.card(body, theme)
    panel_card.pack(side="left", fill="y", padx=(12, 0))
    panel = tk.Frame(panel_card, bg=theme["bg_elevated"], width=380)
    panel.pack(fill="both", expand=True, padx=16, pady=14)
    panel.pack_propagate(False)

    def value_of(item):
        return choices.get(item["key"], item["default"])

    def render_tree():
        tree.delete(*tree.get_children())
        for it in state["items"]:
            review = backend.needs_review(it)
            if review_only_var.get() and not review:
                continue
            cat = backend.CATEGORY_LABEL[it["category"]]
            if it["mixed"]:
                cat += "・分割"
            tag = ("decided",) if it["key"] in choices else (("review",) if review else ())
            tree.insert("", "end", iid=it["key"], values=(cat, it["album"], it["tracks"], value_of(it)), tags=tag)
        n_review = sum(1 for it in state["items"] if backend.needs_review(it))
        n_done = sum(1 for it in state["items"] if backend.needs_review(it) and it["key"] in choices)
        summary_var.set(f"アルバム {len(state['items'])}件 ・ 要確認 {n_review}件(選択済み {n_done}件)")

    def refresh_row(item):
        if tree.exists(item["key"]):
            vals = list(tree.item(item["key"], "values"))
            vals[3] = value_of(item)
            tree.item(item["key"], values=vals, tags=("decided",))
        render_summary_only()

    def render_summary_only():
        n_review = sum(1 for it in state["items"] if backend.needs_review(it))
        n_done = sum(1 for it in state["items"] if backend.needs_review(it) and it["key"] in choices)
        summary_var.set(f"アルバム {len(state['items'])}件 ・ 要確認 {n_review}件(選択済み {n_done}件)")

    def render_panel(item):
        for c in panel.winfo_children():
            c.destroy()
        bg = theme["bg_elevated"]
        if item is None:
            tk.Label(panel, text="左の一覧からアルバムを選んでください。", bg=bg, fg=theme["text_faint"],
                      font=(FONT_BODY, 11), wraplength=340, justify="left").pack(anchor="w")
            return
        tk.Label(panel, text=item["album"], bg=bg, fg=theme["text"], font=(FONT_BODY, 12, "bold"),
                  wraplength=340, justify="left", anchor="w").pack(fill="x")
        info = f"{backend.CATEGORY_LABEL[item['category']]} ・ {item['tracks']}曲"
        if item["mixed"]:
            info += " ・ 同じフォルダの同名アルバムから分割"
        tk.Label(panel, text=info, bg=bg, fg=theme["text_muted"], font=(FONT_BODY, 10),
                  wraplength=340, justify="left", anchor="w").pack(fill="x", pady=(2, 0))
        tk.Label(panel, text=item["folder"], bg=bg, fg=theme["text_faint"], font=(FONT_MONO, 9),
                  wraplength=340, justify="left", anchor="w").pack(fill="x", pady=(2, 0))
        if item["original"]:
            tk.Label(panel, text=f"元のタグ: {item['original']}", bg=bg, fg=theme["text_faint"],
                      font=(FONT_BODY, 10), wraplength=340, justify="left", anchor="w").pack(fill="x", pady=(2, 0))

        tk.Label(panel, text="アルバムアーティスト", bg=bg, fg=theme["text"], font=(FONT_BODY, 11, "bold"),
                  anchor="w").pack(fill="x", pady=(12, 4))
        current = value_of(item)
        options = list(item["options"])
        if item["category"] == "various":
            options = [backend.VARIOUS]
        choice_var = tk.StringVar(value=current if current in options else "__manual__")
        manual_var = tk.StringVar(value="" if current in options else current)

        def choose(value):
            value = value.strip()
            if not value:
                return
            choices[item["key"]] = value
            refresh_row(item)

        for opt in options:
            tk.Radiobutton(panel, text=opt, value=opt, variable=choice_var, bg=bg, fg=theme["text"],
                            selectcolor=theme["bg_elevated_2"], activebackground=bg, anchor="w",
                            wraplength=330, justify="left", font=(FONT_BODY, 11),
                            command=lambda o=opt: choose(o)).pack(fill="x")
        manual_row = tk.Frame(panel, bg=bg)
        manual_row.pack(fill="x", pady=(2, 0))
        tk.Radiobutton(manual_row, text="手入力:", value="__manual__", variable=choice_var, bg=bg,
                        fg=theme["text"], selectcolor=theme["bg_elevated_2"], activebackground=bg,
                        font=(FONT_BODY, 11), command=lambda: choose(manual_var.get())).pack(side="left")
        entry = tk.Entry(manual_row, textvariable=manual_var, font=(FONT_BODY, 11))
        entry.pack(side="left", fill="x", expand=True)

        def on_manual(_e=None):
            choice_var.set("__manual__")
            choose(manual_var.get())
        entry.bind("<Return>", on_manual)
        entry.bind("<FocusOut>", lambda e: choice_var.get() == "__manual__" and choose(manual_var.get()))

        whole = item["options"][-1]
        if item["category"] == "multi" and whole.lower() not in (g.lower() for g in backend.load_groups()):
            def register_group():
                if not messagebox.askyesno("グループ名として登録",
                                           f"「{whole}」を1つのグループ名として登録しますか?\n"
                                           "以後は分割せず、そのまま使います(全アルバムを再判定します)。"):
                    return
                backend.add_group(whole)
                choices.pop(item["key"], None)
                load()
            w.button(panel, "並記をグループ名として登録", register_group, theme,
                     kind="secondary").pack(anchor="w", pady=(8, 0))

        tk.Label(panel, text="参加アーティスト(ダブルクリックで採用)", bg=bg, fg=theme["text"],
                  font=(FONT_BODY, 11, "bold"), anchor="w").pack(fill="x", pady=(14, 4))
        lb_frame = tk.Frame(panel, bg=bg)
        lb_frame.pack(fill="both", expand=True)
        lb = tk.Listbox(lb_frame, font=(FONT_BODY, 11), activestyle="none", height=8,
                         bg=theme["bg_elevated_2"], fg=theme["text"], highlightthickness=0, borderwidth=0)
        lsb = ttk.Scrollbar(lb_frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=lsb.set)
        lb.pack(side="left", fill="both", expand=True)
        lsb.pack(side="left", fill="y")
        names = [p for p, _ in item["participants"]]
        for p, c in item["participants"]:
            lb.insert("end", f"{p}  ({c}/{item['tracks']}曲)")

        def on_pick(_e):
            sel = lb.curselection()
            if sel:
                manual_var.set(names[sel[0]])
                choice_var.set("__manual__")
                choose(names[sel[0]])
        lb.bind("<Double-Button-1>", on_pick)

    def on_select(_e):
        sel = tree.selection()
        item = state["by_key"].get(sel[0]) if sel else None
        render_panel(item)
    tree.bind("<<TreeviewSelect>>", on_select)

    def load():
        base, ext = os.path.splitext(excel_path)
        baseline = getattr(app, "baseline_excel_path", None) or f"{base}_baseline{ext}"

        def target(cb):
            cb(0, 1, "アルバムを分析中...")
            return backend.analyze(excel_path, baseline)

        def on_done(items):
            if not tree.winfo_exists():  # the user left this screen while it was loading
                return
            state["items"] = items
            state["by_key"] = {it["key"]: it for it in items}
            for k in [k for k in choices if k not in state["by_key"]]:
                del choices[k]
            render_tree()
            render_panel(None)

        app.run_task("アルバムアーティストを分析中...", target, on_done)

    def do_apply():
        if not state["items"]:
            return
        pending = [it for it in state["items"] if backend.needs_review(it) and it["key"] not in choices]
        if pending and not messagebox.askyesno(
                "未選択があります",
                f"要確認のうち {len(pending)}件が未選択です。\n"
                "未選択のアルバムは一覧に表示している候補(初期値)で書き込みます。続けますか?"):
            return
        file_choices = {}
        for it in state["items"]:
            value = value_of(it)
            for f in it["files"]:
                file_choices[f] = value

        def target(cb):
            return backend.apply(excel_path, file_choices, cb)

        def on_done(result):
            app.completed_steps.add("albumartist")
            app.completed_steps.discard("integrity_check")
            app.last_results["albumartist"] = result
            messagebox.showinfo("反映しました",
                                 f"{result['set']}曲に書き込みました(Various Artists {result['various_artists']}曲)。\n"
                                 "ジャンルはアルバムアーティストから決まるため、STEP 05 も再実行してください。")
            app.show_dashboard()

        app.run_task("アルバムアーティストを書き込み中...", target, on_done)

    render_panel(None)
    load()
