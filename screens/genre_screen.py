# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import genre as backend
from backend import name_lookup as nl


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
    tk.Label(title_row, text="ジャンル分類", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "アルバムごとに Discogs の「ジャンル＞スタイル」を決めます。\n"
                 "MP3 のジャンルタグにはスタイル名(例: Boogie)を書きます。\n"
                 "候補の順:Excelに設定済み → アーティスト既定 → Discogs → Apple Music → Spotify。\n"
                 "「Excelに反映」を押すまで書き込まれません。", wraplength=420
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text="STEP 05 ・ アルバムごとにジャンル・スタイルを決めます", bg=theme["bg"],
              fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    excel_path = app.get_excel_path()
    if not excel_path:
        tk.Label(outer, text="先に STEP 02 (MP3タグ抽出) を実行してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return

    choices = app.genre_overrides  # album key -> (genre, style, "manual")
    state = {"items": [], "by_key": {}, "hierarchy": {}, "artist_default": {}}

    toolbar = tk.Frame(outer, bg=theme["bg"])
    toolbar.pack(fill="x", pady=(0, 8))
    undecided_only_var = tk.BooleanVar(value=False)
    tk.Checkbutton(toolbar, text="未決定のみ表示", variable=undecided_only_var,
                    bg=theme["bg"], fg=theme["text_muted"], selectcolor=theme["bg_elevated_2"],
                    activebackground=theme["bg"], font=(FONT_BODY, 11),
                    command=lambda: render_tree()).pack(side="left")
    summary_var = tk.StringVar(value="")
    tk.Label(toolbar, textvariable=summary_var, bg=theme["bg"], fg=theme["text_faint"],
              font=(FONT_BODY, 10)).pack(side="left", padx=(14, 0))
    w.button(toolbar, "Excelに反映", lambda: do_apply(), theme, kind="primary").pack(side="right")
    w.button(toolbar, "ジャンルを照会", lambda: do_fetch(), theme,
             kind="secondary").pack(side="right", padx=(0, 8))

    keys = nl.load_api_keys()
    available = nl.available_sources(keys)
    saved = app.cfg.get("genre_lookup_sources") or ["discogs", "apple", "spotify"]
    service_row = tk.Frame(outer, bg=theme["bg"])
    service_row.pack(fill="x", pady=(0, 8))
    tk.Label(service_row, text="照会するサービス:", bg=theme["bg"], fg=theme["text_muted"],
              font=(FONT_BODY, 11)).pack(side="left")
    source_vars = {}
    for src in ("discogs", "apple", "spotify"):
        ok = src in available
        source_vars[src] = tk.BooleanVar(value=ok and src in saved)
        note = {"discogs": "ジャンル＞スタイル", "apple": "大まかなジャンル・約3秒に1回",
                "spotify": "アーティストのジャンル"}[src]
        tk.Checkbutton(service_row, text=f"{backend.SERVICE_LABEL[src]}({note if ok else 'キー未設定'})",
                        variable=source_vars[src], bg=theme["bg"], fg=theme["text"],
                        selectcolor=theme["bg_elevated_2"], activebackground=theme["bg"], font=(FONT_BODY, 11),
                        state="normal" if ok else "disabled").pack(side="left", padx=(8, 0))

    body = tk.Frame(outer, bg=theme["bg"])
    body.pack(fill="both", expand=True)

    style = ttk.Style()
    style.configure("GN.Treeview", background=theme["bg_elevated"], fieldbackground=theme["bg_elevated"],
                    foreground=theme["text"], rowheight=24, font=(FONT_BODY, 11))
    style.configure("GN.Treeview.Heading", font=(FONT_BODY, 10, "bold"))
    list_frame = tk.Frame(body, bg=theme["bg"])
    list_frame.pack(side="left", fill="both", expand=True)
    tree = ttk.Treeview(list_frame, columns=("src", "artist", "album", "n", "style"), show="headings",
                         style="GN.Treeview", selectmode="browse")
    for col, text, width, stretch in (("src", "根拠", 130, False), ("artist", "アルバムアーティスト", 180, True),
                                      ("album", "アルバム", 250, True), ("n", "曲数", 45, False),
                                      ("style", "スタイル", 150, True)):
        tree.heading(col, text=text)
        tree.column(col, width=width, stretch=stretch, anchor="e" if col == "n" else "w")
    tree.tag_configure("none", foreground=theme["red"])
    tree.tag_configure("manual", foreground=theme["green"])
    sb = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="left", fill="y")

    panel_card = w.card(body, theme)
    panel_card.pack(side="left", fill="y", padx=(12, 0))
    panel = tk.Frame(panel_card, bg=theme["bg_elevated"], width=380)
    panel.pack(fill="both", expand=True, padx=16, pady=14)
    panel.pack_propagate(False)

    def current(item):
        if item["key"] in choices:
            g, s, src = choices[item["key"]]
            return g, s, src
        return item["genre"], item["style"], item["source"]

    def row_values(item):
        g, s, src = current(item)
        return (backend.SOURCE_LABEL[src], item["albumartist"], item["album"], item["tracks"], s or "")

    def row_tag(item):
        g, s, src = current(item)
        return ("manual",) if src == "manual" else (("none",) if not s else ())

    def update_summary():
        n_none = sum(1 for it in state["items"] if not current(it)[1])
        summary_var.set(f"アルバム {len(state['items'])}件 ・ 未決定 {n_none}件 ・ 手動 {len(choices)}件")

    def render_tree():
        tree.delete(*tree.get_children())
        for it in state["items"]:
            if undecided_only_var.get() and current(it)[1]:
                continue
            tree.insert("", "end", iid=it["key"], values=row_values(it), tags=row_tag(it))
        update_summary()

    def refresh_row(item):
        if tree.exists(item["key"]):
            tree.item(item["key"], values=row_values(item), tags=row_tag(item))
        update_summary()

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
        tk.Label(panel, text=f"{item['albumartist']} ・ {item['tracks']}曲", bg=bg, fg=theme["text_muted"],
                  font=(FONT_BODY, 10), wraplength=340, justify="left", anchor="w").pack(fill="x", pady=(2, 0))
        tk.Label(panel, text=item["folder"], bg=bg, fg=theme["text_faint"], font=(FONT_MONO, 9),
                  wraplength=340, justify="left", anchor="w").pack(fill="x", pady=(2, 0))
        found = item["found"]

        def found_text(src):
            if src not in found:
                return "未照会"
            v = found[src]
            if src == "discogs":
                return (" / ".join(v["genres"]) + " ＞ " + ", ".join(v["styles"])) if v["genres"] else "見つからず"
            if src == "spotify":
                return ", ".join(v) if v else "見つからず"
            return v or "見つからず"
        for src in ("discogs", "apple", "spotify"):
            tk.Label(panel, text=f"{backend.SERVICE_LABEL[src]}: {found_text(src)}", bg=bg, fg=theme["text_muted"],
                      font=(FONT_BODY, 10), anchor="w", wraplength=340, justify="left").pack(fill="x", pady=(4 if src == "discogs" else 0, 0))
        registered = state["artist_default"].get(item["albumartist"].lower())
        tk.Label(panel, text="登録済みのジャンル(アーティスト既定):\n"
                              + (f"{registered[0]} > {registered[1]}" if registered else "未登録"),
                  bg=bg, fg=theme["text"] if registered else theme["text_faint"],
                  font=(FONT_BODY, 10, "bold" if registered else "normal"), anchor="w",
                  wraplength=340, justify="left").pack(fill="x", pady=(2, 0))

        g, s, src = current(item)
        hierarchy = state["hierarchy"]
        genres = sorted(hierarchy, key=lambda x: (x == backend.UNKNOWN_PARENT, x.lower()))

        tk.Label(panel, text="ジャンル", bg=bg, fg=theme["text"], font=(FONT_BODY, 11, "bold"),
                  anchor="w").pack(fill="x", pady=(12, 2))
        genre_var = tk.StringVar(value=g or "")
        genre_cb = ttk.Combobox(panel, textvariable=genre_var, values=genres, state="readonly",
                                 font=(FONT_BODY, 11))
        genre_cb.pack(fill="x")

        tk.Label(panel, text="スタイル(タグに書く値)", bg=bg, fg=theme["text"], font=(FONT_BODY, 11, "bold"),
                  anchor="w").pack(fill="x", pady=(10, 2))
        style_var = tk.StringVar(value=s or "")
        style_cb = ttk.Combobox(panel, textvariable=style_var, state="readonly", font=(FONT_BODY, 11), height=20)
        style_cb.pack(fill="x")
        style_cb.configure(values=hierarchy.get(g, []) if g else [])

        def on_genre(_e=None):
            style_cb.configure(values=hierarchy.get(genre_var.get(), []))
            style_var.set("")

        def on_style(_e=None):
            if genre_var.get() and style_var.get():
                choices[item["key"]] = (genre_var.get(), style_var.get(), "manual")
                refresh_row(item)
        genre_cb.bind("<<ComboboxSelected>>", on_genre)
        style_cb.bind("<<ComboboxSelected>>", on_style)

        def add_style():
            genre = genre_var.get()
            if not genre:
                messagebox.showinfo("ジャンルを選択", "先にジャンルを選んでください。")
                return
            name = simpledialog.askstring("独自スタイルを追加", f"「{genre}」に追加するスタイル名:", parent=panel)
            if not name or not name.strip():
                return
            name = name.strip()
            if name not in hierarchy.get(genre, []):
                backend.add_custom_style(genre, name)
                hierarchy.setdefault(genre, []).append(name)
                hierarchy[genre].sort(key=str.lower)
            style_cb.configure(values=hierarchy[genre])
            style_var.set(name)
            on_style()

        def set_default():
            genre, st = genre_var.get(), style_var.get()
            if not genre or not st:
                messagebox.showinfo("未選択", "ジャンルとスタイルを選んでください。")
                return
            backend.set_artist_default(item["albumartist"], genre, st)
            state["artist_default"][item["albumartist"].lower()] = (genre, st)
            n = 0
            for it in state["items"]:
                if it["albumartist"].lower() == item["albumartist"].lower() and it["key"] not in choices:
                    it["genre"], it["style"], it["source"] = genre, st, "artist"
                    refresh_row(it)
                    n += 1
            # rebuild after the click handler returns: the pressed button itself is destroyed
            panel.after(0, lambda: render_panel(item))
            messagebox.showinfo("既定にしました",
                                 f"「{item['albumartist']}」の既定を {genre} > {st} にしました。\n"
                                 f"手動で選んでいないアルバム {n}件にも当てはめました。")

        btns = tk.Frame(panel, bg=bg)
        btns.pack(fill="x", pady=(12, 0))
        w.button(btns, "独自スタイルを追加", add_style, theme, kind="secondary").pack(anchor="w")
        if not item["various"]:  # a compilation's genre says nothing about "Various Artists"
            w.button(btns, "このアーティストの既定にする", set_default, theme, kind="secondary").pack(anchor="w", pady=(8, 0))
        if item["key"] in choices:
            def reset():
                choices.pop(item["key"], None)
                refresh_row(item)
                render_panel(item)
            w.button(btns, "手動の選択を取り消す", reset, theme, kind="secondary").pack(anchor="w", pady=(8, 0))

    def on_select(_e):
        sel = tree.selection()
        render_panel(state["by_key"].get(sel[0]) if sel else None)
    tree.bind("<<TreeviewSelect>>", on_select)

    def load():
        def target(cb):
            cb(0, 1, "アルバムを集計中...")
            return backend.build_items(excel_path)

        def on_done(result):
            if not tree.winfo_exists():  # the user left this screen while it was loading
                return
            items, hierarchy, artist_default = result
            state["items"], state["hierarchy"], state["artist_default"] = items, hierarchy, artist_default
            state["by_key"] = {it["key"]: it for it in items}
            for k in [k for k in choices if k not in state["by_key"]]:
                del choices[k]
            render_tree()
            render_panel(None)

        app.run_task("ジャンルの候補を集計中...", target, on_done)

    def do_fetch():
        sources = [src for src, var in source_vars.items() if var.get()]
        if not sources:
            messagebox.showinfo("サービス未選択", "照会するサービスを1つ以上選んでください。")
            return
        app.cfg["genre_lookup_sources"] = sources
        from config import save_config
        save_config(app.cfg)
        todo = [it for it in state["items"] if not set(sources) <= set(it["found"])]
        if not todo:
            messagebox.showinfo("照会済み", "選んだサービスでは、すべてのアルバムを照会済みです。")
            return
        per_album = {"discogs": 1.1, "apple": 3.2, "spotify": 0.4}
        minutes = max(1, round(len(todo) * max(per_album[src] for src in sources) / 60))
        if not messagebox.askyesno("ジャンルを照会",
                                   f"{len(todo)}件のアルバムを問い合わせます(約{minutes}分以上)。\n"
                                   "途中でキャンセルしても、照会した分は保存されます。続けますか?"):
            return

        def target(cb):
            return backend.fetch_genres(excel_path, todo, sources, cb)

        def on_done(result):
            load()

        app.run_task("ジャンルを照会中...", target, on_done)

    def do_apply():
        if not state["items"]:
            return
        undecided = [it for it in state["items"] if not current(it)[1]]
        if undecided and not messagebox.askyesno(
                "未決定があります",
                f"{len(undecided)}件のアルバムはスタイルが未決定です。\n"
                "未決定のアルバムはジャンルを空欄で書き込みます。続けますか?"):
            return
        all_choices = {it["key"]: current(it) for it in state["items"]}

        def target(cb):
            return backend.apply(excel_path, all_choices, cb)

        def on_done(result):
            result["unresolved"] = len(undecided)
            app.completed_steps.add("genre")
            app.completed_steps.discard("integrity_check")
            app.last_results["genre"] = result
            messagebox.showinfo("反映しました",
                                 f"{result['set']}曲に書き込みました(未決定 {len(undecided)}アルバム)。\n"
                                 "STEP 07 (ID3 タグ情報確認・適用) を再度実行してください。")
            app.show_dashboard()

        app.run_task("ジャンルを書き込み中...", target, on_done)

    render_panel(None)
    load()
