# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import itunes_verify

import openpyxl


def build(parent, app):
    theme = app.theme
    parent.configure(bg=theme["bg"])

    excel_path = app.get_excel_path()

    outer = tk.Frame(parent, bg=theme["bg"])
    outer.pack(fill="both", expand=True, padx=36, pady=28)

    header = tk.Frame(outer, bg=theme["bg"])
    header.pack(fill="x", pady=(0, 16))
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
    tk.Label(title_row, text="Web再調査(Apple Music)", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "辞書変換では対応できない表記ゆれを、iTunes Search APIで正式名称と照合します。\n"
                 "チェックを入れて選択した項目のみ、Apple Music側の表記に更新されます。"
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text="STEP 03 ・ 既に英語表記の項目を選んで正式表記を確認", bg=theme["bg"],
              fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    if not excel_path:
        tk.Label(outer, text="先に STEP 02 (MP3タグ抽出) を実行してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return

    banner = tk.Frame(outer, bg=theme["bg_elevated_2"])
    banner.pack(fill="x", pady=(0, 12))
    tk.Label(banner, text="選択した項目のみ Apple Music (iTunes Search API) に照会します。件数が多いと時間がかかります。",
              bg=theme["bg_elevated_2"], fg=theme["text_muted"], font=(FONT_BODY, 10),
              padx=14, pady=10, anchor="w", justify="left").pack(fill="x")

    toolbar = tk.Frame(outer, bg=theme["bg"])
    toolbar.pack(fill="x", pady=(0, 8))

    select_all_var = tk.BooleanVar(value=False)
    select_all_chk = tk.Checkbutton(toolbar, text="すべて選択", variable=select_all_var,
                                     bg=theme["bg"], fg=theme["text_muted"], selectcolor=theme["bg_elevated_2"],
                                     activebackground=theme["bg"], font=(FONT_BODY, 11),
                                     command=lambda: toggle_all())
    select_all_chk.pack(side="left")

    selected_count_var = tk.StringVar(value="選択中: 0件")
    tk.Label(toolbar, textvariable=selected_count_var, bg=theme["bg"], fg=theme["text_faint"],
              font=(FONT_BODY, 10)).pack(side="right", padx=(0, 14))
    run_btn_holder = tk.Frame(toolbar, bg=theme["bg"])
    run_btn_holder.pack(side="right")

    list_card = w.card(outer, theme)
    list_card.pack(fill="both", expand=True)
    canvas = tk.Canvas(list_card, bg=theme["bg_elevated"], highlightthickness=0)
    scrollbar = tk.Scrollbar(list_card, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg=theme["bg_elevated"])
    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
    scrollbar.pack(side="right", fill="y")
    w.enable_mousewheel(canvas)

    state = {"artist_vars": {}, "album_vars": {}}

    def update_selected_count():
        n = sum(v.get() for v in state["artist_vars"].values()) + sum(v.get() for v in state["album_vars"].values())
        selected_count_var.set(f"選択中: {n}件")

    def toggle_all():
        val = select_all_var.get()
        for v in list(state["artist_vars"].values()) + list(state["album_vars"].values()):
            v.set(val)
        update_selected_count()

    def section(title_text):
        h = tk.Frame(scroll_frame, bg=theme["bg_elevated"])
        h.pack(fill="x", padx=10, pady=(14, 4))
        tk.Label(h, text=title_text, bg=theme["bg_elevated"], fg=theme["text"],
                  font=(FONT_DISPLAY, 12, "bold")).pack(side="left")

    def add_row(name, var_dict):
        row = tk.Frame(scroll_frame, bg=theme["bg_elevated"])
        row.pack(fill="x", padx=10)
        var = tk.BooleanVar(value=False)
        var_dict[name] = var
        chk = tk.Checkbutton(row, variable=var, bg=theme["bg_elevated"], selectcolor=theme["bg_elevated_2"],
                              activebackground=theme["bg_elevated"], command=update_selected_count)
        chk.pack(side="left")
        tk.Label(row, text=name, bg=theme["bg_elevated"], fg=theme["text"], font=(FONT_MONO, 10),
                  anchor="w").pack(side="left", fill="x", expand=True, padx=(6, 0))

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb["IDTAG"]
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row))
    artist_names, album_names = itunes_verify.collect_english_names(rows)

    if not artist_names and not album_names:
        tk.Label(scroll_frame, text="既に英語表記の項目が見つかりませんでした。", bg=theme["bg_elevated"],
                  fg=theme["text_faint"], font=(FONT_BODY, 11)).pack(pady=30)
    else:
        section(f"アーティスト名 ({len(artist_names)}件)")
        for name in artist_names:
            add_row(name, state["artist_vars"])
        section(f"アルバム名 ({len(album_names)}件)")
        for name in album_names:
            add_row(name, state["album_vars"])

    def do_verify():
        selected_artists = [n for n, v in state["artist_vars"].items() if v.get()]
        selected_albums = [n for n, v in state["album_vars"].items() if v.get()]
        if not selected_artists and not selected_albums:
            messagebox.showinfo("未選択", "確認する項目を選択してください。")
            return

        def target(cb):
            wb2 = openpyxl.load_workbook(excel_path)
            ws2 = wb2["IDTAG"]
            rows2 = list(ws2.iter_rows(min_row=2, max_row=ws2.max_row))
            changes, n_applied = itunes_verify.verify_selected(rows2, selected_artists, selected_albums, cb)
            wb2.save(excel_path)
            return {"changes": changes, "applied": n_applied}

        def on_done(result):
            changes = result["changes"]
            if changes:
                lines = "\n".join(f"・{c['before']} → {c['after']}" for c in changes[:20])
                messagebox.showinfo("再調査完了", f"{result['applied']}件を更新しました。\n\n{lines}")
            else:
                messagebox.showinfo("再調査完了", "変更はありませんでした(すべて正式表記と一致)。")
            app.show_verify()

        app.run_task("Apple Musicへ照会中...", target, on_done)

    run_btn = w.button(run_btn_holder, "選択した項目を再調査", do_verify, theme, kind="primary")
    run_btn.pack()
