# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from platform_util import open_path
from theme import FONT_DISPLAY, FONT_BODY
from backend import name_lookup
from backend import ai_fill


def build(parent, app):
    theme = app.theme
    parent.configure(bg=theme["bg"])

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
    tk.Label(title_row, text="名前照会(Apple・Spotify・Discogs)", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "辞書で変換できずに日本語が残った曲名・アーティスト名・アルバム名を、3か所に照会します。\n"
                 "洋楽は原題(英字)、邦楽は本人・盤の正式表記を採用します。\n"
                 "2か所以上で一致したら緑、1か所だけ・食い違いは橙、候補が無ければ赤(そのまま)。\n"
                 "緑以外は Excel の「確認リスト」シートにまとめます。"
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text="STEP 03 ・ 日本語が残った名前をオンラインで照会",
              bg=theme["bg"], fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    body = tk.Frame(outer, bg=theme["bg"])
    body.pack(fill="both", expand=True)

    excel_path = app.get_excel_path()
    if not excel_path or not os.path.exists(excel_path):
        tk.Label(body, text="先に STEP 02 (MP3タグ抽出) を実行してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return

    keys = name_lookup.load_api_keys()
    available = name_lookup.available_sources(keys)
    labels = {"apple": "Apple Music", "spotify": "Spotify", "discogs": "Discogs"}
    saved = app.cfg.get("name_lookup_sources") or list(name_lookup.ALL_SOURCES)
    card = w.card(body, theme)
    card.pack(fill="x", pady=(0, 16))
    inner = tk.Frame(card, bg=theme["bg_elevated"])
    inner.pack(fill="x", padx=24, pady=16)
    tk.Label(inner, text="照会するサービス", anchor="w", bg=theme["bg_elevated"], fg=theme["text"],
              font=(FONT_BODY, 11, "bold")).pack(fill="x", pady=(0, 4))
    source_vars = {}
    for src in name_lookup.ALL_SOURCES:
        ok = src in available
        var = tk.BooleanVar(value=ok and src in saved)
        source_vars[src] = var
        note = {"apple": "キー不要・約3秒に1回", "spotify": "", "discogs": "約1秒に1回"}[src]
        text = labels[src] + (f"({note})" if ok and note else "") + ("" if ok else "(キー未設定のため使えません)")
        tk.Checkbutton(inner, text=text, variable=var, anchor="w", bg=theme["bg_elevated"], fg=theme["text"],
                        selectcolor=theme["bg_elevated_2"], activebackground=theme["bg_elevated"],
                        font=(FONT_BODY, 11), state="normal" if ok else "disabled").pack(fill="x")

    n = name_lookup.count_targets(excel_path)
    tk.Label(body, text=f"日本語が残っている行：{n}行\n"
                        "照会の結果は Excel の横に保存します。やり直したときは、前回緑になった名前は照会せず、\n"
                        "橙・赤の名前は、まだ照会していないサービスにだけ問い合わせます。\n"
                        "Apple は約3秒に1回しか問い合わせられないため、件数が多いと数十分〜数時間かかります。\n"
                        "問い合わせが多すぎて断られたときは、画面下に「待機中(あと○秒)」と出ます。キャンセルしても照会済みの分は残ります。",
              bg=theme["bg"], fg=theme["text_muted"], font=(FONT_BODY, 11), justify="left").pack(anchor="w", pady=(0, 12))

    result_box = tk.Frame(body, bg=theme["bg"])

    def render(result):
        for c in result_box.winfo_children():
            c.destroy()
        rc = w.card(result_box, theme)
        rc.pack(fill="x")
        ri = tk.Frame(rc, bg=theme["bg_elevated"])
        ri.pack(fill="x", padx=24, pady=16)
        tk.Label(ri, text=f"緑(採用) {result['adopted']}件 ・ 橙(要確認) {result['check']}件 ・ "
                          f"青(正式表記のまま) {result['kept']}件 ・ 赤(未解決) {result['unresolved']}件",
                  bg=theme["bg_elevated"], fg=theme["text"], font=(FONT_BODY, 12)).pack(anchor="w")
        tk.Label(ri, text=f"照会先：{'・'.join(result['sources'])}",
                  bg=theme["bg_elevated"], fg=theme["text_faint"], font=(FONT_BODY, 10)).pack(anchor="w", pady=(4, 0))
        w.button(ri, "Excelで確認する", lambda: open_path(excel_path), theme,
                 kind="secondary").pack(anchor="w", pady=(12, 0))

    def do_run():
        sources = [src for src, var in source_vars.items() if var.get()]
        if not sources:
            from tkinter import messagebox
            messagebox.showinfo("サービス未選択", "照会するサービスを1つ以上選んでください。")
            return
        app.cfg["name_lookup_sources"] = sources
        from config import save_config
        save_config(app.cfg)

        def target(cb):
            return name_lookup.run(excel_path, cb, sources=sources)

        def on_done(result):
            app.last_results["name_lookup"] = result
            app.completed_steps.discard("integrity_check")  # Excel が変わったので STEP 07 はやり直し
            app.show_name_lookup()

        app.run_task("名前を照会中...", target, on_done)

    def do_ai(engine):
        from tkinter import messagebox
        name = ai_fill.ENGINES[engine]
        n_red = ai_fill.count_red(excel_path)
        if not n_red:
            messagebox.showinfo("対象なし", "赤(未解決)のまま日本語が残っているセルはありません。")
            return
        if not messagebox.askyesno(
                f"{name} で補う",
                f"赤(未解決)のセル {n_red}件の元の表記を {name} に聞き、答えをサービスで確かめます。\n"
                f"確かめられたものは緑、確かめられなかったものは {name} の案を橙で入れます。\n"
                "答えが無いものは赤のまま残します。実行しますか?"):
            return
        sources = [src for src, var in source_vars.items() if var.get()]

        def target(cb):
            return ai_fill.run(excel_path, sources, cb, engine=engine)

        def on_done(result):
            app.last_results["ai_fill"] = result
            app.completed_steps.discard("integrity_check")
            messagebox.showinfo(f"{name} で補いました",
                                 f"{result['names']}件の名前を問い合わせました。\n"
                                 f"緑(確認済み) {result['green']}件・橙(要確認) {result['orange']}件・"
                                 f"青(正式表記のまま) {result['kept']}件・赤のまま {result['none']}件(セル数)。\n"
                                 "Excel の「確認リスト」にも載せました。")

        app.run_task(f"{name} で補っています...", target, on_done)

    btn_row = tk.Frame(body, bg=theme["bg"])
    btn_row.pack(anchor="w", pady=(0, 16))
    w.button(btn_row, "照会して Excel に書き込む", do_run, theme, kind="primary").pack(side="left")
    w.button(btn_row, "赤のセルを Claude で補う", lambda: do_ai("claude"), theme,
             kind="secondary").pack(side="left", padx=(8, 0))
    w.button(btn_row, "赤のセルを Codex で補う", lambda: do_ai("codex"), theme,
             kind="secondary").pack(side="left", padx=(8, 0))
    result_box.pack(fill="x")
    if app.last_results.get("name_lookup"):
        render(app.last_results["name_lookup"])
