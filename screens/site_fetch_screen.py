# -*- coding: utf-8 -*-
"""サイトから取り寄せる（2026-11-25）。手元のフォルダを MP3 データベースのサイトと突き合わせ、合ったタグを書く。
Excel を使う STEP 01〜08 とは別の道。A（全部合う）は最初から印あり、B は人が見て決める、C は書かない。
"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from platform_util import open_path
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import site_fetch

GRADE_TEXT = {"A": ("A 書いてよい", "green"), "B": ("B 見て決める", "yellow"), "C": ("C 書かない", "red")}


def build(parent, app):
    theme = app.theme
    parent.configure(bg=theme["bg"])
    state = app.__dict__.setdefault("site_fetch_state", {"root": app.cfg.get("source_folder", ""), "albums": None})

    outer = tk.Frame(parent, bg=theme["bg"])
    outer.pack(fill="both", expand=True, padx=36, pady=28)

    header = tk.Frame(outer, bg=theme["bg"])
    header.pack(fill="x", pady=(0, 14))
    back = tk.Label(header, text="←", bg=theme["bg_elevated_2"], fg=theme["text_muted"], font=(FONT_BODY, 14),
                    width=2, cursor="hand2")
    back.pack(side="left")
    back.bind("<Button-1>", lambda e: app.show_dashboard())
    title_box = tk.Frame(header, bg=theme["bg"])
    title_box.pack(side="left", padx=(12, 0))
    row = tk.Frame(title_box, bg=theme["bg"])
    row.pack(anchor="w")
    tk.Label(row, text="サイトから取り寄せる", bg=theme["bg"], fg=theme["text"], font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(row, theme,
                 "フォルダ（アルバム1枚）ごとに、曲名・曲番号・長さを MP3 データベースのサイトに送り、合ったアルバムのタグを受け取ります。\n"
                 "A：全部の曲が合った（最初から選ばれています）／B：だいたい合った（見て選んでください）／C：合わない（書きません）。\n"
                 "書く前に、元のタグを控えに取ります（STEP02 と同じ控え。元に戻せます）。"
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text="STEP 08 の別の道 ・ Excel を使わずに、サイトのタグを手元の MP3 に書く", bg=theme["bg"],
             fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    # フォルダ
    frow = tk.Frame(outer, bg=theme["bg_elevated"], highlightbackground=theme["border"], highlightthickness=1)
    frow.pack(fill="x")
    root_var = tk.StringVar(value=state["root"])
    tk.Entry(frow, textvariable=root_var, bg=theme["bg_elevated"], fg=theme["text"], insertbackground=theme["text"],
             font=(FONT_MONO, 10), relief="flat", bd=0).pack(side="left", fill="x", expand=True, padx=10, pady=6)

    def browse():
        p = filedialog.askdirectory(initialdir=root_var.get() or "/")
        if p:
            root_var.set(p)
    b = tk.Label(frow, text="参照", bg=theme["bg_elevated"], fg=theme["accent"], font=(FONT_BODY, 10), cursor="hand2", padx=14)
    b.pack(side="right")
    b.bind("<Button-1>", lambda e: browse())

    # 書く項目
    fbox = tk.Frame(outer, bg=theme["bg"])
    fbox.pack(fill="x", pady=(10, 0))
    tk.Label(fbox, text="書く項目：", bg=theme["bg"], fg=theme["text_muted"], font=(FONT_BODY, 10)).grid(row=0, column=0, sticky="nw")
    field_vars = {}
    # 14項目を7つずつ2行に（1行だと Linux の字体で右が切れた。2026-11-26）
    for i, (key, label) in enumerate(site_fetch.FIELDS):
        v = tk.BooleanVar(value=True)
        field_vars[key] = v
        tk.Checkbutton(fbox, text=label, variable=v, bg=theme["bg"], fg=theme["text_muted"], selectcolor=theme["bg_elevated_2"],
                       activebackground=theme["bg"], font=(FONT_BODY, 10), highlightthickness=0, bd=0).grid(
            row=i // 7, column=1 + i % 7, sticky="w", padx=(0, 10))
    opt = tk.Frame(outer, bg=theme["bg"])
    opt.pack(fill="x")
    cover_missing_var = tk.BooleanVar(value=True)
    tk.Checkbutton(opt, text="ジャケットは、画像が無い曲にだけ入れる", variable=cover_missing_var, bg=theme["bg"],
                   fg=theme["text_muted"], selectcolor=theme["bg_elevated_2"], activebackground=theme["bg"],
                   font=(FONT_BODY, 10), highlightthickness=0, bd=0).pack(side="left")
    tk.Label(opt, text=f"ID3 の版：{'v2.3' if app.cfg.get('id3_version') == 'v23' else 'v2.4'}（設定で変えられます）",
             bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 10)).pack(side="left", padx=(16, 0))

    toolbar = tk.Frame(outer, bg=theme["bg"])
    toolbar.pack(fill="x", pady=(10, 8))
    summary_var = tk.StringVar(value="")
    tk.Label(toolbar, textvariable=summary_var, bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 10)).pack(side="left")

    list_card = w.card(outer, theme)
    list_card.pack(fill="both", expand=True)
    canvas = tk.Canvas(list_card, bg=theme["bg_elevated"], highlightthickness=0)
    sb = tk.Scrollbar(list_card, orient="vertical", command=canvas.yview)
    rows = tk.Frame(canvas, bg=theme["bg_elevated"])
    rows.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=rows, anchor="nw")
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    w.enable_mousewheel(canvas)

    sel_vars = []

    def update_summary():
        albums = state["albums"] or []
        if not albums:
            summary_var.set("フォルダを選んで「読み込んで問い合わせる」を押してください。")
            return
        g = [a.get("result", {}).get("grade", "C") for a in albums]
        n_sel = sum(1 for a, v in sel_vars if v.get())
        summary_var.set(f"{len(albums)} 枚：A {g.count('A')}・B {g.count('B')}・C {g.count('C')}　選択中 {n_sel} 枚")

    def render():
        for c in rows.winfo_children():
            c.destroy()
        sel_vars.clear()
        for a in state["albums"] or []:
            r = a.get("result") or {"grade": "C", "score": 0}
            grade = r.get("grade", "C")
            line = tk.Frame(rows, bg=theme["bg_elevated"])
            line.pack(fill="x", padx=12, pady=4)
            v = tk.BooleanVar(value=a.get("selected", grade == "A"))
            sel_vars.append((a, v))
            tk.Checkbutton(line, variable=v, bg=theme["bg_elevated"], selectcolor=theme["bg_elevated_2"],
                           activebackground=theme["bg_elevated"], state="disabled" if grade == "C" else "normal",
                           highlightthickness=0, bd=0,
                           command=update_summary).pack(side="left")
            text, color = GRADE_TEXT.get(grade, GRADE_TEXT["C"])
            tk.Label(line, text=text, bg=theme["bg_elevated"], fg=theme[color], font=(FONT_BODY, 10, "bold"), width=12,
                     anchor="w").pack(side="left")
            tk.Label(line, text=os.path.relpath(a["folder"], state["root"]) or os.path.basename(a["folder"]),
                     bg=theme["bg_elevated"], fg=theme["text"], font=(FONT_BODY, 11), anchor="w", width=46).pack(side="left")
            al = r.get("album")
            info = (f"→ {al['title']}（{al.get('artist') or ''}・{al.get('year') or ''}）  合った曲 {al['tracks_matched']}/{len(a['tracks'])}"
                    f"  点数 {r.get('score', 0)}") if al else (r.get("note") or "合うアルバムがありません")
            tk.Label(line, text=info, bg=theme["bg_elevated"], fg=theme["text_muted"], font=(FONT_BODY, 10),
                     anchor="w").pack(side="left", fill="x", expand=True)
        update_summary()

    def do_ask():
        root = root_var.get().strip()
        if not os.path.isdir(root):
            messagebox.showwarning("フォルダ", "フォルダが見つかりません。")
            return
        state["root"] = root
        url = app.cfg.get("mp3db_api_url") or site_fetch.DEFAULT_URL
        key = app.cfg.get("mp3db_api_key", "")

        def target(cb):
            albums = site_fetch.scan(root, cb)
            return site_fetch.ask(albums, url, key, cb)

        def done(albums):
            state["albums"] = albums
            render()
        app.run_task("サイトから取り寄せ：読み込みと問い合わせ", target, done)

    def do_write():
        albums = state["albums"] or []
        for a, v in sel_vars:
            a["selected"] = v.get()
        chosen = [a for a in albums if a.get("selected")]
        fields = [k for k, v in field_vars.items() if v.get()]
        if not chosen or not fields:
            messagebox.showinfo("書く", "書くアルバムと項目を選んでください。")
            return
        n = sum(1 for a in chosen for t in a["result"].get("tracks", []) if t.get("track_id"))
        if not messagebox.askyesno("確認", f"{len(chosen)} 枚・{n} 曲のタグを書きます。\n"
                                          "書く前に、元のタグを控えに取ります（元に戻せます）。\n\nよろしいですか？"):
            return

        def target(cb):
            res = site_fetch.write(albums, fields, app.cfg.get("id3_version", "v24"), cover_missing_var.get(), cb)
            res["log_path"] = site_fetch.save_log(res, app.cfg.get("excel_folder") or state["root"])
            return res

        def done(res):
            if messagebox.askyesno("書きました", f"{res['files']} 曲を処理し、{len(res['rows'])} か所を書き換えました。\n"
                                                f"元のタグの控え：{res['backup_dir']}\n記録：{res['log_path']}\n\n記録を開きますか？"):
                open_path(res["log_path"])
        app.run_task("サイトから取り寄せ：書き込み", target, done)

    w.button(toolbar, "選んだアルバムに書く", do_write, theme, kind="primary").pack(side="right")
    w.button(toolbar, "読み込んで問い合わせる", do_ask, theme, kind="secondary").pack(side="right", padx=(0, 8))
    render()
