# -*- coding: utf-8 -*-
import datetime
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from platform_util import open_path
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from config import save_config
from backend import site_fetch
from logger import LOG_PATH
from data.genre_mapping import generate_alias_fallback, ALIAS_FALLBACK_FILENAME


def build(parent, app):
    theme = app.theme
    parent.configure(bg=theme["bg"])

    outer = tk.Frame(parent, bg=theme["bg"])
    outer.pack(fill="both", expand=True, padx=28, pady=18)

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
    tk.Label(title_box, text="設定", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 15, "bold")).pack(anchor="w")
    tk.Label(title_box, text="保存先フォルダ・ジャンル分類ルール・表示テーマを管理", bg=theme["bg"],
              fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    # scrollable content area, with the save/cancel footer pinned below it so
    # it's always reachable even when the window is shorter than the content
    scroll_area = tk.Frame(outer, bg=theme["bg"])
    scroll_area.pack(fill="both", expand=True)
    canvas = tk.Canvas(scroll_area, bg=theme["bg"], highlightthickness=0)
    scrollbar = tk.Scrollbar(scroll_area, orient="vertical", command=canvas.yview)
    body = tk.Frame(canvas, bg=theme["bg"])
    body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas_window = canvas.create_window((0, 0), window=body, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    w.enable_mousewheel(canvas)

    path_vars = {}

    _section(body, theme, "フォルダ設定", "MP3の読み込み元と、整理後のファイルの書き出し先を指定します。")
    path_vars["source_folder"] = _folder_field(body, theme, "MP3保存元フォルダ", app.cfg["source_folder"],
                                                is_dir=True)
    path_vars["output_folder"] = _folder_field(body, theme, "出力先フォルダ", app.cfg["output_folder"],
                                                is_dir=True)
    path_vars["excel_folder"] = _folder_field(body, theme, "Excel保存先フォルダ", app.cfg["excel_folder"],
                                               is_dir=True)

    excel_box = tk.Frame(body, bg=theme["bg"])
    excel_box.pack(fill="x", pady=(0, 6))
    excel_label_row = tk.Frame(excel_box, bg=theme["bg"])
    excel_label_row.pack(fill="x")
    tk.Label(excel_label_row, text="現在使用中のExcelファイル", bg=theme["bg"], fg=theme["text_muted"],
              font=(FONT_BODY, 10), anchor="w").pack(side="left")
    w.help_badge(excel_label_row, theme,
                 "STEP02(タグ抽出)を実行すると自動的にここが更新されます。\n"
                 "以前作成したExcelファイルに作業を戻したい場合は「選択」で切り替えられます。"
                 ).pack(side="left", padx=(6, 0))

    excel_row = tk.Frame(excel_box, bg=theme["bg_elevated"], highlightbackground=theme["border"],
                          highlightthickness=1)
    excel_row.pack(fill="x", pady=(4, 0))
    excel_path_var = tk.StringVar(value=app.cfg.get("excel_path") or "(未選択)")
    tk.Label(excel_row, textvariable=excel_path_var, bg=theme["bg_elevated"], fg=theme["text"],
              font=(FONT_MONO, 10), anchor="w").pack(side="left", fill="x", expand=True, padx=10, pady=6)

    def browse_excel_file():
        path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            initialdir=app.cfg.get("excel_folder") or "/",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
        )
        if path:
            app.set_excel_path(path)
            excel_path_var.set(path)

    excel_browse_lbl = tk.Label(excel_row, text="選択", bg=theme["bg_elevated"], fg=theme["accent"],
                                  font=(FONT_BODY, 10), cursor="hand2", padx=14)
    excel_browse_lbl.pack(side="right")
    excel_browse_lbl.bind("<Button-1>", lambda e: browse_excel_file())

    _hr(body, theme)
    _section(body, theme, "辞書データ", "英語変換・ジャンル分類に使う辞書ファイルの場所を指定します。",
             help_text="このフォルダに \"translation_dictionary.xlsx\"(邦題→英題)や \"genre_dictionary.xlsx\"\n"
                        "(アーティスト→ジャンル階層)を置くと、アプリ内蔵の既定辞書の代わりに使われます。\n"
                        "ファイルがない場合は既定の辞書がそのまま使われます。")
    path_vars["genre_rule_folder"] = _folder_field(
        body, theme, "辞書データ フォルダ", app.cfg["genre_rule_folder"], is_dir=True)

    _section(body, theme, "照会サービスのキー",
             "名前照会・ジャンル照会・アルバム写真で Discogs と Spotify を使うときの、キーを書いたテキストファイルです。",
             help_text="ファイルの中に【Discogs】の節(トークン:…)と【Spotify】の節(クライアントID:…・\n"
                        "クライアントシークレット:…)を書いておきます。キーはアプリには保存しません。\n"
                        "未指定のときは Apple Music だけを使います。")
    path_vars["api_keys_file"] = _folder_field(
        body, theme, "キーのファイル", app.cfg.get("api_keys_file", ""), is_dir=False)

    alias_row = tk.Frame(body, bg=theme["bg"])
    alias_row.pack(fill="x", pady=(2, 0))
    alias_status_var = tk.StringVar()
    tk.Label(alias_row, textvariable=alias_status_var, bg=theme["bg"], fg=theme["text_faint"],
              font=(FONT_BODY, 10), anchor="w", justify="left", wraplength=760).pack(side="left", fill="x", expand=True)
    genre_update_btn = w.button(alias_row, "Genre更新", lambda: do_genre_update(), theme, kind="secondary")
    genre_update_btn.pack(side="right")

    def refresh_alias_status():
        folder = path_vars["genre_rule_folder"].get()
        path = os.path.join(folder, ALIAS_FALLBACK_FILENAME) if folder else None
        if path and os.path.exists(path):
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
            alias_status_var.set(f"補助ジャンル辞書: 作成済み(最終更新 {mtime})")
        else:
            alias_status_var.set("補助ジャンル辞書: 未作成(未作成の間はアーティスト辞書のみで判定します)")

    def do_genre_update():
        folder = path_vars["genre_rule_folder"].get()
        if not folder:
            messagebox.showwarning("未設定", "先に辞書データ フォルダを指定してください。")
            return

        def target(cb):
            return generate_alias_fallback(folder, cb)

        def on_done(result):
            messagebox.showinfo(
                "完了",
                f"補助ジャンル辞書を{result['entries']}件生成しました。\n"
                "下の「補助ジャンル辞書を使用する」にチェックを入れると、\n"
                "アーティスト辞書に該当がないアルバムの判定に使われます。",
            )
            refresh_alias_status()

        app.run_task("Genreデータを更新中...", target, on_done)

    fallback_enabled_var = tk.BooleanVar(value=app.cfg.get("genre_alias_fallback_enabled", False))
    tk.Checkbutton(body, text="補助ジャンル辞書を使用する(アーティスト辞書に該当がない場合、MP3のジャンルタグから推定)",
                    variable=fallback_enabled_var, bg=theme["bg"], fg=theme["text"],
                    selectcolor=theme["bg_elevated_2"], activebackground=theme["bg"],
                    font=(FONT_BODY, 11)).pack(anchor="w", pady=(8, 0))

    refresh_alias_status()

    _hr(body, theme)
    _section(body, theme, "MP3タグ設定", "STEP 07 でID3タグを更新するときのバージョンを指定します。",
             help_text="v2.3(UTF-16)は古いカーナビ・オーディオ機器との互換性を重視した形式です。\n"
                        "v2.4(UTF-8)はより新しい標準的な形式で、通常はこちらを選べば問題ありません。")

    id3_version_var = tk.StringVar(value=app.cfg.get("id3_version", "v24"))
    version_row = tk.Frame(body, bg=theme["bg"])
    version_row.pack(fill="x", pady=(4, 8))

    version_pills = {}

    def refresh_version_pills():
        for val, (pill, lbl) in version_pills.items():
            selected = id3_version_var.get() == val
            bg = theme["accent"] if selected else theme["bg_elevated_2"]
            fg = theme["accent_text"] if selected else theme["text_muted"]
            pill.configure(bg=bg)
            lbl.configure(bg=bg, fg=fg)

    def make_version_pill(value, text):
        pill = tk.Frame(version_row, bg=theme["bg_elevated_2"], cursor="hand2")
        pill.pack(side="left", padx=(0, 10))
        lbl = tk.Label(pill, text=text, bg=theme["bg_elevated_2"], fg=theme["text_muted"],
                        font=(FONT_BODY, 11), padx=14, pady=6)
        lbl.pack()

        def on_click(e):
            id3_version_var.set(value)
            refresh_version_pills()

        pill.bind("<Button-1>", on_click)
        lbl.bind("<Button-1>", on_click)
        version_pills[value] = (pill, lbl)

    make_version_pill("v23", "ID3v2.3 (UTF-16)")
    make_version_pill("v24", "ID3v2.4 (UTF-8)")
    refresh_version_pills()

    strip_v1_var = tk.BooleanVar(value=app.cfg.get("strip_id3v1", False))
    tk.Checkbutton(body, text="ID3v1タグ(古い形式)の情報を削除する", variable=strip_v1_var,
                    bg=theme["bg"], fg=theme["text"], selectcolor=theme["bg_elevated_2"],
                    activebackground=theme["bg"], font=(FONT_BODY, 11)).pack(anchor="w", pady=(0, 4))
    tk.Label(body, text="MP3ファイル末尾に付与される旧式のID3v1タグを取り除きます(ID3v2タグ本体はそのまま残ります)。",
              bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 10), anchor="w").pack(anchor="w")

    _hr(body, theme)
    _section(body, theme, "サイトから取り寄せ", "MP3 データベースのサイトから MP3 タグを取り寄せるときの窓口と鍵です。",
             help_text="鍵は、サイトにログインしてマイページの「TagFlow の鍵」で作ります。鍵は人に教えないでください。")
    site_url_var = _text_field(body, theme, "窓口の URL", app.cfg.get("mp3db_api_url") or site_fetch.DEFAULT_URL)
    site_key_var = _text_field(body, theme, "鍵", app.cfg.get("mp3db_api_key", ""), secret=True)

    _hr(body, theme)
    _section(body, theme, "表示テーマ", "画面の色調を選択します。")
    theme_row = tk.Frame(body, bg=theme["bg"])
    theme_row.pack(fill="x", pady=(4, 0))
    _theme_swatch(theme_row, theme, "dark", "ダーク", app, selected=(app.cfg["theme"] == "dark"))
    _theme_swatch(theme_row, theme, "light", "ライト", app, selected=(app.cfg["theme"] == "light"))

    _hr(body, theme)
    _section(body, theme, "動作ログ", "各処理の開始・完了・エラーを記録したログファイルです。トラブル時の確認に使えます。")
    log_row = tk.Frame(body, bg=theme["bg"])
    log_row.pack(fill="x", pady=(4, 0))
    tk.Label(log_row, text=LOG_PATH, bg=theme["bg"], fg=theme["text_faint"],
              font=(FONT_MONO, 9), anchor="w").pack(side="left", fill="x", expand=True)

    def open_log():
        if not os.path.exists(LOG_PATH):
            messagebox.showinfo("ログなし", "まだログファイルが作成されていません。")
            return
        open_path(LOG_PATH)

    log_open_lbl = tk.Label(log_row, text="開く", bg=theme["bg"], fg=theme["accent"],
                              font=(FONT_BODY, 10), cursor="hand2")
    log_open_lbl.pack(side="right")
    log_open_lbl.bind("<Button-1>", lambda e: open_log())

    # footer lives in `outer`, below `scroll_area`, so it stays visible
    # regardless of scroll position
    _hr(outer, theme)
    footer = tk.Frame(outer, bg=theme["bg"])
    footer.pack(fill="x", pady=(10, 0))

    def do_save():
        for key, var in path_vars.items():
            app.cfg[key] = var.get()
        app.cfg["id3_version"] = id3_version_var.get()
        app.cfg["strip_id3v1"] = strip_v1_var.get()
        app.cfg["genre_alias_fallback_enabled"] = fallback_enabled_var.get()
        app.cfg["mp3db_api_url"] = site_url_var.get().strip()
        app.cfg["mp3db_api_key"] = site_key_var.get().strip()
        save_config(app.cfg)
        messagebox.showinfo("保存しました", "設定を保存しました。")
        app.show_dashboard()

    w.button(footer, "変更を保存", do_save, theme, kind="primary").pack(side="right")
    cancel_lbl = tk.Label(footer, text="キャンセル", bg=theme["bg"], fg=theme["text_muted"],
                            font=(FONT_BODY, 11), cursor="hand2")
    cancel_lbl.pack(side="right", padx=(0, 14))
    cancel_lbl.bind("<Button-1>", lambda e: app.show_dashboard())


def _section(parent, theme, title, desc, help_text=None):
    box = tk.Frame(parent, bg=theme["bg"])
    box.pack(fill="x", pady=(0, 8))
    title_row = tk.Frame(box, bg=theme["bg"])
    title_row.pack(fill="x")
    tk.Label(title_row, text=title, bg=theme["bg"], fg=theme["text"], font=(FONT_DISPLAY, 12, "bold"),
              anchor="w").pack(side="left")
    if help_text:
        w.help_badge(title_row, theme, help_text).pack(side="left", padx=(6, 0))
    tk.Label(box, text=desc, bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 10),
              anchor="w").pack(fill="x")


def _hr(parent, theme):
    f = tk.Frame(parent, bg=theme["border"], height=1)
    f.pack(fill="x", pady=10)


def _folder_field(parent, theme, label_text, initial, is_dir=True):
    box = tk.Frame(parent, bg=theme["bg"])
    box.pack(fill="x", pady=(0, 6))
    tk.Label(box, text=label_text, bg=theme["bg"], fg=theme["text_muted"], font=(FONT_BODY, 10),
              anchor="w").pack(fill="x")

    row = tk.Frame(box, bg=theme["bg_elevated"], highlightbackground=theme["border"], highlightthickness=1)
    row.pack(fill="x", pady=(3, 0))

    var = tk.StringVar(value=initial)
    entry = tk.Entry(row, textvariable=var, bg=theme["bg_elevated"], fg=theme["text"],
                      insertbackground=theme["text"], font=(FONT_MONO, 10), relief="flat", bd=0)
    entry.pack(side="left", fill="x", expand=True, padx=10, pady=6)

    def browse():
        if is_dir:
            path = filedialog.askdirectory(initialdir=var.get() or "/")
        else:
            path = filedialog.askopenfilename(initialdir=os.path.dirname(var.get()) if var.get() else "/")
        if path:
            var.set(path)

    browse_lbl = tk.Label(row, text="参照", bg=theme["bg_elevated"], fg=theme["accent"], font=(FONT_BODY, 10),
                            cursor="hand2", padx=14)
    browse_lbl.pack(side="right")
    browse_lbl.bind("<Button-1>", lambda e: browse())

    return var


def _text_field(parent, theme, label_text, initial, secret=False):
    box = tk.Frame(parent, bg=theme["bg"])
    box.pack(fill="x", pady=(0, 6))
    tk.Label(box, text=label_text, bg=theme["bg"], fg=theme["text_muted"], font=(FONT_BODY, 10),
              anchor="w").pack(fill="x")
    row = tk.Frame(box, bg=theme["bg_elevated"], highlightbackground=theme["border"], highlightthickness=1)
    row.pack(fill="x", pady=(3, 0))
    var = tk.StringVar(value=initial)
    tk.Entry(row, textvariable=var, show="•" if secret else "", bg=theme["bg_elevated"], fg=theme["text"],
              insertbackground=theme["text"], font=(FONT_MONO, 10), relief="flat", bd=0).pack(
        side="left", fill="x", expand=True, padx=10, pady=6)
    return var


def _theme_swatch(parent, theme, key, label_text, app, selected):
    from theme import get_theme
    preview_theme = get_theme(key)

    box = tk.Frame(parent, bg=theme["bg"])
    box.pack(side="left", padx=(0, 16))

    border_color = theme["accent"] if selected else theme["border"]
    swatch = tk.Frame(box, bg=preview_theme["bg"], width=108, height=64,
                       highlightbackground=border_color, highlightthickness=2 if selected else 1)
    swatch.pack()
    swatch.pack_propagate(False)
    inner = tk.Frame(swatch, bg=preview_theme["bg_elevated"], width=88, height=20)
    inner.place(x=10, y=9)
    dot = tk.Canvas(inner, width=8, height=8, bg=preview_theme["bg_elevated"], highlightthickness=0)
    dot.create_oval(1, 1, 7, 7, fill=preview_theme["accent"], outline="")
    dot.place(x=6, y=6)
    tk.Frame(swatch, bg=preview_theme["text_faint"], width=54, height=4).place(x=10, y=36)
    tk.Frame(swatch, bg=preview_theme["border"], width=36, height=4).place(x=10, y=44)

    swatch.bind("<Button-1>", lambda e: app.set_theme(key))
    swatch.configure(cursor="hand2")

    lbl_color = theme["text_muted"] if selected else theme["text_faint"]
    tk.Label(box, text=label_text, bg=theme["bg"], fg=lbl_color, font=(FONT_BODY, 10)).pack(pady=(4, 0))
