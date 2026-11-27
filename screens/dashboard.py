# -*- coding: utf-8 -*-
import os
import sys
import datetime
import tkinter as tk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import extract as backend_extract
from backend import translate as backend_translate
import shutil

from PIL import Image, ImageTk

ASSETS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
LOGO_PATHS = {
    "dark": os.path.join(ASSETS_DIR, "logo_dark.png"),   # white mark, shown on the dark theme
    "light": os.path.join(ASSETS_DIR, "logo_light.png"),  # black mark, shown on the light theme
}
APP_ICON_PATH = os.path.join(ASSETS_DIR, "app_icon_1024.png")
LOGO_HEIGHT = 140  # both marks are trimmed of transparent margins and shown at this height


def read_stats(excel_path):
    if not excel_path or not os.path.exists(excel_path):
        return None
    import openpyxl
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb["IDTAG"]
    tracks = 0
    albums = set()
    artists = set()
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if row[2] is None and row[3] is None:
            continue
        tracks += 1
        if row[4]:
            albums.add(row[4])
        aa = row[12] or row[3]
        if aa:
            artists.add(aa)
    return {"tracks": tracks, "albums": len(albums), "artists": len(artists)}


STEP_DEFS = [
    ("cleanup", "小容量ファイル検出・削除", "ダウンロードやコピーでできた不要なファイル(Mac の補足ファイル・壊れた小さいファイルなど)を探し、選んで削除します。", "teal"),
    ("extract", "MP3タグ抽出", "指定フォルダ内のMP3を再帰的にスキャンし、ID3タグをExcelへ書き出します。", "accent"),
    ("translate", "英語表記への変換", "邦題を登録済みの変換辞書で英題に変換し、Title Case表記に統一します。", "green"),
    ("albumartist", "アルバムアーティスト推定", "アルバムごとに、単独・複数並記・主要アーティスト・Various Artistsを判定し、要確認のものを選びます。", "yellow"),
    ("genre", "ジャンル分類", "アルバムごとに Discogs のジャンル＞スタイルを決めます。タグにはスタイル名を書きます。", "teal"),
    ("artwork", "アルバム写真の整合性確認", "Apple Music・Spotify・Discogs の画像と照合し、未登録・不一致のアルバム写真を検出・修正します。(任意)", "cyan"),
    ("integrity", "ID3 タグ情報確認・適用", "変更点・警告を確認し、確認した内容をMP3のID3タグに書き込みます。", "yellow"),
    ("organize", "適用ファイル操作", "ファイル名をトラック名に変えて、アーティスト別フォルダへコピーします(元データの削除も選べます)。", "red"),
]

HELP_TEXT = {
    "cleanup": "指定サイズ以下のファイル(アートワークの断片やキャッシュファイルなど)を検出します。"
               "削除対象は自分で選んで実行するまで何も削除されません。",
    "extract": "MP3のID3タグを読み取り、Excelに一覧化します。ここで作成したExcelが以降すべてのSTEPの作業データになります。",
    "translate": "登録済みの変換辞書に一致した邦題のみ英語に変換します。辞書にない語は赤字で「未解決」として残ります。"
                 "「名前照会」で、残った日本語を Apple・Spotify・Discogs に照会できます。",
    "albumartist": "アルバム(フォルダ)単位で判定します。過半数の曲に参加するアーティストがいればそのアーティスト、"
                   "複数並記は各アーティストか並記のままかを選択、過半数に届かなければVarious Artistsか参加者から選びます。",
    "genre": "Discogs のジャンル＞スタイル(独自スタイルを含む)からアルバムごとに選びます。候補は Excel の設定済みの値、"
             "アーティストの既定、Discogs・Apple Music・Spotify への照会の順。辞書は data/genre_styles.xlsx です。",
    "integrity": "元データとの変更点や、曲順・発売日・ファイル作成日の不整合を確認し、作業用Excelの内容を"
                 "MP3のID3タグに書き込みます。タグを更新するまでSTEP08は実行できません。",
    "artwork": "MP3に埋め込まれた画像と、選んだサービス(Apple Music・Spotify・Discogs)の画像を比較し、大きく異なる場合のみ警告します"
               "(解像度違いなど小さな差は無視されます)。反映しない限り何も変更されません。",
    "organize": "STEP07でタグを更新したファイルを、ファイル名を変えてアルバムアーティスト別フォルダへコピーします。"
                "既定では元データは削除されません(オプションで削除も選べます)。",
}


def build(parent, app):
    theme = app.theme
    parent.configure(bg=theme["bg"])

    outer = tk.Frame(parent, bg=theme["bg"])
    outer.pack(fill="both", expand=True, padx=28, pady=14)

    # header
    header = tk.Frame(outer, bg=theme["bg"])
    header.pack(fill="x", pady=(0, 10))

    mark = tk.Frame(header, bg=theme["accent"], width=32, height=32)
    mark.pack(side="left")
    mark.pack_propagate(False)
    tk.Label(mark, text="♪", bg=theme["accent"], fg=theme["bg"], font=(FONT_DISPLAY, 14, "bold")).pack(expand=True)

    title_box = tk.Frame(header, bg=theme["bg"])
    title_box.pack(side="left", padx=(10, 0))
    tk.Label(title_box, text="Tag FloW", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 20, "bold")).pack(anchor="w")
    tk.Label(title_box, text="MP3 ライブラリ管理", bg=theme["bg"], fg=theme["text_muted"],
              font=(FONT_BODY, 11)).pack(anchor="w")

    stats = read_stats(app.get_excel_path())
    stats_box = tk.Frame(header, bg=theme["bg"])
    stats_box.pack(side="right")
    if stats:
        _stat(stats_box, theme, str(stats["tracks"]), "曲")
        _stat(stats_box, theme, str(stats["albums"]), "アルバム")
        _stat(stats_box, theme, str(stats["artists"]), "アーティスト")
    else:
        tk.Label(stats_box, text="設定：初期設定", bg=theme["bg"], fg=theme["text_faint"],
                  font=(FONT_BODY, 11)).pack(side="left", padx=10)

    # サイト（MP3 データベース）から取り寄せる別の道（2026-11-25）。Excel の STEP の外に置く
    fetch_btn = w.button(header, "⇩ サイトから取り寄せる", app.show_site_fetch, theme, kind="secondary")

    settings_btn = tk.Frame(header, bg=theme["bg_elevated_2"], width=32, height=32,
                             highlightbackground=theme["border"], highlightthickness=1)
    settings_btn.pack(side="right", padx=(14, 0))
    fetch_btn.pack(side="right", padx=(14, 0))
    settings_btn.pack_propagate(False)
    lbl = tk.Label(settings_btn, text="≡", bg=theme["bg_elevated_2"], fg=theme["text_muted"],
                    font=(FONT_BODY, 16))
    lbl.pack(expand=True)
    for widget in (settings_btn, lbl):
        widget.bind("<Button-1>", lambda e: app.show_settings())
        widget.configure(cursor="hand2")

    w.button(header, "タグ一覧編集", lambda: app.show_tag_table(), theme, kind="secondary").pack(side="right", padx=(14, 0))

    # grid of cards
    grid = tk.Frame(outer, bg=theme["bg"])
    grid.pack(fill="both", expand=True)
    for c in range(3):
        grid.columnconfigure(c, weight=1, uniform="col")

    for idx, (key, name, desc, color_key) in enumerate(STEP_DEFS):
        r, c = divmod(idx, 3)
        _build_card(grid, app, theme, idx + 1, key, name, desc, color_key, r, c)

    # the 3-column grid has one empty cell after the last card (8 steps -> row 2, col 2) --
    # a natural, unobtrusive spot for a corner logo
    last_row, last_col = divmod(len(STEP_DEFS), 3)
    _build_logo(grid, app, theme, last_row, last_col)


def _build_logo(grid, app, theme, row, column):
    theme_name = app.cfg.get("theme", "dark")
    photos = []
    # Tag FloW's app icon first, then the Freebear Factory logo
    for path in (APP_ICON_PATH, LOGO_PATHS.get(theme_name)):
        if not path or not os.path.exists(path):
            continue
        try:
            img = Image.open(path).convert("RGBA")
            img = img.crop(img.getchannel("A").getbbox())
            img = img.resize((round(img.width * LOGO_HEIGHT / img.height), LOGO_HEIGHT), Image.LANCZOS)
            photos.append(ImageTk.PhotoImage(img))
        except Exception:
            continue
    if not photos:
        return
    app._dashboard_logo_ref = photos  # keep references alive past this function

    cell = tk.Frame(grid, bg=theme["bg"])
    cell.grid(row=row, column=column, sticky="nsew", padx=6, pady=6)
    holder = tk.Frame(cell, bg=theme["bg"])
    holder.place(relx=1.0, rely=1.0, anchor="se")
    for photo in photos:
        tk.Label(holder, image=photo, bg=theme["bg"]).pack(side="left", padx=(8, 0))


def _stat(parent, theme, value, label):
    box = tk.Frame(parent, bg=theme["bg"])
    box.pack(side="left", padx=10)
    tk.Label(box, text=value, bg=theme["bg"], fg=theme["text"], font=(FONT_DISPLAY, 15, "bold")).pack()
    tk.Label(box, text=label, bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 10)).pack()


def _build_card(grid, app, theme, step_no, key, name, desc, color_key, r, c):
    outer_pad = tk.Frame(grid, bg=theme["bg"])
    outer_pad.grid(row=r, column=c, sticky="nsew", padx=6, pady=6)

    card = w.card(outer_pad, theme)
    card.pack(fill="both", expand=True)

    inner = tk.Frame(card, bg=theme["bg_elevated"])
    inner.pack(fill="both", expand=True, padx=16, pady=12)

    top_row = tk.Frame(inner, bg=theme["bg_elevated"])
    top_row.pack(fill="x")
    tk.Label(top_row, text=f"STEP {step_no:02d}", bg=theme["bg_elevated"], fg=theme["text_faint"],
              font=(FONT_MONO, 10)).pack(side="left")
    if key in HELP_TEXT:
        w.help_badge(top_row, theme, HELP_TEXT[key]).pack(side="left", padx=(6, 0))

    done = key in app.completed_steps
    blocked = key == "organize" and "integrity_check" not in app.completed_steps
    if done:
        status_text, status_color = "完了", "green"
    elif blocked:
        status_text, status_color = "STEP 07待ち", "text_faint"
    elif key == "cleanup":
        status_text, status_color = "確認可能", "text_faint"
    elif key == "artwork":
        status_text, status_color = "任意", "text_faint"
    else:
        status_text, status_color = "未実行", "text_faint"
    tk.Label(top_row, text=status_text, bg=theme["bg_elevated"],
              fg=theme[status_color] if status_color != "text_faint" else theme["text_faint"],
              font=(FONT_BODY, 10)).pack(side="right")

    tk.Label(inner, text=name, bg=theme["bg_elevated"], fg=theme["text"],
              font=(FONT_DISPLAY, 13, "bold"), anchor="w").pack(fill="x", pady=(6, 3))
    tk.Label(inner, text=w.wrap_text(desc, (FONT_BODY, 10), 280), bg=theme["bg_elevated"],
              fg=theme["text_muted"], font=(FONT_BODY, 10), justify="left", anchor="w").pack(fill="x")

    btn_frame = tk.Frame(inner, bg=theme["bg_elevated"])
    btn_frame.pack(fill="x", pady=(8, 0), side="bottom")

    action = _make_action(app, key)
    if key in ("integrity", "artwork"):
        btn_text = "確認する" if not done else "再確認"
    else:
        btn_text = "確認する" if key == "cleanup" else ("実行" if not done else "再実行")
    btn = w.button(btn_frame, btn_text, action, theme, kind="primary")
    btn.pack(side="left")
    if blocked:
        btn.configure(state="disabled")

    if key == "translate":
        verify_btn = w.button(btn_frame, "Web再調査", lambda: app.show_verify(), theme, kind="secondary")
        verify_btn.pack(side="left", padx=(8, 0))
        lookup_btn = w.button(btn_frame, "名前照会", lambda: app.show_name_lookup(), theme, kind="secondary")
        lookup_btn.pack(side="left", padx=(8, 0))

    result = app.last_results.get(key)
    if result:
        note = _result_note(key, result)
        if note:
            tk.Label(btn_frame, text=note, bg=theme["bg_elevated"], fg=theme["text_faint"],
                      font=(FONT_MONO, 9)).pack(side="left", padx=(10, 0))


def _result_note(key, result):
    if key == "extract":
        return f"{result.get('total', 0)}曲"
    if key == "translate":
        note = f"変換{result.get('converted', 0)}件"
        if result.get("unresolved"):
            note += f" ・未解決{result['unresolved']}件"
        return note
    if key == "albumartist":
        return f"設定{result.get('set', 0)}件"
    if key == "genre":
        note = f"設定{result.get('set', 0)}件"
        if result.get("unresolved"):
            note += f" ・未解決{result['unresolved']}件"
        return note
    if key == "integrity":
        return f"変更{result.get('changes', 0)}件・警告{result.get('warnings', 0)}件"
    if key == "artwork":
        return f"反映{result.get('applied', 0)}件"
    if key == "organize":
        return f"{result.get('folders_copied', 0)}フォルダ"
    return ""


def _make_action(app, key):
    if key == "cleanup":
        return lambda: app.show_cleanup()

    if key == "extract":
        def action():
            src = app.cfg.get("source_folder")
            excel_folder = app.cfg.get("excel_folder")
            if not src or not excel_folder:
                _warn(app, "先に設定画面で「MP3保存元フォルダ」と「Excel保存先フォルダ」を指定してください。")
                return
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(excel_folder, f"TagFloW_{ts}.xlsx")

            def target(cb):
                return backend_extract.run(src, out_path, cb)

            def on_done(result):
                app.set_excel_path(result["excel_path"])
                base, ext = os.path.splitext(result["excel_path"])
                baseline_path = f"{base}_baseline{ext}"
                shutil.copy2(result["excel_path"], baseline_path)
                app.baseline_excel_path = baseline_path
                app.completed_steps.add("extract")
                # a fresh extraction invalidates every downstream step's results
                # (last_results uses "integrity" as its key; completed_steps uses "integrity_check")
                for key in ("translate", "albumartist", "genre", "integrity", "artwork", "organize"):
                    app.last_results.pop(key, None)
                for key in ("translate", "albumartist", "genre", "integrity_check", "artwork", "organize"):
                    app.completed_steps.discard(key)
                app.artwork_overrides = {}
                app.artwork_scan_results = None
                app.genre_overrides = {}
                app.albumartist_choices = {}
                app.last_results["extract"] = result
                app.refresh()

            app.run_task("MP3タグ抽出中...", target, on_done)
        return action

    if key == "translate":
        def action():
            path = app.get_excel_path()
            if not path:
                _warn(app, "先に STEP 02 (MP3タグ抽出) を実行してください。")
                return

            def target(cb):
                return backend_translate.run(path, cb, dict_folder=app.cfg.get("genre_rule_folder"))

            def on_done(result):
                app.completed_steps.add("translate")
                app.completed_steps.discard("integrity_check")
                app.last_results["translate"] = result
                app.refresh()

            app.run_task("英語表記へ変換中...", target, on_done)
        return action

    if key == "albumartist":
        def action():
            if not app.get_excel_path():
                _warn(app, "先に STEP 02 (MP3タグ抽出) を実行してください。")
                return
            app.show_albumartist()
        return action

    if key == "genre":
        def action():
            path = app.get_excel_path()
            if not path:
                _warn(app, "先に STEP 02 (MP3タグ抽出) を実行してください。")
                return
            app.show_genre()
        return action

    if key == "integrity":
        return lambda: app.show_integrity()

    if key == "artwork":
        def action():
            path = app.get_excel_path()
            if not path:
                _warn(app, "先に STEP 02 (MP3タグ抽出) を実行してください。")
                return
            app.show_artwork()
        return action

    if key == "organize":
        def action():
            path = app.get_excel_path()
            if not path:
                _warn(app, "先に STEP 02 (MP3タグ抽出) を実行してください。")
                return
            if "integrity_check" not in app.completed_steps:
                _warn(app, "先に STEP 07 (ID3 タグ情報確認・適用) でID3タグを更新してください。")
                return
            app.show_organize()
        return action

    return lambda: None


def _warn(app, message):
    from tkinter import messagebox
    messagebox.showwarning("未実行", message)
