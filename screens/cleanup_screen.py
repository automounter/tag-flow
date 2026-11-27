# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
import unicodedata
from tkinter import messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from platform_util import open_path
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import cleanup as backend_cleanup


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
    tk.Label(title_row, text="小容量ファイル検出・削除", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "ダウンロードやコピーのときにできた不要なファイルを探します。\n"
                 "・Mac の補足ファイル(._ で始まる)、中身が音声でない小さいファイル、\n"
                 "  ダウンロードの付属ファイル(歌詞 .lrc など)は最初からチェック済み\n"
                 "・小さいけれど再生できる MP3 と、MP3 以外のファイルは一覧に出すだけ(必要なら選ぶ)\n"
                 "チェックを入れて削除を押すまで何も削除されません。"
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text=f"STEP 01 ・ {app.cfg.get('source_folder') or '(未設定)'}", bg=theme["bg"],
              fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    if not app.cfg.get("source_folder"):
        tk.Label(outer, text="先に設定画面で「MP3保存元フォルダ」を指定してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return

    scan_btn = w.button(header, "再スキャン", lambda: rescan(), theme, kind="secondary")
    scan_btn.pack(side="right")

    jp_folder_var = tk.StringVar(value="")
    jp_folder_lbl = tk.Label(header, textvariable=jp_folder_var, bg=theme["bg"], fg=theme["yellow"],
                              font=(FONT_BODY, 10, "bold"))
    jp_folder_lbl.pack(side="right", padx=(0, 10))

    def show_japanese_folders(paths):
        win = tk.Toplevel(parent)
        win.title(f"日本語フォルダ({len(paths)}件)")
        win.geometry("820x420")
        win.configure(bg=theme["bg"])
        tk.Label(win, text="ダブルクリックで Finder に表示します", bg=theme["bg"], fg=theme["text_muted"],
                  font=(FONT_BODY, 10)).pack(anchor="w", padx=12, pady=(10, 6))
        frame = tk.Frame(win, bg=theme["bg"])
        frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        listbox = tk.Listbox(frame, bg=theme["bg_elevated"], fg=theme["text"], font=(FONT_BODY, 11),
                              activestyle="none", highlightthickness=0)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        for path in paths:
            listbox.insert("end", unicodedata.normalize("NFC", path))

        def open_selected(event):
            selection = listbox.curselection()
            if selection:
                open_path(paths[selection[0]])

        listbox.bind("<Double-Button-1>", open_selected)

    def check_japanese_folder(result, error):
        if error or not result:
            return
        jp_folder_var.set(f"日本語フォルダ有り({len(result)}件)")
        jp_folder_lbl.configure(cursor="hand2")
        jp_folder_lbl.bind("<Button-1>", lambda e: show_japanese_folders(result))

    w.run_async(header, lambda: backend_cleanup.find_japanese_subfolders(app.cfg.get("source_folder")),
                check_japanese_folder)

    status_var = tk.StringVar(value="スキャン中...")
    status_lbl = tk.Label(outer, textvariable=status_var, bg=theme["bg"], fg=theme["text"],
                           font=(FONT_BODY, 11), anchor="w")
    status_lbl.pack(fill="x", pady=(0, 10))

    toolbar = tk.Frame(outer, bg=theme["bg"])
    toolbar.pack(fill="x", pady=(0, 8))

    tk.Label(toolbar, text="小さいファイルの基準:", bg=theme["bg"], fg=theme["text_muted"],
              font=(FONT_BODY, 11)).pack(side="left")
    threshold_var = tk.StringVar(value=str(app.cfg.get("small_file_threshold_kb", 10)))
    tk.Spinbox(toolbar, from_=1, to=1024, width=5, textvariable=threshold_var,
               font=(FONT_BODY, 11)).pack(side="left", padx=(4, 2))
    tk.Label(toolbar, text="KB 以下", bg=theme["bg"], fg=theme["text_muted"],
              font=(FONT_BODY, 11)).pack(side="left", padx=(0, 16))

    select_all_var = tk.BooleanVar(value=False)
    select_all_chk = tk.Checkbutton(toolbar, text="すべて選択", variable=select_all_var,
                                     bg=theme["bg"], fg=theme["text_muted"], selectcolor=theme["bg_elevated_2"],
                                     activebackground=theme["bg"], font=(FONT_BODY, 11),
                                     command=lambda: toggle_all())
    select_all_chk.pack(side="left")

    selected_count_var = tk.StringVar(value="選択中: 0件")
    tk.Label(toolbar, textvariable=selected_count_var, bg=theme["bg"], fg=theme["text_faint"],
              font=(FONT_BODY, 10)).pack(side="right", padx=(0, 14))

    delete_btn = w.button(toolbar, "選択したファイルを削除", lambda: do_delete(), theme, kind="danger")
    delete_btn.pack(side="right")
    delete_btn.configure(state="disabled")

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

    state = {"files": [], "vars": {}}

    def update_selected_count():
        n = sum(1 for v in state["vars"].values() if v.get())
        selected_count_var.set(f"選択中: {n}件")
        delete_btn.configure(state="normal" if n > 0 else "disabled")

    def toggle_all():
        val = select_all_var.get()
        for v in state["vars"].values():
            v.set(val)
        update_selected_count()

    def render_list():
        for c in scroll_frame.winfo_children():
            c.destroy()
        state["vars"] = {}

        files = state["files"]
        if not files:
            tk.Label(scroll_frame, text="不要なファイルは見つかりませんでした。", bg=theme["bg_elevated"],
                      fg=theme["text_faint"], font=(FONT_BODY, 11)).pack(pady=30)
            return

        by_cat = {}
        for f in files:
            by_cat.setdefault(f["category"], []).append(f)

        for key, heading, checked in backend_cleanup.CATEGORIES:
            items = by_cat.get(key)
            if not items:
                continue
            ch = tk.Frame(scroll_frame, bg=theme["bg_elevated"])
            ch.pack(fill="x", padx=10, pady=(14, 2))
            tk.Label(ch, text=f"{heading} ・ {len(items)}件" + ("" if checked else "(最初はチェックなし)"),
                      bg=theme["bg_elevated"], fg=theme["red"] if checked else theme["text"],
                      font=(FONT_BODY, 11, "bold")).pack(side="left")
            album = None
            for item in items:
                if item["album"] != album:
                    album = item["album"]
                    tk.Label(scroll_frame, text=album, bg=theme["bg_elevated"], fg=theme["text_muted"],
                              font=(FONT_BODY, 10), anchor="w").pack(fill="x", padx=24, pady=(4, 0))
                row = tk.Frame(scroll_frame, bg=theme["bg_elevated"])
                row.pack(fill="x", padx=24)
                var = tk.BooleanVar(value=checked)
                state["vars"][item["path"]] = var
                tk.Checkbutton(row, variable=var, bg=theme["bg_elevated"], selectcolor=theme["bg_elevated_2"],
                                activebackground=theme["bg_elevated"], command=update_selected_count).pack(side="left")
                tk.Label(row, text=item["name"], bg=theme["bg_elevated"], fg=theme["text"],
                          font=(FONT_MONO, 10), anchor="w").pack(side="left", fill="x", expand=True, padx=(6, 0))
                tk.Label(row, text=f"{item['size'] / 1024:.1f} KB", bg=theme["bg_elevated"], fg=theme["text_faint"],
                          font=(FONT_MONO, 9)).pack(side="right")

    def rescan():
        status_var.set("スキャン中...(ネットワークフォルダのため時間がかかる場合があります)")
        scan_btn.configure(state="disabled")
        for c in scroll_frame.winfo_children():
            c.destroy()
        try:
            threshold = max(1, int(threshold_var.get()))
        except ValueError:
            threshold = 10
        threshold_var.set(str(threshold))
        if app.cfg.get("small_file_threshold_kb") != threshold:
            app.cfg["small_file_threshold_kb"] = threshold
            from config import save_config
            save_config(app.cfg)
        src = app.cfg["source_folder"]

        w.run_async(outer, lambda: backend_cleanup.scan_small_files(src, threshold), on_scan_done)

    def on_scan_done(results, error):
        scan_btn.configure(state="normal")
        if error:
            status_var.set(f"スキャンに失敗しました: {error}")
            return
        state["files"] = results
        albums = len(set(f["album"] for f in results))
        if results:
            checked = {key for key, _, on in backend_cleanup.CATEGORIES if on}
            pre = sum(1 for f in results if f["category"] in checked)
            status_var.set(f"不要とみられるファイルなど {len(results)}件 が見つかりました({albums}フォルダ)。"
                           f"そのうち {pre}件 に最初からチェックを入れています。")
        else:
            status_var.set("不要なファイルは見つかりませんでした。")
        render_list()
        select_all_var.set(False)
        update_selected_count()

    def do_delete():
        targets = [p for p, v in state["vars"].items() if v.get()]
        if not targets:
            return
        if not messagebox.askyesno("削除の確認", f"選択した {len(targets)}件のファイルを削除します。よろしいですか？"):
            return
        deleted, errors = backend_cleanup.delete_files(targets)
        if errors:
            messagebox.showwarning("一部エラー", f"{len(deleted)}件削除、{len(errors)}件失敗しました。")
        else:
            messagebox.showinfo("削除完了", f"{len(deleted)}件のファイルを削除しました。")
        rescan()

    rescan()
