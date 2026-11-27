# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from platform_util import open_path
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import organize as backend_organize
from logger import LOG_PATH


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
    tk.Label(title_row, text="適用ファイル操作", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "STEP 07 でID3タグを更新したファイルを、①ファイル名を変えて②出力先へコピーします。\n"
                 "③は既定でオフです。今回対象のフォルダのみを処理し、他の整理済みアルバムには触れません。\n"
                 "「アルバム/Disc 1」「アルバム/Disc 2」のようにディスクをサブフォルダで分けている場合は、\n"
                 "同じアルバムアーティスト・アルバム名を持つフォルダを1つにまとめて整理します。"
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text="STEP 08 ・ ファイル名を変更し、アーティスト別フォルダへコピーします",
              bg=theme["bg"], fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    body = tk.Frame(outer, bg=theme["bg"])
    body.pack(fill="both", expand=True)

    excel_path = app.get_excel_path()
    if not excel_path:
        tk.Label(body, text="先に STEP 02 (MP3タグ抽出) を実行してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return
    if "integrity_check" not in app.completed_steps:
        tk.Label(body, text="先に STEP 07 (ID3 タグ情報確認・適用) で「⑤ ID3タグを更新」を実行してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return
    if not app.cfg.get("output_folder"):
        tk.Label(body, text="先に設定画面で「出力先フォルダ」を指定してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return

    validation = backend_organize.validate_for_organize(excel_path)
    missing_title_n = len(validation["missing_title"])
    missing_track_n = len(validation["missing_track"])
    if missing_title_n or missing_track_n:
        warn_card = w.card(body, theme)
        warn_card.pack(fill="x", pady=(0, 16))
        warn_inner = tk.Frame(warn_card, bg=theme["bg_elevated"])
        warn_inner.pack(fill="x", padx=22, pady=16)
        tk.Label(warn_inner, text="タイトル・トラック番号が不足しているファイルがあります",
                  bg=theme["bg_elevated"], fg=theme["yellow"], font=(FONT_BODY, 11, "bold"),
                  anchor="w").pack(fill="x")
        tk.Label(warn_inner,
                  text=f"タイトル未設定: {missing_title_n}件 ・ トラック番号未設定: {missing_track_n}件"
                       f"(全{validation['total']}件中)\n"
                       "該当ファイルもコピーされますが、①のファイル名変更の対象外となり、"
                       "元のファイル名のまま残ります。",
                  bg=theme["bg_elevated"], fg=theme["text_faint"], font=(FONT_BODY, 10),
                  anchor="w", justify="left", wraplength=760).pack(fill="x", pady=(4, 0))

    delete_var = tk.BooleanVar(value=False)

    option_card = w.card(body, theme)
    option_card.pack(fill="x", pady=(0, 20))
    opt_inner = tk.Frame(option_card, bg=theme["bg_elevated"])
    opt_inner.pack(fill="x", padx=22, pady=18)

    for step_text in ("①ファイル名の変更(<ディスク>_<トラック>_<曲名>.mp3)", "②アーティスト別フォルダに整理(コピー)"):
        tk.Label(opt_inner, text=step_text, bg=theme["bg_elevated"], fg=theme["text"],
                  font=(FONT_BODY, 11), anchor="w").pack(anchor="w", pady=(0, 4))
    tk.Checkbutton(opt_inner, text="③元のデータを削除する(移動後、元のファイル・フォルダを削除)", variable=delete_var,
                    bg=theme["bg_elevated"], fg=theme["text"], selectcolor=theme["bg_elevated_2"],
                    activebackground=theme["bg_elevated"], font=(FONT_BODY, 11), anchor="w",
                    command=lambda: update_warning()).pack(anchor="w", pady=(4, 0))

    warn_var = tk.StringVar(value="")
    warn_lbl = tk.Label(opt_inner, textvariable=warn_var, bg=theme["bg_elevated"], fg=theme["text_faint"],
                          font=(FONT_BODY, 10), wraplength=760, justify="left", anchor="w")
    warn_lbl.pack(anchor="w", pady=(10, 0), fill="x")

    def update_warning():
        if delete_var.get():
            warn_var.set("コピーした後、MP3保存元フォルダから該当ファイル・フォルダが完全に削除されます。"
                          "実行後は元に戻せません。")
            warn_lbl.configure(fg=theme["red"])
        else:
            warn_var.set("元のファイル・フォルダは残り、出力先にコピーが作成されます。")
            warn_lbl.configure(fg=theme["yellow"])

    update_warning()

    conflict_box = tk.Frame(body, bg=theme["bg"])
    conflict_box.pack(fill="both", expand=True, pady=(0, 16))

    result_box = tk.Frame(body, bg=theme["bg"])
    result_box.pack(fill="both", expand=True)

    def render_result(result):
        for c in result_box.winfo_children():
            c.destroy()
        card = w.card(result_box, theme)
        card.pack(fill="x")
        inner = tk.Frame(card, bg=theme["bg_elevated"])
        inner.pack(fill="x", padx=24, pady=20)

        stat_row = tk.Frame(inner, bg=theme["bg_elevated"])
        stat_row.pack(fill="x")
        _stat(stat_row, theme, str(result.get("folders_copied", 0)), "フォルダ")
        _stat(stat_row, theme, str(result.get("renamed", 0)), "リネーム対象フォルダ")
        _stat(stat_row, theme, str(result.get("originals_deleted", 0)), "元データ削除")

        errors = (result.get("copy_errors") or []) + (result.get("delete_errors") or [])
        if errors:
            tk.Label(inner, text=f"エラー {len(errors)}件が発生しました。詳細はログを確認してください。",
                      bg=theme["bg_elevated"], fg=theme["red"], font=(FONT_BODY, 10)).pack(anchor="w", pady=(12, 0))

        rename_notes = result.get("rename_errors") or []
        if rename_notes:
            tk.Label(inner, text=f"ファイル名変更に関する注記 {len(rename_notes)}件"
                                  "(タイトル/トラック番号不足によるスキップ、リネーム失敗などを含む)。"
                                  "詳細はログを確認してください。",
                      bg=theme["bg_elevated"], fg=theme["yellow"], font=(FONT_BODY, 10)).pack(anchor="w", pady=(6, 0))

        def open_log():
            if not os.path.exists(LOG_PATH):
                messagebox.showinfo("ログなし", "まだログファイルが作成されていません。")
                return
            open_path(LOG_PATH)

        w.button(inner, "ログを開く", open_log, theme, kind="secondary").pack(anchor="w", pady=(14, 0))

    def do_run(folder_actions=None, file_actions=None):
        if delete_var.get():
            ok = messagebox.askyesno(
                "確認",
                "元のファイル・フォルダを削除します。この操作は元に戻せません。\n本当に実行しますか？",
                icon="warning",
            )
            if not ok:
                return

        for c in result_box.winfo_children():
            c.destroy()

        def target(cb):
            applied = app.last_results.get("tags_applied") or {}
            return backend_organize.run(excel_path, app.cfg["output_folder"], cb,
                                         do_delete_originals=delete_var.get(),
                                         source_root=app.cfg.get("source_folder"),
                                         keep_folders=applied.get("tag_failed_folders"),
                                         folder_actions=folder_actions,
                                         file_actions=file_actions)

        def on_done(result):
            app.completed_steps.add("organize")
            app.last_results["organize"] = result
            render_result(result)

        app.run_task("ファイルをコピー中...", target, on_done)

    FOLDER_CONFLICT_OPTIONS = [("overwrite", "上書き"), ("merge", "統合"), ("new", "新しいフォルダを作成")]
    FILE_CONFLICT_OPTIONS = [("overwrite", "上書き"), ("rename", "ファイル名に「_X」を付与")]

    def make_choice_pills(parent, options, default, on_change=None):
        """A row of mutually-exclusive pill buttons (like STEP06's artwork
        choice). Returns (row_frame, {"get": fn, "set": fn})."""
        row = tk.Frame(parent, bg=theme["bg_elevated_2"])
        box = {"value": default}
        pills = {}

        def refresh():
            for key, (pill, lbl) in pills.items():
                selected = key == box["value"]
                bg = theme["accent"] if selected else theme["bg"]
                fg = theme["accent_text"] if selected else theme["text_muted"]
                pill.configure(bg=bg)
                lbl.configure(bg=bg, fg=fg)

        def set_value(key):
            box["value"] = key
            refresh()
            if on_change:
                on_change(key)

        for key, label in options:
            pill = tk.Frame(row, bg=theme["bg"], cursor="hand2")
            pill.pack(side="left", padx=(0, 6))
            lbl = tk.Label(pill, text=label, bg=theme["bg"], fg=theme["text_muted"],
                            font=(FONT_BODY, 9), padx=8, pady=4)
            lbl.pack()
            pill.bind("<Button-1>", lambda e, k=key: set_value(k))
            lbl.bind("<Button-1>", lambda e, k=key: set_value(k))
            pills[key] = (pill, lbl)

        refresh()
        return row, {"get": lambda: box["value"], "set": set_value}

    def show_conflict_panel(conflicts):
        for c in conflict_box.winfo_children():
            c.destroy()
        for c in result_box.winfo_children():
            c.destroy()

        folder_states = {}
        file_states = {}

        card = w.card(conflict_box, theme)
        card.pack(fill="both", expand=True)
        inner = tk.Frame(card, bg=theme["bg_elevated"])
        inner.pack(fill="both", expand=True, padx=22, pady=18)

        tk.Label(inner, text="コピー先の確認", bg=theme["bg_elevated"], fg=theme["text"],
                  font=(FONT_DISPLAY, 13, "bold"), anchor="w").pack(fill="x")
        tk.Label(inner,
                  text="以下のフォルダ・ファイルは既にコピー先に存在します。操作を選んでください"
                       "(未選択のままなら、フォルダは「統合」、ファイルは「_Xを付与」として扱われます)。",
                  bg=theme["bg_elevated"], fg=theme["text_muted"], font=(FONT_BODY, 10),
                  anchor="w", justify="left", wraplength=880).pack(fill="x", pady=(4, 10))

        overwrite_all_var = tk.BooleanVar(value=False)

        def toggle_overwrite_all():
            # unchecking goes back to the defaults (folders: merge, files: add "_X")
            on = overwrite_all_var.get()
            for st in folder_states.values():
                st["set"]("overwrite" if on else "merge")
            for st in file_states.values():
                st["set"]("overwrite" if on else "rename")

        tk.Checkbutton(inner, text="すべて上書きにする", variable=overwrite_all_var,
                        bg=theme["bg_elevated"], fg=theme["text_muted"], selectcolor=theme["bg_elevated_2"],
                        activebackground=theme["bg_elevated"], font=(FONT_BODY, 11),
                        command=toggle_overwrite_all).pack(anchor="w", pady=(0, 12))

        scroll_wrap = tk.Frame(inner, bg=theme["bg_elevated"])
        scroll_wrap.pack(fill="both", expand=True)
        canvas = tk.Canvas(scroll_wrap, bg=theme["bg_elevated"], highlightthickness=0, height=360)
        scrollbar = tk.Scrollbar(scroll_wrap, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=theme["bg_elevated"])
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        w.enable_mousewheel(canvas)

        for f in conflicts["folders"]:
            dest = f["dest_folder"]
            row_card = tk.Frame(scroll_frame, bg=theme["bg_elevated_2"])
            row_card.pack(fill="x", padx=4, pady=6)

            head = tk.Frame(row_card, bg=theme["bg_elevated_2"])
            head.pack(fill="x", padx=12, pady=(10, 4))
            tk.Label(head, text=f"{f['albumartist']} / {f['album']}", bg=theme["bg_elevated_2"],
                      fg=theme["text"], font=(FONT_BODY, 11, "bold"), anchor="w").pack(fill="x")
            tk.Label(head, text=dest, bg=theme["bg_elevated_2"], fg=theme["text_faint"],
                      font=(FONT_MONO, 9), anchor="w").pack(fill="x")

            file_conflicts = conflicts["files"].get(dest, [])
            file_panel_holder = tk.Frame(row_card, bg=theme["bg_elevated_2"])

            def update_file_panel_visibility(dest=dest, holder=file_panel_holder):
                if folder_states[dest]["get"]() == "merge" and file_conflicts:
                    holder.pack(fill="x", padx=24, pady=(0, 10))
                else:
                    holder.pack_forget()

            def on_folder_change(key, dest=dest):
                if key != "overwrite":
                    overwrite_all_var.set(False)
                update_file_panel_visibility(dest=dest)

            pill_row, fstate = make_choice_pills(row_card, FOLDER_CONFLICT_OPTIONS, "merge",
                                                 on_change=on_folder_change)
            pill_row.pack(fill="x", padx=12, pady=(0, 8))
            folder_states[dest] = fstate

            if file_conflicts:
                tk.Label(file_panel_holder, text=f"重複するファイル ({len(file_conflicts)}件)",
                          bg=theme["bg_elevated_2"], fg=theme["text_muted"], font=(FONT_BODY, 9, "bold"),
                          anchor="w").pack(fill="x", pady=(4, 4))
                for fc in file_conflicts:
                    frow = tk.Frame(file_panel_holder, bg=theme["bg_elevated_2"])
                    frow.pack(fill="x", pady=2)
                    tk.Label(frow, text=fc["rel_path"], bg=theme["bg_elevated_2"], fg=theme["text_faint"],
                              font=(FONT_MONO, 9), anchor="w").pack(side="left", fill="x", expand=True)
                    fpill_row, ffstate = make_choice_pills(
                        frow, FILE_CONFLICT_OPTIONS, "rename",
                        on_change=lambda key: key != "overwrite" and overwrite_all_var.set(False))
                    fpill_row.pack(side="right")
                    file_states[(dest, fc["rel_path"])] = ffstate

            update_file_panel_visibility()

        btn_row = tk.Frame(inner, bg=theme["bg_elevated"])
        btn_row.pack(fill="x", pady=(14, 0))

        def confirm():
            folder_actions = {dest: st["get"]() for dest, st in folder_states.items()}
            file_actions = {key: st["get"]() for key, st in file_states.items()}
            for c in conflict_box.winfo_children():
                c.destroy()
            do_run(folder_actions=folder_actions, file_actions=file_actions)

        def cancel():
            for c in conflict_box.winfo_children():
                c.destroy()

        w.button(btn_row, "反映して実行", confirm, theme, kind="primary").pack(side="left")
        w.button(btn_row, "キャンセル", cancel, theme, kind="secondary").pack(side="left", padx=(8, 0))

    def start_run():
        for c in conflict_box.winfo_children():
            c.destroy()

        run_btn.configure(state="disabled")

        def on_checked(conflicts, error):
            run_btn.configure(state="normal")
            if error:
                messagebox.showerror("エラー", f"コピー先の確認に失敗しました。\n{error}")
                return
            if conflicts["folders"]:
                show_conflict_panel(conflicts)
            else:
                do_run()

        w.run_async(body, lambda: backend_organize.find_conflicts(
            excel_path, app.cfg["output_folder"]), on_checked)

    run_btn = w.button(body, "実行", start_run, theme, kind="primary")
    # directly under the options card, not below the (expanding) conflict/result areas
    run_btn.pack(anchor="w", pady=(0, 20), before=conflict_box)

    existing = app.last_results.get("organize")
    if existing:
        render_result(existing)


def _stat(parent, theme, value, label):
    box = tk.Frame(parent, bg=theme["bg_elevated"])
    box.pack(side="left", padx=(0, 28))
    tk.Label(box, text=value, bg=theme["bg_elevated"], fg=theme["text"],
              font=(FONT_DISPLAY, 20, "bold")).pack(anchor="w")
    tk.Label(box, text=label, bg=theme["bg_elevated"], fg=theme["text_faint"],
              font=(FONT_BODY, 10)).pack(anchor="w")
