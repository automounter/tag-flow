# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from platform_util import open_path
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import integrity_check
from backend import organize as backend_organize
from backend import color_legend


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
    tk.Label(title_row, text="ID3 タグ情報確認・適用", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "①STEP02直後のデータ(ベースライン)と現在のExcelを比較し、変更点と警告をレポートにします。\n"
                 "③で開くExcel(作業用のExcel)の内容が、⑤でMP3のID3タグに書き込まれます。\n"
                 "STEP06で反映したアートワークも⑤で書き込まれます。\n"
                 "④にチェックを入れるまで⑤は実行できません。"
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text="STEP 07 ・ 変更点を確認し、ID3タグをMP3に書き込みます",
              bg=theme["bg"], fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    body = tk.Frame(outer, bg=theme["bg"])
    body.pack(fill="both", expand=True)

    excel_path = app.get_excel_path()
    baseline_path = getattr(app, "baseline_excel_path", None)
    if not baseline_path and excel_path:
        # after a restart the path is gone, but STEP02 saved the file next to the Excel
        base, ext = os.path.splitext(excel_path)
        baseline_path = f"{base}_baseline{ext}"

    if not excel_path or not baseline_path or not os.path.exists(baseline_path):
        tk.Label(body, text="先に STEP 02〜05 を実行してください(元データとの比較にはSTEP02実行直後の\nバックアップが必要です)。",
                  bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 12), justify="left").pack(pady=40)
        return

    bg = theme["bg_elevated"]
    card = w.card(body, theme)
    card.pack(fill="x")
    inner = tk.Frame(card, bg=bg)
    inner.pack(fill="x", padx=24, pady=20)

    stat_row = tk.Frame(inner, bg=bg)
    stat_row.pack(fill="x")
    changes_var = tk.StringVar(value="—")
    warnings_var = tk.StringVar(value="—")
    _stat(stat_row, theme, changes_var, "変更点")
    _stat(stat_row, theme, warnings_var, "警告")
    report_note_var = tk.StringVar(value="①を実行すると件数が表示されます。")
    tk.Label(stat_row, textvariable=report_note_var, bg=bg, fg=theme["text_faint"],
              font=(FONT_BODY, 10)).pack(side="left", anchor="s", pady=(0, 6))

    tk.Frame(inner, bg=theme["border"], height=1).pack(fill="x", pady=16)

    def row(label_widget_fn):
        r = tk.Frame(inner, bg=bg)
        r.pack(fill="x", pady=(0, 10))
        return label_widget_fn(r)

    check_btn = row(lambda r: w.button(r, "① 整合性を確認", lambda: do_check(), theme, kind="secondary"))
    report_btn = row(lambda r: w.button(r, "② 整合性確認レポートをExcelで表示", lambda: open_report(), theme,
                                        kind="secondary"))
    def open_working():
        def target(cb):
            cb(0, 1, "色の説明を書き込み中...")
            return color_legend.ensure_working_legend(excel_path)

        app.run_task("Excelを開く準備中...", target, lambda ok: open_path(excel_path))

    excel_btn = row(lambda r: w.button(r, "③ 適用案の確認内容をExcelで表示", open_working, theme,
                                       kind="secondary"))
    for b in (check_btn, report_btn, excel_btn):
        b.pack(side="left")
    tk.Label(excel_btn.master, text="このExcel(IDTAGシート)の曲名・アーティスト・アルバム・ジャンル・アルバムアーティストが、"
                                    "⑤でそのままMP3に書き込まれます。修正はこのExcelで行い、保存してから⑤を実行してください。"
                                    "セルの色の意味は「色の説明」のシートにあります。",
              bg=bg, fg=theme["text_faint"], font=(FONT_BODY, 10), wraplength=560, justify="left",
              anchor="w").pack(side="left", padx=(12, 0))

    confirmed_var = tk.BooleanVar(value=False)
    tk.Checkbutton(inner, text="④ 適用案の内容確認および修正完了", variable=confirmed_var,
                    bg=bg, fg=theme["text"], selectcolor=theme["bg_elevated_2"],
                    activebackground=bg, font=(FONT_BODY, 11),
                    command=lambda: update_apply_btn()).pack(anchor="w", pady=(4, 10))

    apply_row = tk.Frame(inner, bg=bg)
    apply_row.pack(fill="x")
    apply_btn = w.button(apply_row, "⑤ ID3タグを更新", lambda: do_apply(), theme, kind="primary")
    apply_btn.pack(side="left")
    apply_btn.configure(state="disabled")
    w.button(apply_row, "元に戻す(最初に読み込んだときのタグ)", lambda: do_restore(), theme,
             kind="secondary").pack(side="left", padx=(12, 0))
    tk.Label(inner,
              text="※ MP3保存元フォルダの元のファイルのタグを書き換えます。STEP 02 で最初に読み込んだときのタグを"
                   "控えとして保存してあり、「元に戻す」でその状態に戻せます(空欄だった項目・アートワーク・ID3のバージョンも含む)。\n"
                   "STEP 08 で元データを削除したファイルは戻せません。",
              bg=bg, fg=theme["yellow"], font=(FONT_BODY, 10), wraplength=900, justify="left",
              anchor="w").pack(fill="x", pady=(8, 0))

    result_var = tk.StringVar(value="")
    result_lbl = tk.Label(inner, textvariable=result_var, bg=bg, fg=theme["green"], font=(FONT_BODY, 11, "bold"),
                           wraplength=900, justify="left", anchor="w")
    result_lbl.pack(fill="x", pady=(12, 0))

    def update_apply_btn():
        apply_btn.configure(state="normal" if confirmed_var.get() else "disabled")

    def show_report_result(result):
        changes_var.set(str(result["changes"]))
        warnings_var.set(str(result["warnings"]))
        report_note_var.set(os.path.basename(result["report_path"]))
        report_btn.configure(state="normal")

    def open_report():
        result = app.last_results.get("integrity")
        if not result or not os.path.exists(result["report_path"]):
            messagebox.showinfo("レポートなし", "先に「① 整合性を確認」を実行してください。")
            return
        open_path(result["report_path"])

    def do_check():
        output_dir = app.cfg.get("excel_folder")

        def target(cb):
            return integrity_check.run(baseline_path, excel_path, output_dir, cb)

        def on_done(result):
            app.last_results["integrity"] = result
            app.integrity_report_path = result["report_path"]
            show_report_result(result)

        app.run_task("整合性を確認中...", target, on_done)

    def do_apply():
        if not messagebox.askyesno(
                "ID3タグを更新",
                "作業用Excelの内容を、MP3保存元フォルダの元のファイルのタグに書き込みます。\n"
                "Excelを編集した場合は、保存してから実行してください。実行しますか?"):
            return

        def target(cb):
            return backend_organize.write_tags(excel_path, cb,
                                               id3_version=app.cfg.get("id3_version", "v24"),
                                               strip_id3v1=app.cfg.get("strip_id3v1", False),
                                               artwork_overrides=app.artwork_overrides)

        def on_done(result):
            app.last_results["tags_applied"] = result
            errors = result["tag_errors"]
            if errors:
                result_lbl.configure(fg=theme["red"])
                result_var.set(f"{result['tags_written']}曲を更新しました(変更なし {result.get('tags_unchanged', 0)}曲)。"
                               f"エラー {len(errors)}件(詳細はログ)。もう一度⑤を押すと、エラーの曲だけ書き込み直します"
                               "(済んだ曲は飛ばします)。エラーのあったフォルダは、STEP 08 で元データを削除しません。")
                app.log.error(f"ID3タグ更新エラー: {errors}")
            else:
                result_lbl.configure(fg=theme["green"])
                result_var.set(f"{result['tags_written']}曲のID3タグを更新しました(変更なし {result.get('tags_unchanged', 0)}曲)。"
                               "STEP 08 (適用ファイル操作) に進めます。")
            app.completed_steps.add("integrity_check")

        app.run_task("ID3タグを更新中...", target, on_done)

    def do_restore():
        if not messagebox.askyesno(
                "元に戻す",
                "STEP 02 で最初に読み込んだときに保存した控えから、元のファイルのタグを元に戻します。\n"
                "(空欄だった項目・アートワーク・ID3のバージョンも元に戻ります)\n\n"
                "STEP 08 で元データを削除したファイルは対象外です。実行しますか?", icon="warning"):
            return

        def target(cb):
            return backend_organize.restore_tags(excel_path, cb)

        def on_done(result):
            app.completed_steps.discard("integrity_check")
            app.last_results.pop("tags_applied", None)
            problems = []
            if result["no_snapshot"]:
                problems.append(f"控えが無い {len(result['no_snapshot'])}件")
            if result["not_found"]:
                problems.append(f"ファイルが見つからない {len(result['not_found'])}件")
            if result["errors"]:
                problems.append(f"エラー {len(result['errors'])}件")
                app.log.error(f"タグを元に戻す処理のエラー: {result['errors']}")
            result_lbl.configure(fg=theme["red"] if problems else theme["green"])
            result_var.set(f"{result['restored']}/{result['total']}曲のタグを元に戻しました。"
                           + ("(" + "・".join(problems) + ")" if problems else ""))

        app.run_task("タグを元に戻しています...", target, on_done)

    existing = app.last_results.get("integrity")
    if existing:
        show_report_result(existing)
    else:
        report_btn.configure(state="disabled")
    applied = app.last_results.get("tags_applied")
    if applied and "integrity_check" in app.completed_steps:
        result_var.set(f"適用済み:{applied['tags_written']}/{applied['tags_total']}曲。")


def _stat(parent, theme, value_var, label):
    box = tk.Frame(parent, bg=theme["bg_elevated"])
    box.pack(side="left", padx=(0, 32))
    tk.Label(box, textvariable=value_var, bg=theme["bg_elevated"], fg=theme["text"],
              font=(FONT_DISPLAY, 22, "bold")).pack(anchor="w")
    tk.Label(box, text=label, bg=theme["bg_elevated"], fg=theme["text_faint"],
              font=(FONT_BODY, 10)).pack(anchor="w")
