# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog

from tksheet import Sheet

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from theme import FONT_DISPLAY, FONT_BODY
from config import save_config
from backend import tag_table as backend_tag_table

HEADERS = ["フォルダ", "ファイル名"] + [name for _, name in backend_tag_table.FIELDS] + ["path"]
COL_FOLDER = 0
COL_FILENAME = 1
COL_FIRST_FIELD = 2
# the file path rides along in a hidden last column so it stays attached to its
# row when rows are sorted
COL_PATH = len(HEADERS) - 1
COLUMN_WIDTHS = [160, 220, 220, 160, 200, 160, 60, 60, 70, 200, 120, 160, 60, 1]

PATTERN_HELP = ("使えるタグ名: %title% %artist% %album% %albumartist% %track% %disc% "
                "%year% %genre% %composer% %comment% %bpm%")


def build(parent, app):
    theme = app.theme
    parent.configure(bg=theme["bg"])
    state = {"folder": None, "original": {}, "dirty": 0}

    outer = tk.Frame(parent, bg=theme["bg"])
    outer.pack(fill="both", expand=True, padx=28, pady=20)

    header = tk.Frame(outer, bg=theme["bg"])
    header.pack(fill="x", pady=(0, 10))
    back = tk.Frame(header, bg=theme["bg_elevated_2"], width=32, height=32,
                     highlightbackground=theme["border"], highlightthickness=1)
    back.pack(side="left")
    back.pack_propagate(False)
    back_lbl = tk.Label(back, text="←", bg=theme["bg_elevated_2"], fg=theme["text_muted"], font=(FONT_BODY, 14))
    back_lbl.pack(expand=True)
    for widget in (back, back_lbl):
        widget.bind("<Button-1>", lambda e: go_back())
        widget.configure(cursor="hand2")

    title_box = tk.Frame(header, bg=theme["bg"])
    title_box.pack(side="left", padx=(12, 0))
    title_row = tk.Frame(title_box, bg=theme["bg"])
    title_row.pack(anchor="w")
    tk.Label(title_row, text="タグ一覧編集", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "フォルダ内のMP3のタグを表形式で直接編集します。\n"
                 "セルをダブルクリックまたは入力で編集、⌘C/⌘Vでコピー・貼り付け、⌘Zで元に戻す、⌘Fで検索・置換。\n"
                 "変更したセルは色が付き、「保存」を押すまでファイルは変更されません。\n"
                 "ファイル名の列を書き換えると、保存時にファイル名も変更されます。"
                 ).pack(side="left", padx=(6, 0))
    folder_var = tk.StringVar(value="フォルダが開かれていません")
    tk.Label(title_box, textvariable=folder_var, bg=theme["bg"], fg=theme["text_muted"],
              font=(FONT_BODY, 10)).pack(anchor="w")

    save_btn = w.button(header, "保存", lambda: save(), theme, kind="primary")
    save_btn.pack(side="right")
    save_btn.configure(state="disabled")
    w.button(header, "フォルダを開く", lambda: open_folder(), theme, kind="secondary").pack(side="right", padx=(0, 8))

    toolbar = tk.Frame(outer, bg=theme["bg"])
    toolbar.pack(fill="x", pady=(0, 8))
    w.button(toolbar, "選択範囲に一括入力", lambda: fill_selection(), theme, kind="secondary").pack(side="left")
    w.button(toolbar, "ファイル名→タグ", lambda: filename_to_tags(), theme, kind="secondary").pack(side="left", padx=(8, 0))
    w.button(toolbar, "タグ→ファイル名", lambda: tags_to_filename(), theme, kind="secondary").pack(side="left", padx=(8, 0))
    w.button(toolbar, "検索・置換", lambda: sheet.open_find(focus=True), theme, kind="secondary").pack(side="left", padx=(8, 0))
    count_var = tk.StringVar(value="")
    tk.Label(toolbar, textvariable=count_var, bg=theme["bg"], fg=theme["text_muted"],
              font=(FONT_BODY, 11)).pack(side="right")

    sheet = Sheet(outer, headers=HEADERS, data=[], font=(FONT_BODY, 12, "normal"),
                  header_font=(FONT_BODY, 11, "bold"), show_row_index=True)
    sheet.change_theme("dark" if app.cfg.get("theme", "dark") == "dark" else "light blue")
    sheet.pack(fill="both", expand=True)
    sheet.enable_bindings("single_select", "drag_select", "ctrl_select", "select_all", "column_select",
                          "row_select", "column_width_resize", "double_click_column_resize", "arrowkeys",
                          "right_click_popup_menu", "rc_select", "copy", "cut", "paste", "delete",
                          "undo", "edit_cell", "find", "replace", "sort_rows")
    sheet.set_column_widths(COLUMN_WIDTHS)
    sheet.readonly_columns(columns=[COL_FOLDER, COL_PATH])
    sheet.hide_columns(columns=COL_PATH, data_indexes=True)
    sheet.bind("<<SheetModified>>", lambda e: refresh_changes())
    changed_bg = theme["yellow"] if app.cfg.get("theme", "dark") == "dark" else "#FFF1B8"
    changed_fg = theme["accent_text"] if app.cfg.get("theme", "dark") == "dark" else theme["text"]

    # ---------- helpers ----------
    def changed_cells(data):
        cells = []
        for r, row in enumerate(data):
            orig = state["original"].get(row[COL_PATH])
            if orig is None:
                continue
            for c in range(COL_FILENAME, COL_PATH):
                if (row[c] or "") != orig[c]:
                    cells.append((r, c))
        return cells

    def refresh_changes():
        data = sheet.get_sheet_data()
        cells = changed_cells(data)
        sheet.dehighlight_all(redraw=False)
        sheet.highlight_cells(cells=cells, bg=changed_bg, fg=changed_fg, redraw=True)
        state["dirty"] = len({r for r, _ in cells})
        count_var.set(f"{len(data)}曲" + (f" ・ 未保存の変更 {state['dirty']}曲" if state["dirty"] else ""))
        save_btn.configure(state="normal" if state["dirty"] else "disabled")

    def confirm_discard():
        if not state["dirty"]:
            return True
        return messagebox.askyesno("確認", f"未保存の変更が{state['dirty']}曲あります。\n破棄してよろしいですか？",
                                    icon="warning")

    def go_back():
        if confirm_discard():
            app.show_dashboard()

    def target_rows():
        rows = sorted({r for r, _ in sheet.get_selected_cells(get_rows=True, get_columns=True)})
        return rows or list(range(sheet.get_total_rows()))

    def apply_updates(updates):
        # writes {(row, col): value} as one block so a single ⌘Z undoes the whole operation
        data = sheet.get_sheet_data()
        updates = {k: v for k, v in updates.items() if (data[k[0]][k[1]] or "") != v}
        if not updates:
            messagebox.showinfo("変更なし", "変更される項目はありませんでした。")
            return
        rows = [r for r, _ in updates]
        cols = [c for _, c in updates]
        r0, c0 = min(rows), min(cols)
        block = [[updates.get((r, c), data[r][c]) for c in range(c0, max(cols) + 1)]
                 for r in range(r0, max(rows) + 1)]
        sheet.set_data((r0, c0), data=block, undo=True, emit_event=True)
        refresh_changes()

    def ask_pattern(title, cfg_key, default):
        value = simpledialog.askstring(title, f"パターンを入力してください。\n{PATTERN_HELP}",
                                        initialvalue=app.cfg.get(cfg_key) or default, parent=parent)
        if value:
            app.cfg[cfg_key] = value
            save_config(app.cfg)
        return value

    def row_tags(row):
        return {key: row[COL_FIRST_FIELD + i] or "" for i, key in enumerate(backend_tag_table.FIELD_KEYS)}

    # ---------- actions ----------
    def open_folder():
        if app.task_running:
            messagebox.showinfo("実行中", "他の処理が実行中です。完了をお待ちください。")
            return
        if not confirm_discard():
            return
        folder = filedialog.askdirectory(
            title="編集するフォルダを選択",
            initialdir=app.cfg.get("tag_table_folder") or app.cfg.get("source_folder") or os.path.expanduser("~"))
        if not folder:
            return
        app.cfg["tag_table_folder"] = folder
        save_config(app.cfg)

        def target(cb):
            return backend_tag_table.scan(folder, cb)

        def on_done(rows):
            data = [[r["folder"], r["filename"]] + [r[k] for k in backend_tag_table.FIELD_KEYS] + [r["path"]]
                    for r in rows]
            state["folder"] = folder
            state["original"] = {row[COL_PATH]: list(row) for row in data}
            folder_var.set(folder)
            sheet.set_sheet_data(data, reset_col_positions=False, reset_highlights=True)
            sheet.hide_columns(columns=COL_PATH, data_indexes=True)
            refresh_changes()

        app.run_task("タグ読み込み中...", target, on_done)

    def fill_selection():
        cells = [(r, c) for r, c in sheet.get_selected_cells(get_rows=True, get_columns=True)
                 if COL_FILENAME <= c < COL_PATH]
        if not cells:
            messagebox.showinfo("未選択", "入力先のセルを選択してください（フォルダ列は編集できません）。")
            return
        value = simpledialog.askstring("一括入力", f"選択中の{len(cells)}セルに入力する値:", parent=parent)
        if value is None:
            return
        apply_updates({cell: value for cell in cells})

    def filename_to_tags():
        rows = target_rows()
        if not rows:
            return
        pattern = ask_pattern("ファイル名→タグ", "tag_table_parse_pattern", "%track% %title%")
        if not pattern:
            return
        data = sheet.get_sheet_data()
        updates = {}
        unmatched = 0
        try:
            for r in rows:
                parsed = backend_tag_table.parse_filename(pattern, data[r][COL_FILENAME] or "")
                if parsed is None:
                    unmatched += 1
                    continue
                for key, value in parsed.items():
                    updates[(r, COL_FIRST_FIELD + backend_tag_table.FIELD_KEYS.index(key))] = value
        except ValueError as e:
            messagebox.showerror("パターンエラー", str(e))
            return
        if unmatched:
            messagebox.showwarning("一部不一致", f"{unmatched}曲はファイル名がパターンに合わないため変更しません。")
        apply_updates(updates)

    def tags_to_filename():
        rows = target_rows()
        if not rows:
            return
        pattern = ask_pattern("タグ→ファイル名", "tag_table_build_pattern", "%track% %title%")
        if not pattern:
            return
        data = sheet.get_sheet_data()
        updates = {}
        try:
            for r in rows:
                name = backend_tag_table.build_filename(pattern, row_tags(data[r]))
                if name:
                    updates[(r, COL_FILENAME)] = name
        except ValueError as e:
            messagebox.showerror("パターンエラー", str(e))
            return
        apply_updates(updates)

    def save():
        if app.task_running:
            messagebox.showinfo("実行中", "他の処理が実行中です。完了をお待ちください。")
            return
        data = sheet.get_sheet_data()
        changes = []
        bad_names = []
        for row in data:
            path = row[COL_PATH]
            orig = state["original"][path]
            tags = {key: row[COL_FIRST_FIELD + i] or "" for i, key in enumerate(backend_tag_table.FIELD_KEYS)
                    if (row[COL_FIRST_FIELD + i] or "") != orig[COL_FIRST_FIELD + i]}
            new_name = None
            if (row[COL_FILENAME] or "") != orig[COL_FILENAME]:
                new_name = row[COL_FILENAME] or ""
                err = backend_tag_table.invalid_filename(new_name)
                if err:
                    bad_names.append(f"{orig[COL_FILENAME]} → {new_name}: {err}")
            if tags or new_name is not None:
                changes.append({"path": path, "tags": tags, "filename": new_name})
        if bad_names:
            messagebox.showerror("ファイル名エラー", "次のファイル名は使えません。修正してから保存してください。\n\n"
                                  + "\n".join(bad_names[:10]))
            return
        if not changes:
            return
        if not messagebox.askyesno("保存の確認", f"{len(changes)}曲のファイルを書き換えます。よろしいですか？"):
            return

        def target(cb):
            return backend_tag_table.save(changes, app.cfg.get("id3_version", "v24"),
                                          app.cfg.get("strip_id3v1", False), cb)

        # snapshot of what is being written, so edits made while saving stay marked as unsaved
        saved_rows = {row[COL_PATH]: [v or "" for v in row] for row in data}

        def on_done(result):
            failed = {path for path, _ in result["errors"]}
            renames = result["renames"]
            for ch in changes:
                path = ch["path"]
                if path in failed:
                    continue
                new_path = renames.get(path, path)
                state["original"].pop(path)
                state["original"][new_path] = saved_rows[path][:COL_PATH] + [new_path]
            if renames:
                for r, row in enumerate(sheet.get_sheet_data()):
                    if row[COL_PATH] in renames:
                        sheet.set_cell_data(r, COL_PATH, renames[row[COL_PATH]], redraw=False)
            refresh_changes()
            if result["errors"]:
                lines = [f"{os.path.basename(p)}: {msg}" for p, msg in result["errors"][:10]]
                messagebox.showwarning("一部失敗", f"{len(result['errors'])}曲で失敗しました。\n\n" + "\n".join(lines))
            else:
                messagebox.showinfo("保存完了", f"{result['saved']}曲を保存しました"
                                    + (f"（名前変更 {result['renamed']}件）" if result["renamed"] else "") + "。")

        app.run_task("タグ保存中...", target, on_done)

    refresh_changes()
