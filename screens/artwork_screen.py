# -*- coding: utf-8 -*-
import io
import os
import sys
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import widgets as w
from theme import FONT_DISPLAY, FONT_BODY, FONT_MONO
from backend import artwork_check
from backend import name_lookup as nl

from PIL import Image, ImageTk

THUMB_SIZE = 64
STATUS_LABEL = {"missing": "画像が見つかりません", "mismatch": "公式画像と異なる可能性があります"}
SOURCE_LABEL = {"local": "フォルダから選択済み", "url": "URLから指定済み", "official": "推奨画像を採用済み",
                "delete": "アートワーク削除を選択済み", None: "変更しない"}


_run_async = w.run_async


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
    tk.Label(title_row, text="アルバム写真の整合性確認", bg=theme["bg"], fg=theme["text"],
              font=(FONT_DISPLAY, 16, "bold")).pack(side="left")
    w.help_badge(title_row, theme,
                 "埋め込み画像と、選んだサービスの画像(Apple Music→Spotify→Discogsの順で最初に見つかったもの)を\n"
                 "8×8の縮小比較(perceptual hash)で照合します。ほかのサービスの画像も候補として切り替えられます。\n"
                 "解像度違いなどの小さな差は無視し、明らかに異なる場合のみ警告します。\n"
                 "「反映」を押すまでMP3ファイル自体は変更されません。"
                 ).pack(side="left", padx=(6, 0))
    tk.Label(title_box, text="STEP 06 (任意) ・ Apple Music・Spotify・Discogs の画像と比較します", bg=theme["bg"],
              fg=theme["text_muted"], font=(FONT_BODY, 10)).pack(anchor="w")

    excel_path = app.get_excel_path()
    if not excel_path:
        tk.Label(outer, text="先に STEP 02 (MP3タグ抽出) を実行してください。", bg=theme["bg"],
                  fg=theme["text_faint"], font=(FONT_BODY, 12)).pack(pady=40)
        return

    status_var = tk.StringVar(value="「スキャン開始」を押すとアルバムごとに公式画像と比較します(時間がかかります)。"
                                     " 画像が異なる場合は各アルバムごとに①推奨画像を採用②フォルダから選択③URLから指定④アートワーク削除、"
                                     "のいずれかを選んでください(未選択のままなら変更しません)。")
    status_lbl = tk.Label(outer, textvariable=status_var, bg=theme["bg"], fg=theme["text_muted"],
                           font=(FONT_BODY, 11), anchor="w", justify="left", wraplength=1080)
    status_lbl.pack(fill="x", pady=(0, 10))

    keys = nl.load_api_keys()
    available = nl.available_sources(keys)
    saved = app.cfg.get("artwork_sources") or list(artwork_check.ARTWORK_SOURCES)
    service_row = tk.Frame(outer, bg=theme["bg"])
    service_row.pack(fill="x", pady=(0, 8))
    tk.Label(service_row, text="照会するサービス:", bg=theme["bg"], fg=theme["text_muted"],
              font=(FONT_BODY, 11)).pack(side="left")
    source_vars = {}
    for src in artwork_check.ARTWORK_SOURCES:
        ok = src in available
        source_vars[src] = tk.BooleanVar(value=ok and src in saved)
        note = {"apple": "1000px・約3秒に1回", "spotify": "最大640px", "discogs": "利用者の画像・約1秒に1回"}[src]
        tk.Checkbutton(service_row, text=f"{artwork_check.SOURCE_NAMES[src]}({note if ok else 'キー未設定'})",
                        variable=source_vars[src], bg=theme["bg"], fg=theme["text"],
                        selectcolor=theme["bg_elevated_2"], activebackground=theme["bg"], font=(FONT_BODY, 11),
                        state="normal" if ok else "disabled").pack(side="left", padx=(8, 0))

    toolbar = tk.Frame(outer, bg=theme["bg"])
    toolbar.pack(fill="x", pady=(0, 8))

    selected_count_var = tk.StringVar(value="選択中: 0件")
    tk.Label(toolbar, textvariable=selected_count_var, bg=theme["bg"], fg=theme["text_faint"],
              font=(FONT_BODY, 10)).pack(side="right", padx=(0, 14))

    action_holder = tk.Frame(toolbar, bg=theme["bg"])
    action_holder.pack(side="right")

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

    state = {"items": getattr(app, "artwork_scan_results", None) or [], "choices": {}, "images": []}

    def to_thumb(image_bytes):
        try:
            img = Image.open(io.BytesIO(image_bytes))
            img.thumbnail((THUMB_SIZE, THUMB_SIZE), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            state["images"].append(photo)  # keep a reference alive
            return photo
        except Exception:
            return None

    def update_selected_count():
        n = sum(1 for c in state["choices"].values() if c.get("source") in ("local", "url", "official", "delete"))
        selected_count_var.set(f"選択中: {n}件")

    def render_list():
        for c in scroll_frame.winfo_children():
            c.destroy()
        state["choices"] = {}
        state["images"] = []

        items = state["items"]
        if not items:
            tk.Label(scroll_frame, text="対象がありません(未スキャン、または問題は見つかりませんでした)。",
                      bg=theme["bg_elevated"], fg=theme["text_faint"], font=(FONT_BODY, 11)).pack(pady=30)
            update_selected_count()
            return

        for idx, item in enumerate(items):
            state["choices"][idx] = {"source": None, "value": None}
            _build_row(scroll_frame, theme, idx, item, state, to_thumb, update_selected_count)

        update_selected_count()

    def do_scan():
        sources = [src for src, var in source_vars.items() if var.get()]
        if not sources:
            messagebox.showinfo("サービス未選択", "照会するサービスを1つ以上選んでください。")
            return
        app.cfg["artwork_sources"] = sources
        from config import save_config
        save_config(app.cfg)
        status_var.set("スキャン中...(ネットワーク経由のため時間がかかります)")

        def target(cb):
            return artwork_check.scan(excel_path, cb, sources)

        def on_done(results):
            state["items"] = results
            app.artwork_scan_results = results
            status_var.set(f"{len(results)}件のアルバムで画像の問題が見つかりました。各アルバムごとに更新方法を選んでください。"
                            if results else "問題のあるアルバムは見つかりませんでした。")
            render_list()

        app.run_task("アートワークを確認中...", target, on_done)

    def do_apply():
        choices = []
        for idx, c in state["choices"].items():
            if c.get("source") in ("local", "url", "official", "delete"):
                item = state["items"][idx]
                choices.append({"folder": item["folder"], "album": item["album"],
                                 "source": c["source"], "value": c.get("value")})
        if not choices:
            messagebox.showinfo("未選択", "反映する項目がありません。①推奨画像を採用、②フォルダから選択、"
                                          "③URLから指定、④アートワーク削除のいずれかを選んでください。")
            return

        def target(cb):
            overrides, errors = artwork_check.apply_choices(choices, cb)
            applied = [c for c in choices if c["folder"] in overrides]
            artwork_check.write_excel_results(excel_path, applied)
            return overrides, errors

        def on_done(result):
            overrides, errors = result
            existing = getattr(app, "artwork_overrides", {}) or {}
            existing.update(overrides)
            app.artwork_overrides = existing
            if overrides:
                app.completed_steps.add("artwork")
                prev = app.last_results.get("artwork", {}).get("applied", 0)
                app.last_results["artwork"] = {"applied": prev + len(overrides)}
            msg = f"{len(overrides)}件を反映しました(Excelにも記録しました)。\nSTEP 07 の「⑤ ID3タグを更新」で書き込まれます。"
            if errors:
                msg += "\n\nエラー:\n" + "\n".join(f"・{album}: {err}" for album, err in errors)
            messagebox.showinfo("反映しました", msg)
            app.show_artwork()

        app.run_task("画像を確認・保存中...", target, on_done)

    scan_btn = w.button(action_holder, "スキャン開始", do_scan, theme, kind="secondary")
    scan_btn.pack(side="left", padx=(0, 8))
    apply_btn = w.button(action_holder, "選択した項目を反映", do_apply, theme, kind="primary")
    apply_btn.pack(side="left")

    render_list()


def _build_row(scroll_frame, theme, idx, item, state, to_thumb, update_selected_count):
    row_card = tk.Frame(scroll_frame, bg=theme["bg_elevated_2"])
    row_card.pack(fill="x", padx=10, pady=6)

    top = tk.Frame(row_card, bg=theme["bg_elevated_2"])
    top.pack(fill="x", padx=12, pady=(10, 6))

    cur_box = tk.Frame(top, bg=theme["bg"], width=THUMB_SIZE, height=THUMB_SIZE)
    cur_box.pack(side="left")
    cur_box.pack_propagate(False)
    local_bytes = item.get("local_bytes")
    if local_bytes:
        thumb = to_thumb(local_bytes)
        if thumb:
            tk.Label(cur_box, image=thumb, bg=theme["bg"]).pack(expand=True)
    else:
        tk.Label(cur_box, text="なし", bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 9)).pack(expand=True)

    tk.Label(top, text="→", bg=theme["bg_elevated_2"], fg=theme["text_faint"],
              font=(FONT_BODY, 12)).pack(side="left", padx=6)

    ref_box = tk.Frame(top, bg=theme["bg"], width=THUMB_SIZE, height=THUMB_SIZE)
    ref_box.pack(side="left")
    ref_box.pack_propagate(False)
    ref_placeholder = tk.Label(ref_box, text="読込中" if item.get("artwork_url") else "未取得",
                                bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 9))
    ref_placeholder.pack(expand=True)

    ref = {"url": item.get("artwork_url")}  # the candidate shown (and adopted by "推奨画像を採用")

    def show_reference(url):
        ref["url"] = url
        for c in ref_box.winfo_children():
            c.destroy()
        tk.Label(ref_box, text="読込中", bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 9)).pack(expand=True)

        def on_reference_fetched(data, err):
            if err or not data or ref["url"] != url:
                return
            thumb = to_thumb(data)
            if thumb:
                for c in ref_box.winfo_children():
                    c.destroy()
                tk.Label(ref_box, image=thumb, bg=theme["bg"]).pack(expand=True)

        _run_async(ref_box, lambda: artwork_check._request_bytes(url), on_reference_fetched)

    if ref["url"]:
        show_reference(ref["url"])

    info_box = tk.Frame(top, bg=theme["bg_elevated_2"])
    info_box.pack(side="left", fill="x", expand=True, padx=(10, 0))
    tk.Label(info_box, text=item["album"], bg=theme["bg_elevated_2"], fg=theme["text"],
              font=(FONT_BODY, 11, "bold"), anchor="w").pack(fill="x")
    tk.Label(info_box, text=f"{item['albumartist']} ・ {STATUS_LABEL.get(item['status'], item['status'])}",
              bg=theme["bg_elevated_2"], fg=theme["text_faint"], font=(FONT_MONO, 9), anchor="w").pack(fill="x")
    # older scan results (before services were selectable) only have Apple's URL
    candidates = item.get("candidates") or ([{"source": "apple", "size": 1000, "url": ref["url"]}] if ref["url"] else [])
    cand_row = tk.Frame(info_box, bg=theme["bg_elevated_2"])
    cand_row.pack(fill="x", pady=(2, 0))
    cand_labels = []

    def refresh_candidates():
        for c, lbl in cand_labels:
            small = c["size"] < artwork_check.MAX_ARTWORK_SIZE
            chosen = c["url"] == ref["url"]
            lbl.configure(fg=theme["yellow"] if small else theme["text_muted"],
                          font=(FONT_BODY, 9, "bold" if chosen else "normal"))

    def pick_candidate(c):
        show_reference(c["url"])
        if state["choices"][idx]["source"] == "official":  # adopting follows the shown candidate
            choose_official()
        refresh_candidates()

    if candidates:
        tk.Label(cand_row, text="候補(クリックで切替):", bg=theme["bg_elevated_2"], fg=theme["text_faint"],
                  font=(FONT_BODY, 9)).pack(side="left")
    for c in candidates:
        size = f"{c['size']}px" if c["size"] else "大きさ不明"
        warn = "・1000px未満" if c["size"] < artwork_check.MAX_ARTWORK_SIZE else ""
        lbl = tk.Label(cand_row, text=f"{artwork_check.SOURCE_NAMES[c['source']]} {size}{warn}",
                        bg=theme["bg_elevated_2"], cursor="hand2", padx=6)
        lbl.pack(side="left")
        lbl.bind("<Button-1>", lambda e, c=c: pick_candidate(c))
        cand_labels.append((c, lbl))
    refresh_candidates()

    bottom = tk.Frame(row_card, bg=theme["bg_elevated_2"])
    bottom.pack(fill="x", padx=12, pady=(0, 10))

    preview_box = tk.Frame(bottom, bg=theme["bg"], width=THUMB_SIZE, height=THUMB_SIZE)
    preview_box.pack(side="left")
    preview_box.pack_propagate(False)
    tk.Label(preview_box, text="---", bg=theme["bg"], fg=theme["text_faint"], font=(FONT_BODY, 9)).pack(expand=True)

    status_var = tk.StringVar(value=SOURCE_LABEL[None])
    status_lbl = tk.Label(bottom, textvariable=status_var, bg=theme["bg_elevated_2"], fg=theme["text_faint"],
                           font=(FONT_BODY, 9))

    def set_preview(image_bytes=None, placeholder="---"):
        for c in preview_box.winfo_children():
            c.destroy()
        thumb = to_thumb(image_bytes) if image_bytes else None
        if thumb:
            tk.Label(preview_box, image=thumb, bg=theme["bg"]).pack(expand=True)
        else:
            tk.Label(preview_box, text=placeholder, bg=theme["bg"], fg=theme["text_faint"],
                      font=(FONT_BODY, 9)).pack(expand=True)

    pills = {}

    def refresh_pills():
        chosen = state["choices"][idx]["source"]
        for key, (pill, lbl) in pills.items():
            selected = key == chosen
            bg = theme["accent"] if selected else theme["bg"]
            fg = theme["accent_text"] if selected else theme["text_muted"]
            pill.configure(bg=bg)
            lbl.configure(bg=bg, fg=fg)

    def choose_delete():
        state["choices"][idx] = {"source": "delete", "value": None}
        status_var.set(SOURCE_LABEL["delete"])
        set_preview(None, placeholder="削除")
        refresh_pills()
        update_selected_count()

    def choose_official():
        ref_url = ref["url"]
        if not ref_url:
            return
        status_var.set("画像を取得中...")

        def on_fetched(result, err):
            if err:
                status_var.set(SOURCE_LABEL[state["choices"][idx]["source"]])
                messagebox.showerror("取得エラー", f"画像を取得できませんでした。\n{err}")
                return
            data, _ext = result
            state["choices"][idx] = {"source": "official", "value": ref_url}
            status_var.set(SOURCE_LABEL["official"])
            set_preview(data)
            refresh_pills()
            update_selected_count()

        _run_async(scroll_frame, lambda: artwork_check.fetch_url_image(ref_url), on_fetched)

    def choose_local():
        initial = item["folder"] if os.path.isdir(item["folder"]) else None
        path = filedialog.askopenfilename(
            title="画像ファイルを選択",
            initialdir=initial,
            filetypes=[("画像ファイル", "*.jpg *.jpeg *.png"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        try:
            data, _ext = artwork_check.load_local_image(path)
        except Exception as e:
            messagebox.showerror("読み込みエラー", f"画像を読み込めませんでした。\n{e}")
            return
        state["choices"][idx] = {"source": "local", "value": path}
        status_var.set(f"{SOURCE_LABEL['local']}: {os.path.basename(path)}")
        set_preview(data)
        refresh_pills()
        update_selected_count()

    def choose_url():
        url = simpledialog.askstring("URL入力", "JPEGまたはPNG画像のURLを入力してください:")
        if not url:
            return
        url = url.strip()
        status_var.set("画像を取得中...")

        def on_fetched(result, err):
            if err:
                status_var.set(SOURCE_LABEL[state["choices"][idx]["source"]])
                messagebox.showerror("取得エラー", f"画像を取得できませんでした。\n{err}")
                return
            data, _ext = result
            state["choices"][idx] = {"source": "url", "value": url}
            status_var.set(SOURCE_LABEL["url"])
            set_preview(data)
            refresh_pills()
            update_selected_count()

        _run_async(scroll_frame, lambda: artwork_check.fetch_url_image(url), on_fetched)

    def make_pill(parent, key, text, command):
        pill = tk.Frame(parent, bg=theme["bg"], cursor="hand2")
        pill.pack(side="left", padx=(0, 8))
        lbl = tk.Label(pill, text=text, bg=theme["bg"], fg=theme["text_muted"], font=(FONT_BODY, 10), padx=10, pady=5)
        lbl.pack()
        pill.bind("<Button-1>", lambda e: command())
        lbl.bind("<Button-1>", lambda e: command())
        pills[key] = (pill, lbl)

    pill_row = tk.Frame(bottom, bg=theme["bg_elevated_2"])
    pill_row.pack(side="left", padx=(10, 0))
    if ref["url"]:
        make_pill(pill_row, "official", "①推奨画像を採用", choose_official)
        make_pill(pill_row, "local", "②フォルダから選択", choose_local)
        make_pill(pill_row, "url", "③URLから指定", choose_url)
        make_pill(pill_row, "delete", "④アートワーク削除", choose_delete)
    else:
        make_pill(pill_row, "local", "①フォルダから選択", choose_local)
        make_pill(pill_row, "url", "②URLから指定", choose_url)
        make_pill(pill_row, "delete", "③アートワーク削除", choose_delete)

    status_lbl.pack(side="left", padx=(14, 0))
    refresh_pills()
