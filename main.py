# -*- coding: utf-8 -*-
import os
import sys
import threading
import queue
import time
import tkinter as tk
from tkinter import messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from theme import get_theme, FONT_DISPLAY, FONT_BODY, FONT_MONO
from config import load_config, save_config
import widgets as w
from logger import setup_logging

from screens import (dashboard, cleanup_screen, settings_screen, verify_screen, integrity_screen,
                      artwork_screen, organize_screen, genre_screen, tag_table_screen,
                      albumartist_screen, name_lookup_screen, site_fetch_screen)
from backend import extract as backend_extract
from backend import translate as backend_translate
from backend import genre as backend_genre
from backend import organize as backend_organize


class TaskCancelled(Exception):
    pass


class TagFlowApp:
    def __init__(self):
        self.log = setup_logging()
        self.log.info("Tag FloW 起動")

        self.cfg = load_config()
        self.theme = get_theme(self.cfg.get("theme", "dark"))

        self.root = tk.Tk()
        self.root.title("Tag FloW")
        # Linux では入力欄に白い枠（フォーカスの縁取り）が出るので消す
        self.root.option_add("*Entry.highlightThickness", 0)
        self.root.geometry("1180x820")
        self.root.minsize(980, 700)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.progress_queue = queue.Queue()
        self.completed_steps = set()
        self.last_results = {}
        self.baseline_excel_path = None
        self.integrity_report_path = None
        self.artwork_overrides = {}
        self.artwork_scan_results = None
        self.genre_overrides = {}
        self.albumartist_choices = {}
        self.task_running = False
        self.task_cancel_requested = False
        self.task_start_time = None
        self.task_label_var = tk.StringVar(value="")
        self.task_pct_var = tk.StringVar(value="")
        self.task_eta_var = tk.StringVar(value="")

        self.container = tk.Frame(self.root, bg=self.theme["bg"])
        self.container.pack(fill="both", expand=True)

        self.status_bar = tk.Frame(self.root, bg=self.theme["bg_elevated"], height=40)
        self.status_bar.pack(fill="x", side="bottom")
        self._build_status_bar()

        self.current_screen_name = "dashboard"
        self.show_dashboard()

        # Tk swallows callback errors to stderr, which nobody sees when started from run.command
        self.root.report_callback_exception = lambda *exc: self.log.error("画面の処理でエラー", exc_info=exc)
        self.root.after(100, self._poll_progress)

    def _on_close(self):
        if self.task_running:
            ok = messagebox.askyesno(
                "確認",
                "処理を実行中です。終了すると処理が中断され、\n"
                "ファイルが不完全な状態になる可能性があります。\n本当に終了しますか？",
                icon="warning",
            )
            if not ok:
                return
            self.log.warning("処理実行中にアプリを終了")
        self.log.info("Tag FloW 終了")
        self.root.destroy()

    # ---------- theme / navigation ----------
    def set_theme(self, name):
        self.cfg["theme"] = name
        self.theme = get_theme(name)
        save_config(self.cfg)
        self._rebuild_status_bar()
        self.refresh()

    def refresh(self):
        if self.current_screen_name == "dashboard":
            self.show_dashboard()
        elif self.current_screen_name == "cleanup":
            self.show_cleanup()
        elif self.current_screen_name == "settings":
            self.show_settings()
        elif self.current_screen_name == "verify":
            self.show_verify()
        elif self.current_screen_name == "name_lookup":
            self.show_name_lookup()
        elif self.current_screen_name == "integrity":
            self.show_integrity()
        elif self.current_screen_name == "artwork":
            self.show_artwork()
        elif self.current_screen_name == "organize":
            self.show_organize()
        elif self.current_screen_name == "albumartist":
            self.show_albumartist()
        elif self.current_screen_name == "genre":
            self.show_genre()
        elif self.current_screen_name == "tag_table":
            self.show_tag_table()
        elif self.current_screen_name == "site_fetch":
            self.show_site_fetch()

    def _clear_container(self):
        for child in self.container.winfo_children():
            child.destroy()

    def show_dashboard(self):
        self.current_screen_name = "dashboard"
        self.root.configure(bg=self.theme["bg"])
        self._clear_container()
        dashboard.build(self.container, self)

    def show_cleanup(self):
        self.current_screen_name = "cleanup"
        self._clear_container()
        cleanup_screen.build(self.container, self)

    def show_settings(self):
        self.current_screen_name = "settings"
        self._clear_container()
        settings_screen.build(self.container, self)

    def show_verify(self):
        self.current_screen_name = "verify"
        self._clear_container()
        verify_screen.build(self.container, self)

    def show_name_lookup(self):
        self.current_screen_name = "name_lookup"
        self._clear_container()
        name_lookup_screen.build(self.container, self)

    def show_integrity(self):
        self.current_screen_name = "integrity"
        self._clear_container()
        integrity_screen.build(self.container, self)

    def show_artwork(self):
        self.current_screen_name = "artwork"
        self._clear_container()
        artwork_screen.build(self.container, self)

    def show_organize(self):
        self.current_screen_name = "organize"
        self._clear_container()
        organize_screen.build(self.container, self)

    def show_albumartist(self):
        self.current_screen_name = "albumartist"
        self._clear_container()
        albumartist_screen.build(self.container, self)

    def show_genre(self):
        self.current_screen_name = "genre"
        self._clear_container()
        genre_screen.build(self.container, self)

    def show_site_fetch(self):
        self.current_screen_name = "site_fetch"
        self._clear_container()
        site_fetch_screen.build(self.container, self)

    def show_tag_table(self):
        self.current_screen_name = "tag_table"
        self._clear_container()
        tag_table_screen.build(self.container, self)

    # ---------- status bar ----------
    def _build_status_bar(self):
        self.status_bar.configure(bg=self.theme["bg_elevated"])
        for c in self.status_bar.winfo_children():
            c.destroy()
        inner = tk.Frame(self.status_bar, bg=self.theme["bg_elevated"])
        inner.pack(fill="both", expand=True, padx=20, pady=6)

        self.status_label = tk.Label(inner, textvariable=self.task_label_var, bg=self.theme["bg_elevated"],
                                      fg=self.theme["text_muted"], font=(FONT_BODY, 11), anchor="w")
        self.status_label.pack(side="left")

        self.status_cancel_btn = w.button(inner, "キャンセル", self._request_cancel, self.theme, kind="danger")
        self.status_cancel_btn.pack(side="left", padx=(10, 0))
        self.status_cancel_btn.configure(state="disabled")

        self.status_eta = tk.Label(inner, textvariable=self.task_eta_var, bg=self.theme["bg_elevated"],
                                    fg=self.theme["text_faint"], font=(FONT_MONO, 10), anchor="e")
        self.status_eta.pack(side="right")

        self.status_pct = tk.Label(inner, textvariable=self.task_pct_var, bg=self.theme["bg_elevated"],
                                    fg=self.theme["accent"], font=(FONT_MONO, 11, "bold"), anchor="e")
        self.status_pct.pack(side="right", padx=(0, 14))

        self.status_progress = w.progress_bar(inner, self.theme)
        self.status_progress.pack(side="right", fill="x", expand=True, padx=14)

    def _rebuild_status_bar(self):
        self._build_status_bar()

    def _request_cancel(self):
        if not self.task_running:
            return
        self.task_cancel_requested = True
        self.log.info("ユーザーが処理のキャンセルを要求")
        self.status_cancel_btn.configure(state="disabled")

    # ---------- background task runner ----------
    def run_task(self, name, target, on_done=None):
        if self.task_running:
            messagebox.showinfo("実行中", "他の処理が実行中です。完了をお待ちください。")
            return
        self.task_running = True
        self.task_cancel_requested = False
        self.task_start_time = time.time()
        self.task_label_var.set(name)
        self.task_pct_var.set("0%")
        self.task_eta_var.set("")
        self.status_cancel_btn.configure(state="normal")
        self.log.info(f"処理開始: {name}")

        def progress_cb(current, total, message):
            if self.task_cancel_requested:
                raise TaskCancelled()
            self.progress_queue.put(("progress", current, total, message))

        def worker():
            try:
                result = target(progress_cb)
                self.progress_queue.put(("done", result, None))
            except TaskCancelled:
                self.progress_queue.put(("cancelled", None, None))
            except Exception as e:
                self.progress_queue.put(("error", None, str(e)))

        self._on_done_cb = on_done
        threading.Thread(target=worker, daemon=True).start()

    def _poll_progress(self):
        try:
            while True:
                item = self.progress_queue.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, current, total, message = item
                    pct = int(current / total * 100) if total else 0
                    self.task_pct_var.set(f"{pct}%")
                    self.task_label_var.set(message)
                    self.status_progress["value"] = pct
                    elapsed = time.time() - self.task_start_time
                    if current > 0 and total:
                        rate = current / elapsed
                        remaining = (total - current) / rate if rate > 0 else 0
                        self.task_eta_var.set(f"残り 約{int(remaining // 60)}分{int(remaining % 60)}秒")
                elif kind == "done":
                    _, result, _ = item
                    self.task_running = False
                    self.status_cancel_btn.configure(state="disabled")
                    self.task_label_var.set("完了")
                    self.task_pct_var.set("100%")
                    self.task_eta_var.set("")
                    self.status_progress["value"] = 100
                    result_str = str(result)
                    if len(result_str) > 2000:
                        # some results (e.g. STEP06's artwork scan) embed raw
                        # image bytes -- logging those in full has blown a
                        # single log line up to several MB, forcing rotation
                        # to discard the rest of the log's history almost
                        # immediately
                        result_str = f"{result_str[:2000]}...(以下省略、全{len(result_str)}文字)"
                    self.log.info(f"処理完了: {result_str}")
                    if self._on_done_cb:
                        self._on_done_cb(result)
                elif kind == "cancelled":
                    self.task_running = False
                    self.status_cancel_btn.configure(state="disabled")
                    self.task_label_var.set("キャンセルしました")
                    self.task_pct_var.set("")
                    self.task_eta_var.set("")
                    self.status_progress["value"] = 0
                    self.log.warning("処理をキャンセル")
                elif kind == "error":
                    _, _, err = item
                    self.task_running = False
                    self.status_cancel_btn.configure(state="disabled")
                    self.task_label_var.set("エラー")
                    self.log.error(f"処理エラー: {err}")
                    messagebox.showerror("エラー", err)
        except queue.Empty:
            pass
        except Exception:
            # an error in a screen's on_done must not stop this loop: if it does, every
            # later task runs in the background but never shows progress or finishes on screen
            self.log.exception("画面の更新中にエラー")
            messagebox.showerror("エラー", "画面の更新中にエラーが起きました。詳細はログを確認してください。")
        self.root.after(150, self._poll_progress)

    # ---------- current excel path ----------
    def get_excel_path(self):
        path = self.cfg.get("excel_path") or ""
        if path and os.path.exists(path):
            return path
        return None

    def set_excel_path(self, path):
        self.cfg["excel_path"] = path
        save_config(self.cfg)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = TagFlowApp()
    app.run()
