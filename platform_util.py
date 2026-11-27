# -*- coding: utf-8 -*-
"""OS ごとの違いをここにまとめる（2026-11-26。Windows 版・Linux 版（Debian 系・RHEL 系）のため）。
Mac での動きは今まで通り。"""
import os
import subprocess
import sys

IS_MAC = sys.platform == "darwin"
IS_WINDOWS = sys.platform.startswith("win")


def config_dir():
    if IS_MAC:
        return os.path.expanduser("~/Library/Application Support/TagFlow")
    if IS_WINDOWS:
        return os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "TagFlow")
    return os.path.join(os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config"), "TagFlow")


def open_path(path):
    """ファイルやフォルダを、その OS の既定のアプリで開く（Excel の記録など）。"""
    if IS_MAC:
        subprocess.run(["open", path])
    elif IS_WINDOWS:
        os.startfile(path)
    else:
        subprocess.Popen(["xdg-open", path])


# 字体：その OS に必ずある、日本語の出る字体を選ぶ
if IS_MAC:
    FONT_UI, FONT_FIXED = "Helvetica Neue", "Menlo"
elif IS_WINDOWS:
    FONT_UI, FONT_FIXED = "Yu Gothic UI", "Consolas"
else:
    FONT_UI, FONT_FIXED = "Noto Sans CJK JP", "DejaVu Sans Mono"
