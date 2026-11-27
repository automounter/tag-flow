# -*- coding: utf-8 -*-
import os
import re

JP_RE = re.compile(r'[぀-ゟ゠-ヿ一-鿿]')


def find_japanese_subfolders(source_folder):
    """Returns the paths of the immediate subfolders of source_folder
    (typically album folders) whose name contains Japanese characters. Used
    by the STEP01 screen to show a heads-up notice; only checks direct
    children, not the full tree, to stay fast on large/network-mounted
    libraries."""
    if not source_folder or not os.path.isdir(source_folder):
        return []
    try:
        entries = os.listdir(source_folder)
    except OSError:
        return []
    paths = [os.path.join(source_folder, name) for name in entries if JP_RE.search(name)]
    return sorted(p for p in paths if os.path.isdir(p))


# kinds of files STEP01 lists, in display order: (key, heading, checked by default)
CATEGORIES = [
    ("appledouble", "Mac の補足ファイル(._ で始まる)", True),
    ("broken", "中身が音声でない小さいファイル(0バイト・読めない MP3)", True),
    ("download_extra", "ダウンロードの付属ファイル(歌詞 .lrc・途中のファイルなど)", True),
    ("small_audio", "小さいけれど再生できる MP3(短い効果音など)", False),
    ("other", "MP3 以外のファイル(画像・ほかの形式の音声など)", False),
]


# left behind by music downloaders: synced lyrics, unfinished downloads
DOWNLOAD_EXTRA_EXTS = {".lrc", ".part", ".tmp", ".crdownload", ".download"}


def _playable(path):
    try:
        from mutagen.mp3 import MP3
        return MP3(path).info.length > 0.5
    except Exception:
        return False


def scan_small_files(source_folder, threshold_kb=10, exclude_dirs=("★Output",)):
    """Lists files that are probably leftovers of downloads/copies, each with a category
    (see CATEGORIES). Only "appledouble", "broken" and "download_extra" are pre-selected:
    a small MP3 that still plays, or any non-MP3 file (a cover image, another audio
    format), is shown but left for the user to decide."""
    threshold_bytes = threshold_kb * 1024
    results = []
    for root, dirs, files in os.walk(source_folder):
        dirs[:] = [d for d in dirs if d not in exclude_dirs and not d.startswith("★")]
        for fn in files:
            if fn == ".DS_Store":
                continue
            path = os.path.join(root, fn)
            try:
                size = os.path.getsize(path)
            except OSError:
                continue
            is_mp3 = fn.lower().endswith(".mp3")
            if fn.startswith("._"):
                category = "appledouble"
            elif size == 0:
                category = "broken"
            elif os.path.splitext(fn)[1].lower() in DOWNLOAD_EXTRA_EXTS:
                category = "download_extra"
            elif is_mp3:
                if size > threshold_bytes:
                    continue
                category = "small_audio" if _playable(path) else "broken"
            else:
                category = "other"
            results.append({"path": path, "name": fn, "album": os.path.basename(root),
                            "size": size, "category": category})
    order = {key: i for i, (key, _, _) in enumerate(CATEGORIES)}
    results.sort(key=lambda x: (order[x["category"]], x["album"], x["name"]))
    return results


def delete_files(paths):
    deleted = []
    errors = []
    for p in paths:
        try:
            os.remove(p)
            deleted.append(p)
        except OSError as e:
            errors.append((p, str(e)))
    return deleted, errors
