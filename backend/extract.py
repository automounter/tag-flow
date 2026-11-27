# -*- coding: utf-8 -*-
import os
import unicodedata
import datetime
import openpyxl
from mutagen.easyid3 import EasyID3
from mutagen.id3 import ID3
from backend import tag_backup

HEADERS = ["No", "", "トラック名", "アーティスト名", "アルバム名", "トラック番号", "トラック総数",
           "ディスク番号", "ディスク総数", "日付", "ジャンル", "", "アルバムアーティスト",
           "ファイルサイズ", "フォルダパス", "フォーマット", "MP3ファイルパス（照合用・編集しないでください）",
           "アートワーク", "元ジャンルタグ（照合用・編集しないでください）"]


def nfc(s):
    return unicodedata.normalize("NFC", s) if s else s


def split_num(val):
    if not val:
        return None, None
    val = val[0] if isinstance(val, list) else val
    if "/" in val:
        a, b = val.split("/", 1)
        a = int(a) if a.strip().isdigit() else None
        b = int(b) if b.strip().isdigit() else None
        return a, b
    return (int(val) if val.strip().isdigit() else None), None


def parse_date(val):
    if not val:
        return None
    val = val[0] if isinstance(val, list) else val
    val = val.strip()
    for fmt in ("%Y-%m-%d", "%Y-%m", "%Y"):
        try:
            d = datetime.datetime.strptime(val, fmt)
            if fmt == "%Y-%m":
                return datetime.date(d.year, d.month, 1)
            if fmt == "%Y":
                return datetime.date(d.year, 1, 1)
            return d.date()
        except ValueError:
            continue
    return val


def find_mp3_files(source_folder):
    files = []
    for root, dirs, filenames in os.walk(source_folder):
        for fn in filenames:
            if fn.lower().endswith(".mp3"):
                files.append(os.path.join(root, fn))
    files.sort(key=lambda p: (os.path.dirname(p).lower(), os.path.basename(p).lower()))
    return files


def run(source_folder, excel_out_path, progress_callback=None):
    files = find_mp3_files(source_folder)
    total = len(files)
    rows = []
    errors = []

    n_snapshots = 0
    for i, f in enumerate(files, 1):
        try:
            n_snapshots += tag_backup.snapshot(f)
        except Exception as e:
            errors.append((f, f"タグの控えを保存できませんでした: {e}"))
        try:
            easy = EasyID3(f)
        except Exception as e:
            easy = {}
            errors.append((f, str(e)))

        def g(key):
            v = easy.get(key)
            return nfc(v[0]) if v else None

        title = g("title")
        artist = g("artist")
        album = g("album")
        track_no, track_total = split_num(easy.get("tracknumber"))
        disc_no, disc_total = split_num(easy.get("discnumber"))
        date_val = parse_date(easy.get("date"))
        genre = g("genre")
        albumartist = g("albumartist") or g("performer")

        if genre is None or albumartist is None:
            try:
                raw = ID3(f)
                if genre is None and "TCON" in raw:
                    genre = nfc(str(raw["TCON"]))
                if albumartist is None and "TPE2" in raw:
                    albumartist = nfc(str(raw["TPE2"]))
            except Exception:
                pass

        try:
            size_mb = round(os.path.getsize(f) / (1024 * 1024), 2)
        except OSError:
            size_mb = None

        rows.append({
            "title": title, "artist": artist, "album": album,
            "track_no": track_no, "track_total": track_total,
            "disc_no": disc_no, "disc_total": disc_total,
            "date": date_val, "genre": genre, "albumartist": albumartist,
            "size_mb": size_mb, "folder_path": nfc(os.path.dirname(f)),
            "format": "MP3", "file_path": nfc(f),
        })

        if progress_callback and i % 20 == 0:
            progress_callback(i, total, f"タグ抽出中 {i}/{total}")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IDTAG"
    ws.append(HEADERS)
    for i, r in enumerate(rows, 1):
        ws.append([
            i, None, r["title"], r["artist"], r["album"], r["track_no"], r["track_total"],
            r["disc_no"], r["disc_total"], r["date"], r["genre"], None, r["albumartist"],
            r["size_mb"], r["folder_path"], r["format"], r["file_path"], None, r["genre"],
        ])
    for col, width in zip("ABCDEFGHIJKLMNOPQRS", [6, 3, 32, 30, 13, 9, 13, 13, 13, 11, 12, 3, 22, 9, 55, 20, 60, 22, 22]):
        ws.column_dimensions[col].width = width
    for cell in ws["J"]:
        cell.number_format = "yyyy-mm-dd"

    os.makedirs(os.path.dirname(excel_out_path), exist_ok=True)
    wb.save(excel_out_path)

    if progress_callback:
        progress_callback(total, total, "完了")

    return {"total": total, "errors": errors, "excel_path": excel_out_path, "new_snapshots": n_snapshots}
