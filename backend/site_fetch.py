# -*- coding: utf-8 -*-
"""サイト（how-to-make-house.com の MP3 データベース）から MP3 タグを取り寄せて書く（2026-11-25）。

流れ：フォルダを読む（scan）→ アルバム（フォルダ）ごとにサイトの窓口 api/match.php に問い合わせる（ask）
→ 人が選ぶ → 元のタグを控え（tag_backup。STEP02 と同じ「元に戻す」の控え）に取ってから書く（write）。
サイトの答えの grade：A は書いてよい、B は人が見て決める、C は書かない。
"""
import datetime
import json
import os
import urllib.error
import urllib.request

from mutagen.id3 import ID3, ID3NoHeaderError, APIC, TXXX, Frames
from mutagen.mp3 import MP3

from backend import tag_backup

DEFAULT_URL = "https://how-to-make-house.com/mp3db/api/match.php"
BATCH = 20  # 1回に問い合わせる枚数（サイトは 50 枚まで）

# 書ける項目。画面のチェックの並びもこの順
FIELDS = [
    ("TIT2", "曲名"), ("TPE1", "アーティスト"), ("TPE2", "アルバムアーティスト"), ("TALB", "アルバム"),
    ("TRCK", "曲番号"), ("TPOS", "ディスク番号"), ("TDRC", "発売日"), ("TCON", "ジャンル"),
    ("TCOM", "作曲者"), ("TSRC", "ISRC"), ("TBPM", "BPM"), ("TKEY", "キー"),
    ("TXXX:MP3DB_ALBUM_ID", "アルバムの番号"), ("APIC", "ジャケット"),
]


def _text(tags, key):
    fr = tags.get(key)
    return str(fr.text[0]) if fr is not None and getattr(fr, "text", None) else ""


def _num(s):
    s = (s or "").split("/")[0].strip()
    return int(s) if s.isdigit() else 0


def scan(root, progress=None):
    """root より下の MP3 を、フォルダごとにまとめて読む。"""
    folders = []
    for top, dirs, names in os.walk(root):
        dirs[:] = sorted(d for d in dirs if not d.startswith((".", "@")))
        mp3s = sorted(n for n in names if n.lower().endswith(".mp3") and not n.startswith("._"))
        if mp3s:
            folders.append((top, mp3s))
    total = sum(len(m) for _, m in folders)
    albums, done = [], 0
    for top, mp3s in folders:
        tracks = []
        album_ids = set()
        for n in mp3s:
            path = os.path.join(top, n)
            t = {"file": n, "path": path, "title": "", "artist": "", "album": "", "track": 0, "disc": 0, "seconds": 0}
            try:
                tags = ID3(path)
                t.update(title=_text(tags, "TIT2"), artist=_text(tags, "TPE1"), album=_text(tags, "TALB"),
                         track=_num(_text(tags, "TRCK")), disc=_num(_text(tags, "TPOS")))
                for fr in tags.getall("TXXX"):
                    if fr.desc == "MP3DB_ALBUM_ID" and fr.text:
                        album_ids.add(str(fr.text[0]))
            except ID3NoHeaderError:
                pass
            except Exception:
                pass
            try:
                t["seconds"] = round(MP3(path).info.length, 1)
            except Exception:
                pass
            tracks.append(t)
            done += 1
            if progress:
                progress(done, total, f"読み込み中 {done}/{total}：{os.path.basename(top)}")
        albums.append({"folder": top, "album_id": album_ids.pop() if len(album_ids) == 1 else "", "tracks": tracks})
    return albums


def _post(url, key, payload):
    req = urllib.request.Request(url, data=json.dumps(payload, ensure_ascii=False).encode("utf-8"), method="POST",
                                 headers={"Content-Type": "application/json", "Authorization": f"Bearer {key}",
                                          "User-Agent": "TagFlow"})
    try:
        with urllib.request.urlopen(req, timeout=60) as res:
            return json.loads(res.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read().decode("utf-8")).get("error", "")
        except Exception:
            msg = ""
        raise RuntimeError(f"サイトの窓口がエラーを返しました（{e.code}）：{msg}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"サイトにつながりません：{e.reason}") from e


def ask(albums, url, key, progress=None):
    """アルバムごとにサイトへ問い合わせ、各アルバムに result を付けて返す。"""
    if not key:
        raise RuntimeError("サイトの鍵が設定されていません。設定の画面で、マイページの「TagFlow の鍵」で作った鍵を入れてください。")
    for i in range(0, len(albums), BATCH):
        part = albums[i:i + BATCH]
        payload = {"albums": [{"folder": os.path.basename(a["folder"]), "album_id": a["album_id"],
                               "tracks": [{k: t[k] for k in ("file", "title", "artist", "album", "track", "disc", "seconds")}
                                          for t in a["tracks"]]} for a in part]}
        res = _post(url, key, payload)
        for a, r in zip(part, res.get("results", [])):
            a["result"] = r
        if progress:
            progress(min(i + BATCH, len(albums)), len(albums), f"サイトに問い合わせ中 {min(i + BATCH, len(albums))}/{len(albums)} 枚")
    return albums


def _frame(key, value):
    if key.startswith("TXXX:"):
        return TXXX(encoding=3, desc=key[5:], text=[value])
    return Frames[key](encoding=3, text=[value])


def _cover(url, cache):
    if url not in cache:
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "TagFlow"}), timeout=30) as res:
                data = res.read()
            mime = "image/png" if data[:8] == b"\x89PNG\r\n\x1a\n" else "image/jpeg"
            cache[url] = (mime, data)
        except Exception:
            cache[url] = None
    return cache[url]


def write(albums, fields, id3_version="v23", cover_only_if_missing=True, progress=None):
    """選ばれたアルバム（album["selected"]）の、合った曲に書く。書く前に元のタグを tag_backup に取る
    （一度取ったファイルは取り直さない。いちばん初めのタグに戻せる）。戻り値：記録の行（ファイル・項目・前・後）など。"""
    v2 = 3 if id3_version == "v23" else 4
    targets = [(a, t, r) for a in albums if a.get("selected") and a.get("result")
               for t, r in zip(a["tracks"], a["result"].get("tracks", [])) if r.get("track_id") and r.get("tags")]
    log, covers = [], {}
    for n, (a, t, r) in enumerate(targets, 1):
        path = t["path"]
        tag_backup.snapshot(path)
        try:
            tags = ID3(path)
        except ID3NoHeaderError:
            tags = ID3()
        changed = False
        for key in fields:
            value = r["tags"].get(key)
            if not value:
                continue
            if key == "APIC":
                if cover_only_if_missing and tags.getall("APIC"):
                    continue
                img = _cover(value, covers)
                if not img:
                    continue
                tags.delall("APIC")
                tags.add(APIC(encoding=3, mime=img[0], type=3, desc="Cover", data=img[1]))
                log.append((path, key, "", "（画像）"))
                changed = True
                continue
            before = ""
            if key.startswith("TXXX:"):
                before = next((str(f.text[0]) for f in tags.getall("TXXX") if f.desc == key[5:] and f.text), "")
            else:
                before = _text(tags, key)
            if before == str(value):
                continue
            if key == "TBPM":
                try:
                    if abs(float(before) - float(value)) < 0.01:  # 「110.540000」と「110.54」は同じとみる
                        continue
                except ValueError:
                    pass
            if key.startswith("TXXX:"):
                tags.delall(key)
                tags.add(_frame(key, str(value)))
            else:
                tags.setall(key, [_frame(key, str(value))])
            log.append((path, key, before, str(value)))
            changed = True
        if changed:
            # 更新日時は変わってよい（NAS への写しと DB の読み直しは、大きさと更新日時で変わったファイルを見つけるため）
            tags.save(path, v2_version=v2)
        if progress:
            progress(n, len(targets), f"書き込み中 {n}/{len(targets)}：{os.path.basename(path)}")
    return {"rows": log, "backup_dir": tag_backup.BACKUP_DIR, "files": len(targets)}


def save_log(result, folder):
    """書いた記録を Excel に残す（STEP08 と同じく、あとで見返せるように）。"""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "取り寄せの記録"
    ws.append(["ファイル", "項目", "前", "後"])
    for row in result["rows"]:
        ws.append(list(row))
    ws.column_dimensions["A"].width = 80
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 40
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"サイトから取り寄せ_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb.save(path)
    return path
