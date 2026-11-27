# -*- coding: utf-8 -*-
"""Keeps a copy of each MP3's original ID3 tag, taken the first time Tag FloW
reads the file (STEP02), so "元に戻す" can put it back exactly -- including
fields that were empty, artwork, other frames and the ID3 version.

Layout under BACKUP_DIR:
  meta/<key>.json  path, whether a v2/v1 tag existed, v2 version, artwork refs
  tags/<key>.id3   the v2 tag without artwork (mutagen writes a tag-only file)
  art/<sha1>.bin   artwork, shared between tracks with the same image
A file is never snapshotted twice: re-running STEP02 after tags were written
must not replace the original with the edited state."""
import os
import json
import hashlib
import datetime
import unicodedata
from mutagen.id3 import ID3, APIC, ID3NoHeaderError
from mutagen.id3 import delete as delete_id3

from config import CONFIG_DIR

BACKUP_DIR = os.path.join(CONFIG_DIR, "タグの控え")


def nfc(s):
    # own copy: importing it from backend.extract would be circular
    return unicodedata.normalize("NFC", s) if s else s


def _key(path):
    return hashlib.sha1(nfc(path).encode("utf-8")).hexdigest()


def _paths(path):
    k = _key(path)
    return (os.path.join(BACKUP_DIR, "meta", f"{k}.json"),
            os.path.join(BACKUP_DIR, "tags", f"{k}.id3"))


def has_snapshot(path):
    return os.path.exists(_paths(path)[0])


def _has_v1(path):
    with open(path, "rb") as f:
        f.seek(0, os.SEEK_END)
        if f.tell() < 128:
            return False
        f.seek(-128, os.SEEK_END)
        return f.read(3) == b"TAG"


def snapshot(path):
    """Saves the file's current tag unless a snapshot already exists.
    Returns True when a new snapshot was written."""
    meta_path, tag_path = _paths(path)
    if os.path.exists(meta_path):
        return False
    for sub in ("meta", "tags", "art"):
        os.makedirs(os.path.join(BACKUP_DIR, sub), exist_ok=True)

    meta = {"path": nfc(path), "had_v1": _has_v1(path), "had_v2": False, "v2_version": None,
            "artwork": [], "saved_at": datetime.datetime.now().isoformat(timespec="seconds")}
    try:
        tags = ID3(path)
    except ID3NoHeaderError:
        tags = None
    if tags is not None:
        meta["had_v2"] = True
        meta["v2_version"] = tags.version[1]
        for frame in tags.getall("APIC"):
            digest = hashlib.sha1(frame.data).hexdigest()
            art_path = os.path.join(BACKUP_DIR, "art", f"{digest}.bin")
            if not os.path.exists(art_path):
                with open(art_path, "wb") as f:
                    f.write(frame.data)
            meta["artwork"].append({"sha1": digest, "mime": frame.mime, "type": int(frame.type),
                                    "desc": frame.desc, "encoding": int(frame.encoding)})
        tags.delall("APIC")
        open(tag_path, "wb").close()
        # v2.2 can't be written by mutagen; it is stored (and restored) as v2.4
        tags.save(tag_path, v1=0, v2_version=3 if meta["v2_version"] == 3 else 4)
    # meta last: its presence marks the snapshot as complete
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)
    return True


def restore(path):
    """Writes the snapshot back into the file. Raises FileNotFoundError when
    there is no snapshot."""
    meta_path, tag_path = _paths(path)
    with open(meta_path, encoding="utf-8") as f:
        meta = json.load(f)
    if not meta["had_v2"]:
        delete_id3(path, delete_v1=not meta["had_v1"], delete_v2=True)
        return
    tags = ID3(tag_path)
    for art in meta["artwork"]:
        with open(os.path.join(BACKUP_DIR, "art", f"{art['sha1']}.bin"), "rb") as f:
            data = f.read()
        tags.add(APIC(encoding=art["encoding"], mime=art["mime"], type=art["type"],
                      desc=art["desc"], data=data))
    tags.save(path, v1=2 if meta["had_v1"] else 0, v2_version=3 if meta["v2_version"] == 3 else 4)


def restore_files(paths, progress_callback=None):
    restored, missing, errors = 0, [], []
    total = len(paths)
    for i, p in enumerate(paths, 1):
        try:
            if not has_snapshot(p):
                missing.append(p)
            else:
                restore(p)
                restored += 1
        except Exception as e:
            errors.append((p, str(e)))
        if progress_callback and i % 50 == 0:
            progress_callback(i, total, f"タグを元に戻しています {i}/{total}")
    if progress_callback:
        progress_callback(total, total, "完了")
    return {"restored": restored, "total": total, "no_snapshot": missing, "errors": errors}
