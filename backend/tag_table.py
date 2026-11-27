# -*- coding: utf-8 -*-
import os
import re
from mutagen.easyid3 import EasyID3
from mutagen.id3 import ID3, COMM, ID3NoHeaderError

from backend.extract import find_mp3_files, nfc
from backend.organize import INVALID_CHARS

# (EasyID3 key or "comment", 表示名)
FIELDS = [
    ("title", "タイトル"),
    ("artist", "アーティスト"),
    ("album", "アルバム"),
    ("albumartist", "アルバムアーティスト"),
    ("tracknumber", "トラック"),
    ("discnumber", "ディスク"),
    ("date", "年"),
    ("genre", "ジャンル"),
    ("composer", "作曲者"),
    ("comment", "コメント"),
    ("bpm", "BPM"),
]
FIELD_KEYS = [k for k, _ in FIELDS]

# %token% -> field key, for the filename <-> tag patterns
PATTERN_TOKENS = {
    "title": "title", "artist": "artist", "album": "album", "albumartist": "albumartist",
    "track": "tracknumber", "disc": "discnumber", "year": "date", "genre": "genre",
    "composer": "composer", "comment": "comment", "bpm": "bpm",
}
TOKEN_RE = re.compile(r"%([a-z]+)%")


# frame IDs behind the EasyID3 keys, so reading needs only one pass over the
# file (EasyID3 plus a separate ID3 read for the comment doubled the time on the NAS)
FRAME_IDS = {
    "title": "TIT2", "artist": "TPE1", "album": "TALB", "albumartist": "TPE2",
    "tracknumber": "TRCK", "discnumber": "TPOS", "date": "TDRC", "genre": "TCON",
    "composer": "TCOM", "bpm": "TBPM",
}


def read_file(path):
    try:
        id3 = ID3(path)
    except Exception:
        id3 = None
    tags = {}
    for key, frame_id in FRAME_IDS.items():
        frame = id3.get(frame_id) if id3 else None
        if frame is None or not frame.text:
            tags[key] = ""
        elif frame_id == "TCON":
            tags[key] = nfc(frame.genres[0]) if frame.genres else ""
        else:
            tags[key] = nfc(str(frame.text[0]))
    tags["comment"] = ""
    # prefer the plain (no description) comment; iTunes adds extra COMM frames like iTunNORM
    for f in (id3.getall("COMM") if id3 else []):
        if not f.desc:
            tags["comment"] = nfc(str(f))
            break
    return tags


def scan(folder, progress_callback=None):
    """Returns [{"path", "folder", "filename", <field key>: str, ...}] for every
    MP3 under folder. "folder" is relative to folder, "filename" has no extension."""
    files = find_mp3_files(folder)
    total = len(files)
    rows = []
    for i, path in enumerate(files, 1):
        rel_folder = os.path.relpath(os.path.dirname(path), folder)
        row = {
            "path": path,
            "folder": "" if rel_folder == "." else nfc(rel_folder),
            "filename": nfc(os.path.splitext(os.path.basename(path))[0]),
        }
        row.update(read_file(path))
        rows.append(row)
        if progress_callback and i % 20 == 0:
            progress_callback(i, total, f"タグ読み込み中 {i}/{total}")
    if progress_callback:
        progress_callback(total, total, "完了")
    return rows


def _write_tags(path, tags, v2_version, v1_mode):
    easy_tags = {k: v for k, v in tags.items() if k != "comment"}
    if easy_tags:
        try:
            audio = EasyID3(path)
        except ID3NoHeaderError:
            audio = EasyID3()
        for key, value in easy_tags.items():
            if value:
                audio[key] = value
            elif key in audio:
                del audio[key]
        audio.save(path, v2_version=v2_version, v1=v1_mode)

    if "comment" in tags:
        try:
            id3 = ID3(path)
        except ID3NoHeaderError:
            id3 = ID3()
        for f in id3.getall("COMM"):
            if not f.desc:
                id3.delall(f.HashKey)
        if tags["comment"]:
            id3.add(COMM(encoding=3, lang="eng", desc="", text=tags["comment"]))
        id3.save(path, v2_version=v2_version, v1=v1_mode)


def save(changes, id3_version="v24", strip_id3v1=False, progress_callback=None):
    """changes: [{"path", "tags": {field key: new value (only changed ones)},
    "filename": new name without extension or None}]. Tags are written first,
    then the file is renamed. Returns {"saved", "renamed", "errors", "renames"}
    where renames maps old path -> new path."""
    v2_version = 3 if id3_version == "v23" else 4
    v1_mode = 0 if strip_id3v1 else 1
    total = len(changes)
    saved = 0
    errors = []
    renames = {}

    for i, ch in enumerate(changes, 1):
        path = ch["path"]
        try:
            if ch.get("tags"):
                _write_tags(path, ch["tags"], v2_version, v1_mode)
            saved += 1
        except Exception as e:
            errors.append((path, f"タグ書き込み失敗: {e}"))
            continue

        new_name = ch.get("filename")
        if new_name:
            new_path = os.path.join(os.path.dirname(path), new_name + os.path.splitext(path)[1])
            # case-only renames are the same file on the case-insensitive macOS filesystem
            if os.path.exists(new_path) and new_path.lower() != path.lower():
                errors.append((path, f"同名のファイルがあるため名前を変更できません: {os.path.basename(new_path)}"))
            else:
                try:
                    os.rename(path, new_path)
                    renames[path] = new_path
                except OSError as e:
                    errors.append((path, f"名前の変更失敗: {e}"))

        if progress_callback and i % 20 == 0:
            progress_callback(i, total, f"保存中 {i}/{total}")

    if progress_callback:
        progress_callback(total, total, "完了")
    return {"saved": saved, "renamed": len(renames), "errors": errors, "renames": renames}


def invalid_filename(name):
    """Returns an error message for a filename the filesystem can't take, else None."""
    if not name.strip(" ."):
        return "ファイル名が空です"
    if INVALID_CHARS.search(name):
        return '使えない文字(\\ / : * ? " < > |)が含まれています'
    return None


def parse_filename(pattern, filename):
    """Applies a pattern like "%track% - %artist% - %title%" to filename.
    Returns {field key: value} or None if the filename doesn't match."""
    regex = ""
    keys = []
    pos = 0
    for m in TOKEN_RE.finditer(pattern):
        key = PATTERN_TOKENS.get(m.group(1))
        if key is None:
            raise ValueError(f"不明なタグ名です: %{m.group(1)}%")
        regex += re.escape(pattern[pos:m.start()]) + "(.+?)"
        keys.append(key)
        pos = m.end()
    regex += re.escape(pattern[pos:])
    m = re.fullmatch(regex, filename)
    if not m:
        return None
    return {k: v.strip() for k, v in zip(keys, m.groups())}


def build_filename(pattern, tags):
    """Fills a pattern like "%track% %title%" from tags. The track/disc number
    drops any "/total" part, and the track number is zero-padded to 2 digits."""
    def repl(m):
        key = PATTERN_TOKENS.get(m.group(1))
        if key is None:
            raise ValueError(f"不明なタグ名です: %{m.group(1)}%")
        value = tags.get(key, "")
        if key in ("tracknumber", "discnumber"):
            value = value.split("/")[0].strip()
            if key == "tracknumber" and value.isdigit():
                value = value.zfill(2)
        return value
    name = TOKEN_RE.sub(repl, pattern)
    return INVALID_CHARS.sub("_", name).strip(" .")
