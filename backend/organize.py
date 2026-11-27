# -*- coding: utf-8 -*-
import os
import re
import shutil
import unicodedata
from collections import defaultdict
import openpyxl
from mutagen.easyid3 import EasyID3
from mutagen.id3 import ID3, APIC, ID3NoHeaderError
from backend.artwork_check import ARTWORK_DELETE
from backend.extract import nfc
from backend import tag_backup

INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')


def sanitize(name):
    if not name:
        return "Unknown"
    name = INVALID_CHARS.sub("_", name)
    name = name.strip(" .")
    return name or "Unknown"


def _planned_name(path, title=None):
    """Returns (final filename, note) for a file about to be copied: MP3s become
    '<disc>_<track>_<title>.mp3' (both 2-digit, zero-padded; disc defaults to 01
    even for single-disc albums). A file that can't get that name keeps its own,
    and note says why. `title` overrides the file's title tag, for the pre-flight
    check that runs before the tags from Excel have been written."""
    fname = os.path.basename(path)
    if not fname.lower().endswith(".mp3"):
        return fname, None
    try:
        audio = EasyID3(path)
    except Exception as e:
        return fname, f"read error: {e}"
    title = title or (audio.get("title") or [None])[0]
    track_no = _parse_num((audio.get("tracknumber") or [None])[0])
    disc_no = _parse_num((audio.get("discnumber") or [None])[0]) or 1
    if not title or track_no is None:
        missing = "タイトル" if not title else "トラック番号"
        return fname, f"{missing}がないため、ファイル名は変更されませんでした"
    return f"{disc_no:02d}_{track_no:02d}_{sanitize(str(title))}.mp3", None


def _copy_file(src_file, dest_file):
    shutil.copy2(src_file, dest_file)
    # the original may be deleted later in this run, so a short copy must not pass
    if os.path.getsize(src_file) != os.path.getsize(dest_file):
        raise OSError(f"コピー後のサイズが一致しません: {dest_file}")


def _copy_folder_merge(src_folder, dest_folder, file_actions=None, orig_dest_for_actions=None):
    """Recursively copies src_folder's contents into dest_folder under their
    final names (see _planned_name). A filename that already exists in
    dest_folder is never silently overwritten by default -- it's disambiguated
    with a "_X" suffix instead. This matters when merging same-named files from
    multiple disc subfolders of one release into a single destination: silently
    overwriting would otherwise lose a track.

    orig_dest_for_actions/file_actions let a caller override that default
    per file, for files that were already present in dest_folder *before*
    this run (see find_conflicts): file_actions is
    {(orig_dest_for_actions, rel_path): "overwrite"|"rename"}. Passing
    orig_dest_for_actions=None (the default; used whenever dest_folder is
    brand new or was just cleared for this run) means no pre-existing file
    could possibly be found, so every collision seen here is between files
    from this run alone and is always disambiguated automatically.

    Returns (list of dest file paths, [(source file path, note), ...])."""
    file_actions = file_actions or {}
    dest_files = []
    notes = []
    for root, dirs, files in os.walk(src_folder):
        rel = os.path.relpath(root, src_folder)
        dest_dir = dest_folder if rel == "." else os.path.join(dest_folder, rel)
        os.makedirs(dest_dir, exist_ok=True)
        for fname in sorted(files):
            if fname == ".DS_Store":
                continue
            src_file = os.path.join(root, fname)
            final_name, note = _planned_name(src_file)
            if note:
                notes.append((src_file, note))
            rel_path = final_name if rel == "." else os.path.join(rel, final_name)
            dest_file = os.path.join(dest_dir, final_name)
            pre_existing = orig_dest_for_actions is not None and os.path.exists(dest_file)
            if pre_existing and file_actions.get((orig_dest_for_actions, rel_path), "rename") == "overwrite":
                _copy_file(src_file, dest_file)
                dest_files.append(dest_file)
                continue
            if os.path.exists(dest_file):
                base, ext = os.path.splitext(final_name)
                n = 2
                while os.path.exists(dest_file):
                    dest_file = os.path.join(dest_dir, f"{base}_{n}{ext}")
                    n += 1
            _copy_file(src_file, dest_file)
            dest_files.append(dest_file)
    return dest_files, notes


def _load_rows_and_groups(excel_path):
    """Shared by run() and find_conflicts(): reads the IDTAG sheet and
    groups source folders the same way STEP08 does (folders sharing the
    same majority (albumartist, album) -- e.g. "Disc 1"/"Disc 2" of one
    release -- are grouped to share a single destination folder)."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb["IDTAG"]

    rows = []
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        filepath = r[16].value
        if not filepath:
            continue
        rows.append({
            "title": r[2].value, "artist": r[3].value, "album": r[4].value,
            "genre": r[10].value, "albumartist": r[12].value,
            "folder_path": r[14].value, "filepath": filepath,
        })

    folder_rows = defaultdict(list)
    for r in rows:
        folder_rows[r["folder_path"]].append(r)

    folder_albumartist = {}
    folder_album = {}
    for folder, rlist in folder_rows.items():
        aa_vals = [r["albumartist"] for r in rlist if r["albumartist"]]
        folder_albumartist[folder] = sanitize(max(set(aa_vals), key=aa_vals.count) if aa_vals else "Unknown Artist")
        album_vals = [r["album"] for r in rlist if r["album"]]
        folder_album[folder] = sanitize(max(set(album_vals), key=album_vals.count)) if album_vals else None

    album_groups = defaultdict(list)
    for folder in folder_rows:
        album_groups[(folder_albumartist[folder], folder_album[folder])].append(folder)

    return rows, folder_rows, album_groups


def group_base_name(album_name, group_folders):
    if len(group_folders) > 1 and album_name:
        return album_name
    return sanitize(os.path.basename(group_folders[0]))


def find_conflicts(excel_path, output_root):
    """Pre-flight check for STEP08's copy phase: for each album group's
    planned destination folder, reports whether it already exists on disk
    (e.g. left over from an earlier STEP08 run) and, if so, which of its
    files would collide with files this run would copy into it. The
    organize screen shows this to the user and lets them decide an action
    per folder/file; the decisions are passed back into run() as
    folder_actions/file_actions. Returns
    {"folders": [{dest_folder, albumartist, album, source_folders}, ...],
     "files": {dest_folder: [{rel_path, src_file, dest_file}, ...]}}."""
    rows, _, album_groups = _load_rows_and_groups(excel_path)
    titles = {r["filepath"]: r["title"] for r in rows}
    folders = []
    files_by_dest = {}
    for (aa_name, album_name), group_folders in album_groups.items():
        base_name = group_base_name(album_name, group_folders)
        dest = os.path.join(output_root, aa_name, base_name)
        if not os.path.isdir(dest):
            continue
        folders.append({
            "dest_folder": dest, "albumartist": aa_name,
            "album": album_name or base_name, "source_folders": list(group_folders),
        })
        conflicts = []
        for folder in group_folders:
            if not os.path.isdir(folder):
                continue
            for root, dirs, fnames in os.walk(folder):
                rel_dir = os.path.relpath(root, folder)
                for fname in fnames:
                    if fname == ".DS_Store":
                        continue
                    src_file = os.path.join(root, fname)
                    final_name, _ = _planned_name(src_file, titles.get(nfc(src_file)))
                    rel_path = final_name if rel_dir == "." else os.path.join(rel_dir, final_name)
                    dest_file = os.path.join(dest, rel_path)
                    if os.path.exists(dest_file):
                        conflicts.append({
                            "rel_path": rel_path,
                            "src_file": src_file,
                            "dest_file": dest_file,
                        })
        if conflicts:
            files_by_dest[dest] = conflicts
    return {"folders": folders, "files": files_by_dest}


def validate_for_organize(excel_path):
    """Checks whether each row has the title + track number that STEP08's
    final filename (_planned_name) needs. A row missing either is
    still copied and tagged normally -- it just keeps its original filename
    -- so this is informational, not a hard blocker. Returns
    {"total": n, "missing_title": [filepath, ...], "missing_track": [filepath, ...]}."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb["IDTAG"]

    total = 0
    missing_title = []
    missing_track = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        filepath = row[16].value
        if not filepath:
            continue
        total += 1
        if not row[2].value:
            missing_title.append(filepath)
        if not row[5].value:
            missing_track.append(filepath)

    return {"total": total, "missing_title": missing_title, "missing_track": missing_track}


def _actual_paths(folders):
    """Excel stores NFC paths, but the names on disk may not be NFC."""
    actual = {}
    for folder in folders:
        for root, dirs, fnames in os.walk(folder):
            for fname in fnames:
                path = os.path.join(root, fname)
                actual[nfc(path)] = path
    return actual


def _openable(path):
    """On the NAS a path mixing composed and decomposed accents (folder from the
    Excel, file name from the listing) has failed with "No such file" for
    names like "Björk" -- fall back to a uniformly normalized spelling."""
    if os.path.exists(path):
        return path
    for form in ("NFC", "NFD"):
        alt = unicodedata.normalize(form, path)
        if os.path.exists(alt):
            return alt
    return path


def write_tags(excel_path, progress_callback=None, id3_version="v24", strip_id3v1=False,
               artwork_overrides=None):
    """STEP07: writes the Excel's tags (and STEP06's artwork choices) into the
    original MP3 files in place. Returns the folders where a write failed, so
    STEP08 can keep those originals instead of deleting them."""
    v2_version = 3 if id3_version == "v23" else 4
    v1_mode = 0 if strip_id3v1 else 1
    artwork_overrides = artwork_overrides or {}
    rows, folder_rows, _ = _load_rows_and_groups(excel_path)
    actual_paths = _actual_paths(folder_rows)
    total_rows = len(rows)

    tag_written = tag_unchanged = 0
    tag_errors = []
    tag_failed_folders = set()
    for i, r in enumerate(rows, 1):
        src_file = actual_paths.get(nfc(r["filepath"]))
        if src_file is None:
            tag_errors.append((r["filepath"], "not found"))
            tag_failed_folders.add(r["folder_path"])
            continue
        stage = "控えの保存"
        try:
            # Excels made before tag_backup existed have no snapshot yet
            tag_backup.snapshot(src_file)
            stage = "タグの読み込み"
            src_file = _openable(src_file)
            try:
                audio = EasyID3(src_file)
            except ID3NoHeaderError:
                # a file with no tag at all: start an empty one
                audio = EasyID3()
                audio.filename = src_file
            wanted = {k: str(r[col]) for k, col in (("title", "title"), ("artist", "artist"), ("album", "album"),
                                                    ("genre", "genre"), ("albumartist", "albumartist")) if r[col]}
            override_path = artwork_overrides.get(r["folder_path"])
            # rewriting a file whose tag grows rewrites the whole MP3 over the network,
            # so files already matching the Excel are left alone (a re-run only touches the rest)
            if not override_path and all(audio.get(k) == [v] for k, v in wanted.items()):
                tag_unchanged += 1
                if progress_callback and i % 50 == 0:
                    progress_callback(i, total_rows, f"ID3タグを更新中 {i}/{total_rows}")
                continue
            for k, v in wanted.items():
                audio[k] = v
            stage = "タグの書き込み"
            audio.save(v2_version=v2_version, v1=v1_mode)

            stage = "アートワークの書き込み"
            if override_path == ARTWORK_DELETE:
                id3 = ID3(src_file)
                id3.delall("APIC")
                id3.save(src_file, v2_version=v2_version, v1=v1_mode)
            elif override_path and os.path.exists(override_path):
                with open(override_path, "rb") as imgf:
                    img_data = imgf.read()
                mime = "image/png" if override_path.lower().endswith(".png") else "image/jpeg"
                id3 = ID3(src_file)
                id3.delall("APIC")
                id3.add(APIC(encoding=3, mime=mime, type=3, desc="Cover", data=img_data))
                id3.save(src_file, v2_version=v2_version, v1=v1_mode)

            tag_written += 1
        except Exception as e:
            tag_errors.append((src_file, f"{stage}: {e}"))
            tag_failed_folders.add(r["folder_path"])
        if progress_callback and i % 50 == 0:
            progress_callback(i, total_rows, f"ID3タグを更新中 {i}/{total_rows}")

    if progress_callback:
        progress_callback(total_rows, total_rows, "完了")
    return {"tags_written": tag_written, "tags_unchanged": tag_unchanged, "tags_total": total_rows,
            "tag_errors": tag_errors,
            "tag_failed_folders": sorted(f for f in tag_failed_folders if f)}


def restore_tags(excel_path, progress_callback=None):
    """Puts back the original tags (tag_backup snapshots) of the Excel's files."""
    rows, folder_rows, _ = _load_rows_and_groups(excel_path)
    actual_paths = _actual_paths(folder_rows)
    paths, not_found = [], []
    for r in rows:
        p = actual_paths.get(nfc(r["filepath"]))
        if p:
            paths.append(p)
        else:
            not_found.append(r["filepath"])
    result = tag_backup.restore_files(paths, progress_callback)
    result["not_found"] = not_found
    result["total"] += len(not_found)
    return result


def run(excel_path, output_root, progress_callback=None, do_delete_originals=False,
        source_root=None, folder_actions=None, file_actions=None, keep_folders=None):
    """STEP08: copies the (already tagged, see write_tags) original files to the
    output under their final names, then (only if do_delete_originals) deletes
    the originals. Originals are kept when their copy failed or when they are in
    keep_folders (STEP07 couldn't write their tags). `source_root` (the
    configured MP3 source folder) is never itself removed by the empty-parent
    cleanup below, even if every album under it ends up deleted.

    folder_actions/file_actions are the user's decisions for conflicts
    reported by find_conflicts() (see the organize screen's confirmation
    panel): folder_actions is {dest_folder: "overwrite"|"merge"|"new"}
    (default "merge" for any conflicting folder left unspecified) and
    file_actions is {(dest_folder, rel_path): "overwrite"|"rename"}
    (default "rename") for files inside a "merge" folder that already
    existed there before this run."""
    folder_actions = folder_actions or {}
    file_actions = file_actions or {}
    tag_failed_folders = set(keep_folders or [])
    rows, folder_rows, album_groups = _load_rows_and_groups(excel_path)

    os.makedirs(output_root, exist_ok=True)

    total_folders = len(folder_rows)
    total_steps = total_folders

    dest_folder_map = {}
    copy_errors = []
    renamed_dirs = set()
    rename_notes = []
    i = 0
    for (aa_name, album_name), group_folders in album_groups.items():
        base_name = group_base_name(album_name, group_folders)
        artist_dir = os.path.join(output_root, aa_name)
        dest = os.path.join(artist_dir, base_name)
        dest_for_file_actions = None
        try:
            os.makedirs(artist_dir, exist_ok=True)
            if os.path.isdir(dest):
                # this destination already exists (e.g. left over from an earlier
                # STEP08 run) -- apply the user's conflict-resolution choice
                # instead of silently merging or auto-suffixing
                action = folder_actions.get(dest, "merge")
                if action == "overwrite":
                    shutil.rmtree(dest)
                    os.makedirs(dest, exist_ok=True)
                elif action == "new":
                    suffix = 2
                    orig_dest = dest
                    while os.path.exists(dest):
                        dest = f"{orig_dest} ({suffix})"
                        suffix += 1
                else:  # "merge"
                    dest_for_file_actions = dest
        except Exception as e:
            # one album's destination failing must not abort the albums after it
            for folder in group_folders:
                copy_errors.append((folder, str(e)))
            i += len(group_folders)
            continue

        for folder in group_folders:
            i += 1
            if not os.path.isdir(folder):
                copy_errors.append((folder, "source not found"))
            else:
                try:
                    dest_files, notes = _copy_folder_merge(
                        folder, dest, file_actions=file_actions,
                        orig_dest_for_actions=dest_for_file_actions)
                    dest_folder_map[folder] = dest
                    renamed_dirs.update(os.path.dirname(f) for f in dest_files if f.lower().endswith(".mp3"))
                    rename_notes.extend(notes)
                except Exception as e:
                    copy_errors.append((folder, str(e)))
            if progress_callback:
                progress_callback(i, total_steps, f"フォルダコピー中 {i}/{total_folders}")

    deleted_originals = []
    delete_errors = []
    if do_delete_originals:
        source_root_norm = os.path.normpath(os.path.abspath(source_root)) if source_root else None
        for i, src_folder in enumerate(dest_folder_map, 1):
            if src_folder in tag_failed_folders:
                delete_errors.append((src_folder, "STEP 07 でタグを更新できなかったファイルがあるため、元データを残しました"))
                continue
            try:
                shutil.rmtree(src_folder)
                deleted_originals.append(src_folder)
                # a merged multi-disc release can leave its now-empty parent
                # folder (e.g. "Big Release/" once "Disc 1" and "Disc 2" are
                # both gone) behind -- remove it too, but only if truly empty,
                # and never the configured source root itself (deleting enough
                # albums in one run could otherwise empty it out completely)
                parent = os.path.dirname(src_folder)
                try:
                    if (parent and os.path.normpath(os.path.abspath(parent)) != source_root_norm
                            and not os.listdir(parent)):
                        os.rmdir(parent)
                except OSError:
                    pass
            except Exception as e:
                delete_errors.append((src_folder, str(e)))
            if progress_callback:
                progress_callback(i, len(dest_folder_map), f"元データを削除中 {i}/{len(dest_folder_map)}")

    if progress_callback:
        progress_callback(1, 1, "完了")

    return {
        "folders_copied": len(dest_folder_map), "folders_total": total_folders, "copy_errors": copy_errors,
        "renamed": len(renamed_dirs), "rename_errors": rename_notes,
        "originals_deleted": len(deleted_originals), "delete_errors": delete_errors,
    }


def _parse_num(value):
    """'2/9' or '2' -> 2. Returns None if missing/unparsable."""
    if not value:
        return None
    head = str(value).split('/')[0].strip()
    return int(head) if head.isdigit() else None
