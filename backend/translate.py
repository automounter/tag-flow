# -*- coding: utf-8 -*-
import re
import sys
import os
import unicodedata
from collections import Counter, defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from data.mappings import load_translation_dicts

JP_RE = re.compile(r'[぀-ゟ゠-ヿ一-鿿]')
GREEN = PatternFill(fgColor="C6EFCE", fill_type="solid")
BLUE = PatternFill(fgColor="DDEBF7", fill_type="solid")
LAVENDER = PatternFill(fgColor="E6D9F2", fill_type="solid")  # already-English, formatting unified

OVERRIDES_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                              "data", "translation_overrides.xlsx")

SHIHO_JP_ALBUMS = {"ぎりぎり・誘惑 - Single", "ルシファー - EP"}
WHITESPACE_RE = re.compile(r'\s+')


def clean_format(s):
    """Normalize an already-English string: full-width->half-width, trim, collapse spaces."""
    if not s:
        return s
    result = unicodedata.normalize("NFKC", s)
    result = WHITESPACE_RE.sub(' ', result).strip()
    return result


def normalize_english_fields(rows, progress_callback=None):
    """Second pass: for cells that are NOT Japanese, fix whitespace/width glitches and
    unify inconsistent capitalization of the same artist/album across rows (most common
    variant wins). Does not force Title Case on already-correct English text, since some
    artist names are intentionally stylized (e.g. deadmau5)."""
    total = len(rows)

    # pass 1: collect canonical casing per normalized key for artist / album
    artist_variants = defaultdict(Counter)
    album_variants = defaultdict(Counter)
    for row in rows:
        artist, album = row[3].value, row[4].value
        if artist and not JP_RE.search(str(artist)):
            cleaned = clean_format(str(artist))
            artist_variants[cleaned.lower()][cleaned] += 1
        if album and not JP_RE.search(str(album)):
            cleaned = clean_format(str(album))
            album_variants[cleaned.lower()][cleaned] += 1

    artist_canonical = {k: c.most_common(1)[0][0] for k, c in artist_variants.items()}
    album_canonical = {k: c.most_common(1)[0][0] for k, c in album_variants.items()}

    n_fixed = 0
    for i, row in enumerate(rows, 1):
        title_cell, artist_cell, album_cell = row[2], row[3], row[4]

        title = title_cell.value
        if title and not JP_RE.search(str(title)):
            cleaned = clean_format(str(title))
            if cleaned != title:
                title_cell.value = cleaned
                title_cell.fill = LAVENDER
                title_cell.font = Font(bold=True)
                n_fixed += 1

        artist = artist_cell.value
        if artist and not JP_RE.search(str(artist)):
            cleaned = clean_format(str(artist))
            canonical = artist_canonical.get(cleaned.lower(), cleaned)
            if canonical != artist:
                artist_cell.value = canonical
                artist_cell.fill = LAVENDER
                artist_cell.font = Font(bold=True)
                n_fixed += 1

        album = album_cell.value
        if album and not JP_RE.search(str(album)):
            cleaned = clean_format(str(album))
            canonical = album_canonical.get(cleaned.lower(), cleaned)
            if canonical != album:
                album_cell.value = canonical
                album_cell.fill = LAVENDER
                album_cell.font = Font(bold=True)
                n_fixed += 1

        if progress_callback and i % 200 == 0:
            progress_callback(i, total, f"表記整合性チェック中 {i}/{total}")

    return n_fixed


def title_case(s):
    if not s:
        return s
    words = s.split(' ')
    out = []
    for w in words:
        letters = [c for c in w if c.isalpha()]
        if letters and all(c.isupper() for c in letters):
            out.append(w)
            continue
        idx = None
        for i, c in enumerate(w):
            if c.isalpha():
                idx = i
                break
        if idx is None:
            out.append(w)
        else:
            out.append(w[:idx] + w[idx].upper() + w[idx + 1:])
    return ' '.join(out)


def convert_artist(artist, artist_core):
    if not artist or not JP_RE.search(artist):
        return artist, False
    parts = re.split(r'(\s*[,&]\s*)', artist)
    out, changed = [], False
    for p in parts:
        if re.fullmatch(r'\s*[,&]\s*', p):
            out.append(p)
            continue
        key = p.strip()
        if key in artist_core:
            out.append(artist_core[key])
            changed = True
        else:
            out.append(key)
    return "".join(out), changed


def convert_album(album, album_map):
    if not album or not JP_RE.search(album):
        return album, False
    if album in SHIHO_JP_ALBUMS:
        return album, False
    if album in album_map:
        return album_map[album], True
    return album, False


def convert_title(title, artist, title_map):
    if not title or not JP_RE.search(title):
        return title, False
    if artist == "SHIHO":
        return title, False
    if title in title_map:
        return title_map[title], True
    keys = sorted(title_map.keys(), key=len, reverse=True)
    result, changed = title, False
    for k in keys:
        if k in result:
            result = result.replace(k, title_map[k])
            changed = True
    return result, changed


def load_overrides():
    """Exact-match corrections checked before the dictionary (whose partial
    replacement can mangle titles, e.g. "If・You・アスクト・ミー・トゥ")."""
    if not os.path.exists(OVERRIDES_FILE):
        return {}, {}, {}
    wb = openpyxl.load_workbook(OVERRIDES_FILE, read_only=True)
    artists = {a: b for a, b in wb["アーティスト"].iter_rows(min_row=2, values_only=True) if a and b}
    albums = {a: b for a, b in wb["アルバム"].iter_rows(min_row=2, values_only=True) if a and b}
    titles = {(a or "", t): n for a, t, n in wb["曲名"].iter_rows(min_row=2, values_only=True) if t and n}
    wb.close()
    return artists, albums, titles


def _mark(cell, value, original):
    cell.value = value
    cell.fill = BLUE if value == original else GREEN
    cell.font = Font(bold=True)


def run(excel_path, progress_callback=None, dict_folder=None):
    artist_core, album_map, title_map = load_translation_dicts(dict_folder)
    ov_artist, ov_album, ov_title = load_overrides()
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["IDTAG"]
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row))
    total = len(rows)

    n_title = n_artist = n_album = 0
    n_kept_title = n_kept_album = 0
    n_unresolved = 0
    n_artist_unresolved = 0

    for i, row in enumerate(rows, 1):
        title_cell, artist_cell, album_cell = row[2], row[3], row[4]
        title, artist, album = title_cell.value, artist_cell.value, album_cell.value

        if (artist or "", title) in ov_title:
            _mark(title_cell, ov_title[(artist or "", title)], title)
            n_title += 1
            title = None
        if artist in ov_artist:
            _mark(artist_cell, ov_artist[artist], artist)
            n_artist += 1
            artist = None
        if album in ov_album:
            _mark(album_cell, ov_album[album], album)
            n_album += 1
            album = None

        new_artist, a_changed = convert_artist(artist, artist_core)
        new_album, al_changed = convert_album(album, album_map)
        new_title, t_changed = convert_title(title, artist, title_map)

        if new_artist and JP_RE.search(str(new_artist)):
            # still has unmapped Japanese after conversion (even if part of a
            # compound name was resolved) -- flag instead of passing silently
            if a_changed:
                artist_cell.value = new_artist
            artist_cell.font = Font(bold=True, color="FFFF0000")
            n_artist_unresolved += 1
        elif a_changed:
            new_artist = title_case(new_artist)
            artist_cell.value = new_artist
            artist_cell.fill = GREEN
            artist_cell.font = Font(bold=True)
            n_artist += 1
        if al_changed:
            new_album = title_case(new_album)
            album_cell.value = new_album
            album_cell.fill = GREEN
            album_cell.font = Font(bold=True)
            n_album += 1
        elif album and JP_RE.search(str(album)):
            album_cell.fill = BLUE
            album_cell.font = Font(bold=True)
            n_kept_album += 1

        if t_changed:
            new_title = title_case(new_title)
            title_cell.value = new_title
            title_cell.fill = GREEN
            title_cell.font = Font(bold=True)
            n_title += 1
        elif title and JP_RE.search(str(title)):
            if artist == "SHIHO":
                title_cell.fill = BLUE
                title_cell.font = Font(bold=True)
                n_kept_title += 1
            else:
                title_cell.font = Font(bold=True, color="FFFF0000")
                n_unresolved += 1

        if progress_callback and i % 100 == 0:
            progress_callback(i, total, f"英語変換中 {i}/{total}")

    n_formatting_fixed = normalize_english_fields(rows, progress_callback)

    wb.save(excel_path)
    if progress_callback:
        progress_callback(total, total, "完了")

    return {
        "converted": n_title + n_artist + n_album,
        "converted_title": n_title, "converted_artist": n_artist, "converted_album": n_album,
        "kept_japanese": n_kept_title + n_kept_album,
        "unresolved": n_unresolved + n_artist_unresolved,
        "unresolved_title": n_unresolved, "unresolved_artist": n_artist_unresolved,
        "formatting_fixed": n_formatting_fixed,
    }
