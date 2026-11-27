# -*- coding: utf-8 -*-
"""STEP05: genre by album, Discogs style (Genre > Style).

The MP3 genre tag (TCON, the "ジャンル" column) gets the style name only
(e.g. "Boogie"); the parent genre and where the value came from go in two
extra columns. Decided 2026-11-07."""
import os
import re
import json
import urllib.parse
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import PatternFill, Font

from backend.albumartist import album_key, split_artists, load_groups
from backend import name_lookup as nl

STYLES_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "data", "genre_styles.xlsx")
UNKNOWN_PARENT = "（親ジャンル未定）"

GENRE_COL = 10            # ジャンル (0-indexed) -> TCON
PARENT_COL = 19           # T: ジャンル(Discogs)
SOURCE_COL = 20           # U: ジャンルの根拠
PARENT_HEADER = "ジャンル（Discogs）"
SOURCE_HEADER = "ジャンルの根拠"

SOURCE_LABEL = {"excel": "設定済み", "artist": "アーティスト既定", "discogs": "Discogs",
                "apple": "Apple Music", "spotify": "Spotify", "manual": "手動", "none": "未決定"}

MANUAL_FILL = PatternFill(fgColor="FCE4EC", fill_type="solid")
APPLE_FILL = PatternFill(fgColor="FFF2CC", fill_type="solid")



# ---------- dictionary (data/genre_styles.xlsx) ----------

def load_dictionary():
    wb = openpyxl.load_workbook(STYLES_FILE, read_only=True)
    hierarchy = defaultdict(list)  # genre -> [style]
    for g, s, _src in wb["スタイル"].iter_rows(min_row=2, values_only=True):
        if g and s and s not in hierarchy[g]:
            hierarchy[g].append(s)
    artist_default = {}
    for a, g, s in wb["アーティスト既定"].iter_rows(min_row=2, values_only=True):
        if a and g and s:
            artist_default[a.strip().lower()] = (g, s)
    apple = {}
    for a, g, s in wb["Apple対応"].iter_rows(min_row=2, values_only=True):
        if a:
            apple[a.strip().lower()] = (g, s)
    wb.close()
    for g in hierarchy:
        hierarchy[g].sort(key=str.lower)
    return {"hierarchy": dict(hierarchy), "artist_default": artist_default, "apple": apple}


def _append_row(sheet, values):
    wb = openpyxl.load_workbook(STYLES_FILE)
    wb[sheet].append(values)
    wb.save(STYLES_FILE)


def set_artist_default(artist, genre, style):
    wb = openpyxl.load_workbook(STYLES_FILE)
    ws = wb["アーティスト既定"]
    for row in ws.iter_rows(min_row=2):
        if row[0].value and row[0].value.strip().lower() == artist.strip().lower():
            row[1].value, row[2].value = genre, style
            break
    else:
        ws.append((artist, genre, style))
    wb.save(STYLES_FILE)


def add_custom_style(genre, style):
    _append_row("スタイル", (genre, style, "独自"))


# ---------- online lookup (Discogs・Apple Music・Spotify) ----------

SERVICE_LABEL = {"discogs": "Discogs", "apple": "Apple Music", "spotify": "Spotify"}
VARIOUS = "various artists"


def _lookup_cache_path(excel_path):
    base, _ = os.path.splitext(excel_path)
    return f"{base}_genre_lookup.json"


def load_lookup_cache(excel_path):
    try:
        with open(_lookup_cache_path(excel_path), encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _apple_genre(http, albumartist, album):
    various = albumartist.lower() == VARIOUS
    q = urllib.parse.urlencode({"term": album if various else f"{albumartist} {album}",
                                "entity": "album", "limit": 10, "country": "US"})
    for r in (http.get("apple", "https://itunes.apple.com/search?" + q) or {}).get("results", []):
        if nl.same(r.get("collectionName"), album) and (various or nl.same(r.get("artistName"), albumartist)):
            return r.get("primaryGenreName") or ""
    return ""


def _discogs_genre(http, albumartist, album):
    """Discogs search results already carry genre/style, so one request per album."""
    various = albumartist.lower() == VARIOUS
    params = {"release_title": nl.core(album), "per_page": 10}
    if not various:
        params["artist"] = albumartist
    d = http.get("discogs", "https://api.discogs.com/database/search?" + urllib.parse.urlencode(params)) or {}
    for x in d.get("results", []):
        title = x.get("title", "")
        album_part = title.split(" - ", 1)[1] if " - " in title else title
        if any(nl.same(p, album) for p in nl.discogs_parts(album_part)):
            return {"genres": x.get("genre") or [], "styles": x.get("style") or []}
    return {"genres": [], "styles": []}


def _spotify_genres(http, albumartist, cache):
    """Spotify has no album genres; the album artist's genres stand in."""
    if albumartist.lower() == VARIOUS:
        return []
    artists = cache.setdefault("_spotify_artists", {})
    if albumartist not in artists:
        q = urllib.parse.urlencode({"q": albumartist, "type": "artist", "limit": 5})
        items = ((http.get("spotify", "https://api.spotify.com/v1/search?" + q) or {}).get("artists") or {}).get("items", [])
        hit = next((x for x in items if nl.same(x.get("name"), albumartist)), None)
        artists[albumartist] = (hit or {}).get("genres") or []
    return artists[albumartist]


def fetch_genres(excel_path, items, sources, progress_callback=None):
    """Asks the chosen services about albums that have no answer from them yet.
    Results are kept next to the Excel, also when cancelled."""
    keys = nl.load_api_keys()
    sources = [s for s in nl.available_sources(keys) if s in sources]
    http = nl.Http(keys)
    cache = load_lookup_cache(excel_path)

    def save():
        with open(_lookup_cache_path(excel_path), "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)

    todo = [it for it in items if not set(sources) <= set(cache.get(it["key"], {}))]
    state = {"message": "", "i": 0}

    def report(message):
        try:
            if progress_callback:
                progress_callback(state["i"], len(todo), message)
        except BaseException:  # Cancel: keep what was fetched
            save()
            raise

    def on_wait(msg):
        report(f"{state['message']} ／ {msg}")
    http.on_wait = on_wait

    for i, it in enumerate(todo, 1):
        state["i"] = i - 1
        state["message"] = f"アルバムのジャンルを照会中 {i}/{len(todo)}：{it['album']}"
        report(state["message"])
        entry = cache.setdefault(it["key"], {})
        if "discogs" in sources and "discogs" not in entry:
            entry["discogs"] = _discogs_genre(http, it["albumartist"], it["album"])
        if "apple" in sources and "apple" not in entry:
            entry["apple"] = _apple_genre(http, it["albumartist"], it["album"])
        if "spotify" in sources and "spotify" not in entry:
            entry["spotify"] = _spotify_genres(http, it["albumartist"], cache)
        if i % 10 == 0:
            save()
    save()
    if progress_callback:
        progress_callback(len(todo), len(todo), "完了")
    return {"searched": len(todo), "sources": sources}


def _from_discogs(info, style_parents):
    """First Discogs style, under the Discogs genre it belongs to."""
    if not info or not info.get("styles"):
        return None, None
    style = info["styles"][0]
    parents = style_parents.get(style, [])
    genre = next((g for g in info.get("genres", []) if g in parents), None)
    return genre or (parents[0] if parents else (info.get("genres") or [UNKNOWN_PARENT])[0]), style


def _from_spotify(genres, style_parents):
    by_norm = {re.sub(r'[\s\-&/]+', '', s.lower()): s for s in style_parents}
    for g in genres or []:
        style = by_norm.get(re.sub(r'[\s\-&/]+', '', g.lower()))
        if style:
            return style_parents[style][0], style
    return None, None


# ---------- album list ----------

def build_items(excel_path):
    d = load_dictionary()
    hierarchy = d["hierarchy"]
    style_parents = defaultdict(list)
    for g, styles in hierarchy.items():
        for s in styles:
            style_parents[s].append(g)
    lookup = load_lookup_cache(excel_path)
    groups = load_groups()

    wb = openpyxl.load_workbook(excel_path, read_only=True)
    albums = {}
    for row in wb["IDTAG"].iter_rows(min_row=2, values_only=True):
        if row[4] is None:
            continue
        albumartist = row[12] or row[3] or ""
        key = f"{album_key(row[14])}|{row[4]}|{albumartist}"
        a = albums.setdefault(key, {"key": key, "album": row[4], "albumartist": albumartist,
                                    "folder": album_key(row[14]), "tracks": 0, "excel": Counter()})
        a["tracks"] += 1
        parent = row[PARENT_COL] if len(row) > PARENT_COL else None
        if row[GENRE_COL]:
            a["excel"][(parent, row[GENRE_COL])] += 1
    wb.close()

    items = []
    for a in albums.values():
        found = lookup.get(a["key"], {})
        apple_genre = found.get("apple")
        various = a["albumartist"].lower() == VARIOUS
        genre = style = None
        source = "none"
        if a["excel"]:
            parent, st = a["excel"].most_common(1)[0][0]
            if st in style_parents:  # an old 4-level master value is not a style; ignore it
                genre, style, source = (parent if parent in style_parents[st] else style_parents[st][0]), st, "excel"
        if not style and not various:  # a compilation's genre has nothing to do with "Various Artists"
            # "Dirty Projectors & Björk" falls back to the first listed artist that has a default
            for name in [a["albumartist"]] + split_artists(a["albumartist"], groups):
                if name.lower() in d["artist_default"]:
                    (genre, style), source = d["artist_default"][name.lower()], "artist"
                    break
        if not style:
            g, st = _from_discogs(found.get("discogs"), style_parents)
            if st:
                genre, style, source = g, st, "discogs"
        if not style and apple_genre:
            g, st = d["apple"].get(apple_genre.lower(), (None, None))
            if g:
                genre, style, source = g, st, ("apple" if st else "none")
        if not style:
            g, st = _from_spotify(found.get("spotify"), style_parents)
            if st:
                genre, style, source = g, st, "spotify"
        items.append({**{k: a[k] for k in ("key", "album", "albumartist", "folder", "tracks")},
                      "found": found, "various": various,
                      "genre": genre, "style": style, "source": source})
    items.sort(key=lambda it: (it["style"] is not None, it["albumartist"].lower(), it["album"].lower()))
    return items, hierarchy, d["artist_default"]


def apply(excel_path, choices, progress_callback=None):
    """choices: {album key: (genre, style, source)}"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["IDTAG"]
    ws.cell(row=1, column=PARENT_COL + 1, value=PARENT_HEADER)
    ws.cell(row=1, column=SOURCE_COL + 1, value=SOURCE_HEADER)
    ws.column_dimensions["T"].width = 22
    ws.column_dimensions["U"].width = 18
    total = ws.max_row - 1
    n_set = 0
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), 1):
        if row[4].value is None:
            continue
        albumartist = row[12].value or row[3].value or ""
        choice = choices.get(f"{album_key(row[14].value)}|{row[4].value}|{albumartist}")
        if not choice:
            continue
        genre, style, source = choice
        r = row[0].row
        # ws.cell(value=None) leaves the old value in place, so assign explicitly
        cell = ws.cell(row=r, column=GENRE_COL + 1)
        cell.value = style
        ws.cell(row=r, column=PARENT_COL + 1).value = genre
        ws.cell(row=r, column=SOURCE_COL + 1).value = SOURCE_LABEL.get(source, source)
        cell.font = Font(bold=True)
        cell.fill = MANUAL_FILL if source == "manual" else (APPLE_FILL if source == "apple" else PatternFill(fill_type=None))
        n_set += 1
        if progress_callback and i % 200 == 0:
            progress_callback(i, total, f"ジャンル書き込み中 {i}/{total}")
    wb.save(excel_path)
    if progress_callback:
        progress_callback(total, total, "完了")
    return {"set": n_set, "unresolved": 0}
