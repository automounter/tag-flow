# -*- coding: utf-8 -*-
"""Optional online re-verification of already-English artist/album names,
using Apple's public iTunes Search API (no key required)."""
import json
import re
import time
import urllib.parse
import urllib.request
from collections import defaultdict
from openpyxl.styles import PatternFill, Font

JP_RE = re.compile(r'[぀-ゟ゠-ヿ一-鿿]')
VIOLET = PatternFill(fgColor="D9D2F0", fill_type="solid")

API_URL = "https://itunes.apple.com/search"


def _fetch(term, entity, limit=5):
    url = API_URL + "?" + urllib.parse.urlencode({"term": term, "entity": entity, "limit": limit})
    req = urllib.request.Request(url, headers={"User-Agent": "TagFlow/1.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.load(resp)


def _same(a, b):
    # only spelling/casing/punctuation fixes are accepted; a different name is never a "correction"
    from backend.name_lookup import same
    return same(a, b)


def lookup_artist(name):
    try:
        data = _fetch(name, "musicArtist", limit=5)
    except Exception:
        return None
    return next((r.get("artistName") for r in data.get("results", []) if _same(r.get("artistName"), name)), None)


def lookup_album(artist, album):
    try:
        data = _fetch(f"{artist} {album}", "album", limit=10)
    except Exception:
        return None
    # the first search result used to be taken as-is, which renamed albums to unrelated
    # ones ("Doin' the Thing" -> an ELO best-of); now both names have to match
    for r in data.get("results", []):
        if _same(r.get("artistName", ""), artist) and _same(r.get("collectionName"), album):
            return r.get("collectionName")
    return None


def collect_english_names(rows):
    """Returns (artist_names, album_names) sorted lists of unique already-English values."""
    artists = set()
    albums = set()
    for row in rows:
        artist, album = row[3].value, row[4].value
        if artist and not JP_RE.search(str(artist)):
            for part in re.split(r'\s*[,&]\s*', str(artist)):
                if part:
                    artists.add(part)
        if album and not JP_RE.search(str(album)):
            albums.add(str(album))
    return sorted(artists), sorted(albums)


def verify_selected(rows, selected_artists, selected_albums, progress_callback=None):
    """Look up selected artist/album names on iTunes and apply corrections in-place.
    Returns list of {type, before, after} for names that changed."""
    changes = []
    total = len(selected_artists) + len(selected_albums)
    done = 0

    artist_updates = {}
    for name in selected_artists:
        official = lookup_artist(name)
        done += 1
        if progress_callback:
            progress_callback(done, total, f"Apple Musicで照会中: {name}")
        time.sleep(0.15)  # be polite to the public API
        if official and official != name:
            artist_updates[name] = official
            changes.append({"type": "artist", "before": name, "after": official})

    # album lookup needs an artist context; use the album's most common artist in the sheet
    album_artist_map = defaultdict(lambda: defaultdict(int))
    for row in rows:
        artist, album = row[3].value, row[4].value
        if album in selected_albums and artist:
            album_artist_map[album][str(artist)] += 1

    album_updates = {}
    for album in selected_albums:
        artist_counts = album_artist_map.get(album, {})
        rep_artist = max(artist_counts, key=artist_counts.get) if artist_counts else ""
        official = lookup_album(rep_artist, album) if rep_artist else None
        done += 1
        if progress_callback:
            progress_callback(done, total, f"Apple Musicで照会中: {album}")
        time.sleep(0.15)
        if official and official != album:
            album_updates[album] = official
            changes.append({"type": "album", "before": album, "after": official})

    n_applied = 0
    for row in rows:
        artist_cell, album_cell = row[3], row[4]
        artist, album = artist_cell.value, album_cell.value
        if artist:
            parts = re.split(r'(\s*[,&]\s*)', str(artist))
            new_parts = [artist_updates.get(p, p) for p in parts]
            new_artist = "".join(new_parts)
            if new_artist != artist:
                artist_cell.value = new_artist
                artist_cell.fill = VIOLET
                artist_cell.font = Font(bold=True)
                n_applied += 1
        if album in album_updates:
            album_cell.value = album_updates[album]
            album_cell.fill = VIOLET
            album_cell.font = Font(bold=True)
            n_applied += 1

    return changes, n_applied
