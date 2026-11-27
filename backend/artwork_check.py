# -*- coding: utf-8 -*-
"""STEP06: compare embedded album artwork against Apple Music's official
cover art (via the free iTunes Search API) and flag albums that are
missing artwork or whose artwork looks different."""
import io
import os
import re
import json
import urllib.parse
import urllib.request
import openpyxl
from openpyxl.styles import PatternFill, Font
from mutagen.id3 import ID3
from PIL import Image
from backend import name_lookup as nl

from config import CONFIG_DIR

CACHE_DIR = os.path.join(CONFIG_DIR, "artwork_cache")
MISMATCH_THRESHOLD = 12  # hamming distance out of 64 bits
ARTWORK_DELETE = "__delete__"  # sentinel value: strip embedded artwork instead of replacing it
MAX_ARTWORK_SIZE = 1000  # px, longest side an embedded cover is capped to
ARTWORK_JPEG_QUALITY = 90

ARTWORK_COL = 18  # "アートワーク" column (1-indexed, matches extract.py's HEADERS)
GREEN = PatternFill(fgColor="C6EFCE", fill_type="solid")
BLUE = PatternFill(fgColor="DDEBF7", fill_type="solid")
WARN = PatternFill(fgColor="F8CBAD", fill_type="solid")


def _request_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "TagFlow/1.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.load(resp)


def _request_bytes(url):
    req = urllib.request.Request(url, headers={"User-Agent": "TagFlow/1.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        return resp.read()


def _phash(img_bytes, size=8):
    img = Image.open(io.BytesIO(img_bytes)).convert("L").resize((size, size), Image.LANCZOS)
    pixels = list(img.getdata())
    avg = sum(pixels) / len(pixels)
    return "".join("1" if p > avg else "0" for p in pixels)


def _hamming(a, b):
    return sum(c1 != c2 for c1, c2 in zip(a, b))


def _local_artwork(mp3_path):
    if not mp3_path or not os.path.exists(mp3_path):
        return None
    try:
        tags = ID3(mp3_path)
        apics = tags.getall("APIC")
        if apics:
            return apics[0].data
    except Exception:
        pass
    return None


ARTWORK_SOURCES = ("apple", "spotify", "discogs")
SOURCE_NAMES = {"apple": "Apple Music", "spotify": "Spotify", "discogs": "Discogs"}


def _apple_candidate(http, artist, album):
    q = urllib.parse.urlencode({"term": f"{artist} {album}", "entity": "album", "limit": 10})
    results = (http.get("apple", "https://itunes.apple.com/search?" + q) or {}).get("results", [])
    best = next((r for r in results if nl.same(r.get("collectionName"), album)
                 and nl.same(r.get("artistName"), artist)), None)
    best = best or next((r for r in results if nl.same(r.get("collectionName"), album)), None)
    if not best or not best.get("artworkUrl100"):
        return None
    return {"source": "apple", "size": MAX_ARTWORK_SIZE,
            "url": re.sub(r"\d+x\d+bb\.jpg$", f"{MAX_ARTWORK_SIZE}x{MAX_ARTWORK_SIZE}bb.jpg", best["artworkUrl100"])}


def _spotify_candidate(http, artist, album):
    q = urllib.parse.urlencode({"q": f"{album} {artist}", "type": "album", "limit": 10})
    items = ((http.get("spotify", "https://api.spotify.com/v1/search?" + q) or {}).get("albums") or {}).get("items", [])
    best = next((x for x in items if nl.same(x.get("name"), album)
                 and any(nl.same(a.get("name"), artist) for a in x.get("artists", []))), None)
    if not best or not best.get("images"):
        return None
    img = max(best["images"], key=lambda i: i.get("width") or 0)  # at most 640px on Spotify
    return {"source": "spotify", "size": img.get("width") or 0, "url": img["url"]}


def _discogs_candidate(http, artist, album):
    params = {"release_title": nl.core(album), "per_page": 10, "type": "master"}
    if artist.lower() != "various artists":
        params["artist"] = artist
    d = http.get("discogs", "https://api.discogs.com/database/search?" + urllib.parse.urlencode(params)) or {}
    for x in d.get("results", []):
        title = x.get("title", "")
        album_part = title.split(" - ", 1)[1] if " - " in title else title
        if any(nl.same(p, album) for p in nl.discogs_parts(album_part)) and x.get("cover_image"):
            # the search result's cover is the release's primary image; its size isn't given
            master = http.get("discogs", x["resource_url"]) or {}
            img = next((i for i in master.get("images", []) if i.get("type") == "primary"), None) \
                or (master.get("images") or [None])[0]
            if img:
                return {"source": "discogs", "size": min(img.get("width") or 0, img.get("height") or 0),
                        "url": img.get("uri") or x["cover_image"]}
            return {"source": "discogs", "size": 0, "url": x["cover_image"]}
    return None


def scan(excel_path, progress_callback=None, sources=ARTWORK_SOURCES):
    """Compares each album's embedded artwork with the first image found among
    the chosen services (Apple Music → Spotify → Discogs). Every image found is
    kept as a candidate so the screen can offer the others too."""
    keys = nl.load_api_keys()
    sources = [s for s in ARTWORK_SOURCES if s in sources and s in nl.available_sources(keys)]
    http = nl.Http(keys)
    finders = {"apple": _apple_candidate, "spotify": _spotify_candidate, "discogs": _discogs_candidate}

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb["IDTAG"]

    folder_info = {}
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        filepath = r[16].value
        folder = r[14].value
        if not filepath or not folder or folder in folder_info:
            continue
        folder_info[folder] = {
            "album": r[4].value,
            "albumartist": r[12].value or r[3].value,
            "sample_file": filepath,
        }

    items = [(f, info) for f, info in folder_info.items() if info["album"] and info["albumartist"]]
    total = len(items)
    results = []
    state = {"message": "", "i": 0}

    def on_wait(msg):
        if progress_callback:
            progress_callback(state["i"], total, f"{state['message']} ／ {msg}")
    http.on_wait = on_wait

    for i, (folder, info) in enumerate(items, 1):
        album, artist = info["album"], info["albumartist"]
        state["i"], state["message"] = i - 1, f"アートワーク確認中 {i}/{total} ({album})"
        if progress_callback:
            progress_callback(i - 1, total, state["message"])
        local_bytes = _local_artwork(info["sample_file"])
        candidates = []
        for src in sources:
            try:
                c = finders[src](http, artist, album)
            except Exception:
                c = None
            if c:
                candidates.append(c)
        artwork_url = candidates[0]["url"] if candidates else None

        status = None
        if local_bytes is None:
            status = "missing"
        elif artwork_url:
            try:
                remote_bytes = _request_bytes(artwork_url)
                dist = _hamming(_phash(local_bytes), _phash(remote_bytes))
                if dist > MISMATCH_THRESHOLD:
                    status = "mismatch"
            except Exception:
                pass

        if status:
            results.append({
                "folder": folder, "album": album, "albumartist": artist,
                "status": status, "artwork_url": artwork_url, "local_bytes": local_bytes,
                "candidates": candidates,
            })

    if progress_callback:
        progress_callback(total, total, "完了")
    return results


def _normalize_image(data):
    """Validates the bytes are a JPEG or PNG image, caps them to
    MAX_ARTWORK_SIZE per side, and re-encodes as JPEG so every embedded
    cover is bounded in both dimensions and file size regardless of what
    the user supplied. Returns (jpeg_bytes, ".jpg")."""
    img = Image.open(io.BytesIO(data))
    if img.format not in ("PNG", "JPEG", "MPO"):
        raise ValueError(f"対応していない画像形式です ({img.format})。JPEGまたはPNGを指定してください。")

    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        img = background
    else:
        img = img.convert("RGB")

    if img.width > MAX_ARTWORK_SIZE or img.height > MAX_ARTWORK_SIZE:
        img.thumbnail((MAX_ARTWORK_SIZE, MAX_ARTWORK_SIZE), Image.LANCZOS)

    out = io.BytesIO()
    img.save(out, format="JPEG", quality=ARTWORK_JPEG_QUALITY, optimize=True)
    return out.getvalue(), ".jpg"


def load_local_image(file_path):
    """Reads a local image file, validates and normalizes it, returns (bytes, ext)."""
    with open(file_path, "rb") as f:
        data = f.read()
    return _normalize_image(data)


def fetch_url_image(url):
    """Downloads an image from a URL, validates and normalizes it, returns (bytes, ext)."""
    data = _request_bytes(url)
    return _normalize_image(data)


def apply_choices(choices, progress_callback=None):
    """Applies one artwork decision per flagged album.

    choices: list of {folder, album, source, value} where source is
    "local" (value = local file path), "url" or "official" (value = image
    URL; "official" is the recommended Apple Music image picked via the
    scan result rather than typed in by hand), or "delete" (strip the
    embedded artwork; value unused). Items the caller left undecided never
    reach here. Returns
    ({folder_path: cached_image_path_or_ARTWORK_DELETE}, [(album, error), ...])."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    overrides = {}
    errors = []
    total = len(choices)
    for i, c in enumerate(choices, 1):
        try:
            if c["source"] == "delete":
                overrides[c["folder"]] = ARTWORK_DELETE
            elif c["source"] == "local":
                data, ext = load_local_image(c["value"])
                safe_name = re.sub(r'[^\w\-]', '_', c["album"])[:60]
                cache_path = os.path.join(CACHE_DIR, f"{safe_name}_{abs(hash(c['folder']))}{ext}")
                with open(cache_path, "wb") as f:
                    f.write(data)
                overrides[c["folder"]] = cache_path
            elif c["source"] in ("url", "official"):
                data, ext = fetch_url_image(c["value"])
                safe_name = re.sub(r'[^\w\-]', '_', c["album"])[:60]
                cache_path = os.path.join(CACHE_DIR, f"{safe_name}_{abs(hash(c['folder']))}{ext}")
                with open(cache_path, "wb") as f:
                    f.write(data)
                overrides[c["folder"]] = cache_path
            else:
                continue
        except Exception as e:
            errors.append((c["album"], str(e)))
        if progress_callback:
            progress_callback(i, total, f"画像を保存中 {i}/{total}")
    return overrides, errors


def write_excel_results(excel_path, applied_choices):
    """Records each applied artwork decision into the IDTAG sheet's
    "アートワーク" column (18th, added alongside STEP08's other columns),
    on every row belonging to that album's folder -- so the decision is
    visible in the Excel data before STEP08 ever writes it into the MP3s.
    applied_choices: choices (see apply_choices) that succeeded, i.e. whose
    folder ended up in apply_choices' returned overrides dict."""
    if not applied_choices:
        return

    by_folder = {c["folder"]: c for c in applied_choices}

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["IDTAG"]
    if ws.cell(row=1, column=ARTWORK_COL).value != "アートワーク":
        ws.cell(row=1, column=ARTWORK_COL).value = "アートワーク"

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        folder = row[14].value
        choice = by_folder.get(folder)
        if choice is None:
            continue
        cell = row[ARTWORK_COL - 1]
        if choice["source"] == "local":
            cell.value = f"変更(フォルダ): {os.path.basename(choice['value'])}"
            cell.fill = GREEN
        elif choice["source"] == "url":
            cell.value = "変更(URL)"
            cell.fill = BLUE
        elif choice["source"] == "official":
            cell.value = "変更(推奨画像)"
            cell.fill = BLUE
        elif choice["source"] == "delete":
            cell.value = "削除"
            cell.fill = WARN
        else:
            continue
        cell.font = Font(bold=True)

    wb.save(excel_path)
