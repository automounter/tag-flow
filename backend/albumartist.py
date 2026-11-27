# -*- coding: utf-8 -*-
"""Album-artist decision per album (folder), not per track.

Categories:
  single   - every track has the same single artist (auto)
  group    - every track has the same name, and it is a known group such as
             "Earth, Wind & Fire" (auto, never split)
  multi    - the album's artist is several artists listed together
             ("Paul McCartney & Stevie Wonder"): choose one of them or the whole name
  major    - one artist appears on more than half of the tracks (auto)
  various  - nobody appears on more than half: "Various Artists" or pick a participant
"""
import os
import re
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import PatternFill, Font

YELLOW = PatternFill(fgColor="FFF2CC", fill_type="solid")
DARK_YELLOW = PatternFill(fgColor="FFD700", fill_type="solid")

VARIOUS = "Various Artists"
NEEDS_REVIEW = ("multi", "various")
CATEGORY_LABEL = {
    "single": "単独",
    "group": "グループ名",
    "major": "主要アーティスト",
    "multi": "複数並記",
    "various": "過半数なし",
}

GROUPS_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "data", "artist_groups.txt")

# "and" and " x " are not separators: "KC and The Sunshine Band", "Chloe x Halle"
SPLIT_RE = re.compile(r'\s*(?:,|&|\bfeat\.|\bfeat\b|\bft\.|\bfeaturing\b|\bwith\b|\bvs\.?)\s*', re.I)
# "Gladys Knight & The Pips", "Quincy Jones & His Orchestra" are one act
BACKING_RE = re.compile(r'^(?:the|his|her|their|friends)\b', re.I)
FEAT_RE = re.compile(r'\b(?:feat\.?|ft\.|featuring|with)\s', re.I)
DISC_DIR_RE = re.compile(r'^(?:disc|disk|cd)\s*\d+$', re.I)


def load_groups():
    try:
        with open(GROUPS_FILE, encoding="utf-8") as f:
            return [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
    except FileNotFoundError:
        return []


def add_group(name):
    groups = load_groups()
    if name.lower() in (g.lower() for g in groups):
        return
    with open(GROUPS_FILE, "a", encoding="utf-8") as f:
        f.write(name.strip() + "\n")


def split_artists(name, groups, standalone=()):
    """'The Emotions & Earth, Wind & Fire' -> ['The Emotions', 'Earth, Wind & Fire']
    standalone: names seen on their own elsewhere ("The Orb"), never merged as a backing band"""
    if not name:
        return []
    protected = {}
    text = name
    for i, g in enumerate(sorted(groups, key=len, reverse=True)):
        pattern = re.compile(r'(?<![\w])' + re.escape(g) + r'(?![\w])', re.I)
        if pattern.search(text):
            token = f"\x00{i}\x00"
            protected[token] = g
            text = pattern.sub(token, text)

    parts = []
    for p in SPLIT_RE.split(text):
        p = p.strip()
        if not p:
            continue
        if parts and BACKING_RE.match(p) and p not in standalone:
            parts[-1] = f"{parts[-1]} & {p}"
            continue
        parts.append(p)

    out = []
    for p in parts:
        for token, g in protected.items():
            p = p.replace(token, g)
        if p.lower() not in (o.lower() for o in out):
            out.append(p)
    return out


def album_key(folder):
    if folder and DISC_DIR_RE.match(os.path.basename(folder)):
        return os.path.dirname(folder)
    return folder or ""


def _year(v):
    if hasattr(v, "year"):
        return v.year
    return str(v)[:4] if v else ""


def _has_collision(tracks):
    seen = Counter((t["disc"], t["no"]) for t in tracks if t["no"])
    return any(c > 1 for c in seen.values())


def _split_mixed(tracks, groups, standalone):
    """A folder sometimes holds two different albums with the same name
    ("Greatest Hits" by Björk and by Quincy Jones). Track numbers then repeat;
    split by year first, then by the first-credited artist."""
    if not _has_collision(tracks):
        return [(tracks, False)]
    by_year = defaultdict(list)
    for t in tracks:
        by_year[t["year"]].append(t)
    out = []
    for sub in by_year.values():
        if not _has_collision(sub):
            out.append((sub, True))
            continue
        by_artist = defaultdict(list)
        for t in sub:
            parts = split_artists(t["artist"], groups, standalone)
            by_artist[parts[0] if parts else ""].append(t)
        # guest tracks ("Patti Austin & James Ingram" on a Quincy Jones album) go back
        # into the first bigger cluster whose track numbers they don't clash with
        buckets = []
        for cluster in sorted(by_artist.values(), key=len, reverse=True):
            for b in buckets:
                if not _has_collision(b + cluster):
                    b.extend(cluster)
                    break
            else:
                buckets.append(list(cluster))
        out.extend((b, True) for b in buckets)
    return out


def _decide(artists, groups, standalone):
    n = len(artists)
    full_counts = Counter(artists)
    credit = Counter()
    for t in artists:
        for p in split_artists(t, groups, standalone):
            credit[p] += 1
    majority = [p for p, c in credit.most_common() if c * 2 > n]

    if len(full_counts) == 1:
        name = artists[0]
        parts = split_artists(name, groups, standalone)
        if len(parts) <= 1:
            category = "group" if SPLIT_RE.search(name) else "single"
            return category, [name], name, credit
        # "Incognito featuring Linda Lee Hopkins" -> the main act is the natural pick
        default = parts[0] if FEAT_RE.search(name) else name
        return "multi", parts + [name], default, credit
    if len(majority) == 1:
        return "major", [majority[0]], majority[0], credit
    if len(majority) >= 2:
        combined = full_counts.most_common(1)[0][0]
        if set(split_artists(combined, groups, standalone)) != set(majority):
            combined = " & ".join(majority)
        return "multi", majority + [combined], combined, credit
    return "various", [VARIOUS], VARIOUS, credit


def analyze(excel_path, baseline_path=None):
    groups = load_groups()
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    ws = wb["IDTAG"]
    folders = defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[4] is None:
            continue
        folders[(album_key(row[14]), row[4])].append({
            "artist": row[3] or "", "album": row[4], "no": row[5], "disc": row[7],
            "year": _year(row[9]), "current": row[12] or "", "file": row[16],
        })
    wb.close()

    standalone = {t["artist"] for ts in folders.values() for t in ts if not SPLIT_RE.search(t["artist"])}

    original = {}
    if baseline_path and os.path.exists(baseline_path):
        bwb = openpyxl.load_workbook(baseline_path, read_only=True)
        for row in bwb["IDTAG"].iter_rows(min_row=2, values_only=True):
            if row[16]:
                original[row[16]] = row[12]
        bwb.close()

    items = []
    for (folder, album), tracks in folders.items():
        for sub, mixed in _split_mixed(tracks, groups, standalone):
            category, options, default, credit = _decide([t["artist"] for t in sub], groups, standalone)
            files = [t["file"] for t in sub]
            orig = Counter(original.get(f) for f in files if original.get(f))
            items.append({
                "key": f"{folder}|{album}|{sub[0]['year']}|{sub[0]['artist']}" if mixed else f"{folder}|{album}",
                "album": album,
                "folder": folder,
                "mixed": mixed,
                "tracks": len(sub),
                "category": category,
                "options": options,
                "default": default,
                "participants": credit.most_common(),
                "current": Counter(t["current"] for t in sub).most_common(1)[0][0],
                "original": orig.most_common(1)[0][0] if orig else "",
                "files": files,
            })
    items.sort(key=lambda it: (not needs_review(it), it["album"].lower()))
    return items


def needs_review(item):
    # a split-off fragment of one or two tracks is often a guest track that landed apart
    return item["category"] in NEEDS_REVIEW or (item["mixed"] and item["tracks"] <= 2)


def apply(excel_path, choices, progress_callback=None):
    """choices: {mp3 file path: album artist}"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["IDTAG"]
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row))
    n_set = n_diff = n_various = 0
    for i, row in enumerate(rows, 1):
        if row[4].value is None:
            continue
        value = choices.get(row[16].value)
        if not value:
            continue
        cell = row[12]
        cell.value = value
        n_set += 1
        if value == VARIOUS:
            cell.fill = DARK_YELLOW
            cell.font = Font(bold=True)
            n_various += 1
        elif value != row[3].value:
            cell.fill = YELLOW
            cell.font = Font(bold=True)
            n_diff += 1
        else:
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font()
        if progress_callback and i % 200 == 0:
            progress_callback(i, len(rows), f"アルバムアーティスト書き込み中 {i}/{len(rows)}")
    wb.save(excel_path)
    if progress_callback:
        progress_callback(len(rows), len(rows), "完了")
    return {"set": n_set, "differ": n_diff, "various_artists": n_various}
