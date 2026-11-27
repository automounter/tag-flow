# -*- coding: utf-8 -*-
import os
import re
import openpyxl
from openpyxl.styles import Font

JP_RE = re.compile(r'[぀-ゟ゠-ヿ一-鿿]')
ALIAS_FALLBACK_FILENAME = "genre_alias_fallback.xlsx"

_BUNDLED_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "genre_dictionary.xlsx")
_FILENAME = "genre_dictionary.xlsx"


def load_genre_dict(folder=None):
    """Returns (priority_order, core_genre, default_genre) read from
    genre_dictionary.xlsx. Row order in the "Genre" sheet is the match
    priority (top row checked first, as a substring of the album artist);
    the row whose albumartist is "*" supplies the fallback. Reads
    `<folder>/genre_dictionary.xlsx` when present, else the bundled default
    shipped alongside this file."""
    path = os.path.join(folder, _FILENAME) if folder else None
    if not path or not os.path.exists(path):
        path = _BUNDLED_PATH

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Genre"]

    priority_order = []
    core_genre = {}
    default_genre = ("Dance", "Eurodance", "Eurodance", "Eurodance")
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        albumartist, cat, genre, sub, style = (row + (None,) * 5)[:5]
        path_tuple = (cat or "", genre or "", sub or "", style or "")
        if albumartist == "*":
            default_genre = path_tuple
            continue
        priority_order.append(str(albumartist))
        core_genre[str(albumartist)] = path_tuple

    return priority_order, core_genre, default_genre


def genre_path(cat, genre, sub, style):
    parts = [cat, genre]
    if sub and sub != genre:
        parts.append(sub)
    if style and style != sub and style != genre:
        parts.append(style)
    return " > ".join(parts)


def decide_genre(albumartist, priority_order, core_genre, default_genre,
                  embedded_genre=None, alias_fallback=None):
    """Returns a genre path string, or None when the artist is an unmapped
    Japanese name with no usable fallback -- in that case we refuse to guess
    rather than silently fall back to a default genre.

    When the artist isn't in the curated priority list, `alias_fallback`
    (see load_alias_fallback) is checked next using the album's own embedded
    genre tag (`embedded_genre`) before giving up -- this can resolve
    albums that would otherwise be left unmapped."""
    if not albumartist:
        return genre_path(*default_genre)
    for name in priority_order:
        if name in albumartist:
            return genre_path(*core_genre[name])
    if albumartist == "Various Artists":
        return genre_path(*default_genre)
    if alias_fallback and embedded_genre:
        key = str(embedded_genre).strip().lower()
        if key in alias_fallback:
            return alias_fallback[key]
    if JP_RE.search(albumartist):
        return None
    return genre_path(*default_genre)


def load_genre_hierarchy(genre_rule_folder=None):
    """Returns a nested dict describing every Category > Genre > Subgenre >
    Style combination available for manual selection in the STEP05 review
    screen: {category: {genre: {subgenre: [style, ...]}}}. Reads the
    "Genre Master" sheet of <genre_rule_folder>/genre_master_integrated.xlsx
    (the same reference dataset generate_alias_fallback reads from its
    "Aliases" sheet), and always also folds in genre_dictionary.xlsx's own
    curated (category, genre, subgenre, style) combinations (see
    load_genre_dict) -- so a value already decided from the curated
    dictionary is guaranteed to have a path back to its four components
    even if it's missing from the reference dataset, and manual selection
    still works with just the built-in dictionary when the reference
    dataset isn't configured at all."""
    hierarchy = {}

    def _add(cat, genre, sub, style):
        if not cat or not genre:
            return
        sub = sub or genre
        style = style or sub
        styles = hierarchy.setdefault(cat, {}).setdefault(genre, {}).setdefault(sub, [])
        if style not in styles:
            styles.append(style)

    master_path = os.path.join(genre_rule_folder, "genre_master_integrated.xlsx") if genre_rule_folder else None
    if master_path and os.path.exists(master_path):
        wb = openpyxl.load_workbook(master_path, data_only=True)
        if "Genre Master" in wb.sheetnames:
            ws = wb["Genre Master"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[1]:
                    continue
                _add(row[1], row[4], row[7], row[10])

    _, core_genre, default_genre = load_genre_dict(genre_rule_folder)
    for tup in list(core_genre.values()) + [default_genre]:
        _add(*tup)

    return hierarchy


def genre_reverse_lookup(hierarchy):
    """Inverts load_genre_hierarchy()'s tree into
    {display_string: (category, genre, subgenre, style)}, using the same
    abbreviation rule as genre_path() so a value currently stored in the
    "ジャンル" column (from the automatic classifier or a previous manual
    review) can be mapped back to its four components for the STEP05
    review screen's initial selection, without fragile string parsing."""
    reverse = {}
    for cat, genres in hierarchy.items():
        for genre, subs in genres.items():
            for sub, styles in subs.items():
                for style in styles:
                    display = genre_path(cat, genre, sub, style)
                    reverse.setdefault(display, (cat, genre, sub, style))
    return reverse


def generate_alias_fallback(genre_rule_folder, progress_callback=None):
    """Reads genre_master_integrated.xlsx's "Aliases" sheet from
    genre_rule_folder (the reference genre-string normalization dataset)
    and writes genre_alias_fallback.xlsx into the same folder: a
    normalized-genre-tag -> Category/Genre/Subgenre/Style table, used as a
    secondary fallback (see decide_genre) when an album's artist isn't in
    the curated priority dictionary. Returns {"entries": n, "output_path": p}."""
    source_path = os.path.join(genre_rule_folder, "genre_master_integrated.xlsx")
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"genre_master_integrated.xlsx が見つかりません: {source_path}")

    src_wb = openpyxl.load_workbook(source_path, data_only=True)
    ws = src_wb["Aliases"]
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    idx = {name: i for i, name in enumerate(header_row) if name}

    required = ("normalized_alias", "category_en", "genre_en")
    missing = [c for c in required if c not in idx]
    if missing:
        raise ValueError(f"Aliasesシートに必要な列がありません: {', '.join(missing)}")

    entries = {}
    all_rows = list(ws.iter_rows(min_row=2, values_only=True))
    total = len(all_rows)
    for i, row in enumerate(all_rows, 1):
        normalized_alias = row[idx["normalized_alias"]] if idx["normalized_alias"] < len(row) else None
        category = row[idx["category_en"]] if idx["category_en"] < len(row) else None
        genre = row[idx["genre_en"]] if idx["genre_en"] < len(row) else None
        subgenre = row[idx["subgenre_en"]] if "subgenre_en" in idx and idx["subgenre_en"] < len(row) else None
        style = row[idx["style_en"]] if "style_en" in idx and idx["style_en"] < len(row) else None
        if normalized_alias and category and genre:
            key = str(normalized_alias).strip().lower()
            entries.setdefault(key, (category, genre, subgenre or "", style or ""))
        if progress_callback and i % 200 == 0:
            progress_callback(i, total, f"補助ジャンル辞書を生成中 {i}/{total}")

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "AliasFallback"
    out_ws.append(["normalized_alias", "category", "genre", "subgenre", "style"])
    for cell in out_ws[1]:
        cell.font = Font(bold=True)
    for key in sorted(entries):
        cat, genre, sub, style = entries[key]
        out_ws.append([key, cat, genre, sub, style])
    for col, width in zip("ABCDE", [26, 18, 22, 22, 22]):
        out_ws.column_dimensions[col].width = width

    output_path = os.path.join(genre_rule_folder, ALIAS_FALLBACK_FILENAME)
    out_wb.save(output_path)
    if progress_callback:
        progress_callback(total, total, "完了")
    return {"entries": len(entries), "output_path": output_path}


def load_alias_fallback(folder):
    """Returns {normalized_alias: genre_path_string} from
    genre_alias_fallback.xlsx in `folder`, or {} if that file hasn't been
    generated yet (see generate_alias_fallback)."""
    if not folder:
        return {}
    path = os.path.join(folder, ALIAS_FALLBACK_FILENAME)
    if not os.path.exists(path):
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["AliasFallback"]
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        key, cat, genre, sub, style = (row + (None,) * 5)[:5]
        result[str(key)] = genre_path(cat or "", genre or "", sub or "", style or "")
    return result
