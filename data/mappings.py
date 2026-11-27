# -*- coding: utf-8 -*-
import os
import openpyxl

_BUNDLED_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "translation_dictionary.xlsx")
_FILENAME = "translation_dictionary.xlsx"


def _read_sheet(ws):
    d = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0] or len(row) < 2 or not row[1]:
            continue
        d[str(row[0])] = str(row[1])
    return d


def load_translation_dicts(folder=None):
    """Returns (artist_core, album_map, title_map) read from
    translation_dictionary.xlsx's "Artist"/"Album"/"Title" sheets (each a
    plain 日本語 -> English 2-column table). Reads
    `<folder>/translation_dictionary.xlsx` when present, else the bundled
    default shipped alongside this file."""
    path = os.path.join(folder, _FILENAME) if folder else None
    if not path or not os.path.exists(path):
        path = _BUNDLED_PATH

    wb = openpyxl.load_workbook(path, data_only=True)
    return _read_sheet(wb["Artist"]), _read_sheet(wb["Album"]), _read_sheet(wb["Title"])
