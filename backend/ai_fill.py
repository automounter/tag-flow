# -*- coding: utf-8 -*-
"""After the name lookup (STEP 03): asks an AI on this Mac (`claude -p` or
`codex exec`) for the original titles of names that are still red, then checks
each answer against the music services. Confirmed answers turn green; the rest
are written in orange as the AI's suggestion, to be checked by eye."""
import json
import os
import re
import shutil
import subprocess
import tempfile
import urllib.parse

import openpyxl
from openpyxl.styles import Font

from backend import name_lookup as nl

COLS = {nl.COL_TITLE: "トラック名", nl.COL_ARTIST: "アーティスト名", nl.COL_ALBUM: "アルバム名", nl.COL_AA: "アルバムアーティスト"}
KIND = {nl.COL_TITLE: "曲名", nl.COL_ARTIST: "アーティスト名", nl.COL_ALBUM: "アルバム名", nl.COL_AA: "アーティスト名"}
RED_RGB = "FF9999"
BATCH = 60
CLAUDE_TIMEOUT = 600

PROMPT = """あなたは音楽の資料係です。日本の MP3 のタグに日本語で入っている名前を、正式な表記に直します。
決まり:
- 洋楽(海外のアーティスト)の曲名・アルバム名・アーティスト名は、日本で付けた邦題やカタカナ表記を使わず、元の題名・表記(英語など)にする。例:「素顔のままで」(Billy Joel) → "Just The Way You Are"
- 括弧の補足(Live・Instrumental・Remix など)は元の題名の書き方に合わせて残す
- 邦楽(日本のアーティスト)は、本人・盤の正式な表記のまま。カタカナが正式ならカタカナのまま返す。例:「いとしのエリー」→「いとしのエリー」
- 日本独自の盤・編集盤などで原題が無くても、名前が英語をカタカナにしただけのもの(例:「ライヴ・コレクション」→ "Live Collection"、「グレイテスト・ヒッツ」→ "Greatest Hits")は、その英語の綴りを答えてよい
- それ以外で確信が持てないものは answer を null にする。それらしい題名を作らないこと
入力は JSON の配列です。各要素の id ごとに、次の形の JSON の配列だけを返してください(説明の文章は書かない):
[{"id": 1, "answer": "元の表記 または null", "western": true または false}]

入力:
"""


ENGINES = {"claude": "Claude", "codex": "Codex"}


def _cli(name):
    return shutil.which(name) or os.path.expanduser(f"~/.local/bin/{name}")


def _run_engine(engine, prompt):
    """Returns the AI's reply text. Both CLIs are the ones already signed in on this Mac."""
    if engine == "codex":
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "answer.txt")
            # stdin must be closed or codex waits for more input; read-only: it only has to answer
            proc = subprocess.run([_cli("codex"), "exec", "--skip-git-repo-check", "--ephemeral", "-s", "read-only",
                                   "-C", tmp, "-o", out, "--", prompt],
                                  capture_output=True, text=True, timeout=CLAUDE_TIMEOUT, stdin=subprocess.DEVNULL)
            if proc.returncode != 0 or not os.path.exists(out):
                raise RuntimeError(f"codex の実行に失敗しました: {(proc.stderr or proc.stdout).strip()[-300:]}")
            with open(out, encoding="utf-8") as f:
                return f.read()
    proc = subprocess.run([_cli("claude"), "-p", prompt, "--output-format", "json"],
                          capture_output=True, text=True, timeout=CLAUDE_TIMEOUT, stdin=subprocess.DEVNULL)
    if proc.returncode != 0:
        raise RuntimeError(f"claude の実行に失敗しました: {proc.stderr.strip()[:300]}")
    return json.loads(proc.stdout).get("result", "")


def ask_ai(items, engine="claude"):
    """items: [{"id", "kind", "text", "artist", "album"}] -> {id: (answer or None, western)}"""
    text = _run_engine(engine, PROMPT + json.dumps(items, ensure_ascii=False))
    m = re.search(r"\[.*\]", text, re.S)
    if not m:
        raise RuntimeError(f"{ENGINES[engine]} の答えを読み取れませんでした")
    out = {}
    for x in json.loads(m.group(0)):
        ans = x.get("answer")
        out[x.get("id")] = ((ans.strip() if isinstance(ans, str) and ans.strip() else None), bool(x.get("western")))
    return out


def _found(http, sources, kind, answer, artist):
    """Returns the service that lists `answer` (for this artist), or None."""
    if "spotify" in sources:
        t = {"曲名": "track", "アルバム名": "album", "アーティスト名": "artist"}[kind]
        q = answer if kind == "アーティスト名" else f"{answer} {artist}"
        d = http.get("spotify", "https://api.spotify.com/v1/search?" +
                     urllib.parse.urlencode({"q": q, "type": t, "limit": 10})) or {}
        for x in (d.get(t + "s") or {}).get("items", []):
            if nl.same(x.get("name"), answer):
                return "Spotify"
    if "apple" in sources:
        ent = {"曲名": "song", "アルバム名": "album", "アーティスト名": "musicArtist"}[kind]
        key = {"曲名": "trackName", "アルバム名": "collectionName", "アーティスト名": "artistName"}[kind]
        term = answer if kind == "アーティスト名" else f"{answer} {artist}"
        d = http.get("apple", "https://itunes.apple.com/search?" +
                     urllib.parse.urlencode({"term": term, "entity": ent, "limit": 10, "country": "US"})) or {}
        if any(nl.same(x.get(key), answer) for x in d.get("results", [])):
            return "Apple Music"
    if "discogs" in sources and kind != "曲名":
        p = {"type": "release" if kind == "アルバム名" else "artist", "per_page": 10}
        p["release_title" if kind == "アルバム名" else "q"] = answer
        d = http.get("discogs", "https://api.discogs.com/database/search?" + urllib.parse.urlencode(p)) or {}
        for x in d.get("results", []):
            title = x.get("title", "")
            name = title.split(" - ", 1)[1] if (kind == "アルバム名" and " - " in title) else title
            if nl.same(nl.discogs_name(name), answer):
                return "Discogs"
    return None


def count_red(excel_path):
    wb = openpyxl.load_workbook(excel_path, read_only=False)
    n = len(_red_cells(wb["IDTAG"]))
    wb.close()
    return n


def _red_cells(ws):
    cells = []
    for row in ws.iter_rows(min_row=2):
        for col in COLS:
            c = row[col]
            if c.value and nl.JP_RE.search(str(c.value)) and c.fill is not None and c.fill.fill_type \
                    and str(c.fill.fgColor.rgb).upper().endswith(RED_RGB):
                cells.append((row, col))
    return cells


def run(excel_path, sources, progress_callback=None, engine="claude"):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["IDTAG"]
    cells = _red_cells(ws)

    # one question per distinct name in its context (the same title can appear on several rows)
    items, key_of = [], {}
    for row, col in cells:
        k = (col, str(row[col].value), str(row[nl.COL_ARTIST].value or ""), str(row[nl.COL_ALBUM].value or ""))
        if k not in key_of:
            key_of[k] = len(items) + 1
            items.append({"id": key_of[k], "kind": KIND[col], "text": k[1], "artist": k[2], "album": k[3]})

    name = ENGINES[engine]
    answers = {}
    for start in range(0, len(items), BATCH):
        batch = items[start:start + BATCH]
        if progress_callback:
            progress_callback(start, len(items), f"{name} に問い合わせ中 {start + 1}〜{start + len(batch)}/{len(items)}")
        answers.update(ask_ai(batch, engine))

    keys = nl.load_api_keys()
    sources = [s for s in nl.available_sources(keys) if s in sources]
    http = nl.Http(keys)
    verified = {}
    for n, it in enumerate(items, 1):
        if progress_callback:
            progress_callback(n - 1, len(items), f"{name} の答えを確かめています {n}/{len(items)}：{it['text']}")
        ans, western = answers.get(it["id"], (None, False))
        if not ans or ans == it["text"]:
            continue
        artist = it["artist"] if not nl.JP_RE.search(it["artist"]) else ""
        verified[it["id"]] = _found(http, sources, it["kind"], ans, artist) if sources else None

    stats = {"green": 0, "orange": 0, "kept": 0, "none": 0}
    review = {}
    for row, col in cells:
        c = row[col]
        k = (col, str(c.value), str(row[nl.COL_ARTIST].value or ""), str(row[nl.COL_ALBUM].value or ""))
        i = key_of[k]
        ans, western = answers.get(i, (None, False))
        old = c.value
        if not ans:
            stats["none"] += 1
            continue
        if ans == old:  # the AI says the Japanese is the official name (邦楽)
            c.fill, kind = nl.BLUE, "kept"
        elif verified.get(i):
            c.value, c.fill, kind = ans, nl.GREEN, "green"
        else:
            c.value, c.fill, kind = ans, nl.ORANGE, "orange"
        c.font = Font(bold=True)
        stats[kind] += 1
        rk = (kind, COLS[col], str(old), ans)
        review[rk] = (review.get(rk, (0, None))[0] + 1, verified.get(i))

    if nl.REVIEW_SHEET in wb.sheetnames:
        rv = wb[nl.REVIEW_SHEET]
    else:
        rv = wb.create_sheet(nl.REVIEW_SHEET)
        for i, h in enumerate(["区分", "列", "修正前", "修正後", "行数"], 1):
            rv.cell(row=1, column=i, value=h)
    label = {"green": f"{name} の案（サービスで確認済み・採用）", "orange": f"{name} の案（要確認・裏付けなし）",
             "kept": f"{name}：正式表記のまま"}
    fill = {"green": nl.GREEN, "orange": nl.ORANGE, "kept": nl.BLUE}
    for (kind, col_name, old, new), (count, src) in sorted(review.items(), key=lambda x: (x[0][0] != "orange", x[0][1])):
        rv.append([label[kind], col_name, old, new, count] + ([f"確認：{src}"] if src else []))
        for cell in rv[rv.max_row]:
            cell.fill = fill[kind]

    wb.save(excel_path)
    if progress_callback:
        progress_callback(len(items), len(items), "完了")
    return {"cells": len(cells), "names": len(items), **stats}
