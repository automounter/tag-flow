# -*- coding: utf-8 -*-
"""STEP 03 の後：辞書で変換できなかった日本語の曲名・アーティスト名・アルバム名を
Apple Music・Spotify・Discogs に照会し、2か所以上で一致した表記を採用する。

方針（利用者の決定）
- 洋楽：日本語の題は正式名と認めず、原題（英字）を採用する。
- 邦楽：本人・盤の正式表記のまま（カタカナが正式ならカタカナ、英字が正式なら英字）。
- やり直し：前回の結果で緑になるものは照会しない。それ以外は未照会のサービスにだけ問い合わせる。
"""
import base64
import json
import os
import re
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, OrderedDict, defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill

JP_RE = re.compile(r'[぀-ゟ゠-ヿ一-鿿々]')
KANA_KANJI_RE = re.compile(r'[぀-ゟ一-鿿々]')  # ひらがな・漢字（カタカナだけの名前は洋楽の音訳が多い）
# 「K.C.&ザ・サンシャイン・バンド」のように空白なしの & は名前の一部なので区切らない
SPLIT_RE = re.compile(r'(\s*,\s*|\s+&\s+|\s*、\s*|\s+(?:feat|ft)\.\s+)', re.I)

GREEN = PatternFill(fgColor="C6EFCE", fill_type="solid")
ORANGE = PatternFill(fgColor="FFC000", fill_type="solid")
BLUE = PatternFill(fgColor="DDEBF7", fill_type="solid")
RED = PatternFill(fgColor="FF9999", fill_type="solid")

# Apple の日本のストアのジャンルで邦楽と判断するもの
JP_GENRES = {"J-Pop", "歌謡曲", "演歌", "アニメ", "J-Rock", "J-Hip-Hop", "J-Dance", "J-R&B",
             "J-Jazz", "Kayokyoku", "Enka", "Anime", "Japanese Pop"}

COL_TITLE, COL_ARTIST, COL_ALBUM, COL_TRACK, COL_DISC, COL_AA, COL_FOLDER, COL_PATH = 2, 3, 4, 5, 7, 12, 14, 16
REVIEW_SHEET = "確認リスト"


# ---------- キー ----------

def load_api_keys(path=None):
    """設定画面で指定したキーのファイルの【Discogs】【Spotify …】の節からキーを読む。値はどこにも書き出さない。"""
    if not path:
        from config import load_config
        path = load_config().get("api_keys_file")
    if not path:
        return {}
    path = os.path.expanduser(path)
    keys = {}
    if not os.path.exists(path):
        return keys
    section = ""
    with open(path, encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if s.startswith("【"):
                section = s.lower()
                continue
            m = re.match(r'^([^：:]+)[：:]\s*(\S+)\s*$', s)
            if not m:
                continue
            label, value = m.group(1).strip().lower(), m.group(2)
            if "discogs" in section and "token" in label:
                keys["discogs_token"] = value
            elif "spotify" in section and label in ("クライアントid", "client id"):
                keys["spotify_id"] = value
            elif "spotify" in section and label in ("クライアントシークレット", "client secret"):
                keys["spotify_secret"] = value
    return keys


# ---------- HTTP ----------

class Http:
    INTERVAL = {"apple": 3.2, "discogs": 1.1, "spotify": 0.2}

    def __init__(self, keys):
        self.keys = keys
        self.last = defaultdict(float)
        self.spotify_token = None
        self.on_wait = None  # fn(message); may raise to cancel (the app's progress callback does)

    def _sleep(self, seconds, reason):
        # sleep in small steps so the status bar shows the wait and Cancel takes effect quickly
        end = time.time() + seconds
        while True:
            left = end - time.time()
            if left <= 0:
                return
            if self.on_wait:
                self.on_wait(f"{reason}(あと{int(left) + 1}秒)")
            time.sleep(min(1.0, left))

    def _wait(self, src):
        w = self.INTERVAL[src] - (time.time() - self.last[src])
        if w > 0:
            time.sleep(w)
        self.last[src] = time.time()

    def get(self, src, url, headers=None):
        headers = dict(headers or {})
        headers.setdefault("User-Agent", "TagFlow/1.0")
        names = {"apple": "Apple Music", "spotify": "Spotify", "discogs": "Discogs"}
        for attempt in range(3):
            self._wait(src)
            if src == "discogs":
                headers["Authorization"] = f"Discogs token={self.keys['discogs_token']}"
            if src == "spotify":
                if not self.spotify_token:
                    self._spotify_auth()
                headers["Authorization"] = f"Bearer {self.spotify_token}"
            try:
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=25) as r:
                    return json.load(r)
            except urllib.error.HTTPError as e:
                if e.code == 404:
                    return None
                if e.code == 401 and src == "spotify":
                    self.spotify_token = None
                    continue
                if e.code in (403, 429, 500, 502, 503):
                    self._sleep(min(int(e.headers.get("Retry-After") or 30), 120),
                                f"{names[src]} の問い合わせ制限のため待機中 {attempt + 1}/3回目")
                    continue
                raise
            except (urllib.error.URLError, TimeoutError, ConnectionError):
                self._sleep(10, f"{names[src]} に接続できないため待機中 {attempt + 1}/3回目")
        return None

    def _spotify_auth(self):
        cred = base64.b64encode(f"{self.keys['spotify_id']}:{self.keys['spotify_secret']}".encode()).decode()
        req = urllib.request.Request("https://accounts.spotify.com/api/token", data=b"grant_type=client_credentials",
                                     headers={"Authorization": "Basic " + cred})
        with urllib.request.urlopen(req, timeout=20) as r:
            self.spotify_token = json.load(r)["access_token"]


# ---------- 文字の比較 ----------

def norm(s):
    s = unicodedata.normalize("NFKC", str(s or "")).lower().replace("&", "and")
    s = re.sub(r'\s*-\s*(single|ep)$', '', s)
    s = re.sub(r'^the\s+', '', s)  # The Commodores と Commodores を同じとみなす
    return re.sub(r'[\s・･.,\'’"`\-_!?:;/()\[\]（）「」『』~〜]', '', s)


def core(s):
    """括弧の補足（feat. やバージョン表記）と、Spotify 式の「 - Single Version」などを除いた部分。"""
    s = str(s or "")
    c = re.sub(r'\s+-\s+[^-]*$', '', s)
    c = re.sub(r'\s*[\(（\[].*$', '', c).strip()
    return c or s


def same(a, b):
    return bool(a and b) and (norm(a) == norm(b) or norm(core(a)) == norm(core(b)))


def split_artist(s):
    return [p.strip() for p in SPLIT_RE.split(str(s or "")) if p.strip() and not SPLIT_RE.fullmatch(p)]


def discogs_name(s):
    return re.sub(r'\s*\(\d+\)$', '', (s or "").rstrip("*")).strip()


def discogs_parts(s):
    """Discogs の日本盤は「素顔のままで = Just The Way You Are」のように併記される。"""
    return [p.strip() for p in str(s or "").split(" = ") if p.strip()]


# ---------- 照会 ----------

ALL_SOURCES = ("apple", "spotify", "discogs")


def available_sources(keys):
    ok = {"apple": True, "spotify": bool(keys.get("spotify_id") and keys.get("spotify_secret")),
          "discogs": bool(keys.get("discogs_token"))}
    return [s for s in ALL_SOURCES if ok[s]]


class Lookup:
    def __init__(self, keys, cache, sources=None):
        self.http = Http(keys)
        self.cache = cache
        allowed = set(available_sources(keys)) & set(sources or ALL_SOURCES)
        self.use_apple = "apple" in allowed
        self.use_discogs = "discogs" in allowed
        self.use_spotify = "spotify" in allowed
        self.sources = sorted(allowed)
        self.seen = set()  # 今回の実行で判断済みの (kind, key)。同じものを2回照会しない

    def _need(self, kind, key, confirmed=None):
        """返り値：(info, 今回問い合わせるサービス)。
        前回の結果で緑になるもの（confirmed が True）と、今のサービスをすべて照会済みのものは問い合わせない。
        それ以外は、まだ照会していないサービスにだけ問い合わせ、前回の結果に足す。
        （サービスを選べるようになる前の結果は、全サービスで照会したもの）"""
        entry = self.cache.setdefault(kind, {}).get(key)
        if entry is None:
            return {}, set(self.sources)
        done = set(entry.get("_sources", ALL_SOURCES))
        missing = set(self.sources) - done
        if (kind, key) in self.seen or not missing or (confirmed and confirmed(entry)):
            self.seen.add((kind, key))
            return entry, set()
        return dict(entry, _sources=sorted(done)), missing

    def _store(self, kind, key, info, asked):
        info["_sources"] = sorted(set(info.get("_sources", [])) | asked)
        self.cache.setdefault(kind, {})[key] = info
        self.seen.add((kind, key))
        return info

    def _apple_search(self, term, entity, limit=10):
        q = urllib.parse.urlencode({"term": term, "entity": entity, "country": "JP", "lang": "ja_jp", "limit": limit})
        d = self.http.get("apple", "https://itunes.apple.com/search?" + q)
        return (d or {}).get("results", [])

    def _apple_lookup(self, ids, country, entity=None):
        p = {"id": ",".join(map(str, ids)), "country": country, "lang": "ja_jp" if country == "JP" else "en_us"}
        if entity:
            p["entity"] = entity
        d = self.http.get("apple", "https://itunes.apple.com/lookup?" + urllib.parse.urlencode(p))
        return (d or {}).get("results", [])

    # --- アーティスト ---
    def artist(self, name, confirmed=None):
        info, ask = self._need("artist", name, confirmed)
        if not ask:
            return info
        if "apple" in ask:
            info.update(apple_jp=None, apple_us=None, apple_genre=None)
            # 日本のストアでも洋楽のアーティスト名は英字で返るので、名前の一致ではなく先頭の結果を使う
            res = self._apple_search(name, "musicArtist", 5)
            if res:
                x = res[0]
                info["apple_jp"], info["apple_genre"] = x["artistName"], x.get("primaryGenreName")
                us = self._apple_lookup([x["artistId"]], "US")
                info["apple_us"] = us[0].get("artistName") if us else None
        if "spotify" in ask:
            # Spotify には同じ人の日本語名の別登録（グローリア・ゲイナー と Gloria Gaynor など）があるので、
            # 上位3件から日本語名と英字名をそれぞれ拾う
            q = urllib.parse.urlencode({"q": name, "type": "artist", "market": "JP", "limit": 5})
            d = self.http.get("spotify", "https://api.spotify.com/v1/search?" + q)
            names = [x.get("name") or "" for x in ((d or {}).get("artists") or {}).get("items", [])[:3]]
            info["spotify_ja"] = next((n for n in names if same(n, name)), names[0] if names else None)
            info["spotify_en"] = next((n for n in names if not JP_RE.search(n)), None)
        return self._store("artist", name, info, ask)

    # --- アルバム（曲目ごと） ---
    def album(self, album, artist, latin=None, confirmed=None):
        key = f"{album}\t{artist}"
        info, ask = self._need("album", key, confirmed)
        if not ask:
            return info
        if "apple" in ask:
            info["apple"] = self._apple_album(album, artist)
        if "spotify" in ask:
            info["spotify"] = self._spotify_album(album, artist)
            for alt in (latin, ""):
                if not info["spotify"] and alt != artist:
                    info["spotify"] = self._spotify_album(album, alt or "")
        if "discogs" in ask:
            # Discogs のアーティスト名は英字なので、英字名で検索する
            info["discogs"] = self._discogs_album(album, latin or artist)
        return self._store("album", key, info, ask)

    def _apple_album(self, album, artist):
        res = self._apple_search(f"{album} {artist}".strip(), "album") or self._apple_search(album, "album")
        hit = next((x for x in res if same(x.get("collectionName"), album)), None)
        if not hit:
            return None
        jp = [x for x in self._apple_lookup([hit["collectionId"]], "JP", "song") if x.get("wrapperType") == "track"]
        us = {x["trackId"]: x for x in self._apple_lookup([hit["collectionId"]], "US", "song")
              if x.get("wrapperType") == "track"}
        us_album = next((x.get("collectionName") for x in us.values()), None)
        tracks = []
        for x in jp:
            u = us.get(x["trackId"], {})
            tracks.append({"disc": x.get("discNumber"), "track": x.get("trackNumber"),
                           "jp": x.get("trackName"), "us": u.get("trackName"),
                           "artist_jp": x.get("artistName"), "artist_us": u.get("artistName")})
        return {"jp": hit.get("collectionName"), "us": us_album, "genre": hit.get("primaryGenreName"), "tracks": tracks}

    def _spotify_album(self, album, artist):
        q = urllib.parse.urlencode({"q": f"{album} {artist}".strip(), "type": "album", "market": "JP", "limit": 5})
        d = self.http.get("spotify", "https://api.spotify.com/v1/search?" + q)
        hit = next((x for x in ((d or {}).get("albums") or {}).get("items", []) if same(x.get("name"), album)), None)
        if not hit:
            return None
        tracks, url = [], f"https://api.spotify.com/v1/albums/{hit['id']}/tracks?market=JP&limit=50"
        while url:
            d = self.http.get("spotify", url) or {}
            for x in d.get("items", []):
                tracks.append({"disc": x.get("disc_number"), "track": x.get("track_number"), "name": x.get("name"),
                               "artist": ", ".join(a["name"] for a in x.get("artists", []))})
            url = d.get("next")
        return {"name": hit.get("name"), "tracks": tracks}

    def _discogs_album(self, album, artist):
        q = urllib.parse.urlencode({"q": f"{artist} {album}".strip(), "type": "release", "per_page": 8})
        d = self.http.get("discogs", "https://api.discogs.com/database/search?" + q) or {}
        hit = None
        for x in d.get("results", []):
            title = x.get("title", "")
            album_part = title.split(" - ", 1)[1] if " - " in title else title
            if any(same(p, album) for p in discogs_parts(album_part)):
                hit = x
                break
        if not hit:
            return None
        rel = self.http.get("discogs", f"https://api.discogs.com/releases/{hit['id']}") or {}
        rel_artist = ", ".join(discogs_name(a.get("name")) for a in rel.get("artists", []))
        tracks = [{"title": t.get("title"),
                   "artist": ", ".join(discogs_name(a.get("name")) for a in t.get("artists", [])) or rel_artist}
                  for t in rel.get("tracklist", []) if t.get("type_", "track") == "track"]
        return {"title": rel.get("title"), "artist": rel_artist, "tracks": tracks}

    # --- アルバムが見つからなかった曲 ---
    def song(self, title, artist, latin=None):
        key = f"{title}\t{artist}"
        info, ask = self._need("song", key)
        if not ask:
            return info

        def artist_ok(n):
            return same(n, artist) or (latin and (same(n, latin) or norm(latin) in norm(n)))

        if "apple" in ask:
            info["apple"] = None
            res = self._apple_search(f"{core(title)} {artist}".strip(), "song", 25)
            hit = next((x for x in res if same(x.get("trackName"), title) and artist_ok(x.get("artistName"))), None)
            if hit:
                us = self._apple_lookup([hit["trackId"]], "US")
                info["apple"] = {"jp": hit.get("trackName"), "us": us[0].get("trackName") if us else None}
        if "spotify" in ask:
            info["spotify"] = None
            q = urllib.parse.urlencode({"q": f"{core(title)} {artist}".strip(), "type": "track", "market": "JP", "limit": 5})
            d = self.http.get("spotify", "https://api.spotify.com/v1/search?" + q)
            for x in ((d or {}).get("tracks") or {}).get("items", []):
                if any(artist_ok(a["name"]) for a in x.get("artists", [])):
                    info["spotify"] = x.get("name")
                    break
        return self._store("song", key, info, ask)


# ---------- 判定 ----------

def is_japanese_artist(info, name):
    if info.get("apple_genre") in JP_GENRES:
        return True
    if info.get("apple_jp") and JP_RE.search(info["apple_jp"]):
        return True  # Apple の日本のストアは洋楽を英字、邦楽を日本語で表示する
    if info.get("apple_jp") or info.get("spotify_ja"):
        return False
    return bool(KANA_KANJI_RE.search(name))  # どこにも無い：ひらがな・漢字を含む名前は邦楽とみなす


def decide(current, cands, western):
    """cands: [(source, name)]。洋楽では日本語の候補を捨てる。
    返り値：(新しい値, 区分) 区分は apple(緑)/check(橙)/keep(青)/none(赤)。"""
    cands = [(s, n) for s, n in cands if n]
    if western:
        cands = [(s, n) for s, n in cands if not JP_RE.search(n)]
    if not cands:
        return current, "none"
    groups = OrderedDict()
    for s, n in cands:
        groups.setdefault(norm(core(n)), []).append((s, n))
    best = max(groups.values(), key=lambda g: len({s for s, _ in g}))
    value = re.sub(r'\s+-\s+[^-]*$', '', best[0][1]).strip()  # Spotify 式の「 - Single Version」は付けない
    suffix = re.search(r'\s*[\(（][^()（）]*[\)）]$', str(current))
    if suffix and not JP_RE.search(suffix.group()) and "(" not in value:
        value += " " + suffix.group().strip()  # 元のタグの英字の補足（(Remastered) など）は残す
    if norm(value) == norm(current):
        return current, "keep"
    return value, ("apple" if len({s for s, _ in best}) >= 2 else "check")


def _match_track(tracks, title, disc, track, name_keys):
    for t in tracks:
        if any(same(t.get(k), title) for k in name_keys):
            return t
        for k in name_keys:
            if any(same(p, title) for p in discogs_parts(t.get(k))):
                return t
    if track:
        for t in tracks:
            if t.get("track") == track and (t.get("disc") or 1) == (disc or 1):
                return t
    return None


def _num(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


# ---------- 実行 ----------

def cache_path(excel_path):
    base, _ = os.path.splitext(excel_path)
    return f"{base}_name_lookup.json"


def count_targets(excel_path):
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    n = 0
    for r in wb["IDTAG"].iter_rows(min_row=2, values_only=True):
        if any(r[c] and JP_RE.search(str(r[c])) for c in (COL_TITLE, COL_ARTIST, COL_ALBUM, COL_AA)):
            n += 1
    wb.close()
    return n


def run(excel_path, progress_callback=None, keys_file=None, sources=None):
    keys = load_api_keys(keys_file)
    cpath = cache_path(excel_path)
    try:
        with open(cpath, encoding="utf-8") as f:
            cache = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        cache = {}
    lk = Lookup(keys, cache, sources)
    if not lk.sources:
        raise ValueError("照会先のサービスが1つも選ばれていません。")
    state = {"done": 0, "total": 1, "message": ""}

    def save_cache():
        with open(cpath, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["IDTAG"]
    base, ext = os.path.splitext(excel_path)
    original = {}
    if os.path.exists(f"{base}_baseline{ext}"):
        bw = openpyxl.load_workbook(f"{base}_baseline{ext}", read_only=True)
        original = {r[COL_PATH]: r for r in bw["IDTAG"].iter_rows(min_row=2, values_only=True) if r[COL_PATH]}
        bw.close()

    rows = [r for r in ws.iter_rows(min_row=2)
            if any(r[c].value and JP_RE.search(str(r[c].value)) for c in (COL_TITLE, COL_ARTIST, COL_ALBUM, COL_AA))]

    def orig(r, col):
        o = original.get(r[COL_PATH].value)
        return (o[col] if o and o[col] else None) or r[col].value or ""

    # 1) アーティスト（名前全体と、区切った各部分のうち日本語のもの）
    names = set()
    for r in rows:
        for c in (COL_ARTIST, COL_AA):
            v = str(r[c].value or "")
            if JP_RE.search(v):
                names.add(v.strip())
                names.update(p for p in split_artist(v) if JP_RE.search(p))
    names = sorted(names)
    albums = OrderedDict()
    for r in rows:
        main = (split_artist(orig(r, COL_AA)) or split_artist(orig(r, COL_ARTIST)) or [""])[0]
        albums.setdefault((orig(r, COL_ALBUM), main), []).append(r)
    total = len(names) + len(albums)
    done = 0
    state["total"] = total

    def report(current, total_, message):
        # a Cancel raises out of progress_callback: keep what was looked up so far
        state["done"] = current
        try:
            progress_callback(current, total_, message)
            state["message"] = message
        except BaseException:
            save_cache()
            raise

    if progress_callback:
        # keep "アルバムを照会中 5/71：…" visible while waiting
        def on_wait(msg):
            base = state["message"]
            report(state["done"], state["total"], f"{base} ／ {msg}" if base else msg)
            state["message"] = base
        lk.http.on_wait = on_wait

    # Discogs の盤のアーティスト名（英字）も、そのアーティストの票に加える（2) の後で集める）
    discogs_votes = defaultdict(list)

    def artist_one(name, info=None):
        if info is None:
            info = lk.artist(name)
        if is_japanese_artist(info, name):
            cands = [("apple", info.get("apple_jp")), ("spotify", info.get("spotify_ja"))]
            return decide(name, cands, False) + (cands,)
        cands = [("apple", info.get("apple_us") or info.get("apple_jp")), ("spotify", info.get("spotify_en"))]
        votes = Counter(discogs_votes.get(name, []))
        if votes:
            cands.append(("discogs", votes.most_common(1)[0][0]))
        return decide(name, cands, True) + (cands,)

    def resolve_artist(value):
        value = str(value).strip()
        whole = artist_one(value)
        # 「ピーチズ & ハーブ」のようなグループ名は全体で採用する。ただし結果が1人の名前なら
        # （パティ・ラベル & マイケル・マクドナルド → Michael McDonald）人ごとに分けて照会し直す
        if not SPLIT_RE.search(value) or (whole[1] == "apple" and re.search(r'&|,| and ', whole[0], re.I)):
            return whole
        out, kinds, all_cands = [], set(), []
        for tok in SPLIT_RE.split(value):
            if not tok:
                continue
            if SPLIT_RE.fullmatch(tok) or not JP_RE.search(tok):
                out.append(tok)
                continue
            v, k, c = artist_one(tok.strip())
            out.append(v)
            kinds.add(k)
            all_cands += c
        kind = next(k for k in ("none", "check", "apple", "keep") if k in kinds)
        if kind != "apple" and whole[1] == "check" and re.search(r'&|,| and ', whole[0], re.I):
            return whole  # 分けても決まらないときは、グループ名としての候補を橙で出す
        joined = OrderedDict()
        for src, n in all_cands:
            joined.setdefault(src, []).append(n or "-")
        return "".join(out), kind, [(src, " / ".join(ns)) for src, ns in joined.items()]

    # 1) アーティスト（前回の結果で緑になるものは問い合わせない）
    for n in names:
        if progress_callback:
            report(done, total, f"アーティストを照会中 {done + 1}/{total}：{n}")
        lk.artist(n, confirmed=lambda info, n=n: artist_one(n, info)[1] == "apple")
        done += 1
        if done % 10 == 0:
            save_cache()
    save_cache()

    def latin_name(name):
        if not name or not JP_RE.search(name):
            return name
        info = lk.artist(name)
        return next((n for n in (info.get("apple_us"), info.get("apple_jp"), info.get("spotify_en"))
                     if n and not JP_RE.search(n)), None)

    def judge_album(album, main, album_rows, info, query=True):
        """トラック名・アルバム名の判定。返り値：[(セル, 新しい値, 区分, 列名, 候補)]。
        query=False では曲の照会をせず、前回の結果だけで判定する（緑かどうかの事前確認用）。"""
        ap, sp, dg = info.get("apple"), info.get("spotify"), info.get("discogs")
        if main and JP_RE.search(main):
            western = not is_japanese_artist(lk.artist(main), main)
        elif ap and ap.get("genre"):
            western = ap["genre"] not in JP_GENRES  # 英字名の邦楽（DREAMS COME TRUE など）はジャンルで見分ける
        else:
            western = True
        out = []
        for r in album_rows:
            # トラック名
            cur = r[COL_TITLE].value
            if cur and JP_RE.search(str(cur)):
                title = orig(r, COL_TITLE)
                row_artist = (split_artist(orig(r, COL_ARTIST)) or [main])[0]
                row_western = western
                if row_artist != main and JP_RE.search(row_artist):  # オムニバスは曲ごとのアーティストで判断
                    row_western = not is_japanese_artist(lk.artist(row_artist), row_artist)
                disc, track = _num(r[COL_DISC].value), _num(r[COL_TRACK].value)
                cands = []
                a = _match_track(ap["tracks"], title, disc, track, ("jp",)) if ap else None
                s = _match_track(sp["tracks"], title, disc, track, ("name",)) if sp else None
                d = _match_track(dg["tracks"], title, None, None, ("title",)) if dg else None
                if a:
                    cands.append(("apple", (a["us"] or a["jp"]) if row_western else a["jp"]))
                if s:
                    cands.append(("spotify", s["name"]))
                if d:
                    parts = discogs_parts(d["title"])
                    pick = [p for p in parts if bool(JP_RE.search(p)) != row_western] or parts
                    cands.append(("discogs", pick[0]))
                if not a and not s:
                    if query:
                        sg = lk.song(title, row_artist, latin_name(row_artist))
                    else:
                        sg = lk.cache.get("song", {}).get(f"{title}\t{row_artist}") or {}
                    if sg.get("apple"):
                        cands.append(("apple", (sg["apple"]["us"] or sg["apple"]["jp"]) if row_western else sg["apple"]["jp"]))
                    if sg.get("spotify"):
                        cands.append(("spotify", sg["spotify"]))
                new, kind = decide(str(cur), cands, row_western)
                out.append((r[COL_TITLE], new, kind, "トラック名", cands))

            # アルバム名
            cur = r[COL_ALBUM].value
            if cur and JP_RE.search(str(cur)):
                cands = []
                if ap:
                    cands.append(("apple", (ap["us"] or ap["jp"]) if western else ap["jp"]))
                if sp:
                    cands.append(("spotify", sp["name"]))
                if dg:
                    parts = discogs_parts(dg["title"])
                    pick = [p for p in parts if bool(JP_RE.search(p)) != western] or parts
                    cands.append(("discogs", pick[0]))
                new, kind = decide(str(cur), cands, western)
                out.append((r[COL_ALBUM], new, kind, "アルバム名", cands))
        return out

    # 2) アルバム（曲目ごと。前回の結果でトラック名・アルバム名がすべて緑になるものは問い合わせない）
    for (album, main), album_rows in albums.items():
        if progress_callback:
            report(done, total, f"アルバムを照会中 {done + 1}/{total}：{album}")
        needs_album = any(r[c].value and JP_RE.search(str(r[c].value))
                          for r in album_rows for c in (COL_TITLE, COL_ALBUM))
        if album and needs_album:
            lk.album(album, main, latin_name(main),
                     confirmed=lambda info, a=album, m=main, rs=album_rows:
                     all(k == "apple" for _, _, k, _, _ in judge_album(a, m, rs, info, query=False)))
        done += 1
        if done % 5 == 0:
            save_cache()
    save_cache()

    # 3) 判定と書き込み
    for (album, main) in albums:
        dg = (lk.cache.get("album", {}).get(f"{album}\t{main}") or {}).get("discogs")
        if dg and dg.get("artist") and JP_RE.search(main):
            discogs_votes[main].append(dg["artist"])

    stats = Counter()
    review = OrderedDict()
    sources_of = {}

    def set_cell(cell, new, kind, col_name, srcs):
        stats[kind] += 1
        old = cell.value
        if kind != "none":
            cell.value = new
        cell.fill = {"apple": GREEN, "check": ORANGE, "keep": BLUE, "none": RED}[kind]
        cell.font = Font(bold=True)
        if kind != "apple":
            k = (kind, col_name, str(old), str(cell.value))
            review[k] = review.get(k, 0) + 1
            sources_of[k] = srcs

    for i, ((album, main), album_rows) in enumerate(albums.items(), 1):
        if progress_callback:
            report(total, total, f"表記を判定中 {i}/{len(albums)}：{album}")
        info = lk.cache.get("album", {}).get(f"{album}\t{main}") or {}
        for cell, new, kind, col_name, cands in judge_album(album, main, album_rows, info):
            set_cell(cell, new, kind, col_name, cands)
        # アーティスト名・アルバムアーティスト
        for r in album_rows:
            for col, col_name in ((COL_ARTIST, "アーティスト名"), (COL_AA, "アルバムアーティスト")):
                cur = r[col].value
                if cur and JP_RE.search(str(cur)):
                    new, kind, cands = resolve_artist(cur)
                    set_cell(r[col], new, kind, col_name, cands)

    save_cache()

    # 確認リスト（緑以外）
    # rows added by "AI で補う" (ai_fill) describe cells this run no longer touches; keep them
    ai_rows = []
    if REVIEW_SHEET in wb.sheetnames:
        for r in wb[REVIEW_SHEET].iter_rows(min_row=2):
            if str(r[0].value or "").startswith(("Claude", "Codex")):
                ai_rows.append([(c.value, c.fill.fgColor.rgb if c.fill.fill_type else None) for c in r])
        del wb[REVIEW_SHEET]
    rv = wb.create_sheet(REVIEW_SHEET)
    rv.append(["区分", "列", "修正前", "修正後", "行数", "Apple", "Spotify", "Discogs"])
    label = {"none": "未解決（候補なし・そのまま）", "check": "要確認（1か所だけ・食い違い）", "keep": "正式表記のまま"}
    fill = {"none": RED, "check": ORANGE, "keep": BLUE}
    rank = {"none": 0, "check": 1, "keep": 2}
    for k in sorted(review, key=lambda k: (rank[k[0]], k[1], k[2])):
        srcs = dict(sources_of.get(k) or [])
        rv.append([label[k[0]], k[1], k[2], k[3], review[k], srcs.get("apple"), srcs.get("spotify"), srcs.get("discogs")])
        for cell in rv[rv.max_row]:
            cell.fill = fill[k[0]]
    for saved in ai_rows:
        rv.append([v for v, _ in saved])
        for cell, (_, rgb) in zip(rv[rv.max_row], saved):
            if rgb:
                cell.fill = PatternFill(fgColor=rgb, fill_type="solid")
    for col, width in zip("ABCDEFGH", [28, 18, 48, 48, 7, 36, 36, 36]):
        rv.column_dimensions[col].width = width
    rv.freeze_panes = "A2"

    wb.save(excel_path)
    if progress_callback:
        progress_callback(total, total, "完了")
    return {"rows": len(rows), "adopted": stats["apple"], "check": stats["check"], "kept": stats["keep"],
            "unresolved": stats["none"], "sources": [s for s, ok in (("Apple", lk.use_apple), ("Spotify", lk.use_spotify),
                                                                      ("Discogs", lk.use_discogs)) if ok]}
