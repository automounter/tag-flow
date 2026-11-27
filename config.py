# -*- coding: utf-8 -*-
import json
import os

from platform_util import config_dir

CONFIG_DIR = config_dir()
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")

_LEGACY_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

DEFAULTS = {
    "source_folder": "",
    "output_folder": "",
    "excel_folder": "",
    "excel_path": "",
    "genre_rule_folder": "",
    "genre_alias_fallback_enabled": False,
    "theme": "dark",
    "small_file_threshold_kb": 10,
    "id3_version": "v24",
    "strip_id3v1": False,
    "mp3db_api_url": "",
    "mp3db_api_key": "",
    "api_keys_file": "",  # text file with the Discogs / Spotify keys (kept outside the app)
}

def load_config():
    path = CONFIG_PATH
    if not os.path.exists(path) and os.path.exists(_LEGACY_CONFIG_PATH):
        path = _LEGACY_CONFIG_PATH  # migrate from pre-packaged dev location
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            merged = dict(DEFAULTS)
            merged.update(data)
            return merged
        except Exception:
            pass
    return dict(DEFAULTS)

def save_config(cfg):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
