# -*- coding: utf-8 -*-

DARK = {
    "bg": "#1C1915",
    "bg_elevated": "#242019",
    "bg_elevated_2": "#2C2720",
    "border": "#463F34",
    "text": "#F5F0E6",
    "text_muted": "#B3A996",
    "text_faint": "#847A69",
    "accent": "#F0A868",
    "accent_text": "#1C1915",
    "green": "#4CAF74",
    "cyan": "#6FB8D6",
    "red": "#E0685F",
    "yellow": "#E0B84A",
    "gold": "#C99A3E",
    "teal": "#4FB6B6",
}

LIGHT = {
    "bg": "#F7F3EA",
    "bg_elevated": "#FFFFFF",
    "bg_elevated_2": "#ECE5D6",
    "border": "#D9D0BC",
    "text": "#2B2620",
    "text_muted": "#6B6151",
    "text_faint": "#948A76",
    "accent": "#C97A2E",
    "accent_text": "#FFFFFF",
    "green": "#2E8B57",
    "cyan": "#2E7A99",
    "red": "#C0392B",
    "yellow": "#B8860B",
    "gold": "#9C6B0E",
    "teal": "#1F7A7A",
}

from platform_util import FONT_UI, FONT_FIXED

FONT_DISPLAY = FONT_UI
FONT_BODY = FONT_UI
FONT_MONO = FONT_FIXED

def get_theme(name):
    return DARK if name == "dark" else LIGHT
