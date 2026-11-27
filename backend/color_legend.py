# -*- coding: utf-8 -*-
"""「色の説明」sheets for the working Excel (STEP 07 ③) and the integrity report (②)."""
from openpyxl.styles import PatternFill, Font

SHEET = "色の説明"

# (fill RGB or None, font RGB or None, column, meaning, step)
WORKING = [
    ("C6EFCE", None, "曲名・アーティスト名・アルバム名", "英語に変換した(辞書・修正辞書、名前照会で2か所以上一致、AI の案をサービスで確認)", "STEP 03"),
    ("DDEBF7", None, "曲名・アーティスト名・アルバム名", "日本語のまま残した(邦楽の正式な表記など)", "STEP 03"),
    ("FFC000", None, "曲名・アーティスト名・アルバム名", "要確認(照会で1か所だけ・食い違い、AI の案で裏付けなし)", "STEP 03 名前照会・AI で補う"),
    ("FF9999", None, "曲名・アーティスト名・アルバム名", "未解決(候補なし。日本語のまま)", "STEP 03 名前照会"),
    (None, "FF0000", "曲名・アーティスト名", "赤い文字:辞書で変換できなかった(名前照会の前)", "STEP 03"),
    ("E6D9F2", None, "曲名・アーティスト名・アルバム名", "英語の表記ゆれ(空白・全角・大文字小文字)をそろえた", "STEP 03"),
    ("D9D2F0", None, "アーティスト名・アルバム名", "Web再調査(Apple Music)で表記を直した", "STEP 03 Web再調査"),
    ("FFF2CC", None, "アルバムアーティスト", "トラックのアーティストと違う値にした", "STEP 04"),
    ("FFD700", None, "アルバムアーティスト", "Various Artists にした", "STEP 04"),
    ("FFF2CC", None, "ジャンル", "Apple Music のジャンルの読み替えで決めた", "STEP 05"),
    ("FCE4EC", None, "ジャンル", "手で選んだ", "STEP 05"),
    ("C6EFCE", None, "アートワーク", "フォルダの画像に変更", "STEP 06"),
    ("DDEBF7", None, "アートワーク", "URL・推奨画像に変更", "STEP 06"),
    ("F8CBAD", None, "アートワーク", "アートワークを削除", "STEP 06"),
]

REPORT = [
    ("C6EFCE", "変更点一覧", "トラック名が変わった"),
    ("DDEBF7", "変更点一覧", "アーティスト名が変わった"),
    ("FFF2CC", "変更点一覧", "アルバム名が変わった"),
    ("FDE4CF", "変更点一覧", "アルバムアーティストが変わった"),
    ("E6D9F2", "変更点一覧", "ジャンルが変わった"),
    ("F8CBAD", "整合性チェック", "警告(曲順の重複・欠番、発売日の不一致、ファイル作成日の外れ値)"),
]

BOLD = Font(bold=True)


def _fill(rgb):
    return PatternFill(fgColor=rgb, fill_type="solid")


def _new_sheet(wb):
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    return wb.create_sheet(SHEET)


def add_working_legend(wb):
    ws = _new_sheet(wb)
    ws.append(["色", "列", "意味", "付ける段階"])
    for c in ws[1]:
        c.font = BOLD
    for fill, font, col, meaning, step in WORKING:
        ws.append(["見本", col, meaning, step])
        cell = ws.cell(row=ws.max_row, column=1)
        if fill:
            cell.fill = _fill(fill)
        if font:
            cell.font = Font(bold=True, color="FF" + font)
    ws.append([])
    ws.append(["", "", "同じ色でも列によって意味が違います(例:薄い黄色はアルバムアーティストとジャンルで別の意味)。"
                       "名前照会と AI で補った結果は「確認リスト」のシートにも載っています。"])
    for col, width in zip("ABCD", [8, 30, 80, 26]):
        ws.column_dimensions[col].width = width


def add_report_legend(wb):
    ws = _new_sheet(wb)
    ws.append(["色", "シート", "意味"])
    for c in ws[1]:
        c.font = BOLD
    for fill, sheet, meaning in REPORT:
        ws.append(["見本", sheet, meaning])
        ws.cell(row=ws.max_row, column=1).fill = _fill(fill)
    for col, width in zip("ABC", [8, 18, 70]):
        ws.column_dimensions[col].width = width


def ensure_working_legend(excel_path):
    """Adds (or refreshes) the legend sheet in the working Excel. Returns False when the
    file could not be saved (for example while it is open in Excel)."""
    import openpyxl
    wb = openpyxl.load_workbook(excel_path)
    add_working_legend(wb)
    try:
        wb.save(excel_path)
    except OSError:
        return False
    return True
