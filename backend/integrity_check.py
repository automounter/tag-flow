# -*- coding: utf-8 -*-
import os
import datetime
from collections import defaultdict, Counter
import openpyxl
from backend import color_legend
from openpyxl.styles import PatternFill, Font

FIELD_COLORS = {
    "トラック名": PatternFill(fgColor="C6EFCE", fill_type="solid"),
    "アーティスト名": PatternFill(fgColor="DDEBF7", fill_type="solid"),
    "アルバム名": PatternFill(fgColor="FFF2CC", fill_type="solid"),
    "アルバムアーティスト": PatternFill(fgColor="FDE4CF", fill_type="solid"),
    "ジャンル": PatternFill(fgColor="E6D9F2", fill_type="solid"),
}
WARN_FILL = PatternFill(fgColor="F8CBAD", fill_type="solid")
BOLD = Font(bold=True)
BOLD_RED = Font(bold=True, color="FFCC0000")

FIELD_COLS = [("トラック名", 2), ("アーティスト名", 3), ("アルバム名", 4),
              ("アルバムアーティスト", 12), ("ジャンル", 10)]


def _load_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["IDTAG"]
    rows = []
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        filepath = r[16].value
        if not filepath:
            continue
        rows.append({
            "no": r[0].value, "title": r[2].value, "artist": r[3].value, "album": r[4].value,
            "track_no": r[5].value, "track_total": r[6].value, "disc_no": r[7].value,
            "date": r[9].value, "genre": r[10].value, "albumartist": r[12].value,
            "filepath": filepath,
        })
    return rows


def _build_changes(baseline_rows, current_rows):
    baseline_by_path = {r["filepath"]: r for r in baseline_rows}
    changes = []
    for cur in current_rows:
        base = baseline_by_path.get(cur["filepath"])
        if base is None:
            continue
        for label, key in [("トラック名", "title"), ("アーティスト名", "artist"),
                            ("アルバム名", "album"), ("アルバムアーティスト", "albumartist"),
                            ("ジャンル", "genre")]:
            before, after = base.get(key), cur.get(key)
            if (before or "") != (after or ""):
                changes.append({
                    "no": cur["no"], "filepath": cur["filepath"], "field": label,
                    "before": before or "", "after": after or "",
                })
    return changes


def _check_consistency(current_rows, progress_callback=None):
    warnings = []
    groups = defaultdict(list)
    for r in current_rows:
        groups[(r["album"] or "", r["albumartist"] or r["artist"] or "")].append(r)

    # track order: grouped per-disc, since track numbers restart on each disc
    disc_groups = defaultdict(list)
    for r in current_rows:
        if not r["album"]:
            continue
        disc_groups[(r["album"], r["albumartist"] or r["artist"] or "", r["disc_no"] or 1)].append(r)

    for (album, albumartist, disc_no), items in disc_groups.items():
        track_nos = [r["track_no"] for r in items if r["track_no"]]
        dup = [t for t, c in Counter(track_nos).items() if c > 1]
        disc_label = f"(Disc {disc_no}) " if len(set(d for (a, aa, d) in disc_groups if a == album and aa == albumartist)) > 1 else ""
        if dup:
            warnings.append({
                "album": album, "albumartist": albumartist, "type": "曲順異常",
                "detail": f"{disc_label}トラック番号が重複しています: {sorted(dup)}",
            })
        if track_nos:
            expected_total = items[0].get("track_total") or len(items)
            missing = sorted(set(range(1, expected_total + 1)) - set(track_nos))
            if missing:
                warnings.append({
                    "album": album, "albumartist": albumartist, "type": "曲順異常",
                    "detail": f"{disc_label}欠番の可能性があります: {missing}",
                })

    # the per-file stat() below is the slow part on a NAS, so report progress here
    # (otherwise the status bar's remaining-time estimate never gets a data point)
    total_groups = len(groups)
    for gi, ((album, albumartist), items) in enumerate(groups.items(), 1):
        if progress_callback and gi % 20 == 0:
            progress_callback(gi, total_groups, f"ファイル作成日を確認中 {gi}/{total_groups}")
        if not album:
            continue

        # release date consistency
        dates = [str(r["date"]) for r in items if r["date"]]
        distinct_dates = sorted(set(dates))
        if len(distinct_dates) > 1:
            warnings.append({
                "album": album, "albumartist": albumartist, "type": "発売日不一致",
                "detail": f"同一アルバム内で日付が揃っていません: {', '.join(distinct_dates)}",
            })

        # file creation date outliers
        creation_dates = {}
        for r in items:
            path = r["filepath"]
            if not path or not os.path.exists(path):
                continue
            try:
                ts = os.stat(path).st_birthtime
            except (OSError, AttributeError):
                try:
                    ts = os.stat(path).st_mtime
                except OSError:
                    continue
            creation_dates[path] = datetime.date.fromtimestamp(ts)
        if creation_dates:
            day_counts = Counter(creation_dates.values())
            majority_day, _ = day_counts.most_common(1)[0]
            outliers = [p for p, d in creation_dates.items()
                        if abs((d - majority_day).days) > 1]
            if outliers:
                names = [os.path.basename(p) for p in outliers]
                warnings.append({
                    "album": album, "albumartist": albumartist, "type": "作成日の外れ値",
                    "detail": f"ファイル作成日が他の曲と大きく異なります({majority_day}が多数派): {', '.join(names)}",
                })

    return warnings


def run(baseline_path, current_path, output_dir, progress_callback=None):
    if progress_callback:
        progress_callback(0, 1, "ベースラインを読み込み中")
    baseline_rows = _load_rows(baseline_path)

    if progress_callback:
        progress_callback(0, 1, "現在のデータを読み込み中")
    current_rows = _load_rows(current_path)

    changes = _build_changes(baseline_rows, current_rows)
    warnings = _check_consistency(current_rows, progress_callback)

    out_wb = openpyxl.Workbook()
    ws1 = out_wb.active
    ws1.title = "変更点一覧"
    ws1.append(["No", "MP3ファイルパス", "項目", "変更前", "変更後"])
    for c in ws1[1]:
        c.font = BOLD
    for ch in changes:
        ws1.append([ch["no"], ch["filepath"], ch["field"], ch["before"], ch["after"]])
        row_idx = ws1.max_row
        fill = FIELD_COLORS.get(ch["field"])
        for col in range(1, 6):
            cell = ws1.cell(row=row_idx, column=col)
            cell.font = BOLD
            if fill:
                cell.fill = fill
    widths1 = [6, 55, 16, 34, 34]
    for col, wdt in zip("ABCDE", widths1):
        ws1.column_dimensions[col].width = wdt

    ws2 = out_wb.create_sheet("整合性チェック")
    ws2.append(["アルバム", "アルバムアーティスト", "警告種別", "詳細"])
    for c in ws2[1]:
        c.font = BOLD
    for wn in warnings:
        ws2.append([wn["album"], wn["albumartist"], wn["type"], wn["detail"]])
        row_idx = ws2.max_row
        for col in range(1, 5):
            cell = ws2.cell(row=row_idx, column=col)
            cell.fill = WARN_FILL
            cell.font = BOLD_RED if col == 3 else BOLD
    widths2 = [30, 26, 14, 60]
    for col, wdt in zip("ABCD", widths2):
        ws2.column_dimensions[col].width = wdt

    color_legend.add_report_legend(out_wb)

    os.makedirs(output_dir, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(output_dir, f"TagFloW_{ts}_整合性確認.xlsx")
    out_wb.save(out_path)

    if progress_callback:
        progress_callback(1, 1, "完了")

    return {"changes": len(changes), "warnings": len(warnings), "report_path": out_path}
