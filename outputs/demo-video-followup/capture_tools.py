"""Capture native PNG evidence and read-only Excel state for the new film."""
from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
from datetime import datetime
import json
from pathlib import Path
import subprocess
import sys
import time

import imageio_ffmpeg
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent
WORKSPACE = ROOT.parents[1]
CAPTURES = ROOT / "captures"
WINDOWS_PATH = ROOT / "windows.json"
WINDOWS = json.loads(WINDOWS_PATH.read_text(encoding="utf-8")) if WINDOWS_PATH.exists() else {"celly": 5443164, "excel": 395498}
USER32 = ctypes.windll.user32
USER32.SetProcessDPIAware()


def write_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def geometry():
    result = {}
    for name, handle in WINDOWS.items():
        rect = wintypes.RECT()
        if not USER32.GetWindowRect(handle, ctypes.byref(rect)):
            raise RuntimeError(f"Cannot read {name} window rectangle")
        result[name] = [rect.left, rect.top, rect.right, rect.bottom]
    return result


def snapshot(label):
    CAPTURES.mkdir(exist_ok=True)
    started = time.time()
    path = CAPTURES / f"{label}-desktop.png"
    subprocess.run([
        imageio_ffmpeg.get_ffmpeg_exe(), "-hide_banner", "-loglevel", "error",
        "-f", "gdigrab", "-framerate", "1", "-draw_mouse", "0",
        "-offset_x", "0", "-offset_y", "0", "-video_size", "1920x1020",
        "-i", "desktop", "-frames:v", "1", "-update", "1", "-y", str(path),
    ], check=True, timeout=20)
    rectangles = geometry()
    with Image.open(path) as desktop:
        for name, rect in rectangles.items():
            desktop.crop(rect).save(CAPTURES / f"{label}-{name}.png")
    point = wintypes.POINT()
    USER32.GetCursorPos(ctypes.byref(point))
    event = {
        "label": label, "time": started, "finished": time.time(),
        "windows": rectangles, "pointer": [point.x, point.y],
        "desktop": str(path),
    }
    manifest_path = ROOT / "capture.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else []
    manifest.append(event)
    write_json(manifest_path, manifest)
    print(json.dumps(event, ensure_ascii=False))


def excel_state(label):
    # Read only: never selects, activates, edits, formats, saves, or evaluates cells.
    import win32com.client
    excel = win32com.client.GetActiveObject("Excel.Application")
    matches = [book for book in excel.Workbooks if book.Name == "거래내역_테스트용_500행.xlsx"]
    if len(matches) != 1:
        raise RuntimeError("Expected one open demo workbook")
    book = matches[0]
    sheet = book.ActiveSheet
    used = sheet.UsedRange
    rows, cols = used.Rows.Count, used.Columns.Count
    if rows * cols > 20000:
        raise RuntimeError("Unexpected sheet size")
    values = used.Value
    formulas = used.Formula2
    formats = [[used.Cells(r, c).NumberFormat for c in range(1, cols + 1)] for r in range(1, rows + 1)]
    state = {
        "time": time.time(), "workbook": book.Name, "fullName": book.FullName,
        "saved": bool(book.Saved), "sheet": sheet.Name,
        "sheets": [s.Name for s in book.Worksheets], "address": used.Address,
        "values": values, "formulas": formulas, "formats": formats,
    }
    write_json(ROOT / f"excel-{label}.json", state)
    print(json.dumps({k: v for k, v in state.items() if k not in {"values", "formulas", "formats"}}, ensure_ascii=False))
    if rows <= 30:
        print(json.dumps(values, ensure_ascii=False, default=str))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("command", choices=["snapshot", "excel-state"])
    parser.add_argument("label")
    args = parser.parse_args()
    {"snapshot": snapshot, "excel-state": excel_state}[args.command](args.label)
