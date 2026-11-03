"""Record the desktop and drive Celly for the demo shoot.

record  : run ffmpeg (30fps, real cursor) until stop.flag appears, then quit cleanly
shot X  : save one desktop PNG to shots/X.png
move/click X Y, type TEXT, key NAME, windows
"""
import ctypes
from ctypes import wintypes
import json
from pathlib import Path
import subprocess
import sys
import time

import imageio_ffmpeg

ROOT = Path(__file__).resolve().parent
FF = imageio_ffmpeg.get_ffmpeg_exe()
U = ctypes.windll.user32
U.SetProcessDPIAware()


def record():
    out = ROOT / "raw.mkv"
    p = subprocess.Popen([FF, "-hide_banner", "-loglevel", "error", "-y", "-f", "gdigrab", "-framerate", "30",
                          "-draw_mouse", "1", "-i", "desktop", "-c:v", "libx264", "-preset", "ultrafast", "-crf", "16",
                          "-pix_fmt", "yuv444p", str(out)], stdin=subprocess.PIPE)
    (ROOT / "record.json").write_text(json.dumps({"start": time.time()}), encoding="utf-8")
    flag = ROOT / "stop.flag"
    while not flag.exists() and p.poll() is None:
        time.sleep(.2)
    if p.poll() is None:
        p.stdin.write(b"q")
        p.stdin.flush()
        p.wait(30)
    flag.unlink(missing_ok=True)
    print("recorded", out)


def shot(name):
    (ROOT / "shots").mkdir(exist_ok=True)
    subprocess.run([FF, "-hide_banner", "-loglevel", "error", "-y", "-f", "gdigrab", "-framerate", "1", "-draw_mouse", "1",
                    "-i", "desktop", "-frames:v", "1", str(ROOT / "shots" / f"{name}.png")], check=True, timeout=20)
    log("shot", name)


def log(*event):
    with (ROOT / "events.jsonl").open("a", encoding="utf-8") as f:
        f.write(json.dumps({"t": time.time(), "event": event}, ensure_ascii=False) + "\n")


def move(x, y, seconds=.6):
    p = wintypes.POINT()
    U.GetCursorPos(ctypes.byref(p))
    steps = max(1, int(seconds * 60))
    for i in range(1, steps + 1):
        k = i / steps
        k = k * k * (3 - 2 * k)  # ease in-out
        U.SetCursorPos(round(p.x + (x - p.x) * k), round(p.y + (y - p.y) * k))
        time.sleep(seconds / steps)


def click(x, y):
    move(x, y)
    time.sleep(.15)
    U.mouse_event(2, 0, 0, 0, 0)
    time.sleep(.05)
    U.mouse_event(4, 0, 0, 0, 0)
    log("click", x, y)


class KI(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]


class INPUT(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [("ki", KI), ("pad", ctypes.c_byte * 32)]
    _anonymous_ = ("u",)
    _fields_ = [("type", wintypes.DWORD), ("u", _U)]


def send(vk=0, scan=0, flags=0):
    i = INPUT(type=1)
    i.ki = KI(vk, scan, flags, 0, 0)
    U.SendInput(1, ctypes.byref(i), ctypes.sizeof(i))


def type_text(text, delay=.075):
    log("type", text)
    for ch in text:
        send(scan=ord(ch), flags=4)  # KEYEVENTF_UNICODE
        send(scan=ord(ch), flags=4 | 2)
        time.sleep(delay)


KEYS = {"enter": 0x0D, "ctrl+s": (0x11, 0x53), "ctrl+a": (0x11, 0x41), "delete": 0x2E}


def key(name):
    vk = KEYS[name]
    vks = vk if isinstance(vk, tuple) else (vk,)
    for v in vks:
        send(vk=v)
    for v in reversed(vks):
        send(vk=v, flags=2)
    log("key", name)


def windows():
    found = {}

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def each(h, _):
        if U.IsWindowVisible(h):
            buf = ctypes.create_unicode_buffer(256)
            U.GetWindowTextW(h, buf, 256)
            r = wintypes.RECT()
            U.GetWindowRect(h, ctypes.byref(r))
            if buf.value in ("셀리",) or "거래내역_테스트용_500행" in buf.value:
                found[buf.value] = [h, r.left, r.top, r.right, r.bottom]
        return True

    U.EnumWindows(each, 0)
    print(json.dumps(found, ensure_ascii=False))


if __name__ == "__main__":
    cmd, *args = sys.argv[1:]
    if cmd == "record":
        record()
    elif cmd == "shot":
        shot(args[0])
    elif cmd == "move":
        move(int(args[0]), int(args[1]))
    elif cmd == "click":
        click(int(args[0]), int(args[1]))
    elif cmd == "type":
        type_text(args[0])
    elif cmd == "key":
        key(args[0])
    elif cmd == "windows":
        windows()
    elif cmd == "front":
        for h in map(int, args):
            U.keybd_event(0x12, 0, 0, 0)  # Alt tap lets SetForegroundWindow succeed
            U.keybd_event(0x12, 0, 2, 0)
            U.ShowWindow(h, 9)
            U.SetForegroundWindow(h)
            time.sleep(.4)
