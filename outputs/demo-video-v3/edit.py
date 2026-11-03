"""Cut raw.mkv into the ~75s Celly demo: follow-zoom camera, step pill, sped-up waits."""
from pathlib import Path
import subprocess

import imageio_ffmpeg
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parent
FONT = ROOT.parents[1] / "app" / "fonts" / "Pretendard-Bold.ttf"
FF = imageio_ffmpeg.get_ffmpeg_exe()
W, H, FPS = 1920, 1080, 30
MUSIC = ROOT.parents[1] / "MP_시작하는 마음.mp3"
MUSIC_START, MUSIC_GAIN_DB = 35, -5  # user asked to start at 0:35; -5dB keeps it under the presenter's voice
BG = (238, 241, 240)
PILL, INK, KEY = (228, 233, 230), (28, 40, 37), (47, 125, 225)

# Camera rects on the stage (desktop rows 0..1020 placed at y=30): (x, y, width), height = width*9/16.
FULL = (0, 0, 1920)
CEL_IN = (880, 495, 1040)     # Celly input + lower chat, right edge of Excel
CEL_TOP = (640, 90, 1280)     # Celly conversation
TABLE = (0, 170, 1100)        # Excel result table

S1, S2, S3 = ("1", "집계표 만들기"), ("2", "합계 추가"), ("3", "천원 단위 표시")
# Result captions: segment index -> (from, to output seconds into segment, [(text, key?)]).
CAPTIONS = {
    7: (1.2, 3.9, [("새 시트에 ", 0), ("회사별·연도별 집계표", 1), (" 완성", 0)]),
    10: (2.0, 6.8, [("아래 ", 0), ("연도별 합계", 1), (", 오른쪽 ", 0), ("회사별 합계", 1)]),
    13: (0.6, 7.0, [("표시만 천원 단위", 1), (" · 실제 값은 그대로", 0)], 760),  # full view: over empty rows, not the sheet tabs
}
# Cells that changed, in raw desktop pixels (sheet at 100%): segment index -> [(from, to, box)].
ROW13 = (46, 521, 790, 545)       # A13:F13 yearly totals
COL_F = (664, 245, 790, 545)      # F1:F13 company totals
AMOUNTS = (171, 268, 790, 545)    # B2:F13 display format changed
HIGHLIGHTS = {
    10: [(2.0, 6.8, ROW13), (2.5, 6.8, COL_F)],
    13: [(0.6, 7.0, AMOUNTS)],
}
# Celly '작업 완료' reply bubbles (raw px): segment index -> (from, to, box). Glow + one light sweep + sparkles.
DONE = {
    5: (0.6, 3.8, (1405, 223, 1872, 524)),
    10: (0.1, 2.2, (1398, 529, 1864, 754)),
    13: (0.1, 3.2, (1397, 551, 1864, 754)),
}
# (raw start, raw end, speed, step, camera keyframes [(output seconds into segment, rect)])
SEGMENTS = [
    (7.0, 9.5, 1, S1, [(0, FULL)]),
    (9.5, 17.5, 1, S1, [(0, FULL), (1.2, CEL_IN)]),
    (18.0, 19.5, 1, S1, [(0, CEL_IN)]),
    (28.3, 31.3, 1, S1, [(0, CEL_IN)]),
    (31.3, 41.7, 3, S1, [(0, CEL_IN)]),
    (41.7, 46.0, 1, S1, [(0, CEL_IN), (1.0, CEL_TOP)]),
    (62.5, 63.6, 1, S1, [(0, CEL_TOP), (0.8, FULL)]),
    (103.6, 110.4, 1, S1, [(0, FULL), (0.5, FULL), (1.5, TABLE), (3.2, TABLE), (4.0, CEL_IN)]),
    (124.8, 134.2, 1, S2, [(0, CEL_IN)]),
    (134.2, 145.0, 3, S2, [(0, CEL_IN)]),
    (145.0, 152.0, 1, S2, [(0, CEL_IN), (1.2, CEL_IN), (2.2, TABLE)]),
    # Step 3 stays on the full workspace (no zoom), as requested.
    (170.8, 179.4, 1, S3, [(0, TABLE), (1.0, FULL)]),
    (179.4, 182.75, 1, S3, [(0, FULL)]),
    (182.75, 192.0, 1, S3, [(0, FULL)]),
]
FADE = 8  # frames of crossfade where the raw footage jumps


def font(size):
    return ImageFont.truetype(str(FONT), size)


def stage_mask():
    """Keep the two app windows; matte the desktop/other apps behind them."""
    m = Image.new("L", (W, 1020), 0)
    d = ImageDraw.Draw(m)
    d.rectangle((9, 12, 1359, 990), fill=255)      # Excel
    d.rectangle((1366, 10, 1911, 1013), fill=255)  # Celly
    return m


MASK = stage_mask()


def stage(raw):
    canvas = Image.new("RGB", (W, H), BG)
    canvas.paste(raw.crop((0, 0, W, 1020)), (0, 30), MASK)
    return canvas


def highlight(canvas, boxes, t):
    """Fade in a translucent blue box over each changed range (drawn on the stage, so it zooms with the camera)."""
    overlay = Image.new("RGBA", canvas.size)
    d = ImageDraw.Draw(overlay)
    for t0, t1, (x0, y0, x1, y1) in boxes:
        a = min(1, max(0, (t - t0) / .35), max(0, (t1 - t) / .25))
        if a > 0:
            d.rounded_rectangle((x0 - 3, y0 + 27, x1 + 3, y1 + 33), radius=6,
                                fill=KEY + (round(46 * a),), outline=KEY + (round(255 * a),), width=4)
    canvas.paste(overlay, (0, 0), overlay)


def star(d, cx, cy, r, color):
    k = r * .28
    d.polygon([(cx, cy - r), (cx + k, cy - k), (cx + r, cy), (cx + k, cy + k),
               (cx, cy + r), (cx - k, cy + k), (cx - r, cy), (cx - k, cy - k)], fill=color)


def done_glow(canvas, spec, t):
    """Soft blue glow around the reply bubble, one diagonal light sweep, and three twinkles."""
    t0, t1, (x0, y0, x1, y1) = spec
    a = min(1, max(0, (t - t0) / .3), max(0, (t1 - t) / .4))
    if a <= 0:
        return
    m, w, h = 40, x1 - x0, y1 - y0
    size = (w + 2 * m, h + 2 * m)
    # Blur alpha masks only, then fill with a solid colour (blurring RGBA directly leaves grey fringes).
    glow = Image.new("L", size)
    ImageDraw.Draw(glow).rounded_rectangle((m, m, m + w, m + h), radius=18, outline=round(170 * a), width=12)
    halo = Image.new("RGBA", size, KEY + (0,))
    halo.putalpha(glow.filter(ImageFilter.GaussianBlur(10)))
    ImageDraw.Draw(halo).rounded_rectangle((m, m, m + w, m + h), radius=18, outline=KEY + (round(255 * a),), width=3)
    p = (t - t0 - .15) / .8  # sweep progress
    if 0 < p < 1:
        streak = Image.new("L", size)
        bx, bw = m - 120 + p * (w + 240), 70
        ImageDraw.Draw(streak).polygon([(bx, m), (bx + bw, m), (bx + bw - 60, m + h), (bx - 60, m + h)], fill=150)
        clip = Image.new("L", size)
        ImageDraw.Draw(clip).rounded_rectangle((m + 3, m + 3, m + w - 3, m + h - 3), radius=16, fill=255)
        band = Image.new("RGBA", size, (190, 218, 255, 0))
        band.putalpha(Image.fromarray(np.minimum(np.asarray(streak.filter(ImageFilter.GaussianBlur(14))), np.asarray(clip))))
        halo.alpha_composite(band)
    d = ImageDraw.Draw(halo)
    for i, (sx, sy, r) in enumerate([(m + w - 6, m + 4, 16), (m + w + 14, m + 34, 9), (m + 10, m + h - 4, 11)]):
        k = (t - t0 - .25 - i * .18) / .7
        if 0 < k < 1:
            star(d, sx, sy, r * np.sin(np.pi * k), KEY + (round(255 * a),))
    canvas.paste(halo, (x0 - m, y0 + 30 - m), halo)


def ease(k):
    return k * k * (3 - 2 * k)


def camera(keys, t):
    if t <= keys[0][0]:
        return keys[0][1]
    for (t0, a), (t1, b) in zip(keys, keys[1:]):
        if t <= t1:
            k = ease((t - t0) / (t1 - t0))
            return tuple(a[i] + (b[i] - a[i]) * k for i in range(3))
    return keys[-1][1]


def pill(canvas, xy, parts, size=30, anchor="left"):
    f = font(size)
    width = sum(f.getlength(t) for t, _ in parts)
    x, y = xy
    if anchor == "right":
        x -= width + 56
    elif anchor == "bottom":  # centred caption; above the bottom edge unless a top is given
        x, y = x - (width + 56) / 2, H - 48 - (size + 30) if y is None else y
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle((x, y, x + width + 56, y + size + 30), radius=14, fill=PILL)
    cx = x + 28
    for text, color in parts:
        d.text((cx, y + 15 + size / 2), text, font=f, fill=color, anchor="lm")
        cx += f.getlength(text)


def frames(start, end, speed):
    """Yield raw frames from start..end, keeping every `speed`-th one."""
    p = subprocess.Popen([FF, "-loglevel", "error", "-ss", f"{start:.3f}", "-i", str(ROOT / "raw.mkv"),
                          "-t", f"{end - start:.3f}", "-f", "rawvideo", "-pix_fmt", "rgb24", "pipe:1"],
                         stdout=subprocess.PIPE)
    n = 0
    count = round((end - start) * FPS)
    while n < count:
        buf = p.stdout.read(W * H * 3)
        if len(buf) < W * H * 3:
            break
        if n % speed == 0:
            yield Image.frombuffer("RGB", (W, H), buf)
        n += 1
    p.stdout.close()
    p.wait()


def render(out):
    enc = subprocess.Popen([FF, "-hide_banner", "-loglevel", "warning", "-y", "-f", "rawvideo", "-pix_fmt", "rgb24",
                            "-s", f"{W}x{H}", "-r", str(FPS), "-i", "pipe:0", "-c:v", "libx264", "-preset", "slow",
                            "-crf", "16", "-pix_fmt", "yuv420p", "-movflags", "+faststart", str(out)],
                           stdin=subprocess.PIPE)
    last, total, prev_end = None, 0, None
    for si, (start, end, speed, step, keys) in enumerate(SEGMENTS):
        jump = prev_end is not None and abs(start - prev_end) > .05
        prev_end = end
        for i, raw in enumerate(frames(start, end, speed)):
            t = i / FPS
            x, y, w = camera(keys, t)
            h = w * 9 / 16
            st = stage(raw)
            if si in HIGHLIGHTS:
                highlight(st, HIGHLIGHTS[si], t)
            if si in DONE:
                done_glow(st, DONE[si], t)
            frame = st.crop((round(x), round(y), round(x + w), round(y + h))).resize((W, H), Image.LANCZOS)
            pill(frame, (36, 30), [(step[0] + "  ", KEY), (step[1], INK)])
            if speed > 1:
                pill(frame, (W - 36, 30), [("처리 대기 ", INK), (f"{speed}배속", KEY)], size=26, anchor="right")
            if si in CAPTIONS:
                c0, c1, parts, *top = CAPTIONS[si]
                alpha = min(1, max(0, (t - c0) / .25), max(0, (c1 - t) / .25))
                if alpha > 0:
                    layer = frame.copy()
                    pill(layer, (W / 2, top[0] if top else None), [(text, KEY if key else INK) for text, key in parts],
                         size=40, anchor="bottom")
                    frame = Image.blend(frame, layer, alpha)
            if jump and i < FADE and last is not None:
                frame = Image.blend(last, frame, (i + 1) / (FADE + 1))
            enc.stdin.write(frame.tobytes())
            last = frame
            total += 1
        print(f"segment {si:2d} raw {start:6.2f}-{end:6.2f} x{speed} -> out {total / FPS:5.1f}s", flush=True)
    enc.stdin.close()
    if enc.wait():
        raise RuntimeError("ffmpeg failed")
    print(f"done: {total / FPS:.1f}s")
    return total / FPS


def add_music(silent, out, seconds):
    """Mux the music from MUSIC_START: 0.5s fade-in (mid-song start), 3s fade-out at the end."""
    subprocess.run([FF, "-hide_banner", "-loglevel", "error", "-y", "-i", str(silent),
                    "-ss", str(MUSIC_START), "-t", f"{seconds:.3f}", "-i", str(MUSIC),
                    "-map", "0:v", "-map", "1:a", "-c:v", "copy", "-c:a", "aac", "-b:a", "192k",
                    "-af", f"volume={MUSIC_GAIN_DB}dB,afade=t=in:d=0.5,afade=t=out:st={seconds - 3:.3f}:d=3",
                    "-shortest", "-movflags", "+faststart", str(out)], check=True)
    print("music added:", out.name)


if __name__ == "__main__":
    silent = ROOT / "셀리_시연영상_무음.mp4"
    add_music(silent, ROOT / "셀리_시연영상.mp4", render(silent))
