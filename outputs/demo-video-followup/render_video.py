"""Render a new 1080p Celly film exclusively from this take's native captures."""
from __future__ import annotations

import argparse
from functools import lru_cache
import json
import math
from pathlib import Path
import subprocess
import sys
import wave

import imageio_ffmpeg
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent
PROJECT = ROOT.parents[1]
W, H, FPS = 1920, 1080, 30
FONT_DIR = PROJECT / "app" / "fonts"
GROUND = "#F3F5F4"
SURFACE = "#E7ECE9"
INK = "#1C2825"
MUTED = "#52645C"
ACCENT = "#087E70"
BLUE = ACCENT
ICON = Image.open(PROJECT / "app" / "assets" / "celly-icon.png").convert("RGBA")
ROLES = {}
SHOTS = []
CAPTURE_LOG = {}


@lru_cache(maxsize=80)
def font(size, bold=False):
    return ImageFont.truetype(str(FONT_DIR / ("Pretendard-Bold.ttf" if bold else "Pretendard-Regular.ttf")), size)


def center(draw, xy, text, size, fill, bold=False):
    draw.text(xy, text, font=font(size, bold), fill=fill, anchor="mt")


def brand(canvas, x=38, y=20, size=64):
    icon = ICON.resize((size - 8, size - 8), Image.Resampling.LANCZOS)
    canvas.paste(icon, (x + 4, y + 4), icon)


def base(phase):
    canvas = Image.new("RGB", (W, H), GROUND)
    brand(canvas,x=98,y=18,size=44)
    d = ImageDraw.Draw(canvas)
    d.text((152,24),"셀리",font=font(34,True),fill=INK)
    d.text((284,32),"열린 Excel에서 계속",font=font(22),fill=MUTED)
    d.text((1822,32),phase,font=font(22),fill=INK,anchor="rt")
    return canvas


PILL = "#E4E9E6"
KEY = "#2F7DE1"


def caption(canvas, text):
    """Centered rounded pill; [bracketed] words are drawn in the key colour."""
    parts = [(chunk, i % 2 == 1) for i, chunk in enumerate(text.replace("\n", " ").replace("]", "[").split("[")) if chunk]
    size = 32
    while size > 22 and sum(font(size, True).getlength(t) for t, _ in parts) > 1500:
        size -= 1
    width = sum(font(size, True).getlength(t) for t, _ in parts)
    x, cy = 960 - width / 2, 1036
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle((x - 30, cy - 30, x + width + 30, cy + 30), radius=14, fill=PILL)
    for chunk, key in parts:
        d.text((x, cy), chunk, font=font(size, True), fill=KEY if key else INK, anchor="lm")
        x += font(size, True).getlength(chunk)


def tag(canvas, text, x=142, y=145):
    d = ImageDraw.Draw(canvas)
    d.text((x,y),text,font=font(22),fill=MUTED)


@lru_cache(maxsize=80)
def capture(role, window):
    return Image.open(ROOT / "captures" / f"{ROLES[role]}-{window}.png").convert("RGB")


def screenshot(canvas, role, window, box, crop=None):
    source = capture(role, window)
    if crop:
        source = source.crop(tuple(crop))
    x, y, width, height = box
    scale = min(width / source.width, height / source.height)
    nw, nh = round(source.width * scale), round(source.height * scale)
    x, y = round(x + (width - nw) / 2), round(y + (height - nh) / 2)
    image = source.resize((nw, nh), Image.Resampling.LANCZOS)
    canvas.paste(image,(x,y))
    return x, y, scale


@lru_cache(maxsize=80)
def screen_layer(role, window, box, crop):
    layer = Image.new("RGB", (W, H), GROUND)
    placement = screenshot(layer, role, window, box, crop)
    return layer, placement


@lru_cache(maxsize=200)
def live_image(name, frame):
    return Image.open(ROOT/"motion"/name/f"{frame+1:03d}.png").convert("RGB")


def live_panel(canvas, name, frame, box):
    source = live_image(name, frame)
    x,y,width,height=box
    scale=min(width/source.width,height/source.height)
    size=(round(source.width*scale),round(source.height*scale))
    x,y=round(x+(width-size[0])/2),round(y+(height-size[1])/2)
    image=source.resize(size,Image.Resampling.LANCZOS)
    canvas.paste(image,(x,y))


def workspace_only(source):
    """Keep both complete application windows and matte out the surrounding desktop."""
    workspace=Image.new("RGB",source.size,GROUND)
    for rect in [(9,12,1360,990),(1383,10,1910,1012)]:
        workspace.paste(source.crop(rect),(rect[0],rect[1]))
    return workspace


@lru_cache(maxsize=12)
def raw_desktop_still(role):
    return Image.open(ROOT/"captures"/f"{ROLES[role]}-desktop.png").convert("RGB")


@lru_cache(maxsize=4)
def raw_desktop_frame(name,frame):
    return Image.open(ROOT/"wide-motion"/name/f"{frame+1:03d}.png").convert("RGB")


@lru_cache(maxsize=12)
def desktop_still(role):
    return workspace_only(raw_desktop_still(role)).resize((1724,916),Image.Resampling.LANCZOS)


@lru_cache(maxsize=4)
def desktop_frame(name,frame):
    return workspace_only(raw_desktop_frame(name,frame)).resize((1724,916),Image.Resampling.LANCZOS)


def loupe(canvas,source,spec,local,seconds):
    """Magnify genuine pixels locally while the complete workspace stays fixed."""
    enter=min(1,max(0,local/.32))
    leave=min(1,max(0,(seconds-local)/.24))
    strength=(1-(1-enter)**4)*leave
    if strength<=0:
        return
    crop=spec["source"]
    # A source outline connects each loupe to its actual UI location.
    source_box=(98+crop[0]*1724/1920-4,76+crop[1]*916/1020-4,
                98+crop[2]*1724/1920+4,76+crop[3]*916/1020+4)
    marker=Image.new("RGBA",canvas.size)
    md=ImageDraw.Draw(marker)
    md.rounded_rectangle(source_box,radius=6,fill=(8,126,112,round(26*strength)),outline=(8,126,112,round(255*strength)),width=3)
    canvas.paste(marker,(0,0),marker)
    if spec["shape"]=="outline":
        return
    zoom=spec["zoom"]
    width,height=round((crop[2]-crop[0])*zoom),round((crop[3]-crop[1])*zoom)
    if spec.get("capture"):
        # Cursor-free raw frame just before send, so the magnified text is unobstructed.
        source=Image.open(ROOT/"captures"/f"{spec['capture']}-desktop.png").convert("RGB")
    image=source.crop(tuple(round(v) for v in crop)).resize((width,height),Image.Resampling.LANCZOS)
    padding=8
    layer=Image.new("RGBA",(width+padding*2,height+padding*2),(255,255,255,0))
    mask=Image.new("L",layer.size)
    ImageDraw.Draw(mask).rounded_rectangle((0,0,layer.width-1,layer.height-1),radius=10,fill=255)
    layer.paste("white",(0,0,*layer.size))
    layer.paste(image,(padding,padding))
    ImageDraw.Draw(layer).rounded_rectangle((0,0,layer.width-1,layer.height-1),radius=10,outline=ACCENT,width=3)
    layer.putalpha(mask)
    x,y=spec["position"]
    y+=round(12*(1-strength))
    shadow=Image.new("RGBA",(layer.width+48,layer.height+48))
    shadow.paste((22,38,30,48),(24,27),mask)
    shadow=shadow.filter(ImageFilter.GaussianBlur(14))
    shadow.putalpha(shadow.getchannel("A").point(lambda a:round(a*strength)))
    canvas.paste(shadow,(x-24,y-24),shadow)
    layer.putalpha(mask.point(lambda a:round(a*strength)))
    canvas.paste(layer,(x,y),layer)
    label=Image.new("RGBA",(950,40))
    ImageDraw.Draw(label).text((0,0),spec["title"],font=font(24,True),fill=ACCENT)
    label.putalpha(label.getchannel("A").point(lambda a:round(a*strength)))
    canvas.paste(label,(x,y-38),label)


def draw_desktop_shot(shot,local):
    canvas=base(shot["phase"])
    if shot.get("desktopLive"):
        frame=min(round(shot["seconds"]*12)-1,int(local*12))
        desktop=desktop_frame(shot["desktopLive"],frame)
        raw=raw_desktop_frame(shot["desktopLive"],frame)
    else:
        desktop=desktop_still(shot["desktopRole"])
        raw=raw_desktop_still(shot["desktopRole"])
    canvas.paste(desktop,(98,76))
    if shot.get("lens"):
        loupe(canvas,raw,shot["lens"],local,shot["seconds"])
    caption(canvas,shot["caption"])
    if shot.get("note"):
        ImageDraw.Draw(canvas).text((1822,1048),shot["note"],font=font(22),fill=MUTED,anchor="rt")
    return canvas


def callout(canvas, box, amount=1):
    overlay = Image.new("RGBA", (W, H))
    d = ImageDraw.Draw(overlay)
    d.rounded_rectangle(box, radius=12, outline=(45, 111, 237, round(230 * amount)), width=5)
    canvas.paste(overlay, (0, 0), overlay)


def observed_pointer(canvas, shot, placement, local):
    pointer = shot.get("pointer")
    if not pointer:
        return
    event = CAPTURE_LOG[ROLES[pointer["role"]]]
    window = pointer["window"]
    native = event["windows"][window]
    panel, (x, y, scale) = placement[window]
    crop = panel.get("crop", [0, 0])
    px = x + (event["pointer"][0] - native[0] - crop[0]) * scale
    py = y + (event["pointer"][1] - native[1] - crop[1]) * scale
    overlay = Image.new("RGBA", (W, H))
    d = ImageDraw.Draw(overlay)
    elapsed = local - .4
    if 0 <= elapsed < .6:
        radius = 15 + elapsed * 70
        d.ellipse((px-radius,py-radius,px+radius,py+radius),outline=(48,119,244,round(220*(1-elapsed/.6))),width=5)
    points=[(px,py),(px+8,py+34),(px+16,py+23),(px+26,py+35),(px+32,py+29),(px+21,py+18),(px+34,py+14)]
    d.polygon(points,fill="white",outline=INK,width=2)
    canvas.paste(overlay,(0,0),overlay)


def card(kind, title, subtitle="", local=0):
    canvas = base("")
    d = ImageDraw.Draw(canvas)
    if kind == "opening" and local < 3:
        center(d, (960, 350), "표 만들고, 합계 넣고,", 72, INK, True)
        center(d, (960, 452), "단위 바꾸고.", 84, INK, True)
        center(d, (960, 618), "반복되는 Excel 작업을 이어서 다듬는 방법", 31, MUTED)
    elif kind == "opening":
        icon = ICON.resize((132, 132), Image.Resampling.LANCZOS)
        canvas.paste(icon, (894, 242), icon)
        center(d, (960, 412), "셀리", 100, INK, True)
        center(d, (960, 564), subtitle, 43, INK, True)
        center(d, (960, 657), "집계표 → 합계 추가 → 천원 단위 표시", 30, MUTED)
    elif kind == "ending":
        icon = ICON.resize((112, 112), Image.Resampling.LANCZOS)
        canvas.paste(icon, (904, 218), icon)
        center(d, (960, 379), title, 58, INK, True)
        center(d, (960, 501), "지금 열린 Excel에서,", 42, INK)
        center(d, (960, 567), "바꾸고 싶은 일을 입력해 보세요.", 48, INK, True)
        center(d, (960, 708), "셀리  CELLY", 29, MUTED, True)
    else:
        _, _, text = title.partition(". ")
        center(d, (960, 473), text, 68, INK, True)
        center(d, (960, 608), "파일 첨부 없이, 열린 Excel에서 계속", 31, MUTED)
    return canvas


def draw_shot(shot, local):
    if shot["kind"] in {"opening", "chapter", "ending"}:
        return card(shot["kind"], shot["title"], shot.get("subtitle", ""), local)
    if shot["kind"]=="desktop":
        return draw_desktop_shot(shot,local)
    canvas = base(shot["phase"])
    placement = {}
    for panel in shot["panels"]:
        if shot.get("live"):
            live_panel(canvas,shot["live"],min(59,int(local*12)),panel["box"])
        else:
            layer, position=screen_layer(panel["role"],panel["window"],tuple(panel["box"]),tuple(panel["crop"]) if panel.get("crop") else None)
            # Only copy the bounded panel/shadow region, retaining the header.
            x,y,w,h=panel["box"]
            region=(max(0,x-40),max(104,y-35),min(W,x+w+40),min(934,y+h+40))
            canvas.paste(layer.crop(region),(region[0],region[1]))
            placement[panel["window"]] = (panel,position)
    if shot.get("label"):
        tag(canvas, shot["label"])
    if shot.get("note"):
        d = ImageDraw.Draw(canvas)
        center(d, (960, 880), shot["note"], 23, MUTED)
    if shot.get("unit"):
        tag(canvas, shot["unit"], x=shot.get("unitX",1500), y=shot.get("unitY",145))
    for box in shot.get("highlights", []):
        if local < shot["seconds"] - 3:
            callout(canvas, box)
    if shot.get("shortcut"):
        center(ImageDraw.Draw(canvas), (1610, 157), shot["shortcut"], 27, INK, True)
    observed_pointer(canvas, shot, placement, local)
    caption(canvas, shot["caption"])
    return canvas


def build_timeline():
    config = json.loads((ROOT / "film-edit.json").read_text(encoding="utf-8"))
    global ROLES, SHOTS, CAPTURE_LOG
    ROLES = config["roles"]
    SHOTS = config["shots"]
    CAPTURE_LOG = {item["label"]: item for item in json.loads((ROOT/"capture.json").read_text(encoding="utf-8"))}
    position = 0
    for shot in SHOTS:
        shot["start"] = position
        position += shot["seconds"]
        shot["end"] = position
    (ROOT / "edit-timeline.json").write_text(json.dumps(SHOTS, ensure_ascii=False, indent=2), encoding="utf-8")
    return position


def soundtrack(duration):
    rate, bpm = 48000, 112
    beat = 60 / bpm
    samples = np.zeros((math.ceil(duration * rate), 2), dtype=np.float64)

    def add(signal, start, gain, pan=0):
        index = round(start * rate)
        if index < 0 or index >= len(samples):
            return
        signal = signal[:len(samples)-index] * gain
        samples[index:index+len(signal), 0] += signal * math.sqrt((1-pan)/2)
        samples[index:index+len(signal), 1] += signal * math.sqrt((1+pan)/2)

    def tone(midi, length, style="key"):
        t = np.arange(round(rate * length)) / rate
        freq = 440 * 2 ** ((midi - 69) / 12)
        attack = np.minimum(t / 0.012, 1)
        release = np.minimum((length - t) / 0.08, 1)
        decay = np.exp(-t * (4.8 if style == "key" else 1.6))
        wave = np.sin(2*np.pi*freq*t) + .22*np.sin(4*np.pi*freq*t) + .05*np.sin(6*np.pi*freq*t)
        return wave * attack * release * decay

    rng = np.random.default_rng(20261103)
    chords = [(48, 55, 60, 64), (43, 50, 55, 59), (45, 52, 57, 60), (41, 48, 53, 57)]
    melody = [76, None, 79, 76, 74, None, 71, None, 72, None, 76, 79, 77, 76, 72, None]
    for step in range(math.ceil(duration / beat)):
        t0 = step * beat
        chord = chords[(step // 4) % 4]
        if step % 2 == 0:
            for note in chord[1:]:
                add(tone(note, beat * 1.7), t0, .042, -.2)
            add(tone(chord[0]-12, beat * .8, "bass"), t0, .10)
        kick_t = np.arange(int(rate * .16)) / rate
        kick = np.sin(2*np.pi*(55*kick_t + 4.5*(1-np.exp(-kick_t*30)))) * np.exp(-kick_t*27)
        add(kick, t0, .10)
        if step % 4 in (1, 3):
            nt = np.arange(int(rate*.11)) / rate
            snare = rng.normal(size=len(nt)) * np.exp(-nt*45)
            add(snare, t0, .019, .1)
        for eighth in (0, .5):
            nt = np.arange(int(rate*.045)) / rate
            noise = rng.normal(size=len(nt))
            highpass = noise - np.convolve(noise, np.ones(9)/9, mode="same")
            add(highpass*np.exp(-nt*95), t0 + eighth*beat, .011, .35)
        m = melody[step % len(melody)]
        if m is not None:
            add(tone(m, beat*.8), t0 + .015, .026, .23)
            add(tone(m, beat*.8), t0 + beat*.75, .007, -.3)
    for shot in SHOTS:
        if shot.get("click"):
            nt = np.arange(int(rate*.07)) / rate
            click = np.sin(2*np.pi*1250*nt) * np.exp(-nt*90)
            add(click, shot["start"] + shot.get("clickAt",.4), .10)
    t = np.arange(len(samples)) / rate
    fade = np.minimum(t / 1.4, 1) * np.minimum((duration-t)/2.0, 1)
    samples *= fade[:, None]
    peak = np.abs(samples).max()
    if peak > .55:
        samples *= .55/peak
    with wave.open(str(ROOT/"soundtrack.wav"), "wb") as out:
        out.setnchannels(2)
        out.setsampwidth(2)
        out.setframerate(rate)
        out.writeframes((samples*32767).astype("<i2").tobytes())


def render():
    duration = build_timeline()
    soundtrack(duration)
    command = [imageio_ffmpeg.get_ffmpeg_exe(), "-hide_banner", "-loglevel", "warning", "-y",
               "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}", "-r", str(FPS), "-i", "pipe:0",
               "-i", str(ROOT/"soundtrack.wav"), "-c:v", "libx264", "-preset", "medium", "-crf", "14",
               "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", "-shortest",
               str(ROOT/"셀리_소개_완성본.mp4")]
    with (ROOT/"render.log").open("wb") as log:
        process = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=log, stderr=log)
        previous_frame=None
        for shot in SHOTS:
            count = round(shot["seconds"] * FPS)
            # Preserve the workspace geometry; animate only the local lens after 38s.
            static_frame = None
            for frame in range(count):
                local = frame/FPS
                if static_frame is None or shot["kind"] == "opening" or shot.get("highlights") or shot.get("pointer") or shot.get("live") or shot.get("desktopLive") or shot.get("lens"):
                    current = draw_shot(shot, local)
                    static_frame = current
                else:
                    current = static_frame
                if previous_frame is not None and local < .2 and shot["start"] < 38:
                    current=Image.blend(previous_frame,current,local/.2)
                process.stdin.write(current.tobytes())
            previous_frame=static_frame
            print(f"Rendered {shot['start']:5.1f}-{shot['end']:5.1f}: {shot.get('phase',shot.get('title'))}", flush=True)
        process.stdin.close()
        code = process.wait()
        if code:
            raise RuntimeError(f"FFmpeg returned {code}; see render.log")
    print(f"Video complete: {duration}s", flush=True)


def previews():
    build_timeline()
    (ROOT/"previews").mkdir(exist_ok=True)
    frames = []
    for i, shot in enumerate(SHOTS):
        local = min(shot["seconds"] / 2, shot["seconds"] - .5)
        im = draw_shot(shot, local)
        im.save(ROOT/"previews"/f"shot-{i:02d}.png")
        small = im.resize((480,270), Image.Resampling.LANCZOS)
        tile = Image.new("RGB", (480,300), "#F4F7FC")
        tile.paste(small,(0,0))
        ImageDraw.Draw(tile).text((10,275), f"{i:02d}  {shot['start']:g}s  {shot.get('phase',shot.get('title'))}",font=font(16),fill=INK)
        frames.append(tile)
    contact = Image.new("RGB", (1920, math.ceil(len(frames)/4)*300), "#E7EEF9")
    for i, frame in enumerate(frames):
        contact.paste(frame, ((i%4)*480,(i//4)*300))
    contact.save(ROOT/"contact-sheet.png")
    print("Previews complete")


def gif():
    build_timeline()
    config = json.loads((ROOT/"film-edit.json").read_text(encoding="utf-8"))
    frames = []
    representatives=[]
    for segment in config["gif"]:
        shot = dict(segment["composition"] if "composition" in segment else SHOTS[segment["shot"]])
        shot["caption"] = segment["caption"]
        shot["note"] = "실제 작업 화면 · 대기 일부 생략"
        frame = draw_shot(shot, 1).resize((1152,648), Image.Resampling.LANCZOS)
        representatives.append(frame)
        for _ in range(int(segment["seconds"]*8)):
            frames.append(frame)
    palette = Image.new("RGB",(1152,648*len(representatives)))
    for i, frame in enumerate(representatives):
        palette.paste(frame,(0,i*648))
    palette = palette.quantize(colors=240,method=Image.Quantize.MEDIANCUT)
    indexed=[im.quantize(palette=palette,dither=Image.Dither.NONE) for im in frames]
    indexed[0].save(ROOT/"셀리_10초_채팅용.gif",save_all=True,append_images=indexed[1:],duration=125,loop=0,optimize=True,disposal=2)
    print("GIF complete: 10s")


if __name__ == "__main__":
    parser=argparse.ArgumentParser()
    parser.add_argument("command",choices=["previews","render","gif"])
    args=parser.parse_args()
    {"previews":previews,"render":render,"gif":gif}[args.command]()
