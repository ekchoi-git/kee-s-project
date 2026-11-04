"""Render the approved Celly demo from captured, real app states."""

from __future__ import annotations

import math
import subprocess
import wave
from pathlib import Path

import imageio_ffmpeg
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont


ROOT = Path(__file__).parent
W, H, FPS, DURATION = 1920, 1080, 15, 60
GREEN = '#193f35'
DEEP = '#102d27'
CREAM = '#f8f5ec'
CORAL = '#c67851'
MUTED = '#a6b7ad'
FONT_DIR = Path(__file__).parents[2] / 'app' / 'fonts'
REG = str(FONT_DIR / 'Pretendard-Regular.ttf')
MED = str(FONT_DIR / 'Pretendard-Medium.ttf')
BOLD = str(FONT_DIR / 'Pretendard-Bold.ttf')

before_excel = Image.open(ROOT / 'excel_before.png').convert('RGB')
after_excel = Image.open(ROOT / 'excel_after.png').convert('RGB')
saved_excel = Image.open(ROOT / 'excel_saved.png').convert('RGB')
before_celly = Image.open(ROOT / 'celly_before.png').convert('RGB')
input_celly = Image.open(ROOT / 'celly_input.png').convert('RGB')
result_celly = Image.open(ROOT / 'celly_result.png').convert('RGB')
icon = Image.open(Path(__file__).parents[2] / 'app' / 'assets' / 'celly-icon.png').convert('RGBA')


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(BOLD if bold else REG, size)


def text_center(draw: ImageDraw.ImageDraw, xy, value, fnt, fill):
    box = draw.textbbox((0, 0), value, font=fnt)
    draw.text((xy[0] - (box[2] - box[0]) / 2, xy[1]), value, font=fnt, fill=fill)


def panel(canvas: Image.Image, im: Image.Image, box: tuple[int, int, int, int]):
    x, y, pw, ph = box
    layer = im.resize((pw, ph), Image.Resampling.LANCZOS)
    mask = Image.new('L', (pw, ph), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle((0, 0, pw - 1, ph - 1), radius=18, fill=255)
    shadow = Image.new('RGBA', (pw + 36, ph + 36), (0, 0, 0, 0))
    shd = ImageDraw.Draw(shadow)
    shd.rounded_rectangle((16, 16, pw + 16, ph + 16), radius=22, fill=(6, 31, 23, 95))
    shadow = shadow.filter(ImageFilter.GaussianBlur(16))
    canvas.paste(shadow, (x - 18, y - 10), shadow)
    canvas.paste(layer, (x, y), mask)


def base(t: float) -> Image.Image:
    canvas = Image.new('RGB', (W, H), CREAM)
    d = ImageDraw.Draw(canvas)
    d.rectangle((0, 0, W, 101), fill=DEEP)
    d.rounded_rectangle((42, 23, 103, 84), radius=14, fill='#ffffff')
    canvas.paste(icon.resize((54, 54), Image.Resampling.LANCZOS), (46, 26), icon.resize((54, 54), Image.Resampling.LANCZOS))
    d.text((125, 31), '셀리', font=font(43, True), fill='white')
    d.text((242, 45), 'Excel 작업을 말로 간편하게', font=font(25), fill='#c9d8d0')
    phase = '연결 준비' if t < 15 else '말로 찾기' if t < 31 else '결과 확인과 저장'
    d.rounded_rectangle((1580, 31, 1852, 75), radius=22, fill='#315a4c')
    text_center(d, (1716, 40), phase, font(24, True), 'white')
    return canvas


def caption(canvas: Image.Image, value: str):
    overlay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    d.rounded_rectangle((193, 934, 1727, 1020), radius=25, fill=(10, 30, 25, 228))
    text_center(d, (960, 954), value, font(36, True), '#ffffff')
    canvas.alpha_composite(overlay)


def highlight(canvas: Image.Image, box, alpha=180):
    overlay = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    d.rounded_rectangle(box, radius=15, outline=(198, 120, 81, alpha), width=7)
    canvas.alpha_composite(overlay)


def cursor(canvas: Image.Image, x: float, y: float, click: float = -1):
    ov = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    x, y = int(x), int(y)
    pts = [(x, y), (x + 8, y + 33), (x + 15, y + 22), (x + 26, y + 35),
           (x + 32, y + 29), (x + 19, y + 17), (x + 32, y + 13)]
    d.polygon([(px + 3, py + 4) for px, py in pts], fill=(0, 0, 0, 100))
    d.polygon(pts, fill=(255, 255, 255, 255), outline=(20, 42, 35, 255), width=3)
    if 0 <= click < 0.55:
        r = 17 + int(click * 80)
        a = int(210 * (1 - click / 0.55))
        d.ellipse((x-r, y-r, x+r, y+r), outline=(198, 120, 81, a), width=6)
    canvas.alpha_composite(ov)


def ease(v):
    v = max(0, min(1, v))
    return v*v*(3-2*v)


def pos(t, points):
    for (t0, x0, y0), (t1, x1, y1) in zip(points, points[1:]):
        if t <= t1:
            v = ease((t-t0)/(t1-t0))
            return x0+(x1-x0)*v, y0+(y1-y0)*v
    return points[-1][1:]


def scene_content(t: float) -> Image.Image:
    canvas = base(t).convert('RGBA')
    if t < 15:
        ex, ce = before_excel, before_celly
        caption_text = '열린 Excel 파일을 셀리와 연결합니다.'
    elif t < 31:
        ex = before_excel
        ce = before_celly if t < 21 else input_celly
        caption_text = '원하는 조건을 말로 입력하면, 셀리가 Excel에 바로 적용합니다.'
    elif t < 49:
        ex, ce = after_excel, result_celly
        caption_text = '적용 내용을 확인하고, Excel에서 바뀐 데이터를 살펴보세요.'
    else:
        ex, ce = saved_excel, result_celly
        caption_text = '확인한 결과는 Excel 파일로 저장하세요.'
    panel(canvas, ex, (72, 124, 1280, 884))
    panel(canvas, ce, (1417, 124, 422, 800))
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle((89, 148, 242, 187), radius=18, fill='#264e40')
    d.text((111, 153), 'Microsoft Excel', font=font(21, True), fill='white')
    d.rounded_rectangle((1434, 148, 1535, 187), radius=18, fill='#264e40')
    d.text((1460, 153), '셀리', font=font(21, True), fill='white')
    if t < 15:
        highlight(canvas, (1433, 755, 1816, 813) if t < 11 else (1433, 813, 1816, 849))
        x,y=pos(t,[(6,700,570),(9,1610,783),(12,1630,827),(15,1660,829)])
        click_t=9.4
    elif t < 31:
        highlight(canvas, (1434, 813, 1814, 916) if t < 28 else (1775, 865, 1826, 919))
        x,y=pos(t,[(17,1560,600),(21,1580,853),(27,1600,870),(29,1805,884),(31,1805,884)])
        click_t=29.1
    elif t < 49:
        highlight(canvas, (1445, 281, 1805, 482) if t < 40 else (628, 352, 836, 658))
        x,y=pos(t,[(33,1500,300),(38,1530,392),(43,760,420),(49,760,540)])
        click_t=37.5
    else:
        highlight(canvas, (235, 131, 283, 165) if t < 53 else (318, 135, 631, 164))
        x,y=pos(t,[(49,770,540),(51,255,147),(55,460,150)])
        click_t=51.0
    cursor(canvas,x,y,t-click_t)
    caption(canvas,caption_text)
    return canvas.convert('RGB')


def card(title: str, subtitle: str = '', chapter: bool = False) -> Image.Image:
    im=Image.new('RGB',(W,H),DEEP)
    d=ImageDraw.Draw(im)
    for i in range(11):
        x=1270+i*76
        d.line((x,0,x,H),fill='#23473b',width=2)
    for i in range(12):
        y=90+i*76
        d.line((1250,y,W,y),fill='#23473b',width=2)
    d.rounded_rectangle((170,287,350,467),radius=41,fill='white')
    ic=icon.resize((150,150),Image.Resampling.LANCZOS)
    im.paste(ic,(185,302),ic)
    if chapter:
        d.text((410,336),title,font=font(84,True),fill='#ffffff')
        d.rectangle((412,454,977,460),fill=CORAL)
    else:
        d.text((408,303),title,font=font(112,True),fill='#ffffff')
        d.text((417,463),subtitle,font=font(54),fill='#d4e1d8')
    d.text((175,923),'CELLY  ·  EXCEL ASSISTANT',font=font(26,True),fill=MUTED)
    return im


opening=card('셀리','Excel 작업을 말로 간편하게')
endcard=card('셀리','반복되는 Excel 작업, 셀리에게 말하세요.')
chapters=[card('1. 연결 준비',chapter=True),card('2. 말로 찾기',chapter=True),card('3. 결과 확인과 저장',chapter=True)]


def frame(t):
    if t < 4:
        im=opening.copy().convert('RGBA')
        caption(im,'Excel 작업, 이제 셀리에게 말해보세요.')
        return im.convert('RGB')
    if t < 6:
        return chapters[0]
    if t < 15:
        return scene_content(t)
    if t < 17:
        return chapters[1]
    if t < 31:
        return scene_content(t)
    if t < 33:
        return chapters[2]
    if t < 55:
        return scene_content(t)
    im=endcard.copy().convert('RGBA')
    caption(im,'셀리와 함께, Excel 작업을 더 간단하게.')
    return im.convert('RGB')


def make_audio():
    sr=48000
    n=sr*DURATION
    t=np.arange(n,dtype=np.float32)/sr
    audio=np.zeros(n,dtype=np.float32)
    # Quiet original underscore: four warm chords, soft amplitude pulses.
    chords=[(196,246.94,293.66),(174.61,220,261.63),(164.81,207.65,246.94),(196,246.94,293.66)]
    for i in range(4):
        lo,hi=i*15*sr,(i+1)*15*sr
        seg=t[lo:hi]
        env=np.minimum(1,(seg-i*15)/1.4)*np.minimum(1,((i+1)*15-seg)/1.4)
        wave_part=sum(np.sin(2*np.pi*f*seg) for f in chords[i])/3
        audio[lo:hi]+=0.055*wave_part*env*(0.75+0.25*np.sin(2*np.pi*0.16*seg))
    audio*=np.minimum(1,t/1.5)*np.minimum(1,(DURATION-t)/2)
    for sec in (9.4,21.0,29.1,37.5,51.0):
        lo=int(sec*sr)
        dur=int(.12*sr)
        k=np.arange(dur)/sr
        click=(np.sin(2*np.pi*840*k)+.35*np.sin(2*np.pi*1230*k))*np.exp(-36*k)*.115
        audio[lo:lo+dur]+=click.astype(np.float32)
    pcm=np.clip(audio,-1,1)
    stereo=np.stack([pcm,pcm],axis=1)
    with wave.open(str(ROOT/'soundtrack.wav'),'wb') as f:
        f.setnchannels(2);f.setsampwidth(2);f.setframerate(sr)
        f.writeframes((stereo*32767).astype('<i2').tobytes())


def render():
    make_audio()
    # Save representative frames for visual QA.
    for sec in (2,10,24,35,44,52,57):
        frame(sec).save(ROOT/f'preview_{sec:02}.jpg',quality=90)
    ff=imageio_ffmpeg.get_ffmpeg_exe()
    cmd=[ff,'-y','-f','rawvideo','-pixel_format','rgb24','-video_size',f'{W}x{H}',
         '-framerate',str(FPS),'-i','pipe:0','-i',str(ROOT/'soundtrack.wav'),
         '-c:v','libx264','-preset','veryfast','-crf','20','-pix_fmt','yuv420p',
         '-c:a','aac','-b:a','128k','-r',str(FPS),'-t',str(DURATION),
         '-movflags','+faststart',str(ROOT/'셀리_60초_소개.mp4')]
    process=subprocess.Popen(cmd,stdin=subprocess.PIPE,stdout=subprocess.DEVNULL,stderr=(ROOT/'ffmpeg.log').open('wb'))
    try:
        for i in range(DURATION*FPS):
            process.stdin.write(frame(i/FPS).tobytes())
            if i%150==0:
                print(f'{i/FPS:.0f}/{DURATION}s',flush=True)
    finally:
        process.stdin.close()
    result=process.wait()
    if result:
        raise RuntimeError(f'ffmpeg failed with code {result}; see ffmpeg.log')
    print(ROOT/'셀리_60초_소개.mp4')


if __name__=='__main__':
    render()
