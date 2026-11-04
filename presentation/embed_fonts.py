"""Embed Pretendard into index.html, subset to the deck's characters plus all Hangul.

Run again after editing slide text:  python -I presentation/embed_fonts.py
The full TTFs in assets/ stay as a second family, so any character added later still renders in Pretendard
when the assets folder is present.
"""
import base64
import io
from pathlib import Path
import re
import string

from fontTools import subset
from fontTools.ttLib import TTFont

HERE = Path(__file__).resolve().parent
HTML = HERE / "index.html"
START, END = "/* embedded-fonts:start */", "/* embedded-fonts:end */"

text = "".join(p.read_text(encoding="utf-8") for p in
               [HTML, HERE / "assets" / "step-window.js", HERE / "assets" / "step-window.css"])
chars = set(text) | set(string.printable) | set("·…→←↑↓“”‘’「」×—–₩%")
chars |= {chr(c) for c in range(0xAC00, 0xD7A4)}  # every Hangul syllable: text typed in edit mode must keep the font


def subset_b64(weight):
    font = TTFont(HERE / "assets" / f"Pretendard-{weight}.ttf")
    opts = subset.Options()
    opts.flavor = "woff"  # zlib-compressed; woff2 would need brotli
    opts.layout_features = ["*"]
    s = subset.Subsetter(opts)
    s.populate(text="".join(chars))
    s.subset(font)
    buf = io.BytesIO()
    font.flavor = "woff"
    font.save(buf)
    return base64.b64encode(buf.getvalue()).decode()


faces = "\n".join(
    f'@font-face {{ font-family: "Pretendard Embedded"; src: url(data:font/woff;base64,{subset_b64(w)}) format("woff"); '
    f'font-weight: {n}; font-display: block; }}' for w, n in [("Regular", 400), ("Bold", 700)])
block = f"{START}\n{faces}\n{END}"

html = HTML.read_text(encoding="utf-8")
if START in html:
    html = re.sub(re.escape(START) + ".*?" + re.escape(END), lambda _: block, html, flags=re.S)
else:
    html = html.replace("<style>\n", "<style>\n" + block + "\n", 1)
HTML.write_text(html, encoding="utf-8")
print(f"embedded {len(chars)} characters; index.html is now {HTML.stat().st_size / 1e6:.2f} MB")
