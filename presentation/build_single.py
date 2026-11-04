"""Build 셀리_발표자료.html: the whole deck in one file, for running on another PC.

Edit index.html as usual, then run:  python -I presentation/build_single.py
"""
import base64
from pathlib import Path
import re
import runpy
import subprocess

HERE = Path(__file__).resolve().parent
ASSETS = HERE / "assets"
OUT = HERE / "셀리_발표자료.html"

# Text edited directly in the built file (edit mode) would be overwritten: stop unless it is committed.
if subprocess.run(["git", "diff", "--quiet", "--", str(OUT)], cwd=HERE).returncode:
    raise SystemExit(f"{OUT.name} has uncommitted direct edits. Commit or move them into index.html first.")

runpy.run_path(str(HERE / "embed_fonts.py"))  # new slide text must be in the embedded fonts: the TTFs are dropped below


def b64(name):
    return base64.b64encode((ASSETS / name).read_bytes()).decode()


def swap(html, old, new):
    assert html.count(old) == 1, old
    return html.replace(old, new)


css = (ASSETS / "step-window.css").read_text(encoding="utf-8")
js = (ASSETS / "step-window.js").read_text(encoding="utf-8")
assert "</style" not in css and "</script" not in js

html = (HERE / "index.html").read_text(encoding="utf-8")
html = swap(html, 'href="assets/celly-icon.png"', f'href="data:image/png;base64,{b64("celly-icon.png")}"')
html = swap(html, '<link rel="stylesheet" href="assets/step-window.css">', f"<style>\n{css}\n</style>")
html, faces = re.subn(r'@font-face \{ font-family: Pretendard; src: url\(assets/[^\n]*\n', "", html)
assert faces == 2
# A Blob keeps seeking reliable; a 21MB data: URL in src does not on every browser.
# The data stays in the page so edit mode can save it back; this runs after the editor captured the clean <main>.
html = swap(html, '<video src="assets/demo.mp4" controls', '<video id="demo-video" controls')
html = swap(html, '<script src="assets/step-window.js"></script>', f"""<script type="text/plain" id="demo-video-data">{b64("demo.mp4")}</script>
<script>
{{
  const text = atob(document.getElementById('demo-video-data').textContent);
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i);
  document.getElementById('demo-video').src = URL.createObjectURL(new Blob([bytes], {{ type: 'video/mp4' }}));
}}
</script>
<script>
{js}
</script>""")

assert "assets/" not in re.sub(r"[A-Za-z0-9+/=]{200,}", "", html), "deck still points at a separate file"
OUT.write_text(html, encoding="utf-8")
print(f"{OUT.name}: {OUT.stat().st_size / 1e6:.1f} MB")
