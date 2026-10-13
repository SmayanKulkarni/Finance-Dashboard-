"""
Render README screenshots of the three dashboard pages into docs/screenshots/.

Needs LibreOffice (soffice) and poppler (pdftoppm):
    python make_screenshots.py
"""
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageChops

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "docs", "screenshots")
PAGES = ["summary", "overview", "details"]


def crop(path):
    """Trim the white page margin so only the dashboard remains."""
    im = Image.open(path).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    box = ImageChops.difference(im, bg).convert("L").point(lambda p: 255 if p > 12 else 0).getbbox()
    return im.crop(box)


def main():
    tmp = tempfile.mkdtemp()
    xlsx = os.path.join(tmp, "shot.xlsx")
    env = dict(os.environ, BLD_SHOT="1", BLD_OUT=xlsx)
    subprocess.run([sys.executable, os.path.join(HERE, "build_dashboard.py")], env=env, check=True)
    prof = os.path.join(tmp, "profile")
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless",
                    "--convert-to", "pdf", "--outdir", tmp, xlsx],
                   check=True, capture_output=True, timeout=1200)
    subprocess.run(["pdftoppm", "-r", "220", "-png", "-f", "1", "-l", "3",
                    os.path.join(tmp, "shot.pdf"), os.path.join(tmp, "pg")], check=True)
    os.makedirs(OUT_DIR, exist_ok=True)
    pngs = sorted(f for f in os.listdir(tmp) if f.startswith("pg") and f.endswith(".png"))
    for name, f in zip(PAGES, pngs):
        out = os.path.join(OUT_DIR, f"{name}.png")
        crop(os.path.join(tmp, f)).save(out, optimize=True)
        print("wrote", out)
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
