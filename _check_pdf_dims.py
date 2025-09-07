import sys
from pypdf import PdfReader
from pathlib import Path

path = Path(sys.argv[1])
reader = PdfReader(str(path))
print(f"Pages: {len(reader.pages)}")
for idx, page in enumerate(reader.pages, start=1):
    rot = int(page.get('/Rotate', 0) or 0) % 360
    mb = page.mediabox
    cb = getattr(page, 'cropbox', None) or mb
    llx, lly, urx, ury = float(cb.left), float(cb.bottom), float(cb.right), float(cb.top)
    w, h = urx-llx, ury-lly
    visual_w = h if rot in (90, 270) else w
    visual_h = w if rot in (90, 270) else h
    print(f"{idx:>4}: rotate={rot:>3}  crop=({w:.2f} x {h:.2f})  ll=({llx:.2f},{lly:.2f}) ur=({urx:.2f},{ury:.2f})  visual=({visual_w:.2f} x {visual_h:.2f})")
