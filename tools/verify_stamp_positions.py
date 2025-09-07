import os
import sys
import fitz  # PyMuPDF

# Allow running from repo root
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from pdf_attachments_ui import insert_text_to_pdf_safe


def find_span(page, text):
    # Try full text
    d = page.get_text("dict")
    for blk in d.get("blocks", []):
        if blk.get("type", 0) != 0:
            continue
        for line in blk.get("lines", []):
            for span in line.get("spans", []):
                s = span.get("text", "")
                if text and text in s:
                    return span
                if "Prilog" in s:
                    return span
    return None


def main():
    if len(sys.argv) < 3:
        print("Usage: python tools/verify_stamp_positions.py <pdf> <text>")
        sys.exit(2)

    src = sys.argv[1]
    text = sys.argv[2]

    if not os.path.exists(src):
        print(f"File not found: {src}")
        sys.exit(1)

    # Produce a stamped copy alongside the source
    prefix = "__verify__"
    insert_text_to_pdf_safe(src, text, True, prefix)
    stamped = os.path.join(os.path.dirname(src), f"{prefix}_{os.path.basename(src)}")

    if not os.path.exists(stamped):
        print("Failed to generate stamped PDF")
        sys.exit(1)

    print(f"Stamped file: {stamped}")
    doc = fitz.open(stamped)
    all_ok = True
    for idx, page in enumerate(doc):
        rot = page.rotation % 360
        pr = page.rect
        span = find_span(page, text)
        if not span:
            print(f"- Page {idx+1}: stamp text NOT FOUND")
            all_ok = False
            continue
        x0, y0, x1, y1 = span["bbox"]
        inside = (x0 >= pr.x0 - 1 and y0 >= pr.y0 - 1 and x1 <= pr.x1 + 1 and y1 <= pr.y1 + 1)
        pos = f"bbox=({x0:.2f},{y0:.2f},{x1:.2f},{y1:.2f}) page=({pr.x0:.2f},{pr.y0:.2f},{pr.x1:.2f},{pr.y1:.2f}) rot={rot}"
        if not inside:
            print(f"- Page {idx+1}: OUTSIDE visible area -> {pos}")
            all_ok = False
            continue

        # Corner expectation
        tol = 200
        near = []
        if rot == 0:
            near.append("right" if x1 >= pr.x1 - tol else "left")
            near.append("top" if y0 <= pr.y0 + tol else "bottom")
        elif rot == 90:
            near.append("left" if x0 <= pr.x0 + tol else "right")
            near.append("top" if y0 <= pr.y0 + tol else "bottom")
        elif rot == 180:
            near.append("left" if x0 <= pr.x0 + tol else "right")
            near.append("bottom" if y1 >= pr.y1 - tol else "top")
        elif rot == 270:
            near.append("right" if x1 >= pr.x1 - tol else "left")
            near.append("bottom" if y1 >= pr.y1 - tol else "top")

        print(f"- Page {idx+1}: OK inside, near {'/'.join(near)} -> {pos}")

    if not all_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()

