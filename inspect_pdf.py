import sys, os
from math import isclose
try:
    from pypdf import PdfReader
except Exception as e:
    sys.stderr.write(f"pypdf import failed: {e}\n")
    sys.exit(2)

PT_TO_MM = 25.4/72.0

def rect_info(rect):
    try:
        return {
            'left': float(rect.left), 'bottom': float(rect.bottom),
            'right': float(rect.right), 'top': float(rect.top),
            'width': float(rect.width), 'height': float(rect.height),
        }
    except Exception:
        # Fallback for older versions
        try:
            ll = rect.lower_left
            ur = rect.upper_right
            left, bottom = float(ll[0]), float(ll[1])
            right, top = float(ur[0]), float(ur[1])
            return {
                'left': left, 'bottom': bottom,
                'right': right, 'top': top,
                'width': right-left, 'height': top-bottom,
            }
        except Exception:
            return None

def page_params(page):
    rot = 0
    for key in ('/Rotate', 'Rotate', 'rotation'):
        try:
            val = page.get(key, 0) if hasattr(page, 'get') else getattr(page, key, 0)
            if isinstance(val, int):
                rot = val
                break
        except Exception:
            pass
    # Some versions expose property 'rotation'
    try:
        rot_prop = getattr(page, 'rotation')
        if isinstance(rot_prop, int):
            rot = rot_prop
    except Exception:
        pass

    media = rect_info(page.mediabox)
    crop = rect_info(getattr(page, 'cropbox', None) or page.mediabox)
    bleed = rect_info(getattr(page, 'bleedbox', None) or page.mediabox)
    art = rect_info(getattr(page, 'artbox', None) or page.mediabox)

    width_pt = media['width'] if media else None
    height_pt = media['height'] if media else None
    return {
        'rotation': rot or 0,
        'mediabox': media,
        'cropbox': crop,
        'bleedbox': bleed,
        'artbox': art,
        'width_pt': width_pt,
        'height_pt': height_pt,
        'width_mm': (width_pt * PT_TO_MM) if width_pt is not None else None,
        'height_mm': (height_pt * PT_TO_MM) if height_pt is not None else None,
    }


def analyze_file(path):
    r = PdfReader(path)
    pages = r.pages
    params = [page_params(p) for p in pages]
    return params


def fmt_mm(v):
    return None if v is None else round(v, 2)


def safe_str(s: str):
    try:
        return s.encode('utf-8', 'backslashreplace').decode('utf-8')
    except Exception:
        return repr(s)

def main():
    if len(sys.argv) < 3:
        print("Usage: python inspect_pdf.py <pdf1> <pdf2>")
        sys.exit(1)
    p1, p2 = sys.argv[1], sys.argv[2]
    # Try to use UTF-8 for stdout to avoid Windows codepage issues
    try:
        sys.stdout.reconfigure(encoding='utf-8')  # type: ignore[attr-defined]
    except Exception:
        pass
    print("FILE1:", safe_str(p1))
    a = analyze_file(p1)
    if len(a) >= 2:
        p2a = a[1]
        print("Page 2:")
        print({
            'rotation': p2a['rotation'],
            'media_w_pt': round(p2a['mediabox']['width'], 3),
            'media_h_pt': round(p2a['mediabox']['height'], 3),
            'media_w_mm': fmt_mm(p2a['width_mm']),
            'media_h_mm': fmt_mm(p2a['height_mm']),
            'crop==media': p2a['cropbox'] == p2a['mediabox'],
        })
        print("Page 2 boxes:")
        print({'mediabox': p2a['mediabox'], 'cropbox': p2a['cropbox'], 'bleedbox': p2a['bleedbox'], 'artbox': p2a['artbox']})
    else:
        print("File1 has <2 pages")

    print("\nFILE2:", safe_str(p2))
    b = analyze_file(p2)
    print(f"Pages: {len(b)}")
    # unique size/rotation combos
    combos = {}
    for i, p in enumerate(b, start=1):
        key = (round(p['mediabox']['width'],3), round(p['mediabox']['height'],3), int(p['rotation']))
        combos.setdefault(key, []).append(i)
    print("Unique (width_pt,height_pt,rotation) -> pages:")
    for k, idxs in combos.items():
        w,h,r = k
        print({
            'size_pt': (w,h),
            'size_mm': (round(w*PT_TO_MM,2), round(h*PT_TO_MM,2)),
            'rotation': r,
            'pages': idxs[:50] + (["..."] if len(idxs)>50 else []),
        })
    # Any cropbox differing from mediabox?
    diffs = []
    for i,p in enumerate(b, start=1):
        if p['cropbox'] != p['mediabox']:
            diffs.append(i)
    print("Pages with cropbox != mediabox:", diffs or None)

    # Compare against file1 page 2
    if len(a) >= 2:
        f1 = a[1]
        ref = (round(f1['mediabox']['width'],3), round(f1['mediabox']['height'],3), int(f1['rotation']))
        print("\nComparison vs file1 page2 (size_pt, rotation):", ref)
        matching = [i for i,p in enumerate(b, start=1)
                    if (round(p['mediabox']['width'],3), round(p['mediabox']['height'],3), int(p['rotation'])) == ref]
        print("Pages in file2 matching file1 page2:", matching or None)
        non_matching = [i for i in range(1, len(b)+1) if i not in matching]
        print("Pages in file2 NOT matching file1 page2:", non_matching or None)

if __name__ == '__main__':
    main()

