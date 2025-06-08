import os
import json
import re
import glob
import time
import datetime
import shutil
import requests
import getpass
import tkinter as tk

import cv2
import fitz       # PyMuPDF
import numpy as np
import pandas as pd
import pdfplumber

from tkinter import filedialog
from DecalExtract_helper import get_valid_api_key, fetch_pdf_via_api

# ── Configuration ──────────────────────────────────────────────────────────────
SITE_ID         = 733
DOWNLOAD_TIMEOUT= 8     # seconds to wait for PDF generation
DPI             = 300
THICKNESS_IN    = 0.004
MATERIAL_DENSITY= 0.035
FACTOR          = 166
STEP_DELAY      = 5 #whenever you need a short delay insert: time.sleep(STEP_DELAY)

# Map keyword labels to BGR fill colors
COLOR_MAP = {
    'green':  ( 81, 167,   0),
    'yellow': ( 24, 241, 244),
    'red':    ( 15,  46, 209),
    'blue':   (204, 102,   0),
    'black':  (  0,   0,   0),
}
# ── The Secret Sauce ──────────────────────────────────────────────────────────
KEY_FILE = os.path.expanduser("~/.decal_api_key.json")
API_ENDPOINT = "https://hal4ecrr1tk.execute-api.us-east-1.amazonaws.com/prod/get_current_drawing"

# ── Utility Functions ─────────────────────────────────────────────────────────
def get_valid_api_key() -> str:
    """
    Prompt the user once for X-API-KEY, store it in ~/.decal_api_key.json,
    and return it.  On subsequent runs, re-use the saved key.
    """
    # 1) Try to load existing key
    if os.path.exists(KEY_FILE):
        try:
            with open(KEY_FILE, "r") as f:
                data = json.load(f)
                key = data.get("x_api_key")
                if key:
                    return key
        except Exception:
            pass

    # 2) Ask the user to paste in their API key
    print("Please paste your X-API-KEY for the signed-URL service:")
    key = getpass.getpass(prompt="X-API-KEY: ")

    # 3) Save it for next time
    try:
        with open(KEY_FILE, "w") as f:
            json.dump({"x_api_key": key}, f)
        os.chmod(KEY_FILE, 0o600)
    except Exception as e:
        print(f"Warning: could not save key to {KEY_FILE}: {e}")

    return key
    
def render_pdf_color_page(pdf_path, dpi=300):
    """Load the first page of PDF at `dpi` into a BGR numpy image."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    scale = dpi / 72
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8)
    img = arr.reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        return cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    else:
        return cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        
def find_union_of_ink_contours(img_color, min_area=500, pad_pct=0.05, dbg_dir=None, dbg_name=None):
    """
    Instrumented “union of all ink” fallback.  Detect every non-white contour ≥ min_area,
    print out its area and bounding box, then union them all and pad by pad_pct.

    Parameters:
    - img_color : np.ndarray (BGR) of the full-page image
    - min_area   : int  → discard any contour whose area < this (default 500)
    - pad_pct    : float→ pad the final union-outwards by pad_pct * (width/height)
    - dbg_dir    : str  → (optional) path to your debugging folder (e.g. 'debugging')
    - dbg_name   : str  → (optional) base filename for the debug image (e.g. 'part1234')

    Returns:
    - (x0p, y0p, x1p, y1p) or None
    """

    import cv2
    import numpy as np
    import os

    # 0) Prepare gray + threshold mask (inverse: ink = white=255, background=0)
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)

    # 1) Find all external contours on that mask
    cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    total_cnts = len(cnts)
    print(f"      · [DEBUG] find_union_of_ink_contours: found {total_cnts} total contours")

    if not cnts:
        print("         → No contours found at all.")
        return None

    # 2) Keep only contours whose area >= min_area
    big_boxes = []
    for idx, c in enumerate(cnts):
        area = cv2.contourArea(c)
        x, y, w, h = cv2.boundingRect(c)

        print(f"         → Contour #{idx}: area={area:.0f}, bbox=({x},{y},{w},{h})")
        if area < min_area:
            print(f"            (discarded, area < {min_area})")
            continue

        print(f"            (accepted)")
        big_boxes.append((x, y, w, h))

    if not big_boxes:
        print(f"      · [DEBUG] No contours ≥ min_area({min_area}) → return None")
        return None

    # 3) Union all those bounding rects into one big box
    x0 = min(box[0] for box in big_boxes)
    y0 = min(box[1] for box in big_boxes)
    x1 = max(box[0] + box[2] for box in big_boxes)
    y1 = max(box[1] + box[3] for box in big_boxes)
    print(f"      · [DEBUG] Union of accepted boxes = ({x0}, {y0}, {x1}, {y1}) before padding")

    # 4) Pad that union OUTWARDS by pad_pct in each direction
    img_h, img_w = img_color.shape[:2]
    rect_w = x1 - x0
    rect_h = y1 - y0
    pad_x = int(rect_w * pad_pct)
    pad_y = int(rect_h * pad_pct)

    x0p = max(x0 - pad_x, 0)
    y0p = max(y0 - pad_y, 0)
    x1p = min(x1 + pad_x, img_w)
    y1p = min(y1 + pad_y, img_h)
    print(f"      · [DEBUG] After pad_pct={pad_pct*100:.0f}%, padded box = ({x0p}, {y0p}, {x1p}, {y1p})")

    if x1p <= x0p or y1p <= y0p:
        print("      · [DEBUG] Invalid padded box (zero or negative area). Returning None.")
        return None

    # 5) (Optional) Write a debug image showing each accepted box in GREEN
    #    and the final padded union in RED.  Uncomment if you want to save it.
    if dbg_dir and dbg_name:
        try:
            debug_vis = img_color.copy()
            # draw each accepted box in GREEN:
            for (bx, by, bw, bh) in big_boxes:
                cv2.rectangle(debug_vis,
                              (bx, by),
                              (bx + bw, by + bh),
                              (0, 255, 0), 2)

            # draw the final padded union in RED:
            cv2.rectangle(debug_vis,
                          (x0p, y0p),
                          (x1p, y1p),
                          (0, 0, 255), 3)

            os.makedirs(dbg_dir, exist_ok=True)
            dbg_path = os.path.join(dbg_dir, f"{dbg_name}_union_debug.png")
            cv2.imwrite(dbg_path, debug_vis)
            print(f"      · [DEBUG] Wrote debug image → {dbg_path}")
        except Exception as ex:
            print(f"      · [DEBUG] Failed to write debug image: {ex}")

    return (x0p, y0p, x1p, y1p)

def extract_color_label(pdf_path):
    """
    Stub: read text near the top of the PDF and return one of
    'green','yellow','red','blue','black'.  Use pdfplumber or regex.
    """
    # TODO: open pdfplumber, search for the crop-band label and return lowercase key
    return 'green'
    
def load_template_sets(root='templates'):
    """
    Look under root/set1…setN for quad-templates (top_left, top_right, bottom_left, bottom_right).
    We now consider jpg/jpeg/png extensions. Returns a list of (templates, offsets).
    """
    quad_names = ['top_left','top_right','bottom_left','bottom_right']
    sets = []
    for folder in sorted(glob.glob(os.path.join(root, 'set*'))):
        tpl_dict, off_dict = {}, {}
        for quad in quad_names:
            for ext in ('jpg','jpeg','png'):
                path = os.path.join(folder, f"{quad}.{ext}")
                if os.path.exists(path):
                    img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
                    edges = cv2.Canny(img, 50, 150)
                    h, w = edges.shape
                    tpl_dict[quad] = edges
                    # offsets inside the small crop marks
                    if quad == 'top_left':
                        off_dict[quad] = (w-1, h-1)
                    elif quad == 'top_right':
                        off_dict[quad] = (0,   h-1)
                    elif quad == 'bottom_left':
                        off_dict[quad] = (w-1, 0)
                    else:  # bottom_right
                        off_dict[quad] = (0,   0)
                    break
        if len(tpl_dict) == 4:
            sets.append((tpl_dict, off_dict))

    if not sets:
        raise FileNotFoundError("No complete template-sets found under "+root)
    return sets

def detect_with_one_set(img_gray, templates, offsets):
    """Run matchTemplate for each of the 4 corners in this single set."""
    H, W = img_gray.shape
    rois = {
        'top_left':     (0,     0,   W//2,   H//2),
        'top_right':    (W//2,  0,   W,      H//2),
        'bottom_left':  (0,     H//2,W//2,   H),
        'bottom_right': (W//2,  H//2,W,      H)
    }
    dets = {}
    for q in templates:
        x1, y1, x2, y2 = rois[q]
        roi = img_gray[y1:y2, x1:x2]
        edges_roi = cv2.Canny(roi, 50, 150)
        res = cv2.matchTemplate(edges_roi, templates[q], cv2.TM_CCOEFF_NORMED)
        _, _, _, loc = cv2.minMaxLoc(res)
        offx, offy = offsets[q]
        dets[q] = (x1 + loc[0] + offx, y1 + loc[1] + offy)
    return dets
    
def find_nearby_blob_group(
    img_color,
    min_area: int = 500,
    tol: int    = 50,
    pad: int    = 20
) -> tuple[int,int,int,int] | None:
    """
    Locate ALL non-white contours in img_color. Pick the single largest contour
    (by area), then find any other contours whose bottom‐edge is within `tol` pixels
    of that largest contour's bottom edge. If at least two contours qualify, union
    their bounding boxes into one rectangle, pad by `pad` pixels on each side, and return it.
    Otherwise return None.

    - min_area : ignore any contour whose w*h < min_area
    - tol      : vertical tolerance (pixels) to group bottoms of contours
    - pad      : pad (pixels) to expand the unioned bounding box (clamped)
    """
    import cv2
    import numpy as np

    # 1) Grayscale + threshold → every pixel < 250 becomes “ink” (255 in mask), white → 0.
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)

    # 2) Find all external contours
    cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None

    # 3) Keep only those whose bounding‐rect area >= min_area
    boxes = []
    for c in cnts:
        x, y, w, h = cv2.boundingRect(c)
        if w*h < min_area:
            continue
        boxes.append((x, y, w, h))

    if len(boxes) < 2:
        return None

    # 4) Identify the largest box by area (w*h)
    main = max(boxes, key=lambda b: b[2]*b[3])
    mx, my, mw, mh = main
    main_bottom = my + mh

    # 5) Gather any other box whose bottom is within tol pixels of main_bottom
    group = [main]
    for (x, y, w, h) in boxes:
        if (x, y, w, h) == main:
            continue
        bottom = y + h
        if abs(bottom - main_bottom) <= tol:
            group.append((x, y, w, h))

    if len(group) < 2:
        return None

    # 6) Union all group‐boxes into a single bounding rectangle
    xs = [b[0] for b in group] + [b[0] + b[2] for b in group]
    ys = [b[1] for b in group] + [b[1] + b[3] for b in group]
    x0 = min(xs)
    y0 = min(ys)
    x1 = max(xs)
    y1 = max(ys)

    # 7) Apply uniform padding → clamp within image
    img_h, img_w = img_color.shape[:2]
    x0p = max(x0 - pad, 0)
    y0p = max(y0 - pad, 0)
    x1p = min(x1 + pad, img_w)
    y1p = min(y1 + pad, img_h)

    return (x0p, y0p, x1p, y1p)

def crop_blob_bbox(img_gray):
    """Return bounding box (x0,y0,x1,y1) of the largest dark blob."""
    _, th = cv2.threshold(img_gray, 250, 255, cv2.THRESH_BINARY_INV)
    cnts, _ = cv2.findContours(th, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None
    c = max(cnts, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(c)
    return (x, y, x + w, y + h)
    
def detect_enclosed_box(img_gray, min_area=5000):
    """
    Find the contour with the largest perimeter in a binary‐inverted version of img_gray,
    then return its bounding‐rectangle. This reliably catches a single rounded‐corner border
    even if the top edge is lightly anti‐aliased.
    - img_gray: a BGR→Gray frame (numpy array)
    - min_area: ignore tiny contours smaller than this (pixels^2)
    Returns (x0, y0, x1, y1) or None.
    """
    # 1) Invert threshold so that nearly‐black border+ink → white (255), background → 0
    _, thresh = cv2.threshold(img_gray, 250, 255, cv2.THRESH_BINARY_INV)

    # 2) Find all external contours on that mask (CHAIN_APPROX_NONE to preserve curves)
    cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)

    best_contour = None
    best_peri = 0

    for c in cnts:
        area = cv2.contourArea(c)
        if area < min_area:
            continue  # ignore tiny specks
        peri = cv2.arcLength(c, True)
        if peri > best_peri:
            best_peri = peri
            best_contour = c

    if best_contour is None:
        return None

    # 3) Return the bounding‐rectangle of that “longest perimeter” contour
    x, y, w, h = cv2.boundingRect(best_contour)
    return (x, y, x + w, y + h)

def rect_intersection(a, b):
    """Intersection area of two rects a=(x0,y0,x1,y1), b likewise."""
    x0 = max(a[0], b[0]); y0 = max(a[1], b[1])
    x1 = min(a[2], b[2]); y1 = min(a[3], b[3])
    if x1 <= x0 or y1 <= y0:
        return 0
    return (x1 - x0) * (y1 - y0)

def detect_best_crop(img_color, template_sets):
    """
    Try each template-set; score by how much of the  blob
    sits inside the resulting crop. Return the best (x0,y0,x1,y1).
    """
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    blob_box = crop_blob_bbox(gray) or (0, 0, gray.shape[1], gray.shape[0])
    blob_area = (blob_box[2] - blob_box[0]) * (blob_box[3] - blob_box[1])

    best_score, best_rect = -1, (0, 0, gray.shape[1], gray.shape[0])
    for templates, offsets in template_sets:
        try:
            corners = detect_with_one_set(gray, templates, offsets)
            # compute average-rectangle
            tl, tr = corners['top_left'], corners['top_right']
            bl, br = corners['bottom_left'], corners['bottom_right']
            x0 = int((tl[0] + bl[0]) / 2)
            x1 = int((tr[0] + br[0]) / 2)
            y0 = int((tl[1] + tr[1]) / 2)
            y1 = int((bl[1] + br[1]) / 2)
        except:
            continue

        rect = (x0, y0, x1, y1)
        score = rect_intersection(blob_box, rect) / float(blob_area or 1)
        if score > best_score and score > 0.8:
            best_score, best_rect = score, rect

    if best_score < 0.8:
        # fallback
        x0, y0, x1, y1 = blob_box
        if x1 <= x0 or y1 <= y0:
            h, w = gray.shape
            m = int(0.01 * min(h, w))
            x0, y0, x1, y1 = m, m, w - m, h - m
        best_rect = (x0, y0, x1, y1)

    print(f"    · Chosen crop={best_rect} (score={best_score:.2f})")
    return best_rect

def select_all_crop_candidates(img_color, template_sets, penalty_thresh=0.1):
    """
    Returns a list of non-overlapping (x0,y0,x1,y1) rectangles
    whose penalty (edge-ink on crop border) < penalty_thresh,
    all having virtually the same aspect ratio.
    """
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    _, blob = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
    blob = (blob > 0).astype(np.uint8)
    H, W = gray.shape
    cands = []

    for tpl, offs in template_sets:
        try:
            corners = detect_with_one_set(gray, tpl, offs)
            x0 = int((corners['top_left'][0] + corners['bottom_left'][0]) / 2)
            x1 = int((corners['top_right'][0] + corners['bottom_right'][0]) / 2)
            y0 = int((corners['top_left'][1] + corners['top_right'][1]) / 2)
            y1 = int((corners['bottom_left'][1] + corners['bottom_right'][1]) / 2)
        except:
            continue

        # penalty = ink on 5-pixel wide border
        e = 5
        top    = blob[y0:y0 + e, x0:x1]
        bot    = blob[y1 - e:y1, x0:x1]
        left   = blob[y0:y1, x0:x0 + e]
        right  = blob[y0:y1, x1 - e:x1]
        penalty = float(top.sum() + bot.sum() + left.sum() + right.sum()) / ((x1 - x0) * (y1 - y0))
        if penalty > penalty_thresh or x1 <= x0 or y1 <= y0:
            continue
        cands.append(((x1 - x0) / (y1 - y0), (x0, y0, x1, y1)))

    if not cands:
        return []

    # group by ratio (within 5%)
    base_ratio = cands[0][0]
    boxes = []
    for ratio, box in cands:
        if abs(ratio - base_ratio) / base_ratio < 0.05:
            if all(rect_intersection(box, b) == 0 for b in boxes):
                boxes.append(box)
    return sorted(boxes, key=lambda b: b[0])
    
def recolor_layer(image, color_bgr):
    """
    Make every true‐black pixel → color_bgr, everything else transparent.
    Returns 4-channel BGRA.
    """
    bgr = image.copy()
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    alpha = (gray == 0).astype(np.uint8) * 255
    # overlay fill color where alpha=255
    for c in range(3):
        bgr[:, :, c] = np.where(alpha == 255, color_bgr[c], 0)
    return cv2.cvtColor(bgr, cv2.COLOR_BGR2BGRA)


def parse_dimensions_from_pdf(pdf_path):
    """
    Scan the first page of pdf for:
      1) “Dimensions (h x w): 1.25" x 5.75"”
      2) “OVER ALL LENGTH IS 14 INCHES”
      3) “450 mm x 129 mm”  (millimeters → inches)
      4) any free “#″ x #″” fallback (inches)

    Returns:
      - (height_in, width_in) in inches as floats,
      - OR (length_in, None) if only an “OVER ALL LENGTH” was present,
      - OR (0.0, 0.0) if nothing parseable found.

    Notes:
      • If you find an “mm” match, you convert each number with / 25.4 → inch.
      • This order of checks ensures that “mm” lines get priority over a generic “#″ x #″” fallback.
    """
    import pdfplumber
    import re

    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""

    # 1) Explicit (h x w) in inches
    m = re.search(
        r'Dimensions\s*\(\s*h\s*[x×]\s*w\s*\)\s*:\s*([\d.]+)\s*["”]?\s*[x×]\s*([\d.]+)\s*["”]?',
        text, re.IGNORECASE
    )
    if m:
        return float(m.group(1)), float(m.group(2))

    # 2) “OVER ALL LENGTH IS 14 INCHES”
    m2 = re.search(
        r'OVER\s*ALL\s*LENGTH\s*(?:IS|=)\s*([\d.]+)\s*INCH',
        text, re.IGNORECASE
    )
    if m2:
        length_in = float(m2.group(1))
        return length_in, None

    # 3) Millimeter line: “450 mm x 129 mm” (may use lowercase mm or uppercase MM)
    mmm = re.search(
        r'([\d.]+)\s*mm\s*[x×]\s*([\d.]+)\s*mm',
        text, re.IGNORECASE
    )
    if mmm:
        # Convert each mm → inches by dividing by 25.4
        h_mm = float(mmm.group(1))
        w_mm = float(mmm.group(2))
        h_in = h_mm / 25.4
        w_in = w_mm / 25.4
        return h_in, w_in

    # 4) Any free “#″ x #″” fallback (inches)
    m3 = re.search(r'([\d.]+)\s*["”]\s*[x×]\s*([\d.]+)\s*["”]?', text)
    if m3:
        return float(m3.group(1)), float(m3.group(2))

    # Nothing matched
    return 0.0, 0.0
    
def extract_color_label(pdf_path: str,
                        crop_y0: float = None) -> str:
    """
    Scan the first page of pdf_path for any of the COLOR_MAP keys
    ('green','yellow','red','blue','black').  If crop_y0 is provided,
    only consider labels whose bottom is <= crop_y0 (i.e. text above
    the decal).  Return the closest label to crop_y0, or 'black' as default.
    """
    keywords = set(COLOR_MAP.keys())
    best = None
    best_bottom = -1

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        # each word has: text, x0,x1,top,bottom
        for w in page.extract_words():
            txt = w["text"].strip().lower()
            if txt in keywords:
                btm = float(w["bottom"])
                # if crop_y0 given, ignore words below the decal region
                if crop_y0 is not None and btm > crop_y0:
                    continue
                # pick the word whose bottom is closest to crop_y0
                if crop_y0 is None or btm > best_bottom:
                    best_bottom = btm
                    best = txt

    return best or "black"
                            
def crop_full_logo(pdf_path, dpi=300, margin_pt=5):
    """
    Find the “…mm” dimension line in the PDF and return
    its top‐Y in pixels (minus a small margin). If nothing
    is found, returns None.
    """
    with pdfplumber.open(pdf_path) as pdf:
        words = pdf.pages[0].extract_words()
    dims = [w for w in words if w["text"].lower().endswith("mm")]
    if not dims:
        return None
    dim_top_pt = min(w["top"] for w in dims)
    return int((dim_top_pt - margin_pt) * dpi / 72)

def match_one_corner(img_color, tpl_edges, offset, quadrant):
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    H, W = gray.shape
    rois = {
        'top-left':    (0,     0,     W//2,  H//2),
        'top-right':   (W//2,  0,     W,     H//2),
        'bottom-left': (0,     H//2,  W//2,  H),
        'bottom-right':(W//2,  H//2,  W,     H),
    }
    x1, y1, x2, y2 = rois[quadrant]
    roi = gray[y1:y2, x1:x2]
    edges_roi = cv2.Canny(roi, 50, 150)
    res = cv2.matchTemplate(edges_roi, tpl_edges, cv2.TM_CCOEFF_NORMED)
    _, _, _, maxloc = cv2.minMaxLoc(res)
    ox, oy = offset
    return (x1 + maxloc[0] + ox, y1 + maxloc[1] + oy)

def select_best_crop_box(img_color, template_sets, expected_ratio=None, edge=5, ar_weight=1000):
    """
    Try each template-set to get a candidate box, then score by:
      • penalty: how much “ink” lies on the 5px border
      • strict aspect-ratio check (±5%)
      • corner-template match-confidence (>= 0.85)
    Return the (x0,y0,x1,y1) with the lowest total_score.
    """
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    # everything below 250 is “ink”
    _, blob = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
    blob = (blob > 0).astype(np.uint8)

    candidates = []
    H, W = blob.shape

    for templates, offsets in template_sets:
        try:
            # 1) Try to detect all four corner-brackets with high confidence
            corners = {}
            for quad in ('top_left','top_right','bottom_left','bottom_right'):
                # matchTemplate returns (minVal, maxVal, minLoc, maxLoc)
                # but we want to inspect maxVal to ensure it >= 0.85
                x1, y1, x2, y2 = {
                    'top_left':      (0,     0,    W//2,  H//2),
                    'top_right':     (W//2,  0,    W,     H//2),
                    'bottom_left':   (0,     H//2, W//2,  H),
                    'bottom_right':  (W//2,  H//2, W,     H),
                }[quad]
                roi = gray[y1:y2, x1:x2]
                edges_roi = cv2.Canny(roi, 50, 150)
                res = cv2.matchTemplate(edges_roi, templates[quad], cv2.TM_CCOEFF_NORMED)
                _, maxVal, _, maxLoc = cv2.minMaxLoc(res)

                # **(a)** If confidence < 0.85, abort this template-set entirely
                if maxVal < 0.85:
                    raise ValueError(f"{quad} corner match too weak ({maxVal:.2f})")

                offx, offy = offsets[quad]
                corners[quad] = (x1 + maxLoc[0] + offx, y1 + maxLoc[1] + offy)

            # 2) Average the four corners into a rectangle
            tl, tr = corners['top_left'], corners['top_right']
            bl, br = corners['bottom_left'], corners['bottom_right']

            x0 = int((tl[0] + bl[0]) / 2)
            y0 = int((tl[1] + tr[1]) / 2)
            x1 = int((tr[0] + br[0]) / 2)
            y1 = int((bl[1] + br[1]) / 2)

        except Exception:
            # this template-set failed (either low confidence or corner detection failed)
            continue

        # 3) Clip into image bounds
        x0n, y0n = max(0, x0), max(0, y0)
        x1n, y1n = min(W, x1), min(H, y1)
        if x1n <= x0n or y1n <= y0n:
            continue

        # 4) Compute border-ink penalty (5px wide)
        top    = blob[y0n:y0n+edge, x0n:x1n]
        bottom = blob[y1n-edge:y1n, x0n:x1n]
        left   = blob[y0n:y1n, x0n:x0n+edge]
        right  = blob[y0n:y1n, x1n-edge:x1n]
        penalty = int(top.sum() + bottom.sum() + left.sum() + right.sum())

        # 5) Strict aspect-ratio penalty
        w_rect = float(x1n - x0n)
        h_rect = float(y1n - y0n)
        ar = (w_rect / h_rect) if (h_rect > 0) else 0

        # If expected_ratio provided, reject if >5% off
        if expected_ratio:
            if abs(ar - expected_ratio)/expected_ratio > 0.05:
                # reject this candidate completely
                continue
            ar_penalty = abs(ar - expected_ratio) * ar_weight
        else:
            ar_penalty = 0

        total_score = penalty + ar_penalty
        candidates.append((total_score, (x0n, y0n, x1n, y1n)))

    # 6) If no “good” corner-based candidate, fallback
    if not candidates:
        raise RuntimeError("No valid crop candidates found")

    # 7) Pick the rectangle with the lowest combined score
    _, best_box = min(candidates, key=lambda t: t[0])
    return best_box

def wait_for_login(driver, timeout=300):
    """
    Block until the IRR search field reappears,
    i.e. you’ve logged back in.
    """
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, 'docLibContainer_search_field'))
    )

def find_horizontal_aligned_union(img_color, min_area=2000, tol=250, pad_pct=0.05, min_ratio=0.5):
    """
    Group only those “big” contours (area ≥ min_area) whose vertical centers
    lie within `tol` pixels of the largest contour’s center AND whose aspect
    ratio (width/height) ≥ min_ratio.  Then return the union of those bounding
    boxes, padded by pad_pct.  If no suitable contour ≥ min_area is found, return None.
    """

    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
    cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # 1) Collect all contours with area >= min_area
    big = []
    for c in cnts:
        area = cv2.contourArea(c)
        if area < min_area:
            continue
        x, y, w, h = cv2.boundingRect(c)
        cy = y + (h / 2)
        ratio = float(w) / float(h) if h > 0 else 0.0
        big.append({'bbox': (x, y, x + w, y + h), 'area': area, 'cy': cy, 'ratio': ratio})

    if not big:
        return None

    # 2) Find the single largest contour (“main blob”)
    main = max(big, key=lambda b: b['area'])
    cy_main = main['cy']

    # 3) Always include the main blob.  Then group any other ‘big’ contour whose
    #    vertical center is within tol AND whose aspect ratio >= min_ratio.
    group = [main]
    for entry in big:
        if entry is main:
            continue
        if abs(entry['cy'] - cy_main) <= tol and entry['ratio'] >= min_ratio:
            group.append(entry)

    # 4) Compute union of all bounding boxes in that group
    xs = []
    ys = []
    for entry in group:
        x0_, y0_, x1_, y1_ = entry['bbox']
        xs.extend([x0_, x1_])
        ys.extend([y0_, y1_])

    x0u = min(xs)
    y0u = min(ys)
    x1u = max(xs)
    y1u = max(ys)

    # 5) Pad the union‐box by pad_pct on all sides (clamp to image edges)
    h_img, w_img = img_color.shape[:2]
    rect_w = x1u - x0u
    rect_h = y1u - y0u
    pad_x = int(rect_w * pad_pct)
    pad_y = int(rect_h * pad_pct)

    x0p = max(x0u - pad_x, 0)
    y0p = max(y0u - pad_y, 0)
    x1p = min(x1u + pad_x, w_img)
    y1p = min(y1u + pad_y, h_img)

    return (x0p, y0p, x1p, y1p)

def download_pdf_via_api(part_number: str, pdf_dir: str, api_key: str) -> str:
    """
    1) POST {part_number} + x-api-key
    2) parse JSON or raw text for a signed CloudFront URL
    3) GET that URL → save PDF under pdf_dir
    4) return local path, or None on failure
    """
    body = {"part_number": part_number}
    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key
    }

    # a) call the API
    try:
        resp = requests.post(API_ENDPOINT,
                             headers=headers,
                             data=json.dumps(body),
                             timeout=30)
        resp.raise_for_status()
    except Exception as e:
        print(f"   · [ERROR] API call failed for '{part_number}': {e}")
        return None

    # b) extract signed URL
    signed_url = None
    try:
        payload = resp.json()
        if isinstance(payload, dict) and "url" in payload:
            signed_url = payload["url"]
        elif isinstance(payload, str) and payload.startswith("http"):
            signed_url = payload
        else:
            raise ValueError(f"Bad payload: {payload!r}")
    except ValueError:
        text = resp.text.strip()
        if text.startswith("http"):
            signed_url = text
        else:
            print(f"   · [ERROR] Unexpected API response for '{part_number}': {resp.text!r}")
            return None

    # c) download the PDF bytes
    try:
        dl = requests.get(signed_url, timeout=60)
        dl.raise_for_status()
    except Exception as e:
        print(f"   · [ERROR] Could not GET PDF for '{part_number}': {e}")
        return None

    # d) save to disk
    os.makedirs(pdf_dir, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = part_number.replace(" ", "_")
    filename = f"{safe_name}_{ts}.pdf"
    out_path = os.path.join(pdf_dir, filename)
    try:
        with open(out_path, "wb") as f:
            f.write(dl.content)
    except Exception as e:
        print(f"   · [ERROR] Writing PDF to disk failed: {e}")
        return None

    print(f"   · [API] Downloaded PDF → {out_path}")
    return out_path
    
def find_grouped_union_of_ink_contours(img_color, min_area=500, pad_pct=0.05, proximity_px=50):
    """
    1) Threshold `img_color` so that any pixel <250→foreground (ink).
    2) Find all external contours in that thresholded mask.
    3) Keep only those contours whose area >= min_area.
    4) Cluster together any contours whose bounding boxes come within
       `proximity_px` pixels horizontally (and that overlap vertically at all).
    5) Compute one big bounding box around that cluster (group) and then pad it
       outward by pad_pct * (width_of_group) horizontally and pad_pct * (height_of_group) vertically.
    Returns (x0, y0, x1, y1) or None if no contour was found.
    """

    # Step 1: Create a binary “ink mask”
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)

    # Step 2: Find all external contours
    cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None

    # Step 3: Filter by area >= min_area
    boxes = []
    for c in cnts:
        area = cv2.contourArea(c)
        if area < min_area:
            continue
        x, y, w, h = cv2.boundingRect(c)
        boxes.append((x, y, w, h))
    if not boxes:
        return None

    # Step 4: Sort boxes by x-coordinate (left edge)
    boxes = sorted(boxes, key=lambda b: b[0])

    # We'll form a single “cluster” by starting from the largest‐area contour,
    # then absorbing any other box whose bounding box’s x-range comes within
    # proximity_px, AND whose y-range overlaps at all.  (This guarantees we pull
    # in the entire disconnected logo, but nothing extremely far away.)
    # First, identify the “largest” blob by area.
    #    (We know each box is (x, y, w, h).)
    areas = [w*h for (x,y,w,h) in boxes]
    largest_idx = int(np.argmax(areas))
    gx, gy, gw, gh = boxes[largest_idx]

    # Our “group’s bounding rect so far”:
    group_x0 = gx
    group_y0 = gy
    group_x1 = gx + gw
    group_y1 = gy + gh

    # Now attempt to absorb any other contours that lie “close enough.”
    # We do a single pass over all boxes (including ones on either side).  If a box’s
    # x‐range [x, x+w] is within proximity_px of our current group’s [group_x0, group_x1],
    # and its y‐range [y, y+h] overlaps at all with our group’s [group_y0, group_y1],
    # we absorb it and expand our group.  We repeat until no new box can be absorbed.
    absorbed = True
    used = set([largest_idx])

    while absorbed:
        absorbed = False
        for i, (x, y, w, h) in enumerate(boxes):
            if i in used:
                continue
            # Horizontal proximity check:
            #    We say “close enough” if box’s x is within proximity_px of group_x1,
            #    OR if group_x0 is within proximity_px of box’s x+w.
            bx0, bx1 = x, x + w
            dist_horiz = 0
            if bx1 < group_x0:
                dist_horiz = group_x0 - bx1
            elif bx0 > group_x1:
                dist_horiz = bx0 - group_x1
            else:
                dist_horiz = 0  # they overlap horizontally already

            # Vertical overlap check: do they share any y-range?
            #    (i.e. box’s y..y+h intersects group_y0..group_y1)
            vy0, vy1 = y, y + h
            overlap_vert = not (vy1 < group_y0 or vy0 > group_y1)

            if dist_horiz <= proximity_px and overlap_vert:
                # absorb it
                used.add(i)
                absorbed = True
                group_x0 = min(group_x0, x)
                group_y0 = min(group_y0, y)
                group_x1 = max(group_x1, x + w)
                group_y1 = max(group_y1, y + h)

    # At this point, (group_x0, group_y0) … (group_x1, group_y1) covers
    # all contours in that cluster.  Now pad this bounding box outward by pad_pct:
    img_h, img_w = img_color.shape[:2]
    gw = group_x1 - group_x0
    gh = group_y1 - group_y0
    pad_x = int(gw * pad_pct)
    pad_y = int(gh * pad_pct)

    x0p = max(group_x0 - pad_x, 0)
    y0p = max(group_y0 - pad_y, 0)
    x1p = min(group_x1 + pad_x, img_w)
    y1p = min(group_y1 + pad_y, img_h)

    # If our final padded box is degenerate, return None:
    if x1p <= x0p or y1p <= y0p:
        return None

    return (x0p, y0p, x1p, y1p)

    
def find_aligned_blob_group(img_color, min_area=10000, tol=10, pad=20):
    """
    Locate connected components ≥min_area, group those whose
    bottom-y are within tol pixels of each other. If ≥2 found,
    return their combined bbox padded by `pad`. Else None.
    """
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    _, th = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
    cnts, _ = cv2.findContours(th, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    boxes = []
    for c in cnts:
        x, y, w, h = cv2.boundingRect(c)
        if w * h < min_area:
            continue
        boxes.append((x, y, w, h))

    if len(boxes) < 2:
        return None

    bottoms = [y + h for x, y, w, h in boxes]
    avg_b = sum(bottoms) / len(bottoms)
    group = [b for b in boxes if abs((b[1] + b[3]) - avg_b) <= tol]

    if len(group) < 2:
        return None

    xs = [x for x, y, w, h in group] + [x + w for x, y, w, h in group]
    ys = [y for x, y, w, h in group] + [y + h for x, y, w, h in group]
    x0 = max(min(xs) - pad, 0)
    y0 = max(min(ys) - pad, 0)
    x1 = min(max(xs) + pad, img_color.shape[1])
    y1 = min(max(ys) + pad, img_color.shape[0])
    return (x0, y0, x1, y1)

def main(input_sheet, output_root, seq=105):
    api_key = get_valid_api_key()

    # ─── Prepare output directories ────────────────────────────────────────────
    today     = datetime.datetime.now().strftime('%m%d%Y')
    base_name = f"decal_output_{today}"
    out_dir   = os.path.join(output_root, base_name)
    idx = 1
    while os.path.exists(out_dir):
        out_dir = os.path.join(output_root, f"{base_name}_{idx}")
        idx += 1
    os.makedirs(out_dir)
    imgs_dir = os.path.join(out_dir, 'images')
    dbg_dir  = os.path.join(out_dir, 'debugging')
    cub_dir  = os.path.join(out_dir, 'cubiscan')
    tmp_dir  = os.path.join(out_dir, 'temp_pdfs')
    for d in (imgs_dir, dbg_dir, cub_dir, tmp_dir):
        os.makedirs(d, exist_ok=True)

    # ─── Load templates once ───────────────────────────────────────────────────
    template_sets = load_template_sets('templates')
    print(f"· Loaded {len(template_sets)} template sets for corner detection")

    # ─── Read parts list ───────────────────────────────────────────────────────
    df = pd.read_excel(input_sheet, dtype=str)
    df.columns = df.columns.str.upper()
    df.rename(columns={df.columns[0]: 'PART', df.columns[1]: 'TMS'}, inplace=True)
    records = []
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M') + '00'

    # ─── Loop over each row ────────────────────────────────────────────────────
    for i, row in df.iterrows():
        original_part = row['PART'].strip()
        tms           = row['TMS']

        print(f"[{i}] ➡️ Processing part={original_part}, TMS={tms}")

        # 1) Download PDF via API
        pdf_path = download_pdf_via_api(original_part, tmp_dir, api_key)
        if not pdf_path:
            print(f"    · No document found for {original_part}; skipping.")
            records.append({
                'ITEM_ID': original_part,
                # ... fill in the rest of your “skip” record fields ...
                'NET_LENGTH': 0,
                'NET_WIDTH': 0,
                'NET_HEIGHT': THICKNESS_IN,
                'IMAGE_FILE_NAME': '',
                'UPDATED': 'N',
                'TIME_STAMP': ts,
                'SITE_ID': SITE_ID,
                'FACTOR': FACTOR,
                # etc.
            })
            continue

        print(f"    · PDF downloaded → {pdf_path}")

        # a) Render first page to BGR image & build “ink” mask
        img = render_pdf_color_page(pdf_path, dpi=DPI)
        h_img, w_img = img.shape[:2]
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, blob = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)

        # b) Compute y0_art for bracket clamping
        ys = np.where(blob.sum(axis=1) > 0)[0]
        y0_art = int(ys.min()) if ys.size else 0
        y0_art = max(0, y0_art - 5)

        # c) Parse dimensions
        h_in, w_in = parse_dimensions_from_pdf(pdf_path)
        expected_ar = (w_in / h_in) if (h_in and w_in) else None

        # d) Crop logic (unchanged)…
        print("   · Attempting crop-mark → bracket → union-of-all-ink…")
        gray_for_rect = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        enclosed_rect = detect_enclosed_box(gray_for_rect, min_area=5000)
        if enclosed_rect:
            ex0, ey0, ex1, ey1 = enclosed_rect
            if ey0 > int(0.20 * h_img):
                enclosed_rect = None

        # … insert the rest of your cropping fallbacks here exactly as before …

        # f) Save the final crop as JPEG
        jpg_name = f"{tms}.{original_part}.{seq}.jpg"
        out_jpg = os.path.join(imgs_dir, jpg_name)
        cv2.imwrite(out_jpg, crop_img)
        print(f"   · Writing JPEG → {out_jpg}")

        # g) Clean up
        os.remove(pdf_path)
        time.sleep(STEP_DELAY)

        # 8) Record row
        vol  = h_in * w_in * THICKNESS_IN
        wgt  = vol * MATERIAL_DENSITY
        records.append({
            'ITEM_ID':         original_part,
            'NET_LENGTH':      h_in,
            'NET_WIDTH':       w_in,
            'NET_HEIGHT':      THICKNESS_IN,
            'NET_WEIGHT':      wgt,
            'NET_VOLUME':      vol,
            'IMAGE_FILE_NAME': jpg_name,
            'UPDATED':         'Y',
            'TIME_STAMP':      ts,
            'SITE_ID':         SITE_ID,
            'FACTOR':          FACTOR,
            # etc.
        })
        print(f"[{i}] ✅ Done\n")

    # ─── Tear down & write CSV ───────────────────────────────────────────────────
    shutil.rmtree(tmp_dir, ignore_errors=True)
    df_out = pd.DataFrame(records)
    cols = [
        'ITEM_ID','ITEM_TYPE','DESCRIPTION','NET_LENGTH','NET_WIDTH','NET_HEIGHT',
        'NET_WEIGHT','NET_VOLUME','NET_DIM_WGT','DIM_UNIT','WGT_UNIT','VOL_UNIT',
        'FACTOR','SITE_ID','TIME_STAMP','OPT_INFO_1','OPT_INFO_2','OPT_INFO_3',
        'OPT_INFO_4','OPT_INFO_5','OPT_INFO_6','OPT_INFO_7','OPT_INFO_8',
        'IMAGE_FILE_NAME','UPDATED'
    ]
    df_out = df_out.reindex(columns=cols)
    out_csv = os.path.join(cub_dir, f"{SITE_ID}_{ts}.csv")
    df_out.to_csv(out_csv, index=False)
    print("All done →", out_csv)

    # ── (Legacy block #6: not usually reached) ─────────────────────────────────────────
    print("    · Rendering page to image…")
    img_color = render_pdf_color_page(pdf_path, dpi=DPI)
    h_img, w_img = img_color.shape[:2]

    # 6a) Try 4-corner bracket crop with all template sets
    try:
        print("    · Selecting best crop box…")
        x0, y0, x1, y1 = select_best_crop_box(img_color, template_sets)
        print(f"    · Bracket crop box: {(x0, y0, x1, y1)}")
        crop_region = img_color[y0:y1, x0:x1]
    except Exception as e:
        print(f"    · Template crop failed ({e}); falling back to blob/full-page…")
        gray2 = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray2, 250, 255, cv2.THRESH_BINARY_INV)
        cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if cnts:
            bx, by, bw, bh = cv2.boundingRect(max(cnts, key=cv2.contourArea))
            print(f"    · Blob crop box: {(bx, by, bx + bw, by + bh)}")
            crop_region = img_color[by:by + bh, bx:bx + bw]
        else:
            m = int(0.01 * min(h_img, w_img))
            print(f"    · Full-page margin crop: {(m, m, w_img - m, h_img - m)}")
            crop_region = img_color[m:h_img - m, m:w_img - m]

    # 6c) Extract that region
    region = img_color[y0:y1, x0:x1]
    rh, rw = crop_region.shape[:2]

    # 6d) Legacy multi-layer? 3 bands side-by-side
    if rw > rh * 1.8:
        print("    · Detected legacy multi-layer → slicing bands…")
        third = rw // 3
        green = crop_region[:, third:2 * third]
        black = crop_region[:, 2 * third:3 * third]

        def recolor(band, bgr):
            g = cv2.cvtColor(band, cv2.COLOR_BGR2GRAY)
            _, mask = cv2.threshold(g, 250, 255, cv2.THRESH_BINARY_INV)
            fill = np.zeros_like(band); fill[:] = bgr
            return np.where(mask[:, :, None] > 0, fill, band)

        band_g = recolor(green, COLOR_MAP['green'])
        band_b = recolor(black, COLOR_MAP['black'])
        stacked = band_b.copy()
        mask_g = cv2.cvtColor(band_g, cv2.COLOR_BGR2GRAY) < 250
        for c in range(3):
            stacked[:, :, c] = np.where(mask_g, band_g[:, :, c], stacked[:, :, c])
        crop = stacked
    else:
        crop = crop_region

    # ── 7) Save JPEG ─────────────────────────────────────────
    jpg_name = f"{tms}.{part}.{seq}.jpg"
    out_jpg = os.path.join(imgs_dir, jpg_name)
    print(f"    · Writing JPEG → {out_jpg}")
    cv2.imwrite(out_jpg, crop)
    time.sleep(STEP_DELAY)

    # ── 8) Compute dims, volume, weight ────────────────────
    h_in2, w_in2 = parse_dimensions_from_pdf(pdf_path)
    vol2 = h_in2 * w_in2 * THICKNESS_IN
    wgt2 = vol2 * MATERIAL_DENSITY
    dim_wgt2 = vol2 / FACTOR

    # ── 9) Clean up ─────────────────────────────────────────
    os.remove(pdf_path)

    # ── 10) Record ──────────────────────────────────────────
    records.append({
        'ITEM_ID':         part,
        'ITEM_TYPE':       '',
        'DESCRIPTION':     '',
        'NET_LENGTH':      h_in2,
        'NET_WIDTH':       w_in2,
        'NET_HEIGHT':      THICKNESS_IN,
        'NET_WEIGHT':      wgt2,
        'NET_VOLUME':      vol2,
        'NET_DIM_WGT':     dim_wgt2,
        'DIM_UNIT':        'in',
        'WGT_UNIT':        'lb',
        'VOL_UNIT':        'in',
        'FACTOR':          FACTOR,
        'SITE_ID':         SITE_ID,
        'TIME_STAMP':      ts,
        'OPT_INFO_1':      '',
        'OPT_INFO_2':      'Y',
        'OPT_INFO_3':      'N',
        'OPT_INFO_4':      '',
        'OPT_INFO_5':      '',
        'OPT_INFO_6':      '',
        'OPT_INFO_7':      '',
        'OPT_INFO_8':      0,
        'IMAGE_FILE_NAME': '',
        'UPDATED':         'Y'
    })
    print(f"[{i}] ✅ Done\n")

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    
    sheet = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel files", "*.xlsx")]
    )
    
    out_root = filedialog.askdirectory(
        title="Select output directory"
    )
    main(sheet, out_root, seq=105)

