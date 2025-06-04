import os
import re
import glob
import time
import datetime
import shutil
import requests

import cv2
import fitz       # PyMuPDF
import numpy as np
import pandas as pd
import pdfplumber

import tkinter as tk
from urllib.parse import urljoin
from tkinter import filedialog, simpledialog

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium.common.exceptions import TimeoutException

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

# ── Utility Functions ─────────────────────────────────────────────────────────
def choose_chrome_profile():
    """
    Two‐button chooser:
      • Default → %LOCALAPPDATA%/Google/Chrome/User Data/Default
      • Custom  → prompt for folder (blank ⇒ new session)
    Returns the chosen path, or None.
    """
    # hidden root
    root = tk.Tk()
    root.withdraw()

    choice = {'profile': None}
    dlg = tk.Toplevel(root)
    dlg.title("Select Chrome Profile")
    dlg.geometry("360x140")
    dlg.resizable(False, False)

    tk.Label(
        dlg,
        text=(
            "Choose a Chrome profile:\n\n"
            "• Default: all Chrome windows must be closed first\n"
            "  (uses your existing Default profile folder)\n"
            "• Custom: browse or type a folder (blank ⇒ new session)"
        ),
        justify="left",
        wraplength=350,
        padx=10, pady=10
    ).pack()

    def _use_default():
        local = os.environ.get("LOCALAPPDATA")
        if local:
            choice['profile'] = os.path.join(
                local, "Google", "Chrome", "User Data", "Default"
            )
        else:
            choice['profile'] = None
        dlg.destroy()

    def _use_custom():
        p = simpledialog.askstring(
            "Custom Profile",
            "Enter full path to Chrome profile folder\n(leave blank for new):",
            parent=dlg
        )
        choice['profile'] = p or None
        dlg.destroy()

    frm = tk.Frame(dlg, pady=5)
    frm.pack()
    tk.Button(frm, text="Default", width=14, command=_use_default).pack(side="left", padx=8)
    tk.Button(frm, text="Custom",  width=14, command=_use_custom).pack(side="left", padx=8)

    # wait for the dialog to go away
    root.wait_window(dlg)
    root.destroy()
    return choice['profile']
        
def strip_gt_suffix(part: str) -> str:
    """
    Remove a trailing “GT” from the part number, if present,
    but leave any other letters (e.g. DU, FR, etc.) untouched.
    """
    if part.endswith("GT"):
        return part[:-2]
    return part

def clear_filters(driver, timeout=10):
    """Click any existing remove-filter buttons, wait for them to disappear."""
    remove_btns = driver.find_elements(By.CSS_SELECTOR, 'button.a-IRR-button--remove')
    for b in remove_btns:
        try:
            b.click()
        except:
            pass
    WebDriverWait(driver, timeout).until(
        EC.invisibility_of_element_located((By.CSS_SELECTOR, 'button.a-IRR-button--remove'))
    )

def init_driver(download_dir, profile_dir=None, headless=False):
    """Configure Chrome for headless PDF downloads into download_dir."""
    opts = Options()
    if profile_dir:
        # point at both the user‐data dir and the Default subfolder
        opts.add_argument(f"--user-data-dir={profile_dir}")
        opts.add_argument("--profile-directory=Default")
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        # force external download (not in‐browser)
        "plugins.always_open_pdf_externally": True,
    }
    opts.add_experimental_option("prefs", prefs)
    if headless:
        opts.add_argument("--headless")
    return webdriver.Chrome(options=opts)

def download_pdf_for_part(part, tmp_dir, driver, base_url):
    """
    Search for 'part' in the library, download its PDF into tmp_dir, and return the local path.
    - If “No data found” appears, return None.
    - If the sign-in lightbox appears, raise to trigger safe_download retry.
    - If no exact match <td> is found, return None (skip).
    - Otherwise download the PDF (inline URL or click+poll).
    """
    wait = WebDriverWait(driver, 20)

    # 1) Search for the part
    driver.get(base_url)
    wait.until(EC.presence_of_element_located((By.ID, 'docLibContainer_search_field')))
    clean = strip_gt_suffix(part)
    fld = wait.until(EC.element_to_be_clickable((By.ID, 'docLibContainer_search_field')))
    fld.clear()
    fld.send_keys(clean)
    driver.find_element(By.ID, 'docLibContainer_search_button').click()

    # 1a) Bail if “No data found”
    time.sleep(1)
    if driver.find_elements(By.CSS_SELECTOR, 'div.a-IRR-noDataMsg'):
        return None

    # 1b) Detect locked-out sign-in lightbox
    if driver.find_elements(By.CSS_SELECTOR, 'div.sign-in-box.ext-sign-in-box'):
        # Only this exact condition triggers a restart
        raise WebDriverException("Session locked, need to re-login")

    # 2) Try to click the row’s “Documents” button for an exact match.
    #    If no exact <td> for 'clean' appears, bail out immediately.
    try:
        td = driver.find_element(
            By.XPATH,
            f"//td[normalize-space(text())='{clean}']"
        )
    except NoSuchElementException:
        return None

    tr = td.find_element(By.XPATH, "./ancestor::tr")
    docs_btn = tr.find_element(By.XPATH, ".//button[contains(., 'Documents')]")
    docs_btn.click()

    # 3) Switch into the Documents iframe
    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.CSS_SELECTOR, "iframe[title='Documents']")
    ))

    # 4) Try inline URL download
    dl = wait.until(EC.presence_of_element_located((By.ID, "downloadBtn")))
    onclick = dl.get_attribute("onclick") or ""
    m = re.search(r"doDownload\('([^']+)','([^']+)'\)", onclick)
    if m:
        url_frag, filename = m.groups()
        full_url = urljoin(base_url, url_frag)
        cookies = {c["name"]: c["value"] for c in driver.get_cookies()}
        resp = requests.get(full_url, cookies=cookies, timeout=60)
        resp.raise_for_status()
        out_path = os.path.join(tmp_dir, filename)
        with open(out_path, "wb") as f:
            f.write(resp.content)
        driver.switch_to.default_content()
        return out_path

    # 5) Fallback: click + poll for the PDF file
    dl.click()
    driver.switch_to.default_content()
    return wait_for_pdf(tmp_dir, clean, timeout=DOWNLOAD_TIMEOUT)

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
        
def find_union_of_ink_contours(img_color, min_area=500, pad_pct=0.05, proximity_px=50):
    """
    From a BGR image (300 dpi) of a PDF page, build a mask of any non-white pixels ("ink"),
    find all contours whose area is >= min_area, then cluster together any contours whose
    bounding-boxes lie within `proximity_px` pixels (horizontally) and overlap vertically.
    Finally, take the union of that cluster’s bounding box, pad it outward by pad_pct,
    and return (x0, y0, x1, y1). Returns None if no contour of sufficient size exists.

    - img_color: full-resolution (300 dpi) BGR numpy array.
    - min_area: ignore any contour smaller than this (in pixels²). Default=500.
    - pad_pct: fraction of the union‐box’s width/height to expand OUTWARDS on each side. Default=0.05 (5%).
    - proximity_px: maximum horizontal gap (in pixels) for merging two contours into the same group. Default=50.
    """
    import cv2
    import numpy as np

    # 1) Build an “ink” mask = any pixel that isn’t nearly white.
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)

    # 2) Find all external contours on that mask
    cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None

    # 3) Keep only contours whose area >= min_area
    boxes = []
    for c in cnts:
        area = cv2.contourArea(c)
        if area < min_area:
            continue
        x, y, w, h = cv2.boundingRect(c)
        boxes.append((x, y, w, h))
    if not boxes:
        return None

    # 4) Sort boxes by their left‐edge x-coordinate
    boxes = sorted(boxes, key=lambda b: b[0])

    # 5) Find the single largest‐area contour to seed our cluster
    areas = [w * h for (x, y, w, h) in boxes]
    largest_idx = int(np.argmax(areas))
    gx, gy, gw, gh = boxes[largest_idx]

    # Group bounding‐box so far = that largest contour
    group_x0 = gx
    group_y0 = gy
    group_x1 = gx + gw
    group_y1 = gy + gh

    absorbed = True
    used = set([largest_idx])

    # 6) Grow the cluster by absorbing any box whose x-range is within proximity_px
    #    AND whose y-range overlaps with the current group vertically.
    while absorbed:
        absorbed = False
        for i, (x, y, w, h) in enumerate(boxes):
            if i in used:
                continue

            # a) Horizontal‐gap check: how far is this box from our current group?
            bx0, bx1 = x, x + w
            if bx1 < group_x0:
                dist_horiz = group_x0 - bx1
            elif bx0 > group_x1:
                dist_horiz = bx0 - group_x1
            else:
                dist_horiz = 0

            # b) Vertical‐overlap check: does y..y+h overlap group_y0..group_y1?
            vy0, vy1 = y, y + h
            overlap_vert = not (vy1 < group_y0 or vy0 > group_y1)

            if dist_horiz <= proximity_px and overlap_vert:
                # absorb this contour
                used.add(i)
                absorbed = True
                group_x0 = min(group_x0, x)
                group_y0 = min(group_y0, y)
                group_x1 = max(group_x1, x + w)
                group_y1 = max(group_y1, y + h)

    # 7) Pad that final union‐box by pad_pct (uniformly) and clamp to image bounds
    img_h, img_w = img_color.shape[:2]
    rect_w = group_x1 - group_x0
    rect_h = group_y1 - group_y0
    pad_x = int(rect_w * pad_pct)
    pad_y = int(rect_h * pad_pct)

    x0p = max(group_x0 - pad_x, 0)
    y0p = max(group_y0 - pad_y, 0)
    x1p = min(group_x1 + pad_x, img_w)
    y1p = min(group_y1 + pad_y, img_h)

    if x1p <= x0p or y1p <= y0p:
        return None

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

def wait_for_pdf(tmp_dir, part, timeout=DOWNLOAD_TIMEOUT):
    """
    Poll tmp_dir until a PDF whose filename contains `part` appears
    and its size is >0. Returns its full path.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        for fn in os.listdir(tmp_dir):
            if fn.lower().endswith('.pdf') and part in fn:
                full = os.path.join(tmp_dir, fn)
                if os.path.getsize(full) > 0:
                    return full
        time.sleep(0.5)
    raise RuntimeError(f"Timeout waiting for PDF containing '{part}'")

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

def safe_download(part, tmp_dir, driver, base_url, profile):
    """
    Attempt to download PDF for `part`. On WebDriverException or locked-out lightbox,
    fully quit Chrome, wait (1 minute for the first retry; 2 minutes thereafter),
    re-initialize, wait for manual re-login, then retry exactly this `part`.

    Returns (pdf_path, driver) once the PDF is downloaded, or (None, driver) if “No data found”.
    """
    # We will attempt a 1-minute wait first; if that fails, switch to 2-minute
    # intervals forever after (rather than continually growing).
    first_wait = 120       # 1 minute
    subsequent_wait = 200  # 3 minutes 20 seconds
    attempt = 0            # counter

    while True:
        try:
            # If driver is None, create it fresh (first iteration or after quit)
            if driver is None:
                driver = init_driver(tmp_dir, profile_dir=profile, headless=False)
                driver.get(base_url)
                # Wait until the IRR search field is present (manual login if needed)
                WebDriverWait(driver, 300).until(
                    EC.presence_of_element_located((By.ID, 'docLibContainer_search_field'))
                )

            # Attempt to download the PDF normally
            pdf_path = download_pdf_for_part(part, tmp_dir, driver, base_url)

            # If download_pdf_for_part returns None → no document found
            if not pdf_path:
                return None, driver

            return pdf_path, driver

        except WebDriverException as e:
            msg = str(e)
            # Specifically handle our locked-out lightbox by forcing a re-login
            if "Session locked" in msg or "lightbox" in msg or "sign-in-box" in msg.lower():
                wait_time = first_wait if attempt == 0 else subsequent_wait
                print(f"    · Locked out ({msg}); closing browser and retrying in {wait_time}s…")
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = None
                attempt += 1
                time.sleep(wait_time)

                # Re-open fresh browser and wait for manual login again
                driver = init_driver(tmp_dir, profile_dir=profile, headless=False)
                driver.get(base_url)
                print("    · Waiting for library page to become available…")
                try:
                    WebDriverWait(driver, 300).until(
                        EC.presence_of_element_located((By.ID, 'docLibContainer_search_field'))
                    )
                    print("    · Library page detected; resuming download for part:", part)
                except TimeoutException:
                    # If still locked out, we’ll loop back and retry after another wait_time
                    print(f"    · Still locked out after {wait_time}s; will retry.")
                    continue

            else:
                # Any other WebDriverException (e.g. Chrome crashed). Retry after a wait.
                wait_time = first_wait if attempt == 0 else subsequent_wait
                print(f"    · WebDriverException ({msg}); closing browser and retrying in {wait_time}s…")
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = None
                attempt += 1
                time.sleep(wait_time)
                continue

        except TimeoutException as te:
            # Timeout waiting for PDF or for search field – treat like other exceptions
            wait_time = first_wait if attempt == 0 else subsequent_wait
            print(f"    · TimeoutException ({te}); closing browser and retrying in {wait_time}s…")
            try:
                driver.quit()
            except Exception:
                pass
            driver = None
            attempt += 1
            time.sleep(wait_time)
            continue

def _wait_for_library_or_lock(driver, poll_interval=5):
    """
    Block until we see the real library search field appear.
    If the locked-out lightbox appears first, ignore it and keep polling.
    i.e. we never timeout here; we only return once '#docLibContainer_search_field' is found.
    """
    from selenium.common.exceptions import TimeoutException

    while True:
        # 1) If the library search field is present, we’re done.
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.ID, 'docLibContainer_search_field'))
            )
            return  # library is available again
        except TimeoutException:
            pass

        # 2) If the lockout screen is present, ignore it and keep waiting
        #    (we do not raise here; just give the user more time to log in).
        if driver.find_elements(By.CSS_SELECTOR, 'div.sign-in-box.ext-sign-in-box'):
            print("    · Found lockout box—still waiting for manual re-login…")
            # (do NOT quit; just keep looping)

        # 3) Otherwise, sleep and try again.
        time.sleep(poll_interval)
        
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

def main(input_sheet, output_root, base_url, profile=None, seq=105):
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

    # ─── Launch browser & load library ─────────────────────────────────────────
    driver = init_driver(tmp_dir, profile_dir=profile)
    print("· Browser launched")
    driver.get(base_url)
    wait = WebDriverWait(driver, 20)
    wait.until(EC.presence_of_element_located((By.ID, 'docLibContainer_search_field')))
    print("· Library page ready")

    # ─── Load templates once ───────────────────────────────────────────────────
    template_sets = load_template_sets('templates')
    print(f"· Loaded {len(template_sets)} template sets for corner detection")

    # ─── Read parts list ───────────────────────────────────────────────────────
    df = pd.read_excel(input_sheet, dtype=str)
    df.columns = df.columns.str.upper()
    df.rename(columns={df.columns[0]:'PART', df.columns[1]:'TMS'}, inplace=True)
    records = []
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M') + '00'

    # ─── Loop over each row ────────────────────────────────────────────────────
    for i, row in df.iterrows():
        original_part = row['PART'].strip()
        tms          = row['TMS']
        search_part  = strip_gt_suffix(original_part)

        print(f"[{i}] ➡️ Processing part={search_part} (orig={original_part}), TMS={tms}")

        try:
            # 1) Download PDF (auto-retries on WebDriver errors)
            result = safe_download(search_part, tmp_dir, driver, base_url, profile)
            if isinstance(result, tuple):
                pdf_path, driver = result
            else:
                pdf_path = result

            # 1a) skip if no document for this part
            if not pdf_path:
                print(f"    · No document found for {original_part}; skipping.")
                records.append({
                    'ITEM_ID':         original_part,
                    'ITEM_TYPE':       '',
                    'DESCRIPTION':     '',
                    'NET_LENGTH':      0,
                    'NET_WIDTH':       0,
                    'NET_HEIGHT':      THICKNESS_IN,
                    'NET_WEIGHT':      0,
                    'NET_VOLUME':      0,
                    'NET_DIM_WGT':     0,
                    'DIM_UNIT':        'in',
                    'WGT_UNIT':        'lb',
                    'VOL_UNIT':        'in',
                    'FACTOR':          FACTOR,
                    'SITE_ID':         SITE_ID,
                    'TIME_STAMP':      ts,
                    'OPT_INFO_2':      'N',
                    'OPT_INFO_3':      'N',
                    'OPT_INFO_8':      0,
                    'IMAGE_FILE_NAME': '',
                    'UPDATED':         'N'
                })
                continue

            print(f"    · PDF downloaded → {pdf_path}")

            # a)  Render first page to BGR image & build “ink” mask
            img = render_pdf_color_page(pdf_path, dpi=DPI)
            h_img, w_img = img.shape[:2]
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, blob = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)

            # b)  Find the first row with any ink (to compute y0_art for bracket clamping)
            ys = np.where(blob.sum(axis=1) > 0)[0]
            y0_art = int(ys.min()) if ys.size else 0
            PAD_TOP = 5
            y0_art = max(0, y0_art - PAD_TOP)

            # c)  Parse (h × w) from PDF so we have expected aspect ratio
            h_in, w_in = parse_dimensions_from_pdf(pdf_path)
            expected_ar = (w_in / h_in) if (h_in and w_in) else None

            # d)  Attempt crop‐mark → bracket → union‐of‐all‐ink (with 5% padding)
            print("   · Attempting crop‐mark → bracket → union‐of‐all‐ink…")

            # Prepare grayscale for crop‐mark detection:
            gray_for_rect = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            enclosed_rect = detect_enclosed_box(gray_for_rect, min_area=5000)

            crop_img = None

            # 1) If we found a rounded‐corner border, expand it uniformly by 5% of the smaller side:
            if enclosed_rect:
                ex0, ey0, ex1, ey1 = enclosed_rect
                rect_w = ex1 - ex0
                rect_h = ey1 - ey0

                # 5% of the smaller dimension → uniform padding on all sides
                pad = int(min(rect_w, rect_h) * 0.05)

                x0c = max(ex0 - pad, 0)
                y0c = max(ey0 - pad, 0)
                x1c = min(ex1 + pad, w_img)
                y1c = min(ey1 + pad, h_img)

                crop_img = img[y0c:y1c, x0c:x1c]
                print(f"   · Using crop‐mark + 5% pad: {(x0c, y0c, x1c, y1c)}")

            else:
                # 2) Fallback #1: try bracket‐template detection
                try:
                    x0b, y0b, x1b, y1b = select_best_crop_box(
                        img,
                        template_sets,
                        expected_ratio=expected_ar
                    )
                    b_w = x1b - x0b
                    b_h = y1b - y0b

                    # 5% of the smaller dimension → uniform padding on all sides
                    pad = int(min(b_w, b_h) * 0.05)

                    x0c = max(x0b - pad, 0)
                    y0c = max(y0b - pad, 0)
                    x1c = min(x1b + pad, w_img)
                    y1c = min(y1b + pad, h_img)

                    crop_img = img[y0c:y1c, x0c:x1c]
                    print(f"   · Using bracket‐crop + 5% pad: {(x0c, y0c, x1c, y1c)}")

                    print("   · Bracket‐template crop (dim‐guided)…")

                except RuntimeError as e:
                    # ── NEW FALLBACK BLOCK STARTS HERE ──
                    print(f"   · No valid bracket candidates ({e}); falling back…")

                    # 1) Fallback #1: aligned‐blobs group (if at least two blobs share a common baseline)
                    grp = find_aligned_blob_group(img, min_area=5000, tol=10, pad=20)
                    if grp:
                        x0g, y0g, x1g, y1g = grp
                        print(f"   · Aligned blob group crop: {grp}")
                        # Crop 20px extra inside image bounds
                        crop_img = img[
                            max(0, y0g - 20) : min(y1g + 20, h_img),
                            max(0, x0g - 20) : min(x1g + 20, w_img)
                        ]

                    else:
                        # 2) Fallback #2: enclosed rectangle (rounded border)
                        gray_fb = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                        rect2 = detect_enclosed_box(gray_fb, min_area=5000)
                        if rect2:
                            x0e2, y0e2, x1e2, y1e2 = rect2
                            print(f"   · Enclosed rectangle crop: {(x0e2, y0e2, x1e2, y1e2)}")
                            crop_img = img[y0e2:y1e2, x0e2:x1e2]

                        else:
                            # 3) Fallback #3: “Full‐logo” via “…mm” line (top‐portion crop)
                            y_crop = crop_full_logo(pdf_path, dpi=DPI)
                            if y_crop:
                                print(f"   · Full‐logo crop at y={y_crop}px")
                                crop_img = img[:y_crop, :]

                            else:
                                # 4) Fallback #4: UNION‐OF‐INK‐CONTOURS (catch disconnected pieces)
                                print("   · No enclosed rectangle; attempting union of ALL ink contours…")
                                # Call our new, “clustered” union‐of‐ink function:
                                rect_union = find_union_of_ink_contours(
                                    img,
                                    min_area=500,
                                    pad_pct=0.05,
                                    proximity_px=50
                                )
                                if rect_union:
                                    x0u, y0u, x1u, y1u = rect_union
                                    print(f"   · Union‐of‐ink‐contours crop: {(x0u, y0u, x1u, y1u)}")
                                    crop_img = img[y0u:y1u, x0u:x1u]

                                else:
                                    # 5) Fallback #5: as a last resort, just do a 1% full‐page margin crop
                                    margin = int(0.01 * min(h_img, w_img))
                                    print("   · Union‐of‐ink failed; doing full‐page margin crop")
                                    crop_img = img[
                                        margin : h_img - margin,
                                        margin : w_img - margin
                                    ]
                # ── At this point, `crop_img` is set (either by “using crop‐mark,” “bracket,” or one of the fallbacks) ──

            # e)  Print final crop size, save JPEG, clean up, and record:
            print(f"   · Final crop size: {crop_img.shape[1]}×{crop_img.shape[0]} (w×h)")

            # f)  Save the final crop as a JPEG
            jpg_name = f"{tms}.{original_part}.{seq}.jpg"
            out_jpg = os.path.join(imgs_dir, jpg_name)
            print(f"   · Writing JPEG → {out_jpg}")
            cv2.imwrite(out_jpg, crop_img)

            # g)  Clean up the temporary PDF
            print("   · Removing temp PDF")
            os.remove(pdf_path)
            time.sleep(STEP_DELAY)

            # ─── 8) Record row (dimensions already parsed) ──────────────────────────
            vol  = h_in * w_in * THICKNESS_IN
            wgt  = vol * MATERIAL_DENSITY
            dimw = vol / FACTOR

            records.append({
                'ITEM_ID':         original_part,
                'ITEM_TYPE':       '',
                'DESCRIPTION':     '',
                'NET_LENGTH':      h_in,
                'NET_WIDTH':       w_in,
                'NET_HEIGHT':      THICKNESS_IN,
                'NET_WEIGHT':      wgt,
                'NET_VOLUME':      vol,
                'NET_DIM_WGT':     dimw,
                'DIM_UNIT':        'in',
                'WGT_UNIT':        'lb',
                'VOL_UNIT':        'in',
                'FACTOR':          FACTOR,
                'SITE_ID':         SITE_ID,
                'TIME_STAMP':      ts,
                'OPT_INFO_2':      'Y',
                'OPT_INFO_3':      'N',
                'OPT_INFO_8':      0,
                'IMAGE_FILE_NAME': jpg_name,
                'UPDATED':         'Y'
            })
            print(f"[{i}] ✅ Done\n")
            time.sleep(STEP_DELAY)

        except Exception as e:
            print(f"[{i}] ❌ ERROR: {e}")
            with open(os.path.join(dbg_dir, 'errors.log'), 'a', encoding='utf-8') as f:
                f.write(f"{original_part}: {e}\n")
            continue

    # ─── Tear down & write CSV ───────────────────────────────────────────────────
    try:
        driver.quit()
    except:
        pass
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
    # 0) pick profile first
    print("→ About to ask for Chrome profile…")
    profile = choose_chrome_profile()
    print("→ Using profile:", profile or "<new session>")

    # 1) pick parts sheet
    sheet = filedialog.askopenfilename(
        title='Select parts sheet',
        filetypes=[('Excel/CSV','*.xlsx *.xls *.csv')]
    )
    if not sheet:
        print("No sheet selected, exiting.")
        exit()

    # 2) pick output folder
    out_root = filedialog.askdirectory(title='Select output folder')
    if not out_root:
        print("No output folder selected, exiting.")
        exit()

    # 3) enter the library URL
    url = simpledialog.askstring(
        'Document Library URL',
        'Enter the library URL:'
    )
    if not url:
        print("No URL provided, exiting.")
        exit()

    # 4) run!
    main(
        input_sheet=sheet,
        output_root=out_root,
        base_url=url,
        profile=profile,
        seq=105
    )
