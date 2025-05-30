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
STEP_DELAY      = 0.5 #whenever you need a short delay insert: time.sleep(STEP_DELAY)

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
        try: b.click()
        except: pass
    WebDriverWait(driver, timeout).until(
        EC.invisibility_of_element_located((By.CSS_SELECTOR,'button.a-IRR-button--remove'))
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
    Search for `part` in the library and download its PDF.
    - If we never even see the search field → raise WebDriverException("Session locked…")
    - If we see the search field but no matching row → return None
    - Otherwise download (inline or click+poll) and return the local PDF path.
    """
    wait = WebDriverWait(driver, 20)

    # 1) Go to library home
    driver.get(base_url)

    # 1a) Wait for the search field.  If it never appears → we're locked out
    try:
        wait.until(EC.presence_of_element_located((By.ID, 'docLibContainer_search_field')))
    except TimeoutException:
        raise WebDriverException("Session locked, need to re-login")

    clean = strip_gt_suffix(part)

    # 1b) Enter your search term
    fld = wait.until(EC.element_to_be_clickable((By.ID, 'docLibContainer_search_field')))
    fld.clear()
    fld.send_keys(clean)
    driver.find_element(By.ID, 'docLibContainer_search_button').click()

    # short pause for “no data found” to appear
    time.sleep(1)
    if driver.find_elements(By.CSS_SELECTOR, 'div.a-IRR-noDataMsg'):
        return None

    # 2) Look up the row for this part.  If it never shows → no document
    try:
        td = WebDriverWait(driver, 5).until(  # shorter wait here
            EC.presence_of_element_located((
                By.XPATH,
                f"//td[normalize-space(text())='{clean}']"
            ))
        )
    except TimeoutException:
        # no matching part on this page → give up and move on
        driver.switch_to.default_content()
        return None

    # 3) Click the “Documents” button in that row
    tr = td.find_element(By.XPATH, "./ancestor::tr")
    docs_btn = tr.find_element(By.XPATH, ".//button[contains(., 'Documents')]")
    docs_btn.click()

    # 4) Switch into the Documents iframe
    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.CSS_SELECTOR, "iframe[title='Documents']")
    ))

    # 5) Try inline‐URL download
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

    # 6) Fallback: click + poll
    dl.click()
    driver.switch_to.default_content()
    return wait_for_pdf(tmp_dir, clean, timeout=DOWNLOAD_TIMEOUT)

def render_pdf_color_page(pdf_path, dpi=300):
    """Load the first page of PDF at `dpi` into a BGR numpy image."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    scale = dpi/72
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8)
    img = arr.reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        return cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    else:
        return cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

def extract_color_label(pdf_path):
    """
    Stub: read text near the top of the PDF and return one of
    'green','yellow','red','blue','black'.  Use pdfplumber or regex.
    """
    # TODO: open pdfplumber, search for the crop-band label and return lowercase key
    return 'green'
    
def load_template_sets(root='templates'):
    """
    Look under root/set1…set6 for quad-templates.
    Returns a list of (templates, offsets) for each set.
    """
    quad_names = ['top_left','top_right','bottom_left','bottom_right']
    sets = []
    for folder in sorted(glob.glob(os.path.join(root, 'set*'))):
        tpl_dict, off_dict = {}, {}
        for quad in quad_names:
            for ext in ('jpg','jpeg'):
                path = os.path.join(folder, f"{quad}.{ext}")
                if os.path.exists(path):
                    img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
                    edges = cv2.Canny(img, 50, 150)
                    h, w = edges.shape
                    tpl_dict[quad] = edges
                    # offsets inside the small crop marks
                    if quad=='top_left':      off_dict[quad] = (w-1, h-1)
                    elif quad=='top_right':   off_dict[quad] = (0,   h-1)
                    elif quad=='bottom_left': off_dict[quad] = (w-1, 0)
                    else:                     off_dict[quad] = (0,   0)
                    break
        if len(tpl_dict)==4:
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
        x1,y1,x2,y2 = rois[q]
        roi = img_gray[y1:y2, x1:x2]
        edges_roi = cv2.Canny(roi, 50, 150)
        res = cv2.matchTemplate(edges_roi, templates[q], cv2.TM_CCOEFF_NORMED)
        _, _, _, loc = cv2.minMaxLoc(res)
        offx, offy = offsets[q]
        dets[q] = (x1 + loc[0] + offx, y1 + loc[1] + offy)
    return dets

def crop_blob_bbox(img_gray):
    """Return bounding box (x0,y0,x1,y1) of the largest dark blob."""
    _,th = cv2.threshold(img_gray, 250, 255, cv2.THRESH_BINARY_INV)
    cnts,_ = cv2.findContours(th, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None
    c = max(cnts, key=cv2.contourArea)
    x,y,w,h = cv2.boundingRect(c)
    return (x,y,x+w,y+h)

def rect_intersection(a, b):
    """Intersection area of two rects a=(x0,y0,x1,y1)."""
    x0 = max(a[0], b[0]); y0 = max(a[1], b[1])
    x1 = min(a[2], b[2]); y1 = min(a[3], b[3])
    if x1<=x0 or y1<=y0: return 0
    return (x1-x0)*(y1-y0)

def detect_best_crop(img_color, template_sets):
    """
    Try each template-set; score by how much of the main blob
    sits inside the resulting crop. Return the best (x0,y0,x1,y1).
    """
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    blob_box = crop_blob_bbox(gray) or (0,0,gray.shape[1],gray.shape[0])
    blob_area = (blob_box[2]-blob_box[0])*(blob_box[3]-blob_box[1])

    best_score, best_rect = -1, (0,0,gray.shape[1],gray.shape[0])
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

        rect = (x0,y0,x1,y1)
        score = rect_intersection(blob_box, rect) / float(blob_area or 1)
        if score > best_score and score > 0.8:
            best_score, best_rect = score, rect

    if best_score < 0.8:
        # fallback
        x0,y0,x1,y1 = blob_box
        if x1<=x0 or y1<=y0:
            h,w = gray.shape
            m = int(0.01*min(h,w))
            x0,y0,x1,y1 = m,m,w-m,h-m
        best_rect = (x0,y0,x1,y1)

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
    blob = (blob>0).astype(np.uint8)
    H,W = gray.shape
    cands = []
    for tpl, offs in template_sets:
        try:
            corners = detect_with_one_set(gray, tpl, offs)
            x0 = int((corners['top_left'][0]+corners['bottom_left'][0])/2)
            x1 = int((corners['top_right'][0]+corners['bottom_right'][0])/2)
            y0 = int((corners['top_left'][1]+corners['top_right'][1])/2)
            y1 = int((corners['bottom_left'][1]+corners['bottom_right'][1])/2)
        except:
            continue
        # penalty = ink on 5-pixel wide border
        e=5
        top    = blob[y0:y0+e, x0:x1]
        bot    = blob[y1-e:y1, x0:x1]
        left   = blob[y0:y1, x0:x0+e]
        right  = blob[y0:y1, x1-e:x1]
        penalty = float(top.sum()+bot.sum()+left.sum()+right.sum())/((x1-x0)*(y1-y0))
        if penalty>penalty_thresh or x1<=x0 or y1<=y0:
            continue
        cands.append(( (x1-x0)/(y1-y0), (x0,y0,x1,y1) ))
    if not cands:
        return []
    # group by ratio (within 5%)
    base_ratio = cands[0][0]
    boxes=[]
    for ratio, box in cands:
        if abs(ratio-base_ratio)/base_ratio<0.05:
            if all(rect_intersection(box, b)==0 for b in boxes):
                boxes.append(box)
    return sorted(boxes, key=lambda b: b[0])
    
def recolor_layer(image, color_bgr):
    """
    Make every true‐black pixel → color_bgr, everything else transparent.
    Returns 4-channel BGRA.
    """
    bgr = image.copy()
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    alpha = (gray==0).astype(np.uint8)*255
    # overlay fill color where alpha=255
    for c in range(3):
        bgr[:,:,c] = np.where(alpha==255, color_bgr[c], 0)
    return cv2.cvtColor(bgr, cv2.COLOR_BGR2BGRA)

def parse_dimensions_from_pdf(pdf_path):
    """
    Scan the first page of pdf for:
      1) “Dimensions (h x w): 1.25" x 5.75"”
      2) “OVER ALL LENGTH IS 14 INCHES”
      3) any “#″ x #″” pattern
    Returns (height_in, width_in) as floats, or (0.0, 0.0).
    """
    import re
    import pdfplumber

    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""

    # 1) Explicit (h x w) line
    m = re.search(
        r'Dimensions\s*\(h\s*[x×]\s*w\)\s*:\s*([\d.]+)\s*["”]?\s*[x×]\s*([\d.]+)\s*["”]?',
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
        length = float(m2.group(1))
        # width unknown → leave as 0.0 (or derive from aspect ratio if you like)
        return length, 0.0

    # 3) Any free “#″ x #″” fallback
    m3 = re.search(r'([\d.]+)\s*["”]\s*[x×]\s*([\d.]+)\s*["”]?', text)
    if m3:
        return float(m3.group(1)), float(m3.group(2))

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
        'top-left':    (0,    0,    W//2,  H//2),
        'top-right':   (W//2, 0,    W,     H//2),
        'bottom-left': (0,    H//2, W//2,  H),
        'bottom-right':(W//2, H//2, W,     H),
    }
    x1, y1, x2, y2 = rois[quadrant]
    roi = gray[y1:y2, x1:x2]
    edges_roi = cv2.Canny(roi, 50, 150)
    res = cv2.matchTemplate(edges_roi, tpl_edges, cv2.TM_CCOEFF_NORMED)
    _, _, _, maxloc = cv2.minMaxLoc(res)
    ox, oy = offset
    return (x1 + maxloc[0] + ox, y1 + maxloc[1] + oy)

def select_best_crop_box(img_color,
                         template_sets,
                         expected_ratio=None,
                         edge=5,
                         ar_weight=1000,
                         edge_penalty_weight=1):
    """
    Try each template‐set to get a candidate box, then score by:
      • border‐ink penalty   (how much “ink” lies on the crop border)
      • edge‐penalty         (how many Canny edges on the border)
      • aspect‐ratio penalty (how far the box’s AR is from expected_ratio)
      • blob‐containment     (skip any box that cuts through the main decal blob)
    Return the (x0,y0,x1,y1) with the lowest combined score.
    """
    import cv2
    import numpy as np

    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    H, W = gray.shape

    # 1) find the main dark‐blob (the decal) and its bounding box
    _, blob_mask = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
    blob_mask = (blob_mask > 0).astype(np.uint8)
    cnts, _ = cv2.findContours(blob_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if cnts:
        c = max(cnts, key=cv2.contourArea)
        bx, by, bw, bh = cv2.boundingRect(c)
        blob_box = (bx, by, bx + bw, by + bh)
    else:
        blob_box = (0, 0, W, H)

    candidates = []

    for templates, offsets in template_sets:
        # 2) detect the four crop‐mark corners
        try:
            corners = detect_with_one_set(gray, templates, offsets)
            x0 = int((corners['top_left'][0]    + corners['bottom_left'][0]) / 2)
            y0 = int((corners['top_left'][1]    + corners['top_right'][1])   / 2)
            x1 = int((corners['top_right'][0]   + corners['bottom_right'][0]) / 2)
            y1 = int((corners['bottom_left'][1] + corners['bottom_right'][1]) / 2)
        except Exception:
            continue

        # clip to image bounds
        x0n, y0n = max(0, x0),          max(0, y0)
        x1n, y1n = min(W, x1),          min(H, y1)
        if x1n <= x0n or y1n <= y0n:
            continue

        # 3) ensure we fully contain the decal blob
        bx0, by0, bx1, by1 = blob_box
        if not (x0n <= bx0 and y0n <= by0 and x1n >= bx1 and y1n >= by1):
            continue

        # 4) border‐ink penalty
        top    = blob_mask[y0n:y0n+edge, x0n:x1n]
        bottom = blob_mask[y1n-edge:y1n, x0n:x1n]
        left   = blob_mask[y0n:y1n, x0n:x0n+edge]
        right  = blob_mask[y0n:y1n, x1n-edge:x1n]
        ink_penalty = int(top.sum() + bottom.sum() + left.sum() + right.sum())

        # 5) edge‐detection penalty on the *border*
        edges = cv2.Canny(gray[y0n:y1n, x0n:x1n], 50, 150)
        # count edge pixels along border
        border_pixels = (
            edges[0, :].sum() + edges[-1, :].sum() +
            edges[:, 0].sum() + edges[:, -1].sum()
        ) / 255
        edge_penalty = int(border_pixels * edge_penalty_weight)

        # 6) aspect‐ratio penalty
        ar = (x1n - x0n) / float(y1n - y0n)
        ar_penalty = abs(ar - expected_ratio) * ar_weight if expected_ratio else 0

        total_score = ink_penalty + edge_penalty + ar_penalty
        candidates.append((total_score, (x0n, y0n, x1n, y1n)))

    if not candidates:
        raise RuntimeError("No valid crop candidates found")

    # pick the candidate with minimal combined score


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
    Attempt to download the PDF for `part`, retrying on lockout by
    fully restarting the browser with increasing delays.
    Returns a tuple (pdf_path, driver).
    """
    delays = [10, 20, 30, 40, 50]
    for delay in delays:
        try:
            # Try the normal download path; will raise if locked out
            return download_pdf_for_part(part, tmp_dir, driver, base_url), driver
        except (WebDriverException, TimeoutException) as e:
            print(f"    · Locked out ({e}); closing browser and retrying in {delay}s…")
            # Kill the hanging session
            try:
                driver.quit()
            except:
                pass
            time.sleep(delay)
            # Start fresh
            driver = init_driver(tmp_dir, profile_dir=profile, headless=False)
            driver.get(base_url)
            print("    · Waiting for library page to become available…")
            # Block until the search field shows up again (or timeout)
            try:
                wait_for_login(driver, timeout=60)
            except TimeoutException:
                print(f"    · Still locked out after {delay}s; will retry.")
                continue

    # Final attempt (no further delays)
    return download_pdf_for_part(part, tmp_dir, driver, base_url), driver
    
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
        x,y,w,h = cv2.boundingRect(c)
        if w*h < min_area: continue
        boxes.append((x,y,w,h))

    if len(boxes) < 2:
        return None

    bottoms = [y+h for x,y,w,h in boxes]
    avg_b = sum(bottoms)/len(bottoms)
    group = [b for b in boxes if abs((b[1]+b[3]) - avg_b) <= tol]

    if len(group) < 2:
        return None

    xs = [x for x,y,w,h in group] + [x+w for x,y,w,h in group]
    ys = [y for x,y,w,h in group] + [y+h for x,y,w,h in group]
    x0 = max(min(xs)-pad, 0)
    y0 = max(min(ys)-pad, 0)
    x1 = min(max(xs)+pad, img_color.shape[1])
    y1 = min(max(ys)+pad, img_color.shape[0])
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

    wait = WebDriverWait(driver, 20)
    delays = [10, 20, 30]  # seconds to wait before each retry

    for delay in delays:
        driver.get(base_url)
        time.sleep(1)  # let the page start rendering

        # 1) immediate lock-out check by page title
        if "Sign in to your account" in driver.title:
            print(f"· Locked out at startup; closing browser and retrying in {delay}s…")
            driver.quit()
            time.sleep(delay)
            driver = init_driver(tmp_dir, profile_dir=profile)
            continue

        # 2) try waiting for the library search field
        try:
            wait.until(EC.presence_of_element_located((By.ID, 'docLibContainer_search_field')))
            print("· Library page ready")
            break
        except TimeoutException:
            print(f"· Timeout waiting for library; closing browser and retrying in {delay}s…")
            driver.quit()
            time.sleep(delay)
            driver = init_driver(tmp_dir, profile_dir=profile)
    else:
        raise RuntimeError("❌ Could not reach Document Library after multiple retries")

    # ─── Load templates once ───────────────────────────────────────────────────
    template_sets = load_template_sets('templates')
    print(f"· Loaded {len(template_sets)} templet sets for corner detection")

    # ─── Read parts list ───────────────────────────────────────────────────────
    df = pd.read_excel(input_sheet, dtype=str)
    df.columns = df.columns.str.upper()
    df.rename(columns={df.columns[0]:'PART', df.columns[1]:'TMS'}, inplace=True)
    records = []
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M') + '00'

    # ─── Loop over each row ────────────────────────────────────────────────────
    for i, row in df.iterrows():
        original_part = row['PART'].strip()
        tms           = row['TMS']
        search_part   = strip_gt_suffix(original_part)
    
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

            # ─── 2a) Render first page to BGR image & build ink mask ──────────────────
            img = render_pdf_color_page(pdf_path, dpi=DPI)
            h_img, w_img = img.shape[:2]
            # build a binary mask of “ink” pixels
            gray     = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, blob  = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
            # find first row with any ink
            ys = np.where(blob.sum(axis=1) > 0)[0]
            y0_art = int(ys.min()) if ys.size else 0
            # optional pad above artwork
            PAD_TOP = 5
            y0_art = max(0, y0_art - PAD_TOP)
            
            # ─── 2b) Now choose your crop in X and Y as before,
            # but clamp the top at y0_art:
            # e.g. for bracket-template:
            print("    · Bracket-template crop…")
            x0, y0_old, x1, y1 = select_best_crop_box(img, template_sets)
            # enforce no header above artwork
            y0 = max(y0_old, y0_art)
            crop_img = img[y0:y1, x0:x1]
            time.sleep(STEP_DELAY)
    
            # 3) Try bracket‐template crop (guided by the PDF’s stated dims)
            print("    · Bracket-template crop…")
            # parse the PDF for its H×W dims (inches)
            h_in, w_in = parse_dimensions_from_pdf(pdf_path)
            expected_ar = (w_in / h_in) if (h_in and w_in) else None
        
            try:
                # pass expected_ar into the selector (None if dims absent)
                x0, y0, x1, y1 = select_best_crop_box(
                    img,
                    template_sets,
                    expected_ratio=expected_ar
                )
                print(f"    · Bracket crop box: {(x0, y0, x1, y1)}")
                crop_img = img[y0:y1, x0:x1]
        
            except RuntimeError as e:
                # fallback if no good bracket candidates
                print(f"    · No valid bracket candidates ({e}); falling back…")
                # your existing blob or full‐logo fallback here, e.g.:
                grp = find_aligned_blob_group(img, min_area=5000, tol=10, pad=20)
                if grp:
                    x0, y0, x1, y1 = grp
                    print(f"    · Aligned blob group crop: {(x0, y0, x1, y1)}")
                    crop_img = img[y0:y1, x0:x1]
                else:
                    # your full‐logo or blob fallback
                    y_crop = crop_full_logo(pdf_path, dpi=DPI)
                    if y_crop:
                        print(f"    · Full-logo crop at y={y_crop}px")
                        crop_img = img[:y_crop, :]
                    else:
                        # last‐ditch: everything
                        h_img, w_img = img.shape[:2]
                        margin = int(0.01 * min(h_img, w_img))
                        print("    · Full-page margin crop")
                        crop_img = img[margin:h_img-margin, margin:w_img-margin]
    
            # 4) Decide which crop method to use
            pad = 20  # padding for blob-based crops
            crop_img = None
    
            # a) full-logo crop via “…mm” line
            y_crop = crop_full_logo(pdf_path, dpi=DPI)
            if y_crop and y_crop > 0:
                print(f"    · Full-logo crop at y={y_crop}px")
                crop_img = img[:y_crop, :]
    
            # b) large decal spanning multiple blobs
            else:
                grp = find_aligned_blob_group(img, min_area=5000, tol=10, pad=pad)
                if grp:
                    x0,y0,x1,y1 = grp
                    print(f"    · Aligned blob group crop: {(x0,y0,x1,y1)}")
                    crop_img = img[max(0,y0-pad):min(y1+pad,h_img),
                                   max(0,x0-pad):min(x1+pad,w_img)]
    
            # c) legacy multi-layer via multiple template sets
            if crop_img is None:
                boxes = select_all_crop_candidates(img, template_sets, penalty_thresh=0.1)
                if len(boxes) > 1:
                    print("    · Legacy multi-layer detected → compositing layers…")
                    layers = []
                    for bx0,by0,bx1,by1 in boxes:
                        region = img[by0:by1, bx0:bx1]
                        color = extract_color_label(pdf_path, crop_y0=by0)
                        layers.append(recolor_layer(region, COLOR_MAP[color]))
                    base = layers[0]
                    for layer in layers[1:]:
                        alpha = layer[:,:,3].astype(float)/255.0
                        for c in range(3):
                            base[:,:,c] = layer[:,:,c]*alpha + base[:,:,c]*(1-alpha)
                    crop_img = cv2.cvtColor(base, cv2.COLOR_BGRA2BGR)
    
            # d) single best bracket crop
            if crop_img is None:
                try:
                    print("    · Bracket-template crop…")
                    # bias the template matcher toward the PDF’s H×W ratio
                    x0,y0,x1,y1 = select_best_crop_box(img, template_sets, expected_ratio=expected_ar)
                    print(f"    · Bracket crop box: {(x0,y0,x1,y1)}")
                    crop_img = img[y0:y1, x0:x1]
                except RuntimeError:
                            # e) fallback to single blob bbox + margin
                            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                            blob = crop_blob_bbox(gray) or (0,0,w_img,h_img)
                            bx,by,bx2,by2 = blob
                            print(f"    · Blob-fallback crop: {blob}")
                            crop_img = img[max(0,by-pad):min(by2+pad,h_img),
                                           max(0,bx-pad):min(bx2+pad,w_img)]
    
            # 5) Save JPEG
            jpg_name = f"{tms}.{original_part}.{seq}.jpg"
            out_jpg  = os.path.join(imgs_dir, jpg_name)
            print(f"    · Writing JPEG → {out_jpg}")
            cv2.imwrite(out_jpg, crop_img)
    
            # 6) Clean up PDF
            print("    · Removing temp PDF")
            os.remove(pdf_path)
            time.sleep(STEP_DELAY)
    
            # 7) Record row (dimensions already parsed)
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
                'IMAGE_FILE_NAME': '',
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
    driver.quit()
    shutil.rmtree(tmp_dir)
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



    # ── 6) RENDER & CROP ─────────────────────────────────────────
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
        gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
        cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if cnts:
            bx, by, bw, bh = cv2.boundingRect(max(cnts, key=cv2.contourArea))
            print(f"    · Blob crop box: {(bx, by, bx+bw, by+bh)}")
            crop_region = img_color[by:by+bh, bx:bx+bw]
        else:
            m = int(0.01 * min(h_img, w_img))
            print(f"    · Full-page margin crop: {(m, m, w_img-m, h_img-m)}")
            crop_region = img_color[m:h_img-m, m:w_img-m]

    # 6c) Extract that region
    region = img_color[y0:y1, x0:x1]
    rh, rw = crop_region.shape[:2]

    # 6d) Legacy multi-layer? 3 bands side-by-side
    if rw > rh * 1.8:
        print("    · Detected legacy multi-layer → slicing bands…")
        third = rw // 3
        green = crop_region[:, third:2*third]
        black = crop_region[:, 2*third:3*third]
        def recolor(band, bgr):
            g = cv2.cvtColor(band, cv2.COLOR_BGR2GRAY)
            _, mask = cv2.threshold(g, 250, 255, cv2.THRESH_BINARY_INV)
            fill = np.zeros_like(band); fill[:] = bgr
            return np.where(mask[:,:,None]>0, fill, band)
        band_g = recolor(green, COLOR_MAP['green'])
        band_b = recolor(black, COLOR_MAP['black'])
        stacked = band_b.copy()
        mask_g = cv2.cvtColor(band_g, cv2.COLOR_BGR2GRAY) < 250
        for c in range(3):
            stacked[:,:,c] = np.where(mask_g, band_g[:,:,c], stacked[:,:,c])
        crop = stacked
    else:
        crop = crop_region

    # ── 7) Save JPEG ─────────────────────────────────────────
    jpg_name = f"{tms}.{part}.{seq}.jpg"
    out_jpg  = os.path.join(imgs_dir, jpg_name)
    print(f"    · Writing JPEG → {out_jpg}")
    cv2.imwrite(out_jpg, crop)
    time.sleep(STEP_DELAY)

    # ── 8) Compute dims, volume, weight ────────────────────
    h_in, w_in = parse_dimensions_from_pdf(pdf_path)
    vol     = h_in * w_in * THICKNESS_IN
    wgt     = vol * MATERIAL_DENSITY
    dim_wgt = vol / FACTOR

    # ── 9) Clean up ─────────────────────────────────────────
    os.remove(pdf_path)

    # ── 10) Record ──────────────────────────────────────────
    records.append({
        'ITEM_ID':         part,
        'ITEM_TYPE':       '',
        'DESCRIPTION':     '',
        'NET_LENGTH':      h_in,
        'NET_WIDTH':       w_in,
        'NET_HEIGHT':      THICKNESS_IN,
        'NET_WEIGHT':      wgt,
        'NET_VOLUME':      vol,
        'NET_DIM_WGT':     dim_wgt,
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
