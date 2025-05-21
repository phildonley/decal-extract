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
from selenium.common.exeptions import WebDriverException

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
def strip_suffix(part: str) -> str:
    """
    Remove the last two characters only if they are both letters.
    E.g. "1293217GT" -> "1293217", but "65417" -> "65417"
    """
    if len(part) > 2 and part[-2:].isalpha():
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
        opts.add_argument(f"--user-data-dir={profile_dir}")
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }
    opts.add_experimental_option("prefs", prefs)
    if headless:
        opts.add_argument("--headless")
    return webdriver.Chrome(options=opts)

def download_pdf_for_part(part, tmp_dir, driver, base_url):
    """
    Search for the exact part, open its Documents modal,
    then either fetch via requests (if onclick URL exists),
    or click-and-poll tmp_dir for the PDF.
    Returns the local pdf_path (string).
    """
    wait = WebDriverWait(driver, 20)

    # 1) Go back to main list & search
    driver.get(base_url)
    wait.until(EC.presence_of_element_located((By.ID, 'docLibContainer_search_field')))

    clean = strip_suffix(part)

    fld = wait.until(EC.element_to_be_clickable((By.ID, 'docLibContainer_search_field')))
    fld.clear()
    fld.send_keys(clean)
    driver.find_element(By.ID, 'docLibContainer_search_button').click()

    # 2) Find the <td> whose text == our part, then get its <tr>
    td = wait.until(EC.presence_of_element_located((
        By.XPATH,
        f"//td[normalize-space(text())='{clean}']"
    )))
    tr = td.find_element(By.XPATH, "./ancestor::tr")

    # 3) Click that row’s Documents button
    docs_btn = tr.find_element(By.XPATH, ".//button[contains(., 'Documents')]")
    docs_btn.click()

    # 4) Into the Documents iframe
    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.CSS_SELECTOR, "iframe[title='Documents']")
    ))

    # 5) Try to parse an inline URL & filename
    dl = wait.until(EC.presence_of_element_located((By.ID, "downloadBtn")))
    onclick = dl.get_attribute("onclick") or ""
    m = re.search(r"doDownload\('([^']+)','([^']+)'\)", onclick)

    if m:
        # fetch via requests (so we can inject cookies)
        url_frag, filename = m.groups()
        full_url = urljoin(base_url, url_frag)
        cookies = {c["name"]: c["value"] for c in driver.get_cookies()}
        resp = requests.get(full_url, cookies=cookies, timeout=60)
        resp.raise_for_status()

        out_path = os.path.join(tmp_dir, filename)
        with open(out_path, "wb") as f:
            f.write(resp.content)

        driver.switch_to.default_content()
        return out_path, driver

    # 6) Fallback: click + poll for the file to appear
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

def parse_dimensions_from_pdf(pdf_path):
    pat = r"Dimensions\s*\(h\s*x\s*w\)\s*:\s*([\d.]+)\s*[xX]\s*([\d.]+)"
    with pdfplumber.open(pdf_path) as pdf:
        txt = pdf.pages[0].extract_text() or ""
    m = re.search(pat, txt)
    if m:
        return float(m.group(1)), float(m.group(2))
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

def select_best_crop_box(img_color, template_sets, edge=5):
    """
    Try each template‐set to get a candidate box, then score by
    how many non‐white pixels lie on the crop boundary.
    Return the (x0,y0,x1,y1) with the lowest edge‐penalty.
    """
    gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
    # everything below 250 is “ink”
    _, blob = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
    blob = (blob > 0).astype(np.uint8)

    candidates = []
    for templates, offsets in template_sets:
        try:
            # run your 4‐corner match for this set
            corners = detect_with_one_set(gray, templates, offsets)
            x0, y0, x1, y1 = average_rectangle({
                'top-left':    corners['top_left'],
                'top-right':   corners['top_right'],
                'bottom-left': corners['bottom_left'],
                'bottom-right':corners['bottom_right'],
            })
        except Exception:
            continue

        # clip into image
        x0, y0 = max(0, x0), max(0, y0)
        x1, y1 = min(blob.shape[1], x1), min(blob.shape[0], y1)
        if x1 <= x0 or y1 <= y0:
            continue

        # build 4 edge masks
        top    = blob[y0:y0+edge, x0:x1]
        bottom = blob[y1-edge:y1, x0:x1]
        left   = blob[y0:y1, x0:x0+edge]
        right  = blob[y0:y1, x1-edge:x1]
        penalty = int(top.sum() + bottom.sum() + left.sum() + right.sum())

        candidates.append((penalty, (x0, y0, x1, y1)))

    if not candidates:
        raise RuntimeError("No valid crop candidates found")

    best_penalty, best_box = min(candidates, key=lambda t: t[0])
    print(f"    · Chosen crop box {best_box} with penalty {best_penalty}")
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
    Wrap download_pdf_for_part in a try/except.
    If Selenium hiccups, restart the browser once and retry.
    Returns a tuple (pdf_path, driver), where driver may be a new session.
    """
    try:
        pdf_path = download_pdf_for_part(part, tmp_dir, driver, base_url)
        return pdf_path, driver
    except WebDriverException as e:
        print(f"    · WebDriver hiccup ({e}); restarting browser…")
        # tear down the old session
        try:
            driver.quit()
        except:
            pass
        time.sleep(2)

        # start a fresh browser session
        new_driver = init_driver(tmp_dir, profile_dir=profile, headless=False)
        new_driver.get(base_url)
        wait_for_login(new_driver)

        # retry the download
        pdf_path = download_pdf_for_part(part, tmp_dir, new_driver, base_url)
        return pdf_path, new_driver

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
    print(f"· Loaded {len(template_sets)} templet sets for corner detection")

    # ─── Read parts list ───────────────────────────────────────────────────────
    df = pd.read_excel(input_sheet, dtype=str)
    df.columns = df.columns.str.upper()
    df.rename(columns={df.columns[0]:'PART', df.columns[1]:'TMS'}, inplace=True)
    records = []
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M') + '00'

    # ─── Loop over each row ────────────────────────────────────────────────────
    for i, row in df.iterrows():
        raw_part, tms = row['PART'], row['TMS']
        part = strip_suffix(raw_part)
        print(f"[{i}] ➡️ Processing part={part}, TMS={tms}")

        try:
            # 1) Download PDF (auto-retries on WebDriver errors)
            pdf_path, driver = safe_download(part, tmp_dir, driver, base_url, profile)
            print(f"   · PDF downloaded → {pdf_path}")
            time.sleep(STEP_DELAY)

            # 2) Render to BGR image
            print("    · Rendering to image…")
            img_color = render_pdf_color_page(pdf_path, dpi=DPI)
            h_img, w_img = img_color.shape[:2]
            time.sleep(STEP_DELAY)

             # 3) Select the best crop box across all template‐sets
            try:
                print("    · Selecting best crop box…")
                x0, y0, x1, y1 = select_best_crop_box(img_color, template_sets)
            except Exception as e:
                print(f"    · Crop‐by‐templates failed ({e}); falling back to blob/full‐page…")
                gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
                _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
                cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                if cnts:
                    c = max(cnts, key=cv2.contourArea)
                    bx, by, bw, bh = cv2.boundingRect(c)
                    x0, y0, x1, y1 = bx, by, bx + bw, by + bh
                    print(f"    · Blob crop box: {(x0, y0, x1, y1)}")
                else:
                    margin = int(0.01 * min(h_img, w_img))
                    x0, y0, x1, y1 = margin, margin, w_img - margin, h_img - margin
                    print(f"    · Full‐page margin box: {(x0, y0, x1, y1)}")

            # 4) Extract region & handle legacy bands
            region = img_color[y0:y1, x0:x1]
            rh, rw = region.shape[:2]
            if rw > rh*1.8:
                print("    · Detected legacy multi‐layer → slicing bands…")
                third = rw//3
                green = region[:, third:2*third]
                black = region[:, 2*third:3*third]
                def recolor_band(band, bgr):
                    g = cv2.cvtColor(band, cv2.COLOR_BGR2GRAY)
                    _, mask = cv2.threshold(g, 250, 255, cv2.THRESH_BINARY_INV)
                    fill = np.zeros_like(band)
                    fill[:] = bgr
                    return np.where(mask[:,:,None]>0, fill, band)
                band_g = recolor_band(green, COLOR_MAP['green'])
                band_b = recolor_band(black, COLOR_MAP['black'])
                stacked = band_b.copy()
                mask_g = cv2.cvtColor(band_g, cv2.COLOR_BGR2GRAY)<250
                for c in range(3):
                    stacked[:,:,c] = np.where(mask_g, band_g[:,:,c], stacked[:,:,c])
                crop = stacked
            else:
                crop = region

            # 5) Save JPEG
            jpg_name = f"{tms}.{part}.{seq}.jpg"
            out_jpg  = os.path.join(imgs_dir, jpg_name)
            print(f"    · Writing JPEG → {out_jpg}")
            cv2.imwrite(out_jpg, crop)
            time.sleep(STEP_DELAY)

            # 6) Parse dim/compute
            print("    · Parsing dimensions & computing…")
            h_in, w_in = parse_dimensions_from_pdf(pdf_path)
            vol  = h_in * w_in * THICKNESS_IN
            wgt  = vol * MATERIAL_DENSITY
            dimw = vol / FACTOR
            time.sleep(STEP_DELAY)

            # 7) Clean up PDF
            print("    · Removing temp PDF")
            os.remove(pdf_path)
            time.sleep(STEP_DELAY)

            # 8) Record row
            records.append({
                'ITEM_ID':      part,
                'ITEM_TYPE':    '',
                'DESCRIPTION':  '',
                'NET_LENGTH':   h_in,
                'NET_WIDTH':    w_in,
                'NET_HEIGHT':   THICKNESS_IN,
                'NET_WEIGHT':   wgt,
                'NET_VOLUME':   vol,
                'NET_DIM_WGT':  dimw,
                'DIM_UNIT':     'in',
                'WGT_UNIT':     'lb',
                'VOL_UNIT':     'in',
                'FACTOR':       FACTOR,
                'SITE_ID':      SITE_ID,
                'TIME_STAMP':   ts,
                'OPT_INFO_2':   'Y',
                'OPT_INFO_3':   'N',
                'OPT_INFO_8':   0,
                'IMAGE_FILE_NAME': jpg_name,
                'UPDATED':        'Y'
            })
            print(f"[{i}] ✅ Done\n")
            time.sleep(STEP_DELAY)

        except Exception as e:
            print(f"[{i}] ❌ ERROR: {e}")
            with open(os.path.join(dbg_dir,'errors.log'),'a',encoding='utf-8') as f:
                f.write(f"{raw_part}: {e}\n")
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
        print("    · Detecting corner brackets…")
        corners = detect_bracket_corners(img_color, templates, offsets)
        x0, y0, x1, y1 = average_rectangle(corners)
        print(f"    · Bracket crop box: {(x0, y0, x1, y1)}")
    except Exception as e:
        print(f"    · No brackets ({e}); falling back to blob/full-page…")
        gray = cv2.cvtColor(img_color, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
        cnts, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if cnts:
            c = max(cnts, key=cv2.contourArea)
            bx, by, bw, bh = cv2.boundingRect(c)
            x0, y0, x1, y1 = bx, by, bx + bw, by + bh
            print(f"    · Blob crop box: {(x0, y0, x1, y1)}")
        else:
            margin = int(0.01 * min(h_img, w_img))
            x0, y0, x1, y1 = margin, margin, w_img - margin, h_img - margin
            print(f"    · Full-page margin box: {(x0, y0, x1, y1)}")

    # 6c) Extract that region
    region = img_color[y0:y1, x0:x1]
    rh, rw = region.shape[:2]

    # 6d) Legacy multi-layer? 3 bands side-by-side
    if rw > rh * 1.8:
        print("    · Detected legacy multi-layer → slicing into thirds…")
        third = rw // 3
        green_band = region[:, third:2*third]
        black_band = region[:, 2*third:3*third]

        def recolor_band(band, color_bgr):
            g = cv2.cvtColor(band, cv2.COLOR_BGR2GRAY)
            _, mask = cv2.threshold(g, 250, 255, cv2.THRESH_BINARY_INV)
            fill = np.zeros_like(band)
            fill[:] = color_bgr
            return np.where(mask[:, :, None] > 0, fill, band)

        # recolor each band
        band_g = recolor_band(green_band, COLOR_MAP['green'])
        band_b = recolor_band(black_band, COLOR_MAP['black'])

        # overlay green atop black
        stacked = band_b.copy()
        mask_g = cv2.cvtColor(band_g, cv2.COLOR_BGR2GRAY) < 250
        for ch in range(3):
            stacked[:, :, ch] = np.where(mask_g, band_g[:, :, ch], stacked[:, :, ch])

        crop = stacked

    else:
        # single-image crop
        crop = region

    # ── 7) Save JPEG ─────────────────────────────────────────
    jpg_name = f"{tms}.{part}.{seq}.jpg"
    out_jpg  = os.path.join(imgs_dir, jpg_name)
    print(f"    · Saving JPEG → {out_jpg}")
    cv2.imwrite(out_jpg, crop)

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
        'IMAGE_FILE_NAME': jpg_name,
        'UPDATED':         'Y'
    })
    print(f"[{i}] ✅ Done\n")

    # ── 6) Tear down & write CSV ────────────────────────────────────────────────
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


# ── Entry Point ────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    tk.Tk().withdraw()
    sheet = filedialog.askopenfilename(
        title='Select parts sheet',
        filetypes=[('Excel/CSV','*.xlsx *.xls *.csv')]
    )
    if not sheet:
        exit()

    out_root = filedialog.askdirectory(title='Select output folder')
    if not out_root:
        exit()

    url = simpledialog.askstring('Document Library URL', 'Enter the library URL:')
    if not url:
        exit()

    profile = simpledialog.askstring(
        'Chrome Profile (optional)',
        'Enter Chrome profile path, or leave blank:'
    )

    main(
        input_sheet=sheet,
        output_root=out_root,
        base_url=url,
        profile=profile or None,
        seq=105
    )
