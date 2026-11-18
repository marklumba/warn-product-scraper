
"""
WARN.com Product IMAGE Scraper
Built on the same framework as warn_scraper_full.py.

For every MPN in Warn_Scrape_List_*.xlsx:
  1. Search the MPN on warn.com and open the product page.
  2. Collect every image in the main product gallery, in gallery order.
  3. Download each image at 1800 px wide and convert to JPG.
  4. Save all images in ONE main folder using a double underscore:
       WAR-17801__1.jpg
       WAR-17801__2.jpg
       WAR-85360__1.jpg
  5. Generate an Excel log:
       - Image Log
       - Parts To Review

Important:
  If the requested MPN is not found on the opened product page,
  the product is skipped and Product Title remains blank in both
  Excel sheets. The incorrect title is retained in the application
  log for debugging purposes.

Requirements:
  pip install pillow requests pandas openpyxl seleniumbase
"""

import io
import os
import re
import json
import time
import base64
import shutil
import logging
import tempfile

from datetime import datetime
from pathlib import Path
from urllib.parse import (
    quote,
    urljoin,
    urlparse,
    urlunparse,
    parse_qsl,
    urlencode,
)

import pandas as pd
import requests

from openpyxl.utils import get_column_letter
from PIL import Image

from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


# ---------------- Constants ----------------

WEBSITE = "https://www.warn.com/"
SEARCH_URL = "https://www.warn.com/search?q={mpn}"

INPUT_FILE = r"Warn_Scrape_List_2026-11-16.xlsx"

ELEMENT_WAIT_TIME = 20
PAGE_PAUSE = 3
SEARCH_RESULT_WAIT = 12
SAVE_LOG_EVERY = 5


# ---------------- Image Settings ----------------

TARGET_WIDTH = 1800

UPSCALE_SMALLER_IMAGES = True

OUTPUT_FORMAT = "jpg"

JPEG_QUALITY = 95

SKIP_EXISTING = True

NAME_SEPARATOR = "__"

REVIEW_MAX_IMAGES = 1


# ---------------- Output Paths ----------------

OUTPUT_ROOT = (
    Path(os.path.expanduser("~"))
    / "Desktop"
    / f"Warn_Images_{datetime.now():%Y-%m-%d}"
)

LOG_FILE = (
    Path(os.path.expanduser("~"))
    / "Desktop"
    / f"Warn_Image_Log_{datetime.now():%Y-%m-%d}.xlsx"
)


# ---------------- Gallery Selectors ----------------

GALLERY_SELECTORS = [
    'media-gallery ul[id^="Slider-Gallery"] li.product__media-item img',
    'media-gallery ul[id^="Slider-Gallery"] img',
    'ul[id^="Slider-Gallery"] img',
    "div.product__media img",
]


# ---------------- Logging ----------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s: %(message)s",
    handlers=[
        logging.FileHandler("warn_image_scraper.log"),
        logging.StreamHandler(),
    ],
)


# ---------------- Driver ----------------

def setup_driver():
    user_data_dir = tempfile.mkdtemp()

    driver = Driver(uc=True)
    driver.user_data_dir = user_data_dir
    driver.set_page_load_timeout(30)

    return driver


def wait_for_captcha(driver):
    print("If a CAPTCHA / bot check appears, solve it manually.")
    input("Press Enter to continue...")


def cleanup(driver):
    try:
        driver.quit()
    except Exception as e:
        print(f"Driver quit error (can be ignored): {e}")

    if hasattr(driver, "user_data_dir"):
        if os.path.exists(driver.user_data_dir):
            shutil.rmtree(
                driver.user_data_dir,
                ignore_errors=True,
            )


# ---------------- Input / Helpers ----------------

def load_mpn_list(path):
    df = pd.read_excel(path, dtype=str)

    inv_col = next(
        (c for c in df.columns if "inventory" in c.lower()),
        df.columns[0],
    )

    df = df[[inv_col]].copy()
    df.columns = ["Inventory Number"]

    df = df.dropna(subset=["Inventory Number"])

    df["Inventory Number"] = df["Inventory Number"].str.strip()

    # Extract numeric portion after the last hyphen.
    # Example: WAR-102255 -> 102255

    df["MPN"] = (
        df["Inventory Number"]
        .str.extract(r"-(\d+)$")[0]
        .fillna(df["Inventory Number"])
    )

    return (
        df.drop_duplicates(subset="MPN")
        .reset_index(drop=True)
    )


def clean(text):
    return re.sub(
        r"\s+",
        " ",
        (text or "").replace("\xa0", " "),
    ).strip()


def safe_name(text):
    """Make text safe for Windows file names."""

    t = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "",
        clean(text),
    )

    return t.strip(" .") or "Untitled"


# ---------------- Product URL Index ----------------

_SITEMAP_JS = r"""
const done = arguments[arguments.length - 1];

const getText = u =>
  fetch(u, {credentials: 'include'}).then(r => {
    if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + u);
    return r.text();
  });

const locs = t =>
  Array.from(t.matchAll(/<loc>([^<]+)<\/loc>/g))
    .map(m => m[1].replace(/&amp;/g, '&'));

getText('/sitemap.xml').then(async idx => {

  const maps = locs(idx).filter(
    u => u.indexOf('sitemap_products') !== -1
  );

  let urls = [];

  for (const m of maps) {
    const t = await getText(m);

    urls = urls.concat(
      locs(t).filter(u => u.indexOf('/products/') !== -1)
    );
  }

  done(JSON.stringify(urls));

}).catch(e => done('ERR:' + e));
"""


def load_product_url_index(driver):
    """
    List every product URL from the website sitemap.
    Returns an empty list if the sitemap cannot be read.
    """

    try:
        driver.set_script_timeout(120)

        res = driver.execute_async_script(_SITEMAP_JS)

        if not isinstance(res, str) or res.startswith("ERR:"):
            logging.warning(
                f"Product sitemap could not be read ({res}); "
                "the search page will be used instead"
            )
            return []

        return [
            u.split("?")[0]
            for u in json.loads(res)
        ]

    except Exception as e:
        logging.warning(
            f"Product sitemap could not be read ({e}); "
            "the search page will be used instead"
        )

        return []


def url_from_index(url_index, mpn):
    """
    Product URLs whose handle contains the MPN as its own word.
    Example: m12-winch-17801
    """

    pat = re.compile(
        r"(?:^|-)" + re.escape(mpn.lower()) + r"(?:-|$)"
    )

    hits = [
        u for u in url_index
        if pat.search(
            u.split("/products/")[-1].lower()
        )
    ]

    hits.sort(
        key=lambda u: (
            not u.lower().rstrip("/").endswith(
                "-" + mpn.lower()
            ),
            len(u),
        )
    )

    return hits


# ---------------- Predictive Search ----------------

_SUGGEST_JS = """
const q = arguments[0], done = arguments[arguments.length - 1];

fetch(
  '/search/suggest.json?q=' + encodeURIComponent(q) +
  '&resources[type]=product&resources[limit]=10',
  {credentials: 'include'}
)
.then(r => r.json())
.then(j => done(JSON.stringify(j)))
.catch(e => done('ERR:' + e));
"""


def _product_links(driver, mpn):
    """
    Return:
      - all product links on the page
      - product links whose URL contains the MPN
    """

    hrefs = []

    for a in driver.find_elements(
        By.CSS_SELECTOR,
        'a[href*="/products/"]',
    ):

        h = (
            a.get_attribute("href") or ""
        ).split("?")[0]

        if "/products/" in h and h not in hrefs:
            hrefs.append(h)

    exact = [
        h for h in hrefs
        if mpn in h.split("/products/")[-1]
    ]

    return hrefs, exact


def _suggest_urls(driver, mpn):
    """Product URLs returned by predictive search."""

    try:
        driver.set_script_timeout(30)

        res = driver.execute_async_script(
            _SUGGEST_JS,
            mpn,
        )

        if not isinstance(res, str) or res.startswith("ERR:"):
            return []

        products = (
            json.loads(res)
            .get("resources", {})
            .get("results", {})
            .get("products", [])
        )

        urls = [
            urljoin(WEBSITE, p.get("url", "")).split("?")[0]
            for p in products
            if p.get("url")
        ]

        return [
            u for u in urls
            if mpn in u.split("/products/")[-1]
        ]

    except Exception as e:
        logging.warning(
            f"Predictive search failed for {mpn}: {e}"
        )

        return []


# ---------------- Find Product URL ----------------

def find_product_url(driver, mpn, url_index=None):
    """
    Return (product URL, status).

    Search order:
      1. Product sitemap
      2. Search page exact URL match
      3. Predictive search exact URL match
      4. First search result as a last-resort candidate

    IMPORTANT:
    The last-resort candidate is NOT considered a valid product.
    The caller must independently validate the MPN on the page.
    """

    if url_index:

        hits = url_from_index(
            url_index,
            mpn,
        )

        if hits:

            if len(hits) > 1:
                logging.warning(
                    f"MPN {mpn}: {len(hits)} products match "
                    f"in the sitemap, using {hits[0]}"
                )

            return hits[0], "OK (sitemap)"

    driver.get(
        SEARCH_URL.format(mpn=quote(mpn))
    )

    time.sleep(PAGE_PAUSE)

    if "/products/" in driver.current_url:
        return (
            driver.current_url.split("?")[0],
            "OK (direct)",
        )

    try:
        WebDriverWait(
            driver,
            SEARCH_RESULT_WAIT,
        ).until(
            lambda d: _product_links(d, mpn)[1]
        )

    except TimeoutException:
        pass

    hrefs, exact = _product_links(
        driver,
        mpn,
    )

    if exact:
        return exact[0], "OK"

    suggested = _suggest_urls(
        driver,
        mpn,
    )

    if suggested:
        return suggested[0], "OK (predictive search)"

    logging.warning(
        f"MPN {mpn}: not found in sitemap, "
        "search results or predictive search"
    )

    if hrefs:
        return (
            hrefs[0],
            "CHECK - MPN not in URL (first result used)",
        )

    return "", "No search results"


# ---------------- Product Title ----------------

def scrape_title(driver):
    for sel in (
        "div.product_title h1",
        "div.product_info-container h1",
        "h1",
    ):

        try:
            return clean(
                driver.find_element(
                    By.CSS_SELECTOR,
                    sel,
                ).get_attribute("textContent")
            )

        except NoSuchElementException:
            continue

    return ""


# ---------------- Image URLs ----------------

def _best_img_url(img):
    """
    Get the best image URL from src, data-src,
    data-original, or srcset.
    """

    src = (
        img.get_attribute("src") or ""
    ).strip()

    if src and not src.startswith("data:"):
        return src

    for attr in (
        "data-src",
        "data-original",
    ):

        v = (
            img.get_attribute(attr) or ""
        ).strip()

        if v and not v.startswith("data:"):
            return v

    srcset = (
        img.get_attribute("srcset")
        or img.get_attribute("data-srcset")
        or ""
    ).strip()

    best = ""
    best_w = -1

    for part in srcset.split(","):

        bits = part.strip().split()

        if not bits:
            continue

        w = (
            int(bits[1][:-1])
            if len(bits) > 1
            and bits[1][:-1].isdigit()
            else 0
        )

        if w > best_w:
            best = bits[0]
            best_w = w

    return best


def sized_url(url, width=TARGET_WIDTH):
    """Request a specific image width from Shopify CDN."""

    u = urlparse(
        urljoin(WEBSITE, url)
    )

    q = dict(
        parse_qsl(u.query)
    )

    q["width"] = str(width)
    q.pop("height", None)

    return urlunparse(
        u._replace(
            query=urlencode(q)
        )
    )


def _image_key(url):
    """Image identity ignoring width/height parameters."""

    u = urlparse(
        urljoin(WEBSITE, url)
    )

    q = dict(
        parse_qsl(u.query)
    )

    return (
        u.path
        + "?v="
        + q.get("v", "")
    )


def get_gallery_image_urls(driver):
    """
    Return:
      - list of image URLs in gallery order
      - expected image count from gallery, if available
    """

    expected = None

    try:
        gal = driver.find_element(
            By.CSS_SELECTOR,
            "media-gallery",
        )

        v = (
            gal.get_attribute("data-num-images")
            or ""
        ).strip()

        expected = (
            int(v)
            if v.isdigit()
            else None
        )

    except NoSuchElementException:
        pass

    urls = []
    seen = set()

    for sel in GALLERY_SELECTORS:

        for img in driver.find_elements(
            By.CSS_SELECTOR,
            sel,
        ):

            raw = _best_img_url(img)

            if not raw:
                continue

            key = _image_key(raw)

            if key in seen:
                continue

            seen.add(key)

            urls.append(
                sized_url(raw)
            )

        if urls:
            break

    return urls, expected


# ---------------- Download + Resize ----------------

def fetch_bytes(driver, url):
    """
    Download through the browser first.
    Falls back to requests using browser cookies.
    """

    js = """
    const url = arguments[0], done = arguments[arguments.length - 1];

    fetch(url, {credentials: 'include'})
      .then(r => {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.arrayBuffer();
      })
      .then(buf => {
        const a = new Uint8Array(buf);
        let s = '';
        const CH = 0x8000;

        for (let i = 0; i < a.length; i += CH) {
          s += String.fromCharCode.apply(
            null,
            a.subarray(i, i + CH)
          );
        }

        done(btoa(s));
      })
      .catch(e => done('ERR:' + e));
    """

    try:
        driver.set_script_timeout(60)

        res = driver.execute_async_script(
            js,
            url,
        )

        if isinstance(res, str) and not res.startswith("ERR:"):
            return base64.b64decode(res)

        logging.warning(
            f"Browser fetch failed ({res}); "
            f"trying requests: {url}"
        )

    except Exception as e:
        logging.warning(
            f"Browser fetch error ({e}); "
            f"trying requests: {url}"
        )

    ua = driver.execute_script(
        "return navigator.userAgent"
    )

    cookies = {
        c["name"]: c["value"]
        for c in driver.get_cookies()
    }

    r = requests.get(
        url,
        headers={
            "User-Agent": ua,
            "Referer": WEBSITE,
        },
        cookies=cookies,
        timeout=60,
    )

    r.raise_for_status()

    return r.content


def process_image(raw):
    """
    Convert image to RGB and resize to TARGET_WIDTH.
    Maintain original aspect ratio.
    """

    img = Image.open(
        io.BytesIO(raw)
    )

    img.load()

    src_w, src_h = img.size

    if (
        img.mode in ("RGBA", "LA")
        or (
            img.mode == "P"
            and "transparency" in img.info
        )
    ):

        rgba = img.convert("RGBA")

        bg = Image.new(
            "RGB",
            rgba.size,
            (255, 255, 255),
        )

        bg.paste(
            rgba,
            mask=rgba.split()[-1],
        )

        img = bg

    elif img.mode != "RGB":
        img = img.convert("RGB")

    note = ""

    if src_w != TARGET_WIDTH:

        if src_w > TARGET_WIDTH or UPSCALE_SMALLER_IMAGES:

            new_h = max(
                1,
                round(
                    src_h * TARGET_WIDTH / src_w
                ),
            )

            img = img.resize(
                (TARGET_WIDTH, new_h),
                Image.LANCZOS,
            )

            note = (
                f"resized {src_w}px -> {TARGET_WIDTH}px"
                + (
                    " (enlarged)"
                    if src_w < TARGET_WIDTH
                    else ""
                )
            )

        else:

            note = (
                f"smaller than {TARGET_WIDTH}px "
                f"({src_w}px), not enlarged"
            )

    return img, note


def save_image(img, path):

    if OUTPUT_FORMAT.lower() == "png":

        img.save(
            path,
            "PNG",
            optimize=True,
        )

    else:

        img.save(
            path,
            "JPEG",
            quality=JPEG_QUALITY,
            optimize=True,
        )


def image_path(inv, index):
    """
    ONE main folder, no subfolders.

    Example:
      WAR-127348__1.jpg
      WAR-127348__2.jpg
    """

    ext = (
        "png"
        if OUTPUT_FORMAT.lower() == "png"
        else "jpg"
    )

    return (
        OUTPUT_ROOT
        / f"{safe_name(inv)}{NAME_SEPARATOR}{index}.{ext}"
    )


# ---------------- Excel Log ----------------

def _format_log_sheet(ws, max_width=60):
    """
    Autofit column widths.
    Add filters and freeze the header row and first column.
    """

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_cells in enumerate(
        ws.columns,
        start=1,
    ):

        longest = max(
            (
                len(str(c.value))
                for c in column_cells
                if c.value is not None
            ),
            default=0,
        )

        ws.column_dimensions[
            get_column_letter(col_idx)
        ].width = min(
            max(longest + 2, 8),
            max_width,
        )


def save_log(log_rows, review_rows=None):
    """
    Sheet 1: Image Log
    Sheet 2: Parts To Review
    """

    if not log_rows and not review_rows:
        return

    log_cols = [
        "Inventory Number",
        "MPN",
        "Product Title",
        "Image #",
        "File Name",
        "Width",
        "Height",
        "Status",
        "Note",
        "Source URL",
    ]

    review_cols = [
        "Inventory Number",
        "MPN",
        "Product Title",
        "Image Count",
        "Review Status",
        "Product URL",
    ]

    with pd.ExcelWriter(
        LOG_FILE,
        engine="openpyxl",
    ) as writer:

        pd.DataFrame(
            log_rows or [],
            columns=log_cols,
        ).to_excel(
            writer,
            index=False,
            sheet_name="Image Log",
        )

        pd.DataFrame(
            review_rows or [],
            columns=review_cols,
        ).to_excel(
            writer,
            index=False,
            sheet_name="Parts To Review",
        )

        for ws in writer.book.worksheets:
            _format_log_sheet(ws)

    print(
        f"Log saved: {LOG_FILE} "
        f"({len(review_rows or [])} parts to review)"
    )


# ---------------- Main ----------------

def main(test_mpns=None):
    """
    test_mpns:
        Optional list of MPN strings to run instead of the full list.

    Example:
        main(test_mpns=["28500", "17801", "26502"])
    """

    driver = None

    log_rows = []
    review_rows = []

    try:

        mpns = load_mpn_list(
            INPUT_FILE
        )

        if test_mpns:

            mpns = (
                mpns[
                    mpns["MPN"].isin(test_mpns)
                ]
                .reset_index(drop=True)
            )

        print(
            f"Loaded {len(mpns)} MPNs"
        )

        OUTPUT_ROOT.mkdir(
            parents=True,
            exist_ok=True,
        )

        driver = setup_driver()

        driver.get(WEBSITE)

        wait_for_captcha(driver)

        url_index = load_product_url_index(
            driver
        )

        print(
            f"Product index: {len(url_index)} "
            "product URLs read from the sitemap"
        )

        for i, r in mpns.iterrows():

            mpn = r["MPN"]
            inv = r["Inventory Number"]

            print(
                f"[{i + 1}/{len(mpns)}] "
                f"MPN {mpn} ({inv})"
            )

            title = ""
            url = ""
            count = 0

            def flag(review_status):
                review_rows.append({
                    "Inventory Number": inv,
                    "MPN": mpn,
                    "Product Title": title,
                    "Image Count": count,
                    "Review Status": review_status,
                    "Product URL": url,
                })

            try:

                url, status = find_product_url(
                    driver,
                    mpn,
                    url_index,
                )

                if not url:

                    logging.warning(
                        f"MPN {mpn}: {status}"
                    )

                    log_rows.append({
                        "Inventory Number": inv,
                        "MPN": mpn,
                        "Status": status,
                    })

                    flag(
                        "Product not found - check"
                    )

                else:

                    if not status.startswith("OK"):

                        logging.warning(
                            f"MPN {mpn} needs a manual check: "
                            f"{status} {url}"
                        )

                    driver.get(url)

                    time.sleep(PAGE_PAUSE)

                    WebDriverWait(
                        driver,
                        ELEMENT_WAIT_TIME,
                    ).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "h1")
                        )
                    )

                    title = scrape_title(driver)

                    # -----------------------------------------
                    # MPN VALIDATION
                    # -----------------------------------------

                    if mpn not in driver.page_source:

                        # Preserve the incorrect title for
                        # debugging in the application log.
                        rejected_title = title

                        logging.error(
                            f"MPN {mpn} ({inv}): opened page "
                            f"'{rejected_title}' does not contain "
                            f"the MPN. Skipped (wrong product?): {url}"
                        )

                        # IMPORTANT:
                        # Clear the title so the wrong product
                        # is not associated with the requested MPN.
                        title = ""

                        log_rows.append({
                            "Inventory Number": inv,
                            "MPN": mpn,
                            "Product Title": "",
                            "Status": (
                                "MPN not found on page - "
                                "skipped (wrong product?)"
                            ),
                        })

                        # Product Title will also be blank
                        # in the Parts To Review sheet.
                        flag(
                            "Wrong product opened - check"
                        )

                    else:

                        # -------------------------------------
                        # VALID PRODUCT: COLLECT GALLERY IMAGES
                        # -------------------------------------

                        urls, expected = get_gallery_image_urls(
                            driver
                        )

                        count = len(urls)

                        if count <= REVIEW_MAX_IMAGES:

                            flag(
                                "No images"
                                if count == 0
                                else f"{count} image - possible placeholder"
                            )

                        if not urls:

                            logging.warning(
                                f"MPN {mpn} ({inv}): "
                                "no gallery images found"
                            )

                            log_rows.append({
                                "Inventory Number": inv,
                                "MPN": mpn,
                                "Product Title": title,
                                "Status": "No images found",
                            })

                        else:

                            if (
                                expected is not None
                                and expected != count
                            ):

                                logging.warning(
                                    f"MPN {mpn} ({inv}): "
                                    f"gallery says {expected} images, "
                                    f"found {count}"
                                )

                            print(
                                f"   -> {title} | {count} images "
                                f"(gallery says {expected})"
                            )

                            # ---------------------------------
                            # DOWNLOAD IMAGES
                            # ---------------------------------

                            for idx, img_url in enumerate(
                                urls,
                                start=1,
                            ):

                                path = image_path(
                                    inv,
                                    idx,
                                )

                                row = {
                                    "Inventory Number": inv,
                                    "MPN": mpn,
                                    "Product Title": title,
                                    "Image #": idx,
                                    "File Name": path.name,
                                    "Source URL": img_url,
                                }

                                try:

                                    if (
                                        SKIP_EXISTING
                                        and path.exists()
                                    ):

                                        with Image.open(path) as existing:

                                            row.update(
                                                Width=existing.width,
                                                Height=existing.height,
                                                Status="Skipped (already exists)",
                                            )

                                    else:

                                        img, note = process_image(
                                            fetch_bytes(
                                                driver,
                                                img_url,
                                            )
                                        )

                                        save_image(
                                            img,
                                            path,
                                        )

                                        row.update(
                                            Width=img.width,
                                            Height=img.height,
                                            Status="OK",
                                            Note=note,
                                        )

                                        if img.width != TARGET_WIDTH:

                                            row["Status"] = (
                                                f"CHECK - width is "
                                                f"{img.width}, not "
                                                f"{TARGET_WIDTH}"
                                            )

                                    print(
                                        f"      {path.name} "
                                        f"{row.get('Width')}x"
                                        f"{row.get('Height')} "
                                        f"{row['Status']}"
                                    )

                                except Exception as e:

                                    logging.error(
                                        f"MPN {mpn} image {idx} failed: {e}"
                                    )

                                    row["Status"] = f"ERROR: {e}"

                                log_rows.append(row)

            except Exception as e:

                logging.error(
                    f"MPN {mpn} failed: {e}"
                )

                log_rows.append({
                    "Inventory Number": inv,
                    "MPN": mpn,
                    "Product Title": title,
                    "Status": f"ERROR: {e}",
                })

                flag(
                    "ERROR - check"
                )

            # ---------------------------------------------
            # PERIODIC EXCEL SAVE
            # ---------------------------------------------

            if (i + 1) % SAVE_LOG_EVERY == 0:

                save_log(
                    log_rows,
                    review_rows,
                )

    finally:

        # Save final results even if the scraper encounters
        # an error or is interrupted during execution.
        save_log(
            log_rows,
            review_rows,
        )

        if driver:
            cleanup(driver)


# ---------------- Run ----------------

if __name__ == "__main__":

    # TEST MODE:
    # Uncomment to test selected MPNs.
    #
    #main(test_mpns=["85360", "101450", "80156"])

    # FULL RUN:
    main()
