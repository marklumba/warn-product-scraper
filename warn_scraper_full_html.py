# """
# WARN.com scraper (built on the JEGS script framework)

# For every MPN in Warn_Scrape_List_*.xlsx:
#   1. Search the MPN on warn.com and open the product page
#   2. Scrape:  Product Title | Product Summary | Description (paragraphs + bullets) |
#               Product-Description (HTML) | Additional Description | Universal |
#               Specs (name = column, value = cell, A-Z)
#   3. Scrape the Vehicle Compatibility List (fitment table) -> Sheet "Warn Fitment"
#   4. Scrape the links in each product page's own "Product Literature" section (Winching /
#      Installation and Operational / Warranties / Replacement Parts URLs) -> Sheet "Product Literature URLs"
#      (only the documents listed on that product's page, as Matt asked; no "See All" page)
#   5. Save everything to Warn_Product_Data_<date>.xlsx on the Desktop

# Progress is saved every 5 products, so a crash does not lose the work.
# """
# import os
# import re
# import json
# import time
# import logging
# import tempfile
# import shutil
# from html import escape as html_escape
# from datetime import datetime
# from urllib.parse import quote, urljoin

# import pandas as pd
# from openpyxl import load_workbook
# from openpyxl.styles import Font, PatternFill
# from openpyxl.utils import get_column_letter
# from seleniumbase import Driver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import TimeoutException, NoSuchElementException

# # ---------------- Constants ----------------
# WEBSITE = "https://www.warn.com/"
# SEARCH_URL = "https://www.warn.com/search?q={mpn}"
# INPUT_FILE = r"Warn_Scrape_List_2026-10-27.xlsx"   # <-- full path if not in the same folder
# ELEMENT_WAIT_TIME = 20
# PAGE_PAUSE = 3
# SEARCH_RESULT_WAIT = 12       # seconds to wait for the search results to show the MPN
# SAVE_EVERY = 5

# # ---- Product-Description (HTML) ----
# # Layout: <ul> bullets </ul> + USA line + <p> paragraphs </p> + "Specifications:" + 6-column table + Prop 65 footer
# ADD_USA_LINE = True          # flag image + "Designed, Engineered, and Assembled in the USA"
# ADD_PROP65_FOOTER = True     # California Prop 65 warning at the end
# SPEC_PAIRS_PER_ROW = 3       # 3 pairs (name + value) per row = 6 columns, as Matt asked
# USA_LINE = (
#     '<img alt="" src="https://d3d71ba2asa5oz.cloudfront.net/12015492/images/us_flag_small_54h.jpg"> '
#     '<a>Designed, Engineered, and Assembled in the USA &nbsp;&nbsp;&nbsp;&nbsp;</a><br><br>'
# )
# PROP65_FOOTER = (
#     '<br><br><b>California Residents: <a href="http://www.P65Warnings.ca.gov">Prop 65 Warning</a></b> '
#     '<img alt="" src="https://d3d71ba2asa5oz.cloudfront.net/12015492/images/prop65_warning_image_small.jpg">'
# )
# # Inline styles (a website description field usually has no access to our CSS classes)
# SPEC_TABLE_STYLE = "border-collapse:collapse;width:100%;"
# SPEC_NAME_STYLE = ("width:14%;background-color:#f4f4f4;color:#555555;font-weight:bold;font-size:12px;"
#                    "text-transform:uppercase;padding:10px 12px;border:1px solid #e5e5e5;")
# SPEC_VALUE_STYLE = "width:19%;background-color:#ffffff;font-size:13px;padding:10px 12px;border:1px solid #e5e5e5;"

# # ---- Warn Fitment: model name corrections (exact match on the whole Model value) ----
# FITMENT_MODEL_FIXES = {"CRV": "CR-V"}

# # ---- Product literature (links inside each product page's own "Product Literature" section) ----
# LIT_OUTPUT_COLS = ["Part", "Installation Guides and Operational Manuals", "Winching Manuals", "Warranty Information", "Replacement Parts"]
# LIT_DATA_COLS = [c for c in LIT_OUTPUT_COLS if c != "Part"]
# LIT_PRODUCT_WAIT = 12            # seconds to wait for the product's literature links to appear
# LIT_SKIP_PATTERN = re.compile(r"\binternational\b", re.I)   # Matt: skip "International" documents (not English)
# # The section has headers (div.prod-lit__col-header). The header above a link decides its column:
# LIT_HEADING_KEYWORDS = [
#     (("replacement",), "Replacement Parts"),
#     (("warrant",), "Warranty Information"),
#     (("install", "operation"), "Installation Guides and Operational Manuals"),
#     (("winch",), "Winching Manuals"),
# ]

# logging.basicConfig(
#     level=logging.INFO,
#     format="%(asctime)s - %(levelname)s: %(message)s",
#     handlers=[logging.FileHandler("warn_scraper.log"), logging.StreamHandler()],
# )


# # ---------------- Driver ----------------
# def setup_driver():
#     user_data_dir = tempfile.mkdtemp()
#     driver = Driver(uc=True)
#     driver.user_data_dir = user_data_dir
#     driver.set_page_load_timeout(30)
#     return driver


# def wait_for_captcha(driver):
#     print("If a CAPTCHA / bot check appears, solve it manually.")
#     input("Press Enter to continue...")


# # ---------------- Input ----------------
# def load_mpn_list(path):
#     df = pd.read_excel(path, dtype=str)
#     df = df.iloc[:, :2]
#     df.columns = ["Inventory Number", "MPN"]
#     df = df.dropna(subset=["MPN"])
#     df["MPN"] = df["MPN"].str.strip()
#     return df.drop_duplicates(subset="MPN").reset_index(drop=True)


# # ---------------- Helpers ----------------
# def clean(text):
#     """Collapse whitespace / non-breaking spaces."""
#     return re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()


# def text_of(element):
#     """textContent also works for collapsed accordions (hidden elements)."""
#     return clean(element.get_attribute("textContent"))


# # ---------------- Search ----------------
# # Product index from the site's own sitemap: every product URL, independent of the search page
# _SITEMAP_JS = r"""
# const done = arguments[arguments.length - 1];
# const getText = u => fetch(u, {credentials: 'include'}).then(r => {
#   if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + u);
#   return r.text();
# });
# const locs = t => Array.from(t.matchAll(/<loc>([^<]+)<\/loc>/g)).map(m => m[1].replace(/&amp;/g, '&'));
# getText('/sitemap.xml').then(async idx => {
#   const maps = locs(idx).filter(u => u.indexOf('sitemap_products') !== -1);
#   let urls = [];
#   for (const m of maps) {
#     const t = await getText(m);
#     urls = urls.concat(locs(t).filter(u => u.indexOf('/products/') !== -1));
#   }
#   done(JSON.stringify(urls));
# }).catch(e => done('ERR:' + e));
# """


# def load_product_url_index(driver):
#     """List of every product URL on the site (from sitemap.xml). Empty list if it cannot be read."""
#     try:
#         driver.set_script_timeout(120)
#         res = driver.execute_async_script(_SITEMAP_JS)
#         if not isinstance(res, str) or res.startswith("ERR:"):
#             logging.warning(f"Product sitemap could not be read ({res}); the search page will be used instead")
#             return []
#         return [u.split("?")[0] for u in json.loads(res)]
#     except Exception as e:
#         logging.warning(f"Product sitemap could not be read ({e}); the search page will be used instead")
#         return []


# def url_from_index(url_index, mpn):
#     """Product URLs whose handle contains the MPN as its own word (e.g. m12-winch-17801), best match first."""
#     pat = re.compile(r"(?:^|-)" + re.escape(mpn.lower()) + r"(?:-|$)")
#     hits = [u for u in url_index if pat.search(u.split("/products/")[-1].lower())]
#     hits.sort(key=lambda u: (not u.lower().rstrip("/").endswith("-" + mpn.lower()), len(u)))
#     return hits


# # Shopify predictive-search endpoint: lists the products that match the typed MPN
# _SUGGEST_JS = """
# const q = arguments[0], done = arguments[arguments.length - 1];
# fetch('/search/suggest.json?q=' + encodeURIComponent(q) + '&resources[type]=product&resources[limit]=10',
#       {credentials: 'include'})
#   .then(r => r.json()).then(j => done(JSON.stringify(j))).catch(e => done('ERR:' + e));
# """


# def _product_links(driver, mpn):
#     """(all product links on the page, those whose URL contains the MPN) - snapshot of the page right now."""
#     hrefs = []
#     for a in driver.find_elements(By.CSS_SELECTOR, 'a[href*="/products/"]'):
#         h = (a.get_attribute("href") or "").split("?")[0]
#         if "/products/" in h and h not in hrefs:
#             hrefs.append(h)
#     return hrefs, [h for h in hrefs if mpn in h.split("/products/")[-1]]


# def _suggest_urls(driver, mpn):
#     """Product URLs returned by the site's own predictive search for this MPN."""
#     try:
#         driver.set_script_timeout(30)
#         res = driver.execute_async_script(_SUGGEST_JS, mpn)
#         if not isinstance(res, str) or res.startswith("ERR:"):
#             return []
#         products = json.loads(res).get("resources", {}).get("results", {}).get("products", [])
#         urls = [urljoin(WEBSITE, p.get("url", "")).split("?")[0] for p in products if p.get("url")]
#         return [u for u in urls if mpn in u.split("/products/")[-1]]
#     except Exception as e:
#         logging.warning(f"Predictive search failed for {mpn}: {e}")
#         return []


# def find_product_url(driver, mpn, url_index=None):
#     """
#     Return (product URL, status) for an MPN. Tried in this order:
#       1. the site's product sitemap (url_index) - does not depend on the search page at all
#       2. the search page, waiting until a result whose URL contains the MPN appears
#       3. the site's predictive search
#       4. last resort: first search result (the caller discards the data if the page lacks the MPN)
#     """
#     if url_index:
#         hits = url_from_index(url_index, mpn)
#         if hits:
#             if len(hits) > 1:
#                 logging.warning(f"MPN {mpn}: {len(hits)} products match in the sitemap, using {hits[0]} (others: {hits[1:3]})")
#             return hits[0], "OK (sitemap)"

#     driver.get(SEARCH_URL.format(mpn=quote(mpn)))
#     time.sleep(PAGE_PAUSE)

#     if "/products/" in driver.current_url:          # search redirected straight to the product
#         return driver.current_url.split("?")[0], "OK (direct)"

#     try:
#         WebDriverWait(driver, SEARCH_RESULT_WAIT).until(lambda d: _product_links(d, mpn)[1])
#     except TimeoutException:
#         pass
#     hrefs, exact = _product_links(driver, mpn)
#     if exact:
#         return exact[0], "OK"

#     suggested = _suggest_urls(driver, mpn)
#     if suggested:
#         return suggested[0], "OK (predictive search)"

#     logging.warning(f"MPN {mpn}: not found in sitemap, search results or predictive search")
#     if hrefs:
#         return hrefs[0], "CHECK - MPN not in URL (first result used)"
#     return "", "No search results"


# def scrape_title(driver):
#     for sel in ("div.product_title h1", "div.product_info-container h1", "h1"):
#         try:
#             return text_of(driver.find_element(By.CSS_SELECTOR, sel))
#         except NoSuchElementException:
#             continue
#     return ""


# def texts_from(driver, selectors, unique=True):
#     """Try each CSS selector in order; return the cleaned, non-empty texts of the first one that matches."""
#     for sel in selectors:
#         try:
#             found = [text_of(e) for e in driver.find_elements(By.CSS_SELECTOR, sel)]
#         except Exception:
#             continue
#         found = [t for t in found if t]
#         if unique:   # drop exact duplicates and texts already contained in a longer one (nested <p>)
#             found = [t for i, t in enumerate(found)
#                      if t not in found[:i] and not any(t != o and t in o for o in found)]
#         if found:
#             return found
#     return []


# def scrape_summary(driver):
#     """Product Summary: short text under the title (metafield rich text)."""
#     return "\n".join(texts_from(driver, [
#         'section[id^="ProductInfo"] div[class*="metafield-rich_text_field"] p',
#         'div.product__info-container div[class*="metafield-rich_text_field"] p',
#         'div[class*="metafield-rich_text_field"] p',
#     ]))


# def scrape_description_parts(driver):
#     """Returns (paragraphs, bullets) as two lists of clean text."""
#     paragraphs = texts_from(driver, [
#         "div.prod-desc__body p",
#         "div.prod-desc_body p",
#         "div.prod-desc p",
#     ])
#     bullets = texts_from(driver, [
#         "div.prod-desc__bullets li.feat-item",
#         "div.prod-desc_bullets li.feat-item",
#         "ul.feat-list li.feat-item",
#     ], unique=False)
#     return paragraphs, bullets


# def format_description(paragraphs, bullets):
#     """'Description' column: paragraphs followed by the feature bullets, as plain text."""
#     parts = []
#     if paragraphs:
#         parts.append("\n\n".join(paragraphs))
#     if bullets:
#         parts.append("\n".join(f"• {b}" for b in bullets))
#     return "\n\n".join(parts)


# def build_specs_table_html(specs):
#     """Specifications as an HTML table with SPEC_PAIRS_PER_ROW name/value pairs per row (6 columns),
#     A-Z reading left to right, then top to bottom. Returns '' when there are no specs."""
#     items = sorted(((n, v) for n, v in (specs or {}).items() if n and v), key=lambda kv: kv[0].lower())
#     if not items:
#         return ""
#     n = SPEC_PAIRS_PER_ROW
#     rows = []
#     for i in range(0, len(items), n):
#         chunk = items[i:i + n]
#         tds = "".join(
#             f'<td style="{SPEC_NAME_STYLE}">{html_escape(name, quote=False)}</td>'
#             f'<td style="{SPEC_VALUE_STYLE}">{html_escape(value, quote=False)}</td>'
#             for name, value in chunk
#         )
#         if len(chunk) < n:                                   # pad the last row to the full 6 columns
#             tds += f'<td colspan="{(n - len(chunk)) * 2}"></td>'
#         rows.append(f"<tr>{tds}</tr>")
#     return f'<table style="{SPEC_TABLE_STYLE}"><tbody>' + "".join(rows) + "</tbody></table>"


# def build_description_html(paragraphs, bullets, specs=None):
#     """
#     'Product-Description' column, Matt's layout:
#       <ul><li>bullets</li></ul> + USA line + <p>paragraph<br><br>paragraph</p>
#       + <b>Specifications:</b> + 6-column spec table + Prop 65 footer
#     Text is HTML-escaped (& < >); quotes and apostrophes stay as typed.
#     Returns '' when the product has no description at all.
#     """
#     if not paragraphs and not bullets:
#         return ""
#     out = ""
#     if bullets:
#         out += "<ul>" + "".join(f"<li>{html_escape(b, quote=False)}</li>" for b in bullets) + "</ul>"
#     if ADD_USA_LINE:
#         out += USA_LINE
#     if paragraphs:
#         out += "<p>" + "<br><br>".join(html_escape(p, quote=False) for p in paragraphs) + "</p>"
#     table = build_specs_table_html(specs)
#     if table:
#         out += "<b>Specifications:</b>" + table
#     if ADD_PROP65_FOOTER:
#         out += PROP65_FOOTER
#     return out


# def scrape_additional_description(driver):
#     """
#     Heading (e.g. 'XD9 Winch') + slpa-intro paragraphs from the accordion that holds the
#     product story. The 'Vehicle Compatibility List' accordion is ignored.
#     Returns '' when the product has no such text.
#     """
#     paragraphs = texts_from(driver, ["div.slpa-intro p"])
#     paragraphs = [p for p in paragraphs if "vehicle compatibility list" not in p.lower()]
#     if not paragraphs:
#         return ""

#     heading = ""
#     try:
#         h = driver.find_element(
#             By.XPATH,
#             "//details[.//div[contains(@class,'slpa')]]//h2[contains(@class,'accordion__title')]",
#         )
#         heading = text_of(h)
#     except NoSuchElementException:
#         pass
#     if "vehicle compatibility list" in heading.lower():
#         heading = ""

#     return "\n".join(([heading] if heading else []) + paragraphs)


# def scrape_specs(driver):
#     """Returns {spec name: spec value} -> becomes column headers / cell values."""
#     specs = {}
#     for item in driver.find_elements(By.CSS_SELECTOR, "div.specs-grid div.specs-item"):
#         try:
#             name = text_of(item.find_element(By.CSS_SELECTOR, "div.specs-name"))
#             value = text_of(item.find_element(By.CSS_SELECTOR, "div.specs-value"))
#         except NoSuchElementException:
#             continue                                  # empty spec item
#         if name:
#             specs[name] = value
#     return specs


# # ---------------- Fitment (Vehicle Compatibility List) ----------------
# def _cell_value(row_el, field_class):
#     """Extract fitment-cell__value text from a fitment row element."""
#     for sel in (
#         f"div.fitment-cell.fitment-field-{field_class} div.fitment-cell__value",
#         f'div[class*="fitment-field-{field_class}"] [class*="fitment-cell__value"]',
#     ):
#         els = row_el.find_elements(By.CSS_SELECTOR, sel)
#         if els:
#             return clean(els[0].get_attribute("textContent"))
#     return ""


# def _normalise_year(year_text):
#     """'1977 - 1986' -> '1977-1986'; single years unchanged."""
#     return re.sub(r"\s*-\s*", "-", year_text)


# def scrape_fitment(driver, inv):
#     """
#     Scrape the Vehicle Compatibility List table.
#     Returns a list of dicts: Part, Year, Make, Model, Submodel1-3 (Custom).
#     Uses textContent so hidden rows (initiallyHidden block) are captured without clicking.
#     """
#     containers = driver.find_elements(By.CSS_SELECTOR, "div#cm-fitment-table")
#     if not containers:
#         containers = driver.find_elements(By.CSS_SELECTOR, 'div[id*="fitment-table"]')
#     if not containers:
#         return []

#     fitment_rows = []
#     for block_sel in (
#         "div.fitment-table-body.fitment-table-fitment",
#         'div[class*="fitment-table-fitment"]',
#     ):
#         blocks = containers[0].find_elements(By.CSS_SELECTOR, block_sel)
#         if not blocks:
#             continue
#         for block in blocks:
#             for row_el in block.find_elements(By.CSS_SELECTOR, "div.fitment-row"):
#                 year_text = _cell_value(row_el, "year")
#                 make = _cell_value(row_el, "make")
#                 model = FITMENT_MODEL_FIXES.get(_cell_value(row_el, "model"), _cell_value(row_el, "model"))
#                 if not (year_text or make or model):
#                     continue
#                 # The site can show several ranges in one cell ("1974-2023, 2026-2027").
#                 # The upload format is ONE range per row, so split them into separate rows.
#                 years = [_normalise_year(y.strip()) for y in year_text.split(",") if y.strip()] or [""]
#                 for year in years:
#                     fitment_rows.append({
#                         "Part": inv,
#                         "Year": year,
#                         "Make": make,
#                         "Model": model,
#                         "Submodel1 (Custom)": "",
#                         "Submodel2 (Custom)": "",
#                         "Submodel3 (Custom)": "",
#                     })
#         if fitment_rows:   # first selector that yields results wins
#             break

#     return fitment_rows


# # ---------------- Product literature (only the links in the product page's own section) ----------------
# def _is_lit_url(href):
#     """literature.warn.com documents, plus the Shopify-CDN PDF (International Warranty Policy)."""
#     h = (href or "").strip()
#     return "literature.warn.com" in h or ("cdn.shopify.com" in h and ".pdf" in h.lower())


# def _norm_href(href):
#     """Normalise a URL for deduplication: strip query string, fragment, and trailing slash."""
#     return re.sub(r"[?#].*$", "", (href or "").strip()).rstrip("/")


# def _lit_column(header):
#     """Output column for a link based solely on the header text above it; None = unknown."""
#     t = (header or "").lower()
#     for keys, col in LIT_HEADING_KEYWORDS:
#         if any(k in t for k in keys):
#             return col
#     return None


# # Every link inside div.prod-lit, in page order, each with the column header above it
# _PROD_LIT_JS = """
# const root = document.querySelector('div.prod-lit');
# if (!root) return [];
# const out = []; let header = '';
# root.querySelectorAll('[class*="col-header"], a[href]').forEach(el => {
#   if (el.tagName === 'A') {
#     out.push({href: el.href, header: header, cls: el.className || '',
#               text: (el.textContent || '').replace(/\\s+/g, ' ').trim()});
#   } else {
#     header = (el.textContent || '').replace(/\\s+/g, ' ').trim();
#   }
# });
# return out;
# """


# def scrape_product_literature(driver, mpn=""):
#     """
#     URLs from the "Product Literature" section of the CURRENT product page (nothing from the
#     "See All Product Literature" page), grouped by output column:
#     {"Winching": [...], "Installation and Operational": [...], "Warranties": [...], "Replacement Parts": [...]}
#     """
#     result = {c: [] for c in LIT_DATA_COLS}

#     # make sure the accordion is open and in view (the links are read from the page either way)
#     try:
#         driver.execute_script(
#             "document.querySelectorAll('details').forEach(d => { if (d.querySelector('div.prod-lit')) d.open = true; });"
#             "const el = document.querySelector('div.prod-lit'); if (el) el.scrollIntoView({block: 'center'});"
#         )
#     except Exception:
#         pass

#     try:
#         WebDriverWait(driver, LIT_PRODUCT_WAIT).until(
#             lambda d: any(_is_lit_url(a.get_attribute("href"))
#                           for a in d.find_elements(By.CSS_SELECTOR, "div.prod-lit a[href]"))
#         )
#     except TimeoutException:
#         return result          # this product has no literature links

#     seen = set()   # global dedup across all columns using normalised URLs
#     unknown = []
#     for it in driver.execute_script(_PROD_LIT_JS) or []:
#         href = (it.get("href") or "").strip()
#         if not _is_lit_url(href):
#             continue                                   # "See All Product Literature", "View All Replacement Parts"
#         key = _norm_href(href)
#         if key in seen:
#             continue
#         seen.add(key)
#         header = it.get("header") or ""
#         col = _lit_column(header)
#         if re.search(r"part\s*#", it.get("text") or "", re.I):
#             col = "Replacement Parts"                  # rows that list "Part #: ..." are replacement parts
#         if col is None:
#             unknown.append((header, href))
#             continue
#         if LIT_SKIP_PATTERN.search(it.get("text") or ""):   # Matt: skip "International" documents (not English)
#             continue                                        # but still count them as seen so they are not repeated in another column
    
#         result[col].append(href)

#     if unknown:
#         logging.warning(f"MPN {mpn}: {len(unknown)} literature link(s) under an unknown header were skipped: "
#                         f"{unknown[:3]}")
#         try:
#             with open(f"prod_lit_debug_{mpn or 'product'}.html", "w", encoding="utf-8") as f:
#                 f.write(driver.page_source)
#         except Exception:
#             pass
#     return result


# # ---------------- Per-product scrape ----------------
# def scrape_product(driver, url, inv, mpn=""):
#     driver.get(url)
#     time.sleep(PAGE_PAUSE)
#     WebDriverWait(driver, ELEMENT_WAIT_TIME).until(
#         EC.presence_of_element_located((By.CSS_SELECTOR, "h1"))
#     )
#     fitment = scrape_fitment(driver, inv)
#     paragraphs, bullets = scrape_description_parts(driver)
#     specs = scrape_specs(driver)
#     return {
#         "Product Title": scrape_title(driver),
#         "Product Summary": scrape_summary(driver),
#         "Description": format_description(paragraphs, bullets),
#         "Product-Description": build_description_html(paragraphs, bullets, specs),
#         "Additional Description": scrape_additional_description(driver),
#         "Universal": "No" if fitment else "Yes",     # No = has Vehicle Fitment, Yes = no fitment list
#         "_specs": specs,
#         "_fitment": fitment,
#         "_literature": scrape_product_literature(driver, mpn),
#     }


# # ---------------- Output ----------------
# def build_dataframe(rows):
#     """Spec names become their own columns (union across all products), sorted A-Z from column I."""
#     base_cols = ["Inventory Number", "MPN",
#                  "Product Title", "Product Summary", "Description", "Product-Description",
#                  "Additional Description", "Universal"]
#     spec_cols = []
#     for r in rows:
#         for k in r.get("_specs", {}):
#             if k not in spec_cols:
#                 spec_cols.append(k)
#     spec_cols = sorted(spec_cols, key=str.lower)   # A-Z, ignores upper/lower case

#     flat = []
#     for r in rows:
#         d = {c: r.get(c, "") for c in base_cols}
#         for k in spec_cols:
#             d[k] = r.get("_specs", {}).get(k, "")
#         flat.append(d)
#     return pd.DataFrame(flat, columns=base_cols + spec_cols)


# def build_fitment_dataframe(rows):
#     fitment_cols = ["Part", "Year", "Make", "Model",
#                     "Submodel1 (Custom)", "Submodel2 (Custom)", "Submodel3 (Custom)"]
#     all_fitment = []
#     for r in rows:
#         all_fitment.extend(r.get("_fitment", []))
#     if not all_fitment:
#         return pd.DataFrame(columns=fitment_cols)
#     return pd.DataFrame(all_fitment, columns=fitment_cols)


# def build_literature_dataframe(rows):
#     """
#     One row per product per URL slot.
#     - Column with more than 1 URL: fills its own slots, blank after its last URL.
#     - Column with exactly 1 URL: repeated across all slots (matches Matt's sample).
#     - Column with 0 URLs: blank.
#     Rows per product = the largest URL count across the columns.
#     """
#     data_cols = [c for c in LIT_OUTPUT_COLS if c != "Part"]
#     out_rows = []
#     for r in rows:
#         inv = r.get("Inventory Number", "")
#         lit = r.get("_literature", {})
#         if not lit or not any(lit.get(c) for c in data_cols):
#             continue
#         n_slots = max((len(lit.get(c, [])) for c in data_cols), default=0)
#         for slot in range(n_slots):
#             row_out = {"Part": inv}
#             for col in data_cols:
#                 urls = lit.get(col, [])
#                 if not urls:
#                     row_out[col] = ""
#                 if slot < len(urls):
#                     row_out[col] = urls[slot]
#                 else:
#                     row_out[col] = ""
#                 # elif len(urls) == 1:
#                 #     row_out[col] = urls[0]        # single URL: repeat across all slots
#                 # elif slot < len(urls):
#                 #     row_out[col] = urls[slot]     # multiple URLs: fill own slots only
#                 # else:
#                 #     row_out[col] = ""             # exhausted: leave blank
#             out_rows.append(row_out)
#     if not out_rows:
#         return pd.DataFrame(columns=LIT_OUTPUT_COLS)
#     return pd.DataFrame(out_rows, columns=LIT_OUTPUT_COLS)


# def output_path():
#     return os.path.join(os.path.expanduser("~"), "Desktop",
#                         f"Warn_Product_Data_{datetime.now():%Y-%m-%d}.xlsx")


# def _format_sheet(ws):
#     """Bold grey header, auto-filter, freeze row 1, auto-width."""
#     ws.freeze_panes = "A2"
#     ws.auto_filter.ref = ws.dimensions
#     header_fill = PatternFill("solid", fgColor="C8C8C8")
#     for cell in ws[1]:
#         cell.fill = header_fill
#         cell.font = Font(bold=True)
#     for col_idx, column_cells in enumerate(ws.columns, start=1):
#         width = max(len(str(cell.value or "")) for cell in column_cells) + 2
#         ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width, 8), 50)


# def save_to_excel(rows, format_file=False):
#     df1 = build_dataframe(rows)
#     df_fit = build_fitment_dataframe(rows)
#     df_lit = build_literature_dataframe(rows)
#     path = output_path()

#     with pd.ExcelWriter(path, engine="openpyxl") as writer:
#         df1.to_excel(writer, index=False, sheet_name="Warn Products")
#         df_fit.to_excel(writer, index=False, sheet_name="Warn Fitment")
#         df_lit.to_excel(writer, index=False, sheet_name="Product Literature URLs")
#         wb = writer.book
#         # Year column as text so Excel doesn't turn ranges into dates
#         ws_fit = wb["Warn Fitment"]
#         year_col_idx = df_fit.columns.get_loc("Year") + 1
#         for row_idx in range(2, ws_fit.max_row + 1):
#             ws_fit.cell(row=row_idx, column=year_col_idx).number_format = "@"

#     if format_file:
#         wb = load_workbook(path)
#         for name in ("Warn Products", "Warn Fitment", "Product Literature URLs"):
#             _format_sheet(wb[name])
#             wb[name].freeze_panes = "B2"
#         wb.save(path)

#     print(f"Saved: {path}")


# def cleanup(driver):
#     try:
#         driver.quit()
#     except Exception as e:
#         print(f"Driver quit error (can be ignored): {e}")
#     if hasattr(driver, "user_data_dir") and os.path.exists(driver.user_data_dir):
#         shutil.rmtree(driver.user_data_dir, ignore_errors=True)


# # ---------------- Main ----------------
# def main(test_mpns=None):
#     """
#     test_mpns: optional list of MPN strings to run instead of the full list.
#     Example: main(test_mpns=["28500", "17801", "26502"])
#     """
#     driver = None
#     rows = []
#     try:
#         mpns = load_mpn_list(INPUT_FILE)
#         if test_mpns:
#             mpns = mpns[mpns["MPN"].isin(test_mpns)].reset_index(drop=True)
#         print(f"Loaded {len(mpns)} MPNs")

#         driver = setup_driver()
#         driver.get(WEBSITE)
#         wait_for_captcha(driver)

#         url_index = load_product_url_index(driver)
#         print(f"Product index: {len(url_index)} product URLs read from the sitemap")

#         for i, r in mpns.iterrows():
#             mpn, inv = r["MPN"], r["Inventory Number"]
#             row = {"Inventory Number": inv, "MPN": mpn, "Product URL": "", "Status": ""}
#             print(f"[{i + 1}/{len(mpns)}] MPN {mpn}")
#             url = ""
#             try:
#                 url, status = find_product_url(driver, mpn, url_index)
#                 row["Product URL"], row["Status"] = url, status
#                 if url:
#                     data = scrape_product(driver, url, inv, mpn)
#                     if mpn not in driver.page_source:
#                         # wrong product opened: keep NOTHING from it (a blank row is better than wrong data)
#                         row["Status"] += " | MPN not found on page - data discarded"
#                         logging.error(f"MPN {mpn} ({inv}): opened page '{data.get('Product Title', '')}' "
#                                       f"does not contain the MPN, data discarded: {url}")
#                     else:
#                         row.update(data)

#                     fitment = row.get("_fitment", [])
#                     if not fitment:
#                         logging.warning(f"MPN {mpn} ({inv}): no Vehicle Compatibility List found")
#                     else:
#                         first, last = fitment[0], fitment[-1]
#                         print(f"   -> fitment rows: {len(fitment)} | "
#                               f"first: {first['Year']} {first['Make']} {first['Model']} | "
#                               f"last:  {last['Year']} {last['Make']} {last['Model']}")

#                 lit = row.get("_literature", {})
#                 print("   -> literature | " + ", ".join(f"{c}: {len(lit.get(c, []))}" for c in LIT_DATA_COLS))
#                 if url and not any(lit.get(c) for c in LIT_DATA_COLS):
#                     logging.warning(f"MPN {mpn} ({inv}): no literature links on the product page")

#                 if not row["Status"].startswith("OK") or "MPN not found" in row["Status"]:
#                     logging.warning(f"MPN {mpn} needs a manual check: {row['Status']} {url}")
#                 print(f"   -> {row['Status']} | {row.get('Product Title', '')} | "
#                       f"{len(row.get('_specs', {}))} specs")
#             except Exception as e:
#                 logging.error(f"MPN {mpn} failed: {e}")
#                 row["Status"] = f"ERROR: {e}"
#             rows.append(row)

#             if len(rows) % SAVE_EVERY == 0:
#                 save_to_excel(rows)
#     finally:
#         if rows:
#             save_to_excel(rows, format_file=True)
#         if driver:
#             cleanup(driver)


# if __name__ == "__main__":
#     # ---- TEST MODE: comment the line below and uncomment main() to run the full list ----
#     # main(test_mpns=["28500", "17801", "26502"])
#     main()



"""
WARN.com scraper (built on the JEGS script framework)

For every SKU in Warn_Scrape_List_*.xlsx (Inventory Number, e.g. WAR-85360; the MPN is the number after the hyphen):
  1. Find the product on warn.com and open its page
  2. Scrape:  Product Title | Product Summary | Description (paragraphs + bullets) |
              Product-Description (HTML) | Additional Description | Universal |
              Specs (name = column, value = cell, A-Z)                         -> Sheet "Warn Products"
  3. Scrape the Vehicle Compatibility List (fitment table)                    -> Sheet "Warn Fitment"
  4. Scrape the links in each product page's own "Product Literature" section  -> Sheet "Product Literature URLs"
     (Installation Guides and Operational Manuals / Winching Manuals / Warranty Information /
      Replacement Parts; only the documents listed on that product's page; "International" ones skipped)
  5. SKUs that are not on the Warn site (or could not be scraped)             -> Sheet "Items Not Found"
  6. Save everything to ONE file: Warn_Product_Data_<date>.xlsx on the Desktop

Progress is saved every 5 SKUs, so a crash does not lose the work.
"""
import os
import re
import json
import time
import logging
import tempfile
import shutil
from html import escape as html_escape
from datetime import datetime
from urllib.parse import quote, urljoin

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ---------------- Constants ----------------
WEBSITE = "https://www.warn.com/"
SEARCH_URL = "https://www.warn.com/search?q={mpn}"
INPUT_FILE = r"remaining_partnumber.xlsx"   # <-- full path if not in the same folder
ELEMENT_WAIT_TIME = 20
PAGE_PAUSE = 3
SEARCH_RESULT_WAIT = 12       # seconds to wait for the search results to show the MPN
# True  = if the site's product list (sitemap) was read and the SKU is not in it, mark it "not found" straight away
#         (much faster when many SKUs are missing, about 15 s saved per missing SKU)
# False = also try the search page and the site's suggestions before giving up (slower, safest)
SKIP_SEARCH_IF_NOT_IN_SITEMAP = False
SAVE_EVERY = 5

# ---- Product-Description (HTML) ----
# Layout: <ul> bullets </ul> + USA line + <p> paragraphs </p> + "Specifications:" + 6-column table + Prop 65 footer
ADD_USA_LINE = True          # flag image + "Designed, Engineered, and Assembled in the USA"
ADD_PROP65_FOOTER = True     # California Prop 65 warning at the end
SPEC_PAIRS_PER_ROW = 3       # 3 pairs (name + value) per row = 6 columns, as Matt asked
USA_LINE = (
    '<img alt="" src="https://d3d71ba2asa5oz.cloudfront.net/12015492/images/us_flag_small_54h.jpg"> '
    '<a>Designed, Engineered, and Assembled in the USA &nbsp;&nbsp;&nbsp;&nbsp;</a><br><br>'
)
PROP65_FOOTER = (
    '<br><br><b>California Residents: <a href="http://www.P65Warnings.ca.gov">Prop 65 Warning</a></b> '
    '<img alt="" src="https://d3d71ba2asa5oz.cloudfront.net/12015492/images/prop65_warning_image_small.jpg">'
)
# Inline styles (a website description field usually has no access to our CSS classes)
SPEC_TABLE_STYLE = "border-collapse:collapse;width:100%;"
SPEC_NAME_STYLE = ("width:14%;background-color:#f4f4f4;color:#555555;font-weight:bold;font-size:12px;"
                   "text-transform:uppercase;padding:10px 12px;border:1px solid #e5e5e5;")
SPEC_VALUE_STYLE = "width:19%;background-color:#ffffff;font-size:13px;padding:10px 12px;border:1px solid #e5e5e5;"

# ---- Warn Fitment: model name corrections (exact match on the whole Model value) ----
FITMENT_MODEL_FIXES = {"CRV": "CR-V"}

# ---- Product literature (links inside each product page's own "Product Literature" section) ----
LIT_OUTPUT_COLS = ["Part", "Installation Guides and Operational Manuals", "Winching Manuals", "Warranty Information", "Replacement Parts"]
LIT_DATA_COLS = [c for c in LIT_OUTPUT_COLS if c != "Part"]
LIT_PRODUCT_WAIT = 12            # seconds to wait for the product's literature links to appear
# The section has headers (div.prod-lit__col-header). The header above a link decides its column:
LIT_SKIP_PATTERN = re.compile(r"\binternational\b", re.I)   # Matt: skip "International" documents (not English)
LIT_HEADING_KEYWORDS = [
    (("replacement",), "Replacement Parts"),
    (("warrant",), "Warranty Information"),
    (("install", "operation"), "Installation Guides and Operational Manuals"),
    (("winch",), "Winching Manuals"),
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s: %(message)s",
    handlers=[logging.FileHandler("warn_scraper.log"), logging.StreamHandler()],
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


# ---------------- Input ----------------
def load_mpn_list(path):
    """
    Reads the SKU list. Needs an 'Inventory Number' column (e.g. WAR-85360).
    MPN = the file's 'MPN' column if it has one, otherwise the number after the last hyphen of the
    Inventory Number (WAR-85360 -> 85360; anything without a number, e.g. WAR--FLATLINK-PROMO, stays as is).
    'Description' is kept so missing items can be identified on the "Items Not Found" sheet.
    """
    df = pd.read_excel(path, dtype=str)
    inv_col = next((c for c in df.columns if "inventory" in str(c).lower()), df.columns[0])
    mpn_col = next((c for c in df.columns if str(c).strip().lower() == "mpn"), None)
    desc_col = next((c for c in df.columns if str(c).strip().lower() == "description"), None)

    out = pd.DataFrame()
    out["Inventory Number"] = df[inv_col].str.strip()
    derived = out["Inventory Number"].str.extract(r"-(\d+)$")[0].fillna(out["Inventory Number"])
    out["MPN"] = df[mpn_col].str.strip() if mpn_col else derived
    out["Description"] = df[desc_col].fillna("") if desc_col else ""
    out = out.dropna(subset=["Inventory Number", "MPN"])
    out = out[out["MPN"].str.len() > 0]
    return out.drop_duplicates(subset="MPN").reset_index(drop=True)


# ---------------- Helpers ----------------
def clean(text):
    """Collapse whitespace / non-breaking spaces."""
    return re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()


def text_of(element):
    """textContent also works for collapsed accordions (hidden elements)."""
    return clean(element.get_attribute("textContent"))


# ---------------- Search ----------------
# Product index from the site's own sitemap: every product URL, independent of the search page
_SITEMAP_JS = r"""
const done = arguments[arguments.length - 1];
const getText = u => fetch(u, {credentials: 'include'}).then(r => {
  if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + u);
  return r.text();
});
const locs = t => Array.from(t.matchAll(/<loc>([^<]+)<\/loc>/g)).map(m => m[1].replace(/&amp;/g, '&'));
getText('/sitemap.xml').then(async idx => {
  const maps = locs(idx).filter(u => u.indexOf('sitemap_products') !== -1);
  let urls = [];
  for (const m of maps) {
    const t = await getText(m);
    urls = urls.concat(locs(t).filter(u => u.indexOf('/products/') !== -1));
  }
  done(JSON.stringify(urls));
}).catch(e => done('ERR:' + e));
"""


def load_product_url_index(driver):
    """List of every product URL on the site (from sitemap.xml). Empty list if it cannot be read."""
    try:
        driver.set_script_timeout(120)
        res = driver.execute_async_script(_SITEMAP_JS)
        if not isinstance(res, str) or res.startswith("ERR:"):
            logging.warning(f"Product sitemap could not be read ({res}); the search page will be used instead")
            return []
        return [u.split("?")[0] for u in json.loads(res)]
    except Exception as e:
        logging.warning(f"Product sitemap could not be read ({e}); the search page will be used instead")
        return []


def url_from_index(url_index, mpn):
    """Product URLs whose handle contains the MPN as its own word (e.g. m12-winch-17801), best match first."""
    pat = re.compile(r"(?:^|-)" + re.escape(mpn.lower()) + r"(?:-|$)")
    hits = [u for u in url_index if pat.search(u.split("/products/")[-1].lower())]
    hits.sort(key=lambda u: (not u.lower().rstrip("/").endswith("-" + mpn.lower()), len(u)))
    return hits


# Shopify predictive-search endpoint: lists the products that match the typed MPN
_SUGGEST_JS = """
const q = arguments[0], done = arguments[arguments.length - 1];
fetch('/search/suggest.json?q=' + encodeURIComponent(q) + '&resources[type]=product&resources[limit]=10',
      {credentials: 'include'})
  .then(r => r.json()).then(j => done(JSON.stringify(j))).catch(e => done('ERR:' + e));
"""


def _product_links(driver, mpn):
    """(all product links on the page, those whose URL contains the MPN) - snapshot of the page right now."""
    hrefs = []
    for a in driver.find_elements(By.CSS_SELECTOR, 'a[href*="/products/"]'):
        h = (a.get_attribute("href") or "").split("?")[0]
        if "/products/" in h and h not in hrefs:
            hrefs.append(h)
    return hrefs, [h for h in hrefs if mpn in h.split("/products/")[-1]]


def _suggest_urls(driver, mpn):
    """Product URLs returned by the site's own predictive search for this MPN."""
    try:
        driver.set_script_timeout(30)
        res = driver.execute_async_script(_SUGGEST_JS, mpn)
        if not isinstance(res, str) or res.startswith("ERR:"):
            return []
        products = json.loads(res).get("resources", {}).get("results", {}).get("products", [])
        urls = [urljoin(WEBSITE, p.get("url", "")).split("?")[0] for p in products if p.get("url")]
        return [u for u in urls if mpn in u.split("/products/")[-1]]
    except Exception as e:
        logging.warning(f"Predictive search failed for {mpn}: {e}")
        return []


def find_product_url(driver, mpn, url_index=None):
    """
    Return (product URL, status) for an MPN. Tried in this order:
      1. the site's product sitemap (url_index) - does not depend on the search page at all
      2. the search page, waiting until a result whose URL contains the MPN appears
      3. the site's predictive search
      4. last resort: first search result (the caller discards the data if the page lacks the MPN)
    """
    if url_index:
        hits = url_from_index(url_index, mpn)
        if hits:
            if len(hits) > 1:
                logging.warning(f"MPN {mpn}: {len(hits)} products match in the sitemap, using {hits[0]} (others: {hits[1:3]})")
            return hits[0], "OK (sitemap)"

    if url_index and SKIP_SEARCH_IF_NOT_IN_SITEMAP:
        return "", "Not in the site's product list (search skipped)"

    driver.get(SEARCH_URL.format(mpn=quote(mpn)))
    time.sleep(PAGE_PAUSE)

    if "/products/" in driver.current_url:          # search redirected straight to the product
        return driver.current_url.split("?")[0], "OK (direct)"

    try:
        WebDriverWait(driver, SEARCH_RESULT_WAIT).until(lambda d: _product_links(d, mpn)[1])
    except TimeoutException:
        pass
    hrefs, exact = _product_links(driver, mpn)
    if exact:
        return exact[0], "OK"

    suggested = _suggest_urls(driver, mpn)
    if suggested:
        return suggested[0], "OK (predictive search)"

    logging.warning(f"MPN {mpn}: not found in sitemap, search results or predictive search")
    if hrefs:
        return hrefs[0], "CHECK - MPN not in URL (first result used)"
    return "", "No search results"


def scrape_title(driver):
    for sel in ("div.product_title h1", "div.product_info-container h1", "h1"):
        try:
            return text_of(driver.find_element(By.CSS_SELECTOR, sel))
        except NoSuchElementException:
            continue
    return ""


def texts_from(driver, selectors, unique=True):
    """Try each CSS selector in order; return the cleaned, non-empty texts of the first one that matches."""
    for sel in selectors:
        try:
            found = [text_of(e) for e in driver.find_elements(By.CSS_SELECTOR, sel)]
        except Exception:
            continue
        found = [t for t in found if t]
        if unique:   # drop exact duplicates and texts already contained in a longer one (nested <p>)
            found = [t for i, t in enumerate(found)
                     if t not in found[:i] and not any(t != o and t in o for o in found)]
        if found:
            return found
    return []


def scrape_summary(driver):
    """Product Summary: short text under the title (metafield rich text)."""
    return "\n".join(texts_from(driver, [
        'section[id^="ProductInfo"] div[class*="metafield-rich_text_field"] p',
        'div.product__info-container div[class*="metafield-rich_text_field"] p',
        'div[class*="metafield-rich_text_field"] p',
    ]))


def scrape_description_parts(driver):
    """Returns (paragraphs, bullets) as two lists of clean text."""
    paragraphs = texts_from(driver, [
        "div.prod-desc__body p",
        "div.prod-desc_body p",
        "div.prod-desc p",
    ])
    bullets = texts_from(driver, [
        "div.prod-desc__bullets li.feat-item",
        "div.prod-desc_bullets li.feat-item",
        "ul.feat-list li.feat-item",
    ], unique=False)
    return paragraphs, bullets


def format_description(paragraphs, bullets):
    """'Description' column: paragraphs followed by the feature bullets, as plain text."""
    parts = []
    if paragraphs:
        parts.append("\n\n".join(paragraphs))
    if bullets:
        parts.append("\n".join(f"• {b}" for b in bullets))
    return "\n\n".join(parts)


def build_specs_table_html(specs):
    """Specifications as an HTML table with SPEC_PAIRS_PER_ROW name/value pairs per row (6 columns),
    A-Z reading left to right, then top to bottom. Returns '' when there are no specs."""
    items = sorted(((n, v) for n, v in (specs or {}).items() if n and v), key=lambda kv: kv[0].lower())
    if not items:
        return ""
    n = SPEC_PAIRS_PER_ROW
    rows = []
    for i in range(0, len(items), n):
        chunk = items[i:i + n]
        tds = "".join(
            f'<td style="{SPEC_NAME_STYLE}">{html_escape(name, quote=False)}</td>'
            f'<td style="{SPEC_VALUE_STYLE}">{html_escape(value, quote=False)}</td>'
            for name, value in chunk
        )
        if len(chunk) < n:                                   # pad the last row to the full 6 columns
            tds += f'<td colspan="{(n - len(chunk)) * 2}"></td>'
        rows.append(f"<tr>{tds}</tr>")
    return f'<table style="{SPEC_TABLE_STYLE}"><tbody>' + "".join(rows) + "</tbody></table>"


def build_description_html(paragraphs, bullets, specs=None):
    """
    'Product-Description' column, Matt's layout:
      <ul><li>bullets</li></ul> + USA line + <p>paragraph<br><br>paragraph</p>
      + <b>Specifications:</b> + 6-column spec table + Prop 65 footer
    Text is HTML-escaped (& < >); quotes and apostrophes stay as typed.
    Returns '' when the product has no description at all.
    """
    if not paragraphs and not bullets:
        return ""
    out = ""
    if bullets:
        out += "<ul>" + "".join(f"<li>{html_escape(b, quote=False)}</li>" for b in bullets) + "</ul>"
    if ADD_USA_LINE:
        out += USA_LINE
    if paragraphs:
        out += "<p>" + "<br><br>".join(html_escape(p, quote=False) for p in paragraphs) + "</p>"
    table = build_specs_table_html(specs)
    if table:
        out += "<b>Specifications:</b>" + table
    if ADD_PROP65_FOOTER:
        out += PROP65_FOOTER
    return out


def scrape_additional_description(driver):
    """
    Heading (e.g. 'XD9 Winch') + slpa-intro paragraphs from the accordion that holds the
    product story. The 'Vehicle Compatibility List' accordion is ignored.
    Returns '' when the product has no such text.
    """
    paragraphs = texts_from(driver, ["div.slpa-intro p"])
    paragraphs = [p for p in paragraphs if "vehicle compatibility list" not in p.lower()]
    if not paragraphs:
        return ""

    heading = ""
    try:
        h = driver.find_element(
            By.XPATH,
            "//details[.//div[contains(@class,'slpa')]]//h2[contains(@class,'accordion__title')]",
        )
        heading = text_of(h)
    except NoSuchElementException:
        pass
    if "vehicle compatibility list" in heading.lower():
        heading = ""

    return "\n".join(([heading] if heading else []) + paragraphs)


def scrape_specs(driver):
    """Returns {spec name: spec value} -> becomes column headers / cell values."""
    specs = {}
    for item in driver.find_elements(By.CSS_SELECTOR, "div.specs-grid div.specs-item"):
        try:
            name = text_of(item.find_element(By.CSS_SELECTOR, "div.specs-name"))
            value = text_of(item.find_element(By.CSS_SELECTOR, "div.specs-value"))
        except NoSuchElementException:
            continue                                  # empty spec item
        if name:
            specs[name] = value
    return specs


# ---------------- Fitment (Vehicle Compatibility List) ----------------
def _cell_value(row_el, field_class):
    """Extract fitment-cell__value text from a fitment row element."""
    for sel in (
        f"div.fitment-cell.fitment-field-{field_class} div.fitment-cell__value",
        f'div[class*="fitment-field-{field_class}"] [class*="fitment-cell__value"]',
    ):
        els = row_el.find_elements(By.CSS_SELECTOR, sel)
        if els:
            return clean(els[0].get_attribute("textContent"))
    return ""


def _normalise_year(year_text):
    """'1977 - 1986' -> '1977-1986'; single years unchanged."""
    return re.sub(r"\s*-\s*", "-", year_text)


def scrape_fitment(driver, inv):
    """
    Scrape the Vehicle Compatibility List table.
    Returns a list of dicts: Part, Year, Make, Model, Submodel1-3 (Custom).
    Uses textContent so hidden rows (initiallyHidden block) are captured without clicking.
    """
    containers = driver.find_elements(By.CSS_SELECTOR, "div#cm-fitment-table")
    if not containers:
        containers = driver.find_elements(By.CSS_SELECTOR, 'div[id*="fitment-table"]')
    if not containers:
        return []

    fitment_rows = []
    for block_sel in (
        "div.fitment-table-body.fitment-table-fitment",
        'div[class*="fitment-table-fitment"]',
    ):
        blocks = containers[0].find_elements(By.CSS_SELECTOR, block_sel)
        if not blocks:
            continue
        for block in blocks:
            for row_el in block.find_elements(By.CSS_SELECTOR, "div.fitment-row"):
                year_text = _cell_value(row_el, "year")
                make = _cell_value(row_el, "make")
                model = FITMENT_MODEL_FIXES.get(_cell_value(row_el, "model"), _cell_value(row_el, "model"))
                if not (year_text or make or model):
                    continue
                # The site can show several ranges in one cell ("1974-2023, 2026-2027").
                # The upload format is ONE range per row, so split them into separate rows.
                years = [_normalise_year(y.strip()) for y in year_text.split(",") if y.strip()] or [""]
                for year in years:
                    fitment_rows.append({
                        "Part": inv,
                        "Year": year,
                        "Make": make,
                        "Model": model,
                        "Submodel1 (Custom)": "",
                        "Submodel2 (Custom)": "",
                        "Submodel3 (Custom)": "",
                    })
        if fitment_rows:   # first selector that yields results wins
            break

    return fitment_rows


# ---------------- Product literature (only the links in the product page's own section) ----------------
def _is_lit_url(href):
    """literature.warn.com documents, plus the Shopify-CDN PDF."""
    h = (href or "").strip()
    return "literature.warn.com" in h or ("cdn.shopify.com" in h and ".pdf" in h.lower())


def _norm_href(href):
    """Normalise a URL for deduplication: strip query string, fragment, and trailing slash."""
    return re.sub(r"[?#].*$", "", (href or "").strip()).rstrip("/")


def _lit_column(header):
    """Output column for a link based on the header text above it; None = unknown."""
    t = (header or "").lower()
    for keys, col in LIT_HEADING_KEYWORDS:
        if any(k in t for k in keys):
            return col
    return None


# Every link inside div.prod-lit, in page order, each with the column header above it
_PROD_LIT_JS = """
const root = document.querySelector('div.prod-lit');
if (!root) return [];
const out = []; let header = '';
root.querySelectorAll('[class*="col-header"], a[href]').forEach(el => {
  if (el.tagName === 'A') {
    out.push({href: el.href, header: header, cls: el.className || '',
              text: (el.textContent || '').replace(/\\s+/g, ' ').trim()});
  } else {
    header = (el.textContent || '').replace(/\\s+/g, ' ').trim();
  }
});
return out;
"""


def scrape_product_literature(driver, mpn=""):
    """
    URLs from the "Product Literature" section of the CURRENT product page (nothing from the
    "See All Product Literature" page), grouped by the LIT_DATA_COLS output columns.
    Documents with "International" in the text are skipped; each URL is kept once per product.
    """
    result = {c: [] for c in LIT_DATA_COLS}

    # make sure the accordion is open and in view (the links are read from the page either way)
    try:
        driver.execute_script(
            "document.querySelectorAll('details').forEach(d => { if (d.querySelector('div.prod-lit')) d.open = true; });"
            "const el = document.querySelector('div.prod-lit'); if (el) el.scrollIntoView({block: 'center'});"
        )
    except Exception:
        pass

    try:
        WebDriverWait(driver, LIT_PRODUCT_WAIT).until(
            lambda d: any(_is_lit_url(a.get_attribute("href"))
                          for a in d.find_elements(By.CSS_SELECTOR, "div.prod-lit a[href]"))
        )
    except TimeoutException:
        return result          # this product has no literature links

    seen = set()   # dedup across all columns using normalised URLs
    unknown = []
    for it in driver.execute_script(_PROD_LIT_JS) or []:
        href = (it.get("href") or "").strip()
        if not _is_lit_url(href):
            continue                                   # "See All Product Literature", "View All Replacement Parts"
        if LIT_SKIP_PATTERN.search(it.get("text") or ""):
            continue                                   # Matt: skip "International" documents (not English)
        key = _norm_href(href)
        if key in seen:
            continue
        seen.add(key)
        header = it.get("header") or ""
        col = _lit_column(header)
        if re.search(r"part\s*#", it.get("text") or "", re.I):
            col = "Replacement Parts"                  # rows that list "Part #: ..." are replacement parts
        if col is None:
            unknown.append((header, href))
            continue
        result[col].append(href)

    if unknown:
        logging.warning(f"MPN {mpn}: {len(unknown)} literature link(s) under an unknown header were skipped: "
                        f"{unknown[:3]}")
        try:
            with open(f"prod_lit_debug_{mpn or 'product'}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
        except Exception:
            pass
    return result


# ---------------- Per-product scrape ----------------
def scrape_product(driver, url, inv, mpn=""):
    driver.get(url)
    time.sleep(PAGE_PAUSE)
    WebDriverWait(driver, ELEMENT_WAIT_TIME).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "h1"))
    )
    fitment = scrape_fitment(driver, inv)
    paragraphs, bullets = scrape_description_parts(driver)
    specs = scrape_specs(driver)
    return {
        "Product Title": scrape_title(driver),
        "Product Summary": scrape_summary(driver),
        "Description": format_description(paragraphs, bullets),
        "Product-Description": build_description_html(paragraphs, bullets, specs),
        "Additional Description": scrape_additional_description(driver),
        "Universal": "No" if fitment else "Yes",     # No = has Vehicle Fitment, Yes = no fitment list
        "_specs": specs,
        "_fitment": fitment,
        "_literature": scrape_product_literature(driver, mpn),
    }


# ---------------- Output ----------------
def build_dataframe(rows):
    """Spec names become their own columns (union across all products), sorted A-Z from column I."""
    base_cols = ["Inventory Number", "MPN",
                 "Product Title", "Product Summary", "Description", "Product-Description",
                 "Additional Description", "Universal"]
    rows = [r for r in rows if not r.get("Not Found")]     # SKUs not on the site go to "Items Not Found" instead
    spec_cols = []
    for r in rows:
        for k in r.get("_specs", {}):
            if k not in spec_cols:
                spec_cols.append(k)
    spec_cols = sorted(spec_cols, key=str.lower)   # A-Z, ignores upper/lower case

    flat = []
    for r in rows:
        d = {c: r.get(c, "") for c in base_cols}
        for k in spec_cols:
            d[k] = r.get("_specs", {}).get(k, "")
        flat.append(d)
    return pd.DataFrame(flat, columns=base_cols + spec_cols)


def build_fitment_dataframe(rows):
    fitment_cols = ["Part", "Year", "Make", "Model",
                    "Submodel1 (Custom)", "Submodel2 (Custom)", "Submodel3 (Custom)"]
    all_fitment = []
    for r in rows:
        all_fitment.extend(r.get("_fitment", []))
    if not all_fitment:
        return pd.DataFrame(columns=fitment_cols)
    return pd.DataFrame(all_fitment, columns=fitment_cols)


def build_literature_dataframe(rows):
    """
    One row per product per URL slot: each column lists its own URLs once, blank after its last URL.
    Rows per product = the largest URL count across the columns.
    """
    out_rows = []
    for r in rows:
        inv = r.get("Inventory Number", "")
        lit = r.get("_literature", {})
        if not lit or not any(lit.get(c) for c in LIT_DATA_COLS):
            continue
        n_slots = max(len(lit.get(c, [])) for c in LIT_DATA_COLS)
        for slot in range(n_slots):
            row_out = {"Part": inv}
            for col in LIT_DATA_COLS:
                urls = lit.get(col, [])
                row_out[col] = urls[slot] if slot < len(urls) else ""
            out_rows.append(row_out)
    if not out_rows:
        return pd.DataFrame(columns=LIT_OUTPUT_COLS)
    return pd.DataFrame(out_rows, columns=LIT_OUTPUT_COLS)


def build_not_found_dataframe(rows):
    """SKUs that are not on the Warn site (or could not be scraped), with the reason."""
    cols = ["Inventory Number", "MPN", "Description", "Reason", "Details", "Product URL Checked"]
    out = [{"Inventory Number": r.get("Inventory Number", ""), "MPN": r.get("MPN", ""),
            "Description": r.get("SKU Description", ""), "Reason": r.get("Not Found", ""),
            "Details": r.get("Not Found Details", ""), "Product URL Checked": r.get("Product URL", "")}
           for r in rows if r.get("Not Found")]
    return pd.DataFrame(out, columns=cols)


def output_path():
    return os.path.join(os.path.expanduser("~"), "Desktop",
                        f"Warn_Product_Data_{datetime.now():%Y-%m-%d}.xlsx")


def _format_sheet(ws):
    """Bold grey header, auto-filter, freeze row 1, auto-width."""
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    header_fill = PatternFill("solid", fgColor="C8C8C8")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        width = max(len(str(cell.value or "")) for cell in column_cells) + 2
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width, 8), 50)


def save_to_excel(rows, format_file=False):
    df1 = build_dataframe(rows)
    df_fit = build_fitment_dataframe(rows)
    df_lit = build_literature_dataframe(rows)
    df_nf = build_not_found_dataframe(rows)
    path = output_path()

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Warn Products")
        df_fit.to_excel(writer, index=False, sheet_name="Warn Fitment")
        df_lit.to_excel(writer, index=False, sheet_name="Product Literature URLs")
        df_nf.to_excel(writer, index=False, sheet_name="Items Not Found")
        wb = writer.book
        # Year column as text so Excel doesn't turn ranges into dates
        ws_fit = wb["Warn Fitment"]
        year_col_idx = df_fit.columns.get_loc("Year") + 1
        for row_idx in range(2, ws_fit.max_row + 1):
            ws_fit.cell(row=row_idx, column=year_col_idx).number_format = "@"

    if format_file:
        wb = load_workbook(path)
        for name in ("Warn Products", "Warn Fitment", "Product Literature URLs", "Items Not Found"):
            _format_sheet(wb[name])
            wb[name].freeze_panes = "B2"
        wb.save(path)

    print(f"Saved: {path}")


def cleanup(driver):
    try:
        driver.quit()
    except Exception as e:
        print(f"Driver quit error (can be ignored): {e}")
    if hasattr(driver, "user_data_dir") and os.path.exists(driver.user_data_dir):
        shutil.rmtree(driver.user_data_dir, ignore_errors=True)


# ---------------- Main ----------------
def main(test_mpns=None):
    """
    test_mpns: optional list of MPN strings to run instead of the full list.
    Example: main(test_mpns=["85360", "101450", "80156"])
    """
    driver = None
    rows = []
    try:
        mpns = load_mpn_list(INPUT_FILE)
        if test_mpns:
            mpns = mpns[mpns["MPN"].isin(test_mpns)].reset_index(drop=True)
        print(f"Loaded {len(mpns)} SKUs")

        driver = setup_driver()
        driver.get(WEBSITE)
        wait_for_captcha(driver)

        url_index = load_product_url_index(driver)
        print(f"Product index: {len(url_index)} product URLs read from the sitemap")

        for i, r in mpns.iterrows():
            mpn, inv = r["MPN"], r["Inventory Number"]
            row = {"Inventory Number": inv, "MPN": mpn, "SKU Description": r.get("Description", ""),
                   "Product URL": "", "Status": "", "Not Found": "", "Not Found Details": ""}
            print(f"[{i + 1}/{len(mpns)}] {inv} (MPN {mpn})")
            url = ""
            try:
                url, status = find_product_url(driver, mpn, url_index)
                row["Product URL"], row["Status"] = url, status

                if not url:
                    row["Not Found"] = "Not found on warn.com"
                    row["Not Found Details"] = status
                    logging.warning(f"{inv}: not found on warn.com ({status})")
                else:
                    if not status.startswith("OK"):
                        logging.warning(f"{inv} needs a manual check: {status} {url}")
                    data = scrape_product(driver, url, inv, mpn)

                    if mpn not in driver.page_source:
                        # wrong product opened: keep NOTHING from it
                        row["Not Found"] = "Not found on warn.com"
                        row["Not Found Details"] = (f"closest page ('{data.get('Product Title', '')}') "
                                                    f"does not contain this MPN")
                        logging.error(f"{inv}: opened page '{data.get('Product Title', '')}' does not "
                                      f"contain the MPN, data discarded: {url}")
                    else:
                        row.update(data)
                        fitment = row.get("_fitment", [])
                        if not fitment:
                            logging.warning(f"{inv}: no Vehicle Compatibility List found")
                        else:
                            first, last = fitment[0], fitment[-1]
                            print(f"   -> fitment rows: {len(fitment)} | "
                                  f"first: {first['Year']} {first['Make']} {first['Model']} | "
                                  f"last:  {last['Year']} {last['Make']} {last['Model']}")
                        lit = row.get("_literature", {})
                        print("   -> literature | " + ", ".join(f"{c}: {len(lit.get(c, []))}" for c in LIT_DATA_COLS))
                        if not any(lit.get(c) for c in LIT_DATA_COLS):
                            logging.warning(f"{inv}: no literature links on the product page")
                        print(f"   -> {row['Status']} | {row.get('Product Title', '')} | "
                              f"{len(row.get('_specs', {}))} specs")
            except Exception as e:
                logging.error(f"{inv} failed: {e}")
                row["Status"] = f"ERROR: {e}"
                row["Not Found"] = "ERROR while scraping - run this SKU again"
                row["Not Found Details"] = str(e)[:300]
            rows.append(row)

            if len(rows) % SAVE_EVERY == 0:
                save_to_excel(rows)
    finally:
        if rows:
            save_to_excel(rows, format_file=True)
            missing = sum(1 for r in rows if r.get("Not Found") == "Not found on warn.com")
            errors = sum(1 for r in rows if str(r.get("Not Found", "")).startswith("ERROR"))
            print(f"\nDone: {len(rows)} SKUs | scraped {len(rows) - missing - errors} | "
                  f"not found {missing} | errors {errors} (see sheet 'Items Not Found')")
        if driver:
            cleanup(driver)


if __name__ == "__main__":
    # ---- TEST MODE: comment the line below and uncomment main() to run the full list ----
    # main(test_mpns=["85360", "101450", "80156"])
    main()