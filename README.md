# WARN.com Product Data Scraper

A Python web scraper that collects product data from [warn.com](https://www.warn.com) and exports it to a structured Excel file.

## What It Does

For every SKU in the input Excel file, the scraper:

1. Finds the product on warn.com
2. Scrapes product details into **Sheet: Warn Products**
   - Product Title, Summary, Description, HTML Description, Additional Description, Universal flag, and all Specs (A–Z)
3. Scrapes the Vehicle Compatibility List into **Sheet: Warn Fitment**
4. Scrapes Product Literature links into **Sheet: Product Literature URLs**
   - Installation Guides, Winching Manuals, Warranty Info, Replacement Parts
5. Logs SKUs not found on the site into **Sheet: Items Not Found**
6. Saves everything to `Warn_Product_Data_<date>.xlsx` on the Desktop

Progress is auto-saved every 5 SKUs to prevent data loss on crashes.

---

## Requirements

- Python 3.8+
- Google Chrome installed

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Setup

1. Place your SKU list Excel file in the project folder (default: `remaining_partnumber.xlsx`)
2. The file must have an **Inventory Number** column (e.g. `WAR-85360`)
   - Optional: `MPN` column (if absent, the number after the last hyphen is used)
   - Optional: `Description` column (used on the "Items Not Found" sheet)
3. Update `INPUT_FILE` in the script if your filename differs

---

## Usage

```bash
python warn_scraper_full_html.py
```

- A browser window will open — solve any CAPTCHA manually, then press **Enter** to continue
- Output file is saved to your Desktop as `Warn_Product_Data_YYYY-MM-DD.xlsx`

### Test Mode

To test with specific MPNs only, edit the bottom of the script:

```python
if __name__ == "__main__":
    main(test_mpns=["28500", "17801", "26502"])
```

---

## Configuration

Key settings at the top of the script:

| Constant | Default | Description |
|---|---|---|
| `INPUT_FILE` | `remaining_partnumber.xlsx` | Path to your SKU list |
| `SAVE_EVERY` | `5` | Auto-save interval (every N SKUs) |
| `SKIP_SEARCH_IF_NOT_IN_SITEMAP` | `False` | Skip search page if SKU not in sitemap |
| `ADD_USA_LINE` | `True` | Include "Made in USA" line in HTML description |
| `ADD_PROP65_FOOTER` | `True` | Include Prop 65 warning in HTML description |
| `SPEC_PAIRS_PER_ROW` | `3` | Spec columns per row in the HTML table (6 total) |

---

## Output

| Sheet | Contents |
|---|---|
| Warn Products | Title, Summary, Description, HTML, Specs |
| Warn Fitment | Year, Make, Model, Submodel per SKU |
| Product Literature URLs | PDF links by category |
| Items Not Found | SKUs with no match on warn.com |

---

## Notes

- Logs are written to `warn_scraper.log` in the project folder
- The scraper uses [SeleniumBase](https://github.com/seleniumbase/SeleniumBase) with undetected-chromedriver mode (`uc=True`) to handle bot protection
- "International" literature documents are skipped (English only)
