# ANZO Product Data Scraper

A Python web scraper that collects product data from [anzousa.com](https://anzousa.com) and exports it to a structured Excel file.

## What It Does

For every SKU in the input Excel file, the scraper:

1. Finds the product on anzousa.com (via the site's product list, then search)
2. Scrapes product details into **Sheet: Anzo Product**
   - Product Description, MPN, and all Specifications (in site order or A–Z)
   - Installation Video / PDF cells contain direct URLs
3. Scrapes the Vehicle Fitment table into **Sheet: Anzo Fitment**
   - Consecutive years are collapsed into ranges (e.g. `1992-1994`)
   - Trim maps to `Submodel1 (Custom)`
4. Logs SKUs not found or errors into **Sheet: Items Not Found**
5. Saves everything to `Anzo_Product_Data_<date>.xlsx` on the Desktop

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

1. Place your SKU list Excel file in the project folder (default: `ANZO_Scrape_List_2026-11-04.xlsx`)
2. The file must have an **Inventory Number** column (e.g. `ANZ-311302`)
   - Optional: `MPN` column (if absent, the number after the last hyphen is used)
   - Optional: `Description` column (used on the "Items Not Found" sheet)
3. Update `INPUT_FILE` in the script if your filename differs

---

## Usage

```bash
python anzo_scraper.py
```

- A browser window will open — solve any CAPTCHA manually, then press **Enter** to continue
- Output file is saved to your Desktop as `Anzo_Product_Data_YYYY-MM-DD.xlsx`

### Test Mode

To test with specific MPNs only, edit the bottom of the script:

```python
if __name__ == "__main__":
    main(test_mpns=["211014", "211153", "211015"])
```

---

## Configuration

Key settings at the top of the script:

| Constant | Default | Description |
|---|---|---|
| `INPUT_FILE` | `ANZO_Scrape_List_2026-11-04.xlsx` | Path to your SKU list |
| `SAVE_EVERY` | `5` | Auto-save interval (every N SKUs) |
| `USE_PRODUCT_LIST` | `True` | Read the full product list first (faster) |
| `SKIP_SEARCH_IF_NOT_IN_LIST` | `False` | Skip search if SKU not in product list |
| `SPEC_ORDER` | `site` | Spec column order: `site` (website order) or `alphabetical` |
| `MAX_CANDIDATES` | `5` | Max search results checked per SKU |

---

## Output

| Sheet | Contents |
|---|---|
| Anzo Product | Product Description, MPN, Specifications |
| Anzo Fitment | Make, Model, Year (ranges), Submodel per SKU |
| Items Not Found | SKUs with no match or scraping errors |

---

## Notes

- Logs are written to `anzo_scraper.log` in the project folder
- The scraper uses [SeleniumBase](https://github.com/seleniumbase/SeleniumBase) with undetected-chromedriver mode (`uc=True`) to handle bot protection
- SKU matching supports exact match and MPN + letter suffix (e.g. `111637X` matches `111637`)
