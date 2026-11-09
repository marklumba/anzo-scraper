"""
ANZO (anzousa.com) scraper - SeleniumBase UC driver, manual CAPTCHA step,
checkpoint saves, and "Items Not Found" list.

For every SKU in the input list:
  1. Find the product on anzousa.com (product list, then search).
  2. Scrape the product page -> Sheet "Anzo Product":
       Product Description
       Specifications (same order/grouping as the website)
       SKU specification renamed to MPN
       Installation Video / PDF cells contain URLs
  3. Scrape the Fitment table -> Sheet "Anzo Fitment":
       Trim -> Submodel1 (Custom)
       Consecutive years -> one range
  4. Missing products / scraping errors -> Sheet "Items Not Found"
  5. Save everything to one Excel file.

Progress is saved every 5 SKUs.

Needs:
    pip install seleniumbase pandas openpyxl
"""

import os
import re
import json
import time
import logging
import tempfile
import shutil

from datetime import datetime
from urllib.parse import quote, urljoin, urlparse

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


# ---------------- Constants ----------------

WEBSITE = "https://anzousa.com/"
SEARCH_URL = "https://anzousa.com/search?q={mpn}&type=product"

INPUT_FILE = r"ANZO_Scrape_List_2026-11-04.xlsx"

ELEMENT_WAIT_TIME = 20
PAGE_PAUSE = 3
SEARCH_RESULT_WAIT = 12
MAX_CANDIDATES = 5
SAVE_EVERY = 5

USE_PRODUCT_LIST = True
SKIP_SEARCH_IF_NOT_IN_LIST = False

PRODUCT_SHEET = "Anzo Product"
FITMENT_SHEET = "Anzo Fitment"
NOT_FOUND_SHEET = "Items Not Found"

BASE_COLS = [
    "Inventory Number",
    "MPN",
    "Product Description",
]

# "site" = preserve the order of specifications on the ANZO website.
# "alphabetical" = sort specification columns A-Z.
SPEC_ORDER = "site"

FITMENT_COLS = [
    "Part",
    "Make",
    "Model",
    "Year",
    "Submodel1 (Custom)",
    "Submodel2 (Custom)",
    "Submodel3 (Custom)",
    "Notes",
]

SPEC_RENAMES = {
    "sku": "MPN"
}


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s: %(message)s",
    handlers=[
        logging.FileHandler("anzo_scraper.log"),
        logging.StreamHandler()
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

    if hasattr(driver, "user_data_dir") and os.path.exists(driver.user_data_dir):
        shutil.rmtree(driver.user_data_dir, ignore_errors=True)


# ---------------- Input ----------------

def load_mpn_list(path):
    """
    Reads the SKU list.

    Requires an Inventory Number column.
    MPN is taken from the MPN column if available;
    otherwise, it is derived from the Inventory Number.
    """

    df = pd.read_excel(path, dtype=str)

    inv_col = next(
        (c for c in df.columns if "inventory" in str(c).lower()),
        df.columns[0]
    )

    mpn_col = next(
        (c for c in df.columns if str(c).strip().lower() == "mpn"),
        None
    )

    desc_col = next(
        (c for c in df.columns if str(c).strip().lower() == "description"),
        None
    )

    out = pd.DataFrame()

    out["Inventory Number"] = df[inv_col].str.strip()

    derived = (
        out["Inventory Number"]
        .str.extract(r"-(\d+)$")[0]
        .fillna(out["Inventory Number"])
    )

    out["MPN"] = df[mpn_col].str.strip() if mpn_col else derived

    out["Description"] = (
        df[desc_col].fillna("") if desc_col else ""
    )

    out = out.dropna(subset=["Inventory Number", "MPN"])
    out = out[out["MPN"].str.len() > 0]

    return out.drop_duplicates(subset="MPN").reset_index(drop=True)


# ---------------- Helpers ----------------

def clean(text):
    """Collapse whitespace and non-breaking spaces."""

    return re.sub(
        r"\s+",
        " ",
        (text or "").replace("\xa0", " ")
    ).strip()


def product_path(url):
    """
    Normalize Shopify product URLs.
    """

    p = urlparse(urljoin(WEBSITE, url))

    path = re.sub(
        r"^/collections/[^/]+/products/",
        "/products/",
        p.path
    )

    return f"{p.scheme}://{p.netloc}{path}"


def sku_matches(sku, mpn):
    """
    Match exact MPN or MPN followed by letters only.
    Example: 111637X matches 111637.
    """

    s = (sku or "").strip().upper()
    m = (mpn or "").strip().upper()

    if not s or not m:
        return False

    return (
        s == m
        or (
            s.startswith(m)
            and not re.search(r"\d", s[len(m):])
        )
    )


# ---------------- Finding the product ----------------

_PRODUCTS_JSON_JS = r"""
const done = arguments[arguments.length - 1];

(async () => {
  const out = [];

  for (let page = 1; page <= 200; page++) {
    const r = await fetch(
      '/products.json?limit=250&page=' + page,
      {credentials: 'include'}
    );

    if (!r.ok) throw new Error('HTTP ' + r.status);

    const j = await r.json();

    if (!j.products || !j.products.length) break;

    for (const p of j.products) {
      for (const v of (p.variants || [])) {
        if (v.sku) out.push([v.sku, p.handle]);
      }
    }
  }

  done(JSON.stringify(out));

})().catch(e => done('ERR:' + e));
"""


_PRODUCT_SKUS_JS = r"""
const url = arguments[0];
const done = arguments[arguments.length - 1];

fetch(url.split('?')[0] + '.js', {credentials: 'include'})
  .then(r => {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.json();
  })
  .then(p => done(JSON.stringify(
    (p.variants || []).map(v => v.sku || '')
  )))
  .catch(e => done('ERR:' + e));
"""


_SUGGEST_JS = """
const q = arguments[0];
const done = arguments[arguments.length - 1];

fetch(
  '/search/suggest.json?q=' +
  encodeURIComponent(q) +
  '&resources[type]=product&resources[limit]=10',
  {credentials: 'include'}
)
  .then(r => r.json())
  .then(j => done(JSON.stringify(j)))
  .catch(e => done('ERR:' + e));
"""


def load_sku_index(driver):
    """
    Returns a dictionary:
        {SKU: product URL}
    """

    try:
        driver.set_script_timeout(300)

        res = driver.execute_async_script(_PRODUCTS_JSON_JS)

        if not isinstance(res, str) or res.startswith("ERR:"):
            logging.warning(
                f"Product list could not be read ({res}); "
                "the search will be used instead"
            )
            return {}

        return {
            sku.strip().upper(): f"{WEBSITE}products/{handle}"
            for sku, handle in json.loads(res)
        }

    except Exception as e:
        logging.warning(
            f"Product list could not be read ({e}); "
            "the search will be used instead"
        )
        return {}


def url_from_sku_index(sku_index, mpn):
    """Find an exact SKU first, then an MPN with letter suffix."""

    m = (mpn or "").strip().upper()

    if m in sku_index:
        return sku_index[m]

    for sku, url in sku_index.items():
        if sku_matches(sku, m):
            return url

    return ""


def _suggest_urls(driver, mpn):
    try:
        driver.set_script_timeout(30)

        res = driver.execute_async_script(_SUGGEST_JS, mpn)

        if not isinstance(res, str) or res.startswith("ERR:"):
            return []

        products = (
            json.loads(res)
            .get("resources", {})
            .get("results", {})
            .get("products", [])
        )

        return [
            product_path(p["url"])
            for p in products
            if p.get("url")
        ]

    except Exception as e:
        logging.warning(
            f"Predictive search failed for {mpn}: {e}"
        )
        return []


def _search_page_urls(driver, mpn):
    """Find product links on the search results page."""

    driver.get(SEARCH_URL.format(mpn=quote(mpn)))
    time.sleep(PAGE_PAUSE)

    if "/products/" in driver.current_url:
        return [product_path(driver.current_url)]

    sel = (
        '#MainContent a[href*="/products/"], '
        'main a[href*="/products/"]'
    )

    try:
        WebDriverWait(
            driver,
            SEARCH_RESULT_WAIT
        ).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, sel)) > 0
        )

    except TimeoutException:
        return []

    urls = []

    for a in driver.find_elements(By.CSS_SELECTOR, sel):
        u = product_path(a.get_attribute("href") or "")

        if "/products/" in u and u not in urls:
            urls.append(u)

    return urls


def verify_candidate(driver, url, mpn):
    """
    Returns:
        True  = SKU verified
        False = SKU does not match
        None  = verification unavailable
    """

    try:
        driver.set_script_timeout(30)

        res = driver.execute_async_script(
            _PRODUCT_SKUS_JS,
            url
        )

        if not isinstance(res, str) or res.startswith("ERR:"):
            return None

        return any(
            sku_matches(s, mpn)
            for s in json.loads(res)
        )

    except Exception:
        return None


def find_product_url(driver, mpn, sku_index=None):
    """
    Search order:
      1. Product list
      2. Predictive search
      3. Search results page
      4. First unchecked candidate as a last resort
    """

    if sku_index:
        hit = url_from_sku_index(sku_index, mpn)

        if hit:
            return hit, "OK (product list)"

        if SKIP_SEARCH_IF_NOT_IN_LIST:
            return "", "Not in the site's product list (search skipped)"

    tried = set()
    unchecked = ""

    for source in (_suggest_urls, _search_page_urls):

        for url in source(driver, mpn)[:MAX_CANDIDATES]:

            if url in tried:
                continue

            tried.add(url)

            ok = verify_candidate(driver, url, mpn)

            if ok:
                return url, "OK (search)"

            if ok is None and not unchecked:
                unchecked = url

    if unchecked:
        return (
            unchecked,
            "CHECK - SKU could not be verified (first result used)"
        )

    logging.warning(
        f"MPN {mpn}: no matching product found"
    )

    return "", "No search results"


# ---------------- Scraping the product page ----------------

_PRODUCT_JS = r"""
const txt = el => el
  ? (el.textContent || '').replace(/[\s\u00a0]+/g, ' ').trim()
  : '';

const tab = name => Array.from(
  document.querySelectorAll('details')
).find(d => {
  const h = d.querySelector('summary h3, summary h2, summary');

  return h && txt(h).toLowerCase() === name;
});

const body = d => d
  ? (
      d.querySelector('div.rte') ||
      d.querySelector('div[class*="content"]') ||
      d
    )
  : null;

// Get the actual URL from a link.
const hrefOf = a => {
  const h = (a.getAttribute('href') || '').trim();

  return /^https?:\/\//i.test(h) ? h : a.href;
};

const out = {
  title: '',
  specs: [],
  fitment: {
    headers: [],
    rows: []
  }
};

// Product title.
out.title = txt(
  document.querySelector('h1.product_title') ||
  document.querySelector('div.product_info-container h1') ||
  document.querySelector('h1')
);

// Product Description extraction has been removed.

// Specifications.
const specTab = tab('specifications');

if (specTab && !specTab.open) {
  specTab.open = true;
}

const specRoot = specTab
  ? (body(specTab) || specTab)
  : null;

if (specRoot) {

  specRoot.querySelectorAll('table tr').forEach(tr => {

    const cells = Array.from(
      tr.querySelectorAll('td, th')
    );

    if (cells.length >= 2) {

      // Link cells contain their actual URL.
      const links = Array.from(
        cells[1].querySelectorAll('a[href]')
      )
        .filter(a => !/^(javascript:|#)/i.test(
          (a.getAttribute('href') || '').trim()
        ))
        .map(hrefOf);

      out.specs.push([
        txt(cells[0]),
        links.length
          ? links.join(' | ')
          : txt(cells[1])
      ]);
    }
  });
}

// Fitment table.
const fit = document.querySelector(
  '#fitmentTableWrapper table'
);

if (fit) {

  out.fitment.headers = Array.from(
    fit.querySelectorAll('thead th')
  ).map(txt);

  fit.querySelectorAll('tbody tr').forEach(tr => {

    const cells = Array.from(
      tr.querySelectorAll('td')
    ).map(txt);

    if (cells.some(c => c)) {
      out.fitment.rows.push(cells);
    }
  });
}

return out;
"""


def parse_specs(pairs):
    """
    Convert specification pairs into a dictionary.

    SKU is renamed to MPN.
    First occurrence of a specification name wins.
    """

    specs = {}

    for name, value in pairs or []:

        name = clean(name)

        if not name:
            continue

        name = SPEC_RENAMES.get(name.lower(), name)

        specs.setdefault(name, clean(value))

    return specs


def parse_fitment(fit, inv):
    """
    Convert the website fitment table into FITMENT_COLS.
    Trim maps to Submodel1 (Custom).
    """

    headers = [
        clean(h).lower()
        for h in (fit or {}).get("headers", [])
    ]

    rows = []

    for cells in (fit or {}).get("rows", []):

        rec = {
            c: ""
            for c in FITMENT_COLS
        }

        rec["Part"] = inv

        header_map = {}

        for h in headers:

            if h == "trim":
                header_map[h] = "Submodel1 (Custom)"

            else:
                for col in FITMENT_COLS[1:]:

                    if h == col.lower():
                        header_map[h] = col
                        break

        for h, col in header_map.items():

            idx = headers.index(h)

            rec[col] = (
                clean(cells[idx])
                if idx < len(cells)
                else ""
            )

        if not headers:

            site_order = [
                "Make",
                "Model",
                "Year",
                "Submodel1 (Custom)",
                "Notes"
            ]

            for i, col in enumerate(site_order):
                rec[col] = (
                    clean(cells[i])
                    if i < len(cells)
                    else ""
                )

        if any(rec[c] for c in FITMENT_COLS[1:]):
            rows.append(rec)

    return rows


def scrape_product(driver, url, inv):
    """Scrape product title, specifications, and fitment."""

    driver.get(url)
    time.sleep(PAGE_PAUSE)

    WebDriverWait(
        driver,
        ELEMENT_WAIT_TIME
    ).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "h1")
        )
    )

    data = driver.execute_script(_PRODUCT_JS) or {}

    return {
        "Product Description": clean(data.get("title", "")),
        "_specs": parse_specs(data.get("specs")),
        "_fitment": parse_fitment(data.get("fitment"), inv),
    }


# ---------------- Output ----------------

def merge_spec_order(master, names):
    """
    Add new specification names while preserving their order
    relative to specifications already found on the website.
    """

    prev = None

    for i, name in enumerate(names):

        if name in master:
            prev = name
            continue

        if prev is not None:

            master.insert(
                master.index(prev) + 1,
                name
            )

        else:

            nxt = next(
                (
                    n for n in names[i + 1:]
                    if n in master
                ),
                None
            )

            if nxt is not None:
                master.insert(master.index(nxt), name)
            else:
                master.append(name)

        prev = name

    return master


def build_product_dataframe(rows):
    """
    Excel layout:

    Inventory Number | MPN | Product Description | Specifications
    """

    rows = [
        r for r in rows
        if not r.get("Not Found")
    ]

    if SPEC_ORDER == "alphabetical":

        spec_cols = sorted(
            {
                k
                for r in rows
                for k in r.get("_specs", {})
                if k != "MPN"
            },
            key=str.lower
        )

    else:

        spec_cols = []

        for r in rows:

            merge_spec_order(
                spec_cols,
                [
                    k
                    for k in r.get("_specs", {})
                    if k != "MPN"
                ]
            )

    flat = []

    for r in rows:

        d = {
            c: r.get(c, "")
            for c in BASE_COLS
        }

        for k in spec_cols:
            d[k] = r.get("_specs", {}).get(k, "")

        flat.append(d)

    return pd.DataFrame(
        flat,
        columns=BASE_COLS + spec_cols
    )


# def _compress_years(df):
#     """
#     Collapse consecutive year rows for the same
#     Part/Make/Model/Submodels/Notes into a year range.
#     """

#     if df.empty:
#         return df

#     key_cols = [
#         "Part",
#         "Make",
#         "Model",
#         "Submodel1 (Custom)",
#         "Submodel2 (Custom)",
#         "Submodel3 (Custom)",
#         "Notes"
#     ]

#     out = []

#     for key, grp in df.groupby(
#         key_cols,
#         sort=False,
#         dropna=False
#     ):

#         years = []

#         for y in grp["Year"]:

#             try:
#                 years.append(int(str(y).strip()))

#             except (ValueError, TypeError):
#                 years.append(str(y).strip())

#         int_years = sorted(
#             set(y for y in years if isinstance(y, int))
#         )

#         other_years = [
#             y for y in years
#             if not isinstance(y, int)
#         ]

#         ranges = []

#         if int_years:

#             start = end = int_years[0]

#             for y in int_years[1:]:

#                 if y == end + 1:
#                     end = y

#                 else:

#                     ranges.append(
#                         str(start)
#                         if start == end
#                         else f"{start}-{end}"
#                     )

#                     start = end = y

#             ranges.append(
#                 str(start)
#                 if start == end
#                 else f"{start}-{end}"
#             )

#         year_str = ", ".join(
#             ranges + [str(y) for y in other_years]
#         )

#         rec = dict(
#             zip(
#                 key_cols,
#                 key if isinstance(key, tuple) else (key,)
#             )
#         )

#         rec["Year"] = year_str

#         out.append(rec)

#     return pd.DataFrame(
#         out,
#         columns=FITMENT_COLS
#     )



def _compress_years(df):
    """
    One row per year range.
    Consecutive years are joined into a range (1992-1994); a gap in the years starts a NEW ROW,
    so 1973, 1974, 1979 ... 1991 becomes two rows: 1973-1974 and 1979-1991.
    """
    if df.empty:
        return df

    key_cols = [
        "Part", "Make", "Model",
        "Submodel1 (Custom)", "Submodel2 (Custom)", "Submodel3 (Custom)",
        "Notes",
    ]
    out = []

    for key, grp in df.groupby(key_cols, sort=False, dropna=False):

        years = []
        for y in grp["Year"]:
            try:
                years.append(int(str(y).strip()))
            except (ValueError, TypeError):
                years.append(str(y).strip())

        int_years = sorted(set(y for y in years if isinstance(y, int)))
        other_years = list(dict.fromkeys(y for y in years if not isinstance(y, int) and y))

        ranges = []
        if int_years:
            start = end = int_years[0]
            for y in int_years[1:]:
                if y == end + 1:
                    end = y
                else:
                    ranges.append(str(start) if start == end else f"{start}-{end}")
                    start = end = y
            ranges.append(str(start) if start == end else f"{start}-{end}")

        base = dict(zip(key_cols, key if isinstance(key, tuple) else (key,)))

        # one output row for every range (and for every non-numeric year text)
        for year_str in ranges + other_years:
            out.append({**base, "Year": year_str})

    return pd.DataFrame(out, columns=FITMENT_COLS)

def build_fitment_dataframe(rows):

    allrows = [
        rec
        for r in rows
        for rec in r.get("_fitment", [])
    ]

    if not allrows:
        return pd.DataFrame(
            columns=FITMENT_COLS
        )

    df = pd.DataFrame(
        allrows,
        columns=FITMENT_COLS
    )

    return _compress_years(df)


def build_not_found_dataframe(rows):

    cols = [
        "Inventory Number",
        "MPN",
        "Description",
        "Reason",
        "Details",
        "Product URL Checked"
    ]

    out = [
        {
            "Inventory Number": r.get("Inventory Number", ""),
            "MPN": r.get("MPN", ""),
            "Description": r.get("SKU Description", ""),
            "Reason": r.get("Not Found", ""),
            "Details": r.get("Not Found Details", ""),
            "Product URL Checked": r.get("Product URL", "")
        }
        for r in rows
        if r.get("Not Found")
    ]

    return pd.DataFrame(
        out,
        columns=cols
    )


def output_path():

    return os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        f"Anzo_Product_Data_{datetime.now():%Y-%m-%d}.xlsx"
    )


def _format_sheet(ws):
    """Format header, filters, freeze panes, and column widths."""

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    fill = PatternFill(
        "solid",
        fgColor="C8C8C8"
    )

    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(bold=True)

    for col_idx, cells in enumerate(
        ws.columns,
        start=1
    ):

        width = (
            max(
                len(str(c.value or ""))
                for c in cells
            )
            + 2
        )

        ws.column_dimensions[
            get_column_letter(col_idx)
        ].width = min(
            max(width, 8),
            50
        )


def save_to_excel(rows, format_file=False):

    df_prod = build_product_dataframe(rows)
    df_fit = build_fitment_dataframe(rows)
    df_nf = build_not_found_dataframe(rows)

    path = output_path()

    with pd.ExcelWriter(
        path,
        engine="openpyxl"
    ) as writer:

        df_prod.to_excel(
            writer,
            index=False,
            sheet_name=PRODUCT_SHEET
        )

        df_fit.to_excel(
            writer,
            index=False,
            sheet_name=FITMENT_SHEET
        )

        df_nf.to_excel(
            writer,
            index=False,
            sheet_name=NOT_FOUND_SHEET
        )

        # Keep Year as text.
        ws_fit = writer.book[FITMENT_SHEET]

        year_col = FITMENT_COLS.index("Year") + 1

        for r in range(2, ws_fit.max_row + 1):

            ws_fit.cell(
                row=r,
                column=year_col
            ).number_format = "@"

    if format_file:

        wb = load_workbook(path)

        for ws in wb.worksheets:
            _format_sheet(ws)

        wb.save(path)

    print(f"Saved: {path}")


# ---------------- Main ----------------

def main(test_mpns=None):
    """
    test_mpns:
        Optional list of MPN strings to run instead of the full list.

    Example:
        main(test_mpns=["111506", "511032", "211140"])
    """

    driver = None
    rows = []

    try:

        mpns = load_mpn_list(INPUT_FILE)

        if test_mpns:

            mpns = mpns[
                mpns["MPN"].isin(test_mpns)
            ].reset_index(drop=True)

        print(f"Loaded {len(mpns)} SKUs")

        driver = setup_driver()

        driver.get(WEBSITE)

        wait_for_captcha(driver)

        sku_index = (
            load_sku_index(driver)
            if USE_PRODUCT_LIST
            else {}
        )

        print(
            f"Product list: {len(sku_index)} SKUs read from the site"
            + (
                ""
                if sku_index
                else " (not available - the search will be used)"
            )
        )

        for i, r in mpns.iterrows():

            mpn = r["MPN"]
            inv = r["Inventory Number"]

            row = {
                "Inventory Number": inv,
                "MPN": mpn,
                "SKU Description": r.get("Description", ""),
                "Product URL": "",
                "Status": "",
                "Not Found": "",
                "Not Found Details": ""
            }

            print(
                f"[{i + 1}/{len(mpns)}] {inv} (MPN {mpn})"
            )

            try:

                url, status = find_product_url(
                    driver,
                    mpn,
                    sku_index
                )

                row["Product URL"] = url
                row["Status"] = status

                if not url:

                    row["Not Found"] = (
                        "Not found on anzousa.com"
                    )

                    row["Not Found Details"] = status

                    logging.warning(
                        f"{inv}: not found on anzousa.com ({status})"
                    )

                else:

                    if not status.startswith("OK"):

                        logging.warning(
                            f"{inv} needs a manual check: {status} {url}"
                        )

                    data = scrape_product(
                        driver,
                        url,
                        inv
                    )

                    if mpn not in driver.page_source:

                        # Wrong product opened: discard its data.
                        row["Not Found"] = (
                            "Not found on anzousa.com"
                        )

                        row["Not Found Details"] = (
                            f"closest page ('{data.get('Product Description', '')}') "
                            "does not contain this SKU"
                        )

                        logging.error(
                            f"{inv}: opened page "
                            f"'{data.get('Product Description', '')}' "
                            f"does not contain the SKU. "
                            f"Data discarded: {url}"
                        )

                    else:

                        row.update(data)

                        row["MPN"] = (
                            row["_specs"].pop("MPN", "")
                            or mpn
                        )

                        if not row["_fitment"]:

                            logging.warning(
                                f"{inv}: no fitment table found"
                            )

                        print(
                            f"   -> {row['Product Description'][:70]} | "
                            f"{len(row['_specs'])} specs | "
                            f"fitment rows: {len(row['_fitment'])} | "
                            f"{status}"
                        )

            except Exception as e:

                logging.error(
                    f"{inv} failed: {e}"
                )

                row["Status"] = f"ERROR: {e}"

                row["Not Found"] = (
                    "ERROR while scraping - run this SKU again"
                )

                row["Not Found Details"] = str(e)[:300]

            rows.append(row)

            if len(rows) % SAVE_EVERY == 0:
                save_to_excel(rows)

    finally:

        if rows:

            save_to_excel(
                rows,
                format_file=True
            )

            missing = sum(
                1
                for r in rows
                if r.get("Not Found") == "Not found on anzousa.com"
            )

            errors = sum(
                1
                for r in rows
                if str(r.get("Not Found", "")).startswith("ERROR")
            )

            print(
                f"\nDone: {len(rows)} SKUs | "
                f"scraped {len(rows) - missing - errors} | "
                f"not found {missing} | errors {errors} "
                f"(see sheet '{NOT_FOUND_SHEET}')"
            )

        if driver:
            cleanup(driver)


if __name__ == "__main__":

    # TEST MODE: run selected MPNs.
    # Comment this line and uncomment main() to run the full list.

    # main(test_mpns=["211014", "211153", "211015"])

    main()
