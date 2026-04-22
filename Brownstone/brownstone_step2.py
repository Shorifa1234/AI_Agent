# brownstone_step2_details.py
# deps: pip install "selenium>=4.23,<5" chromedriver-autoinstaller pandas openpyxl

from __future__ import annotations
import os, re, time
from typing import Dict, List, Optional
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
import chromedriver_autoinstaller as cda

# ============= CONFIG =============
STEP1_FILE = "brownstone_step1_list.xlsx"
OUTFILE    = "brownstone_step2_details.xlsx"

PAGELOAD_TIMEOUT = 90
IMPLICIT_WAIT    = 4
SAVE_EVERY_N     = 15  # Save progress every 15 products

INFO_BLOCK_SEL   = "//div[contains(@class, 'uk-width-expand') and contains(@class, 'uk-first-column')]"
DIMENSIONS_SPAN  = "span.dimensions"
DESC_BLOCK_SEL   = "div.product-info-fulldescription"
DESC_H4_SEL      = "h4.wp-block-heading"
# ==================================

def make_driver() -> webdriver.Chrome:
    cda.install()
    opts = Options()
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,1200")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    drv = webdriver.Chrome(options=opts)
    drv.set_page_load_timeout(PAGELOAD_TIMEOUT)
    drv.implicitly_wait(IMPLICIT_WAIT)
    return drv

def clean_text(s: Optional[str]) -> str:
    if not s: return ""
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def clean_html_text(s: str) -> str:
    s = re.sub(r"<br\s*/?>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def get_text_or_empty(el) -> str:
    try:
        return clean_text(el.text)
    except Exception:
        return ""

def get_labeled_value(container, label: str) -> str:
    """Extracts VALUE from <p><strong>LABEL:</strong><br>VALUE</p>"""
    try:
        el = container.find_element(By.XPATH, f".//p[strong[normalize-space()='{label}:']]")
    except Exception:
        return ""
    # Prefer innerHTML after </strong><br>
    try:
        html = el.get_attribute("innerHTML") or ""
        m = re.search(r"</strong>\s*<br\s*/?>\s*(.+)$", html, flags=re.I|re.S)
        if m:
            return clean_html_text(m.group(1))
    except Exception:
        pass
    txt = (el.text or "").strip()
    txt = re.sub(rf"^{re.escape(label)}:\s*", "", txt, flags=re.I).strip()
    return txt

def extract_dimensions_from_spans(block_el) -> Dict[str, str]:
    """
    Map span-based and compact dimensions into:
      Width, Depth, Height, Dia + Dimensions Raw
    Handles:
      29" wide / 20" deep / 26" height
      26"w x 18"d x 28"h
      Ø36" / 36" dia / diameter 36"
    """
    width = depth = height = dia = ""
    raw_parts: List[str] = []

    # --- 1) span.dimensions style ---
    try:
        spans = block_el.find_elements(By.CSS_SELECTOR, DIMENSIONS_SPAN)
        for sp in spans:
            t = get_text_or_empty(sp)
            if not t: 
                continue
            raw_parts.append(t)

            # 29" wide
            m_w = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*(?:w(?:ide)?\b)', t, re.I)
            m_d = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*(?:d(?:eep)?\b)', t, re.I)
            m_h = re.search(r'(\d+(?:\.\d+)?)\s*"?\s*(?:h(?:eight)?\b)', t, re.I)
            m_o = re.search(r'(?:ø|diameter|dia)\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*"?', t, re.I)

            if m_w and not width:  width  = m_w.group(1)  # Keep only numeric value (no inches symbol)
            if m_d and not depth:  depth  = m_d.group(1)  # Keep only numeric value (no inches symbol)
            if m_h and not height: height = m_h.group(1)  # Keep only numeric value (no inches symbol)
            if m_o and not dia:    dia    = m_o.group(1)  # Keep only numeric value (no inches symbol)
    except Exception:
        pass

    # --- 2) Fallback: compact patterns in the whole block text ---
    block_text = get_text_or_empty(block_el)
    if block_text:
        # W x D x H
        m_compact = re.search(
            r'(\d+(?:\.\d+)?)\s*"?\s*w[^a-zA-Z0-9]+(\d+(?:\.\d+)?)\s*"?\s*d[^a-zA-Z0-9]+(\d+(?:\.\d+)?)\s*"?\s*h',
            block_text, re.I)
        if m_compact:
            if not width:  width  = m_compact.group(1)
            if not depth:  depth  = m_compact.group(2)
            if not height: height = m_compact.group(3)
            raw_parts.append(m_compact.group(0))

        # Diameter-only patterns
        m_dia = re.search(r'(?:ø|diameter|dia)\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*"?', block_text, re.I)
        if m_dia and not dia:
            dia = m_dia.group(1)
            raw_parts.append(m_dia.group(0))

        # Also accept "Width: 29", "Depth: 20", "Height: 26", "Dia: 36"
        def kv(label):
            m = re.search(rf'{label}\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*"?', block_text, re.I)
            return (m.group(1)) if m else ""  # Removed the inches symbol

        if not width:  width  = kv("Width")
        if not depth:  depth  = kv("Depth")
        if not height: height = kv("Height")
        if not dia:    dia    = kv("Dia|Diameter|Ø")

    dims_raw = clean_text(" | ".join(dict.fromkeys(raw_parts)))  # unique, joined

    # Print values in the console
    print(f"Width: {width}, Depth: {depth}, Height: {height}, Dia: {dia}")
    
    return {"Dimensions Raw": dims_raw, "Width": width, "Depth": depth, "Height": height, "Dia": dia}

def find_weight_everywhere(driver) -> str:
    try:
        body_text = clean_text(driver.find_element(By.TAG_NAME, "body").text)
    except Exception:
        body_text = ""
    m = re.search(r'Weight\s*:\s*([0-9]+(?:\.\d+)?)\s*(?:lb|lbs|pounds)?', body_text, re.I)
    if m:
        weight = m.group(1)  # Extract numeric weight only (no unit)
        print(f"Weight: {weight}")  # Print to console
        return weight
    m2 = re.search(r'([0-9]+(?:\.\d+)?)\s*(?:lb|lbs)\b', body_text, re.I)
    if m2:
        weight = m2.group(1)
        print(f"Weight: {weight}")  # Print to console
        return weight
    return ""

def extract_description(driver) -> str:
    parts: List[str] = []
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, DESC_H4_SEL):
            txt = get_text_or_empty(el)
            if txt: parts.append(txt)
    except Exception:
        pass
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, DESC_BLOCK_SEL):
            txt = get_text_or_empty(el)
            if txt: parts.append(txt)
    except Exception:
        pass
    seen, uniq = set(), []
    for p in parts:
        if p not in seen:
            uniq.append(p); seen.add(p)
    return clean_text("\n\n".join(uniq))

def extract_product_name(driver) -> str:
    """
    Extracts the product name from <h4> tags, removing unwanted tags like <strong> or "New" badge.
    Handles cases where the product name may be formatted with extra HTML tags.
    """
    try:
        # Find <h4> tags, which often contain the product name
        product_name_element = driver.find_element(By.CSS_SELECTOR, 'h4.wp-block-heading')
        product_name = get_text_or_empty(product_name_element).strip()

        # Remove "New" badge or any additional formatting
        product_name = re.sub(r"\s*New\s*", "", product_name, flags=re.I)

        return clean_text(product_name)  # Clean extra spaces or special chars
    except Exception as e:
        print(f"[Error] Failed to extract product name: {e}")
        return ""

def extract_info_block(driver) -> Dict[str, str]:
    data = {
        "SKU": "", "Finish": "", "Made In": "", "Materials": "",
        "Dimensions Raw": "", "Width": "", "Depth": "", "Height": "", "Dia": ""
    }
    try:
        block = driver.find_element(By.XPATH, INFO_BLOCK_SEL)
    except NoSuchElementException:
        return data

    data["SKU"]       = get_labeled_value(block, "SKU")
    data["Finish"]    = get_labeled_value(block, "Finish")
    data["Made In"]   = get_labeled_value(block, "Made In")
    data["Materials"] = get_labeled_value(block, "Materials")

    dims = extract_dimensions_from_spans(block)
    data.update(dims)
    return data

def enrich_row(driver, row: Dict[str, str]) -> Dict[str, str]:
    url = row.get("Product URL", "").strip()
    if not url:
        return row

    for attempt in (1, 2):
        try:
            driver.get(url)
            break
        except Exception:
            time.sleep(1)
    time.sleep(1.2)

    # Extract product name correctly
    product_name = extract_product_name(driver)
    
    # Now enrich the rest of the details
    info   = extract_info_block(driver)
    weight = find_weight_everywhere(driver)  # Weight extraction with print to console
    descr  = extract_description(driver)

    # Update the row with the enriched data
    row.update(info)
    row["Product Name"] = product_name  # Add the extracted product name here
    row["Weight"] = weight
    row["Description"] = descr
    return row

def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols_order = [
        "Product URL", "Image URL", "Product Name",
        "SKU", "Finish", "Made In", "Materials",
        "Dimensions Raw", "Width", "Depth", "Height", "Dia",
        "Weight", "Description"
    ]
    for c in cols_order:
        if c not in df.columns:
            df[c] = ""
    return df[cols_order]

def main():
    if not os.path.exists(STEP1_FILE):
        raise FileNotFoundError(f"Step-1 file not found: {STEP1_FILE}")

    base = pd.read_excel(STEP1_FILE)
    if base.empty:
        print("Step-1 file is empty; nothing to enrich.")
        return

    base["Product URL"] = base["Product URL"].astype(str).fillna("").str.strip()
    base = base.dropna(subset=["Product URL"])
    base = base.drop_duplicates(subset=["Product URL"], keep="first").reset_index(drop=True)
    base = ensure_columns(base)

    drv = make_driver()
    try:
        enriched: List[Dict[str, str]] = []
        total = len(base)
        for i, row in base.iterrows():
            d = row.to_dict()
            product_name = str(d.get('Product Name', '')).strip()
            print(f"[{i+1}/{total}] {product_name[:70]}")
            d = enrich_row(drv, d)
            enriched.append(d)

            if (i+1) % SAVE_EVERY_N == 0:
                pd.DataFrame(enriched).pipe(ensure_columns).to_excel(OUTFILE, index=False)
                print(f"[save] {i+1} rows -> {OUTFILE}")

        pd.DataFrame(enriched).pipe(ensure_columns).to_excel(OUTFILE, index=False)
        print(f"\n[done] Enriched {len(enriched)} rows -> {OUTFILE}")

    finally:
        try:
            drv.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
