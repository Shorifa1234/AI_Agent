# listpage.py
# Step 1 – Paul Ferrante (Lighting) list scraper
# Python 3.9+ | selenium 4.x | pandas | openpyxl
#
# IMPORTANT: Start Chrome manually with remote debugging BEFORE running:
# "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:/workspace/chrome_debug"

import time
import random
import re
import socket
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ===================== CONFIG =====================
LIST_URL = "https://paulferrante.com/product-category/lighting/ceiling-fixtures/"
OUT_XLSX = "paulferrante_step1_full.xlsx"
TARGET_COUNT = None  # None = collect all available
INITIAL_DWELL_SEC = 2.0

# Choose driver setup:
USE_SELENIUM_MANAGER = True  # True = no path needed; False = use explicit path below
CHROMEDRIVER_PATH = r"C:\chromedriver.exe"  # used only if USE_SELENIUM_MANAGER=False

REMOTE_DEBUG_ADDR = "127.0.0.1:9222"  # must match the Chrome you started
REMOTE_DEBUG_HOST = "127.0.0.1"
REMOTE_DEBUG_PORT = 9222
# ==================================================


def _port_open(host="127.0.0.1", port=9222, timeout=1.0) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def make_driver():
    """
    Attach to the already-open Chrome via remote debugging.
    Ensure you've started Chrome with --remote-debugging-port=9222.
    """
    if not _port_open(REMOTE_DEBUG_HOST, REMOTE_DEBUG_PORT, timeout=1.0):
        raise RuntimeError(
            "Chrome is not reachable on 127.0.0.1:9222.\n"
            "Start it first:\n"
            r'  "C:\Program Files\Google\Chrome\Application\chrome.exe" '
            r'--remote-debugging-port=9222 --user-data-dir="C:/chrome_debug"'
        )

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", REMOTE_DEBUG_ADDR)

    if USE_SELENIUM_MANAGER:
        # ✅ No Service(...) path — let Selenium Manager fetch the right driver
        driver = webdriver.Chrome(options=options)
    else:
        # ✅ Explicit path (make sure path exists & matches Chrome major version)
        service = Service(CHROMEDRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)

    return driver


def page_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
    )


def is_wordfence_block(driver) -> bool:
    html = driver.page_source.lower()
    return ("wordfence" in html and "access to this site has been limited" in html) or \
           ("http response code 503" in html and "blocked" in html)


def gentle_scroll(driver, max_seconds=90):
    """
    Human-like scrolling with jitter to encourage lazy-loading.
    Stops when page height stops growing for a few iterations or time limit reached.
    """
    start = time.time()
    last_doc_h = driver.execute_script("return document.body.scrollHeight")
    stagnant = 0

    while True:
        # Scroll down by a chunk
        driver.execute_script("window.scrollBy(0, Math.min(window.innerHeight*0.75, 700));")
        time.sleep(0.6 + random.uniform(0.2, 0.6))

        # Occasional tiny up/down jiggle
        if random.random() < 0.15:
            driver.execute_script("window.scrollBy(0, -Math.min(window.innerHeight*0.2, 200));")
            time.sleep(0.25 + random.uniform(0.1, 0.3))
            driver.execute_script("window.scrollBy(0, Math.min(window.innerHeight*0.9, 900));")
            time.sleep(0.45 + random.uniform(0.1, 0.3))

        # Try clicking a "Load more" button if present (Elementor variants)
        try:
            load_more = driver.find_element(By.CSS_SELECTOR, "a.elementor-button[href*='load'], button.elementor-button")
            if load_more.is_displayed() and load_more.is_enabled():
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", load_more)
                time.sleep(0.4)
                load_more.click()
                time.sleep(1.2 + random.uniform(0.3, 0.7))
        except Exception:
            pass

        # Check document growth
        doc_h = driver.execute_script("return document.body.scrollHeight")
        if doc_h <= last_doc_h:
            stagnant += 1
        else:
            stagnant = 0
            last_doc_h = doc_h

        # Near bottom? Nudge to absolute bottom
        viewport_bottom = driver.execute_script("return window.pageYOffset + window.innerHeight")
        if viewport_bottom + 50 >= doc_h:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1.2 + random.uniform(0.2, 0.6))
            doc_h2 = driver.execute_script("return document.body.scrollHeight")
            if doc_h2 <= last_doc_h:
                stagnant += 1
            else:
                stagnant = 0
                last_doc_h = doc_h2

        if stagnant >= 6:
            break
        if time.time() - start > max_seconds:
            break


def get_cards(driver):
    # Primary selector for their product cards; add a couple of fallbacks for resilience.
    cards = driver.find_elements(By.CSS_SELECTOR, "div.elementor.product-type-simple")
    if not cards:
        cards = driver.find_elements(By.CSS_SELECTOR, "li.product")  # WooCommerce fallback
    if not cards:
        cards = driver.find_elements(By.CSS_SELECTOR, "[class*='product']")  # last-resort broad match
    return cards


def extract_card(driver, card):
    """
    Extract fields from a single product card:
    - Product Link: href of a.elementor-element (fallback to first anchor if needed)
    - Product Name: text of p.elementor-heading-title (fallbacks added)
    - Product Image: src of img.attachment-large (handles lazyload)
    """
    # Product Link
    product_link = ""
    try:
        link_el = card.find_element(By.CSS_SELECTOR, "a.elementor-element")
        product_link = (link_el.get_attribute("href") or "").strip()
    except Exception:
        try:
            link_el = card.find_elements(By.CSS_SELECTOR, "a[href]")[0]
            product_link = (link_el.get_attribute("href") or "").strip()
        except Exception:
            pass

    # Product Name
    product_name = ""
    # common Elementor/Woo headings
    name_selectors = [
        "p.elementor-heading-title",
        "h2.woocommerce-loop-product__title",
        "h2[class*='product']",
        "h3[class*='product']",
        ".woocommerce-loop-product__title",
    ]
    for sel in name_selectors:
        try:
            el = card.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt:
                product_name = txt
                break
        except Exception:
            continue

    # Product Image (lazy-load aware)
    product_image = ""
    try:
        img_el = None
        for sel in ["img.attachment-large", "img.wp-post-image", "img"]:
            try:
                img_el = card.find_element(By.CSS_SELECTOR, sel)
                break
            except Exception:
                continue

        if img_el:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", img_el)
            time.sleep(0.2 + random.uniform(0.05, 0.25))
            for _ in range(14):
                src = (img_el.get_attribute("src") or "").strip()
                if not src:
                    src = (
                        img_el.get_attribute("data-src")
                        or img_el.get_attribute("data-lazy-src")
                        or img_el.get_attribute("data-original")
                        or ""
                    ).strip()
                if src:
                    try:
                        loaded = driver.execute_script("return arguments[0].naturalWidth > 1", img_el)
                    except Exception:
                        loaded = True
                    if loaded:
                        product_image = src
                        break
                time.sleep(0.3 + random.uniform(0.05, 0.25))
    except Exception:
        pass

    return {
        "Product Link": product_link,
        "Product Name": product_name,
        "Product Image": product_image,
    }


def clean_product_name(name: str) -> str:
    """
    Remove the leading catalog number from names like:
    '1182 Copa Sconce' or '1182 - Copa Sconce' or '1182—Copa Sconce'
    """
    if not name:
        return name
    # strip digits + optional spaces + optional dash/en dash/em dash
    return re.sub(r'^\s*\d+\s*[-–—]*\s*', '', name).strip()


def crawl_step1():
    driver = make_driver()
    rows, seen = [], set()

    try:
        # Open the target in a new tab inside the attached browser
        driver.switch_to.new_window('tab')
        driver.get(LIST_URL)
        page_ready(driver)

        if is_wordfence_block(driver):
            print("Blocked by Wordfence on initial load. Try later or request allow-listing.")
            return

        time.sleep(INITIAL_DWELL_SEC + random.uniform(0.3, 0.8))

        # Keep trying to reveal all products: do a few scroll cycles until
        # the number of cards stops increasing.
        prev_count = -1
        for _ in range(6):
            gentle_scroll(driver, max_seconds=90)
            # Bottom nudges to catch late content
            for _ in range(2):
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1.2 + random.uniform(0.2, 0.6))
            cards_now = len(get_cards(driver))
            if cards_now <= prev_count:
                break
            prev_count = cards_now

        if is_wordfence_block(driver):
            print("Blocked by Wordfence during scrolling. Backing off.")
            return

        # Briefly bring each card into view once (helps images finish loading)
        for card in get_cards(driver):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
                time.sleep(0.08 + random.uniform(0.04, 0.12))
            except Exception:
                pass

        # Extract all unique items (no cap if TARGET_COUNT is None)
        for card in get_cards(driver):
            data = extract_card(driver, card)

            # clean product name: remove leading digits/symbols
            data["Product Name"] = clean_product_name(data.get("Product Name", ""))

            link = data.get("Product Link", "")
            if link and link not in seen:
                seen.add(link)
                rows.append(data)
                if TARGET_COUNT and len(rows) >= TARGET_COUNT:
                    break

        # Save Excel for Step-2
        df = pd.DataFrame(rows)
        cols = ["Product Link", "Product Image", "Product Name"]
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]

        out = Path(OUT_XLSX).resolve()
        df.to_excel(out, index=False)
        print(f"Saved {len(df)} rows to {out}")

    finally:
        # Keep attached Chrome open; do not quit() to avoid closing your session.
        pass


if __name__ == "__main__":
    crawl_step1()
