# detailspage.py
# Step 2 – Paul Ferrante product details scraper
# Python 3.9+ | selenium 4.x | pandas | openpyxl
# VPN NEED

import os
import re
import time
import random
import shutil
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchDriverException, WebDriverException

# ========= I/O =========
IN_XLSX  = "paulferrante_step1_full.xlsx"
OUT_XLSX = "paulferrante_step2.xlsx"

# ========= Timeouts / pacing =========
PAGE_TIMEOUT = 30
TABLE_TIMEOUT = 8

# ========= Driver config (UPDATED) =========
CHROME_DEBUG_ADDR = "127.0.0.1:9222"  # your launched Chrome's port

# Chrome install path (adjust if different)
CHROME_BINARY = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Location where you placed chromedriver.exe
FALLBACK_DRIVER_DIR = r"C:\workspace"

# Common Chrome binary fallback paths
COMMON_CHROME_BINARIES = [
    CHROME_BINARY,
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


# ========= Helpers =========
def _resolve_chrome_binary() -> Optional[str]:
    if CHROME_BINARY and os.path.exists(CHROME_BINARY):
        return CHROME_BINARY
    for p in COMMON_CHROME_BINARIES:
        if os.path.exists(p):
            return p
    which = shutil.which("chrome") or shutil.which("chrome.exe") or shutil.which("google-chrome")
    return which

def _find_chromedriver_in(base_dir: str) -> Optional[str]:
    if not base_dir or not os.path.isdir(base_dir):
        return None
    candidates = [
        os.path.join(base_dir, "chromedriver.exe"),
        os.path.join(base_dir, "chromedriver", "chromedriver.exe"),
        os.path.join(base_dir, "chromedriver-win64", "chromedriver.exe"),
        os.path.join(base_dir, "chrome-win64", "chromedriver.exe"),
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    for root, _dirs, files in os.walk(base_dir):
        for f in files:
            if f.lower() == "chromedriver.exe":
                return os.path.join(root, f)
    return None

def _attach_with_selenium_manager(chrome_bin: str):
    opts = webdriver.ChromeOptions()
    opts.binary_location = chrome_bin
    opts.add_experimental_option("debuggerAddress", CHROME_DEBUG_ADDR)
    return webdriver.Chrome(options=opts)

def _attach_with_fallback_service(chrome_bin: str, driver_exe: str):
    opts = webdriver.ChromeOptions()
    opts.binary_location = chrome_bin
    opts.add_experimental_option("debuggerAddress", CHROME_DEBUG_ADDR)
    service = Service(driver_exe)
    return webdriver.Chrome(service=service, options=opts)

def make_driver():
    chrome_bin = _resolve_chrome_binary()
    if not chrome_bin:
        raise FileNotFoundError(
            "Chrome binary not found.\n"
            "Set CHROME_BINARY to the path of chrome.exe or install in the default location:\n"
            r"  C:\Program Files\Google\Chrome\Application\chrome.exe"
        )

    # First try: Selenium Manager (auto)
    try:
        drv = _attach_with_selenium_manager(chrome_bin)
        print("[OK] Attached via Selenium Manager (auto driver).")
        print(f"[OK] Chrome: {chrome_bin}")
        return drv
    except (NoSuchDriverException, WebDriverException) as e:
        print(f"[WARN] Selenium Manager attach failed: {e.__class__.__name__}: {e}")

    # Fallback: local chromedriver
    driver_exe = _find_chromedriver_in(FALLBACK_DRIVER_DIR)
    if driver_exe and os.path.exists(driver_exe):
        try:
            drv = _attach_with_fallback_service(chrome_bin, driver_exe)
            print(f"[OK] Fallback attach using local chromedriver: {driver_exe}")
            print(f"[OK] Chrome: {chrome_bin}")
            return drv
        except Exception as e2:
            raise RuntimeError(
                "Fallback attach failed.\n"
                f"  • Tried local chromedriver: {driver_exe}\n"
                f"  • Chrome binary: {chrome_bin}\n"
                "Tips:\n"
                "  1) Make sure Chrome is running with --remote-debugging-port=9222\n"
                "  2) chromedriver version must match Chrome's version"
            ) from e2

    raise RuntimeError(
        "Auto attach via Selenium Manager failed, and fallback chromedriver not found.\n"
        "Solution:\n"
        "  1) Ensure Chrome is running with --remote-debugging-port=9222\n"
        "  2) If internet is stable, Selenium Manager will auto-download chromedriver\n"
        "  3) Or, put chromedriver.exe in C:\\workspace\n"
        "  4) Set CHROME_BINARY to the path of chrome.exe explicitly"
    )


# ========= Page utils =========
def page_ready(driver, timeout=PAGE_TIMEOUT):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
    )

def is_wordfence_block(driver) -> bool:
    html = driver.page_source.lower()
    return ("wordfence" in html and "access to this site has been limited" in html) or \
           ("http response code 503" in html and "blocked" in html)

def click_additional_info_if_present(driver):
    try:
        possibles = [
            ("css", "a[href*='additional']"),
            ("css", "button[aria-controls*='additional']"),
            ("xpath", "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'additional information')]"),
            ("xpath", "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'additional information')]"),
            ("xpath", "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'product details')]"),
            ("xpath", "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'product details')]"),
        ]
        for kind, sel in possibles:
            try:
                if kind == "css":
                    el = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                else:
                    el = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, sel)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                time.sleep(0.15)
                el.click()
                time.sleep(0.4)
                break
            except:
                continue
    except:
        pass

def clean_label(label_text: str) -> str:
    if label_text is None:
        return ""
    t = label_text.strip()
    t = re.sub(r":\s*$", "", t)
    t = re.sub(r"\s+", " ", t)
    return t

def get_attributes(driver) -> Dict[str, str]:
    attrs: Dict[str, str] = {}

    tables = driver.find_elements(By.CSS_SELECTOR, "table.woocommerce-product-attributes.shop_attributes")
    if not tables:
        click_additional_info_if_present(driver)
        tables = driver.find_elements(By.CSS_SELECTOR, "table.woocommerce-product-attributes.shop_attributes")

    if not tables:
        return attrs

    table = tables[0]
    rows = table.find_elements(By.CSS_SELECTOR, "tr.woocommerce-product-attributes-item")
    for row in rows:
        try:
            label_el = row.find_element(By.CSS_SELECTOR, ".woocommerce-product-attributes-item__label")
            value_el = row.find_element(By.CSS_SELECTOR, ".woocommerce-product-attributes-item__value")
            label = clean_label(label_el.text)
            value = value_el.text.strip()
            if not label:
                continue
            base = label
            i = 2
            while label in attrs:
                label = f"{base} ({i})"
                i += 1
            attrs[label] = value
        except:
            continue

    return attrs


# ========= Crawl =========
def process_links(driver, df_in: pd.DataFrame) -> List[dict]:
    results: List[dict] = []

    try:
        driver.switch_to.new_window('tab')
    except Exception:
        pass

    for idx, row in df_in.iterrows():
        link = (row.get("Product Link") or "").strip()
        name = (row.get("Product Name") or "").strip()
        img  = (row.get("Product Image") or "").strip()

        if not link:
            print(f"[{idx}] Empty Product Link. Skipping.")
            continue

        try:
            driver.get(link)
            page_ready(driver)

            if is_wordfence_block(driver):
                print(f"[{idx}] Wordfence block detected at: {link} — skipping.")
                continue

            time.sleep(1.0 + random.uniform(0.2, 0.8))

            try:
                WebDriverWait(driver, TABLE_TIMEOUT).until(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, "table.woocommerce-product-attributes.shop_attributes, tr.woocommerce-product-attributes-item")
                    )
                )
            except:
                pass

            attrs = get_attributes(driver)

            record = {
                "Product Link": link,
                "Product Name": name,
                "Product Image": img,
            }
            record.update(attrs)
            results.append(record)
            print(f"[{idx}] OK: {name or link} ({len(attrs)} attrs)")

            time.sleep(0.8 + random.uniform(0.2, 0.8))

        except Exception as e:
            print(f"[{idx}] Error on {link}: {e}")
            results.append({
                "Product Link": link,
                "Product Name": name,
                "Product Image": img,
            })
            time.sleep(1.5)

    return results


# ========= Main =========
def main():
    in_path = Path(IN_XLSX).resolve()
    if not in_path.exists():
        raise FileNotFoundError(f"Cannot find input Excel: {in_path}")

    df_in = pd.read_excel(in_path)

    for col in ["Product Link", "Product Name", "Product Image"]:
        if col not in df_in.columns:
            df_in[col] = ""

    driver = make_driver()
    try:
        rows = process_links(driver, df_in)

        df_out = pd.DataFrame(rows)
        step1_cols = ["Product Link", "Product Name", "Product Image"]
        attr_cols = [c for c in df_out.columns if c not in step1_cols]
        ordered_cols = step1_cols + sorted(attr_cols, key=lambda x: x.lower())

        df_out = df_out.reindex(columns=ordered_cols)

        out_path = Path(OUT_XLSX).resolve()
        df_out.to_excel(out_path, index=False)
        print(f"\nSaved {len(df_out)} rows to {out_path}")
        print(f"Columns: {', '.join(ordered_cols)}")

    finally:
        pass


if __name__ == "__main__":
    main()
