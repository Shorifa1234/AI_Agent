import os
import time
import re
import pandas as pd
from typing import Dict, Optional, Tuple

# =============== USER CONFIG ===============
CHROMEDRIVER_PATH = r"C:/chromedriver.exe"

# Input/Output in same folder as this script
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_XLSX  = os.path.join(BASE_DIR, "step1_vanguard.xlsx")
OUTPUT_XLSX = os.path.join(BASE_DIR, "vanguard_full.xlsx")

SAVE_EVERY_N = 20   # checkpoint save frequency
SKIP_IF_ALREADY_FILLED = True
HEADLESS = False     # False if you want to watch browser
SLEEP_BETWEEN = (0.8, 1.6)
# ===========================================

# ---------- Selenium setup ----------
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException

def make_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    if HEADLESS:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1280,1800")
    options.add_argument("--log-level=2")
    service = Service(CHROMEDRIVER_PATH)
    return webdriver.Chrome(service=service, options=options)

def sleep_polite():
    import random
    time.sleep(random.uniform(*SLEEP_BETWEEN))

def safe_get_text(el) -> str:
    try:
        return el.text.strip()
    except Exception:
        return ""

def click_to_expand_if_needed(driver, header_locator: Tuple[By, str], wait_secs: int = 8) -> bool:
    try:
        header = WebDriverWait(driver, wait_secs).until(EC.presence_of_element_located(header_locator))
    except TimeoutException:
        return False

    panel_name = header.get_attribute("panel-name") or ""
    body_xpath = f"//*[contains(@class,'CCP_Body') and @panel-name='{panel_name}']"

    def body_visible():
        try:
            body = driver.find_element(By.XPATH, body_xpath)
            style = body.get_attribute("style") or ""
            return "display: none" not in style.lower()
        except NoSuchElementException:
            return False

    if body_visible():
        return True

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", header)
        time.sleep(0.2)
        header.click()
        time.sleep(0.4)
    except (ElementClickInterceptedException, StaleElementReferenceException):
        try:
            driver.execute_script("arguments[0].click();", header)
            time.sleep(0.4)
        except Exception:
            pass

    return body_visible()

def extract_description(driver) -> Optional[str]:
    try:
        wrap = WebDriverWait(driver, 6).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#divRightSide #divFullHeader .Romance"))
        )
        return safe_get_text(wrap)
    except TimeoutException:
        return None

def get_dimensions_text(driver) -> str:
    try:
        dim_el = WebDriverWait(driver, 4).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(@id,'_divDimenions') and contains(@class,'DetailText')]")
            )
        )
        return safe_get_text(dim_el)
    except TimeoutException:
        pass
    try:
        body = driver.find_element(By.XPATH, "//*[contains(@class,'CCP_Body') and @panel-name='Dimensions']")
        return safe_get_text(body)
    except NoSuchElementException:
        return ""

def parse_dimensions(dim_text: str) -> Dict[str, Optional[str]]:
    result = {"Width": None, "Depth": None, "Height": None, "Diameter": None}
    if not dim_text:
        return result
    text = " ".join(dim_text.split())
    patterns = {
        "Width":   r"(?:\bW(?:idth)?\.?\s*[:=]?\s*)(\d+(?:\.\d+)?)\b",
        "Depth":   r"(?:\bD(?:epth)?\.?\s*[:=]?\s*)(\d+(?:\.\d+)?)\b",
        "Height":  r"(?:\bH(?:eight)?\.?\s*[:=]?\s*)(\d+(?:\.\d+)?)\b",
        "Diameter":r"(?:\bDia(?:meter)?\.?\s*[:=]?\s*)(\d+(?:\.\d+)?)\b",
    }
    for key, pat in patterns.items():
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            result[key] = m.group(1)
    if not result["Width"]:
        m = re.search(r"\b(\d+(?:\.\d+)?)\s*[x×]?\s*W\b", text, re.IGNORECASE)
        if m: result["Width"] = m.group(1)
    if not result["Depth"]:
        m = re.search(r"\b(\d+(?:\.\d+)?)\s*[x×]?\s*D\b", text, re.IGNORECASE)
        if m: result["Depth"] = m.group(1)
    if not result["Height"]:
        m = re.search(r"\b(\d+(?:\.\d+)?)\s*[x×]?\s*H\b", text, re.IGNORECASE)
        if m: result["Height"] = m.group(1)
    return result

def extract_weight(driver) -> Optional[str]:
    try:
        body = driver.find_element(By.XPATH, "//*[contains(@class,'CCP_Body') and @panel-name='Shipping']")
        text = safe_get_text(body)
    except NoSuchElementException:
        text = ""
    if not text:
        return None
    m = re.search(r"Weight\s*:\s*([0-9]+(?:\.\d+)?)\s*lb", text, re.IGNORECASE)
    return m.group(1) if m else None

def expand_section(driver, panel_name: str) -> bool:
    locator = (By.XPATH, f"//*[@class='SectionHeader CCP_Header' and @panel-name='{panel_name}']")
    return click_to_expand_if_needed(driver, locator, wait_secs=8)

def process_product(driver, url: str) -> Dict[str, Optional[str]]:
    out = {"Description": None,"Width": None,"Depth": None,"Height": None,"Diameter": None,"Weight": None}
    driver.get(url)
    try:
        WebDriverWait(driver, 10).until(
            EC.any_of(
                EC.presence_of_element_located((By.ID, "divRightSide")),
                EC.presence_of_element_located((By.CSS_SELECTOR, ".SectionHeader.CCP_Header"))
            )
        )
    except TimeoutException:
        pass
    out["Description"] = extract_description(driver)
    if expand_section(driver, "Dimensions"):
        dim_text = get_dimensions_text(driver)
        out.update(parse_dimensions(dim_text))
    if expand_section(driver, "Shipping"):
        out["Weight"] = extract_weight(driver)
    return out

def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    needed = ["Product URL","Image URL","Product Name","SKU","Description","Weight","Width","Depth","Diameter","Height"]
    for col in needed:
        if col not in df.columns:
            df[col] = None
    return df[needed]

def row_already_filled(row: pd.Series) -> bool:
    if not SKIP_IF_ALREADY_FILLED:
        return False
    return any([
        bool(row.get("Description")),
        bool(row.get("Weight")),
        bool(row.get("Width")),
        bool(row.get("Depth")),
        bool(row.get("Height")),
        bool(row.get("Diameter"))
    ])

def main():
    df = pd.read_excel(INPUT_XLSX)
    df = ensure_columns(df)
    driver = make_driver()
    processed = 0
    try:
        for idx, row in df.iterrows():
            url = str(row.get("Product URL") or "").strip()
            if not url: continue
            if row_already_filled(row): continue
            try:
                sleep_polite()
                data = process_product(driver, url)
                for k,v in data.items():
                    df.at[idx,k] = v
            except Exception as e:
                print(f"[WARN] Failed at row {idx} URL={url} :: {e}")
            processed += 1
            if processed % SAVE_EVERY_N == 0:
                tmp_path = OUTPUT_XLSX.replace(".xlsx","_checkpoint.xlsx")
                try:
                    df.to_excel(tmp_path, index=False)
                    print(f"[Checkpoint] Saved: {tmp_path}")
                except Exception as e:
                    print(f"[WARN] Checkpoint save failed: {e}")
    finally:
        driver.quit()
    df = ensure_columns(df)
    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"[DONE] Saved: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
