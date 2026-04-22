# brownstone_step1_list.py
# deps: pip install "selenium>=4.23,<5" chromedriver-autoinstaller pandas openpyxl

from __future__ import annotations
import os, time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import chromedriver_autoinstaller as cda

# ============= CONFIG =============
PAGE_URL = "https://brownstonefurniture.com/products/nightstands/"
OUTPUT_FILE = "brownstone_step1_list.xlsx"

PAGELOAD_TIMEOUT = 90
IMPLICIT_WAIT = 4

PRODUCT_SEL = "div.uk-card-body h4 a"  # Product name and URL
IMAGE_SEL = "img.lazyloaded"  # Image URL
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

def get_text_or_empty(el) -> str:
    try:
        return el.text.strip()
    except Exception:
        return ""

def get_image_url(el) -> str:
    try:
        return el.get_attribute("src").strip()
    except Exception:
        return ""

def extract_product_data(driver) -> list:
    product_data = []
    driver.get(PAGE_URL)
    time.sleep(3)

    # Find product URLs, names, and image URLs
    product_elements = driver.find_elements(By.CSS_SELECTOR, PRODUCT_SEL)
    image_elements = driver.find_elements(By.CSS_SELECTOR, IMAGE_SEL)
    
    for product, image in zip(product_elements, image_elements):
        product_url = product.get_attribute("href")
        product_name = get_text_or_empty(product)
        image_url = get_image_url(image)
        
        # Collect product data
        product_data.append({
            "Product URL": product_url,
            "Image URL": image_url,
            "Product Name": product_name
        })
    
    return product_data

def save_to_excel(data: list) -> None:
    df = pd.DataFrame(data)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved data to {OUTPUT_FILE}")

def main():
    driver = make_driver()
    try:
        data = extract_product_data(driver)
        save_to_excel(data)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
