import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# ---------- SETUP ----------
chrome_driver_path = "C:/chromedriver.exe"
options = Options()
options.add_argument("--start-maximized")

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)

# ---------- SCRAPING ----------
base_url = "https://www.vanguardfurniture.com/styles?Room=LR&ProdType=004"
all_data = []

for page in range(0, 12):  # total 12 pages
    if page == 0:
        url = base_url
    else:
        url = f"{base_url}&PageIndex={page}"

    driver.get(url)
    time.sleep(3)  # allow page to load
    
    # scroll for lazyload
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)

    # find product containers
    products = driver.find_elements(By.CSS_SELECTOR, ".grid_3.SearchResults_Container")

    for product in products:
        try:
            # main <a>
            a_tag = product.find_element(By.TAG_NAME, "a")
            relative_url = a_tag.get_attribute("href")
            if relative_url.startswith("/"):
                product_url = "https://www.vanguardfurniture.com" + relative_url
            else:
                product_url = relative_url

            # image
            img_tag = a_tag.find_element(By.TAG_NAME, "img")
            image_url = img_tag.get_attribute("src")
            sku = img_tag.get_attribute("alt")

            # text lines inside <a>
            text_lines = a_tag.text.strip().split("\n")
            product_name = ""
            if len(text_lines) >= 2:
                sku_text = text_lines[0].strip()
                product_name = text_lines[1].strip()
            else:
                sku_text = sku

            all_data.append({
                "Product URL": product_url,
                "Image URL": image_url,
                "Product Name": product_name,
                "SKU": sku_text
            })
        except Exception as e:
            print("Error:", e)

# ---------- SAVE ----------
df = pd.DataFrame(all_data)
df.to_excel("step1_vanguard.xlsx", index=False)

driver.quit()
print("Step 1 completed. Data saved to step1_vanguard.xlsx")
