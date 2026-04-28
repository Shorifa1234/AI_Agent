"""
Gabby Step 1 — List Page Scraper
Website: https://gabby.com
Method: Shopify JSON API (no Selenium needed)
Output: Gabby_step1.xlsx — one sheet per category
"""

import requests
import openpyxl
import time
import os

BASE_URL = "https://gabby.com"
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "Gabby_step1.xlsx")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

# collection slug → sheet name
CATEGORIES = {
    "indoor-sofas-settees-loveseats":    "Sofas & Loveseats",
    "indoor-sectional-seating":          "Sectionals",
    "indoor-lounge-swivel-chairs":       "Lounge & Swivel Chairs",
    "indoor-accent-chairs":              "Accent Chairs",
    "indoor-benches-banquettes":         "Benches & Banquettes",
    "indoor-ottomans-stools":            "Ottomans & Stools",
    "indoor-coffee-tables":              "Coffee Tables",
    "indoor-side-end-tables":            "Side & End Tables",
    "indoor-console-tables":             "Console Tables",
    "indoor-accent-tables":              "Accent Tables",
    "indoor-cabinets":                   "Cabinets",
    "indoor-bookcases":                  "Bookcases",
    "indoor-credenzas":                  "Credenzas",
    "indoor-sideboards-buffets":         "Sideboards & Buffets",
    "indoor-dining-tables":              "Dining Tables",
    "indoor-dining-chairs":              "Dining Chairs",
    "indoor-bar-counter-height-stools":  "Bar & Counter Stools",
    "beds":                              "Beds",
    "headboards":                        "Headboards",
    "dressers":                          "Dressers",
    "nightstands":                       "Nightstands",
    "desks":                             "Desks",
    "mirrors":                           "Mirrors",
    "ceiling-lights":                    "Ceiling Lights",
    "lamps":                             "Lamps",
    "wall-lights":                       "Wall Lights",
}


def fetch_collection(slug):
    """Fetch all products from a Shopify collection using JSON API with pagination."""
    products = []
    page = 1
    while True:
        url = f"{BASE_URL}/collections/{slug}/products.json?limit=250&page={page}"
        try:
            resp = requests.get(url, headers=HEADERS, timeout=20)
            resp.raise_for_status()
            batch = resp.json().get("products", [])
            if not batch:
                break
            products.extend(batch)
            print(f"    Page {page}: {len(batch)} products fetched")
            page += 1
            time.sleep(0.5)
        except Exception as e:
            print(f"    Error on page {page}: {e}")
            break
    return products


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for slug, category_name in CATEGORIES.items():
        print(f"\n[{category_name}]  ->  /collections/{slug}")
        products = fetch_collection(slug)

        if not products:
            print(f"  No products found, skipping.")
            continue

        ws = wb.create_sheet(title=category_name[:31])
        ws.append([
            "Category", "Product Name", "Manufacturer SKU",
            "Product URL", "Image URL", "Product Type", "Collection Slug"
        ])

        for p in products:
            name      = p.get("title", "")
            handle    = p.get("handle", "")
            url       = f"{BASE_URL}/products/{handle}"
            mfr_sku   = p["variants"][0].get("sku", "") if p.get("variants") else ""
            img       = p["images"][0].get("src", "") if p.get("images") else ""
            if img.startswith("//"):
                img = "https:" + img
            ptype     = p.get("product_type", "")
            ws.append([category_name, name, mfr_sku, url, img, ptype, slug])

        print(f"  Saved {len(products)} products -> sheet '{category_name}'")

    wb.save(OUTPUT_FILE)
    print(f"\nDone! Output: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
