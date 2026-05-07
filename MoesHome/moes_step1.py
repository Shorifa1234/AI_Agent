"""
Moe's Home - Step 1
Fetches product listings per category via GraphQL API.
Output: moes_step1.xlsx
"""

import sys
import re
import time
import requests
import openpyxl

GQL_URL = "https://mcprod.moeshome.com/graphql"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/json",
    "Store": "default",
}
PAGE_SIZE = 48
DEMO = "--demo" in sys.argv
DEMO_LIMIT = 5

CATEGORIES = [
    {"name": "Coffee & Cocktail Tables", "code": "CO", "urls": [
        ("https://www.moeshome.com/furniture/living-room/coffee-tables", "MTI1Ng=="),
    ]},
    {"name": "Dining Tables", "code": "DT", "urls": [
        ("https://www.moeshome.com/furniture/dining-room/dining-tables", "MTI4MA=="),
        ("https://www.moeshome.com/furniture/dining-room/bar-counter-tables", "MTI5NQ=="),
    ]},
    {"name": "Consoles", "code": "CN", "urls": [
        ("https://www.moeshome.com/media-tv-consoles", "MTQzOQ=="),
        ("https://www.moeshome.com/furniture/living-room/console-tables", "MTI2NQ=="),
    ]},
    {"name": "Desks", "code": "DK", "urls": [
        ("https://www.moeshome.com/furniture/office/office-desks", "MTMzNw=="),
    ]},
    {"name": "Bookcases", "code": "BK", "urls": [
        ("https://www.moeshome.com/bookcases-shelves-cabinets", "MTQ3NQ=="),
    ]},
    {"name": "Cabinets", "code": "CA", "urls": [
        ("https://www.moeshome.com/sideboards-storage", "MTQ0OA=="),
    ]},
    {"name": "Accent Tables", "code": "AC", "urls": [
        ("https://www.moeshome.com/furniture/living-room/accent-side-tables", "MTI1OQ=="),
    ]},
    {"name": "Bar Carts", "code": "BA", "urls": [
        ("https://www.moeshome.com/furniture/dining-room/bar-carts-wine-cabinets", "MTQ5Mw=="),
    ]},
    {"name": "Dining Chairs", "code": "DI", "urls": [
        ("https://www.moeshome.com/furniture/dining-room/dining-chairs", "MTI4Mw=="),
    ]},
    {"name": "Bar Stools", "code": "BS", "urls": [
        ("https://www.moeshome.com/furniture/dining-room/bar-counter-stools", "MTI5Mg=="),
    ]},
    {"name": "Sofas & Loveseats", "code": "SO", "urls": [
        ("https://www.moeshome.com/furniture/living-room/sofas",            "MTIzNQ=="),
        ("https://www.moeshome.com/chaises",                                "Mjk5MA=="),
        ("https://www.moeshome.com/furniture/living-room/modular-sofas",    "MTI0MQ=="),
        ("https://www.moeshome.com/furniture/living-room/modular-components","MTk0NA=="),
    ]},
    {"name": "Sectionals", "code": "SE", "urls": [
        ("https://www.moeshome.com/furniture/living-room/sectionals", "MTIzOA=="),
    ]},
    {"name": "Lounge Chairs", "code": "LC", "urls": [
        ("https://www.moeshome.com/accent-lounge-chairs", "MTQzNg=="),
    ]},
    {"name": "Ottomans & Benches", "code": "OB", "urls": [
        ("https://www.moeshome.com/furniture/living-room/footstools-ottomans", "MTI1MA=="),
        ("https://www.moeshome.com/benches-stools",                            "MTIzMg=="),
    ]},
    {"name": "Desk Chairs", "code": "DS", "urls": [
        ("https://www.moeshome.com/furniture/office/office-chairs", "MTM0MA=="),
    ]},
    {"name": "Pendants", "code": "PE", "urls": [
        ("https://www.moeshome.com/decor/lighting/pendant-lamps", "MTQxOA=="),
    ]},
    {"name": "Floor Lamps", "code": "FL", "urls": [
        ("https://www.moeshome.com/decor/lighting/floor-lamps", "MTQxMg=="),
    ]},
    {"name": "Mirrors", "code": "MI", "urls": [
        ("https://www.moeshome.com/decor/wall-decor/mirrors", "MTM4Mg=="),
    ]},
    {"name": "Wall Decor", "code": "WA", "urls": [
        ("https://www.moeshome.com/decor/wall-decor/wall-sculptures", "MTU4MA=="),
        ("https://www.moeshome.com/decor/wall-decor/wall-art",        "MTQzMA=="),
    ]},
    {"name": "Outdoor Seating", "code": "OS", "urls": [
        ("https://www.moeshome.com/furniture/outdoor/outdoor-sofas-sectionals",          "MTM1NQ=="),
        ("https://www.moeshome.com/furniture/outdoor/outdoor-lounge-chairs-chaises",     "MTQ5Ng=="),
        ("https://www.moeshome.com/furniture/outdoor/outdoor-dining-chairs",             "MTM2Nw=="),
        ("https://www.moeshome.com/furniture/outdoor/outdoor-benches-stools",            "MTQ2OQ=="),
        ("https://www.moeshome.com/furniture/collaboration",                             "MzM1OA=="),
    ]},
    {"name": "Outdoor Tables", "code": "OT", "urls": [
        ("https://www.moeshome.com/furniture/outdoor/outdoor-coffee-side-tables", "MTM1OA=="),
        ("https://www.moeshome.com/furniture/outdoor/outdoor-dining-tables",      "MTQ2Ng=="),
    ]},
]

QUERY = """
query GetProducts($catUid: String!, $pageSize: Int!, $currentPage: Int!) {
  products(
    filter: {category_uid: {eq: $catUid}}
    pageSize: $pageSize
    currentPage: $currentPage
    sort: {name: ASC}
  ) {
    total_count
    page_info { current_page total_pages }
    items {
      sku
      name
      url_key
      description { html }
      image { url }
      moes_b2b_shipping_method
      moes_shipping_box_width
      moes_shipping_box_height
      moes_shipping_box_depth
      moes_shipping_box_weight
      ... on ConfigurableProduct {
        variants {
          attributes { label value_index }
          product {
            sku
            name
            image { url }
          }
        }
      }
    }
  }
}
"""

SHIPPING_METHOD_QUERY = """
query {
  customAttributeMetadata(attributes: [{entity_type: "4", attribute_code: "moes_b2b_shipping_method"}]) {
    items {
      attribute_code
      attribute_options { label value }
    }
  }
}
"""


def gql(query, variables=None):
    for attempt in range(3):
        try:
            payload = {"query": query}
            if variables:
                payload["variables"] = variables
            resp = requests.post(GQL_URL, json=payload, headers=HEADERS, timeout=30)
            data = resp.json()
            if "errors" in data:
                print(f"  GQL error: {data['errors'][0]['message']}")
                return None
            return data["data"]
        except Exception as e:
            print(f"  Request error (attempt {attempt+1}): {e}")
            time.sleep(2)
    return None


def fetch_shipping_method_map():
    data = gql(SHIPPING_METHOD_QUERY)
    if not data:
        return {}
    options = data["customAttributeMetadata"]["items"][0]["attribute_options"]
    return {opt["value"]: opt["label"] for opt in options}


def first_box_val(pipe_str):
    if not pipe_str:
        return ""
    val = pipe_str.split("|")[0]
    try:
        f = float(val)
        return "" if f == 0.0 else str(f)
    except Exception:
        return ""


def get_products_for_category(cat_uid, demo=False):
    products = []
    seen_skus = set()
    page = 1
    while True:
        data = gql(QUERY, {"catUid": cat_uid, "pageSize": PAGE_SIZE, "currentPage": page})
        if not data:
            break
        result = data["products"]
        items = result["items"]
        total_pages = result["page_info"]["total_pages"]
        for item in items:
            sku = item["sku"]
            if sku in seen_skus:
                continue
            seen_skus.add(sku)
            products.append(item)
            if demo and len(products) >= DEMO_LIMIT:
                return products
        print(f"    Page {page}/{total_pages} - {len(items)} products")
        if page >= total_pages:
            break
        page += 1
        time.sleep(0.3)
    return products


def main():
    mode = "DEMO" if DEMO else "FULL"
    print(f"=== Moe's Home Step 1 [{mode}] ===")

    print("Fetching shipping method options...")
    ship_map = fetch_shipping_method_map()
    print(f"  Shipping methods: {ship_map}")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for cat in CATEGORIES:
        cat_name = cat["name"]
        cat_code = cat["code"]
        all_products = []
        seen = set()
        primary_url = cat["urls"][0][0]

        print(f"\nCategory: {cat_name}")
        for url, uid in cat["urls"]:
            print(f"  Fetching uid={uid} ({url})")
            products = get_products_for_category(uid, demo=DEMO)
            for p in products:
                if p["sku"] not in seen:
                    seen.add(p["sku"])
                    all_products.append(p)
                if DEMO and len(all_products) >= DEMO_LIMIT:
                    break
            if DEMO and len(all_products) >= DEMO_LIMIT:
                break

        print(f"  Total: {len(all_products)} products")

        ws = wb.create_sheet(title=cat_name[:31])
        ws.append(["Category", "Code", "Primary URL", "SKU", "Name",
                   "Product URL", "Image URL", "Description",
                   "Shipping Method", "Box W", "Box H", "Box D", "Box Wt"])

        for item in all_products:
            prod_url = f"https://www.moeshome.com/{item['url_key']}"
            desc_html = (item.get("description") or {}).get("html") or ""
            desc = re.sub(r"<[^>]+>", " ", desc_html).strip()
            desc = re.sub(r"\s+", " ", desc)
            parent_img = (item.get("image") or {}).get("url") or ""
            parent_img = parent_img.split("?")[0]
            ship_id = str(item.get("moes_b2b_shipping_method") or "")
            ship_label = ship_map.get(ship_id, "")

            box_w = first_box_val(item.get("moes_shipping_box_width"))
            box_h = first_box_val(item.get("moes_shipping_box_height"))
            box_d = first_box_val(item.get("moes_shipping_box_depth"))
            box_wt = first_box_val(item.get("moes_shipping_box_weight"))

            variants = item.get("variants") or []
            if variants:
                for v in variants:
                    vp = v.get("product") or {}
                    v_sku = vp.get("sku") or item["sku"]
                    v_name = vp.get("name") or item["name"]
                    v_img = (vp.get("image") or {}).get("url") or parent_img
                    v_img = v_img.split("?")[0]
                    ws.append([cat_name, cat_code, primary_url,
                               v_sku, v_name, prod_url, v_img, desc,
                               ship_label, box_w, box_h, box_d, box_wt])
            else:
                ws.append([cat_name, cat_code, primary_url,
                           item["sku"], item["name"], prod_url, parent_img, desc,
                           ship_label, box_w, box_h, box_d, box_wt])

    suffix = "_demo" if DEMO else ""
    out_path = f"moes_step1{suffix}.xlsx"
    wb.save(out_path)
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
