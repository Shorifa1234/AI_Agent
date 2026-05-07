import requests
import re
import json
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"}

for url in [
    "https://www.sherrillfurniture.com/catalog/6050",
    "https://www.sherrillfurniture.com/catalog/dc333",
    "https://www.sherrillfurniture.com/catalog/1823",
]:
    print(f"\n{'='*60}")
    print(f"URL: {url}")
    r = requests.get(url, headers=HEADERS, timeout=30)
    soup = BeautifulSoup(r.text, "html.parser")

    # Meta description
    meta = soup.find("meta", attrs={"name": "description"})
    print(f"meta description: {meta.get('content') if meta else 'NOT FOUND'}")

    # h1
    h1 = soup.find("h1")
    print(f"h1: {h1.get_text(strip=True) if h1 else 'NOT FOUND'}")

    # Images
    imgs = [img.get("src","") for img in soup.find_all("img") if "/files/catalog/" in img.get("src","")]
    print(f"catalog images ({len(imgs)}): {imgs[0] if imgs else 'NONE'}")

    # Datalayer
    m = re.search(r"dataLayer\s*=\s*(\[.*?\]);", r.text, re.DOTALL)
    if m:
        try:
            data = json.loads(m.group(1))
            tax = data[0].get("entityTaxonomy", {})
            print(f"room_style: {list(tax.get('room_style',{}).values())}")
            print(f"furniture_type: {list(tax.get('furniture_type',{}).values())}")
            print(f"room_type: {list(tax.get('room_type',{}).values())}")
        except Exception as e:
            print(f"datalayer parse error: {e}")
    else:
        print("datalayer: NOT FOUND")

    # All body text lines around product name
    lines = [l.strip() for l in soup.get_text("\n", strip=True).split("\n") if l.strip()]
    model = h1.get_text(strip=True) if h1 else ""
    for i, line in enumerate(lines):
        if model and line.upper() == model.upper():
            print(f"\nBody lines [{i}:{i+15}]:")
            for j in range(i, min(i+15, len(lines))):
                print(f"  [{j}] {lines[j]}")
            break
