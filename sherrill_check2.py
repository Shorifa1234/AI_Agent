import requests
from bs4 import BeautifulSoup
import re

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"}
BASE = "https://www.sherrillfurniture.com"

# Check detail page with requests
url = BASE + "/catalog/dc333"
r = requests.get(url, headers=HEADERS, timeout=30)
soup = BeautifulSoup(r.text, "html.parser")

print("=== DETAIL PAGE (requests) ===")
print("Title:", soup.title.get_text(strip=True) if soup.title else "")

# h1/h2
for tag in ["h1", "h2", "h3"]:
    els = soup.find_all(tag)
    if els:
        print(f"{tag}:", [e.get_text(strip=True) for e in els[:5]])

# Look for description, dimensions
for attr in ["description", "itemprop"]:
    els = soup.find_all(attrs={attr: True})
    if els:
        print(f"\nElements with {attr}:", [(e.name, e.get(attr), e.get_text(strip=True)[:80]) for e in els[:5]])

# og tags
og_image = soup.find("meta", property="og:image")
og_desc  = soup.find("meta", property="og:description")
print("\nog:image:", og_image.get("content") if og_image else "None")
print("og:desc:", og_desc.get("content")[:100] if og_desc else "None")

# Look for dimension/spec tables
for sel in ["table", ".dimensions", ".specs", ".product-specs", "[class*=spec]", "[class*=dimen]"]:
    els = soup.select(sel)
    if els:
        print(f"\n[{sel}] ({len(els)}):", els[0].get_text(" ", strip=True)[:200])

# Look for any text with dimensions
body_text = soup.get_text(" ", strip=True)
# Find W x D x H patterns
dims = re.findall(r'[\d.]+["\']?\s*[WwDdHh]', body_text)
print("\nDimension patterns found:", dims[:10])

# Image - look for larger version
imgs = soup.find_all("img")
for img in imgs[:10]:
    src = img.get("src") or img.get("data-src") or ""
    if "catalog" in src or "product" in src.lower():
        print("Img:", src[:120])

# Full body text snippet around "dimensions" or model number
idx = body_text.lower().find("dc333")
if idx > -1:
    print("\nText near model number:", body_text[max(0,idx-100):idx+500])
