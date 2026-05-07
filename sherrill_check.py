import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"}

url = "https://www.sherrillfurniture.com/products/dining-chair"
r = requests.get(url, headers=HEADERS, timeout=30)
soup = BeautifulSoup(r.text, "html.parser")

# Product tiles
tiles = soup.select("a.product-results-tile")
print(f"Product tiles: {len(tiles)}")
for t in tiles[:5]:
    href = t.get("href", "")
    img  = t.find("img")
    img_src = img.get("src", "") if img else ""
    name_el = t.find("h3")
    name = name_el.get_text(strip=True) if name_el else ""
    print(f"  href: {href}")
    print(f"  img:  {img_src}")
    print(f"  name: {name}")
    print()

# col-25 div structure
cols = soup.select(".col-25")
print(f"col-25 divs: {len(cols)}")
print("First col-25:", str(cols[0])[:600] if cols else "none")

# Pagination
pag = soup.select(".pager li, .pager-item, [class*=pager]")
print(f"\nPagination elements: {len(pag)}")

# Check detail page
if tiles:
    detail_url = "https://www.sherrillfurniture.com" + tiles[0].get("href", "")
    print(f"\nFetching detail: {detail_url}")
    r2 = requests.get(detail_url, headers=HEADERS, timeout=30)
    soup2 = BeautifulSoup(r2.text, "html.parser")
    print("Detail title:", soup2.title.get_text(strip=True) if soup2.title else "")
    print("Detail HTML snippet:", r2.text[2000:3500])
