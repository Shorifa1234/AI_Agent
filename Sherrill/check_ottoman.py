import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"}
url = "https://www.sherrillfurniture.com/products/ottoman"
r = requests.get(url, headers=HEADERS, timeout=30)
soup = BeautifulSoup(r.text, "html.parser")

print("h1:", soup.find("h1").get_text(strip=True) if soup.find("h1") else "N/A")
meta = soup.find("meta", attrs={"name": "description"})
print("meta:", meta.get("content", "")[:120] if meta else "N/A")

print("\nAll catalog links:")
for a in soup.find_all("a", href=True):
    href = a.get("href", "")
    if "/catalog/" in href:
        cls = a.get("class", [])
        print(f"  href={href} | class={cls} | text={a.get_text(strip=True)[:40]}")

print("\nAll product tile selectors:")
for sel in ["a.product-results-tile", "div.views-row a", ".product-tile a", "li.views-row a"]:
    found = soup.select(sel)
    print(f"  {sel}: {len(found)}")

print("\nAll divs with 'views-row' class:")
rows = soup.select(".views-row")
print(f"  .views-row count: {len(rows)}")
if rows:
    print("  first row html:", str(rows[0])[:200])
