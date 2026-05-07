import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"}
url = "https://www.sherrillfurniture.com/products/ottoman"
r = requests.get(url, headers=HEADERS, timeout=30)
soup = BeautifulSoup(r.text, "html.parser")

print("=== ALL catalog links ===")
for a in soup.find_all("a", href=True):
    href = a.get("href", "")
    if "/catalog/" in href:
        print(f"  class={a.get('class')} | href={href} | text={a.get_text(strip=True)[:50]}")

print("\n=== All anchors with /products/ ===")
for a in soup.find_all("a", href=True):
    href = a.get("href", "")
    if "/products/" in href and href != "/products/ottoman":
        print(f"  href={href} | text={a.get_text(strip=True)[:50]}")

print("\n=== Page snippet around 'ottoman' text ===")
lines = [l.strip() for l in soup.get_text("\n").split("\n") if l.strip()]
for i, line in enumerate(lines):
    if "ottoman" in line.lower() and len(line) < 200:
        print(f"  [{i}] {line}")
