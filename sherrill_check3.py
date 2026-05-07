import requests
from bs4 import BeautifulSoup
import re

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"}
BASE = "https://www.sherrillfurniture.com"

url = BASE + "/catalog/dc333"
r = requests.get(url, headers=HEADERS, timeout=30)
soup = BeautifulSoup(r.text, "html.parser")

# Print all text blocks that might have dimensions
body = soup.get_text("\n", strip=True)
lines = [l.strip() for l in body.split("\n") if l.strip()]

print("=== All lines with numbers/dimensions ===")
for i, line in enumerate(lines):
    if re.search(r'\d+[\.\d]*\s*["\']?\s*(W|D|H|Dia|inch|cm)', line, re.I):
        print(f"  Line {i}: {line}")

print("\n=== Lines near DC333 ===")
for i, line in enumerate(lines):
    if "dc333" in line.lower() or "333" in line:
        start = max(0, i-2)
        end = min(len(lines), i+10)
        for j in range(start, end):
            print(f"  [{j}] {lines[j]}")
        print()

print("\n=== All div/p with meaningful text ===")
for el in soup.find_all(["p", "div", "span", "li"]):
    txt = el.get_text(strip=True)
    if 10 < len(txt) < 300 and not el.find_all(recursive=False):
        if re.search(r'(width|depth|height|dimension|material|finish|description|\d+[\.\d]*")', txt, re.I):
            print(f"  <{el.name} class={el.get('class')}> {txt}")

print("\n=== Images ===")
for img in soup.find_all("img"):
    src = img.get("src") or ""
    if "catalog" in src:
        print(f"  alt={img.get('alt','')} | src={src}")
