import requests
from bs4 import BeautifulSoup
import json
import re

url = "https://www.fdbmobler.com/products/d102-sos-sofabord-eg-olieret-o55-1"

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
}

response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.content, 'html.parser')

print("=" * 80)
print("PRODUCT PAGE ANALYSIS")
print("=" * 80)

# 1. Try to find JSON-LD data
json_ld = soup.find('script', {'type': 'application/ld+json'})
if json_ld:
    print("\n--- JSON-LD Data Found ---")
    data = json.loads(json_ld.string)
    print(json.dumps(data, indent=2)[:1000])

# 2. Try to find product title
title = soup.find('h1')
print(f"\n--- Title ---")
print(title.get_text(strip=True) if title else "Not found")

# 3. Try to find description
desc_div = soup.find('div', {'class': lambda x: x and 'description' in str(x).lower()})
if not desc_div:
    desc_div = soup.find('div', {'class': lambda x: x and 'product' in str(x).lower() and 'content' in str(x).lower()})
print(f"\n--- Description ---")
print(desc_div.get_text(strip=True)[:200] if desc_div else "Not found")

# 4. Try to find dimensions
print(f"\n--- Searching for Dimensions ---")
text = soup.get_text()
# Look for dimension patterns
dim_patterns = [
    r'(\d+(?:\.\d+)?)\s*(?:x|×)\s*(\d+(?:\.\d+)?)\s*(?:x|×)\s*(\d+(?:\.\d+)?)\s*cm',
    r'(?:W|Width):\s*(\d+(?:\.\d+)?)\s*cm',
    r'(?:H|Height):\s*(\d+(?:\.\d+)?)\s*cm',
    r'(?:D|Depth):\s*(\d+(?:\.\d+)?)\s*cm',
    r'(?:Ø|Diameter):\s*(\d+(?:\.\d+)?)\s*cm',
]

for pattern in dim_patterns:
    matches = re.findall(pattern, text, re.IGNORECASE)
    if matches:
        print(f"  Pattern '{pattern}': {matches[:3]}")

# 5. Try to find image
main_img = soup.find('img', {'class': lambda x: x and ('product' in str(x).lower() or 'main' in str(x).lower())})
if not main_img:
    main_img = soup.find('img', {'src': lambda x: x and 'products/' in str(x)})
print(f"\n--- Main Image ---")
if main_img:
    img_src = main_img.get('src', main_img.get('data-src', ''))
    print(img_src[:150])
else:
    print("Not found")

# 6. Look for SKU/Product ID
print(f"\n--- SKU/Product ID ---")
sku_meta = soup.find('meta', {'property': 'product:sku'}) or soup.find('span', {'class': lambda x: x and 'sku' in str(x).lower()})
if sku_meta:
    print(sku_meta.get('content', sku_meta.get_text(strip=True)))
else:
    print("Not found in meta/span")

# 7. Check for product JSON data in script tags
print(f"\n--- Checking for Product JSON in Scripts ---")
for script in soup.find_all('script'):
    if script.string and 'product' in script.string.lower() and '{' in script.string:
        # Try to extract JSON
        try:
            # Look for JSON object
            start = script.string.find('{')
            if start > -1:
                print(f"Found potential JSON data (first 300 chars):")
                print(script.string[start:start+300])
                break
        except:
            pass
