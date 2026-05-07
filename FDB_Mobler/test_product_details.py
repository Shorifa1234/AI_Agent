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
print("EXTRACTING PRODUCT DATA")
print("=" * 80)

# Extract JSON-LD
json_ld_scripts = soup.find_all('script', {'type': 'application/ld+json'})
product_data = {}

for script in json_ld_scripts:
    try:
        data = json.loads(script.string)
        if data.get('@type') == 'Product':
            product_data = data
            break
    except:
        pass

# Print structured data
print("\nProduct Name:", product_data.get('name', 'N/A'))
print("Brand:", product_data.get('brand', {}).get('name', 'N/A'))
print("Category:", product_data.get('category', 'N/A'))
print("Description:", product_data.get('description', 'N/A')[:200])

# Get image
if 'image' in product_data:
    images = product_data['image']
    if isinstance(images, list):
        print("\nMain Image:", images[0] if images else 'N/A')
    else:
        print("\nMain Image:", images)

# Get offers/price info
if 'offers' in product_data:
    offers = product_data['offers']
    print("\nPrice:", offers.get('price', 'N/A'))
    print("Currency:", offers.get('priceCurrency', 'N/A'))
    print("Availability:", offers.get('availability', 'N/A'))

# Look for dimensions in page
print("\n" + "=" * 80)
print("SEARCHING FOR DIMENSIONS")
print("=" * 80)

# Get all text content
page_text = soup.get_text()

# Common dimension patterns
patterns = {
    'WxDxH (cm)': r'(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*cm',
    'Diameter': r'[ØD](?:iameter)?:?\s*(\d+(?:\.\d+)?)\s*cm',
    'Height': r'[Hh](?:eight)?:?\s*(\d+(?:\.\d+)?)\s*cm',
    'Width': r'[Ww](?:idth)?:?\s*(\d+(?:\.\d+)?)\s*cm',
    'Depth': r'[Dd](?:epth)?:?\s*(\d+(?:\.\d+)?)\s*cm',
}

for name, pattern in patterns.items():
    matches = re.findall(pattern, page_text)
    if matches:
        print(f"{name}: {matches[0]}")

# Check product title for dimensions
title = soup.find('h1')
if title:
    title_text = title.get_text()
    print(f"\nTitle: {title_text}")
    # Extract from title (often has Ø55, etc.)
    diam_match = re.search(r'[ØO](\d+)', title_text)
    if diam_match:
        print(f"  -> Diameter from title: {diam_match.group(1)} cm")

# Look for specification table
spec_tables = soup.find_all('table')
if spec_tables:
    print("\n--- Specification Tables ---")
    for table in spec_tables[:2]:
        rows = table.find_all('tr')
        for row in rows[:5]:
            cols = row.find_all(['td', 'th'])
            if len(cols) >= 2:
                print(f"  {cols[0].get_text(strip=True)}: {cols[1].get_text(strip=True)}")

# Look for detail/spec divs
spec_divs = soup.find_all('div', {'class': lambda x: x and any(word in str(x).lower() for word in ['spec', 'detail', 'dimension', 'measurement'])})
if spec_divs:
    print("\n--- Specification Divs ---")
    for div in spec_divs[:2]:
        print(div.get_text(strip=True)[:300])
