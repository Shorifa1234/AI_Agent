import requests
from bs4 import BeautifulSoup
import json
import re
import time

test_urls = [
    "https://www.fdbmobler.com/products/d102-sos-sofabord-eg-olieret-o55-1",
    "https://www.fdbmobler.com/products/a40-spisebord-eg-sortmalet-l220",
    "https://www.fdbmobler.com/products/f212-loftlampe-messing",
]

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
}

def extract_product_data(url):
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'html.parser')

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

    # Extract image URL
    image_url = ''
    if 'image' in product_data:
        images = product_data['image']
        if isinstance(images, list):
            image_url = images[0].get('url', '') if isinstance(images[0], dict) else images[0]
        elif isinstance(images, dict):
            image_url = images.get('url', '')
        else:
            image_url = str(images)

    # Extract dimensions from title
    name = product_data.get('name', '')

    # Look for dimensions in name
    width = depth = height = diameter = ''

    # Diameter pattern (Ø55, O55, etc.)
    diam_match = re.search(r'[ØO](\d+)', name)
    if diam_match:
        diameter = str(round(float(diam_match.group(1)) / 2.54, 2))  # Convert cm to inches

    # WxDxH pattern (220x95x72)
    wdh_match = re.search(r'(\d+)\s*x\s*(\d+)\s*x\s*(\d+)', name)
    if wdh_match:
        width = str(round(float(wdh_match.group(1)) / 2.54, 2))
        depth = str(round(float(wdh_match.group(2)) / 2.54, 2))
        height = str(round(float(wdh_match.group(3)) / 2.54, 2))

    # Length pattern (L220, W180, etc.)
    len_match = re.search(r'[LWl](\d+)', name)
    if len_match and not width:
        width = str(round(float(len_match.group(1)) / 2.54, 2))

    return {
        'name': name,
        'brand': product_data.get('brand', {}).get('name', ''),
        'description': product_data.get('description', ''),
        'image_url': image_url,
        'width': width,
        'depth': depth,
        'height': height,
        'diameter': diameter,
    }

print("=" * 80)
print("TESTING MULTIPLE PRODUCTS")
print("=" * 80)

for url in test_urls:
    print(f"\n\n{url}")
    print("-" * 80)

    data = extract_product_data(url)

    print(f"Name: {data['name']}")
    print(f"Brand: {data['brand']}")
    print(f"Description: {data['description'][:100]}...")
    print(f"Image: {data['image_url'][:80]}...")
    print(f"Width: {data['width']} inches" if data['width'] else "Width: Not found")
    print(f"Depth: {data['depth']} inches" if data['depth'] else "Depth: Not found")
    print(f"Height: {data['height']} inches" if data['height'] else "Height: Not found")
    print(f"Diameter: {data['diameter']} inches" if data['diameter'] else "Diameter: Not found")

    time.sleep(1)
