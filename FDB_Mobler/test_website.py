import requests
from bs4 import BeautifulSoup
import json

# Test if FDB Mobler is a Shopify site
url = "https://www.fdbmobler.com/collections/coffee-tables"

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

print(f"Testing: {url}\n")

try:
    response = requests.get(url, headers=headers, timeout=15)
    print(f"Status Code: {response.status_code}")

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')

        # Check if Shopify
        is_shopify = 'Shopify' in response.text or 'cdn.shopify.com' in response.text
        print(f"Is Shopify: {is_shopify}")

        # Try Shopify JSON API
        if is_shopify:
            print("\n--- Testing Shopify JSON API ---")
            json_url = url + ".json"
            json_response = requests.get(json_url, headers=headers, timeout=15)
            print(f"JSON API Status: {json_response.status_code}")

            if json_response.status_code == 200:
                data = json_response.json()
                print(f"Products in JSON: {len(data.get('products', []))}")
                if data.get('products'):
                    print(f"\nFirst product example:")
                    p = data['products'][0]
                    print(f"  Title: {p.get('title')}")
                    print(f"  Handle: {p.get('handle')}")
                    print(f"  Product URL: https://www.fdbmobler.com/products/{p.get('handle')}")

        # Check pagination
        print("\n--- Checking Pagination ---")
        next_link = soup.find('a', {'rel': 'next'}) or soup.find('a', text='Next')
        print(f"Next page button found: {next_link is not None}")

        # Count products on page
        product_items = soup.find_all('div', class_=lambda x: x and 'product' in x.lower())
        print(f"Product divs found: {len(product_items)}")

except Exception as e:
    print(f"Error: {e}")
