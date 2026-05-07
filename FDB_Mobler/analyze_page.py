import requests
from bs4 import BeautifulSoup

url = "https://www.fdbmobler.com/collections/coffee-tables"
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
}

response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.content, 'html.parser')

# Find product links
product_links = []
for a in soup.find_all('a', href=True):
    href = a['href']
    if '/products/' in href and href not in product_links:
        full_url = 'https://www.fdbmobler.com' + href if href.startswith('/') else href
        product_links.append(full_url)

print(f"Total unique product URLs found: {len(product_links)}\n")
print("First 5 products:")
for i, link in enumerate(product_links[:5], 1):
    print(f"{i}. {link}")

# Check for pagination
print("\n--- Pagination Check ---")
pagination = soup.find_all('a', href=lambda x: x and '?page=' in x)
print(f"Pagination links found: {len(pagination)}")
if pagination:
    for p in pagination[:3]:
        print(f"  {p.get('href')}")

# Try to find product grid
print("\n--- Product Grid Analysis ---")
product_grid = soup.find('div', {'id': 'product-grid'}) or soup.find('div', class_='product-grid')
if product_grid:
    products = product_grid.find_all('div', class_=lambda x: x and 'product-item' in str(x).lower())
    print(f"Products in grid: {len(products)}")
else:
    print("Product grid not found, searching for product cards...")
    # Try to find product items by different selectors
    items = soup.find_all('div', {'class': lambda x: x and ('card' in str(x).lower() or 'item' in str(x).lower())})
    print(f"Potential product items: {len(items)}")
