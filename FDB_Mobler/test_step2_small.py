"""Quick test of Step 2 with just 5 products"""
import sys
sys.path.insert(0, '.')

import requests
from bs4 import BeautifulSoup
import openpyxl
import json
import re

HEADERS = {'User-Agent': 'Mozilla/5.0'}

# Test with first 5 products
wb = openpyxl.load_workbook('FDB_Mobler_step1.xlsx')
ws = wb.worksheets[0]

print("Testing Step 2 with 5 products from first category...")
print(f"Category: {ws.title}\n")

count = 0
for row in ws.iter_rows(min_row=2, max_row=6, values_only=True):
    if row[1]:
        count += 1
        url = row[1]
        print(f"{count}. {url}")

        try:
            response = requests.get(url, headers=HEADERS, timeout=15)
            soup = BeautifulSoup(response.content, 'html.parser')

            # Try JSON-LD
            json_ld = soup.find('script', {'type': 'application/ld+json'})
            if json_ld:
                data = json.loads(json_ld.string)
                if data.get('@type') == 'Product':
                    name = data.get('name', 'N/A')
                    print(f"   Name: {name[:60]}")

                    # Extract dimensions
                    diam = re.search(r'[ØOø](\d+)', name)
                    if diam:
                        print(f"   Diameter: {round(float(diam.group(1))/2.54, 2)} inches")
                    else:
                        print("   Diameter: Not found")
                else:
                    print("   JSON-LD found but not Product type")
            else:
                print("   No JSON-LD found")

        except Exception as e:
            print(f"   ERROR: {e}")

        print()

print("Test complete!")
