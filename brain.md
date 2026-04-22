# SKU Scraping Agent — Brain

## Overview

`agent.py` হলো একটি AI-powered web scraping agent যা:
1. **Vendor name** নিলে Status Tracker থেকে সব category URL বের করে
2. প্রতিটি category page scrape করে সব product link collect করে
3. প্রতিটি product detail page থেকে full data extract করে
4. Output save করে `{VendorName}.xlsx` format এ (Julian Chichester.xlsx এর মতো)

---

## How to Run

```bash
# Interactive mode
python agent.py

# Command-line mode
python agent.py "Julian Chichester"
python agent.py "Brownstone"
python agent.py "Wesley Hall"
```

---

## Requirements

```bash
pip install anthropic requests beautifulsoup4 openpyxl selenium chromedriver-autoinstaller
```

Environment variable দরকার:
```
ANTHROPIC_API_KEY=sk-ant-...
```

---

## File Structure

```
SKU_AGENT/
├── agent.py                          ← Main AI agent (এটাই চালাতে হবে)
├── brain.md                          ← এই ফাইল (documentation)
├── SD_Web Scraping - Status Tracker.xlsx  ← Vendor category URLs এখানে আছে
├── Julian Chichester.xlsx            ← Sample output (এই format এ output হবে)
│
├── Brownstone/                       ← পুরানো manual scripts
│   ├── brownstone_step1.py
│   └── brownstone_step2.py
├── Wesley Hall/
│   ├── step1.py
│   └── step2.py
├── Invisible Collection/
│   ├── invisible_step1.py
│   └── invisible_step2.py
├── Paul Ferrante/
│   ├── listpage.py
│   ├── detailspage.py
│   └── step3.py
├── Vanguard/
│   ├── vanguard_step1.py
│   └── vanguard_step2.py
└── Crystorama/                       ← Empty (not yet done)
```

---

## Status Tracker Structure

`SD_Web Scraping - Status Tracker.xlsx` — Main sheet columns:

| Col | Content |
|-----|---------|
| 1 | Category (e.g. "Dining Tables") |
| 2 | Vendor name (e.g. "Julian Chichester") |
| 3 | Rank |
| 4 | **Sample Link** ← category URL যেটা scrape করতে হবে |
| 5 | Sample status |
| 7 | Scraping Status (True/False) |
| 8 | Manually Status |
| 9 | Submission Done |

---

## Output Excel Format

প্রতিটি output file এ প্রতিটি category এর জন্য আলাদা sheet থাকবে।

**Sheet structure:**
```
Row 1:  Brand        | {Vendor Name}
Row 2:  Category Link| {Category URL}
Row 3:  (empty)
Row 4:  Index | Category | Manufacturer | Source | Image URL | Product Name | SKU | Product Family Id | Description | Width | Depth | Height | Diameter | Finish | [extra cols...]
Row 5+: data...
```

**Sample — Julian Chichester.xlsx sheets:**
- Nightstands → https://julianchichester.com/category/bookshelves
- Dining Table → https://julianchichester.com/category/dining-tables
- Coffee & Cocktail Tables → https://julianchichester.com/category/coffee-tables
- Side & End Tables → https://julianchichester.com/category/occasional-tables
- Consoles → https://julianchichester.com/category/console-tables
- Beds & Headboards → https://julianchichester.com/category/beds-daybeds
- Desks → https://julianchichester.com/category/desks-dressing-tables
- Dining Chairs → https://julianchichester.com/category/dining-chairs
- Lounge Chairs → https://julianchichester.com/category/lounge-chairs
- Bar Stools → https://julianchichester.com/category/bar-stools
- Sofas & Loveseats → https://julianchichester.com/category/settees-sofas-chaises
- Mirrors → https://julianchichester.com/category/mirrors
- Chandeliers, Sconces, Table Lamps, Floor Lamps, Objects

---

## SKU Generation Rule

Format: `[Vendor 3 letters][Category initials][2-digit number]`

| Vendor | Category | SKU Pattern |
|--------|----------|-------------|
| Julian Chichester | Nightstands | JULNI01, JULNI02 … |
| Julian Chichester | Dining Tables | JULDI01, JULDI02 … |
| Julian Chichester | Coffee Tables | JULCO01 … |
| Brownstone | Nightstands | BRONI01 … |
| Wesley Hall | Lounge Chairs | WESLO01 … |
| Vanguard | Bedroom | VANBR01 … |

---

## Tools Available to Claude

### 1. `get_vendor_categories(vendor_name)`
Status Tracker থেকে vendor এর সব category + URL বের করে।

**Returns:**
```json
{
  "vendor": "Julian Chichester",
  "total_categories": 17,
  "categories": [
    {"category": "Nightstands & End Tables", "url": "https://...", "already_scraped": false},
    ...
  ]
}
```

### 2. `fetch_page_html(url, use_selenium=false, wait_seconds=3)`
একটি page এর cleaned HTML return করে।

- `use_selenium=false` → requests দিয়ে fetch (fast, non-JS sites)
- `use_selenium=true` → Chrome browser দিয়ে fetch (JS-heavy sites)

**Returns:**
```json
{
  "url": "https://...",
  "char_count": 15234,
  "html": "cleaned HTML content..."
}
```

### 3. `save_products(vendor_name, category_name, category_url, products, output_file)`
Products save করে Excel এ।

**products array structure:**
```json
[
  {
    "Product Name": "Bay Bookcase",
    "Source": "https://julianchichester.com/product/bay-bookcase",
    "Image URL": "https://...",
    "Description": "Adapted from a 50s design...",
    "Width": "59",
    "Depth": "20",
    "Height": "79",
    "Finish": "High Gloss Teal Vellum"
  }
]
```

---

## Agent Flow (Step by Step)

```
User input: "Julian Chichester"
    ↓
get_vendor_categories("Julian Chichester")
    → 17 categories found
    ↓
For category "Dining Tables":
  fetch_page_html("https://julianchichester.com/category/dining-tables")
    → HTML with product cards
  Claude reads HTML → extracts product URLs (e.g. 12 products)
  If pagination found → fetch next pages too
    ↓
  For each product URL:
    fetch_page_html("https://julianchichester.com/product/canopy-dining-table")
      → Claude extracts: name, image, description, dimensions, finish
    ↓
  save_products(vendor="Julian Chichester", category="Dining Tables", products=[...])
    ↓
For category "Coffee Tables":
  ... (same process)
    ↓
... (all 17 categories)
    ↓
DONE → Julian Chichester.xlsx created
```

---

## Dimension Rules

- সব dimension **INCHES** এ থাকতে হবে
- cm → inch: `value / 2.54` (2 decimal places)
- mm → inch: `value / 25.4` (2 decimal places)
- Round items এর জন্য: `Diameter` column use করো, `Width`/`Depth` empty রাখো
- Range দিলে (e.g. "60–72"): smaller value use করো

---

## JavaScript Sites

কিছু site JavaScript দিয়ে products load করে। এদের জন্য `use_selenium=true` দরকার।

Known JS-heavy vendors:
- **Invisible Collection** — Algolia-based search, needs Selenium
- **Paul Ferrante** — Requires remote debugging Chrome

Non-JS vendors (requests কাজ করে):
- **Julian Chichester** — Standard WordPress/WooCommerce
- **Wesley Hall** — Traditional HTML
- **Brownstone** — Standard HTML (lazy-load images only)
- **Vanguard** — Standard HTML pagination

---

## Existing Manual Scripts Reference

পুরানো vendor-specific scripts যেগুলো এখনো কাজে লাগতে পারে:

| Vendor | Script | What it does |
|--------|--------|-------------|
| Brownstone | `brownstone_step1.py` | Nightstands list page scraper |
| Brownstone | `brownstone_step2.py` | Detail page enricher |
| Wesley Hall | `step1.py` | Category list scraper (requests) |
| Wesley Hall | `step2.py` | Detail page scraper |
| Invisible Collection | `invisible_step1.py` | Selenium + Algolia list scraper |
| Invisible Collection | `invisible_step2.py` | Detail + tearsheet downloader |
| Paul Ferrante | `listpage.py` | Remote debugging Chrome scraper |
| Paul Ferrante | `detailspage.py` | WooCommerce attributes scraper |
| Vanguard | `vanguard_step1.py` | Pagination list scraper |
| Vanguard | `vanguard_step2.py` | Detail page scraper |

---

## Configuration

`agent.py` এর top এ config variables:

```python
MODEL = "claude-sonnet-4-6"    # Claude model
MAX_HTML_CHARS = 30000          # Max HTML per page fetch
MAX_ITERATIONS = 300            # Agent loop safety limit
CHROMEDRIVER_PATH = "C:/chromedriver.exe"  # Selenium fallback path
```

---

## Common Issues

| Problem | Solution |
|---------|---------|
| `No categories found for 'X'` | Vendor name spelling check করো। Exact match না লাগলে partial match হয়। |
| `Selenium fetch failed` | Chrome + chromedriver installed আছে কিনা চেক করো |
| Page loads but no products | Site JS-heavy — `use_selenium=true` দিয়ে retry করবে Claude |
| Huge HTML (>60K chars) | agent.py automatically compress করে key elements রেখে |
| API key error | `ANTHROPIC_API_KEY` environment variable set আছে কিনা দেখো |

---

## Notes on Julian Chichester Website

- Base URL: `https://julianchichester.com`
- Category pages: `/category/{slug}/`
- Product pages: `/product/{slug}/`
- Dimensions format: `"W 59" x D 20" x H 79"` (inches এ already)
- Finish info: product description তে থাকে
- Images: BunnyCDN hosted (`b-cdn.net`)
- No login required — public site
