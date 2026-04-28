# SKU Scraping Agent — Instructions

## Project Overview

This project scrapes furniture/decor vendor websites and outputs Excel files in a standard format (based on Julian Chichester.xlsx).

**Workflow:**
```
Vendor Name Input
    → Status Tracker lookup (SD_Web Scraping - Status Tracker.xlsx)
    → Scrape category list pages → collect product URLs
    → Scrape each product detail page → collect data
    → Save to {VendorName}.xlsx
```

---

## Folder Structure

```
SKU_AGENT/
├── agent.py                              ← AI agent (run this for automated scraping)
├── brain.md                              ← Agent documentation
├── INSTRUCTIONS.md                       ← This file
├── SD_Web Scraping - Status Tracker.xlsx ← Vendor category URLs
├── Julian Chichester.xlsx                ← Sample output format
│
├── Brownstone/           (step1 + step2)
├── Wesley Hall/          (step1 + step2)
├── Invisible Collection/ (step1 + step2)
├── Paul Ferrante/        (listpage + detailspage + step3)
├── Vanguard/             (step1 + step2)
├── Crystorama/           (step1 → xlsx done)
├── CrystoramaLights/     (step1 + step2 → COMPLETE, 2259 products, 6 sheets)
├── BlackmanCruz/         (step1 + step2 → Shopify JSON API)
└── Gabby/                (step1 + step2 → Shopify JSON API, 26 categories)
```

---

## Output Excel Format (Standard for ALL Vendors)

Every vendor output file follows this sheet structure:

| Row | Column A | Column B |
|-----|----------|----------|
| 1 | Brand | {Vendor Name} |
| 2 | Category Link | {Category URL} |
| 3 | *(empty)* | |
| 4 | **Headers** → | Index, Category, Manufacturer, Source, Image URL, Product Name, SKU, Product Family Id, Description, Width, Depth, Height, Diameter, Finish |
| 5+ | Data rows | |

- One **sheet per category**
- Sheet name = category name (e.g. "Dining Tables", "Lounge Chairs")
- One output file per vendor: `{VendorName}.xlsx`

---

## SKU Generation Rule

Format: `[3-letter vendor code][2-letter category code][2-digit sequence]`

**Examples:**

| Vendor | Category | SKU |
|--------|----------|-----|
| Julian Chichester | Dining Tables | JULDI01, JULDI02 … |
| Julian Chichester | Nightstands | JULNI01, JULNI02 … |
| Brownstone | Nightstands | BRONI01, BRONI02 … |
| Wesley Hall | Lounge Chairs | WESLO01, WESLO02 … |
| Vanguard | Bedroom | VANBR01, VANBR02 … |
| Crystorama | Chandeliers | CRYCH01, CRYCH02 … |

**Rules:**
- Vendor code = first 3 letters of vendor name (uppercase)
- Category code = first 2 letters of category (uppercase), or meaningful abbreviation
- Sequence starts at 01, zero-padded to 2 digits

---

## Dimension Rules

- All dimensions must be in **INCHES**
- Convert cm → inch: `value / 2.54` (round to 2 decimal places)
- Convert mm → inch: `value / 25.4` (round to 2 decimal places)
- Round/circular items: use **Diameter** column, leave Width/Depth empty
- Range values (e.g. "60–72 inches"): use the **smaller** value

---

## How to Build a New Vendor Script

### Step 1 — List Page Scraper

**Goal:** Collect all product URLs + basic data from category listing pages.

**What to extract per product:**
- Product URL (detail page link)
- Image URL
- Product Name
- SKU (if visible on list page)

**Output:** Save to `{VendorName}_step1.xlsx`

**Pagination:** Always check for pagination (next page button, URL `?page=2`, etc.) and loop through all pages.

**Selenium vs Requests:**
- Use `requests` + `BeautifulSoup` for standard HTML sites (faster)
- Use `Selenium` + `chromedriver` for JavaScript/lazy-load sites

**ChromeDriver path:** `C:/chromedriver.exe`

---

### Step 2 — Detail Page Scraper

**Goal:** Read product URLs from Step 1 Excel, visit each detail page, extract full data, save final output.

**What to extract per product:**
- Product Name
- SKU
- Description
- Width, Depth, Height (in inches)
- Diameter (if round item)
- Finish / Color / Material
- Image URL
- Product Family / Collection (if available)

**Output:** Save to `{VendorName}.xlsx` in standard format (see Output Format section above)

**Rate limiting:** Add `time.sleep(1)` between requests to avoid getting blocked.

---

## Scraping Methods by Vendor Type

| Vendor Type | Method | Tool |
|-------------|--------|------|
| Standard HTML | `requests` + `BeautifulSoup` | Fast, no browser |
| Lazy-load images | `Selenium` scroll to bottom | Chrome needed |
| JavaScript SPA | `Selenium` | Chrome needed |
| Algolia search | Direct Algolia API call | No Selenium needed |
| JSON-LD data | `json.loads(script tag)` | Beautiful Soup |

---

## Data Fields Reference

| Column | Description | Notes |
|--------|-------------|-------|
| Index | Auto-number | 1, 2, 3 … |
| Category | Category name | e.g. "Dining Tables" |
| Manufacturer | Vendor name | e.g. "Vanguard" |
| Source | Product detail page URL | Full URL |
| Image URL | Main product image | Full URL |
| Product Name | Product title | As on website |
| SKU | Generated SKU | See SKU rules above |
| Product Family Id | Collection/family name | Optional |
| Description | Product description | Clean text, no HTML |
| Width | Width in inches | Numbers only |
| Depth | Depth in inches | Numbers only |
| Height | Height in inches | Numbers only |
| Diameter | Diameter in inches | Only for round items |
| Finish | Finish/color/material | e.g. "Brushed Nickel" |

---

## Running the Scripts

### Manual Vendor Scripts (Step 1 → Step 2)
```bash
# Step 1: collect list page data
python vanguard_step1.py

# Step 2: enrich with detail page data
python vanguard_step2.py
```

### AI Agent (Automated)
```bash
# Run interactively
python agent.py

# Run with vendor name argument
python agent.py "Julian Chichester"
python agent.py "Wesley Hall"
```

---

## Requirements

```bash
pip install anthropic requests beautifulsoup4 openpyxl selenium chromedriver-autoinstaller
```

**Environment variable required for agent.py:**
```
ANTHROPIC_API_KEY=sk-ant-...
```

---

## Status Tracker (SD_Web Scraping - Status Tracker.xlsx)

Main sheet column layout:

| Column | Content |
|--------|---------|
| 1 | Category (e.g. "Dining Tables") |
| 2 | Vendor name (e.g. "Julian Chichester") |
| 3 | Rank |
| 4 | **Category URL** ← this is what we scrape |
| 5 | Sample status |
| 7 | Scraping Status (True/False) |
| 8 | Manually Status |
| 9 | Submission Done |

---

## Completed Vendors

| Vendor | Script | Status | Notes |
|--------|--------|--------|-------|
| Julian Chichester | agent.py | ✅ Complete | Standard HTML, 13+ categories |
| Brownstone | step1 + step2 | ✅ Complete | Selenium, CSS selectors |
| Wesley Hall | step1 + step2 | ✅ Complete | requests + BeautifulSoup |
| Invisible Collection | step1 + step2 | ✅ Complete | Selenium + Algolia |
| Paul Ferrante | listpage + detailspage | ✅ Complete | Remote debugging Chrome |
| Vanguard Designs | step1 + step2 | ✅ Complete | Selenium, 14 categories, reads Status Tracker sheet |
| Crystorama | step1 | ✅ Complete | requests + BeautifulSoup |
| CrystoramaLights | step1 + step2 | ✅ Complete | JSON-LD, 2259 products, 6 sheets |
| Blackman Cruz | step1 + step2 | ✅ Complete | Shopify JSON API, multi-category |
| Gabby | step1 + step2 | ✅ Complete | Shopify JSON API, 26 categories |

---

## Common Issues & Fixes

| Problem | Solution |
|---------|---------|
| No products found on list page | Site may be JS-heavy — switch to Selenium |
| Images not loading | Use Selenium + scroll to bottom (lazy-load) |
| Blocked / 403 error | Add User-Agent header + `time.sleep()` between requests |
| Dimensions in cm | Convert: `cm / 2.54` for inches |
| Wrong product count | Check pagination — may have multiple pages |
| ChromeDriver error | Make sure Chrome browser is installed, chromedriver at `C:/chromedriver.exe` |
| Excel not saving | Close the Excel file before running the script |
| API key error | Set `ANTHROPIC_API_KEY` environment variable |
