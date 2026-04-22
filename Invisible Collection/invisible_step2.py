"""
Selenium scraper for The Invisible Collection
Setup:
  pip install selenium webdriver-manager pandas openpyxl
"""
import pandas as pd
import re
import time
import random
import os
import sys
import glob
import shutil

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
except ImportError:
    os.system(f"{sys.executable} -m pip install selenium webdriver-manager pandas openpyxl")
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains

try:
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    os.system(f"{sys.executable} -m pip install webdriver-manager")
    from webdriver_manager.chrome import ChromeDriverManager

# ============ CONFIG ============
INPUT_FILE = "invisible_collection_Bookcase.xlsx"
OUTPUT_FILE = "product_details_output.xlsx"
HEADLESS = True  # False korle browser dekhte parben
VENDOR_NAME = "Invisible Collection"
CATEGORY_NAME = "Bookcase"     # Change per category
TEARSHEET_FOLDER = os.path.join(os.getcwd(), "tearsheets")
DOWNLOAD_FOLDER = os.path.join(os.getcwd(), "downloads_temp")
# ================================

os.makedirs(TEARSHEET_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)


def generate_sku(vendor, category, index):
    """SKU = first 3 of vendor + first letter each category word + index
    INVDC01, INVDC02..."""
    v = re.sub(r'[^A-Za-z]', '', vendor).upper()[:3]
    words = category.strip().split()
    if len(words) >= 2:
        c = (words[0][0] + words[1][0]).upper()
    else:
        c = re.sub(r'[^A-Za-z]', '', category).upper()[:2]
    return f"{v}{c}{index:02d}"


def get_driver():
    options = Options()
    if HEADLESS:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    prefs = {
        "download.default_directory": DOWNLOAD_FOLDER,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
            window.chrome = {runtime: {}};
        """
    })
    return driver


def clean_text(text):
    if not text:
        return ""
    return re.sub(r'\s+', ' ', text).strip()


def safe_text(driver, css_selector):
    try:
        el = driver.find_element(By.CSS_SELECTOR, css_selector)
        return clean_text(el.text)
    except:
        return ""



def normalize_dim_text(text):
    """
    Pre-process dimension text to normalize all variants before regex parsing.
    - × (U+00D7) multiply sign → x
    - '' double-single-quote → ″ inch symbol
    - comma decimal separators → period  (27,3 → 27.3)
    - L dimension prefix → W  (L 25 x → W 25 x)
    - "Small – for 1 candle:" → "Small:"  (dash + extra words in size label)
    - Truncates secondary item labels (Chain:, Vase:, Left:, Stool: …)
    """
    # × → x
    text = text.replace('\u00d7', 'x')
    # '' → ″
    text = text.replace("''", '\u2033')
    # comma decimals: 27,3 → 27.3
    text = re.sub(r'(\d),(\d)', r'\1.\2', text)
    # L dimension prefix → W  (but not "Large", "Left" etc — must be followed by digit)
    text = re.sub(r'\bL\s*(?=[\d.])', 'W ', text)
    # "Small – for 1 candle:" → "Small:"
    text = re.sub(
        r'\b(Small|Medium|Large|Extra\s+Large)\s*[\u2013\u2014-][^:]*:',
        r'\1:', text, flags=re.IGNORECASE
    )
    # Truncate secondary item labels that appear after main dims
    text = re.sub(
        r'\s*(Chain|Vase\s*\d*|Left|Right|Frame\s+dimensions?|By\s+object|'
        r'Pen\s+holder|Paper\s+clip|Post-it|Desk\s+mat|Stool|Valet|'
        r'Candle\s+diameter|Height\s+with\s+candle|Recommended|Weight|'
        r'Maximum\s+bespoke)\s*:.*',
        '', text, flags=re.IGNORECASE | re.DOTALL
    )
    # Ensure space after size-label colon: 'Small:W' → 'Small: W'
    text = re.sub(
        r'\b(Small|Medium|Large|Square|Extra\s+Large):\s*([WDHL])',
        r'\1: \2', text, flags=re.IGNORECASE
    )
    # Ensure space between dimension letter and digit: 'W46' → 'W 46'
    text = re.sub(r'\b([WDHL])(\d)', r'\1 \2', text)
    return text.strip()

def parse_dimensions(dim_text):
    """
    Parse dimension fields from text - inch values only.

    Priority rules for size variants:
      - Small + Large present  → take Small
      - Medium + Large present → take Medium
      - Only one size          → take that one
      - No size labels         → take first W x D x H inch values found

    ── RUG PATTERNS (W x H, no Depth) ──
      Small: 250 x 300 cm / 98.4″ x 118.1″
      Large: 300 x 400 cm / 118.1″ x 157.4″
      Standard dimensions:\n300 x 400 cm / 118″ x 157″
      300 x 400 cm / 118″ x 157″
    """
    result = {
        "Width (inch)": "", "Depth (inch)": "", "Height (inch)": "",
        "Diameter (inch)": "", "Seat Height (inch)": "", "Seat Depth (inch)": "",
        "Fabric Required (yd)": "",
    }
    # Normalize all variant characters/formats before parsing
    dim_text = normalize_dim_text(dim_text)
    if not dim_text:
        return result

    # ══════════════════════════════════════════════════════════════════════
    # ── RUG Pattern A: size-labelled 2D  (Small/Medium/Large: NNN x NNN cm / W″ x H″)
    # ══════════════════════════════════════════════════════════════════════
    rug_size_map = {}
    rug_size_pat = re.compile(
        r'(Small|Medium|Large|Square)[\s\u2013\u2014-]*:\s*[\d.]+\s*x\s*[\d.]+\s*cm\s*/\s*([\d.]+)[\u2033\u2032\u201d"]*\s*x\s*([\d.]+)',
        re.IGNORECASE
    )
    for m in rug_size_pat.finditer(dim_text):
        rug_size_map[m.group(1).lower()] = (m.group(2), m.group(3))

    if rug_size_map:
        _rug_pri = ['small', 'medium', 'large', 'square']
        chosen = next((rug_size_map[k] for k in _rug_pri if k in rug_size_map),
                      list(rug_size_map.values())[0])
        result["Width (inch)"] = chosen[0]
        result["Height (inch)"] = chosen[1]
        # Rug has no depth — skip W x D x H patterns below
        # Still parse Diameter / Seat fields before returning
    # ══════════════════════════════════════════════════════════════════════
    # ── OBJECT Pattern: size-labelled W x H (no Depth)
    #    Small/Medium/Large/Extra Large: W N x H N cm / W N″ x H N″
    # ══════════════════════════════════════════════════════════════════════
    if not result["Width (inch)"]:
        obj_size_map = {}
        obj_size_pat = re.compile(
            r'(Small|Medium|Large|Extra\s+Large)\s*:\s*'
            r'W\s*[\d.]+\s*x\s*H\s*[\d.]+\s*cm'
            r'\s*/\s*'
            r'W\s*([\d.]+)[\u2033\u201d\u2019\u2032"]*\s*x\s*H\s*([\d.]+)',
            re.IGNORECASE
        )
        for m in obj_size_pat.finditer(dim_text):
            obj_size_map[m.group(1).lower().replace(' ', '_')] = (m.group(2), m.group(3))
        if obj_size_map:
            if "small" in obj_size_map:
                chosen = obj_size_map["small"]
            elif "medium" in obj_size_map:
                chosen = obj_size_map["medium"]
            elif "large" in obj_size_map:
                chosen = obj_size_map["large"]
            elif "extra_large" in obj_size_map:
                chosen = obj_size_map["extra_large"]
            else:
                chosen = list(obj_size_map.values())[0]
            result["Width (inch)"] = chosen[0]
            result["Height (inch)"] = chosen[1]

    # ══════════════════════════════════════════════════════════════════════
    # ── RUG Pattern B: plain 2D  (NNN x NNN cm / W″ x H″)
    #    covers: "300 x 400 cm / 118″ x 157″"
    #         or "Standard dimensions:\n300 x 400 cm / 118″ x 157″"
    # ══════════════════════════════════════════════════════════════════════
    if not result["Width (inch)"]:
        rug_plain = re.search(
            r'[\d.]+\s*x\s*[\d.]+\s*cm\s*/\s*([\d.]+)[″"\'′]?\s*x\s*([\d.]+)',
            dim_text
        )
        if rug_plain:
            result["Width (inch)"] = rug_plain.group(1)
            result["Height (inch)"] = rug_plain.group(2)

    # ══════════════════════════════════════════════════════════════════════
    # ── 3-D patterns (W x D x H) — only if rug patterns found nothing ──
    # ══════════════════════════════════════════════════════════════════════
    if not result["Width (inch)"]:
        # ── Pattern 1: size-labelled W x D x H — split by label for correct priority ──
        size_map = {}
        for _lbl, _chunk in re.findall(
            r'\b(Small|Medium|Large)\s*:\s*(W\s*[\d.]+.+?)(?=\s*\b(?:Small|Medium|Large)\s*:|$)',
            dim_text, re.IGNORECASE
        ):
            _parts = _chunk.split('/')
            if len(_parts) > 1:
                _m = re.search(
                    r'W\s*([\d.]+)[\u2033\u201d"]*\s*x\s*(?:D\s*)?([\d.]+)[\u2033\u201d"]*\s*x\s*H\s*([\d.]+)',
                    _parts[1], re.IGNORECASE
                )
                if _m:
                    size_map[_lbl.lower()] = (_m.group(1), _m.group(2), _m.group(3))
        # also run original row_pattern for formats without /
        _row_pat = re.compile(
            r'(Small|Medium|Large)\s*:\s*'
            r'W\s*[\d.]+\s*x\s*D\s*[\d.]+\s*x\s*H\s*[\d.]+'
            r'\s*/\s*'
            r'W\s*([\d.]+)[\u2033\u201d"]*\s*x\s*(?:D\s*)?([\d.]+)[\u2033\u201d"]*\s*x\s*H\s*([\d.]+)',
            re.IGNORECASE
        )
        for m in _row_pat.finditer(dim_text):
            k = m.group(1).lower()
            if k not in size_map:
                size_map[k] = (m.group(2), m.group(3), m.group(4))

        if size_map:
            _pri = ['small', 'medium', 'large']
            chosen = next((size_map[k] for k in _pri if k in size_map), list(size_map.values())[0])
            result["Width (inch)"] = chosen[0]
            result["Depth (inch)"] = chosen[1]
            result["Height (inch)"] = chosen[2]

    # ── Pattern 2: plain row — no size label, cm / inch format ──
    if not result["Width (inch)"]:
        plain_slash = re.search(
            r'W\s*[\d.]+\s*x\s*D\s*[\d.]+\s*x\s*H\s*[\d.]+'
            r'\s*/\s*'
            r'W\s*([\d.]+)[″\'′"]*\s*x\s*D\s*([\d.]+)[″\'′"]*\s*x\s*H\s*([\d.]+)',
            dim_text
        )
        if plain_slash:
            result["Width (inch)"] = plain_slash.group(1)
            result["Depth (inch)"] = plain_slash.group(2)
            result["Height (inch)"] = plain_slash.group(3)

    # ── Pattern 3: inch line with ″ or " char ──
    if not result["Width (inch)"]:
        inch_line = re.search(
            r'W\s*([\d.]+)[″"]\s*x\s*D\s*([\d.]+)[″"]\s*x\s*H\s*([\d.]+)',
            dim_text
        )
        if inch_line:
            result["Width (inch)"] = inch_line.group(1)
            result["Depth (inch)"] = inch_line.group(2)
            result["Height (inch)"] = inch_line.group(3)

    # ── Pattern: plain W x H (no Depth): "W 47 x H 23 cm W 18.5″ x H 9″" ──
    if not result["Width (inch)"]:
        wh_plain = re.search(
            r'W\s*[\d.]+\s*x\s*H\s*[\d.]+\s*cm\s*W\s*([\d.]+)[\u2033\u2032\u201d"]*'
            r'\s*x\s*H\s*([\d.]+)',
            dim_text
        )
        if wh_plain:
            result["Width (inch)"] = wh_plain.group(1)
            result["Height (inch)"] = wh_plain.group(2)

    # ── Pattern: W x W x H (after L→W normalization): "W 25 x W 4.5 x H 33 cm W 9.9″ x W 1.8″ x H 13″"
    #    Take first W and last W before H as Width/Depth, H as Height ──
    if not result["Width (inch)"]:
        wwh = re.search(
            r'W\s*[\d.]+\s*x\s*W\s*[\d.]+\s*x\s*H\s*[\d.]+\s*cm'
            r'\s*W\s*([\d.]+)[\u2033\u2032\u201d"]*'
            r'\s*x\s*W\s*([\d.]+)[\u2033\u2032\u201d"]*'
            r'\s*x\s*H\s*([\d.]+)',
            dim_text
        )
        if wwh:
            result["Width (inch)"] = wwh.group(1)
            result["Depth (inch)"] = wwh.group(2)
            result["Height (inch)"] = wwh.group(3)

    # ── Pattern: W x H x W (L→W normalized L x H x W): "W 27.3 x H 38.4 x W 7.3 cm W 10.75″ x H 15.12″ x W 2.87″" ──
    if not result["Width (inch)"]:
        whw = re.search(
            r'W\s*[\d.]+\s*x\s*H\s*[\d.]+\s*x\s*W\s*[\d.]+\s*cm'
            r'\s*W\s*([\d.]+)[\u2033\u2032\u201d"]*'
            r'\s*x\s*H\s*([\d.]+)',
            dim_text
        )
        if whw:
            result["Width (inch)"] = whw.group(1)
            result["Height (inch)"] = whw.group(2)

    # ── Fallback: width looks like cm value (>100), re-parse after / ──
    if result["Width (inch)"] and float(result["Width (inch)"]) > 100:
        parts = dim_text.split("/")
        if len(parts) > 1:
            m2 = re.search(
                r'W\s*([\d.]+)[″\'′"]*\s*x\s*D\s*([\d.]+)[″\'′"]*\s*x\s*H\s*([\d.]+)',
                parts[1]
            )
            if m2:
                result["Width (inch)"] = m2.group(1)
                result["Depth (inch)"] = m2.group(2)
                result["Height (inch)"] = m2.group(3)

    # ── Diameter ──
    # NOTE: matches both Ø (U+00D8) and ∅ (U+2205) — site uses both characters
    # Range diameters like "Ø 1.9-2.7″" → take the first (lower) value
    _D = r'[Ø∅]'
    _RNG = r'(?:-[\d.]+)?'   # optional range suffix e.g. "-2.7"

    # ── Size-labelled ∅×H  (Small/Medium/Large: ∅ N-N x H N cm / ∅ N-N″ x H N″) ──
    # ── Size-labelled H-only  (Small/Medium/Large: H N cm / H N″) ──
    if not result["Height (inch)"]:
        h_size_map = {}
        h_size_pat = re.compile(
            r'(Small|Medium|Large|Extra\s+Large)\s*:\s*'
            r'H\s*[\d.]+\s*cm\s*/\s*H\s*([\d.]+)',
            re.IGNORECASE
        )
        for m in h_size_pat.finditer(dim_text):
            h_size_map[m.group(1).lower().replace(' ', '_')] = m.group(2)
        if h_size_map:
            if "small" in h_size_map:
                result["Height (inch)"] = h_size_map["small"]
            elif "medium" in h_size_map:
                result["Height (inch)"] = h_size_map["medium"]
            elif "large" in h_size_map:
                result["Height (inch)"] = h_size_map["large"]
            elif "extra_large" in h_size_map:
                result["Height (inch)"] = h_size_map["extra_large"]
            else:
                result["Height (inch)"] = list(h_size_map.values())[0]

    dh_size_map = {}
    dh_size_pat = re.compile(
        r'(Small|Medium|Large)\s*:\s*'
        r'[\u00d8\u2205]\s*[\d.]+(?:-[\d.]+)?\s*x\s*H\s*[\d.]+\s*cm'
        r'\s*/\s*'
        r'[\u00d8\u2205]\s*([\d.]+)(?:-[\d.]+)?[\u2033\u2032\u201d"\']*\s*x\s*H\s*([\d.]+)',
        re.IGNORECASE
    )
    for m in dh_size_pat.finditer(dim_text):
        dh_size_map[m.group(1).lower()] = (m.group(2), m.group(3))

    if dh_size_map:
        if "small" in dh_size_map:
            chosen = dh_size_map["small"]
        elif "medium" in dh_size_map:
            chosen = dh_size_map["medium"]
        elif "large" in dh_size_map:
            chosen = dh_size_map["large"]
        else:
            chosen = list(dh_size_map.values())[0]
        result["Diameter (inch)"] = chosen[0]
        result["Height (inch)"] = chosen[1]

    # ── Plain ∅ x H (no size label) — checked BEFORE diameter-only fallbacks ──
    # inch-symbol ″ guarantees inch value; wins over bare cm number (dm_bare)
    # e.g. "Ø 18 x H 3 cm Ø 7″ x H 1″" → Diameter=7, Height=1
    if not result["Diameter (inch)"]:
        dh_inch = re.search(
            _D + r'\s*([\d.]+)' + _RNG + r'[″"\'′]\s*x\s*H\s*([\d.]+)[″"\'′]?',
            dim_text
        )
        if dh_inch:
            result["Diameter (inch)"] = dh_inch.group(1)
            if not result["Height (inch)"]:
                result["Height (inch)"] = dh_inch.group(2)
        else:
            dh_slash = re.search(
                _D + r'\s*[\d.]+' + _RNG + r'\s*x\s*H\s*[\d.]+\s*cm\s*/\s*'
                + _D + r'\s*([\d.]+)' + _RNG + r'[″"\'′]?\s*x\s*H\s*([\d.]+)',
                dim_text
            )
            if dh_slash:
                result["Diameter (inch)"] = dh_slash.group(1)
                if not result["Height (inch)"]:
                    result["Height (inch)"] = dh_slash.group(2)

    # ── Plain ∅ (diameter only, no height) ──
    if not result["Diameter (inch)"]:
        dm_slash = re.search(
            _D + r'\s*[\d.]+' + _RNG + r'\s*cm\s*[/\n]\s*' + _D + r'\s*([\d.]+)' + _RNG + r'[″"\'′]?',
            dim_text
        )
        if dm_slash:
            result["Diameter (inch)"] = dm_slash.group(1)
        else:
            dm_named = re.search(
                r'(?:Dia(?:meter)?)[:\s]*([\d.]+)\s*cm\s*/\s*([\d.]+)',
                dim_text
            )
            if dm_named:
                result["Diameter (inch)"] = dm_named.group(2)
            else:
                dm_standalone = re.search(_D + r'\s*([\d.]+)' + _RNG + r'[″"\'′]', dim_text)
                if dm_standalone:
                    result["Diameter (inch)"] = dm_standalone.group(1)
                else:
                    if not re.search(_D + r'\s*[\d.]+\s*cm', dim_text):
                        dm_bare = re.search(_D + r'\s*([\d.]+)', dim_text)
                        if dm_bare:
                            result["Diameter (inch)"] = dm_bare.group(1)


    # ── Plain H-only (no size label): "H 22 cm H 8.7″" ──
    if not result["Height (inch)"]:
        h_plain = re.search(r'\bH\s*[\d.]+\s*cm\s+H\s*([\d.]+)', dim_text, re.IGNORECASE)
        if h_plain:
            result["Height (inch)"] = h_plain.group(1)

    # ── Size-labelled Ø-only (no H): "Small: Ø 6.2 cm / Ø 2.4″" ──
    if not result["Diameter (inch)"]:
        d_only_size_map = {}
        d_only_pat = re.compile(
            r'(Small|Medium|Large|Extra\s+Large)\s*:\s*'
            r'[\u00d8\u2205]\s*[\d.]+(?:-[\d.]+)?\s*cm\s*/?\s*(?:[\u00d8\u2205]\s*)?(\d[\d.]*)',
            re.IGNORECASE
        )
        for m in d_only_pat.finditer(dim_text):
            d_only_size_map[m.group(1).lower().replace(' ', '_')] = m.group(2)
        if d_only_size_map:
            _d_pri = ['small', 'medium', 'large', 'extra_large']
            result["Diameter (inch)"] = next(
                (d_only_size_map[k] for k in _d_pri if k in d_only_size_map),
                list(d_only_size_map.values())[0]
            )

    # ── Seat Height ──
    sm = re.search(r'[Ss]eat(?:ing)?\s*height[:\s]*([\d.]+)\s*cm\s*/\s*([\d.]+)', dim_text)
    if sm:
        result["Seat Height (inch)"] = sm.group(2)

    # ── Seat Depth ──
    sd = re.search(r'[Ss]eat(?:ing)?\s*depth[:\s]*([\d.]+)\s*cm\s*/\s*([\d.]+)', dim_text)
    if sd:
        result["Seat Depth (inch)"] = sd.group(2)

    # ── Fabric Required ──
    fm = re.search(r'[Ff]abric\s*required[:\s]*([\d.]+)\s*m\s*/\s*([\d.]+)\s*yd', dim_text)
    if fm:
        result["Fabric Required (yd)"] = fm.group(2)

    return result


# Lines to skip when extracting finish from a single cst-str paragraph
_FINISH_SKIP = re.compile(
    r'^(?:designed\s+by|available\s+in|bespoke|please\s+discover|the\s+images?\s+showcase)',
    re.IGNORECASE
)


def get_finish_from_page(driver):
    """
    Extract finish from .cst-str second <p> tag first.
    If only one <p>, each <br>-separated line is checked; lines starting with
    'Designed by', 'Available in', 'Bespoke', 'Please discover', or
    'The images showcased' are skipped — the first remaining line is the finish.
    Falls back to select/variable-item options if not found.
    """
    try:
        cst = driver.find_element(By.CSS_SELECTOR, ".pro-extra-item.cst-str")
        paras = cst.find_elements(By.TAG_NAME, "p")
        if len(paras) >= 2:
            finish_text = clean_text(paras[1].text)
            if finish_text:
                return finish_text
        elif len(paras) == 1:
            # paragraph may contain multiple <br>-separated lines
            raw = paras[0].text  # Selenium converts <br> → \n
            for line in raw.splitlines():
                line = line.strip()
                if line and not _FINISH_SKIP.match(line):
                    return line
    except:
        pass

    finishes = []
    try:
        selects = driver.find_elements(By.CSS_SELECTOR, "select[data-attribute_name]")
        for sel in selects:
            attr_name = (sel.get_attribute("data-attribute_name") or "").lower()
            if any(k in attr_name for k in ["finish", "fabric", "wood", "material", "color",
                                              "marble", "stone", "leather", "metal", "veneer"]):
                options = sel.find_elements(By.TAG_NAME, "option")
                for opt in options:
                    val = opt.text.strip()
                    if val and val.lower() not in ("choose an option", ""):
                        if "customer" not in val.lower() and "com" not in val.lower().split("-")[0].strip():
                            finishes.append(val)
    except:
        pass

    if not finishes:
        try:
            items = driver.find_elements(By.CSS_SELECTOR, "li.variable-item[data-title]")
            for item in items:
                title = (item.get_attribute("data-title") or "").strip()
                if title and "customer" not in title.lower() and title.upper() != "COM - CUSTOMER'S OWN FABRIC":
                    finishes.append(title)
        except:
            pass

    return " | ".join(finishes)


def parse_light_info(driver):
    """Extract Colour Temperature, Light Source, Socket from .light-info section"""
    result = {
        "Colour Temperature": "",
        "Light Source": "",
        "Socket": "",
    }
    try:
        light_text = safe_text(driver, ".pro-extra-item.light-info")
        if not light_text:
            items = driver.find_elements(By.CSS_SELECTOR, ".pro-extra-item")
            for item in items:
                txt = clean_text(item.text)
                if "colour temperature" in txt.lower() or "light source" in txt.lower() or "socket" in txt.lower():
                    light_text = txt
                    break

        if not light_text:
            return result

        ct = re.search(r'Colou?r\s*Temperature\s*:\s*(.+?)(?:\n|Light Source|Socket|Dimmable|IP\s|$)', light_text, re.IGNORECASE)
        if ct:
            result["Colour Temperature"] = ct.group(1).strip().rstrip('.')
        else:
            ct2 = re.search(r'Colou?r\s*Temperature\s*:\s*(.+?)(?=\s*(?:Light Source|Socket|Dimmable|IP\s))', light_text, re.IGNORECASE)
            if ct2:
                result["Colour Temperature"] = ct2.group(1).strip().rstrip('.')

        ls = re.search(r'Light\s*Source\s*:\s*(.+?)(?:\n|Socket|Colou?r|Dimmable|IP\s|$)', light_text, re.IGNORECASE)
        if ls:
            val = ls.group(1).strip().rstrip('.')
            if val.lower() not in ("yes", "no"):
                result["Light Source"] = val
        if not result["Light Source"]:
            ls_all = re.findall(r'Light\s*Source\s*:\s*(.+?)(?=\s*(?:Socket|Colou?r|Dimmable|IP\s|Light Source Included|$))', light_text, re.IGNORECASE)
            for val in ls_all:
                val = val.strip().rstrip('.')
                if val.lower() not in ("yes", "no"):
                    result["Light Source"] = val
                    break

        sk = re.search(r'Socket\s*:\s*(.+?)(?:\n|Colou?r|Light Source|Dimmable|IP\s|$)', light_text, re.IGNORECASE)
        if sk:
            result["Socket"] = sk.group(1).strip().rstrip('.')

    except Exception as e:
        print(f"    Light info parse error: {e}")

    return result


def wait_for_download(timeout=30):
    """Wait for PDF to finish downloading"""
    end_time = time.time() + timeout
    while time.time() < end_time:
        files = glob.glob(os.path.join(DOWNLOAD_FOLDER, "*"))
        downloading = [f for f in files if f.endswith(".crdownload") or f.endswith(".tmp")]
        done_files = [f for f in files if f.lower().endswith(".pdf") or f.lower().endswith(".par")]
        if done_files and not downloading:
            done_files.sort(key=os.path.getmtime, reverse=True)
            return done_files[0]
        time.sleep(1)
    return None


def download_tearsheet(driver, product_name):
    """
    Tearsheet download flow:
    1. Click "Download Tearsheet" button
    2. Popup opens with DOWNLOAD button
    3. Click DOWNLOAD -> PDF downloads
    """
    tearsheet_path = ""
    try:
        for f in glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")):
            try:
                os.remove(f)
            except:
                pass

        tearsheet_btn = None
        for sel in ["a#tic_tear_sheet", "a.open-tearsheet-popup", "a[href='#tearsheet_popup']"]:
            try:
                tearsheet_btn = driver.find_element(By.CSS_SELECTOR, sel)
                break
            except:
                continue

        if not tearsheet_btn:
            try:
                all_links = driver.find_elements(By.TAG_NAME, "a")
                for link in all_links:
                    if "download tearsheet" in (link.text or "").lower():
                        tearsheet_btn = link
                        break
            except:
                pass

        if not tearsheet_btn:
            return ""

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tearsheet_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", tearsheet_btn)
        time.sleep(3)

        popup_found = False
        for sel in ["#tearsheet_popup", "[id*='tearsheet']", ".fancybox-content",
                     ".fancybox-slide--current", ".popup-tearsheet"]:
            try:
                WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, sel))
                )
                popup_found = True
                print(f"    Popup found: {sel}")
                break
            except:
                continue

        if not popup_found:
            time.sleep(2)
            try:
                body_text = driver.find_element(By.TAG_NAME, "body").text
                if "download the tearsheet" in body_text.lower():
                    popup_found = True
                    print(f"    Popup detected via text")
            except:
                pass

        if not popup_found:
            print(f"    Tearsheet: popup did not appear")
            return ""

        download_btn = None
        popup_containers = [
            "#tearsheet_popup",
            ".fancybox-content",
            ".fancybox-slide--current",
            "[id*='tearsheet']",
        ]

        for container_sel in popup_containers:
            try:
                container = driver.find_element(By.CSS_SELECTOR, container_sel)
                elements = container.find_elements(By.CSS_SELECTOR, "a, button, input[type='submit']")
                for el in elements:
                    el_text = (el.text or el.get_attribute("value") or "").strip().upper()
                    if el_text == "DOWNLOAD" and el.is_displayed():
                        download_btn = el
                        print(f"    Found DOWNLOAD btn in {container_sel}")
                        break
                if download_btn:
                    break
            except:
                continue

        if not download_btn:
            try:
                xpaths = [
                    "//a[normalize-space(text())='Download' or normalize-space(text())='DOWNLOAD']",
                    "//button[normalize-space(text())='Download' or normalize-space(text())='DOWNLOAD']",
                    "//a[contains(@class,'download')]",
                    "//button[contains(@class,'download')]",
                    "//a[contains(@id,'download')]",
                    "//input[@value='Download' or @value='DOWNLOAD']",
                ]
                for xp in xpaths:
                    try:
                        elements = driver.find_elements(By.XPATH, xp)
                        for el in elements:
                            if el.is_displayed():
                                el_text = (el.text or "").strip().upper()
                                if el_text != "DOWNLOAD TEARSHEET" and el != tearsheet_btn:
                                    download_btn = el
                                    print(f"    Found DOWNLOAD btn via XPath")
                                    break
                        if download_btn:
                            break
                    except:
                        continue
            except:
                pass

        if not download_btn:
            try:
                for sel in [
                    ".tearsheet-download", "#tearsheet_download",
                    ".popup-download-btn", ".tear-download",
                    "a.download-btn", "button.download-btn",
                    ".fancybox-content a.btn", ".fancybox-content button",
                    "#tearsheet_popup a", "#tearsheet_popup button",
                ]:
                    try:
                        el = driver.find_element(By.CSS_SELECTOR, sel)
                        if el.is_displayed() and el != tearsheet_btn:
                            download_btn = el
                            print(f"    Found DOWNLOAD btn: {sel}")
                            break
                    except:
                        continue
            except:
                pass

        if download_btn:
            driver.execute_script("arguments[0].click();", download_btn)
            print(f"    Clicked DOWNLOAD, waiting for PDF...")
            time.sleep(5)

            downloaded = wait_for_download(timeout=30)
            if downloaded:
                safe_name = re.sub(r'[^\w\s-]', '', product_name).strip().replace(' ', '_')[:80]
                dest = os.path.join(TEARSHEET_FOLDER, f"{safe_name}_tearsheet.pdf")
                counter = 1
                while os.path.exists(dest):
                    dest = os.path.join(TEARSHEET_FOLDER, f"{safe_name}_tearsheet_{counter}.pdf")
                    counter += 1
                shutil.move(downloaded, dest)
                tearsheet_path = dest
                print(f"    Tearsheet SAVED: {os.path.basename(dest)}")
            else:
                print(f"    Tearsheet: download timeout / no file found")
        else:
            print(f"    Tearsheet: DOWNLOAD button not found in popup")

        try:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(0.5)
        except:
            pass
        try:
            for sel in [".fancybox-close-small", ".fancybox-close", ".popup-close", "button.close"]:
                try:
                    close = driver.find_element(By.CSS_SELECTOR, sel)
                    if close.is_displayed():
                        close.click()
                        break
                except:
                    continue
        except:
            pass
        time.sleep(1)

    except Exception as e:
        print(f"    Tearsheet error: {e}")

    return tearsheet_path


def scrape_product(driver, url, index):
    data = {
        "Product URL": url, "SKU": "", "Product Name": "", "Designer": "",
        "Description": "", "Dimension (Full)": "",
        "Width (inch)": "", "Depth (inch)": "", "Height (inch)": "",
        "Diameter (inch)": "", "Seat Height (inch)": "", "Seat Depth (inch)": "",
        "Fabric Required (yd)": "", "Production Lead Time": "",
        "Finish": "", "Price": "",
        "Colour Temperature": "", "Light Source": "", "Socket": "",
        "Tearsheet": "",
    }

    try:
        driver.get(url)
        time.sleep(random.uniform(2, 4))

        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1")))
        except:
            print(f"    Page load timeout")
            return data

        # Product Name
        for sel in ["h1.product_title", "h1.pro-title", "h1"]:
            txt = safe_text(driver, sel)
            if txt:
                data["Product Name"] = txt
                break

        if not data["Product Name"]:
            return data

        # SKU
        page_sku = safe_text(driver, ".sku")
        data["SKU"] = page_sku.strip() if page_sku.strip() else generate_sku(VENDOR_NAME, CATEGORY_NAME, index)

        # Designer
        for sel in [".pro-designer-name a", ".pro-designer-name", "a[href*='/designer/']"]:
            txt = safe_text(driver, sel)
            if txt:
                data["Designer"] = txt
                break

        # Description
        for sel in [".pro-content-txt .pro-content-area", ".pro-content-txt"]:
            txt = safe_text(driver, sel)
            if txt:
                txt = re.sub(r'^Photo\s*credit\s*[–-]\s*\w[\w\s]*?\s+(?=[A-Z])', '', txt).strip()
                data["Description"] = txt
                break

        # Price
        data["Price"] = safe_text(driver, ".price .amount") or safe_text(driver, ".woocommerce-Price-amount")

        # ── Dimensions ────────────────────────────────────────────────────────
        dim_text = ""
        try:
            outer = driver.find_element(By.CSS_SELECTOR, ".all-var-outer")
            dim_text = clean_text(outer.text)
        except:
            pass

        if not dim_text or not re.search(r'\d', dim_text):
            for sel in [".pro-extra-item p", ".pro-extra-item", ".all-dc-data"]:
                txt = safe_text(driver, sel)
                if txt and re.search(r'\d', txt):
                    dim_text = txt
                    break

        # ── Also grab ALL .pro-extra-item blocks to catch rug/object dimension div ──
        # Covers: "300 x 400 cm", "W 60 x D 60 x H 60 cm", "∅ 31 x H 14 cm"
        _has_dim = re.compile(
            r'(?:[\d.]+\s*x\s*(?:[A-Z]\s*)?[\d.]'  # N x N  or  N x H N
            r'|[Ø∅]\s*[\d.]'                          # ∅ N
            r'|\bH\s*[\d.]+\s*cm)',                  # H-only: e.g. H 52 cm
            re.IGNORECASE
        )
        if not dim_text or not _has_dim.search(dim_text):
            try:
                all_extra = driver.find_elements(By.CSS_SELECTOR, ".pro-extra-item")
                for item in all_extra:
                    txt = clean_text(item.text)
                    if _has_dim.search(txt) and 'cm' in txt.lower():
                        dim_text = txt
                        break
            except:
                pass

        if dim_text:
            data["Dimension (Full)"] = dim_text
            dims = parse_dimensions(dim_text)
            for k, v in dims.items():
                data[k] = v

        # Lead Time
        lt_text = safe_text(driver, ".pro-lead-time .pro-lt") or safe_text(driver, ".pro-lead-time")
        if lt_text:
            lm = re.search(r'(\d+\s*weeks?)', lt_text, re.IGNORECASE)
            data["Production Lead Time"] = lm.group(1) if lm else lt_text.replace("Production lead time:", "").strip()

        # ── Finish ──
        data["Finish"] = get_finish_from_page(driver)

        # ── Light Info ────────────────────────────────────────────────────────
        light_data = parse_light_info(driver)
        data["Colour Temperature"] = light_data["Colour Temperature"]
        data["Light Source"] = light_data["Light Source"]
        data["Socket"] = light_data["Socket"]

        # Download Tearsheet
        data["Tearsheet"] = download_tearsheet(driver, data["Product Name"])

    except Exception as e:
        print(f"    ERROR: {e}")

    return data


# ══════════════════════════════════════════════════════════════════════════════
# ── AUTO-SAVE helpers ─────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

OUTPUT_COLS = [
    "Product URL", "SKU", "Product Name", "Designer", "Description",
    "Dimension (Full)", "Width (inch)", "Depth (inch)", "Height (inch)",
    "Diameter (inch)", "Seat Height (inch)", "Seat Depth (inch)",
    "Fabric Required (yd)", "Production Lead Time", "Finish", "Price",
    "Colour Temperature", "Light Source", "Socket",
    "Tearsheet"
]


def load_existing_results():
    """
    Load already-saved results from OUTPUT_FILE.
    Returns (list_of_dicts, set_of_done_urls).
    """
    if not os.path.exists(OUTPUT_FILE):
        return [], set()
    try:
        df = pd.read_excel(OUTPUT_FILE, engine="openpyxl")
        records = df.to_dict(orient="records")
        done_urls = {str(r.get("Product URL", "")).strip() for r in records if r.get("Product URL")}
        print(f"  Resume: found {len(done_urls)} already-scraped URLs in '{OUTPUT_FILE}'")
        return records, done_urls
    except Exception as e:
        print(f"  Warning: could not read existing output ({e}), starting fresh.")
        return [], set()


def save_results(results):
    """Save current results list to OUTPUT_FILE."""
    df_out = pd.DataFrame(results)
    for c in OUTPUT_COLS:
        if c not in df_out.columns:
            df_out[c] = ""
    df_out = df_out[OUTPUT_COLS]
    df_out.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")


# ══════════════════════════════════════════════════════════════════════════════


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"ERROR: '{INPUT_FILE}' not found!")
        print(f"Current folder: {os.getcwd()}")
        return

    df_input = pd.read_excel(INPUT_FILE)
    print(f"Loaded {len(df_input)} rows")

    url_col = None
    for col in df_input.columns:
        if "url" in col.lower() or "link" in col.lower():
            url_col = col
            break
    if not url_col:
        url_col = df_input.columns[0]
    print(f"URL column: '{url_col}'")

    urls = df_input[url_col].dropna().tolist()
    print(f"URLs: {len(urls)}")
    print(f"Tearsheets folder: {TEARSHEET_FOLDER}\n")

    # ── Resume: load already-done results ──
    results, done_urls = load_existing_results()

    pending_urls = [str(u).strip() for u in urls
                    if str(u).strip().startswith("http") and str(u).strip() not in done_urls]
    skipped = len(urls) - len(pending_urls)
    print(f"Pending: {len(pending_urls)} URLs  |  Skipped (already done): {skipped}\n")

    if not pending_urls:
        print("All URLs already scraped. Nothing to do.")
        return

    print("Starting Chrome...")
    driver = get_driver()

    print("Visiting homepage...")
    try:
        driver.get("https://theinvisiblecollection.com/")
        time.sleep(4)
    except:
        pass

    success = 0

    try:
        for i, url in enumerate(pending_urls, 1):
            global_index = len(results) + 1  # for SKU numbering
            print(f"\n[{i}/{len(pending_urls)}] {url}")
            data = scrape_product(driver, url, index=global_index)

            if data["Product Name"] and data["Product Name"] not in ("", "FAILED"):
                success += 1
                print(f"    OK: {data['Product Name']} | SKU: {data['SKU']}")
                if data["Width (inch)"]:
                    print(f"    Dims: W={data['Width (inch)']}  D={data['Depth (inch)']}  H={data['Height (inch)']}")
                if data["Diameter (inch)"]:
                    print(f"    Diameter={data['Diameter (inch)']}")
                if data["Seat Height (inch)"]:
                    print(f"    Seat H={data['Seat Height (inch)']}  Seat D={data['Seat Depth (inch)']}")
                if data["Colour Temperature"]:
                    print(f"    Light: {data['Light Source']} | Socket: {data['Socket']} | Temp: {data['Colour Temperature']}")
                if data["Finish"]:
                    print(f"    Finish: {data['Finish']}")
            else:
                print(f"    FAILED")

            results.append(data)

            # ── AUTO-SAVE after every product ──
            try:
                save_results(results)
                print(f"    Auto-saved → '{OUTPUT_FILE}'")
            except Exception as save_err:
                print(f"    Auto-save failed: {save_err}")

            time.sleep(random.uniform(2, 5))

    finally:
        driver.quit()
        print("\nBrowser closed.")

    # Final save
    save_results(results)

    print(f"\n{'='*50}")
    print(f"SUCCESS this run: {success}/{len(pending_urls)} products")
    print(f"Total in output : {len(results)}")
    print(f"Output: '{OUTPUT_FILE}'")
    ts_count = len(glob.glob(os.path.join(TEARSHEET_FOLDER, "*.pdf")))
    print(f"Tearsheets: {ts_count} PDFs")


if __name__ == "__main__":
    main()