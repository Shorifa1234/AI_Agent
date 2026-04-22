import re
import pandas as pd

# -------- CONFIG --------
INPUT_XLSX  = "paulferrante_step2.xlsx"
OUTPUT_XLSX = "paulferrante_step2_full_with_all.xlsx"

# -------- REGEX --------
NUM  = r'(?:\d+\s*\d/\d|\d+\.\d+|\d+)'      # 12, 12.5, 12 1/2
INCH = r'(?:\s*(?:["”]|in(?:ches)?|inch)?)' # ", in, inch
SEP  = r'[\s]*'

PAT_WIDTH   = re.compile(rf'(?:width[:\-]?\s*|)(?P<val>{NUM}){INCH}{SEP}(?:w(?![a-z])|width)', re.I)
PAT_HEIGHT  = re.compile(rf'(?:height[:\-]?\s*|)(?P<val>{NUM}){INCH}{SEP}(?:h(?![a-z])|height|high)', re.I)
PAT_DIAM    = re.compile(rf'(?P<val>{NUM}){INCH}{SEP}(?:dia(?:meter)?|Ø|d(?![a-z]))', re.I)
PAT_DEPTH   = re.compile(rf'(?:depth[:\-]?\s*|)(?P<val>{NUM}){INCH}{SEP}(?:d(?![a-z])|depth)', re.I)  # New regex for Depth

# Other field patterns
PATTERNS = {
    "Finish": re.compile(r'finish[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Socket": re.compile(r'socket[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Wattage": re.compile(r'wattage[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Lightsource": re.compile(r'light\s*source[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Color Temperature": re.compile(r'color\s*temp(?:erature)?[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Extension": re.compile(r'extension[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Rating": re.compile(r'rating[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Shade Details": re.compile(r'(?:shade\s*details?|shade)[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Base": re.compile(r'base[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Canopy": re.compile(r'canopy[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
    "Chain Length": re.compile(r'chain(?:\s*length)?[:\-]?\s*([A-Za-z0-9\s\-\(\)/,&\.]+)', re.I),
}

def clean_num(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r'["”]\s*', '', s)
    s = re.sub(r'\b(?:inches?|inch|in)\b', '', s, flags=re.I)
    return s.strip()

def extract_field(pattern, text):
    m = pattern.search(text)
    if not m:
        return ""
    return m.group(1).strip()

def extract_whd(text: str):
    """Extract Width, Height, Diameter, and Depth."""
    if not isinstance(text, str):
        return "", "", "", ""
    t = text.strip()
    if not t:
        return "", "", "", ""

    w = h = d = dia = ""
    m1 = PAT_WIDTH.search(t)
    m2 = PAT_HEIGHT.search(t)
    m3 = PAT_DIAM.search(t)
    m4 = PAT_DEPTH.search(t)

    if m1: w = clean_num(m1.group('val'))
    if m2: h = clean_num(m2.group('val'))
    if m3: dia = clean_num(m3.group('val'))
    if m4: d = clean_num(m4.group('val'))

    return w, h, dia, d

def extract_all(dim_text: str):
    """Extract all fields from one dimension string."""
    if not isinstance(dim_text, str):
        dim_text = ""
    t = dim_text.strip()

    w, h, dia, d = extract_whd(t)
    results = {
        "Width": w,
        "Height": h,
        "Diameter": dia,
        "Depth": d,  # Add Depth to the results
    }

    # Other text-based fields
    for field, pat in PATTERNS.items():
        results[field] = extract_field(pat, t)

    return results

def main():
    df = pd.read_excel(INPUT_XLSX)

    if 'Dimensions' not in df.columns:
        raise ValueError("Input file must contain a 'Dimensions' column.")

    extracted_rows = []
    for dim in df['Dimensions']:
        extracted_rows.append(extract_all(dim))

    extra_df = pd.DataFrame(extracted_rows)

    merged = pd.concat([df, extra_df], axis=1)
    merged.to_excel(OUTPUT_XLSX, index=False)

    print(f"✅ Done! Extracted Width/Height/Diameter/Depth + all fields → {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
