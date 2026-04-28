import re


def parse_dimensions(dim_str):
    """
    Parse dimension strings into Width, Depth, Height, Diameter (all in inches).

    Handles patterns like:
      18 H x 88 W x 32.5 D
      18 H x 50 Square
      11.5 H x 44 Diam
      35Dia x 17.5H
      42 W x 41 L x 14H
      20.5 H x 27.75 D x 49.25 L
      Adjustable height: 30.5 H - 22H x 4...
    """
    result = {"Width": "", "Depth": "", "Height": "", "Diameter": ""}

    if not dim_str or str(dim_str).strip() == "":
        return result

    text = str(dim_str).strip()
    # normalize: remove inch symbols, fix apostrophe-as-inch, lowercase
    text = text.replace('"', '').replace("'", '').lower()
    # remove "adjustable height:" prefix — take the larger H value
    text = re.sub(r'adjustable\s*height\s*:', '', text)

    # extract all tokens like: 18h  88w  32.5d  44diam  50square  41l
    # token = number immediately followed by (or preceded by) a letter code
    # pattern: optional spaces between number and letter
    token_pattern = re.compile(
        r'(\d+(?:\.\d+)?)\s*(h|w|d|l|diam|dia|square)\b'
        r'|'
        r'\b(diam|dia|square)\s*(\d+(?:\.\d+)?)',
        re.IGNORECASE
    )

    found = {}  # key: dimension code, value: float

    for m in token_pattern.finditer(text):
        if m.group(1) is not None:
            num = float(m.group(1))
            code = m.group(2).lower()
        else:
            num = float(m.group(4))
            code = m.group(3).lower()

        if code == 'h':
            # if multiple H values (adjustable), take the larger one
            if 'h' not in found or num > found['h']:
                found['h'] = num
        elif code == 'w':
            found['w'] = num
        elif code in ('d',):
            found['d'] = num
        elif code == 'l':
            # L = Length → maps to Width column
            found['l'] = num
        elif code in ('diam', 'dia'):
            found['diam'] = num
        elif code == 'square':
            found['square'] = num

    # --- assign to output ---
    result["Height"] = _fmt(found.get('h'))

    if 'diam' in found:
        result["Diameter"] = _fmt(found['diam'])
        # square & diam: no W/D
    elif 'square' in found:
        result["Width"] = _fmt(found['square'])
        result["Depth"] = _fmt(found['square'])
    else:
        # L = Length → Width; if both L and W exist: L→Width, W→Depth
        if 'l' in found and 'w' in found:
            result["Width"] = _fmt(found['l'])
            result["Depth"] = _fmt(found.get('d') or found['w'])
        elif 'l' in found:
            result["Width"] = _fmt(found['l'])
            if 'd' in found:
                result["Depth"] = _fmt(found['d'])
        else:
            if 'w' in found:
                result["Width"] = _fmt(found['w'])
            if 'd' in found:
                result["Depth"] = _fmt(found['d'])

    return result


def _fmt(val):
    if val is None:
        return ""
    # show as integer if whole number, else 2 decimal places
    return str(int(val)) if val == int(val) else f"{val:.2f}"


# ── test ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    test_cases = [
        "18 H x 88 W x 32.5 D",
        "18 H x 50 Square",
        "16 H x 70 W x 38 D",
        "11.5 H x 44 Diam",
        "14.75 H x 41 W x 42 D",
        "42 W x 41 L x 14H",
        "19 H x 42 W x 28.25D",
        "15.5 H x 51 W x 45 D",
        "18 H x 60 W x 36 D",
        "22 H x 38 W x 40 D",
        "13 H x 41 W x 36 D",
        "14.25 H x 57.5 Diam",
        "20.5 H x 27.75 D x 49.25 L",
        "16.5 H x 37 W x 34 D",
        "21.5H x 47W x 23.5D",
        "20 H x 33 W x 25 D",
        "18 H x 35 D",
        "25H x 82W x 36D",
        "19 H x 69 W x 66 D",
        "17 H x 40 Square",
        "17.5 H x 55 Dia",
        "35Dia x 17.5H",
        "14.5 H x 55 Dia",
        "Adjustable height: 30.5 H - 22H x 48W x 24D",
        "10H x 67W x 34D",
    ]

    print(f"{'Input':<45} {'H':>6} {'W':>6} {'D':>6} {'Diam':>6}")
    print("-" * 75)
    for t in test_cases:
        r = parse_dimensions(t)
        print(f"{t:<45} {r['Height']:>6} {r['Width']:>6} {r['Depth']:>6} {r['Diameter']:>6}")
