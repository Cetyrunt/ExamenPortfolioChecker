import pandas as pd

def clean_val(val):
    """Clean a single Excel cell and format dates to EU format (DD-MM-YYYY)."""
    if pd.isna(val):
        return "---"

    if isinstance(val, pd.Timestamp):
        if val.year == 2000:
            return "---"
        return val.strftime("%d-%m-%Y")

    s = str(val).strip()
    if s.startswith("2000-01-01") or s.startswith("1899-12-30"):
        return "---"

    # Convert ISO-like strings to EU date
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        try:
            y, m, d = s[:10].split("-")
            return f"{d}-{m}-{y}"
        except Exception:
            pass

    if "dropdown" in s.lower() or s == "0":
        return "---"

    return s
