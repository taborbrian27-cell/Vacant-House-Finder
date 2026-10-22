import pandas as pd
import re
from pathlib import Path

FOLDER = Path(__file__).resolve().parent

HOUSE_FILE = FOLDER / "Prince_Georges_Vacancy_Leads_HOUSE_VERIFIED.xlsx"
REDFIN_FILE = FOLDER / "Redfin_All_Downloads_Combined.xlsx"
OUTPUT_FILE = FOLDER / "Prince_Georges_Houses_For_Sale_FINAL.xlsx"

def clean_col(name):
    return re.sub(r"[^a-z0-9]+", " ", str(name).strip().lower()).strip()

def find_col(df, candidates):
    lookup = {clean_col(c): c for c in df.columns}
    for candidate in candidates:
        key = clean_col(candidate)
        if key in lookup:
            return lookup[key]
    # Also allow partial matches.
    for c in df.columns:
        cc = clean_col(c)
        for candidate in candidates:
            if clean_col(candidate) in cc:
                return c
    return None

def normalize_address(value):
    if pd.isna(value):
        return ""
    s = str(value).upper().strip()
    s = s.replace("&", " AND ")
    s = re.sub(r"[.,#'’]", " ", s)
    s = re.sub(r"\s+", " ", s)

    # Normalize common street suffixes and directional abbreviations.
    replacements = {
        r"\bSTREET\b": "ST",
        r"\bST\b": "ST",
        r"\bROAD\b": "RD",
        r"\bRD\b": "RD",
        r"\bAVENUE\b": "AVE",
        r"\bAV\b": "AVE",
        r"\bDRIVE\b": "DR",
        r"\bDR\b": "DR",
        r"\bLANE\b": "LN",
        r"\bLN\b": "LN",
        r"\bCOURT\b": "CT",
        r"\bCT\b": "CT",
        r"\bCIRCLE\b": "CIR",
        r"\bCIR\b": "CIR",
        r"\bBOULEVARD\b": "BLVD",
        r"\bBLVD\b": "BLVD",
        r"\bHIGHWAY\b": "HWY",
        r"\bHWY\b": "HWY",
        r"\bPARKWAY\b": "PKWY",
        r"\bPKWY\b": "PKWY",
        r"\bPLACE\b": "PL",
        r"\bPL\b": "PL",
        r"\bTERRACE\b": "TER",
        r"\bTER\b": "TER",
        r"\bTRAIL\b": "TRL",
        r"\bTRL\b": "TRL",
        r"\bNORTH\b": "N",
        r"\bSOUTH\b": "S",
        r"\bEAST\b": "E",
        r"\bWEST\b": "W",
    }
    for pattern, replacement in replacements.items():
        s = re.sub(pattern, replacement, s)
    return s.strip()

print("Loading House-Verified properties...")
houses = pd.read_excel(HOUSE_FILE)

print("Loading combined Redfin listings...")
redfin = pd.read_excel(REDFIN_FILE)

house_address_col = find_col(houses, ["Address", "Property Address", "Street Address"])
redfin_address_col = find_col(redfin, ["Property Address", "Address", "Street Address"])
redfin_status_col = find_col(redfin, ["Status", "Listing Status"])
redfin_city_col = find_col(redfin, ["City"])
redfin_state_col = find_col(redfin, ["State Or Province", "State", "State/Province"])
redfin_zip_col = find_col(redfin, ["Zip Or Postal Code", "Zip Code", "ZIP", "Postal Code"])

if not house_address_col:
    raise ValueError("Could not find Address in the House-Verified workbook.")
if not redfin_address_col:
    raise ValueError("Could not find Property Address in the Redfin workbook.")

print(f"House address column: {house_address_col}")
print(f"Redfin address column: {redfin_address_col}")

# Keep only Redfin listings that are explicitly Active when a status column exists.
if redfin_status_col:
    status = redfin[redfin_status_col].fillna("").astype(str).str.strip().str.upper()
    active = redfin[status.eq("ACTIVE")].copy()
    print(f"Redfin rows before Active filter: {len(redfin):,}")
    print(f"Redfin Active rows: {len(active):,}")
else:
    active = redfin.copy()
    print("No Redfin Status column found; all combined Redfin rows will be used.")

# Build normalized street-address keys.
houses["_match_address"] = houses[house_address_col].map(normalize_address)
active["_match_address"] = active[redfin_address_col].map(normalize_address)

# Remove blank address keys.
houses = houses[houses["_match_address"] != ""].copy()
active = active[active["_match_address"] != ""].copy()

# If multiple Redfin records have the same address, keep one active listing.
active = active.drop_duplicates(subset=["_match_address"], keep="first")

# Match the Redfin addresses to House-Verified SDAT addresses.
matched = houses.merge(
    active,
    on="_match_address",
    how="inner",
    suffixes=("", " - Redfin")
)

print(f"House-Verified properties: {len(houses):,}")
print(f"Matched House-Verified + Redfin properties: {len(matched):,}")

# Remove helper columns and place important Redfin fields near the front.
matched = matched.drop(columns=["_match_address"], errors="ignore")

preferred = [
    "Account ID",
    house_address_col,
    "House Verified",
    "Property Condition",
    "Boarded Up Windows",
    "Overgrown Vegetation",
    "Vacancy Score",
    "Priority",
    redfin_status_col,
    find_col(active, ["Price", "List Price", "Listing Price"]),
    find_col(active, ["Beds", "Bedrooms"]),
    find_col(active, ["Baths", "Bathrooms"]),
    find_col(active, ["Square Feet", "Sq Ft", "Square Footage"]),
    find_col(active, ["Year Built"]),
    find_col(active, ["Days On Redfin", "Days On Market"]),
    find_col(active, ["URL", "Listing URL"]),
]

preferred = [c for c in preferred if c and c in matched.columns]
remaining = [c for c in matched.columns if c not in preferred]
matched = matched[preferred + remaining]

matched.to_excel(OUTPUT_FILE, index=False)

print()
print("==============================================")
print("FINAL FOR-SALE MATCH COMPLETE")
print("==============================================")
print(f"Matching properties: {len(matched):,}")
print(f"Created: {OUTPUT_FILE.name}")
print("Only House-Verified properties with Active Redfin listings are included.")
print("==============================================")
