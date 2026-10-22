import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ============================================================
# PRINCE GEORGE'S COUNTY VACANCY LEAD FINDER
# VERSION: HOUSE VERIFIED + CONDITION SCORING
# ============================================================

input_file = "prince_georges_houses.csv"
output_file = "Prince_Georges_Vacancy_Leads_HOUSE_VERIFIED.xlsx"

print()
print("================================================")
print(" HOUSE-VERIFIED VACANCY LEAD FINDER")
print("================================================")
print()
print("Loading property data...")
print()

# ------------------------------------------------------------
# HELPER: FIND A COLUMN BY KEYWORDS
# ------------------------------------------------------------

columns = pd.read_csv(input_file, nrows=0).columns.tolist()


def find_column(required_keywords, optional_keywords=None):
    optional_keywords = optional_keywords or []

    # First try columns containing every required keyword.
    matches = [
        c for c in columns
        if all(k.upper() in c.upper() for k in required_keywords)
    ]

    if not matches:
        raise ValueError(
            "Could not find a required column containing: "
            + ", ".join(required_keywords)
        )

    # Prefer matches containing optional keywords.
    for c in matches:
        if all(k.upper() in c.upper() for k in optional_keywords):
            return c

    return matches[0]


# ------------------------------------------------------------
# FIND IMPORTANT COLUMNS
# ------------------------------------------------------------

account_column = find_column(["ACCOUNT NUMBER"])

address_column = find_column(
    ["MDP STREET ADDRESS", "MDP FIELD: ADDRESS"]
)

occupancy_column = find_column(["OWNER OCCUPANCY CODE"])

land_use_column = find_column(["LAND USE CODE"])

# Current-cycle values, not historical sale values.
improvement_column = find_column(
    ["CURRENT CYCLE DATA", "IMPROVEMENTS VALUE"]
)

land_value_column = find_column(
    ["CURRENT CYCLE DATA", "LAND VALUE"]
)

# Structure / dwelling fields.
year_built_column = find_column(
    ["YEAR BUILT"],
    ["PRIMARY STRUCTURE"]
)

dwelling_units_column = find_column(
    ["NUMBER OF DWELLING UNITS"]
)

structure_area_column = find_column(
    ["STRUCTURE AREA"]
)

dwelling_type_column = find_column(
    ["DWELLING TYPE"]
)

# Stories are useful supporting evidence but are not required
# because the main house filter already uses dwelling units,
# structure area, year built, land use, and improvement value.
try:
    stories_column = find_column(["NUMBER OF STORIES"])
except ValueError:
    stories_column = None

# ------------------------------------------------------------
# FIND SALES SEGMENTS 1, 2, AND 3
# ------------------------------------------------------------

sale_segments = []

for segment in ["1", "2", "3"]:
    try:
        date_col = find_column(
            [f"SALES SEGMENT {segment}", "TRANSFER DATE"]
        )
    except ValueError:
        date_col = None

    try:
        amount_col = find_column(
            [f"SALES SEGMENT {segment}", "CONSIDERATION"]
        )
    except ValueError:
        amount_col = None

    sale_segments.append((date_col, amount_col))


# ------------------------------------------------------------
# LOAD ONLY THE COLUMNS WE NEED
# ------------------------------------------------------------

use_columns = [
    account_column,
    address_column,
    occupancy_column,
    land_use_column,
    improvement_column,
    land_value_column,
    year_built_column,
    dwelling_units_column,
    structure_area_column,
    dwelling_type_column
]

if stories_column:
    use_columns.append(stories_column)

for date_col, amount_col in sale_segments:
    if date_col:
        use_columns.append(date_col)
    if amount_col:
        use_columns.append(amount_col)

# Remove duplicates while preserving order.
use_columns = list(dict.fromkeys(use_columns))

properties = pd.read_csv(
    input_file,
    usecols=use_columns,
    low_memory=False
)

# ------------------------------------------------------------
# RENAME CORE COLUMNS
# ------------------------------------------------------------

rename_map = {
    account_column: "Account ID",
    address_column: "Address",
    occupancy_column: "Owner Occupancy",
    land_use_column: "Land Use",
    improvement_column: "Current Improvement Value",
    land_value_column: "Current Land Value",
    year_built_column: "Year Built",
    dwelling_units_column: "Dwelling Units",
    structure_area_column: "Structure Area Sq Ft",
    dwelling_type_column: "Dwelling Type"
}

if stories_column:
    rename_map[stories_column] = "Stories"

properties = properties.rename(columns=rename_map)

# ------------------------------------------------------------
# CLEAN NUMERIC FIELDS
# ------------------------------------------------------------

def clean_number(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip(),
        errors="coerce"
    )


for column in [
    "Current Improvement Value",
    "Current Land Value",
    "Year Built",
    "Dwelling Units",
    "Structure Area Sq Ft"
]:
    properties[column] = clean_number(properties[column])

if "Stories" in properties.columns:
    properties["Stories"] = clean_number(properties["Stories"])

# ------------------------------------------------------------
# BUILD SALES DATA FROM SEGMENTS 1, 2, AND 3
# ------------------------------------------------------------

sale_date_series = []
sale_amount_series = []

for i, (date_col, amount_col) in enumerate(sale_segments, start=1):

    date_name = f"_Sale_Date_{i}"
    amount_name = f"_Sale_Amount_{i}"

    if date_col:
        properties[date_name] = pd.to_datetime(
            properties[date_col],
            errors="coerce"
        )
        sale_date_series.append(date_name)

    if amount_col:
        properties[amount_name] = clean_number(
            properties[amount_col]
        )
        sale_amount_series.append(amount_name)

# Find the most recent sale among all three segments.
properties["Last Sale Date"] = pd.NaT
properties["Last Sale Amount"] = pd.NA

for idx, row in properties.iterrows():

    best_date = pd.NaT
    best_amount = pd.NA

    for i in range(1, 4):

        date_name = f"_Sale_Date_{i}"
        amount_name = f"_Sale_Amount_{i}"

        if date_name not in properties.columns:
            continue

        current_date = row[date_name]

        if pd.notna(current_date):
            if pd.isna(best_date) or current_date > best_date:
                best_date = current_date

                if amount_name in properties.columns:
                    best_amount = row[amount_name]
                else:
                    best_amount = pd.NA

    properties.at[idx, "Last Sale Date"] = best_date
    properties.at[idx, "Last Sale Amount"] = best_amount

# ------------------------------------------------------------
# REMOVE TEMPORARY SALES COLUMNS
# ------------------------------------------------------------

temporary_columns = [
    c for c in properties.columns
    if c.startswith("_Sale_")
]

temporary_original_sales = [
    c for c in properties.columns
    if c.startswith("SALES SEGMENT")
]

# Drop temporary calculated columns only.
if temporary_columns:
    properties = properties.drop(columns=temporary_columns)

# Drop original segment columns because the final report uses
# the calculated most-recent sale.
for c in list(properties.columns):
    if "SALES SEGMENT" in c.upper():
        properties = properties.drop(columns=[c])

# ------------------------------------------------------------
# CLEAN BASIC DATA
# ------------------------------------------------------------

properties["Address"] = properties["Address"].astype(str).str.strip()
properties["Owner Occupancy"] = (
    properties["Owner Occupancy"].astype(str).str.strip().str.upper()
)
properties["Land Use"] = properties["Land Use"].astype(str).str.strip()

properties = properties[
    properties["Address"].notna() &
    (properties["Address"] != "") &
    (properties["Address"].str.upper() != "NAN")
].copy()

# ------------------------------------------------------------
# HOUSE VERIFICATION
# ------------------------------------------------------------

current_year = pd.Timestamp.today().year

valid_year = (
    properties["Year Built"].notna() &
    properties["Year Built"].between(1700, current_year)
)

valid_structure_area = (
    properties["Structure Area Sq Ft"].notna() &
    (properties["Structure Area Sq Ft"] > 0)
)

valid_dwelling_units = (
    properties["Dwelling Units"].notna() &
    (properties["Dwelling Units"] >= 1)
)

valid_improvement = (
    properties["Current Improvement Value"].notna() &
    (properties["Current Improvement Value"] > 0)
)

residential_land_use = properties["Land Use"].str.contains(
    r"Residential \(R\)|Town House \(TH\)|Residential|Town House",
    case=False,
    na=False,
    regex=True
)

dwelling_type_present = (
    properties["Dwelling Type"].notna() &
    properties["Dwelling Type"].astype(str).str.strip().ne("") &
    properties["Dwelling Type"].astype(str).str.upper().ne("NAN")
)

properties["House Verified"] = (
    valid_improvement &
    valid_structure_area &
    valid_year &
    valid_dwelling_units &
    residential_land_use &
    dwelling_type_present
).map({True: "YES", False: "NO"})

properties["House Verification Reason"] = ""

properties.loc[
    valid_improvement,
    "House Verification Reason"
] += "Improvement value > $0; "

properties.loc[
    valid_structure_area,
    "House Verification Reason"
] += "Structure area > 0; "

properties.loc[
    valid_year,
    "House Verification Reason"
] += "Valid year built; "

properties.loc[
    valid_dwelling_units,
    "House Verification Reason"
] += "Dwelling units >= 1; "

properties.loc[
    residential_land_use,
    "House Verification Reason"
] += "Residential/Town House land use; "

properties.loc[
    dwelling_type_present,
    "House Verification Reason"
] += "Dwelling type present; "

properties["House Verification Reason"] = (
    properties["House Verification Reason"]
    .str.rstrip("; ")
)

# IMPORTANT: only keep records that pass every house check.
properties = properties[
    properties["House Verified"] == "YES"
].copy()

# ------------------------------------------------------------
# CALCULATE YEARS SINCE LAST SALE
# ------------------------------------------------------------

today = pd.Timestamp.today()

properties["Years Since Sale"] = (
    today - properties["Last Sale Date"]
).dt.days / 365.25

# ------------------------------------------------------------
# MANUAL VISUAL INSPECTION FIELDS
# ------------------------------------------------------------
# The SDAT property file does NOT establish whether windows are
# boarded up or vegetation is overgrown. These are therefore
# manual inspection fields.
#
# Allowed entries:
# Property Condition: POOR / VERY POOR / FAIR / GOOD / UNKNOWN
# Boarded Up Windows: YES / NO / UNKNOWN
# Overgrown Vegetation: YES / NO / UNKNOWN
#
# A POOR or VERY POOR condition = +30 points.
# Boarded Up Windows YES = +30 points.
# Overgrown Vegetation YES = +30 points.

properties["Property Condition"] = "UNKNOWN"
properties["Boarded Up Windows"] = "UNKNOWN"
properties["Overgrown Vegetation"] = "UNKNOWN"

# ------------------------------------------------------------
# BASE VACANCY SCORE
# ------------------------------------------------------------

properties["Vacancy Score"] = 0
properties["Reason"] = ""

# Indicator 1: Non-owner occupied
mask = properties["Owner Occupancy"].eq("N")
properties.loc[mask, "Vacancy Score"] += 30
properties.loc[mask, "Reason"] += "Non-owner occupied; "

# Indicator 2: No recorded sale
mask = properties["Last Sale Date"].isna()
properties.loc[mask, "Vacancy Score"] += 20
properties.loc[mask, "Reason"] += "No recorded sale date; "

# Indicator 3: Sale 15+ years ago
mask = properties["Years Since Sale"] >= 15
properties.loc[mask, "Vacancy Score"] += 20
properties.loc[mask, "Reason"] += "Last sale 15+ years ago; "

# Indicator 4: Sale 10-15 years ago
mask = (
    (properties["Years Since Sale"] >= 10) &
    (properties["Years Since Sale"] < 15)
)
properties.loc[mask, "Vacancy Score"] += 12
properties.loc[mask, "Reason"] += "Last sale 10-15 years ago; "

# Indicator 5: Sale 5-10 years ago
mask = (
    (properties["Years Since Sale"] >= 5) &
    (properties["Years Since Sale"] < 10)
)
properties.loc[mask, "Vacancy Score"] += 6
properties.loc[mask, "Reason"] += "Last sale 5-10 years ago; "

# Indicator 6: No sale amount
mask = properties["Last Sale Amount"].isna()
properties.loc[mask, "Vacancy Score"] += 5
properties.loc[mask, "Reason"] += "No recorded sale amount; "

# ------------------------------------------------------------
# VISUAL CONDITION POINTS
# ------------------------------------------------------------

# Property condition: POOR or VERY POOR = +30
mask = properties["Property Condition"].isin(
    ["POOR", "VERY POOR"]
)
properties.loc[mask, "Vacancy Score"] += 30
properties.loc[mask, "Reason"] += "Poor property condition; "

# Boarded windows = +30
mask = properties["Boarded Up Windows"].eq("YES")
properties.loc[mask, "Vacancy Score"] += 30
properties.loc[mask, "Reason"] += "Boarded-up windows; "

# Overgrown vegetation = +30
mask = properties["Overgrown Vegetation"].eq("YES")
properties.loc[mask, "Vacancy Score"] += 30
properties.loc[mask, "Reason"] += "Overgrown vegetation; "

# ------------------------------------------------------------
# LIMIT SCORE TO 100
# ------------------------------------------------------------

properties["Vacancy Score"] = properties[
    "Vacancy Score"
].clip(upper=100)

# ------------------------------------------------------------
# PRIORITY
# ------------------------------------------------------------

def get_priority(score):
    if score >= 70:
        return "HIGH"
    elif score >= 50:
        return "MEDIUM"
    elif score >= 30:
        return "LOW"
    else:
        return "VERY LOW"


properties["Priority"] = properties[
    "Vacancy Score"
].apply(get_priority)

# ------------------------------------------------------------
# CLEAN REASON
# ------------------------------------------------------------

properties["Reason"] = (
    properties["Reason"]
    .str.rstrip("; ")
)

properties.loc[
    properties["Reason"] == "",
    "Reason"
] = "No major vacancy indicators found"

# ------------------------------------------------------------
# SORT
# ------------------------------------------------------------

properties = properties.sort_values(
    ["Vacancy Score", "Years Since Sale"],
    ascending=[False, False]
)

# ------------------------------------------------------------
# FINAL COLUMNS
# ------------------------------------------------------------

final_columns = [
    "Account ID",
    "Address",
    "Owner Occupancy",
    "Land Use",
    "House Verified",
    "House Verification Reason",
    "Year Built",
    "Dwelling Units",
    "Stories",
    "Structure Area Sq Ft",
    "Dwelling Type",
    "Last Sale Date",
    "Last Sale Amount",
    "Current Improvement Value",
    "Current Land Value",
    "Years Since Sale",
    "Property Condition",
    "Boarded Up Windows",
    "Overgrown Vegetation",
    "Vacancy Score",
    "Priority",
    "Reason"
]

# Only include Stories if it exists.
final_columns = [
    c for c in final_columns
    if c in properties.columns
]

properties = properties[final_columns]

# ------------------------------------------------------------
# CREATE EXCEL REPORT
# ------------------------------------------------------------

print("Creating Excel report...")
print()

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    properties.to_excel(
        writer,
        sheet_name="All Properties",
        index=False
    )

    properties[
        properties["Priority"] == "HIGH"
    ].to_excel(
        writer,
        sheet_name="HIGH Priority",
        index=False
    )

    properties[
        properties["Priority"] == "MEDIUM"
    ].to_excel(
        writer,
        sheet_name="MEDIUM Priority",
        index=False
    )

    properties[
        properties["Priority"] == "LOW"
    ].to_excel(
        writer,
        sheet_name="LOW Priority",
        index=False
    )

    # Instructions sheet
    instructions = pd.DataFrame({
        "Field": [
            "Property Condition",
            "Boarded Up Windows",
            "Overgrown Vegetation",
            "House Verified"
        ],
        "How to use": [
            "Enter POOR or VERY POOR for +30 vacancy points. FAIR/GOOD/UNKNOWN = 0 points.",
            "Enter YES for +30 vacancy points. NO/UNKNOWN = 0 points.",
            "Enter YES for +30 vacancy points. NO/UNKNOWN = 0 points.",
            "YES means the property passed all six residential-structure checks."
        ]
    })

    instructions.to_excel(
        writer,
        sheet_name="Instructions",
        index=False
    )

# ------------------------------------------------------------
# FORMAT EXCEL
# ------------------------------------------------------------

print("Formatting Excel report...")

workbook = load_workbook(output_file)

for worksheet in workbook.worksheets:

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    # Header row formatting and widths
    for column_cells in worksheet.columns:

        column_letter = get_column_letter(
            column_cells[0].column
        )

        max_length = 0

        for cell in column_cells:
            try:
                max_length = max(
                    max_length,
                    len(str(cell.value))
                )
            except:
                pass

        worksheet.column_dimensions[
            column_letter
        ].width = min(max_length + 2, 45)

    # Find columns by header name so formatting does not
    # break if columns are added or moved.
    header_map = {
        cell.value: cell.column
        for cell in worksheet[1]
    }

    # Date
    if "Last Sale Date" in header_map:
        col = get_column_letter(header_map["Last Sale Date"])
        for cell in worksheet[col][1:]:
            if cell.value is not None:
                cell.number_format = "yyyy-mm-dd"

    # Money
    for name in [
        "Last Sale Amount",
        "Current Improvement Value",
        "Current Land Value"
    ]:
        if name in header_map:
            col = get_column_letter(header_map[name])
            for cell in worksheet[col][1:]:
                if cell.value is not None:
                    cell.number_format = '$#,##0'

    # Numeric structure fields
    for name in [
        "Year Built",
        "Dwelling Units",
        "Stories",
        "Structure Area Sq Ft"
    ]:
        if name in header_map:
            col = get_column_letter(header_map[name])
            for cell in worksheet[col][1:]:
                if cell.value is not None:
                    cell.number_format = '#,##0'

    # Vacancy score color scale
    if "Vacancy Score" in header_map and worksheet.max_row >= 2:
        score_col = get_column_letter(
            header_map["Vacancy Score"]
        )

        worksheet.conditional_formatting.add(
            f"{score_col}2:{score_col}{worksheet.max_row}",
            ColorScaleRule(
                start_type="min",
                start_color="FFFFFF",
                mid_type="percentile",
                mid_value=50,
                mid_color="FFFF00",
                end_type="max",
                end_color="FF0000"
            )
        )

# ------------------------------------------------------------
# ADD EXCEL DROPDOWNS FOR MANUAL INSPECTION
# ------------------------------------------------------------

for sheet_name in [
    "All Properties",
    "HIGH Priority",
    "MEDIUM Priority",
    "LOW Priority"
]:

    if sheet_name not in workbook.sheetnames:
        continue

    worksheet = workbook[sheet_name]

    headers = {
        cell.value: cell.column
        for cell in worksheet[1]
    }

    # Property Condition dropdown
    if "Property Condition" in headers:
        condition_col = get_column_letter(
            headers["Property Condition"]
        )

        validation = DataValidation(
            type="list",
            formula1='"POOR,VERY POOR,FAIR,GOOD,UNKNOWN"',
            allow_blank=False
        )

        worksheet.add_data_validation(validation)
        # Only add the dropdown when the sheet has data rows.
        if worksheet.max_row >= 2:
            validation.add(
                f"{condition_col}2:{condition_col}{worksheet.max_row}"
            )

    # YES/NO/UNKNOWN dropdowns
    for field in [
        "Boarded Up Windows",
        "Overgrown Vegetation"
    ]:

        if field in headers:
            field_col = get_column_letter(
                headers[field]
            )

            validation = DataValidation(
                type="list",
                formula1='"YES,NO,UNKNOWN"',
                allow_blank=False
            )

            worksheet.add_data_validation(validation)
            # Only add the dropdown when the sheet has data rows.
            if worksheet.max_row >= 2:
                validation.add(
                    f"{field_col}2:{field_col}{worksheet.max_row}"
                )

workbook.save(output_file)

# ------------------------------------------------------------
# SUMMARY
# ------------------------------------------------------------

high_count = len(
    properties[properties["Priority"] == "HIGH"]
)

medium_count = len(
    properties[properties["Priority"] == "MEDIUM"]
)

low_count = len(
    properties[properties["Priority"] == "LOW"]
)

very_low_count = len(
    properties[properties["Priority"] == "VERY LOW"]
)

print()
print("================================================")
print(" HOUSE-VERIFIED REPORT CREATED - FIXED")
print("================================================")
print()
print("File:")
print(output_file)
print()
print("House-verified properties:", f"{len(properties):,}")
print()
print("HIGH PRIORITY:     ", f"{high_count:,}")
print("MEDIUM PRIORITY:   ", f"{medium_count:,}")
print("LOW PRIORITY:      ", f"{low_count:,}")
print("VERY LOW PRIORITY: ", f"{very_low_count:,}")
print()
print("================================================")
print("Finished!")
print("================================================")
