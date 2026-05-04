"""
transform_master.py
Transforms Master.xlsx into the Fuel.xlsx format expected by the dashboard.

Usage:
    python app/transform_master.py
"""

import pandas as pd
from pathlib import Path

MASTER = Path(r"C:\Users\halea\NextCloud\Documents\Petroleum\Fuel Company\Master.xlsx")
OUTPUT = Path(__file__).resolve().parent.parent / "data" / "Fuel.xlsx"

# ---------------------------------------------------------------------------
# 1. ACTUAL sheet  (from Sheet5 which already has a proper 'Date' column)
# ---------------------------------------------------------------------------
actual_raw = pd.read_excel(MASTER, sheet_name="Sheet5")
actual_raw.columns = actual_raw.columns.str.strip()

# Parse dates (e.g. '17/03/26' or datetime objects)
actual_raw["Date"] = pd.to_datetime(actual_raw["Date"], dayfirst=True, errors="coerce")

# Drop placeholder rows: monthly template entries (4th of each month) where
# Offtake is 0 — these are auto-extended future rows with no real data.
actual_raw["_offtake"] = pd.to_numeric(actual_raw["Offtake"], errors="coerce").fillna(0)
actual_raw["_day"] = actual_raw["Date"].dt.day
placeholder_mask = (actual_raw["_offtake"] == 0) & (actual_raw["_day"] == 4)
actual_raw = actual_raw[~placeholder_mask].drop(columns=["_offtake", "_day"])

# Also cap at today so no future-dated rows sneak through
actual_raw = actual_raw[actual_raw["Date"] <= pd.Timestamp.today()]

# Rename to match dashboard expectations
actual = actual_raw[["Date", "Closing Stock", "Offtake", "Fuel Type", "Company", "Location"]].copy()
actual["Tonga Power Offtake"] = 0  # not in source; fill with 0

# ---------------------------------------------------------------------------
# 2. RESUPPLY sheet  (pivot wide → long so each row is one fuel-type qty)
# ---------------------------------------------------------------------------
resupply_raw = pd.read_excel(MASTER, sheet_name="Resupply")
resupply_raw.columns = resupply_raw.columns.str.strip()
resupply_raw["Date"] = pd.to_datetime(resupply_raw["Date"], dayfirst=True, errors="coerce")

fuel_cols = [c for c in resupply_raw.columns if c in ("Petrol", "Diesel", "Jet-A1")]
resupply = resupply_raw.melt(
    id_vars=["Date", "Company", "Location"],
    value_vars=fuel_cols,
    var_name="Fuel Type",
    value_name="Quantity",
).dropna(subset=["Quantity"])
resupply = resupply[["Date", "Quantity", "Fuel Type", "Company", "Location"]]

# ---------------------------------------------------------------------------
# 3. TERMINAL sheet  (empty in source — create a minimal placeholder)
# ---------------------------------------------------------------------------
terminal = pd.DataFrame(columns=["Date", "Quantity", "Fuel Type", "Company", "Location"])

# ---------------------------------------------------------------------------
# 4. Write output
# ---------------------------------------------------------------------------
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

# Load existing FuelPrice_Long and Tariff_Long from current Fuel.xlsx if it
# exists, so we don't lose price/tariff data already in place.
fuel_price = pd.DataFrame(columns=["Date", "Fuel Type", "Price"])
tariff = pd.DataFrame(columns=["Date", "Value"])

existing = OUTPUT
if existing.exists():
    try:
        xls = pd.ExcelFile(existing)
        if "FuelPrice_Long" in xls.sheet_names:
            fuel_price = pd.read_excel(existing, sheet_name="FuelPrice_Long")
        if "Tariff_Long" in xls.sheet_names:
            tariff = pd.read_excel(existing, sheet_name="Tariff_Long")
    except Exception as e:
        print(f"Warning: could not read existing price/tariff sheets: {e}")

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    actual.to_excel(writer, sheet_name="Actual", index=False)
    resupply.to_excel(writer, sheet_name="Resupply", index=False)
    terminal.to_excel(writer, sheet_name="Terminal", index=False)
    fuel_price.to_excel(writer, sheet_name="FuelPrice_Long", index=False)
    tariff.to_excel(writer, sheet_name="Tariff_Long", index=False)

print(f"Done. Written to: {OUTPUT}")
print(f"  Actual rows    : {len(actual)}")
print(f"  Resupply rows  : {len(resupply)}")
print(f"  Terminal rows  : {len(terminal)}")
print(f"  FuelPrice rows : {len(fuel_price)}")
print(f"  Tariff rows    : {len(tariff)}")
