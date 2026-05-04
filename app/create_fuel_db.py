import pandas as pd
from pathlib import Path

# Paths to source files
DATA_DIR = Path(__file__).resolve().parent.parent / "data"
source_fuel = DATA_DIR / "Oil_Data_Consolidated.xlsx"
source_price = DATA_DIR / "Transformed_for_Analysis.xlsx"
target = DATA_DIR / "Fuel.xlsx"

# Read sheets from source files
actual = pd.read_excel(source_fuel, sheet_name="Actual")
resupply = pd.read_excel(source_fuel, sheet_name="Resupply")
terminal = pd.read_excel(source_fuel, sheet_name="Terminal")
fuel_price = pd.read_excel(source_price, sheet_name="FuelPrice_Long")
tariff = pd.read_excel(source_price, sheet_name="Tariff_Long")

# Write all sheets to the new Fuel.xlsx
with pd.ExcelWriter(target, engine="openpyxl") as writer:
    actual.to_excel(writer, sheet_name="Actual", index=False)
    resupply.to_excel(writer, sheet_name="Resupply", index=False)
    terminal.to_excel(writer, sheet_name="Terminal", index=False)
    fuel_price.to_excel(writer, sheet_name="FuelPrice_Long", index=False)
    tariff.to_excel(writer, sheet_name="Tariff_Long", index=False)

print(f"Created {target} with all required sheets.")
