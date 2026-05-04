"""Upsert Pacific Energy Apr 25 – May 3 data into Fuel.xlsx."""
import pandas as pd
from pathlib import Path

FUEL = Path(__file__).resolve().parent.parent / "data" / "Fuel.xlsx"

# Tongatapu: (date, fuel_type, closing_stock, offtake)
tongatapu = [
    ("2026-04-25","Petrol",305347,0),("2026-04-25","Diesel",88764,0),("2026-04-25","Kerosene",267678,0),
    ("2026-04-26","Petrol",305347,0),("2026-04-26","Diesel",88764,0),("2026-04-26","Kerosene",267678,0),
    ("2026-04-27","Petrol",203390,101200),("2026-04-27","Diesel",64564,24200),("2026-04-27","Kerosene",246367,20900),
    ("2026-04-28","Petrol",990115,43600),("2026-04-28","Diesel",789057,39600),("2026-04-28","Kerosene",420864,3000),
    ("2026-04-29","Petrol",932749,57700),("2026-04-29","Diesel",747002,42000),("2026-04-29","Kerosene",382169,38600),
    ("2026-04-30","Petrol",799353,133396),("2026-04-30","Diesel",551422,195580),("2026-04-30","Kerosene",367458,14711),
    ("2026-05-01","Petrol",712197,86300),("2026-05-01","Diesel",490867,60600),("2026-05-01","Kerosene",347079,20100),
    ("2026-05-02","Petrol",712197,0),("2026-05-02","Diesel",490867,0),("2026-05-02","Kerosene",347079,0),
    ("2026-05-03","Petrol",712197,0),("2026-05-03","Diesel",490867,0),("2026-05-03","Kerosene",347079,0),
]

# Vava'u: (date, fuel_type, closing_stock, offtake)
vavau = [
    ("2026-04-25","Petrol",206291,0),("2026-04-25","Diesel",147301,0),
    ("2026-04-26","Petrol",206291,0),("2026-04-26","Diesel",147301,0),
    ("2026-04-27","Petrol",198308,8000),("2026-04-27","Diesel",143549,4000),
    ("2026-04-28","Petrol",193238,5000),("2026-04-28","Diesel",140574,3000),
    ("2026-04-29","Petrol",185155,8000),("2026-04-29","Diesel",138599,2000),
    ("2026-04-30","Petrol",138513,46500),("2026-04-30","Diesel",112018,26600),
    ("2026-05-01","Petrol",138212,0),("2026-05-01","Diesel",111890,0),
    ("2026-05-02","Petrol",138212,0),("2026-05-02","Diesel",111890,0),
    ("2026-05-03","Petrol",138212,0),("2026-05-03","Diesel",111890,0),
]

new_rows = []
for d, ft, cs, oft in tongatapu:
    new_rows.append({"Date": pd.Timestamp(d), "Closing Stock": cs, "Offtake": oft,
                     "Fuel Type": ft, "Company": "Pacific Energy",
                     "Location": "Tongatapu", "Tonga Power Offtake": 0})
for d, ft, cs, oft in vavau:
    new_rows.append({"Date": pd.Timestamp(d), "Closing Stock": cs, "Offtake": oft,
                     "Fuel Type": ft, "Company": "Pacific Energy",
                     "Location": "Vava'u", "Tonga Power Offtake": 0})

new_df = pd.DataFrame(new_rows)
dates  = new_df["Date"].unique()

actual   = pd.read_excel(str(FUEL), sheet_name="Actual")
resupply = pd.read_excel(str(FUEL), sheet_name="Resupply")
terminal = pd.read_excel(str(FUEL), sheet_name="Terminal")
fp       = pd.read_excel(str(FUEL), sheet_name="FuelPrice_Long")
tariff   = pd.read_excel(str(FUEL), sheet_name="Tariff_Long")

actual["Date"] = pd.to_datetime(actual["Date"], errors="coerce")
mask   = ~((actual["Company"] == "Pacific Energy") & (actual["Date"].isin(dates)))
actual = pd.concat([actual[mask], new_df], ignore_index=True).sort_values(
    ["Date", "Company", "Location", "Fuel Type"])

with pd.ExcelWriter(str(FUEL), engine="openpyxl") as w:
    actual.to_excel(w,   sheet_name="Actual",         index=False)
    resupply.to_excel(w, sheet_name="Resupply",       index=False)
    terminal.to_excel(w, sheet_name="Terminal",       index=False)
    fp.to_excel(w,       sheet_name="FuelPrice_Long", index=False)
    tariff.to_excel(w,   sheet_name="Tariff_Long",    index=False)

pe = actual[actual["Company"] == "Pacific Energy"]
print("Done.")
print(f"  Tongatapu rows : {len(pe[pe['Location'] == 'Tongatapu'])}")
vavau_loc = "Vava'u"
print(f"  Vava'u rows    : {len(pe[pe['Location'] == vavau_loc])}")
print(f"  Total rows     : {len(actual)}")
print(f"  Date range     : {actual['Date'].min().date()} -> {actual['Date'].max().date()}")
