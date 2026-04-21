# Dashboard Calculations Reference

This document explains how the dashboard calculates KPIs, status labels, and chart values in `dashboard.py`.

## 1. Data Sources

- Main operational data workbook: `Oil_Data_Consolidated.xlsx`
  - Sheet `Actual`
  - Sheet `Resupply`
  - Sheet `Terminal`
- Price and tariff workbook: `Transformed_for_Analysis.xlsx`
  - Sheet `FuelPrice_Long`
  - Sheet `Tariff_Long`

## 2. Data Preparation Rules

Before calculations, the app standardizes key fields:

- Dates are parsed with `pd.to_datetime(..., errors="coerce")`
- Numeric fields are parsed with `pd.to_numeric(..., errors="coerce")`
- Invalid values become null (`NaT` or `NaN`) and are excluded where required

## 3. Filter Logic

### Fuel Supply section filters
Applied to both `Actual` and `Resupply` using:

- Company
- Location
- Fuel Type
- Month (derived from Date as `YYYY-MM`)

### Terminal section filters
Applied only to `Terminal`:

- Terminal Location
- Terminal Type (`Terminal Info`)

### Prices and Tariffs filters
Applied only inside Prices and Tariffs section:

- Fuel prices: Price Type, Fuel
- Tariff: Tariff Component

## 4. Core KPI Calculations (Fuel Supply)

All formulas below are based on the filtered data in Fuel Supply.

### 4.1 Total Fuel (Total Stock)

Method:
1. Keep rows with non-null `Date` and `Closing Stock`
2. Sort by `Date`
3. For each unique `(Company, Location, Fuel Type)`, keep only the latest row
4. Sum `Closing Stock`

Formula (conceptual):

`Total Fuel = sum(latest Closing Stock per Company-Location-Fuel Type)`

### 4.2 Total Offtake and Non-Power Offtake

- `Total Offtake` is the sum of `Offtake`
- `Tonga Power` is the sum of `Tonga Power Offtake`
- `Non-Power` is computed as:

`Non-Power = max(Offtake - Tonga Power Offtake, 0)` per row, then summed

The `clip(lower=0)` rule prevents negative non-power values.

### 4.3 Upcoming Supply

Method:
- Keep `Resupply` rows with non-null `Quantity`
- Sum `Quantity`

Formula:

`Upcoming Supply = sum(Resupply Quantity)`

### 4.4 Average Daily Offtake

Method:
- Use total consumption (`Total Offtake`)
- Divide by number of unique dates in filtered Actual data

Formula:

`Average Daily Offtake = Total Offtake / count(unique Date)`

If there are no dates, this is set to `0`.

### 4.5 Days of Cover

Formula:

`Days of Cover = Total Fuel / Average Daily Offtake`

If `Average Daily Offtake <= 0`, Days of Cover is set to `0`.

### 4.6 Cover Status Thresholds

- `Safe` if Days of Cover >= 45
- `Watch` if Days of Cover >= 30 and < 45
- `Critical` if Days of Cover < 30

## 5. Stock by Fuel Type

Method:
1. Keep rows with non-null `Date` and `Closing Stock`
2. Keep latest row per `(Company, Location, Fuel Type)`
3. Group by `Fuel Type` and sum `Closing Stock`

Used in the summary and PDF.

## 6. Fuel Price Calculations

### 6.1 Latest Retail Prices KPI

Method:
1. Filter to `Price_Type == Retail`
2. Keep rows with non-null `Date` and `Price`
3. Sort by `Date`
4. Keep latest row per `Fuel`

Displayed as current retail price per fuel.

### 6.2 Fuel Price Trend Chart

Method:
1. Apply selected Price Type and Fuel filters
2. Keep non-null Date and Price
3. Sort by Date
4. Plot line chart of `Price` over `Date`

## 7. Tariff Calculations

### 7.1 Tariff date conversion (fiscal year handling)

Tariff records contain `Year` (format like `2025/26`) and `Month` (e.g., `Jul`, `Jan`).

Conversion logic:
- Start year = first part of `Year` (before `/`)
- If month is Jul-Dec, calendar year = start year
- If month is Jan-Jun, calendar year = start year + 1
- Period is set to first day of that month

### 7.2 Components included

Only these components are used for tariff KPI/trend:

- Fuel Component
- Non Fuel component
- Total Tariff

### 7.3 Current Tariff KPI

Method:
1. Take latest tariff year present in data
2. Within that year, find max `Period` (latest month)
3. Group by Component and average Value for that period

Result: Current Tariff values by component for latest month.

### 7.4 Average Tariff KPI

Method:
1. Take latest tariff year
2. Group by Component and average `Value` over all months in that year

Result: Average Tariff values by component for latest tariff year.

### 7.5 Tariff Trend Chart

Method:
1. Apply selected component filters
2. Sort by `Period`
3. Plot `Value` over `Period` by component

## 8. Fuel-Price vs Tariff Dependency Heatmap

Method:
1. Build monthly fuel-price matrix:
   - index: monthly Period
   - columns: `Price_Type + Fuel`
   - values: mean Price
2. Build monthly tariff matrix:
   - index: monthly Period
   - columns: Component
   - values: mean Value
3. Inner join both matrices by month
4. Compute correlation matrix
5. Show only cross-section: Fuel Price series vs Tariff components

Shown only when:
- both matrices are non-empty
- correlation is non-empty
- overlapping months count is at least 3

## 9. Terminal Capacities Chart

Method:
1. Apply terminal filters (Location and Terminal Type)
2. Keep rows with non-null `Quantity`
3. Plot grouped bars:
   - x: Terminal Info
   - y: Quantity
   - color: Fuel Type
   - facet: Location

## 10. PDF Summary Calculations

The PDF uses the same already-computed values from the dashboard session:

- Total Fuel
- Upcoming Supply
- Total Offtake
- Days of Cover and Status
- Stock by Fuel Type
- Latest prices by type/fuel
- Average tariff values for latest year
- Active filter selections
- Last sync timestamp

So PDF numbers are aligned with on-screen dashboard calculations at generation time.
