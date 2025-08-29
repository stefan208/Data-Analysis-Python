import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# Config
FIN_PATH = Path("data/financial_kpis.xlsx")
OPS_PATH = Path("data/operational_kpis.xlsx")
DATE_COL = "date"
REVENUE_COL = "revenue"  

# Load
financial_df = pd.read_excel(FIN_PATH)
operational_df = pd.read_excel(OPS_PATH)

# Standardize & parse
financial_df.columns = financial_df.columns.str.lower().str.strip()
operational_df.columns = operational_df.columns.str.lower().str.strip()

# Ensure date column exists
if DATE_COL not in financial_df.columns or DATE_COL not in operational_df.columns:
    raise ValueError(f"'{DATE_COL}' column must be present in both Excel files.")

# Parse dates
financial_df[DATE_COL] = pd.to_datetime(financial_df[DATE_COL])
operational_df[DATE_COL] = pd.to_datetime(operational_df[DATE_COL])

# Merge & basic cleaning
df = pd.merge(financial_df, operational_df, on=DATE_COL, how="inner")
df = df.drop_duplicates().sort_values(DATE_COL)

# Forward/backward fill numeric gaps (after sorting)
numeric_cols = df.select_dtypes(include="number").columns.tolist()
df[numeric_cols] = df[numeric_cols].ffill().bfill()

# Set index for time-series convenience
df = df.set_index(DATE_COL)

# Sanity checks / validation

if REVENUE_COL not in df.columns:
    raise ValueError(f"'{REVENUE_COL}' column not found after merge.")

if df.index.is_monotonic_increasing is False:
    df = df.sort_index()

# Check for remaining missing values in key fields
if df[REVENUE_COL].isna().any():
    raise ValueError("Revenue still contains missing values after filling.")


# KPI calculations

def calculate_cagr(series: pd.Series) -> float:
    # Assumes first and last values are valid and index spans multiple years
    start, end = series.iloc[0], series.iloc[-1]
    n_years = (series.index[-1] - series.index[0]).days / 365.25
    if n_years <= 0:
        return float("nan")
    return (end / start) ** (1 / n_years) - 1

# Year-over-Year growth (expects monthly data; 12-period difference)
df["revenue_yoy"] = df[REVENUE_COL].pct_change(periods=12)

cagr = calculate_cagr(df[REVENUE_COL])

print("=== KPI Summary ===")
print(f"CAGR (Revenue): {cagr:.2%}" if pd.notna(cagr) else "CAGR (Revenue): N/A")
last_yoy = df["revenue_yoy"].iloc[-1]
print(f"Latest YoY (Revenue): {last_yoy:.2%}" if pd.notna(last_yoy) else "Latest YoY (Revenue): N/A")

# Visualization (single, simple chart)
plt.figure(figsize=(10, 5))
plt.plot(df.index, df[REVENUE_COL])
plt.title("Revenue Over Time")
plt.xlabel("Date")
plt.ylabel("Revenue")
plt.tight_layout()
plt.show()
Add data analysis script
