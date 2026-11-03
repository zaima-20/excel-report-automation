"""
Automate Excel Reports with Python
-----------------------------------
Reads a raw sales Excel file, cleans it (duplicates, missing values, bad types),
builds summary tables, and saves a formatted multi-sheet report with a chart.

Requirements:  pip install pandas openpyxl
Run:           python excel_report_automation.py
"""

from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment

RAW_FILE = Path("raw_sales.xlsx")
REPORT_FILE = Path("sales_report.xlsx")


# ---------------------------------------------------------------
# Step 0: Create a messy sample file (so readers can follow along)
# ---------------------------------------------------------------
def create_sample_data(path: Path) -> None:
    rng = np.random.default_rng(42)
    n = 200
    df = pd.DataFrame({
        "Order ID": range(1001, 1001 + n),
        "Date": pd.date_range("2026-01-01", periods=n, freq="D").strftime("%Y-%m-%d"),
        "Region": rng.choice(["North", "South", "East", "West"], n),
        "Product": rng.choice(["Laptop", "Mouse", "Keyboard", "Monitor"], n),
        "Quantity": rng.integers(1, 10, n).astype(float),
        "Unit Price": rng.choice([250.0, 15.0, 40.0, 180.0], n),
    })
    # Make it messy, like real-world data
    df.loc[rng.choice(n, 12, replace=False), "Quantity"] = np.nan
    df.loc[rng.choice(n, 8, replace=False), "Region"] = np.nan
    df.loc[rng.choice(n, 5, replace=False), "Unit Price"] = np.nan
    df["Region"] = df["Region"].where(rng.random(n) > 0.1, df["Region"].str.lower() + " ")
    df = pd.concat([df, df.sample(15, random_state=1)])  # duplicate rows
    df.to_excel(path, index=False)
    print(f"Sample data created: {path} ({len(df)} rows)")


# ---------------------------------------------------------------
# Step 1: Read the Excel file
# ---------------------------------------------------------------
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    print(f"Loaded {len(df)} rows, {df.shape[1]} columns")
    return df


# ---------------------------------------------------------------
# Step 2: Clean the data
# ---------------------------------------------------------------
def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    stats = {"rows_before": len(df)}

    # 2a. Remove exact duplicate rows
    stats["duplicates_removed"] = int(df.duplicated().sum())
    df = df.drop_duplicates().copy()

    # 2b. Standardise text columns ("north " -> "North")
    df["Region"] = df["Region"].str.strip().str.title()

    # 2c. Fix data types
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # 2d. Handle missing values
    stats["missing_before"] = df.isna().sum().to_dict()
    df["Region"] = df["Region"].fillna("Unknown")
    df["Quantity"] = df["Quantity"].fillna(df["Quantity"].median())
    # Fill price from the same product's typical price
    df["Unit Price"] = df.groupby("Product")["Unit Price"].transform(
        lambda s: s.fillna(s.median())
    )

    # 2e. Add calculated column
    df["Revenue"] = df["Quantity"] * df["Unit Price"]
    df["Month"] = df["Date"].dt.to_period("M").astype(str)

    stats["rows_after"] = len(df)
    return df, stats


# ---------------------------------------------------------------
# Step 3: Build summary tables
# ---------------------------------------------------------------
def build_summaries(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    by_region = (df.groupby("Region", as_index=False)
                   .agg(Orders=("Order ID", "count"),
                        Units=("Quantity", "sum"),
                        Revenue=("Revenue", "sum"))
                   .sort_values("Revenue", ascending=False))

    by_product = (df.groupby("Product", as_index=False)
                    .agg(Units=("Quantity", "sum"), Revenue=("Revenue", "sum"))
                    .sort_values("Revenue", ascending=False))

    monthly = df.pivot_table(index="Month", columns="Region",
                             values="Revenue", aggfunc="sum", fill_value=0)
    monthly = monthly.reset_index()

    return {"By Region": by_region, "By Product": by_product, "Monthly": monthly}


# ---------------------------------------------------------------
# Step 4: Save everything to one Excel report
# ---------------------------------------------------------------
def save_report(df: pd.DataFrame, summaries: dict, stats: dict, path: Path) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        overview = pd.DataFrame({
            "Metric": ["Rows before cleaning", "Duplicates removed",
                       "Rows after cleaning", "Total revenue"],
            "Value": [stats["rows_before"], stats["duplicates_removed"],
                      stats["rows_after"], round(df["Revenue"].sum(), 2)],
        })
        overview.to_excel(writer, sheet_name="Overview", index=False)
        for name, table in summaries.items():
            table.to_excel(writer, sheet_name=name, index=False)
        df.to_excel(writer, sheet_name="Clean Data", index=False)


# ---------------------------------------------------------------
# Step 5: Format the report and add a chart
# ---------------------------------------------------------------
def format_report(path: Path) -> None:
    wb = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")

    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(width + 3, 30)
        ws.freeze_panes = "A2"

    # Bar chart of revenue by region
    ws = wb["By Region"]
    chart = BarChart()
    chart.title = "Revenue by Region"
    chart.y_axis.title = "Revenue"
    data = Reference(ws, min_col=4, min_row=1, max_row=ws.max_row)
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "F2")

    wb.save(path)


def main() -> None:
    if not RAW_FILE.exists():
        create_sample_data(RAW_FILE)
    df = load_data(RAW_FILE)
    clean, stats = clean_data(df)
    summaries = build_summaries(clean)
    save_report(clean, summaries, stats, REPORT_FILE)
    format_report(REPORT_FILE)
    print(f"Duplicates removed: {stats['duplicates_removed']}")
    print(f"Report saved: {REPORT_FILE}")


if __name__ == "__main__":
    main()
