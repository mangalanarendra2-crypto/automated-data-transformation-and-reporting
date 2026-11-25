"""
Automated Data Transformation and Reporting
===========================================

A simple ETL-style pipeline:

    1. EXTRACT    - read raw CSV files from the data/ folder
    2. TRANSFORM  - clean the data and derive new columns
    3. REPORT     - build summary tables, charts, and an Excel report

Run:
    python pipeline.py
    python pipeline.py --input data/raw_sales.csv --output output
"""

import argparse
import logging
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # no display needed; works on servers / schedulers
import matplotlib.pyplot as plt
import pandas as pd

log = logging.getLogger("pipeline")


# --------------------------------------------------------------------------
# 1. EXTRACT
# --------------------------------------------------------------------------
def extract(path: Path) -> pd.DataFrame:
    """Read the raw CSV file into a DataFrame."""
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    df = pd.read_csv(path, dtype=str)  # read everything as text, clean later
    log.info("Extracted %d rows from %s", len(df), path)
    return df


# --------------------------------------------------------------------------
# 2. TRANSFORM
# --------------------------------------------------------------------------
def parse_date(value) -> pd.Timestamp:
    """Accept both YYYY-MM-DD and DD/MM/YYYY; return NaT if unparseable."""
    if pd.isna(value):
        return pd.NaT
    value = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return pd.Timestamp(datetime.strptime(value, fmt))
        except ValueError:
            continue
    return pd.NaT


def transform(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Clean the raw data and add derived columns.

    Returns the cleaned DataFrame and a dict of data-quality statistics.
    """
    stats = {"rows_in": len(df)}

    # Tidy text columns: trim spaces and use consistent capitalisation
    for col in ["customer", "region", "product"]:
        df[col] = df[col].str.strip().str.title()

    # Convert types
    df["order_date"] = df["order_date"].apply(parse_date)
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")

    # Remove exact duplicate orders
    before = len(df)
    df = df.drop_duplicates(subset="order_id", keep="first")
    stats["duplicates_removed"] = before - len(df)

    # Drop rows that are unusable (missing key fields or invalid values)
    before = len(df)
    required = ["order_date", "customer", "quantity", "unit_price"]
    df = df.dropna(subset=required)
    df = df[(df["quantity"] > 0) & (df["unit_price"] > 0)]
    stats["invalid_rows_removed"] = before - len(df)

    # Derived columns
    df = df.copy()
    df["revenue"] = df["quantity"] * df["unit_price"]
    df["month"] = df["order_date"].dt.to_period("M").astype(str)
    df["quantity"] = df["quantity"].astype(int)

    df = df.sort_values("order_date").reset_index(drop=True)
    stats["rows_out"] = len(df)
    log.info(
        "Transformed: %d in -> %d out (%d duplicates, %d invalid)",
        stats["rows_in"],
        stats["rows_out"],
        stats["duplicates_removed"],
        stats["invalid_rows_removed"],
    )
    return df, stats


# --------------------------------------------------------------------------
# 3. REPORT
# --------------------------------------------------------------------------
def build_summaries(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Create the summary tables used in the report."""
    monthly = df.groupby("month", as_index=False).agg(
        orders=("order_id", "count"), units=("quantity", "sum"), revenue=("revenue", "sum")
    )
    monthly["growth_pct"] = (monthly["revenue"].pct_change() * 100).round(1)

    by_region = (
        df.groupby("region", as_index=False)
        .agg(orders=("order_id", "count"), revenue=("revenue", "sum"))
        .sort_values("revenue", ascending=False)
    )
    by_region["share_pct"] = (by_region["revenue"] / by_region["revenue"].sum() * 100).round(1)

    by_product = (
        df.groupby("product", as_index=False)
        .agg(units=("quantity", "sum"), revenue=("revenue", "sum"))
        .sort_values("revenue", ascending=False)
    )

    top_customers = (
        df.groupby("customer", as_index=False)
        .agg(orders=("order_id", "count"), revenue=("revenue", "sum"))
        .sort_values("revenue", ascending=False)
        .head(5)
    )

    return {
        "Monthly": monthly,
        "By Region": by_region,
        "By Product": by_product,
        "Top Customers": top_customers,
    }


def make_charts(summaries: dict[str, pd.DataFrame], out_dir: Path) -> list[Path]:
    """Save a couple of PNG charts and return their paths."""
    paths = []

    monthly = summaries["Monthly"]
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.plot(monthly["month"], monthly["revenue"], marker="o")
    ax.set_title("Monthly Revenue")
    ax.set_ylabel("Revenue (INR)")
    ax.grid(alpha=0.3)
    fig.autofmt_xdate()
    p = out_dir / "monthly_revenue.png"
    fig.tight_layout()
    fig.savefig(p, dpi=120)
    plt.close(fig)
    paths.append(p)

    region = summaries["By Region"]
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.bar(region["region"], region["revenue"])
    ax.set_title("Revenue by Region")
    ax.set_ylabel("Revenue (INR)")
    p = out_dir / "revenue_by_region.png"
    fig.tight_layout()
    fig.savefig(p, dpi=120)
    plt.close(fig)
    paths.append(p)

    return paths


def write_excel_report(
    df: pd.DataFrame, summaries: dict[str, pd.DataFrame], stats: dict, chart_paths: list[Path], path: Path
) -> None:
    """Write everything into one Excel workbook."""
    overview = pd.DataFrame(
        {
            "Metric": [
                "Report generated",
                "Rows in raw file",
                "Duplicates removed",
                "Invalid rows removed",
                "Clean rows",
                "Total revenue",
                "Total orders",
                "Average order value",
            ],
            "Value": [
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                stats["rows_in"],
                stats["duplicates_removed"],
                stats["invalid_rows_removed"],
                stats["rows_out"],
                f"{df['revenue'].sum():,.0f}",
                df["order_id"].nunique(),
                f"{df['revenue'].mean():,.0f}",
            ],
        }
    )

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        overview.to_excel(writer, sheet_name="Overview", index=False)
        for name, table in summaries.items():
            table.to_excel(writer, sheet_name=name, index=False)
        df.to_excel(writer, sheet_name="Clean Data", index=False)

        # Light formatting: auto-fit column widths
        for ws in writer.book.worksheets:
            for column in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in column)
                ws.column_dimensions[column[0].column_letter].width = min(width + 3, 40)

        # Embed charts in the Overview sheet
        from openpyxl.drawing.image import Image

        ws = writer.book["Overview"]
        for i, chart in enumerate(chart_paths):
            img = Image(str(chart))
            img.width, img.height = 480, 275
            ws.add_image(img, f"D{2 + i * 16}")

    log.info("Excel report written to %s", path)


# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="Automated Data Transformation and Reporting")
    parser.add_argument("--input", default="data/raw_sales.csv", help="raw CSV file")
    parser.add_argument("--output", default="output", help="folder for results")
    args = parser.parse_args()

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(out_dir / "pipeline.log"), logging.StreamHandler()],
    )

    log.info("Pipeline started")
    raw = extract(Path(args.input))
    clean, stats = transform(raw)

    clean.to_csv(out_dir / "cleaned_sales.csv", index=False)
    summaries = build_summaries(clean)
    charts = make_charts(summaries, out_dir)
    write_excel_report(clean, summaries, stats, charts, out_dir / "sales_report.xlsx")
    log.info("Pipeline finished successfully")


if __name__ == "__main__":
    main()
