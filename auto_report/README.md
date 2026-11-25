# Automated Data Transformation and Reporting

A small Python pipeline that takes a messy raw sales CSV, cleans and transforms it,
and automatically produces a report (Excel workbook + charts + cleaned CSV + log).

## Project structure

```
auto_report/
├── data/raw_sales.csv     # sample raw input (with deliberate data problems)
├── pipeline.py            # the whole pipeline: extract -> transform -> report
├── requirements.txt
└── output/                # created on first run
    ├── cleaned_sales.csv
    ├── sales_report.xlsx
    ├── monthly_revenue.png
    ├── revenue_by_region.png
    └── pipeline.log
```

## How to run

```bash
pip install -r requirements.txt
python pipeline.py
```

Use your own file or output folder:

```bash
python pipeline.py --input data/my_file.csv --output results
```

## What the pipeline does

1. **Extract** - reads the raw CSV.
2. **Transform**
   - trims spaces and standardises capitalisation (`  Asha Rao ` / `SOUTH` -> `Asha Rao` / `South`)
   - parses dates in two formats (`2026-01-05` and `05/01/2026`)
   - removes duplicate orders
   - removes invalid rows (missing customer/quantity/date, negative quantity)
   - adds `revenue` and `month` columns
3. **Report** - builds monthly, region, product and top-customer summaries,
   draws two charts, and saves everything to `sales_report.xlsx`.

## Automating it

Run it on a schedule so the report refreshes by itself.

- **Linux/macOS (cron)**, every day at 8 AM:
  ```
  0 8 * * * cd /path/to/auto_report && python pipeline.py
  ```
- **Windows**: Task Scheduler -> create a task that runs `python pipeline.py`.
