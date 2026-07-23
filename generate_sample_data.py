"""
Generates sample Azure Cost Management-style export files for demo/testing purposes.
Mimics the columns found in a real Azure Cost Management "Cost Details" export.
Run once to populate the data/ folder with 4 months of sample data, then delete
these sample files and replace them with your real monthly exports.
"""
import os
import random
from datetime import date, timedelta

import pandas as pd

random.seed(42)

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
os.makedirs(DATA_DIR, exist_ok=True)

SUBSCRIPTION = "Enterprise Subscription 1"

RESOURCE_GROUPS = ["rg-prod-app", "rg-prod-data", "rg-shared-network", "rg-dev-test"]

# (ServiceName, MeterCategory, base daily cost, growth per month, volatility)
SERVICES = [
    ("Virtual Machines", "Compute", 42.0, 0.06, 0.15),
    ("Storage", "Storage", 11.0, 0.03, 0.10),
    ("Azure SQL Database", "Databases", 28.0, 0.08, 0.12),
    ("App Service", "Web", 9.0, 0.02, 0.10),
    ("Azure Kubernetes Service", "Compute", 19.0, 0.10, 0.18),
    ("Networking", "Networking", 6.0, 0.01, 0.08),
    ("Backup", "Storage", 3.0, 0.02, 0.05),
    ("Monitor", "Analytics", 2.5, 0.04, 0.08),
]

MONTHS = [(2026, 3), (2026, 4), (2026, 5), (2026, 6)]


def month_days(year, month):
    start = date(year, month, 1)
    if month == 12:
        end = date(year + 1, 1, 1)
    else:
        end = date(year, month + 1, 1)
    d = start
    while d < end:
        yield d
        d += timedelta(days=1)


for month_index, (year, month) in enumerate(MONTHS):
    rows = []
    for svc_name, meter_cat, base_cost, growth, volatility in SERVICES:
        trended_base = base_cost * ((1 + growth) ** month_index)
        for day in month_days(year, month):
            rg = random.choice(RESOURCE_GROUPS)
            noise = random.uniform(1 - volatility, 1 + volatility)
            cost = round(trended_base / len(RESOURCE_GROUPS) * noise, 2)
            # skip some rows randomly to mimic real sparse usage data
            if random.random() < 0.05:
                continue
            rows.append(
                {
                    "UsageDateTime": day.strftime("%Y-%m-%d"),
                    "SubscriptionName": SUBSCRIPTION,
                    "ResourceGroupName": rg,
                    "ServiceName": svc_name,
                    "MeterCategory": meter_cat,
                    "Cost": cost,
                    "Currency": "USD",
                }
            )

    df = pd.DataFrame(rows)
    out_path = os.path.join(DATA_DIR, f"AzureCost_{year}-{month:02d}.xlsx")
    df.to_excel(out_path, index=False, sheet_name="Cost Details")
    print(f"Wrote {out_path} ({len(df)} rows, total cost {df['Cost'].sum():.2f})")

print("\nSample data generated. Replace files in data/ with your real monthly exports when ready.")
