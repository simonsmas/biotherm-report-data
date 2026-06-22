import json
import sys
from pathlib import Path
from datetime import datetime, timezone
import pandas as pd

CAMPAIGN = "Biotherm"
TIMEZONE_LABEL = "ET"
TIMEZONE = "America/New_York"

EXCEL_FILE_CANDIDATES = [
    "biotherm_daily_total.xlsx",
    "Biotherm_daily_total.xlsx",
    "biotherm_daily_totals.xlsx",
    "Biotherm_daily_totals.xlsx",
    "daily_totals.xlsx",
]

OUTPUT_JSON = "biotherm_totals.json"

VARIANT_ORDER = ["Pulsation Berry", "Citrus Tonic"]

STOCK = {
    "Pulsation Berry": 4700,
    "Citrus Tonic": 5000,
}

def find_input_file():
    for filename in EXCEL_FILE_CANDIDATES:
        path = Path(filename)
        if path.exists():
            return path
    raise FileNotFoundError("Could not find the Biotherm daily totals Excel file.")

def normalise_col(value):
    return str(value).strip().lower().replace("_", " ")

def find_column(df, names):
    lookup = {normalise_col(c): c for c in df.columns}
    for name in names:
        key = normalise_col(name)
        if key in lookup:
            return lookup[key]
    raise ValueError(f"Missing required column. Tried: {', '.join(names)}")

def read_daily_totals(path):
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]

    date_col = find_column(df, ["Date", "Order Date", "Day"])
    variant_cols = {variant: find_column(df, [variant]) for variant in VARIANT_ORDER}

    df = df[df[date_col].notna()].copy()
    if df.empty:
        return pd.DataFrame(columns=["Date"] + VARIANT_ORDER)

    out = pd.DataFrame()
    out["Date"] = pd.to_datetime(df[date_col], errors="coerce")

    valid_mask = out["Date"].notna()
    if not valid_mask.all():
        df = df.loc[valid_mask].copy()
        out = out.loc[valid_mask].copy()

    if out.empty:
        return pd.DataFrame(columns=["Date"] + VARIANT_ORDER)

    for variant, col in variant_cols.items():
        out[variant] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return out.sort_values("Date").reset_index(drop=True)

def pct(value):
    return round(float(value), 1)

def ordinal(n):
    if 10 <= n % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"

def format_us_date(dt):
    return f"{dt.strftime('%B')} {ordinal(int(dt.day))} {dt.year}"

def format_report_updated_label():
    now_utc = datetime.now(timezone.utc)
    now_et = pd.Timestamp(now_utc).tz_convert(TIMEZONE).to_pydatetime()
    return f"{format_us_date(now_et)} {now_et.strftime('%-I:%M %p')} {TIMEZONE_LABEL}"

def build_summary_rows(claimed_by_variant):
    rows = []
    total_available = 0
    total_claimed = 0

    for variant in VARIANT_ORDER:
        available = int(STOCK[variant])
        claimed = int(claimed_by_variant.get(variant, 0))
        remaining = available - claimed
        percent_remaining = pct(remaining / available * 100) if available else 0

        rows.append({
            "variant": variant,
            "available": available,
            "claimed": claimed,
            "remaining": remaining,
            "percent_remaining": percent_remaining,
            "percent_remaining_display": f"{percent_remaining:g}%",
        })

        total_available += available
        total_claimed += claimed

    total_remaining = total_available - total_claimed
    total_percent_remaining = pct(total_remaining / total_available * 100) if total_available else 0

    rows.append({
        "variant": "TOTAL",
        "available": total_available,
        "claimed": total_claimed,
        "remaining": total_remaining,
        "percent_remaining": total_percent_remaining,
        "percent_remaining_display": f"{total_percent_remaining:g}%",
    })

    return rows

def build_json(df):
    report_generated_utc = datetime.now(timezone.utc).isoformat()
    report_updated_label = format_report_updated_label()

    if df.empty:
        claimed = {variant: 0 for variant in VARIANT_ORDER}
        return {
            "campaign": CAMPAIGN,
            "timezone": TIMEZONE_LABEL,
            "last_updated_utc": report_generated_utc,
            "report_updated_at_label": report_updated_label,
            "updated_to": None,
            "includes_orders_up_to_label": "No orders reported yet",
            "rows": build_summary_rows(claimed),
            "redemptions_by_day": {
                "dates": [],
                "date_iso": [],
                "rows": [{"variant": variant, "values": [], "total": 0} for variant in VARIANT_ORDER] + [{"variant": "Total", "values": [], "total": 0}],
                "max_columns_per_table": 20,
            },
            "_debug": {"source_rows": 0, "latest_date": None, "claimed_by_variant": claimed, "status": "empty_or_pre_launch_file"}
        }

    latest_date = df["Date"].max()
    claimed = {variant: int(df[variant].sum()) for variant in VARIANT_ORDER}
    dates = [d.strftime("%m/%d") for d in df["Date"]]
    date_iso = [d.strftime("%Y-%m-%d") for d in df["Date"]]

    redemptions_rows = []
    for variant in VARIANT_ORDER:
        values = [int(v) for v in df[variant].tolist()]
        redemptions_rows.append({"variant": variant, "values": values, "total": int(sum(values))})

    total_values = [int(sum(row[variant] for variant in VARIANT_ORDER)) for _, row in df.iterrows()]
    redemptions_rows.append({"variant": "Total", "values": total_values, "total": int(sum(total_values))})

    return {
        "campaign": CAMPAIGN,
        "timezone": TIMEZONE_LABEL,
        "last_updated_utc": report_generated_utc,
        "report_updated_at_label": report_updated_label,
        "updated_to": latest_date.strftime("%Y-%m-%d"),
        "includes_orders_up_to_label": f"{format_us_date(latest_date)} {TIMEZONE_LABEL}",
        "rows": build_summary_rows(claimed),
        "redemptions_by_day": {
            "dates": dates,
            "date_iso": date_iso,
            "rows": redemptions_rows,
            "max_columns_per_table": 20,
        },
        "_debug": {"source_rows": int(len(df)), "latest_date": latest_date.strftime("%Y-%m-%d"), "claimed_by_variant": claimed, "status": "ok"}
    }

def main():
    output_json = sys.argv[1] if len(sys.argv) > 1 else OUTPUT_JSON
    input_file = find_input_file()
    df = read_daily_totals(input_file)
    data = build_json(df)
    Path(output_json).write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"Created {output_json}")
    print(f"Source file: {input_file}")
    print(f"Rows/days processed: {len(df):,}")
    print(f"Status: {data['_debug']['status']}")
    print(f"Includes orders up to: {data['includes_orders_up_to_label']}")
    print(f"Last updated: {data['report_updated_at_label']}")

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
