"""
Marketing Analytics Toolkit

Computes channel-level KPIs (CVR, CPC, CAC, AOV, ROAS) from Excel 'sessions' and
'orders' sheets, produces summary tables/charts, and fits a linear
Revenue ~ Spend model to approximate marginal ROAS. Designed for reproducible analysis and CLI use.
"""

import argparse
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# --- I/O ----------------------------------------------------------------------

def load_data(xlsx_path: str):
    """Read required sheets from a workbook."""
    sessions = pd.read_excel(xlsx_path, sheet_name="sessions")
    orders = pd.read_excel(xlsx_path, sheet_name="orders")
    return sessions, orders


# --- Cleaning -----------------------------------------------------------------

def clean_sessions(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise session fields: dates, numeric types, and string trimming."""
    df = df.copy()
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"]).dt.date
    for c in ["Sessions", "Clicks", "Spend_gbp", "Adj_spend_gbp", "CTR", "CPC"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["Channel", "Source", "Medium", "Campaign", "Department"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df


def clean_orders(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise order fields: dates, numeric types, and channel labels."""
    df = df.copy()
    if "Order_date" in df.columns:
        df["Order_date"] = pd.to_datetime(df["Order_date"]).dt.date
        df["Date"] = df["Order_date"]
    if "Revenue_gbp" in df.columns:
        df["Revenue_gbp"] = pd.to_numeric(df["Revenue_gbp"], errors="coerce")
    if "Items" in df.columns:
        df["Items"] = pd.to_numeric(df["Items"], errors="coerce")
    for c in ["Channel_ft", "Channel"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df


# --- KPI assembly -------------------------------------------------------------

def compute_daily_channel_kpis(sessions: pd.DataFrame, orders: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate sessions and orders to Date × Channel, join them,
    and compute KPI columns.
    """
    s = sessions.copy()
    s["Spend"] = s.get("Adj_spend_gbp", s.get("Spend_gbp", 0.0)).fillna(0.0)
    s["Date"] = pd.to_datetime(s["Date"]).dt.date
    s_agg = (
        s.groupby(["Date", "Channel"], as_index=False)
         .agg(Sessions=("Sessions", "sum"),
              Clicks=("Clicks", "sum"),
              Spend=("Spend", "sum"))
    )

    o = orders.copy()
    chan_col = "Channel_ft" if "Channel_ft" in o.columns else ("Channel" if "Channel" in o.columns else None)
    if chan_col is None:
        o["Channel_ft"] = "All"
        chan_col = "Channel_ft"
    if "Date" not in o.columns:
        o["Date"] = pd.to_datetime(o["Order_date"]).dt.date

    o_agg = (
        o.groupby(["Date", chan_col], as_index=False)
         .agg(Orders=("Order_id", "nunique"),
              Revenue=("Revenue_gbp", "sum"))
         .rename(columns={chan_col: "Channel"})
    )

    df = (
        pd.merge(s_agg, o_agg, on=["Date", "Channel"], how="left")
          .fillna({"Orders": 0, "Revenue": 0.0})
    )

    # KPI definitions
    df["CVR"]      = np.where(df["Clicks"] > 0, df["Orders"] / df["Clicks"], np.nan)
    df["CPC_calc"] = np.where(df["Clicks"] > 0, df["Spend"]  / df["Clicks"], np.nan)
    df["CAC"]      = np.where(df["Orders"] > 0, df["Spend"]  / df["Orders"], np.nan)
    df["AOV"]      = np.where(df["Orders"] > 0, df["Revenue"]/ df["Orders"], np.nan)
    df["ROAS"]     = np.where(df["Spend"]  > 0, df["Revenue"]/ df["Spend"],  np.nan)
    return df


def channel_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse to channel totals and recompute KPIs on the aggregated values
    (avoids averaging ratios).
    """
    agg = (
        df.groupby("Channel", as_index=False)
          .agg(Sessions=("Sessions", "sum"),
               Clicks=("Clicks", "sum"),
               Spend=("Spend", "sum"),
               Orders=("Orders", "sum"),
               Revenue=("Revenue", "sum"))
    )
    agg["CVR"]  = np.where(agg["Clicks"] > 0, agg["Orders"]  / agg["Clicks"], np.nan)
    agg["CPC"]  = np.where(agg["Clicks"] > 0, agg["Spend"]   / agg["Clicks"], np.nan)
    agg["CAC"]  = np.where(agg["Orders"] > 0, agg["Spend"]   / agg["Orders"], np.nan)
    agg["AOV"]  = np.where(agg["Orders"] > 0, agg["Revenue"] / agg["Orders"], np.nan)
    agg["ROAS"] = np.where(agg["Spend"]  > 0, agg["Revenue"] / agg["Spend"],  np.nan)
    return agg.sort_values("ROAS", ascending=False)


# --- Visualisation ------------------------------------------------------------

def plot_roas_by_channel(summary_df: pd.DataFrame, out_dir: str):
    """Bar chart for ROAS by channel."""
    os.makedirs(out_dir, exist_ok=True)
    plt.figure()
    plt.bar(summary_df["Channel"], summary_df["ROAS"])
    plt.title("ROAS by Channel")
    plt.ylabel("ROAS (Revenue / Spend)")
    plt.xticks(rotation=30, ha="right")
    out = os.path.join(out_dir, "roas_by_channel.png")
    plt.tight_layout()
    plt.savefig(out, dpi=120, bbox_inches="tight")
    return out


def plot_revenue_vs_spend_over_time(daily_df: pd.DataFrame, out_dir: str):
    """Line chart of total Revenue and Spend by day."""
    os.makedirs(out_dir, exist_ok=True)
    ts = (
        daily_df.groupby("Date", as_index=False)
                .agg(Revenue=("Revenue", "sum"), Spend=("Spend", "sum"))
                .sort_values("Date")
    )
    plt.figure()
    plt.plot(ts["Date"], ts["Revenue"], label="Revenue")
    plt.plot(ts["Date"], ts["Spend"], label="Spend")
    plt.title("Revenue vs Spend Over Time (Daily)")
    plt.ylabel("GBP")
    plt.xlabel("Date")
    plt.legend()
    out = os.path.join(out_dir, "revenue_vs_spend_over_time.png")
    plt.tight_layout()
    plt.savefig(out, dpi=120, bbox_inches="tight")
    return out


# --- Simple response curve ----------------------------------------------------

def simple_response_curve(daily_df: pd.DataFrame, out_dir: str):
    """
    Fit a linear model Revenue ≈ a + b·Spend on daily totals.
    The slope b approximates marginal ROAS over the observed range.
    """
    os.makedirs(out_dir, exist_ok=True)
    df = (
        daily_df.groupby("Date", as_index=False)
                .agg(Revenue=("Revenue", "sum"), Spend=("Spend", "sum"))
                .dropna(subset=["Revenue", "Spend"])
    )
    if len(df) < 2 or df["Spend"].sum() == 0:
        return None, None, None

    b, a = np.polyfit(df["Spend"].values, df["Revenue"].values, 1)

    xs = np.linspace(df["Spend"].min(), df["Spend"].max(), 100)
    ys = a + b * xs

    plt.figure()
    plt.scatter(df["Spend"], df["Revenue"], alpha=0.7)
    plt.plot(xs, ys, linewidth=2)
    plt.title("Revenue vs Spend (linear fit)")
    plt.xlabel("Spend (GBP)")
    plt.ylabel("Revenue (GBP)")
    out = os.path.join(out_dir, "revenue_vs_spend_fit.png")
    plt.tight_layout()
    plt.savefig(out, dpi=120, bbox_inches="tight")
    return float(a), float(b), out


# --- CLI entry point ----------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", required=True, help="Path to the Excel file (with 'sessions' and 'orders' sheets)")
    parser.add_argument("--out", default="outputs", help="Directory to save outputs (tables and charts)")
    args = parser.parse_args()

    sessions_raw, orders_raw = load_data(args.xlsx)
    sessions = clean_sessions(sessions_raw)
    orders = clean_orders(orders_raw)

    daily = compute_daily_channel_kpis(sessions, orders)
    chan = channel_summary(daily)

    # Persist table
    os.makedirs(args.out, exist_ok=True)
    chan_out = chan.copy()
    for c in ["CVR", "CPC", "CAC", "AOV", "ROAS"]:
        if c in chan_out.columns:
            chan_out[c] = chan_out[c].round(3)
    chan_path = os.path.join(args.out, "channel_kpis.csv")
    chan_out.to_csv(chan_path, index=False)

    # Persist charts
    roas_plot_path = plot_roas_by_channel(chan_out, args.out)
    timeseries_plot_path = plot_revenue_vs_spend_over_time(daily, args.out)
    a, b, fit_plot_path = simple_response_curve(daily, args.out)

    # Lightweight run log
    print("Saved:")
    print(" -", chan_path)
    print(" -", roas_plot_path)
    print(" -", timeseries_plot_path)
    if a is not None:
        print(" -", fit_plot_path)
        print(f"Linear model: Revenue ≈ {a:.2f} + {b:.3f} · Spend (b ≈ marginal ROAS)")


if __name__ == "__main__":
    main()
