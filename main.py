#!/usr/bin/env python3
"""
Capitaline Consolidator

Consolidates multiple Capitaline Excel/CSV files into a single Excel with 3 sheets:
1. Div Adj Close Price (NSE prioritized, fallback BSE)
2. Daily Total Return (%) (NSE prioritized, fallback BSE)
3. Average Marketcap (avg of NSE+BSE)

Default: INFO logs (minimal, user-friendly)
Optional: --debug flag for detailed DEBUG logs
"""

import os, sys, glob, logging, time, argparse
import pandas as pd

# ---------------- Paths ----------------
ASSETS_DIR = os.path.join(os.getcwd(), "assets")
LOGS_DIR = os.path.join(os.getcwd(), "logs")
OUTPUT_XLSX = os.path.join(os.getcwd(), "consolidated_output.xlsx")

# ---------------- Column synonyms ----------------
HEADER_SYNONYMS = {
    "company_name": {"company name", "company", "stock name"},
    "trading_date": {"trading date", "date"},
    "nse_price": {"nse div adj close price", "nse price"},
    "bse_price": {"bse div adj close price", "bse price"},
    "nse_return": {"nse daily total return (%)", "nse return"},
    "bse_return": {"bse daily total return (%)", "bse return"},
    "nse_mcap": {"nse marketcap", "nse market cap"},
    "bse_mcap": {"bse marketcap", "bse market cap"},
}
NEEDED_KEYS = set(HEADER_SYNONYMS.keys())


# ---------------- Logging ----------------
def setup_logger(debug=False):
    """Configure logger with INFO (default) or DEBUG (--debug)."""
    os.makedirs(LOGS_DIR, exist_ok=True)
    log_file = os.path.join(LOGS_DIR, "consolidation.log")

    logger = logging.getLogger("consolidator")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)
    if logger.hasHandlers():
        logger.handlers.clear()

    # Console logs
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)
    ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))

    # File logs (always DEBUG, overwrite each run)
    fh = logging.FileHandler(log_file, mode="w", encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(
        logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
        )
    )

    logger.addHandler(ch)
    logger.addHandler(fh)
    return logger


# ---------------- Helpers ----------------
def normalize_columns(df, fname, logger):
    """Rename messy headers → canonical names."""
    lower_map = {c: c.strip().lower() for c in df.columns}
    mapping = {}
    for canon, variants in HEADER_SYNONYMS.items():
        for orig, low in lower_map.items():
            if low in variants:
                mapping[orig] = canon
                break
    df = df.rename(columns=mapping)
    for key in NEEDED_KEYS:
        if key not in df.columns:
            df[key] = pd.NA
    logger.debug(f"{fname}: normalized columns → {list(df.columns)}")
    return df


# ---------------- Main ----------------
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--debug", action="store_true", help="Enable verbose debug logging"
    )
    args = parser.parse_args()

    global logger
    logger = setup_logger(debug=args.debug)

    start = time.time()

    # Collect input files
    files = []
    for ext in ("*.csv", "*.xlsx", "*.xls"):
        for f in glob.glob(os.path.join(ASSETS_DIR, ext)):
            if not os.path.basename(f).startswith("~$"):  # skip Excel lockfiles
                files.append(f)
    files = sorted(files)

    if not files:
        logger.error(f"No files in {ASSETS_DIR}")
        sys.exit(1)

    logger.info(f"Found {len(files)} files")

    frames = []
    for f in files:
        try:
            logger.info(f"Reading {os.path.basename(f)}")
            if f.endswith((".xlsx", ".xls")):
                raw = pd.read_excel(
                    f, engine="openpyxl", header=1
                )  # real header at row 2
            else:
                raw = pd.read_csv(f, header=1)

            if args.debug:
                logger.debug(f"Preview of {os.path.basename(f)}:\n{raw.head()}")

            df = normalize_columns(raw, os.path.basename(f), logger)

            # Parse dates & numerics
            df["trading_date"] = pd.to_datetime(df["trading_date"], errors="coerce")
            for col in [
                "nse_price",
                "bse_price",
                "nse_return",
                "bse_return",
                "nse_mcap",
                "bse_mcap",
            ]:
                df[col] = pd.to_numeric(df[col], errors="coerce")

            before = len(df)
            df = df[df["trading_date"].notna()]
            logger.info(f"{os.path.basename(f)}: {before} rows → {len(df)} valid")

            frames.append(df)
        except Exception:
            logger.exception(f"Failed reading {f}, skipping")

    if not frames:
        logger.error("No valid data loaded")
        sys.exit(1)

    all_df = pd.concat(frames, ignore_index=True)
    logger.info(
        f"Total merged rows: {len(all_df)} "
        f"(companies={all_df['company_name'].nunique()}, "
        f"dates={all_df['trading_date'].nunique()})"
    )

    # Aggregation (vectorized)
    agg_df = (
        all_df.groupby(["company_name", "trading_date"])
        .agg(
            {
                "nse_price": "last",
                "bse_price": "last",
                "nse_return": "last",
                "bse_return": "last",
                "nse_mcap": "last",
                "bse_mcap": "last",
            }
        )
        .reset_index()
    )

    agg_df["final_price"] = agg_df["nse_price"].where(
        agg_df["nse_price"].notna(), agg_df["bse_price"]
    )
    agg_df["final_return"] = agg_df["nse_return"].where(
        agg_df["nse_return"].notna(), agg_df["bse_return"]
    )
    agg_df["avg_mcap"] = agg_df[["nse_mcap", "bse_mcap"]].mean(axis=1, skipna=True)

    logger.info(
        f"After aggregation: {len(agg_df)} rows, "
        f"{agg_df['company_name'].nunique()} companies, "
        f"{agg_df['trading_date'].nunique()} dates"
    )

    if agg_df.empty:
        logger.error("Final dataframe is empty! Check column mapping.")
        sys.exit(1)

    # Pivot to wide format
    prices = agg_df.pivot(
        index="company_name", columns="trading_date", values="final_price"
    )
    returns = agg_df.pivot(
        index="company_name", columns="trading_date", values="final_return"
    )
    mcaps = agg_df.pivot(
        index="company_name", columns="trading_date", values="avg_mcap"
    )

    # Write Excel
    with pd.ExcelWriter(
        OUTPUT_XLSX, engine="xlsxwriter", datetime_format="yyyy-mm-dd"
    ) as writer:
        prices.to_excel(writer, sheet_name="Div Adj Close Price")
        returns.to_excel(writer, sheet_name="Daily Total Return (%)")
        mcaps.to_excel(writer, sheet_name="Average Marketcap")
        for sheet in writer.sheets.values():
            sheet.freeze_panes(1, 1)
            sheet.set_column(0, 0, 28)

    logger.info(f"✅ File written: {OUTPUT_XLSX}")

    end = time.time()
    logger.info(f"⏱️ Script finished in {end - start:.2f} seconds")


if __name__ == "__main__":
    main()
