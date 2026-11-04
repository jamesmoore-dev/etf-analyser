"""
ETF Analyser and Portfolio Builder
----------------------------------
Reads a list of funds from a spreadsheet (funds.xlsx), downloads their price
history, calculates key performance and risk numbers, and saves:
  etf_results.xlsx           the funds grouped under your headings, a flat sortable
                             sheet, a sector summary, a correlation table and charts
  etf_portfolio_builder.html an interactive page (open it in your browser): pick
                             funds, set weights, add money and see how it would have
                             grown, plus the sector summary and correlations
  growth_chart.png and risk_return_chart.png

The fund list lives in funds.xlsx, not in this script. The first time you run it,
a starter funds.xlsx is created containing Global X's 53 funds (Fund Map, July
2026). Edit that spreadsheet, or make another one, to analyse any funds you like.

Usage examples:
    python3 etf_analyser.py                          # uses funds.xlsx, last 5 years
    python3 etf_analyser.py --funds mylist.xlsx --out mylist   # a different list
    python3 etf_analyser.py --period 3y              # change the time window
    python3 etf_analyser.py --tickers GXAI SEMI TECH # only analyse specific funds
    python3 etf_analyser.py --chart-top 0            # growth chart shows every fund
    python3 etf_analyser.py --common-start           # chart all funds from the same start date

Note: prices are "adjusted" (they include reinvested distributions), so the
returns are total returns, not just price changes. Past performance does not
predict future performance, and this is not financial advice.
"""

import argparse
import json
import os
import sys
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. THE FUND LIST
# ---------------------------------------------------------------------------
# The funds to analyse are read from a spreadsheet (see load_funds below). This
# built-in list is only used to create the starter spreadsheet the first time you
# run the script. It is Global X's Fund Map (July 2026). Each entry is:
#     (heading, sub-heading, [(ticker, name), ...])
DEFAULT_FUND_MAP = [
    ("Thematic Growth", "Disruptive Technology", [
        ("ACDC", "Battery Tech & Lithium ETF"),
        ("BUGG", "Cybersecurity ETF"),
        ("DTEC", "Defence Tech ETF"),
        ("FTEC", "Fintech & Blockchain ETF"),
        ("GXAI", "Artificial Intelligence ETF"),
        ("HMND", "Humanoid Robotics ETF"),
        ("MOON", "Space Tech ETF"),
        ("ROBO", "ROBO Global Robotics & Automation ETF"),
        ("SEMI", "Semiconductor ETF"),
        ("TECH", "Morningstar Global Technology ETF"),
    ]),
    ("Thematic Growth", "Multi-Theme", [
        ("DRGN", "China Tech ETF"),
        ("FANG", "FANG+ ETF"),
        ("FHNG", "FANG+ (Currency Hedged) ETF"),
    ]),
    ("Thematic Growth", "People & Demographics", [
        ("CURE", "S&P Biotech ETF"),
    ]),
    ("Thematic Growth", "Physical Environment & Infrastructure", [
        ("AINF", "Artificial Intelligence Infrastructure ETF"),
        ("HGEN", "Hydrogen ETF"),
        ("PAVE", "US Infrastructure Development ETF"),
    ]),
    ("Commodities", "Miners", [
        ("ATOM", "Uranium ETF"),
        ("GMTL", "Rare Earth and Critical Metals ETF"),
        ("SLVM", "Silver Miners ETF"),
        ("WIRE", "Copper Miners ETF"),
    ]),
    ("Commodities", "Synthetic", [
        ("BCOM", "Bloomberg Commodity Complex ETF"),
    ]),
    ("Commodities", "Physical", [
        ("ETPMAG", "Physical Silver Structured"),
        ("ETPMPD", "Physical Palladium Structured"),
        ("ETPMPM", "Physical Precious Metals Basket Structured"),
        ("ETPMPT", "Physical Platinum Structured"),
        ("GHLD", "Gold Bullion (Currency Hedged) ETF"),
        ("GOLD", "Physical Gold Structured"),
        ("GXLD", "Gold Bullion ETF"),
    ]),
    ("Income", "Fixed Income", [
        ("BANK", "Australian Bank Credit ETF"),
        ("USHY", "USD High Yield Bond (Currency Hedged) ETF"),
        ("USIG", "USD Corporate Bond (Currency Hedged) ETF"),
        ("USTB", "US Treasury Bond (Currency Hedged) ETF"),
    ]),
    ("Income", "Covered Call", [
        ("AYLD", "S&P/ASX 200 Covered Call Complex ETF"),
        ("QYLD", "Nasdaq 100 Covered Call Complex ETF"),
        ("UYLD", "S&P 500 Covered Call Complex ETF"),
    ]),
    ("Income", "Dividend", [
        ("ZYAU", "S&P/ASX 200 High Dividend ETF"),
        ("ZYUS", "S&P 500 High Yield Low Volatility ETF"),
    ]),
    ("Core", "", [
        ("A300", "Australia 300 ETF"),
        ("GARP", "S&P World ex Australia GARP ETF"),
        ("GHRP", "S&P World ex Australia GARP (Currency Hedged) ETF"),
        ("GRPA", "S&P Australia GARP ETF"),
        ("ISMD", "MSCI International Small and Mid Cap ETF"),
        ("OZXX", "Australia ex Financials & Resources ETF"),
        ("RSSL", "Russell 2000 ETF"),
        ("U100", "US 100 ETF"),
    ]),
    ("International Access", "Regional", [
        ("ESTX", "EURO STOXX 50 ETF"),
    ]),
    ("International Access", "Single Country", [
        ("J100", "Japan TOPIX 100 ETF"),
        ("NDIA", "India Nifty 50 ETF"),
    ]),
    ("Digital Assets", "", [
        ("EBTC", "21Shares Bitcoin ETF"),
        ("EETH", "21Shares Ethereum ETF"),
    ]),
    ("Leveraged & Inverse", "", [
        ("LNAS", "Ultra Long Nasdaq 100 Complex ETF"),
        ("SNAS", "Ultra Short Nasdaq 100 Complex ETF"),
    ]),
]

# Funds listed on TMX Australia (the exchange formerly known as Cboe Australia).
# Only used when creating the starter spreadsheet; everything else is on the ASX.
DEFAULT_TMX_AUSTRALIA = {"EBTC", "EETH", "FTEC"}

# The fund list currently in use. These are filled in by load_funds().
FUND_MAP = []         # [(heading, sub-heading, [(ticker, name), ...]), ...]
FUNDS = {}            # ticker -> name
SECTOR_OF = {}        # ticker -> (heading, sub-heading)
EXCHANGE_OF = {}      # ticker -> exchange label shown in the results
YAHOO_SYMBOL_OF = {}  # ticker -> the symbol Yahoo Finance uses

# What you can write in the spreadsheet's Exchange column, and the ending Yahoo
# Finance adds for it. For any other exchange, fill in the "Yahoo symbol" column.
EXCHANGES = {
    "ASX": ("ASX", ".AX"),                       # e.g. GXAI -> GXAI.AX
    "TMX AUSTRALIA": ("TMX Australia", ".XA"),   # e.g. EBTC -> EBTC.XA
    "TMX": ("TMX Australia", ".XA"),
    "CBOE": ("TMX Australia", ".XA"),            # Cboe Australia is now TMX Australia
    "CBOE AUSTRALIA": ("TMX Australia", ".XA"),
    "US": ("US", ""),                            # NYSE / Nasdaq, e.g. AAPL
}

# Used when calculating the return-per-unit-of-risk ratio. Edit to taste.
RISK_FREE_RATE = 0.04  # 4% a year

TRADING_DAYS_PER_YEAR = 252

NUMBER_COLUMNS = ["Years", "Latest price", "Total return %", "Annual return %",
                  "Volatility %", "Max drawdown %", "Return per risk"]


# ---------------------------------------------------------------------------
# 1b. LOADING THE FUND LIST FROM A SPREADSHEET
# ---------------------------------------------------------------------------
FUNDS_COLUMNS = ["Ticker", "Name", "Heading", "Sub-heading", "Exchange", "Yahoo symbol"]

HOW_TO_USE = [
    "How to use this spreadsheet",
    "",
    "Each row on the Funds sheet is one fund to analyse. Add, remove or edit rows, "
    "save the file, then run the script again.",
    "",
    "Ticker (required): the fund's code, for example VAS. Write it without any ending like .AX.",
    "Name (optional): only used for display.",
    "Heading and Sub-heading (optional): how funds are grouped in the results. Make up your "
    "own groups. Funds with no heading go under 'Other'.",
    "Exchange (optional): ASX, TMX Australia or US. Leave it blank for ASX.",
    "Yahoo symbol (optional): only needed for a fund on any other exchange. Type the exact "
    "symbol shown on Yahoo Finance (for example VOD.L).",
    "",
    "Please keep the column headings in row 1 as they are.",
    "If a fund shows 'No price data found' in the results, check its ticker and exchange "
    "on finance.yahoo.com.",
    "",
    "To keep several lists, save each as its own file and run, for example:",
    "python3 etf_analyser.py --funds mylist.xlsx --out mylist",
]


def create_starter_file(path):
    """Write a starter spreadsheet containing the built-in Global X fund list."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Funds"
    for col, name in enumerate(FUNDS_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.fill = COLUMN_HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF")
    row = 2
    for heading, sub, funds in DEFAULT_FUND_MAP:
        for ticker, name in funds:
            exchange = "TMX Australia" if ticker in DEFAULT_TMX_AUSTRALIA else "ASX"
            for col, value in enumerate([ticker, name, heading, sub, exchange, ""], start=1):
                ws.cell(row=row, column=col, value=value or None)
            row += 1
    for col, width in enumerate([10, 48, 24, 38, 16, 16], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"

    notes = wb.create_sheet("How to use")
    for i, line in enumerate(HOW_TO_USE, start=1):
        notes.cell(row=i, column=1, value=line or None)
    notes["A1"].font = Font(bold=True, size=13)
    notes.column_dimensions["A"].width = 110
    for i in range(1, len(HOW_TO_USE) + 1):
        notes.cell(row=i, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)


def load_funds(path):
    """Read the fund list from a spreadsheet (.xlsx) or a CSV and set it up for use."""
    global FUND_MAP, FUNDS, SECTOR_OF, EXCHANGE_OF, YAHOO_SYMBOL_OF
    try:
        if os.path.splitext(path)[1].lower() == ".csv":
            raw = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
        else:
            raw = pd.read_excel(path, sheet_name=0, dtype=str)
    except Exception as err:
        sys.exit(f"Couldn't read {path}: {err}")

    # Be forgiving about how the column headings are written
    raw.columns = [str(c).strip().lower().replace("-", " ").replace("_", " ")
                   for c in raw.columns]
    raw = raw.rename(columns={
        "code": "ticker", "symbol": "ticker", "asx code": "ticker", "fund": "name",
        "category": "heading", "sector": "heading", "sub heading": "subheading",
        "sub category": "subheading", "subcategory": "subheading", "yahoo": "yahoo symbol"})
    if "ticker" not in raw.columns:
        sys.exit(f"{path} needs a column headed 'Ticker'. The columns found were: "
                 f"{list(raw.columns)}")

    def cell(row, column):
        value = row.get(column, "")
        return "" if pd.isna(value) else str(value).strip()

    rows, seen = [], set()
    for _, row in raw.iterrows():
        ticker = cell(row, "ticker").upper()
        if not ticker:
            continue
        if ticker in seen:
            print(f"Note: {ticker} is listed more than once in {path}; using the first.")
            continue
        yahoo, exchange = cell(row, "yahoo symbol"), cell(row, "exchange")
        if yahoo:
            symbol, label = yahoo.upper(), exchange or "Other"
        else:
            key = (exchange or "ASX").upper()
            if key not in EXCHANGES:
                print(f"Note: skipped {ticker} - unknown exchange '{exchange}'. Use ASX, "
                      "TMX Australia or US, or fill in the 'Yahoo symbol' column.")
                continue
            label, suffix = EXCHANGES[key]
            symbol = ticker + suffix
        seen.add(ticker)
        rows.append((ticker, cell(row, "name"), cell(row, "heading") or "Other",
                     cell(row, "subheading"), label, symbol))
    if not rows:
        sys.exit(f"No funds found in {path}. Put at least one ticker under the 'Ticker' heading.")

    # Group the funds by heading and sub-heading, keeping the order you wrote them in
    fund_map = []
    for heading in dict.fromkeys(r[2] for r in rows):
        for sub in dict.fromkeys(r[3] for r in rows if r[2] == heading):
            fund_map.append((heading, sub, [(r[0], r[1]) for r in rows
                                            if r[2] == heading and r[3] == sub]))
    FUND_MAP = fund_map
    FUNDS = {t: n for _, _, funds in fund_map for t, n in funds}
    SECTOR_OF = {r[0]: (r[2], r[3]) for r in rows}
    EXCHANGE_OF = {r[0]: r[4] for r in rows}
    YAHOO_SYMBOL_OF = {r[0]: r[5] for r in rows}


# ---------------------------------------------------------------------------
# 2. DOWNLOADING THE DATA
# ---------------------------------------------------------------------------
def exchange_of(code):
    """Which exchange a fund is listed on."""
    return EXCHANGE_OF[code]


def yahoo_symbol(code):
    """The symbol Yahoo uses for a fund, e.g. GXAI -> GXAI.AX, EBTC -> EBTC.XA"""
    return YAHOO_SYMBOL_OF[code]


def download_prices(codes, period):
    """Return a DataFrame of adjusted closing prices: one column per ETF."""
    symbol_to_code = {yahoo_symbol(c): c for c in codes}
    yahoo_codes = list(symbol_to_code)
    raw = yf.download(
        yahoo_codes,
        period=period,
        auto_adjust=True,  # adjusts prices for distributions -> total return
        progress=False,
    )
    if raw.empty:
        return pd.DataFrame()

    prices = raw["Close"]
    if isinstance(prices, pd.Series):  # happens when only one ETF is requested
        prices = prices.to_frame(name=yahoo_codes[0])

    # Turn Yahoo symbols (e.g. GXAI.AX) back into plain ETF codes (GXAI)
    prices.columns = [symbol_to_code.get(str(c), str(c)) for c in prices.columns]
    prices = prices.dropna(axis=1, how="all")  # drop ETFs with no data at all
    return prices


# ---------------------------------------------------------------------------
# 3. THE MATHS
# ---------------------------------------------------------------------------
def compute_metrics(prices, codes):
    """
    Turn a table of prices into a table of performance and risk numbers.
    Every code gets a row, even if there is no usable data, so no fund
    ever goes missing from the results.
    """
    nan = float("nan")
    rows = []
    for code in codes:
        heading, sub = SECTOR_OF[code]
        row = {
            "Code": code,
            "Name": FUNDS[code],
            "Category": heading,
            "Sub-category": sub,
            "Exchange": exchange_of(code),
            "Data from": None,
            "Years": nan,
            "Latest price": nan,
            "Total return %": nan,
            "Annual return %": nan,
            "Volatility %": nan,
            "Max drawdown %": nan,
            "Return per risk": nan,
            "Status": "No price data found",
        }

        series = prices[code].dropna() if code in prices.columns else pd.Series(dtype=float)
        if 0 < len(series) < 30:
            row["Status"] = "Not enough history (under 30 trading days)"
        elif len(series) >= 30:
            daily_returns = series.pct_change(fill_method=None).dropna()
            years = (series.index[-1] - series.index[0]).days / 365.25
            total_return = series.iloc[-1] / series.iloc[0] - 1

            # Annualised numbers are only meaningful with at least a year of data
            cagr = (1 + total_return) ** (1 / years) - 1 if years >= 1 else nan

            # Volatility = how much the price swings, scaled up to a yearly figure
            volatility = daily_returns.std() * (TRADING_DAYS_PER_YEAR ** 0.5)

            # Max drawdown = biggest fall from a previous peak to a later low
            max_drawdown = (series / series.cummax() - 1).min()

            # Return per unit of risk (a Sharpe-style ratio)
            if years >= 1 and volatility > 0:
                return_per_risk = (cagr - RISK_FREE_RATE) / volatility
            else:
                return_per_risk = nan

            row.update({
                "Data from": series.index[0].date(),
                "Years": round(years, 2),
                "Latest price": round(series.iloc[-1], 2),
                "Total return %": round(total_return * 100, 1),
                "Annual return %": round(cagr * 100, 1),
                "Volatility %": round(volatility * 100, 1),
                "Max drawdown %": round(max_drawdown * 100, 1),
                "Return per risk": round(return_per_risk, 2),
                "Status": "OK" if years >= 1
                          else "Under 1 year of history - no annualised figures",
            })
        rows.append(row)

    return pd.DataFrame(rows).set_index("Code")


# ---------------------------------------------------------------------------
# 4. THE CHARTS
# ---------------------------------------------------------------------------
def plot_growth(prices, common_start, filename, limited=False):
    """Line chart: what $10,000 invested at the start would have become."""
    data = prices.copy()
    if common_start:
        data = data.dropna()  # keep only dates where every ETF has a price

    fig, ax = plt.subplots(figsize=(14, 7))
    for code in data.columns:
        series = data[code].dropna()
        if series.empty:
            continue
        ax.plot(series.index, series / series.iloc[0] * 10_000, label=code, linewidth=1)

    title = "Growth of $10,000 (distributions reinvested)"
    if limited:
        title += f" - top {len(data.columns)} by annual return"
    ax.set_title(title)
    ax.set_ylabel("Value (AUD)")
    # Legend sits outside the chart so it doesn't cover the lines
    ax.legend(loc="center left", bbox_to_anchor=(1.01, 0.5),
              ncol=1 if len(data.columns) <= 15 else 3, fontsize=8)
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(filename, dpi=150, bbox_inches="tight")
    plt.close(fig)


def plot_risk_return(metrics, filename):
    """Scatter plot: each ETF's annual return against its volatility."""
    usable = metrics.dropna(subset=["Annual return %"])
    if usable.empty:
        return False

    fig, ax = plt.subplots(figsize=(9, 6))
    ax.scatter(usable["Volatility %"], usable["Annual return %"])
    for code, row in usable.iterrows():
        ax.annotate(code, (row["Volatility %"], row["Annual return %"]),
                    textcoords="offset points", xytext=(5, 4), fontsize=8)

    ax.axhline(0, color="grey", linewidth=0.8)
    ax.set_title("Risk vs return")
    ax.set_xlabel("Volatility (% a year) - further right = more risk")
    ax.set_ylabel("Annual return (%) - higher = better")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(filename, dpi=150)
    plt.close(fig)
    return True


# ---------------------------------------------------------------------------
# 4b. SECTOR SUMMARY AND CORRELATIONS
# ---------------------------------------------------------------------------
def compute_sector_summary(metrics):
    """
    One row per Fund Map heading (and per sub-heading), plus a total row.
    These are simple averages across the funds in each group, using only funds
    with at least a year of history. They are not the return of a sector index.
    """
    def summarise(level, name, group):
        ok = group.dropna(subset=["Annual return %"])
        row = {"Level": level, "Sector": name, "Funds": len(group),
               "Funds with 1y+ data": len(ok),
               "Avg annual return %": None, "Median annual return %": None,
               "Avg volatility %": None, "Avg max drawdown %": None,
               "Best fund": "", "Best annual return %": None,
               "Worst fund": "", "Worst annual return %": None}
        if not ok.empty:
            best = ok["Annual return %"].idxmax()
            worst = ok["Annual return %"].idxmin()
            row.update({
                "Avg annual return %": round(ok["Annual return %"].mean(), 1),
                "Median annual return %": round(ok["Annual return %"].median(), 1),
                "Avg volatility %": round(ok["Volatility %"].mean(), 1),
                "Avg max drawdown %": round(ok["Max drawdown %"].mean(), 1),
                "Best fund": best,
                "Best annual return %": ok.loc[best, "Annual return %"],
                "Worst fund": worst,
                "Worst annual return %": ok.loc[worst, "Annual return %"],
            })
        return row

    rows = []
    for heading in dict.fromkeys(h for h, _, _ in FUND_MAP):
        group = metrics[metrics["Category"] == heading]
        if group.empty:
            continue
        rows.append(summarise("heading", heading, group))
        subs = [s for s in dict.fromkeys(s for h, s, _ in FUND_MAP if h == heading) if s]
        for sub in subs:
            in_sub = group[group["Sub-category"] == sub]
            if not in_sub.empty:
                rows.append(summarise("sub", sub, in_sub))
    rows.append(summarise("total", "All funds", metrics))
    return pd.DataFrame(rows)


def compute_correlations(prices, min_weeks=52):
    """
    How closely each pair of funds moves together (+1 = in step, 0 = unrelated,
    -1 = opposite), using WEEKLY returns. Weekly is used because funds that hold
    overseas shares trade in different time zones to Australian ones, which
    blurs daily comparisons. Funds with under a year of data are left out.
    """
    weekly = prices.ffill().resample("W-FRI").last()
    returns = weekly.pct_change(fill_method=None)
    corr = returns.corr(min_periods=min_weeks)
    keep = [c for c in FUNDS if c in corr.columns and corr[c].notna().any()]
    return corr.loc[keep, keep]


# ---------------------------------------------------------------------------
# 5. SAVING TO EXCEL
# ---------------------------------------------------------------------------
SECTOR_COLUMNS = ["Code", "Name", "Exchange", "Data from", "Years", "Latest price",
                  "Total return %", "Annual return %", "Volatility %",
                  "Max drawdown %", "Return per risk", "Status"]
FLAT_COLUMNS = ["Code", "Name", "Category", "Sub-category", "Exchange", "Data from",
                "Years", "Latest price", "Total return %", "Annual return %",
                "Volatility %", "Max drawdown %", "Return per risk", "Status"]

NUMBER_FORMATS = {
    "Data from": "dd mmm yyyy",
    "Years": "0.00",
    "Latest price": "0.00",
    "Total return %": "0.0;[Red]-0.0",
    "Annual return %": "0.0;[Red]-0.0",
    "Volatility %": "0.0",
    "Max drawdown %": "0.0;[Red]-0.0",
    "Return per risk": "0.00;[Red]-0.00",
}
COLUMN_WIDTHS = {
    "Code": 10, "Name": 48, "Category": 22, "Sub-category": 36, "Exchange": 15,
    "Data from": 13, "Years": 8, "Latest price": 12, "Total return %": 15,
    "Annual return %": 16, "Volatility %": 13, "Max drawdown %": 16,
    "Return per risk": 16, "Status": 50,
}

# Colours borrowed from the Fund Map
HEADING_FILL = PatternFill("solid", fgColor="0F3D4A")      # dark teal
SUBHEADING_FILL = PatternFill("solid", fgColor="E3E7E9")   # light grey
COLUMN_HEADER_FILL = PatternFill("solid", fgColor="E8562A")  # orange


def _band(ws, row, ncols, text, fill, font):
    """Fill a whole row with a colour and put a label in the first cell."""
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
    ws.cell(row=row, column=1, value=text)


def _write_row(ws, row, columns, values):
    for col, (name, value) in enumerate(zip(columns, values), start=1):
        if hasattr(value, "item"):   # numpy number -> plain Python number
            value = value.item()
        if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
            value = None
        cell = ws.cell(row=row, column=col, value=value)
        if name in NUMBER_FORMATS:
            cell.number_format = NUMBER_FORMATS[name]


def _setup_sheet(ws, columns, as_of):
    ws["A1"] = as_of
    ws["A1"].font = Font(italic=True)
    for col, name in enumerate(columns, start=1):
        cell = ws.cell(row=3, column=col, value=name)
        cell.fill = COLUMN_HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF")
        ws.column_dimensions[cell.column_letter].width = COLUMN_WIDTHS[name]
    ws.freeze_panes = "C4"   # keeps the header, Code and Name in view when scrolling


SUMMARY_COLUMNS = ["Sector", "Funds", "Funds with 1y+ data", "Avg annual return %",
                   "Median annual return %", "Avg volatility %", "Avg max drawdown %",
                   "Best fund", "Best annual return %", "Worst fund",
                   "Worst annual return %"]
for _name in ("Avg annual return %", "Median annual return %", "Avg max drawdown %",
              "Best annual return %", "Worst annual return %"):
    NUMBER_FORMATS[_name] = "0.0;[Red]-0.0"
NUMBER_FORMATS["Avg volatility %"] = "0.0"


def _write_sector_sheet(ws, sector_df, as_of):
    ws["A1"] = ("Simple averages across each group's funds that have 1+ year of data "
                "(not a sector index). " + as_of)
    ws["A1"].font = Font(italic=True)
    for col, name in enumerate(SUMMARY_COLUMNS, start=1):
        cell = ws.cell(row=3, column=col, value=name)
        cell.fill = COLUMN_HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[cell.column_letter].width = 36 if name == "Sector" else 14
    ws.row_dimensions[3].height = 32
    for i, rec in enumerate(sector_df.to_dict("records"), start=4):
        _write_row(ws, i, SUMMARY_COLUMNS, [rec[c] for c in SUMMARY_COLUMNS])
        if rec["Level"] == "heading":
            for col in range(1, len(SUMMARY_COLUMNS) + 1):
                ws.cell(row=i, column=col).fill = HEADING_FILL
                ws.cell(row=i, column=col).font = Font(bold=True, color="FFFFFF")
            ws.cell(row=i, column=1).value = rec["Sector"].upper()
        elif rec["Level"] == "sub":
            ws.cell(row=i, column=1).alignment = Alignment(indent=2)
        else:  # total row
            for col in range(1, len(SUMMARY_COLUMNS) + 1):
                ws.cell(row=i, column=col).fill = SUBHEADING_FILL
                ws.cell(row=i, column=col).font = Font(bold=True)
    ws.freeze_panes = "B4"


def _write_correlation_sheet(ws, corr_df, as_of):
    ws["A1"] = ("Correlation of weekly returns: +1 moves in step, 0 unrelated, -1 "
                "opposite. Funds with under 1 year of data are left out. " + as_of)
    ws["A1"].font = Font(italic=True)
    codes = list(corr_df.columns)
    for j, code in enumerate(codes, start=2):
        cell = ws.cell(row=3, column=j, value=code)
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(textRotation=90, horizontal="center")
        ws.column_dimensions[get_column_letter(j)].width = 5.5
    ws.column_dimensions["A"].width = 10
    ws.row_dimensions[3].height = 52
    for i, row_code in enumerate(codes, start=4):
        ws.cell(row=i, column=1, value=row_code).font = Font(bold=True, size=9)
        for j, col_code in enumerate(codes, start=2):
            v = corr_df.loc[row_code, col_code]
            cell = ws.cell(row=i, column=j, value=None if pd.isna(v) else round(float(v), 2))
            cell.number_format = "0.00"
            cell.font = Font(size=8)
            cell.alignment = Alignment(horizontal="center")
    last = f"{get_column_letter(len(codes) + 1)}{len(codes) + 3}"
    ws.conditional_formatting.add(f"B4:{last}", ColorScaleRule(
        start_type="num", start_value=-1, start_color="2E75B6",
        mid_type="num", mid_value=0, mid_color="FFFFFF",
        end_type="num", end_value=1, end_color="E8562A"))
    ws.freeze_panes = "B4"


def save_excel(metrics, filename, as_of, sector_df=None, corr_df=None,
               chart_files=("growth_chart.png", "risk_return_chart.png")):
    """
    Write the results into one Excel workbook:
      By Sector      - every fund grouped under its Fund Map heading and sub-heading
      All ETFs       - one flat table, best annual return first, with filters
      Sector Summary - averages for each heading and sub-heading
      Correlations   - how closely each pair of funds moves together
      Charts         - the two charts
    """
    wb = Workbook()

    # --- Sheet 1: grouped like the Fund Map ---------------------------------
    ws = wb.active
    ws.title = "By Sector"
    _setup_sheet(ws, SECTOR_COLUMNS, as_of)
    row = 4
    headings = list(dict.fromkeys(h for h, _, _ in FUND_MAP))   # keeps map order
    for heading in headings:
        in_heading = metrics[metrics["Category"] == heading]
        if in_heading.empty:
            continue
        _band(ws, row, len(SECTOR_COLUMNS), heading.upper(), HEADING_FILL,
              Font(bold=True, color="FFFFFF"))
        row += 1
        for sub in dict.fromkeys(s for h, s, _ in FUND_MAP if h == heading):
            in_sub = in_heading[in_heading["Sub-category"] == sub]
            if in_sub.empty:
                continue
            if sub:
                _band(ws, row, len(SECTOR_COLUMNS), sub, SUBHEADING_FILL,
                      Font(bold=True))
                row += 1
            for code, r in in_sub.iterrows():
                values = [code if c == "Code" else r[c] for c in SECTOR_COLUMNS]
                _write_row(ws, row, SECTOR_COLUMNS, values)
                row += 1

    # --- Sheet 2: flat and sortable -----------------------------------------
    flat = wb.create_sheet("All ETFs")
    _setup_sheet(flat, FLAT_COLUMNS, as_of)
    ranked = metrics.sort_values("Annual return %", ascending=False, na_position="last")
    for i, (code, r) in enumerate(ranked.iterrows(), start=4):
        values = [code if c == "Code" else r[c] for c in FLAT_COLUMNS]
        _write_row(flat, i, FLAT_COLUMNS, values)
    flat.auto_filter.ref = f"A3:{flat.cell(row=3, column=len(FLAT_COLUMNS)).column_letter}{3 + len(ranked)}"

    # --- Sheets 3 and 4: sector summary and correlations --------------------
    if sector_df is not None and not sector_df.empty:
        _write_sector_sheet(wb.create_sheet("Sector Summary"), sector_df, as_of)
    if corr_df is not None and not corr_df.empty:
        _write_correlation_sheet(wb.create_sheet("Correlations"), corr_df, as_of)

    # --- Sheet 5: charts ----------------------------------------------------
    charts = wb.create_sheet("Charts")
    growth_file, risk_file = chart_files
    if growth_file and os.path.exists(growth_file):
        charts.add_image(XLImage(growth_file), "A1")
    if risk_file and os.path.exists(risk_file):
        charts.add_image(XLImage(risk_file), "A40")

    wb.save(filename)


# ---------------------------------------------------------------------------
# 6. THE PORTFOLIO BUILDER PAGE (HTML)
# ---------------------------------------------------------------------------
def _num(value, digits=4):
    """A plain number for JSON, or None if it is missing."""
    return None if pd.isna(value) else round(float(value), digits)


def build_html(prices, metrics, sector_df, corr_df, as_of, filename, source=""):
    """
    Write one self-contained HTML page (no internet needed to use it). The price
    history is embedded in the page, so it can run portfolio simulations on its
    own. Re-run the script to refresh the prices.
    """
    filled = prices.ffill()   # carry the last price over holidays and gaps
    dates = [d.strftime("%Y-%m-%d") for d in prices.index]

    funds = []
    for code, r in metrics.iterrows():
        has_data = pd.notna(r["Latest price"])
        first = None
        series = None
        if has_data:
            first = int(prices.index.get_loc(prices[code].first_valid_index()))
            series = [_num(v) for v in filled[code]]
        funds.append({
            "code": code, "name": r["Name"], "category": r["Category"],
            "sub": r["Sub-category"], "exchange": r["Exchange"],
            "status": r["Status"], "hasData": bool(has_data), "first": first,
            "annual": _num(r["Annual return %"], 1), "prices": series,
        })

    sectors = []
    for rec in sector_df.to_dict("records"):
        sectors.append({
            "level": rec["Level"], "sector": rec["Sector"], "funds": int(rec["Funds"]),
            "withData": int(rec["Funds with 1y+ data"]),
            "avgAnnual": _num(rec["Avg annual return %"], 1),
            "medAnnual": _num(rec["Median annual return %"], 1),
            "avgVol": _num(rec["Avg volatility %"], 1),
            "avgDD": _num(rec["Avg max drawdown %"], 1),
            "best": rec["Best fund"], "bestRet": _num(rec["Best annual return %"], 1),
            "worst": rec["Worst fund"], "worstRet": _num(rec["Worst annual return %"], 1),
        })

    corr = {"codes": list(corr_df.columns),
            "matrix": [[_num(v, 2) for v in row] for row in corr_df.to_numpy()]}

    payload = json.dumps({"asOf": as_of, "source": source, "dates": dates,
                          "funds": funds, "sectors": sectors, "corr": corr},
                         separators=(",", ":"), allow_nan=False)
    payload = payload.replace("</", "<\\/")   # keeps the data safe inside <script>
    with open(filename, "w", encoding="utf-8") as f:
        f.write(HTML_TEMPLATE.replace("__DATA__", payload))


HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ETF Portfolio Builder</title>
<style>
:root{--teal:#0f3d4a;--orange:#e8562a;--band:#e3e7e9;--text:#1c2b33;--muted:#5d6b73;--line:#d5dade;--bg:#f4f6f7;--card:#fff;--pos:#1a7f4b;--neg:#c0392b}
*{box-sizing:border-box}
body{margin:0;font:14px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:var(--text);background:var(--bg)}
header{background:var(--teal);color:#fff;padding:18px 24px 0}
header h1{margin:0 0 4px;font-size:20px;font-weight:600}
header p{margin:0 0 12px;font-size:12px;opacity:.8}
nav{display:flex;gap:4px;flex-wrap:wrap}
nav button{border:0;background:transparent;color:#cfe0e5;padding:10px 16px;font:inherit;cursor:pointer;border-bottom:3px solid transparent}
nav button.active{color:#fff;border-bottom-color:var(--orange);font-weight:600}
main{max-width:1240px;margin:0 auto;padding:20px 24px 48px}
section.tab{display:none}
section.tab.active{display:block}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:16px;margin-bottom:16px}
.card h2{margin:0 0 10px;font-size:15px}
.card h3{margin:18px 0 8px;font-size:13px}
.grid{display:grid;gap:16px;grid-template-columns:minmax(300px,400px) 1fr;align-items:start}
@media(max-width:860px){.grid{grid-template-columns:1fr}}
.muted{color:var(--muted)}
.fine{font-size:12px;color:var(--muted);margin:12px 0 0}
input,select,button{font:inherit}
input[type=text],input[type=number],input[type=date],select{border:1px solid var(--line);border-radius:6px;padding:6px 8px;background:#fff;color:var(--text)}
input[type=number]{width:110px}
#search{width:100%;margin-bottom:8px}
button.btn{border:1px solid var(--line);background:#fff;border-radius:6px;padding:5px 10px;cursor:pointer}
button.btn:hover{border-color:var(--teal)}
button.x{border:0;background:transparent;color:var(--muted);cursor:pointer;font-size:16px;line-height:1}
.picker{max-height:540px;overflow:auto;border:1px solid var(--line);border-radius:6px}
.picker details{border-bottom:1px solid var(--line)}
.picker summary{cursor:pointer;background:var(--teal);color:#fff;font-weight:600;padding:7px 10px;list-style:none;display:flex;justify-content:space-between;align-items:center}
.picker summary::-webkit-details-marker{display:none}
.picker summary .link{font-size:11px;font-weight:400;text-decoration:underline;opacity:.9}
.picker .sub{font-size:11px;text-transform:uppercase;letter-spacing:.03em;color:var(--muted);background:var(--band);padding:4px 10px}
.fund{display:flex;gap:8px;align-items:center;padding:4px 10px;cursor:pointer}
.fund:hover{background:#f0f4f5}
.fund.disabled{opacity:.45;cursor:not-allowed}
.fund .code{font-weight:600;width:58px}
.fund .nm{flex:1;color:var(--muted);font-size:12px}
.fund .hint{font-size:11px;white-space:nowrap}
table{border-collapse:collapse;width:100%}
th,td{padding:6px 8px;border-bottom:1px solid var(--line);text-align:right;vertical-align:middle}
th:first-child,td:first-child{text-align:left}
th{font-size:12px;color:var(--muted);font-weight:600}
.form{display:grid;grid-template-columns:auto 1fr;gap:8px 12px;align-items:center}
.pos{color:var(--pos)}.neg{color:var(--neg)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:14px}
.kpi{border:1px solid var(--line);border-radius:6px;padding:10px 12px;background:#fff}
.kpi .l{font-size:11px;color:var(--muted)}
.kpi .v{font-size:20px;font-weight:600}
.kpi .s{font-size:11px;color:var(--muted)}
.note{background:#fff6e5;border:1px solid #f0d9a8;border-radius:6px;padding:8px 10px;margin-bottom:12px;font-size:13px}
#chartWrap{position:relative}
#chart{width:100%;display:block}
.tip{position:absolute;pointer-events:none;background:#fff;border:1px solid var(--line);box-shadow:0 2px 8px rgba(0,0,0,.12);border-radius:6px;padding:6px 8px;font-size:12px;display:none;white-space:nowrap;z-index:5}
.legend{display:flex;gap:16px;font-size:12px;color:var(--muted);margin:4px 0 0 64px}
.legend i{display:inline-block;width:14px;height:3px;vertical-align:middle;margin-right:5px}
table.sec th,table.sec td{font-size:13px}
table.sec tr.heading td{background:var(--teal);color:#fff;font-weight:600}
table.sec tr.heading td .pos,table.sec tr.heading td .neg{color:#fff}
table.sec tr.sub td:first-child{padding-left:24px}
table.sec tr.total td{background:var(--band);font-weight:600}
.corrWrap{overflow:auto;max-height:78vh;border:1px solid var(--line);border-radius:6px;background:#fff}
table.corr{border-collapse:separate;border-spacing:0;font-size:10px;width:auto}
table.corr th{position:sticky;background:#fff;z-index:2;padding:2px;border:0;color:var(--text)}
table.corr thead th{top:0}
table.corr th.row{left:0;text-align:left;padding:0 6px;font-size:11px;font-weight:600}
table.corr th.corner{left:0;z-index:3}
table.corr th.col span{writing-mode:vertical-rl;transform:rotate(180deg);display:inline-block;font-size:11px;font-weight:600;padding:2px 0}
table.corr td{width:36px;min-width:36px;height:24px;text-align:center;padding:0;border:1px solid #fff}
.scale{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--muted);margin:8px 0}
.scale div{width:200px;height:10px;border-radius:5px;background:linear-gradient(90deg,#2e75b6,#fff,#e8562a);border:1px solid var(--line)}
</style>
</head>
<body>
<header>
  <h1>ETF Portfolio Builder</h1>
  <p id="asof"></p>
  <nav>
    <button class="active" data-tab="portfolio">Portfolio builder</button>
    <button data-tab="sectors">Sector summary</button>
    <button data-tab="corr">Correlations</button>
  </nav>
</header>
<main>

<section class="tab active" id="tab-portfolio">
  <div class="grid">
    <div class="card">
      <h2>1. Pick funds</h2>
      <input type="text" id="search" placeholder="Search by code or name">
      <div class="picker" id="picker"></div>
    </div>
    <div>
      <div class="card"><h2>2. Weights</h2><div id="weights"></div></div>
      <div class="card">
        <h2>3. Money</h2>
        <div class="form">
          <label for="init">Starting amount ($)</label><input type="number" id="init" min="0" step="100" value="10000">
          <label for="contrib">Add money ($)</label>
          <div><input type="number" id="contrib" min="0" step="50" value="0">
            <select id="freq"><option value="week">every week</option><option value="fortnight">every fortnight</option><option value="month" selected>every month</option></select></div>
          <label for="start">Start date</label><input type="date" id="start">
        </div>
        <p class="fine">If a chosen fund only has prices from a later date, the portfolio starts on that date instead.</p>
      </div>
    </div>
  </div>

  <div class="card" id="results">
    <h2>4. How it would have grown</h2>
    <div id="msg"></div>
    <div class="kpis" id="kpis"></div>
    <div id="chartWrap"><svg id="chart"></svg><div class="tip" id="tip"></div></div>
    <div class="legend" id="legend"></div>
    <h3>Holdings</h3>
    <div id="holdings"></div>
    <p class="fine">Assumes no brokerage, fees or tax, and that each deposit buys at that day's closing price (part-units allowed). Prices include reinvested distributions. Deposits are split by your weights, with no rebalancing afterwards. This shows what would have happened in the past; it does not predict the future and is not financial advice.</p>
  </div>
</section>

<section class="tab" id="tab-sectors">
  <div class="card">
    <h2>Sector summary</h2>
    <p class="fine" style="margin-top:0">Simple averages across the funds in each group that have at least 1 year of history. This is not the return of a sector index, and funds that launched recently are left out of the averages.</p>
    <div id="sectors"></div>
  </div>
</section>

<section class="tab" id="tab-corr">
  <div class="card">
    <h2>Correlations</h2>
    <p class="fine" style="margin-top:0">How closely two funds move together, using weekly returns. +1 means they move in step, 0 means unrelated, and negative means they tend to move in opposite directions. Funds with under a year of history are left out. Hover over a cell for the pair.</p>
    <p><label>Show <select id="corrMode"><option value="all">all funds</option><option value="mine">only funds in my portfolio</option></select></label></p>
    <div class="scale"><span>-1</span><div></div><span>+1</span></div>
    <div class="corrWrap" id="corr"></div>
  </div>
</section>

</main>

<script id="data" type="application/json">__DATA__</script>
<script>
//--CORE-START--
// Pure calculation code (no page elements), so it can be tested on its own.
function lowerBound(arr, x){let lo=0,hi=arr.length;while(lo<hi){const m=(lo+hi)>>1;if(arr[m]<x)lo=m+1;else hi=m;}return lo;}
function addDaysISO(iso,n){const d=new Date(iso+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10);}
function addMonthsISO(iso,n){
  const p=iso.split('-').map(Number);
  const t=new Date(Date.UTC(p[0],p[1]-1+n,1));
  const last=new Date(Date.UTC(t.getUTCFullYear(),t.getUTCMonth()+1,0)).getUTCDate();
  t.setUTCDate(Math.min(p[2],last));
  return t.toISOString().slice(0,10);
}
// Indexes of the trading days on which money goes in. The first is the starting amount.
function schedule(dates,startIdx,endIdx,freq){
  const idxs=[startIdx];
  if(freq==='none') return idxs;
  const startISO=dates[startIdx];
  for(let k=1;;k++){
    const target=freq==='week'?addDaysISO(startISO,7*k):freq==='fortnight'?addDaysISO(startISO,14*k):addMonthsISO(startISO,k);
    const i=lowerBound(dates,target);
    if(i>endIdx) break;
    idxs.push(i);
  }
  return idxs;
}
// Money-weighted annual return (like Excel's XIRR). flows: [{t: days, cf: cash flow}]
function xirr(flows){
  const npv=r=>flows.reduce((s,f)=>s+f.cf/Math.pow(1+r,f.t/365),0);
  let lo=-0.99,hi=10,flo=npv(lo),fhi=npv(hi);
  if(!(flo*fhi<0)) return null;
  for(let i=0;i<200;i++){
    const mid=(lo+hi)/2,fm=npv(mid);
    if(flo*fm<=0){hi=mid;fhi=fm;}else{lo=mid;flo=fm;}
  }
  return (lo+hi)/2;
}
// picks: [{code, weight}] with weights adding to 1.  opts: {initial, recurring, freq, startISO}
function simulate(D,byCode,picks,opts){
  const dates=D.dates,endIdx=dates.length-1;
  let firstIdx=0,limiter=null;
  picks.forEach(p=>{const f=byCode[p.code];if(f.first>firstIdx){firstIdx=f.first;limiter=p.code;}});
  let userIdx=opts.startISO?lowerBound(dates,opts.startISO):0;
  if(userIdx>endIdx) userIdx=endIdx;
  const startIdx=Math.max(firstIdx,userIdx);
  const limited=firstIdx>userIdx;
  const sched=schedule(dates,startIdx,endIdx,opts.freq);
  const flowAt=new Map();
  sched.forEach((idx,k)=>{const amt=k===0?opts.initial:opts.recurring;if(amt>0)flowAt.set(idx,(flowAt.get(idx)||0)+amt);});
  const units=picks.map(()=>0),invested=picks.map(()=>0);
  const len=endIdx-startIdx+1;
  const value=new Array(len),put=new Array(len);
  const twIdx=[];let tw=1,cum=0,prevV=0;const rets=[];
  for(let i=startIdx;i<=endIdx;i++){
    const flow=flowAt.get(i)||0;
    if(flow>0){
      picks.forEach((p,j)=>{const px=byCode[p.code].prices[i];units[j]+=flow*p.weight/px;invested[j]+=flow*p.weight;});
      cum+=flow;
    }
    let v=0;picks.forEach((p,j)=>{v+=units[j]*byCode[p.code].prices[i];});
    if(prevV>0){const r=(v-flow)/prevV-1;tw*=1+r;rets.push(r);}
    value[i-startIdx]=v;put[i-startIdx]=cum;twIdx.push(tw);prevV=v;
  }
  const day=i=>Date.parse(dates[i]+'T00:00:00Z')/864e5;
  const days=day(endIdx)-day(startIdx);
  const final=value[len-1];
  const flows=[];flowAt.forEach((amt,idx)=>flows.push({t:day(idx)-day(startIdx),cf:-amt}));
  flows.push({t:days,cf:final});
  const irr=(days>=365&&cum>0)?xirr(flows):null;
  const twrAnnual=(days>=365&&tw>0)?Math.pow(tw,365/days)-1:null;
  let vol=null;
  if(rets.length>1){const m=rets.reduce((a,b)=>a+b,0)/rets.length;vol=Math.sqrt(rets.reduce((a,b)=>a+(b-m)*(b-m),0)/(rets.length-1))*Math.sqrt(252);}
  let peak=-Infinity,maxDD=0;twIdx.forEach(x=>{if(x>peak)peak=x;const dd=x/peak-1;if(dd<maxDD)maxDD=dd;});
  const holdings=picks.map((p,j)=>{const v=units[j]*byCode[p.code].prices[endIdx];return{code:p.code,weight:p.weight,invested:invested[j],value:v};});
  return{startIdx,endIdx,limited,limiter,days,final,cum,profit:final-cum,irr,twrAnnual,vol,maxDD,holdings,value,put};
}
//--CORE-END--

const D=JSON.parse(document.getElementById('data').textContent);
const byCode={};D.funds.forEach(f=>byCode[f.code]=f);
const sel=new Map();   // code -> weight
const $=id=>document.getElementById(id);
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const money=v=>new Intl.NumberFormat('en-AU',{style:'currency',currency:'AUD',maximumFractionDigits:0}).format(v);
const moneyShort=v=>v>=1e6?'$'+(v/1e6).toFixed(v>=1e7?0:1)+'m':v>=1e3?'$'+(v/1e3).toFixed(v>=1e4?0:1)+'k':'$'+Math.round(v);
const pct=(v,d=1)=>(v==null||isNaN(v))?'—':(v>0?'+':'')+v.toFixed(d)+'%';
const cls=v=>v>0?'pos':v<0?'neg':'';
const colored=(v,d=1)=>`<span class="${cls(v)}">${pct(v,d)}</span>`;
const fmtDate=iso=>new Date(iso+'T00:00:00Z').toLocaleDateString('en-AU',{day:'numeric',month:'short',year:'numeric',timeZone:'UTC'});
const monthYear=iso=>new Date(iso+'T00:00:00Z').toLocaleDateString('en-AU',{month:'short',year:'numeric',timeZone:'UTC'});

$('asof').textContent=(D.source?'Fund list: '+D.source+'. ':'')+D.asOf;

// ---- tabs ----
document.querySelectorAll('nav button').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('nav button').forEach(x=>x.classList.toggle('active',x===b));
  document.querySelectorAll('section.tab').forEach(s=>s.classList.toggle('active',s.id==='tab-'+b.dataset.tab));
  if(b.dataset.tab==='corr') renderCorr();
  if(b.dataset.tab==='portfolio') drawChart();
});

// ---- fund picker ----
function buildPicker(){
  const picker=$('picker');picker.innerHTML='';
  let details=null,lastCat=null,lastSub=null;
  D.funds.forEach(f=>{
    if(f.category!==lastCat){
      details=document.createElement('details');details.open=true;details.dataset.cat=f.category;
      const s=document.createElement('summary');
      s.innerHTML=`<span>${esc(f.category)}</span><span class="link" data-cat="${esc(f.category)}">add / remove all</span>`;
      details.appendChild(s);picker.appendChild(details);lastCat=f.category;lastSub=null;
    }
    if(f.sub&&f.sub!==lastSub){const d=document.createElement('div');d.className='sub';d.textContent=f.sub;details.appendChild(d);}
    lastSub=f.sub;
    const row=document.createElement('label');
    row.className='fund'+(f.hasData?'':' disabled');
    row.dataset.search=(f.code+' '+f.name+' '+f.category+' '+f.sub).toLowerCase();
    row.title=f.hasData?f.name+' ('+f.exchange+')':f.name+': '+f.status;
    const hint=f.hasData?(f.annual==null?'<span class="muted">new</span>':colored(f.annual)+' <span class="muted">p.a.</span>'):'<span class="muted">no data</span>';
    row.innerHTML=`<input type="checkbox" data-code="${f.code}" ${f.hasData?'':'disabled'}><span class="code">${f.code}</span><span class="nm">${esc(f.name)}</span><span class="hint">${hint}</span>`;
    details.appendChild(row);
  });
  picker.querySelectorAll('input[type=checkbox]').forEach(cb=>cb.onchange=()=>{
    if(cb.checked) addFund(cb.dataset.code); else sel.delete(cb.dataset.code);
    renderWeights();update();
  });
  picker.querySelectorAll('.link').forEach(l=>l.onclick=e=>{
    e.preventDefault();e.stopPropagation();
    const codes=D.funds.filter(f=>f.category===l.dataset.cat&&f.hasData).map(f=>f.code);
    const allOn=codes.every(c=>sel.has(c));
    codes.forEach(c=>{if(allOn)sel.delete(c);else if(!sel.has(c))addFund(c);});
    syncChecks();renderWeights();update();
  });
}
function addFund(code){
  if(sel.has(code)) return;
  const vals=[...sel.values()];
  sel.set(code,vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:100);
}
function syncChecks(){document.querySelectorAll('#picker input[type=checkbox]').forEach(cb=>cb.checked=sel.has(cb.dataset.code));}
$('search').oninput=e=>{
  const q=e.target.value.trim().toLowerCase();
  document.querySelectorAll('#picker .fund').forEach(r=>r.style.display=(!q||r.dataset.search.includes(q))?'':'none');
  document.querySelectorAll('#picker .sub').forEach(s=>s.style.display=q?'none':'');
  document.querySelectorAll('#picker details').forEach(d=>{
    const any=[...d.querySelectorAll('.fund')].some(r=>r.style.display!=='none');
    d.style.display=any?'':'none';if(q&&any)d.open=true;
  });
};

// ---- weights ----
function renderWeights(){
  const box=$('weights');
  if(!sel.size){box.innerHTML='<p class="muted" style="margin:0">Tick at least one fund on the left.</p>';return;}
  const total=[...sel.values()].reduce((a,b)=>a+b,0)||1;
  let h='<table><thead><tr><th>Fund</th><th>Weight</th><th>Share</th><th></th></tr></thead><tbody>';
  sel.forEach((w,code)=>{
    h+=`<tr><td><b>${code}</b> <span class="muted">${esc(byCode[code].name)}</span></td><td><input type="number" class="w" data-code="${code}" min="0" step="1" value="${+w.toFixed(2)}"></td><td class="muted pct" data-code="${code}">${(w/total*100).toFixed(1)}%</td><td><button class="x" data-code="${code}" title="Remove">×</button></td></tr>`;
  });
  h+='</tbody></table><p style="margin:10px 0 0"><button class="btn" id="equal">Equal weights</button> <button class="btn" id="clear">Clear all</button> <span class="fine">Weights are scaled to add up to 100%.</span></p>';
  box.innerHTML=h;
  box.querySelectorAll('input.w').forEach(inp=>inp.oninput=()=>{
    sel.set(inp.dataset.code,Math.max(0,+inp.value||0));
    const t=[...sel.values()].reduce((a,b)=>a+b,0)||1;
    box.querySelectorAll('.pct').forEach(c=>c.textContent=(sel.get(c.dataset.code)/t*100).toFixed(1)+'%');
    update();
  });
  box.querySelectorAll('button.x').forEach(b=>b.onclick=()=>{sel.delete(b.dataset.code);syncChecks();renderWeights();update();});
  $('equal').onclick=()=>{const w=+(100/sel.size).toFixed(2);sel.forEach((_,c)=>sel.set(c,w));renderWeights();update();};
  $('clear').onclick=()=>{sel.clear();syncChecks();renderWeights();update();};
}

// ---- results ----
let lastChart=null;
function update(){
  const picks=[...sel.entries()].map(([code,weight])=>({code,weight}));
  const total=picks.reduce((a,p)=>a+p.weight,0);
  const msg=$('msg');msg.innerHTML='';msg.className='';
  const blank=t=>{lastChart=null;if(t){msg.className='note';msg.textContent=t;}$('kpis').innerHTML='';$('chart').innerHTML='';$('legend').innerHTML='';$('holdings').innerHTML='';$('tip').style.display='none';};
  if(!picks.length) return blank('Pick some funds to see how they would have grown.');
  if(total<=0) return blank('Give at least one fund a weight above 0.');
  picks.forEach(p=>p.weight/=total);
  const initial=Math.max(0,+$('init').value||0),recurring=Math.max(0,+$('contrib').value||0);
  if(initial+recurring<=0) return blank('Enter a starting amount or a regular amount to add.');
  const freq=recurring>0?$('freq').value:'none';
  const r=simulate(D,byCode,picks,{initial,recurring,freq,startISO:$('start').value||D.dates[0]});
  const startISO=D.dates[r.startIdx],endISO=D.dates[r.endIdx];
  let note='';
  if(r.limited) note+=`The portfolio starts on ${fmtDate(startISO)} because ${r.limiter} only has prices from then. `;
  if(r.days<365) note+='Under a year of history, so annual figures are not shown. ';
  if(note){msg.className='note';msg.textContent=note;}
  const years=r.days/365.25;
  const kp=[
    ['Final value',money(r.final),`${fmtDate(startISO)} to ${fmtDate(endISO)} (${years.toFixed(1)} years)`],
    ['Money put in',money(r.cum),initial>0?`${money(initial)} to start${recurring>0?' + '+money(recurring)+' '+({week:'weekly',fortnight:'fortnightly',month:'monthly'})[freq]:''}`:'regular deposits'],
    ['Profit',`<span class="${cls(r.profit)}">${r.profit<0?'-':''}${money(Math.abs(r.profit))}</span>`,r.cum>0?pct(r.profit/r.cum*100)+' on money put in':''],
    ['Your annual return',r.irr==null?'—':colored(r.irr*100),'money-weighted: accounts for when you added money'],
    ['Portfolio annual return',r.twrAnnual==null?'—':colored(r.twrAnnual*100),'time-weighted: ignores the timing of deposits'],
    ['Biggest fall',pct(r.maxDD*100),'largest drop from a previous peak'],
    ['Volatility',r.vol==null?'—':(r.vol*100).toFixed(1)+'%','how much it swings, per year']
  ];
  $('kpis').innerHTML=kp.map(k=>`<div class="kpi"><div class="l">${k[0]}</div><div class="v">${k[1]}</div><div class="s">${k[2]}</div></div>`).join('');
  $('legend').innerHTML='<span><i style="background:#0f3d4a"></i>Portfolio value</span><span><i style="background:#e8562a"></i>Money put in</span>';
  let h='<table><thead><tr><th>Fund</th><th>Weight</th><th>Put in</th><th>Value now</th><th>Profit</th><th>Return</th></tr></thead><tbody>';
  r.holdings.forEach(x=>{const pr=x.value-x.invested;h+=`<tr><td><b>${x.code}</b> <span class="muted">${esc(byCode[x.code].name)}</span></td><td>${(x.weight*100).toFixed(1)}%</td><td>${money(x.invested)}</td><td>${money(x.value)}</td><td class="${cls(pr)}">${pr<0?'-':''}${money(Math.abs(pr))}</td><td>${x.invested>0?colored(pr/x.invested*100):'—'}</td></tr>`;});
  $('holdings').innerHTML=h+'</tbody></table>';
  lastChart={startIdx:r.startIdx,value:r.value,put:r.put};
  drawChart();
}
['init','contrib','freq','start'].forEach(id=>$(id).oninput=update);

// ---- chart ----
function niceTicks(max,count){
  const raw=max/count,mag=Math.pow(10,Math.floor(Math.log10(raw||1))),n=raw/mag;
  const step=(n<=1?1:n<=2?2:n<=5?5:10)*mag;const out=[];
  for(let v=0;v<=max+step*0.5;v+=step)out.push(v);
  return out;
}
function drawChart(){
  if(!lastChart) return;
  const svg=$('chart'),wrap=$('chartWrap');
  const W=wrap.clientWidth||800,H=340,pl=64,pr=16,pt=14,pb=28,pw=W-pl-pr,ph=H-pt-pb;
  const {startIdx,value,put}=lastChart,n=value.length;
  const ticks=niceTicks(Math.max(...value,...put)*1.02||1,5),top=ticks[ticks.length-1];
  const X=i=>pl+(n>1?i/(n-1)*pw:pw/2),Y=v=>pt+ph-v/top*ph;
  const line=a=>a.map((v,i)=>(i?'L':'M')+X(i).toFixed(1)+' '+Y(v).toFixed(1)).join('');
  let s='';
  ticks.forEach(t=>{s+=`<line x1="${pl}" x2="${W-pr}" y1="${Y(t)}" y2="${Y(t)}" stroke="#e3e7e9"/><text x="${pl-8}" y="${Y(t)+4}" text-anchor="end" font-size="11" fill="#5d6b73">${moneyShort(t)}</text>`;});
  const nx=Math.min(6,n);
  for(let k=0;k<nx;k++){const i=nx>1?Math.round(k*(n-1)/(nx-1)):0;s+=`<text x="${X(i)}" y="${H-8}" text-anchor="${k===0?'start':k===nx-1?'end':'middle'}" font-size="11" fill="#5d6b73">${monthYear(D.dates[startIdx+i])}</text>`;}
  s+=`<path d="${line(put)}" fill="none" stroke="#e8562a" stroke-width="2"/><path d="${line(value)}" fill="none" stroke="#0f3d4a" stroke-width="2.5"/>`;
  s+=`<line id="hl" y1="${pt}" y2="${pt+ph}" stroke="#9aa7ad" style="display:none"/><circle id="hd1" r="4" fill="#0f3d4a" style="display:none"/><circle id="hd2" r="4" fill="#e8562a" style="display:none"/>`;
  s+=`<rect id="ov" x="${pl}" y="${pt}" width="${pw}" height="${ph}" fill="transparent"/>`;
  svg.setAttribute('viewBox',`0 0 ${W} ${H}`);svg.setAttribute('height',H);svg.innerHTML=s;
  const tip=$('tip'),ov=$('ov');
  ov.onmousemove=e=>{
    const r=svg.getBoundingClientRect(),x=(e.clientX-r.left)*(W/r.width)-pl;
    const i=Math.max(0,Math.min(n-1,n>1?Math.round(x/pw*(n-1)):0));
    $('hl').setAttribute('x1',X(i));$('hl').setAttribute('x2',X(i));$('hl').style.display='';
    $('hd1').setAttribute('cx',X(i));$('hd1').setAttribute('cy',Y(value[i]));$('hd1').style.display='';
    $('hd2').setAttribute('cx',X(i));$('hd2').setAttribute('cy',Y(put[i]));$('hd2').style.display='';
    tip.innerHTML=`<b>${fmtDate(D.dates[startIdx+i])}</b><br>Value: ${money(value[i])}<br>Put in: ${money(put[i])}<br>Profit: ${money(value[i]-put[i])}`;
    tip.style.display='block';
    const wr=wrap.getBoundingClientRect();let left=e.clientX-wr.left+14;
    if(left+tip.offsetWidth>wr.width) left=e.clientX-wr.left-tip.offsetWidth-14;
    tip.style.left=left+'px';tip.style.top=Math.max(0,e.clientY-wr.top-30)+'px';
  };
  ov.onmouseleave=()=>{tip.style.display='none';['hl','hd1','hd2'].forEach(id=>$(id).style.display='none');};
}
window.addEventListener('resize',drawChart);

// ---- sector summary ----
function renderSectors(){
  let h='<table class="sec"><thead><tr><th>Sector</th><th>Funds</th><th>With 1y+ data</th><th>Avg annual return</th><th>Median</th><th>Avg volatility</th><th>Avg biggest fall</th><th>Best fund</th><th>Worst fund</th></tr></thead><tbody>';
  D.sectors.forEach(r=>{
    const name=r.level==='heading'?r.sector.toUpperCase():esc(r.sector);
    h+=`<tr class="${r.level}"><td>${name}</td><td>${r.funds}</td><td>${r.withData}</td><td>${colored(r.avgAnnual)}</td><td>${colored(r.medAnnual)}</td><td>${r.avgVol==null?'—':r.avgVol.toFixed(1)+'%'}</td><td>${pct(r.avgDD)}</td><td>${r.best?`${r.best} (${pct(r.bestRet)})`:'—'}</td><td>${r.worst?`${r.worst} (${pct(r.worstRet)})`:'—'}</td></tr>`;
  });
  $('sectors').innerHTML=h+'</tbody></table>';
}

// ---- correlations ----
function corrColor(v){return v>=0?`rgba(232,86,42,${v.toFixed(2)})`:`rgba(46,117,182,${(-v).toFixed(2)})`;}
function renderCorr(){
  const codes=D.corr.codes,m=D.corr.matrix;
  let idx=codes.map((_,i)=>i);
  if($('corrMode').value==='mine') idx=idx.filter(i=>sel.has(codes[i]));
  const box=$('corr');
  if(idx.length<2){box.innerHTML='<p class="muted" style="padding:12px">Pick at least two funds (with 1+ year of data) on the Portfolio builder tab to see their correlations.</p>';return;}
  let h='<table class="corr"><thead><tr><th class="corner"></th>'+idx.map(i=>`<th class="col" title="${esc(byCode[codes[i]].name)}"><span>${codes[i]}</span></th>`).join('')+'</tr></thead><tbody>';
  idx.forEach(i=>{
    h+=`<tr><th class="row" title="${esc(byCode[codes[i]].name)}">${codes[i]}</th>`;
    idx.forEach(j=>{const v=m[i][j];h+=v==null?'<td></td>':`<td style="background:${corrColor(v)};${Math.abs(v)>0.7?'color:#fff;':''}" title="${codes[i]} vs ${codes[j]}: ${v.toFixed(2)}">${v.toFixed(2)}</td>`;});
    h+='</tr>';
  });
  box.innerHTML=h+'</tbody></table>';
}
$('corrMode').onchange=renderCorr;

// ---- start ----
$('start').min=D.dates[0];$('start').max=D.dates[D.dates.length-1];$('start').value=D.dates[0];
buildPicker();renderWeights();renderSectors();update();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# 7. RUNNING IT
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Analyse the funds listed in a spreadsheet and build a portfolio page")
    parser.add_argument("--funds", default="funds.xlsx",
                        help="spreadsheet (.xlsx or .csv) listing the funds to analyse "
                             "(default: funds.xlsx)")
    parser.add_argument("--out", default=None,
                        help="start of the output file names, so different lists don't "
                             "overwrite each other, e.g. --out mylist gives "
                             "mylist_results.xlsx (default: etf)")
    parser.add_argument("--period", default="5y",
                        help="history to download, e.g. 1y, 3y, 5y, max (default: 5y)")
    parser.add_argument("--tickers", nargs="+",
                        help="only analyse these tickers from your list, e.g. GXAI SEMI")
    parser.add_argument("--common-start", action="store_true",
                        help="start the growth chart on the first date all funds have data")
    parser.add_argument("--chart-top", type=int, default=10,
                        help="how many funds to show on the growth chart, best annual "
                             "return first (default: 10, use 0 for all)")
    args = parser.parse_args()

    # Load the fund list. On the very first run, create the starter spreadsheet.
    if not os.path.exists(args.funds):
        if args.funds != "funds.xlsx":
            sys.exit(f"Couldn't find {args.funds}.")
        create_starter_file(args.funds)
        print("Created funds.xlsx with the 53 Global X funds to get you started. "
              "Edit it any time to analyse different funds.")
    load_funds(args.funds)
    print(f"Using {args.funds}: {len(FUNDS)} funds")

    if args.out:
        results_file = f"{args.out}_results.xlsx"
        html_file = f"{args.out}_portfolio_builder.html"
        growth_file = f"{args.out}_growth_chart.png"
        risk_file = f"{args.out}_risk_return_chart.png"
    else:
        results_file, html_file = "etf_results.xlsx", "etf_portfolio_builder.html"
        growth_file, risk_file = "growth_chart.png", "risk_return_chart.png"

    if args.tickers:
        wanted = {t.upper() for t in args.tickers}
        unknown = sorted(wanted - set(FUNDS))
        if unknown:
            print(f"Not in {args.funds} (add them to the spreadsheet): {unknown}")
        codes = [c for c in FUNDS if c in wanted]   # keep the spreadsheet order
    else:
        codes = list(FUNDS)
    if not codes:
        sys.exit("No valid tickers to analyse.")

    print(f"Downloading {len(codes)} funds ({args.period} of history)...")
    prices = download_prices(codes, args.period)
    if prices.empty:
        sys.exit("No data came back. Check your internet connection and the tickers, "
                 "then try again.")

    metrics = compute_metrics(prices, codes)

    # Record exactly when this data was pulled (uses your computer's clock)
    now = datetime.now()
    as_of = (f"Data pulled {now:%A %d %B %Y, %I:%M %p}. "
             f"Latest price date in data: {prices.index[-1]:%d %b %Y} "
             f"(may be mid-session if the market is open). "
             f"History requested: {args.period}.")

    ranked = metrics.dropna(subset=["Annual return %"]).sort_values(
        "Annual return %", ascending=False)
    with_data = metrics[metrics["Latest price"].notna()]
    no_data = [c for c in metrics.index if pd.isna(metrics.loc[c, "Latest price"])]

    pd.set_option("display.width", 220)
    pd.set_option("display.max_columns", None)
    print("\n" + as_of)
    shown = ["Exchange", "Data from", "Years", "Latest price", "Total return %",
             "Annual return %", "Volatility %", "Max drawdown %", "Return per risk"]
    ordered = pd.concat([ranked, with_data.drop(ranked.index)])  # short-history funds last
    print("\n" + ordered[shown].to_string())
    if no_data:
        print(f"\nNo usable price data for {len(no_data)} fund(s): {no_data}")
        print("They are still listed in the Excel file, with a note in the Status column.")

    # With dozens of funds the growth chart becomes unreadable, so by default it
    # shows only the best performers (the Excel file still lists every fund)
    pool = ranked if not ranked.empty else with_data
    chart_codes = list(pool.index[:args.chart_top]) if args.chart_top > 0 else list(pool.index)
    if chart_codes:
        plot_growth(prices[chart_codes], args.common_start, growth_file,
                    limited=len(chart_codes) < len(pool))
    made_scatter = plot_risk_return(metrics, risk_file)
    sector_df = compute_sector_summary(metrics)
    corr_df = compute_correlations(prices)
    save_excel(metrics, results_file, as_of, sector_df, corr_df,
               chart_files=(growth_file if chart_codes else None,
                            risk_file if made_scatter else None))
    build_html(prices, metrics, sector_df, corr_df, as_of, html_file,
               source=os.path.basename(args.funds))

    saved = [results_file, html_file] + ([growth_file] if chart_codes else []) \
            + ([risk_file] if made_scatter else [])
    print(f"\nSaved: {', '.join(saved)}")
    print(f"Open {html_file} in your browser for the portfolio builder.")
    print("Note: funds with under 1 year of history have no annualised figures.")


if __name__ == "__main__":
    main()
