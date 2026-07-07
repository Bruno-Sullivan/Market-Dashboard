#!/usr/bin/env python
# coding: utf-8

# In[2]:


from io import StringIO
from pathlib import Path
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import xml.etree.ElementTree as ET

import pandas as pd
import yfinance as yf
import requests

try:
    from IPython.display import display, HTML as JupyterHTML
    IN_JUPYTER = True
except Exception:
    IN_JUPYTER = False


HTML_FILE = Path("morning_dashboard.html")
EXCEL_FILE = Path("morning_market_tracker.xlsx")


# ============================================================
# MARKET DATA
# ============================================================

YAHOO_TICKERS = {
    # Equity indices
    "S&P 500": ("^GSPC", "Equity Indices", "index points"),
    "Dow Jones Industrial Average": ("^DJI", "Equity Indices", "index points"),
    "Nasdaq Composite": ("^IXIC", "Equity Indices", "index points"),
    "Russell 2000": ("^RUT", "Equity Indices", "index points"),
    "TSX Composite": ("^GSPTSE", "Equity Indices", "index points"),
    "FTSE 100": ("^FTSE", "Equity Indices", "index points"),
    "Nikkei 225": ("^N225", "Equity Indices", "index points"),
    "VIX": ("^VIX", "Equity Indices", "index points"),

    # Commodities
    "Gold Futures": ("GC=F", "Commodities", "USD/oz"),
    "WTI Crude Futures": ("CL=F", "Commodities", "USD/bbl"),
    "Brent Crude Futures": ("BZ=F", "Commodities", "USD/bbl"),
    "Copper Futures": ("HG=F", "Commodities", "USD/lb"),
    "Natural Gas Futures": ("NG=F", "Commodities", "USD/MMBtu"),
    "Silver Futures": ("SI=F", "Commodities", "USD/oz"),

    # Mag 7
    "Apple": ("AAPL", "Mag 7", "USD/share"),
    "Microsoft": ("MSFT", "Mag 7", "USD/share"),
    "Nvidia": ("NVDA", "Mag 7", "USD/share"),
    "Amazon": ("AMZN", "Mag 7", "USD/share"),
    "Alphabet": ("GOOGL", "Mag 7", "USD/share"),
    "Meta": ("META", "Mag 7", "USD/share"),
    "Tesla": ("TSLA", "Mag 7", "USD/share"),

    # Top Canadian large-cap stocks
    "Royal Bank of Canada": ("RY.TO", "Top Canadian Stocks", "CAD/share"),
    "Toronto-Dominion Bank": ("TD.TO", "Top Canadian Stocks", "CAD/share"),
    "Shopify": ("SHOP.TO", "Top Canadian Stocks", "CAD/share"),
    "Enbridge": ("ENB.TO", "Top Canadian Stocks", "CAD/share"),
    "Bank of Montreal": ("BMO.TO", "Top Canadian Stocks", "CAD/share"),
    "Brookfield Corp": ("BN.TO", "Top Canadian Stocks", "CAD/share"),
    "CIBC": ("CM.TO", "Top Canadian Stocks", "CAD/share"),
    "Agnico Eagle Mines": ("AEM.TO", "Top Canadian Stocks", "CAD/share"),
    "Scotiabank": ("BNS.TO", "Top Canadian Stocks", "CAD/share"),
    "Canadian Natural Resources": ("CNQ.TO", "Top Canadian Stocks", "CAD/share"),
}


def make_row(series, source, code, category, unit, date=None, value=None, prev_value=None, error=None):
    change = None
    change_pct = None

    if value is not None and prev_value is not None:
        change = value - prev_value
        if prev_value != 0:
            change_pct = change / abs(prev_value) * 100

    return {
        "Category": category,
        "Series": series,
        "Value": value,
        "Change": change,
        "Change_Pct": change_pct,
        "Date": date,
        "Source": source,
        "Code": code,
        "Unit": unit,
        "Updated_At": datetime.now(),
        "Error": error,
    }


def error_row(series, source, code, category, unit, error):
    return make_row(
        series=series,
        source=source,
        code=code,
        category=category,
        unit=unit,
        error=str(error),
    )


def latest_boc_value(code):
    url = f"https://www.bankofcanada.ca/valet/observations/{code}/json?recent=10"

    response = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=20,
    )
    response.raise_for_status()

    data = response.json()
    observations = data.get("observations", [])

    values = []

    for obs in observations:
        raw = obs.get(code, {}).get("v")

        if raw not in [None, "", "."]:
            values.append({
                "Date": pd.to_datetime(obs["d"]).date(),
                "Value": float(raw),
            })

    if not values:
        raise ValueError(f"No BoC data found for {code}")

    values = sorted(values, key=lambda x: x["Date"])

    latest = values[-1]
    previous = values[-2] if len(values) >= 2 else None

    return latest, previous


def boc_fx_rows():
    rows = []

    try:
        usdcad_latest, usdcad_prev = latest_boc_value("FXUSDCAD")
        eurcad_latest, eurcad_prev = latest_boc_value("FXEURCAD")
        gbpcad_latest, gbpcad_prev = latest_boc_value("FXGBPCAD")

        rows.append(
            make_row(
                series="CAD/USD",
                source="Bank of Canada",
                code="1 / FXUSDCAD",
                category="FX",
                unit="USD per CAD",
                date=usdcad_latest["Date"],
                value=1 / usdcad_latest["Value"],
                prev_value=1 / usdcad_prev["Value"] if usdcad_prev is not None else None,
            )
        )

        rows.append(
            make_row(
                series="CAD/EUR",
                source="Bank of Canada",
                code="1 / FXEURCAD",
                category="FX",
                unit="EUR per CAD",
                date=eurcad_latest["Date"],
                value=1 / eurcad_latest["Value"],
                prev_value=1 / eurcad_prev["Value"] if eurcad_prev is not None else None,
            )
        )

        rows.append(
            make_row(
                series="CAD/GBP",
                source="Bank of Canada",
                code="1 / FXGBPCAD",
                category="FX",
                unit="GBP per CAD",
                date=gbpcad_latest["Date"],
                value=1 / gbpcad_latest["Value"],
                prev_value=1 / gbpcad_prev["Value"] if gbpcad_prev is not None else None,
            )
        )

        rows.append(
            make_row(
                series="USD/EUR",
                source="Bank of Canada",
                code="FXUSDCAD / FXEURCAD",
                category="FX",
                unit="EUR per USD",
                date=max(usdcad_latest["Date"], eurcad_latest["Date"]),
                value=usdcad_latest["Value"] / eurcad_latest["Value"],
                prev_value=(
                    usdcad_prev["Value"] / eurcad_prev["Value"]
                    if usdcad_prev is not None and eurcad_prev is not None
                    else None
                ),
            )
        )

    except Exception as e:
        rows.append(error_row("FX Basket", "Bank of Canada", "BoC FX", "FX", "", e))

    return rows


def yahoo_rows_batch():
    ticker_to_info = {
        ticker: (series_name, category, unit)
        for series_name, (ticker, category, unit) in YAHOO_TICKERS.items()
    }

    tickers = list(ticker_to_info.keys())

    data = yf.download(
        tickers=tickers,
        period="7d",
        interval="1d",
        group_by="ticker",
        auto_adjust=False,
        progress=False,
        threads=True,
    )

    rows = []

    for ticker in tickers:
        series_name, category, unit = ticker_to_info[ticker]

        try:
            if isinstance(data.columns, pd.MultiIndex):
                close = data[ticker]["Close"].dropna()
            else:
                close = data["Close"].dropna()

            close = close.sort_index()

            if close.empty:
                raise ValueError("No close data returned")

            latest_value = float(close.iloc[-1])
            prev_value = float(close.iloc[-2]) if len(close) >= 2 else None
            latest_date = close.index[-1].date()

            rows.append(
                make_row(
                    series=series_name,
                    source="Yahoo Finance",
                    code=ticker,
                    category=category,
                    unit=unit,
                    date=latest_date,
                    value=latest_value,
                    prev_value=prev_value,
                )
            )

        except Exception as e:
            rows.append(
                error_row(
                    series=series_name,
                    source="Yahoo Finance",
                    code=ticker,
                    category=category,
                    unit=unit,
                    error=e,
                )
            )

    return rows


def add_average_move(df, series_name, component_names, category):
    components = df[
        (df["Series"].isin(component_names)) &
        (df["Error"].isna()) &
        (df["Change_Pct"].notna())
    ]

    if components.empty:
        return None

    avg_pct = float(components["Change_Pct"].mean())

    return {
        "Category": category,
        "Series": series_name,
        "Value": None,
        "Change": None,
        "Change_Pct": avg_pct,
        "Date": components["Date"].max(),
        "Source": "Calculated",
        "Code": "Average 1D % move",
        "Unit": "%",
        "Updated_At": datetime.now(),
        "Error": None,
    }


def build_dashboard():
    rows = []

    rows.extend(yahoo_rows_batch())
    rows.extend(boc_fx_rows())

    df = pd.DataFrame(rows)

    extra_rows = []

    mag7_names = [
        "Apple",
        "Microsoft",
        "Nvidia",
        "Amazon",
        "Alphabet",
        "Meta",
        "Tesla",
    ]

    canada_top_names = [
        "Royal Bank of Canada",
        "Toronto-Dominion Bank",
        "Shopify",
        "Enbridge",
        "Bank of Montreal",
        "Brookfield Corp",
        "CIBC",
        "Agnico Eagle Mines",
        "Scotiabank",
        "Canadian Natural Resources",
    ]

    mag7_avg = add_average_move(
        df,
        "Average 1D Move",
        mag7_names,
        category="Mag 7",
    )

    canada_avg = add_average_move(
        df,
        "Average 1D Move",
        canada_top_names,
        category="Top Canadian Stocks",
    )

    if mag7_avg is not None:
        extra_rows.append(mag7_avg)

    if canada_avg is not None:
        extra_rows.append(canada_avg)

    if extra_rows:
        df = pd.concat([df, pd.DataFrame(extra_rows)], ignore_index=True)

    order = [
        "Equity Indices",
        "FX",
        "Commodities",
        "Mag 7",
        "Top Canadian Stocks",
    ]

    df["Category_Order"] = df["Category"].apply(
        lambda x: order.index(x) if x in order else 999
    )

    df = df.sort_values(["Category_Order", "Category", "Series"]).drop(
        columns=["Category_Order"]
    )

    return df


def save_excel_dashboard(df):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="Latest", index=False)


# ============================================================
# LIVE RATES DATA
# ============================================================

def make_rate_row(country, series, value=None, prev_value=None, date=None, source=None, code=None, error=None):
    change = None

    if value is not None and prev_value is not None:
        change = value - prev_value

    return {
        "Country": country,
        "Series": series,
        "Value": value,
        "Change": change,
        "Date": date,
        "Source": source,
        "Code": code,
        "Error": error,
    }


def rate_error_row(country, series, source, code, error):
    return make_rate_row(
        country=country,
        series=series,
        source=source,
        code=code,
        error=str(error),
    )


CANADA_LOOKUP_SERIES = {
    "3 Month": ("V39065", "lookup_tbill_yields.php", "2012-10-16"),
    "6 Month": ("V39066", "lookup_tbill_yields.php", "2012-10-16"),
    "1 Year": ("V39067", "lookup_tbill_yields.php", "2012-10-16"),
    "2 Year": ("V39051", "lookup_bond_yields.php", "2012-02-06"),
    "5 Year": ("V39053", "lookup_bond_yields.php", "2012-02-06"),
    "10 Year": ("V39055", "lookup_bond_yields.php", "2012-02-06"),
    "Long Term": ("V39056", "lookup_bond_yields.php", "2012-02-06"),
}


def fetch_boc_lookup_series(series_name, code, lookup_page, start_reference):
    end_date = datetime.today().date()
    start_date = end_date - timedelta(days=45)

    url = (
        "https://www.bankofcanada.ca/stats/results//csv?"
        f"dF={start_date}&dT={end_date}"
        f"&lP={lookup_page}"
        "&rangeType=dates"
        f"&sR={start_reference}"
        f"&se=LOOKUPS_{code}"
    )

    response = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=15,
    )
    response.raise_for_status()

    lines = response.text.splitlines()

    header_index = None

    for i, line in enumerate(lines):
        cleaned = line.replace('"', "").strip()
        if cleaned.startswith(f"Date,{code}"):
            header_index = i
            break

    if header_index is None:
        raise ValueError(f"Could not find data table for {code}")

    df = pd.read_csv(StringIO("\n".join(lines[header_index:])))
    df.columns = ["Date", "Value"]

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Value"] = pd.to_numeric(
        df["Value"].replace([".", "NA", "Bank holiday"], pd.NA),
        errors="coerce",
    )

    df = df.dropna(subset=["Date", "Value"]).sort_values("Date")

    if df.empty:
        raise ValueError(f"No valid data returned for {code}")

    latest = df.iloc[-1]
    previous = df.iloc[-2] if len(df) >= 2 else None

    return make_rate_row(
        country="Canada",
        series=series_name,
        value=float(latest["Value"]),
        prev_value=float(previous["Value"]) if previous is not None else None,
        date=latest["Date"].date(),
        source="Bank of Canada",
        code=code,
    )


def canada_yield_rows():
    rows = []

    with ThreadPoolExecutor(max_workers=7) as executor:
        future_to_name = {}

        for series_name, (code, lookup_page, start_reference) in CANADA_LOOKUP_SERIES.items():
            future = executor.submit(
                fetch_boc_lookup_series,
                series_name,
                code,
                lookup_page,
                start_reference,
            )

            future_to_name[future] = (series_name, code)

        for future in as_completed(future_to_name):
            series_name, code = future_to_name[future]

            try:
                rows.append(future.result())
            except Exception as e:
                rows.append(rate_error_row("Canada", series_name, "Bank of Canada", code, e))

    return rows


CANADA_OVERNIGHT_SERIES = {
    "BoC Overnight Target Rate": "V39079",
    "CORRA Overnight Rate": "AVG.INTWO",
}


def fetch_boc_direct(series_name, code):
    url = f"https://www.bankofcanada.ca/valet/observations/{code}/json?recent=10"

    response = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=15,
    )
    response.raise_for_status()

    data = response.json()
    observations = data.get("observations", [])

    values = []

    for obs in observations:
        raw = obs.get(code, {}).get("v")

        if raw not in [None, "", "."]:
            values.append({
                "Date": pd.to_datetime(obs["d"]).date(),
                "Value": float(raw),
            })

    if not values:
        raise ValueError(f"No valid BoC direct data for {code}")

    values = sorted(values, key=lambda x: x["Date"])

    latest = values[-1]
    previous = values[-2] if len(values) >= 2 else None

    return make_rate_row(
        country="Canada",
        series=series_name,
        value=latest["Value"],
        prev_value=previous["Value"] if previous is not None else None,
        date=latest["Date"],
        source="Bank of Canada",
        code=code,
    )


def canada_overnight_rows():
    rows = []

    for series_name, code in CANADA_OVERNIGHT_SERIES.items():
        try:
            rows.append(fetch_boc_direct(series_name, code))
        except Exception as e:
            rows.append(rate_error_row("Canada", series_name, "Bank of Canada", code, e))

    return rows


US_TREASURY_MAP = {
    "3 Month": "BC_3MONTH",
    "6 Month": "BC_6MONTH",
    "1 Year": "BC_1YEAR",
    "2 Year": "BC_2YEAR",
    "5 Year": "BC_5YEAR",
    "10 Year": "BC_10YEAR",
    "30 Year": "BC_30YEAR",
}


def _localname(tag):
    return tag.split("}", 1)[-1].upper()


def fetch_us_treasury_table():
    years = [datetime.today().year, datetime.today().year - 1]
    records = []

    for year in years:
        url = (
            "https://home.treasury.gov/resource-center/data-chart-center/"
            "interest-rates/pages/xml"
            f"?data=daily_treasury_yield_curve&field_tdr_date_value={year}"
        )

        response = requests.get(
            url,
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=25,
        )
        response.raise_for_status()

        root = ET.fromstring(response.content)

        for elem in root.iter():
            if _localname(elem.tag) == "PROPERTIES":
                row = {}

                for child in list(elem):
                    row[_localname(child.tag)] = child.text

                if "NEW_DATE" in row:
                    records.append(row)

    if not records:
        raise ValueError("No U.S. Treasury XML records returned")

    df = pd.DataFrame(records)
    df["Date"] = pd.to_datetime(df["NEW_DATE"], errors="coerce")

    for col in US_TREASURY_MAP.values():
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["Date"]).sort_values("Date")

    return df


def us_treasury_rows():
    rows = []

    try:
        df = fetch_us_treasury_table()

        for series_name, col in US_TREASURY_MAP.items():
            try:
                if col not in df.columns:
                    raise ValueError(f"{col} not found in Treasury XML")

                temp = df[["Date", col]].copy()
                temp.columns = ["Date", "Value"]

                temp = temp.dropna(subset=["Date", "Value"]).sort_values("Date")

                if temp.empty:
                    raise ValueError(f"No valid Treasury data for {series_name}")

                latest = temp.iloc[-1]
                previous = temp.iloc[-2] if len(temp) >= 2 else None

                rows.append(
                    make_rate_row(
                        country="USA",
                        series=series_name,
                        value=float(latest["Value"]),
                        prev_value=float(previous["Value"]) if previous is not None else None,
                        date=latest["Date"].date(),
                        source="U.S. Treasury XML",
                        code=col,
                    )
                )

            except Exception as e:
                rows.append(rate_error_row("USA", series_name, "U.S. Treasury XML", col, e))

    except Exception as e:
        for series_name, col in US_TREASURY_MAP.items():
            rows.append(rate_error_row("USA", series_name, "U.S. Treasury XML", col, e))

    return rows


def _nyfed_reference_rate_row(series_name, endpoint, code):
    end_date = datetime.today().date()
    start_date = end_date - timedelta(days=20)

    url = (
        f"https://markets.newyorkfed.org/api/rates/{endpoint}/search.json"
        f"?startDate={start_date}&endDate={end_date}&type=rate"
    )

    try:
        response = requests.get(
            url,
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=25,
        )
        response.raise_for_status()

        payload = response.json()

        records = (
            payload.get("refRates")
            or payload.get("rates")
            or payload.get("data")
            or []
        )

        if not records:
            raise ValueError(f"No records returned from New York Fed API. Keys: {list(payload.keys())}")

        df = pd.DataFrame(records)
        df.columns = [str(c).strip() for c in df.columns]

        date_col = None
        for c in df.columns:
            cl = c.lower()
            if cl in ["effectivedate", "effective_date", "date"] or "date" in cl:
                date_col = c
                break

        rate_col = None
        for preferred in ["percentRate", "rate", "effectiveRate"]:
            if preferred in df.columns:
                rate_col = preferred
                break

        if rate_col is None:
            numeric_cols = []
            for c in df.columns:
                numeric = pd.to_numeric(df[c], errors="coerce")
                if numeric.notna().sum() >= 2:
                    numeric_cols.append(c)

            if numeric_cols:
                rate_col = numeric_cols[0]

        if date_col is None:
            raise ValueError(f"Could not find date column. Columns were: {list(df.columns)}")

        if rate_col is None:
            raise ValueError(f"Could not find rate column. Columns were: {list(df.columns)}")

        temp = df[[date_col, rate_col]].copy()
        temp.columns = ["Date", "Value"]

        temp["Date"] = pd.to_datetime(temp["Date"], errors="coerce")
        temp["Value"] = pd.to_numeric(temp["Value"], errors="coerce")

        temp = temp.dropna(subset=["Date", "Value"]).sort_values("Date")

        if temp.empty:
            raise ValueError("No valid rate rows after cleaning")

        latest = temp.iloc[-1]
        previous = temp.iloc[-2] if len(temp) >= 2 else None

        return make_rate_row(
            country="USA",
            series=series_name,
            value=float(latest["Value"]),
            prev_value=float(previous["Value"]) if previous is not None else None,
            date=latest["Date"].date(),
            source="New York Fed",
            code=code,
        )

    except Exception as e:
        return rate_error_row(
            "USA",
            series_name,
            "New York Fed",
            code,
            e,
        )


def fed_funds_row():
    return _nyfed_reference_rate_row(
        series_name="Effective Fed Funds Rate",
        endpoint="unsecured/effr",
        code="EFFR",
    )


def sofr_row():
    return _nyfed_reference_rate_row(
        series_name="SOFR Repo Rate",
        endpoint="secured/sofr",
        code="SOFR",
    )


def build_rates_dashboard():
    rows = []

    rows.extend(canada_yield_rows())
    rows.extend(canada_overnight_rows())

    rows.extend(us_treasury_rows())
    rows.append(fed_funds_row())
    rows.append(sofr_row())

    df = pd.DataFrame(rows)

    country_order = {
        "Canada": 0,
        "USA": 1,
    }

    series_order = {
        "3 Month": 0,
        "6 Month": 1,
        "1 Year": 2,
        "2 Year": 3,
        "5 Year": 4,
        "10 Year": 5,
        "30 Year": 6,
        "Long Term": 7,
        "BoC Overnight Target Rate": 8,
        "CORRA Overnight Rate": 9,
        "Effective Fed Funds Rate": 10,
        "SOFR Repo Rate": 11,
    }

    df["Country_Order"] = df["Country"].map(country_order).fillna(999)
    df["Series_Order"] = df["Series"].map(series_order).fillna(999)

    df = df.sort_values(["Country_Order", "Series_Order", "Series"]).drop(
        columns=["Country_Order", "Series_Order"]
    )

    return df


# ============================================================
# PRETTY HTML FORMATTERS
# ============================================================

def change_class(x):
    if pd.isna(x):
        return ""

    if x > 0:
        return "positive"

    if x < 0:
        return "negative"

    return ""


def pretty_format_value(row):
    value = row["Value"]
    unit = row.get("Unit", "")

    if pd.isna(value):
        return ""

    if unit == "%":
        return f"{value:,.2f}%"

    if unit == "index points":
        return f"{value:,.2f}"

    if "share" in str(unit):
        return f"{value:,.2f}"

    if (
        "USD" in str(unit)
        or "CAD" in str(unit)
        or "GBP" in str(unit)
        or "EUR" in str(unit)
    ):
        return f"{value:,.4f}"

    return f"{value:,.4f}"


def pretty_format_change(row):
    change = row["Change"]

    if pd.isna(change):
        return ""

    return f"{change:+,.4f}"


def pretty_format_rate(row):
    pct = row["Change_Pct"]

    if pd.isna(pct):
        return ""

    return f"{pct:+.2f}%"


def format_live_rate_value(x):
    if pd.isna(x):
        return ""
    return f"{x:.3f}%"


def format_live_rate_change(x):
    if pd.isna(x):
        return ""
    return f"{x * 100:+.1f} bps"


def build_live_rates_html(rates_df=None):
    try:
        if rates_df is None:
            rates_df = build_rates_dashboard()
    except Exception as e:
        return f"""
            <div class="manual-rates-title">Live Rates & Yields</div>
            <section class="sources-box">
                <h2>Rates Error</h2>
                <div class="source-line">
                    <span>Error</span>
                    <span>{str(e)}</span>
                </div>
            </section>
        """

    shown = rates_df.copy()

    shown["Value_Display"] = shown["Value"].apply(format_live_rate_value)
    shown["Change_Display"] = shown["Change"].apply(format_live_rate_change)
    shown["Change_Class"] = shown["Change"].apply(change_class)
    shown["Error"] = shown["Error"].fillna("")

    cards = ""

    for country in ["Canada", "USA"]:
        section = shown[
            (shown["Country"] == country) &
            (shown["Value"].notna())
        ].copy()

        if section.empty:
            continue

        rows_html = ""

        for _, row in section.iterrows():
            rows_html += f"""
                <tr>
                    <td class="series">{row['Series']}</td>
                    <td class="num">{row['Value_Display']}</td>
                    <td class="num {row['Change_Class']}">{row['Change_Display']}</td>
                </tr>
            """

        cards += f"""
            <section class="market-section live-rate-card">
                <h2>{country} Rates & Yields</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Series</th>
                            <th>Rate</th>
                            <th>Change</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows_html}
                    </tbody>
                </table>
            </section>
        """

    errors = shown[
        shown["Error"].notna() &
        (shown["Error"].astype(str).str.strip() != "")
    ]

    errors_html = ""

    if not errors.empty:
        error_rows = ""

        for _, row in errors.iterrows():
            error_rows += f"""
                <tr>
                    <td>{row['Country']}</td>
                    <td>{row['Series']}</td>
                    <td>{row['Source']}</td>
                    <td>{row['Code']}</td>
                    <td>{row['Error']}</td>
                </tr>
            """

        errors_html = f"""
            <section class="sources-box">
                <h2>Rates Errors</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Country</th>
                            <th>Series</th>
                            <th>Source</th>
                            <th>Code</th>
                            <th>Error</th>
                        </tr>
                    </thead>
                    <tbody>
                        {error_rows}
                    </tbody>
                </table>
            </section>
        """

    return f"""
        <div class="manual-rates-title">Live Rates & Yields</div>
        <div class="grid manual-rates-grid">
            {cards}
        </div>
        {errors_html}
    """

YIELD_CURVE_ORDER = {
    "3 Month": 0,
    "6 Month": 1,
    "1 Year": 2,
    "2 Year": 3,
    "5 Year": 4,
    "10 Year": 5,
    "30 Year": 6,
    "Long Term": 6,
}


def build_single_yield_curve_card(country, rates_df):
    curve = rates_df[
        (rates_df["Country"] == country) &
        (rates_df["Series"].isin(YIELD_CURVE_ORDER.keys())) &
        (rates_df["Value"].notna())
    ].copy()

    if "Error" in curve.columns:
        curve = curve[
            curve["Error"].isna() |
            (curve["Error"].astype(str).str.strip() == "")
        ]

    curve["Curve_Order"] = curve["Series"].map(YIELD_CURVE_ORDER)
    curve = curve.dropna(subset=["Curve_Order"]).sort_values(["Curve_Order", "Series"])

    if curve.shape[0] < 2:
        return f"""
            <section class="market-section yield-curve-card">
                <h2>{country} Yield Curve</h2>
                <div class="curve-empty">Not enough valid yield data to draw a curve.</div>
            </section>
        """

    width = 640
    height = 300
    left = 58
    right = 24
    top = 30
    bottom = 58

    plot_w = width - left - right
    plot_h = height - top - bottom

    y_min = float(curve["Value"].min())
    y_max = float(curve["Value"].max())

    if y_min == y_max:
        y_min -= 0.25
        y_max += 0.25
    else:
        pad = (y_max - y_min) * 0.18
        y_min -= pad
        y_max += pad

    points = []
    n = len(curve)

    for i, (_, row) in enumerate(curve.iterrows()):
        x = left + (plot_w * i / (n - 1))
        y = top + ((y_max - float(row["Value"])) / (y_max - y_min)) * plot_h

        points.append({
            "x": x,
            "y": y,
            "series": str(row["Series"]),
            "value": float(row["Value"]),
        })

    polyline_points = " ".join(
        f"{p['x']:.1f},{p['y']:.1f}" for p in points
    )

    y_ticks_html = ""

    for j in range(5):
        tick_value = y_min + (y_max - y_min) * j / 4
        y = top + ((y_max - tick_value) / (y_max - y_min)) * plot_h

        y_ticks_html += f"""
            <line class="grid-line" x1="{left}" y1="{y:.1f}" x2="{width - right}" y2="{y:.1f}" />
            <text class="axis-label" x="{left - 8}" y="{y + 3:.1f}" text-anchor="end">{tick_value:.2f}%</text>
        """

    x_labels_html = ""
    circles_html = ""
    value_labels_html = ""

    for p in points:
        x_labels_html += f"""
            <text class="axis-label" x="{p['x']:.1f}" y="{height - 24}" text-anchor="middle">{p['series']}</text>
        """

        circles_html += f"""
            <circle class="curve-point" cx="{p['x']:.1f}" cy="{p['y']:.1f}" r="4.5" />
        """

        value_y = p["y"] - 10

        if value_y < top + 8:
            value_y = p["y"] + 18

        value_labels_html += f"""
            <text class="curve-label" x="{p['x']:.1f}" y="{value_y:.1f}" text-anchor="middle">{p['value']:.2f}%</text>
        """

    latest_date = curve["Date"].max()

    return f"""
        <section class="market-section yield-curve-card">
            <h2>{country} Yield Curve</h2>
            <div class="curve-date">Latest available date: {latest_date}</div>

            <svg class="yield-curve-svg" viewBox="0 0 {width} {height}" role="img" aria-label="{country} yield curve">
                {y_ticks_html}

                <line class="axis-line" x1="{left}" y1="{height - bottom}" x2="{width - right}" y2="{height - bottom}" />
                <line class="axis-line" x1="{left}" y1="{top}" x2="{left}" y2="{height - bottom}" />

                <polyline class="curve-line" points="{polyline_points}" />

                {circles_html}
                {value_labels_html}
                {x_labels_html}
            </svg>
        </section>
    """


def build_yield_curve_html(rates_df=None):
    try:
        if rates_df is None:
            rates_df = build_rates_dashboard()

    except Exception as e:
        return f"""
            <div class="yield-curve-title">Yield Curves</div>
            <section class="sources-box">
                <h2>Yield Curve Error</h2>
                <div class="source-line">
                    <span>Error</span>
                    <span>{str(e)}</span>
                </div>
            </section>
        """

    cards = ""

    for country in ["Canada", "USA"]:
        cards += build_single_yield_curve_card(country, rates_df)

    return f"""
        <div class="yield-curve-title">Yield Curves</div>
        <div class="grid yield-curve-grid">
            {cards}
        </div>
    """

    shown = rates_df.copy()

    shown["Value_Display"] = shown["Value"].apply(format_live_rate_value)
    shown["Change_Display"] = shown["Change"].apply(format_live_rate_change)
    shown["Change_Class"] = shown["Change"].apply(change_class)
    shown["Error"] = shown["Error"].fillna("")

    cards = ""

    for country in ["Canada", "USA"]:
        section = shown[
            (shown["Country"] == country) &
            (shown["Value"].notna())
        ].copy()

        if section.empty:
            continue

        rows_html = ""

        for _, row in section.iterrows():
            rows_html += f"""
                <tr>
                    <td class="series">{row['Series']}</td>
                    <td class="num">{row['Value_Display']}</td>
                    <td class="num {row['Change_Class']}">{row['Change_Display']}</td>
                </tr>
            """

        cards += f"""
            <section class="market-section live-rate-card">
                <h2>{country} Rates & Yields</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Series</th>
                            <th>Rate</th>
                            <th>Change</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows_html}
                    </tbody>
                </table>
            </section>
        """

    errors = shown[
        shown["Error"].notna() &
        (shown["Error"].astype(str).str.strip() != "")
    ]

    errors_html = ""

    if not errors.empty:
        error_rows = ""

        for _, row in errors.iterrows():
            error_rows += f"""
                <tr>
                    <td>{row['Country']}</td>
                    <td>{row['Series']}</td>
                    <td>{row['Source']}</td>
                    <td>{row['Code']}</td>
                    <td>{row['Error']}</td>
                </tr>
            """

        errors_html = f"""
            <section class="sources-box">
                <h2>Rates Errors</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Country</th>
                            <th>Series</th>
                            <th>Source</th>
                            <th>Code</th>
                            <th>Error</th>
                        </tr>
                    </thead>
                    <tbody>
                        {error_rows}
                    </tbody>
                </table>
            </section>
        """

    return f"""
        <div class="manual-rates-title">Live Rates & Yields</div>
        <div class="grid manual-rates-grid">
            {cards}
        </div>
        {errors_html}
    """


def save_pretty_dashboard(df):
    shown = df.copy()

    shown["Series_Display"] = shown["Series"].astype(str)
    shown["Value_Display"] = shown.apply(pretty_format_value, axis=1)
    shown["Change_Display"] = shown.apply(pretty_format_change, axis=1)
    shown["Rate_Display"] = shown.apply(pretty_format_rate, axis=1)
    shown["Change_Class"] = shown["Change"].apply(change_class)
    shown["Rate_Class"] = shown["Change_Pct"].apply(change_class)

    category_order = [
        "Equity Indices",
        "FX",
        "Commodities",
        "Mag 7",
        "Top Canadian Stocks",
    ]

    categories = [
        c for c in category_order
        if c in shown["Category"].dropna().unique()
    ]

    extra_categories = [
        c for c in shown["Category"].dropna().unique()
        if c not in categories
    ]

    categories += extra_categories

    sections_html = ""

    for category in categories:
        section = shown[shown["Category"] == category].copy()

        section = section[
            section["Value"].notna() |
            section["Change"].notna() |
            section["Change_Pct"].notna()
        ]

        if section.empty:
            continue

        section["Is_Average_Row"] = section["Code"].eq("Average 1D % move").astype(int)

        section = section.sort_values(
            ["Is_Average_Row", "Series"],
            ascending=[True, True],
        )

        rows_html = ""

        for _, row in section.iterrows():
            avg_class = "avg-row" if row["Code"] == "Average 1D % move" else ""

            rows_html += f"""
                <tr class="{avg_class}">
                    <td class="series">{row['Series_Display']}</td>
                    <td class="num">{row['Value_Display']}</td>
                    <td class="num {row['Change_Class']}">{row['Change_Display']}</td>
                    <td class="num {row['Rate_Class']}">{row['Rate_Display']}</td>
                </tr>
            """

        sections_html += f"""
            <section class="market-section">
                <h2>{category}</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Series</th>
                            <th>Value</th>
                            <th>Change</th>
                            <th>Rate</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows_html}
                    </tbody>
                </table>
            </section>
        """

    try:
        rates_df = build_rates_dashboard()
    except Exception:
        rates_df = None

    live_rates_html = build_live_rates_html(rates_df)
    yield_curve_html = build_yield_curve_html(rates_df)

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Morning Market Dashboard</title>
        <style>
            body {{
                margin: 0;
                padding: 34px;
                background: #050505;
                color: #f7f7f7;
                font-family: Menlo, Consolas, Monaco, monospace;
            }}

            .dashboard {{
                max-width: 1500px;
                margin: 0 auto;
            }}

            h1 {{
                text-align: center;
                font-family: Georgia, serif;
                font-style: italic;
                font-weight: 400;
                font-size: 26px;
                letter-spacing: 0.5px;
                margin: 0 0 6px 0;
            }}

            .timestamp {{
                text-align: center;
                color: #bcbcbc;
                font-size: 12px;
                margin-bottom: 28px;
            }}

            .grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(430px, 1fr));
                gap: 22px;
            }}

            .market-section {{
                border: 2px solid #ffffff;
                border-radius: 14px;
                overflow: hidden;
                background: #0d0d0d;
                box-shadow: 0 0 18px rgba(255, 255, 255, 0.08);
            }}

            .market-section h2 {{
                margin: 0;
                padding: 12px 16px;
                border-bottom: 2px solid #ffffff;
                font-size: 15px;
                text-transform: uppercase;
                letter-spacing: 1.2px;
                background: #151515;
            }}

            table {{
                width: 100%;
                border-collapse: collapse;
            }}

            th {{
                padding: 9px 12px;
                border-bottom: 1px solid #ffffff;
                color: #ffffff;
                font-size: 12px;
                text-align: left;
                text-transform: uppercase;
                letter-spacing: 0.7px;
                background: #090909;
            }}

            td {{
                padding: 8px 12px;
                border-bottom: 1px solid #3d3d3d;
                font-size: 12px;
            }}

            tr:nth-child(even) {{
                background: #1b1b1b;
            }}

            tr:last-child td {{
                border-bottom: none;
            }}

            tr:hover {{
                background: #2a2a2a;
            }}

            .series {{
                width: 48%;
                font-weight: 600;
            }}

            .num {{
                text-align: right;
                font-variant-numeric: tabular-nums;
                white-space: nowrap;
            }}

            .positive {{
                color: #9cffb1 !important;
            }}

            .negative {{
                color: #ff9c9c !important;
            }}

            .avg-row td {{
                border-top: 2px solid #ffffff !important;
                background: #151515 !important;
                font-weight: 800;
            }}

            .avg-row .series {{
                text-transform: uppercase;
                letter-spacing: 0.6px;
            }}

            .manual-rates-title {{
                margin: 38px 0 18px 0;
                padding-top: 24px;
                border-top: 2px solid #ffffff;
                text-align: center;
                font-family: Georgia, serif;
                font-style: italic;
                font-size: 22px;
                letter-spacing: 0.8px;
            }}

                         .yield-curve-title {{
                margin: 38px 0 18px 0;
                padding-top: 24px;
                border-top: 2px solid #ffffff;
                text-align: center;
                font-family: Georgia, serif;
                font-style: italic;
                font-size: 22px;
                letter-spacing: 0.8px;
            }}

            .yield-curve-grid {{
                grid-template-columns: repeat(auto-fit, minmax(520px, 1fr));
            }}

            .yield-curve-card {{
                padding-bottom: 10px;
            }}

            .curve-date {{
                padding: 10px 16px 0 16px;
                color: #bcbcbc;
                font-size: 11px;
            }}

            .yield-curve-svg {{
                width: 100%;
                height: auto;
                display: block;
                padding: 6px 8px 10px 8px;
                box-sizing: border-box;
            }}

            .curve-line {{
                fill: none;
                stroke: #ffffff;
                stroke-width: 2.4;
            }}

            .curve-point {{
                fill: #0d0d0d;
                stroke: #ffffff;
                stroke-width: 2;
            }}

            .axis-line {{
                stroke: #ffffff;
                stroke-width: 1;
            }}

            .grid-line {{
                stroke: #333333;
                stroke-width: 1;
            }}

            .curve-label {{
                fill: #f7f7f7;
                font-size: 10px;
                font-weight: 700;
            }}

            .axis-label {{
                fill: #bcbcbc;
                font-size: 10px;
            }}

            .curve-empty {{
                padding: 18px;
                color: #bcbcbc;
                font-size: 12px;
            }}

            .manual-rates-grid {{
                grid-template-columns: repeat(auto-fit, minmax(380px, 1fr));
            }}

            .live-rate-card td:nth-child(2),
            .live-rate-card td:nth-child(3),
            .live-rate-card th:nth-child(2),
            .live-rate-card th:nth-child(3) {{
                text-align: right;
            }}

            .sources-box {{
                margin-top: 22px;
                border: 2px solid #ffffff;
                border-radius: 14px;
                background: #0d0d0d;
                padding: 16px 18px;
                box-shadow: 0 0 18px rgba(255, 255, 255, 0.08);
            }}

            .sources-box h2 {{
                margin: 0 0 12px 0;
                font-size: 14px;
                text-transform: uppercase;
                letter-spacing: 1.2px;
            }}

            .source-line {{
                display: grid;
                grid-template-columns: 180px 1fr;
                gap: 12px;
                font-size: 12px;
                padding: 5px 0;
                border-top: 1px solid #333333;
            }}

            .source-line span {{
                color: #f7f7f7;
                font-weight: 700;
            }}

            .sources-box td {{
                color: #ffb4b4;
                font-size: 11px;
            }}

            @media (max-width: 700px) {{
                body {{
                    padding: 18px;
                }}

                .grid {{
                    grid-template-columns: 1fr;
                }}

                .source-line {{
                    grid-template-columns: 1fr;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="dashboard">
            <h1>Morning Market Dashboard</h1>
            <div class="timestamp">Updated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div>

            <div class="grid">
                {sections_html}
            </div>

            {live_rates_html}

            {yield_curve_html}
        </div>
    </body>
    </html>
    """

    HTML_FILE.write_text(html, encoding="utf-8")

    if IN_JUPYTER:
        display(JupyterHTML(html))



dashboard = build_dashboard()
save_excel_dashboard(dashboard)
save_pretty_dashboard(dashboard)


# In[ ]:


jupyter nbconvert --to script "Market Daily.ipynb"

