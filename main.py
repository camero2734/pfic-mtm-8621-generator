import argparse
import html
import getpass
import glob
import json
import logging
import os
import re
import sys
from decimal import Decimal

import numpy as np
import pandas as pd
import pandas_datareader.data as web
import pikepdf
import weasyprint

# ---------------------------------------------------------------------------
# Script directory (for locating template PDF regardless of cwd)
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Value used to mark a checkbox as checked (PDF name /1)
CHECKED = "/1"

# ---------------------------------------------------------------------------
# AcroForm field name constants (as they appear in f8621.pdf widget /T keys)
# ---------------------------------------------------------------------------

# Page 1 – personal info
F_NAME = "f1_1[0]"  # Name of shareholder
F_ADDRESS_1 = "f1_2[0]"  # Street address line 1
F_ADDRESS_2 = "f1_3[0]"  # Street address line 2 (city/state/zip overflow)
F_CITY_OR_TOWN = "f1_4[0]"  # City
F_STATE = "f1_5[0]"  # State
F_COUNTRY = "f1_6[0]"  # Country
F_POSTAL_CODE = "f1_7[0]"  # Postal code
F_TAX_YEAR = "f1_9[0]"  # 2-digit tax year (top-right corner)
F_IDENTIFYING_NUM = "f1_8[0]"  # SSN / EIN

# Page 1 – shareholder type checkboxes (c1_1[0..5] = individual/partnership/S-corp/trust/estate/other)
C_SHAREHOLDER_INDIVIDUAL = "c1_1[0]"

# Page 1 – PFIC info
F_PFIC_NAME = "f1_14[0]"  # Name of PFIC
F_PFIC_ADDRESS = "f1_15[0]"  # Address of PFIC (multi-line field)
F_PFIC_REF_ID = "f1_17[0]"  # Reference ID number of PFIC
F_PFIC_SHARE_CLASS = "f1_23[0]"  # Description of each class of shares

# Page 1 – Part I
F_DATE_ACQUISITION = "f1_24[0]"  # Date of acquisition
F_NUM_SHARES = "f1_25[0]"  # Number of shares
F_AMOUNT_1291 = "f1_27[0]"  # Amount subject to section 1291
F_AMOUNT_1293 = "f1_28[0]"  # Amount subject to section 1293
F_AMOUNT_1296 = "f1_29[0]"  # Amount subject to section 1296

# Page 1 – Part I value-of-PFIC checkboxes (≤$50k / $50k–$100k / $100k–$150k / $150k–$200k)
C_VALUE_LE_50K = "c1_5[0]"
C_VALUE_50_100K = "c1_5[1]"
C_VALUE_100_150K = "c1_5[2]"
C_VALUE_150_200K = "c1_5[3]"
F_VALUE_OVER_200K = "f1_26[0]"  # free-text field for values > $200k

# Page 1 – Part I section type c checkbox
C_SECTION_TYPE_C = "c1_8[0]"

# Page 1 – Part II (MTM election)
C_PART2_MTM = "c1_11[0]"

# Page 2 – Part IV (MTM annual calculations, one page per lot)
F_10A = "f2_15[0]"  # FMV at year-end
F_10B = "f2_16[0]"  # Adjusted basis at year-end
F_10C = "f2_17[0]"  # Gain (loss) from line 10a - 10b
F_11 = "f2_18[0]"  # Unreversed inclusions (holding)
F_12 = "f2_19[0]"  # Ordinary loss limited by line 11 (holding)
F_13A = "f2_20[0]"  # Sale proceeds
F_13B = "f2_21[0]"  # Adjusted basis at date of sale
F_13C = "f2_22[0]"  # Gain (loss) from line 13a - 13b
F_14A = "f2_23[0]"  # Unreversed inclusions (sale)
F_14B = "f2_24[0]"  # Ordinary loss limited by line 14a (sale)
F_14C = "f2_25[0]"  # Capital loss (sale, basis ≤ original)

# ---------------------------------------------------------------------------
# Expected XLSX column names
# ---------------------------------------------------------------------------
EOY_COLUMNS = ["Year", "Price"]
PFIC_COLUMNS = ["PFIC Name", "PFIC Address", "PFIC Reference ID", "PFIC Share Class", "Currency"]
TRANSACTION_COLUMNS = [
    "Date",
    "Type",
    "Number of shares",
    "Total Value",
]
BUY_TYPES = {"buy", "purchase", "reinvestment"}
SELL_TYPES = {"sell", "sale", "distribution"}

# FRED exchange rate series codes.
# These map a currency code to the FRED series that gives "USD per 1 unit of foreign currency".
# Rates are expressed as: foreign_amount * rate = USD_amount.
FRED_FX_SERIES = {
    "AUD": "DEXUSAL",
    "BRL": "DEXBZUS",
    "CAD": "DEXCAUS",
    "CHF": "DEXSZUS",
    "CNY": "DEXCHUS",
    "DKK": "DEXDNUS",
    "EUR": "DEXUSEU",
    "GBP": "DEXUSUK",
    "HKD": "DEEXHKUS",
    "INR": "DEXINUS",
    "JPY": "DEXJPUS",
    "KRW": "DEXKOUS",
    "MXN": "DEXMXUS",
    "NOK": "DEXNOUS",
    "NZD": "DEXUSNZ",
    "SEK": "DEXSDUS",
    "SGD": "DEXSIUS",
    "TWD": "DEXTWUS",
    "THB": "DEXTHUS",
    "ZAR": "DEEXZAUS",
}

# Cache: {currency: {date: rate}}
_fx_cache: dict = {}


def get_exchange_rate(currency: str, date) -> float:
    """Look up the USD exchange rate for a currency on a given date via FRED.
    Returns the rate such that: USD_amount = foreign_amount * rate.
    For currencies with a direct USD series (e.g. DEXUSEU for EUR), uses it directly.
    For others, falls back to converting via EUR.
    """
    currency = currency.upper().strip()
    if currency == "USD":
        return 1.0

    date = pd.Timestamp(date)
    if currency not in _fx_cache or date not in _fx_cache[currency]:
        _prefetch_rates(currency, date)

    if currency in _fx_cache and date in _fx_cache[currency]:
        rate = _fx_cache[currency][date]
        logging.info(f"  💱 {currency}/USD on {date.date()}: {rate:.4f}")
        return rate

    raise ValueError(
        f"Could not find exchange rate for {currency} on {date.date()}. "
        f"Ensure the date is a business day or add the rate manually."
    )


def _prefetch_rates(currency: str, date):
    """Fetch a range of rates from FRED for the given currency around the date.
    FRED only publishes rates on US business days, so we fetch a small window
    and use the nearest available rate.
    """
    if currency not in FRED_FX_SERIES:
        raise ValueError(
            f"Currency '{currency}' is not in the FRED series mapping. "
            f"Supported currencies: {sorted(FRED_FX_SERIES.keys())}"
        )

    series = FRED_FX_SERIES[currency]
    start = date - pd.Timedelta(days=10)
    end = date + pd.Timedelta(days=10)

    try:
        df = web.DataReader(series, "fred", start, end)
    except Exception as e:
        raise ValueError(f"Failed to fetch FRED series {series}: {e}") from e

    if df.empty:
        raise ValueError(f"No data returned from FRED for {series} around {date.date()}")

    series_col = df.columns[0]

    # Most FRED FX series are "foreign currency per USD" (e.g. DEXJPUS = JPY/USD).
    # A few are "USD per foreign unit" (DEXUSEU = USD/EUR, DEXUSUK = USD/GBP, etc.).
    # We normalize everything to "USD per 1 unit of foreign currency".
    # Series starting with DEXUS are already USD/foreign; others need inversion.
    if series.startswith("DEXUS") or series.startswith("DEEX"):
        usd_per_foreign = df[series_col]
    else:
        usd_per_foreign = 1.0 / df[series_col]

    if currency not in _fx_cache:
        _fx_cache[currency] = {}

    for idx, val in usd_per_foreign.items():
        if pd.notna(val):
            _fx_cache[currency][pd.Timestamp(idx)] = float(val)

    # If exact date not available (weekend/holiday), use nearest business day
    if date not in _fx_cache[currency]:
        nearest = min(_fx_cache[currency].keys(), key=lambda d: abs(d - date))
        _fx_cache[currency][date] = _fx_cache[currency][nearest]


# ---------------------------------------------------------------------------
# Data classes for computation results
# ---------------------------------------------------------------------------


class Part1Result:
    """Results from Part I computation."""

    def __init__(self, date_of_acq, unsold_shares, value_of_pfic):
        self.date_of_acq = date_of_acq
        self.unsold_shares = unsold_shares
        self.value_of_pfic = value_of_pfic


class LotResult:
    """Results from Part IV computation for a single lot."""

    def __init__(
        self,
        lot_index,
        is_holding,
        fmv,
        adjusted_basis,
        original_basis,
        gain_loss,
        proceeds=None,
        sale_gain_loss=None,
        unreversed=None,
        ordinary_loss=None,
        capital_loss=None,
        skipped=False,
        roll_forward=None,
        calc_detail=None,
    ):
        self.lot_index = lot_index
        self.is_holding = is_holding
        self.fmv = fmv
        self.adjusted_basis = adjusted_basis
        self.original_basis = original_basis
        self.gain_loss = gain_loss
        self.proceeds = proceeds
        self.sale_gain_loss = sale_gain_loss
        self.unreversed = unreversed
        self.ordinary_loss = ordinary_loss
        self.capital_loss = capital_loss
        self.skipped = skipped
        self.roll_forward = roll_forward or []
        self.calc_detail = calc_detail or {}

    @property
    def lot_summary(self) -> dict:
        summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}
        if self.skipped:
            return summary
        if self.is_holding:
            if self.gain_loss < 0:
                if self.ordinary_loss is not None and self.ordinary_loss != 0:
                    summary["ordinary_losses"] += abs(self.ordinary_loss)
            else:
                summary["ordinary_gains"] += self.gain_loss
        else:
            if self.sale_gain_loss is not None and self.sale_gain_loss >= 0:
                summary["ordinary_gains"] += self.sale_gain_loss
            elif self.sale_gain_loss is not None and self.sale_gain_loss < 0:
                if self.ordinary_loss is not None and self.ordinary_loss != 0:
                    summary["ordinary_losses"] += abs(self.ordinary_loss)
                if self.capital_loss is not None and self.capital_loss != 0:
                    summary["capital_losses"] += abs(self.capital_loss)
        return summary


# ---------------------------------------------------------------------------
# Input validation
# ---------------------------------------------------------------------------


def validate_tax_year(tax_year_str: str) -> int:
    """Parse and validate a 2-digit tax year string. Returns the full year (e.g. 2025)."""
    tax_year_str = tax_year_str.strip()
    if not tax_year_str.isdigit():
        logging.error("💥 Tax year must be numeric (e.g. '25' for 2025).")
        sys.exit(1)
    year_int = int(tax_year_str)
    if year_int < 0 or year_int > 99:
        logging.error("💥 Tax year must be a 2-digit number between 00 and 99.")
        sys.exit(1)
    return 2000 + year_int


def validate_xlsx_columns(df: pd.DataFrame, expected: list, sheet_name: str, filepath: str):
    """Validate that a DataFrame has all expected columns. Exits with error if missing."""
    missing = [col for col in expected if col not in df.columns]
    if missing:
        logging.error(f"💥 Missing columns in sheet '{sheet_name}' of {filepath}: {', '.join(missing)}")
        logging.error(f"   Expected: {', '.join(expected)}")
        logging.error(f"   Found:    {', '.join(df.columns)}")
        sys.exit(1)


def validate_reference_id(ref_id: str):
    """Validate that the reference ID is alphanumeric per IRS rules."""
    if not re.match(r"^[A-Z0-9]{1,50}$", ref_id, re.IGNORECASE):
        logging.error(
            f"💥 Invalid PFIC Reference ID: '{ref_id}'. "
            "Must be 1-50 characters, alphanumeric only (A-Z, 0-9), "
            "no spaces or special characters."
        )
        sys.exit(1)


def get_eoy_value(df_eoy: pd.DataFrame, year: int, column: str, filepath: str):
    """Safely look up an EOY value for a given year. Exits with error if year is missing."""
    rows = df_eoy[df_eoy["Year"] == year]
    if len(rows) == 0:
        logging.error(
            f"💥 No EOY data found for year {year} in {filepath}. "
            f"Available years: {', '.join(str(y) for y in df_eoy['Year'].values)}"
        )
        sys.exit(1)
    return rows[column].values[0]


# ---------------------------------------------------------------------------
# FIFO lot construction from transactions
# ---------------------------------------------------------------------------


def fifo_lots_from_transactions(df_txn: pd.DataFrame, currency: str) -> pd.DataFrame:
    """Convert a DataFrame of buy/sell transactions into Lot Details rows
    using FIFO share matching.  Buys open lots; sells close lots in
    buy-date order, splitting lots when a partial fill is needed.
    Exchange rates are looked up from FRED based on the currency."""

    txns = df_txn.sort_values("Date").reset_index(drop=True)

    open_lots = []
    for _, row in txns.iterrows():
        ttype = str(row["Type"]).strip().lower()
        if ttype in BUY_TYPES:
            total_value = row["Total Value"]
            num_shares = Decimal(str(row["Number of shares"]))
            er = get_exchange_rate(currency, row["Date"])
            open_lots.append(
                {
                    "Date: Acquisition": row["Date"],
                    "Price per share: Acquisition": total_value / float(num_shares),
                    "Number of shares": num_shares,
                    "Cost: Acquisition": total_value,
                    "Exchange Rate: Acquisition": er,
                    "Date: Sale": np.nan,
                    "Price per share: Sale": np.nan,
                    "Exchange Rate: Sale": np.nan,
                    "_remaining": num_shares,
                }
            )
        elif ttype in SELL_TYPES:
            to_sell = Decimal(str(row["Number of shares"]))
            sale_date = row["Date"]
            sale_price = row["Total Value"] / row["Number of shares"]
            sale_er = get_exchange_rate(currency, sale_date)
            new_lots = []
            for lot in open_lots:
                if to_sell <= 0:
                    new_lots.append(lot)
                    continue
                if lot["_remaining"] <= 0:
                    new_lots.append(lot)
                    continue
                fill = min(lot["_remaining"], to_sell)
                cost_per_share = lot["Cost: Acquisition"] / float(lot["_remaining"])
                if fill < lot["_remaining"]:
                    new_lots.append(
                        {
                            "Date: Acquisition": lot["Date: Acquisition"],
                            "Price per share: Acquisition": lot["Price per share: Acquisition"],
                            "Number of shares": fill,
                            "Cost: Acquisition": float(fill) * cost_per_share,
                            "Exchange Rate: Acquisition": lot["Exchange Rate: Acquisition"],
                            "Date: Sale": sale_date,
                            "Price per share: Sale": sale_price,
                            "Exchange Rate: Sale": sale_er,
                            "_remaining": Decimal(0),
                        }
                    )
                    remaining = lot["_remaining"] - fill
                    lot["Number of shares"] = remaining
                    lot["Cost: Acquisition"] = float(remaining) * cost_per_share
                    lot["_remaining"] = remaining
                    new_lots.append(lot)
                else:
                    lot["Number of shares"] = fill
                    lot["Cost: Acquisition"] = float(fill) * cost_per_share
                    lot["Date: Sale"] = sale_date
                    lot["Price per share: Sale"] = sale_price
                    lot["Exchange Rate: Sale"] = sale_er
                    lot["_remaining"] = Decimal(0)
                    new_lots.append(lot)
                to_sell -= fill
            open_lots = new_lots
            if to_sell > 0:
                logging.warning(
                    f"⚠️  Sell of {row['Number of shares']} shares on {sale_date} "
                    f"exceeds available lots by {to_sell} shares."
                )

    result = []
    for lot in open_lots:
        if lot["_remaining"] > 0:
            lot["Number of shares"] = lot["_remaining"]
        del lot["_remaining"]
        result.append(lot)

    return pd.DataFrame(result)


# ---------------------------------------------------------------------------
# Computation logic (shared between PDF and text output)
# ---------------------------------------------------------------------------


def compute_part1(df_lot: pd.DataFrame, df_eoy: pd.DataFrame, current_year: int, filepath: str) -> Part1Result:
    """Compute Part I results (shared logic for both PDF and text)."""
    date_of_acq = (
        pd.to_datetime(df_lot["Date: Acquisition"].values[0]).strftime("%Y-%m-%d")
        if len(df_lot.index) == 1
        else "Multiple"
    )
    unsold_shares = Decimal(0)
    for lot in range(len(df_lot.index)):
        if np.isnan(df_lot["Price per share: Sale"][lot]):
            unsold_shares += Decimal(str(df_lot["Number of shares"][lot]))

    last_er = get_eoy_value(df_eoy, current_year, "Exchange Rate", filepath)
    last_price = get_eoy_value(df_eoy, current_year, "Price", filepath)
    value_of_pfic = round(float(unsold_shares) * last_price * last_er)

    return Part1Result(date_of_acq, unsold_shares, value_of_pfic)


def compute_lot(
    df_lot: pd.DataFrame,
    df_eoy: pd.DataFrame,
    lot: int,
    current_year: int,
    filepath: str,
) -> LotResult:
    """Compute Part IV results for a single lot. Returns LotResult with skipped=True
    if lot was sold in a prior year."""
    year_of_acquisition = df_lot["Date: Acquisition"][lot].year
    cost_acquisition = df_lot["Cost: Acquisition"][lot]
    er_of_acquisition = df_lot["Exchange Rate: Acquisition"][lot]
    num_shares = float(df_lot["Number of shares"][lot])
    original_basis = cost_acquisition * er_of_acquisition

    adb = original_basis
    uni = 0.0
    roll_forward = []
    for year in range(year_of_acquisition, current_year):
        price = get_eoy_value(df_eoy, year, "Price", filepath)
        fx = get_eoy_value(df_eoy, year, "Exchange Rate", filepath)
        fmv = round(num_shares * price * fx)
        raw_mtm = fmv - adb
        adb_before = adb
        if raw_mtm >= 0:
            adb = adb + raw_mtm
            uni = uni + raw_mtm
            allowed_loss = 0
        else:
            allowed_loss = min(-raw_mtm, uni)
            adb = adb - allowed_loss
            uni = uni - allowed_loss
        roll_forward.append(
            {
                "year": year,
                "eoy_price": price,
                "eoy_fx": fx,
                "fmv": fmv,
                "adb_begin": round(adb_before),
                "raw_mtm": raw_mtm,
                "allowed_loss": allowed_loss,
                "adb_end": round(adb),
                "uni_end": round(uni),
            }
        )
    adjusted_basis = round(adb)
    unreversed_amount = round(uni)

    if not np.isnan(df_lot["Price per share: Sale"][lot]):
        # Sold lot
        sale_er = df_lot["Exchange Rate: Sale"][lot]
        sale_price = df_lot["Price per share: Sale"][lot]
        year_of_sale = df_lot["Date: Sale"][lot].year

        if year_of_sale < current_year:
            return LotResult(
                lot_index=lot,
                is_holding=False,
                fmv=0,
                adjusted_basis=adjusted_basis,
                original_basis=original_basis,
                gain_loss=0,
                skipped=True,
                roll_forward=roll_forward,
            )

        proceeds = round(num_shares * sale_price * sale_er)
        sale_gain_loss = proceeds - adjusted_basis

        calc_detail = {
            "shares": num_shares,
            "sale_price": sale_price,
            "sale_fx": sale_er,
            "sale_date": str(df_lot["Date: Sale"][lot].date())
            if hasattr(df_lot["Date: Sale"][lot], "date")
            else str(df_lot["Date: Sale"][lot]),
        }

        if sale_gain_loss < 0:
            if unreversed_amount > 0:
                unreversed = unreversed_amount
                ordinary_loss = -min(unreversed, -sale_gain_loss)
                capital_loss = None
                logging.info(f"    📉 Lot {lot + 1}: Ordinary loss of ${abs(ordinary_loss)}")
            else:
                unreversed = None
                ordinary_loss = 0
                capital_loss = sale_gain_loss
                logging.info(f"    📉 Lot {lot + 1}: Capital loss of ${abs(sale_gain_loss)}")
        else:
            unreversed = None
            ordinary_loss = None
            capital_loss = None
            logging.info(f"    📈 Lot {lot + 1}: Ordinary gain of ${sale_gain_loss}")

        return LotResult(
            lot_index=lot,
            is_holding=False,
            fmv=0,
            adjusted_basis=adjusted_basis,
            original_basis=original_basis,
            gain_loss=0,
            proceeds=proceeds,
            sale_gain_loss=sale_gain_loss,
            unreversed=unreversed,
            ordinary_loss=ordinary_loss,
            capital_loss=capital_loss,
            roll_forward=roll_forward,
            calc_detail=calc_detail,
        )

    else:
        # Holding at year-end
        last_er = get_eoy_value(df_eoy, current_year, "Exchange Rate", filepath)
        last_price = get_eoy_value(df_eoy, current_year, "Price", filepath)
        fmv = round(num_shares * last_price * last_er)

        gain_loss = fmv - adjusted_basis
        logging.info(f"    📈 Lot {lot + 1}: No sale (holding position)")

        calc_detail = {
            "shares": num_shares,
            "eoy_price": last_price,
            "eoy_fx": last_er,
            "eoy_date": f"{current_year}-12-31",
        }

        if gain_loss < 0:
            if unreversed_amount > 0:
                unreversed = unreversed_amount
                ordinary_loss = -min(unreversed, -gain_loss)
                logging.info(f"    📉 Lot {lot + 1}: Ordinary loss of ${abs(ordinary_loss)}")
            else:
                unreversed = 0
                ordinary_loss = 0
                logging.info(f"    📉 Lot {lot + 1}: Unrecognizable loss of ${abs(gain_loss)}")
            capital_loss = None
        else:
            unreversed = None
            ordinary_loss = None
            capital_loss = None
            logging.info(f"    📈 Lot {lot + 1}: Ordinary gain of ${gain_loss}")

        return LotResult(
            lot_index=lot,
            is_holding=True,
            fmv=fmv,
            adjusted_basis=adjusted_basis,
            original_basis=original_basis,
            gain_loss=gain_loss,
            unreversed=unreversed,
            ordinary_loss=ordinary_loss,
            capital_loss=capital_loss,
            roll_forward=roll_forward,
            calc_detail=calc_detail,
        )


# ---------------------------------------------------------------------------
# Form section builders  (return dicts of field_name -> value)
# ---------------------------------------------------------------------------


def _personal_info_fields(data_dict: dict) -> dict:
    fields = {
        F_NAME: data_dict["Name of shareholder"],
        F_ADDRESS_1: data_dict["Address"],
        F_CITY_OR_TOWN: data_dict["City"],
        F_STATE: data_dict["State"],
        F_COUNTRY: data_dict["Country"],
        F_POSTAL_CODE: data_dict["Postal Code"],
        F_IDENTIFYING_NUM: data_dict["Identifying Number"],
        F_TAX_YEAR: data_dict["Tax year"],
        C_SHAREHOLDER_INDIVIDUAL: CHECKED,  # always individual
    }
    # Fill address line 2 if provided
    if data_dict.get("Address line 2"):
        fields[F_ADDRESS_2] = data_dict["Address line 2"]
    return fields


def _pfic_info_fields(df_pfic: pd.DataFrame) -> dict:
    return {
        F_PFIC_NAME: str(df_pfic["PFIC Name"].values[0]),
        F_PFIC_ADDRESS: str(df_pfic["PFIC Address"].values[0]),
        F_PFIC_REF_ID: str(df_pfic["PFIC Reference ID"].values[0]),
        F_PFIC_SHARE_CLASS: str(df_pfic["PFIC Share Class"].values[0]),
    }


def _part1_fields(part1: Part1Result) -> dict:
    """Build PDF field dict from Part1Result."""
    fields = {
        F_DATE_ACQUISITION: str(part1.date_of_acq),
        F_NUM_SHARES: str(part1.unsold_shares),
        F_AMOUNT_1291: "",
        F_AMOUNT_1293: "",
        F_AMOUNT_1296: str(part1.value_of_pfic),
        C_SECTION_TYPE_C: CHECKED,
        C_PART2_MTM: CHECKED,
    }

    # Value-of-PFIC checkboxes
    v = part1.value_of_pfic
    if 0 <= v <= 50_000:
        fields[C_VALUE_LE_50K] = CHECKED
    elif v <= 100_000:
        fields[C_VALUE_50_100K] = CHECKED
    elif v <= 150_000:
        fields[C_VALUE_100_150K] = CHECKED
    elif v <= 200_000:
        fields[C_VALUE_150_200K] = CHECKED
    else:
        fields[F_VALUE_OVER_200K] = str(v)

    return fields


def _lot_fields(lot_result: LotResult) -> dict | None:
    """Build PDF field dict for a single lot. Returns None if the lot should be skipped."""
    if lot_result.skipped:
        return None

    fields = {}

    if lot_result.is_holding:
        fields[F_10A] = str(lot_result.fmv)
        fields[F_10B] = str(lot_result.adjusted_basis)
        fields[F_10C] = str(lot_result.gain_loss)

        if lot_result.gain_loss < 0:
            fields[F_11] = str(lot_result.unreversed)
            fields[F_12] = str(lot_result.ordinary_loss)
        else:
            fields[F_11] = ""
            fields[F_12] = ""

        fields.update({F_13A: "", F_13B: "", F_13C: "", F_14A: "", F_14B: "", F_14C: ""})

    else:
        # Sold lot
        fields[F_13A] = str(lot_result.proceeds)
        fields[F_13B] = str(lot_result.adjusted_basis)
        fields[F_13C] = str(lot_result.sale_gain_loss)

        if lot_result.sale_gain_loss < 0:
            if lot_result.unreversed is not None and lot_result.unreversed > 0:
                fields[F_14A] = str(lot_result.unreversed)
                fields[F_14B] = str(lot_result.ordinary_loss)
                fields[F_14C] = ""
            else:
                fields[F_14A] = "0"
                fields[F_14B] = "0"
                fields[F_14C] = str(lot_result.capital_loss)
        else:
            fields.update({F_14A: "", F_14B: "", F_14C: ""})

        fields.update({F_10A: "", F_10B: "", F_10C: "", F_11: "", F_12: ""})

    return fields


# ---------------------------------------------------------------------------
# Data loading helper
# ---------------------------------------------------------------------------


def load_xlsx(xlsx_path: str) -> tuple:
    """Load and validate an XLSX input file. Returns (df_lot, df_eoy, df_pfic).
    Exchange rates are looked up from FRED based on the Currency field in PFIC Details.
    """
    logging.info(f"  📂 Reading {xlsx_path}")
    xls = pd.ExcelFile(xlsx_path)

    df_txn = pd.read_excel(xls, sheet_name="Transactions")
    validate_xlsx_columns(df_txn, TRANSACTION_COLUMNS, "Transactions", xlsx_path)

    df_eoy = pd.read_excel(xls, sheet_name="EOY Details")
    validate_xlsx_columns(df_eoy, EOY_COLUMNS, "EOY Details", xlsx_path)

    df_pfic = pd.read_excel(xls, sheet_name="PFIC Details")
    validate_xlsx_columns(df_pfic, PFIC_COLUMNS, "PFIC Details", xlsx_path)

    validate_reference_id(str(df_pfic["PFIC Reference ID"].values[0]))

    currency = str(df_pfic["Currency"].values[0]).strip().upper()
    logging.info(f"  💱 Currency: {currency}")

    df_lot = fifo_lots_from_transactions(df_txn, currency)
    df_eoy = add_eoy_exchange_rates(df_eoy, currency)

    return df_lot, df_eoy, df_pfic


def add_eoy_exchange_rates(df_eoy: pd.DataFrame, currency: str) -> pd.DataFrame:
    """Add an Exchange Rate column to the EOY DataFrame by looking up
    December 31 exchange rates from FRED for each year."""
    df_eoy = df_eoy.copy()
    rates = []
    for _, row in df_eoy.iterrows():
        date = pd.Timestamp(year=int(row["Year"]), month=12, day=31)
        rate = get_exchange_rate(currency, date)
        rates.append(rate)
    df_eoy["Exchange Rate"] = rates
    return df_eoy


# ---------------------------------------------------------------------------
# PDF output
# ---------------------------------------------------------------------------


def create_filled_pdf(output_path: str, data_dict: dict, xlsx: str):
    """
    Build and save a filled Form 8621 PDF directly via AcroForm fields.
    Returns (number_of_lots, pfic_summary).
    """
    tax_year = validate_tax_year(data_dict["Tax year"])
    df_lot, df_eoy, df_pfic = load_xlsx(xlsx)

    number_of_lots = len(df_lot.index)
    logging.info(f"  📊 Found {number_of_lots} lot(s) in input")

    pfic_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}

    # Compute shared data
    part1 = compute_part1(df_lot, df_eoy, tax_year, xlsx)

    # Page 0 (page 1 of the form) – fixed fields
    page0_fields: dict = {}
    page0_fields.update(_personal_info_fields(data_dict))
    page0_fields.update(_pfic_info_fields(df_pfic))
    page0_fields.update(_part1_fields(part1))

    # Page 1 (page 2 of the form) – one set of Part IV fields per lot
    # The form only has one page 2; for multiple lots we need multiple copies.
    # We build a list of per-lot field dicts and then assemble a multi-page PDF.
    lot_pages: list[dict] = []
    actual_lots = 0

    for lot in range(number_of_lots):
        logging.info(f"  🔄 Processing lot {lot + 1}/{number_of_lots}")
        lot_result = compute_lot(df_lot, df_eoy, lot, tax_year, xlsx)
        if lot_result.skipped:
            logging.info(f"    ⏭️ Skipping lot {lot + 1} (sale in different year)")
            continue
        fields = _lot_fields(lot_result)
        lot_pages.append(fields)
        actual_lots += 1
        pfic_summary["ordinary_gains"] += lot_result.lot_summary["ordinary_gains"]
        pfic_summary["ordinary_losses"] += lot_result.lot_summary["ordinary_losses"]
        pfic_summary["capital_losses"] += lot_result.lot_summary["capital_losses"]

    # Build a PDF with page 1 followed by one copy of page 2 per lot
    template_path = os.path.join(SCRIPT_DIR, "f8621.pdf")
    _assemble_and_fill(
        template_path=template_path,
        output_path=output_path,
        page0_fields=page0_fields,
        lot_pages=lot_pages,
    )

    return actual_lots, pfic_summary


def _assemble_and_fill(
    template_path: str,
    output_path: str,
    page0_fields: dict,
    lot_pages: list[dict],
) -> None:
    """
    Assemble a final PDF:
      - page 1 of the template (filled with page0_fields)
      - one copy of template page 2 per lot (each filled with the corresponding lot fields)
    """
    template = pikepdf.Pdf.open(template_path)
    out = pikepdf.Pdf.new()

    # Page 1
    out.pages.append(template.pages[0])
    _fill_page(out.pages[0], page0_fields)

    # One lot page per lot
    for lot_fields in lot_pages:
        out.pages.append(template.pages[1])
        out.acroform.fix_copied_annotations(
            pikepdf.Page(out.pages[-1]),
            pikepdf.Page(template.pages[1]),
            template.acroform,
        )
        _fill_page(out.pages[-1], lot_fields)

    out.Root.AcroForm.NeedAppearances = True
    out.save(output_path)


def _fill_page(page, fields: dict) -> None:
    """Fill AcroForm widget annotations on a single page object in-place."""
    annots = page.get("/Annots")
    if not annots:
        return
    for annot in annots:
        t = annot.get("/T")
        if t is None:
            continue
        name = str(t)
        if name not in fields:
            continue
        ft = annot.get("/FT")
        if str(ft) == "/Btn":
            # Checkboxes: set both /V (value) and /AS (appearance state).
            val = pikepdf.Name(fields[name])
            annot["/V"] = val
            annot["/AS"] = val
        else:
            # Text fields: write value and clear any stale appearance stream
            # so NeedAppearances triggers a fresh render by the viewer.
            annot["/V"] = pikepdf.String(fields[name])
            if "/AP" in annot:
                del annot["/AP"]


# ---------------------------------------------------------------------------
# Text output
# ---------------------------------------------------------------------------


def generate_text_output(path: str, data_dict: dict, xlsx: str):
    """
    Generate a plain-text summary of Form 8621 fields, suitable for
    manual entry into tax software.  Returns (number_of_lots, pfic_summary)
    matching the same contract as create_filled_pdf().
    """
    tax_year = validate_tax_year(data_dict["Tax year"])
    df_lot, df_eoy, df_pfic = load_xlsx(xlsx)

    number_of_lots = len(df_lot.index)

    pfic_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}
    lines = []

    # ── Header / personal info ──────────────────────────────────────────────
    lines.append("=" * 60)
    lines.append("FORM 8621 — Mark-to-Market Election")
    lines.append("=" * 60)
    lines.append(f"Name of shareholder  : {data_dict['Name of shareholder']}")
    lines.append(f"Identifying Number   : {data_dict['Identifying Number']}")
    lines.append(f"Address              : {data_dict['Address']}")
    lines.append(
        f"City/State/Zip       : {data_dict['City']}, {data_dict['State']}, {data_dict['Country']}, {data_dict['Postal Code']}"
    )
    lines.append(f"Tax year             : 20{data_dict['Tax year']}")
    lines.append("Type of shareholder  : Individual")
    lines.append("")

    # ── PFIC info ───────────────────────────────────────────────────────────
    lines.append("── PFIC Information ──")
    lines.append(f"PFIC Name            : {df_pfic['PFIC Name'].values[0]}")
    lines.append(f"PFIC Address         : {df_pfic['PFIC Address'].values[0]}")
    lines.append(f"PFIC Reference ID    : {df_pfic['PFIC Reference ID'].values[0]}")
    lines.append(f"Share Class          : {df_pfic['PFIC Share Class'].values[0]}")
    lines.append("")

    # ── Part I ──────────────────────────────────────────────────────────────
    lines.append("── Part I — PFIC Information ──")
    part1 = compute_part1(df_lot, df_eoy, tax_year, xlsx)
    lines.append(f"Date of Acquisition  : {part1.date_of_acq}")
    lines.append(f"Number of Shares     : {part1.unsold_shares}")
    lines.append(f"FMV (line 1f / §1296): ${part1.value_of_pfic}")
    lines.append("Part II election     : Mark-to-Market (checked)")
    lines.append("")

    # ── Part IV — one block per lot ─────────────────────────────────────────
    actual_lots = 0
    for lot in range(number_of_lots):
        lot_result = compute_lot(df_lot, df_eoy, lot, tax_year, xlsx)
        if lot_result.skipped:
            continue

        if lot_result.is_holding:
            lines.append(f"── Part IV — Lot {actual_lots + 1} (held at year-end) ──")
            lines.append(f"  10a  FMV at year-end          : ${lot_result.fmv}")
            lines.append(f"  10b  Adjusted basis            : ${lot_result.adjusted_basis}")
            lines.append(f"  10c  Gain / (Loss) [10a - 10b] : ${lot_result.gain_loss}")

            if lot_result.gain_loss < 0:
                if lot_result.ordinary_loss is not None and lot_result.ordinary_loss != 0:
                    lines.append(f"  11   Unreversed inclusions    : ${lot_result.unreversed}")
                    lines.append(f"  12   Ordinary loss             : ${lot_result.ordinary_loss}")
                else:
                    lines.append("  11   Unreversed inclusions    : $0")
                    lines.append("  12   Ordinary loss             : $0  (non-deductible)")
            else:
                lines.append("  11   Unreversed inclusions    : (n/a)")
                lines.append("  12   Ordinary loss             : (n/a)")

        else:
            lines.append(f"── Part IV — Lot {actual_lots + 1} (sold in {tax_year}) ──")
            lines.append(f"  13a  Sale proceeds              : ${lot_result.proceeds}")
            lines.append(f"  13b  Adjusted basis at sale     : ${lot_result.adjusted_basis}")
            lines.append(f"  13c  Gain / (Loss) [13a - 13b]  : ${lot_result.sale_gain_loss}")

            if lot_result.sale_gain_loss < 0:
                if lot_result.adjusted_basis > lot_result.original_basis:
                    lines.append(f"  14a  Unreversed inclusions    : ${lot_result.unreversed}")
                    lines.append(f"  14b  Ordinary loss             : ${lot_result.ordinary_loss}")
                    lines.append("  14c  Capital loss              : (n/a)")
                else:
                    lines.append("  14a  Unreversed inclusions    : $0")
                    lines.append("  14b  Ordinary loss             : $0")
                    lines.append(f"  14c  Capital loss              : ${lot_result.capital_loss}")
            else:
                lines.append("  14a  Unreversed inclusions    : (n/a)")
                lines.append("  14b  Ordinary loss             : (n/a)")
                lines.append("  14c  Capital loss              : (n/a)")

        lines.append("")
        actual_lots += 1
        pfic_summary["ordinary_gains"] += lot_result.lot_summary["ordinary_gains"]
        pfic_summary["ordinary_losses"] += lot_result.lot_summary["ordinary_losses"]
        pfic_summary["capital_losses"] += lot_result.lot_summary["capital_losses"]

    with open(path, "w") as f:
        f.write("\n".join(lines))

    return actual_lots, pfic_summary


# ---------------------------------------------------------------------------
# Supporting XLSX output (calculation worksheet)
# ---------------------------------------------------------------------------


def generate_supporting_pdf(output_path: str, data_dict: dict, xlsx_files: list):
    """
    Generate a single PDF with one page per PFIC showing the AdB/UNI roll-forward
    calculations for each lot.  Returns a list of pfic_summary dicts, one per PFIC file.
    """
    tax_year = validate_tax_year(data_dict["Tax year"])

    all_summaries = []
    sections = []

    CURRENCY_SYMBOLS = {"EUR": "€", "GBP": "£", "JPY": "¥", "CAD": "C$", "AUD": "A$", "CHF": "CHF", "USD": "$"}

    def fmt_money(v, currency=None):
        if v is None:
            return "—"
        symbol = CURRENCY_SYMBOLS.get(currency, "$") if currency else "$"
        if float(v) == int(v):
            v = int(v)
            if v < 0:
                return f"({symbol}{abs(v):,})"
            return f"{symbol}{v:,}"
        if v < 0:
            return f"({symbol}{abs(v):,.2f})"
        return f"{symbol}{v:,.2f}"

    def fmt_fx(v):
        return f"{v:.4f}" if v is not None else "—"

    CSS = """
    @page { size: letter; margin: 0.75in 0.6in; }
    body { font-family: Helvetica, Arial, sans-serif; font-size: 9.5pt; color: #222; }
    h1 { font-size: 14pt; margin: 0 0 4pt 0; }
    .subtitle { font-size: 10pt; color: #555; margin: 0 0 12pt 0; }
    h2 {
      font-size: 11pt; color: #fff; background: #4472C4;
      padding: 4pt 8pt; margin: 16pt 0 6pt 0; break-after: avoid;
    }
    .kv-table { border-collapse: collapse; margin: 0 0 6pt 0; }
    .kv-table td { padding: 2pt 8pt 2pt 0; vertical-align: top; }
    .kv-table td:first-child { font-weight: bold; white-space: nowrap; }
    h3 {
      font-size: 9.5pt; background: #D9E2F3;
      padding: 3pt 6pt; margin: 10pt 0 4pt 0; break-after: avoid;
    }
    .glossary { font-size: 8pt; color: #555; margin: 0 0 6pt 0; line-height: 1.3; }
    table.roll {
      border-collapse: collapse; width: 100%;
      margin: 0 0 6pt 0; font-size: 8.5pt;
    }
    table.roll th, table.roll td {
      border: 1px solid #999; padding: 3pt 5pt; text-align: right;
    }
    table.roll th { background: #f0f0f0; font-weight: bold; font-size: 8pt; }
    table.roll td:first-child { text-align: center; }
    table.line-items { border-collapse: collapse; margin: 0 0 6pt 0; }
    table.line-items td { padding: 2pt 8pt 2pt 0; }
    table.line-items td:first-child { font-weight: bold; white-space: nowrap; }
    table.line-items td:last-child { text-align: right; font-variant-numeric: tabular-nums; }
    .line-items tr.calc-detail td { font-size: 8pt; color: #666; font-weight: normal; padding-top: 0; }
    .summary-table { border-collapse: collapse; margin: 6pt 0; }
    .summary-table td { padding: 2pt 10pt 2pt 0; }
    .summary-table td:first-child { font-weight: bold; }
    .summary-table td:last-child { text-align: right; font-variant-numeric: tabular-nums; }
    .page-break { page-break-before: always; }
    """

    sections = []

    for xlsx in xlsx_files:
        file_name = os.path.splitext(os.path.basename(xlsx))[0]
        df_lot, df_eoy, df_pfic = load_xlsx(xlsx)
        number_of_lots = len(df_lot.index)
        pfic_name = str(df_pfic["PFIC Name"].values[0])
        currency = str(df_pfic["Currency"].values[0]).strip().upper()
        pfic_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}
        part1 = compute_part1(df_lot, df_eoy, tax_year, xlsx)

        parts = []
        parts.append(f"<h1>{html.escape(file_name.upper())} ({html.escape(pfic_name)})</h1>")
        parts.append(
            f'<p class="subtitle">Shareholder: {html.escape(data_dict["Name of shareholder"])} '
            f"&bull; Tax Year 20{html.escape(data_dict['Tax year'])} "
            f"&bull; Currency: {html.escape(currency)}</p>"
        )

        parts.append("<h2>Part I &mdash; PFIC Information</h2>")
        parts.append('<table class="kv-table">')
        parts.append(f"<tr><td>Date of Acquisition</td><td>{html.escape(str(part1.date_of_acq))}</td></tr>")
        parts.append(f"<tr><td>Unsold Shares</td><td>{html.escape(str(part1.unsold_shares))}</td></tr>")
        parts.append(f"<tr><td>FMV (&sect;1296)</td><td>{fmt_money(part1.value_of_pfic)}</td></tr>")
        parts.append("</table>")

        actual_lots = 0

        for lot in range(number_of_lots):
            lot_result = compute_lot(df_lot, df_eoy, lot, tax_year, xlsx)
            if lot_result.skipped:
                continue

            actual_lots += 1
            pfic_summary["ordinary_gains"] += lot_result.lot_summary["ordinary_gains"]
            pfic_summary["ordinary_losses"] += lot_result.lot_summary["ordinary_losses"]
            pfic_summary["capital_losses"] += lot_result.lot_summary["capital_losses"]

            status = "held at year-end" if lot_result.is_holding else f"sold in {tax_year}"
            parts.append(f"<h2>Lot {actual_lots} &mdash; {status}</h2>")

            acq_date = df_lot["Date: Acquisition"][lot]
            acq_str = str(acq_date.date()) if hasattr(acq_date, "date") else str(acq_date)
            num_shares = float(df_lot["Number of shares"][lot])
            cost_acq = df_lot["Cost: Acquisition"][lot]
            er_acq = df_lot["Exchange Rate: Acquisition"][lot]

            parts.append('<table class="kv-table">')
            parts.append(f"<tr><td>Acquisition Date</td><td>{html.escape(acq_str)}</td></tr>")
            parts.append(f"<tr><td>Shares</td><td>{num_shares:,.2f}</td></tr>")
            parts.append(f"<tr><td>Cost ({html.escape(currency)})</td><td>{fmt_money(cost_acq, currency)}</td></tr>")
            fred_series = FRED_FX_SERIES.get(currency, "")
            series_note = f" (FRED {fred_series})" if fred_series else ""
            parts.append(
                f"<tr><td>Acquisition FX Rate</td><td>{fmt_fx(er_acq)} as of {html.escape(acq_str)}{series_note}</td></tr>"
            )
            parts.append(f"<tr><td>Original Basis (USD)</td><td>{fmt_money(lot_result.original_basis)}</td></tr>")

            if not lot_result.is_holding:
                sale_date = df_lot["Date: Sale"][lot]
                sale_str = str(sale_date.date()) if hasattr(sale_date, "date") else str(sale_date)
                sale_price = df_lot["Price per share: Sale"][lot]
                sale_er = df_lot["Exchange Rate: Sale"][lot]
                parts.append(f"<tr><td>Sale Date</td><td>{html.escape(sale_str)}</td></tr>")
                parts.append(
                    f"<tr><td>Sale Price/Share ({html.escape(currency)})</td><td>{fmt_money(sale_price, currency)}</td></tr>"
                )
                parts.append(
                    f"<tr><td>Sale FX Rate</td><td>{fmt_fx(sale_er)} as of {html.escape(sale_str)}{series_note}</td></tr>"
                )

            parts.append("</table>")

            parts.append("<h3>AdB / UNI Roll-Forward</h3>")
            parts.append(
                '<p class="glossary">AdB = Adjusted Basis &bull; UNI = Unreversed Inclusions &bull; MTM = Mark-to-Market &bull; FMV = Fair Market Value</p>'
            )
            parts.append('<table class="roll">')
            parts.append("<thead><tr>")
            for hdr in [
                "Year",
                f"EOY Price ({html.escape(currency)})",
                "EOY FX Rate" + (f" ({fred_series})" if fred_series else ""),
                "FMV (USD)",
                "AdB Begin (USD)",
                "Raw MTM (USD)",
                "Allowed Loss (USD)",
                "AdB End (USD)",
                "UNI End (USD)",
            ]:
                parts.append(f"<th>{hdr}</th>")
            parts.append("</tr></thead><tbody>")

            for entry in lot_result.roll_forward:
                parts.append("<tr>")
                parts.append(f"<td>{entry['year']}</td>")
                parts.append(f"<td>{fmt_money(entry['eoy_price'], currency)}</td>")
                parts.append(f"<td>{fmt_fx(entry['eoy_fx'])}</td>")
                parts.append(f"<td>{fmt_money(entry['fmv'])}</td>")
                parts.append(f"<td>{fmt_money(entry['adb_begin'])}</td>")
                parts.append(f"<td>{fmt_money(entry['raw_mtm'])}</td>")
                parts.append(f"<td>{fmt_money(entry['allowed_loss'])}</td>")
                parts.append(f"<td>{fmt_money(entry['adb_end'])}</td>")
                parts.append(f"<td>{fmt_money(entry['uni_end'])}</td>")
                parts.append("</tr>")

            parts.append("</tbody></table>")

            parts.append("<h3>Form 8621 Line Items</h3>")
            parts.append('<table class="line-items">')

            if lot_result.is_holding:
                cd = lot_result.calc_detail
                items = [
                    ("10a &mdash; FMV at year-end", lot_result.fmv),
                    (
                        f"&emsp;{cd['shares']:,.2f} shares &times; {html.escape(currency)} {fmt_money(cd['eoy_price'], currency)}/share &times; {fmt_fx(cd['eoy_fx'])} FX as of {html.escape(cd['eoy_date'])}{series_note}",
                        None,
                    ),
                    ("10b &mdash; Adjusted basis at year-end", lot_result.adjusted_basis),
                    ("10c &mdash; Gain/(Loss)", lot_result.gain_loss),
                ]
                if lot_result.gain_loss < 0:
                    items.append(
                        (
                            "11 &mdash; Unreversed inclusions",
                            lot_result.unreversed if lot_result.unreversed is not None else 0,
                        )
                    )
                    items.append(
                        (
                            "12 &mdash; Ordinary loss",
                            lot_result.ordinary_loss if lot_result.ordinary_loss is not None else 0,
                        )
                    )
            else:
                cd = lot_result.calc_detail
                items = [
                    ("13a &mdash; Sale proceeds", lot_result.proceeds),
                    (
                        f"&emsp;{cd['shares']:,.2f} shares &times; {html.escape(currency)} {fmt_money(cd['sale_price'], currency)}/share &times; {fmt_fx(cd['sale_fx'])} FX as of {html.escape(cd['sale_date'])}{series_note}",
                        None,
                    ),
                    ("13b &mdash; Adjusted basis at sale", lot_result.adjusted_basis),
                    ("13c &mdash; Gain/(Loss)", lot_result.sale_gain_loss),
                ]
                if lot_result.sale_gain_loss is not None and lot_result.sale_gain_loss < 0:
                    if lot_result.unreversed is not None and lot_result.unreversed > 0:
                        items.append(("14a &mdash; Unreversed inclusions", lot_result.unreversed))
                        items.append(("14b &mdash; Ordinary loss", lot_result.ordinary_loss))
                    else:
                        items.append(("14a &mdash; Unreversed inclusions", 0))
                        items.append(("14b &mdash; Ordinary loss", 0))
                        items.append(("14c &mdash; Capital loss", lot_result.capital_loss))

            for label, val in items:
                if val is None:
                    parts.append(f'<tr class="calc-detail"><td>{label}</td><td></td></tr>')
                else:
                    parts.append(f"<tr><td>{label}</td><td>{fmt_money(val)}</td></tr>")
            parts.append("</table>")

        parts.append("<h2>Summary of Gains and Losses</h2>")
        parts.append('<table class="summary-table">')
        parts.append(f"<tr><td>Total Ordinary Gains</td><td>{fmt_money(pfic_summary['ordinary_gains'])}</td></tr>")
        parts.append(f"<tr><td>Total Ordinary Losses</td><td>{fmt_money(pfic_summary['ordinary_losses'])}</td></tr>")
        parts.append(f"<tr><td>Total Capital Losses</td><td>{fmt_money(pfic_summary['capital_losses'])}</td></tr>")
        parts.append("</table>")

        sections.append("\n".join(parts))
        all_summaries.append(pfic_summary)

    html_content = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{CSS}</style></head><body>
{'<div class="page-break">'.join(sections)}
</body></html>"""

    logging.getLogger("weasyprint").setLevel(logging.ERROR)
    weasyprint.HTML(string=html_content).write_pdf(output_path)

    return all_summaries


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args():
    parser = argparse.ArgumentParser(description="Fill out IRS Form 8621 for Mark-to-Market (MTM) elections.")
    parser.add_argument(
        "--format",
        choices=["pdf", "txt"],
        default="pdf",
        help="Output format: 'pdf' for filled PDF (default), 'txt' for plain-text summary.",
    )
    parser.add_argument(
        "--inputs-dir",
        default=None,
        help="Directory containing XLSX input files (default: inputs/ next to this script).",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Output directory (default: outputs/20YY/ next to this script).",
    )
    parser.add_argument(
        "--remember",
        action="store_true",
        help="Save entered details to inputs/__details.json for reuse on future runs.",
    )
    return parser.parse_args()


DETAILS_FILENAME = "__details.json"


def _details_path(inputs_dir: str) -> str:
    return os.path.join(inputs_dir, DETAILS_FILENAME)


def _load_details(inputs_dir: str) -> dict | None:
    path = _details_path(inputs_dir)
    if not os.path.isfile(path):
        return None
    try:
        with open(path) as f:
            data = json.load(f)
        logging.info(f"📋 Loaded saved details from {path}")
        return data
    except (json.JSONDecodeError, OSError) as exc:
        logging.warning(f"⚠️  Could not read {path}: {exc}")
        return None


def _save_details(data_dict: dict, inputs_dir: str) -> None:
    path = _details_path(inputs_dir)
    os.makedirs(inputs_dir, exist_ok=True)
    with open(path, "w") as f:
        json.dump(data_dict, f, indent=2)
    logging.info(f"💾 Details saved to {path}")


def read_inputs(args):
    inputs_dir = args.inputs_dir or os.path.join(SCRIPT_DIR, "inputs")
    saved = _load_details(inputs_dir)

    if saved is not None:
        data_dict = saved
        logging.info("📋 Using saved details (skipping input prompts)")
    else:
        data_dict = {}

        logging.info("📝 Enter shareholder details:")

        data_dict["Name of shareholder"] = input("👤 Name of shareholder: ")
        data_dict["Identifying Number"] = getpass.getpass("🆔 Identifying Number (e.g., SSN)")
        data_dict["Address"] = input("🏠 Address (Street + House Number): ")
        address_line_2 = input("🏠 Address line 2 (or press Enter to skip): ").strip()
        data_dict["Address line 2"] = address_line_2 if address_line_2 else ""
        data_dict["City"] = input("🏙️ City: ")
        data_dict["State"] = input("🗺️ State/Province: ")
        data_dict["Country"] = input("🌍 Country: ")
        data_dict["Postal Code"] = input("📮 Postal Code: ")

        data_dict["Tax year"] = input("📅 Tax year (last two digits): ")

    # Validate tax year early
    validate_tax_year(data_dict["Tax year"])

    # Determine format
    if args.format == "txt":
        data_dict["output_format"] = "txt"
    else:
        data_dict["output_format"] = "pdf"

    # Persist if --remember was requested and details weren't already saved
    if args.remember and saved is None:
        _save_details(data_dict, inputs_dir)

    files = glob.glob(os.path.join(inputs_dir, "*.xlsx"))

    # Filter out backup/temp files
    files = [f for f in files if not os.path.basename(f).startswith("~") and not f.endswith(".bak")]

    return data_dict, files


def main():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    logging.getLogger("weasyprint").setLevel(logging.ERROR)
    logging.getLogger("fontTools.subset").setLevel(logging.WARNING)
    logging.getLogger("fontTools.ttLib.ttFont").setLevel(logging.WARNING)
    logging.info("🚀 Form 8621 Filler Initialized")

    args = parse_args()

    try:
        data_dict, files = read_inputs(args)
        if not files:
            logging.error(
                "💥 No input files found in the inputs directory. Please consult the README for instructions."
            )
            sys.exit(1)

        # Build output path (prevents path traversal from tax year)
        tax_year_raw = data_dict["Tax year"].strip()
        if not tax_year_raw.isdigit():
            logging.error("💥 Invalid tax year.")
            sys.exit(1)
        output_dir = args.output_dir or os.path.join(SCRIPT_DIR, "outputs", f"20{tax_year_raw}")
        os.makedirs(output_dir, exist_ok=True)
        logging.info(f"📁 Output directory: {output_dir}")

        total_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}

        for file in files:
            file_name = os.path.splitext(os.path.basename(file))[0]
            logging.info(f"📂 Processing PFIC: {file_name}")

            if data_dict["output_format"] == "txt":
                form_output_path = os.path.join(output_dir, f"{file_name}.txt")
                number_of_lots, pfic_summary = generate_text_output(
                    path=form_output_path, data_dict=data_dict, xlsx=file
                )
            else:
                form_output_path = os.path.join(output_dir, f"{file_name}.pdf")
                number_of_lots, pfic_summary = create_filled_pdf(
                    output_path=form_output_path, data_dict=data_dict, xlsx=file
                )

            total_summary["ordinary_gains"] += pfic_summary["ordinary_gains"]
            total_summary["ordinary_losses"] += pfic_summary["ordinary_losses"]
            total_summary["capital_losses"] += pfic_summary["capital_losses"]

            logging.info(f"  ✅ Form completed and saved to {form_output_path}")

        supporting_path = os.path.join(output_dir, f"20{data_dict['Tax year']}_supporting.pdf")
        generate_supporting_pdf(output_path=supporting_path, data_dict=data_dict, xlsx_files=files)
        logging.info(f"📊 Supporting document saved to {supporting_path}")

        logging.info("✅ All forms processed successfully!")

        logging.info("")
        logging.info("=" * 60)
        logging.info(f"📋 SUMMARY OF GAINS AND LOSSES FOR TAX YEAR 20{data_dict['Tax year']}")
        logging.info("=" * 60)

        if total_summary["ordinary_gains"] > 0:
            logging.info(f"💰 Total Ordinary Gains: ${total_summary['ordinary_gains']:.2f}")
            logging.info("   ➡️  Add this amount to your ordinary income on your tax return")
            logging.info("")
        if total_summary["ordinary_losses"] > 0:
            logging.info(f"📉 Total Ordinary Losses: ${total_summary['ordinary_losses']:.2f}")
            logging.info("   ➡️  Include this amount as an ordinary loss on your tax return")
            logging.info("")
        if total_summary["capital_losses"] > 0:
            logging.info(f"📉 Total Capital Losses: ${total_summary['capital_losses']:.2f}")
            logging.info("   ➡️  Report according to capital loss rules in the Code and regulations")
            logging.info("")

        if all(v == 0 for v in total_summary.values()):
            logging.info("📊 No gains or losses to report this year")
            logging.info("")

    except SystemExit:
        raise
    except Exception:
        logging.exception("💥 An error occurred")
    finally:
        logging.info("👋 Shutting down. Goodbye!")


if __name__ == "__main__":
    main()
