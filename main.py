import argparse
import getpass
import glob
import json
import logging
import os
import re
import sys

import numpy as np
import pandas as pd
import pikepdf

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
LOT_COLUMNS = [
    "Date: Acquisition",
    "Price per share: Acquisition",
    "Number of shares",
    "Cost: Acquisition",
    "Exchange Rate: Acquisition",
    "Date: Sale",
    "Price per share: Sale",
    "Exchange Rate: Sale",
]
EOY_COLUMNS = ["Year", "Price", "Exchange Rate"]
PFIC_COLUMNS = ["PFIC Name", "PFIC Address", "PFIC Reference ID", "PFIC Share Class"]

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


def validate_xlsx_columns(
    df: pd.DataFrame, expected: list, sheet_name: str, filepath: str
):
    """Validate that a DataFrame has all expected columns. Exits with error if missing."""
    missing = [col for col in expected if col not in df.columns]
    if missing:
        logging.error(
            f"💥 Missing columns in sheet '{sheet_name}' of {filepath}: "
            f"{', '.join(missing)}"
        )
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


def validate_sale_data_completeness(df_lot: pd.DataFrame, filepath: str):
    """Ensure that rows with a sale price also have a sale exchange rate and date."""
    for i in range(len(df_lot.index)):
        sale_price = df_lot["Price per share: Sale"][i]
        sale_er = df_lot["Exchange Rate: Sale"][i]
        sale_date = df_lot["Date: Sale"][i]
        has_price = not pd.isna(sale_price)
        has_er = not pd.isna(sale_er)
        has_date = not pd.isna(sale_date)

        if has_price and (not has_er or not has_date):
            missing = []
            if not has_er:
                missing.append("Exchange Rate: Sale")
            if not has_date:
                missing.append("Date: Sale")
            logging.error(
                f"💥 Lot {i + 1} in {filepath} has a sale price but is missing: "
                f"{', '.join(missing)}"
            )
            sys.exit(1)
        if (not has_price) and (has_er or has_date):
            logging.warning(
                f"⚠️  Lot {i + 1} in {filepath} has sale details but no sale price — "
                "sale data will be ignored."
            )


# ---------------------------------------------------------------------------
# Computation logic (shared between PDF and text output)
# ---------------------------------------------------------------------------


def compute_part1(
    df_lot: pd.DataFrame, df_eoy: pd.DataFrame, current_year: int, filepath: str
) -> Part1Result:
    """Compute Part I results (shared logic for both PDF and text)."""
    date_of_acq = (
        pd.to_datetime(df_lot["Date: Acquisition"].values[0]).strftime("%Y-%m-%d")
        if len(df_lot.index) == 1
        else "Multiple"
    )
    unsold_shares = 0
    for lot in range(len(df_lot.index)):
        if np.isnan(df_lot["Price per share: Sale"][lot]):
            unsold_shares += df_lot["Number of shares"][lot]

    last_er = get_eoy_value(df_eoy, current_year, "Exchange Rate", filepath)
    last_price = get_eoy_value(df_eoy, current_year, "Price", filepath)
    value_of_pfic = round(unsold_shares * last_price * last_er)

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
    num_shares = df_lot["Number of shares"][lot]
    original_basis = cost_acquisition * er_of_acquisition

    aab = original_basis
    uni = 0.0
    for year in range(year_of_acquisition, current_year):
        price = get_eoy_value(df_eoy, year, "Price", filepath)
        fx = get_eoy_value(df_eoy, year, "Exchange Rate", filepath)
        fmv = round(num_shares * price * fx)
        raw_mtm = fmv - aab
        if raw_mtm >= 0:
            aab = aab + raw_mtm
            uni = uni + raw_mtm
        else:
            allowed_loss = min(-raw_mtm, uni)
            aab = aab - allowed_loss
            uni = uni - allowed_loss
    adjusted_basis = round(aab)
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
            )

        proceeds = round(num_shares * sale_price * sale_er)
        sale_gain_loss = proceeds - adjusted_basis

        if sale_gain_loss < 0:
            if unreversed_amount > 0:
                unreversed = unreversed_amount
                ordinary_loss = -min(unreversed, -sale_gain_loss)
                capital_loss = None
                logging.info(
                    f"    📉 Lot {lot + 1}: Ordinary loss of ${abs(ordinary_loss)}"
                )
            else:
                unreversed = None
                ordinary_loss = 0
                capital_loss = sale_gain_loss
                logging.info(
                    f"    📉 Lot {lot + 1}: Capital loss of ${abs(sale_gain_loss)}"
                )
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
        )

    else:
        # Holding at year-end
        last_er = get_eoy_value(df_eoy, current_year, "Exchange Rate", filepath)
        last_price = get_eoy_value(df_eoy, current_year, "Price", filepath)
        fmv = round(num_shares * last_price * last_er)

        gain_loss = fmv - adjusted_basis
        logging.info(f"    📈 Lot {lot + 1}: No sale (holding position)")

        if gain_loss < 0:
            if unreversed_amount > 0:
                unreversed = unreversed_amount
                ordinary_loss = -min(unreversed, -gain_loss)
                logging.info(
                    f"    📉 Lot {lot + 1}: Ordinary loss of ${abs(ordinary_loss)}"
                )
            else:
                unreversed = 0
                ordinary_loss = 0
                logging.info(
                    f"    📉 Lot {lot + 1}: Unrecognizable loss of ${abs(gain_loss)}"
                )
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

        fields.update(
            {F_13A: "", F_13B: "", F_13C: "", F_14A: "", F_14B: "", F_14C: ""}
        )

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
    """Load and validate an XLSX input file. Returns (df_lot, df_eoy, df_pfic)."""
    logging.info(f"  📂 Reading {xlsx_path}")
    df_lot = pd.read_excel(xlsx_path, sheet_name="Lot Details")
    df_eoy = pd.read_excel(xlsx_path, sheet_name="EOY Details")
    df_pfic = pd.read_excel(xlsx_path, sheet_name="PFIC Details")

    validate_xlsx_columns(df_lot, LOT_COLUMNS, "Lot Details", xlsx_path)
    validate_xlsx_columns(df_eoy, EOY_COLUMNS, "EOY Details", xlsx_path)
    validate_xlsx_columns(df_pfic, PFIC_COLUMNS, "PFIC Details", xlsx_path)

    validate_reference_id(str(df_pfic["PFIC Reference ID"].values[0]))
    validate_sale_data_completeness(df_lot, xlsx_path)

    return df_lot, df_eoy, df_pfic


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
    lines.append(f"Type of shareholder  : Individual")
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
    lines.append(f"Part II election     : Mark-to-Market (checked)")
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
                    lines.append(f"  11   Unreversed inclusions    : $0")
                    lines.append(
                        f"  12   Ordinary loss             : $0  (non-deductible)"
                    )
            else:
                lines.append(f"  11   Unreversed inclusions    : (n/a)")
                lines.append(f"  12   Ordinary loss             : (n/a)")

        else:
            lines.append(f"── Part IV — Lot {actual_lots + 1} (sold in {tax_year}) ──")
            lines.append(f"  13a  Sale proceeds              : ${lot_result.proceeds}")
            lines.append(f"  13b  Adjusted basis at sale     : ${lot_result.adjusted_basis}")
            lines.append(f"  13c  Gain / (Loss) [13a - 13b]  : ${lot_result.sale_gain_loss}")

            if lot_result.sale_gain_loss < 0:
                if lot_result.adjusted_basis > lot_result.original_basis:
                    lines.append(f"  14a  Unreversed inclusions    : ${lot_result.unreversed}")
                    lines.append(f"  14b  Ordinary loss             : ${lot_result.ordinary_loss}")
                    lines.append(f"  14c  Capital loss              : (n/a)")
                else:
                    lines.append(f"  14a  Unreversed inclusions    : $0")
                    lines.append(f"  14b  Ordinary loss             : $0")
                    lines.append(f"  14c  Capital loss              : ${lot_result.capital_loss}")
            else:
                lines.append(f"  14a  Unreversed inclusions    : (n/a)")
                lines.append(f"  14b  Ordinary loss             : (n/a)")
                lines.append(f"  14c  Capital loss              : (n/a)")

        lines.append("")
        actual_lots += 1
        pfic_summary["ordinary_gains"] += lot_result.lot_summary["ordinary_gains"]
        pfic_summary["ordinary_losses"] += lot_result.lot_summary["ordinary_losses"]
        pfic_summary["capital_losses"] += lot_result.lot_summary["capital_losses"]

    with open(path, "w") as f:
        f.write("\n".join(lines))

    return actual_lots, pfic_summary


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args():
    parser = argparse.ArgumentParser(
        description="Fill out IRS Form 8621 for Mark-to-Market (MTM) elections."
    )
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
        data_dict["Identifying Number"] = getpass.getpass(
            "🆔 Identifying Number (e.g., SSN)"
        )
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
    logging.info("🚀 Form 8621 Filler Initialized")

    args = parse_args()

    try:
        data_dict, files = read_inputs(args)
        if not files:
            logging.error(
                "💥 No input files found in the inputs directory. "
                "Please consult the README for instructions."
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

        logging.info("✅ All forms processed successfully!")

        logging.info("")
        logging.info("=" * 60)
        logging.info(
            f"📋 SUMMARY OF GAINS AND LOSSES FOR TAX YEAR 20{data_dict['Tax year']}"
        )
        logging.info("=" * 60)

        if total_summary["ordinary_gains"] > 0:
            logging.info(
                f"💰 Total Ordinary Gains: ${total_summary['ordinary_gains']:.2f}"
            )
            logging.info(
                "   ➡️  Add this amount to your ordinary income on your tax return"
            )
            logging.info("")
        if total_summary["ordinary_losses"] > 0:
            logging.info(
                f"📉 Total Ordinary Losses: ${total_summary['ordinary_losses']:.2f}"
            )
            logging.info(
                "   ➡️  Include this amount as an ordinary loss on your tax return"
            )
            logging.info("")
        if total_summary["capital_losses"] > 0:
            logging.info(
                f"📉 Total Capital Losses: ${total_summary['capital_losses']:.2f}"
            )
            logging.info(
                "   ➡️  Report according to capital loss rules in the Code and regulations"
            )
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
