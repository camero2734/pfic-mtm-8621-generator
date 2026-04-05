import pdfrw
import pandas as pd
import numpy as np
import glob
import os
import logging

# Value written into a checkbox field to mark it checked
CHECKED = pdfrw.PdfObject("/1")


# ---------------------------------------------------------------------------
# AcroForm field name constants (as they appear in f8621.pdf widget /T keys)
# ---------------------------------------------------------------------------

# Page 1 – personal info
F_NAME = "f1_1[0]"  # Name of shareholder
F_ADDRESS_1 = "f1_2[0]"  # Street address line 1
F_ADDRESS_2 = "f1_3[0]"  # Street address line 2 (city/state/zip overflow)
F_CITY_STATE_ZIP = "f1_4[0]"  # City, state, zip
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
# PDF helpers
# ---------------------------------------------------------------------------


def _fill_fields(pdf: pdfrw.PdfReader, fields: dict[str, str], page_index: int) -> None:
    """Write *fields* (field_name -> value) into widget annotations on *page_index*."""
    page = pdf.pages[page_index]
    annotations = page.Annots
    if not annotations:
        return
    for annot in annotations:
        if annot.Subtype != "/Widget":
            continue
        raw_t = annot.T
        if not raw_t:
            continue
        # Decode UTF-16-BE field name
        s = str(raw_t).strip("<>")
        try:
            name = bytes.fromhex(s).decode("utf-16-be").lstrip("\ufeff")
        except Exception:
            name = str(raw_t)
        if name in fields:
            annot.update(pdfrw.PdfDict(V=fields[name], AP=""))


def fill_pdf(
    input_path: str, output_path: str, page_fields: list[dict[str, str]]
) -> None:
    """
    Fill *input_path* and write to *output_path*.

    *page_fields* is a list of dicts, one per PDF page (0-indexed).
    Each dict maps AcroForm field names to string values (or PdfObjects for checkboxes).
    """
    pdf = pdfrw.PdfReader(input_path)
    for page_index, fields in enumerate(page_fields):
        if fields:
            _fill_fields(pdf, fields, page_index)
    pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject("true")))
    pdfrw.PdfWriter().write(output_path, pdf)


# ---------------------------------------------------------------------------
# Form section builders  (return dicts of field_name -> value)
# ---------------------------------------------------------------------------


def _personal_info_fields(data_dict: dict) -> dict:
    return {
        F_NAME: data_dict["Name of shareholder"],
        F_ADDRESS_1: data_dict["Address"],
        F_CITY_STATE_ZIP: data_dict["City, State, Zip, Country"],
        F_IDENTIFYING_NUM: data_dict["Identifying Number"],
        F_TAX_YEAR: data_dict["Tax year"],
        C_SHAREHOLDER_INDIVIDUAL: CHECKED,  # always individual
    }


def _pfic_info_fields(df_pfic: pd.DataFrame) -> dict:
    return {
        F_PFIC_NAME: str(df_pfic["PFIC Name"].values[0]),
        F_PFIC_ADDRESS: str(df_pfic["PFIC Address"].values[0]),
        F_PFIC_REF_ID: str(df_pfic["PFIC Reference ID"].values[0]),
        F_PFIC_SHARE_CLASS: str(df_pfic["PFIC Share Class"].values[0]),
    }


def _part1_fields(
    df_lot: pd.DataFrame, df_eoy: pd.DataFrame, current_year: int
) -> dict:
    date_of_acq = (
        pd.to_datetime(df_lot["Date: Acquisition"].values[0]).strftime("%Y-%m-%d")
        if len(df_lot.index) == 1
        else "Multiple"
    )
    unsold_shares = 0
    for lot in range(len(df_lot.index)):
        if np.isnan(df_lot["Price per share: Sale"][lot]):
            unsold_shares += df_lot["Number of shares"][lot]

    last_er = df_eoy[df_eoy["Year"] == current_year]["Exchange Rate"].values[0]
    last_price = df_eoy[df_eoy["Year"] == current_year]["Price"].values[0]
    value_of_pfic = round(unsold_shares * last_price / last_er)

    fields = {
        F_DATE_ACQUISITION: str(date_of_acq),
        F_NUM_SHARES: str(unsold_shares),
        F_AMOUNT_1291: "",
        F_AMOUNT_1293: "",
        F_AMOUNT_1296: str(value_of_pfic),
        C_SECTION_TYPE_C: CHECKED,
        C_PART2_MTM: CHECKED,
    }

    # Value-of-PFIC checkboxes
    if 0 <= value_of_pfic <= 50_000:
        fields[C_VALUE_LE_50K] = CHECKED
    elif value_of_pfic <= 100_000:
        fields[C_VALUE_50_100K] = CHECKED
    elif value_of_pfic <= 150_000:
        fields[C_VALUE_100_150K] = CHECKED
    elif value_of_pfic <= 200_000:
        fields[C_VALUE_150_200K] = CHECKED
    else:
        fields[F_VALUE_OVER_200K] = str(value_of_pfic)

    return fields


def _part4_fields(
    df_lot: pd.DataFrame, df_eoy: pd.DataFrame, lot: int, current_year: int
):
    """
    Returns (fields_dict, lot_summary) for one lot, or (None, lot_summary) if the lot
    should be skipped (sold in a prior year).
    """
    lot_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}

    year_of_acquisition = df_lot["Date: Acquisition"][lot].year
    cost_acquisition = df_lot["Cost: Acquisition"][lot]
    er_of_acquisition = df_lot["Exchange Rate: Acquisition"][lot]
    num_shares = df_lot["Number of shares"][lot]
    original_basis = cost_acquisition / er_of_acquisition

    if current_year > year_of_acquisition:
        prev_er = df_eoy[df_eoy["Year"] == current_year - 1]["Exchange Rate"].values[0]
        prev_price = df_eoy[df_eoy["Year"] == current_year - 1]["Price"].values[0]
        adjusted_basis = round(num_shares * prev_price / prev_er)
    else:
        adjusted_basis = round(original_basis)

    fields = {}

    if np.isnan(df_lot["Price per share: Sale"][lot]):
        # Holding at year-end
        logging.info(f"    📈 Lot {lot + 1}: No sale (holding position)")
        last_er = df_eoy[df_eoy["Year"] == current_year]["Exchange Rate"].values[0]
        last_price = df_eoy[df_eoy["Year"] == current_year]["Price"].values[0]
        fmv = round(num_shares * last_price / last_er)

        fields[F_10A] = str(fmv)
        fields[F_10B] = str(adjusted_basis)
        fields[F_10C] = str(fmv - adjusted_basis)

        gain_loss = fmv - adjusted_basis
        if gain_loss < 0:
            if adjusted_basis > original_basis:
                unreversed = round(adjusted_basis - original_basis)
                ordinary_loss = -min(unreversed, -gain_loss)
                fields[F_11] = str(unreversed)
                fields[F_12] = str(ordinary_loss)
                logging.info(
                    f"    📉 Lot {lot + 1}: Ordinary loss of ${abs(ordinary_loss)}"
                )
                lot_summary["ordinary_losses"] += abs(ordinary_loss)
            else:
                logging.info(
                    f"    📉 Lot {lot + 1}: Unrecognizable loss of ${abs(gain_loss)}"
                )
                fields[F_11] = "0"
                fields[F_12] = "0"
        else:
            fields[F_11] = ""
            fields[F_12] = ""
            logging.info(f"    📈 Lot {lot + 1}: Ordinary gain of ${gain_loss}")
            lot_summary["ordinary_gains"] += gain_loss

        fields.update(
            {F_13A: "", F_13B: "", F_13C: "", F_14A: "", F_14B: "", F_14C: ""}
        )

    else:
        # Sold lot
        logging.info(f"    💰 Lot {lot + 1}: Sale detected")
        last_er = df_lot["Exchange Rate: Sale"][lot]
        last_price = df_lot["Price per share: Sale"][lot]
        year_of_sale = df_lot["Date: Sale"][lot].year
        if year_of_sale < current_year:
            return None, lot_summary
        proceeds = round(num_shares * last_price / last_er)
        sale_gain_loss = proceeds - adjusted_basis

        fields[F_13A] = str(proceeds)
        fields[F_13B] = str(adjusted_basis)
        fields[F_13C] = str(sale_gain_loss)

        if sale_gain_loss < 0:
            if adjusted_basis > original_basis:
                unreversed = round(adjusted_basis - original_basis)
                ordinary_loss = -min(unreversed, -sale_gain_loss)
                fields[F_14A] = str(unreversed)
                fields[F_14B] = str(ordinary_loss)
                fields[F_14C] = ""
                logging.info(
                    f"    📉 Lot {lot + 1}: Ordinary loss of ${abs(ordinary_loss)}"
                )
                lot_summary["ordinary_losses"] += abs(ordinary_loss)
            else:
                capital_loss = sale_gain_loss
                fields[F_14A] = "0"
                fields[F_14B] = "0"
                fields[F_14C] = str(capital_loss)
                logging.info(
                    f"    📉 Lot {lot + 1}: Capital loss of ${abs(capital_loss)}"
                )
                lot_summary["capital_losses"] += abs(capital_loss)
        else:
            fields.update({F_14A: "", F_14B: "", F_14C: ""})
            logging.info(f"    📈 Lot {lot + 1}: Ordinary gain of ${sale_gain_loss}")
            lot_summary["ordinary_gains"] += sale_gain_loss

        fields.update({F_10A: "", F_10B: "", F_10C: "", F_11: "", F_12: ""})

    return fields, lot_summary


# ---------------------------------------------------------------------------
# Main PDF generation
# ---------------------------------------------------------------------------


def create_filled_pdf(output_path: str, data_dict: dict, xlsx: str):
    """
    Build and save a filled Form 8621 PDF directly via AcroForm fields.
    Returns (number_of_lots, pfic_summary).
    """
    tax_year = 2000 + int(data_dict["Tax year"])
    df_lot = pd.read_excel(xlsx, sheet_name="Lot Details")
    df_eoy = pd.read_excel(xlsx, sheet_name="EOY Details")
    df_pfic = pd.read_excel(xlsx, sheet_name="PFIC Details")
    number_of_lots = len(df_lot.index)
    logging.info(f"  📊 Found {number_of_lots} lots to process")
    logging.debug(f"  📊 Lot details dataframe:\n{df_lot}")

    print(df_pfic)

    pfic_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}

    # Page 0 (page 1 of the form) – fixed fields
    page0_fields: dict = {}
    page0_fields.update(_personal_info_fields(data_dict))
    page0_fields.update(_pfic_info_fields(df_pfic))
    page0_fields.update(_part1_fields(df_lot, df_eoy, tax_year))

    # Page 1 (page 2 of the form) – one set of Part IV fields per lot
    # The form only has one page 2; for multiple lots we need multiple copies.
    # We build a list of per-lot field dicts and then assemble a multi-page PDF.
    lot_pages: list[dict] = []
    actual_lots = 0

    for lot in range(number_of_lots):
        logging.info(f"  🔄 Processing lot {lot + 1}/{number_of_lots}")
        fields, lot_summary = _part4_fields(df_lot, df_eoy, lot, tax_year)
        if fields is None:
            logging.info(f"    ⏭️ Skipping lot {lot + 1} (sale in different year)")
            continue
        lot_pages.append(fields)
        actual_lots += 1
        pfic_summary["ordinary_gains"] += lot_summary["ordinary_gains"]
        pfic_summary["ordinary_losses"] += lot_summary["ordinary_losses"]
        pfic_summary["capital_losses"] += lot_summary["capital_losses"]

    # Build a PDF with page 1 followed by one copy of page 2 per lot
    _assemble_and_fill(
        template_path="f8621.pdf",
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

    We modify the page tree (/Kids, /Count) directly so that the AcroForm dictionary
    is preserved intact when pdfrw writes the file.
    """
    template = pdfrw.PdfReader(template_path)

    # Fill page 1 in-place
    _fill_page(template.pages[0], page0_fields)

    # Fill the template's own page 2 with the first lot
    if lot_pages:
        _fill_page(template.pages[1], lot_pages[0])

    # For additional lots, re-read the template to get fresh independent page copies
    extra_pages = []
    for lot_fields in lot_pages[1:]:
        fresh = pdfrw.PdfReader(template_path)
        page = fresh.pages[1]
        _fill_page(page, lot_fields)
        extra_pages.append(page)

    # Rebuild the page tree: page 1 + one page 2 per lot (drop template pages 3 & 4)
    new_kids = (
        [template.pages[0]] + ([template.pages[1]] if lot_pages else []) + extra_pages
    )
    template.Root.Pages.Kids = pdfrw.PdfArray(new_kids)
    template.Root.Pages.Count = pdfrw.PdfObject(len(new_kids))

    template.Root.AcroForm.update(
        pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject("true"))
    )
    pdfrw.PdfWriter().write(output_path, template)


def _fill_page(page, fields: dict) -> None:
    """Fill AcroForm widget annotations on a single page object in-place."""
    annotations = page.Annots
    if not annotations:
        return
    for annot in annotations:
        if annot.Subtype != "/Widget":
            continue
        raw_t = annot.T
        if not raw_t:
            continue
        s = str(raw_t).strip("<>")
        try:
            name = bytes.fromhex(s).decode("utf-16-be").lstrip("\ufeff")
        except Exception:
            name = str(raw_t)
        if name in fields:
            annot.update(pdfrw.PdfDict(V=fields[name], AP=""))


# ---------------------------------------------------------------------------
# Text output (unchanged logic, kept for --txt mode)
# ---------------------------------------------------------------------------


def generate_text_output(path: str, data_dict: dict, xlsx: str):
    """
    Generate a plain-text summary of Form 8621 fields, suitable for
    manual entry into tax software.  Returns (number_of_lots, pfic_summary)
    matching the same contract as create_filled_pdf().
    """
    tax_year = 2000 + int(data_dict["Tax year"])
    df_lot = pd.read_excel(xlsx, sheet_name="Lot Details")
    df_eoy = pd.read_excel(xlsx, sheet_name="EOY Details")
    df_pfic = pd.read_excel(xlsx, sheet_name="PFIC Details")
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
    lines.append(f"City/State/Zip       : {data_dict['City, State, Zip, Country']}")
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
    total_unsold_shares = 0
    date_of_acq = (
        pd.to_datetime(df_lot["Date: Acquisition"].values[0]).strftime("%Y-%m-%d")
        if len(df_lot.index) == 1
        else "Multiple"
    )
    for lot in range(len(df_lot.index)):
        if np.isnan(df_lot["Price per share: Sale"][lot]):
            total_unsold_shares += df_lot["Number of shares"][lot]

    last_er = df_eoy[df_eoy["Year"] == tax_year]["Exchange Rate"].values[0]
    last_price = df_eoy[df_eoy["Year"] == tax_year]["Price"].values[0]
    fmv_total = round(total_unsold_shares * last_price / last_er)

    lines.append(f"Date of Acquisition  : {date_of_acq}")
    lines.append(f"Number of Shares     : {total_unsold_shares}")
    lines.append(f"FMV (line 1f / §1296): ${fmv_total}")
    lines.append(f"Part II election     : Mark-to-Market (checked)")
    lines.append("")

    # ── Part IV — one block per lot ─────────────────────────────────────────
    actual_lots = 0
    for lot in range(number_of_lots):
        year_of_acquisition = df_lot["Date: Acquisition"][lot].year
        cost_acquisition = df_lot["Cost: Acquisition"][lot]
        er_of_acquisition = df_lot["Exchange Rate: Acquisition"][lot]
        num_shares = df_lot["Number of shares"][lot]
        original_basis = cost_acquisition / er_of_acquisition

        if tax_year > year_of_acquisition:
            prev_er = df_eoy[df_eoy["Year"] == tax_year - 1]["Exchange Rate"].values[0]
            prev_price = df_eoy[df_eoy["Year"] == tax_year - 1]["Price"].values[0]
            adjusted_basis = round(num_shares * prev_price / prev_er)
        else:
            adjusted_basis = round(original_basis)

        lot_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}

        if np.isnan(df_lot["Price per share: Sale"][lot]):
            lot_er = df_eoy[df_eoy["Year"] == tax_year]["Exchange Rate"].values[0]
            lot_price = df_eoy[df_eoy["Year"] == tax_year]["Price"].values[0]
            fmv = round(num_shares * lot_price / lot_er)
            gain_loss = fmv - adjusted_basis

            lines.append(f"── Part IV — Lot {actual_lots + 1} (held at year-end) ──")
            lines.append(f"  10a  FMV at year-end          : ${fmv}")
            lines.append(f"  10b  Adjusted basis            : ${adjusted_basis}")
            lines.append(f"  10c  Gain / (Loss) [10a - 10b] : ${gain_loss}")

            if gain_loss < 0:
                if adjusted_basis > original_basis:
                    unreversed = round(adjusted_basis - original_basis)
                    ordinary_loss = -min(unreversed, -gain_loss)
                    lines.append(f"  11   Unreversed inclusions    : ${unreversed}")
                    lines.append(f"  12   Ordinary loss             : ${ordinary_loss}")
                    lot_summary["ordinary_losses"] += abs(ordinary_loss)
                else:
                    lines.append(f"  11   Unreversed inclusions    : $0")
                    lines.append(
                        f"  12   Ordinary loss             : $0  (non-deductible)"
                    )
            else:
                lines.append(f"  11   Unreversed inclusions    : (n/a)")
                lines.append(f"  12   Ordinary loss             : (n/a)")
                lot_summary["ordinary_gains"] += gain_loss

        else:
            sale_er = df_lot["Exchange Rate: Sale"][lot]
            sale_price = df_lot["Price per share: Sale"][lot]
            year_of_sale = df_lot["Date: Sale"][lot].year
            if year_of_sale < tax_year:
                number_of_lots -= 1
                continue
            proceeds = round(num_shares * sale_price / sale_er)
            sale_gain_loss = proceeds - adjusted_basis

            lines.append(f"── Part IV — Lot {actual_lots + 1} (sold in {tax_year}) ──")
            lines.append(f"  13a  Sale proceeds              : ${proceeds}")
            lines.append(f"  13b  Adjusted basis at sale     : ${adjusted_basis}")
            lines.append(f"  13c  Gain / (Loss) [13a - 13b]  : ${sale_gain_loss}")

            if sale_gain_loss < 0:
                if adjusted_basis > original_basis:
                    unreversed = round(adjusted_basis - original_basis)
                    ordinary_loss = -min(unreversed, -sale_gain_loss)
                    lines.append(f"  14a  Unreversed inclusions    : ${unreversed}")
                    lines.append(f"  14b  Ordinary loss             : ${ordinary_loss}")
                    lines.append(f"  14c  Capital loss              : (n/a)")
                    lot_summary["ordinary_losses"] += abs(ordinary_loss)
                else:
                    capital_loss = sale_gain_loss
                    lines.append(f"  14a  Unreversed inclusions    : $0")
                    lines.append(f"  14b  Ordinary loss             : $0")
                    lines.append(f"  14c  Capital loss              : ${capital_loss}")
                    lot_summary["capital_losses"] += abs(capital_loss)
            else:
                lines.append(f"  14a  Unreversed inclusions    : (n/a)")
                lines.append(f"  14b  Ordinary loss             : (n/a)")
                lines.append(f"  14c  Capital loss              : (n/a)")
                lot_summary["ordinary_gains"] += sale_gain_loss

        lines.append("")
        actual_lots += 1
        pfic_summary["ordinary_gains"] += lot_summary["ordinary_gains"]
        pfic_summary["ordinary_losses"] += lot_summary["ordinary_losses"]
        pfic_summary["capital_losses"] += lot_summary["capital_losses"]

    number_of_lots = actual_lots

    with open(path, "w") as f:
        f.write("\n".join(lines))

    return number_of_lots, pfic_summary


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def read_inputs():
    data_dict = {}

    logging.info("📝 First, enter some details:")

    data_dict["Name of shareholder"] = input("👤 Name of shareholder: ")
    data_dict["Identifying Number"] = input("🆔 Identifying Number (e.g., SSN): ")
    data_dict["City, State, Zip, Country"] = input("🌍 City, State, Zip, Country: ")
    data_dict["Address"] = input("🏠 Address: ")
    data_dict["Tax year"] = input("📅 Tax year (last two digits): ")

    output_format = (
        input("📄 Output format — [P]DF (default) or [T]XT for tax software entry: ")
        .strip()
        .lower()
    )
    data_dict["output_format"] = "txt" if output_format in ("t", "txt") else "pdf"

    files = glob.glob("inputs/*.xlsx")
    return data_dict, files


def main():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    logging.info("🚀 Form 8621 Filler Initialized")

    try:
        data_dict, files = read_inputs()
        if not files:
            logging.error(
                "💥 No input files found in 'inputs/' directory. Please consult the README for instructions."
            )
            exit(1)

        OUTPUT_FOLDER = f"./outputs/20{data_dict['Tax year']}/"
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        logging.info(f"📁 Output directory: {OUTPUT_FOLDER}")

        total_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}

        for file in files:
            file_name = file.split("/")[-1].split(".")[0]
            logging.info(f"📂 Processing PFIC: {file_name}")

            if data_dict["output_format"] == "txt":
                FORM_OUTPUT_PATH = f"{OUTPUT_FOLDER}{file_name}.txt"
                number_of_lots, pfic_summary = generate_text_output(
                    path=FORM_OUTPUT_PATH, data_dict=data_dict, xlsx=file
                )
            else:
                FORM_OUTPUT_PATH = f"{OUTPUT_FOLDER}{file_name}.pdf"
                number_of_lots, pfic_summary = create_filled_pdf(
                    output_path=FORM_OUTPUT_PATH, data_dict=data_dict, xlsx=file
                )

            total_summary["ordinary_gains"] += pfic_summary["ordinary_gains"]
            total_summary["ordinary_losses"] += pfic_summary["ordinary_losses"]
            total_summary["capital_losses"] += pfic_summary["capital_losses"]

            logging.info(f"  ✅ Form completed and saved to {FORM_OUTPUT_PATH}")

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

    except Exception as e:
        logging.error(f"💥 An error occurred: {e}")
    finally:
        logging.info("👋 Shutting down. Goodbye!")


if __name__ == "__main__":
    main()
