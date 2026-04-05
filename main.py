import pdfrw
from reportlab.pdfgen import canvas
import pandas as pd
import numpy as np
import glob
import os
from f8621_xy_coordinates import get_coordinates
import logging

CHECKMARK = "\u2713"


def create_overlay(path: str, data_dict: dict, xlsx: str):
    """
    Create the data that will be overlayed on top
    of the form that we want to fill
    """
    tax_year = 2000 + int(data_dict["Tax year"])
    df_lot = pd.read_excel(xlsx, sheet_name="Lot Details")
    df_eoy = pd.read_excel(xlsx, sheet_name="EOY Details")
    df_pfic = pd.read_excel(xlsx, sheet_name="PFIC Details")
    number_of_lots = len(df_lot.index)
    logging.info(f"  📊 Found {number_of_lots} lots to process")
    logging.debug(f"  📊 Lot details dataframe:\n{df_lot}")
    c = canvas.Canvas(path)
    coordinates = get_coordinates()
    add_personal_info(c, coordinates, data_dict)
    add_pfic_info(c, coordinates, df_pfic)
    add_part_1(c, coordinates, df_lot, df_eoy, tax_year)
    add_part_2(c, coordinates, data_dict)

    # Track gains and losses for this PFIC
    pfic_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}

    for lot in range(number_of_lots):
        logging.info(f"  🔄 Processing lot {lot + 1}/{number_of_lots}")
        result = add_part_4(c, coordinates, df_lot, df_eoy, lot, tax_year)
        if isinstance(result, tuple):
            processed, lot_summary = result
            if not processed:
                logging.info(f"    ⏭️ Skipping lot {lot + 1} (sale in different year)")
                number_of_lots = number_of_lots - 1
            else:
                # Add to PFIC summary
                pfic_summary["ordinary_gains"] += lot_summary.get("ordinary_gains", 0)
                pfic_summary["ordinary_losses"] += lot_summary.get("ordinary_losses", 0)
                pfic_summary["capital_losses"] += lot_summary.get("capital_losses", 0)
        else:
            if not result:
                logging.info(f"    ⏭️ Skipping lot {lot + 1} (sale in different year)")
                number_of_lots = number_of_lots - 1

    c.save()
    return number_of_lots, pfic_summary


def add_personal_info(c, coordinates, data_dict):
    keys = [
        "Name of shareholder",
        "Identifying Number",
        "Address",
        "City, State, Zip, Country",
        "Tax year",
        "Type of Shareholder",
    ]
    for key in keys:
        c.drawString(coordinates[key][0], coordinates[key][1], data_dict[key])

    c.drawString(196, 627, CHECKMARK)  # type of shareholder


def add_pfic_info(c, coordinates, df_pfic: pd.DataFrame):
    print(df_pfic)
    keys = ["PFIC Name", "PFIC Address", "PFIC Reference ID", "PFIC Share Class"]
    for key in keys:
        value = df_pfic[key].values[0]
        if key == "PFIC Address" and isinstance(value, str) and "\n" in value:
            lines = value.split("\n")
            y = coordinates[key][1]
            for line in lines:
                c.drawString(coordinates[key][0], y, line)
                y -= 12
        else:
            c.drawString(coordinates[key][0], coordinates[key][1], value)


def add_part_1(c, coordinates, df_lot, df_eoy, current_year):
    part_1_dict = {}
    part_1_dict["Date of Acquisition"] = (
        pd.to_datetime(df_lot["Date: Acquisition"].values[0]).strftime("%Y-%m-%d")
        if len(df_lot.index) == 1
        else "Multiple"
    )
    part_1_dict["Number of Shares"] = 0
    part_1_dict["Amount of 1291"] = ""
    part_1_dict["Amount of 1293"] = ""
    for lot in range(len(df_lot.index)):
        # Check if lot was sold and get last price and ER
        if np.isnan(df_lot["Price per share: Sale"][lot]):
            part_1_dict["Number of Shares"] = (
                part_1_dict["Number of Shares"] + df_lot["Number of shares"][lot]
            )

    last_er = df_eoy[df_eoy["Year"] == current_year]["Exchange Rate"].values[0]
    last_price = df_eoy[df_eoy["Year"] == current_year]["Price"].values[0]
    part_1_dict["Amount of 1296"] = round(
        part_1_dict["Number of Shares"] * last_price / last_er
    )

    for key in part_1_dict.keys():
        c.drawString(
            coordinates[key][0], coordinates[key][1], "{}".format(part_1_dict[key])
        )

    value_of_pfic = part_1_dict["Amount of 1296"]
    if (value_of_pfic >= 0) and (value_of_pfic <= 50000):
        c.drawString(79.2, 373.5, CHECKMARK)  # value of pfic
    elif (value_of_pfic > 50000) and (value_of_pfic <= 100000):
        c.drawString(151.2, 373.5, CHECKMARK)  # value of pfic
    elif (value_of_pfic > 100000) and (value_of_pfic <= 150000):
        c.drawString(245, 373.5, CHECKMARK)  # value of pfic
    elif (value_of_pfic > 150000) and (value_of_pfic <= 200000):
        c.drawString(345.6, 373.5, CHECKMARK)  # value of pfic
    else:
        c.drawString(199, 362, "{}".format(value_of_pfic))  # value of pfic

    # Check marks
    c.drawString(79.2, 290, CHECKMARK)  # type of PFIC type c


def add_part_2(c, coordinates, data_dict):
    c.drawString(52.4, 205.5, CHECKMARK)  # Part II election to MTM PFIC stock


def add_part_4(c, coordinates, df_lot, df_eoy, lot, current_year):
    etf_dict = {}
    # Get info about origianl aquisition
    year_of_aqiusition = df_lot["Date: Acquisition"][lot].year
    cost_aquisition = df_lot["Cost: Acquisition"][lot]
    er_of_aqiusition = df_lot["Exchange Rate: Acquisition"][lot]

    number_of_shares = df_lot["Number of shares"][lot]
    original_basis = cost_aquisition / er_of_aqiusition

    logging.debug(
        f"    📊 Lot {lot + 1} details - Shares: {number_of_shares:.2f}, Original basis: ${original_basis:.2f}"
    )

    # Get last year's basis
    if current_year > year_of_aqiusition:
        prev_year_er = df_eoy[df_eoy["Year"] == current_year - 1][
            "Exchange Rate"
        ].values[0]
        prev_year_price = df_eoy[df_eoy["Year"] == current_year - 1]["Price"].values[0]
        adjusted_basis = round(number_of_shares * prev_year_price / prev_year_er)
    else:
        adjusted_basis = round(original_basis)

    # Track gains/losses for this lot
    lot_summary = {"ordinary_gains": 0, "ordinary_losses": 0, "capital_losses": 0}

    # Check if lot was sold and get last price and ER
    if np.isnan(df_lot["Price per share: Sale"][lot]):
        logging.info(f"    📈 Lot {lot + 1}: No sale (holding position)")
        last_er = df_eoy[df_eoy["Year"] == current_year]["Exchange Rate"].values[0]
        last_price = df_eoy[df_eoy["Year"] == current_year]["Price"].values[0]
        fmv_dollars = round(number_of_shares * last_price / last_er)
        logging.debug(f"    💱 Exchange rate: {last_er}, Price: ${last_price}")
        logging.debug(f"    💰 FMV: ${fmv_dollars}, Adjusted basis: ${adjusted_basis}")

        etf_dict["10a"] = fmv_dollars
        etf_dict["10b"] = adjusted_basis
        etf_dict["10c"] = etf_dict["10a"] - etf_dict["10b"]
        if etf_dict["10c"] < 0:
            if adjusted_basis > original_basis:
                unreversed_inclusions = round(adjusted_basis - original_basis)
                if unreversed_inclusions > (-1 * etf_dict["10c"]):
                    loss_from_ten_c = etf_dict["10c"]
                else:
                    loss_from_ten_c = -1 * unreversed_inclusions
                etf_dict["11"] = unreversed_inclusions
                etf_dict["12"] = loss_from_ten_c
                logging.info(
                    f"    📉 Lot {lot + 1}: Ordinary loss of ${abs(etf_dict['12'])}"
                )
                lot_summary["ordinary_losses"] += abs(etf_dict["12"])
            else:
                logging.info(
                    f"    📉 Lot {lot + 1}: Unrecognizable loss of ${abs(etf_dict['10c'])}"
                )
                etf_dict["11"] = "0"
                etf_dict["12"] = "0"
        else:
            etf_dict["11"] = ""
            etf_dict["12"] = ""
            logging.info(f"    📈 Lot {lot + 1}: Ordinary gain of ${etf_dict['10c']}")
            lot_summary["ordinary_gains"] += etf_dict["10c"]
        etf_dict["13a"] = ""
        etf_dict["13b"] = ""
        etf_dict["13c"] = ""
        etf_dict["14a"] = ""
        etf_dict["14b"] = ""
        etf_dict["14c"] = ""

    else:
        logging.info(f"    💰 Lot {lot + 1}: Sale detected")
        last_er = df_lot["Exchange Rate: Sale"][lot]
        last_price = df_lot["Price per share: Sale"][lot]
        year_of_sale = df_lot["Date: Sale"][lot].year
        if year_of_sale < current_year:
            return False, lot_summary
        fmv_dollars = round(number_of_shares * last_price / last_er)
        logging.debug(
            f"    💱 Sale exchange rate: {last_er}, Sale price: ${last_price}"
        )
        logging.debug(
            f"    💰 Sale proceeds: ${fmv_dollars}, Adjusted basis: ${adjusted_basis}"
        )
        etf_dict["13a"] = round(fmv_dollars)
        etf_dict["13b"] = round(adjusted_basis)
        etf_dict["13c"] = etf_dict["13a"] - etf_dict["13b"]
        if etf_dict["13c"] < 0:
            if adjusted_basis > original_basis:
                unreversed_inclusions = round(adjusted_basis - original_basis)
                if unreversed_inclusions > (-1 * etf_dict["13c"]):
                    loss_from_thirteen_c = etf_dict["13c"]
                else:
                    loss_from_thirteen_c = -1 * unreversed_inclusions
                etf_dict["14a"] = unreversed_inclusions
                etf_dict["14b"] = loss_from_thirteen_c
                etf_dict["14c"] = ""
                logging.info(
                    f"    📉 Lot {lot + 1}: Ordinary loss of ${abs(etf_dict['14b'])}"
                )
                lot_summary["ordinary_losses"] += abs(etf_dict["14b"])
            else:
                etf_dict["14a"] = 0
                etf_dict["14b"] = 0
                etf_dict["14c"] = etf_dict["13c"]
                logging.info(
                    f"    📉 Lot {lot + 1}: Capital loss of ${abs(etf_dict['14c'])}"
                )
                lot_summary["capital_losses"] += abs(etf_dict["14c"])
        else:
            etf_dict["14a"] = ""
            etf_dict["14b"] = ""
            etf_dict["14c"] = ""
            logging.info(f"    📈 Lot {lot + 1}: Ordinary gain of ${etf_dict['13c']}")
            lot_summary["ordinary_gains"] += etf_dict["13c"]

    c.showPage()
    for key in etf_dict.keys():
        if key in coordinates:
            c.drawString(
                coordinates[key][0], coordinates[key][1], "{}".format(etf_dict[key])
            )
        else:
            logging.warning(f"    ⚠️ Coordinate missing for {key} in lot {lot + 1}")
    return True, lot_summary


def merge_pdfs(pdf_1, pdf_2, output):
    """
    Merge the specified fillable form PDF with the
    overlay PDF and save the output
    """
    form = pdfrw.PdfReader(pdf_1)
    olay = pdfrw.PdfReader(pdf_2)

    for form_page, overlay_page in zip(form.pages, olay.pages):
        merge_obj = pdfrw.PageMerge()
        overlay = merge_obj.add(overlay_page)[0]
        pdfrw.PageMerge(form_page).add(overlay).render()

    writer = pdfrw.PdfWriter()
    writer.write(output, form)


def split(path, page, output):
    pdf_obj = pdfrw.PdfReader(path)
    total_pages = len(pdf_obj.pages)

    writer = pdfrw.PdfWriter()

    if page <= total_pages:
        writer.addpage(pdf_obj.pages[page])

    writer.write(output)


def concatenate(paths, output):
    writer = pdfrw.PdfWriter()

    for path in paths:
        reader = pdfrw.PdfReader(path)
        writer.addpages(reader.pages)

    writer.write(output)


def create_full_8621(path, number_of_page_2, output):
    orig_path = path + ".pdf"
    page_1_path = path + "page1.pdf"
    page_2_path = path + "page2.pdf"
    split(orig_path, 0, page_1_path)
    split(orig_path, 1, page_2_path)
    concatenate([page_1_path, page_2_path], output)
    for page in range(number_of_page_2 - 1):
        concatenate([output, page_2_path], output)

    os.remove(page_1_path)
    os.remove(page_2_path)


def read_inputs():
    data_dict = {}

    logging.info("📝 First, enter some details:")

    data_dict["Name of shareholder"] = input("👤 Name of shareholder: ")
    data_dict["Identifying Number"] = input("🆔 Identifying Number (e.g., SSN): ")
    data_dict["City, State, Zip, Country"] = input("🌍 City, State, Zip, Country: ")
    data_dict["Address"] = input("🏠 Address: ")
    data_dict["Tax year"] = input("📅 Tax year (last two digits): ")
    data_dict["Type of Shareholder"] = CHECKMARK  # assuming always an individual

    output_format = (
        input("📄 Output format — [P]DF (default) or [T]XT for tax software entry: ")
        .strip()
        .lower()
    )
    data_dict["output_format"] = "txt" if output_format in ("t", "txt") else "pdf"

    # Placeholder values:
    # data_dict['Name of shareholder'] = 'John Doe'
    # data_dict['Identifying Number'] = '123-45-6789'
    # data_dict['City, State, Zip, Country'] = 'Anytown, ST 12345'
    # data_dict['Address'] = '123 Main St'
    # data_dict['Tax year'] = '25'
    # data_dict['Type of Shareholder'] = '\u2713'  # assuming always an individual

    files = glob.glob("inputs/*.xlsx")

    return data_dict, files


def generate_text_output(path: str, data_dict: dict, xlsx: str):
    """
    Generate a plain-text summary of Form 8621 fields, suitable for
    manual entry into tax software.  Returns (number_of_lots, pfic_summary)
    matching the same contract as create_overlay().
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
            # Unsold lot
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
            # Sold lot
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


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    logging.info("🚀 Form 8621 Filler Initialized")

    try:
        data_dict, files = read_inputs()
        if not files:
            logging.error(
                "💥 No input files found in 'inputs/' directory. Please consult the README for instructions."
            )
            exit(1)

        form = "f8621"

        OUTPUT_FOLDER = f"./outputs/20{data_dict['Tax year']}/"
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        logging.info(f"📁 Output directory: {OUTPUT_FOLDER}")

        # Track totals across all PFICs
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
                FORM_FULL_PATH = f"{OUTPUT_FOLDER}{file_name}_full.pdf"
                FORM_OVERLAY_PATH = f"{OUTPUT_FOLDER}{file_name}_overlay.pdf"
                FORM_OUTPUT_PATH = f"{OUTPUT_FOLDER}{file_name}.pdf"

                number_of_lots, pfic_summary = create_overlay(
                    path=FORM_OVERLAY_PATH, data_dict=data_dict, xlsx=file
                )
                create_full_8621(form, number_of_lots, FORM_FULL_PATH)
                merge_pdfs(FORM_FULL_PATH, FORM_OVERLAY_PATH, FORM_OUTPUT_PATH)

                # Delete intermediate files
                os.remove(FORM_FULL_PATH)
                os.remove(FORM_OVERLAY_PATH)

            # Add to total summary
            total_summary["ordinary_gains"] += pfic_summary["ordinary_gains"]
            total_summary["ordinary_losses"] += pfic_summary["ordinary_losses"]
            total_summary["capital_losses"] += pfic_summary["capital_losses"]

            logging.info(f"  ✅ Form completed and saved to {FORM_OUTPUT_PATH}")

        logging.info("✅ All forms processed successfully!")

        # Display summary and instructions with visual emphasis
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
