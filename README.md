# Form 8621 Filler

This program fills out IRS Form 8621, specifically for Mark-to-Market (MTM) elections.

<img width="500" alt="image" src="https://github.com/user-attachments/assets/970426a5-dd5a-46a1-a0a0-97e6c16de4f6" />

---

> [!CAUTION]
> This tool does not provide not tax advice, cannot be guaranteed to generate a compliant output, and is not responsible for any errors in the generated forms. It is your responsibility to ensure that the filled-out forms are correct and compliant with IRS regulations. Consult a tax professional if needed.

## Getting started
For each PFIC you'd like to generate a form for, copy the `f8621.xlsx` template file into the `inputs` folder and rename it to something meaningful (e.g. the reference ID of the PFIC).

Example:
```bash
cp ./f8621.xlsx ./inputs/vwce.xlsx # Fund 1
cp ./f8621.xlsx ./inputs/spyy.xlsx # Fund 2
```

Now you need to fill out the XLSX files with the relevant data for each fund. The program will read these files and generate a filled-out Form 8621 PDF for each.

## Filling out the form

### Transactions
In the `Transactions` sheet, list every buy and sell transaction for the PFIC. The program will automatically construct FIFO share lots from these transactions. Columns:

- **Date** — Date of the transaction
- **Type** — `buy` (or `purchase`, `reinvestment`) for acquisitions; `sell` (or `sale`, `distribution`) for dispositions
- **Number of shares** — Number of shares bought or sold
- **Total Value** — Total value of the transaction in the local currency, including any commissions or fees (for buys, this is the cost; for sells, this is the proceeds)

Exchange rates are looked up automatically from the Federal Reserve's historical daily rates.

### EOY Details
In the `EOY Details` sheet, fill out for each year:
- The year
- The FMV of a share of the PFIC on December 31st (in local currency)

Exchange rates for December 31st of each year are looked up automatically.

### PFIC Details
In the `PFIC Details` sheet, fill out:
- The name of the PFIC (e.g. Vanguard FTSE All-World UCITS ETF)
- The address of the PFIC (you can usually find this in the Prospectus by searching for "Registered Office")
- The reference ID
  - Recommended to use the ticker name, like VWCE or SPYY, without any special characters
  - Can be whatever you want, but must be consistent year-to-year
  - From the IRS:
    > The reference ID number must be alphanumeric [A-Z, 0-9] and no special characters or spaces are permitted. The length of a given reference ID number is limited to 50 characters.
- The share class (you can usually find this in the Prospectus by searching for "Share Class")
- The currency — ISO 4217 code (e.g. EUR, GBP, JPY) for the local currency the PFIC is denominated in

## Running the program
To run the program, first install `uv` by following the instructions in the [uv documentation](https://docs.astral.sh/uv/getting-started/installation/).

Next, install the dependencies:
```bash
uv install
```

Finally, run the program:
```bash
uv run main.py
```

This will ask you several questions in the terminal:
1. Name of the Shareholder (that's you!)
2. Identifying number (SSN or ITIN)
3. City, State, ZIP, Country (e.g. Amsterdam 1012AB, Netherlands)
4. Address (e.g. Some Street 123)
5. Tax year, just the last 2 digits (e.g. 19 for 2019)

That's all of the data needed. It will then generate a PDF for each PFIC in the `outputs/YEAR` folder, named `REFERENCE_ID.pdf`. For example, if you have `vwce.xlsx` and `spyy.xlsx` in the `inputs` folder, and run the program for 2025, it will generate:
- `outputs/2025/vwce.pdf`
- `outputs/2025/spyy.pdf`

## Running tests
You can run tests via:

```sh
uv run pytest
```