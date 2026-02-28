# QuickAriba Importer

This project contains:

- `ariba_excel_import.py`: imports non-catalog items from Excel into SAP Ariba Buying.
- `reichelt_offer_to_excel.py`: converts a Reichelt offer PDF into that Excel format.
- `thorlabs_cart_to_excel.py`: converts a Thorlabs cart CSV into that Excel format.

The login/2FA step remains manual. The script attaches to your already logged-in browser tab.

## Excel format

Use this exact layout in the first worksheet:

- `A1`: `Supplier name`
- `B1`: supplier value (example: `Reichelt`, `Thorlabs`)
- Row `3` headers exactly:
  - `Product name`
  - `Description`
  - `Quantity`
  - `Unit price`
- Data starts at row `4`

The script stops when it reaches the first completely empty data row.

## Reichelt PDF to Excel

Use this script when you receive a Reichelt offer PDF and want to generate the Excel file for the importer.

```powershell
.\.venv\Scripts\python .\reichelt_offer_to_excel.py --pdf ".\ExampleReicheltOffer.pdf" --out ".\orders_from_reichelt.xlsx"
```

Optional flags:

- `--supplier` default: `Reichelt` (written to cell `B1`)

Column mapping used:

- Reichelt `Item No.` -> Excel `Product name`
- Reichelt `Description` -> Excel `Description`
- Reichelt `Quantity` -> Excel `Quantity`
- Reichelt `Price` -> Excel `Unit price`

Ignored Reichelt columns:

- `Category of goods`
- `Price on all`

## Thorlabs CSV to Excel

Use this script when you export a Thorlabs cart CSV and want to generate the Excel file for the importer.

Typical workflow:

1. Export cart from Thorlabs as CSV.
2. Convert CSV to Ariba import Excel with the command below.
3. Run `ariba_excel_import.py` with the generated `.xlsx` file.

```powershell
.\.venv\Scripts\python .\thorlabs_cart_to_excel.py --csv ".\ExampleThorlabsCart.csv" --out ".\orders_from_thorlabs.xlsx"
```

Optional flags:

- `--supplier` default: `Thorlabs` (written to cell `B1`)

Column mapping used:

- Thorlabs `Item Number` -> Excel `Product name`
- Excel `Description` -> `<Thorlabs Description>` on first line, then ` URL: <Thorlabs URL>` on second line
- Thorlabs `Quantity` -> Excel `Quantity`
- Thorlabs `Unit Price` -> Excel `Unit price`

## Install

Recommended on Windows PowerShell (no venv activation needed):

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python -m playwright install chromium
```

Optional (if you prefer activation):

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m playwright install chromium
```

## Start browser for manual login + 2FA

Close existing Chrome/Edge windows first, then start one with remote debugging enabled.

Chrome example:

```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="$env:TEMP\quickariba-profile"
```

Edge example:

```powershell
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="$env:TEMP\quickariba-profile"
```

Then log into Ariba in that browser and open the `Non-catalog request` page.

## Run importer

Recommended (no activation):

```powershell
.\.venv\Scripts\python .\ariba_excel_import.py --excel ".\orders.xlsx"
```

Optional (with activation):

```powershell
.\.venv\Scripts\Activate.ps1
python .\ariba_excel_import.py --excel ".\orders.xlsx"
```

Optional flags:

- `--cdp-url` default: `http://127.0.0.1:9222`
- `--page-contains` default: `ariba`

## Automated Ariba flow

For each row:

1. Ensure supplier is selected (from `B1`) via `View all suppliers`.
2. Set category to `900006 (Laboratory - Items and disposables)`.
3. Fill `Product name`, `Description`, `Quantity`, `Unit price`.
4. Click `Add to cart`.
5. If more rows remain: open three-dot menu and click `Create new`.

At the end, the script stops and you continue checkout manually.

## Notes

- The script validates headers and numeric fields before starting browser actions.
- If supplier search returns no match, it fails fast with a clear error.
- Ariba UI updates can break selectors; if that happens, share a screenshot and we can adjust quickly.

## Troubleshooting

`Activate.ps1 cannot be loaded because running scripts is disabled`

- Use the no-activation commands (recommended):

```powershell
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python -m playwright install chromium
.\.venv\Scripts\python .\ariba_excel_import.py --excel ".\orders.xlsx"
```

- Or allow activation in the current PowerShell session only:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

`Automation failed: connect ECONNREFUSED 127.0.0.1:9222`

- Browser was not started with remote debugging. Start Chrome/Edge with `--remote-debugging-port=9222` as shown above.
- Make sure all old browser windows are closed before starting the debugging instance.

`Excel validation error: Row 3 headers must be exactly ...`

- Confirm row 3 is exactly: `Product name`, `Description`, `Quantity`, `Unit price`.
- Confirm data starts at row 4.

`Supplier search found no match for '<name>'`

- Check the value in `B1` and try a more specific supplier name.
- Open `View all suppliers` manually and verify the supplier is searchable in your tenant.

`No offer rows found in PDF table`

- Verify the PDF is a Reichelt offer table similar to `ExampleReicheltOffer.pdf`.
- If extraction still fails, install dependencies again: `.\.venv\Scripts\python -m pip install -r requirements.txt`.

