# IBAN Checker for Excel

Excel add-in for the [ibanchecker.cash](https://ibanchecker.cash) IBAN validation API. It checks a column of IBANs in place: whether each one is valid, why it failed, its country, its bank and the bank's BIC.

Validation is done by the ibanchecker.cash API, the same engine as the website, national check digits included, so a workbook and the site never disagree about an IBAN. No IBAN is stored or logged by the API.

It runs in Excel for Microsoft 365 on Windows and Mac, Excel 2021 and 2024, and Excel on the web. Excel on iPad does not run add-in formulas.

## Install

Until the add-in is listed in Microsoft Marketplace, add it from its manifest, [ibanchecker.cash/excel/manifest.xml](https://ibanchecker.cash/excel/manifest.xml):

- Excel on the web: Home > Add-ins > More Settings > Upload My Add-in.
- Excel for Mac: copy the file to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (create `wef` if missing) and restart Excel.
- Microsoft 365 administrators: admin center > Settings > Integrated apps > Add-ins > Deploy Add-in.

Then open **Home > IBAN Checker > Settings** and paste your key. A free key is available at [ibanchecker.cash/api-docs](https://ibanchecker.cash/api-docs).

## Formulas

Every formula takes a single cell or a range and returns a result of the same shape.

| Formula | Returns |
| --- | --- |
| `=IBAN.VALID(A2:A100)` | `TRUE` or `FALSE` |
| `=IBAN.ERROR(A2:A100)` | why the IBAN failed, empty when it is valid |
| `=IBAN.BANK(A2:A100)` | the bank's name |
| `=IBAN.BIC(A2:A100)` | the bank's BIC |
| `=IBAN.CHECK(A2:A100)` | five columns: valid, error, country, bank, BIC |

Spaces and lower case are accepted. A cell without any digit (a header such as `IBAN`) is skipped and not counted. A valid IBAN whose bank is missing from our directory reads `Not in our bank directory`. One formula checks at most 300 IBANs that were not checked in the last six hours; for a longer list use the button.

## Check selected IBANs

Select the IBANs and choose **Home > IBAN Checker > Check selected IBANs**. The results are written as values in the five columns to the right of the selection, with a header row when the selection sits under one. Nothing is overwritten: if those columns hold data the add-in asks first and inserts five new columns instead. Up to 5,000 rows per run.

## What it costs

Each IBAN checked counts as one request on your API key, the same as a call to `POST /api/v1/validate`. The same IBAN appearing several times in one formula or run is sent once. Results are kept for six hours in Office's storage on your device, under a SHA-256 hash rather than the IBAN, so checking the same IBAN again within six hours is not counted. Plans and limits: [ibanchecker.cash/pricing](https://ibanchecker.cash/pricing).

## Where the key is kept

In `OfficeRuntime.storage` on your device, never in the workbook, so sharing a workbook does not share your key.

## Develop

The add-in is plain JavaScript with no build step. `src/core.js` holds the logic and is tested in Node; `src/functions.js` and `src/taskpane.js` connect it to Excel. The files are served from `https://ibanchecker.cash/excel/`, the same origin as the API.

```bash
npm test
```

Microsoft and Excel are trademarks of the Microsoft group of companies.
