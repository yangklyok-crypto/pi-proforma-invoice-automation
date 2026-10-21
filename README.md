# PI / Proforma Invoice Automation

A local Python web demo for generating proforma invoices. It is a generic public adaptation of a private v3.9 workflow, not a copy of its company workbook or product database.

Features: multiple line items, NGN or USD, an optional 7.5% VAT portion calculated *within* quoted prices, automatic invoice numbers, Excel and PDF exports, and optional product specification PDF pages appended once per selected file. The included example uses invented names and prices.

## Requirements

Python 3.11 or newer. Tested with the versions in `requirements.txt`. No Microsoft Excel installation is required. Dependencies are installed from their own distributions and retain their respective licenses; the MIT license here applies to this repository's original code only.

## Run locally

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:8000`. The default seller is `Example Supplier`. Set `PI_SELLER_NAME`, `PI_SELLER_DETAILS`, `PI_DATA_DIR`, `PI_HOST`, and `PI_PORT` as environment variables if needed; `.env.example` shows their names. The app does not automatically load `.env`.

The app is intended for one trusted user on localhost. It has no authentication or multiuser access controls. Do not expose its port to the internet or a public tunnel. Customer information, uploaded specification PDFs, generated invoices, and the number database live in `data/`, which is ignored by Git. If you customize the app, inspect every file and Git commit before making your repository public.

## API and examples

`examples/sample_invoice.json` contains fictitious values. POST the same JSON shape to `/api/invoices`. The result contains download links for the PDF and `.xlsx`. Upload a PDF (up to 10 MB and 30 pages) to `/api/specs`, then put its returned `id` in a line's `spec_id`. Omit the ID or use `null` to skip specification pages. Uploads are stored under random IDs, with no original filenames retained.

Invoice numbers use `PFI-YEAR-NNNN` and a local SQLite counter. Number allocation can leave a gap if rendering fails. This demo does not generate tax compliant invoices, exchange rates, bank instructions, approvals, email delivery, or accounting integrations. Review your own invoice and tax requirements before business use.

The Excel document uses a new generic fixed layout created in code. It does not include or reproduce any private template, logo, customer list, product catalog, real quote, or specification PDF. To change the layout, edit `render_xlsx` in `invoice_core.py`.

## Test

```bash
python -m unittest discover -s tests -v
```

The tests cover totals, inclusive VAT, numbering, exports, one time PDF attachment, and invalid input. They use temporary folders and invented data.

## Source and release boundary

The private v3.9 package informed the feature list. This repository is a clean implementation for public demonstration. It deliberately excludes all private archives and their Git history, the original Excel workbook and PDF parameter library, company branding and contact details, real customer and sales records, and product master data. A user should inspect the exact staged files before first public release.

