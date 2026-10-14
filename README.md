# Anshita's Document Studio

Anshita's Document Studio is a polished Flask application for creating professional invoices and academic result documents from one browser-based workspace. It combines a presentation-focused landing page with practical document automation:

- Tax-aware invoice PDF generation
- Individual and bulk academic marksheet PDF generation
- Class performance calculation with grades, CGPA, ranks, and pass/fail status
- Excel class-summary export
- Drag-and-drop CSV workflows
- A built-in page-copy report for reviewing the product's visible content

The project is intentionally lightweight: Flask provides the web application and API layer, ReportLab creates PDFs, and openpyxl creates `.xlsx` reports. Generated files are written to the local `outputs/` directory.

## Project status

This is a local-first project intended for portfolio demonstrations, academic administration prototypes, and small internal workflows. It does not currently include authentication, a database, cloud storage, or multi-user access control.

## Features

### Invoice generation

- Manual invoice entry with company and client details
- Dynamic line items with quantity, rate, and tax percentage
- Automatic subtotal, tax, discount, grand-total, and due-date calculations
- Support for CGST/SGST-style tax presentation
- Amount-in-words output in Indian currency format
- Downloadable, styled A4 PDF invoices

### Academic result generation

- Configurable subjects, maximum marks, pass marks, and credits
- Multiple student rows with roll numbers and personal details
- Per-subject marks validation and calculations
- Percentage, CGPA, grade, rank, and result status
- Failed-subject tracking
- Individual student marksheet PDFs
- Bulk marksheet PDF generation
- Class-summary Excel export
- Preview table before downloading generated files

### Workspace experience

- Personalizable student/user greeting
- Theme selector with indigo, rose, emerald, and violet palettes
- Autosaved workspace mode and user-name preferences through browser local storage
- Responsive glassmorphism interface
- Landing page, workspace, and page-report routes

## Technology stack

| Layer | Technology |
| --- | --- |
| Backend | Python, Flask |
| Invoice and marksheet PDFs | ReportLab |
| Spreadsheet export | openpyxl |
| Frontend | HTML, Tailwind CSS CDN, vanilla JavaScript |
| Icons and typography | Phosphor Icons CDN, Google Fonts |
| Persistence | Local filesystem and browser `localStorage` |

## Requirements

- Python 3.10 or newer recommended
- `pip`
- A modern browser

The application currently imports these Python packages:

- `Flask`
- `reportlab`
- `openpyxl`

## Installation

### 1. Clone or copy the project

```bash
cd "Invoice and result generator"
```

### 2. Create a virtual environment

macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Windows PowerShell:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```bash
python -m pip install Flask reportlab openpyxl
```

### 4. Start the application

```bash
python app.py
```

The application starts on:

```text
http://127.0.0.1:5000
```

Open the URL in a browser and select **Launch Workspace** to begin creating documents.

## Application routes

| Route | Method | Purpose |
| --- | --- | --- |
| `/` | `GET` | Landing page |
| `/home` | `GET` | Landing-page alias |
| `/workspace` | `GET` | Invoice and marksheet workspace |
| `/report` | `GET` | Visible-copy/page-content report |
| `/api/generate-invoice` | `POST` | Generate an invoice PDF |
| `/api/generate-marksheet` | `POST` | Generate marksheet PDFs and/or an Excel summary |
| `/download/<filename>` | `GET` | Download a generated file |

## API usage

### Generate an invoice

Send JSON to `/api/generate-invoice`:

```json
{
  "company": {
    "name": "Anshita's Document Studio",
    "tagline": "Professional document automation",
    "address": "Your business address",
    "gstin": "N/A",
    "email": "your-email@example.com",
    "phone": "+91 00000 00000",
    "website": "",
    "bank_details": "Payment details"
  },
  "client": {
    "name": "Client Organization",
    "address": "Client billing address",
    "gstin": "N/A",
    "email": "client@example.com",
    "contact": "Accounts Team"
  },
  "meta": {
    "invoice_no": "INV-001",
    "invoice_date": "2026-10-14",
    "due_days": 30,
    "place_of_supply": "India",
    "currency_symbol": "Rs.",
    "notes": "Thank you for choosing us."
  },
  "tax_mode": "cgst_sgst",
  "line_items": [
    {
      "description": "Professional Service",
      "qty": 1,
      "rate": 1000,
      "tax_pct": 18,
      "discount_pct": 0
    }
  ]
}
```

Successful responses include the generated filename and download URL:

```json
{
  "status": "ok",
  "filename": "invoice_INV-001.pdf",
  "download": "/download/invoice_INV-001.pdf"
}
```

### Generate marksheets

Send JSON to `/api/generate-marksheet`:

```json
{
  "institution": {
    "name": "Anshita's Learning Studio",
    "exam_name": "Final Examination",
    "program": "Computer Science",
    "semester": "Semester 1",
    "logo_text": "AS",
    "address": "Your institution address",
    "affiliation": "Anshita's Project"
  },
  "subjects": [
    {
      "code": "CS101",
      "name": "Computer Science",
      "max_marks": 100,
      "pass_marks": 40,
      "credits": 4
    }
  ],
  "students": [
    {
      "roll_no": "AS-001",
      "name": "Demo Student",
      "father": "",
      "dob": "",
      "marks": [85]
    }
  ],
  "mode": "all"
}
```

Supported `mode` values:

- `single`: create a PDF for the first student
- `all`: create PDFs for all students and one class-summary Excel file
- `excel`: create only the class-summary Excel file

The response includes:

- `files`: generated file names and download URLs
- `preview`: computed totals, percentages, CGPA, grades, ranks, and result status

## CSV workflow

The workspace supports CSV-based bulk entry through the browser interface. Use the bulk-upload controls in the invoice or marksheet module, select a compatible CSV file, review the imported rows, and then generate the requested outputs.

For marksheet imports, keep the file structure aligned with the visible workspace fields: student identity columns followed by subject marks. Always review the preview table before downloading final documents.

## Result calculation model

The backend calculates academic results in this order:

1. Total obtained marks for each student
2. Total maximum marks
3. Overall percentage
4. Grade using the configured grade scale
5. CGPA derived from the grade-point scale
6. Pass/fail status based on subject pass marks
7. Rank based on descending percentage

The current grade scale is:

| Percentage | Grade | Grade points |
| ---: | --- | ---: |
| 90–100 | O | 10 |
| 80–89 | A+ | 9 |
| 70–79 | A | 8 |
| 60–69 | B+ | 7 |
| 50–59 | B | 6 |
| 40–49 | C | 5 |
| Below 40 | F | 0 |

## Project structure

```text
Invoice and result generator/
├── app.py                 # Flask application, calculations, PDF/XLSX builders, API routes
├── templates/
│   ├── home.html          # Landing page
│   ├── index.html         # Interactive invoice and marksheet workspace
│   └── report.html        # Page-copy report
├── outputs/               # Generated PDFs and Excel files
└── README.md              # Project documentation
```

## Generated files

Generated documents are saved locally under `outputs/`. The directory is created automatically when the application starts.

Typical outputs include:

- Invoice PDF files
- Individual marksheet PDFs
- Bulk marksheet PDFs
- Class result `.xlsx` summaries

Generated files should be treated as local working artifacts. Add an appropriate ignore rule for `outputs/` if these documents should not be committed to version control.

## Customization

### Change the project identity

Update the following areas together to keep the application consistent:

- Page titles and visible labels in `templates/home.html`
- Workspace defaults in `templates/index.html`
- Generated invoice and marksheet defaults in `templates/index.html`
- Page-report copy in `app.py`
- PDF styling constants near the top of `app.py`

### Change colors and themes

- Landing-page colors are configured in the Tailwind configuration inside `home.html`.
- Workspace themes are defined through CSS variables in `index.html`.
- PDF colors are controlled by the `BRAND`, `DARK`, `LIGHT`, and related constants in `app.py`.

### Change grading rules

Edit `GRADE_SCALE` in `app.py`, then verify both the result preview and generated marksheet output.

## Development notes

- The app runs in Flask debug mode when launched directly with `python app.py`.
- The frontend currently loads Tailwind CSS, Phosphor Icons, and Google Fonts from CDNs, so an internet connection is useful for the full visual experience.
- There is no database: browser preferences use `localStorage`, while generated documents use the local `outputs/` directory.
- API errors are returned as JSON with HTTP status `400`.
- Downloads return `404` when the requested generated file does not exist.

For production deployment, use a production WSGI server, disable debug mode, add authentication if documents contain sensitive information, validate uploaded files and request payloads, and configure a controlled storage location for generated documents.

## Troubleshooting

### The page loads without styling or icons

Check the browser's network connection. Tailwind CSS, Phosphor Icons, and Google Fonts are loaded from external CDNs.

### Generated files are not visible

Check the `outputs/` directory and confirm that the Flask process has write permission in the project directory.

### An API request returns an error

Inspect the JSON response's `message` field. Confirm that required objects such as `company`, `client`, `institution`, `subjects`, and `students` are present and correctly shaped.

### The marksheet result is incorrect

Verify each student's marks array has the same number of entries as the `subjects` array and that the subject maximum/pass marks are correct.

## License

No license has currently been specified for this project. Add a license before distributing it publicly or incorporating third-party contributions.