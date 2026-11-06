# Expense reports: GBP reimbursement and missing receipts

Create a company-ready CSV and a portable evidence report for business travel. This version supports US trip expenses paid on a GBP card and reimbursed in GBP. It records actual posted GBP amounts, rather than estimating exchange rates. Manual report generation needs no packages or login. Optional Gmail collection uses Google OAuth and the dependencies below. No LaTeX installation is needed. Use Python 3.10 or newer on Mac, Windows, or Linux.

## Start a trip

From this repository folder:

```sh
python3 expense_report.py init trips/us-visit
```

On Windows, use `py` instead of `python3` if appropriate.

1. Fill in `trips/us-visit/trip.json`: claimant, company, project, destination, trip dates, business purpose, and the date you will review the declaration. Dates use `YYYY-MM-DD`. Optionally add `approver` and `trip_evidence_files`, a list such as `["evidence/itinerary.pdf"]`.
2. Put receipts, card statement extracts, and itinerary/meeting evidence in that trip's `evidence` folder. Supported evidence: PDF, PNG, JPG, JPEG, WebP and plain text (including imported email snapshots). Redact full card numbers and unrelated transactions before adding evidence. Keep unredacted originals privately.
3. Open `expenses.csv` in a spreadsheet editor and add one row per purchase. Preserve the headers and save as CSV. Use a text editor for `trip.json`.
4. Generate a new output folder:

```sh
python3 expense_report.py build trips/us-visit --output output/us-visit-v1
```

Open `output/us-visit-v1/report.html` in your browser. Review the CSV, report, missing-receipt declaration, and all evidence. Sign the declaration only after verifying it. Share the **whole output folder** (zip it if needed): it contains the CSV, HTML report, trip details, and numbered evidence copies. You can print the HTML to PDF using the browser, but the PDF does **not** embed the linked evidence; send the evidence folder too. The approver should retain approval and payment records.

Existing output folders are never overwritten. Use `--output output/us-visit-v2` after making corrections. Keep outputs separate from trip inputs.

## CSV fields

| Field | What to enter |
| --- | --- |
| expense_id | Unique reference, e.g. M001; letters, digits, underscores or hyphens |
| purchase_date | Date of the purchase, within the trip dates |
| posted_date | Date the settled charge appeared on the card; may be after the trip |
| merchant / location | Merchant name and city/country |
| category | Use `meal` for breakfasts/lunches/dinners; identify meals separately from other travel |
| description | What was purchased, e.g. Breakfast, solo |
| business_purpose | Connection to the trip/project; not just “business” |
| attendees | Optional; identify other diners and business relationship when relevant |
| original_amount / original_currency | Original USD charge if known; fill both or leave both blank |
| posted_gbp | Actual settled GBP card charge, including any tip/tax already included; no currency symbol |
| receipt_status | `receipt`, `missing`, or `requested` |
| missing_reason | Required for missing/requested receipts; state what happened and any attempt to obtain a duplicate |
| receipt_file | Relative path to merchant receipt, e.g. evidence/M001.jpg; required for `receipt` status |
| card_evidence_file | Relative path to card statement extract; required for missing/requested receipts |
| notes | Mark reconstructed details explicitly; record relevant context, tips or known limitations |

Use decimal points without thousands separators. Positive amounts only: refunds/chargebacks must be reconciled before submission; this version does not model them. Currency conversion and bank CSV layouts vary, so this version uses a reviewed expense template rather than automatic bank import. Do not guess unknown original amounts or itemized tax. Record separate card fees separately, only if the company agrees to reimburse them; do not duplicate fees already included in a posted charge. Use one settled entry per purchase, not both pending and posted entries.

## Missing receipts and Canadian company records

For small meals without a receipt, keep the actual card charge, merchant/location, purchase date, meal description, trip business purpose, and explanation. Request a duplicate when possible. A signed declaration is **your own statement**, not a replacement merchant receipt or a guarantee of CRA acceptance. The report remains labelled “awaiting signature”; it does not claim a declaration has been signed or a reimbursement approved.

The company can reimburse an agreed full amount even where its tax deduction is limited. Meals are separately identified so its accountant can apply the applicable treatment (generally a 50% limit, with exceptions). US sales tax is not Canadian GST/HST. This tool does not calculate tax credits, tax deductibility, or CAD book values. The company handles its CAD accounting conversion separately from your GBP reimbursement.

Useful official guidance:
- [CRA business records](https://www.canada.ca/en/revenue-agency/services/tax/businesses/topics/sole-proprietorships-partnerships/business-records.html)
- [CRA business expenses and meals](https://www.canada.ca/en/revenue-agency/services/tax/businesses/topics/sole-proprietorships-partnerships/business-expenses.html)
- [CRA out-of-pocket expenses for suppliers/contractors](https://www.canada.ca/en/revenue-agency/services/forms-publications/publications/gi-197/out-of-pocket-expenses.html)

## Privacy and maintenance

Keep trip inputs, bank statements, receipts, generated reports, credentials, and login tokens out of Git. `.gitignore` excludes default private folders, but review `git status` before committing if you put records elsewhere. No evidence is uploaded by this generator. CSV text that could be treated as a spreadsheet formula is prefixed with an apostrophe for safety.

Run checks with:

```sh
python3 -m unittest -v
```

`mk_expense_doc.py` now invokes the updated Gmail collector with the same arguments as `gmail_import.py`. The unsafe, hardcoded 2022 script has been replaced; the old no-argument invocation is no longer supported. LaTeX templates remain for reference and are not used by the new workflow. Bank reconciliation remains manual.


## Collect receipts from Gmail

First initialize a trip as above. Create a Gmail label such as `Expenses` and apply it to the relevant receipt emails, or use a company subject search as before. Both read and unread emails are included. Nothing is sent, marked read, deleted or relabelled.

1. Follow [Google's Python quickstart](https://developers.google.com/workspace/gmail/api/quickstart/python) to enable Gmail API in a Google Cloud project, configure the OAuth consent screen (add yourself as a test user if applicable), and create a **Desktop app** OAuth client. Download its JSON file as `credentials.json` in the repository folder. Never commit it.
2. Install the optional dependencies:

```sh
python3 -m pip install -r requirements-gmail.txt
```

3. Collect the trip emails (replace the example dates):

```sh
python3 gmail_import.py trips/us-visit --query 'label:Expenses' --start 2026-10-01 --end 2026-10-10
```

The first run opens Google's sign-in/consent page requesting read-only Gmail access. Later runs reuse `token.json`. You can supply `--credentials PATH --token PATH` for a different account; use private paths. Revoke access through Google when no longer needed. A test-mode OAuth application can require periodic sign-in again.

The dates filter **email receipt dates, inclusively in UTC**, not purchase dates. Widen the window to include late/forwarded receipts. For the old subject convention, use `--query 'subject:CompanyName'`. Queries use Gmail API search semantics; OR searches are grouped with the date filter. The collector retrieves every result page, recursively reads MIME parts, stores supported attachments under collision-resistant names, and records unsupported/empty attachments for review.

### Review the import and build

- `gmail_candidates.csv` contains one draft row per email, with email text and evidence paths in the notes. Actual purchase/posting dates, merchant, business purpose, GBP amount and receipt status are deliberately blank: email arrival dates and receipt currency do not establish settled card amounts. Review these against your card statement.
- Supported attachments and a readable `*-email.txt` snapshot are saved in `evidence/`. HTML-only email bodies are converted to text without rendering remote content. Exact email metadata, readable body, source message ID and attachment paths are retained in `gmail-import/*.json`; unsupported attachment types are listed there and in the draft notes. Check Gmail for those attachments manually.
- Select the correct receipt file and fill `receipt_file`; logos and signature images can also be attachments, so nothing is automatically labelled a valid receipt. Body-only merchant receipts can use the generated email text file after review. A forwarded note describing your purchase is your own explanation, not a merchant receipt.
- Copy reviewed rows (without a second header) into `expenses.csv`. If one email covers multiple purchases, split it into separate rows with distinct IDs and the correct amounts/evidence. Add card evidence and an explanation for genuinely missing receipts. Do not duplicate entries already recorded manually.
- For multiple supporting files, put the main receipt in `receipt_file` and add the other relevant paths, including the email snapshot if useful, to `trip.json`'s `trip_evidence_files` list. Only referenced evidence is included in the generated pack. Replace draft review notes with useful submission context; source JSON remains private unless you intentionally share it.
- Run the usual `expense_report.py build` command. The same validation applies to manually entered and email-derived expenses.

Repeat the Gmail import to append new emails. Existing candidate rows and `expenses.csv` are left intact. Deduplication uses the Gmail account plus message ID: keep existing candidate rows even after copying them to expenses, and do not change their IDs. Forwarded copies with different Gmail message IDs still need manual duplicate review. Imported emails are treated as snapshots; the command does not refresh existing rows. An interrupted API request fails visibly and leaves the existing candidate CSV intact; rerunning reuses matching downloaded evidence.

Tests simulate pagination, deeply nested and duplicate-named attachments, attachment API fetches, HTML-only and attachment-free emails, date boundaries, repeated imports and interrupted downloads, as well as report generation. No real mailbox is contacted by the tests. Live OAuth/mailbox access must be checked with your own credentials.
