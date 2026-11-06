"""Generate a local GBP reimbursement CSV and printable evidence report. No dependencies."""
import argparse
import csv
from datetime import date
from decimal import Decimal, InvalidOperation
import html
import json
from pathlib import Path
import re
import shutil

FIELDS = ['expense_id', 'purchase_date', 'posted_date', 'merchant', 'location',
          'category', 'description', 'business_purpose', 'attendees',
          'original_amount', 'original_currency', 'posted_gbp', 'receipt_status',
          'missing_reason', 'receipt_file', 'card_evidence_file', 'notes']
STATUSES = {'receipt', 'missing', 'requested'}


def money(value, field):
    try:
        amount = Decimal(value)
    except InvalidOperation:
        raise ValueError(f'{field}: enter a decimal amount without a currency symbol') from None
    if not amount.is_finite() or amount <= 0 or amount != amount.quantize(Decimal('.01')):
        raise ValueError(f'{field}: amount must be positive with at most two decimal places')
    return amount


def load_rows(folder):
    with (folder / 'expenses.csv').open(newline='', encoding='utf-8-sig') as stream:
        reader = csv.DictReader(stream)
        if reader.fieldnames != FIELDS:
            raise ValueError('expenses.csv headers must match the generated template')
        rows = []
        seen = set()
        for line, raw in enumerate(reader, 2):
            if None in raw or any(v is None for v in raw.values()):
                raise ValueError(f'CSV row {line}: wrong number of columns')
            row = {k: v.strip() for k, v in raw.items()}
            if not any(row.values()):
                continue
            try:
                for key in ['expense_id', 'purchase_date', 'posted_date', 'merchant', 'location',
                            'category', 'description', 'business_purpose', 'posted_gbp', 'receipt_status']:
                    if not row[key]:
                        raise ValueError(f'{key} is required')
                ident = row['expense_id']
                if not re.fullmatch(r'[A-Za-z0-9_-]+', ident) or ident in seen:
                    raise ValueError('expense_id must be unique and use letters, digits, _ or -')
                seen.add(ident)
                for key in ['purchase_date', 'posted_date']:
                    date.fromisoformat(row[key])
                row['posted_gbp'] = format(money(row['posted_gbp'], 'posted_gbp'), '.2f')
                if bool(row['original_amount']) != bool(row['original_currency']):
                    raise ValueError('supply both original_amount and original_currency, or leave both blank')
                if row['original_amount']:
                    money(row['original_amount'], 'original_amount')
                    if not re.fullmatch('[A-Z]{3}', row['original_currency']):
                        raise ValueError('original_currency must be a three-letter uppercase code')
                if row['receipt_status'] not in STATUSES:
                    raise ValueError('receipt_status must be receipt, missing or requested')
                if row['receipt_status'] == 'receipt' and not row['receipt_file']:
                    raise ValueError('receipt_file is required for receipt status')
                if row['receipt_status'] != 'receipt' and (not row['missing_reason'] or not row['card_evidence_file']):
                    raise ValueError('missing/requested receipts require missing_reason and card_evidence_file')
                for key in ['receipt_file', 'card_evidence_file']:
                    if row[key]:
                        path = (folder / row[key]).resolve()
                        if not path.is_relative_to(folder.resolve()) or not path.is_file():
                            raise ValueError(f'{key} must be an existing file inside the trip folder')
                        if path.suffix.lower() not in {'.pdf', '.png', '.jpg', '.jpeg', '.webp'}:
                            raise ValueError(f'{key}: use PDF, PNG, JPG or WebP evidence')
                rows.append(row)
            except ValueError as error:
                raise ValueError(f'CSV row {line}: {error}') from None
    if not rows:
        raise ValueError('Add at least one expense to expenses.csv')
    return sorted(rows, key=lambda row: (row['purchase_date'], row['expense_id']))


def init_trip(folder):
    if folder.exists() and any(folder.iterdir()):
        raise ValueError('Choose a new or empty trip folder; existing files will not be overwritten')
    folder.mkdir(parents=True, exist_ok=True)
    (folder / 'evidence').mkdir()
    trip = dict(claimant='', company='', project='', destination='', start_date='', end_date='',
                business_purpose='', trip_evidence_files=[], declaration_date='', approver='')
    (folder / 'trip.json').write_text(json.dumps(trip, indent=2) + '\n', encoding='utf-8')
    with (folder / 'expenses.csv').open('w', newline='', encoding='utf-8') as stream:
        csv.writer(stream).writerow(FIELDS)
    print(f'Created {folder}. Fill trip.json and expenses.csv; put evidence in evidence/.')


def csv_safe(value):
    # Prevent merchant names / notes being interpreted as spreadsheet formulas.
    return "'" + value if value.lstrip().startswith(('=', '+', '-', '@')) else value


def generate(folder, output):
    trip = json.loads((folder / 'trip.json').read_text(encoding='utf-8'))
    for key in ['claimant', 'company', 'project', 'destination', 'start_date', 'end_date', 'business_purpose', 'declaration_date']:
        if not isinstance(trip.get(key), str) or not trip[key].strip():
            raise ValueError(f'trip.json: {key} is required')
    for key in ['start_date', 'end_date', 'declaration_date']:
        date.fromisoformat(trip[key])
    if trip['start_date'] > trip['end_date']:
        raise ValueError('Trip start_date must be on or before end_date')
    rows = load_rows(folder)
    for row in rows:
        if not trip['start_date'] <= row['purchase_date'] <= trip['end_date']:
            raise ValueError(f"{row['expense_id']}: purchase_date is outside the trip dates")
    files = {}
    for row in rows:
        for key in ['receipt_file', 'card_evidence_file']:
            if row[key]:
                files[row[key]] = (folder / row[key]).resolve()
    extras = trip.get('trip_evidence_files', [])
    if not isinstance(extras, list) or any(not isinstance(p, str) for p in extras):
        raise ValueError('trip_evidence_files must be a list of file paths')
    for name in extras:
        path = (folder / name).resolve()
        if not path.is_relative_to(folder.resolve()) or not path.is_file() or path.suffix.lower() not in {'.pdf', '.png', '.jpg', '.jpeg', '.webp'}:
            raise ValueError('Trip evidence must be PDF or image files inside the trip folder')
        files[name] = path
    output = output.resolve()
    # Keep generated material separate from all source records.
    if output.is_relative_to(folder.resolve()) or folder.resolve().is_relative_to(output):
        raise ValueError('Output must be separate from the trip folder')
    if output.exists():
        raise ValueError('Output already exists; choose a new output folder')
    output.mkdir(parents=True)
    (output / 'evidence').mkdir()
    links = {}
    for index, (name, source) in enumerate(files.items(), 1):
        target = f'evidence/{index:03d}{source.suffix.lower()}'
        shutil.copyfile(source, output / target)
        links[name] = target
    exported = []
    for row in rows:
        item = row.copy()
        for key in ['receipt_file', 'card_evidence_file']:
            item[key] = links.get(row[key], '')
        exported.append(item)
    with (output / 'expenses.csv').open('w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows({k: csv_safe(v) for k, v in row.items()} for row in exported)
    esc = lambda value: html.escape(str(value), quote=True)
    total = sum((Decimal(r['posted_gbp']) for r in rows), Decimal('0'))
    missing = [r for r in rows if r['receipt_status'] != 'receipt']
    sections = []
    for row in exported:
        details = ''.join(f'<dt>{esc(k.replace("_", " "))}</dt><dd>{esc(v)}</dd>' for k, v in row.items() if v and k not in {'receipt_file', 'card_evidence_file'})
        evidence = ''.join(f'<p><a href="{esc(row[k])}">{esc(k.replace("_", " "))}</a></p>' for k in ['receipt_file', 'card_evidence_file'] if row[k])
        sections.append(f'<section><h2 id="{esc(row["expense_id"])}">{esc(row["expense_id"])}</h2><dl>{details}</dl>{evidence}</section>')
    summary = ''.join(f'<tr><td><a href="#{esc(r["expense_id"])}">{esc(r["expense_id"])}</a></td><td>{esc(r["purchase_date"])}</td><td>{esc(r["merchant"])}</td><td>{esc(r["category"])}</td><td>{esc(r["receipt_status"])}</td><td>{esc(r["posted_gbp"])}</td></tr>' for r in rows)
    declaration = ''
    if missing:
        declaration = f'''<section><h2>Missing-receipt declaration — awaiting signature</h2>
<p>Affected entries: {esc(', '.join(r['expense_id'] for r in missing))}.</p>
<p>I confirm that these expenses were incurred by me for the business purposes stated, that the amounts claimed reflect my posted GBP card charges, and that I have not claimed them elsewhere. Itemized merchant receipts are unavailable or requested for the reasons recorded against each entry. Any reconstructed descriptions are identified in the notes.</p>
<p>Review this statement before signing. This is a claimant declaration, not a merchant receipt or assurance of tax deductibility.</p>
<p>Claimant: {esc(trip['claimant'])} · Declaration date: {esc(trip['declaration_date'])}</p><p>Signature: __________________________</p></section>'''
    metadata = ''.join(f'<dt>{esc(k.replace("_", " "))}</dt><dd>{esc(trip[k])}</dd>' for k in ['claimant', 'company', 'project', 'destination', 'start_date', 'end_date', 'business_purpose'])
    trip_links = ''.join(f'<li><a href="{esc(links[p])}">Trip evidence {i}</a></li>' for i, p in enumerate(extras, 1))
    report = f'''<!doctype html><html lang="en"><meta charset="utf-8"><title>Expense report</title>
<style>body{{font:16px system-ui;max-width:1000px;margin:40px auto;padding:0 20px;color:#182534}}table{{border-collapse:collapse;width:100%}}td,th{{padding:9px;border-bottom:1px solid #bbb;text-align:left}}dt{{font-weight:bold;margin-top:8px}}dd{{margin-left:0;white-space:pre-wrap;overflow-wrap:anywhere}}section{{border-top:2px solid #aaa;margin-top:30px;padding-top:12px}}@media print{{section{{break-before:page}}a{{color:inherit}}}}</style>
<h1>GBP expense reimbursement report</h1><dl>{metadata}</dl><ul>{trip_links}</ul>
<p>Reimbursement requested: <strong>GBP {total:.2f}</strong>. Missing/requested receipts: {len(missing)}.</p>
<table><thead><tr><th>Reference</th><th>Date</th><th>Merchant</th><th>Category</th><th>Receipt</th><th>GBP</th></tr></thead><tbody>{summary}</tbody></table>
{''.join(sections)}{declaration}
<section><h2>Company review — awaiting approval</h2><p>Approver: {esc(trip.get('approver', ''))}</p><p>Approved amount (GBP): __________ Date: __________ Signature/reference: __________</p><p>Payment date/reference: ____________________</p>
<p>Amounts are actual GBP reimbursement requests, not calculated tax deductions. Meals are separately categorized for the company's accountant. US sales tax is not Canadian GST/HST. No Canadian tax credit or CAD conversion is calculated here. Approval does not guarantee CRA acceptance.</p></section></html>'''
    (output / 'report.html').write_text(report, encoding='utf-8')
    (output / 'trip.json').write_text(json.dumps(trip, indent=2) + '\n', encoding='utf-8')
    print(f'Generated {output}: GBP {total:.2f}; {len(missing)} missing/requested receipts. Open report.html.')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    commands.add_parser('init', help='Create blank trip templates').add_argument('folder', type=Path)
    build = commands.add_parser('build', help='Validate and generate a self-contained report folder')
    build.add_argument('folder', type=Path)
    build.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    try:
        if args.command == 'init':
            init_trip(args.folder)
        else:
            generate(args.folder, args.output)
    except (ValueError, OSError, InvalidOperation) as error:
        parser.exit(2, f'Error: {error}\n')


if __name__ == '__main__':
    main()
