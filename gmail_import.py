"""Read-only Gmail collection into a review CSV; never guesses expense amounts."""
import argparse
import base64
import csv
from datetime import date, datetime, time, timedelta, timezone
import hashlib
from html.parser import HTMLParser
import json
import os
from pathlib import Path
import re
import tempfile

from expense_report import FIELDS, csv_safe

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
EXTENSIONS = {'.pdf', '.png', '.jpg', '.jpeg', '.webp'}


class EmailText(HTMLParser):
    """Extract readable text without rendering or loading email HTML."""
    def __init__(self):
        super().__init__()
        self.chunks = []
        self.hidden = 0

    def handle_starttag(self, tag, attrs):
        if tag in {'script', 'style'}:
            self.hidden += 1
        elif tag in {'p', 'br', 'div', 'tr', 'li'} and not self.hidden:
            self.chunks.append('\n')

    def handle_endtag(self, tag):
        if tag in {'script', 'style'}:
            self.hidden = max(0, self.hidden - 1)

    def handle_data(self, data):
        if not self.hidden:
            self.chunks.append(data)


def atomic_write(path, text):
    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=path.parent,
                                     delete=False, newline='') as stream:
        temp = Path(stream.name)
        try:
            stream.write(text)
        except BaseException:
            temp.unlink(missing_ok=True)
            raise
    try:
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def connect(credentials, token):
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError:
        raise ValueError('Install Gmail dependencies: python3 -m pip install -r requirements-gmail.txt') from None
    creds = Credentials.from_authorized_user_file(str(token), SCOPES) if token.exists() else None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials), SCOPES)
            creds = flow.run_local_server(port=0)
        token.parent.mkdir(parents=True, exist_ok=True)
        atomic_write(token, creds.to_json())
        token.chmod(0o600)
    return build('gmail', 'v1', credentials=creds, cache_discovery=False)


def walk_parts(part):
    yield part
    for child in part.get('parts', []):
        yield from walk_parts(child)


def decode(data):
    return base64.urlsafe_b64decode(data + '=' * (-len(data) % 4))


def body_data(service, message_id, part):
    body = part.get('body', {})
    if 'data' in body:
        return decode(body['data'])
    if 'attachmentId' in body:
        result = service.users().messages().attachments().get(
            userId='me', messageId=message_id, id=body['attachmentId']).execute(num_retries=3)
        return decode(result['data'])
    return b''


def collect(service, folder, query, start, end):
    """Inclusive UTC email-received dates, distinct from reviewed purchase dates."""
    if end < start:
        raise ValueError('--end must be on or after --start')
    if not query.strip():
        raise ValueError('Supply a Gmail query, ideally label:Expenses or a company subject filter')
    folder = folder.resolve()
    if not (folder / 'trip.json').is_file() or not (folder / 'expenses.csv').is_file():
        raise ValueError('Initialize a trip folder first with expense_report.py init')
    account = service.users().getProfile(userId='me').execute(num_retries=3)['emailAddress'].lower()
    since = int(datetime.combine(start, time.min, timezone.utc).timestamp())
    until = int(datetime.combine(end + timedelta(days=1), time.min, timezone.utc).timestamp())
    # Group the caller's query so OR cannot bypass the date boundaries.
    search = f'({query}) after:{since - 1} before:{until}'
    imported = folder / 'gmail-import'
    imported.mkdir(exist_ok=True)
    if imported.is_symlink():
        raise ValueError('gmail-import must not be a symlink')
    evidence = folder / 'evidence'
    evidence.mkdir(exist_ok=True)
    if evidence.is_symlink():
        raise ValueError('evidence must not be a symlink')
    candidates = folder / 'gmail_candidates.csv'
    existing = []
    if candidates.exists():
        with candidates.open(encoding='utf-8-sig', newline='') as stream:
            reader = csv.DictReader(stream)
            if reader.fieldnames != FIELDS:
                raise ValueError('gmail_candidates.csv has unexpected headers; restore the original headers')
            existing = list(reader)
            if any(None in r or any(v is None for v in r.values()) for r in existing):
                raise ValueError('gmail_candidates.csv has malformed rows')
    known = {r['expense_id'] for r in existing}
    additions = []
    token = None
    visited = set()
    while True:
        kwargs = dict(userId='me', q=search, maxResults=100)
        if token:
            kwargs['pageToken'] = token
        page = service.users().messages().list(**kwargs).execute(num_retries=3)
        for item in page.get('messages', []):
            message_id = item['id']
            ident = 'G' + hashlib.sha256(f'{account}:{message_id}'.encode()).hexdigest()[:20]
            if ident in known:
                continue
            message = service.users().messages().get(userId='me', id=message_id, format='full').execute(num_retries=3)
            received = datetime.fromtimestamp(int(message['internalDate']) / 1000, timezone.utc)
            if not start <= received.date() <= end:
                continue
            payload = message.get('payload', {})
            headers = {h['name'].lower(): h['value'] for h in payload.get('headers', [])}
            paths = []
            unsupported = []
            texts = []
            html_texts = []
            for index, part in enumerate(walk_parts(payload)):
                filename = part.get('filename', '')
                if filename:
                    extension = Path(filename.replace('\\', '/')).suffix.lower()
                    if extension not in EXTENSIONS:
                        unsupported.append(filename)
                        continue
                    # Never use the sender's filename as a filesystem path.
                    relative = f'evidence/{ident}-{index:03d}{extension}'
                    target = folder / relative
                    content = body_data(service, message_id, part)
                    if not content:
                        unsupported.append(f'{filename} (empty)')
                        continue
                    if target.exists():
                        if target.is_symlink() or target.read_bytes() != content:
                            raise ValueError(f'Evidence file conflict: {relative}')
                    else:
                        with target.open('xb') as stream:
                            stream.write(content)
                    paths.append(relative)
                elif part.get('mimeType') in {'text/plain', 'text/html'}:
                    # Decode MIME charset if supplied; replace invalid text, never execute it.
                    part_headers = {h['name'].lower(): h['value'] for h in part.get('headers', [])}
                    match = re.search(r'charset=["\']?([^;"\'\s]+)', part_headers.get('content-type', ''), re.I)
                    charset = match.group(1) if match else 'utf-8'
                    try:
                        decoded = body_data(service, message_id, part).decode(charset, errors='replace')
                    except LookupError:
                        decoded = body_data(service, message_id, part).decode('utf-8', errors='replace')
                    if part.get('mimeType') == 'text/html':
                        parser = EmailText()
                        parser.feed(decoded)
                        html_texts.append(''.join(parser.chunks))
                    else:
                        texts.append(decoded)
            text = '\n\n'.join(texts if texts else html_texts)
            record = dict(expense_id=ident, account=account, message_id=message_id,
                          received_utc=received.isoformat(), subject=headers.get('subject', ''),
                          sender=headers.get('from', ''), body=text, snippet=message.get('snippet', ''),
                          evidence_files=paths, unsupported_attachments=unsupported,
                          query=search, gmail_url=f'https://mail.google.com/mail/u/?authuser={account}#all/{message_id}')
            email_path = f'evidence/{ident}-email.txt'
            email_text = (f"From: {record['sender']}\nSubject: {record['subject']}\n"
                          f"Received UTC: {record['received_utc']}\nGmail message: {message_id}\n\n{text or record['snippet']}\n")
            target = folder / email_path
            if target.is_symlink():
                raise ValueError('Email evidence must not be a symlink')
            atomic_write(target, email_text)
            paths.append(email_path)
            atomic_write(imported / f'{ident}.json', json.dumps(record, indent=2, ensure_ascii=False) + '\n')
            row = dict.fromkeys(FIELDS, '')
            # Leave purchase/posting dates, merchant, GBP amount and receipt status for review.
            row.update(expense_id=ident, description=text.strip() or message.get('snippet', ''),
                       notes=f"REVIEW REQUIRED. Email received {received.date()} UTC; this is not the purchase date. "
                             f"Subject: {headers.get('subject', '')}. Source: gmail-import/{ident}.json. "
                             f"Evidence candidates (verify these are receipts, not logos): {', '.join(paths) or 'none'}. "
                             f"Unsupported/empty attachments: {', '.join(unsupported) or 'none'}. "
                             "One email may contain several purchases; split into unique expense rows if necessary.")
            additions.append({k: csv_safe(v) for k, v in row.items()})
            known.add(ident)
        token = page.get('nextPageToken')
        if not token:
            break
        if token in visited:
            raise ValueError('Gmail repeated a pagination token; retry the import')
        visited.add(token)
    # Commit the review queue only after all API requests complete. Source/evidence files
    # left by interrupted imports are reused after verifying their bytes on the next run.
    import io
    buffer = io.StringIO(newline='')
    writer = csv.DictWriter(buffer, fieldnames=FIELDS)
    writer.writeheader()
    writer.writerows(existing + additions)
    atomic_write(candidates, '\ufeff' + buffer.getvalue())
    print(f'Imported {len(additions)} new emails into {candidates}. Existing review rows preserved; expenses.csv unchanged.')
    return len(additions)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('folder', type=Path)
    parser.add_argument('--query', required=True, help='Gmail search, e.g. label:Expenses or subject:Tidal')
    parser.add_argument('--start', required=True, type=date.fromisoformat, help='First received date (UTC, inclusive)')
    parser.add_argument('--end', required=True, type=date.fromisoformat, help='Last received date (UTC, inclusive)')
    parser.add_argument('--credentials', type=Path, default=Path('credentials.json'))
    parser.add_argument('--token', type=Path, default=Path('token.json'))
    args = parser.parse_args()
    try:
        if args.end < args.start:
            raise ValueError('--end must be on or after --start')
        collect(connect(args.credentials, args.token), args.folder, args.query, args.start, args.end)
    except Exception as error:
        # API/auth/network errors must fail visibly, not silently create an incomplete queue.
        parser.exit(2, f'Gmail import failed: {error}\n')


if __name__ == '__main__':
    main()
