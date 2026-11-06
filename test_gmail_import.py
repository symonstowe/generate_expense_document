import base64
import csv
from datetime import date, datetime, timezone
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock

from expense_report import FIELDS, generate, init_trip
from gmail_import import collect


def encoded(text):
    return base64.urlsafe_b64encode(text.encode()).decode().rstrip('=')


def response(value):
    request = Mock()
    request.execute.return_value = value
    return request


class GmailImport(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.trip = self.root / 'trip'
        init_trip(self.trip)
        self.api = Mock()
        self.api.users().getProfile.return_value = response({'emailAddress': 'test@example.com'})
        self.messages = self.api.users().messages()
        stamp = str(int(datetime(2026, 10, 2, tzinfo=timezone.utc).timestamp() * 1000))
        self.fixtures = {
            'a': {'id': 'a', 'internalDate': stamp, 'payload': {
                'headers': [{'name': 'subject', 'value': 'Company breakfast'}],
                'parts': [{'parts': [{'parts': [
                    {'mimeType': 'text/plain', 'body': {'data': encoded('Breakfast during site visit')}},
                    {'filename': '../../receipt.pdf', 'body': {'attachmentId': 'pdf1'}},
                    {'filename': 'receipt.pdf', 'body': {'data': encoded('second receipt')}},
                    {'filename': 'script.exe', 'body': {'data': encoded('unsupported')}}]}]}]}},
            'b': {'id': 'b', 'internalDate': stamp, 'payload': {
                'mimeType': 'text/plain', 'body': {'data': encoded('Lunch, no receipt')}}},
            'old': {'id': 'old', 'internalDate': '0', 'payload': {}},
        }
        self.messages.get.side_effect = lambda **kw: response(self.fixtures[kw['id']])
        self.messages.attachments().get.return_value = response({'data': encoded('first receipt')})
        self.pages()

    def pages(self):
        self.messages.list.side_effect = [response({'messages': [{'id': 'a'}], 'nextPageToken': 'next'}),
                                          response({'messages': [{'id': 'b'}, {'id': 'old'}]})]

    def collect(self):
        return collect(self.api, self.trip, 'label:Expenses OR subject:Company', date(2026, 10, 1), date(2026, 10, 3))

    def rows(self):
        with (self.trip / 'gmail_candidates.csv').open(encoding='utf-8-sig', newline='') as stream:
            return list(csv.DictReader(stream))

    def test_pagination_nested_attachments_and_review(self):
        self.assertEqual(self.collect(), 2)
        first_query = self.messages.list.call_args_list[0].kwargs['q']
        self.assertTrue(first_query.startswith('(label:Expenses OR subject:Company) after:'))
        self.assertEqual(self.messages.list.call_args_list[1].kwargs['pageToken'], 'next')
        rows = self.rows()
        self.assertIn('Breakfast during site visit', rows[0]['description'])
        self.assertEqual(rows[1]['description'], 'Lunch, no receipt')
        self.assertEqual(rows[0]['purchase_date'], '')
        self.assertEqual(rows[0]['posted_gbp'], '')
        self.assertEqual(rows[0]['receipt_status'], '')
        records = [json.loads(p.read_text()) for p in (self.trip / 'gmail-import').glob('*.json')]
        receipt = next(r for r in records if r['message_id'] == 'a')
        self.assertEqual(len(receipt['evidence_files']), 3)
        self.assertEqual(receipt['unsupported_attachments'], ['script.exe'])
        self.assertEqual(len(list((self.trip / 'evidence').iterdir())), 4)
        self.assertFalse((self.root / 'receipt.pdf').exists())
        # Reviewed row becomes a valid report expense, and its receipt is copied.
        config = dict(claimant='Test', company='Company', project='Visit', destination='US',
                      start_date='2026-10-01', end_date='2026-10-03', business_purpose='Meetings',
                      declaration_date='2026-11-06', trip_evidence_files=receipt['evidence_files'][1:])
        (self.trip / 'trip.json').write_text(json.dumps(config))
        row = rows[0]
        row.update(purchase_date='2026-10-02', posted_date='2026-10-03', merchant='Cafe',
                   location='Boston US', category='meal', business_purpose='Trip breakfast',
                   posted_gbp='10.00', receipt_status='receipt', receipt_file=receipt['evidence_files'][0])
        with (self.trip / 'expenses.csv').open('w', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=FIELDS)
            writer.writeheader()
            writer.writerow(row)
        generate(self.trip, self.root / 'report')
        self.assertIn('GBP 10.00', (self.root / 'report/report.html').read_text())
        self.assertEqual(len(list((self.root / 'report/evidence').iterdir())), 3)

    def test_rerun_preserves_review_edits_without_duplicates(self):
        self.collect()
        path = self.trip / 'gmail_candidates.csv'
        edited = path.read_text(encoding='utf-8-sig').replace('Lunch, no receipt', 'Reviewed lunch')
        path.write_text(edited, encoding='utf-8-sig')
        self.pages()
        self.assertEqual(self.collect(), 0)
        self.assertEqual(len(self.rows()), 2)
        self.assertEqual(self.rows()[1]['description'], 'Reviewed lunch')

    def test_api_failure_does_not_commit_partial_queue_and_can_retry(self):
        self.messages.list.side_effect = [response({'messages': [{'id': 'a'}], 'nextPageToken': 'next'}),
                                          RuntimeError('Connection failed')]
        with self.assertRaisesRegex(RuntimeError, 'Connection failed'):
            self.collect()
        self.assertFalse((self.trip / 'gmail_candidates.csv').exists())
        self.pages()
        self.assertEqual(self.collect(), 2)

    def test_html_only_email_is_text_not_executable(self):
        self.fixtures['b']['payload'] = {'mimeType': 'text/html', 'body': {
            'data': encoded('<p>Lunch &amp; coffee</p><script>secret()</script>')}}
        self.collect()
        self.assertIn('Lunch & coffee', self.rows()[1]['description'])
        self.assertNotIn('secret', self.rows()[1]['description'])
        email_files = list((self.trip / 'evidence').glob('*-email.txt'))
        self.assertTrue(any('Lunch & coffee' in p.read_text() for p in email_files))

    def test_empty_results(self):
        self.messages.list.side_effect = [response({})]
        self.assertEqual(self.collect(), 0)
        self.assertEqual(self.rows(), [])


if __name__ == '__main__':
    unittest.main()
