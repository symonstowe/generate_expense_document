import csv
import json
from pathlib import Path
import tempfile
import unittest
from expense_report import FIELDS, generate, init_trip


class Reports(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.trip = self.root / 'trip'
        init_trip(self.trip)
        config = dict(claimant='Example Person', company='Example Company', project='Site visit',
                      destination='Boston, US', start_date='2026-10-01', end_date='2026-10-03',
                      business_purpose='Project meetings', declaration_date='2026-11-06')
        (self.trip / 'trip.json').write_text(json.dumps(config))
        (self.trip / 'evidence/card.pdf').write_bytes(b'%PDF-1.4\n%%EOF')
        self.row = dict.fromkeys(FIELDS, '')
        self.row.update(expense_id='M01', purchase_date='2026-10-02', posted_date='2026-10-04',
                        merchant='=Example Cafe', location='Boston', category='meal', description='Lunch <solo>',
                        business_purpose='Lunch during site visit', posted_gbp='12.30', receipt_status='missing',
                        missing_reason='Forgot to request receipt', card_evidence_file='evidence/card.pdf')
        self.write_rows([self.row])

    def write_rows(self, rows):
        with (self.trip / 'expenses.csv').open('w', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=FIELDS)
            writer.writeheader()
            writer.writerows(rows)

    def test_report_and_portable_evidence(self):
        second = dict(self.row, expense_id='M02', posted_gbp='0.10')
        self.write_rows([self.row, second])
        out = self.root / 'result'
        generate(self.trip, out)
        report = (out / 'report.html').read_text()
        self.assertIn('GBP 12.40', report)
        self.assertIn('Lunch &lt;solo&gt;', report)
        self.assertIn('awaiting signature', report)
        with (out / 'expenses.csv').open(encoding='utf-8-sig') as stream:
            exported = list(csv.DictReader(stream))
        self.assertEqual(exported[0]['merchant'], "'=Example Cafe")
        self.assertEqual((out / exported[0]['card_evidence_file']).read_bytes(), b'%PDF-1.4\n%%EOF')
        self.assertEqual(len(list((out / 'evidence').iterdir())), 1)
        with self.assertRaises(ValueError):
            generate(self.trip, out)

    def test_reject_invalid_inputs(self):
        for changes in [dict(posted_gbp='NaN'), dict(posted_gbp='1.001'), dict(posted_gbp='-2'),
                        dict(missing_reason=''), dict(card_evidence_file=''),
                        dict(card_evidence_file='../outside.pdf'), dict(purchase_date='2025-01-01'),
                        dict(original_amount='12', original_currency='')]:
            with self.subTest(changes=changes):
                self.write_rows([dict(self.row, **changes)])
                with self.assertRaises(ValueError):
                    generate(self.trip, self.root / 'result')
                self.assertFalse((self.root / 'result').exists())

    def test_duplicate_ids(self):
        self.write_rows([self.row, self.row])
        with self.assertRaisesRegex(ValueError, 'unique'):
            generate(self.trip, self.root / 'result')

    def test_init_does_not_overwrite(self):
        with self.assertRaises(ValueError):
            init_trip(self.trip)


if __name__ == '__main__':
    unittest.main()
