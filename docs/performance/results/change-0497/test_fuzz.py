"""Exercise immutable fuzz-verification receipts without builds or fuzz runs."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

spec = importlib.util.spec_from_file_location('fuzz0497', Path(__file__).with_name('fuzz.py'))
fuzz = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fuzz)


class VerificationReceiptTests(unittest.TestCase):
    def payload(self):
        return {'schema': 'test', 'verified_utc': '2026-09-10T00:00:00+00:00',
                'passed': True, 'binary': {'sha256': 'a' * 64}}

    def test_reverification_is_read_only(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'verify.json'
            payload = self.payload()
            self.assertFalse(fuzz.write_or_check_verification(path, payload))
            before = path.read_bytes(), path.stat().st_mtime_ns
            payload['verified_utc'] = '2026-09-10T01:00:00+00:00'
            self.assertTrue(fuzz.write_or_check_verification(path, payload))
            self.assertEqual(before, (path.read_bytes(), path.stat().st_mtime_ns))

    def test_changed_binary_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'verify.json'
            payload = self.payload()
            fuzz.write_or_check_verification(path, payload)
            payload['binary']['sha256'] = 'b' * 64
            with self.assertRaises(RuntimeError):
                fuzz.write_or_check_verification(path, payload)

    def test_naive_retained_timestamp_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'verify.json'
            payload = self.payload()
            invalid = dict(payload, verified_utc='2026-09-10T00:00:00')
            path.write_text(json.dumps(invalid))
            with self.assertRaises(RuntimeError):
                fuzz.write_or_check_verification(path, payload)


if __name__ == '__main__':
    unittest.main()
