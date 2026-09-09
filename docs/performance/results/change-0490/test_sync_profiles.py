"""Reject ambiguous per-sample sync attribution before using profiler data."""
from pathlib import Path
import unittest

import sync_profiles as profiles


class SyncAlignmentTests(unittest.TestCase):
    def trace(self, count=5):
        return '\n'.join(f'1234 178896900{i}.000001 fdatasync(3</owned/replay/case.replay>) = 0 <0.000123>' for i in range(count))

    def test_one_preflight_and_warmups_are_excluded(self):
        value = profiles.parse_sync(self.trace(), Path('/owned/replay'), samples=2, warmups=2)
        self.assertEqual(len(value['events']), 5)
        self.assertEqual([x['timestamp'] for x in value['measured']], ['1788969003.000001', '1788969004.000001'])
        self.assertEqual([x['duration_ns'] for x in value['measured']], [123000, 123000])

    def test_missing_or_extra_sync_refuses_alignment(self):
        for count in [4, 6]:
            with self.subTest(count=count), self.assertRaises(ValueError):
                profiles.parse_sync(self.trace(count), Path('/owned/replay'), 2, 2)

    def test_foreign_path_and_changed_path_refused(self):
        for text in [self.trace().replace('/owned/', '/other/'), self.trace().replace('case.replay', 'second.replay', 1)]:
            with self.subTest(text=text), self.assertRaises(ValueError):
                profiles.parse_sync(text, Path('/owned/replay'), 2, 2)

    def test_failed_unfinished_and_wrong_kind_refused(self):
        for text in [self.trace().replace('= 0', '= -1 EIO', 1), self.trace() + '\nfdatasync(3 <unfinished ...>', self.trace().replace('fdatasync', 'fsync', 1)]:
            with self.subTest(text=text), self.assertRaises(ValueError):
                profiles.parse_sync(text, Path('/owned/replay'), 2, 2)

    def test_out_of_order_sync_refused(self):
        lines = self.trace().splitlines()
        lines[1], lines[2] = lines[2], lines[1]
        with self.assertRaises(ValueError):
            profiles.parse_sync('\n'.join(lines), Path('/owned/replay'), 2, 2)

    def test_phase_order_is_reversed_and_inventory_explicit(self):
        rows = profiles.runs()
        self.assertEqual(len(rows), 8)
        self.assertEqual(len({r['label'] for r in rows}), 8)
        self.assertEqual([r['phase'] for r in rows[:4]], ['before', 'after', 'after', 'before'])
        self.assertEqual([r['samples'] for r in rows], [30] * 4 + [1] * 4)

    def summary(self):
        names = ['fdatasync', 'openat', 'close', 'read', 'write', 'unlink', 'statx']
        return '% time seconds usecs/call calls errors syscall\n' + '\n'.join(
            f'1.00 0.000003 1 3 {name}' for name in names
        ) + '\n100.00 0.000021 1 21 total\n'

    def test_summary_requires_sync_and_file_lifecycle_counts(self):
        value = profiles.parse_summary(self.summary(), samples=1, warmups=1)
        self.assertEqual(value['syscalls']['fdatasync']['calls'], 3)
        self.assertEqual(value['total']['calls'], 21)

    def test_summary_rejects_empty_malformed_irrelevant_and_bad_totals(self):
        for text in ['', 'garbage', self.summary().replace('fdatasync', 'fsync'), self.summary().replace('21 total', '20 total'), self.summary().replace('statx', 'unknown'), self.summary().replace('3 fdatasync', '2 fdatasync')]:
            with self.subTest(text=text), self.assertRaises(ValueError):
                profiles.parse_summary(text, samples=1, warmups=1)


if __name__ == '__main__':
    unittest.main()
