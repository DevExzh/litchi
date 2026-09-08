import tempfile
from pathlib import Path
import unittest
from analyze import counters


class CounterTests(unittest.TestCase):
    def parse(self, text, expected=('cycles:u',)):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'counters.csv'
            path.write_text(text)
            return counters(path, expected)

    def test_scaled_reported_zero_and_unavailable_remain_distinct(self):
        rows = self.parse('0;;cycles:u;123;83.00;;\n<not supported>;;LLC-load-misses:u;0;100.00;;\n',
                          ('cycles:u', 'LLC-load-misses:u'))
        self.assertEqual(rows['cycles:u']['count'], 0)
        self.assertEqual(rows['cycles:u']['running_percent'], 83)
        self.assertIsNone(rows['LLC-load-misses:u']['count'])
        self.assertEqual(rows['LLC-load-misses:u']['status'], 'not_supported')

    def test_bad_counter_rows_rejected(self):
        for row in ('12;;cycles:u;1;101;;', '-1;;cycles:u;1;80;;',
                    '12;;cycles:u;1;nan;;', '12;;unknown;1;80;;',
                    '12;;cycles:u;1;80;;\n12;;cycles:u;1;80;;', ''):
            with self.subTest(row=row), self.assertRaises(AssertionError):
                self.parse(row)


if __name__ == '__main__':
    unittest.main()
