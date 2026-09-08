import re
import unittest
from heap_analyze import descriptor_pattern


class DescriptorFilterTests(unittest.TestCase):
    def test_factored_filter_matches_exact_identifier_set(self):
        selected = {0, 10, 0xab, 0xabc, *range(0x800, 0x980, 3)}
        pattern = re.compile(b'(?:' + descriptor_pattern(selected) + b')')
        for identifier in range(0x1000):
            self.assertEqual(pattern.fullmatch(f'{identifier:x}'.encode()) is not None,
                             identifier in selected)
        self.assertIsNone(pattern.fullmatch(b''))
        self.assertIsNone(pattern.fullmatch(b'00'))

    def test_selected_malformed_tail_is_not_silently_skipped(self):
        pattern = re.compile(rb'(?m)^([+-]) (' + descriptor_pattern([10, 0xabc]) + rb')((?: [^\n]*)?)\n')
        rows = list(pattern.finditer(b'+ ab\n+ a extra\n- abc\n+ abcd\n'))
        self.assertEqual([row.group(2) for row in rows], [b'a', b'abc'])
        self.assertEqual(rows[0].group(3), b' extra')


if __name__ == '__main__':
    unittest.main()
