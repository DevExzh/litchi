"""Focused regression for Rust slice symbols in Callgrind annotations."""
import unittest
from analyze_publication import display_name, h

class AnnotationNames(unittest.TestCase):
    def test_slice_type_brackets_are_not_object_suffix(self):
        name='core::slice::<impl [T]>::sort_unstable_by'
        self.assertEqual(display_name('???:'+name+' (4x) [/tmp/executable]'),name)
        parsed=h.parse_annotation('28 * ???:'+name+' [/tmp/executable]\n12 > ???:other (1x) [/tmp/executable]\n',name,'fixture')
        self.assertEqual(parsed['selected_ir'],28)
        self.assertEqual(parsed['direct'][0]['name'],'other')

    def test_regular_and_libc_names(self):
        self.assertEqual(display_name('???:module::function (2x) [/tmp/executable]'),'module::function')
        self.assertEqual(display_name('./source/file.S:__memcpy_avx (4x) [/lib/libc.so]'),'__memcpy_avx')

if __name__=='__main__':
    unittest.main()
