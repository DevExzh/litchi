//! Allocation-free ODF 1.4 full-width/half-width text mappings.
//!
//! The scalar evaluator owns UTF-8 traversal, output reservation, and result
//! construction.  These functions only map one Unicode scalar at a time.  An
//! [`asc`] result may carry one half-width voiced mark; [`jis`] reports when a
//! following voiced mark was consumed as part of the current full-width
//! katakana character.

/// Apply ODF 1.4 §6.20.2 Table 33 (ASC) to one Unicode scalar.
///
/// The optional second character is the half-width voiced or semi-voiced mark
/// that follows a voiced katakana base in the ASC result.
pub(super) fn asc(c: char) -> (char, Option<char>) {
    let code = c as u32;
    match code {
        0x30a1..=0x30aa if code & 1 == 0 => (mapped((code - 0x30a2) / 2 + 0xff71), None),
        0x30a1..=0x30aa => (mapped((code - 0x30a1) / 2 + 0xff67), None),
        0x30ab..=0x30c2 if code & 1 == 0 => {
            (mapped((code - 0x30ac) / 2 + 0xff76), Some('\u{ff9e}'))
        },
        0x30ab..=0x30c2 => (mapped((code - 0x30ab) / 2 + 0xff76), None),
        0x30c3 => ('\u{ff6f}', None),
        0x30c4..=0x30c9 if code & 1 == 1 => {
            (mapped((code - 0x30c5) / 2 + 0xff82), Some('\u{ff9e}'))
        },
        0x30c4..=0x30c9 => (mapped((code - 0x30c4) / 2 + 0xff82), None),
        0x30ca..=0x30ce => (mapped(code - 0x30ca + 0xff85), None),
        0x30cf..=0x30dd if code % 3 == 1 => {
            (mapped((code - 0x30d0) / 3 + 0xff8a), Some('\u{ff9e}'))
        },
        0x30cf..=0x30dd if code % 3 == 2 => {
            (mapped((code - 0x30d1) / 3 + 0xff8a), Some('\u{ff9f}'))
        },
        0x30cf..=0x30dd => (mapped((code - 0x30cf) / 3 + 0xff8a), None),
        0x30de..=0x30e2 => (mapped(code - 0x30de + 0xff8f), None),
        0x30e3..=0x30e8 if code & 1 == 0 => (mapped((code - 0x30e4) / 2 + 0xff94), None),
        0x30e3..=0x30e8 => (mapped((code - 0x30e3) / 2 + 0xff6c), None),
        0x30e9..=0x30ed => (mapped(code - 0x30e9 + 0xff97), None),
        0x30ef => ('\u{ff9c}', None),
        0x30f2 => ('\u{ff66}', None),
        0x30f3 => ('\u{ff9d}', None),
        0xff01..=0xff5e => (mapped(code - 0xff01 + 0x21), None),
        0x2015 => ('\u{ff70}', None),
        0x2018 => ('\u{0060}', None),
        0x2019 => ('\u{0027}', None),
        0x201d => ('\u{0022}', None),
        0x3001 => ('\u{ff64}', None),
        0x3002 => ('\u{ff61}', None),
        0x300c => ('\u{ff62}', None),
        0x300d => ('\u{ff63}', None),
        0x309b => ('\u{ff9e}', None),
        0x309c => ('\u{ff9f}', None),
        0x30fb => ('\u{ff65}', None),
        0x30fc => ('\u{ff70}', None),
        0xffe5 => ('\u{005c}', None),
        _ => (c, None),
    }
}

/// Apply ODF 1.4 §6.20.11 Table 34 (JIS) to one Unicode scalar.
///
/// `next` is the immediately following scalar, when present.  The returned
/// Boolean is true only when that scalar is a half-width voiced or semi-voiced
/// mark consumed by the current katakana base.
pub(super) fn jis(c: char, next: Option<char>) -> (char, bool) {
    let code = c as u32;
    match code {
        0x22 => ('\u{201d}', false),
        0x5c => ('\u{ffe5}', false),
        0x60 => ('\u{2018}', false),
        0x27 => ('\u{2019}', false),
        0x21..=0x7e => (mapped(code - 0x21 + 0xff01), false),
        0xff66 => ('\u{30f2}', false),
        0xff67..=0xff6b => (mapped((code - 0xff67) * 2 + 0x30a1), false),
        0xff6c..=0xff6e => (mapped((code - 0xff6c) * 2 + 0x30e3), false),
        0xff6f => ('\u{30c3}', false),
        0xff71..=0xff75 => (mapped((code - 0xff71) * 2 + 0x30a2), false),
        0xff76..=0xff81 if next == Some('\u{ff9e}') => (mapped((code - 0xff76) * 2 + 0x30ac), true),
        0xff76..=0xff81 => (mapped((code - 0xff76) * 2 + 0x30ab), false),
        0xff82..=0xff84 if next == Some('\u{ff9e}') => (mapped((code - 0xff82) * 2 + 0x30c5), true),
        0xff82..=0xff84 => (mapped((code - 0xff82) * 2 + 0x30c4), false),
        0xff85..=0xff89 => (mapped(code - 0xff85 + 0x30ca), false),
        0xff8a..=0xff8e if next == Some('\u{ff9e}') => (mapped((code - 0xff8a) * 3 + 0x30d0), true),
        0xff8a..=0xff8e if next == Some('\u{ff9f}') => (mapped((code - 0xff8a) * 3 + 0x30d1), true),
        0xff8a..=0xff8e => (mapped((code - 0xff8a) * 3 + 0x30cf), false),
        0xff8f..=0xff93 => (mapped(code - 0xff8f + 0x30de), false),
        0xff94..=0xff96 => (mapped((code - 0xff94) * 2 + 0x30e4), false),
        0xff97..=0xff9b => (mapped(code - 0xff97 + 0x30e9), false),
        0xff9c => ('\u{30ef}', false),
        0xff9d => ('\u{30f3}', false),
        0xff9e => ('\u{309b}', false),
        0xff9f => ('\u{309c}', false),
        0xff70 => ('\u{30fc}', false),
        0xff61 => ('\u{3002}', false),
        0xff62 => ('\u{300c}', false),
        0xff63 => ('\u{300d}', false),
        0xff64 => ('\u{3001}', false),
        0xff65 => ('\u{30fb}', false),
        _ => (c, false),
    }
}

#[inline]
fn mapped(value: u32) -> char {
    // Every value produced by Tables 33 and 34 is a Unicode scalar value.
    char::from_u32(value).unwrap_or('\0')
}

#[cfg(test)]
mod tests {
    use super::{asc, jis};

    fn cp(value: u32) -> char {
        char::from_u32(value).expect("test code point is a scalar")
    }

    fn assert_asc_table(table: &[(u32, u32, Option<u32>)]) {
        for &(input, output, suffix) in table {
            assert_eq!(
                asc(cp(input)),
                (cp(output), suffix.map(cp)),
                "U+{input:04X}"
            );
        }
    }

    fn assert_asc_range(start: u32, outputs: &[u32]) {
        for (offset, &output) in outputs.iter().enumerate() {
            let input = start + u32::try_from(offset).expect("test range fits");
            assert_eq!(asc(cp(input)), (cp(output), None), "U+{input:04X}");
        }
    }

    #[test]
    fn asc_covers_every_named_katakana_range() {
        assert_asc_range(
            0x30a1,
            &[
                0xff67, 0xff71, 0xff68, 0xff72, 0xff69, 0xff73, 0xff6a, 0xff74, 0xff6b, 0xff75,
            ],
        );
        assert_asc_table(&[
            (0x30ab, 0xff76, None),
            (0x30ac, 0xff76, Some(0xff9e)),
            (0x30ad, 0xff77, None),
            (0x30ae, 0xff77, Some(0xff9e)),
            (0x30af, 0xff78, None),
            (0x30b0, 0xff78, Some(0xff9e)),
            (0x30b1, 0xff79, None),
            (0x30b2, 0xff79, Some(0xff9e)),
            (0x30b3, 0xff7a, None),
            (0x30b4, 0xff7a, Some(0xff9e)),
            (0x30b5, 0xff7b, None),
            (0x30b6, 0xff7b, Some(0xff9e)),
            (0x30b7, 0xff7c, None),
            (0x30b8, 0xff7c, Some(0xff9e)),
            (0x30b9, 0xff7d, None),
            (0x30ba, 0xff7d, Some(0xff9e)),
            (0x30bb, 0xff7e, None),
            (0x30bc, 0xff7e, Some(0xff9e)),
            (0x30bd, 0xff7f, None),
            (0x30be, 0xff7f, Some(0xff9e)),
            (0x30bf, 0xff80, None),
            (0x30c0, 0xff80, Some(0xff9e)),
            (0x30c1, 0xff81, None),
            (0x30c2, 0xff81, Some(0xff9e)),
        ]);
        assert_asc_table(&[
            (0x30c3, 0xff6f, None),
            (0x30c4, 0xff82, None),
            (0x30c5, 0xff82, Some(0xff9e)),
            (0x30c6, 0xff83, None),
            (0x30c7, 0xff83, Some(0xff9e)),
            (0x30c8, 0xff84, None),
            (0x30c9, 0xff84, Some(0xff9e)),
        ]);
        assert_asc_range(0x30ca, &[0xff85, 0xff86, 0xff87, 0xff88, 0xff89]);
        assert_asc_table(&[
            (0x30cf, 0xff8a, None),
            (0x30d0, 0xff8a, Some(0xff9e)),
            (0x30d1, 0xff8a, Some(0xff9f)),
            (0x30d2, 0xff8b, None),
            (0x30d3, 0xff8b, Some(0xff9e)),
            (0x30d4, 0xff8b, Some(0xff9f)),
            (0x30d5, 0xff8c, None),
            (0x30d6, 0xff8c, Some(0xff9e)),
            (0x30d7, 0xff8c, Some(0xff9f)),
            (0x30d8, 0xff8d, None),
            (0x30d9, 0xff8d, Some(0xff9e)),
            (0x30da, 0xff8d, Some(0xff9f)),
            (0x30db, 0xff8e, None),
            (0x30dc, 0xff8e, Some(0xff9e)),
            (0x30dd, 0xff8e, Some(0xff9f)),
        ]);
        assert_asc_range(0x30de, &[0xff8f, 0xff90, 0xff91, 0xff92, 0xff93]);
        assert_asc_table(&[
            (0x30e3, 0xff6c, None),
            (0x30e4, 0xff94, None),
            (0x30e5, 0xff6d, None),
            (0x30e6, 0xff95, None),
            (0x30e7, 0xff6e, None),
            (0x30e8, 0xff96, None),
        ]);
        assert_asc_range(0x30e9, &[0xff97, 0xff98, 0xff99, 0xff9a, 0xff9b]);
        assert_asc_table(&[
            (0x30ef, 0xff9c, None),
            (0x30f2, 0xff66, None),
            (0x30f3, 0xff9d, None),
        ]);
    }

    #[test]
    fn asc_covers_ascii_range_and_named_exceptions() {
        let ascii = "!\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~";
        for (offset, expected) in ascii.chars().enumerate() {
            let input = cp(0xff01 + u32::try_from(offset).expect("ASCII range fits"));
            assert_eq!(asc(input), (expected, None), "ASCII offset {offset}");
        }
        assert_asc_table(&[
            (0x2015, 0xff70, None),
            (0x2018, 0x0060, None),
            (0x2019, 0x0027, None),
            (0x201d, 0x0022, None),
            (0x3001, 0xff64, None),
            (0x3002, 0xff61, None),
            (0x300c, 0xff62, None),
            (0x300d, 0xff63, None),
            (0x309b, 0xff9e, None),
            (0x309c, 0xff9f, None),
            (0x30fb, 0xff65, None),
            (0x30fc, 0xff70, None),
            (0xffe5, 0x005c, None),
        ]);
        assert_eq!(asc('A'), ('A', None));
        assert_eq!(asc('\u{3042}'), ('\u{3042}', None));
    }

    fn assert_jis_table(table: &[(u32, Option<u32>, u32, bool)]) {
        for &(input, next, output, consume) in table {
            assert_eq!(
                jis(cp(input), next.map(cp)),
                (cp(output), consume),
                "U+{input:04X} next={next:?}"
            );
        }
    }

    fn assert_jis_range(start: u32, outputs: &[u32]) {
        for (offset, &output) in outputs.iter().enumerate() {
            let input = start + u32::try_from(offset).expect("test range fits");
            assert_eq!(jis(cp(input), None), (cp(output), false), "U+{input:04X}");
        }
    }

    #[test]
    fn jis_covers_ascii_exceptions_and_katakana_ranges() {
        assert_jis_table(&[
            (0x22, None, 0x201d, false),
            (0x5c, None, 0xffe5, false),
            (0x60, None, 0x2018, false),
            (0x27, None, 0x2019, false),
            (0xff66, None, 0x30f2, false),
            (0xff6f, None, 0x30c3, false),
            (0xff9c, None, 0x30ef, false),
            (0xff9d, None, 0x30f3, false),
            (0xff9e, None, 0x309b, false),
            (0xff9f, None, 0x309c, false),
            (0xff70, None, 0x30fc, false),
            (0xff61, None, 0x3002, false),
            (0xff62, None, 0x300c, false),
            (0xff63, None, 0x300d, false),
            (0xff64, None, 0x3001, false),
            (0xff65, None, 0x30fb, false),
        ]);
        let ascii = "!\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~";
        for (offset, input) in ascii.chars().enumerate() {
            let code = input as u32;
            let expected = match code {
                0x22 => 0x201d,
                0x27 => 0x2019,
                0x5c => 0xffe5,
                0x60 => 0x2018,
                _ => code - 0x21 + 0xff01,
            };
            assert_eq!(
                jis(input, None),
                (cp(expected), false),
                "ASCII offset {offset}"
            );
        }
        for (input, expected) in [
            (0xff66, 0x30f2),
            (0xff6f, 0x30c3),
            (0xff9c, 0x30ef),
            (0xff9d, 0x30f3),
        ] {
            assert_eq!(jis(cp(input), None), (cp(expected), false), "U+{input:04X}");
        }
        assert_jis_range(0xff67, &[0x30a1, 0x30a3, 0x30a5, 0x30a7, 0x30a9]);
        assert_jis_range(0xff6c, &[0x30e3, 0x30e5, 0x30e7]);
        assert_jis_range(0xff71, &[0x30a2, 0x30a4, 0x30a6, 0x30a8, 0x30aa]);
        assert_jis_range(
            0xff76,
            &[
                0x30ab, 0x30ad, 0x30af, 0x30b1, 0x30b3, 0x30b5, 0x30b7, 0x30b9, 0x30bb, 0x30bd,
                0x30bf, 0x30c1,
            ],
        );
        assert_jis_range(0xff82, &[0x30c4, 0x30c6, 0x30c8]);
        assert_jis_range(0xff85, &[0x30ca, 0x30cb, 0x30cc, 0x30cd, 0x30ce]);
        assert_jis_range(0xff8a, &[0x30cf, 0x30d2, 0x30d5, 0x30d8, 0x30db]);
        assert_jis_range(0xff8f, &[0x30de, 0x30df, 0x30e0, 0x30e1, 0x30e2]);
        assert_jis_range(0xff94, &[0x30e4, 0x30e6, 0x30e8]);
        assert_jis_range(0xff97, &[0x30e9, 0x30ea, 0x30eb, 0x30ec, 0x30ed]);
    }

    #[test]
    fn jis_consumes_only_matching_voiced_pairs() {
        assert_jis_table(&[
            (0xff76, Some(0xff9e), 0x30ac, true),
            (0xff81, Some(0xff9e), 0x30c2, true),
            (0xff76, Some(0xff9f), 0x30ab, false),
            (0xff82, Some(0xff9e), 0x30c5, true),
            (0xff84, Some(0xff9e), 0x30c9, true),
            (0xff82, Some(0xff9f), 0x30c4, false),
            (0xff8a, Some(0xff9e), 0x30d0, true),
            (0xff8e, Some(0xff9f), 0x30dd, true),
            (0xff8a, Some(0xff9f), 0x30d1, true),
            (0xff8e, Some(0xff9e), 0x30dc, true),
            (0xff8a, None, 0x30cf, false),
            (0xff8a, Some('X' as u32), 0x30cf, false),
        ]);
        assert_eq!(jis('\u{ff9e}', Some('\u{ff9f}')), ('\u{309b}', false));
    }

    #[test]
    fn width_mappings_leave_unlisted_scalars_unchanged() {
        assert_eq!(asc('A'), ('A', None));
        for c in ['\u{3042}', '\u{4e00}', '\u{1f600}'] {
            assert_eq!(asc(c), (c, None));
            assert_eq!(jis(c, Some('\u{ff9e}')), (c, false));
        }
    }
}
