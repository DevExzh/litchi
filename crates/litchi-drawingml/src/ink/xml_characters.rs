//! XML 1.0 character checks on already validated UTF-8.

/// A Rust string cannot contain surrogates or scalars above U+10FFFF. The
/// remaining exclusions in XML 1.0's Char production are ASCII controls other
/// than TAB/LF/CR, and U+FFFE/U+FFFF. UTF-8 continuation bytes are all >= 0x80,
/// so inspecting control bytes cannot reject a valid multibyte character.
pub(super) fn valid(value: &str) -> bool {
    // A reduction lets the compiler check ASCII controls in blocks without
    // decoding every Unicode scalar a second time. Keep the exact non-ASCII
    // exclusions as UTF-8 substring searches; other noncharacters are legal
    // under the XML Char production and must not be rejected.
    let invalid_controls = value.as_bytes().iter().fold(0_u8, |invalid, &byte| {
        invalid | u8::from(byte < 0x20 && !matches!(byte, b'\t' | b'\n' | b'\r'))
    });
    invalid_controls == 0 && !value.contains("\u{fffe}") && !value.contains("\u{ffff}")
}

#[cfg(test)]
mod tests {
    use super::valid;

    #[test]
    fn matches_xml10_char_production_for_every_unicode_scalar() {
        let mut utf8 = [0; 4];
        for codepoint in 0..=0x10_ffff {
            let Some(character) = char::from_u32(codepoint) else {
                continue;
            };
            let expected = matches!(character,
                '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}'
                | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}');
            assert_eq!(
                valid(character.encode_utf8(&mut utf8)),
                expected,
                "U+{codepoint:04X}"
            );
        }
    }

    #[test]
    fn mixed_text_and_exclusions_at_block_boundaries_keep_the_same_contract() {
        let prefix = "handwriting 漢字 \u{10000}\u{10ffff}\t\n\r";
        assert!(valid(prefix));
        assert!(valid(""));
        for offset in 0..128 {
            let prefix = format!("{prefix}{}", "x".repeat(offset));
            for excluded in [
                '\0', '\u{8}', '\u{b}', '\u{c}', '\u{e}', '\u{1f}', '\u{fffe}', '\u{ffff}',
            ] {
                assert!(!valid(&format!("{prefix}{excluded}suffix")));
            }
        }
    }
}
