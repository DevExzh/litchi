use smallvec::SmallVec;
use std::cmp::Ordering;
use std::fmt;

pub(crate) const MAX_DIRECTORY_NAME_CODE_UNITS: usize = 31;
const FORBIDDEN_DIRECTORY_NAME_CHARS: [char; 4] = ['/', '\\', ':', '!'];

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DirectoryNameData {
    pub(crate) utf16: SmallVec<[u16; 32]>,
    pub(crate) comparison: SmallVec<[u16; 32]>,
}

impl DirectoryNameData {
    pub(crate) fn compare(&self, other: &Self) -> Ordering {
        self.utf16
            .len()
            .cmp(&other.utf16.len())
            .then_with(|| self.comparison.cmp(&other.comparison))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum DirectoryNameError {
    Empty,
    ContainsNul,
    ForbiddenCharacter(char),
    TooLong(usize),
}

impl fmt::Display for DirectoryNameError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => formatter.write_str("CFB directory entry names must not be empty"),
            Self::ContainsNul => {
                formatter.write_str("CFB directory entry names must not contain NUL")
            },
            Self::ForbiddenCharacter(character) => write!(
                formatter,
                "CFB directory entry name contains forbidden character {character:?}"
            ),
            Self::TooLong(length) => write!(
                formatter,
                "CFB directory entry name uses {length} UTF-16 code units; maximum is {MAX_DIRECTORY_NAME_CODE_UNITS}"
            ),
        }
    }
}

fn simple_uppercase(character: char) -> char {
    // No ASCII scalar has a multi-character uppercase mapping, and its
    // single-character mapping is the ASCII one, so the ASCII branch returns
    // exactly what the general path returns. CFB stream names are ASCII in
    // practice, and the general path builds a Unicode case-mapping iterator
    // for every character. `uppercase_matches_unicode_mapping_for_every_char`
    // proves the two agree over the whole scalar range.
    if character.is_ascii() {
        return character.to_ascii_uppercase();
    }
    let mut uppercase = character.to_uppercase();
    let first = uppercase.next().unwrap_or(character);
    if uppercase.next().is_some() {
        character
    } else {
        first
    }
}

pub(crate) fn directory_name_data(name: &str) -> Result<DirectoryNameData, DirectoryNameError> {
    if name.is_empty() {
        return Err(DirectoryNameError::Empty);
    }
    if name.contains('\0') {
        return Err(DirectoryNameError::ContainsNul);
    }
    if let Some(character) = name
        .chars()
        .find(|character| FORBIDDEN_DIRECTORY_NAME_CHARS.contains(character))
    {
        return Err(DirectoryNameError::ForbiddenCharacter(character));
    }

    if name.len() <= MAX_DIRECTORY_NAME_CODE_UNITS && name.is_ascii() {
        let mut utf16 = [0u16; 32];
        let mut comparison = [0u16; 32];
        for (index, &byte) in name.as_bytes().iter().enumerate() {
            utf16[index] = u16::from(byte);
            comparison[index] = u16::from(byte.to_ascii_uppercase());
        }
        return Ok(DirectoryNameData {
            utf16: SmallVec::from_buf_and_len(utf16, name.len()),
            comparison: SmallVec::from_buf_and_len(comparison, name.len()),
        });
    }

    let utf16: SmallVec<[u16; 32]> = name.encode_utf16().collect();
    if utf16.len() > MAX_DIRECTORY_NAME_CODE_UNITS {
        return Err(DirectoryNameError::TooLong(utf16.len()));
    }

    let mut comparison = SmallVec::with_capacity(utf16.len());
    for character in name.chars().map(simple_uppercase) {
        let mut encoded = [0u16; 2];
        comparison.extend_from_slice(character.encode_utf16(&mut encoded));
    }
    Ok(DirectoryNameData { utf16, comparison })
}

#[cfg(test)]
mod tests {
    use super::{DirectoryNameData, DirectoryNameError, directory_name_data, simple_uppercase};

    /// The pre-optimization construction expressed without any ASCII
    /// specialization.  Differential tests below keep this small reference
    /// independent of the production fast path while still using the CFB
    /// rule for single-code-point uppercase mappings.
    fn general_reference(name: &str) -> Result<(Vec<u16>, Vec<u16>), DirectoryNameError> {
        if name.is_empty() {
            return Err(DirectoryNameError::Empty);
        }
        if name.contains('\0') {
            return Err(DirectoryNameError::ContainsNul);
        }
        if let Some(character) = name
            .chars()
            .find(|character| ['/', '\\', ':', '!'].contains(character))
        {
            return Err(DirectoryNameError::ForbiddenCharacter(character));
        }

        let utf16: Vec<u16> = name.encode_utf16().collect();
        if utf16.len() > super::MAX_DIRECTORY_NAME_CODE_UNITS {
            return Err(DirectoryNameError::TooLong(utf16.len()));
        }

        let mut comparison = Vec::with_capacity(utf16.len());
        for character in name.chars().map(unicode_simple_uppercase) {
            let mut encoded = [0u16; 2];
            comparison.extend_from_slice(character.encode_utf16(&mut encoded));
        }
        Ok((utf16, comparison))
    }

    fn assert_matches_general_reference(name: &str) {
        let actual = directory_name_data(name);
        let expected = general_reference(name);
        match (actual, expected) {
            (Ok(actual), Ok((expected_utf16, expected_comparison))) => {
                assert_eq!(
                    actual.utf16.as_slice(),
                    expected_utf16.as_slice(),
                    "UTF-16 for {name:?}"
                );
                assert_eq!(
                    actual.comparison.as_slice(),
                    expected_comparison.as_slice(),
                    "comparison key for {name:?}"
                );
            },
            (Err(actual), Err(expected)) => assert_eq!(actual, expected, "error for {name:?}"),
            (actual, expected) => panic!(
                "fast path and general reference disagree for {name:?}: actual={actual:?}, expected={expected:?}"
            ),
        }
    }

    fn unicode_simple_uppercase(character: char) -> char {
        let mut uppercase = character.to_uppercase();
        let first = uppercase.next().unwrap_or(character);
        if uppercase.next().is_some() {
            character
        } else {
            first
        }
    }

    #[test]
    fn uppercase_matches_unicode_mapping_for_every_char() {
        // Exhaustive over the whole Unicode scalar range, so the ASCII branch
        // is proven equal rather than sampled.
        for scalar in 0..=u32::from(char::MAX) {
            let Some(character) = char::from_u32(scalar) else {
                continue;
            };
            assert_eq!(
                simple_uppercase(character),
                unicode_simple_uppercase(character),
                "uppercase mapping diverges for U+{scalar:04X}"
            );
        }
    }

    #[test]
    fn comparison_keys_are_unchanged_for_representative_names() {
        for name in [
            "Workbook",
            "WordDocument",
            "PowerPoint Document",
            "Root Entry",
            "\u{5}SummaryInformation",
            "benchmark_stream_00002.bin",
            "mixedCASE-123",
            "\u{df}stra\u{df}e",
            "\u{fb01}le",
            "\u{130}stanbul",
            "\u{3c3}\u{3c2}\u{3a3}",
        ] {
            let data = directory_name_data(name).expect("representative name is valid");
            let expected: Vec<u16> = name
                .chars()
                .map(unicode_simple_uppercase)
                .flat_map(|character| {
                    let mut encoded = [0u16; 2];
                    character.encode_utf16(&mut encoded).to_vec()
                })
                .collect();
            assert_eq!(
                data.comparison.as_slice(),
                expected.as_slice(),
                "comparison key changed for {name:?}"
            );
            assert_eq!(
                data.utf16.as_slice(),
                name.encode_utf16().collect::<Vec<u16>>().as_slice()
            );
        }
    }

    #[test]
    fn ascii_construction_matches_the_general_reference_for_every_byte_and_pair() {
        // Every one-byte and two-byte ASCII name is small enough to keep this
        // exhaustive while covering NUL, all four forbidden bytes, controls,
        // DEL, and every ordinary byte at both positions.
        for byte in 0_u8..=0x7f {
            let name = String::from_utf8(vec![byte]).unwrap();
            assert_matches_general_reference(&name);
        }
        for first in 0_u8..=0x7f {
            for second in 0_u8..=0x7f {
                let name = String::from_utf8(vec![first, second]).unwrap();
                assert_matches_general_reference(&name);
            }
        }
    }

    #[test]
    fn ascii_length_boundaries_match_the_general_reference() {
        // Include every length through two units past the limit so that both
        // the fixed-array admission boundary and the general too-long path
        // are exercised for every ASCII byte.
        for length in 0..=super::MAX_DIRECTORY_NAME_CODE_UNITS + 2 {
            for byte in 0_u8..=0x7f {
                let name = String::from_utf8(vec![byte; length]).unwrap();
                assert_matches_general_reference(&name);
            }
        }

        // Exercise every byte at every position around the fixed-array
        // boundary, including a late NUL or forbidden character that must be
        // reported before an over-length refusal.
        for length in [3, 30, 31, 32, 33] {
            for position in 0..length {
                for byte in 0_u8..=0x7f {
                    let mut bytes = vec![b'a'; length];
                    bytes[position] = byte;
                    let name = String::from_utf8(bytes).unwrap();
                    assert_matches_general_reference(&name);
                }
            }
        }
    }

    #[test]
    fn validation_precedence_matches_the_general_reference() {
        let mut cases = vec![
            (String::new(), DirectoryNameError::Empty),
            ("\0/".to_owned(), DirectoryNameError::ContainsNul),
            ("/\0".to_owned(), DirectoryNameError::ContainsNul),
            (
                "a!b/c".to_owned(),
                DirectoryNameError::ForbiddenCharacter('!'),
            ),
            (
                "a:/\\b".to_owned(),
                DirectoryNameError::ForbiddenCharacter(':'),
            ),
            (
                format!("/{}", "a".repeat(super::MAX_DIRECTORY_NAME_CODE_UNITS)),
                DirectoryNameError::ForbiddenCharacter('/'),
            ),
            (
                format!("{}:", "a".repeat(super::MAX_DIRECTORY_NAME_CODE_UNITS)),
                DirectoryNameError::ForbiddenCharacter(':'),
            ),
            (
                format!("\0{}", "a".repeat(super::MAX_DIRECTORY_NAME_CODE_UNITS)),
                DirectoryNameError::ContainsNul,
            ),
            (
                "a".repeat(super::MAX_DIRECTORY_NAME_CODE_UNITS + 1),
                DirectoryNameError::TooLong(super::MAX_DIRECTORY_NAME_CODE_UNITS + 1),
            ),
        ];
        for (name, expected) in cases.drain(..) {
            assert_eq!(
                directory_name_data(&name),
                Err(expected.clone()),
                "name={name:?}"
            );
            assert_eq!(
                general_reference(&name),
                Err(expected),
                "reference name={name:?}"
            );
        }
    }

    #[test]
    fn unicode_and_utf16_length_boundaries_match_the_general_reference() {
        let mut names = vec![
            "é".to_owned(),
            "😀".to_owned(),
            "ßstraße".to_owned(),
            "ﬃle".to_owned(),
            "ſtream".to_owned(),
            "İstanbul".to_owned(),
            "σςΣ".to_owned(),
            "雪だるま".to_owned(),
            "a/é".to_owned(),
            "😀\\雪".to_owned(),
        ];

        names.push(format!("{}é", "a".repeat(30)));
        names.push(format!("{}é", "a".repeat(31)));
        names.push(format!("{}😀", "a".repeat(29)));
        names.push(format!("{}😀", "a".repeat(30)));
        names.push(format!("{}雪", "a".repeat(30)));
        names.push(format!("{}雪", "a".repeat(31)));

        for name in names {
            assert_matches_general_reference(&name);
        }
    }

    #[test]
    fn refusals_are_unchanged() {
        assert_eq!(directory_name_data(""), Err(DirectoryNameError::Empty));
        assert_eq!(
            directory_name_data("a\0b"),
            Err(DirectoryNameError::ContainsNul)
        );
        assert_eq!(
            directory_name_data("a/b"),
            Err(DirectoryNameError::ForbiddenCharacter('/'))
        );
        assert_eq!(
            directory_name_data(&"a".repeat(32)),
            Err(DirectoryNameError::TooLong(32))
        );
    }

    #[test]
    fn ascii_case_insensitive_names_compare_equal() {
        let lower: DirectoryNameData = directory_name_data("workbook").unwrap();
        let upper = directory_name_data("WORKBOOK").unwrap();
        assert_eq!(lower.compare(&upper), std::cmp::Ordering::Equal);
    }
}
