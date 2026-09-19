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

/// A validated ASCII query; stored directory keys keep their full UTF-16 form.
pub(crate) struct AsciiLookupKey<'a>(&'a [u8]);

impl AsciiLookupKey<'_> {
    #[inline]
    pub(crate) fn compare(&self, other: &DirectoryNameData) -> Ordering {
        let length = self.0.len().cmp(&other.utf16.len());
        if length != Ordering::Equal {
            return length;
        }
        for (&byte, &unit) in self.0.iter().zip(other.comparison.iter()) {
            let ordering = u16::from(byte.to_ascii_uppercase()).cmp(&unit);
            if ordering != Ordering::Equal {
                return ordering;
            }
        }
        self.0.len().cmp(&other.comparison.len())
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

/// Borrows bounded ASCII queries without constructing owned comparison keys.
/// `None` requires the caller to use `directory_name_data`, including its
/// validation; it does not mean the unhandled name is valid.
pub(crate) fn ascii_lookup_key(
    name: &str,
) -> Result<Option<AsciiLookupKey<'_>>, DirectoryNameError> {
    if name.len() > MAX_DIRECTORY_NAME_CODE_UNITS || !name.is_ascii() {
        return Ok(None);
    }
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
    Ok(Some(AsciiLookupKey(name.as_bytes())))
}

#[cfg(test)]
mod tests {
    use super::{
        DirectoryNameData, DirectoryNameError, ascii_lookup_key, directory_name_data,
        simple_uppercase,
    };

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

    fn assert_ascii_lookup_matches_reference(name: &str) {
        assert_matches_general_reference(name);

        let lookup = ascii_lookup_key(name);
        if name.is_ascii() && name.len() <= super::MAX_DIRECTORY_NAME_CODE_UNITS {
            match directory_name_data(name) {
                Ok(data) => {
                    let key = match lookup {
                        Ok(Some(key)) => key,
                        Ok(None) => panic!("bounded ASCII name did not produce a key"),
                        Err(error) => panic!("valid bounded ASCII name was rejected: {error:?}"),
                    };
                    assert_eq!(
                        key.compare(&data),
                        std::cmp::Ordering::Equal,
                        "borrowed key differs from owned key for {name:?}"
                    );
                },
                Err(expected) => match lookup {
                    Err(actual) => assert_eq!(actual, expected, "error for {name:?}"),
                    Ok(Some(_)) => panic!("invalid bounded ASCII name produced a key"),
                    Ok(None) => panic!("invalid bounded ASCII name fell back unexpectedly"),
                },
            }
        } else {
            assert!(
                matches!(lookup, Ok(None)),
                "non-bounded-ASCII name did not request the general fallback: {name:?}"
            );
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
            assert_ascii_lookup_matches_reference(&name);
        }
        for first in 0_u8..=0x7f {
            for second in 0_u8..=0x7f {
                let name = String::from_utf8(vec![first, second]).unwrap();
                assert_ascii_lookup_matches_reference(&name);
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
                assert_ascii_lookup_matches_reference(&name);
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
                    assert_ascii_lookup_matches_reference(&name);
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

    #[test]
    fn ascii_lookup_key_falls_back_for_unicode_and_long_names() {
        let mut names = vec![
            "é".to_owned(),
            "😀".to_owned(),
            "a/é".to_owned(),
            "\0é".to_owned(),
            format!("{}é", "a".repeat(30)),
            format!("{}😀", "a".repeat(29)),
            format!("{}😀", "a".repeat(30)),
            "é".repeat(32),
            "😀".repeat(16),
            "a".repeat(32),
            format!("{}!", "a".repeat(31)),
        ];

        for name in names.drain(..) {
            assert_ascii_lookup_matches_reference(&name);
        }
    }

    #[test]
    fn ascii_lookup_comparison_matches_owned_ordering_for_unicode_stored_names() {
        let cases = [
            ("z", "aa"),
            ("aa", "z"),
            ("workbook", "WORKBOOK"),
            ("A", "a"),
            ("a", "B"),
            ("A", "é"),
            ("z", "ß"),
            ("z", "😀"),
        ];

        for (lookup_name, stored_name) in cases {
            let lookup = match ascii_lookup_key(lookup_name) {
                Ok(Some(key)) => key,
                Ok(None) => panic!("ASCII lookup name unexpectedly fell back"),
                Err(error) => panic!("ASCII lookup name was rejected: {error:?}"),
            };
            let lookup_data = directory_name_data(lookup_name).unwrap();
            let stored_data = directory_name_data(stored_name).unwrap();
            assert_eq!(
                lookup.compare(&stored_data),
                lookup_data.compare(&stored_data),
                "ordering differs for lookup={lookup_name:?}, stored={stored_name:?}"
            );
        }

        // Preserve the final slice-length ordering even when test-only cached
        // comparison data differs in length from the original UTF-16 name.
        let key = ascii_lookup_key("ab").unwrap().unwrap();
        let owned = directory_name_data("ab").unwrap();
        for comparison in [vec![65], vec![65, 66, 67]] {
            let mut cached = directory_name_data("ab").unwrap();
            cached.comparison = comparison.into();
            assert_eq!(key.compare(&cached), owned.compare(&cached));
        }

        let long_ascii = "z".repeat(31);
        let long_unicode = format!("{}😀", "a".repeat(29));
        let lookup = match ascii_lookup_key(&long_ascii) {
            Ok(Some(key)) => key,
            Ok(None) => panic!("bounded ASCII lookup unexpectedly fell back"),
            Err(error) => panic!("bounded ASCII lookup was rejected: {error:?}"),
        };
        let lookup_data = directory_name_data(&long_ascii).unwrap();
        let stored_data = directory_name_data(&long_unicode).unwrap();
        assert_eq!(
            lookup.compare(&stored_data),
            lookup_data.compare(&stored_data),
            "ordering differs at the supplementary UTF-16 boundary"
        );
    }
}
