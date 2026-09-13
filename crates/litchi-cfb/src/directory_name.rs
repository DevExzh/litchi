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
