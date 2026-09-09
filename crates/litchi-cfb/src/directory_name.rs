use crate::file::OleError;
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
                "CFB directory entry name uses at least {length} UTF-16 code units; maximum is {MAX_DIRECTORY_NAME_CODE_UNITS}"
            ),
        }
    }
}

fn simple_uppercase(character: char) -> char {
    // [MS-CFB] 2.6.4 compares UTF-16 code points, so a supplementary scalar
    // is represented by two surrogate code points and neither code point is
    // uppercased. Rust's scalar `to_uppercase` would otherwise map scripts
    // such as Deseret.
    if character as u32 > 0xFFFF {
        return character;
    }

    // Rust exposes Unicode's full uppercase mapping. These BMP characters
    // have a distinct one-code-point simple uppercase mapping even though
    // their full mapping expands with an iota; all other multi-code-point
    // mappings intentionally remain unchanged.
    match character {
        '\u{1F80}' => '\u{1F88}',
        '\u{1F81}' => '\u{1F89}',
        '\u{1F82}' => '\u{1F8A}',
        '\u{1F83}' => '\u{1F8B}',
        '\u{1F84}' => '\u{1F8C}',
        '\u{1F85}' => '\u{1F8D}',
        '\u{1F86}' => '\u{1F8E}',
        '\u{1F87}' => '\u{1F8F}',
        '\u{1F90}' => '\u{1F98}',
        '\u{1F91}' => '\u{1F99}',
        '\u{1F92}' => '\u{1F9A}',
        '\u{1F93}' => '\u{1F9B}',
        '\u{1F94}' => '\u{1F9C}',
        '\u{1F95}' => '\u{1F9D}',
        '\u{1F96}' => '\u{1F9E}',
        '\u{1F97}' => '\u{1F9F}',
        '\u{1FA0}' => '\u{1FA8}',
        '\u{1FA1}' => '\u{1FA9}',
        '\u{1FA2}' => '\u{1FAA}',
        '\u{1FA3}' => '\u{1FAB}',
        '\u{1FA4}' => '\u{1FAC}',
        '\u{1FA5}' => '\u{1FAD}',
        '\u{1FA6}' => '\u{1FAE}',
        '\u{1FA7}' => '\u{1FAF}',
        '\u{1FB3}' => '\u{1FBC}',
        '\u{1FC3}' => '\u{1FCC}',
        '\u{1FF3}' => '\u{1FFC}',
        _ => {
            let mut uppercase = character.to_uppercase();
            let first = uppercase.next().unwrap_or(character);
            if uppercase.next().is_some() {
                character
            } else {
                first
            }
        },
    }
}

pub(crate) fn directory_name_data(name: &str) -> Result<DirectoryNameData, DirectoryNameError> {
    if name.is_empty() {
        return Err(DirectoryNameError::Empty);
    }
    let utf16_len = name
        .encode_utf16()
        .take(MAX_DIRECTORY_NAME_CODE_UNITS + 1)
        .count();
    if utf16_len > MAX_DIRECTORY_NAME_CODE_UNITS {
        return Err(DirectoryNameError::TooLong(utf16_len));
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

    let mut comparison = SmallVec::with_capacity(utf16.len());
    for character in name.chars().map(simple_uppercase) {
        let mut encoded = [0u16; 2];
        comparison.extend_from_slice(character.encode_utf16(&mut encoded));
    }
    Ok(DirectoryNameData { utf16, comparison })
}

/// Compare two CFB storage or stream names using the format's sorting
/// relationship. Invalid names never compare equal.
pub fn directory_names_equal(left: &str, right: &str) -> bool {
    match (directory_name_data(left), directory_name_data(right)) {
        (Ok(left), Ok(right)) => left.compare(&right) == Ordering::Equal,
        _ => false,
    }
}

/// Validate one CFB storage or stream name before authoring it.
///
/// The name must be nonempty, contain no NUL or `/`, `\\`, `:`, or `!`, and
/// fit within the 31 UTF-16 code-unit payload allowed by a 64-byte directory
/// name field after its terminating NUL is written.
pub fn validate_directory_name(name: &str) -> Result<(), OleError> {
    directory_name_data(name)
        .map(|_| ())
        .map_err(|error| OleError::InvalidData(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    const TOO_LONG_CODE_UNITS: usize = MAX_DIRECTORY_NAME_CODE_UNITS + 1;

    #[test]
    fn equality_uses_cfb_simple_uppercase_mapping() {
        assert!(directory_names_equal("é", "É"));
        assert!(directory_names_equal("ſtream", "STREAM"));
        assert!(!directory_names_equal("ß", "SS"));
        assert!(directory_names_equal("\u{1F80}", "\u{1F88}"));
        assert!(directory_names_equal("\u{1FB3}", "\u{1FBC}"));
        assert!(!directory_names_equal("ﬀ", "FF"));
    }

    #[test]
    fn supplementary_scalars_remain_utf16_surrogate_pairs() {
        assert!(!directory_names_equal("𐐨", "𐐀"));
    }

    #[test]
    fn invalid_names_never_compare_equal() {
        for name in [
            "",
            "nul\0name",
            "bad/name",
            "bad\\name",
            "bad:name",
            "bad!name",
        ] {
            assert!(!directory_names_equal(name, name));
            assert!(validate_directory_name(name).is_err());
        }
    }

    #[test]
    fn validation_counts_utf16_code_units() {
        let valid = "é".repeat(MAX_DIRECTORY_NAME_CODE_UNITS);
        assert!(validate_directory_name(&valid).is_ok());

        let too_long = "é".repeat(MAX_DIRECTORY_NAME_CODE_UNITS + 1);
        assert!(matches!(
            directory_name_data(&too_long),
            Err(DirectoryNameError::TooLong(TOO_LONG_CODE_UNITS))
        ));

        let astral_valid = "😀".repeat(MAX_DIRECTORY_NAME_CODE_UNITS / 2);
        assert!(validate_directory_name(&astral_valid).is_ok());
        let astral_too_long = "😀".repeat(MAX_DIRECTORY_NAME_CODE_UNITS / 2 + 1);
        assert!(validate_directory_name(&astral_too_long).is_err());
    }

    #[test]
    fn giant_name_is_rejected_before_full_utf16_collection() {
        let giant = "😀".repeat(1_000_000);
        assert!(matches!(
            directory_name_data(&giant),
            Err(DirectoryNameError::TooLong(TOO_LONG_CODE_UNITS))
        ));
    }
}
