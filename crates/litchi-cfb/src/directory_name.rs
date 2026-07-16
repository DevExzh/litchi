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
