//! ODS character-reference decoding shared by authoring and worksheet paths.

use core::fmt;

/// A character reference in XML content that cannot be represented under the
/// XML 1.0 character policy used by ODS.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum Error {
    InvalidUtf8,
    InvalidNumber,
    IllegalCharacter,
    UnknownName,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InvalidUtf8 => "character reference name is not UTF-8",
            Self::InvalidNumber => "character reference has an invalid number",
            Self::IllegalCharacter => "character reference names an illegal XML 1.0 character",
            Self::UnknownName => "character reference has an unknown name",
        })
    }
}

/// Resolve the bytes between `&` and `;` in one XML general reference.
///
/// The five XML predefined entities and decimal/hexadecimal character
/// references are accepted. Numeric references are filtered with the XML 1.0
/// character set before becoming a Rust `char`; no replacement character is
/// introduced for malformed or unsupported input.
pub(crate) fn decode(reference: &[u8]) -> Result<char, Error> {
    match reference {
        b"amp" => return Ok('&'),
        b"lt" => return Ok('<'),
        b"gt" => return Ok('>'),
        b"quot" => return Ok('"'),
        b"apos" => return Ok('\''),
        _ => {},
    }

    let reference = core::str::from_utf8(reference).map_err(|_| Error::InvalidUtf8)?;
    let value = if let Some(hexadecimal) = reference.strip_prefix("#x") {
        if hexadecimal.is_empty() {
            return Err(Error::InvalidNumber);
        }
        u32::from_str_radix(hexadecimal, 16).map_err(|_| Error::InvalidNumber)?
    } else if let Some(decimal) = reference.strip_prefix('#') {
        if decimal.is_empty() {
            return Err(Error::InvalidNumber);
        }
        decimal.parse::<u32>().map_err(|_| Error::InvalidNumber)?
    } else {
        return Err(Error::UnknownName);
    };

    let character = char::from_u32(value).ok_or(Error::IllegalCharacter)?;
    if !is_xml_1_0_character(character) {
        return Err(Error::IllegalCharacter);
    }
    Ok(character)
}

const fn is_xml_1_0_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

#[cfg(test)]
mod tests {
    use super::{Error, decode};

    #[test]
    fn decodes_predefined_and_numeric_references_without_second_pass() {
        assert_eq!(decode(b"lt"), Ok('<'));
        assert_eq!(decode(b"amp"), Ok('&'));
        assert_eq!(decode(b"quot"), Ok('"'));
        assert_eq!(decode(b"apos"), Ok('\''));
        assert_eq!(decode(b"#13"), Ok('\r'));
        assert_eq!(decode(b"#x0A"), Ok('\n'));
        assert_eq!(decode(b"#x9"), Ok('\t'));
        assert_eq!(decode(b"#x1F600"), Ok('😀'));
    }

    #[test]
    fn rejects_unknown_and_non_xml_1_0_references() {
        for reference in [b"bogus".as_slice(), b"#".as_slice(), b"#x".as_slice()] {
            assert!(decode(reference).is_err(), "{reference:?}");
        }
        for reference in [
            b"#0".as_slice(),
            b"#x1".as_slice(),
            b"#xB".as_slice(),
            b"#xD800".as_slice(),
            b"#x110000".as_slice(),
        ] {
            assert_eq!(decode(reference), Err(Error::IllegalCharacter), "{reference:?}");
        }
    }
}
