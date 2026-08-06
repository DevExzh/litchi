//! Bounded lexical and semantic validation for `symEx`.

use crate::error::{Error, Result};

use super::model::{MAX_FONT_CHARS, MAX_SYMBOLS, Symbol, Symbols};

/// Maximum source XML retained by one symbol snapshot.
pub(crate) const MAX_XML_BYTES: usize = 4 * 1024 * 1024;
/// Maximum elements scanned in one run fragment.
pub(crate) const MAX_XML_NODES: usize = 16 * 1024;
/// Maximum nesting depth accepted in one run fragment.
pub(crate) const MAX_XML_DEPTH: usize = 64;
/// Target namespace of the Word 2015 `symEx` element and attributes.
pub(crate) const SYMEX_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/word/2015/wordml/symex";
/// Markup-compatibility namespace used by canonical authored symbols.
pub(crate) const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

pub(crate) fn validate_symbol(value: &Symbol) -> Result<()> {
    if let Some(font) = value.font_value()
        && font.chars().count() > MAX_FONT_CHARS
    {
        return Err(Error::Invalid(format!(
            "Word symEx font exceeds {MAX_FONT_CHARS} characters"
        )));
    }

    if let Some(character) = value.character_value()
        && char::from_u32(character).is_none()
    {
        return Err(Error::Invalid(format!(
            "Word symEx char 0x{character:08X} is not a Unicode scalar value"
        )));
    }
    Ok(())
}

pub(crate) fn validate_symbols(value: &Symbols) -> Result<()> {
    if value.values().len() > MAX_SYMBOLS {
        return Err(Error::Invalid(format!(
            "Word run symbols exceed {MAX_SYMBOLS} elements"
        )));
    }
    for symbol in value.values() {
        validate_symbol(symbol)?;
    }
    Ok(())
}

pub(crate) fn parse_char(value: &str) -> Result<u32> {
    if value.is_empty() || value.len() > 8 || !value.bytes().all(is_ascii_hex_digit) {
        return Err(Error::InvalidFormat(format!(
            "invalid Word symEx char '{value}'; expected one to eight hexadecimal digits"
        )));
    }
    let character = u32::from_str_radix(value, 16).map_err(|error| {
        Error::InvalidFormat(format!("invalid Word symEx char '{value}': {error}"))
    })?;
    if char::from_u32(character).is_none() {
        return Err(Error::InvalidFormat(format!(
            "Word symEx char '{value}' is not a Unicode scalar value"
        )));
    }
    Ok(character)
}

#[inline]
pub(crate) fn is_ascii_hex_digit(value: u8) -> bool {
    value.is_ascii_digit() || matches!(value, b'a'..=b'f' | b'A'..=b'F')
}
