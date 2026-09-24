//! Checked source spans and bounded attribute-value edits for SpreadsheetML.

use std::ops::Range;

use crate::error::{Error, Result};

fn invalid(value: impl std::fmt::Display) -> Error {
    crate::error::invalid(value.to_string())
}

pub(crate) fn value_span(xml: &[u8], raw: &[u8]) -> Result<Range<usize>> {
    let start = (raw.as_ptr() as usize)
        .checked_sub(xml.as_ptr() as usize)
        .ok_or_else(|| invalid("XML attribute value is not source-backed"))?;
    let end = start
        .checked_add(raw.len())
        .ok_or_else(|| invalid("XML attribute source range overflow"))?;
    let quote = start.checked_sub(1).and_then(|at| xml.get(at));
    if xml.get(start..end) != Some(raw)
        || !matches!(quote, Some(b'\'' | b'"'))
        || xml.get(end) != quote
    {
        return Err(invalid("invalid XML attribute source range"));
    }
    Ok(start..end)
}

#[cfg(test)]
pub(crate) fn escaped_xstring(value: &str) -> Vec<u8> {
    let mut output = Vec::with_capacity(value.len());
    append_escaped_xstring_direct(&mut output, value);
    output
}

/// Return the encoded XML attribute length without constructing the encoded
/// SpreadsheetML string or its escaped byte buffer.
pub(crate) fn escaped_xstring_len(value: &str) -> Result<usize> {
    let bytes = value.as_bytes();
    let mut length = 0usize;
    for (at, character) in value.char_indices() {
        let added = if character == '_'
            && bytes
                .get(at..at.saturating_add(7))
                .is_some_and(|slice| crate::raw::strings::spreadsheet_escape_at(slice, 0).is_some())
        {
            7
        } else if matches!(character, '\u{9}' | '\u{A}' | '\u{D}') || character >= '\u{20}' {
            match character {
                '&' => 5,
                '<' => 4,
                '"' | '\'' => 6,
                '\t' | '\n' | '\r' => 5,
                '\u{fffe}' | '\u{ffff}' => 7,
                _ => character.len_utf8(),
            }
        } else {
            // encode_spreadsheet_text emits one seven-byte SpreadsheetML
            // escape for each UTF-16 code unit of an XML-illegal scalar.
            character.encode_utf16(&mut [0; 2]).len() * 7
        };
        length = length
            .checked_add(added)
            .ok_or_else(|| invalid("escaped XML attribute length overflows"))?;
    }
    Ok(length)
}

/// Encode one SpreadsheetML attribute value after an exact, fallible size
/// preflight.  The encoder writes directly from the caller's value so it does
/// not construct an intermediate `String` whose capacity would bypass the
/// caller's allocation policy.
pub(crate) fn try_escaped_xstring(value: &str) -> Result<Vec<u8>> {
    let length = escaped_xstring_len(value)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| crate::error::allocation("escaped XML attribute", source))?;
    append_escaped_xstring_direct(&mut output, value);
    debug_assert_eq!(output.len(), length);
    Ok(output)
}

pub(crate) fn append_escaped_xstring(output: &mut Vec<u8>, value: &str) {
    append_escaped_xstring_direct(output, value);
}

fn append_escaped_xstring_direct(output: &mut Vec<u8>, value: &str) {
    let bytes = value.as_bytes();
    for (at, character) in value.char_indices() {
        if character == '_'
            && bytes
                .get(at..at.saturating_add(7))
                .is_some_and(|slice| crate::raw::strings::spreadsheet_escape_at(slice, 0).is_some())
        {
            output.extend_from_slice(b"_x005F_");
            continue;
        }
        if matches!(character, '\u{9}' | '\u{A}' | '\u{D}') || character >= '\u{20}' {
            match character {
                '&' => output.extend_from_slice(b"&amp;"),
                '<' => output.extend_from_slice(b"&lt;"),
                '"' => output.extend_from_slice(b"&quot;"),
                '\'' => output.extend_from_slice(b"&apos;"),
                '\t' => output.extend_from_slice(b"&#x9;"),
                '\n' => output.extend_from_slice(b"&#xA;"),
                '\r' => output.extend_from_slice(b"&#xD;"),
                '\u{fffe}' => output.extend_from_slice(b"_xFFFE_"),
                '\u{ffff}' => output.extend_from_slice(b"_xFFFF_"),
                _ => {
                    let mut encoded = [0; 4];
                    output.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
                },
            }
            continue;
        }
        let mut units = [0; 2];
        for unit in character.encode_utf16(&mut units) {
            let mut escape = [0u8; 7];
            escape[0] = b'_';
            escape[1] = b'x';
            let digits = b"0123456789ABCDEF";
            escape[2] = digits[usize::from(*unit >> 12)];
            escape[3] = digits[usize::from((*unit >> 8) & 0xF)];
            escape[4] = digits[usize::from((*unit >> 4) & 0xF)];
            escape[5] = digits[usize::from(*unit & 0xF)];
            escape[6] = b'_';
            output.extend_from_slice(&escape);
        }
    }
}

pub(crate) fn validate_xml_characters(xml: &[u8]) -> Result<()> {
    let text = std::str::from_utf8(xml).map_err(invalid)?;
    if text.chars().any(|c| !matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
        return Err(invalid("invalid XML character"));
    }
    Ok(())
}
