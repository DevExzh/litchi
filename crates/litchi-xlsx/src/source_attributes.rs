//! Checked source spans and bounded attribute-value edits for SpreadsheetML.

use crate::error::{Error, Result};

fn invalid(value: impl std::fmt::Display) -> Error {
    crate::error::invalid(value.to_string())
}

#[cfg(test)]
pub(crate) fn escaped_xstring(value: &str) -> Vec<u8> {
    let encoded = crate::raw::strings::encode_spreadsheet_text(value);
    let mut output = Vec::with_capacity(encoded.len());
    append_encoded_xstring(&mut output, &encoded);
    output
}

pub(crate) fn append_escaped_xstring(output: &mut Vec<u8>, value: &str) {
    let encoded = crate::raw::strings::encode_spreadsheet_text(value);
    append_encoded_xstring(output, &encoded);
}

fn append_encoded_xstring(output: &mut Vec<u8>, encoded: &str) {
    for character in encoded.chars() {
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
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
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
