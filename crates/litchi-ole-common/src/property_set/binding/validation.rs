//! Wire and semantic checks for standard property-set names.

use super::super::model::invalid;
use litchi_cfb::OleError;

pub(super) const PREFIX: u8 = 0x05;
pub(super) const MAX_NAME_BYTES: usize = 27;
pub(super) const GUID_SUFFIX_BYTES: usize = 26;

pub(super) fn validate_name(value: &str) -> Result<(), OleError> {
    let bytes = value.as_bytes();
    if bytes.first().copied() != Some(PREFIX) {
        return Err(invalid("Property Set binding name must start with 0x05"));
    }
    if is_named_binding(value) {
        return Ok(());
    }
    if bytes.len() != MAX_NAME_BYTES {
        return Err(invalid(
            "GUID-derived Property Set binding name must contain 26 suffix characters",
        ));
    }
    if !bytes[1..].iter().all(|byte| {
        byte.is_ascii_lowercase() || byte.is_ascii_uppercase() || *byte >= b'0' && *byte <= b'5'
    }) {
        return Err(invalid(
            "GUID-derived Property Set binding name contains a character outside A-Z or 0-5",
        ));
    }
    Ok(())
}

pub(super) fn is_named_binding(value: &str) -> bool {
    [
        "\u{0005}SummaryInformation",
        "\u{0005}DocumentSummaryInformation",
        "\u{0005}GlobalInfo",
        "\u{0005}ImageContents",
        "\u{0005}ImageInfo",
    ]
    .into_iter()
    .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

pub(super) fn suffix_value(byte: u8) -> Option<u8> {
    match byte.to_ascii_lowercase() {
        b'a'..=b'z' => Some(byte.to_ascii_lowercase() - b'a'),
        b'0'..=b'5' => Some(26 + byte - b'0'),
        _ => None,
    }
}
