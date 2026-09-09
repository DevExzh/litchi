//! The `[MS-OLEPS]` 2.23 binding-name codec.

use super::super::model::{
    DOCUMENT_SUMMARY_INFORMATION_FMTID, Guid, SUMMARY_INFORMATION_FMTID,
    USER_DEFINED_PROPERTIES_FMTID,
};
use super::model::{BindingName, GLOBAL_INFO_FMTID, IMAGE_CONTENTS_FMTID, IMAGE_INFO_FMTID};
use super::validation::{
    GUID_SUFFIX_BYTES, MAX_NAME_BYTES, PREFIX, is_named_binding, suffix_value, validate_name,
};
use litchi_cfb::OleError;

const ALPHABET: &[u8; 32] = b"abcdefghijklmnopqrstuvwxyz012345";

pub(super) fn encode(format_identifier: Guid) -> BindingName {
    if let Some(name) = named_name(format_identifier) {
        let mut bytes = [0u8; MAX_NAME_BYTES];
        bytes[..name.len()].copy_from_slice(name);
        return BindingName::from_bytes(bytes, name.len());
    }

    let mut bytes = [0u8; MAX_NAME_BYTES];
    bytes[0] = PREFIX;
    let raw = format_identifier.as_bytes();
    // GUID (Packet Version) bits are consumed in packet-byte order, from the
    // least significant bit to the most significant bit of each byte.
    for group in 0..GUID_SUFFIX_BYTES {
        let mut value = 0u8;
        for bit in 0..5 {
            let position = group * 5 + bit;
            let bit_value = if position < 128 {
                (raw[position / 8] >> (position % 8)) & 1
            } else {
                0
            };
            value |= bit_value << bit;
        }
        bytes[group + 1] = ALPHABET[usize::from(value)];
    }
    BindingName::from_bytes(bytes, MAX_NAME_BYTES)
}

pub(super) fn decode(name: &str) -> Result<Guid, OleError> {
    validate_name(name)?;
    if is_named_binding(name) {
        return named_identifier(name)
            .ok_or_else(|| super::super::model::invalid("Unknown named Property Set binding"));
    }

    let mut raw = [0u8; 16];
    for (group, byte) in name.as_bytes()[1..].iter().copied().enumerate() {
        let value = suffix_value(byte).ok_or_else(|| {
            super::super::model::invalid("Invalid GUID-derived Property Set binding character")
        })?;
        for bit in 0..5 {
            let position = group * 5 + bit;
            if position >= 128 {
                if value & (1u8 << bit) != 0 {
                    return Err(super::super::model::invalid(
                        "GUID-derived Property Set binding name has nonzero trailing bits",
                    ));
                }
            } else if value & (1u8 << bit) != 0 {
                raw[position / 8] |= 1u8 << (position % 8);
            }
        }
    }
    Ok(Guid::from_bytes(raw))
}

fn named_name(format_identifier: Guid) -> Option<&'static [u8]> {
    Some(match format_identifier {
        SUMMARY_INFORMATION_FMTID => b"\x05SummaryInformation",
        DOCUMENT_SUMMARY_INFORMATION_FMTID | USER_DEFINED_PROPERTIES_FMTID => {
            b"\x05DocumentSummaryInformation"
        },
        GLOBAL_INFO_FMTID => b"\x05GlobalInfo",
        IMAGE_CONTENTS_FMTID => b"\x05ImageContents",
        IMAGE_INFO_FMTID => b"\x05ImageInfo",
        _ => return None,
    })
}

fn named_identifier(name: &str) -> Option<Guid> {
    if name.eq_ignore_ascii_case("\u{0005}SummaryInformation") {
        Some(SUMMARY_INFORMATION_FMTID)
    } else if name.eq_ignore_ascii_case("\u{0005}DocumentSummaryInformation") {
        Some(DOCUMENT_SUMMARY_INFORMATION_FMTID)
    } else if name.eq_ignore_ascii_case("\u{0005}GlobalInfo") {
        Some(GLOBAL_INFO_FMTID)
    } else if name.eq_ignore_ascii_case("\u{0005}ImageContents") {
        Some(IMAGE_CONTENTS_FMTID)
    } else if name.eq_ignore_ascii_case("\u{0005}ImageInfo") {
        Some(IMAGE_INFO_FMTID)
    } else {
        None
    }
}
