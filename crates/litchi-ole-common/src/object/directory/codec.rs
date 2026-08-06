//! Conversion between the validated CFB directory API and typed metadata.

use super::model::{EntryKind, Links, Metadata, NOSTREAM, Sid};
use super::validation;
use crate::property_set::Guid;
use litchi_cfb::{DirectoryEntry, OleError};
use std::fmt::Write as _;

pub(crate) fn decode(entry: &DirectoryEntry) -> Result<Metadata, OleError> {
    let kind = match entry.entry_type {
        0x01 => EntryKind::Storage,
        0x02 => EntryKind::Stream,
        0x05 => EntryKind::Root,
        value => {
            return Err(OleError::InvalidFormat(format!(
                "unsupported CFB directory object type {value:#04X}"
            )));
        },
    };
    let sid = Sid::new(entry.sid)?;
    let links = Links::from_raw(
        decode_link(entry.sid_left)?,
        decode_link(entry.sid_right)?,
        decode_link(entry.sid_child)?,
    );
    let metadata = Metadata::new(
        sid,
        kind,
        parse_class_id(&entry.clsid)?,
        links,
        entry.start_sector,
        entry.size,
        entry.is_minifat,
    );
    validation::validate(metadata)?;
    Ok(metadata)
}

fn decode_link(raw: u32) -> Result<Option<Sid>, OleError> {
    if raw == NOSTREAM {
        Ok(None)
    } else {
        Sid::new(raw).map(Some)
    }
}

pub(crate) fn parse_class_id(input: &str) -> Result<Option<Guid>, OleError> {
    if input.is_empty() {
        return Ok(None);
    }
    let source = input.as_bytes();
    let (body, wrapped) = match (source.first(), source.last()) {
        (Some(b'{'), Some(b'}')) => (&source[1..source.len() - 1], true),
        _ => (source, false),
    };
    if wrapped && (body.is_empty() || body.iter().any(|byte| *byte == b'{' || *byte == b'}')) {
        return Err(invalid_class_id());
    }
    let canonical = body.len() == 36
        && body[8] == b'-'
        && body[13] == b'-'
        && body[18] == b'-'
        && body[23] == b'-';
    let compact = body.len() == 32 && !body.iter().any(|byte| *byte == b'-');
    if !canonical && !compact {
        return Err(invalid_class_id());
    }

    let mut digits = [0u8; 32];
    let mut count = 0usize;
    for byte in body.iter().copied() {
        if byte == b'-' {
            continue;
        }
        let value = hex(byte).ok_or_else(invalid_class_id)?;
        let slot = digits.get_mut(count).ok_or_else(invalid_class_id)?;
        *slot = value;
        count += 1;
    }
    if count != digits.len() {
        return Err(invalid_class_id());
    }

    let mut display = [0u8; 16];
    for (index, byte) in display.iter_mut().enumerate() {
        *byte = (digits[index * 2] << 4) | digits[index * 2 + 1];
    }
    let cfb = [
        display[3],
        display[2],
        display[1],
        display[0],
        display[5],
        display[4],
        display[7],
        display[6],
        display[8],
        display[9],
        display[10],
        display[11],
        display[12],
        display[13],
        display[14],
        display[15],
    ];
    if cfb == [0; 16] {
        Ok(None)
    } else {
        Ok(Some(Guid::from_bytes(cfb)))
    }
}

fn invalid_class_id() -> OleError {
    OleError::InvalidFormat("CFB directory CLSID is not a valid GUID".into())
}

pub(crate) fn format_class_id(value: Guid) -> String {
    let bytes = value.as_bytes();
    let mut output = String::with_capacity(36);
    for byte in bytes[0..4].iter().rev() {
        let _ = write!(output, "{byte:02X}");
    }
    output.push('-');
    for byte in bytes[4..6].iter().rev() {
        let _ = write!(output, "{byte:02X}");
    }
    output.push('-');
    for byte in bytes[6..8].iter().rev() {
        let _ = write!(output, "{byte:02X}");
    }
    output.push('-');
    for byte in &bytes[8..10] {
        let _ = write!(output, "{byte:02X}");
    }
    output.push('-');
    for byte in &bytes[10..16] {
        let _ = write!(output, "{byte:02X}");
    }
    output
}

fn hex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}
