//! Conversion between the validated CFB directory API and typed metadata.

use super::catalog::Projection;
use super::model::{EntryKind, Links, Metadata, NOSTREAM, Sid};
use super::validation;
use crate::property_set::Guid;
use litchi_cfb::{DirectoryEntry, OleError};

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

pub(crate) fn decode_links(entry: &DirectoryEntry) -> Result<Links, OleError> {
    let links = Links::from_raw(
        decode_link(entry.sid_left)?,
        decode_link(entry.sid_right)?,
        decode_link(entry.sid_child)?,
    );
    let sid = Sid::new(entry.sid)?;
    if [links.left(), links.right(), links.child()]
        .into_iter()
        .flatten()
        .any(|link| link == sid)
    {
        return Err(OleError::InvalidFormat(
            "CFB directory metadata contains a self-referential link".into(),
        ));
    }
    Ok(links)
}

/// Projects one entry while retaining unsupported entry kinds as raw data.
pub(crate) fn project(entry: &DirectoryEntry) -> Result<Projection, OleError> {
    let links = decode_links(entry)?;
    let metadata = match entry.entry_type {
        0x01 | 0x02 | 0x05 => Some(decode(entry)?),
        _ => None,
    };
    Ok(Projection { metadata, links })
}

/// Applies only fields represented by the typed metadata projection.
///
/// The caller supplies both projections so unchanged raw spellings (notably
/// CLSID text) remain byte-for-byte equivalent in the source model.
pub(crate) fn apply_metadata(entry: &mut DirectoryEntry, before: Metadata, after: Metadata) {
    if before.kind() != after.kind() {
        entry.entry_type = after.kind().raw();
    }
    if before.class_id() != after.class_id() {
        entry.clsid = after.class_id().map(format_class_id).unwrap_or_default();
    }
    if before.links() != after.links() {
        entry.sid_left = after.links().left().map_or(NOSTREAM, Sid::raw);
        entry.sid_right = after.links().right().map_or(NOSTREAM, Sid::raw);
        entry.sid_child = after.links().child().map_or(NOSTREAM, Sid::raw);
    }
    if before.start_sector() != after.start_sector() {
        entry.start_sector = after.start_sector();
    }
    if before.stream_size() != after.stream_size() {
        entry.size = after.stream_size();
    }
    if before.uses_mini_stream() != after.uses_mini_stream() {
        entry.is_minifat = after.uses_mini_stream();
    }
}

pub(crate) fn raw_equal(left: &DirectoryEntry, right: &DirectoryEntry) -> bool {
    left.sid == right.sid
        && left.name == right.name
        && left.entry_type == right.entry_type
        && left.sid_left == right.sid_left
        && left.sid_right == right.sid_right
        && left.sid_child == right.sid_child
        && left.clsid == right.clsid
        && left.start_sector == right.start_sector
        && left.size == right.size
        && left.is_minifat == right.is_minifat
        && left.children.len() == right.children.len()
        && left
            .children
            .iter()
            .zip(&right.children)
            .all(|(child_left, child_right)| raw_equal(child_left, child_right))
}

pub(crate) fn raw_catalog_equal(left: &[DirectoryEntry], right: &[DirectoryEntry]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(entry_left, entry_right)| raw_equal(entry_left, entry_right))
}

pub(crate) fn fingerprint(entries: &[DirectoryEntry]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    hash_usize(&mut value, entries.len());
    for entry in entries {
        hash_entry(&mut value, entry);
    }
    value
}

fn hash_entry(value: &mut u64, entry: &DirectoryEntry) {
    hash_u32(value, entry.sid);
    hash_bytes(value, entry.name.as_bytes());
    hash_u8(value, entry.entry_type);
    hash_u32(value, entry.sid_left);
    hash_u32(value, entry.sid_right);
    hash_u32(value, entry.sid_child);
    hash_bytes(value, entry.clsid.as_bytes());
    hash_u32(value, entry.start_sector);
    hash_u64(value, entry.size);
    hash_u8(value, u8::from(entry.is_minifat));
    hash_usize(value, entry.children.len());
    for child in &entry.children {
        hash_entry(value, child);
    }
}

fn hash_bytes(value: &mut u64, bytes: &[u8]) {
    hash_usize(value, bytes.len());
    for byte in bytes {
        hash_u8(value, *byte);
    }
}

fn hash_usize(value: &mut u64, input: usize) {
    hash_u64(value, input as u64);
}

fn hash_u64(value: &mut u64, input: u64) {
    for byte in input.to_le_bytes() {
        hash_u8(value, byte);
    }
}

fn hash_u32(value: &mut u64, input: u32) {
    for byte in input.to_le_bytes() {
        hash_u8(value, byte);
    }
}

fn hash_u8(value: &mut u64, input: u8) {
    *value ^= u64::from(input);
    *value = value.wrapping_mul(0x0000_0100_0000_01b3);
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
    let compact = body.len() == 32 && !body.contains(&b'-');
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
        push_hex_byte(&mut output, *byte);
    }
    output.push('-');
    for byte in bytes[4..6].iter().rev() {
        push_hex_byte(&mut output, *byte);
    }
    output.push('-');
    for byte in bytes[6..8].iter().rev() {
        push_hex_byte(&mut output, *byte);
    }
    output.push('-');
    for byte in &bytes[8..10] {
        push_hex_byte(&mut output, *byte);
    }
    output.push('-');
    for byte in &bytes[10..16] {
        push_hex_byte(&mut output, *byte);
    }
    output
}

fn push_hex_byte(output: &mut String, byte: u8) {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";

    output.push(char::from(HEX[usize::from(byte >> 4)]));
    output.push(char::from(HEX[usize::from(byte & 0x0F)]));
}

fn hex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}
