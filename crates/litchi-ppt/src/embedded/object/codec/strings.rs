//! Bounded UTF-16 and optional child-record codecs shared by OLE containers.

use super::super::model::{MAX_METAFILE_BYTES, MAX_OLE_NAME_UNITS};
use super::wire::{corrupted, record_bytes};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

#[allow(clippy::type_complexity)]
pub(crate) fn parse_optional_ole_children(
    children: &[Record],
) -> Result<(
    Option<String>,
    Option<String>,
    Option<String>,
    Option<Vec<u8>>,
)> {
    let mut menu_name = None;
    let mut program_id = None;
    let mut clipboard_name = None;
    let mut metafile = None;
    let mut last_string_instance = 0u16;
    for child in children {
        if child.record_type == RecordType::CString {
            if metafile.is_some()
                || !(1..=3).contains(&child.instance)
                || child.instance <= last_string_instance
            {
                return corrupted("OLE object string atoms are duplicated or out of order");
            }
            last_string_instance = child.instance;
            let value = parse_ole_string(child, child.instance != 1)?;
            match child.instance {
                1 => menu_name = Some(value),
                2 => program_id = Some(value),
                3 => clipboard_name = Some(value),
                _ => unreachable!("instance was bounded"),
            }
        } else if child.record_type == RecordType::MetaFile {
            if metafile.is_some()
                || child.version != 0
                || child.instance != 0
                || child.data.len() > MAX_METAFILE_BYTES
                || usize::try_from(child.data_length).ok() != Some(child.data.len())
            {
                return corrupted("MetafileBlob has an invalid header, size, or placement");
            }
            metafile = Some(child.data.clone());
        } else {
            return corrupted("OLE object container contains an unexpected child record");
        }
    }
    Ok((menu_name, program_id, clipboard_name, metafile))
}

pub(crate) fn append_optional_ole_children(
    children: &mut Vec<u8>,
    menu_name: Option<&str>,
    program_id: Option<&str>,
    clipboard_name: Option<&str>,
    metafile: Option<&[u8]>,
) -> Result<()> {
    for (instance, value, printable) in [
        (1, menu_name, false),
        (2, program_id, true),
        (3, clipboard_name, true),
    ] {
        if let Some(value) = value {
            children.extend_from_slice(&record_bytes(
                0,
                instance,
                RecordType::CString,
                &encode_ole_string(value, printable)?,
            )?);
        }
    }
    if let Some(metafile) = metafile {
        if metafile.len() > MAX_METAFILE_BYTES {
            return corrupted("MetafileBlob exceeds 64 MiB");
        }
        children.extend_from_slice(&record_bytes(0, 0, RecordType::MetaFile, metafile)?);
    }
    Ok(())
}

fn parse_ole_string(record: &Record, printable: bool) -> Result<String> {
    if record.version != 0
        || record.record_type != RecordType::CString
        || !record.data.len().is_multiple_of(2)
        || record.data.len() / 2 > MAX_OLE_NAME_UNITS
        || usize::try_from(record.data_length).ok() != Some(record.data.len())
    {
        return corrupted("OLE object string atom has an invalid header or size");
    }
    let units = record
        .data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<Vec<_>>();
    if units.contains(&0) {
        return corrupted("OLE object string contains an embedded null");
    }
    let value = String::from_utf16(&units)
        .map_err(|_| Error::Corrupted("OLE object string contains invalid UTF-16".into()))?;
    if printable && value.chars().any(char::is_control) {
        return corrupted("OLE object printable string contains a control character");
    }
    Ok(value)
}

pub(crate) fn encode_ole_string(value: &str, printable: bool) -> Result<Vec<u8>> {
    if value.contains('\0') {
        return corrupted("OLE object string contains an embedded null");
    }
    if printable && value.chars().any(char::is_control) {
        return corrupted("OLE object printable string contains a control character");
    }
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > MAX_OLE_NAME_UNITS {
        return corrupted(format!(
            "OLE object string exceeds {MAX_OLE_NAME_UNITS} UTF-16 units"
        ));
    }
    Ok(units.into_iter().flat_map(u16::to_le_bytes).collect())
}
