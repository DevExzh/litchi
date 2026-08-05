//! BIFF8 NAME (Lbl) record writer.
//!
//! This module encodes workbook-level defined names (named ranges) into
//! BIFF8 `NAME` records, following the [MS-XLS] Lbl structure and the
//! option flag layout used by Apache POI's `NameRecord`.

use std::io::Write;

use crate::writer::DefinedName;
use crate::writer::DefinedNameRecordOptions;
use crate::{DefinedNameKind, NameScope};
use crate::{Error, Result};

use super::{has_multibyte_char, write_record_header};

/// Write a single NAME (Lbl) record.
///
/// The `rgce` argument contains the BIFF8 formula bytes for the defined
/// name (for example, a `PtgRef` or `PtgArea` sequence produced by the
/// formula module).
pub(crate) fn write_name<W: Write>(writer: &mut W, name: &DefinedName, rgce: &[u8]) -> Result<()> {
    if rgce.len() > u16::MAX as usize {
        return Err(Error::InvalidData(
            "Named range formula exceeds BIFF8 length limit".to_string(),
        ));
    }

    // Determine character count and encoding for the Name field. For
    // built-in names we follow the POI convention and store a single
    // byte built-in code with fBuiltin set.
    let (cch, is_16bit, built_in_code, name_str): (u8, bool, Option<u8>, Option<&str>) =
        if name.is_built_in {
            let code = name.built_in_code.ok_or_else(|| {
                Error::InvalidData("Built-in defined name requires a built_in_code".to_string())
            })?;
            (1, false, Some(code), None)
        } else {
            let char_count = name.name.chars().count();
            if char_count == 0 {
                return Err(Error::InvalidData(
                    "Defined name must not be empty".to_string(),
                ));
            }
            if char_count > u8::MAX as usize {
                return Err(Error::InvalidData(
                    "Defined name must be at most 255 characters".to_string(),
                ));
            }
            let cch = char_count as u8;
            let is_16bit = has_multibyte_char(&name.name);
            (cch, is_16bit, None, Some(name.name.as_str()))
        };

    let name_bytes_len: u16 = if is_16bit {
        1u16 + (cch as u16) * 2
    } else {
        1u16 + (cch as u16)
    };

    let cce = rgce.len() as u16;
    let data_len: u16 = 14u16.saturating_add(name_bytes_len).saturating_add(cce);

    // Option flags (grbit) map directly to Apache POI's NameRecord.Option
    // constants for easier comparison:
    //  - 0x0001: hidden name (fHidden)
    //  - 0x0002: function name (fFunc)
    //  - 0x0020: built-in name (fBuiltin)
    let mut grbit: u16 = 0;
    if name.hidden {
        grbit |= 0x0001;
    }
    if name.is_function {
        grbit |= 0x0002;
    }
    if name.is_built_in {
        grbit |= 0x0020;
    }

    write_record_header(writer, 0x0018, data_len)?;

    // grbit
    writer.write_all(&grbit.to_le_bytes())?;

    // chKey (no keyboard shortcut when using defined ranges programmatically)
    writer.write_all(&[0u8])?;

    // cch (character count of Name)
    writer.write_all(&[cch])?;

    // cce (length of rgce in bytes)
    writer.write_all(&cce.to_le_bytes())?;

    // reserved3
    writer.write_all(&0u16.to_le_bytes())?;

    // itab: 0 for workbook-scoped names, or 1-based sheet index for
    // local names.
    let itab = name.local_sheet.unwrap_or(0);
    writer.write_all(&itab.to_le_bytes())?;

    // reserved4..reserved7
    writer.write_all(&[0u8; 4])?;

    // Name (XLUnicodeStringNoCch): flags + characters.
    if let Some(code) = built_in_code {
        // Built-in names are always single-byte codes.
        writer.write_all(&[0x00])?; // compressed 8-bit
        writer.write_all(&[code])?;
    } else if let Some(s) = name_str {
        if is_16bit {
            writer.write_all(&[0x01])?; // UTF-16LE
            for code_unit in s.encode_utf16() {
                writer.write_all(&code_unit.to_le_bytes())?;
            }
        } else {
            writer.write_all(&[0x00])?; // compressed 8-bit
            writer.write_all(s.as_bytes())?;
        }
    }

    // rgce (formula tokens)
    writer.write_all(rgce)?;

    Ok(())
}

fn push_no_cch_string(data: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    data.push(u8::from(!compressed));
    for unit in units {
        if compressed {
            data.push(unit as u8);
        } else {
            data.extend_from_slice(&unit.to_le_bytes());
        }
    }
}

fn write_continued_record<W: Write>(writer: &mut W, record_type: u16, data: &[u8]) -> Result<()> {
    let mut offset = 0;
    let first = data.len().min(8224);
    write_record_header(writer, record_type, first as u16)?;
    writer.write_all(&data[..first])?;
    offset += first;
    while offset < data.len() {
        let end = (offset + 8224).min(data.len());
        write_record_header(writer, 0x003c, (end - offset) as u16)?;
        writer.write_all(&data[offset..end])?;
        offset = end;
    }
    Ok(())
}

pub(crate) fn write_defined_name_record<W: Write>(
    writer: &mut W,
    name: &DefinedNameRecordOptions,
) -> Result<()> {
    let mut flags = u16::from(name.hidden)
        | (u16::from(name.function) << 1)
        | (u16::from(name.vba_procedure) << 2)
        | (u16::from(name.procedure) << 3)
        | (u16::from(name.calculated_expression) << 4)
        | (u16::from(matches!(name.kind, DefinedNameKind::BuiltIn(_))) << 5)
        | (u16::from(name.function_group) << 6)
        | (u16::from(name.published) << 13)
        | (u16::from(name.workbook_parameter) << 14);
    flags &= 0x7fff;
    let mut data = Vec::new();
    data.extend_from_slice(&flags.to_le_bytes());
    data.push(name.shortcut_key.unwrap_or(0));
    let (name_len, built_in) = match name.built_in() {
        Some(value) => (1usize, Some(value.code())),
        None => (name.name.encode_utf16().count(), None),
    };
    data.push(name_len as u8);
    data.extend_from_slice(&(name.formula_tokens.len() as u16).to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    let itab = match name.scope {
        NameScope::Workbook => 0,
        NameScope::Worksheet(index) => u16::try_from(index + 1)
            .map_err(|_| Error::InvalidData("defined name sheet scope exceeds u16".to_string()))?,
    };
    data.extend_from_slice(&itab.to_le_bytes());
    for value in [
        &name.custom_menu,
        &name.description,
        &name.help_topic,
        &name.status_bar,
    ] {
        data.push(value.chars().count() as u8);
    }
    if let Some(code) = built_in {
        data.extend_from_slice(&[0, code]);
    } else {
        push_no_cch_string(&mut data, &name.name);
    }
    data.extend_from_slice(&name.formula_tokens);
    data.extend_from_slice(&name.formula_extra);
    for value in [
        &name.custom_menu,
        &name.description,
        &name.help_topic,
        &name.status_bar,
    ] {
        data.extend(value.chars().map(|character| character as u8));
    }
    write_continued_record(writer, 0x0018, &data)?;
    if let Some(comment) = &name.comment {
        write_name_comment(writer, name.serialized_name(), comment)?;
    }
    Ok(())
}

pub(crate) fn write_name_comment<W: Write>(
    writer: &mut W,
    name: &str,
    comment: &str,
) -> Result<()> {
    let name_len = name.encode_utf16().count();
    let comment_len = comment.encode_utf16().count();
    if name_len > 255 || comment_len > 255 {
        return Err(Error::InvalidData(
            "NameCmt strings exceed 255 UTF-16 units".to_string(),
        ));
    }
    let mut data = Vec::new();
    data.extend_from_slice(&0x0894u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u64.to_le_bytes());
    data.extend_from_slice(&(name_len as u16).to_le_bytes());
    data.extend_from_slice(&(comment_len as u16).to_le_bytes());
    push_no_cch_string(&mut data, name);
    push_no_cch_string(&mut data, comment);
    write_record_header(writer, 0x0894, data.len() as u16)?;
    writer.write_all(&data)?;
    Ok(())
}

fn push_frt_header(data: &mut Vec<u8>, record_type: u16) {
    data.extend_from_slice(&record_type.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u64.to_le_bytes());
}
fn push_xl_name_unicode(data: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16().collect::<Vec<_>>();
    data.extend_from_slice(&(units.len() as u16).to_le_bytes());
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    data.push(u8::from(!compressed));
    for unit in units {
        if compressed {
            data.push(unit as u8)
        } else {
            data.extend_from_slice(&unit.to_le_bytes())
        }
    }
}

pub(crate) fn write_name_function_group<W: Write>(
    writer: &mut W,
    value: &crate::NameFnGrp12,
) -> Result<()> {
    let mut data = Vec::new();
    push_frt_header(&mut data, 0x0899);
    let count = value.function_name.encode_utf16().count();
    data.extend_from_slice(&(count as u16).to_le_bytes());
    data.extend_from_slice(&u16::from(value.category).to_le_bytes());
    push_xl_name_unicode(&mut data, &value.function_name);
    write_record_header(writer, 0x0899, data.len() as u16)?;
    writer.write_all(&data)?;
    Ok(())
}

pub(crate) fn write_name_publish<W: Write>(
    writer: &mut W,
    value: &crate::NamePublish,
) -> Result<()> {
    let mut data = Vec::new();
    push_frt_header(&mut data, 0x0893);
    let flags = u16::from(value.published) | (u16::from(value.workbook_parameter) << 1);
    data.extend_from_slice(&flags.to_le_bytes());
    push_xl_name_unicode(&mut data, &value.name);
    write_record_header(writer, 0x0893, data.len() as u16)?;
    writer.write_all(&data)?;
    Ok(())
}
