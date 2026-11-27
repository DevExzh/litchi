//! Structural and resource validation for DOC embedded-object edits.

use super::super::model::{Editor, FieldMarker, WriteOptions};
use super::{MAX_FIELDS, MAX_PICF, MAX_PIECES, corrupted, u32_at};
use crate::package::Result;

pub(in crate::embedded_object) fn validate_existing_fields(
    fields: &[FieldMarker],
    main_ccp: u32,
) -> Result<()> {
    let mut stack = Vec::new();
    for marker in fields {
        match marker.descriptor[0] & 0x1F {
            0x13 => stack.push((marker.cp, marker.descriptor[1], false)),
            0x14 => {
                let Some(value) = stack.last_mut() else {
                    return Err(corrupted("orphan field separator"));
                };
                if value.2 {
                    return Err(corrupted("duplicate field separator"));
                }
                value.2 = true;
            },
            0x15 => {
                if stack.pop().is_none() {
                    return Err(corrupted("orphan field end"));
                }
            },
            _ => return Err(corrupted("invalid field marker descriptor")),
        }
    }
    if !stack.is_empty() || fields.last().is_some_and(|marker| marker.cp >= main_ccp) {
        return Err(corrupted("unclosed field structure"));
    }
    Ok(())
}

pub(in crate::embedded_object) fn validate_options(
    value: &WriteOptions,
    editor: &Editor,
) -> Result<()> {
    if value.storage_id == 0 || value.storage_id > i32::MAX as u32 {
        return Err(corrupted("storage ID must be a positive signed integer"));
    }
    if value.instruction.is_empty()
        || value.instruction.encode_utf16().count() > 4_096
        || value
            .instruction
            .chars()
            .any(|c| matches!(c, '\u{13}' | '\u{14}' | '\u{15}'))
    {
        return Err(corrupted("object instruction is invalid"));
    }
    if value.picture_data.len() < 4
        || value.picture_data.len() > MAX_PICF
        || u32_at(&value.picture_data, 0)? as usize != value.picture_data.len()
    {
        return Err(corrupted("PICF block length prefix is invalid"));
    }
    if editor.pieces.len() + 2 > MAX_PIECES || editor.fields.len() + 3 > MAX_FIELDS {
        return Err(corrupted("object insertion exceeds resource limits"));
    }
    Ok(())
}
