//! Contextual invariants for main, title, notes, and handout masters.

use super::model::{Context, Limits};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

const SL_MASTER_TITLE: u32 = 2;

pub(super) fn validate(context: Context, root: &Record, limits: Limits) -> Result<()> {
    limits.validate()?;
    let mut count = 0usize;
    walk(root, 1, limits, &mut count)?;

    if root.record_type != context.expected_record_type()
        || root.record_type_raw != context.expected_record_type().as_u16()
        || root.version != 0x0f
        || root.instance != 0
    {
        return corrupted(format!(
            "{} master has an invalid root record header",
            context.name()
        ));
    }

    let mut slide_atom = None;
    let mut notes_atom = None;
    let mut drawings = 0usize;
    for child in &root.children {
        match child.record_type_raw {
            value if value == RecordType::SlideAtom.as_u16() => {
                validate_atom(child, 0x02, 0, 24, "SlideAtom")?;
                if slide_atom.replace(child).is_some() {
                    return corrupted("master contains duplicate SlideAtom records");
                }
            },
            value if value == RecordType::NotesAtom.as_u16() => {
                validate_atom(child, 0x01, 0, 8, "NotesAtom")?;
                if notes_atom.replace(child).is_some() {
                    return corrupted("master contains duplicate NotesAtom records");
                }
            },
            value if value == RecordType::PPDrawing.as_u16() => {
                if child.version != 0x0f || child.instance != 0 {
                    return corrupted("master drawing has an invalid container header");
                }
                drawings += 1;
                if drawings > 1 {
                    return corrupted("master contains duplicate PPDrawing records");
                }
            },
            _ => {},
        }
    }

    match context {
        Context::Main => {
            let atom = slide_atom
                .ok_or_else(|| Error::Corrupted("main master is missing its SlideAtom".into()))?;
            let master_id = read_u32(atom, 12);
            let notes_id = read_u32(atom, 16);
            if master_id != 0 || notes_id != 0 {
                return corrupted(
                    "main master SlideAtom must have null masterIdRef and notesIdRef",
                );
            }
        },
        Context::Title => {
            let atom = slide_atom
                .ok_or_else(|| Error::Corrupted("title master is missing its SlideAtom".into()))?;
            if read_u32(atom, 0) != SL_MASTER_TITLE {
                return corrupted("title master SlideAtom does not use SL_MasterTitle");
            }
            if read_u32(atom, 12) == 0 {
                return corrupted("title master SlideAtom must reference its main master");
            }
        },
        Context::Notes => {
            let atom = notes_atom
                .ok_or_else(|| Error::Corrupted("notes master is missing its NotesAtom".into()))?;
            if read_u32(atom, 0) != 0 || read_u16(atom, 4) != 0 {
                return corrupted(
                    "notes master NotesAtom must have null slideIdRef and slide flags",
                );
            }
        },
        Context::Handout => {},
    }
    if drawings != 1 {
        return corrupted("master must contain exactly one drawing container");
    }
    Ok(())
}

fn walk(record: &Record, depth: usize, limits: Limits, count: &mut usize) -> Result<()> {
    if depth > limits.max_depth {
        return invalid("master-layout record nesting exceeds the configured depth limit");
    }
    *count = count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("master-layout record count overflow".into()))?;
    if *count > limits.max_records {
        return invalid("master-layout record count exceeds the configured limit");
    }
    if record.version > 0x0f || record.instance > 0x0fff {
        return invalid("PPT record header exceeds its version or instance bit field");
    }
    if record.data_length as usize != record.data.len() {
        return corrupted("PPT record data length does not match its payload");
    }
    for child in &record.children {
        walk(child, depth + 1, limits, count)?;
    }
    Ok(())
}

fn validate_atom(
    record: &Record,
    version: u16,
    instance: u16,
    length: usize,
    name: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.data.len() != length
        || !record.children.is_empty()
    {
        return corrupted(format!("{name} has an invalid header or payload"));
    }
    Ok(())
}

fn read_u16(record: &Record, offset: usize) -> u16 {
    u16::from_le_bytes([record.data[offset], record.data[offset + 1]])
}

fn read_u32(record: &Record, offset: usize) -> u32 {
    u32::from_le_bytes([
        record.data[offset],
        record.data[offset + 1],
        record.data[offset + 2],
        record.data[offset + 3],
    ])
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
