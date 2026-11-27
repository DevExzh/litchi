//! Context and record invariants for notes-master text styles.

use super::model::{MAX_PACKAGE_BYTES, Styles};
use crate::consts::RecordType;
use crate::master_layout::Context;
use crate::package::{Error, Result};
use crate::records::Record;

pub(crate) fn validate(context: Context, root: &Record) -> Result<()> {
    if context != Context::Notes {
        return Err(Error::InvalidFormat(
            "notes-master text styles are only valid on a notes master".into(),
        ));
    }
    super::super::validation::validate(context, root)?;

    let mut index = None;
    for (child_index, child) in root.children.iter().enumerate() {
        if !is_record(child) {
            continue;
        }
        validate_record(child)?;
        if index.replace(child_index).is_some() {
            return Err(Error::Corrupted(
                "notes master contains duplicate text-style packages".into(),
            ));
        }
    }
    Ok(())
}

pub(crate) fn styles_index(context: Context, root: &Record) -> Result<Option<usize>> {
    validate(context, root)?;
    Ok(root.children.iter().position(is_record))
}

pub(crate) fn styles_insertion_index(context: Context, root: &Record) -> Result<usize> {
    validate(context, root)?;

    // Notes round-trip atoms form the final typed region of a NotesContainer.
    // Insert after the existing typed region when one exists; otherwise append
    // so the original unknown-record order remains unchanged.
    let mut last_round_trip = None;
    for (index, child) in root.children.iter().enumerate() {
        if matches!(
            child.record_type,
            RecordType::RoundTripTheme12Atom
                | RecordType::RoundTripColorMapping12Atom
                | RecordType::RoundTripNotesMasterTextStyles12Atom
        ) {
            last_round_trip = Some(index);
        }
    }
    Ok(last_round_trip.map_or(root.children.len(), |index| index + 1))
}

pub(crate) fn is_record(record: &Record) -> bool {
    record.record_type == RecordType::RoundTripNotesMasterTextStyles12Atom
        && record.record_type_raw == RecordType::RoundTripNotesMasterTextStyles12Atom.as_u16()
}

pub(crate) fn validate_record(record: &Record) -> Result<()> {
    if record.version != 0 || record.instance != 0 || !record.children.is_empty() {
        return Err(Error::Corrupted(
            "notes-master text-style atom has an invalid header or children".into(),
        ));
    }
    let data_length = usize::try_from(record.data_length).map_err(|_err| {
        Error::Corrupted("notes-master text-style atom length exceeds usize".into())
    })?;
    if data_length != record.data.len() {
        return Err(Error::Corrupted(
            "notes-master text-style atom length does not match its payload".into(),
        ));
    }
    if record.data.len() > MAX_PACKAGE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "notes-master text-style atom exceeds {MAX_PACKAGE_BYTES} bytes"
        )));
    }
    Styles::from_package(record.data.clone()).map(|_| ())
}
