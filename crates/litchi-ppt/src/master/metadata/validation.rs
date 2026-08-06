//! Context and wire invariants for `SlideNameAtom`.

use super::model::Name;
use crate::consts::RecordType;
use crate::master_layout::Context;
use crate::package::{Error, Result};
use crate::records::Record;

pub(super) fn validate(context: Context, root: &Record) -> Result<()> {
    if root.record_type != context.expected_record_type()
        || root.record_type_raw != context.expected_record_type().as_u16()
        || root.version != 0x0f
        || root.instance != 0
    {
        return Err(Error::Corrupted(format!(
            "{} master has an invalid root record header",
            context.name()
        )));
    }

    let mut name_index = None;
    for (index, child) in root.children.iter().enumerate() {
        if !is_name_record(child) {
            continue;
        }
        validate_name_record(child)?;
        if name_index.replace(index).is_some() {
            return Err(Error::Corrupted(
                "master contains duplicate SlideNameAtom records".into(),
            ));
        }
    }
    Ok(())
}

pub(super) fn name_index(context: Context, root: &Record) -> Result<Option<usize>> {
    validate(context, root)?;
    Ok(root.children.iter().position(is_name_record))
}

pub(super) fn name_insertion_index(context: Context, root: &Record) -> usize {
    root.children
        .iter()
        .position(|child| follows_name(context, child))
        .unwrap_or(root.children.len())
}

pub(super) fn is_name_record(record: &Record) -> bool {
    record.record_type == RecordType::CString
        && record.record_type_raw == RecordType::CString.as_u16()
        && record.instance == 3
}

fn validate_name_record(record: &Record) -> Result<()> {
    if record.version != 0 || record.instance != 3 || !record.children.is_empty() {
        return Err(Error::Corrupted(
            "SlideNameAtom has an invalid header or child records".into(),
        ));
    }
    let data_length = usize::try_from(record.data_length)
        .map_err(|_| Error::Corrupted("SlideNameAtom length exceeds usize".into()))?;
    if data_length != record.data.len() {
        return Err(Error::Corrupted(
            "SlideNameAtom length does not match its payload".into(),
        ));
    }
    Name::from_wire(&record.data).map(|_| ())
}

fn follows_name(context: Context, child: &Record) -> bool {
    if child.record_type == RecordType::ProgTags {
        return true;
    }
    if context == Context::Main
        && child.record_type == RecordType::CString
        && child.record_type_raw == RecordType::CString.as_u16()
        && child.instance == 2
    {
        return true;
    }

    matches!(
        (context, child.record_type),
        (
            Context::Main,
            RecordType::RoundTripOriginalMainMasterId12Atom
                | RecordType::RoundTripTheme12Atom
                | RecordType::RoundTripColorMapping12Atom
                | RecordType::RoundTripContentMasterInfo12Atom
                | RecordType::RoundTripOArtTextStyles12Atom
                | RecordType::RoundTripAnimationHash12Atom
                | RecordType::RoundTripAnimation12Atom
                | RecordType::RoundTripCompositeMasterId12Atom
        ) | (
            Context::Title,
            RecordType::RoundTripTheme12Atom
                | RecordType::RoundTripColorMapping12Atom
                | RecordType::RoundTripCompositeMasterId12Atom
                | RecordType::RoundTripSlideSyncInfo12
                | RecordType::RoundTripAnimationHash12Atom
                | RecordType::RoundTripAnimation12Atom
                | RecordType::RoundTripContentMasterId12Atom
        ) | (
            Context::Notes,
            RecordType::RoundTripTheme12Atom
                | RecordType::RoundTripColorMapping12Atom
                | RecordType::RoundTripNotesMasterTextStyles12Atom
        ) | (
            Context::Handout,
            RecordType::RoundTripTheme12Atom | RecordType::RoundTripColorMapping12Atom
        )
    )
}
