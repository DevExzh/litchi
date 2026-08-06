//! Main-master context and TemplateNameAtom wire invariants.

use super::model::Name;
use crate::consts::RecordType;
use crate::master_layout::Context;
use crate::package::{Error, Result};
use crate::records::Record;

pub(super) fn validate(context: Context, root: &Record) -> Result<()> {
    if context != Context::Main {
        return Err(Error::InvalidFormat(
            "TemplateNameAtom is only valid on a main master".into(),
        ));
    }
    super::super::validation::validate(context, root)?;
    let mut index = None;
    for (child_index, child) in root.children.iter().enumerate() {
        if !is_template_record(child) {
            continue;
        }
        validate_record(child)?;
        if index.replace(child_index).is_some() {
            return Err(Error::Corrupted(
                "main master contains duplicate TemplateNameAtom records".into(),
            ));
        }
    }
    Ok(())
}

pub(crate) fn template_index(context: Context, root: &Record) -> Result<Option<usize>> {
    validate(context, root)?;
    Ok(root.children.iter().position(is_template_record))
}

pub(crate) fn template_insertion_index(context: Context, root: &Record) -> Result<usize> {
    validate(context, root)?;
    // TemplateNameAtom is the final named field in MainMasterContainer.
    // Appending keeps every opaque record in its original order and places the
    // typed field after any existing round-trip region.
    Ok(root.children.len())
}

pub(super) fn is_template_record(record: &Record) -> bool {
    record.record_type == RecordType::CString
        && record.record_type_raw == RecordType::CString.as_u16()
        && record.instance == 2
}

pub(super) fn validate_record(record: &Record) -> Result<()> {
    if record.version != 0 || record.instance != 2 || !record.children.is_empty() {
        return Err(Error::Corrupted(
            "TemplateNameAtom has an invalid header or child records".into(),
        ));
    }
    if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
        return Err(Error::Corrupted(
            "TemplateNameAtom length does not match its payload".into(),
        ));
    }
    Name::from_wire(&record.data).map(|_| ())
}
