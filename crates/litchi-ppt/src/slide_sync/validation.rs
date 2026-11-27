//! Structural and resource validation for synchronization snapshots.

use super::model::{LibraryUrl, Limits, ServerId, SystemTime};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

pub(crate) fn validate(root: &Record, limits: Limits) -> Result<()> {
    validate_slide_header(root)?;
    let mut count = 0usize;
    walk(root, 1, &mut count, limits)?;
    let mut sync_count = 0;
    for child in &root.children {
        if is_sync_record(child) {
            sync_count += 1;
            validate_sync_record(child)?;
        }
    }
    if sync_count > 1 {
        return Err(Error::Corrupted(
            "slide contains duplicate RoundTripSlideSyncInfo12 containers".into(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_slide_header(root: &Record) -> Result<()> {
    if root.record_type != RecordType::Slide
        || root.record_type_raw != RecordType::Slide.as_u16()
        || root.version != 0x0f
        || root.instance != 0
    {
        return Err(Error::InvalidFormat(
            "slide synchronization owner requires a SlideContainer root".into(),
        ));
    }
    Ok(())
}

pub(crate) fn is_sync_record(record: &Record) -> bool {
    record.record_type == RecordType::RoundTripSlideSyncInfo12
        && record.record_type_raw == RecordType::RoundTripSlideSyncInfo12.as_u16()
}

pub(crate) fn sync_index(root: &Record) -> Result<Option<usize>> {
    validate_slide_header(root)?;
    let mut index = None;
    for (child_index, child) in root.children.iter().enumerate() {
        if is_sync_record(child) {
            validate_sync_record(child)?;
            if index.replace(child_index).is_some() {
                return Err(Error::Corrupted(
                    "slide contains duplicate RoundTripSlideSyncInfo12 containers".into(),
                ));
            }
        }
    }
    Ok(index)
}

pub(crate) fn insertion_index(root: &Record) -> Result<usize> {
    validate_slide_header(root)?;
    let mut last = None;
    for (index, child) in root.children.iter().enumerate() {
        if is_slide_round_trip(child.record_type) {
            last = Some(index);
        }
    }
    Ok(last.map_or(root.children.len(), |index| index + 1))
}

pub(crate) fn validate_sync_record(record: &Record) -> Result<()> {
    if !is_sync_record(record) || record.version != 0x0f || record.instance != 0 {
        return Err(Error::Corrupted(
            "RoundTripSlideSyncInfo12 has an invalid record header".into(),
        ));
    }
    let children = if record.children.is_empty() {
        Record::parse_sequence_strict(&record.data, "RoundTripSlideSyncInfo12")?
    } else {
        record.children.clone()
    };
    if children.len() != 3 {
        return Err(Error::Corrupted(
            "RoundTripSlideSyncInfo12 must contain exactly three records".into(),
        ));
    }
    validate_text_atom(&children[0], 0, "ServerIdAtom")?;
    validate_text_atom(&children[1], 1, "SlideLibUrlAtom")?;
    if children[2].record_type != RecordType::RoundTripSlideSyncInfoAtom12
        || children[2].record_type_raw != RecordType::RoundTripSlideSyncInfoAtom12.as_u16()
        || children[2].version != 0
        || children[2].instance != 0
        || children[2].data.len() != 32
    {
        return Err(Error::Corrupted(
            "SlideSyncInfoAtom12 has an invalid record header or size".into(),
        ));
    }
    let _ = ServerId::from_wire(&children[0].data)?;
    let _ = LibraryUrl::from_wire(&children[1].data)?;
    let _ = SystemTime::from_wire(&children[2].data[..16], "dateTimeModified")?;
    let _ = SystemTime::from_wire(&children[2].data[16..], "dateTimeInserted")?;
    Ok(())
}

fn validate_text_atom(record: &Record, instance: u16, name: &str) -> Result<()> {
    if record.record_type != RecordType::CString
        || record.record_type_raw != RecordType::CString.as_u16()
        || record.version != 0
        || record.instance != instance
        || !record.data.len().is_multiple_of(2)
    {
        return Err(Error::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    Ok(())
}

fn walk(record: &Record, depth: usize, count: &mut usize, limits: Limits) -> Result<()> {
    if depth > limits.max_depth {
        return Err(Error::InvalidFormat(
            "slide synchronization record nesting exceeds the depth limit".into(),
        ));
    }
    *count = (*count).checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("slide synchronization record count overflow".into())
    })?;
    if *count > limits.max_records {
        return Err(Error::InvalidFormat(
            "slide synchronization record count exceeds the limit".into(),
        ));
    }
    for child in &record.children {
        walk(child, depth + 1, count, limits)?;
    }
    Ok(())
}

fn is_slide_round_trip(record_type: RecordType) -> bool {
    matches!(
        record_type,
        RecordType::RoundTripTheme12Atom
            | RecordType::RoundTripColorMapping12Atom
            | RecordType::RoundTripCompositeMasterId12Atom
            | RecordType::RoundTripSlideSyncInfo12
            | RecordType::RoundTripAnimationHash12Atom
            | RecordType::RoundTripAnimation12Atom
            | RecordType::RoundTripContentMasterId12Atom
    )
}
