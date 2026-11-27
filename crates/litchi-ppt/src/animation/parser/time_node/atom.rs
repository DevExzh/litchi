//! `TimeNodeAtom` decoding.

use super::super::support::{parse_optional_time_value, read_i32, read_u32, require_atom};
use super::model::AtomFlags;
use crate::animation::types::{TimeNodeAtom, TimeNodeFill, TimeNodeKind, TimeNodeRestart};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Parse the exact 32-byte payload of a `TimeNodeAtom`.
///
/// # Errors
///
/// Returns an error if the record header is malformed or a property field
/// holds a value the format does not permit.
pub fn parse_time_node_atom(record: &Record) -> Result<TimeNodeAtom> {
    require_atom(record, RecordType::TimeNode, 0, 32, "TimeNodeAtom")?;
    let flags = AtomFlags::from_raw(read_u32(&record.data, 28));
    let fill = parse_optional_time_value(
        flags.fill_property,
        read_u32(&record.data, 12),
        TimeNodeFill::parse,
        "TimeNodeAtom fill",
    )?;
    let restart = parse_optional_time_value(
        flags.restart_property,
        read_u32(&record.data, 4),
        TimeNodeRestart::parse,
        "TimeNodeAtom restart",
    )?;
    let node_type = parse_optional_time_value(
        flags.grouping_type_property,
        read_u32(&record.data, 8),
        TimeNodeKind::parse,
        "TimeNodeAtom type",
    )?;
    let duration = read_i32(&record.data, 24);
    let duration_ms = if flags.duration_property {
        Some(duration)
    } else if duration == 0 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "TimeNodeAtom duration must be zero when not explicitly set".to_string(),
        ));
    };
    Ok(TimeNodeAtom {
        fill,
        restart,
        node_type,
        duration_ms,
    })
}
