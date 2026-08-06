//! Validation shared by link snapshots and transactional publications.

use super::model::Link;
use litchi_cfb::OleError;
use std::ops::Range;

pub(crate) fn validate(link: &Link) -> Result<(), OleError> {
    let length = link.wire.len();
    if length > super::codec::MAX_BYTES {
        return Err(invalid("OLE link stream exceeds the metadata limit"));
    }
    if length < 16 {
        return Err(invalid("OLE link stream header is truncated"));
    }
    if link.kind != super::model::Kind::from_flags(link.flags) {
        return Err(invalid("OLE link kind does not match its flags"));
    }
    if link.tail_offset > length {
        return Err(invalid("OLE link tail falls outside the source stream"));
    }

    let mut previous_end = 16;
    for (label, range, moniker) in [
        ("reserved moniker", link.reserved_moniker.as_ref(), false),
        ("relative moniker", link.relative_source.as_ref(), true),
        ("absolute moniker", link.absolute_source.as_ref(), true),
    ] {
        if let Some(range) = range {
            previous_end = check_ordered_range(range, previous_end, length, label, moniker)?;
        }
    }

    if let Some(offset) = link.class_id_offset {
        check_field(offset, 16, length, "class identifier")?;
        if link.class_id.is_none() {
            return Err(invalid(
                "OLE link class identifier offset has no typed value",
            ));
        }
        if offset < previous_end
            || offset
                .checked_add(16)
                .is_none_or(|end| end > link.tail_offset)
        {
            return Err(invalid(
                "OLE link class identifier is outside the known field range",
            ));
        }
        previous_end = offset + 16;
    } else if link.class_id.is_some() {
        return Err(invalid("OLE link class identifier has no source offset"));
    }

    if let Some(range) = link.reserved_display.as_ref() {
        previous_end = check_ordered_range(range, previous_end, length, "reserved display", false)?;
    }

    if let Some(offsets) = link.times_offsets {
        let timestamp_start = previous_end
            .checked_add(4)
            .ok_or_else(|| invalid("OLE link timestamp start overflows"))?;
        for (index, offset) in offsets.into_iter().enumerate() {
            check_field(offset, 8, length, "timestamp")?;
            if offset < previous_end
                || offset
                    .checked_add(8)
                    .is_none_or(|end| end > link.tail_offset)
            {
                return Err(invalid(format!(
                    "OLE link timestamp {index} is outside the known field range"
                )));
            }
        }
        if offsets[0] != timestamp_start
            || offsets[0] + 8 != offsets[1]
            || offsets[1] + 8 != offsets[2]
            || link.reserved2.is_none()
            || link.times.is_none()
        {
            return Err(invalid("OLE link timestamp tail has invalid offsets"));
        }
    } else if link.times.is_some() || link.reserved2.is_some() {
        return Err(invalid("OLE link timestamp values have no source offsets"));
    }

    if link.tail_offset < previous_end {
        return Err(invalid("OLE link tail precedes a known field"));
    }
    Ok(())
}

fn check_range(range: &Range<usize>, length: usize, label: &str) -> Result<(), OleError> {
    if range.start > range.end || range.end > length {
        return Err(invalid(format!(
            "OLE link {label} range is outside the source stream"
        )));
    }
    Ok(())
}

fn check_ordered_range(
    range: &Range<usize>,
    previous_end: usize,
    length: usize,
    label: &str,
    moniker: bool,
) -> Result<usize, OleError> {
    check_range(range, length, label)?;
    if range.start < previous_end {
        return Err(invalid(format!(
            "OLE link {label} overlaps an earlier known field"
        )));
    }
    if moniker && range.len() < 16 {
        return Err(invalid(format!(
            "OLE link {label} is shorter than its class identifier"
        )));
    }
    Ok(range.end)
}

fn check_field(offset: usize, width: usize, length: usize, label: &str) -> Result<(), OleError> {
    let end = offset
        .checked_add(width)
        .ok_or_else(|| invalid(format!("OLE link {label} offset overflows")))?;
    if end > length {
        return Err(invalid(format!(
            "OLE link {label} field is outside the source stream"
        )));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}
