//! OLEDS `\x01Ole` wire parsing.

use super::model::{Kind, Link, Times};
use crate::property_set::Guid;
use litchi_cfb::OleError;
use std::ops::Range;
use std::sync::Arc;

pub(crate) const LINKED_FLAG: u32 = 0x0000_0001;
pub(crate) const CACHE_HINT_FLAG: u32 = 0x0000_1000;
const MAX_BYTES: usize = 16 * 1024 * 1024;

struct Parsed {
    kind: Kind,
    flags: u32,
    update_option: u32,
    reserved_moniker: Option<Range<usize>>,
    relative_source: Option<Range<usize>>,
    absolute_source: Option<Range<usize>>,
    class_id: Option<Guid>,
    class_id_offset: Option<usize>,
    reserved_display: Option<Range<usize>>,
    reserved2: Option<u32>,
    times: Option<Times>,
    times_offsets: Option<[usize; 3]>,
    tail_offset: usize,
}

struct Reader<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Reader<'a> {
    const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }

    fn remaining(&self) -> usize {
        self.bytes.len() - self.offset
    }

    fn u32(&mut self, field: &str) -> Result<u32, OleError> {
        Ok(u32::from_le_bytes(self.array(field)?))
    }

    fn u64(&mut self, field: &str) -> Result<u64, OleError> {
        Ok(u64::from_le_bytes(self.array(field)?))
    }

    fn array<const N: usize>(&mut self, field: &str) -> Result<[u8; N], OleError> {
        let bytes = self.take(N, field)?;
        let mut output = [0; N];
        output.copy_from_slice(bytes);
        Ok(output)
    }

    fn take(&mut self, count: usize, field: &str) -> Result<&'a [u8], OleError> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| invalid(format!("{field} length overflows")))?;
        if end > self.bytes.len() {
            return Err(invalid(format!("{field} is truncated")));
        }
        let bytes = &self.bytes[self.offset..end];
        self.offset = end;
        Ok(bytes)
    }

    fn sized(&mut self, field: &str, optional: bool) -> Result<Option<Range<usize>>, OleError> {
        let size_wire = self.u32(&format!("{field} size"))?;
        let size = usize::try_from(size_wire)
            .map_err(|_error| invalid(format!("{field} size exceeds addressable memory")))?;
        if size == 0 {
            if optional {
                return Ok(None);
            }
            return Err(invalid(format!("{field} must be present")));
        }
        if size < 4 {
            return Err(invalid(format!(
                "{field} size is smaller than its size field"
            )));
        }
        let payload = size - 4;
        let start = self.offset;
        self.take(payload, field)?;
        Ok(Some(start..self.offset))
    }
}

pub(crate) fn parse(wire: Arc<[u8]>) -> Result<Link, OleError> {
    if wire.len() > MAX_BYTES {
        return Err(invalid("OLE link stream exceeds the metadata limit"));
    }
    let parsed = parse_fields(wire.as_ref())?;
    Ok(Link {
        wire,
        kind: parsed.kind,
        flags: parsed.flags,
        link_update_option: parsed.update_option,
        reserved_moniker: parsed.reserved_moniker,
        relative_source: parsed.relative_source,
        absolute_source: parsed.absolute_source,
        class_id: parsed.class_id,
        class_id_offset: parsed.class_id_offset,
        reserved_display: parsed.reserved_display,
        reserved2: parsed.reserved2,
        times: parsed.times,
        times_offsets: parsed.times_offsets,
        tail_offset: parsed.tail_offset,
    })
}

fn parse_fields(bytes: &[u8]) -> Result<Parsed, OleError> {
    let mut reader = Reader::new(bytes);
    let version = reader.u32("OLE link version")?;
    if version != Link::VERSION {
        return Err(invalid("OLE link stream has an unsupported version"));
    }
    let flags = reader.u32("OLE link flags")?;
    let kind = Kind::from_flags(flags);
    let update_option = reader.u32("OLE link update option")?;
    let reserved = reader.u32("OLE link reserved field")?;
    if reserved != 0 {
        return Err(invalid("OLE link reserved field must be zero"));
    }

    let reserved_moniker = if reader.remaining() == 0 {
        None
    } else {
        reader.sized("OLE link reserved moniker", true)?
    };
    if kind.is_embedded() {
        return Ok(Parsed {
            kind,
            flags,
            update_option,
            reserved_moniker,
            relative_source: None,
            absolute_source: None,
            class_id: None,
            class_id_offset: None,
            reserved_display: None,
            reserved2: None,
            times: None,
            times_offsets: None,
            tail_offset: reader.offset,
        });
    }

    let relative_source = reader.sized("OLE link relative moniker", true)?;
    validate_moniker(bytes, relative_source.as_ref(), "relative")?;
    let absolute_source = reader.sized("OLE link absolute moniker", false)?;
    validate_moniker(bytes, absolute_source.as_ref(), "absolute")?;

    let (class_id, class_id_offset, reserved_display, reserved2, times, times_offsets) =
        if reader.remaining() == 0 {
            (None, None, None, None, None, None)
        } else {
            let indicator = reader.u32("OLE link class identifier indicator")?;
            if indicator != u32::MAX {
                return Err(invalid(
                    "OLE link class identifier indicator must be 0xFFFFFFFF",
                ));
            }
            let class_id_offset = reader.offset;
            let class_id = Guid::from_bytes(reader.array::<16>("OLE link class identifier")?);
            if reader.remaining() == 0 {
                (
                    Some(class_id),
                    Some(class_id_offset),
                    None,
                    None,
                    None,
                    None,
                )
            } else {
                let display_length_wire = reader.u32("OLE link reserved display length")?;
                let display_length = usize::try_from(display_length_wire).map_err(|_error| {
                    invalid("OLE link reserved display length exceeds addressable memory")
                })?;
                if display_length % 2 != 0 {
                    return Err(invalid(
                        "OLE link reserved display length is not UTF-16 aligned",
                    ));
                }
                let display_start = reader.offset;
                reader.take(display_length, "OLE link reserved display")?;
                let reserved_display = Some(display_start..reader.offset);
                if reader.remaining() < 28 {
                    return Err(invalid(
                        "OLE link timestamp tail is truncated after the display name",
                    ));
                }
                let reserved2 = reader.u32("OLE link Reserved2")?;
                let offsets = [reader.offset, reader.offset + 8, reader.offset + 16];
                let times = Times::new(
                    reader.u64("OLE link local update time")?,
                    reader.u64("OLE link local check time")?,
                    reader.u64("OLE link remote update time")?,
                );
                (
                    Some(class_id),
                    Some(class_id_offset),
                    reserved_display,
                    Some(reserved2),
                    Some(times),
                    Some(offsets),
                )
            }
        };

    let tail_offset = reader.offset;
    Ok(Parsed {
        kind,
        flags,
        update_option,
        reserved_moniker,
        relative_source,
        absolute_source,
        class_id,
        class_id_offset,
        reserved_display,
        reserved2,
        times,
        times_offsets,
        tail_offset,
    })
}

fn validate_moniker(
    wire: &[u8],
    range: Option<&Range<usize>>,
    label: &str,
) -> Result<(), OleError> {
    if let Some(moniker_range) = range
        && wire[moniker_range.clone()].len() < 16
    {
        return Err(invalid(format!(
            "OLE link {label} moniker is shorter than its class identifier"
        )));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}
