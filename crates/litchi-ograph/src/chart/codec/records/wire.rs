//! Shared BIFF framing, scalar access, and bounded allocation helpers for chart records.

use super::super::super::model::Count;
use crate::limits::as_u64;
use crate::{Error, Result};
use litchi_biff::{Encoder, Kind as RecordKind, RecordRef};

pub(super) const CONTINUE: RecordKind = RecordKind::from_wire(0x003C);
pub(super) const SCL: RecordKind = RecordKind::from_wire(0x00A0);
pub(super) const DIMENSIONS: RecordKind = RecordKind::from_wire(0x0200);
pub(super) const GRAPH_BLANK: RecordKind = RecordKind::from_wire(0x0001);
pub(super) const GRAPH_NUMBER: RecordKind = RecordKind::from_wire(0x0003);
pub(super) const EXCEL_BLANK: RecordKind = RecordKind::from_wire(0x0201);
pub(super) const EXCEL_NUMBER: RecordKind = RecordKind::from_wire(0x0203);
pub(super) const EXCEL_BOOL_ERR: RecordKind = RecordKind::from_wire(0x0205);
pub(super) const CELL_LABEL: RecordKind = RecordKind::from_wire(0x0204);
pub(super) const DATA_LAB_EXT: RecordKind = RecordKind::from_wire(0x086A);
pub(super) const DATA_LAB_EXT_CONTENTS: RecordKind = RecordKind::from_wire(0x086B);
pub(super) const CHART_REC: RecordKind = RecordKind::from_wire(0x1002);
pub(super) const SERIES: RecordKind = RecordKind::from_wire(0x1003);
pub(super) const DATA_FORMAT: RecordKind = RecordKind::from_wire(0x1006);
pub(super) const LINE_FORMAT: RecordKind = RecordKind::from_wire(0x1007);
pub(super) const MARKER_FORMAT: RecordKind = RecordKind::from_wire(0x1009);
pub(super) const AREA_FORMAT: RecordKind = RecordKind::from_wire(0x100A);
pub(super) const PIE_FORMAT: RecordKind = RecordKind::from_wire(0x100B);
pub(super) const SERIES_TEXT: RecordKind = RecordKind::from_wire(0x100D);
pub(super) const CHART_FORMAT: RecordKind = RecordKind::from_wire(0x1014);
pub(super) const LEGEND: RecordKind = RecordKind::from_wire(0x1015);
pub(super) const SERIES_LIST: RecordKind = RecordKind::from_wire(0x1016);
pub(super) const BAR: RecordKind = RecordKind::from_wire(0x1017);
pub(super) const LINE: RecordKind = RecordKind::from_wire(0x1018);
pub(super) const PIE: RecordKind = RecordKind::from_wire(0x1019);
pub(super) const AREA: RecordKind = RecordKind::from_wire(0x101A);
pub(super) const SCATTER: RecordKind = RecordKind::from_wire(0x101B);
pub(super) const CRT_LINE: RecordKind = RecordKind::from_wire(0x101C);
pub(super) const AXIS: RecordKind = RecordKind::from_wire(0x101D);
pub(super) const TICK: RecordKind = RecordKind::from_wire(0x101E);
pub(super) const VALUE_RANGE: RecordKind = RecordKind::from_wire(0x101F);
pub(super) const CAT_SER_RANGE: RecordKind = RecordKind::from_wire(0x1020);
pub(super) const AXIS_LINE: RecordKind = RecordKind::from_wire(0x1021);
pub(super) const CRT_LINK: RecordKind = RecordKind::from_wire(0x1022);
pub(super) const DEFAULT_TEXT: RecordKind = RecordKind::from_wire(0x1024);
pub(super) const TEXT: RecordKind = RecordKind::from_wire(0x1025);
pub(super) const FONT_X: RecordKind = RecordKind::from_wire(0x1026);
pub(super) const OBJECT_LINK: RecordKind = RecordKind::from_wire(0x1027);
pub(super) const FRAME: RecordKind = RecordKind::from_wire(0x1032);
pub(super) const BEGIN: RecordKind = RecordKind::from_wire(0x1033);
pub(super) const END: RecordKind = RecordKind::from_wire(0x1034);
pub(crate) const PLOT_AREA: RecordKind = RecordKind::from_wire(0x1035);
pub(super) const DROP_BAR: RecordKind = RecordKind::from_wire(0x103D);
pub(super) const RADAR: RecordKind = RecordKind::from_wire(0x103E);
pub(super) const SURFACE: RecordKind = RecordKind::from_wire(0x103F);
pub(super) const RADAR_AREA: RecordKind = RecordKind::from_wire(0x1040);
pub(super) const AXIS_PARENT: RecordKind = RecordKind::from_wire(0x1041);
pub(crate) const SHT_PROPS: RecordKind = RecordKind::from_wire(0x1044);
pub(super) const SER_TO_CRT: RecordKind = RecordKind::from_wire(0x1045);
pub(super) const AXES_USED: RecordKind = RecordKind::from_wire(0x1046);
pub(super) const SER_PARENT: RecordKind = RecordKind::from_wire(0x104A);
pub(super) const SER_AUX_TREND: RecordKind = RecordKind::from_wire(0x104B);
pub(super) const POS: RecordKind = RecordKind::from_wire(0x104F);
pub(super) const BRAI: RecordKind = RecordKind::from_wire(0x1051);
pub(super) const SER_AUX_ERR_BAR: RecordKind = RecordKind::from_wire(0x105B);
pub(super) const PLOT_GROWTH: RecordKind = RecordKind::from_wire(0x1064);
pub(super) const SI_INDEX: RecordKind = RecordKind::from_wire(0x1065);

pub(super) fn push_record(out: &mut Encoder, kind: RecordKind, payload: &[u8]) -> Result<()> {
    out.push(kind, payload)?;
    Ok(())
}

pub(super) fn count_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<Count> {
    Count::new(u16_at(data, offset, record)?).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "series point count exceeds 32,767",
    })
}

pub(super) fn byte_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<u8> {
    data.get(offset).copied().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record scalar is truncated",
    })
}

pub(super) fn u16_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<u16> {
    let bytes = data
        .get(
            offset..offset.checked_add(2).ok_or(Error::SizeOverflow {
                resource: "record scalar",
            })?,
        )
        .ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "record u16 is truncated",
        })?;
    let lo = bytes.first().copied().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record u16 is truncated",
    })?;
    let hi = bytes.get(1).copied().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record u16 is truncated",
    })?;
    Ok(u16::from_le_bytes([lo, hi]))
}

pub(super) fn i16_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<i16> {
    Ok(i16::from_le_bytes(
        u16_at(data, offset, record)?.to_le_bytes(),
    ))
}

pub(super) fn u32_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<u32> {
    let bytes = data
        .get(
            offset..offset.checked_add(4).ok_or(Error::SizeOverflow {
                resource: "record scalar",
            })?,
        )
        .ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "record u32 is truncated",
        })?;
    let array = <[u8; 4]>::try_from(bytes).ok().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record u32 is truncated",
    })?;
    Ok(u32::from_le_bytes(array))
}

pub(super) fn i32_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<i32> {
    Ok(i32::from_le_bytes(
        u32_at(data, offset, record)?.to_le_bytes(),
    ))
}

pub(super) fn f64_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<f64> {
    let bytes = data
        .get(
            offset..offset.checked_add(8).ok_or(Error::SizeOverflow {
                resource: "record scalar",
            })?,
        )
        .ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "record f64 is truncated",
        })?;
    let array = <[u8; 8]>::try_from(bytes).ok().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record f64 is truncated",
    })?;
    Ok(f64::from_le_bytes(array))
}

pub(super) fn array4_at(data: &[u8], offset: usize, record: RecordRef<'_>) -> Result<[u8; 4]> {
    let bytes = data
        .get(
            offset..offset.checked_add(4).ok_or(Error::SizeOverflow {
                resource: "record array",
            })?,
        )
        .ok_or(Error::InvalidChart {
            offset: record.offset(),
            reason: "record array is truncated",
        })?;
    <[u8; 4]>::try_from(bytes).ok().ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "record array is truncated",
    })
}

pub(super) fn exact(record: RecordRef<'_>, expected: usize) -> Result<()> {
    if record.payload().len() != expected {
        return invalid(record, "record payload has an invalid fixed length");
    }
    Ok(())
}

pub(super) fn invalid<T>(record: RecordRef<'_>, reason: &'static str) -> Result<T> {
    Err(Error::InvalidChart {
        offset: record.offset(),
        reason,
    })
}

pub(super) fn invalid_model<T>(field: &'static str, reason: &'static str) -> Result<T> {
    Err(Error::InvalidModel { field, reason })
}

pub(super) fn limit<T>(resource: &'static str, observed: usize, maximum: usize) -> Result<T> {
    Err(Error::LimitExceeded {
        resource,
        observed: as_u64(observed),
        maximum: as_u64(maximum),
    })
}

pub(super) fn check_add(current: usize, maximum: usize, resource: &'static str) -> Result<()> {
    let observed = current
        .checked_add(1)
        .ok_or(Error::SizeOverflow { resource })?;
    if observed > maximum {
        return limit(resource, observed, maximum);
    }
    Ok(())
}

pub(super) fn push<T>(values: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .ok()
        .ok_or(Error::Allocation { resource })?;
    values.push(value);
    Ok(())
}

pub(super) fn copy(data: &[u8], resource: &'static str, maximum: usize) -> Result<Vec<u8>> {
    if data.len() > maximum {
        return limit(resource, data.len(), maximum);
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(data.len())
        .ok()
        .ok_or(Error::Allocation { resource })?;
    output.extend_from_slice(data);
    Ok(output)
}

pub(super) fn vec_with_capacity(capacity: usize, resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .ok()
        .ok_or(Error::Allocation { resource })?;
    Ok(output)
}

pub(super) fn put_slice(
    output: &mut [u8],
    offset: usize,
    value: &[u8],
    resource: &'static str,
) -> Result<()> {
    let end = offset
        .checked_add(value.len())
        .ok_or(Error::SizeOverflow { resource })?;
    output
        .get_mut(offset..end)
        .ok_or(Error::SizeOverflow { resource })?
        .copy_from_slice(value);
    Ok(())
}

pub(super) fn put_byte(output: &mut [u8], offset: usize, value: u8) -> Result<()> {
    *output.get_mut(offset).ok_or(Error::SizeOverflow {
        resource: "encoded scalar",
    })? = value;
    Ok(())
}

pub(super) fn put_u16(output: &mut [u8], offset: usize, value: u16) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded u16")
}

pub(super) fn put_u32(output: &mut [u8], offset: usize, value: u32) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded u32")
}

pub(super) fn put_i16(output: &mut [u8], offset: usize, value: i16) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded i16")
}

pub(super) fn put_i32(output: &mut [u8], offset: usize, value: i32) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded i32")
}

pub(super) fn put_f64(output: &mut [u8], offset: usize, value: f64) -> Result<()> {
    put_slice(output, offset, &value.to_le_bytes(), "encoded f64")
}
