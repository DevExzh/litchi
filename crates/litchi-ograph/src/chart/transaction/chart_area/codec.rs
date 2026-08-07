//! Lossless fixed-size `Chart` record patching.

use super::validation;
use crate::chart::{Chart, Rect};
use crate::{Error, Result};
use litchi_biff::{Kind, Records};

const CHART: Kind = Kind::from_wire(0x1002);
const PAYLOAD_BYTES: usize = 16;

/// Finds and validates the source `Chart` record without mutating it.
pub(crate) fn locate(chart: &Chart, expected: Rect) -> Result<usize> {
    validation::ensure(expected)?;
    let source = match &chart.origin {
        crate::chart::model::Origin::Parsed(stream) => stream.as_bytes(),
        crate::chart::model::Origin::Fresh => {
            return Err(Error::UnsupportedMutation {
                operation: "chart-area-patch",
                reason: "the source chart is not a parsed replayable stream",
            });
        },
    };
    let mut offset = None;
    for item in Records::new(source) {
        let record = item?;
        if record.kind() != CHART {
            continue;
        }
        if offset.is_some() {
            return Err(Error::UnsupportedMutation {
                operation: "chart-area-patch",
                reason: "source Chart record is duplicated",
            });
        }
        let data = crate::record::payload_bytes(CHART, record.payload(), PAYLOAD_BYTES)?;
        let actual = Rect {
            x: i32::from_le_bytes(data[0..4].try_into().ok().ok_or(Error::SizeOverflow {
                resource: "Chart rectangle",
            })?),
            y: i32::from_le_bytes(data[4..8].try_into().ok().ok_or(Error::SizeOverflow {
                resource: "Chart rectangle",
            })?),
            width: i32::from_le_bytes(data[8..12].try_into().ok().ok_or(Error::SizeOverflow {
                resource: "Chart rectangle",
            })?),
            height: i32::from_le_bytes(data[12..16].try_into().ok().ok_or({
                Error::SizeOverflow {
                    resource: "Chart rectangle",
                }
            })?),
        };
        if actual != expected {
            return Err(Error::UnsupportedMutation {
                operation: "chart-area-patch",
                reason: "source Chart record does not match the snapshot rectangle",
            });
        }
        offset = Some(record.offset());
    }
    offset.ok_or(Error::UnsupportedMutation {
        operation: "chart-area-patch",
        reason: "source Chart record is missing",
    })
}

/// Replaces only the 16-byte `Chart` payload after source validation.
pub(crate) fn patch(chart: &mut Chart, before: Rect, after: Rect) -> Result<()> {
    validation::ensure_pair(before, after)?;
    let offset = locate(chart, before)?;
    let payload_start = offset.checked_add(4).ok_or(Error::SizeOverflow {
        resource: "Chart rectangle offset",
    })?;
    let payload_end = payload_start
        .checked_add(PAYLOAD_BYTES)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle payload",
        })?;
    let bytes = match &mut chart.origin {
        crate::chart::model::Origin::Parsed(stream) => stream.as_bytes_mut(),
        crate::chart::model::Origin::Fresh => {
            return Err(Error::UnsupportedMutation {
                operation: "chart-area-patch",
                reason: "the source chart is not a parsed replayable stream",
            });
        },
    };
    let target = bytes
        .get_mut(payload_start..payload_end)
        .ok_or(Error::UnsupportedMutation {
            operation: "chart-area-patch",
            reason: "source Chart payload falls outside the source stream",
        })?;
    target[0..4].copy_from_slice(&after.x.to_le_bytes());
    target[4..8].copy_from_slice(&after.y.to_le_bytes());
    target[8..12].copy_from_slice(&after.width.to_le_bytes());
    target[12..16].copy_from_slice(&after.height.to_le_bytes());
    Ok(())
}
