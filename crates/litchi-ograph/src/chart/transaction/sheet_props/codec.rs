//! Lossless fixed-size `ShtProps` patching.

use super::validation;
use crate::chart::{Chart, Props};
use crate::{Error, Result};
use litchi_biff::Records;

const PAYLOAD_BYTES: usize = 4;

/// Finds and validates the unique source `ShtProps` record without mutating.
pub(crate) fn locate(chart: &Chart, expected: Props) -> Result<usize> {
    validation::ensure(expected)?;
    let source = match &chart.origin {
        crate::chart::model::Origin::Parsed(stream) => stream.as_bytes(),
        crate::chart::model::Origin::Fresh => {
            return Err(Error::UnsupportedMutation {
                operation: "sheet-props-patch",
                reason: "the source chart is not a parsed replayable stream",
            });
        },
    };

    let mut props_offset = None;
    let mut plot_area_count = 0usize;
    for item in Records::new(source) {
        let record = item?;
        if record.kind() == super::super::super::codec::SHT_PROPS {
            if props_offset.is_some() {
                return Err(Error::UnsupportedMutation {
                    operation: "sheet-props-patch",
                    reason: "source ShtProps record is duplicated",
                });
            }
            let data = crate::record::payload_bytes(
                super::super::super::codec::SHT_PROPS,
                record.payload(),
                PAYLOAD_BYTES,
            )?;
            let flags = u32::from_le_bytes(data.try_into().ok().ok_or(Error::SizeOverflow {
                resource: "ShtProps flags",
            })?);
            if flags != expected.flags {
                return Err(Error::UnsupportedMutation {
                    operation: "sheet-props-patch",
                    reason: "source ShtProps value does not match the snapshot",
                });
            }
            props_offset = Some(record.offset());
        } else if record.kind() == super::super::super::codec::PLOT_AREA {
            crate::record::payload_bytes(
                super::super::super::codec::PLOT_AREA,
                record.payload(),
                0,
            )?;
            plot_area_count = plot_area_count.checked_add(1).ok_or(Error::SizeOverflow {
                resource: "PlotArea record count",
            })?;
        }
    }

    if plot_area_count > 1 {
        return Err(Error::UnsupportedMutation {
            operation: "sheet-props-patch",
            reason: "source PlotArea record is duplicated",
        });
    }
    if (plot_area_count == 1) != expected.plot_area {
        return Err(Error::UnsupportedMutation {
            operation: "sheet-props-patch",
            reason: "source PlotArea presence does not match the snapshot",
        });
    }
    props_offset.ok_or(Error::UnsupportedMutation {
        operation: "sheet-props-patch",
        reason: "source ShtProps record is missing",
    })
}

/// Replaces only the four-byte `ShtProps` payload after complete preflight.
pub(crate) fn patch(
    chart: &mut Chart,
    before: Props,
    after: Props,
    expected_offset: usize,
) -> Result<()> {
    validation::ensure_pair(before, after)?;
    let offset = locate(chart, before)?;
    if offset != expected_offset {
        return Err(Error::UnsupportedMutation {
            operation: "sheet-props-patch",
            reason: "source ShtProps record identity does not match the patch",
        });
    }
    let payload_start = offset.checked_add(4).ok_or(Error::SizeOverflow {
        resource: "ShtProps offset",
    })?;
    let payload_end = payload_start
        .checked_add(PAYLOAD_BYTES)
        .ok_or(Error::SizeOverflow {
            resource: "ShtProps payload",
        })?;
    let bytes = match &mut chart.origin {
        crate::chart::model::Origin::Parsed(stream) => stream.as_bytes_mut(),
        crate::chart::model::Origin::Fresh => {
            return Err(Error::UnsupportedMutation {
                operation: "sheet-props-patch",
                reason: "the source chart is not a parsed replayable stream",
            });
        },
    };
    let target = bytes
        .get_mut(payload_start..payload_end)
        .ok_or(Error::UnsupportedMutation {
            operation: "sheet-props-patch",
            reason: "source ShtProps payload falls outside the source stream",
        })?;
    target.copy_from_slice(&after.flags.to_le_bytes());
    Ok(())
}
