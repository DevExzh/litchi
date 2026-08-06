//! `[MS-OGRAPH]` `ShtProps` mutation guards.

use crate::chart::Props;
use crate::{Error, Result};

/// Validates the complete typed four-byte `ShtProps` value.
pub(super) fn ensure(value: Props) -> Result<()> {
    if crate::chart::codec::valid_props(value.flags) {
        Ok(())
    } else {
        Err(Error::InvalidModel {
            field: "sheet properties",
            reason: "ShtProps flags, blank mode, or plot-area policy is invalid",
        })
    }
}

/// Validates a source/replacement pair without allowing `PlotArea` topology
/// to change.  That record is zero-sized but structural: changing its
/// presence would add or remove a record rather than patch an existing fixed
/// payload.
pub(crate) fn ensure_pair(before: Props, after: Props) -> Result<()> {
    ensure(before)?;
    ensure(after)?;
    if before.plot_area != after.plot_area {
        return Err(Error::UnsupportedMutation {
            operation: "sheet-props-patch",
            reason: "PlotArea record presence cannot change in a fixed-record transaction",
        });
    }
    Ok(())
}
