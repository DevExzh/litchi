//! `[MS-OGRAPH]` chart-area mutation guards.

use crate::chart::Rect;
use crate::{Error, Result};

/// Validates the `Chart` record's semantic fixed-point constraints.
pub(super) fn ensure(value: Rect) -> Result<()> {
    if value.x != 0 || value.y != 0 {
        return Err(Error::InvalidModel {
            field: "chart area",
            reason: "Chart x and y must be zero",
        });
    }
    if value.width < 0 || value.height < 0 {
        return Err(Error::InvalidModel {
            field: "chart area",
            reason: "Chart width and height must be nonnegative",
        });
    }
    Ok(())
}

/// Validates both the retained source value and the requested replacement.
pub(crate) fn ensure_pair(before: Rect, after: Rect) -> Result<()> {
    ensure(before)?;
    ensure(after)
}
