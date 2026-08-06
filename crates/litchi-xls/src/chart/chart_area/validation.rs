//! Semantic guards for the fixed `Chart` geometry.

use litchi_ograph::chart::Rect;

use crate::{Error, Result};

use super::super::wire::CHART;

/// Validates the constraints shared by `[MS-XLS]` and `[MS-OGRAPH]`.
pub(crate) fn ensure(rect: Rect) -> Result<()> {
    if rect.x != 0 || rect.y != 0 {
        return Err(invalid(
            "Chart origin must be zero for a safe chart-area edit",
        ));
    }
    if rect.width < 0 || rect.height < 0 {
        return Err(invalid(
            "Chart width and height must be nonnegative for a safe edit",
        ));
    }
    Ok(())
}

pub(crate) fn ensure_pair(before: Rect, after: Rect) -> Result<()> {
    ensure(before)?;
    ensure(after)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: CHART,
        message: message.into(),
    }
}
