//! Structural validation for section wire tables.

use crate::package::{Error as PackageError, Result};

/// Validate the strictly increasing CP array of a PlcSed table.
pub(super) fn character_positions(cps: &[u32]) -> Result<()> {
    if cps.windows(2).any(|pair| pair[0] >= pair[1]) {
        return Err(PackageError::Corrupted(
            "PlcfSed character positions must be strictly increasing".to_string(),
        ));
    }
    Ok(())
}
