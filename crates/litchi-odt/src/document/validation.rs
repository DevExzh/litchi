//! Safety limits shared by document XML codecs and package validation.

use litchi_core::{Error, Result};

const MAX_REFERENCE_DEPTH: usize = 4_096;
const MAX_REFERENCES: usize = 1_000_000;

pub(super) fn checked_reference_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODF reference nesting depth overflow".to_string()))?;
    if depth > MAX_REFERENCE_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "ODF reference nesting exceeds {MAX_REFERENCE_DEPTH} levels"
        )));
    }
    Ok(depth)
}

pub(super) fn ensure_reference_capacity(length: usize, kind: &str) -> Result<()> {
    if length >= MAX_REFERENCES {
        return Err(Error::InvalidFormat(format!(
            "document exceeds {MAX_REFERENCES} {kind}"
        )));
    }
    Ok(())
}
