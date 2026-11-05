//! Cross-field validation for the cell-watch owner.

use super::model::{MAX_COLUMN, MAX_ROW, Reference, UnknownRecord, Watches};
use crate::package::error::{Error, Result};
use std::collections::HashSet;

pub(crate) fn invalid(typ: &'static str, value: &'static str) -> Error {
    Error::Unrecognized { typ, val: value }
}

pub(crate) fn reference(value: &Reference) -> Result<()> {
    if value.row() > MAX_ROW {
        return Err(invalid("BrtCellWatch row", "exceeds"));
    }
    if value.column() > MAX_COLUMN {
        return Err(invalid("BrtCellWatch column", "exceeds"));
    }
    Ok(())
}

pub(crate) fn watches(value: &Watches) -> Result<()> {
    if value.len() > super::MAX_WATCHES {
        return Err(Error::InvalidLength {
            expected: super::MAX_WATCHES,
            found: value.len(),
        });
    }
    let mut references = HashSet::with_capacity(value.len());
    for watch in value.iter() {
        reference(&watch.reference())?;
        if !references.insert(watch.reference()) {
            return Err(invalid("BrtCellWatch collection", "duplicate cell (, )"));
        }
    }
    Ok(())
}

pub(crate) fn unknown(value: &UnknownRecord) -> Result<()> {
    if value.kind() > 0x3fff {
        return Err(invalid("opaque BIFF12 record kind", "0x"));
    }
    if value.payload().len() > super::MAX_OPAQUE_PAYLOAD {
        return Err(Error::InvalidLength {
            expected: super::MAX_OPAQUE_PAYLOAD,
            found: value.payload().len(),
        });
    }
    Ok(())
}
