//! Cross-field validation for the cell-watch owner.

use super::model::{MAX_COLUMN, MAX_ROW, Reference, UnknownRecord, Watches};
use crate::package::error::{Error, Result};
use std::collections::HashSet;

pub(crate) fn invalid(typ: impl Into<String>, value: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.into(),
        val: value.into(),
    }
}

pub(crate) fn reference(value: &Reference) -> Result<()> {
    if value.row() > MAX_ROW {
        return Err(invalid(
            "BrtCellWatch row",
            format!("{} exceeds {MAX_ROW}", value.row()),
        ));
    }
    if value.column() > MAX_COLUMN {
        return Err(invalid(
            "BrtCellWatch column",
            format!("{} exceeds {MAX_COLUMN}", value.column()),
        ));
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
            return Err(invalid(
                "BrtCellWatch collection",
                format!("duplicate cell ({}, {})", watch.row(), watch.column()),
            ));
        }
    }
    Ok(())
}

pub(crate) fn unknown(value: &UnknownRecord) -> Result<()> {
    if value.kind() > 0x3fff {
        return Err(invalid(
            "opaque BIFF12 record kind",
            format!("0x{:04X}", value.kind()),
        ));
    }
    if value.payload().len() > super::MAX_OPAQUE_PAYLOAD {
        return Err(Error::InvalidLength {
            expected: super::MAX_OPAQUE_PAYLOAD,
            found: value.payload().len(),
        });
    }
    Ok(())
}
