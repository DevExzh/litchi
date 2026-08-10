#![allow(
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Error adaptation at the host/formula boundary.

use crate::package::error::{Error, Result};

/// Convert package-owned resolution errors into the standalone formula error
/// vocabulary consumed by the shared formula renderer.
pub(super) fn owner_formula_resolution<T>(result: Result<T>) -> crate::formula::Result<T> {
    result.map_err(|error| match error {
        Error::InvalidFormula(message) => crate::formula::Error::InvalidFormula(message),
        Error::InvalidCellReference(reference) => {
            crate::formula::Error::InvalidCellReference(reference)
        },
        Error::InvalidLength { expected, found } => {
            crate::formula::Error::InvalidLength { expected, found }
        },
        Error::UnsupportedFeature(feature) => crate::formula::Error::UnsupportedFeature(feature),
        Error::Encoding(message) => crate::formula::Error::Encoding(message),
        error => crate::formula::Error::InvalidFormula(error.to_string()),
    })
}
