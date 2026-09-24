//! Structural validation for the source-backed ink model.

use super::{Document, ValueError};

/// Validate one immutable InkML projection.
///
/// # Errors
///
/// Returns an error when a source span or bounded collection is invalid.
pub(crate) fn validate(document: &Document) -> Result<(), ValueError> {
    document.validate()
}
