//! Atomic worksheet graph edits.

use super::{Sheet, validation};
use litchi_core::Result;

/// Clone only the physical worksheet graph, apply one edit, then validate the
/// candidate before it can be published by a builder or mutable facade.
pub(crate) fn edit<F>(source: &[Sheet], operation: F) -> Result<Vec<Sheet>>
where
    F: FnOnce(&mut Vec<Sheet>) -> Result<()>,
{
    let mut candidate = source.to_vec();
    operation(&mut candidate)?;
    validation::validate_sheets(&candidate)?;
    Ok(candidate)
}
