//! Non-mutating slicer semantic and graph validation.

use super::model::{Definition, Part, Slicers};
use crate::error::Result;
use litchi_opc::OpcPackage;

/// Validate one cache definition.
pub fn definition(value: &Definition) -> Result<()> {
    crate::slicer_cache::validate(value)
}

/// Validate one worksheet slicers collection.
pub fn views(value: &Slicers) -> Result<()> {
    // The canonical writer performs the complete bounded semantic validation
    // without changing the caller's value.
    super::codec::write_views(value).map(|_| ())
}

/// Validate one worksheet part model, including its XML payload.
pub fn part(value: &Part) -> Result<()> {
    views(&value.slicers)
}

/// Validate all feature-owned OPC edges and orphan rules.
pub fn graph(package: &OpcPackage) -> Result<()> {
    super::package::validate_graph(package)
}
