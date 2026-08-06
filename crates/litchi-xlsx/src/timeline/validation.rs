//! Non-mutating timeline semantic and graph validation.

use super::model::{CacheDefinition, Part, Views};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

/// Validate one timeline cache definition.
pub fn cache(value: &CacheDefinition) -> Result<()> {
    super::codec::write_cache(value).map(|_| ())
}

/// Validate one worksheet timelines collection.
pub fn views(value: &Views) -> Result<()> {
    super::codec::write_views(value).map(|_| ())
}

/// Validate one worksheet part model, including its XML payload.
pub fn part(value: &Part) -> Result<()> {
    views(&value.timelines)
}

/// Validate all timeline package edges and orphan rules.
pub fn graph(package: &OpcPackage, workbook: &PackURI) -> Result<()> {
    super::package::validate_graph(package, workbook)
}
