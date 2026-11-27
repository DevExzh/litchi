//! Timeline OPC graph ownership.

use super::model::{Cache, Part};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI, SourceLimits};

/// Load and validate all workbook-owned timeline caches.
pub fn load_caches(package: &OpcPackage, workbook: &PackURI) -> Result<Vec<Cache>> {
    crate::timelines::load_timeline_caches(package, workbook)
}

/// Load and validate every worksheet timelines part.
pub fn load_parts(package: &OpcPackage, workbook: &PackURI) -> Result<Vec<Part>> {
    crate::timelines::load_timelines(package, workbook)
}

/// Validate every timeline cache/view edge and orphan rule.
pub fn validate_graph(package: &OpcPackage, workbook: &PackURI) -> Result<()> {
    let limits = package.read_limits();
    validate_graph_with_limits(package, workbook, &limits)
}

pub(crate) fn validate_graph_with_limits(
    package: &OpcPackage,
    workbook: &PackURI,
    limits: &SourceLimits,
) -> Result<()> {
    crate::timelines::load_timeline_caches_limited(package, workbook, limits)?;
    crate::timelines::load_timelines_limited(package, workbook, limits)?;
    Ok(())
}
