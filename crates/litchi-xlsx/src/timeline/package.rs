//! Timeline OPC graph ownership.

use super::model::{Cache, Part};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

/// Load and validate all workbook-owned timeline caches.
pub fn load_caches(package: &OpcPackage, workbook: &PackURI) -> Result<Vec<Cache>> {
    crate::timelines::load_timeline_caches(package, workbook)
}

/// Store a complete workbook timeline-cache set atomically at the part layer.
pub fn store_caches(package: &mut OpcPackage, workbook: &PackURI, caches: &[Cache]) -> Result<()> {
    crate::timelines::store_timeline_caches(package, workbook, caches)
}

/// Load and validate every worksheet timelines part.
pub fn load_parts(package: &OpcPackage, workbook: &PackURI) -> Result<Vec<Part>> {
    crate::timelines::load_timelines(package, workbook)
}

/// Store one worksheet-owned timelines part and its relationship.
pub fn store_part(package: &mut OpcPackage, workbook: &PackURI, value: &Part) -> Result<()> {
    crate::timelines::store_worksheet_timelines(package, workbook, value)
}

/// Validate every timeline cache/view edge and orphan rule.
pub fn validate_graph(package: &OpcPackage, workbook: &PackURI) -> Result<()> {
    load_caches(package, workbook)?;
    load_parts(package, workbook)?;
    Ok(())
}
