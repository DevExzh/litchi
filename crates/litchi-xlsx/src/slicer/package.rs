//! Slicer OPC graph ownership.

use super::model::{Cache, Part};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI, SourceLimits};

/// Load and validate all workbook-owned slicer caches.
pub fn load_caches(package: &OpcPackage) -> Result<Vec<Cache>> {
    crate::slicer_cache::package::load_slicer_caches(package)
}

/// Load and validate the slicers part owned by one worksheet.
pub fn load_parts(package: &OpcPackage, worksheet: &PackURI) -> Result<Vec<Part>> {
    crate::slicer_cache::views::load_slicer_parts(package, worksheet)
}

/// Validate every slicer cache and worksheet slicers graph in the package.
pub fn validate_graph(package: &OpcPackage) -> Result<()> {
    let limits = package.read_limits();
    validate_graph_with_limits(package, &limits)
}

pub(crate) fn validate_graph_with_limits(
    package: &OpcPackage,
    limits: &SourceLimits,
) -> Result<()> {
    crate::slicer_cache::package::load_slicer_caches_limited(package, limits)?;
    crate::slicer_cache::views::validate_package_graph(package)?;
    let worksheets: Vec<PackURI> = package
        .iter_parts()
        .filter(|part| part.content_type() == litchi_opc::constants::content_type::SML_WORKSHEET)
        .map(|part| part.partname().clone())
        .collect();
    for worksheet in worksheets {
        crate::slicer_cache::views::load_slicer_parts_limited(package, &worksheet, limits)?;
    }
    Ok(())
}
