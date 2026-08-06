//! Slicer OPC graph ownership.

use super::model::{Cache, Part};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

/// Load and validate all workbook-owned slicer caches.
pub fn load_caches(package: &OpcPackage) -> Result<Vec<Cache>> {
    crate::slicer_cache::package::load_slicer_caches(package)
}

/// Store one workbook-owned slicer cache atomically at the part layer.
pub fn store_cache(package: &mut OpcPackage, value: &Cache) -> Result<()> {
    crate::slicer_cache::package::store_slicer_cache(package, value)
}

/// Load and validate the slicers part owned by one worksheet.
pub fn load_parts(package: &OpcPackage, worksheet: &PackURI) -> Result<Vec<Part>> {
    crate::slicer_cache::views::load_slicer_parts(package, worksheet)
}

/// Store one worksheet-owned slicers part and its relationship.
pub fn store_part(package: &mut OpcPackage, worksheet: &PackURI, value: &Part) -> Result<()> {
    crate::slicer_cache::views::store_slicer_part(package, worksheet, value)
}

/// Validate every slicer cache and worksheet slicers graph in the package.
pub fn validate_graph(package: &OpcPackage) -> Result<()> {
    load_caches(package)?;
    crate::slicer_cache::views::validate_package_graph(package)?;
    let worksheets: Vec<PackURI> = package
        .iter_parts()
        .filter(|part| part.content_type() == litchi_opc::constants::content_type::SML_WORKSHEET)
        .map(|part| part.partname().clone())
        .collect();
    for worksheet in worksheets {
        load_parts(package, &worksheet)?;
    }
    Ok(())
}
