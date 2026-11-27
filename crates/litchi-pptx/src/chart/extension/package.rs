//! OPC traversal for Microsoft `ChartEx` parts.

use litchi_opc::part::Part as OpcPart;
use litchi_opc::{OpcPackage, PackURI};

use super::model::Document;
use crate::parts::is_relationship_type;
use crate::{Error, Result};

/// Borrowed `ChartEx` package part.
pub struct Part<'a> {
    pub(crate) part: &'a dyn OpcPart,
}

/// Read and validate a `ChartEx` part together with its related resources.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage, part_name: &str) -> Result<Document> {
    let uri = PackURI::new(part_name).map_err(|error| Error::Uri(error.clone()))?;
    let part = Part::from_part(package.get_part(&uri)?)?;
    part.parse_in_package(package)
}

/// Resolve all `ChartEx` parts directly related to a source part.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn related<'a>(package: &'a OpcPackage, source: &dyn OpcPart) -> Result<Vec<Part<'a>>> {
    let mut parts = Vec::new();
    for relationship in source.rels().iter() {
        if relationship.reltype()
            != "http://schemas.microsoft.com/office/2014/relationships/chartEx"
            && !is_relationship_type(
                relationship.reltype(),
                "http://schemas.microsoft.com/office/2014/relationships/chartEx",
                "chartEx",
            )
        {
            continue;
        }
        if relationship.is_external() {
            return Err(Error::Relationship(
                "ChartEx relationships cannot be external".to_owned(),
            ));
        }
        let target = relationship.target_partname()?;
        parts.push(Part::from_part(package.get_part(&target)?)?);
    }
    Ok(parts)
}

/// Parse a `ChartEx` part without opening its referenced package resources.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn read(part: &Part<'_>) -> Result<Document> {
    part.parse()
}
