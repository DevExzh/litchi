//! OPC graph ownership for the XLSB Volatile Dependencies part.

use litchi_opc::{OpcPackage, PackURI, TargetMode};

use super::model::ReadLimits;
use crate::package::error::{Error, Result};

/// MS-XLSB 2.1.7.60 content type.
pub const CONTENT_TYPE: &str = "application/vnd.ms-excel.volatileDependencies";
/// MS-XLSB 2.1.7.60 Workbook relationship type.
pub const RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies";
/// Canonical path used when authoring a missing owner. Existing paths are
/// retained exactly because the specification does not require a fixed name.
pub const DEFAULT_PART_NAME: &str = "/xl/volatileDependencies.bin";

/// The Workbook relationship and target part which own one stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Graph {
    pub(crate) workbook_name: PackURI,
    pub(crate) relationship_id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) part_name: PackURI,
    pub(crate) content_type: String,
}

/// Discover and validate the complete Workbook-owned Volatile graph.
pub(crate) fn discover_graph(
    package: &OpcPackage,
    workbook_name: &PackURI,
    limits: ReadLimits,
) -> Result<Option<Graph>> {
    let workbook = package.get_part(workbook_name)?;
    if workbook.content_type() != litchi_opc::constants::content_type::XLSB_BIN {
        return Err(invalid(
            "Volatile Dependencies relationship source is not an XLSB Workbook part",
        ));
    }
    scan_feature_relationships(package, workbook_name, limits)?;
    let mut matches = workbook
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == RELATIONSHIP_TYPE);
    let Some(relationship) = matches.next() else {
        ensure_no_orphan_parts(package, None)?;
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(invalid(
            "Workbook declares multiple Volatile Dependencies relationships",
        ));
    }
    if relationship.target_mode() != TargetMode::Internal || relationship.is_external() {
        return Err(invalid(
            "Volatile Dependencies relationship cannot be external",
        ));
    }
    let part_name = relationship.target_partname()?;
    let part = package.get_part(&part_name)?;
    // Keep the stored spelling as the graph identity. OPC part lookup is
    // case-insensitive, but the spelling is part of the source closure and
    // must be restored by an inverse publication.
    let stored_part_name = part.partname().clone();
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "Volatile Dependencies part '{}' has content type '{}'",
            part_name.as_str(),
            part.content_type()
        )));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Volatile Dependencies part must not have relationships",
        ));
    }
    ensure_no_orphan_parts(package, Some(&stored_part_name))?;
    ensure_exclusive_inbound_relationship(
        package,
        &part_name,
        workbook_name,
        relationship.r_id(),
        limits,
    )?;
    Ok(Some(Graph {
        workbook_name: workbook_name.clone(),
        relationship_id: relationship.r_id().to_string(),
        relationship_type: relationship.reltype().to_string(),
        target_ref: relationship.target_ref().to_string(),
        part_name: stored_part_name,
        content_type: part.content_type().to_string(),
    }))
}

/// Scan every relationship before deciding that the optional owner is absent.
/// A feature relationship on the package root or another part is still an
/// ownership conflict; silently ignoring it would make a graph edit publish
/// against the wrong Workbook owner. The scan is bounded before any graph
/// decision is returned.
fn scan_feature_relationships(
    package: &OpcPackage,
    workbook_name: &PackURI,
    limits: ReadLimits,
) -> Result<()> {
    let mut relationship_count = 0usize;
    for relationship in package.rels().iter() {
        relationship_count = checked_relationship_count(relationship_count, limits)?;
        if relationship.reltype() == RELATIONSHIP_TYPE {
            return Err(invalid(
                "Volatile Dependencies relationship cannot originate at the package root",
            ));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            relationship_count = checked_relationship_count(relationship_count, limits)?;
            if relationship.reltype() == RELATIONSHIP_TYPE && source.partname() != workbook_name {
                return Err(invalid(format!(
                    "Volatile Dependencies relationship on unexpected source part '{}'",
                    source.partname().as_str()
                )));
            }
        }
    }
    Ok(())
}

fn checked_relationship_count(current: usize, limits: ReadLimits) -> Result<usize> {
    let actual = current.checked_add(1).ok_or(Error::CapacityOverflow {
        resource: "Volatile Dependencies relationships",
    })?;
    if actual > limits.max_relationships {
        return Err(Error::LimitExceeded {
            resource: "Volatile Dependencies relationships",
            actual,
            maximum: limits.max_relationships,
        });
    }
    Ok(actual)
}

fn ensure_no_orphan_parts(package: &OpcPackage, expected: Option<&PackURI>) -> Result<()> {
    if package.iter_parts().any(|part| {
        part.content_type() == CONTENT_TYPE
            && expected.is_none_or(|expected| !expected.is_equivalent_to(part.partname()))
    }) {
        return Err(invalid(
            "package contains an orphan or additional Volatile Dependencies part",
        ));
    }
    Ok(())
}

fn ensure_exclusive_inbound_relationship(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_relationship_id: &str,
    limits: ReadLimits,
) -> Result<()> {
    let mut relationship_count = 0usize;
    for relationship in package.rels().iter() {
        relationship_count = checked_relationship_count(relationship_count, limits)?;
        if relationship.target_mode() == TargetMode::Internal
            && relationship.target_partname()?.is_equivalent_to(target)
        {
            return Err(invalid(
                "Volatile Dependencies part has an unexpected package-level relationship",
            ));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            relationship_count = checked_relationship_count(relationship_count, limits)?;
            if relationship.target_mode() == TargetMode::Internal
                && relationship.target_partname()?.is_equivalent_to(target)
                && (source.partname() != expected_source
                    || relationship.r_id() != expected_relationship_id)
            {
                return Err(invalid(
                    "Volatile Dependencies part has an unexpected inbound relationship",
                ));
            }
        }
    }
    Ok(())
}

pub(crate) fn default_part_uri() -> Result<PackURI> {
    PackURI::new(DEFAULT_PART_NAME).map_err(|error| Error::InvalidUri(error.to_string()))
}

pub(crate) fn reject_inbound_relationships(
    package: &OpcPackage,
    target: &PackURI,
    limits: ReadLimits,
) -> Result<()> {
    let mut relationship_count = 0usize;
    for relationship in package.rels().iter() {
        relationship_count = checked_relationship_count(relationship_count, limits)?;
        if relationship.target_mode() == TargetMode::Internal
            && relationship.target_partname()?.is_equivalent_to(target)
        {
            return Err(invalid(format!(
                "package relationship targets Volatile Dependencies part '{}'",
                target.as_str()
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            relationship_count = checked_relationship_count(relationship_count, limits)?;
            if relationship.target_mode() == TargetMode::Internal
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                return Err(invalid(format!(
                    "part '{}' targets Volatile Dependencies part '{}'",
                    source.partname().as_str(),
                    target.as_str()
                )));
            }
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
