//! OPC graph ownership for the Calculation Chain part.

use litchi_opc::{OpcPackage, PackURI, TargetMode};

use super::model::ReadLimits;
use crate::package::error::{Error, Result};

/// MS-XLSB 2.1.7.4 content type.
pub const CONTENT_TYPE: &str = "application/vnd.ms-excel.calcChain";
/// MS-XLSB 2.1.7.4 Workbook relationship type.
pub const RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

/// The Workbook relationship and target part which own one opaque stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Graph {
    pub(crate) workbook_name: PackURI,
    pub(crate) relationship_id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) part_name: PackURI,
    pub(crate) content_type: String,
}

/// Discover the owner while also proving that `probe_target` has no hidden
/// inbound relationship when the Calculation Chain owner is absent. The
/// optional probe is used by inverse publication so one bounded relationship
/// pass covers both feature ownership and target topology.
pub(crate) fn discover_graph_with_probe(
    package: &OpcPackage,
    limits: ReadLimits,
    probe_target: Option<&PackURI>,
) -> Result<Option<Graph>> {
    let limits = limits.validate()?;
    // This pass must precede every owned graph value and package clone.  An
    // authored package can contain relationship/part metadata that has never
    // crossed the bounded OPC reader, so the ordinary in-memory clone is an
    // allocation boundary in its own right.  Keep this pass borrowed and
    // check the same finite OPC metadata ceilings used by source capture.
    preflight_metadata(package, limits)?;
    let workbook = package.main_document_part()?;
    let workbook_name = workbook.partname();
    if workbook.content_type() != litchi_opc::constants::content_type::XLSB_BIN {
        return Err(invalid(
            "Calculation Chain relationship source is not an XLSB Workbook part",
        ));
    }
    let mut matches = workbook
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == RELATIONSHIP_TYPE);
    let Some(relationship) = matches.next() else {
        scan_relationship_graph(package, workbook_name, probe_target, None, None, limits)?;
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(invalid(
            "Workbook declares multiple Calculation Chain relationships",
        ));
    }
    if relationship.target_mode() != TargetMode::Internal || relationship.is_external() {
        return Err(invalid("Calculation Chain relationship cannot be external"));
    }

    let part_name = relationship.target_partname()?;
    let part = package.get_part(&part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(Error::InvalidContentType {
            expected: CONTENT_TYPE.to_string(),
            got: part.content_type().to_string(),
        });
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Calculation Chain part must not have relationships",
        ));
    }
    // All ownership and inbound checks run while the package metadata is
    // still borrowed.  Only after this complete graph preflight succeeds do
    // we retain the lexical tokens needed by a source-bound patch.
    let stored_part_name = part.partname();
    scan_relationship_graph(
        package,
        workbook_name,
        Some(stored_part_name),
        Some(workbook_name),
        Some(relationship.r_id()),
        limits,
    )?;

    Ok(Some(Graph {
        workbook_name: workbook_name.clone(),
        relationship_id: relationship.r_id().to_string(),
        relationship_type: relationship.reltype().to_string(),
        target_ref: relationship.target_ref().to_string(),
        part_name: stored_part_name.clone(),
        content_type: part.content_type().to_string(),
    }))
}

/// Preflight the modeled package metadata without cloning any part, URI, or
/// relationship field. `OpcPackage::clone` deep-copies these fields even
/// though part blobs are Arc-backed, so an authored package must be checked
/// before the Calculation Chain snapshot retains its package image. The
/// resulting totals are a conservative metadata budget, not byte-exact
/// allocator telemetry: retained source-map keys are covered by the repeated
/// URI terms, while opaque source members are bounded by OPC ingress and
/// source-capture limits.
pub(crate) fn preflight_metadata(package: &OpcPackage, limits: ReadLimits) -> Result<()> {
    // Keep this metadata pass independent of the per-stream byte ceiling.
    // `max_part_bytes` is enforced against the Workbook/owner and captured
    // source tokens after graph discovery; using it here would make a small
    // stream limit reject ordinary package metadata before the more useful
    // stream-specific error is reported.  The OPC metadata profile itself is
    // still finite and is the same default profile used for in-memory
    // authored metadata.
    let opc_limits = litchi_opc::ReadLimits::default();
    let part_count = package.part_count();
    if part_count > limits.max_graph_parts {
        return Err(Error::LimitExceeded {
            resource: "Calculation Chain graph parts",
            actual: part_count,
            maximum: limits.max_graph_parts,
        });
    }

    let mut relationship_count = 0usize;
    let mut relationship_metadata_bytes = 0usize;
    preflight_relationship_collection(
        package.rels(),
        &opc_limits,
        &mut relationship_count,
        &mut relationship_metadata_bytes,
        limits,
    )?;

    let mut content_type_metadata_bytes = 0usize;
    for part in package.iter_parts() {
        preflight_xml_attribute(
            "PartName",
            part.partname().as_str(),
            opc_limits.max_xml_attribute_bytes(),
        )?;
        preflight_xml_attribute(
            "ContentType",
            part.content_type(),
            opc_limits.max_xml_attribute_bytes(),
        )?;
        content_type_metadata_bytes = checked_metadata_add(
            content_type_metadata_bytes,
            // `OpcPackage::clone` owns both the HashMap key and the part's
            // PackURI, in addition to its content-type string.  Charge both
            // URI copies before the clone boundary.
            metadata_cost([
                part.partname().as_str().len(),
                part.partname().as_str().len(),
                part.content_type().len(),
            ])?,
            "Calculation Chain content-type metadata",
        )?;
        if content_type_metadata_bytes > opc_limits.max_content_types_bytes() {
            return Err(opc_limit(
                litchi_opc::ReadResource::ContentTypesBytes,
                content_type_metadata_bytes,
                opc_limits.max_content_types_bytes(),
            ));
        }
        preflight_relationship_collection(
            part.rels(),
            &opc_limits,
            &mut relationship_count,
            &mut relationship_metadata_bytes,
            limits,
        )?;
    }
    Ok(())
}

fn preflight_relationship_collection(
    relationships: &litchi_opc::Relationships,
    opc_limits: &litchi_opc::ReadLimits,
    relationship_count: &mut usize,
    relationship_metadata_bytes: &mut usize,
    limits: ReadLimits,
) -> Result<()> {
    // A cloned Relationships owns these collection-level strings even when
    // the collection is empty.  They are distinct from the relationship
    // fields below and must be charged without assuming that an owner URI is
    // equal to its physical part name.
    let collection_metadata = metadata_cost(
        std::iter::once(relationships.base_uri().len())
            .chain(relationships.source_uri().into_iter().map(str::len)),
    )?;
    *relationship_metadata_bytes = checked_metadata_add(
        *relationship_metadata_bytes,
        collection_metadata,
        "Calculation Chain relationship owner metadata",
    )?;
    if *relationship_metadata_bytes > opc_limits.max_total_relationship_xml_bytes() {
        return Err(opc_limit(
            litchi_opc::ReadResource::TotalRelationshipXmlBytes,
            *relationship_metadata_bytes,
            opc_limits.max_total_relationship_xml_bytes(),
        ));
    }
    for relationship in relationships.iter() {
        preflight_relationship(
            relationship,
            opc_limits,
            relationship_count,
            relationship_metadata_bytes,
            limits,
        )?;
    }
    Ok(())
}

fn preflight_relationship(
    relationship: &litchi_opc::Relationship,
    opc_limits: &litchi_opc::ReadLimits,
    relationship_count: &mut usize,
    relationship_metadata_bytes: &mut usize,
    limits: ReadLimits,
) -> Result<()> {
    *relationship_count = relationship_count
        .checked_add(1)
        .ok_or(Error::CapacityOverflow {
            resource: "Calculation Chain relationships",
        })?;
    if *relationship_count > limits.max_relationships {
        return Err(Error::LimitExceeded {
            resource: "Calculation Chain relationships",
            actual: *relationship_count,
            maximum: limits.max_relationships,
        });
    }

    preflight_xml_attribute(
        "Id",
        relationship.r_id(),
        opc_limits.max_xml_attribute_bytes(),
    )?;
    preflight_xml_attribute(
        "Type",
        relationship.reltype(),
        opc_limits.max_xml_attribute_bytes(),
    )?;
    preflight_xml_attribute(
        "Target",
        relationship.target_ref(),
        opc_limits.max_xml_attribute_bytes(),
    )?;
    preflight_metadata_text(
        litchi_opc::ReadResource::RelationshipTargetBytes,
        relationship.target_ref().len(),
        opc_limits.max_relationship_target_bytes(),
    )?;
    // `target_partname` resolves a relative target against the source URI and
    // allocates a new PackURI.  Bound that allocation conservatively before
    // the graph code performs it.  External targets are still bounded by the
    // raw target checks above and do not need a resolved URI.
    if relationship.target_mode() == TargetMode::Internal {
        let resolved_upper_bound = relationship
            .base_uri()
            .len()
            .checked_add(relationship.target_ref().len())
            .ok_or(Error::CapacityOverflow {
                resource: "Calculation Chain relationship target metadata",
            })?;
        preflight_metadata_text(
            litchi_opc::ReadResource::XmlAttributeBytes,
            resolved_upper_bound,
            opc_limits.max_xml_attribute_bytes(),
        )?;
    }

    *relationship_metadata_bytes = checked_metadata_add(
        *relationship_metadata_bytes,
        // The HashMap key and Relationship::r_id are separate owned strings;
        // base/source are also owned by every edge.  Count each actual field,
        // including the optional source URI, before OpcPackage::clone.
        metadata_cost(
            std::iter::once(relationship.r_id().len())
                .chain(std::iter::once(relationship.r_id().len()))
                .chain(std::iter::once(relationship.reltype().len()))
                .chain(std::iter::once(relationship.target_ref().len()))
                .chain(std::iter::once(relationship.base_uri().len()))
                .chain(relationship.source_uri().into_iter().map(str::len)),
        )?,
        "Calculation Chain relationship metadata",
    )?;
    if *relationship_metadata_bytes > opc_limits.max_total_relationship_xml_bytes() {
        return Err(opc_limit(
            litchi_opc::ReadResource::TotalRelationshipXmlBytes,
            *relationship_metadata_bytes,
            opc_limits.max_total_relationship_xml_bytes(),
        ));
    }
    Ok(())
}

fn metadata_cost(lengths: impl IntoIterator<Item = usize>) -> Result<usize> {
    lengths
        .into_iter()
        .try_fold(0usize, |total, length| total.checked_add(length))
        // XML escaping can expand a byte several times; the fixed allowance
        // covers element/attribute names and delimiters without allocation.
        .and_then(|length| length.checked_mul(6))
        .and_then(|length| length.checked_add(256))
        .ok_or(Error::CapacityOverflow {
            resource: "Calculation Chain metadata",
        })
}

fn checked_metadata_add(current: usize, addition: usize, resource: &'static str) -> Result<usize> {
    current
        .checked_add(addition)
        .ok_or(Error::CapacityOverflow { resource })
}

fn preflight_metadata_text(
    resource: litchi_opc::ReadResource,
    actual: usize,
    maximum: usize,
) -> Result<()> {
    if actual > maximum {
        return Err(opc_limit(resource, actual, maximum));
    }
    Ok(())
}

fn preflight_xml_attribute(key: &str, value: &str, maximum: usize) -> Result<()> {
    let encoded = value.chars().try_fold(0usize, |length, character| {
        let bytes = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length.checked_add(bytes).ok_or(Error::CapacityOverflow {
            resource: "Calculation Chain XML attribute metadata",
        })
    })?;
    let actual = key
        .len()
        .checked_add(encoded)
        .ok_or(Error::CapacityOverflow {
            resource: "Calculation Chain XML attribute metadata",
        })?;
    preflight_metadata_text(litchi_opc::ReadResource::XmlAttributeBytes, actual, maximum)
}

fn opc_limit(resource: litchi_opc::ReadResource, actual: usize, maximum: usize) -> Error {
    Error::Opc(litchi_opc::OpcError::ReadLimit {
        resource,
        actual: actual as u64,
        maximum: maximum as u64,
    })
}

/// Scan every relationship source once for feature ownership and inbound
/// topology. `target` is optional for an absent-owner read; a probe target is
/// supplied by inverse publication to catch a foreign edge before re-adding a
/// removed physical part.
fn scan_relationship_graph(
    package: &OpcPackage,
    workbook_name: &PackURI,
    target: Option<&PackURI>,
    expected_source: Option<&PackURI>,
    expected_relationship_id: Option<&str>,
    limits: ReadLimits,
) -> Result<()> {
    let mut relationship_count = 0usize;
    let mut part_count = 0usize;
    for relationship in package.rels().iter() {
        relationship_count = checked_relationship_count(relationship_count, limits)?;
        if relationship.reltype() == RELATIONSHIP_TYPE {
            return Err(invalid(
                "Calculation Chain relationship cannot originate at the package root",
            ));
        }
        if let Some(target) = target {
            if relationship.target_mode() == TargetMode::Internal
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                return Err(invalid(
                    "Calculation Chain part has an unexpected package-level relationship",
                ));
            }
        }
    }
    for source in package.iter_parts() {
        part_count = checked_graph_part_count(part_count, limits)?;
        if source.content_type() == CONTENT_TYPE
            && target.is_none_or(|target| !source.partname().is_equivalent_to(target))
        {
            return Err(invalid(
                "package contains an orphan or additional Calculation Chain part",
            ));
        }
        for relationship in source.rels().iter() {
            relationship_count = checked_relationship_count(relationship_count, limits)?;
            if relationship.reltype() == RELATIONSHIP_TYPE
                && !source.partname().is_equivalent_to(workbook_name)
            {
                return Err(invalid(format!(
                    "Calculation Chain relationship on unexpected source part '{}'",
                    source.partname().as_str()
                )));
            }
            if let Some(target) = target {
                if relationship.target_mode() == TargetMode::Internal
                    && relationship.target_partname()?.is_equivalent_to(target)
                    && (expected_source
                        .is_none_or(|expected| !source.partname().is_equivalent_to(expected))
                        || expected_relationship_id
                            .is_none_or(|expected| relationship.r_id() != expected))
                {
                    return Err(invalid(
                        "Calculation Chain part has an unexpected inbound relationship",
                    ));
                }
            }
        }
    }
    Ok(())
}

fn checked_relationship_count(current: usize, limits: ReadLimits) -> Result<usize> {
    let actual = current.checked_add(1).ok_or(Error::CapacityOverflow {
        resource: "Calculation Chain relationships",
    })?;
    if actual > limits.max_relationships {
        return Err(Error::LimitExceeded {
            resource: "Calculation Chain relationships",
            actual,
            maximum: limits.max_relationships,
        });
    }
    Ok(actual)
}

fn checked_graph_part_count(current: usize, limits: ReadLimits) -> Result<usize> {
    let actual = current.checked_add(1).ok_or(Error::CapacityOverflow {
        resource: "Calculation Chain graph parts",
    })?;
    if actual > limits.max_graph_parts {
        return Err(Error::LimitExceeded {
            resource: "Calculation Chain graph parts",
            actual,
            maximum: limits.max_graph_parts,
        });
    }
    Ok(actual)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
