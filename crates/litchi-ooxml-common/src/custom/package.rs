//! Package graph discovery for custom document properties.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Relationships, SourceBackedPackage};

use super::codec;
use super::model::Props;
use super::reserved::{self, Host};
use super::schema::{PART_NAME, invalid};
use crate::leniency::CustomPropertiesDefect;
use crate::properties::package_graph::GraphPart;
use crate::{Error, Result};

impl Props {
    /// Reads the custom-properties relationship and target part from a package.
    ///
    /// A genuinely absent relationship and part produce empty properties.
    /// Orphans, duplicate relationships, external targets, missing targets,
    /// wrong content types, and malformed XML are returned as typed errors.
    /// A malformed or unsupported individual property is skipped and recorded
    /// in [`Props::tolerance_report`] instead (see [`crate::leniency`]).
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Ok(Self::read_selected(package)?.0)
    }

    /// Read the relationship-selected part, returning its name with the
    /// collection.
    fn read_selected(package: &OpcPackage) -> Result<(Self, Option<PackURI>)> {
        let graph = inspect_graph(package)?;
        let Some(part_name) = graph.part_name else {
            return Ok((Self::new(), None));
        };
        let part = package.get_part(&part_name)?;
        let properties = codec::decode_tolerant(part.blob(), part_name.as_str())?;
        Ok((properties, Some(part_name)))
    }

    /// Reads custom properties and validates reserved names for `host`.
    ///
    /// The host-independent parser preserves opaque Microsoft Information
    /// Protection SDK metadata. This method additionally rejects reserved
    /// sensitivity-marking properties that belong to a different host.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    ///
    /// On read the host check is advisory (LENIENCY_AUDIT 8.4): a reserved
    /// property scoped to another host, or one whose value lacks the inferred
    /// format, is skipped and recorded as
    /// [`ReservedProperty`](crate::leniency::CustomPropertiesDefect::ReservedProperty)
    /// rather than failing the document. The semantic metadata transaction
    /// validates the host scope again before publication.
    pub fn read_for(package: &OpcPackage, host: Host) -> Result<Self> {
        let (properties, part_name) = Self::read_selected(package)?;
        properties.admit_for(part_name.as_ref(), host)
    }

    /// Decode the bytes of the relationship-selected part `part_name` as
    /// [`Self::read_for`] decodes them.
    pub(crate) fn decode_for(bytes: &[u8], part_name: &PackURI, host: Host) -> Result<Self> {
        codec::decode_tolerant(bytes, part_name.as_str())?.admit_for(Some(part_name), host)
    }

    /// Skip and record the reserved properties of another host, then
    /// validate the host's own reserved properties ([`Self::read_for`]).
    fn admit_for(self, part_name: Option<&PackURI>, host: Host) -> Result<Self> {
        let mut properties = self;
        let rejected: Vec<String> = properties
            .properties
            .iter()
            .filter(|(name, property)| {
                reserved::validate_property(name, &property.value, Some(host)).is_err()
            })
            .map(|(name, _)| name.clone())
            .collect();
        let part = part_name.map_or(PART_NAME, PackURI::as_str);
        for name in rejected {
            properties.remove(&name);
            properties.tolerance_report.record(
                CustomPropertiesDefect::ReservedProperty,
                part,
                Some(&name),
                Some(host_name(host)),
            );
        }
        properties.validate_for(host)?;
        Ok(properties)
    }
}

/// The validated custom-properties graph of a package.
pub(crate) struct PackageGraph {
    /// The relationship-selected part, when present.
    pub(crate) part_name: Option<PackURI>,
    /// The package-level relationship selecting it, when present.
    pub(crate) relationship_id: Option<String>,
}

/// Inspect the custom-properties graph of an owning package; every anomaly
/// is an error.
pub(crate) fn inspect_graph(package: &OpcPackage) -> Result<PackageGraph> {
    inspect_graph_parts(|| package.iter_parts(), package.rels())
}

/// [`inspect_graph`] over the metadata-only catalog of a source-backed
/// package. No part payload is read.
pub(crate) fn inspect_source_graph(package: &SourceBackedPackage) -> Result<PackageGraph> {
    package.source_version()?;
    let graph = inspect_graph_parts(|| package.iter_parts(), package.rels())?;
    package.source_version()?;
    Ok(graph)
}

fn inspect_graph_parts<F, I, P>(
    parts: F,
    package_relationships: &Relationships,
) -> Result<PackageGraph>
where
    F: Fn() -> I,
    I: IntoIterator<Item = P>,
    P: GraphPart,
{
    let canonical = custom_part_name()?;
    let mut custom_parts = Vec::new();
    for part in parts() {
        if part
            .part_name()
            .as_str()
            .eq_ignore_ascii_case(canonical.as_str())
            && part.part_content_type() != ct::OFC_CUSTOM_PROPERTIES
        {
            return Err(Error::ContentType {
                expected: ct::OFC_CUSTOM_PROPERTIES.to_owned(),
                actual: part.part_content_type().to_owned(),
            });
        }
        if part.part_content_type() == ct::OFC_CUSTOM_PROPERTIES {
            custom_parts.push(part.part_name().clone());
            if custom_parts.len() > 1 {
                return Err(invalid("package contains multiple custom-properties parts"));
            }
        }
        if part
            .part_relationships()
            .iter()
            .any(|relationship| relationship.reltype() == rt::CUSTOM_PROPERTIES)
        {
            return Err(Error::Relationship(format!(
                "custom-properties relationship must be package-level, not owned by '{}'",
                part.part_name().as_str()
            )));
        }
    }

    let relationships: Vec<_> = package_relationships
        .iter()
        .filter(|relationship| relationship.reltype() == rt::CUSTOM_PROPERTIES)
        .collect();
    if relationships.len() > 1 {
        return Err(Error::Relationship(
            "package contains multiple custom-properties relationships".to_owned(),
        ));
    }

    let Some(relationship) = relationships.first().copied() else {
        if let Some(part_name) = custom_parts.first() {
            return Err(Error::Relationship(format!(
                "custom-properties part '{}' is orphaned",
                part_name.as_str()
            )));
        }
        return Ok(PackageGraph {
            part_name: None,
            relationship_id: None,
        });
    };
    if relationship.is_external() {
        return Err(Error::Relationship(
            "custom-properties relationship cannot be external".to_owned(),
        ));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(Error::Relationship(
            "custom-properties relationship target cannot contain a query or fragment".to_owned(),
        ));
    }
    let target = relationship.target_partname().map_err(|error| {
        Error::Relationship(format!(
            "invalid custom-properties relationship target: {error}"
        ))
    })?;
    let target_part = parts()
        .into_iter()
        .find(|part| {
            part.part_name()
                .as_str()
                .eq_ignore_ascii_case(target.as_str())
        })
        .ok_or_else(|| Error::Missing(target.to_string()))?;
    if target_part.part_content_type() != ct::OFC_CUSTOM_PROPERTIES {
        return Err(Error::ContentType {
            expected: ct::OFC_CUSTOM_PROPERTIES.to_owned(),
            actual: target_part.part_content_type().to_owned(),
        });
    }
    let part_name = target_part.part_name().clone();
    if custom_parts
        .first()
        .is_some_and(|candidate| !candidate.as_str().eq_ignore_ascii_case(part_name.as_str()))
    {
        return Err(Error::Relationship(
            "custom-properties relationship does not target the unique custom-properties part"
                .to_owned(),
        ));
    }

    Ok(PackageGraph {
        part_name: Some(part_name),
        relationship_id: Some(relationship.r_id().to_owned()),
    })
}

fn host_name(host: Host) -> &'static str {
    match host {
        Host::Excel => "Excel",
        Host::Word => "Word",
        Host::PowerPoint => "PowerPoint",
    }
}

pub(super) fn custom_part_name() -> Result<PackURI> {
    PackURI::new(PART_NAME).map_err(|error| Error::Uri(error.clone()))
}
