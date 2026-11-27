//! Package graph discovery for extended document properties.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Relationships, SourceBackedPackage};

use super::{Props, STRICT_RELATIONSHIP, codec};
use crate::properties::package_graph::{GraphPart, same_part_name};
use crate::{Error, Result};

/// The validated extended-properties graph of a package.
#[derive(Debug, Clone)]
pub(crate) struct Graph {
    /// The relationship-selected part, when present.
    pub(crate) part: Option<PackURI>,
    /// The package-level relationship selecting it, when present.
    pub(crate) relationship_id: Option<String>,
}

pub(crate) fn is_relationship(relationship_type: &str) -> bool {
    matches!(
        relationship_type,
        rt::EXTENDED_PROPERTIES | STRICT_RELATIONSHIP
    )
}

/// Inspect the extended-properties graph of an owning package; every
/// anomaly is an error.
///
/// # Errors
///
/// Returns [`Error::Relationship`] for more than one package-level
/// relationship, a part-owned, external or query-bearing relationship, or an
/// orphaned part; [`Error::Missing`] for a dangling relationship; and
/// [`Error::ContentType`] for a target of another content type.
pub(crate) fn inspect(package: &OpcPackage) -> Result<Graph> {
    inspect_parts(|| package.iter_parts(), package.rels())
}

/// [`inspect`] over the metadata-only catalog of a source-backed package.
pub(crate) fn inspect_source(package: &SourceBackedPackage) -> Result<Graph> {
    package.source_version()?;
    let graph = inspect_parts(|| package.iter_parts(), package.rels())?;
    package.source_version()?;
    Ok(graph)
}

fn inspect_parts<F, I, P>(parts: F, package_relationships: &Relationships) -> Result<Graph>
where
    F: Fn() -> I,
    I: IntoIterator<Item = P>,
    P: GraphPart,
{
    let mut typed_parts = Vec::new();
    for part in parts() {
        if part.part_content_type() == ct::OFC_EXTENDED_PROPERTIES {
            typed_parts.push(part.part_name().clone());
            if typed_parts.len() > 1 {
                return Err(Error::Invalid(
                    "package contains multiple extended-properties parts".to_owned(),
                ));
            }
        }
        if part
            .part_relationships()
            .iter()
            .any(|relationship| is_relationship(relationship.reltype()))
        {
            return Err(Error::Relationship(format!(
                "extended-properties relationship must be package-level, not owned by '{}'",
                part.part_name().as_str()
            )));
        }
    }
    let mut relationships = package_relationships
        .iter()
        .filter(|relationship| is_relationship(relationship.reltype()));
    let relationship = relationships.next();
    if relationships.next().is_some() {
        return Err(Error::Relationship(
            "package contains multiple extended-properties relationships".to_owned(),
        ));
    }
    let Some(relationship) = relationship else {
        if let Some(part) = typed_parts.first() {
            return Err(Error::Relationship(format!(
                "extended-properties part '{}' is orphaned",
                part.as_str()
            )));
        }
        return Ok(Graph {
            part: None,
            relationship_id: None,
        });
    };
    if relationship.is_external() {
        return Err(Error::Relationship(
            "extended-properties relationship cannot be external".to_owned(),
        ));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(Error::Relationship(
            "extended-properties relationship target cannot contain a query or fragment".to_owned(),
        ));
    }
    let target = relationship.target_partname().map_err(|error| {
        Error::Relationship(format!(
            "invalid extended-properties relationship target: {error}"
        ))
    })?;
    let part = parts()
        .into_iter()
        .find(|part| same_part_name(part.part_name(), &target))
        .ok_or_else(|| Error::Missing(target.to_string()))?;
    if part.part_content_type() != ct::OFC_EXTENDED_PROPERTIES {
        return Err(Error::ContentType {
            expected: ct::OFC_EXTENDED_PROPERTIES.to_owned(),
            actual: part.part_content_type().to_owned(),
        });
    }
    let part = part.part_name().clone();
    if typed_parts
        .first()
        .is_some_and(|candidate| !same_part_name(candidate, &part))
    {
        return Err(Error::Relationship(
            "extended-properties relationship does not target the unique extended-properties part"
                .to_owned(),
        ));
    }
    Ok(Graph {
        part: Some(part),
        relationship_id: Some(relationship.r_id().to_owned()),
    })
}

/// Read the relationship-selected extended properties of a package.
///
/// Absence is `Ok(None)`, distinct from a present but empty part.
///
/// # Errors
///
/// Returns the graph errors of the package-level relationship (more than
/// one, part-owned, external, dangling, or a target of another content
/// type) and the strict decoding errors of the part (see the
/// [module](super) documentation).
pub fn read(package: &OpcPackage) -> Result<Option<Props>> {
    let graph = inspect(package)?;
    let Some(part) = graph.part else {
        return Ok(None);
    };
    codec::decode(package.get_part(&part)?.blob()).map(Some)
}

/// [`read`] over a source-backed package: only the relationship-selected
/// part is read, with the package's source-freshness and execution checks
/// around it.
///
/// # Errors
///
/// The errors of [`read`], and a source-change or cancellation error.
pub fn read_source_backed(package: &SourceBackedPackage) -> Result<Option<Props>> {
    package.check_execution()?;
    let graph = inspect_source(package)?;
    let Some(part) = graph.part else {
        return Ok(None);
    };
    let data = package.part(&part)?.data()?;
    package.check_execution()?;
    let props = codec::decode(data.as_bytes())?;
    package.source_version()?;
    Ok(Some(props))
}
