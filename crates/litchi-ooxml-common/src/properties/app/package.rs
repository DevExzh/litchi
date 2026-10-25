//! Package graph discovery and mutation for extended document properties.

use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::rel::TargetMode;
use litchi_opc::{OpcPackage, PackURI, Relationships, SourceBackedPackage};

use super::{Props, STRICT_RELATIONSHIP, codec};
use crate::properties::package_graph::{
    GraphPart, ensure_only_inbound, free_part_name, same_part_name,
};
use crate::signature::{refuse_signed_change, settle_unsigned_change};
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

/// Write `props` to a package's extended-properties part.
///
/// An existing part is edited in place: only the scalar elements whose
/// value changes are rewritten, so the vector properties, extension content
/// and the bytes of every unchanged element survive. A missing part is
/// created as `/docProps/app.xml` (or the first free `/docProps/appN.xml`)
/// with a package-level relationship. Returns whether the package changed.
///
/// An equal value publishes nothing, so a signed package keeps its
/// signature; a change to a signed package is refused before anything is
/// mutated (ADR 0006): remove the signature explicitly first with
/// [`OpcPackage::unsign`] or the format facade's `unsign`.
///
/// # Errors
///
/// Returns [`Error::Opc`] with
/// [`SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// for a change to a signed package, the errors of [`read`], and an error
/// when a value cannot be written.
pub(crate) fn write(package: &mut OpcPackage, props: Props) -> Result<bool> {
    let graph = inspect(package)?;
    if let Some(part) = graph.part {
        let current = package.get_part(&part)?.blob();
        let Some(xml) = codec::splice(current, &props)? else {
            return Ok(false);
        };
        refuse_signed_change(package)?;
        package.get_part_mut(&part)?.set_blob_shared(Arc::new(xml));
        settle_unsigned_change(package)?;
        return Ok(true);
    }
    let part = free_part_name(
        &*package,
        "/docProps/app",
        crate::LimitKind::ExtendedPropertiesPartNames,
    )?;
    let xml = codec::encode_new(&props)?;
    let mut relationships = package.rels().clone();
    let id = next_relationship_id(&relationships)?;
    relationships.try_add_relationship(
        rt::EXTENDED_PROPERTIES.to_owned(),
        part.as_str().trim_start_matches('/').to_owned(),
        id,
        TargetMode::Internal,
    )?;
    package.validate_new_part_name(&part)?;
    refuse_signed_change(package)?;
    package.add_part(Box::new(BlobPart::new(
        part,
        ct::OFC_EXTENDED_PROPERTIES.to_owned(),
        xml,
    )));
    *package.rels_mut() = relationships;
    settle_unsigned_change(package)?;
    Ok(true)
}

/// Remove a package's extended-properties part and its relationship.
///
/// Absence is a successful no-op, which keeps a signed package's
/// signature. A part another relationship also targets is refused, as is
/// removing a present part from a signed package (ADR 0006). Returns whether
/// the package changed.
///
/// # Errors
///
/// Returns [`Error::Opc`] with
/// [`SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// for a signed package, [`Error::Relationship`] for a shared part, and the
/// graph errors of [`read`].
pub(crate) fn clear(package: &mut OpcPackage) -> Result<bool> {
    let graph = inspect(package)?;
    let (Some(part), Some(id)) = (graph.part, graph.relationship_id) else {
        return Ok(false);
    };
    ensure_only_inbound(&*package, &part, &id, "extended-properties")?;
    let mut relationships = package.rels().clone();
    relationships.remove(&id);
    refuse_signed_change(package)?;
    if !package.remove_part(&part) {
        return Err(Error::Missing(part.to_string()));
    }
    *package.rels_mut() = relationships;
    settle_unsigned_change(package)?;
    Ok(true)
}

/// The first `rIdN` the package-level relationships do not use.
pub(crate) fn next_relationship_id(relationships: &Relationships) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("rId{index}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        kind: crate::LimitKind::PackageRelationshipIdentifiers,
        maximum: u32::MAX as usize,
        observed: usize::MAX,
    })
}
