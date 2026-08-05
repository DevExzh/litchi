//! Package graph discovery and mutation for custom document properties.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

use super::codec;
use super::model::Props;
use super::schema::{PART_NAME, PART_TARGET, invalid};
use crate::{Error, Result};

impl Props {
    /// Reads the custom-properties relationship and target part from a package.
    ///
    /// A genuinely absent relationship and part produce empty properties.
    /// Orphans, duplicate relationships, external targets, missing targets,
    /// wrong content types, and malformed XML are returned as typed errors.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        let graph = inspect_graph(package)?;
        let Some(part_name) = graph.part_name else {
            return Ok(Self::new());
        };
        let part = package.get_part(&part_name)?;
        codec::decode(part.blob())
    }

    /// Writes this collection to a package.
    ///
    /// Empty properties remove both the target part and its package-level
    /// relationship. Non-empty properties update the existing validated target
    /// or create the canonical `/docProps/custom.xml` part and relationship.
    pub fn write(&self, package: &mut OpcPackage) -> Result<()> {
        let graph = inspect_graph(package)?;
        if self.is_empty() {
            if graph.part_name.is_none() && graph.relationship_id.is_none() {
                return Ok(());
            }
            package.unsign();
            if let Some(part_name) = graph.part_name {
                let removed = package.remove_part(&part_name);
                if !removed {
                    return Err(Error::Missing(part_name.to_string()));
                }
            }
            if let Some(relationship_id) = graph.relationship_id {
                let removed = package.rels_mut().remove(&relationship_id);
                if removed.is_none() {
                    return Err(Error::Relationship(format!(
                        "custom-properties relationship '{relationship_id}' disappeared during removal"
                    )));
                }
            }
            return Ok(());
        }

        let xml = codec::encode(self)?;
        match graph.part_name {
            Some(part_name) => {
                if package.get_part(&part_name)?.blob() == xml.as_slice() {
                    return Ok(());
                }
                package.get_part_mut(&part_name)?.set_blob(xml);
                package.unsign();
            },
            None => {
                let part_name = custom_part_name()?;
                package.validate_new_part_name(&part_name)?;
                let part = BlobPart::new(part_name, ct::OFC_CUSTOM_PROPERTIES.to_owned(), xml);
                package.unsign();
                package.add_part(Box::new(part));
                package.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
            },
        }
        Ok(())
    }
}

struct PackageGraph {
    part_name: Option<PackURI>,
    relationship_id: Option<String>,
}

fn inspect_graph(package: &OpcPackage) -> Result<PackageGraph> {
    let canonical = custom_part_name()?;
    let mut custom_parts = Vec::new();
    for part in package.iter_parts() {
        if part
            .partname()
            .as_str()
            .eq_ignore_ascii_case(canonical.as_str())
            && part.content_type() != ct::OFC_CUSTOM_PROPERTIES
        {
            return Err(Error::ContentType {
                expected: ct::OFC_CUSTOM_PROPERTIES.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        if part.content_type() == ct::OFC_CUSTOM_PROPERTIES {
            custom_parts.push(part.partname().clone());
            if custom_parts.len() > 1 {
                return Err(invalid("package contains multiple custom-properties parts"));
            }
        }
        if part
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::CUSTOM_PROPERTIES)
        {
            return Err(Error::Relationship(format!(
                "custom-properties relationship must be package-level, not owned by '{}'",
                part.partname().as_str()
            )));
        }
    }

    let relationships: Vec<_> = package
        .rels()
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
    let target_part = package
        .iter_parts()
        .find(|part| {
            part.partname()
                .as_str()
                .eq_ignore_ascii_case(target.as_str())
        })
        .ok_or_else(|| Error::Missing(target.to_string()))?;
    if target_part.content_type() != ct::OFC_CUSTOM_PROPERTIES {
        return Err(Error::ContentType {
            expected: ct::OFC_CUSTOM_PROPERTIES.to_owned(),
            actual: target_part.content_type().to_owned(),
        });
    }
    let part_name = target_part.partname().clone();
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

pub(super) fn custom_part_name() -> Result<PackURI> {
    PackURI::new(PART_NAME).map_err(|error| Error::Uri(error.to_string()))
}
