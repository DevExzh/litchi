use super::{Dialect, STRICT_CORE_PROPERTIES_RELATIONSHIP};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};

pub(super) struct Graph {
    pub(super) part: Option<PackURI>,
    pub(super) relationship_id: Option<String>,
    pub(super) dialect: Option<Dialect>,
}

pub(super) fn inspect(package: &OpcPackage) -> Result<Graph> {
    let mut core_part = None;
    for part in package.iter_parts() {
        if part.content_type() == ct::OPC_CORE_PROPERTIES {
            if core_part.is_some() {
                return Err(Error::Invalid(
                    "package contains multiple core-properties parts".to_owned(),
                ));
            }
            core_part = Some(part.partname().clone());
        }
        if part
            .rels()
            .iter()
            .any(|relationship| is_core_relationship(relationship.reltype()))
        {
            return Err(Error::Relationship(format!(
                "core-properties relationship must be package-level, not owned by '{}'",
                part.partname().as_str()
            )));
        }
    }

    let mut core_relationship = None;
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| is_core_relationship(relationship.reltype()))
    {
        if core_relationship.is_some() {
            return Err(Error::Relationship(
                "OPC M4.1 permits at most one core-properties relationship".to_owned(),
            ));
        }
        core_relationship = Some(relationship);
    }

    let Some(relationship) = core_relationship else {
        if let Some(part) = core_part {
            return Err(Error::Relationship(format!(
                "core-properties part '{}' is orphaned",
                part.as_str()
            )));
        }
        return Ok(Graph {
            part: None,
            relationship_id: None,
            dialect: None,
        });
    };

    if relationship.is_external() {
        return Err(Error::Relationship(
            "core-properties relationship has an external target".to_owned(),
        ));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(Error::Relationship(
            "core-properties relationship target cannot contain a query or fragment".to_owned(),
        ));
    }
    let requested = relationship.target_partname().map_err(|error| {
        Error::Relationship(format!(
            "invalid core-properties relationship target: {error}"
        ))
    })?;
    let part = package
        .iter_parts()
        .find(|part| same_name(part.partname(), &requested))
        .ok_or_else(|| {
            Error::Missing(format!(
                "core-properties relationship target '{}' does not exist",
                requested.as_str()
            ))
        })?;
    if part.content_type() != ct::OPC_CORE_PROPERTIES {
        return Err(Error::ContentType {
            expected: ct::OPC_CORE_PROPERTIES.to_owned(),
            actual: part.content_type().to_owned(),
        });
    }
    let actual = part.partname().clone();
    if core_part
        .as_ref()
        .is_some_and(|candidate| !same_name(candidate, &actual))
    {
        return Err(Error::Relationship(
            "core-properties relationship does not target the unique core-properties part"
                .to_owned(),
        ));
    }
    Ok(Graph {
        part: Some(actual),
        relationship_id: Some(relationship.r_id().to_owned()),
        dialect: Some(
            if relationship.reltype() == STRICT_CORE_PROPERTIES_RELATIONSHIP {
                Dialect::Strict
            } else {
                Dialect::Transitional
            },
        ),
    })
}

/// Ensures deleting the core-properties part cannot invalidate another owner.
///
/// Read and update intentionally allow differently typed inbound relationships:
/// updating a shared part leaves every edge valid. Deletion is stricter because
/// removing the part would leave those relationships dangling.
pub(super) fn ensure_clear_safe(
    package: &OpcPackage,
    target: &PackURI,
    core_id: &str,
) -> Result<()> {
    reject_other_inbound(package, target, core_id)
}

pub(super) fn is_core_relationship(relationship_type: &str) -> bool {
    matches!(
        relationship_type,
        rt::CORE_PROPERTIES | STRICT_CORE_PROPERTIES_RELATIONSHIP
    )
}

fn reject_other_inbound(package: &OpcPackage, target: &PackURI, core_id: &str) -> Result<()> {
    for relationship in package.rels().iter() {
        if relationship.r_id() == core_id || relationship.is_external() {
            continue;
        }
        let related = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid package relationship target: {error}"))
        })?;
        if same_name(&related, target) {
            return Err(Error::Relationship(format!(
                "core-properties part '{}' has another package-level inbound relationship",
                target.as_str()
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let related = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid relationship target from '{}': {error}",
                    source.partname().as_str()
                ))
            })?;
            if same_name(&related, target) {
                return Err(Error::Relationship(format!(
                    "core-properties part '{}' is shared by '{}'",
                    target.as_str(),
                    source.partname().as_str()
                )));
            }
        }
    }
    Ok(())
}

fn same_name(left: &PackURI, right: &PackURI) -> bool {
    left.as_str().eq_ignore_ascii_case(right.as_str())
}
