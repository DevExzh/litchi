use super::{Dialect, STRICT_CORE_PROPERTIES_RELATIONSHIP};
use crate::leniency::{CorePropertiesDefect, ToleranceReport};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PACKAGE_URI;
use litchi_opc::{
    OpcPackage, PackURI, Part, PartMetadata, PartView, Relationships, SourceBackedPackage,
};

pub(super) struct Graph {
    pub(super) part: Option<PackURI>,
    pub(super) relationship_id: Option<String>,
    pub(super) dialect: Option<Dialect>,
}

/// Inspect the graph strictly, as the writer needs it: every anomaly is an
/// error, so a write never acts on an ambiguous graph.
pub(super) fn inspect(package: &OpcPackage) -> Result<Graph> {
    inspect_parts(|| package.iter_parts(), package.rels(), || Ok(()), None)
}

/// Inspect the graph for a read, repairing and recording the anomalies of
/// LENIENCY_AUDIT 8.2 into `report`.
pub(super) fn inspect_tolerant(
    package: &OpcPackage,
    report: &mut ToleranceReport,
) -> Result<Graph> {
    inspect_parts(
        || package.iter_parts(),
        package.rels(),
        || Ok(()),
        Some(report),
    )
}

/// Inspect the same core-properties graph over the metadata-only
/// source-backed catalog. No ordinary part payload is read here. The
/// metadata-only part projection is deliberately fed through the same graph
/// implementation as the eager package path so strict OPC edge rules cannot
/// drift between ingress modes.
pub(super) fn inspect_source(
    package: &SourceBackedPackage,
    report: &mut ToleranceReport,
) -> Result<Graph> {
    package.source_version()?;
    let result = inspect_parts(
        || package.iter_parts(),
        package.rels(),
        || Ok(package.check_execution()?),
        Some(report),
    );
    package.source_version()?;
    result
}

trait GraphPart {
    fn graph_partname(&self) -> &PackURI;
    fn graph_content_type(&self) -> &str;
    fn graph_rels(&self) -> &Relationships;
}

impl GraphPart for &dyn Part {
    fn graph_partname(&self) -> &PackURI {
        self.partname()
    }

    fn graph_content_type(&self) -> &str {
        self.content_type()
    }

    fn graph_rels(&self) -> &Relationships {
        self.rels()
    }
}

impl GraphPart for PartMetadata<'_> {
    fn graph_partname(&self) -> &PackURI {
        self.partname()
    }

    fn graph_content_type(&self) -> &str {
        self.content_type()
    }

    fn graph_rels(&self) -> &Relationships {
        self.rels()
    }
}

impl GraphPart for PartView<'_> {
    fn graph_partname(&self) -> &PackURI {
        self.partname()
    }

    fn graph_content_type(&self) -> &str {
        self.content_type()
    }

    fn graph_rels(&self) -> &Relationships {
        self.rels()
    }
}

/// Locate the relationship-selected core-properties part.
///
/// With a `report`, the graph anomalies Word, LibreOffice and POI read past
/// (LENIENCY_AUDIT 8.2) are repaired and recorded instead of failing: a
/// core-properties relationship owned by a part is ignored, a second
/// core-typed part is dead weight beside the relationship-selected one, an
/// orphaned part means no core properties, and a dangling relationship means
/// no core properties. Without one (the writer) they stay errors. An
/// external, query/fragment or wrongly typed target, and more than one
/// package-level relationship (OPC M4.1), are errors either way.
fn inspect_parts<F, I, P, C>(
    parts: F,
    package_relationships: &Relationships,
    mut check: C,
    mut report: Option<&mut ToleranceReport>,
) -> Result<Graph>
where
    F: Fn() -> I,
    I: IntoIterator<Item = P>,
    P: GraphPart,
    C: FnMut() -> Result<()>,
{
    let mut core_part = None;
    // Further core-typed parts, tolerated only on the read path.
    let mut extra_core_parts = Vec::new();
    for part in parts() {
        check()?;
        if part.graph_content_type() == ct::OPC_CORE_PROPERTIES {
            if core_part.is_some() {
                if report.is_none() {
                    return Err(Error::Invalid(
                        "package contains multiple core-properties parts".to_owned(),
                    ));
                }
                extra_core_parts.push(part.graph_partname().clone());
            } else {
                core_part = Some(part.graph_partname().clone());
            }
        }
        for relationship in part.graph_rels().iter() {
            check()?;
            if is_core_relationship(relationship.reltype()) {
                let Some(report) = report.as_deref_mut() else {
                    return Err(Error::Relationship(format!(
                        "core-properties relationship must be package-level, not owned by '{}'",
                        part.graph_partname().as_str()
                    )));
                };
                report.record(
                    CorePropertiesDefect::MisplacedRelationship,
                    part.graph_partname().as_str(),
                    Some(relationship.r_id()),
                    Some(relationship.target_ref()),
                );
            }
        }
    }

    let mut core_relationship = None;
    for relationship in package_relationships.iter() {
        check()?;
        if !is_core_relationship(relationship.reltype()) {
            continue;
        }
        if core_relationship.is_some() {
            return Err(Error::Relationship(
                "OPC M4.1 permits at most one core-properties relationship".to_owned(),
            ));
        }
        core_relationship = Some(relationship);
    }

    let Some(relationship) = core_relationship else {
        if let Some(part) = core_part {
            let Some(report) = report else {
                return Err(Error::Relationship(format!(
                    "core-properties part '{}' is orphaned",
                    part.as_str()
                )));
            };
            for orphan in std::iter::once(&part).chain(&extra_core_parts) {
                report.record(
                    CorePropertiesDefect::OrphanedPart,
                    orphan.as_str(),
                    None,
                    None,
                );
            }
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
    let mut actual = None;
    for part in parts() {
        check()?;
        if !same_name(part.graph_partname(), &requested) {
            continue;
        }
        if part.graph_content_type() != ct::OPC_CORE_PROPERTIES {
            return Err(Error::ContentType {
                expected: ct::OPC_CORE_PROPERTIES.to_owned(),
                actual: part.graph_content_type().to_owned(),
            });
        }
        actual = Some(part.graph_partname().clone());
        break;
    }
    let Some(actual) = actual else {
        let Some(report) = report else {
            return Err(Error::Missing(format!(
                "core-properties relationship target '{}' does not exist",
                requested.as_str()
            )));
        };
        report.record(
            CorePropertiesDefect::DanglingRelationship,
            PACKAGE_URI,
            Some(relationship.r_id()),
            Some(relationship.target_ref()),
        );
        for orphan in core_part.iter().chain(&extra_core_parts) {
            report.record(
                CorePropertiesDefect::OrphanedPart,
                orphan.as_str(),
                None,
                None,
            );
        }
        return Ok(Graph {
            part: None,
            relationship_id: None,
            dialect: None,
        });
    };
    if let Some(report) = report {
        for extra in core_part.iter().chain(&extra_core_parts) {
            if !same_name(extra, &actual) {
                report.record(
                    CorePropertiesDefect::DuplicatePart,
                    extra.as_str(),
                    None,
                    None,
                );
            }
        }
    } else if core_part
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
