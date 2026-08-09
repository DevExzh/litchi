use super::super::codec::{invalid, limit};
use super::super::model::Limits;
use super::super::{BTreeSet, Error, HashSet, OpcPackage, PackURI, Part, Result};
use super::{ExistingAddInGraph, PackageGraphIndex, PlannedPart, fold_part_name};
pub(in crate::web) fn preflight_planned_parts(
    package: &OpcPackage,
    index: &PackageGraphIndex,
    planned: &[PlannedPart],
    old_parts: &[PackURI],
    protected: &HashSet<String>,
) -> Result<()> {
    let old_names: HashSet<_> = old_parts.iter().map(fold_part_name).collect();
    let mut planned_names = BTreeSet::new();
    for part in planned {
        let folded = fold_part_name(&part.name);
        if folded_name_conflicts(&planned_names, &folded) {
            return invalid(format!(
                "authored part '{}' conflicts with another planned part",
                part.name.as_str()
            ));
        }
        planned_names.insert(folded.clone());
        if let Some(canonical) = index.canonical(&part.name) {
            if canonical != &part.name {
                return invalid(format!(
                    "authored part name '{}' differs in case from canonical package part '{}'",
                    part.name.as_str(),
                    canonical.as_str()
                ));
            }
            if !old_names.contains(&folded) {
                return invalid(format!(
                    "authored web extension part '{}' already exists outside the replaced graph",
                    part.name.as_str()
                ));
            }
            let existing = package.get_part(canonical)?;
            if existing.content_type() != part.content_type {
                return invalid(format!(
                    "cannot change content type of existing part '{}'",
                    part.name.as_str()
                ));
            }
            if !planned_part_matches_existing(existing, part) && protected.contains(&folded) {
                return Err(Error::Relationship(format!(
                    "cannot replace protected shared web extension part '{}'",
                    part.name.as_str()
                )));
            }
        } else if index.conflicts(&part.name) {
            return invalid(format!(
                "authored part '{}' conflicts with an existing package part",
                part.name.as_str()
            ));
        }
    }
    Ok(())
}

pub(in crate::web) fn folded_name_conflicts(names: &BTreeSet<String>, candidate: &str) -> bool {
    if names.contains(candidate) {
        return true;
    }
    let mut ancestor = candidate;
    while let Some(index) = ancestor.rfind('/') {
        if index == 0 {
            break;
        }
        ancestor = &ancestor[..index];
        if names.contains(ancestor) {
            return true;
        }
    }
    let prefix = format!("{candidate}/");
    names
        .range(prefix.clone()..)
        .next()
        .is_some_and(|name| name.starts_with(&prefix))
}

pub(in crate::web) fn part_names_conflict(left: &PackURI, right: &PackURI) -> bool {
    let left = fold_part_name(left);
    let right = fold_part_name(right);
    left == right
        || left
            .strip_prefix(&right)
            .is_some_and(|suffix| suffix.starts_with('/'))
        || right
            .strip_prefix(&left)
            .is_some_and(|suffix| suffix.starts_with('/'))
}

pub(in crate::web) fn planned_deletions(
    old_parts: &[PackURI],
    retained: &BTreeSet<String>,
    protected: &HashSet<String>,
    limits: &Limits,
) -> Result<Vec<PackURI>> {
    let mut deletions = Vec::new();
    for name in old_parts {
        let folded = fold_part_name(name);
        if protected.contains(&folded) || retained.contains(&folded) {
            continue;
        }
        deletions.push(name.clone());
    }
    if deletions.len() > limits.part_deletions {
        return limit(
            "web extension part deletions",
            limits.part_deletions,
            deletions.len(),
        );
    }
    Ok(deletions)
}

pub(in crate::web) fn validate_plan_counts(
    package: &OpcPackage,
    index: &PackageGraphIndex,
    planned: &[PlannedPart],
    existing_graph: Option<&ExistingAddInGraph>,
    limits: &Limits,
) -> Result<()> {
    let new_parts = planned
        .iter()
        .filter(|part| !index.contains(&part.name))
        .count();
    if new_parts > limits.part_allocations {
        return limit(
            "web extension part allocations",
            limits.part_allocations,
            new_parts,
        );
    }
    let peak_parts = index
        .parts
        .len()
        .checked_add(new_parts)
        .ok_or(Error::Limit {
            resource: "package parts",
            max: limits.package_parts,
            actual: usize::MAX,
        })?;
    if peak_parts > limits.package_parts {
        return limit("package parts", limits.package_parts, peak_parts);
    }

    let mut relationships = index.relationships;
    for part in planned {
        if let Some(canonical) = index.canonical(&part.name) {
            relationships = relationships
                .checked_sub(package.get_part(canonical)?.rels().len())
                .ok_or_else(|| Error::Invalid("relationship count underflow".into()))?;
        }
        relationships =
            relationships
                .checked_add(part.relationships.len())
                .ok_or(Error::Limit {
                    resource: "package relationships",
                    max: limits.package_relationships,
                    actual: usize::MAX,
                })?;
    }
    if existing_graph.is_none() {
        relationships = relationships.checked_add(1).ok_or(Error::Limit {
            resource: "package relationships",
            max: limits.package_relationships,
            actual: usize::MAX,
        })?;
    }
    if relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            relationships,
        );
    }
    Ok(())
}

pub(in crate::web) fn graph_matches_plan(
    package: &OpcPackage,
    planned: &[PlannedPart],
    existing: &ExistingAddInGraph,
    retained: &BTreeSet<String>,
) -> bool {
    existing.owned_parts.len() == retained.len()
        && existing
            .owned_parts
            .iter()
            .all(|name| retained.contains(&fold_part_name(name)))
        && planned.iter().all(|part| {
            package
                .get_part(&part.name)
                .is_ok_and(|existing_part| planned_part_matches_existing(existing_part, part))
        })
}

pub(in crate::web) fn planned_part_matches_existing(
    existing: &dyn Part,
    planned: &PlannedPart,
) -> bool {
    existing.content_type() == planned.content_type
        && existing.blob() == planned.data.as_slice()
        && existing.rels().len() == planned.relationships.len()
        && planned.relationships.iter().all(|planned_relationship| {
            existing
                .rels()
                .get(&planned_relationship.id)
                .is_some_and(|relationship| {
                    relationship.reltype() == planned_relationship.relationship_type
                        && relationship.target_ref() == planned_relationship.target
                        && relationship.is_external() == planned_relationship.external
                })
        })
}
