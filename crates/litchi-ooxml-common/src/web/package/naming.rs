use super::super::codec::*;
use super::super::model::*;
use super::super::*;
use super::*;
pub(in crate::web) fn fold_part_name(name: &PackURI) -> String {
    name.as_str().to_ascii_lowercase()
}

pub(in crate::web) fn existing_web_extension_graph(
    package: &OpcPackage,
    limits: &Limits,
    index: &PackageGraphIndex,
    budget: &mut OperationBudget,
) -> Result<Option<ExistingAddInGraph>> {
    let Some(loaded) = load_with_index_budget(package, limits, index, budget)? else {
        return Ok(None);
    };
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
        .ok_or_else(|| {
            Error::Relationship("loaded task panes have no package relationship".into())
        })?;
    let task_panes_target = checked_internal_target(relationship, "task-pane")?;
    let task_panes_name = index
        .canonical(&task_panes_target)
        .ok_or_else(|| Error::Missing(task_panes_target.to_string()))?
        .clone();
    let task_panes_part = package.get_part(&task_panes_name)?;
    let mut extensions_by_relationship = HashMap::with_capacity(loaded.panes.len());
    let mut owned = HashSet::new();
    owned.insert(task_panes_name.clone());
    for pane in loaded.panes {
        let child_relationship = task_panes_part
            .rels()
            .get(&pane.relationship_id)
            .ok_or_else(|| {
                Error::Relationship(format!(
                    "task pane references missing relationship '{}'",
                    pane.relationship_id
                ))
            })?;
        let extension_target = checked_internal_target(child_relationship, "add-in")?;
        let extension_name = index
            .canonical(&extension_target)
            .ok_or_else(|| Error::Missing(extension_target.to_string()))?
            .clone();
        if extensions_by_relationship
            .insert(pane.relationship_id.clone(), extension_name.clone())
            .is_some()
        {
            return invalid(format!(
                "duplicate task-pane relationship ID '{}'",
                pane.relationship_id
            ));
        }
        owned.insert(extension_name);
        for resource in pane.snapshot_resources {
            if let SnapshotTarget::Internal { part_name, .. } = resource.target {
                owned.insert(part_name);
            }
        }
    }
    let mut owned_parts: Vec<_> = owned.into_iter().collect();
    owned_parts.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    Ok(Some(ExistingAddInGraph {
        root_relationship_id: relationship.r_id().to_owned(),
        task_panes_name,
        extensions_by_relationship,
        owned_parts,
    }))
}

pub(in crate::web) fn next_web_extension_part_name(
    index: &PackageGraphIndex,
    reserved: &BTreeSet<String>,
    preferred_index: usize,
    limits: &Limits,
    allocation_probes: &mut usize,
) -> Result<PackURI> {
    let mut offset = 0usize;
    loop {
        charge_allocation_probe(allocation_probes, limits)?;
        let part_number = preferred_index
            .checked_add(offset)
            .ok_or_else(|| Error::Invalid("web extension part index overflow".into()))?;
        let candidate = PackURI::new(format!("/webextensions/webextension{part_number}.xml"))
            .map_err(Error::Uri)?;
        if !folded_name_conflicts(reserved, &fold_part_name(&candidate))
            && !index.conflicts(&candidate)
        {
            return Ok(candidate);
        }
        offset = offset
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("web extension part index overflow".into()))?;
    }
}

pub(in crate::web) fn next_task_panes_part_name(
    index: &PackageGraphIndex,
    limits: &Limits,
    allocation_probes: &mut usize,
) -> Result<PackURI> {
    let mut attempt = 1usize;
    loop {
        charge_allocation_probe(allocation_probes, limits)?;
        let suffix = if attempt == 1 {
            String::new()
        } else {
            attempt.to_string()
        };
        let candidate =
            PackURI::new(format!("/webextensions/taskpanes{suffix}.xml")).map_err(Error::Uri)?;
        if !index.conflicts(&candidate) {
            return Ok(candidate);
        }
        attempt = attempt
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("task-pane part index overflow".into()))?;
    }
}

pub(in crate::web) fn charge_allocation_probe(probes: &mut usize, limits: &Limits) -> Result<()> {
    let actual = probes.saturating_add(1);
    if *probes >= limits.part_allocations {
        return limit(
            "web extension part-name allocation probes",
            limits.part_allocations,
            actual,
        );
    }
    *probes = actual;
    Ok(())
}

pub(in crate::web) fn next_package_relationship_id(
    package: &OpcPackage,
    limits: &Limits,
) -> Result<String> {
    let attempts = package.rels().len().checked_add(1).ok_or(Error::Limit {
        resource: "web extension relationship IDs",
        max: limits.part_allocations,
        actual: usize::MAX,
    })?;
    let attempts = attempts.min(limits.part_allocations);
    for index in 1..=attempts {
        let candidate = format!("rIdPanes{index}");
        if package.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(Error::Relationship(
        "unable to allocate a task-pane relationship ID".into(),
    ))
}

pub(in crate::web) fn add_or_match_planned_part(
    planned: &mut Vec<PlannedPart>,
    by_name: &mut HashMap<String, usize>,
    part: PlannedPart,
) -> Result<()> {
    let key = fold_part_name(&part.name);
    if let Some(existing) = by_name.get(&key).and_then(|index| planned.get(*index)) {
        if existing == &part {
            return Ok(());
        }
        return invalid(format!(
            "conflicting authored resources target '{}'",
            part.name.as_str()
        ));
    }
    by_name.insert(key, planned.len());
    planned.push(part);
    Ok(())
}
