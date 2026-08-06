use super::super::codec::*;
use super::super::model::*;
use super::super::*;
use super::*;
pub(in crate::web) fn has_task_panes_relationship(
    package: &OpcPackage,
    limits: &Limits,
) -> Result<bool> {
    let relationships = package.rels().len();
    if relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            relationships,
        );
    }
    Ok(package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP))
}

/// Resolve and validate the complete package graph with safe default limits.
pub fn load(package: &OpcPackage) -> Result<Option<Panes>> {
    load_with(package, &Limits::standard())
}

/// Resolve and validate the complete package graph with explicit limits.
pub fn load_with(package: &OpcPackage, limits: &Limits) -> Result<Option<Panes>> {
    if !has_task_panes_relationship(package, limits)? {
        return Ok(None);
    }
    let mut budget = OperationBudget::default();
    let index = PackageGraphIndex::build(package, limits, &mut budget)?;
    load_with_index_budget(package, limits, &index, &mut budget)
}

pub(in crate::web) fn load_with_index_budget(
    package: &OpcPackage,
    limits: &Limits,
    index: &PackageGraphIndex,
    budget: &mut OperationBudget,
) -> Result<Option<Panes>> {
    let relationships: Vec<_> = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
        .collect();
    if relationships.is_empty() {
        return Ok(None);
    }
    if relationships.len() != 1 {
        return Err(Error::Relationship(
            "package has multiple web extension task-pane relationships".into(),
        ));
    }
    let relationship = relationships[0];
    if relationship.is_external() {
        return Err(Error::Relationship(
            "task-pane relationship must be internal".into(),
        ));
    }
    let target = checked_internal_target(relationship, "task-pane")?;
    let part_name = index
        .canonical(&target)
        .ok_or_else(|| Error::Missing(format!("task-pane part '{}'", target.as_str())))?;
    let task_panes_part = package.get_part(part_name).map_err(|error| {
        Error::Missing(format!("task-pane part '{}': {error}", part_name.as_str()))
    })?;
    require_content_type(task_panes_part, TASK_PANES_CONTENT_TYPE)?;
    let parsed_panes = parse_panes_with_budget(task_panes_part.blob(), limits, budget)?;

    let referenced_ids: HashSet<&str> = parsed_panes
        .iter()
        .map(|pane| pane.relationship_id.as_str())
        .collect();
    if referenced_ids.len() != parsed_panes.len() {
        return invalid("task panes contain duplicate relationship IDs".into());
    }
    for child_relationship in task_panes_part.rels().iter() {
        if child_relationship.reltype() != ADD_IN_RELATIONSHIP {
            return Err(Error::Relationship(format!(
                "task-pane part has forbidden relationship '{}' of type '{}'",
                child_relationship.r_id(),
                child_relationship.reltype()
            )));
        }
        if !referenced_ids.contains(child_relationship.r_id()) {
            return Err(Error::Relationship(format!(
                "task-pane part has unreferenced relationship '{}'",
                child_relationship.r_id()
            )));
        }
    }

    let mut panes = Vec::with_capacity(parsed_panes.len());
    let mut total_snapshot_bytes = 0usize;
    let mut snapshot_names = HashSet::new();
    let mut extension_names = HashSet::new();
    for pane in parsed_panes {
        let child_relationship = task_panes_part
            .rels()
            .get(&pane.relationship_id)
            .ok_or_else(|| {
                Error::Relationship(format!(
                    "task pane references missing relationship '{}'",
                    pane.relationship_id
                ))
            })?;
        if child_relationship.is_external() {
            return Err(Error::Relationship(format!(
                "web extension relationship '{}' must be internal",
                pane.relationship_id
            )));
        }
        let extension_target = checked_internal_target(child_relationship, "add-in")?;
        let extension_name = index.canonical(&extension_target).ok_or_else(|| {
            Error::Missing(format!(
                "web extension part '{}'",
                extension_target.as_str()
            ))
        })?;
        let extension_part = package.get_part(extension_name).map_err(|error| {
            Error::Missing(format!(
                "web extension part '{}': {error}",
                extension_name.as_str()
            ))
        })?;
        let extension_name = extension_part.partname().clone();
        if !extension_names.insert(fold_part_name(&extension_name)) {
            return Err(Error::Relationship(format!(
                "multiple task panes target web extension part '{}'",
                extension_name.as_str()
            )));
        }
        require_content_type(extension_part, ADD_IN_CONTENT_TYPE)?;
        let add_in = parse_add_in_with_budget(extension_part.blob(), limits, budget)?;
        let snapshot_resources = load_snapshot_resources(
            package,
            extension_part,
            &add_in,
            &mut total_snapshot_bytes,
            &mut snapshot_names,
            limits,
            index,
        )?;
        panes.push(Pane {
            dock_state: pane.dock_state,
            visible: pane.visible,
            width: pane.width,
            row: pane.row,
            locked: pane.locked,
            relationship_id: pane.relationship_id,
            add_in,
            snapshot_resources,
            extension_list: pane.extension_list,
        });
    }
    let panes = Panes { panes };
    validate_panes(&panes, limits)?;
    Ok(Some(panes))
}
