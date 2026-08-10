//! Workbook-package orchestration for worksheet removal and graph checks.

use super::codec::{
    calculation_chain_removal, compose_part, compose_part_optional, removal_reference_part,
};
use super::{
    ActiveTab, Arc, BTreeSet, Change, Commit, Edit, Error, GraphAction, GraphChange, HashSet,
    OpcPackage, OrderPlan, PackURI, Part, PartChange, Patch, Relationship, RemoveBlock, Result,
    TabAction, TabEditBlock, Target, Visibility, Workbook, WorksheetKind, allocation, invalid, raw,
};

pub(super) fn commit_removals(edit: Edit) -> Result<Commit> {
    if edit.has_non_removal() {
        return Err(edit.remove_block(RemoveBlock::MixedEdit, "transaction"));
    }
    let Edit {
        base,
        panes: _,
        defined_names: _,
        drawings: _,
        active: _,
        order: _,
        sheets: _,
        added: _,
        removed,
    } = edit;
    let first_position = removed
        .iter()
        .next()
        .copied()
        .ok_or_else(|| invalid("worksheet removal plan is empty"))?;
    let first_sheet = base
        .inner
        .sheets
        .get(first_position)
        .ok_or_else(|| invalid("removed worksheet position disappeared"))?;
    let block = |reason, part: &str| Error::SheetRemoveBlocked {
        sheet: first_sheet.name.clone(),
        position: first_position,
        part: part.to_owned(),
        reason,
    };
    let retained_len = base
        .inner
        .sheets
        .len()
        .checked_sub(removed.len())
        .ok_or_else(|| invalid("removed worksheet count exceeds the catalog"))?;
    if retained_len == 0 {
        return Err(block(
            RemoveBlock::LastSheet,
            base.inner.workbook_uri.as_str(),
        ));
    }
    ensure_reorder_supported(&base, &first_sheet.name, first_position)?;

    for position in &removed {
        let sheet = base
            .inner
            .sheets
            .get(*position)
            .ok_or_else(|| invalid("removed worksheet position disappeared"))?;
        if sheet.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name.clone(),
            });
        }
    }

    let main = base.inner.package.get_part(&base.inner.workbook_uri)?;
    if let Some(relationship) = main.rels().iter().find(|relationship| {
        relationship.reltype() == litchi_opc::constants::relationship_type::VBA_PROJECT
    }) {
        return Err(block(RemoveBlock::MacroProject, relationship.target_ref()));
    }

    let visible = base
        .inner
        .sheets
        .iter()
        .enumerate()
        .filter(|(position, sheet)| {
            !removed.contains(position) && sheet.visibility == Visibility::Visible
        })
        .map(|(position, _)| position)
        .collect::<Vec<_>>();
    if visible.is_empty() {
        return Err(Error::TabEditBlocked {
            sheet: first_sheet.name.clone(),
            position: first_position,
            reason: TabEditBlock::LastVisibleTab,
        });
    }
    let current_active = base
        .inner
        .active_sheet
        .ok_or_else(|| invalid("non-empty workbook has no active tab"))?;
    let final_active_identity = if !removed.contains(&current_active)
        && base
            .inner
            .sheets
            .get(current_active)
            .is_some_and(|sheet| sheet.visibility == Visibility::Visible)
    {
        current_active
    } else {
        visible
            .iter()
            .copied()
            .find(|position| *position > current_active)
            .or_else(|| {
                visible
                    .iter()
                    .rev()
                    .copied()
                    .find(|position| *position < current_active)
            })
            .ok_or_else(|| invalid("replacement active worksheet disappeared"))?
    };
    let final_active_position = (0..final_active_identity)
        .filter(|position| !removed.contains(position))
        .count();
    if final_active_position > raw::catalog_edit::MAX_ACTIVE_TAB {
        let sheet = base
            .inner
            .sheets
            .get(final_active_identity)
            .ok_or_else(|| invalid("replacement active worksheet disappeared"))?;
        return Err(Error::TabEditBlocked {
            sheet: sheet.name.clone(),
            position: final_active_position,
            reason: TabEditBlock::ActiveTabLimit,
        });
    }
    let active_sheet = base
        .inner
        .sheets
        .get(final_active_identity)
        .ok_or_else(|| invalid("replacement active worksheet disappeared"))?;

    let removed_relationship_ids = removed
        .iter()
        .map(|position| {
            base.inner
                .sheets
                .get(*position)
                .map(|sheet| sheet.relationship_id.as_str())
                .ok_or_else(|| invalid("removed worksheet relationship disappeared"))
        })
        .collect::<Result<Vec<_>>>()?;
    let local_scopes = base
        .inner
        .defined_names
        .iter()
        .filter(|name| name.local_sheet_id.is_some())
        .count();
    let before_workbook = main.blob_arc();
    let mut after_workbook = raw::catalog_edit::remove(
        &before_workbook,
        raw::catalog_edit::Remove {
            sheet: &first_sheet.name,
            position: first_position,
            relationship_ids: removed_relationship_ids.clone(),
            active: raw::catalog_edit::Active {
                sheet: &active_sheet.name,
                position: final_active_position,
            },
            local_scopes,
        },
    )?;
    after_workbook = raw::recalc::invalidate(&after_workbook)?;
    let catalog = raw::parse_catalog(&after_workbook)?;
    if catalog.sheets.len() != retained_len || catalog.active_sheet_index != final_active_position {
        return Err(invalid("workbook worksheet-removal verification failed"));
    }
    let retained = base
        .inner
        .sheets
        .iter()
        .enumerate()
        .filter(|(position, _)| !removed.contains(position));
    for (actual, (_, expected)) in catalog.sheets.iter().zip(retained) {
        if actual.relationship_id != expected.relationship_id || actual.name != expected.name {
            return Err(invalid(
                "workbook worksheet-removal verification changed a retained tab",
            ));
        }
    }
    verify_removed_defined_names(&base, &catalog, &removed)?;

    let mut changes = Vec::new();
    changes
        .try_reserve(removed.len().saturating_add(1))
        .map_err(|source| allocation("worksheet-removal changes", source))?;
    for position in &removed {
        let sheet = base
            .inner
            .sheets
            .get(*position)
            .ok_or_else(|| invalid("removed worksheet disappeared during patch creation"))?;
        changes.push(Change::Remove {
            sheet: sheet.name.clone().into_boxed_str(),
            position: *position,
            visibility: sheet.visibility.clone(),
        });
    }
    let active_before = active_tab_at(&base, current_active, current_active, None)?;
    let active_after = active_tab_at(&base, final_active_identity, final_active_position, None)?;
    if active_before != active_after {
        changes.push(Change::Active {
            before: active_before,
            after: active_after,
        });
    }

    let mut parts = vec![PartChange {
        uri: base.inner.workbook_uri.clone(),
        before: before_workbook,
        after: Arc::new(after_workbook),
    }];
    if final_active_identity != current_active {
        if !removed.contains(&current_active) {
            let old_active = base
                .inner
                .sheets
                .get(current_active)
                .ok_or_else(|| invalid("previous active worksheet disappeared"))?;
            compose_part(&mut parts, &base, &old_active.part_uri, |content| {
                raw::sheet_view_edit::rewrite(
                    content,
                    false,
                    raw::sheet_view_edit::Context {
                        sheet: &old_active.name,
                        position: current_active,
                    },
                )
            })?;
        }
        compose_part(&mut parts, &base, &active_sheet.part_uri, |content| {
            raw::sheet_view_edit::rewrite(
                content,
                true,
                raw::sheet_view_edit::Context {
                    sheet: &active_sheet.name,
                    position: final_active_position,
                },
            )
        })?;
    }

    let existing_titles = base
        .inner
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    let removed_positions = removed.iter().copied().collect::<Vec<_>>();
    let mut property_parts = base
        .inner
        .package
        .iter_parts()
        .filter(|part| {
            part.content_type() == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES
        })
        .map(|part| part.partname().clone())
        .collect::<Vec<_>>();
    property_parts.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    for uri in property_parts {
        compose_part_optional(&mut parts, &base, &uri, |content| {
            raw::properties_edit::remove_sheets(content, &existing_titles, &removed_positions)
        })?;
    }

    let mut graph = Vec::new();
    graph
        .try_reserve(removed.len().saturating_add(1))
        .map_err(|source| allocation("worksheet graph removals", source))?;
    for position in &removed {
        let sheet = base
            .inner
            .sheets
            .get(*position)
            .ok_or_else(|| invalid("removed worksheet disappeared during graph planning"))?;
        let relationship = main.rels().get(&sheet.relationship_id).ok_or_else(|| {
            invalid(format!(
                "worksheet '{}' relationship disappeared",
                sheet.name
            ))
        })?;
        if relationship.is_external()
            || !relationship
                .target_partname()?
                .is_equivalent_to(&sheet.part_uri)
        {
            return Err(invalid(format!(
                "worksheet '{}' relationship target changed",
                sheet.name
            )));
        }
        ensure_exclusive_sheet_incoming(
            &base.inner.package,
            &sheet.part_uri,
            &base.inner.workbook_uri,
            relationship.r_id(),
            &sheet.name,
            *position,
        )?;
        graph.push(GraphChange {
            action: GraphAction::Remove,
            source: base.inner.workbook_uri.clone(),
            relationship: relationship.clone(),
            part: base.inner.package.get_part(&sheet.part_uri)?.clone_part(),
        });
    }
    graph.extend(calculation_chain_removal(&base)?);

    let detached_workbook_relationships = graph
        .iter()
        .filter(|change| change.source == base.inner.workbook_uri)
        .map(|change| change.relationship.r_id())
        .collect::<Vec<_>>();
    scan_removal_dependencies(&base, &parts, &removed, &detached_workbook_relationships)?;

    let mut package = base.inner.package.clone();
    for part in &parts {
        package
            .get_part_mut(&part.uri)?
            .set_blob_shared(Arc::clone(&part.after));
    }
    for change in &graph {
        change.validate(&package)?;
        change.apply(&mut package)?;
    }
    let workbook = Workbook::from_package_with_styles(package, Some(&base))?;
    Ok(Commit {
        workbook: workbook.clone(),
        patch: Patch {
            changes: changes.into_boxed_slice(),
            package_changes: Box::new([]),
            parts: parts.into_boxed_slice(),
            graph: graph.into_boxed_slice(),
            web: None,
            style_guard: None,
            source: Some(base),
            target: Some(workbook),
        },
    })
}

fn verify_removed_defined_names(
    workbook: &Workbook,
    catalog: &raw::Catalog,
    removed: &BTreeSet<usize>,
) -> Result<()> {
    let mut expected = Vec::new();
    expected
        .try_reserve_exact(workbook.inner.defined_names.len())
        .map_err(|source| allocation("defined-name verification", source))?;
    for name in &workbook.inner.defined_names {
        let scope = name
            .local_sheet_id
            .map(|scope| {
                usize::try_from(scope).map_err(|_| invalid("defined-name scope does not fit usize"))
            })
            .transpose()?;
        if scope.is_some_and(|scope| removed.contains(&scope)) {
            continue;
        }
        let mapped = scope.map(|scope| {
            u32::try_from(
                (0..scope)
                    .filter(|position| !removed.contains(position))
                    .count(),
            )
            .map_err(|_| invalid("remapped defined-name scope does not fit u32"))
        });
        expected.push((name, mapped.transpose()?));
    }
    if expected.len() != catalog.defined_names.len() {
        return Err(invalid(
            "workbook removal changed the effective defined-name count unexpectedly",
        ));
    }
    for ((before, scope), after) in expected.into_iter().zip(&catalog.defined_names) {
        if after.local_sheet_id != scope || !same_defined_name_except_scope(before, after) {
            return Err(invalid(format!(
                "workbook removal changed defined name '{}' unexpectedly",
                before.name
            )));
        }
    }
    Ok(())
}

fn ensure_exclusive_sheet_incoming(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_id: &str,
    sheet: &str,
    position: usize,
) -> Result<()> {
    let blocked = |part: &str| Error::SheetRemoveBlocked {
        sheet: sheet.to_owned(),
        position,
        part: part.to_owned(),
        reason: RemoveBlock::IncomingRelationship,
    };
    let targets = |relationship: &Relationship| -> Result<bool> {
        if relationship.is_external() {
            return Ok(false);
        }
        relationship
            .target_partname()
            .map(|candidate| candidate.as_str().eq_ignore_ascii_case(target.as_str()))
            .map_err(Into::into)
    };
    for relationship in package.rels().iter() {
        if targets(relationship)? {
            return Err(blocked("/"));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if targets(relationship)?
                && !(source.partname() == expected_source && relationship.r_id() == expected_id)
            {
                return Err(blocked(source.partname().as_str()));
            }
        }
    }
    Ok(())
}

fn scan_removal_dependencies(
    workbook: &Workbook,
    parts: &[PartChange],
    removed: &BTreeSet<usize>,
    removed_relationship_ids: &[&str],
) -> Result<()> {
    let catalog = workbook
        .inner
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    let targets = removed
        .iter()
        .map(|position| {
            let sheet =
                workbook.inner.sheets.get(*position).ok_or_else(|| {
                    invalid("removed worksheet disappeared during dependency scan")
                })?;
            Ok(raw::reference_scan::Sheet {
                name: &sheet.name,
                position: *position,
                native_id: sheet.native_id,
                catalog: &catalog,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let reachable = reachable_after_removal(workbook, removed_relationship_ids)?;
    for uri in &reachable {
        let part = workbook.inner.package.get_part(uri)?;
        if !removal_reference_part(part) {
            continue;
        }
        let content = parts
            .iter()
            .find(|change| &change.uri == uri)
            .map_or(part.blob(), |change| change.after.as_slice());
        let Some(hit) = raw::reference_scan::scan(content, &targets)? else {
            continue;
        };
        let sheet = targets
            .get(hit.target)
            .ok_or_else(|| invalid("dependency scan returned an unknown removal target"))?;
        let reason = match hit.dependency {
            raw::reference_scan::Dependency::Modeled => RemoveBlock::IncomingReference,
            raw::reference_scan::Dependency::Unmodeled => RemoveBlock::UnmodeledReference,
            raw::reference_scan::Dependency::MarkupCompatibility => {
                RemoveBlock::MarkupCompatibility
            },
        };
        return Err(Error::SheetRemoveBlocked {
            sheet: sheet.name.to_owned(),
            position: sheet.position,
            part: uri.to_string(),
            reason,
        });
    }
    Ok(())
}

fn reachable_after_removal(
    workbook: &Workbook,
    removed_relationship_ids: &[&str],
) -> Result<Vec<PackURI>> {
    let mut reachable = HashSet::<PackURI>::new();
    reachable
        .try_reserve(workbook.inner.package.part_count())
        .map_err(|source| allocation("reachable package graph", source))?;
    let mut pending = Vec::<PackURI>::new();
    for relationship in workbook.inner.package.rels().iter() {
        if !relationship.is_external() {
            let target = relationship.target_partname()?;
            let part = workbook.inner.package.get_part(&target)?;
            pending.push(part.partname().clone());
        }
    }
    while let Some(uri) = pending.pop() {
        if !reachable.insert(uri.clone()) {
            continue;
        }
        let part = workbook.inner.package.get_part(&uri)?;
        for relationship in part.rels().iter() {
            if uri == workbook.inner.workbook_uri
                && removed_relationship_ids.contains(&relationship.r_id())
            {
                continue;
            }
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            let target = workbook.inner.package.get_part(&target)?.partname().clone();
            if !reachable.contains(&target) {
                pending.push(target);
            }
        }
    }
    let mut reachable = reachable.into_iter().collect::<Vec<_>>();
    reachable.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    Ok(reachable)
}

pub(super) fn raw_visibility_matches(value: &raw::Visibility, action: TabAction) -> bool {
    matches!(
        (value, action),
        (raw::Visibility::Visible, TabAction::Show)
            | (raw::Visibility::Hidden, TabAction::Hide)
            | (raw::Visibility::VeryHidden, TabAction::VeryHide)
    )
}

pub(super) fn active_tab_at(
    workbook: &Workbook,
    identity: usize,
    position: usize,
    name: Option<&str>,
) -> Result<ActiveTab> {
    let sheet = workbook
        .inner
        .sheets
        .get(identity)
        .ok_or_else(|| invalid("active tab points outside the workbook catalog"))?;
    Ok(ActiveTab {
        name: name.unwrap_or(&sheet.name).into(),
        position,
    })
}

pub(super) fn validate_order_plan(order: &OrderPlan, len: usize) -> Result<()> {
    if order.positions.len() != len {
        return Err(invalid("tab-order plan has the wrong number of sheets"));
    }
    let mut seen = Vec::new();
    seen.try_reserve_exact(len)
        .map_err(|source| allocation("tab-order validation", source))?;
    seen.resize(len, false);
    for identity in &order.positions {
        let Some(slot) = seen.get_mut(*identity) else {
            return Err(invalid("tab-order plan contains an out-of-range identity"));
        };
        if *slot {
            return Err(invalid("tab-order plan is not a permutation"));
        }
        *slot = true;
    }

    let mut replay = Vec::new();
    replay
        .try_reserve_exact(len)
        .map_err(|source| allocation("tab-move replay", source))?;
    replay.extend(0..len);
    for moved in &order.moves {
        if moved.from >= replay.len() || moved.to >= replay.len() {
            return Err(invalid("tab move contains an out-of-range position"));
        }
        if replay[moved.from] != moved.sheet {
            return Err(invalid("tab move source does not match the pending order"));
        }
        let identity = replay.remove(moved.from);
        replay.insert(moved.to, identity);
    }
    if replay != order.positions {
        return Err(invalid("tab moves do not produce the pending final order"));
    }
    Ok(())
}

pub(super) fn ensure_reorder_supported(
    workbook: &Workbook,
    sheet: &str,
    position: usize,
) -> Result<()> {
    let main = workbook
        .inner
        .package
        .get_part(&workbook.inner.workbook_uri)?;
    if main
        .rels()
        .iter()
        .any(|relationship| relationship.reltype().ends_with("/revisionHeaders"))
    {
        return Err(Error::TabEditBlocked {
            sheet: sheet.to_owned(),
            position,
            reason: TabEditBlock::TrackedWorkbook,
        });
    }
    Ok(())
}

pub(super) fn verify_defined_name_scopes(
    source: &raw::Catalog,
    catalog: &raw::Catalog,
    base_len: usize,
    order: &[Target],
) -> Result<()> {
    if catalog.defined_names.len() != source.defined_names.len() {
        return Err(invalid(
            "workbook reorder changed the effective defined-name count",
        ));
    }
    let mut old_to_new = Vec::new();
    old_to_new
        .try_reserve_exact(base_len)
        .map_err(|source| allocation("defined-name scope map", source))?;
    old_to_new.resize(base_len, usize::MAX);
    for (new, target) in order.iter().copied().enumerate() {
        let Target::Base(old) = target else {
            continue;
        };
        let slot = old_to_new
            .get_mut(old)
            .ok_or_else(|| invalid("defined-name scope map has an invalid sheet identity"))?;
        if *slot != usize::MAX {
            return Err(invalid("defined-name scope map repeats a sheet identity"));
        }
        *slot = new;
    }
    if old_to_new.contains(&usize::MAX) {
        return Err(invalid("defined-name scope map omits a sheet identity"));
    }
    for (before, after) in source.defined_names.iter().zip(&catalog.defined_names) {
        let expected_scope = match before.local_sheet_id {
            None => None,
            Some(scope) => {
                let scope = usize::try_from(scope)
                    .map_err(|_| invalid("defined-name scope does not fit usize"))?;
                let mapped = old_to_new
                    .get(scope)
                    .copied()
                    .ok_or_else(|| invalid("defined-name scope cannot be remapped"))?;
                Some(
                    u32::try_from(mapped)
                        .map_err(|_| invalid("remapped defined-name scope does not fit u32"))?,
                )
            },
        };
        if after.local_sheet_id != expected_scope || !same_defined_name_except_scope(before, after)
        {
            return Err(invalid(format!(
                "workbook reorder changed defined name '{}' unexpectedly",
                before.name
            )));
        }
    }
    Ok(())
}

fn same_defined_name_except_scope(left: &raw::DefinedName, right: &raw::DefinedName) -> bool {
    left.name == right.name
        && left.reference == right.reference
        && left.comment == right.comment
        && left.custom_menu == right.custom_menu
        && left.description == right.description
        && left.help == right.help
        && left.status_bar == right.status_bar
        && left.shortcut_key == right.shortcut_key
        && left.hidden == right.hidden
        && left.function == right.function
        && left.vb_procedure == right.vb_procedure
        && left.xlm == right.xlm
        && left.function_group_id == right.function_group_id
        && left.publish_to_server == right.publish_to_server
        && left.workbook_parameter == right.workbook_parameter
}
