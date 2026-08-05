//! Package-level orchestration for atomic workbook catalog snapshot edits.

use std::collections::{HashMap, HashSet};

use super::codec::{
    relationship_attribute_from_namespaces, relationship_attribute_name, rewrite_slot, scan,
    sibling_name, write_close, write_tag,
};
use super::model::{
    Active, Create, FIRST_SHEET_SENTINEL, Layout, MAX_ACTIVE_TAB, MAX_RELATIONSHIP_ID_CHARS,
    MAX_SHEET_ID, MAX_SHEETS, Plan, Remove, Rename, Replacement, Slot, Span, State, Tab, Tag,
};
use super::validation::{OrderMap, block, map_removed_position, validate_order};
use crate::error::{Error, Result, TabEditBlock, allocation, invalid};

pub(crate) fn rewrite(content: &[u8], plan: Plan<'_>) -> Result<Vec<u8>> {
    if plan.tabs.is_empty()
        && plan.renames.is_empty()
        && plan.active.is_none()
        && plan.order.is_none()
    {
        return Ok(content.to_vec());
    }
    let layout = scan(content)?;
    let Plan {
        tabs,
        renames,
        active,
        order,
    } = plan;
    let first = tabs.first().copied();
    let first_rename = renames.first().copied();
    let order_context = order.as_ref().map(|order| (order.sheet, order.position));
    let context = order_context
        .or_else(|| active.map(|active| (active.sheet, active.position)))
        .or_else(|| first.map(|tab| (tab.sheet, tab.position)))
        .or_else(|| first_rename.map(|rename| (rename.sheet, rename.position)));
    if layout.protected && (!tabs.is_empty() || !renames.is_empty() || order.is_some()) {
        return Err(block(context, TabEditBlock::ProtectedWorkbook));
    }
    if !renames.is_empty() && (layout.sheets.payload || layout.alternate_dependencies) {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if let Some(active) = active
        && active.position > MAX_ACTIVE_TAB
    {
        return Err(block(
            Some((active.sheet, active.position)),
            TabEditBlock::ActiveTabLimit,
        ));
    }

    let order_map = order
        .as_ref()
        .map(|order| validate_order(&layout, order))
        .transpose()?;

    let mut replacements = Vec::new();
    replacements
        .try_reserve(
            tabs.len()
                .saturating_add(renames.len())
                .saturating_add(order_map.as_ref().map_or(0, |_| layout.sheet_slots.len()))
                .saturating_add(layout.workbook_views.len())
                .saturating_add(layout.defined_name_slots.len()),
        )
        .map_err(|source| allocation("workbook edit plan", source))?;
    sheet_replacements(
        content,
        &layout,
        &tabs,
        &renames,
        order_map.as_ref(),
        &mut replacements,
    )?;

    if let Some(order_map) = order_map.as_ref() {
        view_replacements(
            content,
            &layout,
            order_map,
            active,
            context,
            &mut replacements,
        )?;
        defined_name_replacements(content, &layout, order_map, &mut replacements)?;
    } else if let Some(active) = active {
        replacements.push(active_replacement(
            content,
            &layout,
            active.position,
            Some((active.sheet, active.position)),
        )?);
    }
    replacements.sort_unstable_by_key(|replacement| replacement.span.start);
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping workbook edit replacements"));
    }

    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            let removed = replacement.span.end.checked_sub(replacement.span.start)?;
            size.checked_sub(removed)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("workbook edit output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("workbook edit output", source))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(output)
}

/// Append one checked sheet catalog entry while preserving every existing
/// element byte. Positional dependencies do not move for an append; activation
/// is intentionally handled by the ordinary view rewriter after insertion.
pub(crate) fn append(content: &[u8], create: Create<'_>) -> Result<Vec<u8>> {
    let layout = scan(content)?;
    let context = Some((create.sheet, create.position));
    if layout.protected {
        return Err(block(context, TabEditBlock::ProtectedWorkbook));
    }
    if layout.sheets.payload || layout.alternate_dependencies || layout.alternate_content {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if layout.sheet_slots.len() >= MAX_SHEETS {
        return Err(block(context, TabEditBlock::SheetLimit));
    }
    if create.position != layout.sheet_slots.len() {
        return Err(invalid("new worksheet position is not the catalog tail"));
    }
    if !(1..=MAX_SHEET_ID).contains(&create.sheet_id) {
        return Err(invalid("new worksheet native sheet ID is out of range"));
    }
    if create.relationship_id.is_empty()
        || create.relationship_id.chars().count() > MAX_RELATIONSHIP_ID_CHARS
    {
        return Err(invalid("new worksheet relationship ID is out of range"));
    }
    if layout
        .sheet_slots
        .iter()
        .any(|sheet| sheet.relationship_id.as_ref() == create.relationship_id)
    {
        return Err(invalid(
            "new worksheet relationship ID already exists in the catalog",
        ));
    }

    let sheet_name = layout.sheet_slots.first().map_or_else(
        || sibling_name(&layout.sheets.slot.tag.name, "sheet"),
        |sheet| sheet.slot.tag.name.to_string(),
    );
    let relationship_name = layout
        .sheet_slots
        .first()
        .and_then(relationship_attribute_name)
        .map(str::to_owned)
        .or_else(|| relationship_attribute_from_namespaces(&layout.root))
        .or_else(|| relationship_attribute_from_namespaces(&layout.sheets.slot.tag))
        .ok_or_else(|| block(context, TabEditBlock::MarkupCompatibility))?;
    let tag = Tag {
        name: sheet_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut created = Vec::new();
    let mut attributes = vec![
        ("name", create.sheet.to_owned()),
        ("sheetId", create.sheet_id.to_string()),
        (
            relationship_name.as_str(),
            create.relationship_id.to_owned(),
        ),
    ];
    if let Some(state) = create.state.attribute() {
        attributes.push(("state", state.to_owned()));
    }
    write_tag(&mut created, &tag, true, &[], &attributes);

    if layout.sheets.slot.empty {
        let mut replacement = Vec::new();
        write_tag(&mut replacement, &layout.sheets.slot.tag, false, &[], &[]);
        replacement.extend_from_slice(&created);
        write_close(&mut replacement, &layout.sheets.slot.tag.name);
        let mut output = Vec::new();
        output
            .try_reserve_exact(
                content
                    .len()
                    .checked_sub(layout.sheets.slot.span.end - layout.sheets.slot.span.start)
                    .and_then(|size| size.checked_add(replacement.len()))
                    .ok_or_else(|| invalid("workbook append output size overflow"))?,
            )
            .map_err(|source| allocation("workbook append output", source))?;
        output.extend_from_slice(&content[..layout.sheets.slot.span.start]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&content[layout.sheets.slot.span.end..]);
        return Ok(output);
    }

    let at = layout.sheets.slot.close_start;
    let mut output = Vec::new();
    output
        .try_reserve_exact(
            content
                .len()
                .checked_add(created.len())
                .ok_or_else(|| invalid("workbook append output size overflow"))?,
        )
        .map_err(|source| allocation("workbook append output", source))?;
    output.extend_from_slice(&content[..at]);
    output.extend_from_slice(&created);
    output.extend_from_slice(&content[at..]);
    Ok(output)
}

/// Remove checked catalog records, remap every modeled positional dependency,
/// and drop defined names whose local scope disappears.
pub(crate) fn remove(content: &[u8], plan: Remove<'_>) -> Result<Vec<u8>> {
    let layout = scan(content)?;
    let context = Some((plan.sheet, plan.position));
    if plan.relationship_ids.is_empty() {
        return Ok(content.to_vec());
    }
    if layout.protected {
        return Err(block(context, TabEditBlock::ProtectedWorkbook));
    }
    if layout.sheets.payload
        || layout
            .book_views
            .as_ref()
            .is_some_and(|views| views.payload)
        || layout
            .defined_names
            .as_ref()
            .is_some_and(|names| names.payload)
        || layout.alternate_content
        || layout.alternate_dependencies
        || layout
            .defined_name_slots
            .iter()
            .filter(|name| name.local_sheet_id.is_some())
            .count()
            != plan.local_scopes
    {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }

    let mut removed = HashSet::new();
    removed
        .try_reserve(plan.relationship_ids.len())
        .map_err(|source| allocation("sheet-removal index", source))?;
    for relationship_id in &plan.relationship_ids {
        if !removed.insert(*relationship_id) {
            return Err(invalid(format!(
                "sheet removal repeats relationship '{relationship_id}'"
            )));
        }
    }
    let mut old_to_new = Vec::new();
    old_to_new
        .try_reserve_exact(layout.sheet_slots.len())
        .map_err(|source| allocation("sheet-removal mapping", source))?;
    let mut next = 0usize;
    for slot in &layout.sheet_slots {
        if removed.remove(slot.relationship_id.as_ref()) {
            old_to_new.push(None);
        } else {
            old_to_new.push(Some(next));
            next = next
                .checked_add(1)
                .ok_or_else(|| invalid("retained worksheet count overflow"))?;
        }
    }
    if let Some(relationship_id) = removed.iter().next() {
        return Err(invalid(format!(
            "removed worksheet relationship '{relationship_id}' is absent from the catalog"
        )));
    }
    if next == 0 {
        return Err(invalid("workbook catalog removal cannot leave zero sheets"));
    }
    if plan.active.position >= next {
        return Err(invalid(
            "replacement active tab exceeds the retained catalog",
        ));
    }

    let removed_count = old_to_new
        .iter()
        .filter(|position| position.is_none())
        .count();
    let mut replacements = Vec::new();
    replacements
        .try_reserve(
            removed_count
                .saturating_add(layout.workbook_views.len().max(1))
                .saturating_add(layout.defined_name_slots.len()),
        )
        .map_err(|source| allocation("sheet removals", source))?;
    for (slot, mapped) in layout.sheet_slots.iter().zip(&old_to_new) {
        if mapped.is_none() {
            replacements.push(Replacement {
                span: slot.slot.span,
                bytes: Vec::new(),
            });
        }
    }

    if layout.workbook_views.is_empty() {
        replacements.push(active_replacement(
            content,
            &layout,
            plan.active.position,
            context,
        )?);
    } else {
        for (index, view) in layout.workbook_views.iter().enumerate() {
            let old_active = view.active.unwrap_or(0);
            let mapped_active = map_removed_position(&old_to_new, old_active, context)?;
            let desired_active = if index == 0 {
                plan.active.position
            } else {
                mapped_active
            };
            if desired_active > MAX_ACTIVE_TAB {
                return Err(block(context, TabEditBlock::ActiveTabLimit));
            }
            let old_first = view.first.unwrap_or(0);
            let desired_first = if old_first == FIRST_SHEET_SENTINEL {
                old_first
            } else {
                let old = usize::try_from(old_first)
                    .map_err(|_| block(context, TabEditBlock::ViewIndex))?;
                u32::try_from(map_removed_position(&old_to_new, old, context)?)
                    .map_err(|_| block(context, TabEditBlock::ViewIndex))?
            };
            let active_changed = view
                .active
                .map_or(desired_active != 0, |old| old != desired_active);
            let first_changed = view
                .first
                .map_or(desired_first != 0, |old| old != desired_first);
            if !active_changed && !first_changed {
                continue;
            }
            let mut removed_attributes = Vec::new();
            let mut appended = Vec::new();
            if active_changed {
                removed_attributes.push("activeTab");
                appended.push(("activeTab", desired_active.to_string()));
            }
            if first_changed {
                removed_attributes.push("firstSheet");
                appended.push(("firstSheet", desired_first.to_string()));
            }
            replacements.push(Replacement {
                span: view.slot.span,
                bytes: rewrite_slot(content, &view.slot, &removed_attributes, &appended),
            });
        }
    }

    let removed_names = layout
        .defined_name_slots
        .iter()
        .filter(|name| {
            name.local_sheet_id
                .and_then(|old| old_to_new.get(old))
                .is_some_and(Option::is_none)
        })
        .count();
    if removed_names != 0
        && removed_names == layout.defined_name_slots.len()
        && let Some(container) = &layout.defined_names
    {
        replacements.push(Replacement {
            span: container.slot.span,
            bytes: Vec::new(),
        });
    } else {
        for name in &layout.defined_name_slots {
            let Some(old) = name.local_sheet_id else {
                continue;
            };
            let Some(mapped) = old_to_new.get(old) else {
                return Err(invalid(
                    "defined-name scope exceeds the workbook sheet catalog during removal",
                ));
            };
            match mapped {
                None => replacements.push(Replacement {
                    span: name.slot.span,
                    bytes: Vec::new(),
                }),
                Some(new) if *new != old => replacements.push(Replacement {
                    span: name.slot.span,
                    bytes: rewrite_slot(
                        content,
                        &name.slot,
                        &["localSheetId"],
                        &[("localSheetId", new.to_string())],
                    ),
                }),
                Some(_) => {},
            }
        }
    }

    replacements.sort_unstable_by_key(|replacement| replacement.span.start);
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping workbook removal replacements"));
    }
    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            size.checked_sub(replacement.span.end - replacement.span.start)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("workbook removal output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("workbook removal output", source))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(output)
}

fn sheet_replacements(
    source: &[u8],
    layout: &Layout,
    tabs: &[Tab<'_>],
    renames: &[Rename<'_>],
    order: Option<&OrderMap>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    #[derive(Clone, Copy)]
    struct Update<'a> {
        sheet: &'a str,
        position: usize,
        state: Option<State>,
        name: Option<&'a str>,
    }

    let mut updates = HashMap::<&str, Update<'_>>::new();
    updates
        .try_reserve(tabs.len().saturating_add(renames.len()))
        .map_err(|source| allocation("tab update index", source))?;
    for tab in tabs {
        let update = updates.entry(tab.relationship_id).or_insert(Update {
            sheet: tab.sheet,
            position: tab.position,
            state: None,
            name: None,
        });
        if update.state.replace(tab.state).is_some() {
            return Err(invalid(format!(
                "duplicate tab state for relationship '{}'",
                tab.relationship_id
            )));
        }
    }
    for rename in renames {
        let update = updates.entry(rename.relationship_id).or_insert(Update {
            sheet: rename.sheet,
            position: rename.position,
            state: None,
            name: None,
        });
        if update.name.replace(rename.name).is_some() {
            return Err(invalid(format!(
                "duplicate tab rename for relationship '{}'",
                rename.relationship_id
            )));
        }
    }

    if let Some(order) = order {
        for (new, old) in order.new_to_old.iter().copied().enumerate() {
            let destination = &layout.sheet_slots[new];
            let selected = &layout.sheet_slots[old];
            let update = updates.remove(selected.relationship_id.as_ref());
            if new == old && update.is_none() {
                continue;
            }
            let bytes = update.map_or_else(
                || source[selected.slot.span.start..selected.slot.span.end].to_vec(),
                |update| sheet_replacement(source, &selected.slot, update.state, update.name),
            );
            replacements.push(Replacement {
                span: destination.slot.span,
                bytes,
            });
        }
    } else {
        for found in &layout.sheet_slots {
            let Some(update) = updates.remove(found.relationship_id.as_ref()) else {
                continue;
            };
            replacements.push(Replacement {
                span: found.slot.span,
                bytes: sheet_replacement(source, &found.slot, update.state, update.name),
            });
        }
    }
    if let Some(update) = updates.values().next() {
        return Err(Error::TabEditBlocked {
            sheet: update.sheet.to_owned(),
            position: update.position,
            reason: TabEditBlock::MarkupCompatibility,
        });
    }
    Ok(())
}

fn sheet_replacement(
    source: &[u8],
    slot: &Slot,
    state: Option<State>,
    name: Option<&str>,
) -> Vec<u8> {
    let mut removed = Vec::new();
    let mut appended = Vec::new();
    if let Some(state) = state {
        removed.push("state");
        if let Some(value) = state.attribute() {
            appended.push(("state", value.to_owned()));
        }
    }
    if let Some(name) = name {
        removed.push("name");
        appended.push(("name", name.to_owned()));
    }
    rewrite_slot(source, slot, &removed, &appended)
}

fn view_replacements(
    source: &[u8],
    layout: &Layout,
    order: &OrderMap,
    active: Option<Active<'_>>,
    context: Option<(&str, usize)>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    if layout.workbook_views.is_empty() {
        if let Some(active) = active {
            replacements.push(active_replacement(
                source,
                layout,
                active.position,
                context,
            )?);
        }
        return Ok(());
    }
    for (index, view) in layout.workbook_views.iter().enumerate() {
        let old_active = view.active.unwrap_or(0);
        let Some(&mapped_active) = order.old_to_new.get(old_active) else {
            return Err(block(context, TabEditBlock::ViewIndex));
        };
        let desired_active = if index == 0 {
            active.map_or(mapped_active, |active| active.position)
        } else {
            mapped_active
        };
        if desired_active > MAX_ACTIVE_TAB {
            return Err(block(context, TabEditBlock::ActiveTabLimit));
        }

        let old_first = view.first.unwrap_or(0);
        let desired_first = if old_first == FIRST_SHEET_SENTINEL {
            old_first
        } else {
            let old =
                usize::try_from(old_first).map_err(|_| block(context, TabEditBlock::ViewIndex))?;
            let mapped = order
                .old_to_new
                .get(old)
                .copied()
                .ok_or_else(|| block(context, TabEditBlock::ViewIndex))?;
            u32::try_from(mapped).map_err(|_| block(context, TabEditBlock::ViewIndex))?
        };

        let active_changed = view
            .active
            .map_or(desired_active != 0, |old| old != desired_active);
        let first_changed = view
            .first
            .map_or(desired_first != 0, |old| old != desired_first);
        if !active_changed && !first_changed {
            continue;
        }
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if active_changed {
            removed.push("activeTab");
            appended.push(("activeTab", desired_active.to_string()));
        }
        if first_changed {
            removed.push("firstSheet");
            appended.push(("firstSheet", desired_first.to_string()));
        }
        replacements.push(Replacement {
            span: view.slot.span,
            bytes: rewrite_slot(source, &view.slot, &removed, &appended),
        });
    }
    Ok(())
}

fn defined_name_replacements(
    source: &[u8],
    layout: &Layout,
    order: &OrderMap,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    for name in &layout.defined_name_slots {
        let Some(old) = name.local_sheet_id else {
            continue;
        };
        let Some(&new) = order.old_to_new.get(old) else {
            return Err(invalid(
                "defined-name scope exceeds the workbook sheet order during reorder",
            ));
        };
        if old == new {
            continue;
        }
        replacements.push(Replacement {
            span: name.slot.span,
            bytes: rewrite_slot(
                source,
                &name.slot,
                &["localSheetId"],
                &[("localSheetId", new.to_string())],
            ),
        });
    }
    Ok(())
}

fn active_replacement(
    source: &[u8],
    layout: &Layout,
    active: usize,
    context: Option<(&str, usize)>,
) -> Result<Replacement> {
    let appended = [("activeTab", active.to_string())];
    if let Some(view) = layout.workbook_views.first() {
        return Ok(Replacement {
            span: view.slot.span,
            bytes: rewrite_slot(source, &view.slot, &["activeTab"], &appended),
        });
    }
    if layout.alternate_content {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if let Some(book_views) = &layout.book_views {
        if book_views.payload {
            return Err(block(context, TabEditBlock::MarkupCompatibility));
        }
        let name = sibling_name(&book_views.slot.tag.name, "workbookView");
        let view = Tag {
            name: name.into_boxed_str(),
            attributes: Box::new([]),
        };
        let mut bytes = Vec::new();
        if book_views.slot.empty {
            write_tag(&mut bytes, &book_views.slot.tag, false, &[], &[]);
            write_tag(&mut bytes, &view, true, &[], &appended);
            write_close(&mut bytes, &book_views.slot.tag.name);
        } else {
            bytes.extend_from_slice(
                &source[book_views.slot.span.start..book_views.slot.close_start],
            );
            write_tag(&mut bytes, &view, true, &[], &appended);
            bytes.extend_from_slice(&source[book_views.slot.close_start..book_views.slot.span.end]);
        }
        return Ok(Replacement {
            span: book_views.slot.span,
            bytes,
        });
    }

    let book_views_name = sibling_name(&layout.sheets.slot.tag.name, "bookViews");
    let workbook_view_name = sibling_name(&book_views_name, "workbookView");
    let book_views = Tag {
        name: book_views_name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let workbook_view = Tag {
        name: workbook_view_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut bytes = Vec::new();
    write_tag(&mut bytes, &book_views, false, &[], &[]);
    write_tag(&mut bytes, &workbook_view, true, &[], &appended);
    write_close(&mut bytes, &book_views_name);
    Ok(Replacement {
        span: Span {
            start: layout.sheets.slot.span.start,
            end: layout.sheets.slot.span.start,
        },
        bytes,
    })
}
