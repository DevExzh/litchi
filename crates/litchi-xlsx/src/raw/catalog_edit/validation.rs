//! Validation and positional dependency mapping for catalog rewrites.

use std::collections::{HashMap, HashSet};

use super::model::{Layout, Order};
use crate::error::{Error, Result, TabEditBlock, allocation, invalid};

pub(super) struct OrderMap {
    pub(super) old_to_new: Vec<usize>,
    pub(super) new_to_old: Vec<usize>,
}

pub(super) fn map_removed_position(
    old_to_new: &[Option<usize>],
    old: usize,
    context: Option<(&str, usize)>,
) -> Result<usize> {
    let selected = old_to_new
        .get(old)
        .ok_or_else(|| block(context, TabEditBlock::ViewIndex))?;
    if let Some(mapped) = selected {
        return Ok(*mapped);
    }
    old_to_new
        .iter()
        .skip(old.saturating_add(1))
        .find_map(|mapped| *mapped)
        .or_else(|| old_to_new[..old].iter().rev().find_map(|mapped| *mapped))
        .ok_or_else(|| block(context, TabEditBlock::ViewIndex))
}

pub(super) fn validate_order(layout: &Layout, order: &Order<'_>) -> Result<OrderMap> {
    let context = Some((order.sheet, order.position));
    if layout.sheets.payload
        || layout
            .book_views
            .as_ref()
            .is_some_and(|views| views.payload)
        || layout
            .defined_names
            .as_ref()
            .is_some_and(|names| names.payload)
    {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if layout.alternate_dependencies
        || (layout.alternate_content && layout.workbook_views.is_empty())
        || layout
            .defined_name_slots
            .iter()
            .filter(|name| name.local_sheet_id.is_some())
            .count()
            != order.local_scopes
    {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if order.relationship_ids.len() != layout.sheet_slots.len() {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }

    let mut direct = HashMap::new();
    direct
        .try_reserve(layout.sheet_slots.len())
        .map_err(|source| allocation("sheet-order index", source))?;
    for (old, slot) in layout.sheet_slots.iter().enumerate() {
        if direct.insert(slot.relationship_id.as_ref(), old).is_some() {
            return Err(invalid(format!(
                "duplicate direct workbook sheet relationship '{}' during reorder",
                slot.relationship_id
            )));
        }
    }

    let mut seen = HashSet::new();
    seen.try_reserve(order.relationship_ids.len())
        .map_err(|source| allocation("sheet-order validation", source))?;
    let mut old_to_new = Vec::new();
    old_to_new
        .try_reserve_exact(order.relationship_ids.len())
        .map_err(|source| allocation("reverse sheet-order mapping", source))?;
    old_to_new.resize(order.relationship_ids.len(), 0usize);
    let mut new_to_old = Vec::new();
    new_to_old
        .try_reserve_exact(order.relationship_ids.len())
        .map_err(|source| allocation("sheet-order mapping", source))?;
    for (new, relationship_id) in order.relationship_ids.iter().copied().enumerate() {
        if !seen.insert(relationship_id) {
            return Err(invalid(format!(
                "sheet reorder repeats relationship '{relationship_id}'"
            )));
        }
        let Some(&old) = direct.get(relationship_id) else {
            return Err(block(context, TabEditBlock::MarkupCompatibility));
        };
        old_to_new[old] = new;
        new_to_old.push(old);
    }
    Ok(OrderMap {
        old_to_new,
        new_to_old,
    })
}

pub(super) fn block(context: Option<(&str, usize)>, reason: TabEditBlock) -> Error {
    context.map_or_else(
        || invalid("workbook catalog rewrite has no associated tab change"),
        |(sheet, position)| Error::TabEditBlocked {
            sheet: sheet.to_owned(),
            position,
            reason,
        },
    )
}
