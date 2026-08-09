//! Dependency guards for atomic worksheet edits.

use std::collections::BTreeMap;

use litchi_sheet::{Cell as Address, Column, Rect, Row};

use super::codec::{CellSlot, DimensionTag, Layout, SheetData};
use super::model::{
    Action, ColumnAction, DefaultsAction, DescentEffect, Payload, Plan, RowAction, StyleEffect,
    WidthEffect,
};
use crate::column::Assignments;
use crate::error::{ColumnEditBlock, DefaultsEditBlock, EditBlock, Error, Result, RowEditBlock};
pub(super) fn plan_sets_descent(plan: &Plan) -> bool {
    plan.rows
        .values()
        .any(|action| matches!(action.descent, Some(DescentEffect::Set(_))))
        || plan
            .defaults
            .is_some_and(|action| matches!(action.effects().descent, Some(DescentEffect::Set(_))))
}

pub(super) fn validate_defaults_action(
    layout: &Layout,
    sheet: &str,
    action: Option<DefaultsAction>,
) -> Result<()> {
    let Some(action) = action else {
        return Ok(());
    };
    let reason = if layout.protected {
        Some(DefaultsEditBlock::ProtectedSheet)
    } else if layout.defaults_compatibility {
        Some(DefaultsEditBlock::MarkupCompatibility)
    } else if layout.defaults.is_none()
        && action.materializes()
        && action.effects().height.is_none()
    {
        Some(DefaultsEditBlock::NeedsHeight)
    } else {
        None
    };
    if let Some(reason) = reason {
        return Err(Error::DefaultsEditBlocked {
            sheet: sheet.to_owned(),
            reason,
        });
    }
    Ok(())
}

pub(super) fn validate_column_actions(
    layout: &Layout,
    sheet: &str,
    actions: &BTreeMap<Column, ColumnAction>,
) -> Result<()> {
    if layout.protected
        && let Some(column) = actions.keys().next()
    {
        return Err(Error::ColumnEditBlocked {
            sheet: sheet.to_owned(),
            column: *column,
            reason: ColumnEditBlock::ProtectedSheet,
        });
    }
    let mut owners = Assignments::new()?;
    if let Some(stored) = &layout.columns {
        for (index, column) in stored.columns.iter().enumerate() {
            owners.assign(column.first, column.last, index);
        }
    }
    for (column, action) in actions {
        if !matches!(action.style, Some(StyleEffect::Set(_)))
            || matches!(action.width, Some(WidthEffect::Set(_)))
        {
            continue;
        }
        let has_width = owners
            .get(*column)
            .and_then(|index| layout.columns.as_ref()?.columns.get(index))
            .is_some_and(|stored| {
                stored
                    .tag
                    .attributes
                    .iter()
                    .any(|attribute| attribute.name.as_ref() == "width")
            });
        if !has_width || matches!(action.width, Some(WidthEffect::Reset)) {
            return Err(Error::ColumnEditBlocked {
                sheet: sheet.to_owned(),
                column: *column,
                reason: ColumnEditBlock::StyleNeedsWidth,
            });
        }
    }
    Ok(())
}

pub(super) fn validate_row_actions(
    layout: &Layout,
    sheet: &str,
    actions: &BTreeMap<Row, RowAction>,
) -> Result<()> {
    if layout.protected
        && let Some(row) = actions.keys().next()
    {
        return Err(Error::RowEditBlocked {
            sheet: sheet.to_owned(),
            row: *row,
            reason: RowEditBlock::ProtectedSheet,
        });
    }
    Ok(())
}

pub(super) fn validate_actions(
    layout: &Layout,
    sheet: &str,
    actions: &BTreeMap<Address, Action>,
) -> Result<()> {
    for (address, action) in actions {
        let blocked = if layout.protected {
            Some(EditBlock::ProtectedSheet)
        } else if layout.extended_validation
            || layout
                .validations
                .iter()
                .any(|range| range.contains(*address))
        {
            Some(EditBlock::DataValidation)
        } else if layout
            .formula_ranges
            .iter()
            .any(|range| range.contains(*address))
        {
            Some(EditBlock::GroupFormula)
        } else if layout
            .merged
            .iter()
            .any(|range| range.contains(*address) && !range.starts_at(*address))
        {
            Some(EditBlock::CoveredMerge)
        } else if cell_slot(&layout.sheet_data, *address).is_some_and(|cell| cell.mce_payload) {
            Some(EditBlock::MarkupCompatibility)
        } else {
            None
        };
        if let Some(reason) = blocked {
            return Err(Error::EditBlocked {
                sheet: sheet.to_owned(),
                address: *address,
                reason,
            });
        }
        if let Some(Payload::Set(content)) = action.payload() {
            content.validate_for_write()?;
        }
        if let Some(Payload::SharedString { text, .. }) = action.payload() {
            crate::Content::Value(crate::Value::Text(text.clone())).validate_for_write()?;
        }
    }
    Ok(())
}

fn cell_slot(sheet_data: &SheetData, address: Address) -> Option<&CellSlot> {
    let row = sheet_data
        .rows
        .binary_search_by_key(&(address.row().get() + 1), |row| row.number)
        .ok()
        .and_then(|index| sheet_data.rows.get(index))?;
    row.cells
        .binary_search_by_key(&address, |cell| cell.address)
        .ok()
        .and_then(|index| row.cells.get(index))
}

#[derive(Debug, Default)]
struct CellBounds(Option<Rect>);

impl CellBounds {
    fn push(&mut self, address: Address) {
        let cell = Rect::single(address);
        self.0 = Some(self.0.map_or(cell, |range| range.union(cell)));
    }
}

pub(super) fn expanded_dimension<'a>(
    layout: &'a Layout,
    actions: &BTreeMap<Address, Action>,
) -> Option<(&'a DimensionTag, Rect)> {
    let dimension = layout.dimension.as_ref()?;
    let mut bounds = CellBounds::default();
    for row in &layout.sheet_data.rows {
        for cell in &row.cells {
            if !matches!(actions.get(&cell.address), Some(Action::Remove)) {
                bounds.push(cell.address);
            }
        }
    }
    for (address, action) in actions {
        if action.creates_missing() {
            bounds.push(*address);
        }
    }
    let result = bounds.0?;
    let expanded = dimension.declared.union(result);
    (expanded != dimension.declared).then_some((dimension, expanded))
}
