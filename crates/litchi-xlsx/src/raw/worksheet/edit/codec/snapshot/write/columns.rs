//! Column interval snapshot writer.

use std::collections::{BTreeMap, HashMap};

use litchi_sheet::Column;

use super::super::super::wire::{sibling_name, write_close, write_tag};
use super::super::model::{ColumnSlot, ColumnsSlot, Tag};
use crate::column::Assignments;
use crate::error::{ColumnEditBlock, Error, Result, allocation, invalid};
use crate::outline::Outline;
use crate::raw::worksheet::edit::model::{ColumnAction, StyleEffect, WidthEffect};

#[derive(Debug, Clone, Copy)]
enum ColumnPiece {
    Keep(Column, Column),
    Edit(Column, Column, ColumnAction),
}

pub(crate) fn write_columns(
    output: &mut Vec<u8>,
    source: &[u8],
    stored: &ColumnsSlot,
    actions: BTreeMap<Column, ColumnAction>,
    sheet: &str,
) -> Result<()> {
    let mut owners = Assignments::new()?;
    for (index, column) in stored.columns.iter().enumerate() {
        owners.assign(column.first, column.last, index);
    }
    let mut by_owner = HashMap::<usize, BTreeMap<Column, ColumnAction>>::new();
    let mut implicit = BTreeMap::new();
    for (column, action) in actions {
        if let Some(owner) = owners.get(column) {
            by_owner.entry(owner).or_default().insert(column, action);
        } else if action.materializes() {
            implicit.insert(column, action);
        }
    }

    if stored.payload
        && let Some(column) = implicit.keys().next()
    {
        return Err(Error::ColumnEditBlocked {
            sheet: sheet.to_owned(),
            column: *column,
            reason: ColumnEditBlock::MarkupCompatibility,
        });
    }

    if stored.empty {
        return Err(invalid("worksheet cols contains no col during edit"));
    }
    output.extend_from_slice(&source[stored.span.start..stored.tag_end]);
    let mut cursor = stored.tag_end;
    for (index, column) in stored.columns.iter().enumerate() {
        output.extend_from_slice(&source[cursor..column.span.start]);
        if let Some(edits) = by_owner.remove(&index) {
            let pieces = column_pieces(column, &edits)?;
            if column.payload && pieces.len() > 1 {
                let edited = edits.keys().next().copied().unwrap_or(column.first);
                return Err(Error::ColumnEditBlocked {
                    sheet: sheet.to_owned(),
                    column: edited,
                    reason: ColumnEditBlock::MarkupCompatibility,
                });
            }
            for piece in pieces {
                write_column_piece(output, source, column, piece);
            }
        } else {
            output.extend_from_slice(&source[column.span.start..column.span.end]);
        }
        cursor = column.span.end;
    }
    output.extend_from_slice(&source[cursor..stored.close_start]);
    write_column_actions(output, &stored.tag.name, implicit);
    output.extend_from_slice(&source[stored.close_start..stored.span.end]);
    Ok(())
}

fn column_pieces(
    stored: &ColumnSlot,
    edits: &BTreeMap<Column, ColumnAction>,
) -> Result<Vec<ColumnPiece>> {
    let capacity = edits
        .len()
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| invalid("column edit split count overflow"))?;
    let mut pieces = Vec::new();
    pieces
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("column edit splits", source))?;
    let mut next = stored.first.get();
    for (column, action) in edits {
        if column.get() > next {
            pieces.push(ColumnPiece::Keep(
                Column::new(next)?,
                Column::new(column.get() - 1)?,
            ));
        }
        if let Some(ColumnPiece::Edit(_, last, previous)) = pieces.last_mut()
            && previous == action
            && last.next() == Some(*column)
        {
            *last = *column;
        } else {
            pieces.push(ColumnPiece::Edit(*column, *column, *action));
        }
        next = column.get().saturating_add(1);
    }
    if next <= stored.last.get() {
        pieces.push(ColumnPiece::Keep(Column::new(next)?, stored.last));
    }
    Ok(pieces)
}

fn write_column_piece(
    output: &mut Vec<u8>,
    source: &[u8],
    stored: &ColumnSlot,
    piece: ColumnPiece,
) {
    let (first, last, action) = match piece {
        ColumnPiece::Keep(first, last) => (first, last, None),
        ColumnPiece::Edit(first, last, action) => (first, last, Some(action)),
    };
    let mut removed = vec!["min", "max"];
    let mut appended = vec![
        ("min", (first.get() + 1).to_string()),
        ("max", (last.get() + 1).to_string()),
    ];
    if let Some(action) = action {
        column_effect_attributes(action, &mut removed, &mut appended);
    }
    write_tag(output, &stored.tag, stored.empty, &removed, &appended);
    if !stored.empty {
        output.extend_from_slice(&source[stored.tag_end..stored.close_start]);
        write_close(output, &stored.tag.name);
    }
}

pub(crate) fn write_new_columns(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    actions: BTreeMap<Column, ColumnAction>,
) {
    if !actions.values().any(|action| action.materializes()) {
        return;
    }
    let name = sibling_name(sheet_data_name, "cols");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    write_tag(output, &tag, false, &[], &[]);
    write_column_actions(output, &name, actions);
    write_close(output, &name);
}

fn write_column_actions(
    output: &mut Vec<u8>,
    columns_name: &str,
    actions: BTreeMap<Column, ColumnAction>,
) {
    let name = sibling_name(columns_name, "col");
    let tag = Tag {
        name: name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut pending: Option<(Column, Column, ColumnAction)> = None;
    for (column, action) in actions {
        if !action.materializes() {
            continue;
        }
        match pending {
            Some((first, last, previous)) if previous == action && last.next() == Some(column) => {
                pending = Some((first, column, action));
            },
            Some((first, last, previous)) => {
                write_new_column(output, &tag, first, last, previous);
                pending = Some((column, column, action));
            },
            None => pending = Some((column, column, action)),
        }
    }
    if let Some((first, last, action)) = pending {
        write_new_column(output, &tag, first, last, action);
    }
}

fn write_new_column(
    output: &mut Vec<u8>,
    tag: &Tag,
    first: Column,
    last: Column,
    action: ColumnAction,
) {
    let mut removed = Vec::new();
    let mut appended = vec![
        ("min", (first.get() + 1).to_string()),
        ("max", (last.get() + 1).to_string()),
    ];
    column_effect_attributes(action, &mut removed, &mut appended);
    write_tag(output, tag, true, &removed, &appended);
}

fn column_effect_attributes(
    action: ColumnAction,
    removed: &mut Vec<&'static str>,
    appended: &mut Vec<(&'static str, String)>,
) {
    if let Some(hidden) = action.hidden {
        removed.push("hidden");
        if hidden {
            appended.push(("hidden", "1".to_owned()));
        }
    }
    if let Some(width) = action.width {
        removed.extend(["width", "customWidth"]);
        if let WidthEffect::Set(width) = width {
            appended.push(("width", width.get().to_string()));
            appended.push(("customWidth", "1".to_owned()));
        }
    }
    if let Some(style) = action.style {
        removed.push("style");
        if let StyleEffect::Set(key) = style {
            appended.push(("style", key.to_string()));
        }
    }
    if let Some(best_fit) = action.best_fit {
        removed.push("bestFit");
        if best_fit {
            appended.push(("bestFit", "1".to_owned()));
        }
    }
    if let Some(outline) = action.outline {
        removed.push("outlineLevel");
        if outline != Outline::NONE {
            appended.push(("outlineLevel", outline.get().to_string()));
        }
    }
    if let Some(collapsed) = action.collapsed {
        removed.push("collapsed");
        if collapsed {
            appended.push(("collapsed", "1".to_owned()));
        }
    }
    if let Some(phonetic) = action.phonetic {
        removed.push("phonetic");
        if phonetic {
            appended.push(("phonetic", "1".to_owned()));
        }
    }
}
