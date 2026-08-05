//! Planning and validation for physical Numbers row sorting.

use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::numbers::bnc::{BncCell, CachedScalar, StoredValue};
use crate::numbers::editor::stroke_layers::{
    reorder_rows as reorder_stroke_rows, validate_row_reorder as validate_stroke_row_reorder,
};
use crate::numbers::editor::table_topology::{
    category_grouping_is_enabled, deprecated_category_grouping_is_enabled, filter_has_row_state,
    table_has_spill_state,
};
use crate::table_hidden_axes::{positional_user_hidden_axes, restore_positional_user_hidden_axes};
use litchi_iwa_common::table::axis::HiddenAxes;

use super::*;

mod storage;

use storage::{reorder_body_row_headers, reorder_body_table_tile_rows, reorder_row_uids};

#[derive(Debug)]
pub(super) struct BodySortPlan {
    body_start: usize,
    sources_by_destination: Vec<usize>,
}

impl BodySortPlan {
    fn reorders_rows(&self) -> bool {
        self.sources_by_destination
            .iter()
            .enumerate()
            .any(|(destination_offset, source)| *source != self.body_start + destination_offset)
    }

    fn destinations_by_source(&self) -> Result<Vec<usize>> {
        let mut destinations = vec![None; self.sources_by_destination.len()];
        for (destination_offset, source) in self.sources_by_destination.iter().copied().enumerate()
        {
            let source_offset = source
                .checked_sub(self.body_start)
                .filter(|offset| *offset < self.sources_by_destination.len());
            let Some(source_offset) = source_offset else {
                return Err(Error::InvalidFormat(
                    "Numbers sort plan contains a source row outside the body".to_owned(),
                ));
            };
            if destinations[source_offset]
                .replace(
                    self.body_start
                        .checked_add(destination_offset)
                        .ok_or_else(|| {
                            Error::ParseError("Numbers sort destination overflow".to_owned())
                        })?,
                )
                .is_some()
            {
                return Err(Error::InvalidFormat(
                    "Numbers sort plan contains a source row more than once".to_owned(),
                ));
            }
        }
        destinations
            .into_iter()
            .map(|destination| {
                destination.ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers sort plan does not cover every body row".to_owned(),
                    )
                })
            })
            .collect()
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum SortScalar {
    Number(f64),
    Text(u32),
    Boolean(bool),
    Date(f64),
    Duration(f64),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum SortScalarKind {
    Number,
    Text,
    Boolean,
    Date,
    Duration,
}

impl SortScalar {
    fn kind(self) -> SortScalarKind {
        match self {
            Self::Number(_) => SortScalarKind::Number,
            Self::Text(_) => SortScalarKind::Text,
            Self::Boolean(_) => SortScalarKind::Boolean,
            Self::Date(_) => SortScalarKind::Date,
            Self::Duration(_) => SortScalarKind::Duration,
        }
    }

    fn compare(self, other: Self, text_by_identifier: &HashMap<u32, String>) -> Ordering {
        let kinds = self.kind().cmp(&other.kind());
        if kinds != Ordering::Equal {
            return kinds;
        }
        match (self, other) {
            (Self::Number(left), Self::Number(right))
            | (Self::Date(left), Self::Date(right))
            | (Self::Duration(left), Self::Duration(right)) => compare_finite_numbers(left, right),
            (Self::Text(left), Self::Text(right)) => match (
                text_by_identifier.get(&left),
                text_by_identifier.get(&right),
            ) {
                (Some(left), Some(right)) => left.cmp(right),
                _ => left.cmp(&right),
            },
            (Self::Boolean(left), Self::Boolean(right)) => left.cmp(&right),
            _ => Ordering::Equal,
        }
    }
}

fn compare_finite_numbers(left: f64, right: f64) -> Ordering {
    if left < right {
        Ordering::Less
    } else if left > right {
        Ordering::Greater
    } else {
        Ordering::Equal
    }
}

pub(super) fn apply_attached_table_sort_order(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &Order,
) -> Result<bool> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_sort_order(&descriptor.model, order)?;
    if order.scope() != Scope::EntireTable {
        return Err(Error::ParseError(
            "Cannot execute a selected-row Numbers sort without an explicit row range".to_owned(),
        ));
    }
    let (body_start, body_end) = table_body_bounds(&descriptor.model)?;
    apply_attached_table_sort_range(package, table_id, &descriptor, order, body_start, body_end)
}

pub(super) fn apply_attached_table_sort_order_to_rows(
    package: &mut IWorkPackage,
    table_id: u64,
    order: &Order,
    rows: RowRange,
) -> Result<bool> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_sort_order(&descriptor.model, order)?;
    if order.scope() != Scope::SelectedRows {
        return Err(Error::ParseError(
            "Cannot execute an entire-table Numbers sort through a selected-row range".to_owned(),
        ));
    }
    let (body_start, body_end) = table_body_bounds(&descriptor.model)?;
    let body_rows = body_end - body_start;
    if rows.end() > body_rows {
        return Err(Error::ParseError(format!(
            "Numbers selected-row sort range {}..{} is outside the table's {body_rows} body rows",
            rows.start(),
            rows.end()
        )));
    }
    let selected_start = body_start
        .checked_add(rows.start())
        .ok_or_else(|| Error::ParseError("Numbers selected-row sort start overflow".to_owned()))?;
    let selected_end = body_start
        .checked_add(rows.end())
        .ok_or_else(|| Error::ParseError("Numbers selected-row sort end overflow".to_owned()))?;
    apply_attached_table_sort_range(
        package,
        table_id,
        &descriptor,
        order,
        selected_start,
        selected_end,
    )
}

fn apply_attached_table_sort_range(
    package: &mut IWorkPackage,
    table_id: u64,
    descriptor: &TableDescriptor,
    order: &Order,
    row_start: usize,
    row_end: usize,
) -> Result<bool> {
    if row_end.saturating_sub(row_start) < 2 {
        return Ok(false);
    }

    let locations = object_locations(package)?;
    let positional_hidden_axes = validate_sort_features(package, &locations, descriptor)?;
    let plan = plan_body_sort(
        package,
        &locations,
        &descriptor.model,
        row_start,
        row_end,
        order,
    )?;
    if !plan.reorders_rows() {
        return Ok(false);
    }
    let destinations_by_source = plan.destinations_by_source()?;
    if let Some(sidecar) = &descriptor.model.stroke_sidecar {
        validate_stroke_row_reorder(
            package,
            &locations,
            sidecar.identifier,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns,
            row_start,
            &destinations_by_source,
        )?;
    }
    reorder_body_table_tile_rows(
        package,
        &locations,
        &descriptor.model,
        row_start,
        &destinations_by_source,
    )?;
    reorder_body_row_headers(
        package,
        &locations,
        &descriptor.model,
        row_start,
        &destinations_by_source,
    )?;
    if let Some(sidecar) = &descriptor.model.stroke_sidecar {
        reorder_stroke_rows(
            package,
            &locations,
            sidecar.identifier,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns,
            row_start,
            &destinations_by_source,
        )?;
    }
    let row_uids = descriptor
        .model
        .base_column_row_uids
        .as_ref()
        .ok_or_else(|| {
            Error::InvalidFormat(
                "Cannot safely execute a Numbers sort without a stable row UID map".to_owned(),
            )
        })?;
    reorder_row_uids(
        package,
        &locations,
        row_uids.identifier,
        descriptor.model.number_of_rows as usize,
        row_start,
        &destinations_by_source,
    )?;
    if !positional_hidden_axes.is_empty() {
        restore_positional_user_hidden_axes(package, table_id, &positional_hidden_axes)?;
    }
    Ok(true)
}

fn table_body_bounds(model: &TableModelArchive) -> Result<(usize, usize)> {
    let rows = model.number_of_rows as usize;
    let settings = NumbersTableHeaderSettings::from_model(model)?;
    let body_start = settings.header_row_count();
    let body_end = rows
        .checked_sub(settings.footer_row_count())
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers footer rows exceed the table row count".to_owned())
        })?;
    if body_start > body_end {
        return Err(Error::InvalidFormat(
            "Numbers header and footer rows exceed the table row count".to_owned(),
        ));
    }
    Ok((body_start, body_end))
}

fn validate_sort_features(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    descriptor: &TableDescriptor,
) -> Result<HiddenAxes> {
    let model = &descriptor.model;
    if model.number_of_filtered_rows.unwrap_or(0) != 0
        || model.pivot_owner.is_some()
        || deprecated_category_grouping_is_enabled(model.category_owner_deprecated.as_ref())
        || filter_has_row_state(package, locations, model.row_filter_set_pre_pivot.as_ref())?
        || category_grouping_is_enabled(package, locations, model.category_owner.as_ref())?
        || table_has_spill_state(package, descriptor.table_info_id)?
    {
        return Err(Error::ParseError(
            "Cannot yet execute a Numbers sort on a filtered, grouped, pivot, or spill table"
                .to_owned(),
        ));
    }
    if let Some(table) = &model.base_data_store.conditionalstyletable
        && table_data_list_has_entries(
            package,
            locations,
            table.identifier,
            tst::table_data_list::ListType::ConditionalStyle,
        )?
    {
        return Err(Error::ParseError(
            "Cannot yet execute a Numbers sort on a table with conditional styles".to_owned(),
        ));
    }
    if model.base_column_row_uids.is_none() {
        return Err(Error::ParseError(
            "Cannot safely execute a Numbers sort without a stable row UID map".to_owned(),
        ));
    }
    if !cell_merge::regions_in_package(package, descriptor.object_id)?.is_empty() {
        return Err(Error::ParseError(
            "Cannot yet execute a Numbers sort on a table with merged cells".to_owned(),
        ));
    }
    if let Some(sidecar) = &model.stroke_sidecar {
        validate_stroke_row_reorder(
            package,
            locations,
            sidecar.identifier,
            model.number_of_rows,
            model.number_of_columns,
            0,
            &[],
        )?;
    }
    let declares_hidden_storage = model.hidden_states_owner.is_some()
        || model.hidden_state_formula_owner_for_rows.is_some()
        || model.hidden_state_formula_owner_for_columns.is_some()
        || model.number_of_hidden_rows.unwrap_or_default() != 0
        || model.number_of_user_hidden_rows.unwrap_or_default() != 0
        || model.number_of_hidden_columns.unwrap_or_default() != 0
        || model.number_of_user_hidden_columns.unwrap_or_default() != 0;
    if declares_hidden_storage {
        positional_user_hidden_axes(package, descriptor.object_id)
    } else {
        Ok(HiddenAxes::empty())
    }
}

fn plan_body_sort(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    body_start: usize,
    body_end: usize,
    order: &Order,
) -> Result<BodySortPlan> {
    let rows = model.number_of_rows as usize;
    let columns = model.number_of_columns as usize;
    let tile_size = model.base_data_store.tiles.tile_size.unwrap_or(256);
    if tile_size == 0 {
        return Err(Error::InvalidFormat(
            "Numbers table declares a zero tile size".to_owned(),
        ));
    }
    let mut rule_positions = HashMap::with_capacity(order.rules().len());
    for (position, rule) in order.rules().iter().enumerate() {
        rule_positions.insert(rule.column().get(), position);
    }
    let mut keys_by_body_row = vec![None; body_end - body_start];
    let mut tile_keys = HashSet::with_capacity(model.base_data_store.tiles.tiles.len());

    for reference in &model.base_data_store.tiles.tiles {
        if !tile_keys.insert(reference.tileid) {
            return Err(Error::InvalidFormat(format!(
                "Numbers table repeats tile key {}",
                reference.tileid
            )));
        }
        let archive_name = locations.get(&reference.tile.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers tile object {} is missing",
                reference.tile.identifier
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(reference.tile.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers tile object {} is missing",
                reference.tile.identifier
            ))
        })?;
        let tile = decode_unique_tile(object, reference.tile.identifier)?;
        let mut row_keys = HashSet::with_capacity(tile.row_infos.len());
        for row in &tile.row_infos {
            if row.tile_row_index >= tile_size || !row_keys.insert(row.tile_row_index) {
                return Err(Error::InvalidFormat(
                    "Numbers tile contains an invalid row payload".to_owned(),
                ));
            }
            let global_row = reference
                .tileid
                .checked_mul(tile_size)
                .and_then(|base| base.checked_add(row.tile_row_index))
                .ok_or_else(|| Error::ParseError("Numbers tile row overflow".to_owned()))?;
            let global_row = global_row as usize;
            if global_row >= rows {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table stores row {global_row} outside its {rows} rows"
                )));
            }
            if !(body_start..body_end).contains(&global_row) {
                continue;
            }
            let body_row = global_row - body_start;
            if keys_by_body_row[body_row].is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table stores body row {global_row} more than once"
                )));
            }
            let mut sort_values = vec![None; order.rules().len()];
            for (column, cell) in split_row(row)?.iter().enumerate() {
                let Some(cell) = cell.as_deref() else {
                    continue;
                };
                if column >= columns {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers body row {global_row} stores a cell outside the table columns"
                    )));
                }
                let cell = BncCell::parse(cell)?;
                validate_movable_body_cell(&cell, global_row, column)?;
                if let Some(&rule_position) = rule_positions.get(&column) {
                    sort_values[rule_position] = Some(sort_scalar(&cell, global_row, column)?);
                }
            }
            let sort_values = sort_values
                .into_iter()
                .enumerate()
                .map(|(rule_position, value)| {
                    value.ok_or_else(|| {
                        let column = order.rules()[rule_position].column().get();
                        Error::ParseError(format!(
                            "Numbers body row {global_row} has no scalar sort key in column {column}"
                        ))
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            keys_by_body_row[body_row] = Some(sort_values);
        }
    }

    let keys_by_body_row = keys_by_body_row
        .into_iter()
        .enumerate()
        .map(|(body_row, keys)| {
            keys.ok_or_else(|| {
                Error::ParseError(format!(
                    "Numbers body row {} has no stored sort-key cells",
                    body_start + body_row
                ))
            })
        })
        .collect::<Result<Vec<_>>>()?;
    validate_consistent_scalar_kinds(&keys_by_body_row, order)?;
    let text_identifiers = keys_by_body_row
        .iter()
        .flatten()
        .filter_map(|scalar| match scalar {
            SortScalar::Text(identifier) => Some(*identifier),
            _ => None,
        })
        .collect::<HashSet<_>>();
    let text_by_identifier = resolve_table_string_values(
        package,
        locations,
        model.base_data_store.string_table.identifier,
        &text_identifiers,
    )?;
    validate_sort_text_references(&keys_by_body_row, body_start, order, &text_by_identifier)?;

    let mut source_offsets = (0..keys_by_body_row.len()).collect::<Vec<_>>();
    source_offsets.sort_by(|left, right| {
        compare_body_rows(
            &keys_by_body_row[*left],
            &keys_by_body_row[*right],
            order,
            &text_by_identifier,
        )
        .then_with(|| left.cmp(right))
    });
    Ok(BodySortPlan {
        body_start,
        sources_by_destination: source_offsets
            .into_iter()
            .map(|source_offset| body_start + source_offset)
            .collect(),
    })
}

fn decode_unique_tile(object: &ArchiveObject, tile_id: u64) -> Result<Tile> {
    let tiles = object
        .messages
        .iter()
        .filter_map(|message| Tile::decode(message.data.as_slice()).ok())
        .collect::<Vec<_>>();
    match tiles.as_slice() {
        [tile] => Ok(tile.clone()),
        [] => Err(Error::InvalidFormat(format!(
            "Object {tile_id} has no Numbers tile payload"
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "Object {tile_id} has multiple Numbers tile payloads"
        ))),
    }
}

fn validate_movable_body_cell(cell: &BncCell, row: usize, column: usize) -> Result<()> {
    if cell.formula_error_identifier().is_some() {
        return Err(Error::ParseError(format!(
            "Cannot yet execute a Numbers sort with a formula error in body cell ({row}, {column})"
        )));
    }
    match cell.stored_value() {
        StoredValue::Formula(_) => Err(Error::ParseError(format!(
            "Cannot yet execute a Numbers sort with a formula in body cell ({row}, {column})"
        ))),
        StoredValue::Error => Err(Error::ParseError(format!(
            "Cannot yet execute a Numbers sort with an error in body cell ({row}, {column})"
        ))),
        StoredValue::Unsupported(kind) => Err(Error::ParseError(format!(
            "Cannot yet execute a Numbers sort with unsupported cell type {kind} in body cell ({row}, {column})"
        ))),
        StoredValue::Empty
        | StoredValue::Number
        | StoredValue::Text(_)
        | StoredValue::RichText(_)
        | StoredValue::Date
        | StoredValue::Boolean
        | StoredValue::Duration => Ok(()),
    }
}

fn sort_scalar(cell: &BncCell, row: usize, column: usize) -> Result<SortScalar> {
    if let StoredValue::Text(identifier) = cell.stored_value() {
        return Ok(SortScalar::Text(identifier));
    }
    let scalar = match cell.cached_scalar()? {
        Some(CachedScalar::Number(value)) if value.is_finite() => SortScalar::Number(value),
        Some(CachedScalar::Boolean(value)) => SortScalar::Boolean(value),
        Some(CachedScalar::Date(value)) if value.is_finite() => SortScalar::Date(value),
        Some(CachedScalar::Duration(value)) if value.is_finite() => SortScalar::Duration(value),
        Some(CachedScalar::Number(_))
        | Some(CachedScalar::Date(_))
        | Some(CachedScalar::Duration(_)) => {
            return Err(Error::ParseError(format!(
                "Numbers sort key in body cell ({row}, {column}) is not finite"
            )));
        },
        Some(CachedScalar::Unsupported(kind)) => {
            return Err(Error::ParseError(format!(
                "Numbers sort key in body cell ({row}, {column}) has unsupported cell type {kind}"
            )));
        },
        None => {
            return Err(Error::ParseError(format!(
                "Numbers sort key in body cell ({row}, {column}) is empty"
            )));
        },
    };
    Ok(scalar)
}

fn validate_consistent_scalar_kinds(
    keys_by_body_row: &[Vec<SortScalar>],
    order: &Order,
) -> Result<()> {
    let Some(first) = keys_by_body_row.first() else {
        return Ok(());
    };
    for (rule_position, rule) in order.rules().iter().enumerate() {
        let kind = first[rule_position].kind();
        if keys_by_body_row
            .iter()
            .any(|keys| keys[rule_position].kind() != kind)
        {
            return Err(Error::ParseError(format!(
                "Numbers sort column {} mixes scalar types; use one Number, plain Text, Boolean, Date, or Duration type per rule",
                rule.column().get()
            )));
        }
    }
    Ok(())
}

fn validate_sort_text_references(
    keys_by_body_row: &[Vec<SortScalar>],
    body_start: usize,
    order: &Order,
    text_by_identifier: &HashMap<u32, String>,
) -> Result<()> {
    for (body_row, keys) in keys_by_body_row.iter().enumerate() {
        for (rule_position, scalar) in keys.iter().enumerate() {
            let SortScalar::Text(identifier) = scalar else {
                continue;
            };
            if text_by_identifier.contains_key(identifier) {
                continue;
            }
            let row = body_start
                .checked_add(body_row)
                .ok_or_else(|| Error::ParseError("Numbers sort body row overflow".to_owned()))?;
            let column = order.rules()[rule_position].column().get();
            return Err(Error::InvalidFormat(format!(
                "Numbers sort key in body cell ({row}, {column}) references missing string {identifier}"
            )));
        }
    }
    Ok(())
}

fn compare_body_rows(
    left: &[SortScalar],
    right: &[SortScalar],
    order: &Order,
    text_by_identifier: &HashMap<u32, String>,
) -> Ordering {
    for ((left, right), rule) in left.iter().zip(right).zip(order.rules()) {
        let ordering = left.compare(*right, text_by_identifier);
        let ordering = match rule.direction() {
            Direction::Ascending => ordering,
            Direction::Descending => ordering.reverse(),
        };
        if ordering != Ordering::Equal {
            return ordering;
        }
    }
    Ordering::Equal
}
