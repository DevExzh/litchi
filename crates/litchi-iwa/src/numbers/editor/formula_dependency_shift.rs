//! Formula dependency coordinate shifts for inserted and deleted table axes.

use super::*;

mod ast;
mod wire;

use ast::rewrite_formula_asts;
use wire::{
    rewrite_shifted_dependency_tile_wire, rewrite_shifted_formula_owner_wire,
    rewrite_shifted_range_tile_wire,
};

/// Rewrite only the coordinate-bearing AST fields in an already-validated
/// formula archive while retaining unknown protobuf fields byte-for-byte.
///
/// Merge formulas are stored outside the ordinary calculation-engine formula
/// table, but use the same AST wire representation. Keeping this primitive
/// here prevents the merge implementation from duplicating the delicate
/// lossless AST rewrite logic.
pub(super) fn rewrite_formula_archive_wire(
    data: &[u8],
    previous: &tsce::FormulaArchive,
    current: &tsce::FormulaArchive,
) -> Result<Vec<u8>> {
    ast::rewrite_formula_archive_wire(data, previous, current)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum DependencyAxis {
    Column,
    Row,
}

/// A formula host retained at its current post-deletion coordinate even though
/// its source cell lies in the deleted table band.
///
/// Native merged cells carry their leading formula anchor into the next
/// physical cell before compacting the axis. After compaction, that formula
/// remains at this coordinate.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct FormulaHostCoordinate {
    row: u32,
    column: u32,
}

impl FormulaHostCoordinate {
    pub(super) fn from_table_coordinates(row: usize, column: usize) -> Result<Self> {
        Ok(Self {
            row: u32::try_from(row)
                .map_err(|_| Error::ParseError("Numbers formula row exceeds u32".to_owned()))?,
            column: u32::try_from(column)
                .map_err(|_| Error::ParseError("Numbers formula column exceeds u32".to_owned()))?,
        })
    }

    const fn coordinate(self, axis: DependencyAxis) -> u32 {
        match axis {
            DependencyAxis::Column => self.column,
            DependencyAxis::Row => self.row,
        }
    }
}

/// Validated formula hosts carried out of a deleted axis by merged-cell
/// anchor relocation.
#[derive(Debug, Default)]
struct FormulaHostRetentions {
    hosts: Vec<FormulaHostCoordinate>,
}

impl FormulaHostRetentions {
    fn new(
        mut hosts: Vec<FormulaHostCoordinate>,
        axis: DependencyAxis,
        deletion: u32,
    ) -> Result<Self> {
        if hosts.iter().any(|host| host.coordinate(axis) != deletion) {
            return Err(Error::InvalidFormat(
                "A retained iWork formula host does not lie in the deleted table axis".to_owned(),
            ));
        }
        hosts.sort_unstable();
        if hosts.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(Error::InvalidFormat(
                "An iWork formula host is retained more than once".to_owned(),
            ));
        }
        Ok(Self { hosts })
    }

    fn contains(&self, row: u32, column: u32) -> bool {
        self.hosts
            .binary_search(&FormulaHostCoordinate { row, column })
            .is_ok()
    }

    fn shifted_host(
        &self,
        row: u32,
        column: u32,
        axis: DependencyAxis,
        position: u32,
        mutation: DependencyMutation,
        what: &str,
    ) -> Result<(u32, u32)> {
        if self.contains(row, column) {
            return Ok((row, column));
        }
        match axis {
            DependencyAxis::Column => Ok((row, mutation.coordinate(column, position, what)?)),
            DependencyAxis::Row => Ok((mutation.coordinate(row, position, what)?, column)),
        }
    }

    fn is_empty(&self) -> bool {
        self.hosts.is_empty()
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum FooterRangeInsertion {
    Body,
    FixedSection,
}

impl FooterRangeInsertion {
    const fn expands_footer_ranges(self) -> bool {
        matches!(self, Self::Body)
    }
}

impl DependencyAxis {
    fn noun(self) -> &'static str {
        match self {
            Self::Column => "column",
            Self::Row => "row",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DependencyMutation {
    Insert,
    Delete,
}

#[derive(Debug, Default)]
struct FormulaDependencyAdjustments {
    local_precedents: BTreeMap<(u32, u32), LocalPrecedentAdjustments>,
}

#[derive(Debug, Default)]
struct LocalPrecedentAdjustments {
    insert: Vec<(u32, u32)>,
    remove: Vec<(u32, u32)>,
}

impl LocalPrecedentAdjustments {
    fn normalize(&mut self) {
        self.insert.sort_unstable();
        self.insert.dedup();
        self.remove.sort_unstable();
        self.remove.dedup();
    }

    fn is_empty(&self) -> bool {
        self.insert.is_empty() && self.remove.is_empty()
    }
}

impl DependencyMutation {
    fn verb(self) -> &'static str {
        match self {
            Self::Insert => "insert",
            Self::Delete => "delete",
        }
    }

    fn coordinate(self, value: u32, position: u32, what: &str) -> Result<u32> {
        match self {
            Self::Insert if value < position => Ok(value),
            Self::Insert => value
                .checked_add(1)
                .ok_or_else(|| Error::ParseError(format!("Numbers {what} overflow"))),
            Self::Delete if value < position => Ok(value),
            Self::Delete if value > position => Ok(value - 1),
            Self::Delete => Err(Error::ParseError(format!(
                "Cannot delete Numbers {} {position}: a formula still references it",
                what
            ))),
        }
    }
}

pub(super) fn shift_formula_dependencies(
    package: &mut IWorkPackage,
    table_info_id: u64,
    axis: DependencyAxis,
    insertion: u32,
    footer_range_insertion: FooterRangeInsertion,
) -> Result<()> {
    let retained_hosts = FormulaHostRetentions::default();
    mutate_formula_dependencies(
        package,
        table_info_id,
        axis,
        insertion,
        DependencyMutation::Insert,
        footer_range_insertion.expands_footer_ranges(),
        &retained_hosts,
    )
}

pub(super) fn delete_formula_dependencies(
    package: &mut IWorkPackage,
    table_info_id: u64,
    axis: DependencyAxis,
    deletion: u32,
    retained_hosts: Vec<FormulaHostCoordinate>,
) -> Result<()> {
    let retained_hosts = FormulaHostRetentions::new(retained_hosts, axis, deletion)?;
    mutate_formula_dependencies(
        package,
        table_info_id,
        axis,
        deletion,
        DependencyMutation::Delete,
        false,
        &retained_hosts,
    )
}

fn mutate_formula_dependencies(
    package: &mut IWorkPackage,
    table_info_id: u64,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    expand_footer_ranges: bool,
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    let Some(component) = package.calculation_engine_entry_name()?.map(str::to_owned) else {
        return Ok(());
    };
    let adjustments = rewrite_formula_asts(
        package,
        &component,
        table_info_id,
        axis,
        position,
        mutation,
        expand_footer_ranges,
        retained_hosts,
    )?;
    package.update_archive(&component, |archive| {
        let Some((owner_id, message_index)) = archive.objects.iter().find_map(|object| {
            object.messages.iter().enumerate().find_map(|(index, message)| {
                if message.type_ != 4008 {
                    return None;
                }
                let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                    .ok()?;
                (owner.formula_owner.as_ref().map(|reference| reference.identifier)
                    == Some(table_info_id))
                .then_some((object.archive_info.identifier?, index))
            })
        }) else {
            return Err(Error::InvalidFormat(format!(
                "Numbers table info {table_info_id} has no formula dependency owner"
            )));
        };
        let owner_object = archive.object(owner_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers formula owner {owner_id} is missing"))
        })?;
        let original = owner_object.messages[message_index].data.clone();
        let previous = tsce::FormulaOwnerDependenciesArchive::decode(original.as_slice())?;
        validate_shiftable_formula_owner(&previous, axis, mutation)?;
        let internal_owner_id = previous.internal_formula_owner_id;
        reject_incoming_dependencies(
            archive,
            owner_id,
            internal_owner_id,
            &previous.formula_owner_uid,
            axis,
            mutation,
        )?;
        let range_dependency_hosts = range_dependency_hosts(archive, &previous)?;
        let mut current = previous.clone();
        mutate_spanning_ranges(
            &mut current.spanning_column_dependencies,
            axis,
            mutation,
        )?;
        mutate_spanning_ranges(
            &mut current.spanning_row_dependencies,
            axis,
            mutation,
        )?;
        mutate_range_dependencies(
            &mut current.range_dependencies,
            axis,
            position,
            internal_owner_id,
            mutation,
            &adjustments,
            retained_hosts,
        )?;
        mutate_uuid_references(
            &mut current.uuid_references,
            axis,
            position,
            mutation,
            retained_hosts,
        )?;
        if let Some(dependencies) = &mut current.cell_dependencies {
            for record in &mut dependencies.cell_record {
                mutate_dependency_record(
                    record,
                    axis,
                    position,
                    internal_owner_id,
                    mutation,
                    &adjustments,
                    &range_dependency_hosts,
                    retained_hosts,
                )?;
            }
        }

        let tile_ids = previous
            .tiled_cell_dependencies
            .as_ref()
            .map(|dependencies| {
                dependencies
                    .cell_record_tiles
                    .iter()
                    .map(|reference| reference.identifier)
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        for tile_id in tile_ids {
            let object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers formula dependency tile {tile_id} is missing"
                ))
            })?;
            let tile_message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == 4009)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers formula dependency tile {tile_id} has no payload"
                    ))
                })?;
            let tile_original = object.messages[tile_message_index].data.clone();
            let tile_previous = tsce::CellRecordTileArchive::decode(tile_original.as_slice())?;
            if tile_previous.internal_owner_id != internal_owner_id {
                return Err(Error::InvalidFormat(format!(
                    "Numbers dependency tile {tile_id} belongs to another owner"
                )));
            }
            let mut tile_current = tile_previous.clone();
            for record in &mut tile_current.cell_records {
                mutate_dependency_record(
                    record,
                    axis,
                    position,
                    internal_owner_id,
                    mutation,
                    &adjustments,
                    &range_dependency_hosts,
                    retained_hosts,
                )?;
                let (coordinate, tile_begin, tile_size) = match axis {
                    DependencyAxis::Column => (
                        record.column,
                        tile_current.tile_column_begin,
                        FORMULA_DEPENDENCY_TILE_COLUMNS,
                    ),
                    DependencyAxis::Row => (
                        record.row,
                        tile_current.tile_row_begin,
                        FORMULA_DEPENDENCY_TILE_ROWS,
                    ),
                };
                let expected_begin = coordinate / tile_size * tile_size;
                if expected_begin != tile_begin {
                    return Err(Error::ParseError(format!(
                        "Cannot {} Numbers {}: formula dependency at {coordinate} would cross a dependency-tile boundary",
                        mutation.verb(), axis.noun()
                    )));
                }
            }
            let data = rewrite_shifted_dependency_tile_wire(
                &tile_original,
                &tile_previous,
                &tile_current,
                axis,
            )?;
            object.replace_message(
                tile_message_index,
                RawMessage {
                    type_: 4009,
                    data,
                },
            )?;
        }

        let range_tile_ids = previous
            .tiled_range_dependencies
            .as_ref()
            .map(|dependencies| {
                dependencies
                    .range_precedents_tile
                    .iter()
                    .map(|reference| reference.identifier)
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        for tile_id in range_tile_ids {
            let object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers range dependency tile {tile_id} is missing"
                ))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == 4010)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers range dependency tile {tile_id} has no payload"
                    ))
                })?;
            let original = object.messages[message_index].data.clone();
            let previous_tile =
                tsce::RangePrecedentsTileArchive::decode(original.as_slice())?;
            let mut current_tile = previous_tile.clone();
            mutate_range_tile(
                &mut current_tile,
                axis,
                position,
                internal_owner_id,
                mutation,
                &adjustments,
                retained_hosts,
            )?;
            let data = rewrite_shifted_range_tile_wire(
                &original,
                &previous_tile,
                &current_tile,
            )?;
            let message_type = object.messages[message_index].type_;
            object.replace_message(message_index, RawMessage { type_: message_type, data })?;
        }

        let data = rewrite_shifted_formula_owner_wire(&original, &previous, &current, axis)?;
        let object = archive.object_mut(owner_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers formula owner {owner_id} is missing"))
        })?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

fn range_dependency_hosts(
    archive: &Archive,
    owner: &tsce::FormulaOwnerDependenciesArchive,
) -> Result<HashSet<(u32, u32)>> {
    let mut hosts = owner
        .range_dependencies
        .as_ref()
        .into_iter()
        .flat_map(|dependencies| &dependencies.back_dependency)
        .map(|dependency| (dependency.cell_coord_row, dependency.cell_coord_column))
        .collect::<HashSet<_>>();
    for reference in owner
        .tiled_range_dependencies
        .as_ref()
        .into_iter()
        .flat_map(|dependencies| &dependencies.range_precedents_tile)
    {
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers range dependency tile {} is missing",
                reference.identifier
            ))
        })?;
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == 4010)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers range dependency tile {} has no payload",
                    reference.identifier
                ))
            })?;
        let tile = tsce::RangePrecedentsTileArchive::decode(message.data.as_slice())?;
        for dependency in &tile.from_to_range {
            hosts.insert(explicit_cell_coordinate(
                &dependency.from_coord,
                "range-tile host",
            )?);
        }
    }
    Ok(hosts)
}

fn mutate_range_dependencies(
    dependencies: &mut Option<tsce::RangeDependenciesArchive>,
    axis: DependencyAxis,
    position: u32,
    internal_owner_id: u32,
    mutation: DependencyMutation,
    adjustments: &FormulaDependencyAdjustments,
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    let Some(dependencies) = dependencies else {
        return Ok(());
    };
    for dependency in &mut dependencies.back_dependency {
        let host = (dependency.cell_coord_row, dependency.cell_coord_column);
        let local_adjustments = adjustments.local_precedents.get(&host);
        let (row, column) = retained_hosts.shifted_host(
            dependency.cell_coord_row,
            dependency.cell_coord_column,
            axis,
            position,
            mutation,
            "range dependency host",
        )?;
        dependency.cell_coord_row = row;
        dependency.cell_coord_column = column;
        if dependency.range_reference.is_some() && dependency.internal_range_reference.is_some() {
            return Err(Error::InvalidFormat(
                "iWork range dependency has both external and internal references".to_owned(),
            ));
        }
        if let Some(reference) = &mut dependency.internal_range_reference
            && reference.owner_id == internal_owner_id
        {
            mutate_range_coordinate(
                &mut reference.range,
                axis,
                position,
                mutation,
                local_adjustments,
            )?;
        }
    }
    Ok(())
}

fn mutate_range_tile(
    tile: &mut tsce::RangePrecedentsTileArchive,
    axis: DependencyAxis,
    position: u32,
    internal_owner_id: u32,
    mutation: DependencyMutation,
    adjustments: &FormulaDependencyAdjustments,
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    for dependency in &mut tile.from_to_range {
        let host = explicit_cell_coordinate(&dependency.from_coord, "range-tile host")?;
        let local_adjustments = adjustments.local_precedents.get(&host);
        let (row, column) = retained_hosts.shifted_host(
            host.0,
            host.1,
            axis,
            position,
            mutation,
            "range-tile host",
        )?;
        dependency.from_coord.row = Some(row);
        dependency.from_coord.column = Some(column);
        if tile.to_owner_id == internal_owner_id {
            mutate_cell_rect(
                &mut dependency.refers_to_rect,
                axis,
                position,
                mutation,
                local_adjustments,
            )?;
        }
    }
    Ok(())
}

fn mutate_cell_rect(
    rect: &mut tsce::CellRectArchive,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    adjustments: Option<&LocalPrecedentAdjustments>,
) -> Result<()> {
    let (origin_row, origin_column) = explicit_cell_coordinate(&rect.origin, "range origin")?;
    let row_count = rect.size.num_rows.unwrap_or(1);
    let column_count = rect.size.num_columns.unwrap_or(1);
    let mut range = tsce::RangeCoordinateArchive {
        top_left_column: origin_column,
        top_left_row: origin_row,
        bottom_right_column: origin_column
            .checked_add(column_count.saturating_sub(1))
            .ok_or_else(|| Error::ParseError("iWork range column overflow".to_owned()))?,
        bottom_right_row: origin_row
            .checked_add(row_count.saturating_sub(1))
            .ok_or_else(|| Error::ParseError("iWork range row overflow".to_owned()))?,
    };
    mutate_range_coordinate(&mut range, axis, position, mutation, adjustments)?;
    rect.origin.column = Some(range.top_left_column);
    rect.origin.row = Some(range.top_left_row);
    let columns = range
        .bottom_right_column
        .checked_sub(range.top_left_column)
        .and_then(|size| size.checked_add(1))
        .ok_or_else(|| Error::InvalidFormat("iWork range has inverted columns".to_owned()))?;
    let rows = range
        .bottom_right_row
        .checked_sub(range.top_left_row)
        .and_then(|size| size.checked_add(1))
        .ok_or_else(|| Error::InvalidFormat("iWork range has inverted rows".to_owned()))?;
    rect.size.num_columns = (columns != 1).then_some(columns);
    rect.size.num_rows = (rows != 1).then_some(rows);
    Ok(())
}

fn mutate_range_coordinate(
    range: &mut tsce::RangeCoordinateArchive,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    adjustments: Option<&LocalPrecedentAdjustments>,
) -> Result<()> {
    let footer_override = adjustments.and_then(|adjustments| {
        (axis == DependencyAxis::Row)
            .then(|| match mutation {
                DependencyMutation::Insert
                    if range.bottom_right_row.checked_add(1) == Some(position)
                        && adjustments.insert.iter().any(|&(row, column)| {
                            row == position
                                && (range.top_left_column..=range.bottom_right_column)
                                    .contains(&column)
                        }) =>
                {
                    Some(position)
                },
                DependencyMutation::Delete
                    if range.bottom_right_row == position
                        && range.top_left_row < position
                        && adjustments.remove.iter().any(|&(row, column)| {
                            row == position
                                && (range.top_left_column..=range.bottom_right_column)
                                    .contains(&column)
                        }) =>
                {
                    position.checked_sub(1)
                },
                _ => None,
            })
            .flatten()
    });
    match axis {
        DependencyAxis::Column => {
            range.top_left_column =
                mutation.coordinate(range.top_left_column, position, "range start column")?;
            range.bottom_right_column =
                mutation.coordinate(range.bottom_right_column, position, "range end column")?;
        },
        DependencyAxis::Row => {
            range.top_left_row =
                mutation.coordinate(range.top_left_row, position, "range start row")?;
            range.bottom_right_row = footer_override.map(Ok).unwrap_or_else(|| {
                mutation.coordinate(range.bottom_right_row, position, "range end row")
            })?;
        },
    }
    Ok(())
}

fn mutate_uuid_references(
    references: &mut Option<tsce::UuidReferencesArchive>,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    let Some(references) = references else {
        return Ok(());
    };
    if !retained_hosts.is_empty()
        && (!references.table_refs.is_empty() || !references.table_uuid_refs.is_empty())
    {
        return Err(Error::ParseError(
            "Cannot yet relocate a merged iWork formula anchor with UUID-reference dependencies"
                .to_owned(),
        ));
    }
    for reference in &mut references.table_refs {
        if let Some(coordinates) = &mut reference.coord_set {
            mutate_cell_coord_set(coordinates, axis, position, mutation)?;
        }
    }
    for table in &mut references.table_uuid_refs {
        for reference in &mut table.uuid_refs {
            if let Some(coordinates) = &mut reference.coord_set {
                mutate_cell_coord_set(coordinates, axis, position, mutation)?;
            }
        }
    }
    Ok(())
}

fn mutate_cell_coord_set(
    coordinates: &mut tsce::CellCoordSetArchive,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
) -> Result<()> {
    for column in &mut coordinates.column_entries {
        if axis == DependencyAxis::Column {
            column.column =
                mutation.coordinate(column.column, position, "UUID-reference host column")?;
        } else {
            for entry in &mut column.row_set.entries {
                let begin = u32::try_from(entry.range_begin).map_err(|_| {
                    Error::InvalidFormat("iWork UUID-reference row is negative".to_owned())
                })?;
                let end = entry
                    .range_end
                    .map(u32::try_from)
                    .transpose()
                    .map_err(|_| {
                        Error::InvalidFormat("iWork UUID-reference row is negative".to_owned())
                    })?;
                if end.is_some_and(|end| begin < position && position <= end) {
                    return Err(Error::ParseError(format!(
                        "Cannot {} an iWork row through a compact UUID-reference host range",
                        mutation.verb()
                    )));
                }
                entry.range_begin = i32::try_from(mutation.coordinate(
                    begin,
                    position,
                    "UUID-reference host row",
                )?)
                .map_err(|_| Error::ParseError("iWork UUID-reference row overflow".to_owned()))?;
                entry.range_end = end
                    .map(|end| {
                        mutation
                            .coordinate(end, position, "UUID-reference host row")
                            .and_then(|end| {
                                i32::try_from(end).map_err(|_| {
                                    Error::ParseError(
                                        "iWork UUID-reference row overflow".to_owned(),
                                    )
                                })
                            })
                    })
                    .transpose()?;
            }
        }
    }
    Ok(())
}

fn explicit_cell_coordinate(
    coordinate: &tsce::CellCoordinateArchive,
    what: &str,
) -> Result<(u32, u32)> {
    if coordinate.packed_data.is_some() {
        return Err(Error::ParseError(format!(
            "Cannot rewrite packed iWork {what} coordinates"
        )));
    }
    Ok((
        coordinate
            .row
            .ok_or_else(|| Error::InvalidFormat(format!("iWork {what} row is missing")))?,
        coordinate
            .column
            .ok_or_else(|| Error::InvalidFormat(format!("iWork {what} column is missing")))?,
    ))
}

fn reject_incoming_dependencies(
    archive: &Archive,
    target_object_id: u64,
    target_internal_owner_id: u32,
    target_owner_uid: &tsp::Uuid,
    axis: DependencyAxis,
    mutation: DependencyMutation,
) -> Result<()> {
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ != 4008 {
                continue;
            }
            let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            if object.archive_info.identifier == Some(target_object_id) {
                continue;
            }
            validate_shiftable_formula_owner(&owner, axis, mutation)?;
            if owner
                .cell_dependencies
                .as_ref()
                .is_some_and(|dependencies| {
                    dependencies
                        .cell_record
                        .iter()
                        .any(|record| record_has_external_owner(record, target_internal_owner_id))
                })
            {
                return Err(incoming_dependency_error(axis, mutation));
            }
            if owner
                .range_dependencies
                .as_ref()
                .is_some_and(|dependencies| {
                    dependencies.back_dependency.iter().any(|dependency| {
                        dependency
                            .internal_range_reference
                            .as_ref()
                            .is_some_and(|reference| reference.owner_id == target_internal_owner_id)
                            || dependency.range_reference.is_some()
                    })
                })
                || uuid_references_owner(&owner.uuid_references, target_owner_uid)
            {
                return Err(incoming_dependency_error(axis, mutation));
            }
            for reference in owner
                .tiled_cell_dependencies
                .as_ref()
                .into_iter()
                .flat_map(|dependencies| &dependencies.cell_record_tiles)
            {
                let tile_object = archive.object(reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers formula dependency tile {} is missing",
                        reference.identifier
                    ))
                })?;
                let tile_message = tile_object
                    .messages
                    .iter()
                    .find(|message| message.type_ == 4009)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers formula dependency tile {} has no payload",
                            reference.identifier
                        ))
                    })?;
                let tile = tsce::CellRecordTileArchive::decode(tile_message.data.as_slice())?;
                if tile
                    .cell_records
                    .iter()
                    .any(|record| record_has_external_owner(record, target_internal_owner_id))
                {
                    return Err(incoming_dependency_error(axis, mutation));
                }
            }
            for reference in owner
                .tiled_range_dependencies
                .as_ref()
                .into_iter()
                .flat_map(|dependencies| &dependencies.range_precedents_tile)
            {
                let tile_object = archive.object(reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers range dependency tile {} is missing",
                        reference.identifier
                    ))
                })?;
                let tile_message = tile_object
                    .messages
                    .iter()
                    .find(|message| message.type_ == 4010)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers range dependency tile {} has no payload",
                            reference.identifier
                        ))
                    })?;
                let tile = tsce::RangePrecedentsTileArchive::decode(tile_message.data.as_slice())?;
                if tile.to_owner_id == target_internal_owner_id && !tile.from_to_range.is_empty() {
                    return Err(incoming_dependency_error(axis, mutation));
                }
            }
        }
    }
    Ok(())
}

fn uuid_references_owner(
    references: &Option<tsce::UuidReferencesArchive>,
    owner_uid: &tsp::Uuid,
) -> bool {
    references.as_ref().is_some_and(|references| {
        references
            .table_refs
            .iter()
            .any(|reference| reference.owner_uuid == *owner_uid)
            || references
                .table_uuid_refs
                .iter()
                .any(|reference| reference.owner_uuid == *owner_uid)
    })
}

fn record_has_external_owner(
    record: &tsce::CellRecordExpandedArchive,
    target_internal_owner_id: u32,
) -> bool {
    record.expanded_edges.as_ref().is_some_and(|edges| {
        edges
            .internal_owner_id_for_edge
            .contains(&target_internal_owner_id)
    })
}

fn incoming_dependency_error(axis: DependencyAxis, mutation: DependencyMutation) -> Error {
    Error::ParseError(format!(
        "Cannot yet {} a Numbers {} referenced by formulas in another table",
        mutation.verb(),
        axis.noun()
    ))
}

fn validate_shiftable_formula_owner(
    owner: &tsce::FormulaOwnerDependenciesArchive,
    axis: DependencyAxis,
    mutation: DependencyMutation,
) -> Result<()> {
    let volatile_dependencies = owner
        .volatile_dependencies
        .as_ref()
        .is_some_and(volatile_dependencies_are_populated);
    let spanning_dependencies = owner
        .spanning_column_dependencies
        .as_ref()
        .is_some_and(|dependencies| !dependencies.coord_refers_to_spans.is_empty())
        || owner
            .spanning_row_dependencies
            .as_ref()
            .is_some_and(|dependencies| !dependencies.coord_refers_to_spans.is_empty());
    let whole_owner_dependencies = owner
        .whole_owner_dependencies
        .as_ref()
        .and_then(|dependencies| dependencies.dependent_cells.as_ref())
        .is_some_and(|cells| !cells.owner_entries.is_empty());
    let cell_errors = owner
        .cell_errors
        .as_ref()
        .is_some_and(|errors| !errors.errors.is_empty() || !errors.enhanced_errors.is_empty());
    let spills = owner
        .spill_range_sizes
        .as_ref()
        .is_some_and(|spills| !spills.spills.is_empty());
    if volatile_dependencies
        || spanning_dependencies
        || whole_owner_dependencies
        || cell_errors
        || spills
    {
        return Err(Error::ParseError(format!(
            "Cannot yet {} a {} in a Numbers table with volatile, spanning, error, or spill dependency state",
            mutation.verb(),
            axis.noun()
        )));
    }
    Ok(())
}

fn volatile_dependencies_are_populated(
    dependencies: &tsce::VolatileDependenciesExpandedArchive,
) -> bool {
    let populated_cells = [
        dependencies.volatile_time_cells.as_ref(),
        dependencies.volatile_random_cells.as_ref(),
        dependencies.volatile_locale_cells.as_ref(),
        dependencies.volatile_sheet_table_name_cells.as_ref(),
        dependencies.volatile_remote_data_cells.as_ref(),
    ]
    .into_iter()
    .flatten()
    .any(|cells| !cells.column_entries.is_empty());
    let populated_refs = dependencies
        .volatile_geometry_cell_refs
        .as_ref()
        .is_some_and(|cells| !cells.owner_entries.is_empty());
    populated_cells || populated_refs
}

fn mutate_spanning_ranges(
    dependencies: &mut Option<tsce::SpanningDependenciesExpandedArchive>,
    axis: DependencyAxis,
    mutation: DependencyMutation,
) -> Result<()> {
    let Some(dependencies) = dependencies else {
        return Ok(());
    };
    for range in [
        dependencies.total_range_for_table.as_mut(),
        dependencies.body_range_for_table.as_mut(),
    ]
    .into_iter()
    .flatten()
    {
        match axis {
            DependencyAxis::Column => {
                range.bottom_right_column =
                    mutate_range_end(range.bottom_right_column, mutation, "column")?;
            },
            DependencyAxis::Row => {
                range.bottom_right_row = mutate_range_end(range.bottom_right_row, mutation, "row")?;
            },
        }
    }
    Ok(())
}

fn mutate_range_end(value: u32, mutation: DependencyMutation, axis: &str) -> Result<u32> {
    match mutation {
        DependencyMutation::Insert => value.checked_add(1).ok_or_else(|| {
            Error::ParseError(format!("Numbers formula dependency range {axis} overflow"))
        }),
        DependencyMutation::Delete => value.checked_sub(1).ok_or_else(|| {
            Error::ParseError(format!("Numbers formula dependency range {axis} underflow"))
        }),
    }
}

fn mutate_dependency_record(
    record: &mut tsce::CellRecordExpandedArchive,
    axis: DependencyAxis,
    position: u32,
    internal_owner_id: u32,
    mutation: DependencyMutation,
    adjustments: &FormulaDependencyAdjustments,
    range_dependency_hosts: &HashSet<(u32, u32)>,
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    let previous_host = (record.row, record.column);
    let (row, column) = retained_hosts.shifted_host(
        record.row,
        record.column,
        axis,
        position,
        mutation,
        "formula host",
    )?;
    record.row = row;
    record.column = column;
    let Some(edges) = &mut record.expanded_edges else {
        return Ok(());
    };
    if edges.edge_without_owner_rows.len() != edges.edge_without_owner_columns.len()
        || edges.edge_with_owner_rows.len() != edges.edge_with_owner_columns.len()
        || edges.edge_with_owner_rows.len() != edges.internal_owner_id_for_edge.len()
    {
        return Err(Error::InvalidFormat(
            "Numbers formula dependency edge arrays have inconsistent lengths".to_owned(),
        ));
    }
    if let Some(adjustment) = adjustments
        .local_precedents
        .get(&previous_host)
        .filter(|_| !range_dependency_hosts.contains(&previous_host))
    {
        if adjustment.remove.iter().any(|coordinate| {
            !edges
                .edge_without_owner_rows
                .iter()
                .copied()
                .zip(edges.edge_without_owner_columns.iter().copied())
                .any(|existing| existing == *coordinate)
        }) {
            return Err(Error::InvalidFormat(format!(
                "iWork footer formula at ({}, {}) does not contain every dependency selected for contraction",
                previous_host.0, previous_host.1
            )));
        }
        let mut retained = edges
            .edge_without_owner_rows
            .iter()
            .copied()
            .zip(edges.edge_without_owner_columns.iter().copied())
            .filter(|coordinate| adjustment.remove.binary_search(coordinate).is_err())
            .collect::<Vec<_>>();
        for coordinate in &mut retained {
            match axis {
                DependencyAxis::Column => {
                    coordinate.1 = mutation.coordinate(
                        coordinate.1,
                        position,
                        "local formula precedent column",
                    )?;
                },
                DependencyAxis::Row => {
                    coordinate.0 = mutation.coordinate(
                        coordinate.0,
                        position,
                        "local formula precedent row",
                    )?;
                },
            }
        }
        retained.extend(adjustment.insert.iter().copied());
        retained.sort_unstable();
        retained.dedup();
        edges.edge_without_owner_rows = retained.iter().map(|(row, _)| *row).collect();
        edges.edge_without_owner_columns = retained.iter().map(|(_, column)| *column).collect();
    } else {
        let local_coordinates = match axis {
            DependencyAxis::Column => &mut edges.edge_without_owner_columns,
            DependencyAxis::Row => &mut edges.edge_without_owner_rows,
        };
        for coordinate in local_coordinates {
            *coordinate = mutation.coordinate(
                *coordinate,
                position,
                match axis {
                    DependencyAxis::Column => "local formula precedent column",
                    DependencyAxis::Row => "local formula precedent row",
                },
            )?;
        }
    }
    let owned_coordinates = match axis {
        DependencyAxis::Column => &mut edges.edge_with_owner_columns,
        DependencyAxis::Row => &mut edges.edge_with_owner_rows,
    };
    for (coordinate, owner) in owned_coordinates
        .iter_mut()
        .zip(&edges.internal_owner_id_for_edge)
    {
        if *owner == internal_owner_id {
            *coordinate = mutation.coordinate(
                *coordinate,
                position,
                match axis {
                    DependencyAxis::Column => "formula precedent column",
                    DependencyAxis::Row => "formula precedent row",
                },
            )?;
        }
    }
    Ok(())
}
