//! Formula dependency coordinate shifts for inserted and deleted table axes.

use super::*;
use std::collections::BTreeSet;

mod ast;
mod wire;

use ast::rewrite_formula_asts;
use wire::{
    rewrite_shifted_dependency_tile_wire, rewrite_shifted_formula_owner_wire,
    rewrite_shifted_range_tile_wire,
};

const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const CELL_DEPENDENCY_TILE_MESSAGE_TYPE: u32 = 4_009;
const RANGE_DEPENDENCY_TILE_MESSAGE_TYPE: u32 = 4_010;

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

    /// Return the adjustment to a relative cross-table reference held by a
    /// formula host as its table axis changes.
    ///
    /// An ordinary formula moves with its host, so its relative cross-table
    /// coordinates compensate for that host movement and continue to point at
    /// the same external cell. A merged anchor retained through deletion is
    /// different: native iWork first relocates it one cell forward, then
    /// compacts the deleted band. Its relative external coordinates therefore
    /// advance by one even though the final host coordinate is unchanged.
    fn cross_table_relative_offset(
        &self,
        row: u32,
        column: u32,
        axis: DependencyAxis,
        position: u32,
        mutation: DependencyMutation,
        what: &str,
    ) -> Result<i64> {
        if self.contains(row, column) {
            return match mutation {
                DependencyMutation::Delete => Ok(1),
                DependencyMutation::Insert => Err(Error::InvalidFormat(
                    "An iWork formula host cannot be retained during axis insertion".to_owned(),
                )),
            };
        }
        let (shifted_row, shifted_column) =
            self.shifted_host(row, column, axis, position, mutation, what)?;
        let (host, shifted_host) = match axis {
            DependencyAxis::Column => (column, shifted_column),
            DependencyAxis::Row => (row, shifted_row),
        };
        Ok(i64::from(host) - i64::from(shifted_host))
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
    external_precedents: BTreeMap<(u32, u32), ExternalPrecedentAdjustments>,
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

/// One explicit cross-table calculation-engine precedent.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct ExternalCellPrecedent {
    owner_id: u32,
    row: u32,
    column: u32,
}

/// One normalized rectangular cross-table calculation-engine precedent.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ExternalRangePrecedent {
    owner_id: u32,
    top: u32,
    left: u32,
    bottom: u32,
    right: u32,
}

impl ExternalRangePrecedent {
    fn from_range_coordinate(owner_id: u32, range: &tsce::RangeCoordinateArchive) -> Result<Self> {
        let range = Self {
            owner_id,
            top: range.top_left_row,
            left: range.top_left_column,
            bottom: range.bottom_right_row,
            right: range.bottom_right_column,
        };
        range.validate()
    }

    fn from_cell_rect(owner_id: u32, rect: &tsce::CellRectArchive) -> Result<Self> {
        let (top, left) = explicit_cell_coordinate(&rect.origin, "cross-table range origin")?;
        let rows = rect.size.num_rows.unwrap_or(1);
        let columns = rect.size.num_columns.unwrap_or(1);
        if rows == 0 || columns == 0 {
            return Err(Error::InvalidFormat(
                "iWork cross-table range has an empty rectangle".to_owned(),
            ));
        }
        let range = Self {
            owner_id,
            top,
            left,
            bottom: top.checked_add(rows - 1).ok_or_else(|| {
                Error::ParseError("iWork cross-table range row overflow".to_owned())
            })?,
            right: left.checked_add(columns - 1).ok_or_else(|| {
                Error::ParseError("iWork cross-table range column overflow".to_owned())
            })?,
        };
        range.validate()
    }

    fn validate(self) -> Result<Self> {
        if self.top > self.bottom || self.left > self.right {
            return Err(Error::InvalidFormat(
                "iWork cross-table range is inverted".to_owned(),
            ));
        }
        Ok(self)
    }

    fn write_range_coordinate(self, range: &mut tsce::RangeCoordinateArchive) {
        range.top_left_column = self.left;
        range.top_left_row = self.top;
        range.bottom_right_column = self.right;
        range.bottom_right_row = self.bottom;
    }

    fn write_cell_rect(self, rect: &mut tsce::CellRectArchive) -> Result<()> {
        let columns = self
            .right
            .checked_sub(self.left)
            .and_then(|width| width.checked_add(1))
            .ok_or_else(|| {
                Error::InvalidFormat("iWork cross-table range is inverted".to_owned())
            })?;
        let rows = self
            .bottom
            .checked_sub(self.top)
            .and_then(|height| height.checked_add(1))
            .ok_or_else(|| {
                Error::InvalidFormat("iWork cross-table range is inverted".to_owned())
            })?;
        rect.origin.column = Some(self.left);
        rect.origin.row = Some(self.top);
        rect.size.num_columns = (columns != 1).then_some(columns);
        rect.size.num_rows = (rows != 1).then_some(rows);
        Ok(())
    }

    fn insert_cells(self, cells: &mut BTreeSet<ExternalCellPrecedent>) {
        for row in self.top..=self.bottom {
            for column in self.left..=self.right {
                cells.insert(ExternalCellPrecedent {
                    owner_id: self.owner_id,
                    row,
                    column,
                });
            }
        }
    }
}

/// Cross-table precedents resolved from one formula AST.
#[derive(Debug, Default)]
struct CrossTablePrecedents {
    direct: BTreeSet<ExternalCellPrecedent>,
    ranges: Vec<ExternalRangePrecedent>,
}

impl CrossTablePrecedents {
    fn expanded(&self) -> BTreeSet<ExternalCellPrecedent> {
        let mut cells = self.direct.clone();
        for range in &self.ranges {
            range.insert_cells(&mut cells);
        }
        cells
    }
}

/// A rebased external range and its exact source representation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ExternalRangeAdjustment {
    previous: ExternalRangePrecedent,
    current: ExternalRangePrecedent,
}

/// Exact cross-table precedents before and after an AST rewrite.
///
/// Native Numbers packages represent cross-table ranges as dedicated range
/// records, while new documents may store the same cells as expanded edges.
/// Retained merge anchors need both representations to follow the rewritten
/// AST atomically.
#[derive(Debug)]
struct ExternalPrecedentAdjustments {
    direct_previous: BTreeSet<ExternalCellPrecedent>,
    direct_current: BTreeSet<ExternalCellPrecedent>,
    ranges: Vec<ExternalRangeAdjustment>,
}

impl ExternalPrecedentAdjustments {
    fn new(previous: CrossTablePrecedents, current: CrossTablePrecedents) -> Result<Self> {
        let CrossTablePrecedents {
            direct: direct_previous,
            ranges: previous_ranges,
        } = previous;
        let CrossTablePrecedents {
            direct: direct_current,
            ranges: current_ranges,
        } = current;
        if previous_ranges.len() != current_ranges.len() {
            return Err(Error::InvalidFormat(
                "iWork cross-table formula range count changed during coordinate mutation"
                    .to_owned(),
            ));
        }
        let mut ranges = Vec::with_capacity(previous_ranges.len());
        for (previous, current) in previous_ranges.into_iter().zip(current_ranges) {
            if previous.owner_id != current.owner_id {
                return Err(Error::InvalidFormat(
                    "iWork cross-table formula range owner changed during coordinate mutation"
                        .to_owned(),
                ));
            }
            ranges.push(ExternalRangeAdjustment { previous, current });
        }
        Ok(Self {
            direct_previous,
            direct_current,
            ranges,
        })
    }

    fn changed(&self) -> bool {
        self.direct_previous != self.direct_current
            || self
                .ranges
                .iter()
                .any(|adjustment| adjustment.previous != adjustment.current)
    }

    fn has_direct_rewrite(&self) -> bool {
        self.direct_previous != self.direct_current
    }

    fn replacement_for_range(
        &self,
        previous: ExternalRangePrecedent,
    ) -> Result<Option<ExternalRangePrecedent>> {
        let mut matches = self
            .ranges
            .iter()
            .filter(|adjustment| adjustment.previous == previous)
            .map(|adjustment| adjustment.current);
        let Some(current) = matches.next() else {
            return Ok(None);
        };
        if matches.any(|candidate| candidate != current) {
            return Err(Error::InvalidFormat(
                "iWork cross-table range dependency has ambiguous AST replacements".to_owned(),
            ));
        }
        Ok(Some(current))
    }

    fn rewrite_expanded_edges(
        &self,
        edges: &mut tsce::ExpandedEdgesArchive,
        previous_host: (u32, u32),
    ) -> Result<()> {
        let existing = edges
            .internal_owner_id_for_edge
            .iter()
            .copied()
            .zip(edges.edge_with_owner_rows.iter().copied())
            .zip(edges.edge_with_owner_columns.iter().copied())
            .map(|((owner_id, row), column)| ExternalCellPrecedent {
                owner_id,
                row,
                column,
            })
            .collect::<BTreeSet<_>>();
        if existing == self.direct_previous {
            return Self::write_external_edges(edges, &self.direct_current, previous_host);
        }
        let expanded_previous = CrossTablePrecedents {
            direct: self.direct_previous.clone(),
            ranges: self
                .ranges
                .iter()
                .map(|adjustment| adjustment.previous)
                .collect(),
        }
        .expanded();
        if existing != expanded_previous {
            return Err(Error::InvalidFormat(format!(
                "iWork formula host ({}, {}) has cross-table dependency edges that do not match its AST",
                previous_host.0, previous_host.1
            )));
        }
        let expanded_current = CrossTablePrecedents {
            direct: self.direct_current.clone(),
            ranges: self
                .ranges
                .iter()
                .map(|adjustment| adjustment.current)
                .collect(),
        }
        .expanded();
        Self::write_external_edges(edges, &expanded_current, previous_host)
    }

    fn write_external_edges(
        edges: &mut tsce::ExpandedEdgesArchive,
        current: &BTreeSet<ExternalCellPrecedent>,
        previous_host: (u32, u32),
    ) -> Result<()> {
        let expected_owners = current
            .iter()
            .map(|precedent| precedent.owner_id)
            .collect::<Vec<_>>();
        if edges.internal_owner_id_for_edge != expected_owners {
            return Err(Error::ParseError(format!(
                "Cannot relocate iWork formula host ({}, {}) because its cross-table dependency owner sequence would change",
                previous_host.0, previous_host.1
            )));
        }
        edges.edge_with_owner_rows = current.iter().map(|precedent| precedent.row).collect();
        edges.edge_with_owner_columns = current.iter().map(|precedent| precedent.column).collect();
        Ok(())
    }
}

/// Calculation-engine identity and bounds for one table addressable from a
/// cross-table formula AST.
#[derive(Clone, Copy, Debug)]
struct ExternalFormulaOwner {
    internal_owner_id: u32,
    rows: u32,
    columns: u32,
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
                if message.type_ != FORMULA_OWNER_MESSAGE_TYPE {
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
        mutate_incoming_range_dependencies(
            archive,
            owner_id,
            internal_owner_id,
            &previous.formula_owner_uid,
            axis,
            position,
            mutation,
            retained_hosts,
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
                .position(|message| message.type_ == CELL_DEPENDENCY_TILE_MESSAGE_TYPE)
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
                    type_: CELL_DEPENDENCY_TILE_MESSAGE_TYPE,
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
                .position(|message| message.type_ == RANGE_DEPENDENCY_TILE_MESSAGE_TYPE)
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
    })?;
    if adjustments.external_precedents.is_empty() {
        return Ok(());
    }
    let table_id = attached_table_descriptors(package)?
        .into_iter()
        .find(|table| table.table_info_id == table_info_id)
        .map(|table| table.object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table info {table_info_id} has no attached table model"
            ))
        })?;
    let hosts = adjustments
        .external_precedents
        .keys()
        .map(|&(row, column)| {
            Ok((
                usize::try_from(row)
                    .map_err(|_| Error::ParseError("iWork formula row exceeds usize".to_owned()))?,
                usize::try_from(column).map_err(|_| {
                    Error::ParseError("iWork formula column exceeds usize".to_owned())
                })?,
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    formula_cache::refresh_formula_caches_at_hosts(package, table_id, &hosts)?;
    Ok(())
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
            .find(|message| message.type_ == RANGE_DEPENDENCY_TILE_MESSAGE_TYPE)
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
        let external_adjustment = adjustments.external_precedents.get(&host);
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
        } else if let Some(reference) = &mut dependency.internal_range_reference
            && let Some(adjustment) = external_adjustment
            && let Some(replacement) =
                adjustment.replacement_for_range(ExternalRangePrecedent::from_range_coordinate(
                    reference.owner_id,
                    &reference.range,
                )?)?
        {
            replacement.write_range_coordinate(&mut reference.range);
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
        let external_adjustment = adjustments.external_precedents.get(&host);
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
        } else if let Some(adjustment) = external_adjustment
            && let Some(replacement) =
                adjustment.replacement_for_range(ExternalRangePrecedent::from_cell_rect(
                    tile.to_owner_id,
                    &dependency.refers_to_rect,
                )?)?
        {
            replacement.write_cell_rect(&mut dependency.refers_to_rect)?;
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
    for reference in &mut references.table_refs {
        if let Some(coordinates) = &mut reference.coord_set {
            mutate_cell_coord_set(coordinates, axis, position, mutation, retained_hosts)?;
        }
    }
    for table in &mut references.table_uuid_refs {
        for reference in &mut table.uuid_refs {
            if let Some(coordinates) = &mut reference.coord_set {
                mutate_cell_coord_set(coordinates, axis, position, mutation, retained_hosts)?;
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
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    for column in &mut coordinates.column_entries {
        if axis == DependencyAxis::Column {
            if mutation == DependencyMutation::Delete && column.column == position {
                if !row_set_is_fully_retained(&column.row_set, column.column, retained_hosts)? {
                    return Err(Error::ParseError(
                        "Cannot delete an iWork UUID-reference host column containing formulas that are not retained by a merged cell"
                            .to_owned(),
                    ));
                }
            } else {
                column.column =
                    mutation.coordinate(column.column, position, "UUID-reference host column")?;
            }
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
                let finish = end.unwrap_or(begin);
                if finish < begin {
                    return Err(Error::InvalidFormat(
                        "iWork UUID-reference row range is inverted".to_owned(),
                    ));
                }
                if mutation == DependencyMutation::Delete
                    && (begin..=finish).contains(&position)
                    && retained_hosts.contains(position, column.column)
                {
                    if finish > position {
                        entry.range_end = Some(i32::try_from(finish - 1).map_err(|_| {
                            Error::ParseError("iWork UUID-reference row overflow".to_owned())
                        })?);
                    }
                    continue;
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

/// Return whether every coordinate encoded by one UUID column entry belongs to
/// an anchor retained through the current deletion.
///
/// A column deletion cannot split a `ColumnEntry` without changing its native
/// protobuf shape. It is safe to preserve that entry only when every one of
/// its rows is a retained formula host. Row deletion has a denser native
/// representation and is handled directly above.
fn row_set_is_fully_retained(
    rows: &tsce::IndexSetArchive,
    column: u32,
    retained_hosts: &FormulaHostRetentions,
) -> Result<bool> {
    if rows.entries.is_empty() {
        return Ok(false);
    }
    for entry in &rows.entries {
        let begin = u32::try_from(entry.range_begin)
            .map_err(|_| Error::InvalidFormat("iWork UUID-reference row is negative".to_owned()))?;
        let finish = entry
            .range_end
            .map(u32::try_from)
            .transpose()
            .map_err(|_| Error::InvalidFormat("iWork UUID-reference row is negative".to_owned()))?
            .unwrap_or(begin);
        if finish < begin {
            return Err(Error::InvalidFormat(
                "iWork UUID-reference row range is inverted".to_owned(),
            ));
        }
        let count = u64::from(finish)
            .checked_sub(u64::from(begin))
            .and_then(|difference| difference.checked_add(1))
            .ok_or_else(|| {
                Error::ParseError("iWork UUID-reference row range overflow".to_owned())
            })?;
        if count
            > u64::try_from(retained_hosts.hosts.len()).map_err(|_| {
                Error::ParseError("iWork formula-host retention count exceeds u64".to_owned())
            })?
        {
            return Ok(false);
        }
        for row in begin..=finish {
            if !retained_hosts.contains(row, column) {
                return Ok(false);
            }
        }
    }
    Ok(true)
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

/// Rebase native range-proxy owners that point into a merged formula anchor.
///
/// Native Numbers represents merged formula geometry with detached calculation
/// owners whose ranges point back at the visible table. Those are not user
/// formulas, so they have no formula AST to rewrite. When a retained merge
/// anchor survives a deletion, its proxy range must contract around the
/// retained host in the same archive transaction. Ordinary cross-table
/// formulas remain deliberately rejected here because their AST lives in a
/// different table and needs a separate rewrite operation.
fn mutate_incoming_range_dependencies(
    archive: &mut Archive,
    target_object_id: u64,
    target_internal_owner_id: u32,
    target_owner_uid: &tsp::Uuid,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    let mut incoming_owners = Vec::new();
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ != FORMULA_OWNER_MESSAGE_TYPE {
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
                    dependencies
                        .back_dependency
                        .iter()
                        .any(|dependency| dependency.range_reference.is_some())
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
                    .find(|message| message.type_ == CELL_DEPENDENCY_TILE_MESSAGE_TYPE)
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
            let mut incoming_range_tiles = Vec::new();
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
                    .find(|message| message.type_ == RANGE_DEPENDENCY_TILE_MESSAGE_TYPE)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers range dependency tile {} has no payload",
                            reference.identifier
                        ))
                    })?;
                let tile = tsce::RangePrecedentsTileArchive::decode(tile_message.data.as_slice())?;
                if tile.to_owner_id == target_internal_owner_id && !tile.from_to_range.is_empty() {
                    incoming_range_tiles.push(reference.identifier);
                }
            }
            let has_incoming_range =
                owner
                    .range_dependencies
                    .as_ref()
                    .is_some_and(|dependencies| {
                        dependencies.back_dependency.iter().any(|dependency| {
                            dependency
                                .internal_range_reference
                                .as_ref()
                                .is_some_and(|reference| {
                                    reference.owner_id == target_internal_owner_id
                                })
                        })
                    })
                    || !incoming_range_tiles.is_empty();
            if !has_incoming_range {
                continue;
            }
            if owner.formula_owner.is_some() {
                return Err(incoming_dependency_error(axis, mutation));
            }
            incoming_owners.push((
                object.archive_info.identifier.ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers incoming range-proxy owner is missing its object ID".to_owned(),
                    )
                })?,
                owner,
                incoming_range_tiles,
            ));
        }
    }

    let mut rewritten_range_tiles = BTreeSet::new();
    for (object_id, previous, range_tile_ids) in incoming_owners {
        let mut current = previous.clone();
        if let Some(dependencies) = &mut current.range_dependencies {
            for dependency in &mut dependencies.back_dependency {
                if let Some(reference) = &mut dependency.internal_range_reference
                    && reference.owner_id == target_internal_owner_id
                {
                    mutate_incoming_target_range_coordinate(
                        &mut reference.range,
                        axis,
                        position,
                        mutation,
                        retained_hosts,
                    )?;
                }
            }
        }
        let source_object = archive.object(object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers incoming range-proxy owner {object_id} is missing"
            ))
        })?;
        let message_index = source_object
            .messages
            .iter()
            .position(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers incoming range-proxy owner {object_id} has no payload"
                ))
            })?;
        let original = source_object.messages[message_index].data.clone();
        let data = rewrite_shifted_formula_owner_wire(&original, &previous, &current, axis)?;
        let source_object = archive.object_mut(object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers incoming range-proxy owner {object_id} disappeared"
            ))
        })?;
        let message_type = source_object.messages[message_index].type_;
        source_object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;

        for tile_id in range_tile_ids {
            if !rewritten_range_tiles.insert(tile_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers incoming range-proxy owners share range tile {tile_id}"
                )));
            }
            let tile_object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers incoming range dependency tile {tile_id} is missing"
                ))
            })?;
            let message_index = tile_object
                .messages
                .iter()
                .position(|message| message.type_ == RANGE_DEPENDENCY_TILE_MESSAGE_TYPE)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers incoming range dependency tile {tile_id} has no payload"
                    ))
                })?;
            let original = tile_object.messages[message_index].data.clone();
            let previous_tile = tsce::RangePrecedentsTileArchive::decode(original.as_slice())?;
            if previous_tile.to_owner_id != target_internal_owner_id {
                return Err(Error::InvalidFormat(format!(
                    "Numbers incoming range dependency tile {tile_id} has an unexpected target owner"
                )));
            }
            let mut current_tile = previous_tile.clone();
            for dependency in &mut current_tile.from_to_range {
                mutate_incoming_target_cell_rect(
                    &mut dependency.refers_to_rect,
                    axis,
                    position,
                    mutation,
                    retained_hosts,
                )?;
            }
            let data = rewrite_shifted_range_tile_wire(&original, &previous_tile, &current_tile)?;
            let message_type = tile_object.messages[message_index].type_;
            tile_object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
        }
    }
    Ok(())
}

fn mutate_incoming_target_range_coordinate(
    range: &mut tsce::RangeCoordinateArchive,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    let retains_deleted_coordinate = mutation == DependencyMutation::Delete
        && range_contains_retained_host(range, axis, position, retained_hosts);
    let mutate_coordinate = |coordinate: u32, what: &str| {
        if retains_deleted_coordinate && coordinate == position {
            Ok(position)
        } else {
            mutation.coordinate(coordinate, position, what)
        }
    };
    match axis {
        DependencyAxis::Column => {
            range.top_left_column = mutate_coordinate(range.top_left_column, "range start column")?;
            range.bottom_right_column =
                mutate_coordinate(range.bottom_right_column, "range end column")?;
        },
        DependencyAxis::Row => {
            range.top_left_row = mutate_coordinate(range.top_left_row, "range start row")?;
            range.bottom_right_row = mutate_coordinate(range.bottom_right_row, "range end row")?;
        },
    }
    if range.top_left_row > range.bottom_right_row
        || range.top_left_column > range.bottom_right_column
    {
        return Err(Error::InvalidFormat(
            "Numbers incoming range-proxy target became inverted".to_owned(),
        ));
    }
    Ok(())
}

fn range_contains_retained_host(
    range: &tsce::RangeCoordinateArchive,
    axis: DependencyAxis,
    position: u32,
    retained_hosts: &FormulaHostRetentions,
) -> bool {
    let contains_deleted_axis = match axis {
        DependencyAxis::Column => {
            (range.top_left_column..=range.bottom_right_column).contains(&position)
        },
        DependencyAxis::Row => (range.top_left_row..=range.bottom_right_row).contains(&position),
    };
    contains_deleted_axis
        && retained_hosts.hosts.iter().any(|host| {
            (range.top_left_row..=range.bottom_right_row).contains(&host.row)
                && (range.top_left_column..=range.bottom_right_column).contains(&host.column)
        })
}

fn mutate_incoming_target_cell_rect(
    rect: &mut tsce::CellRectArchive,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    retained_hosts: &FormulaHostRetentions,
) -> Result<()> {
    let (top, left) = explicit_cell_coordinate(&rect.origin, "incoming range origin")?;
    let rows = rect.size.num_rows.unwrap_or(1);
    let columns = rect.size.num_columns.unwrap_or(1);
    if rows == 0 || columns == 0 {
        return Err(Error::InvalidFormat(
            "Numbers incoming range-proxy tile has an empty rectangle".to_owned(),
        ));
    }
    let mut range = tsce::RangeCoordinateArchive {
        top_left_column: left,
        top_left_row: top,
        bottom_right_column: left.checked_add(columns - 1).ok_or_else(|| {
            Error::ParseError("Numbers incoming range-proxy column overflow".to_owned())
        })?,
        bottom_right_row: top.checked_add(rows - 1).ok_or_else(|| {
            Error::ParseError("Numbers incoming range-proxy row overflow".to_owned())
        })?,
    };
    mutate_incoming_target_range_coordinate(&mut range, axis, position, mutation, retained_hosts)?;
    rect.origin.column = Some(range.top_left_column);
    rect.origin.row = Some(range.top_left_row);
    let columns = range
        .bottom_right_column
        .checked_sub(range.top_left_column)
        .and_then(|width| width.checked_add(1))
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers incoming range-proxy is inverted".to_owned())
        })?;
    let rows = range
        .bottom_right_row
        .checked_sub(range.top_left_row)
        .and_then(|height| height.checked_add(1))
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers incoming range-proxy is inverted".to_owned())
        })?;
    rect.size.num_columns = (columns != 1).then_some(columns);
    rect.size.num_rows = (rows != 1).then_some(rows);
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
    let external_adjustment = adjustments.external_precedents.get(&previous_host);
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
        if external_adjustment.is_some_and(ExternalPrecedentAdjustments::has_direct_rewrite) {
            return Err(Error::InvalidFormat(format!(
                "iWork formula host ({}, {}) has cross-table references but no dependency edges",
                previous_host.0, previous_host.1
            )));
        }
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
    if let Some(adjustment) = external_adjustment {
        adjustment.rewrite_expanded_edges(edges, previous_host)?;
    }
    Ok(())
}
