//! Implementation of the public [`crate::cell::text_style`] transactions:
//! character properties (bold, italic, underline, size, font, colour) of
//! table cells (`IWORK_API_REQUESTS.md` request 7, record 0309).
//!
//! A Numbers cell's text style is a key in its table's style list
//! (`DataStore.styleTable`) whose entry references a
//! `TSWP.ParagraphStyleArchive`; a cell without a key uses its section's
//! base text style (header rows, footer rows, header columns, body). A set
//! creates a paragraph-style variation of the cell's current style that
//! overrides the requested character properties (every other property
//! stays inherited), inserts it beside the stylesheet and lists it there,
//! appends a style-list entry for it and points the cell at the new key;
//! the old key's entry loses the cell's reference. Each cell takes its own
//! variation, and a property equal to the inherited one stays inherited,
//! as the host editor writes them (record 0370). Ported from the host
//! editor's `cell_paragraph_style` (new-style path).

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_table::{
    cell_text_style as shared, cell_text_style_headers as headers, cell_text_style_writer as writer,
};
use litchi_iwa_text_wire::CharacterProperties;
use litchi_numbers_wire::{
    Limits as BncLimits, table_topology::storage::HEADER_BUCKET_ROWS as SHARED_HEADER_BUCKET_ROWS,
};
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{
    self, Fault, Working,
    metadata::{self, Transition},
};
use super::table_topology::{self as topology, TILE_MESSAGE_TYPE};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;

structure::structural_family!("cell-text-style", Wire(litchi_iwa_common::LimitKind));

/// `TST.HeaderStorageBucket`.
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
/// A typed, content-redacted cell text-style failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted cell text-style source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} cell text-style items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("cell text-style candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("cell text-style patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// The cell is outside the table.
    #[error("cell outside the table")]
    CellNotFound,
    /// The style sets no property.
    #[error("the style sets no property")]
    EmptyStyle,
    /// Commit was requested before staging a cell.
    #[error("no cell text style is staged")]
    NoStagedOperation,
    /// The cell's style lives outside what this transaction rewrites (no
    /// stylesheet, a segmented list, a style in another component).
    #[error("a surviving object depends on the cell's style graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support cell text styles")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound { .. } | Self::TableNotFound { .. } | Self::CellNotFound { .. } => {
                RefusalKind::NotFound
            },
            Self::EmptyStyle { .. } => RefusalKind::InvalidValue,
            Self::NoStagedOperation { .. } => RefusalKind::UnsupportedOperation,
            Self::UnsupportedDependency { .. } | Self::UnsupportedSource { .. } => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

pub(super) fn from_topology(error: topology::Error) -> Error {
    use topology::Error as E;
    match error {
        E::Refused(crate::package::table_topology::Refusal::SheetNotFound) => {
            Error::Refused(Refusal::SheetNotFound)
        },
        E::Refused(crate::package::table_topology::Refusal::TableNotFound) => {
            Error::Refused(Refusal::TableNotFound)
        },
        E::Allocation { amount } => Error::Allocation { amount },
        E::Refused(crate::package::table_topology::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        _ => Error::InvalidSource,
    }
}

#[cfg(test)]
fn from_style(error: litchi_iwa_text_wire::CharStyleError) -> Error {
    match error {
        litchi_iwa_text_wire::CharStyleError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_text_wire::CharStyleError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// One immutable edit that sets the character properties of cells of one
/// table.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
    cells: Vec<(At, CharacterProperties)>,
    limits: shared::Limits,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("cells", &self.cells.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage `properties` on the stored cell at `position`; a later stage
    /// of the same cell adds to the earlier one.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Refusal::EmptyStyle`].
    pub fn set(mut self, position: At, properties: CharacterProperties) -> Result<Self, Error> {
        if properties.is_empty() {
            return Err(Error::Refused(Refusal::EmptyStyle));
        }
        if position.row() >= self.rows || position.column() >= self.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        self.limits
            .admit_character_properties(&properties)
            .map_err(from_shared)?;
        if let Some((_, staged)) = self.cells.iter_mut().find(|(cell, _)| *cell == position) {
            self.limits
                .admit_character_properties(staged)
                .map_err(from_shared)?;
            let mut merged = properties;
            merged.inherit(staged);
            *staged = merged;
            return Ok(self);
        }
        if self.cells.len() >= litchi_iwa_table::TableSource::max_objects(self.source) {
            return Err(Error::Allocation { amount: 1 });
        }
        self.limits.admit_allocations(1).map_err(from_shared)?;
        self.cells
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.cells.push((position, properties));
        Ok(self)
    }

    /// Validate and atomically publish the staged cell styles.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] (an empty cell takes a style-only
    /// cell); [`Refusal::UnsupportedDependency`] for a style graph
    /// this transaction does not rewrite; source, resource and
    /// verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.cells.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let (published, limits) = publish(
            self.source,
            self.sheet,
            self.table,
            &self.cells,
            self.limits,
        )?;
        let effective_values = if self.cells.len() == 1 {
            limits.admit_allocations(1).map_err(from_shared)?;
            vec![text_style_with_limits(
                &published.package,
                self.sheet,
                self.table,
                self.cells[0].0,
                &limits,
            )?]
        } else {
            let mut positions = Vec::new();
            limits
                .reserve(&mut positions, self.cells.len())
                .map_err(from_shared)?;
            positions.extend(self.cells.iter().map(|(position, _)| *position));
            text_style_many_with_limits(
                &published.package,
                self.sheet,
                self.table,
                &positions,
                &limits,
            )?
        };
        for ((_, properties), effective) in self.cells.iter().zip(effective_values) {
            limits
                .admit_character_properties(properties)
                .map_err(from_shared)?;
            limits
                .admit_character_properties(&effective)
                .map_err(from_shared)?;
            let mut expected = properties.clone();
            expected.inherit(&effective);
            limits
                .admit_character_properties(&expected)
                .map_err(from_shared)?;
            limits
                .admit_character_properties(&effective)
                .map_err(from_shared)?;
            if with_flags_off(expected.clone()) != with_flags_off(effective.clone()) {
                return Err(Error::Verification);
            }
        }
        // Every stored cell survives; an empty cell may gain a style-only
        // cell (record 0370).
        let count = |package: &Package| {
            package.workbook_ref().sheets()[self.sheet]
                .tables()
                .nth(self.table)
                .map(|table| table.iter_cells().count())
        };
        let (Some(before), Some(after)) = (count(self.source), count(&published.package)) else {
            return Err(Error::Verification);
        };
        if after < before || after > before.saturating_add(self.cells.len()) {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

/// The shared rich-text engine's error (record 0412).
impl From<litchi_iwa_table::cell_rich_text::Error> for Error {
    fn from(error: litchi_iwa_table::cell_rich_text::Error) -> Self {
        use litchi_iwa_table::cell_rich_text::{Error as E, LimitKind as K, Refusal as R};
        match error {
            E::Refused(R::CellNotFound) => Self::Refused(Refusal::CellNotFound),
            E::Refused(R::UnsupportedSource) => Self::Refused(Refusal::UnsupportedSource),
            E::Refused(_) => Self::Refused(Refusal::UnsupportedDependency),
            E::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    K::PackageBytes => LimitKind::PackageBytes,
                    K::Identifiers => LimitKind::Identifiers,
                    K::Wire(kind) => LimitKind::Wire(kind),
                    _ => LimitKind::PayloadBytes,
                },
                observed,
                maximum,
            },
            E::Allocation { amount } => Self::Allocation { amount },
            E::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

fn text_style(
    package: &Package,
    sheet: usize,
    table: usize,
    position: At,
) -> Result<CharacterProperties, Error> {
    let limits = shared::Limits::from_source(package).map_err(from_shared)?;
    text_style_with_limits(package, sheet, table, position, &limits)
}

fn text_style_with_limits(
    package: &Package,
    sheet: usize,
    table: usize,
    position: At,
    limits: &shared::Limits,
) -> Result<CharacterProperties, Error> {
    let working = Working::new(package)?;
    let target = topology::locate(&working, sheet, table).map_err(from_topology)?;
    shared::read(
        package,
        shared::Target {
            model_identifier: target.model,
            row: position.row(),
            column: position.column(),
        },
        limits,
    )
    .map_err(from_shared)
}

fn text_style_many_with_limits(
    package: &Package,
    sheet: usize,
    table: usize,
    positions: &[At],
    limits: &shared::Limits,
) -> Result<Vec<CharacterProperties>, Error> {
    let working = Working::new(package)?;
    let target = topology::locate(&working, sheet, table).map_err(from_topology)?;
    let mut targets = Vec::new();
    limits
        .reserve(&mut targets, positions.len())
        .map_err(from_shared)?;
    targets.extend(positions.iter().map(|position| shared::Target {
        model_identifier: target.model,
        row: position.row(),
        column: position.column(),
    }));
    shared::read_many(package, &targets, limits).map_err(from_shared)
}

fn from_shared(error: shared::Error) -> Error {
    match error {
        shared::Error::InvalidSource(_) => Error::InvalidSource,
        shared::Error::Refused(refusal) => match refusal {
            shared::Refusal::OutsideTable => Error::Refused(Refusal::CellNotFound),
            shared::Refusal::UnsupportedStyleList | shared::Refusal::UnsupportedDependency => {
                Error::Refused(Refusal::UnsupportedDependency)
            },
            shared::Refusal::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
            _ => Error::InvalidSource,
        },
        shared::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed,
            maximum,
        },
        shared::Error::Allocation { amount } => Error::Allocation { amount },
        shared::Error::NeedsIdentifier => Error::InvalidSource,
        shared::Error::BatchUnsupported => Error::Refused(Refusal::UnsupportedDependency),
        _ => Error::InvalidSource,
    }
}

/// Apply the shared format-neutral style graph plan through Numbers'
/// immutable working set and metadata publisher.  All graph decisions and
/// payload construction come from `litchi-iwa-table`; this adapter owns only
/// identifier allocation, archive headers, registry transitions and cell
/// publication.
fn publish(
    source: &Package,
    sheet: usize,
    table: usize,
    cells: &[(At, CharacterProperties)],
    limits: shared::Limits,
) -> Result<(structure::Published, shared::Limits), Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let target = topology::locate(&working, sheet, table).map_err(from_topology)?;
    if cells.len() == 1 {
        let (position, properties) = &cells[0];
        let shared_target = shared::Target {
            model_identifier: target.model,
            row: position.row(),
            column: position.column(),
        };
        let plan =
            match shared::plan_set_precharged(source, shared_target, properties, None, &limits) {
                Ok(plan) => plan,
                Err(shared::Error::NeedsIdentifier) => {
                    limits.admit_allocations(1).map_err(from_shared)?;
                    let fresh = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
                    shared::plan_set_precharged(
                        source,
                        shared_target,
                        properties,
                        Some(fresh),
                        &limits,
                    )
                    .map_err(from_shared)?
                },
                Err(error) => return Err(from_shared(error)),
            };
        let Some(plan) = plan else {
            return working
                .publish(&Transition::default())
                .map(|published| (published, limits))
                .map_err(Error::from);
        };
        let mut transition = Transition::default();
        apply_shared_plan(source, &mut working, &mut transition, plan, &limits)?;
        return working
            .publish(&transition)
            .map(|published| (published, limits))
            .map_err(Error::from);
    }
    let mut requests = Vec::new();
    limits
        .reserve(&mut requests, cells.len())
        .map_err(from_shared)?;
    for (position, properties) in cells {
        limits
            .admit_character_properties(properties)
            .map_err(from_shared)?;
        let shared_target = shared::Target {
            model_identifier: target.model,
            row: position.row(),
            column: position.column(),
        };
        requests.push(shared::Request {
            target: shared_target,
            properties: properties.clone(),
            fresh: None,
        });
    }
    let prepare_batch = cells.len() >= 64;
    let mut allocate = || {
        limits.admit_allocations(1)?;
        working
            .allocate(1)
            .map_err(|_| shared::Error::Allocation { amount: 1 })?
            .first()
            .copied()
            .ok_or(shared::Error::InvalidSource("identifier allocation failed"))
    };
    let plan = shared::plan_set_many_with_allocator_precharged(
        source,
        &requests,
        &limits,
        prepare_batch,
        &mut allocate,
    )
    .map_err(from_shared)?;
    let Some(plan) = plan else {
        return working
            .publish(&Transition::default())
            .map(|published| (published, limits))
            .map_err(Error::from);
    };
    let mut transition = Transition::default();
    apply_shared_plan(source, &mut working, &mut transition, plan, &limits)?;
    working
        .publish(&transition)
        .map(|published| (published, limits))
        .map_err(Error::from)
}

fn apply_shared_plan(
    source: &Package,
    working: &mut Working<'_>,
    transition: &mut Transition,
    plan: shared::Plan,
    limits: &shared::Limits,
) -> Result<(), Error> {
    admit_publication(&plan, limits)?;
    let registry =
        metadata::registry(working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    for rewrite in &plan.rewrites {
        let object = working.object(rewrite.identifier)?;
        let index = structure::unique_message(object, &[rewrite.message_type])?;
        if rewrite.add_references.is_empty() && rewrite.remove_references.is_empty() {
            replace_message(working, rewrite.identifier, index, rewrite.data.clone())?;
        } else {
            match replace_with_references(
                working,
                rewrite.identifier,
                index,
                rewrite.data.clone(),
                &rewrite.remove_references,
                &rewrite.add_references,
            ) {
                Ok(()) => {},
                Err(Fault::Invalid) => {
                    replace_message(working, rewrite.identifier, index, rewrite.data.clone())?;
                    let info = &mut working
                        .object_mut(rewrite.identifier)?
                        .archive_info
                        .message_infos
                        .get_mut(index)
                        .ok_or(Fault::Invalid)?;
                    info.object_references
                        .retain(|reference| !rewrite.remove_references.contains(reference));
                    for reference in &rewrite.add_references {
                        if !info.object_references.contains(reference) {
                            info.object_references.push(*reference);
                        }
                    }
                },
                Err(error) => return Err(error.into()),
            }
        }
    }
    for insertion in &plan.insertions {
        let parent = insertion
            .object_references
            .first()
            .copied()
            .ok_or(Error::InvalidSource)?;
        let template = working.object(parent)?;
        let object = structure::new_object(
            insertion.identifier,
            Some(template),
            insertion.message_type,
            insertion.data.clone(),
            insertion.object_references.clone(),
            working.archive_limits(),
        )?;
        let entry = working
            .locate(insertion.beside)?
            .ok_or(Error::InvalidSource)?;
        working.insert(&entry, object)?;
    }
    for registration in &plan.registrations {
        match *registration {
            shared::Registration::Uuid { object } => {
                let entry = working.locate(object)?.ok_or(Error::InvalidSource)?;
                let component = registry.component_for_entry(&entry)?.identifier;
                let seed = metadata::identity_seed(working, object)?;
                transition.register_uuid(&registry, component, object, seed)?;
            },
            shared::Registration::ExternalReference { from, object } => {
                let from_entry = working.locate(from)?.ok_or(Error::InvalidSource)?;
                let object_entry = working.locate(object)?.ok_or(Error::InvalidSource)?;
                let from_component = registry.component_for_entry(&from_entry)?.identifier;
                let object_component = registry.component_for_entry(&object_entry)?.identifier;
                transition.reference(&registry, from_component, object_component, object)?;
            },
            shared::Registration::Removed { object } => {
                transition.remove_objects(&registry, &[object])?;
                transition.release_suffix(&[object])?;
                working.remove(object)?;
            },
        }
    }
    write_style_keys(source, working, &plan.cell_updates, limits)?;
    Ok(())
}

/// Admit the second owned copy made when a neutral plan is published into the
/// mutable Numbers working set.  The shared planner charges its own source
/// graph and generated output; this boundary separately owns the copies held
/// by `RawMessage`, `NewObject`, and the compatibility replacement path.
fn admit_publication(plan: &shared::Plan, limits: &shared::Limits) -> Result<(), Error> {
    let mut copied = 0_usize;
    let mut allocations = 0_usize;
    let mut references = 0_usize;
    let mut work = 0_usize;
    for rewrite in &plan.rewrites {
        let copies = usize::from(
            !(rewrite.add_references.is_empty() && rewrite.remove_references.is_empty()),
        )
        .saturating_add(1);
        copied = copied.saturating_add(rewrite.data.len().saturating_mul(copies));
        allocations = allocations.saturating_add(1);
        references = references
            .saturating_add(rewrite.add_references.len())
            .saturating_add(rewrite.remove_references.len());
        work = work.saturating_add(rewrite.data.len()).saturating_add(1);
    }
    for insertion in &plan.insertions {
        copied = copied.saturating_add(insertion.data.len());
        allocations = allocations.saturating_add(2);
        references = references.saturating_add(insertion.object_references.len());
        work = work.saturating_add(insertion.data.len()).saturating_add(1);
    }
    allocations = allocations
        .saturating_add(plan.registrations.len())
        .saturating_add(plan.cell_updates.len());
    work = work
        .saturating_add(plan.registrations.len())
        .saturating_add(plan.cell_updates.len());
    limits.admit_copies(copied).map_err(from_shared)?;
    limits.admit_allocations(allocations).map_err(from_shared)?;
    limits.admit_references(references).map_err(from_shared)?;
    limits.admit_work(work).map_err(from_shared)
}

/// Publish all cell-key updates through the shared bounded tile writer. The
/// writer owns existing-cell, empty-slot and missing-row handling under the
/// same staged-to-readback transaction ledger.
fn write_style_keys(
    source: &Package,
    working: &mut Working<'_>,
    updates: &[(shared::Target, u32)],
    limits: &shared::Limits,
) -> Result<(), Error> {
    let Some((first_target, _)) = updates.first().copied() else {
        return Ok(());
    };
    limits.admit_work(updates.len()).map_err(from_shared)?;
    if updates
        .iter()
        .any(|(target, _)| target.model_identifier != first_target.model_identifier)
    {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    let facts = shared::table_facts(source, first_target, limits).map_err(from_shared)?;
    let shared::TableFacts {
        rows: table_rows,
        columns: table_columns,
        tile_size,
        row_headers,
        column_headers,
        tiles,
    } = facts;
    limits.admit_work(updates.len()).map_err(from_shared)?;
    limits
        .admit_work(updates.len().saturating_mul(tiles.len()))
        .map_err(from_shared)?;
    limits
        .admit_work(updates.len().saturating_mul(updates.len()))
        .map_err(from_shared)?;
    let mut groups: Vec<(u64, Vec<writer::Update>)> = Vec::new();
    limits.reserve(&mut groups, 1).map_err(from_shared)?;
    for (target, key) in updates {
        if target.row >= table_rows || target.column >= table_columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        let tile_key = u64::from(target.row) / tile_size;
        let local =
            u32::try_from(u64::from(target.row) % tile_size).map_err(|_| Error::InvalidSource)?;
        let tile = tiles
            .iter()
            .find(|(candidate, _)| u64::from(*candidate) == tile_key)
            .map(|(_, identifier)| *identifier)
            .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
        let update = writer::Update {
            row: local,
            column: target.column,
            key: Some(*key),
        };
        if let Some((_, rows)) = groups
            .iter_mut()
            .find(|(identifier, _)| *identifier == tile)
        {
            limits.reserve(rows, 1).map_err(from_shared)?;
            rows.push(update);
        } else {
            let mut rows = Vec::new();
            limits.reserve(&mut rows, 1).map_err(from_shared)?;
            rows.push(update);
            limits.reserve(&mut groups, 1).map_err(from_shared)?;
            groups.push((tile, rows));
        }
    }
    let bnc_limits = BncLimits::from_wire_limits(limits.wire);
    limits
        .admit_work(groups.len().saturating_mul(tiles.len()))
        .map_err(from_shared)?;
    for (tile, rows) in groups {
        let (tile_index, tile_data) = topology::message(working, tile, &[TILE_MESSAGE_TYPE])?;
        let write = writer::set_text_style_keys(
            tile_data,
            table_columns,
            &rows,
            writer::RowPlacement::FieldOrder,
            limits,
            &bnc_limits,
        )
        .map_err(from_shared)?;
        let tile_key = tiles
            .iter()
            .find(|(_, identifier)| *identifier == tile)
            .map(|(key, _)| *key)
            .ok_or(Error::InvalidSource)?;
        let row_offset = tile_key
            .checked_mul(u32::try_from(tile_size).map_err(|_| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        let deltas = headers::materialized_deltas(&rows, write.previous(), row_offset, limits)
            .map_err(from_shared)?;
        rewrite_materialized_headers(working, &row_headers, column_headers, &deltas, limits)?;
        if write.changed() {
            replace_message(working, tile, tile_index, write.into_tile())?;
        }
    }
    Ok(())
}

fn rewrite_materialized_headers(
    working: &mut Working<'_>,
    row_headers: &[(u32, u64)],
    column_headers: Option<u64>,
    deltas: &headers::MaterializedHeaderDeltas,
    limits: &shared::Limits,
) -> Result<(), Error> {
    if deltas.is_empty() {
        return Ok(());
    }
    let row_groups = materialized_header_groups(row_headers, column_headers, deltas, limits)?;
    limits
        .admit_work(row_groups.len().saturating_mul(row_headers.len()))
        .map_err(from_shared)?;
    for (bucket, bucket_deltas) in row_groups {
        let row_identifier = row_headers
            .iter()
            .find(|(candidate, _)| u64::from(*candidate) == bucket)
            .map(|(_, identifier)| *identifier)
            .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
        let (index, payload) =
            topology::message(working, row_identifier, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let rewrite = headers::rewrite_header_buckets(payload, &[], &bucket_deltas, limits)
            .map_err(from_shared)?;
        let (row, _) = rewrite.into_parts();
        let row = row.ok_or(Error::InvalidSource)?;
        replace_message(working, row_identifier, index, row)?;
    }
    if !deltas.columns().is_empty() {
        let column_identifier =
            column_headers.ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
        let column_deltas = deltas.columns_only(limits).map_err(from_shared)?;
        let (index, payload) =
            topology::message(working, column_identifier, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let rewrite = headers::rewrite_header_buckets(&[], payload, &column_deltas, limits)
            .map_err(from_shared)?;
        let (_, column) = rewrite.into_parts();
        let column = column.ok_or(Error::InvalidSource)?;
        replace_message(working, column_identifier, index, column)?;
    }
    Ok(())
}

fn materialized_header_groups(
    row_headers: &[(u32, u64)],
    column_headers: Option<u64>,
    deltas: &headers::MaterializedHeaderDeltas,
    limits: &shared::Limits,
) -> Result<Vec<(u64, headers::MaterializedHeaderDeltas)>, Error> {
    let groups = deltas
        .split_rows_by_bucket(SHARED_HEADER_BUCKET_ROWS, limits)
        .map_err(from_shared)?;
    limits
        .admit_work(groups.len().saturating_mul(row_headers.len()))
        .map_err(from_shared)?;
    if groups.iter().any(|(bucket, _)| {
        !row_headers
            .iter()
            .any(|(candidate, _)| u64::from(*candidate) == *bucket)
    }) {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    if !deltas.columns().is_empty() && column_headers.is_none() {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(groups)
}

/// `properties` with every unset emphasis flag read as off.
fn with_flags_off(mut properties: CharacterProperties) -> CharacterProperties {
    for flag in [
        &mut properties.bold,
        &mut properties.italic,
        &mut properties.underline,
        &mut properties.strikethrough,
    ] {
        flag.get_or_insert(false);
    }
    properties
}

/// Replace one message, transitioning its aggregate object references, or
/// keeping its header (the shared engine's helpers, record 0412).
pub(super) use litchi_iwa_table::cell_rich_text::{replace_message, replace_with_references};

impl From<topology::Error> for Error {
    fn from(error: topology::Error) -> Self {
        from_topology(error)
    }
}

impl Package {
    /// The effective character properties of one cell's text (its text
    /// style and that style's parents; a cell without its own style uses
    /// its section's base style).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::CellNotFound`], source errors.
    pub fn table_cell_text_style<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<CharacterProperties, Error> {
        let (sheet, table) = self.resolve_text_style_target(sheet.into(), table.into())?;
        text_style(self, sheet, table, position)
    }

    /// Start an immutable cell text-style edit of one table.
    ///
    /// ```no_run
    /// use litchi_numbers::cell::text_style::CharacterProperties;
    /// use litchi_numbers::{At, Package};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut bold = CharacterProperties::new();
    /// bold.bold = Some(true);
    /// let commit = package
    ///     .edit_table_cell_text_style(0, 0)?
    ///     .set(At::new(1, 1), bold)?
    ///     .commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_table_cell_text_style(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_cell_text_style<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet, table) = self.resolve_text_style_target(sheet.into(), table.into())?;
        let leaf = self.workbook_ref().sheets()[sheet]
            .tables()
            .nth(table)
            .ok_or(Error::Refused(Refusal::TableNotFound))?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            rows: leaf.row_count(),
            columns: leaf.column_count(),
            cells: Vec::new(),
            limits: shared::Limits::from_source(self).map_err(from_shared)?,
        })
    }

    /// Apply a cell text-style patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_cell_text_style(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }

    fn resolve_text_style_target(
        &self,
        sheet: SheetSelector<'_>,
        table: TableSelector<'_>,
    ) -> Result<(usize, usize), Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = topology::table_position(self, sheet, table).map_err(from_topology)?;
        Ok((sheet, table))
    }
}

#[cfg(test)]
mod wire_error_tests {
    use super::{Error, LimitKind, from_style, materialized_header_groups};
    use litchi_iwa_table::{
        cell_text_style as shared, cell_text_style_headers as headers,
        cell_text_style_writer as writer,
    };

    #[test]
    fn character_style_resource_errors_survive_the_owner_boundary() {
        let error = from_style(litchi_iwa_text_wire::CharStyleError::Limit {
            kind: litchi_iwa_common::LimitKind::RewriteWork,
            observed: 11,
            maximum: 10,
        });
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Wire(litchi_iwa_common::LimitKind::RewriteWork),
                observed: 11,
                maximum: 10,
            }
        ));
        assert!(matches!(
            from_style(litchi_iwa_text_wire::CharStyleError::Allocation { amount: 42 }),
            Error::Allocation { amount: 42 }
        ));
        assert!(matches!(
            from_style(litchi_iwa_text_wire::CharStyleError::Invalid),
            Error::InvalidSource
        ));
    }

    #[test]
    fn materialization_without_header_ownership_is_typed() {
        let package = super::Package::new();
        let limits = shared::Limits::from_source(&package).expect("limits");
        let updates = [writer::Update {
            row: 0,
            column: 0,
            key: Some(1),
        }];
        let deltas = headers::materialized_deltas(&updates, &[None], 0, &limits)
            .expect("materialized deltas");
        assert!(matches!(
            materialized_header_groups(&[], None, &deltas, &limits),
            Err(Error::Refused(super::Refusal::UnsupportedDependency))
        ));
    }
}
