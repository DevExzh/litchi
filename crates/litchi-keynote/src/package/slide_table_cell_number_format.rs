//! Exact-source transactions for an existing Keynote table-cell Number
//! format.
//!
//! The package adapter owns the native table-model, tile, BNC, and format-list
//! graph. Public callers select a slide and table semantically and pass the
//! checked [`crate::slide::table::number_format::CellPosition`] coordinate;
//! native object identifiers and list keys never cross this module's public
//! boundary.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::too_many_arguments,
    reason = "The package boundary deliberately redacts native failure detail."
)]

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use crate::slide::table::number_format::CellPosition;
use crate::slide::table::number_format::{
    DecimalPlaces, FixedDecimalPlaces, NegativeStyle, Number, ThousandsSeparator,
};
use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{decode_varint_from_bytes, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    numbers_table_cell_number_format_codec as number_codec,
    numbers_table_cell_storage_codec as storage_codec,
};
use litchi_numbers_wire::BncCell;
use thiserror::Error;

use super::slide_table_core as core;
use super::{Package, PayloadLimitKind, PhysicalSource, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::TableSelector;

const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE: u32 = 6_201;
const FORMAT_LIST_TYPE: i32 = 2;

/// Finite resources charged by an existing Keynote table-cell Number-format
/// transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableCellNumberFormatLimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    PayloadObjects,
    PayloadMessages,
    References,
    WireFields,
    WireNesting,
    WireWork,
    Allocations,
    Retained,
    Scratch,
    Components,
}

impl fmt::Display for SlideTableCellNumberFormatLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-free semantic location for one Keynote table-cell Number-format
/// operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableCellNumberFormatPath {
    Package,
    Cell {
        slide: Position,
        table: Position,
        position: CellPosition,
    },
}

impl fmt::Display for SlideTableCellNumberFormatPath {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Cell {
                slide,
                table,
                position,
            } => write!(
                formatter,
                "slide {} table {} cell {}:{} Number format",
                slide.get(),
                table.get(),
                position.row(),
                position.column()
            ),
        }
    }
}

/// Failure from a Keynote table-cell Number-format read or transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTableCellNumberFormatError {
    #[error("this Keynote source does not support physical table-cell Number-format edits")]
    UnsupportedSource,
    #[error("the requested Keynote table-cell Number-format graph is outside the supported scope")]
    UnsupportedDependency,
    #[error("the requested Keynote table-cell Number-format topology is unsupported")]
    UnsupportedTopology,
    #[error("the Keynote table-cell Number-format selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    #[error("the selected Keynote table has no cell at position {position:?}")]
    CellPositionNotFound { position: CellPosition },
    #[error("the selected Keynote table is locked")]
    Locked,
    #[error("the selected Keynote table-cell Number format is not a Number format")]
    WrongFormatFamily,
    #[error("the selected Keynote table-cell Number-format source is invalid")]
    InvalidSource,
    #[error(
        "Keynote table-cell Number format {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTableCellNumberFormatLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error(
        "could not allocate {amount} units for the Keynote table-cell Number-format transaction"
    )]
    Allocation { amount: usize },
    #[error("the edited Keynote table-cell Number format failed semantic verification")]
    Verification,
    #[error("the Keynote table-cell Number-format patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable Number-format edit staged against an immutable package
/// snapshot.
pub struct SlideTableCellNumberFormatEdit<'a> {
    source: &'a Package,
    selection: CellSelection,
    after: Option<Number>,
}

impl fmt::Debug for SlideTableCellNumberFormatEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableCellNumberFormatEdit")
            .field("path", &self.selection.path)
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableCellNumberFormatEdit<'_> {
    #[must_use]
    pub const fn path(&self) -> SlideTableCellNumberFormatPath {
        self.selection.path
    }

    #[must_use]
    pub const fn before(&self) -> Option<&Number> {
        self.selection.before.as_ref()
    }

    #[must_use]
    pub const fn after(&self) -> Option<&Number> {
        self.after.as_ref()
    }

    #[must_use]
    pub const fn number(&self) -> Option<&Number> {
        self.after.as_ref()
    }

    #[must_use]
    pub const fn format(&self) -> Option<&Number> {
        self.after.as_ref()
    }

    #[must_use]
    pub fn set(mut self, value: Number) -> Self {
        self.after = Some(value);
        self
    }

    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self
    }

    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    pub fn commit(
        self,
    ) -> Result<SlideTableCellNumberFormatCommit, SlideTableCellNumberFormatError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible Number-format patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableCellNumberFormatPatch {
    artifacts: ExactArtifacts,
    path: SlideTableCellNumberFormatPath,
    before: Option<Number>,
    after: Option<Number>,
    touched_components: usize,
}

impl fmt::Debug for SlideTableCellNumberFormatPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableCellNumberFormatPatch")
            .field("path", &self.path)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableCellNumberFormatPatch {
    #[must_use]
    pub const fn path(&self) -> SlideTableCellNumberFormatPath {
        self.path
    }

    #[must_use]
    pub const fn before(&self) -> Option<&Number> {
        self.before.as_ref()
    }

    #[must_use]
    pub const fn after(&self) -> Option<&Number> {
        self.after.as_ref()
    }

    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after,
            after: self.before,
            touched_components: self.touched_components,
        }
    }
}

/// Compact diagnostics for one published Number-format transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableCellNumberFormatDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SlideTableCellNumberFormatDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one Number-format transaction.
#[must_use = "a Keynote Number-format commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableCellNumberFormatCommit {
    package: Package,
    patch: SlideTableCellNumberFormatPatch,
    diagnostics: SlideTableCellNumberFormatDiagnostics,
}

impl SlideTableCellNumberFormatCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &SlideTableCellNumberFormatPatch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableCellNumberFormatDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
struct CellSelection {
    path: SlideTableCellNumberFormatPath,
    target: core::Target,
    position: CellPosition,
    tile: core::ObjectLocation,
    tile_message_index: usize,
    format: core::ObjectLocation,
    format_message_index: usize,
    before: Option<Number>,
    old_format_key: Option<u32>,
    tile_references: Vec<(u32, u64)>,
    budget: core::Budget,
}

impl fmt::Debug for CellSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CellSelection")
            .field("path", &self.path)
            .field("before", &self.before)
            .field("locked", &self.target.locked)
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Read one existing slide-table cell's explicit Number format.
    ///
    /// `None` is the automatic/no-explicit-Number state. Other native format
    /// families fail with [`SlideTableCellNumberFormatError::WrongFormatFamily`].
    pub fn slide_table_cell_number_format<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        position: CellPosition,
    ) -> Result<Option<Number>, SlideTableCellNumberFormatError> {
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        Ok(select_cell(self, slide.into(), table.into(), position, &mut budget)?.before)
    }

    /// Begin an immutable exact edit of one existing slide-table cell's
    /// Number format.
    pub fn edit_slide_table_cell_number_format<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        position: CellPosition,
    ) -> Result<SlideTableCellNumberFormatEdit<'_>, SlideTableCellNumberFormatError> {
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        let selection = select_cell(self, slide.into(), table.into(), position, &mut budget)?;
        let after = selection.before;
        Ok(SlideTableCellNumberFormatEdit {
            source: self,
            selection: CellSelection {
                budget,
                ..selection
            },
            after,
        })
    }

    /// Apply an exact-source checked reversible Number-format patch.
    pub fn apply_slide_table_cell_number_format(
        &self,
        patch: &SlideTableCellNumberFormatPatch,
    ) -> Result<SlideTableCellNumberFormatCommit, SlideTableCellNumberFormatError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideTableCellNumberFormatError::PatchConflict);
        }
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        let current = select_path(self, patch.path, &mut budget)?;
        if current.before != patch.before {
            return Err(SlideTableCellNumberFormatError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideTableCellNumberFormatCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableCellNumberFormatDiagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch, budget)
    }
}

fn select_path(
    package: &Package,
    path: SlideTableCellNumberFormatPath,
    budget: &mut core::Budget,
) -> Result<CellSelection, SlideTableCellNumberFormatError> {
    let SlideTableCellNumberFormatPath::Cell {
        slide,
        table,
        position,
    } = path
    else {
        return Err(SlideTableCellNumberFormatError::PatchConflict);
    };
    select_cell(
        package,
        SlideSelector::position(slide),
        TableSelector::position(table),
        position,
        budget,
    )
}

fn select_cell(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    position: CellPosition,
    budget: &mut core::Budget,
) -> Result<CellSelection, SlideTableCellNumberFormatError> {
    let target = core::select_table(package, slide, table, budget).map_err(map_core_error)?;
    if position.row() >= target.rows || position.column() >= target.columns {
        return Err(SlideTableCellNumberFormatError::CellPositionNotFound { position });
    }

    let model_payload = core::model_payload(package, &target).map_err(map_core_error)?;
    let options = storage_options(package, budget).map_err(map_core_error)?;
    let (model, report) = storage_codec::decode_table_model_with_data_store_and_visitor(
        model_payload,
        options,
        &mut (),
    )
    .map_err(map_storage_error)?;
    charge_storage_report(budget, report).map_err(map_core_error)?;
    if model.model().number_of_rows() != target.rows
        || model.model().number_of_columns() != target.columns
    {
        return Err(SlideTableCellNumberFormatError::InvalidSource);
    }
    let store = model.data_store();
    let (tile_storage, tile_references) = decode_tile_references(store.tiles(), package, budget)?;
    let tile_size = tile_storage
        .tile_size()
        .filter(|size| *size != 0)
        .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
    let tile_id = position.row() / tile_size;
    let mut tile_matches = tile_references.iter().filter(|(id, _)| *id == tile_id);
    let tile_identifier = tile_matches
        .next()
        .map(|(_, identifier)| *identifier)
        .ok_or(SlideTableCellNumberFormatError::CellPositionNotFound { position })?;
    if tile_matches.next().is_some() || tile_identifier == target.model.identifier {
        return Err(SlideTableCellNumberFormatError::InvalidSource);
    }
    let tile = core::locate_object(package, tile_identifier).map_err(map_core_error)?;
    let tile_object = core::object_at(package, &tile).map_err(map_core_error)?;
    let (tile_message_index, tile_payload) =
        core::unique_message(tile_object, TILE_MESSAGE_TYPE, budget).map_err(map_core_error)?;
    let cell_source = tile_cell(tile_payload, position.row(), position.column())
        .map_err(|error| map_cell_wire_error(error, position))?;
    charge_bnc_parse(budget, cell_source.len()).map_err(map_core_error)?;
    let cell =
        BncCell::parse(cell_source).map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?;
    let old_format_key = classify_cell(&cell)?;

    // Number BNC identifiers in current Keynote files resolve through the
    // post-BNC format table (field 22). Older producers may only expose the
    // pre-BNC route (field 11), so retain that as a compatibility fallback.
    let format_identifier = store
        .format_table()
        .or_else(|| Some(store.format_table_pre_bnc()))
        .map(|reference| reference.identifier())
        .filter(|identifier| *identifier != 0)
        .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
    let format = core::locate_object(package, format_identifier).map_err(map_core_error)?;
    if format.identifier == target.model.identifier || format.identifier == tile_identifier {
        return Err(SlideTableCellNumberFormatError::UnsupportedDependency);
    }
    let format_object = core::object_at(package, &format).map_err(map_core_error)?;
    let (format_message_index, format_payload) = core::unique_message_any(
        format_object,
        &[
            TABLE_DATA_LIST_MESSAGE_TYPE,
            TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE,
        ],
        budget,
    )
    .map_err(map_core_error)?;
    let facts = decode_format_facts(format_payload, package, budget)?;
    validate_format_refcounts(package, &tile_references, &facts, tile_size, budget)?;
    let before = match old_format_key {
        Some(key) => {
            let entry = facts
                .entries
                .iter()
                .find(|entry| entry.key == key)
                .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
            if entry.ref_count == 0 {
                return Err(SlideTableCellNumberFormatError::InvalidSource);
            }
            let format_type = format_type(&entry.payload, package, budget)?;
            if format_type != number_codec::NATIVE_NUMBER_FORMAT_TYPE {
                return Err(SlideTableCellNumberFormatError::WrongFormatFamily);
            }
            let number = decode_number(&entry.payload, package, budget)?;
            if cell.explicit_format_flags() == 0 {
                None
            } else {
                Some(number)
            }
        },
        None => None,
    };
    Ok(CellSelection {
        path: SlideTableCellNumberFormatPath::Cell {
            slide: target.slide_position,
            table: target.table_position,
            position,
        },
        target,
        position,
        tile,
        tile_message_index,
        format,
        format_message_index,
        before,
        old_format_key,
        tile_references,
        budget: *budget,
    })
}

#[derive(Default)]
struct TileReferenceCollector {
    references: Vec<(u32, u64)>,
    tile_ids: HashSet<u32>,
    identifiers: HashSet<u64>,
}

impl storage_codec::StorageVisitor for TileReferenceCollector {
    fn visit_tile_reference(
        &mut self,
        record: storage_codec::TileReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.references
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.tile_ids
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.identifiers
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        if !self.tile_ids.insert(record.tile_id())
            || !self.identifiers.insert(record.reference().identifier())
        {
            return Err(storage_codec::DecodeError::invalid_visitor_result());
        }
        self.references
            .push((record.tile_id(), record.reference().identifier()));
        Ok(())
    }
}

fn decode_tile_references(
    source: &[u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<(storage_codec::TileStorageSnapshot, Vec<(u32, u64)>), SlideTableCellNumberFormatError>
{
    budget.allocations(1).map_err(map_core_error)?;
    let options = storage_options(package, budget).map_err(map_core_error)?;
    let mut collector = TileReferenceCollector::default();
    let (snapshot, report) =
        storage_codec::decode_tile_storage_with_visitor(source, options, &mut collector)
            .map_err(map_storage_error)?;
    charge_storage_report(budget, report).map_err(map_core_error)?;
    let retained = collector
        .references
        .len()
        .checked_mul(size_of::<(u32, u64)>() + 4 * size_of::<usize>())
        .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
    budget
        .allocations(
            collector
                .references
                .len()
                .checked_mul(3)
                .ok_or(core::Error::InvalidSource)
                .map_err(map_core_error)?,
        )
        .and_then(|_| budget.retained(retained))
        .and_then(|_| budget.work(collector.references.len()))
        .map_err(map_core_error)?;
    if collector.references.is_empty()
        || collector
            .references
            .iter()
            .any(|(_, identifier)| *identifier == 0)
    {
        return Err(SlideTableCellNumberFormatError::InvalidSource);
    }
    Ok((snapshot, collector.references))
}

struct FormatEntry {
    key: u32,
    ref_count: u32,
    payload: Vec<u8>,
}

struct FormatFacts {
    next_list_id: u32,
    entries: Vec<FormatEntry>,
    keys: HashSet<u32>,
}

#[derive(Default)]
struct FormatCollector {
    entries: Vec<FormatEntry>,
    keys: HashSet<u32>,
    segments: usize,
}

impl storage_codec::StorageVisitor for FormatCollector {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        let snapshot = record.snapshot();
        let Some(payload) = snapshot.format() else {
            return Err(storage_codec::DecodeError::invalid_visitor_result());
        };
        if snapshot.string_value().is_some()
            || snapshot.reference().is_some()
            || snapshot.formula().is_some()
            || snapshot.custom_format().is_some()
            || snapshot.rich_text_payload().is_some()
            || snapshot.comment_storage().is_some()
            || snapshot.import_warning_set().is_some()
            || snapshot.cell_spec().is_some()
            || snapshot.key() == 0
            || snapshot.ref_count() == 0
            || payload.is_empty()
        {
            return Err(storage_codec::DecodeError::invalid_visitor_result());
        }
        self.entries
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.keys
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        if !self.keys.insert(snapshot.key()) {
            return Err(storage_codec::DecodeError::invalid_visitor_result());
        }
        let mut owned_payload = Vec::new();
        owned_payload
            .try_reserve_exact(payload.len())
            .map_err(|_| storage_codec::DecodeError::allocation(payload.len()))?;
        owned_payload.extend_from_slice(payload);
        self.entries.push(FormatEntry {
            key: snapshot.key(),
            ref_count: snapshot.ref_count(),
            payload: owned_payload,
        });
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.segments = self
            .segments
            .checked_add(1)
            .ok_or_else(storage_codec::DecodeError::invalid_visitor_result)?;
        Ok(())
    }
}

fn decode_format_facts(
    source: &[u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<FormatFacts, SlideTableCellNumberFormatError> {
    let options = storage_options(package, budget).map_err(map_core_error)?;
    let mut collector = FormatCollector::default();
    let (snapshot, report) =
        storage_codec::decode_table_data_list_with_visitor(source, options, &mut collector)
            .map_err(map_storage_error)?;
    charge_storage_report(budget, report).map_err(map_core_error)?;
    if snapshot.list_type() != FORMAT_LIST_TYPE || collector.segments != 0 {
        return Err(SlideTableCellNumberFormatError::UnsupportedDependency);
    }
    if snapshot.next_list_id() == 0
        || collector
            .entries
            .iter()
            .any(|entry| entry.key >= snapshot.next_list_id())
    {
        return Err(SlideTableCellNumberFormatError::InvalidSource);
    }
    budget
        .retained(
            collector
                .entries
                .len()
                .saturating_mul(size_of::<FormatEntry>()),
        )
        .and_then(|_| {
            budget.retained(
                collector
                    .keys
                    .len()
                    .saturating_mul(size_of::<u32>() + 4 * size_of::<usize>()),
            )
        })
        .and_then(|_| budget.allocations(collector.entries.len().saturating_mul(2)))
        .and_then(|_| budget.work(collector.entries.len()))
        .map_err(map_core_error)?;
    Ok(FormatFacts {
        next_list_id: snapshot.next_list_id(),
        entries: collector.entries,
        keys: collector.keys,
    })
}

fn classify_cell(cell: &BncCell) -> Result<Option<u32>, SlideTableCellNumberFormatError> {
    if cell.control_cell_spec_identifier().is_some() || cell.secondary_format_identifier().is_some()
    {
        return Err(SlideTableCellNumberFormatError::WrongFormatFamily);
    }
    match cell.cell_format_kind() {
        Some(litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND) => {
            if !cell.has_only_decimal_format_metadata()
                || (cell.explicit_format_flags() != 0
                    && cell.explicit_format_flags() != litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT)
            {
                return Err(SlideTableCellNumberFormatError::WrongFormatFamily);
            }
            let key = cell
                .format_identifier()
                .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
            if key == 0 {
                return Err(SlideTableCellNumberFormatError::InvalidSource);
            }
            Ok(Some(key))
        },
        Some(_) => Err(SlideTableCellNumberFormatError::WrongFormatFamily),
        None if cell.format_identifier().is_some() || cell.explicit_format_flags() != 0 => {
            Err(SlideTableCellNumberFormatError::WrongFormatFamily)
        },
        None => Ok(None),
    }
}

fn format_type(
    source: &[u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<u32, SlideTableCellNumberFormatError> {
    let limits = budget.residual(package).map_err(map_core_error)?;
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?;
    let mut selected = None;
    for field in view.fields() {
        budget.fields(1).map_err(map_core_error)?;
        budget.work(field.raw().len()).map_err(map_core_error)?;
        if field.number() != 1 {
            continue;
        }
        if selected.is_some() || field.wire_type() != 0 {
            return Err(SlideTableCellNumberFormatError::InvalidSource);
        }
        let (value, consumed) = decode_varint_from_bytes(field.payload())
            .map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?;
        if consumed != field.payload().len() {
            return Err(SlideTableCellNumberFormatError::InvalidSource);
        }
        selected =
            Some(u32::try_from(value).map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?);
    }
    selected.ok_or(SlideTableCellNumberFormatError::InvalidSource)
}

fn number_options(
    source: &[u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<number_codec::DecodeOptions, SlideTableCellNumberFormatError> {
    let limits = budget.residual(package).map_err(map_core_error)?;
    let output = source
        .len()
        .max(1)
        .saturating_mul(4)
        .min(limits.max_output_bytes());
    Ok(number_codec::DecodeOptions::for_source(source).with_max_output_bytes(output.max(1)))
}

fn decode_number(
    source: &[u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Number, SlideTableCellNumberFormatError> {
    let options = number_options(source, package, budget)?;
    let (snapshot, report) = number_codec::decode_number_format_with_report(source, options)
        .map_err(map_number_codec_error)?;
    charge_number_report(budget, report).map_err(map_core_error)?;
    let places = if snapshot.decimal_places() == number_codec::NATIVE_AUTOMATIC_DECIMAL_PLACES {
        DecimalPlaces::Automatic
    } else {
        DecimalPlaces::Fixed(
            FixedDecimalPlaces::new(
                u8::try_from(snapshot.decimal_places())
                    .map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?,
            )
            .map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?,
        )
    };
    let negative = match snapshot.negative_style() {
        0 => NegativeStyle::MinusSign,
        1 => NegativeStyle::Red,
        2 => NegativeStyle::Parentheses,
        3 => NegativeStyle::RedParentheses,
        _ => return Err(SlideTableCellNumberFormatError::InvalidSource),
    };
    let separator = if snapshot.show_thousands_separator() {
        ThousandsSeparator::Shown
    } else {
        ThousandsSeparator::Hidden
    };
    Ok(Number::new(places, negative, separator))
}

fn number_write(value: Number) -> number_codec::NumberFormatWrite {
    let decimal_places = match value.decimal_places() {
        DecimalPlaces::Automatic => number_codec::NATIVE_AUTOMATIC_DECIMAL_PLACES,
        DecimalPlaces::Fixed(value) => u32::from(value.value()),
    };
    let negative_style = match value.negative_style() {
        NegativeStyle::MinusSign => 0,
        NegativeStyle::Red => 1,
        NegativeStyle::Parentheses => 2,
        NegativeStyle::RedParentheses => 3,
    };
    number_codec::NumberFormatWrite::new(
        decimal_places,
        negative_style,
        matches!(value.thousands_separator(), ThousandsSeparator::Shown),
    )
}

#[derive(Default)]
struct BncReferenceCensus {
    formats: HashMap<u32, usize>,
    invalid: bool,
}

impl BncReferenceCensus {
    fn increment(&mut self, identifier: u32) -> bool {
        if identifier == 0 {
            return false;
        }
        if let Some(count) = self.formats.get_mut(&identifier) {
            let Some(next) = (*count).checked_add(1) else {
                return false;
            };
            *count = next;
        } else {
            if self.formats.try_reserve(1).is_err() {
                return false;
            }
            self.formats.insert(identifier, 1);
        }
        true
    }

    fn count_cell(&mut self, source: &[u8]) -> bool {
        let Ok(cell) = BncCell::parse(source) else {
            return false;
        };
        let Some(primary) = cell.format_identifier() else {
            return cell.secondary_format_identifier().is_none()
                && cell.control_cell_spec_identifier().is_none();
        };
        if !self.increment(primary) {
            return false;
        }
        cell.secondary_format_identifier()
            .is_none_or(|identifier| self.increment(identifier))
    }
}

impl storage_codec::StorageVisitor for BncReferenceCensus {
    fn visit_tile_row(
        &mut self,
        row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        if self.invalid {
            return Ok(());
        }
        let storage = row
            .cell_storage_buffer()
            .unwrap_or_else(|| row.cell_storage_buffer_pre_bnc());
        let offsets = row
            .cell_offsets()
            .unwrap_or_else(|| row.cell_offsets_pre_bnc());
        if offsets.is_empty() {
            match row.cell_count() {
                0 => return Ok(()),
                1 if self.count_cell(storage) => return Ok(()),
                _ => {
                    self.invalid = true;
                    return Ok(());
                },
            }
        }
        if !offsets.len().is_multiple_of(2) {
            self.invalid = true;
            return Ok(());
        }
        let slot_count = offsets.len() / 2;
        let expected = usize::try_from(row.cell_count()).unwrap_or(usize::MAX);
        if expected > slot_count {
            self.invalid = true;
            return Ok(());
        }
        let unit = if row.has_wide_offsets().unwrap_or(false) {
            4usize
        } else {
            1usize
        };
        let mut occupied = 0usize;
        let mut previous = None;
        for encoded in offsets.chunks_exact(2) {
            let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
            if raw == u16::MAX {
                continue;
            }
            let Some(start) = usize::from(raw).checked_mul(unit) else {
                self.invalid = true;
                return Ok(());
            };
            if start > storage.len() || previous.is_some_and(|prior| prior > start) {
                self.invalid = true;
                return Ok(());
            }
            if let Some(prior) = previous {
                if !self.count_cell(&storage[prior..start]) {
                    self.invalid = true;
                    return Ok(());
                }
            }
            let Some(next) = occupied.checked_add(1) else {
                self.invalid = true;
                return Ok(());
            };
            occupied = next;
            previous = Some(start);
        }
        if let Some(start) = previous {
            if !self.count_cell(&storage[start..]) {
                self.invalid = true;
                return Ok(());
            }
        }
        if occupied != expected {
            self.invalid = true;
        }
        Ok(())
    }
}

fn validate_format_refcounts(
    package: &Package,
    tile_references: &[(u32, u64)],
    facts: &FormatFacts,
    _tile_size: u32,
    budget: &mut core::Budget,
) -> Result<(), SlideTableCellNumberFormatError> {
    let mut census = BncReferenceCensus::default();
    for (_, identifier) in tile_references {
        let location = core::locate_object(package, *identifier).map_err(map_core_error)?;
        let object = core::object_at(package, &location).map_err(map_core_error)?;
        let (_, payload) =
            core::unique_message(object, TILE_MESSAGE_TYPE, budget).map_err(map_core_error)?;
        let options = storage_options(package, budget).map_err(map_core_error)?;
        let (_, report) = storage_codec::decode_tile_with_visitor(payload, options, &mut census)
            .map_err(map_storage_error)?;
        charge_storage_report(budget, report).map_err(map_core_error)?;
        if census.invalid {
            return Err(SlideTableCellNumberFormatError::InvalidSource);
        }
    }
    for entry in &facts.entries {
        let observed = census.formats.get(&entry.key).copied().unwrap_or(0);
        if observed != usize::try_from(entry.ref_count).unwrap_or(usize::MAX) {
            return Err(SlideTableCellNumberFormatError::InvalidSource);
        }
    }
    if census.formats.keys().any(|key| !facts.keys.contains(key)) {
        return Err(SlideTableCellNumberFormatError::InvalidSource);
    }
    let retained = census
        .formats
        .len()
        .checked_mul(size_of::<(u32, usize)>() + 4 * size_of::<usize>())
        .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
    budget
        .allocations(census.formats.len())
        .and_then(|_| budget.retained(retained))
        .and_then(|_| budget.work(census.formats.len()))
        .map_err(map_core_error)?;
    Ok(())
}

fn storage_options(
    package: &Package,
    budget: &core::Budget,
) -> Result<storage_codec::DecodeOptions, core::Error> {
    let limits = budget.residual(package)?;
    let references = budget.remaining_references()?;
    Ok(storage_codec::DecodeOptions::new(
        limits.max_input_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        references,
        limits.max_input_bytes(),
    ))
}

fn charge_storage_report(
    budget: &mut core::Budget,
    report: storage_codec::DecodeReport,
) -> Result<(), core::Error> {
    budget.input(report.source_bytes())?;
    budget.fields(report.fields())?;
    budget.work(report.work_bytes())?;
    budget.references(report.references())?;
    budget.nesting(report.max_depth() as usize)
}

fn charge_bnc_parse(budget: &mut core::Budget, bytes: usize) -> Result<(), core::Error> {
    budget.input(bytes)?;
    budget.work(bytes)?;
    budget.allocations(1)?;
    budget.retained(bytes)
}

fn charge_tile_patch(
    budget: &mut core::Budget,
    tile_bytes: usize,
    replacement_bytes: usize,
) -> Result<(), core::Error> {
    let candidate = tile_bytes
        .checked_add(replacement_bytes)
        .and_then(|bytes| bytes.checked_add(64))
        .ok_or(core::Error::InvalidSource)?;
    let scratch = candidate.checked_mul(3).ok_or(core::Error::InvalidSource)?;
    budget.allocations(5)?;
    budget.scratch(scratch)?;
    budget.retained(candidate)?;
    budget.work(
        candidate
            .checked_add(scratch)
            .ok_or(core::Error::InvalidSource)?,
    )
}

fn charge_number_report(
    budget: &mut core::Budget,
    report: number_codec::DecodeReport,
) -> Result<(), core::Error> {
    budget.input(report.input_bytes())?;
    budget.output(report.output_bytes())?;
    budget.fields(report.fields())?;
    budget.work(report.work_bytes())?;
    budget.references(report.references())?;
    budget.allocations(report.allocations())?;
    budget.retained(report.retained_bytes())?;
    budget.scratch(report.scratch_bytes())?;
    budget.nesting(report.max_depth() as usize)
}

fn charge_number_requirements(
    budget: &mut core::Budget,
    requirements: number_codec::RewriteExecutionRequirements,
) -> Result<(), core::Error> {
    budget.output(requirements.output_bytes())?;
    budget.fields(requirements.fields())?;
    budget.work(requirements.work_bytes())?;
    budget.references(requirements.references())?;
    budget.allocations(requirements.allocations())?;
    budget.retained(requirements.retained_bytes())?;
    budget.scratch(requirements.scratch_bytes())?;
    budget.nesting(requirements.max_depth() as usize)
}

fn prepare_number_append(
    value: Number,
    source: &[u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Vec<u8>, SlideTableCellNumberFormatError> {
    let options = number_options(source, package, budget)?;
    let prepared = number_codec::prepare_number_format_write(number_write(value), options)
        .map_err(map_number_codec_error)?;
    let requirements = prepared.execution_requirements();
    charge_number_requirements(budget, requirements).map_err(map_core_error)?;
    let output = prepared
        .execute(number_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(map_number_codec_error)?;
    if output.report().output_bytes() != requirements.output_bytes()
        || output.report().fields() != requirements.fields()
        || output.report().work_bytes() != requirements.work_bytes()
        || output.report().allocations() != requirements.allocations()
        || output.report().retained_bytes() != requirements.retained_bytes()
    {
        return Err(SlideTableCellNumberFormatError::Verification);
    }
    Ok(output.into_bytes())
}

fn find_matching_format(
    facts: &FormatFacts,
    desired: Number,
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Option<u32>, SlideTableCellNumberFormatError> {
    for entry in &facts.entries {
        if format_type(&entry.payload, package, budget)? != number_codec::NATIVE_NUMBER_FORMAT_TYPE
        {
            continue;
        }
        if decode_number(&entry.payload, package, budget)? == desired {
            return Ok(Some(entry.key));
        }
    }
    Ok(None)
}

fn validate_current_entry(
    facts: &FormatFacts,
    old_key: Option<u32>,
    cell: &BncCell,
    before: Option<Number>,
    package: &Package,
    budget: &mut core::Budget,
) -> Result<(), SlideTableCellNumberFormatError> {
    let Some(key) = old_key else {
        if before.is_some() {
            return Err(SlideTableCellNumberFormatError::PatchConflict);
        }
        return Ok(());
    };
    let entry = facts
        .entries
        .iter()
        .find(|entry| entry.key == key)
        .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
    if entry.ref_count == 0
        || format_type(&entry.payload, package, budget)? != number_codec::NATIVE_NUMBER_FORMAT_TYPE
    {
        return Err(SlideTableCellNumberFormatError::WrongFormatFamily);
    }
    let current = decode_number(&entry.payload, package, budget)?;
    if cell.explicit_format_flags() == litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT {
        if before != Some(current) {
            return Err(SlideTableCellNumberFormatError::PatchConflict);
        }
    } else if before.is_some() {
        return Err(SlideTableCellNumberFormatError::PatchConflict);
    }
    Ok(())
}

fn mutate_format_list(
    source: &[u8],
    facts: &FormatFacts,
    old_key: Option<u32>,
    desired_key: Option<u32>,
    desired_payload: Option<&[u8]>,
    package: &Package,
    budget: &mut core::Budget,
) -> Result<(Option<Vec<u8>>, Option<u32>), SlideTableCellNumberFormatError> {
    if old_key == desired_key {
        return Ok((None, desired_key));
    }
    let mut output: Option<Vec<u8>> = None;
    if let Some(old_key) = old_key {
        let entry = facts
            .entries
            .iter()
            .find(|entry| entry.key == old_key)
            .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
        let mutation = if entry.ref_count > 1 {
            storage_codec::TableDataListEntryMutation::RefCount(
                storage_codec::TableDataListEntryRefCountEdit::new(
                    old_key,
                    entry.ref_count,
                    entry.ref_count - 1,
                ),
            )
        } else {
            storage_codec::TableDataListEntryMutation::Remove(
                storage_codec::TableDataListEntryRemovalSpec::format(
                    old_key,
                    entry.ref_count,
                    &entry.payload,
                ),
            )
        };
        output = Some(apply_list_mutation(source, mutation, package, budget)?);
    }
    let current_source = output.as_deref().unwrap_or(source);
    let Some(desired_key) = desired_key else {
        if desired_payload.is_some() {
            return Err(SlideTableCellNumberFormatError::InvalidSource);
        }
        return Ok((output, None));
    };
    let existing = facts.entries.iter().find(|entry| entry.key == desired_key);
    let next = if let Some(entry) = existing {
        let replacement = entry
            .ref_count
            .checked_add(1)
            .ok_or(SlideTableCellNumberFormatError::UnsupportedDependency)?;
        storage_codec::TableDataListEntryMutation::RefCount(
            storage_codec::TableDataListEntryRefCountEdit::new(
                desired_key,
                entry.ref_count,
                replacement,
            ),
        )
    } else {
        let payload = desired_payload.ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
        if desired_key != facts.next_list_id {
            return Err(SlideTableCellNumberFormatError::UnsupportedDependency);
        }
        let replacement = desired_key
            .checked_add(1)
            .ok_or(SlideTableCellNumberFormatError::UnsupportedDependency)?;
        storage_codec::TableDataListEntryMutation::Append(
            storage_codec::TableDataListEntryAppend::format(desired_key, 1, payload)
                .advance_next_list_id(desired_key, replacement),
        )
    };
    let rewritten = apply_list_mutation(current_source, next, package, budget)?;
    Ok((Some(rewritten), Some(desired_key)))
}

fn apply_list_mutation(
    source: &[u8],
    mutation: storage_codec::TableDataListEntryMutation<'_>,
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Vec<u8>, SlideTableCellNumberFormatError> {
    let options = storage_options(package, budget).map_err(map_core_error)?;
    let prepared = storage_codec::prepare_table_data_list_entry_rewrite(source, mutation, options)
        .map_err(map_storage_error)?;
    let preparation = prepared.prepare_report();
    budget
        .input(preparation.source_bytes())
        .map_err(map_core_error)?;
    let requirements = prepared.requirements();
    budget
        .fields(requirements.fields())
        .map_err(map_core_error)?;
    budget
        .work(requirements.work_bytes())
        .map_err(map_core_error)?;
    budget
        .references(requirements.references())
        .map_err(map_core_error)?;
    budget
        .nesting(requirements.max_depth() as usize)
        .map_err(map_core_error)?;
    budget
        .allocations(requirements.allocations())
        .map_err(map_core_error)?;
    budget
        .retained(requirements.retained_bytes())
        .map_err(map_core_error)?;
    budget
        .scratch(requirements.scratch_bytes())
        .map_err(map_core_error)?;
    budget
        .output(requirements.output_bytes())
        .map_err(map_core_error)?;
    let (output, report) = prepared
        .execute(requirements.exact_limits())
        .map_err(map_storage_error)?;
    if report.output_bytes() != requirements.output_bytes()
        || report.fields() != requirements.fields()
        || report.work_bytes() != requirements.work_bytes()
        || report.references() != requirements.references()
        || report.allocations() != requirements.allocations()
        || report.retained_bytes() != requirements.retained_bytes()
    {
        return Err(SlideTableCellNumberFormatError::Verification);
    }
    Ok(output)
}

fn commit_edit(
    source: &Package,
    selection: &CellSelection,
    after: Option<Number>,
) -> Result<SlideTableCellNumberFormatCommit, SlideTableCellNumberFormatError> {
    let mut budget = selection.budget;
    if selection.before == after {
        let bytes = shared_source_artifact(source, &mut budget)?;
        return Ok(SlideTableCellNumberFormatCommit {
            package: source.snapshot(),
            patch: SlideTableCellNumberFormatPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                path: selection.path,
                before: selection.before,
                after,
                touched_components: 0,
            },
            diagnostics: SlideTableCellNumberFormatDiagnostics::unchanged(),
        });
    }
    if selection.target.locked {
        return Err(SlideTableCellNumberFormatError::Locked);
    }
    let (candidate, touched_components) =
        rewrite(source, selection.path, selection.before, after, &mut budget)?;
    let mut reopen_budget = budget;
    let reopened = select_path(&candidate, selection.path, &mut reopen_budget)?;
    if !same_target(&reopened, selection) || reopened.before != after {
        return Err(SlideTableCellNumberFormatError::Verification);
    }
    verify_package_locality(source, &candidate, selection, &reopened, &mut reopen_budget)?;
    let source_bytes = shared_source_artifact(source, &mut reopen_budget)?;
    let target_bytes = shared_source_artifact(&candidate, &mut reopen_budget)?;
    Ok(SlideTableCellNumberFormatCommit {
        package: candidate,
        patch: SlideTableCellNumberFormatPatch {
            artifacts: ExactArtifacts::new(source_bytes, target_bytes),
            path: selection.path,
            before: selection.before,
            after,
            touched_components,
        },
        diagnostics: SlideTableCellNumberFormatDiagnostics::published(touched_components),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &SlideTableCellNumberFormatPatch,
    mut budget: core::Budget,
) -> Result<SlideTableCellNumberFormatCommit, SlideTableCellNumberFormatError> {
    let source_selection = select_path(source, patch.path, &mut budget)?;
    if source_selection.before != patch.before {
        return Err(SlideTableCellNumberFormatError::PatchConflict);
    }
    let candidate = parse_candidate(patch.artifacts.target(), source, &mut budget)?;
    let candidate_selection = select_path(&candidate, patch.path, &mut budget)?;
    if !same_target(&candidate_selection, &source_selection)
        || candidate_selection.before != patch.after
    {
        return Err(SlideTableCellNumberFormatError::Verification);
    }
    verify_package_locality(
        source,
        &candidate,
        &source_selection,
        &candidate_selection,
        &mut budget,
    )?;
    Ok(SlideTableCellNumberFormatCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableCellNumberFormatDiagnostics::published(patch.touched_components),
    })
}

fn rewrite(
    source: &Package,
    path: SlideTableCellNumberFormatPath,
    before: Option<Number>,
    after: Option<Number>,
    budget: &mut core::Budget,
) -> Result<(Package, usize), SlideTableCellNumberFormatError> {
    let selection = select_path(source, path, budget)?;
    if selection.before != before {
        return Err(SlideTableCellNumberFormatError::PatchConflict);
    }
    if selection.target.locked {
        return Err(SlideTableCellNumberFormatError::Locked);
    }
    let tile_object = core::object_at(source, &selection.tile).map_err(map_core_error)?;
    let (_, tile_payload) =
        core::unique_message(tile_object, TILE_MESSAGE_TYPE, budget).map_err(map_core_error)?;
    let cell_source = tile_cell(
        tile_payload,
        selection.position.row(),
        selection.position.column(),
    )
    .map_err(|error| map_cell_wire_error(error, selection.position))?;
    charge_bnc_parse(budget, cell_source.len()).map_err(map_core_error)?;
    let cell =
        BncCell::parse(cell_source).map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?;
    if classify_cell(&cell)? != selection.old_format_key {
        return Err(SlideTableCellNumberFormatError::PatchConflict);
    }

    let format_object = core::object_at(source, &selection.format).map_err(map_core_error)?;
    let (_, format_payload) = core::unique_message_any(
        format_object,
        &[
            TABLE_DATA_LIST_MESSAGE_TYPE,
            TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE,
        ],
        budget,
    )
    .map_err(map_core_error)?;
    let facts = decode_format_facts(format_payload, source, budget)?;
    validate_format_refcounts(source, &selection.tile_references, &facts, 0, budget)?;
    validate_current_entry(
        &facts,
        selection.old_format_key,
        &cell,
        before,
        source,
        budget,
    )?;

    let desired_key = if let Some(value) = after {
        find_matching_format(&facts, value, source, budget)?.or(Some(facts.next_list_id))
    } else {
        None
    };
    if desired_key == Some(facts.next_list_id) {
        facts
            .next_list_id
            .checked_add(1)
            .ok_or(SlideTableCellNumberFormatError::UnsupportedDependency)?;
    }
    let desired_payload = match (after, desired_key == Some(facts.next_list_id)) {
        (Some(value), true) => Some(prepare_number_append(
            value,
            format_payload,
            source,
            budget,
        )?),
        _ => None,
    };
    let (format_replacement, selected_key) = mutate_format_list(
        format_payload,
        &facts,
        selection.old_format_key,
        desired_key,
        desired_payload.as_deref(),
        source,
        budget,
    )?;
    if after.is_some() && selected_key.is_none() {
        return Err(SlideTableCellNumberFormatError::InvalidSource);
    }

    let mut replacement_cell =
        BncCell::parse(cell_source).map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?;
    replacement_cell
        .set_number_or_percentage_format_identifier_preserving_value(selected_key)
        .map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?;
    let replacement_cell = replacement_cell.encode();
    if replacement_cell == cell_source && format_replacement.is_none() {
        return Err(SlideTableCellNumberFormatError::Verification);
    }
    charge_tile_patch(budget, tile_payload.len(), replacement_cell.len())
        .map_err(map_core_error)?;
    let patched_tile = patch_tile_cell(
        tile_payload,
        selection.position.row(),
        selection.position.column(),
        &replacement_cell,
    )
    .map_err(|error| map_cell_wire_error(error, selection.position))?;

    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?;
    let mut component_names = vec![
        selection.tile.component.as_ref(),
        selection.format.component.as_ref(),
    ];
    component_names.sort_unstable();
    component_names.dedup();
    let mut archives = Vec::new();
    archives
        .try_reserve_exact(component_names.len())
        .map_err(|_| SlideTableCellNumberFormatError::Allocation {
            amount: component_names.len(),
        })?;
    for component_name in &component_names {
        let component = core::component(source, component_name).map_err(map_core_error)?;
        let encoded = component
            .archive()
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_native_error)?;
        budget.allocations(1).map_err(map_core_error)?;
        budget.retained(encoded).map_err(map_core_error)?;
        budget.scratch(encoded).map_err(map_core_error)?;
        budget.work(encoded).map_err(map_core_error)?;
        archives.push(((*component_name).to_owned(), component.archive().clone()));
    }
    replace_archive_message(
        &mut archives,
        selection.tile.component.as_ref(),
        selection.tile.identifier,
        selection.tile_message_index,
        patched_tile,
        archive_limits,
    )?;
    if let Some(format_replacement) = format_replacement {
        replace_archive_message(
            &mut archives,
            selection.format.component.as_ref(),
            selection.format.identifier,
            selection.format_message_index,
            format_replacement,
            archive_limits,
        )?;
    }

    let mut compressed = Vec::new();
    compressed.try_reserve_exact(archives.len()).map_err(|_| {
        SlideTableCellNumberFormatError::Allocation {
            amount: archives.len(),
        }
    })?;
    for (name, archive) in &archives {
        let encoded_bound = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_native_error)?;
        let compressed_bound =
            SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_core_native_error)?;
        budget
            .output(
                encoded_bound
                    .checked_add(compressed_bound)
                    .ok_or(core::Error::InvalidSource)
                    .map_err(map_core_error)?,
            )
            .map_err(map_core_error)?;
        budget.allocations(2).map_err(map_core_error)?;
        budget.retained(encoded_bound).map_err(map_core_error)?;
        budget.retained(compressed_bound).map_err(map_core_error)?;
        budget.scratch(encoded_bound).map_err(map_core_error)?;
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_native_error)?;
        if bytes.len() != encoded_bound {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
        let bytes = SnappyStream::compress(&bytes).map_err(map_core_native_error)?;
        if bytes.len() > compressed_bound {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
        compressed.push((name.as_str(), bytes));
    }

    let edits = compressed
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes.as_slice()))
        .collect::<Vec<_>>();
    let catalog = physical_catalog(source)?;
    budget
        .preflight_reassembly(
            catalog,
            compressed.iter().map(|(_, bytes)| bytes.len()).sum(),
            0,
        )
        .map_err(map_core_error)?;
    let prepared = catalog
        .prepare_reassembly_with_deletions(&edits, &[], source.state.options.archive())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements).map_err(map_core_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate = parse_candidate(output.into(), source, budget)?;
    Ok((candidate, component_names.len()))
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideTableCellNumberFormatError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTableCellNumberFormatError::UnsupportedSource),
    }
}

fn shared_source_artifact(
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Arc<[u8]>, SlideTableCellNumberFormatError> {
    let source = physical_catalog(package)?;
    budget
        .artifact(source.package().source_bytes().len(), 0)
        .map_err(map_core_error)?;
    Ok(source.shared_source())
}

fn parse_candidate(
    source: Arc<[u8]>,
    original: &Package,
    budget: &mut core::Budget,
) -> Result<Package, SlideTableCellNumberFormatError> {
    budget
        .preflight_candidate(original, source.len())
        .map_err(map_core_error)?;
    budget
        .preflight_semantic_scan(original)
        .map_err(map_core_error)?;
    let candidate = Package::from_source_with_options(source, original.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    Ok(candidate)
}

fn same_target(left: &CellSelection, right: &CellSelection) -> bool {
    core::same_target(&left.target, &right.target)
        && left.position == right.position
        && left.tile == right.tile
        && left.tile_message_index == right.tile_message_index
        && left.format == right.format
        && left.format_message_index == right.format_message_index
}

fn map_cell_wire_error(
    error: CellWireError,
    position: CellPosition,
) -> SlideTableCellNumberFormatError {
    match error {
        CellWireError::Missing => {
            SlideTableCellNumberFormatError::CellPositionNotFound { position }
        },
        CellWireError::Unsupported => SlideTableCellNumberFormatError::UnsupportedDependency,
        CellWireError::Invalid => SlideTableCellNumberFormatError::InvalidSource,
    }
}

fn map_core_error(error: core::Error) -> SlideTableCellNumberFormatError {
    match error {
        core::Error::UnsupportedSource => SlideTableCellNumberFormatError::UnsupportedSource,
        core::Error::UnsupportedDependency => {
            SlideTableCellNumberFormatError::UnsupportedDependency
        },
        core::Error::UnsupportedTopology => SlideTableCellNumberFormatError::UnsupportedTopology,
        core::Error::AmbiguousSelector => SlideTableCellNumberFormatError::AmbiguousSelector,
        core::Error::EmptySlideName => SlideTableCellNumberFormatError::EmptySlideName,
        core::Error::SlideNameNotFound => SlideTableCellNumberFormatError::SlideNameNotFound,
        core::Error::SlidePositionNotFound(position) => {
            SlideTableCellNumberFormatError::SlidePositionNotFound { position }
        },
        core::Error::TablePositionNotFound(position) => {
            SlideTableCellNumberFormatError::TablePositionNotFound { position }
        },
        core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCellNumberFormatError::LimitExceeded {
            kind: map_limit_kind(kind),
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => SlideTableCellNumberFormatError::Allocation { amount },
        core::Error::InvalidSource
        | core::Error::Read
        | core::Error::Wire
        | core::Error::Codec
        | core::Error::Archive
        | core::Error::Verification => SlideTableCellNumberFormatError::InvalidSource,
    }
}

const fn map_limit_kind(kind: core::LimitKind) -> SlideTableCellNumberFormatLimitKind {
    match kind {
        core::LimitKind::InputBytes => SlideTableCellNumberFormatLimitKind::InputBytes,
        core::LimitKind::OutputBytes => SlideTableCellNumberFormatLimitKind::OutputBytes,
        core::LimitKind::Entries => SlideTableCellNumberFormatLimitKind::Entries,
        core::LimitKind::EntryBytes => SlideTableCellNumberFormatLimitKind::EntryBytes,
        core::LimitKind::TotalBytes => SlideTableCellNumberFormatLimitKind::TotalBytes,
        core::LimitKind::PayloadObjects => SlideTableCellNumberFormatLimitKind::PayloadObjects,
        core::LimitKind::PayloadMessages => SlideTableCellNumberFormatLimitKind::PayloadMessages,
        core::LimitKind::References => SlideTableCellNumberFormatLimitKind::References,
        core::LimitKind::WireFields => SlideTableCellNumberFormatLimitKind::WireFields,
        core::LimitKind::WireNesting => SlideTableCellNumberFormatLimitKind::WireNesting,
        core::LimitKind::WireWork => SlideTableCellNumberFormatLimitKind::WireWork,
        core::LimitKind::Allocations => SlideTableCellNumberFormatLimitKind::Allocations,
        core::LimitKind::Retained => SlideTableCellNumberFormatLimitKind::Retained,
        core::LimitKind::Scratch => SlideTableCellNumberFormatLimitKind::Scratch,
        core::LimitKind::Components => SlideTableCellNumberFormatLimitKind::Components,
    }
}

fn map_read_error(error: ReadError) -> SlideTableCellNumberFormatError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableCellNumberFormatError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableCellNumberFormatLimitKind::References,
                _ => SlideTableCellNumberFormatLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableCellNumberFormatError::LimitExceeded {
            kind: match kind {
                PayloadLimitKind::Bytes => SlideTableCellNumberFormatLimitKind::InputBytes,
                PayloadLimitKind::Fields => SlideTableCellNumberFormatLimitKind::WireFields,
                PayloadLimitKind::Nesting => SlideTableCellNumberFormatLimitKind::WireNesting,
                PayloadLimitKind::Work => SlideTableCellNumberFormatLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => {
            SlideTableCellNumberFormatError::Allocation { amount }
        },
        ReadError::Archive(error) => map_archive_error(error),
        _ => SlideTableCellNumberFormatError::InvalidSource,
    }
}

fn map_storage_error(error: storage_codec::DecodeError) -> SlideTableCellNumberFormatError {
    let Some(limit) = error.resource_limit() else {
        return SlideTableCellNumberFormatError::InvalidSource;
    };
    match limit {
        storage_codec::DecodeLimit::Bytes { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::InputBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        storage_codec::DecodeLimit::References { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::References,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        storage_codec::DecodeLimit::Text { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        storage_codec::DecodeLimit::Fields { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::WireFields,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        storage_codec::DecodeLimit::Work { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        storage_codec::DecodeLimit::Nesting { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        storage_codec::DecodeLimit::Allocation { requested } => {
            SlideTableCellNumberFormatError::Allocation { amount: requested }
        },
        storage_codec::DecodeLimit::Retained { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::Retained,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        _ => SlideTableCellNumberFormatError::InvalidSource,
    }
}

fn map_number_codec_error(error: number_codec::DecodeError) -> SlideTableCellNumberFormatError {
    let Some(limit) = error.resource_limit() else {
        return SlideTableCellNumberFormatError::InvalidSource;
    };
    match limit {
        number_codec::DecodeLimit::InputBytes { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::InputBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        number_codec::DecodeLimit::OutputBytes { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::OutputBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        number_codec::DecodeLimit::Fields { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::WireFields,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        number_codec::DecodeLimit::Work { observed, maximum }
        | number_codec::DecodeLimit::Items { observed, maximum }
        | number_codec::DecodeLimit::Text { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        number_codec::DecodeLimit::Nesting { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        number_codec::DecodeLimit::References { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::References,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        number_codec::DecodeLimit::Allocation { requested } => {
            SlideTableCellNumberFormatError::Allocation { amount: requested }
        },
        number_codec::DecodeLimit::Retained { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::Retained,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        number_codec::DecodeLimit::Scratch { observed, maximum } => {
            SlideTableCellNumberFormatError::LimitExceeded {
                kind: SlideTableCellNumberFormatLimitKind::Scratch,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        _ => SlideTableCellNumberFormatError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTableCellNumberFormatError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCellNumberFormatError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    SlideTableCellNumberFormatLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideTableCellNumberFormatLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => {
                    SlideTableCellNumberFormatLimitKind::Entries
                },
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideTableCellNumberFormatLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    SlideTableCellNumberFormatLimitKind::TotalBytes
                },
                _ => SlideTableCellNumberFormatLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableCellNumberFormatError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_native_error(error),
        _ => SlideTableCellNumberFormatError::InvalidSource,
    }
}

fn map_core_native_error(error: litchi_iwa_core::Error) -> SlideTableCellNumberFormatError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCellNumberFormatError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => {
                    SlideTableCellNumberFormatLimitKind::PayloadObjects
                },
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableCellNumberFormatLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideTableCellNumberFormatLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    SlideTableCellNumberFormatLimitKind::EntryBytes
                },
                _ => SlideTableCellNumberFormatLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableCellNumberFormatError::Allocation { amount: requested }
        },
        _ => SlideTableCellNumberFormatError::InvalidSource,
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CellWireError {
    Invalid,
    Missing,
    Unsupported,
}

fn field_varint(view: &WireView<'_>, number: u32) -> Result<Option<u64>, CellWireError> {
    let mut value = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 0 {
            return Err(CellWireError::Invalid);
        }
        let (candidate, consumed) =
            decode_varint_from_bytes(field.payload()).map_err(|_| CellWireError::Invalid)?;
        if consumed != field.payload().len() || value.replace(candidate).is_some() {
            return Err(CellWireError::Invalid);
        }
    }
    Ok(value)
}

fn field_bytes<'a>(view: &WireView<'a>, number: u32) -> Result<Option<&'a [u8]>, CellWireError> {
    let mut value = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 {
            return Err(CellWireError::Invalid);
        }
        let payload = field
            .canonical_payload()
            .map_err(|_| CellWireError::Invalid)?;
        if value.replace(payload).is_some() {
            return Err(CellWireError::Invalid);
        }
    }
    Ok(value)
}

fn tile_cell(source: &[u8], row: u32, column: u32) -> Result<&[u8], CellWireError> {
    let view = WireView::parse(source).map_err(|_| CellWireError::Invalid)?;
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == 5) {
        if field.wire_type() != 2 {
            return Err(CellWireError::Invalid);
        }
        let payload = field
            .canonical_payload()
            .map_err(|_| CellWireError::Invalid)?;
        let row_view = WireView::parse(payload).map_err(|_| CellWireError::Invalid)?;
        let Some(row_index) = field_varint(&row_view, 1)? else {
            return Err(CellWireError::Invalid);
        };
        let row_index = u32::try_from(row_index).map_err(|_| CellWireError::Invalid)?;
        if row_index != row {
            continue;
        }
        if selected.is_some() {
            return Err(CellWireError::Invalid);
        }
        let cell_count = field_varint(&row_view, 2)?.ok_or(CellWireError::Invalid)?;
        let buffer = field_bytes(&row_view, 6)?.ok_or(CellWireError::Invalid)?;
        selected = Some(select_row_cell(&row_view, buffer, cell_count, column)?);
    }
    selected.ok_or(CellWireError::Missing)
}

fn patch_tile_cell(
    source: &[u8],
    row: u32,
    column: u32,
    replacement: &[u8],
) -> Result<Vec<u8>, CellWireError> {
    let view = WireView::parse(source).map_err(|_| CellWireError::Invalid)?;
    let mut output = Vec::new();
    output
        .try_reserve(source.len().saturating_add(replacement.len()))
        .map_err(|_| CellWireError::Invalid)?;
    let mut selected = false;
    for field in view.fields() {
        if field.number() != 5 {
            output.extend_from_slice(field.raw());
            continue;
        }
        if field.wire_type() != 2 {
            return Err(CellWireError::Invalid);
        }
        let payload = field
            .canonical_payload()
            .map_err(|_| CellWireError::Invalid)?;
        let row_view = WireView::parse(payload).map_err(|_| CellWireError::Invalid)?;
        let row_index = field_varint(&row_view, 1)?
            .ok_or(CellWireError::Invalid)
            .and_then(|value| u32::try_from(value).map_err(|_| CellWireError::Invalid))?;
        if row_index != row {
            output.extend_from_slice(field.raw());
            continue;
        }
        if selected {
            return Err(CellWireError::Invalid);
        }
        let patched_row = patch_row_cell(payload, column, replacement)?;
        litchi_iwa_common::wire::append_length_delimited_field(&mut output, 5, &patched_row)
            .map_err(|_| CellWireError::Invalid)?;
        selected = true;
    }
    if !selected {
        return Err(CellWireError::Missing);
    }
    Ok(output)
}

fn select_row_cell<'a>(
    row: &WireView<'a>,
    buffer: &'a [u8],
    cell_count: u64,
    column: u32,
) -> Result<&'a [u8], CellWireError> {
    let count = usize::try_from(cell_count).map_err(|_| CellWireError::Invalid)?;
    let index = usize::try_from(column).map_err(|_| CellWireError::Unsupported)?;
    if count == 0 {
        return Err(CellWireError::Invalid);
    }
    let offset_field = field_bytes(row, 7)?;
    if count == 1 && offset_field.is_none() {
        return (index == 0)
            .then_some(buffer)
            .ok_or(CellWireError::Unsupported);
    }
    let offsets = offset_field.ok_or(CellWireError::Unsupported)?;
    let wide = row_wide_offsets(row)?;
    let unit = if wide { 4usize } else { 1usize };
    if offsets.len() % 2 != 0 {
        return Err(CellWireError::Invalid);
    }
    let slot_count = offsets.len() / 2;
    if slot_count < count || index >= slot_count {
        return Err(CellWireError::Unsupported);
    }
    let starts = decode_row_offsets(offsets, slot_count, unit)?;
    if starts.iter().flatten().count() != count {
        return Err(CellWireError::Invalid);
    }
    let start = starts[index].ok_or(CellWireError::Unsupported)?;
    let end = starts
        .iter()
        .skip(index + 1)
        .flatten()
        .next()
        .copied()
        .unwrap_or(buffer.len());
    if start > end || end > buffer.len() {
        return Err(CellWireError::Invalid);
    }
    Ok(&buffer[start..end])
}

fn decode_row_offsets(
    offsets: &[u8],
    count: usize,
    unit: usize,
) -> Result<Vec<Option<usize>>, CellWireError> {
    let expected = count.checked_mul(2).ok_or(CellWireError::Invalid)?;
    if offsets.len() != expected {
        return Err(CellWireError::Unsupported);
    }
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(count)
        .map_err(|_| CellWireError::Invalid)?;
    let mut previous = None;
    for encoded in offsets.chunks_exact(2) {
        let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
        if raw == u16::MAX {
            starts.push(None);
            continue;
        }
        let start = usize::from(raw)
            .checked_mul(unit)
            .ok_or(CellWireError::Invalid)?;
        if previous.is_some_and(|prior| prior > start) {
            return Err(CellWireError::Invalid);
        }
        starts.push(Some(start));
        previous = Some(start);
    }
    Ok(starts)
}

fn row_wide_offsets(row: &WireView<'_>) -> Result<bool, CellWireError> {
    let Some(value) = field_varint(row, 8)? else {
        return Ok(false);
    };
    if value > 1 {
        return Err(CellWireError::Invalid);
    }
    Ok(value != 0)
}

fn patch_row_cell(
    source: &[u8],
    column: u32,
    replacement: &[u8],
) -> Result<Vec<u8>, CellWireError> {
    let view = WireView::parse(source).map_err(|_| CellWireError::Invalid)?;
    let cell_count = field_varint(&view, 2)?.ok_or(CellWireError::Invalid)?;
    let buffer = field_bytes(&view, 6)?.ok_or(CellWireError::Invalid)?;
    let (patched_buffer, patched_offsets) =
        patch_row_buffer(&view, buffer, cell_count, column, replacement)?;
    let mut output = Vec::new();
    output
        .try_reserve(source.len().saturating_add(replacement.len()))
        .map_err(|_| CellWireError::Invalid)?;
    let mut replaced = false;
    let mut offsets_replaced = false;
    for field in view.fields() {
        if field.number() == 6 {
            if replaced {
                return Err(CellWireError::Invalid);
            }
            litchi_iwa_common::wire::append_length_delimited_field(&mut output, 6, &patched_buffer)
                .map_err(|_| CellWireError::Invalid)?;
            replaced = true;
        } else if field.number() == 7 {
            if let Some(offsets) = patched_offsets.as_deref() {
                if offsets_replaced {
                    return Err(CellWireError::Invalid);
                }
                litchi_iwa_common::wire::append_length_delimited_field(&mut output, 7, offsets)
                    .map_err(|_| CellWireError::Invalid)?;
                offsets_replaced = true;
            } else {
                output.extend_from_slice(field.raw());
            }
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(CellWireError::Invalid);
    }
    Ok(output)
}

fn patch_row_buffer(
    row: &WireView<'_>,
    buffer: &[u8],
    cell_count: u64,
    column: u32,
    replacement: &[u8],
) -> Result<(Vec<u8>, Option<Vec<u8>>), CellWireError> {
    let count = usize::try_from(cell_count).map_err(|_| CellWireError::Invalid)?;
    let index = usize::try_from(column).map_err(|_| CellWireError::Unsupported)?;
    if count == 0 {
        return Err(CellWireError::Invalid);
    }
    let offset_field = field_bytes(row, 7)?;
    if count == 1 && offset_field.is_none() {
        if index != 0 {
            return Err(CellWireError::Unsupported);
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(replacement.len())
            .map_err(|_| CellWireError::Invalid)?;
        output.extend_from_slice(replacement);
        return Ok((output, None));
    }
    let offsets = offset_field.ok_or(CellWireError::Unsupported)?;
    let wide = row_wide_offsets(row)?;
    let unit = if wide { 4usize } else { 1usize };
    if offsets.len() % 2 != 0 {
        return Err(CellWireError::Invalid);
    }
    let slot_count = offsets.len() / 2;
    if slot_count < count || index >= slot_count {
        return Err(CellWireError::Unsupported);
    }
    let starts = decode_row_offsets(offsets, slot_count, unit)?;
    if starts.iter().flatten().count() != count {
        return Err(CellWireError::Invalid);
    }
    let start = starts[index].ok_or(CellWireError::Unsupported)?;
    let end = starts
        .iter()
        .skip(index + 1)
        .flatten()
        .next()
        .copied()
        .unwrap_or(buffer.len());
    if start > end || end > buffer.len() {
        return Err(CellWireError::Invalid);
    }
    let old_len = end - start;
    let capacity = buffer
        .len()
        .checked_sub(old_len)
        .and_then(|value| value.checked_add(replacement.len()))
        .ok_or(CellWireError::Invalid)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_| CellWireError::Invalid)?;
    output.extend_from_slice(&buffer[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&buffer[end..]);

    let mut encoded_offsets = Vec::new();
    encoded_offsets
        .try_reserve_exact(offsets.len())
        .map_err(|_| CellWireError::Invalid)?;
    let delta = replacement.len() as isize - old_len as isize;
    for (offset_index, encoded) in offsets.chunks_exact(2).enumerate() {
        let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
        if raw == u16::MAX {
            encoded_offsets.extend_from_slice(&u16::MAX.to_le_bytes());
            continue;
        }
        let original = usize::from(raw)
            .checked_mul(unit)
            .ok_or(CellWireError::Invalid)?;
        let adjusted = if offset_index > index {
            if delta.is_negative() {
                original.checked_sub(delta.unsigned_abs())
            } else {
                original.checked_add(delta as usize)
            }
        } else {
            Some(original)
        }
        .ok_or(CellWireError::Invalid)?;
        if !adjusted.is_multiple_of(unit) || adjusted / unit > usize::from(u16::MAX) {
            return Err(CellWireError::Unsupported);
        }
        encoded_offsets.extend_from_slice(
            &u16::try_from(adjusted / unit)
                .map_err(|_| CellWireError::Unsupported)?
                .to_le_bytes(),
        );
    }
    let replacement_offsets = (delta != 0).then_some(encoded_offsets);
    Ok((output, replacement_offsets))
}

fn replace_archive_message(
    archives: &mut [(String, Archive)],
    component_name: &str,
    identifier: u64,
    message_index: usize,
    payload: Vec<u8>,
    limits: litchi_iwa_core::Limits,
) -> Result<(), SlideTableCellNumberFormatError> {
    let archive = archives
        .iter_mut()
        .find(|(name, _)| name == component_name)
        .map(|(_, archive)| archive)
        .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
    let object = archive
        .object_mut(identifier)
        .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
    let message_type = object
        .messages
        .get(message_index)
        .map(|message| message.type_)
        .ok_or(SlideTableCellNumberFormatError::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: message_type,
                data: payload,
            },
            limits,
        )
        .map_err(map_core_native_error)?;
    Ok(())
}

fn verify_package_locality(
    source: &Package,
    candidate: &Package,
    selection: &CellSelection,
    candidate_selection: &CellSelection,
    budget: &mut core::Budget,
) -> Result<(), SlideTableCellNumberFormatError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let source_entries = source_catalog.package();
    let candidate_entries = candidate_catalog.package();
    budget
        .entries(source_entries.len().saturating_add(candidate_entries.len()))
        .map_err(map_core_error)?;
    budget
        .allocations(source_entries.len().saturating_add(candidate_entries.len()))
        .map_err(map_core_error)?;

    let mut source_by_name = HashMap::new();
    source_by_name
        .try_reserve(source_entries.len())
        .map_err(|_| SlideTableCellNumberFormatError::Allocation {
            amount: source_entries.len(),
        })?;
    let mut candidate_by_name = HashMap::new();
    candidate_by_name
        .try_reserve(candidate_entries.len())
        .map_err(|_| SlideTableCellNumberFormatError::Allocation {
            amount: candidate_entries.len(),
        })?;
    for entry in source_entries.iter() {
        budget
            .work(entry.name().len().saturating_add(entry.data().len()))
            .map_err(map_core_error)?;
        if source_by_name.insert(entry.name(), entry).is_some() {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
    }
    for entry in candidate_entries.iter() {
        budget
            .work(entry.name().len().saturating_add(entry.data().len()))
            .map_err(map_core_error)?;
        if candidate_by_name.insert(entry.name(), entry).is_some() {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
    }
    if source_by_name.len() != candidate_by_name.len()
        || candidate_by_name
            .keys()
            .any(|name| !source_by_name.contains_key(name))
    {
        return Err(SlideTableCellNumberFormatError::Verification);
    }

    let mut changed_components = HashSet::new();
    changed_components.insert(selection.tile.component.as_ref());
    changed_components.insert(selection.format.component.as_ref());
    for (name, source_entry) in &source_by_name {
        let candidate_entry = candidate_by_name
            .get(name)
            .copied()
            .ok_or(SlideTableCellNumberFormatError::Verification)?;
        if !changed_components.contains(name)
            && (source_entry.data() != candidate_entry.data()
                || source_entry.metadata() != candidate_entry.metadata()
                || source_entry.raw_record().local_record()
                    != candidate_entry.raw_record().local_record()
                || !same_central_directory_record(
                    source_entry.raw_record().central_directory_record(),
                    candidate_entry.raw_record().central_directory_record(),
                ))
        {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
    }

    let limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| SlideTableCellNumberFormatError::InvalidSource)?;
    for component_name in changed_components {
        let source_component = core::component(source, component_name).map_err(map_core_error)?;
        let candidate_component =
            core::component(candidate, component_name).map_err(map_core_error)?;
        let source_archive = source_component.archive();
        let candidate_archive = candidate_component.archive();
        if source_archive.objects.len() != candidate_archive.objects.len() {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
        for source_object in &source_archive.objects {
            let identifier = source_object
                .archive_info
                .identifier
                .ok_or(SlideTableCellNumberFormatError::Verification)?;
            let candidate_object = candidate_archive
                .object(identifier)
                .ok_or(SlideTableCellNumberFormatError::Verification)?;
            let selected_message = if identifier == selection.tile.identifier {
                Some((selection.tile_message_index, TILE_MESSAGE_TYPE))
            } else if identifier == selection.format.identifier {
                source_object
                    .messages
                    .get(selection.format_message_index)
                    .map(|message| (selection.format_message_index, message.type_))
            } else {
                None
            };
            if identifier == selection.format.identifier
                && source_object
                    .messages
                    .get(selection.format_message_index)
                    .is_some_and(|message| message.type_ == TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE)
            {
                // The selected format-list root may use the native compatibility
                // type. Keep that exact type admitted by the locality fence.
                let index = selection.format_message_index;
                if candidate_object
                    .messages
                    .get(index)
                    .is_none_or(|message| message.type_ != TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE)
                {
                    return Err(SlideTableCellNumberFormatError::Verification);
                }
            }
            verify_object_delta(
                source_object,
                candidate_object,
                selected_message,
                limits,
                budget,
            )?;
        }
    }
    let _ = candidate_selection;
    Ok(())
}

fn verify_object_delta(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    selected_message: Option<(usize, u32)>,
    limits: litchi_iwa_core::Limits,
    budget: &mut core::Budget,
) -> Result<(), SlideTableCellNumberFormatError> {
    let identifier = source
        .archive_info
        .identifier
        .ok_or(SlideTableCellNumberFormatError::Verification)?;
    if candidate.archive_info.identifier != Some(identifier)
        || source.messages.len() != source.archive_info.message_infos.len()
        || candidate.messages.len() != candidate.archive_info.message_infos.len()
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
        || source.archive_info.should_merge != candidate.archive_info.should_merge
    {
        return Err(SlideTableCellNumberFormatError::Verification);
    }
    let mut expected = source.clone();
    let mut changed = false;
    for (index, (source_message, candidate_message)) in
        source.messages.iter().zip(&candidate.messages).enumerate()
    {
        let source_info = source
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideTableCellNumberFormatError::Verification)?;
        let candidate_info = candidate
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideTableCellNumberFormatError::Verification)?;
        if source_info.type_ != source_message.type_
            || candidate_info.type_ != candidate_message.type_
            || source_info.type_ != candidate_info.type_
            || source_info.versions != candidate_info.versions
            || source_info.field_infos != candidate_info.field_infos
            || source_info.object_references != candidate_info.object_references
            || source_info.data_references != candidate_info.data_references
            || source_info.base_message_index != candidate_info.base_message_index
            || source_info.diff_merge_version != candidate_info.diff_merge_version
            || source_info.diff_field_path != candidate_info.diff_field_path
            || source_info.fields_to_remove != candidate_info.fields_to_remove
            || source_info.diff_read_version != candidate_info.diff_read_version
            || u32::try_from(source_message.data.len()).ok() != Some(source_info.length)
            || u32::try_from(candidate_message.data.len()).ok() != Some(candidate_info.length)
        {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
        if source_message.data == candidate_message.data {
            continue;
        }
        let Some((allowed_index, allowed_type)) = selected_message else {
            return Err(SlideTableCellNumberFormatError::Verification);
        };
        if index != allowed_index || source_message.type_ != allowed_type {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
        expected
            .replace_message_preserving_header_with_limits(index, candidate_message.clone(), limits)
            .map_err(map_core_native_error)?;
        changed = true;
    }
    if !changed {
        if !source.same_content_ignoring_offsets(candidate) {
            return Err(SlideTableCellNumberFormatError::Verification);
        }
        return Ok(());
    }
    let ((header_length, data_length), expected) =
        physical_object_framing(expected, limits, budget)?;
    if candidate.header_length != header_length || candidate.data_length != data_length {
        return Err(SlideTableCellNumberFormatError::Verification);
    }
    if !expected.same_content_ignoring_offsets(candidate) {
        return Err(SlideTableCellNumberFormatError::Verification);
    }
    Ok(())
}

fn physical_object_framing(
    object: ArchiveObject,
    limits: litchi_iwa_core::Limits,
    budget: &mut core::Budget,
) -> Result<((u64, u64), ArchiveObject), SlideTableCellNumberFormatError> {
    let payload_length = object
        .messages
        .iter()
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(SlideTableCellNumberFormatError::Verification)?;
    let mut archive = Archive::new();
    archive
        .objects
        .try_reserve_exact(1)
        .map_err(|_| SlideTableCellNumberFormatError::Allocation { amount: 1 })?;
    archive.objects.push(object);
    budget
        .allocations(1)
        .and_then(|_| budget.retained(size_of::<ArchiveObject>()))
        .and_then(|_| budget.work(payload_length))
        .map_err(map_core_error)?;
    let encoded_length = archive
        .encoded_len_with_limits(limits)
        .map_err(map_core_native_error)?;
    let header_length = encoded_length
        .checked_sub(payload_length)
        .ok_or(SlideTableCellNumberFormatError::Verification)?;
    let object = archive
        .objects
        .pop()
        .ok_or(SlideTableCellNumberFormatError::Verification)?;
    let mut object = object;
    object.header_length =
        u64::try_from(header_length).map_err(|_| SlideTableCellNumberFormatError::Verification)?;
    object.data_length =
        u64::try_from(payload_length).map_err(|_| SlideTableCellNumberFormatError::Verification)?;
    Ok(((object.header_length, object.data_length), object))
}

fn same_central_directory_record(left: &[u8], right: &[u8]) -> bool {
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;
    left.len() == right.len()
        && left.len() >= LOCAL_HEADER_OFFSET.end
        && left[..LOCAL_HEADER_OFFSET.start] == right[..LOCAL_HEADER_OFFSET.start]
        && left[LOCAL_HEADER_OFFSET.end..] == right[LOCAL_HEADER_OFFSET.end..]
}

#[cfg(test)]
mod tests {
    use super::{
        CellPosition, DecimalPlaces, FixedDecimalPlaces, NegativeStyle, Number, Package,
        SlideTableCellNumberFormatError, ThousandsSeparator,
    };
    use crate::slide::table::TableSelector;
    use crate::{Position, SlideSelector};

    fn fixed_number(decimal_places: u8, separator: ThousandsSeparator) -> Number {
        Number::new(
            DecimalPlaces::Fixed(
                FixedDecimalPlaces::new(decimal_places)
                    .expect("test precision is within the checked Number domain"),
            ),
            NegativeStyle::MinusSign,
            separator,
        )
    }

    #[test]
    fn semantic_values_are_archive_free_and_checked() {
        assert_eq!(
            CellPosition::from_a1("B2").unwrap(),
            CellPosition::new(1, 1)
        );
        assert_eq!(CellPosition::new(1, 1).to_string(), "B2");
        assert!(CellPosition::from_a1("A0").is_err());
        assert!(FixedDecimalPlaces::new(31).is_err());
        assert_eq!(
            fixed_number(2, ThousandsSeparator::Shown).decimal_places(),
            DecimalPlaces::Fixed(FixedDecimalPlaces::TWO)
        );
    }

    /// Run against the verified native Keynote fixture without making the
    /// repository depend on a machine-local path. Set the
    /// LITCHI_KEYNOTE_NUMBER_SOURCE environment variable to the fixture path
    /// to exercise native readback, change, inverse, and exact patch conflict
    /// behavior.
    #[test]
    fn native_fixture_number_transaction_round_trips_when_configured() {
        let Some(path) = std::env::var_os("LITCHI_KEYNOTE_NUMBER_SOURCE") else {
            return;
        };
        let bytes = std::fs::read(path).expect("configured Keynote fixture is readable");
        let source = Package::from_bytes(&bytes).expect("native Keynote fixture reopens");
        let slide = SlideSelector::position(Position::new(0));
        let table = TableSelector::position(Position::new(0));
        let position = CellPosition::new(1, 1);
        let before = source
            .slide_table_cell_number_format(slide, table, position)
            .expect("fixture cell Number format is readable");
        assert_eq!(before, Some(fixed_number(2, ThousandsSeparator::Hidden)));

        let changed = fixed_number(3, ThousandsSeparator::Hidden);
        let commit = source
            .edit_slide_table_cell_number_format(slide, table, position)
            .expect("fixture cell edit opens")
            .set(changed)
            .commit()
            .expect("fixture Number change commits and reopens");
        assert_eq!(
            commit
                .package()
                .slide_table_cell_number_format(slide, table, position)
                .expect("changed candidate reads back"),
            Some(changed)
        );

        let patch = commit.patch().clone();
        assert!(!patch.is_noop());
        let applied = source
            .apply_slide_table_cell_number_format(&patch)
            .expect("exact patch applies to its source");
        assert_eq!(
            applied
                .package()
                .slide_table_cell_number_format(slide, table, position)
                .expect("applied patch reads back"),
            Some(changed)
        );
        let inverse = patch.inverse();
        let reverted = applied
            .package()
            .apply_slide_table_cell_number_format(&inverse)
            .expect("inverse patch applies to the changed candidate");
        assert_eq!(
            reverted
                .package()
                .slide_table_cell_number_format(slide, table, position)
                .expect("inverse reads back"),
            before
        );

        let cleared = source
            .edit_slide_table_cell_number_format(slide, table, position)
            .expect("fixture cell can be reopened for reset")
            .reset()
            .commit()
            .expect("reset commits and reopens");
        assert_eq!(
            cleared
                .package()
                .slide_table_cell_number_format(slide, table, position)
                .expect("reset candidate reads back"),
            None
        );

        let conflict = source
            .apply_slide_table_cell_number_format(&inverse)
            .expect_err("inverse must reject the original source");
        assert_eq!(conflict, SlideTableCellNumberFormatError::PatchConflict);
    }
}
