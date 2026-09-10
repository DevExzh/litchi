//! Private selector-to-storage handoff for focused Keynote table reads.
//!
//! The public semantic read is assembled by this package owner once the
//! shared sidecar decoders have resolved text, formulas, errors, rich text,
//! and comments.  This module owns only the Keynote-specific part of that
//! handoff: proving the selected model/DataStore route, collecting tile
//! references, and resolving archive messages without retaining native
//! objects in the result.

use std::cell::Cell as FlagCell;
use std::collections::HashMap;
use std::fmt;
use std::fmt::Write as _;

use litchi_core::Position;
use litchi_iwa_common::formula::render::FormulaRenderBudget;
use litchi_iwa_common::table::coordinate::CellPosition;
use litchi_iwa_common::table::model::{Cell, Dimensions};
use litchi_iwa_common::table::read::{
    CellComment, Comment, CommentAuthor, CommentReply, TableRead,
};
use litchi_iwa_common::wire::{WireDescent, WireView, preflight_wire_tree_with_limits};
use litchi_iwa_common::{WireLimits, decode_varint_from_bytes};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;
use litchi_iwa_protos::{annotation_author_codec, comment_storage_codec, numbers_formula_codec};
use litchi_numbers_wire::formula_envelope;
use litchi_numbers_wire::formula_names;
use litchi_numbers_wire::formula_render::{
    FormulaCategoryId, FormulaEventRenderBudget, FormulaTablePrefix, ReferenceResolver,
};
use litchi_numbers_wire::table_cells::{self as wire_cells, CellValueSink, TableCellReadBudget};
use litchi_numbers_wire::table_sidecars::{
    self as sidecars, CellSidecarResolver, SidecarAllocation, SidecarIssue, SidecarKind,
    SidecarReadBudget, SidecarReference, SidecarTables,
};
use std::mem::size_of;
use thiserror::Error;

use super::Package;
use super::slide_table_core as core;
use crate::{SlideSelector, slide::table::TableSelector};

/// One source-backed model/DataStore selection owned by a focused Keynote
/// table read.
///
/// The snapshots borrow the package's immutable archive bytes.  The target
/// retains the graph proof and is never exposed through the semantic API.
#[derive(Debug)]
pub(crate) struct SelectedSlideTableStorage<'source> {
    model: storage::TableModelSnapshot<'source>,
    data_store: storage::DataStoreSnapshot<'source>,
    tile_size: u32,
    tiles: Vec<wire_cells::TileReference>,
}

impl<'source> SelectedSlideTableStorage<'source> {
    #[must_use]
    pub(crate) fn model(&self) -> storage::TableModelSnapshot<'source> {
        self.model
    }

    #[must_use]
    pub(crate) fn data_store(&self) -> storage::DataStoreSnapshot<'source> {
        self.data_store
    }

    #[must_use]
    pub(crate) fn dimensions(&self) -> wire_cells::TableDimensions {
        wire_cells::TableDimensions::new(
            self.model.number_of_rows(),
            self.model.number_of_columns(),
            self.tile_size,
        )
    }

    /// Walk the selected tiles through the shared borrowed reader.
    pub(crate) fn read_into<Budget, Sink, Resolve>(
        &self,
        resolve_tile: Resolve,
        budget: &mut Budget,
        sink: &mut Sink,
    ) -> Result<wire_cells::TableCellReadReport, Budget::Error>
    where
        Budget: TableCellReadBudget,
        Sink: CellValueSink<Budget>,
        Resolve: FnMut(u64, &mut Budget) -> Result<Option<NativeMessages<'source>>, Budget::Error>,
    {
        wire_cells::read_table_cells(
            self.dimensions(),
            self.tiles.iter().copied(),
            resolve_tile,
            budget,
            sink,
        )
    }
}

/// Select and strictly decode one Keynote table's model/DataStore envelope.
///
/// Tile payloads and sidecar objects remain unresolved.  Their object
/// identities are retained only in this private source-backed value until the
/// shared cell and sidecar readers consume them under the same package budget.
pub(crate) fn selected_slide_table_storage<'source>(
    package: &'source Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    budget: &mut core::Budget,
) -> core::Result<SelectedSlideTableStorage<'source>> {
    let target = core::select_table(package, slide, table, budget)?;
    let payload = core::model_payload(package, &target)?;
    let options = budget.storage_codec_options(package)?;
    let mut routes = TileRoutes::default();
    let (projection, report) =
        storage::decode_table_model_with_data_store_and_visitor(payload, options, &mut routes)
            .map_err(|_| core::Error::Codec)?;
    budget.storage_codec_report(report)?;

    let model = projection.model();
    if model.number_of_rows() != target.rows || model.number_of_columns() != target.columns {
        return Err(core::Error::InvalidSource);
    }

    let data_store = projection.data_store();
    let tile_options = budget.storage_codec_options(package)?;
    let (tile_storage, tile_report) =
        storage::decode_tile_storage_with_report(data_store.tiles(), tile_options)
            .map_err(|_| core::Error::Codec)?;
    budget.storage_codec_report(tile_report)?;
    let tile_size = tile_storage.tile_size().ok_or(core::Error::InvalidSource)?;
    if tile_size == 0 {
        return Err(core::Error::InvalidSource);
    }

    let tile_count = usize::try_from(target.rows)
        .map_err(|_| core::Error::InvalidSource)?
        .checked_add(
            usize::try_from(tile_size)
                .map_err(|_| core::Error::InvalidSource)?
                .saturating_sub(1),
        )
        .and_then(|value| value.checked_div(usize::try_from(tile_size).ok()?))
        .ok_or(core::Error::InvalidSource)?;
    // A valid table may have no tile references when it has no rows or when
    // every addressable cell is absent from native sparse storage.  The
    // shared reader treats that as an empty semantic result; only routes that
    // cannot fit the selected tile grid are malformed.
    if routes.values.len() > tile_count {
        return Err(core::Error::UnsupportedDependency);
    }
    for tile in &routes.values {
        let _ = core::locate_object(package, tile.object_id())?;
        core::ensure_unique_identity(package, tile.object_id(), budget)?;
    }
    budget.allocations(routes.values.len())?;
    budget.retained(
        routes
            .values
            .len()
            .checked_mul(size_of::<wire_cells::TileReference>())
            .ok_or(core::Error::InvalidSource)?,
    )?;

    Ok(SelectedSlideTableStorage {
        model,
        data_store,
        tile_size,
        tiles: routes.values,
    })
}

/// Resolve one selected object to borrowed native messages.
///
/// The object index is authoritative and the identity check was completed for
/// every selected tile during admission, so this hot path performs only the
/// logarithmic lookup and exact object check. No object or message is cloned.
pub(crate) fn resolve_messages<'source>(
    package: &'source Package,
    object_id: u64,
    budget: &mut core::Budget,
) -> core::Result<Option<NativeMessages<'source>>> {
    if object_id == 0 {
        return Err(core::Error::InvalidSource);
    }
    let location = core::locate_object(package, object_id)?;
    let object = core::object_at(package, &location)?;
    budget.payload_messages(object.messages.len())?;
    Ok(Some(NativeMessages {
        messages: object.messages.as_slice(),
    }))
}

#[derive(Default)]
struct TileRoutes {
    values: Vec<wire_cells::TileReference>,
}

impl storage::StorageVisitor for TileRoutes {
    fn visit_tile_reference(
        &mut self,
        record: storage::TileReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        if self
            .values
            .iter()
            .any(|route| route.tile_index() == record.tile_id())
        {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        self.values
            .try_reserve(1)
            .map_err(|_| storage::DecodeError::allocation(1))?;
        let object_id = record.reference().identifier();
        if object_id == 0 {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        if self
            .values
            .iter()
            .any(|route| route.object_id() == object_id)
        {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        self.values
            .push(wire_cells::TileReference::new(record.tile_id(), object_id));
        Ok(())
    }
}

/// Borrowed messages from one archive object.
pub(crate) struct NativeMessages<'source> {
    messages: &'source [RawMessage],
}

impl<'source> IntoIterator for NativeMessages<'source> {
    type Item = wire_cells::Message<'source>;
    type IntoIter = NativeMessageIter<'source>;

    fn into_iter(self) -> Self::IntoIter {
        NativeMessageIter {
            messages: self.messages.iter(),
        }
    }
}

pub(crate) struct NativeMessageIter<'source> {
    messages: std::slice::Iter<'source, RawMessage>,
}

impl<'source> Iterator for NativeMessageIter<'source> {
    type Item = wire_cells::Message<'source>;

    fn next(&mut self) -> Option<Self::Item> {
        self.messages
            .next()
            .map(|message| wire_cells::Message::new(message.type_, message.data.as_slice()))
    }
}

/// Finite resources charged by a focused Keynote slide-table cell read.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[non_exhaustive]
pub enum SlideTableCellsLimitKind {
    /// Source bytes inspected by the selected graph and storage readers.
    InputBytes,
    /// Semantic output bytes retained by formula/rich-text rendering.
    OutputBytes,
    /// Native archive entries inspected by the package proof.
    Entries,
    /// Native entry bytes inspected by the package proof.
    EntryBytes,
    /// Aggregate native entry bytes inspected by the package proof.
    TotalBytes,
    /// Native payload objects inspected by the package proof.
    PayloadObjects,
    /// Native payload messages inspected by the package proof.
    PayloadMessages,
    /// Materialized semantic cells and sidecar records.
    PayloadItems,
    /// Native references inspected by the package proof.
    References,
    /// Wire fields inspected by a selected storage or sidecar decoder.
    WireFields,
    /// Maximum nested wire depth inspected by a selected decoder.
    WireNesting,
    /// Aggregate wire work performed by the read.
    WireWork,
    /// Fallible allocations admitted by the read.
    Allocations,
    /// Bytes retained by semantic values and sidecars.
    Retained,
    /// Temporary bytes retained by decoder staging.
    Scratch,
    /// Native components inspected by the package proof.
    Components,
}

impl fmt::Display for SlideTableCellsLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
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

/// Content-free semantic location for a focused slide-table cell read.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[non_exhaustive]
pub enum SlideTableCellsPath {
    /// The complete Keynote package.
    Package,
    /// One selected slide/table pair.
    Table { slide: Position, table: Position },
}

impl fmt::Display for SlideTableCellsPath {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Table { slide, table } => {
                write!(
                    formatter,
                    "slide {} table {} cells",
                    slide.get(),
                    table.get()
                )
            },
        }
    }
}

/// Failure from a focused Keynote slide-table cell read.
#[derive(Clone, Debug, Eq, PartialEq, Error)]
#[non_exhaustive]
pub enum SlideTableCellsError {
    /// The source is not a physical Keynote package.
    #[error("this Keynote source does not support focused slide-table cell reads")]
    UnsupportedSource,
    /// The selected table depends on a graph shape this reader does not own.
    #[error("the requested Keynote slide-table cell graph has an unsupported dependency")]
    UnsupportedDependency,
    /// The selected table graph is malformed or unsupported.
    #[error("the requested Keynote slide-table cell topology is unsupported")]
    UnsupportedTopology,
    /// The selector matched more than one semantic object.
    #[error("the Keynote slide-table cell selector is ambiguous")]
    AmbiguousSelector,
    /// The requested slide name is empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide matched the requested name.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// No slide exists at the requested position.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// No table exists at the requested position.
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    /// The selected archive graph or payload is malformed.
    #[error("the selected Keynote slide-table cell source is invalid")]
    InvalidSource,
    /// A finite read resource ceiling was exceeded.
    #[error(
        "Keynote slide-table cells {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: SlideTableCellsLimitKind,
        /// Observed resource count.
        observed: u64,
        /// Configured resource ceiling.
        maximum: u64,
    },
    /// A bounded staging allocation failed.
    #[error("could not allocate {amount} units for the Keynote slide-table cell read")]
    Allocation { amount: usize },
}

const MAX_MATERIALIZED_CELLS: usize = 1 << 20;
const MAX_SIDECAR_ENTRIES: usize = 1_000_000;
const RICH_TEXT_PAYLOAD_MESSAGE_KIND: u32 = 6_218;
const RICH_TEXT_STORAGE_MESSAGE_KINDS: [u32; 2] = [2_001, 2_022];
const ANNOTATION_AUTHOR_MESSAGE_KIND: u32 = 212;

/// One aggregate budget adapter shared by model, tile, sidecar, and semantic
/// value reads.  The package's core budget remains the single authority for
/// limits; these local counters cover only the shared reader's cumulative
/// cell/output observations.
struct ReadBudget<'package> {
    package: &'package Package,
    budget: core::Budget,
    materialized_cells: usize,
    output_bytes: usize,
}

impl<'package> ReadBudget<'package> {
    fn new(package: &'package Package) -> Result<Self, SlideTableCellsError> {
        Ok(Self {
            package,
            budget: core::Budget::new(package).map_err(map_core_error)?,
            materialized_cells: 0,
            output_bytes: 0,
        })
    }

    fn limits(&self) -> Result<WireLimits, SlideTableCellsError> {
        let residual = self.budget.residual(self.package).map_err(map_core_error)?;
        let input = self
            .budget
            .remaining_input()
            .map_err(map_core_error)?
            .clamp(1, WireLimits::MAX_INPUT_BYTES);
        let output = self
            .budget
            .remaining_output()
            .map_err(map_core_error)?
            .clamp(1, WireLimits::MAX_OUTPUT_BYTES);
        residual
            .with_input_bytes(input)
            .and_then(|limits| limits.with_output_bytes(output))
            .map_err(map_common_wire_error)
    }

    fn storage_options(
        &self,
        source: &[u8],
    ) -> Result<storage::DecodeOptions, SlideTableCellsError> {
        let limits = self.limits()?;
        let input = source.len().max(1);
        if input > limits.max_input_bytes() {
            return Err(limit_error(
                SlideTableCellsLimitKind::InputBytes,
                input,
                limits.max_input_bytes(),
            ));
        }
        let references = self.budget.remaining_references().map_err(map_core_error)?;
        Ok(storage::DecodeOptions::new(
            input,
            limits.max_fields().max(1),
            limits.max_rewrite_work().max(1),
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            references.max(1),
            limits.max_input_bytes().max(1),
        ))
    }

    fn charge_storage_report(
        &mut self,
        report: storage::DecodeReport,
    ) -> Result<(), SlideTableCellsError> {
        self.budget
            .input(report.source_bytes())
            .map_err(map_core_error)?;
        self.budget
            .fields(report.fields())
            .map_err(map_core_error)?;
        self.budget
            .work(report.work_bytes())
            .map_err(map_core_error)?;
        self.budget
            .references(report.references())
            .map_err(map_core_error)?;
        self.budget
            .nesting(report.max_depth() as usize)
            .map_err(map_core_error)
    }

    fn charge_budget_report(
        &mut self,
        fields: usize,
        work: usize,
        references: usize,
        depth: usize,
    ) -> Result<(), SlideTableCellsError> {
        self.budget.fields(fields).map_err(map_core_error)?;
        self.budget.work(work).map_err(map_core_error)?;
        self.budget.references(references).map_err(map_core_error)?;
        self.budget.nesting(depth).map_err(map_core_error)
    }

    fn charge_text(&mut self, amount: usize) -> Result<(), SlideTableCellsError> {
        self.budget.retained(amount).map_err(map_core_error)?;
        self.budget.output(amount).map_err(map_core_error)
    }

    fn copy_owned_text(&mut self, source: &str) -> Result<String, SlideTableCellsError> {
        self.budget.allocations(1).map_err(map_core_error)?;
        self.budget.retained(source.len()).map_err(map_core_error)?;
        self.budget.work(source.len()).map_err(map_core_error)?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(source.len())
            .map_err(|_| SlideTableCellsError::Allocation {
                amount: source.len(),
            })?;
        owned.push_str(source);
        Ok(owned)
    }
}

impl TableCellReadBudget for ReadBudget<'_> {
    type Error = SlideTableCellsError;

    fn storage_options(&mut self, source: &[u8]) -> Result<storage::DecodeOptions, Self::Error> {
        ReadBudget::storage_options(self, source)
    }

    fn charge_storage_report(&mut self, report: storage::DecodeReport) -> Result<(), Self::Error> {
        self.charge_storage_report(report)
    }

    fn check_materialized_cells(&mut self, observed: usize) -> Result<(), Self::Error> {
        check_materialized_cell_limit(observed)?;
        let additional = observed.saturating_sub(self.materialized_cells);
        self.budget.work(additional).map_err(map_core_error)?;
        self.materialized_cells = observed;
        Ok(())
    }

    fn charge_cell_source(&mut self, bytes: usize) -> Result<(), Self::Error> {
        self.budget.work(bytes).map_err(map_core_error)
    }

    fn charge_allocation(
        &mut self,
        _target: wire_cells::AllocationTarget,
        amount: usize,
    ) -> Result<(), Self::Error> {
        self.budget.allocations(amount).map_err(map_core_error)?;
        let bytes = amount
            .checked_mul(size_of::<u32>())
            .ok_or(SlideTableCellsError::InvalidSource)?;
        self.budget.retained(bytes).map_err(map_core_error)
    }

    fn map_issue(&mut self, issue: wire_cells::TableCellIssue) -> Self::Error {
        map_table_cell_issue(issue)
    }
}

fn check_materialized_cell_limit(observed: usize) -> Result<(), SlideTableCellsError> {
    if observed > MAX_MATERIALIZED_CELLS {
        return Err(limit_error(
            SlideTableCellsLimitKind::PayloadItems,
            observed,
            MAX_MATERIALIZED_CELLS,
        ));
    }
    Ok(())
}

impl SidecarReadBudget for ReadBudget<'_> {
    type Error = SlideTableCellsError;

    fn list_options(
        &mut self,
        _kind: SidecarKind,
        source: &[u8],
    ) -> Result<storage::DecodeOptions, Self::Error> {
        self.storage_options(source)
    }

    fn charge_list_report(
        &mut self,
        _kind: SidecarKind,
        report: storage::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.charge_storage_report(report)
    }

    fn charge_list_probe_report(
        &mut self,
        _kind: SidecarKind,
        report: storage::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.budget
            .input(report.source_bytes())
            .map_err(map_core_error)?;
        self.budget
            .fields(report.fields())
            .map_err(map_core_error)?;
        self.budget
            .work(report.work_bytes())
            .map_err(map_core_error)?;
        self.budget
            .nesting(report.max_depth() as usize)
            .map_err(map_core_error)
    }

    fn charge_retained(
        &mut self,
        _kind: SidecarKind,
        target: SidecarAllocation,
        amount: usize,
    ) -> Result<(), Self::Error> {
        match target {
            SidecarAllocation::Text | SidecarAllocation::FormulaBytes => {
                self.budget.retained(amount).map_err(map_core_error)?;
            },
            SidecarAllocation::Keys
            | SidecarAllocation::Values
            | SidecarAllocation::Segments
            | SidecarAllocation::Comments => {
                self.budget.allocations(amount).map_err(map_core_error)?;
                self.budget
                    .retained(
                        amount
                            .checked_mul(size_of::<u64>())
                            .ok_or(SlideTableCellsError::InvalidSource)?,
                    )
                    .map_err(map_core_error)?;
            },
        }
        Ok(())
    }

    fn charge_entry(
        &mut self,
        _kind: SidecarKind,
        _key: u32,
        source_bytes: usize,
    ) -> Result<(), Self::Error> {
        self.budget.work(source_bytes).map_err(map_core_error)?;
        self.budget.entries(1).map_err(map_core_error)
    }

    fn max_entries(&self, _kind: SidecarKind) -> usize {
        MAX_SIDECAR_ENTRIES.min(self.package.semantic_limits().max_references())
    }

    fn map_issue(&mut self, issue: SidecarIssue) -> Self::Error {
        map_sidecar_issue(issue)
    }

    fn comment_options(
        &mut self,
        source: &[u8],
    ) -> Result<comment_storage_codec::DecodeOptions, Self::Error> {
        let limits = self.limits()?;
        let input = source.len().max(1);
        if input > limits.max_input_bytes() {
            return Err(limit_error(
                SlideTableCellsLimitKind::InputBytes,
                input,
                limits.max_input_bytes(),
            ));
        }
        Ok(comment_storage_codec::DecodeOptions::new(
            input,
            limits.max_fields().max(1),
            limits.max_rewrite_work().max(1),
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            self.budget
                .remaining_references()
                .map_err(map_core_error)?
                .max(1),
            limits.max_input_bytes().max(1),
        ))
    }

    fn charge_comment_report(
        &mut self,
        report: comment_storage_codec::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.budget
            .input(report.source_bytes())
            .map_err(map_core_error)?;
        self.charge_budget_report(
            report.fields(),
            report.work_bytes(),
            report.references(),
            report.max_depth() as usize,
        )?;
        self.budget
            .work(report.reference_bytes())
            .map_err(map_core_error)?;
        self.budget
            .work(report.text_bytes())
            .map_err(map_core_error)?;
        self.budget.work(report.replies()).map_err(map_core_error)
    }

    fn formula_options(
        &mut self,
        source: &[u8],
    ) -> Result<numbers_formula_codec::DecodeOptions, Self::Error> {
        let limits = self.limits()?;
        let input = source.len().max(1);
        if input > limits.max_input_bytes() {
            return Err(limit_error(
                SlideTableCellsLimitKind::InputBytes,
                input,
                limits.max_input_bytes(),
            ));
        }
        Ok(numbers_formula_codec::DecodeOptions::new(
            input,
            limits.max_fields().max(1),
            limits.max_rewrite_work().max(1),
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            limits.max_fields().max(1),
            limits.max_input_bytes().max(1),
        )
        .with_opaque_unknown_fields(true)
        .with_unknown_functions(true)
        .with_render_recursion_limit(u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX)))
    }

    fn charge_formula_report(
        &mut self,
        report: numbers_formula_codec::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.budget.input(report.bytes()).map_err(map_core_error)?;
        self.charge_budget_report(
            report.fields(),
            report.work(),
            0,
            report.max_depth() as usize,
        )?;
        self.budget
            .work(report.text_bytes())
            .map_err(map_core_error)
    }

    fn formula_envelope_limits(
        &mut self,
        source: &[u8],
    ) -> Result<formula_envelope::FormulaEnvelopeLimits, Self::Error> {
        let limits = self.limits()?;
        let input = source.len().max(1);
        if input > limits.max_input_bytes() {
            return Err(limit_error(
                SlideTableCellsLimitKind::InputBytes,
                input,
                limits.max_input_bytes(),
            ));
        }
        let max_fields = limits.max_fields();
        let max_work = limits.max_rewrite_work();
        Ok(formula_envelope::FormulaEnvelopeLimits {
            max_fields,
            // The scanner reports aggregate parent and nested payload bytes;
            // the source length is only the initial admission check. Give it
            // the residual operation budget so a valid nested envelope is
            // not rejected once recursive wire work exceeds its root length.
            max_input_bytes: limits.max_input_bytes(),
            max_work,
            // `residual` already subtracts all earlier work from these
            // ceilings, so this scan starts with a zero local base.
            base_fields: 0,
            base_work: 0,
        })
    }

    fn charge_formula_envelope_report(
        &mut self,
        report: formula_envelope::FormulaEnvelopeReport,
    ) -> Result<(), Self::Error> {
        let Some(wire) = report.wire_preflight() else {
            return Ok(());
        };
        self.budget.fields(wire.fields()).map_err(map_core_error)?;
        self.budget
            .work(wire.scanned_bytes())
            .map_err(map_core_error)?;
        self.budget
            .nesting(wire.max_depth())
            .map_err(map_core_error)
    }

    fn retain_formula_envelope_cost(
        &mut self,
        cost: formula_envelope::AttemptedFormulaEnvelopeCost,
    ) {
        let _ = self.budget.input(cost.work());
        let _ = self.budget.fields(cost.fields());
        let _ = self.budget.work(cost.work());
    }

    fn map_formula_envelope_error(&mut self, error: litchi_iwa_common::Error) -> Self::Error {
        map_common_wire_error(error)
    }
}

impl FormulaRenderBudget for ReadBudget<'_> {
    type Error = SlideTableCellsError;

    fn output_limit(&self, observed: usize) -> Self::Error {
        limit_error(
            SlideTableCellsLimitKind::OutputBytes,
            observed,
            self.package.limits().max_total_bytes() as usize,
        )
    }

    fn allocation(&self, _resource: &'static str, amount: usize) -> Self::Error {
        SlideTableCellsError::Allocation { amount }
    }

    fn invalid(&self, _message: &'static str) -> Self::Error {
        SlideTableCellsError::InvalidSource
    }

    fn check(&self, amount: usize) -> Result<(), Self::Error> {
        let remaining = self.budget.remaining_output().map_err(map_core_error)?;
        if amount > remaining {
            let observed = self
                .output_bytes
                .checked_add(amount)
                .ok_or_else(|| self.output_limit(usize::MAX))?;
            return Err(self.output_limit(observed));
        }
        Ok(())
    }

    fn check_structure(&self, nodes: usize, parts: usize) -> Result<(), Self::Error> {
        let maximum = self.budget.remaining_work().map_err(map_core_error)?;
        if nodes > maximum || parts > maximum {
            return Err(limit_error(
                SlideTableCellsLimitKind::WireWork,
                nodes.max(parts),
                maximum,
            ));
        }
        Ok(())
    }

    fn charge(&mut self, amount: usize) -> Result<(), Self::Error> {
        self.check(amount)?;
        self.budget.output(amount).map_err(map_core_error)?;
        self.output_bytes = self
            .output_bytes
            .checked_add(amount)
            .ok_or(SlideTableCellsError::InvalidSource)?;
        Ok(())
    }
}

impl FormulaEventRenderBudget for ReadBudget<'_> {
    fn check_render_depth(&self, depth: usize) -> Result<(), Self::Error> {
        let maximum = self.limits()?.max_nesting();
        if depth > maximum {
            return Err(limit_error(
                SlideTableCellsLimitKind::WireNesting,
                depth,
                maximum,
            ));
        }
        Ok(())
    }

    fn parse_error(&self, _message: String) -> Self::Error {
        SlideTableCellsError::InvalidSource
    }

    fn invalid_format(&self, _message: String) -> Self::Error {
        SlideTableCellsError::InvalidSource
    }
}

/// Bounded formula names resolved from rooted Keynote slide tables.
///
/// A formula owner points at a native `TableInfo` object.  The table-name
/// catalog is built from the same slide-owned graph proof as table selection;
/// owner records and category nodes are then projected by the shared borrowed
/// wire readers.  Unknown identities remain a hard error after rendering so a
/// renderer compatibility fallback cannot become semantic data.
struct FormulaNameResolver {
    table_names: Vec<Box<str>>,
    owners: HashMap<[u32; 4], usize>,
    categories: HashMap<[u64; 2], Box<str>>,
    unresolved_table: FlagCell<bool>,
    unresolved_category: FlagCell<bool>,
}

impl FormulaNameResolver {
    fn has_unresolved_reference(&self) -> bool {
        self.unresolved_table.get() || self.unresolved_category.get()
    }
}

impl ReferenceResolver for FormulaNameResolver {
    fn table_prefix(
        &self,
        _id: &numbers_formula_codec::FormulaRenderCfuuid,
    ) -> Option<FormulaTablePrefix<'_>> {
        // Keynote has no proven sheet prefix in this focused profile.  Keep
        // the prefix unresolved so a renderer fallback cannot invent one.
        self.unresolved_table.set(true);
        None
    }

    fn table_only_name(&self, id: &numbers_formula_codec::FormulaRenderCfuuid) -> Option<&str> {
        let key = match (id.word0, id.word1, id.word2, id.word3) {
            (Some(word0), Some(word1), Some(word2), Some(word3)) => [word0, word1, word2, word3],
            _ => {
                self.unresolved_table.set(true);
                return None;
            },
        };
        let Some(index) = self.owners.get(&key).copied() else {
            self.unresolved_table.set(true);
            return None;
        };
        let Some(name) = self.table_names.get(index) else {
            self.unresolved_table.set(true);
            return None;
        };
        Some(name)
    }

    fn category_name(&self, id: FormulaCategoryId) -> Option<&str> {
        if let Some(name) = self.categories.get(&[id.lower, id.upper]) {
            return Some(name);
        }
        if id.lower == 1 && id.upper == 0 {
            return Some("Grand Total");
        }
        self.unresolved_category.set(true);
        None
    }

    fn function_name(&self, index: u32) -> Option<&str> {
        litchi_numbers_wire::function_map::function_name(index)
    }
}

const FORMULA_OWNER_MESSAGE_KIND: u32 = 4_008;
const FORMULA_CATEGORY_MESSAGE_KIND: u32 = 6_383;

impl FormulaNameResolver {
    fn build(package: &Package, budget: &mut ReadBudget<'_>) -> Result<Self, SlideTableCellsError> {
        let rooted =
            core::rooted_table_name_catalog(package, &mut budget.budget).map_err(map_core_error)?;
        let target_count = rooted.len();
        let mut table_names = Vec::new();
        if target_count != 0 {
            budget.budget.allocations(1).map_err(map_core_error)?;
            budget
                .budget
                .retained(
                    target_count
                        .checked_mul(size_of::<Box<str>>())
                        .ok_or(SlideTableCellsError::InvalidSource)?,
                )
                .map_err(map_core_error)?;
            table_names.try_reserve_exact(target_count).map_err(|_| {
                SlideTableCellsError::Allocation {
                    amount: target_count,
                }
            })?;
        }
        let mut table_indices = HashMap::new();
        if target_count != 0 {
            budget.budget.allocations(1).map_err(map_core_error)?;
            budget
                .budget
                .retained(
                    target_count
                        .checked_mul(size_of::<(u64, usize)>())
                        .ok_or(SlideTableCellsError::InvalidSource)?,
                )
                .map_err(map_core_error)?;
            table_indices.try_reserve(target_count).map_err(|_| {
                SlideTableCellsError::Allocation {
                    amount: target_count,
                }
            })?;
        }
        for table in rooted {
            let index = table_names.len();
            if table_indices
                .insert(table.table_info_identifier, index)
                .is_some()
            {
                return Err(SlideTableCellsError::InvalidSource);
            }
            table_names.push(table.name);
        }

        let mut owners = HashMap::new();
        let mut categories = HashMap::new();
        for component in package.state.source.components().iter() {
            budget.budget.components(1).map_err(map_core_error)?;
            for object in &component.archive().objects {
                budget.budget.payload_objects(1).map_err(map_core_error)?;
                for message in &object.messages {
                    budget.budget.payload_messages(1).map_err(map_core_error)?;
                    let work = message
                        .data
                        .len()
                        .checked_add(1)
                        .ok_or(SlideTableCellsError::InvalidSource)?;
                    budget.budget.work(work).map_err(map_core_error)?;
                    if message.type_ == FORMULA_CATEGORY_MESSAGE_KIND {
                        read_formula_categories(&message.data, &mut categories, budget)?;
                    }
                    if message.type_ != FORMULA_OWNER_MESSAGE_KIND {
                        continue;
                    }
                    let limits = formula_name_wire_limits(budget, message.data.len())?;
                    let projection = match formula_names::read_formula_owner_dependencies(
                        &message.data,
                        formula_names::ReadLimits { wire: limits },
                    ) {
                        Ok(projection) => {
                            charge_formula_name_report(
                                budget,
                                projection.report().input_bytes(),
                                projection.report().fields(),
                                projection.report().work(),
                            )?;
                            projection
                        },
                        Err(failure) => {
                            let attempted = failure.attempted();
                            charge_formula_name_report(
                                budget,
                                attempted.input_bytes,
                                attempted.fields,
                                attempted.work,
                            )?;
                            if matches!(
                                failure.error(),
                                litchi_iwa_common::Error::LimitExceeded { .. }
                            ) {
                                return Err(map_common_wire_error(failure.error().clone()));
                            }
                            continue;
                        },
                    };
                    let Some(&table_index) = table_indices.get(&projection.table_info_ref()) else {
                        continue;
                    };
                    budget.budget.allocations(1).map_err(map_core_error)?;
                    budget
                        .budget
                        .retained(size_of::<([u32; 4], usize)>())
                        .map_err(map_core_error)?;
                    owners
                        .try_reserve(1)
                        .map_err(|_| SlideTableCellsError::Allocation { amount: 1 })?;
                    if owners
                        .insert(projection.cfuuid_words(), table_index)
                        .is_some()
                    {
                        return Err(SlideTableCellsError::InvalidSource);
                    }
                }
            }
        }

        Ok(Self {
            table_names,
            owners,
            categories,
            unresolved_table: FlagCell::new(false),
            unresolved_category: FlagCell::new(false),
        })
    }
}

fn formula_name_wire_limits(
    budget: &ReadBudget<'_>,
    source_len: usize,
) -> Result<WireLimits, SlideTableCellsError> {
    let limits = budget.limits()?;
    let fields = limits.max_fields();
    let work = limits.max_rewrite_work();
    if fields == 0 {
        return Err(limit_error(SlideTableCellsLimitKind::WireFields, 1, fields));
    }
    if work == 0 {
        return Err(limit_error(SlideTableCellsLimitKind::WireWork, 1, work));
    }
    let input = work.max(source_len);
    if input > limits.max_input_bytes() {
        return Err(limit_error(
            SlideTableCellsLimitKind::InputBytes,
            input,
            limits.max_input_bytes(),
        ));
    }
    WireLimits::default()
        .with_input_bytes(input.max(1))
        .and_then(|value| value.with_fields(fields.clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|value| value.with_nesting(limits.max_nesting().min(WireLimits::MAX_NESTING)))
        .and_then(|value| value.with_rewrite_work(work.clamp(1, WireLimits::MAX_REWRITE_WORK)))
        .map_err(map_common_wire_error)
}

fn charge_formula_name_report(
    budget: &mut ReadBudget<'_>,
    input_bytes: usize,
    fields: usize,
    work: usize,
) -> Result<(), SlideTableCellsError> {
    budget.budget.input(input_bytes).map_err(map_core_error)?;
    budget.charge_budget_report(fields, work, 0, 0)
}

fn read_formula_categories(
    source: &[u8],
    categories: &mut HashMap<[u64; 2], Box<str>>,
    budget: &mut ReadBudget<'_>,
) -> Result<(), SlideTableCellsError> {
    let limits = formula_name_wire_limits(budget, source.len())?;
    let category_limits = formula_names::CategoryReadLimits {
        wire: limits,
        max_nodes: limits.max_fields().clamp(1, WireLimits::MAX_FIELDS),
    };
    let projection = match formula_names::read_formula_category_names(source, category_limits) {
        Ok(projection) => {
            charge_formula_name_report(
                budget,
                projection.report.input_bytes(),
                projection.report.fields(),
                projection.report.work(),
            )?;
            projection
        },
        Err(failure) => {
            let attempted = failure.attempted();
            charge_formula_name_report(
                budget,
                attempted.input_bytes,
                attempted.fields,
                attempted.work,
            )?;
            if matches!(
                failure.error(),
                litchi_iwa_common::Error::LimitExceeded { .. }
            ) {
                return Err(map_common_wire_error(failure.error().clone()));
            }
            return Ok(());
        },
    };
    if projection.entries.is_empty() {
        return Ok(());
    }
    budget.budget.allocations(1).map_err(map_core_error)?;
    budget
        .budget
        .retained(
            projection
                .entries
                .len()
                .checked_mul(size_of::<([u64; 2], Box<str>)>())
                .ok_or(SlideTableCellsError::InvalidSource)?,
        )
        .map_err(map_core_error)?;
    categories
        .try_reserve(projection.entries.len())
        .map_err(|_| SlideTableCellsError::Allocation {
            amount: projection.entries.len(),
        })?;
    for entry in projection.entries {
        let reserve = entry.label().formatted_len();
        budget.budget.work(reserve).map_err(map_core_error)?;
        budget.budget.allocations(1).map_err(map_core_error)?;
        budget.budget.retained(reserve).map_err(map_core_error)?;
        let mut label = String::new();
        label
            .try_reserve_exact(reserve)
            .map_err(|_| SlideTableCellsError::Allocation { amount: reserve })?;
        write!(&mut label, "{}", entry.label()).map_err(|_| SlideTableCellsError::InvalidSource)?;
        if categories
            .insert(entry.id(), label.into_boxed_str())
            .is_some()
        {
            return Err(SlideTableCellsError::InvalidSource);
        }
    }
    Ok(())
}

fn read_sidecars(
    package: &Package,
    data_store: storage::DataStoreSnapshot<'_>,
    budget: &mut ReadBudget<'_>,
) -> Result<SidecarTables, SlideTableCellsError> {
    let mut tables = SidecarTables::new();
    let required = [
        (SidecarKind::Strings, data_store.string_table().identifier()),
        (
            SidecarKind::Formulas,
            data_store.formula_table().identifier(),
        ),
    ];
    for (kind, identifier) in required {
        let list = read_one_sidecar(package, kind, identifier, budget)?;
        tables
            .insert(list)
            .map_err(|issue| SidecarReadBudget::map_issue(budget, issue))?;
    }
    for (kind, reference) in [
        (SidecarKind::FormulaErrors, data_store.formula_error_table()),
        (SidecarKind::RichTextPayloads, data_store.rich_text_table()),
        (SidecarKind::Comments, data_store.comment_storage_table()),
    ] {
        let Some(reference) = reference else {
            continue;
        };
        let list = read_one_sidecar(package, kind, reference.identifier(), budget)?;
        tables
            .insert(list)
            .map_err(|issue| SidecarReadBudget::map_issue(budget, issue))?;
    }
    Ok(tables)
}

fn read_one_sidecar(
    package: &Package,
    kind: SidecarKind,
    identifier: u64,
    budget: &mut ReadBudget<'_>,
) -> Result<sidecars::SidecarList, SlideTableCellsError> {
    let Some(root) =
        resolve_messages(package, identifier, &mut budget.budget).map_err(map_core_error)?
    else {
        return Err(SidecarReadBudget::map_issue(
            budget,
            SidecarIssue::MissingPayload { kind },
        ));
    };
    sidecars::read_sidecar_list(
        kind,
        identifier,
        root,
        |segment_id, sidecar_budget| {
            resolve_messages(package, segment_id, &mut sidecar_budget.budget)
                .map_err(map_core_error)
        },
        budget,
    )
}

struct KeynoteCellSink<'sidecars, 'names> {
    table: litchi_iwa_common::table::model::Builder,
    comments: Vec<CellComment>,
    sidecars: &'sidecars SidecarTables,
    formula_names: Option<&'names FormulaNameResolver>,
    rows: u32,
    columns: u32,
}

impl<'sidecars, 'names> KeynoteCellSink<'sidecars, 'names> {
    fn new(
        name: String,
        dimensions: Dimensions,
        sidecars: &'sidecars SidecarTables,
        formula_names: Option<&'names FormulaNameResolver>,
    ) -> Self {
        Self {
            table: litchi_iwa_common::table::model::Builder::new(name, dimensions),
            comments: Vec::new(),
            sidecars,
            formula_names,
            rows: dimensions.rows(),
            columns: dimensions.columns(),
        }
    }

    fn finish(self) -> Result<TableRead, SlideTableCellsError> {
        let table = self.table.finish().map_err(map_common_table_error)?;
        TableRead::try_from_owned_parts(table, self.comments).map_err(map_common_table_error)
    }
}

impl CellValueSink<ReadBudget<'_>> for KeynoteCellSink<'_, '_> {
    fn visit_cell(
        &mut self,
        cell: wire_cells::CellSource<'_>,
        budget: &mut ReadBudget<'_>,
    ) -> Result<(), SlideTableCellsError> {
        budget.budget.allocations(1).map_err(map_core_error)?;
        budget
            .budget
            .retained(size_of::<Cell>())
            .map_err(map_core_error)?;
        let mut resolver = KeynoteCellSidecarResolver {
            budget,
            formula_names: self.formula_names,
            rows: self.rows,
            columns: self.columns,
        };
        let (value, comment) = self.sidecars.materialize_cell(
            cell.value(),
            cell.row(),
            cell.column(),
            &mut resolver,
        )?;
        self.table
            .push(Cell::new(
                CellPosition::new(cell.row(), cell.column()),
                value,
            ))
            .map_err(|error| map_common_table_error(error.into_parts().0))?;
        if let Some(comment) = comment {
            budget.budget.allocations(1).map_err(map_core_error)?;
            budget
                .budget
                .retained(size_of::<CellComment>())
                .map_err(map_core_error)?;
            self.comments
                .try_reserve(1)
                .map_err(|_| SlideTableCellsError::Allocation {
                    amount: self.comments.len().saturating_add(1),
                })?;
            self.comments.push(CellComment::new(
                CellPosition::new(cell.row(), cell.column()),
                comment,
            ));
        }
        Ok(())
    }
}

struct KeynoteCellSidecarResolver<'package, 'budget, 'names> {
    budget: &'budget mut ReadBudget<'package>,
    formula_names: Option<&'names FormulaNameResolver>,
    rows: u32,
    columns: u32,
}

impl<'package, 'budget, 'names> KeynoteCellSidecarResolver<'package, 'budget, 'names> {
    fn canonical_message<'source>(
        messages: NativeMessages<'source>,
        kind: u32,
    ) -> Result<&'source [u8], SlideTableCellsError> {
        let mut payload = None;
        for message in messages {
            if message.kind != kind {
                continue;
            }
            if payload.replace(message.data).is_some() {
                return Err(SlideTableCellsError::InvalidSource);
            }
        }
        payload.ok_or(SlideTableCellsError::InvalidSource)
    }

    fn canonical_message_kinds<'source>(
        messages: NativeMessages<'source>,
        kinds: &[u32; 2],
    ) -> Result<&'source [u8], SlideTableCellsError> {
        let mut payload = None;
        for message in messages {
            if !kinds.contains(&message.kind) {
                continue;
            }
            if payload.replace(message.data).is_some() {
                return Err(SlideTableCellsError::InvalidSource);
            }
        }
        payload.ok_or(SlideTableCellsError::InvalidSource)
    }

    fn resolve_object_payload(
        &mut self,
        reference: SidecarReference,
        kind: u32,
    ) -> Result<&'package [u8], SlideTableCellsError> {
        let messages = resolve_messages(
            self.budget.package,
            reference.identifier(),
            &mut self.budget.budget,
        )
        .map_err(map_core_error)?
        .ok_or(SlideTableCellsError::InvalidSource)?;
        Self::canonical_message(messages, kind)
    }

    fn read_author(
        &mut self,
        reference: SidecarReference,
    ) -> Result<CommentAuthor, SlideTableCellsError> {
        let payload = self.resolve_object_payload(reference, ANNOTATION_AUTHOR_MESSAGE_KIND)?;
        let limits = self.budget.limits()?;
        let bytes = payload.len().max(1);
        if bytes > limits.max_input_bytes() {
            return Err(limit_error(
                SlideTableCellsLimitKind::InputBytes,
                bytes,
                limits.max_input_bytes(),
            ));
        }
        let fields = limits.max_fields().max(1);
        let work = limits.max_rewrite_work().max(1);
        let options = annotation_author_codec::DecodeOptions::new(
            bytes,
            fields,
            work,
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            self.budget
                .budget
                .remaining_references()
                .map_err(map_core_error)?
                .max(1),
            limits.max_input_bytes().max(1),
            self.budget
                .budget
                .remaining_allocations()
                .map_err(map_core_error)?
                .max(1),
        );
        let (snapshot, report) =
            annotation_author_codec::decode_annotation_author_with_report(payload, options)
                .map_err(|_| SlideTableCellsError::InvalidSource)?;
        self.budget
            .budget
            .input(report.source_bytes())
            .map_err(map_core_error)?;
        self.budget.charge_budget_report(
            report.fields(),
            report.work_bytes(),
            report.references(),
            report.max_depth() as usize,
        )?;
        self.budget
            .budget
            .allocations(report.allocations())
            .map_err(map_core_error)?;
        self.budget.charge_text(report.text_bytes())?;
        let owned_strings = usize::from(snapshot.name().is_some_and(|value| !value.is_empty()))
            + usize::from(snapshot.public_id().is_some_and(|value| !value.is_empty()));
        self.budget
            .budget
            .allocations(owned_strings)
            .map_err(map_core_error)?;
        CommentAuthor::try_new(snapshot.name(), snapshot.public_id())
            .map_err(map_common_table_error)
    }

    fn read_reply(
        &mut self,
        reference: SidecarReference,
    ) -> Result<CommentReply, SlideTableCellsError> {
        let messages = resolve_messages(
            self.budget.package,
            reference.identifier(),
            &mut self.budget.budget,
        )
        .map_err(map_core_error)?
        .ok_or(SlideTableCellsError::InvalidSource)?;
        let storage = sidecars::read_comment_storage(reference, messages, &mut *self.budget)?;
        let (text, timestamp, author_reference, replies, _storage_uuid) = storage.into_parts();
        if !replies.is_empty() {
            return Err(SlideTableCellsError::UnsupportedDependency);
        }
        let author = author_reference
            .map(|reference| self.read_author(reference))
            .transpose()?;
        Ok(CommentReply::from_owned_parts(text, timestamp, author))
    }
}

impl CellSidecarResolver for KeynoteCellSidecarResolver<'_, '_, '_> {
    type Error = SlideTableCellsError;

    fn rich_text(&mut self, reference: SidecarReference) -> Result<String, Self::Error> {
        let payload = self.resolve_object_payload(reference, RICH_TEXT_PAYLOAD_MESSAGE_KIND)?;
        let limits = self.budget.limits()?;
        let (storage_identifier, report) = preflight_rich_text_payload(payload, limits)?;
        self.budget.charge_budget_report(
            report.fields(),
            report.scanned_bytes(),
            0,
            report.max_depth(),
        )?;
        let storage_reference =
            SidecarReference::new(storage_identifier).ok_or(SlideTableCellsError::InvalidSource)?;
        let messages = resolve_messages(
            self.budget.package,
            storage_reference.identifier(),
            &mut self.budget.budget,
        )
        .map_err(map_core_error)?
        .ok_or(SlideTableCellsError::InvalidSource)?;
        let storage_payload =
            Self::canonical_message_kinds(messages, &RICH_TEXT_STORAGE_MESSAGE_KINDS)?;
        self.budget
            .budget
            .work(storage_payload.len())
            .map_err(map_core_error)?;
        let text_limits = litchi_iwa_text_wire::Limits::new(
            storage_payload.len().max(1),
            limits.max_fields().max(1),
            limits.max_fields().max(1),
            limits.max_rewrite_work().max(1),
        )
        .map_err(|_| SlideTableCellsError::InvalidSource)?;
        let storage = litchi_iwa_text_wire::from_bytes_with_limits(storage_payload, text_limits)
            .map_err(|_| SlideTableCellsError::InvalidSource)?;
        self.budget.charge_text(storage.len())?;
        Ok(storage.into_text())
    }

    fn formula(
        &mut self,
        _key: u32,
        source: &[u8],
        row: u32,
        column: u32,
    ) -> Result<String, Self::Error> {
        let Some(resolver) = self.formula_names else {
            return Err(SlideTableCellsError::UnsupportedDependency);
        };
        let rendered = sidecars::render_formula(
            source,
            1,
            row,
            column,
            self.rows,
            self.columns,
            resolver,
            &mut *self.budget,
        )?;
        if resolver.has_unresolved_reference() {
            return Err(SlideTableCellsError::UnsupportedDependency);
        }
        Ok(rendered)
    }

    fn comment(&mut self, reference: SidecarReference) -> Result<Comment, Self::Error> {
        let messages = resolve_messages(
            self.budget.package,
            reference.identifier(),
            &mut self.budget.budget,
        )
        .map_err(map_core_error)?
        .ok_or(SlideTableCellsError::InvalidSource)?;
        let storage = sidecars::read_comment_storage(reference, messages, &mut *self.budget)?;
        let (text, timestamp, author_reference, reply_references, _storage_uuid) =
            storage.into_parts();
        let author = author_reference
            .map(|reference| self.read_author(reference))
            .transpose()?;
        if !reply_references.is_empty() {
            self.budget.budget.allocations(1).map_err(map_core_error)?;
            self.budget
                .budget
                .retained(
                    reply_references
                        .len()
                        .checked_mul(size_of::<CommentReply>())
                        .ok_or(SlideTableCellsError::InvalidSource)?,
                )
                .map_err(map_core_error)?;
        }
        let mut replies = Vec::new();
        replies
            .try_reserve_exact(reply_references.len())
            .map_err(|_| SlideTableCellsError::Allocation {
                amount: reply_references.len(),
            })?;
        for reply in reply_references {
            replies.push(self.read_reply(reply)?);
        }
        Ok(Comment::from_owned_parts(
            text,
            timestamp,
            author,
            Some(replies.into_boxed_slice()),
        ))
    }

    fn retain_text(
        &mut self,
        _kind: SidecarKind,
        _key: u32,
        source: &str,
    ) -> Result<String, Self::Error> {
        self.budget.charge_text(source.len())?;
        let mut value = String::new();
        value
            .try_reserve_exact(source.len())
            .map_err(|_| SlideTableCellsError::Allocation {
                amount: source.len(),
            })?;
        value.push_str(source);
        Ok(value)
    }

    fn missing(&mut self, _kind: SidecarKind, _key: u32) -> Self::Error {
        SlideTableCellsError::InvalidSource
    }

    fn invalid_scalar(&mut self) -> Self::Error {
        SlideTableCellsError::InvalidSource
    }
}

fn preflight_rich_text_payload(
    source: &[u8],
    limits: WireLimits,
) -> Result<(u64, litchi_iwa_common::wire::WirePreflight), SlideTableCellsError> {
    let input = source.len().max(1);
    if input > limits.max_input_bytes() {
        return Err(limit_error(
            SlideTableCellsLimitKind::InputBytes,
            input,
            limits.max_input_bytes(),
        ));
    }
    let mut storage_identifier = None;
    let mut has_cell_owner = false;
    let report = preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        match field.number() {
            1 => {
                if storage_identifier.is_some() || field.wire_type() != 2 {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "invalid rich-text storage reference".to_owned(),
                    ));
                }
                field.validate_canonical_framing()?;
                storage_identifier = Some(parse_local_reference(field.payload(), limits)?);
            },
            3 => {
                if has_cell_owner || field.wire_type() != 2 {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "invalid rich-text cell owner".to_owned(),
                    ));
                }
                field.validate_canonical_framing()?;
                has_cell_owner = true;
            },
            _ => {},
        }
        Ok(WireDescent::Skip)
    })
    .map_err(|_| SlideTableCellsError::InvalidSource)?;
    if !has_cell_owner {
        return Err(SlideTableCellsError::InvalidSource);
    }
    Ok((
        storage_identifier.ok_or(SlideTableCellsError::InvalidSource)?,
        report,
    ))
}

fn parse_local_reference(
    source: &[u8],
    limits: WireLimits,
) -> Result<u64, litchi_iwa_common::Error> {
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|_| litchi_iwa_common::Error::InvalidFormat("invalid local reference".into()))?;
    let mut identifier = None;
    for field in view.fields() {
        if field.number() != 1 || field.wire_type() != 0 || identifier.is_some() {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "invalid local reference".into(),
            ));
        }
        field.validate_canonical_framing().map_err(|_| {
            litchi_iwa_common::Error::InvalidFormat("invalid local reference".into())
        })?;
        let (value, length) = decode_varint_from_bytes(field.payload()).map_err(|_| {
            litchi_iwa_common::Error::InvalidFormat("invalid local reference".into())
        })?;
        if length != field.payload().len() || value == 0 {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "invalid local reference".into(),
            ));
        }
        identifier = Some(value);
    }
    identifier
        .ok_or_else(|| litchi_iwa_common::Error::InvalidFormat("invalid local reference".into()))
}

impl Package {
    /// Read one selected slide table into the archive-free common table model.
    ///
    /// The selected graph, model/DataStore, tile payloads, sidecar lists, and
    /// semantic values consume one aggregate budget.  All native objects,
    /// generated protobuf values, and object identifiers are discarded before
    /// this method returns.
    pub fn slide_table_cells<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<TableRead, SlideTableCellsError> {
        let mut budget = ReadBudget::new(self)?;
        let selected =
            selected_slide_table_storage(self, slide.into(), table.into(), &mut budget.budget)
                .map_err(map_core_error)?;
        let sidecars = read_sidecars(self, selected.data_store(), &mut budget)?;
        let formula_names = if sidecars
            .list(SidecarKind::Formulas)
            .is_some_and(|list| !list.is_empty())
        {
            Some(FormulaNameResolver::build(self, &mut budget)?)
        } else {
            None
        };
        let dimensions = Dimensions::new(
            selected.model().number_of_rows(),
            selected.model().number_of_columns(),
        );
        let table_name = budget.copy_owned_text(selected.model().table_name())?;
        let mut sink =
            KeynoteCellSink::new(table_name, dimensions, &sidecars, formula_names.as_ref());
        selected.read_into(
            |object_id, cell_budget| {
                resolve_messages(cell_budget.package, object_id, &mut cell_budget.budget)
                    .map_err(map_core_error)
            },
            &mut budget,
            &mut sink,
        )?;
        sink.finish()
    }
}

fn limit_error(
    kind: SlideTableCellsLimitKind,
    observed: usize,
    maximum: usize,
) -> SlideTableCellsError {
    SlideTableCellsError::LimitExceeded {
        kind,
        observed: usize_as_u64(observed),
        maximum: usize_as_u64(maximum),
    }
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn map_core_error(error: core::Error) -> SlideTableCellsError {
    match error {
        core::Error::UnsupportedSource => SlideTableCellsError::UnsupportedSource,
        core::Error::UnsupportedDependency => SlideTableCellsError::UnsupportedDependency,
        core::Error::UnsupportedTopology => SlideTableCellsError::UnsupportedTopology,
        core::Error::AmbiguousSelector => SlideTableCellsError::AmbiguousSelector,
        core::Error::EmptySlideName => SlideTableCellsError::EmptySlideName,
        core::Error::SlideNameNotFound => SlideTableCellsError::SlideNameNotFound,
        core::Error::SlidePositionNotFound(position) => {
            SlideTableCellsError::SlidePositionNotFound { position }
        },
        core::Error::TablePositionNotFound(position) => {
            SlideTableCellsError::TablePositionNotFound { position }
        },
        core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCellsError::LimitExceeded {
            kind: map_core_limit(kind),
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => SlideTableCellsError::Allocation { amount },
        core::Error::InvalidSource
        | core::Error::Read
        | core::Error::Wire
        | core::Error::Codec
        | core::Error::Archive
        | core::Error::Verification => SlideTableCellsError::InvalidSource,
    }
}

fn map_core_limit(kind: core::LimitKind) -> SlideTableCellsLimitKind {
    match kind {
        core::LimitKind::InputBytes => SlideTableCellsLimitKind::InputBytes,
        core::LimitKind::OutputBytes => SlideTableCellsLimitKind::OutputBytes,
        core::LimitKind::Entries => SlideTableCellsLimitKind::Entries,
        core::LimitKind::EntryBytes => SlideTableCellsLimitKind::EntryBytes,
        core::LimitKind::TotalBytes => SlideTableCellsLimitKind::TotalBytes,
        core::LimitKind::PayloadObjects => SlideTableCellsLimitKind::PayloadObjects,
        core::LimitKind::PayloadMessages => SlideTableCellsLimitKind::PayloadMessages,
        core::LimitKind::References => SlideTableCellsLimitKind::References,
        core::LimitKind::WireFields => SlideTableCellsLimitKind::WireFields,
        core::LimitKind::WireNesting => SlideTableCellsLimitKind::WireNesting,
        core::LimitKind::WireWork => SlideTableCellsLimitKind::WireWork,
        core::LimitKind::Allocations => SlideTableCellsLimitKind::Allocations,
        core::LimitKind::Retained => SlideTableCellsLimitKind::Retained,
        core::LimitKind::Scratch => SlideTableCellsLimitKind::Scratch,
        core::LimitKind::Components => SlideTableCellsLimitKind::Components,
    }
}

fn map_table_cell_issue(issue: wire_cells::TableCellIssue) -> SlideTableCellsError {
    match issue {
        wire_cells::TableCellIssue::Allocation { amount, .. } => {
            SlideTableCellsError::Allocation { amount }
        },
        wire_cells::TableCellIssue::StorageDecode(error) => map_storage_error(error),
        wire_cells::TableCellIssue::CellValueDecode { error, .. } => map_cell_value_error(error),
        wire_cells::TableCellIssue::CounterOverflow
        | wire_cells::TableCellIssue::MissingTile { .. }
        | wire_cells::TableCellIssue::MissingTilePayload { .. }
        | wire_cells::TableCellIssue::DuplicateTilePayload { .. }
        | wire_cells::TableCellIssue::TileIndexOutOfBounds { .. }
        | wire_cells::TableCellIssue::TileRowOutOfBounds { .. }
        | wire_cells::TableCellIssue::RowCoordinateOverflow { .. }
        | wire_cells::TableCellIssue::TableRowOutOfBounds { .. }
        | wire_cells::TableCellIssue::DuplicateTileRow { .. }
        | wire_cells::TableCellIssue::CellStorageOutOfBounds { .. } => {
            SlideTableCellsError::InvalidSource
        },
    }
}

fn map_storage_error(error: storage::DecodeError) -> SlideTableCellsError {
    let Some(limit) = error.resource_limit() else {
        return SlideTableCellsError::InvalidSource;
    };
    match limit {
        storage::DecodeLimit::Bytes { observed, maximum }
        | storage::DecodeLimit::Retained { observed, maximum } => {
            limit_error(SlideTableCellsLimitKind::InputBytes, observed, maximum)
        },
        storage::DecodeLimit::References { observed, maximum } => {
            limit_error(SlideTableCellsLimitKind::References, observed, maximum)
        },
        storage::DecodeLimit::Text { observed, maximum } => {
            limit_error(SlideTableCellsLimitKind::Retained, observed, maximum)
        },
        storage::DecodeLimit::Fields { observed, maximum } => {
            limit_error(SlideTableCellsLimitKind::WireFields, observed, maximum)
        },
        storage::DecodeLimit::Work { observed, maximum } => {
            limit_error(SlideTableCellsLimitKind::WireWork, observed, maximum)
        },
        storage::DecodeLimit::Nesting { observed, maximum } => limit_error(
            SlideTableCellsLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        storage::DecodeLimit::Allocation { requested } => {
            SlideTableCellsError::Allocation { amount: requested }
        },
        _ => SlideTableCellsError::InvalidSource,
    }
}

fn map_cell_value_error(
    _error: litchi_numbers_wire::cell_value::DecodeError,
) -> SlideTableCellsError {
    SlideTableCellsError::InvalidSource
}

fn map_common_wire_error(error: litchi_iwa_common::Error) -> SlideTableCellsError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => limit_error(
            match kind {
                litchi_iwa_common::LimitKind::InputBytes => SlideTableCellsLimitKind::InputBytes,
                litchi_iwa_common::LimitKind::Fields => SlideTableCellsLimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => SlideTableCellsLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Nesting => SlideTableCellsLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SlideTableCellsLimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SlideTableCellsLimitKind::PayloadItems
                },
            },
            observed,
            limit,
        ),
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideTableCellsError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => SlideTableCellsError::InvalidSource,
    }
}

fn map_sidecar_issue(issue: SidecarIssue) -> SlideTableCellsError {
    match issue {
        SidecarIssue::Allocation { amount, .. } => SlideTableCellsError::Allocation { amount },
        SidecarIssue::Coordinator(_)
        | SidecarIssue::StorageDecode { .. }
        | SidecarIssue::CommentDecode(_)
        | SidecarIssue::FormulaDecode(_)
        | SidecarIssue::InvalidEntry { .. }
        | SidecarIssue::ZeroReference { .. }
        | SidecarIssue::MissingPayload { .. }
        | SidecarIssue::DuplicatePayload { .. }
        | SidecarIssue::DuplicateKey { .. }
        | SidecarIssue::DuplicateSegmentReference { .. }
        | SidecarIssue::DuplicateList { .. }
        | SidecarIssue::InvalidReplies => SlideTableCellsError::InvalidSource,
        _ => SlideTableCellsError::InvalidSource,
    }
}

fn map_common_table_error(error: litchi_iwa_common::table::model::Error) -> SlideTableCellsError {
    match error {
        litchi_iwa_common::table::model::Error::Allocation { amount, .. } => {
            SlideTableCellsError::Allocation { amount }
        },
        litchi_iwa_common::table::model::Error::BudgetExceeded { requested, maximum } => {
            limit_error(SlideTableCellsLimitKind::PayloadItems, requested, maximum)
        },
        _ => SlideTableCellsError::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_common::table::cell::value::Value;

    #[test]
    fn native_slide_table_read_resolves_values_comments_and_empty_replies() {
        let source =
            include_bytes!("../../../../test-data/iwork/keynote/slide-table-read-native.key");
        let package = Package::from_bytes(source).expect("native Keynote table fixture");
        let read = package
            .slide_table_cells(0, 0)
            .expect("focused Keynote table read");

        assert_eq!(read.dimensions(), Dimensions::new(5, 4));
        assert_eq!(
            read.get_a1("A1").unwrap(),
            Some(&Value::Text("Buffa discovery\n".into()))
        );
        assert_eq!(
            read.get_a1("A2").unwrap(),
            Some(&Value::Text("Café 北京".into()))
        );
        let number = Value::number(12.5).unwrap();
        assert_eq!(read.get_a1("B2").unwrap(), Some(&number));
        assert_eq!(read.get_a1("C2").unwrap(), Some(&Value::Boolean(true)));
        assert_eq!(
            read.get_a1("D2").unwrap(),
            Some(&Value::Formula("=(B2+1)".into()))
        );

        let comment = read.get_comment_a1("D2").unwrap().expect("formula comment");
        assert_eq!(comment.text(), "Focused Keynote table read — Café 北京");
        assert_eq!(
            comment.author().and_then(CommentAuthor::display_name),
            Some("Ryker Zhu")
        );
        assert!(comment.timestamp().is_some());
        assert_eq!(comment.replies(), Some([].as_slice()));
    }

    #[test]
    fn materialized_cell_ceiling_is_inclusive_and_cumulative() {
        let source =
            include_bytes!("../../../../test-data/iwork/keynote/slide-table-read-native.key");
        let package = Package::from_bytes(source).expect("native Keynote table fixture");
        let mut budget = ReadBudget::new(&package).expect("focused read budget");
        let first = MAX_MATERIALIZED_CELLS - 257;
        let second = 257;
        let total = first.checked_add(second).expect("test count fits usize");
        assert_eq!(total, MAX_MATERIALIZED_CELLS);
        TableCellReadBudget::check_materialized_cells(&mut budget, first)
            .expect("first bounded batch");
        TableCellReadBudget::check_materialized_cells(&mut budget, total)
            .expect("legacy ceiling is inclusive");
        assert_eq!(budget.materialized_cells, MAX_MATERIALIZED_CELLS);

        let error = TableCellReadBudget::check_materialized_cells(&mut budget, total + 1)
            .expect_err("the next cumulative cell must be refused");
        assert_eq!(
            error,
            SlideTableCellsError::LimitExceeded {
                kind: SlideTableCellsLimitKind::PayloadItems,
                observed: usize_as_u64(MAX_MATERIALIZED_CELLS + 1),
                maximum: usize_as_u64(MAX_MATERIALIZED_CELLS),
            }
        );
    }

    #[test]
    fn unknown_formula_owner_is_refused_instead_of_getting_a_synthetic_prefix() {
        let resolver = FormulaNameResolver {
            table_names: vec!["Table 1".into()],
            owners: HashMap::new(),
            categories: HashMap::new(),
            unresolved_table: FlagCell::new(false),
            unresolved_category: FlagCell::new(false),
        };
        let owner = numbers_formula_codec::FormulaRenderCfuuid {
            has_uuid_bytes: false,
            word0: Some(1),
            word1: Some(2),
            word2: Some(3),
            word3: Some(4),
        };
        assert!(resolver.table_prefix(&owner).is_none());
        assert!(resolver.has_unresolved_reference());
    }

    #[test]
    fn rooted_formula_names_render_table_only_and_known_categories() {
        let resolver = FormulaNameResolver {
            table_names: vec!["Table 1".into()],
            owners: HashMap::from([([1, 2, 3, 4], 0)]),
            categories: HashMap::from([([9, 8], "Region".into())]),
            unresolved_table: FlagCell::new(false),
            unresolved_category: FlagCell::new(false),
        };
        let owner = numbers_formula_codec::FormulaRenderCfuuid {
            has_uuid_bytes: true,
            word0: Some(1),
            word1: Some(2),
            word2: Some(3),
            word3: Some(4),
        };
        assert_eq!(resolver.table_only_name(&owner), Some("Table 1"));
        assert_eq!(
            resolver.category_name(FormulaCategoryId { lower: 9, upper: 8 }),
            Some("Region")
        );
        assert_eq!(
            resolver.category_name(FormulaCategoryId { lower: 1, upper: 0 }),
            Some("Grand Total")
        );
        assert!(!resolver.has_unresolved_reference());
    }

    #[test]
    fn rooted_native_table_catalog_proves_table_info_name() {
        let source =
            include_bytes!("../../../../test-data/iwork/keynote/slide-table-read-native.key");
        let package = Package::from_bytes(source).expect("native Keynote table fixture");
        let mut budget = ReadBudget::new(&package).expect("focused read budget");
        let names = core::rooted_table_name_catalog(&package, &mut budget.budget)
            .expect("rooted table catalog");
        assert_eq!(names.len(), 1);
        assert_eq!(names[0].name.as_ref(), "Table 1");
        assert_ne!(names[0].table_info_identifier, 0);
    }

    #[test]
    fn rich_text_reference_preflight_rejects_duplicate_storage_fields() {
        let source = [
            0x0a, 0x02, 0x08, 0x07, // field 1: storage ref
            0x0a, 0x02, 0x08, 0x08, // duplicate field 1
            0x1a, 0x01, 0x00, // field 3: cell owner marker
        ];
        assert!(matches!(
            preflight_rich_text_payload(&source, WireLimits::default()),
            Err(SlideTableCellsError::InvalidSource)
        ));
    }

    #[test]
    fn rich_text_storage_rejects_cross_kind_duplicates_in_either_order() {
        for messages in [
            [
                RawMessage {
                    type_: RICH_TEXT_STORAGE_MESSAGE_KINDS[0],
                    data: vec![1],
                },
                RawMessage {
                    type_: RICH_TEXT_STORAGE_MESSAGE_KINDS[1],
                    data: vec![2],
                },
            ],
            [
                RawMessage {
                    type_: RICH_TEXT_STORAGE_MESSAGE_KINDS[1],
                    data: vec![1],
                },
                RawMessage {
                    type_: RICH_TEXT_STORAGE_MESSAGE_KINDS[0],
                    data: vec![2],
                },
            ],
        ] {
            assert!(matches!(
                KeynoteCellSidecarResolver::canonical_message_kinds(
                    NativeMessages {
                        messages: &messages
                    },
                    &RICH_TEXT_STORAGE_MESSAGE_KINDS,
                ),
                Err(SlideTableCellsError::InvalidSource)
            ));
        }
    }

    #[test]
    fn formula_envelope_budget_covers_recursive_wire_scan() {
        let source =
            include_bytes!("../../../../test-data/iwork/keynote/slide-table-read-native.key");
        let package = Package::from_bytes(source).expect("native Keynote table fixture");
        let mut budget = ReadBudget::new(&package).expect("focused read budget");
        let selected = selected_slide_table_storage(
            &package,
            SlideSelector::position(Position::new(0)),
            TableSelector::position(Position::new(0)),
            &mut budget.budget,
        )
        .expect("selected native table");
        let formulas = read_one_sidecar(
            &package,
            SidecarKind::Formulas,
            selected.data_store().formula_table().identifier(),
            &mut budget,
        )
        .expect("native formula sidecar");
        let formula = formulas
            .iter()
            .find_map(|(_key, value)| value.formula())
            .expect("native formula source");
        let limits = budget
            .formula_envelope_limits(formula)
            .expect("residual formula limits");
        assert!(limits.max_input_bytes > formula.len());
        let report = formula_envelope::preflight_formula_envelope(formula, limits)
            .expect("recursive formula envelope");
        assert!(report.scanned_bytes() > formula.len());
    }

    #[test]
    fn formula_render_budget_accepts_exact_residual_output() {
        let source =
            include_bytes!("../../../../test-data/iwork/keynote/slide-table-read-native.key");
        let package = Package::from_bytes(source).expect("native Keynote table fixture");
        let mut budget = ReadBudget::new(&package).expect("focused read budget");
        FormulaRenderBudget::charge(&mut budget, 1).expect("initial output byte");
        let remaining = budget
            .budget
            .remaining_output()
            .expect("remaining output budget");
        FormulaRenderBudget::check(&budget, remaining)
            .expect("exact residual output is still admissible");
        FormulaRenderBudget::charge(&mut budget, remaining).expect("exact residual output charge");
        assert_eq!(budget.output_bytes, remaining.saturating_add(1));
        assert!(budget.budget.remaining_output().is_err());
    }
}
