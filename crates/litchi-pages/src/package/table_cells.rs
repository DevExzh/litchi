//! Bounded, selector-first Pages body-table cell readback.
//!
//! Pages owns the rooted attachment/model proof and the archive lookup in this
//! module.  The selected model and DataStore remain borrowed from the source;
//! the shared Numbers-wire reader owns only tile/row topology and cell-value
//! classification.  Native identifiers are kept in this private adapter and
//! are never part of the semantic API.

use std::cell::Cell as StateCell;
use std::collections::HashMap;
use std::fmt;
use std::fmt::Write as _;
use std::mem::size_of;
use std::num::NonZeroU64;

use litchi_iwa_common::formula::render::FormulaRenderBudget;
use litchi_iwa_common::table::coordinate::CellPosition;
use litchi_iwa_common::table::model::{Cell as TableCell, Dimensions};
use litchi_iwa_common::table::read::{
    CellComment, Comment, CommentAuthor, CommentReply, TableRead,
};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;
use litchi_numbers_wire::formula_names;
use litchi_numbers_wire::formula_render::{
    FormulaCategoryId, FormulaEventRenderBudget, FormulaTablePrefix, ReferenceResolver,
};
use litchi_numbers_wire::table_cells::{self as wire_cells, CellValueSink, TableCellReadBudget};
use litchi_numbers_wire::table_sidecars::{
    self as sidecars, CellSidecarResolver, SidecarAllocation, SidecarIssue, SidecarKind,
    SidecarReadBudget, SidecarReference, SidecarTables,
};
use thiserror::Error;

use super::{Package, table_lock};
use crate::selector::BodyTableSelector;

/// Finite resources charged while reading one Pages body-table cell source.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableCellsLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete package output bytes (reserved for shared transaction budgets).
    OutputBytes,
    /// ZIP members inspected while resolving the rooted table.
    Entries,
    /// Bytes retained by one ZIP member.
    EntryBytes,
    /// Aggregate ZIP member bytes.
    TotalEntryBytes,
    /// ZIP metadata bytes.
    PackageBytes,
    /// Bytes in one decoded native payload.
    PayloadBytes,
    /// Aggregate decoded native payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected.
    PayloadObjects,
    /// Native payload messages inspected.
    PayloadMessages,
    /// Native payload metadata items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict table-cell wire bytes.
    WireBytes,
    /// Strict table-cell wire output bytes (unused by this read).
    WireOutputBytes,
    /// Strict table-cell wire fields.
    WireFields,
    /// Strict table-cell wire nesting.
    WireNesting,
    /// Aggregate table-cell wire work.
    WireWork,
}

impl fmt::Display for BodyTableCellsLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "ZIP entries",
            Self::EntryBytes => "ZIP entry bytes",
            Self::TotalEntryBytes => "total ZIP entry bytes",
            Self::PackageBytes => "package metadata bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure while resolving or decoding one Pages body-table cell source.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableCellsError {
    /// No rooted body table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one rooted body table matched an exact name.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The rooted selector or ownership graph is ambiguous.
    #[error("the Pages body-table cell selector is ambiguous")]
    AmbiguousSelector,
    /// The source is not admitted by the native Pages table-cell profile.
    #[error("the Pages package source does not support body-table cell reads")]
    UnsupportedSource,
    /// The selected rooted graph or storage payload is malformed.
    #[error("the selected Pages body-table cell source is invalid")]
    InvalidSource,
    /// A finite read resource ceiling was exceeded.
    #[error("Pages body-table cells {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: BodyTableCellsLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded staging allocation failed.
    #[error("could not allocate {amount} units for Pages body-table cells")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
}

/// Borrowed strict model/DataStore proof for one selected Pages table.
///
/// The proof is crate-private because its fields include native routing state.
/// It is deliberately separate from the semantic result: a caller can pass
/// this source-backed projection to the shared cell reader without exposing
/// package objects or generated protobuf values.
pub(crate) struct SelectedBodyTableStorage<'source> {
    model: storage::TableModelSnapshot<'source>,
    data_store: storage::DataStoreSnapshot<'source>,
    tile_size: u32,
    tiles: Vec<wire_cells::TileReference>,
}

impl<'source> SelectedBodyTableStorage<'source> {
    /// Return the selected table's checked dimensions for the shared reader.
    #[must_use]
    pub(crate) fn dimensions(&self) -> wire_cells::TableDimensions {
        wire_cells::TableDimensions::new(
            self.model.number_of_rows(),
            self.model.number_of_columns(),
            self.tile_size,
        )
    }

    /// Borrow the model snapshot without exposing it outside this crate.
    #[must_use]
    pub(crate) const fn model(&self) -> storage::TableModelSnapshot<'source> {
        self.model
    }

    /// Borrow the DataStore snapshot without exposing it outside this crate.
    #[must_use]
    pub(crate) const fn data_store(&self) -> storage::DataStoreSnapshot<'source> {
        self.data_store
    }

    /// Feed all selected cells to a Pages-owned semantic sink.
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
        let tile_references = self.tiles.iter().copied();
        wire_cells::read_table_cells(
            self.dimensions(),
            tile_references,
            resolve_tile,
            budget,
            sink,
        )
    }
}

/// A bounded package-local object index used by all selected sidecar and tile
/// lookups in one read.  The index retains only borrowed message slices; it
/// does not clone native objects or protobuf payloads.  Building it once also
/// prevents a table with many tiles from repeatedly scanning every component.
struct ObjectIndex<'source> {
    entries: Vec<(NonZeroU64, &'source [RawMessage])>,
}

impl<'source> ObjectIndex<'source> {
    fn new(
        package: &'source Package,
        budget: &mut table_lock::WireBudget,
    ) -> Result<Self, BodyTableCellsError> {
        let components = package.state.source.components();
        let object_count = components
            .iter()
            .try_fold(0usize, |count, component| {
                count.checked_add(component.archive().objects.len())
            })
            .ok_or(BodyTableCellsError::InvalidSource)?;
        budget
            .charge_payload_work(object_count)
            .map_err(map_lock_error)?;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(object_count)
            .map_err(|_| BodyTableCellsError::Allocation {
                amount: object_count,
            })?;
        for component in components.iter() {
            for object in &component.archive().objects {
                budget
                    .charge_payload_work(object.messages.len())
                    .map_err(map_lock_error)?;
                let Some(identifier) = object.archive_info.identifier else {
                    continue;
                };
                let Some(identifier) = NonZeroU64::new(identifier) else {
                    return Err(BodyTableCellsError::InvalidSource);
                };
                entries.push((identifier, object.messages.as_slice()));
            }
        }
        entries.sort_unstable_by_key(|(identifier, _)| *identifier);
        if entries.windows(2).any(|pair| pair[0].0 == pair[1].0) {
            return Err(BodyTableCellsError::InvalidSource);
        }
        Ok(Self { entries })
    }

    fn resolve(
        &self,
        identifier: u64,
    ) -> Result<Option<NativeMessages<'source>>, BodyTableCellsError> {
        let Some(identifier) = NonZeroU64::new(identifier) else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        Ok(self
            .entries
            .binary_search_by_key(&identifier, |(key, _)| *key)
            .ok()
            .map(|index| NativeMessages {
                messages: self.entries[index].1,
            }))
    }
}

/// Borrowed message candidates for one physical object.
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

const MAX_MATERIALIZED_CELLS: usize = 1_000_000;
const MAX_SIDECAR_ENTRIES: usize = 1_000_000;

/// Pages' aggregate ledger adapter for the shared cell and sidecar readers.
///
/// Formula output is tracked locally because the table-lock wire budget keeps
/// output and input counters separate.  All other counters are charged
/// directly to the same [`table_lock::WireBudget`] used by rooted selection
/// and model/DataStore admission.
struct PagesCellReadBudget {
    wire: table_lock::WireBudget,
    materialized_cells: usize,
    output_bytes: usize,
}

impl PagesCellReadBudget {
    fn new(wire: table_lock::WireBudget) -> Self {
        Self {
            wire,
            materialized_cells: 0,
            output_bytes: 0,
        }
    }

    fn wire(&self) -> &table_lock::WireBudget {
        &self.wire
    }

    fn wire_mut(&mut self) -> &mut table_lock::WireBudget {
        &mut self.wire
    }
}

impl TableCellReadBudget for PagesCellReadBudget {
    type Error = BodyTableCellsError;

    fn storage_options(&mut self, source: &[u8]) -> Result<storage::DecodeOptions, Self::Error> {
        storage_options(&self.wire, source)
    }

    fn charge_storage_report(&mut self, report: storage::DecodeReport) -> Result<(), Self::Error> {
        charge_storage_report(&mut self.wire, report)
    }

    fn check_materialized_cells(&mut self, observed: usize) -> Result<(), Self::Error> {
        if observed > MAX_MATERIALIZED_CELLS {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::PayloadItems,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(MAX_MATERIALIZED_CELLS),
            });
        }
        let additional = observed.saturating_sub(self.materialized_cells);
        self.wire
            .charge_payload_items(additional)
            .map_err(map_lock_error)?;
        self.wire
            .charge_payload_work(additional)
            .map_err(map_lock_error)?;
        self.materialized_cells = observed;
        Ok(())
    }

    fn charge_cell_source(&mut self, bytes: usize) -> Result<(), Self::Error> {
        self.wire.charge_payload_work(bytes).map_err(map_lock_error)
    }

    fn charge_allocation(
        &mut self,
        _target: wire_cells::AllocationTarget,
        amount: usize,
    ) -> Result<(), Self::Error> {
        self.wire
            .charge_payload_work(amount)
            .map_err(map_lock_error)
    }

    fn map_issue(&mut self, issue: wire_cells::TableCellIssue) -> Self::Error {
        map_table_cell_issue(issue)
    }
}

impl SidecarReadBudget for PagesCellReadBudget {
    type Error = BodyTableCellsError;

    fn list_options(
        &mut self,
        _kind: SidecarKind,
        source: &[u8],
    ) -> Result<storage::DecodeOptions, Self::Error> {
        storage_options(&self.wire, source)
    }

    fn charge_list_report(
        &mut self,
        _kind: SidecarKind,
        report: storage::DecodeReport,
    ) -> Result<(), Self::Error> {
        charge_storage_report(&mut self.wire, report)
    }

    fn charge_list_probe_report(
        &mut self,
        _kind: SidecarKind,
        report: storage::DecodeReport,
    ) -> Result<(), Self::Error> {
        charge_storage_probe_report(&mut self.wire, report)
    }

    fn charge_retained(
        &mut self,
        _kind: SidecarKind,
        target: SidecarAllocation,
        amount: usize,
    ) -> Result<(), Self::Error> {
        match target {
            SidecarAllocation::Keys
            | SidecarAllocation::Values
            | SidecarAllocation::Segments
            | SidecarAllocation::Comments => {
                self.wire
                    .charge_payload_items(amount)
                    .map_err(map_lock_error)?;
                self.wire
                    .charge_payload_work(amount)
                    .map_err(map_lock_error)?;
            },
            SidecarAllocation::Text | SidecarAllocation::FormulaBytes => {
                self.wire
                    .charge_payload_work(amount)
                    .map_err(map_lock_error)?;
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
        self.wire
            .charge_payload_work(source_bytes)
            .map_err(map_lock_error)?;
        self.wire.charge_payload_items(1).map_err(map_lock_error)
    }

    fn max_entries(&self, _kind: SidecarKind) -> usize {
        MAX_SIDECAR_ENTRIES.min(self.wire.remaining_wire_fields())
    }

    fn map_issue(&mut self, issue: SidecarIssue) -> Self::Error {
        map_sidecar_issue(issue)
    }

    fn comment_options(
        &mut self,
        source: &[u8],
    ) -> Result<litchi_iwa_protos::comment_storage_codec::DecodeOptions, Self::Error> {
        let limits = self.wire.wire_limits();
        let input = source.len().max(1);
        if input > limits.max_input_bytes() {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireBytes,
                observed: usize_as_u64(input),
                maximum: usize_as_u64(limits.max_input_bytes()),
            });
        }
        let fields = self.wire.remaining_wire_fields();
        let work = self.wire.remaining_wire_work();
        if fields == 0 {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireFields,
                observed: 1,
                maximum: 0,
            });
        }
        if work == 0 {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireWork,
                observed: 1,
                maximum: 0,
            });
        }
        Ok(
            litchi_iwa_protos::comment_storage_codec::DecodeOptions::new(
                input,
                fields,
                work,
                u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
                self.wire.remaining_payload_references(),
                work.min(limits.max_input_bytes()),
            ),
        )
    }

    fn charge_comment_report(
        &mut self,
        report: litchi_iwa_protos::comment_storage_codec::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.wire
            .charge_payload_work(report.source_bytes())
            .map_err(map_lock_error)?;
        self.wire
            .charge_codec_report(
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            )
            .map_err(map_lock_error)?;
        self.wire
            .charge_payload_work(report.reference_bytes())
            .map_err(map_lock_error)?;
        self.wire
            .charge_payload_work(report.text_bytes())
            .map_err(map_lock_error)?;
        self.wire
            .charge_payload_items(report.replies())
            .map_err(map_lock_error)
    }

    fn formula_options(
        &mut self,
        source: &[u8],
    ) -> Result<litchi_iwa_protos::numbers_formula_codec::DecodeOptions, Self::Error> {
        let limits = self.wire.wire_limits();
        let input = source.len().max(1);
        if input > limits.max_input_bytes() {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireBytes,
                observed: usize_as_u64(input),
                maximum: usize_as_u64(limits.max_input_bytes()),
            });
        }
        let fields = self.wire.remaining_wire_fields();
        let work = self.wire.remaining_wire_work();
        if fields == 0 {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireFields,
                observed: 1,
                maximum: 0,
            });
        }
        if work == 0 {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireWork,
                observed: 1,
                maximum: 0,
            });
        }
        Ok(
            litchi_iwa_protos::numbers_formula_codec::DecodeOptions::new(
                input,
                fields,
                work,
                u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
                fields,
                work.min(limits.max_input_bytes()),
            )
            .with_opaque_unknown_fields(true)
            .with_unknown_functions(true)
            .with_render_recursion_limit(u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX)),
        )
    }

    fn charge_formula_report(
        &mut self,
        report: litchi_iwa_protos::numbers_formula_codec::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.wire
            .charge_payload_work(report.bytes())
            .map_err(map_lock_error)?;
        self.wire
            .charge_codec_report(report.fields(), report.work(), report.max_depth(), 0)
            .map_err(map_lock_error)?;
        self.wire
            .charge_payload_work(report.text_bytes())
            .map_err(map_lock_error)
    }

    fn formula_envelope_limits(
        &mut self,
        source: &[u8],
    ) -> Result<litchi_numbers_wire::formula_envelope::FormulaEnvelopeLimits, Self::Error> {
        let limits = self.wire.wire_limits();
        let input = source.len().max(1);
        if input > limits.max_input_bytes() {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireBytes,
                observed: usize_as_u64(input),
                maximum: usize_as_u64(limits.max_input_bytes()),
            });
        }
        let max_fields = limits.max_fields();
        let max_work = limits.max_rewrite_work();
        // Formula-envelope preflight charges the root and every descended
        // child into the same wire-work ledger. Its input total can therefore
        // exceed the root slice length while each nested payload remains
        // individually bounded. Use the remaining aggregate work as the
        // finite scan ceiling instead of the single-message input ceiling.
        let max_input_bytes = self
            .wire
            .remaining_wire_work()
            .max(input)
            .clamp(1, litchi_iwa_common::WireLimits::MAX_INPUT_BYTES);
        Ok(
            litchi_numbers_wire::formula_envelope::FormulaEnvelopeLimits {
                max_fields,
                max_input_bytes,
                max_work,
                base_fields: max_fields.saturating_sub(self.wire.remaining_wire_fields()),
                base_work: max_work.saturating_sub(self.wire.remaining_wire_work()),
            },
        )
    }

    fn charge_formula_envelope_report(
        &mut self,
        report: litchi_numbers_wire::formula_envelope::FormulaEnvelopeReport,
    ) -> Result<(), Self::Error> {
        let Some(wire) = report.wire_preflight() else {
            return Ok(());
        };
        self.wire
            .charge_payload_work(wire.scanned_bytes())
            .map_err(map_lock_error)?;
        self.wire
            .charge_codec_report(
                wire.fields(),
                0,
                u32::try_from(wire.max_depth()).unwrap_or(u32::MAX),
                0,
            )
            .map_err(map_lock_error)
    }

    fn retain_formula_envelope_cost(
        &mut self,
        cost: litchi_numbers_wire::formula_envelope::AttemptedFormulaEnvelopeCost,
    ) {
        // A rejected envelope is never published, but its bounded attempted
        // scan remains part of this transaction's monotonic accounting.  The
        // ledger cannot return a second error while the original decoder
        // error is being propagated, so deliberately ignore a repeated
        // ceiling failure here.
        let _ = self.wire.charge_payload_work(cost.work());
        let _ = self.wire.charge_codec_report(cost.fields(), 0, 0, 0);
    }

    fn map_formula_envelope_error(&mut self, error: litchi_iwa_common::Error) -> Self::Error {
        map_common_wire_error(error)
    }
}

impl FormulaRenderBudget for PagesCellReadBudget {
    type Error = BodyTableCellsError;

    fn output_limit(&self, observed: usize) -> Self::Error {
        BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireOutputBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(self.wire.wire_limits().max_output_bytes()),
        }
    }

    fn allocation(&self, _resource: &'static str, amount: usize) -> Self::Error {
        BodyTableCellsError::Allocation { amount }
    }

    fn invalid(&self, _message: &'static str) -> Self::Error {
        BodyTableCellsError::InvalidSource
    }

    fn check(&self, amount: usize) -> Result<(), Self::Error> {
        let Some(observed) = self.output_bytes.checked_add(amount) else {
            return Err(self.output_limit(usize::MAX));
        };
        if observed > self.wire.wire_limits().max_output_bytes() {
            return Err(self.output_limit(observed));
        }
        Ok(())
    }

    fn check_structure(&self, nodes: usize, parts: usize) -> Result<(), Self::Error> {
        let maximum = self.wire.remaining_wire_work();
        if nodes > maximum || parts > maximum {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireWork,
                observed: nodes.max(parts) as u64,
                maximum: usize_as_u64(maximum),
            });
        }
        Ok(())
    }

    fn charge(&mut self, amount: usize) -> Result<(), Self::Error> {
        self.check(amount)?;
        self.wire
            .charge_output_bytes(amount)
            .map_err(map_lock_error)?;
        self.output_bytes = self.output_bytes.saturating_add(amount);
        Ok(())
    }
}

impl FormulaEventRenderBudget for PagesCellReadBudget {
    fn check_render_depth(&self, depth: usize) -> Result<(), Self::Error> {
        if depth > self.wire.wire_limits().max_nesting() {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireNesting,
                observed: usize_as_u64(depth),
                maximum: usize_as_u64(self.wire.wire_limits().max_nesting()),
            });
        }
        Ok(())
    }

    fn parse_error(&self, _message: String) -> Self::Error {
        BodyTableCellsError::InvalidSource
    }

    fn invalid_format(&self, _message: String) -> Self::Error {
        BodyTableCellsError::InvalidSource
    }
}

/// Bounded formula names resolved from the rooted Pages body-table catalog.
///
/// The shared renderer only needs borrowed names while it visits one formula.
/// The adapter therefore retains table names and owner/category indexes, but
/// never publishes a native object identifier or a generated dependency
/// archive.  Unknown identities remain a hard read failure after rendering so
/// the renderer's compatibility fallback cannot become semantic data.
struct FormulaNameResolver {
    table_names: Vec<Box<str>>,
    owners: HashMap<[u32; 4], usize>,
    categories: HashMap<[u64; 2], Box<str>>,
    unresolved_table: StateCell<bool>,
    unresolved_category: StateCell<bool>,
}

impl FormulaNameResolver {
    fn has_unresolved_reference(&self) -> bool {
        self.unresolved_table.get() || self.unresolved_category.get()
    }
}

impl ReferenceResolver for FormulaNameResolver {
    fn table_prefix(
        &self,
        _id: &litchi_iwa_protos::numbers_formula_codec::FormulaRenderCfuuid,
    ) -> Option<FormulaTablePrefix<'_>> {
        self.unresolved_table.set(true);
        None
    }

    fn table_only_name(
        &self,
        id: &litchi_iwa_protos::numbers_formula_codec::FormulaRenderCfuuid,
    ) -> Option<&str> {
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
    fn build(
        package: &Package,
        index: &ObjectIndex<'_>,
        budget: &mut PagesCellReadBudget,
    ) -> Result<Self, BodyTableCellsError> {
        // The full rooted catalog is the authority for table display names.
        // A selected table alone is insufficient for a cross-table formula,
        // and no synthetic sheet name is valid for the Pages profile.
        let targets = table_lock::body_table_catalog_with_budget(package, budget.wire_mut())
            .map_err(map_lock_error)?;
        let target_count = targets.len();
        budget
            .wire_mut()
            .charge_payload_items(target_count)
            .map_err(map_lock_error)?;
        budget
            .wire_mut()
            .charge_payload_work(target_count)
            .map_err(map_lock_error)?;
        let mut table_names = Vec::new();
        table_names.try_reserve_exact(target_count).map_err(|_| {
            BodyTableCellsError::Allocation {
                amount: target_count,
            }
        })?;
        budget
            .wire_mut()
            .charge_payload_work(target_count)
            .map_err(map_lock_error)?;
        let mut table_indices = HashMap::new();
        table_indices
            .try_reserve(target_count)
            .map_err(|_| BodyTableCellsError::Allocation {
                amount: target_count,
            })?;
        for target in targets {
            let table_info = target.drawable_identifier.get();
            let index = table_names.len();
            if table_indices.insert(table_info, index).is_some() {
                return Err(BodyTableCellsError::InvalidSource);
            }
            table_names.push(target.table_name);
        }

        let mut owners = HashMap::new();
        let mut categories = HashMap::new();
        for (_, messages) in &index.entries {
            for message in *messages {
                budget
                    .wire_mut()
                    .charge_payload_work(1)
                    .map_err(map_lock_error)?;
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
                    // A valid owner can refer to a table outside the selected
                    // rooted body.  Keep it opaque; the formula renderer will
                    // fail closed if a cell actually points to that owner.
                    continue;
                };
                budget
                    .wire_mut()
                    .charge_payload_items(1)
                    .map_err(map_lock_error)?;
                owners
                    .try_reserve(1)
                    .map_err(|_| BodyTableCellsError::Allocation {
                        amount: owners.len().saturating_add(1),
                    })?;
                if owners
                    .insert(projection.cfuuid_words(), table_index)
                    .is_some()
                {
                    return Err(BodyTableCellsError::InvalidSource);
                }
            }
        }

        Ok(Self {
            table_names,
            owners,
            categories,
            unresolved_table: StateCell::new(false),
            unresolved_category: StateCell::new(false),
        })
    }
}

fn formula_name_wire_limits(
    budget: &PagesCellReadBudget,
    source_len: usize,
) -> Result<litchi_iwa_common::WireLimits, BodyTableCellsError> {
    let limits = budget.wire().wire_limits();
    let fields = budget.wire().remaining_wire_fields();
    let work = budget.wire().remaining_wire_work();
    if fields == 0 {
        return Err(BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireFields,
            observed: 1,
            maximum: 0,
        });
    }
    if work == 0 {
        return Err(BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let input = work.max(source_len);
    if input > litchi_iwa_common::WireLimits::MAX_INPUT_BYTES {
        return Err(BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireBytes,
            observed: usize_as_u64(input),
            maximum: usize_as_u64(litchi_iwa_common::WireLimits::MAX_INPUT_BYTES),
        });
    }
    litchi_iwa_common::WireLimits::default()
        .with_input_bytes(input.max(1))
        .and_then(|value| {
            value.with_fields(fields.clamp(1, litchi_iwa_common::WireLimits::MAX_FIELDS))
        })
        .and_then(|value| {
            value.with_nesting(
                limits
                    .max_nesting()
                    .min(litchi_iwa_common::WireLimits::MAX_NESTING),
            )
        })
        .and_then(|value| {
            value.with_rewrite_work(work.clamp(1, litchi_iwa_common::WireLimits::MAX_REWRITE_WORK))
        })
        .map_err(map_common_wire_error)
}

fn charge_formula_name_report(
    budget: &mut PagesCellReadBudget,
    input_bytes: usize,
    fields: usize,
    work: usize,
) -> Result<(), BodyTableCellsError> {
    budget
        .wire_mut()
        .charge_payload_work(input_bytes)
        .map_err(map_lock_error)?;
    budget
        .wire_mut()
        .charge_codec_report(fields, work, 0, 0)
        .map_err(map_lock_error)
}

fn read_formula_categories(
    source: &[u8],
    categories: &mut HashMap<[u64; 2], Box<str>>,
    budget: &mut PagesCellReadBudget,
) -> Result<(), BodyTableCellsError> {
    let limits = formula_name_wire_limits(budget, source.len())?;
    let category_limits = formula_names::CategoryReadLimits {
        wire: limits,
        max_nodes: budget
            .wire()
            .remaining_wire_fields()
            .clamp(1, litchi_iwa_common::WireLimits::MAX_FIELDS),
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
    budget
        .wire_mut()
        .charge_payload_items(projection.entries.len())
        .map_err(map_lock_error)?;
    categories
        .try_reserve(projection.entries.len())
        .map_err(|_| BodyTableCellsError::Allocation {
            amount: projection.entries.len(),
        })?;
    for entry in projection.entries {
        let reserve = entry.label().formatted_len();
        budget
            .wire_mut()
            .charge_payload_work(reserve)
            .map_err(map_lock_error)?;
        let mut label = String::new();
        label
            .try_reserve_exact(reserve)
            .map_err(|_| BodyTableCellsError::Allocation { amount: reserve })?;
        write!(&mut label, "{}", entry.label()).map_err(|_| BodyTableCellsError::InvalidSource)?;
        let key = entry.id();
        if categories.insert(key, label.into_boxed_str()).is_some() {
            return Err(BodyTableCellsError::InvalidSource);
        }
    }
    Ok(())
}

fn read_sidecars(
    index: &ObjectIndex<'_>,
    data_store: storage::DataStoreSnapshot<'_>,
    budget: &mut PagesCellReadBudget,
) -> Result<SidecarTables, BodyTableCellsError> {
    let mut tables = SidecarTables::new();
    let required = [
        (SidecarKind::Strings, data_store.string_table().identifier()),
        (
            SidecarKind::Formulas,
            data_store.formula_table().identifier(),
        ),
    ];
    for (kind, identifier) in required {
        let list = read_one_sidecar(index, kind, identifier, budget)?;
        tables
            .insert(list)
            .map_err(|issue| SidecarReadBudget::map_issue(budget, issue))?;
    }
    if let Some(reference) = data_store.formula_error_table() {
        let list = read_one_sidecar(
            index,
            SidecarKind::FormulaErrors,
            reference.identifier(),
            budget,
        )?;
        tables
            .insert(list)
            .map_err(|issue| SidecarReadBudget::map_issue(budget, issue))?;
    }
    if let Some(reference) = data_store.rich_text_table() {
        let list = read_one_sidecar(
            index,
            SidecarKind::RichTextPayloads,
            reference.identifier(),
            budget,
        )?;
        tables
            .insert(list)
            .map_err(|issue| SidecarReadBudget::map_issue(budget, issue))?;
    }
    if let Some(reference) = data_store.comment_storage_table() {
        let list = read_one_sidecar(index, SidecarKind::Comments, reference.identifier(), budget)?;
        tables
            .insert(list)
            .map_err(|issue| SidecarReadBudget::map_issue(budget, issue))?;
    }
    Ok(tables)
}

fn read_one_sidecar(
    index: &ObjectIndex<'_>,
    kind: SidecarKind,
    identifier: u64,
    budget: &mut PagesCellReadBudget,
) -> Result<sidecars::SidecarList, BodyTableCellsError> {
    let Some(root) = index.resolve(identifier)? else {
        return Err(SidecarReadBudget::map_issue(
            budget,
            SidecarIssue::MissingPayload { kind },
        ));
    };
    sidecars::read_sidecar_list(
        kind,
        identifier,
        root,
        |segment_id, _budget| index.resolve(segment_id),
        budget,
    )
}

struct PagesCellSink<'sidecars, 'source, 'names> {
    table: litchi_iwa_common::table::model::Builder,
    comments: Vec<CellComment>,
    sidecars: &'sidecars SidecarTables,
    index: &'sidecars ObjectIndex<'source>,
    formula_names: Option<&'names FormulaNameResolver>,
    rows: u32,
    columns: u32,
}

impl<'sidecars, 'source, 'names> PagesCellSink<'sidecars, 'source, 'names> {
    fn new(
        name: String,
        dimensions: Dimensions,
        sidecars: &'sidecars SidecarTables,
        index: &'sidecars ObjectIndex<'source>,
        formula_names: Option<&'names FormulaNameResolver>,
    ) -> Self {
        let table = litchi_iwa_common::table::model::Builder::new(name, dimensions);
        let comments = Vec::new();
        Self {
            table,
            comments,
            sidecars,
            index,
            formula_names,
            rows: dimensions.rows(),
            columns: dimensions.columns(),
        }
    }

    fn finish(self) -> Result<TableRead, BodyTableCellsError> {
        let table = self.table.finish().map_err(map_common_table_error)?;
        TableRead::try_from_owned_parts(table, self.comments).map_err(map_common_table_error)
    }
}

fn retain_table_name(
    name: &str,
    budget: &mut PagesCellReadBudget,
) -> Result<String, BodyTableCellsError> {
    budget
        .wire_mut()
        .charge_payload_work(name.len())
        .map_err(map_lock_error)?;
    let mut retained = String::new();
    retained
        .try_reserve_exact(name.len())
        .map_err(|_| BodyTableCellsError::Allocation { amount: name.len() })?;
    retained.push_str(name);
    Ok(retained)
}

impl CellValueSink<PagesCellReadBudget> for PagesCellSink<'_, '_, '_> {
    fn visit_cell(
        &mut self,
        cell: wire_cells::CellSource<'_>,
        budget: &mut PagesCellReadBudget,
    ) -> Result<(), BodyTableCellsError> {
        let mut resolver = PagesCellSidecarResolver {
            index: self.index,
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
        budget
            .wire_mut()
            .charge_payload_work(1)
            .map_err(map_lock_error)?;
        self.table
            .push(TableCell::new(
                CellPosition::new(cell.row(), cell.column()),
                value,
            ))
            .map_err(|error| map_common_table_error(error.into_parts().0))?;
        if let Some(comment) = comment {
            budget
                .wire_mut()
                .charge_payload_items(1)
                .map_err(map_lock_error)?;
            self.comments
                .try_reserve(1)
                .map_err(|_| BodyTableCellsError::Allocation {
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

struct PagesCellSidecarResolver<'index, 'budget, 'source, 'names> {
    index: &'index ObjectIndex<'source>,
    budget: &'budget mut PagesCellReadBudget,
    formula_names: Option<&'names FormulaNameResolver>,
    rows: u32,
    columns: u32,
}

impl CellSidecarResolver for PagesCellSidecarResolver<'_, '_, '_, '_> {
    type Error = BodyTableCellsError;

    fn rich_text(&mut self, reference: SidecarReference) -> Result<String, Self::Error> {
        let Some(messages) = self.index.resolve(reference.identifier())? else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        let mut payload = None;
        for message in messages {
            if message.kind != RICH_TEXT_PAYLOAD_MESSAGE_KIND {
                continue;
            }
            if payload.replace(message.data).is_some() {
                return Err(BodyTableCellsError::InvalidSource);
            }
        }
        let Some(payload) = payload else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        let (storage_identifier, report) = preflight_rich_text_payload(payload)?;
        self.budget
            .wire_mut()
            .charge_payload_work(report.scanned_bytes())
            .map_err(map_lock_error)?;
        self.budget
            .wire_mut()
            .charge_codec_report(
                report.fields(),
                report.scanned_bytes(),
                u32::try_from(report.max_depth()).unwrap_or(u32::MAX),
                1,
            )
            .map_err(map_lock_error)?;
        let Some(messages) = self.index.resolve(storage_identifier)? else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        let mut storage_payload = None;
        for message in messages {
            if !matches!(message.kind, 2_001 | 2_022) {
                continue;
            }
            if storage_payload.replace(message.data).is_some() {
                return Err(BodyTableCellsError::InvalidSource);
            }
        }
        let Some(storage_payload) = storage_payload else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        self.budget
            .wire_mut()
            .charge_payload_work(storage_payload.len())
            .map_err(map_lock_error)?;
        let limits = litchi_iwa_text_wire::Limits::new(
            storage_payload.len().max(1),
            self.budget.wire().remaining_wire_fields().max(1),
            self.budget.wire().remaining_wire_fields().max(1),
            self.budget.wire().remaining_wire_work().max(1),
        )
        .map_err(|_| BodyTableCellsError::InvalidSource)?;
        let storage = litchi_iwa_text_wire::from_bytes_with_limits(storage_payload, limits)
            .map_err(|_| BodyTableCellsError::InvalidSource)?;
        self.budget.charge(storage.len())?;
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
            return Err(BodyTableCellsError::UnsupportedSource);
        };
        let rendered = sidecars::render_formula(
            source,
            1,
            row,
            column,
            self.rows,
            self.columns,
            resolver,
            self.budget,
        )?;
        if resolver.has_unresolved_reference() {
            return Err(BodyTableCellsError::UnsupportedSource);
        }
        Ok(rendered)
    }

    fn comment(&mut self, reference: SidecarReference) -> Result<Comment, Self::Error> {
        let Some(messages) = self.index.resolve(reference.identifier())? else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        let storage = sidecars::read_comment_storage(reference, messages, self.budget)?;
        let (text, timestamp, author_reference, replies, _storage_uuid) = storage.into_parts();
        let author = author_reference
            .map(|reference| self.author(reference))
            .transpose()?;
        if !replies.is_empty() {
            let reply_bytes = replies
                .len()
                .checked_mul(size_of::<CommentReply>())
                .ok_or(BodyTableCellsError::InvalidSource)?;
            self.budget
                .wire_mut()
                .charge_payload_work(reply_bytes)
                .map_err(map_lock_error)?;
            self.budget
                .wire_mut()
                .charge_payload_items(1)
                .map_err(map_lock_error)?;
        }
        let mut retained_replies = Vec::new();
        retained_replies
            .try_reserve_exact(replies.len())
            .map_err(|_| BodyTableCellsError::Allocation {
                amount: replies.len(),
            })?;
        for reply in replies {
            retained_replies.push(self.read_reply(reply)?);
        }
        Ok(Comment::from_owned_parts(
            text,
            timestamp,
            author,
            Some(retained_replies.into_boxed_slice()),
        ))
    }

    fn retain_text(
        &mut self,
        _kind: SidecarKind,
        _key: u32,
        source: &str,
    ) -> Result<String, Self::Error> {
        self.budget.charge(source.len())?;
        let mut value = String::new();
        value
            .try_reserve_exact(source.len())
            .map_err(|_| BodyTableCellsError::Allocation {
                amount: source.len(),
            })?;
        value.push_str(source);
        Ok(value)
    }

    fn missing(&mut self, _kind: SidecarKind, _key: u32) -> Self::Error {
        BodyTableCellsError::InvalidSource
    }

    fn invalid_scalar(&mut self) -> Self::Error {
        BodyTableCellsError::InvalidSource
    }
}

impl PagesCellSidecarResolver<'_, '_, '_, '_> {
    fn read_reply(
        &mut self,
        reference: SidecarReference,
    ) -> Result<CommentReply, BodyTableCellsError> {
        let Some(messages) = self.index.resolve(reference.identifier())? else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        let storage = sidecars::read_comment_storage(reference, messages, self.budget)?;
        let (text, timestamp, author_reference, replies, _storage_uuid) = storage.into_parts();
        if !replies.is_empty() {
            return Err(BodyTableCellsError::UnsupportedSource);
        }
        let author = author_reference
            .map(|reference| self.author(reference))
            .transpose()?;
        Ok(CommentReply::from_owned_parts(text, timestamp, author))
    }

    fn author(
        &mut self,
        reference: SidecarReference,
    ) -> Result<CommentAuthor, BodyTableCellsError> {
        let Some(messages) = self.index.resolve(reference.identifier())? else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        let mut payload = None;
        for message in messages {
            if message.kind != ANNOTATION_AUTHOR_MESSAGE_KIND {
                continue;
            }
            if payload.replace(message.data).is_some() {
                return Err(BodyTableCellsError::InvalidSource);
            }
        }
        let Some(payload) = payload else {
            return Err(BodyTableCellsError::InvalidSource);
        };
        let limits = self.budget.wire().wire_limits();
        let input = payload.len().max(1);
        if input > limits.max_input_bytes() {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireBytes,
                observed: usize_as_u64(input),
                maximum: usize_as_u64(limits.max_input_bytes()),
            });
        }
        let fields = self.budget.wire().remaining_wire_fields();
        let work = self.budget.wire().remaining_wire_work();
        if fields == 0 {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireFields,
                observed: 1,
                maximum: 0,
            });
        }
        if work == 0 {
            return Err(BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireWork,
                observed: 1,
                maximum: 0,
            });
        }
        let options = litchi_iwa_protos::annotation_author_codec::DecodeOptions::new(
            input,
            fields,
            work,
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            self.budget.wire().remaining_payload_references(),
            work.min(limits.max_input_bytes()),
            fields,
        );
        let (author, report) =
            litchi_iwa_protos::annotation_author_codec::decode_annotation_author_with_report(
                payload, options,
            )
            .map_err(map_author_decode_error)?;
        self.budget
            .wire_mut()
            .charge_payload_work(report.source_bytes())
            .map_err(map_lock_error)?;
        self.budget
            .wire_mut()
            .charge_codec_report(
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            )
            .map_err(map_lock_error)?;
        self.budget
            .wire_mut()
            .charge_payload_work(report.text_bytes())
            .map_err(map_lock_error)?;
        self.budget
            .wire_mut()
            .charge_payload_items(report.allocations())
            .map_err(map_lock_error)?;
        let owned_strings = usize::from(author.name().is_some_and(|value| !value.is_empty()))
            + usize::from(author.public_id().is_some_and(|value| !value.is_empty()));
        self.budget
            .wire_mut()
            .charge_payload_items(owned_strings)
            .map_err(map_lock_error)?;
        CommentAuthor::try_new(author.name(), author.public_id()).map_err(map_common_table_error)
    }
}

const RICH_TEXT_PAYLOAD_MESSAGE_KIND: u32 = 6_218;
const ANNOTATION_AUTHOR_MESSAGE_KIND: u32 = 212;

fn preflight_rich_text_payload(
    source: &[u8],
) -> Result<(u64, litchi_iwa_common::wire::WirePreflight), BodyTableCellsError> {
    use litchi_iwa_common::wire::{WireDescent, preflight_wire_tree_with_limits};

    let limits = litchi_iwa_common::WireLimits::default()
        .with_input_bytes(
            source
                .len()
                .clamp(1, litchi_iwa_common::WireLimits::MAX_INPUT_BYTES),
        )
        .and_then(|limits| {
            limits.with_fields(
                source
                    .len()
                    .clamp(1, litchi_iwa_common::WireLimits::MAX_FIELDS),
            )
        })
        .and_then(|limits| limits.with_nesting(1))
        .map_err(|_| BodyTableCellsError::InvalidSource)?;
    let mut storage = None;
    let mut has_cell_owner = false;
    let report = preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        match field.number() {
            1 => {
                if storage.is_some() || field.wire_type() != 2 {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "invalid rich-text storage reference".to_owned(),
                    ));
                }
                field.validate_canonical_framing()?;
                storage = Some(parse_local_reference(field.payload())?);
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
    .map_err(|_| BodyTableCellsError::InvalidSource)?;
    if !has_cell_owner {
        return Err(BodyTableCellsError::InvalidSource);
    }
    let storage = storage.ok_or(BodyTableCellsError::InvalidSource)?;
    Ok((storage, report))
}

fn parse_local_reference(source: &[u8]) -> Result<u64, litchi_iwa_common::Error> {
    use litchi_iwa_common::decode_varint_from_bytes;
    let mut identifier = None;
    for field in litchi_iwa_common::wire::WireView::parse(source)
        .map_err(|_| litchi_iwa_common::Error::InvalidFormat("invalid local reference".into()))?
        .fields()
    {
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

fn map_table_cell_issue(issue: wire_cells::TableCellIssue) -> BodyTableCellsError {
    match issue {
        wire_cells::TableCellIssue::StorageDecode(error) => map_storage_error(error),
        wire_cells::TableCellIssue::Allocation { amount, .. } => {
            BodyTableCellsError::Allocation { amount }
        },
        wire_cells::TableCellIssue::MissingTile { .. }
        | wire_cells::TableCellIssue::MissingTilePayload { .. }
        | wire_cells::TableCellIssue::DuplicateTilePayload { .. }
        | wire_cells::TableCellIssue::TileIndexOutOfBounds { .. }
        | wire_cells::TableCellIssue::TileRowOutOfBounds { .. }
        | wire_cells::TableCellIssue::RowCoordinateOverflow { .. }
        | wire_cells::TableCellIssue::TableRowOutOfBounds { .. }
        | wire_cells::TableCellIssue::DuplicateTileRow { .. }
        | wire_cells::TableCellIssue::CounterOverflow
        | wire_cells::TableCellIssue::CellStorageOutOfBounds { .. }
        | wire_cells::TableCellIssue::CellValueDecode { .. } => BodyTableCellsError::InvalidSource,
    }
}

fn map_sidecar_issue(issue: SidecarIssue) -> BodyTableCellsError {
    match issue {
        SidecarIssue::StorageDecode { error, .. } => map_storage_error(error),
        SidecarIssue::CommentDecode(error) => map_comment_decode_error(error),
        SidecarIssue::FormulaDecode(error) => map_formula_decode_error(error),
        SidecarIssue::Coordinator(issue) => map_coordinator_issue(issue),
        SidecarIssue::Allocation { amount, .. } => BodyTableCellsError::Allocation { amount },
        SidecarIssue::InvalidEntry { .. }
        | SidecarIssue::ZeroReference { .. }
        | SidecarIssue::MissingPayload { .. }
        | SidecarIssue::DuplicatePayload { .. }
        | SidecarIssue::DuplicateKey { .. }
        | SidecarIssue::DuplicateSegmentReference { .. }
        | SidecarIssue::DuplicateList { .. }
        | SidecarIssue::ZeroUuid
        | SidecarIssue::InvalidReplies => BodyTableCellsError::InvalidSource,
    }
}

fn map_coordinator_issue(
    issue: litchi_numbers_wire::table_data_list::CoordinatorIssue,
) -> BodyTableCellsError {
    use litchi_numbers_wire::table_data_list::CoordinatorIssue;
    match issue {
        CoordinatorIssue::EntryLimit { observed, maximum } => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::PayloadItems,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        CoordinatorIssue::Allocation { amount, .. } => BodyTableCellsError::Allocation { amount },
        CoordinatorIssue::MissingRoot { .. }
        | CoordinatorIssue::DuplicateRoot { .. }
        | CoordinatorIssue::MissingSegment { .. }
        | CoordinatorIssue::MissingSegmentPayload { .. }
        | CoordinatorIssue::DuplicateSegmentPayload { .. }
        | CoordinatorIssue::WrongSegmentType { .. }
        | CoordinatorIssue::KeyRangeOverflow { .. }
        | CoordinatorIssue::MissingKeyRange { .. }
        | CoordinatorIssue::EntryOutsideKeyRange { .. }
        | CoordinatorIssue::DuplicateEntryKey { .. } => BodyTableCellsError::InvalidSource,
    }
}

fn map_comment_decode_error(
    error: litchi_iwa_protos::comment_storage_codec::DecodeError,
) -> BodyTableCellsError {
    use litchi_iwa_protos::comment_storage_codec::DecodeLimit;
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::OutputBytes { observed, maximum }) => {
            BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireOutputBytes,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(maximum),
            }
        },
        Some(DecodeLimit::References { observed, maximum })
        | Some(DecodeLimit::Replies { observed, maximum })
        | Some(DecodeLimit::ReferenceBytes { observed, maximum }) => {
            BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::PayloadReferences,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(maximum),
            }
        },
        Some(DecodeLimit::Text { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::PayloadItems,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Fields { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Work { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        Some(DecodeLimit::Allocations { observed, maximum })
        | Some(DecodeLimit::Scratch { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Retained { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::PayloadBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(_) | None => BodyTableCellsError::InvalidSource,
    }
}

fn map_formula_decode_error(
    error: litchi_iwa_protos::numbers_formula_codec::DecodeError,
) -> BodyTableCellsError {
    use litchi_iwa_protos::numbers_formula_codec::DecodeLimit;
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Fields { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Work { observed, maximum })
        | Some(DecodeLimit::Nodes { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        Some(DecodeLimit::Text { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireOutputBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Allocation { requested }) => {
            BodyTableCellsError::Allocation { amount: requested }
        },
        None => BodyTableCellsError::InvalidSource,
    }
}

fn map_author_decode_error(
    error: litchi_iwa_protos::annotation_author_codec::DecodeError,
) -> BodyTableCellsError {
    use litchi_iwa_protos::annotation_author_codec::DecodeLimit;
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::OutputBytes { observed, maximum }) => {
            BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireOutputBytes,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(maximum),
            }
        },
        Some(DecodeLimit::Fields { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Work { observed, maximum })
        | Some(DecodeLimit::Allocations { observed, maximum })
        | Some(DecodeLimit::Scratch { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        Some(DecodeLimit::References { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::PayloadReferences,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Text { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::PayloadItems,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(DecodeLimit::Retained { observed, maximum }) => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::PayloadBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        Some(_) | None => BodyTableCellsError::InvalidSource,
    }
}

fn map_common_wire_error(error: litchi_iwa_common::Error) -> BodyTableCellsError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => BodyTableCellsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => BodyTableCellsLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::Fields => BodyTableCellsLimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => {
                    BodyTableCellsLimitKind::WireOutputBytes
                },
                litchi_iwa_common::LimitKind::Nesting => BodyTableCellsLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => BodyTableCellsLimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    BodyTableCellsLimitKind::PayloadItems
                },
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            BodyTableCellsError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => BodyTableCellsError::InvalidSource,
    }
}

fn map_common_table_error(error: litchi_iwa_common::table::model::Error) -> BodyTableCellsError {
    match error {
        litchi_iwa_common::table::model::Error::Allocation { amount, .. } => {
            BodyTableCellsError::Allocation { amount }
        },
        litchi_iwa_common::table::model::Error::BudgetExceeded { requested, maximum } => {
            BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::PayloadItems,
                observed: usize_as_u64(requested),
                maximum: usize_as_u64(maximum),
            }
        },
        litchi_iwa_common::table::model::Error::InvalidAddress { .. }
        | litchi_iwa_common::table::model::Error::OutOfBounds { .. }
        | litchi_iwa_common::table::model::Error::InvalidRange { .. }
        | litchi_iwa_common::table::model::Error::CoordinateOverflow { .. }
        | litchi_iwa_common::table::model::Error::DuplicatePosition { .. }
        | litchi_iwa_common::table::model::Error::DuplicateTableName { .. } => {
            BodyTableCellsError::InvalidSource
        },
    }
}

impl Package {
    /// Read one rooted Pages body table into the archive-free shared table
    /// model, resolving only the sidecars referenced by its selected cells.
    ///
    /// Table selection, model/DataStore admission, sidecar decoding, tile
    /// traversal, and semantic publication consume one aggregate budget.  The
    /// returned value contains no package objects, generated protobuf values,
    /// or native identifiers.
    pub fn body_table_cells<'table, S>(&self, selector: S) -> Result<TableRead, BodyTableCellsError>
    where
        S: Into<BodyTableSelector<'table>>,
    {
        let mut wire =
            table_lock::WireBudget::new(self.state.source.limits()).map_err(map_lock_error)?;
        let selected = selected_body_table_storage(self, selector.into(), &mut wire)?;
        let index = ObjectIndex::new(self, &mut wire)?;
        let mut budget = PagesCellReadBudget::new(wire);
        let sidecars = read_sidecars(&index, selected.data_store(), &mut budget)?;
        let formula_names = if sidecars
            .list(SidecarKind::Formulas)
            .is_some_and(|list| !list.is_empty())
        {
            Some(FormulaNameResolver::build(self, &index, &mut budget)?)
        } else {
            None
        };
        let dimensions = Dimensions::new(
            selected.model().number_of_rows(),
            selected.model().number_of_columns(),
        );
        let table_name = retain_table_name(selected.model().table_name(), &mut budget)?;
        let mut sink = PagesCellSink::new(
            table_name,
            dimensions,
            &sidecars,
            &index,
            formula_names.as_ref(),
        );
        selected.read_into(
            |object_id, _budget| index.resolve(object_id),
            &mut budget,
            &mut sink,
        )?;
        sink.finish()
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

/// Resolve one table-model object and decode its strict nested DataStore.
///
/// This is the Pages-owned admission point.  Selection and ownership are
/// revalidated before any retained tile route is allocated; model dimensions
/// and visible name must agree with the rooted target.  The nested DataStore
/// and tile-storage reports are charged into the same caller budget.
pub(crate) fn selected_body_table_storage<'source>(
    package: &'source Package,
    selector: BodyTableSelector<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<SelectedBodyTableStorage<'source>, BodyTableCellsError> {
    let target = package
        .resolve_body_table_with_budget(selector, budget)
        .map_err(map_lock_error)?;
    table_lock::validate_body_table_target(package, &target, budget).map_err(map_lock_error)?;
    let model_message = model_message(package, &target)?;
    let source = model_message.data.as_slice();
    let options = storage_options(budget, source)?;
    let mut routes = TileRoutes::default();
    let decoded =
        storage::decode_table_model_with_data_store_and_visitor(source, options, &mut routes)
            .map_err(map_storage_error)?;
    let (projection, report) = decoded;
    charge_storage_report(budget, report)?;
    let model = projection.model();
    if model.number_of_rows() != target.table_rows
        || model.number_of_columns() != target.table_columns
        || model.table_name() != target.table_name.as_ref()
    {
        return Err(BodyTableCellsError::InvalidSource);
    }

    // The model traversal validates and streams tile references, but its
    // visitor intentionally exposes no TileStorage scalar fields.  Decode the
    // same borrowed envelope once more to obtain tile_size for the shared
    // topology reader.  Its repeated object references were already charged
    // by the model traversal; fields/work remain charged as a second bounded
    // codec pass while references are retained by the first pass.
    let data_store = projection.data_store();
    let tile_options = storage_options(budget, data_store.tiles())?;
    let (tile_storage, tile_report) =
        storage::decode_tile_storage_with_report(data_store.tiles(), tile_options)
            .map_err(map_storage_error)?;
    charge_storage_report_without_references(budget, tile_report)?;
    let tile_size = tile_storage.tile_size().unwrap_or(256);
    if tile_size == 0 {
        return Err(BodyTableCellsError::InvalidSource);
    }

    let max_tiles = tile_count(model.number_of_rows(), tile_size)?;
    if routes.values.len() > max_tiles {
        return Err(BodyTableCellsError::InvalidSource);
    }
    Ok(SelectedBodyTableStorage {
        model,
        data_store,
        tile_size,
        tiles: routes.values,
    })
}

#[derive(Default)]
struct TileRoutes {
    values: Vec<wire_cells::TileReference>,
    tile_ids: std::collections::HashSet<u32>,
    object_ids: std::collections::HashSet<u64>,
}

impl storage::StorageVisitor for TileRoutes {
    fn visit_tile_reference(
        &mut self,
        record: storage::TileReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        let tile_id = record.tile_id();
        let object_id = record.reference().identifier();
        if self.tile_ids.contains(&tile_id) || self.object_ids.contains(&object_id) {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        self.values
            .try_reserve(1)
            .map_err(|_| storage::DecodeError::allocation(self.values.len().saturating_add(1)))?;
        self.tile_ids
            .try_reserve(1)
            .map_err(|_| storage::DecodeError::allocation(self.tile_ids.len().saturating_add(1)))?;
        self.object_ids.try_reserve(1).map_err(|_| {
            storage::DecodeError::allocation(self.object_ids.len().saturating_add(1))
        })?;
        self.tile_ids.insert(tile_id);
        self.object_ids.insert(object_id);
        self.values
            .push(wire_cells::TileReference::new(tile_id, object_id));
        Ok(())
    }
}

fn model_message<'source>(
    package: &'source Package,
    target: &table_lock::BodyTableTarget,
) -> Result<&'source RawMessage, BodyTableCellsError> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableCellsError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableCellsError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableCellsError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableCellsError::InvalidSource)?;
    object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableCellsError::InvalidSource)?;
    Ok(message)
}

fn storage_options(
    budget: &table_lock::WireBudget,
    source: &[u8],
) -> Result<storage::DecodeOptions, BodyTableCellsError> {
    let limits = budget.wire_limits();
    let input = source.len().max(1);
    if input > limits.max_input_bytes() {
        return Err(BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireBytes,
            observed: usize_as_u64(input),
            maximum: usize_as_u64(limits.max_input_bytes()),
        });
    }
    let fields = budget.remaining_wire_fields();
    if fields == 0 {
        return Err(BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireFields,
            observed: 1,
            maximum: 0,
        });
    }
    let work = budget.remaining_wire_work();
    if work == 0 {
        return Err(BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let references = budget.remaining_payload_references();
    let recursion = u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX);
    Ok(storage::DecodeOptions::new(
        input, fields, work, recursion, references, input,
    ))
}

fn charge_storage_report(
    budget: &mut table_lock::WireBudget,
    report: storage::DecodeReport,
) -> Result<(), BodyTableCellsError> {
    budget
        .charge_payload_work(report.source_bytes())
        .map_err(map_lock_error)?;
    budget
        .charge_codec_report(
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
        )
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(report.text_bytes())
        .map_err(map_lock_error)
}

fn charge_storage_probe_report(
    budget: &mut table_lock::WireBudget,
    report: storage::DecodeReport,
) -> Result<(), BodyTableCellsError> {
    budget
        .charge_payload_work(report.source_bytes())
        .map_err(map_lock_error)?;
    budget
        .charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)
        .map_err(map_lock_error)
}

fn charge_storage_report_without_references(
    budget: &mut table_lock::WireBudget,
    report: storage::DecodeReport,
) -> Result<(), BodyTableCellsError> {
    budget
        .charge_payload_work(report.source_bytes())
        .map_err(map_lock_error)?;
    budget
        .charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(report.text_bytes())
        .map_err(map_lock_error)
}

fn tile_count(rows: u32, tile_size: u32) -> Result<usize, BodyTableCellsError> {
    let rows = usize::try_from(rows).map_err(|_| BodyTableCellsError::InvalidSource)?;
    let tile_size = usize::try_from(tile_size).map_err(|_| BodyTableCellsError::InvalidSource)?;
    rows.checked_add(tile_size.saturating_sub(1))
        .and_then(|value| value.checked_div(tile_size))
        .ok_or(BodyTableCellsError::InvalidSource)
}

fn map_storage_error(error: storage::DecodeError) -> BodyTableCellsError {
    let Some(limit) = error.resource_limit() else {
        return BodyTableCellsError::InvalidSource;
    };
    match limit {
        storage::DecodeLimit::Bytes { observed, maximum } => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        storage::DecodeLimit::References { observed, maximum } => {
            BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::PayloadReferences,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(maximum),
            }
        },
        storage::DecodeLimit::Text { observed, maximum } => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::PayloadItems,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        storage::DecodeLimit::Fields { observed, maximum } => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        storage::DecodeLimit::Work { observed, maximum } => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        storage::DecodeLimit::Nesting { observed, maximum } => BodyTableCellsError::LimitExceeded {
            kind: BodyTableCellsLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        storage::DecodeLimit::Allocation { requested }
        | storage::DecodeLimit::Retained {
            observed: requested,
            maximum: _,
        } => BodyTableCellsError::Allocation { amount: requested },
        _ => BodyTableCellsError::InvalidSource,
    }
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableCellsError {
    match error {
        table_lock::BodyTableLockError::TableNotFound => BodyTableCellsError::TableNotFound,
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableCellsError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::AmbiguousSelector => BodyTableCellsError::AmbiguousSelector,
        table_lock::BodyTableLockError::UnsupportedSource => BodyTableCellsError::UnsupportedSource,
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableCellsError::LimitExceeded {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableCellsError::Allocation { amount }
        },
        table_lock::BodyTableLockError::InvalidSource
        | table_lock::BodyTableLockError::Verification
        | table_lock::BodyTableLockError::PatchConflict => BodyTableCellsError::InvalidSource,
    }
}

const fn map_lock_limit(kind: table_lock::BodyTableLockLimitKind) -> BodyTableCellsLimitKind {
    use table_lock::BodyTableLockLimitKind as Lock;
    match kind {
        Lock::InputBytes => BodyTableCellsLimitKind::InputBytes,
        Lock::OutputBytes => BodyTableCellsLimitKind::OutputBytes,
        Lock::Entries => BodyTableCellsLimitKind::Entries,
        Lock::EntryBytes => BodyTableCellsLimitKind::EntryBytes,
        Lock::TotalEntryBytes => BodyTableCellsLimitKind::TotalEntryBytes,
        Lock::PackageBytes => BodyTableCellsLimitKind::PackageBytes,
        Lock::PayloadBytes => BodyTableCellsLimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => BodyTableCellsLimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => BodyTableCellsLimitKind::PayloadObjects,
        Lock::PayloadMessages => BodyTableCellsLimitKind::PayloadMessages,
        Lock::PayloadItems => BodyTableCellsLimitKind::PayloadItems,
        Lock::PayloadReferences => BodyTableCellsLimitKind::PayloadReferences,
        Lock::WireBytes => BodyTableCellsLimitKind::WireBytes,
        Lock::WireFields => BodyTableCellsLimitKind::WireFields,
        Lock::WireNesting => BodyTableCellsLimitKind::WireNesting,
        Lock::WireWork => BodyTableCellsLimitKind::WireWork,
    }
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn formula_names_resolve_complete_owner_and_fail_closed_for_unknown_ids() {
        let mut owners = HashMap::new();
        owners.insert([1, 2, 3, 4], 0);
        let resolver = FormulaNameResolver {
            table_names: vec!["Table 2".into()],
            owners,
            categories: HashMap::new(),
            unresolved_table: StateCell::new(false),
            unresolved_category: StateCell::new(false),
        };
        let known = litchi_iwa_protos::numbers_formula_codec::FormulaRenderCfuuid {
            has_uuid_bytes: true,
            word0: Some(1),
            word1: Some(2),
            word2: Some(3),
            word3: Some(4),
        };
        assert_eq!(resolver.table_only_name(&known), Some("Table 2"));
        assert_eq!(
            resolver.category_name(FormulaCategoryId { lower: 1, upper: 0 }),
            Some("Grand Total")
        );

        let unknown = litchi_iwa_protos::numbers_formula_codec::FormulaRenderCfuuid {
            has_uuid_bytes: true,
            word0: Some(4),
            word1: Some(3),
            word2: Some(2),
            word3: Some(1),
        };
        assert_eq!(resolver.table_only_name(&unknown), None);
        assert_eq!(
            resolver.category_name(FormulaCategoryId { lower: 9, upper: 9 }),
            None
        );
        assert!(resolver.has_unresolved_reference());
    }

    #[test]
    fn body_table_cells_rejects_missing_name_and_position_selectors() {
        let package = Package::from_bytes(include_bytes!(
            "../../../../test-data/iwork/pages/body-table-read-native.pages"
        ))
        .expect("native Pages table fixture");
        assert_eq!(
            package.body_table_cells("missing table"),
            Err(BodyTableCellsError::TableNotFound)
        );
        assert_eq!(
            package.body_table_cells(usize::MAX),
            Err(BodyTableCellsError::TableNotFound)
        );
    }

    #[test]
    fn body_table_cells_honors_aggregate_wire_field_ceiling() {
        let source =
            include_bytes!("../../../../test-data/iwork/pages/body-table-read-native.pages");
        let archive_limits = litchi_iwa_core::Limits::default()
            .with_header_fields(256)
            .expect("valid archive field limit");
        let limits = litchi_iwa_archive::Limits::default()
            .with_archive_limits(archive_limits)
            .expect("valid package limits");
        let package = Package::from_bytes_with_limits(source, limits)
            .expect("native fixture must open before the focused read budget is spent");
        let error = package
            .body_table_cells(0usize)
            .expect_err("selection and cell traversal must exceed the shared field ceiling");
        assert!(matches!(
            error,
            BodyTableCellsError::LimitExceeded {
                kind: BodyTableCellsLimitKind::WireFields,
                observed,
                maximum,
            } if observed > maximum && maximum < 256
        ));
    }

    #[test]
    fn body_table_cells_reads_direct_comment_replies() {
        let package = Package::from_bytes(include_bytes!(
            "../../../../test-data/iwork/pages/body-table-comment-hidden-native-saved.pages"
        ))
        .expect("native Pages comment fixture");
        let read = package
            .body_table_cells("Table 1")
            .expect("native table cells and comment should decode");
        let comment = read
            .get_comment_a1("B2")
            .expect("valid cell address")
            .expect("native B2 comment");
        let replies = comment
            .replies()
            .expect("the focused reader should materialize direct replies");
        assert_eq!(replies.len(), 1);
        assert_eq!(replies[0].text(), "Native hidden-axis control reply");
        assert!(replies[0].timestamp().is_some());
        assert!(replies[0].author().is_some());
    }
}
