//! Exact-source transactions for persisted Keynote slide-table sort order.
//!
//! This owner deliberately edits only field 44 of one canonical
//! `TST.TableModelArchive`.  Physical row reordering, selected-row execution,
//! legacy type-6000 models, and native object identity remain compatibility
//! host concerns.  The public surface contains only semantic values and
//! checked slide/table selectors.  Field 45 (the native tracker) and every
//! other model field are intentionally opaque and are never rewritten.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "the package boundary redacts lower-layer failure details"
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceScope, ArchiveReferenceVisitor, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{table_info_codec, table_sort_order_codec as codec};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::TableSelector;
use crate::slide::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const TABLE_MODEL_COLUMNS_FIELD: u32 = 7;

/// Finite resource categories enforced by a slide-table sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableSortLimitKind {
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

impl fmt::Display for SlideTableSortLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
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

/// Content-free semantic path for a slide-table sort operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableSortPath {
    Package,
    Table { slide: Position, table: Position },
}

impl fmt::Display for SlideTableSortPath {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => f.write_str("package"),
            Self::Table { slide, table } => {
                write!(f, "slide {} table {}", slide.get(), table.get())
            },
        }
    }
}

/// Failure from a Keynote slide-table sort read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTableSortError {
    #[error("this Keynote source does not support persisted slide-table sort edits")]
    UnsupportedSource,
    #[error("the requested Keynote slide-table sort graph is outside the supported scope")]
    UnsupportedDependency,
    #[error("the requested Keynote slide-table sort topology is unsupported")]
    UnsupportedTopology,
    #[error("the Keynote slide-table sort selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    #[error("the selected Keynote slide-table is locked")]
    Locked,
    #[error("the selected Keynote slide-table sort source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table sort {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTableSortLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote slide-table sort transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote slide-table sort failed semantic verification")]
    Verification,
    #[error("the Keynote slide-table sort patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy)]
struct SortBudget {
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
}

impl SortBudget {
    fn new(package: &Package) -> Result<Self, SlideTableSortError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let source = usize::try_from(package.state.options.archive().max_input_bytes())
            .map_err(|_| SlideTableSortError::InvalidSource)?;
        let aggregate = source
            .checked_mul(4)
            .ok_or(SlideTableSortError::InvalidSource)?;
        Ok(Self {
            max_input: aggregate,
            max_output: aggregate,
            max_fields: wire.max_fields(),
            max_work: wire.max_rewrite_work(),
            max_nesting: wire.max_nesting(),
            max_references: package.semantic_limits().max_references(),
            max_allocations: aggregate,
            max_retained: aggregate,
            max_scratch: aggregate,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
        })
    }

    fn add(
        cur: &mut usize,
        amount: usize,
        max: usize,
        kind: SlideTableSortLimitKind,
    ) -> Result<(), SlideTableSortError> {
        let observed = cur
            .checked_add(amount)
            .ok_or(SlideTableSortError::InvalidSource)?;
        if observed > max {
            return Err(SlideTableSortError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: max as u64,
            });
        }
        *cur = observed;
        Ok(())
    }

    fn input(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        Self::add(
            &mut self.input,
            n,
            self.max_input,
            SlideTableSortLimitKind::InputBytes,
        )
    }
    fn output(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        Self::add(
            &mut self.output,
            n,
            self.max_output,
            SlideTableSortLimitKind::OutputBytes,
        )
    }
    fn fields(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        Self::add(
            &mut self.fields,
            n,
            self.max_fields,
            SlideTableSortLimitKind::WireFields,
        )
    }
    fn work(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        Self::add(
            &mut self.work,
            n,
            self.max_work,
            SlideTableSortLimitKind::WireWork,
        )
    }
    fn references(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        Self::add(
            &mut self.references,
            n,
            self.max_references,
            SlideTableSortLimitKind::References,
        )
    }
    fn allocations(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        Self::add(
            &mut self.allocations,
            n,
            self.max_allocations,
            SlideTableSortLimitKind::Allocations,
        )
    }
    fn retained(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        Self::add(
            &mut self.retained,
            n,
            self.max_retained,
            SlideTableSortLimitKind::Retained,
        )
    }
    fn scratch(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        Self::add(
            &mut self.scratch,
            n,
            self.max_scratch,
            SlideTableSortLimitKind::Scratch,
        )
    }

    fn physical(&mut self, n: usize) -> Result<(), SlideTableSortError> {
        self.input(n)?;
        self.work(n)
    }

    fn codec(&mut self, report: codec::DecodeReport) -> Result<(), SlideTableSortError> {
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.references(report.rules())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableSortError::LimitExceeded {
                kind: SlideTableSortLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())
    }

    fn residual(&self, package: &Package) -> Result<WireLimits, SlideTableSortError> {
        let base = package.wire_limits().map_err(map_wire_error)?;
        base.with_input_bytes(
            base.max_input_bytes()
                .min(self.max_input.saturating_sub(self.input).max(1)),
        )
        .and_then(|v| {
            v.with_fields(
                base.max_fields()
                    .min(self.max_fields.saturating_sub(self.fields).max(1)),
            )
        })
        .and_then(|v| {
            v.with_rewrite_work(
                base.max_rewrite_work()
                    .min(self.max_work.saturating_sub(self.work).max(1)),
            )
        })
        .and_then(|v| v.with_nesting(base.max_nesting().min(self.max_nesting)))
        .map_err(map_wire_error)
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideTableSortError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.work(requirements.output_bytes())
    }
}

/// One mutable sort value staged against an immutable package snapshot.
pub struct SlideTableSortEdit<'a> {
    source: &'a Package,
    selection: SortSelection,
    after: Option<Order>,
}

impl fmt::Debug for SlideTableSortEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideTableSortEdit")
            .field("path", &self.selection.path())
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableSortEdit<'_> {
    #[must_use]
    pub fn path(&self) -> SlideTableSortPath {
        self.selection.path()
    }
    #[must_use]
    pub fn order(&self) -> Option<&Order> {
        self.after.as_ref()
    }
    #[must_use]
    pub fn before(&self) -> Option<&Order> {
        self.selection.before.as_ref()
    }
    #[must_use]
    pub fn set(mut self, order: Order) -> Self {
        self.after = Some(order);
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
    pub fn commit(self) -> Result<SlideTableSortCommit, SlideTableSortError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible slide-table sort patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableSortPatch {
    artifacts: ExactArtifacts,
    selection: SortSelection,
    before: Option<Order>,
    after: Option<Order>,
}

impl fmt::Debug for SlideTableSortPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideTableSortPatch")
            .field("path", &self.selection.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableSortPatch {
    #[must_use]
    pub fn before(&self) -> Option<&Order> {
        self.before.as_ref()
    }
    #[must_use]
    pub fn after(&self) -> Option<&Order> {
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
            selection: self.selection.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Diagnostics for one persisted slide-table sort publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableSortDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideTableSortDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }
    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews: 0,
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
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one slide-table sort transaction.
#[must_use = "a Keynote slide-table sort commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableSortCommit {
    package: Package,
    patch: SlideTableSortPatch,
    diagnostics: SlideTableSortDiagnostics,
}

impl SlideTableSortCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    pub const fn patch(&self) -> &SlideTableSortPatch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableSortDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
struct SortSelection {
    slide_position: Position,
    table_position: Position,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    table_info_message_index: usize,
    model_message_index: usize,
    component_name: Arc<str>,
    before: Option<Order>,
    locked: bool,
    column_count: u32,
}

impl fmt::Debug for SortSelection {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SortSelection")
            .field("slide_position", &self.slide_position)
            .field("table_position", &self.table_position)
            .field("before", &self.before)
            .field("locked", &self.locked)
            .finish_non_exhaustive()
    }
}

impl SortSelection {
    const fn path(&self) -> SlideTableSortPath {
        SlideTableSortPath::Table {
            slide: self.slide_position,
            table: self.table_position,
        }
    }
}

impl Package {
    /// Read the persisted sort configuration of one slide table.
    pub fn slide_table_sort_order<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Option<Order>, SlideTableSortError> {
        Ok(select_table(self, slide.into(), table.into())?.before)
    }

    /// Begin an immutable exact edit of one persisted slide-table sort.
    pub fn edit_slide_table_sort_order<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTableSortEdit<'_>, SlideTableSortError> {
        let selection = select_table(self, slide.into(), table.into())?;
        Ok(SlideTableSortEdit {
            source: self,
            after: selection.before.clone(),
            selection,
        })
    }

    /// Apply an exact-source checked reversible slide-table sort patch.
    pub fn apply_slide_table_sort_order(
        &self,
        patch: &SlideTableSortPatch,
    ) -> Result<SlideTableSortCommit, SlideTableSortError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideTableSortError::PatchConflict);
        }
        let current = select_table(
            self,
            SlideSelector::position(patch.selection.slide_position),
            TableSelector::position(patch.selection.table_position),
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SlideTableSortError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideTableSortCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableSortDiagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch)
    }
}

fn commit_edit(
    source: &Package,
    selection: &SortSelection,
    after: Option<Order>,
) -> Result<SlideTableSortCommit, SlideTableSortError> {
    if selection.before == after {
        let bytes: Arc<[u8]> = Arc::from(source.source_bytes());
        return Ok(SlideTableSortCommit {
            package: source.snapshot(),
            patch: SlideTableSortPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before.clone(),
                after,
            },
            diagnostics: SlideTableSortDiagnostics::unchanged(),
        });
    }
    if selection.locked {
        return Err(SlideTableSortError::Locked);
    }
    let mut budget = SortBudget::new(source)?;
    let candidate = rewrite_sort(source, selection, after.clone(), &mut budget)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(selection.slide_position),
        TableSelector::position(selection.table_position),
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(SlideTableSortError::Verification);
    }
    verify_locality(source, &candidate, selection, &mut budget)?;
    let target = physical_catalog(&candidate)?.shared_source();
    Ok(SlideTableSortCommit {
        package: candidate,
        patch: SlideTableSortPatch {
            artifacts: ExactArtifacts::new(Arc::from(source.source_bytes()), target),
            selection: selection.clone(),
            before: selection.before.clone(),
            after,
        },
        diagnostics: SlideTableSortDiagnostics::published(),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &SlideTableSortPatch,
) -> Result<SlideTableSortCommit, SlideTableSortError> {
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        TableSelector::position(patch.selection.table_position),
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SlideTableSortError::Verification);
    }
    let mut budget = SortBudget::new(source)?;
    budget.input(patch.artifacts.target().len())?;
    verify_locality(source, &candidate, &patch.selection, &mut budget)?;
    Ok(SlideTableSortCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableSortDiagnostics::published(),
    })
}

fn select_table(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
) -> Result<SortSelection, SlideTableSortError> {
    let mut budget = SortBudget::new(package)?;
    let catalog = physical_catalog(package)?;
    budget.input(package.source_bytes().len())?;
    budget.allocations(catalog.package().len())?;
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTableSortError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideTableSortError::InvalidSource)?;
    let (slide_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
    let limits = budget.residual(package)?;
    let owned = repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits)?;
    let z_order = repeated_references(slide_payload, SLIDE_Z_ORDER_FIELD, limits)?;
    budget.references(
        owned
            .len()
            .checked_add(z_order.len())
            .ok_or(SlideTableSortError::InvalidSource)?,
    )?;
    reject_duplicates(&owned)?;
    reject_duplicates(&z_order)?;
    validate_slide_metadata(slide, slide_index, &owned, &z_order)?;
    let mut tables: Vec<(u64, u64, Arc<str>, usize, usize, Option<Order>, bool, u32)> = Vec::new();
    tables
        .try_reserve_exact(z_order.len())
        .map_err(|_| SlideTableSortError::Allocation {
            amount: z_order.len(),
        })?;
    for identifier in z_order {
        let Some((_owner, object)) = package.object_with_component(identifier) else {
            return Err(SlideTableSortError::InvalidSource);
        };
        if !object
            .messages
            .iter()
            .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        {
            continue;
        }
        if owned.iter().filter(|value| **value == identifier).count() != 1 {
            return Err(SlideTableSortError::InvalidSource);
        }
        let (info_index, info_payload) = unique_message(object, TABLE_INFO_MESSAGE_TYPE)?;
        let info = table_info_codec::decode_table_info(
            info_payload,
            table_info_codec::DecodeOptions::new(
                info_payload.len().max(1),
                limits.max_fields(),
                limits.max_rewrite_work(),
                u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            ),
        )
        .map_err(|_| SlideTableSortError::InvalidSource)?;
        let parent = table_parent(info_payload, limits)?;
        if parent != record.slide_identifier {
            return Err(SlideTableSortError::InvalidSource);
        }
        let model_identifier = info.table_model().identifier().get();
        validate_table_info_metadata(object, info_index, parent, model_identifier)?;
        let (model_component, model) = package
            .object_with_component(model_identifier)
            .ok_or(SlideTableSortError::InvalidSource)?;
        if model
            .messages
            .iter()
            .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        {
            return Err(SlideTableSortError::UnsupportedDependency);
        }
        let (model_index, model_payload) = unique_message(model, TABLE_MODEL_MESSAGE_TYPE)?;
        let before = decode_sort(model_payload, package, &mut budget)?;
        let columns = model_column_count(model_payload, limits)?;
        if let Some(order) = &before {
            validate_columns(order, columns)?;
        }
        tables.push((
            identifier,
            model_identifier,
            Arc::from(model_component),
            info_index,
            model_index,
            before,
            info.locked().unwrap_or(false),
            columns,
        ));
    }
    let table_position = table_selector.as_position();
    let (
        table_info_identifier,
        model_identifier,
        component_name,
        table_info_message_index,
        model_message_index,
        before,
        locked,
        column_count,
    ) = tables.get(table_position.get()).cloned().ok_or(
        SlideTableSortError::TablePositionNotFound {
            position: table_position,
        },
    )?;
    ensure_unique_identity(package, table_info_identifier)?;
    ensure_unique_identity(package, model_identifier)?;
    ensure_unique_table_owner(
        package,
        record.slide_identifier,
        table_info_identifier,
        model_identifier,
        limits,
    )?;
    validate_global_inbound_references(
        package,
        record.slide_identifier,
        slide_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
    )?;
    Ok(SortSelection {
        slide_position,
        table_position,
        slide_identifier: record.slide_identifier,
        table_info_identifier,
        model_identifier,
        table_info_message_index,
        model_message_index,
        component_name,
        before,
        locked,
        column_count,
    })
}

fn rewrite_sort(
    source: &Package,
    selection: &SortSelection,
    after: Option<Order>,
    budget: &mut SortBudget,
) -> Result<Package, SlideTableSortError> {
    if let Some(order) = &after {
        validate_columns(order, selection.column_count)?;
    }
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.component_name.as_ref())
        .ok_or(SlideTableSortError::InvalidSource)?;
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.physical(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    let object = archive
        .object(selection.model_identifier)
        .ok_or(SlideTableSortError::InvalidSource)?;
    let original = object
        .messages
        .get(selection.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(SlideTableSortError::InvalidSource)?
        .data
        .as_slice();
    let current = decode_sort(original, source, budget)?;
    if current != selection.before {
        return Err(SlideTableSortError::InvalidSource);
    }
    let residual = budget.residual(source)?;
    let options = codec_options(original, residual, budget.remaining_allocations());
    // Preparation has conservative source/candidate accounting; charge a
    // bounded envelope before allowing the codec to allocate its plan.
    let envelope = original
        .len()
        .checked_mul(16)
        .ok_or(SlideTableSortError::InvalidSource)?;
    budget.work(envelope)?;
    budget.allocations(16)?;
    budget.scratch(envelope)?;
    let prepared = codec::prepare_table_model_sort_order_rewrite(
        original,
        order_to_snapshot(after.as_ref())?,
        options,
    )
    .map_err(map_codec_error)?;
    let requirements = prepared.execution_requirements();
    budget.codec_requirements(requirements)?;
    let execution_limits = requirements.exact_limits();
    let output = prepared
        .execute(execution_limits)
        .map_err(map_codec_error)?;
    let report = output.report();
    if report.output_bytes() != requirements.output_bytes
        || report.fields() != requirements.fields
        || report.work_bytes() != requirements.work_bytes
        || report.allocations() != requirements.allocations
        || report.retained_bytes() != requirements.retained_bytes
        || report.scratch_bytes() != requirements.scratch_bytes
    {
        return Err(SlideTableSortError::Verification);
    }
    let rewritten = output.into_bytes();
    let verified = decode_sort(&rewritten, source, budget)?;
    if verified != after {
        return Err(SlideTableSortError::Verification);
    }
    archive
        .object_mut(selection.model_identifier)
        .ok_or(SlideTableSortError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_len = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_len).map_err(map_core_error)?;
    budget.output(
        encoded_len
            .checked_add(compressed_bound)
            .ok_or(SlideTableSortError::InvalidSource)?,
    )?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if bytes.len() != encoded_len {
        return Err(SlideTableSortError::Verification);
    }
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(SlideTableSortError::Verification);
    }
    let edits = [EntryEdit::new(
        selection.component_name.as_ref(),
        &compressed,
    )];
    let prepared_reassembly = catalog
        .prepare_reassembly(&edits, physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared_reassembly.execution_requirements();
    budget.reassembly(requirements)?;
    let output = prepared_reassembly
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    Package::from_source_with_options(output.into(), source.state.options).map_err(map_read_error)
}

fn order_to_snapshot(
    order: Option<&Order>,
) -> Result<Option<codec::SortOrderSnapshot>, SlideTableSortError> {
    order
        .map(|value| {
            let scope = match value.scope() {
                Scope::EntireTable => codec::SortScope::EntireTable,
                Scope::SelectedRows => codec::SortScope::SelectedRows,
            };
            let rules = value.rules().iter().map(|rule| {
                codec::SortRule::new(
                    rule.column().native_value(),
                    match rule.direction() {
                        Direction::Ascending => codec::SortDirection::Ascending,
                        Direction::Descending => codec::SortDirection::Descending,
                    },
                )
            });
            codec::SortOrderSnapshot::new(scope, rules)
                .map_err(|_| SlideTableSortError::InvalidSource)
        })
        .transpose()
}

fn snapshot_to_order(
    snapshot: Option<codec::SortOrderSnapshot>,
) -> Result<Option<Order>, SlideTableSortError> {
    snapshot
        .map(|value| {
            let scope = match value.scope() {
                codec::SortScope::EntireTable => Scope::EntireTable,
                codec::SortScope::SelectedRows => Scope::SelectedRows,
            };
            let mut rules = Vec::new();
            rules.try_reserve_exact(value.rules().len()).map_err(|_| {
                SlideTableSortError::Allocation {
                    amount: value.rules().len(),
                }
            })?;
            for rule in value.rules() {
                let column = ColumnIndex::from_native(rule.column())
                    .map_err(|_| SlideTableSortError::InvalidSource)?;
                let direction = match rule.direction() {
                    codec::SortDirection::Ascending => Direction::Ascending,
                    codec::SortDirection::Descending => Direction::Descending,
                };
                rules.push(Rule::new(column, direction));
            }
            Order::with_scope(scope, rules).map_err(|_| SlideTableSortError::InvalidSource)
        })
        .transpose()
}

fn decode_sort(
    payload: &[u8],
    package: &Package,
    budget: &mut SortBudget,
) -> Result<Option<Order>, SlideTableSortError> {
    let limits = budget.residual(package)?;
    let options = codec_options(payload, limits, budget.remaining_allocations());
    let (snapshot, report) = codec::decode_table_model_sort_order_with_report(payload, options)
        .map_err(map_codec_error)?;
    budget.codec(report)?;
    snapshot_to_order(snapshot)
}

fn codec_options(
    payload: &[u8],
    limits: WireLimits,
    max_allocations: usize,
) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len().max(1)),
        limits
            .max_output_bytes()
            .min(payload.len().saturating_add(256).max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        payload.len().max(1),
        usize::try_from(u32::MAX).unwrap_or(usize::MAX),
    )
    .with_max_allocations(max_allocations)
}

fn validate_columns(order: &Order, count: u32) -> Result<(), SlideTableSortError> {
    order.rules().iter().try_for_each(|rule| {
        if rule.column().native_value() >= count {
            Err(SlideTableSortError::InvalidSource)
        } else {
            Ok(())
        }
    })
}

impl SortBudget {
    fn codec_requirements(
        &mut self,
        req: codec::RewriteExecutionRequirements,
    ) -> Result<(), SlideTableSortError> {
        self.output(req.output_bytes)?;
        self.fields(req.fields)?;
        self.work(req.work_bytes)?;
        self.allocations(req.allocations)?;
        self.retained(req.retained_bytes)?;
        self.scratch(req.scratch_bytes)?;
        self.nesting = self.nesting.max(req.max_depth as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableSortError::LimitExceeded {
                kind: SlideTableSortLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn remaining_allocations(&self) -> usize {
        self.max_allocations.saturating_sub(self.allocations).max(1)
    }
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &SortSelection,
    budget: &mut SortBudget,
) -> Result<(), SlideTableSortError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    for entry in source_catalog.package().iter() {
        let other = candidate_catalog
            .package()
            .iter()
            .find(|value| value.name() == entry.name())
            .ok_or(SlideTableSortError::Verification)?;
        if entry.name() != selection.component_name.as_ref()
            && (entry.data() != other.data() || entry.metadata() != other.metadata())
        {
            return Err(SlideTableSortError::Verification);
        }
        budget.work(entry.data().len())?;
    }
    let source_archive = component_archive(source, selection.component_name.as_ref())?;
    let candidate_archive = component_archive(candidate, selection.component_name.as_ref())?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SlideTableSortError::Verification);
    }
    let limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    for source_object in &source_archive.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(SlideTableSortError::Verification)?;
        let other = candidate_archive
            .object(identifier)
            .ok_or(SlideTableSortError::Verification)?;
        if identifier != selection.model_identifier
            && !source_object.same_content_ignoring_offsets(other)
        {
            return Err(SlideTableSortError::Verification);
        }
        if identifier == selection.model_identifier {
            let message = other
                .messages
                .get(selection.model_message_index)
                .ok_or(SlideTableSortError::Verification)?
                .clone();
            let mut expected = source_object.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    selection.model_message_index,
                    message,
                    limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = other.header_length;
            expected.data_length = other.data_length;
            if !expected.same_content_ignoring_offsets(other) {
                return Err(SlideTableSortError::Verification);
            }
        }
    }
    Ok(())
}

fn component_archive(package: &Package, name: &str) -> Result<Archive, SlideTableSortError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(SlideTableSortError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableSortError::InvalidSource);
    }
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        package
            .state
            .options
            .archive()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    Archive::parse_with_limits(
        stream.as_bytes(),
        package
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideTableSortError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideTableSortError::EmptySlideName);
            }
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(|_| SlideTableSortError::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideTableSortError::SlideNameNotFound)
        },
    }
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), SlideTableSortError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableSortError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        validate_message_header(object, index)?;
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideTableSortError::InvalidSource);
        }
    }
    selected.ok_or(SlideTableSortError::InvalidSource)
}

fn validate_message_header(
    object: &ArchiveObject,
    index: usize,
) -> Result<(), SlideTableSortError> {
    let message = object
        .messages
        .get(index)
        .ok_or(SlideTableSortError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(SlideTableSortError::InvalidSource)?;
    if message.type_ != info.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SlideTableSortError::InvalidSource);
    }
    Ok(())
}

fn repeated_references(
    payload: &[u8],
    number: u32,
    limits: WireLimits,
) -> Result<Vec<u64>, SlideTableSortError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let count = fields
        .fields()
        .filter(|field| field.number() == number)
        .count();
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|_| SlideTableSortError::Allocation { amount: count })?;
    for field in fields.fields().filter(|field| field.number() == number) {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideTableSortError::InvalidSource);
        }
        values.push(strict_reference(field.payload(), limits)?);
    }
    Ok(values)
}

fn table_parent(payload: &[u8], limits: WireLimits) -> Result<u64, SlideTableSortError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let super_fields: Vec<_> = fields
        .fields()
        .filter(|field| field.number() == TABLE_SUPER_FIELD)
        .collect();
    if super_fields.len() != 1 || super_fields[0].wire_type() != 2 {
        return Err(SlideTableSortError::InvalidSource);
    }
    let drawable =
        WireView::parse_with_limits(super_fields[0].payload(), limits).map_err(map_wire_error)?;
    let parents: Vec<_> = drawable
        .fields()
        .filter(|field| field.number() == DRAWABLE_PARENT_FIELD)
        .collect();
    if parents.len() != 1 || parents[0].wire_type() != 2 {
        return Err(SlideTableSortError::InvalidSource);
    }
    strict_reference(parents[0].payload(), limits)
}

fn strict_reference(payload: &[u8], limits: WireLimits) -> Result<u64, SlideTableSortError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideTableSortError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideTableSortError::InvalidSource)?;
                if value == 0 || width != encoded_len(value) {
                    return Err(SlideTableSortError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideTableSortError::UnsupportedDependency),
            _ => {},
        }
    }
    identifier.ok_or(SlideTableSortError::InvalidSource)
}

fn model_column_count(payload: &[u8], limits: WireLimits) -> Result<u32, SlideTableSortError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut value = None;
    for field in fields
        .fields()
        .filter(|field| field.number() == TABLE_MODEL_COLUMNS_FIELD)
    {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 0 || value.is_some() {
            return Err(SlideTableSortError::InvalidSource);
        }
        let (decoded, width) = decode_varint_from_bytes(field.payload())
            .map_err(|_| SlideTableSortError::InvalidSource)?;
        if decoded > u64::from(u32::MAX) || width != encoded_len(decoded) {
            return Err(SlideTableSortError::InvalidSource);
        }
        value = Some(decoded as u32);
    }
    match value {
        Some(value) if value != 0 => Ok(value),
        _ => Err(SlideTableSortError::InvalidSource),
    }
}

fn reject_duplicates(values: &[u64]) -> Result<(), SlideTableSortError> {
    if values
        .iter()
        .enumerate()
        .any(|(i, value)| values[..i].contains(value))
    {
        Err(SlideTableSortError::InvalidSource)
    } else {
        Ok(())
    }
}

fn validate_slide_metadata(
    object: &ArchiveObject,
    index: usize,
    owned: &[u64],
    z_order: &[u64],
) -> Result<(), SlideTableSortError> {
    validate_message_header(object, index)?;
    let info = &object.archive_info.message_infos[index];
    if info
        .object_references
        .iter()
        .enumerate()
        .any(|(i, id)| info.object_references[..i].contains(id))
    {
        return Err(SlideTableSortError::InvalidSource);
    }
    for identifier in owned.iter().chain(z_order) {
        if !info.object_references.is_empty()
            && info
                .object_references
                .iter()
                .filter(|id| *id == identifier)
                .count()
                != 1
        {
            return Err(SlideTableSortError::InvalidSource);
        }
    }
    for field in &info.field_infos {
        if field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD]
            && field.object_references.as_slice() != owned
        {
            return Err(SlideTableSortError::InvalidSource);
        }
        if field.path.as_slice() == [SLIDE_Z_ORDER_FIELD]
            && field.object_references.as_slice() != z_order
        {
            return Err(SlideTableSortError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_table_info_metadata(
    object: &ArchiveObject,
    index: usize,
    _parent: u64,
    model: u64,
) -> Result<(), SlideTableSortError> {
    validate_message_header(object, index)?;
    let info = &object.archive_info.message_infos[index];
    // Native Keynote treats the drawable parent as a weak/rooting route: it
    // is authoritative in the TableInfo payload but is not repeated in the
    // MessageInfo aggregate.  The model is the strong object edge and must
    // occur exactly once whenever the producer supplied an aggregate.
    if !info.object_references.is_empty()
        && info
            .object_references
            .iter()
            .filter(|id| **id == model)
            .count()
            != 1
    {
        return Err(SlideTableSortError::InvalidSource);
    }
    for field in &info.field_infos {
        if field.path.as_slice() == [TABLE_MODEL_FIELD]
            && field.object_references.as_slice() != [model]
        {
            return Err(SlideTableSortError::InvalidSource);
        }
    }
    Ok(())
}

fn ensure_unique_identity(package: &Package, identifier: u64) -> Result<(), SlideTableSortError> {
    let mut count = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.archive_info.identifier == Some(identifier) {
                count += 1;
            }
        }
    }
    if count == 1 {
        Ok(())
    } else {
        Err(SlideTableSortError::UnsupportedDependency)
    }
}

fn ensure_unique_table_owner(
    package: &Package,
    slide: u64,
    table_info: u64,
    model: u64,
    limits: WireLimits,
) -> Result<(), SlideTableSortError> {
    let mut owned_count = 0usize;
    let mut z_count = 0usize;
    let mut selected = false;
    let mut model_owners = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ == SLIDE_MESSAGE_TYPE {
                    let owned =
                        repeated_references(&message.data, SLIDE_OWNED_DRAWABLES_FIELD, limits)?;
                    let z = repeated_references(&message.data, SLIDE_Z_ORDER_FIELD, limits)?;
                    let oh = owned.iter().filter(|id| **id == table_info).count();
                    let zh = z.iter().filter(|id| **id == table_info).count();
                    owned_count += oh;
                    z_count += zh;
                    if object.archive_info.identifier == Some(slide) && oh == 1 && zh == 1 {
                        selected = true;
                    }
                }
                if message.type_ == TABLE_INFO_MESSAGE_TYPE {
                    let info = table_info_codec::decode_table_info(
                        &message.data,
                        table_info_codec::DecodeOptions::new(
                            message.data.len().max(1),
                            limits.max_fields(),
                            limits.max_rewrite_work(),
                            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
                        ),
                    )
                    .map_err(|_| SlideTableSortError::InvalidSource)?;
                    if info.table_model().identifier().get() == model {
                        model_owners += 1;
                    }
                }
            }
        }
    }
    if owned_count == 1 && z_count == 1 && selected && model_owners == 1 {
        Ok(())
    } else {
        Err(SlideTableSortError::UnsupportedDependency)
    }
}

/// Prove that the selected table graph has no opaque or foreign inbound
/// owners.  A package may place the model in a different physical member, but
/// only the rooted slide and selected TableInfo are allowed to point at the
/// selected table/model pair.  This is intentionally a read/edit admission
/// check: a future ArchiveInfo field must not silently become a second owner
/// of a model that this narrow field-44 transaction would rewrite.
fn validate_global_inbound_references(
    package: &Package,
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
) -> Result<(), SlideTableSortError> {
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut census = InboundReferenceCensus {
        slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
        model_edges: 0,
        invalid: false,
    };
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            object
                .inspect_references_with_policy_and_limits(
                    &mut census,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
    }
    if census.invalid
        || census.model_edges == 0
        || !validate_selected_edge_paths(
            package,
            slide_identifier,
            slide_message_index,
            table_info_identifier,
            table_info_message_index,
            model_identifier,
        )?
    {
        return Err(SlideTableSortError::UnsupportedDependency);
    }
    Ok(())
}

fn validate_selected_edge_paths(
    package: &Package,
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
) -> Result<bool, SlideTableSortError> {
    let Some((_slide_component, slide)) = package.object_with_component(slide_identifier) else {
        return Ok(false);
    };
    let Some(slide_info) = slide.archive_info.message_infos.get(slide_message_index) else {
        return Ok(false);
    };
    for field in &slide_info.field_infos {
        if field.object_references.contains(&table_info_identifier)
            && field.path.as_slice() != [SLIDE_OWNED_DRAWABLES_FIELD]
            && field.path.as_slice() != [SLIDE_Z_ORDER_FIELD]
        {
            return Ok(false);
        }
    }
    let Some((_table_component, table_info)) = package.object_with_component(table_info_identifier)
    else {
        return Ok(false);
    };
    let Some(table_info_info) = table_info
        .archive_info
        .message_infos
        .get(table_info_message_index)
    else {
        return Ok(false);
    };
    let mut model_field_count = 0usize;
    for field in &table_info_info.field_infos {
        if field.object_references.contains(&model_identifier)
            && field.path.as_slice() != [TABLE_MODEL_FIELD]
        {
            return Ok(false);
        }
        if field.path.as_slice() == [TABLE_MODEL_FIELD]
            && field.object_references.as_slice() != [model_identifier]
        {
            return Ok(false);
        }
        if field.path.as_slice() == [TABLE_MODEL_FIELD] {
            model_field_count = model_field_count.saturating_add(1);
        }
    }
    Ok(model_field_count <= 1)
}

struct InboundReferenceCensus {
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    model_edges: usize,
    invalid: bool,
}

impl ArchiveReferenceVisitor for InboundReferenceCensus {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.referenced_identifier == self.model_identifier {
            if occurrence.object_identifier == self.table_info_identifier
                && occurrence.message_index == self.table_info_message_index
            {
                self.model_edges = self.model_edges.saturating_add(1);
            } else {
                self.invalid = true;
            }
        }
        if occurrence.referenced_identifier == self.table_info_identifier
            && (occurrence.object_identifier != self.slide_identifier
                || occurrence.message_index != self.slide_message_index)
        {
            self.invalid = true;
        }
        // Keep the scope in the implementation so the exception remains
        // visibly limited to the rooted aggregate/field metadata edge; the
        // core visitor itself has already rejected opaque unknown fields.
        let _scope_is_known = matches!(
            occurrence.scope,
            ArchiveReferenceScope::Message | ArchiveReferenceScope::Field { .. }
        );
        Ok(())
    }
}

fn same_selection(a: &SortSelection, b: &SortSelection) -> bool {
    a.slide_position == b.slide_position
        && a.table_position == b.table_position
        && a.slide_identifier == b.slide_identifier
        && a.table_info_identifier == b.table_info_identifier
        && a.model_identifier == b.model_identifier
        && a.model_message_index == b.model_message_index
        && a.component_name == b.component_name
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideTableSortError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTableSortError::UnsupportedSource),
    }
}

fn map_read_error(error: ReadError) -> SlideTableSortError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableSortError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableSortLimitKind::References,
                _ => SlideTableSortLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableSortError::Allocation { amount },
        _ => SlideTableSortError::InvalidSource,
    }
}
fn map_codec_error(error: codec::DecodeError) -> SlideTableSortError {
    if let Some(amount) = error.allocation_amount() {
        return SlideTableSortError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return SlideTableSortError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::InputBytes { observed, maximum } => {
            (SlideTableSortLimitKind::InputBytes, observed, maximum)
        },
        codec::DecodeLimit::OutputBytes { observed, maximum } => {
            (SlideTableSortLimitKind::OutputBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (SlideTableSortLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::WorkBytes { observed, maximum } => {
            (SlideTableSortLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => (
            SlideTableSortLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        codec::DecodeLimit::Rules { observed, maximum } => {
            (SlideTableSortLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Columns { observed, maximum } => {
            (SlideTableSortLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Allocations { observed, maximum } => {
            (SlideTableSortLimitKind::Allocations, observed, maximum)
        },
        codec::DecodeLimit::RetainedBytes { observed, maximum } => {
            (SlideTableSortLimitKind::Retained, observed, maximum)
        },
        codec::DecodeLimit::ScratchBytes { observed, maximum } => {
            (SlideTableSortLimitKind::Scratch, observed, maximum)
        },
        _ => return SlideTableSortError::InvalidSource,
    };
    SlideTableSortError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}
fn map_wire_error(_error: litchi_iwa_common::Error) -> SlideTableSortError {
    SlideTableSortError::InvalidSource
}
fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTableSortError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableSortError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideTableSortLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SlideTableSortLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SlideTableSortLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideTableSortLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => SlideTableSortLimitKind::TotalBytes,
                _ => SlideTableSortLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableSortError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTableSortError::InvalidSource,
    }
}
fn map_core_error(error: litchi_iwa_core::Error) -> SlideTableSortError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableSortError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => SlideTableSortLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableSortLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => SlideTableSortLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => SlideTableSortLimitKind::EntryBytes,
                _ => SlideTableSortLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableSortError::Allocation { amount: requested }
        },
        _ => SlideTableSortError::InvalidSource,
    }
}
