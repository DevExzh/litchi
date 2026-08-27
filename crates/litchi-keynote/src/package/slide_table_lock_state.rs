//! Exact-source transactions for persisted Keynote slide-table lock state.
//!
//! This owner exposes only a semantic lock value and checked slide/table
//! selectors. Native identifiers, archive objects, protobuf messages, and
//! package bytes stay private to this adapter. The transaction rewrites only
//! the optional lock field in one canonical TableInfo object.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "the package boundary redacts lower-layer failure details"
)]

use std::sync::Arc;
use std::{collections::HashSet, fmt};

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceScope, ArchiveReferenceVisitor, RawMessage,
    SnappyStream,
};
use litchi_iwa_protos::{package_metadata_codec, table_info_codec, table_model_discovery_codec};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::TableSelector;
use crate::slide::table::lock::State;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const LEGACY_TABLE_MODEL_MESSAGE_TYPE: u32 = 6_000;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;

/// Finite resource categories enforced by one slide-table lock transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableLockStateLimitKind {
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

impl fmt::Display for SlideTableLockStateLimitKind {
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

/// Content-free semantic path for a slide-table lock operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableLockStatePath {
    Package,
    Table { slide: Position, table: Position },
}

impl fmt::Display for SlideTableLockStatePath {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Table { slide, table } => {
                write!(formatter, "slide {} table {}", slide.get(), table.get())
            },
        }
    }
}

/// Failure from a Keynote slide-table lock read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTableLockStateError {
    #[error("this Keynote source does not support persisted slide-table lock edits")]
    UnsupportedSource,
    #[error("the requested Keynote slide-table lock graph is outside the supported scope")]
    UnsupportedDependency,
    #[error("the Keynote slide-table lock selector is ambiguous")]
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
    #[error("the selected Keynote slide-table lock source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table lock {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTableLockStateLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote slide-table lock transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote slide-table lock failed semantic verification")]
    Verification,
    #[error("the Keynote slide-table lock patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct LockBudget {
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    max_entries: usize,
    max_entry_bytes: usize,
    max_total_bytes: usize,
    max_payload_objects: usize,
    max_payload_messages: usize,
    max_components: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
    entries: usize,
    entry_bytes: usize,
    total_bytes: usize,
    payload_objects: usize,
    payload_messages: usize,
    components: usize,
}

impl LockBudget {
    fn new(package: &Package) -> Result<Self, SlideTableLockStateError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let source = usize::try_from(package.state.options.archive().max_input_bytes())
            .map_err(|_| SlideTableLockStateError::InvalidSource)?;
        let outer_limits = package.state.options.archive();
        let archive_limits = outer_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        // One exact transaction may inventory the source, the rewritten
        // candidate, and the candidate again during semantic/locality reopen.
        // Keep these physical axes aggregate across those bounded passes.
        let operation_passes = 4usize;
        let max_components = outer_limits
            .max_entries()
            .checked_mul(operation_passes)
            .unwrap_or(usize::MAX);
        let max_payload_objects = archive_limits
            .max_objects()
            .checked_mul(max_components)
            .unwrap_or(usize::MAX);
        let max_payload_messages = archive_limits
            .max_messages()
            .checked_mul(max_components)
            .unwrap_or(usize::MAX);
        let aggregate = source
            .checked_mul(4)
            .ok_or(SlideTableLockStateError::InvalidSource)?;
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
            max_entries: outer_limits
                .max_entries()
                .checked_mul(operation_passes)
                .unwrap_or(usize::MAX),
            max_entry_bytes: usize::try_from(outer_limits.max_total_bytes())
                .map_err(|_| SlideTableLockStateError::InvalidSource)?
                .checked_mul(operation_passes)
                .unwrap_or(usize::MAX),
            max_total_bytes: usize::try_from(outer_limits.max_total_bytes())
                .map_err(|_| SlideTableLockStateError::InvalidSource)?
                .checked_mul(operation_passes)
                .unwrap_or(usize::MAX),
            max_payload_objects,
            max_payload_messages,
            max_components,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
            entries: 0,
            entry_bytes: 0,
            total_bytes: 0,
            payload_objects: 0,
            payload_messages: 0,
            components: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideTableLockStateLimitKind,
    ) -> Result<(), SlideTableLockStateError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideTableLockStateError::InvalidSource)?;
        if observed > maximum {
            return Err(SlideTableLockStateError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn input(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            SlideTableLockStateLimitKind::InputBytes,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            SlideTableLockStateLimitKind::OutputBytes,
        )
    }

    fn fields(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            SlideTableLockStateLimitKind::WireFields,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            SlideTableLockStateLimitKind::WireWork,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideTableLockStateLimitKind::References,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideTableLockStateLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideTableLockStateLimitKind::Retained,
        )
    }

    fn scratch(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            SlideTableLockStateLimitKind::Scratch,
        )
    }

    fn entries(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.entries,
            amount,
            self.max_entries,
            SlideTableLockStateLimitKind::Entries,
        )
    }

    fn entry_bytes(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.entry_bytes,
            amount,
            self.max_entry_bytes,
            SlideTableLockStateLimitKind::EntryBytes,
        )
    }

    fn total_bytes(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.total_bytes,
            amount,
            self.max_total_bytes,
            SlideTableLockStateLimitKind::TotalBytes,
        )
    }

    fn payload_objects(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.payload_objects,
            amount,
            self.max_payload_objects,
            SlideTableLockStateLimitKind::PayloadObjects,
        )
    }

    fn payload_messages(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.payload_messages,
            amount,
            self.max_payload_messages,
            SlideTableLockStateLimitKind::PayloadMessages,
        )
    }

    fn components(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        Self::add(
            &mut self.components,
            amount,
            self.max_components,
            SlideTableLockStateLimitKind::Components,
        )
    }

    fn inventory(
        &mut self,
        package: &Package,
        catalog: &litchi_iwa_archive::SourceCatalog,
    ) -> Result<(), SlideTableLockStateError> {
        self.entries(catalog.package().len())?;
        self.components(package.state.source.components().len())?;
        for entry in catalog.package().iter() {
            let bytes = entry.data().len();
            self.entry_bytes(bytes)?;
            self.total_bytes(bytes)?;
        }
        for component in package.state.source.components().iter() {
            self.payload_objects(component.archive().objects.len())?;
            let message_count =
                component
                    .archive()
                    .objects
                    .iter()
                    .try_fold(0usize, |count, object| {
                        count
                            .checked_add(object.messages.len())
                            .ok_or(SlideTableLockStateError::InvalidSource)
                    })?;
            self.payload_messages(message_count)?;
        }
        Ok(())
    }

    fn physical(&mut self, amount: usize) -> Result<(), SlideTableLockStateError> {
        self.input(amount)?;
        self.work(amount)
    }

    fn codec_requirements(
        &mut self,
        requirements: table_info_codec::RewriteExecutionRequirements,
    ) -> Result<(), SlideTableLockStateError> {
        self.input(requirements.input_bytes())?;
        self.output(requirements.output_bytes())?;
        self.fields(requirements.fields())?;
        self.work(requirements.work_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.nesting = self.nesting.max(requirements.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableLockStateError::LimitExceeded {
                kind: SlideTableLockStateLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn metadata_report(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<(), SlideTableLockStateError> {
        self.input(report.input_bytes())?;
        self.output(report.output_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.components(report.components_scanned())?;
        self.references(
            report
                .references_scanned()
                .checked_add(report.source_references_scanned())
                .ok_or(SlideTableLockStateError::InvalidSource)?,
        )?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableLockStateError::LimitExceeded {
                kind: SlideTableLockStateLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideTableLockStateError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.work(requirements.output_bytes())
    }

    fn residual(&self, package: &Package) -> Result<WireLimits, SlideTableLockStateError> {
        let base = package.wire_limits().map_err(map_wire_error)?;
        base.with_input_bytes(
            base.max_input_bytes()
                .min(self.max_input.saturating_sub(self.input).max(1)),
        )
        .and_then(|limits| {
            limits.with_fields(
                base.max_fields()
                    .min(self.max_fields.saturating_sub(self.fields).max(1)),
            )
        })
        .and_then(|limits| {
            limits.with_output_bytes(
                base.max_output_bytes()
                    .min(self.max_output.saturating_sub(self.output).max(1)),
            )
        })
        .and_then(|limits| {
            limits.with_rewrite_work(
                base.max_rewrite_work()
                    .min(self.max_work.saturating_sub(self.work).max(1)),
            )
        })
        .and_then(|limits| limits.with_nesting(base.max_nesting().min(self.max_nesting)))
        .map_err(map_wire_error)
    }
}

/// One mutable lock value staged against an immutable package snapshot.
pub struct SlideTableLockStateEdit<'a> {
    source: &'a Package,
    selection: LockSelection,
    after: State,
}

impl fmt::Debug for SlideTableLockStateEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableLockStateEdit")
            .field("path", &self.selection.path())
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableLockStateEdit<'_> {
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    #[must_use]
    pub const fn table_position(&self) -> Position {
        self.selection.table_position
    }

    #[must_use]
    pub const fn path(&self) -> SlideTableLockStatePath {
        self.selection.path()
    }

    #[must_use]
    pub const fn before(&self) -> State {
        self.selection.before
    }

    #[must_use]
    pub const fn state(&self) -> State {
        self.after
    }

    #[must_use]
    pub const fn after(&self) -> State {
        self.after
    }

    #[must_use]
    pub fn set(mut self, state: State) -> Self {
        self.after = state;
        self
    }

    pub fn set_state(&mut self, state: State) -> &mut Self {
        self.after = state;
        self
    }

    pub fn lock(&mut self) -> &mut Self {
        self.set_state(State::Locked)
    }

    pub fn unlock(&mut self) -> &mut Self {
        self.set_state(State::Unlocked)
    }

    pub fn commit(self) -> Result<SlideTableLockStateCommit, SlideTableLockStateError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible slide-table lock patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableLockStatePatch {
    artifacts: ExactArtifacts,
    selection: LockSelection,
    before: State,
    after: State,
}

impl fmt::Debug for SlideTableLockStatePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableLockStatePatch")
            .field("path", &self.selection.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableLockStatePatch {
    #[must_use]
    pub const fn before(&self) -> State {
        self.before
    }

    #[must_use]
    pub const fn after(&self) -> State {
        self.after
    }

    #[must_use]
    pub const fn path(&self) -> SlideTableLockStatePath {
        self.selection.path()
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
            before: self.after,
            after: self.before,
        }
    }
}

/// Compact publication diagnostics for one slide-table lock transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableLockStateDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideTableLockStateDiagnostics {
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

/// Fully verified result of one slide-table lock transaction.
#[must_use = "a Keynote slide-table lock commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableLockStateCommit {
    package: Package,
    patch: SlideTableLockStatePatch,
    diagnostics: SlideTableLockStateDiagnostics,
}

impl SlideTableLockStateCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &SlideTableLockStatePatch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableLockStateDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
struct LockSelection {
    slide_position: Position,
    table_position: Position,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    slide_message_index: usize,
    table_info_message_index: usize,
    model_message_index: usize,
    component_name: Arc<str>,
    before: State,
    locked: bool,
    admission_budget: LockBudget,
}

impl fmt::Debug for LockSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("LockSelection")
            .field("slide_position", &self.slide_position)
            .field("table_position", &self.table_position)
            .field("before", &self.before)
            .field("locked", &self.locked)
            .finish_non_exhaustive()
    }
}

impl LockSelection {
    const fn path(&self) -> SlideTableLockStatePath {
        SlideTableLockStatePath::Table {
            slide: self.slide_position,
            table: self.table_position,
        }
    }
}

#[derive(Clone, Copy)]
struct LockCandidate {
    table_info_identifier: u64,
    model_identifier: u64,
    info_message_index: usize,
    model_message_index: usize,
    before: State,
}

impl Package {
    /// Read one existing slide table's persisted interactive lock state.
    pub fn slide_table_lock_state<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<State, SlideTableLockStateError> {
        Ok(select_table(self, slide.into(), table.into())?.before)
    }

    /// Start an immutable exact edit of one slide table's persisted lock state.
    pub fn edit_slide_table_lock_state<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTableLockStateEdit<'_>, SlideTableLockStateError> {
        let selection = select_table(self, slide.into(), table.into())?;
        Ok(SlideTableLockStateEdit {
            source: self,
            after: selection.before,
            selection,
        })
    }

    /// Apply an exact-source checked reversible slide-table lock patch.
    pub fn apply_slide_table_lock_state(
        &self,
        patch: &SlideTableLockStatePatch,
    ) -> Result<SlideTableLockStateCommit, SlideTableLockStateError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideTableLockStateError::PatchConflict);
        }
        let current = select_table(
            self,
            SlideSelector::position(patch.selection.slide_position),
            TableSelector::position(patch.selection.table_position),
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SlideTableLockStateError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(SlideTableLockStateCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableLockStateDiagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch, current.admission_budget)
    }
}

fn commit_edit(
    source: &Package,
    selection: &LockSelection,
    after: State,
) -> Result<SlideTableLockStateCommit, SlideTableLockStateError> {
    let catalog = physical_catalog(source)?;
    if selection.before == after {
        let bytes = catalog.shared_source();
        return Ok(SlideTableLockStateCommit {
            package: source.snapshot(),
            patch: SlideTableLockStatePatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before,
                after,
            },
            diagnostics: SlideTableLockStateDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SlideTableLockStateError::UnsupportedSource);
    }
    let mut budget = selection.admission_budget;
    let candidate = rewrite_lock(source, selection, after, &mut budget)?;
    budget.work(candidate.source_bytes().len())?;
    budget.allocations(1)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_table_with_budget(
        &candidate,
        SlideSelector::position(selection.slide_position),
        TableSelector::position(selection.table_position),
        &mut budget,
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(SlideTableLockStateError::Verification);
    }
    verify_locality(source, &candidate, selection, &mut budget)?;
    let target = physical_catalog(&candidate)?.shared_source();
    Ok(SlideTableLockStateCommit {
        package: candidate,
        patch: SlideTableLockStatePatch {
            artifacts: ExactArtifacts::new(Arc::from(source.source_bytes()), target),
            selection: selection.clone(),
            before: selection.before,
            after,
        },
        diagnostics: SlideTableLockStateDiagnostics::published(),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &SlideTableLockStatePatch,
    mut budget: LockBudget,
) -> Result<SlideTableLockStateCommit, SlideTableLockStateError> {
    budget.input(patch.artifacts.target().len())?;
    budget.work(patch.artifacts.target().len())?;
    budget.allocations(physical_catalog(source)?.package().len())?;
    budget.retained(patch.artifacts.target().len())?;
    budget.scratch(patch.artifacts.target().len())?;
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_table_with_budget(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        TableSelector::position(patch.selection.table_position),
        &mut budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SlideTableLockStateError::Verification);
    }
    verify_locality(source, &candidate, &patch.selection, &mut budget)?;
    Ok(SlideTableLockStateCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableLockStateDiagnostics::published(),
    })
}

fn select_table(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
) -> Result<LockSelection, SlideTableLockStateError> {
    let mut budget = LockBudget::new(package)?;
    select_table_with_budget(package, slide_selector, table_selector, &mut budget)
}

fn select_table_with_budget(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
    budget: &mut LockBudget,
) -> Result<LockSelection, SlideTableLockStateError> {
    let catalog = physical_catalog(package)?;
    budget.input(package.source_bytes().len())?;
    budget.inventory(package, catalog)?;
    budget.allocations(catalog.package().len())?;
    validate_package_metadata(package, budget)?;
    // `slide_record_at` and name selectors share the lazy semantic document
    // decoder. Reserve a conservative package-sized work envelope before
    // entering it; later graph scans continue consuming this same budget.
    budget.work(package.source_bytes().len())?;
    budget.allocations(1)?;
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTableLockStateError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (_slide_component, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let (slide_message_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE, budget)?;

    let wire_limits = budget.residual(package)?;
    let owned = repeated_references(
        slide_payload,
        SLIDE_OWNED_DRAWABLES_FIELD,
        wire_limits,
        budget,
    )?;
    let z_order = repeated_references(slide_payload, SLIDE_Z_ORDER_FIELD, wire_limits, budget)?;
    budget.references(
        owned
            .len()
            .checked_add(z_order.len())
            .ok_or(SlideTableLockStateError::InvalidSource)?,
    )?;
    reject_duplicates(&owned, budget)?;
    reject_duplicates(&z_order, budget)?;
    validate_slide_metadata(slide, slide_message_index, &owned, &z_order, budget)?;

    let mut candidates = Vec::new();
    budget.allocations(z_order.len())?;
    candidates.try_reserve_exact(z_order.len()).map_err(|_| {
        SlideTableLockStateError::Allocation {
            amount: z_order.len(),
        }
    })?;

    for table_info_identifier in z_order {
        let Some((_owner_component, info_object)) =
            package.object_with_component(table_info_identifier)
        else {
            return Err(SlideTableLockStateError::InvalidSource);
        };
        let info_count = info_object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .count();
        let has_role_alias = info_object.messages.iter().any(|message| {
            matches!(
                message.type_,
                TABLE_MODEL_MESSAGE_TYPE
                    | TABLE_STYLE_MESSAGE_TYPE
                    | TABLE_STYLE_PRESET_MESSAGE_TYPE
                    | TABLE_STYLE_NETWORK_MESSAGE_TYPE
                    | STYLESHEET_MESSAGE_TYPE
            )
        });
        if info_count == 0 {
            if has_role_alias {
                return Err(SlideTableLockStateError::UnsupportedDependency);
            }
            continue;
        }
        if info_count != 1 || has_role_alias {
            return Err(SlideTableLockStateError::UnsupportedDependency);
        }
        if owned
            .iter()
            .filter(|identifier| **identifier == table_info_identifier)
            .count()
            != 1
        {
            return Err(SlideTableLockStateError::InvalidSource);
        }
        let (info_message_index, info_payload) =
            unique_message(info_object, TABLE_INFO_MESSAGE_TYPE, budget)?;
        let info = decode_table_info(info_payload, package, budget)?;
        let parent = table_parent(info_payload, wire_limits, budget)?;
        if parent != record.slide_identifier {
            return Err(SlideTableLockStateError::InvalidSource);
        }
        let model_identifier = info.table_model().identifier().get();
        validate_table_info_metadata(
            info_object,
            info_message_index,
            record.slide_identifier,
            model_identifier,
            budget,
        )?;
        let (model_component, model) = package
            .object_with_component(model_identifier)
            .ok_or(SlideTableLockStateError::InvalidSource)?;
        let model_count = model
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .count();
        let has_model_role_alias = model.messages.iter().any(|message| {
            matches!(
                message.type_,
                LEGACY_TABLE_MODEL_MESSAGE_TYPE
                    | TABLE_STYLE_MESSAGE_TYPE
                    | TABLE_STYLE_PRESET_MESSAGE_TYPE
                    | TABLE_STYLE_NETWORK_MESSAGE_TYPE
                    | STYLESHEET_MESSAGE_TYPE
            )
        });
        if model_count != 1 || has_model_role_alias {
            return Err(SlideTableLockStateError::UnsupportedDependency);
        }
        let (model_message_index, model_payload) =
            unique_message(model, TABLE_MODEL_MESSAGE_TYPE, budget)?;
        validate_model_payload(model_payload, package, budget)?;
        candidates.push(LockCandidate {
            table_info_identifier,
            model_identifier,
            info_message_index,
            model_message_index,
            before: State::from_locked(info.locked().unwrap_or(false)),
        });
        let _ = model_component;
    }

    let table_position = table_selector.as_position();
    let candidate = candidates.get(table_position.get()).copied().ok_or(
        SlideTableLockStateError::TablePositionNotFound {
            position: table_position,
        },
    )?;
    let (component_name, _info_object) = package
        .object_with_component(candidate.table_info_identifier)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    ensure_unique_identity(package, candidate.table_info_identifier, budget)?;
    ensure_unique_identity(package, candidate.model_identifier, budget)?;
    ensure_unique_table_owner(
        package,
        record.slide_identifier,
        candidate.table_info_identifier,
        candidate.model_identifier,
        wire_limits,
        budget,
    )?;
    validate_global_inbound_references(
        package,
        record.slide_identifier,
        candidate.table_info_identifier,
        candidate.model_identifier,
        budget,
    )?;

    Ok(LockSelection {
        slide_position,
        table_position,
        slide_identifier: record.slide_identifier,
        table_info_identifier: candidate.table_info_identifier,
        model_identifier: candidate.model_identifier,
        slide_message_index,
        table_info_message_index: candidate.info_message_index,
        model_message_index: candidate.model_message_index,
        component_name: Arc::from(component_name),
        before: candidate.before,
        locked: candidate.before.is_locked(),
        admission_budget: *budget,
    })
}

fn validate_package_metadata(
    package: &Package,
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

    let mut payload = None;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                budget.work(
                    message
                        .data
                        .len()
                        .checked_add(1)
                        .ok_or(SlideTableLockStateError::InvalidSource)?,
                )?;
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    if payload.replace(message.data.as_slice()).is_some() {
                        return Err(SlideTableLockStateError::UnsupportedDependency);
                    }
                }
            }
        }
    }
    let payload = payload.ok_or(SlideTableLockStateError::InvalidSource)?;
    let wire = budget.residual(package)?;
    let recursion =
        u32::try_from(wire.max_nesting()).map_err(|_| SlideTableLockStateError::InvalidSource)?;
    let remaining_references = budget
        .max_references
        .saturating_sub(budget.references)
        .max(1);
    let remaining_components = budget
        .max_components
        .saturating_sub(budget.components)
        .max(1);
    let options = package_metadata_codec::RewriteOptions::new(
        wire.max_input_bytes().min(payload.len().max(1)),
        wire.max_output_bytes(),
        wire.max_fields(),
        wire.max_rewrite_work(),
        recursion,
        remaining_components,
        remaining_references,
        remaining_references,
    );
    let mut visitor = StrictPackageMetadataVisitor { unknown: false };
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        payload,
        options,
        &mut visitor,
    )
    .map_err(map_metadata_error)?;
    budget.metadata_report(inspection.report())?;
    if visitor.unknown {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    Ok(())
}

#[derive(Default)]
struct StrictPackageMetadataVisitor {
    unknown: bool,
}

impl package_metadata_codec::PackageMetadataVisitor for StrictPackageMetadataVisitor {
    fn visit_unknown_field(&mut self) -> Result<(), package_metadata_codec::RewriteError> {
        self.unknown = true;
        Ok(())
    }
}

fn rewrite_lock(
    source: &Package,
    selection: &LockSelection,
    after: State,
    budget: &mut LockBudget,
) -> Result<Package, SlideTableLockStateError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.component_name.as_ref())
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableLockStateError::UnsupportedSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;

    budget.physical(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    budget.scratch(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    charge_archive_inventory(&archive, budget)?;
    let object = archive
        .object(selection.table_info_identifier)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let original = object
        .messages
        .get(selection.table_info_message_index)
        .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .ok_or(SlideTableLockStateError::InvalidSource)?
        .data
        .as_slice();
    let current = decode_table_info(original, source, budget)?;
    if State::from_locked(current.locked().unwrap_or(false)) != selection.before {
        return Err(SlideTableLockStateError::InvalidSource);
    }

    let write = table_info_codec::TableInfoLockWrite::from_locked(after.is_locked());
    let residual = budget.residual(source)?;
    let recursion = u32::try_from(residual.max_nesting())
        .map_err(|_| SlideTableLockStateError::InvalidSource)?;
    let options = table_info_codec::DecodeOptions::new(
        residual.max_input_bytes().min(original.len().max(1)),
        residual.max_fields(),
        residual.max_rewrite_work(),
        recursion,
    )
    .with_max_output_bytes(residual.max_output_bytes())
    .with_max_allocations(
        budget
            .max_allocations
            .saturating_sub(budget.allocations)
            .max(1),
    )
    .with_max_retained_bytes(budget.max_retained.saturating_sub(budget.retained).max(1))
    .with_max_scratch_bytes(budget.max_scratch.saturating_sub(budget.scratch).max(1));
    let prepared = table_info_codec::prepare_table_info_lock_rewrite(original, write, options)
        .map_err(map_table_info_error)?;
    let requirements = prepared.execution_requirements();
    budget.codec_requirements(requirements)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_table_info_error)?;
    let report = output.report();
    if report.input_bytes() != requirements.input_bytes()
        || report.output_bytes() != requirements.output_bytes()
        || report.fields() != requirements.fields()
        || report.work_bytes() != requirements.work_bytes()
        || report.max_depth() != requirements.max_depth()
        || report.allocations() != requirements.allocations()
        || report.retained_bytes() != requirements.retained_bytes()
        || report.scratch_bytes() != requirements.scratch_bytes()
    {
        return Err(SlideTableLockStateError::Verification);
    }
    let rewritten = output.into_bytes();
    let verified = decode_table_info(&rewritten, source, budget)?;
    if State::from_locked(verified.locked().unwrap_or(false)) != after {
        return Err(SlideTableLockStateError::Verification);
    }

    let object = archive
        .object_mut(selection.table_info_identifier)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            selection.table_info_message_index,
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;

    let encoded_len = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.output(encoded_len)?;
    budget.allocations(1)?;
    budget.retained(encoded_len)?;
    budget.scratch(encoded_len)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if bytes.len() != encoded_len {
        return Err(SlideTableLockStateError::Verification);
    }
    budget.work(bytes.len())?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_len).map_err(map_core_error)?;
    budget.output(compressed_bound)?;
    budget.allocations(1)?;
    budget.retained(compressed_bound)?;
    budget.scratch(compressed_bound)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(SlideTableLockStateError::Verification);
    }

    let entry_edits = [EntryEdit::new(
        selection.component_name.as_ref(),
        &compressed,
    )];
    budget.allocations(catalog.package().len())?;
    budget.work(catalog.package().len())?;
    let prepared_reassembly = catalog
        .prepare_reassembly_with_deletions(&entry_edits, &[], physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared_reassembly.execution_requirements();
    budget.reassembly(requirements)?;
    let output = prepared_reassembly
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    budget.input(output.len())?;
    budget.work(output.len())?;
    budget.allocations(catalog.package().len())?;
    budget.retained(output.len())?;
    budget.scratch(output.len())?;
    Package::from_source_with_options(output.into(), source.state.options).map_err(map_read_error)
}

fn decode_table_info(
    payload: &[u8],
    package: &Package,
    budget: &mut LockBudget,
) -> Result<table_info_codec::TableInfoSnapshot, SlideTableLockStateError> {
    let limits = budget.residual(package)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideTableLockStateError::InvalidSource)?;
    budget.work(
        payload
            .len()
            .checked_mul(2)
            .ok_or(SlideTableLockStateError::InvalidSource)?,
    )?;
    table_info_codec::decode_table_info(
        payload,
        table_info_codec::DecodeOptions::new(
            limits.max_input_bytes().min(payload.len().max(1)),
            limits.max_fields(),
            limits.max_rewrite_work(),
            recursion,
        ),
    )
    .map_err(map_table_info_error)
}

fn validate_model_payload(
    payload: &[u8],
    package: &Package,
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    let limits = budget.residual(package)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideTableLockStateError::InvalidSource)?;
    let options = table_model_discovery_codec::DecodeOptions::for_source(payload)
        .with_max_input_bytes(limits.max_input_bytes().min(payload.len().max(1)))
        .with_max_fields(limits.max_fields())
        .with_max_work_bytes(limits.max_rewrite_work())
        .with_max_text_bytes(limits.max_input_bytes())
        .with_recursion_limit(recursion);
    let (_, report) = table_model_discovery_codec::decode_table_model_with_report(payload, options)
        .map_err(map_table_model_error)?;
    budget.input(report.input_bytes())?;
    budget.fields(report.fields())?;
    budget.work(report.work_bytes())?;
    budget.allocations(report.allocations())?;
    budget.retained(report.retained_bytes())?;
    budget.scratch(report.scratch_bytes())?;
    let depth = report.max_depth() as usize;
    if depth > budget.nesting {
        budget.nesting = depth;
    }
    if budget.nesting > budget.max_nesting {
        return Err(SlideTableLockStateError::LimitExceeded {
            kind: SlideTableLockStateLimitKind::WireNesting,
            observed: budget.nesting as u64,
            maximum: budget.max_nesting as u64,
        });
    }
    Ok(())
}

fn validate_slide_metadata(
    object: &ArchiveObject,
    message_index: usize,
    owned: &[u64],
    z_order: &[u64],
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    if info.object_references.is_empty() {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    if !info.data_references.is_empty() {
        return Err(SlideTableLockStateError::UnsupportedDependency);
    }
    reject_duplicates(&info.object_references, budget)?;
    budget.allocations(info.object_references.len())?;
    let mut declared = HashSet::new();
    declared
        .try_reserve(info.object_references.len())
        .map_err(|_| SlideTableLockStateError::Allocation {
            amount: info.object_references.len(),
        })?;
    for identifier in &info.object_references {
        declared.insert(*identifier);
    }
    for identifier in owned.iter().chain(z_order) {
        budget.work(1)?;
        if !declared.contains(identifier) {
            return Err(SlideTableLockStateError::InvalidSource);
        }
    }
    budget.references(info.object_references.len())?;
    budget.fields(info.field_infos.len())?;
    budget.work(info.field_infos.len())?;
    let mut saw_owned = false;
    let mut saw_z_order = false;
    for field in &info.field_infos {
        budget.references(
            field
                .object_references
                .len()
                .checked_add(field.data_references.len())
                .ok_or(SlideTableLockStateError::InvalidSource)?,
        )?;
        if !field.data_references.is_empty() {
            return Err(SlideTableLockStateError::UnsupportedDependency);
        }
        match field.path.as_slice() {
            [SLIDE_OWNED_DRAWABLES_FIELD] => {
                if saw_owned
                    || !is_message_reference_field(field)
                    || field.object_references.as_slice() != owned
                {
                    return Err(SlideTableLockStateError::InvalidSource);
                }
                saw_owned = true;
            },
            [SLIDE_Z_ORDER_FIELD] => {
                if saw_z_order
                    || !is_message_reference_field(field)
                    || field.object_references.as_slice() != z_order
                {
                    return Err(SlideTableLockStateError::InvalidSource);
                }
                saw_z_order = true;
            },
            _ if !field.object_references.is_empty() => {
                return Err(SlideTableLockStateError::UnsupportedDependency);
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_table_info_metadata(
    object: &ArchiveObject,
    message_index: usize,
    parent: u64,
    model: u64,
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    if info.object_references.is_empty() {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    if !info.data_references.is_empty() {
        return Err(SlideTableLockStateError::UnsupportedDependency);
    }
    reject_duplicates(&info.object_references, budget)?;
    if info
        .object_references
        .iter()
        .filter(|value| **value == parent)
        .count()
        != 1
        || info
            .object_references
            .iter()
            .filter(|value| **value == model)
            .count()
            != 1
    {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    budget.references(info.object_references.len())?;
    budget.fields(info.field_infos.len())?;
    budget.work(info.field_infos.len())?;
    let mut saw_model = false;
    let mut saw_parent = false;
    for field in &info.field_infos {
        budget.references(
            field
                .object_references
                .len()
                .checked_add(field.data_references.len())
                .ok_or(SlideTableLockStateError::InvalidSource)?,
        )?;
        if !field.data_references.is_empty() {
            return Err(SlideTableLockStateError::UnsupportedDependency);
        }
        match field.path.as_slice() {
            [TABLE_MODEL_FIELD] => {
                if saw_model
                    || !is_message_reference_field(field)
                    || field.object_references.as_slice() != [model]
                    || !field.data_references.is_empty()
                {
                    return Err(SlideTableLockStateError::InvalidSource);
                }
                saw_model = true;
            },
            [TABLE_SUPER_FIELD, DRAWABLE_PARENT_FIELD] => {
                if saw_parent
                    || !is_message_reference_field(field)
                    || field.object_references.as_slice() != [parent]
                    || !field.data_references.is_empty()
                {
                    return Err(SlideTableLockStateError::InvalidSource);
                }
                saw_parent = true;
            },
            _ if !field.object_references.is_empty() => {
                return Err(SlideTableLockStateError::UnsupportedDependency);
            },
            _ => {},
        }
    }
    Ok(())
}

fn is_message_reference_field(field: &litchi_iwa_core::FieldInfo) -> bool {
    field
        .r#type
        .is_none_or(|field_type| field_type == litchi_iwa_core::FieldType::Message)
}

fn validate_message_header(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), SlideTableLockStateError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    let message = object
        .messages
        .get(message_index)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    if message.type_ != info.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    Ok(())
}

fn unique_message<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    budget: &mut LockBudget,
) -> Result<(usize, &'a [u8]), SlideTableLockStateError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        budget.work(
            message
                .data
                .len()
                .checked_add(1)
                .ok_or(SlideTableLockStateError::InvalidSource)?,
        )?;
        validate_message_header(object, index)?;
        if message.type_ == message_type {
            if selected.is_some() {
                return Err(SlideTableLockStateError::UnsupportedDependency);
            }
            selected = Some((index, message.data.as_slice()));
        }
    }
    selected.ok_or(SlideTableLockStateError::InvalidSource)
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut LockBudget,
) -> Result<Vec<u64>, SlideTableLockStateError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let count = fields
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    let mut result = Vec::new();
    budget.allocations(count)?;
    result
        .try_reserve_exact(count)
        .map_err(|_| SlideTableLockStateError::Allocation { amount: count })?;
    budget.fields(count)?;
    budget.work(payload.len())?;
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideTableLockStateError::InvalidSource);
        }
        result.push(strict_reference(field.payload(), limits, budget)?);
    }
    Ok(result)
}

fn table_parent(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LockBudget,
) -> Result<u64, SlideTableLockStateError> {
    budget.work(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut super_payload = None;
    let mut super_count = 0usize;
    for field in fields.fields() {
        budget.fields(1)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == TABLE_SUPER_FIELD {
            super_count = super_count
                .checked_add(1)
                .ok_or(SlideTableLockStateError::InvalidSource)?;
            super_payload = Some(field);
        }
    }
    let Some(super_field) = super_payload else {
        return Err(SlideTableLockStateError::InvalidSource);
    };
    if super_count != 1 || super_field.wire_type() != 2 {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    let drawable =
        WireView::parse_with_limits(super_field.payload(), limits).map_err(map_wire_error)?;
    let mut parent_payload = None;
    let mut parent_count = 0usize;
    for field in drawable.fields() {
        budget.fields(1)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == DRAWABLE_PARENT_FIELD {
            parent_count = parent_count
                .checked_add(1)
                .ok_or(SlideTableLockStateError::InvalidSource)?;
            parent_payload = Some(field);
        }
    }
    let Some(parent_field) = parent_payload else {
        return Err(SlideTableLockStateError::InvalidSource);
    };
    if parent_count != 1 || parent_field.wire_type() != 2 {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    strict_reference(parent_field.payload(), limits, budget)
}

fn strict_reference(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LockBudget,
) -> Result<u64, SlideTableLockStateError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        budget.fields(1)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideTableLockStateError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideTableLockStateError::InvalidSource)?;
                if value == 0 || width != encoded_len(value) {
                    return Err(SlideTableLockStateError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideTableLockStateError::UnsupportedDependency),
            _ => return Err(SlideTableLockStateError::InvalidSource),
        }
    }
    identifier.ok_or(SlideTableLockStateError::InvalidSource)
}

fn reject_duplicates(
    values: &[u64],
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    budget.allocations(values.len())?;
    let mut seen = HashSet::new();
    seen.try_reserve(values.len())
        .map_err(|_| SlideTableLockStateError::Allocation {
            amount: values.len(),
        })?;
    for value in values {
        budget.work(1)?;
        if !seen.insert(*value) {
            return Err(SlideTableLockStateError::InvalidSource);
        }
    }
    Ok(())
}

fn ensure_unique_identity(
    package: &Package,
    identifier: u64,
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    let mut count = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            budget.work(1)?;
            if object.archive_info.identifier == Some(identifier) {
                count = count
                    .checked_add(1)
                    .ok_or(SlideTableLockStateError::InvalidSource)?;
            }
        }
    }
    if count != 1 {
        return Err(SlideTableLockStateError::UnsupportedDependency);
    }
    Ok(())
}

fn ensure_unique_table_owner(
    package: &Package,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    limits: WireLimits,
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    let mut owned_count = 0usize;
    let mut z_order_count = 0usize;
    let mut selected_slide = false;
    let mut model_owners = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for (message_index, message) in object.messages.iter().enumerate() {
                budget.work(1)?;
                if message.type_ == SLIDE_MESSAGE_TYPE {
                    let owned = repeated_references(
                        &message.data,
                        SLIDE_OWNED_DRAWABLES_FIELD,
                        limits,
                        budget,
                    )?;
                    let z_order =
                        repeated_references(&message.data, SLIDE_Z_ORDER_FIELD, limits, budget)?;
                    let owned_hits = owned
                        .iter()
                        .filter(|value| **value == table_info_identifier)
                        .count();
                    let z_hits = z_order
                        .iter()
                        .filter(|value| **value == table_info_identifier)
                        .count();
                    owned_count = owned_count.saturating_add(owned_hits);
                    z_order_count = z_order_count.saturating_add(z_hits);
                    if object.archive_info.identifier == Some(slide_identifier)
                        && owned_hits == 1
                        && z_hits == 1
                    {
                        selected_slide = true;
                    }
                }
                if message.type_ == TABLE_INFO_MESSAGE_TYPE {
                    let info = decode_table_info(&message.data, package, budget)?;
                    if info.table_model().identifier().get() == model_identifier {
                        model_owners = model_owners.saturating_add(1);
                    }
                }
                let _ = message_index;
            }
        }
    }
    if owned_count == 1 && z_order_count == 1 && selected_slide && model_owners == 1 {
        Ok(())
    } else {
        Err(SlideTableLockStateError::UnsupportedDependency)
    }
}

fn validate_global_inbound_references(
    package: &Package,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let (_, slide) = package
        .object_with_component(slide_identifier)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let slide_message_index = slide
        .messages
        .iter()
        .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let slide_info = slide
        .archive_info
        .message_infos
        .get(slide_message_index)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let (_, table_info) = package
        .object_with_component(table_info_identifier)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let table_info_message_index = table_info
        .messages
        .iter()
        .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let table_info_info = table_info
        .archive_info
        .message_infos
        .get(table_info_message_index)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    let expected_info_edges = reference_occurrence_count(slide_info, table_info_identifier)?;
    let expected_model_edges = reference_occurrence_count(table_info_info, model_identifier)?;
    if expected_info_edges == 0 || expected_model_edges == 0 {
        return Err(SlideTableLockStateError::UnsupportedDependency);
    }
    budget.work(
        slide_info
            .field_infos
            .len()
            .checked_add(table_info_info.field_infos.len())
            .ok_or(SlideTableLockStateError::InvalidSource)?,
    )?;
    let mut census = InboundReferenceCensus {
        slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
        slide_field_count: slide_info.field_infos.len(),
        table_info_field_count: table_info_info.field_infos.len(),
        slide_route_fields: route_field_indices(
            slide_info,
            [&[SLIDE_OWNED_DRAWABLES_FIELD], &[SLIDE_Z_ORDER_FIELD]],
        ),
        table_info_route_fields: route_field_indices(
            table_info_info,
            [
                &[TABLE_MODEL_FIELD],
                &[TABLE_SUPER_FIELD, DRAWABLE_PARENT_FIELD],
            ],
        ),
        expected_info_edges,
        expected_model_edges,
        info_edges: 0,
        model_edges: 0,
        invalid: false,
    };
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.archive_info.identifier.is_none()
                || object.messages.len() != object.archive_info.message_infos.len()
            {
                return Err(SlideTableLockStateError::InvalidSource);
            }
            let field_count = object
                .archive_info
                .message_infos
                .iter()
                .map(|info| info.field_infos.len())
                .try_fold(0usize, |count, fields| {
                    count
                        .checked_add(fields)
                        .ok_or(SlideTableLockStateError::InvalidSource)
                })?;
            let reference_count =
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .try_fold(0usize, |count, info| {
                        let field_references =
                            info.field_infos.iter().try_fold(0usize, |count, field| {
                                count
                                    .checked_add(field.object_references.len())
                                    .and_then(|count| {
                                        count.checked_add(field.data_references.len())
                                    })
                                    .ok_or(SlideTableLockStateError::InvalidSource)
                            })?;
                        let references = info
                            .object_references
                            .len()
                            .checked_add(info.data_references.len())
                            .and_then(|count| count.checked_add(field_references))
                            .ok_or(SlideTableLockStateError::InvalidSource)?;
                        count
                            .checked_add(references)
                            .ok_or(SlideTableLockStateError::InvalidSource)
                    })?;
            let message_bytes = object
                .messages
                .iter()
                .map(|message| message.data.len())
                .try_fold(0usize, |count, bytes| {
                    count
                        .checked_add(bytes)
                        .ok_or(SlideTableLockStateError::InvalidSource)
                })?;
            let header_bytes = usize::try_from(object.header_length).unwrap_or(usize::MAX);
            // `inspect_references_with_policy_and_limits` canonicalizes and
            // walks the complete ArchiveInfo header. Debit that work before
            // entering the visitor so a hostile header cannot spend the
            // transaction budget after the scan has already run.
            budget.fields(field_count)?;
            budget.references(reference_count)?;
            budget.allocations(
                object
                    .messages
                    .len()
                    .checked_add(field_count)
                    .and_then(|value| value.checked_add(reference_count))
                    .and_then(|value| value.checked_add(1))
                    .ok_or(SlideTableLockStateError::InvalidSource)?,
            )?;
            budget.retained(header_bytes)?;
            budget.scratch(header_bytes)?;
            budget.work(
                object
                    .messages
                    .len()
                    .checked_add(message_bytes)
                    .and_then(|value| value.checked_add(field_count))
                    .and_then(|value| value.checked_add(header_bytes))
                    .ok_or(SlideTableLockStateError::InvalidSource)?,
            )?;
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
        || census.info_edges != census.expected_info_edges
        || census.model_edges != census.expected_model_edges
    {
        return Err(SlideTableLockStateError::UnsupportedDependency);
    }
    Ok(())
}

fn reference_occurrence_count(
    info: &litchi_iwa_core::MessageInfo,
    identifier: u64,
) -> Result<usize, SlideTableLockStateError> {
    let direct = info
        .object_references
        .iter()
        .filter(|value| **value == identifier)
        .count();
    info.field_infos.iter().try_fold(direct, |count, field| {
        let nested = field
            .object_references
            .iter()
            .filter(|value| **value == identifier)
            .count();
        count
            .checked_add(nested)
            .ok_or(SlideTableLockStateError::InvalidSource)
    })
}

struct InboundReferenceCensus {
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    slide_field_count: usize,
    table_info_field_count: usize,
    slide_route_fields: [Option<usize>; 2],
    table_info_route_fields: [Option<usize>; 2],
    expected_info_edges: usize,
    expected_model_edges: usize,
    info_edges: usize,
    model_edges: usize,
    invalid: bool,
}

fn route_field_indices(
    info: &litchi_iwa_core::MessageInfo,
    paths: [&[u32]; 2],
) -> [Option<usize>; 2] {
    [
        info.field_infos
            .iter()
            .position(|field| field.path.as_slice() == paths[0]),
        info.field_infos
            .iter()
            .position(|field| field.path.as_slice() == paths[1]),
    ]
}

impl ArchiveReferenceVisitor for InboundReferenceCensus {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.referenced_identifier == self.table_info_identifier {
            let scope_is_allowed = match occurrence.scope {
                ArchiveReferenceScope::Message => true,
                ArchiveReferenceScope::Field { field_index } => {
                    field_index < self.slide_field_count
                        && self.slide_route_fields.contains(&Some(field_index))
                },
            };
            if occurrence.kind != ArchiveReferenceKind::Object
                || occurrence.object_identifier != self.slide_identifier
                || occurrence.message_index != self.slide_message_index
                || !scope_is_allowed
            {
                self.invalid = true;
            } else {
                self.info_edges = self.info_edges.saturating_add(1);
            }
        }
        if occurrence.referenced_identifier == self.model_identifier {
            let scope_is_allowed = match occurrence.scope {
                ArchiveReferenceScope::Message => true,
                ArchiveReferenceScope::Field { field_index } => {
                    field_index < self.table_info_field_count
                        && self.table_info_route_fields.contains(&Some(field_index))
                },
            };
            if occurrence.kind != ArchiveReferenceKind::Object
                || occurrence.object_identifier != self.table_info_identifier
                || occurrence.message_index != self.table_info_message_index
                || !scope_is_allowed
            {
                self.invalid = true;
            } else {
                self.model_edges = self.model_edges.saturating_add(1);
            }
        }
        let _scope_is_known = matches!(
            occurrence.scope,
            ArchiveReferenceScope::Message | ArchiveReferenceScope::Field { .. }
        );
        Ok(())
    }
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &LockSelection,
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    if source_catalog.package().len() != candidate_catalog.package().len() {
        return Err(SlideTableLockStateError::Verification);
    }
    budget.entries(
        source_catalog
            .package()
            .len()
            .checked_add(candidate_catalog.package().len())
            .ok_or(SlideTableLockStateError::InvalidSource)?,
    )?;
    for (entry, other) in source_catalog
        .package()
        .iter()
        .zip(candidate_catalog.package().iter())
    {
        let entry_bytes = entry
            .data()
            .len()
            .checked_add(other.data().len())
            .ok_or(SlideTableLockStateError::InvalidSource)?;
        budget.entry_bytes(entry_bytes)?;
        budget.total_bytes(entry_bytes)?;
        budget.work(
            entry
                .data()
                .len()
                .checked_add(entry.name().len())
                .and_then(|value| value.checked_add(other.data().len()))
                .ok_or(SlideTableLockStateError::InvalidSource)?,
        )?;
        if other.name() != entry.name() {
            return Err(SlideTableLockStateError::Verification);
        }
        if entry.name() != selection.component_name.as_ref()
            && (entry.data() != other.data() || entry.metadata() != other.metadata())
        {
            return Err(SlideTableLockStateError::Verification);
        }
    }
    let source_archive = component_archive(source, selection.component_name.as_ref(), budget)?;
    let candidate_archive =
        component_archive(candidate, selection.component_name.as_ref(), budget)?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SlideTableLockStateError::Verification);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    for source_object in &source_archive.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(SlideTableLockStateError::Verification)?;
        let candidate_object = candidate_archive
            .object(identifier)
            .ok_or(SlideTableLockStateError::Verification)?;
        if identifier == selection.table_info_identifier {
            let candidate_message = candidate_object
                .messages
                .get(selection.table_info_message_index)
                .ok_or(SlideTableLockStateError::Verification)?;
            let mut expected = source_object.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    selection.table_info_message_index,
                    candidate_message.clone(),
                    archive_limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = candidate_object.header_length;
            expected.data_length = candidate_object.data_length;
            if !expected.same_content_ignoring_offsets(candidate_object) {
                return Err(SlideTableLockStateError::Verification);
            }
        } else if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(SlideTableLockStateError::Verification);
        }
    }
    Ok(())
}

fn component_archive(
    package: &Package,
    name: &str,
    budget: &mut LockBudget,
) -> Result<Archive, SlideTableLockStateError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(SlideTableLockStateError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableLockStateError::InvalidSource);
    }
    budget.physical(entry.data().len())?;
    let snappy = package
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream =
        SnappyStream::decompress_with_limits(entry.data(), snappy).map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    budget.scratch(stream.as_bytes().len())?;
    let archive = Archive::parse_with_limits(stream.as_bytes(), limits).map_err(map_core_error)?;
    charge_archive_inventory(&archive, budget)?;
    Ok(archive)
}

fn charge_archive_inventory(
    archive: &Archive,
    budget: &mut LockBudget,
) -> Result<(), SlideTableLockStateError> {
    let objects = archive.objects.len();
    let messages = archive
        .objects
        .iter()
        .map(|object| object.messages.len())
        .try_fold(0usize, |count, value| {
            count
                .checked_add(value)
                .ok_or(SlideTableLockStateError::InvalidSource)
        })?;
    let fields = archive
        .objects
        .iter()
        .flat_map(|object| object.archive_info.message_infos.iter())
        .map(|info| info.field_infos.len())
        .try_fold(0usize, |count, value| {
            count
                .checked_add(value)
                .ok_or(SlideTableLockStateError::InvalidSource)
        })?;
    let references = archive
        .objects
        .iter()
        .flat_map(|object| object.archive_info.message_infos.iter())
        .try_fold(0usize, |count, info| {
            let nested = info.field_infos.iter().try_fold(0usize, |count, field| {
                count
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
                    .ok_or(SlideTableLockStateError::InvalidSource)
            })?;
            count
                .checked_add(info.object_references.len())
                .and_then(|value| value.checked_add(info.data_references.len()))
                .and_then(|value| value.checked_add(nested))
                .ok_or(SlideTableLockStateError::InvalidSource)
        })?;
    let message_bytes = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .map(|message| message.data.len())
        .try_fold(0usize, |count, value| {
            count
                .checked_add(value)
                .ok_or(SlideTableLockStateError::InvalidSource)
        })?;
    budget.payload_objects(objects)?;
    budget.payload_messages(messages)?;
    budget.fields(fields)?;
    budget.references(references)?;
    budget.allocations(
        objects
            .checked_add(messages)
            .and_then(|value| value.checked_add(fields))
            .and_then(|value| value.checked_add(references))
            .ok_or(SlideTableLockStateError::InvalidSource)?,
    )?;
    budget.work(
        objects
            .checked_add(messages)
            .and_then(|value| value.checked_add(fields))
            .and_then(|value| value.checked_add(references))
            .and_then(|value| value.checked_add(message_bytes))
            .ok_or(SlideTableLockStateError::InvalidSource)?,
    )
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideTableLockStateError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideTableLockStateError::EmptySlideName);
            }
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(|_| SlideTableLockStateError::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideTableLockStateError::SlideNameNotFound)
        },
    }
}

fn same_selection(left: &LockSelection, right: &LockSelection) -> bool {
    left.slide_position == right.slide_position
        && left.table_position == right.table_position
        && left.slide_identifier == right.slide_identifier
        && left.table_info_identifier == right.table_info_identifier
        && left.model_identifier == right.model_identifier
        && left.slide_message_index == right.slide_message_index
        && left.table_info_message_index == right.table_info_message_index
        && left.model_message_index == right.model_message_index
        && left.component_name == right.component_name
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideTableLockStateError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTableLockStateError::UnsupportedSource),
    }
}

fn map_read_error(error: ReadError) -> SlideTableLockStateError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableLockStateError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableLockStateLimitKind::References,
                _ => SlideTableLockStateLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableLockStateError::Allocation { amount },
        _ => SlideTableLockStateError::InvalidSource,
    }
}

fn map_table_info_error(error: table_info_codec::DecodeError) -> SlideTableLockStateError {
    if let Some(amount) = error.allocation_amount() {
        return SlideTableLockStateError::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideTableLockStateError::LimitExceeded {
            kind: SlideTableLockStateLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideTableLockStateError::LimitExceeded {
            kind: SlideTableLockStateLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.output_limit_values() {
        return SlideTableLockStateError::LimitExceeded {
            kind: SlideTableLockStateLimitKind::OutputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.allocation_limit_values() {
        return SlideTableLockStateError::LimitExceeded {
            kind: SlideTableLockStateLimitKind::Allocations,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.retained_limit_values() {
        return SlideTableLockStateError::LimitExceeded {
            kind: SlideTableLockStateLimitKind::Retained,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.scratch_limit_values() {
        return SlideTableLockStateError::LimitExceeded {
            kind: SlideTableLockStateLimitKind::Scratch,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            table_info_codec::WireResourceLimit::Bytes { observed, maximum } => {
                SlideTableLockStateError::LimitExceeded {
                    kind: SlideTableLockStateLimitKind::InputBytes,
                    observed: observed.unwrap_or_default() as u64,
                    maximum: maximum.unwrap_or_default() as u64,
                }
            },
            table_info_codec::WireResourceLimit::Nesting { observed, maximum } => {
                SlideTableLockStateError::LimitExceeded {
                    kind: SlideTableLockStateLimitKind::WireNesting,
                    observed: observed.unwrap_or_default() as u64,
                    maximum: maximum.unwrap_or_default() as u64,
                }
            },
            _ => SlideTableLockStateError::InvalidSource,
        };
    }
    SlideTableLockStateError::InvalidSource
}

fn map_metadata_error(error: package_metadata_codec::RewriteError) -> SlideTableLockStateError {
    if let Some(amount) = error.allocation_request() {
        return SlideTableLockStateError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return SlideTableLockStateError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
            (SlideTableLockStateLimitKind::InputBytes, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
            (SlideTableLockStateLimitKind::OutputBytes, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
            (SlideTableLockStateLimitKind::WireFields, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
            (SlideTableLockStateLimitKind::WireWork, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
            return SlideTableLockStateError::LimitExceeded {
                kind: SlideTableLockStateLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        package_metadata_codec::RewriteLimit::Components { observed, maximum } => {
            (SlideTableLockStateLimitKind::Components, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::References { observed, maximum } => {
            (SlideTableLockStateLimitKind::References, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::Additions { observed, maximum } => {
            (SlideTableLockStateLimitKind::Allocations, observed, maximum)
        },
        _ => return SlideTableLockStateError::InvalidSource,
    };
    SlideTableLockStateError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_table_model_error(
    error: table_model_discovery_codec::DecodeError,
) -> SlideTableLockStateError {
    let Some(limit) = error.resource_limit() else {
        return SlideTableLockStateError::InvalidSource;
    };
    match limit {
        table_model_discovery_codec::DecodeLimit::Bytes { observed, maximum } => {
            SlideTableLockStateError::LimitExceeded {
                kind: SlideTableLockStateLimitKind::InputBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        table_model_discovery_codec::DecodeLimit::Fields { observed, maximum } => {
            SlideTableLockStateError::LimitExceeded {
                kind: SlideTableLockStateLimitKind::WireFields,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        table_model_discovery_codec::DecodeLimit::Work { observed, maximum } => {
            SlideTableLockStateError::LimitExceeded {
                kind: SlideTableLockStateLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        table_model_discovery_codec::DecodeLimit::Text { observed, maximum } => {
            SlideTableLockStateError::LimitExceeded {
                kind: SlideTableLockStateLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        table_model_discovery_codec::DecodeLimit::Nesting { observed, maximum } => {
            SlideTableLockStateError::LimitExceeded {
                kind: SlideTableLockStateLimitKind::WireNesting,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        _ => SlideTableLockStateError::InvalidSource,
    }
}

fn map_wire_error(_error: litchi_iwa_common::Error) -> SlideTableLockStateError {
    SlideTableLockStateError::InvalidSource
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTableLockStateError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableLockStateError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    SlideTableLockStateLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideTableLockStateLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideTableLockStateLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideTableLockStateLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    SlideTableLockStateLimitKind::TotalBytes
                },
                _ => SlideTableLockStateLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableLockStateError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTableLockStateError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideTableLockStateError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableLockStateError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => SlideTableLockStateLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableLockStateLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideTableLockStateLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    SlideTableLockStateLimitKind::EntryBytes
                },
                _ => SlideTableLockStateLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableLockStateError::Allocation { amount: requested }
        },
        _ => SlideTableLockStateError::InvalidSource,
    }
}
