//! Selector-first, source-preserving row and column sizing for Keynote tables.
//!
//! The public surface contains only checked selectors and the archive-free
//! [`Dimension`], [`Points`], and [`Size`] values.  Native table-model,
//! header-bucket, geometry, and ZIP records stay private to this module.  A
//! changed operation rewrites the selected header bucket and the owning
//! `TSD.DrawableArchive` geometry in one package transaction.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::too_many_arguments,
    clippy::wildcard_enum_match_arm,
    reason = "The focused package boundary keeps native proof details private."
)]

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{
        NestedFieldEdit, NestedFieldReplacement, WireView, patch_nested_fields_batched_with_limits,
    },
};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceVisitor, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{package_metadata_codec, table_dimension_codec as codec, table_info_codec};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::{
    TableSelector,
    dimension::{Dimension, Points, Size},
};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const DRAWABLE_GEOMETRY_FIELD: u32 = 1;
const GEOMETRY_SIZE_FIELD: u32 = 2;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;
const MODEL_STORAGE_FIELD: u32 = 4;
const MODEL_ROW_BUCKET_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 1, 2];
const MODEL_COLUMN_BUCKET_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 2];
const DEFAULT_HEADER_BUCKET_ROWS: u32 = 65_536;
const MAX_ROLE_MESSAGE_TYPES: [u32; 7] = [
    TABLE_INFO_MESSAGE_TYPE,
    TABLE_MODEL_MESSAGE_TYPE,
    HEADER_BUCKET_MESSAGE_TYPE,
    6_003,
    6_008,
    6_247,
    401,
];

/// Finite resources governed by one slide-table dimension operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableDimensionLimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    PayloadObjects,
    PayloadMessages,
    PayloadItems,
    References,
    WireBytes,
    WireOutputBytes,
    WireFields,
    WireNesting,
    WireWork,
    Allocations,
    Retained,
    Scratch,
    Components,
    TransactionWork,
}

impl fmt::Display for SlideTableDimensionLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::References => "references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
            Self::TransactionWork => "transaction work",
        })
    }
}

/// Content-free semantic path for a slide-table dimension operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableDimensionPath {
    /// The complete package.
    Package,
    /// One checked slide/table and dimension.
    Table {
        slide: Position,
        table: Position,
        dimension: Dimension,
    },
}

impl fmt::Display for SlideTableDimensionPath {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => f.write_str("package"),
            Self::Table {
                slide,
                table,
                dimension,
            } => write!(
                f,
                "slide {} table {} {} {}",
                slide.get(),
                table.get(),
                dimension.noun(),
                dimension.index()
            ),
        }
    }
}

/// Failure from a Keynote slide-table dimension read or transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTableDimensionError {
    #[error("this Keynote source does not support physical slide-table dimension edits")]
    UnsupportedSource,
    #[error("the requested Keynote slide-table dimension graph has an unsupported dependency")]
    UnsupportedDependency,
    #[error("the requested Keynote slide-table dimension topology is unsupported")]
    UnsupportedTopology,
    #[error("the Keynote slide-table dimension selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    #[error("the selected Keynote slide table is locked")]
    Locked,
    #[error("the selected Keynote slide-table dimension source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table dimensions {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTableDimensionLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote slide-table dimension transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote slide-table dimension failed semantic verification")]
    Verification,
    #[error("the Keynote slide-table dimension patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy)]
struct DimensionBudget {
    max_input: usize,
    max_output: usize,
    max_entries: usize,
    max_entry_bytes: usize,
    max_total_bytes: usize,
    max_objects: usize,
    max_messages: usize,
    max_items: usize,
    max_references: usize,
    max_wire_bytes: usize,
    max_wire_output: usize,
    max_fields: usize,
    max_nesting: usize,
    max_work: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    max_components: usize,
    max_transaction_work: usize,
    input: usize,
    output: usize,
    entries: usize,
    entry_bytes: usize,
    total_bytes: usize,
    objects: usize,
    messages: usize,
    items: usize,
    references: usize,
    wire_bytes: usize,
    wire_output: usize,
    fields: usize,
    nesting: usize,
    work: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
    components: usize,
    transaction_work: usize,
}

impl DimensionBudget {
    fn new(package: &Package) -> Result<Self, SlideTableDimensionError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let archive = package.state.options.archive();
        let source = usize::try_from(archive.max_input_bytes())
            .map_err(|_| SlideTableDimensionError::InvalidSource)?
            .max(1);
        let total = usize::try_from(archive.max_total_bytes())
            .unwrap_or(usize::MAX)
            .max(source);
        let aggregate = source
            .checked_mul(8)
            .ok_or(SlideTableDimensionError::InvalidSource)?;
        let semantic = package.semantic_limits();
        let components = package.state.source.components().len().max(1);
        Ok(Self {
            max_input: aggregate,
            max_output: aggregate.min(total.max(1)),
            max_entries: aggregate,
            max_entry_bytes: total,
            max_total_bytes: total,
            max_objects: semantic.max_objects().saturating_mul(8).max(1),
            max_messages: semantic.max_objects().saturating_mul(16).max(1),
            max_items: semantic.max_slides().saturating_mul(8).max(1),
            max_references: semantic.max_references().saturating_mul(8).max(1),
            max_wire_bytes: aggregate,
            max_wire_output: aggregate,
            max_fields: wire
                .max_fields()
                .saturating_mul(8)
                .min(WireLimits::MAX_FIELDS),
            max_nesting: wire.max_nesting(),
            max_work: wire.max_rewrite_work(),
            max_allocations: aggregate,
            max_retained: total.min(aggregate.max(1)),
            max_scratch: aggregate,
            // A transaction may select the source more than once (edit
            // revalidation, candidate reopen, and inverse/apply checks), and
            // each selection performs both metadata and physical authority
            // passes.  Keep the component axis aggregate across those bounded
            // passes rather than rejecting a valid package after the eighth
            // scan of each physical component.
            // Selection performs bounded component scans for ownership,
            // bucket routing, identity checks, rewrite, candidate reopen,
            // and locality.  Keep the logical component envelope aggregate
            // across those passes; physical archive limits remain the hard
            // per-package bound.
            max_components: components.saturating_mul(32).max(1),
            max_transaction_work: aggregate,
            input: 0,
            output: 0,
            entries: 0,
            entry_bytes: 0,
            total_bytes: 0,
            objects: 0,
            messages: 0,
            items: 0,
            references: 0,
            wire_bytes: 0,
            wire_output: 0,
            fields: 0,
            nesting: 0,
            work: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
            components: 0,
            transaction_work: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideTableDimensionLimitKind,
    ) -> Result<(), SlideTableDimensionError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideTableDimensionError::InvalidSource)?;
        if observed > maximum {
            return Err(SlideTableDimensionError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn input(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            SlideTableDimensionLimitKind::InputBytes,
        )
    }
    fn output(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            SlideTableDimensionLimitKind::OutputBytes,
        )
    }
    fn entries(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.entries,
            amount,
            self.max_entries,
            SlideTableDimensionLimitKind::Entries,
        )
    }
    fn entry_bytes(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.entry_bytes,
            amount,
            self.max_entry_bytes,
            SlideTableDimensionLimitKind::EntryBytes,
        )
    }
    fn total_bytes(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.total_bytes,
            amount,
            self.max_total_bytes,
            SlideTableDimensionLimitKind::TotalBytes,
        )
    }
    fn objects(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.objects,
            amount,
            self.max_objects,
            SlideTableDimensionLimitKind::PayloadObjects,
        )
    }
    fn messages(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.messages,
            amount,
            self.max_messages,
            SlideTableDimensionLimitKind::PayloadMessages,
        )
    }
    fn items(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.items,
            amount,
            self.max_items,
            SlideTableDimensionLimitKind::PayloadItems,
        )
    }
    fn references(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideTableDimensionLimitKind::References,
        )
    }
    fn wire_bytes(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.wire_bytes,
            amount,
            self.max_wire_bytes,
            SlideTableDimensionLimitKind::WireBytes,
        )
    }
    fn wire_output(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.wire_output,
            amount,
            self.max_wire_output,
            SlideTableDimensionLimitKind::WireOutputBytes,
        )
    }
    fn fields(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            SlideTableDimensionLimitKind::WireFields,
        )
    }
    fn work(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            SlideTableDimensionLimitKind::WireWork,
        )
    }
    fn allocations(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideTableDimensionLimitKind::Allocations,
        )
    }
    fn retained(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideTableDimensionLimitKind::Retained,
        )
    }
    fn scratch(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            SlideTableDimensionLimitKind::Scratch,
        )
    }
    fn components(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.components,
            amount,
            self.max_components,
            SlideTableDimensionLimitKind::Components,
        )
    }
    fn transaction_work(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        Self::add(
            &mut self.transaction_work,
            amount,
            self.max_transaction_work,
            SlideTableDimensionLimitKind::TransactionWork,
        )
    }

    fn physical(&mut self, amount: usize) -> Result<(), SlideTableDimensionError> {
        self.input(amount)?;
        self.entry_bytes(amount)?;
        self.total_bytes(amount)?;
        self.work(amount)
    }

    fn residual_wire(&self, package: &Package) -> Result<WireLimits, SlideTableDimensionError> {
        let base = package.wire_limits().map_err(map_wire_error)?;
        let input = self.remaining_wire(
            self.wire_bytes,
            self.max_wire_bytes,
            SlideTableDimensionLimitKind::WireBytes,
        )?;
        let output = self.remaining_wire(
            self.wire_output,
            self.max_wire_output,
            SlideTableDimensionLimitKind::WireOutputBytes,
        )?;
        let fields = self.remaining_wire(
            self.fields,
            self.max_fields,
            SlideTableDimensionLimitKind::WireFields,
        )?;
        let work = self.remaining_wire(
            self.work,
            self.max_work,
            SlideTableDimensionLimitKind::WireWork,
        )?;
        if self.max_nesting == 0 {
            return Err(SlideTableDimensionError::LimitExceeded {
                kind: SlideTableDimensionLimitKind::WireNesting,
                observed: 1,
                maximum: 0,
            });
        }
        let nesting = self.max_nesting;
        base.with_input_bytes(base.max_input_bytes().min(input))
            .and_then(|v| v.with_output_bytes(base.max_output_bytes().min(output)))
            .and_then(|v| v.with_fields(base.max_fields().min(fields)))
            .and_then(|v| v.with_rewrite_work(base.max_rewrite_work().min(work)))
            .and_then(|v| v.with_nesting(base.max_nesting().min(nesting)))
            .map_err(map_wire_error)
    }

    fn remaining_wire(
        &self,
        used: usize,
        maximum: usize,
        kind: SlideTableDimensionLimitKind,
    ) -> Result<usize, SlideTableDimensionError> {
        maximum
            .checked_sub(used)
            .filter(|remaining| *remaining > 0)
            .ok_or(SlideTableDimensionError::LimitExceeded {
                kind,
                observed: used.saturating_add(1) as u64,
                maximum: maximum as u64,
            })
    }

    fn codec_options(
        &self,
        package: &Package,
        _bytes: usize,
    ) -> Result<codec::DecodeOptions, SlideTableDimensionError> {
        let limits = self.residual_wire(package)?;
        let references = self.remaining_wire(
            self.references,
            self.max_references,
            SlideTableDimensionLimitKind::References,
        )?;
        Ok(codec::DecodeOptions::new(
            limits.max_input_bytes(),
            limits.max_fields(),
            limits.max_rewrite_work(),
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            references,
            limits.max_input_bytes(),
        ))
    }

    fn codec_report(
        &mut self,
        report: codec::DecodeReport,
    ) -> Result<(), SlideTableDimensionError> {
        self.input(report.source_bytes())?;
        self.wire_bytes(report.source_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.references(report.references())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableDimensionError::LimitExceeded {
                kind: SlideTableDimensionLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideTableDimensionError> {
        self.output(requirements.output_bytes())?;
        self.wire_output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.transaction_work(requirements.output_bytes())
    }
}

/// One selected table and its private native proof.
#[derive(Clone, PartialEq)]
struct DimensionSelection {
    slide_position: Position,
    table_position: Position,
    slide_identifier: u64,
    table_info_identifier: u64,
    table_info_component: Arc<str>,
    table_info_object: usize,
    table_info_message: usize,
    model_identifier: u64,
    model_component: Arc<str>,
    model_object: usize,
    model_message: usize,
    bucket_identifier: u64,
    bucket_component: Arc<str>,
    bucket_object: usize,
    bucket_message: usize,
    rows: u32,
    columns: u32,
    default_row_height: f64,
    default_column_width: f64,
    row_total: f32,
    column_total: f32,
    geometry_width: f32,
    geometry_height: f32,
    locked: bool,
    before: Size,
    dimension: Dimension,
}

impl fmt::Debug for DimensionSelection {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("DimensionSelection")
            .field("slide_position", &self.slide_position)
            .field("table_position", &self.table_position)
            .field("dimension", &self.dimension)
            .field("before", &self.before)
            .field("locked", &self.locked)
            .finish_non_exhaustive()
    }
}

impl DimensionSelection {
    const fn path(&self) -> SlideTableDimensionPath {
        SlideTableDimensionPath::Table {
            slide: self.slide_position,
            table: self.table_position,
            dimension: self.dimension,
        }
    }
}

/// One mutable dimension value staged against an immutable package snapshot.
pub struct SlideTableDimensionEdit<'a> {
    source: &'a Package,
    selection: DimensionSelection,
    after: Size,
}

impl fmt::Debug for SlideTableDimensionEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideTableDimensionEdit")
            .field("path", &self.selection.path())
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableDimensionEdit<'_> {
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    #[must_use]
    pub const fn table_position(&self) -> Position {
        self.selection.table_position
    }

    #[must_use]
    pub const fn dimension(&self) -> Dimension {
        self.selection.dimension
    }

    #[must_use]
    pub const fn path(&self) -> SlideTableDimensionPath {
        self.selection.path()
    }

    #[must_use]
    pub const fn before(&self) -> Size {
        self.selection.before
    }

    #[must_use]
    pub const fn size(&self) -> Size {
        self.after
    }

    #[must_use]
    pub const fn after(&self) -> Size {
        self.after
    }

    #[must_use]
    pub fn set(mut self, size: Size) -> Self {
        self.after = size;
        self
    }

    #[must_use]
    pub fn set_points(self, points: Points) -> Self {
        self.set(Size::Points(points))
    }

    #[must_use]
    pub fn reset(self) -> Self {
        self.set(Size::Default)
    }

    pub fn commit(self) -> Result<SlideTableDimensionCommit, SlideTableDimensionError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source reversible slide-table dimension patch.
#[derive(Clone, PartialEq)]
pub struct SlideTableDimensionPatch {
    artifacts: ExactArtifacts,
    selection: DimensionSelection,
    before: Size,
    after: Size,
    touched_components: usize,
    deleted_previews: usize,
    source_previews_absent: bool,
    target_previews_absent: bool,
}

impl fmt::Debug for SlideTableDimensionPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideTableDimensionPatch")
            .field("path", &self.selection.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableDimensionPatch {
    #[must_use]
    pub const fn path(&self) -> SlideTableDimensionPath {
        self.selection.path()
    }

    #[must_use]
    pub const fn dimension(&self) -> Dimension {
        self.selection.dimension
    }

    #[must_use]
    pub const fn before(&self) -> Size {
        self.before
    }

    #[must_use]
    pub const fn after(&self) -> Size {
        self.after
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
            touched_components: self.touched_components,
            deleted_previews: self.deleted_previews,
            source_previews_absent: self.target_previews_absent,
            target_previews_absent: self.source_previews_absent,
        }
    }
}

/// Compact publication diagnostics for one dimension transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableDimensionDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideTableDimensionDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
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

/// Fully verified result of one dimension transaction.
#[must_use = "a slide-table dimension commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableDimensionCommit {
    package: Package,
    patch: SlideTableDimensionPatch,
    diagnostics: SlideTableDimensionDiagnostics,
}

impl SlideTableDimensionCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &SlideTableDimensionPatch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableDimensionDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted slide table's explicit or default row/column size.
    pub fn slide_table_dimension_size<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        dimension: Dimension,
    ) -> Result<Size, SlideTableDimensionError> {
        let mut budget = DimensionBudget::new(self)?;
        Ok(select_table(self, slide.into(), table.into(), dimension, &mut budget)?.before)
    }

    /// Start a selector-first immutable row/column size edit.
    pub fn edit_slide_table_dimension_size<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        dimension: Dimension,
    ) -> Result<SlideTableDimensionEdit<'_>, SlideTableDimensionError> {
        let mut budget = DimensionBudget::new(self)?;
        let selection = select_table(self, slide.into(), table.into(), dimension, &mut budget)?;
        Ok(SlideTableDimensionEdit {
            source: self,
            after: selection.before,
            selection,
        })
    }

    /// Apply an exact-source checked reversible dimension patch.
    pub fn apply_slide_table_dimension_size(
        &self,
        patch: &SlideTableDimensionPatch,
    ) -> Result<SlideTableDimensionCommit, SlideTableDimensionError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideTableDimensionError::PatchConflict);
        }
        let mut budget = DimensionBudget::new(self)?;
        let current = select_table(
            self,
            SlideSelector::position(patch.selection.slide_position),
            TableSelector::position(patch.selection.table_position),
            patch.selection.dimension,
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SlideTableDimensionError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideTableDimensionCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableDimensionDiagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch, &mut budget)
    }
}

fn commit_edit(
    source: &Package,
    selection: &DimensionSelection,
    after: Size,
) -> Result<SlideTableDimensionCommit, SlideTableDimensionError> {
    let catalog = physical_catalog(source)?;
    let mut budget = DimensionBudget::new(source)?;
    let current = select_table(
        source,
        SlideSelector::position(selection.slide_position),
        TableSelector::position(selection.table_position),
        selection.dimension,
        &mut budget,
    )?;
    if !same_selection(&current, selection) || current.before != selection.before {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    let source_previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| SlideTableDimensionError::InvalidSource)?;
    if after == selection.before {
        let bytes = catalog.shared_source();
        return Ok(SlideTableDimensionCommit {
            package: source.snapshot(),
            patch: SlideTableDimensionPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before,
                after,
                touched_components: 0,
                deleted_previews: 0,
                source_previews_absent: source_previews.len() == 0,
                target_previews_absent: source_previews.len() == 0,
            },
            diagnostics: SlideTableDimensionDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SlideTableDimensionError::UnsupportedSource);
    }
    if selection.locked {
        return Err(SlideTableDimensionError::Locked);
    }
    let candidate = rewrite_dimension(
        source,
        &current,
        after,
        source_previews.names(),
        &mut budget,
    )?;
    candidate.validate().map_err(map_read_error)?;
    let candidate_selection = select_table(
        &candidate,
        SlideSelector::position(selection.slide_position),
        TableSelector::position(selection.table_position),
        selection.dimension,
        &mut budget,
    )?;
    if !same_selection(&candidate_selection, selection) || candidate_selection.before != after {
        return Err(SlideTableDimensionError::Verification);
    }
    let target_catalog = physical_catalog(&candidate)?;
    let target_previews =
        super::rendering_invalidation::root_preview_deletions(target_catalog.package())
            .map_err(|_| SlideTableDimensionError::Verification)?;
    verify_locality(
        source,
        &candidate,
        selection,
        &current,
        after,
        source_previews.names(),
        target_previews.names(),
        &mut budget,
    )?;
    budget.allocations(1)?;
    budget.retained(source.source_bytes().len())?;
    let target = target_catalog.shared_source();
    let touched = usize::from(selection.table_info_component != selection.bucket_component) + 1;
    Ok(SlideTableDimensionCommit {
        package: candidate,
        patch: SlideTableDimensionPatch {
            artifacts: ExactArtifacts::new(Arc::from(source.source_bytes()), target),
            selection: selection.clone(),
            before: selection.before,
            after,
            touched_components: touched,
            deleted_previews: source_previews.len().saturating_sub(target_previews.len()),
            source_previews_absent: source_previews.len() == 0,
            target_previews_absent: target_previews.len() == 0,
        },
        diagnostics: SlideTableDimensionDiagnostics::published(
            touched,
            source_previews.len().saturating_sub(target_previews.len()),
        ),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &SlideTableDimensionPatch,
    budget: &mut DimensionBudget,
) -> Result<SlideTableDimensionCommit, SlideTableDimensionError> {
    budget.input(patch.artifacts.target().len())?;
    budget.transaction_work(patch.artifacts.target().len())?;
    budget.allocations(1)?;
    budget.retained(patch.artifacts.target().len())?;
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        TableSelector::position(patch.selection.table_position),
        patch.selection.dimension,
        budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SlideTableDimensionError::Verification);
    }
    let source_catalog = physical_catalog(source)?;
    let target_catalog = physical_catalog(&candidate)?;
    let source_previews =
        super::rendering_invalidation::root_preview_deletions(source_catalog.package())
            .map_err(|_| SlideTableDimensionError::Verification)?;
    let target_previews =
        super::rendering_invalidation::root_preview_deletions(target_catalog.package())
            .map_err(|_| SlideTableDimensionError::Verification)?;
    verify_locality(
        source,
        &candidate,
        &patch.selection,
        &patch.selection,
        patch.after,
        source_previews.names(),
        target_previews.names(),
        budget,
    )?;
    Ok(SlideTableDimensionCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableDimensionDiagnostics::published(
            patch.touched_components,
            patch.deleted_previews,
        ),
    })
}

fn select_table(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
    dimension: Dimension,
    budget: &mut DimensionBudget,
) -> Result<DimensionSelection, SlideTableDimensionError> {
    let catalog = physical_catalog(package)?;
    budget.entries(catalog.package().len())?;
    budget.input(package.source_bytes().len())?;
    budget.transaction_work(package.source_bytes().len())?;
    let slide_position = resolve_slide_position(package, slide_selector)?;
    budget.items(1)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTableDimensionError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (_slide_component, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    ensure_unique_identity(package, record.slide_identifier, budget)?;
    let (slide_message_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
    let owned = repeated_references(
        slide_payload,
        SLIDE_OWNED_DRAWABLES_FIELD,
        budget.residual_wire(package)?,
        budget,
    )?;
    let z_order = repeated_references(
        slide_payload,
        SLIDE_Z_ORDER_FIELD,
        budget.residual_wire(package)?,
        budget,
    )?;
    reject_duplicates(&owned, budget)?;
    reject_duplicates(&z_order, budget)?;
    validate_slide_metadata(slide, slide_message_index, &owned, &z_order, budget)?;
    budget.allocations(owned.len())?;
    let mut owned_set = HashSet::new();
    owned_set
        .try_reserve(owned.len())
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: owned.len(),
        })?;
    owned_set.extend(owned.iter().copied());

    let mut table_candidates = Vec::new();
    budget.allocations(z_order.len())?;
    table_candidates
        .try_reserve_exact(z_order.len())
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: z_order.len(),
        })?;
    for table_info_identifier in z_order.iter().copied() {
        let (table_info_component, table_info_object) = package
            .object_with_component(table_info_identifier)
            .ok_or(SlideTableDimensionError::InvalidSource)?;
        let mut role_count = 0usize;
        let mut role_type = None;
        for message in &table_info_object.messages {
            if is_role_type(message.type_) {
                role_count = role_count.saturating_add(1);
                role_type = Some(message.type_);
            }
        }
        if role_type != Some(TABLE_INFO_MESSAGE_TYPE) {
            if role_count != 0 {
                return Err(SlideTableDimensionError::UnsupportedTopology);
            }
            continue;
        }
        if role_count != 1 || !owned_set.contains(&table_info_identifier) {
            return Err(SlideTableDimensionError::InvalidSource);
        }
        let (table_info_message, info_payload) =
            unique_message(table_info_object, TABLE_INFO_MESSAGE_TYPE)?;
        let info = decode_table_info(info_payload, package, budget)?;
        let parent = table_parent(info_payload, budget.residual_wire(package)?, budget)?;
        if parent != record.slide_identifier {
            return Err(SlideTableDimensionError::InvalidSource);
        }
        let model_identifier = info.table_model().identifier().get();
        validate_table_info_metadata(
            table_info_object,
            table_info_message,
            model_identifier,
            budget,
        )?;
        let (model_component, model_object) = package
            .object_with_component(model_identifier)
            .ok_or(SlideTableDimensionError::InvalidSource)?;
        let mut model_role_count = 0usize;
        let mut model_role_type = None;
        for message in &model_object.messages {
            if is_role_type(message.type_) {
                model_role_count = model_role_count.saturating_add(1);
                model_role_type = Some(message.type_);
            }
        }
        if model_role_count != 1 || model_role_type != Some(TABLE_MODEL_MESSAGE_TYPE) {
            return Err(SlideTableDimensionError::UnsupportedTopology);
        }
        let (model_message, model_payload) =
            unique_message(model_object, TABLE_MODEL_MESSAGE_TYPE)?;
        let options = budget.codec_options(package, model_payload.len())?;
        let (model, report) = codec::decode_table_model_with_report(model_payload, options)
            .map_err(map_codec_error)?;
        budget.codec_report(report)?;
        if model.number_of_rows() == 0 || model.number_of_columns() == 0 {
            return Err(SlideTableDimensionError::UnsupportedTopology);
        }
        let (default_row_height, default_column_width) =
            parse_model_defaults(model_payload, budget.residual_wire(package)?, budget)?;
        let geometry = parse_table_geometry(info_payload, budget.residual_wire(package)?, budget)?;
        let locked = info.locked().unwrap_or(false);
        let table_info_object_index = object_index_for(package, table_info_identifier, budget)?;
        let model_object_index = object_index_for(package, model_identifier, budget)?;
        budget.allocations(2)?;
        table_candidates.push((
            table_info_identifier,
            Arc::<str>::from(table_info_component),
            table_info_object_index,
            table_info_message,
            model_identifier,
            Arc::<str>::from(model_component),
            model_object_index,
            model_message,
            model.number_of_rows(),
            model.number_of_columns(),
            default_row_height,
            default_column_width,
            geometry,
            locked,
        ));
    }

    let table_position = table_selector.as_position();
    let (
        table_info_identifier,
        table_info_component,
        table_info_object,
        table_info_message,
        model_identifier,
        model_component,
        model_object,
        model_message,
        rows,
        columns,
        default_row_height,
        default_column_width,
        geometry,
        locked,
    ) = table_candidates.get(table_position.get()).cloned().ok_or(
        SlideTableDimensionError::TablePositionNotFound {
            position: table_position,
        },
    )?;
    ensure_unique_identity(package, table_info_identifier, budget)?;
    ensure_unique_identity(package, model_identifier, budget)?;
    ensure_unique_table_owner(
        package,
        record.slide_identifier,
        table_info_identifier,
        model_identifier,
        budget,
    )?;

    let model_payload = object_message_payload(
        package,
        model_identifier,
        model_message,
        TABLE_MODEL_MESSAGE_TYPE,
    )?;
    let model_options = budget.codec_options(package, model_payload.len())?;
    let (model, report) = codec::decode_table_model_with_report(model_payload, model_options)
        .map_err(map_codec_error)?;
    budget.codec_report(report)?;
    if model.number_of_rows() != rows || model.number_of_columns() != columns {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    let store_options = budget.codec_options(package, model.base_data_store().len())?;
    let (store, store_report) =
        codec::decode_data_store_with_report(model.base_data_store(), store_options)
            .map_err(map_codec_error)?;
    budget.codec_report(store_report)?;

    let row_options = budget.codec_options(package, store.row_headers().len())?;
    let (_, row_report) =
        codec::decode_header_storage_with_report(store.row_headers(), row_options)
            .map_err(map_codec_error)?;
    budget.codec_report(row_report)?;
    let row_references =
        storage_references(store.row_headers(), budget.residual_wire(package)?, budget)?;
    let expected_rows = usize::try_from(rows.div_ceil(DEFAULT_HEADER_BUCKET_ROWS))
        .map_err(|_| SlideTableDimensionError::InvalidSource)?;
    if row_references.len() != expected_rows {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    let column_identifier = checked_reference(store.column_headers())?;
    if row_references.contains(&column_identifier) {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    budget.allocations(row_references.len())?;
    let mut row_locations = Vec::new();
    row_locations
        .try_reserve_exact(row_references.len())
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: row_references.len(),
        })?;
    reject_duplicates(&row_references, budget)?;
    for identifier in row_references {
        row_locations.push(locate_bucket(package, identifier, budget)?);
    }
    let column_location = locate_bucket(package, column_identifier, budget)?;
    validate_model_storage_references(
        package,
        model_identifier,
        model_message,
        &row_locations,
        &column_location,
        budget,
    )?;
    let row_dimension = match dimension {
        Dimension::Row(_) => dimension,
        Dimension::Column(_) => Dimension::Row(0),
    };
    let column_dimension = match dimension {
        Dimension::Row(_) => Dimension::Column(0),
        Dimension::Column(_) => dimension,
    };
    let (row_total, row_selected) = collect_axis(
        package,
        &row_locations,
        rows,
        row_dimension,
        default_row_height,
        budget,
    )?;
    let (column_total, column_selected) = collect_axis(
        package,
        std::slice::from_ref(&column_location),
        columns,
        column_dimension,
        default_column_width,
        budget,
    )?;
    let (bucket, before) = match dimension {
        Dimension::Row(index) => {
            let slot = index / usize::try_from(DEFAULT_HEADER_BUCKET_ROWS).unwrap_or(usize::MAX);
            let bucket = row_locations
                .get(slot)
                .cloned()
                .ok_or(SlideTableDimensionError::InvalidSource)?;
            (bucket, row_selected.unwrap_or(Size::Default))
        },
        Dimension::Column(_) => (
            column_location.clone(),
            column_selected.unwrap_or(Size::Default),
        ),
    };
    validate_geometry_matches(geometry, row_total, column_total)?;
    validate_global_references(
        package,
        record.slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message,
        model_identifier,
        model_message,
        &row_locations,
        &column_location,
        budget,
    )?;
    validate_package_metadata(
        package,
        record.slide_identifier,
        _slide_component,
        table_info_identifier,
        table_info_component.as_ref(),
        model_identifier,
        model_component.as_ref(),
        &row_locations,
        &column_location,
        budget,
    )?;
    Ok(DimensionSelection {
        slide_position,
        table_position,
        slide_identifier: record.slide_identifier,
        table_info_identifier,
        table_info_component,
        table_info_object,
        table_info_message,
        model_identifier,
        model_component,
        model_object,
        model_message,
        bucket_identifier: bucket.identifier,
        bucket_component: bucket.component,
        bucket_object: bucket.object,
        bucket_message: bucket.message,
        rows,
        columns,
        default_row_height,
        default_column_width,
        row_total,
        column_total,
        geometry_width: geometry.width,
        geometry_height: geometry.height,
        locked,
        before,
        dimension,
    })
}

fn validate_model_storage_references(
    package: &Package,
    model_identifier: u64,
    model_message_index: usize,
    row_locations: &[BucketLocation],
    column_location: &BucketLocation,
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    let (_, model) = package
        .object_with_component(model_identifier)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    let info = model
        .archive_info
        .message_infos
        .get(model_message_index)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    budget.items(info.object_references.len().saturating_add(1))?;
    budget.fields(info.field_infos.len())?;
    budget.references(
        info.data_references.len().saturating_add(
            info.field_infos
                .iter()
                .map(|field| {
                    field
                        .object_references
                        .len()
                        .saturating_add(field.data_references.len())
                })
                .sum::<usize>(),
        ),
    )?;
    budget.transaction_work(info.field_infos.len())?;
    let storage_count = row_locations.len().saturating_add(1);
    budget.allocations(storage_count.saturating_mul(4))?;
    let mut expected = HashMap::new();
    expected
        .try_reserve(storage_count)
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: storage_count,
        })?;
    let mut expected_paths = HashMap::new();
    expected_paths.try_reserve(storage_count).map_err(|_| {
        SlideTableDimensionError::Allocation {
            amount: storage_count,
        }
    })?;
    for location in row_locations {
        expected.insert(location.identifier, 0usize);
        expected_paths.insert(location.identifier, MODEL_ROW_BUCKET_PATH);
    }
    expected.insert(column_location.identifier, 0usize);
    expected_paths.insert(column_location.identifier, MODEL_COLUMN_BUCKET_PATH);
    if info
        .data_references
        .iter()
        .any(|identifier| expected.contains_key(identifier))
    {
        return Err(SlideTableDimensionError::UnsupportedDependency);
    }
    for identifier in &info.object_references {
        if let Some(count) = expected.get_mut(identifier) {
            *count = (*count).saturating_add(1);
        }
    }
    let mut field_counts = HashMap::new();
    field_counts
        .try_reserve(storage_count)
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: storage_count,
        })?;
    for field in &info.field_infos {
        if field
            .data_references
            .iter()
            .any(|identifier| expected.contains_key(identifier))
        {
            return Err(SlideTableDimensionError::UnsupportedDependency);
        }
        let selected = field.object_references.iter().find_map(|identifier| {
            expected_paths
                .get(identifier)
                .map(|path| (*identifier, *path))
        });
        if let Some((identifier, expected_path)) = selected {
            if field.object_references.as_slice() != [identifier]
                || field.path.as_slice() != expected_path
            {
                return Err(SlideTableDimensionError::UnsupportedDependency);
            }
            let count = field_counts.entry(identifier).or_insert(0usize);
            *count = (*count).saturating_add(1);
        } else if field.path.as_slice() == MODEL_ROW_BUCKET_PATH
            || field.path.as_slice() == MODEL_COLUMN_BUCKET_PATH
        {
            return Err(SlideTableDimensionError::InvalidSource);
        }
    }
    if expected.values().any(|count| *count != 1)
        || expected
            .keys()
            .any(|identifier| field_counts.get(identifier) != Some(&1))
    {
        return Err(SlideTableDimensionError::UnsupportedDependency);
    }
    Ok(())
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideTableDimensionError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTableDimensionError::UnsupportedSource),
    }
}

fn validate_package_metadata(
    package: &Package,
    slide_identifier: u64,
    slide_component: &str,
    table_info_identifier: u64,
    table_info_component: &str,
    model_identifier: u64,
    model_component: &str,
    row_locations: &[BucketLocation],
    column_location: &BucketLocation,
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    let physical_capacity = package
        .state
        .source
        .components()
        .iter()
        .map(|component| component.archive().objects.len())
        .try_fold(0usize, |total, count| total.checked_add(count))
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    budget.allocations(physical_capacity)?;
    budget.transaction_work(physical_capacity)?;
    let mut physical_identifiers = HashSet::new();
    physical_identifiers
        .try_reserve(physical_capacity)
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: physical_capacity,
        })?;
    let mut physical_maximum = 0u64;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.objects(1)?;
            let identifier = object
                .archive_info
                .identifier
                .ok_or(SlideTableDimensionError::InvalidSource)?;
            if !physical_identifiers.insert(identifier) {
                return Err(SlideTableDimensionError::UnsupportedDependency);
            }
            physical_maximum = physical_maximum.max(identifier);
        }
    }
    let target_capacity = row_locations.len().saturating_add(4);
    budget.allocations(target_capacity)?;
    let mut selected_targets = Vec::new();
    selected_targets
        .try_reserve_exact(target_capacity)
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: target_capacity,
        })?;
    selected_targets.push(MetadataTarget {
        identifier: slide_identifier,
        component_name: slide_component,
    });
    selected_targets.push(MetadataTarget {
        identifier: table_info_identifier,
        component_name: table_info_component,
    });
    selected_targets.push(MetadataTarget {
        identifier: model_identifier,
        component_name: model_component,
    });
    for location in row_locations.iter().chain(std::iter::once(column_location)) {
        selected_targets.push(MetadataTarget {
            identifier: location.identifier,
            component_name: location.component.as_ref(),
        });
    }
    let mut payload = None;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                validate_message_header(object, index)?;
                budget.messages(1)?;
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                    && payload.replace(message.data.as_slice()).is_some()
                {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
            }
        }
    }
    let payload = payload.ok_or(SlideTableDimensionError::InvalidSource)?;
    let limits = budget.residual_wire(package)?;
    let remaining_components = budget.remaining_wire(
        budget.components,
        budget.max_components,
        SlideTableDimensionLimitKind::Components,
    )?;
    let remaining_references = budget.remaining_wire(
        budget.references,
        budget.max_references,
        SlideTableDimensionLimitKind::References,
    )?;
    let remaining_items = budget.remaining_wire(
        budget.items,
        budget.max_items,
        SlideTableDimensionLimitKind::PayloadItems,
    )?;
    let options = package_metadata_codec::RewriteOptions::new(
        payload.len().max(1).min(limits.max_input_bytes()),
        limits.max_output_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        remaining_components,
        remaining_references,
        remaining_items,
    );
    let mut visitor = StrictPackageMetadataVisitor::new(&physical_identifiers, &selected_targets)?;
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        payload,
        options,
        &mut visitor,
    )
    .map_err(map_metadata_error)?;
    let report = inspection.report();
    budget.input(report.input_bytes())?;
    budget.output(report.output_bytes())?;
    budget.fields(report.fields())?;
    budget.work(report.work_bytes())?;
    budget.components(report.components_scanned())?;
    budget.references(report.references_scanned())?;
    budget.allocations(report.allocations())?;
    budget.retained(report.retained_bytes())?;
    budget.scratch(report.scratch_bytes())?;
    budget.nesting = budget.nesting.max(report.max_depth() as usize);
    if budget.nesting > budget.max_nesting {
        return Err(SlideTableDimensionError::LimitExceeded {
            kind: SlideTableDimensionLimitKind::WireNesting,
            observed: budget.nesting as u64,
            maximum: budget.max_nesting as u64,
        });
    }
    if visitor.unknown {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    if visitor
        .external_component_identifiers
        .iter()
        .any(|identifier| !visitor.component_identifiers.contains(identifier))
        || visitor
            .external_object_identifiers
            .iter()
            .any(|identifier| !physical_identifiers.contains(identifier))
        || inspection.last_object_identifier() < physical_maximum
        || visitor.authority_invalid
        || visitor.duplicate_uuid
        || visitor.duplicate_component
        || visitor.selected_mismatch
        || selected_targets
            .iter()
            .any(|target| !visitor.object_components.contains_key(&target.identifier))
    {
        return Err(SlideTableDimensionError::UnsupportedDependency);
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct MetadataTarget<'a> {
    identifier: u64,
    component_name: &'a str,
}

fn metadata_component_matches_physical(metadata: &str, physical: &str) -> bool {
    fn basename(name: &str) -> &str {
        name.rsplit('/').next().unwrap_or(name)
    }
    fn without_iwa(name: &str) -> &str {
        name.strip_suffix(".iwa").unwrap_or(name)
    }
    let metadata = without_iwa(basename(metadata));
    let physical = without_iwa(basename(physical));
    metadata == physical
}

struct StrictPackageMetadataVisitor<'a> {
    physical_identifiers: &'a HashSet<u64>,
    selected_targets: &'a [MetadataTarget<'a>],
    unknown: bool,
    authority_invalid: bool,
    duplicate_uuid: bool,
    duplicate_component: bool,
    selected_mismatch: bool,
    component_identifiers: HashSet<u64>,
    external_component_identifiers: HashSet<u64>,
    external_object_identifiers: HashSet<u64>,
    uuid_pairs: HashSet<(u64, u64)>,
    object_bindings: HashMap<(u64, u64), (u64, u64)>,
    object_components: HashMap<u64, u64>,
}

impl<'a> StrictPackageMetadataVisitor<'a> {
    fn new(
        physical_identifiers: &'a HashSet<u64>,
        selected_targets: &'a [MetadataTarget<'a>],
    ) -> Result<Self, SlideTableDimensionError> {
        let capacity = physical_identifiers.len();
        let mut component_identifiers = HashSet::new();
        component_identifiers
            .try_reserve(capacity)
            .map_err(|_| SlideTableDimensionError::Allocation { amount: capacity })?;
        let mut external_component_identifiers = HashSet::new();
        external_component_identifiers
            .try_reserve(capacity)
            .map_err(|_| SlideTableDimensionError::Allocation { amount: capacity })?;
        let mut external_object_identifiers = HashSet::new();
        external_object_identifiers
            .try_reserve(capacity)
            .map_err(|_| SlideTableDimensionError::Allocation { amount: capacity })?;
        let mut uuid_pairs = HashSet::new();
        uuid_pairs
            .try_reserve(capacity)
            .map_err(|_| SlideTableDimensionError::Allocation { amount: capacity })?;
        let mut object_bindings = HashMap::new();
        object_bindings
            .try_reserve(capacity)
            .map_err(|_| SlideTableDimensionError::Allocation { amount: capacity })?;
        let mut object_components = HashMap::new();
        object_components
            .try_reserve(capacity)
            .map_err(|_| SlideTableDimensionError::Allocation { amount: capacity })?;
        Ok(Self {
            physical_identifiers,
            selected_targets,
            unknown: false,
            authority_invalid: false,
            duplicate_uuid: false,
            duplicate_component: false,
            selected_mismatch: false,
            component_identifiers,
            external_component_identifiers,
            external_object_identifiers,
            uuid_pairs,
            object_bindings,
            object_components,
        })
    }
}

impl package_metadata_codec::PackageMetadataVisitor for StrictPackageMetadataVisitor<'_> {
    fn visit_unknown_field(&mut self) -> Result<(), package_metadata_codec::RewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if !component.is_current() {
            self.authority_invalid = true;
        }
        self.component_identifiers
            .try_reserve(1)
            .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        if !self.component_identifiers.insert(component.identifier()) {
            self.duplicate_component = true;
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if !binding.component().is_current() {
            self.authority_invalid = true;
        }
        if !self
            .physical_identifiers
            .contains(&binding.object_identifier())
        {
            self.authority_invalid = true;
        }
        let uuid = binding.uuid();
        let component_identifier = binding.component().identifier();
        let object_identifier = binding.object_identifier();
        let pair = (uuid.lower(), uuid.upper());
        self.uuid_pairs
            .try_reserve(1)
            .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        if pair == (0, 0) || !self.uuid_pairs.insert(pair) {
            self.duplicate_uuid = true;
        }
        if self
            .object_bindings
            .insert(
                (component_identifier, object_identifier),
                (uuid.lower(), uuid.upper()),
            )
            .is_some()
        {
            self.duplicate_uuid = true;
        }
        if self
            .object_components
            .insert(object_identifier, component_identifier)
            .is_some()
        {
            self.duplicate_uuid = true;
        }
        for target in self.selected_targets {
            if target.identifier == object_identifier
                && !metadata_component_matches_physical(
                    binding.component().effective_locator(),
                    target.component_name,
                )
            {
                self.selected_mismatch = true;
            }
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if !reference.source().is_current() || reference.is_versioned() {
            self.authority_invalid = true;
        }
        self.external_component_identifiers
            .try_reserve(1)
            .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        self.external_component_identifiers
            .insert(reference.target_component_identifier());
        if let Some(object_identifier) = reference.object_identifier() {
            self.external_object_identifiers
                .try_reserve(1)
                .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
            self.external_object_identifiers.insert(object_identifier);
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        _reference: package_metadata_codec::DataReferenceDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        self.authority_invalid = true;
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        _owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        self.authority_invalid = true;
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: package_metadata_codec::ComponentDescriptor<'_>,
        _identifier: u64,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        self.authority_invalid = true;
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        _object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        self.authority_invalid = true;
        Ok(())
    }
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideTableDimensionError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideTableDimensionError::EmptySlideName);
            }
            let show = package.show().map_err(map_read_error)?;
            show.select_slide(selector)
                .map_err(|_| SlideTableDimensionError::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideTableDimensionError::SlideNameNotFound)
        },
    }
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), SlideTableDimensionError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    let mut selected = None;
    let mut recognized_role = None;
    for (index, message) in object.messages.iter().enumerate() {
        validate_message_header(object, index)?;
        if is_role_type(message.type_) {
            if recognized_role.replace(message.type_).is_some() || message.type_ != message_type {
                return Err(SlideTableDimensionError::InvalidSource);
            }
        }
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideTableDimensionError::InvalidSource);
        }
    }
    selected.ok_or(SlideTableDimensionError::InvalidSource)
}

fn validate_message_header(
    object: &ArchiveObject,
    index: usize,
) -> Result<(), SlideTableDimensionError> {
    let message = object
        .messages
        .get(index)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    if message.type_ != info.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    Ok(())
}

fn is_role_type(message_type: u32) -> bool {
    MAX_ROLE_MESSAGE_TYPES.contains(&message_type)
}

fn repeated_references(
    payload: &[u8],
    number: u32,
    limits: WireLimits,
    budget: &mut DimensionBudget,
) -> Result<Vec<u64>, SlideTableDimensionError> {
    budget.wire_bytes(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let count = fields
        .fields()
        .filter(|field| field.number() == number)
        .count();
    budget.fields(fields.fields().count())?;
    budget.work(payload.len())?;
    budget.allocations(count)?;
    budget.references(count)?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|_| SlideTableDimensionError::Allocation { amount: count })?;
    for field in fields.fields().filter(|field| field.number() == number) {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideTableDimensionError::InvalidSource);
        }
        values.push(strict_reference(field.payload(), limits, budget)?);
    }
    Ok(values)
}

fn table_parent(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut DimensionBudget,
) -> Result<u64, SlideTableDimensionError> {
    budget.wire_bytes(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut parent = None;
    for field in fields.fields() {
        if field.number() == TABLE_SUPER_FIELD {
            if parent.replace(field).is_some() || field.wire_type() != 2 {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            field.validate_canonical_framing().map_err(map_wire_error)?;
        }
    }
    budget.fields(fields.fields().count())?;
    budget.work(payload.len())?;
    let parent = parent.ok_or(SlideTableDimensionError::InvalidSource)?;
    if parent.wire_type() != 2 {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    budget.wire_bytes(parent.payload().len())?;
    let drawable = WireView::parse_with_limits(parent.payload(), limits).map_err(map_wire_error)?;
    let mut parent_field = None;
    for field in drawable.fields() {
        if field.number() == DRAWABLE_PARENT_FIELD {
            if parent_field.replace(field).is_some() || field.wire_type() != 2 {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            field.validate_canonical_framing().map_err(map_wire_error)?;
        }
    }
    budget.fields(drawable.fields().count())?;
    budget.work(parent.payload().len())?;
    let parent_field = parent_field.ok_or(SlideTableDimensionError::InvalidSource)?;
    if parent_field.wire_type() != 2 {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    strict_reference(parent_field.payload(), limits, budget)
}

fn strict_reference(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut DimensionBudget,
) -> Result<u64, SlideTableDimensionError> {
    budget.wire_bytes(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideTableDimensionError::InvalidSource)?;
                if value == 0 || width != encoded_len(value) {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideTableDimensionError::UnsupportedDependency),
            _ => return Err(SlideTableDimensionError::InvalidSource),
        }
    }
    budget.fields(fields.fields().count())?;
    budget.work(payload.len())?;
    identifier.ok_or(SlideTableDimensionError::InvalidSource)
}

fn checked_reference(reference: codec::ReferenceSnapshot) -> Result<u64, SlideTableDimensionError> {
    if reference.identifier() == 0 || reference.deprecated_is_external() == Some(true) {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    Ok(reference.identifier())
}

fn decode_table_info(
    payload: &[u8],
    package: &Package,
    budget: &mut DimensionBudget,
) -> Result<table_info_codec::TableInfoSnapshot, SlideTableDimensionError> {
    let limits = budget.residual_wire(package)?;
    budget.wire_bytes(payload.len())?;
    let options = table_info_codec::DecodeOptions::new(
        payload.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    );
    let snapshot =
        table_info_codec::decode_table_info(payload, options).map_err(map_table_info_error)?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.fields(view.fields().count())?;
    budget.work(payload.len())?;
    Ok(snapshot)
}

fn parse_model_defaults(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut DimensionBudget,
) -> Result<(f64, f64), SlideTableDimensionError> {
    budget.wire_bytes(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut row = None;
    let mut column = None;
    for field in fields.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            16 | 17 => {
                if field.wire_type() != 1 {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
                let value = u64::from_le_bytes(
                    field
                        .payload()
                        .try_into()
                        .map_err(|_| SlideTableDimensionError::InvalidSource)?,
                );
                let value = f64::from_bits(value);
                if !value.is_finite() || value <= 0.0 {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
                let target = if field.number() == 16 {
                    &mut row
                } else {
                    &mut column
                };
                if target.replace(value).is_some() {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
            },
            _ => {},
        }
    }
    budget.fields(fields.fields().count())?;
    budget.work(payload.len())?;
    Ok((
        row.ok_or(SlideTableDimensionError::InvalidSource)?,
        column.ok_or(SlideTableDimensionError::InvalidSource)?,
    ))
}

#[derive(Clone, Copy)]
struct Geometry {
    width: f32,
    height: f32,
}

fn parse_table_geometry(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut DimensionBudget,
) -> Result<Geometry, SlideTableDimensionError> {
    budget.wire_bytes(payload.len())?;
    let root = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut super_field = None;
    for field in root.fields() {
        if field.number() == TABLE_SUPER_FIELD {
            if super_field.replace(field).is_some() || field.wire_type() != 2 {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            field.validate_canonical_framing().map_err(map_wire_error)?;
        }
    }
    let super_field = super_field.ok_or(SlideTableDimensionError::InvalidSource)?;
    if super_field.wire_type() != 2 {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    budget.wire_bytes(super_field.payload().len())?;
    let drawable =
        WireView::parse_with_limits(super_field.payload(), limits).map_err(map_wire_error)?;
    let mut geometry = None;
    for field in drawable.fields() {
        if field.number() == DRAWABLE_GEOMETRY_FIELD {
            if geometry.replace(field).is_some() || field.wire_type() != 2 {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            field.validate_canonical_framing().map_err(map_wire_error)?;
        }
    }
    let geometry = geometry.ok_or(SlideTableDimensionError::InvalidSource)?;
    if geometry.wire_type() != 2 {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    budget.wire_bytes(geometry.payload().len())?;
    let geometry =
        WireView::parse_with_limits(geometry.payload(), limits).map_err(map_wire_error)?;
    let mut size = None;
    for field in geometry.fields() {
        if field.number() == GEOMETRY_SIZE_FIELD {
            if size.replace(field).is_some() || field.wire_type() != 2 {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            field.validate_canonical_framing().map_err(map_wire_error)?;
        }
    }
    let size = size.ok_or(SlideTableDimensionError::InvalidSource)?;
    if size.wire_type() != 2 {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    budget.wire_bytes(size.payload().len())?;
    let size = WireView::parse_with_limits(size.payload(), limits).map_err(map_wire_error)?;
    let mut width = None;
    let mut height = None;
    for field in size.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            SIZE_WIDTH_FIELD | SIZE_HEIGHT_FIELD => {
                if field.wire_type() != 5 || field.payload().len() != 4 {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
                let value = f32::from_bits(u32::from_le_bytes(
                    field
                        .payload()
                        .try_into()
                        .map_err(|_| SlideTableDimensionError::InvalidSource)?,
                ));
                if !value.is_finite() || value <= 0.0 {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
                let target = if field.number() == SIZE_WIDTH_FIELD {
                    &mut width
                } else {
                    &mut height
                };
                if target.replace(value).is_some() {
                    return Err(SlideTableDimensionError::InvalidSource);
                }
            },
            _ => {},
        }
    }
    budget.fields(
        root.fields()
            .count()
            .saturating_add(drawable.fields().count())
            .saturating_add(geometry.fields().count())
            .saturating_add(size.fields().count()),
    )?;
    budget.work(
        payload
            .len()
            .saturating_add(super_field.payload().len())
            .saturating_add(geometry.fields().count())
            .saturating_add(size.fields().count()),
    )?;
    Ok(Geometry {
        width: width.ok_or(SlideTableDimensionError::InvalidSource)?,
        height: height.ok_or(SlideTableDimensionError::InvalidSource)?,
    })
}

fn object_message_payload(
    package: &Package,
    identifier: u64,
    message_index: usize,
    message_type: u32,
) -> Result<&[u8], SlideTableDimensionError> {
    let (_, object) = package
        .object_with_component(identifier)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    let message = object
        .messages
        .get(message_index)
        .filter(|message| message.type_ == message_type)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    validate_message_header(object, message_index)?;
    Ok(message.data.as_slice())
}

fn object_index_for(
    package: &Package,
    identifier: u64,
    budget: &mut DimensionBudget,
) -> Result<usize, SlideTableDimensionError> {
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for (index, object) in component.archive().objects.iter().enumerate() {
            budget.objects(1)?;
            budget.transaction_work(1)?;
            if object.archive_info.identifier == Some(identifier) {
                return Ok(index);
            }
        }
    }
    Err(SlideTableDimensionError::InvalidSource)
}

#[derive(Clone)]
struct BucketLocation {
    identifier: u64,
    component: Arc<str>,
    object: usize,
    message: usize,
}

fn storage_references(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut DimensionBudget,
) -> Result<Vec<u64>, SlideTableDimensionError> {
    budget.wire_bytes(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let count = fields.fields().filter(|field| field.number() == 2).count();
    budget.fields(fields.fields().count())?;
    budget.work(payload.len())?;
    budget.references(count)?;
    budget.allocations(count)?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(count)
        .map_err(|_| SlideTableDimensionError::Allocation { amount: count })?;
    for field in fields.fields().filter(|field| field.number() == 2) {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideTableDimensionError::InvalidSource);
        }
        references.push(strict_reference(field.payload(), limits, budget)?);
    }
    Ok(references)
}

struct HeaderReader {
    selected: u32,
    minimum: u32,
    maximum: u32,
    maximum_entries: usize,
    found: Option<u32>,
    seen: HashSet<u32>,
    total: f64,
    invalid: bool,
}

impl codec::StorageVisitor for HeaderReader {
    fn visit_header(&mut self, header: codec::HeaderSnapshot) -> Result<(), codec::DecodeError> {
        let index = header.index();
        let bits = header.size_bits();
        let value = f32::from_bits(bits);
        if index < self.minimum
            || index >= self.maximum
            || !value.is_finite()
            || value < 0.0
            || (value == 0.0 && bits != 0)
        {
            self.invalid = true;
        }
        if self.seen.len() >= self.maximum_entries && !self.seen.contains(&index) {
            self.invalid = true;
        } else if !self.seen.insert(index) {
            self.invalid = true;
        }
        let effective = if bits == 0 { 0.0 } else { f64::from(value) };
        self.total += effective;
        if index == self.selected && self.found.replace(bits).is_some() {
            self.invalid = true;
        }
        Ok(())
    }
}

fn locate_bucket(
    package: &Package,
    identifier: u64,
    budget: &mut DimensionBudget,
) -> Result<BucketLocation, SlideTableDimensionError> {
    let mut found = None;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for (object, value) in component.archive().objects.iter().enumerate() {
            budget.objects(1)?;
            if value.archive_info.identifier != Some(identifier) {
                continue;
            }
            if found.is_some() {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            let mut message_index = None;
            let mut role_count = 0usize;
            let mut wrong_role = false;
            for (index, message) in value.messages.iter().enumerate() {
                validate_message_header(value, index)?;
                budget.messages(1)?;
                budget.transaction_work(message.data.len().saturating_add(1))?;
                if is_role_type(message.type_) {
                    role_count = role_count.saturating_add(1);
                    if message.type_ == HEADER_BUCKET_MESSAGE_TYPE {
                        if message_index.replace(index).is_some() {
                            return Err(SlideTableDimensionError::InvalidSource);
                        }
                    } else {
                        wrong_role = true;
                    }
                }
            }
            if wrong_role || role_count != 1 {
                return Err(SlideTableDimensionError::UnsupportedTopology);
            }
            let message = message_index.ok_or(SlideTableDimensionError::InvalidSource)?;
            let message_info = value
                .archive_info
                .message_infos
                .get(message)
                .ok_or(SlideTableDimensionError::InvalidSource)?;
            // Header-storage buckets are leaf records.  A selected bucket
            // must not carry an alternate ArchiveInfo edge or FieldInfo route
            // that could make the in-place rewrite shared or ambiguous.
            if !message_info.object_references.is_empty()
                || !message_info.data_references.is_empty()
                || !message_info.field_infos.is_empty()
            {
                return Err(SlideTableDimensionError::UnsupportedDependency);
            }
            budget.allocations(1)?;
            found = Some(BucketLocation {
                identifier,
                component: Arc::<str>::from(component.name()),
                object,
                message,
            });
        }
    }
    found.ok_or(SlideTableDimensionError::InvalidSource)
}

fn collect_axis(
    package: &Package,
    locations: &[BucketLocation],
    limit: u32,
    dimension: Dimension,
    default: f64,
    budget: &mut DimensionBudget,
) -> Result<(f32, Option<Size>), SlideTableDimensionError> {
    let selected =
        u32::try_from(dimension.index()).map_err(|_| SlideTableDimensionError::InvalidSource)?;
    if selected >= limit {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    let mut total = 0.0f64;
    let mut selected_size = None;
    for (slot, location) in locations.iter().enumerate() {
        let payload = object_message_payload(
            package,
            location.identifier,
            location.message,
            HEADER_BUCKET_MESSAGE_TYPE,
        )?;
        let (minimum, maximum) = match dimension {
            Dimension::Row(_) => {
                let minimum = u32::try_from(slot)
                    .ok()
                    .and_then(|slot| slot.checked_mul(DEFAULT_HEADER_BUCKET_ROWS))
                    .ok_or(SlideTableDimensionError::InvalidSource)?;
                (
                    minimum,
                    minimum
                        .saturating_add(DEFAULT_HEADER_BUCKET_ROWS)
                        .min(limit),
                )
            },
            Dimension::Column(_) => (0, limit),
        };
        let count = usize::try_from(maximum.saturating_sub(minimum))
            .map_err(|_| SlideTableDimensionError::InvalidSource)?;
        budget.allocations(count)?;
        let mut seen = HashSet::new();
        seen.try_reserve(count)
            .map_err(|_| SlideTableDimensionError::Allocation { amount: count })?;
        let mut reader = HeaderReader {
            selected,
            minimum,
            maximum,
            maximum_entries: count,
            found: None,
            seen,
            total: 0.0,
            invalid: false,
        };
        let options = budget.codec_options(package, payload.len())?;
        let (_, report) =
            codec::decode_header_storage_bucket_with_visitor(payload, options, &mut reader)
                .map_err(map_codec_error)?;
        budget.codec_report(report)?;
        let missing = count.saturating_sub(reader.seen.len());
        let bucket_total =
            reader.total + default * f64::from(u32::try_from(missing).unwrap_or(u32::MAX));
        total += bucket_total;
        if reader.invalid {
            return Err(SlideTableDimensionError::InvalidSource);
        }
        if reader.found.is_some() {
            selected_size = Some(
                reader
                    .found
                    .map(size_from_bits)
                    .transpose()?
                    .ok_or(SlideTableDimensionError::InvalidSource)?,
            );
        }
    }
    let total = total as f32;
    Ok((total, selected_size))
}

fn size_from_bits(bits: u32) -> Result<Size, SlideTableDimensionError> {
    if bits == 0 {
        return Ok(Size::Default);
    }
    Points::new(f32::from_bits(bits))
        .map(Size::Points)
        .map_err(|_| SlideTableDimensionError::InvalidSource)
}

fn validate_geometry_matches(
    geometry: Geometry,
    rows: f32,
    columns: f32,
) -> Result<(), SlideTableDimensionError> {
    fn close(left: f32, right: f32) -> bool {
        let tolerance = 0.01f32.max(left.abs().max(right.abs()) * 0.0001);
        (left - right).abs() <= tolerance
    }
    if !close(geometry.width, columns) || !close(geometry.height, rows) {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    Ok(())
}

fn reject_duplicates(
    values: &[u64],
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    budget.allocations(values.len())?;
    budget.transaction_work(values.len())?;
    let mut seen = HashSet::new();
    seen.try_reserve(values.len())
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: values.len(),
        })?;
    for value in values {
        if !seen.insert(*value) {
            return Err(SlideTableDimensionError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_slide_metadata(
    object: &ArchiveObject,
    message_index: usize,
    owned: &[u64],
    z_order: &[u64],
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    budget.fields(info.field_infos.len())?;
    budget.references(
        info.object_references
            .len()
            .saturating_add(info.data_references.len()),
    )?;
    let expected_capacity = owned.len().saturating_add(z_order.len());
    budget.allocations(expected_capacity)?;
    let mut expected_counts = HashMap::new();
    expected_counts
        .try_reserve(expected_capacity)
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: expected_capacity,
        })?;
    for identifier in owned.iter().chain(z_order) {
        expected_counts.entry(*identifier).or_insert(0usize);
    }
    budget.transaction_work(
        info.field_infos
            .len()
            .saturating_add(info.object_references.len()),
    )?;
    for identifier in &info.object_references {
        if let Some(count) = expected_counts.get_mut(identifier) {
            *count = (*count).saturating_add(1);
        }
    }
    let mut owned_field = false;
    let mut z_order_field = false;
    for field in &info.field_infos {
        budget.references(
            field
                .object_references
                .len()
                .saturating_add(field.data_references.len()),
        )?;
        if field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD] {
            if owned_field
                || field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != owned
            {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            owned_field = true;
            for identifier in &field.object_references {
                if let Some(count) = expected_counts.get_mut(identifier) {
                    *count = (*count).saturating_add(1);
                }
            }
        } else if field.path.as_slice() == [SLIDE_Z_ORDER_FIELD] {
            if z_order_field
                || field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != z_order
            {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            z_order_field = true;
            for identifier in &field.object_references {
                if let Some(count) = expected_counts.get_mut(identifier) {
                    *count = (*count).saturating_add(1);
                }
            }
        } else if !field.object_references.is_empty() || !field.data_references.is_empty() {
            return Err(SlideTableDimensionError::UnsupportedTopology);
        }
    }
    if owned_field != z_order_field {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    if expected_counts
        .values()
        .any(|count| *count == 0 || *count > 3)
    {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    Ok(())
}

fn validate_table_info_metadata(
    object: &ArchiveObject,
    message_index: usize,
    model: u64,
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    budget.fields(info.field_infos.len())?;
    budget.references(
        info.object_references
            .len()
            .saturating_add(info.data_references.len()),
    )?;
    if info
        .object_references
        .iter()
        .filter(|id| **id == model)
        .count()
        != 1
    {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    let mut model_field = false;
    for field in &info.field_infos {
        budget.references(
            field
                .object_references
                .len()
                .saturating_add(field.data_references.len()),
        )?;
        if field.path.as_slice() == [TABLE_MODEL_FIELD] {
            if model_field
                || field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != [model]
            {
                return Err(SlideTableDimensionError::InvalidSource);
            }
            model_field = true;
        } else if !field.object_references.is_empty() || !field.data_references.is_empty() {
            return Err(SlideTableDimensionError::UnsupportedTopology);
        }
    }
    Ok(())
}

fn ensure_unique_identity(
    package: &Package,
    identifier: u64,
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    let mut count = 0usize;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.objects(1)?;
            if object.archive_info.identifier == Some(identifier) {
                count = count.saturating_add(1);
            }
        }
    }
    if count == 1 {
        Ok(())
    } else {
        Err(SlideTableDimensionError::UnsupportedDependency)
    }
}

fn ensure_unique_table_owner(
    package: &Package,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    let mut owned = 0usize;
    let mut z_order = 0usize;
    let mut selected = false;
    let mut model_owners = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                budget.messages(1)?;
                if message.type_ == SLIDE_MESSAGE_TYPE {
                    let owned_refs = repeated_references(
                        &message.data,
                        SLIDE_OWNED_DRAWABLES_FIELD,
                        budget.residual_wire(package)?,
                        budget,
                    )?;
                    let z_refs = repeated_references(
                        &message.data,
                        SLIDE_Z_ORDER_FIELD,
                        budget.residual_wire(package)?,
                        budget,
                    )?;
                    let owned_hits = owned_refs
                        .iter()
                        .filter(|id| **id == table_info_identifier)
                        .count();
                    let z_hits = z_refs
                        .iter()
                        .filter(|id| **id == table_info_identifier)
                        .count();
                    owned = owned.saturating_add(owned_hits);
                    z_order = z_order.saturating_add(z_hits);
                    if object.archive_info.identifier == Some(slide_identifier)
                        && owned_hits == 1
                        && z_hits == 1
                    {
                        selected = true;
                    }
                } else if message.type_ == TABLE_INFO_MESSAGE_TYPE {
                    let info = decode_table_info(&message.data, package, budget)?;
                    if info.table_model().identifier().get() == model_identifier {
                        model_owners = model_owners.saturating_add(1);
                    }
                }
                let _ = index;
            }
        }
    }
    if owned == 1 && z_order == 1 && selected && model_owners == 1 {
        Ok(())
    } else {
        Err(SlideTableDimensionError::UnsupportedDependency)
    }
}

fn validate_global_references(
    package: &Package,
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    model_message_index: usize,
    row_locations: &[BucketLocation],
    column_location: &BucketLocation,
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    let limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let slide = package
        .object_with_component(slide_identifier)
        .map(|(_, object)| object)
        .ok_or(SlideTableDimensionError::UnsupportedDependency)?;
    let slide_info = slide
        .archive_info
        .message_infos
        .get(slide_message_index)
        .ok_or(SlideTableDimensionError::UnsupportedDependency)?;
    let expected_table_edges = slide_info
        .object_references
        .iter()
        .chain(
            slide_info
                .field_infos
                .iter()
                .flat_map(|field| field.object_references.iter()),
        )
        .filter(|identifier| **identifier == table_info_identifier)
        .count();
    let table_info = package
        .object_with_component(table_info_identifier)
        .map(|(_, object)| object)
        .ok_or(SlideTableDimensionError::UnsupportedDependency)?;
    let table_info_info = table_info
        .archive_info
        .message_infos
        .get(table_info_message_index)
        .ok_or(SlideTableDimensionError::UnsupportedDependency)?;
    let expected_model_edges = table_info_info
        .object_references
        .iter()
        .chain(
            table_info_info
                .field_infos
                .iter()
                .flat_map(|field| field.object_references.iter()),
        )
        .filter(|identifier| **identifier == model_identifier)
        .count();
    if expected_table_edges == 0 || expected_model_edges == 0 {
        return Err(SlideTableDimensionError::UnsupportedDependency);
    }
    let storage_capacity = row_locations.len().saturating_add(1);
    budget.allocations(storage_capacity)?;
    let mut storage_identifiers = HashSet::new();
    storage_identifiers
        .try_reserve(storage_capacity)
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: storage_capacity,
        })?;
    for location in row_locations.iter().chain(std::iter::once(column_location)) {
        if !storage_identifiers.insert(location.identifier) {
            return Err(SlideTableDimensionError::UnsupportedDependency);
        }
    }
    budget.transaction_work(storage_capacity)?;
    let mut visitor = GlobalReferenceVisitor {
        slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
        model_message_index,
        table_edges: 0,
        model_edges: 0,
        expected_table_edges,
        expected_model_edges,
        storage_identifiers: &storage_identifiers,
        invalid: false,
    };
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let info_items = object
                .archive_info
                .message_infos
                .iter()
                .map(|info| {
                    info.field_infos
                        .len()
                        .saturating_add(info.object_references.len())
                        .saturating_add(info.data_references.len())
                        .saturating_add(1)
                })
                .sum::<usize>();
            budget.items(info_items)?;
            budget.references(
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .map(|info| {
                        info.object_references
                            .len()
                            .saturating_add(info.data_references.len())
                            .saturating_add(
                                info.field_infos
                                    .iter()
                                    .map(|field| {
                                        field
                                            .object_references
                                            .len()
                                            .saturating_add(field.data_references.len())
                                    })
                                    .sum::<usize>(),
                            )
                    })
                    .sum::<usize>(),
            )?;
            budget.work(object.data_length as usize)?;
            object
                .inspect_references_with_policy_and_limits(
                    &mut visitor,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    limits,
                )
                .map_err(map_core_error)?;
        }
    }
    if visitor.invalid
        || visitor.table_edges != visitor.expected_table_edges
        || visitor.model_edges != visitor.expected_model_edges
    {
        return Err(SlideTableDimensionError::UnsupportedDependency);
    }
    // Every storage bucket must have one physical object identity.  This also
    // prevents a later rewrite from silently selecting an alias in another
    // component.
    for location in row_locations.iter().chain(std::iter::once(column_location)) {
        ensure_unique_identity(package, location.identifier, budget)?;
    }
    Ok(())
}

struct GlobalReferenceVisitor<'a> {
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    model_message_index: usize,
    table_edges: usize,
    model_edges: usize,
    expected_table_edges: usize,
    expected_model_edges: usize,
    storage_identifiers: &'a HashSet<u64>,
    invalid: bool,
}

impl ArchiveReferenceVisitor for GlobalReferenceVisitor<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        let storage_target = self
            .storage_identifiers
            .contains(&occurrence.referenced_identifier);
        if storage_target
            && (occurrence.kind != ArchiveReferenceKind::Object
                || occurrence.object_identifier != self.model_identifier
                || occurrence.message_index != self.model_message_index)
        {
            self.invalid = true;
        }
        if occurrence.kind != ArchiveReferenceKind::Object {
            if occurrence.referenced_identifier == self.table_info_identifier
                || occurrence.referenced_identifier == self.model_identifier
            {
                self.invalid = true;
            }
            return Ok(());
        }
        if occurrence.referenced_identifier == self.table_info_identifier {
            if occurrence.object_identifier == self.slide_identifier
                && occurrence.message_index == self.slide_message_index
            {
                self.table_edges = self.table_edges.saturating_add(1);
            } else {
                self.invalid = true;
            }
        }
        if occurrence.referenced_identifier == self.model_identifier {
            if occurrence.object_identifier == self.table_info_identifier
                && occurrence.message_index == self.table_info_message_index
            {
                self.model_edges = self.model_edges.saturating_add(1);
            } else {
                self.invalid = true;
            }
        }
        // A selected model may point at storage buckets.  Their exact route is
        // checked against the model's wire payload and the bucket identities;
        // metadata references to unrelated selected roots are not accepted.
        Ok(())
    }
}

fn same_selection(left: &DimensionSelection, right: &DimensionSelection) -> bool {
    left.slide_position == right.slide_position
        && left.table_position == right.table_position
        && left.slide_identifier == right.slide_identifier
        && left.table_info_identifier == right.table_info_identifier
        && left.table_info_component == right.table_info_component
        && left.table_info_object == right.table_info_object
        && left.table_info_message == right.table_info_message
        && left.model_identifier == right.model_identifier
        && left.model_component == right.model_component
        && left.model_object == right.model_object
        && left.model_message == right.model_message
        && left.bucket_identifier == right.bucket_identifier
        && left.bucket_component == right.bucket_component
        && left.bucket_object == right.bucket_object
        && left.bucket_message == right.bucket_message
        && left.dimension == right.dimension
}

fn rewrite_dimension(
    source: &Package,
    selection: &DimensionSelection,
    after: Size,
    previews: &[&'static str],
    budget: &mut DimensionBudget,
) -> Result<Package, SlideTableDimensionError> {
    let catalog = physical_catalog(source)?;
    let physical_limits = source.state.options.archive();
    let limits = budget.residual_wire(source)?;
    let (after_value, row_total, column_total) = match selection.dimension {
        Dimension::Row(_) => {
            let value = effective_size(after, selection.default_row_height)?;
            let old = effective_size(selection.before, selection.default_row_height)?;
            (
                value,
                checked_total(selection.row_total, old, value)?,
                selection.column_total,
            )
        },
        Dimension::Column(_) => {
            let value = effective_size(after, selection.default_column_width)?;
            let old = effective_size(selection.before, selection.default_column_width)?;
            (
                value,
                selection.row_total,
                checked_total(selection.column_total, old, value)?,
            )
        },
    };
    let _ = after_value;
    let bucket_payload = object_message_payload(
        source,
        selection.bucket_identifier,
        selection.bucket_message,
        HEADER_BUCKET_MESSAGE_TYPE,
    )?;
    let index = u32::try_from(selection.dimension.index())
        .map_err(|_| SlideTableDimensionError::InvalidSource)?;
    let limit = match selection.dimension {
        Dimension::Row(_) => selection.rows,
        Dimension::Column(_) => selection.columns,
    };
    let edit = match after {
        Size::Default => codec::HeaderSizeEdit::remove(index),
        Size::Points(points) => codec::HeaderSizeEdit::set(index, points.value().to_bits()),
    };
    let planning_options = budget.codec_options(source, bucket_payload.len())?;
    let plan =
        codec::plan_header_storage_bucket_sizes(bucket_payload, limit, &[edit], planning_options)
            .map_err(map_codec_error)?;
    let requirements = plan.requirements();
    budget.codec_report(requirements.source())?;
    let upper = requirements.result_upper_bound();
    budget.input(upper.source_bytes())?;
    budget.fields(upper.fields())?;
    budget.work(upper.work_bytes())?;
    budget.references(upper.references())?;
    budget.transaction_work(requirements.rewrite_work_bytes())?;
    let execution_options = budget.codec_options(source, requirements.output_bytes())?;
    let (rewritten_bucket, report) =
        codec::execute_header_storage_bucket_size_plan(plan, execution_options)
            .map_err(map_codec_error)?;
    budget.codec_report(report.result())?;

    let table_info_payload = object_message_payload(
        source,
        selection.table_info_identifier,
        selection.table_info_message,
        TABLE_INFO_MESSAGE_TYPE,
    )?;
    let (width, height) = (column_total, row_total);
    let width_path = [
        TABLE_SUPER_FIELD,
        DRAWABLE_GEOMETRY_FIELD,
        GEOMETRY_SIZE_FIELD,
        SIZE_WIDTH_FIELD,
    ];
    let height_path = [
        TABLE_SUPER_FIELD,
        DRAWABLE_GEOMETRY_FIELD,
        GEOMETRY_SIZE_FIELD,
        SIZE_HEIGHT_FIELD,
    ];
    let geometry_edits = [
        NestedFieldEdit::new(
            &width_path,
            true,
            NestedFieldReplacement::Fixed32(Some(width.to_bits())),
        ),
        NestedFieldEdit::new(
            &height_path,
            true,
            NestedFieldReplacement::Fixed32(Some(height.to_bits())),
        ),
    ];
    let rewritten_geometry =
        patch_nested_fields_batched_with_limits(table_info_payload, &geometry_edits, limits)
            .map_err(map_wire_error)?;
    budget.wire_bytes(table_info_payload.len())?;
    budget.wire_output(rewritten_geometry.len())?;
    budget.fields(geometry_edits.len())?;
    budget.work(
        table_info_payload
            .len()
            .saturating_add(rewritten_geometry.len()),
    )?;
    budget.allocations(rewritten_geometry.len())?;

    let bucket_replacement = NativeReplacement {
        identifier: selection.bucket_identifier,
        message: selection.bucket_message,
        message_type: HEADER_BUCKET_MESSAGE_TYPE,
        payload: &rewritten_bucket,
    };
    let table_replacement = NativeReplacement {
        identifier: selection.table_info_identifier,
        message: selection.table_info_message,
        message_type: TABLE_INFO_MESSAGE_TYPE,
        payload: &rewritten_geometry,
    };
    let replacement_count = if selection.bucket_component == selection.table_info_component {
        1
    } else {
        2
    };
    let mut compressed_entries = Vec::new();
    compressed_entries
        .try_reserve_exact(replacement_count)
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: replacement_count,
        })?;
    budget.allocations(replacement_count)?;
    if selection.bucket_component == selection.table_info_component {
        let replacements = [bucket_replacement, table_replacement];
        compressed_entries.push(rewrite_component(
            source,
            selection.bucket_component.as_ref(),
            &replacements,
            budget,
        )?);
    } else {
        compressed_entries.push(rewrite_component(
            source,
            selection.bucket_component.as_ref(),
            std::slice::from_ref(&bucket_replacement),
            budget,
        )?);
        compressed_entries.push(rewrite_component(
            source,
            selection.table_info_component.as_ref(),
            std::slice::from_ref(&table_replacement),
            budget,
        )?);
    }
    let mut edits = Vec::new();
    edits.try_reserve_exact(replacement_count).map_err(|_| {
        SlideTableDimensionError::Allocation {
            amount: replacement_count,
        }
    })?;
    budget.allocations(replacement_count)?;
    if selection.bucket_component == selection.table_info_component {
        edits.push(EntryEdit::new(
            selection.bucket_component.as_ref(),
            compressed_entries[0].as_slice(),
        ));
    } else {
        edits.push(EntryEdit::new(
            selection.bucket_component.as_ref(),
            compressed_entries[0].as_slice(),
        ));
        edits.push(EntryEdit::new(
            selection.table_info_component.as_ref(),
            compressed_entries[1].as_slice(),
        ));
    }
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, previews, physical_limits)
        .map_err(map_archive_error)?;
    let reassembly = prepared.execution_requirements();
    budget.reassembly(reassembly)?;
    let output = prepared
        .execute(reassembly.exact_limits())
        .map_err(map_archive_error)?;
    budget.input(output.len())?;
    budget.transaction_work(output.len())?;
    budget.allocations(1)?;
    budget.retained(output.len())?;
    Package::from_source_with_options(output.into(), source.state.options).map_err(map_read_error)
}

#[derive(Clone, Copy)]
struct NativeReplacement<'a> {
    identifier: u64,
    message: usize,
    message_type: u32,
    payload: &'a [u8],
}

fn rewrite_component(
    source: &Package,
    component_name: &str,
    replacements: &[NativeReplacement<'_>],
    budget: &mut DimensionBudget,
) -> Result<Vec<u8>, SlideTableDimensionError> {
    let catalog = physical_catalog(source)?;
    budget.transaction_work(catalog.package().len())?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableDimensionError::UnsupportedSource);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    budget.allocations(1)?;
    budget.retained(entry.data().len())?;
    budget.physical(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        source
            .state
            .options
            .archive()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    budget.physical(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    charge_archive_inventory(&archive, budget)?;
    for replacement in replacements {
        let object = archive
            .object_mut(replacement.identifier)
            .ok_or(SlideTableDimensionError::InvalidSource)?;
        validate_message_header(object, replacement.message)?;
        let message = object
            .messages
            .get(replacement.message)
            .ok_or(SlideTableDimensionError::InvalidSource)?;
        if message.type_ != replacement.message_type {
            return Err(SlideTableDimensionError::InvalidSource);
        }
        budget.allocations(replacement.payload.len())?;
        budget.retained(replacement.payload.len())?;
        object
            .replace_message_preserving_header_with_limits(
                replacement.message,
                RawMessage {
                    type_: replacement.message_type,
                    data: replacement.payload.to_vec(),
                },
                archive_limits,
            )
            .map_err(map_core_error)?;
    }
    let encoded_length = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_length).map_err(map_core_error)?;
    budget.output(encoded_length.saturating_add(compressed_bound))?;
    budget.allocations(encoded_length)?;
    budget.retained(encoded_length)?;
    budget.allocations(compressed_bound)?;
    budget.retained(compressed_bound)?;
    let encoded = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if encoded.len() != encoded_length {
        return Err(SlideTableDimensionError::Verification);
    }
    let compressed = SnappyStream::compress(&encoded).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(SlideTableDimensionError::Verification);
    }
    Ok(compressed)
}

fn effective_size(size: Size, default: f64) -> Result<f32, SlideTableDimensionError> {
    let value = match size {
        Size::Default => default,
        Size::Points(points) => f64::from(points.value()),
    };
    if !value.is_finite() || value <= 0.0 || value > f64::from(f32::MAX) {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    Ok(value as f32)
}

fn checked_total(total: f32, old: f32, new: f32) -> Result<f32, SlideTableDimensionError> {
    let value = f64::from(total) - f64::from(old) + f64::from(new);
    if !value.is_finite() || value <= 0.0 || value > f64::from(f32::MAX) {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    Ok(value as f32)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &DimensionSelection,
    _source_selection: &DimensionSelection,
    _after: Size,
    source_previews: &[&'static str],
    target_previews: &[&'static str],
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    budget.entries(
        source_catalog
            .package()
            .len()
            .saturating_add(candidate_catalog.package().len()),
    )?;
    budget.allocations(
        source_catalog
            .package()
            .len()
            .saturating_add(candidate_catalog.package().len()),
    )?;
    let mut source_names = HashSet::new();
    source_names
        .try_reserve(source_catalog.package().len())
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: source_catalog.package().len(),
        })?;
    for entry in source_catalog.package().iter() {
        if !source_names.insert(entry.name()) {
            return Err(SlideTableDimensionError::Verification);
        }
    }
    let mut candidate_entries = HashMap::new();
    candidate_entries
        .try_reserve(candidate_catalog.package().len())
        .map_err(|_| SlideTableDimensionError::Allocation {
            amount: candidate_catalog.package().len(),
        })?;
    for entry in candidate_catalog.package().iter() {
        if candidate_entries.insert(entry.name(), entry).is_some() {
            return Err(SlideTableDimensionError::Verification);
        }
    }
    for entry in source_catalog.package().iter() {
        budget.work(
            entry
                .data()
                .len()
                .saturating_add(entry.raw_record().local_record().len())
                .saturating_add(entry.raw_record().central_directory_record().len()),
        )?;
        let candidate_entry = candidate_entries.get(entry.name()).copied();
        let preview = source_previews.contains(&entry.name());
        if preview && !target_previews.contains(&entry.name()) {
            if candidate_entry.is_some() {
                return Err(SlideTableDimensionError::Verification);
            }
            continue;
        }
        let candidate_entry = candidate_entry.ok_or(SlideTableDimensionError::Verification)?;
        let changed = entry.name() == selection.bucket_component.as_ref()
            || entry.name() == selection.table_info_component.as_ref();
        if !changed
            && (entry.data() != candidate_entry.data()
                || entry.metadata() != candidate_entry.metadata()
                || entry.raw_record().local_record() != candidate_entry.raw_record().local_record()
                || !same_central_directory_record(
                    entry.raw_record().central_directory_record(),
                    candidate_entry.raw_record().central_directory_record(),
                ))
        {
            return Err(SlideTableDimensionError::Verification);
        }
    }
    for entry in candidate_catalog.package().iter() {
        if !source_names.contains(entry.name()) && !target_previews.contains(&entry.name()) {
            return Err(SlideTableDimensionError::Verification);
        }
    }

    let mut checked_components = [
        selection.bucket_component.as_ref(),
        selection.table_info_component.as_ref(),
    ];
    if checked_components[0] == checked_components[1] {
        checked_components[1] = "";
    }
    for component_name in checked_components {
        if component_name.is_empty() {
            continue;
        }
        let source_archive = component_archive(source, component_name, budget)?;
        let candidate_archive = component_archive(candidate, component_name, budget)?;
        if source_archive.objects.len() != candidate_archive.objects.len() {
            return Err(SlideTableDimensionError::Verification);
        }
        for source_object in &source_archive.objects {
            let identifier = source_object
                .archive_info
                .identifier
                .ok_or(SlideTableDimensionError::Verification)?;
            let candidate_object = candidate_archive
                .object(identifier)
                .ok_or(SlideTableDimensionError::Verification)?;
            let selected_object = identifier == selection.bucket_identifier
                || identifier == selection.table_info_identifier;
            if !selected_object && !source_object.same_content_ignoring_offsets(candidate_object) {
                return Err(SlideTableDimensionError::Verification);
            }
            if selected_object {
                if !same_archive_info_ignoring_message_length(
                    &source_object.archive_info,
                    &candidate_object.archive_info,
                    allowed_message_for_identifier(identifier, selection),
                ) {
                    return Err(SlideTableDimensionError::Verification);
                }
                let allowed_message = allowed_message_for_identifier(identifier, selection);
                for (index, source_message) in source_object.messages.iter().enumerate() {
                    if index != allowed_message {
                        let candidate_message = candidate_object
                            .messages
                            .get(index)
                            .ok_or(SlideTableDimensionError::Verification)?;
                        if source_message.type_ != candidate_message.type_
                            || source_message.data != candidate_message.data
                        {
                            return Err(SlideTableDimensionError::Verification);
                        }
                    }
                }
            }
        }
    }
    Ok(())
}

fn allowed_message_for_identifier(identifier: u64, selection: &DimensionSelection) -> usize {
    if identifier == selection.bucket_identifier {
        selection.bucket_message
    } else {
        selection.table_info_message
    }
}

fn same_archive_info_ignoring_message_length(
    left: &litchi_iwa_core::ArchiveInfo,
    right: &litchi_iwa_core::ArchiveInfo,
    allowed_message: usize,
) -> bool {
    left.identifier == right.identifier
        && left.should_merge == right.should_merge
        && left.message_infos.len() == right.message_infos.len()
        && left
            .message_infos
            .iter()
            .zip(&right.message_infos)
            .enumerate()
            .all(|(index, (left, right))| {
                (index == allowed_message || left.length == right.length)
                    && left.type_ == right.type_
                    && left.versions == right.versions
                    && left.field_infos == right.field_infos
                    && left.object_references == right.object_references
                    && left.data_references == right.data_references
                    && left.base_message_index == right.base_message_index
                    && left.diff_merge_version == right.diff_merge_version
                    && left.diff_field_path == right.diff_field_path
                    && left.fields_to_remove == right.fields_to_remove
                    && left.diff_read_version == right.diff_read_version
            })
}

fn same_central_directory_record(left: &[u8], right: &[u8]) -> bool {
    // The relative local-header offset is the only central-directory field
    // expected to move when an earlier member changes size.  Preserve every
    // other producer-authored central record byte, including names/extras and
    // flags, while allowing that ZIP bookkeeping field to churn.
    left.len() == right.len()
        && (left == right
            || (left.len() >= 46 && left[..42] == right[..42] && left[46..] == right[46..]))
}

fn charge_archive_inventory(
    archive: &Archive,
    budget: &mut DimensionBudget,
) -> Result<(), SlideTableDimensionError> {
    budget.objects(archive.objects.len())?;
    let mut messages = 0usize;
    let mut fields = 0usize;
    let mut references = 0usize;
    let mut items = 0usize;
    let mut work = 0usize;
    for object in &archive.objects {
        messages = messages.saturating_add(object.messages.len());
        work = work.saturating_add(object.data_length as usize);
        for info in &object.archive_info.message_infos {
            fields = fields.saturating_add(info.field_infos.len());
            references = references.saturating_add(
                info.object_references
                    .len()
                    .saturating_add(info.data_references.len()),
            );
            items = items.saturating_add(1 + info.field_infos.len());
            for field in &info.field_infos {
                references = references.saturating_add(
                    field
                        .object_references
                        .len()
                        .saturating_add(field.data_references.len()),
                );
            }
        }
    }
    budget.messages(messages)?;
    budget.fields(fields)?;
    budget.references(references)?;
    budget.items(items)?;
    budget.work(work)?;
    Ok(())
}

fn component_archive(
    package: &Package,
    name: &str,
    budget: &mut DimensionBudget,
) -> Result<Archive, SlideTableDimensionError> {
    let catalog = physical_catalog(package)?;
    budget.transaction_work(catalog.package().len())?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(SlideTableDimensionError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableDimensionError::InvalidSource);
    }
    budget.allocations(1)?;
    budget.retained(entry.data().len())?;
    budget.physical(entry.data().len())?;
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
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    budget.physical(stream.as_bytes().len())?;
    let archive = Archive::parse_with_limits(
        stream.as_bytes(),
        package
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    charge_archive_inventory(&archive, budget)?;
    Ok(archive)
}

fn map_read_error(error: ReadError) -> SlideTableDimensionError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableDimensionError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableDimensionLimitKind::References,
                _ => SlideTableDimensionLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableDimensionError::Allocation { amount },
        _ => SlideTableDimensionError::InvalidSource,
    }
}

fn map_codec_error(error: codec::DecodeError) -> SlideTableDimensionError {
    if let Some(amount) = error.allocation_requested() {
        return SlideTableDimensionError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return SlideTableDimensionError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::Bytes { observed, maximum } => {
            (SlideTableDimensionLimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::References { observed, maximum } => {
            (SlideTableDimensionLimitKind::References, observed, maximum)
        },
        codec::DecodeLimit::Text { observed, maximum } => {
            (SlideTableDimensionLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (SlideTableDimensionLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::Work { observed, maximum } => {
            (SlideTableDimensionLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => {
            return SlideTableDimensionError::LimitExceeded {
                kind: SlideTableDimensionLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        codec::DecodeLimit::Allocation { requested } => {
            return SlideTableDimensionError::Allocation { amount: requested };
        },
        codec::DecodeLimit::Retained { observed, maximum } => {
            (SlideTableDimensionLimitKind::Retained, observed, maximum)
        },
        _ => return SlideTableDimensionError::InvalidSource,
    };
    SlideTableDimensionError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_table_info_error(error: table_info_codec::DecodeError) -> SlideTableDimensionError {
    if let Some(amount) = error.allocation_amount() {
        return SlideTableDimensionError::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideTableDimensionError::LimitExceeded {
            kind: SlideTableDimensionLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideTableDimensionError::LimitExceeded {
            kind: SlideTableDimensionLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.output_limit_values() {
        return SlideTableDimensionError::LimitExceeded {
            kind: SlideTableDimensionLimitKind::WireOutputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.allocation_limit_values() {
        return SlideTableDimensionError::LimitExceeded {
            kind: SlideTableDimensionLimitKind::Allocations,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.retained_limit_values() {
        return SlideTableDimensionError::LimitExceeded {
            kind: SlideTableDimensionLimitKind::Retained,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.scratch_limit_values() {
        return SlideTableDimensionError::LimitExceeded {
            kind: SlideTableDimensionLimitKind::Scratch,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            table_info_codec::WireResourceLimit::Bytes { observed, maximum } => {
                SlideTableDimensionError::LimitExceeded {
                    kind: SlideTableDimensionLimitKind::WireBytes,
                    observed: observed.unwrap_or_default() as u64,
                    maximum: maximum.unwrap_or_default() as u64,
                }
            },
            table_info_codec::WireResourceLimit::Nesting { observed, maximum } => {
                SlideTableDimensionError::LimitExceeded {
                    kind: SlideTableDimensionLimitKind::WireNesting,
                    observed: observed.unwrap_or_default() as u64,
                    maximum: maximum.unwrap_or_default() as u64,
                }
            },
            _ => SlideTableDimensionError::InvalidSource,
        };
    }
    SlideTableDimensionError::InvalidSource
}

fn map_wire_error(_error: litchi_iwa_common::Error) -> SlideTableDimensionError {
    SlideTableDimensionError::InvalidSource
}

fn map_metadata_error(error: package_metadata_codec::RewriteError) -> SlideTableDimensionError {
    if let Some(amount) = error.allocation_request() {
        return SlideTableDimensionError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return SlideTableDimensionError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
            (SlideTableDimensionLimitKind::InputBytes, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
            (SlideTableDimensionLimitKind::OutputBytes, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
            (SlideTableDimensionLimitKind::WireFields, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
            (SlideTableDimensionLimitKind::WireWork, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
            return SlideTableDimensionError::LimitExceeded {
                kind: SlideTableDimensionLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        package_metadata_codec::RewriteLimit::Components { observed, maximum } => {
            (SlideTableDimensionLimitKind::Components, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::References { observed, maximum } => {
            (SlideTableDimensionLimitKind::References, observed, maximum)
        },
        package_metadata_codec::RewriteLimit::Additions { observed, maximum } => {
            (SlideTableDimensionLimitKind::Allocations, observed, maximum)
        },
        _ => return SlideTableDimensionError::InvalidSource,
    };
    SlideTableDimensionError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTableDimensionError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableDimensionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    SlideTableDimensionLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideTableDimensionLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideTableDimensionLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideTableDimensionLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    SlideTableDimensionLimitKind::TotalBytes
                },
                _ => SlideTableDimensionLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableDimensionError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTableDimensionError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideTableDimensionError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableDimensionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => SlideTableDimensionLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableDimensionLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideTableDimensionLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    SlideTableDimensionLimitKind::EntryBytes
                },
                _ => SlideTableDimensionLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableDimensionError::Allocation { amount: requested }
        },
        _ => SlideTableDimensionError::InvalidSource,
    }
}
