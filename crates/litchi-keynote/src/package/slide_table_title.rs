//! Exact-source transactions for existing Keynote slide-table title settings.
//!
//! This owner is deliberately narrow. It admits canonical type-6001 table
//! models whose slide, TableInfo, model, and title-style objects are uniquely
//! owned. The ownership resolver may cross physical members, while a changed
//! transaction edits only the member that owns the selected model. Legacy
//! type-6000 models remain compatibility-host scope.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "the semantic boundary redacts lower-layer failure details"
)]

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{Catalog, Entry, EntryEdit, ExactArtifacts};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{
        NestedFieldEdit, NestedFieldReplacement, WireView, patch_nested_fields_batched_with_limits,
    },
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{numbers_table_title_codec, table_info_codec};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::TableSelector;
use crate::slide::table::title::Settings;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const TITLE_VISIBLE_FIELD: u32 = 22;
const TITLE_STYLE_FIELD: u32 = 30;
const TITLE_SHAPE_STYLE_FIELD: u32 = 36;
const TITLE_OUTLINED_FIELD: u32 = 37;

/// Finite resource categories enforced by a slide-table title transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableTitleLimitKind {
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

impl fmt::Display for SlideTableTitleLimitKind {
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

/// Failure from a Keynote slide-table title read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTableTitleError {
    #[error("this Keynote source does not support physical slide-table title edits")]
    UnsupportedSource,
    #[error("the requested Keynote slide-table title graph is outside the supported scope")]
    UnsupportedDependency,
    #[error("the Keynote slide-table title selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    #[error("the selected Keynote table is locked")]
    Locked,
    #[error("the selected Keynote slide-table title source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table title {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTableTitleLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote slide-table title transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote slide-table title failed semantic verification")]
    Verification,
    #[error("the Keynote slide-table title patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy)]
struct TitleBudget {
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

impl TitleBudget {
    fn new(package: &Package) -> Result<Self, SlideTableTitleError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let source: usize = package
            .state
            .options
            .archive()
            .max_input_bytes()
            .try_into()
            .map_err(|_| SlideTableTitleError::InvalidSource)?;
        let aggregate = source
            .checked_mul(4)
            .ok_or(SlideTableTitleError::InvalidSource)?;
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
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideTableTitleLimitKind,
    ) -> Result<(), SlideTableTitleError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideTableTitleError::InvalidSource)?;
        if observed > maximum {
            return Err(SlideTableTitleError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn input(&mut self, amount: usize) -> Result<(), SlideTableTitleError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            SlideTableTitleLimitKind::InputBytes,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), SlideTableTitleError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            SlideTableTitleLimitKind::OutputBytes,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), SlideTableTitleError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            SlideTableTitleLimitKind::WireWork,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), SlideTableTitleError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideTableTitleLimitKind::References,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), SlideTableTitleError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideTableTitleLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), SlideTableTitleError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideTableTitleLimitKind::Retained,
        )
    }

    fn scratch(&mut self, amount: usize) -> Result<(), SlideTableTitleError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            SlideTableTitleLimitKind::Scratch,
        )
    }

    fn physical(&mut self, bytes: usize) -> Result<(), SlideTableTitleError> {
        self.input(bytes)?;
        self.work(bytes)
    }

    fn codec_report(
        &mut self,
        report: numbers_table_title_codec::DecodeReport,
    ) -> Result<(), SlideTableTitleError> {
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            SlideTableTitleLimitKind::WireFields,
        )?;
        self.work(report.work_bytes())?;
        self.references(report.references())?;
        self.work(report.reference_bytes())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableTitleError::LimitExceeded {
                kind: SlideTableTitleLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideTableTitleError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.work(requirements.output_bytes())
    }

    fn residual(&self, package: &Package) -> Result<WireLimits, SlideTableTitleError> {
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
            limits.with_rewrite_work(
                base.max_rewrite_work()
                    .min(self.max_work.saturating_sub(self.work).max(1)),
            )
        })
        .and_then(|limits| limits.with_nesting(base.max_nesting().min(self.max_nesting)))
        .map_err(map_wire_error)
    }
}

/// One mutable title value staged against an immutable package snapshot.
pub struct SlideTableTitleEdit<'a> {
    source: &'a Package,
    selection: TitleSelection,
    after: Settings,
}

impl fmt::Debug for SlideTableTitleEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableTitleEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("table_position", &self.selection.table_position)
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableTitleEdit<'_> {
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    #[must_use]
    pub const fn table_position(&self) -> Position {
        self.selection.table_position
    }

    #[must_use]
    pub const fn before(&self) -> Settings {
        self.selection.before
    }

    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }

    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.after = settings;
        self
    }

    pub fn commit(self) -> Result<SlideTableTitleCommit, SlideTableTitleError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible slide-table title patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableTitlePatch {
    artifacts: ExactArtifacts,
    selection: TitleSelection,
    before: Settings,
    after: Settings,
    touched_components: usize,
    deleted_previews: usize,
    source_previews_absent: bool,
    target_previews_absent: bool,
}

impl fmt::Debug for SlideTableTitlePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableTitlePatch")
            .field("slide_position", &self.selection.slide_position)
            .field("table_position", &self.selection.table_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableTitlePatch {
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.before
    }

    #[must_use]
    pub const fn after(&self) -> Settings {
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

/// Compact slide-table title publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableTitleDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideTableTitleDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
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

/// Fully verified result of one slide-table title transaction.
#[must_use = "a Keynote slide-table title commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableTitleCommit {
    package: Package,
    patch: SlideTableTitlePatch,
    diagnostics: SlideTableTitleDiagnostics,
}

impl SlideTableTitleCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &SlideTableTitlePatch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableTitleDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
struct TitleSelection {
    slide_position: Position,
    table_position: Position,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    model_message_index: usize,
    component_name: Arc<str>,
    before: Settings,
    locked: bool,
}

impl fmt::Debug for TitleSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TitleSelection")
            .field("slide_position", &self.slide_position)
            .field("table_position", &self.table_position)
            .field("before", &self.before)
            .field("locked", &self.locked)
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Read one existing slide table's lossless title visibility and outline settings.
    pub fn slide_table_title_settings<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Settings, SlideTableTitleError> {
        Ok(select_table(self, slide.into(), table.into())?.before)
    }

    /// Begin an immutable exact edit of one existing slide-table title.
    pub fn edit_slide_table_title<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTableTitleEdit<'_>, SlideTableTitleError> {
        let selection = select_table(self, slide.into(), table.into())?;
        let after = selection.before;
        Ok(SlideTableTitleEdit {
            source: self,
            selection,
            after,
        })
    }

    /// Apply an exact-source checked reversible slide-table title patch.
    pub fn apply_slide_table_title(
        &self,
        patch: &SlideTableTitlePatch,
    ) -> Result<SlideTableTitleCommit, SlideTableTitleError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideTableTitleError::PatchConflict);
        }
        if previews_absent(self)? != patch.source_previews_absent {
            return Err(SlideTableTitleError::PatchConflict);
        }
        let current = select_table(
            self,
            SlideSelector::position(patch.selection.slide_position),
            TableSelector::position(patch.selection.table_position),
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SlideTableTitleError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideTableTitleCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableTitleDiagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch)
    }
}

fn commit_edit(
    source: &Package,
    selection: &TitleSelection,
    after: Settings,
) -> Result<SlideTableTitleCommit, SlideTableTitleError> {
    if selection.before == after {
        let bytes: Arc<[u8]> = Arc::from(source.source_bytes());
        return Ok(SlideTableTitleCommit {
            package: source.snapshot(),
            patch: SlideTableTitlePatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before,
                after,
                touched_components: 0,
                deleted_previews: 0,
                source_previews_absent: previews_absent(source)?,
                target_previews_absent: previews_absent(source)?,
            },
            diagnostics: SlideTableTitleDiagnostics::unchanged(),
        });
    }
    if selection.locked {
        return Err(SlideTableTitleError::Locked);
    }
    let mut budget = TitleBudget::new(source)?;
    let (candidate, deleted_previews) = rewrite_title(source, selection, after, &mut budget)?;
    candidate.validate().map_err(map_read_error)?;
    if !previews_absent(&candidate)? {
        return Err(SlideTableTitleError::Verification);
    }
    let selected = select_table(
        &candidate,
        SlideSelector::position(selection.slide_position),
        TableSelector::position(selection.table_position),
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(SlideTableTitleError::Verification);
    }
    verify_locality(source, &candidate, selection, true, &mut budget)?;
    let target = physical_catalog(&candidate)?.shared_source();
    Ok(SlideTableTitleCommit {
        package: candidate,
        patch: SlideTableTitlePatch {
            artifacts: ExactArtifacts::new(Arc::from(source.source_bytes()), target),
            selection: selection.clone(),
            before: selection.before,
            after,
            touched_components: 1,
            deleted_previews,
            source_previews_absent: previews_absent(source)?,
            target_previews_absent: true,
        },
        diagnostics: SlideTableTitleDiagnostics::published(deleted_previews),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &SlideTableTitlePatch,
) -> Result<SlideTableTitleCommit, SlideTableTitleError> {
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    if previews_absent(&candidate)? != patch.target_previews_absent {
        return Err(SlideTableTitleError::Verification);
    }
    let selected = select_table(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        TableSelector::position(patch.selection.table_position),
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SlideTableTitleError::Verification);
    }
    let mut budget = TitleBudget::new(source)?;
    budget.input(patch.artifacts.target().len())?;
    verify_locality(
        source,
        &candidate,
        &patch.selection,
        patch.target_previews_absent,
        &mut budget,
    )?;
    Ok(SlideTableTitleCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableTitleDiagnostics::published(patch.deleted_previews),
    })
}

fn select_table(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
) -> Result<TitleSelection, SlideTableTitleError> {
    let mut budget = TitleBudget::new(package)?;
    let catalog = physical_catalog(package)?;
    budget.input(package.source_bytes().len())?;
    budget.allocations(catalog.package().len())?;
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTableTitleError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (_slide_component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideTableTitleError::InvalidSource)?;
    let (slide_message_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
    let wire_limits = budget.residual(package)?;
    let owned = repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, wire_limits)?;
    let z_order = repeated_references(slide_payload, SLIDE_Z_ORDER_FIELD, wire_limits)?;
    budget.references(owned.len().saturating_add(z_order.len()))?;
    let owned_set = checked_reference_set(&owned, &mut budget)?;
    let _z_order_set = checked_reference_set(&z_order, &mut budget)?;
    validate_slide_metadata(slide, slide_message_index, &owned, &z_order, &mut budget)?;

    let mut tables = Vec::new();
    tables
        .try_reserve_exact(z_order.len())
        .map_err(|_| SlideTableTitleError::Allocation {
            amount: z_order.len(),
        })?;
    budget.allocations(1)?;
    for identifier in z_order {
        let Some((_owner_component, object)) = package.object_with_component(identifier) else {
            return Err(SlideTableTitleError::InvalidSource);
        };
        if object
            .messages
            .iter()
            .all(|message| message.type_ != TABLE_INFO_MESSAGE_TYPE)
        {
            continue;
        }
        if !owned_set.contains(&identifier) {
            return Err(SlideTableTitleError::InvalidSource);
        }
        let (info_index, info_payload) = unique_message(object, TABLE_INFO_MESSAGE_TYPE)?;
        let info = decode_table_info(info_payload, package, &mut budget)?;
        let parent = table_parent(info_payload, wire_limits)?;
        if parent != record.slide_identifier {
            return Err(SlideTableTitleError::InvalidSource);
        }
        let model_identifier = info.table_model().identifier().get();
        validate_table_info_metadata(object, info_index, parent, model_identifier, &mut budget)?;
        let (model_component, model) = package
            .object_with_component(model_identifier)
            .ok_or(SlideTableTitleError::InvalidSource)?;
        if model
            .messages
            .iter()
            .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        {
            return Err(SlideTableTitleError::UnsupportedDependency);
        }
        let (model_message_index, model_payload) = unique_message(model, TABLE_MODEL_MESSAGE_TYPE)?;
        let before_snapshot = decode_title(model_payload, package, &mut budget)?;
        let before = Settings::new(
            before_snapshot.table_name_enabled(),
            before_snapshot.table_name_border_enabled(),
        );
        if before.is_visible() {
            validate_visible_prerequisites(
                package,
                model,
                model_message_index,
                before_snapshot,
                &mut budget,
            )?;
        }
        tables.push((
            identifier,
            model_identifier,
            Arc::<str>::from(model_component),
            model_message_index,
            before,
            info.locked().unwrap_or(false),
        ));
    }
    let table_position = table_selector.as_position();
    let (
        table_info_identifier,
        model_identifier,
        model_component_name,
        model_message_index,
        before,
        locked,
    ) = tables.get(table_position.get()).cloned().ok_or(
        SlideTableTitleError::TablePositionNotFound {
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
        wire_limits,
        &mut budget,
    )?;
    Ok(TitleSelection {
        slide_position,
        table_position,
        slide_identifier: record.slide_identifier,
        table_info_identifier,
        model_identifier,
        model_message_index,
        component_name: model_component_name,
        before,
        locked,
    })
}

fn rewrite_title(
    source: &Package,
    selection: &TitleSelection,
    after: Settings,
    budget: &mut TitleBudget,
) -> Result<(Package, usize), SlideTableTitleError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.component_name.as_ref())
        .ok_or(SlideTableTitleError::InvalidSource)?;
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
        .ok_or(SlideTableTitleError::InvalidSource)?;
    let original = object
        .messages
        .get(selection.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(SlideTableTitleError::InvalidSource)?
        .data
        .as_slice();
    let snapshot = decode_title(original, source, budget)?;
    if Settings::new(
        snapshot.table_name_enabled(),
        snapshot.table_name_border_enabled(),
    ) != selection.before
    {
        return Err(SlideTableTitleError::InvalidSource);
    }
    validate_visible_prerequisites(
        source,
        object,
        selection.model_message_index,
        snapshot,
        budget,
    )
    .or_else(|error| {
        if after.is_visible() {
            Err(error)
        } else {
            Ok(())
        }
    })?;

    let paths = [[TITLE_VISIBLE_FIELD], [TITLE_OUTLINED_FIELD]];
    let edits = [
        NestedFieldEdit::new(
            &paths[0],
            selection.before.visible().is_some(),
            NestedFieldReplacement::Varint(after.visible().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &paths[1],
            selection.before.outlined().is_some(),
            NestedFieldReplacement::Varint(after.outlined().map(u64::from)),
        ),
    ];
    let output_bound = original
        .len()
        .checked_add(64)
        .ok_or(SlideTableTitleError::InvalidSource)?;
    budget.output(output_bound)?;
    budget.work(original.len().saturating_mul(8).max(1))?;
    let limits = budget
        .residual(source)?
        .with_output_bytes(output_bound)
        .and_then(|limits| limits.with_rewrite_work(original.len().saturating_mul(8).max(1)))
        .map_err(map_wire_error)?;
    let rewritten = patch_nested_fields_batched_with_limits(original, &edits, limits)
        .map_err(map_wire_error)?;
    let verified = decode_title(&rewritten, source, budget)?;
    if Settings::new(
        verified.table_name_enabled(),
        verified.table_name_border_enabled(),
    ) != after
    {
        return Err(SlideTableTitleError::Verification);
    }
    archive
        .object_mut(selection.model_identifier)
        .ok_or(SlideTableTitleError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_core_error)?;
    budget.output(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideTableTitleError::InvalidSource)?,
    )?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if bytes.len() != encoded_bound {
        return Err(SlideTableTitleError::Verification);
    }
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(SlideTableTitleError::Verification);
    }
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| SlideTableTitleError::InvalidSource)?;
    let edits = [EntryEdit::new(
        selection.component_name.as_ref(),
        &compressed,
    )];
    let prepared = catalog
        .prepare_reassembly_with_deletions(&edits, previews.names(), physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements)?;
    budget.input(requirements.output_bytes())?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, previews.len()))
}

fn decode_title(
    payload: &[u8],
    package: &Package,
    budget: &mut TitleBudget,
) -> Result<numbers_table_title_codec::TableTitleSettingsSnapshot, SlideTableTitleError> {
    let limits = budget.residual(package)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideTableTitleError::InvalidSource)?;
    let options = numbers_table_title_codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
        budget
            .max_references
            .saturating_sub(budget.references)
            .max(1),
    );
    let (snapshot, report) =
        numbers_table_title_codec::decode_table_title_settings_with_report(payload, options)
            .map_err(map_title_codec_error)?;
    budget.codec_report(report)?;
    Ok(snapshot)
}

fn decode_table_info(
    payload: &[u8],
    package: &Package,
    budget: &mut TitleBudget,
) -> Result<table_info_codec::TableInfoSnapshot, SlideTableTitleError> {
    let limits = budget.residual(package)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideTableTitleError::InvalidSource)?;
    budget.work(
        payload
            .len()
            .checked_mul(4)
            .ok_or(SlideTableTitleError::InvalidSource)?,
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
    .map_err(|_| SlideTableTitleError::InvalidSource)
}

fn validate_visible_prerequisites(
    package: &Package,
    model: &ArchiveObject,
    message_index: usize,
    snapshot: numbers_table_title_codec::TableTitleSettingsSnapshot,
    budget: &mut TitleBudget,
) -> Result<(), SlideTableTitleError> {
    let height = f64::from_bits(
        snapshot
            .table_name_height_bits()
            .ok_or(SlideTableTitleError::InvalidSource)?,
    );
    if !height.is_finite() || height < 0.0 {
        return Err(SlideTableTitleError::InvalidSource);
    }
    let style = snapshot
        .table_name_style()
        .ok_or(SlideTableTitleError::InvalidSource)?;
    let shape = snapshot
        .table_name_shape_style()
        .ok_or(SlideTableTitleError::InvalidSource)?;
    if style.identifier() == 0
        || shape.identifier() == 0
        || style.identifier() == shape.identifier()
        || style.deprecated_type().is_some()
        || shape.deprecated_type().is_some()
        || style.deprecated_is_external().is_some()
        || shape.deprecated_is_external().is_some()
    {
        return Err(SlideTableTitleError::UnsupportedDependency);
    }
    validate_model_style_metadata(
        model,
        message_index,
        style.identifier(),
        shape.identifier(),
        budget,
    )?;
    require_style_object(
        package,
        style.identifier(),
        PARAGRAPH_STYLE_MESSAGE_TYPE,
        budget,
    )?;
    require_style_object(
        package,
        shape.identifier(),
        SHAPE_STYLE_MESSAGE_TYPE,
        budget,
    )
}

fn require_style_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
    budget: &mut TitleBudget,
) -> Result<(), SlideTableTitleError> {
    let mut matched = None;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            budget.work(1)?;
            if object.archive_info.identifier == Some(identifier) {
                if matched.replace((component.name(), object)).is_some() {
                    return Err(SlideTableTitleError::UnsupportedDependency);
                }
            }
        }
    }
    let (_owner_component, object) = matched.ok_or(SlideTableTitleError::InvalidSource)?;
    if object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
        || object.messages[0].type_ != message_type
    {
        return Err(SlideTableTitleError::UnsupportedDependency);
    }
    validate_message_header(object, 0)
}

fn validate_model_style_metadata(
    object: &ArchiveObject,
    message_index: usize,
    style: u64,
    shape: u64,
    budget: &mut TitleBudget,
) -> Result<(), SlideTableTitleError> {
    validate_message_header(object, message_index)?;
    let info = &object.archive_info.message_infos[message_index];
    if !info.object_references.is_empty() {
        let frequencies = reference_frequencies(&info.object_references, budget)?;
        if frequencies.get(&style).copied() != Some(1)
            || frequencies.get(&shape).copied() != Some(1)
        {
            return Err(SlideTableTitleError::InvalidSource);
        }
    }
    for (field_number, identifier) in [(TITLE_STYLE_FIELD, style), (TITLE_SHAPE_STYLE_FIELD, shape)]
    {
        let fields = info
            .field_infos
            .iter()
            .filter(|field| field.path.as_slice() == [field_number])
            .collect::<Vec<_>>();
        if fields.len() > 1
            || fields.first().is_some_and(|field| {
                field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
                    || !field.data_references.is_empty()
                    || field.object_references.as_slice() != [identifier]
            })
        {
            return Err(SlideTableTitleError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_slide_metadata(
    object: &ArchiveObject,
    message_index: usize,
    owned: &[u64],
    z_order: &[u64],
    budget: &mut TitleBudget,
) -> Result<(), SlideTableTitleError> {
    validate_message_header(object, message_index)?;
    let info = &object.archive_info.message_infos[message_index];
    if !info.object_references.is_empty() {
        let frequencies = reference_frequencies(&info.object_references, budget)?;
        if frequencies.len() != info.object_references.len() {
            return Err(SlideTableTitleError::InvalidSource);
        }
        for identifier in owned.iter().chain(z_order) {
            if frequencies.get(identifier).copied() != Some(1) {
                return Err(SlideTableTitleError::InvalidSource);
            }
        }
    }
    for field in &info.field_infos {
        if field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD]
            && field.object_references.as_slice() != owned
        {
            return Err(SlideTableTitleError::InvalidSource);
        }
        if field.path.as_slice() == [SLIDE_Z_ORDER_FIELD]
            && field.object_references.as_slice() != z_order
        {
            return Err(SlideTableTitleError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_table_info_metadata(
    object: &ArchiveObject,
    message_index: usize,
    parent: u64,
    model: u64,
    budget: &mut TitleBudget,
) -> Result<(), SlideTableTitleError> {
    validate_message_header(object, message_index)?;
    let info = &object.archive_info.message_infos[message_index];
    if !info.object_references.is_empty() {
        let frequencies = reference_frequencies(&info.object_references, budget)?;
        if frequencies.get(&parent).copied() != Some(1)
            || frequencies.get(&model).copied() != Some(1)
        {
            return Err(SlideTableTitleError::InvalidSource);
        }
    }
    for field in &info.field_infos {
        if field.path.as_slice() == [TABLE_MODEL_FIELD]
            && field.object_references.as_slice() != [model]
        {
            return Err(SlideTableTitleError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_message_header(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), SlideTableTitleError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableTitleError::InvalidSource);
    }
    let message = object
        .messages
        .get(message_index)
        .ok_or(SlideTableTitleError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableTitleError::InvalidSource)?;
    if message.type_ != info.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SlideTableTitleError::InvalidSource);
    }
    Ok(())
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), SlideTableTitleError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableTitleError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        validate_message_header(object, index)?;
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideTableTitleError::InvalidSource);
        }
    }
    selected.ok_or(SlideTableTitleError::InvalidSource)
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
) -> Result<Vec<u64>, SlideTableTitleError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(
            fields
                .fields()
                .filter(|field| field.number() == field_number)
                .count(),
        )
        .map_err(|_| SlideTableTitleError::Allocation {
            amount: payload.len(),
        })?;
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideTableTitleError::InvalidSource);
        }
        result.push(strict_reference(field.payload(), limits)?);
    }
    Ok(result)
}

fn table_parent(payload: &[u8], limits: WireLimits) -> Result<u64, SlideTableTitleError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let super_fields = fields
        .fields()
        .filter(|field| field.number() == TABLE_SUPER_FIELD)
        .collect::<Vec<_>>();
    if super_fields.len() != 1 || super_fields[0].wire_type() != 2 {
        return Err(SlideTableTitleError::InvalidSource);
    }
    let drawable =
        WireView::parse_with_limits(super_fields[0].payload(), limits).map_err(map_wire_error)?;
    let parents = drawable
        .fields()
        .filter(|field| field.number() == DRAWABLE_PARENT_FIELD)
        .collect::<Vec<_>>();
    if parents.len() != 1 || parents[0].wire_type() != 2 {
        return Err(SlideTableTitleError::InvalidSource);
    }
    strict_reference(parents[0].payload(), limits)
}

fn strict_reference(payload: &[u8], limits: WireLimits) -> Result<u64, SlideTableTitleError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideTableTitleError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideTableTitleError::InvalidSource)?;
                if value == 0 || width != encoded_len(value) {
                    return Err(SlideTableTitleError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideTableTitleError::UnsupportedDependency),
            _ => {},
        }
    }
    identifier.ok_or(SlideTableTitleError::InvalidSource)
}

fn checked_reference_set(
    values: &[u64],
    budget: &mut TitleBudget,
) -> Result<HashSet<u64>, SlideTableTitleError> {
    budget.allocations(usize::from(!values.is_empty()))?;
    budget.retained(
        values
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableTitleError::InvalidSource)?,
    )?;
    let mut set = HashSet::new();
    set.try_reserve(values.len())
        .map_err(|_| SlideTableTitleError::Allocation {
            amount: values.len(),
        })?;
    for value in values {
        if !set.insert(*value) {
            return Err(SlideTableTitleError::InvalidSource);
        }
    }
    Ok(set)
}

fn reference_frequencies(
    values: &[u64],
    budget: &mut TitleBudget,
) -> Result<HashMap<u64, usize>, SlideTableTitleError> {
    budget.allocations(usize::from(!values.is_empty()))?;
    budget.retained(
        values
            .len()
            .checked_mul(size_of::<(u64, usize)>())
            .ok_or(SlideTableTitleError::InvalidSource)?,
    )?;
    let mut frequencies = HashMap::new();
    frequencies
        .try_reserve(values.len())
        .map_err(|_| SlideTableTitleError::Allocation {
            amount: values.len(),
        })?;
    for value in values {
        let count = frequencies.entry(*value).or_insert(0usize);
        *count = count
            .checked_add(1)
            .ok_or(SlideTableTitleError::InvalidSource)?;
    }
    Ok(frequencies)
}

fn ensure_unique_identity(package: &Package, identifier: u64) -> Result<(), SlideTableTitleError> {
    let mut total = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.archive_info.identifier == Some(identifier) {
                total += 1;
            }
        }
    }
    if total != 1 {
        return Err(SlideTableTitleError::UnsupportedDependency);
    }
    Ok(())
}

fn ensure_unique_table_owner(
    package: &Package,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    limits: WireLimits,
    budget: &mut TitleBudget,
) -> Result<(), SlideTableTitleError> {
    let mut owned_count = 0usize;
    let mut z_order_count = 0usize;
    let mut selected_slide = false;
    let mut model_owners = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ == SLIDE_MESSAGE_TYPE {
                    let owned =
                        repeated_references(&message.data, SLIDE_OWNED_DRAWABLES_FIELD, limits)?;
                    let z_order = repeated_references(&message.data, SLIDE_Z_ORDER_FIELD, limits)?;
                    let owned_hits = owned.iter().fold(0usize, |count, identifier| {
                        count + usize::from(*identifier == table_info_identifier)
                    });
                    let z_hits = z_order.iter().fold(0usize, |count, identifier| {
                        count + usize::from(*identifier == table_info_identifier)
                    });
                    owned_count = owned_count
                        .checked_add(owned_hits)
                        .ok_or(SlideTableTitleError::InvalidSource)?;
                    z_order_count = z_order_count
                        .checked_add(z_hits)
                        .ok_or(SlideTableTitleError::InvalidSource)?;
                    budget.work(
                        owned
                            .len()
                            .checked_add(z_order.len())
                            .ok_or(SlideTableTitleError::InvalidSource)?,
                    )?;
                    if object.archive_info.identifier == Some(slide_identifier)
                        && owned_hits == 1
                        && z_hits == 1
                    {
                        selected_slide = true;
                    }
                }
                if message.type_ == TABLE_INFO_MESSAGE_TYPE {
                    if let Ok(info) = table_info_codec::decode_table_info(
                        &message.data,
                        table_info_codec::DecodeOptions::new(
                            message.data.len().max(1),
                            limits.max_fields(),
                            limits.max_rewrite_work(),
                            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
                        ),
                    ) {
                        if info.table_model().identifier().get() == model_identifier {
                            model_owners += 1;
                        }
                    }
                }
            }
        }
    }
    if owned_count != 1 || z_order_count != 1 || !selected_slide || model_owners != 1 {
        return Err(SlideTableTitleError::UnsupportedDependency);
    }
    Ok(())
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &TitleSelection,
    target_previews_absent: bool,
    budget: &mut TitleBudget,
) -> Result<(), SlideTableTitleError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let previews = super::rendering_invalidation::root_preview_deletions(source_catalog.package())
        .map_err(|_| SlideTableTitleError::Verification)?;
    let source_entries = entry_index(source_catalog.package(), budget)?;
    let candidate_entries = entry_index(candidate_catalog.package(), budget)?;
    for entry in source_catalog.package().iter() {
        budget.work(entry.data().len())?;
        let candidate_entry = candidate_entries.get(entry.name()).copied();
        if target_previews_absent && previews.names().contains(&entry.name()) {
            if candidate_entry.is_some() {
                return Err(SlideTableTitleError::Verification);
            }
            continue;
        }
        let other = candidate_entry.ok_or(SlideTableTitleError::Verification)?;
        if entry.name() != selection.component_name.as_ref()
            && (entry.data() != other.data() || entry.metadata() != other.metadata())
        {
            return Err(SlideTableTitleError::Verification);
        }
    }
    for entry in candidate_catalog.package().iter() {
        if !source_entries.contains_key(entry.name()) {
            return Err(SlideTableTitleError::Verification);
        }
    }
    let source_archive = component_archive(source, selection.component_name.as_ref())?;
    let candidate_archive = component_archive(candidate, selection.component_name.as_ref())?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SlideTableTitleError::Verification);
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
            .ok_or(SlideTableTitleError::Verification)?;
        let candidate_object = candidate_archive
            .object(identifier)
            .ok_or(SlideTableTitleError::Verification)?;
        if identifier == selection.model_identifier {
            let candidate_message = candidate_object
                .messages
                .get(selection.model_message_index)
                .ok_or(SlideTableTitleError::Verification)?;
            let mut expected = source_object.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    selection.model_message_index,
                    candidate_message.clone(),
                    archive_limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = candidate_object.header_length;
            expected.data_length = candidate_object.data_length;
            if !expected.same_content_ignoring_offsets(candidate_object) {
                return Err(SlideTableTitleError::Verification);
            }
        } else if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(SlideTableTitleError::Verification);
        }
    }
    Ok(())
}

/// Index physical member names once for a locality verification pass.
///
/// Duplicate names cannot be assigned an unambiguous provenance.  The index
/// therefore rejects them before any candidate bytes are compared.
fn entry_index<'a>(
    catalog: &'a Catalog,
    budget: &mut TitleBudget,
) -> Result<HashMap<&'a str, &'a Entry>, SlideTableTitleError> {
    let count = catalog.iter().count();
    budget.allocations(usize::from(count != 0))?;
    budget.retained(
        count
            .checked_mul(size_of::<(&str, &Entry)>())
            .ok_or(SlideTableTitleError::InvalidSource)?,
    )?;
    let mut index = HashMap::new();
    index
        .try_reserve(count)
        .map_err(|_| SlideTableTitleError::Allocation { amount: count })?;
    for entry in catalog.iter() {
        budget.work(
            entry
                .name()
                .len()
                .checked_add(1)
                .ok_or(SlideTableTitleError::InvalidSource)?,
        )?;
        if index.insert(entry.name(), entry).is_some() {
            return Err(SlideTableTitleError::Verification);
        }
    }
    Ok(index)
}

fn component_archive(package: &Package, name: &str) -> Result<Archive, SlideTableTitleError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(SlideTableTitleError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableTitleError::InvalidSource);
    }
    let snappy = package
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream =
        SnappyStream::decompress_with_limits(entry.data(), snappy).map_err(map_core_error)?;
    Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideTableTitleError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideTableTitleError::EmptySlideName);
            }
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(|_| SlideTableTitleError::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideTableTitleError::SlideNameNotFound)
        },
    }
}

fn same_selection(left: &TitleSelection, right: &TitleSelection) -> bool {
    left.slide_position == right.slide_position
        && left.table_position == right.table_position
        && left.slide_identifier == right.slide_identifier
        && left.table_info_identifier == right.table_info_identifier
        && left.model_identifier == right.model_identifier
        && left.model_message_index == right.model_message_index
        && left.component_name == right.component_name
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideTableTitleError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTableTitleError::UnsupportedSource),
    }
}

fn previews_absent(package: &Package) -> Result<bool, SlideTableTitleError> {
    super::rendering_invalidation::root_previews_absent(physical_catalog(package)?.package())
        .map_err(|_| SlideTableTitleError::Verification)
}

fn map_read_error(error: ReadError) -> SlideTableTitleError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableTitleError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableTitleLimitKind::References,
                _ => SlideTableTitleLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableTitleError::Allocation { amount },
        _ => SlideTableTitleError::InvalidSource,
    }
}

fn map_title_codec_error(error: numbers_table_title_codec::DecodeError) -> SlideTableTitleError {
    let Some(limit) = error.resource_limit() else {
        return SlideTableTitleError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        numbers_table_title_codec::DecodeLimit::Bytes { observed, maximum } => (
            SlideTableTitleLimitKind::InputBytes,
            observed as u64,
            maximum as u64,
        ),
        numbers_table_title_codec::DecodeLimit::Fields { observed, maximum } => (
            SlideTableTitleLimitKind::WireFields,
            observed as u64,
            maximum as u64,
        ),
        numbers_table_title_codec::DecodeLimit::Work { observed, maximum } => (
            SlideTableTitleLimitKind::WireWork,
            observed as u64,
            maximum as u64,
        ),
        numbers_table_title_codec::DecodeLimit::Nesting { observed, maximum } => (
            SlideTableTitleLimitKind::WireNesting,
            u64::from(observed),
            u64::from(maximum),
        ),
        numbers_table_title_codec::DecodeLimit::References { observed, maximum } => (
            SlideTableTitleLimitKind::References,
            observed as u64,
            maximum as u64,
        ),
        _ => return SlideTableTitleError::InvalidSource,
    };
    SlideTableTitleError::LimitExceeded {
        kind,
        observed,
        maximum,
    }
}

fn map_wire_error(_error: litchi_iwa_common::Error) -> SlideTableTitleError {
    SlideTableTitleError::InvalidSource
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTableTitleError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideTableTitleLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SlideTableTitleLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SlideTableTitleLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideTableTitleLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => SlideTableTitleLimitKind::TotalBytes,
                _ => SlideTableTitleLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableTitleError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTableTitleError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideTableTitleError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => SlideTableTitleLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableTitleLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => SlideTableTitleLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => SlideTableTitleLimitKind::EntryBytes,
                _ => SlideTableTitleLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableTitleError::Allocation { amount: requested }
        },
        _ => SlideTableTitleError::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn unbounded_budget() -> TitleBudget {
        TitleBudget {
            max_input: usize::MAX,
            max_output: usize::MAX,
            max_fields: usize::MAX,
            max_work: usize::MAX,
            max_nesting: usize::MAX,
            max_references: usize::MAX,
            max_allocations: usize::MAX,
            max_retained: usize::MAX,
            max_scratch: usize::MAX,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
        }
    }

    #[test]
    fn reference_index_rejects_duplicate_routes() {
        let mut budget = unbounded_budget();
        assert!(matches!(
            checked_reference_set(&[11, 22, 11], &mut budget),
            Err(SlideTableTitleError::InvalidSource)
        ));
    }

    #[test]
    fn reference_frequency_index_counts_once_per_route() {
        let mut budget = unbounded_budget();
        let frequencies = reference_frequencies(&[11, 22, 11], &mut budget).unwrap();
        assert_eq!(frequencies.get(&11), Some(&2));
        assert_eq!(frequencies.get(&22), Some(&1));
    }
}
