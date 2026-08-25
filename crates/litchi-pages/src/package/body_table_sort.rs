//! Exact-source ownership of a rooted Pages table's persisted sort order.
//!
//! This transaction changes only `TableModelArchive.sort_order` (field 44).
//! Field 45 and every other model field remain opaque source bytes.  Rooted
//! table selection and lock/ownership proof are delegated to `table_lock`;
//! the strict source-preserving field projection is provided by the hidden
//! `table_sort_order_codec`.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused transaction keeps its proof beside the rewrite."
)]

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_archive::{Error as ArchiveError, LimitKind as ArchiveLimitKind, SourceCatalog};
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len};
use litchi_iwa_core::{Error as CoreError, LimitKind as CoreLimitKind, RawMessage};
use litchi_iwa_protos::table_sort_order_codec as codec;
use thiserror::Error;

use super::{Package, PackageError, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_MODEL_COLUMNS_FIELD: u32 = 7;
const SORT_TRACKER_FIELD: u32 = 45;
const SORT_TRACKER_REFERENCE_FIELD: u32 = 1;
const SORT_TRACKER_REFERENCE_TYPE_FIELD: u32 = 2;
const SORT_TRACKER_REFERENCE_EXTERNAL_FIELD: u32 = 3;

/// Content-free location associated with a body-table sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableSortPath {
    /// The complete Pages package.
    Package,
    /// One rooted body table at a checked zero-based position.
    Table { table: usize },
}

/// Finite resources governed by one Pages persisted-sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableSortLimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalEntryBytes,
    PackageBytes,
    PayloadBytes,
    TotalPayloadBytes,
    PayloadObjects,
    PayloadMessages,
    PayloadItems,
    PayloadReferences,
    WireBytes,
    WireOutputBytes,
    WireFields,
    WireNesting,
    WireWork,
    WireRules,
    WireColumns,
    WireAllocations,
    WireRetainedBytes,
    WireScratchBytes,
}

impl fmt::Display for BodyTableSortLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "ZIP entries",
            Self::EntryBytes => "ZIP entry bytes",
            Self::TotalEntryBytes => "total ZIP entry bytes",
            Self::PackageBytes => "package bytes",
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
            Self::WireRules => "wire rules",
            Self::WireColumns => "wire columns",
            Self::WireAllocations => "wire allocations",
            Self::WireRetainedBytes => "wire retained bytes",
            Self::WireScratchBytes => "wire scratch bytes",
        })
    }
}

/// Failure from a Pages body-table persisted-sort read or transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum BodyTableSortError {
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    #[error("the Pages body-table sort selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Pages package source does not support exact body-table sort editing")]
    UnsupportedSource,
    #[error("the selected Pages body-table sort source is invalid")]
    InvalidSource,
    #[error("the selected Pages body table is locked")]
    TableLocked,
    #[error("Pages body-table sort {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        kind: BodyTableSortLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for Pages body-table sort")]
    Allocation { amount: usize },
    #[error("the edited Pages body-table sort failed semantic verification")]
    Verification,
    #[error("the Pages body-table sort patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Clone, PartialEq, Eq)]
struct Target {
    native: table_lock::BodyTableTarget,
    columns: u32,
    before: Option<Order>,
}

/// A mutable semantic persisted-sort edit.
pub struct BodyTableSortEdit<'a> {
    source: &'a Package,
    target: Target,
    after: Option<Order>,
}

impl fmt::Debug for BodyTableSortEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableSortEdit")
            .field("table_position", &self.target.native.table_position)
            .field("before", &self.target.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyTableSortEdit<'_> {
    /// Return the selected table path without exposing native identifiers.
    #[must_use]
    pub const fn path(&self) -> BodyTableSortPath {
        BodyTableSortPath::Table {
            table: self.target.native.table_position,
        }
    }

    /// Return the staged order, or `None` when the native field is absent or cleared.
    #[must_use]
    pub fn order(&self) -> Option<&Order> {
        self.after.as_ref()
    }

    /// Stage a persisted sort order without moving table rows.
    #[must_use]
    pub fn set(mut self, order: Order) -> Self {
        self.after = Some(order);
        self
    }

    /// Remove the persisted sort configuration.
    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self
    }

    /// Alias for [`Self::clear`].
    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    /// Validate and publish this edit atomically.
    pub fn commit(self) -> Result<BodyTableSortCommit, BodyTableSortError> {
        commit_edit(self)
    }
}

/// Reversible exact-source patch for one body-table sort field.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyTableSortPatch {
    artifacts: OwnedExactArtifacts,
    target: Target,
    before: Option<Order>,
    after: Option<Order>,
}

impl fmt::Debug for BodyTableSortPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableSortPatch")
            .field("table_position", &self.target.native.table_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyTableSortPatch {
    /// Return the selected table path without exposing native identifiers.
    #[must_use]
    pub const fn path(&self) -> BodyTableSortPath {
        BodyTableSortPath::Table {
            table: self.target.native.table_position,
        }
    }

    /// Return the source semantic order.
    #[must_use]
    pub fn before(&self) -> Option<&Order> {
        self.before.as_ref()
    }

    /// Return the target semantic order.
    #[must_use]
    pub fn after(&self) -> Option<&Order> {
        self.after.as_ref()
    }

    /// Return the source fingerprint used for conflict detection.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target fingerprint used for conflict detection.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Whether this patch leaves both semantic state and exact bytes unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            target: self.target.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Content-free transaction diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyTableSortDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl BodyTableSortDiagnostics {
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

/// Fully reopened result of one body-table sort transaction.
#[must_use = "a body-table sort commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyTableSortCommit {
    package: Package,
    patch: BodyTableSortPatch,
    diagnostics: BodyTableSortDiagnostics,
}

impl BodyTableSortCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &BodyTableSortPatch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableSortDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body table's persisted sort configuration.
    pub fn body_table_sort_order<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Option<Order>, BodyTableSortError> {
        let mut budget = transaction_budget(self)?;
        let target = resolve_target(self, selector.into(), &mut budget)?;
        Ok(target.before)
    }

    /// Start a selector-first persisted-sort edit.
    pub fn edit_body_table_sort_order<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<BodyTableSortEdit<'_>, BodyTableSortError> {
        let mut budget = transaction_budget(self)?;
        let target = resolve_target(self, selector.into(), &mut budget)?;
        Ok(BodyTableSortEdit {
            source: self,
            after: target.before.clone(),
            target,
        })
    }

    /// Apply a reversible exact-source persisted-sort patch.
    pub fn apply_body_table_sort_order(
        &self,
        patch: &BodyTableSortPatch,
    ) -> Result<BodyTableSortCommit, BodyTableSortError> {
        let mut budget = transaction_budget(self)?;
        let current = self.state.source.source_bytes();
        budget
            .charge_payload_work(current.len())
            .map_err(map_lock_error)?;
        if fingerprint(current) != patch.source_fingerprint()
            || current != patch.artifacts.source_owner().as_ref()
        {
            return Err(BodyTableSortError::PatchConflict);
        }
        let selected = resolve_target(
            self,
            BodyTableSelector::index(patch.target.native.table_position),
            &mut budget,
        )?;
        if selected.native.model_identifier != patch.target.native.model_identifier
            || selected.before != patch.before
        {
            return Err(BodyTableSortError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyTableSortCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyTableSortDiagnostics::unchanged(),
            });
        }
        let target_owner = patch.artifacts.target_owner();
        budget
            .charge_payload_work(target_owner.as_ref().len())
            .map_err(map_lock_error)?;
        if !self.state.source.source_is_exact()
            || fingerprint(target_owner.as_ref()) != patch.target_fingerprint()
        {
            return Err(BodyTableSortError::PatchConflict);
        }
        let candidate = reopen_candidate(self, &target_owner, &mut budget)?;
        let verified = resolve_target(
            &candidate,
            BodyTableSelector::index(patch.target.native.table_position),
            &mut budget,
        )?;
        if verified.native.model_identifier != patch.target.native.model_identifier
            || verified.before != patch.after
        {
            return Err(BodyTableSortError::Verification);
        }
        verify_locality(self, &candidate, &patch.target.native, &mut budget)?;
        Ok(BodyTableSortCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableSortDiagnostics::published(),
        })
    }
}

fn resolve_target(
    package: &Package,
    selector: BodyTableSelector<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<Target, BodyTableSortError> {
    let native = package
        .resolve_body_table_with_budget(selector, budget)
        .map_err(map_lock_error)?;
    if native.model_message_type != TABLE_MODEL_MESSAGE_TYPE {
        return Err(BodyTableSortError::InvalidSource);
    }
    table_lock::validate_body_table_target(package, &native, budget).map_err(map_lock_error)?;
    let message = model_message(package, &native)?;
    validate_sort_tracker(package, &native, message, budget)?;
    validate_model_ownership(package, &native, budget)?;
    let columns = model_columns(&message.data, budget)?;
    let before = decode_sort(&message.data, columns, budget)?;
    Ok(Target {
        native,
        columns,
        before,
    })
}

fn transaction_budget(package: &Package) -> Result<table_lock::WireBudget, BodyTableSortError> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    Ok(budget)
}

fn commit_edit(edit: BodyTableSortEdit<'_>) -> Result<BodyTableSortCommit, BodyTableSortError> {
    let source = edit.source;
    if edit.target.before == edit.after {
        return Ok(BodyTableSortCommit {
            package: source.snapshot(),
            patch: BodyTableSortPatch {
                artifacts: OwnedExactArtifacts::new(
                    SharedBytes::from_shared_slice(source.state.source.shared_source()),
                    SharedBytes::from_shared_slice(source.state.source.shared_source()),
                ),
                target: edit.target,
                before: edit.after.clone(),
                after: edit.after,
            },
            diagnostics: BodyTableSortDiagnostics::unchanged(),
        });
    }
    if edit.target.native.explicit_locked == Some(true) {
        return Err(BodyTableSortError::TableLocked);
    }
    if let Some(order) = &edit.after {
        validate_order_columns(order, edit.target.columns)?;
    }
    if !source.state.source.source_is_exact() {
        return Err(BodyTableSortError::UnsupportedSource);
    }
    let mut budget =
        table_lock::WireBudget::new(source.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&source.state.source)
        .map_err(map_lock_error)?;
    table_lock::validate_body_table_target(source, &edit.target.native, &mut budget)
        .map_err(map_lock_error)?;
    let package = rewrite_model(source, &edit.target, edit.after.clone(), &mut budget)?;
    let candidate_target = resolve_target(
        &package,
        BodyTableSelector::index(edit.target.native.table_position),
        &mut budget,
    )?;
    if candidate_target.native.model_identifier != edit.target.native.model_identifier
        || candidate_target.before != edit.after
    {
        return Err(BodyTableSortError::Verification);
    }
    verify_locality(source, &package, &edit.target.native, &mut budget)?;
    let package_source = SharedBytes::from_shared_slice(package.state.source.shared_source());
    Ok(BodyTableSortCommit {
        package,
        patch: BodyTableSortPatch {
            artifacts: OwnedExactArtifacts::new(
                SharedBytes::from_shared_slice(source.state.source.shared_source()),
                package_source,
            ),
            target: edit.target.clone(),
            before: edit.target.before,
            after: edit.after,
        },
        diagnostics: BodyTableSortDiagnostics::published(),
    })
}

fn model_message<'a>(
    package: &'a Package,
    target: &table_lock::BodyTableTarget,
) -> Result<&'a RawMessage, BodyTableSortError> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableSortError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableSortError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableSortError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(BodyTableSortError::InvalidSource)?;
    if object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .is_none()
    {
        return Err(BodyTableSortError::InvalidSource);
    }
    Ok(message)
}

fn validate_sort_tracker(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    model: &RawMessage,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableSortError> {
    let model_view = budget.parse(&model.data, 1).map_err(map_lock_error)?;
    let mut tracker_payload = None;
    for field in model_view
        .fields()
        .filter(|field| field.number() == SORT_TRACKER_FIELD)
    {
        if field.wire_type() != 2 || tracker_payload.is_some() {
            return Err(invalid_source());
        }
        field
            .validate_canonical_framing()
            .map_err(|_| invalid_source())?;
        tracker_payload = Some(field.payload());
    }
    let Some(tracker_payload) = tracker_payload else {
        return Ok(());
    };
    let tracker_view = budget.parse(tracker_payload, 2).map_err(map_lock_error)?;
    let mut target_identifier = None;
    for field in tracker_view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| invalid_source())?;
        if field.number() == SORT_TRACKER_REFERENCE_FIELD {
            if field.wire_type() != 2 || target_identifier.is_some() {
                return Err(invalid_source());
            }
            let reference_view = budget.parse(field.payload(), 3).map_err(map_lock_error)?;
            let mut identifier = None;
            let mut reference_type = None;
            let mut external = None;
            for reference_field in reference_view.fields() {
                reference_field
                    .validate_canonical_framing()
                    .map_err(|_| invalid_source())?;
                match reference_field.number() {
                    1 => {
                        if reference_field.wire_type() != 0 || identifier.is_some() {
                            return Err(invalid_source());
                        }
                        identifier = Some(decode_nonzero_varint(reference_field.payload())?);
                    },
                    SORT_TRACKER_REFERENCE_TYPE_FIELD => {
                        if reference_field.wire_type() != 0 || reference_type.is_some() {
                            return Err(invalid_source());
                        }
                        reference_type = Some(decode_canonical_varint(reference_field.payload())?);
                    },
                    SORT_TRACKER_REFERENCE_EXTERNAL_FIELD => {
                        if reference_field.wire_type() != 0 || external.is_some() {
                            return Err(invalid_source());
                        }
                        external = Some(decode_canonical_varint(reference_field.payload())?);
                    },
                    _ => {},
                }
            }
            if reference_type.is_some_and(|value| value != 0)
                || external.is_some_and(|value| value != 0)
            {
                return Err(invalid_source());
            }
            target_identifier = Some(identifier.ok_or_else(invalid_source)?);
        }
    }
    let identifier = target_identifier.ok_or_else(invalid_source)?;
    let model_info = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .and_then(|component| component.archive().objects.get(target.model_object_index))
        .and_then(|object| {
            object
                .archive_info
                .message_infos
                .get(target.model_message_index)
        })
        .ok_or(BodyTableSortError::InvalidSource)?;
    if model_info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier.get())
        .count()
        != 1
        || model_info.data_references.contains(&identifier.get())
    {
        return Err(invalid_source());
    }
    let mut field_declaration_count = 0usize;
    for field_info in &model_info.field_infos {
        let occurrences = field_info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier.get())
            .count();
        if occurrences == 0 {
            continue;
        }
        if occurrences != 1 || field_info.path.as_slice() != [SORT_TRACKER_FIELD, 1] {
            return Err(invalid_source());
        }
        field_declaration_count = field_declaration_count
            .checked_add(1)
            .ok_or_else(invalid_source)?;
    }
    // Pages 14.4's native table-model producer records this tracker in the
    // aggregate object-reference list but omits the corresponding FieldInfo.
    // Accept that producer form only when the aggregate is the exact singleton
    // proved above; if a FieldInfo is present it must be the canonical path.
    if field_declaration_count > 1 {
        return Err(invalid_source());
    }
    let mut physical = None;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for object in &component.archive().objects {
            if object.archive_info.identifier == Some(identifier.get()) {
                if physical.replace(component_index).is_some() {
                    return Err(invalid_source());
                }
            }
        }
    }
    if physical != Some(target.model_component_index) {
        return Err(invalid_source());
    }
    Ok(())
}

fn validate_model_ownership(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableSortError> {
    let mut physical = None;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            budget.charge_payload_work(1).map_err(map_lock_error)?;
            if object.archive_info.identifier == Some(target.model_identifier.get()) {
                if physical.replace((component_index, object_index)).is_some() {
                    return Err(invalid_source());
                }
            }
        }
    }
    if physical != Some((target.model_component_index, target.model_object_index)) {
        return Err(invalid_source());
    }
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            for (message_index, info) in object.archive_info.message_infos.iter().enumerate() {
                let aggregate_count = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == target.model_identifier.get())
                    .count();
                let data_count = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == target.model_identifier.get())
                    .count();
                let field_count = info
                    .field_infos
                    .iter()
                    .map(|field| {
                        field
                            .object_references
                            .iter()
                            .filter(|identifier| **identifier == target.model_identifier.get())
                            .count()
                    })
                    .sum::<usize>();
                if aggregate_count == 0 && data_count == 0 && field_count == 0 {
                    continue;
                }
                let selected_info = component_index == target.component_index
                    && object_index == target.object_index
                    && message_index == target.info_message_index;
                let valid_selected_field_info = field_count == 0
                    || (field_count == 1
                        && info.field_infos.iter().all(|field| {
                            !field
                                .object_references
                                .contains(&target.model_identifier.get())
                                || field.path.as_slice() == [2]
                        }));
                if !selected_info
                    || aggregate_count != 1
                    || data_count != 0
                    || !valid_selected_field_info
                {
                    return Err(invalid_source());
                }
            }
        }
    }
    Ok(())
}

fn decode_nonzero_varint(source: &[u8]) -> Result<std::num::NonZeroU64, BodyTableSortError> {
    std::num::NonZeroU64::new(decode_canonical_varint(source)?).ok_or_else(invalid_source)
}

fn decode_canonical_varint(source: &[u8]) -> Result<u64, BodyTableSortError> {
    let (value, length) = decode_varint_from_bytes(source).map_err(|_| invalid_source())?;
    if length != source.len() || encoded_len(value) != length {
        return Err(invalid_source());
    }
    Ok(value)
}

fn model_columns(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<u32, BodyTableSortError> {
    let view = budget.parse(source, 1).map_err(map_lock_error)?;
    budget
        .charge_payload_work(view.len())
        .map_err(map_lock_error)?;
    let mut columns = None;
    for field in view
        .fields()
        .filter(|field| field.number() == TABLE_MODEL_COLUMNS_FIELD)
    {
        if field.wire_type() != 0 || columns.is_some() {
            return Err(invalid_source());
        }
        field
            .validate_canonical_framing()
            .map_err(|_| invalid_source())?;
        let (value, length) =
            decode_varint_from_bytes(field.payload()).map_err(|_| invalid_source())?;
        if length != field.payload().len() {
            return Err(invalid_source());
        }
        columns = Some(u32::try_from(value).map_err(|_| invalid_source())?);
    }
    columns
        .filter(|columns| *columns != 0)
        .ok_or_else(invalid_source)
}

fn decode_sort(
    source: &[u8],
    columns: u32,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<Order>, BodyTableSortError> {
    let options = codec_options(budget, source.len(), columns)?;
    let (snapshot, report) =
        codec::decode_table_sort_order_with_report(source, options).map_err(map_codec_error)?;
    budget
        .charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(report.allocations())
        .and_then(|_| budget.charge_payload_work(report.retained_bytes()))
        .and_then(|_| budget.charge_payload_work(report.scratch_bytes()))
        .map_err(map_lock_error)?;
    let Some(snapshot) = snapshot else {
        return Ok(None);
    };
    let scope = match snapshot.scope() {
        codec::SortScope::EntireTable => Scope::EntireTable,
        codec::SortScope::SelectedRows => Scope::SelectedRows,
    };
    let rules = snapshot
        .rules()
        .iter()
        .map(|rule| {
            let column = ColumnIndex::from_native(rule.column()).map_err(|_| invalid_source())?;
            let direction = match rule.direction() {
                codec::SortDirection::Ascending => Direction::Ascending,
                codec::SortDirection::Descending => Direction::Descending,
            };
            Ok(Rule::new(column, direction))
        })
        .collect::<Result<Vec<_>, BodyTableSortError>>()?;
    let order = Order::with_scope(scope, rules).map_err(|_| invalid_source())?;
    validate_order_columns(&order, columns)?;
    Ok(Some(order))
}

fn snapshot_from_order(order: &Order) -> Result<codec::SortOrderSnapshot, BodyTableSortError> {
    let scope = match order.scope() {
        Scope::EntireTable => codec::SortScope::EntireTable,
        Scope::SelectedRows => codec::SortScope::SelectedRows,
    };
    let rules = order
        .rules()
        .iter()
        .map(|rule| {
            let direction = match rule.direction() {
                Direction::Ascending => codec::SortDirection::Ascending,
                Direction::Descending => codec::SortDirection::Descending,
            };
            codec::SortRule::new(rule.column().native_value(), direction)
        })
        .collect::<Vec<_>>();
    codec::SortOrderSnapshot::new(scope, rules).map_err(map_codec_error)
}

fn rewrite_model(
    source: &Package,
    target: &Target,
    after: Option<Order>,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableSortError> {
    let catalog = &source.state.source;
    // The conservative rewritten-message/archive/Snappy/ZIP and candidate
    // inventory bounds are charged before editable archive or codec-owned
    // candidate allocations.  Prepared reassembly and the final Package
    // reopen retain their private staging allocations; their exact
    // requirements/work are charged on the same budget before execution.
    let bounds = preflight_sort_rewrite(source, target, after.as_ref(), budget)?;
    let component = catalog
        .components()
        .get_index(target.native.model_component_index)
        .ok_or(BodyTableSortError::InvalidSource)?;
    let component_name = component.name();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(BodyTableSortError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyTableSortError::UnsupportedSource);
    }
    let (mut archive, archive_limits) =
        page_layout::editable_archive(source, component_name).map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(target.native.model_object_index)
        .ok_or(BodyTableSortError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.native.model_identifier.get()) {
        return Err(BodyTableSortError::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, target.native.model_message_index)
        .map_err(|_| invalid_source())?;
    let original = object
        .messages
        .get(target.native.model_message_index)
        .filter(|message| message.type_ == target.native.model_message_type)
        .ok_or(BodyTableSortError::InvalidSource)?
        .data
        .clone();
    if decode_sort(&original, target.columns, budget)? != target.before {
        return Err(BodyTableSortError::InvalidSource);
    }
    let desired = after.as_ref().map(snapshot_from_order).transpose()?;
    let options = codec_options(budget, original.len(), target.columns)?;
    let prepared = codec::prepare_table_sort_order_rewrite(&original, desired, options)
        .map_err(map_codec_error)?;
    let requirements = prepared.execution_requirements();
    charge_codec_requirements(budget, requirements, original.len())?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_codec_error)?;
    let report = output.report();
    if report.output_bytes() != requirements.output_bytes
        || report.fields() != requirements.fields
        || report.work_bytes() != requirements.work_bytes
        || report.max_depth() != requirements.max_depth
        || report.rules() != requirements.rules
        || report.allocations() != requirements.allocations
        || report.retained_bytes() != requirements.retained_bytes
        || report.scratch_bytes() != requirements.scratch_bytes
    {
        return Err(BodyTableSortError::Verification);
    }
    let replacement = output.into_bytes();
    object
        .replace_message_preserving_header_with_limits(
            target.native.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: replacement,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed =
        page_layout::compress_archive(archive, archive_limits).map_err(map_page_layout_error)?;
    if compressed.len() > bounds.compressed_bound {
        return Err(BodyTableSortError::Verification);
    }
    let edits = [EntryEdit::new(component_name, compressed.as_slice())];
    let prepared_reassembly = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], catalog.limits())
        .map_err(map_archive_error)?;
    let reassembly_requirements = prepared_reassembly.execution_requirements();
    if reassembly_requirements.output_bytes() > bounds.package_output_bound {
        return Err(BodyTableSortError::Verification);
    }
    charge_reassembly_requirements(budget, reassembly_requirements)?;
    let output = prepared_reassembly
        .execute(reassembly_requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), catalog.limits())
            .map_err(map_archive_error)?;
    table_lock::charge_reopen_work(&candidate_source, budget).map_err(map_lock_error)?;
    Package::from_source_catalog(candidate_source).map_err(map_package_error)
}

#[derive(Clone, Copy)]
struct SortRewriteBounds {
    compressed_bound: usize,
    package_output_bound: usize,
}

fn preflight_sort_rewrite(
    source: &Package,
    target: &Target,
    after: Option<&Order>,
    budget: &mut table_lock::WireBudget,
) -> Result<SortRewriteBounds, BodyTableSortError> {
    const MAX_VARINT_BYTES: usize = 10;
    const MAX_RULE_REWRITE_BYTES: usize = 64;

    let catalog = &source.state.source;
    let component = catalog
        .components()
        .get_index(target.native.model_component_index)
        .ok_or(BodyTableSortError::InvalidSource)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(BodyTableSortError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyTableSortError::UnsupportedSource);
    }
    let original = model_message(source, &target.native)?;
    page_layout::validate_selected_metadata(
        component
            .archive()
            .objects
            .get(target.native.model_object_index)
            .ok_or(BodyTableSortError::InvalidSource)?,
        target.native.model_message_index,
    )
    .map_err(map_page_layout_error)?;
    let rule_count = after.map_or(0, |order| order.rules().len());
    let rewritten_message_bound = original
        .data
        .len()
        .checked_add(
            rule_count
                .checked_mul(MAX_RULE_REWRITE_BYTES)
                .ok_or(BodyTableSortError::InvalidSource)?,
        )
        .and_then(|value| value.checked_add(MAX_RULE_REWRITE_BYTES))
        .ok_or(BodyTableSortError::InvalidSource)?;
    let archive_source_length = parsed_archive_source_length(component.archive())?;
    let archive_bound = archive_source_length
        .checked_sub(original.data.len())
        .and_then(|value| value.checked_add(rewritten_message_bound))
        .and_then(|value| value.checked_add(MAX_VARINT_BYTES.saturating_mul(2)))
        .ok_or(BodyTableSortError::InvalidSource)?;
    let compressed_bound = table_lock::snappy_compressed_bound(archive_bound)
        .ok_or(BodyTableSortError::InvalidSource)?;
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| {
            BodyTableSortError::LimitExceeded {
                kind: BodyTableSortLimitKind::EntryBytes,
                observed: u64::MAX,
                maximum: catalog.limits().max_entry_bytes(),
            }
        })?;
    let package_output_bound = catalog
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(BodyTableSortError::InvalidSource)?;
    budget
        .charge_output_bytes(package_output_bound)
        .and_then(|_| {
            budget.precharge_candidate_reopen(
                catalog,
                package_output_bound,
                target.native.model_component_index,
                compressed_bound,
                archive_bound,
                target.native.model_object_index,
                target.native.model_message_index,
                rewritten_message_bound,
            )
        })
        .map_err(map_lock_error)?;
    Ok(SortRewriteBounds {
        compressed_bound,
        package_output_bound,
    })
}

fn parsed_archive_source_length(
    archive: &litchi_iwa_core::Archive,
) -> Result<usize, BodyTableSortError> {
    let Some(last) = archive.objects.last() else {
        return Ok(0);
    };
    let length = last
        .data_offset
        .checked_add(last.data_length)
        .ok_or(BodyTableSortError::InvalidSource)?;
    usize::try_from(length).map_err(|_| BodyTableSortError::InvalidSource)
}

fn reopen_candidate(
    source: &Package,
    bytes: &SharedBytes,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableSortError> {
    // Apply receives an already materialized exact target artifact.  Its
    // complete output and reopen work are charged on the shared transaction
    // budget before ZIP parsing; the ZIP/catalog parser itself remains a
    // staged private allocation boundary.
    budget
        .charge_output_bytes(bytes.len())
        .and_then(|_| budget.charge_payload_work(bytes.len()))
        .map_err(map_lock_error)?;
    let catalog = SourceCatalog::from_shared_bytes_with_limits(
        Arc::<[u8]>::from(bytes.as_ref()),
        source.state.source.limits(),
    )
    .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .and_then(|_| table_lock::charge_reopen_work(&catalog, budget))
        .map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn validate_order_columns(order: &Order, columns: u32) -> Result<(), BodyTableSortError> {
    if order
        .rules()
        .iter()
        .any(|rule| u64::from(rule.column().native_value()) >= u64::from(columns))
    {
        return Err(invalid_source());
    }
    Ok(())
}

fn codec_options(
    budget: &table_lock::WireBudget,
    source_bytes: usize,
    columns: u32,
) -> Result<codec::DecodeOptions, BodyTableSortError> {
    let limits = budget.wire_limits();
    let fields = budget.remaining_wire_fields();
    if fields == 0 {
        return Err(BodyTableSortError::LimitExceeded {
            kind: BodyTableSortLimitKind::WireFields,
            observed: 1,
            maximum: 0,
        });
    }
    let work = budget.remaining_wire_work();
    if work == 0 {
        return Err(BodyTableSortError::LimitExceeded {
            kind: BodyTableSortLimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let references = budget.remaining_payload_references();
    if references == 0 {
        return Err(BodyTableSortError::LimitExceeded {
            kind: BodyTableSortLimitKind::WireRules,
            observed: 1,
            maximum: 0,
        });
    }
    Ok(codec::DecodeOptions::new(
        source_bytes.max(1).min(limits.max_input_bytes()),
        limits.max_output_bytes().max(1),
        fields,
        work,
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        references,
        usize::try_from(columns).unwrap_or(usize::MAX).max(1),
    )
    .with_max_allocations(fields))
}

fn charge_codec_requirements(
    budget: &mut table_lock::WireBudget,
    requirements: codec::RewriteExecutionRequirements,
    source_bytes: usize,
) -> Result<(), BodyTableSortError> {
    budget
        .charge_payload_work(source_bytes)
        .and_then(|_| {
            budget.charge_codec_report(
                requirements.fields,
                requirements.work_bytes,
                requirements.max_depth,
                0,
            )
        })
        .and_then(|_| budget.charge_payload_work(requirements.output_bytes))
        .and_then(|_| budget.charge_payload_work(requirements.allocations))
        .and_then(|_| budget.charge_payload_work(requirements.retained_bytes))
        .and_then(|_| budget.charge_payload_work(requirements.scratch_bytes))
        .map_err(map_lock_error)
}

fn charge_reassembly_requirements(
    budget: &mut table_lock::WireBudget,
    requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
) -> Result<(), BodyTableSortError> {
    budget
        .charge_payload_work(requirements.output_bytes())
        .and_then(|_| budget.charge_payload_work(requirements.retained_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.scratch_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .map_err(map_lock_error)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableSortError> {
    let left = &source.state.source;
    let right = &candidate.state.source;
    if left.components().len() != right.components().len()
        || left.package().len() != right.package().len()
    {
        return Err(BodyTableSortError::Verification);
    }
    for (before, after) in left.package().iter().zip(right.package().iter()) {
        budget
            .charge_payload_work(
                before
                    .data()
                    .len()
                    .saturating_add(after.data().len())
                    .saturating_add(before.raw_name().len())
                    .saturating_add(after.raw_name().len()),
            )
            .map_err(map_lock_error)?;
        let selected = before.name()
            == left
                .components()
                .get_index(target.model_component_index)
                .ok_or(BodyTableSortError::Verification)?
                .name();
        if before.name() != after.name()
            || before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || (!selected
                && (before.raw_record().local_record() != after.raw_record().local_record()
                    || !central_record_preserved_except_offset(
                        before.raw_record().central_directory_record(),
                        after.raw_record().central_directory_record(),
                    )
                    || before.data() != after.data()))
        {
            return Err(BodyTableSortError::Verification);
        }
    }
    for (component_index, (before, after)) in left
        .components()
        .iter()
        .zip(right.components().iter())
        .enumerate()
    {
        if before.name() != after.name()
            || before.archive().objects.len() != after.archive().objects.len()
        {
            return Err(BodyTableSortError::Verification);
        }
        for (object_index, (before_object, after_object)) in before
            .archive()
            .objects
            .iter()
            .zip(after.archive().objects.iter())
            .enumerate()
        {
            budget
                .charge_payload_work(
                    before_object
                        .messages
                        .iter()
                        .map(|message| message.data.len())
                        .sum::<usize>()
                        .saturating_add(
                            after_object
                                .messages
                                .iter()
                                .map(|message| message.data.len())
                                .sum::<usize>(),
                        )
                        .saturating_add(
                            before_object
                                .archive_info
                                .message_infos
                                .iter()
                                .map(|info| {
                                    info.field_infos
                                        .len()
                                        .saturating_add(info.object_references.len())
                                        .saturating_add(info.data_references.len())
                                })
                                .sum::<usize>(),
                        )
                        .saturating_add(
                            after_object
                                .archive_info
                                .message_infos
                                .iter()
                                .map(|info| {
                                    info.field_infos
                                        .len()
                                        .saturating_add(info.object_references.len())
                                        .saturating_add(info.data_references.len())
                                })
                                .sum::<usize>(),
                        ),
                )
                .map_err(map_lock_error)?;
            if component_index != target.model_component_index
                || object_index != target.model_object_index
            {
                if !before_object.same_content_ignoring_offsets(after_object) {
                    return Err(BodyTableSortError::Verification);
                }
                continue;
            }
            if before_object.archive_info.identifier != after_object.archive_info.identifier
                || before_object.archive_info.should_merge != after_object.archive_info.should_merge
                || before_object.messages.len() != after_object.messages.len()
                || before_object.archive_info.message_infos.len()
                    != after_object.archive_info.message_infos.len()
                || before_object.header_length != after_object.header_length
            {
                return Err(BodyTableSortError::Verification);
            }
            let mut expected_archive_info = before_object.archive_info.clone();
            let selected_before_info = expected_archive_info
                .message_infos
                .get_mut(target.model_message_index)
                .ok_or(BodyTableSortError::Verification)?;
            let selected_after_info = after_object
                .archive_info
                .message_infos
                .get(target.model_message_index)
                .ok_or(BodyTableSortError::Verification)?;
            selected_before_info.length = selected_after_info.length;
            if expected_archive_info != after_object.archive_info {
                return Err(BodyTableSortError::Verification);
            }
            for (index, (before_message, after_message)) in before_object
                .messages
                .iter()
                .zip(after_object.messages.iter())
                .enumerate()
            {
                let before_info = before_object
                    .archive_info
                    .message_infos
                    .get(index)
                    .ok_or(BodyTableSortError::Verification)?;
                let after_info = after_object
                    .archive_info
                    .message_infos
                    .get(index)
                    .ok_or(BodyTableSortError::Verification)?;
                if index == target.model_message_index {
                    let mut normalized_info = before_info.clone();
                    normalized_info.length = after_info.length;
                    if before_message.type_ != after_message.type_ || normalized_info != *after_info
                    {
                        return Err(BodyTableSortError::Verification);
                    }
                } else if before_message != after_message || before_info != after_info {
                    return Err(BodyTableSortError::Verification);
                }
            }
        }
    }
    Ok(())
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

fn invalid_source() -> BodyTableSortError {
    BodyTableSortError::InvalidSource
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableSortError {
    match error {
        table_lock::BodyTableLockError::TableNotFound => BodyTableSortError::TableNotFound,
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableSortError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::AmbiguousSelector => BodyTableSortError::AmbiguousSelector,
        table_lock::BodyTableLockError::UnsupportedSource => BodyTableSortError::UnsupportedSource,
        table_lock::BodyTableLockError::InvalidSource => BodyTableSortError::InvalidSource,
        table_lock::BodyTableLockError::PatchConflict => BodyTableSortError::PatchConflict,
        table_lock::BodyTableLockError::Verification => BodyTableSortError::Verification,
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableSortError::Allocation { amount }
        },
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableSortError::LimitExceeded {
            kind: match kind {
                table_lock::BodyTableLockLimitKind::InputBytes => {
                    BodyTableSortLimitKind::InputBytes
                },
                table_lock::BodyTableLockLimitKind::OutputBytes => {
                    BodyTableSortLimitKind::OutputBytes
                },
                table_lock::BodyTableLockLimitKind::Entries => BodyTableSortLimitKind::Entries,
                table_lock::BodyTableLockLimitKind::EntryBytes => {
                    BodyTableSortLimitKind::EntryBytes
                },
                table_lock::BodyTableLockLimitKind::TotalEntryBytes => {
                    BodyTableSortLimitKind::TotalEntryBytes
                },
                table_lock::BodyTableLockLimitKind::PackageBytes => {
                    BodyTableSortLimitKind::PackageBytes
                },
                table_lock::BodyTableLockLimitKind::PayloadBytes => {
                    BodyTableSortLimitKind::PayloadBytes
                },
                table_lock::BodyTableLockLimitKind::TotalPayloadBytes => {
                    BodyTableSortLimitKind::TotalPayloadBytes
                },
                table_lock::BodyTableLockLimitKind::PayloadObjects => {
                    BodyTableSortLimitKind::PayloadObjects
                },
                table_lock::BodyTableLockLimitKind::PayloadMessages => {
                    BodyTableSortLimitKind::PayloadMessages
                },
                table_lock::BodyTableLockLimitKind::PayloadItems => {
                    BodyTableSortLimitKind::PayloadItems
                },
                table_lock::BodyTableLockLimitKind::PayloadReferences => {
                    BodyTableSortLimitKind::PayloadReferences
                },
                table_lock::BodyTableLockLimitKind::WireBytes => BodyTableSortLimitKind::WireBytes,
                table_lock::BodyTableLockLimitKind::WireFields => {
                    BodyTableSortLimitKind::WireFields
                },
                table_lock::BodyTableLockLimitKind::WireNesting => {
                    BodyTableSortLimitKind::WireNesting
                },
                table_lock::BodyTableLockLimitKind::WireWork => BodyTableSortLimitKind::WireWork,
            },
            observed,
            maximum,
        },
    }
}

fn map_codec_error(error: codec::DecodeError) -> BodyTableSortError {
    if let Some(amount) = error.allocation_amount() {
        return BodyTableSortError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return invalid_source();
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::InputBytes { observed, maximum } => {
            (BodyTableSortLimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::OutputBytes { observed, maximum } => {
            (BodyTableSortLimitKind::WireOutputBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (BodyTableSortLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::WorkBytes { observed, maximum } => {
            (BodyTableSortLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => {
            return BodyTableSortError::LimitExceeded {
                kind: BodyTableSortLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        codec::DecodeLimit::Rules { observed, maximum } => {
            (BodyTableSortLimitKind::WireRules, observed, maximum)
        },
        codec::DecodeLimit::Columns { observed, maximum } => {
            (BodyTableSortLimitKind::WireColumns, observed, maximum)
        },
        codec::DecodeLimit::Allocations { observed, maximum } => {
            (BodyTableSortLimitKind::WireAllocations, observed, maximum)
        },
        codec::DecodeLimit::RetainedBytes { observed, maximum } => {
            (BodyTableSortLimitKind::WireRetainedBytes, observed, maximum)
        },
        codec::DecodeLimit::ScratchBytes { observed, maximum } => {
            (BodyTableSortLimitKind::WireScratchBytes, observed, maximum)
        },
        _ => return invalid_source(),
    };
    BodyTableSortError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_archive_error(error: ArchiveError) -> BodyTableSortError {
    match error {
        ArchiveError::Allocation { amount, .. } => BodyTableSortError::Allocation { amount },
        ArchiveError::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableSortError::LimitExceeded {
            kind: match kind {
                ArchiveLimitKind::InputBytes => BodyTableSortLimitKind::InputBytes,
                ArchiveLimitKind::OutputBytes => BodyTableSortLimitKind::OutputBytes,
                ArchiveLimitKind::Entries => BodyTableSortLimitKind::Entries,
                ArchiveLimitKind::MemberNameBytes | ArchiveLimitKind::MetadataBytes => {
                    BodyTableSortLimitKind::PackageBytes
                },
                ArchiveLimitKind::CompressedEntryBytes | ArchiveLimitKind::EntryBytes => {
                    BodyTableSortLimitKind::EntryBytes
                },
                ArchiveLimitKind::TotalBytes => BodyTableSortLimitKind::TotalEntryBytes,
                ArchiveLimitKind::IwaStreamBytes => BodyTableSortLimitKind::PayloadBytes,
                ArchiveLimitKind::IwaTotalBytes => BodyTableSortLimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        ArchiveError::Iwa(error) => map_core_error(error),
        _ => invalid_source(),
    }
}

fn map_page_layout_error(error: page_layout::PageLayoutError) -> BodyTableSortError {
    match error {
        page_layout::PageLayoutError::UnsupportedSource => BodyTableSortError::UnsupportedSource,
        page_layout::PageLayoutError::InvalidSource
        | page_layout::PageLayoutError::InvalidLayout(_) => BodyTableSortError::InvalidSource,
        page_layout::PageLayoutError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableSortError::LimitExceeded {
            kind: match kind {
                page_layout::PageLayoutLimitKind::InputBytes => BodyTableSortLimitKind::InputBytes,
                page_layout::PageLayoutLimitKind::OutputBytes => {
                    BodyTableSortLimitKind::OutputBytes
                },
                page_layout::PageLayoutLimitKind::Entries => BodyTableSortLimitKind::Entries,
                page_layout::PageLayoutLimitKind::EntryBytes => BodyTableSortLimitKind::EntryBytes,
                page_layout::PageLayoutLimitKind::TotalEntryBytes => {
                    BodyTableSortLimitKind::TotalEntryBytes
                },
                page_layout::PageLayoutLimitKind::PackageBytes => {
                    BodyTableSortLimitKind::PackageBytes
                },
                page_layout::PageLayoutLimitKind::PayloadBytes => {
                    BodyTableSortLimitKind::PayloadBytes
                },
                page_layout::PageLayoutLimitKind::TotalPayloadBytes => {
                    BodyTableSortLimitKind::TotalPayloadBytes
                },
                page_layout::PageLayoutLimitKind::PayloadObjects => {
                    BodyTableSortLimitKind::PayloadObjects
                },
                page_layout::PageLayoutLimitKind::PayloadMessages => {
                    BodyTableSortLimitKind::PayloadMessages
                },
                page_layout::PageLayoutLimitKind::PayloadItems => {
                    BodyTableSortLimitKind::PayloadItems
                },
                page_layout::PageLayoutLimitKind::WireBytes => BodyTableSortLimitKind::WireBytes,
                page_layout::PageLayoutLimitKind::WireFields => BodyTableSortLimitKind::WireFields,
                page_layout::PageLayoutLimitKind::WireNesting => {
                    BodyTableSortLimitKind::WireNesting
                },
                page_layout::PageLayoutLimitKind::WireWork => BodyTableSortLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        page_layout::PageLayoutError::Allocation { amount } => {
            BodyTableSortError::Allocation { amount }
        },
        page_layout::PageLayoutError::Verification => BodyTableSortError::Verification,
        page_layout::PageLayoutError::PatchConflict => BodyTableSortError::PatchConflict,
    }
}

fn map_package_error(error: PackageError) -> BodyTableSortError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => BodyTableSortError::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => BodyTableSortError::LimitExceeded {
            kind: BodyTableSortLimitKind::PayloadObjects,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::PayloadLimit { observed, limit }
        | PackageError::SectionNamesTooLarge { observed, limit } => {
            BodyTableSortError::LimitExceeded {
                kind: BodyTableSortLimitKind::PayloadBytes,
                observed: observed as u64,
                maximum: limit as u64,
            }
        },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::NotPages
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(_) => BodyTableSortError::InvalidSource,
    }
}

fn map_core_error(error: CoreError) -> BodyTableSortError {
    match error {
        CoreError::Allocation { requested, .. } => {
            BodyTableSortError::Allocation { amount: requested }
        },
        CoreError::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableSortError::LimitExceeded {
            kind: match kind {
                CoreLimitKind::ArchiveBytes => BodyTableSortLimitKind::TotalPayloadBytes,
                CoreLimitKind::Objects => BodyTableSortLimitKind::PayloadObjects,
                CoreLimitKind::Messages | CoreLimitKind::MessagesPerObject => {
                    BodyTableSortLimitKind::PayloadMessages
                },
                CoreLimitKind::ObjectBytes
                | CoreLimitKind::MessageBytes
                | CoreLimitKind::HeaderBytes
                | CoreLimitKind::HeaderMemoryBytes => BodyTableSortLimitKind::PayloadBytes,
                CoreLimitKind::HeaderFields | CoreLimitKind::MetadataItems => {
                    BodyTableSortLimitKind::PayloadItems
                },
                CoreLimitKind::HeaderNesting => BodyTableSortLimitKind::WireNesting,
                _ => BodyTableSortLimitKind::PayloadBytes,
            },
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        _ => invalid_source(),
    }
}

fn central_record_preserved_except_offset(before: &[u8], after: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    before.len() == after.len()
        && before.len() >= OFFSET.end
        && before[..OFFSET.start] == after[..OFFSET.start]
        && before[OFFSET.end..] == after[OFFSET.end..]
}
