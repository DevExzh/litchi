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
use litchi_iwa_archive::{Error as ArchiveError, LimitKind as ArchiveLimitKind};
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len};
use litchi_iwa_core::{Error as CoreError, LimitKind as CoreLimitKind, RawMessage};
use litchi_iwa_protos::table_sort_order_codec as codec;
use thiserror::Error as ThisError;

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
pub enum Path {
    /// The complete Pages package.
    Package,
    /// One rooted body table at a checked zero-based position.
    Table { table: usize },
}

/// Finite resources governed by one Pages persisted-sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
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

impl fmt::Display for LimitKind {
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
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the selected Pages body-table sort source is invalid")]
    InvalidSource,
    #[error("Pages body-table sort {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
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

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    #[error("the Pages body-table sort selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Pages package source does not support exact body-table sort editing")]
    UnsupportedSource,
    #[error("the selected Pages body table is locked")]
    TableLocked,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::AmbiguousTableName { .. } | Self::AmbiguousSelector { .. } => {
                RefusalKind::InvalidValue
            },
            Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
            Self::TableLocked { .. } => RefusalKind::Protected,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

#[derive(Clone, PartialEq, Eq)]
struct Target {
    native: table_lock::BodyTableTarget,
    columns: u32,
    before: Option<Order>,
}

/// A mutable semantic persisted-sort edit.
pub struct Edit<'a> {
    source: &'a Package,
    selector: table_lock::OwnedBodyTableSelector,
    target: Option<Target>,
    after: Option<Order>,
    staged: bool,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("resolved", &self.target.is_some())
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn target(&mut self) -> Result<&Target, Error> {
        if self.target.is_none() {
            let mut budget = transaction_budget(self.source)?;
            let target = resolve_target_owned(self.source, &self.selector, &mut budget)?;
            if !self.staged {
                self.after = target.before.clone();
            }
            self.target = Some(target);
        }
        self.target.as_ref().ok_or(Error::InvalidSource)
    }

    /// Return the selected table path without exposing native identifiers.
    pub fn path(&mut self) -> Result<Path, Error> {
        self.target()?;
        self.target
            .as_ref()
            .map(|target| Path::Table {
                table: target.native.table_position,
            })
            .ok_or(Error::InvalidSource)
    }

    /// Return the staged order, or `None` when the native field is absent or cleared.
    pub fn order(&mut self) -> Result<Option<&Order>, Error> {
        self.target()?;
        Ok(self.after.as_ref())
    }

    /// Stage a persisted sort order without moving table rows.
    #[must_use]
    pub fn set(mut self, order: Order) -> Self {
        self.after = Some(order);
        self.staged = true;
        self
    }

    /// Remove the persisted sort configuration.
    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self.staged = true;
        self
    }

    /// Alias for [`Self::clear`].
    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    /// Validate and publish this edit atomically.
    pub fn commit(mut self) -> Result<Commit, Error> {
        self.target()?;
        commit_edit(self)
    }
}

/// Reversible exact-source patch for one body-table sort field.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    target: Target,
    before: Option<Order>,
    after: Option<Order>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("table_position", &self.target.native.table_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected table path without exposing native identifiers.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
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
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
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
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body table's persisted sort configuration.
    pub fn body_table_sort_order<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Option<Order>, Error> {
        let mut budget = transaction_budget(self)?;
        let target = resolve_target(self, selector.into(), &mut budget)?;
        Ok(target.before)
    }

    /// Start a selector-first persisted-sort edit.
    pub fn edit_body_table_sort_order<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            selector: table_lock::OwnedBodyTableSelector::try_from(self, selector.into())
                .map_err(map_lock_error)?,
            target: None,
            after: None,
            staged: false,
        })
    }

    /// Apply a reversible exact-source persisted-sort patch.
    pub fn apply_body_table_sort_order(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut budget = transaction_budget(self)?;
        let current = self.state.source.source_bytes();
        budget
            .charge_payload_work(current.len())
            .map_err(map_lock_error)?;
        if fingerprint(current) != patch.source_fingerprint()
            || current != patch.artifacts.source_owner().as_ref()
        {
            return Err(Error::PatchConflict);
        }
        let selected = resolve_target(
            self,
            BodyTableSelector::index(patch.target.native.table_position),
            &mut budget,
        )?;
        if selected.native.model_identifier != patch.target.native.model_identifier
            || selected.before != patch.before
        {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let target_owner = patch.artifacts.target_owner();
        budget
            .charge_payload_work(target_owner.as_ref().len())
            .map_err(map_lock_error)?;
        if !self.state.source.source_is_exact()
            || fingerprint(target_owner.as_ref()) != patch.target_fingerprint()
        {
            return Err(Error::PatchConflict);
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
            return Err(Error::Verification);
        }
        verify_locality(self, &candidate, &patch.target.native, &mut budget)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }
}

fn resolve_target(
    package: &Package,
    selector: BodyTableSelector<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<Target, Error> {
    let native = package
        .resolve_body_table_with_budget(selector, budget)
        .map_err(map_lock_error)?;
    if native.model_message_type != TABLE_MODEL_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
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

fn resolve_target_owned(
    package: &Package,
    selector: &table_lock::OwnedBodyTableSelector,
    budget: &mut table_lock::WireBudget,
) -> Result<Target, Error> {
    let native = selector
        .resolve_with_budget(package, budget)
        .map_err(map_lock_error)?;
    if native.model_message_type != TABLE_MODEL_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
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

fn transaction_budget(package: &Package) -> Result<table_lock::WireBudget, Error> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    Ok(budget)
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let target = edit.target.ok_or(Error::InvalidSource)?;
    if target.before == edit.after {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(
                    SharedBytes::from_shared_slice(source.state.source.shared_source()),
                    SharedBytes::from_shared_slice(source.state.source.shared_source()),
                ),
                target,
                before: edit.after.clone(),
                after: edit.after,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if target.native.explicit_locked == Some(true) {
        return Err(Error::Refused(Refusal::TableLocked));
    }
    if let Some(order) = &edit.after {
        validate_order_columns(order, target.columns)?;
    }
    if !source.state.source.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let mut budget =
        table_lock::WireBudget::new(source.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&source.state.source)
        .map_err(map_lock_error)?;
    table_lock::validate_body_table_target(source, &target.native, &mut budget)
        .map_err(map_lock_error)?;
    let package = rewrite_model(source, &target, edit.after.clone(), &mut budget)?;
    let candidate_target = resolve_target(
        &package,
        BodyTableSelector::index(target.native.table_position),
        &mut budget,
    )?;
    if candidate_target.native.model_identifier != target.native.model_identifier
        || candidate_target.before != edit.after
    {
        return Err(Error::Verification);
    }
    verify_locality(source, &package, &target.native, &mut budget)?;
    let package_source = SharedBytes::from_shared_slice(package.state.source.shared_source());
    Ok(Commit {
        package,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(
                SharedBytes::from_shared_slice(source.state.source.shared_source()),
                package_source,
            ),
            target: target.clone(),
            before: target.before,
            after: edit.after,
        },
        diagnostics: Diagnostics::published(),
    })
}

fn model_message<'a>(
    package: &'a Package,
    target: &table_lock::BodyTableTarget,
) -> Result<&'a RawMessage, Error> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    if object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .is_none()
    {
        return Err(Error::InvalidSource);
    }
    Ok(message)
}

fn validate_sort_tracker(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    model: &RawMessage,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
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
        .ok_or(Error::InvalidSource)?;
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
) -> Result<(), Error> {
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

fn decode_nonzero_varint(source: &[u8]) -> Result<std::num::NonZeroU64, Error> {
    std::num::NonZeroU64::new(decode_canonical_varint(source)?).ok_or_else(invalid_source)
}

fn decode_canonical_varint(source: &[u8]) -> Result<u64, Error> {
    let (value, length) = decode_varint_from_bytes(source).map_err(|_| invalid_source())?;
    if length != source.len() || encoded_len(value) != length {
        return Err(invalid_source());
    }
    Ok(value)
}

fn model_columns(source: &[u8], budget: &mut table_lock::WireBudget) -> Result<u32, Error> {
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
) -> Result<Option<Order>, Error> {
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
        .collect::<Result<Vec<_>, Error>>()?;
    let order = Order::with_scope(scope, rules).map_err(|_| invalid_source())?;
    validate_order_columns(&order, columns)?;
    Ok(Some(order))
}

fn snapshot_from_order(order: &Order) -> Result<codec::SortOrderSnapshot, Error> {
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
) -> Result<Package, Error> {
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
        .ok_or(Error::InvalidSource)?;
    let component_name = component.name();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let (mut archive, archive_limits) =
        page_layout::editable_archive(source, component_name).map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(target.native.model_object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.native.model_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, target.native.model_message_index)
        .map_err(|_| invalid_source())?;
    let original = object
        .messages
        .get(target.native.model_message_index)
        .filter(|message| message.type_ == target.native.model_message_type)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    if decode_sort(&original, target.columns, budget)? != target.before {
        return Err(Error::InvalidSource);
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
        return Err(Error::Verification);
    }
    let replacement = output.into_bytes();
    object
        .replace_message_preserving_header(
            target.native.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: replacement,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed = page_layout::compress_archive(source, archive, archive_limits)
        .map_err(map_page_layout_error)?;
    if compressed.len() > bounds.compressed_bound {
        return Err(Error::Verification);
    }
    let edits = [EntryEdit::new(component_name, compressed.as_slice())];
    let prepared_reassembly = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], catalog.limits())
        .map_err(map_archive_error)?;
    let reassembly_requirements = prepared_reassembly.execution_requirements();
    if reassembly_requirements.output_bytes() > bounds.package_output_bound {
        return Err(Error::Verification);
    }
    charge_reassembly_requirements(budget, reassembly_requirements)?;
    let output = prepared_reassembly
        .execute(reassembly_requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate_source = catalog.reopen(output.into()).map_err(map_archive_error)?;
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
) -> Result<SortRewriteBounds, Error> {
    const MAX_VARINT_BYTES: usize = 10;
    const MAX_RULE_REWRITE_BYTES: usize = 64;

    let catalog = &source.state.source;
    let component = catalog
        .components()
        .get_index(target.native.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let original = model_message(source, &target.native)?;
    page_layout::validate_selected_metadata(
        component
            .archive()
            .objects
            .get(target.native.model_object_index)
            .ok_or(Error::InvalidSource)?,
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
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|value| value.checked_add(MAX_RULE_REWRITE_BYTES))
        .ok_or(Error::InvalidSource)?;
    let archive_source_length = parsed_archive_source_length(component.archive())?;
    let archive_bound = archive_source_length
        .checked_sub(original.data.len())
        .and_then(|value| value.checked_add(rewritten_message_bound))
        .and_then(|value| value.checked_add(MAX_VARINT_BYTES.saturating_mul(2)))
        .ok_or(Error::InvalidSource)?;
    let compressed_bound =
        table_lock::snappy_compressed_bound(archive_bound).ok_or(Error::InvalidSource)?;
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| Error::Limit {
            kind: LimitKind::EntryBytes,
            observed: u64::MAX,
            maximum: catalog.limits().max_entry_bytes(),
        })?;
    let package_output_bound = catalog
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(Error::InvalidSource)?;
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

fn parsed_archive_source_length(archive: &litchi_iwa_core::Archive) -> Result<usize, Error> {
    let Some(last) = archive.objects.last() else {
        return Ok(0);
    };
    let length = last
        .data_offset
        .checked_add(last.data_length)
        .ok_or(Error::InvalidSource)?;
    usize::try_from(length).map_err(|_| Error::InvalidSource)
}

fn reopen_candidate(
    source: &Package,
    bytes: &SharedBytes,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, Error> {
    // Apply receives an already materialized exact target artifact.  Its
    // complete output and reopen work are charged on the shared transaction
    // budget before ZIP parsing; the ZIP/catalog parser itself remains a
    // staged private allocation boundary.
    budget
        .charge_output_bytes(bytes.len())
        .and_then(|_| budget.charge_payload_work(bytes.len()))
        .map_err(map_lock_error)?;
    let catalog = source
        .state
        .source
        .reopen(Arc::<[u8]>::from(bytes.as_ref()))
        .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .and_then(|_| table_lock::charge_reopen_work(&catalog, budget))
        .map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn validate_order_columns(order: &Order, columns: u32) -> Result<(), Error> {
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
) -> Result<codec::DecodeOptions, Error> {
    let limits = budget.wire_limits();
    let fields = budget.remaining_wire_fields();
    if fields == 0 {
        return Err(Error::Limit {
            kind: LimitKind::WireFields,
            observed: 1,
            maximum: 0,
        });
    }
    let work = budget.remaining_wire_work();
    if work == 0 {
        return Err(Error::Limit {
            kind: LimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let references = budget.remaining_payload_references();
    if references == 0 {
        return Err(Error::Limit {
            kind: LimitKind::WireRules,
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
) -> Result<(), Error> {
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
) -> Result<(), Error> {
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
) -> Result<(), Error> {
    let left = &source.state.source;
    let right = &candidate.state.source;
    if left.components().len() != right.components().len()
        || left.package().len() != right.package().len()
    {
        return Err(Error::Verification);
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
                .ok_or(Error::Verification)?
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
            return Err(Error::Verification);
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
            return Err(Error::Verification);
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
                    return Err(Error::Verification);
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
                return Err(Error::Verification);
            }
            let mut expected_archive_info = before_object.archive_info.clone();
            let selected_before_info = expected_archive_info
                .message_infos
                .get_mut(target.model_message_index)
                .ok_or(Error::Verification)?;
            let selected_after_info = after_object
                .archive_info
                .message_infos
                .get(target.model_message_index)
                .ok_or(Error::Verification)?;
            selected_before_info.length = selected_after_info.length;
            if expected_archive_info != after_object.archive_info {
                return Err(Error::Verification);
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
                    .ok_or(Error::Verification)?;
                let after_info = after_object
                    .archive_info
                    .message_infos
                    .get(index)
                    .ok_or(Error::Verification)?;
                if index == target.model_message_index {
                    let mut normalized_info = before_info.clone();
                    normalized_info.length = after_info.length;
                    if before_message.type_ != after_message.type_ || normalized_info != *after_info
                    {
                        return Err(Error::Verification);
                    }
                } else if before_message != after_message || before_info != after_info {
                    return Err(Error::Verification);
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

fn invalid_source() -> Error {
    Error::InvalidSource
}

fn map_lock_error(error: table_lock::Error) -> Error {
    match error {
        table_lock::Error::Refused(crate::table::lock::Refusal::TableNotFound) => {
            Error::Refused(Refusal::TableNotFound)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousTableName) => {
            Error::Refused(Refusal::AmbiguousTableName)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousSelector) => {
            Error::Refused(Refusal::AmbiguousSelector)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        table_lock::Error::InvalidSource => Error::InvalidSource,
        table_lock::Error::PatchConflict => Error::PatchConflict,
        table_lock::Error::Verification => Error::Verification,
        table_lock::Error::Allocation { amount } => Error::Allocation { amount },
        table_lock::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                table_lock::LimitKind::InputBytes => LimitKind::InputBytes,
                table_lock::LimitKind::OutputBytes => LimitKind::OutputBytes,
                table_lock::LimitKind::Entries => LimitKind::Entries,
                table_lock::LimitKind::EntryBytes => LimitKind::EntryBytes,
                table_lock::LimitKind::TotalEntryBytes => LimitKind::TotalEntryBytes,
                table_lock::LimitKind::PackageBytes => LimitKind::PackageBytes,
                table_lock::LimitKind::PayloadBytes => LimitKind::PayloadBytes,
                table_lock::LimitKind::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
                table_lock::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
                table_lock::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
                table_lock::LimitKind::PayloadItems => LimitKind::PayloadItems,
                table_lock::LimitKind::PayloadReferences => LimitKind::PayloadReferences,
                table_lock::LimitKind::WireBytes => LimitKind::WireBytes,
                table_lock::LimitKind::WireFields => LimitKind::WireFields,
                table_lock::LimitKind::WireNesting => LimitKind::WireNesting,
                table_lock::LimitKind::WireWork => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
    }
}

fn map_codec_error(error: codec::DecodeError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return invalid_source();
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::InputBytes { observed, maximum } => {
            (LimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::OutputBytes { observed, maximum } => {
            (LimitKind::WireOutputBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::WorkBytes { observed, maximum } => {
            (LimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => {
            return Error::Limit {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        codec::DecodeLimit::Rules { observed, maximum } => {
            (LimitKind::WireRules, observed, maximum)
        },
        codec::DecodeLimit::Columns { observed, maximum } => {
            (LimitKind::WireColumns, observed, maximum)
        },
        codec::DecodeLimit::Allocations { observed, maximum } => {
            (LimitKind::WireAllocations, observed, maximum)
        },
        codec::DecodeLimit::RetainedBytes { observed, maximum } => {
            (LimitKind::WireRetainedBytes, observed, maximum)
        },
        codec::DecodeLimit::ScratchBytes { observed, maximum } => {
            (LimitKind::WireScratchBytes, observed, maximum)
        },
        _ => return invalid_source(),
    };
    Error::Limit {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_archive_error(error: ArchiveError) -> Error {
    match error {
        ArchiveError::Allocation { amount, .. } => Error::Allocation { amount },
        ArchiveError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                ArchiveLimitKind::InputBytes => LimitKind::InputBytes,
                ArchiveLimitKind::OutputBytes => LimitKind::OutputBytes,
                ArchiveLimitKind::Entries => LimitKind::Entries,
                ArchiveLimitKind::MemberNameBytes | ArchiveLimitKind::MetadataBytes => {
                    LimitKind::PackageBytes
                },
                ArchiveLimitKind::CompressedEntryBytes | ArchiveLimitKind::EntryBytes => {
                    LimitKind::EntryBytes
                },
                ArchiveLimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                ArchiveLimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                ArchiveLimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        ArchiveError::Iwa(error) => map_core_error(error),
        _ => invalid_source(),
    }
}

fn map_page_layout_error(error: page_layout::Error) -> Error {
    match error {
        page_layout::Error::Refused(
            crate::page_layout::transaction::Refusal::UnsupportedSource,
        ) => Error::Refused(Refusal::UnsupportedSource),
        page_layout::Error::InvalidSource | page_layout::Error::InvalidLayout(_) => {
            Error::InvalidSource
        },
        page_layout::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                page_layout::LimitKind::InputBytes => LimitKind::InputBytes,
                page_layout::LimitKind::OutputBytes => LimitKind::OutputBytes,
                page_layout::LimitKind::Entries => LimitKind::Entries,
                page_layout::LimitKind::EntryBytes => LimitKind::EntryBytes,
                page_layout::LimitKind::TotalEntryBytes => LimitKind::TotalEntryBytes,
                page_layout::LimitKind::PackageBytes => LimitKind::PackageBytes,
                page_layout::LimitKind::PayloadBytes => LimitKind::PayloadBytes,
                page_layout::LimitKind::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
                page_layout::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
                page_layout::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
                page_layout::LimitKind::PayloadItems => LimitKind::PayloadItems,
                page_layout::LimitKind::WireBytes => LimitKind::WireBytes,
                page_layout::LimitKind::WireFields => LimitKind::WireFields,
                page_layout::LimitKind::WireNesting => LimitKind::WireNesting,
                page_layout::LimitKind::WireWork => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
        page_layout::Error::Allocation { amount } => Error::Allocation { amount },
        page_layout::Error::Verification => Error::Verification,
        page_layout::Error::PatchConflict => Error::PatchConflict,
    }
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadObjects,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::PayloadLimit { observed, limit }
        | PackageError::SectionNamesTooLarge { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::UnexpectedFormat { .. }
        | PackageError::NotOfficeFile
        | PackageError::MissingIdentity
        | PackageError::InvalidLimit { .. }
        | PackageError::Limit { .. }
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(_) => Error::InvalidSource,
    }
}

fn map_core_error(error: CoreError) -> Error {
    match error {
        CoreError::Allocation { requested, .. } => Error::Allocation { amount: requested },
        CoreError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                CoreLimitKind::ArchiveBytes => LimitKind::TotalPayloadBytes,
                CoreLimitKind::Objects => LimitKind::PayloadObjects,
                CoreLimitKind::Messages | CoreLimitKind::MessagesPerObject => {
                    LimitKind::PayloadMessages
                },
                CoreLimitKind::ObjectBytes
                | CoreLimitKind::MessageBytes
                | CoreLimitKind::HeaderBytes
                | CoreLimitKind::HeaderMemoryBytes => LimitKind::PayloadBytes,
                CoreLimitKind::HeaderFields | CoreLimitKind::MetadataItems => {
                    LimitKind::PayloadItems
                },
                CoreLimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::PayloadBytes,
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
