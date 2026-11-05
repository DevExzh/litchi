//! Exact-source transactions for Pages body-table names.
//!
//! The rooted body-table graph is resolved by `table_lock`; this module only
//! owns the table-model name projection and immutable transaction around it.
//! The strict discovery codec validates the selected model fields while
//! the original model payload remains the source-preservation authority.
//!
//! Resource accounting is a conservative logical operation envelope over
//! package, wire, rewrite, reopen, and locality work. It is not allocator/RSS
//! telemetry for package caches or already-decompressed archives.

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_core::RawMessage;
use litchi_iwa_core::archive::{
    ArchiveReferenceOccurrence, ArchiveReferencePolicy, ArchiveReferenceVisitor,
};
use litchi_iwa_protos::package_metadata_codec::{
    ComponentDescriptor, DataReferenceDescriptor, DataReferenceOwnerDescriptor,
    ExternalReferenceDescriptor, ObjectUuidDescriptor, PackageMetadataVisitor,
    RewriteError as MetadataError, RewriteOptions, inspect_package_metadata_with_visitor,
};
use litchi_iwa_protos::table_model_discovery_codec::{self, DecodeLimit, TableModelSnapshot};
use thiserror::Error as ThisError;

use super::{Package, PackageError, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::name::Name;

mod dependencies;

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// A content-free location associated with one body-table name operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Pages package.
    Package,
    /// One rooted table at a checked zero-based position.
    Table { table: usize },
}

/// Finite resource categories enforced by a body-table name transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source package bytes.
    InputBytes,
    /// Complete candidate package bytes.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical package entry.
    EntryBytes,
    /// Aggregate physical package bytes.
    TotalEntryBytes,
    /// Decoded native payload bytes.
    PayloadBytes,
    /// Aggregate decoded native payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected.
    PayloadObjects,
    /// Native payload messages inspected.
    PayloadMessages,
    /// Native framing/metadata items inspected.
    PayloadItems,
    /// Native references inspected.
    PayloadReferences,
    /// Strict name wire bytes.
    WireBytes,
    /// Strict name output bytes.
    WireOutputBytes,
    /// Strict name fields.
    WireFields,
    /// Strict name nesting.
    WireNesting,
    /// Strict name work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
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

/// Failure from a Pages body-table name read or transaction.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted native graph or selected name payload is malformed.
    #[error("the selected Pages body-table name source is invalid")]
    InvalidSource,
    /// The requested public name violates a value invariant.
    #[error("invalid Pages body-table name: {0}")]
    InvalidName(crate::table::name::Error),
    /// A finite transaction ceiling was exceeded.
    #[error("Pages body-table name {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Maximum configured amount.
        maximum: u64,
    },
    /// A fallible bounded allocation failed.
    #[error("could not allocate {amount} units for the Pages body-table name transaction")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Candidate reopening did not reproduce the requested semantic state.
    #[error("the edited Pages body-table name failed semantic verification")]
    Verification,
    /// The patch was created from another exact package artifact.
    #[error("the Pages body-table name patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// No rooted body table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one rooted body table matched a name selector.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The selector or rooted native graph was ambiguous.
    #[error("the Pages body-table name selector is ambiguous")]
    AmbiguousSelector,
    /// The selected source cannot be edited while preserving exact bytes.
    #[error("the Pages package source does not support exact body-table name editing")]
    UnsupportedSource,
    /// A changed operation targeted a locked table.
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

/// Mutable semantic name staged against one immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    selector: table_lock::OwnedBodyTableSelector,
    target: Option<table_lock::BodyTableTarget>,
    before: Option<Name>,
    name: Option<Name>,
    budget: Option<table_lock::WireBudget>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("resolved", &self.target.is_some())
            .field("name_present", &self.name.is_some())
            .field(
                "name_bytes",
                &self.name.as_ref().map_or(0, |name| name.as_str().len()),
            )
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn resolve(&mut self) -> Result<(), Error> {
        if self.target.is_none() {
            let mut budget = match self.budget.take() {
                Some(budget) => budget,
                None => table_lock::WireBudget::new(self.source.state.source.limits())
                    .map_err(map_lock_error)?,
            };
            let target = match resolve_target_owned(self.source, &self.selector, &mut budget) {
                Ok(target) => target,
                Err(error) => {
                    self.budget = Some(budget);
                    return Err(error);
                },
            };
            let before =
                match resolved_name_at_target_with_budget(self.source, &target, &mut budget) {
                    Ok(before) => before,
                    Err(error) => {
                        self.budget = Some(budget);
                        return Err(error);
                    },
                };
            if self.name.is_none() {
                self.name = Some(before.clone());
            }
            self.before = Some(before);
            self.target = Some(target);
            self.budget = Some(budget);
        }
        Ok(())
    }

    /// Return the selected table path.
    pub fn path(&mut self) -> Result<Path, Error> {
        self.resolve()?;
        self.target
            .as_ref()
            .map(|target| Path::Table {
                table: target.table_position,
            })
            .ok_or(Error::InvalidSource)
    }

    /// Return the source name.
    pub fn before(&mut self) -> Result<&Name, Error> {
        self.resolve()?;
        self.before.as_ref().ok_or(Error::InvalidSource)
    }

    /// Return the staged name.
    pub fn name(&mut self) -> Result<&Name, Error> {
        self.resolve()?;
        self.name.as_ref().ok_or(Error::InvalidSource)
    }

    /// Replace the complete staged lossless name.
    #[must_use]
    pub fn set(mut self, name: Name) -> Self {
        self.name = Some(name);
        self
    }

    /// Validate and stage a borrowed name.
    pub fn set_name(self, name: &str) -> Result<Self, Error> {
        Name::new(name)
            .map(|name| self.set(name))
            .map_err(Error::InvalidName)
    }

    /// Validate and publish the staged name atomically.
    pub fn commit(mut self) -> Result<Commit, Error> {
        self.resolve()?;
        commit_edit(self)
    }
}

/// Exact-source reversible name patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: table_lock::BodyTableTarget,
    before: Name,
    after: Name,
    source_preview_count: usize,
    target_preview_count: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected table path.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            table: self.proof.table_position,
        }
    }

    /// Return the source semantic name.
    #[must_use]
    pub fn before(&self) -> &Name {
        &self.before
    }

    /// Return the target semantic name.
    #[must_use]
    pub fn after(&self) -> &Name {
        &self.after
    }

    /// Return the source fingerprint used for exact conflict detection.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the target fingerprint used for exact conflict detection.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether the semantic and physical artifacts are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source, &self.target) || self.source == self.target)
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            proof: self.proof.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
        }
    }
}

/// Content-free diagnostics from one name publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
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

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Whether exact package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of root previews removed by the edit.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one name transaction.
#[must_use = "a body-table name commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body's lossless name.
    pub fn body_table_name<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Name, Error> {
        let mut budget =
            table_lock::WireBudget::new(self.state.source.limits()).map_err(map_lock_error)?;
        let target = resolve_target_with_budget(self, selector.into(), &mut budget)?;
        resolved_name_at_target_with_budget(self, &target, &mut budget)
    }

    /// Start a selector-first immutable body-table name edit.
    pub fn edit_body_table_name<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            selector: table_lock::OwnedBodyTableSelector::try_from(self, selector.into())
                .map_err(map_lock_error)?,
            target: None,
            before: None,
            name: None,
            budget: None,
        })
    }

    /// Apply an exact-source-checked reversible name patch.
    pub fn apply_body_table_name(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut budget =
            table_lock::WireBudget::new(self.state.source.limits()).map_err(map_lock_error)?;
        charge_fingerprint(self.source_bytes(), &mut budget)?;
        if page_layout::fingerprint(self.source_bytes()) != patch.source_fingerprint
            || self.source_bytes() != patch.source.as_ref()
        {
            return Err(Error::PatchConflict);
        }
        if name_at_target_with_budget(self, &patch.proof, &mut budget)? != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source_preview_count != patch.target_preview_count {
                return Err(Error::PatchConflict);
            }
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact()
            || preview_count_with_budget(self, &mut budget)? != patch.source_preview_count
        {
            return Err(Error::PatchConflict);
        }
        dependencies::validate(self, &patch.proof, &mut budget)?;
        charge_fingerprint(patch.target.as_ref(), &mut budget)?;
        if page_layout::fingerprint(patch.target.as_ref()) != patch.target_fingerprint {
            return Err(Error::PatchConflict);
        }
        let candidate = reopen_target(self, Arc::clone(&patch.target), &mut budget)?;
        if name_at_target_with_budget(&candidate, &patch.proof, &mut budget)? != patch.after
            || preview_count_with_budget(&candidate, &mut budget)? != patch.target_preview_count
        {
            return Err(Error::Verification);
        }
        verify_locality(self, &candidate, &patch.proof, &mut budget)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch
                    .source_preview_count
                    .saturating_sub(patch.target_preview_count),
            ),
        })
    }
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let target = edit.target.ok_or(Error::InvalidSource)?;
    let before = edit.before.ok_or(Error::InvalidSource)?;
    let name = edit.name.ok_or(Error::InvalidSource)?;
    let mut budget = edit.budget.ok_or(Error::InvalidSource)?;
    charge_fingerprint(source.source_bytes(), &mut budget)?;
    let source_preview_count = preview_count_with_budget(source, &mut budget)?;
    let source_bytes: Arc<[u8]> = source.state.source.shared_source();
    let source_fingerprint = page_layout::fingerprint(source_bytes.as_ref());
    if before == name {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                source: Arc::clone(&source_bytes),
                target: source_bytes,
                source_fingerprint,
                target_fingerprint: source_fingerprint,
                proof: target,
                before: before.clone(),
                after: name.clone(),
                source_preview_count,
                target_preview_count: source_preview_count,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    if target.explicit_locked == Some(true) {
        return Err(Error::Refused(Refusal::TableLocked));
    }
    reject_name_collision(source, &target, &name, &mut budget)?;
    dependencies::validate(source, &target, &mut budget)?;
    let package = rewrite_name(source, &target, &before, &name, &mut budget)?;
    let target_owner = package.state.source.shared_source();
    charge_fingerprint(target_owner.as_ref(), &mut budget)?;
    let target_fingerprint = page_layout::fingerprint(target_owner.as_ref());
    let target_preview_count = preview_count_with_budget(&package, &mut budget)?;
    Ok(Commit {
        package,
        patch: Patch {
            source: source_bytes,
            target: target_owner,
            source_fingerprint,
            target_fingerprint,
            proof: target,
            before,
            after: name,
            source_preview_count,
            target_preview_count,
        },
        diagnostics: Diagnostics::published(
            source_preview_count.saturating_sub(target_preview_count),
        ),
    })
}

fn resolve_target_with_budget(
    package: &Package,
    selector: BodyTableSelector<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<table_lock::BodyTableTarget, Error> {
    let target = package
        .resolve_body_table_with_budget(selector, budget)
        .map_err(map_lock_error)?;
    validate_name_authority(package, &target, budget)?;
    Ok(target)
}

fn resolve_target_owned(
    package: &Package,
    selector: &table_lock::OwnedBodyTableSelector,
    budget: &mut table_lock::WireBudget,
) -> Result<table_lock::BodyTableTarget, Error> {
    let target = selector
        .resolve_with_budget(package, budget)
        .map_err(map_lock_error)?;
    validate_name_authority(package, &target, budget)?;
    Ok(target)
}

fn name_at_target_with_budget(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<Name, Error> {
    table_lock::validate_body_table_target(package, target, budget).map_err(map_lock_error)?;
    validate_name_authority(package, target, budget)?;
    resolved_name_at_target_with_budget(package, target, budget)
}

fn resolved_name_at_target_with_budget(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<Name, Error> {
    let message = model_message(package, target)?;
    let snapshot = decode_snapshot(&message.data, budget)?;
    own_name(snapshot.table_name(), budget)
}

fn own_name(value: &str, budget: &mut table_lock::WireBudget) -> Result<Name, Error> {
    budget
        .charge_payload_items(value.len())
        .and_then(|_| budget.charge_payload_work(value.len()))
        .map_err(map_lock_error)?;
    Name::new(value).map_err(Error::InvalidName)
}

struct ArchiveMetadataAuthority;

impl ArchiveReferenceVisitor for ArchiveMetadataAuthority {
    fn visit_reference(
        &mut self,
        _occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        Ok(())
    }
}

fn charge_fingerprint(bytes: &[u8], budget: &mut table_lock::WireBudget) -> Result<(), Error> {
    budget
        .charge_input_source(bytes)
        .and_then(|_| budget.charge_payload_work(bytes.len()))
        .map_err(map_lock_error)
}

fn preview_count_with_budget(
    package: &Package,
    budget: &mut table_lock::WireBudget,
) -> Result<usize, Error> {
    let entries = package.state.source.package();
    let comparison_work = entries
        .len()
        .checked_mul(PREVIEW_ENTRY_NAMES.len())
        .and_then(|work| work.checked_add(PREVIEW_ENTRY_NAMES.len()))
        .ok_or(Error::InvalidSource)?;
    budget
        .charge_payload_work(comparison_work)
        .map_err(map_lock_error)?;
    Ok(PREVIEW_ENTRY_NAMES
        .iter()
        .filter(|name| entries.iter().any(|entry| entry.name() == **name))
        .count())
}

/// Validate the archive-header and package-metadata authority for the model
/// selected by the rooted table graph.  Table-lock proves the local route;
/// this additional pass prevents an unrelated archive object or stale UUID
/// map from becoming a second owner of the name.
fn validate_name_authority(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    // Precharge a source-sized envelope before any package-wide role,
    // ArchiveInfo, inbound-reference, or metadata-message traversal. The
    // more precise object/field/reference counters below remain additive.
    budget
        .charge_payload_work(package.source_bytes().len())
        .map_err(map_lock_error)?;
    validate_role_messages(package, target)?;
    let mut object_count = 0usize;
    let mut message_count = 0usize;
    let mut metadata_count = 0usize;
    let mut reference_count = 0usize;
    for component in package.state.source.components().iter() {
        object_count = object_count
            .checked_add(component.archive().objects.len())
            .ok_or(Error::InvalidSource)?;
        for object in &component.archive().objects {
            message_count = message_count
                .checked_add(object.messages.len())
                .ok_or(Error::InvalidSource)?;
            metadata_count = metadata_count
                .checked_add(object.archive_info.message_infos.len())
                .ok_or(Error::InvalidSource)?;
            for info in &object.archive_info.message_infos {
                reference_count = reference_count
                    .checked_add(info.object_references.len())
                    .and_then(|count| count.checked_add(info.data_references.len()))
                    .ok_or(Error::InvalidSource)?;
                for field in &info.field_infos {
                    metadata_count = metadata_count.checked_add(1).ok_or(Error::InvalidSource)?;
                    reference_count = reference_count
                        .checked_add(field.object_references.len())
                        .and_then(|count| count.checked_add(field.data_references.len()))
                        .ok_or(Error::InvalidSource)?;
                }
            }
        }
    }
    budget
        .charge_payload_objects(object_count)
        .and_then(|_| budget.charge_payload_messages(message_count))
        .and_then(|_| budget.charge_payload_items(metadata_count))
        .and_then(|_| budget.charge_payload_references(reference_count))
        .and_then(|_| {
            budget.charge_payload_work(
                object_count
                    .saturating_add(message_count)
                    .saturating_add(metadata_count)
                    .saturating_add(reference_count),
            )
        })
        .map_err(map_lock_error)?;
    let archive_limits = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut metadata_visitor = ArchiveMetadataAuthority;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            object
                .inspect_references_with_policy(
                    &mut metadata_visitor,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
    }
    let mut identifiers = HashSet::new();
    identifiers
        .try_reserve(package.stats().total_objects())
        .map_err(|_| Error::Allocation {
            amount: package.stats().total_objects(),
        })?;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            budget.charge_payload_work(1).map_err(map_lock_error)?;
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            if !identifiers.insert(identifier) {
                return Err(Error::InvalidSource);
            }
        }
    }
    let model_info = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .and_then(|c| c.archive().objects.get(target.model_object_index))
        .and_then(|o| o.archive_info.message_infos.get(target.model_message_index))
        .ok_or(Error::InvalidSource)?;
    let mut inbound = 0usize;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            for (message_index, info) in object.archive_info.message_infos.iter().enumerate() {
                if info
                    .object_references
                    .iter()
                    .any(|reference| !identifiers.contains(reference))
                    || info.field_infos.iter().any(|field| {
                        field
                            .object_references
                            .iter()
                            .any(|reference| !identifiers.contains(reference))
                    })
                {
                    return Err(Error::InvalidSource);
                }
                let aggregate = info
                    .object_references
                    .iter()
                    .filter(|id| **id == target.model_identifier.get())
                    .count();
                let data = info
                    .data_references
                    .iter()
                    .filter(|id| **id == target.model_identifier.get())
                    .count();
                let fields = info
                    .field_infos
                    .iter()
                    .map(|field| {
                        field
                            .object_references
                            .iter()
                            .filter(|id| **id == target.model_identifier.get())
                            .count()
                    })
                    .sum::<usize>();
                let field_data = info
                    .field_infos
                    .iter()
                    .map(|field| {
                        field
                            .data_references
                            .iter()
                            .filter(|id| **id == target.model_identifier.get())
                            .count()
                    })
                    .sum::<usize>();
                if aggregate == 0 && data == 0 && fields == 0 && field_data == 0 {
                    continue;
                }
                let selected = component_index == target.component_index
                    && object_index == target.object_index
                    && message_index == target.info_message_index;
                if !selected
                    || aggregate != 1
                    || data != 0
                    || fields > 1
                    || field_data != 0
                    || info.field_infos.iter().any(|field| {
                        field
                            .object_references
                            .contains(&target.model_identifier.get())
                            && field.path.as_slice() != [2]
                    })
                {
                    return Err(Error::InvalidSource);
                }
                inbound = inbound.saturating_add(aggregate);
            }
        }
    }
    if inbound != 1 {
        return Err(Error::InvalidSource);
    }
    // Every selected model-header reference must resolve to a physical object;
    // this rejects stale/unknown archive-info references without constraining
    // opaque fields in unrelated native messages.
    for reference in model_info.object_references.iter() {
        if !identifiers.contains(reference) {
            return Err(Error::InvalidSource);
        }
    }
    for field in &model_info.field_infos {
        for reference in field.object_references.iter() {
            if !identifiers.contains(reference) {
                return Err(Error::InvalidSource);
            }
        }
    }
    validate_metadata_authority(package, target, &identifiers, budget)
}

fn validate_role_messages(
    package: &Package,
    target: &table_lock::BodyTableTarget,
) -> Result<(), Error> {
    let drawable = package
        .state
        .source
        .components()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource)?;
    let table_info_count = drawable
        .messages
        .iter()
        .filter(|message| matches!(message.type_, 6_000 | 6_001 | 6_003))
        .count();
    if target.message_type != 6_000 || table_info_count != 1 {
        return Err(Error::InvalidSource);
    }
    let model = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .and_then(|component| component.archive().objects.get(target.model_object_index))
        .ok_or(Error::InvalidSource)?;
    let model_count = model
        .messages
        .iter()
        .filter(|message| matches!(message.type_, 6_000 | 6_001 | 6_003))
        .count();
    if target.model_message_type != TABLE_MODEL_MESSAGE_TYPE || model_count != 1 {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn reject_name_collision(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    name: &Name,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    if name.as_str() == target.table_name.as_ref() {
        return Ok(());
    }
    match package.resolve_body_table_with_budget(BodyTableSelector::Name(name.as_str()), budget) {
        Ok(found) if found.model_identifier == target.model_identifier => Ok(()),
        Ok(_)
        | Err(table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousTableName)) => {
            Err(Error::InvalidSource)
        },
        Err(table_lock::Error::Refused(crate::table::lock::Refusal::TableNotFound)) => Ok(()),
        Err(error) => Err(map_lock_error(error)),
    }
}

fn validate_metadata_authority(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    identifiers: &HashSet<u64>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    // This transaction owns only the selected model's current component and
    // locator authority. UUID bit patterns and component assignments for
    // unrelated objects have no independent physical truth in this slice, so
    // they remain opaque; duplicate bindings, a unique current locator, and
    // every selected external/ambiguous/data route are still checked below.
    let Some(component) = package
        .state
        .source
        .components()
        .iter()
        .find(|c| c.name() == "Index/Metadata.iwa")
    else {
        return Ok(());
    };
    let mut metadata = None;
    for object in &component.archive().objects {
        for message in &object.messages {
            if message.type_ == 11_006 {
                if metadata.replace(message.data.as_slice()).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    let Some(metadata) = metadata else {
        return Err(Error::InvalidSource);
    };
    budget
        .charge_payload_work(metadata.len())
        .map_err(map_lock_error)?;
    let fields = budget.remaining_wire_fields().max(1);
    let work = budget.remaining_wire_work().max(1);
    let options = RewriteOptions::new(
        metadata
            .len()
            .max(1)
            .min(budget.wire_limits().max_input_bytes()),
        budget.wire_limits().max_output_bytes(),
        fields,
        work,
        u32::try_from(budget.wire_limits().max_nesting()).unwrap_or(u32::MAX),
        package
            .state
            .source
            .components()
            .len()
            .max(identifiers.len())
            .max(1),
        budget.remaining_payload_references().max(1),
        1,
    );
    let expected_locator = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .map(|component| {
            component
                .name()
                .strip_prefix("Index/")
                .unwrap_or(component.name())
        })
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or("Document");
    let mut visitor = MetadataAuthorityVisitor::with_capacity(
        identifiers.len(),
        target.model_identifier.get(),
        expected_locator,
    )?;
    let report = inspect_package_metadata_with_visitor(metadata, options, &mut visitor)
        .map_err(map_metadata_error)?;
    budget
        .charge_codec_report(
            report.report().fields(),
            report.report().work_bytes(),
            report.report().max_depth(),
            report.report().references_scanned(),
        )
        .map_err(map_lock_error)?;
    if visitor.unknown
        || visitor.invalid_authority
        || visitor.duplicate_component
        || visitor.duplicate_object
        || visitor.duplicate_locator
        || !visitor.selected_model
        || visitor.current_locator_identifier != visitor.selected_component_identifier
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

#[derive(Default)]
struct MetadataAuthorityVisitor<'expected> {
    unknown: bool,
    invalid_authority: bool,
    duplicate_component: bool,
    duplicate_object: bool,
    selected_model: bool,
    expected_identifier: u64,
    expected_locator: &'expected str,
    components: HashSet<u64>,
    object_ids: HashSet<u64>,
    current_locator_identifier: Option<u64>,
    selected_component_identifier: Option<u64>,
    duplicate_locator: bool,
}

impl<'expected> MetadataAuthorityVisitor<'expected> {
    fn with_capacity(
        capacity: usize,
        expected_identifier: u64,
        expected_locator: &'expected str,
    ) -> Result<Self, Error> {
        let mut visitor = Self {
            expected_identifier,
            expected_locator,
            ..Self::default()
        };
        visitor
            .components
            .try_reserve(capacity)
            .map_err(|_| Error::Allocation { amount: capacity })?;
        visitor
            .object_ids
            .try_reserve(capacity)
            .map_err(|_| Error::Allocation { amount: capacity })?;
        Ok(visitor)
    }
}

impl PackageMetadataVisitor for MetadataAuthorityVisitor<'_> {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), MetadataError> {
        if !self.components.insert(component.identifier()) {
            self.duplicate_component = true;
        }
        if component.is_current() && component.effective_locator() == self.expected_locator {
            if self
                .current_locator_identifier
                .replace(component.identifier())
                .is_some()
            {
                self.duplicate_locator = true;
            }
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataError> {
        if !self.object_ids.insert(binding.object_identifier()) {
            self.duplicate_object = true;
        }
        if binding.object_identifier() == self.expected_identifier {
            let component = binding.component();
            if !component.is_current() || component.effective_locator() != self.expected_locator {
                self.invalid_authority = true;
            }
            self.selected_component_identifier = Some(component.identifier());
            self.selected_model = true;
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataError> {
        if reference.object_identifier() == Some(self.expected_identifier) {
            self.invalid_authority = true;
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        reference: DataReferenceDescriptor<'_>,
    ) -> Result<(), MetadataError> {
        if reference.data_identifier() == self.expected_identifier {
            self.invalid_authority = true;
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataError> {
        if owner.object_identifier() == self.expected_identifier
            || owner.data_identifier() == self.expected_identifier
        {
            self.invalid_authority = true;
        }
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataError> {
        if identifier == self.expected_identifier {
            self.invalid_authority = true;
        }
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), MetadataError> {
        if object_identifier == self.expected_identifier {
            self.invalid_authority = true;
        }
        Ok(())
    }
}

fn map_metadata_error(error: MetadataError) -> Error {
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    match error.resource_limit() {
        Some(limit) => {
            let (kind, observed, maximum) = match limit {
                litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
                    observed,
                    maximum,
                } => (LimitKind::WireBytes, observed, maximum),
                litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
                    observed,
                    maximum,
                } => (LimitKind::WireOutputBytes, observed, maximum),
                litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields {
                    observed,
                    maximum,
                } => (LimitKind::WireFields, observed, maximum),
                litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work {
                    observed,
                    maximum,
                } => (LimitKind::WireWork, observed, maximum),
                litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting {
                    observed,
                    maximum,
                } => (LimitKind::WireNesting, observed as usize, maximum as usize),
                litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
                    observed,
                    maximum,
                } => (LimitKind::PayloadItems, observed, maximum),
                litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
                    observed,
                    maximum,
                } => (LimitKind::PayloadReferences, observed, maximum),
                litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
                    observed,
                    maximum,
                } => (LimitKind::PayloadItems, observed, maximum),
                _ => (LimitKind::PayloadItems, 1, 0),
            };
            Error::Limit {
                kind,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        None => Error::InvalidSource,
    }
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

fn decode_snapshot<'a>(
    source: &'a [u8],
    budget: &mut table_lock::WireBudget,
) -> Result<TableModelSnapshot<'a>, Error> {
    let limits = budget.wire_limits();
    let options = table_model_discovery_codec::DecodeOptions::new(
        source.len().max(1).min(limits.max_input_bytes()),
        budget.remaining_wire_fields(),
        budget.remaining_wire_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_text_bytes(limits.max_input_bytes())
    .with_max_output_bytes(limits.max_output_bytes())
    .with_max_allocations(budget.remaining_wire_fields().max(1))
    .with_max_retained_bytes(limits.max_output_bytes())
    .with_max_scratch_bytes(limits.max_output_bytes());
    let (snapshot, report) =
        table_model_discovery_codec::decode_table_model_with_report(source, options)
            .map_err(map_codec_error)?;
    budget
        .charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(report.text_bytes())
        .and_then(|_| budget.charge_payload_work(report.allocations()))
        .and_then(|_| budget.charge_payload_work(report.retained_bytes()))
        .and_then(|_| budget.charge_payload_work(report.scratch_bytes()))
        .map_err(map_lock_error)?;
    Ok(snapshot)
}

fn rewrite_name(
    source: &Package,
    target: &table_lock::BodyTableTarget,
    before: &Name,
    after: &Name,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, Error> {
    table_lock::validate_body_table_target(source, target, budget).map_err(map_lock_error)?;
    let (component_name, stream_length) =
        table_lock::preflight_body_table_component(source, target, budget)
            .map_err(map_lock_error)?;
    let (mut archive, archive_limits) =
        page_layout::editable_archive(source, component_name).map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(target.model_object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, target.model_message_index)
        .map_err(map_page_layout_error)?;
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let original_message_length = message.data.len();
    let old = decode_snapshot(&message.data, budget)?;
    let current = own_name(old.table_name(), budget)?;
    if &current != before {
        return Err(Error::InvalidSource);
    }
    let options = table_model_discovery_codec::DecodeOptions::new(
        message
            .data
            .len()
            .max(1)
            .min(budget.wire_limits().max_input_bytes()),
        budget.remaining_wire_fields(),
        budget.remaining_wire_work(),
        u32::try_from(budget.wire_limits().max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_text_bytes(budget.wire_limits().max_input_bytes())
    .with_max_output_bytes(budget.wire_limits().max_output_bytes())
    .with_max_allocations(budget.remaining_wire_fields().max(1))
    .with_max_retained_bytes(budget.wire_limits().max_output_bytes())
    .with_max_scratch_bytes(budget.wire_limits().max_output_bytes());
    let prepared = table_model_discovery_codec::prepare_table_model_name_rewrite(
        &message.data,
        table_model_discovery_codec::TableModelNameWrite::new(after.as_str()).with_fingerprint(
            table_model_discovery_codec::table_model_source_fingerprint(&message.data),
        ),
        options,
    )
    .map_err(map_codec_error)?;
    let requirements = prepared.execution_requirements();
    let package_bound = source
        .source_bytes()
        .len()
        .checked_add(stream_length)
        .and_then(|n| n.checked_add(requirements.output_bytes().saturating_mul(2)))
        .and_then(|n| n.checked_add(1024))
        .ok_or(Error::InvalidSource)?;
    budget
        .charge_codec_report(
            requirements.fields(),
            requirements.work_bytes(),
            requirements.max_depth(),
            0,
        )
        .and_then(|_| budget.charge_output_bytes(requirements.output_bytes()))
        .and_then(|_| budget.charge_output_bytes(package_bound))
        .and_then(|_| budget.charge_payload_bytes(requirements.output_bytes()))
        .and_then(|_| budget.charge_total_payload_bytes(requirements.output_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.work_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.fields()))
        .and_then(|_| budget.charge_payload_work(requirements.text_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .and_then(|_| budget.charge_payload_work(requirements.retained_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.scratch_bytes()))
        .map_err(map_lock_error)?;
    let rewritten = prepared
        .execute(requirements.exact())
        .map_err(map_codec_error)?
        .into_bytes();
    let verified = decode_snapshot(&rewritten, budget)?;
    let verified_name = own_name(verified.table_name(), budget)?;
    if &verified_name != after {
        return Err(Error::Verification);
    }
    object
        .replace_message_preserving_header(
            target.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed = page_layout::compress_archive(source, archive, archive_limits)
        .map_err(map_page_layout_error)?;
    let archive_bound = stream_length
        .checked_sub(original_message_length)
        .and_then(|length| length.checked_add(requirements.output_bytes()))
        .and_then(|length| length.checked_add(1024))
        .ok_or(Error::InvalidSource)?;
    let source_entry = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    let changed_entry_bound = match source_entry.metadata().central().compression_method() {
        0 => compressed.len(),
        8 => table_lock::deflate_compressed_bound(compressed.len()).ok_or(Error::InvalidSource)?,
        _ => return Err(Error::Refused(Refusal::UnsupportedSource)),
    };
    budget
        .precharge_candidate_reopen(
            &source.state.source,
            package_bound,
            target.model_component_index,
            changed_entry_bound,
            archive_bound,
            target.model_object_index,
            target.model_message_index,
            requirements.output_bytes(),
        )
        .map_err(map_lock_error)?;
    let mut deletions = Vec::new();
    deletions
        .try_reserve_exact(PREVIEW_ENTRY_NAMES.len())
        .map_err(|_| Error::Allocation {
            amount: PREVIEW_ENTRY_NAMES.len(),
        })?;
    for name in PREVIEW_ENTRY_NAMES {
        if source
            .state
            .source
            .package()
            .iter()
            .any(|entry| entry.name() == name)
        {
            deletions.push(name);
        }
    }
    let output = source
        .state
        .source
        .package()
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            &deletions,
            source.state.source.limits(),
        )
        .map_err(map_archive_error)?;
    if output.len() > package_bound {
        return Err(Error::Verification);
    }
    budget
        .charge_input_source(&output)
        .and_then(|_| budget.charge_payload_work(output.len()))
        .map_err(map_lock_error)?;
    let catalog = source
        .state
        .source
        .reopen(output.into())
        .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .map_err(map_lock_error)?;
    table_lock::charge_reopen_work(&catalog, budget).map_err(map_lock_error)?;
    let candidate = Package::from_source_catalog(catalog).map_err(map_package_error)?;
    if name_at_target_with_budget(&candidate, target, budget)? != *after {
        return Err(Error::Verification);
    }
    reject_name_collision(&candidate, target, after, budget)?;
    verify_locality(source, &candidate, target, budget)?;
    Ok(candidate)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    let before_entries = source.state.source.package();
    let after_entries = candidate.state.source.package();
    budget
        .charge_payload_work(
            before_entries
                .len()
                .saturating_add(after_entries.len())
                .saturating_add(
                    before_entries
                        .len()
                        .saturating_mul(PREVIEW_ENTRY_NAMES.len()),
                ),
        )
        .map_err(map_lock_error)?;
    let mut after_by_name = HashMap::new();
    after_by_name
        .try_reserve(after_entries.len())
        .map_err(|_| Error::Allocation {
            amount: after_entries.len(),
        })?;
    for after in after_entries.iter() {
        if after_by_name.insert(after.name(), after).is_some() {
            return Err(Error::Verification);
        }
    }
    for preview in PREVIEW_ENTRY_NAMES {
        after_by_name.remove(preview);
    }
    for before in before_entries.iter() {
        if PREVIEW_ENTRY_NAMES
            .iter()
            .any(|name| *name == before.name())
        {
            continue;
        }
        let after = after_by_name
            .remove(before.name())
            .ok_or(Error::Verification)?;
        let selected = before.name()
            == source
                .state
                .source
                .components()
                .get_index(target.model_component_index)
                .ok_or(Error::Verification)?
                .name();
        if before.name() != after.name()
            || before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || (!selected && before.data() != after.data())
        {
            return Err(Error::Verification);
        }
    }
    if !after_by_name.is_empty() {
        return Err(Error::Verification);
    }
    let before_components = source.state.source.components();
    let after_components = candidate.state.source.components();
    if before_components.len() != after_components.len() {
        return Err(Error::Verification);
    }
    for (component_index, (before, after)) in before_components
        .iter()
        .zip(after_components.iter())
        .enumerate()
    {
        if before.name() != after.name()
            || before.archive().objects.len() != after.archive().objects.len()
        {
            return Err(Error::Verification);
        }
        budget
            .charge_payload_work(before.archive().objects.len())
            .map_err(map_lock_error)?;
        for (object_index, (before_object, after_object)) in before
            .archive()
            .objects
            .iter()
            .zip(after.archive().objects.iter())
            .enumerate()
        {
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
                || before_object.header_length != after_object.header_length
                || before_object.messages.len() != after_object.messages.len()
                || before_object.archive_info.message_infos.len()
                    != after_object.archive_info.message_infos.len()
            {
                return Err(Error::Verification);
            }
            for (message_index, (before_message, after_message)) in before_object
                .messages
                .iter()
                .zip(after_object.messages.iter())
                .enumerate()
            {
                if before_message.type_ != after_message.type_
                    || (message_index != target.model_message_index
                        && before_message.data != after_message.data)
                {
                    return Err(Error::Verification);
                }
            }
        }
    }
    Ok(())
}

fn reopen_target(
    source: &Package,
    target: Arc<[u8]>,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, Error> {
    budget
        .charge_output_bytes(target.len())
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(target.len())
        .map_err(map_lock_error)?;
    let catalog = source
        .state
        .source
        .reopen(target)
        .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .map_err(map_lock_error)?;
    table_lock::charge_reopen_work(&catalog, budget).map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
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
        table_lock::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::Error::Allocation { amount } => Error::Allocation { amount },
        table_lock::Error::Verification => Error::Verification,
        table_lock::Error::PatchConflict => Error::PatchConflict,
    }
}

const fn map_lock_limit(kind: table_lock::LimitKind) -> LimitKind {
    use table_lock::LimitKind as Lock;
    match kind {
        Lock::InputBytes => LimitKind::InputBytes,
        Lock::OutputBytes => LimitKind::OutputBytes,
        Lock::Entries => LimitKind::Entries,
        Lock::EntryBytes => LimitKind::EntryBytes,
        Lock::TotalEntryBytes => LimitKind::TotalEntryBytes,
        Lock::PackageBytes => LimitKind::PayloadItems,
        Lock::PayloadBytes => LimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => LimitKind::PayloadObjects,
        Lock::PayloadMessages => LimitKind::PayloadMessages,
        Lock::PayloadItems => LimitKind::PayloadItems,
        Lock::PayloadReferences => LimitKind::PayloadReferences,
        Lock::WireBytes => LimitKind::WireBytes,
        Lock::WireFields => LimitKind::WireFields,
        Lock::WireNesting => LimitKind::WireNesting,
        Lock::WireWork => LimitKind::WireWork,
    }
}

fn map_page_layout_error(error: page_layout::Error) -> Error {
    match error {
        page_layout::Error::Refused(
            crate::page_layout::transaction::Refusal::UnsupportedSource,
        ) => Error::Refused(Refusal::UnsupportedSource),
        page_layout::Error::InvalidSource => Error::InvalidSource,
        page_layout::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_page_limit(kind),
            observed,
            maximum,
        },
        page_layout::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

const fn map_page_limit(kind: page_layout::LimitKind) -> LimitKind {
    use page_layout::LimitKind as Layout;
    match kind {
        Layout::InputBytes => LimitKind::InputBytes,
        Layout::OutputBytes => LimitKind::OutputBytes,
        Layout::Entries => LimitKind::Entries,
        Layout::EntryBytes => LimitKind::EntryBytes,
        Layout::TotalEntryBytes => LimitKind::TotalEntryBytes,
        Layout::PackageBytes => LimitKind::PayloadItems,
        Layout::PayloadBytes => LimitKind::PayloadBytes,
        Layout::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        Layout::PayloadObjects => LimitKind::PayloadObjects,
        Layout::PayloadMessages => LimitKind::PayloadMessages,
        Layout::PayloadItems => LimitKind::PayloadItems,
        Layout::WireBytes => LimitKind::WireBytes,
        Layout::WireFields => LimitKind::WireFields,
        Layout::WireNesting => LimitKind::WireNesting,
        Layout::WireWork => LimitKind::WireWork,
    }
}

fn map_codec_error(error: table_model_discovery_codec::DecodeError) -> Error {
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Fields { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Work { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        Some(DecodeLimit::Text { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Output { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireOutputBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Allocations { observed, maximum })
        | Some(DecodeLimit::Retained { observed, maximum })
        | Some(DecodeLimit::Scratch { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(_) | None => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::PayloadItems,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Reassembly(_) => Error::Refused(Refusal::UnsupportedSource),
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::PayloadBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadObjects,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        _ => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::{Error, LimitKind, Package, Refusal};
    use crate::selector::BodyTableSelector;

    fn fixture() -> Result<Package, Box<dyn std::error::Error>> {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/pages/body-table-visible.pages");
        Ok(Package::open(path)?)
    }

    #[test]
    fn edit_debug_redacts_staged_table_name() -> Result<(), Box<dyn std::error::Error>> {
        let package = fixture()?;
        let mut edit = package
            .edit_body_table_name(BodyTableSelector::index(0))?
            .set_name("body-table-debug-sentinel")?;
        edit.name()?;

        let debug = format!("{edit:?}");
        assert!(!debug.contains("body-table-debug-sentinel"));
        assert!(debug.contains("name_present: true"));
        assert!(debug.contains("name_bytes:"));
        Ok(())
    }

    #[test]
    fn deferred_resolution_preserves_wire_budget_after_failure()
    -> Result<(), Box<dyn std::error::Error>> {
        let package = fixture()?;
        let mut edit = package.edit_body_table_name(BodyTableSelector::index(usize::MAX))?;

        assert!(matches!(
            edit.path(),
            Err(Error::Refused(Refusal::TableNotFound))
        ));
        let budget = edit
            .budget
            .as_mut()
            .expect("a failed lazy resolution retains its wire budget");
        budget.exhaust_payload_work_for_test();
        assert!(matches!(
            edit.path(),
            Err(Error::Limit {
                kind: LimitKind::WireWork,
                ..
            })
        ));
        Ok(())
    }
}
