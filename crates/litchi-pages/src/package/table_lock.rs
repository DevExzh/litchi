//! Immutable, exact-source transactions for Pages body-table locks.

use std::collections::HashMap;
use std::fmt;
use std::mem::size_of;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_archive::{ComponentCatalog, Limits, SourceCatalog};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireFieldView, WireView, patch_varint_field, transform_length_delimited_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{table_model_discovery_codec, text_storage_codec};
use thiserror::Error as ThisError;

use super::{
    Package, PackageError, RootReferences, decode_body_storage, root_references_with_limits,
};
use crate::selector::BodyTableSelector;
use crate::table::lock::State;

const DRAWABLE_ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const LEGACY_TABLE_INFO_MESSAGE_TYPE: u32 = 6_003;
const TABLE_MODEL_MESSAGE_TYPES: [u32; 2] = [6_000, 6_001];
const ROOT_OBJECT_IDENTIFIER: u64 = 1;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const ROOT_BODY_FIELD: u32 = 4;
const TABLE_BODY_FIELD: u32 = 9;
const TABLE_ENTRY_FIELD: u32 = 1;
const TABLE_ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const TABLE_ENTRY_OBJECT_FIELD: u32 = 2;
const DRAWABLE_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const TABLE_INFO_SUPER_FIELD: u32 = 1;
const TABLE_INFO_MODEL_FIELD: u32 = 2;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const OBJECT_REPLACEMENT_CHARACTER: u16 = 0xfffc;
const MAX_BODY_TABLES: usize = 4_096;
const SNAPPY_RAW_MAX_OVERHEAD: usize = 32;
const SNAPPY_RAW_MAX_EXPANSION_DIVISOR: usize = 6;
const SNAPPY_FRAME_HEADER_BYTES: usize = 4;
const DEFLATE_MAX_LITERAL_BITS: usize = 9;
const DEFLATE_BLOCK_OVERHEAD_BYTES: usize = 3;
const MAX_VARINT_BYTES: usize = 10;

/// Finite resources enforced while reading or publishing one body-table lock
/// transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete edited package output bytes.
    OutputBytes,
    /// ZIP members retained by the package.
    Entries,
    /// Bytes retained by one ZIP member.
    EntryBytes,
    /// Aggregate bytes retained by ZIP members.
    TotalEntryBytes,
    /// ZIP container names or structural metadata bytes.
    PackageBytes,
    /// Bytes in one decoded native payload container.
    PayloadBytes,
    /// Aggregate decoded native payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected by the transaction.
    PayloadObjects,
    /// Native payload messages inspected by the transaction.
    PayloadMessages,
    /// Native payload framing or metadata items inspected by the transaction.
    PayloadItems,
    /// Native object references inspected while proving ownership.
    PayloadReferences,
    /// Bytes inspected by the body-table wire projection.
    WireBytes,
    /// Fields inspected by the body-table wire projection.
    WireFields,
    /// Nesting used by the body-table wire projection.
    WireNesting,
    /// Aggregate work charged by the body-table wire projection.
    WireWork,
}

impl fmt::Display for LimitKind {
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
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure from a semantic body-table lock read or immutable transaction.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// Rooted semantic selection did not resolve to one supported native payload.
    #[error("the selected Pages body table has no unambiguous editable lock payload")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages body-table lock {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested resource amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Pages body-table lock transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
    },
    /// Complete candidate reopening did not reproduce the requested lock state.
    #[error("the edited Pages body-table lock failed semantic verification")]
    Verification,
    /// The supplied patch was not created from this exact package artifact.
    #[error("the Pages body-table lock patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The rooted Pages body has no table matching the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one body table has the requested exact name.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// A selector could not be resolved uniquely.
    #[error("the Pages body-table selector is ambiguous")]
    AmbiguousSelector,
    /// The package source cannot publish a preservation-safe changed edit.
    #[error("the Pages package source does not support exact body-table lock editing")]
    UnsupportedSource,
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
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// A mutable semantic body-table lock state staged against one immutable
/// package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    target: BodyTableTarget,
    before: State,
    state: State,
}

/// Owned selector retained by lazy edit constructors.
///
/// Name selectors are copied once at the API boundary so an edit can defer
/// the rooted body graph walk until its first operation without borrowing the
/// caller's temporary string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum OwnedBodyTableSelector {
    Name(Box<str>),
    Position(crate::Position),
}

impl OwnedBodyTableSelector {
    pub(crate) fn try_from(
        package: &Package,
        selector: BodyTableSelector<'_>,
    ) -> Result<Self, Error> {
        match selector {
            BodyTableSelector::Name(name) => {
                let mut budget = WireBudget::new(package.state.source.limits())?;
                budget.admit_selector_name(name.len())?;
                let mut owned = String::new();
                owned
                    .try_reserve_exact(name.len())
                    .map_err(|_| Error::Allocation { amount: name.len() })?;
                owned.push_str(name);
                Ok(Self::Name(owned.into_boxed_str()))
            },
            BodyTableSelector::Position(position) => Ok(Self::Position(position)),
        }
    }

    pub(crate) fn resolve(&self, package: &Package) -> Result<BodyTableTarget, Error> {
        let mut budget = WireBudget::new(package.state.source.limits())?;
        self.resolve_with_budget(package, &mut budget)
    }

    pub(crate) fn resolve_with_budget(
        &self,
        package: &Package,
        budget: &mut WireBudget,
    ) -> Result<BodyTableTarget, Error> {
        match self {
            Self::Name(name) => {
                package.resolve_body_table_with_budget(BodyTableSelector::Name(name), budget)
            },
            Self::Position(position) => package
                .resolve_body_table_with_budget(BodyTableSelector::Position(*position), budget),
        }
    }
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // The source snapshot and target contain native identifiers and raw
        // package details; expose only the semantic lock values.
        formatter
            .debug_struct("Edit")
            .field("before", &self.before)
            .field("state", &self.state)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the lock state that would be published by this edit.
    #[must_use]
    pub const fn state(&self) -> State {
        self.state
    }

    /// Replace the staged semantic lock state.
    pub fn set_state(&mut self, state: State) -> &mut Self {
        self.state = state;
        self
    }

    /// Protect the selected table from interactive editing.
    pub fn lock(&mut self) -> &mut Self {
        self.set_state(State::Locked)
    }

    /// Allow interactive edits to the selected table.
    pub fn unlock(&mut self) -> &mut Self {
        self.set_state(State::Unlocked)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let source = physical_source(self.source)?;
        let mut budget = WireBudget::new(source.limits())?;
        let source_bytes = source.shared_source();
        let source_fingerprint = fingerprint(source.source_bytes(), &mut budget)?;
        if self.before == self.state {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch {
                    source: Arc::clone(&source_bytes),
                    target: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    proof: self.target,
                    before: self.before,
                    after: self.state,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }

        let package = rewrite_lock_state(
            self.source,
            &self.target,
            self.before,
            self.state,
            &mut budget,
        )?;
        let target = physical_source(&package)?.shared_source();
        let target_fingerprint = fingerprint(target.as_ref(), &mut budget)?;
        Ok(Commit {
            package,
            patch: Patch {
                source: source_bytes,
                target,
                source_fingerprint,
                target_fingerprint,
                proof: self.target,
                before: self.before,
                after: self.state,
            },
            diagnostics: Diagnostics::published(),
        })
    }
}

/// Reversible patch bound to the exact source and target package artifacts.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: BodyTableTarget,
    before: State,
    after: State,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("table_position", &self.proof.table_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the semantic state required before this patch can apply.
    #[must_use]
    pub const fn before(&self) -> State {
        self.before
    }

    /// Return the semantic state produced by this patch.
    #[must_use]
    pub const fn after(&self) -> State {
        self.after
    }

    /// Return whether this patch preserves both semantic state and exact bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source, &self.target)
                || self.source.as_ref() == self.target.as_ref())
    }

    /// Return an exact-source inverse that restores the original artifact.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            proof: self.proof.clone(),
            before: self.after,
            after: self.before,
        }
    }
}

/// Compact evidence describing work performed by one committed transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs semantically from source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one body-table lock transaction.
#[must_use = "a Pages body-table lock commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its fully reopened package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct BodyTableTarget {
    pub(crate) table_position: usize,
    pub(crate) table_name: Box<str>,
    pub(crate) table_rows: u32,
    pub(crate) table_columns: u32,
    /// UTF-16 character offset of the body's object-replacement entry.
    ///
    /// The deletion transaction uses this source witness to remove exactly
    /// one body-table anchor after the graph has been proved.  Keeping the
    /// offset beside the private target avoids a second text scan in sibling
    /// phases while leaving the public selector and snapshot archive-free.
    pub(crate) anchor_character_index: usize,
    pub(crate) sheet_identifier: NonZeroU64,
    pub(crate) sheet_component_index: usize,
    pub(crate) sheet_object_index: usize,
    pub(crate) sheet_message_index: usize,
    pub(crate) sheet_message_type: u32,
    pub(crate) attachment_identifier: NonZeroU64,
    pub(crate) attachment_component_index: usize,
    pub(crate) attachment_object_index: usize,
    pub(crate) attachment_message_index: usize,
    pub(crate) drawable_identifier: NonZeroU64,
    pub(crate) model_identifier: NonZeroU64,
    pub(crate) model_component_index: usize,
    pub(crate) model_object_index: usize,
    pub(crate) model_message_index: usize,
    pub(crate) model_message_type: u32,
    pub(crate) component_index: usize,
    pub(crate) object_index: usize,
    pub(crate) message_index: usize,
    pub(crate) message_type: u32,
    pub(crate) info_message_index: usize,
    pub(crate) body_component_index: usize,
    pub(crate) body_object_index: usize,
    pub(crate) body_message_index: usize,
    pub(crate) body_message_type: u32,
    pub(crate) body_identifier: NonZeroU64,
    pub(crate) explicit_locked: Option<bool>,
}

/// The small, presence-preserving `TST.TableInfoArchive` projection owned by
/// this Pages adapter. All unselected fields remain borrowed opaque bytes in
/// the source message and are never represented by the public API.
#[derive(Debug, Clone, Copy)]
struct TableInfoSnapshot {
    table_model: NonZeroU64,
    locked: Option<bool>,
}

#[derive(Debug, Clone, Copy)]
struct ObjectLocation<'a> {
    component_index: usize,
    object_index: usize,
    object: &'a ArchiveObject,
}

impl Package {
    /// Read one body-attached table's effective interactive lock state.
    pub fn body_table_lock<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<State, Error> {
        let target = self.resolve_body_table(selector.into())?;
        Ok(State::from_locked(target.explicit_locked.unwrap_or(false)))
    }

    /// Start a selector-first immutable body-table lock edit.
    pub fn edit_body_table_lock<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let target = self.resolve_body_table(selector.into())?;
        let before = State::from_locked(target.explicit_locked.unwrap_or(false));
        Ok(Edit {
            source: self,
            target,
            before,
            state: before,
        })
    }

    /// Apply an exact-source-checked reversible body-table lock patch.
    pub fn apply_body_table_lock(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        let mut budget = WireBudget::new(source.limits())?;
        if fingerprint(source.source_bytes(), &mut budget)? != patch.source_fingerprint {
            return Err(Error::PatchConflict);
        }
        if !bytes_equal_with_budget(source.source_bytes(), patch.source.as_ref(), &mut budget)? {
            return Err(Error::PatchConflict);
        }
        if body_table_lock_at_target_with_budget(self, &patch.proof, &mut budget)? != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source_fingerprint != patch.target_fingerprint
                || !bytes_equal_with_budget(
                    patch.source.as_ref(),
                    patch.target.as_ref(),
                    &mut budget,
                )?
            {
                return Err(Error::PatchConflict);
            }
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !source.source_is_exact()
            || fingerprint(patch.target.as_ref(), &mut budget)? != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        let candidate = reopen_shared(Arc::clone(&patch.target), source, &mut budget)?;
        if body_table_lock_at_target_with_budget(&candidate, &patch.proof, &mut budget)?
            != patch.after
        {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }

    pub(crate) fn resolve_body_table(
        &self,
        selector: BodyTableSelector<'_>,
    ) -> Result<BodyTableTarget, Error> {
        let source = physical_source(self)?;
        let mut budget = WireBudget::new(source.limits())?;
        self.resolve_body_table_with_budget(selector, &mut budget)
    }

    pub(crate) fn resolve_body_table_with_budget(
        &self,
        selector: BodyTableSelector<'_>,
        budget: &mut WireBudget,
    ) -> Result<BodyTableTarget, Error> {
        let source = physical_source(self)?;
        budget.charge_source_catalog(source)?;
        let targets = native_body_table_targets_with_budget(self, budget)?;
        match selector {
            BodyTableSelector::Position(position) => {
                budget.charge_payload_work(position.get().saturating_add(1).min(targets.len()))?;
                targets
                    .into_iter()
                    .nth(position.get())
                    .ok_or(Error::Refused(Refusal::TableNotFound))
            },
            BodyTableSelector::Name(name) => {
                let mut first = None;
                for target in targets {
                    budget
                        .charge_payload_work(target.table_name.len().saturating_add(name.len()))?;
                    if target.table_name.as_ref() != name {
                        continue;
                    }
                    if first.replace(target).is_some() {
                        return Err(Error::Refused(Refusal::AmbiguousTableName));
                    }
                }
                let Some(first) = first else {
                    return Err(Error::Refused(Refusal::TableNotFound));
                };
                Ok(first)
            },
        }
    }
}

/// Resolve and revalidate one rooted body-table graph for sibling semantic
/// adapters.  The returned target contains only crate-private native proof;
/// it never crosses the Pages public API.
pub(crate) fn validate_body_table_target(
    package: &Package,
    target: &BodyTableTarget,
    budget: &mut WireBudget,
) -> Result<(), Error> {
    validate_selected_ownership(package, target, budget)
}

pub(crate) fn native_body_table_targets_with_budget(
    package: &Package,
    budget: &mut WireBudget,
) -> Result<Vec<BodyTableTarget>, Error> {
    native_body_table_targets_with_budget_mode(package, budget, false)
}

/// Discover the semantic body-table facts used by the focused read catalog.
///
/// The resolver intentionally shares the complete rooted graph walk with the
/// selector-first mutation adapters.  `allow_empty` only changes the normal
/// result for a valid body-less or table-less root; malformed rooted metadata
/// continues to fail closed in the shared walk.
pub(crate) fn body_table_catalog_with_budget(
    package: &Package,
    budget: &mut WireBudget,
) -> Result<Vec<BodyTableTarget>, Error> {
    native_body_table_targets_with_budget_mode(package, budget, true)
}

fn native_body_table_targets_with_budget_mode(
    package: &Package,
    budget: &mut WireBudget,
    allow_empty: bool,
) -> Result<Vec<BodyTableTarget>, Error> {
    let source = physical_source(package)?;
    native_body_table_targets_from_components_with_budget_mode(
        source.components(),
        source.limits(),
        budget,
        allow_empty,
        Some(package),
    )
}

/// Discover rooted Pages body-table targets from an already parsed component
/// catalog.  The caller owns physical ingress validation; this path performs
/// only the format-specific ownership proof and selected metadata projection.
/// It is used by the lazy merge reader so constructing a focused reader does
/// not require constructing the full semantic [`Document`](crate::Document).
pub(crate) fn body_table_targets_from_components_with_budget(
    components: &ComponentCatalog,
    limits: Limits,
    budget: &mut WireBudget,
) -> Result<Vec<BodyTableTarget>, Error> {
    native_body_table_targets_from_components_with_budget_mode(
        components, limits, budget, false, None,
    )
}

fn native_body_table_targets_from_components_with_budget_mode(
    components: &ComponentCatalog,
    limits: Limits,
    budget: &mut WireBudget,
    allow_empty: bool,
    package: Option<&Package>,
) -> Result<Vec<BodyTableTarget>, Error> {
    let object_index = index_objects(components, budget)?;
    let root_component_index = components
        .iter()
        .position(|component| component.name() == "Index/Document.iwa")
        .ok_or(Error::InvalidSource)?;
    let root_location = object_index
        .get(&ROOT_OBJECT_IDENTIFIER)
        .copied()
        .ok_or(Error::InvalidSource)?;
    if root_location.component_index != root_component_index
        || root_location.object.archive_info.identifier != Some(ROOT_OBJECT_IDENTIFIER)
    {
        return Err(Error::InvalidSource);
    }
    let Some((root_message_index, root_message)) = unique_message_index(
        root_location.object.messages.as_slice(),
        ROOT_MESSAGE_TYPE,
        budget,
    )?
    else {
        return Err(Error::InvalidSource);
    };
    validate_selected_metadata(root_location.object, root_message_index)?;
    // Root projection and body-storage decoding are part of the same read
    // transaction. Charge both strict and lazy projection passes before
    // invoking the package-owned decoders, which do not know about this
    // adapter budget.
    budget.charge_payload_work(root_message.data.len())?;
    budget.charge_payload_work(root_message.data.len())?;
    let RootReferences { body, .. } =
        root_references_with_limits(components, limits).map_err(map_package_error)?;
    let root_message_info = root_location
        .object
        .archive_info
        .message_infos
        .get(root_message_index)
        .ok_or(Error::InvalidSource)?;
    let Some(body_identifier) = body else {
        // A body-less Pages document is a valid empty/rootless source, but a
        // retained field-4 object or data edge is contradictory archive
        // metadata. Do not collapse that tampered rooted edge into an
        // ordinary selector miss.
        if root_message_info.field_infos.iter().any(|field| {
            field.path.as_slice() == [ROOT_BODY_FIELD]
                && (!field.object_references.is_empty() || !field.data_references.is_empty())
        }) {
            return Err(Error::InvalidSource);
        }
        return if allow_empty {
            Ok(Vec::new())
        } else {
            Err(Error::Refused(Refusal::TableNotFound))
        };
    };
    if !message_declares_reference(
        root_message_info,
        body_identifier.get(),
        &[ROOT_BODY_FIELD],
        true,
        false,
        budget,
    )? {
        return Err(Error::InvalidSource);
    }
    let body_location = object_index
        .get(&body_identifier.get())
        .copied()
        .ok_or(Error::InvalidSource)?;
    let body_message = unique_text_message(body_location.object, body_identifier, budget)?;
    validate_selected_metadata(body_location.object, body_message.0)?;
    let body_payload = body_message.1.data.as_slice();
    // The body decoder/projection performs strict and lazy passes. Reserve
    // both passes in the shared transaction budget before any decoder-owned
    // allocation can occur.
    budget.charge_payload_work(body_payload.len())?;
    budget.charge_payload_work(body_payload.len())?;
    let body_view = budget.parse(body_payload, 0)?;
    let text_fragment_count = body_view
        .fields()
        .filter(|field| field.number() == 3)
        .count();
    budget.charge_payload_items(text_fragment_count)?;
    if let Some(section_field) = unique_field(budget, &body_view, 17, 2)? {
        let section_view = budget.parse(section_field.payload(), 1)?;
        let section_count = section_view
            .fields()
            .filter(|field| field.number() == 1)
            .count();
        budget.charge_payload_items(section_count)?;
    }
    let body_storage = if package.is_some() {
        // Keep the package mutation path's strict body-storage validation at
        // its historical point in the rooted walk.  This preserves failures
        // for malformed body text even when the table selector later misses.
        Some(
            decode_body_storage(
                &body_location.object.messages,
                body_identifier,
                super::MAX_SECTIONS,
                super::effective_text_limit(limits),
                limits,
            )
            .map_err(map_package_error)?
            .0,
        )
    } else {
        None
    };
    let body_validation = if package.is_none() {
        Some(validate_body_storage_wire(body_payload, limits, budget)?)
    } else {
        None
    };
    let body_message_info = body_location
        .object
        .archive_info
        .message_infos
        .get(body_message.0)
        .ok_or(Error::InvalidSource)?;
    let Some(table_field) = unique_field(budget, &body_view, TABLE_BODY_FIELD, 2)? else {
        // A missing payload table is a normal selector miss only when its
        // rooted archive metadata also carries no field-9 attachment edges.
        // Validate the empty inventory so stale declarations cannot be
        // silently treated as an absent table.
        validate_body_table_ownership(body_message_info, &[], budget)?;
        return if allow_empty {
            Ok(Vec::new())
        } else {
            Err(Error::Refused(Refusal::TableNotFound))
        };
    };
    let table_view = budget.parse(table_field.payload(), 1)?;
    let entry_count = table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRY_FIELD)
        .count();
    if entry_count > MAX_BODY_TABLES {
        return Err(Error::Limit {
            kind: LimitKind::PayloadItems,
            observed: usize_as_u64(entry_count),
            maximum: usize_as_u64(MAX_BODY_TABLES),
        });
    }
    // Charge the complete entry inventory before retaining parsed entries so
    // a bounded transaction cannot be bypassed by the reserve itself.
    budget.charge_payload_items(entry_count)?;
    budget.charge_payload_work(entry_count)?;
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(entry_count)
        .map_err(|_| Error::Allocation {
            amount: entry_count,
        })?;
    for field in table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRY_FIELD)
    {
        validate_field(field, 2)?;
        entries.push(parse_body_table_entry(budget, field.payload())?);
    }
    budget.charge_sort_work(entries.len())?;
    entries.sort_unstable_by_key(|entry| entry.character_index);
    budget.charge_payload_work(entries.len())?;
    if entries
        .windows(2)
        .any(|pair| pair[0].character_index == pair[1].character_index)
    {
        return Err(Error::InvalidSource);
    }
    // The body archive header is shared by every table attachment. Prove its
    // aggregate reference inventory once, before the per-table graph walk, so
    // work and reference charges describe the transaction rather than
    // multiplying with the number of selected tables. The empty case is also
    // checked: a field-9 declaration without a payload entry is malformed,
    // not an ordinary selector miss.
    validate_body_table_ownership(body_message_info, &entries, budget)?;

    let mut targets = Vec::new();
    budget.charge_payload_items(entries.len())?;
    targets
        .try_reserve(entries.len())
        .map_err(|_| Error::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_attachments = Vec::new();
    budget.charge_payload_items(entries.len())?;
    seen_attachments
        .try_reserve(entries.len())
        .map_err(|_| Error::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_drawables = Vec::new();
    budget.charge_payload_items(entries.len())?;
    seen_drawables
        .try_reserve(entries.len())
        .map_err(|_| Error::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_models = Vec::new();
    budget.charge_payload_items(entries.len())?;
    seen_models
        .try_reserve(entries.len())
        .map_err(|_| Error::Allocation {
            amount: entries.len(),
        })?;
    if let Some(body_storage) = body_storage.as_ref() {
        validate_body_placeholders_in_text(body_storage.text(), &entries, budget)?;
    } else {
        let body_validation = body_validation.ok_or(Error::InvalidSource)?;
        validate_body_placeholders_from_wire(body_payload, &entries, body_validation, budget)?;
    }
    for entry in entries {
        let Some(attachment_location) = object_index.get(&entry.identifier.get()).copied() else {
            return Err(Error::InvalidSource);
        };
        budget.charge_payload_work(seen_attachments.len())?;
        budget.charge_payload_work(seen_drawables.len().saturating_add(seen_models.len()))?;
        if seen_attachments.contains(&entry.identifier)
            || seen_drawables.contains(&entry.identifier)
            || seen_models.contains(&entry.identifier)
            || entry.identifier.get() == ROOT_OBJECT_IDENTIFIER
            || entry.identifier.get() == body_identifier.get()
        {
            return Err(Error::InvalidSource);
        }
        seen_attachments.push(entry.identifier);
        let Some((attachment_message_index, attachment_message)) = unique_message_index(
            attachment_location.object.messages.as_slice(),
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            budget,
        )?
        else {
            return Err(Error::InvalidSource);
        };
        validate_selected_metadata(attachment_location.object, attachment_message_index)?;
        let drawable_identifier = parse_drawable_attachment(budget, &attachment_message.data)?;
        let Some(drawable_location) = object_index.get(&drawable_identifier.get()).copied() else {
            return Err(Error::InvalidSource);
        };
        budget.charge_payload_work(seen_drawables.len())?;
        budget.charge_payload_work(seen_attachments.len().saturating_add(seen_models.len()))?;
        if seen_drawables.contains(&drawable_identifier)
            || seen_attachments.contains(&drawable_identifier)
            || seen_models.contains(&drawable_identifier)
            || drawable_identifier.get() == ROOT_OBJECT_IDENTIFIER
            || drawable_identifier.get() == body_identifier.get()
        {
            return Err(Error::InvalidSource);
        }
        seen_drawables.push(drawable_identifier);
        let Some((message_index, message)) =
            unique_table_info(drawable_location.object.messages.as_slice(), budget)?
        else {
            return Err(Error::InvalidSource);
        };
        validate_selected_metadata(drawable_location.object, message_index)?;
        let snapshot = decode_table_info(&message.data, budget, body_identifier)?;
        let model_identifier = snapshot.table_model;
        let Some(model_location) = object_index.get(&model_identifier.get()).copied() else {
            return Err(Error::InvalidSource);
        };
        if model_location.object.archive_info.identifier != Some(model_identifier.get()) {
            return Err(Error::InvalidSource);
        }
        let Some((model_message_index, model_message)) =
            unique_table_model(model_location.object.messages.as_slice(), budget)?
        else {
            return Err(Error::InvalidSource);
        };
        validate_selected_metadata(model_location.object, model_message_index)?;
        if role_identifiers_are_aliased(
            ROOT_OBJECT_IDENTIFIER,
            body_identifier.get(),
            entry.identifier.get(),
            drawable_identifier.get(),
            model_identifier.get(),
        ) {
            return Err(Error::InvalidSource);
        }
        budget.charge_payload_work(seen_models.len())?;
        budget.charge_payload_work(seen_attachments.len().saturating_add(seen_drawables.len()))?;
        if seen_models.contains(&model_identifier)
            || seen_attachments.contains(&model_identifier)
            || seen_drawables.contains(&model_identifier)
            || model_identifier.get() == ROOT_OBJECT_IDENTIFIER
            || model_identifier.get() == body_identifier.get()
        {
            return Err(Error::InvalidSource);
        }
        seen_models.push(model_identifier);
        let (table_name, table_rows, table_columns) =
            decode_table_model_facts(budget, &model_message.data)?;
        let table_position = targets.len();
        let target = BodyTableTarget {
            table_position,
            table_name,
            table_rows,
            table_columns,
            anchor_character_index: entry.character_index,
            // Pages' rooted body is the section/sheet owner for table
            // attachment dependencies. Keep this proof on the target so
            // sibling semantic adapters can inspect the exact body payload
            // without rebuilding the document graph.
            sheet_identifier: body_identifier,
            sheet_component_index: body_location.component_index,
            sheet_object_index: body_location.object_index,
            sheet_message_index: body_message.0,
            sheet_message_type: body_message.1.type_,
            attachment_identifier: entry.identifier,
            attachment_component_index: attachment_location.component_index,
            attachment_object_index: attachment_location.object_index,
            attachment_message_index,
            drawable_identifier,
            model_identifier,
            model_component_index: model_location.component_index,
            model_object_index: model_location.object_index,
            model_message_index,
            model_message_type: model_message.type_,
            component_index: drawable_location.component_index,
            object_index: drawable_location.object_index,
            message_index,
            message_type: message.type_,
            info_message_index: message_index,
            body_component_index: body_location.component_index,
            body_object_index: body_location.object_index,
            body_message_index: body_message.0,
            body_message_type: body_message.1.type_,
            body_identifier,
            explicit_locked: snapshot.locked,
        };
        // Reads and semantic no-ops must carry the same rooted ownership
        // proof as changed publication. Payload links alone are not enough
        // when archive-header metadata is missing, aliased, or redirected.
        if let Some(package) = package {
            validate_selected_ownership(package, &target, budget)?;
        } else {
            validate_selected_ownership_in_components(components, &target, budget)?;
        }
        targets.push(target);
    }
    if targets.is_empty() {
        return if allow_empty {
            Ok(targets)
        } else {
            Err(Error::Refused(Refusal::TableNotFound))
        };
    }
    Ok(targets)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct BodyTableEntry {
    character_index: usize,
    identifier: NonZeroU64,
}

/// Check the character positions used by the package-owned mutation path.
///
/// This deliberately keeps the historical full-text traversal.  The mutable
/// [`Package`] path still uses the semantic storage projection so its strict
/// section and run validation remains unchanged.
fn validate_body_placeholders_in_text(
    text: &str,
    entries: &[BodyTableEntry],
    budget: &mut WireBudget,
) -> Result<(), Error> {
    let mut entry_index = 0;
    for (character_index, character) in text.encode_utf16().enumerate() {
        budget.charge_payload_work(1)?;
        if entries
            .get(entry_index)
            .is_some_and(|entry| entry.character_index == character_index)
        {
            if character != OBJECT_REPLACEMENT_CHARACTER {
                return Err(Error::InvalidSource);
            }
            entry_index += 1;
        }
    }
    (entry_index == entries.len())
        .then_some(())
        .ok_or(Error::InvalidSource)
}

/// Validate and inspect only the borrowed field-3 text projection needed by
/// the metadata-only reader.
///
/// The text-wire validator still checks the complete known TSWP storage tree,
/// canonical framing, UTF-8/UTF-16 lengths, references, and bounded work. Its
/// result is a compact report; no semantic `String` or run vector is built.
/// Buffa then exposes borrowed text fragments for the placeholder check. Once
/// every selected table attachment is proven, trailing text is intentionally
/// left untouched by this focused pass.
#[derive(Clone, Copy)]
struct BodyWireValidation {
    limits: litchi_iwa_text_wire::RewriteLimits,
    validation: litchi_iwa_text_wire::StorageValidation,
}

fn validate_body_storage_wire(
    payload: &[u8],
    limits: Limits,
    budget: &mut WireBudget,
) -> Result<BodyWireValidation, Error> {
    let wire_limits = super::storage_rewrite_limits(limits).map_err(|error| match error {
        super::StorageWireLimitsError::Physical(error) => map_archive_error(error),
        super::StorageWireLimitsError::Wire(_) => Error::InvalidSource,
    })?;
    let validation = litchi_iwa_text_wire::validate_storage(payload, wire_limits)
        .map_err(map_body_storage_wire_error)?;
    budget.charge_payload_work(validation.validation_work())?;
    Ok(BodyWireValidation {
        limits: wire_limits,
        validation,
    })
}

fn validate_body_placeholders_from_wire(
    payload: &[u8],
    entries: &[BodyTableEntry],
    body: BodyWireValidation,
    budget: &mut WireBudget,
) -> Result<(), Error> {
    if entries.is_empty() {
        return Ok(());
    }
    let element_memory = body
        .validation
        .fragments()
        .checked_mul(size_of::<&str>())
        .ok_or(Error::InvalidSource)?;
    let recursion_limit = u32::try_from(body.limits.max_nesting()).unwrap_or(u32::MAX);
    let options = text_storage_codec::DecodeOptions::new(
        body.limits.max_message_bytes(),
        0,
        element_memory,
        recursion_limit,
    );
    let view = text_storage_codec::decode_storage_text(payload, options)
        .map_err(|_| Error::InvalidSource)?;
    if view.len() != body.validation.fragments() {
        return Err(Error::InvalidSource);
    }

    let mut entry_index = 0;
    let mut character_index = 0usize;
    for fragment in view.fragments() {
        for character in fragment.encode_utf16() {
            if entry_index == entries.len() {
                return Ok(());
            }
            budget.charge_payload_work(1)?;
            if entries[entry_index].character_index == character_index {
                if character != OBJECT_REPLACEMENT_CHARACTER {
                    return Err(Error::InvalidSource);
                }
                entry_index += 1;
            }
            character_index = character_index.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    (entry_index == entries.len())
        .then_some(())
        .ok_or(Error::InvalidSource)
}

#[cfg(test)]
fn validate_body_placeholders_in_wire(
    payload: &[u8],
    entries: &[BodyTableEntry],
    limits: Limits,
    budget: &mut WireBudget,
) -> Result<(), Error> {
    let body = validate_body_storage_wire(payload, limits, budget)?;
    validate_body_placeholders_from_wire(payload, entries, body, budget)
}

fn map_body_storage_wire_error(error: litchi_iwa_text_wire::RewriteError) -> Error {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::Limit {
            kind: match resource {
                "message bytes" => LimitKind::WireBytes,
                "fields" => LimitKind::WireFields,
                "nesting" => LimitKind::WireNesting,
                "text fragments" | "table entries" => LimitKind::PayloadItems,
                "text bytes" => LimitKind::PayloadBytes,
                "object references" => LimitKind::PayloadReferences,
                _ => LimitKind::WireWork,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            Error::Allocation { amount }
        },
        litchi_iwa_text_wire::RewriteError::InvalidFormat(_)
        | litchi_iwa_text_wire::RewriteError::Projection(_)
        | litchi_iwa_text_wire::RewriteError::InvalidLimit { .. }
        | litchi_iwa_text_wire::RewriteError::ReversedRange { .. }
        | litchi_iwa_text_wire::RewriteError::RangeOutOfBounds { .. }
        | litchi_iwa_text_wire::RewriteError::SurrogateSplit { .. }
        | litchi_iwa_text_wire::RewriteError::ArithmeticOverflow { .. } => Error::InvalidSource,
        _ => Error::InvalidSource,
    }
}

fn parse_body_table_entry(budget: &mut WireBudget, source: &[u8]) -> Result<BodyTableEntry, Error> {
    let view = budget.parse(source, 2)?;
    let index = unique_field(budget, &view, TABLE_ENTRY_CHARACTER_INDEX_FIELD, 0)?
        .ok_or(Error::InvalidSource)?;
    let character_index = parse_u32(index)? as usize;
    let reference = unique_field(budget, &view, TABLE_ENTRY_OBJECT_FIELD, 2)?;
    let Some(reference) = reference else {
        return Err(Error::InvalidSource);
    };
    Ok(BodyTableEntry {
        character_index,
        identifier: parse_local_reference(budget, reference.payload(), 3)?,
    })
}

fn parse_drawable_attachment(budget: &mut WireBudget, source: &[u8]) -> Result<NonZeroU64, Error> {
    let view = budget.parse(source, 2)?;
    let field = unique_field(budget, &view, DRAWABLE_FIELD, 2)?.ok_or(Error::InvalidSource)?;
    parse_local_reference(budget, field.payload(), 3)
}

fn decode_table_model_facts(
    budget: &mut WireBudget,
    source: &[u8],
) -> Result<(Box<str>, u32, u32), Error> {
    let limits = budget.wire_limits();
    let options = table_model_discovery_codec::DecodeOptions::new(
        source.len().max(1).min(limits.max_input_bytes()),
        budget.remaining_wire_fields(),
        budget.remaining_wire_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_text_bytes(limits.max_input_bytes());
    let (snapshot, report) =
        table_model_discovery_codec::decode_table_model_with_report(source, options)
            .map_err(map_table_model_codec_error)?;
    budget.charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)?;
    budget.charge_payload_work(report.text_bytes())?;
    let text = snapshot.table_name();
    let mut owned = String::new();
    budget.charge_payload_items(text.len())?;
    budget.charge_payload_work(text.len())?;
    owned
        .try_reserve_exact(text.len())
        .map_err(|_| Error::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok((
        owned.into_boxed_str(),
        snapshot.number_of_rows(),
        snapshot.number_of_columns(),
    ))
}

fn map_table_model_codec_error(error: table_model_discovery_codec::DecodeError) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        table_model_discovery_codec::DecodeLimit::Bytes { observed, maximum } => (
            LimitKind::WireBytes,
            usize_as_u64(observed),
            usize_as_u64(maximum),
        ),
        table_model_discovery_codec::DecodeLimit::Fields { observed, maximum } => (
            LimitKind::WireFields,
            usize_as_u64(observed),
            usize_as_u64(maximum),
        ),
        table_model_discovery_codec::DecodeLimit::Work { observed, maximum } => (
            LimitKind::WireWork,
            usize_as_u64(observed),
            usize_as_u64(maximum),
        ),
        table_model_discovery_codec::DecodeLimit::Text { observed, maximum }
        | table_model_discovery_codec::DecodeLimit::Allocations { observed, maximum }
        | table_model_discovery_codec::DecodeLimit::Retained { observed, maximum }
        | table_model_discovery_codec::DecodeLimit::Scratch { observed, maximum }
        | table_model_discovery_codec::DecodeLimit::Output { observed, maximum } => (
            LimitKind::PayloadItems,
            usize_as_u64(observed),
            usize_as_u64(maximum),
        ),
        table_model_discovery_codec::DecodeLimit::Nesting { observed, maximum } => (
            LimitKind::WireNesting,
            u64::from(observed),
            u64::from(maximum),
        ),
        _ => return Error::InvalidSource,
    };
    Error::Limit {
        kind,
        observed,
        maximum,
    }
}

pub(crate) fn parse_local_reference(
    budget: &mut WireBudget,
    source: &[u8],
    depth: usize,
) -> Result<NonZeroU64, Error> {
    budget.charge_payload_references(1)?;
    let view = budget.parse(source, depth)?;
    let mut identifier = None;
    let mut deprecated_type_seen = false;
    let mut external_seen = false;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(Error::InvalidSource);
                }
                identifier = NonZeroU64::new(decode_varint(field.payload())?);
                if identifier.is_none() {
                    return Err(Error::InvalidSource);
                }
            },
            2 => {
                if field.wire_type() != 0 || deprecated_type_seen {
                    return Err(Error::InvalidSource);
                }
                deprecated_type_seen = true;
                let _deprecated_type = decode_varint(field.payload())?;
            },
            3 => {
                if field.wire_type() != 0 || external_seen {
                    return Err(Error::InvalidSource);
                }
                external_seen = true;
                if decode_varint(field.payload())? > 0 {
                    return Err(Error::InvalidSource);
                }
            },
            _ => {},
        }
    }
    identifier.ok_or(Error::InvalidSource)
}

fn parse_u32(field: WireFieldView<'_>) -> Result<u32, Error> {
    let value = decode_varint(field.payload())?;
    u32::try_from(value).map_err(|_| Error::InvalidSource)
}

fn decode_varint(source: &[u8]) -> Result<u64, Error> {
    let (value, length) = decode_varint_from_bytes(source).map_err(|_| Error::InvalidSource)?;
    if length != source.len() || encoded_len(value) != length {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

fn unique_field<'a>(
    budget: &mut WireBudget,
    view: &'a WireView<'a>,
    number: u32,
    wire_type: u8,
) -> Result<Option<WireFieldView<'a>>, Error> {
    budget.charge_payload_work(view.len())?;
    let mut result = None;
    for field in view.fields().filter(|field| field.number() == number) {
        validate_field(field, wire_type)?;
        if result.replace(field).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(result)
}

fn validate_field(field: WireFieldView<'_>, wire_type: u8) -> Result<(), Error> {
    if field.wire_type() != wire_type {
        return Err(Error::InvalidSource);
    }
    field
        .validate_canonical_framing()
        .map_err(|_| Error::InvalidSource)
}

fn unique_text_message<'a>(
    object: &'a ArchiveObject,
    identifier: NonZeroU64,
    budget: &mut WireBudget,
) -> Result<(usize, &'a RawMessage), Error> {
    budget.charge_payload_messages(object.messages.len())?;
    let mut result = None;
    for (index, message) in object.messages.iter().enumerate() {
        budget.charge_payload_work(1)?;
        if matches!(message.type_, 2_001 | 2_022) {
            if result.replace((index, message)).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    result.ok_or_else(|| {
        let _ = identifier;
        Error::InvalidSource
    })
}

fn unique_message_index<'a>(
    messages: &'a [RawMessage],
    message_type: u32,
    budget: &mut WireBudget,
) -> Result<Option<(usize, &'a RawMessage)>, Error> {
    budget.charge_payload_messages(messages.len())?;
    let mut result = None;
    for (index, message) in messages.iter().enumerate() {
        budget.charge_payload_work(1)?;
        if message.type_ != message_type {
            continue;
        }
        if result.replace((index, message)).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(result)
}

fn unique_table_info<'a>(
    messages: &'a [RawMessage],
    budget: &mut WireBudget,
) -> Result<Option<(usize, &'a RawMessage)>, Error> {
    budget.charge_payload_messages(messages.len())?;
    let mut result = None;
    for (index, message) in messages.iter().enumerate() {
        budget.charge_payload_work(1)?;
        if !matches!(
            message.type_,
            TABLE_INFO_MESSAGE_TYPE | LEGACY_TABLE_INFO_MESSAGE_TYPE
        ) {
            continue;
        }
        if result.replace((index, message)).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(result)
}

fn unique_table_model<'a>(
    messages: &'a [RawMessage],
    budget: &mut WireBudget,
) -> Result<Option<(usize, &'a RawMessage)>, Error> {
    budget.charge_payload_messages(messages.len())?;
    let mut result = None;
    for (index, message) in messages.iter().enumerate() {
        budget.charge_payload_work(1)?;
        if !TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_) {
            continue;
        }
        if result.replace((index, message)).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(result)
}

fn role_identifiers_are_aliased(
    root_identifier: u64,
    body_identifier: u64,
    attachment_identifier: u64,
    drawable_identifier: u64,
    model_identifier: u64,
) -> bool {
    let identifiers = [
        root_identifier,
        body_identifier,
        attachment_identifier,
        drawable_identifier,
        model_identifier,
    ];
    identifiers
        .iter()
        .enumerate()
        .any(|(index, identifier)| identifiers[..index].contains(identifier))
}

fn decode_table_info(
    source: &[u8],
    budget: &mut WireBudget,
    body_identifier: NonZeroU64,
) -> Result<TableInfoSnapshot, Error> {
    let view = budget.parse(source, 0)?;
    let super_field =
        unique_field(budget, &view, TABLE_INFO_SUPER_FIELD, 2)?.ok_or(Error::InvalidSource)?;
    let model_field =
        unique_field(budget, &view, TABLE_INFO_MODEL_FIELD, 2)?.ok_or(Error::InvalidSource)?;
    let drawable = budget.parse(super_field.payload(), 1)?;
    let parent =
        unique_field(budget, &drawable, DRAWABLE_PARENT_FIELD, 2)?.ok_or(Error::InvalidSource)?;
    if parse_local_reference(budget, parent.payload(), 3)? != body_identifier {
        return Err(Error::InvalidSource);
    }
    let locked = unique_field(budget, &drawable, DRAWABLE_LOCKED_FIELD, 0)?
        .map(|field| match decode_varint(field.payload())? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(Error::InvalidSource),
        })
        .transpose()?;
    let table_model = parse_local_reference(budget, model_field.payload(), 1)?;
    Ok(TableInfoSnapshot {
        table_model,
        locked,
    })
}

fn index_objects<'a>(
    components: &'a ComponentCatalog,
    budget: &mut WireBudget,
) -> Result<HashMap<u64, ObjectLocation<'a>>, Error> {
    let object_count = components.iter().try_fold(0usize, |count, component| {
        count
            .checked_add(component.archive().objects.len())
            .ok_or(Error::Allocation { amount: usize::MAX })
    })?;
    // Charge the complete object inventory before reserving the index so a
    // small transaction budget cannot be bypassed by a large preallocation.
    budget.charge_payload_objects(object_count)?;
    budget.charge_payload_work(object_count)?;
    let mut index = HashMap::new();
    index
        .try_reserve(object_count)
        .map_err(|_| Error::Allocation {
            amount: object_count,
        })?;
    for (component_index, component) in components.iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            if index
                .insert(
                    identifier,
                    ObjectLocation {
                        component_index,
                        object_index,
                        object,
                    },
                )
                .is_some()
            {
                // A native reference is not semantically resolvable when the
                // same identifier is present in more than one component.
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(index)
}

/// Read the selected table through a previously resolved physical ownership
/// proof.  The candidate package is produced by a same-topology splice (or
/// is the exact target captured by a patch), so re-indexing every native
/// object would only repeat the locate scan that established this proof.
fn body_table_lock_at_target_with_budget(
    package: &Package,
    target: &BodyTableTarget,
    budget: &mut WireBudget,
) -> Result<State, Error> {
    validate_selected_ownership(package, target, budget)?;
    let source = physical_source(package)?;
    let component = source
        .components()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let message = object
        .messages
        .get(target.message_index)
        .filter(|message| message.type_ == target.message_type)
        .ok_or(Error::InvalidSource)?;
    Ok(State::from_locked(
        decode_table_info(&message.data, budget, target.body_identifier)?
            .locked
            .unwrap_or(false),
    ))
}

fn rewrite_lock_state(
    source: &Package,
    target: &BodyTableTarget,
    before: State,
    after: State,
    budget: &mut WireBudget,
) -> Result<Package, Error> {
    let source_catalog = physical_source(source)?;
    let physical_limits = source_catalog.limits();
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    if State::from_locked(target.explicit_locked.unwrap_or(false)) != before {
        return Err(Error::InvalidSource);
    }
    validate_selected_ownership(source, target, budget)?;
    let component = source_catalog
        .components()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource)?;
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    // Account for the compressed component scan and the complete decoded
    // archive inventory before Snappy decompression or Archive::parse can
    // reserve output. The already-parsed source catalog retains the exact
    // decompressed object extents needed for this allocation-free preflight.
    budget.charge_payload_work(entry.data().len())?;
    let expected_stream_length = parsed_archive_source_length(component.archive())?;
    budget.charge_archive_inventory(expected_stream_length, component.archive())?;
    let stream = SnappyStream::decompress(entry.data(), snappy_limits).map_err(map_core_error)?;
    let stream_length = stream.as_bytes().len();
    if stream_length != expected_stream_length {
        return Err(Error::InvalidSource);
    }
    let mut archive = Archive::parse(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive, budget)?;
    let object = archive
        .objects
        .get_mut(target.object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let message = object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != target.message_type {
        return Err(Error::InvalidSource);
    }
    validate_selected_metadata(object, target.message_index)?;
    let original_message_length = message.data.len();
    let patched_bound =
        table_info_rewrite_bound(original_message_length, target.explicit_locked.is_some()).ok_or(
            Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: u64::MAX,
                maximum: physical_limits.max_input_bytes(),
            },
        )?;
    let rewritten_bound = stream_length
        .checked_sub(original_message_length)
        .and_then(|value| value.checked_add(patched_bound))
        .and_then(|value| value.checked_add(16))
        .ok_or(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: physical_limits.max_input_bytes(),
        })?;
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| Error::Limit {
            kind: LimitKind::EntryBytes,
            observed: u64::MAX,
            maximum: physical_limits.max_entry_bytes(),
        })?;
    let compressed_bound = snappy_compressed_bound(rewritten_bound).ok_or(Error::Limit {
        kind: LimitKind::OutputBytes,
        observed: u64::MAX,
        maximum: physical_limits.max_input_bytes(),
    })?;
    let replacement_compressed_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => deflate_compressed_bound(compressed_bound).ok_or(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: physical_limits.max_input_bytes(),
        })?,
        _ => return Err(Error::Refused(Refusal::UnsupportedSource)),
    };
    let package_output_bound = source_catalog
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(replacement_compressed_bound))
        .ok_or(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: physical_limits.max_input_bytes(),
        })?;
    // Reserve every changed-output bound before the first fallible rewrite
    // allocation.  The patched payload is immediately parsed again for
    // semantic verification, so its bounded work is charged up front too;
    // hostile source limits therefore fail before `transform_*` can allocate.
    budget.charge_output_bytes(rewritten_bound)?;
    budget.charge_output_bytes(compressed_bound)?;
    budget.charge_output_bytes(package_output_bound)?;
    budget.charge_payload_bytes(rewritten_bound)?;
    budget.charge_total_payload_bytes(rewritten_bound)?;
    budget.charge_payload_work(original_message_length)?;
    budget.charge_payload_work(patched_bound)?;
    budget.charge_payload_work(rewritten_bound)?;
    let patched = transform_length_delimited_field::<_, litchi_iwa_common::Error>(
        &message.data,
        TABLE_INFO_SUPER_FIELD,
        |drawable| {
            patch_varint_field(
                drawable,
                DRAWABLE_LOCKED_FIELD,
                target.explicit_locked.is_some(),
                Some(u64::from(after.is_locked())),
            )
        },
    )
    .map_err(map_wire_error)?;
    let patched_length = patched.len();
    if patched_length > patched_bound {
        return Err(Error::Verification);
    }
    if State::from_locked(
        decode_table_info(&patched, budget, target.body_identifier)?
            .locked
            .unwrap_or(false),
    ) != after
    {
        return Err(Error::Verification);
    }
    object
        .replace_message_preserving_header(
            target.message_index,
            RawMessage {
                type_: target.message_type,
                data: patched,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let rewritten = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    if rewritten.len() > rewritten_bound {
        return Err(Error::Verification);
    }
    drop(stream);
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten, snappy_limits).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(Error::Verification);
    }
    drop(rewritten);
    budget.charge_payload_work(package_output_bound)?;
    let output = source_catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            physical_limits,
        )
        .map_err(map_archive_error)?;
    if output.len() > package_output_bound {
        return Err(Error::Verification);
    }
    drop(compressed);
    let candidate_source: Arc<[u8]> = output.into();
    let candidate = reopen_shared(candidate_source, source_catalog, budget)?;
    if body_table_lock_at_target_with_budget(&candidate, target, budget)? != after {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

/// Return a checked upper bound for the table-info payload produced by the
/// nested lock rewrite.  A missing lock field adds its one-byte key and
/// one-byte canonical value; both the nested and outer length prefixes may
/// grow, so reserve the complete varint headroom before the transformer gets
/// an opportunity to allocate its output buffer.
fn table_info_rewrite_bound(input_len: usize, lock_present: bool) -> Option<usize> {
    let nested_growth = usize::from(!lock_present).checked_mul(2)?;
    let prefix_headroom = MAX_VARINT_BYTES.checked_mul(2)?;
    input_len
        .checked_add(nested_growth)?
        .checked_add(prefix_headroom)
}

/// Return a checked upper bound for the framed Snappy output produced by
/// [`SnappyStream::compress`].  The raw encoder's documented bound is
/// `32 + input + input / 6` for each independently encoded chunk; every IWA
/// frame adds its four-byte header.  Keeping this reservation conservative is
/// required because `compress_vec` allocates that raw bound before returning
/// the actual compressed bytes.
pub(crate) fn snappy_compressed_bound(input_len: usize) -> Option<usize> {
    let full_chunk_bound = SnappyStream::WRITE_CHUNK_SIZE
        .checked_add(SnappyStream::WRITE_CHUNK_SIZE / SNAPPY_RAW_MAX_EXPANSION_DIVISOR)?
        .checked_add(SNAPPY_RAW_MAX_OVERHEAD)?
        .checked_add(SNAPPY_FRAME_HEADER_BYTES)?;
    let full_chunks = input_len / SnappyStream::WRITE_CHUNK_SIZE;
    let remainder = input_len % SnappyStream::WRITE_CHUNK_SIZE;
    let mut bound = full_chunks.checked_mul(full_chunk_bound)?;
    if remainder != 0 {
        let remainder_bound = remainder
            .checked_add(remainder / SNAPPY_RAW_MAX_EXPANSION_DIVISOR)?
            .checked_add(SNAPPY_RAW_MAX_OVERHEAD)?
            .checked_add(SNAPPY_FRAME_HEADER_BYTES)?;
        bound = bound.checked_add(remainder_bound)?;
    }
    Some(bound)
}

/// Return a checked upper bound for the raw Deflate output emitted by the
/// workspace's `flate2` backend.  This mirrors its fixed-block worst-case
/// literal bound without a zlib wrapper: at most nine bits per input byte,
/// plus the small-input and final-block overheads.
pub(crate) fn deflate_compressed_bound(input_len: usize) -> Option<usize> {
    let literal_bits = input_len
        .checked_mul(DEFLATE_MAX_LITERAL_BITS - 8)?
        .checked_add(7)?;
    let literal_bytes = literal_bits / 8;
    let small_input_overhead = usize::from(input_len == 0) + usize::from(input_len < 9);
    input_len
        .checked_add(literal_bytes)
        .and_then(|value| value.checked_add(small_input_overhead))
        .and_then(|value| value.checked_add(DEFLATE_BLOCK_OVERHEAD_BYTES))
}

/// Validate the selected table's local graph and physical slots.
///
/// The rooted body's complete table-attachment inventory is proved once by
/// [`validate_body_table_ownership`]; keeping that shared proof out of this
/// per-target helper prevents repeated reference scans while retaining the
/// attachment/drawable/model locality checks.
fn validate_selected_ownership(
    package: &Package,
    target: &BodyTableTarget,
    budget: &mut WireBudget,
) -> Result<(), Error> {
    let source = physical_source(package)?;
    validate_selected_ownership_in_components(source.components(), target, budget)
}

fn validate_selected_ownership_in_components(
    components: &ComponentCatalog,
    target: &BodyTableTarget,
    budget: &mut WireBudget,
) -> Result<(), Error> {
    let component = components
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let message_info = object
        .archive_info
        .message_infos
        .get(target.message_index)
        .ok_or(Error::InvalidSource)?;
    validate_selected_metadata(object, target.message_index)?;
    if message_info.type_ != target.message_type {
        return Err(Error::InvalidSource);
    }
    message_declares_reference(
        message_info,
        target.model_identifier.get(),
        &[TABLE_INFO_MODEL_FIELD],
        true,
        false,
        budget,
    )?;
    message_declares_reference(
        message_info,
        target.body_identifier.get(),
        &[TABLE_INFO_SUPER_FIELD, DRAWABLE_PARENT_FIELD],
        false,
        false,
        budget,
    )?;
    let attachment_component = components
        .get_index(target.attachment_component_index)
        .ok_or(Error::InvalidSource)?;
    let attachment = attachment_component
        .archive()
        .objects
        .get(target.attachment_object_index)
        .ok_or(Error::InvalidSource)?;
    if attachment.archive_info.identifier != Some(target.attachment_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let attachment_message_info = attachment
        .archive_info
        .message_infos
        .get(target.attachment_message_index)
        .ok_or(Error::InvalidSource)?;
    if attachment
        .messages
        .get(target.attachment_message_index)
        .is_none_or(|message| message.type_ != DRAWABLE_ATTACHMENT_MESSAGE_TYPE)
    {
        return Err(Error::InvalidSource);
    }
    validate_selected_metadata(attachment, target.attachment_message_index)?;
    if !message_declares_reference(
        attachment_message_info,
        target.drawable_identifier.get(),
        &[DRAWABLE_FIELD],
        true,
        false,
        budget,
    )? {
        return Err(Error::InvalidSource);
    }
    let body_component = components
        .get_index(target.body_component_index)
        .ok_or(Error::InvalidSource)?;
    let body = body_component
        .archive()
        .objects
        .get(target.body_object_index)
        .ok_or(Error::InvalidSource)?;
    if body.archive_info.identifier != Some(target.body_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    validate_selected_metadata(body, target.body_message_index)?;
    let body_info = body
        .archive_info
        .message_infos
        .get(target.body_message_index)
        .ok_or(Error::InvalidSource)?;
    if body_info.type_ != target.body_message_type {
        return Err(Error::InvalidSource);
    }

    // Retain and re-check the terminal model header as part of the selected
    // ownership proof.  The model is not rewritten, but its physical slot and
    // message metadata must remain the same object reached from TableInfo;
    // otherwise a changed transaction could publish against a redirected
    // archive header after resolving the drawable graph.
    let model_component = components
        .get_index(target.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let model = model_component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(Error::InvalidSource)?;
    if model.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let model_message = model
        .messages
        .get(target.model_message_index)
        .ok_or(Error::InvalidSource)?;
    if model_message.type_ != target.model_message_type
        || !TABLE_MODEL_MESSAGE_TYPES.contains(&target.model_message_type)
    {
        return Err(Error::InvalidSource);
    }
    validate_selected_metadata(model, target.model_message_index)?;
    Ok(())
}

fn message_declares_reference(
    message: &litchi_iwa_core::MessageInfo,
    identifier: u64,
    accepted_path: &[u32],
    require_aggregate: bool,
    require_field: bool,
    budget: &mut WireBudget,
) -> Result<bool, Error> {
    budget.charge_payload_items(message.field_infos.len())?;
    budget.charge_payload_references(message.object_references.len())?;
    budget.charge_payload_references(message.data_references.len())?;
    budget.charge_payload_work(message.field_infos.len())?;
    let aggregate_occurrences = message
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate_occurrences > 1 || (require_aggregate && aggregate_occurrences != 1) {
        return Err(Error::InvalidSource);
    }
    if message.data_references.contains(&identifier) {
        return Err(Error::InvalidSource);
    }
    let mut field_declarations = 0usize;
    let mut accepted_fields = 0usize;
    for field in &message.field_infos {
        budget.charge_payload_references(field.object_references.len())?;
        budget.charge_payload_references(field.data_references.len())?;
        budget.charge_payload_work(field.path.path.len())?;
        if field.data_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if field.path.as_slice() == accepted_path {
            accepted_fields = accepted_fields.checked_add(1).ok_or(Error::InvalidSource)?;
            if field.object_references.as_slice() != [identifier]
                || !field.data_references.is_empty()
                || aggregate_occurrences != 1
            {
                return Err(Error::InvalidSource);
            }
        }
        if occurrences != 0 {
            if occurrences != 1 || field.path.as_slice() != accepted_path {
                return Err(Error::InvalidSource);
            }
            field_declarations = field_declarations
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            if field_declarations > 1 {
                return Err(Error::InvalidSource);
            }
        }
    }
    if require_field && accepted_fields != 1 {
        return Err(Error::InvalidSource);
    }
    Ok(true)
}

/// Prove the body-level table attachment inventory once for the complete
/// table projection.  Each table target still proves its attachment,
/// drawable, model, and physical slots independently; only the shared body
/// header scan is coalesced.  This keeps aggregate reference/work charges
/// linear in the body metadata while retaining the exact field-9 ownership
/// requirement for every table entry.
fn validate_body_table_ownership(
    message: &litchi_iwa_core::MessageInfo,
    entries: &[BodyTableEntry],
    budget: &mut WireBudget,
) -> Result<(), Error> {
    budget.charge_payload_items(message.field_infos.len())?;
    budget.charge_payload_references(message.object_references.len())?;
    budget.charge_payload_references(message.data_references.len())?;
    budget.charge_payload_work(message.field_infos.len())?;
    budget.charge_payload_work(entries.len())?;

    // Current Pages writers commonly retain table-attachment ownership only
    // in MessageInfo.object_references and omit field-local declarations. The
    // rooted payload remains authoritative for the field-9 edge, so accept
    // that aggregate-only form while still rejecting duplicates and any
    // contradictory field-local attribution.
    let mut declarations = HashMap::new();
    declarations
        .try_reserve(entries.len())
        .map_err(|_| Error::Allocation {
            amount: entries.len(),
        })?;
    for entry in entries {
        if declarations
            .insert(entry.identifier.get(), (0usize, 0usize))
            .is_some()
        {
            return Err(Error::InvalidSource);
        }
    }
    for identifier in &message.object_references {
        budget.charge_payload_work(1)?;
        if let Some((aggregate_count, _)) = declarations.get_mut(identifier) {
            *aggregate_count = aggregate_count.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    for identifier in &message.data_references {
        budget.charge_payload_work(1)?;
        if declarations.contains_key(identifier) {
            return Err(Error::InvalidSource);
        }
    }
    for field in &message.field_infos {
        budget.charge_payload_references(field.object_references.len())?;
        budget.charge_payload_references(field.data_references.len())?;
        budget.charge_payload_work(field.path.path.len())?;
        for identifier in &field.data_references {
            budget.charge_payload_work(1)?;
            if declarations.contains_key(identifier) {
                return Err(Error::InvalidSource);
            }
        }
        if field.path.as_slice() == [TABLE_BODY_FIELD] {
            if field.object_references.len() != 1 || !field.data_references.is_empty() {
                return Err(Error::InvalidSource);
            }
            let identifier = field.object_references[0];
            let Some((_, field_count)) = declarations.get_mut(&identifier) else {
                return Err(Error::InvalidSource);
            };
            *field_count = field_count.checked_add(1).ok_or(Error::InvalidSource)?;
            if *field_count > 1 {
                return Err(Error::InvalidSource);
            }
            continue;
        }
        for identifier in &field.object_references {
            budget.charge_payload_work(1)?;
            if declarations.contains_key(identifier) {
                return Err(Error::InvalidSource);
            }
        }
    }
    if declarations
        .values()
        .any(|(aggregate_count, _)| *aggregate_count != 1)
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

#[cfg(test)]
fn message_declares_reference_prefix(
    message: &litchi_iwa_core::MessageInfo,
    identifier: u64,
    accepted_prefix: &[u32],
    budget: &mut WireBudget,
) -> Result<bool, Error> {
    budget.charge_payload_items(message.field_infos.len())?;
    budget.charge_payload_references(message.object_references.len())?;
    budget.charge_payload_references(message.data_references.len())?;
    budget.charge_payload_work(message.field_infos.len())?;
    // Retain one aggregate occurrence/declaration counter for this complete
    // message. It is shared by every field check below, avoiding a fresh
    // reference scan for each table attachment.
    let field_reference_capacity =
        message
            .field_infos
            .iter()
            .try_fold(0usize, |capacity, field| {
                capacity
                    .checked_add(field.object_references.len())
                    .ok_or(Error::Allocation { amount: usize::MAX })
            })?;
    let declaration_capacity = message
        .object_references
        .len()
        .checked_add(field_reference_capacity)
        .ok_or(Error::Allocation { amount: usize::MAX })?;
    // Charge the complete nested reference inventory before reserving the
    // aggregate declaration map.  Otherwise a low reference ceiling could
    // be discovered only after this allocation has already taken place.
    for field in &message.field_infos {
        budget.charge_payload_references(field.object_references.len())?;
        budget.charge_payload_references(field.data_references.len())?;
    }
    let mut declarations = HashMap::new();
    budget.charge_payload_work(declaration_capacity)?;
    declarations
        .try_reserve(declaration_capacity)
        .map_err(|_| Error::Allocation {
            amount: declaration_capacity,
        })?;
    for aggregate_identifier in &message.object_references {
        budget.charge_payload_work(1)?;
        let counts = declarations
            .entry(*aggregate_identifier)
            .or_insert((0usize, 0usize));
        counts.1 = counts.1.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    if declarations.get(&identifier).map_or(0, |counts| counts.1) != 1 {
        return Err(Error::InvalidSource);
    }
    if message.data_references.contains(&identifier) {
        return Err(Error::InvalidSource);
    }
    let mut field_declarations = 0usize;
    for field in &message.field_infos {
        budget.charge_payload_work(field.path.path.len())?;
        if field.data_references.iter().any(|candidate| {
            declarations
                .get(candidate)
                .is_some_and(|counts| counts.1 != 0)
        }) {
            return Err(Error::InvalidSource);
        }
        if field.data_references.is_empty() {
            for field_identifier in &field.object_references {
                budget.charge_payload_work(1)?;
                let counts = declarations
                    .entry(*field_identifier)
                    .or_insert((0usize, 0usize));
                counts.0 = counts.0.checked_add(1).ok_or(Error::InvalidSource)?;
            }
        }
        // The selected body table is owned by the exact table-attachment
        // field.  A descendant path is not interchangeable with that edge;
        // only unrelated aggregate references may use their own paths.
        let has_accepted_prefix = field.path.as_slice() == accepted_prefix;

        if has_accepted_prefix {
            // An aggregate edge is only attributable when its FieldInfo is
            // present and exact.  In particular, an empty or data-only field
            // must not be treated as an optional declaration, and a second
            // record at the accepted path must not be ignored.
            if field.object_references.len() != 1 || !field.data_references.is_empty() {
                return Err(Error::InvalidSource);
            }
            let field_identifier = field.object_references[0];
            if declarations
                .get(&field_identifier)
                .map_or(0, |counts| counts.1)
                != 1
            {
                return Err(Error::InvalidSource);
            }
            if field_identifier == identifier {
                field_declarations = field_declarations
                    .checked_add(1)
                    .ok_or(Error::InvalidSource)?;
                if field_declarations > 1 {
                    return Err(Error::InvalidSource);
                }
            }
        } else {
            // Body storage messages can carry more than the selected table
            // attachment inventory.  For example, section boundaries live at
            // field 17 while table attachments live at field 9.  An
            // aggregate reference outside the selected prefix is valid when
            // its own FieldInfo path accounts for that exact occurrence; it
            // must not be reinterpreted as a field-9 table edge merely because
            // this validator is proving one selected table.
            for (field_index, field_identifier) in field.object_references.iter().enumerate() {
                if declarations
                    .get(field_identifier)
                    .map_or(0, |counts| counts.1)
                    != 1
                {
                    return Err(Error::InvalidSource);
                }
                budget.charge_payload_work(field_index)?;
                if field.object_references[..field_index].contains(field_identifier) {
                    return Err(Error::InvalidSource);
                }
                if *field_identifier == identifier {
                    // The selected attachment must stay rooted at the exact
                    // table-attachment path, even when unrelated aggregate
                    // references are accepted elsewhere in the body.
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    if field_declarations != 1 {
        return Err(Error::InvalidSource);
    }
    // Every aggregate reference must have exactly one field-local declaration.
    // The counters above make this final check linear in the reference count.
    for aggregate_identifier in &message.object_references {
        budget.charge_payload_work(1)?;
        let Some((field_count, aggregate_count)) = declarations.get(aggregate_identifier) else {
            return Err(Error::InvalidSource);
        };
        if *aggregate_count != 1 || *field_count != 1 {
            return Err(Error::InvalidSource);
        }
    }
    Ok(true)
}

fn validate_selected_metadata(object: &ArchiveObject, message_index: usize) -> Result<(), Error> {
    let message = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let message_info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != message_info.type_
        || object.archive_info.should_merge == Some(true)
        || message_info.base_message_index.is_some()
        || !message_info.diff_merge_version.is_empty()
        || message_info.diff_field_path.is_some()
        || !message_info.fields_to_remove.is_empty()
        || !message_info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn parsed_archive_source_length(archive: &Archive) -> Result<usize, Error> {
    let Some(last) = archive.objects.last() else {
        return Ok(0);
    };
    let length = last
        .data_offset
        .checked_add(last.data_length)
        .ok_or(Error::InvalidSource)?;
    usize::try_from(length).map_err(|_| Error::InvalidSource)
}

/// Preflight one sibling body-table payload before its archive decoder is
/// allowed to allocate.  The source catalog and selected compressed member
/// are charged through the same transaction budget used by the lock owner.
pub(crate) fn preflight_body_table_component<'a>(
    package: &'a Package,
    target: &BodyTableTarget,
    budget: &mut WireBudget,
) -> Result<(&'a str, usize), Error> {
    let source = physical_source(package)?;
    budget.charge_source_catalog(source)?;
    let component = source
        .components()
        .get_index(target.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let entry = source
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    budget.charge_payload_work(entry.data().len())?;
    let stream_length = parsed_archive_source_length(component.archive())?;
    budget.charge_archive_inventory(stream_length, component.archive())?;
    Ok((component.name(), stream_length))
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
    budget: &mut WireBudget,
) -> Result<(), Error> {
    for object in &archive.objects {
        budget.charge_payload_work(1)?;
        let offset = usize::try_from(object.header_offset).map_err(|_| Error::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(Error::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_| Error::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(Error::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(u64::try_from(prefix_bytes).map_err(|_| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(Error::InvalidSource)?
                != object.data_offset
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn physical_source(package: &Package) -> Result<&SourceCatalog, Error> {
    Ok(&package.state.source)
}

fn reopen_shared(
    source: Arc<[u8]>,
    previous: &SourceCatalog,
    budget: &mut WireBudget,
) -> Result<Package, Error> {
    budget.charge_input_source(source.as_ref())?;
    budget.charge_payload_work(source.len())?;
    let catalog = previous.reopen(source).map_err(map_archive_error)?;
    budget.charge_source_catalog(&catalog)?;
    charge_reopen_work(&catalog, budget)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

/// Account for the aggregate scans and retained metadata performed by
/// [`Package::from_source_catalog`] before allowing that parser to allocate
/// its semantic snapshot.  The parser owns the concrete allocations, while
/// this adapter owns the transaction-wide ceiling.
pub(crate) fn charge_reopen_work(
    source: &SourceCatalog,
    budget: &mut WireBudget,
) -> Result<(), Error> {
    let components = source.components();
    budget.charge_payload_work(components.len())?;
    for component in components.iter() {
        let archive = component.archive();
        budget.charge_payload_work(archive.objects.len())?;
        for object in &archive.objects {
            budget.charge_payload_work(object.messages.len())?;
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                budget.charge_payload_work(message.data.len())?;
                if message.type_ == ROOT_MESSAGE_TYPE {
                    // Root projection performs strict and lazy passes.
                    budget.charge_payload_work(message.data.len())?;
                } else if matches!(message.type_, 2_001 | 2_022) {
                    // Body storage validation performs a strict tree pass,
                    // a lazy projection pass, and materialization checks.
                    budget.charge_payload_work(message.data.len())?;
                    budget.charge_payload_work(message.data.len())?;
                }
                budget.charge_payload_work(
                    info.field_infos
                        .len()
                        .saturating_add(info.object_references.len())
                        .saturating_add(info.data_references.len()),
                )?;
            }
        }
    }
    Ok(())
}

fn bytes_equal_with_budget(
    left: &[u8],
    right: &[u8],
    budget: &mut WireBudget,
) -> Result<bool, Error> {
    if left.len() != right.len() {
        return Ok(false);
    }
    budget.charge_payload_work(left.len())?;
    Ok(left == right)
}

fn fingerprint(bytes: &[u8], budget: &mut WireBudget) -> Result<u64, Error> {
    budget.charge_input_source(bytes)?;
    budget.charge_payload_work(bytes.len())?;
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    Ok(value)
}

pub(crate) struct WireBudget {
    limits: WireLimits,
    physical_limits: Limits,
    total_bytes: usize,
    total_fields: usize,
    total_work: usize,
    payload_work: usize,
    payload_references: usize,
    input_bytes: usize,
    output_bytes: usize,
    entries: usize,
    total_entry_bytes: usize,
    package_bytes: usize,
    payload_bytes: usize,
    total_payload_bytes: usize,
    payload_objects: usize,
    payload_messages: usize,
    payload_items: usize,
    source_keys: [Option<(usize, usize)>; 2],
    catalog_keys: [Option<(usize, usize)>; 2],
    source_key_count: usize,
    catalog_key_count: usize,
    maximum_payload_work: usize,
    maximum_payload_references: usize,
}

impl WireBudget {
    pub(crate) fn new(physical_limits: Limits) -> Result<Self, Error> {
        let archive = physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let common = WireLimits::default();
        let input = archive
            .max_message_bytes()
            .min(common.max_input_bytes())
            .max(1);
        let fields = archive.max_header_fields().min(common.max_fields()).max(1);
        let nesting = archive
            .max_header_nesting()
            .min(common.max_nesting())
            .max(1);
        let wire_limits = common
            .with_input_bytes(input)
            .and_then(|value| value.with_fields(fields))
            .and_then(|value| {
                value.with_output_bytes(
                    archive
                        .max_archive_bytes()
                        .min(common.max_output_bytes())
                        .max(1),
                )
            })
            .and_then(|value| value.with_nesting(nesting))
            .map_err(map_wire_error)?;
        let maximum_payload_work = physical_limits
            .max_iwa_stream_bytes()
            .checked_mul(16)
            .unwrap_or(usize::MAX);
        let maximum_payload_references = archive
            .max_metadata_items()
            .checked_mul(physical_limits.max_entries().max(1))
            .unwrap_or(usize::MAX);
        Ok(Self {
            limits: wire_limits,
            physical_limits,
            total_bytes: 0,
            total_fields: 0,
            total_work: 0,
            payload_work: 0,
            payload_references: 0,
            input_bytes: 0,
            output_bytes: 0,
            entries: 0,
            total_entry_bytes: 0,
            package_bytes: 0,
            payload_bytes: 0,
            total_payload_bytes: 0,
            payload_objects: 0,
            payload_messages: 0,
            payload_items: 0,
            source_keys: [None; 2],
            catalog_keys: [None; 2],
            source_key_count: 0,
            catalog_key_count: 0,
            maximum_payload_work,
            maximum_payload_references,
        })
    }

    pub(crate) fn wire_limits(&self) -> WireLimits {
        self.limits
    }

    pub(crate) fn remaining_wire_fields(&self) -> usize {
        self.limits.max_fields().saturating_sub(self.total_fields)
    }

    pub(crate) fn remaining_wire_work(&self) -> usize {
        self.limits
            .max_rewrite_work()
            .saturating_sub(self.total_work)
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "candidate reservation needs the exact selected archive/message bounds"
    )]
    pub(crate) fn precharge_candidate_reopen(
        &mut self,
        source: &SourceCatalog,
        candidate_package_bound: usize,
        changed_component_index: usize,
        changed_entry_bound: usize,
        changed_archive_bound: usize,
        changed_object_index: usize,
        changed_message_index: usize,
        changed_message_bound: usize,
    ) -> Result<(), Error> {
        self.charge_input_bytes(candidate_package_bound)?;
        self.charge_entries(source.package().len())?;
        let changed_component = source
            .components()
            .get_index(changed_component_index)
            .ok_or(Error::InvalidSource)?;
        for entry in source.package().iter() {
            let entry_bytes = if entry.name() == changed_component.name() {
                changed_entry_bound
            } else {
                entry.data().len()
            };
            self.charge_entry_bytes(entry_bytes)?;
            self.charge_total_entry_bytes(entry_bytes)?;
            self.charge_payload_work(entry_bytes)?;
            let metadata_bytes = entry
                .raw_name()
                .len()
                .checked_add(entry.metadata().local().name().len())
                .and_then(|value| value.checked_add(entry.metadata().local().extra().len()))
                .and_then(|value| value.checked_add(entry.metadata().local().comment().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().name().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().extra().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().comment().len()))
                .ok_or(Error::Limit {
                    kind: LimitKind::PackageBytes,
                    observed: u64::MAX,
                    maximum: self.physical_limits.max_input_bytes(),
                })?;
            self.charge_package_bytes(metadata_bytes)?;
            self.charge_payload_work(metadata_bytes)?;
        }

        let components = source.components();
        self.charge_payload_work(components.len())?;
        for (component_index, component) in components.iter().enumerate() {
            let archive = component.archive();
            let archive_bytes = if component_index == changed_component_index {
                changed_archive_bound
            } else {
                parsed_archive_source_length(archive)?
            };
            self.charge_archive_inventory(archive_bytes, archive)?;
            self.charge_payload_work(archive.objects.len())?;
            for (object_index, object) in archive.objects.iter().enumerate() {
                self.charge_payload_work(object.messages.len())?;
                for (message_index, (message, info)) in object
                    .messages
                    .iter()
                    .zip(&object.archive_info.message_infos)
                    .enumerate()
                {
                    let message_bytes = if component_index == changed_component_index
                        && object_index == changed_object_index
                        && message_index == changed_message_index
                    {
                        changed_message_bound
                    } else {
                        message.data.len()
                    };
                    self.charge_payload_work(message_bytes)?;
                    if message.type_ == ROOT_MESSAGE_TYPE {
                        self.charge_payload_work(message_bytes)?;
                    } else if matches!(message.type_, 2_001 | 2_022) {
                        self.charge_payload_work(message_bytes)?;
                        self.charge_payload_work(message_bytes)?;
                    }
                    self.charge_payload_work(
                        info.field_infos
                            .len()
                            .saturating_add(info.object_references.len())
                            .saturating_add(info.data_references.len()),
                    )?;
                }
            }
        }
        Ok(())
    }

    pub(crate) fn maximum_payload_references(&self) -> usize {
        self.maximum_payload_references
    }

    pub(crate) fn remaining_payload_references(&self) -> usize {
        self.maximum_payload_references
            .saturating_sub(self.payload_references)
    }

    pub(crate) fn charge_codec_report(
        &mut self,
        fields: usize,
        work_bytes: usize,
        max_depth: u32,
        references: usize,
    ) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.total_fields,
            fields,
            usize_as_u64(self.limits.max_fields()),
            LimitKind::WireFields,
        )?;
        self.charge_wire_work(work_bytes)?;
        if usize::try_from(max_depth).unwrap_or(usize::MAX) > self.limits.max_nesting() {
            return Err(Error::Limit {
                kind: LimitKind::WireNesting,
                observed: u64::from(max_depth),
                maximum: usize_as_u64(self.limits.max_nesting()),
            });
        }
        self.charge_payload_references(references)
    }

    fn charge_counter(
        counter: &mut usize,
        amount: usize,
        maximum: u64,
        kind: LimitKind,
    ) -> Result<(), Error> {
        let observed = counter.checked_add(amount).ok_or(Error::Limit {
            kind,
            observed: u64::MAX,
            maximum,
        })?;
        let observed_u64 = usize_as_u64(observed);
        if observed_u64 > maximum {
            return Err(Error::Limit {
                kind,
                observed: observed_u64,
                maximum,
            });
        }
        *counter = observed;
        Ok(())
    }

    fn charge_wire_work(&mut self, amount: usize) -> Result<(), Error> {
        let work = self.total_work.checked_add(amount).ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_as_u64(self.limits.max_rewrite_work()),
        })?;
        if work > self.limits.max_rewrite_work() {
            return Err(Error::Limit {
                kind: LimitKind::WireWork,
                observed: usize_as_u64(work),
                maximum: usize_as_u64(self.limits.max_rewrite_work()),
            });
        }
        self.total_work = work;
        Ok(())
    }

    pub(crate) fn charge_payload_work(&mut self, amount: usize) -> Result<(), Error> {
        let work = self.payload_work.checked_add(amount).ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_as_u64(self.maximum_payload_work),
        })?;
        if work > self.maximum_payload_work {
            return Err(Error::Limit {
                kind: LimitKind::WireWork,
                observed: usize_as_u64(work),
                maximum: usize_as_u64(self.maximum_payload_work),
            });
        }
        self.charge_wire_work(amount)?;
        self.payload_work = work;
        Ok(())
    }

    #[cfg(test)]
    pub(crate) fn exhaust_payload_work_for_test(&mut self) {
        self.payload_work = self.maximum_payload_work;
    }

    pub(crate) fn charge_payload_references(&mut self, amount: usize) -> Result<(), Error> {
        self.charge_wire_work(amount)?;
        Self::charge_counter(
            &mut self.payload_references,
            amount,
            usize_as_u64(self.maximum_payload_references),
            LimitKind::PayloadReferences,
        )
    }

    fn charge_input_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.input_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            LimitKind::InputBytes,
        )
    }

    /// Admit a caller-owned selector name before retaining its bytes.
    ///
    /// This is intentionally limited to scalar byte ceilings.  Rooted table
    /// discovery remains deferred until the edit is first used.
    pub(crate) fn admit_selector_name(&mut self, amount: usize) -> Result<(), Error> {
        self.charge_input_bytes(amount)?;
        self.charge_payload_bytes(amount)?;
        self.charge_total_payload_bytes(amount)
    }

    pub(crate) fn charge_output_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.output_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            LimitKind::OutputBytes,
        )
    }

    fn charge_entries(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.entries,
            amount,
            usize_as_u64(self.physical_limits.max_entries()),
            LimitKind::Entries,
        )
    }

    fn charge_entry_bytes(&mut self, amount: usize) -> Result<(), Error> {
        let observed = usize_as_u64(amount);
        let maximum = self.physical_limits.max_entry_bytes();
        if observed > maximum {
            return Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed,
                maximum,
            });
        }
        Ok(())
    }

    fn charge_total_entry_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.total_entry_bytes,
            amount,
            self.physical_limits.max_total_bytes(),
            LimitKind::TotalEntryBytes,
        )
    }

    fn charge_package_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.package_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            LimitKind::PackageBytes,
        )
    }

    pub(crate) fn charge_payload_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.payload_bytes,
            amount,
            usize_as_u64(self.physical_limits.max_iwa_stream_bytes()),
            LimitKind::PayloadBytes,
        )
    }

    pub(crate) fn charge_total_payload_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.total_payload_bytes,
            amount,
            self.physical_limits.max_total_bytes(),
            LimitKind::TotalPayloadBytes,
        )
    }

    pub(crate) fn charge_payload_objects(&mut self, amount: usize) -> Result<(), Error> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_objects,
            amount,
            usize_as_u64(limits.max_objects()),
            LimitKind::PayloadObjects,
        )
    }

    pub(crate) fn charge_payload_messages(&mut self, amount: usize) -> Result<(), Error> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_messages,
            amount,
            usize_as_u64(limits.max_messages()),
            LimitKind::PayloadMessages,
        )
    }

    pub(crate) fn charge_payload_items(&mut self, amount: usize) -> Result<(), Error> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_items,
            amount,
            usize_as_u64(limits.max_metadata_items()),
            LimitKind::PayloadItems,
        )
    }

    fn remember_source_key(&mut self, bytes: &[u8]) -> bool {
        let key = (bytes.as_ptr() as usize, bytes.len());
        if self.source_keys[..self.source_key_count].contains(&Some(key)) {
            return false;
        }
        if self.source_key_count < self.source_keys.len() {
            self.source_keys[self.source_key_count] = Some(key);
            self.source_key_count += 1;
        }
        true
    }

    fn remember_catalog_key(&mut self, bytes: &[u8]) -> bool {
        let key = (bytes.as_ptr() as usize, bytes.len());
        if self.catalog_keys[..self.catalog_key_count].contains(&Some(key)) {
            return false;
        }
        if self.catalog_key_count < self.catalog_keys.len() {
            self.catalog_keys[self.catalog_key_count] = Some(key);
            self.catalog_key_count += 1;
        }
        true
    }

    pub(crate) fn charge_input_source(&mut self, bytes: &[u8]) -> Result<(), Error> {
        if self.remember_source_key(bytes) {
            self.charge_input_bytes(bytes.len())?;
        }
        Ok(())
    }

    pub(crate) fn charge_source_catalog(&mut self, source: &SourceCatalog) -> Result<(), Error> {
        let source_bytes = source.source_bytes();
        self.charge_input_source(source_bytes)?;
        if !self.remember_catalog_key(source_bytes) {
            return Ok(());
        }
        let package = source.package();
        self.charge_entries(package.len())?;
        for entry in package.iter() {
            let entry_bytes =
                usize::try_from(entry.metadata().uncompressed_size()).map_err(|_| {
                    Error::Limit {
                        kind: LimitKind::EntryBytes,
                        observed: u64::MAX,
                        maximum: self.physical_limits.max_entry_bytes(),
                    }
                })?;
            // EntryBytes is a per-member ceiling. The aggregate is charged
            // separately so several members can each fit individually while
            // still being bounded by TotalEntryBytes.
            self.charge_entry_bytes(entry_bytes)?;
            self.charge_total_entry_bytes(entry_bytes)?;
            self.charge_payload_work(entry.data().len())?;
            let metadata_bytes = entry
                .raw_name()
                .len()
                .checked_add(entry.metadata().local().name().len())
                .and_then(|value| value.checked_add(entry.metadata().local().extra().len()))
                .and_then(|value| value.checked_add(entry.metadata().local().comment().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().name().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().extra().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().comment().len()))
                .ok_or(Error::Limit {
                    kind: LimitKind::PackageBytes,
                    observed: u64::MAX,
                    maximum: self.physical_limits.max_input_bytes(),
                })?;
            self.charge_package_bytes(metadata_bytes)?;
            self.charge_payload_work(metadata_bytes)?;
        }
        Ok(())
    }

    /// Charge the parsed IWA storage retained by a component-only reader.
    ///
    /// A shared component catalog has already passed physical ZIP, Snappy, and
    /// neutral archive limits at its originating ingress.  The focused query
    /// still needs an operation-local accounting point for the immutable
    /// archive bytes it can inspect, while object/message/index work is
    /// charged by the rooted ownership walk itself.
    pub(crate) fn charge_component_catalog(
        &mut self,
        components: &ComponentCatalog,
    ) -> Result<(), Error> {
        self.charge_payload_work(components.len())?;
        for component in components.iter() {
            let archive_bytes = parsed_archive_source_length(component.archive())?;
            if archive_bytes > self.physical_limits.max_iwa_stream_bytes() {
                return Err(Error::Limit {
                    kind: LimitKind::PayloadBytes,
                    observed: usize_as_u64(archive_bytes),
                    maximum: usize_as_u64(self.physical_limits.max_iwa_stream_bytes()),
                });
            }
            self.charge_total_payload_bytes(archive_bytes)?;
            self.charge_payload_work(archive_bytes)?;
        }
        Ok(())
    }

    fn charge_archive_inventory(
        &mut self,
        archive_bytes: usize,
        archive: &Archive,
    ) -> Result<(), Error> {
        self.charge_payload_bytes(archive_bytes)?;
        self.charge_total_payload_bytes(archive_bytes)?;
        self.charge_payload_objects(archive.objects.len())?;
        let mut message_count = 0usize;
        let mut item_count = 0usize;
        for object in &archive.objects {
            message_count =
                message_count
                    .checked_add(object.messages.len())
                    .ok_or(Error::Limit {
                        kind: LimitKind::PayloadMessages,
                        observed: u64::MAX,
                        maximum: u64::MAX,
                    })?;
            for info in &object.archive_info.message_infos {
                item_count = item_count
                    .checked_add(1)
                    .and_then(|value| value.checked_add(info.field_infos.len()))
                    .and_then(|value| value.checked_add(info.object_references.len()))
                    .and_then(|value| value.checked_add(info.data_references.len()))
                    .ok_or(Error::Limit {
                        kind: LimitKind::PayloadItems,
                        observed: u64::MAX,
                        maximum: u64::MAX,
                    })?;
            }
        }
        self.charge_payload_messages(message_count)?;
        self.charge_payload_items(item_count)?;
        let work = archive_bytes
            .checked_add(archive.objects.len())
            .and_then(|value| value.checked_add(message_count))
            .and_then(|value| value.checked_add(item_count))
            .ok_or(Error::Limit {
                kind: LimitKind::WireWork,
                observed: u64::MAX,
                maximum: usize_as_u64(self.maximum_payload_work),
            })?;
        self.charge_payload_work(work)
    }

    pub(crate) fn charge_sort_work(&mut self, length: usize) -> Result<(), Error> {
        let levels = if length <= 1 {
            0
        } else {
            (usize::BITS - (length - 1).leading_zeros()) as usize
        };
        let amount = length.checked_mul(levels).ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_as_u64(self.maximum_payload_work),
        })?;
        self.charge_payload_work(amount)
    }

    pub(crate) fn parse<'a>(
        &mut self,
        source: &'a [u8],
        depth: usize,
    ) -> Result<WireView<'a>, Error> {
        if depth > self.limits.max_nesting() {
            return Err(Error::Limit {
                kind: LimitKind::WireNesting,
                observed: usize_as_u64(depth),
                maximum: usize_as_u64(self.limits.max_nesting()),
            });
        }
        let bytes = self
            .total_bytes
            .checked_add(source.len())
            .ok_or(Error::Limit {
                kind: LimitKind::WireBytes,
                observed: u64::MAX,
                maximum: usize_as_u64(self.limits.max_input_bytes()),
            })?;
        if bytes > self.limits.max_input_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::WireBytes,
                observed: usize_as_u64(bytes),
                maximum: usize_as_u64(self.limits.max_input_bytes()),
            });
        }
        let view = WireView::parse(source, self.limits).map_err(map_wire_error)?;
        let fields = self
            .total_fields
            .checked_add(view.len())
            .ok_or(Error::Limit {
                kind: LimitKind::WireFields,
                observed: u64::MAX,
                maximum: usize_as_u64(self.limits.max_fields()),
            })?;
        if fields > self.limits.max_fields() {
            return Err(Error::Limit {
                kind: LimitKind::WireFields,
                observed: usize_as_u64(fields),
                maximum: usize_as_u64(self.limits.max_fields()),
            });
        }
        let work = source.len().checked_add(view.len()).ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_as_u64(self.limits.max_rewrite_work()),
        })?;
        self.charge_wire_work(work)?;
        self.total_bytes = bytes;
        self.total_fields = fields;
        Ok(view)
    }
}

pub(crate) fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadItems,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadObjects,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        PackageError::Semantic(error) => match error {
            crate::Error::TooManySections { actual, limit }
            | crate::Error::TooManyBodyStorages { actual, limit } => Error::Limit {
                kind: LimitKind::PayloadItems,
                observed: usize_as_u64(actual),
                maximum: usize_as_u64(limit),
            },
            crate::Error::TextTooLarge { observed, limit } => Error::Limit {
                kind: LimitKind::PayloadBytes,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(limit),
            },
            crate::Error::InvalidSectionIndex { .. } => Error::InvalidSource,
        },
        PackageError::SectionNamesTooLarge { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        PackageError::InvalidFormat(_)
        | PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::UnexpectedFormat { .. }
        | PackageError::NotOfficeFile
        | PackageError::MissingIdentity
        | PackageError::InvalidLimit { .. }
        | PackageError::Limit { .. } => Error::InvalidSource,
    }
}

pub(crate) fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
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
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::PackageBytes,
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
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Reassembly(_) => Error::Refused(Refusal::UnsupportedSource),
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
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    LimitKind::PayloadBytes
                },
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::PayloadItems,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_core::FieldInfo;
    use litchi_iwa_protos::tswp;
    use prost::Message as _;

    fn table_info_payload() -> Vec<u8> {
        let mut reference = Vec::new();
        litchi_iwa_common::wire::append_varint_field(&mut reference, 1, 42)
            .expect("reference field");

        let mut drawable = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut drawable,
            2,
            &reference,
            WireLimits::DEFAULT,
        )
        .expect("drawable parent field");

        let mut table_info = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut table_info,
            1,
            &drawable,
            WireLimits::DEFAULT,
        )
        .expect("table-info super field");
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut table_info,
            2,
            &reference,
            WireLimits::DEFAULT,
        )
        .expect("table-info model field");
        table_info
    }

    fn budget_with_wire_limits(
        input_bytes: usize,
        fields: usize,
        rewrite_work: usize,
    ) -> WireBudget {
        let mut budget = WireBudget::new(Limits::default()).expect("default wire budget");
        budget.limits = WireLimits::default()
            .with_input_bytes(input_bytes)
            .and_then(|limits| limits.with_fields(fields))
            .and_then(|limits| limits.with_rewrite_work(rewrite_work))
            .expect("test wire limits");
        budget
    }

    #[test]
    fn borrowed_body_placeholder_scan_stops_before_an_oversized_trailing_run() {
        let trailing = "x".repeat(32 * 1024);
        let payload = tswp::StorageArchive {
            text: vec!["\u{fffc}".to_owned(), trailing],
            ..tswp::StorageArchive::default()
        }
        .encode_to_vec();
        let rewrite_limits = super::super::storage_rewrite_limits(Limits::default())
            .unwrap_or_else(|_| panic!("default text-wire limits"));
        let validation = litchi_iwa_text_wire::validate_storage(&payload, rewrite_limits)
            .expect("oversized trailing run is within the test source ceiling");
        let exact_work = validation
            .validation_work()
            .checked_add(1)
            .expect("test work fits usize");
        let mut budget = budget_with_wire_limits(payload.len(), 1_024, exact_work);
        budget.maximum_payload_work = exact_work;
        let entry = BodyTableEntry {
            character_index: 0,
            identifier: NonZeroU64::new(1).expect("test identifier is non-zero"),
        };

        validate_body_placeholders_in_wire(&payload, &[entry], Limits::default(), &mut budget)
            .expect("reader should stop after the selected placeholder");
        assert_eq!(budget.payload_work, exact_work);
    }

    #[test]
    fn borrowed_body_placeholder_scan_rejects_a_non_placeholder_at_the_selected_index() {
        let payload = tswp::StorageArchive {
            text: vec!["ordinary text".to_owned()],
            ..tswp::StorageArchive::default()
        }
        .encode_to_vec();
        let entry = BodyTableEntry {
            character_index: 0,
            identifier: NonZeroU64::new(1).expect("test identifier is non-zero"),
        };
        let mut budget = budget_with_wire_limits(payload.len(), 1_024, 1_024 * 1_024);

        assert_eq!(
            validate_body_placeholders_in_wire(&payload, &[entry], Limits::default(), &mut budget,),
            Err(Error::InvalidSource)
        );
    }

    #[test]
    fn body_table_lock_edit_debug_redacts_native_target_details() {
        let package = Package::from_bytes(include_bytes!(
            "../../../../test-data/iwork/pages/basic.pages"
        ))
        .expect("Pages fixture");
        let target = BodyTableTarget {
            table_position: 17,
            table_name: "private-table-name".into(),
            table_rows: 5,
            table_columns: 4,
            anchor_character_index: 0,
            sheet_identifier: NonZeroU64::new(0xfeed_babe).expect("identifier"),
            sheet_component_index: 1,
            sheet_object_index: 2,
            sheet_message_index: 4,
            sheet_message_type: 10_001,
            attachment_identifier: NonZeroU64::new(0xfeed_face).expect("identifier"),
            attachment_component_index: 2,
            attachment_object_index: 3,
            attachment_message_index: 5,
            drawable_identifier: NonZeroU64::new(0xcafe_babe).expect("identifier"),
            model_identifier: NonZeroU64::new(0xdead_beef).expect("identifier"),
            model_component_index: 7,
            model_object_index: 11,
            model_message_index: 13,
            model_message_type: 6_000,
            component_index: 19,
            object_index: 23,
            message_index: 29,
            message_type: 6_003,
            info_message_index: 29,
            body_component_index: 31,
            body_object_index: 37,
            body_message_index: 41,
            body_message_type: 2_001,
            body_identifier: NonZeroU64::new(0xabad_1dea).expect("identifier"),
            explicit_locked: Some(true),
        };
        let edit = Edit {
            source: &package,
            target,
            before: State::Unlocked,
            state: State::Locked,
        };

        assert_eq!(
            format!("{edit:?}"),
            "Edit { before: Unlocked, state: Locked, .. }"
        );
    }

    #[test]
    fn parsed_archive_source_length_matches_exact_stream_extent() {
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: vec![1, 2, 3],
                    }],
                    Default::default(),
                )
                .expect("first object"),
                ArchiveObject::new(
                    2,
                    vec![RawMessage {
                        type_: 2_001,
                        data: vec![4, 5, 6, 7],
                    }],
                    Default::default(),
                )
                .expect("second object"),
            ],
        };
        let source = archive.to_bytes(Default::default()).expect("archive bytes");
        let parsed = Archive::parse(&source, Default::default()).expect("parsed archive");
        assert_eq!(
            parsed_archive_source_length(&parsed).expect("source extent"),
            source.len()
        );
        assert_eq!(
            parsed_archive_source_length(&Archive {
                objects: Vec::new()
            })
            .expect("empty extent"),
            0
        );
    }

    #[test]
    fn exact_byte_comparisons_charge_only_equal_length_scans() {
        let mut budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        assert!(!bytes_equal_with_budget(b"abc", b"abcd", &mut budget).expect("different lengths"));
        assert_eq!(budget.payload_work, 0);
        assert!(bytes_equal_with_budget(b"abc", b"abc", &mut budget).expect("equal bytes"));
        assert_eq!(budget.payload_work, 3);
        assert!(!bytes_equal_with_budget(b"abc", b"abd", &mut budget).expect("different bytes"));
        assert_eq!(budget.payload_work, 6);
    }

    #[test]
    fn table_info_payload_bytes_fields_and_work_are_aggregate() {
        let payload = table_info_payload();
        let body_identifier = NonZeroU64::new(42).expect("non-zero body identifier");
        let mut baseline = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        decode_table_info(&payload, &mut baseline, body_identifier).expect("one table-info decode");
        let one_bytes = baseline.total_bytes;
        let one_fields = baseline.total_fields;
        let one_work = baseline.total_work;

        let dimensions = [
            (LimitKind::WireBytes, one_bytes * 2 - 1, 1024, 1024 * 1024),
            (LimitKind::WireFields, 1024, one_fields * 2 - 1, 1024 * 1024),
            (LimitKind::WireWork, 1024, 1024, one_work * 2 - 1),
        ];

        for (kind, input_bytes, fields, rewrite_work) in dimensions {
            let mut budget = budget_with_wire_limits(input_bytes, fields, rewrite_work);
            decode_table_info(&payload, &mut budget, body_identifier)
                .expect("first bounded table-info decode");
            assert!(matches!(
                decode_table_info(&payload, &mut budget, body_identifier),
                Err(Error::Limit { kind: observed, .. }) if observed == kind
            ));
        }
    }

    fn tiny_physical_limits() -> Limits {
        let archive = litchi_iwa_core::Limits::default()
            .with_archive_bytes(8)
            .and_then(|limits| limits.with_objects(2))
            .and_then(|limits| limits.with_messages(4))
            .and_then(|limits| limits.with_metadata_items(4))
            .expect("test archive limits");
        Limits::new(16, 2, 8, 16, 8)
            .expect("test physical limits")
            .with_archive_limits(archive)
            .expect("test physical profile")
    }

    fn assert_physical_boundary(
        kind: LimitKind,
        maximum: usize,
        charge: fn(&mut WireBudget, usize) -> Result<(), Error>,
    ) {
        let mut budget = WireBudget::new(tiny_physical_limits()).expect("test wire budget");
        charge(&mut budget, maximum).expect("maximum is accepted");
        assert!(matches!(
            charge(&mut budget, 1),
            Err(Error::Limit { kind: observed, .. }) if observed == kind
        ));
    }

    #[test]
    fn physical_budget_boundaries_reject_max_plus_one() {
        assert_physical_boundary(LimitKind::InputBytes, 16, WireBudget::charge_input_bytes);
        assert_physical_boundary(LimitKind::OutputBytes, 16, WireBudget::charge_output_bytes);
        assert_physical_boundary(LimitKind::Entries, 2, WireBudget::charge_entries);
        assert_physical_boundary(
            LimitKind::TotalEntryBytes,
            16,
            WireBudget::charge_total_entry_bytes,
        );
        assert_physical_boundary(
            LimitKind::PackageBytes,
            16,
            WireBudget::charge_package_bytes,
        );
        assert_physical_boundary(LimitKind::PayloadBytes, 8, WireBudget::charge_payload_bytes);
        assert_physical_boundary(
            LimitKind::TotalPayloadBytes,
            16,
            WireBudget::charge_total_payload_bytes,
        );
        assert_physical_boundary(
            LimitKind::PayloadObjects,
            2,
            WireBudget::charge_payload_objects,
        );
        assert_physical_boundary(
            LimitKind::PayloadMessages,
            4,
            WireBudget::charge_payload_messages,
        );
        assert_physical_boundary(LimitKind::PayloadItems, 4, WireBudget::charge_payload_items);
        assert_physical_boundary(
            LimitKind::PayloadReferences,
            8,
            WireBudget::charge_payload_references,
        );
    }

    #[test]
    fn entry_bytes_is_bounded_per_member() {
        let mut budget = WireBudget::new(tiny_physical_limits()).expect("test wire budget");
        budget
            .charge_entry_bytes(8)
            .expect("one member at the maximum is accepted");
        budget
            .charge_entry_bytes(8)
            .expect("a second member at the maximum is accepted");
        assert!(matches!(
            budget.charge_entry_bytes(9),
            Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: 9,
                maximum: 8,
            })
        ));
    }

    #[test]
    fn source_catalog_charges_multi_member_entry_bytes_aggregate_separately() {
        let limits = Limits::new(512, 4, 8, 16, 64).expect("test physical limits");
        let first = [0_u8; 8];
        let second = [1_u8; 8];
        let source = litchi_iwa_archive::package::to_bytes(
            [
                ("Data/first.bin", first.as_slice()),
                ("Data/second.bin", second.as_slice()),
            ],
            limits,
        )
        .expect("multi-member package");
        let catalog = SourceCatalog::from_bytes(&source, limits).expect("source catalog");
        let mut budget = WireBudget::new(limits).expect("test wire budget");

        budget
            .charge_source_catalog(&catalog)
            .expect("each member fits while the aggregate reaches its ceiling");
        assert_eq!(budget.total_entry_bytes, 16);
    }

    #[test]
    fn snappy_bound_covers_incompressible_multi_frame_output() {
        let mut state = 0x9e37_79b9_7f4a_7c15_u64;
        let mut input = vec![0_u8; SnappyStream::WRITE_CHUNK_SIZE * 16 + 17];
        for byte in &mut input {
            state ^= state << 7;
            state ^= state >> 9;
            state ^= state << 8;
            *byte = state as u8;
        }

        let bound = snappy_compressed_bound(input.len()).expect("checked Snappy bound");
        let compressed =
            SnappyStream::compress(&input, Default::default()).expect("Snappy compression");

        assert!(bound > input.len() + 64);
        assert!(compressed.len() <= bound);
        assert!(snappy_compressed_bound(usize::MAX).is_none());
    }

    #[test]
    fn deflate_bound_covers_incompressible_stored_blocks() {
        const MAX_STORED_BLOCK_BYTES: usize = u16::MAX as usize;
        let input_len = MAX_STORED_BLOCK_BYTES * 16 + 17;
        let mut state = 0x517c_c1b7_2722_0a95_u64;
        let mut deflated = Vec::new();
        let mut remaining = input_len;
        while remaining != 0 {
            let block_len = remaining.min(MAX_STORED_BLOCK_BYTES);
            let final_block = block_len == remaining;
            // Stored Deflate blocks begin byte-aligned after their three-bit
            // header and five zero padding bits.
            deflated.push(u8::from(final_block));
            let length = u16::try_from(block_len).expect("stored block length");
            deflated.extend_from_slice(&length.to_le_bytes());
            deflated.extend_from_slice(&(!length).to_le_bytes());
            for _ in 0..block_len {
                state ^= state << 7;
                state ^= state >> 9;
                state ^= state << 8;
                deflated.push(state as u8);
            }
            remaining -= block_len;
        }

        let bound = deflate_compressed_bound(input_len).expect("checked Deflate bound");
        assert!(deflated.len() > input_len + 64);
        assert!(deflated.len() <= bound);
        assert!(deflate_compressed_bound(usize::MAX).is_none());
    }

    #[test]
    fn role_identifiers_reject_every_pairwise_alias_including_root() {
        let baseline = [1, 42, 100, 200, 300];
        for left in 0..baseline.len() {
            for right in (left + 1)..baseline.len() {
                let mut aliases = baseline;
                aliases[right] = aliases[left];
                assert!(role_identifiers_are_aliased(
                    aliases[0], aliases[1], aliases[2], aliases[3], aliases[4],
                ));
            }
        }
        assert!(!role_identifiers_are_aliased(
            baseline[0],
            baseline[1],
            baseline[2],
            baseline[3],
            baseline[4],
        ));
    }

    #[test]
    fn body_reference_metadata_accepts_unrelated_paths_but_keeps_selected_path_strict() {
        let mut message =
            litchi_iwa_core::MessageInfo::new(2_001, 0, litchi_iwa_core::Limits::default())
                .unwrap();
        message.object_references = vec![100, 900, 901];
        let mut selected = FieldInfo::new(
            litchi_iwa_core::FieldPath::new(
                vec![TABLE_BODY_FIELD],
                litchi_iwa_core::Limits::default(),
            )
            .unwrap(),
            litchi_iwa_core::Limits::default(),
        )
        .unwrap();
        selected.object_references.push(100);
        let mut sections = FieldInfo::new(
            litchi_iwa_core::FieldPath::new(vec![17], litchi_iwa_core::Limits::default()).unwrap(),
            litchi_iwa_core::Limits::default(),
        )
        .unwrap();
        sections.object_references.extend([900, 901]);
        message.field_infos = vec![selected, sections];

        let mut budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        assert!(
            message_declares_reference_prefix(&message, 100, &[TABLE_BODY_FIELD], &mut budget)
                .expect("unrelated aggregate references use their own path")
        );

        message.field_infos[0].path =
            litchi_iwa_core::FieldPath::new(vec![17], litchi_iwa_core::Limits::default()).unwrap();
        let mut strict_budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        assert!(matches!(
            message_declares_reference_prefix(
                &message,
                100,
                &[TABLE_BODY_FIELD],
                &mut strict_budget,
            ),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn data_reference_scans_are_bounded_before_semantic_checks() {
        let mut message_data =
            litchi_iwa_core::MessageInfo::new(2_002, 0, litchi_iwa_core::Limits::default())
                .unwrap();
        message_data.object_references = vec![100];
        message_data.data_references = vec![100];
        message_data.field_infos = vec![
            FieldInfo::new(
                litchi_iwa_core::FieldPath::new(
                    vec![TABLE_BODY_FIELD],
                    litchi_iwa_core::Limits::default(),
                )
                .unwrap(),
                litchi_iwa_core::Limits::default(),
            )
            .unwrap(),
        ];

        let mut message_budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        message_budget.maximum_payload_references = 1;
        assert!(matches!(
            message_declares_reference(
                &message_data,
                100,
                &[TABLE_BODY_FIELD],
                true,
                false,
                &mut message_budget,
            ),
            Err(Error::Limit {
                kind: LimitKind::PayloadReferences,
                ..
            })
        ));

        let mut prefix_budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        prefix_budget.maximum_payload_references = 1;
        assert!(matches!(
            message_declares_reference_prefix(
                &message_data,
                100,
                &[TABLE_BODY_FIELD],
                &mut prefix_budget,
            ),
            Err(Error::Limit {
                kind: LimitKind::PayloadReferences,
                ..
            })
        ));

        let mut field_data =
            litchi_iwa_core::MessageInfo::new(2_003, 0, litchi_iwa_core::Limits::default())
                .unwrap();
        field_data.object_references = vec![100];
        let mut field = FieldInfo::new(
            litchi_iwa_core::FieldPath::new(
                vec![TABLE_BODY_FIELD],
                litchi_iwa_core::Limits::default(),
            )
            .unwrap(),
            litchi_iwa_core::Limits::default(),
        )
        .unwrap();
        field.object_references = vec![100];
        field.data_references = vec![100];
        field_data.field_infos = vec![field];

        let mut field_budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        field_budget.maximum_payload_references = 2;
        assert!(matches!(
            message_declares_reference(
                &field_data,
                100,
                &[TABLE_BODY_FIELD],
                true,
                true,
                &mut field_budget,
            ),
            Err(Error::Limit {
                kind: LimitKind::PayloadReferences,
                ..
            })
        ));

        let mut field_prefix_budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        field_prefix_budget.maximum_payload_references = 2;
        assert!(matches!(
            message_declares_reference_prefix(
                &field_data,
                100,
                &[TABLE_BODY_FIELD],
                &mut field_prefix_budget,
            ),
            Err(Error::Limit {
                kind: LimitKind::PayloadReferences,
                ..
            })
        ));
    }

    #[test]
    fn aggregate_reference_budget_is_charged_before_declaration_reserve() {
        let mut message =
            litchi_iwa_core::MessageInfo::new(2_001, 0, litchi_iwa_core::Limits::default())
                .unwrap();
        let mut field = FieldInfo::new(
            litchi_iwa_core::FieldPath::new(
                vec![TABLE_BODY_FIELD],
                litchi_iwa_core::Limits::default(),
            )
            .unwrap(),
            litchi_iwa_core::Limits::default(),
        )
        .unwrap();
        field.object_references.push(100);
        message.field_infos.push(field);

        let mut budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        budget.maximum_payload_references = 0;
        assert!(matches!(
            message_declares_reference_prefix(&message, 100, &[TABLE_BODY_FIELD], &mut budget,),
            Err(Error::Limit {
                kind: LimitKind::PayloadReferences,
                ..
            })
        ));
        // One unit is charged for the FieldInfo inventory and one for its
        // nested reference scan; the declaration-map reservation must not
        // charge or allocate before the reference ceiling rejects the input.
        assert_eq!(budget.total_work, 2);
    }

    #[test]
    fn empty_body_inventory_charges_references_before_declaration_reserve() {
        let mut message =
            litchi_iwa_core::MessageInfo::new(2_001, 0, litchi_iwa_core::Limits::default())
                .unwrap();
        let mut field = FieldInfo::new(
            litchi_iwa_core::FieldPath::new(
                vec![TABLE_BODY_FIELD],
                litchi_iwa_core::Limits::default(),
            )
            .unwrap(),
            litchi_iwa_core::Limits::default(),
        )
        .unwrap();
        field.object_references.push(100);
        message.object_references.push(100);
        message.field_infos.push(field);

        let mut budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        budget.maximum_payload_references = 0;
        assert!(matches!(
            validate_body_table_ownership(&message, &[], &mut budget),
            Err(Error::Limit {
                kind: LimitKind::PayloadReferences,
                ..
            })
        ));
        assert_eq!(budget.total_work, 1);
    }

    #[test]
    fn table_info_rewrite_bound_covers_optional_lock_growth() {
        assert_eq!(
            table_info_rewrite_bound(10, true),
            Some(10 + MAX_VARINT_BYTES * 2)
        );
        assert_eq!(
            table_info_rewrite_bound(10, false),
            Some(10 + 2 + MAX_VARINT_BYTES * 2)
        );
        assert!(table_info_rewrite_bound(usize::MAX, false).is_none());
    }

    #[test]
    fn is_noop_requires_semantic_and_exact_artifact_identity() {
        let identifier = NonZeroU64::new(1).expect("non-zero test identifier");
        let target = BodyTableTarget {
            table_position: 0,
            table_name: "test-table".into(),
            table_rows: 5,
            table_columns: 4,
            anchor_character_index: 0,
            sheet_identifier: identifier,
            sheet_component_index: 0,
            sheet_object_index: 0,
            sheet_message_index: 0,
            sheet_message_type: 0,
            attachment_identifier: identifier,
            attachment_component_index: 0,
            attachment_object_index: 0,
            attachment_message_index: 0,
            drawable_identifier: identifier,
            model_identifier: identifier,
            model_component_index: 0,
            model_object_index: 0,
            model_message_index: 0,
            model_message_type: 0,
            component_index: 0,
            object_index: 0,
            message_index: 0,
            message_type: 0,
            info_message_index: 0,
            body_component_index: 0,
            body_object_index: 0,
            body_message_index: 0,
            body_message_type: 0,
            body_identifier: identifier,
            explicit_locked: Some(false),
        };
        let source: Arc<[u8]> = Arc::from(&b"source"[..]);
        let different_target: Arc<[u8]> = Arc::from(&b"target"[..]);

        let semantic_only = Patch {
            source: Arc::clone(&source),
            target: different_target,
            source_fingerprint: 7,
            target_fingerprint: 7,
            proof: target.clone(),
            before: State::Unlocked,
            after: State::Unlocked,
        };
        assert!(!semantic_only.is_noop());

        let equivalent_target: Arc<[u8]> = Arc::from(&b"source"[..]);
        let exact = Patch {
            source: Arc::clone(&source),
            target: equivalent_target,
            source_fingerprint: 7,
            target_fingerprint: 7,
            proof: target,
            before: State::Unlocked,
            after: State::Unlocked,
        };
        assert!(exact.is_noop());
    }
}
