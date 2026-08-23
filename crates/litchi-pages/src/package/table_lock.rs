//! Immutable, exact-source transactions for Pages body-table locks.

use std::collections::HashMap;
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireFieldView, WireView, patch_varint_field, transform_length_delimited_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use thiserror::Error;

use super::{
    Package, PackageError, RootReferences, decode_body_storage, effective_text_limit,
    root_references_with_limits,
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
const TABLE_MODEL_NAME_FIELD: u32 = 8;
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
pub enum BodyTableLockLimitKind {
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

impl fmt::Display for BodyTableLockLimitKind {
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
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableLockError {
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
    /// Rooted semantic selection did not resolve to one supported native payload.
    #[error("the selected Pages body table has no unambiguous editable lock payload")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages body-table lock {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: BodyTableLockLimitKind,
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

/// A mutable semantic body-table lock state staged against one immutable
/// package snapshot.
#[derive(Debug)]
pub struct BodyTableLockEdit<'a> {
    source: &'a Package,
    target: BodyTableTarget,
    before: State,
    state: State,
}

impl BodyTableLockEdit<'_> {
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
    pub fn commit(self) -> Result<BodyTableLockCommit, BodyTableLockError> {
        let source = physical_source(self.source)?;
        let mut budget = WireBudget::new(source.limits())?;
        let source_bytes = source.shared_source();
        let source_fingerprint = fingerprint(source.source_bytes(), &mut budget)?;
        if self.before == self.state {
            return Ok(BodyTableLockCommit {
                package: self.source.snapshot(),
                patch: BodyTableLockPatch {
                    source: Arc::clone(&source_bytes),
                    target: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    proof: self.target,
                    before: self.before,
                    after: self.state,
                },
                diagnostics: BodyTableLockDiagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() {
            return Err(BodyTableLockError::UnsupportedSource);
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
        Ok(BodyTableLockCommit {
            package,
            patch: BodyTableLockPatch {
                source: source_bytes,
                target,
                source_fingerprint,
                target_fingerprint,
                proof: self.target,
                before: self.before,
                after: self.state,
            },
            diagnostics: BodyTableLockDiagnostics::published(),
        })
    }
}

/// Reversible patch bound to the exact source and target package artifacts.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyTableLockPatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: BodyTableTarget,
    before: State,
    after: State,
}

impl fmt::Debug for BodyTableLockPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableLockPatch")
            .field("table_position", &self.proof.table_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyTableLockPatch {
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

    /// Return whether applying this patch changes semantic state.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
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
pub struct BodyTableLockDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl BodyTableLockDiagnostics {
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
pub struct BodyTableLockCommit {
    package: Package,
    patch: BodyTableLockPatch,
    diagnostics: BodyTableLockDiagnostics,
}

impl BodyTableLockCommit {
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

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyTableLockPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableLockDiagnostics {
        &self.diagnostics
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct BodyTableTarget {
    table_position: usize,
    table_name: Box<str>,
    attachment_identifier: NonZeroU64,
    attachment_component_index: usize,
    attachment_object_index: usize,
    attachment_message_index: usize,
    drawable_identifier: NonZeroU64,
    model_identifier: NonZeroU64,
    model_component_index: usize,
    model_object_index: usize,
    model_message_index: usize,
    model_message_type: u32,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
    body_component_index: usize,
    body_object_index: usize,
    body_message_index: usize,
    body_message_type: u32,
    body_identifier: NonZeroU64,
    explicit_locked: Option<bool>,
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
    ) -> Result<State, BodyTableLockError> {
        let target = self.resolve_body_table(selector.into())?;
        Ok(State::from_locked(target.explicit_locked.unwrap_or(false)))
    }

    /// Start a selector-first immutable body-table lock edit.
    pub fn edit_body_table_lock<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<BodyTableLockEdit<'_>, BodyTableLockError> {
        let target = self.resolve_body_table(selector.into())?;
        let before = State::from_locked(target.explicit_locked.unwrap_or(false));
        Ok(BodyTableLockEdit {
            source: self,
            target,
            before,
            state: before,
        })
    }

    /// Apply an exact-source-checked reversible body-table lock patch.
    pub fn apply_body_table_lock(
        &self,
        patch: &BodyTableLockPatch,
    ) -> Result<BodyTableLockCommit, BodyTableLockError> {
        let source = physical_source(self)?;
        let mut budget = WireBudget::new(source.limits())?;
        if fingerprint(source.source_bytes(), &mut budget)? != patch.source_fingerprint
            || source.source_bytes() != patch.source.as_ref()
            || body_table_lock_at_target_with_budget(self, &patch.proof, &mut budget)?
                != patch.before
        {
            return Err(BodyTableLockError::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source.as_ref() != patch.target.as_ref()
                || patch.source_fingerprint != patch.target_fingerprint
            {
                return Err(BodyTableLockError::PatchConflict);
            }
            return Ok(BodyTableLockCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyTableLockDiagnostics::unchanged(),
            });
        }
        if !source.source_is_exact()
            || fingerprint(patch.target.as_ref(), &mut budget)? != patch.target_fingerprint
        {
            return Err(BodyTableLockError::PatchConflict);
        }
        let candidate = reopen_shared(Arc::clone(&patch.target), source.limits(), &mut budget)?;
        if body_table_lock_at_target_with_budget(&candidate, &patch.proof, &mut budget)?
            != patch.after
        {
            return Err(BodyTableLockError::Verification);
        }
        Ok(BodyTableLockCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableLockDiagnostics::published(),
        })
    }

    fn resolve_body_table(
        &self,
        selector: BodyTableSelector<'_>,
    ) -> Result<BodyTableTarget, BodyTableLockError> {
        let source = physical_source(self)?;
        let mut budget = WireBudget::new(source.limits())?;
        self.resolve_body_table_with_budget(selector, &mut budget)
    }

    fn resolve_body_table_with_budget(
        &self,
        selector: BodyTableSelector<'_>,
        budget: &mut WireBudget,
    ) -> Result<BodyTableTarget, BodyTableLockError> {
        let source = physical_source(self)?;
        budget.charge_source_catalog(source)?;
        let targets = native_body_table_targets_with_budget(self, budget)?;
        match selector {
            BodyTableSelector::Position(position) => targets
                .into_iter()
                .nth(position.get())
                .ok_or(BodyTableLockError::TableNotFound),
            BodyTableSelector::Name(name) => {
                let mut matching = targets
                    .into_iter()
                    .filter(|target| target.table_name.as_ref() == name);
                let Some(first) = matching.next() else {
                    return Err(BodyTableLockError::TableNotFound);
                };
                if matching.next().is_some() {
                    return Err(BodyTableLockError::AmbiguousTableName);
                }
                Ok(first)
            },
        }
    }
}

fn native_body_table_targets_with_budget(
    package: &Package,
    budget: &mut WireBudget,
) -> Result<Vec<BodyTableTarget>, BodyTableLockError> {
    let source = physical_source(package)?;
    let components = source.components();
    let object_index = index_objects(components, budget)?;
    let root_component_index = components
        .iter()
        .position(|component| component.name() == "Index/Document.iwa")
        .ok_or(BodyTableLockError::InvalidSource)?;
    let root_location = object_index
        .get(&ROOT_OBJECT_IDENTIFIER)
        .copied()
        .ok_or(BodyTableLockError::InvalidSource)?;
    if root_location.component_index != root_component_index
        || root_location.object.archive_info.identifier != Some(ROOT_OBJECT_IDENTIFIER)
    {
        return Err(BodyTableLockError::InvalidSource);
    }
    let Some((root_message_index, root_message)) = unique_message_index(
        root_location.object.messages.as_slice(),
        ROOT_MESSAGE_TYPE,
        budget,
    )?
    else {
        return Err(BodyTableLockError::InvalidSource);
    };
    validate_selected_metadata(root_location.object, root_message_index)?;
    // Root projection and body-storage decoding are part of the same read
    // transaction. Charge both strict and lazy projection passes before
    // invoking the package-owned decoders, which do not know about this
    // adapter budget.
    budget.charge_payload_work(root_message.data.len())?;
    budget.charge_payload_work(root_message.data.len())?;
    let RootReferences { body, .. } =
        root_references_with_limits(components, source.limits()).map_err(map_package_error)?;
    let root_message_info = root_location
        .object
        .archive_info
        .message_infos
        .get(root_message_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    let Some(body_identifier) = body else {
        // A body-less Pages document is a valid empty/rootless source, but a
        // retained field-4 object edge is contradictory archive metadata. Do
        // not collapse that tampered rooted edge into an ordinary selector
        // miss.
        if root_message_info.field_infos.iter().any(|field| {
            field.path.as_slice() == [ROOT_BODY_FIELD] && !field.object_references.is_empty()
        }) {
            return Err(BodyTableLockError::InvalidSource);
        }
        return Err(BodyTableLockError::TableNotFound);
    };
    if !message_declares_reference(
        root_message_info,
        body_identifier.get(),
        &[ROOT_BODY_FIELD],
        true,
        budget,
    )? {
        return Err(BodyTableLockError::InvalidSource);
    }
    let body_location = object_index
        .get(&body_identifier.get())
        .copied()
        .ok_or(BodyTableLockError::InvalidSource)?;
    let body_message = unique_text_message(body_location.object, body_identifier, budget)?;
    validate_selected_metadata(body_location.object, body_message.0)?;
    let body_payload = body_message.1.data.as_slice();
    // The body decoder performs a strict wire pass and a lazy/materializing
    // pass. Reserve both passes in the shared transaction budget before any
    // decoder-owned allocation can occur.
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
    let (body_storage, _) = decode_body_storage(
        &body_location.object.messages,
        body_identifier,
        super::MAX_SECTIONS,
        effective_text_limit(source.limits()),
        source.limits(),
    )
    .map_err(map_package_error)?;
    let Some(table_field) = unique_field(budget, &body_view, TABLE_BODY_FIELD, 2)? else {
        return Err(BodyTableLockError::TableNotFound);
    };
    let table_view = budget.parse(table_field.payload(), 1)?;
    let entry_count = table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRY_FIELD)
        .count();
    if entry_count > MAX_BODY_TABLES {
        return Err(BodyTableLockError::LimitExceeded {
            kind: BodyTableLockLimitKind::PayloadItems,
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
        .map_err(|_| BodyTableLockError::Allocation {
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
        return Err(BodyTableLockError::InvalidSource);
    }
    if let Some(first_entry) = entries.first() {
        let body_message_info = body_location
            .object
            .archive_info
            .message_infos
            .get(body_message.0)
            .ok_or(BodyTableLockError::InvalidSource)?;
        // The body archive header is shared by every table attachment. Prove
        // its aggregate reference inventory once, before the per-table graph
        // walk, so work and reference charges describe the transaction rather
        // than multiplying with the number of selected tables.
        validate_body_table_ownership(body_message_info, first_entry.identifier, &entries, budget)?;
    }

    let mut targets = Vec::new();
    budget.charge_payload_items(entries.len())?;
    targets
        .try_reserve(entries.len())
        .map_err(|_| BodyTableLockError::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_attachments = Vec::new();
    budget.charge_payload_items(entries.len())?;
    seen_attachments
        .try_reserve(entries.len())
        .map_err(|_| BodyTableLockError::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_drawables = Vec::new();
    budget.charge_payload_items(entries.len())?;
    seen_drawables
        .try_reserve(entries.len())
        .map_err(|_| BodyTableLockError::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_models = Vec::new();
    budget.charge_payload_items(entries.len())?;
    seen_models
        .try_reserve(entries.len())
        .map_err(|_| BodyTableLockError::Allocation {
            amount: entries.len(),
        })?;
    let mut entry_index = 0;
    for (character_index, character) in body_storage.text().encode_utf16().enumerate() {
        budget.charge_payload_work(1)?;
        if entries
            .get(entry_index)
            .is_some_and(|entry| entry.character_index == character_index)
        {
            if character != OBJECT_REPLACEMENT_CHARACTER {
                return Err(BodyTableLockError::InvalidSource);
            }
            entry_index += 1;
        }
    }
    if entry_index != entries.len() {
        return Err(BodyTableLockError::InvalidSource);
    }
    for entry in entries {
        let Some(attachment_location) = object_index.get(&entry.identifier.get()).copied() else {
            return Err(BodyTableLockError::InvalidSource);
        };
        budget.charge_payload_work(seen_attachments.len())?;
        budget.charge_payload_work(seen_drawables.len().saturating_add(seen_models.len()))?;
        if seen_attachments.contains(&entry.identifier)
            || seen_drawables.contains(&entry.identifier)
            || seen_models.contains(&entry.identifier)
            || entry.identifier.get() == ROOT_OBJECT_IDENTIFIER
            || entry.identifier.get() == body_identifier.get()
        {
            return Err(BodyTableLockError::InvalidSource);
        }
        seen_attachments.push(entry.identifier);
        let Some((attachment_message_index, attachment_message)) = unique_message_index(
            attachment_location.object.messages.as_slice(),
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            budget,
        )?
        else {
            return Err(BodyTableLockError::InvalidSource);
        };
        validate_selected_metadata(attachment_location.object, attachment_message_index)?;
        let drawable_identifier = parse_drawable_attachment(budget, &attachment_message.data)?;
        let Some(drawable_location) = object_index.get(&drawable_identifier.get()).copied() else {
            return Err(BodyTableLockError::InvalidSource);
        };
        budget.charge_payload_work(seen_drawables.len())?;
        budget.charge_payload_work(seen_attachments.len().saturating_add(seen_models.len()))?;
        if seen_drawables.contains(&drawable_identifier)
            || seen_attachments.contains(&drawable_identifier)
            || seen_models.contains(&drawable_identifier)
            || drawable_identifier.get() == ROOT_OBJECT_IDENTIFIER
            || drawable_identifier.get() == body_identifier.get()
        {
            return Err(BodyTableLockError::InvalidSource);
        }
        seen_drawables.push(drawable_identifier);
        let Some((message_index, message)) =
            unique_table_info(drawable_location.object.messages.as_slice(), budget)?
        else {
            return Err(BodyTableLockError::InvalidSource);
        };
        validate_selected_metadata(drawable_location.object, message_index)?;
        let snapshot = decode_table_info(&message.data, budget, body_identifier)?;
        let model_identifier = snapshot.table_model;
        let Some(model_location) = object_index.get(&model_identifier.get()).copied() else {
            return Err(BodyTableLockError::InvalidSource);
        };
        if model_location.object.archive_info.identifier != Some(model_identifier.get()) {
            return Err(BodyTableLockError::InvalidSource);
        }
        let Some((model_message_index, model_message)) =
            unique_table_model(model_location.object.messages.as_slice(), budget)?
        else {
            return Err(BodyTableLockError::InvalidSource);
        };
        validate_selected_metadata(model_location.object, model_message_index)?;
        if role_identifiers_are_aliased(
            ROOT_OBJECT_IDENTIFIER,
            body_identifier.get(),
            entry.identifier.get(),
            drawable_identifier.get(),
            model_identifier.get(),
        ) {
            return Err(BodyTableLockError::InvalidSource);
        }
        budget.charge_payload_work(seen_models.len())?;
        budget.charge_payload_work(seen_attachments.len().saturating_add(seen_drawables.len()))?;
        if seen_models.contains(&model_identifier)
            || seen_attachments.contains(&model_identifier)
            || seen_drawables.contains(&model_identifier)
            || model_identifier.get() == ROOT_OBJECT_IDENTIFIER
            || model_identifier.get() == body_identifier.get()
        {
            return Err(BodyTableLockError::InvalidSource);
        }
        seen_models.push(model_identifier);
        let table_name = decode_table_model_name(budget, &model_message.data)?;
        let table_position = targets.len();
        let target = BodyTableTarget {
            table_position,
            table_name,
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
        validate_selected_ownership(package, &target, budget)?;
        targets.push(target);
    }
    if targets.is_empty() {
        return Err(BodyTableLockError::TableNotFound);
    }
    Ok(targets)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct BodyTableEntry {
    character_index: usize,
    identifier: NonZeroU64,
}

fn parse_body_table_entry(
    budget: &mut WireBudget,
    source: &[u8],
) -> Result<BodyTableEntry, BodyTableLockError> {
    let view = budget.parse(source, 2)?;
    let index = unique_field(budget, &view, TABLE_ENTRY_CHARACTER_INDEX_FIELD, 0)?
        .ok_or(BodyTableLockError::InvalidSource)?;
    let character_index = parse_u32(index)? as usize;
    let reference = unique_field(budget, &view, TABLE_ENTRY_OBJECT_FIELD, 2)?;
    let Some(reference) = reference else {
        return Err(BodyTableLockError::InvalidSource);
    };
    Ok(BodyTableEntry {
        character_index,
        identifier: parse_reference(budget, reference.payload(), 3)?,
    })
}

fn parse_drawable_attachment(
    budget: &mut WireBudget,
    source: &[u8],
) -> Result<NonZeroU64, BodyTableLockError> {
    let view = budget.parse(source, 2)?;
    let field =
        unique_field(budget, &view, DRAWABLE_FIELD, 2)?.ok_or(BodyTableLockError::InvalidSource)?;
    parse_reference(budget, field.payload(), 3)
}

fn decode_table_model_name(
    budget: &mut WireBudget,
    source: &[u8],
) -> Result<Box<str>, BodyTableLockError> {
    let view = budget.parse(source, 1)?;
    let name = unique_field(budget, &view, TABLE_MODEL_NAME_FIELD, 2)?
        .ok_or(BodyTableLockError::InvalidSource)?;
    let text =
        std::str::from_utf8(name.payload()).map_err(|_| BodyTableLockError::InvalidSource)?;
    let mut owned = String::new();
    budget.charge_payload_items(text.len())?;
    budget.charge_payload_work(text.len())?;
    owned
        .try_reserve_exact(text.len())
        .map_err(|_| BodyTableLockError::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned.into_boxed_str())
}

fn parse_reference(
    budget: &mut WireBudget,
    source: &[u8],
    depth: usize,
) -> Result<NonZeroU64, BodyTableLockError> {
    budget.charge_payload_references(1)?;
    let view = budget.parse(source, depth)?;
    let field = unique_field(budget, &view, 1, 0)?.ok_or(BodyTableLockError::InvalidSource)?;
    let value = decode_varint(field.payload())?;
    NonZeroU64::new(value).ok_or(BodyTableLockError::InvalidSource)
}

fn parse_u32(field: WireFieldView<'_>) -> Result<u32, BodyTableLockError> {
    let value = decode_varint(field.payload())?;
    u32::try_from(value).map_err(|_| BodyTableLockError::InvalidSource)
}

fn decode_varint(source: &[u8]) -> Result<u64, BodyTableLockError> {
    let (value, length) =
        decode_varint_from_bytes(source).map_err(|_| BodyTableLockError::InvalidSource)?;
    if length != source.len() || encoded_len(value) != length {
        return Err(BodyTableLockError::InvalidSource);
    }
    Ok(value)
}

fn unique_field<'a>(
    budget: &mut WireBudget,
    view: &'a WireView<'a>,
    number: u32,
    wire_type: u8,
) -> Result<Option<WireFieldView<'a>>, BodyTableLockError> {
    budget.charge_payload_work(view.len())?;
    let mut result = None;
    for field in view.fields().filter(|field| field.number() == number) {
        validate_field(field, wire_type)?;
        if result.replace(field).is_some() {
            return Err(BodyTableLockError::InvalidSource);
        }
    }
    Ok(result)
}

fn validate_field(field: WireFieldView<'_>, wire_type: u8) -> Result<(), BodyTableLockError> {
    if field.wire_type() != wire_type {
        return Err(BodyTableLockError::InvalidSource);
    }
    field
        .validate_canonical_framing()
        .map_err(|_| BodyTableLockError::InvalidSource)
}

fn unique_text_message<'a>(
    object: &'a ArchiveObject,
    identifier: NonZeroU64,
    budget: &mut WireBudget,
) -> Result<(usize, &'a RawMessage), BodyTableLockError> {
    budget.charge_payload_messages(object.messages.len())?;
    let mut result = None;
    for (index, message) in object.messages.iter().enumerate() {
        budget.charge_payload_work(1)?;
        if matches!(message.type_, 2_001 | 2_022) {
            if result.replace((index, message)).is_some() {
                return Err(BodyTableLockError::InvalidSource);
            }
        }
    }
    result.ok_or_else(|| {
        let _ = identifier;
        BodyTableLockError::InvalidSource
    })
}

fn unique_message_index<'a>(
    messages: &'a [RawMessage],
    message_type: u32,
    budget: &mut WireBudget,
) -> Result<Option<(usize, &'a RawMessage)>, BodyTableLockError> {
    budget.charge_payload_messages(messages.len())?;
    let mut result = None;
    for (index, message) in messages.iter().enumerate() {
        budget.charge_payload_work(1)?;
        if message.type_ != message_type {
            continue;
        }
        if result.replace((index, message)).is_some() {
            return Err(BodyTableLockError::InvalidSource);
        }
    }
    Ok(result)
}

fn unique_table_info<'a>(
    messages: &'a [RawMessage],
    budget: &mut WireBudget,
) -> Result<Option<(usize, &'a RawMessage)>, BodyTableLockError> {
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
            return Err(BodyTableLockError::InvalidSource);
        }
    }
    Ok(result)
}

fn unique_table_model<'a>(
    messages: &'a [RawMessage],
    budget: &mut WireBudget,
) -> Result<Option<(usize, &'a RawMessage)>, BodyTableLockError> {
    budget.charge_payload_messages(messages.len())?;
    let mut result = None;
    for (index, message) in messages.iter().enumerate() {
        budget.charge_payload_work(1)?;
        if !TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_) {
            continue;
        }
        if result.replace((index, message)).is_some() {
            return Err(BodyTableLockError::InvalidSource);
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
) -> Result<TableInfoSnapshot, BodyTableLockError> {
    let view = budget.parse(source, 0)?;
    let super_field = unique_field(budget, &view, TABLE_INFO_SUPER_FIELD, 2)?
        .ok_or(BodyTableLockError::InvalidSource)?;
    let model_field = unique_field(budget, &view, TABLE_INFO_MODEL_FIELD, 2)?
        .ok_or(BodyTableLockError::InvalidSource)?;
    let drawable = budget.parse(super_field.payload(), 1)?;
    let parent = unique_field(budget, &drawable, DRAWABLE_PARENT_FIELD, 2)?
        .ok_or(BodyTableLockError::InvalidSource)?;
    if parse_reference(budget, parent.payload(), 3)? != body_identifier {
        return Err(BodyTableLockError::InvalidSource);
    }
    let locked = unique_field(budget, &drawable, DRAWABLE_LOCKED_FIELD, 0)?
        .map(|field| match decode_varint(field.payload())? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(BodyTableLockError::InvalidSource),
        })
        .transpose()?;
    let table_model = parse_reference(budget, model_field.payload(), 1)?;
    Ok(TableInfoSnapshot {
        table_model,
        locked,
    })
}

fn index_objects<'a>(
    components: &'a litchi_iwa_archive::ComponentCatalog,
    budget: &mut WireBudget,
) -> Result<HashMap<u64, ObjectLocation<'a>>, BodyTableLockError> {
    let object_count = components.iter().try_fold(0usize, |count, component| {
        count
            .checked_add(component.archive().objects.len())
            .ok_or(BodyTableLockError::Allocation { amount: usize::MAX })
    })?;
    // Charge the complete object inventory before reserving the index so a
    // small transaction budget cannot be bypassed by a large preallocation.
    budget.charge_payload_objects(object_count)?;
    budget.charge_payload_work(object_count)?;
    let mut index = HashMap::new();
    index
        .try_reserve(object_count)
        .map_err(|_| BodyTableLockError::Allocation {
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
                return Err(BodyTableLockError::InvalidSource);
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
) -> Result<State, BodyTableLockError> {
    validate_selected_ownership(package, target, budget)?;
    let source = physical_source(package)?;
    let component = source
        .components()
        .get_index(target.component_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.object_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(BodyTableLockError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.message_index)
        .filter(|message| message.type_ == target.message_type)
        .ok_or(BodyTableLockError::InvalidSource)?;
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
) -> Result<Package, BodyTableLockError> {
    let source_catalog = physical_source(source)?;
    let physical_limits = source_catalog.limits();
    if State::from_locked(target.explicit_locked.unwrap_or(false)) != before {
        return Err(BodyTableLockError::InvalidSource);
    }
    validate_selected_ownership(source, target, budget)?;
    let component = source_catalog
        .components()
        .get_index(target.component_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyTableLockError::UnsupportedSource);
    }
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    // Account for the compressed component scan before decompression and for
    // the decoded archive inventory before Archive::parse can reserve it.
    budget.charge_payload_work(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let stream_length = stream.as_bytes().len();
    budget.charge_archive_inventory(stream_length, component.archive())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive, budget)?;
    let object = archive
        .objects
        .get_mut(target.object_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(BodyTableLockError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.message_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if message.type_ != target.message_type {
        return Err(BodyTableLockError::InvalidSource);
    }
    validate_selected_metadata(object, target.message_index)?;
    let original_message_length = message.data.len();
    let patched_bound =
        table_info_rewrite_bound(original_message_length, target.explicit_locked.is_some()).ok_or(
            BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::OutputBytes,
                observed: u64::MAX,
                maximum: physical_limits.max_input_bytes(),
            },
        )?;
    let rewritten_bound = stream_length
        .checked_sub(original_message_length)
        .and_then(|value| value.checked_add(patched_bound))
        .and_then(|value| value.checked_add(16))
        .ok_or(BodyTableLockError::LimitExceeded {
            kind: BodyTableLockLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: physical_limits.max_input_bytes(),
        })?;
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| {
            BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::EntryBytes,
                observed: u64::MAX,
                maximum: physical_limits.max_entry_bytes(),
            }
        })?;
    let compressed_bound =
        snappy_compressed_bound(rewritten_bound).ok_or(BodyTableLockError::LimitExceeded {
            kind: BodyTableLockLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: physical_limits.max_input_bytes(),
        })?;
    let replacement_compressed_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => {
            deflate_compressed_bound(compressed_bound).ok_or(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::OutputBytes,
                observed: u64::MAX,
                maximum: physical_limits.max_input_bytes(),
            })?
        },
        _ => return Err(BodyTableLockError::UnsupportedSource),
    };
    let package_output_bound = source_catalog
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(replacement_compressed_bound))
        .ok_or(BodyTableLockError::LimitExceeded {
            kind: BodyTableLockLimitKind::OutputBytes,
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
        return Err(BodyTableLockError::Verification);
    }
    if State::from_locked(
        decode_table_info(&patched, budget, target.body_identifier)?
            .locked
            .unwrap_or(false),
    ) != after
    {
        return Err(BodyTableLockError::Verification);
    }
    object
        .replace_message_preserving_header_with_limits(
            target.message_index,
            RawMessage {
                type_: target.message_type,
                data: patched,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if rewritten.len() > rewritten_bound {
        return Err(BodyTableLockError::Verification);
    }
    drop(stream);
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(BodyTableLockError::Verification);
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
        return Err(BodyTableLockError::Verification);
    }
    drop(compressed);
    let candidate_source: Arc<[u8]> = output.into();
    let candidate = reopen_shared(candidate_source, physical_limits, budget)?;
    if body_table_lock_at_target_with_budget(&candidate, target, budget)? != after {
        return Err(BodyTableLockError::Verification);
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
fn snappy_compressed_bound(input_len: usize) -> Option<usize> {
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
fn deflate_compressed_bound(input_len: usize) -> Option<usize> {
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
) -> Result<(), BodyTableLockError> {
    let source = physical_source(package)?;
    let component = source
        .components()
        .get_index(target.component_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.object_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(BodyTableLockError::InvalidSource);
    }
    let message_info = object
        .archive_info
        .message_infos
        .get(target.message_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    validate_selected_metadata(object, target.message_index)?;
    if message_info.type_ != target.message_type
        || !message_declares_reference(
            message_info,
            target.model_identifier.get(),
            &[TABLE_INFO_MODEL_FIELD],
            true,
            budget,
        )?
        || !message_declares_reference(
            message_info,
            target.body_identifier.get(),
            &[TABLE_INFO_SUPER_FIELD, DRAWABLE_PARENT_FIELD],
            true,
            budget,
        )?
    {
        return Err(BodyTableLockError::InvalidSource);
    }
    let attachment_component = source
        .components()
        .get_index(target.attachment_component_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    let attachment = attachment_component
        .archive()
        .objects
        .get(target.attachment_object_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if attachment.archive_info.identifier != Some(target.attachment_identifier.get()) {
        return Err(BodyTableLockError::InvalidSource);
    }
    let attachment_message_info = attachment
        .archive_info
        .message_infos
        .get(target.attachment_message_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if attachment
        .messages
        .get(target.attachment_message_index)
        .is_none_or(|message| message.type_ != DRAWABLE_ATTACHMENT_MESSAGE_TYPE)
    {
        return Err(BodyTableLockError::InvalidSource);
    }
    validate_selected_metadata(attachment, target.attachment_message_index)?;
    if !message_declares_reference(
        attachment_message_info,
        target.drawable_identifier.get(),
        &[DRAWABLE_FIELD],
        true,
        budget,
    )? {
        return Err(BodyTableLockError::InvalidSource);
    }
    let body_component = source
        .components()
        .get_index(target.body_component_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    let body = body_component
        .archive()
        .objects
        .get(target.body_object_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if body.archive_info.identifier != Some(target.body_identifier.get()) {
        return Err(BodyTableLockError::InvalidSource);
    }
    validate_selected_metadata(body, target.body_message_index)?;
    let body_info = body
        .archive_info
        .message_infos
        .get(target.body_message_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if body_info.type_ != target.body_message_type {
        return Err(BodyTableLockError::InvalidSource);
    }

    // Retain and re-check the terminal model header as part of the selected
    // ownership proof.  The model is not rewritten, but its physical slot and
    // message metadata must remain the same object reached from TableInfo;
    // otherwise a changed transaction could publish against a redirected
    // archive header after resolving the drawable graph.
    let model_component = source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    let model = model_component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if model.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableLockError::InvalidSource);
    }
    let model_message = model
        .messages
        .get(target.model_message_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if model_message.type_ != target.model_message_type
        || !TABLE_MODEL_MESSAGE_TYPES.contains(&target.model_message_type)
    {
        return Err(BodyTableLockError::InvalidSource);
    }
    validate_selected_metadata(model, target.model_message_index)?;
    Ok(())
}

fn message_declares_reference(
    message: &litchi_iwa_core::MessageInfo,
    identifier: u64,
    accepted_path: &[u32],
    require_field: bool,
    budget: &mut WireBudget,
) -> Result<bool, BodyTableLockError> {
    budget.charge_payload_items(message.field_infos.len())?;
    budget.charge_payload_references(message.object_references.len())?;
    budget.charge_payload_references(message.data_references.len())?;
    budget.charge_payload_work(message.field_infos.len())?;
    let aggregate_occurrences = message
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate_occurrences != 1 {
        return Err(BodyTableLockError::InvalidSource);
    }
    if message
        .data_references
        .iter()
        .any(|candidate| *candidate == identifier)
    {
        return Err(BodyTableLockError::InvalidSource);
    }
    let mut field_declarations = 0usize;
    let mut accepted_fields = 0usize;
    for field in &message.field_infos {
        budget.charge_payload_references(field.object_references.len())?;
        budget.charge_payload_references(field.data_references.len())?;
        budget.charge_payload_work(field.path.path.len())?;
        if field
            .data_references
            .iter()
            .any(|candidate| *candidate == identifier)
        {
            return Err(BodyTableLockError::InvalidSource);
        }
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if field.path.as_slice() == accepted_path {
            accepted_fields = accepted_fields
                .checked_add(1)
                .ok_or(BodyTableLockError::InvalidSource)?;
            if (require_field || !field.object_references.is_empty())
                && (field.object_references.as_slice() != [identifier]
                    || !field.data_references.is_empty())
            {
                return Err(BodyTableLockError::InvalidSource);
            }
        }
        if occurrences != 0 {
            if occurrences != 1 || field.path.as_slice() != accepted_path {
                return Err(BodyTableLockError::InvalidSource);
            }
            field_declarations = field_declarations
                .checked_add(1)
                .ok_or(BodyTableLockError::InvalidSource)?;
            if field_declarations > 1 {
                return Err(BodyTableLockError::InvalidSource);
            }
        }
    }
    if require_field && accepted_fields != 1 {
        return Err(BodyTableLockError::InvalidSource);
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
    selected_identifier: NonZeroU64,
    entries: &[BodyTableEntry],
    budget: &mut WireBudget,
) -> Result<(), BodyTableLockError> {
    message_declares_reference_prefix(
        message,
        selected_identifier.get(),
        &[TABLE_BODY_FIELD],
        budget,
    )?;

    // Reserve for every possible field-9 declaration, rather than only the
    // selected entry inventory.  Malformed input may carry extra declaration
    // edges; charging and reserving the complete checked count keeps those
    // inserts from growing the map outside the transaction budget.
    let declaration_capacity = message
        .field_infos
        .iter()
        .try_fold(0usize, |count, field| {
            if field.path.as_slice() == [TABLE_BODY_FIELD] {
                count.checked_add(1)
            } else {
                Some(count)
            }
        })
        .ok_or(BodyTableLockError::Allocation { amount: usize::MAX })?;
    budget.charge_payload_work(message.field_infos.len())?;
    budget.charge_payload_work(declaration_capacity)?;
    let mut declared = HashMap::new();
    declared
        .try_reserve(declaration_capacity)
        .map_err(|_| BodyTableLockError::Allocation {
            amount: declaration_capacity,
        })?;
    for field in &message.field_infos {
        if field.path.as_slice() != [TABLE_BODY_FIELD] {
            continue;
        }
        budget.charge_payload_work(1)?;
        let identifier = *field
            .object_references
            .first()
            .ok_or(BodyTableLockError::InvalidSource)?;
        if declared.insert(identifier, ()).is_some() {
            return Err(BodyTableLockError::InvalidSource);
        }
    }
    if declared.len() != entries.len() {
        return Err(BodyTableLockError::InvalidSource);
    }
    for entry in entries {
        budget.charge_payload_work(1)?;
        if !declared.contains_key(&entry.identifier.get()) {
            return Err(BodyTableLockError::InvalidSource);
        }
    }
    Ok(())
}

fn message_declares_reference_prefix(
    message: &litchi_iwa_core::MessageInfo,
    identifier: u64,
    accepted_prefix: &[u32],
    budget: &mut WireBudget,
) -> Result<bool, BodyTableLockError> {
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
                    .ok_or(BodyTableLockError::Allocation { amount: usize::MAX })
            })?;
    let declaration_capacity = message
        .object_references
        .len()
        .checked_add(field_reference_capacity)
        .ok_or(BodyTableLockError::Allocation { amount: usize::MAX })?;
    let mut declarations = HashMap::new();
    budget.charge_payload_work(declaration_capacity)?;
    declarations
        .try_reserve(declaration_capacity)
        .map_err(|_| BodyTableLockError::Allocation {
            amount: declaration_capacity,
        })?;
    for aggregate_identifier in &message.object_references {
        budget.charge_payload_work(1)?;
        let counts = declarations
            .entry(*aggregate_identifier)
            .or_insert((0usize, 0usize));
        counts.1 = counts
            .1
            .checked_add(1)
            .ok_or(BodyTableLockError::InvalidSource)?;
    }
    if declarations.get(&identifier).map_or(0, |counts| counts.1) != 1 {
        return Err(BodyTableLockError::InvalidSource);
    }
    if message
        .data_references
        .iter()
        .any(|candidate| *candidate == identifier)
    {
        return Err(BodyTableLockError::InvalidSource);
    }
    let mut field_declarations = 0usize;
    for field in &message.field_infos {
        budget.charge_payload_references(field.object_references.len())?;
        budget.charge_payload_references(field.data_references.len())?;
        budget.charge_payload_work(field.path.path.len())?;
        if field.data_references.iter().any(|candidate| {
            declarations
                .get(candidate)
                .is_some_and(|counts| counts.1 != 0)
        }) {
            return Err(BodyTableLockError::InvalidSource);
        }
        if field.data_references.is_empty() {
            for field_identifier in &field.object_references {
                budget.charge_payload_work(1)?;
                let counts = declarations
                    .entry(*field_identifier)
                    .or_insert((0usize, 0usize));
                counts.0 = counts
                    .0
                    .checked_add(1)
                    .ok_or(BodyTableLockError::InvalidSource)?;
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
                return Err(BodyTableLockError::InvalidSource);
            }
            let field_identifier = field.object_references[0];
            if declarations
                .get(&field_identifier)
                .map_or(0, |counts| counts.1)
                != 1
            {
                return Err(BodyTableLockError::InvalidSource);
            }
            if field_identifier == identifier {
                field_declarations = field_declarations
                    .checked_add(1)
                    .ok_or(BodyTableLockError::InvalidSource)?;
                if field_declarations > 1 {
                    return Err(BodyTableLockError::InvalidSource);
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
                    return Err(BodyTableLockError::InvalidSource);
                }
                budget.charge_payload_work(field_index)?;
                if field.object_references[..field_index].contains(field_identifier) {
                    return Err(BodyTableLockError::InvalidSource);
                }
                if *field_identifier == identifier {
                    // The selected attachment must stay rooted at the exact
                    // table-attachment path, even when unrelated aggregate
                    // references are accepted elsewhere in the body.
                    return Err(BodyTableLockError::InvalidSource);
                }
            }
        }
    }
    if field_declarations != 1 {
        return Err(BodyTableLockError::InvalidSource);
    }
    // Every aggregate reference must have exactly one field-local declaration.
    // The counters above make this final check linear in the reference count.
    for aggregate_identifier in &message.object_references {
        budget.charge_payload_work(1)?;
        let Some((field_count, aggregate_count)) = declarations.get(aggregate_identifier) else {
            return Err(BodyTableLockError::InvalidSource);
        };
        if *aggregate_count != 1 || *field_count != 1 {
            return Err(BodyTableLockError::InvalidSource);
        }
    }
    Ok(true)
}

fn validate_selected_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), BodyTableLockError> {
    let message = object
        .messages
        .get(message_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    let message_info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableLockError::InvalidSource)?;
    if message.type_ != message_info.type_
        || object.archive_info.should_merge == Some(true)
        || message_info.base_message_index.is_some()
        || !message_info.diff_merge_version.is_empty()
        || message_info.diff_field_path.is_some()
        || !message_info.fields_to_remove.is_empty()
        || !message_info.diff_read_version.is_empty()
    {
        return Err(BodyTableLockError::InvalidSource);
    }
    Ok(())
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
    budget: &mut WireBudget,
) -> Result<(), BodyTableLockError> {
    for object in &archive.objects {
        budget.charge_payload_work(1)?;
        let offset =
            usize::try_from(object.header_offset).map_err(|_| BodyTableLockError::InvalidSource)?;
        let remaining = source
            .get(offset..)
            .ok_or(BodyTableLockError::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_| BodyTableLockError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(BodyTableLockError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes).map_err(|_| BodyTableLockError::InvalidSource)?,
            )
            .ok_or(BodyTableLockError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(BodyTableLockError::InvalidSource)?
                != object.data_offset
        {
            return Err(BodyTableLockError::InvalidSource);
        }
    }
    Ok(())
}

fn physical_source(package: &Package) -> Result<&SourceCatalog, BodyTableLockError> {
    Ok(&package.state.source)
}

fn reopen_shared(
    source: Arc<[u8]>,
    limits: litchi_iwa_archive::Limits,
    budget: &mut WireBudget,
) -> Result<Package, BodyTableLockError> {
    budget.charge_input_source(source.as_ref())?;
    budget.charge_payload_work(source.len())?;
    let catalog =
        SourceCatalog::from_shared_bytes_with_limits(source, limits).map_err(map_archive_error)?;
    budget.charge_source_catalog(&catalog)?;
    charge_reopen_work(&catalog, budget)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

/// Account for the aggregate scans and retained metadata performed by
/// [`Package::from_source_catalog`] before allowing that parser to allocate
/// its semantic snapshot.  The parser owns the concrete allocations, while
/// this adapter owns the transaction-wide ceiling.
fn charge_reopen_work(
    source: &SourceCatalog,
    budget: &mut WireBudget,
) -> Result<(), BodyTableLockError> {
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

fn fingerprint(bytes: &[u8], budget: &mut WireBudget) -> Result<u64, BodyTableLockError> {
    budget.charge_input_source(bytes)?;
    budget.charge_payload_work(bytes.len())?;
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    Ok(value)
}

struct WireBudget {
    limits: WireLimits,
    physical_limits: litchi_iwa_archive::Limits,
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
    fn new(physical_limits: litchi_iwa_archive::Limits) -> Result<Self, BodyTableLockError> {
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

    fn charge_counter(
        counter: &mut usize,
        amount: usize,
        maximum: u64,
        kind: BodyTableLockLimitKind,
    ) -> Result<(), BodyTableLockError> {
        let observed = counter
            .checked_add(amount)
            .ok_or(BodyTableLockError::LimitExceeded {
                kind,
                observed: u64::MAX,
                maximum,
            })?;
        let observed_u64 = usize_as_u64(observed);
        if observed_u64 > maximum {
            return Err(BodyTableLockError::LimitExceeded {
                kind,
                observed: observed_u64,
                maximum,
            });
        }
        *counter = observed;
        Ok(())
    }

    fn charge_wire_work(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        let work =
            self.total_work
                .checked_add(amount)
                .ok_or(BodyTableLockError::LimitExceeded {
                    kind: BodyTableLockLimitKind::WireWork,
                    observed: u64::MAX,
                    maximum: usize_as_u64(self.limits.max_rewrite_work()),
                })?;
        if work > self.limits.max_rewrite_work() {
            return Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireWork,
                observed: usize_as_u64(work),
                maximum: usize_as_u64(self.limits.max_rewrite_work()),
            });
        }
        self.total_work = work;
        Ok(())
    }

    fn charge_payload_work(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        let work =
            self.payload_work
                .checked_add(amount)
                .ok_or(BodyTableLockError::LimitExceeded {
                    kind: BodyTableLockLimitKind::WireWork,
                    observed: u64::MAX,
                    maximum: usize_as_u64(self.maximum_payload_work),
                })?;
        if work > self.maximum_payload_work {
            return Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireWork,
                observed: usize_as_u64(work),
                maximum: usize_as_u64(self.maximum_payload_work),
            });
        }
        self.charge_wire_work(amount)?;
        self.payload_work = work;
        Ok(())
    }

    fn charge_payload_references(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        self.charge_wire_work(amount)?;
        Self::charge_counter(
            &mut self.payload_references,
            amount,
            usize_as_u64(self.maximum_payload_references),
            BodyTableLockLimitKind::PayloadReferences,
        )
    }

    fn charge_input_bytes(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        Self::charge_counter(
            &mut self.input_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            BodyTableLockLimitKind::InputBytes,
        )
    }

    fn charge_output_bytes(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        Self::charge_counter(
            &mut self.output_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            BodyTableLockLimitKind::OutputBytes,
        )
    }

    fn charge_entries(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        Self::charge_counter(
            &mut self.entries,
            amount,
            usize_as_u64(self.physical_limits.max_entries()),
            BodyTableLockLimitKind::Entries,
        )
    }

    fn charge_entry_bytes(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        let observed = usize_as_u64(amount);
        let maximum = self.physical_limits.max_entry_bytes();
        if observed > maximum {
            return Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::EntryBytes,
                observed,
                maximum,
            });
        }
        Ok(())
    }

    fn charge_total_entry_bytes(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        Self::charge_counter(
            &mut self.total_entry_bytes,
            amount,
            self.physical_limits.max_total_bytes(),
            BodyTableLockLimitKind::TotalEntryBytes,
        )
    }

    fn charge_package_bytes(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        Self::charge_counter(
            &mut self.package_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            BodyTableLockLimitKind::PackageBytes,
        )
    }

    fn charge_payload_bytes(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        Self::charge_counter(
            &mut self.payload_bytes,
            amount,
            usize_as_u64(self.physical_limits.max_iwa_stream_bytes()),
            BodyTableLockLimitKind::PayloadBytes,
        )
    }

    fn charge_total_payload_bytes(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        Self::charge_counter(
            &mut self.total_payload_bytes,
            amount,
            self.physical_limits.max_total_bytes(),
            BodyTableLockLimitKind::TotalPayloadBytes,
        )
    }

    fn charge_payload_objects(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_objects,
            amount,
            usize_as_u64(limits.max_objects()),
            BodyTableLockLimitKind::PayloadObjects,
        )
    }

    fn charge_payload_messages(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_messages,
            amount,
            usize_as_u64(limits.max_messages()),
            BodyTableLockLimitKind::PayloadMessages,
        )
    }

    fn charge_payload_items(&mut self, amount: usize) -> Result<(), BodyTableLockError> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_items,
            amount,
            usize_as_u64(limits.max_metadata_items()),
            BodyTableLockLimitKind::PayloadItems,
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

    fn charge_input_source(&mut self, bytes: &[u8]) -> Result<(), BodyTableLockError> {
        if self.remember_source_key(bytes) {
            self.charge_input_bytes(bytes.len())?;
        }
        Ok(())
    }

    fn charge_source_catalog(&mut self, source: &SourceCatalog) -> Result<(), BodyTableLockError> {
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
                    BodyTableLockError::LimitExceeded {
                        kind: BodyTableLockLimitKind::EntryBytes,
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
                .ok_or(BodyTableLockError::LimitExceeded {
                    kind: BodyTableLockLimitKind::PackageBytes,
                    observed: u64::MAX,
                    maximum: self.physical_limits.max_input_bytes(),
                })?;
            self.charge_package_bytes(metadata_bytes)?;
            self.charge_payload_work(metadata_bytes)?;
        }
        Ok(())
    }

    fn charge_archive_inventory(
        &mut self,
        archive_bytes: usize,
        archive: &Archive,
    ) -> Result<(), BodyTableLockError> {
        self.charge_payload_bytes(archive_bytes)?;
        self.charge_total_payload_bytes(archive_bytes)?;
        self.charge_payload_objects(archive.objects.len())?;
        let mut message_count = 0usize;
        let mut item_count = 0usize;
        for object in &archive.objects {
            message_count = message_count.checked_add(object.messages.len()).ok_or(
                BodyTableLockError::LimitExceeded {
                    kind: BodyTableLockLimitKind::PayloadMessages,
                    observed: u64::MAX,
                    maximum: u64::MAX,
                },
            )?;
            for info in &object.archive_info.message_infos {
                item_count = item_count
                    .checked_add(1)
                    .and_then(|value| value.checked_add(info.field_infos.len()))
                    .and_then(|value| value.checked_add(info.object_references.len()))
                    .and_then(|value| value.checked_add(info.data_references.len()))
                    .ok_or(BodyTableLockError::LimitExceeded {
                        kind: BodyTableLockLimitKind::PayloadItems,
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
            .ok_or(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireWork,
                observed: u64::MAX,
                maximum: usize_as_u64(self.maximum_payload_work),
            })?;
        self.charge_payload_work(work)
    }

    fn charge_sort_work(&mut self, length: usize) -> Result<(), BodyTableLockError> {
        let levels = if length <= 1 {
            0
        } else {
            (usize::BITS - (length - 1).leading_zeros()) as usize
        };
        let amount = length
            .checked_mul(levels)
            .ok_or(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireWork,
                observed: u64::MAX,
                maximum: usize_as_u64(self.maximum_payload_work),
            })?;
        self.charge_payload_work(amount)
    }

    fn parse<'a>(
        &mut self,
        source: &'a [u8],
        depth: usize,
    ) -> Result<WireView<'a>, BodyTableLockError> {
        if depth > self.limits.max_nesting() {
            return Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireNesting,
                observed: usize_as_u64(depth),
                maximum: usize_as_u64(self.limits.max_nesting()),
            });
        }
        let bytes = self.total_bytes.checked_add(source.len()).ok_or(
            BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireBytes,
                observed: u64::MAX,
                maximum: usize_as_u64(self.limits.max_input_bytes()),
            },
        )?;
        if bytes > self.limits.max_input_bytes() {
            return Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireBytes,
                observed: usize_as_u64(bytes),
                maximum: usize_as_u64(self.limits.max_input_bytes()),
            });
        }
        let view = WireView::parse_with_limits(source, self.limits).map_err(map_wire_error)?;
        let fields =
            self.total_fields
                .checked_add(view.len())
                .ok_or(BodyTableLockError::LimitExceeded {
                    kind: BodyTableLockLimitKind::WireFields,
                    observed: u64::MAX,
                    maximum: usize_as_u64(self.limits.max_fields()),
                })?;
        if fields > self.limits.max_fields() {
            return Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireFields,
                observed: usize_as_u64(fields),
                maximum: usize_as_u64(self.limits.max_fields()),
            });
        }
        let work =
            source
                .len()
                .checked_add(view.len())
                .ok_or(BodyTableLockError::LimitExceeded {
                    kind: BodyTableLockLimitKind::WireWork,
                    observed: u64::MAX,
                    maximum: usize_as_u64(self.limits.max_rewrite_work()),
                })?;
        self.charge_wire_work(work)?;
        self.total_bytes = bytes;
        self.total_fields = fields;
        Ok(view)
    }
}

fn map_package_error(error: PackageError) -> BodyTableLockError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => BodyTableLockError::Allocation { amount },
        PackageError::PayloadLimit { observed, limit } => BodyTableLockError::LimitExceeded {
            kind: BodyTableLockLimitKind::PayloadItems,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => BodyTableLockError::LimitExceeded {
            kind: BodyTableLockLimitKind::PayloadObjects,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        PackageError::Semantic(error) => match error {
            crate::Error::TooManySections { actual, limit }
            | crate::Error::TooManyBodyStorages { actual, limit } => {
                BodyTableLockError::LimitExceeded {
                    kind: BodyTableLockLimitKind::PayloadItems,
                    observed: usize_as_u64(actual),
                    maximum: usize_as_u64(limit),
                }
            },
            crate::Error::TextTooLarge { observed, limit } => BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::PayloadBytes,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(limit),
            },
            crate::Error::InvalidSectionIndex { .. } => BodyTableLockError::InvalidSource,
        },
        PackageError::SectionNamesTooLarge { observed, limit } => {
            BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::PayloadBytes,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(limit),
            }
        },
        PackageError::InvalidFormat(_)
        | PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::NotPages => BodyTableLockError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyTableLockError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableLockError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => BodyTableLockLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => BodyTableLockLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => BodyTableLockLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    BodyTableLockLimitKind::PackageBytes
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => BodyTableLockLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    BodyTableLockLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyTableLockLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    BodyTableLockLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyTableLockError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Reassembly(_) => BodyTableLockError::UnsupportedSource,
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::InvalidBundle(_) => BodyTableLockError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyTableLockError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableLockError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => BodyTableLockLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    BodyTableLockLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => BodyTableLockLimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => BodyTableLockLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    BodyTableLockLimitKind::PayloadBytes
                },
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableLockError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => BodyTableLockError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> BodyTableLockError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => BodyTableLockError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => BodyTableLockLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => BodyTableLockLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => BodyTableLockLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => BodyTableLockLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => BodyTableLockLimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    BodyTableLockLimitKind::PayloadItems
                },
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            BodyTableLockError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => BodyTableLockError::InvalidSource,
    }
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_core::{FieldInfo, FieldPath};

    fn table_info_payload() -> Vec<u8> {
        let mut reference = Vec::new();
        litchi_iwa_common::wire::append_varint_field(&mut reference, 1, 42)
            .expect("reference field");

        let mut drawable = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(&mut drawable, 2, &reference)
            .expect("drawable parent field");

        let mut table_info = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(&mut table_info, 1, &drawable)
            .expect("table-info super field");
        litchi_iwa_common::wire::append_length_delimited_field(&mut table_info, 2, &reference)
            .expect("table-info model field");
        table_info
    }

    fn budget_with_wire_limits(
        input_bytes: usize,
        fields: usize,
        rewrite_work: usize,
    ) -> WireBudget {
        let mut budget =
            WireBudget::new(litchi_iwa_archive::Limits::default()).expect("default wire budget");
        budget.limits = WireLimits::default()
            .with_input_bytes(input_bytes)
            .and_then(|limits| limits.with_fields(fields))
            .and_then(|limits| limits.with_rewrite_work(rewrite_work))
            .expect("test wire limits");
        budget
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
            (
                BodyTableLockLimitKind::WireBytes,
                one_bytes * 2 - 1,
                1024,
                1024 * 1024,
            ),
            (
                BodyTableLockLimitKind::WireFields,
                1024,
                one_fields * 2 - 1,
                1024 * 1024,
            ),
            (
                BodyTableLockLimitKind::WireWork,
                1024,
                1024,
                one_work * 2 - 1,
            ),
        ];

        for (kind, input_bytes, fields, rewrite_work) in dimensions {
            let mut budget = budget_with_wire_limits(input_bytes, fields, rewrite_work);
            decode_table_info(&payload, &mut budget, body_identifier)
                .expect("first bounded table-info decode");
            assert!(matches!(
                decode_table_info(&payload, &mut budget, body_identifier),
                Err(BodyTableLockError::LimitExceeded { kind: observed, .. }) if observed == kind
            ));
        }
    }

    fn tiny_physical_limits() -> litchi_iwa_archive::Limits {
        let archive = litchi_iwa_core::Limits::default()
            .with_archive_bytes(8)
            .and_then(|limits| limits.with_objects(2))
            .and_then(|limits| limits.with_messages(4))
            .and_then(|limits| limits.with_metadata_items(4))
            .expect("test archive limits");
        litchi_iwa_archive::Limits::new(16, 2, 8, 16, 8)
            .expect("test physical limits")
            .with_archive_limits(archive)
            .expect("test physical profile")
    }

    fn assert_physical_boundary(
        kind: BodyTableLockLimitKind,
        maximum: usize,
        charge: fn(&mut WireBudget, usize) -> Result<(), BodyTableLockError>,
    ) {
        let mut budget = WireBudget::new(tiny_physical_limits()).expect("test wire budget");
        charge(&mut budget, maximum).expect("maximum is accepted");
        assert!(matches!(
            charge(&mut budget, 1),
            Err(BodyTableLockError::LimitExceeded { kind: observed, .. }) if observed == kind
        ));
    }

    #[test]
    fn physical_budget_boundaries_reject_max_plus_one() {
        assert_physical_boundary(
            BodyTableLockLimitKind::InputBytes,
            16,
            WireBudget::charge_input_bytes,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::OutputBytes,
            16,
            WireBudget::charge_output_bytes,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::Entries,
            2,
            WireBudget::charge_entries,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::TotalEntryBytes,
            16,
            WireBudget::charge_total_entry_bytes,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::PackageBytes,
            16,
            WireBudget::charge_package_bytes,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::PayloadBytes,
            8,
            WireBudget::charge_payload_bytes,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::TotalPayloadBytes,
            16,
            WireBudget::charge_total_payload_bytes,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::PayloadObjects,
            2,
            WireBudget::charge_payload_objects,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::PayloadMessages,
            4,
            WireBudget::charge_payload_messages,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::PayloadItems,
            4,
            WireBudget::charge_payload_items,
        );
        assert_physical_boundary(
            BodyTableLockLimitKind::PayloadReferences,
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
            Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::EntryBytes,
                observed: 9,
                maximum: 8,
            })
        ));
    }

    #[test]
    fn source_catalog_charges_multi_member_entry_bytes_aggregate_separately() {
        let limits =
            litchi_iwa_archive::Limits::new(512, 4, 8, 16, 64).expect("test physical limits");
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
        let catalog =
            SourceCatalog::from_bytes_with_limits(&source, limits).expect("source catalog");
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
        let compressed = SnappyStream::compress(&input).expect("Snappy compression");

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
        let mut message = litchi_iwa_core::MessageInfo::new(2_001, 0);
        message.object_references = vec![100, 900, 901];
        let mut selected = FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD]));
        selected.object_references.push(100);
        let mut sections = FieldInfo::new(FieldPath::new(vec![17]));
        sections.object_references.extend([900, 901]);
        message.field_infos = vec![selected, sections];

        let mut budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        assert!(
            message_declares_reference_prefix(&message, 100, &[TABLE_BODY_FIELD], &mut budget)
                .expect("unrelated aggregate references use their own path")
        );

        message.field_infos[0].path = FieldPath::new(vec![17]);
        let mut strict_budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        assert!(matches!(
            message_declares_reference_prefix(
                &message,
                100,
                &[TABLE_BODY_FIELD],
                &mut strict_budget,
            ),
            Err(BodyTableLockError::InvalidSource)
        ));
    }

    #[test]
    fn data_reference_scans_are_bounded_before_semantic_checks() {
        let mut message_data = litchi_iwa_core::MessageInfo::new(2_002, 0);
        message_data.object_references = vec![100];
        message_data.data_references = vec![100];
        message_data.field_infos = vec![FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD]))];

        let mut message_budget = budget_with_wire_limits(1024, 1024, 1024 * 1024);
        message_budget.maximum_payload_references = 1;
        assert!(matches!(
            message_declares_reference(
                &message_data,
                100,
                &[TABLE_BODY_FIELD],
                false,
                &mut message_budget,
            ),
            Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::PayloadReferences,
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
            Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::PayloadReferences,
                ..
            })
        ));

        let mut field_data = litchi_iwa_core::MessageInfo::new(2_003, 0);
        field_data.object_references = vec![100];
        let mut field = FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD]));
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
                &mut field_budget,
            ),
            Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::PayloadReferences,
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
            Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::PayloadReferences,
                ..
            })
        ));
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
}
