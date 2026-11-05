//! Immutable, exact-source transactions for attached Numbers table locks.

use std::fmt;

use litchi_iwa_archive::package::{EntryEdit, SharedBytes};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    varint::encoded_len,
    wire::{patch_varint_field, transform_length_delimited_field},
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::table_info_codec::{self, WireResourceLimit};
use thiserror::Error as ThisError;

use super::{
    Error as ReadError, LEGACY_TABLE_INFO_MESSAGE_TYPE, Package, Resolved, TABLE_INFO_MESSAGE_TYPE,
    table_info_decode_options,
};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::lock::State;

const TABLE_DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_LOCKED_FIELD: u32 = 5;

/// A finite resource enforced while reading or publishing a table-lock edit.
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
    /// Bytes inspected by the `TableInfo` wire projection.
    WireBytes,
    /// Fields inspected by the `TableInfo` wire projection.
    WireFields,
    /// Nesting used by the `TableInfo` wire projection.
    WireNesting,
    /// Aggregate work charged by the `TableInfo` wire projection.
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

/// Failure from a semantic table-lock read or immutable transaction.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// Rooted semantic selection did not resolve to one supported native payload.
    #[error("the selected Numbers table has no unambiguous editable lock payload")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Numbers table-lock {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested resource amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Numbers table-lock transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
    },
    /// Complete candidate reopening did not reproduce the requested lock state.
    #[error("the edited Numbers table lock failed semantic verification")]
    Verification,
    /// The supplied patch was not created from this exact package artifact.
    #[error("the Numbers table-lock patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The selected sheet does not exist in the rooted workbook.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// The selected sheet has no table matching the requested selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// The package source cannot publish a preservation-safe changed edit.
    #[error("the Numbers package source does not support exact table-lock editing")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound { .. } | Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// A mutable semantic setting staged against one immutable package snapshot.
///
/// An equal-state commit preserves the source's exact lock-field presence.
#[derive(Debug)]
pub struct Edit<'a> {
    source: &'a Package,
    sheet_position: usize,
    table_position: usize,
    before: State,
    state: State,
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
    ///
    /// Equal-state commits preserve absent versus explicit-false encoding.
    /// The inverse of a changed commit restores the exact original encoding.
    ///
    /// # Costs
    ///
    /// A changed edit rewrites one bounded package component and fully reopens
    /// the candidate. A no-op shares the original immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the source is unsupported or inconsistent, a
    /// resource ceiling is exceeded, or candidate verification fails.
    pub fn commit(self) -> Result<Commit, Error> {
        let source = physical_source(self.source)?;
        let source_bytes = source.source_owner();
        let source_fingerprint = fingerprint(&source_bytes);
        if self.before == self.state {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    source: source_bytes.clone(),
                    target: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    sheet_position: self.sheet_position,
                    table_position: self.table_position,
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
            self.sheet_position,
            self.table_position,
            self.before,
            self.state,
        )?;
        let target = physical_source(&package)?.source_owner();
        let target_fingerprint = fingerprint(&target);
        Ok(Commit {
            package,
            patch: Patch {
                source: source_bytes,
                target,
                source_fingerprint,
                target_fingerprint,
                sheet_position: self.sheet_position,
                table_position: self.table_position,
                before: self.before,
                after: self.state,
            },
            diagnostics: Diagnostics::published(),
        })
    }
}

/// A reversible patch bound to the exact source and target package artifacts.
///
/// Its inverse restores the source artifact exactly, including whether an
/// unlocked table encoded the lock field as absent or explicit false.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: SharedBytes,
    target: SharedBytes,
    source_fingerprint: u64,
    target_fingerprint: u64,
    sheet_position: usize,
    table_position: usize,
    before: State,
    after: State,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("sheet_position", &self.sheet_position)
            .field("table_position", &self.table_position)
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

    /// Return whether the patch is both semantically and byte-wise unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source.as_ref() == self.target.as_ref()
            && self.source_fingerprint == self.target_fingerprint
    }

    /// Return an exact-source inverse that restores the original artifact.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            sheet_position: self.sheet_position,
            table_position: self.table_position,
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

    /// Return whether the committed package differs semantically from its source.
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

/// The fully reopened immutable result of one table-lock transaction.
#[must_use = "a Numbers table-lock commit contains the validated package snapshot"]
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

    /// Consume the commit and return the fully reopened package snapshot.
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

#[derive(Debug, Clone, Copy)]
struct NativeTarget {
    drawable_identifier: u64,
    model_identifier: u64,
    owner_identifier: u64,
    owner_component_index: usize,
    owner_object_index: usize,
    owner_message_index: usize,
    owner_message_type: u32,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
    explicit_locked: Option<bool>,
}

impl Package {
    /// Read one attached table's effective interactive lock state.
    ///
    /// Both an absent native lock field and an explicit `false` read as
    /// [`State::Unlocked`].
    ///
    /// # Costs
    ///
    /// Selection uses the package's retained semantic and native indexes.
    ///
    /// # Errors
    ///
    /// Returns an error when either selector has no unique rooted match or the
    /// selected native payload cannot be strictly decoded.
    pub fn table_lock<'sheet, 'table>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<State, Error> {
        let (sheet_position, table_position) =
            self.resolve_table_selectors(sheet_selector, table)?;
        self.table_lock_at(sheet_position, table_position)
    }

    /// Start a selector-first immutable table-lock edit.
    ///
    /// An equal-state commit preserves existing lock-field presence; a changed
    /// commit's inverse restores the exact original presence and bytes.
    ///
    /// # Costs
    ///
    /// The edit borrows this package and does not copy package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when either selector has no unique rooted match or the
    /// selected native payload cannot be strictly decoded.
    pub fn edit_table_lock<'sheet, 'table>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet_position, table_position) =
            self.resolve_table_selectors(sheet_selector, table)?;
        let before = self.table_lock_at(sheet_position, table_position)?;
        Ok(Edit {
            source: self,
            sheet_position,
            table_position,
            before,
            state: before,
        })
    }

    /// Apply an exact-source-checked reversible table-lock patch.
    ///
    /// # Costs
    ///
    /// A changed patch fully reopens its bounded target artifact. A no-op
    /// shares the current immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch does not match this exact artifact or
    /// when the target cannot be reopened and semantically verified.
    pub fn apply_table_lock(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if fingerprint(source.source_bytes()) != patch.source_fingerprint
            || source.source_bytes() != patch.source.as_ref()
            || self.table_lock_at(patch.sheet_position, patch.table_position)? != patch.before
        {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source.as_ref() != patch.target.as_ref()
                || patch.source_fingerprint != patch.target_fingerprint
            {
                return Err(Error::PatchConflict);
            }
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() || fingerprint(&patch.target) != patch.target_fingerprint {
            return Err(Error::PatchConflict);
        }

        let candidate =
            Package::from_source_owner_with_options(patch.target.clone(), self.state.options)
                .map_err(map_candidate_read_error)?;
        if candidate.table_lock_at(patch.sheet_position, patch.table_position)? != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }

    fn resolve_table_selectors<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<(usize, usize), Error> {
        let selected_sheet = self
            .state
            .document
            .sheet(sheet)
            .map_err(|error| map_read_error(ReadError::Semantic(error)))?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?;
        let selected = selected_sheet
            .select(table.into())
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::TableNotFound))?;
        let table_position = selected_sheet
            .tables()
            .position(|candidate| std::ptr::eq(candidate, selected))
            .ok_or(Error::InvalidSource)?;
        Ok((selected_sheet.index(), table_position))
    }

    fn table_lock_at(&self, sheet_position: usize, table_position: usize) -> Result<State, Error> {
        let target = self.native_table_lock_target(sheet_position, table_position)?;
        Ok(State::from_locked(target.explicit_locked.unwrap_or(false)))
    }

    fn native_table_lock_target(
        &self,
        sheet_position: usize,
        table_position: usize,
    ) -> Result<NativeTarget, Error> {
        let document = Self::root_document(&self.state.components).map_err(map_read_error)?;
        let sheet_reference = document
            .sheet_references()
            .get(sheet_position)
            .ok_or(Error::Refused(Refusal::SheetNotFound))?;
        let sheet_object = self
            .state
            .index
            .resolve_ref_id(&self.state.components, sheet_reference.identifier())
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource)?;
        let owner_message_index = unique_sheet_message_index(sheet_object.messages)?;
        let owner_message_type = sheet_object
            .messages
            .get(owner_message_index)
            .ok_or(Error::InvalidSource)?
            .type_;
        let owner_component_index = sheet_object.component_index;
        let owner_object_index = sheet_object.object_index;
        let path = super::SemanticPath::Sheet {
            index: sheet_position,
        };
        let decoded_sheet =
            super::sheet_projection(sheet_object.messages, path).map_err(map_read_error)?;

        let mut table_index = 0usize;
        for drawable in decoded_sheet.drawable_infos {
            let resolved = self
                .state
                .index
                .resolve_ref_id(&self.state.components, drawable.identifier)
                .map_err(map_read_error)?
                .ok_or(Error::InvalidSource)?;
            let Some((message_index, message)) = unique_table_info(resolved)? else {
                continue;
            };
            let snapshot = decode_table_info(&message.data)?;
            if table_index == table_position {
                return Ok(NativeTarget {
                    drawable_identifier: drawable.identifier,
                    model_identifier: snapshot.table_model().identifier().get(),
                    owner_identifier: sheet_reference.identifier(),
                    owner_component_index,
                    owner_object_index,
                    owner_message_index,
                    owner_message_type,
                    component_index: resolved.component_index,
                    object_index: resolved.object_index,
                    message_index,
                    message_type: message.type_,
                    explicit_locked: snapshot.locked(),
                });
            }
            table_index = table_index.checked_add(1).ok_or(Error::InvalidSource)?;
        }
        Err(Error::Refused(Refusal::TableNotFound))
    }
}

fn unique_sheet_message_index(messages: &[RawMessage]) -> Result<usize, Error> {
    let sheet = unique_message_index(messages, super::SHEET_MESSAGE_TYPE)?;
    let form = unique_message_index(messages, super::FORM_BASED_SHEET_MESSAGE_TYPE)?;
    match (sheet, form) {
        (Some(_), Some(_)) | (None, None) => Err(Error::InvalidSource),
        (Some((index, _)), None) | (None, Some((index, _))) => Ok(index),
    }
}

fn unique_table_info(resolved: Resolved<'_>) -> Result<Option<(usize, &RawMessage)>, Error> {
    let canonical = unique_message_index(resolved.messages, TABLE_INFO_MESSAGE_TYPE)?;
    let legacy = unique_message_index(resolved.messages, LEGACY_TABLE_INFO_MESSAGE_TYPE)?;
    match (canonical, legacy) {
        (Some(_), Some(_)) => Err(Error::InvalidSource),
        (Some(message), None) | (None, Some(message)) => Ok(Some(message)),
        (None, None) => Ok(None),
    }
}

fn unique_message_index(
    messages: &[RawMessage],
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, Error> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == message_type);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(first)
}

fn decode_table_info(source: &[u8]) -> Result<table_info_codec::TableInfoSnapshot, Error> {
    table_info_codec::decode_table_info(source, table_info_decode_options(source)).map_err(
        |error| {
            if let Some((observed, maximum)) = error.field_limit_values() {
                return Error::Limit {
                    kind: LimitKind::WireFields,
                    observed: usize_as_u64(observed),
                    maximum: usize_as_u64(maximum),
                };
            }
            if let Some((observed, maximum)) = error.work_limit_values() {
                return Error::Limit {
                    kind: LimitKind::WireWork,
                    observed: usize_as_u64(observed),
                    maximum: usize_as_u64(maximum),
                };
            }
            match error.wire_resource_limit() {
                Some(WireResourceLimit::Bytes {
                    observed,
                    maximum: Some(maximum),
                }) => Error::Limit {
                    kind: LimitKind::WireBytes,
                    observed: usize_as_u64(observed.unwrap_or_else(|| maximum.saturating_add(1))),
                    maximum: usize_as_u64(maximum),
                },
                Some(WireResourceLimit::Nesting {
                    observed,
                    maximum: Some(maximum),
                }) => Error::Limit {
                    kind: LimitKind::WireNesting,
                    observed: u64::from(observed.unwrap_or_else(|| maximum.saturating_add(1))),
                    maximum: u64::from(maximum),
                },
                _ => Error::InvalidSource,
            }
        },
    )
}

fn rewrite_lock_state(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    before: State,
    after: State,
) -> Result<Package, Error> {
    let target = source.native_table_lock_target(sheet_position, table_position)?;
    if State::from_locked(target.explicit_locked.unwrap_or(false)) != before {
        return Err(Error::InvalidSource);
    }
    let source_catalog = physical_source(source)?;
    validate_selected_ownership(source, target)?;
    let component = source
        .state
        .components
        .catalog()
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
    let physical_limits = source_catalog.limits();
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive = Archive::parse(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    drop(stream);
    let object = archive
        .objects
        .get_mut(target.object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier) {
        return Err(Error::InvalidSource);
    }
    let message = object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != target.message_type {
        return Err(Error::InvalidSource);
    }
    let snapshot = decode_table_info(&message.data)?;
    validate_selected_metadata(object, target.message_index)?;
    if snapshot.locked() != target.explicit_locked {
        return Err(Error::InvalidSource);
    }
    let explicit_before = target.explicit_locked;
    let patched = transform_length_delimited_field::<_, litchi_iwa_common::Error>(
        &message.data,
        TABLE_DRAWABLE_SUPER_FIELD,
        |drawable| {
            patch_varint_field(
                drawable,
                DRAWABLE_LOCKED_FIELD,
                explicit_before.is_some(),
                Some(u64::from(after.is_locked())),
            )
        },
    )
    .map_err(map_wire_error)?;
    if State::from_locked(decode_table_info(&patched)?.locked().unwrap_or(false)) != after {
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
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten, snappy_limits).map_err(map_core_error)?;
    drop(rewritten);
    let output = source_catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            physical_limits,
        )
        .map_err(map_archive_error)?;
    drop(compressed);
    let candidate = source
        .reopen_candidate(output.into())
        .map_err(map_candidate_read_error)?;
    if candidate.table_lock_at(sheet_position, table_position)? != after {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

fn physical_source(package: &Package) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    package
        .state
        .components
        .physical()
        .ok_or(Error::Refused(Refusal::UnsupportedSource))
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

fn validate_selected_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
) -> Result<(), Error> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_selected_ownership(package: &Package, target: NativeTarget) -> Result<(), Error> {
    let catalog = package.state.components.catalog();
    let sheet_object = catalog
        .get_index(target.owner_component_index)
        .and_then(|component| component.archive().objects.get(target.owner_object_index))
        .ok_or(Error::InvalidSource)?;
    if sheet_object.archive_info.identifier != Some(target.owner_identifier) {
        return Err(Error::InvalidSource);
    }
    let sheet = super::sheet_projection(&sheet_object.messages, super::SemanticPath::Package)
        .map_err(map_read_error)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|drawable| drawable.identifier == target.drawable_identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource);
    }
    let sheet_message = sheet_object
        .archive_info
        .message_infos
        .get(target.owner_message_index)
        .ok_or(Error::InvalidSource)?;
    if sheet_message.type_ != target.owner_message_type {
        return Err(Error::InvalidSource);
    }
    let owner_path: &[u32] = match target.owner_message_type {
        super::SHEET_MESSAGE_TYPE => &[2],
        super::FORM_BASED_SHEET_MESSAGE_TYPE => &[1, 2],
        _ => return Err(Error::InvalidSource),
    };
    if !message_declares_reference(sheet_message, target.drawable_identifier, owner_path)? {
        return Err(Error::InvalidSource);
    }

    let table_object = catalog
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource)?;
    if table_object.archive_info.identifier != Some(target.drawable_identifier) {
        return Err(Error::InvalidSource);
    }
    let table_message = table_object
        .archive_info
        .message_infos
        .get(target.message_index)
        .ok_or(Error::InvalidSource)?;
    if table_message.type_ != target.message_type {
        return Err(Error::InvalidSource);
    }
    if !message_declares_reference(table_message, target.model_identifier, &[2])? {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn message_declares_reference(
    message: &litchi_iwa_core::MessageInfo,
    identifier: u64,
    accepted_field_path: &[u32],
) -> Result<bool, Error> {
    let mut declared = message.object_references.contains(&identifier);
    for field in &message.field_infos {
        if field.object_references.contains(&identifier) {
            if field.path.as_slice() != accepted_field_path {
                return Err(Error::InvalidSource);
            }
            declared = true;
        }
    }
    Ok(declared)
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), Error> {
    for object in &archive.objects {
        let offset =
            usize::try_from(object.header_offset).map_err(|_error| Error::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(Error::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_error| Error::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(Error::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(u64::try_from(prefix_bytes).map_err(|_error| Error::InvalidSource)?)
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

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn map_candidate_read_error(error: ReadError) -> Error {
    match error {
        ReadError::InputTooLarge { observed, maximum }
        | ReadError::Archive(litchi_iwa_archive::Error::Limit {
            kind: litchi_iwa_archive::LimitKind::InputBytes,
            observed,
            maximum,
        }) => Error::Limit {
            kind: LimitKind::OutputBytes,
            observed,
            maximum,
        },
        other @ (ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::Archive(_)
        | ReadError::MalformedPayload { .. }
        | ReadError::UnexpectedFormat { .. }
        | ReadError::InvalidLimit { .. }
        | ReadError::Common(_)
        | ReadError::InvalidFormat(_)
        | ReadError::ParseError(_)
        | ReadError::Semantic(_)
        | ReadError::SemanticLimit { .. }) => map_read_error(other),
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Common(wire_error) => map_wire_error(wire_error),
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::PayloadObjects,
                super::SemanticLimitKind::References => LimitKind::PayloadReferences,
                super::SemanticLimitKind::OutputTextBytes
                | super::SemanticLimitKind::FormulaCategoryBytes
                | super::SemanticLimitKind::TextBytes => LimitKind::PayloadBytes,
                super::SemanticLimitKind::FormulaRenderDepth
                | super::SemanticLimitKind::FormulaDepth => LimitKind::WireNesting,
                super::SemanticLimitKind::FormulaRenderWork
                | super::SemanticLimitKind::FormulaWork => LimitKind::WireWork,
                super::SemanticLimitKind::Sheets
                | super::SemanticLimitKind::Tables
                | super::SemanticLimitKind::MaterializedCells => LimitKind::PayloadItems,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        ReadError::InputTooLarge { observed, maximum } => Error::Limit {
            kind: LimitKind::InputBytes,
            observed,
            maximum,
        },
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::MalformedPayload { .. }
        | ReadError::UnexpectedFormat { .. }
        | ReadError::InvalidLimit { .. }
        | ReadError::InvalidFormat(_)
        | ReadError::ParseError(_)
        | ReadError::Semantic(_) => Error::InvalidSource,
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
        litchi_iwa_archive::Error::Reassembly(_) => Error::Refused(Refusal::UnsupportedSource),
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        _ => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "this mapper is passed directly to Result::map_err"
)]
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

#[allow(
    clippy::needless_pass_by_value,
    reason = "this mapper is passed directly to Result::map_err"
)]
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

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use litchi_iwa_archive::package::SharedBytes;

    use super::{Patch, State, fingerprint};

    fn patch(source: SharedBytes, target: SharedBytes) -> Patch {
        Patch {
            source_fingerprint: fingerprint(source.as_ref()),
            target_fingerprint: fingerprint(target.as_ref()),
            source,
            target,
            sheet_position: 0,
            table_position: 0,
            before: State::Unlocked,
            after: State::Unlocked,
        }
    }

    #[test]
    fn is_noop_requires_exact_source_and_target_artifacts() {
        let source = SharedBytes::from_shared_slice(Arc::from([1_u8, 2]));
        let different_target = SharedBytes::from_shared_slice(Arc::from([1_u8, 3]));
        assert!(!patch(source.clone(), different_target).is_noop());

        let equivalent_target = SharedBytes::from_shared_slice(Arc::from([1_u8, 2]));
        assert!(patch(source, equivalent_target).is_noop());
    }
}
