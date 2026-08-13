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
use thiserror::Error;

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
pub enum TableLockLimitKind {
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

impl fmt::Display for TableLockLimitKind {
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
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TableLockError {
    /// The selected sheet does not exist in the rooted workbook.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// The selected sheet has no table matching the requested selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// The package source cannot publish a preservation-safe changed edit.
    #[error("the Numbers package source does not support exact table-lock editing")]
    UnsupportedSource,
    /// Rooted semantic selection did not resolve to one supported native payload.
    #[error("the selected Numbers table has no unambiguous editable lock payload")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Numbers table-lock {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: TableLockLimitKind,
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

/// A mutable semantic setting staged against one immutable package snapshot.
///
/// An equal-state commit preserves the source's exact lock-field presence.
#[derive(Debug)]
pub struct TableLockEdit<'a> {
    source: &'a Package,
    sheet_position: usize,
    table_position: usize,
    before: State,
    state: State,
}

impl TableLockEdit<'_> {
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
    pub fn commit(self) -> Result<TableLockCommit, TableLockError> {
        let source = physical_source(self.source)?;
        let source_bytes = source.__source_owner();
        let source_fingerprint = fingerprint(&source_bytes);
        if self.before == self.state {
            return Ok(TableLockCommit {
                package: self.source.snapshot(),
                patch: TableLockPatch {
                    source: source_bytes.clone(),
                    target: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    sheet_position: self.sheet_position,
                    table_position: self.table_position,
                    before: self.before,
                    after: self.state,
                },
                diagnostics: TableLockDiagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() {
            return Err(TableLockError::UnsupportedSource);
        }

        let package = rewrite_lock_state(
            self.source,
            self.sheet_position,
            self.table_position,
            self.before,
            self.state,
        )?;
        let target = physical_source(&package)?.__source_owner();
        let target_fingerprint = fingerprint(&target);
        Ok(TableLockCommit {
            package,
            patch: TableLockPatch {
                source: source_bytes,
                target,
                source_fingerprint,
                target_fingerprint,
                sheet_position: self.sheet_position,
                table_position: self.table_position,
                before: self.before,
                after: self.state,
            },
            diagnostics: TableLockDiagnostics::published(),
        })
    }
}

/// A reversible patch bound to the exact source and target package artifacts.
///
/// Its inverse restores the source artifact exactly, including whether an
/// unlocked table encoded the lock field as absent or explicit false.
#[derive(Clone, PartialEq, Eq)]
pub struct TableLockPatch {
    source: SharedBytes,
    target: SharedBytes,
    source_fingerprint: u64,
    target_fingerprint: u64,
    sheet_position: usize,
    table_position: usize,
    before: State,
    after: State,
}

impl fmt::Debug for TableLockPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TableLockPatch")
            .field("sheet_position", &self.sheet_position)
            .field("table_position", &self.table_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl TableLockPatch {
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
pub struct TableLockDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl TableLockDiagnostics {
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
pub struct TableLockCommit {
    package: Package,
    patch: TableLockPatch,
    diagnostics: TableLockDiagnostics,
}

impl TableLockCommit {
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

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &TableLockPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &TableLockDiagnostics {
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
    ) -> Result<State, TableLockError> {
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
    ) -> Result<TableLockEdit<'_>, TableLockError> {
        let (sheet_position, table_position) =
            self.resolve_table_selectors(sheet_selector, table)?;
        let before = self.table_lock_at(sheet_position, table_position)?;
        Ok(TableLockEdit {
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
    pub fn apply_table_lock(
        &self,
        patch: &TableLockPatch,
    ) -> Result<TableLockCommit, TableLockError> {
        let source = physical_source(self)?;
        if fingerprint(source.source_bytes()) != patch.source_fingerprint
            || source.source_bytes() != patch.source.as_ref()
            || self.table_lock_at(patch.sheet_position, patch.table_position)? != patch.before
        {
            return Err(TableLockError::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source.as_ref() != patch.target.as_ref()
                || patch.source_fingerprint != patch.target_fingerprint
            {
                return Err(TableLockError::PatchConflict);
            }
            return Ok(TableLockCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: TableLockDiagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() || fingerprint(&patch.target) != patch.target_fingerprint {
            return Err(TableLockError::PatchConflict);
        }

        let candidate =
            Package::from_source_owner_with_options(patch.target.clone(), self.state.options)
                .map_err(map_candidate_read_error)?;
        if candidate.table_lock_at(patch.sheet_position, patch.table_position)? != patch.after {
            return Err(TableLockError::Verification);
        }
        Ok(TableLockCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: TableLockDiagnostics::published(),
        })
    }

    fn resolve_table_selectors<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<(usize, usize), TableLockError> {
        let selected_sheet = self
            .state
            .document
            .sheet(sheet)
            .map_err(|error| map_read_error(ReadError::Semantic(error)))?
            .ok_or(TableLockError::SheetNotFound)?;
        let selected = selected_sheet
            .select(table.into())
            .map_err(|_error| TableLockError::InvalidSource)?
            .ok_or(TableLockError::TableNotFound)?;
        let table_position = selected_sheet
            .tables()
            .position(|candidate| std::ptr::eq(candidate, selected))
            .ok_or(TableLockError::InvalidSource)?;
        Ok((selected_sheet.index(), table_position))
    }

    fn table_lock_at(
        &self,
        sheet_position: usize,
        table_position: usize,
    ) -> Result<State, TableLockError> {
        let target = self.native_table_lock_target(sheet_position, table_position)?;
        Ok(State::from_locked(target.explicit_locked.unwrap_or(false)))
    }

    fn native_table_lock_target(
        &self,
        sheet_position: usize,
        table_position: usize,
    ) -> Result<NativeTarget, TableLockError> {
        let document = Self::root_document(&self.state.components).map_err(map_read_error)?;
        let sheet_reference = document
            .sheets
            .get(sheet_position)
            .ok_or(TableLockError::SheetNotFound)?;
        let sheet_object = self
            .state
            .index
            .resolve_ref_id(&self.state.components, sheet_reference.identifier)
            .map_err(map_read_error)?
            .ok_or(TableLockError::InvalidSource)?;
        let owner_message_index = unique_sheet_message_index(sheet_object.messages)?;
        let owner_message_type = sheet_object
            .messages
            .get(owner_message_index)
            .ok_or(TableLockError::InvalidSource)?
            .type_;
        let owner_component_index = sheet_object.component_index;
        let owner_object_index = sheet_object.object_index;
        let path = super::SemanticPath::Sheet {
            index: sheet_position,
        };
        let decoded_sheet =
            super::decode_sheet_payload(sheet_object.messages, path).map_err(map_read_error)?;

        let mut table_index = 0usize;
        for drawable in decoded_sheet.drawable_infos {
            let resolved = self
                .state
                .index
                .resolve_ref_id(&self.state.components, drawable.identifier)
                .map_err(map_read_error)?
                .ok_or(TableLockError::InvalidSource)?;
            let Some((message_index, message)) = unique_table_info(resolved)? else {
                continue;
            };
            let snapshot = decode_table_info(&message.data)?;
            if table_index == table_position {
                return Ok(NativeTarget {
                    drawable_identifier: drawable.identifier,
                    model_identifier: snapshot.table_model().identifier().get(),
                    owner_identifier: sheet_reference.identifier,
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
            table_index = table_index
                .checked_add(1)
                .ok_or(TableLockError::InvalidSource)?;
        }
        Err(TableLockError::TableNotFound)
    }
}

fn unique_sheet_message_index(messages: &[RawMessage]) -> Result<usize, TableLockError> {
    let sheet = unique_message_index(messages, super::SHEET_MESSAGE_TYPE)?;
    let form = unique_message_index(messages, super::FORM_BASED_SHEET_MESSAGE_TYPE)?;
    match (sheet, form) {
        (Some(_), Some(_)) | (None, None) => Err(TableLockError::InvalidSource),
        (Some((index, _)), None) | (None, Some((index, _))) => Ok(index),
    }
}

fn unique_table_info(
    resolved: Resolved<'_>,
) -> Result<Option<(usize, &RawMessage)>, TableLockError> {
    let canonical = unique_message_index(resolved.messages, TABLE_INFO_MESSAGE_TYPE)?;
    let legacy = unique_message_index(resolved.messages, LEGACY_TABLE_INFO_MESSAGE_TYPE)?;
    match (canonical, legacy) {
        (Some(_), Some(_)) => Err(TableLockError::InvalidSource),
        (Some(message), None) | (None, Some(message)) => Ok(Some(message)),
        (None, None) => Ok(None),
    }
}

fn unique_message_index(
    messages: &[RawMessage],
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, TableLockError> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == message_type);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(TableLockError::InvalidSource);
    }
    Ok(first)
}

fn decode_table_info(source: &[u8]) -> Result<table_info_codec::TableInfoSnapshot, TableLockError> {
    table_info_codec::decode_table_info(source, table_info_decode_options(source)).map_err(
        |error| {
            if let Some((observed, maximum)) = error.field_limit_values() {
                return TableLockError::LimitExceeded {
                    kind: TableLockLimitKind::WireFields,
                    observed: usize_as_u64(observed),
                    maximum: usize_as_u64(maximum),
                };
            }
            if let Some((observed, maximum)) = error.work_limit_values() {
                return TableLockError::LimitExceeded {
                    kind: TableLockLimitKind::WireWork,
                    observed: usize_as_u64(observed),
                    maximum: usize_as_u64(maximum),
                };
            }
            match error.wire_resource_limit() {
                Some(WireResourceLimit::Bytes {
                    observed,
                    maximum: Some(maximum),
                }) => TableLockError::LimitExceeded {
                    kind: TableLockLimitKind::WireBytes,
                    observed: usize_as_u64(observed.unwrap_or_else(|| maximum.saturating_add(1))),
                    maximum: usize_as_u64(maximum),
                },
                Some(WireResourceLimit::Nesting {
                    observed,
                    maximum: Some(maximum),
                }) => TableLockError::LimitExceeded {
                    kind: TableLockLimitKind::WireNesting,
                    observed: u64::from(observed.unwrap_or_else(|| maximum.saturating_add(1))),
                    maximum: u64::from(maximum),
                },
                _ => TableLockError::InvalidSource,
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
) -> Result<Package, TableLockError> {
    let target = source.native_table_lock_target(sheet_position, table_position)?;
    if State::from_locked(target.explicit_locked.unwrap_or(false)) != before {
        return Err(TableLockError::InvalidSource);
    }
    let source_catalog = physical_source(source)?;
    validate_selected_ownership(source, target)?;
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(TableLockError::InvalidSource)?;
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(TableLockError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(TableLockError::UnsupportedSource);
    }
    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    drop(stream);
    let object = archive
        .objects
        .get_mut(target.object_index)
        .ok_or(TableLockError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier) {
        return Err(TableLockError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.message_index)
        .ok_or(TableLockError::InvalidSource)?;
    if message.type_ != target.message_type {
        return Err(TableLockError::InvalidSource);
    }
    let snapshot = decode_table_info(&message.data)?;
    validate_selected_metadata(object, target.message_index)?;
    if snapshot.locked() != target.explicit_locked {
        return Err(TableLockError::InvalidSource);
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
        return Err(TableLockError::Verification);
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
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    drop(rewritten);
    let output = source_catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            physical_limits,
        )
        .map_err(map_archive_error)?;
    drop(compressed);
    let candidate = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(map_candidate_read_error)?;
    if candidate.table_lock_at(sheet_position, table_position)? != after {
        return Err(TableLockError::Verification);
    }
    Ok(candidate)
}

fn physical_source(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, TableLockError> {
    package
        .state
        .components
        .physical()
        .ok_or(TableLockError::UnsupportedSource)
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
) -> Result<(), TableLockError> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(TableLockError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(TableLockError::InvalidSource);
    }
    Ok(())
}

fn validate_selected_ownership(
    package: &Package,
    target: NativeTarget,
) -> Result<(), TableLockError> {
    let catalog = package.state.components.catalog();
    let sheet_object = catalog
        .get_index(target.owner_component_index)
        .and_then(|component| component.archive().objects.get(target.owner_object_index))
        .ok_or(TableLockError::InvalidSource)?;
    if sheet_object.archive_info.identifier != Some(target.owner_identifier) {
        return Err(TableLockError::InvalidSource);
    }
    let sheet = super::decode_sheet_payload(&sheet_object.messages, super::SemanticPath::Package)
        .map_err(map_read_error)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|drawable| drawable.identifier == target.drawable_identifier)
        .count()
        != 1
    {
        return Err(TableLockError::InvalidSource);
    }
    let sheet_message = sheet_object
        .archive_info
        .message_infos
        .get(target.owner_message_index)
        .ok_or(TableLockError::InvalidSource)?;
    if sheet_message.type_ != target.owner_message_type {
        return Err(TableLockError::InvalidSource);
    }
    let owner_path: &[u32] = match target.owner_message_type {
        super::SHEET_MESSAGE_TYPE => &[2],
        super::FORM_BASED_SHEET_MESSAGE_TYPE => &[1, 2],
        _ => return Err(TableLockError::InvalidSource),
    };
    if !message_declares_reference(sheet_message, target.drawable_identifier, owner_path)? {
        return Err(TableLockError::InvalidSource);
    }

    let table_object = catalog
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(TableLockError::InvalidSource)?;
    if table_object.archive_info.identifier != Some(target.drawable_identifier) {
        return Err(TableLockError::InvalidSource);
    }
    let table_message = table_object
        .archive_info
        .message_infos
        .get(target.message_index)
        .ok_or(TableLockError::InvalidSource)?;
    if table_message.type_ != target.message_type {
        return Err(TableLockError::InvalidSource);
    }
    if !message_declares_reference(table_message, target.model_identifier, &[2])? {
        return Err(TableLockError::InvalidSource);
    }
    Ok(())
}

fn message_declares_reference(
    message: &litchi_iwa_core::MessageInfo,
    identifier: u64,
    accepted_field_path: &[u32],
) -> Result<bool, TableLockError> {
    let mut declared = message.object_references.contains(&identifier);
    for field in &message.field_infos {
        if field.object_references.contains(&identifier) {
            if field.path.as_slice() != accepted_field_path {
                return Err(TableLockError::InvalidSource);
            }
            declared = true;
        }
    }
    Ok(declared)
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), TableLockError> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset)
            .map_err(|_error| TableLockError::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(TableLockError::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_error| TableLockError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(TableLockError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes).map_err(|_error| TableLockError::InvalidSource)?,
            )
            .ok_or(TableLockError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(TableLockError::InvalidSource)?
                != object.data_offset
        {
            return Err(TableLockError::InvalidSource);
        }
    }
    Ok(())
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn map_candidate_read_error(error: ReadError) -> TableLockError {
    match error {
        ReadError::InputTooLarge { observed, maximum }
        | ReadError::Archive(litchi_iwa_archive::Error::Limit {
            kind: litchi_iwa_archive::LimitKind::InputBytes,
            observed,
            maximum,
        }) => TableLockError::LimitExceeded {
            kind: TableLockLimitKind::OutputBytes,
            observed,
            maximum,
        },
        other @ (ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::Archive(_)
        | ReadError::MalformedPayload { .. }
        | ReadError::NotNumbers
        | ReadError::Common(_)
        | ReadError::InvalidFormat(_)
        | ReadError::ParseError(_)
        | ReadError::Semantic(_)
        | ReadError::SemanticLimit { .. }) => map_read_error(other),
    }
}

fn map_read_error(error: ReadError) -> TableLockError {
    match error {
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Common(wire_error) => map_wire_error(wire_error),
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => TableLockError::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => TableLockLimitKind::PayloadObjects,
                super::SemanticLimitKind::References => TableLockLimitKind::PayloadReferences,
                super::SemanticLimitKind::OutputTextBytes
                | super::SemanticLimitKind::FormulaWireBytes
                | super::SemanticLimitKind::TextBytes => TableLockLimitKind::PayloadBytes,
                super::SemanticLimitKind::FormulaRenderDepth
                | super::SemanticLimitKind::FormulaDepth => TableLockLimitKind::WireNesting,
                super::SemanticLimitKind::FormulaRenderWork
                | super::SemanticLimitKind::FormulaWork => TableLockLimitKind::WireWork,
                super::SemanticLimitKind::Sheets
                | super::SemanticLimitKind::Tables
                | super::SemanticLimitKind::MaterializedCells => TableLockLimitKind::PayloadItems,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        ReadError::InputTooLarge { observed, maximum } => TableLockError::LimitExceeded {
            kind: TableLockLimitKind::InputBytes,
            observed,
            maximum,
        },
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::MalformedPayload { .. }
        | ReadError::NotNumbers
        | ReadError::InvalidFormat(_)
        | ReadError::ParseError(_)
        | ReadError::Semantic(_) => TableLockError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> TableLockError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => TableLockError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => TableLockLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => TableLockLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => TableLockLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => TableLockLimitKind::PackageBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => TableLockLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => TableLockLimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => TableLockLimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    TableLockLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            TableLockError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Reassembly(_) => TableLockError::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::InvalidBundle(_) => TableLockError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "this mapper is passed directly to Result::map_err"
)]
fn map_core_error(error: litchi_iwa_core::Error) -> TableLockError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => TableLockError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => TableLockLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    TableLockLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => TableLockLimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => TableLockLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    TableLockLimitKind::PayloadBytes
                },
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            TableLockError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => TableLockError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "this mapper is passed directly to Result::map_err"
)]
fn map_wire_error(error: litchi_iwa_common::Error) -> TableLockError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => TableLockError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => TableLockLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => TableLockLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => TableLockLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => TableLockLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => TableLockLimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    TableLockLimitKind::PayloadItems
                },
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            TableLockError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => TableLockError::InvalidSource,
    }
}
