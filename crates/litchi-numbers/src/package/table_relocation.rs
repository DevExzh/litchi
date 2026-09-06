//! Exact-source ownership of moving one rooted Numbers table between sheets.
//!
//! A relocation is a graph operation, not a semantic table copy.  The source
//! sheet loses one `TSP.Reference`, the destination sheet gains the exact
//! reference payload, and the table-info drawable's parent is changed to the
//! destination sheet.  All three owners are resolved from the rooted
//! document before any bytes are staged.  Rewrites are grouped by physical
//! component so a component shared by two owners is decompressed, encoded,
//! and reassembled only once.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused transaction keeps its public value types beside the exact-source machinery."
)]

use std::fmt;

use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, OwnedExactArtifacts},
};
use litchi_iwa_common::wire::{
    patch_length_delimited_field, patch_nested_varint_field,
    rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{Archive, ArchiveObject, Error as CoreError, RawMessage, SnappyStream};
use litchi_iwa_protos::table_info_codec;
use thiserror::Error as ThisError;

use super::{Package, table_headers};
use crate::{
    selector::{SheetSelector, TableSelector},
    table::lock::State as LockState,
};

const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const LEGACY_TABLE_INFO_MESSAGE_TYPE: u32 = 6_003;
const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// A content-free location associated with a table-relocation transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// A source table and its existing destination sheet.
    Table {
        /// Zero-based source sheet position.
        source_sheet: usize,
        /// Zero-based table position within the source sheet.
        table: usize,
        /// Zero-based destination sheet position.
        destination_sheet: usize,
    },
}

/// A finite resource governed by table relocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source package bytes.
    InputBytes,
    /// Complete candidate package bytes.
    OutputBytes,
    /// Physical package members.
    Entries,
    /// Bytes in one physical member.
    EntryBytes,
    /// Aggregate physical member bytes.
    TotalEntryBytes,
    /// Bytes in a decoded IWA payload.
    PayloadBytes,
    /// Aggregate decoded IWA payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native metadata or reference items inspected.
    PayloadItems,
    /// Protobuf fields and rewrite work.
    WireWork,
    /// Aggregate transaction work.
    TransactionWork,
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
            Self::WireWork => "wire work",
            Self::TransactionWork => "transaction work",
        })
    }
}

/// A content-redacted table-relocation failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No rooted sheet matched a selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No table on the selected source sheet matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// A changed edit targeted an effectively locked table.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked { path: Path },
    /// The exact physical source cannot be rewritten safely.
    #[error("this Numbers source does not support exact table relocation")]
    UnsupportedSource,
    /// Rooted ownership, metadata, or wire framing is invalid.
    #[error("the Numbers table-relocation source is invalid at {path:?}")]
    InvalidSource { path: Path },
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Numbers table relocation {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
        /// Content-free location where the limit was observed.
        path: Path,
    },
    /// A bounded allocation failed before publication.
    #[error("could not allocate {amount} Numbers table-relocation units at {path:?}")]
    Allocation { amount: usize, path: Path },
    /// Candidate reopening or focused locality verification failed.
    #[error("the edited Numbers table relocation failed semantic verification")]
    Verification,
    /// The patch was produced from a different exact source artifact.
    #[error("the Numbers table-relocation patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Operation {
    pub(super) source_sheet: usize,
    pub(super) source_table: usize,
    pub(super) destination_sheet: usize,
    pub(super) destination_table: usize,
    pub(super) source_sheet_identifier: u64,
    pub(super) destination_sheet_identifier: u64,
    pub(super) drawable_identifier: u64,
    pub(super) model_identifier: u64,
}

impl Operation {
    pub(super) const fn path(self) -> Path {
        Path::Table {
            source_sheet: self.source_sheet,
            table: self.source_table,
            destination_sheet: self.destination_sheet,
        }
    }

    const fn inverse(self) -> Self {
        Self {
            source_sheet: self.destination_sheet,
            source_table: self.destination_table,
            destination_sheet: self.source_sheet,
            destination_table: self.source_table,
            source_sheet_identifier: self.destination_sheet_identifier,
            destination_sheet_identifier: self.source_sheet_identifier,
            drawable_identifier: self.drawable_identifier,
            model_identifier: self.model_identifier,
        }
    }

    pub(super) const fn is_same_sheet(self) -> bool {
        self.source_sheet == self.destination_sheet
            && self.source_sheet_identifier == self.destination_sheet_identifier
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct SheetTarget {
    pub(super) position: usize,
    pub(super) identifier: u64,
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
}

/// One immutable selector-first table-relocation edit.
pub struct Edit<'a> {
    source: &'a Package,
    operation: Operation,
    table: table_headers::Target,
    source_sheet: SheetTarget,
    destination_sheet: SheetTarget,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path())
            .field("destination_table", &self.operation.destination_table)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the semantic source/destination path.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.operation.path()
    }

    /// Return the source sheet's zero-based position.
    #[must_use]
    pub const fn source_sheet_position(&self) -> usize {
        self.operation.source_sheet
    }

    /// Return the source table's zero-based position.
    #[must_use]
    pub const fn table_position(&self) -> usize {
        self.operation.source_table
    }

    /// Return the destination sheet's zero-based position.
    #[must_use]
    pub const fn destination_sheet_position(&self) -> usize {
        self.operation.destination_sheet
    }

    /// Return the table position after an append to the destination sheet.
    #[must_use]
    pub const fn destination_table_position(&self) -> usize {
        self.operation.destination_table
    }

    /// Publish the exact-source edit atomically.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// A reversible patch bound to exact source and target package artifacts.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    operation: Operation,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path())
            .field("destination_table", &self.operation.destination_table)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the semantic source/destination path.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.operation.path()
    }

    /// Return the source sheet's zero-based position.
    #[must_use]
    pub const fn source_sheet_position(&self) -> usize {
        self.operation.source_sheet
    }

    /// Return the source table's zero-based position.
    #[must_use]
    pub const fn table_position(&self) -> usize {
        self.operation.source_table
    }

    /// Return the destination sheet's zero-based position.
    #[must_use]
    pub const fn destination_sheet_position(&self) -> usize {
        self.operation.destination_sheet
    }

    /// Return the table position after relocation.
    #[must_use]
    pub const fn destination_table_position(&self) -> usize {
        self.operation.destination_table
    }

    /// Return a stable, non-authorizing source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return a stable, non-authorizing target fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Whether this patch leaves the exact package artifact unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.operation.is_same_sheet() && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            operation: if self.operation.is_same_sheet() {
                self.operation
            } else {
                self.operation.inverse()
            },
        }
    }
}

/// Content-free diagnostics for one completed relocation.
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

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of root preview assets deleted in this direction.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the candidate artifact was fully reopened and checked.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One fully verified immutable relocation publication.
#[must_use = "a table-relocation commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start a selector-first edit that moves one source-sheet table to an
    /// existing destination sheet.
    ///
    /// The table selector is local to `source_sheet`, and the destination is
    /// appended after the destination sheet's existing semantic tables. A
    /// same-sheet request is an exact no-op after strict source resolution.
    pub fn edit_table_relocation<'source_sheet, 'table, 'destination_sheet>(
        &self,
        source_sheet: impl Into<SheetSelector<'source_sheet>>,
        table: impl Into<TableSelector<'table>>,
        destination_sheet: impl Into<SheetSelector<'destination_sheet>>,
    ) -> Result<Edit<'_>, Error> {
        let source_sheet_semantic = self
            .state
            .document
            .sheet(source_sheet.into())
            .map_err(|_error| invalid_source(Path::Package))?
            .ok_or(Error::SheetNotFound)?;
        let source_sheet_position = source_sheet_semantic.index();
        let source_table_position = resolve_table_position(
            source_sheet_semantic,
            table.into(),
            Path::Table {
                source_sheet: source_sheet_position,
                table: 0,
                destination_sheet: source_sheet_position,
            },
        )?;
        let destination_sheet_semantic = self
            .state
            .document
            .sheet(destination_sheet.into())
            .map_err(|_error| invalid_source(Path::Package))?
            .ok_or(Error::SheetNotFound)?;
        let destination_sheet_position = destination_sheet_semantic.index();
        let path = Path::Table {
            source_sheet: source_sheet_position,
            table: source_table_position,
            destination_sheet: destination_sheet_position,
        };
        let table_target =
            resolve_table_target(self, source_sheet_position, source_table_position, path)?;
        let source_target = resolve_sheet_target(self, source_sheet_position, path)?;
        let destination_target = resolve_sheet_target(self, destination_sheet_position, path)?;
        if source_target.identifier != table_target.sheet_identifier {
            return Err(invalid_source(path));
        }
        let destination_table_position =
            if source_target.identifier == destination_target.identifier {
                source_table_position
            } else {
                destination_sheet_semantic.tables().len()
            };
        let operation = Operation {
            source_sheet: source_sheet_position,
            source_table: source_table_position,
            destination_sheet: destination_sheet_position,
            destination_table: destination_table_position,
            source_sheet_identifier: source_target.identifier,
            destination_sheet_identifier: destination_target.identifier,
            drawable_identifier: table_target.drawable_identifier,
            model_identifier: table_target.model_identifier,
        };
        Ok(Edit {
            source: self,
            operation,
            table: table_target,
            source_sheet: source_target,
            destination_sheet: destination_target,
        })
    }

    /// Move one rooted table and publish the validated immutable package.
    pub fn move_table<'source_sheet, 'table, 'destination_sheet>(
        &self,
        source_sheet: impl Into<SheetSelector<'source_sheet>>,
        table: impl Into<TableSelector<'table>>,
        destination_sheet: impl Into<SheetSelector<'destination_sheet>>,
    ) -> Result<Commit, Error> {
        self.edit_table_relocation(source_sheet, table, destination_sheet)?
            .commit()
    }

    /// Apply a reversible exact-source relocation patch.
    pub fn apply_table_relocation(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_source(self)?;
        let source = catalog.__source_owner();
        if !patch.artifacts.authorizes_owner(&source) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        verify_source_state(self, patch.operation)?;
        let target_owner = patch.artifacts.target_owner();
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_error| Error::Verification)?;
        verify_target_state(&candidate, patch.operation)?;
        verify_locality(self, &candidate, patch.operation)?;
        let source_previews = root_preview_deletions(catalog)?;
        let candidate_previews = root_preview_deletions(physical_source(&candidate)?)?;
        let current_table = resolve_table_target(
            self,
            patch.operation.source_sheet,
            patch.operation.source_table,
            patch.operation.path(),
        )?;
        let current_source_sheet =
            resolve_sheet_target(self, patch.operation.source_sheet, patch.operation.path())?;
        let current_destination_sheet = resolve_sheet_target(
            self,
            patch.operation.destination_sheet,
            patch.operation.path(),
        )?;
        let touched_components = touched_component_count(
            current_table,
            current_source_sheet,
            current_destination_sheet,
        );
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                touched_components,
                source_previews
                    .len()
                    .saturating_sub(candidate_previews.len()),
            ),
        })
    }

    /// Compatibility spelling for callers that name the operation as a move.
    pub fn apply_table_move(&self, patch: &Patch) -> Result<Commit, Error> {
        self.apply_table_relocation(patch)
    }
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let catalog = physical_source(edit.source)?;
    let source_owner = catalog.__source_owner();
    if edit.operation.is_same_sheet() {
        return Ok(Commit {
            package: edit.source.snapshot(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                operation: edit.operation,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if edit.table.locked == LockState::Locked {
        return Err(Error::TableLocked { path: edit.path() });
    }
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    verify_source_state(edit.source, edit.operation)?;
    let source_previews = root_preview_deletions(catalog)?;
    let package = rewrite(
        edit.source,
        edit.operation,
        edit.table,
        edit.source_sheet,
        edit.destination_sheet,
        &source_previews,
    )?;
    verify_target_state(&package, edit.operation)?;
    verify_locality(edit.source, &package, edit.operation)?;
    let target_owner = physical_source(&package)?.__source_owner();
    let touched_components =
        touched_component_count(edit.table, edit.source_sheet, edit.destination_sheet);
    Ok(Commit {
        package,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            operation: edit.operation,
        },
        diagnostics: Diagnostics::published(touched_components, source_previews.len()),
    })
}

fn resolve_table_position(
    sheet: &crate::Sheet,
    selector: TableSelector<'_>,
    path: Path,
) -> Result<usize, Error> {
    match selector {
        TableSelector::Index(index) => sheet
            .tables()
            .nth(index)
            .map(|_table| index)
            .ok_or(Error::TableNotFound),
        TableSelector::Name(name) => {
            let mut matches = sheet
                .tables()
                .enumerate()
                .filter(|(_index, table)| table.name() == name);
            let Some((index, _table)) = matches.next() else {
                return Err(Error::TableNotFound);
            };
            if matches.next().is_some() {
                return Err(invalid_source(path));
            }
            Ok(index)
        },
    }
}

fn resolve_table_target(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    path: Path,
) -> Result<table_headers::Target, Error> {
    // Relocation changes only rooted sheet ownership and the TableInfo parent.
    // Header/storage settings are intentionally outside this transaction, so
    // admission validates the physical table graph without projecting cells.
    table_headers::resolve::resolve_target_physical(source, sheet_position, table_position)
        .map_err(|error| map_header_error(error, path))
}

pub(super) fn resolve_sheet_target(
    source: &Package,
    sheet_position: usize,
    path: Path,
) -> Result<SheetTarget, Error> {
    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or_else(|| invalid_source(path))?;
    let (_message_index, document_message) = table_headers::resolve::unique_message_index(
        &document_object.messages,
        super::DOCUMENT_MESSAGE_TYPE,
    )
    .map_err(|error| map_header_error(error, path))?
    .ok_or_else(|| invalid_source(path))?;
    let sheet_payloads =
        table_headers::resolve::repeated_length_payloads(&document_message.data, 1)
            .map_err(|error| map_header_error(error, path))?;
    let sheet_identifier = table_headers::resolve::local_reference_identifier(
        sheet_payloads
            .get(sheet_position)
            .ok_or(Error::SheetNotFound)?,
    )
    .map_err(|error| map_header_error(error, path))?;
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, sheet_identifier)
        .map_err(|_error| invalid_source(path))?
        .ok_or_else(|| invalid_source(path))?;
    let message_index = table_headers::resolve::unique_sheet_message_index(resolved.messages)
        .map_err(|error| map_header_error(error, path))?;
    let message = resolved
        .messages
        .get(message_index)
        .ok_or_else(|| invalid_source(path))?;
    let object = source
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or_else(|| invalid_source(path))?;
    table_headers::resolve::validate_message_metadata(object, message_index)
        .map_err(|error| map_header_error(error, path))?;
    if object.archive_info.identifier != Some(sheet_identifier) {
        return Err(invalid_source(path));
    }
    Ok(SheetTarget {
        position: sheet_position,
        identifier: sheet_identifier,
        component_index: resolved.component_index,
        object_index: resolved.object_index,
        message_index,
        message_type: message.type_,
    })
}

fn rewrite(
    source: &Package,
    operation: Operation,
    table: table_headers::Target,
    source_sheet: SheetTarget,
    destination_sheet: SheetTarget,
    deleted_previews: &[&str],
) -> Result<Package, Error> {
    let output = rewrite_bytes(
        source,
        operation,
        table,
        source_sheet,
        destination_sheet,
        deleted_previews,
    )?;
    Package::from_owned_bytes_with_options(output, source.state.options)
        .map_err(|_error| Error::Verification)
}

/// Rewrite the exact physical relocation artifact without semantic reopening.
pub(super) fn rewrite_bytes(
    source: &Package,
    operation: Operation,
    table: table_headers::Target,
    source_sheet: SheetTarget,
    destination_sheet: SheetTarget,
    deleted_previews: &[&str],
) -> Result<Vec<u8>, Error> {
    rewrite_bytes_with_parent_policy(
        source,
        operation,
        table,
        source_sheet,
        destination_sheet,
        deleted_previews,
        false,
    )
}

/// Rewrite a physical relocation for the legacy migration host.
///
/// Source-built host packages can contain a rooted table-info drawable whose
/// payload intentionally omits the optional parent edge. The public owner
/// requires that edge for a strict relocation, while this compatibility seam
/// preserves the producer-authored omission and still moves the rooted sheet
/// ownership. All archive, wire, and verification work remains in this owner.
#[cfg(feature = "internal-iwork-source")]
pub(super) fn rewrite_bytes_for_compatibility(
    source: &Package,
    operation: Operation,
    table: table_headers::Target,
    source_sheet: SheetTarget,
    destination_sheet: SheetTarget,
    deleted_previews: &[&str],
) -> Result<Vec<u8>, Error> {
    rewrite_bytes_with_parent_policy(
        source,
        operation,
        table,
        source_sheet,
        destination_sheet,
        deleted_previews,
        true,
    )
}

fn rewrite_bytes_with_parent_policy(
    source: &Package,
    operation: Operation,
    table: table_headers::Target,
    source_sheet: SheetTarget,
    destination_sheet: SheetTarget,
    deleted_previews: &[&str],
    preserve_missing_parent: bool,
) -> Result<Vec<u8>, Error> {
    let source_catalog = physical_source(source)?;
    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(|_error| invalid_source(operation.path()))?;
    let reference_payload = table_reference_payload(source, table, operation.path())?;
    let mut component_indices = [
        table.sheet_component_index,
        table.info_component_index,
        source_sheet.component_index,
        destination_sheet.component_index,
    ];
    component_indices.sort_unstable();
    let mut rewritten = Vec::new();
    rewritten
        .try_reserve_exact(component_indices.len())
        .map_err(|_error| Error::Allocation {
            amount: component_indices.len(),
            path: operation.path(),
        })?;
    let mut previous_component = None;
    for component_index in component_indices {
        if previous_component == Some(component_index) {
            continue;
        }
        previous_component = Some(component_index);
        let component = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or_else(|| invalid_source(operation.path()))?;
        let name = component.name();
        let entry = source_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == name)
            .ok_or_else(|| invalid_source(operation.path()))?;
        if entry.is_opaque() {
            return Err(Error::UnsupportedSource);
        }
        let stream = SnappyStream::decompress_with_limits(
            entry.data(),
            physical_limits
                .snappy_limits()
                .map_err(|_error| invalid_source(operation.path()))?,
        )
        .map_err(|error| map_core_error(error, operation.path()))?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(|error| map_core_error(error, operation.path()))?;
        archive
            .validate_canonical_object_framing(stream.as_bytes())
            .map_err(|error| map_core_error(error, operation.path()))?;
        if source_sheet.component_index == component_index {
            mutate_sheet(
                &mut archive,
                source_sheet,
                operation.drawable_identifier,
                &reference_payload,
                false,
                physical_limits,
                operation.path(),
            )?;
        }
        if destination_sheet.component_index == component_index {
            mutate_sheet(
                &mut archive,
                destination_sheet,
                operation.drawable_identifier,
                &reference_payload,
                true,
                physical_limits,
                operation.path(),
            )?;
        }
        if table.info_component_index == component_index {
            mutate_table_info(
                &mut archive,
                table,
                operation.source_sheet_identifier,
                operation.destination_sheet_identifier,
                operation.model_identifier,
                preserve_missing_parent,
                physical_limits,
                operation.path(),
            )?;
        }
        let encoded = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|error| map_core_error(error, operation.path()))?;
        let compressed = SnappyStream::compress(&encoded)
            .map_err(|error| map_core_error(error, operation.path()))?;
        rewritten.push((name, compressed));
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(rewritten.len())
        .map_err(|_error| Error::Allocation {
            amount: rewritten.len(),
            path: operation.path(),
        })?;
    for (name, data) in &rewritten {
        edits.push(EntryEdit::new(name, data.as_slice()));
    }
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, deleted_previews, physical_limits)
        .map_err(|_error| Error::UnsupportedSource)?;
    Ok(output)
}

fn mutate_sheet(
    archive: &mut Archive,
    target: SheetTarget,
    drawable_identifier: u64,
    reference_payload: &[u8],
    adding: bool,
    limits: litchi_iwa_archive::Limits,
    path: Path,
) -> Result<(), Error> {
    let object = archive
        .objects
        .get_mut(target.object_index)
        .ok_or_else(|| invalid_source(path))?;
    if object.archive_info.identifier != Some(target.identifier) {
        return Err(invalid_source(path));
    }
    let message = object
        .messages
        .get(target.message_index)
        .ok_or_else(|| invalid_source(path))?;
    if message.type_ != target.message_type {
        return Err(invalid_source(path));
    }
    let data = rewrite_sheet_data(
        &message.data,
        target.message_type,
        drawable_identifier,
        reference_payload,
        adding,
        path,
    )?;
    let transition = if adding {
        reference_transition_append(
            object,
            target.message_index,
            target.message_type,
            drawable_identifier,
            path,
        )?
    } else {
        reference_transition_remove(
            object,
            target.message_index,
            target.message_type,
            drawable_identifier,
            path,
        )?
    };
    replace_with_transition(object, target.message_index, data, transition, limits, path)
}

fn mutate_table_info(
    archive: &mut Archive,
    target: table_headers::Target,
    source_sheet_identifier: u64,
    destination_sheet_identifier: u64,
    model_identifier: u64,
    preserve_missing_parent: bool,
    limits: litchi_iwa_archive::Limits,
    path: Path,
) -> Result<(), Error> {
    let object = archive
        .objects
        .get_mut(target.info_object_index)
        .ok_or_else(|| invalid_source(path))?;
    if object.archive_info.identifier != Some(target.drawable_identifier) {
        return Err(invalid_source(path));
    }
    let message = object
        .messages
        .get(target.info_message_index)
        .ok_or_else(|| invalid_source(path))?;
    if ![TABLE_INFO_MESSAGE_TYPE, LEGACY_TABLE_INFO_MESSAGE_TYPE].contains(&message.type_) {
        return Err(invalid_source(path));
    }
    let parent = table_parent_identifier(&message.data, path)?;
    if parent != Some(source_sheet_identifier) && !(preserve_missing_parent && parent.is_none()) {
        return Err(invalid_source(path));
    }
    let before_model = table_model_identifier(&message.data, path)?;
    if before_model != model_identifier {
        return Err(invalid_source(path));
    }
    if parent.is_none() {
        // A compatibility source may omit the optional TableInfo parent edge.
        // Keep that producer-authored shape unchanged; a metadata-only edge
        // would be ambiguous because there is no payload edge to relocate.
        let transition = reference_transition_replace(
            object,
            target.info_message_index,
            source_sheet_identifier,
            destination_sheet_identifier,
            path,
        )?;
        if transition.is_some() {
            return Err(invalid_source(path));
        }
        let data = try_copy_bytes(&message.data, path)?;
        object
            .replace_message_preserving_header_with_limits(
                target.info_message_index,
                RawMessage {
                    type_: message.type_,
                    data,
                },
                limits
                    .effective_archive_limits()
                    .map_err(|_error| invalid_source(path))?,
            )
            .map_err(|error| map_core_error(error, path))
            .map(|_old| ())?;
        return Ok(());
    }
    let data = patch_table_parent(
        &message.data,
        source_sheet_identifier,
        destination_sheet_identifier,
        model_identifier,
        path,
    )?;
    let message_type = message.type_;
    let transition = reference_transition_replace(
        object,
        target.info_message_index,
        source_sheet_identifier,
        destination_sheet_identifier,
        path,
    )?;
    match transition {
        Some(transition) => {
            // The core transition primitive introduces new identifiers as a
            // suffix so retained raw reference occurrences stay attributable.
            // TableInfo parent ownership is positional, however: preserve the
            // parent's original aggregate slot with a second, fully checked
            // reorder. This also makes two independent legacy moves restore
            // canonical parent/model metadata order; public patch inversion
            // remains exact-source and byte-backed.
            let desired_order = replacement_reference_order(
                &transition.aggregate_before,
                source_sheet_identifier,
                destination_sheet_identifier,
                path,
            )?;
            let needs_reorder = desired_order != transition.aggregate_after;
            let transition_data = if needs_reorder {
                try_copy_bytes(&data, path)?
            } else {
                Vec::new()
            };
            if needs_reorder {
                replace_with_transition(
                    object,
                    target.info_message_index,
                    transition_data,
                    transition,
                    limits,
                    path,
                )?;
                object
                    .replace_message_reordering_object_references_preserving_header_with_limits(
                        target.info_message_index,
                        RawMessage {
                            type_: message_type,
                            data,
                        },
                        &desired_order,
                        limits
                            .effective_archive_limits()
                            .map_err(|_error| invalid_source(path))?,
                    )
                    .map_err(|error| map_core_error(error, path))?;
            } else {
                replace_with_transition(
                    object,
                    target.info_message_index,
                    data,
                    transition,
                    limits,
                    path,
                )?;
            }
            Ok(())
        },
        // Some Apple-produced TableInfo messages carry the parent edge in
        // the payload but intentionally omit it from ArchiveInfo reference
        // metadata. Preserve that header verbatim after changing the
        // payload; attempting to synthesize a metadata transition would
        // either reject the valid shape or invent an ownership edge.
        None => object
            .replace_message_preserving_header_with_limits(
                target.info_message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
                limits
                    .effective_archive_limits()
                    .map_err(|_error| invalid_source(path))?,
            )
            .map_err(|error| map_core_error(error, path))
            .map(|_old| ()),
    }
}

fn replacement_reference_order(
    before: &[u64],
    source_identifier: u64,
    destination_identifier: u64,
    path: Path,
) -> Result<Vec<u64>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(before.len())
        .map_err(|_error| Error::Allocation {
            amount: before.len(),
            path,
        })?;
    let mut replaced = false;
    for identifier in before {
        if *identifier == source_identifier {
            if replaced {
                return Err(invalid_source(path));
            }
            output.push(destination_identifier);
            replaced = true;
        } else {
            output.push(*identifier);
        }
    }
    if !replaced {
        return Err(invalid_source(path));
    }
    Ok(output)
}

fn try_copy_bytes(source: &[u8], path: Path) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_error| Error::Allocation {
            amount: source.len(),
            path,
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

fn rewrite_sheet_data(
    source: &[u8],
    message_type: u32,
    drawable_identifier: u64,
    reference_payload: &[u8],
    adding: bool,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let envelope = match message_type {
        SHEET_MESSAGE_TYPE => source,
        FORM_BASED_SHEET_MESSAGE_TYPE => table_headers::resolve::singular_length_payload(source, 1)
            .map_err(|error| map_header_error(error, path))?,
        _ => return Err(invalid_source(path)),
    };
    let payloads = table_headers::resolve::repeated_length_payloads(envelope, 2)
        .map_err(|error| map_header_error(error, path))?;
    let mut matching = 0usize;
    let replacement_capacity = payloads
        .len()
        .checked_add(usize::from(adding))
        .ok_or_else(|| invalid_source(path))?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(replacement_capacity)
        .map_err(|_error| Error::Allocation {
            amount: replacement_capacity,
            path,
        })?;
    for payload in payloads {
        let identifier = table_headers::resolve::local_reference_identifier(payload)
            .map_err(|error| map_header_error(error, path))?;
        if identifier == drawable_identifier {
            matching = matching
                .checked_add(1)
                .ok_or_else(|| invalid_source(path))?;
            if !adding {
                continue;
            }
            return Err(invalid_source(path));
        }
        replacements.push(try_copy_bytes(payload, path)?);
    }
    if adding {
        replacements.push(try_copy_bytes(reference_payload, path)?);
    } else if matching != 1 {
        return Err(invalid_source(path));
    }
    let rewritten_envelope = rewrite_repeated_length_delimited_fields(envelope, 2, &replacements)
        .map_err(|error| map_wire_error(error, path))?;
    let output = match message_type {
        SHEET_MESSAGE_TYPE => rewritten_envelope,
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            patch_length_delimited_field(source, 1, true, Some(rewritten_envelope.as_slice()))
                .map_err(|error| map_wire_error(error, path))?
        },
        _ => return Err(invalid_source(path)),
    };
    let verified = match message_type {
        SHEET_MESSAGE_TYPE => output.as_slice(),
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            table_headers::resolve::singular_length_payload(&output, 1)
                .map_err(|error| map_header_error(error, path))?
        },
        _ => return Err(invalid_source(path)),
    };
    let verified_payloads = table_headers::resolve::repeated_length_payloads(verified, 2)
        .map_err(|error| map_header_error(error, path))?;
    let verified_count = verified_payloads
        .iter()
        .filter(|payload| {
            table_headers::resolve::local_reference_identifier(payload)
                .is_ok_and(|identifier| identifier == drawable_identifier)
        })
        .count();
    if verified_count != usize::from(adding) {
        return Err(invalid_source(path));
    }
    Ok(output)
}

fn table_reference_payload(
    source: &Package,
    target: table_headers::Target,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(target.sheet_component_index)
        .and_then(|component| component.archive().objects.get(target.sheet_object_index))
        .ok_or_else(|| invalid_source(path))?;
    let message = object
        .messages
        .get(target.sheet_message_index)
        .ok_or_else(|| invalid_source(path))?;
    let payloads = table_headers::resolve::sheet_drawable_payloads(message.type_, &message.data)
        .map_err(|error| map_header_error(error, path))?;
    let mut matching = payloads.iter().filter(|payload| {
        table_headers::resolve::local_reference_identifier(payload)
            .is_ok_and(|identifier| identifier == target.drawable_identifier)
    });
    let payload = matching.next().ok_or_else(|| invalid_source(path))?;
    if matching.next().is_some() {
        return Err(invalid_source(path));
    }
    try_copy_bytes(payload, path)
}

fn table_parent_identifier(source: &[u8], path: Path) -> Result<Option<u64>, Error> {
    let super_payload = table_headers::resolve::singular_length_payload(source, 1)
        .map_err(|error| map_header_error(error, path))?;
    let parents = table_headers::resolve::repeated_length_payloads(super_payload, 2)
        .map_err(|error| map_header_error(error, path))?;
    match parents.as_slice() {
        [] => Ok(None),
        [parent] => table_headers::resolve::local_reference_identifier(parent)
            .map(Some)
            .map_err(|error| map_header_error(error, path)),
        _ => Err(invalid_source(path)),
    }
}

fn table_model_identifier(source: &[u8], path: Path) -> Result<u64, Error> {
    table_info_codec::decode_table_model_reference(source, super::table_info_decode_options(source))
        .map(|reference| reference.identifier().get())
        .map_err(|_error| invalid_source(path))
}

fn patch_table_parent(
    source: &[u8],
    source_sheet_identifier: u64,
    destination_sheet_identifier: u64,
    model_identifier: u64,
    path: Path,
) -> Result<Vec<u8>, Error> {
    if table_parent_identifier(source, path)? != Some(source_sheet_identifier)
        || table_model_identifier(source, path)? != model_identifier
    {
        return Err(invalid_source(path));
    }
    // TableInfoArchive.super.parent is root field 1, DrawableArchive.parent
    // is nested field 2, and TSP.Reference.identifier is leaf field 1.  The
    // wire helper edits only that canonical leaf and copies every unknown
    // field in the two envelopes byte-for-byte.
    let output =
        patch_nested_varint_field(source, &[1, 2, 1], true, Some(destination_sheet_identifier))
            .map_err(|error| map_wire_error(error, path))?;
    if table_parent_identifier(&output, path)? != Some(destination_sheet_identifier)
        || table_model_identifier(&output, path)? != model_identifier
    {
        return Err(invalid_source(path));
    }
    Ok(output)
}

struct FieldTransitionOwned {
    field_info_index: usize,
    path: Vec<u32>,
    before: Vec<u64>,
    after: Vec<u64>,
}

struct ReferenceTransitionOwned {
    aggregate_before: Vec<u64>,
    aggregate_after: Vec<u64>,
    fields: Vec<FieldTransitionOwned>,
}

fn reference_transition_append(
    object: &ArchiveObject,
    message_index: usize,
    message_type: u32,
    identifier: u64,
    path: Path,
) -> Result<ReferenceTransitionOwned, Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or_else(|| invalid_source(path))?;
    if info.object_references.contains(&identifier) {
        return Err(invalid_source(path));
    }
    let aggregate_before = try_copy_slice(&info.object_references, path)?;
    let mut aggregate_after = try_copy_slice(&aggregate_before, path)?;
    aggregate_after
        .try_reserve(1)
        .map_err(|_error| Error::Allocation { amount: 1, path })?;
    aggregate_after.push(identifier);
    let expected_path = sheet_drawable_reference_path(message_type, path)?;
    let mut selected = None;
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        if field.path.as_slice() != expected_path {
            continue;
        }
        if selected.replace((field_info_index, field)).is_some()
            || !is_message_reference_field(field)
            || field.object_references.contains(&identifier)
        {
            return Err(invalid_source(path));
        }
    }
    let mut fields = Vec::new();
    if let Some((field_info_index, field)) = selected {
        fields
            .try_reserve_exact(1)
            .map_err(|_error| Error::Allocation { amount: 1, path })?;
        let before = try_copy_slice(&field.object_references, path)?;
        let mut after = try_copy_slice(&before, path)?;
        after
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1, path })?;
        after.push(identifier);
        fields.push(FieldTransitionOwned {
            field_info_index,
            path: try_copy_slice(field.path.as_slice(), path)?,
            before,
            after,
        });
    }
    Ok(ReferenceTransitionOwned {
        aggregate_before,
        aggregate_after,
        fields,
    })
}

fn reference_transition_remove(
    object: &ArchiveObject,
    message_index: usize,
    message_type: u32,
    identifier: u64,
    path: Path,
) -> Result<ReferenceTransitionOwned, Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or_else(|| invalid_source(path))?;
    if info
        .object_references
        .iter()
        .filter(|reference| **reference == identifier)
        .count()
        != 1
    {
        return Err(invalid_source(path));
    }
    let aggregate_before = try_copy_slice(&info.object_references, path)?;
    let mut aggregate_after = Vec::new();
    aggregate_after
        .try_reserve_exact(aggregate_before.len().saturating_sub(1))
        .map_err(|_error| Error::Allocation {
            amount: aggregate_before.len().saturating_sub(1),
            path,
        })?;
    aggregate_after.extend(
        aggregate_before
            .iter()
            .copied()
            .filter(|reference| *reference != identifier),
    );

    let expected_path = sheet_drawable_reference_path(message_type, path)?;
    let mut selected = None;
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        let occurrences = field
            .object_references
            .iter()
            .filter(|reference| **reference == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if occurrences != 1
            || field.path.as_slice() != expected_path
            || !is_message_reference_field(field)
            || selected.replace((field_info_index, field)).is_some()
        {
            return Err(invalid_source(path));
        }
    }
    let mut fields = Vec::new();
    if let Some((field_info_index, field)) = selected {
        fields
            .try_reserve_exact(1)
            .map_err(|_error| Error::Allocation { amount: 1, path })?;
        let before = try_copy_slice(&field.object_references, path)?;
        let mut after = Vec::new();
        after
            .try_reserve_exact(before.len().saturating_sub(1))
            .map_err(|_error| Error::Allocation {
                amount: before.len().saturating_sub(1),
                path,
            })?;
        after.extend(
            before
                .iter()
                .copied()
                .filter(|reference| *reference != identifier),
        );
        fields.push(FieldTransitionOwned {
            field_info_index,
            path: try_copy_slice(field.path.as_slice(), path)?,
            before,
            after,
        });
    }
    Ok(ReferenceTransitionOwned {
        aggregate_before,
        aggregate_after,
        fields,
    })
}

fn reference_transition_replace(
    object: &ArchiveObject,
    message_index: usize,
    source_identifier: u64,
    destination_identifier: u64,
    path: Path,
) -> Result<Option<ReferenceTransitionOwned>, Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or_else(|| invalid_source(path))?;
    let source_count = info
        .object_references
        .iter()
        .filter(|reference| **reference == source_identifier)
        .count();
    let mut source_field = None;
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        let occurrences = field
            .object_references
            .iter()
            .filter(|reference| **reference == source_identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if occurrences != 1
            || field.path.as_slice() != [1, 2]
            || !is_message_reference_field(field)
            || field.object_references.last() != Some(&source_identifier)
            || source_field.replace((field_info_index, field)).is_some()
        {
            return Err(invalid_source(path));
        }
    }
    let destination_in_metadata = info.object_references.contains(&destination_identifier)
        || info
            .field_infos
            .iter()
            .any(|field| field.object_references.contains(&destination_identifier));

    if source_count == 0 {
        if source_field.is_some() || destination_in_metadata {
            // A partially-described parent edge is ambiguous: changing the
            // payload while retaining these metadata references could leave
            // a stale or duplicate ownership edge.
            return Err(invalid_source(path));
        }
        return Ok(None);
    }
    if source_count != 1 || destination_in_metadata {
        return Err(invalid_source(path));
    }
    let aggregate_before = try_copy_slice(&info.object_references, path)?;
    // The exact header-preserving transition contract keeps retained
    // references in their source order and appends newly introduced
    // references as a suffix.  A replacement is therefore represented as a
    // removal followed by an append, rather than an in-place substitution;
    // this also makes the inverse transition restore the original list.
    let mut aggregate_after = Vec::new();
    aggregate_after
        .try_reserve_exact(aggregate_before.len())
        .map_err(|_error| Error::Allocation {
            amount: aggregate_before.len(),
            path,
        })?;
    aggregate_after.extend(
        aggregate_before
            .iter()
            .copied()
            .filter(|reference| *reference != source_identifier),
    );
    aggregate_after.push(destination_identifier);
    let mut fields = Vec::new();
    if let Some((field_info_index, field)) = source_field {
        fields
            .try_reserve_exact(1)
            .map_err(|_error| Error::Allocation { amount: 1, path })?;
        let before = try_copy_slice(&field.object_references, path)?;
        let mut after = Vec::new();
        after
            .try_reserve_exact(before.len())
            .map_err(|_error| Error::Allocation {
                amount: before.len(),
                path,
            })?;
        for reference in &before {
            after.push(if *reference == source_identifier {
                destination_identifier
            } else {
                *reference
            });
        }
        fields.push(FieldTransitionOwned {
            field_info_index,
            path: try_copy_slice(field.path.as_slice(), path)?,
            before,
            after,
        });
    }
    Ok(Some(ReferenceTransitionOwned {
        aggregate_before,
        aggregate_after,
        fields,
    }))
}

fn sheet_drawable_reference_path(message_type: u32, path: Path) -> Result<&'static [u32], Error> {
    match message_type {
        SHEET_MESSAGE_TYPE => Ok(&[2]),
        FORM_BASED_SHEET_MESSAGE_TYPE => Ok(&[1, 2]),
        _ => Err(invalid_source(path)),
    }
}

fn is_message_reference_field(field: &litchi_iwa_core::FieldInfo) -> bool {
    field
        .r#type
        .is_none_or(|field_type| field_type == litchi_iwa_core::FieldType::Message)
}

fn try_copy_slice<T: Copy>(source: &[T], path: Path) -> Result<Vec<T>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_error| Error::Allocation {
            amount: source.len(),
            path,
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

fn replace_with_transition(
    object: &mut ArchiveObject,
    message_index: usize,
    data: Vec<u8>,
    owned: ReferenceTransitionOwned,
    limits: litchi_iwa_archive::Limits,
    path: Path,
) -> Result<(), Error> {
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(owned.fields.len())
        .map_err(|_error| Error::Allocation {
            amount: owned.fields.len(),
            path,
        })?;
    for field in &owned.fields {
        fields.push(FieldObjectReferenceTransition {
            field_info_index: field.field_info_index,
            expected_path: field.path.as_slice(),
            before: field.before.as_slice(),
            after: field.after.as_slice(),
        });
    }
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: object
                    .messages
                    .get(message_index)
                    .ok_or_else(|| invalid_source(path))?
                    .type_,
                data,
            },
            ObjectReferenceTransition {
                aggregate_before: owned.aggregate_before.as_slice(),
                aggregate_after: owned.aggregate_after.as_slice(),
                fields: fields.as_slice(),
            },
            limits
                .effective_archive_limits()
                .map_err(|_error| invalid_source(path))?,
        )
        .map_err(|error| map_core_error(error, path))?;
    Ok(())
}

pub(super) fn verify_source_state(source: &Package, operation: Operation) -> Result<(), Error> {
    let target = resolve_table_target(
        source,
        operation.source_sheet,
        operation.source_table,
        operation.path(),
    )?;
    if target.sheet_identifier != operation.source_sheet_identifier
        || target.drawable_identifier != operation.drawable_identifier
        || target.model_identifier != operation.model_identifier
        || table_parent_identifier_for_target(source, target, operation.path())?
            != Some(operation.source_sheet_identifier)
    {
        return Err(Error::PatchConflict);
    }
    let destination = resolve_sheet_target(source, operation.destination_sheet, operation.path())?;
    if destination.identifier != operation.destination_sheet_identifier {
        return Err(Error::PatchConflict);
    }
    if operation.is_same_sheet() {
        return Ok(());
    }
    if sheet_reference_count(
        source,
        destination,
        operation.drawable_identifier,
        operation.path(),
    )? != 0
    {
        return Err(Error::PatchConflict);
    }
    Ok(())
}

/// Verify the rooted physical source state used by the migration-host seam.
///
/// Unlike [`verify_source_state`], this admission does not consult the
/// semantic document. A source-built host snapshot can retain a valid rooted
/// table graph while its table storage is outside the focused semantic
/// projection. The optional TableInfo parent is preserved when absent; a
/// present parent must still identify the source sheet.
#[cfg(feature = "internal-iwork-source")]
pub(super) fn verify_physical_source_state(
    source: &Package,
    operation: Operation,
) -> Result<Option<u64>, Error> {
    let target = resolve_table_target(
        source,
        operation.source_sheet,
        operation.source_table,
        operation.path(),
    )?;
    let parent = table_parent_identifier_for_target(source, target, operation.path())?;
    if target.sheet_identifier != operation.source_sheet_identifier
        || target.drawable_identifier != operation.drawable_identifier
        || target.model_identifier != operation.model_identifier
        || parent.is_some_and(|identifier| identifier != operation.source_sheet_identifier)
    {
        return Err(Error::PatchConflict);
    }
    let destination = resolve_sheet_target(source, operation.destination_sheet, operation.path())?;
    if destination.identifier != operation.destination_sheet_identifier {
        return Err(Error::PatchConflict);
    }
    if operation.is_same_sheet() {
        return Ok(parent);
    }
    if sheet_reference_count(
        source,
        destination,
        operation.drawable_identifier,
        operation.path(),
    )? != 0
    {
        return Err(Error::PatchConflict);
    }
    Ok(parent)
}

pub(super) fn verify_target_state(candidate: &Package, operation: Operation) -> Result<(), Error> {
    let target = resolve_table_target(
        candidate,
        operation.destination_sheet,
        operation.destination_table,
        operation.path(),
    )?;
    if target.sheet_identifier != operation.destination_sheet_identifier
        || target.drawable_identifier != operation.drawable_identifier
        || target.model_identifier != operation.model_identifier
        || table_parent_identifier_for_target(candidate, target, operation.path())?
            != Some(operation.destination_sheet_identifier)
    {
        return Err(Error::Verification);
    }
    let source = resolve_sheet_target(candidate, operation.source_sheet, operation.path())?;
    let destination =
        resolve_sheet_target(candidate, operation.destination_sheet, operation.path())?;
    if sheet_reference_count(
        candidate,
        source,
        operation.drawable_identifier,
        operation.path(),
    )? != 0
        || sheet_reference_count(
            candidate,
            destination,
            operation.drawable_identifier,
            operation.path(),
        )? != 1
    {
        return Err(Error::Verification);
    }
    Ok(())
}

/// Verify the rooted physical target state used by the migration-host seam.
///
/// The target must retain the source's parent-edge shape: a present parent is
/// rewritten to the destination sheet, while an omitted parent remains
/// omitted. Sheet ownership and cardinality are checked in both cases.
#[cfg(feature = "internal-iwork-source")]
pub(super) fn verify_physical_target_state(
    candidate: &Package,
    operation: Operation,
    source_parent: Option<u64>,
) -> Result<(), Error> {
    let target = resolve_table_target(
        candidate,
        operation.destination_sheet,
        operation.destination_table,
        operation.path(),
    )?;
    let expected_parent = source_parent.map(|_identifier| operation.destination_sheet_identifier);
    if target.sheet_identifier != operation.destination_sheet_identifier
        || target.drawable_identifier != operation.drawable_identifier
        || target.model_identifier != operation.model_identifier
        || table_parent_identifier_for_target(candidate, target, operation.path())?
            != expected_parent
    {
        return Err(Error::Verification);
    }
    let source = resolve_sheet_target(candidate, operation.source_sheet, operation.path())?;
    let destination =
        resolve_sheet_target(candidate, operation.destination_sheet, operation.path())?;
    if sheet_reference_count(
        candidate,
        source,
        operation.drawable_identifier,
        operation.path(),
    )? != 0
        || sheet_reference_count(
            candidate,
            destination,
            operation.drawable_identifier,
            operation.path(),
        )? != 1
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn table_parent_identifier_for_target(
    source: &Package,
    target: table_headers::Target,
    path: Path,
) -> Result<Option<u64>, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(target.info_component_index)
        .and_then(|component| component.archive().objects.get(target.info_object_index))
        .ok_or_else(|| invalid_source(path))?;
    let message = object
        .messages
        .get(target.info_message_index)
        .ok_or_else(|| invalid_source(path))?;
    table_parent_identifier(&message.data, path)
}

fn sheet_reference_count(
    source: &Package,
    sheet: SheetTarget,
    drawable_identifier: u64,
    path: Path,
) -> Result<usize, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(sheet.component_index)
        .and_then(|component| component.archive().objects.get(sheet.object_index))
        .ok_or_else(|| invalid_source(path))?;
    let message = object
        .messages
        .get(sheet.message_index)
        .ok_or_else(|| invalid_source(path))?;
    let payloads = table_headers::resolve::sheet_drawable_payloads(message.type_, &message.data)
        .map_err(|error| map_header_error(error, path))?;
    payloads.iter().try_fold(0usize, |count, payload| {
        let identifier = table_headers::resolve::local_reference_identifier(payload)
            .map_err(|error| map_header_error(error, path))?;
        count
            .checked_add(usize::from(identifier == drawable_identifier))
            .ok_or_else(|| invalid_source(path))
    })
}

pub(super) fn verify_locality(
    source: &Package,
    candidate: &Package,
    operation: Operation,
) -> Result<(), Error> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    let table = resolve_table_target(
        source,
        operation.source_sheet,
        operation.source_table,
        operation.path(),
    )?;
    let mut touched = [
        table.sheet_component_index,
        table.info_component_index,
        resolve_sheet_target(source, operation.source_sheet, operation.path())?.component_index,
        resolve_sheet_target(source, operation.destination_sheet, operation.path())?
            .component_index,
    ];
    touched.sort_unstable();
    root_preview_deletions(source_catalog)?;
    root_preview_deletions(candidate_catalog)?;

    let mut before_entries = source_catalog
        .package()
        .iter()
        .filter(|entry| !ROOT_PREVIEWS.contains(&entry.name()));
    let mut after_entries = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !ROOT_PREVIEWS.contains(&entry.name()));
    loop {
        let before = before_entries.next();
        let after = after_entries.next();
        let (Some(before), Some(after)) = (before, after) else {
            if before.is_some() || after.is_some() {
                return Err(Error::Verification);
            }
            break;
        };
        if before.name() != after.name() {
            return Err(Error::Verification);
        }
        if touched.iter().any(|index| {
            source
                .state
                .components
                .catalog()
                .get_index(*index)
                .is_some_and(|component| component.name() == before.name())
        }) {
            continue;
        }
        if !table_headers::rewrite::package_member_preserved(before, after) {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn touched_component_count(
    table: table_headers::Target,
    source_sheet: SheetTarget,
    destination_sheet: SheetTarget,
) -> usize {
    let mut components = [
        table.sheet_component_index,
        table.info_component_index,
        source_sheet.component_index,
        destination_sheet.component_index,
    ];
    components.sort_unstable();
    1 + components
        .windows(2)
        .filter(|pair| pair[0] != pair[1])
        .count()
}

pub(super) fn root_preview_deletions(source: &SourceCatalog) -> Result<Vec<&'static str>, Error> {
    let mut previews = Vec::new();
    previews
        .try_reserve_exact(ROOT_PREVIEWS.len())
        .map_err(|_error| Error::Allocation {
            amount: ROOT_PREVIEWS.len(),
            path: Path::Package,
        })?;
    for name in ROOT_PREVIEWS {
        let count = source
            .package()
            .iter()
            .filter(|entry| entry.name() == name)
            .count();
        if count > 1 {
            return Err(invalid_source(Path::Package));
        }
        if count == 1 {
            previews.push(name);
        }
    }
    Ok(previews)
}

pub(super) fn physical_source(source: &Package) -> Result<&SourceCatalog, Error> {
    source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)
}

fn invalid_source(path: Path) -> Error {
    Error::InvalidSource { path }
}

pub(super) fn map_header_error(error: table_headers::Error, path: Path) -> Error {
    match error {
        table_headers::Error::SheetNotFound => Error::SheetNotFound,
        table_headers::Error::TableNotFound => Error::TableNotFound,
        table_headers::Error::TableLocked { .. } => Error::TableLocked { path },
        table_headers::Error::UnsupportedSource => Error::UnsupportedSource,
        table_headers::Error::LimitExceeded {
            observed, maximum, ..
        } => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed,
            maximum,
            path,
        },
        table_headers::Error::Allocation { amount, .. } => Error::Allocation { amount, path },
        _ => invalid_source(path),
    }
}

fn map_wire_error(error: litchi_iwa_common::Error, path: Path) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            observed, limit, ..
        } => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(limit).unwrap_or(u64::MAX),
            path,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount, path },
        _ => invalid_source(path),
    }
}

fn map_core_error(error: CoreError, path: Path) -> Error {
    match error {
        CoreError::Limit {
            observed, maximum, ..
        } => Error::LimitExceeded {
            kind: LimitKind::PayloadBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path,
        },
        CoreError::Allocation { requested, .. } => Error::Allocation {
            amount: requested,
            path,
        },
        _ => invalid_source(path),
    }
}
