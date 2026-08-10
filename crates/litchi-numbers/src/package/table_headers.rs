//! Exact-source transactions for one rooted table's header and footer settings.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "the focused boundary exhaustively redacts lower-layer error families"
)]

use std::{collections::HashSet, fmt, sync::Arc};

use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{
        NestedFieldEdit, NestedFieldReplacement, WireDescent, WireView,
        patch_nested_fields_batched_with_limits, preflight_wire_tree_with_limits,
    },
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{numbers_table_header_settings_codec, table_info_codec};
use thiserror::Error as ThisError;

use super::{
    Error as ReadError, FORM_BASED_SHEET_MESSAGE_TYPE, LEGACY_TABLE_INFO_MESSAGE_TYPE, Package,
    Resolved, SHEET_MESSAGE_TYPE, TABLE_INFO_MESSAGE_TYPE, TABLE_MODEL_MESSAGE_TYPE,
    table_info_decode_options,
};
use crate::{
    selector::{SheetSelector, TableSelector},
    table::{headers::Settings, lock::State as LockState},
};

const LEGACY_TABLE_MODEL_MESSAGE_TYPE: u32 = 6_000;
const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const HEADER_ROWS_FIELD: u32 = 9;
const HEADER_COLUMNS_FIELD: u32 = 10;
const FOOTER_ROWS_FIELD: u32 = 11;
const HEADER_ROWS_FROZEN_FIELD: u32 = 12;
const HEADER_COLUMNS_FROZEN_FIELD: u32 = 13;
const REPEATING_HEADER_ROWS_FIELD: u32 = 29;
const REPEATING_HEADER_COLUMNS_FIELD: u32 = 32;
const MIN_SIGN_EXTENDED_I32: u64 = u64::MAX - 2_147_483_647;
const CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE: u32 = 6_372;
const GROUP_BY_MESSAGE_TYPE: u32 = 6_373;

/// A content-free semantic location associated with a transaction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One rooted table at checked zero-based positions.
    Table {
        /// Zero-based rooted sheet position.
        sheet: usize,
        /// Zero-based table position within the sheet.
        table: usize,
    },
}

/// A finite resource governed by the focused transaction.
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
    /// Aggregate member bytes.
    TotalEntryBytes,
    /// Physical container names and metadata.
    PackageBytes,
    /// Bytes in one decoded payload container.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native framing or metadata items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Bytes inspected by one wire projection.
    WireBytes,
    /// Bytes produced by one wire rewrite.
    WireOutputBytes,
    /// Protobuf fields inspected.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Work performed by one strict codec operation.
    WireWork,
    /// Aggregate transaction topology and codec work.
    TransactionWork,
}

/// Content-free validation reason for requested settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum InvalidReason {
    /// Leading and trailing row sections exceed the table's row count.
    RowSectionsExceedTable {
        /// Requested leading header rows.
        header_rows: u8,
        /// Requested trailing footer rows.
        footer_rows: u8,
        /// Rows declared by the selected table.
        table_rows: u32,
    },
    /// Leading header columns exceed the table's column count.
    HeaderColumnsExceedTable {
        /// Requested leading header columns.
        header_columns: u8,
        /// Columns declared by the selected table.
        table_columns: u32,
    },
}

impl fmt::Display for InvalidReason {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::RowSectionsExceedTable {
                header_rows,
                footer_rows,
                table_rows,
            } => write!(
                formatter,
                "header rows {header_rows} plus footer rows {footer_rows} exceed {table_rows} table rows"
            ),
            Self::HeaderColumnsExceedTable {
                header_columns,
                table_columns,
            } => write!(
                formatter,
                "header columns {header_columns} exceed {table_columns} table columns"
            ),
        }
    }
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
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
            Self::TransactionWork => "transaction work",
        })
    }
}

/// A content-redacted table header/footer transaction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No rooted sheet matched the selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No table on the selected sheet matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// Requested settings exceed the selected table's dimensions.
    #[error("the requested table header and footer settings are invalid at {path:?}: {reason}")]
    InvalidSettings {
        /// Semantic location of the rejected settings.
        path: Path,
        /// Content-free validation reason.
        reason: InvalidReason,
    },
    /// A changed edit targeted an effectively locked table.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked {
        /// Selected semantic table.
        path: Path,
    },
    /// A changed edit would stale a rooted dependent topology.
    #[error("the selected Numbers table has a dependent header topology at {path:?}")]
    UnsupportedDependency {
        /// Selected semantic table.
        path: Path,
    },
    /// The source lacks exact physical provenance for changed publication.
    #[error("this Numbers source does not support exact table header editing")]
    UnsupportedSource,
    /// Rooted ownership, metadata, or wire framing is invalid.
    #[error("the Numbers table-header source is invalid at {path:?}")]
    InvalidSource {
        /// Content-free semantic failure location.
        path: Path,
    },
    /// A finite transaction resource ceiling was exceeded.
    #[error("Numbers table-header {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
        /// Content-free semantic failure location.
        path: Path,
    },
    /// A bounded allocation failed before publication.
    #[error("could not allocate {amount} units for the Numbers table-header transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
        /// Content-free semantic failure location.
        path: Path,
    },
    /// Candidate reopening or locality verification failed.
    #[error("the edited Numbers table headers failed semantic verification")]
    Verification,
    /// The patch was applied to a package other than its exact source.
    #[error("the table-header patch does not match the exact source package")]
    PatchConflict,
}

/// Mutable settings staged against one immutable package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    sheet_position: usize,
    table_position: usize,
    before: Settings,
    settings: Settings,
    target: Target,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path())
            .field("settings", &self.settings)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the selected semantic table path.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            sheet: self.sheet_position,
            table: self.table_position,
        }
    }

    /// Return the settings that would be published.
    #[must_use]
    pub const fn settings(&self) -> Settings {
        self.settings
    }

    /// Replace the staged settings.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.settings = settings;
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// # Errors
    ///
    /// Returns a typed selection, source, lock, settings, resource, or
    /// verification error. Failure never publishes a partial package.
    ///
    /// # Costs
    ///
    /// An exact no-op shares the source snapshot. A change parses and rewrites
    /// one component, reassembles once, and fully reopens one candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let catalog = physical_source(self.source)?;
        let source = catalog.shared_source();
        if self.before == self.settings {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source), source),
                    sheet_position: self.sheet_position,
                    table_position: self.table_position,
                    before: self.before,
                    after: self.settings,
                    target: self.target,
                    source_payload: None,
                    target_payload: None,
                    touched_components: 0,
                    source_previews: 0,
                    target_previews: 0,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::UnsupportedSource);
        }
        preflight_transaction_work(self.source, None)?;
        let path = self.path();
        let target = self.target;
        if target.settings != self.before {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        validate_requested(self.settings, target.rows, target.columns, path)?;
        validate_selected_ownership(self.source, target)?;
        if target.locked == LockState::Locked {
            return Err(Error::TableLocked { path });
        }
        validate_dependencies(self.source, target, self.before, self.settings)?;
        let previews = root_preview_deletions(catalog)?;
        let source_payload = clone_selected_payload(self.source, target)?;
        let (package, target_payload) = rewrite(self.source, target, self.settings, &previews)?;
        verify_exact_locality(self.source, &package, target, &previews, 0, &target_payload)?;
        let target_bytes = physical_source(&package)?.shared_source();
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(source, Arc::clone(&target_bytes)),
                sheet_position: self.sheet_position,
                table_position: self.table_position,
                before: self.before,
                after: self.settings,
                target,
                source_payload: Some(source_payload),
                target_payload: Some(Arc::clone(&target_payload)),
                touched_components: 1,
                source_previews: previews.len(),
                target_previews: 0,
            },
            diagnostics: Diagnostics::published(previews.len()),
        })
    }
}

/// A reversible, process-local exact-source patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    sheet_position: usize,
    table_position: usize,
    before: Settings,
    after: Settings,
    target: Target,
    source_payload: Option<Arc<[u8]>>,
    target_payload: Option<Arc<[u8]>>,
    touched_components: usize,
    source_previews: usize,
    target_previews: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected semantic table path.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            sheet: self.sheet_position,
            table: self.table_position,
        }
    }
    /// Return the exact settings required in the source artifact.
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.before
    }
    /// Return the exact settings retained in the target artifact.
    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }
    /// Return the source artifact's diagnostic fingerprint.
    ///
    /// This value is not collision-resistant and never authorizes patch
    /// application; exact retained bytes provide authorization.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target artifact's diagnostic fingerprint.
    ///
    /// This value is not collision-resistant and is not artifact identity.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
    /// Return whether settings and retained artifacts are exact no-ops.
    ///
    /// # Costs
    ///
    /// Uses allocation identity first and otherwise compares complete package
    /// artifacts in `O(package bytes)` time.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }
    /// Return the exact target-to-source inverse.
    ///
    /// # Costs
    ///
    /// Swaps shared artifacts and compact metadata in `O(1)` time.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet_position: self.sheet_position,
            table_position: self.table_position,
            before: self.after,
            after: self.before,
            target: self.target,
            source_payload: self.target_payload.clone(),
            target_payload: self.source_payload.clone(),
            touched_components: self.touched_components,
            source_previews: self.target_previews,
            target_previews: self.source_previews,
        }
    }
}

/// Compact publication evidence.
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
    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews,
            full_reparse_performed: true,
        }
    }
    /// Return whether publication changed package bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    /// Return the number of rewritten payload components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    /// Return the number of root previews deleted in this direction.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
    /// Return whether the candidate was fully reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one transaction.
#[must_use = "a table-header commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully verified package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    /// Consume this result and return the package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    /// Borrow the reversible exact-source patch.
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Target {
    sheet_position: usize,
    table_position: usize,
    model_identifier: u64,
    sheet_identifier: u64,
    drawable_identifier: u64,
    drawable_position: usize,
    sheet_component_index: usize,
    sheet_object_index: usize,
    sheet_message_index: usize,
    sheet_message_type: u32,
    info_component_index: usize,
    info_object_index: usize,
    info_message_index: usize,
    info_message_type: u32,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
    settings: Settings,
    rows: u32,
    columns: u32,
    locked: LockState,
}

impl Package {
    /// Read one rooted table's lossless header and footer settings.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, allocation, or resource error.
    ///
    /// # Costs
    ///
    /// Uses the retained semantic and native indexes and strictly scans the
    /// selected model payload once.
    pub fn table_header_settings<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Settings, Error> {
        let (sheet_position, table_position) = self.resolve_header_selectors(sheet, table)?;
        Ok(resolve_target(self, sheet_position, table_position)?.settings)
    }

    /// Start a selector-first immutable header and footer settings edit.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, allocation, or resource error.
    ///
    /// # Costs
    ///
    /// Borrows this package and strictly scans the selected model once; it
    /// does not copy package bytes.
    pub fn edit_table_headers<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet_position, table_position) = self.resolve_header_selectors(sheet, table)?;
        let target = resolve_target(self, sheet_position, table_position)?;
        let before = target.settings;
        Ok(Edit {
            source: self,
            sheet_position,
            table_position,
            before,
            settings: before,
            target,
        })
    }

    /// Apply an exact-source-checked reversible header settings patch.
    ///
    /// # Errors
    ///
    /// Returns a conflict when this is not the retained exact source. A valid
    /// changed target must reopen and reproduce its requested semantic state.
    ///
    /// # Costs
    ///
    /// A no-op shares this snapshot. A changed patch fully reopens one retained
    /// target artifact and verifies its semantic state and physical locality.
    pub fn apply_table_headers(&self, patch: &Patch) -> Result<Commit, Error> {
        let source_catalog = physical_source(self)?;
        let source = source_catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if settings_at_target(self, patch.target)? != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.source_payload.as_deref() != Some(selected_payload(self, patch.target)?) {
            return Err(Error::PatchConflict);
        }
        if !source_catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let target_bytes = patch.artifacts.target();
        preflight_transaction_work(self, Some(&target_bytes))?;
        let candidate = Package::from_shared_bytes_with_options(target_bytes, self.state.options)
            .map_err(map_candidate_read_error)?;
        if settings_at_target(&candidate, patch.target)? != patch.after {
            return Err(Error::Verification);
        }
        let source_previews = root_preview_deletions(source_catalog)?;
        if source_previews.len() != patch.source_previews {
            return Err(Error::PatchConflict);
        }
        verify_exact_locality(
            self,
            &candidate,
            patch.target,
            &source_previews,
            patch.target_previews,
            patch
                .target_payload
                .as_deref()
                .ok_or(Error::PatchConflict)?,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.source_previews.saturating_sub(patch.target_previews),
            ),
        })
    }

    fn resolve_header_selectors<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<(usize, usize), Error> {
        let selected_sheet = self
            .state
            .document
            .sheet(sheet)
            .map_err(|error| map_read_error(ReadError::Semantic(error)))?
            .ok_or(Error::SheetNotFound)?;
        let table_position = match table.into() {
            TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
            TableSelector::Name(name) => selected_sheet
                .tables()
                .position(|candidate_table| candidate_table.name() == name),
        }
        .ok_or(Error::TableNotFound)?;
        Ok((selected_sheet.index(), table_position))
    }
}

fn resolve_target(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
) -> Result<Target, Error> {
    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (_document_index, document_message) =
        unique_message_index(&document_object.messages, super::DOCUMENT_MESSAGE_TYPE)?.ok_or(
            Error::InvalidSource {
                path: Path::Package,
            },
        )?;
    let sheet_payloads = repeated_length_payloads(&document_message.data, 1)?;
    let sheet_identifier = local_reference_identifier(
        sheet_payloads
            .get(sheet_position)
            .ok_or(Error::SheetNotFound)?,
    )?;
    let sheet = source
        .state
        .index
        .resolve_ref_id(&source.state.components, sheet_identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let sheet_message_index = unique_sheet_message_index(sheet.messages)?;
    let sheet_message = sheet
        .messages
        .get(sheet_message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let drawable_payloads = sheet_drawable_payloads(sheet_message.type_, &sheet_message.data)?;
    let mut semantic_table = 0usize;
    for (drawable_position, drawable_payload) in drawable_payloads.iter().enumerate() {
        let drawable_identifier = local_reference_identifier(drawable_payload)?;
        let info = source
            .state
            .index
            .resolve_ref_id(&source.state.components, drawable_identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let Some((info_message_index, info_message)) = unique_table_info(info)? else {
            continue;
        };
        let info_snapshot = table_info_codec::decode_table_info(
            &info_message.data,
            table_info_decode_options(&info_message.data),
        )
        .map_err(map_table_info_codec_error)?;
        if semantic_table != table_position {
            semantic_table = semantic_table.checked_add(1).ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
            continue;
        }
        let model_identifier = info_snapshot.table_model().identifier().get();
        let model = source
            .state
            .index
            .resolve_ref_id(&source.state.components, model_identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let (message_index, message) = unique_table_model(model.messages)?;
        let decoded = decode_settings(&message.data)?;
        let settings = settings_from_snapshot(&decoded)?;
        validate_stored(settings, decoded.rows(), decoded.columns())?;
        if sheet_identifier == drawable_identifier
            || sheet_identifier == model_identifier
            || drawable_identifier == model_identifier
        {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        return Ok(Target {
            sheet_position,
            table_position,
            model_identifier,
            sheet_identifier,
            drawable_identifier,
            drawable_position,
            sheet_component_index: sheet.component_index,
            sheet_object_index: sheet.object_index,
            sheet_message_index,
            sheet_message_type: sheet_message.type_,
            info_component_index: info.component_index,
            info_object_index: info.object_index,
            info_message_index,
            info_message_type: info_message.type_,
            component_index: model.component_index,
            object_index: model.object_index,
            message_index,
            message_type: message.type_,
            settings,
            rows: decoded.rows(),
            columns: decoded.columns(),
            locked: LockState::from_locked(info_snapshot.locked().unwrap_or(false)),
        });
    }
    Err(Error::TableNotFound)
}

fn settings_at_target(source: &Package, target: Target) -> Result<Settings, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if object.archive_info.identifier != Some(target.model_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(object, target.message_index)?;
    let message = object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if message.type_ != target.message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let snapshot = decode_settings(&message.data)?;
    let settings = settings_from_snapshot(&snapshot)?;
    validate_stored(settings, snapshot.rows(), snapshot.columns())?;
    Ok(settings)
}

fn resolved_object<'a>(
    source: &'a Package,
    resolved: Resolved<'a>,
) -> Result<&'a litchi_iwa_core::ArchiveObject, Error> {
    source
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })
}

fn unique_sheet_message_index(messages: &[RawMessage]) -> Result<usize, Error> {
    let sheet = unique_message_index(messages, SHEET_MESSAGE_TYPE)?;
    let form = unique_message_index(messages, FORM_BASED_SHEET_MESSAGE_TYPE)?;
    match (sheet, form) {
        (Some(_), Some(_)) | (None, None) => Err(Error::InvalidSource {
            path: Path::Package,
        }),
        (Some((index, _)), None) | (None, Some((index, _))) => Ok(index),
    }
}

fn unique_table_info(resolved: Resolved<'_>) -> Result<Option<(usize, &RawMessage)>, Error> {
    let canonical = unique_message_index(resolved.messages, TABLE_INFO_MESSAGE_TYPE)?;
    let legacy = unique_message_index(resolved.messages, LEGACY_TABLE_INFO_MESSAGE_TYPE)?;
    match (canonical, legacy) {
        (Some(_), Some(_)) => Err(Error::InvalidSource {
            path: Path::Package,
        }),
        (Some(message), None) | (None, Some(message)) => Ok(Some(message)),
        (None, None) => Ok(None),
    }
}

fn unique_table_model(messages: &[RawMessage]) -> Result<(usize, &RawMessage), Error> {
    let canonical = unique_message_index(messages, TABLE_MODEL_MESSAGE_TYPE)?;
    let legacy = unique_message_index(messages, LEGACY_TABLE_MODEL_MESSAGE_TYPE)?;
    match (canonical, legacy) {
        (Some(_), Some(_)) | (None, None) => Err(Error::InvalidSource {
            path: Path::Package,
        }),
        (Some(message), None) | (None, Some(message)) => Ok(message),
    }
}

fn unique_message_index(
    messages: &[RawMessage],
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, Error> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == message_type);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(first)
}

fn decode_settings(
    source: &[u8],
) -> Result<numbers_table_header_settings_codec::TableHeaderSettingsSnapshot, Error> {
    let options = numbers_table_header_settings_codec::DecodeOptions::new(
        source.len().max(1),
        source.len().max(1),
        source.len().saturating_mul(4).max(1),
        u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
    );
    numbers_table_header_settings_codec::decode_table_header_settings(source, options)
        .map_err(map_header_codec_error)
}

fn settings_from_snapshot(
    snapshot: &numbers_table_header_settings_codec::TableHeaderSettingsSnapshot,
) -> Result<Settings, Error> {
    fn count(raw: Option<u32>) -> Result<Option<crate::table::headers::Count>, Error> {
        raw.map(|raw_count| {
            usize::try_from(raw_count)
                .ok()
                .and_then(|count| crate::table::headers::Count::new(count).ok())
                .ok_or(Error::InvalidSource {
                    path: Path::Package,
                })
        })
        .transpose()
    }
    Ok(Settings {
        header_rows: count(snapshot.header_rows())?,
        header_columns: count(snapshot.header_columns())?,
        footer_rows: count(snapshot.footer_rows())?,
        header_rows_frozen: snapshot.header_rows_frozen(),
        header_columns_frozen: snapshot.header_columns_frozen(),
        repeating_header_rows_enabled: snapshot.repeating_header_rows_enabled(),
        repeating_header_columns_enabled: snapshot.repeating_header_columns_enabled(),
    })
}

fn validate_requested(
    settings: Settings,
    rows: u32,
    columns: u32,
    path: Path,
) -> Result<(), Error> {
    let header_rows = u8::try_from(settings.header_row_count()).unwrap_or(u8::MAX);
    let footer_rows = u8::try_from(settings.footer_row_count()).unwrap_or(u8::MAX);
    if u16::from(header_rows).saturating_add(u16::from(footer_rows))
        > u16::try_from(rows).unwrap_or(u16::MAX)
    {
        return Err(Error::InvalidSettings {
            path,
            reason: InvalidReason::RowSectionsExceedTable {
                header_rows,
                footer_rows,
                table_rows: rows,
            },
        });
    }
    let header_columns = u8::try_from(settings.header_column_count()).unwrap_or(u8::MAX);
    if u32::from(header_columns) > columns {
        return Err(Error::InvalidSettings {
            path,
            reason: InvalidReason::HeaderColumnsExceedTable {
                header_columns,
                table_columns: columns,
            },
        });
    }
    Ok(())
}

fn validate_message_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
) -> Result<(), Error> {
    let info =
        object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if info.type_ != message.type_
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(())
}

fn require_declared_reference(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    identifier: u64,
    accepted_path: &[u32],
) -> Result<(), Error> {
    let info =
        object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let mut field_occurrence = false;
    for field in &info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if count != 0 {
            if count != 1 || field_occurrence || field.path.as_slice() != accepted_path {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            field_occurrence = true;
        }
    }
    Ok(())
}

fn repeated_length_payloads(source: &[u8], field_number: u32) -> Result<Vec<&[u8]>, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let count = view
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|_allocation| Error::Allocation {
            amount: count,
            path: Path::Package,
        })?;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        values.push(field.payload());
    }
    Ok(values)
}

fn singular_length_payload(source: &[u8], field_number: u32) -> Result<&[u8], Error> {
    let values = repeated_length_payloads(source, field_number)?;
    match values.as_slice() {
        [value] => Ok(*value),
        _ => Err(Error::InvalidSource {
            path: Path::Package,
        }),
    }
}

fn sheet_drawable_payloads(message_type: u32, source: &[u8]) -> Result<Vec<&[u8]>, Error> {
    match message_type {
        SHEET_MESSAGE_TYPE => repeated_length_payloads(source, 2),
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            repeated_length_payloads(singular_length_payload(source, 1)?, 2)
        },
        _ => Err(Error::InvalidSource {
            path: Path::Package,
        }),
    }
}

fn canonical_varint(source: &[u8]) -> Result<u64, Error> {
    let (value, length) =
        decode_varint_from_bytes(source).map_err(|_error| Error::InvalidSource {
            path: Path::Package,
        })?;
    if length != source.len() || encoded_len(value) != length {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(value)
}

fn require_local_reference(source: &[u8], expected: u64) -> Result<(), Error> {
    if local_reference_identifier(source)? != expected {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(())
}

fn local_reference_identifier(source: &[u8]) -> Result<u64, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 if identifier.is_none() && field.wire_type() == 0 => {
                identifier = Some(canonical_varint(field.payload())?);
            },
            2 if deprecated_type.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > u64::from(i32::MAX.unsigned_abs()) && value < MIN_SIGN_EXTENDED_I32 {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                deprecated_type = Some(value);
            },
            3 if external.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > 1 {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                external = Some(value != 0);
            },
            1..=3 => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
            _ => {},
        }
    }
    let resolved_identifier = identifier.ok_or(Error::InvalidSource {
        path: Path::Package,
    })?;
    if resolved_identifier == 0 || external == Some(true) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(resolved_identifier)
}

fn validate_stored(settings: Settings, rows: u32, columns: u32) -> Result<(), Error> {
    if u64::try_from(settings.header_row_count())
        .unwrap_or(u64::MAX)
        .saturating_add(u64::try_from(settings.footer_row_count()).unwrap_or(u64::MAX))
        > u64::from(rows)
        || u64::try_from(settings.header_column_count()).unwrap_or(u64::MAX) > u64::from(columns)
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(())
}

fn validate_selected_ownership(source: &Package, target: Target) -> Result<(), Error> {
    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (document_message_index, document_message) =
        unique_message_index(&document_object.messages, super::DOCUMENT_MESSAGE_TYPE)?.ok_or(
            Error::InvalidSource {
                path: Path::Package,
            },
        )?;
    validate_message_metadata(document_object, document_message_index)?;
    let mut wire_work = 0usize;
    let maximum_work = source.state.options.archive().max_iwa_stream_bytes();
    charge_work(
        &mut wire_work,
        document_message.data.len().saturating_mul(2),
        maximum_work,
    )?;
    let sheet_payloads = repeated_length_payloads(&document_message.data, 1)?;
    require_local_reference(
        sheet_payloads
            .get(target.sheet_position)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?,
        target.sheet_identifier,
    )?;
    require_declared_reference(
        document_object,
        document_message_index,
        target.sheet_identifier,
        &[1],
    )?;

    let sheet_object = source
        .state
        .components
        .catalog()
        .get_index(target.sheet_component_index)
        .and_then(|component| component.archive().objects.get(target.sheet_object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if sheet_object.archive_info.identifier != Some(target.sheet_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(sheet_object, target.sheet_message_index)?;
    let sheet_message = sheet_object
        .messages
        .get(target.sheet_message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if sheet_message.type_ != target.sheet_message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let drawable_payloads =
        sheet_drawable_payloads(target.sheet_message_type, &sheet_message.data)?;
    require_local_reference(
        drawable_payloads
            .get(target.drawable_position)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?,
        target.drawable_identifier,
    )?;
    let sheet_path: &[u32] = match target.sheet_message_type {
        SHEET_MESSAGE_TYPE => &[2],
        FORM_BASED_SHEET_MESSAGE_TYPE => &[1, 2],
        _ => {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        },
    };
    require_declared_reference(
        sheet_object,
        target.sheet_message_index,
        target.drawable_identifier,
        sheet_path,
    )?;

    let info_object = source
        .state
        .components
        .catalog()
        .get_index(target.info_component_index)
        .and_then(|component| component.archive().objects.get(target.info_object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if info_object.archive_info.identifier != Some(target.drawable_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(info_object, target.info_message_index)?;
    let info_message =
        info_object
            .messages
            .get(target.info_message_index)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    if info_message.type_ != target.info_message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    require_local_reference(
        singular_length_payload(&info_message.data, 2)?,
        target.model_identifier,
    )?;
    require_declared_reference(
        info_object,
        target.info_message_index,
        target.model_identifier,
        &[2],
    )?;
    let model_object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if model_object.archive_info.identifier != Some(target.model_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(model_object, target.message_index)?;

    let mut owners = 0usize;
    if target.sheet_identifier == 1
        || target.drawable_identifier == 1
        || target.model_identifier == 1
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    for sheet_payload in sheet_payloads {
        charge_work(&mut wire_work, sheet_payload.len(), maximum_work)?;
        let sheet_identifier = local_reference_identifier(sheet_payload)?;
        if sheet_identifier == 1
            || sheet_identifier == target.drawable_identifier
            || sheet_identifier == target.model_identifier
        {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        let sheet = source
            .state
            .index
            .resolve_ref_id(&source.state.components, sheet_identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let sheet_message_index = unique_sheet_message_index(sheet.messages)?;
        let owner_sheet_message =
            sheet
                .messages
                .get(sheet_message_index)
                .ok_or(Error::InvalidSource {
                    path: Path::Package,
                })?;
        let sheet_multiplier = if owner_sheet_message.type_ == FORM_BASED_SHEET_MESSAGE_TYPE {
            4
        } else {
            2
        };
        charge_work(
            &mut wire_work,
            owner_sheet_message
                .data
                .len()
                .saturating_mul(sheet_multiplier),
            maximum_work,
        )?;
        for drawable_payload in
            sheet_drawable_payloads(owner_sheet_message.type_, &owner_sheet_message.data)?
        {
            charge_work(&mut wire_work, drawable_payload.len(), maximum_work)?;
            let drawable_identifier = local_reference_identifier(drawable_payload)?;
            if drawable_identifier == target.sheet_identifier
                || drawable_identifier == target.model_identifier
            {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            let info = source
                .state
                .index
                .resolve_ref_id(&source.state.components, drawable_identifier)
                .map_err(map_read_error)?
                .ok_or(Error::InvalidSource {
                    path: Path::Package,
                })?;
            let Some((_index, message)) = unique_table_info(info)? else {
                continue;
            };
            charge_work(
                &mut wire_work,
                message.data.len().saturating_mul(4),
                maximum_work,
            )?;
            let snapshot = table_info_codec::decode_table_info(
                &message.data,
                table_info_decode_options(&message.data),
            )
            .map_err(map_table_info_codec_error)?;
            let model_identifier = snapshot.table_model().identifier().get();
            if model_identifier == target.sheet_identifier
                || model_identifier == target.drawable_identifier
            {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            if model_identifier == target.model_identifier {
                owners = owners.checked_add(1).ok_or(Error::InvalidSource {
                    path: Path::Package,
                })?;
                if owners > 1 {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
            }
        }
    }
    if owners != 1 {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(())
}

fn charge_work(total: &mut usize, amount: usize, maximum: usize) -> Result<(), Error> {
    *total = total.checked_add(amount).ok_or(Error::LimitExceeded {
        kind: LimitKind::TransactionWork,
        observed: u64::MAX,
        maximum: usize_as_u64(maximum),
        path: Path::Package,
    })?;
    if *total > maximum {
        return Err(Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: usize_as_u64(*total),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        });
    }
    Ok(())
}

fn validate_dependencies(
    source: &Package,
    target: Target,
    before: Settings,
    after: Settings,
) -> Result<(), Error> {
    let path = Path::Table {
        sheet: target.sheet_position,
        table: target.table_position,
    };
    let model_object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let model = model_object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let model_view = WireView::parse(&model.data).map_err(map_wire_error)?;
    let header_counts_changed =
        before.header_rows != after.header_rows || before.header_columns != after.header_columns;
    let section_counts_changed = header_counts_changed || before.footer_rows != after.footer_rows;
    let mut active_pivot_or_group = false;
    let mut header_count_dependency = false;
    let mut pivot_seen = false;
    let mut unsupported_pivot_dependency = false;
    let mut group_seen = false;
    let mut other_seen = [false; 3];
    for field in model_view.fields() {
        match field.number() {
            83 if !group_seen && field.wire_type() == 2 => {
                field.validate_canonical_framing().map_err(map_wire_error)?;
                group_seen = true;
                active_pivot_or_group |= !field.payload().is_empty();
                header_count_dependency |= !field.payload().is_empty();
            },
            85 if !pivot_seen && field.wire_type() == 2 => {
                field.validate_canonical_framing().map_err(map_wire_error)?;
                pivot_seen = true;
                let identifier = local_reference_identifier(field.payload())?;
                if identifier == 1
                    || identifier == target.sheet_identifier
                    || identifier == target.drawable_identifier
                    || identifier == target.model_identifier
                {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                require_declared_reference(model_object, target.message_index, identifier, &[85])?;
                unsupported_pivot_dependency = true;
                active_pivot_or_group = true;
                header_count_dependency = true;
            },
            81 | 84 | 86 if field.wire_type() == 2 => {
                let slot = match field.number() {
                    81 => 0,
                    84 => 1,
                    86 => 2,
                    _ => {
                        return Err(Error::InvalidSource {
                            path: Path::Package,
                        });
                    },
                };
                if other_seen[slot] {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                other_seen[slot] = true;
                field.validate_canonical_framing().map_err(map_wire_error)?;
                header_count_dependency = true;
                if field.number() == 81 {
                    active_pivot_or_group |= deprecated_category_grouping_active(field.payload())?;
                } else if field.number() == 86 {
                    active_pivot_or_group |=
                        category_owner_reference_active(source, target, field.payload())?;
                }
            },
            81 | 83 | 84 | 85 | 86 => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
            _ => {},
        }
    }
    if unsupported_pivot_dependency {
        return Err(Error::UnsupportedDependency { path });
    }
    if header_counts_changed && header_count_dependency {
        return Err(Error::UnsupportedDependency { path });
    }
    if section_counts_changed && active_pivot_or_group {
        return Err(Error::UnsupportedDependency { path });
    }
    if table_info_has_count_dependency(
        source,
        target,
        header_counts_changed,
        section_counts_changed,
    )? {
        return Err(Error::UnsupportedDependency { path });
    }
    let repeating_changed = before.repeating_header_rows_enabled
        != after.repeating_header_rows_enabled
        || before.repeating_header_columns_enabled != after.repeating_header_columns_enabled;
    if repeating_changed {
        let sheet = source
            .state
            .components
            .catalog()
            .get_index(target.sheet_component_index)
            .and_then(|component| component.archive().objects.get(target.sheet_object_index))
            .and_then(|object| object.messages.get(target.sheet_message_index))
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let payload = if target.sheet_message_type == FORM_BASED_SHEET_MESSAGE_TYPE {
            singular_length_payload(&sheet.data, 1)?
        } else {
            &sheet.data
        };
        if WireView::parse(payload)
            .map_err(map_wire_error)?
            .fields()
            .any(|field| field.number() == 4)
        {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    if header_counts_changed && has_rooted_header_name_manager(source)? {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok(())
}

fn table_info_has_count_dependency(
    source: &Package,
    target: Target,
    header_counts_changed: bool,
    section_counts_changed: bool,
) -> Result<bool, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(target.info_component_index)
        .and_then(|component| component.archive().objects.get(target.info_object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let message = object
        .messages
        .get(target.info_message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let view = WireView::parse(&message.data).map_err(map_wire_error)?;
    let mut seen = [false; 7];
    let mut header_active = false;
    let mut section_active = false;
    for field in view.fields() {
        let slot_option = match field.number() {
            4 => Some(0),
            5 => Some(1),
            7 => Some(2),
            8 => Some(3),
            15 => Some(4),
            16 => Some(5),
            17 => Some(6),
            _ => None,
        };
        let Some(slot_index) = slot_option else {
            continue;
        };
        if seen[slot_index] {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        seen[slot_index] = true;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            4 | 5 | 15 | 17 if field.wire_type() == 2 => {
                let identifier = local_reference_identifier(field.payload())?;
                if identifier == 1
                    || identifier == target.sheet_identifier
                    || identifier == target.drawable_identifier
                    || identifier == target.model_identifier
                {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                require_declared_reference(
                    object,
                    target.info_message_index,
                    identifier,
                    &[field.number()],
                )?;
                header_active = true;
                section_active |= matches!(field.number(), 5 | 15 | 17);
            },
            7 | 8 if field.wire_type() == 2 => header_active = true,
            16 if field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > 1 {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                header_active |= value == 1;
                section_active |= value == 1;
            },
            _ => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
        }
    }
    Ok((header_counts_changed && header_active) || (section_counts_changed && section_active))
}

fn deprecated_category_grouping_active(source: &[u8]) -> Result<bool, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let mut active = false;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == 2 {
            if field.wire_type() != 2 {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            active |= group_by_enabled(field.payload())?.ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        }
    }
    Ok(active)
}

fn group_by_enabled(source: &[u8]) -> Result<Option<bool>, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let mut enabled = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == 6 {
            if enabled.is_some() || field.wire_type() != 0 {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            let value = canonical_varint(field.payload())?;
            if value > 1 {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            enabled = Some(value == 1);
        }
    }
    Ok(enabled)
}

fn category_owner_reference_active(
    source: &Package,
    target: Target,
    payload: &[u8],
) -> Result<bool, Error> {
    let owner_identifier = local_reference_identifier(payload)?;
    if owner_identifier == 1
        || owner_identifier == target.sheet_identifier
        || owner_identifier == target.drawable_identifier
        || owner_identifier == target.model_identifier
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let model_object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    require_declared_reference(model_object, target.message_index, owner_identifier, &[86])?;
    let owner = source
        .state
        .index
        .resolve_ref_id(&source.state.components, owner_identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let owner_object = resolved_object(source, owner)?;
    let (owner_message_index, owner_message) =
        unique_message_index(owner.messages, CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE)?.ok_or(
            Error::InvalidSource {
                path: Path::Package,
            },
        )?;
    let references = repeated_length_payloads(&owner_message.data, 1)?;
    validate_message_metadata(owner_object, owner_message_index)?;
    let mut group_identifiers = HashSet::new();
    group_identifiers
        .try_reserve(references.len())
        .map_err(|_allocation| Error::Allocation {
            amount: references.len(),
            path: Path::Package,
        })?;
    let mut active = false;
    for reference in references {
        let identifier = local_reference_identifier(reference)?;
        if identifier == 1
            || identifier == target.sheet_identifier
            || identifier == target.drawable_identifier
            || identifier == target.model_identifier
            || identifier == owner_identifier
            || !group_identifiers.insert(identifier)
        {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        require_declared_reference(owner_object, owner_message_index, identifier, &[1])?;
        let group = source
            .state
            .index
            .resolve_ref_id(&source.state.components, identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let group_object = resolved_object(source, group)?;
        let (message_index, message) = unique_message_index(group.messages, GROUP_BY_MESSAGE_TYPE)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let enabled = group_by_enabled(&message.data)?.ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
        validate_message_metadata(group_object, message_index)?;
        active |= enabled;
    }
    Ok(active)
}

fn has_rooted_header_name_manager(source: &Package) -> Result<bool, Error> {
    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (document_index, document_message) =
        unique_message_index(&document_object.messages, super::DOCUMENT_MESSAGE_TYPE)?.ok_or(
            Error::InvalidSource {
                path: Path::Package,
            },
        )?;
    let legacy = repeated_length_payloads(&document_message.data, 3)?;
    let super_payload = singular_length_payload(&document_message.data, 8)?;
    let primary = repeated_length_payloads(super_payload, 4)?;
    let (engine_payload, engine_path): (&[u8], &[u32]) =
        match (primary.as_slice(), legacy.as_slice()) {
            ([], []) => return Ok(false),
            ([payload], []) => (*payload, &[8, 4]),
            ([], [payload]) => (*payload, &[3]),
            _ => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
        };
    let engine_identifier = local_reference_identifier(engine_payload)?;
    require_declared_reference(
        document_object,
        document_index,
        engine_identifier,
        engine_path,
    )?;
    let engine = source
        .state
        .index
        .resolve_ref_id(&source.state.components, engine_identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (engine_index, engine_message) =
        unique_message_index(engine.messages, 4_000)?.ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let engine_object = resolved_object(source, engine)?;
    validate_message_metadata(engine_object, engine_index)?;
    let manager_payloads = repeated_length_payloads(&engine_message.data, 14)?;
    let manager_payload = match manager_payloads.as_slice() {
        [] => return Ok(false),
        [payload] => *payload,
        _ => {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        },
    };
    let manager_identifier = local_reference_identifier(manager_payload)?;
    require_declared_reference(engine_object, engine_index, manager_identifier, &[14])?;
    let manager = source
        .state
        .index
        .resolve_ref_id(&source.state.components, manager_identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (manager_index, _message) =
        unique_message_index(manager.messages, 6_366)?.ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    validate_message_metadata(resolved_object(source, manager)?, manager_index)?;
    Ok(true)
}

fn rewritten_payload(source: &[u8], before: Settings, after: Settings) -> Result<Vec<u8>, Error> {
    if settings_from_snapshot(&decode_settings(source)?)? != before {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let paths = [
        [HEADER_ROWS_FIELD],
        [HEADER_COLUMNS_FIELD],
        [FOOTER_ROWS_FIELD],
        [HEADER_ROWS_FROZEN_FIELD],
        [HEADER_COLUMNS_FROZEN_FIELD],
        [REPEATING_HEADER_ROWS_FIELD],
        [REPEATING_HEADER_COLUMNS_FIELD],
    ];
    let before_values = [
        before
            .header_rows
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        before
            .header_columns
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        before
            .footer_rows
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        before.header_rows_frozen.map(u64::from),
        before.header_columns_frozen.map(u64::from),
        before.repeating_header_rows_enabled.map(u64::from),
        before.repeating_header_columns_enabled.map(u64::from),
    ];
    let after_values = [
        after
            .header_rows
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        after
            .header_columns
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        after
            .footer_rows
            .map(|value| u64::try_from(value.get()).unwrap_or(u64::MAX)),
        after.header_rows_frozen.map(u64::from),
        after.header_columns_frozen.map(u64::from),
        after.repeating_header_rows_enabled.map(u64::from),
        after.repeating_header_columns_enabled.map(u64::from),
    ];
    let edits = [
        NestedFieldEdit::new(
            &paths[0],
            before_values[0].is_some(),
            NestedFieldReplacement::Varint(after_values[0]),
        ),
        NestedFieldEdit::new(
            &paths[1],
            before_values[1].is_some(),
            NestedFieldReplacement::Varint(after_values[1]),
        ),
        NestedFieldEdit::new(
            &paths[2],
            before_values[2].is_some(),
            NestedFieldReplacement::Varint(after_values[2]),
        ),
        NestedFieldEdit::new(
            &paths[3],
            before_values[3].is_some(),
            NestedFieldReplacement::Varint(after_values[3]),
        ),
        NestedFieldEdit::new(
            &paths[4],
            before_values[4].is_some(),
            NestedFieldReplacement::Varint(after_values[4]),
        ),
        NestedFieldEdit::new(
            &paths[5],
            before_values[5].is_some(),
            NestedFieldReplacement::Varint(after_values[5]),
        ),
        NestedFieldEdit::new(
            &paths[6],
            before_values[6].is_some(),
            NestedFieldReplacement::Varint(after_values[6]),
        ),
    ];
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(map_wire_error)?
        .with_fields(source.len().max(1))
        .map_err(map_wire_error)?
        .with_nesting(WireLimits::MAX_NESTING)
        .map_err(map_wire_error)?
        .with_output_bytes(source.len().saturating_add(64).max(1))
        .map_err(map_wire_error)?
        .with_rewrite_work(source.len().saturating_mul(4).max(1))
        .map_err(map_wire_error)?;
    preflight_wire_tree_with_limits(source, limits, |_visit| Ok(WireDescent::Skip))
        .map_err(map_wire_error)?;
    let output =
        patch_nested_fields_batched_with_limits(source, &edits, limits).map_err(map_wire_error)?;
    if settings_from_snapshot(&decode_settings(&output)?)? != after {
        return Err(Error::Verification);
    }
    Ok(output)
}

fn rewrite(
    source: &Package,
    target: Target,
    after: Settings,
    previews: &[&str],
) -> Result<(Package, Arc<[u8]>), Error> {
    let source_catalog = physical_source(source)?;
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if entry.is_opaque() {
        return Err(Error::UnsupportedSource);
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
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    drop(stream);
    let object = archive
        .objects
        .get_mut(target.object_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if object.archive_info.identifier != Some(target.model_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(object, target.message_index)?;
    let message = object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if message.type_ != target.message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let data = rewritten_payload(&message.data, target.settings, after)?;
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(data.len())
        .map_err(|_allocation| Error::Allocation {
            amount: data.len(),
            path: Path::Table {
                sheet: target.sheet_position,
                table: target.table_position,
            },
        })?;
    retained.extend_from_slice(&data);
    let retained_payload: Arc<[u8]> = retained.into();
    object
        .replace_message_preserving_header_with_limits(
            target.message_index,
            RawMessage {
                type_: target.message_type,
                data,
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
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            previews,
            physical_limits,
        )
        .map_err(map_archive_error)?;
    drop(compressed);
    let candidate = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(map_candidate_read_error)?;
    let selected = resolve_target(&candidate, target.sheet_position, target.table_position)?;
    if selected.settings != after {
        return Err(Error::Verification);
    }
    Ok((candidate, retained_payload))
}

fn clone_selected_payload(source: &Package, target: Target) -> Result<Arc<[u8]>, Error> {
    let payload = selected_payload(source, target)?;
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(payload.len())
        .map_err(|_allocation| Error::Allocation {
            amount: payload.len(),
            path: Path::Table {
                sheet: target.sheet_position,
                table: target.table_position,
            },
        })?;
    retained.extend_from_slice(payload);
    Ok(retained.into())
}

fn selected_payload(source: &Package, target: Target) -> Result<&[u8], Error> {
    source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })
}

fn root_preview_deletions(source: &SourceCatalog) -> Result<Vec<&'static str>, Error> {
    let mut previews = Vec::new();
    previews
        .try_reserve_exact(ROOT_PREVIEWS.len())
        .map_err(|_allocation| Error::Allocation {
            amount: ROOT_PREVIEWS.len(),
            path: Path::Package,
        })?;
    for name in ROOT_PREVIEWS {
        match source
            .package()
            .iter()
            .filter(|entry| entry.name() == name)
            .count()
        {
            0 => {},
            1 => previews.push(name),
            _ => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
        }
    }
    Ok(previews)
}

fn verify_exact_locality(
    source: &Package,
    candidate: &Package,
    target: Target,
    source_previews: &[&str],
    expected_candidate_previews: usize,
    expected_payload: &[u8],
) -> Result<(), Error> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    let candidate_previews = root_preview_deletions(candidate_catalog)?;
    if candidate_previews.len() != expected_candidate_previews {
        return Err(Error::Verification);
    }
    let mut before_entries = source_catalog
        .package()
        .iter()
        .filter(|entry| !source_previews.contains(&entry.name()));
    let mut after_entries = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !candidate_previews.contains(&entry.name()));
    loop {
        match (before_entries.next(), after_entries.next()) {
            (Some(left), Some(right)) if left.name() == right.name() => {
                let selected = source
                    .state
                    .components
                    .catalog()
                    .get_index(target.component_index)
                    .is_some_and(|component| component.name() == left.name());
                let preserved = if selected {
                    selected_package_member_preserved(left, right)
                } else {
                    package_member_preserved(left, right)
                };
                if !preserved {
                    return Err(Error::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    let before_component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::Verification)?;
    let after_component = candidate
        .state
        .components
        .catalog()
        .get(before_component.name())
        .ok_or(Error::Verification)?;
    if before_component.archive().objects.len() != after_component.archive().objects.len() {
        return Err(Error::Verification);
    }
    for (index, (left, right)) in before_component
        .archive()
        .objects
        .iter()
        .zip(&after_component.archive().objects)
        .enumerate()
    {
        if index != target.object_index {
            if left.archive_info != right.archive_info || left.messages != right.messages {
                return Err(Error::Verification);
            }
            continue;
        }
        if left.messages.len() != right.messages.len()
            || left.archive_info.identifier != right.archive_info.identifier
            || left.archive_info.should_merge != right.archive_info.should_merge
            || left.archive_info.message_infos.len() != right.archive_info.message_infos.len()
        {
            return Err(Error::Verification);
        }
        for (message_index, (left_message, right_message)) in
            left.messages.iter().zip(&right.messages).enumerate()
        {
            if message_index == target.message_index {
                if left_message.type_ != right_message.type_
                    || expected_payload != right_message.data
                    || !message_info_preserved_except_length(
                        left.archive_info
                            .message_infos
                            .get(message_index)
                            .ok_or(Error::Verification)?,
                        right
                            .archive_info
                            .message_infos
                            .get(message_index)
                            .ok_or(Error::Verification)?,
                    )
                {
                    return Err(Error::Verification);
                }
            } else if left_message != right_message
                || left.archive_info.message_infos.get(message_index)
                    != right.archive_info.message_infos.get(message_index)
            {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..OFFSET.start] == candidate[..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

fn package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.metadata().local() == candidate.metadata().local()
        && source.metadata().central() == candidate.metadata().central()
        && selected_local_record_preserved(source, candidate)
        && selected_central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_local_record_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
    let left = source.raw_record().local_record();
    let right = candidate.raw_record().local_record();
    let (Some(left_header), Some(right_header)) = (
        zip_local_header_length(left),
        zip_local_header_length(right),
    ) else {
        return false;
    };
    if left_header != right_header
        || left[..CRC_AND_SIZES.start] != right[..CRC_AND_SIZES.start]
        || left[CRC_AND_SIZES.end..left_header] != right[CRC_AND_SIZES.end..right_header]
    {
        return false;
    }
    let Some(left_end) = left_header
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= left.len())
    else {
        return false;
    };
    let Some(right_end) = right_header
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= right.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &left[left_end..],
        &right[right_end..],
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

fn selected_local_suffix_preserved(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source == candidate;
    }
    let left_prefix = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let right_prefix = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    left_prefix == right_prefix
        && source.len() == candidate.len()
        && source.len() >= left_prefix + 12
        && source[..left_prefix] == candidate[..right_prefix]
        && source[left_prefix + 12..] == candidate[right_prefix + 12..]
}

fn selected_central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..CRC_AND_SIZES.start] == candidate[..CRC_AND_SIZES.start]
        && source[CRC_AND_SIZES.end..OFFSET.start] == candidate[CRC_AND_SIZES.end..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

fn message_info_preserved_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn physical_source(package: &Package) -> Result<&SourceCatalog, Error> {
    package
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)
}

fn preflight_transaction_work(
    source: &Package,
    retained_target: Option<&[u8]>,
) -> Result<(), Error> {
    let maximum =
        usize::try_from(source.state.options.archive().max_total_bytes()).unwrap_or(usize::MAX);
    let source_bytes = source.source_bytes();
    let mut observed = 0;
    charge_work(&mut observed, source_bytes.len().saturating_mul(2), maximum)?;
    if let Some(target_bytes) = retained_target
        && target_bytes != source_bytes
    {
        charge_work(&mut observed, target_bytes.len().saturating_mul(2), maximum)?;
    }
    for object in source.state.components.iter_objects() {
        for message in &object.messages {
            charge_work(
                &mut observed,
                message.data.len().saturating_mul(16),
                maximum,
            )?;
        }
    }
    Ok(())
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn map_header_codec_error(error: numbers_table_header_settings_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        };
    }
    match error.wire_resource_limit() {
        Some(numbers_table_header_settings_codec::WireResourceLimit::Bytes {
            observed,
            maximum,
        }) => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        },
        Some(numbers_table_header_settings_codec::WireResourceLimit::Nesting {
            observed,
            maximum,
        }) => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_table_info_codec_error(error: table_info_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        };
    }
    Error::InvalidSource {
        path: Path::Package,
    }
}

fn map_candidate_read_error(error: ReadError) -> Error {
    match error {
        ReadError::InputTooLarge { observed, maximum }
        | ReadError::Archive(litchi_iwa_archive::Error::Limit {
            kind: litchi_iwa_archive::LimitKind::InputBytes,
            observed,
            maximum,
        }) => Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed,
            maximum,
            path: Path::Package,
        },
        other => map_read_error(other),
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Common(common_error) => map_wire_error(common_error),
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::PayloadObjects,
                super::SemanticLimitKind::References => LimitKind::PayloadReferences,
                super::SemanticLimitKind::FormulaRenderDepth
                | super::SemanticLimitKind::FormulaDepth => LimitKind::WireNesting,
                super::SemanticLimitKind::FormulaRenderWork
                | super::SemanticLimitKind::FormulaWork => LimitKind::WireWork,
                super::SemanticLimitKind::OutputTextBytes
                | super::SemanticLimitKind::FormulaWireBytes
                | super::SemanticLimitKind::TextBytes => LimitKind::PayloadBytes,
                _ => LimitKind::PayloadItems,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        },
        ReadError::InputTooLarge { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed,
            maximum,
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
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
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Reassembly(_) => Error::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadItems,
                _ => LimitKind::PayloadBytes,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => Error::Allocation {
            amount: requested,
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::WireOutputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                _ => LimitKind::PayloadItems,
            },
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(limit),
            path: Path::Package,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}
