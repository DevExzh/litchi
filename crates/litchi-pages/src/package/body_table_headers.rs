//! Exact-source transactions for Pages body-table header and footer settings.
//!
//! The rooted body-table graph is resolved by [`table_lock`].  This adapter
//! owns the complete, presence-preserving header projection and the narrow
//! dependency proof required before changing row/column partitions.  Native
//! identifiers stay in the private proof and never cross the public API.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused package boundary keeps transaction vocabulary beside its proof"
)]

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireDescent, WireView, preflight_wire_tree_with_limits},
};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::table_header_settings_codec::{
    self, DecodeError, DecodeOptions, TableHeaderSettingsSnapshot, TableHeaderSettingsWrite,
    WireResourceLimit,
};
use thiserror::Error;

use super::{Package, PackageError, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::headers::Settings;

const TABLE_INFO_MESSAGE_TYPES: [u32; 2] = [6_000, 6_003];
const ROOT_PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const MAX_VARINT_BYTES: usize = 10;

/// Finite resources charged by one body-table header transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableHeaderSettingsLimitKind {
    /// Complete source package bytes.
    InputBytes,
    /// Complete candidate package bytes.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// One physical entry's bytes.
    EntryBytes,
    /// Aggregate physical entry bytes.
    TotalEntryBytes,
    /// Decoded payload bytes.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected.
    PayloadObjects,
    /// Native payload messages inspected.
    PayloadMessages,
    /// Native framing and metadata items inspected.
    PayloadItems,
    /// Native references inspected.
    PayloadReferences,
    /// Header wire bytes.
    WireBytes,
    /// Header output bytes.
    WireOutputBytes,
    /// Header fields.
    WireFields,
    /// Header nesting.
    WireNesting,
    /// Header work.
    WireWork,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for BodyTableHeaderSettingsLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
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
            Self::TransactionWork => "transaction work",
        })
    }
}

/// Content-free semantic reason for rejecting requested partitions.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableHeaderSettingsInvalidReason {
    /// Header rows plus footer rows exceed the table's rows.
    RowSectionsExceedTable {
        /// Requested header rows.
        header_rows: u8,
        /// Requested footer rows.
        footer_rows: u8,
        /// Native table rows.
        table_rows: u32,
    },
    /// Header columns exceed the table's columns.
    HeaderColumnsExceedTable {
        /// Requested header columns.
        header_columns: u8,
        /// Native table columns.
        table_columns: u32,
    },
}

impl fmt::Display for BodyTableHeaderSettingsInvalidReason {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::RowSectionsExceedTable {
                header_rows,
                footer_rows,
                table_rows,
            } => write!(
                f,
                "header rows {header_rows} plus footer rows {footer_rows} exceed {table_rows} table rows"
            ),
            Self::HeaderColumnsExceedTable {
                header_columns,
                table_columns,
            } => {
                write!(
                    f,
                    "header columns {header_columns} exceed {table_columns} table columns"
                )
            },
        }
    }
}

/// Failure from a Pages body-table header read or transaction.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableHeaderSettingsError {
    /// No rooted body table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// A name selector matched more than one rooted table.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The selector or rooted graph is ambiguous.
    #[error("the Pages body-table header selector is ambiguous")]
    AmbiguousSelector,
    /// Requested settings violate table dimensions.
    #[error("the requested Pages body-table header settings are invalid: {reason}")]
    InvalidSettings {
        /// Content-free reason.
        reason: BodyTableHeaderSettingsInvalidReason,
    },
    /// The source has no exact physical provenance.
    #[error("the Pages package source does not support exact body-table header editing")]
    UnsupportedSource,
    /// A rooted graph or selected payload is malformed.
    #[error("the selected Pages body-table header source is invalid")]
    InvalidSource,
    /// A supported dependency would become stale.
    #[error("the selected Pages body table has a dependent header topology")]
    UnsupportedDependency,
    /// The selected table is protected from editing.
    #[error("the selected Pages body table is locked")]
    TableLocked,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Pages body-table headers {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: BodyTableHeaderSettingsLimitKind,
        /// Observed amount.
        observed: u64,
        /// Maximum amount.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} units for Pages body-table headers")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Candidate reopening did not reproduce the requested settings.
    #[error("the edited Pages body-table headers failed semantic verification")]
    Verification,
    /// The patch was created from another exact source artifact.
    #[error("the Pages body-table header patch does not match the exact source package")]
    PatchConflict,
}

/// Mutable semantic header settings staged against one immutable package.
pub struct BodyTableHeaderSettingsEdit<'a> {
    source: &'a Package,
    target: table_lock::BodyTableTarget,
    before: Settings,
    settings: Settings,
}

impl fmt::Debug for BodyTableHeaderSettingsEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("BodyTableHeaderSettingsEdit")
            .field("before", &self.before)
            .field("settings", &self.settings)
            .finish_non_exhaustive()
    }
}

impl BodyTableHeaderSettingsEdit<'_> {
    /// Return the staged settings.
    #[must_use]
    pub const fn settings(&self) -> Settings {
        self.settings
    }

    /// Replace the complete staged settings.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.settings = settings;
        self
    }

    /// Validate and publish the staged settings atomically.
    pub fn commit(self) -> Result<BodyTableHeaderSettingsCommit, BodyTableHeaderSettingsError> {
        commit_edit(self)
    }
}

/// Exact-source reversible header settings patch.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyTableHeaderSettingsPatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: table_lock::BodyTableTarget,
    before: Settings,
    after: Settings,
    source_preview_count: usize,
    target_preview_count: usize,
}

impl fmt::Debug for BodyTableHeaderSettingsPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("BodyTableHeaderSettingsPatch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyTableHeaderSettingsPatch {
    /// Source settings required before applying this patch.
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.before
    }

    /// Settings produced by this patch.
    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }

    /// Source fingerprint used for diagnostics and conflict checks.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Target fingerprint used for diagnostics and conflict checks.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Whether this patch is an exact byte-and-semantic no-op.
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
            before: self.after,
            after: self.before,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
        }
    }
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct BodyTableHeaderSettingsDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl BodyTableHeaderSettingsDiagnostics {
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

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of removed root previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the candidate was fully reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one header transaction.
#[must_use = "a body-table header commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyTableHeaderSettingsCommit {
    package: Package,
    patch: BodyTableHeaderSettingsPatch,
    diagnostics: BodyTableHeaderSettingsDiagnostics,
}

impl BodyTableHeaderSettingsCommit {
    /// Borrow the validated package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume and return the validated package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyTableHeaderSettingsPatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableHeaderSettingsDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body's lossless header/footer settings.
    pub fn body_table_header_settings<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Settings, BodyTableHeaderSettingsError> {
        let target = resolve_target(self, selector.into())?;
        settings_at_target(self, &target)
    }

    /// Start a selector-first immutable header/footer edit.
    pub fn edit_body_table_header_settings<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<BodyTableHeaderSettingsEdit<'_>, BodyTableHeaderSettingsError> {
        let target = resolve_target(self, selector.into())?;
        let before = settings_at_target(self, &target)?;
        Ok(BodyTableHeaderSettingsEdit {
            source: self,
            target,
            before,
            settings: before,
        })
    }

    /// Apply an exact-source-checked reversible header patch.
    pub fn apply_body_table_header_settings(
        &self,
        patch: &BodyTableHeaderSettingsPatch,
    ) -> Result<BodyTableHeaderSettingsCommit, BodyTableHeaderSettingsError> {
        let mut budget =
            table_lock::WireBudget::new(self.state.source.limits()).map_err(map_lock_error)?;
        if fingerprint(self.source_bytes(), &mut budget)? != patch.source_fingerprint
            || !bytes_equal(self.source_bytes(), patch.source.as_ref(), &mut budget)?
        {
            return Err(BodyTableHeaderSettingsError::PatchConflict);
        }
        if settings_at_target_with_budget(self, &patch.proof, &mut budget)? != patch.before {
            return Err(BodyTableHeaderSettingsError::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source_preview_count != patch.target_preview_count {
                return Err(BodyTableHeaderSettingsError::PatchConflict);
            }
            return Ok(BodyTableHeaderSettingsCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyTableHeaderSettingsDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact()
            || preview_count(self, &mut budget)? != patch.source_preview_count
            || fingerprint(patch.target.as_ref(), &mut budget)? != patch.target_fingerprint
        {
            return Err(BodyTableHeaderSettingsError::PatchConflict);
        }
        let candidate = reopen_target(self, Arc::clone(&patch.target), &mut budget)?;
        if settings_at_target_with_budget(&candidate, &patch.proof, &mut budget)? != patch.after
            || preview_count(&candidate, &mut budget)? != patch.target_preview_count
        {
            return Err(BodyTableHeaderSettingsError::Verification);
        }
        Ok(BodyTableHeaderSettingsCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableHeaderSettingsDiagnostics::published(
                patch
                    .source_preview_count
                    .saturating_sub(patch.target_preview_count),
            ),
        })
    }
}

fn commit_edit(
    edit: BodyTableHeaderSettingsEdit<'_>,
) -> Result<BodyTableHeaderSettingsCommit, BodyTableHeaderSettingsError> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    let mut budget =
        table_lock::WireBudget::new(source.state.source.limits()).map_err(map_lock_error)?;
    let source_fingerprint = fingerprint(source_bytes.as_ref(), &mut budget)?;
    let source_preview_count = preview_count(source, &mut budget)?;
    if edit.before == edit.settings {
        return Ok(BodyTableHeaderSettingsCommit {
            package: source.snapshot(),
            patch: BodyTableHeaderSettingsPatch {
                source: Arc::clone(&source_bytes),
                target: source_bytes,
                source_fingerprint,
                target_fingerprint: source_fingerprint,
                proof: edit.target,
                before: edit.before,
                after: edit.settings,
                source_preview_count,
                target_preview_count: source_preview_count,
            },
            diagnostics: BodyTableHeaderSettingsDiagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(BodyTableHeaderSettingsError::UnsupportedSource);
    }
    let package = rewrite_headers(
        source,
        &edit.target,
        edit.before,
        edit.settings,
        &mut budget,
    )?;
    let target = package.state.source.shared_source();
    let target_fingerprint = fingerprint(target.as_ref(), &mut budget)?;
    let target_preview_count = preview_count(&package, &mut budget)?;
    Ok(BodyTableHeaderSettingsCommit {
        package,
        patch: BodyTableHeaderSettingsPatch {
            source: source_bytes,
            target,
            source_fingerprint,
            target_fingerprint,
            proof: edit.target,
            before: edit.before,
            after: edit.settings,
            source_preview_count,
            target_preview_count,
        },
        diagnostics: BodyTableHeaderSettingsDiagnostics::published(
            source_preview_count.saturating_sub(target_preview_count),
        ),
    })
}

fn resolve_target(
    package: &Package,
    selector: BodyTableSelector<'_>,
) -> Result<table_lock::BodyTableTarget, BodyTableHeaderSettingsError> {
    package.resolve_body_table(selector).map_err(map_lock_error)
}

fn settings_at_target(
    package: &Package,
    target: &table_lock::BodyTableTarget,
) -> Result<Settings, BodyTableHeaderSettingsError> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    settings_at_target_with_budget(package, target, &mut budget)
}

fn settings_at_target_with_budget(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<Settings, BodyTableHeaderSettingsError> {
    table_lock::validate_body_table_target(package, target, budget).map_err(map_lock_error)?;
    let message = model_message(package, target)?;
    let snapshot = decode_snapshot(&message.data, budget)?;
    settings_from_snapshot(snapshot)
}

fn model_message<'a>(
    package: &'a Package,
    target: &table_lock::BodyTableTarget,
) -> Result<&'a RawMessage, BodyTableHeaderSettingsError> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableHeaderSettingsError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    Ok(message)
}

fn decode_snapshot(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<TableHeaderSettingsSnapshot, BodyTableHeaderSettingsError> {
    let limits = budget.wire_limits();
    let preflight = preflight_wire_tree_with_limits(source, limits, |_| Ok(WireDescent::Skip))
        .map_err(map_common_wire_error)?;
    budget
        .charge_codec_report(
            preflight.fields(),
            preflight.scanned_bytes().saturating_mul(2),
            u32::try_from(preflight.max_depth()).unwrap_or(u32::MAX),
            0,
        )
        .map_err(map_lock_error)?;
    let options = DecodeOptions::new(
        source.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    );
    let snapshot = table_header_settings_codec::decode_table_header_settings(source, options)
        .map_err(map_codec_error)?;
    Ok(snapshot)
}

fn settings_from_snapshot(
    snapshot: TableHeaderSettingsSnapshot,
) -> Result<Settings, BodyTableHeaderSettingsError> {
    let count = |value: Option<u32>| {
        value
            .map(|value| {
                usize::try_from(value)
                    .ok()
                    .and_then(|value| crate::table::headers::Count::new(value).ok())
                    .ok_or(BodyTableHeaderSettingsError::InvalidSource)
            })
            .transpose()
    };
    let settings = Settings {
        header_rows: count(snapshot.header_rows())?,
        header_columns: count(snapshot.header_columns())?,
        footer_rows: count(snapshot.footer_rows())?,
        header_rows_frozen: snapshot.header_rows_frozen(),
        header_columns_frozen: snapshot.header_columns_frozen(),
        repeating_header_rows_enabled: snapshot.repeating_header_rows_enabled(),
        repeating_header_columns_enabled: snapshot.repeating_header_columns_enabled(),
    };
    validate_requested(settings, snapshot.rows(), snapshot.columns())?;
    Ok(settings)
}

fn validate_requested(
    settings: Settings,
    rows: u32,
    columns: u32,
) -> Result<(), BodyTableHeaderSettingsError> {
    let header_rows = u8::try_from(settings.header_row_count()).unwrap_or(u8::MAX);
    let footer_rows = u8::try_from(settings.footer_row_count()).unwrap_or(u8::MAX);
    if u16::from(header_rows).saturating_add(u16::from(footer_rows))
        > u16::try_from(rows).unwrap_or(u16::MAX)
    {
        return Err(BodyTableHeaderSettingsError::InvalidSettings {
            reason: BodyTableHeaderSettingsInvalidReason::RowSectionsExceedTable {
                header_rows,
                footer_rows,
                table_rows: rows,
            },
        });
    }
    let header_columns = u8::try_from(settings.header_column_count()).unwrap_or(u8::MAX);
    if u32::from(header_columns) > columns {
        return Err(BodyTableHeaderSettingsError::InvalidSettings {
            reason: BodyTableHeaderSettingsInvalidReason::HeaderColumnsExceedTable {
                header_columns,
                table_columns: columns,
            },
        });
    }
    Ok(())
}

fn rewrite_headers(
    source: &Package,
    target: &table_lock::BodyTableTarget,
    before: Settings,
    after: Settings,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableHeaderSettingsError> {
    let source_catalog = &source.state.source;
    let physical_limits = source_catalog.limits();
    budget
        .charge_source_catalog(source_catalog)
        .map_err(map_lock_error)?;
    table_lock::validate_body_table_target(source, target, budget).map_err(map_lock_error)?;
    let source_message = model_message(source, target)?;
    let snapshot = decode_snapshot(&source_message.data, budget)?;
    if settings_from_snapshot(snapshot)? != before {
        return Err(BodyTableHeaderSettingsError::InvalidSource);
    }
    validate_requested(after, snapshot.rows(), snapshot.columns())?;
    if target.explicit_locked == Some(true) {
        return Err(BodyTableHeaderSettingsError::TableLocked);
    }
    validate_dependencies(source, target, before, after, budget)?;
    let (component_name, stream_length) =
        table_lock::preflight_body_table_component(source, target, budget)
            .map_err(map_lock_error)?;
    let component = source_catalog
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    if component.name() != component_name {
        return Err(BodyTableHeaderSettingsError::InvalidSource);
    }
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyTableHeaderSettingsError::UnsupportedSource);
    }
    budget
        .charge_payload_work(entry.data().len())
        .map_err(map_lock_error)?;
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let original_message_length = source_message.data.len();
    let rewritten_message_bound = table_header_rewrite_bound(original_message_length)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let message_count = component
        .archive()
        .objects
        .iter()
        .try_fold(0usize, |count, object| {
            count.checked_add(object.messages.len())
        })
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let archive_bound = stream_length
        .checked_sub(original_message_length)
        .and_then(|value| value.checked_add(rewritten_message_bound))
        .and_then(|value| {
            value.checked_add(
                message_count
                    .checked_mul(MAX_VARINT_BYTES.saturating_mul(3))?
                    .checked_add(MAX_VARINT_BYTES.saturating_mul(2))?,
            )
        })
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let compressed_bound = table_lock::snappy_compressed_bound(archive_bound)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| {
            BodyTableHeaderSettingsError::LimitExceeded {
                kind: BodyTableHeaderSettingsLimitKind::EntryBytes,
                observed: u64::MAX,
                maximum: physical_limits.max_entry_bytes(),
            }
        })?;
    let replacement_compressed_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => table_lock::deflate_compressed_bound(compressed_bound)
            .ok_or(BodyTableHeaderSettingsError::InvalidSource)?,
        _ => return Err(BodyTableHeaderSettingsError::UnsupportedSource),
    };
    let package_output_bound = source_catalog
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(replacement_compressed_bound))
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    budget
        .charge_output_bytes(rewritten_message_bound)
        .and_then(|_| budget.charge_output_bytes(archive_bound))
        .and_then(|_| budget.charge_output_bytes(compressed_bound))
        .and_then(|_| budget.charge_output_bytes(replacement_compressed_bound))
        .and_then(|_| budget.charge_output_bytes(package_output_bound))
        .and_then(|_| budget.charge_payload_bytes(archive_bound))
        .and_then(|_| budget.charge_total_payload_bytes(archive_bound))
        .and_then(|_| budget.charge_payload_work(archive_bound))
        .and_then(|_| budget.charge_payload_work(compressed_bound))
        .and_then(|_| budget.charge_payload_work(replacement_compressed_bound))
        .and_then(|_| budget.charge_payload_work(package_output_bound))
        .map_err(map_lock_error)?;
    let stream = litchi_iwa_core::SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    if stream.as_bytes().len() != stream_length {
        return Err(BodyTableHeaderSettingsError::InvalidSource);
    }
    let mut archive =
        litchi_iwa_core::Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    let object = archive
        .objects
        .get_mut(target.model_object_index)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableHeaderSettingsError::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, target.model_message_index)
        .map_err(map_page_layout_error)?;
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let actual = decode_snapshot(&message.data, budget)?;
    let actual_settings = settings_from_snapshot(actual)?;
    if actual_settings != before {
        return Err(BodyTableHeaderSettingsError::InvalidSource);
    }
    let after_write = TableHeaderSettingsWrite::new(
        after
            .header_rows
            .map(|value| value.get().try_into().unwrap_or(u32::MAX)),
        after
            .header_columns
            .map(|value| value.get().try_into().unwrap_or(u32::MAX)),
        after
            .footer_rows
            .map(|value| value.get().try_into().unwrap_or(u32::MAX)),
        after.header_rows_frozen,
        after.header_columns_frozen,
        after.repeating_header_rows_enabled,
        after.repeating_header_columns_enabled,
    );
    let wire_limits = budget.wire_limits();
    let options = DecodeOptions::new(
        message.data.len().max(1).min(wire_limits.max_input_bytes()),
        wire_limits.max_fields(),
        wire_limits.max_rewrite_work(),
        u32::try_from(wire_limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_output_bytes(rewritten_message_bound.min(wire_limits.max_output_bytes()));
    let (rewritten, report) =
        table_header_settings_codec::rewrite_table_header_settings_with_report(
            &message.data,
            after_write,
            options,
        )
        .map_err(map_codec_error)?;
    charge_rewrite_report(budget, report)?;
    let verified = settings_from_snapshot(decode_snapshot(&rewritten, budget)?)?;
    if verified != after {
        return Err(BodyTableHeaderSettingsError::Verification);
    }
    object
        .replace_message_preserving_header_with_limits(
            target.model_message_index,
            RawMessage {
                type_: target.model_message_type,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let raw = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if raw.len() > archive_bound {
        return Err(BodyTableHeaderSettingsError::Verification);
    }
    let compressed = litchi_iwa_core::SnappyStream::compress(&raw).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(BodyTableHeaderSettingsError::Verification);
    }
    let previews = root_preview_deletions(source_catalog, budget)?;
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            &previews,
            physical_limits,
        )
        .map_err(map_archive_error)?;
    if output.len() > package_output_bound {
        return Err(BodyTableHeaderSettingsError::Verification);
    }
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), physical_limits)
            .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&candidate_source)
        .map_err(map_lock_error)?;
    table_lock::charge_reopen_work(&candidate_source, budget).map_err(map_lock_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    if settings_at_target_with_budget(&candidate, target, budget)? != after
        || preview_count(&candidate, budget)? != 0
    {
        return Err(BodyTableHeaderSettingsError::Verification);
    }
    Ok(candidate)
}

fn charge_rewrite_report(
    budget: &mut table_lock::WireBudget,
    report: table_header_settings_codec::RewriteReport,
) -> Result<(), BodyTableHeaderSettingsError> {
    budget
        .charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)
        .map_err(map_lock_error)
}

fn table_header_rewrite_bound(input_len: usize) -> Option<usize> {
    input_len.checked_add(
        7usize
            .checked_mul(MAX_VARINT_BYTES.saturating_add(1))?
            .checked_add(MAX_VARINT_BYTES.saturating_mul(2))?,
    )
}

fn validate_dependencies(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    before: Settings,
    after: Settings,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHeaderSettingsError> {
    let header_counts_changed =
        before.header_rows != after.header_rows || before.header_columns != after.header_columns;
    let section_counts_changed = header_counts_changed || before.footer_rows != after.footer_rows;
    let repeating_changed = before.repeating_header_rows_enabled
        != after.repeating_header_rows_enabled
        || before.repeating_header_columns_enabled != after.repeating_header_columns_enabled;
    let model = model_message(package, target)?;
    let model_view =
        WireView::parse(&model.data).map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
    budget
        .charge_payload_work(model.data.len())
        .map_err(map_lock_error)?;
    validate_known_dependency_fields(
        &model_view,
        header_counts_changed,
        section_counts_changed,
        budget,
    )?;
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
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    validate_known_reference_fields(model_info, &model_view, &[85, 86], target, budget)?;
    reject_unsupported_metadata(
        package,
        target.model_component_index,
        target.model_object_index,
        target.model_message_index,
        section_counts_changed,
        budget,
    )?;
    let info_object = package
        .state
        .source
        .components()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let info_message = info_object
        .messages
        .get(target.info_message_index)
        .filter(|message| TABLE_INFO_MESSAGE_TYPES.contains(&message.type_))
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let info_view = WireView::parse(&info_message.data)
        .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
    budget
        .charge_payload_work(info_message.data.len())
        .map_err(map_lock_error)?;
    validate_table_info_dependency_fields(
        &info_view,
        header_counts_changed,
        section_counts_changed,
    )?;
    let info_metadata = info_object
        .archive_info
        .message_infos
        .get(target.info_message_index)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    validate_known_reference_fields(info_metadata, &info_view, &[4, 5, 15, 17], target, budget)?;
    reject_unsupported_metadata(
        package,
        target.component_index,
        target.object_index,
        target.info_message_index,
        section_counts_changed,
        budget,
    )?;
    if repeating_changed {
        let sheet = package
            .state
            .source
            .components()
            .get_index(target.sheet_component_index)
            .and_then(|component| component.archive().objects.get(target.sheet_object_index))
            .and_then(|object| object.messages.get(target.sheet_message_index))
            .filter(|message| message.type_ == target.sheet_message_type)
            .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
        let view = WireView::parse(&sheet.data)
            .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
        budget
            .charge_payload_work(sheet.data.len())
            .map_err(map_lock_error)?;
        if view.fields().any(|field| field.number() == 4) {
            return Err(BodyTableHeaderSettingsError::UnsupportedDependency);
        }
        reject_unsupported_metadata(
            package,
            target.body_component_index,
            target.body_object_index,
            target.body_message_index,
            repeating_changed,
            budget,
        )?;
    }
    Ok(())
}

fn validate_known_dependency_fields(
    view: &WireView<'_>,
    header_counts_changed: bool,
    section_counts_changed: bool,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHeaderSettingsError> {
    let mut seen = [false; 5];
    let mut active_pivot_or_group = false;
    let mut header_count_dependency = false;
    let mut unsupported_pivot_dependency = false;
    for field in view.fields() {
        let Some(slot) = (match field.number() {
            81 => Some(0),
            83 => Some(1),
            84 => Some(2),
            85 => Some(3),
            86 => Some(4),
            _ => None,
        }) else {
            continue;
        };
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
        if seen[slot] || field.wire_type() != 2 {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        seen[slot] = true;
        if !field.payload().is_empty() {
            header_count_dependency = true;
            match field.number() {
                81 => {
                    active_pivot_or_group |=
                        category_owner_grouping_active(field.payload(), budget)?;
                },
                83 => active_pivot_or_group = true,
                85 => {
                    unsupported_pivot_dependency = true;
                    active_pivot_or_group = true;
                },
                86 => active_pivot_or_group = true,
                _ => {},
            }
        }
    }
    if unsupported_pivot_dependency
        || header_counts_changed && header_count_dependency
        || section_counts_changed && active_pivot_or_group
    {
        return Err(BodyTableHeaderSettingsError::UnsupportedDependency);
    }
    Ok(())
}

fn category_owner_grouping_active(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<bool, BodyTableHeaderSettingsError> {
    let view = WireView::parse_with_limits(source, budget.wire_limits())
        .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
    budget
        .charge_codec_report(view.len(), source.len().saturating_mul(2), 1, 0)
        .map_err(map_lock_error)?;
    let mut active = false;
    let mut seen = false;
    for field in view.fields() {
        if field.number() != 2 {
            continue;
        }
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
        if field.wire_type() != 2 {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        seen = true;
        let group = WireView::parse_with_limits(field.payload(), budget.wire_limits())
            .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
        budget
            .charge_codec_report(group.len(), field.payload().len().saturating_mul(2), 2, 0)
            .map_err(map_lock_error)?;
        let mut enabled = None;
        for nested in group.fields() {
            if nested.number() != 6 {
                continue;
            }
            nested
                .validate_canonical_framing()
                .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
            if nested.wire_type() != 0 || enabled.is_some() {
                return Err(BodyTableHeaderSettingsError::InvalidSource);
            }
            let (value, length) = decode_varint_from_bytes(nested.payload())
                .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
            if length != nested.payload().len() || value > 1 {
                return Err(BodyTableHeaderSettingsError::InvalidSource);
            }
            enabled = Some(value == 1);
        }
        active |= enabled.ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    }
    if !seen {
        return Err(BodyTableHeaderSettingsError::InvalidSource);
    }
    Ok(active)
}

fn validate_table_info_dependency_fields(
    view: &WireView<'_>,
    header_counts_changed: bool,
    section_counts_changed: bool,
) -> Result<(), BodyTableHeaderSettingsError> {
    let mut seen = [false; 7];
    for field in view.fields() {
        let Some(slot) = (match field.number() {
            4 => Some(0),
            5 => Some(1),
            7 => Some(2),
            8 => Some(3),
            15 => Some(4),
            16 => Some(5),
            17 => Some(6),
            _ => None,
        }) else {
            continue;
        };
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
        if seen[slot] {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        seen[slot] = true;
        let active = match field.number() {
            4 | 7 | 8 => header_counts_changed && !field.payload().is_empty(),
            5 | 15 | 17 => {
                (header_counts_changed || section_counts_changed) && !field.payload().is_empty()
            },
            16 => {
                if field.wire_type() != 0 {
                    return Err(BodyTableHeaderSettingsError::InvalidSource);
                }
                let (value, length) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
                if length != field.payload().len() || value > 1 {
                    return Err(BodyTableHeaderSettingsError::InvalidSource);
                }
                (header_counts_changed || section_counts_changed) && value == 1
            },
            _ => false,
        };
        if active {
            return Err(BodyTableHeaderSettingsError::UnsupportedDependency);
        }
    }
    Ok(())
}

fn reject_unsupported_metadata(
    package: &Package,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    changed: bool,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHeaderSettingsError> {
    if !changed {
        return Ok(());
    }
    let object = package
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(object_index))
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(BodyTableHeaderSettingsError::InvalidSource)?;
    budget
        .charge_payload_items(info.field_infos.len().saturating_add(1))
        .and_then(|_| {
            budget.charge_payload_references(
                info.object_references
                    .len()
                    .saturating_add(info.data_references.len()),
            )
        })
        .and_then(|_| budget.charge_payload_work(info.field_infos.len()))
        .map_err(map_lock_error)?;
    for field in &info.field_infos {
        budget
            .charge_payload_references(
                field
                    .object_references
                    .len()
                    .saturating_add(field.data_references.len()),
            )
            .and_then(|_| budget.charge_payload_work(field.path.as_slice().len()))
            .map_err(map_lock_error)?;
    }
    if info.type_ != message.type_
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
        || !info.data_references.is_empty()
    {
        return Err(BodyTableHeaderSettingsError::InvalidSource);
    }
    for field in &info.field_infos {
        if !field.data_references.is_empty() {
            return Err(BodyTableHeaderSettingsError::UnsupportedDependency);
        }
    }
    Ok(())
}

fn validate_known_reference_fields(
    info: &litchi_iwa_core::MessageInfo,
    view: &WireView<'_>,
    fields: &[u32],
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHeaderSettingsError> {
    for field in view.fields() {
        if !fields.contains(&field.number()) {
            continue;
        }
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
        if field.wire_type() != 2 {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        let identifier = reference_identifier(field.payload(), budget)?;
        if identifier == 0
            || identifier == 1
            || identifier == target.sheet_identifier.get()
            || identifier == target.drawable_identifier.get()
            || identifier == target.model_identifier.get()
        {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        if info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count()
            != 1
        {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        let mut declarations = info.field_infos.iter().filter(|candidate| {
            candidate.path.as_slice() == [field.number()]
                && candidate.object_references.as_slice() == [identifier]
                && candidate.data_references.is_empty()
        });
        if declarations.next().is_none() || declarations.next().is_some() {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
    }
    // Check the reverse direction as well: a FieldInfo declaration on one of
    // these schema paths must have exactly one matching raw field and the
    // same identifier.  A declaration on a descendant/unknown path is not a
    // substitute for the selected dependency path.
    for declaration in &info.field_infos {
        let Some(&field_number) = fields
            .iter()
            .find(|number| declaration.path.as_slice() == [**number])
        else {
            continue;
        };
        if declaration.object_references.len() != 1 || !declaration.data_references.is_empty() {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        let mut raw = view.fields().filter(|field| field.number() == field_number);
        let Some(raw_field) = raw.next() else {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        };
        if raw.next().is_some() || raw_field.wire_type() != 2 {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        let identifier = reference_identifier(raw_field.payload(), budget)?;
        if declaration.object_references[0] != identifier {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
    }
    Ok(())
}

fn reference_identifier(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<u64, BodyTableHeaderSettingsError> {
    let view = WireView::parse_with_limits(source, budget.wire_limits())
        .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
    budget
        .charge_codec_report(view.len(), source.len().saturating_mul(2), 1, 0)
        .map_err(map_lock_error)?;
    let mut identifier = None;
    for field in view.fields() {
        if field.number() != 1 {
            continue;
        }
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
        if field.wire_type() != 0 || identifier.is_some() {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        let (value, length) = decode_varint_from_bytes(field.payload())
            .map_err(|_| BodyTableHeaderSettingsError::InvalidSource)?;
        if length != field.payload().len() {
            return Err(BodyTableHeaderSettingsError::InvalidSource);
        }
        identifier = Some(value);
    }
    identifier.ok_or(BodyTableHeaderSettingsError::InvalidSource)
}

fn root_preview_deletions(
    catalog: &SourceCatalog,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<&'static str>, BodyTableHeaderSettingsError> {
    let mut deletions = Vec::new();
    deletions
        .try_reserve_exact(ROOT_PREVIEW_NAMES.len())
        .map_err(|_| BodyTableHeaderSettingsError::Allocation {
            amount: ROOT_PREVIEW_NAMES.len(),
        })?;
    for name in ROOT_PREVIEW_NAMES {
        budget
            .charge_payload_work(name.len())
            .map_err(map_lock_error)?;
        if catalog.package().iter().any(|entry| entry.name() == name) {
            deletions.push(name);
        }
    }
    Ok(deletions)
}

fn preview_count(
    package: &Package,
    budget: &mut table_lock::WireBudget,
) -> Result<usize, BodyTableHeaderSettingsError> {
    let mut count = 0;
    for name in ROOT_PREVIEW_NAMES {
        budget
            .charge_payload_work(name.len())
            .map_err(map_lock_error)?;
        if package
            .state
            .source
            .package()
            .iter()
            .any(|entry| entry.name() == name)
        {
            count += 1;
        }
    }
    Ok(count)
}

fn fingerprint(
    bytes: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<u64, BodyTableHeaderSettingsError> {
    budget.charge_input_source(bytes).map_err(map_lock_error)?;
    budget
        .charge_payload_work(bytes.len())
        .map_err(map_lock_error)?;
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    Ok(value)
}

fn bytes_equal(
    left: &[u8],
    right: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<bool, BodyTableHeaderSettingsError> {
    if left.len() != right.len() {
        return Ok(false);
    }
    budget
        .charge_payload_work(left.len())
        .map_err(map_lock_error)?;
    Ok(left == right)
}

fn reopen_target(
    source: &Package,
    target: Arc<[u8]>,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableHeaderSettingsError> {
    budget
        .charge_input_source(target.as_ref())
        .map_err(map_lock_error)?;
    let catalog =
        SourceCatalog::from_shared_bytes_with_limits(target, source.state.source.limits())
            .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .map_err(map_lock_error)?;
    table_lock::charge_reopen_work(&catalog, budget).map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableHeaderSettingsError {
    match error {
        table_lock::BodyTableLockError::TableNotFound => {
            BodyTableHeaderSettingsError::TableNotFound
        },
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableHeaderSettingsError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::AmbiguousSelector => {
            BodyTableHeaderSettingsError::AmbiguousSelector
        },
        table_lock::BodyTableLockError::UnsupportedSource => {
            BodyTableHeaderSettingsError::UnsupportedSource
        },
        table_lock::BodyTableLockError::InvalidSource => {
            BodyTableHeaderSettingsError::InvalidSource
        },
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableHeaderSettingsError::LimitExceeded {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableHeaderSettingsError::Allocation { amount }
        },
        table_lock::BodyTableLockError::Verification => BodyTableHeaderSettingsError::Verification,
        table_lock::BodyTableLockError::PatchConflict => {
            BodyTableHeaderSettingsError::PatchConflict
        },
    }
}

const fn map_lock_limit(
    kind: table_lock::BodyTableLockLimitKind,
) -> BodyTableHeaderSettingsLimitKind {
    use table_lock::BodyTableLockLimitKind as Lock;
    match kind {
        Lock::InputBytes => BodyTableHeaderSettingsLimitKind::InputBytes,
        Lock::OutputBytes => BodyTableHeaderSettingsLimitKind::OutputBytes,
        Lock::Entries => BodyTableHeaderSettingsLimitKind::Entries,
        Lock::EntryBytes => BodyTableHeaderSettingsLimitKind::EntryBytes,
        Lock::TotalEntryBytes => BodyTableHeaderSettingsLimitKind::TotalEntryBytes,
        Lock::PackageBytes => BodyTableHeaderSettingsLimitKind::PayloadItems,
        Lock::PayloadBytes => BodyTableHeaderSettingsLimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => BodyTableHeaderSettingsLimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => BodyTableHeaderSettingsLimitKind::PayloadObjects,
        Lock::PayloadMessages => BodyTableHeaderSettingsLimitKind::PayloadMessages,
        Lock::PayloadItems => BodyTableHeaderSettingsLimitKind::PayloadItems,
        Lock::PayloadReferences => BodyTableHeaderSettingsLimitKind::PayloadReferences,
        Lock::WireBytes => BodyTableHeaderSettingsLimitKind::WireBytes,
        Lock::WireFields => BodyTableHeaderSettingsLimitKind::WireFields,
        Lock::WireNesting => BodyTableHeaderSettingsLimitKind::WireNesting,
        Lock::WireWork => BodyTableHeaderSettingsLimitKind::WireWork,
    }
}

fn map_common_wire_error(error: litchi_iwa_common::Error) -> BodyTableHeaderSettingsError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => BodyTableHeaderSettingsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    BodyTableHeaderSettingsLimitKind::WireBytes
                },
                litchi_iwa_common::LimitKind::Fields => {
                    BodyTableHeaderSettingsLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => {
                    BodyTableHeaderSettingsLimitKind::WireNesting
                },
                litchi_iwa_common::LimitKind::RewriteWork => {
                    BodyTableHeaderSettingsLimitKind::WireWork
                },
                litchi_iwa_common::LimitKind::OutputBytes
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    BodyTableHeaderSettingsLimitKind::WireBytes
                },
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            BodyTableHeaderSettingsError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => {
            BodyTableHeaderSettingsError::InvalidSource
        },
    }
}

fn map_codec_error(error: DecodeError) -> BodyTableHeaderSettingsError {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return BodyTableHeaderSettingsError::LimitExceeded {
            kind: BodyTableHeaderSettingsLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return BodyTableHeaderSettingsError::LimitExceeded {
            kind: BodyTableHeaderSettingsLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.output_limit_values() {
        return BodyTableHeaderSettingsError::LimitExceeded {
            kind: BodyTableHeaderSettingsLimitKind::WireOutputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return BodyTableHeaderSettingsError::Allocation { amount };
    }
    match error.wire_resource_limit() {
        Some(WireResourceLimit::Bytes { observed, maximum }) => {
            BodyTableHeaderSettingsError::LimitExceeded {
                kind: BodyTableHeaderSettingsLimitKind::WireBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(WireResourceLimit::Nesting { observed, maximum }) => {
            BodyTableHeaderSettingsError::LimitExceeded {
                kind: BodyTableHeaderSettingsLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        Some(_) => BodyTableHeaderSettingsError::InvalidSource,
        None => BodyTableHeaderSettingsError::InvalidSource,
    }
}

fn map_page_layout_error(error: page_layout::PageLayoutError) -> BodyTableHeaderSettingsError {
    match error {
        page_layout::PageLayoutError::UnsupportedSource => {
            BodyTableHeaderSettingsError::UnsupportedSource
        },
        page_layout::PageLayoutError::InvalidSource => BodyTableHeaderSettingsError::InvalidSource,
        page_layout::PageLayoutError::LimitExceeded {
            observed, maximum, ..
        } => BodyTableHeaderSettingsError::LimitExceeded {
            kind: BodyTableHeaderSettingsLimitKind::PayloadBytes,
            observed,
            maximum,
        },
        page_layout::PageLayoutError::Allocation { amount } => {
            BodyTableHeaderSettingsError::Allocation { amount }
        },
        _ => BodyTableHeaderSettingsError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyTableHeaderSettingsError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableHeaderSettingsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    BodyTableHeaderSettingsLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    BodyTableHeaderSettingsLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => BodyTableHeaderSettingsLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    BodyTableHeaderSettingsLimitKind::PayloadItems
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => {
                    BodyTableHeaderSettingsLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    BodyTableHeaderSettingsLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyTableHeaderSettingsLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    BodyTableHeaderSettingsLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyTableHeaderSettingsError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => BodyTableHeaderSettingsError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyTableHeaderSettingsError {
    match error {
        litchi_iwa_core::Error::Limit {
            observed, maximum, ..
        } => BodyTableHeaderSettingsError::LimitExceeded {
            kind: BodyTableHeaderSettingsLimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableHeaderSettingsError::Allocation { amount: requested }
        },
        _ => BodyTableHeaderSettingsError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> BodyTableHeaderSettingsError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => BodyTableHeaderSettingsError::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => {
            BodyTableHeaderSettingsError::LimitExceeded {
                kind: BodyTableHeaderSettingsLimitKind::PayloadObjects,
                observed: observed as u64,
                maximum: limit as u64,
            }
        },
        PackageError::PayloadLimit { observed, limit } => {
            BodyTableHeaderSettingsError::LimitExceeded {
                kind: BodyTableHeaderSettingsLimitKind::PayloadBytes,
                observed: observed as u64,
                maximum: limit as u64,
            }
        },
        _ => BodyTableHeaderSettingsError::InvalidSource,
    }
}
