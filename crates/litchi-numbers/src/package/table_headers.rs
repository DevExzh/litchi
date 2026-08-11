//! Exact-source transactions for one rooted table's header and footer settings.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "the focused boundary exhaustively redacts lower-layer error families"
)]

mod api;
mod dependencies;
mod error;
pub(super) mod ownership;
pub(super) mod resolve;
pub(super) mod rewrite;

use std::{fmt, sync::Arc};

use litchi_iwa_archive::package::ExactArtifacts;
use thiserror::Error as ThisError;

use super::Package;
use crate::table::{headers::Settings, lock::State as LockState};

use dependencies::validate_dependencies;
use ownership::validate_selected_ownership;
use resolve::validate_requested;
use rewrite::{
    clone_selected_payload, physical_source, preflight_transaction_work, rewrite,
    root_preview_deletions, verify_exact_locality,
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
pub(super) struct Target {
    pub(super) sheet_position: usize,
    pub(super) table_position: usize,
    pub(super) model_identifier: u64,
    pub(super) sheet_identifier: u64,
    pub(super) drawable_identifier: u64,
    pub(super) drawable_position: usize,
    pub(super) sheet_component_index: usize,
    pub(super) sheet_object_index: usize,
    pub(super) sheet_message_index: usize,
    pub(super) sheet_message_type: u32,
    pub(super) info_component_index: usize,
    pub(super) info_object_index: usize,
    pub(super) info_message_index: usize,
    pub(super) info_message_type: u32,
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) settings: Settings,
    pub(super) rows: u32,
    pub(super) columns: u32,
    pub(super) locked: LockState,
}
