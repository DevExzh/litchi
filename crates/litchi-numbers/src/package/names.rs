//! Exact-source transactions for rooted Numbers sheet and table names.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction keeps its public API before private graph/rewrite helpers and maps lower-layer errors into one redacted Numbers boundary."
)]

use std::collections::HashSet;
use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::{
    SourceCatalog,
    package::{Entry, EntryEdit},
};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireView, patch_length_delimited_field, transform_length_delimited_field},
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{numbers_names_codec, table_info_codec, tn, tsce, tst};
use prost::Message;
use thiserror::Error as ThisError;

use super::{
    Error as ReadError, FORM_BASED_SHEET_MESSAGE_TYPE, LEGACY_TABLE_INFO_MESSAGE_TYPE, Package,
    Resolved, SHEET_MESSAGE_TYPE, TABLE_INFO_MESSAGE_TYPE, TABLE_MODEL_MESSAGE_TYPE,
    table_info_decode_options,
};
use crate::table::lock::State as LockState;
use crate::{SheetSelector, TableSelector};

const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const LEGACY_TABLE_MODEL_MESSAGE_TYPE: u32 = 6_000;
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;
const SHEET_NAME_FIELD: u32 = 1;
const FORM_SHEET_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_NAME_FIELD: u32 = 8;
const MIN_SIGN_EXTENDED_I32: u64 = 0xffff_ffff_8000_0000;

/// A content-free semantic location in a names transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// One rooted sheet at its zero-based workbook position.
    Sheet { position: usize },
    /// One rooted table at its zero-based position within a sheet.
    Table { sheet: usize, table: usize },
}

/// The public invariant violated by a requested name.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum InvalidReason {
    /// Names must contain at least one UTF-8 byte.
    Empty,
    /// Names cannot contain the NUL scalar value.
    ContainsNul,
}

impl fmt::Display for InvalidReason {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Empty => "the name is empty",
            Self::ContainsNul => "the name contains NUL",
        })
    }
}

/// A finite resource governed while names are staged or published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// ZIP members retained by the package.
    Entries,
    /// Bytes in one ZIP member.
    EntryBytes,
    /// Aggregate ZIP member bytes.
    TotalEntryBytes,
    /// Package names or structural metadata bytes.
    PackageBytes,
    /// Bytes in one decoded native payload container.
    PayloadBytes,
    /// Aggregate decoded native payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected by the transaction.
    PayloadObjects,
    /// Native payload messages inspected by the transaction.
    PayloadMessages,
    /// Native framing or metadata items inspected by the transaction.
    PayloadItems,
    /// Native object references inspected by the transaction.
    PayloadReferences,
    /// Bytes inspected by a strict name projection.
    WireBytes,
    /// Fields inspected by a strict name projection.
    WireFields,
    /// Nested depth inspected by a strict name projection.
    WireNesting,
    /// Aggregate strict wire work.
    WireWork,
    /// Staged semantic rename operations.
    Operations,
    /// Aggregate retained UTF-8 name bytes.
    NameBytes,
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
            Self::Operations => "rename operations",
            Self::NameBytes => "name bytes",
        })
    }
}

/// A content-redacted failure from a Numbers names transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No rooted sheet matched a selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No rooted table matched a selector inside the selected sheet.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// A requested name violates a public semantic invariant.
    #[error("invalid Numbers name: {reason}")]
    InvalidName {
        /// Stable content-free reason.
        reason: InvalidReason,
    },
    /// The final batch would contain a duplicate name in one semantic namespace.
    #[error("the Numbers names transaction would create a duplicate semantic name")]
    DuplicateName {
        /// Location of the later colliding value.
        path: Path,
    },
    /// The same semantic owner was selected more than once in one batch.
    #[error("the Numbers names transaction repeats a previously selected target")]
    DuplicateTarget {
        /// Repeated semantic target.
        path: Path,
    },
    /// The package source cannot publish a preservation-safe changed edit.
    #[error("the Numbers package source does not support exact name editing")]
    UnsupportedSource,
    /// The selected rooted native graph cannot be edited without ambiguity.
    #[error("the selected Numbers name owner is invalid")]
    InvalidSource,
    /// A changed table-name operation targets an interactively locked table.
    #[error("the selected Numbers table is locked")]
    TableLocked {
        /// Locked semantic table location.
        path: Path,
    },
    /// A modeled dependency cannot be updated by this bounded vertical.
    #[error("a Numbers name dependency prevents this transaction")]
    UnsupportedDependency {
        /// Selected name whose dependency blocks publication.
        path: Path,
    },
    /// A finite transaction resource ceiling was exceeded.
    #[error("Numbers names {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Numbers names transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
    },
    /// Complete candidate reopening did not reproduce the requested final names.
    #[error("the edited Numbers names failed semantic verification")]
    Verification,
    /// The supplied patch was not created from this exact package artifact.
    #[error("the Numbers names patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
enum Location {
    Sheet(usize),
    Table(usize, usize),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Namespace {
    Sheets,
    Tables(usize),
}

#[derive(Clone, PartialEq, Eq, Hash)]
struct Destination {
    namespace: Namespace,
    name: Arc<str>,
}

impl Location {
    const fn path(self) -> Path {
        match self {
            Self::Sheet(position) => Path::Sheet { position },
            Self::Table(sheet, table) => Path::Table { sheet, table },
        }
    }
}

#[derive(Clone, PartialEq, Eq)]
struct Operation {
    location: Location,
    before: Arc<str>,
    after: Arc<str>,
}

impl fmt::Debug for Operation {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Operation")
            .field("path", &self.location.path())
            .finish_non_exhaustive()
    }
}

/// A batch of sheet/table renames resolved against one immutable base snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    operations: Vec<Operation>,
    targets: HashSet<Location>,
    destinations: HashSet<Destination>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    pub(super) fn new(source: &'a Package) -> Self {
        Self {
            source,
            operations: Vec::new(),
            targets: HashSet::new(),
            destinations: HashSet::new(),
        }
    }

    /// Stage one rooted sheet rename using a selector resolved against the base snapshot.
    ///
    /// # Costs
    ///
    /// Name selection visits at most every rooted sheet. Indexed staged-target
    /// validation is expected `O(1)`; replacement text is copied once.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, name, collision, limit, or allocation failure.
    pub fn rename_sheet<'selector>(
        mut self,
        selector: impl Into<SheetSelector<'selector>>,
        name: &str,
    ) -> Result<Self, Error> {
        validate_requested_name(name)?;
        let sheet = self
            .source
            .document()
            .sheet(selector)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::SheetNotFound)?;
        let location = Location::Sheet(sheet.index());
        self.stage(location, sheet.name(), name)?;
        Ok(self)
    }

    /// Stage one rooted table rename using selectors resolved against the base snapshot.
    ///
    /// Table names are unique only within their owning sheet. Both selectors
    /// always resolve against the immutable base, even after earlier staged
    /// sheet or table renames.
    ///
    /// # Costs
    ///
    /// Each selector visits its rooted namespace at most once. Indexed
    /// staged-target validation is expected `O(1)`; replacement text is copied
    /// once.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, name, collision, limit, or allocation failure.
    pub fn rename_table<'sheet, 'table>(
        mut self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        name: &str,
    ) -> Result<Self, Error> {
        validate_requested_name(name)?;
        let sheet = self
            .source
            .document()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::SheetNotFound)?;
        let (position, semantic_table) = match table.into() {
            TableSelector::Index(position) => (
                position,
                sheet.tables().nth(position).ok_or(Error::TableNotFound)?,
            ),
            TableSelector::Name(requested_name) => sheet
                .tables()
                .enumerate()
                .find(|(_position, candidate)| candidate.name() == requested_name)
                .ok_or(Error::TableNotFound)?,
        };
        let location = Location::Table(sheet.index(), position);
        self.stage(location, semantic_table.name(), name)?;
        Ok(self)
    }

    fn stage(&mut self, location: Location, before: &str, after: &str) -> Result<(), Error> {
        if self.targets.contains(&location) {
            return Err(Error::DuplicateTarget {
                path: location.path(),
            });
        }
        let maximum = self
            .source
            .state
            .options
            .semantic()
            .max_sheets()
            .saturating_add(self.source.state.options.semantic().max_tables());
        let observed = self.operations.len().saturating_add(1);
        if observed > maximum {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Operations,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(maximum),
            });
        }
        let maximum_name_bytes = self.source.state.options.semantic().max_output_text_bytes();
        if after.len() > maximum_name_bytes {
            return Err(Error::LimitExceeded {
                kind: LimitKind::NameBytes,
                observed: usize_as_u64(after.len()),
                maximum: usize_as_u64(maximum_name_bytes),
            });
        }
        let before = try_arc_str(before)?;
        let after = try_arc_str(after)?;
        let destination = Destination {
            namespace: match location {
                Location::Sheet(_) => Namespace::Sheets,
                Location::Table(sheet, _) => Namespace::Tables(sheet),
            },
            name: Arc::clone(&after),
        };
        if self.destinations.contains(&destination) {
            return Err(Error::DuplicateName {
                path: location.path(),
            });
        }
        let operation = Operation {
            location,
            before,
            after,
        };
        self.operations
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                amount: self.operations.len().saturating_add(1),
            })?;
        self.targets
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                amount: self.targets.len().saturating_add(1),
            })?;
        self.destinations
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                amount: self.destinations.len().saturating_add(1),
            })?;
        let inserted_target = self.targets.insert(location);
        let inserted_destination = self.destinations.insert(destination);
        debug_assert!(inserted_target && inserted_destination);
        self.operations.push(operation);
        Ok(())
    }

    /// Validate and atomically publish the staged final name set.
    ///
    /// A semantic no-op shares the original snapshot and bypasses changed-only
    /// framing, cache, lock, and dependency guards. A changed batch rewrites
    /// every touched component once, removes existing root previews, and fully
    /// reopens the candidate under the retained package limits.
    ///
    /// # Costs
    ///
    /// Final-name validation is linear in the rooted semantic catalog. The
    /// conservative dependency/ownership plan is bounded before native work by
    /// [`LimitKind::WireWork`]; each touched component is then parsed and
    /// rewritten once, followed by one complete candidate reopen and an exact
    /// locality scan. Changed commits retain exact source and target artifacts
    /// in the process-local reversible patch.
    ///
    /// # Errors
    ///
    /// Returns without modifying the source when final-name validation,
    /// protection, dependencies, native ownership, resource limits,
    /// preservation-safe reassembly, or semantic readback fails.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self.source, self.operations)
    }
}

/// An exact-source-checked, reversible process-local names patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    operations: Arc<[Operation]>,
    native: Arc<[NativeOperation]>,
    direction: Direction,
    touched_components: usize,
    previews: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Direction {
    Forward,
    Reverse,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return whether the patch changes no semantic name.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.operations
            .iter()
            .all(|operation| operation.before == operation.after)
    }

    /// Return the number of semantic name operations retained by this patch.
    #[must_use]
    pub fn operation_count(&self) -> usize {
        self.operations
            .iter()
            .filter(|operation| operation.before != operation.after)
            .count()
    }

    /// Return an exact inverse that restores the complete original artifact.
    ///
    /// # Costs
    ///
    /// This is `O(1)` and shares both artifacts and the private operation plan.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            operations: Arc::clone(&self.operations),
            native: Arc::clone(&self.native),
            direction: match self.direction {
                Direction::Forward => Direction::Reverse,
                Direction::Reverse => Direction::Forward,
            },
            touched_components: self.touched_components,
            previews: self.previews,
        }
    }
}

/// Compact, content-free evidence about a committed names transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    operations: usize,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            operations: 0,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(
        operations: usize,
        touched_components: usize,
        deleted_previews: usize,
    ) -> Self {
        Self {
            operations,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return the number of changed semantic rename operations.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }

    /// Return whether the transaction changed at least one semantic name.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.operations != 0
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of exact root preview members deleted by this publication.
    ///
    /// Applying an inverse restores retained source bytes and therefore reports
    /// zero deletions even when its forward patch originally removed previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the complete target package was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully reopened immutable result of one names transaction.
#[must_use = "a Numbers names commit contains the validated package snapshot"]
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

impl Package {
    /// Start an infallible empty names batch against this immutable snapshot.
    ///
    /// # Costs
    ///
    /// This is `O(1)` and does not allocate or inspect native package data.
    #[must_use]
    pub fn edit_names(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Apply an exact-source-checked reversible names patch.
    ///
    /// # Costs
    ///
    /// A changed patch reopens its retained target artifact once, verifies
    /// semantic state, and scans package members plus touched native components
    /// for exact locality. A no-op shares this package snapshot without
    /// reassembly or reparsing.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a stale, replayed, tampered, or
    /// cross-artifact patch, or a typed verification/resource error when the
    /// retained target cannot be reopened safely.
    pub fn apply_names(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if fingerprint(source.source_bytes()) != patch.source_fingerprint
            || source.source_bytes() != patch.source.as_ref()
            || !verify_operation_state(self, &patch.operations, patch.direction, false)
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
            Package::from_shared_bytes_with_options(Arc::clone(&patch.target), self.state.options)
                .map_err(map_candidate_read_error)?;
        if !verify_operation_state(&candidate, &patch.operations, patch.direction, true) {
            return Err(Error::Verification);
        }
        verify_exact_locality(
            self,
            &candidate,
            &patch.native,
            patch.direction,
            patch.previews,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.operation_count(),
                patch.touched_components,
                if matches!(patch.direction, Direction::Forward) {
                    patch.previews
                } else {
                    0
                },
            ),
        })
    }
}

fn validate_requested_name(name: &str) -> Result<(), Error> {
    if name.is_empty() {
        return Err(Error::InvalidName {
            reason: InvalidReason::Empty,
        });
    }
    if name.contains('\0') {
        return Err(Error::InvalidName {
            reason: InvalidReason::ContainsNul,
        });
    }
    Ok(())
}

fn try_arc_str(value: &str) -> Result<Arc<str>, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_allocation| Error::Allocation {
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(Arc::from(owned.into_boxed_str()))
}

fn commit_edit(source: &Package, mut operations: Vec<Operation>) -> Result<Commit, Error> {
    operations.sort_unstable_by_key(|operation| operation.location);
    let changed = operations
        .iter()
        .filter(|operation| operation.before != operation.after)
        .count();
    let source_catalog = physical_source(source)?;
    let source_bytes = source_catalog.shared_source();
    let source_fingerprint = fingerprint(&source_bytes);
    if changed == 0 {
        let patch = Patch {
            source: Arc::clone(&source_bytes),
            target: source_bytes,
            source_fingerprint,
            target_fingerprint: source_fingerprint,
            operations: operations.into(),
            native: Arc::from([]),
            direction: Direction::Forward,
            touched_components: 0,
            previews: 0,
        };
        return Ok(Commit {
            package: source.snapshot(),
            patch,
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if !source_catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    validate_final_names(source, &operations)?;
    validate_changed_work_budget(source, &operations)?;
    validate_changed_guards(source, &operations)?;
    let native = resolve_native_operations(source, &operations)?;
    let preview_names = root_preview_deletions(source_catalog)?;
    let (package, touched_components) = rewrite_names(source, &native, &preview_names)?;
    if !verify_operation_state(&package, &operations, Direction::Forward, true) {
        return Err(Error::Verification);
    }
    verify_exact_locality(
        source,
        &package,
        &native,
        Direction::Forward,
        preview_names.len(),
    )?;
    let target = physical_source(&package)?.shared_source();
    let target_fingerprint = fingerprint(&target);
    let patch = Patch {
        source: source_bytes,
        target,
        source_fingerprint,
        target_fingerprint,
        operations: operations.into(),
        native: native.into(),
        direction: Direction::Forward,
        touched_components,
        previews: preview_names.len(),
    };
    Ok(Commit {
        package,
        patch,
        diagnostics: Diagnostics::published(changed, touched_components, preview_names.len()),
    })
}

fn validate_final_names(source: &Package, operations: &[Operation]) -> Result<(), Error> {
    let limits = source.state.options.semantic();
    let mut names = HashSet::new();
    names
        .try_reserve(source.document().sheet_count())
        .map_err(|_allocation| Error::Allocation {
            amount: source.document().sheet_count(),
        })?;
    let mut total_name_bytes = 0usize;
    for sheet in source.document().sheets() {
        let sheet_name = final_name(operations, Location::Sheet(sheet.index()), sheet.name());
        if !names.insert(sheet_name) {
            return Err(Error::DuplicateName {
                path: Path::Sheet {
                    position: sheet.index(),
                },
            });
        }
        total_name_bytes = checked_name_bytes(
            total_name_bytes,
            sheet_name.len(),
            limits.max_output_text_bytes(),
        )?;
        let mut tables = HashSet::new();
        tables
            .try_reserve(sheet.table_count())
            .map_err(|_allocation| Error::Allocation {
                amount: sheet.table_count(),
            })?;
        for (table_position, table) in sheet.tables().enumerate() {
            let location = Location::Table(sheet.index(), table_position);
            let table_name = final_name(operations, location, table.name());
            if !tables.insert(table_name) {
                return Err(Error::DuplicateName {
                    path: location.path(),
                });
            }
            total_name_bytes = checked_name_bytes(
                total_name_bytes,
                table_name.len(),
                limits.max_output_text_bytes(),
            )?;
        }
    }
    Ok(())
}

fn checked_name_bytes(current: usize, added: usize, maximum: usize) -> Result<usize, Error> {
    let observed = current.saturating_add(added);
    if observed > maximum {
        return Err(Error::LimitExceeded {
            kind: LimitKind::NameBytes,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        });
    }
    Ok(observed)
}

fn validate_changed_work_budget(source: &Package, operations: &[Operation]) -> Result<(), Error> {
    let sheets = source.document().sheet_count();
    let tables = source
        .document()
        .sheets()
        .iter()
        .fold(0usize, |count, sheet| {
            count.saturating_add(sheet.table_count())
        });
    let objects = source.state.components.iter_objects().count();
    let changed = operations
        .iter()
        .filter(|operation| operation.before != operation.after)
        .count();
    // This deliberately overcharges every selected owner lookup as a complete
    // rooted-topology visit and the conservative pivot guard as a quadratic
    // table visit. It is checked before any changed-only native scan.
    let topology = sheets.saturating_add(tables).saturating_add(objects);
    let observed = changed
        .saturating_mul(topology)
        .saturating_add(tables.saturating_mul(tables))
        .saturating_add(objects);
    let maximum = source.state.options.semantic().max_formula_render_work();
    if observed > maximum {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        });
    }
    Ok(())
}

fn final_name<'a>(operations: &'a [Operation], location: Location, fallback: &'a str) -> &'a str {
    operations
        .binary_search_by_key(&location, |operation| operation.location)
        .ok()
        .and_then(|index| operations.get(index))
        .map_or(fallback, |operation| operation.after.as_ref())
}

fn verify_operation_state(
    package: &Package,
    operations: &[Operation],
    direction: Direction,
    target: bool,
) -> bool {
    operations.iter().all(|operation| {
        let forward_after = matches!(direction, Direction::Forward) == target;
        let expected = if forward_after {
            operation.after.as_ref()
        } else {
            operation.before.as_ref()
        };
        match operation.location {
            Location::Sheet(position) => package
                .document()
                .sheets()
                .get(position)
                .is_some_and(|sheet| sheet.name() == expected),
            Location::Table(sheet, table) => package
                .document()
                .sheets()
                .get(sheet)
                .and_then(|sheet| sheet.tables().nth(table))
                .is_some_and(|table| table.name() == expected),
        }
    })
}

fn validate_changed_guards(source: &Package, operations: &[Operation]) -> Result<(), Error> {
    validate_unique_rooted_model_ownership(source, operations)?;
    for operation in operations
        .iter()
        .filter(|operation| operation.before != operation.after)
    {
        if let Location::Table(sheet, table) = operation.location
            && source
                .table_lock(SheetSelector::index(sheet), TableSelector::index(table))
                .map_err(map_lock_error)?
                == LockState::Locked
        {
            return Err(Error::TableLocked {
                path: operation.location.path(),
            });
        }
    }
    validate_name_volatile_dependencies(source, operations)?;
    validate_rooted_pivot_dependencies(source, operations)?;
    Ok(())
}

fn validate_unique_rooted_model_ownership(
    source: &Package,
    operations: &[Operation],
) -> Result<(), Error> {
    let selected_count = operations
        .iter()
        .filter(|operation| {
            operation.before != operation.after && matches!(operation.location, Location::Table(..))
        })
        .count();
    if selected_count == 0 {
        return Ok(());
    }
    let mut selected_models = Vec::new();
    selected_models
        .try_reserve_exact(selected_count)
        .map_err(|_allocation| Error::Allocation {
            amount: selected_count,
        })?;
    for operation in operations.iter().filter(|operation| {
        operation.before != operation.after && matches!(operation.location, Location::Table(..))
    }) {
        let Location::Table(sheet, table) = operation.location else {
            continue;
        };
        let identifier = resolve_native_table(source, sheet, table)?.model_identifier;
        if selected_models.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        selected_models.push(identifier);
    }
    let rooted_infos = rooted_table_info_identifiers(source)?;
    let mut ownership_counts = Vec::new();
    ownership_counts
        .try_reserve_exact(selected_models.len())
        .map_err(|_allocation| Error::Allocation {
            amount: selected_models.len(),
        })?;
    ownership_counts.resize(selected_models.len(), 0usize);
    for info_identifier in rooted_infos {
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, info_identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource)?;
        let (_message_index, message) = unique_table_info(resolved)?.ok_or(Error::InvalidSource)?;
        let model_identifier = decode_table_info(&message.data)?
            .table_model()
            .identifier()
            .get();
        if let Some(position) = selected_models
            .iter()
            .position(|selected| *selected == model_identifier)
        {
            ownership_counts[position] = ownership_counts[position].saturating_add(1);
        }
    }
    if ownership_counts.iter().any(|count| *count != 1) {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_name_volatile_dependencies(
    source: &Package,
    operations: &[Operation],
) -> Result<(), Error> {
    let Some(path) = operations
        .iter()
        .find(|operation| operation.before != operation.after)
        .map(|operation| operation.location.path())
    else {
        return Ok(());
    };
    for identifier in rooted_formula_dependency_identifiers(source)? {
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource)?;
        let (message_index, message) =
            unique_message_index(resolved.messages, FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE)?
                .ok_or(Error::InvalidSource)?;
        let object = resolved_object(source, resolved)?;
        validate_message_metadata(object, message_index)?;
        let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
            .map_err(|_error| Error::InvalidSource)?;
        if let Some(owner_reference) = owner.formula_owner.as_ref() {
            require_local_reference(
                singular_length_payload(&message.data, 11)?,
                owner_reference.identifier,
            )?;
        }
        let nonempty = owner
            .volatile_dependencies
            .as_ref()
            .and_then(|dependencies| dependencies.volatile_sheet_table_name_cells.as_ref())
            .is_some_and(|cells| !cells.column_entries.is_empty());
        if nonempty {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    Ok(())
}

#[allow(
    deprecated,
    reason = "Numbers still roots its native calculation engine through the deprecated document field."
)]
fn rooted_formula_dependency_identifiers(source: &Package) -> Result<HashSet<u64>, Error> {
    let component = source
        .state
        .components
        .catalog()
        .get("Index/Document.iwa")
        .ok_or(Error::InvalidSource)?;
    let document_object = component
        .archive()
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(1))
        .ok_or(Error::InvalidSource)?;
    let (document_index, document_message) =
        unique_message_index(&document_object.messages, DOCUMENT_MESSAGE_TYPE)?
            .ok_or(Error::InvalidSource)?;
    validate_message_metadata(document_object, document_index)?;
    let document = tn::DocumentArchive::decode(document_message.data.as_slice())
        .map_err(|_error| Error::InvalidSource)?;
    let Some(calculation_engine) = document.calculation_engine.as_ref() else {
        return Ok(HashSet::new());
    };
    require_local_reference(
        singular_length_payload(&document_message.data, 3)?,
        calculation_engine.identifier,
    )?;
    require_declared_reference(
        document_object,
        document_index,
        calculation_engine.identifier,
        &[3],
    )?;
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, calculation_engine.identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource)?;
    let (engine_index, engine_message) =
        unique_message_index(resolved.messages, CALCULATION_ENGINE_MESSAGE_TYPE)?
            .ok_or(Error::InvalidSource)?;
    let engine_object = resolved_object(source, resolved)?;
    validate_message_metadata(engine_object, engine_index)?;
    let engine = tsce::CalculationEngineArchive::decode(engine_message.data.as_slice())
        .map_err(|_error| Error::InvalidSource)?;
    let references = engine.dependency_tracker.formula_owner_dependencies;
    let tracker = singular_length_payload(&engine_message.data, 2)?;
    let raw_references = repeated_length_payloads(tracker, 6)?;
    if raw_references.len() != references.len() {
        return Err(Error::InvalidSource);
    }
    let mut identifiers = HashSet::new();
    identifiers
        .try_reserve(references.len())
        .map_err(|_allocation| Error::Allocation {
            amount: references.len(),
        })?;
    for (reference, raw) in references.iter().zip(raw_references) {
        require_local_reference(raw, reference.identifier)?;
        require_declared_reference(engine_object, engine_index, reference.identifier, &[2, 6])?;
        if !identifiers.insert(reference.identifier) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(identifiers)
}

fn rooted_table_info_identifiers(source: &Package) -> Result<HashSet<u64>, Error> {
    let expected = source
        .document()
        .sheets()
        .iter()
        .fold(0usize, |count, sheet| {
            count.saturating_add(sheet.table_count())
        });
    let mut identifiers = HashSet::new();
    identifiers
        .try_reserve(expected)
        .map_err(|_allocation| Error::Allocation { amount: expected })?;
    for (sheet_position, semantic_sheet) in source.document().sheets().iter().enumerate() {
        let target = resolve_native_sheet(source, sheet_position)?;
        let object = source
            .state
            .components
            .catalog()
            .get_index(target.component_index)
            .and_then(|component| component.archive().objects.get(target.object_index))
            .ok_or(Error::InvalidSource)?;
        let decoded = super::decode_sheet_payload(
            object.messages.as_slice(),
            super::SemanticPath::Sheet {
                index: sheet_position,
            },
        )
        .map_err(map_read_error)?;
        let mut table_count = 0usize;
        for drawable in decoded.drawable_infos {
            let resolved = source
                .state
                .index
                .resolve_ref_id(&source.state.components, drawable.identifier)
                .map_err(map_read_error)?
                .ok_or(Error::InvalidSource)?;
            if unique_table_info(resolved)?.is_some() {
                table_count = table_count.saturating_add(1);
                if !identifiers.insert(drawable.identifier) {
                    return Err(Error::InvalidSource);
                }
            }
        }
        if table_count != semantic_sheet.table_count() {
            return Err(Error::InvalidSource);
        }
    }
    if identifiers.len() != expected {
        return Err(Error::InvalidSource);
    }
    Ok(identifiers)
}

fn validate_rooted_pivot_dependencies(
    source: &Package,
    operations: &[Operation],
) -> Result<(), Error> {
    let Some(path) = operations.iter().find_map(|operation| {
        (operation.before != operation.after)
            .then_some(operation.location)
            .and_then(|location| match location {
                Location::Table(_, _) => Some(location.path()),
                Location::Sheet(_) => None,
            })
    }) else {
        return Ok(());
    };
    for (sheet_position, sheet) in source.document().sheets().iter().enumerate() {
        for table_position in 0..sheet.table_count() {
            let target = resolve_native_table(source, sheet_position, table_position)?;
            let model = tst::TableModelArchive::decode(target.model_message.data.as_slice())
                .map_err(|_error| Error::InvalidSource)?;
            if model.pivot_owner.is_some() {
                return Err(Error::UnsupportedDependency { path });
            }
        }
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct SheetTarget<'a> {
    identifier: u64,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
    message: &'a RawMessage,
}

#[derive(Debug, Clone, Copy)]
#[allow(
    clippy::struct_field_names,
    reason = "The model prefix distinguishes every coordinate from the enclosing table-info owner."
)]
struct TableTarget<'a> {
    model_identifier: u64,
    model_component_index: usize,
    model_object_index: usize,
    model_message_index: usize,
    model_message_type: u32,
    model_message: &'a RawMessage,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct NativeOperation {
    location: Location,
    before: Arc<str>,
    after: Arc<str>,
    identifier: u64,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
}

fn resolve_native_operations(
    source: &Package,
    operations: &[Operation],
) -> Result<Vec<NativeOperation>, Error> {
    validate_root_document(source)?;
    let changed_count = operations
        .iter()
        .filter(|operation| operation.before != operation.after)
        .count();
    let mut native = Vec::new();
    native
        .try_reserve_exact(changed_count)
        .map_err(|_allocation| Error::Allocation {
            amount: changed_count,
        })?;
    for operation in operations
        .iter()
        .filter(|operation| operation.before != operation.after)
    {
        let resolved = match operation.location {
            Location::Sheet(position) => {
                let target = resolve_native_sheet(source, position)?;
                if decode_sheet_name(target.message_type, &target.message.data)?
                    != operation.before.as_ref()
                {
                    return Err(Error::InvalidSource);
                }
                NativeOperation {
                    location: operation.location,
                    before: Arc::clone(&operation.before),
                    after: Arc::clone(&operation.after),
                    identifier: target.identifier,
                    component_index: target.component_index,
                    object_index: target.object_index,
                    message_index: target.message_index,
                    message_type: target.message_type,
                }
            },
            Location::Table(sheet, table) => {
                let target = resolve_native_table(source, sheet, table)?;
                if decode_table_name(&target.model_message.data)? != operation.before.as_ref() {
                    return Err(Error::InvalidSource);
                }
                NativeOperation {
                    location: operation.location,
                    before: Arc::clone(&operation.before),
                    after: Arc::clone(&operation.after),
                    identifier: target.model_identifier,
                    component_index: target.model_component_index,
                    object_index: target.model_object_index,
                    message_index: target.model_message_index,
                    message_type: target.model_message_type,
                }
            },
        };
        native.push(resolved);
    }
    native.sort_unstable_by_key(|operation| {
        (
            operation.component_index,
            operation.object_index,
            operation.message_index,
        )
    });
    Ok(native)
}

fn validate_root_document(source: &Package) -> Result<(), Error> {
    let component = source
        .state
        .components
        .catalog()
        .get("Index/Document.iwa")
        .ok_or(Error::InvalidSource)?;
    let object = component.archive().object(1).ok_or(Error::InvalidSource)?;
    let (message_index, message) = unique_message_index(&object.messages, DOCUMENT_MESSAGE_TYPE)?
        .ok_or(Error::InvalidSource)?;
    validate_message_metadata(object, message_index)?;
    let document = tn::DocumentArchive::decode(message.data.as_slice())
        .map_err(|_error| Error::InvalidSource)?;
    let references = repeated_length_payloads(&message.data, 1)?;
    if references.len() != document.sheets.len() {
        return Err(Error::InvalidSource);
    }
    for (payload, reference) in references.into_iter().zip(document.sheets) {
        require_local_reference(payload, reference.identifier)?;
        require_declared_reference(object, message_index, reference.identifier, &[1])?;
    }
    Ok(())
}

fn resolve_native_sheet(source: &Package, sheet_position: usize) -> Result<SheetTarget<'_>, Error> {
    let document = Package::root_document(&source.state.components).map_err(map_read_error)?;
    let reference = document
        .sheets
        .get(sheet_position)
        .ok_or(Error::SheetNotFound)?;
    if reference.identifier == 0 || reference.deprecated_is_external == Some(true) {
        return Err(Error::InvalidSource);
    }
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, reference.identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource)?;
    let message_index = unique_sheet_message_index(resolved.messages)?;
    let message = resolved
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let object = resolved_object(source, resolved)?;
    validate_message_metadata(object, message_index)?;
    let decoded = super::decode_sheet_payload(
        resolved.messages,
        super::SemanticPath::Sheet {
            index: sheet_position,
        },
    )
    .map_err(map_read_error)?;
    let semantic = source
        .document()
        .sheets()
        .get(sheet_position)
        .ok_or(Error::SheetNotFound)?;
    if decoded.name != semantic.name()
        || decode_sheet_name(message.type_, &message.data)? != semantic.name()
    {
        return Err(Error::InvalidSource);
    }
    Ok(SheetTarget {
        identifier: reference.identifier,
        component_index: resolved.component_index,
        object_index: resolved.object_index,
        message_index,
        message_type: message.type_,
        message,
    })
}

fn resolve_native_table(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
) -> Result<TableTarget<'_>, Error> {
    let sheet = resolve_native_sheet(source, sheet_position)?;
    let decoded = super::decode_sheet_payload(
        source
            .state
            .components
            .catalog()
            .get_index(sheet.component_index)
            .and_then(|component| component.archive().objects.get(sheet.object_index))
            .ok_or(Error::InvalidSource)?
            .messages
            .as_slice(),
        super::SemanticPath::Sheet {
            index: sheet_position,
        },
    )
    .map_err(map_read_error)?;
    let drawable_payloads = sheet_drawable_payloads(sheet.message_type, &sheet.message.data)?;
    if drawable_payloads.len() != decoded.drawable_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut semantic_table_position = 0usize;
    for (drawable_position, drawable) in decoded.drawable_infos.iter().enumerate() {
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, drawable.identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource)?;
        let Some((info_message_index, info_message)) = unique_table_info(resolved)? else {
            continue;
        };
        let info = decode_table_info(&info_message.data)?;
        if semantic_table_position != table_position {
            semantic_table_position = semantic_table_position
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            continue;
        }
        require_local_reference(
            drawable_payloads
                .get(drawable_position)
                .copied()
                .ok_or(Error::InvalidSource)?,
            drawable.identifier,
        )?;
        let sheet_object = source
            .state
            .components
            .catalog()
            .get_index(sheet.component_index)
            .and_then(|component| component.archive().objects.get(sheet.object_index))
            .ok_or(Error::InvalidSource)?;
        let drawable_path: &[u32] = match sheet.message_type {
            SHEET_MESSAGE_TYPE => &[2],
            FORM_BASED_SHEET_MESSAGE_TYPE => &[1, 2],
            _ => return Err(Error::InvalidSource),
        };
        require_declared_reference(
            sheet_object,
            sheet.message_index,
            drawable.identifier,
            drawable_path,
        )?;
        let info_object = resolved_object(source, resolved)?;
        validate_message_metadata(info_object, info_message_index)?;
        let model_identifier = info.table_model().identifier().get();
        require_declared_reference(info_object, info_message_index, model_identifier, &[2])?;
        let model_reference = singular_length_payload(&info_message.data, 2)?;
        require_local_reference(model_reference, model_identifier)?;
        let model = source
            .state
            .index
            .resolve_ref_id(&source.state.components, model_identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource)?;
        let (model_message_index, model_message) = unique_table_model(model.messages)?;
        let model_object = resolved_object(source, model)?;
        validate_message_metadata(model_object, model_message_index)?;
        let semantic = source
            .document()
            .sheets()
            .get(sheet_position)
            .and_then(|semantic_sheet| semantic_sheet.tables().nth(table_position))
            .ok_or(Error::TableNotFound)?;
        if decode_table_name(&model_message.data)? != semantic.name() {
            return Err(Error::InvalidSource);
        }
        return Ok(TableTarget {
            model_identifier,
            model_component_index: model.component_index,
            model_object_index: model.object_index,
            model_message_index,
            model_message_type: model_message.type_,
            model_message,
        });
    }
    Err(Error::TableNotFound)
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
        .ok_or(Error::InvalidSource)
}

fn unique_sheet_message_index(messages: &[RawMessage]) -> Result<usize, Error> {
    let sheet = unique_message_index(messages, SHEET_MESSAGE_TYPE)?;
    let form = unique_message_index(messages, FORM_BASED_SHEET_MESSAGE_TYPE)?;
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

fn unique_table_model(messages: &[RawMessage]) -> Result<(usize, &RawMessage), Error> {
    let canonical = unique_message_index(messages, TABLE_MODEL_MESSAGE_TYPE)?;
    let legacy = unique_message_index(messages, LEGACY_TABLE_MODEL_MESSAGE_TYPE)?;
    match (canonical, legacy) {
        (Some(_), Some(_)) | (None, None) => Err(Error::InvalidSource),
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
        .filter(|(_index, message)| message.type_ == message_type);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(first)
}

fn validate_message_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
) -> Result<(), Error> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_
        != object
            .messages
            .get(message_index)
            .ok_or(Error::InvalidSource)?
            .type_
        || object.archive_info.should_merge == Some(true)
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

fn require_declared_reference(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    identifier: u64,
    accepted_path: &[u32],
) -> Result<(), Error> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_
        != object
            .messages
            .get(message_index)
            .ok_or(Error::InvalidSource)?
            .type_
        || message
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count()
            != 1
    {
        return Err(Error::InvalidSource);
    }
    let mut field_occurrence = false;
    for field in &message.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if count == 0 {
            continue;
        }
        if count != 1 || field_occurrence || field.path.as_slice() != accepted_path {
            return Err(Error::InvalidSource);
        }
        field_occurrence = true;
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
        .map_err(|_allocation| Error::Allocation { amount: count })?;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
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
        _ => Err(Error::InvalidSource),
    }
}

fn sheet_drawable_payloads(message_type: u32, source: &[u8]) -> Result<Vec<&[u8]>, Error> {
    match message_type {
        SHEET_MESSAGE_TYPE => repeated_length_payloads(source, 2),
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            let sheet = singular_length_payload(source, FORM_SHEET_SUPER_FIELD)?;
            repeated_length_payloads(sheet, 2)
        },
        _ => Err(Error::InvalidSource),
    }
}

fn require_local_reference(source: &[u8], expected_identifier: u64) -> Result<(), Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if identifier.is_some() || field.wire_type() != 0 {
                    return Err(Error::InvalidSource);
                }
                identifier = Some(canonical_varint(field.payload())?);
            },
            2 => {
                if deprecated_type.is_some() || field.wire_type() != 0 {
                    return Err(Error::InvalidSource);
                }
                let value = canonical_varint(field.payload())?;
                if value > u64::from(i32::MAX.unsigned_abs()) && value < MIN_SIGN_EXTENDED_I32 {
                    return Err(Error::InvalidSource);
                }
                deprecated_type = Some(value);
            },
            3 => {
                if external.is_some() || field.wire_type() != 0 {
                    return Err(Error::InvalidSource);
                }
                let value = canonical_varint(field.payload())?;
                if value > 1 {
                    return Err(Error::InvalidSource);
                }
                external = Some(value != 0);
            },
            _ => {},
        }
    }
    if identifier != Some(expected_identifier) || expected_identifier == 0 || external == Some(true)
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn canonical_varint(source: &[u8]) -> Result<u64, Error> {
    let (value, length) =
        decode_varint_from_bytes(source).map_err(|_error| Error::InvalidSource)?;
    if length != source.len() || encoded_len(value) != length {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

struct OwnedEntryEdit<'a> {
    name: &'a str,
    data: Vec<u8>,
}

fn rewrite_names(
    source: &Package,
    operations: &[NativeOperation],
    deleted_previews: &[&str],
) -> Result<(Package, usize), Error> {
    let source_catalog = physical_source(source)?;
    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut rewritten_entries = Vec::new();
    rewritten_entries
        .try_reserve_exact(component_group_count(operations))
        .map_err(|_allocation| Error::Allocation {
            amount: component_group_count(operations),
        })?;
    let mut begin = 0usize;
    while begin < operations.len() {
        let component_index = operations[begin].component_index;
        let mut end = begin + 1;
        while end < operations.len() && operations[end].component_index == component_index {
            end += 1;
        }
        let component = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::InvalidSource)?;
        let component_name = component.name();
        let entry = source_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == component_name)
            .ok_or(Error::InvalidSource)?;
        if entry.is_opaque() {
            return Err(Error::UnsupportedSource);
        }
        let stream = SnappyStream::decompress_with_limits(
            entry.data(),
            physical_limits.snappy_limits().map_err(map_archive_error)?,
        )
        .map_err(map_core_error)?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
        drop(stream);
        for operation in &operations[begin..end] {
            rewrite_native_operation(&mut archive, operation, archive_limits)?;
        }
        let rewritten = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        drop(archive);
        let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
        drop(rewritten);
        rewritten_entries.push(OwnedEntryEdit {
            name: component_name,
            data: compressed,
        });
        begin = end;
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(rewritten_entries.len())
        .map_err(|_allocation| Error::Allocation {
            amount: rewritten_entries.len(),
        })?;
    edits.extend(
        rewritten_entries
            .iter()
            .map(|entry| EntryEdit::new(entry.name, &entry.data)),
    );
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, deleted_previews, physical_limits)
        .map_err(map_archive_error)?;
    drop(edits);
    drop(rewritten_entries);
    let candidate = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(map_candidate_read_error)?;
    Ok((candidate, component_group_count(operations)))
}

fn verify_exact_locality(
    source: &Package,
    candidate: &Package,
    operations: &[NativeOperation],
    direction: Direction,
    preview_count: usize,
) -> Result<(), Error> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    let source_previews = root_preview_deletions(source_catalog)?;
    let candidate_previews = root_preview_deletions(candidate_catalog)?;
    let expected = match direction {
        Direction::Forward => (preview_count, 0),
        Direction::Reverse => (0, preview_count),
    };
    if (source_previews.len(), candidate_previews.len()) != expected {
        return Err(Error::Verification);
    }
    verify_package_members(
        source_catalog,
        candidate_catalog,
        operations,
        &source_previews,
        &candidate_previews,
    )?;
    verify_component_locality(source, candidate, operations, direction)
}

fn verify_package_members(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    operations: &[NativeOperation],
    source_previews: &[&str],
    candidate_previews: &[&str],
) -> Result<(), Error> {
    let mut before = source
        .package()
        .iter()
        .filter(|entry| !source_previews.contains(&entry.name()));
    let mut after = candidate
        .package()
        .iter()
        .filter(|entry| !candidate_previews.contains(&entry.name()));
    loop {
        match (before.next(), after.next()) {
            (Some(before), Some(after)) if before.name() == after.name() => {
                let selected = operations.iter().any(|operation| {
                    source
                        .components()
                        .get_index(operation.component_index)
                        .is_some_and(|component| component.name() == before.name())
                });
                let preserved = if selected {
                    selected_package_member_preserved(before, after)
                } else {
                    package_member_preserved(before, after)
                };
                if !preserved {
                    return Err(Error::Verification);
                }
            },
            (None, None) => return Ok(()),
            _ => return Err(Error::Verification),
        }
    }
}

fn package_member_preserved(source: &Entry, candidate: &Entry) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_package_member_preserved(source: &Entry, candidate: &Entry) -> bool {
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

fn central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= LOCAL_HEADER_OFFSET.end
        && source[..LOCAL_HEADER_OFFSET.start] == candidate[..LOCAL_HEADER_OFFSET.start]
        && source[LOCAL_HEADER_OFFSET.end..] == candidate[LOCAL_HEADER_OFFSET.end..]
}

fn selected_local_record_preserved(source: &Entry, candidate: &Entry) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
    let source_record = source.raw_record().local_record();
    let candidate_record = candidate.raw_record().local_record();
    let Some(source_header) = zip_local_header_length(source_record) else {
        return false;
    };
    let Some(candidate_header) = zip_local_header_length(candidate_record) else {
        return false;
    };
    if source_header != candidate_header
        || source_record[..CRC_AND_SIZES.start] != candidate_record[..CRC_AND_SIZES.start]
        || source_record[CRC_AND_SIZES.end..source_header]
            != candidate_record[CRC_AND_SIZES.end..candidate_header]
    {
        return false;
    }
    let Some(source_payload_end) = source_header
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= source_record.len())
    else {
        return false;
    };
    let Some(candidate_payload_end) = candidate_header
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= candidate_record.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &source_record[source_payload_end..],
        &candidate_record[candidate_payload_end..],
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
    let source_prefix = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let candidate_prefix = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    source_prefix == candidate_prefix
        && source.len() == candidate.len()
        && source.len() >= source_prefix + 12
        && source[..source_prefix] == candidate[..candidate_prefix]
        && source[source_prefix + 12..] == candidate[candidate_prefix + 12..]
}

fn selected_central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= LOCAL_HEADER_OFFSET.end
        && source[..CRC_AND_SIZES.start] == candidate[..CRC_AND_SIZES.start]
        && source[CRC_AND_SIZES.end..LOCAL_HEADER_OFFSET.start]
            == candidate[CRC_AND_SIZES.end..LOCAL_HEADER_OFFSET.start]
        && source[LOCAL_HEADER_OFFSET.end..] == candidate[LOCAL_HEADER_OFFSET.end..]
}

fn verify_component_locality(
    source: &Package,
    candidate: &Package,
    operations: &[NativeOperation],
    direction: Direction,
) -> Result<(), Error> {
    let mut begin = 0usize;
    while begin < operations.len() {
        let component_index = operations[begin].component_index;
        let mut end = begin + 1;
        while end < operations.len() && operations[end].component_index == component_index {
            end += 1;
        }
        let source_component = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::Verification)?;
        let candidate_component = candidate
            .state
            .components
            .catalog()
            .get(source_component.name())
            .ok_or(Error::Verification)?;
        let source_objects = &source_component.archive().objects;
        let candidate_objects = &candidate_component.archive().objects;
        if source_objects.len() != candidate_objects.len() {
            return Err(Error::Verification);
        }
        for (object_index, (before, after)) in
            source_objects.iter().zip(candidate_objects).enumerate()
        {
            let selected = &operations[begin..end];
            if !selected
                .iter()
                .any(|operation| operation.object_index == object_index)
            {
                if before.archive_info != after.archive_info || before.messages != after.messages {
                    return Err(Error::Verification);
                }
                continue;
            }
            verify_selected_object(before, after, object_index, selected, direction)?;
        }
        begin = end;
    }
    Ok(())
}

fn verify_selected_object(
    source: &litchi_iwa_core::ArchiveObject,
    candidate: &litchi_iwa_core::ArchiveObject,
    object_index: usize,
    operations: &[NativeOperation],
    direction: Direction,
) -> Result<(), Error> {
    if source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
    {
        return Err(Error::Verification);
    }
    for (message_index, (before, after)) in
        source.messages.iter().zip(&candidate.messages).enumerate()
    {
        let operation = operations.iter().find(|operation| {
            operation.object_index == object_index && operation.message_index == message_index
        });
        if let Some(operation) = operation {
            if before.type_ != after.type_
                || rewritten_operation_payload(before, operation, direction)? != after.data
                || !message_info_preserved_except_length(
                    source
                        .archive_info
                        .message_infos
                        .get(message_index)
                        .ok_or(Error::Verification)?,
                    candidate
                        .archive_info
                        .message_infos
                        .get(message_index)
                        .ok_or(Error::Verification)?,
                )
            {
                return Err(Error::Verification);
            }
        } else if before != after
            || source.archive_info.message_infos.get(message_index)
                != candidate.archive_info.message_infos.get(message_index)
        {
            return Err(Error::Verification);
        }
    }
    Ok(())
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

fn component_group_count(operations: &[NativeOperation]) -> usize {
    operations
        .iter()
        .enumerate()
        .filter(|(index, operation)| {
            *index == 0
                || operations
                    .get(index.saturating_sub(1))
                    .is_none_or(|previous| previous.component_index != operation.component_index)
        })
        .count()
}

fn rewrite_native_operation(
    archive: &mut Archive,
    operation: &NativeOperation,
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let object = archive
        .objects
        .get_mut(operation.object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(operation.identifier) {
        return Err(Error::InvalidSource);
    }
    validate_message_metadata(object, operation.message_index)?;
    let message = object
        .messages
        .get(operation.message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != operation.message_type {
        return Err(Error::InvalidSource);
    }
    let data = rewritten_operation_payload(message, operation, Direction::Forward)?;
    object
        .replace_message_preserving_header_with_limits(
            operation.message_index,
            RawMessage {
                type_: operation.message_type,
                data,
            },
            limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn rewritten_operation_payload(
    message: &RawMessage,
    operation: &NativeOperation,
    direction: Direction,
) -> Result<Vec<u8>, Error> {
    let (before, after) = match direction {
        Direction::Forward => (operation.before.as_ref(), operation.after.as_ref()),
        Direction::Reverse => (operation.after.as_ref(), operation.before.as_ref()),
    };
    match operation.location {
        Location::Sheet(_) => {
            if decode_sheet_name(message.type_, &message.data)? != before {
                return Err(Error::InvalidSource);
            }
            let data = match message.type_ {
                SHEET_MESSAGE_TYPE => patch_length_delimited_field(
                    &message.data,
                    SHEET_NAME_FIELD,
                    true,
                    Some(after.as_bytes()),
                )
                .map_err(map_wire_error)?,
                FORM_BASED_SHEET_MESSAGE_TYPE => {
                    transform_length_delimited_field::<_, litchi_iwa_common::Error>(
                        &message.data,
                        FORM_SHEET_SUPER_FIELD,
                        |sheet| {
                            patch_length_delimited_field(
                                sheet,
                                SHEET_NAME_FIELD,
                                true,
                                Some(after.as_bytes()),
                            )
                        },
                    )
                    .map_err(map_wire_error)?
                },
                _ => return Err(Error::InvalidSource),
            };
            if decode_sheet_name(message.type_, &data)? != after {
                return Err(Error::Verification);
            }
            Ok(data)
        },
        Location::Table(_, _) => {
            if decode_table_name(&message.data)? != before {
                return Err(Error::InvalidSource);
            }
            let data = patch_length_delimited_field(
                &message.data,
                TABLE_MODEL_NAME_FIELD,
                true,
                Some(after.as_bytes()),
            )
            .map_err(map_wire_error)?;
            if decode_table_name(&data)? != after {
                return Err(Error::Verification);
            }
            Ok(data)
        },
    }
}

fn root_preview_deletions(source: &SourceCatalog) -> Result<Vec<&'static str>, Error> {
    let mut names = Vec::new();
    names
        .try_reserve_exact(ROOT_PREVIEWS.len())
        .map_err(|_allocation| Error::Allocation {
            amount: ROOT_PREVIEWS.len(),
        })?;
    for preview in ROOT_PREVIEWS {
        let count = source
            .package()
            .iter()
            .filter(|entry| entry.name() == preview)
            .count();
        match count {
            0 => {},
            1 => names.push(preview),
            _ => return Err(Error::InvalidSource),
        }
    }
    Ok(names)
}

fn names_decode_options(source: &[u8]) -> numbers_names_codec::DecodeOptions {
    numbers_names_codec::DecodeOptions::new(
        source.len().max(1),
        source.len().max(1),
        source.len().saturating_mul(4).max(1),
        2,
    )
}

fn decode_sheet_name(message_type: u32, source: &[u8]) -> Result<&str, Error> {
    let options = names_decode_options(source);
    let snapshot = match message_type {
        SHEET_MESSAGE_TYPE => numbers_names_codec::decode_sheet_name(source, options),
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            numbers_names_codec::decode_form_sheet_name(source, options)
        },
        _ => return Err(Error::InvalidSource),
    }
    .map_err(map_names_codec_error)?;
    Ok(snapshot.name())
}

fn decode_table_name(source: &[u8]) -> Result<&str, Error> {
    numbers_names_codec::decode_table_names(source, names_decode_options(source))
        .map(numbers_names_codec::TableNamesSnapshot::table_name)
        .map_err(map_names_codec_error)
}

fn decode_table_info(source: &[u8]) -> Result<table_info_codec::TableInfoSnapshot, Error> {
    table_info_codec::decode_table_info(source, table_info_decode_options(source))
        .map_err(map_table_info_error)
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
        let framed = header_bytes
            .checked_add(u64::try_from(prefix_bytes).map_err(|_error| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        if framed != object.header_length
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
    package
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

fn map_names_codec_error(error: numbers_names_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        };
    }
    match error.wire_resource_limit() {
        Some(numbers_names_codec::WireResourceLimit::Bytes { observed, maximum }) => {
            Error::LimitExceeded {
                kind: LimitKind::WireBytes,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(maximum),
            }
        },
        Some(numbers_names_codec::WireResourceLimit::Nesting { observed, maximum }) => {
            Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        None | Some(_) => Error::InvalidSource,
    }
}

fn map_table_info_error(error: table_info_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: usize_as_u64(observed),
            maximum: usize_as_u64(maximum),
        };
    }
    match error.wire_resource_limit() {
        Some(table_info_codec::WireResourceLimit::Bytes {
            observed,
            maximum: Some(maximum),
        }) => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: usize_as_u64(observed.unwrap_or_else(|| maximum.saturating_add(1))),
            maximum: usize_as_u64(maximum),
        },
        Some(table_info_codec::WireResourceLimit::Nesting {
            observed,
            maximum: Some(maximum),
        }) => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed.unwrap_or_else(|| maximum.saturating_add(1))),
            maximum: u64::from(maximum),
        },
        _ => Error::InvalidSource,
    }
}

fn map_lock_error(error: super::TableLockError) -> Error {
    match error {
        super::TableLockError::SheetNotFound => Error::SheetNotFound,
        super::TableLockError::TableNotFound => Error::TableNotFound,
        super::TableLockError::UnsupportedSource => Error::UnsupportedSource,
        super::TableLockError::LimitExceeded {
            kind: _,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: LimitKind::PayloadItems,
            observed,
            maximum,
        },
        super::TableLockError::Allocation { amount } => Error::Allocation { amount },
        super::TableLockError::InvalidSource
        | super::TableLockError::Verification
        | super::TableLockError::PatchConflict => Error::InvalidSource,
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
        },
        other => map_read_error(other),
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::Archive(error) => map_archive_error(error),
        ReadError::Common(error) => map_wire_error(error),
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::PayloadObjects,
                super::SemanticLimitKind::References => LimitKind::PayloadReferences,
                super::SemanticLimitKind::OutputTextBytes
                | super::SemanticLimitKind::FormulaWireBytes
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
        ReadError::InputTooLarge { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed,
            maximum,
        },
        ReadError::Io(_)
        | ReadError::MalformedPayload { .. }
        | ReadError::NotNumbers
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
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Reassembly(_) => Error::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::InvalidBundle(_) => Error::InvalidSource,
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
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => Error::InvalidSource,
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
                litchi_iwa_common::LimitKind::InputBytes
                | litchi_iwa_common::LimitKind::OutputBytes => LimitKind::WireBytes,
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
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => Error::InvalidSource,
    }
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
