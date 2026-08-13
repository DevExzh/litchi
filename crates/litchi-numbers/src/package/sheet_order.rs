//! Implementation of the public [`crate::sheet::order`] sheet-order API.
//!
//! The public module intentionally exposes semantic selectors, positions, and
//! content-free diagnostics only. This implementation keeps retained artifact
//! identity and format details private.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "the focused boundary redacts lower-layer physical error families"
)]

mod error;
mod resolve;
mod rewrite;

use std::{fmt, sync::Arc};

use litchi_iwa_archive::package::OwnedExactArtifacts;
use thiserror::Error as ThisError;

use super::Package;
use crate::selector::SheetSelector;

use error::{map_candidate_read_error, map_read_error};
use resolve::{NativeTarget, TransactionBudget, resolve_native_target};
use rewrite::{
    ReopenCost, physical_source, preflight_transaction_work, root_preview_deletions,
    verify_exact_locality,
};

const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// A finite resource governed by a sheet-order transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source artifact bytes.
    InputBytes,
    /// Complete candidate artifact bytes.
    OutputBytes,
    /// Retained artifact entries.
    Entries,
    /// Bytes in one retained entry.
    EntryBytes,
    /// Aggregate retained-entry bytes.
    TotalEntryBytes,
    /// Retained artifact naming and metadata bytes.
    PackageBytes,
    /// Bytes in one decoded record container.
    PayloadBytes,
    /// Aggregate decoded-record bytes.
    TotalPayloadBytes,
    /// Decoded records inspected.
    PayloadObjects,
    /// Decoded message records inspected.
    PayloadMessages,
    /// Aggregate decoded metadata items.
    PayloadItems,
    /// Rooted semantic references inspected.
    PayloadReferences,
    /// Bytes inspected by the strict format parser.
    WireBytes,
    /// Bytes emitted by the focused rewrite.
    WireOutputBytes,
    /// Encoded fields inspected.
    WireFields,
    /// Nested encoded depth.
    WireNesting,
    /// Aggregate format traversal and rewrite work.
    WireWork,
    /// Aggregate focused transaction work.
    TransactionWork,
    /// Rooted semantic sheets.
    Sheets,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "package entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PackageBytes => "package metadata bytes",
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
            Self::Sheets => "rooted sheets",
        })
    }
}

/// A typed, content-redacted sheet-order transaction failure.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The semantic selector did not resolve one source sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The destination is outside the existing sheet sequence.
    #[error("destination position {position} is outside {sheet_count} rooted sheets")]
    DestinationOutOfRange {
        /// Requested zero-based destination.
        position: usize,
        /// Number of source sheets.
        sheet_count: usize,
    },
    /// This edit already contains its single permitted move.
    #[error("a sheet-order operation is already staged")]
    OperationAlreadyStaged,
    /// Commit was requested before staging a move.
    #[error("no sheet-order operation is staged")]
    NoStagedOperation,
    /// An exact source artifact or its required preview profile is unavailable
    /// for a changed move.
    #[error("source representation does not support changed sheet ordering")]
    UnsupportedSource,
    /// The retained source cannot support an unambiguous sheet-order move.
    #[error("rooted sheet-order source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Rejected observed amount.
        observed: u64,
        /// Retained maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} sheet-order items or bytes")]
    Allocation {
        /// Requested capacity or byte count.
        amount: usize,
    },
    /// The validated candidate did not reproduce the requested move exactly.
    #[error("sheet-order candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact retained source.
    #[error("sheet-order patch conflicts with this package")]
    PatchConflict,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Operation {
    source_position: usize,
    destination_position: usize,
}

/// One immutable edit that can stage exactly one sheet move.
pub struct Edit<'a> {
    source: &'a Package,
    operation: Option<Operation>,
    target: Option<NativeTarget>,
    budget: Option<TransactionBudget>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("operation", &self.operation)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage the selected sheet at one existing zero-based final destination.
    ///
    /// The selector resolves in the immutable source snapshot. The destination
    /// is interpreted after removing the selected sheet, as a position in the
    /// final sequence. This edit accepts exactly one move. Equal source and
    /// destination positions stage an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns [`Error::SheetNotFound`] when the selector does not resolve,
    /// [`Error::DestinationOutOfRange`] when `destination` is not an existing
    /// source position, and [`Error::OperationAlreadyStaged`] after one move.
    /// A changed move can also return content-redacted source, allocation, or
    /// resource errors. Failure leaves the edit unchanged.
    ///
    /// # Costs
    ///
    /// Selector resolution is linear in source sheets. Staging an equal
    /// positional no-op does not inspect retained format records. A changed
    /// move performs bounded target validation; publication work occurs only
    /// in [`Self::commit`].
    pub fn move_sheet<'selector>(
        mut self,
        selector: impl Into<SheetSelector<'selector>>,
        destination: usize,
    ) -> Result<Self, Error> {
        if self.operation.is_some() {
            return Err(Error::OperationAlreadyStaged);
        }
        let selected = self
            .source
            .state
            .document
            .sheet(selector)
            .map_err(|semantic_error| map_read_error(super::Error::Semantic(semantic_error)))?
            .ok_or(Error::SheetNotFound)?;
        let sheet_count = self.source.state.document.sheet_count();
        if destination >= sheet_count {
            return Err(Error::DestinationOutOfRange {
                position: destination,
                sheet_count,
            });
        }
        let operation = Operation {
            source_position: selected.index(),
            destination_position: destination,
        };
        let target = if operation.source_position == operation.destination_position {
            None
        } else {
            let mut budget = TransactionBudget::new(self.source);
            let target = resolve_native_target(self.source, &mut budget)?;
            self.budget = Some(budget);
            Some(target)
        };
        self.operation = Some(operation);
        self.target = target;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// A changed move requires exactly the three canonical root preview assets,
    /// updates both order owners together, removes those assets (three to
    /// zero), and preserves all unrelated retained data. An equal-position
    /// move shares the source artifact unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NoStagedOperation`] for an empty edit. A changed move
    /// returns [`Error::UnsupportedSource`] if the exact source or any required
    /// canonical preview asset is unavailable; it can also return
    /// content-redacted resource, allocation, or verification errors. Failure
    /// never publishes a partial package.
    ///
    /// # Costs
    ///
    /// A positional no-op shares the source artifact and reads it once for a
    /// diagnostic fingerprint. A changed move validates the complete
    /// three-preview profile, rewrites the two order owners, removes the three
    /// previews, reassembles once, and fully reopens one candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self.operation.ok_or(Error::NoStagedOperation)?;
        let source_catalog = physical_source(self.source)?;
        let source_bytes = source_catalog.__source_owner();
        if operation.source_position == operation.destination_position {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: OwnedExactArtifacts::new(source_bytes.clone(), source_bytes),
                    operation,
                    moved_sheet_identifier: None,
                    source_previews: 0,
                    target_previews: 0,
                    source_target: None,
                    target_target: None,
                    source_reopen: ReopenCost::default(),
                    target_reopen: ReopenCost::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !source_catalog.source_is_exact() {
            return Err(Error::UnsupportedSource);
        }
        let mut budget = self.budget.ok_or(Error::InvalidSource)?;
        preflight_transaction_work(self.source, None, &mut budget)?;
        let target = self.target.ok_or(Error::InvalidSource)?;
        let previews = root_preview_deletions(source_catalog)?;
        if previews.len() != ROOT_PREVIEWS.len() {
            return Err(Error::UnsupportedSource);
        }
        let rewritten = rewrite::rewrite(
            self.source,
            &target,
            &mut budget,
            operation.source_position,
            operation.destination_position,
            &previews,
        )?;
        verify_exact_locality(
            self.source,
            &rewritten.package,
            &target,
            &rewritten.target,
            &mut budget,
            operation.source_position,
            operation.destination_position,
            &previews,
            0,
        )?;
        let target_bytes = physical_source(&rewritten.package)?.__source_owner();
        budget.charge_transaction_work(
            source_bytes
                .len()
                .checked_add(target_bytes.len())
                .ok_or(Error::InvalidSource)?,
        )?;
        let source_target = Arc::new(target);
        let target_target = Arc::new(rewritten.target);
        let moved_sheet_identifier = source_target
            .sheet_identifier(operation.source_position)
            .ok_or(Error::InvalidSource)?;
        Ok(Commit {
            package: rewritten.package,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_bytes, target_bytes.clone()),
                operation,
                source_target: Some(source_target),
                target_target: Some(target_target),
                moved_sheet_identifier: Some(moved_sheet_identifier),
                source_previews: previews.len(),
                target_previews: 0,
                source_reopen: rewritten.source_reopen,
                target_reopen: rewritten.target_reopen,
            },
            diagnostics: Diagnostics::published(previews.len()),
        })
    }
}

/// A reversible, process-local exact-source sheet-order patch.
///
/// A patch authorizes application only to its exact retained source artifact;
/// fingerprints are diagnostic values and never authorize a write. It exposes
/// positions but no retained content or implementation identifiers.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    operation: Operation,
    source_target: Option<Arc<NativeTarget>>,
    target_target: Option<Arc<NativeTarget>>,
    moved_sheet_identifier: Option<u64>,
    source_previews: usize,
    target_previews: usize,
    source_reopen: ReopenCost,
    target_reopen: ReopenCost,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_position", &self.operation.source_position)
            .field("destination_position", &self.operation.destination_position)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected sheet's zero-based source position.
    #[must_use]
    pub const fn source_position(&self) -> usize {
        self.operation.source_position
    }

    /// Return the zero-based destination position in the final sequence.
    ///
    /// This is the position after removing the selected source sheet, not a
    /// pre-removal insertion offset.
    #[must_use]
    pub const fn destination_position(&self) -> usize {
        self.operation.destination_position
    }

    /// Return the source artifact's non-authorizing diagnostic fingerprint.
    ///
    /// This value is suitable only for observability; exact retained artifact
    /// identity authorizes patch application.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target artifact's non-authorizing diagnostic fingerprint.
    ///
    /// This value is suitable only for observability; exact retained artifact
    /// identity authorizes patch application.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether the staged move is an exact artifact no-op.
    ///
    /// This is true only when source and final destination positions are equal
    /// and the retained source and target artifacts are exactly unchanged.
    ///
    /// # Costs
    ///
    /// Uses allocation identity first, then at most one complete byte compare.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.operation.source_position == self.operation.destination_position
            && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse patch.
    ///
    /// Apply the returned patch to this patch's resulting package to restore
    /// the original artifact byte-for-byte. For a changed patch, the forward
    /// direction removes all three canonical previews (three to zero); this
    /// inverse restores them (zero to three).
    ///
    /// # Costs
    ///
    /// Swaps shared handles and compact metadata in `O(1)` time.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            operation: Operation {
                source_position: self.operation.destination_position,
                destination_position: self.operation.source_position,
            },
            source_target: self.target_target.as_ref().map(Arc::clone),
            target_target: self.source_target.as_ref().map(Arc::clone),
            moved_sheet_identifier: self.moved_sheet_identifier,
            source_previews: self.target_previews,
            target_previews: self.source_previews,
            source_reopen: self.target_reopen,
            target_reopen: self.source_reopen,
        }
    }
}

/// Content-free diagnostics for one completed sheet-order transaction.
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

    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten retained components.
    ///
    /// A changed sheet-order move reports one component: it contains the two
    /// order owners updated together. Preview deletion is reported separately
    /// by [`Self::deleted_previews`].
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of obsolete root preview assets deleted in this direction.
    ///
    /// A forward changed move deletes exactly three; a no-op deletes zero. An
    /// inverse changed move restores those three and therefore reports zero
    /// deletions in its direction.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether a complete candidate artifact was reopened for validation.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One successfully validated immutable publication.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    ///
    /// Call [`Package::write_to`] on this snapshot to stream its exact retained
    /// artifact. Sink durability and destination replacement remain the
    /// caller's policy.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free transaction diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start an immutable single-operation sheet-order edit.
    ///
    /// The resulting [`Edit`] accepts one semantic selector and one final
    /// zero-based destination. The public transaction types live in
    /// [`crate::sheet::order`], not at this crate's root.
    ///
    /// # Costs
    ///
    /// This only borrows the package; no artifact bytes are copied or parsed.
    #[must_use]
    pub const fn edit_sheet_order(&self) -> Edit<'_> {
        Edit {
            source: self,
            operation: None,
            target: None,
            budget: None,
        }
    }

    /// Apply a reversible exact-source sheet-order patch.
    ///
    /// Application is exact-source only: the patch must be applied to the
    /// immutable artifact from which it was created. A changed forward patch
    /// publishes its already retained target after checking the two order
    /// owners and the canonical preview transition from three to zero. Its
    /// inverse checks and restores the reverse zero-to-three transition. Apply
    /// [`Patch::inverse`] to the resulting package to recover the original
    /// artifact byte-for-byte.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the retained
    /// exact source. A changed patch can also return content-redacted resource,
    /// source, allocation, or verification errors. Failure never publishes a
    /// partial package.
    ///
    /// # Costs
    ///
    /// A no-op shares this snapshot. A changed patch reopens one retained
    /// artifact and performs one linear locality verification.
    pub fn apply_sheet_order(&self, patch: &Patch) -> Result<Commit, Error> {
        let source_catalog = physical_source(self)?;
        let source = source_catalog.__source_owner();
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
        if !source_catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let source_target = patch.source_target.as_deref().ok_or(Error::PatchConflict)?;
        let expected_target = patch.target_target.as_deref().ok_or(Error::PatchConflict)?;
        if !matches!(
            (patch.source_previews, patch.target_previews),
            (3, 0) | (0, 3)
        ) {
            return Err(Error::PatchConflict);
        }
        let mut budget = TransactionBudget::new(self);
        if source_target.sheet_identifier(patch.operation.source_position)
            != patch.moved_sheet_identifier
        {
            return Err(Error::PatchConflict);
        }
        let target_bytes = patch.artifacts.target_owner();
        preflight_transaction_work(self, Some(target_bytes.as_ref()), &mut budget)?;
        budget.charge_transaction_work(patch.target_reopen.work)?;
        budget.charge_references(patch.target_reopen.references)?;
        let candidate = Package::from_source_owner_with_options(target_bytes, self.state.options)
            .map_err(map_candidate_read_error)?;
        let candidate_target = resolve_native_target(&candidate, &mut budget)?;
        if candidate_target != *expected_target
            || candidate_target.sheet_identifier(patch.operation.destination_position)
                != patch.moved_sheet_identifier
        {
            return Err(Error::Verification);
        }
        let source_previews = root_preview_deletions(source_catalog)?;
        if source_previews.len() != patch.source_previews {
            return Err(Error::PatchConflict);
        }
        verify_exact_locality(
            self,
            &candidate,
            source_target,
            &candidate_target,
            &mut budget,
            patch.operation.source_position,
            patch.operation.destination_position,
            &source_previews,
            patch.target_previews,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.source_previews.saturating_sub(patch.target_previews),
            ),
        })
    }
}
