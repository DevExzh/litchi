//! Exact-source per-slide placeholder-visibility transactions.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "the focused transaction maps non-exhaustive lower-layer errors into a content-redacted boundary"
)]

mod errors;
mod resolve;
mod rewrite;
mod slide_number;
mod verification;

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    wire::{WireFieldView, WireView},
};
use litchi_iwa_core::{ArchiveObject, MessageInfo, SnappyStream};
use litchi_iwa_protos::keynote_placeholder_text_codec;
use thiserror::Error as ThisError;

use super::{
    PLACEHOLDER_MESSAGE_TYPE, Package, PhysicalSource, ReadError, SLIDE_MESSAGE_TYPE,
    SLIDE_NODE_MESSAGE_TYPE,
};
use crate::{
    SlideSelector,
    slide::placeholder::{Kind, State},
};

use errors::{
    map_archive_error, map_core_error, map_placeholder_error, map_read_error,
    map_rendering_invalidation_error, map_slide_preview_error, map_wire_error, physical_catalog,
    placeholder_options, root_preview_count,
};
use resolve::focused_slide;
use rewrite::rewrite_visibility;
use slide_number::{
    PLACEHOLDER_FIELD as SLIDE_NUMBER_PLACEHOLDER_FIELD, node_visible as slide_number_node_visible,
    placeholder_owner as slide_number_placeholder_owner,
    validate_global_ownership as validate_global_slide_number_ownership,
    validate_node_references as validate_slide_number_node_references,
    validate_storage as validate_slide_number_storage,
};
use verification::verify_artifact_delta;

const TITLE_REFERENCE_FIELD: u32 = 5;
const BODY_REFERENCE_FIELD: u32 = 6;
const OWNED_DRAWABLES_FIELD: u32 = 7;
const Z_ORDER_FIELD: u32 = 42;
const OBJECT_PLACEHOLDER_FIELD: u32 = 30;
const SLIDE_STYLE_FIELD: u32 = 1;
const SLIDE_BUILDS_FIELD: u32 = 2;
const SLIDE_LAYERING_FIELD: u32 = 41;
const SLIDE_TITLE_CACHE_FIELD: u32 = 37;
const SLIDE_BODY_CACHE_FIELD: u32 = 38;
const SLIDE_STYLE_MESSAGE_TYPE: u32 = 9;
const SLIDE_TEMPLATE_REFERENCE_FIELD: u32 = 17;
const SLIDE_KNOWN_REFERENCE_FIELDS: [u32; 8] = [27, 29, 31, 35, 36, 39, 43, 44];

/// A finite resource governed by a placeholder-visibility transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes consumed while opening the source package.
    InputBytes,
    /// Bytes produced for the resulting package.
    OutputBytes,
    /// Package entries.
    Entries,
    /// Bytes in one package entry.
    EntryBytes,
    /// Aggregate bytes across package entries.
    TotalBytes,
    /// Slides visited while resolving a selector.
    Slides,
    /// Relationships visited while proving the selected state is safe.
    References,
    /// Bytes in selected encoded data.
    WireBytes,
    /// Fields visited in selected encoded data.
    WireFields,
    /// Nested encoded-data depth.
    WireNesting,
    /// Aggregate scanning and update work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// A content-redacted placeholder-visibility failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The source cannot support a changed exact-source visibility transaction.
    #[error("this Keynote source does not support placeholder-visibility edits")]
    UnsupportedSource,
    /// More than one rooted slide has the requested exact name.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// No rooted slide has the requested exact name.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// The requested zero-based position is outside the rooted slide list.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Requested zero-based semantic position.
        position: Position,
    },
    /// The selected slide has no existing placeholder for the requested role.
    #[error("the selected Keynote slide has no existing {kind} placeholder")]
    PlaceholderNotFound {
        /// Missing semantic role.
        kind: Kind,
    },
    /// The selected presentation state cannot safely support the requested change.
    #[error("the Keynote placeholder visibility cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Keynote placeholder-visibility {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource whose ceiling was exceeded.
        kind: LimitKind,
        /// Exact observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Keynote placeholder-visibility transaction")]
    Allocation {
        /// Requested elements or bytes.
        amount: usize,
    },
    /// Candidate semantic or rendering verification failed.
    #[error("the edited Keynote placeholder visibility failed verification")]
    Verification,
    /// A process-local patch was applied to a different exact package snapshot.
    #[error("the Keynote placeholder-visibility patch does not match the exact source package")]
    PatchConflict,
}

/// One staged visibility value against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    position: Position,
    kind: Kind,
    placeholder_identifier: u64,
    before: State,
    state: State,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.position)
            .field("kind", &self.kind)
            .field("before", &self.before)
            .field("state", &self.state)
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn new(source: &'a Package, selection: Selection<'_>, kind: Kind) -> Self {
        Self {
            source,
            position: selection.position,
            kind,
            placeholder_identifier: selection.placeholder_identifier,
            before: selection.state,
            state: selection.state,
        }
    }

    #[must_use]
    /// Return the selected zero-based slide position.
    ///
    /// This is a constant-time accessor.
    pub const fn position(&self) -> Position {
        self.position
    }
    #[must_use]
    /// Return the selected placeholder role.
    ///
    /// This is a constant-time accessor.
    pub const fn kind(&self) -> Kind {
        self.kind
    }
    #[must_use]
    /// Return the currently staged visibility.
    ///
    /// This is a constant-time accessor.
    pub const fn state(&self) -> State {
        self.state
    }
    #[must_use]
    /// Stage an explicit visibility value without publishing it.
    ///
    /// This consumes the edit, is allocation-free, and never creates or
    /// deletes a placeholder. It affects only the already selected slide and
    /// role.
    pub const fn set(mut self, state: State) -> Self {
        self.state = state;
        self
    }
    #[must_use]
    /// Stage [`State::Visible`] without publishing it.
    ///
    /// This has the same constant-time, allocation-free cost as [`Self::set`].
    pub const fn show(self) -> Self {
        self.set(State::Visible)
    }
    #[must_use]
    /// Stage [`State::Hidden`] without publishing it.
    ///
    /// This has the same constant-time, allocation-free cost as [`Self::set`].
    pub const fn hide(self) -> Self {
        self.set(State::Hidden)
    }

    /// Publish the staged visibility atomically.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedSource`] when a changed exact-source
    /// publication is unavailable, or a typed source, resource, allocation, or
    /// verification error. Failures do not publish a partial package.
    ///
    /// # Costs
    ///
    /// An exact semantic no-op returns the existing immutable package snapshot
    /// and performs no publication or candidate reopen. A changed commit
    /// validates the selected slide's supported closure, updates only the
    /// affected presentation state, invalidates derived rendering state,
    /// removes stale package-root previews, and reopens one complete candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let catalog = physical_catalog(self.source)?;
        let source = catalog.shared_source();
        if self.before == self.state {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source), source),
                    position: self.position,
                    kind: self.kind,
                    placeholder_identifier: self.placeholder_identifier,
                    before: self.before,
                    after: self.state,
                    touched_components: 0,
                    source_preview_count: 0,
                    target_preview_count: 0,
                    source_node_invalidated: false,
                    target_node_invalidated: false,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::UnsupportedSource);
        }
        let mut budget = TransactionBudget::new(self.source)?;
        let selection = select_with_budget(
            self.source,
            SlideSelector::Position(self.position),
            self.kind,
            true,
            &mut budget,
        )?
        .ok_or(Error::InvalidSource)?;
        if selection.placeholder_identifier != self.placeholder_identifier
            || selection.state != self.before
        {
            return Err(Error::InvalidSource);
        }
        let source_preview_count = root_preview_count(self.source)?;
        let (package, touched_components, source_node_invalidated) =
            rewrite_visibility(self.source, &selection, self.state, &mut budget)?;
        let target = physical_catalog(&package)?.shared_source();
        Ok(Commit {
            patch: Patch {
                artifacts: ExactArtifacts::new(source, Arc::clone(&target)),
                position: self.position,
                kind: self.kind,
                placeholder_identifier: self.placeholder_identifier,
                before: self.before,
                after: self.state,
                touched_components,
                source_preview_count,
                target_preview_count: 0,
                source_node_invalidated,
                target_node_invalidated: if self.kind == Kind::SlideNumber {
                    self.state == State::Visible
                } else {
                    true
                },
            },
            package,
            diagnostics: Diagnostics::published(touched_components, source_preview_count),
        })
    }
}

/// A process-local exact-source reversible visibility patch.
///
/// A patch is not a serialized interchange format. It is applicable only to
/// the exact package snapshot from which it was committed.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    position: Position,
    kind: Kind,
    placeholder_identifier: u64,
    before: State,
    after: State,
    touched_components: usize,
    source_preview_count: usize,
    target_preview_count: usize,
    source_node_invalidated: bool,
    target_node_invalidated: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .field("kind", &self.kind)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    #[must_use]
    /// Return the selected zero-based slide position.
    pub const fn position(&self) -> Position {
        self.position
    }
    #[must_use]
    /// Return the selected placeholder role.
    pub const fn kind(&self) -> Kind {
        self.kind
    }
    #[must_use]
    /// Return the exact source semantic state.
    pub const fn before(&self) -> State {
        self.before
    }
    #[must_use]
    /// Return the exact target semantic state.
    pub const fn after(&self) -> State {
        self.after
    }
    #[must_use]
    /// Return a diagnostic fingerprint of the retained source package snapshot.
    ///
    /// Fingerprints never authorize application; the retained exact source
    /// snapshot does.
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }
    #[must_use]
    /// Return a diagnostic fingerprint of the retained target package snapshot.
    ///
    /// This value is process-local diagnostic evidence, not a stable package
    /// identity or patch-application authority.
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
    #[must_use]
    /// Return whether semantic state and retained package snapshots are both exact no-ops.
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }
    #[must_use]
    /// Build the exact inverse in `O(1)` by swapping retained package snapshots.
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            position: self.position,
            kind: self.kind,
            placeholder_identifier: self.placeholder_identifier,
            before: self.after,
            after: self.before,
            touched_components: self.touched_components,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
            source_node_invalidated: self.target_node_invalidated,
            target_node_invalidated: self.source_node_invalidated,
        }
    }
}

/// Compact evidence for one visibility commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
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
    #[must_use]
    /// Return whether publication changed the exact package snapshot.
    pub const fn changed(self) -> bool {
        self.changed
    }
    #[must_use]
    /// Return how many distinct underlying components changed.
    ///
    /// Removing stale package-root previews is not included in this count.
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    #[must_use]
    /// Return how many stale package-root previews were deleted directionally.
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
    #[must_use]
    /// Return whether a changed candidate was reopened for verification.
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one visibility transaction.
#[must_use]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}
impl Commit {
    #[must_use]
    /// Borrow the verified resulting package.
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    /// Consume the commit and return its resulting package.
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    /// Borrow the process-local exact-source patch.
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    #[must_use]
    /// Borrow compact transaction diagnostics.
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

pub(super) struct Selection<'a> {
    position: Position,
    kind: Kind,
    node_identifier: u64,
    slide_identifier: u64,
    placeholder_identifier: u64,
    slide_component: &'a str,
    node_component: &'a str,
    slide_payload: &'a [u8],
    node_payload: &'a [u8],
    state: State,
}

pub(super) struct TransactionBudget {
    references: usize,
    fields: usize,
    work: usize,
    maximum_references: usize,
    maximum_fields: usize,
    maximum_work: usize,
}

#[derive(Clone, Copy)]
enum PlaceholderDependencyPath {
    ShapeStyle,
    TextFlow,
    Comment,
    PencilAnnotation,
    Title,
    Caption,
}

struct PlaceholderDependencyDeclaration {
    path: PlaceholderDependencyPath,
    aggregate_occurrences: usize,
    field_attributed: bool,
}

impl PlaceholderDependencyPath {
    const fn from_path(path: &[u32]) -> Option<Self> {
        match path {
            [1, 1, 2] => Some(Self::ShapeStyle),
            [1, 3] => Some(Self::TextFlow),
            [1, 1, 1, 6] => Some(Self::Comment),
            [1, 1, 1, 9] => Some(Self::PencilAnnotation),
            [1, 1, 1, 10] => Some(Self::Title),
            [1, 1, 1, 11] => Some(Self::Caption),
            _ => None,
        }
    }

    const fn as_slice(self) -> &'static [u32] {
        match self {
            Self::ShapeStyle => &[1, 1, 2],
            Self::TextFlow => &[1, 3],
            Self::Comment => &[1, 1, 1, 6],
            Self::PencilAnnotation => &[1, 1, 1, 9],
            Self::Title => &[1, 1, 1, 10],
            Self::Caption => &[1, 1, 1, 11],
        }
    }
}

impl TransactionBudget {
    fn new(package: &Package) -> Result<Self, Error> {
        let wire_limits = package.wire_limits().map_err(map_wire_error)?;
        Ok(Self {
            references: 0,
            fields: 0,
            work: 0,
            maximum_references: package.semantic_limits().max_references(),
            maximum_fields: wire_limits.max_fields(),
            maximum_work: wire_limits.max_rewrite_work(),
        })
    }

    pub(super) fn charge_reference(&mut self) -> Result<(), Error> {
        self.charge_references(1)
    }

    pub(super) fn charge_references(&mut self, amount: usize) -> Result<(), Error> {
        self.references = self
            .references
            .checked_add(amount)
            .ok_or(Error::InvalidSource)?;
        if self.references > self.maximum_references {
            return Err(Error::LimitExceeded {
                kind: LimitKind::References,
                observed: self.references as u64,
                maximum: self.maximum_references as u64,
            });
        }
        Ok(())
    }

    pub(super) fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        self.work = self.work.checked_add(amount).ok_or(Error::InvalidSource)?;
        if self.work > self.maximum_work {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed: self.work as u64,
                maximum: self.maximum_work as u64,
            });
        }
        Ok(())
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), Error> {
        self.fields = self
            .fields
            .checked_add(amount)
            .ok_or(Error::InvalidSource)?;
        if self.fields > self.maximum_fields {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireFields,
                observed: self.fields as u64,
                maximum: self.maximum_fields as u64,
            });
        }
        Ok(())
    }

    pub(super) const fn preview_allowance(
        &self,
    ) -> crate::package::slide_preview::InvalidationAllowance {
        crate::package::slide_preview::InvalidationAllowance::new(
            self.maximum_work.saturating_sub(self.work),
            self.maximum_references.saturating_sub(self.references),
        )
    }

    pub(super) fn charge_preview_report(
        &mut self,
        report: crate::package::slide_preview::InvalidationReport,
    ) -> Result<(), Error> {
        self.charge_references(report.references())?;
        self.charge_work(report.work())
    }

    pub(super) fn map_preview_budget_error(
        &self,
        error: crate::package::slide_preview::BudgetedInvalidationError,
    ) -> Error {
        match error {
            crate::package::slide_preview::BudgetedInvalidationError::Invalidation(inner) => {
                map_slide_preview_error(inner)
            },
            crate::package::slide_preview::BudgetedInvalidationError::BudgetExceeded {
                kind,
                observed,
                maximum: _,
            } => match kind {
                crate::package::slide_preview::InvalidationBudgetKind::References => {
                    Error::LimitExceeded {
                        kind: LimitKind::References,
                        observed: self.references.saturating_add(observed) as u64,
                        maximum: self.maximum_references as u64,
                    }
                },
                crate::package::slide_preview::InvalidationBudgetKind::Work => {
                    Error::LimitExceeded {
                        kind: LimitKind::WireWork,
                        observed: self.work.saturating_add(observed) as u64,
                        maximum: self.maximum_work as u64,
                    }
                },
            },
        }
    }
}

impl Package {
    /// Read an existing placeholder's per-slide visibility.
    ///
    /// `None` means that the selected slide has no existing placeholder for
    /// the requested role. [`Some(State::Hidden)`](State::Hidden) means that
    /// the role exists but does not display on that slide; hidden title and
    /// body roles retain their text.
    ///
    /// [`Kind::SlideNumber`] is a per-slide setting. It does not read or
    /// change the separate show-wide slide-number preference.
    ///
    /// # Errors
    /// Returns a typed selector error for a missing or ambiguous slide, a
    /// source error when the selected presentation state is unsupported, or a
    /// resource or allocation error from bounded focused resolution.
    ///
    /// # Costs
    ///
    /// Position selection reads selected slide metadata without decoding title
    /// or body text. Name selection scans navigator names. This read publishes
    /// no output and reopens no candidate package.
    pub fn slide_placeholder_visibility<'s>(
        &self,
        selector: impl Into<SlideSelector<'s>>,
        kind: Kind,
    ) -> Result<Option<State>, Error> {
        Ok(select(self, selector.into(), kind, false)?.map(|selection| selection.state))
    }

    /// Start a selector-first per-slide placeholder-visibility edit.
    ///
    /// The edit affects only the selected slide and role. In particular,
    /// selecting [`Kind::SlideNumber`] does not change the show-wide
    /// slide-number preference.
    ///
    /// # Errors
    /// Returns [`Error::PlaceholderNotFound`] when the requested role is
    /// absent, selector errors for a missing or ambiguous slide, or typed
    /// source, resource, and allocation failures from focused resolution.
    ///
    /// # Costs
    ///
    /// Resolves the same focused rooted chain as
    /// [`Package::slide_placeholder_visibility`] and retains only compact
    /// semantic state plus an immutable package borrow. It performs no rewrite.
    pub fn edit_slide_placeholder_visibility<'s>(
        &self,
        selector: impl Into<SlideSelector<'s>>,
        kind: Kind,
    ) -> Result<Edit<'_>, Error> {
        let selection = select(self, selector.into(), kind, false)?
            .ok_or(Error::PlaceholderNotFound { kind })?;
        Ok(Edit::new(self, selection, kind))
    }

    /// Apply an exact-source visibility patch.
    ///
    /// # Errors
    /// Returns [`Error::PatchConflict`] when the patch does not authorize the
    /// exact source package snapshot or its selected state. Changed application can
    /// also return typed source, resource, allocation, candidate-reopen, or
    /// verification failures.
    ///
    /// # Costs
    ///
    /// An authorized exact no-op returns immediately after package-snapshot
    /// checking, without selector resolution or cache inspection. Changed
    /// application opens the retained target once and verifies the selected
    /// semantic state, preservation of unselected content, and directional
    /// rendering-cache and package-root-preview state.
    pub fn apply_slide_placeholder_visibility(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_catalog(self)?.shared_source();
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
        let mut budget = TransactionBudget::new(self)?;
        let selection = select_with_budget(
            self,
            SlideSelector::Position(patch.position),
            patch.kind,
            true,
            &mut budget,
        )?
        .ok_or(Error::PatchConflict)?;
        if selection.placeholder_identifier != patch.placeholder_identifier
            || selection.state != patch.before
        {
            return Err(Error::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        let touched = verify_artifact_delta(
            self,
            &candidate,
            &selection,
            patch.after,
            patch.target_preview_count,
            patch.source_node_invalidated,
            patch.target_node_invalidated,
            &mut budget,
        )?;
        if touched != patch.touched_components {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.touched_components,
                patch
                    .source_preview_count
                    .saturating_sub(patch.target_preview_count),
            ),
        })
    }
}

fn select<'a>(
    package: &'a Package,
    selector: SlideSelector<'_>,
    kind: Kind,
    mutation: bool,
) -> Result<Option<Selection<'a>>, Error> {
    let mut budget = TransactionBudget::new(package)?;
    select_with_budget(package, selector, kind, mutation, &mut budget)
}

pub(super) fn select_with_budget<'a>(
    package: &'a Package,
    selector: SlideSelector<'_>,
    kind: Kind,
    mutation: bool,
    budget: &mut TransactionBudget,
) -> Result<Option<Selection<'a>>, Error> {
    let record = focused_slide(package, selector, mutation, budget)?;
    let (node_component, node) = package
        .object_with_component(record.node_identifier)
        .ok_or(Error::InvalidSource)?;
    let (slide_component, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(Error::InvalidSource)?;
    let (_node_index, node_payload) = selected_message(node, SLIDE_NODE_MESSAGE_TYPE)?;
    let (slide_index, slide_payload) = selected_message(slide, SLIDE_MESSAGE_TYPE)?;
    budget.charge_work(slide_payload.len())?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let node_slide_number_visible = if kind == Kind::SlideNumber {
        slide_number_node_visible(node_payload, limits, budget)?
    } else {
        false
    };
    let snapshot_scan_work = slide_payload
        .len()
        .checked_mul(3)
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(snapshot_scan_work)?;
    let snapshot = visibility_snapshot(slide_payload, kind, node_slide_number_visible, limits)?;
    let Some((placeholder_identifier, state)) = snapshot else {
        return Ok(None);
    };
    let (placeholder_component, placeholder) = package
        .object_with_component(placeholder_identifier)
        .ok_or(Error::InvalidSource)?;
    let (placeholder_index, placeholder_payload) =
        selected_message(placeholder, PLACEHOLDER_MESSAGE_TYPE)?;
    budget.charge_work(placeholder_payload.len())?;
    let owner = if kind == Kind::SlideNumber {
        None
    } else {
        Some(
            keynote_placeholder_text_codec::decode_placeholder_text_owner(
                placeholder_payload,
                placeholder_options(package, placeholder_payload)?,
            )
            .map_err(map_placeholder_error)?,
        )
    };
    if mutation {
        let mutation_limits = package.wire_limits().map_err(map_wire_error)?;
        let reserved_count = record
            .rooted_node_identifiers
            .len()
            .checked_add(record.rooted_slide_identifiers.len())
            .and_then(|count| count.checked_add(3))
            .ok_or(Error::InvalidSource)?;
        let mut reserved_identifiers = HashSet::new();
        reserved_identifiers
            .try_reserve(reserved_count)
            .map_err(|_allocation| Error::Allocation {
                amount: reserved_count,
            })?;
        if !reserved_identifiers.insert(1)
            || !reserved_identifiers.insert(record.show_identifier)
            || !reserved_identifiers.insert(placeholder_identifier)
            || record
                .rooted_node_identifiers
                .iter()
                .chain(&record.rooted_slide_identifiers)
                .any(|identifier| !reserved_identifiers.insert(*identifier))
        {
            return Err(Error::InvalidSource);
        }
        validate_rooted_placeholder_locality(
            package,
            &record.rooted_slide_identifiers,
            record.slide_identifier,
            placeholder_identifier,
            mutation_limits,
            budget,
        )?;
        if kind == Kind::SlideNumber {
            validate_global_slide_number_ownership(
                package,
                record.slide_identifier,
                placeholder_identifier,
                mutation_limits,
                budget,
            )?;
        }
        let (owner_kind, storage_identifier) = if kind == Kind::SlideNumber {
            slide_number_placeholder_owner(placeholder_payload, mutation_limits)?
        } else {
            let decoded_owner = owner.as_ref().ok_or(Error::InvalidSource)?;
            (
                decoded_owner.kind(),
                decoded_owner
                    .owned_storage()
                    .ok_or(Error::InvalidSource)?
                    .identifier()
                    .get(),
            )
        };
        if slide_component != placeholder_component
            || owner_kind
                != Some(match kind {
                    Kind::Title => 2,
                    Kind::Body => 3,
                    Kind::SlideNumber => 1,
                })
        {
            return Err(Error::InvalidSource);
        }
        validate_selected_metadata(slide, slide_index, placeholder_identifier, kind)?;
        validate_placeholder_metadata(
            placeholder,
            placeholder_index,
            record.slide_identifier,
            (storage_identifier != 0).then_some(storage_identifier),
        )?;
        let placeholder_dependencies = validate_placeholder_reference_closure(
            placeholder,
            placeholder_index,
            placeholder_payload,
            &reserved_identifiers,
            storage_identifier,
            mutation_limits,
            budget,
        )?;
        let extra_roles = usize::from(storage_identifier != 0)
            .checked_add(usize::from(
                kind == Kind::SlideNumber && storage_identifier != 0,
            ))
            .and_then(|count| count.checked_add(placeholder_dependencies.len()))
            .ok_or(Error::InvalidSource)?;
        reserved_identifiers
            .try_reserve(extra_roles)
            .map_err(|_allocation| Error::Allocation {
                amount: extra_roles,
            })?;
        if storage_identifier != 0 && !reserved_identifiers.insert(storage_identifier) {
            return Err(Error::InvalidSource);
        }
        reserved_identifiers.extend(placeholder_dependencies);
        if kind == Kind::SlideNumber && storage_identifier != 0 {
            validate_slide_number_storage(
                package,
                slide_component,
                storage_identifier,
                &mut reserved_identifiers,
                mutation_limits,
                budget,
            )?;
        }
        validate_component_framing(package, slide_component, budget)?;
        if node_component != slide_component {
            validate_component_framing(package, node_component, budget)?;
        }
        validate_mutation_payload(
            package,
            slide,
            slide_index,
            slide_payload,
            placeholder_payload,
            record.slide_identifier,
            placeholder_identifier,
            kind,
            &mut reserved_identifiers,
            budget,
        )?;
        if kind == Kind::SlideNumber {
            validate_slide_number_node_references(
                node,
                node_payload,
                &reserved_identifiers,
                mutation_limits,
                budget,
            )?;
        }
    }
    Ok(Some(Selection {
        position: record.position,
        kind,
        node_identifier: record.node_identifier,
        slide_identifier: record.slide_identifier,
        placeholder_identifier,
        slide_component,
        node_component,
        slide_payload,
        node_payload,
        state,
    }))
}

fn validate_rooted_placeholder_locality(
    package: &Package,
    rooted_slide_identifiers: &[u64],
    selected_slide_identifier: u64,
    placeholder_identifier: u64,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    const DIRECT_REFERENCE_FIELDS: [u32; 17] = [
        1, 2, 5, 6, 7, 17, 20, 27, 29, 30, 31, 35, 36, 39, 42, 43, 44,
    ];
    for slide_identifier in rooted_slide_identifiers {
        if *slide_identifier == selected_slide_identifier {
            continue;
        }
        let (_component, slide) = package
            .object_with_component(*slide_identifier)
            .ok_or(Error::InvalidSource)?;
        let (index, payload) = selected_message(slide, SLIDE_MESSAGE_TYPE)?;
        budget.charge_work(payload.len())?;
        let info = slide
            .archive_info
            .message_infos
            .get(index)
            .ok_or(Error::InvalidSource)?;
        validate_merge_metadata(slide, info)?;
        charge_reference_metadata_scan(slide, index, budget)?;
        if info.object_references.contains(&placeholder_identifier) {
            return Err(Error::InvalidSource);
        }
        let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
        for field in view.fields() {
            field.validate_canonical_framing().map_err(map_wire_error)?;
            if matches!(field.wire_type(), 3 | 4) {
                return Err(Error::InvalidSource);
            }
            if DIRECT_REFERENCE_FIELDS.contains(&field.number()) {
                budget.charge_reference()?;
                if strict_reference(field, limits)? == placeholder_identifier {
                    return Err(Error::InvalidSource);
                }
            }
        }
        for path in [&[28, 2][..], &[45, 1, 1][..]] {
            if let Some(field) = nested_unique_field(payload, path, limits)? {
                budget.charge_reference()?;
                if strict_reference(field, limits)? == placeholder_identifier {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    Ok(())
}

fn visibility_snapshot(
    source: &[u8],
    kind: Kind,
    node_slide_number_visible: bool,
    limits: WireLimits,
) -> Result<Option<(u64, State)>, Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let role_field = match kind {
        Kind::Title => TITLE_REFERENCE_FIELD,
        Kind::Body => BODY_REFERENCE_FIELD,
        Kind::SlideNumber => SLIDE_NUMBER_PLACEHOLDER_FIELD,
    };
    let other_field = match kind {
        Kind::Title => BODY_REFERENCE_FIELD,
        Kind::Body | Kind::SlideNumber => TITLE_REFERENCE_FIELD,
    };
    let mut role = None;
    let mut other = None;
    for field in view.fields() {
        if field.number() == role_field {
            if role.replace(strict_reference(field, limits)?).is_some() {
                return Err(Error::InvalidSource);
            }
        } else if field.number() == other_field
            && other.replace(strict_reference(field, limits)?).is_some()
        {
            return Err(Error::InvalidSource);
        }
    }
    let Some(identifier) = role else {
        return if kind == Kind::SlideNumber && node_slide_number_visible {
            Err(Error::InvalidSource)
        } else {
            Ok(None)
        };
    };
    if other == Some(identifier) {
        return Err(Error::InvalidSource);
    }
    for field in view.fields() {
        if matches!(
            field.number(),
            TITLE_REFERENCE_FIELD | BODY_REFERENCE_FIELD | SLIDE_NUMBER_PLACEHOLDER_FIELD
        ) && field.number() != role_field
            && strict_reference(field, limits)? == identifier
        {
            return Err(Error::InvalidSource);
        }
    }
    let mut owned = 0usize;
    let mut z_order = 0usize;
    for field in view.fields() {
        if matches!(field.number(), OWNED_DRAWABLES_FIELD | Z_ORDER_FIELD) {
            let candidate = strict_reference(field, limits)?;
            if candidate == identifier {
                if field.number() == OWNED_DRAWABLES_FIELD {
                    owned += 1;
                } else {
                    z_order += 1;
                }
            }
        }
    }
    match (kind, owned, z_order, node_slide_number_visible) {
        (Kind::SlideNumber, 0, 0, false) => Ok(Some((identifier, State::Hidden))),
        (Kind::SlideNumber, 1, 1, true) => Ok(Some((identifier, State::Visible))),
        (Kind::SlideNumber, _, _, _) => Err(Error::InvalidSource),
        (_, 0, 0, _) => Ok(Some((identifier, State::Hidden))),
        (_, 1, 1, _) => Ok(Some((identifier, State::Visible))),
        _ => Err(Error::InvalidSource),
    }
}

fn strict_reference(field: WireFieldView<'_>, limits: WireLimits) -> Result<u64, Error> {
    strict_reference_with_zero(field, limits, false)
}

fn strict_reference_with_zero(
    field: WireFieldView<'_>,
    limits: WireLimits,
    allow_zero: bool,
) -> Result<u64, Error> {
    field.validate_canonical_framing().map_err(map_wire_error)?;
    if field.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    let view = WireView::parse_with_limits(field.payload(), limits).map_err(map_wire_error)?;
    let mut identifier = None;
    let mut deprecated_type_seen = false;
    let mut external = None;
    for nested in view.fields() {
        nested
            .validate_canonical_framing()
            .map_err(map_wire_error)?;
        match nested.number() {
            1 => {
                if identifier.replace(canonical_varint(nested)?).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            2 => {
                if std::mem::replace(&mut deprecated_type_seen, true) {
                    return Err(Error::InvalidSource);
                }
                let deprecated_type = canonical_varint(nested)?;
                if deprecated_type > i32::MAX as u64 && deprecated_type < 0xffff_ffff_8000_0000 {
                    return Err(Error::InvalidSource);
                }
            },
            3 => {
                if external.replace(canonical_varint(nested)?).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            _ if matches!(nested.wire_type(), 3 | 4) => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    let resolved_identifier = identifier.ok_or(Error::InvalidSource)?;
    if resolved_identifier == 0 && !allow_zero {
        return Err(Error::InvalidSource);
    }
    if !matches!(external, None | Some(0)) {
        return Err(Error::InvalidSource);
    }
    Ok(resolved_identifier)
}

fn canonical_varint(field: WireFieldView<'_>) -> Result<u64, Error> {
    if field.wire_type() != 0 {
        return Err(Error::InvalidSource);
    }
    let (value, consumed) =
        decode_varint_from_bytes(field.payload()).map_err(|_| Error::InvalidSource)?;
    if consumed != field.payload().len() || consumed != canonical_varint_len(value) {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

const fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn selected_message(object: &ArchiveObject, kind: u32) -> Result<(usize, &[u8]), Error> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut selected = None;
    for (index, (message, info)) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
        .enumerate()
    {
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
        if message.type_ == kind && selected.replace((index, message.data.as_slice())).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn validate_selected_metadata(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    kind: Kind,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    validate_merge_metadata(object, info)?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource);
    }
    let role_path = [match kind {
        Kind::Title => TITLE_REFERENCE_FIELD,
        Kind::Body => BODY_REFERENCE_FIELD,
        Kind::SlideNumber => SLIDE_NUMBER_PLACEHOLDER_FIELD,
    }];
    let mut role_seen = false;
    for field in &info.field_infos {
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if field.path.as_slice() == role_path {
            if std::mem::replace(&mut role_seen, true) || occurrences > 1 {
                return Err(Error::InvalidSource);
            }
        } else if occurrences != 0 {
            // List-local ownership would become stale when the payload member
            // is removed or inserted. V1 refuses it rather than pruning the
            // still-live aggregate and role ownership evidence.
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn validate_placeholder_metadata(
    object: &ArchiveObject,
    message_index: usize,
    slide_identifier: u64,
    storage_identifier: Option<u64>,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    validate_merge_metadata(object, info)?;
    if storage_identifier.is_none()
        && (info.object_references.contains(&0)
            || info
                .field_infos
                .iter()
                .any(|field| field.object_references.contains(&0)))
    {
        return Err(Error::InvalidSource);
    }
    if storage_identifier == Some(slide_identifier)
        || info
            .object_references
            .iter()
            .filter(|identifier| **identifier == slide_identifier)
            .count()
            > 1
        || storage_identifier.is_some_and(|candidate_storage_identifier| {
            info.object_references
                .iter()
                .filter(|identifier| **identifier == candidate_storage_identifier)
                .count()
                != 1
        })
    {
        return Err(Error::InvalidSource);
    }
    let mut parent_path_seen = false;
    let mut deprecated_storage_path_seen = false;
    let mut modern_storage_path_seen = false;
    for field in &info.field_infos {
        let parent_occurrences = field
            .object_references
            .iter()
            .filter(|identifier| **identifier == slide_identifier)
            .count();
        let storage_occurrences = storage_identifier.map_or(0, |candidate_storage_identifier| {
            field
                .object_references
                .iter()
                .filter(|identifier| **identifier == candidate_storage_identifier)
                .count()
        });
        if parent_occurrences != 0
            && (field.path.as_slice() != [1, 1, 1, 2]
                || parent_occurrences != 1
                || std::mem::replace(&mut parent_path_seen, true))
        {
            return Err(Error::InvalidSource);
        }
        if storage_occurrences != 0 {
            let seen = match field.path.as_slice() {
                [1, 2] => &mut deprecated_storage_path_seen,
                [1, 4] => &mut modern_storage_path_seen,
                _ => return Err(Error::InvalidSource),
            };
            if storage_occurrences != 1 || std::mem::replace(seen, true) {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(())
}

fn validate_placeholder_reference_closure(
    object: &ArchiveObject,
    message_index: usize,
    payload: &[u8],
    reserved_identifiers: &HashSet<u64>,
    storage_identifier: u64,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<HashSet<u64>, Error> {
    let drawable = nested_unique_field(payload, &[1, 1, 1], limits)?.ok_or(Error::InvalidSource)?;
    if drawable.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    let view = WireView::parse_with_limits(drawable.payload(), limits).map_err(map_wire_error)?;
    let dependency_count = view
        .fields()
        .filter(|field| matches!(field.number(), 6 | 9 | 10 | 11))
        .count()
        .checked_add(2)
        .ok_or(Error::InvalidSource)?;
    let mut dependencies = HashMap::new();
    dependencies
        .try_reserve(dependency_count)
        .map_err(|_allocation| Error::Allocation {
            amount: dependency_count,
        })?;
    for path in [&[1, 1, 2][..], &[1, 3][..]] {
        if let Some(field) = nested_unique_field(payload, path, limits)? {
            collect_placeholder_dependency(
                field,
                PlaceholderDependencyPath::from_path(path).ok_or(Error::InvalidSource)?,
                reserved_identifiers,
                storage_identifier,
                &mut dependencies,
                budget,
                limits,
            )?;
        }
    }
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
        if matches!(field.number(), 6 | 9 | 10 | 11) {
            let path = [1, 1, 1, field.number()];
            collect_placeholder_dependency(
                field,
                PlaceholderDependencyPath::from_path(&path).ok_or(Error::InvalidSource)?,
                reserved_identifiers,
                storage_identifier,
                &mut dependencies,
                budget,
                limits,
            )?;
        }
    }
    validate_placeholder_dependency_metadata(object, message_index, &mut dependencies, budget)?;
    let mut identifiers = HashSet::new();
    identifiers
        .try_reserve(dependencies.len())
        .map_err(|_allocation| Error::Allocation {
            amount: dependencies.len(),
        })?;
    identifiers.extend(dependencies.into_keys());
    Ok(identifiers)
}

fn collect_placeholder_dependency(
    field: WireFieldView<'_>,
    path: PlaceholderDependencyPath,
    reserved_identifiers: &HashSet<u64>,
    storage_identifier: u64,
    dependencies: &mut HashMap<u64, PlaceholderDependencyDeclaration>,
    budget: &mut TransactionBudget,
    limits: WireLimits,
) -> Result<(), Error> {
    budget.charge_reference()?;
    let identifier = strict_reference(field, limits)?;
    if reserved_identifiers.contains(&identifier)
        || identifier == storage_identifier
        || dependencies
            .insert(
                identifier,
                PlaceholderDependencyDeclaration {
                    path,
                    aggregate_occurrences: 0,
                    field_attributed: false,
                },
            )
            .is_some()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_placeholder_dependency_metadata(
    object: &ArchiveObject,
    message_index: usize,
    dependencies: &mut HashMap<u64, PlaceholderDependencyDeclaration>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    validate_merge_metadata(object, info)?;

    // Build the declaration index once, then validate K dependencies through
    // O(1) lookups instead of rescanning all aggregate and FieldInfo metadata
    // K times. Charge every inspected metadata scalar to the transaction.
    for identifier in &info.object_references {
        budget.charge_work(8)?;
        if let Some(dependency) = dependencies.get_mut(identifier) {
            dependency.aggregate_occurrences = dependency
                .aggregate_occurrences
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
        }
    }
    for field in &info.field_infos {
        let path_work = field
            .path
            .as_slice()
            .len()
            .checked_mul(4)
            .ok_or(Error::InvalidSource)?;
        budget.charge_work(path_work)?;
        for identifier in &field.object_references {
            budget.charge_work(8)?;
            let Some(dependency) = dependencies.get_mut(identifier) else {
                continue;
            };
            if field.path.as_slice() != dependency.path.as_slice()
                || std::mem::replace(&mut dependency.field_attributed, true)
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    if dependencies
        .values()
        .any(|dependency| dependency.aggregate_occurrences != 1)
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_merge_metadata(object: &ArchiveObject, info: &MessageInfo) -> Result<(), Error> {
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_reference_metadata(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    path: &[u32],
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    validate_merge_metadata(object, info)?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource);
    }
    let mut matching_path = false;
    for field in &info.field_infos {
        if field.path.as_slice() == path {
            if std::mem::replace(&mut matching_path, true)
                || field
                    .object_references
                    .iter()
                    .filter(|candidate| **candidate == identifier)
                    .count()
                    > 1
            {
                return Err(Error::InvalidSource);
            }
        } else if field.object_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn charge_reference_metadata_scan(
    object: &ArchiveObject,
    message_index: usize,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let metadata_work = info
        .object_references
        .len()
        .checked_mul(8)
        .and_then(|work| work.checked_add(1))
        .and_then(|initial_work| {
            info.field_infos
                .iter()
                .try_fold(initial_work, |accumulated_work, field| {
                    field
                        .path
                        .as_slice()
                        .len()
                        .checked_mul(4)
                        .and_then(|path| accumulated_work.checked_add(path))
                        .and_then(|path_work| {
                            field
                                .object_references
                                .len()
                                .checked_mul(8)
                                .and_then(|refs| path_work.checked_add(refs))
                        })
                        .and_then(|work| work.checked_add(1))
                })
        })
        .and_then(|one_pass| one_pass.checked_mul(2))
        .ok_or(Error::InvalidSource)?;
    let references = info
        .object_references
        .len()
        .checked_add(
            info.field_infos
                .iter()
                .try_fold(0usize, |count, field| {
                    count.checked_add(field.object_references.len())
                })
                .ok_or(Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(metadata_work)?;
    budget.charge_references(references)
}

pub(super) fn validate_reference_metadata_set(
    object: &ArchiveObject,
    message_index: usize,
    identifiers: &HashSet<u64>,
    path: &[u32],
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    validate_merge_metadata(object, info)?;
    let mut counts = HashMap::new();
    counts
        .try_reserve(identifiers.len())
        .map_err(|_allocation| Error::Allocation {
            amount: identifiers.len(),
        })?;
    for identifier in identifiers {
        counts.insert(*identifier, 0usize);
    }
    for identifier in &info.object_references {
        if let Some(count) = counts.get_mut(identifier) {
            *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    if counts.values().any(|count| *count != 1) {
        return Err(Error::InvalidSource);
    }
    let mut attributed = HashSet::new();
    attributed
        .try_reserve(identifiers.len())
        .map_err(|_allocation| Error::Allocation {
            amount: identifiers.len(),
        })?;
    for field in &info.field_infos {
        for identifier in &field.object_references {
            if identifiers.contains(identifier)
                && (field.path.as_slice() != path || !attributed.insert(*identifier))
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(())
}

fn validate_component_framing(
    package: &Package,
    component_name: &str,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let catalog = physical_catalog(package)?;
    let component = catalog
        .components()
        .get(component_name)
        .ok_or(Error::InvalidSource)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        package
            .state
            .options
            .archive()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    budget.charge_work(stream.as_bytes().len())?;
    component
        .archive()
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)
}

fn validate_mutation_payload(
    package: &Package,
    slide_object: &ArchiveObject,
    slide_message_index: usize,
    slide_payload: &[u8],
    placeholder_payload: &[u8],
    slide_identifier: u64,
    placeholder_identifier: u64,
    kind: Kind,
    reserved_identifiers: &mut HashSet<u64>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let slide = WireView::parse_with_limits(slide_payload, limits).map_err(map_wire_error)?;
    let field_count = slide.fields().count();
    let mut builds = HashSet::new();
    builds
        .try_reserve(field_count.min(package.semantic_limits().max_references()))
        .map_err(|_allocation| Error::Allocation {
            amount: field_count,
        })?;
    let slide_work = field_count
        .checked_mul(2)
        .and_then(|field_work| slide_payload.len().checked_add(field_work))
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(slide_work)?;
    let mut style = None;
    let mut other_roles = HashSet::new();
    if kind == Kind::SlideNumber {
        other_roles
            .try_reserve(field_count)
            .map_err(|_allocation| Error::Allocation {
                amount: field_count,
            })?;
    }
    for field in slide.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
        match field.number() {
            SLIDE_STYLE_FIELD => {
                budget.charge_reference()?;
                if style.replace(strict_reference(field, limits)?).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            SLIDE_BUILDS_FIELD => {
                budget.charge_reference()?;
                if !builds.insert(strict_reference(field, limits)?) {
                    return Err(Error::InvalidSource);
                }
            },
            3 => return Err(Error::UnsupportedSource),
            TITLE_REFERENCE_FIELD
            | BODY_REFERENCE_FIELD
            | SLIDE_NUMBER_PLACEHOLDER_FIELD
            | OWNED_DRAWABLES_FIELD
            | Z_ORDER_FIELD => {
                budget.charge_reference()?;
                let identifier = strict_reference(field, limits)?;
                if kind == Kind::SlideNumber {
                    let controlled = field.number() == SLIDE_NUMBER_PLACEHOLDER_FIELD
                        || matches!(field.number(), OWNED_DRAWABLES_FIELD | Z_ORDER_FIELD);
                    if identifier == placeholder_identifier && !controlled {
                        return Err(Error::InvalidSource);
                    }
                    if identifier != placeholder_identifier {
                        let inserted = other_roles.insert(identifier);
                        if !inserted
                            && !matches!(field.number(), OWNED_DRAWABLES_FIELD | Z_ORDER_FIELD)
                        {
                            return Err(Error::InvalidSource);
                        }
                    }
                }
            },
            OBJECT_PLACEHOLDER_FIELD | SLIDE_TEMPLATE_REFERENCE_FIELD => {
                budget.charge_reference()?;
                let identifier = strict_reference(field, limits)?;
                if identifier == placeholder_identifier {
                    return Err(Error::InvalidSource);
                }
                if kind == Kind::SlideNumber && !other_roles.insert(identifier) {
                    return Err(Error::InvalidSource);
                }
            },
            number if SLIDE_KNOWN_REFERENCE_FIELDS.contains(&number) => {
                budget.charge_reference()?;
                let identifier = strict_reference(field, limits)?;
                if identifier == placeholder_identifier {
                    return Err(Error::InvalidSource);
                }
                if number == 43 {
                    return Err(Error::UnsupportedSource);
                }
                if kind == Kind::SlideNumber && !other_roles.insert(identifier) {
                    return Err(Error::InvalidSource);
                }
            },
            SLIDE_LAYERING_FIELD => {
                if canonical_varint(field)? != 0 {
                    return Err(Error::InvalidSource);
                }
            },
            SLIDE_TITLE_CACHE_FIELD | SLIDE_BODY_CACHE_FIELD => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    for path in [&[28, 2][..], &[45, 1, 1][..]] {
        if let Some(field) = nested_unique_field(slide_payload, path, limits)? {
            budget.charge_reference()?;
            let identifier = strict_reference(field, limits)?;
            if identifier == placeholder_identifier {
                return Err(Error::InvalidSource);
            }
            if kind == Kind::SlideNumber {
                // These maps point back to already-owned drawables in native
                // slides, so their exact reference may legitimately repeat a
                // field-7/42 dependency. They still cannot alias the selected
                // slide-number placeholder or any reserved closure role.
                other_roles.insert(identifier);
            }
        }
    }
    let style_identifier = style.ok_or(Error::InvalidSource)?;
    if reserved_identifiers.contains(&style_identifier) || builds.contains(&style_identifier) {
        return Err(Error::InvalidSource);
    }
    if builds
        .iter()
        .any(|identifier| reserved_identifiers.contains(identifier))
    {
        return Err(Error::InvalidSource);
    }
    if kind == Kind::SlideNumber
        && other_roles.iter().any(|identifier| {
            reserved_identifiers.contains(identifier)
                || *identifier == style_identifier
                || builds.contains(identifier)
        })
    {
        return Err(Error::InvalidSource);
    }
    validate_slide_dependency_metadata(
        slide_object,
        slide_message_index,
        style_identifier,
        &builds,
    )?;
    let added_roles = builds
        .len()
        .checked_add(1)
        .and_then(|count| count.checked_add(other_roles.len()))
        .ok_or(Error::InvalidSource)?;
    reserved_identifiers
        .try_reserve(added_roles)
        .map_err(|_allocation| Error::Allocation {
            amount: added_roles,
        })?;
    if !reserved_identifiers.insert(style_identifier) {
        return Err(Error::InvalidSource);
    }
    for build_identifier in &builds {
        if !reserved_identifiers.insert(*build_identifier) {
            return Err(Error::InvalidSource);
        }
    }
    for identifier in other_roles {
        if !reserved_identifiers.insert(identifier) {
            return Err(Error::InvalidSource);
        }
    }
    for build_identifier in builds {
        validate_build_dependency(
            package,
            build_identifier,
            placeholder_identifier,
            limits,
            budget,
        )?;
    }
    validate_style_visibility(package, style_identifier, kind, budget)?;
    validate_placeholder_owner(placeholder_payload, slide_identifier, kind, limits, budget)?;
    Ok(())
}

fn validate_slide_dependency_metadata(
    object: &ArchiveObject,
    message_index: usize,
    style_identifier: u64,
    build_identifiers: &HashSet<u64>,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let dependency_count = build_identifiers
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    let mut aggregate_counts = HashMap::new();
    aggregate_counts
        .try_reserve(dependency_count)
        .map_err(|_allocation| Error::Allocation {
            amount: dependency_count,
        })?;
    aggregate_counts.insert(style_identifier, 0usize);
    for identifier in build_identifiers {
        aggregate_counts.insert(*identifier, 0);
    }
    for identifier in &info.object_references {
        if let Some(count) = aggregate_counts.get_mut(identifier) {
            *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    if aggregate_counts.values().any(|count| *count != 1) {
        return Err(Error::InvalidSource);
    }
    let mut attributed = HashSet::new();
    attributed
        .try_reserve(dependency_count)
        .map_err(|_allocation| Error::Allocation {
            amount: dependency_count,
        })?;
    for field in &info.field_infos {
        for identifier in &field.object_references {
            if !aggregate_counts.contains_key(identifier) {
                continue;
            }
            let expected_path = if *identifier == style_identifier {
                [SLIDE_STYLE_FIELD]
            } else {
                [SLIDE_BUILDS_FIELD]
            };
            if field.path.as_slice() != expected_path || !attributed.insert(*identifier) {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(())
}

fn validate_placeholder_owner(
    payload: &[u8],
    slide_identifier: u64,
    kind: Kind,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    reject_groups(payload, limits)?;
    let parent =
        nested_unique_field(payload, &[1, 1, 1, 2], limits)?.ok_or(Error::InvalidSource)?;
    budget.charge_reference()?;
    if strict_reference(parent, limits)? != slide_identifier {
        return Err(Error::InvalidSource);
    }
    let locked = nested_unique_field(payload, &[1, 1, 1, 5], limits)?
        .map(canonical_varint)
        .transpose()?;
    if locked.is_some_and(|value| value != 0) {
        return Err(Error::InvalidSource);
    }
    let deprecated = nested_unique_field(payload, &[1, 2], limits)?
        .map(|field| {
            let identifier = strict_reference_with_zero(field, limits, kind == Kind::SlideNumber)?;
            if identifier != 0 {
                budget.charge_reference()?;
            }
            Ok(identifier)
        })
        .transpose()?;
    let modern = nested_unique_field(payload, &[1, 4], limits)?
        .map(|field| {
            let identifier = strict_reference_with_zero(field, limits, kind == Kind::SlideNumber)?;
            if identifier != 0 {
                budget.charge_reference()?;
            }
            Ok(identifier)
        })
        .transpose()?;
    if deprecated.is_some() && modern.is_some() && deprecated != modern {
        return Err(Error::InvalidSource);
    }
    let expected_kind = match kind {
        Kind::Title => 2,
        Kind::Body => 3,
        Kind::SlideNumber => 1,
    };
    let kind_field = nested_unique_field(payload, &[2], limits)?.ok_or(Error::InvalidSource)?;
    if canonical_varint(kind_field)? != expected_kind {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_style_visibility(
    package: &Package,
    identifier: u64,
    kind: Kind,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let (_component, object) = package
        .object_with_component(identifier)
        .ok_or(Error::InvalidSource)?;
    let (index, payload) = selected_message(object, SLIDE_STYLE_MESSAGE_TYPE)?;
    validate_merge_metadata(
        object,
        object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(Error::InvalidSource)?,
    )?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    budget.charge_work(payload.len())?;
    reject_groups(payload, limits)?;
    if let Some(properties) = nested_unique_field(payload, &[11], limits)? {
        let view =
            WireView::parse_with_limits(properties.payload(), limits).map_err(map_wire_error)?;
        for field in view.fields() {
            field.validate_canonical_framing().map_err(map_wire_error)?;
            let selected_override = match kind {
                Kind::Title | Kind::Body => matches!(field.number(), 4 | 5),
                Kind::SlideNumber => field.number() == 6,
            };
            if selected_override {
                return Err(Error::UnsupportedSource);
            }
        }
    }
    Ok(())
}

fn validate_build_dependency(
    package: &Package,
    build_identifier: u64,
    placeholder_identifier: u64,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    const BUILD_MESSAGE_TYPE: u32 = 8;
    let (_component, object) = package
        .object_with_component(build_identifier)
        .ok_or(Error::InvalidSource)?;
    let (index, payload) = selected_message(object, BUILD_MESSAGE_TYPE)?;
    validate_merge_metadata(
        object,
        object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(Error::InvalidSource)?,
    )?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let field_count = view.fields().count();
    let payload_work = field_count
        .checked_mul(2)
        .and_then(|field_work| payload.len().checked_add(field_work))
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(payload_work)?;
    let mut build_target = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
        if field.number() == 1 {
            budget.charge_reference()?;
            if build_target
                .replace(strict_reference(field, limits)?)
                .is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    if build_target == Some(placeholder_identifier) {
        return Err(Error::UnsupportedSource);
    }
    if let Some(target) = build_target {
        validate_reference_metadata(object, index, target, &[1])?;
    }
    Ok(())
}

fn reject_groups(source: &[u8], limits: WireLimits) -> Result<(), Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn nested_unique_field<'a>(
    source: &'a [u8],
    path: &[u32],
    limits: WireLimits,
) -> Result<Option<WireFieldView<'a>>, Error> {
    let Some((&number, rest)) = path.split_first() else {
        return Err(Error::InvalidSource);
    };
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut found = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if matches!(field.wire_type(), 3 | 4) {
            return Err(Error::InvalidSource);
        }
        if field.number() == number && found.replace(field).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let Some(selected) = found else {
        return Ok(None);
    };
    if rest.is_empty() {
        return Ok(Some(selected));
    }
    if selected.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    nested_unique_field(selected.payload(), rest, limits)
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::{FieldInfo, FieldPath};

    use super::*;

    fn dependency_declarations(
        identifiers: impl IntoIterator<Item = u64>,
    ) -> HashMap<u64, PlaceholderDependencyDeclaration> {
        identifiers
            .into_iter()
            .map(|identifier| {
                (
                    identifier,
                    PlaceholderDependencyDeclaration {
                        path: PlaceholderDependencyPath::PencilAnnotation,
                        aggregate_occurrences: 0,
                        field_attributed: false,
                    },
                )
            })
            .collect()
    }

    #[test]
    fn dependency_metadata_is_indexed_once_for_many_annotations()
    -> Result<(), Box<dyn std::error::Error>> {
        const DEPENDENCIES: usize = 128;
        let fixture = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/basic.key");
        let package = Package::open(fixture)?;
        let selection = select(&package, SlideSelector::index(0), Kind::Title, false)?
            .ok_or(Error::InvalidSource)?;
        let (_component, source) = package
            .object_with_component(selection.placeholder_identifier)
            .ok_or(Error::InvalidSource)?;
        let (message_index, _payload) = selected_message(source, PLACEHOLDER_MESSAGE_TYPE)?;
        let mut object = source.clone();
        let info = object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .ok_or(Error::InvalidSource)?;
        let identifiers: Vec<_> = (0..DEPENDENCIES)
            .map(|offset| 9_000_000_u64 + offset as u64)
            .collect();
        for identifier in &identifiers {
            info.object_references.push(*identifier);
            info.field_infos.push(FieldInfo {
                path: FieldPath::new(vec![1, 1, 1, 9]),
                object_references: vec![*identifier],
                ..FieldInfo::default()
            });
        }
        let mut declarations = dependency_declarations(identifiers.iter().copied());
        let mut budget = TransactionBudget::new(&package)?;
        validate_placeholder_dependency_metadata(
            &object,
            message_index,
            &mut declarations,
            &mut budget,
        )?;

        object.archive_info.message_infos[message_index]
            .field_infos
            .push(FieldInfo {
                path: FieldPath::new(vec![1, 1, 1, 9]),
                object_references: vec![identifiers[0]],
                ..FieldInfo::default()
            });
        let mut declarations = dependency_declarations(identifiers);
        let mut budget = TransactionBudget::new(&package)?;
        assert!(matches!(
            validate_placeholder_dependency_metadata(
                &object,
                message_index,
                &mut declarations,
                &mut budget,
            ),
            Err(Error::InvalidSource)
        ));
        Ok(())
    }
}
