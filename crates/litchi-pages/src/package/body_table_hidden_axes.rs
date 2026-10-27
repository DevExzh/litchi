//! Exact-source hidden-row/hidden-column transactions for Pages tables.
//!
//! The public value is deliberately archive free.  This adapter resolves the
//! rooted table through `table_lock`; the shared engine
//! (`litchi_iwa_table::hidden_axes`, record 0379) proves its hidden-state
//! dependency closure and rewrites it.  Payload surgery is limited to the
//! selected table-info/model fields; all unrelated wire bytes and all
//! unrelated package members remain source authoritative.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused transaction keeps its proof and rewrite helpers together."
)]

use std::fmt;
use std::sync::Arc;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_table::hidden_axes::{self as engine, support::fingerprint};
use thiserror::Error as ThisError;

use super::{Package, PackageError, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::hidden_axes::HiddenAxes;

/// A content-free location associated with one hidden-axis operation.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Pages package.
    Package,
    /// One rooted table at a checked body position.
    Table { table: usize },
}

impl fmt::Debug for Path {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // A table position is useful to the private adapter, but public
        // diagnostics are intentionally content-free.  In particular, do not
        // expose the selector, native proof, or any package member here.
        formatter.write_str(match self {
            Self::Package => "Path::Package",
            Self::Table { .. } => "Path::Table",
        })
    }
}

/// Finite resources governed by one hidden-axis transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalEntryBytes,
    PackageBytes,
    PayloadBytes,
    TotalPayloadBytes,
    PayloadObjects,
    PayloadMessages,
    PayloadItems,
    PayloadReferences,
    WireBytes,
    WireOutputBytes,
    WireFields,
    WireNesting,
    WireWork,
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
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

/// A redacted failure from a body-table hidden-axis read or transaction.
#[derive(ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the selected Pages body-table hidden-axis source is invalid")]
    InvalidSource,
    #[error(
        "Pages body-table hidden axes {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for Pages body-table hidden axes")]
    Allocation { amount: usize },
    #[error("the edited Pages body-table hidden axes failed semantic verification")]
    Verification,
    #[error("the Pages body-table hidden-axis patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    #[error("more than one Pages body table has the requested name")]
    AmbiguousTableName,
    #[error("the selected Pages body-table selector is ambiguous")]
    AmbiguousSelector,
    #[error("this Pages source does not support exact body-table hidden-axis editing")]
    UnsupportedSource,
    #[error("the selected Pages body table is locked")]
    TableLocked,
    #[error("the selected Pages body-table hidden-axis dependency is unsupported")]
    UnsupportedDependency,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::AmbiguousTableName { .. } | Self::AmbiguousSelector { .. } => {
                RefusalKind::InvalidValue
            },
            Self::UnsupportedSource { .. } | Self::UnsupportedDependency { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::TableLocked { .. } => RefusalKind::Protected,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl fmt::Debug for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Limit {
                kind,
                observed,
                maximum,
            } => {
                // Resource categories and amounts are bounded diagnostics;
                // no source fingerprint, payload, UUID, or native identity is
                // ever included in this public representation.
                formatter
                    .debug_struct("Error::Limit")
                    .field("kind", kind)
                    .field("observed", observed)
                    .field("maximum", maximum)
                    .finish_non_exhaustive()
            },
            Self::Allocation { amount } => formatter
                .debug_struct("Error::Allocation")
                .field("allocation_units", amount)
                .finish_non_exhaustive(),
            Self::Refused(Refusal::TableNotFound) => formatter.write_str("Error::TableNotFound"),
            Self::Refused(Refusal::AmbiguousTableName) => {
                formatter.write_str("Error::AmbiguousTableName")
            },
            Self::Refused(Refusal::AmbiguousSelector) => {
                formatter.write_str("Error::AmbiguousSelector")
            },
            Self::Refused(Refusal::UnsupportedSource) => {
                formatter.write_str("Error::UnsupportedSource")
            },
            Self::InvalidSource => formatter.write_str("Error::InvalidSource"),
            Self::Refused(Refusal::TableLocked) => formatter.write_str("Error::TableLocked"),
            Self::Refused(Refusal::UnsupportedDependency) => {
                formatter.write_str("Error::UnsupportedDependency")
            },
            Self::Verification => formatter.write_str("Error::Verification"),
            Self::PatchConflict => formatter.write_str("Error::PatchConflict"),
        }
    }
}

/// Selector-first immutable hidden-axis edit.
pub struct Edit<'a> {
    source: &'a Package,
    selector: table_lock::OwnedBodyTableSelector,
    target: Option<engine::Target>,
    before: Option<HiddenAxes>,
    axes: HiddenAxes,
    staged: bool,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            // Hidden row/column positions are semantic content.  Public
            // diagnostics may report only bounded counts and state changes;
            // the source, proof, and selected package member stay private.
            .field("before_resolved", &self.before.is_some())
            .field("axes_count", &self.axes.as_slice().len())
            .field(
                "changed",
                &self
                    .before
                    .as_ref()
                    .is_some_and(|before| before != &self.axes),
            )
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn resolve(&mut self) -> Result<(), Error> {
        if self.target.is_none() {
            let mut budget = transaction_budget(self.source)?;
            let selector = match &self.selector {
                table_lock::OwnedBodyTableSelector::Name(name) => BodyTableSelector::Name(name),
                table_lock::OwnedBodyTableSelector::Position(position) => {
                    BodyTableSelector::Position(*position)
                },
            };
            let graph = engine::resolve_graph(&Leaf::of(self.source), selector, &mut budget)?;
            if self.before.is_none() {
                let before = graph.before().clone();
                if !self.staged {
                    self.axes = before.clone();
                }
                self.before = Some(before);
            }
            self.target = Some(graph.target().clone());
        }
        Ok(())
    }

    pub fn path(&mut self) -> Result<Path, Error> {
        self.resolve()?;
        self.target
            .as_ref()
            .map(|target| Path::Table {
                table: target.table_position,
            })
            .ok_or(Error::InvalidSource)
    }

    pub fn before(&mut self) -> Result<&HiddenAxes, Error> {
        self.resolve()?;
        self.before.as_ref().ok_or(Error::InvalidSource)
    }

    pub fn axes(&mut self) -> Result<&HiddenAxes, Error> {
        self.resolve()?;
        Ok(&self.axes)
    }

    pub fn hidden_axes(&mut self) -> Result<&HiddenAxes, Error> {
        self.resolve()?;
        Ok(&self.axes)
    }

    #[must_use]
    pub fn set(mut self, axes: HiddenAxes) -> Self {
        self.axes = axes;
        self.staged = true;
        self
    }

    #[must_use]
    pub fn clear(self) -> Self {
        self.set(HiddenAxes::empty())
    }

    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    pub fn commit(mut self) -> Result<Commit, Error> {
        self.resolve()?;
        commit_edit(self)
    }
}

/// Exact-source reversible hidden-axis patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: engine::Target,
    before: HiddenAxes,
    after: HiddenAxes,
    source_previews: usize,
    target_previews: usize,
    touched_components: usize,
    added_object_ids: Arc<[u64]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("before_count", &self.before.as_slice().len())
            .field("after_count", &self.after.as_slice().len())
            .field("changed", &(self.before != self.after))
            .field("noop", &self.is_noop())
            .field("touched_components", &self.touched_components)
            .finish_non_exhaustive()
    }
}

impl Patch {
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            table: self.proof.table_position,
        }
    }

    #[must_use]
    pub const fn before(&self) -> &HiddenAxes {
        &self.before
    }

    #[must_use]
    pub const fn after(&self) -> &HiddenAxes {
        &self.after
    }

    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && self.added_object_ids.is_empty()
            && self.touched_components == 0
            && (Arc::ptr_eq(&self.source, &self.target) || self.source == self.target)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            proof: self.proof.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
            source_previews: self.target_previews,
            target_previews: self.source_previews,
            touched_components: self.touched_components,
            added_object_ids: Arc::clone(&self.added_object_ids),
        }
    }
}

/// Content-free publication diagnostics.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl fmt::Debug for Diagnostics {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Diagnostics")
            .field("changed", &self.changed)
            .field("touched_components", &self.touched_components)
            .field("deleted_previews", &self.deleted_previews)
            .field("full_reparse_performed", &self.full_reparse_performed)
            .finish_non_exhaustive()
    }
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

    const fn published(deleted_previews: usize, touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one hidden-axis transaction.
#[must_use = "a body-table hidden-axis commit contains the validated package snapshot"]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Formatting `package` or `patch` recursively would make this type's
        // redaction depend on every nested implementation.  Keep the public
        // commit representation to the already-safe diagnostics summary.
        formatter
            .debug_struct("Commit")
            .field("changed", &self.diagnostics.changed)
            .field("touched_components", &self.diagnostics.touched_components)
            .field("deleted_previews", &self.diagnostics.deleted_previews)
            .field(
                "full_reparse_performed",
                &self.diagnostics.full_reparse_performed,
            )
            .finish_non_exhaustive()
    }
}

impl Commit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body-table's user-hidden rows and columns.
    pub fn body_table_hidden_axes<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<HiddenAxes, Error> {
        let mut budget = transaction_budget(self)?;
        let graph = engine::resolve_graph(&Leaf::of(self), selector.into(), &mut budget)?;
        Ok(graph.before().clone())
    }

    /// Start a selector-first hidden-axis edit.
    pub fn edit_body_table_hidden_axes<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            selector: table_lock::OwnedBodyTableSelector::try_from(self, selector.into())
                .map_err(map_lock_error)?,
            target: None,
            before: None,
            axes: HiddenAxes::empty(),
            staged: false,
        })
    }

    /// Apply a patch only when both its exact source and semantic proof match.
    pub fn apply_body_table_hidden_axes(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut budget = transaction_budget(self)?;
        let source = self.state.source.source_bytes();
        if fingerprint(source) != patch.source_fingerprint || source != patch.source.as_ref() {
            return Err(Error::PatchConflict);
        }
        let graph = engine::resolve_graph(
            &Leaf::of(self),
            BodyTableSelector::index(patch.proof.table_position),
            &mut budget,
        )?;
        if graph.target().model_identifier != patch.proof.model_identifier
            || *graph.before() != patch.before
        {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        graph.admit_change()?;
        if !self.state.source.source_is_exact()
            || fingerprint(patch.target.as_ref()) != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        let candidate = reopen(self, Arc::clone(&patch.target), &mut budget)?;
        let verified = engine::resolve_graph(
            &Leaf::of(&candidate),
            BodyTableSelector::index(patch.proof.table_position),
            &mut budget,
        )?;
        if verified.target().model_identifier != patch.proof.model_identifier
            || *verified.before() != patch.after
        {
            return Err(Error::Verification);
        }
        engine::verify_replay(
            self,
            &candidate,
            &patch.proof,
            patch.touched_components,
            patch.added_object_ids.as_ref(),
            &mut budget,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.source_previews.saturating_sub(patch.target_previews),
                patch.touched_components,
            ),
        })
    }
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let target = edit.target.ok_or(Error::InvalidSource)?;
    let before = edit.before.ok_or(Error::InvalidSource)?;
    let mut budget = transaction_budget(source)?;
    let source_bytes = source.state.source.shared_source();
    let source_fingerprint = fingerprint(source_bytes.as_ref());
    let source_previews = engine::previews(source);
    if before == edit.axes {
        // Re-read against the captured target so an edit cannot silently
        // publish after its source graph changed underneath it.
        let graph = engine::resolve_graph(
            &Leaf::of(source),
            BodyTableSelector::index(target.table_position),
            &mut budget,
        )?;
        if graph.target().model_identifier != target.model_identifier || *graph.before() != before {
            return Err(Error::InvalidSource);
        }
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                source: Arc::clone(&source_bytes),
                target: source_bytes,
                source_fingerprint,
                target_fingerprint: source_fingerprint,
                proof: target,
                before: before.clone(),
                after: edit.axes,
                source_previews,
                target_previews: source_previews,
                touched_components: 0,
                added_object_ids: Arc::from(Vec::<u64>::new()),
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    if target.explicit_locked == Some(true) {
        return Err(Error::Refused(Refusal::TableLocked));
    }
    let graph = engine::resolve_graph(
        &Leaf::of(source),
        BodyTableSelector::index(target.table_position),
        &mut budget,
    )?;
    if *graph.before() != before {
        return Err(Error::InvalidSource);
    }
    let engine::Published {
        package: Leaf(candidate),
        touched_components,
        added_object_ids,
    } = engine::publish(&Leaf::of(source), &graph, &edit.axes, &mut budget)?;
    let target_owner = candidate.state.source.shared_source();
    let target_fingerprint = fingerprint(target_owner.as_ref());
    let target_previews = engine::previews(&candidate);
    Ok(Commit {
        package: candidate,
        patch: Patch {
            source: source_bytes,
            target: target_owner,
            source_fingerprint,
            target_fingerprint,
            proof: target,
            before,
            after: edit.axes,
            source_previews,
            target_previews,
            touched_components,
            added_object_ids,
        },
        diagnostics: Diagnostics::published(
            source_previews.saturating_sub(target_previews),
            touched_components,
        ),
    })
}

fn transaction_budget(package: &Package) -> Result<table_lock::WireBudget, Error> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    Ok(budget)
}

fn reopen(
    source: &Package,
    target: Arc<[u8]>,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, Error> {
    budget
        .charge_input_source(target.as_ref())
        .map_err(map_lock_error)?;
    let catalog = source
        .state
        .source
        .reopen(target)
        .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .map_err(map_lock_error)?;
    table_lock::charge_reopen_work(&catalog, budget).map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn map_lock_error(error: table_lock::Error) -> Error {
    match error {
        table_lock::Error::Refused(crate::table::lock::Refusal::TableNotFound) => {
            Error::Refused(Refusal::TableNotFound)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousTableName) => {
            Error::Refused(Refusal::AmbiguousTableName)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousSelector) => {
            Error::Refused(Refusal::AmbiguousSelector)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        table_lock::Error::InvalidSource => Error::InvalidSource,
        table_lock::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::Error::Allocation { amount } => Error::Allocation { amount },
        table_lock::Error::Verification => Error::Verification,
        table_lock::Error::PatchConflict => Error::PatchConflict,
    }
}

fn map_lock_limit(kind: table_lock::LimitKind) -> LimitKind {
    match kind {
        table_lock::LimitKind::InputBytes => LimitKind::InputBytes,
        table_lock::LimitKind::OutputBytes => LimitKind::OutputBytes,
        table_lock::LimitKind::Entries => LimitKind::Entries,
        table_lock::LimitKind::EntryBytes => LimitKind::EntryBytes,
        table_lock::LimitKind::TotalEntryBytes => LimitKind::TotalEntryBytes,
        table_lock::LimitKind::PackageBytes => LimitKind::PackageBytes,
        table_lock::LimitKind::PayloadBytes => LimitKind::PayloadBytes,
        table_lock::LimitKind::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        table_lock::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
        table_lock::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
        table_lock::LimitKind::PayloadItems => LimitKind::PayloadItems,
        table_lock::LimitKind::PayloadReferences => LimitKind::PayloadReferences,
        table_lock::LimitKind::WireBytes => LimitKind::WireBytes,
        table_lock::LimitKind::WireFields => LimitKind::WireFields,
        table_lock::LimitKind::WireNesting => LimitKind::WireNesting,
        table_lock::LimitKind::WireWork => LimitKind::WireWork,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
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
        _ => Error::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadObjects,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: limit as u64,
        },
        _ => Error::InvalidSource,
    }
}

// The shared engine reports the same failures; the conversion is lossless
// in both directions.
const fn engine_limit_kind(kind: engine::LimitKind) -> LimitKind {
    use engine::LimitKind as E;
    match kind {
        E::InputBytes => LimitKind::InputBytes,
        E::OutputBytes => LimitKind::OutputBytes,
        E::Entries => LimitKind::Entries,
        E::EntryBytes => LimitKind::EntryBytes,
        E::TotalEntryBytes => LimitKind::TotalEntryBytes,
        E::PackageBytes => LimitKind::PackageBytes,
        E::PayloadBytes => LimitKind::PayloadBytes,
        E::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        E::PayloadObjects => LimitKind::PayloadObjects,
        E::PayloadMessages => LimitKind::PayloadMessages,
        E::PayloadItems => LimitKind::PayloadItems,
        E::PayloadReferences => LimitKind::PayloadReferences,
        E::WireBytes => LimitKind::WireBytes,
        E::WireOutputBytes => LimitKind::WireOutputBytes,
        E::WireFields => LimitKind::WireFields,
        E::WireNesting => LimitKind::WireNesting,
        E::WireWork => LimitKind::WireWork,
        E::TransactionWork => LimitKind::TransactionWork,
    }
}

const fn leaf_limit_kind(kind: LimitKind) -> engine::LimitKind {
    use engine::LimitKind as E;
    match kind {
        LimitKind::InputBytes => E::InputBytes,
        LimitKind::OutputBytes => E::OutputBytes,
        LimitKind::Entries => E::Entries,
        LimitKind::EntryBytes => E::EntryBytes,
        LimitKind::TotalEntryBytes => E::TotalEntryBytes,
        LimitKind::PackageBytes => E::PackageBytes,
        LimitKind::PayloadBytes => E::PayloadBytes,
        LimitKind::TotalPayloadBytes => E::TotalPayloadBytes,
        LimitKind::PayloadObjects => E::PayloadObjects,
        LimitKind::PayloadMessages => E::PayloadMessages,
        LimitKind::PayloadItems => E::PayloadItems,
        LimitKind::PayloadReferences => E::PayloadReferences,
        LimitKind::WireBytes => E::WireBytes,
        LimitKind::WireOutputBytes => E::WireOutputBytes,
        LimitKind::WireFields => E::WireFields,
        LimitKind::WireNesting => E::WireNesting,
        LimitKind::WireWork => E::WireWork,
        LimitKind::TransactionWork => E::TransactionWork,
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        match error {
            engine::Error::Refused(refusal) => Self::Refused(match refusal {
                engine::Refusal::TableNotFound => Refusal::TableNotFound,
                engine::Refusal::AmbiguousTableName => Refusal::AmbiguousTableName,
                engine::Refusal::AmbiguousSelector => Refusal::AmbiguousSelector,
                engine::Refusal::UnsupportedSource => Refusal::UnsupportedSource,
                engine::Refusal::TableLocked => Refusal::TableLocked,
                engine::Refusal::UnsupportedDependency => Refusal::UnsupportedDependency,
            }),
            engine::Error::InvalidSource => Self::InvalidSource,
            engine::Error::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: engine_limit_kind(kind),
                observed,
                maximum,
            },
            engine::Error::Allocation { amount } => Self::Allocation { amount },
            engine::Error::Verification => Self::Verification,
            engine::Error::PatchConflict => Self::PatchConflict,
        }
    }
}

fn to_engine(error: Error) -> engine::Error {
    match error {
        Error::Refused(refusal) => engine::Error::Refused(match refusal {
            Refusal::TableNotFound => engine::Refusal::TableNotFound,
            Refusal::AmbiguousTableName => engine::Refusal::AmbiguousTableName,
            Refusal::AmbiguousSelector => engine::Refusal::AmbiguousSelector,
            Refusal::UnsupportedSource => engine::Refusal::UnsupportedSource,
            Refusal::TableLocked => engine::Refusal::TableLocked,
            Refusal::UnsupportedDependency => engine::Refusal::UnsupportedDependency,
        }),
        Error::InvalidSource => engine::Error::InvalidSource,
        Error::Limit {
            kind,
            observed,
            maximum,
        } => engine::Error::Limit {
            kind: leaf_limit_kind(kind),
            observed,
            maximum,
        },
        Error::Allocation { amount } => engine::Error::Allocation { amount },
        Error::Verification => engine::Error::Verification,
        Error::PatchConflict => engine::Error::PatchConflict,
    }
}

fn lock(error: table_lock::Error) -> engine::Error {
    to_engine(map_lock_error(error))
}

impl engine::Budget for table_lock::WireBudget {
    fn wire_limits(&self) -> WireLimits {
        Self::wire_limits(self)
    }
    fn remaining_wire_fields(&self) -> usize {
        Self::remaining_wire_fields(self)
    }
    fn remaining_wire_work(&self) -> usize {
        Self::remaining_wire_work(self)
    }
    fn maximum_payload_references(&self) -> usize {
        Self::maximum_payload_references(self)
    }
    fn charge_payload_work(&mut self, amount: usize) -> Result<(), engine::Error> {
        Self::charge_payload_work(self, amount).map_err(lock)
    }
    fn charge_payload_references(&mut self, amount: usize) -> Result<(), engine::Error> {
        Self::charge_payload_references(self, amount).map_err(lock)
    }
    fn charge_payload_objects(&mut self, amount: usize) -> Result<(), engine::Error> {
        Self::charge_payload_objects(self, amount).map_err(lock)
    }
    fn charge_payload_items(&mut self, amount: usize) -> Result<(), engine::Error> {
        Self::charge_payload_items(self, amount).map_err(lock)
    }
    fn charge_payload_messages(&mut self, amount: usize) -> Result<(), engine::Error> {
        Self::charge_payload_messages(self, amount).map_err(lock)
    }
    fn charge_payload_bytes(&mut self, amount: usize) -> Result<(), engine::Error> {
        Self::charge_payload_bytes(self, amount).map_err(lock)
    }
    fn charge_total_payload_bytes(&mut self, amount: usize) -> Result<(), engine::Error> {
        Self::charge_total_payload_bytes(self, amount).map_err(lock)
    }
    fn charge_output_bytes(&mut self, amount: usize) -> Result<(), engine::Error> {
        Self::charge_output_bytes(self, amount).map_err(lock)
    }
    fn charge_codec_report(
        &mut self,
        fields: usize,
        work_bytes: usize,
        max_depth: u32,
        references: usize,
    ) -> Result<(), engine::Error> {
        Self::charge_codec_report(self, fields, work_bytes, max_depth, references).map_err(lock)
    }
    fn parse<'a>(&mut self, source: &'a [u8], depth: usize) -> Result<WireView<'a>, engine::Error> {
        Self::parse(self, source, depth).map_err(lock)
    }
    fn precharge_candidate_reopen(
        &mut self,
        source: &litchi_iwa_archive::SourceCatalog,
        candidate_package_bound: usize,
        changed_component_index: usize,
        changed_entry_bound: usize,
        changed_archive_bound: usize,
        changed_object_index: usize,
        changed_message_index: usize,
        changed_message_bound: usize,
    ) -> Result<(), engine::Error> {
        Self::precharge_candidate_reopen(
            self,
            source,
            candidate_package_bound,
            changed_component_index,
            changed_entry_bound,
            changed_archive_bound,
            changed_object_index,
            changed_message_index,
            changed_message_bound,
        )
        .map_err(lock)
    }
}

/// The engine's view of a Pages package. A private wrapper keeps the
/// crate-private table-lock budget out of `Package`'s public trait impls.
#[derive(Clone)]
struct Leaf(Package);

impl Leaf {
    fn of(package: &Package) -> Self {
        Self(package.clone())
    }
}

impl litchi_iwa_table::TableSource for Leaf {
    fn catalog(&self) -> &litchi_iwa_archive::ComponentCatalog {
        litchi_iwa_table::TableSource::catalog(&self.0)
    }
    fn physical(&self) -> Option<&litchi_iwa_archive::SourceCatalog> {
        litchi_iwa_table::TableSource::physical(&self.0)
    }
    fn resolve_object(
        &self,
        identifier: u64,
    ) -> Result<Option<litchi_iwa_table::ResolvedObject<'_>>, litchi_iwa_table::ResolveError> {
        litchi_iwa_table::TableSource::resolve_object(&self.0, identifier)
    }
    fn source_bytes(&self) -> &[u8] {
        litchi_iwa_table::TableSource::source_bytes(&self.0)
    }
    fn archive_limits(&self) -> litchi_iwa_archive::Limits {
        litchi_iwa_table::TableSource::archive_limits(&self.0)
    }
    fn max_objects(&self) -> usize {
        litchi_iwa_table::TableSource::max_objects(&self.0)
    }
    fn max_references(&self) -> usize {
        litchi_iwa_table::TableSource::max_references(&self.0)
    }
}

impl engine::Leaf for Leaf {
    type Selector<'s> = BodyTableSelector<'s>;
    type Budget = table_lock::WireBudget;

    fn position_selector(_scope: usize, table: usize) -> Self::Selector<'static> {
        BodyTableSelector::index(table)
    }

    fn resolve_target(
        &self,
        selector: Self::Selector<'_>,
        budget: &mut Self::Budget,
    ) -> Result<engine::Target, engine::Error> {
        let target = self
            .0
            .resolve_body_table_with_budget(selector, budget)
            .map_err(lock)?;
        table_lock::validate_body_table_target(&self.0, &target, budget).map_err(lock)?;
        Ok(engine::Target {
            scope: 0,
            table_position: target.table_position,
            model_identifier: target.model_identifier,
            model_component_index: target.model_component_index,
            model_object_index: target.model_object_index,
            model_message_index: target.model_message_index,
            model_message_type: target.model_message_type,
            drawable_identifier: target.drawable_identifier,
            component_index: target.component_index,
            object_index: target.object_index,
            info_message_index: target.info_message_index,
            message_type: target.message_type,
            explicit_locked: target.explicit_locked,
        })
    }

    fn reopen_exact(&self, candidate: Arc<[u8]>) -> Result<Self, engine::Error> {
        let catalog = self
            .0
            .state
            .source
            .reopen(candidate)
            .map_err(|error| to_engine(map_archive_error(error)))?;
        Package::from_source_catalog(catalog)
            .map(Self)
            .map_err(|error| to_engine(map_package_error(error)))
    }
}
