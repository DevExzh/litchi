//! Exact-source hidden-row/hidden-column transactions for Keynote slide
//! tables (record 0379).
//!
//! The slide and table selectors resolve to the table's native route; the
//! shared engine (`litchi_iwa_table::hidden_axes`, the one the Pages owner
//! uses) proves the hidden-state graph and rewrites it. Untouched members
//! and objects keep their exact bytes; the patch keeps both exact artifacts,
//! so its inverse restores the source byte for byte.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_table::hidden_axes::{self as engine};
use thiserror::Error as ThisError;

use super::{Package, slide_table_core as core};
use crate::SlideSelector;
use crate::slide::table::TableSelector;
use crate::slide::table::hidden_axes::HiddenAxes;

/// Finite resources governed by one hidden-axis transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate package output bytes.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical entry.
    EntryBytes,
    /// Aggregate physical entry bytes.
    TotalEntryBytes,
    /// Physical package metadata bytes.
    PackageBytes,
    /// Decoded native payload bytes.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native framing items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict codec input bytes.
    WireBytes,
    /// Strict codec output bytes.
    WireOutputBytes,
    /// Strict codec fields inspected.
    WireFields,
    /// Strict codec nesting depth.
    WireNesting,
    /// Strict codec work.
    WireWork,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A redacted failure from a slide-table hidden-axis read or transaction.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`]
    /// classifies it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted table graph is malformed or ambiguous.
    #[error("the selected Keynote slide-table hidden-axis source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error(
        "Keynote slide-table hidden axes {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} units for Keynote slide-table hidden axes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the staged hidden axes.
    #[error("the edited Keynote slide-table hidden axes failed verification")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("the Keynote slide-table hidden-axis patch does not match the exact source")]
    PatchConflict,
}

/// A well-formed request this family refuses.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// No slide or slide table matches the selectors.
    #[error("no Keynote slide or slide table matches the selectors")]
    TableNotFound,
    /// A selector is ambiguous.
    #[error("the Keynote slide or table selector is ambiguous")]
    AmbiguousSelector,
    /// The source cannot publish an exact changed artifact.
    #[error("this Keynote source does not support exact slide-table hidden-axis editing")]
    UnsupportedSource,
    /// The table is locked.
    #[error("the selected Keynote slide table is locked")]
    TableLocked,
    /// A hidden-state dependency is outside the shared engine.
    #[error("the selected Keynote slide-table hidden-axis dependency is unsupported")]
    UnsupportedDependency,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound => RefusalKind::NotFound,
            Self::AmbiguousSelector => RefusalKind::InvalidValue,
            Self::UnsupportedSource | Self::UnsupportedDependency => {
                RefusalKind::UnsupportedContent
            },
            Self::TableLocked => RefusalKind::Protected,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

const fn limit_kind(kind: engine::LimitKind) -> LimitKind {
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

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        match error {
            engine::Error::Refused(refusal) => Self::Refused(match refusal {
                engine::Refusal::TableNotFound => Refusal::TableNotFound,
                engine::Refusal::AmbiguousTableName | engine::Refusal::AmbiguousSelector => {
                    Refusal::AmbiguousSelector
                },
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
                kind: limit_kind(kind),
                observed,
                maximum,
            },
            engine::Error::Allocation { amount } => Self::Allocation { amount },
            engine::Error::Verification => Self::Verification,
            engine::Error::PatchConflict => Self::PatchConflict,
        }
    }
}

fn select_error(error: core::Error) -> engine::Error {
    match error {
        core::Error::SlideNameNotFound
        | core::Error::SlidePositionNotFound(_)
        | core::Error::TablePositionNotFound(_)
        | core::Error::EmptySlideName => engine::Error::Refused(engine::Refusal::TableNotFound),
        core::Error::AmbiguousSelector => {
            engine::Error::Refused(engine::Refusal::AmbiguousSelector)
        },
        core::Error::UnsupportedSource => {
            engine::Error::Refused(engine::Refusal::UnsupportedSource)
        },
        core::Error::UnsupportedDependency | core::Error::UnsupportedTopology => {
            engine::Error::Refused(engine::Refusal::UnsupportedDependency)
        },
        core::Error::Limit {
            observed, maximum, ..
        } => engine::Error::Limit {
            kind: engine::LimitKind::TransactionWork,
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => engine::Error::Allocation { amount },
        _ => engine::Error::InvalidSource,
    }
}

/// The engine's view of a Keynote package: a private wrapper, so the
/// engine's selector and budget types stay out of `Package`'s public impls.
#[derive(Clone)]
struct Leaf(Package);

impl Leaf {
    fn of(package: &Package) -> Self {
        Self(package.snapshot())
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

/// A slide and a table on it, both by position.
#[derive(Clone, Copy)]
struct Selection<'s> {
    slide: SlideSelector<'s>,
    table: TableSelector,
}

impl engine::Leaf for Leaf {
    type Selector<'s> = Selection<'s>;
    type Budget = engine::WireBudget;

    const DELETES_ROOT_PREVIEWS: bool = false;

    fn position_selector(scope: usize, table: usize) -> Self::Selector<'static> {
        Selection {
            slide: SlideSelector::index(scope),
            table: TableSelector::index(table),
        }
    }

    fn resolve_target(
        &self,
        selector: Self::Selector<'_>,
        _budget: &mut Self::Budget,
    ) -> Result<engine::Target, engine::Error> {
        let mut budget = core::Budget::new(&self.0).map_err(select_error)?;
        let selected = core::select_table(&self.0, selector.slide, selector.table, &mut budget)
            .map_err(select_error)?;
        let components = self.0.state.source.components();
        let index_of = |name: &str| {
            components
                .iter()
                .position(|component| component.name() == name)
                .ok_or(engine::Error::InvalidSource)
        };
        let model_component_index = index_of(&selected.model.component)?;
        let component_index = index_of(&selected.table_info.component)?;
        let message_type = |component: usize, object: usize, message: usize| {
            components
                .get_index(component)
                .and_then(|component| component.archive().objects.get(object))
                .and_then(|object| object.messages.get(message))
                .map(|message| message.type_)
                .ok_or(engine::Error::InvalidSource)
        };
        Ok(engine::Target {
            scope: selected.slide_position.get(),
            table_position: selected.table_position.get(),
            model_identifier: std::num::NonZeroU64::new(selected.model.identifier)
                .ok_or(engine::Error::InvalidSource)?,
            model_component_index,
            model_object_index: selected.model.object_index,
            model_message_index: selected.model_message_index,
            model_message_type: message_type(
                model_component_index,
                selected.model.object_index,
                selected.model_message_index,
            )?,
            drawable_identifier: std::num::NonZeroU64::new(selected.table_info.identifier)
                .ok_or(engine::Error::InvalidSource)?,
            component_index,
            object_index: selected.table_info.object_index,
            info_message_index: selected.table_info_message_index,
            message_type: message_type(
                component_index,
                selected.table_info.object_index,
                selected.table_info_message_index,
            )?,
            explicit_locked: Some(selected.locked),
        })
    }

    fn reopen_exact(&self, candidate: Arc<[u8]>) -> Result<Self, engine::Error> {
        let package = Package::from_source_with_options(candidate, self.0.state.options)
            .map_err(|_| engine::Error::Verification)?;
        package
            .validate()
            .map_err(|_| engine::Error::Verification)?;
        Ok(Self(package))
    }
}

fn budget(package: &Package) -> Result<engine::WireBudget, Error> {
    let catalog =
        core::physical_catalog(package).map_err(|_| Error::Refused(Refusal::UnsupportedSource))?;
    let mut budget = engine::WireBudget::new(catalog.limits())?;
    budget.charge_source_catalog(catalog)?;
    Ok(budget)
}

fn graph(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    budget: &mut engine::WireBudget,
) -> Result<engine::Graph, Error> {
    Ok(engine::resolve_graph(
        &Leaf::of(package),
        Selection { slide, table },
        budget,
    )?)
}

fn artifact(package: &Package) -> Result<Arc<[u8]>, Error> {
    Ok(core::physical_catalog(package)
        .map_err(|_| Error::Refused(Refusal::UnsupportedSource))?
        .shared_source())
}

/// A reversible exact-source slide-table hidden-axis patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    proof: engine::Target,
    before: HiddenAxes,
    after: HiddenAxes,
    touched_components: usize,
    added_object_ids: Arc<[u64]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_count", &self.before.as_slice().len())
            .field("after_count", &self.after.as_slice().len())
            .field("touched_components", &self.touched_components)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The source hidden axes.
    #[must_use]
    pub const fn before(&self) -> &HiddenAxes {
        &self.before
    }

    /// The target hidden axes.
    #[must_use]
    pub const fn after(&self) -> &HiddenAxes {
        &self.after
    }

    /// Whether the patch is an exact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            proof: self.proof.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
            touched_components: self.touched_components,
            added_object_ids: Arc::clone(&self.added_object_ids),
        }
    }
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Rewritten package members.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
}

/// One validated slide-table hidden-axis publication.
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

    /// Consume this publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume this publication and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// One staged hidden-axis change of one slide table.
pub struct Edit<'a> {
    source: &'a Package,
    proof: engine::Target,
    before: HiddenAxes,
    after: HiddenAxes,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("before_count", &self.before.as_slice().len())
            .field("after_count", &self.after.as_slice().len())
            .finish_non_exhaustive()
    }
}

fn proof_selection(proof: &engine::Target) -> (SlideSelector<'static>, TableSelector) {
    (
        SlideSelector::index(proof.scope),
        TableSelector::index(proof.table_position),
    )
}

impl Edit<'_> {
    /// The hidden axes observed when the edit was created.
    #[must_use]
    pub const fn before(&self) -> &HiddenAxes {
        &self.before
    }

    /// The staged hidden axes.
    #[must_use]
    pub const fn hidden_axes(&self) -> &HiddenAxes {
        &self.after
    }

    /// Stage `axes` as the table's complete user-hidden rows and columns.
    #[must_use]
    pub fn set(mut self, axes: HiddenAxes) -> Self {
        self.after = axes;
        self
    }

    /// Stage no hidden rows or columns.
    #[must_use]
    pub fn clear(self) -> Self {
        self.set(HiddenAxes::empty())
    }

    /// Validate and atomically publish the staged hidden axes.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableLocked`], [`Refusal::UnsupportedDependency`] for a
    /// graph outside the engine, source, resource and verification failures.
    pub fn commit(self) -> Result<Commit, Error> {
        let source_artifact = artifact(self.source)?;
        if self.before == self.after {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_artifact), source_artifact),
                    proof: self.proof,
                    before: self.before,
                    after: self.after,
                    touched_components: 0,
                    added_object_ids: Arc::from(Vec::<u64>::new()),
                },
                diagnostics: Diagnostics::default(),
            });
        }
        if !core::physical_catalog(self.source)
            .map_err(|_| Error::Refused(Refusal::UnsupportedSource))?
            .source_is_exact()
        {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        let mut budget = budget(self.source)?;
        let (slide, table) = proof_selection(&self.proof);
        let graph = graph(self.source, slide, table, &mut budget)?;
        if *graph.before() != self.before
            || graph.target().model_identifier != self.proof.model_identifier
        {
            return Err(Error::PatchConflict);
        }
        let engine::Published {
            package: Leaf(candidate),
            touched_components,
            added_object_ids,
        } = engine::publish(&Leaf::of(self.source), &graph, &self.after, &mut budget)?;
        let target_artifact = artifact(&candidate)?;
        Ok(Commit {
            package: candidate,
            patch: Patch {
                artifacts: ExactArtifacts::new(source_artifact, target_artifact),
                proof: self.proof,
                before: self.before,
                after: self.after,
                touched_components,
                added_object_ids,
            },
            diagnostics: Diagnostics {
                changed: true,
                touched_components,
            },
        })
    }
}

impl Package {
    /// Read one slide table's user-hidden rows and columns.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::UnsupportedDependency`] for a
    /// hidden-state graph outside the engine, and source failures.
    pub fn slide_table_hidden_axes<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<HiddenAxes, Error> {
        let mut budget = budget(self)?;
        Ok(graph(self, slide.into(), table.into(), &mut budget)?
            .before()
            .clone())
    }

    /// Start a hidden-axis edit of one slide table.
    ///
    /// # Errors
    ///
    /// As [`Self::slide_table_hidden_axes`].
    pub fn edit_slide_table_hidden_axes<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = budget(self)?;
        let graph = graph(self, slide.into(), table.into(), &mut budget)?;
        Ok(Edit {
            source: self,
            proof: graph.target().clone(),
            before: graph.before().clone(),
            after: graph.before().clone(),
        })
    }

    /// Replay a slide-table hidden-axis patch on its exact source.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] on another artifact or a changed table;
    /// [`Error::Verification`] when the target does not reproduce the patch.
    pub fn apply_slide_table_hidden_axes(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_source(&artifact(self)?) {
            return Err(Error::PatchConflict);
        }
        let mut budget = budget(self)?;
        let (slide, table) = proof_selection(&patch.proof);
        let source_graph = graph(self, slide, table, &mut budget)?;
        if source_graph.target().model_identifier != patch.proof.model_identifier
            || *source_graph.before() != patch.before
        {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::default(),
            });
        }
        source_graph.admit_change()?;
        let Leaf(candidate) =
            engine::Leaf::reopen_exact(&Leaf::of(self), patch.artifacts.target())?;
        let verified = graph(&candidate, slide, table, &mut budget)?;
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
            diagnostics: Diagnostics {
                changed: true,
                touched_components: patch.touched_components,
            },
        })
    }
}
