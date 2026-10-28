//! Exact-source hidden-row/hidden-column transactions for Numbers tables
//! (record 0379).
//!
//! The sheet and table selectors resolve to the table's native route; the
//! shared engine (`litchi_iwa_table::hidden_axes`, the one the Pages and
//! Keynote owners use) proves the hidden-state graph and rewrites it.
//! Untouched members and objects keep their exact bytes; the patch keeps both
//! exact artifacts, so its inverse restores the source byte for byte.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_table::hidden_axes::{self as engine};
use thiserror::Error as ThisError;

use super::Package;
use crate::table::hidden_axes::HiddenAxes;
use crate::{SheetSelector, TableSelector};

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
    /// Logical output allocations reserved by one strict codec operation.
    WireAllocations,
    /// Retained bytes charged by one strict codec operation.
    WireRetainedBytes,
    /// Scratch bytes charged by one strict codec operation.
    WireScratchBytes,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A redacted failure from a table hidden-axis read or transaction.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`]
    /// classifies it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted table graph is malformed or ambiguous.
    #[error("the selected Numbers table hidden-axis source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error(
        "Numbers table hidden axes {kind} limit exceeded: observed {observed}, maximum {maximum}"
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
    #[error("could not allocate {amount} units for Numbers table hidden axes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the staged hidden axes.
    #[error("the edited Numbers table hidden axes failed verification")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("the Numbers table hidden-axis patch does not match the exact source")]
    PatchConflict,
}

/// A well-formed request this family refuses.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// No sheet or table matches the selectors.
    #[error("no Numbers sheet or table matches the selectors")]
    TableNotFound,
    /// A selector is ambiguous.
    #[error("the Numbers sheet or table selector is ambiguous")]
    AmbiguousSelector,
    /// The source cannot publish an exact changed artifact.
    #[error("this Numbers source does not support exact table hidden-axis editing")]
    UnsupportedSource,
    /// The table is locked.
    #[error("the selected Numbers table is locked")]
    TableLocked,
    /// A hidden-state dependency is outside the shared engine.
    #[error("the selected Numbers table hidden-axis dependency is unsupported")]
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

/// The engine's view of a Numbers package: a private wrapper, so the
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

/// A sheet and a table on it.
#[derive(Clone, Copy)]
struct Selection<'s> {
    sheet: SheetSelector<'s>,
    table: TableSelector<'s>,
}

impl engine::Leaf for Leaf {
    type Selector<'s> = Selection<'s>;
    type Budget = engine::WireBudget;

    // Numbers cell-level families keep the root previews (record 0370).
    const DELETES_ROOT_PREVIEWS: bool = false;

    fn position_selector(scope: usize, table: usize) -> Self::Selector<'static> {
        Selection {
            sheet: SheetSelector::index(scope),
            table: TableSelector::index(table),
        }
    }

    fn resolve_target(
        &self,
        selector: Self::Selector<'_>,
        _budget: &mut Self::Budget,
    ) -> Result<engine::Target, engine::Error> {
        let not_found = engine::Error::Refused(engine::Refusal::TableNotFound);
        let sheet = self
            .0
            .state
            .document
            .sheet(selector.sheet)
            .map_err(|_| engine::Error::InvalidSource)?
            .ok_or(not_found)?;
        let table = match selector.table {
            TableSelector::Position(index) => {
                let index = index.get();
                sheet.tables().nth(index).map(|_| index)
            },
            TableSelector::Name(name) => {
                let mut matches = sheet
                    .tables()
                    .enumerate()
                    .filter(|(_, table)| table.name() == name);
                let first = matches.next().map(|(index, _)| index);
                if matches.next().is_some() {
                    return Err(engine::Error::Refused(engine::Refusal::AmbiguousTableName));
                }
                first
            },
        }
        .ok_or(not_found)?;
        let target = super::table_headers::resolve::resolve_target(&self.0, sheet.index(), table)
            .map_err(map_header_error)?;
        Ok(engine::Target {
            scope: sheet.index(),
            table_position: table,
            model_identifier: std::num::NonZeroU64::new(target.model_identifier)
                .ok_or(engine::Error::InvalidSource)?,
            model_component_index: target.component_index,
            model_object_index: target.object_index,
            model_message_index: target.message_index,
            model_message_type: target.message_type,
            drawable_identifier: std::num::NonZeroU64::new(target.drawable_identifier)
                .ok_or(engine::Error::InvalidSource)?,
            component_index: target.info_component_index,
            object_index: target.info_object_index,
            info_message_index: target.info_message_index,
            message_type: target.info_message_type,
            explicit_locked: Some(target.locked == crate::table::lock::State::Locked),
        })
    }

    fn reopen_exact(&self, candidate: Arc<[u8]>) -> Result<Self, engine::Error> {
        self.0
            .reopen_candidate(candidate)
            .map(Self)
            .map_err(map_package_error)
    }
}

fn map_header_error(error: super::table_headers::Error) -> engine::Error {
    use super::table_headers::{Error as HeaderError, Refusal as HeaderRefusal};
    match error {
        HeaderError::Refused(refusal) => engine::Error::Refused(match refusal {
            HeaderRefusal::SheetNotFound | HeaderRefusal::TableNotFound => {
                engine::Refusal::TableNotFound
            },
            HeaderRefusal::TableLocked { .. } => engine::Refusal::TableLocked,
            HeaderRefusal::UnsupportedDependency { .. } => engine::Refusal::UnsupportedDependency,
            HeaderRefusal::UnsupportedSource => engine::Refusal::UnsupportedSource,
            HeaderRefusal::InvalidSettings { .. } => return engine::Error::InvalidSource,
        }),
        HeaderError::InvalidSource { .. } => engine::Error::InvalidSource,
        HeaderError::Limit {
            kind,
            observed,
            maximum,
            ..
        } => engine::Error::Limit {
            kind: map_header_limit_kind(kind),
            observed,
            maximum,
        },
        HeaderError::Allocation { amount, .. } => engine::Error::Allocation { amount },
        HeaderError::Verification => engine::Error::Verification,
        HeaderError::PatchConflict => engine::Error::PatchConflict,
    }
}

const fn map_header_limit_kind(kind: super::table_headers::LimitKind) -> engine::LimitKind {
    use super::table_headers::LimitKind as H;
    use engine::LimitKind as E;
    match kind {
        H::InputBytes => E::InputBytes,
        H::OutputBytes => E::OutputBytes,
        H::Entries => E::Entries,
        H::EntryBytes => E::EntryBytes,
        H::TotalEntryBytes => E::TotalEntryBytes,
        H::PackageBytes => E::PackageBytes,
        H::PayloadBytes => E::PayloadBytes,
        H::TotalPayloadBytes => E::TotalPayloadBytes,
        H::PayloadObjects => E::PayloadObjects,
        H::PayloadMessages => E::PayloadMessages,
        H::PayloadItems => E::PayloadItems,
        H::PayloadReferences => E::PayloadReferences,
        H::WireBytes => E::WireBytes,
        H::WireOutputBytes => E::WireOutputBytes,
        H::WireFields => E::WireFields,
        H::WireNesting => E::WireNesting,
        H::WireWork => E::WireWork,
        // The hidden-axis engine predates the table-info codec's distinct
        // allocation/retained/scratch ceilings and reports those legacy
        // adapter failures under its aggregate wire-work budget.  The
        // focused comments and table-header routes retain the dimensions
        // before this unrelated adapter boundary.
        H::WireAllocations | H::WireRetainedBytes | H::WireScratchBytes => E::WireWork,
        H::TransactionWork => E::TransactionWork,
    }
}

fn map_package_error(error: crate::Error) -> engine::Error {
    match error {
        crate::Error::Archive(error) => map_archive_error(error),
        crate::Error::Common(error) => map_common_error(error),
        crate::Error::Semantic(error) => map_workbook_error(error),
        crate::Error::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => engine::Error::Limit {
            kind: map_semantic_limit_kind(kind),
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        crate::Error::InputTooLarge { observed, maximum } => engine::Error::Limit {
            kind: engine::LimitKind::InputBytes,
            observed,
            maximum,
        },
        _ => engine::Error::InvalidSource,
    }
}

fn map_workbook_error(error: crate::workbook::Error) -> engine::Error {
    match error {
        crate::workbook::Error::TooManySheets { actual, limit }
        | crate::workbook::Error::TooManyTables { actual, limit }
        | crate::workbook::Error::TooManyMaterializedCells { actual, limit } => {
            engine::Error::Limit {
                kind: engine::LimitKind::PayloadItems,
                observed: u64::try_from(actual).unwrap_or(u64::MAX),
                maximum: u64::try_from(limit).unwrap_or(u64::MAX),
            }
        },
        crate::workbook::Error::TextTooLarge { observed, limit } => engine::Error::Limit {
            kind: engine::LimitKind::PayloadBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(limit).unwrap_or(u64::MAX),
        },
        crate::workbook::Error::Allocation { amount } => engine::Error::Allocation { amount },
        _ => engine::Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> engine::Error {
    match error {
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            engine::Error::Allocation { amount }
        },
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => engine::Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => engine::LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => engine::LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => engine::LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => engine::LimitKind::PackageBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => engine::LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => engine::LimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => engine::LimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    engine::LimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => engine::Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> engine::Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => engine::Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => engine::LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    engine::LimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => engine::LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => engine::LimitKind::PayloadItems,
                _ => engine::LimitKind::PayloadBytes,
            },
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            engine::Error::Allocation { amount: requested }
        },
        _ => engine::Error::InvalidSource,
    }
}

fn map_common_error(error: litchi_iwa_common::Error) -> engine::Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => engine::Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => engine::LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::Fields => engine::LimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => engine::LimitKind::WireOutputBytes,
                litchi_iwa_common::LimitKind::Nesting => engine::LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => engine::LimitKind::WireWork,
                _ => engine::LimitKind::PayloadItems,
            },
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(limit).unwrap_or(u64::MAX),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => engine::Error::Allocation { amount },
        _ => engine::Error::InvalidSource,
    }
}

const fn map_semantic_limit_kind(kind: crate::SemanticLimitKind) -> engine::LimitKind {
    use crate::SemanticLimitKind as S;
    use engine::LimitKind as E;
    match kind {
        S::Objects => E::PayloadObjects,
        S::References => E::PayloadReferences,
        S::FormulaRenderDepth | S::FormulaDepth => E::WireNesting,
        S::FormulaRenderWork | S::FormulaWork => E::WireWork,
        S::OutputTextBytes | S::FormulaCategoryBytes | S::TextBytes => E::PayloadBytes,
        S::Sheets | S::Tables | S::MaterializedCells => E::PayloadItems,
    }
}

fn physical(package: &Package) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    package
        .state
        .components
        .physical()
        .ok_or(Error::Refused(Refusal::UnsupportedSource))
}

fn budget(package: &Package) -> Result<engine::WireBudget, Error> {
    let catalog = physical(package)?;
    let mut budget = engine::WireBudget::new(catalog.limits())?;
    budget.charge_source_catalog(catalog)?;
    Ok(budget)
}

fn graph(
    package: &Package,
    sheet: SheetSelector<'_>,
    table: TableSelector<'_>,
    budget: &mut engine::WireBudget,
) -> Result<engine::Graph, Error> {
    Ok(engine::resolve_graph(
        &Leaf::of(package),
        Selection { sheet, table },
        budget,
    )?)
}

fn artifact(package: &Package) -> Result<Arc<[u8]>, Error> {
    Ok(physical(package)?.shared_source())
}

/// A reversible exact-source table hidden-axis patch.
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

/// One validated table hidden-axis publication.
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

/// One staged hidden-axis change of one table.
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

fn proof_selection(proof: &engine::Target) -> (SheetSelector<'static>, TableSelector<'static>) {
    (
        SheetSelector::index(proof.scope),
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
        if !physical(self.source)?.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        let mut budget = budget(self.source)?;
        let (sheet, table) = proof_selection(&self.proof);
        let graph = graph(self.source, sheet, table, &mut budget)?;
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
    /// Read one table's user-hidden rows and columns.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::UnsupportedDependency`] for a
    /// hidden-state graph outside the engine, and source failures.
    pub fn table_hidden_axes<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<HiddenAxes, Error> {
        let mut budget = budget(self)?;
        Ok(graph(self, sheet.into(), table.into(), &mut budget)?
            .before()
            .clone())
    }

    /// Start a hidden-axis edit of one table.
    ///
    /// # Errors
    ///
    /// As [`Self::table_hidden_axes`].
    pub fn edit_table_hidden_axes<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = budget(self)?;
        let graph = graph(self, sheet.into(), table.into(), &mut budget)?;
        Ok(Edit {
            source: self,
            proof: graph.target().clone(),
            before: graph.before().clone(),
            after: graph.before().clone(),
        })
    }

    /// Replay a table hidden-axis patch on its exact source.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] on another artifact or a changed table;
    /// [`Error::Verification`] when the target does not reproduce the patch.
    pub fn apply_table_hidden_axes(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_source(&artifact(self)?) {
            return Err(Error::PatchConflict);
        }
        let mut budget = budget(self)?;
        let (sheet, table) = proof_selection(&patch.proof);
        let source_graph = graph(self, sheet, table, &mut budget)?;
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
        let verified = graph(&candidate, sheet, table, &mut budget)?;
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
