//! Selector-first Keynote slide-chart data-symbol transactions (D7-A).
//!
//! The shared W1-E series graph owns the symbol mark/size split.  This leaf
//! adapter supplies only Keynote chart selection, semantic `Symbol` values,
//! exact-source publication, candidate reopen/readback, and typed refusal
//! mapping.

#![allow(
    clippy::cast_sign_loss,
    clippy::needless_pass_by_value,
    reason = "The focused adapter keeps the semantic boundary beside its staged transaction."
)]

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_core::{Archive, Limits as ArchiveLimits};
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions};
use litchi_iwa_protos::chart_style_codec;
use litchi_iwa_style::graph::{self, Components, Workspace};
use litchi_iwa_style::series as engine;
use thiserror::Error as ThisError;

use super::chart_axis_support::{self, AxisSupportError, AxisSupportSelectorError};
use super::slide_chart_font::{self, FieldLedger};
use super::slide_table_core::{Budget as TransactionBudget, Error as TransactionError};
use super::{Package, PhysicalSource};
use crate::{ChartSelector, SlideSelector};
use litchi_core::Position;
use litchi_iwa_common::chart::series::{Symbol, SymbolFamily};

/// The one per-series symbol property owned by this slice.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// A symbol per visible series; `None` hides that series' marker.
    Symbols(Vec<Option<Symbol>>),
}

impl Property {
    /// Return the number of series named by this property.
    #[must_use]
    pub fn len(&self) -> usize {
        match self {
            Self::Symbols(values) => values.len(),
        }
    }

    /// Whether this property names no series.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

/// Finite resource dimensions exposed by the focused route.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Wire input bytes.
    InputBytes,
    /// Wire output bytes.
    OutputBytes,
    /// Bytes in one protobuf payload or archive member.
    WireBytes,
    /// Package entries, objects and messages.
    Entries,
    /// Graph references.
    References,
    /// Parsed wire fields.
    WireFields,
    /// Maximum protobuf nesting depth.
    WireNesting,
    /// Aggregate graph work.
    WireWork,
    /// Graph allocation events.
    Allocations,
    /// Retained graph bytes.
    RetainedBytes,
    /// Temporary graph scratch bytes.
    ScratchBytes,
    /// Caller bytes retained by deferred selectors.
    SelectorBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::SelectorBytes => "selector bytes",
        })
    }
}

/// A well-formed request this route refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The source is not an exact physical package.
    #[error("this Keynote source does not support chart-series edits")]
    UnsupportedSource,
    /// The selector did not resolve one slide/chart.
    #[error("the selected Keynote chart was not found")]
    ChartNotFound,
    /// An exact-name selector was ambiguous.
    #[error("the Keynote chart-series selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked semantic chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The chart kind has no symbol family.
    #[error("the selected chart kind has no data-symbol family")]
    UnsupportedKind,
    /// The value does not name the chart's visible series count.
    #[error("chart series value count does not match the chart")]
    SeriesCount,
    /// Commit was requested without a staged value.
    #[error("the Keynote chart-series transaction has no staged operation")]
    NoStagedOperation,
    /// The graph is valid but outside the focused ownership profile.
    #[error("the selected chart-series graph is outside the focused route")]
    UnsupportedDependency,
}

impl Refusal {
    /// Map this semantic refusal to the unified refusal taxonomy.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource | Self::UnsupportedDependency => {
                RefusalKind::UnsupportedContent
            },
            Self::ChartNotFound
            | Self::SlideNameNotFound
            | Self::SlidePositionNotFound { .. }
            | Self::ChartNameNotFound
            | Self::ChartPositionNotFound { .. } => RefusalKind::NotFound,
            Self::AmbiguousSelector | Self::EmptySlideName | Self::EmptyChartName => {
                RefusalKind::InvalidValue
            },
            Self::UnsupportedKind | Self::SeriesCount => RefusalKind::InvalidValue,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
        }
    }
}

/// A content-redacted chart-series failure.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted graph is malformed.
    #[error("the Keynote chart-series source is invalid")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart-series {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource dimension.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} Keynote chart-series units")]
    Allocation { amount: usize },
    /// A finite finalization calculation overflowed before allocation.
    #[error("Keynote chart-series finalization arithmetic overflow")]
    Overflow,
    /// The reopened candidate failed semantic verification.
    #[error("the Keynote chart-series candidate failed verification")]
    Verification,
    /// The patch was created from another exact source.
    #[error("the Keynote chart-series patch does not match the exact source")]
    PatchConflict,
}

/// Compact evidence describing one chart-series commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    /// Whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Whether the candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Selection {
    slide: Position,
    chart: Position,
    identifier: u64,
}

fn select(
    package: &Package,
    slide: SlideSelector<'_>,
    chart: ChartSelector<'_>,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
) -> Result<Selection, Error> {
    let selected = slide_chart_font::select_chart_for_series(package, slide, chart, budget, fields)
        .map_err(map_axis_support_error)?;
    Ok(Selection {
        slide: selected.slide_position,
        chart: selected.chart_position,
        identifier: selected.chart_identifier,
    })
}

fn map_axis_support_error(error: AxisSupportError) -> Error {
    match error {
        AxisSupportError::Selector(selector) => match selector {
            AxisSupportSelectorError::UnsupportedSource => Refusal::UnsupportedSource.into(),
            AxisSupportSelectorError::AmbiguousSelector => Refusal::AmbiguousSelector.into(),
            AxisSupportSelectorError::EmptySlideName => Refusal::EmptySlideName.into(),
            AxisSupportSelectorError::SlideNameNotFound => Refusal::SlideNameNotFound.into(),
            AxisSupportSelectorError::SlidePositionNotFound { position } => {
                Refusal::SlidePositionNotFound { position }.into()
            },
            AxisSupportSelectorError::ChartNameNotFound => Refusal::ChartNameNotFound.into(),
            AxisSupportSelectorError::ChartPositionNotFound { position } => {
                Refusal::ChartPositionNotFound { position }.into()
            },
            AxisSupportSelectorError::EmptyChartName => Refusal::EmptyChartName.into(),
        },
        AxisSupportError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_axis_support::AxisSupportLimitKind::InputBytes => LimitKind::InputBytes,
                chart_axis_support::AxisSupportLimitKind::OutputBytes => LimitKind::OutputBytes,
                chart_axis_support::AxisSupportLimitKind::WireBytes => LimitKind::WireBytes,
                chart_axis_support::AxisSupportLimitKind::Entries
                | chart_axis_support::AxisSupportLimitKind::EntryBytes
                | chart_axis_support::AxisSupportLimitKind::TotalBytes
                | chart_axis_support::AxisSupportLimitKind::Slides
                | chart_axis_support::AxisSupportLimitKind::TextStorages
                | chart_axis_support::AxisSupportLimitKind::TextFragments => LimitKind::Entries,
                chart_axis_support::AxisSupportLimitKind::References => LimitKind::References,
                chart_axis_support::AxisSupportLimitKind::TextBytes => LimitKind::WireBytes,
                chart_axis_support::AxisSupportLimitKind::WireFields => LimitKind::WireFields,
                chart_axis_support::AxisSupportLimitKind::WireNesting => LimitKind::WireNesting,
                chart_axis_support::AxisSupportLimitKind::WireWork => LimitKind::WireWork,
                chart_axis_support::AxisSupportLimitKind::TitleBytes => LimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        AxisSupportError::Allocation { amount } => Error::Allocation { amount },
        AxisSupportError::InvalidSource => Error::InvalidSource,
    }
}

fn family_and_count(
    package: &Package,
    selection: Selection,
    budget: &mut TransactionBudget,
) -> Result<(SymbolFamily, usize), Error> {
    let payload = package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| {
            component
                .archive()
                .object(selection.identifier)
                .and_then(|object| {
                    object
                        .messages
                        .iter()
                        .find(|message| message.type_ == 5_021)
                        .map(|message| message.data.as_slice())
                })
        })
        .ok_or(Error::InvalidSource)?;
    let limits = budget.residual(package).map_err(map_transaction_error)?;
    let maximum_depth = u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?;
    let options = DecodeOptions::new(
        payload.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields().max(1),
        limits.max_rewrite_work().max(1),
        maximum_depth,
        limits.max_fields().clamp(1, 1_000_000),
        limits.max_input_bytes().clamp(1, 64 * 1024 * 1024),
    );
    let (snapshot, report) =
        match chart_metadata_codec::decode_modern_with_report(payload, &options) {
            Ok(decoded) => decoded,
            Err(error) => {
                charge_metadata_report(budget, error.report())?;
                return Err(map_chart_metadata_decode_error(error));
            },
        };
    charge_metadata_report(budget, report)?;
    let kind = litchi_iwa_common::chart::kind::Kind::from_native(snapshot.chart_type());
    let family = SymbolFamily::for_chart_kind(kind).ok_or(Refusal::UnsupportedKind)?;
    Ok((family, snapshot.series_count()))
}

fn charge_metadata_report(
    budget: &mut TransactionBudget,
    report: chart_metadata_codec::DecodeReport,
) -> Result<(), Error> {
    budget
        .input(report.source_bytes())
        .and_then(|_| budget.fields(report.fields()))
        .and_then(|_| budget.work(report.work_bytes()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.nesting(report.max_depth() as usize))
        .map_err(map_transaction_error)
}

fn map_chart_metadata_decode_error(error: chart_metadata_codec::DecodeError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.input_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.depth_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireNesting,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

struct LeafComponents<'a> {
    components: Vec<(&'a str, &'a Archive)>,
}

impl<'a> LeafComponents<'a> {
    fn view_requirements(package: &Package) -> Result<(usize, usize), Error> {
        let count = package.state.source.components().len();
        let bytes = count
            .checked_mul(size_of::<(&str, &Archive)>())
            .and_then(|value| value.checked_add(size_of::<Self>()))
            .ok_or(Error::InvalidSource)?;
        Ok((usize::from(count != 0).saturating_add(1), bytes))
    }

    fn new(package: &'a Package) -> Result<Self, Error> {
        let count = package.state.source.components().len();
        let mut components = Vec::new();
        components
            .try_reserve_exact(count)
            .map_err(|_| Error::Allocation { amount: count })?;
        for index in 0..count {
            let component = package
                .state
                .source
                .components()
                .get_index(index)
                .ok_or(Error::InvalidSource)?;
            components.push((component.name(), component.archive()));
        }
        Ok(Self { components })
    }
}

impl Components for LeafComponents<'_> {
    fn count(&self) -> usize {
        self.components.len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.components.get(index).copied()
    }
}

fn graph_limits(package: &Package) -> Result<ArchiveLimits, Error> {
    package
        .structure_archive_limits()
        .map_err(|_| Error::InvalidSource)
}

fn transaction_budget(package: &Package) -> Result<TransactionBudget, Error> {
    let mut budget = TransactionBudget::new(package).map_err(map_transaction_error)?;
    budget
        .preflight_semantic_scan(package)
        .map_err(map_transaction_error)?;
    budget
        .inventory_components(package)
        .map_err(map_transaction_error)?;
    Ok(budget)
}

fn transaction_fields(package: &Package) -> Result<FieldLedger, Error> {
    slide_chart_font::field_ledger_for_series(package).map_err(map_axis_support_error)
}

struct WorkspaceAdmission<'a> {
    components: Box<LeafComponents<'a>>,
    requirements: graph::ConstructionRequirements,
    limits: graph::ConstructionLimits,
}

fn map_construction(error: graph::ConstructionError) -> Error {
    match error {
        graph::ConstructionError::Graph(error) => map_graph(error),
        graph::ConstructionError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                graph::ConstructionLimitKind::Work => LimitKind::WireWork,
                graph::ConstructionLimitKind::Allocations => LimitKind::Allocations,
                graph::ConstructionLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                graph::ConstructionLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => return Error::InvalidSource,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        graph::ConstructionError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn admit_workspace<'a>(
    package: &'a Package,
    budget: &mut TransactionBudget,
) -> Result<WorkspaceAdmission<'a>, Error> {
    let components = package.state.source.components();
    let component_count = components.len();
    let object_count = package.state.total_objects;
    let (view_allocations, view_retained) = LeafComponents::view_requirements(package)?;
    budget
        .allocations(view_allocations)
        .and_then(|_| budget.retained(view_retained))
        .map_err(map_transaction_error)?;
    let components_view = Box::new(LeafComponents::new(package)?);
    let construction_limits = budget
        .construction_limits()
        .map_err(map_transaction_error)?;
    let requirements = Workspace::construction_requirements(&*components_view, construction_limits)
        .map_err(map_construction)?;
    budget
        .components(component_count)
        .and_then(|_| budget.payload_objects(object_count))
        .and_then(|_| budget.charge_construction(requirements, 2))
        .map_err(map_transaction_error)
        .map(|()| WorkspaceAdmission {
            components: components_view,
            requirements,
            limits: construction_limits,
        })
}

fn map_transaction_error(error: TransactionError) -> Error {
    match error {
        TransactionError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                super::slide_table_core::LimitKind::InputBytes => LimitKind::InputBytes,
                super::slide_table_core::LimitKind::OutputBytes => LimitKind::OutputBytes,
                super::slide_table_core::LimitKind::WireFields => LimitKind::WireFields,
                super::slide_table_core::LimitKind::WireWork => LimitKind::WireWork,
                super::slide_table_core::LimitKind::Allocations => LimitKind::Allocations,
                super::slide_table_core::LimitKind::Retained => LimitKind::RetainedBytes,
                super::slide_table_core::LimitKind::Scratch => LimitKind::ScratchBytes,
                super::slide_table_core::LimitKind::WireNesting => LimitKind::WireNesting,
                super::slide_table_core::LimitKind::References => LimitKind::References,
                super::slide_table_core::LimitKind::Entries
                | super::slide_table_core::LimitKind::EntryBytes
                | super::slide_table_core::LimitKind::TotalBytes
                | super::slide_table_core::LimitKind::Components
                | super::slide_table_core::LimitKind::PayloadObjects
                | super::slide_table_core::LimitKind::PayloadMessages => LimitKind::Entries,
            },
            observed,
            maximum,
        },
        TransactionError::Allocation(amount) => Error::Allocation { amount },
        TransactionError::UnsupportedSource => Refusal::UnsupportedSource.into(),
        TransactionError::UnsupportedDependency => Refusal::UnsupportedDependency.into(),
        TransactionError::InvalidSource
        | TransactionError::UnsupportedTopology
        | TransactionError::AmbiguousSelector
        | TransactionError::EmptySlideName
        | TransactionError::SlideNameNotFound
        | TransactionError::SlidePositionNotFound(_)
        | TransactionError::TablePositionNotFound(_)
        | TransactionError::Read
        | TransactionError::Wire
        | TransactionError::Codec
        | TransactionError::Archive
        | TransactionError::Verification => Error::InvalidSource,
    }
}

fn map_selector_error(error: super::edit::SelectorError) -> Error {
    match error {
        super::edit::SelectorError::Limit { observed, maximum } => Error::Limit {
            kind: LimitKind::SelectorBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        super::edit::SelectorError::Allocation { amount } => Error::Allocation { amount },
    }
}

fn map_graph(error: graph::Error) -> Error {
    match error {
        graph::Error::Limit {
            kind: graph::GraphLimitKind::ScratchBytes,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::ScratchBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        graph::Error::Limit {
            kind: graph::GraphLimitKind::AllocationEvents,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Allocations,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        graph::Error::Archive(error) => match error {
            litchi_iwa_core::Error::Allocation { requested, .. } => {
                Error::Allocation { amount: requested }
            },
            litchi_iwa_core::Error::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: map_core_limit(kind),
                observed: observed as u64,
                maximum: maximum as u64,
            },
            _ => Error::InvalidSource,
        },
        graph::Error::Registry(error) => map_registry_error(error),
        graph::Error::Wire(error) => map_wire_error(error),
        graph::Error::Codec(error) => map_codec_error(error),
        graph::Error::MissingObject(_)
        | graph::Error::AmbiguousObject(_)
        | graph::Error::IdentifierOverflow => Error::InvalidSource,
        graph::Error::Invalid("a series style is shared by several charts") => {
            Refusal::UnsupportedDependency.into()
        },
        graph::Error::Invalid(_) => Error::InvalidSource,
        _ => Error::InvalidSource,
    }
}

fn map_finish_error(error: graph::FinishError) -> Error {
    match error {
        graph::FinishError::Graph(error) => map_graph(error),
        graph::FinishError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                graph::FinishLimitKind::Work => LimitKind::WireWork,
                graph::FinishLimitKind::Allocations => LimitKind::Allocations,
                graph::FinishLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                graph::FinishLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => return Error::InvalidSource,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        graph::FinishError::Allocation { amount } => Error::Allocation { amount },
        graph::FinishError::Overflow(_) => Error::Overflow,
        _ => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::Entries,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_codec_error(error: chart_style_codec::Error) -> Error {
    match error {
        chart_style_codec::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_style_codec::LimitKind::Bytes => LimitKind::WireBytes,
                chart_style_codec::LimitKind::Fields => LimitKind::WireFields,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        chart_style_codec::Error::Allocation { bytes } => Error::Allocation { amount: bytes },
        _ => Error::InvalidSource,
    }
}

fn map_registry_error(error: litchi_iwa_protos::package_metadata_codec::RewriteError) -> Error {
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
                observed,
                maximum,
            } => (LimitKind::InputBytes, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
                observed,
                maximum,
            } => (LimitKind::OutputBytes, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields {
                observed,
                maximum,
            } => (LimitKind::WireFields, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (LimitKind::WireWork, observed, maximum)
            },
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting {
                observed,
                maximum,
            } => (LimitKind::WireNesting, observed as usize, maximum as usize),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
                observed,
                maximum,
            } => (LimitKind::Entries, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
                observed,
                maximum,
            } => (LimitKind::References, observed, maximum),
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

fn map_core_limit(kind: litchi_iwa_core::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_core::LimitKind::ArchiveBytes => LimitKind::Entries,
        litchi_iwa_core::LimitKind::Objects
        | litchi_iwa_core::LimitKind::Messages
        | litchi_iwa_core::LimitKind::MessagesPerObject
        | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
        litchi_iwa_core::LimitKind::ObjectBytes
        | litchi_iwa_core::LimitKind::MessageBytes
        | litchi_iwa_core::LimitKind::HeaderBytes
        | litchi_iwa_core::LimitKind::HeaderMemoryBytes
        | litchi_iwa_core::LimitKind::SnappyChunkBytes
        | litchi_iwa_core::LimitKind::SnappyStreamBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => LimitKind::WireBytes,
        litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
        litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
        litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::WireWork,
    }
}

const fn is_series_rewritable_type(message_type: u32) -> bool {
    matches!(message_type, 401 | 5_021 | 5_028 | 5_029 | 11_006 | 11_008)
}

fn map_font_error(error: slide_chart_font::Error) -> Error {
    match error {
        slide_chart_font::Error::Refused(slide_chart_font::Refusal::UnsupportedArchiveInfo) => {
            Refusal::UnsupportedDependency.into()
        },
        slide_chart_font::Error::Refused(_) => Refusal::UnsupportedSource.into(),
        slide_chart_font::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                slide_chart_font::LimitKind::InputBytes => LimitKind::InputBytes,
                slide_chart_font::LimitKind::OutputBytes => LimitKind::OutputBytes,
                slide_chart_font::LimitKind::WireBytes => LimitKind::WireBytes,
                slide_chart_font::LimitKind::Entries
                | slide_chart_font::LimitKind::EntryBytes
                | slide_chart_font::LimitKind::TotalBytes
                | slide_chart_font::LimitKind::Slides
                | slide_chart_font::LimitKind::Identifiers => LimitKind::Entries,
                slide_chart_font::LimitKind::References => LimitKind::References,
                slide_chart_font::LimitKind::WireFields => LimitKind::WireFields,
                slide_chart_font::LimitKind::WireNesting => LimitKind::WireNesting,
                slide_chart_font::LimitKind::WireWork => LimitKind::WireWork,
                slide_chart_font::LimitKind::Allocations => LimitKind::Allocations,
                slide_chart_font::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                slide_chart_font::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                slide_chart_font::LimitKind::SelectorBytes => LimitKind::SelectorBytes,
                slide_chart_font::LimitKind::TextBytes => LimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        slide_chart_font::Error::Allocation { amount } => Error::Allocation { amount },
        slide_chart_font::Error::InvalidSource => Error::InvalidSource,
        slide_chart_font::Error::Verification => Error::Verification,
        slide_chart_font::Error::PatchConflict => Error::PatchConflict,
    }
}

fn map_series_error(error: engine::SymbolError) -> Error {
    match error {
        engine::SymbolError::Graph(error) => map_graph(error),
        engine::SymbolError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                engine::LimitKind::InputBytes => LimitKind::InputBytes,
                engine::LimitKind::Fields => LimitKind::WireFields,
                engine::LimitKind::OutputBytes => LimitKind::OutputBytes,
                engine::LimitKind::Nesting => LimitKind::WireNesting,
                engine::LimitKind::Work => LimitKind::WireWork,
                engine::LimitKind::Allocations => LimitKind::Allocations,
                engine::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                engine::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                engine::LimitKind::References | engine::LimitKind::RegistryEntries => {
                    LimitKind::References
                },
                engine::LimitKind::Identifiers => LimitKind::Entries,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        engine::SymbolError::Allocation { amount } => Error::Allocation { amount },
        engine::SymbolError::Invalid("a series style is shared by several charts") => {
            Refusal::UnsupportedDependency.into()
        },
        engine::SymbolError::Invalid(_) => Error::InvalidSource,
        engine::SymbolError::IdentifierOverflow => Refusal::UnsupportedDependency.into(),
        _ => Error::InvalidSource,
    }
}

fn series_limits(
    package: &Package,
    budget: &TransactionBudget,
    fields: &FieldLedger,
) -> Result<engine::Limits, Error> {
    const CHART_CODEC_MAX_BYTES: usize = 256 * 1024 * 1024;
    const CHART_CODEC_MAX_FIELDS: usize = litchi_iwa_common::WireLimits::MAX_FIELDS;
    let wire = budget.residual(package).map_err(map_transaction_error)?;
    let max_fields = fields
        .remaining()
        .map_err(map_font_error)?
        .min(wire.max_fields())
        .clamp(1, CHART_CODEC_MAX_FIELDS);
    let max_input_bytes = wire.max_input_bytes().clamp(1, CHART_CODEC_MAX_BYTES);
    let max_output_bytes = wire.max_output_bytes().clamp(1, CHART_CODEC_MAX_BYTES);
    let references = budget
        .remaining_references()
        .map_err(map_transaction_error)?;
    Ok(engine::Limits::new(
        max_input_bytes,
        max_fields,
        max_output_bytes,
        u32::try_from(wire.max_nesting()).map_err(|_| Error::InvalidSource)?,
        wire.max_rewrite_work(),
        budget
            .remaining_allocations()
            .map_err(map_transaction_error)?,
        budget.remaining_retained().map_err(map_transaction_error)?,
        budget.remaining_scratch().map_err(map_transaction_error)?,
        references,
        references,
        references,
    ))
}

fn charge_series_report(
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
    report: engine::Report,
) -> Result<(), Error> {
    let references = report
        .references()
        .checked_add(report.registry_entries())
        .and_then(|value| value.checked_add(report.identifiers()))
        .ok_or(Error::InvalidSource)?;
    budget
        .input(report.input_bytes())
        .and_then(|_| budget.output(report.output_bytes()))
        .and_then(|_| budget.work(report.work_bytes()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .and_then(|_| budget.references(references))
        .and_then(|_| budget.nesting(report.max_depth() as usize))
        .map_err(map_transaction_error)?;
    fields.charge(report.fields()).map_err(map_font_error)
}

fn finish_workspace(
    workspace: Workspace<'_>,
    budget: &mut TransactionBudget,
) -> Result<graph::Changes, Error> {
    let requirements = workspace.finish_requirements().map_err(map_finish_error)?;
    let limits = graph::FinishLimits::new(
        if requirements.work() == 0 {
            0
        } else {
            budget.remaining_work().map_err(map_transaction_error)?
        },
        if requirements.allocations() == 0 {
            0
        } else {
            budget
                .remaining_allocations()
                .map_err(map_transaction_error)?
        },
        if requirements.retained_bytes() == 0 {
            0
        } else {
            budget.remaining_retained().map_err(map_transaction_error)?
        },
        if requirements.scratch_bytes() == 0 {
            0
        } else {
            budget.remaining_scratch().map_err(map_transaction_error)?
        },
    );
    for (kind, observed, maximum) in [
        (LimitKind::WireWork, requirements.work(), limits.max_work()),
        (
            LimitKind::Allocations,
            requirements.allocations(),
            limits.max_allocations(),
        ),
        (
            LimitKind::RetainedBytes,
            requirements.retained_bytes(),
            limits.max_retained_bytes(),
        ),
        (
            LimitKind::ScratchBytes,
            requirements.scratch_bytes(),
            limits.max_scratch_bytes(),
        ),
    ] {
        if observed > maximum {
            return Err(Error::Limit {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
    }
    let (changes, report) = workspace.finish_bounded(limits).map_err(map_finish_error)?;
    // Component/index admission is charged by `admit_workspace`; this report
    // covers only the final vectors, copied names, and their traversal.
    budget
        .output(report.output_bytes())
        .and_then(|_| budget.work(report.work()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .map_err(map_transaction_error)?;
    Ok(changes)
}

fn publish_changes(
    source: &Package,
    changes: graph::Changes,
    selected_chart: u64,
    scope: &engine::MutationScope,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
) -> Result<(Package, usize, slide_chart_font::MutationPlan), Error> {
    let touched_components = changes
        .edits
        .len()
        .checked_add(changes.insertions.len())
        .and_then(|count| count.checked_add(changes.deletions.len()))
        .ok_or(Error::InvalidSource)?;
    let extra_deletions = match &source.state.source {
        PhysicalSource::Package(catalog) => {
            super::rendering_invalidation::root_preview_deletions(catalog.package())
                .map_err(|_| Error::InvalidSource)?
        },
        PhysicalSource::Semantic(_) => return Err(Refusal::UnsupportedSource.into()),
    };
    let plan = slide_chart_font::mutation_plan_for_series_scope(
        source,
        selected_chart,
        &changes,
        scope,
        extra_deletions.names(),
        budget,
    )
    .map_err(map_font_error)?;
    let candidate = slide_chart_font::publish_with_policy(
        source,
        changes,
        budget,
        fields,
        &plan,
        extra_deletions.names(),
        is_series_rewritable_type,
        &[17, 19],
    )
    .map_err(map_font_error)?;
    Ok((candidate, touched_components, plan))
}

/// A compact exact-source patch for one chart-series symbol value.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    slide: Position,
    chart: Position,
    before: Property,
    after: Property,
    plan: slide_chart_font::MutationPlan,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide", &self.slide)
            .field("chart", &self.chart)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch is an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide: self.slide,
            chart: self.chart,
            before: self.after.clone(),
            after: self.before.clone(),
            plan: slide_chart_font::inverse_mutation_plan(&self.plan),
        }
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide
    }

    /// Return the selected semantic chart position.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart
    }

    /// Borrow the source property.
    #[must_use]
    pub const fn before(&self) -> &Property {
        &self.before
    }

    /// Borrow the target property.
    #[must_use]
    pub const fn after(&self) -> &Property {
        &self.after
    }

    /// Return the exact-source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the exact-target fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
}

/// One verified chart-series commit.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the candidate package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Borrow the exact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit and return its package and exact patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Consume this commit.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
}

/// One immutable selector-first chart-series edit.
pub struct Edit<'a> {
    source: &'a Package,
    slide: super::edit::OwnedSlideSelector,
    chart: super::edit::OwnedChartSelector,
    staged: Option<Property>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged", &self.staged.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn new<'s, 'c>(
        source: &'a Package,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
    ) -> Result<Self, Error> {
        let mut retained = 0;
        let maximum = source.semantic_limits().max_text_bytes();
        let slide =
            super::edit::OwnedSlideSelector::from_selector(slide.into(), maximum, &mut retained)
                .map_err(map_selector_error)?;
        let chart =
            super::edit::OwnedChartSelector::from_selector(chart.into(), maximum, &mut retained)
                .map_err(map_selector_error)?;
        Ok(Self {
            source,
            slide,
            chart,
            staged: None,
        })
    }

    /// Stage all per-series symbol values.
    pub fn set(&mut self, symbols: Vec<Option<Symbol>>) -> Result<&mut Self, Error> {
        if symbols.is_empty() {
            return Err(Refusal::SeriesCount.into());
        }
        self.staged = Some(Property::Symbols(symbols));
        Ok(self)
    }

    /// Commit the staged symbol value.
    pub fn commit(self) -> Result<Commit, Error> {
        let staged = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let mut budget = transaction_budget(self.source)?;
        let mut fields = transaction_fields(self.source)?;
        let selection = select(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
            &mut budget,
            &mut fields,
        )?;
        let (family, series) = family_and_count(self.source, selection, &mut budget)?;
        if staged.len() != series {
            return Err(Refusal::SeriesCount.into());
        }
        let before = read_symbols(
            self.source,
            selection,
            family,
            series,
            &mut budget,
            &mut fields,
        )?;
        let Property::Symbols(staged_values) = &staged;
        if before.as_slice() == staged_values.as_slice() {
            let source = physical_source(self.source)?;
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source), source),
                    slide: selection.slide,
                    chart: selection.chart,
                    before: Property::Symbols(before.clone()),
                    after: staged,
                    plan: slide_chart_font::MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !self.source.structure_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        budget
            .preflight_semantic_scan(self.source)
            .map_err(map_transaction_error)?;
        let admission = admit_workspace(self.source, &mut budget)?;
        let WorkspaceAdmission {
            components,
            requirements,
            limits,
        } = admission;
        let mut workspace = Workspace::new_bounded(
            &*components,
            graph_limits(self.source)?,
            requirements,
            limits,
        )
        .map_err(map_construction)?;
        let limits = series_limits(self.source, &budget, &fields)?;
        let (changed, report) = engine::set_symbols(
            &mut workspace,
            selection.identifier,
            family,
            match &staged {
                Property::Symbols(values) => values,
            },
            limits,
        )
        .map_err(map_series_error)?;
        let scope = report.scope().cloned().ok_or(Error::Verification)?;
        charge_series_report(&mut budget, &mut fields, report)?;
        if !changed {
            return Err(Error::InvalidSource);
        }
        let (candidate, touched_components, plan) = publish_changes(
            self.source,
            finish_workspace(workspace, &mut budget)?,
            selection.identifier,
            &scope,
            &mut budget,
            &mut fields,
        )?;
        let candidate_selection = select(
            &candidate,
            self.slide.as_selector(),
            self.chart.as_selector(),
            &mut budget,
            &mut fields,
        )?;
        let (candidate_family, candidate_series) =
            family_and_count(&candidate, candidate_selection, &mut budget)?;
        if candidate_family != family || candidate_series != series {
            return Err(Error::Verification);
        }
        let after = read_symbols(
            &candidate,
            candidate_selection,
            family,
            series,
            &mut budget,
            &mut fields,
        )?;
        if after.as_slice() != staged_values.as_slice() {
            return Err(Error::Verification);
        }
        let source = physical_source(self.source)?;
        let target = physical_source(&candidate)?;
        Ok(Commit {
            package: candidate,
            patch: Patch {
                artifacts: ExactArtifacts::new(source, Arc::clone(&target)),
                slide: selection.slide,
                chart: selection.chart,
                before: Property::Symbols(before),
                after: Property::Symbols(after),
                plan,
            },
            diagnostics: Diagnostics::published(touched_components),
        })
    }
}

fn read_symbols(
    package: &Package,
    selection: Selection,
    family: SymbolFamily,
    series: usize,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
) -> Result<Vec<Option<Symbol>>, Error> {
    let admission = admit_workspace(package, budget)?;
    let WorkspaceAdmission {
        components,
        requirements,
        limits,
    } = admission;
    let workspace =
        Workspace::new_bounded(&*components, graph_limits(package)?, requirements, limits)
            .map_err(map_construction)?;
    let limits = series_limits(package, budget, fields)?;
    let (values, report) =
        engine::symbols(&workspace, selection.identifier, family, series, limits)
            .map_err(map_series_error)?;
    charge_series_report(budget, fields, report)?;
    Ok(values)
}

fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(catalog) => Ok(catalog.shared_source()),
        PhysicalSource::Semantic(_) => Err(Refusal::UnsupportedSource.into()),
    }
}

impl Package {
    /// Read one chart's data symbols, one optional value per visible series.
    pub fn slide_chart_series<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
    ) -> Result<Property, Error> {
        let mut budget = transaction_budget(self)?;
        let mut fields = transaction_fields(self)?;
        let selection = select(self, slide.into(), chart.into(), &mut budget, &mut fields)?;
        let (family, series) = family_and_count(self, selection, &mut budget)?;
        Ok(Property::Symbols(read_symbols(
            self,
            selection,
            family,
            series,
            &mut budget,
            &mut fields,
        )?))
    }

    /// Start an exact-source chart-series symbol edit.
    pub fn edit_slide_chart_series<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(self, slide, chart)
    }

    /// Apply a patch only to its exact source package.
    pub fn apply_slide_chart_series(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut budget = transaction_budget(self)?;
        let mut fields = transaction_fields(self)?;
        let source = physical_source(self)?;
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
        let target = patch.artifacts.target().to_vec();
        budget
            .preflight_candidate(self, target.len())
            .map_err(map_transaction_error)?;
        let candidate = self
            .structure_reopen(target)
            .map_err(|_| Error::Verification)?;
        slide_chart_font::verify_locality_members(
            self,
            &candidate,
            &patch.plan,
            &mut budget,
            &mut fields,
            is_series_rewritable_type,
            &[17, 19],
        )
        .map_err(map_font_error)?;
        let selection = select(
            &candidate,
            SlideSelector::position(patch.slide),
            ChartSelector::position(patch.chart),
            &mut budget,
            &mut fields,
        )?;
        let (family, series) = family_and_count(&candidate, selection, &mut budget)?;
        let after = read_symbols(
            &candidate,
            selection,
            family,
            series,
            &mut budget,
            &mut fields,
        )?;
        let Property::Symbols(expected) = &patch.after;
        if after.as_slice() != expected.as_slice() {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(0),
        })
    }
}
