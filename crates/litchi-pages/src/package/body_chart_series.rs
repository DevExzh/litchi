//! Selector-first Pages body-chart series-symbol transactions.
//!
//! This adapter owns only Pages selection, bounded admission and exact-source
//! publication.  The series override graph (including symbol visibility,
//! shape and point-size inheritance) remains in `litchi_iwa_style`.

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{Entry, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::series::{Symbol, SymbolFamily};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage};
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions};
use litchi_iwa_style::graph::{self as style_graph, Components, Workspace};
use litchi_iwa_style::series as series_engine;
use thiserror::Error as ThisError;

use super::Package;
use super::body_chart_arrangement::{self as arrangement, ArrangementBudget, ChartTarget};
use super::body_drawable_graph as package_graph;
use crate::selector::BodyChartSelector;

/// A finite resource charged by one Pages body-chart series-symbol operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// Native payload bytes.
    PayloadBytes,
    /// Native references.
    PayloadReferences,
    /// Wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Wire and graph work.
    WireWork,
    /// Bounded allocations.
    Allocations,
    /// Retained source/candidate bytes.
    RetainedBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
    /// Physical archive resource.
    Archive(litchi_iwa_core::LimitKind),
    /// Package ingress/reassembly resource.
    Package(litchi_iwa_archive::LimitKind),
    /// Shared wire resource.
    Wire(litchi_iwa_common::LimitKind),
    /// Chart-style codec resource.
    Codec(litchi_iwa_protos::chart_style_codec::LimitKind),
    /// Package-metadata registry resource.
    Registry(litchi_iwa_protos::package_metadata_codec::RewriteLimit),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InputBytes => formatter.write_str("input bytes"),
            Self::OutputBytes => formatter.write_str("output bytes"),
            Self::PayloadBytes => formatter.write_str("payload bytes"),
            Self::PayloadReferences => formatter.write_str("payload references"),
            Self::WireFields => formatter.write_str("wire fields"),
            Self::WireNesting => formatter.write_str("wire nesting"),
            Self::WireWork => formatter.write_str("wire work"),
            Self::Allocations => formatter.write_str("allocations"),
            Self::RetainedBytes => formatter.write_str("retained bytes"),
            Self::ScratchBytes => formatter.write_str("scratch bytes"),
            Self::Archive(kind) => write!(formatter, "archive {kind}"),
            Self::Package(kind) => write!(formatter, "package {kind:?}"),
            Self::Wire(kind) => write!(formatter, "wire {kind}"),
            Self::Codec(kind) => write!(formatter, "codec {kind:?}"),
            Self::Registry(kind) => write!(formatter, "registry {kind:?}"),
        }
    }
}

/// A well-formed symbol request rejected without host fallback.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No chart matched the checked body position.
    #[error("the Pages body has no chart at position {position:?}")]
    ChartNotFound { position: Position },
    /// The source is not an exact package artifact.
    #[error("this Pages source does not support exact body-chart series edits")]
    UnsupportedSource,
    /// The chart kind has no native series-symbol family.
    #[error("the selected chart kind does not support series symbols")]
    UnsupportedKind,
    /// The symbol vector does not match the chart's visible series count.
    #[error("the series-symbol vector has {given} values; expected {expected}")]
    SeriesCount { expected: usize, given: usize },
    /// A style object is shared by more than one chart.
    #[error("the selected chart series style is shared by another chart")]
    SharedStyle,
    /// No operation was staged.
    #[error("no body-chart series-symbol change is staged")]
    NoStagedOperation,
    /// The patch belongs to another semantic property family.
    #[error("the patch does not belong to body-chart series symbols")]
    WrongProperty,
    /// The engine produced a publication shape this route does not own.
    #[error("the series-symbol graph produced unsupported publication changes")]
    UnsupportedPublication,
}

impl Refusal {
    /// Classify this refusal for the unified-operations boundary.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::ChartNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedSource
            | Self::UnsupportedKind
            | Self::SeriesCount { .. }
            | Self::WrongProperty
            | Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::SharedStyle | Self::UnsupportedPublication => RefusalKind::UnsupportedContent,
        }
    }
}

/// Failure from a Pages body-chart series-symbol read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A typed terminal refusal.
    #[error(transparent)]
    Refused(Refusal),
    /// The selected chart graph or metadata is malformed.
    #[error("the selected Pages body-chart series-symbol source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body-chart series symbols {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} Pages body-chart series-symbol units")]
    Allocation { amount: usize },
    /// Candidate semantic readback did not reproduce the requested vector.
    #[error("the edited Pages body-chart series symbols failed semantic verification")]
    Verification,
    /// The patch was created from another exact source artifact.
    #[error("the Pages body-chart series-symbol patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(value: Refusal) -> Self {
        Self::Refused(value)
    }
}

/// The only supported series-symbol property.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// One optional symbol value per visible chart series; `None` hides it.
    Symbols(Vec<Option<Symbol>>),
}

impl Property {
    /// Return the property kind.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        PropertyKind::Symbols
    }

    /// Return the number of values supplied.
    #[must_use]
    pub fn len(&self) -> usize {
        match self {
            Self::Symbols(values) => values.len(),
        }
    }

    /// Whether the property contains no series values.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

/// The supported body-chart series-symbol property kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// Symbol visibility, shape and optional point size.
    Symbols,
}

impl PropertyKind {
    /// Return the shared symbol property kind.
    #[must_use]
    pub const fn kind(self) -> Self {
        self
    }
}

#[derive(Clone, PartialEq)]
struct Target {
    chart: ChartTarget,
    kind: ChartKind,
    series: usize,
    family: SymbolFamily,
}

struct PackageComponents<'a>(&'a Package);

impl Components for PackageComponents<'_> {
    fn count(&self) -> usize {
        self.0.state.source.components().len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.0
            .state
            .source
            .components()
            .get_index(index)
            .map(|component| (component.name(), component.archive()))
    }
}

fn map_arrangement_limit(kind: arrangement::LimitKind) -> LimitKind {
    match kind {
        arrangement::LimitKind::InputBytes
        | arrangement::LimitKind::PayloadBytes
        | arrangement::LimitKind::TotalPayloadBytes => LimitKind::InputBytes,
        arrangement::LimitKind::PayloadReferences => LimitKind::PayloadReferences,
        arrangement::LimitKind::WireBytes | arrangement::LimitKind::WireOutputBytes => {
            LimitKind::InputBytes
        },
        arrangement::LimitKind::WireFields
        | arrangement::LimitKind::PayloadItems
        | arrangement::LimitKind::PayloadMessages
        | arrangement::LimitKind::PayloadObjects => LimitKind::WireFields,
        arrangement::LimitKind::WireNesting => LimitKind::WireNesting,
        arrangement::LimitKind::WireWork => LimitKind::WireWork,
        arrangement::LimitKind::WireAllocations => LimitKind::Allocations,
        arrangement::LimitKind::WireRetainedBytes => LimitKind::RetainedBytes,
        arrangement::LimitKind::WireScratchBytes => LimitKind::ScratchBytes,
        arrangement::LimitKind::OutputBytes
        | arrangement::LimitKind::Entries
        | arrangement::LimitKind::EntryBytes
        | arrangement::LimitKind::TotalEntryBytes
        | arrangement::LimitKind::PackageBytes => LimitKind::OutputBytes,
    }
}

fn map_arrangement_error(error: arrangement::Error) -> Error {
    match error {
        arrangement::Error::Refused(arrangement::Refusal::ChartNotFound { position })
        | arrangement::Error::Refused(arrangement::Refusal::ChartPositionNotFound { position }) => {
            Refusal::ChartNotFound { position }.into()
        },
        arrangement::Error::Refused(arrangement::Refusal::UnsupportedSource) => {
            Refusal::UnsupportedSource.into()
        },
        arrangement::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_arrangement_limit(kind),
            observed,
            maximum,
        },
        arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        arrangement::Error::InvalidSource
        | arrangement::Error::Verification
        | arrangement::Error::PatchConflict => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: LimitKind::Package(kind),
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Archive(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_registry_error(error: litchi_iwa_protos::package_metadata_codec::RewriteError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (observed, maximum) = match limit {
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work { observed, maximum }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
                observed,
                maximum,
            } => (observed as u64, maximum as u64),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting {
                observed,
                maximum,
            } => (u64::from(observed), u64::from(maximum)),
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind: LimitKind::Registry(limit),
            observed,
            maximum,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    Error::InvalidSource
}

fn map_graph_error(error: style_graph::Error) -> Error {
    match error {
        style_graph::Error::Archive(error) => map_core_error(error),
        style_graph::Error::Wire(error) => match error {
            litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit,
            } => Error::Limit {
                kind: LimitKind::Wire(kind),
                observed: observed as u64,
                maximum: limit as u64,
            },
            litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
            _ => Error::InvalidSource,
        },
        style_graph::Error::Codec(error) => match error {
            litchi_iwa_protos::chart_style_codec::Error::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: LimitKind::Codec(kind),
                observed: observed as u64,
                maximum: maximum as u64,
            },
            litchi_iwa_protos::chart_style_codec::Error::Allocation { bytes } => {
                Error::Allocation { amount: bytes }
            },
            _ => Error::InvalidSource,
        },
        style_graph::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                style_graph::GraphLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                style_graph::GraphLimitKind::AllocationEvents => LimitKind::Allocations,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        style_graph::Error::Registry(error) => map_registry_error(error),
        style_graph::Error::IdentifierOverflow => Refusal::UnsupportedPublication.into(),
        style_graph::Error::Invalid("a series style is shared by several charts") => {
            Refusal::SharedStyle.into()
        },
        _ => Error::InvalidSource,
    }
}

fn map_package_graph_error(error: package_graph::GraphError) -> Error {
    match error {
        package_graph::GraphError::InvalidSource => Error::InvalidSource,
        package_graph::GraphError::Allocation { amount } => Error::Allocation { amount },
        package_graph::GraphError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                package_graph::GraphLimit::InputBytes => LimitKind::InputBytes,
                package_graph::GraphLimit::OutputBytes => LimitKind::OutputBytes,
                package_graph::GraphLimit::WireFields => LimitKind::WireFields,
                package_graph::GraphLimit::WireNesting => LimitKind::WireNesting,
                package_graph::GraphLimit::WireWork => LimitKind::WireWork,
                package_graph::GraphLimit::SemanticItems => LimitKind::PayloadBytes,
                package_graph::GraphLimit::SemanticTextBytes => LimitKind::PayloadBytes,
            },
            observed,
            maximum,
        },
    }
}

fn archive_limits(package: &Package) -> Result<litchi_iwa_core::Limits, Error> {
    package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)
}

fn target_with_budget(
    package: &Package,
    selector: BodyChartSelector,
    budget: &mut ArrangementBudget,
) -> Result<ChartTarget, Error> {
    arrangement::resolve_target(package, selector, budget).map_err(map_arrangement_error)
}

fn admit_workspace(package: &Package, budget: &mut ArrangementBudget) -> Result<(), Error> {
    let source_bytes = package.state.source.shared_source().len();
    let objects = package.state.object_count.max(1);
    if !budget.workspace_source_admitted {
        budget.input(source_bytes).map_err(map_arrangement_error)?;
        budget
            .retained(source_bytes)
            .map_err(map_arrangement_error)?;
        budget.workspace_source_admitted = true;
    }
    budget.work(objects).map_err(map_arrangement_error)?;
    budget.allocations(objects).map_err(map_arrangement_error)?;
    budget
        .scratch(objects.saturating_mul(size_of::<u64>() * 2))
        .map_err(map_arrangement_error)
}

fn metadata_options(budget: &ArrangementBudget, source: &[u8]) -> Result<DecodeOptions, Error> {
    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    Ok(DecodeOptions::new(
        source.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields().max(1),
        limits.max_rewrite_work().max(1),
        u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?,
        limits.max_fields().clamp(1, 1_000_000),
        limits.max_input_bytes().clamp(1, 64 * 1024 * 1024),
    ))
}

fn charge_metadata_report(
    budget: &mut ArrangementBudget,
    report: chart_metadata_codec::DecodeReport,
) -> Result<(), Error> {
    budget
        .input(report.source_bytes())
        .map_err(map_arrangement_error)?;
    budget
        .fields(report.fields())
        .map_err(map_arrangement_error)?;
    budget
        .work(report.work_bytes())
        .map_err(map_arrangement_error)?;
    budget
        .allocations(report.allocations())
        .map_err(map_arrangement_error)?;
    budget
        .retained(report.retained_bytes())
        .map_err(map_arrangement_error)
}

fn map_metadata_error(error: chart_metadata_codec::DecodeError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.input_limit_values() {
        return Error::Limit {
            kind: LimitKind::InputBytes,
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

fn series_count(kind: ChartKind, components: usize) -> Result<usize, Error> {
    let count = match kind {
        ChartKind::Scatter2d if components == 2 => 1,
        ChartKind::Bubble2d if components == 3 => 1,
        ChartKind::Scatter2d | ChartKind::Bubble2d => return Err(Refusal::UnsupportedKind.into()),
        ChartKind::MultiDataScatter2d | ChartKind::MultiDataBubble2d => 1,
        _ => components,
    };
    if count == 0 {
        return Err(Error::InvalidSource);
    }
    Ok(count)
}

fn target_with_metadata(
    package: &Package,
    selector: BodyChartSelector,
    budget: &mut ArrangementBudget,
) -> Result<Target, Error> {
    let chart = target_with_budget(package, selector, budget)?;
    let source =
        arrangement::selected_chart_payload(package, &chart).map_err(map_arrangement_error)?;
    let options = metadata_options(budget, source)?;
    let (snapshot, report) = match chart_metadata_codec::decode_modern_with_report(source, &options)
    {
        Ok(decoded) => decoded,
        Err(error) => {
            charge_metadata_report(budget, error.report())?;
            return Err(map_metadata_error(error));
        },
    };
    charge_metadata_report(budget, report)?;
    let kind = ChartKind::from_native(snapshot.chart_type());
    let family = SymbolFamily::for_chart_kind(kind).ok_or(Refusal::UnsupportedKind)?;
    let series = series_count(kind, snapshot.series_count())?;
    Ok(Target {
        chart,
        kind,
        series,
        family,
    })
}

fn read_symbols(
    package: &Package,
    target: &Target,
    budget: &mut ArrangementBudget,
) -> Result<Vec<Option<Symbol>>, Error> {
    admit_workspace(package, budget)?;
    let components = PackageComponents(package);
    preflight_series_engine(package, target.series, budget)?;
    let limits = archive_limits(package)?;
    let workspace = Workspace::new(&components, limits).map_err(map_graph_error)?;
    let values = series_engine::symbols(
        &workspace,
        target.chart.drawable_identifier.get(),
        target.family,
        target.series,
    )
    .map_err(map_graph_error)?;
    charge_series_engine(package, target.series, budget)?;
    budget
        .work(target.series.saturating_mul(size_of::<Symbol>()))
        .map_err(map_arrangement_error)?;
    budget
        .allocations(usize::from(!values.is_empty()))
        .map_err(map_arrangement_error)?;
    budget
        .retained(values.len().saturating_mul(size_of::<Option<Symbol>>()))
        .map_err(map_arrangement_error)?;
    Ok(values)
}

fn archive_rewrite_bound(package: &Package, archive: &Archive) -> Result<usize, Error> {
    archive
        .encoded_len(archive_limits(package)?)
        .map_err(map_core_error)
}

fn preflight_publication(
    source: &Package,
    edits: &[(String, Archive)],
    insertions: &[(String, Vec<u8>)],
    budget: &ArrangementBudget,
) -> Result<usize, Error> {
    let mut projected = source.state.source.shared_source().len();
    for (name, archive) in edits {
        let bound = archive_rewrite_bound(source, archive)?;
        projected = projected
            .checked_add(name.len())
            .and_then(|value| value.checked_add(bound))
            .ok_or(Error::InvalidSource)?;
    }
    for (name, data) in insertions {
        projected = projected
            .checked_add(name.len())
            .and_then(|value| value.checked_add(data.len()))
            .ok_or(Error::InvalidSource)?;
    }
    budget
        .preflight_output(projected)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_work(projected)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_allocations(
            edits
                .len()
                .saturating_add(insertions.len())
                .saturating_add(1),
        )
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(projected)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_scratch(projected)
        .map_err(map_arrangement_error)?;
    Ok(projected)
}

fn charge_publication(
    candidate: &Package,
    edit_count: usize,
    insertion_count: usize,
    projected: usize,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let actual = candidate.state.source.shared_source().len();
    if actual > projected {
        return Err(Error::Verification);
    }
    budget.output(actual).map_err(map_arrangement_error)?;
    budget.work(actual).map_err(map_arrangement_error)?;
    budget
        .allocations(edit_count.saturating_add(insertion_count).saturating_add(1))
        .map_err(map_arrangement_error)?;
    budget.retained(actual).map_err(map_arrangement_error)?;
    budget.scratch(projected).map_err(map_arrangement_error)
}

fn preflight_series_engine(
    package: &Package,
    series: usize,
    budget: &ArrangementBudget,
) -> Result<(), Error> {
    // `Workspace` and the shared graph engine take borrowed physical limits,
    // not the Pages transaction ledger.  Admit a conservative bounded census
    // before constructing them so graph scans and symbol-vector capacity
    // cannot run ahead of the caller's residual budget.
    let source_bytes = package.state.source.shared_source().len();
    let objects = package.state.object_count.max(1);
    budget
        .preflight_work(source_bytes.saturating_mul(16).saturating_add(objects))
        .map_err(map_arrangement_error)?;
    budget
        .preflight_allocations(objects.saturating_add(series.saturating_mul(4)))
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(
            source_bytes.saturating_add(series.saturating_mul(size_of::<Option<Symbol>>())),
        )
        .map_err(map_arrangement_error)?;
    budget
        .preflight_scratch(source_bytes.saturating_mul(2))
        .map_err(map_arrangement_error)
}

fn charge_series_engine(
    package: &Package,
    series: usize,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let source_bytes = package.state.source.shared_source().len();
    let objects = package.state.object_count.max(1);
    budget
        .work(source_bytes.saturating_mul(16).saturating_add(objects))
        .map_err(map_arrangement_error)?;
    budget
        .allocations(objects.saturating_add(series.saturating_mul(4)))
        .map_err(map_arrangement_error)?;
    budget
        .retained(source_bytes.saturating_add(series.saturating_mul(size_of::<Option<Symbol>>())))
        .map_err(map_arrangement_error)?;
    budget
        .scratch(source_bytes.saturating_mul(2))
        .map_err(map_arrangement_error)
}

fn changed_object_ids(
    source: &Package,
    candidate: &Package,
    changed_components: &[Arc<str>],
) -> Result<Vec<(Arc<str>, Vec<u64>)>, Error> {
    let mut result = Vec::new();
    for name in changed_components {
        let Some(before) = source.state.source.components().get(name.as_ref()) else {
            continue;
        };
        let Some(after) = candidate.state.source.components().get(name.as_ref()) else {
            continue;
        };
        let mut changed = Vec::new();
        for object in &before.archive().objects {
            let identifier = object.archive_info.identifier.ok_or(Error::Verification)?;
            let differs = after
                .archive()
                .object(identifier)
                .is_none_or(|other| !object.same_content_ignoring_offsets(other));
            if differs {
                changed.push(identifier);
            }
        }
        for object in &after.archive().objects {
            let identifier = object.archive_info.identifier.ok_or(Error::Verification)?;
            if before.archive().object(identifier).is_none() {
                changed.push(identifier);
            }
        }
        result.push((Arc::clone(name), changed));
    }
    Ok(result)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    changed_components: &[Arc<str>],
    changed_objects: &[(Arc<str>, Vec<u64>)],
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let changed = |name: &str| changed_components.iter().any(|item| item.as_ref() == name);
    let source_entries = source.state.source.package();
    let candidate_entries = candidate.state.source.package();
    for before in source_entries.iter() {
        let after = candidate_entries
            .iter()
            .find(|item| item.name() == before.name());
        let Some(after) = after else {
            if changed(before.name()) {
                continue;
            }
            return Err(Error::Verification);
        };
        budget
            .work(entry_work(before)?.saturating_add(entry_work(after)?))
            .map_err(map_arrangement_error)?;
        if before.name() != after.name()
            || before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || before.metadata().local() != after.metadata().local()
            || before.metadata().central() != after.metadata().central()
        {
            return Err(Error::Verification);
        }
        if changed(before.name()) {
            if !zip_local_header_preserved(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
            ) || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(16, 28), (42, 46)],
            ) {
                return Err(Error::Verification);
            }
        } else if before.data() != after.data()
            || before.raw_record().local_record() != after.raw_record().local_record()
            || before.raw_record().compressed_data() != after.raw_record().compressed_data()
            || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(42, 46)],
            )
        {
            return Err(Error::Verification);
        }
    }
    for after in candidate_entries.iter() {
        if source_entries
            .iter()
            .all(|before| before.name() != after.name())
            && !changed(after.name())
        {
            return Err(Error::Verification);
        }
    }

    for component in source.state.source.components().iter() {
        let name = component.name();
        let Some(after_component) = candidate.state.source.components().get(name) else {
            if changed(name) {
                continue;
            }
            return Err(Error::Verification);
        };
        if !changed(name) {
            if component.archive().objects.len() != after_component.archive().objects.len()
                || component
                    .archive()
                    .objects
                    .iter()
                    .zip(&after_component.archive().objects)
                    .any(|(before, after)| !before.same_content_ignoring_offsets(after))
            {
                return Err(Error::Verification);
            }
            continue;
        }
        let allowed = changed_objects
            .iter()
            .find(|(item, _)| item.as_ref() == name)
            .map_or(&[][..], |(_, objects)| objects.as_slice());
        for before in &component.archive().objects {
            let identifier = before.archive_info.identifier.ok_or(Error::Verification)?;
            let Some(after) = after_component.archive().object(identifier) else {
                if !allowed.contains(&identifier) {
                    return Err(Error::Verification);
                }
                continue;
            };
            budget
                .work(object_work(before)?.saturating_add(object_work(after)?))
                .map_err(map_arrangement_error)?;
            if !allowed.contains(&identifier) && !before.same_content_ignoring_offsets(after) {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn entry_work(entry: &Entry) -> Result<usize, Error> {
    entry
        .name()
        .len()
        .checked_add(entry.raw_name().len())
        .and_then(|amount| amount.checked_add(entry.data().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().local_record().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().central_directory_record().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().compressed_data().len()))
        .ok_or(Error::InvalidSource)
}

fn object_work(object: &ArchiveObject) -> Result<usize, Error> {
    let mut work = size_of::<ArchiveObject>()
        .checked_add(object.messages.len())
        .and_then(|amount| amount.checked_add(object.archive_info.message_infos.len()))
        .ok_or(Error::InvalidSource)?;
    for message in &object.messages {
        work = work
            .checked_add(size_of::<RawMessage>())
            .and_then(|amount| amount.checked_add(message.data.len()))
            .ok_or(Error::InvalidSource)?;
    }
    Ok(work)
}

fn zip_record_equal_outside(before: &[u8], after: &[u8], excluded: &[(usize, usize)]) -> bool {
    if before.len() != after.len() {
        return false;
    }
    let mut cursor = 0;
    for &(start, end) in excluded {
        if start < cursor
            || end < start
            || end > before.len()
            || before[cursor..start] != after[cursor..start]
        {
            return false;
        }
        cursor = end;
    }
    before[cursor..] == after[cursor..]
}

fn zip_local_header_preserved(before: &[u8], after: &[u8]) -> bool {
    let Some(before_header) = zip_local_header_len(before) else {
        return false;
    };
    let Some(after_header) = zip_local_header_len(after) else {
        return false;
    };
    before_header == after_header
        && zip_record_equal_outside(
            &before[..before_header],
            &after[..after_header],
            &[(14, 26)],
        )
}

fn zip_local_header_len(record: &[u8]) -> Option<usize> {
    if record.len() < 30 || record.get(..4) != Some(b"PK\x03\x04") {
        return None;
    }
    let name = usize::from(u16::from_le_bytes([record[26], record[27]]));
    let extra = usize::from(u16::from_le_bytes([record[28], record[29]]));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

/// A staged immutable body-chart series-symbol edit.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    before: Vec<Option<Symbol>>,
    budget: ArrangementBudget,
    change: Option<Property>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.target.chart.position)
            .field("series", &self.target.series)
            .field("staged", &self.change.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage one semantic symbol vector.
    pub fn set(mut self, property: Property) -> Result<Self, Error> {
        if property.len() != self.target.series {
            return Err(Refusal::SeriesCount {
                expected: self.target.series,
                given: property.len(),
            }
            .into());
        }
        self.change = Some(property);
        Ok(self)
    }

    /// Publish the staged exact-source transaction and reopen its candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let change = self
            .change
            .clone()
            .ok_or_else(|| Error::from(Refusal::NoStagedOperation))?;
        commit_edit(self, change)
    }
}

/// Exact-source reversible body-chart series-symbol patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    position: Position,
    target: Target,
    before: Vec<Option<Symbol>>,
    after: Vec<Option<Symbol>>,
    changed_components: Arc<[Arc<str>]>,
    changed_objects: Vec<(Arc<str>, Vec<u64>)>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .field("series", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether this patch is an exact byte and semantic no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop() && self.before == self.after
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            position: self.position,
            target: self.target.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
            changed_components: self.changed_components.clone(),
            changed_objects: self.changed_objects.clone(),
        }
    }
}

/// Compact publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
}

impl Diagnostics {
    /// Whether package bytes or semantic values changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten existing components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
}

/// Fully reopened immutable result of one body-chart series-symbol edit.
#[must_use = "a series-symbol commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

fn unchanged(source: &Package, target: Target, before: Vec<Option<Symbol>>) -> Commit {
    let bytes = source.state.source.shared_source();
    let owner = SharedBytes::from_shared_slice(Arc::clone(&bytes));
    Commit {
        package: source.clone(),
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
            position: target.chart.position,
            target,
            before: before.clone(),
            after: before,
            changed_components: Arc::from([]),
            changed_objects: Vec::new(),
        },
        diagnostics: Diagnostics {
            changed: false,
            touched_components: 0,
        },
    }
}

fn commit_edit(edit: Edit<'_>, change: Property) -> Result<Commit, Error> {
    let source = edit.source;
    let mut budget = edit.budget;
    let before = edit.before;
    let Property::Symbols(values) = &change;
    if values.len() != edit.target.series {
        return Err(Refusal::SeriesCount {
            expected: edit.target.series,
            given: values.len(),
        }
        .into());
    }
    if before == *values {
        return Ok(unchanged(source, edit.target, before));
    }
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }

    admit_workspace(source, &mut budget)?;
    preflight_series_engine(source, edit.target.series, &budget)?;
    let components = PackageComponents(source);
    let limits = archive_limits(source)?;
    let mut workspace = Workspace::new(&components, limits).map_err(map_graph_error)?;
    let changed = series_engine::set_symbols(
        &mut workspace,
        edit.target.chart.drawable_identifier.get(),
        edit.target.family,
        values,
    )
    .map_err(map_graph_error)?;
    charge_series_engine(source, edit.target.series, &mut budget)?;
    if !changed {
        return Ok(unchanged(source, edit.target, before));
    }
    let changes = workspace.finish();
    if changes.is_empty() {
        return Ok(unchanged(source, edit.target, before));
    }

    let edit_count = changes.edits.len();
    let insertion_count = changes.insertions.len();
    let mut edits = changes.edits;
    let mut insertion_data = Vec::with_capacity(insertion_count);
    for insertion in changes.insertions {
        let data = insertion
            .archive
            .to_bytes(archive_limits(source)?)
            .map_err(map_core_error)?;
        insertion_data.push((insertion.name, data));
    }
    let deletions = changes.deletions;
    let projected = preflight_publication(source, &edits, &insertion_data, &budget)?;
    let names = edits
        .iter()
        .map(|(name, _)| Arc::<str>::from(name.as_str()))
        .chain(
            insertion_data
                .iter()
                .map(|(name, _)| Arc::<str>::from(name.as_str())),
        )
        .chain(deletions.iter().map(|name| Arc::<str>::from(name.as_str())))
        .collect::<Vec<_>>();
    let archives = std::mem::take(&mut edits)
        .into_iter()
        .map(|(name, archive)| (Arc::<str>::from(name), archive))
        .collect::<Vec<_>>();
    let insertion_refs = insertion_data
        .iter()
        .map(|(name, data)| (name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    let deletion_refs = deletions.iter().map(String::as_str).collect::<Vec<_>>();
    let candidate =
        package_graph::publish_changes(source, archives, &insertion_refs, &[], &deletion_refs)
            .map_err(map_package_graph_error)?;
    charge_publication(
        &candidate,
        edit_count,
        insertion_count,
        projected,
        &mut budget,
    )?;
    let mut candidate_budget = budget;
    let candidate_target = target_with_metadata(
        &candidate,
        BodyChartSelector::position(edit.target.chart.position),
        &mut candidate_budget,
    )?;
    if candidate_target.kind != edit.target.kind || candidate_target.series != edit.target.series {
        return Err(Error::Verification);
    }
    let after = read_symbols(&candidate, &candidate_target, &mut candidate_budget)?;
    if after != *values {
        return Err(Error::Verification);
    }
    let changed_objects = changed_object_ids(source, &candidate, &names)?;
    let target_bytes = Arc::clone(&candidate.state.source.shared_source());
    verify_locality(
        source,
        &candidate,
        &names,
        &changed_objects,
        &mut candidate_budget,
    )?;
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(
                SharedBytes::from_shared_slice(source.state.source.shared_source()),
                SharedBytes::from_shared_slice(target_bytes),
            ),
            position: edit.target.chart.position,
            target: edit.target,
            before,
            after,
            changed_components: names.into(),
            changed_objects,
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_components: edit_count,
        },
    })
}

fn apply_patch(package: &Package, patch: &Patch) -> Result<Commit, Error> {
    let mut budget = ArrangementBudget::new(package).map_err(map_arrangement_error)?;
    admit_workspace(package, &mut budget)?;
    let source = package.state.source.shared_source();
    let source_owner = SharedBytes::from_shared_slice(Arc::clone(&source));
    budget
        .work(
            source
                .len()
                .checked_add(patch.artifacts.source_owner().as_ref().len())
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_arrangement_error)?;
    if !patch.artifacts.authorizes_owner(&source_owner) {
        return Err(Error::PatchConflict);
    }
    let current = target_with_metadata(
        package,
        BodyChartSelector::position(patch.position),
        &mut budget,
    )?;
    let before = read_symbols(package, &current, &mut budget)?;
    if !current.chart.same_identity(&patch.target.chart)
        || current.kind != patch.target.kind
        || current.series != patch.target.series
        || before != patch.before
    {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit {
            package: package.clone(),
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: false,
                touched_components: 0,
            },
        });
    }
    if !package.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let target_owner = patch.artifacts.target_owner();
    budget
        .preflight_output(target_owner.as_ref().len())
        .map_err(map_arrangement_error)?;
    budget
        .preflight_scratch(target_owner.as_ref().len())
        .map_err(map_arrangement_error)?;
    budget
        .candidate_reopen(target_owner.as_ref().len())
        .map_err(map_arrangement_error)?;
    let candidate_bytes = Arc::<[u8]>::from(target_owner.as_ref());
    let candidate =
        package_graph::reopen(package, candidate_bytes).map_err(map_package_graph_error)?;
    budget
        .output(target_owner.as_ref().len())
        .map_err(map_arrangement_error)?;
    budget
        .scratch(target_owner.as_ref().len())
        .map_err(map_arrangement_error)?;
    let verified = target_with_metadata(
        &candidate,
        BodyChartSelector::position(patch.position),
        &mut budget,
    )?;
    let after = read_symbols(&candidate, &verified, &mut budget)?;
    if !verified.chart.same_identity(&patch.target.chart)
        || verified.kind != patch.target.kind
        || verified.series != patch.target.series
        || after != patch.after
    {
        return Err(Error::Verification);
    }
    verify_locality(
        package,
        &candidate,
        &patch.changed_components,
        &patch.changed_objects,
        &mut budget,
    )?;
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics {
            changed: true,
            touched_components: patch.changed_components.len(),
        },
    })
}

impl Package {
    /// Read the selected body chart's per-series symbols.
    pub fn body_chart_series(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Property, Error> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = target_with_metadata(self, selector.into(), &mut budget)?;
        Ok(Property::Symbols(read_symbols(self, &target, &mut budget)?))
    }

    /// Begin a selector-first per-series symbol transaction.
    pub fn edit_body_chart_series(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = target_with_metadata(self, selector.into(), &mut budget)?;
        let before = read_symbols(self, &target, &mut budget)?;
        Ok(Edit {
            source: self,
            target,
            before,
            budget,
            change: None,
        })
    }

    /// Apply a patch produced from this exact source package.
    pub fn apply_body_chart_series(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
