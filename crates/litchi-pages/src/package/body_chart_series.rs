//! Selector-first Pages body-chart series-symbol transactions.
//!
//! This adapter owns only Pages selection, bounded admission and exact-source
//! publication.  The series override graph (including symbol visibility,
//! shape and point-size inheritance) remains in `litchi_iwa_style`.

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{
    Entry, EntryEdit, EntryInsertion, OwnedExactArtifacts, SharedBytes,
};
use litchi_iwa_archive::unstable::{
    CandidateIngressCharge, CandidateIngressError, CandidateIngressPlan, CandidateIngressVisitor,
};
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::series::{Symbol, SymbolFamily};
use litchi_iwa_common::wire::{WireFieldView, WireView};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions};
use litchi_iwa_protos::chart_style_codec;
use litchi_iwa_style::graph::{self as style_graph, Components, Workspace};
use litchi_iwa_style::series as series_engine;
use thiserror::Error as ThisError;

use super::Package;
use super::body_chart_arrangement::{self as arrangement, ArrangementBudget, ChartTarget};
use super::body_drawable_graph as package_graph;
use crate::selector::BodyChartSelector;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const SERIES_STYLE_MESSAGE_TYPE: u32 = 5_028;
const SERIES_NON_STYLE_MESSAGE_TYPE: u32 = 5_029;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const OBJECT_CONTAINER_MESSAGE_TYPE: u32 = 11_008;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const STYLE_EXTENSION_FIELD: u32 = 10_000;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
struct MutationPlan {
    components: Vec<Arc<str>>,
    objects: Vec<(Arc<str>, Vec<u64>)>,
    allowed_message_types: Vec<u32>,
    allowed_fields: Vec<u32>,
}

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
    /// Native metadata items such as registry entries.
    PayloadItems,
    /// Native object identifiers allocated or released.
    PayloadObjects,
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
    Codec(chart_style_codec::LimitKind),
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
            Self::PayloadItems => formatter.write_str("payload items"),
            Self::PayloadObjects => formatter.write_str("payload objects"),
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
    /// A finite finalization arithmetic operation overflowed.
    #[error("Pages body-chart series-symbol finalization overflow: {reason}")]
    Overflow { reason: &'static str },
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
        arrangement::Error::InvalidSource => Error::InvalidSource,
        arrangement::Error::Verification => Error::Verification,
        arrangement::Error::PatchConflict => Error::PatchConflict,
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
            chart_style_codec::Error::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: LimitKind::Codec(kind),
                observed: observed as u64,
                maximum: maximum as u64,
            },
            chart_style_codec::Error::Allocation { bytes } => Error::Allocation { amount: bytes },
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

fn map_symbol_error(error: series_engine::SymbolError) -> Error {
    match error {
        series_engine::SymbolError::Graph(error) => map_graph_error(error),
        series_engine::SymbolError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                series_engine::LimitKind::InputBytes => LimitKind::InputBytes,
                series_engine::LimitKind::Fields => LimitKind::WireFields,
                series_engine::LimitKind::OutputBytes => LimitKind::OutputBytes,
                series_engine::LimitKind::Nesting => LimitKind::WireNesting,
                series_engine::LimitKind::Work => LimitKind::WireWork,
                series_engine::LimitKind::Allocations => LimitKind::Allocations,
                series_engine::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                series_engine::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                series_engine::LimitKind::References => LimitKind::PayloadReferences,
                series_engine::LimitKind::RegistryEntries => LimitKind::Registry(
                    litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
                        observed,
                        maximum,
                    },
                ),
                series_engine::LimitKind::Identifiers => LimitKind::PayloadReferences,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        series_engine::SymbolError::Allocation { amount } => Error::Allocation { amount },
        series_engine::SymbolError::Invalid(_) => Error::InvalidSource,
        series_engine::SymbolError::IdentifierOverflow => Error::InvalidSource,
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

fn positive_limit(value: usize, kind: LimitKind) -> Result<usize, Error> {
    if value == 0 {
        return Err(Error::Limit {
            kind,
            observed: 1,
            maximum: 0,
        });
    }
    Ok(value)
}

fn engine_limits(
    package: &Package,
    budget: &ArrangementBudget,
) -> Result<series_engine::Limits, Error> {
    let physical = package.state.source.limits();
    let package_bytes = usize::try_from(physical.max_input_bytes())
        .map_err(|_| Error::InvalidSource)?
        .max(1);
    let total_bytes = usize::try_from(physical.max_total_bytes())
        .map_err(|_| Error::InvalidSource)?
        .max(1);
    let wire = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    let codec_max = chart_style_codec::Limits::DEFAULT.max_bytes();
    let input = budget
        .residual_input_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes)
        .min(litchi_iwa_common::WireLimits::MAX_INPUT_BYTES)
        .min(codec_max);
    let input = positive_limit(input, LimitKind::InputBytes)?;
    let output = budget
        .residual_output_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes)
        .min(codec_max);
    let output = positive_limit(output, LimitKind::OutputBytes)?;
    let fields = wire
        .max_fields()
        .min(chart_style_codec::Limits::DEFAULT.max_fields())
        .min(litchi_iwa_common::WireLimits::MAX_FIELDS);
    let fields = positive_limit(fields, LimitKind::WireFields)?;
    let work = budget
        .residual_work_bytes()
        .map_err(map_arrangement_error)?
        .min(total_bytes.saturating_mul(32).max(1))
        .min(litchi_iwa_common::WireLimits::MAX_REWRITE_WORK);
    let work = positive_limit(work, LimitKind::WireWork)?;
    let allocations = budget
        .residual_allocations()
        .map_err(map_arrangement_error)?;
    let allocations = positive_limit(allocations, LimitKind::Allocations)?;
    let retained = budget
        .residual_retained_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes);
    let retained = positive_limit(retained, LimitKind::RetainedBytes)?;
    let scratch = budget
        .residual_scratch_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes);
    let scratch = positive_limit(scratch, LimitKind::ScratchBytes)?;
    let references = budget
        .residual_references()
        .map_err(map_arrangement_error)?;
    let references = positive_limit(references, LimitKind::PayloadReferences)?;
    let registry_entries = budget
        .residual_registry_entries()
        .map_err(map_arrangement_error)?;
    let registry_entries = positive_limit(registry_entries, LimitKind::PayloadItems)?;
    let identifiers = budget
        .residual_identifiers()
        .map_err(map_arrangement_error)?;
    let identifiers = positive_limit(identifiers, LimitKind::PayloadObjects)?;
    Ok(series_engine::Limits::new(
        input,
        fields,
        output,
        u32::try_from(wire.max_nesting()).map_err(|_| Error::InvalidSource)?,
        work,
        allocations,
        retained,
        scratch,
        references,
        registry_entries,
        identifiers,
    ))
}

fn charge_engine_report(
    budget: &mut ArrangementBudget,
    report: &series_engine::Report,
) -> Result<(), Error> {
    let max_nesting = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?
        .max_nesting();
    budget
        .input(report.input_bytes())
        .and_then(|_| budget.output(report.output_bytes()))
        .and_then(|_| budget.fields(report.fields()))
        .and_then(|_| budget.work(report.work_bytes()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .and_then(|_| budget.references(report.references()))
        .and_then(|_| budget.registry_entries(report.registry_entries()))
        .and_then(|_| budget.identifiers(report.identifiers()))
        .map_err(map_arrangement_error)?;
    if report.max_depth() as usize > max_nesting {
        return Err(Error::Limit {
            kind: LimitKind::WireNesting,
            observed: u64::from(report.max_depth()),
            maximum: max_nesting as u64,
        });
    }
    Ok(())
}

fn map_finish_error(error: style_graph::FinishError) -> Error {
    match error {
        style_graph::FinishError::Graph(error) => map_graph_error(error),
        style_graph::FinishError::Overflow(reason) => Error::Overflow { reason },
        style_graph::FinishError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                style_graph::FinishLimitKind::Work => LimitKind::WireWork,
                style_graph::FinishLimitKind::Allocations => LimitKind::Allocations,
                style_graph::FinishLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                style_graph::FinishLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        style_graph::FinishError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn preflight_finish(
    budget: &ArrangementBudget,
    requirements: style_graph::FinishRequirements,
) -> Result<(), Error> {
    budget
        .preflight_work(requirements.work())
        .and_then(|_| budget.preflight_allocations(requirements.allocations()))
        .and_then(|_| budget.preflight_retained(requirements.retained_bytes()))
        .and_then(|_| budget.preflight_scratch(requirements.scratch_bytes()))
        .map_err(map_arrangement_error)
}

fn finish_limits(budget: &ArrangementBudget) -> Result<style_graph::FinishLimits, Error> {
    Ok(style_graph::FinishLimits::new(
        budget
            .residual_work_bytes()
            .map_err(map_arrangement_error)?,
        budget
            .residual_allocations()
            .map_err(map_arrangement_error)?,
        budget
            .residual_retained_bytes()
            .map_err(map_arrangement_error)?,
        budget
            .residual_scratch_bytes()
            .map_err(map_arrangement_error)?,
    ))
}

fn charge_finish_report(
    budget: &mut ArrangementBudget,
    report: style_graph::FinishReport,
) -> Result<(), Error> {
    budget
        .work(report.work())
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .map_err(map_arrangement_error)
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
    if !budget.workspace_source_admitted {
        budget.input(source_bytes).map_err(map_arrangement_error)?;
        budget
            .retained(source_bytes)
            .map_err(map_arrangement_error)?;
        budget.workspace_source_admitted = true;
    }
    Ok(())
}

fn map_construction_error(error: style_graph::ConstructionError) -> Error {
    match error {
        style_graph::ConstructionError::Graph(error) => map_graph_error(error),
        style_graph::ConstructionError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                style_graph::ConstructionLimitKind::Work => LimitKind::WireWork,
                style_graph::ConstructionLimitKind::Allocations => LimitKind::Allocations,
                style_graph::ConstructionLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                style_graph::ConstructionLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        style_graph::ConstructionError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn construction_limits(
    budget: &ArrangementBudget,
) -> Result<style_graph::ConstructionLimits, Error> {
    Ok(style_graph::ConstructionLimits::new(
        budget
            .residual_work_bytes()
            .map_err(map_arrangement_error)?,
        budget
            .residual_allocations()
            .map_err(map_arrangement_error)?,
        budget
            .residual_retained_bytes()
            .map_err(map_arrangement_error)?,
        budget
            .residual_scratch_bytes()
            .map_err(map_arrangement_error)?,
    ))
}

fn charge_workspace_construction(
    budget: &mut ArrangementBudget,
    requirements: style_graph::ConstructionRequirements,
) -> Result<(), Error> {
    budget
        .preflight_work(requirements.source_work())
        .and_then(|_| budget.preflight_allocations(requirements.index_allocations()))
        .and_then(|_| budget.preflight_retained(requirements.index_retained_bytes()))
        .and_then(|_| budget.preflight_scratch(requirements.index_scratch_bytes()))
        .map_err(map_arrangement_error)?;
    budget
        .work(requirements.source_work())
        .and_then(|_| budget.allocations(requirements.index_allocations()))
        .and_then(|_| budget.retained(requirements.index_retained_bytes()))
        .and_then(|_| budget.scratch(requirements.index_scratch_bytes()))
        .map_err(map_arrangement_error)
}

fn charge_workspace_builder(
    budget: &mut ArrangementBudget,
    requirements: style_graph::ConstructionRequirements,
) -> Result<(), Error> {
    budget
        .preflight_work(requirements.builder_work())
        .map_err(map_arrangement_error)?;
    budget
        .work(requirements.builder_work())
        .map_err(map_arrangement_error)
}

fn metadata_options(budget: &ArrangementBudget, source: &[u8]) -> Result<DecodeOptions, Error> {
    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    let input = positive_limit(
        source.len().min(limits.max_input_bytes()),
        LimitKind::InputBytes,
    )?;
    let fields = positive_limit(limits.max_fields(), LimitKind::WireFields)?;
    let work = positive_limit(limits.max_rewrite_work(), LimitKind::WireWork)?;
    let labels = positive_limit(limits.max_fields(), LimitKind::PayloadItems)?;
    let text = positive_limit(limits.max_input_bytes(), LimitKind::PayloadBytes)?;
    Ok(DecodeOptions::new(
        input,
        fields,
        work,
        u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?,
        labels.min(1_000_000),
        text.min(64 * 1024 * 1024),
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
    let limits = archive_limits(package)?;
    let census_limits = construction_limits(budget)?;
    let construction = Workspace::construction_requirements(&components, census_limits)
        .map_err(map_construction_error)?;
    charge_workspace_construction(budget, construction)?;
    let builder_limits = construction_limits(budget)?;
    let workspace = Workspace::new_bounded(&components, limits, construction, builder_limits)
        .map_err(map_construction_error)?;
    charge_workspace_builder(budget, construction)?;
    let engine_limits = engine_limits(package, budget)?;
    let (values, report) = series_engine::symbols(
        &workspace,
        target.chart.drawable_identifier.get(),
        target.family,
        target.series,
        engine_limits,
    )
    .map_err(map_symbol_error)?;
    charge_engine_report(budget, &report)?;
    Ok(values)
}

fn archive_rewrite_bound(package: &Package, archive: &Archive) -> Result<usize, Error> {
    let encoded = archive
        .encoded_len(archive_limits(package)?)
        .map_err(map_core_error)?;
    let snappy = package
        .state
        .source
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    SnappyStream::maximum_compressed_len(encoded, snappy).map_err(map_core_error)
}

fn preflight_publication(
    source: &Package,
    edits: &[(String, Archive)],
    insertions: &[style_graph::Insertion],
    deletions: &[String],
    budget: &mut ArrangementBudget,
) -> Result<usize, Error> {
    let mut projected = source.state.source.shared_source().len();
    let mut transient = 0usize;
    let mut archive_work = 0usize;
    let archive_count = edits
        .len()
        .checked_add(insertions.len())
        .ok_or(Error::InvalidSource)?;
    for (name, archive) in edits {
        let encoded = archive
            .encoded_len(archive_limits(source)?)
            .map_err(map_core_error)?;
        let bound = archive_rewrite_bound(source, archive)?;
        transient = transient
            .checked_add(encoded)
            .and_then(|value| value.checked_add(bound))
            .ok_or(Error::InvalidSource)?;
        archive_work = archive_work
            .checked_add(name.len())
            .and_then(|value| value.checked_add(encoded))
            .and_then(|value| value.checked_add(bound))
            .ok_or(Error::InvalidSource)?;
        projected = projected
            .checked_add(name.len())
            .and_then(|value| value.checked_add(bound))
            .ok_or(Error::InvalidSource)?;
    }
    for insertion in insertions {
        let encoded = insertion
            .archive
            .encoded_len(archive_limits(source)?)
            .map_err(map_core_error)?;
        let bound = archive_rewrite_bound(source, &insertion.archive)?;
        transient = transient
            .checked_add(encoded)
            .and_then(|value| value.checked_add(bound))
            .ok_or(Error::InvalidSource)?;
        archive_work = archive_work
            .checked_add(insertion.name.len())
            .and_then(|value| value.checked_add(encoded))
            .and_then(|value| value.checked_add(bound))
            .ok_or(Error::InvalidSource)?;
        projected = projected
            .checked_add(insertion.name.len())
            .and_then(|value| value.checked_add(bound))
            .ok_or(Error::InvalidSource)?;
    }
    let collection_units = archive_count
        .checked_mul(3)
        .and_then(|value| value.checked_add(source.state.source.package().len()))
        .and_then(|value| value.checked_add(deletions.len()))
        .ok_or(Error::InvalidSource)?;
    let collection_bytes = archive_count
        .checked_mul(size_of::<(String, Vec<u8>)>())
        .and_then(|value| {
            value.checked_add(
                source
                    .state
                    .source
                    .package()
                    .len()
                    .checked_mul(size_of::<&str>())?,
            )
        })
        .and_then(|value| value.checked_add(deletions.len().checked_mul(size_of::<&str>())?))
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_output(projected)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_work(
            archive_work
                .checked_add(collection_bytes)
                .and_then(|value| value.checked_add(projected))
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_arrangement_error)?;
    budget
        .preflight_allocations(collection_units)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(
            transient
                .checked_add(collection_bytes)
                .and_then(|value| value.checked_add(projected))
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_arrangement_error)?;
    budget
        .preflight_scratch(
            transient
                .checked_add(collection_bytes)
                .and_then(|value| value.checked_add(projected))
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_arrangement_error)?;
    budget
        .work(
            archive_work
                .checked_add(collection_bytes)
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_arrangement_error)?;
    let staging_bytes = transient
        .checked_add(collection_bytes)
        .ok_or(Error::InvalidSource)?;
    budget
        .allocations(collection_units)
        .map_err(map_arrangement_error)?;
    budget
        .retained(staging_bytes)
        .map_err(map_arrangement_error)?;
    budget
        .scratch(staging_bytes)
        .map_err(map_arrangement_error)?;
    Ok(projected)
}

struct IngressBudget<'a>(&'a mut ArrangementBudget);

impl CandidateIngressVisitor for IngressBudget<'_> {
    type Error = Error;

    fn admit(&mut self, charge: CandidateIngressCharge) -> Result<(), Self::Error> {
        self.0
            .retained(charge.retained_bytes())
            .and_then(|_| self.0.work(charge.traversal_bytes()))
            .and_then(|_| self.0.allocations(charge.allocations()))
            .map_err(map_arrangement_error)
    }
}

fn map_candidate_ingress_error(error: CandidateIngressError<Error>) -> Error {
    match error {
        CandidateIngressError::Physical(error) => map_archive_error(error),
        CandidateIngressError::Admission(error) => error,
        CandidateIngressError::CloneArithmetic { .. } => Error::InvalidSource,
        _ => Error::InvalidSource,
    }
}

fn charge_candidate_plan(
    plan: &CandidateIngressPlan<'_>,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let requirements = plan.requirements();
    budget
        .input(requirements.input_bytes())
        .and_then(|_| budget.work(requirements.traversal_bytes()))
        .and_then(|_| budget.allocations(requirements.allocations()))
        .and_then(|_| budget.retained(requirements.retained_bytes()))
        .map_err(map_arrangement_error)?;
    Ok(())
}

fn publish_changes(
    source: &Package,
    changes: style_graph::Changes,
    budget: &mut ArrangementBudget,
) -> Result<(Package, usize), Error> {
    let touched_components = changes
        .edits
        .len()
        .checked_add(changes.insertions.len())
        .and_then(|value| value.checked_add(changes.deletions.len()))
        .ok_or(Error::InvalidSource)?;
    preflight_publication(
        source,
        &changes.edits,
        &changes.insertions,
        &changes.deletions,
        budget,
    )?;
    let archive_limits = archive_limits(source)?;
    let snappy_limits = source
        .state
        .source
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let mut edited = Vec::new();
    edited
        .try_reserve_exact(changes.edits.len())
        .map_err(|_| Error::Allocation {
            amount: changes.edits.len(),
        })?;
    for (name, archive) in changes.edits {
        let encoded = archive.to_bytes(archive_limits).map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&encoded, snappy_limits).map_err(map_core_error)?;
        edited.push((name, compressed));
    }

    let catalog = &source.state.source;
    let mut inserted = Vec::new();
    inserted
        .try_reserve_exact(changes.insertions.len())
        .map_err(|_| Error::Allocation {
            amount: changes.insertions.len(),
        })?;
    for insertion in changes.insertions {
        let encoded = insertion
            .archive
            .to_bytes(archive_limits)
            .map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&encoded, snappy_limits).map_err(map_core_error)?;
        inserted.push((insertion.name, compressed, insertion.before));
    }

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(edited.len())
        .map_err(|_| Error::Allocation {
            amount: edited.len(),
        })?;
    edits.extend(
        edited
            .iter()
            .map(|(name, data)| EntryEdit::new(name.as_str(), data.as_slice())),
    );

    let mut insertions = Vec::new();
    insertions
        .try_reserve_exact(inserted.len())
        .map_err(|_| Error::Allocation {
            amount: inserted.len(),
        })?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(catalog.package().len())
        .map_err(|_| Error::Allocation {
            amount: catalog.package().len(),
        })?;
    names.extend(catalog.package().iter().map(|entry| entry.name()));
    for (name, data, before) in &inserted {
        let insertion = match before.as_deref() {
            Some(before) => {
                let position = names
                    .iter()
                    .position(|candidate| *candidate == before)
                    .ok_or(Error::InvalidSource)?;
                EntryInsertion::at(position, name.as_str(), data.as_slice())
            },
            None => EntryInsertion::new(name.as_str(), data.as_slice()),
        };
        insertions.push(insertion);
    }
    let mut deletions = Vec::new();
    deletions
        .try_reserve_exact(changes.deletions.len())
        .map_err(|_| Error::Allocation {
            amount: changes.deletions.len(),
        })?;
    deletions.extend(changes.deletions.iter().map(String::as_str));

    let mut ingress = IngressBudget(budget);
    let prepared = catalog
        .__prepare_reassembly_with_changes_admitted(
            &insertions,
            &edits,
            &deletions,
            catalog.limits(),
            &mut ingress,
        )
        .map_err(map_candidate_ingress_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .output(requirements.output_bytes())
        .and_then(|_| budget.work(requirements.output_bytes()))
        .and_then(|_| budget.allocations(requirements.allocations()))
        .and_then(|_| budget.retained(requirements.retained_bytes()))
        .and_then(|_| budget.scratch(requirements.scratch_bytes()))
        .map_err(map_arrangement_error)?;
    // The prepared execution report is the exact ZIP-size authority. The
    // earlier projection only admits the archive staging phase; it must not
    // reject a valid insertion after preparation has already been admitted.
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate_plan =
        litchi_iwa_archive::SourceCatalog::__candidate_ingress(&output, catalog.limits())
            .map_err(map_archive_error)?;
    charge_candidate_plan(&candidate_plan, budget)?;
    budget
        .candidate_reopen(output.len())
        .map_err(map_arrangement_error)?;
    let target = Arc::<[u8]>::from(output);
    let candidate = package_graph::reopen(source, target).map_err(map_package_graph_error)?;
    Ok((candidate, touched_components))
}

fn package_object(package: &Package, identifier: u64) -> Option<(&str, &ArchiveObject)> {
    package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| {
            component
                .archive()
                .object(identifier)
                .map(|object| (component.name(), object))
        })
}

fn structural_scope_object(package: &Package, identifier: u64) -> bool {
    let Some((component, object)) = package_object(package, identifier) else {
        return false;
    };
    match component {
        "Index/Document.iwa" if identifier == 1 => object
            .messages
            .iter()
            .any(|message| message.type_ == ROOT_MESSAGE_TYPE),
        "Index/Metadata.iwa" => object
            .messages
            .iter()
            .any(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE),
        _ => false,
    }
}

fn add_planned_object(
    objects: &mut Vec<(Arc<str>, Vec<u64>)>,
    component: &str,
    identifier: u64,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    if let Some((_, identifiers)) = objects
        .iter_mut()
        .find(|(name, _)| name.as_ref() == component)
    {
        if !identifiers.contains(&identifier) {
            budget
                .preflight_allocations(1)
                .and_then(|_| budget.preflight_retained(size_of::<u64>()))
                .and_then(|_| budget.preflight_work(1))
                .map_err(map_arrangement_error)?;
            budget
                .allocations(1)
                .and_then(|_| budget.retained(size_of::<u64>()))
                .and_then(|_| budget.work(1))
                .map_err(map_arrangement_error)?;
            identifiers
                .try_reserve_exact(1)
                .map_err(|_| Error::Allocation { amount: 1 })?;
            identifiers.push(identifier);
        }
    } else {
        let retained = component
            .len()
            .checked_add(size_of::<(Arc<str>, Vec<u64>)>())
            .and_then(|value| value.checked_add(size_of::<u64>()))
            .ok_or(Error::InvalidSource)?;
        budget
            .preflight_allocations(3)
            .and_then(|_| budget.preflight_retained(retained))
            .and_then(|_| budget.preflight_work(1))
            .map_err(map_arrangement_error)?;
        budget
            .allocations(3)
            .and_then(|_| budget.retained(retained))
            .and_then(|_| budget.work(1))
            .map_err(map_arrangement_error)?;
        objects
            .try_reserve_exact(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        let mut identifiers = Vec::new();
        identifiers
            .try_reserve_exact(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        identifiers.push(identifier);
        objects.push((Arc::from(component), identifiers));
    }
    Ok(())
}

fn mutation_plan(
    source: &Package,
    components: &[Arc<str>],
    changes: &style_graph::Changes,
    scope: &series_engine::MutationScope,
    budget: &mut ArrangementBudget,
) -> Result<MutationPlan, Error> {
    let member_allowed = |name: &str| {
        scope.allowed_members().iter().any(|member| member == name)
            || scope
                .planned_new_members()
                .iter()
                .any(|member| member == name)
            || scope
                .planned_removed_members()
                .iter()
                .any(|member| member == name)
    };
    let existing_allowed = |identifier: u64| {
        scope.allowed_objects().contains(&identifier)
            || scope.planned_removed_objects().contains(&identifier)
            || structural_scope_object(source, identifier)
    };
    let new_allowed = |identifier: u64| scope.planned_new_objects().contains(&identifier);
    let mut objects = Vec::new();
    let source_object_count = scope
        .allowed_objects()
        .len()
        .checked_add(scope.planned_removed_objects().len())
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_allocations(1)
        .and_then(|_| {
            budget.preflight_retained(source_object_count.saturating_mul(size_of::<u64>()))
        })
        .map_err(map_arrangement_error)?;
    budget
        .allocations(1)
        .and_then(|_| budget.retained(source_object_count.saturating_mul(size_of::<u64>())))
        .map_err(map_arrangement_error)?;
    let mut source_objects = Vec::new();
    source_objects
        .try_reserve_exact(source_object_count)
        .map_err(|_| Error::Allocation {
            amount: source_object_count,
        })?;
    source_objects.extend_from_slice(scope.allowed_objects());
    source_objects.extend_from_slice(scope.planned_removed_objects());
    for identifier in source_objects {
        let Some((component, _)) = package_object(source, identifier) else {
            return Err(Error::Verification);
        };
        add_planned_object(&mut objects, component, identifier, budget)?;
    }
    for (name, archive) in &changes.edits {
        if !member_allowed(name) {
            return Err(Error::Verification);
        }
        for object in &archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::Verification)?;
            if let Some((_, source_object)) = package_object(source, identifier)
                && same_object_metadata(source_object, object)
            {
                continue;
            }
            if package_object(source, identifier).is_some() {
                if !existing_allowed(identifier) {
                    return Err(Error::Verification);
                }
                if structural_scope_object(source, identifier) {
                    add_planned_object(&mut objects, name, identifier, budget)?;
                }
            } else if !new_allowed(identifier)
                || !object
                    .messages
                    .iter()
                    .all(|message| scope.allowed_message_types().contains(&message.type_))
            {
                return Err(Error::Verification);
            } else {
                add_planned_object(&mut objects, name, identifier, budget)?;
            }
        }
    }
    for insertion in &changes.insertions {
        if !member_allowed(&insertion.name) {
            return Err(Error::Verification);
        }
        for object in &insertion.archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::Verification)?;
            if !new_allowed(identifier)
                || !object
                    .messages
                    .iter()
                    .all(|message| scope.allowed_message_types().contains(&message.type_))
            {
                return Err(Error::Verification);
            }
            add_planned_object(&mut objects, &insertion.name, identifier, budget)?;
        }
    }
    for name in components.iter().map(Arc::as_ref) {
        if !member_allowed(name) {
            return Err(Error::Verification);
        }
    }
    let mut scoped_components = Vec::new();
    let component_names = scope
        .allowed_members()
        .iter()
        .chain(scope.planned_new_members())
        .chain(scope.planned_removed_members())
        .map(String::as_str)
        .chain(components.iter().map(Arc::as_ref));
    for name in component_names {
        if scoped_components
            .iter()
            .any(|item: &Arc<str>| item.as_ref() == name)
        {
            continue;
        }
        let retained = name.len().saturating_add(size_of::<Arc<str>>());
        budget
            .preflight_allocations(2)
            .and_then(|_| budget.preflight_retained(retained))
            .map_err(map_arrangement_error)?;
        budget
            .allocations(2)
            .and_then(|_| budget.retained(retained))
            .map_err(map_arrangement_error)?;
        scoped_components
            .try_reserve_exact(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        scoped_components.push(Arc::from(name));
    }
    budget
        .preflight_allocations(2)
        .and_then(|_| {
            budget.preflight_retained(
                scope
                    .allowed_message_types()
                    .len()
                    .saturating_mul(size_of::<u32>()),
            )
        })
        .and_then(|_| {
            budget.preflight_retained(
                scope
                    .allowed_fields()
                    .len()
                    .saturating_mul(size_of::<u32>()),
            )
        })
        .map_err(map_arrangement_error)?;
    budget.allocations(2).map_err(map_arrangement_error)?;
    let mut allowed_message_types = Vec::new();
    allowed_message_types
        .try_reserve_exact(scope.allowed_message_types().len())
        .map_err(|_| Error::Allocation {
            amount: scope.allowed_message_types().len(),
        })?;
    allowed_message_types.extend_from_slice(scope.allowed_message_types());
    let mut allowed_fields = Vec::new();
    allowed_fields
        .try_reserve_exact(scope.allowed_fields().len())
        .map_err(|_| Error::Allocation {
            amount: scope.allowed_fields().len(),
        })?;
    allowed_fields.extend_from_slice(scope.allowed_fields());
    Ok(MutationPlan {
        components: scoped_components,
        objects,
        allowed_message_types,
        allowed_fields,
    })
}

fn planned_contains(plan: &MutationPlan, component: &str, identifier: u64) -> bool {
    plan.objects
        .iter()
        .any(|(name, ids)| name.as_ref() == component && ids.contains(&identifier))
}

impl MutationPlan {
    fn allowed_object(&self, identifier: u64) -> bool {
        self.objects
            .iter()
            .any(|(_, identifiers)| identifiers.contains(&identifier))
    }
}

fn unique_wire_field<'a>(
    view: &'a WireView<'a>,
    number: u32,
) -> Result<Option<WireFieldView<'a>>, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.replace(field).is_some() {
            return Err(Error::Verification);
        }
    }
    Ok(found)
}

fn same_wire_fields_except(
    source: &WireView<'_>,
    candidate: &WireView<'_>,
    excluded: &[u32],
    allowed_fields: &[u32],
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let mut source_fields = source.fields().filter(|field| {
        !(excluded.contains(&field.number()) && allowed_fields.contains(&field.number()))
    });
    let mut candidate_fields = candidate.fields().filter(|field| {
        !(excluded.contains(&field.number()) && allowed_fields.contains(&field.number()))
    });
    loop {
        match (source_fields.next(), candidate_fields.next()) {
            (None, None) => return Ok(()),
            (Some(left), Some(right)) => {
                budget
                    .work(left.raw().len().saturating_add(right.raw().len()))
                    .map_err(map_arrangement_error)?;
                if left.raw() != right.raw() {
                    return Err(Error::Verification);
                }
            },
            _ => {
                return Err(Error::Verification);
            },
        }
    }
}

fn verify_message_unknowns(
    message_type: u32,
    source: &[u8],
    candidate: &[u8],
    family: SymbolFamily,
    plan: &MutationPlan,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let source_view = budget.parse(source, 1).map_err(map_arrangement_error)?;
    let candidate_view = budget.parse(candidate, 1).map_err(map_arrangement_error)?;
    let outer_excluded: &[u32] = match message_type {
        CHART_MESSAGE_TYPE | SERIES_STYLE_MESSAGE_TYPE | SERIES_NON_STYLE_MESSAGE_TYPE => {
            &[STYLE_EXTENSION_FIELD]
        },
        STYLESHEET_MESSAGE_TYPE => &[1],
        OBJECT_CONTAINER_MESSAGE_TYPE => &[2],
        PACKAGE_METADATA_MESSAGE_TYPE => &[1, 3, 8, 10, 12],
        ROOT_MESSAGE_TYPE => &[7, 9],
        _ => return Err(Error::Verification),
    };
    same_wire_fields_except(
        &source_view,
        &candidate_view,
        outer_excluded,
        &plan.allowed_fields,
        budget,
    )?;
    if message_type == PACKAGE_METADATA_MESSAGE_TYPE {
        let mut source_components = source_view.fields().filter(|field| field.number() == 3);
        let mut candidate_components = candidate_view.fields().filter(|field| field.number() == 3);
        loop {
            match (source_components.next(), candidate_components.next()) {
                (None, None) => break,
                (Some(source_component), Some(candidate_component)) => {
                    let source_view = budget
                        .parse(source_component.payload(), 2)
                        .map_err(map_arrangement_error)?;
                    let candidate_view = budget
                        .parse(candidate_component.payload(), 2)
                        .map_err(map_arrangement_error)?;
                    same_wire_fields_except(
                        &source_view,
                        &candidate_view,
                        &[6, 11],
                        &plan.allowed_fields,
                        budget,
                    )?;
                },
                _ => return Err(Error::Verification),
            }
        }
        return Ok(());
    }
    if !outer_excluded.contains(&STYLE_EXTENSION_FIELD) {
        return Ok(());
    }
    match (
        unique_wire_field(&source_view, STYLE_EXTENSION_FIELD)?,
        unique_wire_field(&candidate_view, STYLE_EXTENSION_FIELD)?,
    ) {
        (None, None) => Ok(()),
        (Some(source), Some(candidate)) => {
            let source_view = budget
                .parse(source.payload(), 2)
                .map_err(map_arrangement_error)?;
            let candidate_view = budget
                .parse(candidate.payload(), 2)
                .map_err(map_arrangement_error)?;
            let excluded: &[u32] = match message_type {
                CHART_MESSAGE_TYPE => &[19],
                SERIES_NON_STYLE_MESSAGE_TYPE => &[family.show_field(), family.type_field()],
                SERIES_STYLE_MESSAGE_TYPE => &[family.size_field()],
                _ => &[],
            };
            same_wire_fields_except(
                &source_view,
                &candidate_view,
                excluded,
                &plan.allowed_fields,
                budget,
            )
        },
        _ => Err(Error::Verification),
    }
}

fn verify_object_messages(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    family: SymbolFamily,
    plan: &MutationPlan,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    if source.messages.len() != candidate.messages.len() {
        return Err(Error::Verification);
    }
    for (before, after) in source.messages.iter().zip(&candidate.messages) {
        if before.type_ != after.type_ {
            return Err(Error::Verification);
        }
        if before.data == after.data {
            continue;
        }
        if !plan.allowed_message_types.contains(&before.type_) {
            return Err(Error::Verification);
        }
        verify_message_unknowns(
            before.type_,
            &before.data,
            &after.data,
            family,
            plan,
            budget,
        )?;
    }
    Ok(())
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    plan: &MutationPlan,
    family: SymbolFamily,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let changed = |name: &str| plan.components.iter().any(|item| item.as_ref() == name);
    let source_entries = source.state.source.package();
    let candidate_entries = candidate.state.source.package();
    for before in source_entries.iter() {
        let after = candidate_entries
            .iter()
            .find(|item| item.name() == before.name());
        let Some(after) = after else {
            if !changed(before.name()) {
                return Err(Error::Verification);
            }
            continue;
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
                    .any(|(before, after)| !same_object_metadata(before, after))
            {
                return Err(Error::Verification);
            }
            continue;
        }
        for before in &component.archive().objects {
            let identifier = before.archive_info.identifier.ok_or(Error::Verification)?;
            let Some(after) = after_component.archive().object(identifier) else {
                if !planned_contains(plan, name, identifier) {
                    return Err(Error::Verification);
                }
                continue;
            };
            budget
                .work(object_work(before)?.saturating_add(object_work(after)?))
                .map_err(map_arrangement_error)?;
            let planned = planned_contains(plan, name, identifier);
            let unchanged = same_object_metadata(before, after);
            if !planned && !unchanged {
                return Err(Error::Verification);
            }
            if planned && !same_archive_info_except_references(before, after, true, plan) {
                return Err(Error::Verification);
            }
            if planned && !before.same_source_headers(after) {
                return Err(Error::Verification);
            }
            if !unchanged {
                verify_object_messages(before, after, family, plan, budget)?;
            }
        }
        for after in &after_component.archive().objects {
            let identifier = after.archive_info.identifier.ok_or(Error::Verification)?;
            if component.archive().object(identifier).is_none()
                && !planned_contains(plan, name, identifier)
            {
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

fn same_object_metadata(before: &ArchiveObject, after: &ArchiveObject) -> bool {
    before.same_content_ignoring_offsets(after)
}

fn same_archive_info_except_references(
    before: &ArchiveObject,
    after: &ArchiveObject,
    allow_message_lengths: bool,
    plan: &MutationPlan,
) -> bool {
    if before.archive_info.identifier != after.archive_info.identifier
        || before.archive_info.should_merge != after.archive_info.should_merge
        || before.archive_info.message_infos.len() != after.archive_info.message_infos.len()
    {
        return false;
    }
    before
        .archive_info
        .message_infos
        .iter()
        .zip(&after.archive_info.message_infos)
        .all(|(left, right)| {
            left.type_ == right.type_
                && left.versions == right.versions
                && (allow_message_lengths || left.length == right.length)
                && left.base_message_index == right.base_message_index
                && left.diff_merge_version == right.diff_merge_version
                && left.diff_field_path == right.diff_field_path
                && left.fields_to_remove == right.fields_to_remove
                && left.diff_read_version == right.diff_read_version
                && reference_transition_allowed(
                    &left.object_references,
                    &right.object_references,
                    plan,
                )
                && left.data_references == right.data_references
                && left.field_infos.len() == right.field_infos.len()
                && left
                    .field_infos
                    .iter()
                    .zip(&right.field_infos)
                    .all(|(left, right)| {
                        left.path == right.path
                            && left.r#type == right.r#type
                            && left.unknown_field_rule == right.unknown_field_rule
                            && left.known_field_rule == right.known_field_rule
                            && left.known_field_version == right.known_field_version
                            && left.known_field_feature_identifier
                                == right.known_field_feature_identifier
                            && reference_transition_allowed(
                                &left.object_references,
                                &right.object_references,
                                plan,
                            )
                            && left.data_references == right.data_references
                    })
        })
}

fn reference_transition_allowed(before: &[u64], after: &[u64], plan: &MutationPlan) -> bool {
    let count =
        |values: &[u64], identifier| values.iter().filter(|value| **value == identifier).count();
    before.iter().chain(after).all(|identifier| {
        let left = count(before, *identifier);
        let right = count(after, *identifier);
        left == right || plan.allowed_object(*identifier)
    })
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
    plan: MutationPlan,
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
            plan: self.plan.clone(),
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
            plan: MutationPlan::default(),
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
    let components = PackageComponents(source);
    let limits = archive_limits(source)?;
    let census_limits = construction_limits(&budget)?;
    let construction = Workspace::construction_requirements(&components, census_limits)
        .map_err(map_construction_error)?;
    charge_workspace_construction(&mut budget, construction)?;
    let builder_limits = construction_limits(&budget)?;
    let mut workspace = Workspace::new_bounded(&components, limits, construction, builder_limits)
        .map_err(map_construction_error)?;
    charge_workspace_builder(&mut budget, construction)?;
    let engine_limits = engine_limits(source, &budget)?;
    let (changed, report) = series_engine::set_symbols(
        &mut workspace,
        edit.target.chart.drawable_identifier.get(),
        edit.target.family,
        values,
        engine_limits,
    )
    .map_err(map_symbol_error)?;
    let scope = report.scope().cloned().ok_or(Error::Verification)?;
    charge_engine_report(&mut budget, &report)?;
    if !changed {
        return Ok(unchanged(source, edit.target, before));
    }
    let finish_requirements = workspace.finish_requirements().map_err(map_finish_error)?;
    preflight_finish(&budget, finish_requirements)?;
    let finish_limits = finish_limits(&budget)?;
    let (changes, finish_report) = workspace
        .finish_bounded(finish_limits)
        .map_err(map_finish_error)?;
    charge_finish_report(&mut budget, finish_report)?;
    if changes.is_empty() {
        return Ok(unchanged(source, edit.target, before));
    }

    let name_count = changes
        .edits
        .len()
        .checked_add(changes.insertions.len())
        .and_then(|value| value.checked_add(changes.deletions.len()))
        .ok_or(Error::InvalidSource)?;
    let name_bytes = changes
        .edits
        .iter()
        .map(|(name, _)| name.len())
        .chain(
            changes
                .insertions
                .iter()
                .map(|insertion| insertion.name.len()),
        )
        .chain(changes.deletions.iter().map(String::len))
        .try_fold(0usize, |total, bytes| total.checked_add(bytes))
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_allocations(name_count)
        .and_then(|_| budget.preflight_retained(name_bytes))
        .and_then(|_| budget.preflight_work(name_bytes))
        .map_err(map_arrangement_error)?;
    budget
        .allocations(name_count)
        .and_then(|_| budget.retained(name_bytes))
        .and_then(|_| budget.work(name_bytes))
        .map_err(map_arrangement_error)?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(name_count)
        .map_err(|_| Error::Allocation { amount: name_count })?;
    names.extend(
        changes
            .edits
            .iter()
            .map(|(name, _)| Arc::<str>::from(name.as_str()))
            .chain(
                changes
                    .insertions
                    .iter()
                    .map(|insertion| Arc::<str>::from(insertion.name.as_str())),
            )
            .chain(
                changes
                    .deletions
                    .iter()
                    .map(|name| Arc::<str>::from(name.as_str())),
            ),
    );
    let plan = mutation_plan(source, &names, &changes, &scope, &mut budget)?;
    let (candidate, touched_components) = publish_changes(source, changes, &mut budget)?;
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
    let target_bytes = Arc::clone(&candidate.state.source.shared_source());
    verify_locality(
        source,
        &candidate,
        &plan,
        edit.target.family,
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
            plan,
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_components,
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
    let candidate_plan = litchi_iwa_archive::SourceCatalog::__candidate_ingress(
        target_owner.as_ref(),
        package.state.source.limits(),
    )
    .map_err(map_archive_error)?;
    charge_candidate_plan(&candidate_plan, &mut budget)?;
    budget
        .candidate_reopen(target_owner.as_ref().len())
        .map_err(map_arrangement_error)?;
    let candidate_bytes = Arc::<[u8]>::from(target_owner.as_ref());
    budget
        .output(target_owner.as_ref().len())
        .and_then(|_| budget.scratch(target_owner.as_ref().len()))
        .map_err(map_arrangement_error)?;
    let candidate =
        package_graph::reopen(package, candidate_bytes).map_err(map_package_graph_error)?;
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
        &patch.plan,
        patch.target.family,
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

#[cfg(test)]
mod tests {
    use super::*;

    const NATIVE: &[u8] =
        include_bytes!("../../../../test-data/iwork/pages/chart-data-native.pages");

    fn append_unknown_marker(data: &mut Vec<u8>, marker: u64) {
        litchi_iwa_common::wire::append_varint_field(data, 12_345, marker)
            .expect("bounded unknown marker");
    }

    fn empty_plan() -> MutationPlan {
        MutationPlan::default()
    }

    #[test]
    fn production_locality_rejects_unplanned_object_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let component = package
            .state
            .source
            .components()
            .iter()
            .find(|component| component.archive().objects.len() >= 2)
            .expect("fixture component with two objects");
        let component_name = Arc::<str>::from(component.name());
        let ids = component
            .archive()
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .take(2)
            .collect::<Vec<_>>();
        let allowed = ids[0];
        let unplanned = ids[1];
        let mut archive = component.archive().clone();
        let message = archive
            .object_mut(unplanned)
            .expect("unplanned object")
            .messages
            .first_mut()
            .expect("unplanned message");
        append_unknown_marker(&mut message.data, 0xfeed);
        let candidate =
            package_graph::publish_archives(&package, vec![(Arc::clone(&component_name), archive)])
                .expect("candidate publication");
        let plan = MutationPlan {
            components: vec![Arc::clone(&component_name)],
            objects: vec![(component_name, vec![allowed])],
            ..MutationPlan::default()
        };
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(
                &package,
                &candidate,
                &plan,
                SymbolFamily::Line2d,
                &mut budget,
            ),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_changed_chart_unknown_field() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let (component_name, chart_id) = package
            .state
            .source
            .components()
            .iter()
            .find_map(|component| {
                component
                    .archive()
                    .objects
                    .iter()
                    .find(|object| {
                        object
                            .messages
                            .iter()
                            .any(|message| message.type_ == CHART_MESSAGE_TYPE)
                    })
                    .and_then(|object| object.archive_info.identifier)
                    .map(|identifier| (Arc::<str>::from(component.name()), identifier))
            })
            .expect("fixture chart object");
        let mut archive = package
            .state
            .source
            .components()
            .get(component_name.as_ref())
            .expect("chart component")
            .archive()
            .clone();
        let message = archive
            .object_mut(chart_id)
            .expect("chart object")
            .messages
            .iter_mut()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .expect("chart message");
        append_unknown_marker(&mut message.data, 0xbeef);
        let candidate =
            package_graph::publish_archives(&package, vec![(Arc::clone(&component_name), archive)])
                .expect("candidate publication");
        let plan = MutationPlan {
            components: vec![Arc::clone(&component_name)],
            objects: vec![(component_name, vec![chart_id])],
            ..MutationPlan::default()
        };
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(
                &package,
                &candidate,
                &plan,
                SymbolFamily::Line2d,
                &mut budget,
            ),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_extra_zip_member() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let candidate = package_graph::publish_changes(
            &package,
            Vec::new(),
            &[("Data/unplanned-series.bin", b"unused".as_slice())],
            &[],
            &[],
        )
        .expect("extra-member candidate publication");
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(
                &package,
                &candidate,
                &empty_plan(),
                SymbolFamily::Line2d,
                &mut budget,
            ),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_zip_metadata_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let entry = package
            .state
            .source
            .package()
            .iter()
            .find(|entry| !entry.is_opaque())
            .expect("decoded package member");
        let central = entry.raw_record().central_directory_record();
        let mut bytes = package.state.source.shared_source().to_vec();
        let offset = bytes
            .windows(central.len())
            .position(|window| window == central)
            .expect("central record in source bytes");
        bytes[offset + 38] ^= 1;
        let candidate = Package::from_bytes(&bytes).expect("metadata-tampered candidate");
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(
                &package,
                &candidate,
                &empty_plan(),
                SymbolFamily::Line2d,
                &mut budget,
            ),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_archive_metadata_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let component = package
            .state
            .source
            .components()
            .iter()
            .find(|component| !component.archive().objects.is_empty())
            .expect("fixture component");
        let component_name = Arc::<str>::from(component.name());
        let object_id = component.archive().objects[0]
            .archive_info
            .identifier
            .expect("object identifier");
        let mut archive = component.archive().clone();
        let object = archive.object_mut(object_id).expect("metadata object");
        object.archive_info.should_merge = Some(!object.archive_info.should_merge.unwrap_or(false));
        let candidate =
            package_graph::publish_archives(&package, vec![(Arc::clone(&component_name), archive)])
                .expect("metadata candidate publication");
        let plan = MutationPlan {
            components: vec![Arc::clone(&component_name)],
            objects: vec![(component_name, vec![object_id])],
            ..MutationPlan::default()
        };
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(
                &package,
                &candidate,
                &plan,
                SymbolFamily::Line2d,
                &mut budget,
            ),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_unplanned_registry_object_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let component = package
            .state
            .source
            .components()
            .iter()
            .find(|component| component.name() == "Index/Metadata.iwa")
            .expect("package metadata component");
        let component_name = Arc::<str>::from(component.name());
        let object_id = component.archive().objects[0]
            .archive_info
            .identifier
            .expect("registry object identifier");
        let mut archive = component.archive().clone();
        let message = archive
            .object_mut(object_id)
            .expect("registry object")
            .messages
            .first_mut()
            .expect("registry message");
        append_unknown_marker(&mut message.data, 0xabad1dea);
        let candidate =
            package_graph::publish_archives(&package, vec![(Arc::clone(&component_name), archive)])
                .expect("registry candidate publication");
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(
                &package,
                &candidate,
                &empty_plan(),
                SymbolFamily::Line2d,
                &mut budget,
            ),
            Err(Error::Verification)
        ));
    }
}
