//! Selector-first Pages body-chart font and point-size transactions.
//!
//! Selection and publication stay in the Pages adapter.  Paragraph-style
//! variation planning, registration, repointing and pruning are delegated to
//! `litchi_iwa_style::chart_font`; no native object identity crosses this
//! module's public API.

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{Entry, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_style::chart_font::{self as engine, ChartFont};
use litchi_iwa_style::graph::{self as style_graph, Components, Workspace};
use litchi_iwa_text::character::TextPointSize;
use thiserror::Error as ThisError;

use super::Package;
use super::body_chart_arrangement::{self as arrangement, ArrangementBudget, ChartTarget};
use super::body_drawable_graph as package_graph;
use crate::selector::BodyChartSelector;

/// A finite resource charged by one Pages body-chart font operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package bytes inspected by rooted selection.
    InputBytes,
    /// Candidate package bytes emitted or reopened.
    OutputBytes,
    /// Native payload bytes inspected.
    PayloadBytes,
    /// Native references inspected.
    PayloadReferences,
    /// Wire fields visited.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Wire and graph work.
    WireWork,
    /// Semantic font-name bytes.
    TextBytes,
    /// Bounded temporary allocations.
    Allocations,
    /// Source and candidate bytes retained together.
    RetainedBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
    /// Physical archive resource.
    Archive(litchi_iwa_core::LimitKind),
    /// ZIP/package ingress resource.
    Package(litchi_iwa_archive::LimitKind),
    /// Shared wire resource.
    Wire(litchi_iwa_common::LimitKind),
    /// Generated chart-style codec resource.
    Codec(litchi_iwa_protos::chart_style_codec::LimitKind),
    /// Package-metadata registry rewrite resource.
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
            Self::TextBytes => formatter.write_str("text bytes"),
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

/// A well-formed request refused by this focused family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No chart matched the source-order selector.
    #[error("the Pages body has no chart at position {position:?}")]
    ChartNotFound { position: Position },
    /// The source does not retain an exact native artifact.
    #[error("this Pages source does not support exact body-chart font edits")]
    UnsupportedSource,
    /// The requested operation is for the other property family.
    #[error("the staged operation does not address this chart-font property")]
    WrongProperty,
    /// No operation was staged.
    #[error("no chart-font change is staged")]
    NoStagedOperation,
    /// The graph produced a publication shape this focused family does not own.
    #[error("the chart-font graph produced unsupported component changes")]
    UnsupportedPublication,
    /// The package-wide object identifier space is exhausted.
    #[error("the Pages chart-font object identifier space is exhausted")]
    IdentifierOverflow,
}

impl Refusal {
    /// The shared refusal category for litchi-ops.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::ChartNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedSource | Self::WrongProperty | Self::NoStagedOperation => {
                RefusalKind::UnsupportedOperation
            },
            Self::UnsupportedPublication | Self::IdentifierOverflow => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

/// Failure from a Pages body-chart font read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted chart graph or selected style graph is malformed.
    #[error("the selected Pages body-chart font source is invalid")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Pages body-chart font {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} Pages body-chart font units")]
    Allocation { amount: usize },
    /// Candidate readback did not reproduce the requested value.
    #[error("the edited Pages body-chart font failed semantic verification")]
    Verification,
    /// The patch was produced from another exact source artifact.
    #[error("the Pages body-chart font patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(value: Refusal) -> Self {
        Self::Refused(value)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Font,
    Size,
}

#[derive(Debug, Clone, PartialEq)]
enum Value {
    Font(ChartFont),
    Size(TextPointSize),
}

#[derive(Debug, Clone)]
enum Change {
    Set(Value),
    Reset,
}

/// The package's parsed components as the shared graph engine sees them.
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

fn map_arrangement_error(error: arrangement::Error) -> Error {
    match error {
        arrangement::Error::Refused(
            arrangement::Refusal::ChartNotFound { position }
            | arrangement::Refusal::ChartPositionNotFound { position },
        ) => Refusal::ChartNotFound { position }.into(),
        arrangement::Error::Refused(arrangement::Refusal::UnsupportedSource) => {
            Refusal::UnsupportedSource.into()
        },
        arrangement::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                arrangement::LimitKind::InputBytes => LimitKind::InputBytes,
                arrangement::LimitKind::OutputBytes => LimitKind::OutputBytes,
                arrangement::LimitKind::Entries => LimitKind::PayloadReferences,
                arrangement::LimitKind::EntryBytes
                | arrangement::LimitKind::TotalEntryBytes
                | arrangement::LimitKind::PackageBytes => LimitKind::InputBytes,
                arrangement::LimitKind::PayloadBytes
                | arrangement::LimitKind::TotalPayloadBytes
                | arrangement::LimitKind::PayloadObjects
                | arrangement::LimitKind::PayloadMessages
                | arrangement::LimitKind::PayloadItems => LimitKind::PayloadBytes,
                arrangement::LimitKind::PayloadReferences => LimitKind::PayloadReferences,
                arrangement::LimitKind::WireBytes | arrangement::LimitKind::WireOutputBytes => {
                    LimitKind::PayloadBytes
                },
                arrangement::LimitKind::WireFields => LimitKind::WireFields,
                arrangement::LimitKind::WireNesting => LimitKind::WireNesting,
                arrangement::LimitKind::WireWork => LimitKind::WireWork,
                arrangement::LimitKind::WireAllocations => LimitKind::Allocations,
                arrangement::LimitKind::WireRetainedBytes => LimitKind::RetainedBytes,
                arrangement::LimitKind::WireScratchBytes => LimitKind::ScratchBytes,
            },
            observed,
            maximum,
        },
        arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        arrangement::Error::Verification => Error::Verification,
        arrangement::Error::PatchConflict => Error::PatchConflict,
        arrangement::Error::InvalidSource => Error::InvalidSource,
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

fn map_engine_error(error: engine::Error) -> Error {
    match error {
        engine::Error::InvalidSource(_) => Error::InvalidSource,
        engine::Error::Limit {
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
                engine::LimitKind::TextBytes => LimitKind::TextBytes,
                engine::LimitKind::Allocations => LimitKind::Allocations,
                engine::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                engine::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                engine::LimitKind::Archive(kind) => LimitKind::Archive(kind),
                engine::LimitKind::Wire(kind) => LimitKind::Wire(kind),
                engine::LimitKind::Codec(kind) => LimitKind::Codec(kind),
                _ => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
        engine::Error::Allocation { amount } => Error::Allocation { amount },
        engine::Error::Refused(_) => Refusal::UnsupportedPublication.into(),
        engine::Error::IdentifierOverflow => Refusal::IdentifierOverflow.into(),
        engine::Error::Graph(error) => match *error {
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
            _ => Error::InvalidSource,
        },
        engine::Error::Codec(error) => match error {
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
        engine::Error::Wire(error) => match error {
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

fn archive_limits(package: &Package) -> Result<litchi_iwa_core::Limits, Error> {
    package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)
}

fn engine_limits(package: &Package, budget: &ArrangementBudget) -> Result<engine::Limits, Error> {
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
    let input = budget
        .residual_input_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes)
        .min(litchi_iwa_protos::chart_style_codec::Limits::DEFAULT.max_bytes())
        .max(1);
    let output = budget
        .residual_output_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes)
        .min(litchi_iwa_protos::chart_style_codec::Limits::DEFAULT.max_bytes())
        .max(1);
    let fields = wire
        .max_fields()
        .min(litchi_iwa_protos::chart_style_codec::Limits::DEFAULT.max_fields())
        .max(1);
    let work = budget
        .residual_work_bytes()
        .map_err(map_arrangement_error)?
        .min(total_bytes.saturating_mul(32).max(1))
        .max(1);
    let allocations = budget
        .residual_allocations()
        .map_err(map_arrangement_error)?
        .max(1);
    let retained = budget
        .residual_retained_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes)
        .max(1);
    let scratch = budget
        .residual_scratch_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes)
        .max(1);
    Ok(engine::Limits::new(
        input,
        fields,
        output,
        u32::try_from(wire.max_nesting()).map_err(|_| Error::InvalidSource)?,
        work,
        package_bytes.min(litchi_iwa_protos::chart_style_codec::Limits::DEFAULT.max_bytes()),
        allocations,
        retained,
        scratch,
    ))
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

fn target_with_budget(
    package: &Package,
    selector: BodyChartSelector,
    budget: &mut ArrangementBudget,
) -> Result<ChartTarget, Error> {
    arrangement::resolve_target(package, selector, budget).map_err(map_arrangement_error)
}

fn charge_engine_report(
    budget: &mut ArrangementBudget,
    report: engine::Report,
) -> Result<(), Error> {
    budget
        .input(report.input_bytes())
        .map_err(map_arrangement_error)?;
    budget
        .output(report.output_bytes())
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
        .map_err(map_arrangement_error)?;
    budget
        .scratch(report.scratch_bytes())
        .map_err(map_arrangement_error)?;
    Ok(())
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
        .map_err(map_arrangement_error)?;
    Ok(())
}

fn archive_rewrite_bound(package: &Package, archive: &Archive) -> Result<usize, Error> {
    let encoded = archive
        .encoded_len(archive_limits(package)?)
        .map_err(map_core_error)?;
    let compressed = package
        .state
        .source
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    SnappyStream::maximum_compressed_len(encoded, compressed).map_err(map_core_error)
}

fn preflight_publication(
    source: &Package,
    edits: &[(String, Archive)],
    budget: &ArrangementBudget,
) -> Result<usize, Error> {
    let projected = edits.iter().try_fold(
        source.state.source.shared_source().len(),
        |total, (name, archive)| {
            let bound = archive_rewrite_bound(source, archive)?;
            total
                .checked_add(name.len())
                .and_then(|total| total.checked_add(bound))
                .ok_or(Error::InvalidSource)
        },
    )?;
    budget
        .preflight_output(projected)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_work(projected)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_allocations(edits.len().saturating_add(1))
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
        .allocations(edit_count.saturating_add(1))
        .map_err(map_arrangement_error)?;
    budget.retained(actual).map_err(map_arrangement_error)?;
    budget.scratch(projected).map_err(map_arrangement_error)
}

fn value_with_budget(
    package: &Package,
    target: &ChartTarget,
    kind: Kind,
    budget: &mut ArrangementBudget,
) -> Result<Value, Error> {
    admit_workspace(package, budget)?;
    let components = PackageComponents(package);
    let limits = archive_limits(package)?;
    let workspace = Workspace::new(&components, limits).map_err(map_engine_graph_error)?;
    let _payload =
        arrangement::selected_chart_payload(package, target).map_err(map_arrangement_error)?;
    let limits = engine_limits(package, budget)?;
    match kind {
        Kind::Font => {
            let result = engine::font(&workspace, target.drawable_identifier.get(), limits);
            let (value, report) = result.map_err(map_engine_error)?;
            charge_engine_report(budget, report)?;
            Ok(Value::Font(value))
        },
        Kind::Size => {
            let result = engine::size(&workspace, target.drawable_identifier.get(), limits);
            let (value, report) = result.map_err(map_engine_error)?;
            charge_engine_report(budget, report)?;
            Ok(Value::Size(value))
        },
    }
}

fn map_engine_graph_error(error: style_graph::Error) -> Error {
    map_engine_error(engine::Error::Graph(Box::new(error)))
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
                package_graph::GraphLimit::SemanticTextBytes => LimitKind::TextBytes,
            },
            observed,
            maximum,
        },
    }
}

/// Staged chart-font or chart-point-size edit against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    target: ChartTarget,
    kind: Kind,
    before: Value,
    budget: ArrangementBudget,
    change: Option<Change>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.target.position)
            .field("kind", &self.kind)
            .field("staged", &self.change.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage a new chart-wide font identity and face-trait value.
    pub fn set_font(mut self, font: ChartFont) -> Result<Self, Error> {
        if self.kind != Kind::Font {
            return Err(Refusal::WrongProperty.into());
        }
        self.change = Some(Change::Set(Value::Font(font)));
        Ok(self)
    }

    /// Stage removal of chart-wide font identity and face-trait overrides.
    pub fn reset_font(mut self) -> Result<Self, Error> {
        if self.kind != Kind::Font {
            return Err(Refusal::WrongProperty.into());
        }
        self.change = Some(Change::Reset);
        Ok(self)
    }

    /// Stage a new chart-wide point size.
    pub fn set_size(mut self, size: TextPointSize) -> Result<Self, Error> {
        if self.kind != Kind::Size {
            return Err(Refusal::WrongProperty.into());
        }
        self.change = Some(Change::Set(Value::Size(size)));
        Ok(self)
    }

    /// Stage removal of chart-wide point-size overrides.
    pub fn reset_size(mut self) -> Result<Self, Error> {
        if self.kind != Kind::Size {
            return Err(Refusal::WrongProperty.into());
        }
        self.change = Some(Change::Reset);
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

/// Exact-source reversible chart-font patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    position: Position,
    target: ChartTarget,
    kind: Kind,
    before: Value,
    after: Value,
    changed_components: Arc<[Arc<str>]>,
    changed_objects: Vec<(Arc<str>, Vec<u64>)>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether this patch preserves the exact source bytes.
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
            kind: self.kind,
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
    /// Whether candidate bytes or semantic value changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of IWA components rewritten by the graph publication.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
}

/// Fully reopened immutable result of one chart-font transaction.
#[must_use = "a chart-font commit contains the validated package snapshot"]
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

fn unchanged(source: &Package, target: ChartTarget, kind: Kind, value: Value) -> Commit {
    let bytes = source.state.source.shared_source();
    let owner = SharedBytes::from_shared_slice(Arc::clone(&bytes));
    Commit {
        package: source.clone(),
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
            position: target.position,
            target,
            kind,
            before: value.clone(),
            after: value,
            changed_components: Arc::from([]),
            changed_objects: Vec::new(),
        },
        diagnostics: Diagnostics {
            changed: false,
            touched_components: 0,
        },
    }
}

fn commit_edit(edit: Edit<'_>, change: Change) -> Result<Commit, Error> {
    let source = edit.source;
    let mut budget = edit.budget;
    let before = edit.before;
    let requested = match (&change, edit.kind) {
        (Change::Set(value), Kind::Font) if matches!(value, Value::Font(_)) => Some(value.clone()),
        (Change::Set(value), Kind::Size) if matches!(value, Value::Size(_)) => Some(value.clone()),
        (Change::Reset, _) => None,
        _ => return Err(Refusal::WrongProperty.into()),
    };
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    if requested.as_ref().is_some_and(|value| value == &before) {
        return Ok(unchanged(source, edit.target, edit.kind, before));
    }

    let components = PackageComponents(source);
    admit_workspace(source, &mut budget)?;
    let limits = archive_limits(source)?;
    let mut workspace = Workspace::new(&components, limits).map_err(map_engine_graph_error)?;
    let _payload =
        arrangement::selected_chart_payload(source, &edit.target).map_err(map_arrangement_error)?;
    let report = engine::set(
        &mut workspace,
        edit.target.drawable_identifier.get(),
        &match (&change, edit.kind) {
            (Change::Set(Value::Font(font)), Kind::Font) => engine::Change::Font(font.clone()),
            (Change::Set(Value::Size(size)), Kind::Size) => engine::Change::Size(*size),
            (Change::Reset, Kind::Font) => engine::Change::ResetFont,
            (Change::Reset, Kind::Size) => engine::Change::ResetSize,
            _ => return Err(Refusal::WrongProperty.into()),
        },
        engine_limits(source, &budget)?,
    )
    .map_err(map_engine_error)?;
    charge_engine_report(&mut budget, report)?;
    if !report.changed() {
        return Ok(unchanged(source, edit.target, edit.kind, before));
    }
    let changes = workspace.finish();
    if !changes.insertions.is_empty() || !changes.deletions.is_empty() {
        return Err(Refusal::UnsupportedPublication.into());
    }
    let edit_count = changes.edits.len();
    let projected = preflight_publication(source, &changes.edits, &budget)?;
    let names = changes
        .edits
        .iter()
        .map(|(name, _)| Arc::<str>::from(name.as_str()))
        .collect::<Vec<_>>();
    let archives = changes
        .edits
        .into_iter()
        .map(|(name, archive)| (Arc::<str>::from(name), archive))
        .collect();
    let candidate =
        package_graph::publish_archives(source, archives).map_err(map_package_graph_error)?;
    charge_publication(&candidate, edit_count, projected, &mut budget)?;
    let after = value_with_budget(&candidate, &edit.target, edit.kind, &mut budget)?;
    if requested
        .as_ref()
        .is_some_and(|requested| requested != &after)
    {
        return Err(Error::Verification);
    }
    let changed_objects = changed_object_ids(source, &candidate, &names)?;
    verify_locality(source, &candidate, &names, &changed_objects, &mut budget)?;
    let touched_components = names.len();
    let source_bytes = source.state.source.shared_source();
    let target_bytes = candidate.state.source.shared_source();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(
                SharedBytes::from_shared_slice(source_bytes),
                SharedBytes::from_shared_slice(Arc::clone(&target_bytes)),
            ),
            position: edit.target.position,
            target: edit.target,
            kind: edit.kind,
            before,
            after,
            changed_components: names.into(),
            changed_objects,
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_components,
        },
    })
}

fn apply_patch(package: &Package, patch: &Patch, kind: Kind) -> Result<Commit, Error> {
    if patch.kind != kind {
        return Err(Refusal::WrongProperty.into());
    }
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
    let current = target_with_budget(
        package,
        BodyChartSelector::position(patch.position),
        &mut budget,
    )?;
    if !current.same_identity(&patch.target)
        || value_with_budget(package, &current, kind, &mut budget)? != patch.before
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
    let verified = target_with_budget(
        &candidate,
        BodyChartSelector::position(patch.position),
        &mut budget,
    )?;
    if !verified.same_identity(&patch.target)
        || value_with_budget(&candidate, &verified, kind, &mut budget)? != patch.after
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
    /// Read the selected body chart's uniform effective font.
    pub fn body_chart_font(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<ChartFont, Error> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = target_with_budget(self, selector.into(), &mut budget)?;
        match value_with_budget(self, &target, Kind::Font, &mut budget)? {
            Value::Font(font) => Ok(font),
            Value::Size(_) => Err(Error::InvalidSource),
        }
    }

    /// Begin a selector-first body chart font transaction.
    pub fn edit_body_chart_font(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = target_with_budget(self, selector, &mut budget)?;
        let before = value_with_budget(self, &target, Kind::Font, &mut budget)?;
        Ok(Edit {
            source: self,
            target,
            kind: Kind::Font,
            before,
            budget,
            change: None,
        })
    }

    /// Apply an exact-source body chart font patch.
    pub fn apply_body_chart_font(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch, Kind::Font)
    }

    /// Read the selected body chart's uniform effective point size.
    pub fn body_chart_font_size(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<TextPointSize, Error> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = target_with_budget(self, selector.into(), &mut budget)?;
        match value_with_budget(self, &target, Kind::Size, &mut budget)? {
            Value::Size(size) => Ok(size),
            Value::Font(_) => Err(Error::InvalidSource),
        }
    }

    /// Begin a selector-first body chart point-size transaction.
    pub fn edit_body_chart_font_size(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = target_with_budget(self, selector, &mut budget)?;
        let before = value_with_budget(self, &target, Kind::Size, &mut budget)?;
        Ok(Edit {
            source: self,
            target,
            kind: Kind::Size,
            before,
            budget,
            change: None,
        })
    }

    /// Apply an exact-source body chart point-size patch.
    pub fn apply_body_chart_font_size(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch, Kind::Size)
    }
}

fn changed_object_ids(
    source: &Package,
    candidate: &Package,
    changed_components: &[Arc<str>],
) -> Result<Vec<(Arc<str>, Vec<u64>)>, Error> {
    let mut result = Vec::new();
    for name in changed_components {
        let before = source
            .state
            .source
            .components()
            .get(name.as_ref())
            .ok_or(Error::Verification)?;
        let after = candidate
            .state
            .source
            .components()
            .get(name.as_ref())
            .ok_or(Error::Verification)?;
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
    let source_entries = source.state.source.package();
    let candidate_entries = candidate.state.source.package();
    if source_entries.len() != candidate_entries.len() {
        return Err(Error::Verification);
    }
    for (before, after) in source_entries.iter().zip(candidate_entries.iter()) {
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
        if changed_components
            .iter()
            .any(|name| name.as_ref() == before.name())
        {
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
            || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(42, 46)],
            )
            || before.raw_record().compressed_data() != after.raw_record().compressed_data()
        {
            return Err(Error::Verification);
        }
    }

    let source_components = source.state.source.components();
    let candidate_components = candidate.state.source.components();
    if source_components.len() != candidate_components.len() {
        return Err(Error::Verification);
    }
    for (before_component, after_component) in
        source_components.iter().zip(candidate_components.iter())
    {
        if before_component.name() != after_component.name() {
            return Err(Error::Verification);
        }
        let changed = changed_components
            .iter()
            .any(|name| name.as_ref() == before_component.name());
        if !changed {
            if before_component.archive().objects.len() != after_component.archive().objects.len()
                || before_component
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
        // Every source object that is not one of the graph's changed objects
        // must remain byte-identical. New style objects are allowed only in a
        // component explicitly returned by the shared workspace.
        let allowed = changed_objects
            .iter()
            .find(|(name, _)| name.as_ref() == before_component.name())
            .map_or(&[][..], |(_, objects)| objects.as_slice());
        for before in &before_component.archive().objects {
            let Some(identifier) = before.archive_info.identifier else {
                return Err(Error::Verification);
            };
            let Some(after) = after_component.archive().object(identifier) else {
                // The inverse of an allocation legitimately removes the
                // graph objects created by the forward publication.
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
        for after in &after_component.archive().objects {
            let Some(identifier) = after.archive_info.identifier else {
                return Err(Error::Verification);
            };
            if before_component.archive().object(identifier).is_none()
                && !allowed.contains(&identifier)
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

#[cfg(test)]
mod tests {
    use super::*;

    const NATIVE: &[u8] =
        include_bytes!("../../../../test-data/iwork/pages/chart-data-native.pages");

    #[test]
    fn mixed_font_source_refuses_uniform_read_without_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let source = package.state.source.shared_source();
        assert!(matches!(
            package.body_chart_font(BodyChartSelector::index(0)),
            Err(Error::InvalidSource)
        ));
        assert_eq!(
            package.state.source.shared_source().as_ref(),
            source.as_ref()
        );
    }

    #[test]
    fn mixed_size_source_refuses_uniform_read_without_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let source = package.state.source.shared_source();
        assert!(matches!(
            package.body_chart_font_size(BodyChartSelector::index(0)),
            Err(Error::InvalidSource)
        ));
        assert_eq!(
            package.state.source.shared_source().as_ref(),
            source.as_ref()
        );
    }

    fn append_unknown_marker(data: &mut Vec<u8>, marker: u64) {
        litchi_iwa_common::wire::append_varint_field(data, 12_345, marker)
            .expect("bounded unknown marker");
    }

    fn has_unknown_marker(data: &[u8], marker: u64) -> bool {
        let mut suffix = litchi_iwa_common::varint::encode_varint(12_345_u64 << 3);
        suffix.extend(litchi_iwa_common::varint::encode_varint(marker));
        data.windows(suffix.len()).any(|window| window == suffix)
    }

    #[test]
    fn production_locality_preserves_unknown_changed_and_untouched_objects() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let component = package
            .state
            .source
            .components()
            .iter()
            .find(|component| {
                component
                    .archive()
                    .objects
                    .iter()
                    .filter(|object| !object.messages.is_empty())
                    .count()
                    >= 2
            })
            .expect("fixture component with two message-bearing objects");
        let component_name = Arc::<str>::from(component.name());
        let object_ids = component
            .archive()
            .objects
            .iter()
            .filter_map(|object| {
                (!object.messages.is_empty())
                    .then(|| object.archive_info.identifier)
                    .flatten()
            })
            .take(2)
            .collect::<Vec<_>>();
        let changed_id = object_ids[0];
        let untouched_id = object_ids[1];

        let mut injected_archive = component.archive().clone();
        append_unknown_marker(
            &mut injected_archive
                .object_mut(changed_id)
                .expect("changed object")
                .messages[0]
                .data,
            0xfeed,
        );
        append_unknown_marker(
            &mut injected_archive
                .object_mut(untouched_id)
                .expect("untouched object")
                .messages[0]
                .data,
            0xbeef,
        );
        let injected = package_graph::publish_archives(
            &package,
            vec![(Arc::clone(&component_name), injected_archive)],
        )
        .expect("unknown-field source injection");

        let mut changed_archive = injected
            .state
            .source
            .components()
            .get(component_name.as_ref())
            .expect("injected component")
            .archive()
            .clone();
        append_unknown_marker(
            &mut changed_archive
                .object_mut(changed_id)
                .expect("changed object")
                .messages[0]
                .data,
            0xcafe,
        );
        let changed = package_graph::publish_archives(
            &injected,
            vec![(Arc::clone(&component_name), changed_archive)],
        )
        .expect("known-field candidate mutation");
        let changed_objects =
            changed_object_ids(&injected, &changed, std::slice::from_ref(&component_name))
                .expect("changed-object witness");
        let changed_message = changed
            .state
            .source
            .components()
            .get(component_name.as_ref())
            .expect("changed component")
            .archive()
            .object(changed_id)
            .expect("changed object")
            .messages[0]
            .data
            .as_slice();
        let untouched_message = changed
            .state
            .source
            .components()
            .get(component_name.as_ref())
            .expect("changed component")
            .archive()
            .object(untouched_id)
            .expect("untouched object")
            .messages[0]
            .data
            .as_slice();
        assert!(has_unknown_marker(changed_message, 0xfeed));
        assert!(has_unknown_marker(untouched_message, 0xbeef));
        let mut budget = ArrangementBudget::new(&injected).expect("locality budget");
        verify_locality(
            &injected,
            &changed,
            std::slice::from_ref(&component_name),
            &changed_objects,
            &mut budget,
        )
        .expect("unknown fields are preserved through production locality");

        let mut negative_archive = changed
            .state
            .source
            .components()
            .get(component_name.as_ref())
            .expect("changed component")
            .archive()
            .clone();
        append_unknown_marker(
            &mut negative_archive
                .object_mut(untouched_id)
                .expect("untouched object")
                .messages[0]
                .data,
            0xabcd,
        );
        let negative = package_graph::publish_archives(
            &changed,
            vec![(Arc::clone(&component_name), negative_archive)],
        )
        .expect("negative candidate mutation");
        let mut negative_budget = ArrangementBudget::new(&injected).expect("locality budget");
        assert!(matches!(
            verify_locality(
                &injected,
                &negative,
                std::slice::from_ref(&component_name),
                &changed_objects,
                &mut negative_budget,
            ),
            Err(Error::Verification)
        ));
    }
}
