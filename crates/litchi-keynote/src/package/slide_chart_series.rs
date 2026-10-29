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

use std::collections::HashSet;
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion, ExactArtifacts};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceVisitor, Limits as ArchiveLimits, RawMessage,
    SnappyStream,
};
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions};
use litchi_iwa_protos::chart_style_codec::{self as chart_style_codec, Limits as ChartCodecLimits};
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

fn admit_workspace(package: &Package, budget: &mut TransactionBudget) -> Result<(), Error> {
    let components = package.state.source.components();
    let component_count = components.len();
    let object_count = package.state.total_objects;
    let index_units = object_count
        .checked_mul(4)
        .and_then(|value| value.checked_add(component_count))
        .ok_or(Error::InvalidSource)?;
    let index_bytes = index_units
        .checked_mul(size_of::<(u64, usize)>() * 2)
        .ok_or(Error::InvalidSource)?;
    budget
        .components(component_count)
        .and_then(|_| budget.payload_objects(object_count))
        .and_then(|_| {
            budget.work(
                component_count
                    .checked_add(object_count)
                    .ok_or(TransactionError::InvalidSource)?,
            )
        })
        .and_then(|_| budget.allocations(index_units))
        .and_then(|_| budget.retained(index_bytes))
        .and_then(|_| budget.scratch(index_bytes))
        .map_err(map_transaction_error)
}

fn admit_symbol_result(series: usize, budget: &mut TransactionBudget) -> Result<(), Error> {
    let bytes = series
        .checked_mul(size_of::<Option<Symbol>>())
        .ok_or(Error::InvalidSource)?;
    budget
        .allocations(1)
        .and_then(|_| budget.retained(bytes))
        .and_then(|_| budget.scratch(bytes))
        .and_then(|_| budget.work(series))
        .map_err(map_transaction_error)
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

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_core_limit(kind),
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

fn planned_archive_bytes(archive: &Archive) -> Result<usize, Error> {
    archive
        .objects
        .iter()
        .try_fold(0_usize, |total, object: &ArchiveObject| {
            let payload = object.messages.iter().try_fold(
                0_usize,
                |message_total, message: &RawMessage| {
                    message_total
                        .checked_add(message.data.len())
                        .and_then(|value| value.checked_add(size_of::<RawMessage>()))
                        .ok_or(Error::InvalidSource)
                },
            )?;
            total
                .checked_add(payload)
                .and_then(|value| value.checked_add(size_of::<ArchiveObject>()))
                .ok_or(Error::InvalidSource)
        })
}

fn planned_compressed_bytes(
    archive: &Archive,
    snappy_limits: litchi_iwa_core::SnappyLimits,
) -> Result<usize, Error> {
    let raw = planned_archive_bytes(archive)?;
    SnappyStream::maximum_compressed_len(raw, snappy_limits).map_err(map_core_error)
}

fn publish_changes(
    source: &Package,
    changes: graph::Changes,
    selected_chart: u64,
    budget: &mut TransactionBudget,
) -> Result<(Package, usize), Error> {
    let touched_components = changes
        .edits
        .len()
        .checked_add(changes.insertions.len())
        .and_then(|count| count.checked_add(changes.deletions.len()))
        .ok_or(Error::InvalidSource)?;
    let catalog = match &source.state.source {
        PhysicalSource::Package(catalog) => catalog,
        PhysicalSource::Semantic(_) => return Err(Refusal::UnsupportedSource.into()),
    };
    let archive_limits = source
        .structure_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let snappy_limits = catalog
        .limits()
        .snappy_limits()
        .map_err(|_| Error::InvalidSource)?;
    let planned_bytes =
        changes
            .edits
            .iter()
            .try_fold(catalog.source_bytes().len(), |total, (name, archive)| {
                let archive_bytes = planned_compressed_bytes(archive, snappy_limits)?;
                total
                    .checked_add(name.len())
                    .and_then(|total| total.checked_add(archive_bytes))
                    .ok_or(Error::InvalidSource)
            })?;
    let planned_bytes = changes
        .insertions
        .iter()
        .try_fold(planned_bytes, |total, insertion| {
            let archive_bytes = planned_compressed_bytes(&insertion.archive, snappy_limits)?;
            total
                .checked_add(insertion.name.len())
                .and_then(|total| total.checked_add(archive_bytes))
                .ok_or(Error::InvalidSource)
        })?;
    budget
        .preflight_reassembly(catalog, planned_bytes, changes.deletions.len())
        .map_err(map_transaction_error)?;
    let collection_allocations = changes
        .edits
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_add(changes.insertions.len().checked_mul(2)?))
        .and_then(|count| count.checked_add(catalog.package().len()))
        .and_then(|count| count.checked_add(changes.deletions.len()))
        .ok_or(Error::InvalidSource)?;
    let collection_name_bytes = changes
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
        .try_fold(0_usize, |total, bytes| total.checked_add(bytes))
        .ok_or(Error::InvalidSource)?;
    budget
        .allocations(collection_allocations)
        .and_then(|_| budget.work(collection_name_bytes))
        .map_err(map_transaction_error)?;
    let mut edited = Vec::new();
    edited
        .try_reserve_exact(changes.edits.len())
        .map_err(|_| Error::Allocation {
            amount: changes.edits.len(),
        })?;
    for (name, archive) in &changes.edits {
        let bytes = archive
            .to_bytes(archive_limits)
            .map_err(|_| Error::InvalidSource)?;
        let compressed =
            SnappyStream::compress(&bytes, snappy_limits).map_err(|_| Error::InvalidSource)?;
        edited.push((name.as_str(), compressed));
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
            .map(|(name, bytes)| EntryEdit::new(name, bytes.as_slice())),
    );
    let mut inserted = Vec::new();
    inserted
        .try_reserve_exact(changes.insertions.len())
        .map_err(|_| Error::Allocation {
            amount: changes.insertions.len(),
        })?;
    for insertion in &changes.insertions {
        let bytes = insertion
            .archive
            .to_bytes(archive_limits)
            .map_err(|_| Error::InvalidSource)?;
        let compressed =
            SnappyStream::compress(&bytes, snappy_limits).map_err(|_| Error::InvalidSource)?;
        inserted.push((
            insertion.name.as_str(),
            compressed,
            insertion.before.as_deref(),
        ));
    }
    let mut insertions = Vec::new();
    insertions
        .try_reserve_exact(inserted.len())
        .map_err(|_| Error::Allocation {
            amount: inserted.len(),
        })?;
    let names = catalog
        .package()
        .iter()
        .map(|entry| entry.name())
        .collect::<Vec<_>>();
    for (name, bytes, before) in &inserted {
        let insertion = match before {
            Some(before) => {
                let position = names
                    .iter()
                    .position(|candidate| *candidate == *before)
                    .ok_or(Error::InvalidSource)?;
                EntryInsertion::at(position, name, bytes.as_slice())
            },
            None => EntryInsertion::new(name, bytes.as_slice()),
        };
        insertions.push(insertion);
    }
    let deleted = changes
        .deletions
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    let bytes = source
        .structure_reassemble(&insertions, &edits, &deleted)
        .map_err(|_| Error::InvalidSource)?;
    budget.output(bytes.len()).map_err(map_transaction_error)?;
    budget
        .preflight_candidate(source, bytes.len())
        .map_err(map_transaction_error)?;
    let bytes_len = bytes.len();
    let candidate = source
        .structure_reopen(bytes)
        .map_err(|_| Error::Verification)?;
    if bytes_len == 0 {
        return Err(Error::Verification);
    }
    verify_planned_locality(source, &candidate, &changes, selected_chart, budget)?;
    verify_chart_unknowns(source, &candidate, selected_chart, budget)?;
    verify_member_locality(source, &candidate, &changes)?;
    Ok((candidate, touched_components))
}

fn chart_payload(package: &Package, identifier: u64) -> Result<&[u8], Error> {
    package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| {
            component.archive().object(identifier).and_then(|object| {
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == 5_021)
                    .map(|message| message.data.as_slice())
            })
        })
        .ok_or(Error::Verification)
}

fn verify_wire_preservation(
    source: &[u8],
    candidate: &[u8],
    excluded: impl Fn(u32) -> bool,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let work = source
        .len()
        .checked_add(candidate.len())
        .ok_or(Error::InvalidSource)?;
    budget
        .input(source.len())
        .and_then(|_| budget.input(candidate.len()))
        .and_then(|_| budget.work(work))
        .map_err(map_transaction_error)?;
    let source_fields = litchi_iwa_common::wire::parse_wire_fields(source, limits)
        .map_err(|_| Error::InvalidSource)?;
    let candidate_fields = litchi_iwa_common::wire::parse_wire_fields(candidate, limits)
        .map_err(|_| Error::InvalidSource)?;
    let fields = source_fields
        .len()
        .checked_add(candidate_fields.len())
        .ok_or(Error::InvalidSource)?;
    budget.fields(fields).map_err(map_transaction_error)?;
    let mut source_iter = source_fields
        .iter()
        .filter(|field| !excluded(field.number()));
    let mut candidate_iter = candidate_fields
        .iter()
        .filter(|field| !excluded(field.number()));
    loop {
        match (source_iter.next(), candidate_iter.next()) {
            (None, None) => return Ok(()),
            (Some(source_field), Some(candidate_field)) => {
                if source_field.number() != candidate_field.number()
                    || source_field.wire_type() != candidate_field.wire_type()
                    || source_field.raw(source).map_err(|_| Error::InvalidSource)?
                        != candidate_field
                            .raw(candidate)
                            .map_err(|_| Error::InvalidSource)?
                {
                    return Err(Error::Verification);
                }
            },
            _ => return Err(Error::Verification),
        }
    }
}

fn verify_chart_unknowns(
    source: &Package,
    candidate: &Package,
    identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let source_payload = chart_payload(source, identifier)?;
    let candidate_payload = chart_payload(candidate, identifier)?;
    let limits = budget.residual(source).map_err(map_transaction_error)?;
    verify_wire_preservation(
        source_payload,
        candidate_payload,
        |number| number == 10_000,
        limits,
        budget,
    )?;
    let source_extension = litchi_iwa_common::wire::parse_wire_fields(source_payload, limits)
        .map_err(|_| Error::InvalidSource)?
        .into_iter()
        .find(|field| field.number() == 10_000)
        .ok_or(Error::InvalidSource)?
        .payload(source_payload)
        .map_err(|_| Error::InvalidSource)?;
    let candidate_extension = litchi_iwa_common::wire::parse_wire_fields(candidate_payload, limits)
        .map_err(|_| Error::InvalidSource)?
        .into_iter()
        .find(|field| field.number() == 10_000)
        .ok_or(Error::InvalidSource)?
        .payload(candidate_payload)
        .map_err(|_| Error::InvalidSource)?;
    verify_wire_preservation(
        source_extension,
        candidate_extension,
        |number| number == 17 || number == 19,
        limits,
        budget,
    )
}

#[derive(Default)]
struct ObjectReferenceCollector {
    identifiers: Vec<u64>,
}

impl ArchiveReferenceVisitor for ObjectReferenceCollector {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind == ArchiveReferenceKind::Object && occurrence.referenced_identifier != 0
        {
            self.identifiers.push(occurrence.referenced_identifier);
        }
        Ok(())
    }
}

fn package_object(package: &Package, identifier: u64) -> Option<&ArchiveObject> {
    package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| component.archive().object(identifier))
}

fn reachable_objects(
    package: &Package,
    root: u64,
    budget: &mut TransactionBudget,
) -> Result<HashSet<u64>, Error> {
    let mut pending = vec![root];
    let mut reachable = HashSet::new();
    reachable
        .try_reserve(package.state.total_objects)
        .map_err(|_| Error::Allocation {
            amount: package.state.total_objects,
        })?;
    let limits = graph_limits(package)?;
    while let Some(identifier) = pending.pop() {
        if !reachable.insert(identifier) {
            continue;
        }
        let object = package_object(package, identifier).ok_or(Error::InvalidSource)?;
        let work =
            object
                .messages
                .iter()
                .try_fold(size_of::<ArchiveObject>(), |total, message| {
                    total
                        .checked_add(size_of::<RawMessage>())
                        .and_then(|value| value.checked_add(message.data.len()))
                        .ok_or(Error::InvalidSource)
                })?;
        budget
            .work(work)
            .and_then(|_| budget.allocations(1))
            .and_then(|_| {
                budget.scratch(
                    object
                        .messages
                        .len()
                        .checked_mul(size_of::<u64>())
                        .ok_or(TransactionError::InvalidSource)?,
                )
            })
            .map_err(map_transaction_error)?;
        let mut references = ObjectReferenceCollector::default();
        object
            .inspect_references_with_policy(
                &mut references,
                ArchiveReferencePolicy::KnownReferences,
                limits,
            )
            .map_err(map_core_error)?;
        budget
            .references(references.identifiers.len())
            .and_then(|_| budget.work(references.identifiers.len()))
            .map_err(map_transaction_error)?;
        pending.extend(references.identifiers);
    }
    Ok(reachable)
}

fn reference_identifiers(payload: &[u8], field_number: u32) -> Result<Vec<u64>, Error> {
    let fields = litchi_iwa_common::wire::parse_wire_fields(
        payload,
        litchi_iwa_common::WireLimits::default(),
    )
    .map_err(|_| Error::InvalidSource)?;
    let mut identifiers = Vec::new();
    for field in fields.iter().filter(|field| field.number() == field_number) {
        let nested = field.payload(payload).map_err(|_| Error::InvalidSource)?;
        let nested_fields = litchi_iwa_common::wire::parse_wire_fields(
            nested,
            litchi_iwa_common::WireLimits::default(),
        )
        .map_err(|_| Error::InvalidSource)?;
        let Some(identifier) = nested_fields
            .iter()
            .find(|nested_field| nested_field.number() == 1 && nested_field.wire_type() == 0)
            .map(|nested_field| {
                litchi_iwa_common::decode_varint_from_bytes(
                    &nested[nested_field.key_end()..nested_field.end()],
                )
                .map(|(value, _)| value)
                .map_err(|_| Error::InvalidSource)
            })
            .transpose()?
        else {
            continue;
        };
        if identifier != 0 {
            identifiers.push(identifier);
        }
    }
    Ok(identifiers)
}

fn stylesheet_reachable_objects(package: &Package, chart: u64) -> Result<HashSet<u64>, Error> {
    let mut pending =
        chart_style_codec::references(chart_payload(package, chart)?, ChartCodecLimits::DEFAULT)
            .map_err(|_| Error::InvalidSource)?
            .series_theme_styles()
            .to_vec();
    let mut reachable = HashSet::new();
    while let Some(identifier) = pending.pop() {
        if identifier == 0 || !reachable.insert(identifier) {
            continue;
        }
        let Some(object) = package_object(package, identifier) else {
            return Err(Error::InvalidSource);
        };
        for message in &object.messages {
            if message.type_ == 5_028 || message.type_ == 5_029 {
                let parents = litchi_iwa_common::wire::parse_wire_fields(
                    &message.data,
                    litchi_iwa_common::WireLimits::default(),
                )
                .map_err(|_| Error::InvalidSource)?;
                for parent in parents.iter().filter(|field| field.number() == 1) {
                    let parent = parent
                        .payload(&message.data)
                        .map_err(|_| Error::InvalidSource)?;
                    pending.extend(reference_identifiers(parent, 3)?);
                    pending.extend(reference_identifiers(parent, 5)?);
                }
            }
        }
    }
    Ok(reachable)
}

fn changed_object_ids(source: &Package, changes: &graph::Changes) -> Result<HashSet<u64>, Error> {
    let mut changed = HashSet::new();
    for (name, archive) in &changes.edits {
        let source_archive = source
            .state
            .source
            .components()
            .get(name)
            .ok_or(Error::InvalidSource)?
            .archive();
        for object in &source_archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            if archive
                .objects
                .iter()
                .all(|candidate| candidate.archive_info.identifier != Some(identifier))
            {
                changed.insert(identifier);
            }
        }
        for object in &archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            if source_archive
                .object(identifier)
                .is_none_or(|before| !before.same_content_ignoring_offsets(object))
            {
                changed.insert(identifier);
            }
        }
    }
    for insertion in &changes.insertions {
        for object in &insertion.archive.objects {
            changed.insert(object.archive_info.identifier.ok_or(Error::InvalidSource)?);
        }
    }
    for name in &changes.deletions {
        if let Some(component) = source.state.source.components().get(name) {
            for object in &component.archive().objects {
                changed.insert(object.archive_info.identifier.ok_or(Error::InvalidSource)?);
            }
        }
    }
    Ok(changed)
}

fn verify_planned_locality(
    source: &Package,
    candidate: &Package,
    changes: &graph::Changes,
    root: u64,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let changed = changed_object_ids(source, changes)?;
    let mut candidate_reachable = reachable_objects(candidate, root, budget)?;
    candidate_reachable.extend(stylesheet_reachable_objects(candidate, root)?);
    for identifier in changed {
        if identifier == 2 {
            continue;
        }
        if package_object(candidate, identifier).is_some()
            && !candidate_reachable.contains(&identifier)
        {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn verify_member_locality(
    source: &Package,
    candidate: &Package,
    changes: &graph::Changes,
) -> Result<(), Error> {
    let source_catalog = match &source.state.source {
        PhysicalSource::Package(catalog) => catalog,
        PhysicalSource::Semantic(_) => return Err(Refusal::UnsupportedSource.into()),
    };
    let candidate_catalog = match &candidate.state.source {
        PhysicalSource::Package(catalog) => catalog,
        PhysicalSource::Semantic(_) => return Err(Error::Verification),
    };
    for entry in source_catalog.package().iter() {
        if matches!(
            entry.name(),
            "preview.jpg" | "preview-micro.jpg" | "preview-web.jpg"
        ) {
            continue;
        }
        let changed = changes.edits.iter().any(|(name, _)| name == entry.name())
            || changes.deletions.iter().any(|name| name == entry.name());
        if changed {
            continue;
        }
        let Some(reopened) = candidate_catalog
            .package()
            .iter()
            .find(|candidate| candidate.name() == entry.name())
        else {
            return Err(Error::Verification);
        };
        if reopened.data() != entry.data() {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

/// A compact exact-source patch for one chart-series symbol value.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    slide: Position,
    chart: Position,
    before: Property,
    after: Property,
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
                .map_err(|_| Error::Allocation { amount: retained })?;
        let chart =
            super::edit::OwnedChartSelector::from_selector(chart.into(), maximum, &mut retained)
                .map_err(|_| Error::Allocation { amount: retained })?;
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
        let before = read_symbols(self.source, selection, family, series, &mut budget)?;
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
        admit_workspace(self.source, &mut budget)?;
        let components = PackageComponents(self.source);
        let mut workspace =
            Workspace::new(&components, graph_limits(self.source)?).map_err(map_graph)?;
        let changed = engine::set_symbols(
            &mut workspace,
            selection.identifier,
            family,
            match &staged {
                Property::Symbols(values) => values,
            },
        )
        .map_err(map_graph)?;
        if !changed {
            return Err(Error::InvalidSource);
        }
        let (candidate, touched_components) = publish_changes(
            self.source,
            workspace.finish(),
            selection.identifier,
            &mut budget,
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
        let after = read_symbols(&candidate, candidate_selection, family, series, &mut budget)?;
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
) -> Result<Vec<Option<Symbol>>, Error> {
    admit_symbol_result(series, budget)?;
    admit_workspace(package, budget)?;
    let components = PackageComponents(package);
    let workspace = Workspace::new(&components, graph_limits(package)?).map_err(map_graph)?;
    engine::symbols(&workspace, selection.identifier, family, series).map_err(map_graph)
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
        let selection = select(
            &candidate,
            SlideSelector::position(patch.slide),
            ChartSelector::position(patch.chart),
            &mut budget,
            &mut fields,
        )?;
        verify_chart_unknowns(self, &candidate, selection.identifier, &mut budget)?;
        let (family, series) = family_and_count(&candidate, selection, &mut budget)?;
        let after = read_symbols(&candidate, selection, family, series, &mut budget)?;
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
