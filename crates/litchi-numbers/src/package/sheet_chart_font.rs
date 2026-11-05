//! Exact-source Numbers sheet-chart font transactions.
//!
//! The selector and package publication stay in this facade; paragraph-style
//! variation allocation, chart repointing, registry updates and pruning are
//! performed by [`litchi_iwa_style::chart_font`].  Font identity and point
//! size are one focused family because the native chart font graph stores
//! them in the same inherited paragraph-style variations.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused transaction keeps its public contract beside the bounded graph phases."
)]

use std::fmt;
use std::mem::size_of;

use litchi_core::Position;
use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_core::SnappyStream;
use litchi_iwa_style::chart_font::ChartFont;
use litchi_iwa_style::chart_font::{self, Change};
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_text::character::TextPointSize;
use thiserror::Error as ThisError;

use super::Package;
use super::chart_arrangement::{
    self as arrangement, ChartBudget, ChartSelection, select_chart_with_budget,
};
use super::structure::{self, Fault, FaultLimit};
use super::style_graph::{self, LeafComponents};
use crate::selector::{ChartSelector, SheetSelector};

const MAX_STAGED_OPERATIONS: usize = 4;
// `LeafComponents` owns one tuple per parsed component.  The graph workspace
// then builds a package-wide identifier index and an ambiguity set.  These
// are admission bounds for the constructor phase; the engine report charges
// its corresponding source traversal after construction.
const GRAPH_INDEX_SLOT_BYTES: usize = 64;
const GRAPH_AMBIGUOUS_SLOT_BYTES: usize = 32;
const GRAPH_EXTRA_ALLOCATIONS: usize = 4;
// ZIP local/central records plus a conservative name/EOCD envelope for a
// planned changed member.  Existing edits do not need this envelope, but
// charging it uniformly keeps the bound simple and safely overestimates them.
const ZIP_MEMBER_OVERHEAD_BYTES: usize = 128;

/// A semantic value addressed by a sheet-chart font transaction.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// The effective font identity and face traits.
    Font(ChartFont),
    /// The effective point size in typographic points.
    Size(TextPointSize),
}

/// The property named by a read or reset operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// Font identity and face traits.
    Font,
    /// Point size.
    Size,
}

impl Property {
    /// Return the property kind carried by this value.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::Font(_) => PropertyKind::Font,
            Self::Size(_) => PropertyKind::Size,
        }
    }
}

/// A finite resource governed by one Numbers chart-font transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// One decoded IWA component payload.
    PayloadBytes,
    /// Native chart references inspected by the selector.
    PayloadReferences,
    /// Parsed wire fields or graph work.
    WireWork,
    /// Candidate output bytes.
    OutputBytes,
    /// Bounded temporary allocations.
    Allocations,
    /// Source and candidate bytes retained together.
    RetainedBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
    /// Object identifiers allocated by the shared graph engine.
    Identifiers,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadReferences => "payload references",
            Self::WireWork => "wire work",
            Self::OutputBytes => "output bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::Identifiers => "object identifiers",
        })
    }
}

/// A well-formed request this family refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No sheet matched the semantic selector.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No chart matched the semantic selector.
    #[error("chart selector did not resolve")]
    ChartNotFound,
    /// The source does not retain an exact physical artifact.
    #[error("this Numbers source does not support exact chart-font edits")]
    UnsupportedSource,
    /// A chart-font edit was committed without a staged operation.
    #[error("no chart-font operation was staged")]
    NoStagedOperation,
}

impl Refusal {
    /// Classify the refusal for unified operations.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::ChartNotFound => RefusalKind::NotFound,
            Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
        }
    }
}

/// Failure from a Numbers chart-font read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted chart graph or paragraph-style source is malformed.
    #[error("the selected Numbers chart-font source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("Numbers chart font {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} Numbers chart-font units")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// Candidate reopen or semantic readback failed.
    #[error("the edited Numbers chart font failed semantic verification")]
    Verification,
    /// A patch was applied to a different exact source.
    #[error("the Numbers chart-font patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported => Self::Refused(Refusal::UnsupportedSource),
            Fault::Invalid => Self::InvalidSource,
            Fault::Dependency => Self::InvalidSource,
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

impl From<litchi_iwa_style::graph::Error> for Error {
    fn from(error: litchi_iwa_style::graph::Error) -> Self {
        match style_graph::fault(&error) {
            style_graph::StyleGraphFault::Shared(fault) => fault.into(),
            style_graph::StyleGraphFault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    litchi_iwa_style::graph::GraphLimitKind::ScratchBytes => {
                        LimitKind::ScratchBytes
                    },
                    litchi_iwa_style::graph::GraphLimitKind::AllocationEvents => {
                        LimitKind::Allocations
                    },
                },
                observed,
                maximum,
            },
        }
    }
}

fn map_engine_error(error: chart_font::Error) -> Error {
    use litchi_iwa_style::chart_font::{Error as EngineError, LimitKind as EngineLimitKind};
    match error {
        EngineError::InvalidSource(_) | EngineError::Refused(_) => Error::InvalidSource,
        EngineError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                EngineLimitKind::InputBytes
                | EngineLimitKind::TextBytes
                | EngineLimitKind::Wire(_)
                | EngineLimitKind::Codec(_)
                | EngineLimitKind::Archive(_) => LimitKind::PayloadBytes,
                EngineLimitKind::Fields | EngineLimitKind::Nesting | EngineLimitKind::Work => {
                    LimitKind::WireWork
                },
                EngineLimitKind::OutputBytes => LimitKind::OutputBytes,
                EngineLimitKind::Allocations => LimitKind::Allocations,
                EngineLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                EngineLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
        EngineError::Allocation { amount } => Error::Allocation { amount },
        EngineError::Graph(error) => (*error).into(),
        EngineError::Codec(_) | EngineError::Wire(_) => Error::InvalidSource,
        EngineError::IdentifierOverflow => Error::Limit {
            kind: LimitKind::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        },
        _ => Error::InvalidSource,
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        structure::map_wire(&error).into()
    }
}

/// The selected chart.  Effective values are intentionally read lazily: a
/// chart can have a uniform font while its paragraph styles carry mixed point
/// sizes (and vice versa), so resolving a selector must not read both
/// independent families up front.
#[derive(Debug, Clone, PartialEq)]
struct Target {
    sheet: Position,
    chart: Position,
    chart_identifier: u64,
}

#[derive(Debug, Clone, Copy)]
struct GraphWorkspaceCost {
    work: usize,
    allocations: usize,
    scratch: usize,
}

#[derive(Debug, Clone, Copy)]
struct CandidatePlan {
    output_bytes: usize,
    scratch_bytes: usize,
    allocations: usize,
}

fn graph_workspace_cost(package: &Package) -> Result<GraphWorkspaceCost, Error> {
    let catalog = package.state.components.catalog();
    let object_count = catalog.iter().try_fold(0usize, |count, component| {
        count
            .checked_add(component.archive().objects.len())
            .ok_or(Error::InvalidSource)
    })?;
    let component_count = catalog.len();
    let work = component_count
        .checked_add(object_count)
        .ok_or(Error::InvalidSource)?;
    let allocations = component_count
        .checked_add(object_count.checked_mul(2).ok_or(Error::InvalidSource)?)
        .and_then(|value| value.checked_add(GRAPH_EXTRA_ALLOCATIONS))
        .ok_or(Error::InvalidSource)?;
    let component_scratch = component_count
        .checked_mul(size_of::<(usize, usize)>())
        .ok_or(Error::InvalidSource)?;
    let index_scratch = object_count
        .checked_mul(GRAPH_INDEX_SLOT_BYTES + GRAPH_AMBIGUOUS_SLOT_BYTES)
        .ok_or(Error::InvalidSource)?;
    Ok(GraphWorkspaceCost {
        work,
        allocations,
        scratch: component_scratch
            .checked_add(index_scratch)
            .ok_or(Error::InvalidSource)?,
    })
}

fn preflight_graph_workspace(
    package: &Package,
    budget: &ChartBudget,
) -> Result<GraphWorkspaceCost, Error> {
    let cost = graph_workspace_cost(package)?;
    budget.preflight_work(cost.work).map_err(map_arrangement)?;
    budget
        .preflight_allocations(cost.allocations)
        .map_err(map_arrangement)?;
    budget
        .preflight_scratch(cost.scratch)
        .map_err(map_arrangement)?;
    Ok(cost)
}

fn charge_graph_workspace(budget: &mut ChartBudget, cost: GraphWorkspaceCost) -> Result<(), Error> {
    // Workspace::source_work is included in each engine report, so only the
    // constructor-owned collection allocations and scratch are charged here.
    budget
        .allocations(cost.allocations)
        .map_err(map_arrangement)?;
    budget.scratch(cost.scratch).map_err(map_arrangement)
}

fn archive_compressed_bound(
    archive: &litchi_iwa_core::Archive,
    archive_limits: litchi_iwa_core::Limits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
) -> Result<usize, Error> {
    let encoded = archive
        .encoded_len(archive_limits)
        .map_err(|error| Error::from(structure::map_core(&error)))?;
    SnappyStream::maximum_compressed_len(encoded, snappy_limits)
        .map_err(|error| Error::from(structure::map_core(&error)))
}

fn candidate_plan(
    package: &Package,
    changes: &litchi_iwa_style::graph::Changes,
) -> Result<CandidatePlan, Error> {
    let catalog = structure::exact_source(package)?;
    let archive_limits = archive_limits(package)?;
    let snappy_limits = catalog
        .limits()
        .snappy_limits()
        .map_err(|error| Error::from(structure::map_archive(&error)))?;
    let mut output_bytes = package.source_bytes().len();
    let mut compressed_bytes = 0usize;
    let mut allocations = GRAPH_EXTRA_ALLOCATIONS;
    let mut add_member = |name: &str, archive: &litchi_iwa_core::Archive| -> Result<(), Error> {
        let compressed = archive_compressed_bound(archive, archive_limits, snappy_limits)?;
        compressed_bytes = compressed_bytes
            .checked_add(compressed)
            .ok_or(Error::InvalidSource)?;
        let name_overhead = name
            .len()
            .checked_mul(2)
            .and_then(|value| value.checked_add(ZIP_MEMBER_OVERHEAD_BYTES))
            .ok_or(Error::InvalidSource)?;
        output_bytes = output_bytes
            .checked_add(compressed)
            .and_then(|value| value.checked_add(name_overhead))
            .ok_or(Error::InvalidSource)?;
        allocations = allocations.checked_add(1).ok_or(Error::InvalidSource)?;
        Ok(())
    };
    for (name, archive) in &changes.edits {
        add_member(name, archive)?;
    }
    for insertion in &changes.insertions {
        add_member(&insertion.name, &insertion.archive)?;
    }
    allocations = allocations
        .checked_add(changes.deletions.len())
        .ok_or(Error::InvalidSource)?;
    Ok(CandidatePlan {
        output_bytes,
        scratch_bytes: output_bytes
            .checked_add(compressed_bytes)
            .ok_or(Error::InvalidSource)?,
        allocations,
    })
}

fn map_arrangement(error: arrangement::Error) -> Error {
    match error {
        arrangement::Error::Refused(refusal) => match refusal {
            arrangement::Refusal::SheetNameNotFound
            | arrangement::Refusal::SheetPositionNotFound { .. }
            | arrangement::Refusal::EmptySheetName
            | arrangement::Refusal::AmbiguousSelector => Refusal::SheetNotFound.into(),
            arrangement::Refusal::ChartPositionNotFound { .. } => Refusal::ChartNotFound.into(),
            arrangement::Refusal::UnsupportedSource => Refusal::UnsupportedSource.into(),
            arrangement::Refusal::UnsupportedDependency => Error::InvalidSource,
        },
        arrangement::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                arrangement::LimitKind::References => LimitKind::PayloadReferences,
                arrangement::LimitKind::Identifiers => LimitKind::Identifiers,
                arrangement::LimitKind::WireBytes | arrangement::LimitKind::CompressedBytes => {
                    LimitKind::PayloadBytes
                },
                arrangement::LimitKind::WireFields
                | arrangement::LimitKind::WireNesting
                | arrangement::LimitKind::WireWork => LimitKind::WireWork,
                arrangement::LimitKind::OutputBytes => LimitKind::OutputBytes,
                arrangement::LimitKind::Allocations => LimitKind::Allocations,
                arrangement::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                arrangement::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
            },
            observed,
            maximum,
        },
        arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn archive_limits(package: &Package) -> Result<litchi_iwa_core::Limits, Error> {
    structure::exact_source(package)?
        .limits()
        .effective_archive_limits()
        .map_err(|error| structure::map_archive(&error).into())
}

fn read_workspace(
    workspace: &Workspace<'_>,
    chart: u64,
    kind: PropertyKind,
    limits: chart_font::Limits,
) -> Result<(Property, chart_font::Report), chart_font::Error> {
    Ok(match kind {
        PropertyKind::Font => {
            let (value, report) = chart_font::font(workspace, chart, limits)?;
            (Property::Font(value), report)
        },
        PropertyKind::Size => {
            let (value, report) = chart_font::size(workspace, chart, limits)?;
            (Property::Size(value), report)
        },
    })
}

fn target_from_selection(selection: &ChartSelection) -> Target {
    Target {
        sheet: selection.sheet_position,
        chart: selection.chart_position,
        chart_identifier: selection.chart_identifier,
    }
}

fn resolve_with_budget<'sheet>(
    package: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    chart: impl Into<ChartSelector>,
    budget: &mut ChartBudget,
) -> Result<Target, Error> {
    let selection = select_chart_with_budget(package, sheet.into(), chart.into(), false, budget)
        .map_err(map_arrangement)?;
    Ok(target_from_selection(&selection))
}

fn read_property(
    package: &Package,
    target: &Target,
    kind: PropertyKind,
    budget: &mut ChartBudget,
) -> Result<Property, Error> {
    let graph_cost = preflight_graph_workspace(package, budget)?;
    let components = LeafComponents::new(package).map_err(Error::from)?;
    let workspace = Workspace::new(&components, archive_limits(package)?)?;
    charge_graph_workspace(budget, graph_cost)?;
    let (value, report) = read_workspace(
        &workspace,
        target.chart_identifier,
        kind,
        budget.chart_font_limits(),
    )
    .map_err(map_engine_error)?;
    budget
        .charge_chart_font_report(report)
        .map_err(map_arrangement)?;
    Ok(value)
}

fn property_index(kind: PropertyKind) -> usize {
    match kind {
        PropertyKind::Font => 0,
        PropertyKind::Size => 1,
    }
}

fn change_kind(change: &Change) -> Result<PropertyKind, Error> {
    match change {
        Change::Font(_) | Change::ResetFont => Ok(PropertyKind::Font),
        Change::Size(_) | Change::ResetSize => Ok(PropertyKind::Size),
        _ => Err(Error::InvalidSource),
    }
}

fn change_from_property(property: Property) -> Change {
    match property {
        Property::Font(value) => Change::Font(value),
        Property::Size(value) => Change::Size(value),
    }
}

fn change_from_reset(kind: PropertyKind) -> Change {
    match kind {
        PropertyKind::Font => Change::ResetFont,
        PropertyKind::Size => Change::ResetSize,
    }
}

fn unchanged_commit(
    source: &Package,
    target: &Target,
    before_font: Option<ChartFont>,
    before_size: Option<TextPointSize>,
) -> Result<Commit, Error> {
    let owner = structure::source_owner(source)?;
    Ok(Commit {
        package: source.snapshot(),
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
            sheet: target.sheet,
            chart: target.chart,
            chart_identifier: target.chart_identifier,
            touched: [before_font.is_some(), before_size.is_some()],
            before_font: before_font.clone(),
            after_font: before_font,
            before_size,
            after_size: before_size,
        },
        diagnostics: Diagnostics { changed: false },
    })
}

/// Font changes staged against one chart of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    changes: Vec<Change>,
    budget: ChartBudget,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet_position", &self.target.sheet)
            .field("chart_position", &self.target.chart)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage a semantic font or point-size value. Changes apply in order.
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        if self.changes.len() >= MAX_STAGED_OPERATIONS {
            return Err(Error::Limit {
                kind: LimitKind::Allocations,
                observed: self.changes.len().saturating_add(1) as u64,
                maximum: MAX_STAGED_OPERATIONS as u64,
            });
        }
        self.budget.allocations(1).map_err(map_arrangement)?;
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(change_from_property(property));
        Ok(self)
    }

    /// Stage removal of the selected chart-font override.
    pub fn reset(&mut self, kind: PropertyKind) -> Result<&mut Self, Error> {
        if self.changes.len() >= MAX_STAGED_OPERATIONS {
            return Err(Error::Limit {
                kind: LimitKind::Allocations,
                observed: self.changes.len().saturating_add(1) as u64,
                maximum: MAX_STAGED_OPERATIONS as u64,
            });
        }
        self.budget.allocations(1).map_err(map_arrangement)?;
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(change_from_reset(kind));
        Ok(self)
    }

    /// Validate, publish and strictly reopen all staged changes.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut budget = self.budget;
        // `Edit` borrows an immutable source snapshot and `target` was
        // resolved when the edit was created. Re-running the package-wide
        // selector census here only spends the residual WireWork budget a
        // second time; candidate publication below still re-resolves and
        // read-backs the selected chart before returning a commit.
        let graph_cost = preflight_graph_workspace(self.source, &budget)?;
        let components = LeafComponents::new(self.source).map_err(Error::from)?;
        let mut workspace = Workspace::new(&components, archive_limits(self.source)?)?;
        charge_graph_workspace(&mut budget, graph_cost)?;
        let mut touched = [false; 2];
        let mut before_values: [Option<Property>; 2] = std::array::from_fn(|_| None);
        for change in &self.changes {
            let kind = change_kind(change)?;
            let index = property_index(kind);
            if !touched[index] {
                touched[index] = true;
                let (value, report) = read_workspace(
                    &workspace,
                    self.target.chart_identifier,
                    kind,
                    budget.chart_font_limits(),
                )
                .map_err(map_engine_error)?;
                budget
                    .charge_chart_font_report(report)
                    .map_err(map_arrangement)?;
                before_values[index] = Some(value);
            }
        }
        for change in &self.changes {
            let report = chart_font::set(
                &mut workspace,
                self.target.chart_identifier,
                change,
                budget.chart_font_limits(),
            )
            .map_err(map_engine_error)?;
            budget
                .charge_chart_font_report(report)
                .map_err(map_arrangement)?;
        }
        let mut after_values: [Option<Property>; 2] = std::array::from_fn(|_| None);
        for kind in [PropertyKind::Font, PropertyKind::Size] {
            let index = property_index(kind);
            if touched[index] {
                let (value, report) = read_workspace(
                    &workspace,
                    self.target.chart_identifier,
                    kind,
                    budget.chart_font_limits(),
                )
                .map_err(map_engine_error)?;
                budget
                    .charge_chart_font_report(report)
                    .map_err(map_arrangement)?;
                after_values[index] = Some(value);
            }
        }
        let changed = touched
            .iter()
            .enumerate()
            .any(|(index, selected)| *selected && before_values[index] != after_values[index]);
        let before_font = match before_values[0].take() {
            Some(Property::Font(value)) => Some(value),
            Some(Property::Size(_)) => return Err(Error::InvalidSource),
            None => None,
        };
        let after_font = match after_values[0].take() {
            Some(Property::Font(value)) => Some(value),
            Some(Property::Size(_)) => return Err(Error::InvalidSource),
            None => before_font.clone(),
        };
        let before_size = match before_values[1].take() {
            Some(Property::Size(value)) => Some(value),
            Some(Property::Font(_)) => return Err(Error::InvalidSource),
            None => None,
        };
        let after_size = match after_values[1].take() {
            Some(Property::Size(value)) => Some(value),
            Some(Property::Font(_)) => return Err(Error::InvalidSource),
            None => before_size,
        };
        if !changed {
            return unchanged_commit(self.source, &self.target, before_font, before_size);
        }
        let changes = workspace.finish();
        if changes.is_empty() {
            return Err(Error::Verification);
        }
        let source_bytes = self.source.source_bytes().len();
        let candidate_plan = candidate_plan(self.source, &changes)?;
        budget
            .preflight_output(candidate_plan.output_bytes)
            .map_err(map_arrangement)?;
        budget
            .preflight_retained(source_bytes)
            .and_then(|_| budget.preflight_retained(candidate_plan.output_bytes))
            .map_err(map_arrangement)?;
        budget
            .preflight_allocations(candidate_plan.allocations)
            .map_err(map_arrangement)?;
        budget
            .preflight_scratch(candidate_plan.scratch_bytes)
            .map_err(map_arrangement)?;
        budget
            .output(candidate_plan.output_bytes)
            .map_err(map_arrangement)?;
        budget.retained(source_bytes).map_err(map_arrangement)?;
        budget
            .retained(candidate_plan.output_bytes)
            .map_err(map_arrangement)?;
        budget
            .allocations(candidate_plan.allocations)
            .map_err(map_arrangement)?;
        budget
            .scratch(candidate_plan.scratch_bytes)
            .map_err(map_arrangement)?;
        let published = style_graph::publish(self.source, changes, &[]).map_err(Error::from)?;
        let actual_candidate_bytes = published.package.source_bytes().len();
        if actual_candidate_bytes > candidate_plan.output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: actual_candidate_bytes as u64,
                maximum: candidate_plan.output_bytes as u64,
            });
        }
        let candidate_target = resolve_with_budget(
            &published.package,
            SheetSelector::position(self.target.sheet),
            ChartSelector::position(self.target.chart),
            &mut budget,
        )?;
        if candidate_target.chart_identifier != self.target.chart_identifier {
            return Err(Error::Verification);
        }
        for kind in [PropertyKind::Font, PropertyKind::Size] {
            let index = property_index(kind);
            if !touched[index] {
                continue;
            }
            let expected = match kind {
                PropertyKind::Font => {
                    Property::Font(after_font.clone().ok_or(Error::Verification)?)
                },
                PropertyKind::Size => Property::Size(after_size.ok_or(Error::Verification)?),
            };
            if read_property(&published.package, &candidate_target, kind, &mut budget)? != expected
            {
                return Err(Error::Verification);
            }
        }
        let source_owner = structure::source_owner(self.source)?;
        let target_owner = structure::source_owner(&published.package)?;
        Ok(Commit {
            package: published.package,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
                sheet: self.target.sheet,
                chart: self.target.chart,
                chart_identifier: self.target.chart_identifier,
                touched,
                before_font,
                after_font,
                before_size,
                after_size,
            },
            diagnostics: Diagnostics { changed: true },
        })
    }
}

/// An exact-source reversible chart-font patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    sheet: Position,
    chart: Position,
    chart_identifier: u64,
    touched: [bool; 2],
    before_font: Option<ChartFont>,
    after_font: Option<ChartFont>,
    before_size: Option<TextPointSize>,
    after_size: Option<TextPointSize>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("sheet_position", &self.sheet)
            .field("chart_position", &self.chart)
            .field("chart_identifier", &self.chart_identifier)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether this patch leaves both semantic values and exact bytes unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
            && (!self.touched[0] || self.before_font == self.after_font)
            && (!self.touched[1] || self.before_size == self.after_size)
    }

    /// The value observed before this patch for `kind`.
    #[must_use]
    pub fn before(&self, kind: PropertyKind) -> Option<Property> {
        match kind {
            PropertyKind::Font => self.before_font.clone().map(Property::Font),
            PropertyKind::Size => self.before_size.map(Property::Size),
        }
    }

    /// The value produced by this patch for `kind`.
    #[must_use]
    pub fn after(&self, kind: PropertyKind) -> Option<Property> {
        match kind {
            PropertyKind::Font => self.after_font.clone().map(Property::Font),
            PropertyKind::Size => self.after_size.map(Property::Size),
        }
    }

    /// Return the exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            chart: self.chart,
            chart_identifier: self.chart_identifier,
            touched: self.touched,
            before_font: self.after_font.clone(),
            after_font: self.before_font.clone(),
            before_size: self.after_size,
            after_size: self.before_size,
        }
    }
}

/// Compact publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A verified immutable chart-font publication.
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

    /// Consume the commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return package plus patch.
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

/// Read one selected chart-font value.
pub(super) fn read<'sheet>(
    package: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    chart: impl Into<ChartSelector>,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let mut budget = ChartBudget::for_package(package).map_err(map_arrangement)?;
    let target = resolve_with_budget(package, sheet, chart, &mut budget)?;
    read_property(package, &target, kind, &mut budget)
}

/// Begin a selector-first exact-source chart-font edit.
pub(super) fn edit<'a, 'sheet>(
    package: &'a Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    chart: impl Into<ChartSelector>,
) -> Result<Edit<'a>, Error> {
    let mut budget = ChartBudget::for_package(package).map_err(map_arrangement)?;
    Ok(Edit {
        target: resolve_with_budget(package, sheet, chart, &mut budget)?,
        source: package,
        changes: Vec::new(),
        budget,
    })
}

/// Apply one exact-source chart-font patch.
pub(super) fn apply_patch(package: &Package, patch: &Patch) -> Result<Commit, Error> {
    let mut budget = ChartBudget::for_package(package).map_err(map_arrangement)?;
    let source_owner = structure::source_owner(package)?;
    if !patch.artifacts.authorizes_owner(&source_owner) {
        return Err(Error::PatchConflict);
    }
    let current = resolve_with_budget(
        package,
        SheetSelector::position(patch.sheet),
        ChartSelector::position(patch.chart),
        &mut budget,
    )?;
    if current.chart_identifier != patch.chart_identifier {
        return Err(Error::PatchConflict);
    }
    for kind in [PropertyKind::Font, PropertyKind::Size] {
        let index = property_index(kind);
        if !patch.touched[index] {
            continue;
        }
        let expected = match kind {
            PropertyKind::Font => patch.before_font.clone().map(Property::Font),
            PropertyKind::Size => patch.before_size.map(Property::Size),
        }
        .ok_or(Error::PatchConflict)?;
        if read_property(package, &current, kind, &mut budget)? != expected {
            return Err(Error::PatchConflict);
        }
    }
    if patch.is_noop() {
        return Ok(Commit {
            package: package.snapshot(),
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: false },
        });
    }
    let candidate_bytes = patch.artifacts.target_owner().as_ref().len();
    budget.output(candidate_bytes).map_err(map_arrangement)?;
    budget.retained(candidate_bytes).map_err(map_arrangement)?;
    budget.allocations(1).map_err(map_arrangement)?;
    budget.scratch(candidate_bytes).map_err(map_arrangement)?;
    let candidate = Package::from_source_owner_with_options(
        patch.artifacts.target_owner(),
        package.state.options,
    )
    .map_err(|_| Error::Verification)?;
    let verified = resolve_with_budget(
        &candidate,
        SheetSelector::position(patch.sheet),
        ChartSelector::position(patch.chart),
        &mut budget,
    )?;
    if verified.chart_identifier != patch.chart_identifier {
        return Err(Error::Verification);
    }
    for kind in [PropertyKind::Font, PropertyKind::Size] {
        let index = property_index(kind);
        if !patch.touched[index] {
            continue;
        }
        let expected = match kind {
            PropertyKind::Font => patch.after_font.clone().map(Property::Font),
            PropertyKind::Size => patch.after_size.map(Property::Size),
        }
        .ok_or(Error::Verification)?;
        if read_property(&candidate, &verified, kind, &mut budget)? != expected {
            return Err(Error::Verification);
        }
    }
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics { changed: true },
    })
}

impl Package {
    /// Read one effective font or point-size value of a selected sheet chart.
    pub fn sheet_chart_font<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        read(self, sheet, chart, kind)
    }

    /// Start an immutable exact-source edit of one selected sheet-chart font.
    pub fn edit_sheet_chart_font<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        edit(self, sheet, chart)
    }

    /// Apply an exact-source chart-font patch to this package.
    pub fn apply_sheet_chart_font(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
