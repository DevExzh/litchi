//! The style-variation graph engine of per-series chart overrides (record
//! 0375a): the sparse `TSCH.ChartSeriesNonStyleArchive` objects a chart
//! references from `series_non_styles`, one per series that overrides a
//! behavioral property (connection lines, wedge explosions, label
//! distances, …).
//!
//! [`set_values`] writes one property for every series, as the migration
//! host's editors do byte for byte: an override object is patched in place
//! when only this chart owns it, copied when another chart shares it,
//! allocated when the series had none, and removed (with its stylesheet
//! membership, UUID binding and object-container slot) when it becomes empty.
//! New styled overrides join the chart's document stylesheet; new unstyled
//! (geometry-only) overrides live in a fresh `TSP.ObjectContainer`
//! component, as iWork saves them. The chart's sparse array, the package
//! metadata registry and the issued-identifier watermark follow.
//!
//! A [`Property`] maps one value type to the override payload; the
//! properties of the Format panel are in [`property`].

use std::collections::{BTreeMap, HashSet};
use std::fmt;
use std::mem::size_of;

use litchi_iwa_common::{WireLimits, wire};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage};
use litchi_iwa_protos::chart_style_codec::{self as codec, Limits as CodecLimits};

use crate::graph::{
    self, DOCUMENT_METADATA_MEMBER, Error, Location, METADATA_MEMBER, Result, Workspace,
    push_message_field, push_packed_field, push_varint_field,
};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const SERIES_STYLE_MESSAGE_TYPE: u32 = 5_028;
const SERIES_NON_STYLE_MESSAGE_TYPE: u32 = 5_029;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const OBJECT_CONTAINER_MESSAGE_TYPE: u32 = 11_008;
const CHART_STYLE_MESSAGE_TYPES: [u32; 8] =
    [5_022, 5_023, 5_024, 5_025, 5_026, 5_027, 5_028, 5_029];

const STYLE_PARENT_FIELD: u32 = 1;
const STYLE_PARENT_REFERENCE_FIELD: u32 = 3;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const CONTAINER_IDENTIFIER_FIELD: u32 = 1;
const CONTAINER_OBJECTS_FIELD: u32 = 2;
const PRIMARY_PACKAGE_IDENTIFIER: u64 = 1;
const OBJECT_CONTAINER_LOCATOR: &str = "ObjectContainer";
const OBJECT_CONTAINER_VERSION: [u32; 3] = [2, 0, 0];
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const CAPABILITY_FIELDS: [u32; 3] = [10_001, 10_002, 10_003];

/// The style parent of an override object allocated for a series that had
/// none.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Base {
    /// Joins the chart's document stylesheet (labels, formats, lines).
    Styled,
    /// An empty style parent in a standalone object container (pure
    /// geometry: wedge explosions and label distances).
    Unstyled,
}

/// The finite resources admitted by one data-symbol operation.
///
/// Format owners derive this profile from their residual transaction budget.
/// The bounded symbol entry points never substitute a process-wide codec or
/// archive default after a caller supplies a narrower profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: usize,
    max_fields: usize,
    max_output_bytes: usize,
    max_nesting: u32,
    max_work_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
    max_references: usize,
    max_registry_entries: usize,
    max_identifiers: usize,
}

impl Limits {
    /// Build an explicit finite symbol-operation profile.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_fields: usize,
        max_output_bytes: usize,
        max_nesting: u32,
        max_work_bytes: usize,
        max_allocations: usize,
        max_retained_bytes: usize,
        max_scratch_bytes: usize,
        max_references: usize,
        max_registry_entries: usize,
        max_identifiers: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_fields,
            max_output_bytes,
            max_nesting,
            max_work_bytes,
            max_allocations,
            max_retained_bytes,
            max_scratch_bytes,
            max_references,
            max_registry_entries,
            max_identifiers,
        }
    }

    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    #[must_use]
    pub const fn max_nesting(self) -> u32 {
        self.max_nesting
    }

    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    #[must_use]
    pub const fn max_references(self) -> usize {
        self.max_references
    }

    #[must_use]
    pub const fn max_registry_entries(self) -> usize {
        self.max_registry_entries
    }

    #[must_use]
    pub const fn max_identifiers(self) -> usize {
        self.max_identifiers
    }
}

/// Resource dimension reported by a bounded symbol operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Source payload bytes inspected.
    InputBytes,
    /// Wire fields visited.
    Fields,
    /// Candidate bytes emitted.
    OutputBytes,
    /// Parent-chain depth.
    Nesting,
    /// Aggregate graph and wire work.
    Work,
    /// Fallible collection allocations.
    Allocations,
    /// Source, candidate and owned semantic buffers retained.
    RetainedBytes,
    /// Temporary buffers and planning scratch.
    ScratchBytes,
    /// Graph references visited or emitted.
    References,
    /// Registry entries rewritten.
    RegistryEntries,
    /// Object identifiers allocated or released.
    Identifiers,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Fields => "wire fields",
            Self::OutputBytes => "output bytes",
            Self::Nesting => "wire nesting",
            Self::Work => "work",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::References => "references",
            Self::RegistryEntries => "registry entries",
            Self::Identifiers => "identifiers",
        })
    }
}

/// A typed failure from a bounded symbol operation.
#[derive(Debug)]
#[non_exhaustive]
pub enum SymbolError {
    /// The graph or one of its generated codecs rejected the source.
    Graph(Error),
    /// A finite operation resource was exhausted.
    Limit {
        /// Resource dimension.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// A fallible collection allocation was refused.
    Allocation {
        /// Requested allocation units.
        amount: usize,
    },
    /// The symbol operation found an invalid rooted graph invariant.
    Invalid(&'static str),
    /// The package-wide identifier space is exhausted.
    IdentifierOverflow,
    /// A well-formed request the shared chart engine refuses.
    Refused(Refusal),
}

/// A well-formed chart-property request the shared engine refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Refusal {
    /// The property does not apply to the chart's kind.
    UnsupportedKind,
    /// The style the write would change is shared with another chart (or
    /// with another series of the same chart).
    SharedStyle,
    /// An image fill needs the embedded-data graph the property does not
    /// write.
    ImageFill,
    /// The chart has no style object of the property's archive (or no
    /// primary axis of the requested direction).
    MissingStyle,
    /// The property does not exist on the requested axis.
    UnsupportedAxis,
}

impl fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::UnsupportedKind => "the property does not apply to this chart kind",
            Self::SharedStyle => "the style is shared with another chart",
            Self::ImageFill => "image fills are not written by this property",
            Self::MissingStyle => "the chart has no style object for the property",
            Self::UnsupportedAxis => "the property does not exist on this axis",
        })
    }
}

impl fmt::Display for SymbolError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Graph(error) => write!(formatter, "series graph: {error}"),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "series {kind} limit exceeded: {observed} > {maximum}"
            ),
            Self::Allocation { amount } => {
                write!(formatter, "could not allocate {amount} series-symbol units")
            },
            Self::Invalid(reason) => formatter.write_str(reason),
            Self::IdentifierOverflow => formatter.write_str("series symbol identifier overflow"),
            Self::Refused(refusal) => refusal.fmt(formatter),
        }
    }
}

impl std::error::Error for SymbolError {}

impl From<Error> for SymbolError {
    fn from(error: Error) -> Self {
        Self::Graph(error)
    }
}

/// Result of a bounded chart-series or chart-style engine operation.
pub type SymbolResult<T> = std::result::Result<T, SymbolError>;

/// Source-derived mutation scope for one symbol operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MutationScope {
    selected_chart: u64,
    allowed_objects: Vec<u64>,
    allowed_members: Vec<String>,
    allowed_message_types: Vec<u32>,
    allowed_fields: Vec<u32>,
    planned_new_objects: Vec<u64>,
    planned_removed_objects: Vec<u64>,
    planned_new_members: Vec<String>,
    planned_removed_members: Vec<String>,
    chart_fields: Vec<u32>,
}

/// The chart-drawable extension fields a series write may change (the
/// sparse series override and private-style arrays).
const SERIES_CHART_FIELDS: [u32; 2] = [17, 19];

impl MutationScope {
    /// A scope of in-place payload rewrites: `objects` (in `members`) may
    /// change messages of `message_types`; the fields start with `fields`
    /// and grow with every generated-extension field a bounded patch
    /// changes. Nothing is allocated or removed.
    pub(crate) fn in_place(
        selected_chart: u64,
        objects: Vec<u64>,
        members: Vec<String>,
        message_types: Vec<u32>,
        fields: Vec<u32>,
        chart_fields: Vec<u32>,
    ) -> Self {
        Self {
            selected_chart,
            allowed_objects: objects,
            allowed_members: members,
            allowed_message_types: message_types,
            allowed_fields: fields,
            planned_new_objects: Vec::new(),
            planned_removed_objects: Vec::new(),
            planned_new_members: Vec::new(),
            planned_removed_members: Vec::new(),
            chart_fields,
        }
    }

    /// Admit `object` (in `member`, of `message_type`) to the scope.
    pub(crate) fn allow(&mut self, object: u64, member: &str, message_type: u32) {
        if !self.allowed_objects.contains(&object) {
            self.allowed_objects.push(object);
        }
        if !self.allowed_members.iter().any(|known| known == member) {
            self.allowed_members.push(member.to_owned());
        }
        if !self.allowed_message_types.contains(&message_type) {
            self.allowed_message_types.push(message_type);
        }
    }

    /// Plan `object` (in `member`, of `message_type`) as created (`true`)
    /// or removed (`false`); removing an object this scope planned to
    /// create cancels the creation.
    pub(crate) fn plan(&mut self, object: u64, member: &str, message_type: u32, created: bool) {
        if !self.allowed_members.iter().any(|known| known == member) {
            self.allowed_members.push(member.to_owned());
        }
        if !self.allowed_message_types.contains(&message_type) {
            self.allowed_message_types.push(message_type);
        }
        if created {
            self.planned_new_objects.push(object);
        } else if let Some(index) = self
            .planned_new_objects
            .iter()
            .position(|planned| *planned == object)
        {
            self.planned_new_objects.remove(index);
        } else {
            self.planned_removed_objects.push(object);
        }
    }

    /// Admit chart drawable extension `fields`.
    pub(crate) fn allow_chart_fields(&mut self, fields: &[u32]) {
        for field in fields {
            if !self.chart_fields.contains(field) {
                self.chart_fields.push(*field);
            }
        }
    }

    /// Admit extension or outer `fields`.
    pub(crate) fn allow_fields(&mut self, fields: &[u32]) {
        for field in fields {
            if !self.allowed_fields.contains(field) {
                self.allowed_fields.push(*field);
            }
        }
    }

    /// The fields of the chart drawable's generated extension the operation
    /// may change.
    #[must_use]
    pub fn allowed_chart_fields(&self) -> &[u32] {
        &self.chart_fields
    }

    #[must_use]
    pub const fn selected_chart(&self) -> u64 {
        self.selected_chart
    }

    #[must_use]
    pub fn allowed_objects(&self) -> &[u64] {
        &self.allowed_objects
    }

    #[must_use]
    pub fn allowed_members(&self) -> &[String] {
        &self.allowed_members
    }

    #[must_use]
    pub fn allowed_message_types(&self) -> &[u32] {
        &self.allowed_message_types
    }

    #[must_use]
    pub fn allowed_fields(&self) -> &[u32] {
        &self.allowed_fields
    }

    #[must_use]
    pub fn planned_new_objects(&self) -> &[u64] {
        &self.planned_new_objects
    }

    #[must_use]
    pub fn planned_removed_objects(&self) -> &[u64] {
        &self.planned_removed_objects
    }

    #[must_use]
    pub fn planned_new_members(&self) -> &[String] {
        &self.planned_new_members
    }

    #[must_use]
    pub fn planned_removed_members(&self) -> &[String] {
        &self.planned_removed_members
    }
}

/// Actual accounting from one bounded symbol operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Report {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    max_depth: u32,
    work_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    references: usize,
    registry_entries: usize,
    identifiers: usize,
    changed: bool,
    scope: Option<MutationScope>,
}

impl Report {
    #[must_use]
    pub const fn input_bytes(&self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub const fn output_bytes(&self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn fields(&self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn max_depth(&self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub const fn work_bytes(&self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn allocations(&self) -> usize {
        self.allocations
    }

    #[must_use]
    pub const fn retained_bytes(&self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(&self) -> usize {
        self.scratch_bytes
    }

    #[must_use]
    pub const fn references(&self) -> usize {
        self.references
    }

    #[must_use]
    pub const fn registry_entries(&self) -> usize {
        self.registry_entries
    }

    #[must_use]
    pub const fn identifiers(&self) -> usize {
        self.identifiers
    }

    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    #[must_use]
    pub fn scope(&self) -> Option<&MutationScope> {
        self.scope.as_ref()
    }
}

#[derive(Debug, Clone)]
pub(crate) struct SymbolBudget {
    limits: Limits,
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    max_depth: u32,
    work_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    references: usize,
    registry_entries: usize,
    identifiers: usize,
    scope: Option<MutationScope>,
}

impl SymbolBudget {
    pub(crate) const fn new(limits: Limits) -> Self {
        Self {
            limits,
            input_bytes: 0,
            output_bytes: 0,
            fields: 0,
            max_depth: 0,
            work_bytes: 0,
            allocations: 0,
            retained_bytes: 0,
            scratch_bytes: 0,
            references: 0,
            registry_entries: 0,
            identifiers: 0,
            scope: None,
        }
    }

    pub(crate) fn limit(kind: LimitKind, observed: usize, maximum: usize) -> SymbolError {
        SymbolError::Limit {
            kind,
            observed,
            maximum,
        }
    }

    pub(crate) fn add(
        kind: LimitKind,
        current: &mut usize,
        amount: usize,
        maximum: usize,
    ) -> SymbolResult<()> {
        let observed = current
            .checked_add(amount)
            .ok_or_else(|| Self::limit(kind, usize::MAX, maximum))?;
        if observed > maximum {
            return Err(Self::limit(kind, observed, maximum));
        }
        *current = observed;
        Ok(())
    }

    pub(crate) fn preflight(
        &self,
        kind: LimitKind,
        current: usize,
        amount: usize,
        maximum: usize,
    ) -> SymbolResult<()> {
        let observed = current
            .checked_add(amount)
            .ok_or_else(|| Self::limit(kind, usize::MAX, maximum))?;
        if observed > maximum {
            return Err(Self::limit(kind, observed, maximum));
        }
        Ok(())
    }

    pub(crate) fn preflight_output(&self, amount: usize) -> SymbolResult<()> {
        self.preflight(
            LimitKind::OutputBytes,
            self.output_bytes,
            amount,
            self.limits.max_output_bytes,
        )
    }

    pub(crate) fn preflight_retained(&self, amount: usize) -> SymbolResult<()> {
        self.preflight(
            LimitKind::RetainedBytes,
            self.retained_bytes,
            amount,
            self.limits.max_retained_bytes,
        )
    }

    pub(crate) fn preflight_scratch(&self, amount: usize) -> SymbolResult<()> {
        self.preflight(
            LimitKind::ScratchBytes,
            self.scratch_bytes,
            amount,
            self.limits.max_scratch_bytes,
        )
    }

    pub(crate) fn preflight_allocations(&self, amount: usize) -> SymbolResult<()> {
        self.preflight(
            LimitKind::Allocations,
            self.allocations,
            amount,
            self.limits.max_allocations,
        )
    }

    pub(crate) fn source(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::InputBytes,
            &mut self.input_bytes,
            amount,
            self.limits.max_input_bytes,
        )
    }

    pub(crate) fn output(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::OutputBytes,
            &mut self.output_bytes,
            amount,
            self.limits.max_output_bytes,
        )
    }

    pub(crate) fn fields(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::Fields,
            &mut self.fields,
            amount,
            self.limits.max_fields,
        )
    }

    pub(crate) fn work(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::Work,
            &mut self.work_bytes,
            amount,
            self.limits.max_work_bytes,
        )
    }

    pub(crate) fn allocations(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::Allocations,
            &mut self.allocations,
            amount,
            self.limits.max_allocations,
        )
    }

    pub(crate) fn retained(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::RetainedBytes,
            &mut self.retained_bytes,
            amount,
            self.limits.max_retained_bytes,
        )
    }

    pub(crate) fn scratch(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::ScratchBytes,
            &mut self.scratch_bytes,
            amount,
            self.limits.max_scratch_bytes,
        )
    }

    pub(crate) fn references(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::References,
            &mut self.references,
            amount,
            self.limits.max_references,
        )
    }

    pub(crate) fn registry_entries(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::RegistryEntries,
            &mut self.registry_entries,
            amount,
            self.limits.max_registry_entries,
        )
    }

    pub(crate) fn identifiers(&mut self, amount: usize) -> SymbolResult<()> {
        Self::add(
            LimitKind::Identifiers,
            &mut self.identifiers,
            amount,
            self.limits.max_identifiers,
        )
    }

    pub(crate) fn depth(&mut self, depth: u32) -> SymbolResult<()> {
        if depth > self.limits.max_nesting {
            return Err(Self::limit(
                LimitKind::Nesting,
                depth as usize,
                self.limits.max_nesting as usize,
            ));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    pub(crate) fn remaining(kind: LimitKind, maximum: usize, used: usize) -> SymbolResult<usize> {
        let remaining = maximum.saturating_sub(used);
        if remaining == 0 {
            return Err(Self::limit(kind, used.saturating_add(1), maximum));
        }
        Ok(remaining)
    }

    pub(crate) fn wire_limits(&self) -> SymbolResult<WireLimits> {
        if self.limits.max_nesting == 0 {
            return Err(Self::limit(LimitKind::Nesting, 1, 0));
        }
        let input = Self::remaining(
            LimitKind::InputBytes,
            self.limits.max_input_bytes,
            self.input_bytes,
        )?
        .min(WireLimits::MAX_INPUT_BYTES);
        let fields = Self::remaining(LimitKind::Fields, self.limits.max_fields, self.fields)?
            .min(WireLimits::MAX_FIELDS);
        let work = Self::remaining(LimitKind::Work, self.limits.max_work_bytes, self.work_bytes)?
            .min(WireLimits::MAX_REWRITE_WORK);
        WireLimits::DEFAULT
            .with_input_bytes(input)
            .and_then(|limits| limits.with_fields(fields))
            .and_then(|limits| limits.with_rewrite_work(work))
            .and_then(|limits| {
                limits.with_nesting((self.limits.max_nesting as usize).min(WireLimits::MAX_NESTING))
            })
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))
    }

    pub(crate) fn codec_limits(&self) -> SymbolResult<CodecLimits> {
        let bytes = Self::remaining(
            LimitKind::InputBytes,
            self.limits.max_input_bytes,
            self.input_bytes,
        )?
        .min(Self::remaining(
            LimitKind::OutputBytes,
            self.limits.max_output_bytes,
            self.output_bytes,
        )?)
        .min(CodecLimits::DEFAULT.max_bytes());
        let fields = Self::remaining(LimitKind::Fields, self.limits.max_fields, self.fields)?
            .min(CodecLimits::DEFAULT.max_fields());
        CodecLimits::DEFAULT
            .with_max_bytes(bytes)
            .and_then(|limits| limits.with_max_fields(fields))
            .map_err(|error| SymbolError::Graph(Error::Codec(error)))
    }

    pub(crate) fn reserve<T>(&mut self, count: usize, scratch: bool) -> SymbolResult<()> {
        self.allocations(1)?;
        let bytes = count.checked_mul(size_of::<T>()).ok_or_else(|| {
            Self::limit(
                LimitKind::RetainedBytes,
                usize::MAX,
                self.limits.max_retained_bytes,
            )
        })?;
        self.retained(bytes)?;
        if scratch {
            self.scratch(bytes)?;
        }
        Ok(())
    }

    pub(crate) fn reserve_vec<T>(
        &mut self,
        values: &mut Vec<T>,
        count: usize,
        scratch: bool,
    ) -> SymbolResult<()> {
        self.reserve::<T>(count, scratch)?;
        values
            .try_reserve_exact(count)
            .map_err(|_| SymbolError::Allocation { amount: count })?;
        Ok(())
    }

    pub(crate) fn admit_registry_inspection(
        &mut self,
        workspace: &Workspace<'_>,
    ) -> SymbolResult<()> {
        let Some(report) = workspace
            .registry_inspection_requirements()
            .map_err(SymbolError::from)?
        else {
            return Ok(());
        };
        self.source(report.input_bytes())?;
        self.output(report.output_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        self.references(report.references_scanned())?;
        self.registry_entries(report.components_scanned())?;
        self.preflight_allocations(1)?;
        self.preflight_retained(report.references_scanned().saturating_mul(size_of::<u64>()))?;
        self.preflight_scratch(report.references_scanned().saturating_mul(size_of::<u64>()))?;
        self.allocations(1)?;
        self.retained(report.references_scanned().saturating_mul(size_of::<u64>()))?;
        self.scratch(report.references_scanned().saturating_mul(size_of::<u64>()))?;
        Ok(())
    }

    pub(crate) fn preflight_codec_report(
        &self,
        report: litchi_iwa_protos::chart_style_codec::Report,
    ) -> SymbolResult<()> {
        self.preflight_allocations(report.allocations())?;
        self.preflight_retained(report.input_bytes())?;
        self.preflight_scratch(report.input_bytes())
    }

    pub(crate) fn charge_codec_report(
        &mut self,
        report: litchi_iwa_protos::chart_style_codec::Report,
    ) -> SymbolResult<()> {
        self.source(report.input_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.allocations(report.allocations())?;
        self.retained(report.input_bytes())?;
        self.scratch(report.input_bytes())?;
        Ok(())
    }

    pub(crate) fn parse_fields(
        &mut self,
        source: &[u8],
        depth: u32,
    ) -> SymbolResult<Vec<wire::WireField>> {
        self.depth(depth)?;
        let limits = self.wire_limits()?;
        self.allocations(1)?;
        self.scratch(size_of::<wire::WireField>())?;
        let mut count = 0usize;
        wire::visit_wire_fields(source, limits, |_field| {
            count = count.saturating_add(1);
            Ok(())
        })
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        self.source(source.len())?;
        self.fields(count)?;
        self.work(source.len().saturating_add(count))?;
        let storage = count
            .checked_mul(size_of::<wire::WireField>())
            .ok_or_else(|| {
                Self::limit(
                    LimitKind::ScratchBytes,
                    usize::MAX,
                    self.limits.max_scratch_bytes,
                )
            })?;
        self.scratch(storage)?;
        self.allocations(1)?;
        wire::parse_wire_fields(source, limits)
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))
    }

    pub(crate) fn report(self, changed: bool) -> Report {
        Report {
            input_bytes: self.input_bytes,
            output_bytes: self.output_bytes,
            fields: self.fields,
            max_depth: self.max_depth,
            work_bytes: self.work_bytes,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
            references: self.references,
            registry_entries: self.registry_entries,
            identifiers: self.identifiers,
            changed,
            scope: self.scope,
        }
    }

    pub(crate) fn scope_mut(&mut self) -> Option<&mut MutationScope> {
        self.scope.as_mut()
    }

    pub(crate) fn set_scope(&mut self, scope: MutationScope) {
        self.scope = Some(scope);
    }

    /// Admit one more extension field into the mutation scope.
    pub(crate) fn allow_field(&mut self, field: u32) -> SymbolResult<()> {
        if self
            .scope
            .as_ref()
            .is_none_or(|scope| scope.allowed_fields.contains(&field))
        {
            return Ok(());
        }
        self.preflight_allocations(1)?;
        self.preflight_retained(size_of::<u32>())?;
        self.allocations(1)?;
        self.retained(size_of::<u32>())?;
        let scope = self.scope.as_mut().expect("scope exists");
        scope
            .allowed_fields
            .try_reserve(1)
            .map_err(|_| SymbolError::Allocation { amount: 1 })?;
        scope.allowed_fields.push(field);
        Ok(())
    }

    pub(crate) fn planned_new_object(&mut self, identifier: u64) {
        if let Some(scope) = self.scope.as_mut() {
            scope.planned_new_objects.push(identifier);
        }
    }

    pub(crate) fn planned_removed_object(&mut self, identifier: u64) {
        if let Some(scope) = self.scope.as_mut() {
            scope.planned_removed_objects.push(identifier);
        }
    }

    pub(crate) fn planned_member(&mut self, new_member: bool, member: &str) -> SymbolResult<()> {
        if self.scope.is_none() {
            return Ok(());
        }
        self.preflight_allocations(1)?;
        self.preflight_retained(member.len())?;
        self.allocations(1)?;
        self.retained(member.len())?;
        let scope = self.scope.as_mut().expect("scope exists");
        if new_member {
            scope.planned_new_members.push(member.to_owned());
        } else {
            scope.planned_removed_members.push(member.to_owned());
        }
        Ok(())
    }
}

/// One per-series override property.
pub trait Property {
    /// The value of one series.
    type Value: Clone + PartialEq;

    /// Where a newly allocated override lives.
    fn base(&self) -> Base;

    /// The value of a series without an override.
    fn default_value(&self) -> Self::Value;

    /// Read the value from an override payload.
    ///
    /// # Errors
    ///
    /// The payload is invalid or stores an unsupported value.
    fn read(&self, payload: &[u8]) -> Result<Self::Value>;

    /// Store `value` in an override payload, keeping every other byte.
    ///
    /// # Errors
    ///
    /// The payload is invalid or `value` cannot be stored.
    fn patch(&self, payload: &[u8], value: &Self::Value) -> Result<Vec<u8>>;
}

#[derive(Debug, Clone)]
struct Slot {
    location: Location,
    object: u64,
    message: usize,
}

struct SeriesGraph {
    chart_message: usize,
    slots: Vec<Option<Slot>>,
}

/// Read `property` for each of the `series_count` series of chart drawable
/// `chart`, in series order.
///
/// # Errors
///
/// An invalid chart, sparse array or override, or a value the property
/// rejects.
pub fn values<P: Property>(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    property: &P,
) -> Result<Vec<P::Value>> {
    let graph = series_graph(workspace, chart, series_count)?;
    read_values(workspace, &graph, property)
}

fn read_values<P: Property>(
    workspace: &Workspace<'_>,
    graph: &SeriesGraph,
    property: &P,
) -> Result<Vec<P::Value>> {
    graph
        .slots
        .iter()
        .map(|slot| match slot {
            Some(slot) => property.read(slot_payload(workspace, slot)?),
            None => Ok(property.default_value()),
        })
        .collect()
}

fn slot_payload<'w>(workspace: &'w Workspace<'_>, slot: &Slot) -> Result<&'w [u8]> {
    let object = workspace.object(slot.object)?;
    let message = object
        .messages
        .get(slot.message)
        .ok_or(Error::Invalid("a series override payload index changed"))?;
    if message.type_ != SERIES_NON_STYLE_MESSAGE_TYPE {
        return Err(Error::Invalid("a series override payload type changed"));
    }
    Ok(message.data.as_slice())
}

fn chart_payload<'w>(workspace: &'w Workspace<'_>, chart: u64) -> Result<(usize, &'w [u8])> {
    workspace.unique_payload(
        chart,
        CHART_MESSAGE_TYPE,
        "a chart must have exactly one chart payload",
    )
}

fn series_graph(workspace: &Workspace<'_>, chart: u64, series_count: usize) -> Result<SeriesGraph> {
    let (chart_message, payload) = chart_payload(workspace, chart)?;
    let references = codec::references(payload, CodecLimits::DEFAULT)?;
    let mut slots = vec![None; series_count];
    if let Some(sparse) = references.series_non_styles() {
        if usize::try_from(sparse.count()).ok() != Some(sparse.entries().len()) {
            return Err(Error::Invalid(
                "a chart's series_non_styles count disagrees with its entries",
            ));
        }
        let mut seen = HashSet::new();
        for &(index, object) in sparse.entries() {
            let index = usize::try_from(index)
                .map_err(|_| Error::Invalid("a series index exceeds usize"))?;
            if index >= series_count {
                return Err(Error::Invalid(
                    "a series override index exceeds the series count",
                ));
            }
            if !seen.insert(object) || slots[index].is_some() {
                return Err(Error::Invalid("a series override is repeated"));
            }
            let location = workspace.locate(object)?;
            let (message, _) = workspace.unique_payload(
                object,
                SERIES_NON_STYLE_MESSAGE_TYPE,
                "a series override must have exactly one payload",
            )?;
            slots[index] = Some(Slot {
                location,
                object,
                message,
            });
        }
    }
    Ok(SeriesGraph {
        chart_message,
        slots,
    })
}

/// The document stylesheet of the chart's first series theme style.
fn chart_stylesheet(workspace: &Workspace<'_>, chart: u64) -> Result<u64> {
    let (_, payload) = chart_payload(workspace, chart)?;
    let style = codec::references(payload, CodecLimits::DEFAULT)?
        .series_theme_styles()
        .first()
        .copied()
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::Invalid("the chart has no series theme style"))?;
    let (_, style_payload) = workspace.unique_payload(
        style,
        SERIES_STYLE_MESSAGE_TYPE,
        "a series style must have exactly one payload",
    )?;
    stylesheet_of(style_payload)?
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::Invalid(
            "the series theme style has no document stylesheet",
        ))
}

/// The stylesheet reference of a style payload's `TSS.StyleArchive` parent
/// (`Some(0)` for a reference without identifier).
fn stylesheet_of(payload: &[u8]) -> Result<Option<u64>> {
    let parent = graph::repeated(payload, STYLE_PARENT_FIELD)?;
    let [parent] = parent.as_slice() else {
        return Err(Error::Invalid("a chart style has no single style parent"));
    };
    let stylesheet = graph::repeated(parent, STYLE_STYLESHEET_FIELD)?;
    Ok(match stylesheet.last() {
        Some(reference) => Some(graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?.unwrap_or(0)),
        None => None,
    })
}

fn style_parent(stylesheet: Option<u64>) -> Vec<u8> {
    let mut style = Vec::new();
    if let Some(identifier) = stylesheet {
        let mut reference = Vec::with_capacity(11);
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, identifier);
        push_message_field(&mut style, STYLE_STYLESHEET_FIELD, &reference);
    }
    style
}

fn override_with_parent(stylesheet: Option<u64>) -> Vec<u8> {
    let mut payload = Vec::new();
    push_message_field(&mut payload, STYLE_PARENT_FIELD, &style_parent(stylesheet));
    payload
}

fn append_capabilities(payload: &mut Vec<u8>) {
    for field in CAPABILITY_FIELDS {
        push_varint_field(payload, field, 1);
    }
}

fn canonical_empty(stylesheet: Option<u64>) -> Vec<u8> {
    let mut payload = override_with_parent(stylesheet);
    append_capabilities(&mut payload);
    payload
}

fn set_stylesheet(payload: &[u8], stylesheet: u64) -> Result<Vec<u8>> {
    if stylesheet_of(payload)? == Some(stylesheet) {
        return Ok(payload.to_vec());
    }
    Ok(wire::patch_length_delimited_field(
        payload,
        STYLE_PARENT_FIELD,
        true,
        Some(&style_parent(Some(stylesheet))),
    )?)
}

fn new_override(
    identifier: u64,
    data: Vec<u8>,
    workspace: &Workspace<'_>,
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SERIES_NON_STYLE_MESSAGE_TYPE,
            data,
        }],
        workspace.limits(),
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec().into();
    Ok(object)
}

fn allocate(next: &mut u64) -> Result<u64> {
    let identifier = *next;
    *next = next.checked_add(1).ok_or(Error::IdentifierOverflow)?;
    Ok(identifier)
}

/// How many chart payloads reference override `object`.
fn owner_count(workspace: &Workspace<'_>, object: u64) -> Result<usize> {
    let mut owners = 0usize;
    for location in workspace.locations_iter() {
        for candidate in &workspace.archive_at(location)?.objects {
            for message in &candidate.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let references = codec::references(&message.data, CodecLimits::DEFAULT)?;
                if references
                    .series_non_styles()
                    .is_some_and(|sparse| sparse.entries().iter().any(|(_, id)| *id == object))
                {
                    owners += 1;
                }
            }
        }
    }
    Ok(owners)
}

struct Container {
    component: u64,
    name: String,
    locator: String,
}

fn reserve_container(workspace: &Workspace<'_>, next: &mut u64) -> Result<Container> {
    let component = allocate(next)?;
    let mut suffix: Option<u64> = None;
    loop {
        let locator = suffix.map_or_else(
            || OBJECT_CONTAINER_LOCATOR.to_owned(),
            |value| format!("{OBJECT_CONTAINER_LOCATOR}-{value}"),
        );
        let name = format!("Index/{locator}.iwa");
        if !workspace.has_member(&name) && workspace.component_for_member(&name)?.is_none() {
            return Ok(Container {
                component,
                name,
                locator,
            });
        }
        suffix = Some(match suffix {
            None => component,
            Some(value) => value.checked_add(1).ok_or(Error::IdentifierOverflow)?,
        });
    }
}

fn encode_container(objects: &[u64]) -> Vec<u8> {
    let mut payload = Vec::new();
    push_varint_field(
        &mut payload,
        CONTAINER_IDENTIFIER_FIELD,
        PRIMARY_PACKAGE_IDENTIFIER,
    );
    for &object in objects {
        let mut reference = Vec::with_capacity(11);
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, object);
        push_message_field(&mut payload, CONTAINER_OBJECTS_FIELD, &reference);
    }
    payload
}

fn insert_container(
    workspace: &mut Workspace<'_>,
    chart_member: &str,
    container: Container,
    objects: Vec<ArchiveObject>,
) -> Result<()> {
    let identifiers = objects
        .iter()
        .map(|object| {
            object
                .archive_info
                .identifier
                .ok_or(Error::Invalid("a container member has no identifier"))
        })
        .collect::<Result<Vec<_>>>()?;
    let limits = workspace.limits();
    let mut archive_objects = Vec::with_capacity(objects.len() + 1);
    archive_objects.push(ArchiveObject::new(
        container.component,
        vec![RawMessage {
            type_: OBJECT_CONTAINER_MESSAGE_TYPE,
            data: encode_container(&identifiers),
        }],
        limits,
    )?);
    archive_objects.extend(objects);
    let archive = Archive {
        objects: archive_objects,
    };
    let anchor = if workspace.has_member(DOCUMENT_METADATA_MEMBER) {
        DOCUMENT_METADATA_MEMBER
    } else {
        METADATA_MEMBER
    };
    workspace.create_member(&container.name, archive, Some(anchor))?;
    if !workspace.has_registry() {
        return Ok(());
    }
    let mut registration = Vec::new();
    push_varint_field(&mut registration, 1, container.component);
    push_message_field(&mut registration, 2, container.locator.as_bytes());
    push_packed_field(&mut registration, 4, &OBJECT_CONTAINER_VERSION);
    push_packed_field(&mut registration, 5, &OBJECT_CONTAINER_VERSION);
    push_varint_field(&mut registration, 10, 0);
    if let Some(token) = workspace.save_token()? {
        push_varint_field(&mut registration, 12, token);
    }
    if workspace.has_data_metadata_map()? {
        push_varint_field(&mut registration, 21, PRIMARY_PACKAGE_IDENTIFIER);
    }
    workspace.add_component(container.component, &registration)?;
    workspace.add_object_uuids(container.component, &identifiers)?;
    let source = workspace
        .component_for_member(chart_member)?
        .ok_or(Error::Invalid("the chart component is not registered"))?;
    for identifier in identifiers {
        workspace.add_external_reference(source, container.component, identifier)?;
    }
    Ok(())
}

/// `(container object, payload index, member objects)` of an object
/// container component.
fn container_of(archive: &Archive) -> Result<Option<(u64, usize, Vec<u64>)>> {
    let mut found = None;
    for object in &archive.objects {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == OBJECT_CONTAINER_MESSAGE_TYPE {
                let identifier = object
                    .archive_info
                    .identifier
                    .ok_or(Error::Invalid("an object container has no identifier"))?;
                if found
                    .replace((identifier, index, message.data.as_slice()))
                    .is_some()
                {
                    return Err(Error::Invalid(
                        "a component holds several object containers",
                    ));
                }
            }
        }
    }
    let Some((identifier, index, payload)) = found else {
        return Ok(None);
    };
    if graph::varint(payload, CONTAINER_IDENTIFIER_FIELD)? != Some(PRIMARY_PACKAGE_IDENTIFIER) {
        return Err(Error::Invalid(
            "an object container is not in the primary package",
        ));
    }
    let mut members = Vec::new();
    let mut seen = HashSet::new();
    for reference in graph::repeated(payload, CONTAINER_OBJECTS_FIELD)? {
        let member = graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?.unwrap_or(0);
        if member == 0 || !seen.insert(member) || archive.object(member).is_none() {
            return Err(Error::Invalid("an object container member is invalid"));
        }
        members.push(member);
    }
    Ok(Some((identifier, index, members)))
}

/// Remove `objects` from their container; returns the container's
/// identifier when it became empty (the component is then deleted).
fn remove_from_container(
    workspace: &mut Workspace<'_>,
    location: Location,
    objects: &[u64],
) -> Result<Option<u64>> {
    let member = workspace.member(location).to_owned();
    let (container, message, mut members) = container_of(workspace.archive_at(location)?)?.ok_or(
        Error::Invalid("a series override container is not an object container"),
    )?;
    let requested = objects.iter().copied().collect::<HashSet<_>>();
    if requested.len() != objects.len() || !requested.iter().all(|id| members.contains(id)) {
        return Err(Error::Invalid(
            "an object container does not own every removed override",
        ));
    }
    members.retain(|identifier| !requested.contains(identifier));
    let component = workspace.component_for_member(&member)?;
    if component.is_some_and(|component| component != container) {
        return Err(Error::Invalid(
            "an object container disagrees with its component",
        ));
    }
    if members.is_empty() {
        if let Some(component) = component {
            let locator = workspace
                .component_locator(component)?
                .ok_or(Error::Invalid(
                    "an object container component has no locator",
                ))?;
            workspace.remove_component(component, &locator)?;
        }
        workspace.delete_member(location)?;
        return Ok(Some(container));
    }
    if let Some(component) = component {
        let registered = workspace.component_uuid_objects(component)?;
        let bound = objects
            .iter()
            .copied()
            .filter(|identifier| registered.contains(identifier))
            .collect::<Vec<_>>();
        workspace.remove_object_uuids(component, &bound)?;
        for &identifier in objects {
            workspace.remove_external_references_to(component, identifier)?;
        }
    }
    for &identifier in objects {
        workspace.remove(identifier)?;
    }
    workspace.replace_payload(container, message, encode_container(&members))?;
    Ok(None)
}

fn style_member(workspace: &Workspace<'_>, owner: Location, styles: &[u64]) -> Result<()> {
    let unique = styles.iter().copied().collect::<HashSet<_>>();
    if unique.len() != styles.len() || unique.contains(&0) {
        return Err(Error::Invalid(
            "chart style identifiers must be unique and nonzero",
        ));
    }
    let archive = workspace.archive_at(owner)?;
    for &style in styles {
        let object = archive.object(style).ok_or(Error::Invalid(
            "a chart style is not in its owner component",
        ))?;
        let payloads = object
            .messages
            .iter()
            .filter(|message| CHART_STYLE_MESSAGE_TYPES.contains(&message.type_))
            .count();
        if payloads != 1 {
            return Err(Error::Invalid(
                "a chart style must have exactly one style payload",
            ));
        }
    }
    Ok(())
}

fn register_styles(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    owner: Location,
    styles: &[u64],
) -> Result<()> {
    style_member(workspace, owner, styles)?;
    let sheet_location = workspace.locate(stylesheet)?;
    let owner_member = workspace.member(owner).to_owned();
    let sheet_member = workspace.member(sheet_location).to_owned();
    let owner_component = workspace.component_for_member(&owner_member)?;
    let sheet_component = workspace.component_for_member(&sheet_member)?;
    let (message, payload) = workspace.unique_payload(
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    let mut data = payload.to_vec();
    let registered = graph::repeated(&data, STYLESHEET_STYLES_FIELD)?
        .into_iter()
        .map(|reference| graph::varint(reference, REFERENCE_IDENTIFIER_FIELD))
        .collect::<Result<Vec<_>>>()?;
    for &style in styles {
        if registered.contains(&Some(style)) {
            return Err(Error::Invalid("the stylesheet already lists a new style"));
        }
        let mut reference = Vec::with_capacity(11);
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, style);
        data = wire::append_repeated_length_delimited_field(
            &data,
            STYLESHEET_STYLES_FIELD,
            &reference,
        )?;
    }
    workspace.replace_payload(stylesheet, message, data)?;
    let info = workspace
        .object_mut(stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    for &style in styles {
        if info.object_references.contains(&style) {
            return Err(Error::Invalid(
                "the stylesheet metadata already references a new style",
            ));
        }
        info.object_references.push(style);
    }
    if let (Some(owner), Some(sheet)) = (owner_component, sheet_component)
        && owner != sheet
    {
        workspace.add_external_reference(owner, sheet, stylesheet)?;
        for &style in styles {
            workspace.add_external_reference(sheet, owner, style)?;
        }
    }
    Ok(())
}

fn unregister_styles(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    owner: Location,
    styles: &[u64],
) -> Result<()> {
    let unique = styles.iter().copied().collect::<HashSet<_>>();
    if unique.len() != styles.len() || unique.contains(&0) {
        return Err(Error::Invalid(
            "chart style identifiers must be unique and nonzero",
        ));
    }
    let sheet_location = workspace.locate(stylesheet)?;
    let local = sheet_location == owner;
    let owner_member = workspace.member(owner).to_owned();
    let sheet_member = workspace.member(sheet_location).to_owned();
    let owner_component = workspace.component_for_member(&owner_member)?;
    let sheet_component = workspace.component_for_member(&sheet_member)?;
    let (message, payload) = workspace.unique_payload(
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    let mut kept = Vec::new();
    let mut counts = BTreeMap::<u64, usize>::new();
    for reference in graph::repeated(payload, STYLESHEET_STYLES_FIELD)? {
        let identifier = graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?.unwrap_or(0);
        if unique.contains(&identifier) {
            *counts.entry(identifier).or_default() += 1;
        } else {
            kept.push(reference.to_vec());
        }
    }
    if styles.iter().any(|style| counts.get(style) != Some(&1)) {
        return Err(Error::Invalid(
            "the stylesheet must list each removed style once",
        ));
    }
    let data =
        wire::rewrite_repeated_length_delimited_fields(payload, STYLESHEET_STYLES_FIELD, &kept)?;
    workspace.replace_payload(stylesheet, message, data)?;
    let info = workspace
        .object_mut(stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    for &style in styles {
        let listed = info
            .object_references
            .iter()
            .filter(|identifier| **identifier == style)
            .count();
        if listed > 1 || (!local && listed != 1) {
            return Err(Error::Invalid(
                "the stylesheet metadata has invalid style ownership",
            ));
        }
        info.object_references
            .retain(|identifier| *identifier != style);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| *identifier != style);
        }
    }
    if let (Some(owner), Some(sheet)) = (owner_component, sheet_component)
        && owner != sheet
    {
        for &style in styles {
            workspace.remove_external_reference(sheet, owner, style)?;
        }
    }
    Ok(())
}

fn patch_chart_references(
    workspace: &mut Workspace<'_>,
    chart: u64,
    message: usize,
    previous: &[u64],
    identifiers: &[Option<u64>],
) -> Result<()> {
    let entries = identifiers
        .iter()
        .enumerate()
        .filter_map(|(index, identifier)| identifier.map(|identifier| (index, identifier)))
        .map(|(index, identifier)| {
            u32::try_from(index)
                .map(|index| (index, identifier))
                .map_err(|_| Error::Invalid("a series index exceeds u32"))
        })
        .collect::<Result<Vec<_>>>()?;
    let payload = workspace.object(chart)?.messages[message].data.as_slice();
    let data = codec::rewrite_series_non_styles(payload, &entries, CodecLimits::DEFAULT)?
        .into_bytes()
        .into_owned();
    workspace.replace_payload(chart, message, data)?;
    let info = workspace
        .object_mut(chart)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a chart payload has no metadata"))?;
    info.object_references
        .retain(|identifier| !previous.contains(identifier));
    for identifier in identifiers.iter().copied().flatten() {
        if !info.object_references.contains(&identifier) {
            info.object_references.push(identifier);
        }
    }
    Ok(())
}

/// Write `property` for every series of chart drawable `chart`, in series
/// order (`expected.len()` is the series count). Returns `false` when every
/// series already has its value (nothing is staged).
///
/// # Errors
///
/// An empty `expected`; an invalid chart, sparse array, override,
/// stylesheet, object container or registry; a value the property rejects;
/// a write that does not read back.
#[allow(
    clippy::too_many_lines,
    reason = "one transaction in the migration host's order, kept in one place for review"
)]
pub fn set_values<P: Property>(
    workspace: &mut Workspace<'_>,
    chart: u64,
    property: &P,
    expected: &[P::Value],
) -> Result<bool> {
    if expected.is_empty() {
        return Err(Error::Invalid("a chart series write names no series"));
    }
    let base = property.base();
    let default = property.default_value();
    let graph = series_graph(workspace, chart, expected.len())?;
    let stylesheet = chart_stylesheet(workspace, chart)?;
    let current = read_values(workspace, &graph, property)?;
    if current == expected {
        return Ok(false);
    }
    let chart_location = workspace.locate(chart)?;
    let chart_member = workspace.member(chart_location).to_owned();

    let styled_empty = canonical_empty(Some(stylesheet));
    let unstyled_empty = canonical_empty(None);
    let unstyled_base = override_with_parent(None);
    let new_base = match base {
        Base::Styled => styled_empty.as_slice(),
        Base::Unstyled => unstyled_base.as_slice(),
    };
    let mut next = workspace.next_object_identifier()?;
    let mut final_ids = graph
        .slots
        .iter()
        .map(|slot| slot.as_ref().map(|slot| slot.object))
        .collect::<Vec<_>>();
    let mut updates = Vec::new();
    let mut removals = Vec::new();
    let mut styled_creations = Vec::new();
    let mut chart_created = Vec::new();
    let mut unstyled_creations = Vec::new();
    let mut container = None;
    let mut created = Vec::new();
    let mut registrations = BTreeMap::<Location, Vec<u64>>::new();
    let mut unregistrations = BTreeMap::<Location, Vec<u64>>::new();

    for (index, ((slot, value), replacement)) in
        graph.slots.iter().zip(&current).zip(expected).enumerate()
    {
        if value == replacement {
            continue;
        }
        if let Some(slot) = slot {
            let payload = slot_payload(workspace, slot)?;
            let mut patched = property.patch(payload, replacement)?;
            let registered = stylesheet_of(payload)?;
            if let Some(registered) = registered
                && registered != stylesheet
            {
                return Err(Error::Invalid(
                    "a series override belongs to another stylesheet",
                ));
            }
            let needs_registration =
                base == Base::Styled && replacement != &default && registered.is_none();
            if needs_registration {
                patched = set_stylesheet(&patched, stylesheet)?;
            }
            let is_empty =
                patched == styled_empty || patched == unstyled_empty || patched == unstyled_base;
            let owners = owner_count(workspace, slot.object)?;
            if owners == 0 {
                return Err(Error::Invalid("a series override has no owner"));
            }
            if owners > 1 {
                if is_empty {
                    final_ids[index] = None;
                    continue;
                }
                let allocation = if registered.is_some() {
                    Base::Styled
                } else {
                    base
                };
                if allocation == Base::Unstyled && container.is_none() {
                    container = Some(reserve_container(workspace, &mut next)?);
                }
                let identifier = allocate(&mut next)?;
                final_ids[index] = Some(identifier);
                let object = new_override(identifier, patched, workspace)?;
                match allocation {
                    Base::Styled => {
                        registrations
                            .entry(chart_location)
                            .or_default()
                            .push(identifier);
                        chart_created.push(identifier);
                        styled_creations.push(object);
                    },
                    Base::Unstyled => unstyled_creations.push(object),
                }
                created.push(identifier);
            } else if is_empty {
                final_ids[index] = None;
                if registered.is_some() {
                    unregistrations
                        .entry(slot.location)
                        .or_default()
                        .push(slot.object);
                }
                removals.push(slot.clone());
            } else {
                if needs_registration {
                    registrations
                        .entry(slot.location)
                        .or_default()
                        .push(slot.object);
                }
                updates.push((slot.clone(), patched));
            }
        } else if replacement != &default {
            let mut data = property.patch(new_base, replacement)?;
            if data.as_slice() == new_base {
                return Err(Error::Invalid("a series write produced no native override"));
            }
            if base == Base::Unstyled {
                append_capabilities(&mut data);
            }
            if base == Base::Unstyled && container.is_none() {
                container = Some(reserve_container(workspace, &mut next)?);
            }
            let identifier = allocate(&mut next)?;
            final_ids[index] = Some(identifier);
            let object = new_override(identifier, data, workspace)?;
            match base {
                Base::Styled => {
                    registrations
                        .entry(chart_location)
                        .or_default()
                        .push(identifier);
                    chart_created.push(identifier);
                    styled_creations.push(object);
                },
                Base::Unstyled => unstyled_creations.push(object),
            }
            created.push(identifier);
        }
    }

    for (slot, data) in updates {
        workspace.replace_payload(slot.object, slot.message, data)?;
    }
    for (location, styles) in &unregistrations {
        unregister_styles(workspace, stylesheet, *location, styles)?;
    }
    // Removal bookkeeping uses each override's member name before removal.
    let removal_members = removals
        .iter()
        .map(|slot| (slot.object, workspace.member(slot.location).to_owned()))
        .collect::<Vec<_>>();
    let mut contained = BTreeMap::<Location, Vec<u64>>::new();
    for slot in &removals {
        if container_of(workspace.archive_at(slot.location)?)?.is_some() {
            contained
                .entry(slot.location)
                .or_default()
                .push(slot.object);
        } else {
            workspace.remove(slot.object)?;
        }
    }
    let mut released_containers = Vec::new();
    for (location, objects) in contained {
        if let Some(identifier) = remove_from_container(workspace, location, &objects)? {
            released_containers.push(identifier);
        }
    }
    for object in styled_creations {
        workspace.insert(chart_location, object)?;
    }
    if !unstyled_creations.is_empty() {
        let container = container.take().ok_or(Error::Invalid(
            "an unstyled override lost its object container",
        ))?;
        insert_container(workspace, &chart_member, container, unstyled_creations)?;
    }
    for (location, styles) in &registrations {
        register_styles(workspace, stylesheet, *location, styles)?;
    }
    let previous = graph
        .slots
        .iter()
        .flatten()
        .map(|slot| slot.object)
        .collect::<Vec<_>>();
    patch_chart_references(workspace, chart, graph.chart_message, &previous, &final_ids)?;

    let mut unbound = BTreeMap::<u64, Vec<u64>>::new();
    for (object, member) in &removal_members {
        if let Some(component) = workspace.component_for_member(member)? {
            if workspace
                .component_uuid_objects(component)?
                .contains(object)
            {
                unbound.entry(component).or_default().push(*object);
            }
        }
    }
    for (component, objects) in unbound {
        workspace.remove_object_uuids(component, &objects)?;
    }
    if !chart_created.is_empty()
        && let Some(component) = workspace.component_for_member(&chart_member)?
    {
        workspace.add_object_uuids(component, &chart_created)?;
    }
    if let Some(last) = created.last().copied() {
        workspace.set_last_object_identifier(last)?;
    }
    let mut released = removals.iter().map(|slot| slot.object).collect::<Vec<_>>();
    released.extend(released_containers);
    if !released.is_empty() {
        workspace.release_identifier_suffix(&released)?;
    }
    if values(workspace, chart, expected.len(), property)? != expected {
        return Err(Error::Invalid("a chart series write does not read back"));
    }
    Ok(true)
}

/// One per-series property stored in the series styles (inherited through
/// the `TSS.StyleArchive.parent` chain).
pub trait StyleProperty {
    /// The stored value.
    type Item: Clone + PartialEq;

    /// The value a style payload defines itself (`None`: inherited).
    ///
    /// # Errors
    ///
    /// The payload is invalid or stores an unsupported value.
    fn read_local(&self, payload: &[u8]) -> Result<Option<Self::Item>>;

    /// Define (`Some`) or clear (`None`) the value in a style payload,
    /// keeping every other byte.
    ///
    /// # Errors
    ///
    /// The payload is invalid or the value cannot be stored.
    fn patch(&self, payload: &[u8], value: Option<&Self::Item>) -> Result<Vec<u8>>;
}

trait BoundedProperty: Property {
    fn read_bounded(&self, payload: &[u8], budget: &mut SymbolBudget) -> SymbolResult<Self::Value>;

    fn patch_bounded(
        &self,
        payload: &[u8],
        value: &Self::Value,
        budget: &mut SymbolBudget,
    ) -> SymbolResult<Vec<u8>>;
}

trait BoundedStyleProperty: StyleProperty {
    fn read_local_bounded(
        &self,
        payload: &[u8],
        budget: &mut SymbolBudget,
    ) -> SymbolResult<Option<Self::Item>>;

    fn patch_bounded(
        &self,
        payload: &[u8],
        value: Option<&Self::Item>,
        budget: &mut SymbolBudget,
    ) -> SymbolResult<Vec<u8>>;
}

/// The effective series style of each of the first `series_count` series:
/// the private override of `series_private_styles`, else the theme style.
fn style_slots(workspace: &Workspace<'_>, chart: u64, series_count: usize) -> Result<Vec<u64>> {
    let (_, payload) = chart_payload(workspace, chart)?;
    let references = codec::references(payload, CodecLimits::DEFAULT)?;
    let mut styles = references.series_theme_styles().to_vec();
    if styles.is_empty() {
        return Err(Error::Invalid("the chart has no series styles"));
    }
    let mut seen = HashSet::new();
    if styles
        .iter()
        .any(|identifier| *identifier == 0 || !seen.insert(*identifier))
    {
        return Err(Error::Invalid(
            "a series theme style is invalid or repeated",
        ));
    }
    if let Some(private) = references.series_private_styles() {
        if usize::try_from(private.count()).ok() != Some(private.entries().len()) {
            return Err(Error::Invalid(
                "a chart's series_private_styles count disagrees with its entries",
            ));
        }
        let mut indices = HashSet::new();
        for &(index, identifier) in private.entries() {
            let index = usize::try_from(index)
                .map_err(|_| Error::Invalid("a series index exceeds usize"))?;
            if index >= styles.len() || !indices.insert(index) || !seen.insert(identifier) {
                return Err(Error::Invalid("a private series style is invalid"));
            }
            styles[index] = identifier;
        }
    }
    if styles.len() < series_count {
        return Err(Error::Invalid(
            "the chart has fewer series styles than series",
        ));
    }
    styles.truncate(series_count);
    for &style in &styles {
        workspace.unique_payload(
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
        )?;
    }
    Ok(styles)
}

fn inherited<P: StyleProperty>(
    workspace: &Workspace<'_>,
    style: u64,
    property: &P,
) -> Result<Option<P::Item>> {
    let mut identifier = style;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(identifier) {
            return Err(Error::Invalid("a series style parent chain has a cycle"));
        }
        let (_, payload) = workspace.unique_payload(
            identifier,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
        )?;
        if let Some(value) = property.read_local(payload)? {
            return Ok(Some(value));
        }
        let parent = graph::repeated(payload, STYLE_PARENT_FIELD)?;
        let parent = match parent.last() {
            Some(style) => match graph::repeated(style, STYLE_PARENT_REFERENCE_FIELD)?.last() {
                Some(reference) => graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?
                    .filter(|identifier| *identifier != 0),
                None => None,
            },
            None => None,
        };
        let Some(parent) = parent else {
            return Ok(None);
        };
        identifier = parent;
    }
}

/// Read `property` (inherited) for each of the `series_count` series.
///
/// # Errors
///
/// An invalid chart or series style, or a value the property rejects.
pub fn style_values<P: StyleProperty>(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    property: &P,
) -> Result<Vec<Option<P::Item>>> {
    style_slots(workspace, chart, series_count)?
        .into_iter()
        .map(|style| inherited(workspace, style, property))
        .collect()
}

/// Refuse a write to a series style another chart also uses.
fn ensure_exclusive(workspace: &Workspace<'_>, style: u64) -> Result<()> {
    let mut owners = 0usize;
    for location in workspace.locations_iter() {
        for candidate in &workspace.archive_at(location)?.objects {
            for message in &candidate.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let references = codec::references(&message.data, CodecLimits::DEFAULT)?;
                let private = references
                    .series_private_styles()
                    .is_some_and(|sparse| sparse.entries().iter().any(|(_, id)| *id == style));
                if private || references.series_theme_styles().contains(&style) {
                    owners += 1;
                }
            }
        }
    }
    if owners != 1 {
        return Err(Error::Invalid("a series style is shared by several charts"));
    }
    Ok(())
}

/// Write `property` in the effective series style of every series whose
/// inherited value differs (`expected.len()` is the series count). Returns
/// `false` when nothing changes.
///
/// # Errors
///
/// An invalid chart or series style, a style another chart shares, a value
/// the property rejects, or a write that does not read back.
pub fn set_style_values<P: StyleProperty>(
    workspace: &mut Workspace<'_>,
    chart: u64,
    property: &P,
    expected: &[Option<P::Item>],
) -> Result<bool> {
    let styles = style_slots(workspace, chart, expected.len())?;
    let mut changed = false;
    for (&style, replacement) in styles.iter().zip(expected) {
        if inherited(workspace, style, property)? == *replacement {
            continue;
        }
        ensure_exclusive(workspace, style)?;
        let (message, payload) = workspace.unique_payload(
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
        )?;
        let data = property.patch(payload, replacement.as_ref())?;
        workspace.replace_payload(style, message, data)?;
        changed = true;
    }
    if changed && style_values(workspace, chart, expected.len(), property)? != expected {
        return Err(Error::Invalid("a series style write does not read back"));
    }
    Ok(changed)
}

pub mod style;

/// The per-series properties of the chart Format panel.
pub mod property {
    use litchi_iwa_common::chart::error_bar::{
        CustomValue, CustomValues, Direction as ErrorBarDirection, FixedValue, Percentage,
        Series as ErrorBarSeries, StandardDeviationCount,
    };
    use litchi_iwa_common::chart::kind::Kind;
    use litchi_iwa_common::chart::number_format::{LabelAffixes, NumberFormat};
    use litchi_iwa_common::chart::pie::{LabelDistance, LabelVisibility, WedgeExplosion};
    use litchi_iwa_common::chart::series::{
        ConnectionLine, ConnectionLineFamily, SymbolFamily, SymbolShape, SymbolSize, Trendline,
        TrendlineType,
    };
    use litchi_iwa_common::chart::series_labels::Visibility;
    use litchi_iwa_common::wire;
    use litchi_iwa_protos::chart_style_codec::{
        self as codec, Archive, Limits as CodecLimits, field, property as rules,
    };

    use super::{
        Base, BoundedProperty, BoundedStyleProperty, Property, StyleProperty, SymbolBudget,
        SymbolError, SymbolResult,
    };
    use crate::graph::{Error, Result};
    use crate::number_format::{self, Fields};

    fn extension(payload: &[u8]) -> Result<Option<&[u8]>> {
        Ok(codec::decode(payload, Archive::SeriesNonStyle, CodecLimits::DEFAULT)?.extension())
    }

    fn bounded_extension<'a>(
        payload: &'a [u8],
        archive: Archive,
        budget: &mut SymbolBudget,
    ) -> SymbolResult<Option<&'a [u8]>> {
        let limits = budget.codec_limits()?;
        let _ = budget.parse_fields(payload, 1)?;
        let _ = budget.parse_fields(payload, 1)?;
        let requirements = codec::decode_requirements(payload, archive, limits)
            .map_err(|error| SymbolError::Graph(Error::Codec(error)))?;
        budget.preflight_codec_report(requirements)?;
        budget.charge_codec_report(requirements)?;
        let (properties, report) = codec::decode_with_report(payload, archive, limits)
            .map_err(|error| SymbolError::Graph(Error::Codec(error)))?;
        budget.preflight_codec_report(report)?;
        budget.charge_codec_report(report)?;
        Ok(properties.extension())
    }

    fn bounded_fields(
        extension: &[u8],
        budget: &mut SymbolBudget,
    ) -> SymbolResult<Vec<wire::WireField>> {
        budget.parse_fields(extension, 2)
    }

    fn bounded_varint(
        extension: &[u8],
        fields: &[wire::WireField],
        number: u32,
    ) -> SymbolResult<Option<u64>> {
        let mut value = None;
        for field in fields.iter().filter(|field| field.number() == number) {
            if value.is_some() {
                return Err(SymbolError::Invalid(
                    "a singular override field is repeated",
                ));
            }
            if field.wire_type() != 0 {
                return Err(SymbolError::Invalid(
                    "an override varint has the wrong wire type",
                ));
            }
            let encoded = field
                .payload(extension)
                .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
            let (decoded, _) = litchi_iwa_common::decode_varint_from_bytes(encoded)
                .map_err(|_| SymbolError::Invalid("an override varint is malformed"))?;
            value = Some(decoded);
        }
        Ok(value)
    }

    fn bounded_bool(
        extension: &[u8],
        fields: &[wire::WireField],
        number: u32,
    ) -> SymbolResult<Option<bool>> {
        bounded_varint(extension, fields, number)?.map_or(Ok(None), |value| {
            if value > 1 {
                return Err(SymbolError::Invalid("an override boolean is not 0 or 1"));
            }
            Ok(Some(value == 1))
        })
    }

    fn bounded_enum(
        extension: &[u8],
        fields: &[wire::WireField],
        number: u32,
    ) -> SymbolResult<Option<i32>> {
        bounded_varint(extension, fields, number)?.map_or(Ok(None), |value| {
            i32::try_from(value)
                .map(Some)
                .map_err(|_| SymbolError::Invalid("an override enum exceeds i32"))
        })
    }

    fn bounded_float(
        extension: &[u8],
        fields: &[wire::WireField],
        number: u32,
    ) -> SymbolResult<Option<f32>> {
        let mut value = None;
        for field in fields.iter().filter(|field| field.number() == number) {
            if value.is_some() {
                return Err(SymbolError::Invalid("a singular style field is repeated"));
            }
            if field.wire_type() != 5 {
                return Err(SymbolError::Invalid(
                    "a style float has the wrong wire type",
                ));
            }
            let bytes = field
                .payload(extension)
                .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
            let bytes: [u8; 4] = bytes
                .try_into()
                .map_err(|_| SymbolError::Invalid("a style float has the wrong width"))?;
            value = Some(f32::from_le_bytes(bytes));
        }
        Ok(value)
    }

    fn bounded_patch_mark(
        payload: &[u8],
        family: SymbolFamily,
        value: &Mark,
        budget: &mut SymbolBudget,
    ) -> SymbolResult<Vec<u8>> {
        let existing = bounded_extension(payload, Archive::SeriesNonStyle, budget)?;
        let source = existing.unwrap_or_default();
        let fields = bounded_fields(source, budget)?;
        let show_present = bounded_bool(source, &fields, family.show_field())?.is_some();
        let type_present = bounded_enum(source, &fields, family.type_field())?.is_some();
        let bound = payload.len().saturating_add(32);
        budget.preflight_output(bound)?;
        budget.preflight_retained(bound)?;
        budget.preflight_scratch(bound)?;
        budget.preflight_allocations(2)?;
        let extension = wire::patch_varint_field(
            source,
            family.show_field(),
            show_present,
            value.visible.then_some(1),
        )
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        let native = u64::try_from(value.shape.native_value())
            .map_err(|_| SymbolError::Invalid("a negative symbol shape"))?;
        let extension = wire::patch_varint_field(
            &extension,
            family.type_field(),
            type_present,
            (value.shape != SymbolShape::Circle).then_some(native),
        )
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        budget.allocations(2)?;
        budget.retained(bound)?;
        budget.scratch(bound)?;
        let patched = wire::patch_length_delimited_field(
            payload,
            codec::EXTENSION_FIELD,
            existing.is_some(),
            (!extension.is_empty()).then_some(extension.as_slice()),
        )
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        budget.output(patched.len())?;
        Ok(patched)
    }

    fn strict_varint(extension: &[u8], number: u32) -> Result<Option<u64>> {
        let mut found = None;
        for candidate in crate::graph::fields(extension)? {
            if candidate.number() != number {
                continue;
            }
            if found.is_some() {
                return Err(Error::Invalid("a singular override field is repeated"));
            }
            if candidate.wire_type() != 0 {
                return Err(Error::Invalid("an override varint has the wrong wire type"));
            }
            let (value, _) = litchi_iwa_common::decode_varint_from_bytes(
                &extension[candidate.key_end()..candidate.end()],
            )
            .map_err(|_| Error::Invalid("an override varint is malformed"))?;
            found = Some(value);
        }
        Ok(found)
    }

    fn strict_bool(extension: &[u8], number: u32) -> Result<Option<bool>> {
        strict_varint(extension, number)?
            .map(|value| {
                if value > 1 {
                    return Err(Error::Invalid("an override boolean is not 0 or 1"));
                }
                Ok(value == 1)
            })
            .transpose()
    }

    fn strict_enum(extension: &[u8], number: u32) -> Result<Option<i32>> {
        strict_varint(extension, number)?
            .map(|value| {
                i32::try_from(value).map_err(|_| Error::Invalid("an override enum exceeds i32"))
            })
            .transpose()
    }

    fn rewrite(payload: &[u8], write: &rules::PropertyWrite<'_>) -> Result<Vec<u8>> {
        Ok(codec::rewrite(
            payload,
            write.archive(),
            write.writes(),
            write.empty(),
            CodecLimits::DEFAULT,
        )?
        .into_bytes()
        .into_owned())
    }

    /// How a series connects its data points (line, radar and scatter
    /// charts).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct ConnectionLines(pub ConnectionLineFamily);

    impl Property for ConnectionLines {
        type Value = ConnectionLine;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> ConnectionLine {
            self.0.default_value()
        }

        fn read(&self, payload: &[u8]) -> Result<ConnectionLine> {
            let Some(extension) = extension(payload)? else {
                return Ok(self.0.default_value());
            };
            let line_type = strict_enum(extension, self.0.field())?;
            let show = if self.0 == ConnectionLineFamily::Scatter2d {
                strict_bool(extension, ConnectionLineFamily::SCATTER_SHOW_FIELD)?
            } else {
                None
            };
            self.0
                .read(show, line_type)
                .map_err(|_| Error::Invalid("an unsupported stored connection line"))
        }

        fn patch(&self, payload: &[u8], value: &ConnectionLine) -> Result<Vec<u8>> {
            let family = self.0;
            family
                .validate(*value)
                .map_err(|_| Error::Invalid("the chart kind cannot hide connection lines"))?;
            let existing = extension(payload)?;
            let source = existing.unwrap_or_default();
            let type_present = strict_enum(source, family.field())?.is_some();
            let native = |value: ConnectionLine| {
                u64::try_from(value.native_value())
                    .map_err(|_| Error::Invalid("a negative connection line"))
            };
            let scatter = family == ConnectionLineFamily::Scatter2d;
            let (show, native) = match value {
                ConnectionLine::Hidden => (None, None),
                ConnectionLine::Straight if !scatter => (None, None),
                value if scatter => (Some(1), Some(native(*value)?)),
                value => (None, Some(native(*value)?)),
            };
            let mut extension = source.to_vec();
            if family == ConnectionLineFamily::Scatter2d {
                let show_present =
                    strict_bool(source, ConnectionLineFamily::SCATTER_SHOW_FIELD)?.is_some();
                extension = wire::patch_varint_field(
                    &extension,
                    ConnectionLineFamily::SCATTER_SHOW_FIELD,
                    show_present,
                    show,
                )?;
            }
            let extension =
                wire::patch_varint_field(&extension, family.field(), type_present, native)?;
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a connection-line write does not read back"));
            }
            Ok(patched)
        }
    }

    /// The distance of each pie or donut wedge from the chart center.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct WedgeExplosions;

    impl Property for WedgeExplosions {
        type Value = WedgeExplosion;

        fn base(&self) -> Base {
            Base::Unstyled
        }

        fn default_value(&self) -> WedgeExplosion {
            WedgeExplosion::ZERO
        }

        fn read(&self, payload: &[u8]) -> Result<WedgeExplosion> {
            let properties = codec::decode(payload, Archive::SeriesNonStyle, CodecLimits::DEFAULT)?;
            properties
                .float(field::series_non_style::PIE_WEDGE_EXPLOSION)?
                .map(WedgeExplosion::from_native)
                .transpose()
                .map_err(|_| Error::Invalid("a stored wedge explosion is out of range"))
                .map(Option::unwrap_or_default)
        }

        fn patch(&self, payload: &[u8], value: &WedgeExplosion) -> Result<Vec<u8>> {
            rewrite(
                payload,
                &rules::pie_wedge_explosion(value.native_fraction()),
            )
        }
    }

    /// The distance of each pie or donut wedge's label from the center.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct LabelDistances;

    impl Property for LabelDistances {
        type Value = LabelDistance;

        fn base(&self) -> Base {
            Base::Unstyled
        }

        fn default_value(&self) -> LabelDistance {
            LabelDistance::DEFAULT
        }

        fn read(&self, payload: &[u8]) -> Result<LabelDistance> {
            let properties = codec::decode(payload, Archive::SeriesNonStyle, CodecLimits::DEFAULT)?;
            let current = properties.float(field::series_non_style::PIE_LABEL_EXPLOSION)?;
            let legacy = properties.float(field::series_non_style::PIE_LEGACY_LABEL_EXPLOSION)?;
            if let (Some(current), Some(legacy)) = (current, legacy)
                && current.to_bits() != legacy.to_bits()
            {
                return Err(Error::Invalid("the stored label distances disagree"));
            }
            legacy
                .or(current)
                .map(LabelDistance::from_percent)
                .transpose()
                .map_err(|_| Error::Invalid("a stored label distance is out of range"))
                .map(|value| value.unwrap_or(LabelDistance::DEFAULT))
        }

        fn patch(&self, payload: &[u8], value: &LabelDistance) -> Result<Vec<u8>> {
            rewrite(payload, &rules::pie_label_distance(value.percent()))
        }
    }

    /// A series' data-symbol visibility and shape (the override half of a
    /// data symbol; the size lives in the series style).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct Mark {
        /// Whether the symbol is drawn.
        pub visible: bool,
        /// The marker shape.
        pub shape: SymbolShape,
    }

    impl Mark {
        /// No symbol (the native default).
        pub const HIDDEN: Self = Self {
            visible: false,
            shape: SymbolShape::Circle,
        };
    }

    /// Data-symbol visibility and shape (area, line, radar and scatter
    /// charts).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct SymbolMarks(pub SymbolFamily);

    impl Property for SymbolMarks {
        type Value = Mark;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> Mark {
            Mark::HIDDEN
        }

        fn read(&self, payload: &[u8]) -> Result<Mark> {
            let Some(extension) = extension(payload)? else {
                return Ok(Mark::HIDDEN);
            };
            let visible = strict_bool(extension, self.0.show_field())?.unwrap_or(false);
            let shape =
                SymbolShape::from_native(strict_enum(extension, self.0.type_field())?.unwrap_or(0))
                    .map_err(|_| Error::Invalid("an unsupported stored symbol shape"))?;
            Ok(Mark { visible, shape })
        }

        fn patch(&self, payload: &[u8], value: &Mark) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            let source = existing.unwrap_or_default();
            let show_present = strict_bool(source, self.0.show_field())?.is_some();
            let type_present = strict_enum(source, self.0.type_field())?.is_some();
            let extension = wire::patch_varint_field(
                source,
                self.0.show_field(),
                show_present,
                value.visible.then_some(1),
            )?;
            let native = u64::try_from(value.shape.native_value())
                .map_err(|_| Error::Invalid("a negative symbol shape"))?;
            let extension = wire::patch_varint_field(
                &extension,
                self.0.type_field(),
                type_present,
                (value.shape != SymbolShape::Circle).then_some(native),
            )?;
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a data-symbol write does not read back"));
            }
            Ok(patched)
        }
    }

    impl BoundedProperty for SymbolMarks {
        fn read_bounded(&self, payload: &[u8], budget: &mut SymbolBudget) -> SymbolResult<Mark> {
            let Some(extension) = bounded_extension(payload, Archive::SeriesNonStyle, budget)?
            else {
                return Ok(Mark::HIDDEN);
            };
            let fields = bounded_fields(extension, budget)?;
            let visible = bounded_bool(extension, &fields, self.0.show_field())?.unwrap_or(false);
            let shape = SymbolShape::from_native(
                bounded_enum(extension, &fields, self.0.type_field())?.unwrap_or(0),
            )
            .map_err(|_| SymbolError::Invalid("an unsupported stored symbol shape"))?;
            Ok(Mark { visible, shape })
        }

        fn patch_bounded(
            &self,
            payload: &[u8],
            value: &Mark,
            budget: &mut SymbolBudget,
        ) -> SymbolResult<Vec<u8>> {
            bounded_patch_mark(payload, self.0, value, budget)
        }
    }

    /// Data-symbol size in the series style (`None`: automatic).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct SymbolSizes(pub SymbolFamily);

    impl StyleProperty for SymbolSizes {
        type Item = SymbolSize;

        fn read_local(&self, payload: &[u8]) -> Result<Option<SymbolSize>> {
            codec::decode(payload, Archive::SeriesStyle, CodecLimits::DEFAULT)?
                .float(self.0.size_field())?
                .map(SymbolSize::new)
                .transpose()
                .map_err(|_| Error::Invalid("a stored symbol size is not positive"))
        }

        fn patch(&self, payload: &[u8], value: Option<&SymbolSize>) -> Result<Vec<u8>> {
            rewrite(
                payload,
                &rules::symbol_size(self.0.size_field(), value.map(|size| size.points())),
            )
        }
    }

    impl BoundedStyleProperty for SymbolSizes {
        fn read_local_bounded(
            &self,
            payload: &[u8],
            budget: &mut SymbolBudget,
        ) -> SymbolResult<Option<SymbolSize>> {
            let Some(extension) = bounded_extension(payload, Archive::SeriesStyle, budget)? else {
                return Ok(None);
            };
            let fields = bounded_fields(extension, budget)?;
            bounded_float(extension, &fields, self.0.size_field())?
                .map(SymbolSize::new)
                .transpose()
                .map_err(|_| SymbolError::Invalid("a stored symbol size is not positive"))
        }

        fn patch_bounded(
            &self,
            payload: &[u8],
            value: Option<&SymbolSize>,
            budget: &mut SymbolBudget,
        ) -> SymbolResult<Vec<u8>> {
            let _ = bounded_extension(payload, Archive::SeriesStyle, budget)?;
            let limits = budget.codec_limits()?;
            let requirements = codec::decode_requirements(payload, Archive::SeriesStyle, limits)
                .map_err(|error| SymbolError::Graph(Error::Codec(error)))?;
            budget.preflight_codec_report(requirements)?;
            budget.charge_codec_report(requirements)?;
            let bound = payload.len().saturating_add(16);
            budget.preflight_output(bound)?;
            budget.preflight_retained(bound)?;
            budget.preflight_scratch(bound)?;
            budget.preflight_allocations(1)?;
            budget.allocations(1)?;
            budget.retained(bound)?;
            budget.scratch(bound)?;
            let write = rules::symbol_size(self.0.size_field(), value.map(|value| value.points()));
            let patched = codec::rewrite(
                payload,
                write.archive(),
                write.writes(),
                write.empty(),
                limits,
            )
            .map_err(|error| SymbolError::Graph(Error::Codec(error)))?
            .into_bytes()
            .into_owned();
            budget.output(patched.len())?;
            Ok(patched)
        }
    }

    const SHOW_TRENDLINE: u32 = 37;
    const TRENDLINE_LABEL: u32 = 54;
    const POLYNOMIAL_ORDER: u32 = 55;
    const MOVING_AVERAGE_PERIOD: u32 = 56;
    const SHOW_EQUATION: u32 = 59;
    const SHOW_LABEL: u32 = 60;
    const SHOW_R_SQUARED: u32 = 61;
    const TRENDLINE_TYPE: u32 = 62;

    fn strict_bytes(extension: &[u8], number: u32) -> Result<Option<&[u8]>> {
        let mut found = None;
        for candidate in crate::graph::fields(extension)? {
            if candidate.number() != number {
                continue;
            }
            if found.is_some() {
                return Err(Error::Invalid("a singular override field is repeated"));
            }
            if candidate.wire_type() != 2 {
                return Err(Error::Invalid("an override string has the wrong wire type"));
            }
            found = Some(&extension[candidate.payload_start()..candidate.end()]);
        }
        Ok(found)
    }

    fn patch_varint(data: &[u8], number: u32, value: Option<u64>) -> Result<Vec<u8>> {
        let present = strict_varint(data, number)?.is_some();
        Ok(wire::patch_varint_field(data, number, present, value)?)
    }

    fn patch_bytes(data: &[u8], number: u32, value: Option<&[u8]>) -> Result<Vec<u8>> {
        let present = strict_bytes(data, number)?.is_some();
        Ok(wire::patch_length_delimited_field(
            data, number, present, value,
        )?)
    }

    /// The trendline of each series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Trendlines;

    impl Property for Trendlines {
        type Value = Trendline;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> Trendline {
            Trendline::none()
        }

        fn read(&self, payload: &[u8]) -> Result<Trendline> {
            let Some(extension) = extension(payload)? else {
                return Ok(Trendline::none());
            };
            let label = strict_bytes(extension, TRENDLINE_LABEL)?
                .map(|bytes| {
                    std::str::from_utf8(bytes)
                        .map(str::to_owned)
                        .map_err(|_| Error::Invalid("a trendline label is not UTF-8"))
                })
                .transpose()?;
            Trendline::from_native(
                strict_varint(extension, SHOW_TRENDLINE)?,
                strict_varint(extension, TRENDLINE_TYPE)?,
                strict_varint(extension, POLYNOMIAL_ORDER)?,
                strict_varint(extension, MOVING_AVERAGE_PERIOD)?,
                strict_varint(extension, SHOW_LABEL)?,
                strict_varint(extension, SHOW_EQUATION)?,
                strict_varint(extension, SHOW_R_SQUARED)?,
                label,
            )
            .map_err(|_| Error::Invalid("an unsupported stored trendline"))
        }

        fn patch(&self, payload: &[u8], value: &Trendline) -> Result<Vec<u8>> {
            value
                .validate()
                .map_err(|_| Error::Invalid("an invalid trendline"))?;
            let existing = extension(payload)?;
            if existing.is_none() && *value == Trendline::none() {
                return Ok(payload.to_vec());
            }
            let kind = value.trendline_type();
            let visible = kind.is_visible();
            let mut extension = existing.unwrap_or_default().to_vec();
            extension = patch_varint(&extension, SHOW_TRENDLINE, visible.then_some(1))?;
            extension = patch_varint(
                &extension,
                TRENDLINE_TYPE,
                visible.then_some(i64::from(kind.native_value()) as u64),
            )?;
            if visible {
                extension = patch_bytes(
                    &extension,
                    TRENDLINE_LABEL,
                    value.custom_name().map(str::as_bytes),
                )?;
                extension = patch_varint(
                    &extension,
                    SHOW_LABEL,
                    value.shows_name_in_legend().then_some(1),
                )?;
                match kind {
                    TrendlineType::Polynomial(order) => {
                        extension = patch_varint(
                            &extension,
                            POLYNOMIAL_ORDER,
                            Some(u64::from(order.get())),
                        )?;
                    },
                    TrendlineType::MovingAverage(period) => {
                        extension = patch_varint(
                            &extension,
                            MOVING_AVERAGE_PERIOD,
                            Some(u64::from(period.get())),
                        )?;
                    },
                    _ => {},
                }
                if kind.supports_equation() {
                    extension = patch_varint(
                        &extension,
                        SHOW_EQUATION,
                        value.shows_equation().then_some(1),
                    )?;
                    extension = patch_varint(
                        &extension,
                        SHOW_R_SQUARED,
                        value.shows_r_squared().then_some(1),
                    )?;
                }
            }
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a trendline write does not read back"));
            }
            Ok(patched)
        }
    }

    const ERROR_CUSTOM_NEGATIVE: u32 = 2;
    const ERROR_CUSTOM_POSITIVE: u32 = 4;
    const ERROR_FIXED_VALUE: u32 = 6;
    const ERROR_PERCENTAGE: u32 = 8;
    const ERROR_DIRECTION: u32 = 10;
    const ERROR_STANDARD_DEVIATION: u32 = 12;
    const ERROR_VALUE_TYPE: u32 = 14;
    const SHOW_ERROR_BARS: u32 = 27;
    const CUSTOM_ARRAY_VALUE: u32 = 1;
    const NATIVE_DIRECTION_BOTH: i32 = 1;
    const NATIVE_FIXED_VALUE: i32 = 1;
    const NATIVE_PERCENTAGE: i32 = 2;
    const NATIVE_STANDARD_DEVIATION: i32 = 3;
    const NATIVE_STANDARD_ERROR: i32 = 4;
    const NATIVE_CUSTOM_VALUES: i32 = 5;

    fn strict_f32(extension: &[u8], number: u32) -> Result<Option<f32>> {
        let mut found = None;
        for candidate in crate::graph::fields(extension)? {
            if candidate.number() != number {
                continue;
            }
            if found.is_some() {
                return Err(Error::Invalid("a singular override field is repeated"));
            }
            if candidate.wire_type() != 5 {
                return Err(Error::Invalid("an override float has the wrong wire type"));
            }
            let bytes: [u8; 4] = extension[candidate.payload_start()..candidate.end()]
                .try_into()
                .map_err(|_| Error::Invalid("a truncated override float"))?;
            let value = f32::from_le_bytes(bytes);
            if !value.is_finite() {
                return Err(Error::Invalid("an override float is not finite"));
            }
            found = Some(value);
        }
        Ok(found)
    }

    fn patch_fixed32(data: &[u8], number: u32, value: Option<u32>) -> Result<Vec<u8>> {
        let present = strict_f32(data, number)?.is_some();
        Ok(wire::patch_fixed32_field(data, number, present, value)?)
    }

    fn whole(value: f32, maximum: f32) -> Result<f32> {
        if !value.is_finite() || value.fract() != 0.0 || value < 0.0 || value > maximum {
            return Err(Error::Invalid(
                "an error-bar count is not an unsigned integer",
            ));
        }
        Ok(value)
    }

    fn custom_values(extension: &[u8], number: u32) -> Result<Vec<CustomValue>> {
        let Some(nested) = strict_bytes(extension, number)? else {
            return Ok(Vec::new());
        };
        wire::repeated_fixed64_values(nested, CUSTOM_ARRAY_VALUE)?
            .into_iter()
            .map(|bits| {
                CustomValue::new(f64::from_bits(bits))
                    .map_err(|_| Error::Invalid("an invalid custom error value"))
            })
            .collect()
    }

    fn patch_custom_values(
        extension: &[u8],
        number: u32,
        values: &[CustomValue],
    ) -> Result<Vec<u8>> {
        let existing = strict_bytes(extension, number)?;
        let nested = wire::rewrite_repeated_fixed64_fields(
            existing.unwrap_or_default(),
            CUSTOM_ARRAY_VALUE,
            &values
                .iter()
                .map(|value| value.value().to_bits())
                .collect::<Vec<_>>(),
        )?;
        Ok(wire::patch_length_delimited_field(
            extension,
            number,
            existing.is_some(),
            (!nested.is_empty()).then_some(nested.as_slice()),
        )?)
    }

    /// The error bars of each series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct ErrorBars;

    impl Property for ErrorBars {
        type Value = ErrorBarSeries;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> ErrorBarSeries {
            ErrorBarSeries::None
        }

        fn read(&self, payload: &[u8]) -> Result<ErrorBarSeries> {
            let invalid = |_| Error::Invalid("an unsupported stored error bar");
            let Some(extension) = extension(payload)? else {
                return Ok(ErrorBarSeries::None);
            };
            let visible = strict_bool(extension, SHOW_ERROR_BARS)?.unwrap_or(false);
            let direction =
                strict_enum(extension, ERROR_DIRECTION)?.unwrap_or(NATIVE_DIRECTION_BOTH);
            let native = strict_enum(extension, ERROR_VALUE_TYPE)?.unwrap_or(NATIVE_FIXED_VALUE);
            let fixed = strict_f32(extension, ERROR_FIXED_VALUE)?;
            let percentage = strict_f32(extension, ERROR_PERCENTAGE)?;
            let deviations = strict_f32(extension, ERROR_STANDARD_DEVIATION)?;
            let negative = custom_values(extension, ERROR_CUSTOM_NEGATIVE)?;
            let positive = custom_values(extension, ERROR_CUSTOM_POSITIVE)?;
            if !visible {
                return Ok(ErrorBarSeries::None);
            }
            let direction = ErrorBarDirection::from_native(direction);
            let value = match native {
                NATIVE_FIXED_VALUE => ErrorBarSeries::FixedValue {
                    direction,
                    value: FixedValue::new(fixed.unwrap_or(FixedValue::DEFAULT.value()))
                        .map_err(invalid)?,
                },
                NATIVE_PERCENTAGE => ErrorBarSeries::Percentage {
                    direction,
                    percentage: Percentage::new(whole(
                        percentage.unwrap_or(f32::from(Percentage::DEFAULT.value())),
                        f32::from(u8::MAX),
                    )? as u8)
                    .map_err(invalid)?,
                },
                NATIVE_STANDARD_DEVIATION => ErrorBarSeries::StandardDeviation {
                    direction,
                    deviations: StandardDeviationCount::new(whole(
                        deviations.unwrap_or(StandardDeviationCount::DEFAULT.value() as f32),
                        u32::MAX as f32,
                    )? as u32)
                    .map_err(invalid)?,
                },
                NATIVE_STANDARD_ERROR => ErrorBarSeries::StandardError { direction },
                NATIVE_CUSTOM_VALUES => ErrorBarSeries::CustomValues {
                    direction,
                    values: CustomValues::from_validated(positive, negative).map_err(invalid)?,
                },
                native => ErrorBarSeries::unsupported(direction, native).map_err(invalid)?,
            };
            value.validate().map_err(invalid)?;
            Ok(value)
        }

        fn patch(&self, payload: &[u8], value: &ErrorBarSeries) -> Result<Vec<u8>> {
            value
                .validate()
                .map_err(|_| Error::Invalid("an invalid error bar"))?;
            let existing = extension(payload)?;
            if existing.is_none() && *value == ErrorBarSeries::None {
                return Ok(payload.to_vec());
            }
            let mut extension = existing.unwrap_or_default().to_vec();
            let direction = value.direction();
            extension = patch_varint(
                &extension,
                SHOW_ERROR_BARS,
                direction.is_some().then_some(1),
            )?;
            if let Some(direction) = direction {
                extension = patch_varint(
                    &extension,
                    ERROR_DIRECTION,
                    Some(i64::from(direction.native_value()) as u64),
                )?;
                extension = patch_varint(
                    &extension,
                    ERROR_VALUE_TYPE,
                    value.native_type().map(|native| i64::from(native) as u64),
                )?;
                match value {
                    ErrorBarSeries::FixedValue { value, .. } => {
                        extension = patch_fixed32(
                            &extension,
                            ERROR_FIXED_VALUE,
                            Some(value.value().to_bits()),
                        )?;
                    },
                    ErrorBarSeries::Percentage { percentage, .. } => {
                        extension = patch_fixed32(
                            &extension,
                            ERROR_PERCENTAGE,
                            Some(f32::from(percentage.value()).to_bits()),
                        )?;
                    },
                    ErrorBarSeries::StandardDeviation { deviations, .. } => {
                        extension = patch_fixed32(
                            &extension,
                            ERROR_STANDARD_DEVIATION,
                            Some((deviations.value() as f32).to_bits()),
                        )?;
                    },
                    ErrorBarSeries::CustomValues { values, .. } => {
                        extension = patch_custom_values(
                            &extension,
                            ERROR_CUSTOM_NEGATIVE,
                            values.negative(),
                        )?;
                        extension = patch_custom_values(
                            &extension,
                            ERROR_CUSTOM_POSITIVE,
                            values.positive(),
                        )?;
                    },
                    _ => {},
                }
            }
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("an error-bar write does not read back"));
            }
            Ok(patched)
        }
    }

    const PIE_SHOW_NAMES: u32 = 31;
    const PIE_SHOW_VALUES: u32 = 44;

    /// Data-point-name and value label visibility of each pie or donut
    /// wedge.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct PieLabels;

    impl Property for PieLabels {
        type Value = LabelVisibility;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> LabelVisibility {
            LabelVisibility::DEFAULT
        }

        fn read(&self, payload: &[u8]) -> Result<LabelVisibility> {
            let Some(extension) = extension(payload)? else {
                return Ok(LabelVisibility::DEFAULT);
            };
            Ok(LabelVisibility::new(
                strict_bool(extension, PIE_SHOW_NAMES)?.unwrap_or(false),
                strict_bool(extension, PIE_SHOW_VALUES)?.unwrap_or(true),
            ))
        }

        fn patch(&self, payload: &[u8], value: &LabelVisibility) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            if existing.is_none() && *value == LabelVisibility::DEFAULT {
                return Ok(payload.to_vec());
            }
            let source = existing.unwrap_or_default();
            let names_present = strict_bool(source, PIE_SHOW_NAMES)?.is_some();
            let values_present = strict_bool(source, PIE_SHOW_VALUES)?.is_some();
            let extension = wire::patch_varint_field(
                source,
                PIE_SHOW_NAMES,
                names_present,
                value.data_point_names_visible().then_some(1),
            )?;
            let extension = wire::patch_varint_field(
                &extension,
                PIE_SHOW_VALUES,
                values_present,
                (!value.values_visible()).then_some(0),
            )?;
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a pie-label write does not read back"));
            }
            Ok(patched)
        }
    }

    /// Value-label visibility of each series, in the chart family's field.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct ValueLabels {
        field: u32,
        default: Visibility,
    }

    impl ValueLabels {
        /// The value-label storage of a chart kind, if it has value labels.
        #[must_use]
        pub fn for_chart_kind(kind: Kind) -> Option<Self> {
            let (field, default) = match kind {
                Kind::Area2d | Kind::Area3d | Kind::StackedArea2d | Kind::StackedArea3d => {
                    (38, Visibility::Hidden)
                },
                Kind::Column2d
                | Kind::Column3d
                | Kind::Bar2d
                | Kind::Bar3d
                | Kind::StackedColumn2d
                | Kind::StackedColumn3d
                | Kind::StackedBar2d
                | Kind::StackedBar3d
                | Kind::MultiDataColumn2d
                | Kind::MultiDataBar2d => (39, Visibility::Hidden),
                Kind::Bubble2d | Kind::MultiDataBubble2d => (40, Visibility::Hidden),
                Kind::Line2d | Kind::Line3d => (42, Visibility::Hidden),
                Kind::Mixed2d | Kind::TwoAxis2d => (43, Visibility::Hidden),
                Kind::Pie2d | Kind::Pie3d | Kind::Donut2d | Kind::Donut3d => {
                    (PIE_SHOW_VALUES, Visibility::Visible)
                },
                Kind::Radar2d => (162, Visibility::Hidden),
                Kind::Scatter2d | Kind::MultiDataScatter2d => (45, Visibility::Hidden),
                _ => return None,
            };
            Some(Self { field, default })
        }
    }

    impl Property for ValueLabels {
        type Value = Visibility;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> Visibility {
            self.default
        }

        fn read(&self, payload: &[u8]) -> Result<Visibility> {
            let Some(extension) = extension(payload)? else {
                return Ok(self.default);
            };
            Ok(strict_bool(extension, self.field)?.map_or(self.default, Visibility::from))
        }

        fn patch(&self, payload: &[u8], value: &Visibility) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            if existing.is_none() && *value == self.default {
                return Ok(payload.to_vec());
            }
            let source = existing.unwrap_or_default();
            let present = strict_bool(source, self.field)?.is_some();
            let extension = wire::patch_varint_field(
                source,
                self.field,
                present,
                (*value != self.default).then_some(u64::from(value.is_visible())),
            )?;
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a value-label write does not read back"));
            }
            Ok(patched)
        }
    }

    fn supported(kind: Kind) -> bool {
        kind != Kind::Undefined && !kind.is_unsupported()
    }

    fn store_extension(
        payload: &[u8],
        existing: Option<&[u8]>,
        extension: Option<Vec<u8>>,
    ) -> Result<Vec<u8>> {
        let Some(extension) = extension else {
            return Ok(payload.to_vec());
        };
        Ok(wire::patch_length_delimited_field(
            payload,
            codec::EXTENSION_FIELD,
            existing.is_some(),
            Some(extension.as_slice()),
        )?)
    }

    /// The value-label prefix and suffix of each series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct ValueLabelAffixes(());

    impl ValueLabelAffixes {
        /// The property for a chart kind with value labels.
        #[must_use]
        pub fn for_chart_kind(kind: Kind) -> Option<Self> {
            supported(kind).then_some(Self(()))
        }
    }

    impl Property for ValueLabelAffixes {
        type Value = LabelAffixes;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> LabelAffixes {
            LabelAffixes::default()
        }

        fn read(&self, payload: &[u8]) -> Result<LabelAffixes> {
            number_format::affixes(extension(payload)?, Fields::SERIES_VALUE_LABEL)
        }

        fn patch(&self, payload: &[u8], value: &LabelAffixes) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            if existing.is_none() && value.is_empty() {
                return Ok(payload.to_vec());
            }
            let next = number_format::patch_affixes(
                existing.unwrap_or_default(),
                Fields::SERIES_VALUE_LABEL,
                value,
                NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            )?;
            let patched = store_extension(payload, existing, next)?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid(
                    "a value-label affix write does not read back",
                ));
            }
            Ok(patched)
        }
    }

    /// The value-label number format of each series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct ValueLabelNumberFormats(());

    impl ValueLabelNumberFormats {
        /// The property for a chart kind with value labels.
        #[must_use]
        pub fn for_chart_kind(kind: Kind) -> Option<Self> {
            supported(kind).then_some(Self(()))
        }
    }

    impl Property for ValueLabelNumberFormats {
        type Value = NumberFormat;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> NumberFormat {
            NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT
        }

        fn read(&self, payload: &[u8]) -> Result<NumberFormat> {
            number_format::number_format(
                extension(payload)?,
                Fields::SERIES_VALUE_LABEL,
                NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            )
        }

        fn patch(&self, payload: &[u8], value: &NumberFormat) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            if existing.is_none() && *value == NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT {
                return Ok(payload.to_vec());
            }
            let next = number_format::patch_number_format(
                existing.unwrap_or_default(),
                Fields::SERIES_VALUE_LABEL,
                *value,
                NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            )?;
            let patched = store_extension(payload, existing, next)?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid(
                    "a value-label number-format write does not read back",
                ));
            }
            Ok(patched)
        }
    }
}

pub(crate) fn bounded_references(
    payload: &[u8],
    budget: &mut SymbolBudget,
) -> SymbolResult<codec::StyleReferences> {
    let _ = budget.parse_fields(payload, 1)?;
    let _ = budget.parse_fields(payload, 1)?;
    let limits = budget.codec_limits()?;
    let requirements = codec::reference_requirements(payload, limits)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?;
    budget.preflight_codec_report(requirements)?;
    budget.charge_codec_report(requirements)?;
    let (references, report) = codec::references_with_report(payload, limits)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?;
    budget.preflight_codec_report(report)?;
    budget.charge_codec_report(report)?;
    let count = references.chart_style().is_some() as usize
        + references.chart_non_style().is_some() as usize
        + references.legend_style().is_some() as usize
        + references.legend_non_style().is_some() as usize
        + references.value_axis_styles().len()
        + references.value_axis_non_styles().len()
        + references.category_axis_styles().len()
        + references.category_axis_non_styles().len()
        + references.series_theme_styles().len()
        + references
            .series_private_styles()
            .map_or(0, |sparse| sparse.entries().len())
        + references
            .series_non_styles()
            .map_or(0, |sparse| sparse.entries().len())
        + references.paragraph_styles().len();
    budget.references(count)?;
    Ok(references)
}

pub(crate) fn bounded_payload<'a>(
    workspace: &'a Workspace<'_>,
    identifier: u64,
    message_type: u32,
    reason: &'static str,
    budget: &mut SymbolBudget,
) -> SymbolResult<(usize, &'a [u8])> {
    budget.work(1)?;
    let (message, payload) = workspace
        .unique_payload(identifier, message_type, reason)
        .map_err(SymbolError::from)?;
    Ok((message, payload))
}

pub(crate) fn bounded_chart_payload<'a>(
    workspace: &'a Workspace<'_>,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<(usize, &'a [u8])> {
    let payload = bounded_payload(
        workspace,
        chart,
        CHART_MESSAGE_TYPE,
        "a chart must have exactly one chart payload",
        budget,
    )?;
    Ok(payload)
}

fn bounded_series_graph(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    budget: &mut SymbolBudget,
) -> SymbolResult<SeriesGraph> {
    if series_count == 0 {
        return Err(SymbolError::Invalid(
            "a chart series operation names no series",
        ));
    }
    let (chart_message, payload) = bounded_chart_payload(workspace, chart, budget)?;
    let references = bounded_references(payload, budget)?;
    budget.reserve::<Option<Slot>>(series_count, true)?;
    let mut slots: Vec<Option<Slot>> = Vec::new();
    slots
        .try_reserve_exact(series_count)
        .map_err(|_| SymbolError::Allocation {
            amount: series_count,
        })?;
    slots.resize(series_count, None);
    if let Some(sparse) = references.series_non_styles() {
        if usize::try_from(sparse.count()).ok() != Some(sparse.entries().len()) {
            return Err(SymbolError::Invalid(
                "a chart's series_non_styles count disagrees with its entries",
            ));
        }
        for &(index, object) in sparse.entries() {
            budget.work(1)?;
            let index = usize::try_from(index)
                .map_err(|_| SymbolError::Invalid("a series index exceeds usize"))?;
            if index >= series_count {
                return Err(SymbolError::Invalid(
                    "a series override index exceeds the series count",
                ));
            }
            if slots[index].is_some() || slots.iter().flatten().any(|slot| slot.object == object) {
                return Err(SymbolError::Invalid("a series override is repeated"));
            }
            let location = workspace.locate(object).map_err(SymbolError::from)?;
            let (message, _) = bounded_payload(
                workspace,
                object,
                SERIES_NON_STYLE_MESSAGE_TYPE,
                "a series override must have exactly one payload",
                budget,
            )?;
            slots[index] = Some(Slot {
                location,
                object,
                message,
            });
        }
    }
    Ok(SeriesGraph {
        chart_message,
        slots,
    })
}

fn bounded_slot_payload<'a>(workspace: &'a Workspace<'_>, slot: &Slot) -> SymbolResult<&'a [u8]> {
    let object = workspace.object(slot.object).map_err(SymbolError::from)?;
    let message = object
        .messages
        .get(slot.message)
        .ok_or(SymbolError::Invalid(
            "a series override payload index changed",
        ))?;
    if message.type_ != SERIES_NON_STYLE_MESSAGE_TYPE {
        return Err(SymbolError::Invalid(
            "a series override payload type changed",
        ));
    }
    Ok(message.data.as_slice())
}

fn bounded_read_values<P: BoundedProperty>(
    workspace: &Workspace<'_>,
    graph: &SeriesGraph,
    property: &P,
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<P::Value>> {
    budget.reserve::<P::Value>(graph.slots.len(), true)?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(graph.slots.len())
        .map_err(|_| SymbolError::Allocation {
            amount: graph.slots.len(),
        })?;
    for slot in &graph.slots {
        budget.work(1)?;
        let value = match slot {
            Some(slot) => property.read_bounded(bounded_slot_payload(workspace, slot)?, budget)?,
            None => property.default_value(),
        };
        values.push(value);
    }
    Ok(values)
}

/// The series style objects of a chart drawable's style references: its
/// theme styles with the private overrides applied, one per series slot
/// (the host's `chart_series_style_slots`).
///
/// # Errors
///
/// No series style, a zero or repeated identifier, or a private override
/// outside the theme styles or disagreeing with its count.
pub fn chart_series_styles(references: &codec::StyleReferences) -> SymbolResult<Vec<u64>> {
    let theme = references.series_theme_styles();
    if theme.is_empty() {
        return Err(SymbolError::Invalid("the chart has no series styles"));
    }
    let mut styles = Vec::new();
    styles
        .try_reserve_exact(theme.len())
        .map_err(|_| SymbolError::Allocation {
            amount: theme.len(),
        })?;
    styles.extend_from_slice(theme);
    if styles.contains(&0) {
        return Err(SymbolError::Invalid(
            "a series theme style is invalid or repeated",
        ));
    }
    for left in 0..styles.len() {
        if styles[left + 1..].contains(&styles[left]) {
            return Err(SymbolError::Invalid(
                "a series theme style is invalid or repeated",
            ));
        }
    }
    if let Some(private) = references.series_private_styles() {
        if usize::try_from(private.count()).ok() != Some(private.entries().len()) {
            return Err(SymbolError::Invalid(
                "a chart's series_private_styles count disagrees with its entries",
            ));
        }
        for &(index, identifier) in private.entries() {
            let index = usize::try_from(index)
                .map_err(|_| SymbolError::Invalid("a series index exceeds usize"))?;
            if index >= styles.len()
                || styles[index] == identifier
                || styles[..index].contains(&identifier)
                || styles[index + 1..].contains(&identifier)
            {
                return Err(SymbolError::Invalid("a private series style is invalid"));
            }
            styles[index] = identifier;
        }
    }
    Ok(styles)
}

fn bounded_style_slots(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<u64>> {
    let (_, payload) = bounded_chart_payload(workspace, chart, budget)?;
    let references = bounded_references(payload, budget)?;
    budget.reserve::<u64>(references.series_theme_styles().len(), true)?;
    if let Some(private) = references.series_private_styles() {
        for _ in private.entries() {
            budget.work(1)?;
        }
    }
    let mut styles = chart_series_styles(&references)?;
    // `usize::MAX`: every series style (the chart-level shadow).
    if series_count != usize::MAX {
        if styles.len() < series_count {
            return Err(SymbolError::Invalid(
                "the chart has fewer series styles than series",
            ));
        }
        styles.truncate(series_count);
    }
    for &style in &styles {
        let _ = bounded_payload(
            workspace,
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
            budget,
        )?;
    }
    Ok(styles)
}

fn bounded_style_parent(payload: &[u8], budget: &mut SymbolBudget) -> SymbolResult<Option<u64>> {
    let fields = budget.parse_fields(payload, 1)?;
    let parent = fields
        .iter()
        .rev()
        .find(|field| field.number() == STYLE_PARENT_FIELD)
        .ok_or(SymbolError::Invalid(
            "a chart style has no single style parent",
        ))?;
    if parent.wire_type() != 2 {
        return Err(SymbolError::Invalid(
            "a chart style parent has the wrong wire type",
        ));
    }
    let parent_payload = parent
        .payload(payload)
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    let parent_fields = budget.parse_fields(parent_payload, 2)?;
    let reference = parent_fields
        .iter()
        .rev()
        .find(|field| field.number() == STYLE_PARENT_REFERENCE_FIELD);
    let Some(reference) = reference else {
        return Ok(None);
    };
    if reference.wire_type() != 2 {
        return Err(SymbolError::Invalid(
            "a style parent reference has the wrong wire type",
        ));
    }
    let reference_payload = reference
        .payload(parent_payload)
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    let reference_fields = budget.parse_fields(reference_payload, 3)?;
    let identifier = reference_fields
        .iter()
        .rev()
        .find(|field| field.number() == REFERENCE_IDENTIFIER_FIELD);
    let Some(identifier) = identifier else {
        return Ok(None);
    };
    if identifier.wire_type() != 0 {
        return Err(SymbolError::Invalid(
            "a style parent identifier is not a varint",
        ));
    }
    let bytes = identifier
        .payload(reference_payload)
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    let (value, width) = litchi_iwa_common::decode_varint_from_bytes(bytes)
        .map_err(|_| SymbolError::Invalid("a style parent identifier is malformed"))?;
    if width != bytes.len() || value == 0 {
        return Err(SymbolError::Invalid(
            "a style parent identifier is noncanonical",
        ));
    }
    budget.references(1)?;
    Ok(Some(value))
}

fn bounded_inherited<P: BoundedStyleProperty>(
    workspace: &Workspace<'_>,
    style: u64,
    property: &P,
    budget: &mut SymbolBudget,
) -> SymbolResult<Option<P::Item>> {
    budget.reserve::<u64>(budget.limits.max_nesting as usize, true)?;
    let mut visited = Vec::new();
    visited
        .try_reserve_exact(budget.limits.max_nesting as usize)
        .map_err(|_| SymbolError::Allocation {
            amount: budget.limits.max_nesting as usize,
        })?;
    let mut identifier = style;
    for depth in 1..=budget.limits.max_nesting {
        budget.depth(depth)?;
        if visited.contains(&identifier) {
            return Err(SymbolError::Invalid(
                "a series style parent chain has a cycle",
            ));
        }
        budget.work(1)?;
        visited.push(identifier);
        let (_, payload) = bounded_payload(
            workspace,
            identifier,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
            budget,
        )?;
        if let Some(value) = property.read_local_bounded(payload, budget)? {
            return Ok(Some(value));
        }
        let Some(parent) = bounded_style_parent(payload, budget)? else {
            return Ok(None);
        };
        identifier = parent;
    }
    Err(SymbolError::Limit {
        kind: LimitKind::Nesting,
        observed: budget.limits.max_nesting as usize + 1,
        maximum: budget.limits.max_nesting as usize,
    })
}

fn bounded_read_sizes(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    property: &property::SymbolSizes,
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<Option<litchi_iwa_common::chart::series::SymbolSize>>> {
    let styles = bounded_style_slots(workspace, chart, series_count, budget)?;
    budget.reserve::<Option<litchi_iwa_common::chart::series::SymbolSize>>(styles.len(), true)?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(styles.len())
        .map_err(|_| SymbolError::Allocation {
            amount: styles.len(),
        })?;
    for style in styles {
        values.push(bounded_inherited(workspace, style, property, budget)?);
    }
    Ok(values)
}

fn bounded_mutation_scope(
    workspace: &Workspace<'_>,
    chart: u64,
    property_fields: &[u32],
    series_count: usize,
    budget: &mut SymbolBudget,
) -> SymbolResult<MutationScope> {
    let graph = bounded_series_graph(workspace, chart, series_count, budget)?;
    let styles = bounded_style_slots(workspace, chart, series_count, budget)?;
    let stylesheet = bounded_chart_stylesheet(workspace, chart, budget)?;
    let object_capacity = graph
        .slots
        .iter()
        .flatten()
        .count()
        .saturating_add(styles.len())
        .saturating_add(2);
    budget.reserve::<u64>(object_capacity, true)?;
    let mut allowed_objects = Vec::new();
    allowed_objects
        .try_reserve_exact(object_capacity)
        .map_err(|_| SymbolError::Allocation {
            amount: object_capacity,
        })?;
    for identifier in std::iter::once(chart)
        .chain(graph.slots.iter().flatten().map(|slot| slot.object))
        .chain(styles.iter().copied())
        .chain(std::iter::once(stylesheet))
    {
        if !allowed_objects.contains(&identifier) {
            allowed_objects.push(identifier);
        }
    }
    let mut allowed_members = Vec::new();
    let member_capacity = graph.slots.iter().flatten().count().saturating_add(5);
    budget.preflight_allocations(1)?;
    budget.preflight_retained(member_capacity.saturating_mul(size_of::<String>()))?;
    budget.allocations(1)?;
    budget.retained(member_capacity.saturating_mul(size_of::<String>()))?;
    allowed_members
        .try_reserve_exact(member_capacity)
        .map_err(|_| SymbolError::Allocation {
            amount: member_capacity,
        })?;
    let mut add_member = |location: Location| -> SymbolResult<()> {
        let source_member = workspace.member(location);
        if !allowed_members
            .iter()
            .any(|candidate| candidate == source_member)
        {
            budget.preflight_allocations(1)?;
            budget.preflight_retained(source_member.len())?;
            budget.allocations(1)?;
            let member = source_member.to_owned();
            budget.retained(member.len())?;
            allowed_members.push(member);
        }
        Ok(())
    };
    add_member(workspace.locate(chart).map_err(SymbolError::from)?)?;
    for slot in graph.slots.iter().flatten() {
        add_member(slot.location)?;
    }
    add_member(workspace.locate(stylesheet).map_err(SymbolError::from)?)?;
    if workspace.has_member(METADATA_MEMBER) {
        let location = workspace
            .member_location(METADATA_MEMBER)
            .ok_or(SymbolError::Invalid("metadata member disappeared"))?;
        add_member(location)?;
    }
    if workspace.has_member(DOCUMENT_METADATA_MEMBER) {
        let location = workspace
            .member_location(DOCUMENT_METADATA_MEMBER)
            .ok_or(SymbolError::Invalid("document metadata member disappeared"))?;
        add_member(location)?;
    }
    let mut allowed_message_types = Vec::new();
    let message_types = [
        CHART_MESSAGE_TYPE,
        SERIES_STYLE_MESSAGE_TYPE,
        SERIES_NON_STYLE_MESSAGE_TYPE,
        STYLESHEET_MESSAGE_TYPE,
        OBJECT_CONTAINER_MESSAGE_TYPE,
        11_006,
    ];
    budget.reserve::<u32>(message_types.len(), true)?;
    allowed_message_types
        .try_reserve_exact(message_types.len())
        .map_err(|_| SymbolError::Allocation {
            amount: message_types.len(),
        })?;
    allowed_message_types.extend_from_slice(&message_types);
    let leading = [
        codec::EXTENSION_FIELD,
        17,
        18,
        19,
        STYLE_PARENT_FIELD,
        STYLE_PARENT_REFERENCE_FIELD,
        STYLE_STYLESHEET_FIELD,
        STYLESHEET_STYLES_FIELD,
        CONTAINER_IDENTIFIER_FIELD,
        CONTAINER_OBJECTS_FIELD,
        REFERENCE_IDENTIFIER_FIELD,
        1,
        2,
    ];
    let trailing = [3, 6, 11];
    let field_count = leading
        .len()
        .saturating_add(property_fields.len())
        .saturating_add(trailing.len());
    budget.reserve::<u32>(field_count, true)?;
    let mut allowed_fields = Vec::new();
    allowed_fields
        .try_reserve_exact(field_count)
        .map_err(|_| SymbolError::Allocation {
            amount: field_count,
        })?;
    for field in leading
        .iter()
        .chain(property_fields)
        .chain(trailing.iter())
        .copied()
    {
        if !allowed_fields.contains(&field) {
            allowed_fields.push(field);
        }
    }
    let planned_capacity = series_count.max(1);
    let mut planned_new_objects = Vec::new();
    let mut planned_removed_objects = Vec::new();
    let mut planned_new_members = Vec::new();
    let mut planned_removed_members = Vec::new();
    budget.reserve_vec(&mut planned_new_objects, planned_capacity, true)?;
    budget.reserve_vec(&mut planned_removed_objects, planned_capacity, true)?;
    budget.reserve_vec(&mut planned_new_members, planned_capacity, true)?;
    budget.reserve_vec(&mut planned_removed_members, planned_capacity, true)?;
    Ok(MutationScope {
        selected_chart: chart,
        allowed_objects,
        allowed_members,
        allowed_message_types,
        allowed_fields,
        planned_new_objects,
        planned_removed_objects,
        planned_new_members,
        planned_removed_members,
        chart_fields: SERIES_CHART_FIELDS.to_vec(),
    })
}

fn bounded_stylesheet_of(payload: &[u8], budget: &mut SymbolBudget) -> SymbolResult<Option<u64>> {
    let fields = budget.parse_fields(payload, 1)?;
    let parent = fields
        .iter()
        .rev()
        .find(|field| field.number() == STYLE_PARENT_FIELD)
        .ok_or(SymbolError::Invalid(
            "a chart style has no single style parent",
        ))?;
    if parent.wire_type() != 2 {
        return Err(SymbolError::Invalid(
            "a chart style parent has the wrong wire type",
        ));
    }
    let parent_payload = parent
        .payload(payload)
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    let parent_fields = budget.parse_fields(parent_payload, 2)?;
    let Some(reference) = parent_fields
        .iter()
        .rev()
        .find(|field| field.number() == STYLE_STYLESHEET_FIELD)
    else {
        return Ok(None);
    };
    if reference.wire_type() != 2 {
        return Err(SymbolError::Invalid(
            "a stylesheet reference has the wrong wire type",
        ));
    }
    let reference_payload = reference
        .payload(parent_payload)
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    let reference_fields = budget.parse_fields(reference_payload, 3)?;
    let Some(identifier) = reference_fields
        .iter()
        .rev()
        .find(|field| field.number() == REFERENCE_IDENTIFIER_FIELD)
    else {
        return Ok(Some(0));
    };
    if identifier.wire_type() != 0 {
        return Err(SymbolError::Invalid(
            "a stylesheet identifier is not a varint",
        ));
    }
    let bytes = identifier
        .payload(reference_payload)
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    let (value, width) = litchi_iwa_common::decode_varint_from_bytes(bytes)
        .map_err(|_| SymbolError::Invalid("a stylesheet identifier is malformed"))?;
    if width != bytes.len() {
        return Err(SymbolError::Invalid(
            "a stylesheet identifier is noncanonical",
        ));
    }
    budget.references(1)?;
    Ok(Some(value))
}

fn bounded_chart_stylesheet(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<u64> {
    let (_, payload) = bounded_chart_payload(workspace, chart, budget)?;
    let references = bounded_references(payload, budget)?;
    let style = references
        .series_theme_styles()
        .first()
        .copied()
        .filter(|identifier| *identifier != 0)
        .ok_or(SymbolError::Invalid("the chart has no series theme style"))?;
    let (_, style_payload) = bounded_payload(
        workspace,
        style,
        SERIES_STYLE_MESSAGE_TYPE,
        "a series style must have exactly one payload",
        budget,
    )?;
    bounded_stylesheet_of(style_payload, budget)?
        .filter(|identifier| *identifier != 0)
        .ok_or(SymbolError::Invalid(
            "the series theme style has no document stylesheet",
        ))
}

fn bounded_style_parent_payload(
    stylesheet: Option<u64>,
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<u8>> {
    let mut style = Vec::new();
    if let Some(identifier) = stylesheet {
        budget.preflight_allocations(1)?;
        budget.preflight_retained(16)?;
        budget.allocations(1)?;
        budget.retained(16)?;
        style
            .try_reserve_exact(16)
            .map_err(|_| SymbolError::Allocation { amount: 16 })?;
        budget.preflight_allocations(1)?;
        budget.preflight_retained(11)?;
        budget.allocations(1)?;
        budget.retained(11)?;
        let mut reference = Vec::new();
        reference
            .try_reserve_exact(11)
            .map_err(|_| SymbolError::Allocation { amount: 11 })?;
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, identifier);
        push_message_field(&mut style, STYLE_STYLESHEET_FIELD, &reference);
    }
    let bound = style.len().saturating_add(16);
    budget.preflight_allocations(1)?;
    budget.preflight_retained(bound)?;
    budget.allocations(1)?;
    budget.retained(bound)?;
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(bound)
        .map_err(|_| SymbolError::Allocation { amount: bound })?;
    push_message_field(&mut payload, STYLE_PARENT_FIELD, &style);
    Ok(payload)
}

fn bounded_canonical_empty(
    stylesheet: Option<u64>,
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<u8>> {
    let mut payload = bounded_style_parent_payload(stylesheet, budget)?;
    let additional = CAPABILITY_FIELDS.len().saturating_mul(3);
    budget.preflight_retained(additional)?;
    budget.retained(additional)?;
    payload
        .try_reserve(additional)
        .map_err(|_| SymbolError::Allocation { amount: additional })?;
    append_capabilities(&mut payload);
    Ok(payload)
}

fn bounded_override_with_parent(
    stylesheet: Option<u64>,
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<u8>> {
    bounded_style_parent_payload(stylesheet, budget)
}

fn bounded_owner_count(
    workspace: &Workspace<'_>,
    object: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<usize> {
    let mut owners = 0usize;
    for location in workspace.locations_iter() {
        let archive = workspace.archive_at(location).map_err(SymbolError::from)?;
        budget.work(archive.objects.len())?;
        for candidate in &archive.objects {
            for message in &candidate.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let references = bounded_references(&message.data, budget)?;
                if references
                    .series_non_styles()
                    .is_some_and(|sparse| sparse.entries().iter().any(|(_, id)| *id == object))
                {
                    owners = owners.saturating_add(1);
                }
            }
        }
    }
    Ok(owners)
}

fn bounded_patch_chart_references(
    workspace: &mut Workspace<'_>,
    chart: u64,
    message: usize,
    previous: &[u64],
    identifiers: &[Option<u64>],
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    let entries_count = identifiers
        .iter()
        .filter(|identifier| identifier.is_some())
        .count();
    budget.reserve::<(u32, u64)>(entries_count, true)?;
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(entries_count)
        .map_err(|_| SymbolError::Allocation {
            amount: entries_count,
        })?;
    for (index, identifier) in identifiers.iter().copied().enumerate() {
        let Some(identifier) = identifier else {
            continue;
        };
        budget.work(1)?;
        entries.push((
            u32::try_from(index).map_err(|_| SymbolError::Invalid("a series index exceeds u32"))?,
            identifier,
        ));
    }
    let payload = workspace
        .object(chart)
        .map_err(SymbolError::from)?
        .messages
        .get(message)
        .ok_or(SymbolError::Invalid("a chart payload index changed"))?
        .data
        .as_slice();
    let limits = budget.codec_limits()?;
    let _ = budget.parse_fields(payload, 1)?;
    let _ = budget.parse_fields(payload, 1)?;
    let requirements = codec::reference_requirements(payload, limits)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?;
    budget.preflight_codec_report(requirements)?;
    budget.charge_codec_report(requirements)?;
    let (_, report) = codec::references_with_report(payload, limits)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?;
    budget.preflight_codec_report(report)?;
    budget.charge_codec_report(report)?;
    let bound = payload
        .len()
        .saturating_add(entries.len().saturating_mul(16));
    budget.preflight_output(bound)?;
    budget.preflight_retained(bound)?;
    budget.preflight_scratch(bound)?;
    budget.preflight_allocations(1)?;
    budget.allocations(1)?;
    budget.retained(bound)?;
    budget.scratch(bound)?;
    let data = codec::rewrite_series_non_styles(payload, &entries, limits)
        .map_err(|error| SymbolError::Graph(Error::Codec(error)))?
        .into_bytes()
        .into_owned();
    budget.output(data.len())?;
    workspace
        .replace_payload(chart, message, data)
        .map_err(SymbolError::from)?;
    let info = workspace
        .object_mut(chart)
        .map_err(SymbolError::from)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(SymbolError::Invalid("a chart payload has no metadata"))?;
    info.object_references
        .retain(|identifier| !previous.contains(identifier));
    for identifier in identifiers.iter().copied().flatten() {
        if !info.object_references.contains(&identifier) {
            info.object_references.push(identifier);
        }
    }
    budget.references(entries.len())?;
    Ok(())
}

fn bounded_style_member(
    workspace: &Workspace<'_>,
    owner: Location,
    styles: &[u64],
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    for (index, &style) in styles.iter().enumerate() {
        budget.work(1)?;
        if style == 0 || styles[..index].contains(&style) {
            return Err(SymbolError::Invalid(
                "chart style identifiers must be unique and nonzero",
            ));
        }
        let archive = workspace.archive_at(owner).map_err(SymbolError::from)?;
        let object = archive.object(style).ok_or(SymbolError::Invalid(
            "a chart style is not in its owner component",
        ))?;
        let payloads = object
            .messages
            .iter()
            .filter(|message| CHART_STYLE_MESSAGE_TYPES.contains(&message.type_))
            .count();
        if payloads != 1 {
            return Err(SymbolError::Invalid(
                "a chart style must have exactly one style payload",
            ));
        }
    }
    Ok(())
}

fn bounded_register_styles(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    owner: Location,
    styles: &[u64],
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    bounded_style_member(workspace, owner, styles, budget)?;
    let sheet_location = workspace.locate(stylesheet).map_err(SymbolError::from)?;
    let owner_member = workspace.member(owner).to_owned();
    let sheet_member = workspace.member(sheet_location).to_owned();
    let owner_component = workspace
        .component_for_member(&owner_member)
        .map_err(SymbolError::from)?;
    let sheet_component = workspace
        .component_for_member(&sheet_member)
        .map_err(SymbolError::from)?;
    let (message, payload) = bounded_payload(
        workspace,
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
        budget,
    )?;
    let fields = budget.parse_fields(payload, 1)?;
    let mut data = Vec::new();
    let bound = payload
        .len()
        .saturating_add(styles.len().saturating_mul(16));
    budget.preflight_output(bound)?;
    budget.preflight_retained(bound)?;
    budget.preflight_scratch(bound)?;
    budget.preflight_allocations(1)?;
    budget.allocations(1)?;
    budget.retained(bound)?;
    budget.scratch(bound)?;
    data.try_reserve_exact(bound)
        .map_err(|_| SymbolError::Allocation { amount: bound })?;
    data.extend_from_slice(payload);
    for &style in styles {
        for field in fields
            .iter()
            .filter(|field| field.number() == STYLESHEET_STYLES_FIELD)
        {
            let reference = field
                .payload(payload)
                .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
            let reference_fields = budget.parse_fields(reference, 2)?;
            let existing = reference_fields
                .iter()
                .rev()
                .find(|field| field.number() == REFERENCE_IDENTIFIER_FIELD)
                .map(|field| {
                    field
                        .payload(reference)
                        .map_err(|error| SymbolError::Graph(Error::Wire(error)))
                        .and_then(|bytes| {
                            litchi_iwa_common::decode_varint_from_bytes(bytes)
                                .map(|(value, _)| value)
                                .map_err(|_| {
                                    SymbolError::Invalid("a stylesheet reference is malformed")
                                })
                        })
                })
                .transpose()?;
            if existing == Some(style) {
                return Err(SymbolError::Invalid(
                    "the stylesheet already lists a new style",
                ));
            }
        }
        let mut reference = Vec::new();
        reference
            .try_reserve_exact(11)
            .map_err(|_| SymbolError::Allocation { amount: 11 })?;
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, style);
        data = wire::append_repeated_length_delimited_field(
            &data,
            STYLESHEET_STYLES_FIELD,
            &reference,
        )
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        budget.references(1)?;
    }
    budget.output(data.len())?;
    workspace
        .replace_payload(stylesheet, message, data)
        .map_err(SymbolError::from)?;
    let info = workspace
        .object_mut(stylesheet)
        .map_err(SymbolError::from)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(SymbolError::Invalid("a stylesheet payload has no metadata"))?;
    for &style in styles {
        if info.object_references.contains(&style) {
            return Err(SymbolError::Invalid(
                "the stylesheet metadata already references a new style",
            ));
        }
        info.object_references.push(style);
    }
    budget.registry_entries(styles.len())?;
    if let (Some(owner), Some(sheet)) = (owner_component, sheet_component)
        && owner != sheet
    {
        budget.admit_registry_inspection(workspace)?;
        budget.references(1)?;
        workspace
            .add_external_reference(owner, sheet, stylesheet)
            .map_err(SymbolError::from)?;
        for &style in styles {
            budget.references(1)?;
            workspace
                .add_external_reference(sheet, owner, style)
                .map_err(SymbolError::from)?;
        }
    }
    Ok(())
}

fn bounded_unregister_styles(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    owner: Location,
    styles: &[u64],
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    for (index, &style) in styles.iter().enumerate() {
        if style == 0 || styles[..index].contains(&style) {
            return Err(SymbolError::Invalid(
                "chart style identifiers must be unique and nonzero",
            ));
        }
    }
    let sheet_location = workspace.locate(stylesheet).map_err(SymbolError::from)?;
    let local = sheet_location == owner;
    let owner_member = workspace.member(owner).to_owned();
    let sheet_member = workspace.member(sheet_location).to_owned();
    let owner_component = workspace
        .component_for_member(&owner_member)
        .map_err(SymbolError::from)?;
    let sheet_component = workspace
        .component_for_member(&sheet_member)
        .map_err(SymbolError::from)?;
    let (message, payload) = bounded_payload(
        workspace,
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
        budget,
    )?;
    let fields = budget.parse_fields(payload, 1)?;
    let bound = payload.len();
    budget.preflight_output(bound)?;
    budget.preflight_retained(bound)?;
    budget.preflight_scratch(bound)?;
    budget.preflight_allocations(fields.len().saturating_add(1))?;
    budget.allocations(fields.len().saturating_add(1))?;
    budget.retained(bound)?;
    budget.scratch(bound)?;
    let mut kept = Vec::new();
    kept.try_reserve_exact(fields.len())
        .map_err(|_| SymbolError::Allocation {
            amount: fields.len(),
        })?;
    for field in fields
        .iter()
        .filter(|field| field.number() == STYLESHEET_STYLES_FIELD)
    {
        let reference = field
            .payload(payload)
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        let reference_fields = budget.parse_fields(reference, 2)?;
        let identifier = reference_fields
            .iter()
            .rev()
            .find(|field| field.number() == REFERENCE_IDENTIFIER_FIELD)
            .ok_or(SymbolError::Invalid(
                "a stylesheet reference has no identifier",
            ))?;
        let bytes = identifier
            .payload(reference)
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        let (identifier, width) = litchi_iwa_common::decode_varint_from_bytes(bytes)
            .map_err(|_| SymbolError::Invalid("a stylesheet reference is malformed"))?;
        if width != bytes.len() {
            return Err(SymbolError::Invalid(
                "a stylesheet reference is noncanonical",
            ));
        }
        if styles.contains(&identifier) {
            budget.references(1)?;
        } else {
            kept.push(reference.to_vec());
        }
    }
    let data =
        wire::rewrite_repeated_length_delimited_fields(payload, STYLESHEET_STYLES_FIELD, &kept)
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    budget.output(data.len())?;
    workspace
        .replace_payload(stylesheet, message, data)
        .map_err(SymbolError::from)?;
    let info = workspace
        .object_mut(stylesheet)
        .map_err(SymbolError::from)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(SymbolError::Invalid("a stylesheet payload has no metadata"))?;
    for &style in styles {
        let listed = info
            .object_references
            .iter()
            .filter(|identifier| **identifier == style)
            .count();
        if listed > 1 || (!local && listed != 1) {
            return Err(SymbolError::Invalid(
                "the stylesheet metadata has invalid style ownership",
            ));
        }
        info.object_references
            .retain(|identifier| *identifier != style);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| *identifier != style);
        }
    }
    budget.registry_entries(styles.len())?;
    if let (Some(owner), Some(sheet)) = (owner_component, sheet_component)
        && owner != sheet
    {
        budget.admit_registry_inspection(workspace)?;
        for &style in styles {
            budget.references(1)?;
            workspace
                .remove_external_reference(sheet, owner, style)
                .map_err(SymbolError::from)?;
        }
    }
    Ok(())
}

fn bounded_container_of(
    archive: &Archive,
    budget: &mut SymbolBudget,
) -> SymbolResult<Option<(u64, usize, Vec<u64>)>> {
    let mut found = None;
    for object in &archive.objects {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ != OBJECT_CONTAINER_MESSAGE_TYPE {
                continue;
            }
            let identifier = object.archive_info.identifier.ok_or(SymbolError::Invalid(
                "an object container has no identifier",
            ))?;
            if found
                .replace((identifier, index, message.data.as_slice()))
                .is_some()
            {
                return Err(SymbolError::Invalid(
                    "a component holds several object containers",
                ));
            }
        }
    }
    let Some((identifier, index, payload)) = found else {
        return Ok(None);
    };
    let fields = budget.parse_fields(payload, 1)?;
    let package = fields
        .iter()
        .rev()
        .find(|field| field.number() == CONTAINER_IDENTIFIER_FIELD)
        .ok_or(SymbolError::Invalid(
            "an object container has no package identifier",
        ))?;
    let package_bytes = package
        .payload(payload)
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    let (package, _) = litchi_iwa_common::decode_varint_from_bytes(package_bytes)
        .map_err(|_| SymbolError::Invalid("an object container package is malformed"))?;
    if package != PRIMARY_PACKAGE_IDENTIFIER {
        return Err(SymbolError::Invalid(
            "an object container is not in the primary package",
        ));
    }
    let count = fields
        .iter()
        .filter(|field| field.number() == CONTAINER_OBJECTS_FIELD)
        .count();
    budget.reserve::<u64>(count, true)?;
    let mut members = Vec::new();
    members
        .try_reserve_exact(count)
        .map_err(|_| SymbolError::Allocation { amount: count })?;
    for field in fields
        .iter()
        .filter(|field| field.number() == CONTAINER_OBJECTS_FIELD)
    {
        let reference = field
            .payload(payload)
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        let reference_fields = budget.parse_fields(reference, 2)?;
        let identifier_field = reference_fields
            .iter()
            .rev()
            .find(|field| field.number() == REFERENCE_IDENTIFIER_FIELD)
            .ok_or(SymbolError::Invalid(
                "an object container reference is missing",
            ))?;
        let bytes = identifier_field
            .payload(reference)
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
        let (member, width) = litchi_iwa_common::decode_varint_from_bytes(bytes)
            .map_err(|_| SymbolError::Invalid("an object container member is malformed"))?;
        if width != bytes.len()
            || member == 0
            || members.contains(&member)
            || archive.object(member).is_none()
        {
            return Err(SymbolError::Invalid(
                "an object container member is invalid",
            ));
        }
        members.push(member);
        budget.references(1)?;
    }
    Ok(Some((identifier, index, members)))
}

fn bounded_insert_container(
    workspace: &mut Workspace<'_>,
    chart_member: &str,
    container: Container,
    objects: Vec<ArchiveObject>,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    let identifiers_count = objects.len();
    budget.identifiers(identifiers_count.saturating_add(1))?;
    budget.registry_entries(identifiers_count.saturating_add(2))?;
    let mut identifiers = Vec::new();
    budget.reserve_vec(&mut identifiers, identifiers_count, true)?;
    for object in &objects {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(SymbolError::Invalid("a container member has no identifier"))?;
        identifiers.push(identifier);
    }
    let container_data = bounded_encode_container(&identifiers, budget)?;
    budget.planned_new_object(container.component);
    let mut container_object = ArchiveObject::new(
        container.component,
        vec![RawMessage {
            type_: OBJECT_CONTAINER_MESSAGE_TYPE,
            data: container_data,
        }],
        workspace.limits(),
    )
    .map_err(|error| SymbolError::Graph(Error::Archive(error)))?;
    container_object.archive_info.message_infos[0].versions =
        OBJECT_CONTAINER_VERSION.to_vec().into();
    let mut archive_objects = Vec::new();
    budget.reserve_vec(&mut archive_objects, objects.len().saturating_add(1), true)?;
    archive_objects.push(container_object);
    archive_objects.extend(objects);
    let archive = Archive {
        objects: archive_objects,
    };
    let anchor = if workspace.has_member(DOCUMENT_METADATA_MEMBER) {
        DOCUMENT_METADATA_MEMBER
    } else {
        METADATA_MEMBER
    };
    let member_name = container.name.clone();
    budget.retained(member_name.len().saturating_add(container.locator.len()))?;
    workspace
        .create_member(&member_name, archive, Some(anchor))
        .map_err(SymbolError::from)?;
    if !workspace.has_registry() {
        return Ok(());
    }
    budget.admit_registry_inspection(workspace)?;
    let mut registration = Vec::new();
    registration
        .try_reserve_exact(64)
        .map_err(|_| SymbolError::Allocation { amount: 64 })?;
    push_varint_field(&mut registration, 1, container.component);
    push_message_field(&mut registration, 2, container.locator.as_bytes());
    push_packed_field(&mut registration, 4, &OBJECT_CONTAINER_VERSION);
    push_packed_field(&mut registration, 5, &OBJECT_CONTAINER_VERSION);
    push_varint_field(&mut registration, 10, 0);
    if let Some(token) = workspace.save_token().map_err(SymbolError::from)? {
        push_varint_field(&mut registration, 12, token);
    }
    if workspace
        .has_data_metadata_map()
        .map_err(SymbolError::from)?
    {
        push_varint_field(&mut registration, 21, PRIMARY_PACKAGE_IDENTIFIER);
    }
    workspace
        .add_component(container.component, &registration)
        .map_err(SymbolError::from)?;
    budget.admit_registry_inspection(workspace)?;
    workspace
        .add_object_uuids(container.component, &identifiers)
        .map_err(SymbolError::from)?;
    let source = workspace
        .component_for_member(chart_member)
        .map_err(SymbolError::from)?
        .ok_or(SymbolError::Invalid(
            "the chart component is not registered",
        ))?;
    for identifier in identifiers {
        budget.references(1)?;
        workspace
            .add_external_reference(source, container.component, identifier)
            .map_err(SymbolError::from)?;
    }
    Ok(())
}

fn bounded_encode_container(
    identifiers: &[u64],
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<u8>> {
    let container_size = 1usize
        .saturating_add(1)
        .saturating_add(litchi_iwa_common::varint::encoded_len(
            PRIMARY_PACKAGE_IDENTIFIER,
        ))
        .saturating_add(
            identifiers
                .iter()
                .map(|identifier| {
                    let reference_size =
                        1usize.saturating_add(litchi_iwa_common::varint::encoded_len(*identifier));
                    1usize
                        .saturating_add(litchi_iwa_common::varint::encoded_len(
                            reference_size as u64,
                        ))
                        .saturating_add(reference_size)
                })
                .sum::<usize>(),
        );
    budget.preflight_output(container_size)?;
    budget.preflight_retained(container_size)?;
    budget.preflight_scratch(container_size)?;
    budget.preflight_allocations(1)?;
    budget.allocations(1)?;
    budget.retained(container_size)?;
    budget.scratch(container_size)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(container_size)
        .map_err(|_| SymbolError::Allocation {
            amount: container_size,
        })?;
    push_varint_field(
        &mut output,
        CONTAINER_IDENTIFIER_FIELD,
        PRIMARY_PACKAGE_IDENTIFIER,
    );
    for identifier in identifiers {
        let mut reference = Vec::new();
        reference
            .try_reserve_exact(1 + litchi_iwa_common::varint::encoded_len(*identifier))
            .map_err(|_| SymbolError::Allocation { amount: 1 })?;
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, *identifier);
        push_message_field(&mut output, CONTAINER_OBJECTS_FIELD, &reference);
    }
    budget.output(output.len())?;
    Ok(output)
}

fn bounded_remove_from_container(
    workspace: &mut Workspace<'_>,
    location: Location,
    objects: &[u64],
    budget: &mut SymbolBudget,
) -> SymbolResult<Option<u64>> {
    let member = workspace.member(location).to_owned();
    let (container, message, mut members) = bounded_container_of(
        workspace.archive_at(location).map_err(SymbolError::from)?,
        budget,
    )?
    .ok_or(SymbolError::Invalid(
        "a series override container is not an object container",
    ))?;
    for (index, &object) in objects.iter().enumerate() {
        if object == 0 || objects[..index].contains(&object) || !members.contains(&object) {
            return Err(SymbolError::Invalid(
                "an object container does not own every removed override",
            ));
        }
    }
    members.retain(|identifier| !objects.contains(identifier));
    let component = workspace
        .component_for_member(&member)
        .map_err(SymbolError::from)?;
    if component.is_some_and(|component| component != container) {
        return Err(SymbolError::Invalid(
            "an object container disagrees with its component",
        ));
    }
    let encoded = if members.is_empty() {
        None
    } else {
        Some(bounded_encode_container(&members, budget)?)
    };
    if members.is_empty() {
        if let Some(component) = component {
            let locator = workspace
                .component_locator(component)
                .map_err(SymbolError::from)?
                .ok_or(SymbolError::Invalid(
                    "an object container component has no locator",
                ))?;
            workspace
                .remove_component(component, &locator)
                .map_err(SymbolError::from)?;
        }
        workspace
            .delete_member(location)
            .map_err(SymbolError::from)?;
        budget.registry_entries(1)?;
        return Ok(Some(container));
    }
    if let Some(component) = component {
        budget.admit_registry_inspection(workspace)?;
        let registered = workspace
            .component_uuid_objects(component)
            .map_err(SymbolError::from)?;
        let mut bound = Vec::new();
        budget.reserve_vec(&mut bound, objects.len(), true)?;
        for &identifier in objects {
            if registered.contains(&identifier) {
                bound.push(identifier);
            }
        }
        workspace
            .remove_object_uuids(component, &bound)
            .map_err(SymbolError::from)?;
        for &identifier in objects {
            budget.references(1)?;
            workspace
                .remove_external_references_to(component, identifier)
                .map_err(SymbolError::from)?;
        }
    }
    for &identifier in objects {
        workspace.remove(identifier).map_err(SymbolError::from)?;
    }
    workspace
        .replace_payload(
            container,
            message,
            encoded.expect("nonempty container encoding"),
        )
        .map_err(SymbolError::from)?;
    budget.registry_entries(objects.len())?;
    Ok(None)
}

fn bounded_allocate(next: &mut u64, budget: &mut SymbolBudget) -> SymbolResult<u64> {
    let identifier = *next;
    *next = next.checked_add(1).ok_or(SymbolError::IdentifierOverflow)?;
    budget.identifiers(1)?;
    Ok(identifier)
}

fn bounded_reserve_container(
    workspace: &Workspace<'_>,
    next: &mut u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<Container> {
    let component = bounded_allocate(next, budget)?;
    let mut suffix: Option<u64> = None;
    loop {
        let suffix_bytes = suffix.map_or(0, |value| value.to_string().len().saturating_add(1));
        budget.preflight_allocations(2)?;
        budget.preflight_retained(
            OBJECT_CONTAINER_LOCATOR
                .len()
                .saturating_add(suffix_bytes)
                .saturating_add(16),
        )?;
        budget.allocations(2)?;
        budget.retained(
            OBJECT_CONTAINER_LOCATOR
                .len()
                .saturating_add(suffix_bytes)
                .saturating_add(16),
        )?;
        let locator = suffix.map_or_else(
            || OBJECT_CONTAINER_LOCATOR.to_owned(),
            |value| format!("{OBJECT_CONTAINER_LOCATOR}-{value}"),
        );
        let name = format!("Index/{locator}.iwa");
        if !workspace.has_member(&name)
            && workspace
                .component_for_member(&name)
                .map_err(SymbolError::from)?
                .is_none()
        {
            return Ok(Container {
                component,
                name,
                locator,
            });
        }
        suffix = Some(match suffix {
            None => component,
            Some(value) => value
                .checked_add(1)
                .ok_or(SymbolError::IdentifierOverflow)?,
        });
    }
}

fn bounded_new_override(
    identifier: u64,
    data: Vec<u8>,
    workspace: &Workspace<'_>,
    budget: &mut SymbolBudget,
) -> SymbolResult<ArchiveObject> {
    budget.preflight_allocations(2)?;
    budget.preflight_retained(data.len())?;
    budget.preflight_scratch(data.len())?;
    budget.allocations(2)?;
    budget.retained(data.len())?;
    budget.scratch(data.len())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SERIES_NON_STYLE_MESSAGE_TYPE,
            data,
        }],
        workspace.limits(),
    )
    .map_err(|error| SymbolError::Graph(Error::Archive(error)))?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec().into();
    Ok(object)
}

fn append_location_group(
    groups: &mut Vec<(Location, Vec<u64>)>,
    location: Location,
    identifier: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    if let Some((_, identifiers)) = groups
        .iter_mut()
        .find(|(candidate, _)| *candidate == location)
    {
        budget.preflight_allocations(1)?;
        budget.preflight_retained(size_of::<u64>())?;
        budget.allocations(1)?;
        budget.retained(size_of::<u64>())?;
        identifiers
            .try_reserve_exact(1)
            .map_err(|_| SymbolError::Allocation { amount: 1 })?;
        identifiers.push(identifier);
        return Ok(());
    }
    budget.preflight_allocations(2)?;
    budget
        .preflight_retained(size_of::<(Location, Vec<u64>)>().saturating_add(size_of::<u64>()))?;
    budget.allocations(2)?;
    budget.retained(size_of::<(Location, Vec<u64>)>().saturating_add(size_of::<u64>()))?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(1)
        .map_err(|_| SymbolError::Allocation { amount: 1 })?;
    identifiers.push(identifier);
    groups.push((location, identifiers));
    Ok(())
}

fn append_component_group(
    groups: &mut Vec<(u64, Vec<u64>)>,
    component: u64,
    identifier: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    if let Some((_, identifiers)) = groups
        .iter_mut()
        .find(|(candidate, _)| *candidate == component)
    {
        budget.preflight_allocations(1)?;
        budget.preflight_retained(size_of::<u64>())?;
        budget.allocations(1)?;
        budget.retained(size_of::<u64>())?;
        identifiers
            .try_reserve_exact(1)
            .map_err(|_| SymbolError::Allocation { amount: 1 })?;
        identifiers.push(identifier);
        return Ok(());
    }
    budget.preflight_allocations(2)?;
    budget.preflight_retained(size_of::<(u64, Vec<u64>)>().saturating_add(size_of::<u64>()))?;
    budget.allocations(2)?;
    budget.retained(size_of::<(u64, Vec<u64>)>().saturating_add(size_of::<u64>()))?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(1)
        .map_err(|_| SymbolError::Allocation { amount: 1 })?;
    identifiers.push(identifier);
    groups.push((component, identifiers));
    Ok(())
}

fn bounded_set_values<P: BoundedProperty>(
    workspace: &mut Workspace<'_>,
    chart: u64,
    property: &P,
    expected: &[P::Value],
    budget: &mut SymbolBudget,
) -> SymbolResult<bool> {
    if expected.is_empty() {
        return Err(SymbolError::Invalid("a chart series write names no series"));
    }
    let base = property.base();
    let default = property.default_value();
    let graph = bounded_series_graph(workspace, chart, expected.len(), budget)?;
    let stylesheet = bounded_chart_stylesheet(workspace, chart, budget)?;
    let current = bounded_read_values(workspace, &graph, property, budget)?;
    if current == expected {
        return Ok(false);
    }
    let chart_location = workspace.locate(chart).map_err(SymbolError::from)?;
    let chart_member = workspace.member(chart_location).to_owned();
    let styled_empty = bounded_canonical_empty(Some(stylesheet), budget)?;
    let unstyled_empty = bounded_canonical_empty(None, budget)?;
    let unstyled_base = bounded_override_with_parent(None, budget)?;
    let new_base = match base {
        Base::Styled => styled_empty.as_slice(),
        Base::Unstyled => unstyled_base.as_slice(),
    };
    let mut next = workspace
        .next_object_identifier()
        .map_err(SymbolError::from)?;
    let mut final_ids = Vec::new();
    budget.reserve_vec(&mut final_ids, graph.slots.len(), true)?;
    final_ids.extend(
        graph
            .slots
            .iter()
            .map(|slot| slot.as_ref().map(|slot| slot.object)),
    );
    let mut updates = Vec::new();
    let mut removals = Vec::new();
    let mut styled_creations = Vec::new();
    let mut chart_created = Vec::new();
    let mut unstyled_creations = Vec::new();
    let mut created = Vec::new();
    let mut registrations: Vec<(Location, Vec<u64>)> = Vec::new();
    let mut unregistrations: Vec<(Location, Vec<u64>)> = Vec::new();
    budget.reserve_vec(&mut updates, expected.len(), true)?;
    budget.reserve_vec(&mut removals, expected.len(), true)?;
    budget.reserve_vec(&mut styled_creations, expected.len(), true)?;
    budget.reserve_vec(&mut chart_created, expected.len(), true)?;
    budget.reserve_vec(&mut unstyled_creations, expected.len(), true)?;
    budget.reserve_vec(&mut created, expected.len(), true)?;
    budget.reserve_vec(&mut registrations, expected.len(), true)?;
    budget.reserve_vec(&mut unregistrations, expected.len(), true)?;
    let mut container = None;
    for (index, ((slot, value), replacement)) in
        graph.slots.iter().zip(&current).zip(expected).enumerate()
    {
        budget.work(1)?;
        if value == replacement {
            continue;
        }
        if let Some(slot) = slot {
            let payload = bounded_slot_payload(workspace, slot)?;
            let mut patched = property.patch_bounded(payload, replacement, budget)?;
            let registered = bounded_stylesheet_of(payload, budget)?;
            if let Some(registered) = registered
                && registered != stylesheet
            {
                return Err(SymbolError::Invalid(
                    "a series override belongs to another stylesheet",
                ));
            }
            let needs_registration =
                base == Base::Styled && replacement != &default && registered.is_none();
            if needs_registration {
                patched = bounded_set_stylesheet(&patched, stylesheet, budget)?;
            }
            let is_empty =
                patched == styled_empty || patched == unstyled_empty || patched == unstyled_base;
            let owners = bounded_owner_count(workspace, slot.object, budget)?;
            if owners == 0 {
                return Err(SymbolError::Invalid("a series override has no owner"));
            }
            if owners > 1 {
                if is_empty {
                    final_ids[index] = None;
                    continue;
                }
                let allocation = if registered.is_some() {
                    Base::Styled
                } else {
                    base
                };
                if allocation == Base::Unstyled && container.is_none() {
                    let reserved = bounded_reserve_container(workspace, &mut next, budget)?;
                    budget.planned_member(true, &reserved.name)?;
                    container = Some(reserved);
                }
                let identifier = bounded_allocate(&mut next, budget)?;
                budget.planned_new_object(identifier);
                final_ids[index] = Some(identifier);
                let object = bounded_new_override(identifier, patched, workspace, budget)?;
                match allocation {
                    Base::Styled => {
                        append_location_group(
                            &mut registrations,
                            chart_location,
                            identifier,
                            budget,
                        )?;
                        chart_created.push(identifier);
                        styled_creations.push(object);
                    },
                    Base::Unstyled => unstyled_creations.push(object),
                }
                created.push(identifier);
            } else if is_empty {
                final_ids[index] = None;
                if registered.is_some() {
                    append_location_group(
                        &mut unregistrations,
                        slot.location,
                        slot.object,
                        budget,
                    )?;
                }
                budget.planned_removed_object(slot.object);
                removals.push(slot.clone());
            } else {
                if needs_registration {
                    append_location_group(&mut registrations, slot.location, slot.object, budget)?;
                }
                updates.push((slot.clone(), patched));
            }
        } else if replacement != &default {
            let mut data = property.patch_bounded(new_base, replacement, budget)?;
            if data.as_slice() == new_base {
                return Err(SymbolError::Invalid(
                    "a series write produced no native override",
                ));
            }
            if base == Base::Unstyled {
                budget.preflight_retained(CAPABILITY_FIELDS.len().saturating_mul(3))?;
                budget.retained(CAPABILITY_FIELDS.len().saturating_mul(3))?;
                data.try_reserve(CAPABILITY_FIELDS.len().saturating_mul(3))
                    .map_err(|_| SymbolError::Allocation {
                        amount: CAPABILITY_FIELDS.len(),
                    })?;
                append_capabilities(&mut data);
            }
            if base == Base::Unstyled && container.is_none() {
                let reserved = bounded_reserve_container(workspace, &mut next, budget)?;
                budget.planned_member(true, &reserved.name)?;
                container = Some(reserved);
            }
            let identifier = bounded_allocate(&mut next, budget)?;
            budget.planned_new_object(identifier);
            final_ids[index] = Some(identifier);
            let object = bounded_new_override(identifier, data, workspace, budget)?;
            match base {
                Base::Styled => {
                    append_location_group(&mut registrations, chart_location, identifier, budget)?;
                    chart_created.push(identifier);
                    styled_creations.push(object);
                },
                Base::Unstyled => unstyled_creations.push(object),
            }
            created.push(identifier);
        }
    }

    for (slot, data) in updates {
        workspace
            .replace_payload(slot.object, slot.message, data)
            .map_err(SymbolError::from)?;
    }
    for (location, styles) in &unregistrations {
        bounded_unregister_styles(workspace, stylesheet, *location, styles, budget)?;
    }
    let mut removal_members = Vec::new();
    budget.reserve_vec(&mut removal_members, removals.len(), true)?;
    for slot in &removals {
        let member = workspace.member(slot.location).to_owned();
        budget.retained(member.len())?;
        removal_members.push((slot.object, member));
    }
    let mut contained: Vec<(Location, Vec<u64>)> = Vec::new();
    budget.reserve_vec(&mut contained, removals.len(), true)?;
    for slot in &removals {
        if bounded_container_of(
            workspace
                .archive_at(slot.location)
                .map_err(SymbolError::from)?,
            budget,
        )?
        .is_some()
        {
            append_location_group(&mut contained, slot.location, slot.object, budget)?;
        } else {
            workspace.remove(slot.object).map_err(SymbolError::from)?;
        }
    }
    let mut released_containers = Vec::new();
    budget.reserve_vec(&mut released_containers, removals.len(), true)?;
    for (location, objects) in contained {
        let member = workspace.member(location).to_owned();
        if let Some(identifier) =
            bounded_remove_from_container(workspace, location, &objects, budget)?
        {
            budget.planned_member(false, &member)?;
            budget.planned_removed_object(identifier);
            released_containers.push(identifier);
        }
    }
    for object in styled_creations {
        workspace
            .insert(chart_location, object)
            .map_err(SymbolError::from)?;
    }
    if !unstyled_creations.is_empty() {
        let container = container.take().ok_or(SymbolError::Invalid(
            "an unstyled override lost its object container",
        ))?;
        bounded_insert_container(
            workspace,
            &chart_member,
            container,
            unstyled_creations,
            budget,
        )?;
    }
    for (location, styles) in &registrations {
        bounded_register_styles(workspace, stylesheet, *location, styles, budget)?;
    }
    let previous_count = graph.slots.iter().flatten().count();
    let mut previous = Vec::new();
    budget.reserve_vec(&mut previous, previous_count, true)?;
    previous.extend(graph.slots.iter().flatten().map(|slot| slot.object));
    bounded_patch_chart_references(
        workspace,
        chart,
        graph.chart_message,
        &previous,
        &final_ids,
        budget,
    )?;
    let mut unbound: Vec<(u64, Vec<u64>)> = Vec::new();
    budget.reserve_vec(&mut unbound, removals.len(), true)?;
    for (object, member) in &removal_members {
        if let Some(component) = workspace
            .component_for_member(member)
            .map_err(SymbolError::from)?
        {
            budget.admit_registry_inspection(workspace)?;
            if workspace
                .component_uuid_objects(component)
                .map_err(SymbolError::from)?
                .contains(object)
            {
                append_component_group(&mut unbound, component, *object, budget)?;
            }
        }
    }
    for (component, objects) in unbound {
        workspace
            .remove_object_uuids(component, &objects)
            .map_err(SymbolError::from)?;
    }
    if !chart_created.is_empty()
        && let Some(component) = workspace
            .component_for_member(&chart_member)
            .map_err(SymbolError::from)?
    {
        budget.admit_registry_inspection(workspace)?;
        budget.registry_entries(chart_created.len())?;
        workspace
            .add_object_uuids(component, &chart_created)
            .map_err(SymbolError::from)?;
    }
    if let Some(last) = created.last().copied() {
        budget.admit_registry_inspection(workspace)?;
        budget.registry_entries(1)?;
        workspace
            .set_last_object_identifier(last)
            .map_err(SymbolError::from)?;
    }
    budget.reserve::<u64>(
        removals.len().saturating_add(released_containers.len()),
        true,
    )?;
    let mut released = Vec::new();
    budget.reserve_vec(
        &mut released,
        removals.len().saturating_add(released_containers.len()),
        true,
    )?;
    released.extend(removals.iter().map(|slot| slot.object));
    released.extend(released_containers);
    if !released.is_empty() {
        budget.identifiers(released.len())?;
        budget.admit_registry_inspection(workspace)?;
        budget.registry_entries(1)?;
        workspace
            .release_identifier_suffix(&released)
            .map_err(SymbolError::from)?;
    }
    let graph = bounded_series_graph(workspace, chart, expected.len(), budget)?;
    let actual = bounded_read_values(workspace, &graph, property, budget)?;
    if actual != expected {
        return Err(SymbolError::Invalid(
            "a chart series write does not read back",
        ));
    }
    Ok(true)
}

fn bounded_style_owner_count(
    workspace: &Workspace<'_>,
    style: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<usize> {
    let mut owners = 0usize;
    for location in workspace.locations_iter() {
        let archive = workspace.archive_at(location).map_err(SymbolError::from)?;
        for candidate in &archive.objects {
            for message in &candidate.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let references = bounded_references(&message.data, budget)?;
                let private = references.series_private_styles().is_some_and(|sparse| {
                    sparse
                        .entries()
                        .iter()
                        .any(|(_, identifier)| *identifier == style)
                });
                if private || references.series_theme_styles().contains(&style) {
                    owners = owners.saturating_add(1);
                }
            }
        }
    }
    Ok(owners)
}

fn bounded_set_style_values<P>(
    workspace: &mut Workspace<'_>,
    chart: u64,
    property: &P,
    expected: &[Option<litchi_iwa_common::chart::series::SymbolSize>],
    budget: &mut SymbolBudget,
) -> SymbolResult<bool>
where
    P: BoundedStyleProperty<Item = litchi_iwa_common::chart::series::SymbolSize>,
{
    let styles = bounded_style_slots(workspace, chart, expected.len(), budget)?;
    let mut changed = false;
    for (&style, replacement) in styles.iter().zip(expected) {
        if bounded_inherited(workspace, style, property, budget)? == *replacement {
            continue;
        }
        if bounded_style_owner_count(workspace, style, budget)? != 1 {
            return Err(SymbolError::Invalid(
                "a series style is shared by several charts",
            ));
        }
        let (message, payload) = bounded_payload(
            workspace,
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
            budget,
        )?;
        let data = property.patch_bounded(payload, replacement.as_ref(), budget)?;
        workspace
            .replace_payload(style, message, data)
            .map_err(SymbolError::from)?;
        changed = true;
    }
    Ok(changed)
}

/// Write all data symbols under one shared explicit residual profile.
///
/// The workspace is copy-on-write. A failure after a staged mark or style
/// mutation can leave that private workspace changed, but no source bytes are
/// published; callers must discard the failed workspace and retain the
/// immutable source snapshot.
pub fn set_symbols(
    workspace: &mut Workspace<'_>,
    chart: u64,
    family: litchi_iwa_common::chart::series::SymbolFamily,
    expected: &[Option<litchi_iwa_common::chart::series::Symbol>],
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    if expected.is_empty() {
        return Err(SymbolError::Invalid("a chart series write names no series"));
    }
    let mut budget = SymbolBudget::new(limits);
    budget.work(workspace.source_work())?;
    let scope = bounded_mutation_scope(
        workspace,
        chart,
        &[
            family.show_field(),
            family.type_field(),
            family.size_field(),
        ],
        expected.len(),
        &mut budget,
    )?;
    budget.set_scope(scope);
    budget.reserve::<property::Mark>(expected.len(), true)?;
    budget.reserve::<Option<litchi_iwa_common::chart::series::SymbolSize>>(expected.len(), true)?;
    let mut marks = Vec::new();
    marks
        .try_reserve_exact(expected.len())
        .map_err(|_| SymbolError::Allocation {
            amount: expected.len(),
        })?;
    let mut sizes = Vec::new();
    sizes
        .try_reserve_exact(expected.len())
        .map_err(|_| SymbolError::Allocation {
            amount: expected.len(),
        })?;
    for symbol in expected {
        marks.push(
            symbol.map_or(property::Mark::HIDDEN, |symbol| property::Mark {
                visible: true,
                shape: symbol.shape,
            }),
        );
        sizes.push(symbol.and_then(|symbol| symbol.size));
    }
    let marked = bounded_set_values(
        workspace,
        chart,
        &property::SymbolMarks(family),
        &marks,
        &mut budget,
    )?;
    let sized = bounded_set_style_values(
        workspace,
        chart,
        &property::SymbolSizes(family),
        &sizes,
        &mut budget,
    )?;
    let graph = bounded_series_graph(workspace, chart, expected.len(), &mut budget)?;
    let actual_marks = bounded_read_values(
        workspace,
        &graph,
        &property::SymbolMarks(family),
        &mut budget,
    )?;
    let actual_sizes = bounded_read_sizes(
        workspace,
        chart,
        expected.len(),
        &property::SymbolSizes(family),
        &mut budget,
    )?;
    let mut actual = Vec::new();
    budget.reserve::<Option<litchi_iwa_common::chart::series::Symbol>>(expected.len(), false)?;
    actual
        .try_reserve_exact(expected.len())
        .map_err(|_| SymbolError::Allocation {
            amount: expected.len(),
        })?;
    for (mark, size) in actual_marks.into_iter().zip(actual_sizes) {
        actual.push(
            mark.visible
                .then_some(litchi_iwa_common::chart::series::Symbol {
                    shape: mark.shape,
                    size,
                }),
        );
    }
    if actual != expected {
        return Err(SymbolError::Invalid(
            "a data-symbol write does not read back",
        ));
    }
    Ok((marked || sized, budget.report(marked || sized)))
}

fn bounded_set_stylesheet(
    payload: &[u8],
    stylesheet: u64,
    budget: &mut SymbolBudget,
) -> SymbolResult<Vec<u8>> {
    if bounded_stylesheet_of(payload, budget)? == Some(stylesheet) {
        budget.preflight_retained(payload.len())?;
        budget.retained(payload.len())?;
        let mut copy = Vec::new();
        copy.try_reserve_exact(payload.len())
            .map_err(|_| SymbolError::Allocation {
                amount: payload.len(),
            })?;
        copy.extend_from_slice(payload);
        return Ok(copy);
    }
    let parent = bounded_style_parent_payload(Some(stylesheet), budget)?;
    let bound = payload.len().saturating_add(parent.len()).saturating_add(8);
    budget.preflight_output(bound)?;
    budget.preflight_retained(bound)?;
    budget.preflight_scratch(bound)?;
    budget.preflight_allocations(1)?;
    budget.allocations(1)?;
    budget.retained(bound)?;
    budget.scratch(bound)?;
    let patched = wire::patch_length_delimited_field(
        payload,
        STYLE_PARENT_FIELD,
        true,
        Some(parent.as_slice()),
    )
    .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    budget.output(patched.len())?;
    Ok(patched)
}

/// Read all data symbols under one explicit residual profile.
pub fn symbols(
    workspace: &Workspace<'_>,
    chart: u64,
    family: litchi_iwa_common::chart::series::SymbolFamily,
    series_count: usize,
    limits: Limits,
) -> SymbolResult<(
    Vec<Option<litchi_iwa_common::chart::series::Symbol>>,
    Report,
)> {
    let mut budget = SymbolBudget::new(limits);
    budget.work(workspace.source_work())?;
    let scope = bounded_mutation_scope(
        workspace,
        chart,
        &[
            family.show_field(),
            family.type_field(),
            family.size_field(),
        ],
        series_count,
        &mut budget,
    )?;
    budget.set_scope(scope);
    let graph = bounded_series_graph(workspace, chart, series_count, &mut budget)?;
    let marks = bounded_read_values(
        workspace,
        &graph,
        &property::SymbolMarks(family),
        &mut budget,
    )?;
    let sizes = bounded_read_sizes(
        workspace,
        chart,
        series_count,
        &property::SymbolSizes(family),
        &mut budget,
    )?;
    budget.reserve::<Option<litchi_iwa_common::chart::series::Symbol>>(series_count, false)?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(series_count)
        .map_err(|_| SymbolError::Allocation {
            amount: series_count,
        })?;
    for (mark, size) in marks.into_iter().zip(sizes) {
        values.push(
            mark.visible
                .then_some(litchi_iwa_common::chart::series::Symbol {
                    shape: mark.shape,
                    size,
                }),
        );
    }
    Ok((values, budget.report(false)))
}

/// Extra output a property patch may emit beyond its source payload before
/// the bounded adapter refuses it: the largest generated value the series
/// properties write (a custom error-bar vector or a number format) fits.
const PROPERTY_PATCH_GROWTH: usize = 64 * 1024;

/// The bounded adapter of any [`Property`]: every payload a property reads or
/// patches is charged to the operation's residual profile, and the
/// extension fields a patch changes join the source-derived mutation scope
/// before publication.
struct Charged<'p, P>(&'p P);

impl<P: Property> Property for Charged<'_, P> {
    type Value = P::Value;

    fn base(&self) -> Base {
        self.0.base()
    }

    fn default_value(&self) -> P::Value {
        self.0.default_value()
    }

    fn read(&self, payload: &[u8]) -> Result<P::Value> {
        self.0.read(payload)
    }

    fn patch(&self, payload: &[u8], value: &P::Value) -> Result<Vec<u8>> {
        self.0.patch(payload, value)
    }
}

impl<P: Property> BoundedProperty for Charged<'_, P> {
    fn read_bounded(&self, payload: &[u8], budget: &mut SymbolBudget) -> SymbolResult<P::Value> {
        budget.source(payload.len())?;
        budget.work(payload.len())?;
        self.0.read(payload).map_err(SymbolError::from)
    }

    fn patch_bounded(
        &self,
        payload: &[u8],
        value: &P::Value,
        budget: &mut SymbolBudget,
    ) -> SymbolResult<Vec<u8>> {
        charged_patch(payload, budget, |payload| {
            self.0.patch(payload, value).map_err(SymbolError::from)
        })
    }
}

/// Run one payload patch under the residual profile and admit the
/// extension fields it changes into the mutation scope.
pub(crate) fn charged_patch(
    payload: &[u8],
    budget: &mut SymbolBudget,
    patch: impl FnOnce(&[u8]) -> SymbolResult<Vec<u8>>,
) -> SymbolResult<Vec<u8>> {
    let bound = payload.len().saturating_add(PROPERTY_PATCH_GROWTH);
    budget.preflight_output(bound)?;
    budget.preflight_retained(bound)?;
    budget.preflight_scratch(payload.len())?;
    budget.preflight_allocations(1)?;
    budget.source(payload.len())?;
    budget.work(payload.len())?;
    budget.allocations(1)?;
    budget.scratch(payload.len())?;
    let patched = patch(payload)?;
    budget.output(patched.len())?;
    budget.retained(patched.len())?;
    admit_changed_fields(payload, &patched, budget)?;
    Ok(patched)
}

/// The fields of the generated extension (field 10,000) of a style payload.
fn bounded_extension_fields<'a>(
    payload: &'a [u8],
    budget: &mut SymbolBudget,
) -> SymbolResult<(&'a [u8], Vec<wire::WireField>)> {
    let outer = budget.parse_fields(payload, 1)?;
    let mut extension: &[u8] = &[];
    let mut found = false;
    for field in outer
        .iter()
        .filter(|field| field.number() == codec::EXTENSION_FIELD)
    {
        if found {
            return Err(SymbolError::Invalid(
                "a style payload repeats its generated extension",
            ));
        }
        found = true;
        extension = field
            .payload(payload)
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))?;
    }
    let fields = budget.parse_fields(extension, 2)?;
    Ok((extension, fields))
}

/// Admit into the mutation scope every generated-extension field whose
/// stored bytes differ between `before` and `after`.
pub(crate) fn admit_changed_fields(
    before: &[u8],
    after: &[u8],
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    if budget.scope.is_none() || before == after {
        return Ok(());
    }
    let (source, source_fields) = bounded_extension_fields(before, budget)?;
    let (target, target_fields) = bounded_extension_fields(after, budget)?;
    let raw = |data: &[u8], field: &wire::WireField| {
        field
            .raw(data)
            .map(<[u8]>::to_vec)
            .map_err(|error| SymbolError::Graph(Error::Wire(error)))
    };
    let mut numbers = Vec::new();
    budget.reserve_vec(
        &mut numbers,
        source_fields.len().saturating_add(target_fields.len()),
        true,
    )?;
    for field in source_fields.iter().chain(&target_fields) {
        if !numbers.contains(&field.number()) {
            numbers.push(field.number());
        }
    }
    for number in numbers {
        budget.work(1)?;
        let left = source_fields
            .iter()
            .filter(|field| field.number() == number)
            .map(|field| raw(source, field))
            .collect::<SymbolResult<Vec<_>>>()?;
        let right = target_fields
            .iter()
            .filter(|field| field.number() == number)
            .map(|field| raw(target, field))
            .collect::<SymbolResult<Vec<_>>>()?;
        if left != right {
            budget.allow_field(number)?;
        }
    }
    Ok(())
}

/// The number of series the Format panel shows for a chart of `kind` whose
/// data has `components` series components: a scatter chart pairs X and Y
/// (and a bubble chart X, Y and size) into one series; `None` for an
/// incomplete scatter or bubble grid or a chart without series.
#[must_use]
pub const fn visible_series(
    kind: litchi_iwa_common::chart::kind::Kind,
    components: usize,
) -> Option<usize> {
    use litchi_iwa_common::chart::kind::Kind;
    let count = match kind {
        Kind::Scatter2d if components == 2 => 1,
        Kind::Bubble2d if components == 3 => 1,
        Kind::Scatter2d | Kind::Bubble2d => return None,
        Kind::MultiDataScatter2d | Kind::MultiDataBubble2d => 1,
        _ => components,
    };
    if count == 0 { None } else { Some(count) }
}

/// Read `property` for each of the `series_count` series of chart drawable
/// `chart` under one explicit residual profile; the report carries the
/// source-derived mutation scope of the chart's series graph.
///
/// # Errors
///
/// An invalid chart, sparse array or override, a value the property
/// rejects, or an exhausted limit.
pub fn read<P: Property>(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    property: &P,
    limits: Limits,
) -> SymbolResult<(Vec<P::Value>, Report)> {
    let mut budget = SymbolBudget::new(limits);
    budget.work(workspace.source_work())?;
    let scope = bounded_mutation_scope(workspace, chart, &[], series_count, &mut budget)?;
    budget.set_scope(scope);
    let graph = bounded_series_graph(workspace, chart, series_count, &mut budget)?;
    let values = bounded_read_values(workspace, &graph, &Charged(property), &mut budget)?;
    Ok((values, budget.report(false)))
}

/// Write `property` for every series (`expected.len()` is the series
/// count) under one explicit residual profile, allocating, copying,
/// patching or releasing the series overrides as [`set_values`] does.
/// Returns whether anything changed and the report with the mutation scope
/// (every object, member, message type and extension field the write may
/// change).
///
/// The workspace is copy-on-write: after a failure the caller discards it.
///
/// # Errors
///
/// An invalid chart or override, a value the property rejects, an
/// exhausted limit, or a write that does not read back.
pub fn write<P: Property>(
    workspace: &mut Workspace<'_>,
    chart: u64,
    property: &P,
    expected: &[P::Value],
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    if expected.is_empty() {
        return Err(SymbolError::Invalid("a chart series write names no series"));
    }
    let mut budget = SymbolBudget::new(limits);
    budget.work(workspace.source_work())?;
    let scope = bounded_mutation_scope(workspace, chart, &[], expected.len(), &mut budget)?;
    budget.set_scope(scope);
    let charged = Charged(property);
    let changed = bounded_set_values(workspace, chart, &charged, expected, &mut budget)?;
    let graph = bounded_series_graph(workspace, chart, expected.len(), &mut budget)?;
    if bounded_read_values(workspace, &graph, &charged, &mut budget)? != expected {
        return Err(SymbolError::Invalid(
            "a chart series write does not read back",
        ));
    }
    Ok((changed, budget.report(changed)))
}

#[cfg(test)]
mod symbol_tests {
    use super::*;
    use litchi_iwa_common::chart::series::{Symbol, SymbolFamily, SymbolShape, SymbolSize};
    use litchi_iwa_core::{Archive, ArchiveObject, Limits as CoreLimits, RawMessage};

    const FAMILY: SymbolFamily = SymbolFamily::Line2d;

    fn reference(identifier: u64) -> Vec<u8> {
        let mut value = Vec::new();
        push_varint_field(&mut value, REFERENCE_IDENTIFIER_FIELD, identifier);
        value
    }

    fn style_payload() -> Vec<u8> {
        let mut parent = Vec::new();
        let reference = reference(3);
        push_message_field(&mut parent, STYLE_STYLESHEET_FIELD, &reference);
        let mut payload = Vec::new();
        push_message_field(&mut payload, STYLE_PARENT_FIELD, &parent);
        push_varint_field(&mut payload, 4_242, 7);
        payload
    }

    fn override_payload() -> Vec<u8> {
        let mut parent = Vec::new();
        let reference = reference(3);
        push_message_field(&mut parent, STYLE_STYLESHEET_FIELD, &reference);
        let mut payload = Vec::new();
        push_message_field(&mut payload, STYLE_PARENT_FIELD, &parent);
        push_varint_field(&mut payload, 4_242, 9);
        payload
    }

    fn chart_payload() -> Vec<u8> {
        let mut extension = Vec::new();
        let style = reference(2);
        push_message_field(&mut extension, 17, &style);
        let mut sparse = Vec::new();
        push_varint_field(&mut sparse, 1, 1);
        let mut entry = Vec::new();
        push_varint_field(&mut entry, 1, 0);
        let override_reference = reference(4);
        push_message_field(&mut entry, 2, &override_reference);
        push_message_field(&mut sparse, 2, &entry);
        push_message_field(&mut extension, 19, &sparse);
        let mut payload = Vec::new();
        push_message_field(&mut payload, codec::EXTENSION_FIELD, &extension);
        payload
    }

    fn chart_payload_without_override() -> Vec<u8> {
        let mut extension = Vec::new();
        let style = reference(2);
        push_message_field(&mut extension, 17, &style);
        let mut payload = Vec::new();
        push_message_field(&mut payload, codec::EXTENSION_FIELD, &extension);
        payload
    }

    fn object(identifier: u64, message_type: u32, payload: Vec<u8>) -> ArchiveObject {
        ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: message_type,
                data: payload,
            }],
            CoreLimits::default(),
        )
        .expect("test archive object")
    }

    fn components() -> Vec<(String, Archive)> {
        vec![(
            "Index/Chart.iwa".to_owned(),
            Archive {
                objects: vec![
                    object(1, CHART_MESSAGE_TYPE, chart_payload()),
                    object(2, SERIES_STYLE_MESSAGE_TYPE, style_payload()),
                    object(3, STYLESHEET_MESSAGE_TYPE, Vec::new()),
                    object(4, SERIES_NON_STYLE_MESSAGE_TYPE, override_payload()),
                ],
            },
        )]
    }

    fn shared_components() -> Vec<(String, Archive)> {
        let mut components = components();
        components[0]
            .1
            .objects
            .push(object(5, CHART_MESSAGE_TYPE, chart_payload()));
        components
    }

    fn fresh_components() -> Vec<(String, Archive)> {
        vec![(
            "Index/Chart.iwa".to_owned(),
            Archive {
                objects: vec![
                    object(1, CHART_MESSAGE_TYPE, chart_payload_without_override()),
                    object(2, SERIES_STYLE_MESSAGE_TYPE, style_payload()),
                    object(3, STYLESHEET_MESSAGE_TYPE, Vec::new()),
                ],
            },
        )]
    }

    const fn limits(max_allocations: usize) -> Limits {
        Limits::new(
            256 * 1024,
            16 * 1024,
            256 * 1024,
            16,
            4 * 1024 * 1024,
            max_allocations,
            4 * 1024 * 1024,
            4 * 1024 * 1024,
            4 * 1024,
            4 * 1024,
            4 * 1024,
        )
    }

    const fn limits_with(fields: usize, references: usize) -> Limits {
        Limits::new(
            256 * 1024,
            fields,
            256 * 1024,
            16,
            4 * 1024 * 1024,
            4_096,
            4 * 1024 * 1024,
            4 * 1024 * 1024,
            references,
            4 * 1024,
            4 * 1024,
        )
    }

    const fn large_limits() -> Limits {
        Limits::new(
            usize::MAX,
            usize::MAX,
            usize::MAX,
            u32::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
        )
    }

    fn expected() -> Vec<Option<Symbol>> {
        vec![Some(Symbol {
            shape: SymbolShape::Diamond,
            size: Some(SymbolSize::new(9.0).expect("positive size")),
        })]
    }

    fn expected_without_size() -> Vec<Option<Symbol>> {
        vec![Some(Symbol {
            shape: SymbolShape::Diamond,
            size: None,
        })]
    }

    #[test]
    fn bounded_symbols_preserve_unknowns_and_read_back() {
        let components = components();
        let mut workspace = Workspace::new(&components, CoreLimits::default()).expect("workspace");
        let values = expected();
        let (changed, report) = set_symbols(&mut workspace, 1, FAMILY, &values, limits(4_096))
            .expect("bounded symbol write");
        assert!(changed);
        assert!(report.changed());
        assert!(report.fields() > 0);
        assert!(report.references() > 0);
        let scope = report.scope().expect("source-derived mutation scope");
        assert_eq!(scope.selected_chart(), 1);
        assert!(scope.allowed_objects().contains(&4));
        assert!(
            scope
                .allowed_message_types()
                .contains(&SERIES_STYLE_MESSAGE_TYPE)
        );
        assert!(scope.allowed_fields().contains(&codec::EXTENSION_FIELD));
        let override_payload = workspace
            .object(4)
            .expect("override")
            .messages
            .iter()
            .find(|message| message.type_ == SERIES_NON_STYLE_MESSAGE_TYPE)
            .expect("override payload")
            .data
            .as_slice();
        assert!(
            wire::parse_wire_fields(override_payload, WireLimits::DEFAULT)
                .expect("override fields")
                .iter()
                .any(|field| field.number() == 4_242)
        );
        let style_payload = workspace
            .object(2)
            .expect("style")
            .messages
            .iter()
            .find(|message| message.type_ == SERIES_STYLE_MESSAGE_TYPE)
            .expect("style payload")
            .data
            .as_slice();
        assert!(
            wire::parse_wire_fields(style_payload, WireLimits::DEFAULT)
                .expect("style fields")
                .iter()
                .any(|field| field.number() == 4_242)
        );
    }

    #[test]
    fn zero_allocation_refuses_before_workspace_mutation() {
        let components = components();
        let mut workspace = Workspace::new(&components, CoreLimits::default()).expect("workspace");
        let error = set_symbols(&mut workspace, 1, FAMILY, &expected(), limits(0))
            .expect_err("zero allocations must refuse");
        assert!(matches!(
            error,
            SymbolError::Limit {
                kind: LimitKind::Allocations,
                ..
            }
        ));
        assert!(workspace.finish().is_empty());
    }

    #[test]
    fn zero_wire_fields_and_references_are_typed_before_write() {
        for (limits, kind) in [
            (limits_with(0, 4_096), LimitKind::Fields),
            (limits_with(16 * 1024, 0), LimitKind::References),
        ] {
            let components = components();
            let mut workspace =
                Workspace::new(&components, CoreLimits::default()).expect("workspace");
            let error = set_symbols(&mut workspace, 1, FAMILY, &expected(), limits)
                .expect_err("zero ceiling must refuse");
            assert!(matches!(
                error,
                SymbolError::Limit { kind: actual, .. } if actual == kind
            ));
            assert!(workspace.finish().is_empty());
        }
    }

    #[test]
    fn large_profile_clamps_wire_and_codec_hard_axes_without_headroom() {
        let budget = SymbolBudget::new(large_limits());
        let wire = budget.wire_limits().expect("wire hard caps");
        assert_eq!(wire.max_input_bytes(), WireLimits::MAX_INPUT_BYTES);
        assert_eq!(wire.max_fields(), WireLimits::MAX_FIELDS);
        assert_eq!(wire.max_rewrite_work(), WireLimits::MAX_REWRITE_WORK);
        assert_eq!(wire.max_nesting(), WireLimits::MAX_NESTING);
        let codec = budget.codec_limits().expect("codec hard caps");
        assert_eq!(codec.max_bytes(), CodecLimits::DEFAULT.max_bytes());
        assert_eq!(codec.max_fields(), CodecLimits::DEFAULT.max_fields());
        let components = components();
        let workspace = Workspace::new(&components, CoreLimits::default()).expect("workspace");
        let (_values, _report) =
            symbols(&workspace, 1, FAMILY, 1, large_limits()).expect("large profile valid input");
    }

    #[test]
    fn codec_requirements_match_follow_on_reference_and_decode_passes() {
        let chart = chart_payload();
        let requirements = codec::reference_requirements(&chart, CodecLimits::DEFAULT)
            .expect("reference requirements");
        let (_, report) =
            codec::references_with_report(&chart, CodecLimits::DEFAULT).expect("reference report");
        assert_eq!(requirements.fields(), report.fields());
        assert_eq!(requirements.work_bytes(), report.work_bytes());
        let payload = override_payload();
        let requirements = codec::decode_requirements(
            &payload,
            codec::Archive::SeriesNonStyle,
            CodecLimits::DEFAULT,
        )
        .expect("decode requirements");
        let (_, report) = codec::decode_with_report(
            &payload,
            codec::Archive::SeriesNonStyle,
            CodecLimits::DEFAULT,
        )
        .expect("decode report");
        assert_eq!(requirements.fields(), report.fields());
        assert_eq!(requirements.work_bytes(), report.work_bytes());
    }

    #[test]
    fn exact_symbol_noop_reports_unchanged_after_first_write() {
        let components = components();
        let mut workspace = Workspace::new(&components, CoreLimits::default()).expect("workspace");
        let values = expected();
        let _ =
            set_symbols(&mut workspace, 1, FAMILY, &values, limits(4_096)).expect("first write");
        let (changed, report) =
            set_symbols(&mut workspace, 1, FAMILY, &values, limits(4_096)).expect("second write");
        assert!(!changed);
        assert!(!report.changed());
    }

    #[test]
    fn shared_override_is_copied_and_unselected_chart_stays_exact() {
        let components = shared_components();
        let mut workspace = Workspace::new(&components, CoreLimits::default()).expect("workspace");
        let (changed, report) = set_symbols(
            &mut workspace,
            1,
            FAMILY,
            &expected_without_size(),
            limits(4_096),
        )
        .expect("shared override write");
        assert!(changed);
        assert!(
            !report
                .scope()
                .expect("mutation scope")
                .planned_new_objects()
                .is_empty()
        );
        let chart_one = workspace.object(1).expect("selected chart");
        let chart_two = workspace.object(5).expect("unselected chart");
        let references_one = codec::references(&chart_one.messages[0].data, CodecLimits::DEFAULT)
            .expect("selected references");
        let references_two = codec::references(&chart_two.messages[0].data, CodecLimits::DEFAULT)
            .expect("unselected references");
        let selected = references_one
            .series_non_styles()
            .expect("selected sparse references")
            .entries()[0]
            .1;
        let unselected = references_two
            .series_non_styles()
            .expect("unselected sparse references")
            .entries()[0]
            .1;
        assert_ne!(selected, 4);
        assert_eq!(unselected, 4);
    }

    #[test]
    fn fresh_styled_symbol_allocation_is_bounded_and_readable() {
        let components = fresh_components();
        let mut workspace = Workspace::new(&components, CoreLimits::default()).expect("workspace");
        let (changed, report) = set_symbols(&mut workspace, 1, FAMILY, &expected(), limits(4_096))
            .expect("fresh styled symbol write");
        assert!(changed);
        assert!(report.changed());
        let references = codec::references(
            &workspace.object(1).expect("chart").messages[0].data,
            CodecLimits::DEFAULT,
        )
        .expect("chart references");
        assert!(
            references
                .series_non_styles()
                .is_some_and(|sparse| sparse.entries().len() == 1)
        );
        assert!(workspace.object(4).is_ok());
    }
}
