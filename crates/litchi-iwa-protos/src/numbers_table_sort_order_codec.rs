//! Strict, source-preserving ownership of `TST.TableModelArchive.sort_order`.
//!
//! The package owner resolves the table/model graph.  This module owns only
//! the field-44 wire projection: required sort scope, repeated index/direction
//! rules, canonicality, and an exact prepared rewrite.  Field 45 and every
//! other model field remain opaque source bytes.  Unknown scalar encodings and
//! balanced unknown groups are retained verbatim. `unknown overlong` scalar
//! encodings are accepted and copied without normalization.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict wire pass intentionally precedes its small public projections."
)]

use core::{fmt, mem::size_of};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_table_sort_order_generated::LitchiIwaNumbersTableSortOrderProjection as projection;

const SORT_ORDER_FIELD: u32 = 44;
const SORT_TYPE_FIELD: u32 = 1;
const SORT_RULES_FIELD: u32 = 2;
const SORT_RULE_COLUMN_FIELD: u32 = 1;
const SORT_RULE_DIRECTION_FIELD: u32 = 2;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MAX_RECURSION: u32 = 64;

/// A persisted table-sort scope.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SortScope {
    /// Sort the complete body table.
    EntireTable,
    /// Sort the row range selected by the caller/UI.
    SelectedRows,
}

impl SortScope {
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::EntireTable => 0,
            Self::SelectedRows => 1,
        }
    }

    fn from_native(value: i32) -> Result<Self, DecodeError> {
        match value {
            0 => Ok(Self::EntireTable),
            1 => Ok(Self::SelectedRows),
            _ => Err(DecodeError::invalid("sort scope")),
        }
    }
}

/// A persisted column sort direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SortDirection {
    Ascending,
    Descending,
}

impl SortDirection {
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Ascending => 0,
            Self::Descending => 1,
        }
    }

    fn from_native(value: i32) -> Result<Self, DecodeError> {
        match value {
            0 => Ok(Self::Ascending),
            1 => Ok(Self::Descending),
            _ => Err(DecodeError::invalid("sort direction")),
        }
    }
}

/// One ordered persisted sort rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SortRule {
    column: u32,
    direction: SortDirection,
}

impl SortRule {
    #[must_use]
    pub const fn new(column: u32, direction: SortDirection) -> Self {
        Self { column, direction }
    }

    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }

    #[must_use]
    pub const fn direction(self) -> SortDirection {
        self.direction
    }
}

/// Validated non-empty semantic sort order.  An empty native marker decodes
/// to `None`; it is not exposed as an empty semantic value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortOrderSnapshot {
    scope: SortScope,
    rules: Vec<SortRule>,
}

impl SortOrderSnapshot {
    /// Construct a semantic order, rejecting empty and duplicate-column rules.
    pub fn new(
        scope: SortScope,
        rules: impl IntoIterator<Item = SortRule>,
    ) -> Result<Self, DecodeError> {
        let rules = rules.into_iter().collect::<Vec<_>>();
        validate_rules(scope, &rules, usize::MAX, usize::MAX)?;
        Ok(Self { scope, rules })
    }

    #[must_use]
    pub const fn scope(&self) -> SortScope {
        self.scope
    }

    #[must_use]
    pub fn rules(&self) -> &[SortRule] {
        &self.rules
    }
}

/// Finite limits for strict source decoding and preparation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_rules: usize,
    max_columns: usize,
    max_allocations: usize,
}

impl DecodeOptions {
    /// Construct explicit input/output/field/work/nesting/rule/column limits.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_rules: usize,
        max_columns: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_rules,
            max_columns,
            max_allocations: 16,
        }
    }

    /// Build a conservative finite policy from one source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(2).max(1),
            bytes.saturating_mul(8).max(1),
            bytes.saturating_mul(32).max(1),
            16,
            bytes.max(1),
            usize::try_from(u32::MAX).unwrap_or(usize::MAX),
        )
        .with_max_allocations(1024)
    }

    #[must_use]
    pub const fn with_max_input_bytes(mut self, value: usize) -> Self {
        self.max_input_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_max_fields(mut self, value: usize) -> Self {
        self.max_fields = value;
        self
    }

    #[must_use]
    pub const fn with_max_work_bytes(mut self, value: usize) -> Self {
        self.max_work_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_recursion_limit(mut self, value: u32) -> Self {
        self.recursion_limit = value;
        self
    }

    #[must_use]
    pub const fn with_max_rules(mut self, value: usize) -> Self {
        self.max_rules = value;
        self
    }

    #[must_use]
    pub const fn with_max_columns(mut self, value: usize) -> Self {
        self.max_columns = value;
        self
    }

    #[must_use]
    pub const fn with_max_allocations(mut self, value: usize) -> Self {
        self.max_allocations = value;
        self
    }

    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    #[must_use]
    pub const fn recursion_limit(self) -> u32 {
        self.recursion_limit
    }

    #[must_use]
    pub const fn max_rules(self) -> usize {
        self.max_rules
    }

    #[must_use]
    pub const fn max_columns(self) -> usize {
        self.max_columns
    }

    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_input_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Alias used by package code that describes a strict rewrite policy.
pub type RewriteOptions = DecodeOptions;

/// Typed strict resource failures.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    InputBytes { observed: usize, maximum: usize },
    OutputBytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    WorkBytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    Rules { observed: usize, maximum: usize },
    Columns { observed: usize, maximum: usize },
    Allocations { observed: usize, maximum: usize },
    RetainedBytes { observed: usize, maximum: usize },
    ScratchBytes { observed: usize, maximum: usize },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ErrorKind {
    Invalid(&'static str),
    Limit(DecodeLimit),
    Allocation { requested: usize },
    Projection,
}

/// Strict generated-free sort-order failure.  Diagnostics never include IDs
/// or source bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    kind: ErrorKind,
}

impl DecodeError {
    const fn invalid(reason: &'static str) -> Self {
        Self {
            kind: ErrorKind::Invalid(reason),
        }
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: ErrorKind::Limit(limit),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: ErrorKind::Projection,
        }
    }

    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        match self.kind {
            ErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    #[must_use]
    pub const fn allocation_amount(self) -> Option<usize> {
        match self.kind {
            ErrorKind::Allocation { requested } => Some(requested),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.kind {
            ErrorKind::Invalid(reason) => {
                write!(formatter, "invalid Numbers table sort-order {reason}")
            },
            ErrorKind::Limit(_) => {
                formatter.write_str("Numbers table sort-order resource limit exceeded")
            },
            ErrorKind::Allocation { requested } => {
                write!(
                    formatter,
                    "cannot allocate sort-order candidate for {requested} bytes"
                )
            },
            ErrorKind::Projection => {
                formatter.write_str("sort-order projection verification failed")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

/// Decode/preparation accounting.  Rewrite preparation reports conservative
/// upper bounds for candidate scratch/allocation/work stages; execution
/// replays those bounds before reserving the candidate Vec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    rules: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
    #[must_use]
    pub const fn rules(self) -> usize {
        self.rules
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Prepared execution ceilings.  Preparation performs no candidate-output
/// allocation; the output Vec is reserved only after these limits are checked.
/// Allocation and scratch values are conservative operation-local bounds that
/// include candidate emission and semantic readback staging.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub rules: usize,
    pub allocations: usize,
    pub retained_bytes: usize,
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    #[must_use]
    pub const fn exact(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            rules: self.rules,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }

    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Caller-supplied limits replayed before candidate allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub rules: usize,
    pub allocations: usize,
    pub retained_bytes: usize,
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        requirements.exact()
    }

    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            max_depth: u32::MAX,
            rules: usize::MAX,
            allocations: usize::MAX,
            retained_bytes: usize::MAX,
            scratch_bytes: usize::MAX,
        }
    }

    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }
    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }
    #[must_use]
    pub const fn with_rules(mut self, value: usize) -> Self {
        self.rules = value;
        self
    }
    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }
    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }
}

/// Candidate model bytes and exact execution accounting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: DecodeReport,
}

impl RewriteOutput {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
    #[must_use]
    pub fn output(&self) -> &[u8] {
        &self.bytes
    }
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
    #[must_use]
    pub const fn report(&self) -> DecodeReport {
        self.report
    }
}

/// Prepared source-borrowing field-44 rewrite.
#[derive(Debug, Clone)]
pub struct PreparedTableSortOrderRewrite<'source> {
    source: &'source [u8],
    desired: Option<SortOrderSnapshot>,
    options: DecodeOptions,
    current: ParsedModel,
    report: DecodeReport,
    requirements: RewriteExecutionRequirements,
}

impl<'source> PreparedTableSortOrderRewrite<'source> {
    #[must_use]
    pub const fn prepare_report(&self) -> DecodeReport {
        self.report
    }
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute after all caller ceilings have passed.  The source remains
    /// untouched when any check or fallible reservation fails.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_limits(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| DecodeError {
                kind: ErrorKind::Allocation {
                    requested: self.requirements.output_bytes,
                },
            })?;
        emit_model_rewrite(
            self.source,
            &self.current,
            self.desired.as_ref(),
            &mut output,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::projection());
        }
        let verify_options = DecodeOptions {
            max_input_bytes: self.options.max_input_bytes.max(output.len()),
            max_output_bytes: self.options.max_output_bytes.max(output.len()),
            max_fields: self.options.max_fields.max(self.requirements.fields),
            max_work_bytes: self
                .options
                .max_work_bytes
                .max(self.requirements.work_bytes),
            ..self.options
        };
        let verified = decode_table_sort_order(&output, verify_options)?;
        if verified != self.desired {
            return Err(DecodeError::projection());
        }
        Ok(RewriteOutput {
            bytes: output,
            report: DecodeReport {
                input_bytes: self.source.len(),
                output_bytes: self.requirements.output_bytes,
                fields: self.requirements.fields,
                work_bytes: self.requirements.work_bytes,
                max_depth: self.requirements.max_depth,
                rules: self.requirements.rules,
                allocations: self.requirements.allocations,
                retained_bytes: self.requirements.retained_bytes,
                scratch_bytes: self.requirements.scratch_bytes,
            },
        })
    }
}

/// Decode field 44 from a complete `TST.TableModelArchive` payload.
pub fn decode_table_sort_order(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<SortOrderSnapshot>, DecodeError> {
    Ok(decode_table_sort_order_with_report(source, options)?.0)
}

/// Decode field 44 and return exact strict traversal accounting.
pub fn decode_table_sort_order_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(Option<SortOrderSnapshot>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(options);
    let model = parse_model(source, options, &mut budget)?;
    let semantic = model.sort.as_ref().and_then(ParsedSort::semantic);
    if let Some(sort) = &model.sort {
        force_buffa(&sort.payload, options, sort.scope)?;
    }
    Ok((
        semantic,
        budget.report(source.len(), source.len(), budget.allocations, source.len()),
    ))
}

/// Prepare a strict rewrite of field 44. `None` means clear to an explicit
/// empty marker when one exists; clearing an absent field is an exact no-op.
pub fn prepare_table_sort_order_rewrite<'source>(
    source: &'source [u8],
    desired: Option<SortOrderSnapshot>,
    options: DecodeOptions,
) -> Result<PreparedTableSortOrderRewrite<'source>, DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(options);
    let current = parse_model(source, options, &mut budget)?;
    if let Some(desired_order) = &desired {
        validate_rules(
            desired_order.scope,
            &desired_order.rules,
            options.max_rules,
            options.max_columns,
        )?;
    }
    if let Some(sort) = &current.sort {
        force_buffa(&sort.payload, options, sort.scope)?;
    }
    let output_bytes = measure_model_rewrite(source, &current, desired.as_ref())?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    // Candidate verification may retain nested unknown groups and rule
    // framing. Start with the complete source traversal, then reserve a
    // conservative allowance for a newly inserted outer field and desired
    // rule envelopes rather than undercounting opaque nested fields.
    let desired_field_allowance = desired.as_ref().map_or(0, |order| {
        order.rules.len().saturating_mul(3).saturating_add(1)
    });
    let candidate_fields = budget
        .fields
        .checked_add(desired_field_allowance)
        .ok_or(DecodeError::projection())?;
    let candidate_work = source
        .len()
        .checked_add(output_bytes)
        .and_then(|value| value.checked_mul(8))
        .ok_or(DecodeError::projection())?;
    let fields = budget
        .fields
        .checked_add(candidate_fields)
        .ok_or(DecodeError::projection())?;
    let work_bytes = budget
        .work
        .checked_add(candidate_work)
        .ok_or(DecodeError::projection())?;
    if fields > options.max_fields {
        return Err(DecodeError::limit(DecodeLimit::Fields {
            observed: fields,
            maximum: options.max_fields,
        }));
    }
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError::limit(DecodeLimit::WorkBytes {
            observed: work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let rules = desired.as_ref().map_or(0, |value| value.rules.len());
    let max_rule_count = current
        .sort
        .as_ref()
        .map_or(0, |sort| sort.rules.len())
        .max(rules);
    let scratch = budget
        .scratch_bytes
        .checked_add(source.len().saturating_add(output_bytes).saturating_mul(2))
        .and_then(|value| value.checked_add(max_rule_count.saturating_mul(size_of::<ParsedRule>())))
        .ok_or(DecodeError::projection())?;
    let allocations = budget
        .allocations
        .checked_add(8)
        .and_then(|value| value.checked_add(max_rule_count.saturating_mul(8)))
        .ok_or(DecodeError::projection())?;
    if allocations > options.max_allocations {
        return Err(DecodeError::limit(DecodeLimit::Allocations {
            observed: allocations,
            maximum: options.max_allocations,
        }));
    }
    let report = DecodeReport {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work_bytes,
        max_depth: budget.max_depth,
        rules,
        allocations,
        retained_bytes: output_bytes,
        scratch_bytes: scratch,
    };
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes,
        max_depth: budget.max_depth,
        rules,
        allocations,
        retained_bytes: output_bytes,
        scratch_bytes: scratch,
    };
    Ok(PreparedTableSortOrderRewrite {
        source,
        desired,
        options,
        current,
        report,
        requirements,
    })
}

/// One-shot source-preserving field-44 rewrite.
pub fn rewrite_table_sort_order(
    source: &[u8],
    desired: Option<SortOrderSnapshot>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_table_sort_order_rewrite(source, desired, options)?;
    let limits = prepared.execution_requirements().exact();
    prepared.execute(limits)
}

/// Canonical model-field naming retained for package/boundary ratchets.
pub fn decode_table_model_sort_order(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<SortOrderSnapshot>, DecodeError> {
    decode_table_sort_order(source, options)
}

/// Canonical model-field naming retained for package/boundary ratchets.
pub fn decode_table_model_sort_order_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(Option<SortOrderSnapshot>, DecodeReport), DecodeError> {
    decode_table_sort_order_with_report(source, options)
}

/// Canonical model-field naming retained for package/boundary ratchets.
pub fn prepare_table_model_sort_order_rewrite<'source>(
    source: &'source [u8],
    desired: Option<SortOrderSnapshot>,
    options: DecodeOptions,
) -> Result<PreparedTableSortOrderRewrite<'source>, DecodeError> {
    prepare_table_sort_order_rewrite(source, desired, options)
}

/// Compatibility alias used by the package owner and boundary checker.
pub type PreparedTableModelSortOrderRewrite<'source> = PreparedTableSortOrderRewrite<'source>;

/// Canonical model-field naming retained for package/boundary ratchets.
pub fn rewrite_table_model_sort_order(
    source: &[u8],
    desired: Option<SortOrderSnapshot>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    rewrite_table_sort_order(source, desired, options)
}

/// Build a canonical nested `TableSortOrderArchive` payload.
pub fn canonical_table_sort_order(order: &SortOrderSnapshot) -> Result<Vec<u8>, DecodeError> {
    emit_canonical_payload(order)
}

#[derive(Debug, Clone)]
struct ParsedModel {
    fields: Vec<FieldSpan>,
    sort: Option<ParsedSort>,
}

#[derive(Debug, Clone)]
struct ParsedSort {
    payload: Vec<u8>,
    fields: Vec<FieldSpan>,
    outer: FieldSpan,
    scope: SortScope,
    rules: Vec<ParsedRule>,
}

impl ParsedSort {
    fn semantic(&self) -> Option<SortOrderSnapshot> {
        if self.rules.is_empty() {
            None
        } else {
            Some(SortOrderSnapshot {
                scope: self.scope,
                rules: self
                    .rules
                    .iter()
                    .map(|rule| SortRule::new(rule.column, rule.direction))
                    .collect(),
            })
        }
    }
}

#[derive(Debug, Clone)]
struct ParsedRule {
    column: u32,
    direction: SortDirection,
    payload: Vec<u8>,
    fields: Vec<FieldSpan>,
}

#[derive(Debug, Clone, Copy)]
struct FieldSpan {
    number: u32,
    wire: u8,
    start: usize,
    end: usize,
    value_start: usize,
    value_end: usize,
    key_canonical: bool,
    length_canonical: bool,
}

#[derive(Debug, Clone, Copy)]
struct Varint {
    value: u64,
    canonical: bool,
}

struct Budget {
    options: DecodeOptions,
    fields: usize,
    work: usize,
    max_depth: u32,
    rules: usize,
    scratch_bytes: usize,
    allocations: usize,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            options,
            fields: 0,
            work: 0,
            max_depth: 1,
            rules: 0,
            scratch_bytes: 0,
            allocations: 0,
        }
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        self.fields = self
            .fields
            .checked_add(1)
            .ok_or(DecodeError::projection())?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        Ok(())
    }

    fn work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        self.work = self
            .work
            .checked_add(bytes)
            .ok_or(DecodeError::projection())?;
        if self.work > self.options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::WorkBytes {
                observed: self.work,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        self.max_depth = self.max_depth.max(depth);
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        Ok(())
    }

    fn rules(&mut self) -> Result<(), DecodeError> {
        self.rules = self.rules.checked_add(1).ok_or(DecodeError::projection())?;
        if self.rules > self.options.max_rules {
            return Err(DecodeError::limit(DecodeLimit::Rules {
                observed: self.rules,
                maximum: self.options.max_rules,
            }));
        }
        Ok(())
    }

    fn allocate(&mut self) -> Result<(), DecodeError> {
        self.allocations = self
            .allocations
            .checked_add(1)
            .ok_or(DecodeError::projection())?;
        if self.allocations > self.options.max_allocations {
            return Err(DecodeError::limit(DecodeLimit::Allocations {
                observed: self.allocations,
                maximum: self.options.max_allocations,
            }));
        }
        Ok(())
    }

    fn reserve_spans(&mut self, count: usize) -> Result<(), DecodeError> {
        let bytes = count
            .checked_mul(size_of::<FieldSpan>())
            .ok_or(DecodeError::projection())?;
        self.scratch_bytes = self
            .scratch_bytes
            .checked_add(bytes)
            .ok_or(DecodeError::projection())?;
        Ok(())
    }

    fn report(
        &self,
        input_bytes: usize,
        output_bytes: usize,
        allocations: usize,
        retained_bytes: usize,
    ) -> DecodeReport {
        DecodeReport {
            input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work,
            max_depth: self.max_depth,
            rules: self.rules,
            allocations,
            retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    if source.len() > options.max_input_bytes {
        return Err(DecodeError::limit(DecodeLimit::InputBytes {
            observed: source.len(),
            maximum: options.max_input_bytes,
        }));
    }
    if options.recursion_limit > MAX_RECURSION {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}

fn parse_model(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ParsedModel, DecodeError> {
    budget.work(source.len())?;
    let fields = scan_fields(source, options, budget, 1)?;
    let mut sort = None;
    for field in &fields {
        if field.number != SORT_ORDER_FIELD {
            continue;
        }
        if sort.is_some() {
            return Err(DecodeError::invalid("duplicate sort_order field"));
        }
        if field.wire != 2 || !field.key_canonical || !field.length_canonical {
            return Err(DecodeError::invalid("sort_order wire"));
        }
        let payload = source[field.value_start..field.value_end].to_vec();
        budget.allocate()?;
        sort = Some(parse_sort_payload(payload, *field, options, budget)?);
    }
    Ok(ParsedModel { fields, sort })
}

fn parse_sort_payload(
    payload: Vec<u8>,
    outer: FieldSpan,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ParsedSort, DecodeError> {
    budget.work(payload.len())?;
    let fields = scan_fields(&payload, options, budget, 2)?;
    let mut scope = None;
    let mut rules = Vec::new();
    for field in &fields {
        match field.number {
            SORT_TYPE_FIELD => {
                if scope.is_some() || field.wire != 0 || !field.key_canonical {
                    return Err(DecodeError::invalid("sort type"));
                }
                if !field.length_canonical {
                    return Err(DecodeError::invalid("sort type varint"));
                }
                let value = read_varint_value(&payload, *field)?;
                scope = Some(SortScope::from_native(
                    i32::try_from(value).map_err(|_| DecodeError::invalid("sort scope"))?,
                )?);
            },
            SORT_RULES_FIELD => {
                if field.wire != 2 || !field.key_canonical || !field.length_canonical {
                    return Err(DecodeError::invalid("sort rule wire"));
                }
                budget.rules()?;
                let raw = payload[field.value_start..field.value_end].to_vec();
                budget.allocate()?;
                let parsed = parse_rule(raw, options, budget)?;
                if rules
                    .iter()
                    .any(|rule: &ParsedRule| rule.column == parsed.column)
                {
                    return Err(DecodeError::invalid("duplicate sort column"));
                }
                rules
                    .try_reserve(1)
                    .map_err(|_| DecodeError::invalid("sort rules allocation"))?;
                rules.push(parsed);
            },
            _ => {},
        }
    }
    let scope = scope.ok_or_else(|| DecodeError::invalid("missing sort type"))?;
    if !rules.is_empty() {
        validate_rules(
            scope,
            &rules
                .iter()
                .map(|rule| SortRule::new(rule.column, rule.direction))
                .collect::<Vec<_>>(),
            options.max_rules,
            options.max_columns,
        )?;
    }
    Ok(ParsedSort {
        payload,
        fields,
        outer,
        scope,
        rules,
    })
}

fn parse_rule(
    payload: Vec<u8>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ParsedRule, DecodeError> {
    budget.work(payload.len())?;
    budget.allocate()?;
    let fields = scan_fields(&payload, options, budget, 3)?;
    let mut column = None;
    let mut direction = None;
    for field in &fields {
        match field.number {
            SORT_RULE_COLUMN_FIELD => {
                if column.is_some()
                    || field.wire != 0
                    || !field.key_canonical
                    || !field.length_canonical
                {
                    return Err(DecodeError::invalid("sort rule column"));
                }
                let value = read_varint_value(&payload, *field)?;
                column =
                    Some(u32::try_from(value).map_err(|_| DecodeError::invalid("sort column"))?);
            },
            SORT_RULE_DIRECTION_FIELD => {
                if direction.is_some()
                    || field.wire != 0
                    || !field.key_canonical
                    || !field.length_canonical
                {
                    return Err(DecodeError::invalid("sort rule direction"));
                }
                let value = read_varint_value(&payload, *field)?;
                direction = Some(SortDirection::from_native(
                    i32::try_from(value).map_err(|_| DecodeError::invalid("sort direction"))?,
                )?);
            },
            _ => {},
        }
    }
    Ok(ParsedRule {
        column: column.ok_or_else(|| DecodeError::invalid("missing sort column"))?,
        direction: direction.ok_or_else(|| DecodeError::invalid("missing sort direction"))?,
        payload,
        fields,
    })
}

fn validate_rules(
    _scope: SortScope,
    rules: &[SortRule],
    max_rules: usize,
    max_columns: usize,
) -> Result<(), DecodeError> {
    if rules.is_empty() {
        return Err(DecodeError::invalid("empty semantic sort order"));
    }
    if rules.len() > max_rules {
        return Err(DecodeError::limit(DecodeLimit::Rules {
            observed: rules.len(),
            maximum: max_rules,
        }));
    }
    for (index, rule) in rules.iter().enumerate() {
        let column =
            usize::try_from(rule.column).map_err(|_| DecodeError::invalid("sort column"))?;
        if column >= max_columns {
            return Err(DecodeError::limit(DecodeLimit::Columns {
                observed: column.saturating_add(1),
                maximum: max_columns,
            }));
        }
        if rules[..index]
            .iter()
            .any(|previous| previous.column == rule.column)
        {
            return Err(DecodeError::invalid("duplicate sort column"));
        }
    }
    Ok(())
}

fn force_buffa(
    payload: &[u8],
    options: DecodeOptions,
    scope: SortScope,
) -> Result<(), DecodeError> {
    let view: projection::TableSortOrderArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(payload)
        .map_err(|_| DecodeError::projection())?;
    if view.r#type != scope.native_value() {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn scan_fields(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<Vec<FieldSpan>, DecodeError> {
    budget.depth(depth)?;
    budget.allocate()?;
    let mut fields = Vec::new();
    fields
        .try_reserve(source.len().min(options.max_fields))
        .map_err(|_| DecodeError::invalid("field allocation"))?;
    budget.reserve_spans(source.len().min(options.max_fields))?;
    let mut offset = 0usize;
    while offset < source.len() {
        let field = parse_field(source, &mut offset, budget, depth)?;
        fields.push(field);
    }
    Ok(fields)
}

fn parse_field(
    source: &[u8],
    offset: &mut usize,
    budget: &mut Budget,
    depth: u32,
) -> Result<FieldSpan, DecodeError> {
    let start = *offset;
    let key = read_varint(source, offset)?;
    let number = u32::try_from(key.value >> 3).map_err(|_| DecodeError::invalid("field number"))?;
    let wire = u8::try_from(key.value & 7).map_err(|_| DecodeError::invalid("wire type"))?;
    if number == 0 || number > MAX_FIELD_NUMBER || wire == 4 {
        return Err(DecodeError::invalid("field key"));
    }
    budget.field()?;
    let mut value_start = *offset;
    let (value_end, length_canonical) = match wire {
        0 => {
            let value = read_varint(source, offset)?;
            (*offset, value.canonical)
        },
        1 => {
            let end = offset
                .checked_add(8)
                .ok_or(DecodeError::invalid("fixed64"))?;
            if end > source.len() {
                return Err(DecodeError::invalid("fixed64"));
            }
            *offset = end;
            (end, true)
        },
        2 => {
            let length = read_varint(source, offset)?;
            let length_value =
                usize::try_from(length.value).map_err(|_| DecodeError::invalid("length"))?;
            value_start = *offset;
            let end = offset
                .checked_add(length_value)
                .ok_or(DecodeError::invalid("length"))?;
            if end > source.len() {
                return Err(DecodeError::invalid("length"));
            }
            *offset = end;
            (end, length.canonical)
        },
        3 => {
            let end = skip_group(source, offset, budget, depth, number)?;
            (end, true)
        },
        5 => {
            let end = offset
                .checked_add(4)
                .ok_or(DecodeError::invalid("fixed32"))?;
            if end > source.len() {
                return Err(DecodeError::invalid("fixed32"));
            }
            *offset = end;
            (end, true)
        },
        _ => return Err(DecodeError::invalid("wire type")),
    };
    Ok(FieldSpan {
        number,
        wire,
        start,
        end: *offset,
        value_start,
        value_end,
        key_canonical: key.canonical,
        length_canonical,
    })
}

fn skip_group(
    source: &[u8],
    offset: &mut usize,
    budget: &mut Budget,
    depth: u32,
    opening_number: u32,
) -> Result<usize, DecodeError> {
    let next_depth = depth.checked_add(1).ok_or(DecodeError::projection())?;
    budget.depth(next_depth)?;
    loop {
        if *offset >= source.len() {
            return Err(DecodeError::invalid("unterminated unknown group"));
        }
        let key = read_varint(source, offset)?;
        let number =
            u32::try_from(key.value >> 3).map_err(|_| DecodeError::invalid("group field"))?;
        let wire = u8::try_from(key.value & 7).map_err(|_| DecodeError::invalid("group wire"))?;
        if number == 0 || number > MAX_FIELD_NUMBER {
            return Err(DecodeError::invalid("group field"));
        }
        budget.field()?;
        match wire {
            4 => {
                if number != opening_number {
                    return Err(DecodeError::invalid("mismatched unknown group"));
                }
                return Ok(*offset);
            },
            3 => {
                skip_group(source, offset, budget, next_depth, number)?;
            },
            0 => {
                let _ = read_varint(source, offset)?;
            },
            1 => {
                let end = offset
                    .checked_add(8)
                    .ok_or(DecodeError::invalid("group fixed64"))?;
                if end > source.len() {
                    return Err(DecodeError::invalid("group fixed64"));
                }
                *offset = end;
            },
            2 => {
                let length = read_varint(source, offset)?;
                let length = usize::try_from(length.value)
                    .map_err(|_| DecodeError::invalid("group length"))?;
                let end = offset
                    .checked_add(length)
                    .ok_or(DecodeError::invalid("group length"))?;
                if end > source.len() {
                    return Err(DecodeError::invalid("group length"));
                }
                *offset = end;
            },
            5 => {
                let end = offset
                    .checked_add(4)
                    .ok_or(DecodeError::invalid("group fixed32"))?;
                if end > source.len() {
                    return Err(DecodeError::invalid("group fixed32"));
                }
                *offset = end;
            },
            _ => return Err(DecodeError::invalid("group wire")),
        }
    }
}

fn read_varint(source: &[u8], offset: &mut usize) -> Result<Varint, DecodeError> {
    let start = *offset;
    let mut value = 0u64;
    for shift in (0..70).step_by(7) {
        let byte = *source
            .get(*offset)
            .ok_or(DecodeError::invalid("truncated varint"))?;
        *offset = offset.checked_add(1).ok_or(DecodeError::projection())?;
        if shift == 63 && byte > 1 {
            return Err(DecodeError::invalid("varint overflow"));
        }
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            let canonical = *offset - start == varint_len(value);
            return Ok(Varint { value, canonical });
        }
    }
    Err(DecodeError::invalid("varint overflow"))
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 128 {
        value >>= 7;
        length += 1;
    }
    length
}

fn read_varint_value(source: &[u8], field: FieldSpan) -> Result<u64, DecodeError> {
    let mut offset = field.value_start;
    Ok(read_varint(source, &mut offset)?.value)
}

fn encode_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 128 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn encoded_varint_len(value: u64) -> usize {
    varint_len(value)
}

fn encoded_field_len(number: u32, payload_len: usize) -> Result<usize, DecodeError> {
    let key = encoded_varint_len(u64::from(number) << 3);
    key.checked_add(encoded_varint_len(payload_len as u64))
        .and_then(|value| value.checked_add(payload_len))
        .ok_or(DecodeError::projection())
}

fn measure_model_rewrite(
    source: &[u8],
    current: &ParsedModel,
    desired: Option<&SortOrderSnapshot>,
) -> Result<usize, DecodeError> {
    let Some(sort) = &current.sort else {
        return if desired.is_some() {
            let payload = canonical_payload_len(desired.ok_or(DecodeError::projection())?)?;
            source
                .len()
                .checked_add(encoded_field_len(SORT_ORDER_FIELD, payload)?)
                .ok_or(DecodeError::projection())
        } else {
            Ok(source.len())
        };
    };
    let payload = measure_payload_rewrite(sort, desired)?;
    source
        .len()
        .checked_sub(sort.outer.end - sort.outer.start)
        .and_then(|base| base.checked_add(encoded_field_len(SORT_ORDER_FIELD, payload).ok()?))
        .ok_or(DecodeError::projection())
}

fn canonical_payload_len(order: &SortOrderSnapshot) -> Result<usize, DecodeError> {
    let mut length = encoded_varint_len(u64::from(SORT_TYPE_FIELD) << 3)
        .checked_add(encoded_varint_len(order.scope.native_value() as u64))
        .ok_or(DecodeError::projection())?;
    for rule in &order.rules {
        let rule_len = encoded_varint_len(8)
            .checked_add(encoded_varint_len(u64::from(rule.column)))
            .and_then(|value| value.checked_add(encoded_varint_len(16)))
            .and_then(|value| {
                value.checked_add(encoded_varint_len(rule.direction.native_value() as u64))
            })
            .ok_or(DecodeError::projection())?;
        length = length
            .checked_add(encoded_field_len(SORT_RULES_FIELD, rule_len)?)
            .ok_or(DecodeError::projection())?;
    }
    Ok(length)
}

fn measure_payload_rewrite(
    sort: &ParsedSort,
    desired: Option<&SortOrderSnapshot>,
) -> Result<usize, DecodeError> {
    let mut length = 0usize;
    let mut saw_rules = false;
    for field in &sort.fields {
        match field.number {
            SORT_TYPE_FIELD => {
                length = length
                    .checked_add(
                        encoded_varint_len(8)
                            + encoded_varint_len(sort.scope.native_value() as u64),
                    )
                    .ok_or(DecodeError::projection())?;
            },
            SORT_RULES_FIELD => {
                if !saw_rules {
                    saw_rules = true;
                    if let Some(order) = desired {
                        for rule in &order.rules {
                            let raw_len = matching_rule_len(sort, *rule)?;
                            length = length
                                .checked_add(encoded_field_len(SORT_RULES_FIELD, raw_len)?)
                                .ok_or(DecodeError::projection())?;
                        }
                    }
                }
            },
            _ => {
                length = length
                    .checked_add(field.end - field.start)
                    .ok_or(DecodeError::projection())?
            },
        }
    }
    if !saw_rules {
        if let Some(order) = desired {
            for rule in &order.rules {
                let raw_len = matching_rule_len(sort, *rule)?;
                length = length
                    .checked_add(encoded_field_len(SORT_RULES_FIELD, raw_len)?)
                    .ok_or(DecodeError::projection())?;
            }
        }
    }
    Ok(length)
}

fn matching_rule_len(sort: &ParsedSort, rule: SortRule) -> Result<usize, DecodeError> {
    if let Some(previous) = sort
        .rules
        .iter()
        .find(|previous| previous.column == rule.column)
    {
        let mut length = 0usize;
        for field in &previous.fields {
            if field.number == SORT_RULE_COLUMN_FIELD {
                length = length
                    .checked_add(encoded_varint_len(8) + encoded_varint_len(u64::from(rule.column)))
                    .ok_or(DecodeError::projection())?;
            } else if field.number == SORT_RULE_DIRECTION_FIELD {
                length = length
                    .checked_add(
                        encoded_varint_len(16)
                            + encoded_varint_len(rule.direction.native_value() as u64),
                    )
                    .ok_or(DecodeError::projection())?;
            } else {
                length = length
                    .checked_add(field.end - field.start)
                    .ok_or(DecodeError::projection())?;
            }
        }
        Ok(length)
    } else {
        let length = encoded_varint_len(8)
            .checked_add(encoded_varint_len(u64::from(rule.column)))
            .and_then(|value| value.checked_add(encoded_varint_len(16)))
            .and_then(|value| {
                value.checked_add(encoded_varint_len(rule.direction.native_value() as u64))
            })
            .ok_or(DecodeError::projection())?;
        Ok(length)
    }
}

fn emit_model_rewrite(
    source: &[u8],
    current: &ParsedModel,
    desired: Option<&SortOrderSnapshot>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut replaced = false;
    for field in &current.fields {
        if field.number != SORT_ORDER_FIELD {
            output.extend_from_slice(&source[field.start..field.end]);
            continue;
        }
        if replaced {
            return Err(DecodeError::invalid("duplicate sort_order field"));
        }
        replaced = true;
        let rewritten = emit_payload_rewrite(
            current.sort.as_ref().ok_or(DecodeError::projection())?,
            desired,
        )?;
        output.push(0xe2);
        output.push(0x02);
        encode_varint(rewritten.len() as u64, output);
        output.extend_from_slice(&rewritten);
    }
    if !replaced && let Some(order) = desired {
        let payload = emit_canonical_payload(order)?;
        output.push(0xe2);
        output.push(0x02);
        encode_varint(payload.len() as u64, output);
        output.extend_from_slice(&payload);
    }
    Ok(())
}

fn emit_payload_rewrite(
    sort: &ParsedSort,
    desired: Option<&SortOrderSnapshot>,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    let output_scope = desired.map_or(sort.scope, |order| order.scope);
    let mut replaced_rules = false;
    for field in &sort.fields {
        match field.number {
            SORT_TYPE_FIELD => {
                output.push(8);
                encode_varint(output_scope.native_value() as u64, &mut output);
            },
            SORT_RULES_FIELD => {
                if replaced_rules {
                    continue;
                }
                replaced_rules = true;
                if let Some(order) = desired {
                    for rule in &order.rules {
                        let raw = sort
                            .rules
                            .iter()
                            .find(|previous| previous.column == rule.column);
                        let rule_bytes = raw.map_or_else(
                            || emit_canonical_rule(*rule),
                            |previous| emit_rule_rewrite(previous, *rule),
                        )?;
                        output.push(0x12);
                        encode_varint(rule_bytes.len() as u64, &mut output);
                        output.extend_from_slice(&rule_bytes);
                    }
                }
            },
            _ => output.extend_from_slice(&sort.payload[field.start..field.end]),
        }
    }
    if !replaced_rules && let Some(order) = desired {
        for rule in &order.rules {
            let raw = sort
                .rules
                .iter()
                .find(|previous| previous.column == rule.column);
            let rule_bytes = raw.map_or_else(
                || emit_canonical_rule(*rule),
                |previous| emit_rule_rewrite(previous, *rule),
            )?;
            output.push(0x12);
            encode_varint(rule_bytes.len() as u64, &mut output);
            output.extend_from_slice(&rule_bytes);
        }
    }
    Ok(output)
}

fn emit_canonical_payload(order: &SortOrderSnapshot) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output.push(8);
    encode_varint(order.scope.native_value() as u64, &mut output);
    for rule in &order.rules {
        let bytes = emit_canonical_rule(*rule)?;
        output.push(0x12);
        encode_varint(bytes.len() as u64, &mut output);
        output.extend_from_slice(&bytes);
    }
    Ok(output)
}

fn emit_canonical_rule(rule: SortRule) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output.push(8);
    encode_varint(u64::from(rule.column), &mut output);
    output.push(16);
    encode_varint(rule.direction.native_value() as u64, &mut output);
    Ok(output)
}

fn emit_rule_rewrite(previous: &ParsedRule, rule: SortRule) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    for field in &previous.fields {
        match field.number {
            SORT_RULE_COLUMN_FIELD => {
                output.push(8);
                encode_varint(u64::from(rule.column), &mut output);
            },
            SORT_RULE_DIRECTION_FIELD => {
                output.push(16);
                encode_varint(rule.direction.native_value() as u64, &mut output);
            },
            _ => output.extend_from_slice(&previous.payload[field.start..field.end]),
        }
    }
    Ok(output)
}

fn check_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    if requirements.output_bytes > limits.output_bytes {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: requirements.output_bytes,
            maximum: limits.output_bytes,
        }));
    }
    if requirements.fields > limits.fields {
        return Err(DecodeError::limit(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: limits.fields,
        }));
    }
    if requirements.work_bytes > limits.work_bytes {
        return Err(DecodeError::limit(DecodeLimit::WorkBytes {
            observed: requirements.work_bytes,
            maximum: limits.work_bytes,
        }));
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    if requirements.rules > limits.rules {
        return Err(DecodeError::limit(DecodeLimit::Rules {
            observed: requirements.rules,
            maximum: limits.rules,
        }));
    }
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::limit(DecodeLimit::Allocations {
            observed: requirements.allocations,
            maximum: limits.allocations,
        }));
    }
    if requirements.retained_bytes > limits.retained_bytes {
        return Err(DecodeError::limit(DecodeLimit::RetainedBytes {
            observed: requirements.retained_bytes,
            maximum: limits.retained_bytes,
        }));
    }
    if requirements.scratch_bytes > limits.scratch_bytes {
        return Err(DecodeError::limit(DecodeLimit::ScratchBytes {
            observed: requirements.scratch_bytes,
            maximum: limits.scratch_bytes,
        }));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
            .with_max_columns(8)
            .with_max_rules(8)
    }

    fn model(payload: &[u8]) -> Vec<u8> {
        let mut output = vec![0x22, 0x01, 0x78];
        output.extend_from_slice(&[0xe2, 0x02, payload.len() as u8]);
        output.extend_from_slice(payload);
        output.extend_from_slice(&[0xea, 0x02, 0x01, 0x7f]);
        output
    }

    fn sort(scope: u8, rules: &[(u8, u8)]) -> Vec<u8> {
        let mut output = vec![8, scope];
        for (column, direction) in rules {
            output.extend_from_slice(&[0x12, 0x04, 8, *column, 16, *direction]);
        }
        output
    }

    #[test]
    fn decode_and_rewrite_preserve_tracker_and_unknowns() {
        let source = model(&sort(0, &[(1, 0)]));
        let (decoded, _) = decode_table_sort_order_with_report(&source, options(&source)).unwrap();
        assert_eq!(
            decoded.as_ref().unwrap().rules()[0],
            SortRule::new(1, SortDirection::Ascending)
        );
        let desired = SortOrderSnapshot::new(
            SortScope::SelectedRows,
            [SortRule::new(2, SortDirection::Descending)],
        )
        .unwrap();
        let output = rewrite_table_sort_order(&source, Some(desired), options(&source)).unwrap();
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| window == [0xea, 0x02, 0x01, 0x7f])
        );
        assert_eq!(
            decode_table_sort_order(output.bytes(), options(output.bytes()))
                .unwrap()
                .unwrap()
                .scope(),
            SortScope::SelectedRows
        );
        let back = SortOrderSnapshot::new(
            SortScope::EntireTable,
            [SortRule::new(2, SortDirection::Descending)],
        )
        .unwrap();
        let output =
            rewrite_table_sort_order(output.bytes(), Some(back), options(output.bytes())).unwrap();
        assert_eq!(
            decode_table_sort_order(output.bytes(), options(output.bytes()))
                .unwrap()
                .unwrap()
                .scope(),
            SortScope::EntireTable
        );
    }

    #[test]
    fn clear_keeps_scope_and_unknown_payload_marker() {
        let mut payload = sort(1, &[]);
        payload.extend_from_slice(&[0x98, 0x06, 0x81, 0x00]);
        let source = model(&payload);
        assert_eq!(
            decode_table_sort_order(&source, options(&source)).unwrap(),
            None
        );
        let output = rewrite_table_sort_order(&source, None, options(&source)).unwrap();
        assert_eq!(output.bytes(), source.as_slice());
    }

    #[test]
    fn unknown_overlong_scalar_and_balanced_group_round_trip() {
        let mut payload = sort(0, &[(1, 0)]);
        payload.extend_from_slice(&[0x98, 0x06, 0x81, 0x00, 0x9b, 0x06, 0x08, 0x07, 0x9c, 0x06]);
        let source = model(&payload);
        let decoded = decode_table_sort_order(&source, options(&source)).unwrap();
        assert!(decoded.is_some());
        let desired = SortOrderSnapshot::new(
            SortScope::EntireTable,
            [SortRule::new(1, SortDirection::Descending)],
        )
        .unwrap();
        let output = rewrite_table_sort_order(&source, Some(desired), options(&source)).unwrap();
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| window == [0x98, 0x06, 0x81, 0x00])
        );
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| window == [0x9b, 0x06, 0x08, 0x07])
        );
    }

    #[test]
    fn prepared_exact_limits_fail_before_output_allocation() {
        let source = model(&sort(0, &[(1, 0)]));
        let desired = SortOrderSnapshot::new(
            SortScope::EntireTable,
            [SortRule::new(2, SortDirection::Ascending)],
        )
        .unwrap();
        let prepared =
            prepare_table_sort_order_rewrite(&source, Some(desired), options(&source)).unwrap();
        let requirements = prepared.execution_requirements();
        assert_eq!(
            prepared.prepare_report().output_bytes(),
            requirements.output_bytes
        );
        assert!(
            prepared
                .clone()
                .execute(
                    requirements
                        .exact()
                        .with_output_bytes(requirements.output_bytes - 1)
                )
                .is_err()
        );
        assert!(
            prepared
                .clone()
                .execute(
                    requirements
                        .exact()
                        .with_allocations(requirements.allocations.saturating_sub(1))
                )
                .is_err()
        );
        assert!(
            prepared
                .clone()
                .execute(
                    requirements
                        .exact()
                        .with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1))
                )
                .is_err()
        );
        assert!(prepared.execute(requirements.exact()).is_ok());
    }

    #[test]
    fn malformed_known_duplicates_are_rejected() {
        let source = model(&[8, 0, 8, 1, 0x12, 0x04, 8, 0, 16, 0]);
        assert!(decode_table_sort_order(&source, options(&source)).is_err());
    }
}
