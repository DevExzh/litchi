//! Strict generated-free streaming reader for the table-local scalar formula subset.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Wire helpers stay beside the generated-free formula model."
)]

use core::{fmt, mem::size_of, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_formula_generated::LitchiIwaFormulaProjection as projection;

const MAX_DEPTH: u32 = 32;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const DECIMAL128_EXPONENT_BIAS: i32 = 0x1820;
const DECIMAL128_COEFFICIENT_BITS: u32 = 113;
const DECIMAL128_SIGN_BIT: u32 = 127;

/// Finite aggregate policy for both the preflight and callback passes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_bytes: usize,
    max_fields: usize,
    max_work: usize,
    recursion_limit: u32,
    max_nodes: usize,
    max_text_bytes: usize,
}

impl DecodeOptions {
    #[must_use]
    pub const fn new(
        max_bytes: usize,
        max_fields: usize,
        max_work: usize,
        recursion_limit: u32,
        max_nodes: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_bytes,
            max_fields,
            max_work,
            recursion_limit,
            max_nodes,
            max_text_bytes,
        }
    }
    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Caller-authorized table-local formula context.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaContext {
    owner: u32,
    host_row: u32,
    host_column: u32,
    rows: u32,
    columns: u32,
}

impl FormulaContext {
    #[must_use]
    pub const fn new(owner: u32, host_row: u32, host_column: u32, rows: u32, columns: u32) -> Self {
        Self {
            owner,
            host_row,
            host_column,
            rows,
            columns,
        }
    }
}

/// Supported postfix binary operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOperator {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Concatenate,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
}

/// One already-lowered coordinate in a canonical cell-reference node.
///
/// Relative coordinates are deltas from the formula host. Absolute
/// coordinates are zero-based table positions. Resolution and table-bound
/// validation remain the responsibility of the Numbers semantic adapter,
/// before it creates this archive-boundary value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWriteAxis {
    coordinate: i32,
    absolute: bool,
}

/// Resolved native formula-owner UID. Construction from the two canonical
/// 64-bit UUID halves keeps CFUUID word ordering inside this archive boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct FormulaWriteOwnerUid {
    lower: u64,
    upper: u64,
}

impl FormulaWriteOwnerUid {
    #[must_use]
    pub const fn from_halves(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }
}

/// One semantic cell endpoint resolved to a concrete table before lowering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWriteCellReference {
    row: u32,
    column: u32,
    row_absolute: bool,
    column_absolute: bool,
}

impl FormulaWriteCellReference {
    #[must_use]
    pub const fn new(row: u32, column: u32, row_absolute: bool, column_absolute: bool) -> Self {
        Self {
            row,
            column,
            row_absolute,
            column_absolute,
        }
    }
}

/// One semantic whole-row/whole-column endpoint.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWriteAxisReference {
    index: u32,
    absolute: bool,
}

impl FormulaWriteAxisReference {
    #[must_use]
    pub const fn new(index: u32, absolute: bool) -> Self {
        Self { index, absolute }
    }
}

/// Local host geometry and CalculationEngine owner resolved by Numbers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWriteContext {
    internal_owner: u32,
    host_row: u32,
    host_column: u32,
    rows: u32,
    columns: u32,
}

impl FormulaWriteContext {
    #[must_use]
    pub const fn new(
        internal_owner: u32,
        host_row: u32,
        host_column: u32,
        rows: u32,
        columns: u32,
    ) -> Self {
        Self {
            internal_owner,
            host_row,
            host_column,
            rows,
            columns,
        }
    }
}

/// One privately resolved external formula owner. Registries passed to the
/// planner must be strictly UID-sorted and contain no duplicate UID.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ResolvedFormulaWriteOwner {
    uid: FormulaWriteOwnerUid,
    internal_owner: u32,
    rows: u32,
    columns: u32,
}

impl ResolvedFormulaWriteOwner {
    #[must_use]
    pub const fn new(
        uid: FormulaWriteOwnerUid,
        internal_owner: u32,
        rows: u32,
        columns: u32,
    ) -> Self {
        Self {
            uid,
            internal_owner,
            rows,
            columns,
        }
    }

    #[must_use]
    pub const fn internal_owner(self) -> u32 {
        self.internal_owner
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaReadRegistryRequirements {
    owners: usize,
    retained_elements: usize,
    retained_bytes: usize,
    work: usize,
    allocations: usize,
}

impl FormulaReadRegistryRequirements {
    #[must_use]
    pub const fn owners(self) -> usize {
        self.owners
    }
    /// Total elements retained across the owner and internal-order vectors.
    #[must_use]
    pub const fn retained_elements(self) -> usize {
        self.retained_elements
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
    #[must_use]
    pub const fn work(self) -> usize {
        self.work
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

#[derive(Debug)]
pub struct FormulaReadRegistryPlan<'source> {
    owners: &'source [ResolvedFormulaWriteOwner],
    internal_order: &'source [usize],
    requirements: FormulaReadRegistryRequirements,
}

impl FormulaReadRegistryPlan<'_> {
    #[must_use]
    pub const fn requirements(&self) -> FormulaReadRegistryRequirements {
        self.requirements
    }
}

pub struct ResolvedFormulaReadRegistry {
    owners: Vec<ResolvedFormulaWriteOwner>,
    internal_order: Vec<usize>,
    requirements: FormulaReadRegistryRequirements,
}

impl ResolvedFormulaReadRegistry {
    #[must_use]
    pub const fn requirements(&self) -> FormulaReadRegistryRequirements {
        self.requirements
    }
}

pub fn plan_resolved_formula_read_registry<'source>(
    owners: &'source [ResolvedFormulaWriteOwner],
    internal_order: &'source [usize],
    options: DecodeOptions,
) -> Result<FormulaReadRegistryPlan<'source>, DecodeError> {
    if owners.len() > options.max_nodes || internal_order.len() > options.max_nodes {
        return Err(DecodeError::limited(DecodeLimit::Nodes {
            observed: owners.len().max(internal_order.len()),
            maximum: options.max_nodes,
        }));
    }
    if owners.len() != internal_order.len() {
        return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
    }
    let retained_elements = owners
        .len()
        .checked_add(internal_order.len())
        .ok_or_else(malformed)?;
    let retained_bytes = owners
        .len()
        .checked_mul(size_of::<ResolvedFormulaWriteOwner>())
        .and_then(|value| {
            internal_order
                .len()
                .checked_mul(size_of::<usize>())
                .and_then(|right| value.checked_add(right))
        })
        .ok_or_else(malformed)?;
    // Two validation scans in planning plus two retained-slice copies during
    // execution; callers authorize the composed plan+execute operation once.
    let work = owners.len().checked_mul(4).ok_or_else(malformed)?;
    if work > options.max_work {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: work,
            maximum: options.max_work,
        }));
    }
    if retained_bytes > options.max_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: retained_bytes,
            maximum: options.max_bytes,
        }));
    }
    let mut previous_uid = None;
    for owner in owners {
        if owner.uid == FormulaWriteOwnerUid::from_halves(0, 0)
            || owner.internal_owner == 0
            || owner.rows == 0
            || owner.columns == 0
            || previous_uid.is_some_and(|uid| uid >= owner.uid)
        {
            return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
        }
        previous_uid = Some(owner.uid);
    }
    let mut previous_internal = None;
    for index in internal_order.iter().copied() {
        let owner = owners
            .get(index)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
        if previous_internal.is_some_and(|value| value >= owner.internal_owner) {
            return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
        }
        previous_internal = Some(owner.internal_owner);
    }
    Ok(FormulaReadRegistryPlan {
        owners,
        internal_order,
        requirements: FormulaReadRegistryRequirements {
            owners: owners.len(),
            retained_elements,
            retained_bytes,
            work,
            allocations: usize::from(!owners.is_empty()) * 2,
        },
    })
}

pub fn execute_resolved_formula_read_registry_plan(
    plan: FormulaReadRegistryPlan<'_>,
    options: DecodeOptions,
) -> Result<ResolvedFormulaReadRegistry, DecodeError> {
    if plan.requirements.retained_bytes > options.max_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: plan.requirements.retained_bytes,
            maximum: options.max_bytes,
        }));
    }
    if plan.requirements.owners > options.max_nodes {
        return Err(DecodeError::limited(DecodeLimit::Nodes {
            observed: plan.requirements.owners,
            maximum: options.max_nodes,
        }));
    }
    if plan.requirements.work > options.max_work {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: plan.requirements.work,
            maximum: options.max_work,
        }));
    }
    let mut owners = Vec::new();
    let mut internal_order = Vec::new();
    owners
        .try_reserve_exact(plan.owners.len())
        .map_err(|_error| {
            DecodeError::limited(DecodeLimit::Allocation {
                requested: plan.owners.len(),
            })
        })?;
    internal_order
        .try_reserve_exact(plan.internal_order.len())
        .map_err(|_error| {
            DecodeError::limited(DecodeLimit::Allocation {
                requested: plan.internal_order.len(),
            })
        })?;
    if owners.capacity() != plan.owners.len()
        || internal_order.capacity() != plan.internal_order.len()
    {
        return Err(DecodeError::limited(DecodeLimit::Allocation {
            requested: owners.capacity().max(internal_order.capacity()),
        }));
    }
    owners.extend_from_slice(plan.owners);
    internal_order.extend_from_slice(plan.internal_order);
    Ok(ResolvedFormulaReadRegistry {
        owners,
        internal_order,
        requirements: plan.requirements,
    })
}

impl FormulaWriteAxis {
    #[must_use]
    pub const fn new(coordinate: i32, absolute: bool) -> Self {
        Self {
            coordinate,
            absolute,
        }
    }

    #[must_use]
    pub const fn coordinate(self) -> i32 {
        self.coordinate
    }

    #[must_use]
    pub const fn is_absolute(self) -> bool {
        self.absolute
    }
}

/// Borrowed postfix node accepted by the strict canonical formula writer.
///
/// This is an archive-boundary lowering type, not the public Numbers formula
/// vocabulary. Higher layers privately resolve cross-table owners before
/// constructing the resolved reference variants below. Pivot nodes remain
/// absent until a semantic resolver can provide an exact native shape.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FormulaWriteNode<'text> {
    Number(f64),
    Text(&'text str),
    Boolean(bool),
    CellReference {
        column: FormulaWriteAxis,
        row: FormulaWriteAxis,
    },
    ResolvedCellReference {
        owner: Option<FormulaWriteOwnerUid>,
        reference: FormulaWriteCellReference,
    },
    ResolvedRange {
        owner: Option<FormulaWriteOwnerUid>,
        start: FormulaWriteCellReference,
        end: FormulaWriteCellReference,
    },
    WholeRows {
        owner: Option<FormulaWriteOwnerUid>,
        start: FormulaWriteAxisReference,
        end: FormulaWriteAxisReference,
    },
    WholeColumns {
        owner: Option<FormulaWriteOwnerUid>,
        start: FormulaWriteAxisReference,
        end: FormulaWriteAxisReference,
    },
    Binary(BinaryOperator),
    Negation,
    Percent,
    Function {
        identifier: u32,
        argument_count: u32,
    },
    Range,
}

/// Exact resources computed without allocating an output buffer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWriteRequirements {
    output_bytes: usize,
    ast_bytes: usize,
    fields: usize,
    work_bytes: usize,
    nodes: usize,
    text_bytes: usize,
    max_depth: u32,
    allocations: usize,
    precedent_count: usize,
    range_count: usize,
}

impl FormulaWriteRequirements {
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub const fn ast_bytes(self) -> usize {
        self.ast_bytes
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
    pub const fn nodes(self) -> usize {
        self.nodes
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn precedent_count(self) -> usize {
        self.precedent_count
    }
    #[must_use]
    pub const fn range_count(self) -> usize {
        self.range_count
    }
}

/// Opaque validated formula plan borrowing the caller-authoritative nodes.
pub struct FormulaWritePlan<'nodes, 'text> {
    nodes: &'nodes [FormulaWriteNode<'text>],
    context: Option<FormulaWriteContext>,
    owners: &'nodes [ResolvedFormulaWriteOwner],
    requirements: FormulaWriteRequirements,
}

impl fmt::Debug for FormulaWritePlan<'_, '_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FormulaWritePlan")
            .field("requirements", &self.requirements)
            .field("nodes", &"<redacted>")
            .finish()
    }
}

impl FormulaWritePlan<'_, '_> {
    #[must_use]
    pub const fn requirements(&self) -> FormulaWriteRequirements {
        self.requirements
    }
}

/// Successful canonical formula-write report.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWriteReport {
    requirements: FormulaWriteRequirements,
}

impl FormulaWriteReport {
    #[must_use]
    pub const fn requirements(self) -> FormulaWriteRequirements {
        self.requirements
    }
}

/// One exact CalculationEngine precedent derived from resolved formula nodes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWritePrecedent {
    internal_owner: u32,
    row: u32,
    column: u32,
}

impl FormulaWritePrecedent {
    #[must_use]
    pub const fn internal_owner(self) -> u32 {
        self.internal_owner
    }
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

/// One logical range emitted separately from its expanded interior facts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWriteRange {
    internal_owner: u32,
    top: u32,
    left: u32,
    bottom: u32,
    right: u32,
}

impl FormulaWriteRange {
    #[must_use]
    pub const fn internal_owner(self) -> u32 {
        self.internal_owner
    }
    #[must_use]
    pub const fn top(self) -> u32 {
        self.top
    }
    #[must_use]
    pub const fn left(self) -> u32 {
        self.left
    }
    #[must_use]
    pub const fn bottom(self) -> u32 {
        self.bottom
    }
    #[must_use]
    pub const fn right(self) -> u32 {
        self.right
    }
}

/// Fallible sink for preflighted formula dependency facts. Observations must
/// be discarded if execution returns an error.
pub trait FormulaWriteDependencyVisitor {
    fn visit_precedent(&mut self, _precedent: FormulaWritePrecedent) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_range(&mut self, _range: FormulaWriteRange) -> Result<(), DecodeError> {
        Ok(())
    }
}

impl FormulaWriteDependencyVisitor for () {}

/// Independent ceilings for dependency facts derived from range expansion.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaWriteDependencyLimits {
    max_precedents: usize,
    max_ranges: usize,
}

impl FormulaWriteDependencyLimits {
    #[must_use]
    pub const fn new(max_precedents: usize, max_ranges: usize) -> Self {
        Self {
            max_precedents,
            max_ranges,
        }
    }
}

/// One validated table-local cell coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellCoordinate {
    row: u32,
    column: u32,
}

impl CellCoordinate {
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

/// One source-ordered evaluator node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaNode {
    Binary(BinaryOperator),
    Negation,
    PlusSign,
    Percent,
    Function {
        identifier: u32,
        argument_count: u32,
    },
    Number {
        bits: u64,
    },
    Boolean(bool),
    Empty,
    Token(bool),
    LocalCell {
        coordinate: CellCoordinate,
        row_is_sticky: u32,
        column_is_sticky: u32,
    },
    CellReference {
        coordinate: CellCoordinate,
    },
    ResolvedCellReference {
        owner: u32,
        coordinate: CellCoordinate,
    },
    ResolvedRange {
        owner: u32,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
    },
    Colon,
    ColonWithUids,
    AppendWhitespace,
    PrependWhitespace,
}

/// Owner-aware local precedent emitted beside a reference node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LocalPrecedent {
    owner: u32,
    coordinate: CellCoordinate,
}

/// Canonical table-local node that the scalar evaluator cannot execute.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UnsupportedLocal {
    node_type: u32,
    function_identifier: Option<u32>,
}

impl UnsupportedLocal {
    #[must_use]
    pub const fn node_type(self) -> u32 {
        self.node_type
    }
    #[must_use]
    pub const fn function_identifier(self) -> Option<u32> {
        self.function_identifier
    }
}

impl LocalPrecedent {
    #[must_use]
    pub const fn owner(self) -> u32 {
        self.owner
    }
    #[must_use]
    pub const fn coordinate(self) -> CellCoordinate {
        self.coordinate
    }
}

/// Fallible source-order sink. Observations must be discarded on decode error.
pub trait FormulaVisitor {
    fn visit_node(&mut self, _node: FormulaNode) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_precedent(&mut self, _precedent: LocalPrecedent) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_range(&mut self, _range: FormulaWriteRange) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_text(&mut self, _value: &str) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_unsupported_local(&mut self, _node: UnsupportedLocal) -> Result<(), DecodeError> {
        Ok(())
    }
}

impl FormulaVisitor for () {}

/// Dependency-only sink. Supported evaluator nodes are intentionally omitted.
pub trait FormulaDependencyVisitor {
    fn visit_precedent(&mut self, _precedent: LocalPrecedent) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_range(&mut self, _range: FormulaWriteRange) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_node(&mut self, _node: FormulaNode) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_text(&mut self, _value: &str) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_unsupported_local(&mut self, _node: UnsupportedLocal) -> Result<(), DecodeError> {
        Ok(())
    }
}

impl FormulaDependencyVisitor for () {}

/// Exact successful aggregate report.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    bytes: usize,
    fields: usize,
    work: usize,
    max_depth: u32,
    text_bytes: usize,
    node_count: usize,
    precedent_count: usize,
    range_count: usize,
    evaluator_supported: bool,
    unsupported_local_count: usize,
    allocations: usize,
}

macro_rules! accessors {
    ($(($name:ident, $ty:ty)),+ $(,)?) => {$(
        #[must_use]
        pub const fn $name(self) -> $ty { self.$name }
    )+};
}

impl DecodeReport {
    accessors!(
        (bytes, usize),
        (fields, usize),
        (work, usize),
        (max_depth, u32),
        (text_bytes, usize),
        (node_count, usize),
        (precedent_count, usize),
        (range_count, usize),
        (evaluator_supported, bool),
        (unsupported_local_count, usize),
        (allocations, usize)
    );
}

/// Typed aggregate refusal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    Bytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    Nodes { observed: usize, maximum: usize },
    Text { observed: usize, maximum: usize },
    Allocation { requested: usize },
}

/// Content-free malformed or unsupported classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InvalidReason {
    MalformedWire,
    MissingRequired,
    DuplicateField,
    UnexpectedField,
    InvalidCoordinate,
    UnsupportedFormula,
    ExternalOwner,
    InvalidPostfix,
    InvalidScalar,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
    reason: Option<InvalidReason>,
}

impl DecodeError {
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        self.limit
    }
    #[must_use]
    pub const fn invalid_reason(&self) -> Option<InvalidReason> {
        self.reason
    }
    const fn invalid(reason: InvalidReason) -> Self {
        Self {
            limit: None,
            reason: Some(reason),
        }
    }
    const fn limited(limit: DecodeLimit) -> Self {
        Self {
            limit: Some(limit),
            reason: None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid or unsupported FormulaArchive")
    }
}

impl std::error::Error for DecodeError {}

/// Strict allocation-free inspection for sizing an evaluator stack before
/// requesting callbacks. Its report covers this one complete validation pass.
pub fn inspect_formula_archive(
    source: &[u8],
    context: FormulaContext,
    options: DecodeOptions,
) -> Result<DecodeReport, DecodeError> {
    validate_context(context)?;
    let mut budget = Budget::new(source, options)?;
    decode_formula(
        source,
        context,
        None,
        &mut budget,
        &mut (),
        false,
        DecodeMode::Evaluator,
    )?;
    Ok(budget.report())
}

/// Strict dependency-only inspection for global AST-to-edge proof.
///
/// Canonical evaluator-unsupported local nodes are tagged rather than refused;
/// malformed, noncanonical, and external-owner forms remain hard failures.
pub fn inspect_formula_dependencies_with_visitor<V: FormulaDependencyVisitor>(
    source: &[u8],
    context: FormulaContext,
    options: DecodeOptions,
    visitor: &mut V,
) -> Result<DecodeReport, DecodeError> {
    validate_context(context)?;
    let mut budget = Budget::new(source, options)?;
    decode_formula(
        source,
        context,
        None,
        &mut budget,
        &mut (),
        false,
        DecodeMode::Dependencies,
    )?;
    budget.preflight_callback_pass()?;
    let mut adapter = DependencyAdapter(visitor);
    decode_formula(
        source,
        context,
        None,
        &mut budget,
        &mut adapter,
        true,
        DecodeMode::Dependencies,
    )?;
    Ok(budget.report())
}

struct DependencyAdapter<'visitor, V: ?Sized>(&'visitor mut V);

impl<V: FormulaDependencyVisitor + ?Sized> FormulaVisitor for DependencyAdapter<'_, V> {
    fn visit_node(&mut self, node: FormulaNode) -> Result<(), DecodeError> {
        self.0.visit_node(node)
    }
    fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
        self.0.visit_precedent(precedent)
    }
    fn visit_range(&mut self, range: FormulaWriteRange) -> Result<(), DecodeError> {
        self.0.visit_range(range)
    }
    fn visit_text(&mut self, value: &str) -> Result<(), DecodeError> {
        self.0.visit_text(value)
    }
    fn visit_unsupported_local(&mut self, node: UnsupportedLocal) -> Result<(), DecodeError> {
        self.0.visit_unsupported_local(node)
    }
}

/// Strictly preflight then stream one table-local FormulaArchive.
pub fn decode_formula_archive_with_visitor<V: FormulaVisitor>(
    source: &[u8],
    context: FormulaContext,
    options: DecodeOptions,
    visitor: &mut V,
) -> Result<DecodeReport, DecodeError> {
    validate_context(context)?;
    let mut budget = Budget::new(source, options)?;
    decode_formula(
        source,
        context,
        None,
        &mut budget,
        &mut (),
        false,
        DecodeMode::Evaluator,
    )?;
    budget.preflight_callback_pass()?;
    decode_formula(
        source,
        context,
        None,
        &mut budget,
        visitor,
        true,
        DecodeMode::Evaluator,
    )?;
    Ok(budget.report())
}

/// Strictly inspect dependencies in a writer-canonical archive whose table
/// owner identities were resolved by the Numbers package graph.
pub fn inspect_resolved_formula_dependencies_with_visitor<V: FormulaDependencyVisitor + ?Sized>(
    source: &[u8],
    context: FormulaWriteContext,
    registry: &ResolvedFormulaReadRegistry,
    options: DecodeOptions,
    visitor: &mut V,
) -> Result<DecodeReport, DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let resolution = validate_read_resolution(context, registry, &mut budget)?;
    let local = resolution.local_context();
    decode_formula(
        source,
        local,
        Some(&resolution),
        &mut budget,
        &mut (),
        false,
        DecodeMode::Dependencies,
    )?;
    budget.preflight_callback_pass()?;
    let mut adapter = DependencyAdapter(visitor);
    decode_formula(
        source,
        local,
        Some(&resolution),
        &mut budget,
        &mut adapter,
        true,
        DecodeMode::Dependencies,
    )?;
    Ok(budget.report())
}

/// Strictly preflight and stream a writer-canonical owner-resolved archive.
pub fn decode_resolved_formula_archive_with_visitor<V: FormulaVisitor + ?Sized>(
    source: &[u8],
    context: FormulaWriteContext,
    registry: &ResolvedFormulaReadRegistry,
    options: DecodeOptions,
    visitor: &mut V,
) -> Result<DecodeReport, DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let resolution = validate_read_resolution(context, registry, &mut budget)?;
    let local = resolution.local_context();
    decode_formula(
        source,
        local,
        Some(&resolution),
        &mut budget,
        &mut (),
        false,
        DecodeMode::Evaluator,
    )?;
    budget.preflight_callback_pass()?;
    decode_formula(
        source,
        local,
        Some(&resolution),
        &mut budget,
        visitor,
        true,
        DecodeMode::Evaluator,
    )?;
    Ok(budget.report())
}

struct ReadResolution<'owners> {
    context: FormulaWriteContext,
    registry: &'owners ResolvedFormulaReadRegistry,
}

impl ReadResolution<'_> {
    const fn local_context(&self) -> FormulaContext {
        FormulaContext::new(
            self.context.internal_owner,
            self.context.host_row,
            self.context.host_column,
            self.context.rows,
            self.context.columns,
        )
    }
}

fn validate_read_resolution<'registry>(
    context: FormulaWriteContext,
    registry: &'registry ResolvedFormulaReadRegistry,
    budget: &mut Budget,
) -> Result<ReadResolution<'registry>, DecodeError> {
    if context.internal_owner == 0
        || context.rows == 0
        || context.columns == 0
        || context.host_row >= context.rows
        || context.host_column >= context.columns
    {
        return Err(DecodeError::invalid(InvalidReason::InvalidCoordinate));
    }
    let local = registry_owner_by_internal(registry, context.internal_owner, budget)?;
    if local.rows != context.rows || local.columns != context.columns {
        return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
    }
    Ok(ReadResolution { context, registry })
}

fn registry_owner_by_internal(
    registry: &ResolvedFormulaReadRegistry,
    internal_owner: u32,
    budget: &mut Budget,
) -> Result<ResolvedFormulaWriteOwner, DecodeError> {
    let mut left = 0usize;
    let mut right = registry.internal_order.len();
    while left < right {
        let middle = left + (right - left) / 2;
        budget.one_time_work(1)?;
        let index = registry.internal_order[middle];
        let owner = registry.owners[index];
        match owner.internal_owner.cmp(&internal_owner) {
            core::cmp::Ordering::Less => left = middle + 1,
            core::cmp::Ordering::Greater => right = middle,
            core::cmp::Ordering::Equal => return Ok(owner),
        }
    }
    Err(DecodeError::invalid(InvalidReason::ExternalOwner))
}

/// Validate and exactly size a canonical FormulaArchive without allocating.
///
/// The input must already be a complete postfix expression. The plan borrows
/// that slice, making mutation between planning and execution impossible in
/// safe Rust. Only root field 1 is emitted; optional host/translation fields
/// are contextual metadata and are not invented by this focused authoring
/// seam.
pub fn plan_formula_archive<'nodes, 'text>(
    nodes: &'nodes [FormulaWriteNode<'text>],
    options: DecodeOptions,
) -> Result<FormulaWritePlan<'nodes, 'text>, DecodeError> {
    plan_formula_archive_in(
        nodes,
        None,
        &[],
        FormulaWriteDependencyLimits::new(0, 0),
        options,
    )
}

/// Plan canonical authoring with privately resolved owner geometry.
pub fn plan_resolved_formula_archive<'nodes, 'text>(
    nodes: &'nodes [FormulaWriteNode<'text>],
    context: FormulaWriteContext,
    owners: &'nodes [ResolvedFormulaWriteOwner],
    dependency_limits: FormulaWriteDependencyLimits,
    options: DecodeOptions,
) -> Result<FormulaWritePlan<'nodes, 'text>, DecodeError> {
    plan_formula_archive_in(nodes, Some(context), owners, dependency_limits, options)
}

fn plan_formula_archive_in<'nodes, 'text>(
    nodes: &'nodes [FormulaWriteNode<'text>],
    context: Option<FormulaWriteContext>,
    owners: &'nodes [ResolvedFormulaWriteOwner],
    dependency_limits: FormulaWriteDependencyLimits,
    options: DecodeOptions,
) -> Result<FormulaWritePlan<'nodes, 'text>, DecodeError> {
    if options.recursion_limit == 0
        || options.recursion_limit > MAX_DEPTH
        || options.recursion_limit < 3
    {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: 3,
            maximum: options.recursion_limit.min(MAX_DEPTH),
        }));
    }
    if options.max_fields < 1 {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: 1,
            maximum: options.max_fields,
        }));
    }
    if nodes.len() > options.max_nodes {
        return Err(DecodeError::limited(DecodeLimit::Nodes {
            observed: nodes.len(),
            maximum: options.max_nodes,
        }));
    }
    if owners.len() > options.max_nodes {
        return Err(DecodeError::limited(DecodeLimit::Nodes {
            observed: owners.len(),
            maximum: options.max_nodes,
        }));
    }
    let owner_work = owners
        .len()
        .checked_mul(owners.len())
        .ok_or_else(malformed)?;
    if owner_work > options.max_work {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: owner_work,
            maximum: options.max_work,
        }));
    }
    validate_write_context(context, owners)?;
    let mut stack = 0usize;
    let mut ast_bytes = 0usize;
    let mut fields = 1usize;
    let mut text_bytes = 0usize;
    let mut plan_work_bytes = owner_work;
    let mut max_depth = 3u32;
    let maximum_output = options
        .max_bytes
        .min(usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| malformed())?);

    for node in nodes {
        if matches!(node, FormulaWriteNode::CellReference { .. }) && options.recursion_limit < 4 {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: 4,
                maximum: options.recursion_limit,
            }));
        }
        plan_work_bytes = checked(plan_work_bytes, 8)?;
        if let FormulaWriteNode::Text(text) = node {
            text_bytes = checked(text_bytes, text.len())?;
            if text_bytes > options.max_text_bytes {
                return Err(DecodeError::limited(DecodeLimit::Text {
                    observed: text_bytes,
                    maximum: options.max_text_bytes,
                }));
            }
            plan_work_bytes = checked(plan_work_bytes, text.len())?;
        }
        if plan_work_bytes > options.max_work {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: plan_work_bytes,
                maximum: options.max_work,
            }));
        }
        validate_write_node(node, &mut stack, context, owners)?;
        let wire = wire_node_requirements(node, context, owners)?;
        ast_bytes = checked(ast_bytes, wire.ast_bytes)?;
        fields = checked(fields, wire.fields)?;
        max_depth = max_depth.max(wire.max_depth);
        if fields > options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: fields,
                maximum: options.max_fields,
            }));
        }
        let current_output = length_delimited_field_len(1, ast_bytes)?;
        if current_output > maximum_output {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: current_output,
                maximum: maximum_output,
            }));
        }
    }
    if nodes.is_empty() || stack != 1 {
        return Err(DecodeError::invalid(InvalidReason::InvalidPostfix));
    }
    let output_bytes = length_delimited_field_len(1, ast_bytes)?;
    let (precedent_count, range_count) = formula_dependency_counts(nodes, context, owners)?;
    let dependency_work = checked(nodes.len(), checked(precedent_count, range_count)?)?;
    if precedent_count > dependency_limits.max_precedents {
        return Err(DecodeError::limited(DecodeLimit::Nodes {
            observed: precedent_count,
            maximum: dependency_limits.max_precedents,
        }));
    }
    if range_count > dependency_limits.max_ranges {
        return Err(DecodeError::limited(DecodeLimit::Nodes {
            observed: range_count,
            maximum: dependency_limits.max_ranges,
        }));
    }
    // Encoding, strict raw readback, and private Buffa parity each make at
    // most one linear pass over the canonical bytes. The small per-node term
    // covers stack validation and exact-size arithmetic.
    let execute_work_bytes = output_bytes
        .checked_mul(3)
        .and_then(|work| nodes.len().checked_mul(8).and_then(|n| work.checked_add(n)))
        .ok_or_else(malformed)?;
    let work_bytes = checked(
        checked(plan_work_bytes, execute_work_bytes)?,
        dependency_work,
    )?;
    let requirements = FormulaWriteRequirements {
        output_bytes,
        ast_bytes,
        fields,
        work_bytes,
        nodes: nodes.len(),
        text_bytes,
        max_depth,
        allocations: 1,
        precedent_count,
        range_count,
    };
    validate_write_requirements(requirements, options)?;
    Ok(FormulaWritePlan {
        nodes,
        context,
        owners,
        requirements,
    })
}

/// Execute a validated canonical formula plan with one fallible exact reserve.
///
/// Production encoding is handwritten. Generated types are never materialized
/// and Prost is not used; private Buffa lazy views independently cross-check
/// the completed borrowed bytes before they are returned.
pub fn execute_formula_archive_plan(
    plan: FormulaWritePlan<'_, '_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, FormulaWriteReport), DecodeError> {
    execute_formula_archive_plan_with_visitor(plan, options, &mut ())
}

/// Execute a plan and stream every preflighted owner-qualified dependency.
pub fn execute_formula_archive_plan_with_visitor(
    plan: FormulaWritePlan<'_, '_>,
    options: DecodeOptions,
    visitor: &mut dyn FormulaWriteDependencyVisitor,
) -> Result<(Vec<u8>, FormulaWriteReport), DecodeError> {
    validate_write_requirements(plan.requirements, options)?;
    let mut output = reserve_formula_output(plan.requirements.output_bytes)?;
    put_key(&mut output, 1, 2);
    put_varint(
        &mut output,
        u64::try_from(plan.requirements.ast_bytes).map_err(|_error| malformed())?,
    );
    for node in plan.nodes {
        let node_bytes = encoded_write_node_len_in(node, plan.context, plan.owners)?;
        put_key(&mut output, 1, 2);
        put_varint(
            &mut output,
            u64::try_from(node_bytes).map_err(|_error| malformed())?,
        );
        encode_write_node_in(&mut output, node, plan.context, plan.owners)?;
    }
    if output.len() != plan.requirements.output_bytes {
        return Err(malformed());
    }
    validate_canonical_formula_output(&output, plan.nodes, plan.context, plan.owners, options)?;
    emit_formula_dependencies(plan.nodes, plan.context, plan.owners, visitor)?;
    Ok((
        output,
        FormulaWriteReport {
            requirements: plan.requirements,
        },
    ))
}

/// Convenience wrapper around [`plan_formula_archive`] and
/// [`execute_formula_archive_plan`].
pub fn encode_formula_archive(
    nodes: &[FormulaWriteNode<'_>],
    options: DecodeOptions,
) -> Result<(Vec<u8>, FormulaWriteReport), DecodeError> {
    execute_formula_archive_plan(plan_formula_archive(nodes, options)?, options)
}

fn validate_write_node(
    node: &FormulaWriteNode<'_>,
    stack: &mut usize,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<(), DecodeError> {
    match node {
        FormulaWriteNode::Number(value) => {
            if !value.is_finite() || (value.is_sign_negative() && *value != 0.0) {
                return Err(DecodeError::invalid(InvalidReason::InvalidScalar));
            }
            *stack = checked(*stack, 1)?;
        },
        FormulaWriteNode::Text(text) => {
            if text.contains('\0') {
                return Err(DecodeError::invalid(InvalidReason::InvalidScalar));
            }
            *stack = checked(*stack, 1)?;
        },
        FormulaWriteNode::Boolean(_) | FormulaWriteNode::CellReference { .. } => {
            *stack = checked(*stack, 1)?;
        },
        FormulaWriteNode::ResolvedCellReference { owner, reference } => {
            let target = resolved_target(context, owners, *owner)?;
            validate_cell_endpoint(*reference, target)?;
            *stack = checked(*stack, 1)?;
        },
        FormulaWriteNode::ResolvedRange { owner, start, end } => {
            let target = resolved_target(context, owners, *owner)?;
            validate_cell_endpoint(*start, target)?;
            validate_cell_endpoint(*end, target)?;
            validate_endpoint_order(*start, *end)?;
            *stack = checked(*stack, 1)?;
        },
        FormulaWriteNode::WholeRows { owner, start, end } => {
            let target = resolved_target(context, owners, *owner)?;
            validate_axis_endpoint(*start, target.rows)?;
            validate_axis_endpoint(*end, target.rows)?;
            validate_axis_order(*start, *end)?;
            *stack = checked(*stack, 1)?;
        },
        FormulaWriteNode::WholeColumns { owner, start, end } => {
            let target = resolved_target(context, owners, *owner)?;
            validate_axis_endpoint(*start, target.columns)?;
            validate_axis_endpoint(*end, target.columns)?;
            validate_axis_order(*start, *end)?;
            *stack = checked(*stack, 1)?;
        },
        FormulaWriteNode::Binary(_) | FormulaWriteNode::Range => {
            if *stack < 2 {
                return Err(DecodeError::invalid(InvalidReason::InvalidPostfix));
            }
            *stack -= 1;
        },
        FormulaWriteNode::Negation | FormulaWriteNode::Percent => {
            if *stack == 0 {
                return Err(DecodeError::invalid(InvalidReason::InvalidPostfix));
            }
        },
        FormulaWriteNode::Function {
            identifier,
            argument_count,
        } => {
            if *identifier == 0 {
                return Err(DecodeError::invalid(InvalidReason::InvalidScalar));
            }
            let arguments = usize::try_from(*argument_count).map_err(|_error| malformed())?;
            if *stack < arguments {
                return Err(DecodeError::invalid(InvalidReason::InvalidPostfix));
            }
            *stack = checked(*stack - arguments, 1)?;
        },
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct WriteTarget {
    internal_owner: u32,
    uid: Option<FormulaWriteOwnerUid>,
    rows: u32,
    columns: u32,
}

fn validate_write_context(
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<(), DecodeError> {
    let Some(context) = context else {
        if owners.is_empty() {
            return Ok(());
        }
        return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
    };
    if context.internal_owner == 0
        || context.rows == 0
        || context.columns == 0
        || context.host_row >= context.rows
        || context.host_column >= context.columns
    {
        return Err(DecodeError::invalid(InvalidReason::InvalidCoordinate));
    }
    let mut previous = None;
    for owner in owners {
        if owner.internal_owner == 0
            || owner.internal_owner == context.internal_owner
            || owner.rows == 0
            || owner.columns == 0
            || owner.uid == FormulaWriteOwnerUid::from_halves(0, 0)
            || previous.is_some_and(|uid| uid >= owner.uid)
        {
            return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
        }
        previous = Some(owner.uid);
    }
    for (index, owner) in owners.iter().enumerate() {
        if owners[index + 1..]
            .iter()
            .any(|candidate| candidate.internal_owner == owner.internal_owner)
        {
            return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
        }
    }
    Ok(())
}

fn resolved_target(
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
    uid: Option<FormulaWriteOwnerUid>,
) -> Result<WriteTarget, DecodeError> {
    let context = context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
    match uid {
        None => Ok(WriteTarget {
            internal_owner: context.internal_owner,
            uid: None,
            rows: context.rows,
            columns: context.columns,
        }),
        Some(uid) => {
            let position = owners
                .binary_search_by_key(&uid, |owner| owner.uid)
                .map_err(|_error| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            let owner = owners[position];
            Ok(WriteTarget {
                internal_owner: owner.internal_owner,
                uid: Some(uid),
                rows: owner.rows,
                columns: owner.columns,
            })
        },
    }
}

fn validate_cell_endpoint(
    reference: FormulaWriteCellReference,
    target: WriteTarget,
) -> Result<(), DecodeError> {
    if reference.row >= target.rows || reference.column >= target.columns {
        Err(DecodeError::invalid(InvalidReason::InvalidCoordinate))
    } else {
        Ok(())
    }
}

fn validate_axis_endpoint(
    reference: FormulaWriteAxisReference,
    maximum: u32,
) -> Result<(), DecodeError> {
    if reference.index >= maximum {
        Err(DecodeError::invalid(InvalidReason::InvalidCoordinate))
    } else {
        Ok(())
    }
}

fn validate_endpoint_order(
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
) -> Result<(), DecodeError> {
    if start.row > end.row || start.column > end.column {
        Err(DecodeError::invalid(InvalidReason::InvalidCoordinate))
    } else {
        Ok(())
    }
}

fn validate_axis_order(
    start: FormulaWriteAxisReference,
    end: FormulaWriteAxisReference,
) -> Result<(), DecodeError> {
    if start.index > end.index {
        Err(DecodeError::invalid(InvalidReason::InvalidCoordinate))
    } else {
        Ok(())
    }
}

struct WireNodeRequirements {
    ast_bytes: usize,
    fields: usize,
    max_depth: u32,
}

fn wire_node_requirements(
    node: &FormulaWriteNode<'_>,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<WireNodeRequirements, DecodeError> {
    let node_bytes = encoded_write_node_len_in(node, context, owners)?;
    let direct_fields = match node {
        FormulaWriteNode::ResolvedCellReference { owner, .. }
        | FormulaWriteNode::ResolvedRange { owner, .. }
        | FormulaWriteNode::WholeRows { owner, .. }
        | FormulaWriteNode::WholeColumns { owner, .. } => 3 + usize::from(owner.is_some()),
        _ => write_node_field_count(node),
    };
    let nested_fields = match node {
        FormulaWriteNode::CellReference { .. } => 4,
        FormulaWriteNode::ResolvedCellReference { owner, .. } => {
            if owner.is_some() {
                9
            } else {
                4
            }
        },
        FormulaWriteNode::ResolvedRange { owner, start, end } => {
            let context =
                context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            colon_nested_field_count(*owner, *start, *end, context)?
        },
        FormulaWriteNode::WholeRows { owner, start, end } => {
            let context =
                context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            whole_axis_nested_field_count(*owner, *start, *end, true, context)?
        },
        FormulaWriteNode::WholeColumns { owner, start, end } => {
            let context =
                context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            whole_axis_nested_field_count(*owner, *start, *end, false, context)?
        },
        _ => 0,
    };
    let resolved_deep = matches!(
        node,
        FormulaWriteNode::ResolvedCellReference { owner: Some(_), .. }
            | FormulaWriteNode::ResolvedRange { .. }
            | FormulaWriteNode::WholeRows { .. }
            | FormulaWriteNode::WholeColumns { .. }
    );
    Ok(WireNodeRequirements {
        ast_bytes: length_delimited_field_len(1, node_bytes)?,
        fields: checked(checked(1, direct_fields)?, nested_fields)?,
        max_depth: if matches!(
            node,
            FormulaWriteNode::CellReference { .. }
                | FormulaWriteNode::ResolvedCellReference { .. }
                | FormulaWriteNode::ResolvedRange { .. }
                | FormulaWriteNode::WholeRows { .. }
                | FormulaWriteNode::WholeColumns { .. }
        ) {
            if resolved_deep { 5 } else { 4 }
        } else {
            3
        },
    })
}

fn formula_dependency_counts(
    nodes: &[FormulaWriteNode<'_>],
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<(usize, usize), DecodeError> {
    let mut precedents = 0usize;
    let mut ranges = 0usize;
    for node in nodes {
        let count = match node {
            FormulaWriteNode::ResolvedCellReference { .. } => 1,
            FormulaWriteNode::ResolvedRange { owner, start, end } => {
                let target = resolved_target(context, owners, *owner)?;
                let _internal_owner = target.internal_owner;
                validate_cell_endpoint(*start, target)?;
                validate_cell_endpoint(*end, target)?;
                ranges = checked(ranges, 1)?;
                rectangular_cell_count(*start, *end)?
            },
            FormulaWriteNode::WholeRows { owner, start, end } => {
                let target = resolved_target(context, owners, *owner)?;
                let _internal_owner = target.internal_owner;
                ranges = checked(ranges, 1)?;
                axis_cell_count(*start, *end, target.columns)?
            },
            FormulaWriteNode::WholeColumns { owner, start, end } => {
                let target = resolved_target(context, owners, *owner)?;
                let _internal_owner = target.internal_owner;
                ranges = checked(ranges, 1)?;
                axis_cell_count(*start, *end, target.rows)?
            },
            _ => 0,
        };
        precedents = checked(precedents, count)?;
    }
    Ok((precedents, ranges))
}

fn emit_formula_dependencies(
    nodes: &[FormulaWriteNode<'_>],
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
    visitor: &mut dyn FormulaWriteDependencyVisitor,
) -> Result<(), DecodeError> {
    for node in nodes {
        match node {
            FormulaWriteNode::ResolvedCellReference { owner, reference } => {
                let target = resolved_target(context, owners, *owner)?;
                visitor.visit_precedent(FormulaWritePrecedent {
                    internal_owner: target.internal_owner,
                    row: reference.row,
                    column: reference.column,
                })?;
            },
            FormulaWriteNode::ResolvedRange { owner, start, end } => {
                let target = resolved_target(context, owners, *owner)?;
                emit_rectangle(target, *start, *end, visitor)?;
            },
            FormulaWriteNode::WholeRows { owner, start, end } => {
                let target = resolved_target(context, owners, *owner)?;
                let start_cell =
                    FormulaWriteCellReference::new(start.index, 0, start.absolute, false);
                let end_cell = FormulaWriteCellReference::new(
                    end.index,
                    target.columns - 1,
                    end.absolute,
                    false,
                );
                emit_rectangle(target, start_cell, end_cell, visitor)?;
            },
            FormulaWriteNode::WholeColumns { owner, start, end } => {
                let target = resolved_target(context, owners, *owner)?;
                let start_cell =
                    FormulaWriteCellReference::new(0, start.index, false, start.absolute);
                let end_cell =
                    FormulaWriteCellReference::new(target.rows - 1, end.index, false, end.absolute);
                emit_rectangle(target, start_cell, end_cell, visitor)?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn emit_rectangle(
    target: WriteTarget,
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
    visitor: &mut dyn FormulaWriteDependencyVisitor,
) -> Result<(), DecodeError> {
    if start.row > end.row || start.column > end.column {
        return Err(DecodeError::invalid(InvalidReason::InvalidCoordinate));
    }
    visitor.visit_range(FormulaWriteRange {
        internal_owner: target.internal_owner,
        top: start.row,
        left: start.column,
        bottom: end.row,
        right: end.column,
    })?;
    for row in start.row..=end.row {
        for column in start.column..=end.column {
            visitor.visit_precedent(FormulaWritePrecedent {
                internal_owner: target.internal_owner,
                row,
                column,
            })?;
        }
    }
    Ok(())
}

fn rectangular_cell_count(
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
) -> Result<usize, DecodeError> {
    let rows = usize::try_from(start.row.abs_diff(end.row))
        .map_err(|_error| malformed())?
        .checked_add(1)
        .ok_or_else(malformed)?;
    let columns = usize::try_from(start.column.abs_diff(end.column))
        .map_err(|_error| malformed())?
        .checked_add(1)
        .ok_or_else(malformed)?;
    rows.checked_mul(columns).ok_or_else(malformed)
}

fn axis_cell_count(
    start: FormulaWriteAxisReference,
    end: FormulaWriteAxisReference,
    perpendicular: u32,
) -> Result<usize, DecodeError> {
    let selected = usize::try_from(start.index.abs_diff(end.index))
        .map_err(|_error| malformed())?
        .checked_add(1)
        .ok_or_else(malformed)?;
    selected
        .checked_mul(usize::try_from(perpendicular).map_err(|_error| malformed())?)
        .ok_or_else(malformed)
}

fn lowered_axis(target: u32, host: u32, absolute: bool) -> Result<FormulaWriteAxis, DecodeError> {
    let coordinate = if absolute {
        i64::from(target)
    } else {
        i64::from(target) - i64::from(host)
    };
    Ok(FormulaWriteAxis::new(
        i32::try_from(coordinate)
            .map_err(|_error| DecodeError::invalid(InvalidReason::InvalidCoordinate))?,
        absolute,
    ))
}

fn formula_decimal128_parts(value: f64) -> Result<(u64, u64), DecodeError> {
    if !value.is_finite() {
        return Err(DecodeError::invalid(InvalidReason::InvalidScalar));
    }
    let magnitude = value.abs();
    let mut buffer = ryu::Buffer::new();
    let spelling = if magnitude == 0.0 {
        "0"
    } else {
        buffer.format_finite(magnitude)
    };
    let (mantissa, explicit) = spelling
        .split_once(['e', 'E'])
        .map_or((spelling, 0), |(mantissa, exponent)| {
            (mantissa, exponent.parse::<i32>().unwrap_or(i32::MIN))
        });
    if explicit == i32::MIN {
        return Err(DecodeError::invalid(InvalidReason::InvalidScalar));
    }
    let fractional = mantissa
        .split_once('.')
        .map_or(0usize, |(_, value)| value.len());
    let mut coefficient = 0u128;
    let mut trailing = 0i32;
    for byte in mantissa.bytes() {
        if byte == b'.' {
            continue;
        }
        let digit = byte
            .checked_sub(b'0')
            .filter(|digit| *digit <= 9)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::InvalidScalar))?;
        coefficient = coefficient
            .checked_mul(10)
            .and_then(|value| value.checked_add(u128::from(digit)))
            .ok_or_else(malformed)?;
        trailing = if digit == 0 {
            trailing.checked_add(1).ok_or_else(malformed)?
        } else {
            0
        };
    }
    if coefficient == 0 {
        trailing = 0;
    } else {
        for _ in 0..trailing {
            coefficient /= 10;
        }
    }
    if coefficient >= (1u128 << DECIMAL128_COEFFICIENT_BITS) {
        return Err(DecodeError::invalid(InvalidReason::InvalidScalar));
    }
    let exponent = explicit
        .checked_sub(i32::try_from(fractional).map_err(|_error| malformed())?)
        .and_then(|value| value.checked_add(trailing))
        .ok_or_else(malformed)?;
    let biased = exponent
        .checked_add(DECIMAL128_EXPONENT_BIAS)
        .filter(|value| (0..=0x3fff).contains(value))
        .ok_or_else(|| DecodeError::invalid(InvalidReason::InvalidScalar))?;
    let encoded = coefficient
        | (u128::try_from(biased).map_err(|_error| malformed())? << DECIMAL128_COEFFICIENT_BITS);
    let encoded = if value.is_sign_negative() {
        encoded | (1u128 << DECIMAL128_SIGN_BIT)
    } else {
        encoded
    };
    Ok((encoded as u64, (encoded >> 64) as u64))
}

fn cfuuid_len(uid: FormulaWriteOwnerUid) -> Result<usize, DecodeError> {
    [
        uid.lower as u32,
        (uid.lower >> 32) as u32,
        uid.upper as u32,
        (uid.upper >> 32) as u32,
    ]
    .into_iter()
    .enumerate()
    .try_fold(0usize, |length, (index, word)| {
        checked(
            length,
            varint_field_len(
                u32::try_from(index + 2).map_err(|_error| malformed())?,
                u64::from(word),
            )?,
        )
    })
}

fn cross_extra_len(uid: FormulaWriteOwnerUid) -> Result<usize, DecodeError> {
    length_delimited_field_len(1, cfuuid_len(uid)?)
}

fn sticky_len() -> Result<usize, DecodeError> {
    Ok(varint_field_len(1, 0)?
        + varint_field_len(2, 0)?
        + varint_field_len(3, 0)?
        + varint_field_len(4, 0)?)
}

fn relative_range_len(begin: i32, end: Option<i32>) -> Result<usize, DecodeError> {
    let mut length = varint_field_len(1, begin as u64)?;
    if let Some(end) = end {
        length = checked(length, varint_field_len(2, end as u64)?)?;
    }
    Ok(length)
}

fn absolute_range_len(begin: u32, end: Option<u32>) -> Result<usize, DecodeError> {
    let mut length = varint_field_len(1, u64::from(begin))?;
    if let Some(end) = end {
        length = checked(length, varint_field_len(2, u64::from(end))?)?;
    }
    Ok(length)
}

fn colon_tract_node_len(
    base: usize,
    owner: Option<FormulaWriteOwnerUid>,
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
    _axis: bool,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<usize, DecodeError> {
    let _target = resolved_target(context, owners, owner)?;
    let context = context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
    let tract = colon_tract_len(start, end, context)?;
    let mut length = checked(
        checked(base, length_delimited_field_len(33, sticky_len()?)?)?,
        length_delimited_field_len(40, tract)?,
    )?;
    if let Some(uid) = owner {
        length = checked(
            length,
            length_delimited_field_len(28, cross_extra_len(uid)?)?,
        )?;
    }
    Ok(length)
}

fn whole_axis_node_len(
    base: usize,
    owner: Option<FormulaWriteOwnerUid>,
    start: FormulaWriteAxisReference,
    end: FormulaWriteAxisReference,
    rows: bool,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<usize, DecodeError> {
    let _target = resolved_target(context, owners, owner)?;
    let context = context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
    let tract = whole_axis_tract_len(start, end, rows, context)?;
    let mut length = checked(
        checked(base, length_delimited_field_len(33, sticky_len()?)?)?,
        length_delimited_field_len(40, tract)?,
    )?;
    if let Some(uid) = owner {
        length = checked(
            length,
            length_delimited_field_len(28, cross_extra_len(uid)?)?,
        )?;
    }
    Ok(length)
}

fn colon_tract_len(
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
    context: FormulaWriteContext,
) -> Result<usize, DecodeError> {
    let mut length = 0usize;
    length = checked(
        length,
        axis_ranges_len(
            start.column,
            end.column,
            start.column_absolute,
            end.column_absolute,
            context.host_column,
            1,
            3,
        )?,
    )?;
    length = checked(
        length,
        axis_ranges_len(
            start.row,
            end.row,
            start.row_absolute,
            end.row_absolute,
            context.host_row,
            2,
            4,
        )?,
    )?;
    checked(length, varint_field_len(5, 1)?)
}

fn whole_axis_tract_len(
    start: FormulaWriteAxisReference,
    end: FormulaWriteAxisReference,
    rows: bool,
    context: FormulaWriteContext,
) -> Result<usize, DecodeError> {
    let (host, relative_field, absolute_field, sentinel_field, sentinel) = if rows {
        (context.host_row, 2, 4, 3, i16::MAX as u32)
    } else {
        (context.host_column, 1, 3, 4, i32::MAX as u32)
    };
    let mut length = axis_ranges_len(
        start.index,
        end.index,
        start.absolute,
        end.absolute,
        host,
        relative_field,
        absolute_field,
    )?;
    length = checked(
        length,
        length_delimited_field_len(sentinel_field, absolute_range_len(sentinel, None)?)?,
    )?;
    checked(length, varint_field_len(5, 1)?)
}

fn range_message_field_count(end_present: bool) -> usize {
    if end_present { 2 } else { 1 }
}

fn axis_ranges_field_count(
    start: u32,
    end: u32,
    start_absolute: bool,
    end_absolute: bool,
) -> usize {
    match (start_absolute, end_absolute) {
        (false, false) | (true, true) => 1 + range_message_field_count(end != start),
        _ => 4,
    }
}

fn colon_nested_field_count(
    owner: Option<FormulaWriteOwnerUid>,
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
    _context: FormulaWriteContext,
) -> Result<usize, DecodeError> {
    // The AST field containing CrossTableExtra is counted by the caller;
    // these are its table-id field and the four CFUUID word fields.
    let owner_fields = if owner.is_some() { 5 } else { 0 };
    let sticky_fields = 4;
    let tract_fields = axis_ranges_field_count(
        start.column,
        end.column,
        start.column_absolute,
        end.column_absolute,
    )
    .checked_add(axis_ranges_field_count(
        start.row,
        end.row,
        start.row_absolute,
        end.row_absolute,
    ))
    .and_then(|value| value.checked_add(1))
    .ok_or_else(malformed)?;
    checked(checked(owner_fields, sticky_fields)?, tract_fields)
}

fn whole_axis_nested_field_count(
    owner: Option<FormulaWriteOwnerUid>,
    start: FormulaWriteAxisReference,
    end: FormulaWriteAxisReference,
    _rows: bool,
    _context: FormulaWriteContext,
) -> Result<usize, DecodeError> {
    let owner_fields = if owner.is_some() { 5 } else { 0 };
    let sticky_fields = 4;
    let tract_fields =
        axis_ranges_field_count(start.index, end.index, start.absolute, end.absolute)
            .checked_add(3)
            .ok_or_else(malformed)?;
    checked(checked(owner_fields, sticky_fields)?, tract_fields)
}

fn axis_ranges_len(
    start: u32,
    end: u32,
    start_absolute: bool,
    end_absolute: bool,
    host: u32,
    relative_field: u32,
    absolute_field: u32,
) -> Result<usize, DecodeError> {
    let mut length = 0usize;
    match (start_absolute, end_absolute) {
        (false, false) => {
            let begin = lowered_axis(start, host, false)?.coordinate;
            let finish = (end != start)
                .then(|| lowered_axis(end, host, false))
                .transpose()?
                .map(|axis| axis.coordinate);
            length = checked(
                length,
                length_delimited_field_len(relative_field, relative_range_len(begin, finish)?)?,
            )?;
        },
        (true, true) => {
            length = checked(
                length,
                length_delimited_field_len(
                    absolute_field,
                    absolute_range_len(start, (end != start).then_some(end))?,
                )?,
            )?
        },
        (false, true) => {
            length = checked(
                length,
                length_delimited_field_len(
                    relative_field,
                    relative_range_len(lowered_axis(start, host, false)?.coordinate, None)?,
                )?,
            )?;
            length = checked(
                length,
                length_delimited_field_len(absolute_field, absolute_range_len(end, None)?)?,
            )?;
        },
        (true, false) => {
            length = checked(
                length,
                length_delimited_field_len(
                    relative_field,
                    relative_range_len(lowered_axis(end, host, false)?.coordinate, None)?,
                )?,
            )?;
            length = checked(
                length,
                length_delimited_field_len(absolute_field, absolute_range_len(start, None)?)?,
            )?;
        },
    }
    Ok(length)
}

fn validate_write_requirements(
    requirements: FormulaWriteRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let buffa_max = usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| malformed())?;
    if requirements.output_bytes > options.max_bytes || requirements.output_bytes > buffa_max {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: requirements.output_bytes,
            maximum: options.max_bytes.min(buffa_max),
        }));
    }
    if requirements.fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: options.max_fields,
        }));
    }
    if requirements.work_bytes > options.max_work {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: options.max_work,
        }));
    }
    if requirements.nodes > options.max_nodes {
        return Err(DecodeError::limited(DecodeLimit::Nodes {
            observed: requirements.nodes,
            maximum: options.max_nodes,
        }));
    }
    if requirements.text_bytes > options.max_text_bytes {
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed: requirements.text_bytes,
            maximum: options.max_text_bytes,
        }));
    }
    if options.recursion_limit == 0
        || options.recursion_limit > MAX_DEPTH
        || requirements.max_depth > options.recursion_limit
    {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: requirements.max_depth.max(options.recursion_limit),
            maximum: options.recursion_limit.min(MAX_DEPTH),
        }));
    }
    Ok(())
}

fn write_node_kind(node: &FormulaWriteNode<'_>) -> u32 {
    match node {
        FormulaWriteNode::Number(_) => 17,
        FormulaWriteNode::Text(_) => 19,
        FormulaWriteNode::Boolean(_) => 18,
        FormulaWriteNode::CellReference { .. } => 36,
        FormulaWriteNode::ResolvedCellReference { .. } => 36,
        FormulaWriteNode::ResolvedRange { .. }
        | FormulaWriteNode::WholeRows { .. }
        | FormulaWriteNode::WholeColumns { .. } => 67,
        FormulaWriteNode::Binary(operator) => match operator {
            BinaryOperator::Add => 1,
            BinaryOperator::Subtract => 2,
            BinaryOperator::Multiply => 3,
            BinaryOperator::Divide => 4,
            BinaryOperator::Power => 5,
            BinaryOperator::Concatenate => 6,
            BinaryOperator::GreaterThan => 7,
            BinaryOperator::GreaterThanOrEqual => 8,
            BinaryOperator::LessThan => 9,
            BinaryOperator::LessThanOrEqual => 10,
            BinaryOperator::Equal => 11,
            BinaryOperator::NotEqual => 12,
        },
        FormulaWriteNode::Negation => 13,
        FormulaWriteNode::Percent => 15,
        FormulaWriteNode::Function { .. } => 16,
        FormulaWriteNode::Range => 29,
    }
}

const fn write_node_field_count(node: &FormulaWriteNode<'_>) -> usize {
    match node {
        FormulaWriteNode::Number(_) => 4,
        FormulaWriteNode::Text(_) | FormulaWriteNode::Boolean(_) => 2,
        FormulaWriteNode::CellReference { .. } => 3,
        FormulaWriteNode::ResolvedCellReference { .. } => 3,
        FormulaWriteNode::ResolvedRange { .. }
        | FormulaWriteNode::WholeRows { .. }
        | FormulaWriteNode::WholeColumns { .. } => 3,
        FormulaWriteNode::Function { .. } => 3,
        FormulaWriteNode::Binary(_)
        | FormulaWriteNode::Negation
        | FormulaWriteNode::Percent
        | FormulaWriteNode::Range => 1,
    }
}

fn encoded_write_node_len_in(
    node: &FormulaWriteNode<'_>,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<usize, DecodeError> {
    let base = varint_field_len(1, u64::from(write_node_kind(node)))?;
    match node {
        FormulaWriteNode::Number(value) => {
            let (low, high) = formula_decimal128_parts(*value)?;
            checked(
                checked(checked(base, 9)?, varint_field_len(42, low)?)?,
                varint_field_len(43, high)?,
            )
        },
        FormulaWriteNode::Text(text) => checked(base, length_delimited_field_len(6, text.len())?),
        FormulaWriteNode::Boolean(value) => checked(base, varint_field_len(5, u64::from(*value))?),
        FormulaWriteNode::CellReference { column, row } => {
            let column = encoded_axis_len(*column)?;
            let row = encoded_axis_len(*row)?;
            checked(
                checked(base, length_delimited_field_len(26, column)?)?,
                length_delimited_field_len(27, row)?,
            )
        },
        FormulaWriteNode::ResolvedCellReference { owner, reference } => {
            let target = resolved_target(context, owners, *owner)?;
            let column = lowered_axis(
                reference.column,
                context
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?
                    .host_column,
                reference.column_absolute,
            )?;
            let row = lowered_axis(
                reference.row,
                context
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?
                    .host_row,
                reference.row_absolute,
            )?;
            let mut length = checked(
                checked(
                    base,
                    length_delimited_field_len(26, encoded_axis_len(column)?)?,
                )?,
                length_delimited_field_len(27, encoded_axis_len(row)?)?,
            )?;
            if let Some(uid) = target.uid {
                length = checked(
                    length,
                    length_delimited_field_len(28, cross_extra_len(uid)?)?,
                )?;
            }
            Ok(length)
        },
        FormulaWriteNode::ResolvedRange { owner, start, end } => {
            colon_tract_node_len(base, *owner, *start, *end, false, context, owners)
        },
        FormulaWriteNode::WholeRows { owner, start, end } => {
            whole_axis_node_len(base, *owner, *start, *end, true, context, owners)
        },
        FormulaWriteNode::WholeColumns { owner, start, end } => {
            whole_axis_node_len(base, *owner, *start, *end, false, context, owners)
        },
        FormulaWriteNode::Function {
            identifier,
            argument_count,
        } => checked(
            checked(base, varint_field_len(2, u64::from(*identifier))?)?,
            varint_field_len(3, u64::from(*argument_count))?,
        ),
        FormulaWriteNode::Binary(_)
        | FormulaWriteNode::Negation
        | FormulaWriteNode::Percent
        | FormulaWriteNode::Range => Ok(base),
    }
}

fn encoded_axis_len(axis: FormulaWriteAxis) -> Result<usize, DecodeError> {
    checked(
        varint_field_len(1, u64::from(zigzag_encode_32(axis.coordinate)))?,
        varint_field_len(2, u64::from(axis.absolute))?,
    )
}

fn varint_field_len(field: u32, value: u64) -> Result<usize, DecodeError> {
    checked(varint_len(u64::from(field) << 3), varint_len(value))
}

fn length_delimited_field_len(field: u32, value_len: usize) -> Result<usize, DecodeError> {
    let value_len_u64 = u64::try_from(value_len).map_err(|_error| malformed())?;
    checked(
        checked(
            varint_len((u64::from(field) << 3) | 2),
            varint_len(value_len_u64),
        )?,
        value_len,
    )
}

fn reserve_formula_output(size: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|_error| DecodeError::limited(DecodeLimit::Allocation { requested: size }))?;
    if output.capacity() != size {
        return Err(DecodeError::limited(DecodeLimit::Allocation {
            requested: output.capacity(),
        }));
    }
    Ok(output)
}

fn encode_write_node_in(
    output: &mut Vec<u8>,
    node: &FormulaWriteNode<'_>,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<(), DecodeError> {
    put_varint_field(output, 1, u64::from(write_node_kind(node)));
    match node {
        FormulaWriteNode::Number(value) => {
            put_key(output, 4, 1);
            output.extend_from_slice(&value.to_bits().to_le_bytes());
            let (low, high) = formula_decimal128_parts(*value)?;
            put_varint_field(output, 42, low);
            put_varint_field(output, 43, high);
        },
        FormulaWriteNode::Text(text) => put_bytes_field(output, 6, text.as_bytes())?,
        FormulaWriteNode::Boolean(value) => put_varint_field(output, 5, u64::from(*value)),
        FormulaWriteNode::CellReference { column, row } => {
            put_axis_field(output, 26, *column)?;
            put_axis_field(output, 27, *row)?;
        },
        FormulaWriteNode::ResolvedCellReference { owner, reference } => {
            let target = resolved_target(context, owners, *owner)?;
            let context =
                context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            put_axis_field(
                output,
                26,
                lowered_axis(
                    reference.column,
                    context.host_column,
                    reference.column_absolute,
                )?,
            )?;
            put_axis_field(
                output,
                27,
                lowered_axis(reference.row, context.host_row, reference.row_absolute)?,
            )?;
            if let Some(uid) = target.uid {
                put_cross_extra_field(output, uid)?;
            }
        },
        FormulaWriteNode::ResolvedRange { owner, start, end } => {
            encode_colon_tract_node(output, *owner, *start, *end, context, owners)?;
        },
        FormulaWriteNode::WholeRows { owner, start, end } => {
            encode_whole_axis_node(output, *owner, *start, *end, true, context, owners)?;
        },
        FormulaWriteNode::WholeColumns { owner, start, end } => {
            encode_whole_axis_node(output, *owner, *start, *end, false, context, owners)?;
        },
        FormulaWriteNode::Function {
            identifier,
            argument_count,
        } => {
            put_varint_field(output, 2, u64::from(*identifier));
            put_varint_field(output, 3, u64::from(*argument_count));
        },
        FormulaWriteNode::Binary(_)
        | FormulaWriteNode::Negation
        | FormulaWriteNode::Percent
        | FormulaWriteNode::Range => {},
    }
    Ok(())
}

fn put_axis_field(
    output: &mut Vec<u8>,
    field: u32,
    axis: FormulaWriteAxis,
) -> Result<(), DecodeError> {
    put_key(output, field, 2);
    put_varint(
        output,
        u64::try_from(encoded_axis_len(axis)?).map_err(|_error| malformed())?,
    );
    put_varint_field(output, 1, u64::from(zigzag_encode_32(axis.coordinate)));
    put_varint_field(output, 2, u64::from(axis.absolute));
    Ok(())
}

fn put_cross_extra_field(
    output: &mut Vec<u8>,
    uid: FormulaWriteOwnerUid,
) -> Result<(), DecodeError> {
    put_key(output, 28, 2);
    put_varint(
        output,
        u64::try_from(cross_extra_len(uid)?).map_err(|_error| malformed())?,
    );
    put_key(output, 1, 2);
    put_varint(
        output,
        u64::try_from(cfuuid_len(uid)?).map_err(|_error| malformed())?,
    );
    for (field, word) in [
        (2, uid.lower as u32),
        (3, (uid.lower >> 32) as u32),
        (4, uid.upper as u32),
        (5, (uid.upper >> 32) as u32),
    ] {
        put_varint_field(output, field, u64::from(word));
    }
    Ok(())
}

fn put_sticky_field(
    output: &mut Vec<u8>,
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
) -> Result<(), DecodeError> {
    put_key(output, 33, 2);
    put_varint(
        output,
        u64::try_from(sticky_len()?).map_err(|_error| malformed())?,
    );
    put_varint_field(output, 1, u64::from(start.row_absolute));
    put_varint_field(output, 2, u64::from(start.column_absolute));
    put_varint_field(output, 3, u64::from(end.row_absolute));
    put_varint_field(output, 4, u64::from(end.column_absolute));
    Ok(())
}

fn put_relative_range_field(
    output: &mut Vec<u8>,
    field: u32,
    begin: i32,
    end: Option<i32>,
) -> Result<(), DecodeError> {
    put_key(output, field, 2);
    put_varint(
        output,
        u64::try_from(relative_range_len(begin, end)?).map_err(|_error| malformed())?,
    );
    put_varint_field(output, 1, begin as u64);
    if let Some(end) = end {
        put_varint_field(output, 2, end as u64);
    }
    Ok(())
}

fn put_absolute_range_field(
    output: &mut Vec<u8>,
    field: u32,
    begin: u32,
    end: Option<u32>,
) -> Result<(), DecodeError> {
    put_key(output, field, 2);
    put_varint(
        output,
        u64::try_from(absolute_range_len(begin, end)?).map_err(|_error| malformed())?,
    );
    put_varint_field(output, 1, u64::from(begin));
    if let Some(end) = end {
        put_varint_field(output, 2, u64::from(end));
    }
    Ok(())
}

fn encode_colon_tract_node(
    output: &mut Vec<u8>,
    owner: Option<FormulaWriteOwnerUid>,
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<(), DecodeError> {
    let target = resolved_target(context, owners, owner)?;
    let context = context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
    if let Some(uid) = target.uid {
        put_cross_extra_field(output, uid)?;
    }
    put_sticky_field(output, start, end)?;
    put_key(output, 40, 2);
    put_varint(
        output,
        u64::try_from(colon_tract_len(start, end, context)?).map_err(|_error| malformed())?,
    );
    put_axis_ranges(
        output,
        start.column,
        end.column,
        start.column_absolute,
        end.column_absolute,
        context.host_column,
        1,
        3,
    )?;
    put_axis_ranges(
        output,
        start.row,
        end.row,
        start.row_absolute,
        end.row_absolute,
        context.host_row,
        2,
        4,
    )?;
    put_varint_field(output, 5, 1);
    Ok(())
}

fn encode_whole_axis_node(
    output: &mut Vec<u8>,
    owner: Option<FormulaWriteOwnerUid>,
    start: FormulaWriteAxisReference,
    end: FormulaWriteAxisReference,
    rows: bool,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
) -> Result<(), DecodeError> {
    let target = resolved_target(context, owners, owner)?;
    let context = context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
    if let Some(uid) = target.uid {
        put_cross_extra_field(output, uid)?;
    }
    let sticky_start = if rows {
        FormulaWriteCellReference::new(0, 0, start.absolute, false)
    } else {
        FormulaWriteCellReference::new(0, 0, false, start.absolute)
    };
    let sticky_end = if rows {
        FormulaWriteCellReference::new(0, 0, end.absolute, false)
    } else {
        FormulaWriteCellReference::new(0, 0, false, end.absolute)
    };
    put_sticky_field(output, sticky_start, sticky_end)?;
    put_key(output, 40, 2);
    put_varint(
        output,
        u64::try_from(whole_axis_tract_len(start, end, rows, context)?)
            .map_err(|_error| malformed())?,
    );
    let (host, relative_field, absolute_field, sentinel_field, sentinel) = if rows {
        (context.host_row, 2, 4, 3, i16::MAX as u32)
    } else {
        (context.host_column, 1, 3, 4, i32::MAX as u32)
    };
    put_axis_ranges(
        output,
        start.index,
        end.index,
        start.absolute,
        end.absolute,
        host,
        relative_field,
        absolute_field,
    )?;
    put_absolute_range_field(output, sentinel_field, sentinel, None)?;
    put_varint_field(output, 5, 1);
    Ok(())
}

fn put_axis_ranges(
    output: &mut Vec<u8>,
    start: u32,
    end: u32,
    start_absolute: bool,
    end_absolute: bool,
    host: u32,
    relative_field: u32,
    absolute_field: u32,
) -> Result<(), DecodeError> {
    match (start_absolute, end_absolute) {
        (false, false) => put_relative_range_field(
            output,
            relative_field,
            lowered_axis(start, host, false)?.coordinate,
            (end != start)
                .then(|| lowered_axis(end, host, false))
                .transpose()?
                .map(|axis| axis.coordinate),
        ),
        (true, true) => {
            put_absolute_range_field(output, absolute_field, start, (end != start).then_some(end))
        },
        (false, true) => {
            put_relative_range_field(
                output,
                relative_field,
                lowered_axis(start, host, false)?.coordinate,
                None,
            )?;
            put_absolute_range_field(output, absolute_field, end, None)
        },
        (true, false) => {
            put_relative_range_field(
                output,
                relative_field,
                lowered_axis(end, host, false)?.coordinate,
                None,
            )?;
            put_absolute_range_field(output, absolute_field, start, None)
        },
    }
}

fn put_bytes_field(output: &mut Vec<u8>, field: u32, value: &[u8]) -> Result<(), DecodeError> {
    put_key(output, field, 2);
    put_varint(
        output,
        u64::try_from(value.len()).map_err(|_error| malformed())?,
    );
    output.extend_from_slice(value);
    Ok(())
}

fn put_varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
    put_key(output, field, 0);
    put_varint(output, value);
}

fn put_key(output: &mut Vec<u8>, field: u32, wire: u8) {
    put_varint(output, (u64::from(field) << 3) | u64::from(wire));
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn validate_canonical_formula_output(
    source: &[u8],
    nodes: &[FormulaWriteNode<'_>],
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let mut remaining = source;
    let root = next_field(&mut remaining, &mut budget, 1)?
        .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?;
    if root.number != 1 || !remaining.is_empty() {
        return Err(malformed());
    }
    let ast = root.bytes()?;
    let view: projection::FormulaArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| malformed())?;
    if !view.has_ast_node_array()
        || view.ast_node_array != ast
        || view.host_column.is_some()
        || view.host_row.is_some()
        || view.host_column_is_negative.is_some()
        || view.host_row_is_negative.is_some()
        || view.translation_flags.is_some()
        || view.host_table_uid.is_some()
        || view.host_column_uid.is_some()
        || view.host_row_uid.is_some()
    {
        return Err(malformed());
    }
    let mut remaining = ast;
    for expected in nodes {
        let field = next_field(&mut remaining, &mut budget, 2)?
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?;
        if field.number != 1 {
            return Err(malformed());
        }
        validate_canonical_write_node(
            field.bytes()?,
            expected,
            context,
            owners,
            &mut budget,
            options,
        )?;
    }
    if !remaining.is_empty() {
        return Err(malformed());
    }
    Ok(())
}

#[derive(Default)]
struct WrittenNodeFields<'source> {
    kind: Option<u32>,
    function: Option<u32>,
    arguments: Option<u32>,
    number: Option<u64>,
    boolean: Option<bool>,
    text: Option<&'source str>,
    column: Option<&'source [u8]>,
    row: Option<&'source [u8]>,
    cross_extra: Option<&'source [u8]>,
    sticky: Option<&'source [u8]>,
    colon: Option<&'source [u8]>,
    decimal_low: Option<u64>,
    decimal_high: Option<u64>,
}

fn validate_canonical_write_node(
    source: &[u8],
    expected: &FormulaWriteNode<'_>,
    context: Option<FormulaWriteContext>,
    owners: &[ResolvedFormulaWriteOwner],
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let mut fields = WrittenNodeFields::default();
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 3)? {
        match field.number {
            1 => set_once(&mut fields.kind, field.varint_u32()?)?,
            2 => set_once(&mut fields.function, field.varint_u32()?)?,
            3 => set_once(&mut fields.arguments, field.varint_u32()?)?,
            4 => set_once(&mut fields.number, field.fixed64()?)?,
            5 => set_once(&mut fields.boolean, field.boolean()?)?,
            6 => set_once(&mut fields.text, strict_utf8(field.bytes()?)?)?,
            26 => set_once(&mut fields.column, field.bytes()?)?,
            27 => set_once(&mut fields.row, field.bytes()?)?,
            28 => set_once(&mut fields.cross_extra, field.bytes()?)?,
            33 => set_once(&mut fields.sticky, field.bytes()?)?,
            40 => set_once(&mut fields.colon, field.bytes()?)?,
            42 => set_once(&mut fields.decimal_low, field.varint()?)?,
            43 => set_once(&mut fields.decimal_high, field.varint()?)?,
            _ => return Err(malformed()),
        }
    }
    if fields.kind != Some(write_node_kind(expected)) {
        return Err(malformed());
    }
    match expected {
        FormulaWriteNode::Number(value)
            if fields.number == Some(value.to_bits())
                && Some(formula_decimal128_parts(*value)?)
                    == fields.decimal_low.zip(fields.decimal_high) => {},
        FormulaWriteNode::Text(text) if fields.text == Some(*text) => {},
        FormulaWriteNode::Boolean(value) if fields.boolean == Some(*value) => {},
        FormulaWriteNode::CellReference { column, row } => {
            validate_written_axis(required(fields.column)?, *column, true, budget, options)?;
            validate_written_axis(required(fields.row)?, *row, false, budget, options)?;
        },
        FormulaWriteNode::ResolvedCellReference { owner, reference } => {
            let target = resolved_target(context, owners, *owner)?;
            let context =
                context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            validate_written_axis(
                required(fields.column)?,
                lowered_axis(
                    reference.column,
                    context.host_column,
                    reference.column_absolute,
                )?,
                true,
                budget,
                options,
            )?;
            validate_written_axis(
                required(fields.row)?,
                lowered_axis(reference.row, context.host_row, reference.row_absolute)?,
                false,
                budget,
                options,
            )?;
            validate_cross_extra(fields.cross_extra, target.uid, budget, options)?;
        },
        FormulaWriteNode::ResolvedRange { owner, start, end } => {
            let target = resolved_target(context, owners, *owner)?;
            validate_cross_extra(fields.cross_extra, target.uid, budget, options)?;
            let context =
                context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            validate_colon_shape(
                fields.sticky,
                required(fields.colon)?,
                *start,
                *end,
                context,
                budget,
                options,
            )?;
        },
        FormulaWriteNode::WholeRows { owner, start, end } => {
            let target = resolved_target(context, owners, *owner)?;
            validate_cross_extra(fields.cross_extra, target.uid, budget, options)?;
            let context =
                context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            validate_whole_axis_shape(
                fields.sticky,
                required(fields.colon)?,
                *start,
                *end,
                true,
                context,
                budget,
                options,
            )?;
        },
        FormulaWriteNode::WholeColumns { owner, start, end } => {
            let target = resolved_target(context, owners, *owner)?;
            validate_cross_extra(fields.cross_extra, target.uid, budget, options)?;
            let context =
                context.ok_or_else(|| DecodeError::invalid(InvalidReason::ExternalOwner))?;
            validate_whole_axis_shape(
                fields.sticky,
                required(fields.colon)?,
                *start,
                *end,
                false,
                context,
                budget,
                options,
            )?;
        },
        FormulaWriteNode::Function {
            identifier,
            argument_count,
        } if fields.function == Some(*identifier) && fields.arguments == Some(*argument_count) => {
        },
        FormulaWriteNode::Binary(_)
        | FormulaWriteNode::Negation
        | FormulaWriteNode::Percent
        | FormulaWriteNode::Range => {},
        _ => return Err(malformed()),
    }
    let view: projection::ASTNodeArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| malformed())?;
    if !view.has_node_type()
        || u32::try_from(view.node_type).ok() != fields.kind
        || view.function_index != fields.function
        || view.function_num_args != fields.arguments
        || view.number.map(f64::to_bits) != fields.number
        || view.boolean != fields.boolean
        || view.string != fields.text
        || view.local_cell_reference.is_some()
        || view.cross_table_cell_reference.is_some()
        || view.column != fields.column
        || view.row != fields.row
        || view.cross_table_extra != fields.cross_extra
        || view.sticky_bits != fields.sticky
        || view.colon_tract != fields.colon
        || view.decimal_low != fields.decimal_low
        || view.decimal_high != fields.decimal_high
        || view.uid_coordinate.is_some()
        || view.tract_list.is_some()
    {
        return Err(malformed());
    }
    Ok(())
}

fn validate_written_axis(
    source: &[u8],
    expected: FormulaWriteAxis,
    column: bool,
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let mut coordinate = None;
    let mut absolute = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 4)? {
        match field.number {
            1 => set_once(&mut coordinate, zigzag32(field.varint_u32()?))?,
            2 => set_once(&mut absolute, field.boolean()?)?,
            _ => return Err(malformed()),
        }
    }
    if coordinate != Some(expected.coordinate) || absolute != Some(expected.absolute) {
        return Err(malformed());
    }
    if column {
        let view: projection::ColumnCoordinateArchiveLazyView<'_> = options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| malformed())?;
        if !view.has_column()
            || view.column != expected.coordinate
            || view.absolute != Some(expected.absolute)
        {
            return Err(malformed());
        }
    } else {
        let view: projection::RowCoordinateArchiveLazyView<'_> = options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| malformed())?;
        if !view.has_row()
            || view.row != expected.coordinate
            || view.absolute != Some(expected.absolute)
        {
            return Err(malformed());
        }
    }
    Ok(())
}

fn validate_cross_extra(
    source: Option<&[u8]>,
    expected: Option<FormulaWriteOwnerUid>,
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    match (source, expected) {
        (None, None) => Ok(()),
        (Some(source), Some(expected)) => {
            budget.message(source, 4)?;
            let mut remaining = source;
            let table = next_field(&mut remaining, budget, 4)?
                .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?;
            if table.number != 1 || !remaining.is_empty() {
                return Err(malformed());
            }
            let uuid = table.bytes()?;
            let mut words = [None; 4];
            let mut uuid_remaining = uuid;
            while let Some(field) = next_field(&mut uuid_remaining, budget, 5)? {
                if !(2..=5).contains(&field.number) {
                    return Err(malformed());
                }
                set_once(&mut words[field.number as usize - 2], field.varint_u32()?)?;
            }
            let expected_words = [
                expected.lower as u32,
                (expected.lower >> 32) as u32,
                expected.upper as u32,
                (expected.upper >> 32) as u32,
            ];
            if words != expected_words.map(Some) {
                return Err(malformed());
            }
            let extra: projection::CrossTableExtraArchiveLazyView<'_> = options
                .buffa()
                .decode_lazy_view(source)
                .map_err(|_error| malformed())?;
            let uuid_view: projection::CFUUIDArchiveLazyView<'_> = options
                .buffa()
                .decode_lazy_view(extra.table_id)
                .map_err(|_error| malformed())?;
            if !extra.has_table_id()
                || uuid_view.word0 != words[0]
                || uuid_view.word1 != words[1]
                || uuid_view.word2 != words[2]
                || uuid_view.word3 != words[3]
            {
                return Err(malformed());
            }
            Ok(())
        },
        _ => Err(malformed()),
    }
}

fn validate_colon_shape(
    sticky: Option<&[u8]>,
    colon: &[u8],
    start: FormulaWriteCellReference,
    end: FormulaWriteCellReference,
    context: FormulaWriteContext,
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    validate_sticky(
        required(sticky)?,
        [
            start.row_absolute,
            start.column_absolute,
            end.row_absolute,
            end.column_absolute,
        ],
        budget,
        options,
    )?;
    validate_colon_bytes(colon, Some((start, end, context)), None, budget, options)
}

fn validate_whole_axis_shape(
    sticky: Option<&[u8]>,
    colon: &[u8],
    start: FormulaWriteAxisReference,
    end: FormulaWriteAxisReference,
    rows: bool,
    context: FormulaWriteContext,
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let bits = if rows {
        [start.absolute, false, end.absolute, false]
    } else {
        [false, start.absolute, false, end.absolute]
    };
    validate_sticky(required(sticky)?, bits, budget, options)?;
    validate_colon_bytes(
        colon,
        None,
        Some((start, end, rows, context)),
        budget,
        options,
    )
}

fn validate_sticky(
    source: &[u8],
    expected: [bool; 4],
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    budget.message(source, 4)?;
    let mut values = [None; 4];
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 4)? {
        if !(1..=4).contains(&field.number) {
            return Err(malformed());
        }
        set_once(&mut values[field.number as usize - 1], field.boolean()?)?;
    }
    if values != expected.map(Some) {
        return Err(malformed());
    }
    let view: projection::StickyBitsArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| malformed())?;
    if !view.has_begin_row_absolute()
        || !view.has_begin_column_absolute()
        || !view.has_end_row_absolute()
        || !view.has_end_column_absolute()
        || view.begin_row_absolute != expected[0]
        || view.begin_column_absolute != expected[1]
        || view.end_row_absolute != expected[2]
        || view.end_column_absolute != expected[3]
    {
        return Err(malformed());
    }
    Ok(())
}

fn validate_colon_bytes(
    source: &[u8],
    rectangle: Option<(
        FormulaWriteCellReference,
        FormulaWriteCellReference,
        FormulaWriteContext,
    )>,
    axis: Option<(
        FormulaWriteAxisReference,
        FormulaWriteAxisReference,
        bool,
        FormulaWriteContext,
    )>,
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    const EMPTY_RANGE_FIELD: (u32, (i64, Option<i64>)) = (0, (0, None));
    const MAX_RANGE_FIELDS: usize = 5;

    budget.message(source, 4)?;
    let mut remaining = source;
    let mut seen = [EMPTY_RANGE_FIELD; MAX_RANGE_FIELDS];
    let mut seen_len = 0usize;
    while let Some(field) = next_field(&mut remaining, budget, 4)? {
        let value = match field.number {
            1..=4 => {
                let bytes = field.bytes()?;
                budget.message(bytes, 5)?;
                let pair = decode_range_pair(bytes, field.number <= 2, budget)?;
                if field.number <= 2 {
                    let view: projection::RelativeRangeArchiveLazyView<'_> = options
                        .buffa()
                        .decode_lazy_view(bytes)
                        .map_err(|_error| malformed())?;
                    if !view.has_begin()
                        || i64::from(view.begin) != pair.0
                        || view.end.map(i64::from) != pair.1
                    {
                        return Err(malformed());
                    }
                } else {
                    let view: projection::AbsoluteRangeArchiveLazyView<'_> = options
                        .buffa()
                        .decode_lazy_view(bytes)
                        .map_err(|_error| malformed())?;
                    if !view.has_begin()
                        || i64::from(view.begin) != pair.0
                        || view.end.map(i64::from) != pair.1
                    {
                        return Err(malformed());
                    }
                }
                (field.number, pair)
            },
            5 if field.boolean()? => (5, (0, None)),
            _ => return Err(malformed()),
        };
        push_range_field(&mut seen, &mut seen_len, value)?;
    }
    let mut expected = [EMPTY_RANGE_FIELD; MAX_RANGE_FIELDS];
    let mut expected_len = 0usize;
    if let Some((start, end, context)) = rectangle {
        append_expected_axis(
            &mut expected,
            &mut expected_len,
            start.column,
            end.column,
            start.column_absolute,
            end.column_absolute,
            context.host_column,
            1,
            3,
        )?;
        append_expected_axis(
            &mut expected,
            &mut expected_len,
            start.row,
            end.row,
            start.row_absolute,
            end.row_absolute,
            context.host_row,
            2,
            4,
        )?;
    } else if let Some((start, end, rows, context)) = axis {
        let (host, relative, absolute, sentinel_field, sentinel) = if rows {
            (context.host_row, 2, 4, 3, i16::MAX as u32)
        } else {
            (context.host_column, 1, 3, 4, i32::MAX as u32)
        };
        append_expected_axis(
            &mut expected,
            &mut expected_len,
            start.index,
            end.index,
            start.absolute,
            end.absolute,
            host,
            relative,
            absolute,
        )?;
        push_range_field(
            &mut expected,
            &mut expected_len,
            (sentinel_field, (i64::from(sentinel), None)),
        )?;
    }
    push_range_field(&mut expected, &mut expected_len, (5, (0, None)))?;
    if seen_len != expected_len || seen[..seen_len] != expected[..expected_len] {
        return Err(malformed());
    }
    Ok(())
}

fn push_range_field(
    output: &mut [(u32, (i64, Option<i64>)); 5],
    length: &mut usize,
    value: (u32, (i64, Option<i64>)),
) -> Result<(), DecodeError> {
    let slot = output.get_mut(*length).ok_or_else(malformed)?;
    *slot = value;
    *length = length.checked_add(1).ok_or_else(malformed)?;
    Ok(())
}

fn decode_range_pair(
    source: &[u8],
    relative: bool,
    budget: &mut Budget,
) -> Result<(i64, Option<i64>), DecodeError> {
    let mut begin = None;
    let mut end = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 5)? {
        let value = field.varint()?;
        let value = if relative {
            i64::from(value as i32)
        } else {
            i64::try_from(value).map_err(|_error| malformed())?
        };
        match field.number {
            1 => set_once(&mut begin, value)?,
            2 => set_once(&mut end, value)?,
            _ => return Err(malformed()),
        }
    }
    Ok((required(begin)?, end))
}

fn append_expected_axis(
    output: &mut [(u32, (i64, Option<i64>)); 5],
    length: &mut usize,
    start: u32,
    end: u32,
    start_absolute: bool,
    end_absolute: bool,
    host: u32,
    relative_field: u32,
    absolute_field: u32,
) -> Result<(), DecodeError> {
    match (start_absolute, end_absolute) {
        (false, false) => push_range_field(
            output,
            length,
            (
                relative_field,
                (
                    i64::from(lowered_axis(start, host, false)?.coordinate),
                    (end != start)
                        .then(|| lowered_axis(end, host, false))
                        .transpose()?
                        .map(|v| i64::from(v.coordinate)),
                ),
            ),
        )?,
        (true, true) => push_range_field(
            output,
            length,
            (
                absolute_field,
                (i64::from(start), (end != start).then_some(i64::from(end))),
            ),
        )?,
        (false, true) => {
            push_range_field(
                output,
                length,
                (
                    relative_field,
                    (
                        i64::from(lowered_axis(start, host, false)?.coordinate),
                        None,
                    ),
                ),
            )?;
            push_range_field(output, length, (absolute_field, (i64::from(end), None)))?;
        },
        (true, false) => {
            push_range_field(
                output,
                length,
                (
                    relative_field,
                    (i64::from(lowered_axis(end, host, false)?.coordinate), None),
                ),
            )?;
            push_range_field(output, length, (absolute_field, (i64::from(start), None)))?;
        },
    }
    Ok(())
}

fn validate_context(context: FormulaContext) -> Result<(), DecodeError> {
    if context.owner == 0
        || context.rows == 0
        || context.columns == 0
        || context.host_row >= context.rows
        || context.host_column >= context.columns
    {
        return Err(DecodeError::invalid(InvalidReason::InvalidCoordinate));
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum DecodeMode {
    Evaluator,
    Dependencies,
}

fn decode_formula<V: FormulaVisitor + ?Sized>(
    source: &[u8],
    context: FormulaContext,
    resolution: Option<&ReadResolution<'_>>,
    budget: &mut Budget,
    visitor: &mut V,
    emit: bool,
    mode: DecodeMode,
) -> Result<(), DecodeError> {
    budget.message(source, 1)?;
    let mut ast = None;
    let mut host_column = None;
    let mut host_row = None;
    let mut column_negative = None;
    let mut row_negative = None;
    let mut opaque = [None; 4];
    let mut selected = [false; 9];
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if !(1..=9).contains(&field.number) {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        singular(&mut selected, field.number as usize - 1)?;
        match field.number {
            1 => ast = Some(field.bytes()?),
            2 => host_column = Some(field.varint_u32()?),
            3 => host_row = Some(field.varint_u32()?),
            4 => column_negative = Some(field.boolean()?),
            5 => row_negative = Some(field.boolean()?),
            6..=9 => opaque[field.number as usize - 6] = Some(field.bytes()?),
            _ => unreachable!(),
        }
    }
    let ast = ast.ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?;
    budget.message(source, 1)?;
    let view: projection::FormulaArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_ast_node_array()
        || view.ast_node_array != ast
        || view.host_column != host_column
        || view.host_row != host_row
        || view.host_column_is_negative != column_negative
        || view.host_row_is_negative != row_negative
        || view.translation_flags != opaque[0]
        || view.host_table_uid != opaque[1]
        || view.host_column_uid != opaque[2]
        || view.host_row_uid != opaque[3]
    {
        return Err(DecodeError::invalid(InvalidReason::MalformedWire));
    }
    decode_node_array(ast, context, resolution, budget, visitor, emit, mode, 2)
}

fn decode_node_array<V: FormulaVisitor + ?Sized>(
    source: &[u8],
    context: FormulaContext,
    resolution: Option<&ReadResolution<'_>>,
    budget: &mut Budget,
    visitor: &mut V,
    emit: bool,
    mode: DecodeMode,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child = depth + 1;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number != 1 {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        if !emit {
            budget.node()?;
        }
        decode_node(
            field.bytes()?,
            context,
            resolution,
            budget,
            visitor,
            emit,
            mode,
            child,
        )?;
    }
    Ok(())
}

#[derive(Default)]
struct NodeFields<'source> {
    kind: Option<u32>,
    function: Option<u32>,
    arguments: Option<u32>,
    number: Option<u64>,
    boolean: Option<bool>,
    text: Option<&'source str>,
    token: Option<bool>,
    local: Option<&'source [u8]>,
    cross: Option<&'source [u8]>,
    column: Option<&'source [u8]>,
    row: Option<&'source [u8]>,
    cross_extra: Option<&'source [u8]>,
    sticky: Option<&'source [u8]>,
    colon: Option<&'source [u8]>,
    decimal_low: Option<u64>,
    decimal_high: Option<u64>,
    uid: Option<&'source [u8]>,
    tract: Option<&'source [u8]>,
    whitespace: Option<&'source str>,
    unmodeled: bool,
    present: u64,
}

fn decode_node<V: FormulaVisitor + ?Sized>(
    source: &[u8],
    context: FormulaContext,
    resolution: Option<&ReadResolution<'_>>,
    budget: &mut Budget,
    visitor: &mut V,
    emit: bool,
    mode: DecodeMode,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let mut fields = NodeFields::default();
    let mut seen = [false; 48];
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number == 0 || field.number > 47 {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        singular(&mut seen, field.number as usize)?;
        fields.present |= field_bit(field.number);
        match field.number {
            1 => fields.kind = Some(field.varint_u32()?),
            2 => fields.function = Some(field.varint_u32()?),
            3 => fields.arguments = Some(field.varint_u32()?),
            4 => fields.number = Some(field.fixed64()?),
            5 => fields.boolean = Some(field.boolean()?),
            6 => {
                let text = strict_utf8(field.bytes()?)?;
                budget.text(text.len())?;
                fields.text = Some(text);
            },
            10 => fields.token = Some(field.boolean()?),
            14 => {
                let _thunk = field.bytes()?;
                fields.unmodeled = true;
            },
            15 => fields.local = Some(field.bytes()?),
            16 => fields.cross = Some(field.bytes()?),
            25 => {
                let text = strict_utf8(field.bytes()?)?;
                budget.text(text.len())?;
                fields.whitespace = Some(text);
            },
            26 => fields.column = Some(field.bytes()?),
            27 => fields.row = Some(field.bytes()?),
            28 => fields.cross_extra = Some(field.bytes()?),
            30 => fields.uid = Some(field.bytes()?),
            33 => fields.sticky = Some(field.bytes()?),
            38 => fields.tract = Some(field.bytes()?),
            40 => fields.colon = Some(field.bytes()?),
            42 => fields.decimal_low = Some(field.varint()?),
            43 => fields.decimal_high = Some(field.varint()?),
            _ => {
                validate_unmodeled_field(field, budget)?;
                fields.unmodeled = true;
            },
        }
    }
    let kind = fields
        .kind
        .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?;
    if resolution.is_some() && mode == DecodeMode::Dependencies && matches!(kind, 6 | 19) {
        budget.unsupported_evaluator();
    }
    budget.message(source, depth)?;
    let view: projection::ASTNodeArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_node_type()
        || u32::try_from(view.node_type).ok() != Some(kind)
        || view.function_index != fields.function
        || view.function_num_args != fields.arguments
        || view.number.map(f64::to_bits) != fields.number
        || view.boolean != fields.boolean
        || view.string != fields.text
        || view.token_boolean != fields.token
        || view.local_cell_reference != fields.local
        || view.cross_table_cell_reference != fields.cross
        || view.whitespace != fields.whitespace
        || view.column != fields.column
        || view.row != fields.row
        || view.cross_table_extra != fields.cross_extra
        || view.sticky_bits != fields.sticky
        || view.colon_tract != fields.colon
        || view.decimal_low != fields.decimal_low
        || view.decimal_high != fields.decimal_high
        || view.uid_coordinate != fields.uid
        || view.tract_list != fields.tract
    {
        return Err(DecodeError::invalid(InvalidReason::MalformedWire));
    }
    if fields.cross.is_some()
        || fields.uid.is_some()
        || matches!(kind, 28 | 48 | 63..=66 | 68)
        || (resolution.is_none() && (fields.cross_extra.is_some() || kind == 67))
    {
        return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
    }
    if resolution.is_none() && (fields.decimal_low.is_some() || fields.decimal_high.is_some()) {
        fields.unmodeled = true;
    }
    if resolution.is_some() {
        match (fields.decimal_low, fields.decimal_high) {
            (None, None) => {},
            (Some(low), Some(high))
                if fields
                    .number
                    .map(f64::from_bits)
                    .and_then(|value| formula_decimal128_parts(value).ok())
                    == Some((low, high)) => {},
            _ => return Err(DecodeError::invalid(InvalidReason::InvalidScalar)),
        }
    }
    if resolution.is_some() && kind == 19 {
        if fields.present != field_bit(1) | field_bit(6) {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        if emit {
            visitor.visit_text(required(fields.text)?)?;
        }
        return Ok(());
    }
    let unsupported_function = kind == 16
        && fields
            .function
            .is_some_and(|identifier| !matches!(identifier, 15 | 30 | 84 | 88 | 168));
    let present = fields.present;
    if fields.unmodeled || (fields.tract.is_some() && kind != 45) || unsupported_function {
        if mode == DecodeMode::Dependencies
            && dependency_tolerates(kind, present, unsupported_function)
        {
            return unsupported_local(kind, fields.function, mode, budget, visitor, emit);
        }
        return Err(DecodeError::invalid(InvalidReason::UnsupportedFormula));
    }
    if let Some(resolution) = resolution {
        if let Some(node) =
            classify_resolved_node(kind, &fields, resolution, budget, depth + 1, visitor, emit)?
        {
            if emit {
                visitor.visit_node(node)?;
            }
            return Ok(());
        }
    }
    let (node, precedent) = match classify_node(kind, fields, context, budget, depth + 1) {
        Ok(value) => value,
        Err(error)
            if mode == DecodeMode::Dependencies
                && error.invalid_reason() == Some(InvalidReason::UnsupportedFormula) =>
        {
            if dependency_tolerates(kind, present, false) {
                return unsupported_local(kind, None, mode, budget, visitor, emit);
            }
            return Err(error);
        },
        Err(error) => return Err(error),
    };
    if emit {
        visitor.visit_node(node)?;
        if let Some(precedent) = precedent {
            visitor.visit_precedent(precedent)?;
        }
    } else if precedent.is_some() {
        budget.precedent()?;
    }
    Ok(())
}

const fn field_bit(number: u32) -> u64 {
    1u64 << number
}

fn dependency_tolerates(kind: u32, present: u64, unsupported_function: bool) -> bool {
    let whitespace = field_bit(1) | field_bit(25);
    let allowed = match kind {
        16 if unsupported_function => whitespace | field_bit(2) | field_bit(3),
        6 | 30 | 34 | 35 | 54 | 56 | 57 | 69 | 70 => whitespace,
        19 => whitespace | field_bit(6),
        17 => whitespace | field_bit(4) | field_bit(42) | field_bit(43),
        20 => whitespace | field_bit(7) | field_bit(19) | field_bit(20) | field_bit(21),
        21 => {
            whitespace
                | field_bit(8)
                | field_bit(9)
                | field_bit(22)
                | field_bit(23)
                | field_bit(24)
                | field_bit(29)
        },
        24 => whitespace | field_bit(11) | field_bit(12),
        25 => whitespace | field_bit(13),
        31 => whitespace | field_bit(17) | field_bit(18),
        52 => whitespace | field_bit(34) | field_bit(35) | field_bit(36),
        53 => whitespace | field_bit(37),
        _ => return false,
    };
    present & !allowed == 0
}

fn unsupported_local<V: FormulaVisitor + ?Sized>(
    node_type: u32,
    function_identifier: Option<u32>,
    mode: DecodeMode,
    budget: &mut Budget,
    visitor: &mut V,
    emit: bool,
) -> Result<(), DecodeError> {
    if mode == DecodeMode::Evaluator {
        return Err(DecodeError::invalid(InvalidReason::UnsupportedFormula));
    }
    let unsupported = UnsupportedLocal {
        node_type,
        function_identifier,
    };
    if emit {
        visitor.visit_unsupported_local(unsupported)
    } else {
        budget.unsupported_local()
    }
}

fn classify_resolved_node<V: FormulaVisitor + ?Sized>(
    kind: u32,
    fields: &NodeFields<'_>,
    resolution: &ReadResolution<'_>,
    budget: &mut Budget,
    depth: u32,
    visitor: &mut V,
    emit: bool,
) -> Result<Option<FormulaNode>, DecodeError> {
    if kind == 17 {
        let legacy = field_bit(1) | field_bit(4);
        let writer = legacy | field_bit(42) | field_bit(43);
        if !matches!(fields.present, value if value == legacy || value == writer) {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        return Ok(None);
    }
    if kind == 36 {
        let allowed = field_bit(1)
            | field_bit(26)
            | field_bit(27)
            | if fields.cross_extra.is_some() {
                field_bit(28)
            } else {
                0
            };
        if fields.present != allowed {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        let target = read_target(fields.cross_extra, resolution, budget)?;
        let row = decode_axis(required(fields.row)?, false, budget, depth)?;
        let column = decode_axis(required(fields.column)?, true, budget, depth)?;
        let coordinate = CellCoordinate {
            row: resolve(resolution.context.host_row, row.0, row.1)?,
            column: resolve(resolution.context.host_column, column.0, column.1)?,
        };
        validate_target_coordinate(coordinate, target)?;
        let precedent = LocalPrecedent {
            owner: target.internal_owner,
            coordinate,
        };
        if emit {
            visitor.visit_precedent(precedent)?;
        } else {
            budget.precedent()?;
        }
        return Ok(Some(FormulaNode::ResolvedCellReference {
            owner: target.internal_owner,
            coordinate,
        }));
    }
    if kind != 67 {
        return Ok(None);
    }
    let allowed = field_bit(1)
        | field_bit(33)
        | field_bit(40)
        | if fields.cross_extra.is_some() {
            field_bit(28)
        } else {
            0
        };
    if fields.present != allowed {
        return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
    }
    let target = read_target(fields.cross_extra, resolution, budget)?;
    let (top, left, bottom, right) = decode_resolved_colon(
        required(fields.sticky)?,
        required(fields.colon)?,
        resolution.context,
        target,
        budget,
    )?;
    let range = FormulaWriteRange {
        internal_owner: target.internal_owner,
        top,
        left,
        bottom,
        right,
    };
    let cells = usize::try_from(bottom - top + 1)
        .ok()
        .and_then(|rows| {
            usize::try_from(right - left + 1)
                .ok()
                .and_then(|columns| rows.checked_mul(columns))
        })
        .ok_or_else(malformed)?;
    budget.work(cells.checked_add(1).ok_or_else(malformed)?)?;
    if emit {
        visitor.visit_range(range)?;
        for row in top..=bottom {
            for column in left..=right {
                visitor.visit_precedent(LocalPrecedent {
                    owner: target.internal_owner,
                    coordinate: CellCoordinate { row, column },
                })?;
            }
        }
    } else {
        budget.range()?;
        for _ in 0..cells {
            budget.precedent()?;
        }
    }
    Ok(Some(FormulaNode::ResolvedRange {
        owner: target.internal_owner,
        top,
        left,
        bottom,
        right,
    }))
}

fn read_target(
    cross: Option<&[u8]>,
    resolution: &ReadResolution<'_>,
    budget: &mut Budget,
) -> Result<WriteTarget, DecodeError> {
    let Some(cross) = cross else {
        return Ok(WriteTarget {
            internal_owner: resolution.context.internal_owner,
            uid: None,
            rows: resolution.context.rows,
            columns: resolution.context.columns,
        });
    };
    let uid = decode_cross_uid(cross, budget, budget.options)?;
    let mut left = 0usize;
    let mut right = resolution.registry.owners.len();
    let position = loop {
        if left >= right {
            return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
        }
        let middle = left + (right - left) / 2;
        budget.work(1)?;
        match resolution.registry.owners[middle].uid.cmp(&uid) {
            core::cmp::Ordering::Less => left = middle + 1,
            core::cmp::Ordering::Greater => right = middle,
            core::cmp::Ordering::Equal => break middle,
        }
    };
    let owner = resolution.registry.owners[position];
    if owner.internal_owner == resolution.context.internal_owner {
        return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
    }
    Ok(WriteTarget {
        internal_owner: owner.internal_owner,
        uid: Some(uid),
        rows: owner.rows,
        columns: owner.columns,
    })
}

fn decode_cross_uid(
    source: &[u8],
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<FormulaWriteOwnerUid, DecodeError> {
    budget.message(source, 4)?;
    let mut remaining = source;
    let table = next_field(&mut remaining, budget, 4)?
        .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?;
    if table.number != 1 || !remaining.is_empty() {
        return Err(malformed());
    }
    let uuid = table.bytes()?;
    let mut words = [None; 4];
    let mut uuid_remaining = uuid;
    while let Some(field) = next_field(&mut uuid_remaining, budget, 5)? {
        if !(2..=5).contains(&field.number) {
            return Err(malformed());
        }
        set_once(&mut words[field.number as usize - 2], field.varint_u32()?)?;
    }
    let words = [
        required(words[0])?,
        required(words[1])?,
        required(words[2])?,
        required(words[3])?,
    ];
    let uid = FormulaWriteOwnerUid::from_halves(
        u64::from(words[0]) | (u64::from(words[1]) << 32),
        u64::from(words[2]) | (u64::from(words[3]) << 32),
    );
    validate_cross_extra(Some(source), Some(uid), budget, options)?;
    Ok(uid)
}

fn validate_target_coordinate(
    value: CellCoordinate,
    target: WriteTarget,
) -> Result<(), DecodeError> {
    if value.row >= target.rows || value.column >= target.columns {
        Err(DecodeError::invalid(InvalidReason::InvalidCoordinate))
    } else {
        Ok(())
    }
}

fn decode_resolved_colon(
    sticky: &[u8],
    colon: &[u8],
    context: FormulaWriteContext,
    target: WriteTarget,
    budget: &mut Budget,
) -> Result<(u32, u32, u32, u32), DecodeError> {
    let bits = decode_sticky_bits(sticky, budget, budget.options)?;
    let mut pairs = [None; 4];
    let mut preserve = None;
    let mut remaining = colon;
    budget.message(colon, 4)?;
    while let Some(field) = next_field(&mut remaining, budget, 4)? {
        match field.number {
            1..=4 => {
                let bytes = field.bytes()?;
                budget.message(bytes, 5)?;
                set_once(
                    &mut pairs[field.number as usize - 1],
                    decode_range_pair(bytes, field.number <= 2, budget)?,
                )?;
            },
            5 => set_once(&mut preserve, field.boolean()?)?,
            _ => return Err(malformed()),
        }
    }
    if preserve != Some(true) {
        return Err(malformed());
    }
    let column_sentinel = pairs[2] == Some((i64::from(i16::MAX as u32), None));
    let row_sentinel = pairs[3] == Some((i64::from(i32::MAX as u32), None));
    if !column_sentinel && !row_sentinel {
        let (left, right) = resolved_axis_bounds(
            pairs[0],
            pairs[2],
            bits[1],
            bits[3],
            context.host_column,
            target.columns,
        )?;
        let (top, bottom) = resolved_axis_bounds(
            pairs[1],
            pairs[3],
            bits[0],
            bits[2],
            context.host_row,
            target.rows,
        )?;
        let start = FormulaWriteCellReference::new(top, left, bits[0], bits[1]);
        let end = FormulaWriteCellReference::new(bottom, right, bits[2], bits[3]);
        validate_colon_shape(
            Some(sticky),
            colon,
            start,
            end,
            context,
            budget,
            budget.options,
        )?;
        return Ok((top, left, bottom, right));
    }
    if column_sentinel && row_sentinel {
        return Err(malformed());
    }
    if column_sentinel {
        let (top, bottom) = resolved_axis_bounds(
            pairs[1],
            pairs[3],
            bits[0],
            bits[2],
            context.host_row,
            target.rows,
        )?;
        let start = FormulaWriteAxisReference::new(top, bits[0]);
        let end = FormulaWriteAxisReference::new(bottom, bits[2]);
        validate_whole_axis_shape(
            Some(sticky),
            colon,
            start,
            end,
            true,
            context,
            budget,
            budget.options,
        )?;
        return Ok((top, 0, bottom, target.columns - 1));
    }
    let (left, right) = resolved_axis_bounds(
        pairs[0],
        pairs[2],
        bits[1],
        bits[3],
        context.host_column,
        target.columns,
    )?;
    let start = FormulaWriteAxisReference::new(left, bits[1]);
    let end = FormulaWriteAxisReference::new(right, bits[3]);
    validate_whole_axis_shape(
        Some(sticky),
        colon,
        start,
        end,
        false,
        context,
        budget,
        budget.options,
    )?;
    Ok((0, left, target.rows - 1, right))
}

fn resolved_axis_bounds(
    relative: Option<(i64, Option<i64>)>,
    absolute: Option<(i64, Option<i64>)>,
    start_absolute: bool,
    end_absolute: bool,
    host: u32,
    maximum: u32,
) -> Result<(u32, u32), DecodeError> {
    let resolve_relative = |value: i64| {
        u32::try_from(i64::from(host).checked_add(value).ok_or_else(malformed)?)
            .map_err(|_error| DecodeError::invalid(InvalidReason::InvalidCoordinate))
    };
    let absolute_value = |value: i64| {
        u32::try_from(value)
            .map_err(|_error| DecodeError::invalid(InvalidReason::InvalidCoordinate))
    };
    let (start, end) = match (start_absolute, end_absolute) {
        (false, false) => {
            let (begin, end) = required(relative)?;
            require(absolute.is_none())?;
            (
                resolve_relative(begin)?,
                resolve_relative(end.unwrap_or(begin))?,
            )
        },
        (true, true) => {
            let (begin, end) = required(absolute)?;
            require(relative.is_none())?;
            (
                absolute_value(begin)?,
                absolute_value(end.unwrap_or(begin))?,
            )
        },
        (false, true) => {
            let (relative, relative_end) = required(relative)?;
            let (absolute, absolute_end) = required(absolute)?;
            require(relative_end.is_none() && absolute_end.is_none())?;
            (resolve_relative(relative)?, absolute_value(absolute)?)
        },
        (true, false) => {
            let (relative, relative_end) = required(relative)?;
            let (absolute, absolute_end) = required(absolute)?;
            require(relative_end.is_none() && absolute_end.is_none())?;
            (absolute_value(absolute)?, resolve_relative(relative)?)
        },
    };
    if start > end || end >= maximum {
        return Err(DecodeError::invalid(InvalidReason::InvalidCoordinate));
    }
    Ok((start, end))
}

fn decode_sticky_bits(
    source: &[u8],
    budget: &mut Budget,
    options: DecodeOptions,
) -> Result<[bool; 4], DecodeError> {
    budget.message(source, 4)?;
    let mut values = [None; 4];
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 4)? {
        if !(1..=4).contains(&field.number) {
            return Err(malformed());
        }
        set_once(&mut values[field.number as usize - 1], field.boolean()?)?;
    }
    let values = [
        required(values[0])?,
        required(values[1])?,
        required(values[2])?,
        required(values[3])?,
    ];
    validate_sticky(source, values, budget, options)?;
    Ok(values)
}

fn classify_node(
    kind: u32,
    fields: NodeFields<'_>,
    context: FormulaContext,
    budget: &mut Budget,
    depth: u32,
) -> Result<(FormulaNode, Option<LocalPrecedent>), DecodeError> {
    const FUNCTION: u16 = 1 << 0;
    const ARGUMENTS: u16 = 1 << 1;
    const NUMBER: u16 = 1 << 2;
    const BOOLEAN: u16 = 1 << 3;
    const TOKEN: u16 = 1 << 4;
    const LOCAL: u16 = 1 << 5;
    const COLUMN: u16 = 1 << 6;
    const ROW: u16 = 1 << 7;
    const TRACT: u16 = 1 << 8;
    let payload = present(fields.function.is_some(), FUNCTION)
        | present(fields.arguments.is_some(), ARGUMENTS)
        | present(fields.number.is_some(), NUMBER)
        | present(fields.boolean.is_some(), BOOLEAN)
        | present(fields.token.is_some(), TOKEN)
        | present(fields.local.is_some(), LOCAL)
        | present(fields.column.is_some(), COLUMN)
        | present(fields.row.is_some(), ROW)
        | present(fields.tract.is_some(), TRACT);
    let no_payload = || payload == 0;
    let binary = match kind {
        1 => Some(BinaryOperator::Add),
        2 => Some(BinaryOperator::Subtract),
        3 => Some(BinaryOperator::Multiply),
        4 => Some(BinaryOperator::Divide),
        5 => Some(BinaryOperator::Power),
        6 => Some(BinaryOperator::Concatenate),
        7 => Some(BinaryOperator::GreaterThan),
        8 => Some(BinaryOperator::GreaterThanOrEqual),
        9 => Some(BinaryOperator::LessThan),
        10 => Some(BinaryOperator::LessThanOrEqual),
        11 => Some(BinaryOperator::Equal),
        12 => Some(BinaryOperator::NotEqual),
        _ => None,
    };
    if let Some(operator) = binary {
        require(no_payload())?;
        return Ok((FormulaNode::Binary(operator), None));
    }
    let (node, coordinate) = match kind {
        13 if no_payload() => (FormulaNode::Negation, None),
        14 if no_payload() => (FormulaNode::PlusSign, None),
        15 if no_payload() => (FormulaNode::Percent, None),
        16 => (
            FormulaNode::Function {
                identifier: fields
                    .function
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
                argument_count: fields.arguments.unwrap_or(0),
            },
            None,
        ),
        17 => (
            FormulaNode::Number {
                bits: fields
                    .number
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
            },
            None,
        ),
        18 => (
            FormulaNode::Boolean(
                fields
                    .boolean
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
            ),
            None,
        ),
        22 if no_payload() => (FormulaNode::Empty, None),
        23 => (
            FormulaNode::Token(
                fields
                    .token
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
            ),
            None,
        ),
        27 => {
            let local = decode_local(
                fields
                    .local
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
                context,
                budget,
                depth,
            )?;
            (
                FormulaNode::LocalCell {
                    coordinate: local.0,
                    row_is_sticky: local.1,
                    column_is_sticky: local.2,
                },
                Some(local.0),
            )
        },
        29 if no_payload() => (FormulaNode::Colon, None),
        32 => (FormulaNode::AppendWhitespace, None),
        33 => (FormulaNode::PrependWhitespace, None),
        36 => {
            let row = decode_axis(
                fields
                    .row
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
                false,
                budget,
                depth,
            )?;
            let column = decode_axis(
                fields
                    .column
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
                true,
                budget,
                depth,
            )?;
            let coordinate = CellCoordinate {
                row: resolve(context.host_row, row.0, row.1)?,
                column: resolve(context.host_column, column.0, column.1)?,
            };
            validate_coordinate(coordinate, context)?;
            (FormulaNode::CellReference { coordinate }, Some(coordinate))
        },
        45 if fields.tract.is_none() => (FormulaNode::ColonWithUids, None),
        28 | 48 | 63..=68 => return Err(DecodeError::invalid(InvalidReason::ExternalOwner)),
        _ => return Err(DecodeError::invalid(InvalidReason::UnsupportedFormula)),
    };
    let allowed = match kind {
        16 => FUNCTION | ARGUMENTS,
        17 => NUMBER,
        18 => BOOLEAN,
        23 => TOKEN,
        27 => LOCAL,
        36 => COLUMN | ROW,
        45 => TRACT,
        _ => 0,
    };
    require(payload & !allowed == 0)?;
    let precedent = coordinate.map(|coordinate| LocalPrecedent {
        owner: context.owner,
        coordinate,
    });
    Ok((node, precedent))
}

fn decode_local(
    source: &[u8],
    context: FormulaContext,
    budget: &mut Budget,
    depth: u32,
) -> Result<(CellCoordinate, u32, u32), DecodeError> {
    budget.message(source, depth)?;
    let mut values = [None; 4];
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if !(1..=4).contains(&field.number) {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        let slot = &mut values[field.number as usize - 1];
        set_once(slot, field.varint_u32()?)?;
    }
    let coordinate = CellCoordinate {
        row: required(values[0])?,
        column: required(values[1])?,
    };
    let row_is_sticky = required(values[2])?;
    let column_is_sticky = required(values[3])?;
    validate_coordinate(coordinate, context)?;
    budget.message(source, depth)?;
    let view: projection::LocalCellReferenceArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_row_handle()
        || !view.has_column_handle()
        || !view.has_row_is_sticky()
        || !view.has_column_is_sticky()
        || view.row_handle != coordinate.row
        || view.column_handle != coordinate.column
        || view.row_is_sticky != row_is_sticky
        || view.column_is_sticky != column_is_sticky
    {
        return Err(DecodeError::invalid(InvalidReason::MissingRequired));
    }
    Ok((coordinate, row_is_sticky, column_is_sticky))
}

fn decode_axis(
    source: &[u8],
    column: bool,
    budget: &mut Budget,
    depth: u32,
) -> Result<(i32, bool), DecodeError> {
    budget.message(source, depth)?;
    let mut coordinate = None;
    let mut absolute = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut coordinate, zigzag32(field.varint_u32()?))?,
            2 => set_once(&mut absolute, field.boolean()?)?,
            _ => return Err(DecodeError::invalid(InvalidReason::UnexpectedField)),
        }
    }
    let coordinate = required(coordinate)?;
    budget.message(source, depth)?;
    if column {
        let view: projection::ColumnCoordinateArchiveLazyView<'_> = budget
            .options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if !view.has_column() || view.column != coordinate || view.absolute != absolute {
            return Err(DecodeError::invalid(InvalidReason::MalformedWire));
        }
    } else {
        let view: projection::RowCoordinateArchiveLazyView<'_> = budget
            .options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if !view.has_row() || view.row != coordinate || view.absolute != absolute {
            return Err(DecodeError::invalid(InvalidReason::MalformedWire));
        }
    }
    Ok((coordinate, absolute.unwrap_or(false)))
}

fn resolve(host: u32, coordinate: i32, absolute: bool) -> Result<u32, DecodeError> {
    let value = if absolute {
        i64::from(coordinate)
    } else {
        i64::from(host) + i64::from(coordinate)
    };
    u32::try_from(value).map_err(|_error| DecodeError::invalid(InvalidReason::InvalidCoordinate))
}

fn validate_coordinate(value: CellCoordinate, context: FormulaContext) -> Result<(), DecodeError> {
    if value.row >= context.rows || value.column >= context.columns {
        return Err(DecodeError::invalid(InvalidReason::InvalidCoordinate));
    }
    Ok(())
}

fn require(condition: bool) -> Result<(), DecodeError> {
    if condition {
        Ok(())
    } else {
        Err(DecodeError::invalid(InvalidReason::UnexpectedField))
    }
}

const fn present(value: bool, mask: u16) -> u16 {
    if value { mask } else { 0 }
}

fn required<T: Copy>(value: Option<T>) -> Result<T, DecodeError> {
    value.ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))
}

fn singular(seen: &mut [bool], index: usize) -> Result<(), DecodeError> {
    let value = seen
        .get_mut(index)
        .ok_or_else(|| DecodeError::invalid(InvalidReason::UnexpectedField))?;
    if *value {
        return Err(DecodeError::invalid(InvalidReason::DuplicateField));
    }
    *value = true;
    Ok(())
}

fn set_once<T>(slot: &mut Option<T>, value: T) -> Result<(), DecodeError> {
    if slot.is_some() {
        return Err(DecodeError::invalid(InvalidReason::DuplicateField));
    }
    *slot = Some(value);
    Ok(())
}

fn validate_unmodeled_field(field: Field<'_>, budget: &mut Budget) -> Result<(), DecodeError> {
    match field.number {
        9 | 11..=13 | 18 | 22..=24 | 37 | 42..=43 | 46..=47 => {
            let _value = field.varint()?;
        },
        19..=20 | 29 | 36 => {
            let _value = field.boolean()?;
        },
        7..=8 => {
            let _value = field.fixed64()?;
        },
        6 | 17 | 21 | 34..=35 => {
            let value = strict_utf8(field.bytes()?)?;
            budget.text(value.len())?;
        },
        33 | 39..=41 | 44..=45 => {
            let _value = field.bytes()?;
        },
        _ => return Err(DecodeError::invalid(InvalidReason::UnexpectedField)),
    }
    Ok(())
}

#[derive(Clone)]
struct Budget {
    options: DecodeOptions,
    bytes: usize,
    fields: usize,
    work: usize,
    max_depth: u32,
    text: usize,
    nodes: usize,
    precedents: usize,
    ranges: usize,
    unsupported_local: usize,
    evaluator_supported: bool,
    one_time_work: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        if source.len() > options.max_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_DEPTH {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_DEPTH,
            }));
        }
        Ok(Self {
            options,
            bytes: source.len(),
            fields: 0,
            work: 0,
            max_depth: 0,
            text: 0,
            nodes: 0,
            precedents: 0,
            ranges: 0,
            unsupported_local: 0,
            evaluator_supported: true,
            one_time_work: 0,
        })
    }
    fn message(&mut self, source: &[u8], depth: u32) -> Result<(), DecodeError> {
        self.depth(depth)?;
        self.work(source.len())
    }
    fn field(&mut self) -> Result<(), DecodeError> {
        self.fields = checked(self.fields, 1)?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        Ok(())
    }
    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.work = checked(self.work, amount)?;
        if self.work > self.options.max_work {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work,
                maximum: self.options.max_work,
            }));
        }
        Ok(())
    }
    fn one_time_work(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.one_time_work = checked(self.one_time_work, amount)?;
        self.work(amount)
    }
    fn depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }
    fn node(&mut self) -> Result<(), DecodeError> {
        self.nodes = checked(self.nodes, 1)?;
        if self.nodes > self.options.max_nodes {
            return Err(DecodeError::limited(DecodeLimit::Nodes {
                observed: self.nodes,
                maximum: self.options.max_nodes,
            }));
        }
        Ok(())
    }
    fn precedent(&mut self) -> Result<(), DecodeError> {
        self.precedents = checked(self.precedents, 1)?;
        Ok(())
    }
    fn range(&mut self) -> Result<(), DecodeError> {
        self.ranges = checked(self.ranges, 1)?;
        Ok(())
    }
    fn unsupported_local(&mut self) -> Result<(), DecodeError> {
        self.unsupported_local = checked(self.unsupported_local, 1)?;
        Ok(())
    }
    fn unsupported_evaluator(&mut self) {
        self.evaluator_supported = false;
    }
    fn text(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.text = checked(self.text, amount)?;
        if self.text > self.options.max_text_bytes {
            return Err(DecodeError::limited(DecodeLimit::Text {
                observed: self.text,
                maximum: self.options.max_text_bytes,
            }));
        }
        Ok(())
    }
    fn preflight_callback_pass(&self) -> Result<(), DecodeError> {
        let fields = self.fields.checked_mul(2);
        let pass_work = self
            .work
            .checked_sub(self.one_time_work)
            .ok_or_else(malformed)?;
        let work = pass_work
            .checked_mul(2)
            .and_then(|value| value.checked_add(self.one_time_work));
        let text = self.text.checked_mul(2);
        let probe = Self {
            options: self.options,
            bytes: self.bytes,
            fields: fields.ok_or_else(malformed)?,
            work: work.ok_or_else(malformed)?,
            max_depth: self.max_depth,
            text: text.ok_or_else(malformed)?,
            nodes: self.nodes,
            precedents: self.precedents,
            ranges: self.ranges,
            unsupported_local: self.unsupported_local,
            evaluator_supported: self.evaluator_supported,
            one_time_work: self.one_time_work,
        };
        if probe.fields > probe.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: probe.fields,
                maximum: probe.options.max_fields,
            }));
        }
        if probe.work > probe.options.max_work {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: probe.work,
                maximum: probe.options.max_work,
            }));
        }
        if probe.text > probe.options.max_text_bytes {
            return Err(DecodeError::limited(DecodeLimit::Text {
                observed: probe.text,
                maximum: probe.options.max_text_bytes,
            }));
        }
        Ok(())
    }
    const fn report(&self) -> DecodeReport {
        DecodeReport {
            bytes: self.bytes,
            fields: self.fields,
            work: self.work,
            max_depth: self.max_depth,
            text_bytes: self.text,
            node_count: self.nodes,
            precedent_count: self.precedents,
            range_count: self.ranges,
            unsupported_local_count: self.unsupported_local,
            evaluator_supported: self.evaluator_supported,
            allocations: 0,
        }
    }
}

fn checked(left: usize, right: usize) -> Result<usize, DecodeError> {
    left.checked_add(right).ok_or_else(malformed)
}

const fn malformed() -> DecodeError {
    DecodeError::invalid(InvalidReason::MalformedWire)
}

#[derive(Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire: u8,
    value: Value<'source>,
}

#[derive(Clone, Copy)]
enum Value<'source> {
    Varint(u64),
    Fixed64(u64),
    Bytes(&'source [u8]),
    Fixed32,
}

impl Field<'_> {
    fn varint(self) -> Result<u64, DecodeError> {
        match self.value {
            Value::Varint(value) if self.wire == 0 => Ok(value),
            _ => Err(malformed()),
        }
    }
    fn varint_u32(self) -> Result<u32, DecodeError> {
        u32::try_from(self.varint()?).map_err(|_error| malformed())
    }
    fn boolean(self) -> Result<bool, DecodeError> {
        match self.varint()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(malformed()),
        }
    }
    fn fixed64(self) -> Result<u64, DecodeError> {
        match self.value {
            Value::Fixed64(value) if self.wire == 1 => Ok(value),
            _ => Err(malformed()),
        }
    }
}

impl<'source> Field<'source> {
    fn bytes(self) -> Result<&'source [u8], DecodeError> {
        match self.value {
            Value::Bytes(value) if self.wire == 2 => Ok(value),
            _ => Err(malformed()),
        }
    }
}

fn next_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<Field<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.depth(depth)?;
    budget.field()?;
    let tag = take_varint(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_error| malformed())?;
    let wire = u8::try_from(tag & 7).map_err(|_error| malformed())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(malformed());
    }
    let value = match wire {
        0 => Value::Varint(take_varint(source)?),
        1 => {
            let raw = take(source, 8)?;
            Value::Fixed64(u64::from_le_bytes(
                raw.try_into().map_err(|_error| malformed())?,
            ))
        },
        2 => {
            let length = usize::try_from(take_varint(source)?).map_err(|_error| malformed())?;
            Value::Bytes(take(source, length)?)
        },
        5 => {
            let _raw = take(source, 4)?;
            Value::Fixed32
        },
        _ => return Err(malformed()),
    };
    Ok(Some(Field {
        number,
        wire,
        value,
    }))
}

fn take<'source>(source: &mut &'source [u8], amount: usize) -> Result<&'source [u8], DecodeError> {
    if source.len() < amount {
        return Err(malformed());
    }
    let (value, rest) = source.split_at(amount);
    *source = rest;
    Ok(value)
}

fn take_varint(source: &mut &[u8]) -> Result<u64, DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or_else(malformed)?;
        if index == 9 && byte > 1 {
            return Err(malformed());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            if varint_len(value) != consumed {
                return Err(malformed());
            }
            *source = &original[consumed..];
            return Ok(value);
        }
    }
    Err(malformed())
}

const fn varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn strict_utf8(source: &[u8]) -> Result<&str, DecodeError> {
    str::from_utf8(source).map_err(|_error| malformed())
}

const fn zigzag32(value: u32) -> i32 {
    ((value >> 1) as i32) ^ (-((value & 1) as i32))
}

const fn zigzag_encode_32(value: i32) -> u32 {
    ((value << 1) ^ (value >> 31)) as u32
}

#[cfg(test)]
mod tests {
    use super::*;

    fn put_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = value as u8 & 0x7f;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    fn key(output: &mut Vec<u8>, field: u32, wire: u8) {
        put_varint(output, (u64::from(field) << 3) | u64::from(wire));
    }

    fn varint(output: &mut Vec<u8>, field: u32, value: u64) {
        key(output, field, 0);
        put_varint(output, value);
    }

    fn bytes(output: &mut Vec<u8>, field: u32, value: &[u8]) {
        key(output, field, 2);
        put_varint(output, value.len() as u64);
        output.extend_from_slice(value);
    }

    fn fixed64(output: &mut Vec<u8>, field: u32, value: u64) {
        key(output, field, 1);
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn zigzag(value: i32) -> u32 {
        ((value << 1) ^ (value >> 31)) as u32
    }

    fn node(kind: u32, fields: impl FnOnce(&mut Vec<u8>)) -> Vec<u8> {
        let mut output = Vec::new();
        varint(&mut output, 1, u64::from(kind));
        fields(&mut output);
        output
    }

    fn formula(nodes: &[Vec<u8>]) -> Vec<u8> {
        let mut ast = Vec::new();
        for node in nodes {
            bytes(&mut ast, 1, node);
        }
        let mut formula = Vec::new();
        bytes(&mut formula, 1, &ast);
        formula
    }

    fn context() -> FormulaContext {
        FormulaContext::new(7, 4, 5, 20, 20)
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(source.len(), 2_000_000, 100_000_000, 16, 100_000, 1_000_000)
    }

    #[derive(Default)]
    struct Facts {
        nodes: Vec<FormulaNode>,
        precedents: Vec<LocalPrecedent>,
    }

    impl FormulaVisitor for Facts {
        fn visit_node(&mut self, node: FormulaNode) -> Result<(), DecodeError> {
            self.nodes.push(node);
            Ok(())
        }
        fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
            self.precedents.push(precedent);
            Ok(())
        }
    }

    #[test]
    fn supported_postfix_nodes_and_owner_aware_precedents_stream_in_order() {
        let number = node(17, |node| fixed64(node, 4, 3.5f64.to_bits()));
        let local = node(27, |node| {
            let mut reference = Vec::new();
            varint(&mut reference, 1, 2);
            varint(&mut reference, 2, 3);
            varint(&mut reference, 3, 1);
            varint(&mut reference, 4, 0);
            bytes(node, 15, &reference);
        });
        let relative = node(36, |node| {
            let mut row = Vec::new();
            varint(&mut row, 1, u64::from(zigzag(-1)));
            let mut column = Vec::new();
            varint(&mut column, 1, u64::from(zigzag(2)));
            bytes(node, 26, &column);
            bytes(node, 27, &row);
        });
        let add = node(1, |_| {});
        let function = node(16, |node| {
            varint(node, 2, 168);
            varint(node, 3, 2);
        });
        let source = formula(&[number, local, relative, add, function]);
        let mut facts = Facts::default();
        let report =
            decode_formula_archive_with_visitor(&source, context(), options(&source), &mut facts)
                .unwrap();
        assert_eq!(report.node_count(), 5);
        assert_eq!(report.precedent_count(), 2);
        assert_eq!(report.allocations(), 0);
        assert_eq!(
            facts.nodes[0],
            FormulaNode::Number {
                bits: 3.5f64.to_bits()
            }
        );
        assert_eq!(
            facts.precedents,
            vec![
                LocalPrecedent {
                    owner: 7,
                    coordinate: CellCoordinate { row: 2, column: 3 }
                },
                LocalPrecedent {
                    owner: 7,
                    coordinate: CellCoordinate { row: 3, column: 7 }
                }
            ]
        );
        assert_eq!(
            facts.nodes.last(),
            Some(&FormulaNode::Function {
                identifier: 168,
                argument_count: 2
            })
        );
    }

    fn reason(error: DecodeError) -> InvalidReason {
        error.invalid_reason().unwrap()
    }

    #[test]
    fn malformed_duplicate_unknown_and_external_forms_fail_closed() {
        let mut duplicate = node(17, |node| fixed64(node, 4, 1.0f64.to_bits()));
        varint(&mut duplicate, 1, 17);
        let duplicate = formula(&[duplicate]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &duplicate,
                    context(),
                    options(&duplicate),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::DuplicateField
        );

        let mut unknown = node(17, |node| fixed64(node, 4, 1.0f64.to_bits()));
        varint(&mut unknown, 48, 1);
        let unknown = formula(&[unknown]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &unknown,
                    context(),
                    options(&unknown),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::UnexpectedField
        );

        let wrong_wire = formula(&[node(19, |node| varint(node, 6, 1))]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &wrong_wire,
                    context(),
                    options(&wrong_wire),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let noncanonical = [0x0a, 0x02, 0x0a, 0x80];
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &noncanonical,
                    context(),
                    options(&noncanonical),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let cross = formula(&[node(28, |_| {})]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(&cross, context(), options(&cross), &mut ())
                    .unwrap_err()
            ),
            InvalidReason::ExternalOwner
        );
        let unsupported = formula(&[node(19, |_| {})]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &unsupported,
                    context(),
                    options(&unsupported),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::UnsupportedFormula
        );
    }

    #[derive(Default)]
    struct DependencyFacts {
        precedents: Vec<LocalPrecedent>,
        unsupported: Vec<UnsupportedLocal>,
        calls: usize,
    }

    impl FormulaDependencyVisitor for DependencyFacts {
        fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
            self.calls += 1;
            self.precedents.push(precedent);
            Ok(())
        }
        fn visit_unsupported_local(&mut self, node: UnsupportedLocal) -> Result<(), DecodeError> {
            self.calls += 1;
            self.unsupported.push(node);
            Ok(())
        }
    }

    #[test]
    fn unsupported_local_unreachable_is_dependency_visible_but_impacted_evaluation_refuses() {
        let string = node(19, |node| bytes(node, 6, b"preserve"));
        let local = node(27, |node| {
            let mut reference = Vec::new();
            varint(&mut reference, 1, 2);
            varint(&mut reference, 2, 3);
            varint(&mut reference, 3, 0);
            varint(&mut reference, 4, 0);
            bytes(node, 15, &reference);
        });
        let unknown_function = node(16, |node| {
            varint(node, 2, 999);
            varint(node, 3, 1);
        });
        let source = formula(&[string, local, unknown_function]);
        let mut facts = DependencyFacts::default();
        let report = inspect_formula_dependencies_with_visitor(
            &source,
            context(),
            options(&source),
            &mut facts,
        )
        .unwrap();
        assert_eq!(report.node_count(), 3);
        assert_eq!(report.precedent_count(), 1);
        assert_eq!(report.unsupported_local_count(), 2);
        assert_eq!(facts.precedents[0].owner(), 7);
        assert_eq!(
            facts.unsupported,
            vec![
                UnsupportedLocal {
                    node_type: 19,
                    function_identifier: None
                },
                UnsupportedLocal {
                    node_type: 16,
                    function_identifier: Some(999)
                }
            ]
        );
        assert_eq!(
            reason(inspect_formula_archive(&source, context(), options(&source)).unwrap_err()),
            InvalidReason::UnsupportedFormula
        );

        let external = formula(&[node(28, |_| {})]);
        assert_eq!(
            reason(
                inspect_formula_dependencies_with_visitor(
                    &external,
                    context(),
                    options(&external),
                    &mut DependencyFacts::default()
                )
                .unwrap_err()
            ),
            InvalidReason::ExternalOwner
        );

        let hidden_local = formula(&[node(20, |node| {
            let mut reference = Vec::new();
            varint(&mut reference, 1, 2);
            varint(&mut reference, 2, 3);
            varint(&mut reference, 3, 0);
            varint(&mut reference, 4, 0);
            bytes(node, 15, &reference);
        })]);
        let mut hidden_facts = DependencyFacts::default();
        assert_eq!(
            reason(
                inspect_formula_dependencies_with_visitor(
                    &hidden_local,
                    context(),
                    options(&hidden_local),
                    &mut hidden_facts
                )
                .unwrap_err()
            ),
            InvalidReason::UnsupportedFormula
        );
        assert_eq!(hidden_facts.calls, 0);
    }

    #[test]
    fn dependency_max_minus_one_preempts_unsupported_and_precedent_callbacks() {
        let source = formula(&[
            node(19, |node| bytes(node, 6, b"x")),
            node(17, |node| fixed64(node, 4, 1.0f64.to_bits())),
        ]);
        let report = inspect_formula_dependencies_with_visitor(
            &source,
            context(),
            options(&source),
            &mut DependencyFacts::default(),
        )
        .unwrap();
        for limited in [
            DecodeOptions::new(
                source.len(),
                report.fields() - 1,
                report.work(),
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work() - 1,
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work(),
                report.max_depth(),
                report.node_count() - 1,
                report.text_bytes(),
            ),
        ] {
            let mut facts = DependencyFacts::default();
            let error =
                inspect_formula_dependencies_with_visitor(&source, context(), limited, &mut facts)
                    .unwrap_err();
            assert!(error.resource_limit().is_some());
            assert_eq!(facts.calls, 0);
        }
    }

    #[test]
    fn native_decimal_number_shape_is_local_but_reference_bearing_near_shape_refuses() {
        let decimal = formula(&[node(17, |node| {
            fixed64(node, 4, 1.25f64.to_bits());
            varint(node, 42, 1);
            varint(node, 43, 2);
        })]);
        let mut facts = DependencyFacts::default();
        let report = inspect_formula_dependencies_with_visitor(
            &decimal,
            context(),
            options(&decimal),
            &mut facts,
        )
        .unwrap();
        assert_eq!(report.unsupported_local_count(), 1);
        assert_eq!(facts.unsupported[0].node_type(), 17);
        assert_eq!(
            reason(inspect_formula_archive(&decimal, context(), options(&decimal)).unwrap_err()),
            InvalidReason::UnsupportedFormula
        );

        let hidden_reference = formula(&[node(17, |node| {
            fixed64(node, 4, 1.25f64.to_bits());
            varint(node, 42, 1);
            varint(node, 43, 2);
            let mut reference = Vec::new();
            varint(&mut reference, 1, 2);
            varint(&mut reference, 2, 3);
            varint(&mut reference, 3, 0);
            varint(&mut reference, 4, 0);
            bytes(node, 15, &reference);
        })]);
        let mut hidden_facts = DependencyFacts::default();
        assert_eq!(
            reason(
                inspect_formula_dependencies_with_visitor(
                    &hidden_reference,
                    context(),
                    options(&hidden_reference),
                    &mut hidden_facts
                )
                .unwrap_err()
            ),
            InvalidReason::UnsupportedFormula
        );
        assert_eq!(hidden_facts.calls, 0);
    }

    #[derive(Default)]
    struct Calls(usize);

    impl FormulaVisitor for Calls {
        fn visit_node(&mut self, _node: FormulaNode) -> Result<(), DecodeError> {
            self.0 += 1;
            Ok(())
        }
        fn visit_precedent(&mut self, _precedent: LocalPrecedent) -> Result<(), DecodeError> {
            self.0 += 1;
            Ok(())
        }
    }

    #[test]
    fn exact_limits_are_inclusive_and_max_minus_one_preempts_callbacks() {
        let source = formula(&[
            node(17, |node| fixed64(node, 4, 1.0f64.to_bits())),
            node(32, |node| bytes(node, 25, b" ")),
        ]);
        let report = decode_formula_archive_with_visitor(
            &source,
            context(),
            options(&source),
            &mut Calls::default(),
        )
        .unwrap();
        let inspection = inspect_formula_archive(&source, context(), options(&source)).unwrap();
        assert_eq!(inspection.node_count(), report.node_count());
        assert_eq!(inspection.fields() * 2, report.fields());
        assert_eq!(inspection.work() * 2, report.work());
        let exact = DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work(),
            report.max_depth(),
            report.node_count(),
            report.text_bytes(),
        );
        assert!(
            decode_formula_archive_with_visitor(&source, context(), exact, &mut Calls::default())
                .is_ok()
        );
        let limited = [
            DecodeOptions::new(
                source.len() - 1,
                report.fields(),
                report.work(),
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields() - 1,
                report.work(),
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work() - 1,
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work(),
                report.max_depth(),
                report.node_count() - 1,
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work(),
                report.max_depth(),
                report.node_count(),
                report.text_bytes() - 1,
            ),
        ];
        for options in limited {
            let mut calls = Calls::default();
            let error =
                decode_formula_archive_with_visitor(&source, context(), options, &mut calls)
                    .unwrap_err();
            assert!(error.resource_limit().is_some());
            assert_eq!(calls.0, 0);
        }
    }

    #[test]
    fn four_thousand_to_eight_thousand_nodes_scale_below_two_point_two() {
        fn run(count: usize) -> DecodeReport {
            let nodes = (0..count)
                .map(|index| node(17, |node| fixed64(node, 4, (index as f64).to_bits())))
                .collect::<Vec<_>>();
            let source = formula(&nodes);
            decode_formula_archive_with_visitor(
                &source,
                context(),
                options(&source),
                &mut Calls::default(),
            )
            .unwrap()
        }
        let four = run(4_096);
        let eight = run(8_192);
        assert_eq!(eight.node_count(), four.node_count() * 2);
        assert!(eight.fields() * 100 <= four.fields() * 220);
        assert!(eight.work() * 100 <= four.work() * 220);
        assert!(eight.text_bytes() * 100 <= four.text_bytes().max(1) * 220);
        assert_eq!(eight.allocations(), 0);
    }

    fn write_options(
        bytes: usize,
        fields: usize,
        work: usize,
        nodes: usize,
        text: usize,
        depth: u32,
    ) -> DecodeOptions {
        DecodeOptions::new(bytes, fields, work, depth, nodes, text)
    }

    fn broad_write_options() -> DecodeOptions {
        write_options(
            8 * 1024 * 1024,
            1_000_000,
            100_000_000,
            100_000,
            1_000_000,
            16,
        )
    }

    #[test]
    fn writer_canonicalizes_every_frozen_node_and_matches_strict_and_buffa_views() {
        let nodes = [
            FormulaWriteNode::Number(2.0),
            FormulaWriteNode::Text("x"),
            FormulaWriteNode::Binary(BinaryOperator::Concatenate),
            FormulaWriteNode::Boolean(true),
            FormulaWriteNode::Percent,
            FormulaWriteNode::CellReference {
                column: FormulaWriteAxis::new(-2, false),
                row: FormulaWriteAxis::new(3, true),
            },
            FormulaWriteNode::Range,
            FormulaWriteNode::Function {
                identifier: 168,
                argument_count: 1,
            },
            FormulaWriteNode::Binary(BinaryOperator::Add),
            FormulaWriteNode::Negation,
        ];
        let options = broad_write_options();
        let plan = plan_formula_archive(&nodes, options).unwrap();
        let requirements = plan.requirements();
        let (output, report) = execute_formula_archive_plan(plan, options).unwrap();
        assert_eq!(output.len(), requirements.output_bytes());
        assert_eq!(output.capacity(), requirements.output_bytes());
        assert_eq!(requirements.allocations(), 1);
        assert_eq!(report.requirements(), requirements);

        let root: projection::FormulaArchiveLazyView<'_> =
            options.buffa().decode_lazy_view(output.as_slice()).unwrap();
        assert!(root.has_ast_node_array());
        assert!(root.host_column.is_none());
        let mut budget = Budget::new(&output, options).unwrap();
        let mut ast = root.ast_node_array;
        for expected in &nodes {
            let node = next_field(&mut ast, &mut budget, 2)
                .unwrap()
                .unwrap()
                .bytes()
                .unwrap();
            validate_canonical_write_node(node, expected, None, &[], &mut budget, options).unwrap();
        }
        assert!(ast.is_empty());
    }

    #[test]
    fn writer_rejects_noncanonical_scalars_and_malformed_postfix() {
        let options = broad_write_options();
        for nodes in [
            vec![FormulaWriteNode::Number(f64::NAN)],
            vec![FormulaWriteNode::Number(f64::INFINITY)],
            vec![FormulaWriteNode::Text("bad\0text")],
            vec![FormulaWriteNode::Binary(BinaryOperator::Add)],
            vec![FormulaWriteNode::Number(1.0), FormulaWriteNode::Number(2.0)],
            vec![FormulaWriteNode::Function {
                identifier: 0,
                argument_count: 0,
            }],
        ] {
            assert!(plan_formula_archive(&nodes, options).is_err());
        }
    }

    #[test]
    fn writer_number_and_every_binary_operator_have_frozen_canonical_tags() {
        let options = broad_write_options();
        let (number, _) =
            encode_formula_archive(&[FormulaWriteNode::Number(2.0)], options).unwrap();
        assert_eq!(
            number,
            [
                0x0a, 0x1b, 0x0a, 0x19, 0x08, 0x11, 0x21, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x40, 0xd0, 0x02, 0x02, 0xd8, 0x02, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0xa0,
                0x30,
            ]
        );

        for (operator, kind) in [
            (BinaryOperator::Add, 1),
            (BinaryOperator::Subtract, 2),
            (BinaryOperator::Multiply, 3),
            (BinaryOperator::Divide, 4),
            (BinaryOperator::Power, 5),
            (BinaryOperator::Concatenate, 6),
            (BinaryOperator::GreaterThan, 7),
            (BinaryOperator::GreaterThanOrEqual, 8),
            (BinaryOperator::LessThan, 9),
            (BinaryOperator::LessThanOrEqual, 10),
            (BinaryOperator::Equal, 11),
            (BinaryOperator::NotEqual, 12),
        ] {
            let nodes = [
                FormulaWriteNode::Number(1.0),
                FormulaWriteNode::Number(2.0),
                FormulaWriteNode::Binary(operator),
            ];
            let (output, _) = encode_formula_archive(&nodes, options).unwrap();
            let root: projection::FormulaArchiveLazyView<'_> =
                options.buffa().decode_lazy_view(&output).unwrap();
            let mut ast = root.ast_node_array;
            let mut budget = Budget::new(&output, options).unwrap();
            let _left = next_field(&mut ast, &mut budget, 2).unwrap().unwrap();
            let _right = next_field(&mut ast, &mut budget, 2).unwrap().unwrap();
            let binary = next_field(&mut ast, &mut budget, 2)
                .unwrap()
                .unwrap()
                .bytes()
                .unwrap();
            let view: projection::ASTNodeArchiveLazyView<'_> =
                options.buffa().decode_lazy_view(binary).unwrap();
            assert_eq!(u32::try_from(view.node_type).unwrap(), kind);
            assert!(ast.is_empty());
        }
    }

    #[test]
    fn writer_preserves_direct_negative_zero_and_matching_decimal128_sign() {
        let options = broad_write_options();
        let (output, _) =
            encode_formula_archive(&[FormulaWriteNode::Number(-0.0)], options).unwrap();
        let root: projection::FormulaArchiveLazyView<'_> =
            options.buffa().decode_lazy_view(&output).unwrap();
        let mut ast = root.ast_node_array;
        let mut budget = Budget::new(&output, options).unwrap();
        let node = next_field(&mut ast, &mut budget, 2)
            .unwrap()
            .unwrap()
            .bytes()
            .unwrap();
        let view: projection::ASTNodeArchiveLazyView<'_> =
            options.buffa().decode_lazy_view(node).unwrap();
        assert_eq!(view.number.unwrap().to_bits(), (-0.0f64).to_bits());
        assert_eq!(view.decimal_low, Some(0));
        assert_eq!(view.decimal_high, Some(0xb040_0000_0000_0000));
        assert!(ast.is_empty());

        assert_eq!(
            formula_decimal128_parts(0.75).unwrap(),
            (75, 0x303c_0000_0000_0000)
        );
        assert_eq!(
            formula_decimal128_parts(1e-5).unwrap(),
            (1, 0x3036_0000_0000_0000)
        );
    }

    #[test]
    fn writer_exact_requirements_are_inclusive_and_each_max_minus_one_refuses() {
        let nodes = [
            FormulaWriteNode::Text("bounded"),
            FormulaWriteNode::CellReference {
                column: FormulaWriteAxis::new(1, false),
                row: FormulaWriteAxis::new(2, true),
            },
            FormulaWriteNode::Binary(BinaryOperator::Add),
        ];
        let broad = broad_write_options();
        let requirements = plan_formula_archive(&nodes, broad).unwrap().requirements();
        let exact = write_options(
            requirements.output_bytes(),
            requirements.fields(),
            requirements.work_bytes(),
            requirements.nodes(),
            requirements.text_bytes(),
            requirements.max_depth(),
        );
        let plan = plan_formula_archive(&nodes, exact).unwrap();
        let (output, report) = execute_formula_archive_plan(plan, exact).unwrap();
        assert_eq!(output.capacity(), output.len());
        assert_eq!(report.requirements().allocations(), 1);

        let limited = [
            write_options(
                requirements.output_bytes() - 1,
                requirements.fields(),
                requirements.work_bytes(),
                requirements.nodes(),
                requirements.text_bytes(),
                requirements.max_depth(),
            ),
            write_options(
                requirements.output_bytes(),
                requirements.fields() - 1,
                requirements.work_bytes(),
                requirements.nodes(),
                requirements.text_bytes(),
                requirements.max_depth(),
            ),
            write_options(
                requirements.output_bytes(),
                requirements.fields(),
                requirements.work_bytes() - 1,
                requirements.nodes(),
                requirements.text_bytes(),
                requirements.max_depth(),
            ),
            write_options(
                requirements.output_bytes(),
                requirements.fields(),
                requirements.work_bytes(),
                requirements.nodes() - 1,
                requirements.text_bytes(),
                requirements.max_depth(),
            ),
            write_options(
                requirements.output_bytes(),
                requirements.fields(),
                requirements.work_bytes(),
                requirements.nodes(),
                requirements.text_bytes() - 1,
                requirements.max_depth(),
            ),
            write_options(
                requirements.output_bytes(),
                requirements.fields(),
                requirements.work_bytes(),
                requirements.nodes(),
                requirements.text_bytes(),
                requirements.max_depth() - 1,
            ),
        ];
        for limited in limited {
            let error = plan_formula_archive(&nodes, limited).unwrap_err();
            assert!(error.resource_limit().is_some());
        }
    }

    #[test]
    fn writer_execute_rechecks_lowered_ceiling_before_reserving_output() {
        let nodes = [FormulaWriteNode::Number(1.0)];
        let broad = broad_write_options();
        let plan = plan_formula_archive(&nodes, broad).unwrap();
        let requirements = plan.requirements();
        let lowered = write_options(
            requirements.output_bytes() - 1,
            requirements.fields(),
            requirements.work_bytes(),
            requirements.nodes(),
            requirements.text_bytes(),
            requirements.max_depth(),
        );
        assert!(matches!(
            execute_formula_archive_plan(plan, lowered)
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Bytes { .. })
        ));
    }

    #[test]
    fn writer_allocation_failure_is_typed_and_content_free() {
        assert!(matches!(
            reserve_formula_output(usize::MAX)
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Allocation {
                requested: usize::MAX
            })
        ));
    }

    #[test]
    fn writer_preempts_node_text_work_field_and_byte_limits_during_planning() {
        let many = vec![FormulaWriteNode::Number(1.0); 8_192];
        assert!(matches!(
            plan_formula_archive(
                &many,
                write_options(usize::MAX, usize::MAX, usize::MAX, 1, usize::MAX, 16)
            )
            .unwrap_err()
            .resource_limit(),
            Some(DecodeLimit::Nodes {
                observed: 8192,
                maximum: 1
            })
        ));

        let text = [FormulaWriteNode::Text("later\0")];
        assert!(matches!(
            plan_formula_archive(
                &text,
                write_options(usize::MAX, usize::MAX, usize::MAX, 1, 1, 16)
            )
            .unwrap_err()
            .resource_limit(),
            Some(DecodeLimit::Text { .. })
        ));
        assert!(matches!(
            plan_formula_archive(
                &text,
                write_options(usize::MAX, usize::MAX, 7, 1, usize::MAX, 16)
            )
            .unwrap_err()
            .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));

        let nodes = [FormulaWriteNode::Number(1.0), FormulaWriteNode::Negation];
        assert!(matches!(
            plan_formula_archive(&nodes, write_options(usize::MAX, 2, usize::MAX, 2, 0, 16))
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        assert!(matches!(
            plan_formula_archive(&nodes, write_options(3, usize::MAX, usize::MAX, 2, 0, 16))
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Bytes { .. })
        ));
    }

    #[test]
    fn writer_four_thousand_to_eight_thousand_nodes_is_strictly_linear() {
        fn requirements(count: usize) -> FormulaWriteRequirements {
            let mut nodes = Vec::with_capacity(count);
            nodes.push(FormulaWriteNode::Number(1.0));
            for _ in 1..count {
                nodes.push(FormulaWriteNode::Negation);
            }
            plan_formula_archive(&nodes, broad_write_options())
                .unwrap()
                .requirements()
        }
        let four = requirements(4_096);
        let eight = requirements(8_192);
        assert_eq!(eight.nodes(), four.nodes() * 2);
        assert!(eight.output_bytes() <= four.output_bytes() * 2);
        assert!(eight.ast_bytes() <= four.ast_bytes() * 2);
        assert!(eight.fields() <= four.fields() * 2);
        assert!(eight.work_bytes() <= four.work_bytes() * 2);
        assert_eq!(four.allocations(), 1);
        assert_eq!(eight.allocations(), 1);
    }

    fn resolved_context() -> (FormulaWriteContext, [ResolvedFormulaWriteOwner; 1]) {
        (
            FormulaWriteContext::new(7, 4, 5, 20, 20),
            [ResolvedFormulaWriteOwner::new(
                FormulaWriteOwnerUid::from_halves(0xe24e_294b_3170_bbd8, 0xbd57_cadc_10f6_658a),
                9,
                8,
                10,
            )],
        )
    }

    fn read_registry(owners: &[ResolvedFormulaWriteOwner]) -> ResolvedFormulaReadRegistry {
        let mut order: Vec<_> = (0..owners.len()).collect();
        order.sort_unstable_by_key(|index| owners[*index].internal_owner);
        let options = broad_write_options();
        let plan = plan_resolved_formula_read_registry(owners, &order, options).unwrap();
        execute_resolved_formula_read_registry_plan(plan, options).unwrap()
    }

    #[derive(Default)]
    struct WriteFacts {
        precedents: Vec<FormulaWritePrecedent>,
        ranges: Vec<FormulaWriteRange>,
    }

    impl FormulaWriteDependencyVisitor for WriteFacts {
        fn visit_precedent(&mut self, value: FormulaWritePrecedent) -> Result<(), DecodeError> {
            self.precedents.push(value);
            Ok(())
        }
        fn visit_range(&mut self, value: FormulaWriteRange) -> Result<(), DecodeError> {
            self.ranges.push(value);
            Ok(())
        }
    }

    #[derive(Default)]
    struct ResolvedReadFacts {
        nodes: Vec<FormulaNode>,
        precedents: Vec<LocalPrecedent>,
        ranges: Vec<FormulaWriteRange>,
        text: Vec<String>,
    }

    impl FormulaVisitor for ResolvedReadFacts {
        fn visit_node(&mut self, node: FormulaNode) -> Result<(), DecodeError> {
            self.nodes.push(node);
            Ok(())
        }
        fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
            self.precedents.push(precedent);
            Ok(())
        }
        fn visit_range(&mut self, range: FormulaWriteRange) -> Result<(), DecodeError> {
            self.ranges.push(range);
            Ok(())
        }
        fn visit_text(&mut self, value: &str) -> Result<(), DecodeError> {
            self.text.push(value.to_owned());
            Ok(())
        }
    }

    impl FormulaDependencyVisitor for ResolvedReadFacts {
        fn visit_node(&mut self, node: FormulaNode) -> Result<(), DecodeError> {
            self.nodes.push(node);
            Ok(())
        }
        fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
            self.precedents.push(precedent);
            Ok(())
        }
        fn visit_range(&mut self, range: FormulaWriteRange) -> Result<(), DecodeError> {
            self.ranges.push(range);
            Ok(())
        }
        fn visit_text(&mut self, value: &str) -> Result<(), DecodeError> {
            self.text.push(value.to_owned());
            Ok(())
        }
    }

    #[test]
    fn resolved_local_cross_range_and_whole_axes_encode_and_stream_exact_facts() {
        let (context, owners) = resolved_context();
        let uid = owners[0].uid;
        let cases = [
            vec![FormulaWriteNode::ResolvedCellReference {
                owner: None,
                reference: FormulaWriteCellReference::new(2, 3, false, true),
            }],
            vec![FormulaWriteNode::ResolvedCellReference {
                owner: Some(uid),
                reference: FormulaWriteCellReference::new(1, 2, false, false),
            }],
            vec![FormulaWriteNode::ResolvedRange {
                owner: Some(uid),
                start: FormulaWriteCellReference::new(1, 1, false, false),
                end: FormulaWriteCellReference::new(2, 2, true, true),
            }],
            vec![FormulaWriteNode::WholeRows {
                owner: None,
                start: FormulaWriteAxisReference::new(1, false),
                end: FormulaWriteAxisReference::new(2, true),
            }],
            vec![FormulaWriteNode::WholeColumns {
                owner: Some(uid),
                start: FormulaWriteAxisReference::new(1, false),
                end: FormulaWriteAxisReference::new(2, true),
            }],
        ];
        for nodes in cases {
            let options = broad_write_options();
            let plan = plan_resolved_formula_archive(
                &nodes,
                context,
                &owners,
                FormulaWriteDependencyLimits::new(1_000, 10),
                options,
            )
            .unwrap();
            let requirements = plan.requirements();
            let mut facts = WriteFacts::default();
            let (output, report) =
                execute_formula_archive_plan_with_visitor(plan, options, &mut facts).unwrap();
            assert_eq!(output.len(), requirements.output_bytes());
            assert_eq!(output.capacity(), output.len());
            assert_eq!(facts.precedents.len(), requirements.precedent_count());
            assert_eq!(facts.ranges.len(), requirements.range_count());
            assert_eq!(report.requirements(), requirements);
        }
    }

    #[test]
    fn resolved_writer_outputs_roundtrip_through_evaluator_and_dependency_readers() {
        let (context, external) = resolved_context();
        let uid = external[0].uid;
        let complete = [
            ResolvedFormulaWriteOwner::new(
                FormulaWriteOwnerUid::from_halves(1, 1),
                context.internal_owner,
                context.rows,
                context.columns,
            ),
            external[0],
        ];
        let registry = read_registry(&complete);
        let cases = [
            vec![FormulaWriteNode::Number(-0.0)],
            vec![FormulaWriteNode::Text("text")],
            vec![
                FormulaWriteNode::Text("a"),
                FormulaWriteNode::Text("b"),
                FormulaWriteNode::Binary(BinaryOperator::Concatenate),
            ],
            vec![FormulaWriteNode::Boolean(true)],
            vec![FormulaWriteNode::ResolvedCellReference {
                owner: None,
                reference: FormulaWriteCellReference::new(2, 3, false, true),
            }],
            vec![FormulaWriteNode::ResolvedCellReference {
                owner: Some(uid),
                reference: FormulaWriteCellReference::new(1, 2, false, false),
            }],
            vec![FormulaWriteNode::ResolvedRange {
                owner: None,
                start: FormulaWriteCellReference::new(2, 3, false, false),
                end: FormulaWriteCellReference::new(3, 4, true, true),
            }],
            vec![FormulaWriteNode::ResolvedRange {
                owner: Some(uid),
                start: FormulaWriteCellReference::new(1, 1, false, false),
                end: FormulaWriteCellReference::new(2, 2, true, true),
            }],
            vec![FormulaWriteNode::WholeRows {
                owner: None,
                start: FormulaWriteAxisReference::new(1, false),
                end: FormulaWriteAxisReference::new(2, true),
            }],
            vec![FormulaWriteNode::WholeRows {
                owner: Some(uid),
                start: FormulaWriteAxisReference::new(1, false),
                end: FormulaWriteAxisReference::new(2, true),
            }],
            vec![FormulaWriteNode::WholeColumns {
                owner: None,
                start: FormulaWriteAxisReference::new(1, false),
                end: FormulaWriteAxisReference::new(2, true),
            }],
            vec![FormulaWriteNode::WholeColumns {
                owner: Some(uid),
                start: FormulaWriteAxisReference::new(1, false),
                end: FormulaWriteAxisReference::new(2, true),
            }],
        ];
        for nodes in cases {
            let options = broad_write_options();
            let plan = plan_resolved_formula_archive(
                &nodes,
                context,
                &external,
                FormulaWriteDependencyLimits::new(1_000, 10),
                options,
            )
            .unwrap();
            let (output, _) = execute_formula_archive_plan(plan, options).unwrap();

            let mut evaluator = ResolvedReadFacts::default();
            let report = decode_resolved_formula_archive_with_visitor(
                &output,
                context,
                &registry,
                options,
                &mut evaluator,
            )
            .unwrap_or_else(|error| panic!("resolved evaluator refused {nodes:?}: {error:?}"));
            assert_eq!(report.node_count(), nodes.len());

            let mut dependencies = ResolvedReadFacts::default();
            let dependency_report = inspect_resolved_formula_dependencies_with_visitor(
                &output,
                context,
                &registry,
                options,
                &mut dependencies,
            )
            .unwrap_or_else(|error| panic!("resolved dependency refused {nodes:?}: {error:?}"));
            assert_eq!(dependencies.precedents, evaluator.precedents);
            assert_eq!(dependencies.ranges, evaluator.ranges);
            assert_eq!(
                dependency_report.precedent_count(),
                report.precedent_count()
            );
            assert_eq!(dependency_report.range_count(), report.range_count());
            assert_eq!(
                dependency_report.evaluator_supported(),
                !nodes.iter().any(|node| matches!(
                    node,
                    FormulaWriteNode::Text(_)
                        | FormulaWriteNode::Binary(BinaryOperator::Concatenate)
                ))
            );
        }
    }

    #[test]
    fn resolved_reader_accepts_only_legacy_or_exact_decimal_number_envelopes() {
        let (context, external) = resolved_context();
        let complete = [
            ResolvedFormulaWriteOwner::new(
                FormulaWriteOwnerUid::from_halves(1, 1),
                context.internal_owner,
                context.rows,
                context.columns,
            ),
            external[0],
        ];
        let registry = read_registry(&complete);
        let bits = 0.75f64.to_bits();
        let legacy = formula(&[node(17, |node| fixed64(node, 4, bits))]);
        let mut facts = ResolvedReadFacts::default();
        let report = decode_resolved_formula_archive_with_visitor(
            &legacy,
            context,
            &registry,
            broad_write_options(),
            &mut facts,
        )
        .unwrap();
        assert_eq!(report.node_count(), 1);
        assert_eq!(facts.nodes, vec![FormulaNode::Number { bits }]);

        for hostile in [
            formula(&[node(17, |node| {
                fixed64(node, 4, bits);
                varint(node, 42, 75);
            })]),
            formula(&[node(17, |node| {
                fixed64(node, 4, bits);
                varint(node, 43, 0x303c_0000_0000_0000);
            })]),
            formula(&[node(17, |node| {
                fixed64(node, 4, bits);
                varint(node, 42, 76);
                varint(node, 43, 0x303c_0000_0000_0000);
            })]),
        ] {
            let mut refused = ResolvedReadFacts::default();
            assert!(
                decode_resolved_formula_archive_with_visitor(
                    &hostile,
                    context,
                    &registry,
                    broad_write_options(),
                    &mut refused,
                )
                .is_err()
            );
            assert!(refused.nodes.is_empty());
        }
    }

    #[test]
    fn resolved_three_owner_lookup_work_is_reported_and_max_minus_one_preempts_callbacks() {
        let (context, external) = resolved_context();
        let uid = external[0].uid;
        let complete = [
            ResolvedFormulaWriteOwner::new(
                FormulaWriteOwnerUid::from_halves(1, 1),
                context.internal_owner,
                context.rows,
                context.columns,
            ),
            ResolvedFormulaWriteOwner::new(FormulaWriteOwnerUid::from_halves(2, 2), 8, 4, 4),
            external[0],
        ];
        let registry = read_registry(&complete);
        let nodes = [FormulaWriteNode::ResolvedCellReference {
            owner: Some(uid),
            reference: FormulaWriteCellReference::new(1, 2, false, false),
        }];
        let broad = broad_write_options();
        let plan = plan_resolved_formula_archive(
            &nodes,
            context,
            &external,
            FormulaWriteDependencyLimits::new(1, 0),
            broad,
        )
        .unwrap();
        let (output, _) = execute_formula_archive_plan(plan, broad).unwrap();
        let mut facts = ResolvedReadFacts::default();
        let report = decode_resolved_formula_archive_with_visitor(
            &output, context, &registry, broad, &mut facts,
        )
        .unwrap();
        assert_eq!(facts.precedents.len(), 1);

        let exact = write_options(
            output.len(),
            report.fields(),
            report.work(),
            complete.len(),
            report.text_bytes(),
            report.max_depth(),
        );
        let mut exact_facts = ResolvedReadFacts::default();
        decode_resolved_formula_archive_with_visitor(
            &output,
            context,
            &registry,
            exact,
            &mut exact_facts,
        )
        .unwrap();
        assert_eq!(exact_facts.precedents.len(), 1);

        let max_minus_one = write_options(
            output.len(),
            report.fields(),
            report.work() - 1,
            complete.len(),
            report.text_bytes(),
            report.max_depth(),
        );
        let mut refused = ResolvedReadFacts::default();
        assert!(matches!(
            decode_resolved_formula_archive_with_visitor(
                &output,
                context,
                &registry,
                max_minus_one,
                &mut refused,
            )
            .unwrap_err()
            .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));

        assert!(refused.nodes.is_empty());
        assert!(refused.precedents.is_empty());
        assert!(refused.ranges.is_empty());
    }

    #[test]
    fn resolved_registry_plan_is_linear_exact_and_rejects_hostile_permutations() {
        fn owners(count: usize) -> (Vec<ResolvedFormulaWriteOwner>, Vec<usize>) {
            let mut owners = Vec::with_capacity(count);
            for index in 0..count {
                owners.push(ResolvedFormulaWriteOwner::new(
                    FormulaWriteOwnerUid::from_halves(index as u64 + 1, 1),
                    u32::try_from(count - index).unwrap(),
                    2,
                    2,
                ));
            }
            let internal_order = (0..count).rev().collect();
            (owners, internal_order)
        }
        let (four_owners, four_order) = owners(4_096);
        let (eight_owners, eight_order) = owners(8_192);
        let broad = broad_write_options();
        let four = plan_resolved_formula_read_registry(&four_owners, &four_order, broad)
            .unwrap()
            .requirements();
        let eight = plan_resolved_formula_read_registry(&eight_owners, &eight_order, broad)
            .unwrap()
            .requirements();
        assert_eq!(eight.owners(), four.owners() * 2);
        assert_eq!(four.retained_elements(), four.owners() * 2);
        assert_eq!(eight.retained_elements(), four.retained_elements() * 2);
        assert!(eight.retained_bytes() <= four.retained_bytes() * 2);
        assert!(eight.work() <= four.work() * 2);
        assert_eq!(four.allocations(), 2);
        assert_eq!(eight.allocations(), 2);

        let exact = write_options(
            usize::MAX,
            usize::MAX,
            four.work(),
            four.owners(),
            usize::MAX,
            16,
        );
        let plan = plan_resolved_formula_read_registry(&four_owners, &four_order, exact).unwrap();
        let registry = execute_resolved_formula_read_registry_plan(plan, exact).unwrap();
        assert_eq!(registry.requirements(), four);

        let execute_nodes_max_minus_one = write_options(
            usize::MAX,
            usize::MAX,
            four.work(),
            four.owners() - 1,
            usize::MAX,
            16,
        );
        let plan = plan_resolved_formula_read_registry(&four_owners, &four_order, exact).unwrap();
        let error =
            match execute_resolved_formula_read_registry_plan(plan, execute_nodes_max_minus_one) {
                Ok(_registry) => panic!("lowered node ceiling must refuse before publication"),
                Err(error) => error,
            };
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nodes { .. })
        ));

        let plan = plan_resolved_formula_read_registry(&four_owners, &four_order, exact).unwrap();
        let execute_work_max_minus_one = write_options(
            usize::MAX,
            usize::MAX,
            four.work() - 1,
            four.owners(),
            usize::MAX,
            16,
        );
        let error =
            match execute_resolved_formula_read_registry_plan(plan, execute_work_max_minus_one) {
                Ok(_registry) => panic!("lowered work ceiling must refuse before publication"),
                Err(error) => error,
            };
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));

        let max_minus_one = write_options(
            usize::MAX,
            usize::MAX,
            four.work() - 1,
            four.owners(),
            usize::MAX,
            16,
        );
        assert!(matches!(
            plan_resolved_formula_read_registry(&four_owners, &four_order, max_minus_one)
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));

        let (small, order) = owners(3);
        for hostile in [vec![0, 1], vec![0, 1, 3], vec![0, 0, 2], vec![0, 1, 2]] {
            assert!(plan_resolved_formula_read_registry(&small, &hostile, broad).is_err());
        }
        assert!(plan_resolved_formula_read_registry(&small, &order, broad).is_ok());
    }

    #[test]
    fn resolved_shapes_have_exact_nested_field_and_depth_preflight() {
        let (context, owners) = resolved_context();
        let uid = owners[0].uid;
        let cases = [
            (
                vec![FormulaWriteNode::ResolvedCellReference {
                    owner: None,
                    reference: FormulaWriteCellReference::new(2, 3, false, true),
                }],
                9,
                4,
            ),
            (
                vec![FormulaWriteNode::ResolvedCellReference {
                    owner: Some(uid),
                    reference: FormulaWriteCellReference::new(1, 2, false, false),
                }],
                15,
                5,
            ),
            (
                vec![FormulaWriteNode::ResolvedRange {
                    owner: Some(uid),
                    start: FormulaWriteCellReference::new(1, 1, false, false),
                    end: FormulaWriteCellReference::new(2, 2, true, true),
                }],
                24,
                5,
            ),
            (
                vec![FormulaWriteNode::WholeRows {
                    owner: None,
                    start: FormulaWriteAxisReference::new(1, false),
                    end: FormulaWriteAxisReference::new(2, true),
                }],
                16,
                5,
            ),
            (
                vec![FormulaWriteNode::WholeColumns {
                    owner: Some(uid),
                    start: FormulaWriteAxisReference::new(1, false),
                    end: FormulaWriteAxisReference::new(2, true),
                }],
                22,
                5,
            ),
        ];
        for (nodes, expected_fields, expected_depth) in cases {
            let broad = broad_write_options();
            let requirements = plan_resolved_formula_archive(
                &nodes,
                context,
                &owners,
                FormulaWriteDependencyLimits::new(1_000, 10),
                broad,
            )
            .unwrap()
            .requirements();
            assert_eq!(requirements.fields(), expected_fields);
            assert_eq!(requirements.max_depth(), expected_depth);

            let exact = write_options(
                requirements.output_bytes(),
                requirements.fields(),
                requirements.work_bytes(),
                requirements.nodes(),
                requirements.text_bytes(),
                requirements.max_depth(),
            );
            let plan = plan_resolved_formula_archive(
                &nodes,
                context,
                &owners,
                FormulaWriteDependencyLimits::new(1_000, 10),
                exact,
            )
            .unwrap();
            execute_formula_archive_plan(plan, exact).unwrap();

            let fields_max_minus_one = write_options(
                requirements.output_bytes(),
                requirements.fields() - 1,
                requirements.work_bytes(),
                requirements.nodes(),
                requirements.text_bytes(),
                requirements.max_depth(),
            );
            assert!(matches!(
                plan_resolved_formula_archive(
                    &nodes,
                    context,
                    &owners,
                    FormulaWriteDependencyLimits::new(1_000, 10),
                    fields_max_minus_one,
                )
                .unwrap_err()
                .resource_limit(),
                Some(DecodeLimit::Fields { .. })
            ));

            let depth_max_minus_one = write_options(
                requirements.output_bytes(),
                requirements.fields(),
                requirements.work_bytes(),
                requirements.nodes(),
                requirements.text_bytes(),
                requirements.max_depth() - 1,
            );
            assert!(matches!(
                plan_resolved_formula_archive(
                    &nodes,
                    context,
                    &owners,
                    FormulaWriteDependencyLimits::new(1_000, 10),
                    depth_max_minus_one,
                )
                .unwrap_err()
                .resource_limit(),
                Some(DecodeLimit::Nesting { .. })
            ));
        }
    }

    #[test]
    fn resolved_negative_relative_range_is_sign_extended_canonical_int32() {
        let (context, owners) = resolved_context();
        let nodes = [FormulaWriteNode::ResolvedRange {
            owner: None,
            start: FormulaWriteCellReference::new(2, 3, false, false),
            end: FormulaWriteCellReference::new(3, 4, false, false),
        }];
        let options = broad_write_options();
        let plan = plan_resolved_formula_archive(
            &nodes,
            context,
            &owners,
            FormulaWriteDependencyLimits::new(4, 1),
            options,
        )
        .unwrap();
        let (output, _) = execute_formula_archive_plan(plan, options).unwrap();
        assert!(output.windows(11).any(|window| {
            window[0] == 0x08
                && window[1] == 0xfe
                && window[2..10].iter().all(|byte| *byte == 0xff)
                && window[10] == 0x01
        }));
    }

    #[test]
    fn resolved_owner_authority_and_dependency_admission_fail_before_output() {
        let (context, _owners) = resolved_context();
        let second_uid = FormulaWriteOwnerUid::from_halves(2, 3);
        let duplicate_owner = [
            ResolvedFormulaWriteOwner::new(FormulaWriteOwnerUid::from_halves(1, 1), 9, 2, 2),
            ResolvedFormulaWriteOwner::new(second_uid, 9, 2, 2),
        ];
        assert!(
            plan_resolved_formula_archive(
                &[FormulaWriteNode::Number(1.0)],
                context,
                &duplicate_owner,
                FormulaWriteDependencyLimits::new(0, 0),
                broad_write_options(),
            )
            .is_err()
        );

        let nodes = [FormulaWriteNode::WholeRows {
            owner: None,
            start: FormulaWriteAxisReference::new(0, false),
            end: FormulaWriteAxisReference::new(19, false),
        }];
        let error = plan_resolved_formula_archive(
            &nodes,
            context,
            &[],
            FormulaWriteDependencyLimits::new(399, 1),
            broad_write_options(),
        )
        .unwrap_err();
        assert!(error.resource_limit().is_some());

        let capped = write_options(usize::MAX, usize::MAX, usize::MAX, 1, usize::MAX, 16);
        assert!(matches!(
            plan_resolved_formula_archive(
                &[FormulaWriteNode::Number(1.0)],
                context,
                &duplicate_owner,
                FormulaWriteDependencyLimits::new(0, 0),
                capped,
            )
            .unwrap_err()
            .resource_limit(),
            Some(DecodeLimit::Nodes { .. })
        ));
        let owner_work_max_minus_one = write_options(usize::MAX, usize::MAX, 3, 2, usize::MAX, 16);
        assert!(matches!(
            plan_resolved_formula_archive(
                &[FormulaWriteNode::Number(1.0)],
                context,
                &duplicate_owner,
                FormulaWriteDependencyLimits::new(0, 0),
                owner_work_max_minus_one,
            )
            .unwrap_err()
            .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
    }
}
