//! Final-state formula dependency and display-cache planning.
//!
//! This is deliberately a plan-only layer.  It never publishes a package and
//! it never lets a streaming codec callback alter the graph under
//! construction: callbacks retain a small staged record, which is committed
//! to the graph only after the enclosing strict decoder and its Buffa parity
//! check have succeeded.

use core::mem::size_of;

use litchi_iwa_common::formula::{FiniteF64, FormulaCachedValue};
use litchi_iwa_common::{varint, wire::WireView};
use litchi_iwa_protos::{
    numbers_formula_codec as formula_codec, numbers_table_cell_dependency_codec as dependency,
    numbers_table_cell_storage_codec as storage,
};
#[cfg(test)]
use prost::Message as _;

const AVERAGE_FUNCTION_ID: u32 = 15;
const COUNT_FUNCTION_ID: u32 = 30;
const MAX_FUNCTION_ID: u32 = 84;
const MIN_FUNCTION_ID: u32 = 88;
const SUM_FUNCTION_ID: u32 = 168;
const PERCENT_SCALE: f64 = 100.0;
// The strict formula codec has a smaller structural ceiling than the shared
// dependency codecs. A caller may authorize more nesting, but passing that
// larger value is itself rejected by the formula decoder.
const FORMULA_NESTING_CEILING: u32 = 32;

/// Stable identity of the selected native formula owner/table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct TableIdentity {
    pub(super) owner: u32,
    pub(super) uuid_lower: u64,
    pub(super) uuid_upper: u64,
}

/// A table-local cell coordinate used by native dependency records.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct Coordinate {
    pub(super) row: u32,
    pub(super) column: u32,
}

/// One scalar final-batch overlay value.
#[derive(Debug, Clone)]
pub(super) struct FinalCell {
    pub(super) table: TableIdentity,
    pub(super) coordinate: Coordinate,
    pub(super) value: FinalValue,
}

/// Content-free final overlay value offered to formula evaluation.
///
/// Only numbers and Booleans are representable as evaluator inputs. Text,
/// date, and duration edits use [`Self::unsupported`], so the cache planner
/// never clones authored text or retains authored content it cannot model.
#[derive(Debug, Clone)]
pub(super) struct FinalValue(FinalValueKind);

#[derive(Debug, Clone)]
enum FinalValueKind {
    Clear,
    Supported(FormulaCachedValue),
    Unsupported,
}

impl FinalValue {
    pub(super) const fn clear() -> Self {
        Self(FinalValueKind::Clear)
    }

    pub(super) const fn number(value: FiniteF64) -> Self {
        Self(FinalValueKind::Supported(FormulaCachedValue::Number(value)))
    }

    pub(super) const fn boolean(value: bool) -> Self {
        Self(FinalValueKind::Supported(FormulaCachedValue::Boolean(
            value,
        )))
    }

    pub(super) const fn unsupported() -> Self {
        Self(FinalValueKind::Unsupported)
    }

    fn supported(&self) -> Result<Option<&FormulaCachedValue>, Failure> {
        match &self.0 {
            FinalValueKind::Clear => Ok(None),
            FinalValueKind::Supported(value) => Ok(Some(value)),
            FinalValueKind::Unsupported => {
                Err(Failure::UnsupportedDependency(Unsupported::Formula))
            },
        }
    }
}

/// A raw dependency object selected from the owning IWA component.
#[derive(Debug, Clone, Copy)]
pub(super) struct DependencyPayload<'source> {
    pub(super) identifier: u64,
    pub(super) bytes: &'source [u8],
}

/// One formula cell whose display cache can be rewritten.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct FormulaCell {
    /// Native internal formula-owner ID.
    pub(super) owner: u32,
    /// Physical table coordinate.
    pub(super) coordinate: Coordinate,
    /// IWA object containing this cell's formula cache entry.
    pub(super) cache_object: u64,
}

/// One formula host discovered in the selected owner's dependency records.
///
/// The storage join deliberately remains outside this module: callers match
/// this semantic key to the formula table-data list and only then attach the
/// cache object's route with [`FormulaHost::into_formula_cell`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct FormulaHost {
    pub(super) owner: u32,
    pub(super) coordinate: Coordinate,
}

impl FormulaHost {
    pub(super) const fn into_formula_cell(self, cache_object: u64) -> FormulaCell {
        FormulaCell {
            owner: self.owner,
            coordinate: self.coordinate,
            cache_object,
        }
    }
}

/// All byte sources used to construct a single dependency graph.
#[derive(Debug, Clone, Copy)]
pub(super) struct CacheSource<'source> {
    /// Owner/table to which every overlay coordinate belongs.
    pub(super) selected_table: TableIdentity,
    pub(super) rows: u32,
    pub(super) columns: u32,
    pub(super) header_rows: u32,
    pub(super) header_columns: u32,
    pub(super) footer_rows: u32,
    /// The one rooted `CalculationEngineArchive` payload.
    pub(super) engine: &'source [u8],
    /// Referenced `FormulaOwnerDependenciesArchive` payloads.
    pub(super) owners: &'source [DependencyPayload<'source>],
    /// Referenced `CellRecordTileArchive` payloads.
    pub(super) record_tiles: &'source [DependencyPayload<'source>],
    /// Formula cells in the selected table, indexed without reparsing storage.
    pub(super) formulas: &'source [FormulaCell],
}

/// A finite cache-planning policy.  Every limit is independent.
#[derive(Debug, Clone, Copy)]
pub(super) struct CacheLimits {
    pub(super) wire_bytes: usize,
    pub(super) wire_fields: usize,
    pub(super) wire_work: usize,
    pub(super) wire_references: usize,
    pub(super) wire_text: usize,
    pub(super) nesting: u32,
    pub(super) graph_nodes: usize,
    pub(super) graph_edges: usize,
    pub(super) cache_cells: usize,
    /// Evaluator-reported AST/evaluation operations.
    pub(super) formula_work: usize,
    /// Graph indexing, searches, edge scans, closure, and queue operations.
    pub(super) graph_work: usize,
    /// Total bytes retained by the successful output plan.
    pub(super) retained_bytes: usize,
    /// Conservative peak bytes retained by graph-planning scratch buffers.
    pub(super) scratch_bytes: usize,
    /// Fallible vector-growth events during graph planning.
    pub(super) allocations: usize,
}

/// Exact work performed by a successful cache plan.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub(super) struct CacheUsage {
    pub(super) wire_bytes: u64,
    pub(super) wire_fields: u64,
    pub(super) wire_work: u64,
    pub(super) wire_references: u64,
    pub(super) wire_text_bytes: u64,
    pub(super) lookup_work: u64,
    pub(super) queue_work: u64,
    pub(super) formula_graph_builds: u64,
    pub(super) formula_nodes: u64,
    pub(super) formula_work: u64,
    pub(super) graph_work: u64,
    pub(super) dependency_edges: u64,
    pub(super) dependency_range_queries: u64,
    pub(super) dependency_range_candidates: u64,
    pub(super) cache_cells_read: u64,
    pub(super) cache_hosts_refreshed: u64,
    pub(super) retained_elements: u64,
    pub(super) retained_bytes: u64,
    pub(super) peak_scratch_bytes: u64,
    pub(super) allocations: u64,
}

/// A non-publishing dependency/cache planning failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Failure {
    InvalidSource,
    UnsupportedDependency(Unsupported),
    LimitExceeded { observed: u64, maximum: u64 },
    Allocation { amount: usize },
}

/// The exact dependency feature this conservative scalar evaluator declines.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Unsupported {
    Range,
    Volatile,
    Spanning,
    WholeOwner,
    UuidReference,
    Spill,
    ExternalOwner,
    MissingOwner,
    Formula,
    CacheType,
    HeaderNameManager,
}

/// A cache cell rewrite.  The owner is retained to make raw list rewrites
/// deterministic when one cache object contains more than one formula owner.
#[derive(Debug, Clone)]
pub(super) struct CacheCellRewrite {
    pub(super) owner: u32,
    pub(super) coordinate: Coordinate,
    pub(super) value: FormulaCachedValue,
}

/// All cache-cell rewrites for one IWA cache object.
#[derive(Debug, Clone)]
pub(super) struct CacheRewrite {
    pub(super) cache_object: u64,
    pub(super) cells: Vec<CacheCellRewrite>,
}

/// An old formula dependency entry to delete after a formula cell is replaced
/// by a scalar or cleared in the final batch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct DependencyRemoval {
    pub(super) owner: u32,
    pub(super) coordinate: Coordinate,
}

/// A complete, directional final-state cache plan.
#[derive(Debug, Clone)]
pub(super) struct CachePlan {
    pub(super) rewrites: Vec<CacheRewrite>,
    pub(super) removals: Vec<DependencyRemoval>,
    pub(super) usage: CacheUsage,
}

/// Read a scalar cache value not supplied by the final-batch overlay.
pub(super) trait CacheBaseline {
    fn value(&self, coordinate: Coordinate) -> Result<Option<&FormulaCachedValue>, Failure>;
}

/// A final-state lookup offered to the supported formula evaluator.
pub(super) trait CacheValues {
    fn value(&self, coordinate: Coordinate) -> Result<Option<&FormulaCachedValue>, Failure>;
}

/// The storage kernel owns formula decoding.  This module owns dependency
/// reachability and final-state ordering, and asks the kernel to evaluate only
/// a scalar formula it recognizes. Implementations must honor the supplied
/// remaining limits before allocating or invoking externally visible work.
pub(super) trait CacheEvaluator {
    /// Strictly enumerate this formula's owner-qualified local precedents.
    fn analyze(
        &mut self,
        formula: FormulaCell,
        limits: CacheLimits,
    ) -> Result<FormulaAnalysis, Failure>;

    fn evaluate(
        &mut self,
        formula: FormulaCell,
        values: &dyn CacheValues,
        limits: CacheLimits,
    ) -> Result<Evaluation, Failure>;
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct FormulaPrecedent {
    pub(super) owner: u32,
    pub(super) coordinate: Coordinate,
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub(super) struct FormulaEvaluationUsage {
    pub(super) wire_bytes: usize,
    pub(super) wire_fields: usize,
    pub(super) wire_work: usize,
    pub(super) wire_text_bytes: usize,
    pub(super) formula_work: usize,
    pub(super) scratch_bytes: usize,
    pub(super) allocations: usize,
}

#[derive(Debug)]
pub(super) struct FormulaAnalysis {
    pub(super) precedents: Vec<FormulaPrecedent>,
    pub(super) usage: FormulaEvaluationUsage,
}

/// A scalar evaluator result with exact AST/evaluation work.
#[derive(Debug, Clone)]
pub(super) struct Evaluation {
    pub(super) value: FormulaCachedValue,
    pub(super) usage: FormulaEvaluationUsage,
}

/// One byte-exact formula table entry used by [`StrictEvaluator`].
#[derive(Debug, Clone, Copy)]
pub(super) struct FormulaPayload<'source> {
    pub(super) owner: u32,
    pub(super) coordinate: Coordinate,
    /// Native formula-list key stored in the host cell.
    pub(super) key: u32,
    pub(super) bytes: &'source [u8],
}

/// Table-local, deterministic evaluator for the legacy scalar parity subset.
///
/// The formula payload is decoded only for evaluation and is never re-encoded;
/// the cache plan consequently preserves the original formula AST bytes and
/// dependency objects exactly.
pub(super) struct StrictEvaluator<'source> {
    formulas: &'source [FormulaPayload<'source>],
    rows: u32,
    columns: u32,
}

impl<'source> StrictEvaluator<'source> {
    /// Construct an evaluator only after proving exact formula-list coverage.
    pub(super) fn new(
        entries: &[FormulaListEntry<'_>],
        formulas: &'source [FormulaPayload<'source>],
        rows: u32,
        columns: u32,
        selected: TableIdentity,
        limits: CacheLimits,
    ) -> Result<(Self, CacheUsage), Failure> {
        if rows == 0 || columns == 0 {
            return Err(Failure::InvalidSource);
        }
        let usage = validate_formula_coverage(selected, entries, formulas, limits)?;
        Ok((
            Self {
                formulas,
                rows,
                columns,
            },
            usage,
        ))
    }

    fn payload(&self, formula: FormulaCell) -> Result<FormulaPayload<'source>, Failure> {
        let key = (formula.owner, formula.coordinate);
        let position = self
            .formulas
            .binary_search_by(|entry| (entry.owner, entry.coordinate).cmp(&key))
            .map_err(|_error| Failure::InvalidSource)?;
        Ok(self.formulas[position])
    }

    fn context(&self, formula: FormulaCell) -> formula_codec::FormulaContext {
        formula_codec::FormulaContext::new(
            formula.owner,
            formula.coordinate.row,
            formula.coordinate.column,
            self.rows,
            self.columns,
        )
    }

    fn options(limits: CacheLimits) -> formula_codec::DecodeOptions {
        formula_codec::DecodeOptions::new(
            limits.wire_bytes,
            limits.wire_fields,
            limits.wire_work,
            limits.nesting.min(FORMULA_NESTING_CEILING),
            limits.formula_work.min(limits.graph_work),
            limits.wire_text,
        )
    }
}

impl CacheEvaluator for StrictEvaluator<'_> {
    fn analyze(
        &mut self,
        formula: FormulaCell,
        limits: CacheLimits,
    ) -> Result<FormulaAnalysis, Failure> {
        let payload = self.payload(formula)?;
        let context = self.context(formula);
        let options = Self::options(limits);
        // Global AST-to-edge proof must not require the scalar evaluator to
        // understand every canonical local node.  Run the strict
        // dependency-only projection twice: first allocation-free for exact
        // sizing, then into the exactly reserved precedent buffer.  It still
        // hard-refuses malformed, noncanonical, cross-owner, and UID forms.
        let probe = formula_codec::inspect_formula_dependencies_with_visitor(
            payload.bytes,
            context,
            options,
            &mut (),
        )
        .map_err(map_formula_decode_failure)?;
        let count = probe.precedent_count();
        let ordering_work = sort_work(count)?
            .checked_add(count)
            .ok_or(Failure::InvalidSource)?;
        preflight_formula_dependency_pass(probe, ordering_work, limits)?;
        check_limit(count, limits.graph_edges)?;
        let bytes = count
            .checked_mul(size_of::<FormulaPrecedent>())
            .ok_or(Failure::InvalidSource)?;
        check_limit(bytes, limits.scratch_bytes)?;
        if count != 0 {
            check_limit(1, limits.allocations)?;
        }
        let mut visitor = FormulaPrecedentCollector {
            precedents: Vec::new(),
            failure: None,
            maximum: count,
        };
        reserve(&mut visitor.precedents, count)?;
        if size_of::<FormulaPrecedent>() != 0 && visitor.precedents.capacity() != count {
            return Err(Failure::Allocation { amount: count });
        }
        let report = formula_codec::inspect_formula_dependencies_with_visitor(
            payload.bytes,
            context,
            options,
            &mut visitor,
        )
        .map_err(map_formula_decode_failure)?;
        if let Some(failure) = visitor.failure {
            return Err(failure);
        }
        if visitor.precedents.len() != count {
            return Err(Failure::InvalidSource);
        }
        let mut usage = formula_reports_usage(probe, report)?;
        usage.scratch_bytes = bytes;
        usage.allocations = usize::from(count != 0);
        usage.formula_work = usage
            .formula_work
            .checked_add(ordering_work)
            .ok_or(Failure::InvalidSource)?;
        check_limit(
            usage.formula_work,
            limits.formula_work.min(limits.graph_work),
        )?;
        visitor.precedents.sort_unstable();
        visitor.precedents.dedup();
        Ok(FormulaAnalysis {
            precedents: visitor.precedents,
            usage,
        })
    }

    fn evaluate(
        &mut self,
        formula: FormulaCell,
        values: &dyn CacheValues,
        limits: CacheLimits,
    ) -> Result<Evaluation, Failure> {
        let payload = self.payload(formula)?;
        let context = self.context(formula);
        let options = Self::options(limits);
        let probe = formula_codec::inspect_formula_archive(payload.bytes, context, options)
            .map_err(map_formula_decode_failure)?;
        let count = probe.node_count();
        let remaining_work = preflight_formula_pass(probe, 0, limits)?;
        let bytes = count
            .checked_mul(size_of::<StreamingValue>())
            .ok_or(Failure::InvalidSource)?;
        check_limit(bytes, limits.scratch_bytes)?;
        if count != 0 {
            check_limit(1, limits.allocations)?;
        }
        let maximum_work = limits.formula_work.min(limits.graph_work);
        let fixed_work = maximum_work
            .checked_sub(remaining_work)
            .ok_or(Failure::InvalidSource)?;
        let mut visitor = StreamingFormula::new(values, count, fixed_work, maximum_work)?;
        let report = formula_codec::decode_formula_archive_with_visitor(
            payload.bytes,
            context,
            options,
            &mut visitor,
        )
        .map_err(map_formula_decode_failure)?;
        let (value, evaluation_work) = visitor.finish()?;
        let mut usage = formula_reports_usage(probe, report)?;
        usage.formula_work = usage
            .formula_work
            .checked_add(evaluation_work)
            .ok_or(Failure::InvalidSource)?;
        check_limit(
            usage.formula_work,
            limits.formula_work.min(limits.graph_work),
        )?;
        usage.scratch_bytes = bytes;
        usage.allocations = usize::from(count != 0);
        Ok(Evaluation { value, usage })
    }
}

struct FormulaPrecedentCollector {
    precedents: Vec<FormulaPrecedent>,
    failure: Option<Failure>,
    maximum: usize,
}

impl formula_codec::FormulaDependencyVisitor for FormulaPrecedentCollector {
    fn visit_precedent(
        &mut self,
        precedent: formula_codec::LocalPrecedent,
    ) -> Result<(), formula_codec::DecodeError> {
        if self.failure.is_some() {
            return Ok(());
        }
        if self.precedents.len() >= self.maximum
            || self.precedents.len() == self.precedents.capacity()
        {
            self.failure = Some(Failure::InvalidSource);
            return Ok(());
        }
        let coordinate = precedent.coordinate();
        self.precedents.push(FormulaPrecedent {
            owner: precedent.owner(),
            coordinate: Coordinate {
                row: coordinate.row(),
                column: coordinate.column(),
            },
        });
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum StreamingScalar {
    Empty,
    Number(f64),
    Boolean(bool),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct StreamingRange {
    top: u32,
    left: u32,
    bottom: u32,
    right: u32,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum StreamingValue {
    Scalar(StreamingScalar),
    Reference(Coordinate),
    Range(StreamingRange),
}

struct StreamingFormula<'values> {
    values: &'values dyn CacheValues,
    stack: Vec<StreamingValue>,
    failure: Option<Failure>,
    fixed_work: usize,
    work: usize,
    maximum_work: usize,
}

impl<'values> StreamingFormula<'values> {
    fn new(
        values: &'values dyn CacheValues,
        capacity: usize,
        fixed_work: usize,
        maximum_work: usize,
    ) -> Result<Self, Failure> {
        let mut stack = Vec::new();
        reserve(&mut stack, capacity)?;
        if size_of::<StreamingValue>() != 0 && stack.capacity() != capacity {
            return Err(Failure::Allocation { amount: capacity });
        }
        Ok(Self {
            values,
            stack,
            failure: None,
            fixed_work,
            work: 0,
            maximum_work,
        })
    }

    fn finish(mut self) -> Result<(FormulaCachedValue, usize), Failure> {
        if let Some(failure) = self.failure {
            return Err(failure);
        }
        if self.stack.len() != 1 {
            return Err(Failure::InvalidSource);
        }
        let popped = self.stack.pop().ok_or(Failure::InvalidSource)?;
        let value = self.scalar(popped)?;
        let value = match value {
            StreamingScalar::Empty => FormulaCachedValue::number(0.0),
            StreamingScalar::Number(number) => FormulaCachedValue::number(number),
            StreamingScalar::Boolean(boolean) => Ok(FormulaCachedValue::Boolean(boolean)),
        }
        .map_err(|_error| Failure::UnsupportedDependency(Unsupported::Formula))?;
        Ok((value, self.work))
    }

    fn visit(&mut self, node: formula_codec::FormulaNode) -> Result<(), Failure> {
        use formula_codec::FormulaNode;

        self.step(1)?;
        match node {
            FormulaNode::Number { bits } => self.push(StreamingValue::Scalar(
                streaming_finite_number(f64::from_bits(bits))?,
            )),
            FormulaNode::Boolean(value) | FormulaNode::Token(value) => {
                self.push(StreamingValue::Scalar(StreamingScalar::Boolean(value)))
            },
            FormulaNode::Empty => self.push(StreamingValue::Scalar(StreamingScalar::Empty)),
            FormulaNode::LocalCell { coordinate, .. }
            | FormulaNode::CellReference { coordinate } => {
                self.push(StreamingValue::Reference(Coordinate {
                    row: coordinate.row(),
                    column: coordinate.column(),
                }))
            },
            FormulaNode::Colon | FormulaNode::ColonWithUids => {
                let end = self.pop()?;
                let start = self.pop()?;
                let (StreamingValue::Reference(start), StreamingValue::Reference(end)) =
                    (start, end)
                else {
                    return unsupported_formula();
                };
                self.push(StreamingValue::Range(StreamingRange {
                    top: start.row.min(end.row),
                    left: start.column.min(end.column),
                    bottom: start.row.max(end.row),
                    right: start.column.max(end.column),
                }))
            },
            FormulaNode::Binary(operator) => {
                let popped_right = self.pop()?;
                let right = self.scalar(popped_right)?;
                let popped_left = self.pop()?;
                let left = self.scalar(popped_left)?;
                let value = streaming_binary(operator, left, right)?;
                self.push(StreamingValue::Scalar(value))
            },
            FormulaNode::Negation => {
                let popped = self.pop()?;
                let value = streaming_scalar_number(self.scalar(popped)?);
                let value = streaming_finite_number(-value)?;
                self.push(StreamingValue::Scalar(value))
            },
            FormulaNode::Percent => {
                let popped = self.pop()?;
                let value = streaming_scalar_number(self.scalar(popped)?);
                let value = streaming_finite_number(value / PERCENT_SCALE)?;
                self.push(StreamingValue::Scalar(value))
            },
            FormulaNode::Function {
                identifier,
                argument_count,
            } => {
                let count =
                    usize::try_from(argument_count).map_err(|_error| Failure::InvalidSource)?;
                let start = self
                    .stack
                    .len()
                    .checked_sub(count)
                    .ok_or(Failure::InvalidSource)?;
                let value = self.evaluate_function(identifier, start)?;
                self.stack.truncate(start);
                self.push(StreamingValue::Scalar(value))
            },
            FormulaNode::PlusSign
            | FormulaNode::AppendWhitespace
            | FormulaNode::PrependWhitespace => Ok(()),
        }
    }

    fn push(&mut self, value: StreamingValue) -> Result<(), Failure> {
        if self.stack.len() == self.stack.capacity() {
            return Err(Failure::InvalidSource);
        }
        self.stack.push(value);
        Ok(())
    }

    fn pop(&mut self) -> Result<StreamingValue, Failure> {
        self.stack.pop().ok_or(Failure::InvalidSource)
    }

    fn step(&mut self, increment: usize) -> Result<(), Failure> {
        self.work = self
            .work
            .checked_add(increment)
            .ok_or(Failure::InvalidSource)?;
        let observed = self
            .fixed_work
            .checked_add(self.work)
            .ok_or(Failure::InvalidSource)?;
        check_limit(observed, self.maximum_work)
    }

    fn scalar(&mut self, value: StreamingValue) -> Result<StreamingScalar, Failure> {
        match value {
            StreamingValue::Scalar(value) => Ok(value),
            StreamingValue::Reference(coordinate) => {
                self.step(1)?;
                match self.values.value(coordinate)? {
                    None => Ok(StreamingScalar::Empty),
                    Some(FormulaCachedValue::Number(number)) => {
                        Ok(StreamingScalar::Number(number.get()))
                    },
                    Some(FormulaCachedValue::Boolean(boolean)) => {
                        Ok(StreamingScalar::Boolean(*boolean))
                    },
                    Some(
                        FormulaCachedValue::Text(_)
                        | FormulaCachedValue::Date(_)
                        | FormulaCachedValue::Duration(_),
                    ) => unsupported_formula(),
                }
            },
            StreamingValue::Range(_) => unsupported_formula(),
        }
    }

    fn evaluate_function(
        &mut self,
        identifier: u32,
        start: usize,
    ) -> Result<StreamingScalar, Failure> {
        if !matches!(
            identifier,
            AVERAGE_FUNCTION_ID
                | COUNT_FUNCTION_ID
                | MAX_FUNCTION_ID
                | MIN_FUNCTION_ID
                | SUM_FUNCTION_ID
        ) {
            return unsupported_formula();
        }
        let mut aggregate = StreamingAggregate::default();
        for index in start..self.stack.len() {
            let argument = self.stack[index];
            self.collect(argument, &mut aggregate)?;
        }
        let number = match identifier {
            SUM_FUNCTION_ID => aggregate.sum,
            COUNT_FUNCTION_ID => aggregate.count as f64,
            AVERAGE_FUNCTION_ID if aggregate.count != 0 => aggregate.sum / aggregate.count as f64,
            AVERAGE_FUNCTION_ID => return unsupported_formula(),
            MIN_FUNCTION_ID => aggregate.minimum.unwrap_or(0.0),
            MAX_FUNCTION_ID => aggregate.maximum.unwrap_or(0.0),
            _ => return Err(Failure::InvalidSource),
        };
        streaming_finite_number(number)
    }

    fn collect(
        &mut self,
        value: StreamingValue,
        aggregate: &mut StreamingAggregate,
    ) -> Result<(), Failure> {
        match value {
            StreamingValue::Scalar(StreamingScalar::Number(number)) => aggregate.push(number),
            StreamingValue::Scalar(StreamingScalar::Empty) => Ok(()),
            StreamingValue::Scalar(StreamingScalar::Boolean(_)) => unsupported_formula(),
            StreamingValue::Reference(coordinate) => {
                if let StreamingScalar::Number(number) =
                    self.scalar(StreamingValue::Reference(coordinate))?
                {
                    aggregate.push(number)?;
                }
                Ok(())
            },
            StreamingValue::Range(range) => {
                for row in range.top..=range.bottom {
                    for column in range.left..=range.right {
                        self.step(1)?;
                        if let StreamingScalar::Number(number) =
                            self.scalar(StreamingValue::Reference(Coordinate { row, column }))?
                        {
                            aggregate.push(number)?;
                        }
                    }
                }
                Ok(())
            },
        }
    }
}

impl formula_codec::FormulaVisitor for StreamingFormula<'_> {
    fn visit_node(
        &mut self,
        node: formula_codec::FormulaNode,
    ) -> Result<(), formula_codec::DecodeError> {
        if self.failure.is_none() {
            if let Err(failure) = self.visit(node) {
                self.failure = Some(failure);
            }
        }
        Ok(())
    }
}

#[derive(Debug, Default)]
struct StreamingAggregate {
    sum: f64,
    count: u64,
    minimum: Option<f64>,
    maximum: Option<f64>,
}

impl StreamingAggregate {
    fn push(&mut self, value: f64) -> Result<(), Failure> {
        if !value.is_finite() {
            return unsupported_formula();
        }
        self.sum += value;
        if !self.sum.is_finite() {
            return unsupported_formula();
        }
        self.count = self.count.checked_add(1).ok_or(Failure::InvalidSource)?;
        self.minimum = Some(self.minimum.map_or(value, |current| current.min(value)));
        self.maximum = Some(self.maximum.map_or(value, |current| current.max(value)));
        Ok(())
    }
}

fn streaming_scalar_number(value: StreamingScalar) -> f64 {
    match value {
        StreamingScalar::Empty => 0.0,
        StreamingScalar::Number(number) => number,
        StreamingScalar::Boolean(boolean) => f64::from(u8::from(boolean)),
    }
}

fn streaming_finite_number(value: f64) -> Result<StreamingScalar, Failure> {
    if value.is_finite() {
        Ok(StreamingScalar::Number(value))
    } else {
        unsupported_formula()
    }
}

fn streaming_binary(
    operator: formula_codec::BinaryOperator,
    left: StreamingScalar,
    right: StreamingScalar,
) -> Result<StreamingScalar, Failure> {
    use formula_codec::BinaryOperator;

    match operator {
        BinaryOperator::Add => {
            streaming_finite_number(streaming_scalar_number(left) + streaming_scalar_number(right))
        },
        BinaryOperator::Subtract => {
            streaming_finite_number(streaming_scalar_number(left) - streaming_scalar_number(right))
        },
        BinaryOperator::Multiply => {
            streaming_finite_number(streaming_scalar_number(left) * streaming_scalar_number(right))
        },
        BinaryOperator::Divide => {
            streaming_finite_number(streaming_scalar_number(left) / streaming_scalar_number(right))
        },
        BinaryOperator::Power => streaming_finite_number(
            streaming_scalar_number(left).powf(streaming_scalar_number(right)),
        ),
        BinaryOperator::GreaterThan => Ok(StreamingScalar::Boolean(
            streaming_scalar_number(left) > streaming_scalar_number(right),
        )),
        BinaryOperator::GreaterThanOrEqual => Ok(StreamingScalar::Boolean(
            streaming_scalar_number(left) >= streaming_scalar_number(right),
        )),
        BinaryOperator::LessThan => Ok(StreamingScalar::Boolean(
            streaming_scalar_number(left) < streaming_scalar_number(right),
        )),
        BinaryOperator::LessThanOrEqual => Ok(StreamingScalar::Boolean(
            streaming_scalar_number(left) <= streaming_scalar_number(right),
        )),
        BinaryOperator::Equal => Ok(StreamingScalar::Boolean(left == right)),
        BinaryOperator::NotEqual => Ok(StreamingScalar::Boolean(left != right)),
    }
}

fn formula_reports_usage(
    first: formula_codec::DecodeReport,
    second: formula_codec::DecodeReport,
) -> Result<FormulaEvaluationUsage, Failure> {
    let sum = |left: usize, right: usize| left.checked_add(right).ok_or(Failure::InvalidSource);
    Ok(FormulaEvaluationUsage {
        wire_bytes: sum(first.bytes(), second.bytes())?,
        wire_fields: sum(first.fields(), second.fields())?,
        wire_work: sum(first.work(), second.work())?,
        wire_text_bytes: sum(first.text_bytes(), second.text_bytes())?,
        formula_work: sum(first.node_count(), second.node_count())?
            .checked_add(sum(first.precedent_count(), second.precedent_count())?)
            .ok_or(Failure::InvalidSource)?,
        scratch_bytes: 0,
        allocations: sum(first.allocations(), second.allocations())?,
    })
}

/// Prove the exact second strict decode and fixed post-pass work fit before
/// any evaluator-owned allocation or callback executes.
fn preflight_formula_pass(
    inspection: formula_codec::DecodeReport,
    fixed_formula_work: usize,
    limits: CacheLimits,
) -> Result<usize, Failure> {
    let doubled = |value: usize| value.checked_mul(2).ok_or(Failure::InvalidSource);
    let tripled = |value: usize| value.checked_mul(3).ok_or(Failure::InvalidSource);
    check_limit(doubled(inspection.bytes())?, limits.wire_bytes)?;
    check_limit(tripled(inspection.fields())?, limits.wire_fields)?;
    check_limit(tripled(inspection.work())?, limits.wire_work)?;
    check_limit(tripled(inspection.text_bytes())?, limits.wire_text)?;
    let fixed = doubled(inspection.node_count())?
        .checked_add(doubled(inspection.precedent_count())?)
        .and_then(|work| work.checked_add(fixed_formula_work))
        .ok_or(Failure::InvalidSource)?;
    let maximum = limits.formula_work.min(limits.graph_work);
    check_limit(fixed, maximum)?;
    Ok(maximum - fixed)
}

/// Prove the second two-pass dependency projection and its fixed ordering
/// work fit before retaining any precedents.  Unlike evaluator inspection,
/// the sizing probe already includes both strict preflight and callback
/// passes, so its field/work/text totals are doubled rather than tripled.
fn preflight_formula_dependency_pass(
    inspection: formula_codec::DecodeReport,
    fixed_formula_work: usize,
    limits: CacheLimits,
) -> Result<(), Failure> {
    let doubled = |value: usize| value.checked_mul(2).ok_or(Failure::InvalidSource);
    check_limit(doubled(inspection.bytes())?, limits.wire_bytes)?;
    check_limit(doubled(inspection.fields())?, limits.wire_fields)?;
    check_limit(doubled(inspection.work())?, limits.wire_work)?;
    check_limit(doubled(inspection.text_bytes())?, limits.wire_text)?;
    let fixed = doubled(inspection.node_count())?
        .checked_add(doubled(inspection.precedent_count())?)
        .and_then(|work| work.checked_add(fixed_formula_work))
        .ok_or(Failure::InvalidSource)?;
    check_limit(fixed, limits.formula_work.min(limits.graph_work))
}

fn map_formula_decode_failure(error: formula_codec::DecodeError) -> Failure {
    if let Some(limit) = error.resource_limit() {
        let (observed, maximum) = match limit {
            formula_codec::DecodeLimit::Bytes { observed, maximum }
            | formula_codec::DecodeLimit::Fields { observed, maximum }
            | formula_codec::DecodeLimit::Work { observed, maximum }
            | formula_codec::DecodeLimit::Nodes { observed, maximum }
            | formula_codec::DecodeLimit::Text { observed, maximum } => (
                usize_to_u64_saturating(observed),
                usize_to_u64_saturating(maximum),
            ),
            formula_codec::DecodeLimit::Nesting { observed, maximum } => {
                (u64::from(observed), u64::from(maximum))
            },
        };
        return Failure::LimitExceeded { observed, maximum };
    }
    match error.invalid_reason() {
        Some(formula_codec::InvalidReason::ExternalOwner) => {
            Failure::UnsupportedDependency(Unsupported::ExternalOwner)
        },
        Some(formula_codec::InvalidReason::UnsupportedFormula) => {
            Failure::UnsupportedDependency(Unsupported::Formula)
        },
        _ => Failure::InvalidSource,
    }
}

fn unsupported_formula<T>() -> Result<T, Failure> {
    Err(Failure::UnsupportedDependency(Unsupported::Formula))
}

const FORMULA_LIST_TYPE: i32 = 3;

/// One borrowed segment selected by the root formula-list reference.
#[derive(Debug, Clone, Copy)]
pub(super) struct FormulaListSegment<'source> {
    pub(super) identifier: u64,
    pub(super) bytes: &'source [u8],
}

/// A formula entry retained without copying its native `FormulaArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct FormulaListEntry<'source> {
    pub(super) key: u32,
    pub(super) ref_count: u32,
    pub(super) bytes: &'source [u8],
}

/// Strictly collect one selected table's formula data list.
///
/// The root and every supplied segment are decoded with the generated-free
/// storage codec. Visitor effects remain staged until the whole enclosing
/// decode has passed Buffa parity. The result is sorted and key-unique, which
/// lets the cache commit layer perform one indexed join with formula cells.
pub(super) fn collect_formula_list<'source>(
    selected: TableIdentity,
    root: &'source [u8],
    segments: &[FormulaListSegment<'source>],
    expected_total_entries: usize,
    limits: CacheLimits,
) -> Result<(Vec<FormulaListEntry<'source>>, CacheUsage), Failure> {
    if selected.owner == 0 {
        return Err(Failure::InvalidSource);
    }
    check_limit(expected_total_entries, limits.graph_nodes)?;
    let mut usage = CacheUsage::default();
    let options = storage_decode_options(limits);
    let segment_index = formula_segment_index(segments, &mut usage, limits)?;
    let mut entries = Vec::new();
    retained_allocation::<FormulaListEntry<'source>>(&mut usage, expected_total_entries, limits)?;
    reserve(&mut entries, expected_total_entries)?;

    let mut root_stage = FormulaListStage::new(limits);
    let (root_snapshot, report) =
        storage::decode_table_data_list_with_visitor(root, options, &mut root_stage)
            .map_err(map_storage_decode_failure)?;
    root_stage.ensure_complete()?;
    charge_formula_list_stage(&mut usage, &root_stage, limits)?;
    charge_report(&mut usage, report, limits)?;
    if root_snapshot.list_type() != FORMULA_LIST_TYPE {
        return Err(Failure::UnsupportedDependency(Unsupported::CacheType));
    }
    verify_formula_segments(&mut root_stage.segments, &segment_index)?;
    let mut root_entries =
        decode_formula_entries(root, options, expected_total_entries, limits, &mut usage)?;
    append_formula_list_entries(
        &mut entries,
        &mut root_entries,
        expected_total_entries,
        limits,
        &mut usage,
    )?;

    for index in segment_index {
        graph_step(&mut usage, binary_search_work(segments.len()), limits)?;
        let segment = segments[index.index];
        let mut stage = FormulaListStage::new(limits);
        let (snapshot, report) = storage::decode_table_data_list_segment_with_visitor(
            segment.bytes,
            options,
            &mut stage,
        )
        .map_err(map_storage_decode_failure)?;
        stage.ensure_complete()?;
        charge_formula_list_stage(&mut usage, &stage, limits)?;
        charge_report(&mut usage, report, limits)?;
        if snapshot.list_type() != FORMULA_LIST_TYPE
            || segment_contains_nested_reference(segment.bytes)?
        {
            return Err(Failure::UnsupportedDependency(Unsupported::CacheType));
        }
        let mut segment_entries = decode_formula_entries(
            segment.bytes,
            options,
            expected_total_entries,
            limits,
            &mut usage,
        )?;
        validate_segment_keys(
            &segment_entries,
            snapshot.key_range_location(),
            snapshot.key_range_length(),
        )?;
        append_formula_list_entries(
            &mut entries,
            &mut segment_entries,
            expected_total_entries,
            limits,
            &mut usage,
        )?;
    }
    if entries.len() != expected_total_entries {
        return Err(Failure::InvalidSource);
    }
    graph_step(&mut usage, sort_work(entries.len())?, limits)?;
    entries.sort_unstable();
    graph_step(&mut usage, entries.len(), limits)?;
    if entries.windows(2).any(|pair| pair[0].key == pair[1].key) {
        return Err(Failure::InvalidSource);
    }
    Ok((entries, usage))
}

/// Prove that formula-list reference counts describe exactly the selected
/// dependency-host multiset and that every joined payload is byte-exact.
///
/// This check rejects orphan list entries, missing hosts, duplicate hosts,
/// wrong formula keys, and joins that substitute a different formula payload.
pub(super) fn validate_formula_coverage(
    selected: TableIdentity,
    entries: &[FormulaListEntry<'_>],
    payloads: &[FormulaPayload<'_>],
    limits: CacheLimits,
) -> Result<CacheUsage, Failure> {
    check_limit(entries.len(), limits.graph_nodes)?;
    check_limit(payloads.len(), limits.graph_nodes)?;
    if entries.windows(2).any(|pair| pair[0].key >= pair[1].key)
        || entries
            .iter()
            .any(|entry| entry.key == 0 || entry.ref_count == 0)
        || payloads
            .windows(2)
            .any(|pair| (pair[0].owner, pair[0].coordinate) >= (pair[1].owner, pair[1].coordinate))
        || payloads
            .iter()
            .any(|payload| payload.owner != selected.owner)
    {
        return Err(Failure::InvalidSource);
    }

    let mut usage = CacheUsage::default();
    let graph_work = payloads
        .len()
        .checked_mul(binary_search_work(entries.len()))
        .and_then(|work| work.checked_add(entries.len()))
        .ok_or(Failure::InvalidSource)?;
    graph_step(&mut usage, graph_work, limits)?;
    let mut counts = usize_slots(entries.len(), &mut usage, limits)?;
    for payload in payloads {
        add_usage(&mut usage.lookup_work, 1)?;
        let index = entries
            .binary_search_by_key(&payload.key, |entry| entry.key)
            .map_err(|_error| Failure::InvalidSource)?;
        if entries[index].bytes != payload.bytes {
            return Err(Failure::InvalidSource);
        }
        counts[index] = counts[index].checked_add(1).ok_or(Failure::InvalidSource)?;
    }
    for (entry, count) in entries.iter().zip(counts) {
        if usize::try_from(entry.ref_count).map_err(|_error| Failure::InvalidSource)? != count {
            return Err(Failure::InvalidSource);
        }
    }
    Ok(usage)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaListStageFailure {
    Invalid,
    Allocation(usize),
    Limit { observed: usize, maximum: usize },
}

struct FormulaListStage {
    segments: Vec<u64>,
    failure: Option<FormulaListStageFailure>,
    scratch_bytes: usize,
    allocations: usize,
    limits: CacheLimits,
}

impl FormulaListStage {
    fn new(limits: CacheLimits) -> Self {
        Self {
            segments: Vec::new(),
            failure: None,
            scratch_bytes: 0,
            allocations: 0,
            limits,
        }
    }

    fn push<T>(&mut self, slot: &mut Vec<T>, value: T, maximum: usize) {
        if self.failure.is_some() {
            return;
        }
        let observed = match slot.len().checked_add(1) {
            Some(value) => value,
            None => {
                self.failure = Some(FormulaListStageFailure::Limit {
                    observed: usize::MAX,
                    maximum,
                });
                return;
            },
        };
        if observed > maximum {
            self.failure = Some(FormulaListStageFailure::Limit { observed, maximum });
            return;
        }
        let scratch_bytes = match self.scratch_bytes.checked_add(size_of::<T>()) {
            Some(value) => value,
            None => {
                self.failure = Some(FormulaListStageFailure::Limit {
                    observed: usize::MAX,
                    maximum: self.limits.scratch_bytes,
                });
                return;
            },
        };
        if scratch_bytes > self.limits.scratch_bytes {
            self.failure = Some(FormulaListStageFailure::Limit {
                observed: scratch_bytes,
                maximum: self.limits.scratch_bytes,
            });
            return;
        }
        if slot.len() == slot.capacity() {
            let allocations = match self.allocations.checked_add(1) {
                Some(value) => value,
                None => {
                    self.failure = Some(FormulaListStageFailure::Limit {
                        observed: usize::MAX,
                        maximum: self.limits.allocations,
                    });
                    return;
                },
            };
            if allocations > self.limits.allocations {
                self.failure = Some(FormulaListStageFailure::Limit {
                    observed: allocations,
                    maximum: self.limits.allocations,
                });
                return;
            }
            self.allocations = allocations;
        }
        if slot.try_reserve_exact(1).is_err() {
            self.failure = Some(FormulaListStageFailure::Allocation(1));
            return;
        }
        self.scratch_bytes = scratch_bytes;
        slot.push(value);
    }

    fn push_segment(&mut self, identifier: u64) {
        let mut segments = core::mem::take(&mut self.segments);
        self.push(&mut segments, identifier, self.limits.wire_references);
        self.segments = segments;
    }

    fn ensure_complete(&self) -> Result<(), Failure> {
        match self.failure {
            Some(FormulaListStageFailure::Invalid) => Err(Failure::InvalidSource),
            Some(FormulaListStageFailure::Allocation(amount)) => {
                Err(Failure::Allocation { amount })
            },
            Some(FormulaListStageFailure::Limit { observed, maximum }) => {
                Err(Failure::LimitExceeded {
                    observed: usize_to_u64_saturating(observed),
                    maximum: usize_to_u64_saturating(maximum),
                })
            },
            None => Ok(()),
        }
    }
}

impl storage::StorageVisitor for FormulaListStage {
    fn visit_list_segment(
        &mut self,
        reference: storage::ReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        let reference = reference.reference();
        if reference.identifier() == 0 || reference.deprecated_is_external() == Some(true) {
            self.failure = Some(FormulaListStageFailure::Invalid);
            return Ok(());
        }
        self.push_segment(reference.identifier());
        Ok(())
    }
}

fn storage_decode_options(limits: CacheLimits) -> storage::DecodeOptions {
    storage::DecodeOptions::new(
        limits.wire_bytes,
        limits.wire_fields,
        limits.wire_work,
        limits.nesting,
        limits.wire_references,
        limits.wire_text,
    )
}

fn map_storage_decode_failure(error: storage::DecodeError) -> Failure {
    map_decode_failure(error)
}

fn decode_formula_entries<'source>(
    source: &'source [u8],
    options: storage::DecodeOptions,
    maximum_entries: usize,
    limits: CacheLimits,
    usage: &mut CacheUsage,
) -> Result<Vec<FormulaListEntry<'source>>, Failure> {
    let view = WireView::parse(source).map_err(|_error| Failure::InvalidSource)?;
    let mut output = Vec::new();
    for field in view.fields() {
        if field.number() != 3 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Failure::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_error| Failure::InvalidSource)?;
        let observed = output.len().checked_add(1).ok_or(Failure::InvalidSource)?;
        if observed > maximum_entries {
            return Err(Failure::LimitExceeded {
                observed: usize_to_u64_saturating(observed),
                maximum: usize_to_u64_saturating(maximum_entries),
            });
        }
        let (entry, report) =
            storage::decode_table_data_list_entry_with_report(field.payload(), options)
                .map_err(map_storage_decode_failure)?;
        charge_report(usage, report, limits)?;
        let bytes = entry
            .formula()
            .ok_or(Failure::UnsupportedDependency(Unsupported::CacheType))?;
        if entry.key() == 0
            || entry.ref_count() == 0
            || entry.string_value().is_some()
            || entry.reference().is_some()
            || entry.format().is_some()
            || entry.custom_format().is_some()
            || entry.rich_text_payload().is_some()
            || entry.comment_storage().is_some()
            || entry.import_warning_set().is_some()
            || entry.cell_spec().is_some()
        {
            return Err(Failure::UnsupportedDependency(Unsupported::CacheType));
        }
        scratch_allocation::<FormulaListEntry<'source>>(usage, 1, limits)?;
        reserve(&mut output, 1)?;
        graph_step(usage, 1, limits)?;
        output.push(FormulaListEntry {
            key: entry.key(),
            ref_count: entry.ref_count(),
            bytes,
        });
    }
    Ok(output)
}

fn formula_segment_index(
    segments: &[FormulaListSegment<'_>],
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Vec<PayloadIndex>, Failure> {
    let mut index = Vec::new();
    scratch_allocation::<PayloadIndex>(usage, segments.len(), limits)?;
    reserve(&mut index, segments.len())?;
    for (position, segment) in segments.iter().enumerate() {
        if segment.identifier == 0 {
            return Err(Failure::InvalidSource);
        }
        graph_step(usage, 1, limits)?;
        index.push(PayloadIndex {
            identifier: segment.identifier,
            index: position,
        });
    }
    graph_step(usage, sort_work(index.len())?, limits)?;
    index.sort_unstable();
    if index
        .windows(2)
        .any(|pair| pair[0].identifier == pair[1].identifier)
    {
        return Err(Failure::InvalidSource);
    }
    Ok(index)
}

fn verify_formula_segments(
    references: &mut [u64],
    segments: &[PayloadIndex],
) -> Result<(), Failure> {
    if references.len() != segments.len() {
        return Err(Failure::InvalidSource);
    }
    references.sort_unstable();
    if references.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Failure::InvalidSource);
    }
    for (reference, segment) in references.iter().zip(segments) {
        if *reference != segment.identifier {
            return Err(Failure::InvalidSource);
        }
    }
    Ok(())
}

fn segment_contains_nested_reference(source: &[u8]) -> Result<bool, Failure> {
    let view = WireView::parse(source).map_err(|_error| Failure::InvalidSource)?;
    for field in view.fields() {
        if field.number() == 4 {
            field
                .validate_canonical_framing()
                .map_err(|_error| Failure::InvalidSource)?;
            return Ok(true);
        }
    }
    Ok(false)
}

fn validate_segment_keys(
    entries: &[FormulaListEntry<'_>],
    location: u32,
    length: u32,
) -> Result<(), Failure> {
    let end = location.checked_add(length).ok_or(Failure::InvalidSource)?;
    if length == 0
        || entries
            .iter()
            .any(|entry| entry.key < location || entry.key >= end)
    {
        return Err(Failure::InvalidSource);
    }
    Ok(())
}

fn append_formula_list_entries<'source>(
    output: &mut Vec<FormulaListEntry<'source>>,
    staged: &mut Vec<FormulaListEntry<'source>>,
    expected: usize,
    limits: CacheLimits,
    usage: &mut CacheUsage,
) -> Result<(), Failure> {
    let total = output
        .len()
        .checked_add(staged.len())
        .ok_or(Failure::InvalidSource)?;
    if total > expected {
        return Err(Failure::InvalidSource);
    }
    graph_step(usage, staged.len(), limits)?;
    output.append(staged);
    Ok(())
}

fn charge_formula_list_stage(
    usage: &mut CacheUsage,
    stage: &FormulaListStage,
    limits: CacheLimits,
) -> Result<(), Failure> {
    charge(
        &mut usage.peak_scratch_bytes,
        stage.scratch_bytes,
        limits.scratch_bytes,
    )?;
    charge(
        &mut usage.allocations,
        stage.allocations,
        limits.allocations,
    )
}

/// Decode, sort, and deduplicate inline and tiled formula hosts for one owner.
///
/// This is the narrow integration seam for the storage layer. It keeps all
/// dependency projection details here while leaving the byte-exact formula
/// table-data-list lookup, payload join, and cache-object routing to the
/// caller. The returned hosts are sorted uniquely by `(owner, row, column)`.
pub(super) fn collect_formula_hosts(
    selected: TableIdentity,
    owners: &[DependencyPayload<'_>],
    record_tiles: &[DependencyPayload<'_>],
    limits: CacheLimits,
) -> Result<(Vec<FormulaHost>, CacheUsage), Failure> {
    let mut usage = CacheUsage {
        formula_graph_builds: 1,
        ..CacheUsage::default()
    };
    let owner_index = payload_index(owners, &mut usage, limits)?;
    let tile_index = payload_index(record_tiles, &mut usage, limits)?;
    let options = decode_options(limits);
    let mut candidates = Vec::new();
    let mut selected_owner_seen = false;

    for entry in &owner_index {
        graph_step(&mut usage, 1, limits)?;
        add_usage(&mut usage.lookup_work, 1)?;
        let payload = owners[entry.index];
        let mut stage = Stage::new(limits);
        let (owner, report) = dependency::decode_formula_owner_dependencies_with_visitor(
            payload.bytes,
            options,
            &mut stage,
        )
        .map_err(map_decode_failure)?;
        stage.ensure_complete()?;
        charge_stage_usage(&mut usage, &stage, limits)?;
        charge_report(&mut usage, report, limits)?;
        if owner.internal_formula_owner_id() != selected.owner {
            continue;
        }
        if selected_owner_seen
            || owner.formula_owner_uid().lower() != selected.uuid_lower
            || owner.formula_owner_uid().upper() != selected.uuid_upper
        {
            return Err(Failure::InvalidSource);
        }
        selected_owner_seen = true;
        append_formula_hosts(
            &mut candidates,
            selected.owner,
            &stage.records,
            limits,
            &mut usage,
        )?;

        for tile_reference in stage.tile_references.iter().copied() {
            let tile_payload = indexed_payload(
                record_tiles,
                &tile_index,
                tile_reference,
                &mut usage,
                limits,
            )?;
            let mut tile_stage = Stage::new(limits);
            let (tile, report) = dependency::decode_cell_record_tile_with_visitor(
                tile_payload.bytes,
                options,
                &mut tile_stage,
            )
            .map_err(map_decode_failure)?;
            tile_stage.ensure_complete()?;
            charge_stage_usage(&mut usage, &tile_stage, limits)?;
            charge_report(&mut usage, report, limits)?;
            if tile.internal_owner_id() != selected.owner {
                return Err(Failure::InvalidSource);
            }
            append_formula_hosts(
                &mut candidates,
                selected.owner,
                &tile_stage.records,
                limits,
                &mut usage,
            )?;
        }
    }
    if !selected_owner_seen {
        return Err(Failure::UnsupportedDependency(Unsupported::MissingOwner));
    }

    graph_step(&mut usage, sort_work(candidates.len())?, limits)?;
    candidates.sort_unstable();
    graph_step(&mut usage, candidates.len(), limits)?;
    candidates.dedup();
    check_limit(candidates.len(), limits.graph_nodes)?;
    add_usage(&mut usage.formula_nodes, candidates.len())?;

    let mut hosts = Vec::new();
    graph_step(&mut usage, candidates.len(), limits)?;
    retained_allocation::<FormulaHost>(&mut usage, candidates.len(), limits)?;
    reserve(&mut hosts, candidates.len())?;
    for host in candidates {
        hosts.push(host);
    }
    Ok((hosts, usage))
}

/// Construct one final-state graph and grouped cache rewrite plan.
///
/// `overlay` must be sorted and unique by coordinate; callers normally obtain
/// that compact buffer from the edit planner.  Formula cells addressed by the
/// overlay are deleted from the old dependency graph before propagation, so a
/// final clear/scalar write cannot refresh an obsolete formula cache.
pub(super) fn plan_final_cache(
    source: CacheSource<'_>,
    overlay: &[FinalCell],
    limits: CacheLimits,
    baseline: &dyn CacheBaseline,
    evaluator: &mut dyn CacheEvaluator,
) -> Result<CachePlan, Failure> {
    ensure_overlay(source, overlay)?;
    let mut usage = CacheUsage {
        formula_graph_builds: 1,
        ..CacheUsage::default()
    };
    let owner_index = payload_index(source.owners, &mut usage, limits)?;
    let tile_index = payload_index(source.record_tiles, &mut usage, limits)?;
    let formula_index = formula_index(source, &mut usage, limits)?;
    let coordinate_index = formula_coordinate_index(&formula_index, &mut usage, limits)?;
    check_limit(formula_index.len(), limits.graph_nodes)?;
    add_usage(&mut usage.formula_nodes, formula_index.len())?;

    let options = decode_options(limits);
    let mut engine_stage = Stage::new(limits);
    let (engine, report) = dependency::decode_calculation_engine_with_visitor(
        source.engine,
        options,
        &mut engine_stage,
    )
    .map_err(map_decode_failure)?;
    engine_stage.ensure_complete()?;
    charge_stage_usage(&mut usage, &engine_stage, limits)?;
    charge_report(&mut usage, report, limits)?;
    if engine.header_name_manager().is_some()
        && overlay.iter().any(|cell| {
            cell.coordinate.row < source.header_rows
                || cell.coordinate.column < source.header_columns
        })
    {
        return Err(Failure::UnsupportedDependency(
            Unsupported::HeaderNameManager,
        ));
    }
    verify_owner_references(&mut engine_stage.owner_references, &owner_index)?;

    let mut edges = Vec::new();
    let mut dependency_hosts = Vec::new();
    let mut selected_owner_seen = false;
    for owner_reference in engine_stage.owner_references.iter().copied() {
        let payload = indexed_payload(
            source.owners,
            &owner_index,
            owner_reference,
            &mut usage,
            limits,
        )?;
        let mut stage = Stage::new(limits);
        let (owner, report) = dependency::decode_formula_owner_dependencies_with_visitor(
            payload.bytes,
            options,
            &mut stage,
        )
        .map_err(map_decode_failure)?;
        stage.ensure_complete()?;
        charge_stage_usage(&mut usage, &stage, limits)?;
        charge_report(&mut usage, report, limits)?;
        let selected_owner = owner.internal_formula_owner_id() == source.selected_table.owner;
        if selected_owner {
            if selected_owner_seen
                || owner.formula_owner_uid().lower() != source.selected_table.uuid_lower
                || owner.formula_owner_uid().upper() != source.selected_table.uuid_upper
            {
                return Err(Failure::InvalidSource);
            }
            selected_owner_seen = true;
            reject_unsupported_owner(owner, &stage, source)?;
        }
        if !selected_owner {
            continue;
        }
        append_edges(
            &mut edges,
            owner.internal_formula_owner_id(),
            &stage.records,
            limits,
            &mut usage,
        )?;
        append_formula_hosts(
            &mut dependency_hosts,
            owner.internal_formula_owner_id(),
            &stage.records,
            limits,
            &mut usage,
        )?;

        for tile_reference in stage.tile_references.iter().copied() {
            let tile_payload = indexed_payload(
                source.record_tiles,
                &tile_index,
                tile_reference,
                &mut usage,
                limits,
            )?;
            let mut tile_stage = Stage::new(limits);
            let (tile, report) = dependency::decode_cell_record_tile_with_visitor(
                tile_payload.bytes,
                options,
                &mut tile_stage,
            )
            .map_err(map_decode_failure)?;
            tile_stage.ensure_complete()?;
            charge_stage_usage(&mut usage, &tile_stage, limits)?;
            charge_report(&mut usage, report, limits)?;
            if tile.internal_owner_id() != owner.internal_formula_owner_id() {
                return Err(Failure::InvalidSource);
            }
            append_edges(
                &mut edges,
                tile.internal_owner_id(),
                &tile_stage.records,
                limits,
                &mut usage,
            )?;
            append_formula_hosts(
                &mut dependency_hosts,
                tile.internal_owner_id(),
                &tile_stage.records,
                limits,
                &mut usage,
            )?;
        }
    }
    if (!overlay.is_empty() || !source.formulas.is_empty()) && !selected_owner_seen {
        return Err(Failure::UnsupportedDependency(Unsupported::MissingOwner));
    }
    graph_step(&mut usage, sort_work(edges.len())?, limits)?;
    edges.sort_unstable();
    edges.dedup();
    validate_dependency_hosts(&mut dependency_hosts, &formula_index, limits, &mut usage)?;
    validate_formula_edges(source.formulas, &edges, evaluator, limits, &mut usage)?;

    let removals = changed_formula_removals(
        overlay,
        source.formulas,
        &formula_index,
        &coordinate_index,
        &mut usage,
        limits,
    )?;
    let mut computed = option_slots(source.formulas.len(), &mut usage, limits)?;
    let mut dirty = bool_slots(source.formulas.len(), &mut usage, limits)?;
    // First build the complete dirty closure.  Evaluating while this walk is
    // still discovering descendants can observe a stale predecessor in a
    // fan-in chain.
    let mut closure = Vec::new();
    let closure_capacity = overlay
        .len()
        .checked_add(source.formulas.len())
        .ok_or(Failure::InvalidSource)?;
    scratch_allocation::<FormulaKey>(&mut usage, closure_capacity, limits)?;
    reserve(&mut closure, closure_capacity)?;
    for cell in overlay {
        closure.push(FormulaKey {
            owner: cell.table.owner,
            coordinate: cell.coordinate,
        });
    }
    let mut cursor = 0usize;
    while cursor < closure.len() {
        let changed = closure[cursor];
        cursor = cursor.checked_add(1).ok_or(Failure::InvalidSource)?;
        graph_step(&mut usage, 1, limits)?;
        add_usage(&mut usage.queue_work, 1)?;
        let (start, end) = edge_range(&edges, changed, &mut usage, limits)?;
        add_usage(&mut usage.dependency_range_queries, 1)?;
        for edge in &edges[start..end] {
            graph_step(&mut usage, 1, limits)?;
            add_usage(&mut usage.dependency_range_candidates, 1)?;
            let formula_position = find_formula(&formula_index, edge.target, &mut usage, limits)?
                .ok_or(Failure::UnsupportedDependency(Unsupported::Formula))?;
            if edge.target_is_in_cycle {
                return Err(Failure::UnsupportedDependency(Unsupported::Formula));
            }
            if overlay_has(overlay, edge.target, &mut usage, limits)? || dirty[formula_position] {
                continue;
            }
            dirty[formula_position] = true;
            closure.push(edge.target);
        }
    }

    // Kahn ordering is required because a final overlay can fan out into a
    // chain of formula caches.  An edge is predecessor -> dependent host.
    let mut indegree = usize_slots(source.formulas.len(), &mut usage, limits)?;
    let mut dirty_count = 0usize;
    for is_dirty in dirty.iter().copied() {
        if is_dirty {
            dirty_count = dirty_count.checked_add(1).ok_or(Failure::InvalidSource)?;
        }
    }
    for edge in &edges {
        graph_step(&mut usage, 1, limits)?;
        let source_is_dirty_formula = find_coordinate_formula(
            &coordinate_index,
            edge.source.coordinate,
            &mut usage,
            limits,
        )?
        .is_some_and(|position| dirty[position]);
        if !overlay_has(overlay, edge.source, &mut usage, limits)? && !source_is_dirty_formula {
            continue;
        }
        let target = find_formula(&formula_index, edge.target, &mut usage, limits)?
            .ok_or(Failure::UnsupportedDependency(Unsupported::Formula))?;
        if !dirty[target] {
            continue;
        }
        if let Some(precedent) = find_coordinate_formula(
            &coordinate_index,
            edge.source.coordinate,
            &mut usage,
            limits,
        )? {
            if dirty[precedent] {
                indegree[target] = indegree[target]
                    .checked_add(1)
                    .ok_or(Failure::InvalidSource)?;
            }
        }
    }
    let mut ready = Vec::new();
    scratch_allocation::<usize>(&mut usage, dirty_count, limits)?;
    reserve(&mut ready, dirty_count)?;
    for (position, is_dirty) in dirty.iter().copied().enumerate() {
        if is_dirty && indegree[position] == 0 {
            ready.push(position);
        }
    }
    cursor = 0;
    while cursor < ready.len() {
        let formula_position = ready[cursor];
        cursor = cursor.checked_add(1).ok_or(Failure::InvalidSource)?;
        graph_step(&mut usage, 1, limits)?;
        add_usage(&mut usage.queue_work, 1)?;
        let values = FinalValues {
            overlay,
            coordinate_index: &coordinate_index,
            computed: &computed,
            baseline,
        };
        let evaluator_limits = remaining_limits(limits, &usage)?;
        let evaluation = evaluator.evaluate(
            source.formulas[formula_index[formula_position].index],
            &values,
            evaluator_limits,
        )?;
        if !matches!(
            evaluation.value,
            FormulaCachedValue::Number(_) | FormulaCachedValue::Boolean(_)
        ) {
            return Err(Failure::UnsupportedDependency(Unsupported::CacheType));
        }
        charge_formula_evaluation_usage(&mut usage, evaluation.usage, limits)?;
        computed[formula_position] = Some(evaluation.value);
        let source_formula = source.formulas[formula_index[formula_position].index];
        let source_key = FormulaKey {
            owner: source_formula.owner,
            coordinate: source_formula.coordinate,
        };
        let (start, end) = edge_range(&edges, source_key, &mut usage, limits)?;
        add_usage(&mut usage.dependency_range_queries, 1)?;
        for edge in &edges[start..end] {
            graph_step(&mut usage, 1, limits)?;
            add_usage(&mut usage.dependency_range_candidates, 1)?;
            let target = find_formula(&formula_index, edge.target, &mut usage, limits)?
                .ok_or(Failure::UnsupportedDependency(Unsupported::Formula))?;
            if !dirty[target] {
                continue;
            }
            indegree[target] = indegree[target]
                .checked_sub(1)
                .ok_or(Failure::InvalidSource)?;
            if indegree[target] == 0 {
                ready.push(target);
            }
        }
    }
    if cursor != dirty_count {
        return Err(Failure::UnsupportedDependency(Unsupported::Formula));
    }

    let rewrites = grouped_rewrites(
        source.formulas,
        &formula_index,
        &mut computed,
        limits,
        &mut usage,
    )?;
    Ok(CachePlan {
        rewrites,
        removals,
        usage,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct FormulaKey {
    owner: u32,
    coordinate: Coordinate,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Edge {
    source: FormulaKey,
    target: FormulaKey,
    target_is_in_cycle: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct FormulaEdge {
    source: FormulaKey,
    target: FormulaKey,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct PayloadIndex {
    identifier: u64,
    index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct FormulaIndex {
    key: FormulaKey,
    index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct CoordinateIndex {
    coordinate: Coordinate,
    formula_position: usize,
}

struct Stage {
    owner_references: Vec<u64>,
    tile_references: Vec<u64>,
    range_references: Vec<u64>,
    components: Vec<dependency::ExpandedEdgeComponent>,
    records: Vec<CapturedRecord>,
    allocation: Option<usize>,
    limit: Option<(usize, usize)>,
    retained_bytes: usize,
    allocations: usize,
    limits: CacheLimits,
}

#[derive(Debug)]
struct CapturedRecord {
    coordinate: Coordinate,
    edges: Option<dependency::ExpandedEdgesSnapshot>,
    is_in_cycle: bool,
    components: Vec<dependency::ExpandedEdgeComponent>,
}

impl Stage {
    fn new(limits: CacheLimits) -> Self {
        Self {
            owner_references: Vec::new(),
            tile_references: Vec::new(),
            range_references: Vec::new(),
            components: Vec::new(),
            records: Vec::new(),
            allocation: None,
            limit: None,
            retained_bytes: 0,
            allocations: 0,
            limits,
        }
    }

    #[allow(clippy::too_many_arguments, reason = "streaming stage budget state")]
    fn push<T>(
        slot: &mut Vec<T>,
        value: T,
        maximum: usize,
        allocation: &mut Option<usize>,
        limit: &mut Option<(usize, usize)>,
        retained_bytes: &mut usize,
        allocations: &mut usize,
        limits: CacheLimits,
    ) {
        if allocation.is_some() || limit.is_some() {
            return;
        }
        let length = slot.len();
        let capacity = slot.capacity();
        let observed = match length.checked_add(1) {
            Some(value) => value,
            None => {
                *limit = Some((usize::MAX, maximum));
                return;
            },
        };
        if observed > maximum {
            *limit = Some((observed, maximum));
            return;
        }
        let bytes = match retained_bytes.checked_add(size_of::<T>()) {
            Some(value) => value,
            None => {
                *limit = Some((usize::MAX, limits.scratch_bytes));
                return;
            },
        };
        if bytes > limits.scratch_bytes {
            *limit = Some((bytes, limits.scratch_bytes));
            return;
        }
        if length == capacity {
            let next_allocations = match allocations.checked_add(1) {
                Some(value) => value,
                None => {
                    *limit = Some((usize::MAX, limits.allocations));
                    return;
                },
            };
            if next_allocations > limits.allocations {
                *limit = Some((next_allocations, limits.allocations));
                return;
            }
            *allocations = next_allocations;
        }
        if slot.try_reserve_exact(1).is_err() {
            *allocation = Some(1);
            return;
        }
        *retained_bytes = bytes;
        slot.push(value);
    }

    fn ensure_complete(&self) -> Result<(), Failure> {
        match self.allocation {
            Some(amount) => Err(Failure::Allocation { amount }),
            None if let Some((observed, maximum)) = self.limit => Err(Failure::LimitExceeded {
                observed: usize_to_u64_saturating(observed),
                maximum: usize_to_u64_saturating(maximum),
            }),
            None if !self.components.is_empty() => Err(Failure::InvalidSource),
            None => Ok(()),
        }
    }
}

impl dependency::DependencyVisitor for Stage {
    fn visit_formula_owner_dependency(
        &mut self,
        reference: dependency::ReferenceRecord<'_>,
    ) -> Result<(), dependency::DecodeError> {
        Self::push(
            &mut self.owner_references,
            reference.reference().identifier(),
            self.limits.wire_references,
            &mut self.allocation,
            &mut self.limit,
            &mut self.retained_bytes,
            &mut self.allocations,
            self.limits,
        );
        Ok(())
    }

    fn visit_tiled_cell_dependency(
        &mut self,
        reference: dependency::ReferenceRecord<'_>,
    ) -> Result<(), dependency::DecodeError> {
        Self::push(
            &mut self.tile_references,
            reference.reference().identifier(),
            self.limits.wire_references,
            &mut self.allocation,
            &mut self.limit,
            &mut self.retained_bytes,
            &mut self.allocations,
            self.limits,
        );
        Ok(())
    }

    fn visit_tiled_range_dependency(
        &mut self,
        reference: dependency::ReferenceRecord<'_>,
    ) -> Result<(), dependency::DecodeError> {
        Self::push(
            &mut self.range_references,
            reference.reference().identifier(),
            self.limits.wire_references,
            &mut self.allocation,
            &mut self.limit,
            &mut self.retained_bytes,
            &mut self.allocations,
            self.limits,
        );
        Ok(())
    }

    fn visit_expanded_edge_component(
        &mut self,
        component: dependency::ExpandedEdgeComponent,
    ) -> Result<(), dependency::DecodeError> {
        Self::push(
            &mut self.components,
            component,
            self.limits.graph_edges.saturating_mul(5),
            &mut self.allocation,
            &mut self.limit,
            &mut self.retained_bytes,
            &mut self.allocations,
            self.limits,
        );
        Ok(())
    }

    fn visit_cell_record(
        &mut self,
        record: dependency::CellRecordSnapshot<'_>,
    ) -> Result<(), dependency::DecodeError> {
        let captured = CapturedRecord {
            coordinate: Coordinate {
                row: record.row(),
                column: record.column(),
            },
            edges: record.expanded_edges_snapshot(),
            is_in_cycle: record.is_in_a_cycle().unwrap_or(false),
            components: core::mem::take(&mut self.components),
        };
        Self::push(
            &mut self.records,
            captured,
            self.limits.graph_nodes,
            &mut self.allocation,
            &mut self.limit,
            &mut self.retained_bytes,
            &mut self.allocations,
            self.limits,
        );
        Ok(())
    }
}

fn decode_options(limits: CacheLimits) -> dependency::DecodeOptions {
    dependency::DecodeOptions::new(
        limits.wire_bytes,
        limits.wire_fields,
        limits.wire_work,
        limits.nesting,
        limits.wire_references,
        limits.wire_text,
    )
}

fn map_decode_failure(error: dependency::DecodeError) -> Failure {
    let Some(limit) = error.resource_limit() else {
        return Failure::InvalidSource;
    };
    match limit {
        dependency::DecodeLimit::Bytes { observed, maximum }
        | dependency::DecodeLimit::References { observed, maximum }
        | dependency::DecodeLimit::Text { observed, maximum }
        | dependency::DecodeLimit::Fields { observed, maximum }
        | dependency::DecodeLimit::Work { observed, maximum } => Failure::LimitExceeded {
            observed: usize_to_u64_saturating(observed),
            maximum: usize_to_u64_saturating(maximum),
        },
        dependency::DecodeLimit::Nesting { observed, maximum } => Failure::LimitExceeded {
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        _ => Failure::InvalidSource,
    }
}

fn usize_to_u64_saturating(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn reject_unsupported_owner(
    owner: dependency::FormulaOwnerDependenciesSnapshot<'_>,
    stage: &Stage,
    source: CacheSource<'_>,
) -> Result<(), Failure> {
    let inert_graph = source.formulas.is_empty()
        && stage.owner_references.is_empty()
        && stage.tile_references.is_empty()
        && stage.range_references.is_empty()
        && stage.records.is_empty()
        && stage.components.is_empty();
    if owner
        .range_dependencies()
        .is_some_and(|value| !empty_message(value))
        || !stage.range_references.is_empty()
    {
        return Err(Failure::UnsupportedDependency(Unsupported::Range));
    }
    if owner
        .volatile_dependencies()
        .is_some_and(|value| !empty_volatile_dependencies(value))
    {
        return Err(Failure::UnsupportedDependency(Unsupported::Volatile));
    }
    if owner.spanning_column_dependencies().is_some_and(|value| {
        !(selected_spanning_dependencies(value, source) || inert_graph && empty_message(value))
    }) || owner.spanning_row_dependencies().is_some_and(|value| {
        !(selected_spanning_dependencies(value, source) || inert_graph && empty_message(value))
    }) {
        return Err(Failure::UnsupportedDependency(Unsupported::Spanning));
    }
    if owner
        .whole_owner_dependencies()
        .is_some_and(|value| !empty_whole_owner_dependencies(value))
    {
        return Err(Failure::UnsupportedDependency(Unsupported::WholeOwner));
    }
    if owner
        .uuid_references()
        .is_some_and(|value| !empty_message(value))
    {
        return Err(Failure::UnsupportedDependency(Unsupported::UuidReference));
    }
    if owner
        .spill_range_sizes()
        .is_some_and(|value| !inactive_spill_sizes(value))
    {
        return Err(Failure::UnsupportedDependency(Unsupported::Spill));
    }
    Ok(())
}

fn empty_message(payload: &[u8]) -> bool {
    WireView::parse(payload).is_ok_and(|view| view.fields().next().is_none())
}

fn empty_volatile_dependencies(payload: &[u8]) -> bool {
    let Ok(view) = WireView::parse(payload) else {
        return false;
    };
    let mut seen = [false; 7];
    for field in view.fields() {
        let Ok(index) = usize::try_from(field.number()) else {
            return false;
        };
        if !matches!(index, 1..=5 | 7)
            || seen[index - 1]
            || field.wire_type() != 2
            || field.validate_canonical_framing().is_err()
            || !empty_message(field.payload())
        {
            return false;
        }
        seen[index - 1] = true;
    }
    true
}

fn empty_whole_owner_dependencies(payload: &[u8]) -> bool {
    let Ok(view) = WireView::parse(payload) else {
        return false;
    };
    let mut fields = view.fields();
    let Some(field) = fields.next() else {
        return true;
    };
    field.number() == 1
        && field.wire_type() == 2
        && field.validate_canonical_framing().is_ok()
        && empty_message(field.payload())
        && fields.next().is_none()
}

fn selected_spanning_dependencies(payload: &[u8], source: CacheSource<'_>) -> bool {
    let Some(last_column) = source.columns.checked_sub(1) else {
        return false;
    };
    let Some(last_row) = source.rows.checked_sub(1) else {
        return false;
    };
    let Some(body_last_row) = source
        .rows
        .checked_sub(source.footer_rows)
        .and_then(|rows| rows.checked_sub(1))
    else {
        return false;
    };
    let Ok(view) = WireView::parse(payload) else {
        return false;
    };
    let mut total = false;
    let mut body = false;
    for field in view.fields() {
        if field.wire_type() != 2 || field.validate_canonical_framing().is_err() {
            return false;
        }
        let expected = match field.number() {
            2 if !total => {
                total = true;
                [0, 0, u64::from(last_column), u64::from(last_row)]
            },
            3 if !body => {
                body = true;
                [
                    u64::from(source.header_columns),
                    u64::from(source.header_rows),
                    u64::from(last_column),
                    u64::from(body_last_row),
                ]
            },
            _ => return false,
        };
        if parse_varint_fields(field.payload(), [1, 2, 3, 4]) != Some(expected) {
            return false;
        }
    }
    total && body
}

fn inactive_spill_sizes(payload: &[u8]) -> bool {
    let Ok(view) = WireView::parse(payload) else {
        return false;
    };
    view.fields().next().is_none()
}

fn parse_varint_fields<const N: usize>(source: &[u8], expected: [u32; N]) -> Option<[u64; N]> {
    let view = WireView::parse(source).ok()?;
    let mut fields = view.fields();
    let mut values = [0u64; N];
    for (index, number) in expected.into_iter().enumerate() {
        let field = fields.next()?;
        if field.number() != number
            || field.wire_type() != 0
            || field.validate_canonical_key().is_err()
        {
            return None;
        }
        let (value, consumed) = varint::decode_varint_from_bytes(field.payload()).ok()?;
        if consumed != field.payload().len() || varint::encoded_len(value) != consumed {
            return None;
        }
        values[index] = value;
    }
    fields.next().is_none().then_some(values)
}

fn append_formula_hosts(
    output: &mut Vec<FormulaHost>,
    owner: u32,
    records: &[CapturedRecord],
    limits: CacheLimits,
    usage: &mut CacheUsage,
) -> Result<(), Failure> {
    for record in records {
        let observed = output.len().checked_add(1).ok_or(Failure::InvalidSource)?;
        check_limit(observed, limits.graph_nodes)?;
        graph_step(usage, 1, limits)?;
        scratch_allocation::<FormulaHost>(usage, 1, limits)?;
        reserve(output, 1)?;
        output.push(FormulaHost {
            owner,
            coordinate: record.coordinate,
        });
    }
    Ok(())
}

fn validate_dependency_hosts(
    hosts: &mut Vec<FormulaHost>,
    formulas: &[FormulaIndex],
    limits: CacheLimits,
    usage: &mut CacheUsage,
) -> Result<(), Failure> {
    graph_step(usage, sort_work(hosts.len())?, limits)?;
    hosts.sort_unstable();
    graph_step(usage, hosts.len(), limits)?;
    hosts.dedup();
    if hosts.len() != formulas.len()
        || hosts.iter().zip(formulas).any(|(host, formula)| {
            host.owner != formula.key.owner || host.coordinate != formula.key.coordinate
        })
    {
        return Err(Failure::UnsupportedDependency(Unsupported::Formula));
    }
    Ok(())
}

fn append_edges(
    output: &mut Vec<Edge>,
    owner: u32,
    records: &[CapturedRecord],
    limits: CacheLimits,
    usage: &mut CacheUsage,
) -> Result<(), Failure> {
    for record in records {
        let Some(shape) = record.edges else {
            if !record.components.is_empty() {
                return Err(Failure::InvalidSource);
            }
            continue;
        };
        let local = shape.local();
        let external = shape.external();
        if external != 0 {
            // The expanded record names an owner for these precedents. This
            // table-local transaction cannot safely conflate it with a local
            // coordinate until the raw writer has an owner-aware overlay.
            return Err(Failure::UnsupportedDependency(Unsupported::ExternalOwner));
        }
        let expected = local.checked_mul(2).ok_or(Failure::InvalidSource)?;
        if record.components.len() != expected {
            return Err(Failure::InvalidSource);
        }
        let mut local_rows = Vec::new();
        let mut local_columns = Vec::new();
        scratch_allocation::<u32>(usage, local, limits)?;
        scratch_allocation::<u32>(usage, local, limits)?;
        reserve(&mut local_rows, local)?;
        reserve(&mut local_columns, local)?;
        for component in &record.components {
            match component.kind() {
                dependency::ExpandedEdgeKind::LocalRow => local_rows.push(component.value()),
                dependency::ExpandedEdgeKind::LocalColumn => local_columns.push(component.value()),
                dependency::ExpandedEdgeKind::ExternalRow
                | dependency::ExpandedEdgeKind::ExternalColumn
                | dependency::ExpandedEdgeKind::InternalOwner => {
                    return Err(Failure::InvalidSource);
                },
            }
        }
        if local_rows.len() != local || local_columns.len() != local {
            return Err(Failure::InvalidSource);
        }
        let total = local;
        check_limit(
            output
                .len()
                .checked_add(total)
                .ok_or(Failure::InvalidSource)?,
            limits.graph_edges,
        )?;
        scratch_allocation::<Edge>(usage, total, limits)?;
        reserve(output, total)?;
        for index in 0..local {
            output.push(Edge {
                source: FormulaKey {
                    owner,
                    coordinate: Coordinate {
                        row: local_rows[index],
                        column: local_columns[index],
                    },
                },
                target: FormulaKey {
                    owner,
                    coordinate: record.coordinate,
                },
                target_is_in_cycle: record.is_in_cycle,
            });
        }
        add_usage(&mut usage.dependency_edges, total)?;
    }
    Ok(())
}

fn validate_formula_edges(
    formulas: &[FormulaCell],
    dependency_edges: &[Edge],
    evaluator: &mut dyn CacheEvaluator,
    limits: CacheLimits,
    usage: &mut CacheUsage,
) -> Result<(), Failure> {
    if dependency_edges.iter().any(|edge| edge.target_is_in_cycle) {
        return Err(Failure::UnsupportedDependency(Unsupported::Formula));
    }
    let mut formula_edges = Vec::new();
    for formula in formulas.iter().copied() {
        let evaluator_limits = remaining_limits(limits, usage)?;
        let analysis = evaluator.analyze(formula, evaluator_limits)?;
        charge_formula_evaluation_usage(usage, analysis.usage, limits)?;
        for precedent in analysis.precedents {
            if precedent.owner != formula.owner {
                return Err(Failure::UnsupportedDependency(Unsupported::ExternalOwner));
            }
            let observed = formula_edges
                .len()
                .checked_add(1)
                .ok_or(Failure::InvalidSource)?;
            check_limit(observed, limits.graph_edges)?;
            graph_step(usage, 1, limits)?;
            scratch_allocation::<FormulaEdge>(usage, 1, limits)?;
            reserve(&mut formula_edges, 1)?;
            formula_edges.push(FormulaEdge {
                source: FormulaKey {
                    owner: precedent.owner,
                    coordinate: precedent.coordinate,
                },
                target: FormulaKey {
                    owner: formula.owner,
                    coordinate: formula.coordinate,
                },
            });
        }
    }
    graph_step(usage, sort_work(formula_edges.len())?, limits)?;
    formula_edges.sort_unstable();
    graph_step(usage, formula_edges.len(), limits)?;
    formula_edges.dedup();
    if formula_edges.len() != dependency_edges.len()
        || formula_edges
            .iter()
            .zip(dependency_edges)
            .any(|(formula, dependency)| {
                formula.source != dependency.source || formula.target != dependency.target
            })
    {
        return Err(Failure::UnsupportedDependency(Unsupported::Formula));
    }
    Ok(())
}

fn payload_index(
    payloads: &[DependencyPayload<'_>],
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Vec<PayloadIndex>, Failure> {
    let mut result = Vec::new();
    graph_step(
        usage,
        payloads
            .len()
            .checked_add(sort_work(payloads.len())?)
            .ok_or(Failure::InvalidSource)?,
        limits,
    )?;
    scratch_allocation::<PayloadIndex>(usage, payloads.len(), limits)?;
    reserve(&mut result, payloads.len())?;
    for (index, payload) in payloads.iter().enumerate() {
        add_usage(&mut usage.lookup_work, 1)?;
        result.push(PayloadIndex {
            identifier: payload.identifier,
            index,
        });
    }
    result.sort_unstable();
    if result
        .windows(2)
        .any(|pair| pair[0].identifier == pair[1].identifier)
    {
        return Err(Failure::InvalidSource);
    }
    Ok(result)
}

fn formula_index(
    source: CacheSource<'_>,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Vec<FormulaIndex>, Failure> {
    let formulas = source.formulas;
    let mut result = Vec::new();
    graph_step(
        usage,
        formulas
            .len()
            .checked_add(sort_work(formulas.len())?)
            .ok_or(Failure::InvalidSource)?,
        limits,
    )?;
    scratch_allocation::<FormulaIndex>(usage, formulas.len(), limits)?;
    reserve(&mut result, formulas.len())?;
    for (index, formula) in formulas.iter().enumerate() {
        add_usage(&mut usage.lookup_work, 1)?;
        if formula.owner != source.selected_table.owner
            || formula.coordinate.row >= source.rows
            || formula.coordinate.column >= source.columns
        {
            return Err(Failure::UnsupportedDependency(Unsupported::ExternalOwner));
        }
        result.push(FormulaIndex {
            key: FormulaKey {
                owner: formula.owner,
                coordinate: formula.coordinate,
            },
            index,
        });
    }
    result.sort_unstable();
    if result.windows(2).any(|pair| pair[0].key == pair[1].key) {
        return Err(Failure::InvalidSource);
    }
    Ok(result)
}

fn formula_coordinate_index(
    index: &[FormulaIndex],
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Vec<CoordinateIndex>, Failure> {
    let mut result = Vec::new();
    graph_step(
        usage,
        index
            .len()
            .checked_add(sort_work(index.len())?)
            .ok_or(Failure::InvalidSource)?,
        limits,
    )?;
    scratch_allocation::<CoordinateIndex>(usage, index.len(), limits)?;
    reserve(&mut result, index.len())?;
    for (formula_position, entry) in index.iter().enumerate() {
        add_usage(&mut usage.lookup_work, 1)?;
        result.push(CoordinateIndex {
            coordinate: entry.key.coordinate,
            formula_position,
        });
    }
    result.sort_unstable();
    if result
        .windows(2)
        .any(|pair| pair[0].coordinate == pair[1].coordinate)
    {
        // One selected table must not expose two formula owners for one cell.
        return Err(Failure::InvalidSource);
    }
    Ok(result)
}

fn indexed_payload<'source>(
    payloads: &'source [DependencyPayload<'source>],
    index: &[PayloadIndex],
    identifier: u64,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<DependencyPayload<'source>, Failure> {
    graph_step(usage, binary_search_work(index.len()), limits)?;
    add_usage(&mut usage.lookup_work, 1)?;
    let position = index
        .binary_search_by_key(&identifier, |entry| entry.identifier)
        .map_err(|_error| Failure::InvalidSource)?;
    Ok(payloads[index[position].index])
}

fn verify_owner_references(
    references: &mut [u64],
    owner_index: &[PayloadIndex],
) -> Result<(), Failure> {
    if references.len() != owner_index.len() {
        return Err(Failure::InvalidSource);
    }
    references.sort_unstable();
    if references.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Failure::InvalidSource);
    }
    for (reference, payload) in references.iter().zip(owner_index) {
        if *reference != payload.identifier {
            return Err(Failure::InvalidSource);
        }
    }
    Ok(())
}

fn ensure_overlay(source: CacheSource<'_>, overlay: &[FinalCell]) -> Result<(), Failure> {
    if overlay
        .windows(2)
        .any(|pair| pair[0].coordinate >= pair[1].coordinate)
    {
        return Err(Failure::InvalidSource);
    }
    if overlay.iter().any(|cell| {
        cell.table != source.selected_table
            || cell.coordinate.row >= source.rows
            || cell.coordinate.column >= source.columns
    }) {
        return Err(Failure::UnsupportedDependency(Unsupported::ExternalOwner));
    }
    Ok(())
}

fn changed_formula_removals(
    overlay: &[FinalCell],
    formulas: &[FormulaCell],
    formula_index: &[FormulaIndex],
    coordinate_index: &[CoordinateIndex],
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Vec<DependencyRemoval>, Failure> {
    let mut removals = Vec::new();
    retained_allocation::<DependencyRemoval>(usage, overlay.len(), limits)?;
    reserve(&mut removals, overlay.len())?;
    for changed in overlay {
        let position =
            coordinate_index.binary_search_by_key(&changed.coordinate, |entry| entry.coordinate);
        if let Ok(position) = position {
            let entry = formula_index[coordinate_index[position].formula_position];
            let formula = formulas[entry.index];
            removals.push(DependencyRemoval {
                owner: formula.owner,
                coordinate: formula.coordinate,
            });
        }
    }
    removals.sort_unstable();
    removals.dedup();
    Ok(removals)
}

fn find_formula(
    index: &[FormulaIndex],
    key: FormulaKey,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Option<usize>, Failure> {
    graph_step(usage, binary_search_work(index.len()), limits)?;
    add_usage(&mut usage.lookup_work, 1)?;
    Ok(index.binary_search_by(|entry| entry.key.cmp(&key)).ok())
}

fn find_coordinate_formula(
    index: &[CoordinateIndex],
    coordinate: Coordinate,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Option<usize>, Failure> {
    graph_step(usage, binary_search_work(index.len()), limits)?;
    add_usage(&mut usage.lookup_work, 1)?;
    Ok(index
        .binary_search_by_key(&coordinate, |entry| entry.coordinate)
        .ok()
        .map(|position| index[position].formula_position))
}

fn overlay_has(
    overlay: &[FinalCell],
    key: FormulaKey,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<bool, Failure> {
    graph_step(usage, binary_search_work(overlay.len()), limits)?;
    add_usage(&mut usage.lookup_work, 1)?;
    Ok(overlay
        .binary_search_by(|cell| {
            (cell.table.owner, cell.coordinate).cmp(&(key.owner, key.coordinate))
        })
        .is_ok())
}

fn edge_range(
    edges: &[Edge],
    source: FormulaKey,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<(usize, usize), Failure> {
    graph_step(
        usage,
        binary_search_work(edges.len()).saturating_mul(2),
        limits,
    )?;
    add_usage(&mut usage.lookup_work, 1)?;
    let start = edges.partition_point(|edge| edge.source < source);
    let end = edges.partition_point(|edge| edge.source <= source);
    Ok((start, end))
}

fn option_slots(
    length: usize,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Vec<Option<FormulaCachedValue>>, Failure> {
    let mut slots = Vec::new();
    scratch_allocation::<Option<FormulaCachedValue>>(usage, length, limits)?;
    reserve(&mut slots, length)?;
    slots.resize_with(length, || None);
    Ok(slots)
}

fn bool_slots(
    length: usize,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Vec<bool>, Failure> {
    let mut slots = Vec::new();
    scratch_allocation::<bool>(usage, length, limits)?;
    reserve(&mut slots, length)?;
    slots.resize(length, false);
    Ok(slots)
}

fn usize_slots(
    length: usize,
    usage: &mut CacheUsage,
    limits: CacheLimits,
) -> Result<Vec<usize>, Failure> {
    let mut slots = Vec::new();
    scratch_allocation::<usize>(usage, length, limits)?;
    reserve(&mut slots, length)?;
    slots.resize(length, 0);
    Ok(slots)
}

struct FinalValues<'a> {
    overlay: &'a [FinalCell],
    coordinate_index: &'a [CoordinateIndex],
    computed: &'a [Option<FormulaCachedValue>],
    baseline: &'a dyn CacheBaseline,
}

impl CacheValues for FinalValues<'_> {
    fn value(&self, coordinate: Coordinate) -> Result<Option<&FormulaCachedValue>, Failure> {
        if let Ok(position) = self
            .overlay
            .binary_search_by(|cell| cell.coordinate.cmp(&coordinate))
        {
            return self.overlay[position].value.supported();
        }
        if let Ok(position) = self
            .coordinate_index
            .binary_search_by_key(&coordinate, |entry| entry.coordinate)
        {
            let formula_position = self.coordinate_index[position].formula_position;
            if let Some(value) = self.computed[formula_position].as_ref() {
                return supported_value(Some(value));
            }
        }
        supported_value(self.baseline.value(coordinate)?)
    }
}

fn supported_value(
    value: Option<&FormulaCachedValue>,
) -> Result<Option<&FormulaCachedValue>, Failure> {
    match value {
        None | Some(FormulaCachedValue::Number(_) | FormulaCachedValue::Boolean(_)) => Ok(value),
        Some(
            FormulaCachedValue::Text(_)
            | FormulaCachedValue::Date(_)
            | FormulaCachedValue::Duration(_),
        ) => Err(Failure::UnsupportedDependency(Unsupported::CacheType)),
    }
}

fn grouped_rewrites(
    formulas: &[FormulaCell],
    index: &[FormulaIndex],
    computed: &mut [Option<FormulaCachedValue>],
    limits: CacheLimits,
    usage: &mut CacheUsage,
) -> Result<Vec<CacheRewrite>, Failure> {
    let mut cells = Vec::new();
    scratch_allocation::<(u64, CacheCellRewrite)>(usage, computed.len(), limits)?;
    reserve(&mut cells, computed.len())?;
    for (position, value) in computed.iter_mut().enumerate() {
        let Some(value) = value.take() else {
            continue;
        };
        check_limit(
            cells.len().checked_add(1).ok_or(Failure::InvalidSource)?,
            limits.cache_cells,
        )?;
        let formula = formulas[index[position].index];
        cells.push((
            formula.cache_object,
            CacheCellRewrite {
                owner: formula.owner,
                coordinate: formula.coordinate,
                value,
            },
        ));
    }
    cells.sort_unstable_by(|left, right| {
        left.0
            .cmp(&right.0)
            .then_with(|| left.1.owner.cmp(&right.1.owner))
            .then_with(|| left.1.coordinate.cmp(&right.1.coordinate))
    });
    let mut grouped = Vec::new();
    let mut current_object = None;
    let mut current_cells = Vec::new();
    let refreshed_hosts = cells.len();
    for (object, cell) in cells {
        if current_object.is_some_and(|current| current != object) {
            retained_allocation::<CacheRewrite>(usage, 1, limits)?;
            reserve(&mut grouped, 1)?;
            grouped.push(CacheRewrite {
                cache_object: current_object.ok_or(Failure::InvalidSource)?,
                cells: core::mem::take(&mut current_cells),
            });
        }
        current_object = Some(object);
        retained_allocation::<CacheCellRewrite>(usage, 1, limits)?;
        reserve(&mut current_cells, 1)?;
        current_cells.push(cell);
    }
    if let Some(cache_object) = current_object {
        retained_allocation::<CacheRewrite>(usage, 1, limits)?;
        reserve(&mut grouped, 1)?;
        grouped.push(CacheRewrite {
            cache_object,
            cells: current_cells,
        });
    }
    add_usage(&mut usage.cache_cells_read, computed.len())?;
    add_usage(&mut usage.cache_hosts_refreshed, refreshed_hosts)?;
    Ok(grouped)
}

fn charge_report(
    usage: &mut CacheUsage,
    report: dependency::DecodeReport,
    limits: CacheLimits,
) -> Result<(), Failure> {
    charge(
        &mut usage.wire_bytes,
        report.source_bytes(),
        limits.wire_bytes,
    )?;
    charge(&mut usage.wire_fields, report.fields(), limits.wire_fields)?;
    charge(&mut usage.wire_work, report.work_bytes(), limits.wire_work)?;
    charge(
        &mut usage.wire_references,
        report.references(),
        limits.wire_references,
    )?;
    charge(
        &mut usage.wire_text_bytes,
        report.text_bytes(),
        limits.wire_text,
    )
}

fn charge_formula_evaluation_usage(
    usage: &mut CacheUsage,
    evaluation: FormulaEvaluationUsage,
    limits: CacheLimits,
) -> Result<(), Failure> {
    charge(
        &mut usage.wire_bytes,
        evaluation.wire_bytes,
        limits.wire_bytes,
    )?;
    charge(
        &mut usage.wire_fields,
        evaluation.wire_fields,
        limits.wire_fields,
    )?;
    charge(&mut usage.wire_work, evaluation.wire_work, limits.wire_work)?;
    charge(
        &mut usage.wire_text_bytes,
        evaluation.wire_text_bytes,
        limits.wire_text,
    )?;
    charge(
        &mut usage.formula_work,
        evaluation.formula_work,
        limits.formula_work,
    )?;
    graph_step(usage, evaluation.formula_work, limits)?;
    charge(
        &mut usage.peak_scratch_bytes,
        evaluation.scratch_bytes,
        limits.scratch_bytes,
    )?;
    charge(
        &mut usage.allocations,
        evaluation.allocations,
        limits.allocations,
    )
}

/// Derive exact residual cache limits after a previously completed bounded
/// stage. Callers use this to compose host/list/coverage/planning stages
/// without authorizing the same remaining budget more than once.
pub(super) fn remaining_limits(
    limits: CacheLimits,
    usage: &CacheUsage,
) -> Result<CacheLimits, Failure> {
    fn remaining(maximum: usize, used: u64) -> Result<usize, Failure> {
        let maximum_u64 = u64::try_from(maximum).map_err(|_error| Failure::InvalidSource)?;
        let value = maximum_u64
            .checked_sub(used)
            .ok_or(Failure::InvalidSource)?;
        usize::try_from(value).map_err(|_error| Failure::InvalidSource)
    }

    Ok(CacheLimits {
        wire_bytes: remaining(limits.wire_bytes, usage.wire_bytes)?,
        wire_fields: remaining(limits.wire_fields, usage.wire_fields)?,
        wire_work: remaining(limits.wire_work, usage.wire_work)?,
        wire_references: remaining(limits.wire_references, usage.wire_references)?,
        wire_text: remaining(limits.wire_text, usage.wire_text_bytes)?,
        nesting: limits.nesting,
        graph_nodes: limits.graph_nodes,
        graph_edges: limits.graph_edges,
        cache_cells: limits.cache_cells,
        formula_work: remaining(limits.formula_work, usage.formula_work)?,
        graph_work: remaining(limits.graph_work, usage.graph_work)?,
        retained_bytes: remaining(limits.retained_bytes, usage.retained_bytes)?,
        scratch_bytes: remaining(limits.scratch_bytes, usage.peak_scratch_bytes)?,
        allocations: remaining(limits.allocations, usage.allocations)?,
    })
}

fn charge_stage_usage(
    usage: &mut CacheUsage,
    stage: &Stage,
    limits: CacheLimits,
) -> Result<(), Failure> {
    charge(
        &mut usage.peak_scratch_bytes,
        stage.retained_bytes,
        limits.scratch_bytes,
    )?;
    charge(
        &mut usage.allocations,
        stage.allocations,
        limits.allocations,
    )
}

fn charge(total: &mut u64, increment: usize, maximum: usize) -> Result<(), Failure> {
    let next = total
        .checked_add(u64::try_from(increment).map_err(|_error| Failure::InvalidSource)?)
        .ok_or(Failure::InvalidSource)?;
    let maximum = u64::try_from(maximum).map_err(|_error| Failure::InvalidSource)?;
    if next > maximum {
        return Err(Failure::LimitExceeded {
            observed: next,
            maximum,
        });
    }
    *total = next;
    Ok(())
}

fn check_limit(observed: usize, maximum: usize) -> Result<(), Failure> {
    if observed > maximum {
        return Err(Failure::LimitExceeded {
            observed: u64::try_from(observed).map_err(|_error| Failure::InvalidSource)?,
            maximum: u64::try_from(maximum).map_err(|_error| Failure::InvalidSource)?,
        });
    }
    Ok(())
}

fn graph_step(
    usage: &mut CacheUsage,
    increment: usize,
    limits: CacheLimits,
) -> Result<(), Failure> {
    charge(&mut usage.graph_work, increment, limits.graph_work)
}

fn scratch_allocation<T>(
    usage: &mut CacheUsage,
    elements: usize,
    limits: CacheLimits,
) -> Result<(), Failure> {
    allocation_bytes::<T>(
        &mut usage.peak_scratch_bytes,
        elements,
        limits.scratch_bytes,
    )?;
    if elements != 0 {
        charge(&mut usage.allocations, 1, limits.allocations)?;
    }
    Ok(())
}

fn retained_allocation<T>(
    usage: &mut CacheUsage,
    elements: usize,
    limits: CacheLimits,
) -> Result<(), Failure> {
    add_usage(&mut usage.retained_elements, elements)?;
    allocation_bytes::<T>(&mut usage.retained_bytes, elements, limits.retained_bytes)?;
    if elements != 0 {
        charge(&mut usage.allocations, 1, limits.allocations)?;
    }
    Ok(())
}

fn allocation_bytes<T>(total: &mut u64, elements: usize, maximum: usize) -> Result<(), Failure> {
    let bytes = elements
        .checked_mul(size_of::<T>())
        .ok_or(Failure::InvalidSource)?;
    charge(total, bytes, maximum)
}

fn binary_search_work(length: usize) -> usize {
    if length <= 1 {
        1
    } else {
        usize::try_from(usize::BITS - (length - 1).leading_zeros()).unwrap_or(usize::MAX)
    }
}

fn sort_work(length: usize) -> Result<usize, Failure> {
    length
        .checked_mul(binary_search_work(length))
        .ok_or(Failure::InvalidSource)
}

fn add_usage(total: &mut u64, increment: usize) -> Result<(), Failure> {
    *total = total
        .checked_add(u64::try_from(increment).map_err(|_error| Failure::InvalidSource)?)
        .ok_or(Failure::InvalidSource)?;
    Ok(())
}

fn reserve<T>(values: &mut Vec<T>, additional: usize) -> Result<(), Failure> {
    values
        .try_reserve_exact(additional)
        .map_err(|_error| Failure::Allocation { amount: additional })
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_protos::{tsce, tsp};

    const OWNER: u32 = 7;
    const TABLE: TableIdentity = TableIdentity {
        owner: OWNER,
        uuid_lower: 11,
        uuid_upper: 13,
    };

    struct Baseline(Vec<(Coordinate, FormulaCachedValue)>);

    impl CacheBaseline for Baseline {
        fn value(&self, coordinate: Coordinate) -> Result<Option<&FormulaCachedValue>, Failure> {
            Ok(self
                .0
                .iter()
                .find_map(|(candidate, value)| (*candidate == coordinate).then_some(value)))
        }
    }

    impl CacheValues for Baseline {
        fn value(&self, coordinate: Coordinate) -> Result<Option<&FormulaCachedValue>, Failure> {
            CacheBaseline::value(self, coordinate)
        }
    }

    #[derive(Default)]
    struct ChainEvaluator {
        calls: Vec<Coordinate>,
    }

    impl CacheEvaluator for ChainEvaluator {
        fn analyze(
            &mut self,
            formula: FormulaCell,
            _limits: CacheLimits,
        ) -> Result<FormulaAnalysis, Failure> {
            let column = formula
                .coordinate
                .column
                .checked_sub(1)
                .ok_or(Failure::InvalidSource)?;
            Ok(FormulaAnalysis {
                precedents: vec![FormulaPrecedent {
                    owner: formula.owner,
                    coordinate: Coordinate {
                        row: formula.coordinate.row,
                        column,
                    },
                }],
                usage: FormulaEvaluationUsage {
                    formula_work: 1,
                    allocations: 1,
                    scratch_bytes: size_of::<FormulaPrecedent>(),
                    ..FormulaEvaluationUsage::default()
                },
            })
        }

        fn evaluate(
            &mut self,
            formula: FormulaCell,
            values: &dyn CacheValues,
            _limits: CacheLimits,
        ) -> Result<Evaluation, Failure> {
            self.calls.push(formula.coordinate);
            let preceding = Coordinate {
                row: formula.coordinate.row,
                column: formula
                    .coordinate
                    .column
                    .checked_sub(1)
                    .ok_or(Failure::InvalidSource)?,
            };
            let number = match values.value(preceding)? {
                Some(FormulaCachedValue::Number(number)) => number.get(),
                _ => return Err(Failure::InvalidSource),
            };
            Ok(Evaluation {
                value: FormulaCachedValue::number(number + 1.0)
                    .map_err(|_error| Failure::InvalidSource)?,
                usage: FormulaEvaluationUsage {
                    formula_work: 1,
                    ..FormulaEvaluationUsage::default()
                },
            })
        }
    }

    #[derive(Default)]
    struct FanoutEvaluator;

    impl CacheEvaluator for FanoutEvaluator {
        fn analyze(
            &mut self,
            formula: FormulaCell,
            _limits: CacheLimits,
        ) -> Result<FormulaAnalysis, Failure> {
            Ok(FormulaAnalysis {
                precedents: vec![FormulaPrecedent {
                    owner: formula.owner,
                    coordinate: coordinate(0),
                }],
                usage: FormulaEvaluationUsage {
                    formula_work: 1,
                    allocations: 1,
                    scratch_bytes: size_of::<FormulaPrecedent>(),
                    ..FormulaEvaluationUsage::default()
                },
            })
        }

        fn evaluate(
            &mut self,
            _formula: FormulaCell,
            values: &dyn CacheValues,
            _limits: CacheLimits,
        ) -> Result<Evaluation, Failure> {
            let number = match values.value(coordinate(0))? {
                Some(FormulaCachedValue::Number(number)) => number.get(),
                _ => return Err(Failure::InvalidSource),
            };
            Ok(Evaluation {
                value: FormulaCachedValue::number(number + 1.0)
                    .map_err(|_error| Failure::InvalidSource)?,
                usage: FormulaEvaluationUsage {
                    formula_work: 1,
                    ..FormulaEvaluationUsage::default()
                },
            })
        }
    }

    fn coordinate(column: u32) -> Coordinate {
        Coordinate { row: 0, column }
    }

    fn formula(column: u32, cache_object: u64) -> FormulaCell {
        FormulaCell {
            owner: OWNER,
            coordinate: coordinate(column),
            cache_object,
        }
    }

    fn record(host: u32, precedents: &[u32], cycle: bool) -> tsce::CellRecordExpandedArchive {
        tsce::CellRecordExpandedArchive {
            row: 0,
            column: host,
            is_in_a_cycle: cycle.then_some(true),
            expanded_edges: Some(tsce::ExpandedEdgesArchive {
                edge_without_owner_rows: vec![0; precedents.len()],
                edge_without_owner_columns: precedents.to_vec(),
                ..Default::default()
            }),
            ..Default::default()
        }
    }

    fn encoded_graph(
        records: Vec<tsce::CellRecordExpandedArchive>,
        volatile: bool,
    ) -> (Vec<u8>, Vec<u8>) {
        let owner = tsce::FormulaOwnerDependenciesArchive {
            formula_owner_uid: tsp::Uuid {
                lower: TABLE.uuid_lower,
                upper: TABLE.uuid_upper,
            },
            internal_formula_owner_id: OWNER,
            cell_dependencies: Some(tsce::CellDependenciesExpandedArchive {
                cell_record: records,
            }),
            volatile_dependencies: volatile.then(|| tsce::VolatileDependenciesExpandedArchive {
                volatile_time_cells: Some(tsce::CellCoordSetArchive {
                    column_entries: vec![tsce::cell_coord_set_archive::ColumnEntry {
                        column: 0,
                        row_set: tsce::IndexSetArchive::default(),
                    }],
                }),
                ..Default::default()
            }),
            ..Default::default()
        }
        .encode_to_vec();
        let engine = tsce::CalculationEngineArchive {
            dependency_tracker: tsce::DependencyTrackerArchive {
                formula_owner_dependencies: vec![tsp::Reference {
                    identifier: 1,
                    ..Default::default()
                }],
                ..Default::default()
            },
            ..Default::default()
        }
        .encode_to_vec();
        (engine, owner)
    }

    fn encoded_formula_host_sources() -> (Vec<u8>, Vec<u8>) {
        let owner = tsce::FormulaOwnerDependenciesArchive {
            formula_owner_uid: tsp::Uuid {
                lower: TABLE.uuid_lower,
                upper: TABLE.uuid_upper,
            },
            internal_formula_owner_id: OWNER,
            cell_dependencies: Some(tsce::CellDependenciesExpandedArchive {
                cell_record: vec![record(2, &[0], false)],
            }),
            tiled_cell_dependencies: Some(tsce::CellDependenciesTiledArchive {
                cell_record_tiles: vec![tsp::Reference {
                    identifier: 905_369,
                    ..Default::default()
                }],
            }),
            ..Default::default()
        }
        .encode_to_vec();
        let tile = tsce::CellRecordTileArchive {
            internal_owner_id: OWNER,
            tile_column_begin: 0,
            tile_row_begin: 0,
            cell_records: vec![record(1, &[0], false), record(2, &[0], false)],
        }
        .encode_to_vec();
        (owner, tile)
    }

    fn encoded_native_dormant_graph(
        records: Vec<tsce::CellRecordExpandedArchive>,
        hostile: bool,
    ) -> (Vec<u8>, Vec<u8>) {
        let total = tsce::RangeCoordinateArchive {
            top_left_column: 0,
            top_left_row: 0,
            bottom_right_column: 31,
            bottom_right_row: 0,
        };
        let body = tsce::RangeCoordinateArchive {
            bottom_right_column: if hostile { 30 } else { 31 },
            ..total
        };
        let owner = tsce::FormulaOwnerDependenciesArchive {
            formula_owner_uid: tsp::Uuid {
                lower: TABLE.uuid_lower,
                upper: TABLE.uuid_upper,
            },
            internal_formula_owner_id: OWNER,
            cell_dependencies: Some(tsce::CellDependenciesExpandedArchive {
                cell_record: records,
            }),
            range_dependencies: Some(tsce::RangeDependenciesArchive::default()),
            volatile_dependencies: Some(tsce::VolatileDependenciesExpandedArchive::default()),
            spanning_column_dependencies: Some(tsce::SpanningDependenciesExpandedArchive {
                total_range_for_table: Some(total),
                body_range_for_table: Some(body),
                ..Default::default()
            }),
            spanning_row_dependencies: Some(tsce::SpanningDependenciesExpandedArchive {
                total_range_for_table: Some(total),
                body_range_for_table: Some(body),
                ..Default::default()
            }),
            whole_owner_dependencies: Some(tsce::WholeOwnerDependenciesExpandedArchive::default()),
            uuid_references: Some(tsce::UuidReferencesArchive::default()),
            spill_range_sizes: Some(tsce::CellSpillSizesArchive::default()),
            ..Default::default()
        }
        .encode_to_vec();
        let engine = tsce::CalculationEngineArchive {
            dependency_tracker: tsce::DependencyTrackerArchive {
                formula_owner_dependencies: vec![tsp::Reference {
                    identifier: 1,
                    ..Default::default()
                }],
                ..Default::default()
            },
            ..Default::default()
        }
        .encode_to_vec();
        (engine, owner)
    }

    fn limits() -> CacheLimits {
        CacheLimits {
            wire_bytes: 1 << 20,
            wire_fields: 1 << 20,
            wire_work: 1 << 24,
            wire_references: 1 << 10,
            wire_text: 1 << 20,
            nesting: 64,
            graph_nodes: 1 << 10,
            graph_edges: 1 << 10,
            cache_cells: 1 << 10,
            formula_work: 1 << 10,
            graph_work: 1 << 20,
            retained_bytes: 1 << 20,
            scratch_bytes: 1 << 20,
            allocations: 1 << 10,
        }
    }

    fn plan<'a>(
        engine: &'a [u8],
        owner: &'a [u8],
        formulas: &'a [FormulaCell],
        overlay: &'a [FinalCell],
        limits: CacheLimits,
        evaluator: &mut dyn CacheEvaluator,
    ) -> Result<CachePlan, Failure> {
        plan_final_cache(
            CacheSource {
                selected_table: TABLE,
                rows: 1,
                columns: 32,
                header_rows: 0,
                header_columns: 0,
                footer_rows: 0,
                engine,
                owners: &[DependencyPayload {
                    identifier: 1,
                    bytes: owner,
                }],
                record_tiles: &[],
                formulas,
            },
            overlay,
            limits,
            &Baseline(Vec::new()),
            evaluator,
        )
    }

    #[test]
    fn formula_chain_uses_topological_final_values() {
        let (engine, owner) =
            encoded_graph(vec![record(1, &[0], false), record(2, &[1], false)], false);
        let formulas = [formula(1, 20), formula(2, 20)];
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(2.0).expect("finite")),
        }];
        let mut evaluator = ChainEvaluator::default();
        let result = plan(
            &engine,
            &owner,
            &formulas,
            &overlay,
            limits(),
            &mut evaluator,
        )
        .expect("chain plan");

        assert_eq!(evaluator.calls, vec![coordinate(1), coordinate(2)]);
        assert_eq!(result.rewrites.len(), 1);
        assert_eq!(result.rewrites[0].cells.len(), 2);
        assert!(matches!(
            result.rewrites[0].cells[1].value,
            FormulaCachedValue::Number(value) if value.get() == 4.0
        ));
    }

    #[test]
    fn strict_formula_chain_refreshes_two_hosts_from_final_overlay() {
        use tsce::ast_node_array_archive::{
            AstLocalCellReferenceNodeArchive, AstNodeArchive, AstNodeType,
        };

        let encoded = |precedent_column: u32| {
            tsce::FormulaArchive {
                ast_node_array: tsce::AstNodeArrayArchive {
                    ast_node: vec![
                        AstNodeArchive {
                            ast_node_type: AstNodeType::LocalCellReferenceNode as i32,
                            ast_local_cell_reference_node_reference: Some(
                                AstLocalCellReferenceNodeArchive {
                                    row_handle: 0,
                                    column_handle: precedent_column,
                                    row_is_sticky: 0,
                                    column_is_sticky: 0,
                                },
                            ),
                            ..Default::default()
                        },
                        AstNodeArchive {
                            ast_node_type: AstNodeType::NumberNode as i32,
                            ast_number_node_number: Some(1.0),
                            ..Default::default()
                        },
                        AstNodeArchive {
                            ast_node_type: AstNodeType::AdditionNode as i32,
                            ..Default::default()
                        },
                    ],
                },
                ..Default::default()
            }
            .encode_to_vec()
        };
        let b = encoded(0);
        let c = encoded(1);
        let entries = [
            FormulaListEntry {
                key: 1,
                ref_count: 1,
                bytes: &b,
            },
            FormulaListEntry {
                key: 2,
                ref_count: 1,
                bytes: &c,
            },
        ];
        let payloads = [
            FormulaPayload {
                owner: OWNER,
                coordinate: coordinate(1),
                key: 1,
                bytes: &b,
            },
            FormulaPayload {
                owner: OWNER,
                coordinate: coordinate(2),
                key: 2,
                bytes: &c,
            },
        ];
        let formulas = [formula(1, 20), formula(2, 20)];
        let (mut evaluator, _coverage) =
            StrictEvaluator::new(&entries, &payloads, 1, 4, TABLE, limits())
                .expect("strict evaluator");
        let (engine, owner) =
            encoded_graph(vec![record(1, &[0], false), record(2, &[1], false)], false);
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(2.0).expect("finite")),
        }];
        let result = plan(
            &engine,
            &owner,
            &formulas,
            &overlay,
            limits(),
            &mut evaluator,
        )
        .expect("strict chain plan");
        assert_eq!(result.rewrites[0].cells.len(), 2);
        assert!(matches!(
            result.rewrites[0].cells[1].value,
            FormulaCachedValue::Number(value) if value.get() == 4.0
        ));
    }

    #[test]
    fn unsupported_local_formula_is_preserved_only_when_proven_unreachable() {
        use tsce::ast_node_array_archive::{
            AstLocalCellReferenceNodeArchive, AstNodeArchive, AstNodeType,
        };

        let bytes = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![
                    AstNodeArchive {
                        ast_node_type: AstNodeType::LocalCellReferenceNode as i32,
                        ast_local_cell_reference_node_reference: Some(
                            AstLocalCellReferenceNodeArchive {
                                row_handle: 0,
                                column_handle: 0,
                                row_is_sticky: 0,
                                column_is_sticky: 0,
                            },
                        ),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::FunctionNode as i32,
                        ast_function_node_index: Some(999),
                        ast_function_node_num_args: Some(1),
                        ..Default::default()
                    },
                ],
            },
            ..Default::default()
        }
        .encode_to_vec();
        let entries = [FormulaListEntry {
            key: 1,
            ref_count: 1,
            bytes: &bytes,
        }];
        let payloads = [FormulaPayload {
            owner: OWNER,
            coordinate: coordinate(1),
            key: 1,
            bytes: &bytes,
        }];
        let formulas = [formula(1, 20)];
        let unrelated = [FinalCell {
            table: TABLE,
            coordinate: coordinate(2),
            value: FinalValue::number(FiniteF64::new(9.0).expect("finite")),
        }];
        let impacted = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(9.0).expect("finite")),
        }];
        let (engine, owner) = encoded_graph(vec![record(1, &[0], false)], false);
        let (mut evaluator, _coverage) =
            StrictEvaluator::new(&entries, &payloads, 1, 32, TABLE, limits())
                .expect("strict evaluator");

        let preserved = plan(
            &engine,
            &owner,
            &formulas,
            &unrelated,
            limits(),
            &mut evaluator,
        )
        .expect("unreachable unsupported formula is byte-preserved");
        assert!(preserved.rewrites.is_empty());

        assert!(matches!(
            plan(
                &engine,
                &owner,
                &formulas,
                &impacted,
                limits(),
                &mut evaluator,
            ),
            Err(Failure::UnsupportedDependency(Unsupported::Formula))
        ));

        let (missing_engine, missing_owner) = encoded_graph(vec![record(1, &[], false)], false);
        assert!(matches!(
            plan(
                &missing_engine,
                &missing_owner,
                &formulas,
                &unrelated,
                limits(),
                &mut evaluator,
            ),
            Err(Failure::UnsupportedDependency(Unsupported::Formula))
        ));
    }

    #[test]
    fn formula_hosts_are_sorted_and_deduplicate_inline_and_tiled_records() {
        let (owner, tile) = encoded_formula_host_sources();
        let owners = [DependencyPayload {
            identifier: 904_993,
            bytes: &owner,
        }];
        let tiles = [DependencyPayload {
            identifier: 905_369,
            bytes: &tile,
        }];
        let (hosts, usage) =
            collect_formula_hosts(TABLE, &owners, &tiles, limits()).expect("bounded formula hosts");
        assert_eq!(
            hosts,
            vec![
                FormulaHost {
                    owner: OWNER,
                    coordinate: coordinate(1),
                },
                FormulaHost {
                    owner: OWNER,
                    coordinate: coordinate(2),
                },
            ]
        );
        assert_eq!(usage.formula_graph_builds, 1);
        assert_eq!(usage.formula_nodes, 2);
        assert_eq!(
            hosts[0].into_formula_cell(77),
            FormulaCell {
                owner: OWNER,
                coordinate: coordinate(1),
                cache_object: 77,
            }
        );
    }

    #[test]
    fn formula_host_collection_bounds_raw_inline_and_tiled_candidates() {
        let (owner, tile) = encoded_formula_host_sources();
        let owners = [DependencyPayload {
            identifier: 904_993,
            bytes: &owner,
        }];
        let tiles = [DependencyPayload {
            identifier: 905_369,
            bytes: &tile,
        }];
        let mut bounded = limits();
        bounded.graph_nodes = 2;
        assert!(matches!(
            collect_formula_hosts(TABLE, &owners, &tiles, bounded),
            Err(Failure::LimitExceeded {
                observed: 3,
                maximum: 2,
            })
        ));
    }

    #[test]
    fn final_overlay_replaces_formula_before_downstream_refresh() {
        let (engine, owner) =
            encoded_graph(vec![record(1, &[0], false), record(2, &[1], false)], false);
        let formulas = [formula(1, 20), formula(2, 20)];
        let overlay = [
            FinalCell {
                table: TABLE,
                coordinate: coordinate(0),
                value: FinalValue::number(FiniteF64::new(2.0).expect("finite")),
            },
            FinalCell {
                table: TABLE,
                coordinate: coordinate(1),
                value: FinalValue::number(FiniteF64::new(10.0).expect("finite")),
            },
        ];
        let mut evaluator = ChainEvaluator::default();
        let result = plan(
            &engine,
            &owner,
            &formulas,
            &overlay,
            limits(),
            &mut evaluator,
        )
        .expect("final-state plan");

        assert_eq!(evaluator.calls, vec![coordinate(2)]);
        assert_eq!(result.removals.len(), 1);
        assert_eq!(result.removals[0].coordinate, coordinate(1));
        assert!(matches!(
            result.rewrites[0].cells[0].value,
            FormulaCachedValue::Number(value) if value.get() == 11.0
        ));
    }

    #[test]
    fn reachable_cycle_and_missing_formula_fail_closed_before_evaluation() {
        let formulas = [formula(1, 20), formula(2, 20)];
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(1.0).expect("finite")),
        }];
        let (cycle_engine, cycle_owner) = encoded_graph(
            vec![record(1, &[0, 2], false), record(2, &[1], false)],
            false,
        );
        let mut evaluator = ChainEvaluator::default();
        let cycle = plan(
            &cycle_engine,
            &cycle_owner,
            &formulas,
            &overlay,
            limits(),
            &mut evaluator,
        );
        assert!(matches!(
            cycle,
            Err(Failure::UnsupportedDependency(Unsupported::Formula))
        ));
        assert!(evaluator.calls.is_empty());

        let (missing_engine, missing_owner) = encoded_graph(vec![record(3, &[0], false)], false);
        let missing = plan(
            &missing_engine,
            &missing_owner,
            &formulas,
            &overlay,
            limits(),
            &mut evaluator,
        );
        assert!(matches!(
            missing,
            Err(Failure::UnsupportedDependency(Unsupported::Formula))
        ));
        assert!(evaluator.calls.is_empty());
    }

    #[test]
    fn omitted_dependency_edge_is_refused_before_evaluation() {
        let (engine, owner) = encoded_graph(vec![record(1, &[], false)], false);
        let formulas = [formula(1, 20)];
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(1.0).expect("finite")),
        }];
        let mut evaluator = ChainEvaluator::default();
        assert!(matches!(
            plan(
                &engine,
                &owner,
                &formulas,
                &overlay,
                limits(),
                &mut evaluator,
            ),
            Err(Failure::UnsupportedDependency(Unsupported::Formula))
        ));
        assert!(evaluator.calls.is_empty());
    }

    #[test]
    fn formula_ref_counts_require_exact_host_key_multiset() {
        let bytes = [0x0a, 0x00];
        let entries = [FormulaListEntry {
            key: 9,
            ref_count: 2,
            bytes: &bytes,
        }];
        let payloads = [
            FormulaPayload {
                owner: OWNER,
                coordinate: coordinate(1),
                key: 9,
                bytes: &bytes,
            },
            FormulaPayload {
                owner: OWNER,
                coordinate: coordinate(2),
                key: 9,
                bytes: &bytes,
            },
        ];
        let usage = validate_formula_coverage(TABLE, &entries, &payloads, limits())
            .expect("exact host multiset");
        let exact = usize::try_from(usage.graph_work).expect("test usage fits");
        let mut one_short = limits();
        one_short.graph_work = exact - 1;
        assert!(matches!(
            validate_formula_coverage(TABLE, &entries, &payloads, one_short),
            Err(Failure::LimitExceeded { maximum, .. }) if maximum == usage.graph_work - 1
        ));

        let wrong_count = [FormulaListEntry {
            ref_count: 1,
            ..entries[0]
        }];
        assert!(matches!(
            validate_formula_coverage(TABLE, &wrong_count, &payloads, limits()),
            Err(Failure::InvalidSource)
        ));
    }

    #[test]
    fn unsupported_owner_rolls_back_without_evaluation() {
        let (engine, owner) = encoded_graph(vec![record(1, &[0], false)], true);
        let formulas = [formula(1, 20)];
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(1.0).expect("finite")),
        }];
        let mut evaluator = ChainEvaluator::default();
        let result = plan(
            &engine,
            &owner,
            &formulas,
            &overlay,
            limits(),
            &mut evaluator,
        );
        assert!(matches!(
            result,
            Err(Failure::UnsupportedDependency(Unsupported::Volatile))
        ));
        assert!(evaluator.calls.is_empty());
    }

    #[test]
    fn native_present_dormant_owner_envelopes_are_accepted() {
        let (engine, owner) = encoded_native_dormant_graph(vec![record(1, &[0], false)], false);
        let formulas = [formula(1, 20)];
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(4.0).expect("finite")),
        }];
        let mut evaluator = ChainEvaluator::default();
        let result = plan(
            &engine,
            &owner,
            &formulas,
            &overlay,
            limits(),
            &mut evaluator,
        )
        .expect("dormant native envelopes");
        assert_eq!(evaluator.calls, vec![coordinate(1)]);
        assert!(matches!(
            result.rewrites[0].cells[0].value,
            FormulaCachedValue::Number(value) if value.get() == 5.0
        ));
    }

    #[test]
    fn builder_empty_spans_are_admitted_only_for_a_fully_inert_owner() {
        let engine = tsce::CalculationEngineArchive {
            dependency_tracker: tsce::DependencyTrackerArchive {
                formula_owner_dependencies: vec![tsp::Reference {
                    identifier: 1,
                    ..Default::default()
                }],
                ..Default::default()
            },
            ..Default::default()
        }
        .encode_to_vec();
        let mut owner = tsce::FormulaOwnerDependenciesArchive {
            formula_owner_uid: tsp::Uuid {
                lower: TABLE.uuid_lower,
                upper: TABLE.uuid_upper,
            },
            internal_formula_owner_id: OWNER,
            spanning_column_dependencies: Some(tsce::SpanningDependenciesExpandedArchive::default()),
            spanning_row_dependencies: Some(tsce::SpanningDependenciesExpandedArchive::default()),
            ..Default::default()
        };
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(4.0).expect("finite")),
        }];
        let mut evaluator = ChainEvaluator::default();
        let inert = plan(
            &engine,
            &owner.encode_to_vec(),
            &[],
            &overlay,
            limits(),
            &mut evaluator,
        )
        .expect("builder-empty dependency owner is inert");
        assert!(inert.rewrites.is_empty());
        assert!(inert.removals.is_empty());
        assert!(evaluator.calls.is_empty());

        owner.cell_dependencies = Some(tsce::CellDependenciesExpandedArchive {
            cell_record: vec![record(1, &[0], false)],
        });
        let formulas = [formula(1, 20)];
        assert!(matches!(
            plan(
                &engine,
                &owner.encode_to_vec(),
                &formulas,
                &overlay,
                limits(),
                &mut evaluator,
            ),
            Err(Failure::UnsupportedDependency(Unsupported::Spanning))
        ));
        assert!(evaluator.calls.is_empty());
    }

    #[test]
    fn hostile_native_spans_and_active_ranges_remain_refused() {
        let formulas = [formula(1, 20)];
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(4.0).expect("finite")),
        }];
        let (engine, hostile_span_owner) =
            encoded_native_dormant_graph(vec![record(1, &[0], false)], true);
        let mut evaluator = ChainEvaluator::default();
        assert!(matches!(
            plan(
                &engine,
                &hostile_span_owner,
                &formulas,
                &overlay,
                limits(),
                &mut evaluator,
            ),
            Err(Failure::UnsupportedDependency(Unsupported::Spanning))
        ));
        assert!(evaluator.calls.is_empty());

        let (engine, dormant_owner) =
            encoded_native_dormant_graph(vec![record(1, &[0], false)], false);
        let mut active_owner =
            tsce::FormulaOwnerDependenciesArchive::decode(dormant_owner.as_slice())
                .expect("test owner");
        active_owner.range_dependencies = Some(tsce::RangeDependenciesArchive {
            back_dependency: vec![tsce::RangeBackDependencyArchive {
                cell_coord_row: 0,
                cell_coord_column: 1,
                internal_range_reference: Some(tsce::InternalRangeReferenceArchive {
                    owner_id: OWNER,
                    range: tsce::RangeCoordinateArchive {
                        top_left_column: 0,
                        top_left_row: 0,
                        bottom_right_column: 0,
                        bottom_right_row: 0,
                    },
                }),
                ..Default::default()
            }],
        });
        let active_owner = active_owner.encode_to_vec();
        assert!(matches!(
            plan(
                &engine,
                &active_owner,
                &formulas,
                &overlay,
                limits(),
                &mut evaluator,
            ),
            Err(Failure::UnsupportedDependency(Unsupported::Range))
        ));
        assert!(evaluator.calls.is_empty());
        assert!(!inactive_spill_sizes(&[0x0a, 0x00]));
    }

    #[test]
    fn owner_qualification_and_missing_owner_fail_typed() {
        let (engine, owner) = encoded_graph(vec![], false);
        let formulas = [FormulaCell {
            owner: OWNER + 1,
            coordinate: coordinate(1),
            cache_object: 20,
        }];
        let mut evaluator = ChainEvaluator::default();
        let wrong_owner = plan(&engine, &owner, &formulas, &[], limits(), &mut evaluator);
        assert!(matches!(
            wrong_owner,
            Err(Failure::UnsupportedDependency(Unsupported::ExternalOwner))
        ));

        let absent_engine = tsce::CalculationEngineArchive {
            dependency_tracker: tsce::DependencyTrackerArchive::default(),
            ..Default::default()
        }
        .encode_to_vec();
        let missing = plan(
            &absent_engine,
            &owner,
            &[formula(1, 20)],
            &[],
            limits(),
            &mut evaluator,
        );
        assert!(matches!(
            missing,
            Err(Failure::UnsupportedDependency(Unsupported::MissingOwner))
                | Err(Failure::InvalidSource)
        ));
    }

    #[test]
    fn fanout_graph_work_is_bounded_at_exact_max_minus_one() {
        let records = (1..=16).map(|column| record(column, &[0], false)).collect();
        let (engine, owner) = encoded_graph(records, false);
        let formulas: Vec<_> = (1..=16).map(|column| formula(column, 20)).collect();
        let overlay = [FinalCell {
            table: TABLE,
            coordinate: coordinate(0),
            value: FinalValue::number(FiniteF64::new(1.0).expect("finite")),
        }];
        let mut evaluator = FanoutEvaluator;
        let successful = plan(
            &engine,
            &owner,
            &formulas,
            &overlay,
            limits(),
            &mut evaluator,
        )
        .expect("bounded fanout plan");
        let exact = usize::try_from(successful.usage.graph_work).expect("test work fits usize");
        let mut one_short = limits();
        one_short.graph_work = exact - 1;
        let mut evaluator = FanoutEvaluator;
        let rejected = plan(
            &engine,
            &owner,
            &formulas,
            &overlay,
            one_short,
            &mut evaluator,
        );
        assert!(matches!(
            rejected,
            Err(Failure::LimitExceeded { maximum, .. }) if maximum == successful.usage.graph_work - 1
        ));
    }

    #[test]
    fn strict_evaluator_returns_numeric_and_boolean_caches() {
        use tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};

        let number_formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![
                    AstNodeArchive {
                        ast_node_type: AstNodeType::NumberNode as i32,
                        ast_number_node_number: Some(1.0),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::NumberNode as i32,
                        ast_number_node_number: Some(2.0),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::AdditionNode as i32,
                        ..Default::default()
                    },
                ],
            },
            ..Default::default()
        }
        .encode_to_vec();
        let boolean_formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![
                    AstNodeArchive {
                        ast_node_type: AstNodeType::BooleanNode as i32,
                        ast_boolean_node_boolean: Some(true),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::BooleanNode as i32,
                        ast_boolean_node_boolean: Some(false),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::EqualToNode as i32,
                        ..Default::default()
                    },
                ],
            },
            ..Default::default()
        }
        .encode_to_vec();
        let payloads = [
            FormulaPayload {
                owner: OWNER,
                coordinate: coordinate(1),
                key: 1,
                bytes: &number_formula,
            },
            FormulaPayload {
                owner: OWNER,
                coordinate: coordinate(2),
                key: 2,
                bytes: &boolean_formula,
            },
        ];
        let entries = [
            FormulaListEntry {
                key: 1,
                ref_count: 1,
                bytes: &number_formula,
            },
            FormulaListEntry {
                key: 2,
                ref_count: 1,
                bytes: &boolean_formula,
            },
        ];
        let (mut evaluator, _coverage) =
            StrictEvaluator::new(&entries, &payloads, 1, 4, TABLE, limits())
                .expect("strict evaluator");
        let values = Baseline(Vec::new());
        let number = evaluator
            .evaluate(formula(1, 20), &values, limits())
            .expect("numeric evaluation");
        let boolean = evaluator
            .evaluate(formula(2, 20), &values, limits())
            .expect("Boolean evaluation");
        assert!(matches!(
            number.value,
            FormulaCachedValue::Number(value) if value.get() == 3.0
        ));
        assert!(matches!(boolean.value, FormulaCachedValue::Boolean(false)));
    }

    #[test]
    fn strict_evaluator_refuses_date_duration_and_unmodeled_nodes() {
        use tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};

        for kind in [
            AstNodeType::DateNode,
            AstNodeType::DurationNode,
            AstNodeType::UnknownFunctionNode,
        ] {
            let bytes = tsce::FormulaArchive {
                ast_node_array: tsce::AstNodeArrayArchive {
                    ast_node: vec![AstNodeArchive {
                        ast_node_type: kind as i32,
                        ..Default::default()
                    }],
                },
                ..Default::default()
            }
            .encode_to_vec();
            let payloads = [FormulaPayload {
                owner: OWNER,
                coordinate: coordinate(1),
                key: 1,
                bytes: &bytes,
            }];
            let entries = [FormulaListEntry {
                key: 1,
                ref_count: 1,
                bytes: &bytes,
            }];
            let (mut evaluator, _coverage) =
                StrictEvaluator::new(&entries, &payloads, 1, 4, TABLE, limits())
                    .expect("strict evaluator");
            assert!(matches!(
                evaluator.evaluate(formula(1, 20), &Baseline(Vec::new()), limits()),
                Err(Failure::UnsupportedDependency(Unsupported::Formula))
            ));
        }
    }

    #[test]
    fn strict_formula_decode_is_canonical_bounded_and_max_minus_one() {
        use tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};

        let valid = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![AstNodeArchive {
                    ast_node_type: AstNodeType::NumberNode as i32,
                    ast_number_node_number: Some(2.0),
                    ..Default::default()
                }],
            },
            ..Default::default()
        }
        .encode_to_vec();
        let evaluate = |bytes: &[u8], evaluator_limits: CacheLimits| {
            let entries = [FormulaListEntry {
                key: 1,
                ref_count: 1,
                bytes,
            }];
            let payloads = [FormulaPayload {
                owner: OWNER,
                coordinate: coordinate(1),
                key: 1,
                bytes,
            }];
            let (mut evaluator, _coverage) =
                StrictEvaluator::new(&entries, &payloads, 1, 4, TABLE, evaluator_limits)?;
            evaluator.evaluate(formula(1, 20), &Baseline(Vec::new()), evaluator_limits)
        };

        let successful = evaluate(&valid, limits()).expect("strict formula");
        let exact = successful.usage.formula_work;
        let mut one_short = limits();
        one_short.formula_work = exact - 1;
        assert!(matches!(
            evaluate(&valid, one_short),
            Err(Failure::LimitExceeded { maximum, .. }) if maximum == u64::try_from(exact - 1).unwrap()
        ));

        let mut duplicate = valid.clone();
        duplicate.extend_from_slice(&valid);
        assert!(matches!(
            evaluate(&duplicate, limits()),
            Err(Failure::InvalidSource)
        ));

        let mut noncanonical = vec![0x8a, 0x00];
        noncanonical.extend_from_slice(&valid[1..]);
        assert!(matches!(
            evaluate(&noncanonical, limits()),
            Err(Failure::InvalidSource)
        ));

        let mut unknown = valid.clone();
        unknown.extend_from_slice(&[0x50, 0x00]);
        assert!(matches!(
            evaluate(&unknown, limits()),
            Err(Failure::InvalidSource)
        ));

        let mut oversized = limits();
        oversized.wire_bytes = valid.len() - 1;
        assert!(matches!(
            evaluate(&valid, oversized),
            Err(Failure::LimitExceeded { maximum, .. }) if maximum == u64::try_from(valid.len() - 1).unwrap()
        ));

        let mut no_allocation = limits();
        no_allocation.allocations = 0;
        assert!(matches!(
            evaluate(&valid, no_allocation),
            Err(Failure::LimitExceeded {
                observed: 1,
                maximum: 0,
            })
        ));

        let mut one_byte_short = limits();
        one_byte_short.scratch_bytes = size_of::<StreamingValue>() - 1;
        assert!(matches!(
            evaluate(&valid, one_byte_short),
            Err(Failure::LimitExceeded { maximum, .. })
                if maximum == u64::try_from(size_of::<StreamingValue>() - 1).unwrap()
        ));
    }
}
